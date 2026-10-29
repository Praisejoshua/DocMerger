"""
Word document merging service.

The first uploaded document becomes the "base". For every following document
we deep-copy its body elements into the base, in order, after:

* copying styles the base does not have,
* copying list/numbering definitions with fresh IDs and remapping references,
* re-creating embedded images and hyperlink relationships in the base package,
* renumbering drawing and bookmark IDs so they stay unique,
* inserting a section break so each document starts on a new page and keeps
  its own page size/margins.

Known limitations (reported to the user where detectable)
* Headers/footers of documents 2..n are not copied.
* Footnotes, endnotes and comments are left out.
* Charts, SmartArt and embedded objects cannot be copied and are left out.
* If two documents define the same style ID differently, the first one wins.
* Document defaults and theme (fonts/colours) come from the first document.

Nothing here executes macros or embedded content, and no document text is
ever written to logs.
"""
import copy
import io
import logging
import time
import zipfile
from pathlib import Path

from django.conf import settings
from docx import Document
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.parts.numbering import NumberingPart

logger = logging.getLogger(__name__)

R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"  # legacy .doc OR password-protected .docx

ISSUE_TEXT = {
    "notes": "footnotes and endnotes were left out",
    "comments": "comments were left out",
    "objects": "some charts, diagrams or embedded objects could not be copied and were left out",
    "numbering": "list numbering may differ from the original",
}


class MergeError(Exception):
    """An error whose message is safe and helpful to show to the user."""


def clean_filename(name: str) -> str:
    """Return a safe, display-only filename (no paths, no control characters)."""
    name = (name or "").replace("\\", "/").split("/")[-1]
    name = "".join(ch for ch in name if ch.isprintable()).strip()
    return name[:100] or "document.docx"


def storage_dir() -> Path:
    path = Path(settings.DOCMERGE_STORAGE_DIR)
    path.mkdir(parents=True, exist_ok=True)
    return path


def result_path(token: str) -> Path:
    """Path of a merged result. `token` must already be a validated UUID string."""
    return storage_dir() / f"{token}.docx"


def remove_file(path: Path) -> None:
    try:
        Path(path).unlink(missing_ok=True)
    except OSError:
        pass


def purge_expired_results() -> None:
    """Delete merged files (and leftovers from failed merges) older than the TTL."""
    cutoff = time.time() - settings.DOCMERGE_RESULT_TTL_MINUTES * 60
    folder = storage_dir()
    for pattern in ("*.docx", "*.part"):
        for path in folder.glob(pattern):
            try:
                if path.stat().st_mtime < cutoff:
                    path.unlink()
            except OSError:
                continue


def _to_int(value, default=0) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def validate_docx(data: bytes, display_name: str) -> None:
    """Check that `data` is a readable, unencrypted .docx package."""
    if data[:8] == OLE_MAGIC:
        raise MergeError(
            f"'{display_name}' appears to be password-protected or in an older Word format. "
            "Remove the password or save it as a .docx file, then try again."
        )

    stream = io.BytesIO(data)
    if not zipfile.is_zipfile(stream):
        raise MergeError(f"'{display_name}' is not a valid Word (.docx) file. It may be corrupted.")

    try:
        with zipfile.ZipFile(stream) as archive:
            names = set(archive.namelist())
            if "[Content_Types].xml" not in names or "word/document.xml" not in names:
                raise MergeError(f"'{display_name}' is not a valid Word (.docx) document.")
            unpacked = sum(info.file_size for info in archive.infolist())
            if unpacked > settings.DOCMERGE_MAX_UNCOMPRESSED_MB * 1024 * 1024:
                raise MergeError(f"'{display_name}' is too large once unpacked to be processed safely.")
    except zipfile.BadZipFile as exc:
        raise MergeError(f"'{display_name}' is corrupted and cannot be read.") from exc


class _Merger:
    def __init__(self, base):
        self.base = base
        self.body = base.element.body
        ids = [_to_int(n.get("id")) for n in self.body.iter(qn("wp:docPr"))]
        self._docpr_id = max(ids, default=0)

    def append(self, src, name: str, warnings: list) -> None:
        issues: set = set()
        num_map = self._merge_numbering(src, issues)
        self._merge_styles(src, num_map)

        src_body = src.element.body
        src_sect = src_body.find(qn("w:sectPr"))
        skip = (qn("w:sectPr"), qn("w:commentRangeStart"), qn("w:commentRangeEnd"))
        elements = [copy.deepcopy(c) for c in src_body.iterchildren() if c.tag not in skip]

        rid_map: dict = {}
        bookmark_offset = self._next_bookmark_id()
        for el in elements:
            self._remap_numbering(el, num_map)
            for sect in el.iter(qn("w:sectPr")):
                self._clean_section(sect)
            self._strip_notes_and_comments(el, issues)
            self._remap_relationships(src.part, el, rid_map, issues)
            self._renumber_ids(el, bookmark_offset)

        final_sect = self.body.get_or_add_sectPr()
        self._insert_section_break(final_sect)
        for el in elements:
            final_sect.addprevious(el)

        first_sect = next((s for el in elements for s in el.iter(qn("w:sectPr"))), None)
        if src_sect is not None:
            new_sect = copy.deepcopy(src_sect)
            self._clean_section(new_sect)
            self.body.replace(final_sect, new_sect)
            first_sect = first_sect if first_sect is not None else new_sect
        elif first_sect is None:
            first_sect = final_sect
        for section_type in first_sect.findall(qn("w:type")):
            first_sect.remove(section_type)

        for key in sorted(issues):
            warnings.append(f"'{name}': {ISSUE_TEXT[key]}.")

    def _merge_styles(self, src, num_map) -> None:
        base_styles = self.base.styles.element
        existing = {s.get(qn("w:styleId")) for s in base_styles.findall(qn("w:style"))}
        for style in src.styles.element.findall(qn("w:style")):
            style_id = style.get(qn("w:styleId"))
            if style_id in existing:
                continue
            clone = copy.deepcopy(style)
            clone.attrib.pop(qn("w:default"), None)
            self._remap_numbering(clone, num_map)
            base_styles.append(clone)
            existing.add(style_id)

    def _base_numbering(self):
        part = self.base.part
        try:
            return part.numbering_part._element
        except (KeyError, NotImplementedError):
            pass
        xml = '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
        new_part = NumberingPart(PackURI("/word/numbering.xml"), CT.WML_NUMBERING, parse_xml(xml), part.package)
        part.relate_to(new_part, RT.NUMBERING)
        return new_part._element

    def _merge_numbering(self, src, issues) -> dict:
        try:
            src_numbering = src.part.numbering_part._element
        except (KeyError, NotImplementedError):
            return {}
        src_nums = src_numbering.findall(qn("w:num"))
        if not src_nums:
            return {}

        try:
            base_numbering = self._base_numbering()
        except Exception as exc:
            logger.warning("Could not prepare numbering part (%s)", type(exc).__name__)
            issues.add("numbering")
            return {}

        abs_offset = max(
            (_to_int(a.get(qn("w:abstractNumId"))) for a in base_numbering.findall(qn("w:abstractNum"))),
            default=-1,
        ) + 1
        num_offset = max(
            (_to_int(n.get(qn("w:numId"))) for n in base_numbering.findall(qn("w:num"))),
            default=0,
        ) + 1

        first_num = base_numbering.find(qn("w:num"))
        for abstract in src_numbering.findall(qn("w:abstractNum")):
            clone = copy.deepcopy(abstract)
            clone.set(qn("w:abstractNumId"), str(_to_int(abstract.get(qn("w:abstractNumId"))) + abs_offset))
            for pic_bullet in list(clone.iter(qn("w:lvlPicBulletId"))):
                pic_bullet.getparent().remove(pic_bullet)
            if first_num is not None:
                first_num.addprevious(clone)
            else:
                base_numbering.append(clone)

        num_map = {}
        for num in src_nums:
            clone = copy.deepcopy(num)
            old_id = num.get(qn("w:numId"))
            new_id = str(_to_int(old_id) + num_offset)
            clone.set(qn("w:numId"), new_id)
            ref = clone.find(qn("w:abstractNumId"))
            if ref is not None:
                ref.set(qn("w:val"), str(_to_int(ref.get(qn("w:val"))) + abs_offset))
            base_numbering.append(clone)
            num_map[old_id] = new_id
        return num_map

    @staticmethod
    def _remap_numbering(root, num_map) -> None:
        if not num_map:
            return
        for node in root.iter(qn("w:numId")):
            old = node.get(qn("w:val"))
            if old in num_map:
                node.set(qn("w:val"), num_map[old])

    @staticmethod
    def _clean_section(sect) -> None:
        for tag in ("w:headerReference", "w:footerReference"):
            for ref in sect.findall(qn(tag)):
                sect.remove(ref)

    def _insert_section_break(self, final_sect) -> None:
        """End the previous document's last section so it keeps its own layout."""
        marker = copy.deepcopy(final_sect)
        previous = final_sect.getprevious()
        if previous is not None and previous.tag == qn("w:p"):
            ppr = previous.get_or_add_pPr()
            if ppr.find(qn("w:sectPr")) is None:
                ppr.append(marker)
                return
        paragraph = OxmlElement("w:p")
        ppr = OxmlElement("w:pPr")
        ppr.append(marker)
        paragraph.append(ppr)
        final_sect.addprevious(paragraph)

    @staticmethod
    def _strip_notes_and_comments(root, issues) -> None:
        for tag, issue in (
            ("w:commentRangeStart", "comments"),
            ("w:commentRangeEnd", "comments"),
        ):
            for node in list(root.iter(qn(tag))):
                node.getparent().remove(node)
                issues.add(issue)
        for tag, issue in (
            ("w:commentReference", "comments"),
            ("w:footnoteReference", "notes"),
            ("w:endnoteReference", "notes"),
        ):
            for node in list(root.iter(qn(tag))):
                run = node.getparent()
                if run is not None and run.getparent() is not None:
                    run.getparent().remove(run)
                    issues.add(issue)

    def _copy_relationship(self, src_part, rid):
        """Create the equivalent of src relationship `rid` in the base; return new rId or None."""
        try:
            rel = src_part.rels[rid]
        except KeyError:
            return None
        base_part = self.base.part
        if rel.is_external:
            return base_part.relate_to(rel.target_ref, rel.reltype, is_external=True)
        if rel.reltype == RT.IMAGE:
            try:
                new_rid, _image = base_part.get_or_add_image(io.BytesIO(rel.target_part.blob))
                return new_rid
            except Exception:
                return None
        return None

    def _remap_relationships(self, src_part, root, rid_map, issues) -> None:
        prefix = "{%s}" % R_NS
        for node in list(root.iter("*")):
            for attr, rid in list(node.attrib.items()):
                if not attr.startswith(prefix):
                    continue
                if rid not in rid_map:
                    rid_map[rid] = self._copy_relationship(src_part, rid)
                new_rid = rid_map[rid]
                if new_rid is not None:
                    node.set(attr, new_rid)
                    continue

                if node.tag.endswith("}svgBlip"):
                    ext = node.getparent()
                    if ext is not None and ext.getparent() is not None:
                        ext.getparent().remove(ext)
                    continue
                if node.tag == qn("w:hyperlink"):
                    node.attrib.pop(attr, None)
                    continue
                run = node
                while run is not None and run.tag != qn("w:r"):
                    run = run.getparent()
                if run is not None and run.getparent() is not None:
                    run.getparent().remove(run)
                else:
                    node.attrib.pop(attr, None)
                issues.add("objects")

    def _next_bookmark_id(self) -> int:
        ids = [
            _to_int(n.get(qn("w:id")))
            for n in self.body.iter(qn("w:bookmarkStart"), qn("w:bookmarkEnd"))
        ]
        return max(ids, default=0) + 1

    def _renumber_ids(self, root, bookmark_offset) -> None:
        for node in root.iter(qn("wp:docPr")):
            self._docpr_id += 1
            node.set("id", str(self._docpr_id))
        for node in root.iter(qn("w:bookmarkStart"), qn("w:bookmarkEnd")):
            node.set(qn("w:id"), str(_to_int(node.get(qn("w:id"))) + bookmark_offset))


def merge_documents(sources: list, output_path: Path) -> list:
    """
    Merge documents in the given order and save the result to `output_path`.

    `sources` is a list of (display_name, raw_bytes) tuples.
    Returns a list of human-readable warnings. Raises MergeError on failure.
    """
    if len(sources) < 2:
        raise MergeError("Please provide at least two documents to merge.")

    for display_name, data in sources:
        validate_docx(data, display_name)

    def load(display_name, data):
        try:
            return Document(io.BytesIO(data))
        except Exception as exc:
            logger.warning("A document could not be opened (%s)", type(exc).__name__)
            raise MergeError(
                f"'{display_name}' could not be opened. It may be corrupted, "
                "password-protected, or not a standard Word document."
            ) from exc

    base_name, base_data = sources[0]
    merger = _Merger(load(base_name, base_data))
    warnings: list = []

    for display_name, data in sources[1:]:
        source_doc = load(display_name, data)
        try:
            merger.append(source_doc, display_name, warnings)
        except MergeError:
            raise
        except Exception as exc:
            logger.warning("A document could not be merged (%s)", type(exc).__name__)
            raise MergeError(
                f"'{display_name}' could not be merged. It may use features this tool cannot copy."
            ) from exc

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        merger.base.save(str(output_path))
    except Exception as exc:
        logger.warning("Saving the merged document failed (%s)", type(exc).__name__)
        raise MergeError("The merged document could not be saved. Please try again.") from exc

    return list(dict.fromkeys(warnings))
