import logging
import uuid

from django.conf import settings
from django.core.exceptions import ValidationError
from django.http import FileResponse, HttpResponse, JsonResponse
from django.shortcuts import render
from django.urls import reverse
from django.views.decorators.http import require_GET, require_POST

from .forms import validate_uploads
from .services import (
    MergeError,
    clean_filename,
    merge_documents,
    purge_expired_results,
    remove_file,
    result_path,
)

logger = logging.getLogger(__name__)

SESSION_KEY = "docmerge_results"
OUTPUT_NAME = "merged_document.docx"
DOCX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


def _json_error(message: str, status: int = 400) -> JsonResponse:
    return JsonResponse({"ok": False, "error": message}, status=status)


@require_GET
def home(request):
    return render(
        request,
        "merger/home.html",
        {
            "min_files": settings.DOCMERGE_MIN_FILES,
            "max_files": settings.DOCMERGE_MAX_FILES,
            "max_file_size_mb": settings.DOCMERGE_MAX_FILE_SIZE_MB,
            "ttl_minutes": settings.DOCMERGE_RESULT_TTL_MINUTES,
        },
    )


@require_POST
def merge(request):
    """Validate uploads, merge them in the submitted order, store the result privately."""
    purge_expired_results()

    uploads = request.FILES.getlist("documents")
    logger.info("Merge request received %d file(s)", len(uploads))  # counts only
    try:
        validate_uploads(uploads)
    except ValidationError as exc:
        return _json_error(" ".join(exc.messages))

    sources = []
    for upload in uploads:
        upload.seek(0)
        sources.append((clean_filename(upload.name), upload.read()))

    token = str(uuid.uuid4())
    final_path = result_path(token)
    part_path = final_path.with_suffix(".part")

    try:
        warnings = merge_documents(sources, part_path)
        part_path.replace(final_path)
    except MergeError as exc:
        remove_file(part_path)
        return _json_error(str(exc), status=422)
    except Exception as exc:
        remove_file(part_path)
        remove_file(final_path)
        logger.error("Unexpected merge failure (%s)", type(exc).__name__)
        return _json_error("Something went wrong while merging. Please try again.", status=500)

    tokens = request.session.get(SESSION_KEY, [])
    tokens.append(token)
    request.session[SESSION_KEY] = tokens[-10:]

    logger.info("Merged %d documents", len(sources))
    return JsonResponse(
        {
            "ok": True,
            "filename": OUTPUT_NAME,
            "size": final_path.stat().st_size,
            "warnings": warnings,
            "download_url": reverse("merger:download", args=[token]),
            "discard_url": reverse("merger:discard", args=[token]),
        }
    )


@require_GET
def download(request, token):
    token = str(token)
    if token not in request.session.get(SESSION_KEY, []):
        return HttpResponse("Not found.", status=404, content_type="text/plain")

    path = result_path(token)
    if not path.is_file():
        return HttpResponse(
            "This download has expired. Please merge your documents again.",
            status=404,
            content_type="text/plain",
        )

    response = FileResponse(
        open(path, "rb"),
        as_attachment=True,
        filename=OUTPUT_NAME,
        content_type=DOCX_CONTENT_TYPE,
    )
    response["Cache-Control"] = "no-store"
    return response


@require_POST
def discard(request, token):
    token = str(token)
    tokens = request.session.get(SESSION_KEY, [])
    if token in tokens:
        remove_file(result_path(token))
        request.session[SESSION_KEY] = [t for t in tokens if t != token]
    return JsonResponse({"ok": True})
