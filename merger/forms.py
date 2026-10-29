from django import forms
from django.conf import settings

from .services import clean_filename


def validate_uploads(files):
    """
    Validate the list of uploaded files (count, extension, size, not empty).
    Raises django.forms.ValidationError with user-friendly messages.
    Works the same on every Django version (no multi-file form widget needed).
    """
    min_files = settings.DOCMERGE_MIN_FILES
    max_files = settings.DOCMERGE_MAX_FILES
    max_mb = settings.DOCMERGE_MAX_FILE_SIZE_MB
    max_bytes = max_mb * 1024 * 1024

    if len(files) < min_files:
        raise forms.ValidationError(
            f"Please add at least {min_files} Word documents to merge "
            f"(the server received {len(files)})."
        )
    if len(files) > max_files:
        raise forms.ValidationError(
            f"You can merge up to {max_files} documents at a time."
        )

    errors = []
    for upload in files:
        name = clean_filename(upload.name)
        if not name.lower().endswith(".docx"):
            errors.append(
                f"'{name}' is not a .docx file. Only Word (.docx) documents are supported."
            )
        elif upload.size == 0:
            errors.append(f"'{name}' is empty.")
        elif upload.size > max_bytes:
            errors.append(f"'{name}' is larger than {max_mb} MB.")
    if errors:
        raise forms.ValidationError(errors)
