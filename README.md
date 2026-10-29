# DocMerge – Word Document Merger

## Run it (first time)

Windows (PowerShell):
    python -m venv .venv
    .venv\Scripts\Activate.ps1
    pip install -r requirements.txt
    python manage.py migrate
    python manage.py runserver

macOS / Linux:
    python3 -m venv .venv
    source .venv/bin/activate
    pip install -r requirements.txt
    python manage.py migrate
    python manage.py runserver

Then open http://127.0.0.1:8000/

## Next times
Activate the venv, then: python manage.py runserver

## Run tests
    python manage.py test
