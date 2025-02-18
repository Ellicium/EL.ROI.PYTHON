# EL.ROI.Python
Backend code for AI shouldcosting application


## Fastapi setup

### Linux
- python3.11 -m pip install virtualenv
- python3.11 -m virtualenv .venv
- source .venv/bin/activate
- pip install -r requirements.txt
- gunicorn -w 2 -k uvicorn.workers.UvicornWorker app.main:app -t 600 --reload

### Windows
- python -m venv .venv
- .venv\Scripts\activate
- pip install -r requirements.txt
- uvicorn app.main:app --port 8080 --workers 2 --reload



