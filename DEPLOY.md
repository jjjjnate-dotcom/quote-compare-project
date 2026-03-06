# Web Deployment Guide

This project is now configured for production deployment with `gunicorn`.

## 1) Required environment variables

- `SECRET_KEY`: random long string for Flask session/flash security
- `PORT`: provided automatically by most platforms (do not hardcode)

## 2) Deploy on Render

1. Push this repo to GitHub.
2. In Render, create a new `Web Service` from the repo.
3. Use these settings:
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `gunicorn wsgi:app --bind 0.0.0.0:$PORT --workers 2 --timeout 120`
4. Add environment variable:
   - `SECRET_KEY` = any strong random value
5. Deploy.

## 3) Deploy on Railway

1. Create a new project from the GitHub repo.
2. Add environment variable:
   - `SECRET_KEY` = any strong random value
3. Railway detects `Procfile` and runs:
   - `web: gunicorn wsgi:app --bind 0.0.0.0:$PORT --workers 2 --timeout 120`

## 4) Deploy on a Linux VM (manual)

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
export SECRET_KEY="replace-with-random-secret"
export PORT=8000
gunicorn wsgi:app --bind 0.0.0.0:$PORT --workers 2 --timeout 120
```

## 5) Notes

- Uploaded files are processed in a temporary directory and returned immediately.
- No database is required for current behavior.
- If you put this behind Nginx, keep request body size at least 20MB.
