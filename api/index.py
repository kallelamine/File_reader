# Vercel serverless entry: expose Flask app so all routes are bundled and run here.
# Rewrites send every path to /api/index; normalize so Flask sees /, /upload, etc.

from app import app as flask_app


def app(environ, start_response):
    path = (environ.get("PATH_INFO") or "/").strip()
    # Strip /api or /api/index so Flask receives /, /upload, /download/...
    if path.startswith("/api/index"):
        path = path[10:] or "/"
    elif path.startswith("/api"):
        path = path[4:] or "/"
    environ["PATH_INFO"] = path
    return flask_app(environ, start_response)
