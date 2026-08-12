from app import app

class VercelPrefixMiddleware:
    """Strips Vercel routing prefixes (/api/index, /api) from PATH_INFO so Flask routes work seamlessly."""
    def __init__(self, app):
        self.app = app

    def __call__(self, environ, start_response):
        path = environ.get("PATH_INFO", "")
        if path.startswith("/api/index"):
            environ["PATH_INFO"] = path[10:] or "/"
        elif path.startswith("/api"):
            environ["PATH_INFO"] = path[4:] or "/"
        return self.app(environ, start_response)

app.wsgi_app = VercelPrefixMiddleware(app.wsgi_app)
