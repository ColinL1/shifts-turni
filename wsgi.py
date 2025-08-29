from app import app as application

# Gunicorn entrypoint expects a module-level variable named `application` by default when using `--chdir` or explicit module.
# We also expose `app` for clarity.
app = application
