"""Desktop entry point using pywebview to show the Flask app without an external browser.

Run locally (dev):
    python webview_app.py

When bundled (py2app / pyinstaller), this script is the main entry.
"""
import threading
import time
import socket
import os
import certifi
import sys

# Delay importing webview until after patching pyobjc __file__
def _prepare_pyobjc():
    try:
        import objc  # type: ignore
        import AppKit  # noqa: F401
        import Foundation  # noqa: F401
        import WebKit  # noqa: F401
        if not hasattr(objc._objc, '__file__'):
            bundle_root = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
            search_roots = [bundle_root,
                            os.path.join(bundle_root, 'Frameworks'),
                            os.path.join(os.path.dirname(sys.executable), '..', 'Frameworks')]
            found = None
            for root in search_roots:
                if not os.path.exists(root):
                    continue
                for dirpath, _, filenames in os.walk(root):
                    for fn in filenames:
                        if fn.startswith('_objc') and (fn.endswith('.so') or fn.endswith('.dylib')):
                            found = os.path.join(dirpath, fn)
                            break
                    if found:
                        break
                if found:
                    break
            if found:
                try:
                    setattr(objc._objc, '__file__', found)
                    print(f"[webview_app] Patched objc._objc.__file__ -> {found}")
                except Exception:
                    pass
        return True
    except Exception as e:  # pragma: no cover
        print(f"[webview_app] Warning: pyobjc preparation failed: {e}")
        return False

_prepare_pyobjc()

import webview  # noqa: E402  (import after pyobjc patch)

# Use the existing Flask app
import app as flask_app_module

PORT = 5001


def wait_for_server(host: str, port: int, timeout: float = 15.0):
    """Wait until the Flask server is accepting connections."""
    start = time.time()
    while time.time() - start < timeout:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.settimeout(0.5)
            try:
                sock.connect((host, port))
                return True
            except OSError:
                time.sleep(0.2)
    return False


def start_flask():
    # Disable reloader when running in a thread
    flask_app_module.app.run(port=PORT, debug=False, use_reloader=False)


def main():
    # Ensure cert path for SSL module in frozen environment
    os.environ.setdefault('SSL_CERT_FILE', certifi.where())
    threading.Thread(target=start_flask, daemon=True).start()
    if not wait_for_server('127.0.0.1', PORT):
        raise SystemExit("Failed to start internal server")

    webview.create_window('Analizzatore Turni', f'http://127.0.0.1:{PORT}/', width=1400, height=900, resizable=True, easy_drag=True)
    webview.start(debug=True)


if __name__ == '__main__':
    main()
