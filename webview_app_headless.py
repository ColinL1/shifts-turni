"""Fallback launcher that just opens the default browser instead of embedding a WebView.
Use when pyobjc / WebKit is problematic. Build with: ENTRY_SCRIPT=webview_app_headless.py python setup.py py2app
"""
import threading
import webbrowser
import time

import app as flask_app_module

PORT = 5001


def start_flask():
    flask_app_module.app.run(port=PORT, debug=False, use_reloader=False)


def main():
    threading.Thread(target=start_flask, daemon=True).start()
    # wait briefly for server
    for _ in range(50):
        time.sleep(0.1)
    webbrowser.open(f'http://127.0.0.1:{PORT}/')

if __name__ == '__main__':
    main()
