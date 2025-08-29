from setuptools import setup
import os

APP = [os.environ.get('ENTRY_SCRIPT', 'webview_app.py')]  # Allow override: ENTRY_SCRIPT=webview_app_headless.py
def collect_templates():
    files = []
    for root, _, filenames in os.walk('templates'):
        for fn in filenames:
            if fn.endswith('.html'):
                full = os.path.join(root, fn)
                rel_dir = os.path.relpath(root, 'templates')
                target_dir = os.path.join('templates', rel_dir) if rel_dir != '.' else 'templates'
                files.append((target_dir, [full]))
    return files

DATA_FILES = collect_templates()


def find_libffi_candidates():
    candidates = [
        '/opt/homebrew/opt/libffi/lib/libffi.8.dylib',      # Apple Silicon (Homebrew)
        '/usr/local/opt/libffi/lib/libffi.8.dylib',         # Intel (Homebrew older path)
        '/opt/local/lib/libffi.8.dylib',                    # MacPorts
    ]
    return [c for c in candidates if os.path.exists(c)]

LIBFFI_LIST = find_libffi_candidates()

def find_openssl_libs():
    candidates = [
        '/opt/homebrew/opt/openssl@3/lib/libssl.3.dylib',
        '/opt/homebrew/opt/openssl@3/lib/libcrypto.3.dylib',
        '/usr/local/opt/openssl@3/lib/libssl.3.dylib',
        '/usr/local/opt/openssl@3/lib/libcrypto.3.dylib',
        '/opt/local/lib/libssl.3.dylib',
        '/opt/local/lib/libcrypto.3.dylib',
    ]
    return [c for c in candidates if os.path.exists(c)]

OPENSSL_LIBS = find_openssl_libs()
OPTIONS = {
    'argv_emulation': True,
    'includes': [
        'flask', 'docx', 'openpyxl', 'werkzeug', 'webview', 'python_docx', 'jinja2', 'markupsafe',
        'pkg_resources', 'importlib_metadata', 'jaraco', 'jaraco.text', 'jaraco.functools', 'jaraco.context',
        'jaraco.collections', 'more_itertools', 'packaging',
        'objc', 'Foundation', 'AppKit', 'WebKit'
    ],
    'excludes': ['tkinter', 'tests', 'pip', 'distutils'],
    'iconfile': 'app_icon.icns',
    'plist': {
        'CFBundleName': 'Analizzatore Turni',
        'CFBundleDisplayName': 'Analizzatore Turni',
        'CFBundleGetInfoString': 'Employee shift analyzer',
        'CFBundleIdentifier': 'com.example.analizzatoreturni',
        'CFBundleShortVersionString': '0.1.0',
        'CFBundleVersion': '0.1.0',
        'NSAppTransportSecurity': {'NSAllowsArbitraryLoads': True},
    },
    'packages': ['encodings', 'flask', 'docx', 'openpyxl', 'pkg_resources', 'importlib_metadata', 'webview', 'certifi'],
    'optimize': 0,
}

if LIBFFI_LIST:
    # Ensure libffi dylib(s) are bundled so _ctypes can load
    OPTIONS['frameworks'] = LIBFFI_LIST
else:
    print("[setup.py] WARNING: libffi.8.dylib not found – the built app may fail to start (_ctypes). Install via 'brew install libffi'.")

# Append OpenSSL dylibs (needed for _ssl) if found
if OPENSSL_LIBS:
    OPTIONS.setdefault('frameworks', [])
    for p in OPENSSL_LIBS:
        if p not in OPTIONS['frameworks']:
            OPTIONS['frameworks'].append(p)
else:
    print("[setup.py] WARNING: OpenSSL libssl.3/libcrypto.3 not found – SSL may fail. Install via 'brew install openssl@3'.")

setup(
    app=APP,
    name='Analizzatore Turni',
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
    install_requires=[r.strip() for r in open('requirements.txt').read().splitlines() if r.strip()],
)
