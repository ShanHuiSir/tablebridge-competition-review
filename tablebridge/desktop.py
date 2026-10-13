"""Windows-first native window; also usable on macOS with pywebview installed."""
import os
import sys
import threading
from pathlib import Path


def main():
    if '--sandbox-worker' in sys.argv:
        from sandbox import worker
        worker()
        return
    # A packaged executable must never write inside its read-only resource bundle.
    home = Path(os.getenv('LOCALAPPDATA', str(Path.home() / '.local' / 'share'))) / 'TableBridge'
    home.mkdir(parents=True, exist_ok=True)
    os.environ.setdefault('BRIDGE_DATA_DIR', str(home / 'data'))
    os.environ.setdefault('BRIDGE_ENV_FILE', str(home / '.env'))
    from app import Handler, ThreadingHTTPServer
    server = ThreadingHTTPServer(('127.0.0.1', 0), Handler)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    try:
        import webview
        webview.create_window('表桥 · 跨系统表格转换智能体', f'http://127.0.0.1:{server.server_port}', width=1420, height=920, min_size=(900, 650))
        webview.start()
    finally:
        server.shutdown()
        server.server_close()


if __name__ == '__main__':
    main()
