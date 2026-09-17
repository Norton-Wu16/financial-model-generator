"""一键启动脚本。

用法:
    python run.py

启动 FastAPI 后端 (http://localhost:8000)，同时托管前端单页应用。
打开浏览器访问 http://localhost:8000 即可使用。
"""
import webbrowser
import threading
import uvicorn


def open_browser():
    """延迟打开浏览器，等服务器启动。"""
    import time
    time.sleep(1.5)
    webbrowser.open("http://localhost:8000")


if __name__ == "__main__":
    threading.Thread(target=open_browser, daemon=True).start()
    print("=" * 60)
    print("  财务模型生成器 Financial Model Generator")
    print("  访问 http://localhost:8000")
    print("  按 Ctrl+C 停止")
    print("=" * 60)
    uvicorn.run("backend.main:app", host="0.0.0.0", port=8000, reload=False)
