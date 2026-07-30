"""
数据脱敏工具 - PyInstaller 打包入口
双击此文件即可运行（打包后为 .exe）
"""
import os
import sys
import webbrowser
import threading
import time

# 环境变量方式设置端口（比 CLI 参数更可靠，避免 devMode 冲突）
os.environ["STREAMLIT_GLOBAL_DEVELOPMENT_MODE"] = "false"
os.environ["STREAMLIT_SERVER_PORT"] = "8501"
os.environ["STREAMLIT_SERVER_ADDRESS"] = "localhost"
os.environ["STREAMLIT_SERVER_HEADLESS"] = "true"
os.environ["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"

BASE = os.path.dirname(os.path.abspath(__file__))
APP = os.path.join(BASE, "mask_tool.py")


def open_browser():
    """延迟 2.5 秒打开浏览器，等 Streamlit 启动完毕"""
    time.sleep(2.5)
    webbrowser.open("http://localhost:8501")


if __name__ == "__main__":
    threading.Thread(target=open_browser, daemon=True).start()

    sys.argv = ["streamlit", "run", APP]
    from streamlit.web import cli
    sys.exit(cli.main())
