# -*- coding: utf-8 -*-
"""一键启动：工作台(Flask) + cpolar 固定隧道

固定公网地址：https://testh.cpolar.cn
（cpolar 基础版，子域名 testh；地址永久不变，除非改隧道配置）
"""
import os
import socket
import subprocess
import sys
import time
import webbrowser

PROJ = os.path.dirname(os.path.abspath(__file__))
VENV_PY = os.path.join(PROJ, ".venv", "Scripts", "python.exe")
CPOLAR_EXE = os.path.join(os.path.dirname(PROJ), "cpolar", "cpolar.exe")
PUBLIC_URL = "https://testh.cpolar.cn"
PORT = 5000
NO_WINDOW = 0x08000000          # CREATE_NO_WINDOW


def port_open(port, host="127.0.0.1"):
    s = socket.socket()
    s.settimeout(1)
    try:
        return s.connect_ex((host, port)) == 0
    finally:
        s.close()


def proc_running(name):
    out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq " + name],
                         capture_output=True, text=True, errors="ignore").stdout
    return name.lower() in (out or "").lower()


def main():
    py = VENV_PY if os.path.exists(VENV_PY) else sys.executable

    # 1) 工作台
    if port_open(PORT):
        print("[1/3] 工作台已在运行：http://127.0.0.1:%d" % PORT)
    else:
        print("[1/3] 启动工作台 ...")
        subprocess.Popen([py, "app.py"], cwd=PROJ, creationflags=NO_WINDOW)
        for _ in range(40):
            if port_open(PORT):
                break
            time.sleep(0.5)
        print("      已就绪：http://127.0.0.1:%d" % PORT)

    # 2) cpolar 隧道
    if proc_running("cpolar.exe"):
        print("[2/3] cpolar 已在运行")
    else:
        if not os.path.exists(CPOLAR_EXE):
            print("[2/3] 找不到 cpolar：%s" % CPOLAR_EXE)
            return
        print("[2/3] 启动 cpolar 隧道 ...")
        subprocess.Popen([CPOLAR_EXE, "start-all"],
                         cwd=os.path.dirname(CPOLAR_EXE), creationflags=NO_WINDOW)
        time.sleep(8)

    # 3) 验证公网地址
    print("[3/3] 验证固定地址 %s ..." % PUBLIC_URL)
    import ssl
    import urllib.request
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    ok = False
    for _ in range(3):
        try:
            r = urllib.request.urlopen(PUBLIC_URL + "/login", timeout=12, context=ctx)
            ok = r.status == 200
            if ok:
                break
        except Exception:
            time.sleep(3)
    print("      公网访问：%s" % ("正常 ✓" if ok else "暂时不通（隧道可能还在建链，稍等 10 秒再试）"))
    print()
    print("固定地址：%s   （地址永久不变，可存为书签）" % PUBLIC_URL)
    if "--no-open" not in sys.argv:
        webbrowser.open(PUBLIC_URL)


if __name__ == "__main__":
    main()
