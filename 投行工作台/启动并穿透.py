# -*- coding: utf-8 -*-
"""一键启动：工作台(Flask) + cloudflared 快速隧道，等地址真正可访问后再用新地址打开浏览器。"""
import os
import re
import sys
import time
import ssl
import subprocess
import webbrowser
import urllib.request

PROJ = os.path.dirname(os.path.abspath(__file__))
VENV_PY = os.path.join(PROJ, ".venv", "Scripts", "python.exe")
CLOUDFLARED = r"C:\Program Files (x86)\cloudflared\cloudflared.exe"
PORT = 5000
LOG = os.path.join(PROJ, ".tmp", "tunnel.log")


def port_in_use(port):
    """用 netstat 判断端口占用（Windows 上 socket.bind 因 SO_REUSEADDR 不可靠）。"""
    try:
        out = subprocess.run(["netstat", "-ano"], capture_output=True,
                             text=True, timeout=8).stdout
        key = ":%d " % port
        for line in out.splitlines():
            if key in line and "LISTENING" in line.upper():
                return True
    except Exception:
        pass
    return False


def wait_reachable(url, timeout=90):
    """等公网地址真正返回 200（国内访问 Cloudflare 较慢，最多等 timeout 秒）。"""
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            r = urllib.request.urlopen(url + "/login", timeout=8, context=ctx)
            if r.getcode() == 200:
                return True
        except Exception:
            pass
        time.sleep(2)
    return False


def main():
    os.makedirs(os.path.dirname(LOG), exist_ok=True)
    flask_proc = None
    cf_proc = None

    # 清理上次遗留的隧道进程
    try:
        subprocess.run(["taskkill", "/F", "/IM", "cloudflared.exe"],
                       capture_output=True, timeout=10)
    except Exception:
        pass

    # 1) 本地 Flask
    if port_in_use(PORT):
        print("[本地] 5000 端口已有服务，跳过启动 Flask")
    else:
        print("[本地] 启动工作台 Flask ...")
        flask_proc = subprocess.Popen([VENV_PY, "app.py"], cwd=PROJ)

    # 2) 隧道（最多重试 3 次）
    if not os.path.exists(CLOUDFLARED):
        print("[隧道] 未找到 cloudflared，只启动本地")
        webbrowser.open("http://127.0.0.1:%d" % PORT)
        print("本地地址: http://127.0.0.1:%d" % PORT)
        input("按回车退出 ...")
        return

    url = None
    for attempt in range(1, 4):
        print("[隧道] 第 %d 次启动隧道 ..." % attempt)
        logf = open(LOG, "w", encoding="utf-8", errors="ignore")
        cf_proc = subprocess.Popen(
            [CLOUDFLARED, "tunnel", "--no-autoupdate",
             "--url", "http://localhost:%d" % PORT],
            stdout=logf, stderr=subprocess.STDOUT)

        # 等地址字符串出现
        for _ in range(45):
            time.sleep(1)
            try:
                txt = open(LOG, encoding="utf-8", errors="ignore").read()
            except Exception:
                txt = ""
            m = re.search(r"https://[a-z0-9-]+\.trycloudflare\.com", txt)
            if m:
                url = m.group(0)
                break

        if url:
            print("[隧道] 地址: " + url)
            print("[隧道] 等待地址真正可访问（国内访问 Cloudflare 可能较慢）...")
            if wait_reachable(url, timeout=90):
                print("[隧道] 可访问！")
                break
            print("[隧道] 本次连接超时，准备重试 ...")
            try:
                cf_proc.terminate()
            except Exception:
                pass
            url = None
        else:
            print("[隧道] 未拿到地址，准备重试 ...")
            try:
                cf_proc.terminate()
            except Exception:
                pass

    # 3) 打开浏览器
    if url:
        print("")
        print("=" * 52)
        print("  本次公网地址（每次启动会变）:")
        print("  " + url)
        print("=" * 52)
        webbrowser.open(url)
    else:
        url = "http://127.0.0.1:%d" % PORT
        print("")
        print("[提示] 隧道多次失败，已打开本地地址（外网暂时不通，稍后重试）")
        webbrowser.open(url)

    print("")
    print("本地地址: http://127.0.0.1:%d" % PORT)
    print("登录: admin / th2026（首次登录强制改密）")
    print("")
    print("服务运行中... 关闭本窗口 = 停止全部服务（含隧道）")
    print("")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        pass
    finally:
        for p in (flask_proc, cf_proc):
            if p:
                try:
                    p.terminate()
                except Exception:
                    pass
        print("已停止。")


if __name__ == "__main__":
    main()
