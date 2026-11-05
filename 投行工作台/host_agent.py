# -*- coding: utf-8 -*-
"""
宿主机代理 host_agent.py
========================

作用：让部署在云端的「投行工作台」能安全地请你本机读写文件——但**每一次都要你点确认**。

它做什么
--------
1. 首次运行用配对码换取 token，存在 ~/.workbuddy_host_agent.json
2. 每隔几秒向云端轮询「待执行的文件请求」
3. 每条请求都会打印出来（操作 / 路径 / 内容摘要），等你输入 y 才执行
4. 执行结果回传云端，页面即可看到

它不做什么
----------
- 不接受任何「免确认」「批量静默执行」指令
- 不越过白名单目录（默认：用户目录下的 WorkBuddy 目录 + 显式 --allow 的目录）
- 不主动上传文件内容以外的任何信息

用法
----
    # 首次配对（配对码在 配置管理 → 智能体配置 → 宿主机授权 里生成）
    python host_agent.py --server https://your-app.example.com --code 123456

    # 之后直接运行
    python host_agent.py --server https://your-app.example.com

    # 追加允许目录、自动确认（仅调试用，日常别加 -y）
    python host_agent.py --server ... --allow "D:\\投行资料" -y
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import time
import urllib.error
import urllib.parse
import urllib.request

CONF = os.path.join(os.path.expanduser("~"), ".workbuddy_host_agent.json")


def load_conf():
    if os.path.isfile(CONF):
        try:
            with open(CONF, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_conf(d):
    with open(CONF, "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False, indent=2)


def post(url, payload, timeout=30):
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(url, data=data, method="POST",
                                 headers={"Content-Type": "application/json; charset=utf-8"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read().decode("utf-8", "replace"))


def get(url, timeout=30):
    req = urllib.request.Request(url, headers={"User-Agent": "host-agent/1.0"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read().decode("utf-8", "replace"))


def within(path, dirs):
    """路径是否落在白名单目录内（不区分大小写，防 ../ 穿越）。"""
    if not dirs:
        return False
    real = os.path.realpath(os.path.expandvars(os.path.expanduser(path)))
    for d in dirs:
        base = os.path.realpath(os.path.expandvars(os.path.expanduser(d)))
        if real == base or real.startswith(base + os.sep):
            return True
    return False


def do_task(task):
    op = task.get("op")
    path = os.path.expandvars(os.path.expanduser(task.get("path") or ""))
    if op == "list":
        items = []
        for name in sorted(os.listdir(path))[:200]:
            full = os.path.join(path, name)
            items.append("%s%s" % (name, "/" if os.path.isdir(full) else ""))
        return True, "\n".join(items)
    if op == "read":
        size = os.path.getsize(path)
        with open(path, "rb") as f:
            blob = f.read(min(size, 2 * 1024 * 1024))
        try:
            return True, blob.decode("utf-8")
        except UnicodeDecodeError:
            return True, "[二进制文件 %d 字节，仅回传十六进制前 256 字节]\n%s" % (
                size, blob[:256].hex())
    if op == "write":
        d = os.path.dirname(path)
        if d and not os.path.isdir(d):
            os.makedirs(d, exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            f.write(task.get("content") or "")
        return True, "已写入 %d 字符" % len(task.get("content") or "")
    return False, "未知操作：%s" % op


def main():
    ap = argparse.ArgumentParser(description="WorkBuddy 宿主机代理")
    ap.add_argument("--server", required=True, help="云端地址，如 https://xxx.example.com")
    ap.add_argument("--code", help="首次配对码（6 位数字）")
    ap.add_argument("--allow", action="append", default=[], help="额外允许的目录，可多次指定")
    ap.add_argument("--interval", type=int, default=5, help="轮询间隔秒，默认 5")
    ap.add_argument("-y", "--yes", action="store_true", help="跳过确认（仅调试用）")
    args = ap.parse_args()

    server = args.server.rstrip("/")
    conf = load_conf()
    token = conf.get("token")

    if args.code:
        try:
            r = post(server + "/api/host/pair", {"code": args.code})
            token = r.get("token")
            if not token:
                print("配对失败：%s" % r.get("error"))
                return 1
            conf["token"] = token
            conf["server"] = server
            save_conf(conf)
            print("配对成功，token 已保存到 %s" % CONF)
        except Exception as e:  # noqa: BLE001
            print("配对请求失败：%s" % e)
            return 1

    if not token:
        print("还没有配对。请先到 配置管理 → 智能体配置 → 宿主机授权 生成配对码，再运行：")
        print("  python host_agent.py --server %s --code 123456" % server)
        return 1

    allow = list(conf.get("allow") or []) + args.allow
    if not allow:
        allow = [os.path.join(os.path.expanduser("~"), "WorkBuddy"),
                 os.path.join(os.path.expanduser("~"), "Desktop")]
    print("已连接：%s" % server)
    print("允许目录：")
    for d in allow:
        print("  - %s" % d)
    print("轮询中（Ctrl+C 退出）…\n")

    while True:
        try:
            data = get("%s/api/host/tasks?token=%s" % (server, urllib.parse.quote(token)))
            for t in data.get("tasks") or []:
                print("─" * 56)
                print("请求 #%s  %s" % (t["id"], {"list": "列目录", "read": "读取文件",
                                              "write": "写入文件"}.get(t["op"], t["op"])))
                print("路径：%s" % t["path"])
                if t["op"] == "write":
                    c = t.get("content") or ""
                    print("内容（前 200 字）：\n%s" % c[:200])
                if not within(t["path"], allow):
                    print("× 不在允许目录内，已拒绝")
                    post(server + "/api/host/report",
                         {"token": token, "id": t["id"], "ok": False, "error": "路径不在允许目录内"})
                    continue
                if not args.yes:
                    ans = input("执行吗？(y/N) ").strip().lower()
                    if ans not in ("y", "yes"):
                        post(server + "/api/host/report",
                             {"token": token, "id": t["id"], "ok": False, "error": "用户拒绝"})
                        print("已跳过\n")
                        continue
                try:
                    ok, result = do_task(t)
                    post(server + "/api/host/report",
                         {"token": token, "id": t["id"], "ok": ok, "result": result})
                    print("完成：%s\n" % (result[:200] if result else "ok"))
                except Exception as e:  # noqa: BLE001
                    post(server + "/api/host/report",
                         {"token": token, "id": t["id"], "ok": False, "error": str(e)})
                    print("失败：%s\n" % e)
        except urllib.error.HTTPError as e:
            if e.code == 401:
                print("token 失效，请重新配对（--code）")
                return 1
            print("HTTP %s" % e.code)
        except KeyboardInterrupt:
            print("\n已退出")
            return 0
        except Exception as e:  # noqa: BLE001
            print("轮询异常：%s" % e)
        time.sleep(max(2, args.interval))


if __name__ == "__main__":
    sys.exit(main())
