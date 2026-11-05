# -*- coding: utf-8 -*-
"""把 WorkBuddy 里指向 C 盘旧项目的会话工作目录，改写为 E 盘新路径。

用法（一般由 一键切换工作区.bat 调用）：
    python 切换工作区.py            # 执行改写
    python 切换工作区.py --check    # 只看当前状态，不修改

为什么需要它：WorkBuddy 的工作区是客户端本地状态（~/.workbuddy/app/sessions.json
里的 workDir），没有对外接口；而且该文件在进程退出时会被覆写，
所以必须在客户端关闭后改写，再重新启动才生效。
"""

import json
import os
import shutil
import sys
import time

OLD = r"C:\Users\10949\WorkBuddy\投行项目"
NEW = r"E:\WorkBuddy\投行项目"
SESSIONS = os.path.join(os.path.expanduser("~"), ".workbuddy", "app", "sessions.json")


def load():
    with open(SESSIONS, encoding="utf-8") as f:
        return json.load(f)


def main():
    check_only = "--check" in sys.argv

    if not os.path.isfile(SESSIONS):
        print("[X] 找不到会话记录：%s" % SESSIONS)
        return 1
    if not os.path.isdir(NEW):
        print("[X] 目标目录不存在：%s" % NEW)
        return 1

    data = load()
    sessions = data.get("sessions", [])
    hits = [s for s in sessions if os.path.normcase(s.get("workDir", "")) == os.path.normcase(OLD)]

    print("会话总数 %d，指向 C 盘旧路径的 %d 条" % (len(sessions), len(hits)))
    for s in sessions:
        flag = "  <-- 待改" if s in hits else ""
        print("  %s  %s%s" % (s.get("conversationId", "")[:8], s.get("workDir", ""), flag))

    if check_only:
        return 0
    if not hits:
        print("已经全部指向 E 盘，无需修改。")
        return 0

    bak = SESSIONS + ".bak-%s" % time.strftime("%Y%m%d-%H%M%S")
    shutil.copy2(SESSIONS, bak)
    print("已备份原文件：%s" % bak)

    for s in hits:
        s["workDir"] = NEW
    data["updatedAt"] = time.strftime("%Y-%m-%dT%H:%M:%S.000Z", time.gmtime())

    tmp = SESSIONS + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, SESSIONS)

    print("已改写 %d 条会话的工作目录 -> %s" % (len(hits), NEW))
    return 0


if __name__ == "__main__":
    sys.exit(main())
