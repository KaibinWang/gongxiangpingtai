# -*- coding: utf-8 -*-
"""版本快照：每次迭代自动记一条版本，用于回退与功能追溯。

用法：
    python 版本快照.py -m "这轮改了什么"     # 记版本（推荐）
    python 版本快照.py --list               # 看版本记录
    python 版本快照.py -m "说明" --zip       # 同时打一份 zip 到 版本备份/

做了什么：
    1. 自动算版本号 v<日期>.<当日序号>
    2. 把「版本号 / 时间 / 说明 / 改动文件」追加到 版本记录.md
    3. git add + commit + 打标签（标签与 版本记录.md 一一对应）

回退：
    git checkout <版本号> -- .        回到某版（覆盖工作区，先提交或备份）
    git show <版本号> --stat          看那版改了什么
    git diff <版本A> <版本B>          比较两版
"""
import argparse
import datetime
import os
import subprocess
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
LOG_MD = os.path.join(ROOT, "版本记录.md")

HEADER = """# 版本记录

> 每次迭代自动追加，用于**回退**与**功能追溯**。
> 回退到某版：`git checkout <版本号> -- .`　看某版改动：`git show <版本号> --stat`
> 本文件由 `版本快照.py` 维护，请勿手工删行。

| 版本 | 时间 | 说明 | 改动文件 |
|---|---|---|---|
"""


def run(args, check=True):
    p = subprocess.run(args, cwd=ROOT, capture_output=True, text=True,
                       encoding="utf-8", errors="replace")
    if check and p.returncode != 0:
        print("执行失败：%s\n%s" % (" ".join(args), (p.stderr or "").strip()))
        sys.exit(1)
    return p.stdout.strip()


def next_tag():
    today = datetime.date.today().strftime("v%Y.%m.%d")
    seq = 0
    for t in run(["git", "tag", "-l"]).splitlines():
        if t.startswith(today + "."):
            try:
                seq = max(seq, int(t.rsplit(".", 1)[1]))
            except ValueError:
                pass
    return "%s.%d" % (today, seq + 1)


def changed_files():
    """返回本次改动文件列表（注意：不能对整段输出做 strip，否则首行状态码前的空格会被吃掉）"""
    out = subprocess.run(["git", "-c", "core.quotepath=false", "status", "--porcelain"],
                         cwd=ROOT, capture_output=True, text=True,
                         encoding="utf-8", errors="replace").stdout
    files = []
    for line in out.splitlines():
        path = line[3:].strip()
        if "->" in path:            # 重命名
            path = path.split("->")[-1].strip()
        if path:
            files.append(path.strip('"'))
    return files


def fmt_files(files, limit=4):
    if not files:
        return "（仅打标签）"
    shown = "、".join(files[:limit])
    return shown + ("、…共 %d 个" % len(files) if len(files) > limit else "")


def append_log(tag, stamp, msg, files):
    new = not os.path.isfile(LOG_MD)
    if new:
        with open(LOG_MD, "w", encoding="utf-8") as f:
            f.write(HEADER)
    with open(LOG_MD, "a", encoding="utf-8") as f:
        f.write("| **%s** | %s | %s | %s |\n" % (tag, stamp, msg.replace("|", "/"), fmt_files(files)))


def backfill():
    """首次运行时，把已有的标签补进记录表"""
    if os.path.isfile(LOG_MD):
        return
    tags = run(["git", "for-each-ref", "--sort=creatordate",
                "--format=%(refname:short)|%(creatordate:short)|%(contents:subject)", "refs/tags"])
    with open(LOG_MD, "w", encoding="utf-8") as f:
        f.write(HEADER)
        for line in tags.splitlines():
            parts = line.split("|")
            if len(parts) >= 3:
                f.write("| **%s** | %s | %s | （历史版本，无明细） |\n"
                        % (parts[0], parts[1], parts[2].replace("|", "/")))
    print("已根据 git 标签补录历史版本到 版本记录.md")


def show_list():
    if not os.path.isfile(LOG_MD):
        backfill()
    print(open(LOG_MD, encoding="utf-8").read())
    print("\n当前工作区：")
    st = run(["git", "-c", "core.quotepath=false", "status", "--short"])
    print(st if st else "  无未提交改动")


def make_zip(tag):
    import zipfile
    out_dir = os.path.join(ROOT, "版本备份")
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, "投行工作台_%s.zip" % tag)
    skip_dirs = {".venv", "__pycache__", ".pytest_cache", ".tmp", "版本备份"}
    n = 0
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        for dp, dns, fns in os.walk(ROOT):
            dns[:] = [d for d in dns if d not in skip_dirs]
            for fn in fns:
                if fn.endswith(".pyc") or fn in {"server.log", "Thumbs.db"}:
                    continue
                full = os.path.join(dp, fn)
                z.write(full, os.path.relpath(full, ROOT))
                n += 1
    return out, n


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("-m", "--message", default="", help="这轮改了什么")
    ap.add_argument("--zip", action="store_true", help="同时生成 zip")
    ap.add_argument("--list", action="store_true", help="查看版本记录")
    a = ap.parse_args()

    if not os.path.isdir(os.path.join(ROOT, ".git")):
        print("当前目录不是 git 仓库，先执行 git init")
        sys.exit(1)

    if a.list:
        show_list()
        return

    backfill()
    files = changed_files()
    tag = next_tag()
    msg = a.message or "迭代 %s" % tag
    stamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

    append_log(tag, stamp, msg, files)

    if files:
        run(["git", "add", "-A"])
        run(["git", "commit", "-m", "%s（%s）" % (msg, tag)])
    else:
        run(["git", "add", "版本记录.md"])
        run(["git", "commit", "-m", "记录版本 %s" % tag])

    run(["git", "tag", "-a", tag, "-m", msg])

    print("✅ 版本 %s 已记录" % tag)
    print("   说明：%s" % msg)
    print("   改动：%d 个文件" % len(files))
    for f in files[:8]:
        print("     · %s" % f)
    if len(files) > 8:
        print("     · …另有 %d 个" % (len(files) - 8))

    if a.zip:
        path, n = make_zip(tag)
        print("   zip：%s（%d 个文件，%.1f MB）" % (path, n, os.path.getsize(path) / 1048576))

    print("\n回退这一版：git checkout %s -- ." % tag)
    print("查看全部版本：python 版本快照.py --list")


if __name__ == "__main__":
    main()
