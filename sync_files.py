# -*- coding: utf-8 -*-
"""把素材文件夹里的真实文件同步到云端版，使云端「素材库」可下载、可作初稿参照。

两种用法：

A. 在本项目里直接跑（读 workbench.db，输出到 投行工作台-云端版）
    python sync_files.py "D:\\科技金融"

B. 在别的机器上跑（用随工具附带的 materials.json，输出到 ./out）
    python sync_files.py "D:\\科技金融"

参数：
    --materials  <json>   指定素材清单（默认先找同目录 materials.json，再找 投行工作台/workbench.db）
    --out        <dir>    输出目录（默认自动判断：项目内 → 投行工作台-云端版；否则 → ./out）
    --max-mb     <n>      单文件上限，默认 40
    --total-mb   <n>      总量上限，默认 300（加 --all 时 2000）
    --all                 索引之外的文件也一并上传（全量搬运）
    --list                只列出会同步哪些，不复制
    --clean               清空已输出文件

输出：
    <out>/assets/files/<id>.<ext>      已索引的材料
    <out>/assets/files/x_0001.<ext>    索引之外的文件（--all）
    <out>/assets/data/files.json       清单（items + extras）
"""
import argparse
import json
import os
import re
import shutil
import sqlite3
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
DB = os.path.join(BASE, "投行工作台", "workbench.db")
CLOUD = os.path.join(BASE, "投行工作台-云端版")

EXT_OK = {".docx", ".doc", ".pptx", ".ppt", ".pdf", ".xlsx", ".xls", ".csv", ".txt", ".md"}
BAD_RE = re.compile(r"app\.py\.txt|requirements\.txt|通报格式|课程报读须知|歌德语言", re.I)


def norm(s):
    """归一化：全角转半角、去空格与常见标点差异、小写"""
    out = []
    for ch in str(s or ""):
        o = ord(ch)
        out.append(chr(o - 0xFEE0) if 0xFF01 <= o <= 0xFF5E else ch)
    x = "".join(out)
    x = re.sub(r"[（）()\[\]【】\s\-_—·.。、,，]", "", x)
    return x.lower()


def load_materials(arg):
    """依次尝试：--materials 指定文件 → 本项目 workbench.db → 同目录 materials.json"""
    if arg:
        with open(arg, encoding="utf-8") as f:
            data = json.load(f)
        return data["materials"] if isinstance(data, dict) else data

    if os.path.isfile(DB):
        con = sqlite3.connect(DB)
        con.row_factory = sqlite3.Row
        rows = [dict(r) for r in con.execute("SELECT id,name,rel,kind,size FROM materials ORDER BY id")]
        con.close()
        return rows

    local = os.path.join(BASE, "materials.json")
    if os.path.isfile(local):
        with open(local, encoding="utf-8") as f:
            data = json.load(f)
        return data["materials"] if isinstance(data, dict) else data

    print("[X] 找不到素材清单：请用 --materials 指定 materials.json，或把本工具放回项目根目录")
    sys.exit(1)


def out_paths(arg_out):
    if arg_out:
        root = arg_out
    elif os.path.isdir(CLOUD):
        root = CLOUD
    else:
        root = os.path.join(BASE, "out")
    return os.path.join(root, "assets", "files"), os.path.join(root, "assets", "data", "files.json")


def scan(root):
    """归一化文件名 → 绝对路径（同名取路径最短的）"""
    index = {}
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if not d.startswith(("~$", "."))]
        for fn in filenames:
            if fn.startswith("~$"):
                continue
            stem, ext = os.path.splitext(fn)
            if ext.lower() not in EXT_OK:
                continue
            full = os.path.join(dirpath, fn)
            key = norm(stem)
            old = index.get(key)
            if old is None or len(full) < len(old):
                index[key] = full
    return index


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src", nargs="?", help="素材文件夹路径，如 D:\\科技金融")
    ap.add_argument("--materials", help="素材清单 json")
    ap.add_argument("--out", help="输出目录")
    ap.add_argument("--max-mb", type=float, default=40)
    ap.add_argument("--total-mb", type=float, default=None, help="默认 300；加 --all 时为 2000")
    ap.add_argument("--all", action="store_true", help="索引之外的文件也一并上传")
    ap.add_argument("--list", action="store_true")
    ap.add_argument("--clean", action="store_true")
    args = ap.parse_args()

    if args.total_mb is None:
        args.total_mb = 2000 if args.all else 300

    files_dir, manifest_path = out_paths(args.out)

    if args.clean:
        if os.path.isdir(files_dir):
            shutil.rmtree(files_dir)
        os.makedirs(files_dir, exist_ok=True)
        if os.path.isfile(manifest_path):
            os.remove(manifest_path)
        print("[OK] 已清空", files_dir)
        return

    if not args.src:
        print('[X] 请给出素材文件夹路径，例如：python sync_files.py "D:\\科技金融"')
        sys.exit(1)
    if not os.path.isdir(args.src):
        print("[X] 路径不存在：", args.src)
        sys.exit(1)

    mats = load_materials(args.materials)
    print("素材清单 %d 条，扫描 %s ..." % (len(mats), args.src))
    idx = scan(args.src)
    print("扫描到可用文件 %d 个" % len(idx))

    matched = []
    for m in mats:
        rel = str(m.get("rel") or m.get("name") or "").replace("\\", "/").split("/")[-1]
        hit = idx.get(norm(os.path.splitext(rel)[0]))
        if hit:
            matched.append((m, hit))
    print("文件名匹配成功 %d / %d 条\n" % (len(matched), len(mats)))

    if args.list:
        for m, p in matched:
            print("  %5s  %8.1f KB  %s" % (m["id"], os.path.getsize(p) / 1024, p))
        return

    os.makedirs(files_dir, exist_ok=True)
    manifest, skipped, total = {}, [], 0
    max_b, cap = args.max_mb * 1024 * 1024, args.total_mb * 1024 * 1024

    for m, p in matched:
        size = os.path.getsize(p)
        base = os.path.basename(p)
        if BAD_RE.search(base):
            skipped.append((m["id"], base, "无关条目"))
            continue
        if size > max_b:
            skipped.append((m["id"], base, "%.1fMB 超过单文件上限" % (size / 1048576)))
            continue
        if total + size > cap:
            skipped.append((m["id"], base, "超过总量上限"))
            continue
        ext = os.path.splitext(p)[1].lower()
        dest = "%s%s" % (m["id"], ext)
        shutil.copy2(p, os.path.join(files_dir, dest))
        manifest[str(m["id"])] = {
            "file": dest,
            "name": base,
            "size": round(size / 1024, 1),
            "kind": ext.lstrip("."),
        }
        total += size

    os.makedirs(os.path.dirname(manifest_path), exist_ok=True)

    extras = []
    if not args.all:
        _write(manifest_path, manifest, extras, total)
        _report(manifest, extras, skipped, total, mats)
        return

    # ---- 全量模式：索引之外的文件也一并上传 ----
    used = {os.path.abspath(p) for _, p in matched}
    print("全量模式：继续收集索引之外的文件 ...")
    n = 0
    for key in sorted(idx):
        p = idx[key]
        if os.path.abspath(p) in used:
            continue
        size = os.path.getsize(p)
        base = os.path.basename(p)
        if BAD_RE.search(base):
            skipped.append(("-", base, "无关条目"))
            continue
        if size > max_b:
            skipped.append(("-", base, "%.1fMB 超过单文件上限" % (size / 1048576)))
            continue
        if total + size > cap:
            skipped.append(("-", base, "超过总量上限"))
            continue
        n += 1
        ext = os.path.splitext(p)[1].lower()
        dest = "x_%04d%s" % (n, ext)
        shutil.copy2(p, os.path.join(files_dir, dest))
        rel_dir = os.path.relpath(os.path.dirname(p), args.src)
        extras.append({
            "file": dest,
            "name": base,
            "dir": "" if rel_dir == "." else rel_dir.replace("\\", "/"),
            "size": round(size / 1024, 1),
            "kind": ext.lstrip("."),
        })
        total += size

    _write(manifest_path, manifest, extras, total)
    _report(manifest, extras, skipped, total, mats)


def _write(path, items, extras, total):
    with open(path, "w", encoding="utf-8") as f:
        json.dump({
            "items": items,
            "extras": extras,
            "total_mb": round(total / 1048576, 1),
        }, f, ensure_ascii=False, indent=1)


def _report(manifest, extras, skipped, total, mats):
    print("\n[OK] 已复制 %d 个文件，合计 %.1f MB" % (len(manifest) + len(extras), total / 1048576))
    print("     已索引 %d 个 / 索引之外 %d 个" % (len(manifest), len(extras)))
    if skipped:
        print("\n[!] 跳过 %d 个：" % len(skipped))
        for i, nm, why in skipped[:15]:
            print("     %5s  %s —— %s" % (i, nm, why))
        if len(skipped) > 15:
            print("     …… 其余 %d 个" % (len(skipped) - 15))
    print("\n未上传（云端只显示索引）：%d 条" % (len(mats) - len(manifest)))


if __name__ == "__main__":
    main()
