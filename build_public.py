# -*- coding: utf-8 -*-
"""生成「公网发布版」：在云端版基础上，把数据库里素材文件的原始文件名脱敏。

云端沙盒有访问校验，公网托管是开放地址，因此发布前必须把
materials.name / rel 中的机构名、日期批次、版本号处理掉。

用法：python build_public.py [输出目录]
"""
import json
import os
import re
import shutil
import sqlite3
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(BASE, "投行工作台-云端版")
DST = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
    os.environ.get("TEMP", "/tmp"), "th_public")

RULES = [
    (r"广州分行|深圳分行|总行|分行|支行", ""),
    (r"行领导", ""),
    (r"考核", "评估"),
    (r"专项行动", "工作方案"),
    (r"新台账", "台账"),
    (r"\d{6,8}[-_]?", ""),
    (r"^\d{3,4}年?[-_]?", ""),
    (r"\((\d+)\)|（(\d+)）", ""),
    (r"[Vv]\d+", ""),
    (r"[（(]\s*[）)]", ""),
]


def mask(name):
    x = str(name or "")
    for pat, rep in RULES:
        x = re.sub(pat, rep, x)
    x = re.sub(r"\s{2,}", " ", x).strip(" -_（()）年")
    return x or "业务材料"


def main():
    if os.path.isdir(DST):
        shutil.rmtree(DST)
    shutil.copytree(SRC, DST, ignore=shutil.ignore_patterns("files"))

    db = os.path.join(DST, "assets", "data", "workbench.db")
    con = sqlite3.connect(db)
    rows = con.execute("SELECT id,name,rel FROM materials").fetchall()
    changed = 0
    for mid, name, rel in rows:
        new = mask(name)
        if new != name or rel != new:
            con.execute("UPDATE materials SET name=?, rel=? WHERE id=?", (new, new, mid))
            changed += 1
    con.commit()
    con.close()

    meta_path = os.path.join(DST, "assets", "data", "meta.json")
    with open(meta_path, encoding="utf-8") as f:
        meta = json.load(f)
    meta["version"] = str(meta.get("version", "")) + " · 公网版"
    meta["public"] = True
    with open(meta_path, "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=1)

    total = sum(os.path.getsize(os.path.join(dp, fn))
                for dp, _, fns in os.walk(DST) for fn in fns)
    print("[OK] 公网发布版已生成：", DST)
    print("     素材文件名脱敏 %d 条，体积 %.1f MB" % (changed, total / 1048576))


if __name__ == "__main__":
    main()
