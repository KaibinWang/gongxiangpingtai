# -*- coding: utf-8 -*-
"""从 workbench.db 生成云端版种子数据 seed.js
- 台账数据原库即为脱敏示例，直接采用
- 素材库文件名按 _build_static.py 的映射脱敏（云端为公网可访问地址）
"""
import ast
import json
import os
import re
import sqlite3

SRC = os.path.join(os.path.dirname(os.path.abspath(__file__)), "投行工作台", "workbench.db")
BUILD = os.path.join(os.path.dirname(os.path.abspath(__file__)), "投行工作台", "_build_static.py")
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "投行工作台-云端版", "assets", "seed.js")

# 从既有脚本里取出脱敏映射，避免重复维护
RENAME, DROP_IDS, FALLBACK = {}, set(), []
if os.path.isfile(BUILD):
    tree = ast.parse(open(BUILD, encoding="utf-8").read())
    for node in tree.body:
        if isinstance(node, ast.Assign) and isinstance(node.value, ast.Dict):
            name = getattr(node.targets[0], "id", "")
            if name == "RENAME":
                RENAME = {k.value: v.value for k, v in zip(node.value.keys, node.value.values)}
            elif name == "DROP_IDS":
                DROP_IDS = {e.value for e in node.value.elts}
        if isinstance(node, ast.Assign) and getattr(node.targets[0], "id", "") == "FALLBACK":
            FALLBACK = [tuple(p.value for p in e.elts) for e in node.value.elts]


def sanitize(mid, name):
    if mid in RENAME:
        return RENAME[mid]
    s = name
    for pat, rep in FALLBACK:
        s = re.sub(pat, rep, s)
    return re.sub(r"\s{2,}", " ", s).strip(" -_（）()") or "业务材料"


con = sqlite3.connect(SRC)
con.row_factory = sqlite3.Row


def rows(sql):
    return [dict(r) for r in con.execute(sql)]


deals = rows("SELECT * FROM deals ORDER BY id")
exps = rows("SELECT * FROM expansion_clients ORDER BY id")
govs = rows("SELECT * FROM gov_projects ORDER BY id")
subs = rows("SELECT * FROM gov_subs ORDER BY id")

mats, seen = [], set()
for r in con.execute("SELECT * FROM materials ORDER BY cat, mtime DESC"):
    if r["id"] in DROP_IDS:
        continue
    if (r["kind"] or "") in ("other",):
        continue  # 图片/工程文件不入云端索引
    nm = sanitize(r["id"], r["name"])
    key = (r["cat"], nm)
    if key in seen:
        continue
    seen.add(key)
    mats.append({"id": r["id"], "name": nm, "rel": nm, "kind": r["kind"],
                 "size": r["size"], "mtime": r["mtime"], "cat": r["cat"]})

seed = {
    "deals": deals,
    "expansion_clients": exps,
    "gov_projects": govs,
    "gov_subs": subs,
    "materials": mats,
    "match_runs": [],
    "settings": {"material_root": "", "auto_scan_on": "0", "auto_scan_time": "08:30"},
}
os.makedirs(os.path.dirname(OUT), exist_ok=True)
with open(OUT, "w", encoding="utf-8") as f:
    f.write("// 由 workbench.db 生成的初始数据（素材库文件名已脱敏）\n")
    f.write("window.SEED = " + json.dumps(seed, ensure_ascii=False, indent=1) + ";\n")
print("seed.js ok:", len(deals), "deals,", len(exps), "exp,", len(govs), "gov,", len(mats), "mats")
