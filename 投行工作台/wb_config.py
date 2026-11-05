# -*- coding: utf-8 -*-
"""
投行工作台 · 配置管理 / 智能体配置 / 宿主机授权

设计要点
--------
1. 四类发送（群机器人、个人应用消息、文件、定时）共用一张任务表 + 一个后台线程，
   页面只负责建任务、看进度，不在请求线程里做网络 IO。
2. 群机器人按企业微信官方 webhook 协议发送；文件先 upload_media 拿 media_id 再发 file 消息。
3. 定时任务只存"下一次执行时间"，每分钟 tick 一次；到点转成普通任务入队，避免重复调度。
4. 智能体走 OpenAI 兼容协议（chat/completions），provider 只是预设好的 base_url + 模型清单，
   用户可下拉选也可手填，API Key 手写、落库后页面只显示掩码。
5. 宿主机授权：云端只存"待执行任务"，由用户在宿主机上运行 host_agent.py 轮询并逐条确认，
   执行完回传结果。云端永远拿不到宿主机文件句柄，也没有免确认通道。
"""

from __future__ import annotations

import csv
import datetime
import io
import json
import os
import re
import sqlite3
import threading
import time
import traceback
import urllib.error
import urllib.parse
import urllib.request
import uuid

from flask import Blueprint, Response, abort, jsonify, request, send_file

import adapter_lib as ad
import gd_crawler as gd

bp = Blueprint("cfg", __name__)

_DB = None
_PAGE = None
_FILE_DIR = None
_LOCK = threading.Lock()
_STARTED = False

WECOM_API = "https://qyapi.weixin.qq.com/cgi-bin"

# 企业微信消息长度上限（markdown 4096 字节，text 2048 字节），留点余量
MAX_MD = 3800
MAX_TEXT = 1800

PROVIDERS = {
    "deepseek": ("DeepSeek", "https://api.deepseek.com/v1",
                 ["deepseek-chat", "deepseek-reasoner"]),
    "openai": ("OpenAI", "https://api.openai.com/v1",
               ["gpt-4o-mini", "gpt-4o", "gpt-4.1-mini", "gpt-4.1"]),
    "dashscope": ("通义千问（阿里云）", "https://dashscope.aliyuncs.com/compatible-mode/v1",
                  ["qwen-plus", "qwen-max", "qwen-turbo", "qwen-long"]),
    "hunyuan": ("腾讯混元", "https://api.hunyuan.cloud.tencent.com/v1",
                ["hunyuan-turbos-latest", "hunyuan-pro", "hunyuan-lite"]),
    "moonshot": ("月之暗面 Kimi", "https://api.moonshot.cn/v1",
                 ["kimi-k2-0905-preview", "moonshot-v1-8k", "moonshot-v1-32k"]),
    "zhipu": ("智谱 GLM", "https://open.bigmodel.cn/api/paas/v4",
              ["glm-4-plus", "glm-4-air", "glm-4-flash"]),
    "custom": ("自定义（OpenAI 兼容）", "", []),
}


# ---------------------------------------------------------------------------
# 基础设施
# ---------------------------------------------------------------------------
def init(db_path, page_fn, file_dir):
    global _DB, _PAGE, _FILE_DIR
    _DB, _PAGE, _FILE_DIR = db_path, page_fn, file_dir
    os.makedirs(_FILE_DIR, exist_ok=True)
    _init_schema()
    _start_workers()


def _db():
    conn = sqlite3.connect(_DB, timeout=20)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA busy_timeout=8000")
    return conn


def q(sql, args=()):
    c = _db()
    try:
        return c.execute(sql, args).fetchall()
    finally:
        c.close()


def one(sql, args=()):
    rows = q(sql, args)
    return rows[0] if rows else None


def x(sql, args=()):
    c = _db()
    try:
        cur = c.execute(sql, args)
        c.commit()
        return cur.lastrowid
    finally:
        c.close()


def esc(v):
    return (str(v if v is not None else "")
            .replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            .replace('"', "&quot;"))


def now_str():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def sget(key, default=""):
    r = one("SELECT value FROM settings WHERE key=?", ("cfg_" + key,))
    return r["value"] if r else default


def sset(key, value):
    x("INSERT OR REPLACE INTO settings VALUES (?,?)", ("cfg_" + key, str(value)))


def _init_schema():
    c = _db()
    try:
        c.executescript("""
        CREATE TABLE IF NOT EXISTS cfg_groups (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL, webhook TEXT NOT NULL, category TEXT DEFAULT '默认',
            enabled INTEGER DEFAULT 1, note TEXT, created TEXT, updated TEXT);
        CREATE TABLE IF NOT EXISTS cfg_templates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL, content TEXT, created TEXT, updated TEXT);
        CREATE TABLE IF NOT EXISTS cfg_members (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            userid TEXT UNIQUE NOT NULL, name TEXT, dept TEXT, tags TEXT,
            enabled INTEGER DEFAULT 1, created TEXT);
        CREATE TABLE IF NOT EXISTS cfg_tasks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            kind TEXT NOT NULL, title TEXT, status TEXT DEFAULT 'queued',
            total INTEGER DEFAULT 0, ok INTEGER DEFAULT 0, fail INTEGER DEFAULT 0,
            payload TEXT, created TEXT, started TEXT, finished TEXT, note TEXT);
        CREATE TABLE IF NOT EXISTS cfg_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT, task_id INTEGER,
            target TEXT, label TEXT, status TEXT DEFAULT 'wait', msg TEXT, ts TEXT);
        CREATE TABLE IF NOT EXISTS cfg_schedules (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT, kind TEXT, payload TEXT, repeat TEXT DEFAULT 'once',
            run_at TEXT, run_time TEXT, weekday INTEGER DEFAULT 0,
            enabled INTEGER DEFAULT 1, next_run TEXT, last_run TEXT, created TEXT);
        CREATE TABLE IF NOT EXISTS agent_runs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            provider TEXT, model TEXT, prompt TEXT, sources TEXT, status TEXT,
            tokens INTEGER DEFAULT 0, cost_ms INTEGER DEFAULT 0,
            result TEXT, err TEXT, created TEXT);
        CREATE TABLE IF NOT EXISTS host_tasks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            op TEXT, path TEXT, content TEXT, status TEXT DEFAULT 'pending',
            result TEXT, err TEXT, created TEXT, finished TEXT);
        CREATE TABLE IF NOT EXISTS adapters (
            id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, host TEXT, list_url TEXT,
            spec TEXT, enabled INTEGER DEFAULT 1, created TEXT, updated TEXT);
        CREATE TABLE IF NOT EXISTS gd_runs (
            id INTEGER PRIMARY KEY AUTOINCREMENT, start TEXT, end TEXT, status TEXT,
            pages INTEGER DEFAULT 0, count INTEGER DEFAULT 0, big INTEGER DEFAULT 0,
            file TEXT, err TEXT, preview TEXT, created TEXT, finished TEXT);
        CREATE TABLE IF NOT EXISTS host_pair (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT, token TEXT, created TEXT, expires TEXT, used INTEGER DEFAULT 0);
        """)
        # 老库升级：成员表新增 webhook 列
        cols = [r[1] for r in c.execute("PRAGMA table_info(cfg_members)")]
        if "webhook" not in cols:
            c.execute("ALTER TABLE cfg_members ADD COLUMN webhook TEXT")
        gcols = [r[1] for r in c.execute("PRAGMA table_info(gd_runs)")]
        for col in ("instruction", "template", "headers", "plan"):
            if col not in gcols:
                c.execute("ALTER TABLE gd_runs ADD COLUMN %s TEXT" % col)
        c.commit()
    finally:
        c.close()


# ---------------------------------------------------------------------------
# HTTP 小工具（只用标准库，避免部署时多装依赖）
# ---------------------------------------------------------------------------
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) WorkbenchConfig/1.0"


def _post_json(url, payload, timeout=20, headers=None):
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    hd = {"Content-Type": "application/json; charset=utf-8"}
    if headers:
        hd.update(headers)
    req = urllib.request.Request(url, data=data, headers=hd, method="POST")
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read().decode("utf-8", "replace"))


def _get_json(url, timeout=20, headers=None):
    req = urllib.request.Request(url, headers=headers or {"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read().decode("utf-8", "replace"))


def _http_get_text(url, timeout=25):
    req = urllib.request.Request(url, headers={
        "User-Agent": UA,
        "Accept": "text/html,application/xhtml+xml,application/json;q=0.9,*/*;q=0.8",
        "Accept-Language": "zh-CN,zh;q=0.9",
    })
    with urllib.request.urlopen(req, timeout=timeout) as r:
        raw = r.read()
        ctype = r.headers.get("Content-Type", "")
    enc = "utf-8"
    m = re.search(r"charset=([\w-]+)", ctype, re.I)
    if m:
        enc = m.group(1)
    else:
        m = re.search(rb'charset=["\']?([\w-]+)', raw[:2048], re.I)
        if m:
            enc = m.group(1).decode("ascii", "ignore")
    for cand in (enc, "utf-8", "gb18030"):
        try:
            return raw.decode(cand, "strict")
        except (UnicodeDecodeError, LookupError):
            continue
    return raw.decode("utf-8", "replace")


def _hook_key(hook):
    try:
        return urllib.parse.parse_qs(urllib.parse.urlparse(hook).query).get("key", [""])[0]
    except Exception:
        return ""


def hook_ok(hook):
    return bool(re.match(r"^https://qyapi\.weixin\.qq\.com/cgi-bin/webhook/send\?key=[\w-]+", hook or ""))


def wecom_send(hook, msgtype, payload):
    body = {"msgtype": msgtype, msgtype: payload}
    r = _post_json(hook, body)
    return r.get("errcode", -1) == 0, r.get("errmsg", ""), r


def wecom_upload(hook, filename, blob, mtype="file"):
    """群机器人素材上传，返回 media_id（3 天内有效）。"""
    key = _hook_key(hook)
    url = "%s/webhook/upload_media?key=%s&type=%s" % (WECOM_API, key, mtype)
    boundary = "----wb" + uuid.uuid4().hex
    ct = "application/octet-stream"
    if mtype == "image":
        ct = "image/png"
    body = io.BytesIO()
    body.write(("--%s\r\n" % boundary).encode())
    body.write(('Content-Disposition: form-data; name="media"; filename="%s"; filelength=%d\r\n'
                % (filename, len(blob))).encode("utf-8"))
    body.write(("Content-Type: %s\r\n\r\n" % ct).encode())
    body.write(blob)
    body.write(("\r\n--%s--\r\n" % boundary).encode())
    req = urllib.request.Request(url, data=body.getvalue(), method="POST", headers={
        "Content-Type": "multipart/form-data; boundary=%s" % boundary})
    with urllib.request.urlopen(req, timeout=60) as r:
        d = json.loads(r.read().decode("utf-8", "replace"))
    if d.get("errcode", -1) != 0:
        raise RuntimeError("上传失败：%s" % d.get("errmsg"))
    return d["media_id"]


def wecom_app_token():
    corpid = sget("app_corpid").strip()
    secret = sget("app_secret").strip()
    if not corpid or not secret:
        raise RuntimeError("未配置企业微信应用 corpid / secret")
    cached = sget("app_token_cache")
    if cached:
        try:
            obj = json.loads(cached)
            if obj.get("expire", 0) > time.time() + 60 and obj.get("corpid") == corpid:
                return obj["token"]
        except Exception:
            pass
    url = "%s/gettoken?corpid=%s&corpsecret=%s" % (
        WECOM_API, urllib.parse.quote(corpid), urllib.parse.quote(secret))
    d = _get_json(url)
    if d.get("errcode", -1) != 0:
        raise RuntimeError("获取 access_token 失败：%s" % d.get("errmsg"))
    token = d["access_token"]
    sset("app_token_cache", json.dumps({"token": token, "corpid": corpid,
                                        "expire": time.time() + int(d.get("expires_in", 7200)) - 120}))
    return token


def app_send_text(userid, content):
    token = wecom_app_token()
    agentid = sget("app_agentid").strip() or "0"
    url = "%s/message/send?access_token=%s" % (WECOM_API, token)
    d = _post_json(url, {"touser": userid, "msgtype": "text", "agentid": int(agentid or 0),
                         "text": {"content": content}, "safe": 0})
    ok = d.get("errcode", -1) == 0
    return ok, d.get("errmsg", ""), d


def app_send_file(userid, path):
    token = wecom_app_token()
    agentid = sget("app_agentid").strip() or "0"
    up = "%s/media/upload?access_token=%s&type=file" % (WECOM_API, token)
    filename = os.path.basename(path)
    with open(path, "rb") as f:
        blob = f.read()
    boundary = "----wb" + uuid.uuid4().hex
    body = io.BytesIO()
    body.write(("--%s\r\n" % boundary).encode())
    body.write(('Content-Disposition: form-data; name="media"; filename="%s"; filelength=%d\r\n'
                % (filename, len(blob))).encode("utf-8"))
    body.write(b"Content-Type: application/octet-stream\r\n\r\n")
    body.write(blob)
    body.write(("\r\n--%s--\r\n" % boundary).encode())
    req = urllib.request.Request(up, data=body.getvalue(), method="POST", headers={
        "Content-Type": "multipart/form-data; boundary=%s" % boundary})
    with urllib.request.urlopen(req, timeout=60) as r:
        d = json.loads(r.read().decode("utf-8", "replace"))
    if d.get("errcode", -1) != 0:
        raise RuntimeError("素材上传失败：%s" % d.get("errmsg"))
    media_id = d["media_id"]
    url = "%s/message/send?access_token=%s" % (WECOM_API, token)
    d = _post_json(url, {"touser": userid, "msgtype": "file", "agentid": int(agentid or 0),
                         "file": {"media_id": media_id}})
    return d.get("errcode", -1) == 0, d.get("errmsg", ""), d


# ---------------------------------------------------------------------------
# 任务队列
# ---------------------------------------------------------------------------
def enqueue(kind, title, targets, payload):
    """targets: [(label, target)]"""
    tid = x("INSERT INTO cfg_tasks (kind,title,status,total,payload,created) VALUES (?,?,?,?,?,?)",
            (kind, title, "queued", len(targets), json.dumps(payload, ensure_ascii=False), now_str()))
    c = _db()
    try:
        c.executemany("INSERT INTO cfg_items (task_id,target,label,status) VALUES (?,?,?,?)",
                      [(tid, t, lb, "wait") for lb, t in targets])
        c.commit()
    finally:
        c.close()
    return tid


def _fetch_task():
    """原子领取一个待执行任务。"""
    c = _db()
    try:
        row = c.execute("SELECT * FROM cfg_tasks WHERE status='queued' ORDER BY id LIMIT 1").fetchone()
        if not row:
            return None
        cur = c.execute("UPDATE cfg_tasks SET status='running', started=? WHERE id=? AND status='queued'",
                        (now_str(), row["id"]))
        c.commit()
        if cur.rowcount == 0:
            return None
        return c.execute("SELECT * FROM cfg_tasks WHERE id=?", (row["id"],)).fetchone()
    finally:
        c.close()


IMAGE_EXT = (".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp")


def _send_one(kind, payload, target):
    """群 / 个人统一：有附件先发文字（若填了），再发图片或文件。"""
    content = (payload.get("content") or "").strip()
    fpath = payload.get("file") or ""
    msgtype = payload.get("msgtype", "markdown")

    if fpath and os.path.isfile(fpath):
        errors = []
        if content:
            try:
                ok, msg, _ = wecom_send(target, "markdown", {"content": content[:MAX_MD]})
                if not ok:
                    errors.append("文字：" + msg)
            except Exception as e:  # noqa: BLE001
                errors.append("文字：%s" % e)
        is_img = os.path.splitext(fpath)[1].lower() in IMAGE_EXT
        mtype = "image" if is_img else "file"
        try:
            mid = wecom_upload(target, os.path.basename(fpath), open(fpath, "rb").read(), mtype)
            ok, msg, _ = wecom_send(target, mtype, {"media_id": mid})
        except Exception as e:  # noqa: BLE001
            ok, msg = False, "%s: %s" % (type(e).__name__, e)
        if not ok:
            errors.append(("图片：" if is_img else "附件：") + msg)
        return (not errors), ("；".join(errors) if errors else "ok"), {}

    if not content:
        return False, "内容为空", {}
    if msgtype == "text":
        return wecom_send(target, "text", {"content": content[:MAX_TEXT]})
    return wecom_send(target, "markdown", {"content": content[:MAX_MD]})


def _run_task(task):
    payload = json.loads(task["payload"] or "{}")
    interval = max(0.0, float(payload.get("interval", 3)))
    retry = max(0, int(payload.get("retry", 1)))
    items = q("SELECT * FROM cfg_items WHERE task_id=? ORDER BY id", (task["id"],))
    ok = fail = 0
    for i, it in enumerate(items):
        if task["status"] == "canceled":
            break
        last = ""
        done = False
        for attempt in range(retry + 1):
            try:
                good, msg, _raw = _send_one(task["kind"], payload, it["target"])
            except Exception as e:  # noqa: BLE001
                good, msg = False, "%s: %s" % (type(e).__name__, e)
            if good:
                done = True
                break
            last = msg
            if attempt < retry:
                time.sleep(1.5)
        st = "ok" if done else "fail"
        if done:
            ok += 1
        else:
            fail += 1
        x("UPDATE cfg_items SET status=?, msg=?, ts=? WHERE id=?",
          (st, ("" if done else last)[:300], now_str(), it["id"]))
        x("UPDATE cfg_tasks SET ok=?, fail=? WHERE id=?", (ok, fail, task["id"]))
        if i < len(items) - 1 and interval:
            time.sleep(interval)
    x("UPDATE cfg_tasks SET status=?, finished=? WHERE id=?",
      ("done" if fail == 0 else ("partial" if ok else "failed"), now_str(), task["id"]))


def _worker():
    while True:
        try:
            task = _fetch_task()
            if task:
                _run_task(task)
                continue
        except Exception:  # noqa: BLE001
            traceback.print_exc()
        time.sleep(1.5)


def _next_run(sch, base=None):
    base = base or datetime.datetime.now()
    repeat = sch["repeat"] if "repeat" in sch.keys() else "once"
    if repeat == "once":
        try:
            return datetime.datetime.strptime(sch["run_at"], "%Y-%m-%d %H:%M")
        except Exception:
            return None
    try:
        hh, mm = [int(v) for v in (sch["run_time"] or "09:00").split(":")]
    except Exception:
        hh, mm = 9, 0
    cand = base.replace(hour=hh, minute=mm, second=0, microsecond=0)
    if repeat == "weekly":
        wd = int(sch["weekday"] or 0)
        delta = (wd - cand.weekday()) % 7
        cand += datetime.timedelta(days=delta)
    if cand <= base:
        cand += datetime.timedelta(days=1 if repeat != "weekly" else 7)
    return cand


def _scheduler():
    while True:
        try:
            now = datetime.datetime.now()
            for s in q("SELECT * FROM cfg_schedules WHERE enabled=1 AND next_run IS NOT NULL"):
                try:
                    due = datetime.datetime.strptime(s["next_run"], "%Y-%m-%d %H:%M")
                except Exception:
                    continue
                if due > now:
                    continue
                payload = json.loads(s["payload"] or "{}")
                enqueue(s["kind"] or "group", s["name"] or "定时发送", payload.get("targets", []), payload)
                nxt = _next_run(s, now + datetime.timedelta(minutes=1)) if s["repeat"] != "once" else None
                x("UPDATE cfg_schedules SET last_run=?, next_run=?, enabled=? WHERE id=?",
                  (now_str(), nxt.strftime("%Y-%m-%d %H:%M") if nxt else None,
                   0 if nxt is None else 1, s["id"]))
        except Exception:  # noqa: BLE001
            traceback.print_exc()
        try:
            _gd_tick()
        except Exception:  # noqa: BLE001
            traceback.print_exc()
        time.sleep(20)


def _start_workers():
    global _STARTED
    if _STARTED:
        return
    _STARTED = True
    threading.Thread(target=_worker, name="cfg-worker", daemon=True).start()
    threading.Thread(target=_scheduler, name="cfg-scheduler", daemon=True).start()


# ---------------------------------------------------------------------------
# 变量替换
# ---------------------------------------------------------------------------
VAR_RE = re.compile(r"\{(\w+)\}")


def render_vars(text, values):
    def rep(m):
        return str(values.get(m.group(1), m.group(0)))
    return VAR_RE.sub(rep, text or "")


def sample_values():
    return {"topic": "【示例】某半导体设备公司并购项目推进会",
            "time": datetime.datetime.now().strftime("%Y年%m月%d日 %H:%M"),
            "location": "分行 12 楼会议室",
            "attendees": "公司部、投行部、客户经理",
            "notes": "请提前准备授信方案与尽调清单"}


# ---------------------------------------------------------------------------
# 页面骨架
# ---------------------------------------------------------------------------
CSS = """
<style>
.cfg-tabs{display:flex;gap:2px;border-bottom:1px solid #e5e9f0;flex-wrap:wrap;margin-bottom:12px}
.cfg-tab{padding:9px 16px;font-size:13px;color:#5b6478;border:1px solid transparent;border-bottom:none;
  border-radius:8px 8px 0 0;margin-bottom:-1px;white-space:nowrap}
.cfg-tab:hover{background:#f7f9fc;color:#2456c8}
.cfg-tab.on{background:#fff;border-color:#e5e9f0;color:#2456c8;font-weight:700;box-shadow:inset 0 2px 0 #2456c8}
.cfg-bar{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:10px}
.cfg-bar .sp{flex:1}
.cfg-hint{font-size:11.5px;color:#8b95ab;line-height:1.9}
.cfg-warn{background:#fdf8ec;border:1px solid #efe0be;border-radius:8px;padding:9px 12px;
  font-size:12px;color:#7a6128;line-height:1.8;margin:10px 0}
.cfg-ok{background:#e8f6ee;border:1px solid #bfe5cd;color:#14683d;border-radius:8px;
  padding:9px 12px;font-size:12.5px;margin-bottom:10px}
.cfg-err{background:#fdecea;border:1px solid #f5c9c6;color:#a12722;border-radius:8px;
  padding:9px 12px;font-size:12.5px;margin-bottom:10px}
.cfg-mini{font-size:11.5px;color:#8b95ab}
.cfg-mask{font-family:Consolas,monospace;letter-spacing:1px}
.cfg-list{max-height:340px;overflow:auto;border:1px solid #e8ecf3;border-radius:8px}
.cfg-list table{margin:0}
.cfg-list tr.sel{background:#eef4ff}
.cfg-pre{background:#fbfcfe;border:1px solid #e8ecf3;border-radius:8px;padding:14px 16px;
  font-size:13px;line-height:1.95;color:#2b3444;white-space:pre-wrap;min-height:120px}
.cfg-pre b{color:#12224a}
.cfg-pre .mk-h{font-size:15px;font-weight:700;color:#12224a;display:block;margin:2px 0 6px}
.cfg-pre .mk-q{border-left:3px solid #c7d6f5;padding-left:10px;color:#46536b;margin:4px 0}
.cfg-pre .mk-tag{display:inline-block;background:#eaf0fd;color:#2456c8;border-radius:4px;
  padding:0 6px;font-size:11.5px;font-weight:700}
.cfg-grid2{display:grid;grid-template-columns:1fr 1fr;gap:14px}
.cfg-grid3{display:grid;grid-template-columns:repeat(3,1fr);gap:12px}
@media(max-width:1100px){.cfg-grid2{grid-template-columns:1fr}}
.cfg-prog{height:8px;background:#eef1f7;border-radius:99px;overflow:hidden;margin:8px 0}
.cfg-prog i{display:block;height:100%;background:linear-gradient(90deg,#2456c8,#4d7ee8);width:0;
  transition:width .4s}
.cfg-log{background:#0f1729;color:#c9d6ef;border-radius:8px;padding:10px 12px;font-size:11.5px;
  line-height:1.85;font-family:Consolas,monospace;max-height:200px;overflow:auto}
.cfg-num{font-family:Consolas,monospace}
.cfg-sub{font-size:12px;color:#5b6478;margin:14px 0 6px;font-weight:700}
.cfg-tag{display:inline-block;background:#eef1f7;color:#5b6478;border-radius:6px;padding:1px 8px;
  font-size:11.5px;margin-left:6px}
.cfg-file{display:flex;align-items:center;gap:8px;font-size:12.5px;color:#3d4657;
  background:#f8fafd;border:1px dashed #d7dfec;border-radius:8px;padding:8px 10px}
table.cfg td,table.cfg th{padding:8px 10px}
table.cfg tr:hover td{background:#fafbfe}
.sw{position:relative;display:inline-block;width:34px;height:18px}
.sw input{display:none}
.sw i{position:absolute;inset:0;background:#dfe4ee;border-radius:99px;transition:.2s}
.sw i:after{content:"";position:absolute;width:14px;height:14px;left:2px;top:2px;background:#fff;
  border-radius:50%;transition:.2s}
.sw input:checked + i{background:#2456c8}
.sw input:checked + i:after{transform:translateX(16px)}
.cfg-vars{display:grid;grid-template-columns:repeat(3,1fr);gap:6px 8px;margin:10px 0 8px}
.cfg-vars label{display:flex;flex-direction:column;gap:3px;font-size:11px;color:#8b95ab;margin:0}
.cfg-vars input{padding:5px 8px;font-size:12px}
.modal{position:fixed;inset:0;background:rgba(12,20,40,.45);display:none;align-items:center;
  justify-content:center;z-index:400;padding:20px}
.modal.on{display:flex}
.modal-box{width:520px;max-width:100%;background:#fff;border-radius:14px;padding:18px 20px 16px;
  box-shadow:0 24px 60px rgba(6,15,40,.35);max-height:88vh;overflow:auto}
.modal-h{display:flex;align-items:center;justify-content:space-between;margin-bottom:6px;
  padding-bottom:10px;border-bottom:1px solid #eef1f6}
.modal-h b{font-size:14.5px;color:#12224a}
.modal-x{font-size:20px;color:#8b95ab;line-height:1;text-decoration:none}
.modal-x:hover{color:#d0342c}
.modal-f{display:flex;justify-content:flex-end;gap:8px;margin-top:16px}
.ops{color:#d7dde8;margin:0 6px}
</style>


<script>
/* 统一表单提交：某些预览环境（iframe sandbox）会静默拦截原生表单提交，
   这里改用 fetch / location 自己走，保证"点了有反应"。 */
(function () {
  function toast(msg) {
    var d = document.createElement('div');
    d.textContent = msg;
    d.style.cssText = 'position:fixed;right:22px;top:18px;z-index:9999;background:#14683d;color:#fff;' +
      'padding:9px 15px;border-radius:8px;font-size:13px;box-shadow:0 8px 24px rgba(6,15,40,.22)';
    document.body.appendChild(d);
    setTimeout(function () { d.parentNode && d.parentNode.removeChild(d); }, 2200);
  }
  document.addEventListener('submit', function (e) {
    var f = e.target;
    if (!f || f.tagName !== 'FORM' || f.dataset.native === '1' || f.dataset.ajax === '1') return;
    var method = (f.getAttribute('method') || 'get').toLowerCase();
    e.preventDefault();

    if (method === 'get') {
      var qs = new URLSearchParams(new FormData(f)).toString();
      location.href = (f.getAttribute('action') || location.pathname) + '?' + qs;
      return;
    }

    var btn = f.querySelector('button:not([type=button])');
    var old = btn ? btn.textContent : '';
    if (btn) { btn.disabled = true; btn.textContent = '处理中…'; }

    fetch(f.getAttribute('action') || location.href, {
      method: 'POST', body: new FormData(f), credentials: 'same-origin', redirect: 'follow'
    }).then(function (r) {
      if (r.redirected || r.ok) toast(r.redirected ? '已保存，正在刷新…' : '已完成');
      if (r.redirected) { location.href = r.url; return; }
      if (r.ok) { location.reload(); return; }
      throw new Error('HTTP ' + r.status);
    }).catch(function (err) {
      if (btn) { btn.disabled = false; btn.textContent = old; }
      console.warn('fetch 提交失败，回退原生提交：', err);
      f.dataset.native = '1';
      f.submit();
    });
  }, true);
})();
</script>

"""


def _shell(act, title, sub, body, tab=None):
    tabs = ""
    if tab:
        items = [("groups", "群组管理"), ("template", "通知编辑"), ("send", "发送控制"),
                 ("personal", "个人发送")]
        tabs = "<div class='cfg-tabs'>" + "".join(
            "<a class='cfg-tab %s' href='/config?tab=%s'>%s</a>"
            % ("on" if tab == k else "", k, n) for k, n in items) + "</div>"
    head = ("<div class='card'><h2>%s <span class='mini'>%s</span></h2>%s%s</div>"
            % (esc(title), esc(sub), CSS, tabs))
    return _PAGE(head + body, act)


# ---------------------------------------------------------------------------
# 配置管理总入口
# ---------------------------------------------------------------------------
@bp.errorhandler(Exception)
def _cfg_error(e):
    """页面出错时给出可读的错误页，而不是白屏 500。"""
    from werkzeug.exceptions import HTTPException
    if isinstance(e, HTTPException):
        return e
    tb = traceback.format_exc()
    body = ("<div class='card'><h2>这个页面出错了</h2>"
            "<div class='cfg-err'>%s: %s</div>"
            "<div class='cfg-hint'>把下面这段发给开发即可定位，完整栈也在服务器日志里。</div>"
            "<pre class='cfg-log' style='max-height:320px'>%s</pre>"
            "<div style='margin-top:12px'><a class='btn btn3' href='/'>返回首页</a></div></div>"
            % (esc(type(e).__name__), esc(str(e)[:300]), esc(tb[-1600:])))
    return _PAGE(body, "config"), 500


@bp.route("/config")
def config_home():
    tab = request.args.get("tab", "groups")
    if tab == "template":
        return _tab_template()
    if tab == "send":
        return _tab_send()
    if tab == "personal":
        return _tab_personal()
    return _tab_groups()


# ---------------------------------------------------------------------------
# 一、群组管理
# ---------------------------------------------------------------------------
def _groups():
    return q("SELECT * FROM cfg_groups ORDER BY id")


def _flash(kind, text):
    return "<div class='cfg-%s'>%s</div>" % ("ok" if kind == "ok" else "err", esc(text))


@bp.route("/config/groups")
def groups_redirect():
    return config_home()


def _tab_groups():
    kw = (request.args.get("q") or "").strip()
    edit_id = request.args.get("edit") or ""
    edit = one("SELECT * FROM cfg_groups WHERE id=?", (edit_id,)) if edit_id else None
    show_modal = bool(edit) or request.args.get("add") == "1"

    sql, args = "SELECT * FROM cfg_groups WHERE 1=1", []
    if kw:
        sql += " AND (name LIKE ? OR webhook LIKE ? OR note LIKE ?)"
        args += ["%%%s%%" % kw] * 3
    sql += " ORDER BY id"
    rows = q(sql, args)
    stat = one("SELECT COUNT(*) c, COALESCE(SUM(enabled),0) e FROM cfg_groups")

    flash = ""
    msg = request.args.get("msg") or ""
    if msg:
        kind, _, text = msg.partition(":")
        flash = _flash("ok" if kind == "ok" else "err", text)

    trs = "".join(
        "<tr><td style='width:56px'><label class='sw'><input type='checkbox' %s "
        "onchange=\"location.href='/config/groups/toggle/%d'\"><i></i></label></td>"
        "<td><b>%s</b>%s</td>"
        "<td class='cfg-mask' style='font-size:11px'>%s</td>"
        "<td style='white-space:nowrap;text-align:right'>"
        "<a href='/config/groups/test/%d' onclick=\"return confirm('向该群发送一条测试消息？')\">测试</a>"
        "<span class='ops'>|</span><a href='/config?tab=groups&edit=%d'>编辑</a>"
        "<span class='ops'>|</span><a href='/config/groups/del/%d' "
        "onclick=\"return confirm('确认删除该群？')\" style='color:#d0342c'>删除</a>"
        "</td></tr>"
        % ("checked" if r["enabled"] else "", r["id"], esc(r["name"]),
           ("<span class='cfg-tag'>%s</span>" % esc(r["note"])) if r["note"] else "",
           esc((r["webhook"] or "")[:78]) + ("…" if len(r["webhook"] or "") > 78 else ""),
           r["id"], r["id"], r["id"])
        for r in rows) or ("<tr><td colspan='4' class='mini'>还没有群组，点右上角「＋ 添加群组」新增。</td></tr>")

    d = edit or {"id": "", "name": "", "webhook": "", "note": "", "enabled": 1}
    g = lambda k: d[k] if k in d.keys() and d[k] is not None else ""

    body = """
    <div class="card"><h2>群组台账 <span class="mini">共 %d 个 · 启用 %d 个</span></h2>
      %s
      <form method="get" class="cfg-bar">
        <input type="hidden" name="tab" value="groups">
        <a class="btn btn2 btn-sm" href="/config?tab=groups&add=1">＋ 添加群组</a>
        <input name="q" value="%s" placeholder="搜索 群名称 / webhook / 备注" style="flex:1;min-width:240px">
        <button class="btn btn3 btn-sm">搜索</button>
        <a class="btn btn3 btn-sm" href="/config?tab=groups">重置</a>
      </form>
      <div class="tbl-wrap"><table class="cfg">
        <tr><th>启用</th><th>群名称</th><th>Webhook 地址</th><th style="text-align:right">操作</th></tr>
        %s</table></div>
      <div class="cfg-hint" style="margin-top:10px">
        启用状态只影响「发送控制」的默认勾选，停用的群仍可手动选中发送。
        Webhook 来自企业微信群 → 群设置 → 群机器人 → 添加机器人。</div>
    </div>

    <div class="modal %s" id="gModal" onclick="if(event.target.id==='gModal')closeG()">
      <div class="modal-box">
        <div class="modal-h"><b>%s</b><a href="javascript:closeG()" class="modal-x">×</a></div>
        <form method="post" action="/config/groups/save" class="frm">
          <input type="hidden" name="id" value="%s">
          <label>群名称 *</label>
          <input name="name" required value="%s" placeholder="如 投行部-项目一群">
          <label>Webhook 地址 *</label>
          <input name="webhook" required value="%s"
            placeholder="https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=xxxx">
          <label>备注</label>
          <input name="note" value="%s" placeholder="如 广州分行核心客户群（选填）">
          <label>启用</label>
          <select name="enabled">
            <option value="1" %s>启用</option><option value="0" %s>停用</option>
          </select>
          <div class="modal-f">
            <a class="btn btn3" href="javascript:closeG()">取消</a>
            <button class="btn">%s</button>
          </div>
        </form>
      </div>
    </div>
    """ % (len(rows), stat["e"] or 0, flash, esc(kw), trs,
           "on" if show_modal else "",
           "编辑群组" if g("id") else "添加群组",
           g("id"), esc(g("name")), esc(g("webhook")), esc(g("note")),
           "selected" if str(d["enabled"]) == "1" else "",
           "selected" if str(d["enabled"]) == "0" else "",
           "保存修改" if g("id") else "确认添加")
    return _shell("config", "配置管理", "群组 / 模板 / 发送 / 个人消息", body, "groups")


@bp.route("/config/groups/save", methods=["POST"])
def groups_save():
    gid = request.form.get("id") or ""
    name = (request.form.get("name") or "").strip()
    hook = (request.form.get("webhook") or "").strip()
    note = (request.form.get("note") or "").strip()
    enabled = 1 if request.form.get("enabled") == "1" else 0
    if not name or not hook:
        return _back("err:群名称和 webhook 不能为空")
    if not hook_ok(hook):
        return _back("err:webhook 格式不对，应为 https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=...")
    if gid:
        x("UPDATE cfg_groups SET name=?,webhook=?,note=?,enabled=?,updated=? WHERE id=?",
          (name, hook, note, enabled, now_str(), gid))
        return _back("ok:已更新「%s」" % name)
    if one("SELECT id FROM cfg_groups WHERE webhook=?", (hook,)):
        return _back("err:该 webhook 已存在")
    x("INSERT INTO cfg_groups (name,webhook,category,note,enabled,created,updated) VALUES (?,?,?,?,?,?,?)",
      (name, hook, "默认", note, enabled, now_str(), now_str()))
    return _back("ok:已添加「%s」" % name)


def _back(msg, tab="groups"):
    base = "/agent" if tab in ("model", "fetch", "gd", "adapt") else "/config"
    return __import__("flask").redirect("%s?tab=%s&msg=%s"
                                        % (base, tab, urllib.parse.quote(msg)))


@bp.route("/config/groups/del/<int:gid>")
def groups_del(gid):
    x("DELETE FROM cfg_groups WHERE id=?", (gid,))
    return _back("ok:已删除")


@bp.route("/config/groups/toggle/<int:gid>")
def groups_toggle(gid):
    r = one("SELECT enabled,name FROM cfg_groups WHERE id=?", (gid,))
    if r:
        x("UPDATE cfg_groups SET enabled=? WHERE id=?", (0 if r["enabled"] else 1, gid))
    return _back("ok:已切换启用状态")


@bp.route("/config/groups/test/<int:gid>")
def groups_test(gid):
    r = one("SELECT * FROM cfg_groups WHERE id=?", (gid,))
    if not r:
        return _back("err:群组不存在")
    content = ("## 连接测试\n> 来自：投行部数智化工作台 · 配置管理\n> 时间：%s\n\n"
               "该群 webhook 配置正常，可以接收通知。" % now_str())
    try:
        ok, msg, _ = wecom_send(r["webhook"], "markdown", {"content": content})
    except Exception as e:  # noqa: BLE001
        return _back("err:发送失败 %s" % e)
    return _back(("ok:测试消息已发送到「%s」" % r["name"]) if ok else ("err:%s" % msg))


def template_redirect():
    return config_home()


def _tab_template():
    tpls = q("SELECT * FROM cfg_templates ORDER BY id")
    tid = request.args.get("id") or ""
    cur = one("SELECT * FROM cfg_templates WHERE id=?", (tid,)) if tid else None
    content = cur["content"] if cur else (
        "## 【{topic}】\n"
        "> 会议时间：{time}\n"
        "> 会议地点：{location}\n"
        "> 参会人员：{attendees}\n\n"
        "{notes}\n\n"
        "<font color=\"comment\">投行部 · 数智化工作台</font>")
    groups = [r for r in _groups() if r["enabled"]]
    flash = ""
    msg = request.args.get("msg") or ""
    if msg:
        kind, _, text = msg.partition(":")
        flash = _flash("ok" if kind == "ok" else "err", text)

    vars_html = "".join(
        "<a class='btn btn3 btn-sm' href=\"javascript:insVar('{%s}')\">{%s} %s</a>"
        % (k, k, label) for k, label in
        [("topic", "会议主题"), ("time", "时间"), ("location", "地点"),
         ("attendees", "参会人员"), ("notes", "备注")])

    body = """
    <div class="card"><h2>消息模板 <span class="mini">支持企业微信 Markdown</span></h2>
      %s
      <form method="get" class="cfg-bar">
        <input type="hidden" name="tab" value="template">
        <span class="cfg-mini">选择模板</span>
        <select name="id" onchange="this.form.submit()" style="width:240px">
          <option value="">+ 新建模板</option>
          %s
        </select>
        <a class="btn btn3 btn-sm" href="/config?tab=template">新建</a>
      </form>
      <div class="cfg-grid2">
        <div>
          <form method="post" action="/config/template/save">
            <input type="hidden" name="id" value="%s">
            <label class="cfg-mini">模板名称</label>
            <input name="name" value="%s" placeholder="如 会议通知模板" required>
            <div class="cfg-bar" style="margin-top:10px">
              <span class="cfg-mini">插入变量</span>%s
            </div>
            <textarea id="tplContent" name="content" rows="16" oninput="refreshPreview()"
              style="font-family:Consolas,Microsoft YaHei,monospace;line-height:1.9">%s</textarea>
            <div class="cfg-bar" style="margin-top:10px">
              <button class="btn">保存模板</button>
              <a class="btn btn3" href="/config?tab=template">新建</a>
              %s
            </div>
          </form>
        </div>
        <div>
          <div class="cfg-mini">预览效果（变量按示例值填充）</div>
          <div class="cfg-pre" id="preview"></div>
          <div class="cfg-sub">试发到群</div>
          <form method="post" action="/config/template/test" class="cfg-bar">
            <input type="hidden" name="content" id="testContent">
            <select name="gid" style="flex:1">
              %s
            </select>
            <button class="btn btn2 btn-sm">发送测试</button>
          </form>
          <div class="cfg-hint">测试消息会真的发到所选群，建议先用测试群。</div>
        </div>
      </div>
    </div>
    <script>
    var SAMPLE = %s;
    function insVar(v){var t=document.getElementById('tplContent');var s=t.selectionStart,e=t.selectionEnd;
      t.value=t.value.slice(0,s)+v+t.value.slice(e);t.focus();t.selectionStart=t.selectionEnd=s+v.length;refreshPreview();}
    function md2html(s){
      s=s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
      s=s.replace(/&lt;font color="(\\w+)"&gt;([\\s\\S]*?)&lt;\\/font&gt;/g,'<span style="color:#8b95ab">$2</span>');
      var lines=s.split('\\n'),out=[];
      for(var i=0;i<lines.length;i++){var L=lines[i];
        if(/^###\\s+/.test(L)){out.push('<b>'+L.replace(/^###\\s+/,'')+'</b>');continue;}
        if(/^##\\s+/.test(L)){out.push('<span class="mk-h">'+L.replace(/^##\\s+/,'')+'</span>');continue;}
        if(/^#\\s+/.test(L)){out.push('<span class="mk-h">'+L.replace(/^#\\s+/,'')+'</span>');continue;}
        if(/^&gt;\\s?/.test(L)){out.push('<div class="mk-q">'+L.replace(/^&gt;\\s?/,'')+'</div>');continue;}
        L=L.replace(/\\*\\*(.+?)\\*\\*/g,'<b>$1</b>');
        L=L.replace(/`(.+?)`/g,'<span class="mk-tag">$1</span>');
        out.push(L);
      }
      return out.join('\\n');
    }
    function refreshPreview(){
      var v=document.getElementById('tplContent').value;
      var r=v.replace(/\\{(\\w+)\\}/g,function(m,k){return SAMPLE[k]!==undefined?SAMPLE[k]:m;});
      document.getElementById('preview').innerHTML=md2html(r);
      document.getElementById('testContent').value=r;
    }
    refreshPreview();
    </script>
    """ % (flash,
           "".join("<option value='%d' %s>%s</option>" % (t["id"], "selected" if cur and cur["id"] == t["id"] else "", esc(t["name"]))
                   for t in tpls),
           cur["id"] if cur else "", esc(cur["name"] if cur else ""), vars_html, esc(content),
           ("<a class='btn btn3' href='/config/template/del/%d' onclick=\"return confirm('确认删除该模板？')\" "
            "style='color:#d0342c;margin-left:8px'>删除</a>" % cur["id"]) if cur else "",
           "".join("<option value='%d'>%s</option>" % (g["id"], esc(g["name"])) for g in groups) or "<option value=''>（暂无启用群组）</option>",
           json.dumps(sample_values(), ensure_ascii=False))
    return _shell("config", "配置管理", "群组 / 模板 / 发送 / 个人消息", body, "template")


@bp.route("/config/template/save", methods=["POST"])
def template_save():
    tid = request.form.get("id") or ""
    name = (request.form.get("name") or "").strip()
    content = request.form.get("content") or ""
    if not name:
        return _back("err:模板名称不能为空", "template")
    if tid:
        x("UPDATE cfg_templates SET name=?,content=?,updated=? WHERE id=?", (name, content, now_str(), tid))
        return _back("ok:模板已保存", "template")
    nid = x("INSERT INTO cfg_templates (name,content,created,updated) VALUES (?,?,?,?)",
            (name, content, now_str(), now_str()))
    return __import__("flask").redirect("/config?tab=template&id=%d" % nid)


@bp.route("/config/template/del/<int:tid>")
def template_del(tid):
    x("DELETE FROM cfg_templates WHERE id=?", (tid,))
    return _back("ok:模板已删除", "template")


@bp.route("/config/template/test", methods=["POST"])
def template_test():
    gid = request.form.get("gid") or ""
    content = request.form.get("content") or ""
    g = one("SELECT * FROM cfg_groups WHERE id=?", (gid,))
    if not g:
        return _back("err:请选择目标群", "template")
    try:
        ok, msg, _ = wecom_send(g["webhook"], "markdown", {"content": content[:MAX_MD]})
    except Exception as e:  # noqa: BLE001
        return _back("err:发送失败 %s" % e, "template")
    return _back(("ok:已发送到「%s」" % g["name"]) if ok else ("err:%s" % msg), "template")


# ---------------------------------------------------------------------------
# 三、发送控制
# ---------------------------------------------------------------------------
@bp.route("/config/send")
def send_redirect():
    return config_home()


def _save_upload(f):
    if not f or not f.filename:
        return ""
    safe = re.sub(r"[^\w.\-\u4e00-\u9fa5]", "_", f.filename)
    path = os.path.join(_FILE_DIR, "%s_%s" % (uuid.uuid4().hex[:8], safe))
    f.save(path)
    return path


def _tab_send():
    groups = _groups()
    tpls = q("SELECT * FROM cfg_templates ORDER BY id")
    tasks = q("SELECT * FROM cfg_tasks WHERE kind='group' ORDER BY id DESC LIMIT 12")
    sches = q("SELECT * FROM cfg_schedules ORDER BY id DESC")
    flash = ""
    msg = request.args.get("msg") or ""
    if msg:
        kind, _, text = msg.partition(":")
        flash = _flash("ok" if kind == "ok" else "err", text)

    gtrs = "".join(
        "<tr><td style='width:34px'><input type='checkbox' class='gck' value='%d' %s "
        "data-name='%s' style='width:auto'></td>"
        "<td>%s</td><td class='mini'>%s</td></tr>"
        % (g["id"], "checked" if g["enabled"] else "", esc(g["name"]), esc(g["name"]),
           "已停用" if not g["enabled"] else "")
        for g in groups) or "<tr><td colspan='3' class='mini'>暂无群组，请先到「群组管理」添加。</td></tr>"

    task_rows = "".join(
        "<tr><td class='cfg-num'>%d</td><td>%s</td><td>%s</td><td class='cfg-num'>%d/%d/%d</td>"
        "<td class='mini'>%s</td><td class='mini'>%s</td></tr>"
        % (t["id"], esc(t["title"] or ""), _status_pill(t["status"]), t["ok"] or 0, t["fail"] or 0,
           t["total"] or 0, esc(t["created"] or ""), esc(t["finished"] or "—"))
        for t in tasks) or "<tr><td colspan='6' class='mini'>还没有发送任务</td></tr>"

    sch_rows = "".join(
        "<tr><td>%s</td><td>%s</td><td class='mini'>%s</td><td class='mini'>%s</td>"
        "<td>%s</td><td class='mini'>%s</td>"
        "<td style='white-space:nowrap'><a href='/config/schedule/toggle/%d'>%s</a> "
        "<a href='/config/schedule/del/%d' onclick=\"return confirm('确认删除该定时任务？')\" "
        "style='color:#d0342c'>删除</a></td></tr>"
        % (esc(s["name"]), _repeat_text(s), esc(s["next_run"] or "—"), esc(s["last_run"] or "—"),
           "<span class='pill p2'>启用</span>" if s["enabled"] else "<span class='pill p6'>停用</span>",
           esc(s["created"] or ""), s["id"], "停用" if s["enabled"] else "启用", s["id"])
        for s in sches) or "<tr><td colspan='7' class='mini'>还没有定时任务</td></tr>"

    body = """
    <div class="card"><h2>发送控制 <span class="mini">群机器人批量发送</span></h2>
      %s
      <div class="cfg-grid2">
        <div>
          <div class="cfg-bar"><b class="cfg-mini">目标群组</b><span class="sp"></span>
            <a class="btn btn3 btn-sm" href="javascript:pickAll(1)">全选</a>
            <a class="btn btn3 btn-sm" href="javascript:pickAll(0)">全不选</a>
            <span class="cfg-mini" id="pickInfo">已选 0</span>
          </div>
          <div class="cfg-list"><table class="cfg">%s</table></div>
        </div>
        <div>
          <form method="post" action="/config/send/submit" enctype="multipart/form-data" data-ajax="1">
            <div class="cfg-bar" style="margin-bottom:6px">
              <b style="font-size:12.5px;color:#3d4657">消息内容</b><span class="sp"></span>
              <select id="tplPick" onchange="loadTpl(this)" style="width:170px">
                <option value="">自定义内容</option>
                %s
              </select>
            </div>
            <div class="cfg-vars" title="留空的变量会保留 {xxx} 原样">
              <label>主题<input id="v_topic" placeholder="会议主题"></label>
              <label>时间<input id="v_time" value="%s"></label>
              <label>地点<input id="v_location" placeholder="会议地点"></label>
              <label>参会<input id="v_attendees" placeholder="参会人员"></label>
              <label>备注<input id="v_notes" placeholder="备注"></label>
            </div>
            <textarea id="sendContent" name="content" rows="8"
              placeholder="支持企业微信 Markdown，可用 {topic}{time}{location}{attendees}{notes} 变量"
              style="font-family:Consolas,Microsoft YaHei,monospace;line-height:1.9"></textarea>
            <div class="cfg-bar" style="margin-top:8px">
              <span class="cfg-mini">消息格式</span>
              <select name="msgtype" style="width:120px">
                <option value="markdown">Markdown</option><option value="text">纯文本</option>
              </select>
              <span class="cfg-mini">间隔(秒)</span><input name="interval" type="number" value="3" min="0" step="0.5" style="width:80px">
              <span class="cfg-mini">重试</span><input name="retry" type="number" value="2" min="0" max="5" style="width:64px">
            </div>
            <div class="cfg-bar" style="margin-top:8px">
              <span class="cfg-mini">附件（可选，与正文同时发则先发文字再发附件）</span>
              <input type="file" name="file" style="max-width:250px">
            </div>
            <input type="hidden" name="gids" id="gids">
            <div style="margin-top:12px">
              <button class="btn" style="width:100%%">▶ 提交发送任务</button>
            </div>
          </form>
          <div class="cfg-prog"><i id="progBar"></i></div>
          <div class="cfg-mini" id="progText">尚未提交任务</div>
          <div class="cfg-log" id="logBox" style="margin-top:8px">等待提交…</div>
        </div>
      </div>
    </div>

    <div class="card"><h2>我的任务</h2>
      <div class="tbl-wrap"><table class="cfg">
        <tr><th>ID</th><th>标题</th><th>状态</th><th>成功/失败/总数</th><th>提交时间</th><th>完成时间</th></tr>
        %s</table></div>
    </div>

    <div class="card"><h2>定时任务 <span class="mini">到点自动转成发送任务</span></h2>
      <form method="post" action="/config/schedule/save">
        <div class="row">
          <div style="flex:2"><label>任务名称</label><input name="name" placeholder="如 每周一晨会通知" required></div>
          <div><label>重复</label><select name="repeat" id="repSel" onchange="repChange()">
            <option value="once">一次性</option><option value="daily">每天</option><option value="weekly">每周</option>
          </select></div>
          <div id="onceBox"><label>执行时间</label><input type="datetime-local" name="run_at"></div>
          <div id="timeBox" style="display:none"><label>时间</label><input type="time" name="run_time" value="09:00"></div>
          <div id="weekBox" style="display:none"><label>星期</label><select name="weekday">
            <option value="0">周一</option><option value="1">周二</option><option value="2">周三</option>
            <option value="3">周四</option><option value="4">周五</option><option value="5">周六</option>
            <option value="6">周日</option></select></div>
        </div>
        <label>消息内容</label>
        <textarea name="content" rows="4" placeholder="留空则使用下方选择的模板"></textarea>
        <div class="row" style="margin-top:8px">
          <div><label>使用模板</label><select name="tpl_id"><option value="">不套用模板</option>%s</select></div>
          <div><label>间隔(秒)</label><input name="interval" type="number" value="3" min="0" step="0.5"></div>
        </div>
        <div class="cfg-hint" style="margin-top:8px">目标群沿用上方勾选的群组，请在提交前先勾好。</div>
        <div style="margin-top:12px"><button class="btn btn2">保存定时任务</button></div>
      </form>
      <div class="tbl-wrap" style="margin-top:14px"><table class="cfg">
        <tr><th>名称</th><th>重复</th><th>下次执行</th><th>上次执行</th><th>状态</th><th>创建</th><th>操作</th></tr>
        %s</table></div>
    </div>

    <script>
    var TPLS = %s;
    function pickAll(v){document.querySelectorAll('.gck').forEach(function(c){c.checked=!!v});syncPick();}
    function syncPick(){var n=document.querySelectorAll('.gck:checked').length;
      document.getElementById('pickInfo').textContent='已选 '+n;
      document.getElementById('gids').value=Array.prototype.map.call(document.querySelectorAll('.gck:checked'),
        function(c){return c.value}).join(',');}
    document.querySelectorAll('.gck').forEach(function(c){c.addEventListener('change',syncPick)});syncPick();
    function loadTpl(sel){var t=TPLS[sel.value];if(t)document.getElementById('sendContent').value=t;}
    function repChange(){var v=document.getElementById('repSel').value;
      document.getElementById('onceBox').style.display=(v==='once')?'':'none';
      document.getElementById('timeBox').style.display=(v==='once')?'none':'';
      document.getElementById('weekBox').style.display=(v==='weekly')?'':'none';}
    repChange();
    var FORM=document.querySelector('form[action="/config/send/submit"]');
    FORM.addEventListener('submit',function(ev){
      if(!document.getElementById('gids').value){alert('请先勾选目标群');ev.preventDefault();return;}
      ev.preventDefault();
      var picked=[].slice.call(document.querySelectorAll('.gck:checked'));
      var names=picked.map(function(c){return c.getAttribute('data-name')}).join('、');
      if(!confirm('确认向 '+picked.length+' 个群发送？

'+names)){return;}
      var ta=document.getElementById('sendContent');
      ['topic','time','location','attendees','notes'].forEach(function(k){
        var el=document.getElementById('v_'+k); if(!el||!el.value)return;
        ta.value=ta.value.split('{'+k+'}').join(el.value);
      });
      var fd=new FormData(FORM);
      fetch('/config/send/submit',{method:'POST',body:fd}).then(function(r){return r.json()}).then(function(d){
        if(d.error){alert(d.error);return;}
        document.getElementById('logBox').textContent='任务 #'+d.task_id+' 已提交，开始轮询进度…';
        poll(d.task_id);
      }).catch(function(e){alert('提交失败：'+e)});
    });
    var TIMER=null;
    function poll(id){
      if(TIMER)clearInterval(TIMER);
      TIMER=setInterval(function(){
        fetch('/config/task/'+id+'/json').then(function(r){return r.json()}).then(function(d){
          var done=d.ok+d.fail, pct=d.total?Math.round(done*100/d.total):0;
          document.getElementById('progBar').style.width=pct+'%%';
          document.getElementById('progText').textContent='进度 '+done+'/'+d.total+'　成功 '+d.ok+'　失败 '+d.fail+'　状态 '+d.status;
          var lines=d.items.map(function(it){return (it.status==='ok'?'✔ ':(it.status==='fail'?'✘ ':'· '))+it.label+(it.msg?' — '+it.msg:'')});
          document.getElementById('logBox').textContent=lines.join('\\n');
          if(d.status!=='queued'&&d.status!=='running'){clearInterval(TIMER);TIMER=null;}
        });
      },1200);
    }
    </script>
    """ % (flash, gtrs,
           "".join("<option value='%d'>%s</option>" % (t["id"], esc(t["name"])) for t in tpls),
           datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
           task_rows,
           "".join("<option value='%d'>%s</option>" % (t["id"], esc(t["name"])) for t in tpls),
           sch_rows,
           json.dumps({str(t["id"]): t["content"] for t in tpls}, ensure_ascii=False))
    return _shell("config", "配置管理", "群组 / 模板 / 发送 / 个人消息", body, "send")


def _status_pill(st):
    m = {"queued": ("p3", "排队中"), "running": ("p1", "发送中"), "done": ("p2", "已完成"),
         "partial": ("p3", "部分失败"), "failed": ("p4", "失败"), "canceled": ("p6", "已取消")}
    cls, txt = m.get(st, ("p6", st or "-"))
    return "<span class='pill %s'>%s</span>" % (cls, txt)


def _repeat_text(s):
    if s["repeat"] == "once":
        return "一次性 %s" % (s["run_at"] or "")
    if s["repeat"] == "weekly":
        wd = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"][int(s["weekday"] or 0) % 7]
        return "每周%s %s" % (wd, s["run_time"] or "")
    return "每天 %s" % (s["run_time"] or "")


@bp.route("/config/send/submit", methods=["POST"])
def send_submit():
    gids = [v for v in (request.form.get("gids") or "").split(",") if v.strip()]
    if not gids:
        return jsonify({"error": "请先勾选目标群"})
    content = request.form.get("content") or ""
    msgtype = request.form.get("msgtype") or "markdown"
    interval = request.form.get("interval") or "3"
    retry = request.form.get("retry") or "2"
    fpath = _save_upload(request.files.get("file"))

    if fpath:
        ext = os.path.splitext(fpath)[1].lower()
        payload = {"msgtype": "image" if ext in IMAGE_EXT else "file",
                   "file": fpath, "content": content,
                   "interval": interval, "retry": retry}
        title = "发送附件 · %s" % os.path.basename(fpath).split("_", 1)[-1]
    else:
        if not content.strip():
            return jsonify({"error": "消息内容不能为空"})
        payload = {"msgtype": msgtype, "content": content, "interval": interval, "retry": retry}
        title = ("发送 Markdown · " if msgtype == "markdown" else "发送文本 · ") + content.strip().replace("\n", " ")[:24]

    targets = []
    for gid in gids:
        g = one("SELECT * FROM cfg_groups WHERE id=?", (gid,))
        if g:
            targets.append((g["name"], g["webhook"]))
    if not targets:
        return jsonify({"error": "目标群不存在"})
    tid = enqueue("group", title, targets, payload)
    return jsonify({"task_id": tid})


@bp.route("/config/task/<int:tid>/json")
def task_json(tid):
    t = one("SELECT * FROM cfg_tasks WHERE id=?", (tid,))
    if not t:
        return jsonify({"error": "任务不存在"})
    items = q("SELECT * FROM cfg_items WHERE task_id=? ORDER BY id", (tid,))
    return jsonify({"id": t["id"], "status": t["status"], "total": t["total"],
                    "ok": t["ok"], "fail": t["fail"],
                    "items": [{"label": i["label"], "status": i["status"], "msg": i["msg"]} for i in items]})


@bp.route("/config/schedule/save", methods=["POST"])
def schedule_save():
    name = (request.form.get("name") or "").strip()
    _v = {k: (request.form.get("v_" + k) or "").strip()
          for k in ("topic", "time", "location", "attendees", "notes")}
    _v["time"] = _v["time"] or datetime.datetime.now().strftime("%Y年%m月%d日 %H:%M")
    repeat = request.form.get("repeat") or "once"
    content = request.form.get("content") or ""
    tpl_id = request.form.get("tpl_id") or ""
    if tpl_id:
        t = one("SELECT content FROM cfg_templates WHERE id=?", (tpl_id,))
        if t:
            content = t["content"]
    content = render_vars(content, _v)
    gids = request.form.get("gids") or ""
    # 目标群：优先取页面勾选（由前端随表单提交），否则取全部启用群
    targets = []
    for gid in [g for g in gids.split(",") if g.strip()]:
        g = one("SELECT name,webhook FROM cfg_groups WHERE id=?", (gid,))
        if g:
            targets.append([g["name"], g["webhook"]])
    if not targets:
        targets = [[g["name"], g["webhook"]] for g in _groups() if g["enabled"]]
    payload = {"msgtype": "markdown", "content": content, "interval": request.form.get("interval") or "3",
               "retry": 2, "targets": targets}
    sch = {"repeat": repeat, "run_at": request.form.get("run_at") or "",
           "run_time": request.form.get("run_time") or "09:00",
           "weekday": request.form.get("weekday") or 0}
    nxt = _next_run(sch)
    if not nxt:
        return _back("err:请填写有效的执行时间", "send")
    x("INSERT INTO cfg_schedules (name,kind,payload,repeat,run_at,run_time,weekday,enabled,next_run,created) "
      "VALUES (?,?,?,?,?,?,?,1,?,?)",
      (name or "定时发送", "group", json.dumps(payload, ensure_ascii=False), repeat,
       sch["run_at"], sch["run_time"], sch["weekday"], nxt.strftime("%Y-%m-%d %H:%M"), now_str()))
    return _back("ok:定时任务已保存，下次执行 %s" % nxt.strftime("%m-%d %H:%M"), "send")


@bp.route("/config/schedule/toggle/<int:sid>")
def schedule_toggle(sid):
    s = one("SELECT * FROM cfg_schedules WHERE id=?", (sid,))
    if not s:
        return _back("err:任务不存在", "send")
    if s["enabled"]:
        x("UPDATE cfg_schedules SET enabled=0 WHERE id=?", (sid,))
        return _back("ok:已停用", "send")
    nxt = _next_run(s)
    x("UPDATE cfg_schedules SET enabled=1, next_run=? WHERE id=?",
      (nxt.strftime("%Y-%m-%d %H:%M") if nxt else None, sid))
    return _back("ok:已启用", "send")


@bp.route("/config/schedule/del/<int:sid>")
def schedule_del(sid):
    x("DELETE FROM cfg_schedules WHERE id=?", (sid,))
    return _back("ok:定时任务已删除", "send")


# ---------------------------------------------------------------------------
# 四、个人发送
# ---------------------------------------------------------------------------
@bp.route("/config/personal")
def personal_redirect():
    return config_home()


def _tab_personal():
    members = q("SELECT * FROM cfg_members ORDER BY id")
    tasks = q("SELECT * FROM cfg_tasks WHERE kind='personal' ORDER BY id DESC LIMIT 10")
    flash = ""
    msg = request.args.get("msg") or ""
    if msg:
        kind, _, text = msg.partition(":")
        flash = _flash("ok" if kind == "ok" else "err", text)

    mrows = "".join(
        "<tr><td style='width:34px'><input type='checkbox' class='mck' value='%d' "
        "data-name='%s' style='width:auto'></td>"
        "<td><b>%s</b></td>"
        "<td class='cfg-mask' style='font-size:11px'>%s</td>"
        "<td style='text-align:right;white-space:nowrap'>"
        "<a href='/config/personal/member_del/%d' onclick=\"return confirm('删除该成员？')\" "
        "style='color:#d0342c'>删除</a></td></tr>"
        % (m["id"], esc(m["name"] or ""), esc(m["name"] or "（未命名）"),
           esc(((m["webhook"] if "webhook" in m.keys() else "") or "")[:78]),
           m["id"])
        for m in members) or "<tr><td colspan='4' class='mini'>还没有成员，先在上方添加「姓名 + Webhook」。</td></tr>"

    task_rows = "".join(
        "<tr><td class='cfg-num'>%d</td><td>%s</td><td>%s</td><td class='cfg-num'>%d/%d/%d</td>"
        "<td class='mini'>%s</td></tr>"
        % (t["id"], esc(t["title"] or ""), _status_pill(t["status"]), t["ok"] or 0, t["fail"] or 0,
           t["total"] or 0, esc(t["created"] or ""))
        for t in tasks) or "<tr><td colspan='5' class='mini'>还没有个人消息任务</td></tr>"

    body = """
    <div class="card"><h2>成员 <span class="mini">共 %d 人 · 每人绑定一个群机器人 Webhook</span></h2>
      %s
      <form method="post" action="/config/personal/member_add" class="cfg-bar">
        <input name="name" placeholder="姓名" style="width:130px" required>
        <input name="webhook" placeholder="https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=xxxx"
          style="flex:1;min-width:320px" required>
        <button class="btn btn2 btn-sm">＋ 添加成员</button>
      </form>
      <div class="tbl-wrap"><table class="cfg">
        <tr><th style="width:34px"></th><th>姓名</th><th>Webhook 地址</th><th style="text-align:right">操作</th></tr>
        %s</table></div>
      <div class="cfg-bar" style="margin-top:8px">
        <a class="btn btn3 btn-sm" href="javascript:pickMembers(1)">全选</a>
        <a class="btn btn3 btn-sm" href="javascript:pickMembers(0)">全不选</a>
        <span class="cfg-mini" id="mInfo">已选 0</span>
      </div>
      <div class="cfg-hint">每个成员的 webhook 可以指向不同群，也可以指向同一个群。发送时按人逐个投递。</div>
    </div>

    <div class="card"><h2>发送内容</h2>
      <form method="post" action="/config/personal/submit" enctype="multipart/form-data" data-ajax="1">
        <textarea name="content" rows="6"
          placeholder="支持企业微信 Markdown，可用 {topic}{time}{location}{attendees}{notes} 变量；留空则只发附件"></textarea>
        <div class="cfg-bar" style="margin-top:10px">
          <span class="cfg-mini">附件（图片 / 文件，可只发附件）</span>
          <input type="file" name="file" style="max-width:260px">
          <span class="cfg-mini">间隔(秒)</span>
          <input name="interval" type="number" value="2" min="0" step="0.5" style="width:80px">
          <button class="btn" style="margin-left:auto">▶ 提交发送</button>
        </div>
        <input type="hidden" name="mids" id="mids">
      </form>
      <div class="cfg-hint" style="margin-top:6px">
        文字和附件都填了的话，先发文字再发附件，共两条消息。图片按图片消息发送，其它按文件发送（≤20MB）。</div>
      <div class="tbl-wrap" style="margin-top:12px"><table class="cfg">
        <tr><th>ID</th><th>标题</th><th>状态</th><th>成功/失败/总数</th><th>提交时间</th></tr>
        %s</table></div>
    </div>

    <script>
    function pickMembers(v){document.querySelectorAll('.mck').forEach(function(c){c.checked=!!v});syncM();}
    function syncM(){var n=document.querySelectorAll('.mck:checked').length;
      document.getElementById('mInfo').textContent='已选 '+n;
      document.getElementById('mids').value=[].slice.call(document.querySelectorAll('.mck:checked'))
        .map(function(c){return c.value}).join(',');}
    document.querySelectorAll('.mck').forEach(function(c){c.addEventListener('change',syncM)});syncM();
    document.querySelector('form[action="/config/personal/submit"]').addEventListener('submit',function(ev){
      var n=document.querySelectorAll('.mck:checked').length;
      if(!n){alert('请先勾选成员');ev.preventDefault();return;}
      var names=[].slice.call(document.querySelectorAll('.mck:checked'))
        .map(function(c){return c.getAttribute('data-name')}).join('、');
      if(!confirm('确认向 '+n+' 位成员发送？\\n\\n'+names)){ev.preventDefault();}
    });
    </script>
    """ % (len(members), flash, mrows, task_rows)
    return _shell("config", "配置管理", "群组 / 模板 / 发送 / 个人消息", body, "personal")


@bp.route("/config/personal/member_add", methods=["POST"])
def member_add():
    name = (request.form.get("name") or "").strip()
    hook = (request.form.get("webhook") or "").strip()
    if not name:
        return _back("err:姓名不能为空", "personal")
    if not hook_ok(hook):
        return _back("err:Webhook 地址格式不对，应为 https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=...",
                     "personal")
    if one("SELECT id FROM cfg_members WHERE name=? AND webhook=?", (name, hook)):
        return _back("err:该成员已存在", "personal")
    x("INSERT INTO cfg_members (userid,name,webhook,enabled,created) VALUES (?,?,?,1,?)",
      ("m" + uuid.uuid4().hex[:12], name, hook, now_str()))
    return _back("ok:已添加「%s」" % name, "personal")


@bp.route("/config/personal/member_del/<int:mid>")
def member_del(mid):
    x("DELETE FROM cfg_members WHERE id=?", (mid,))
    return _back("ok:已删除", "personal")


@bp.route("/config/personal/submit", methods=["POST"])
def personal_submit():
    mids = [v for v in (request.form.get("mids") or "").split(",") if v.strip()]
    if not mids:
        return _back("err:请先勾选成员", "personal")
    content = request.form.get("content") or ""
    _v = {k: (request.form.get("v_" + k) or "").strip()
          for k in ("topic", "time", "location", "attendees", "notes")}
    content = render_vars(content, _v)
    fpath = _save_upload(request.files.get("file"))
    if not fpath and not content.strip():
        return _back("err:内容和附件至少填一项", "personal")

    payload = {"msgtype": "markdown", "file": fpath, "content": content,
               "interval": request.form.get("interval") or "2", "retry": 1, "mode": "hook"}
    targets, names = [], []
    for mid in mids:
        m = one("SELECT * FROM cfg_members WHERE id=?", (mid,))
        if not m:
            continue
        hook = (m["webhook"] if "webhook" in m.keys() else "") or ""
        if not hook_ok(hook):
            names.append((m["name"] or "?") + "(缺 webhook)")
            continue
        names.append(m["name"] or "?")
        targets.append((m["name"] or "?", hook))
    if not targets:
        return _back("err:所选成员都没有配置有效的 Webhook", "personal")
    title = ("发送文件 · " if fpath else "个人消息 · ") + (content.strip().replace("\n", " ")[:20] or "附件")
    enqueue("personal", title, targets, payload)
    extra = ("，跳过 %d 人（缺 webhook）" % (len(names) - len(targets))) if len(names) > len(targets) else ""
    return _back("ok:已提交 %d 人的发送任务%s" % (len(targets), extra), "personal")


AGENT_TABS = [("model", "模型配置"), ("fetch", "智能抓取"), ("gd", "专项抓取")]

DEFAULT_PROMPT = """你是投行部的信息分析助手。下面是我给的检索指令，以及从网上抓回来的资料。
请基于资料回答，不要编造；资料里没有的，写「未提及」。

输出格式（Markdown）：
## 结论摘要
（3 句以内）
## 关键事实
- 主体：
- 事件：
- 金额 / 比例：
- 时间：
- 地点 / 来源：
## 可切入的银行业务
- （2-3 条，写清用什么产品、找谁）
## 待核实
- （1-3 条）"""


def _agent_cfg():
    return {"provider": sget("ag_provider", "deepseek"),
            "base_url": sget("ag_base_url", PROVIDERS["deepseek"][1]),
            "model": sget("ag_model", "deepseek-chat"),
            "api_key": sget("ag_key", ""),
            "temperature": sget("ag_temp", "0.2"),
            "max_tokens": sget("ag_max", "1500")}


def _agent_chat(cfg, messages, timeout=120):
    base = (cfg.get("base_url") or "").rstrip("/")
    if not base:
        raise RuntimeError("未配置接口地址")
    if not cfg.get("api_key"):
        raise RuntimeError("未配置 API Key")
    url = base if base.endswith("/chat/completions") else base + "/chat/completions"
    payload = {"model": cfg.get("model") or "gpt-4o-mini", "messages": messages,
               "temperature": float(cfg.get("temperature") or 0.2),
               "max_tokens": int(cfg.get("max_tokens") or 1500), "stream": False}
    t0 = time.time()
    d = _post_json(url, payload, timeout=timeout,
                   headers={"Authorization": "Bearer " + cfg["api_key"]})
    ms = int((time.time() - t0) * 1000)
    if "choices" not in d:
        raise RuntimeError("接口返回异常：%s" % json.dumps(d, ensure_ascii=False)[:300])
    text = d["choices"][0]["message"]["content"]
    usage = d.get("usage") or {}
    return text, int(usage.get("total_tokens") or 0), ms


def _strip_html(html):
    html = re.sub(r"(?is)<(script|style|noscript|svg|iframe)[^>]*>.*?</\1>", " ", html)
    html = re.sub(r"(?is)<br\s*/?>", "\n", html)
    html = re.sub(r"(?is)</(p|div|h[1-6]|li|tr|section|article)>", "\n", html)
    html = re.sub(r"(?s)<[^>]+>", " ", html)
    html = (html.replace("&nbsp;", " ").replace("&amp;", "&").replace("&lt;", "<")
            .replace("&gt;", ">").replace("&quot;", '"').replace("&#39;", "'"))
    html = re.sub(r"[ \t\u3000]+", " ", html)
    html = re.sub(r"\n{3,}", "\n\n", html)
    return html.strip()


def _page_text(url, limit=9000):
    html = _http_get_text(url)
    text = _strip_html(html)
    title = ""
    m = re.search(r"(?is)<title[^>]*>(.*?)</title>", html)
    if m:
        title = _strip_html(m.group(1)).strip()
    return title, text[:limit]


# ---------------------------------------------------------------------------
# 网页检索（无需用户填网址）
# ---------------------------------------------------------------------------
LOTTERY_TYPES = [("双色球", "ssq"), ("七乐彩", "qlc"), ("快乐8", "kl8")]
LOTTERY_API = "https://www.cwl.gov.cn/cwl_admin/front/cwlkj/search/kjxx/findDrawNotice"
LOTTERY_UNSUPPORTED = ("体彩", "大乐透", "排列三", "排列五", "七星彩", "3d", "足彩", "竞彩")


def _lottery_blocks(prompt):
    """彩票开奖这类结构化数据，直接查官方接口，不靠搜索+网页（网页多是 JS 渲染拿不到号码）。"""
    p = (prompt or "")
    low = p.lower()

    if any(k in low for k in LOTTERY_UNSUPPORTED):
        return None          # 体彩/3D 等暂不支持，走普通搜索
    if not any(k in low for k in ("双色球", "福彩", "彩票", "开奖", "七乐彩", "快乐8", "中奖")):
        return None

    lot = ""
    for kw, code in LOTTERY_TYPES:
        if kw in low:
            lot = code
            break
    if not lot:
        lot = "ssq"          # 只说「开奖/彩票/福彩」默认双色球

    m = re.search(r"(?:最近|近|连续)?\s*(\d{1,2})\s*期", p)
    n = int(m.group(1)) if m else 1
    n = max(1, min(n, 20))
    url = "%s?name=%s&issueCount=%d" % (LOTTERY_API, lot, n)

    try:
        req = urllib.request.Request(url, headers={
            "User-Agent": UA,
            "Referer": "https://www.cwl.gov.cn/",
            "Accept": "application/json, text/plain, */*",
        })
        raw = urllib.request.urlopen(req, timeout=15).read().decode("utf-8", "replace")
        data = json.loads(raw)
    except Exception as e:  # noqa: BLE001
        return {"text": "查询福彩官方开奖接口失败：%s。可稍后重试，或到 www.cwl.gov.cn 查看。" % e,
                "src_note": "中国福彩网官方接口"}

    items = (data or {}).get("result") or []
    if not items:
        return {"text": "福彩官方接口没有返回开奖数据（可能该彩种还没开奖或期号不对）。",
                "src_note": "中国福彩网官方接口"}

    lines = ["## 开奖结果（中国福彩网官方数据）"]
    for it in items:
        red = " ".join(str(it.get("red") or "").split(","))
        blue = str(it.get("blue") or "")
        lines.append("")
        lines.append("**第 %s 期 · %s**（%s）" % (it.get("code"), it.get("name", ""), it.get("date")))
        if red:
            lines.append("红球：%s" % red)
        if blue:
            lines.append("蓝球：%s" % blue)
    lines.append("")
    lines.append("（数据来源：中国福彩网官方开奖接口；投注请以官方公告为准）")
    return {"text": "\n".join(lines), "src_note": "中国福彩网官方开奖数据"}


def _search_query(prompt):
    """从指令里提炼检索词：取第一行，去掉常见口语前缀与标点。"""
    line = (prompt or "").strip().splitlines()[0] if (prompt or "").strip() else ""
    line = re.sub(r"^(帮我|请|麻烦|能不能|可以)?\s*(抓取|搜索|查找|查一下|找一下|找找|爬取|获取|看看)\s*", "", line)
    line = re.sub(r"^(这期|本期|最新一期|最新|最近|今天|昨天|本周|本月|现在)的?\s*", "", line)
    line = re.sub(r"(的开奖信息|开奖信息|开奖结果|最新开奖结果|的开奖号码|开奖号码|的信息|信息|结果|号码)\s*$", "", line)
    line = re.sub(r"^的\s*", "", line)
    line = re.sub(r"\s+", " ", line)
    line = re.sub(r"[，。；！？,.;!?]+$", "", line)
    line = line.strip("：: ")
    if any(k in line for k in ("双色球", "彩票", "福彩", "开奖")):
        line += " 开奖号码"
    return line[:60] or "投行 并购 融资 最新"


def _unwrap_bing(u):
    if "bing.com/ck/a" not in u:
        return u
    m = re.search(r"[?&]u=a1([^&]+)", u)
    if not m:
        return u
    try:
        import base64
        raw = m.group(1)
        raw += "=" * (-len(raw) % 4)
        return base64.urlsafe_b64decode(raw).decode("utf-8", "ignore")
    except Exception:  # noqa: BLE001
        return u


def _bing(q, limit=8):
    url = "https://cn.bing.com/search?q=%s&count=%d&setlang=zh-CN" % (urllib.parse.quote(q), limit + 4)
    html = _http_get_text(url, timeout=20)
    out = []
    for block in re.findall(r'<li class="b_algo".*?</li>', html, re.S)[:limit + 4]:
        m = re.search(r'<h2[^>]*>\s*<a[^>]*href="([^"]+)"[^>]*>(.*?)</a>', block, re.S)
        if not m:
            continue
        href = _unwrap_bing(m.group(1))
        title = _strip_html(m.group(2)).strip()
        snippet = _strip_html(re.sub(r"(?is)<h2.*?</h2>", " ", block)).strip()
        if not href.startswith("http") or "%EF%BF%BD" in href or " " in href:
            continue
        out.append({"title": title, "url": href, "snippet": snippet[:300]})
    return out


def _sogou(q, limit=8):
    url = "https://www.sogou.com/web?query=%s" % urllib.parse.quote(q)
    html = _http_get_text(url, timeout=20)
    out = []
    for m in re.finditer(r'<h3[^>]*class="vr-title"[^>]*>\s*<a[^>]*href="([^"]+)"[^>]*>(.*?)</a>', html, re.S):
        href = m.group(1)
        if href.startswith("/"):
            href = "https://www.sogou.com" + href
        title = _strip_html(m.group(2)).strip()
        if href.startswith("http") and title:
            out.append({"title": title, "url": href, "snippet": ""})
        if len(out) >= limit:
            break
    return out


def _web_search(q, limit=8):
    for fn in (_bing, _sogou):
        try:
            hits = fn(q, limit)
            if hits:
                return hits
        except Exception:  # noqa: BLE001
            continue
    return []


def _pick_sources(hits, want=3, min_len=200):
    """抓取候选页面正文，保留内容够长的；同一域名最多 2 篇。"""
    picked, seen_domain = [], {}
    for h in hits:
        if len(picked) >= want:
            break
        dom = urllib.parse.urlparse(h["url"]).netloc
        if seen_domain.get(dom, 0) >= 2:
            continue
        if re.search(r"\.(pdf|docx?|xlsx?|zip|jpg|png|mp4)$", h["url"], re.I):
            continue
        try:
            title, text = _page_text(h["url"])
        except Exception:  # noqa: BLE001
            continue
        if len(text) < min_len:
            continue
        seen_domain[dom] = seen_domain.get(dom, 0) + 1
        picked.append({"title": title or h["title"], "url": h["url"], "text": text})
    return picked


def _run_agent_bg(run_id, cfg, prompt, kw=""):
    try:
        lot = _lottery_blocks(prompt)
        if lot is not None:
            x("UPDATE agent_runs SET status='done', result=?, sources=?, tokens=0, cost_ms=0 WHERE id=?",
              (lot["text"], lot["src_note"], run_id))
            return
        q = (kw or "").strip() or _search_query(prompt)
        hits = _web_search(q)
        if not hits:
            x("UPDATE agent_runs SET status='error', err=? WHERE id=?",
              ("搜索没有返回结果（可能被限流），换个说法再试一次。检索词：%s" % q, run_id))
            return
        picked = _pick_sources(hits, want=3)
        if picked:
            blocks = ["### 来源：%s\n标题：%s\n链接：%s\n\n%s" % (i + 1, p["title"], p["url"], p["text"])
                      for i, p in enumerate(picked)]
            src_note = "；".join(p["url"] for p in picked)
        else:
            # 正文都没抓到，退化为用搜索结果摘要
            blocks = ["### 搜索结果摘要\n" + "\n".join(
                "- %s：%s（%s）" % (h["title"], h["snippet"], h["url"]) for h in hits[:8])]
            src_note = "摘要：" + "；".join(h["url"] for h in hits[:5])
        user = ("检索指令：%s\n检索词：%s\n\n以下是抓回的网页资料，请据此分析。\n\n%s"
                % (prompt.strip()[:500], q, "\n\n---\n\n".join(blocks)))
        text, tokens, ms = _agent_chat(cfg, [{"role": "system", "content": DEFAULT_PROMPT},
                                             {"role": "user", "content": user[:26000]}])
        x("UPDATE agent_runs SET status='done', result=?, tokens=?, cost_ms=?, sources=? WHERE id=?",
          (text, tokens, ms, src_note[:500], run_id))
    except Exception as e:  # noqa: BLE001
        x("UPDATE agent_runs SET status='error', err=? WHERE id=?", ("%s: %s" % (type(e).__name__, e), run_id))


@bp.route("/agent")
def agent_home():
    tab = request.args.get("tab", "model")
    if tab not in [k for k, _ in AGENT_TABS]:
        tab = "model"
    cfg = _agent_cfg()
    flash = ""
    msg = request.args.get("msg") or ""
    if msg:
        kind, _, text = msg.partition(":")
        flash = _flash("ok" if kind == "ok" else "err", text)

    tabs = "<div class='cfg-tabs'>" + "".join(
        "<a class='cfg-tab %s' href='/agent?tab=%s'>%s</a>" % ("on" if tab == k else "", k, n)
        for k, n in AGENT_TABS) + "</div>"

    if tab == "fetch":
        body = _agent_fetch_html(cfg, flash)
    elif tab == "gd":
        body = _gd_html(flash)
    else:
        body = _agent_model_html(cfg, flash)
    return _PAGE("<div class='card'><h2>智能体配置 <span class='mini'>模型 · 抓取 · 适配器</span></h2>"
                 + CSS + tabs + "</div>" + body, "agent")


def _agent_model_html(cfg, flash):
    presets = {k: {"name": v[0], "url": v[1], "models": v[2]} for k, v in PROVIDERS.items()}
    opts = "".join("<option value='%s' %s>%s</option>"
                   % (k, "selected" if cfg["provider"] == k else "", esc(v[0]))
                   for k, v in PROVIDERS.items())
    urls = sorted({v[1] for v in PROVIDERS.values() if v[1]})
    models = []
    for v in PROVIDERS.values():
        models += v[2]
    runs = q("SELECT * FROM agent_runs ORDER BY id DESC LIMIT 8")
    total_tokens = one("SELECT COALESCE(SUM(tokens),0) s FROM agent_runs")["s"]
    usage_rows = "".join(
        "<tr><td class='cfg-num'>%d</td><td>%s</td><td class='mini'>%s</td><td class='cfg-num'>%s</td>"
        "<td class='cfg-num'>%d</td><td class='cfg-num'>%.1fs</td></tr>"
        % (r["id"], esc(r["model"] or ""), _status_pill(r["status"]), esc(r["created"] or ""),
           r["tokens"] or 0, (r["cost_ms"] or 0) / 1000.0)
        for r in runs) or "<tr><td colspan='6' class='mini'>还没有调用记录</td></tr>"
    key_hint = ("已配置：<span class='cfg-mask'>%s</span>（留空表示不修改）" % ("•" * 12 + cfg["api_key"][-4:])
                if cfg["api_key"] else "尚未配置 API Key")
    body = """
    <div class="card"><h2>模型配置</h2>
      %s
      <form method="post" action="/agent/save">
        <div class="row">
          <div><label>服务商</label><select name="provider" id="provSel" onchange="applyPreset()">%s</select></div>
          <div style="flex:2"><label>接口地址（Base URL，可手填或下拉）</label>
            <input name="base_url" id="baseUrl" list="urlList" value="%s" placeholder="https://api.deepseek.com/v1">
            <datalist id="urlList">%s</datalist></div>
          <div style="flex:1.4"><label>模型（可手填或下拉）</label>
            <input name="model" id="modelInput" list="modelList" value="%s" placeholder="deepseek-chat">
            <datalist id="modelList">%s</datalist></div>
        </div>
        <div class="row">
          <div style="flex:2"><label>API Key</label>
            <input name="api_key" type="password" autocomplete="new-password" placeholder="sk-... 粘贴后保存，页面不会回显">
            <div class="cfg-hint" style="margin-top:4px">%s</div></div>
          <div><label>温度</label><input name="temperature" value="%s" placeholder="0.2"></div>
          <div><label>最大 Token</label><input name="max_tokens" value="%s" placeholder="1500"></div>
        </div>
        <div style="margin-top:14px">
          <button class="btn">保存配置</button>
          <a class="btn btn3" href="/agent/test" style="margin-left:8px">测试连接</a>
          <span class="cfg-hint" style="margin-left:10px">测试会真实调用一次模型（约消耗几十个 token）。</span>
        </div>
      </form>
      <div class="cfg-warn">Key 以明文存于本机 SQLite（部署云端时随服务器磁盘走）。如需更严格，建议改为环境变量注入，
        或只给受限的专用 Key（限制额度、限制模型）。</div>
    </div>

    <div class="card"><h2>用量统计 <span class="mini">累计 %s tokens</span></h2>
      <div class="tbl-wrap"><table class="cfg">
        <tr><th>ID</th><th>模型</th><th>状态</th><th>时间</th><th>Tokens</th><th>耗时</th></tr>
        %s</table></div>
    </div>
    <script>
    var PRESETS = %s;
    function applyPreset(){
      var k=document.getElementById('provSel').value, p=PRESETS[k]; if(!p) return;
      if(p.url) document.getElementById('baseUrl').value=p.url;
      var dl=document.getElementById('modelList'); dl.innerHTML='';
      p.models.forEach(function(m){var o=document.createElement('option');o.value=m;dl.appendChild(o)});
      if(p.models.length) document.getElementById('modelInput').value=p.models[0];
    }
    </script>
    """ % (flash, opts, esc(cfg["base_url"]),
           "".join("<option value='%s'>" % esc(u) for u in urls),
           esc(cfg["model"]),
           "".join("<option value='%s'>" % esc(m) for m in sorted(set(models))),
           key_hint, esc(cfg["temperature"]), esc(cfg["max_tokens"]),
           format(total_tokens, ","), usage_rows,
           json.dumps(presets, ensure_ascii=False))
    return body


@bp.route("/agent/save", methods=["POST"])
def agent_save():
    sset("ag_provider", request.form.get("provider") or "deepseek")
    sset("ag_base_url", (request.form.get("base_url") or "").strip())
    sset("ag_model", (request.form.get("model") or "").strip())
    sset("ag_temp", request.form.get("temperature") or "0.2")
    sset("ag_max", request.form.get("max_tokens") or "1500")
    key = (request.form.get("api_key") or "").strip()
    if key:
        sset("ag_key", key)
    return _back("ok:模型配置已保存", "model")


@bp.route("/agent/test")
def agent_test():
    cfg = _agent_cfg()
    try:
        text, tokens, ms = _agent_chat(cfg, [{"role": "user", "content": "只回复两个字：正常"}], timeout=60)
    except Exception as e:  # noqa: BLE001
        return _back("err:连接失败 %s" % e, "model")
    x("INSERT INTO agent_runs (provider,model,prompt,sources,status,tokens,cost_ms,result,created) "
      "VALUES (?,?,?,?,?,?,?,?,?)",
      (cfg["provider"], cfg["model"], "连接测试", "", "done", tokens, ms, text[:200], now_str()))
    return _back("ok:连接正常，返回「%s」，耗时 %.1fs，tokens %d" % (text.strip()[:20], ms / 1000.0, tokens), "model")


def _agent_fetch_html(cfg, flash):
    runs = q("SELECT * FROM agent_runs WHERE prompt<>'连接测试' ORDER BY id DESC LIMIT 10")
    rows = "".join(
        "<tr><td class='cfg-num'>%d</td><td>%s</td><td class='mini'>%s</td><td class='cfg-num'>%d</td>"
        "<td><a href='/agent/run/%d'>查看</a></td></tr>"
        % (r["id"], _status_pill(r["status"]), esc((r["sources"] or "")[:56]), r["tokens"] or 0, r["id"])
        for r in runs) or "<tr><td colspan='5' class='mini'>还没有抓取任务</td></tr>"
    key_ok = bool(cfg["api_key"])
    body = """
    <div class="card"><h2>智能抓取 <span class="mini">说需求，自动搜网页并提炼</span></h2>
      %s
      <div class="cfg-mini">当前模型：<b>%s</b>　接口：%s　%s</div>
      <form id="runForm" style="margin-top:10px" data-ajax="1">
        <label class="cfg-mini">指令（用一句话说清要什么，例：2026 年 10 月广州半导体行业的融资事件）</label>
        <textarea name="prompt" rows="3" placeholder="不用填网址，系统会自动检索公开网页"></textarea>
        <div class="cfg-bar" style="margin-top:8px">
          <span class="cfg-mini">检索词（可选，留空自动从指令提取）</span>
          <input name="kw" placeholder="自动提取" style="width:220px">
          <button class="btn" type="submit" style="margin-left:auto">▶ 开始检索分析</button>
        </div>
      </form>
      <div class="cfg-prog"><i id="agBar"></i></div>
      <div class="cfg-mini" id="agStatus">%s</div>
      <div class="card" id="agBox" style="display:none;margin-top:10px">
        <h2>分析结果 <span class="mini" id="agMeta"></span>
          <span style="margin-left:auto;display:flex;gap:8px">
            <a class="btn btn3 btn-sm" href="javascript:saveDeal()">存为商机草稿</a>
            <a class="btn btn2 btn-sm" href="javascript:saveLocal()">写入本机文件</a></span></h2>
        <div class="cfg-pre" id="agResult"></div>
        <div class="cfg-hint" id="srcList" style="margin-top:8px"></div>
      </div>
    </div>

    <div class="card"><h2>最近任务</h2>
      <div class="tbl-wrap"><table class="cfg">
        <tr><th>ID</th><th>状态</th><th>检索来源</th><th>Tokens</th><th>结果</th></tr>
        %s</table></div>
    </div>

    <div class="modal" id="hostModal" onclick="if(event.target.id==='hostModal')closeHost()">
      <div class="modal-box">
        <div class="modal-h"><b>需要本机授权</b><a href="javascript:closeHost()" class="modal-x">×</a></div>
        <div class="cfg-hint" style="margin:8px 0 10px">
          要把结果写到你电脑上，需要先在本机运行一次配对。以后不会再问。</div>
        <div class="cfg-file" style="font-size:12.5px">
          配对码 <b class="cfg-mask" id="pairCode" style="font-size:22px;letter-spacing:5px">------</b>
        </div>
        <div class="cfg-hint" style="margin-top:10px">在本机执行：</div>
        <div class="cfg-pre" id="pairCmd" style="font-size:12px;padding:10px 12px"></div>
        <div class="modal-f">
          <a class="btn btn3" href="javascript:closeHost()">稍后</a>
          <a class="btn" href="javascript:checkHost()">我已运行，检查授权</a>
        </div>
      </div>
    </div>

    <script>
    var CUR=null, HOST_TASK=null;
    document.getElementById('runForm').addEventListener('submit',function(ev){
      ev.preventDefault();
      var fd=new FormData(this);
      document.getElementById('agStatus').textContent='正在检索网页并调用模型，约 20-60 秒…';
      document.getElementById('agBar').style.width='30%%';
      document.getElementById('agBox').style.display='none';
      fetch('/agent/run',{method:'POST',body:fd}).then(function(r){return r.json()}).then(function(d){
        if(d.error){document.getElementById('agStatus').textContent='失败：'+d.error;
          document.getElementById('agBar').style.width='0';return;}
        poll(d.id);
      }).catch(function(e){document.getElementById('agStatus').textContent='提交失败：'+e});
    });
    function poll(id){
      fetch('/agent/run/'+id+'/json').then(function(r){return r.json()}).then(function(d){
        if(d.status==='running'||d.status==='queued'){
          document.getElementById('agBar').style.width='70%%';
          setTimeout(function(){poll(id)},1500);return;
        }
        document.getElementById('agBar').style.width='100%%';
        if(d.status==='error'){
          document.getElementById('agStatus').textContent='失败：'+d.err;return;
        }
        CUR=d;
        document.getElementById('agStatus').textContent='完成，耗时 '+(d.cost_ms/1000).toFixed(1)+' 秒，消耗 '+d.tokens+' tokens';
        document.getElementById('agMeta').textContent='tokens '+d.tokens;
        document.getElementById('agResult').textContent=d.result;
        document.getElementById('srcList').textContent='来源：'+(d.sources||'—');
        document.getElementById('agBox').style.display='';
      });
    }
    function saveDeal(){
      if(!CUR)return;
      fetch('/agent/save_deal',{method:'POST',headers:{'Content-Type':'application/json'},
        body:JSON.stringify({run_id:CUR.id})}).then(function(r){return r.json()}).then(function(d){
          if(d.error){alert(d.error);return;}
          alert('已存为商机草稿 #'+d.deal_id);
        });
    }
    function saveLocal(){
      if(!CUR)return;
      fetch('/host/quick',{method:'POST',headers:{'Content-Type':'application/json'},
        body:JSON.stringify({content:CUR.result})}).then(function(r){return r.json()}).then(function(d){
          if(d.error){alert(d.error);return;}
          if(d.paired){HOST_TASK=d.task_id;publishWait();return;}
          showHost(d.code);
        });
    }
    function showHost(code){
      document.getElementById('pairCode').textContent=code;
      var cmd='python host_agent.py --server '+location.origin+' --code '+code;
      document.getElementById('pairCmd').textContent=cmd;
      document.getElementById('hostModal').classList.add('on');
    }
    function closeHost(){document.getElementById('hostModal').classList.remove('on')}
    function checkHost(){
      fetch('/host/status').then(function(r){return r.json()}).then(function(d){
        if(d.paired){closeHost();saveLocal();}else{alert('还没检测到授权，确认本机 agent 已启动。')}
      });
    }
    function publishWait(){
      document.getElementById('agStatus').textContent='已提交本机写入请求，请在弹窗里确认（y）…';
      var n=0;
      var t=setInterval(function(){
        n++;
        fetch('/host/task/'+HOST_TASK+'/json').then(function(r){return r.json()}).then(function(d){
          if(d.status==='done'){
            clearInterval(t);
            document.getElementById('agStatus').textContent='已写入本机：'+(d.result||'ok');
          }else if(d.status==='error'){
            clearInterval(t);
            document.getElementById('agStatus').textContent='写入失败：'+(d.err||'本机拒绝');
          }else if(n>40){clearInterval(t);}
        });
      },1500);
    }
    </script>
    """ % (flash, esc(cfg["model"]), esc(cfg["base_url"]),
           "Key 已配置" if key_ok else "<span style='color:#a12722'>未配置 Key，无法运行</span>",
           "尚未运行" if key_ok else "请先到「模型配置」填写 API Key",
           rows)
    return body


@bp.route("/agent/run", methods=["POST"])
def agent_run():
    cfg = _agent_cfg()
    if not cfg["api_key"]:
        return jsonify({"error": "未配置 API Key"})
    prompt = (request.form.get("prompt") or "").strip()
    if len(prompt) < 4:
        return jsonify({"error": "请把需求写清楚一点（至少几个字）"})
    kw = (request.form.get("kw") or "").strip()
    rid = x("INSERT INTO agent_runs (provider,model,prompt,sources,status,created) VALUES (?,?,?,?,?,?)",
            (cfg["provider"], cfg["model"], prompt[:300], "检索中…", "running", now_str()))
    threading.Thread(target=_run_agent_bg, args=(rid, cfg, prompt, kw), daemon=True).start()
    return jsonify({"id": rid})


@bp.route("/agent/run/<int:rid>/json")
def agent_run_json(rid):
    r = one("SELECT * FROM agent_runs WHERE id=?", (rid,))
    if not r:
        return jsonify({"error": "任务不存在"})
    return jsonify({"id": r["id"], "status": r["status"], "result": r["result"] or "",
                    "err": r["err"] or "", "tokens": r["tokens"] or 0,
                    "cost_ms": r["cost_ms"] or 0, "sources": r["sources"] or ""})


@bp.route("/agent/run/<int:rid>")
def agent_run_view(rid):
    r = one("SELECT * FROM agent_runs WHERE id=?", (rid,))
    if not r:
        abort(404)
    body = ("<div class='card'><h2>抓取结果 #%d <span class='mini'>%s</span></h2>"
            "<div class='cfg-mini'>模型 %s ｜ tokens %d ｜ 耗时 %.1fs</div>"
            "<div class='cfg-mini' style='margin-top:4px'>来源：%s</div>"
            "<div class='cfg-pre' style='margin-top:10px'>%s</div>"
            "<div style='margin-top:12px'><a class='btn btn3' href='/agent?tab=fetch'>返回</a></div></div>"
            % (r["id"], esc(r["created"]), esc(r["model"]), r["tokens"] or 0,
               (r["cost_ms"] or 0) / 1000.0, esc(r["sources"]),
               esc(r["result"] or r["err"] or "（无结果）")))
    return _PAGE("<div class='card'><h2>智能体</h2>" + CSS + "</div>" + body, "agent")


@bp.route("/agent/save_deal", methods=["POST"])
def agent_save_deal():
    data = request.get_json(silent=True) or {}
    r = one("SELECT * FROM agent_runs WHERE id=?", (data.get("run_id"),))
    if not r or not r["result"]:
        return jsonify({"error": "结果不存在"})
    title = "【智能体抓取】%s" % (r["prompt"] or "")[:30]
    did = x("INSERT INTO deals (name,client,sector,stage,status,amount,owner,note,created,updated,deal_date) "
            "VALUES (?,?,?,?,?,?,?,?,?,?,?)",
            (title, "", "", "初步接洽", "进行中", 0, "",
             r["result"][:4000], now_str(), now_str(),
             datetime.date.today().strftime("%Y-%m-%d")))
    return jsonify({"deal_id": did})


# ---------------------------------------------------------------------------
# 宿主机授权（按需触发，不再单独占一个 tab）
# ---------------------------------------------------------------------------
def _new_pair_code():
    code = "%06d" % (int(time.time() * 1000) % 1000000)
    token = uuid.uuid4().hex
    exp = (datetime.datetime.now() + datetime.timedelta(minutes=30)).strftime("%Y-%m-%d %H:%M")
    x("UPDATE host_pair SET used=1")
    x("INSERT INTO host_pair (code,token,created,expires,used) VALUES (?,?,?,?,0)",
      (code, token, now_str(), exp))
    return code


def _paired():
    return bool(one("SELECT id FROM host_pair WHERE used=1 ORDER BY id DESC LIMIT 1"))


@bp.route("/host/status")
def host_status():
    return jsonify({"paired": _paired(),
                    "pending": one("SELECT COUNT(*) c FROM host_tasks WHERE status='pending'")["c"]})


@bp.route("/host/quick", methods=["POST"])
def host_quick():
    data = request.get_json(silent=True) or {}
    content = (data.get("content") or "").strip()
    if not content:
        return jsonify({"error": "内容为空"})
    if not _paired():
        return jsonify({"paired": False, "code": _new_pair_code()})
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    path = "~/WorkBuddy/智能体输出/摘要_%s.md" % stamp
    tid = x("INSERT INTO host_tasks (op,path,content,status,created) VALUES (?,?,?,?,?)",
            ("write", path, content, "pending", now_str()))
    return jsonify({"paired": True, "task_id": tid, "path": path})


@bp.route("/host/task/<int:tid>/json")
def host_task_json(tid):
    t = one("SELECT * FROM host_tasks WHERE id=?", (tid,))
    if not t:
        return jsonify({"error": "任务不存在"})
    return jsonify({"id": t["id"], "status": t["status"],
                    "result": (t["result"] or t["err"] or "")[:300], "path": t["path"]})


@bp.route("/api/host/pair", methods=["POST"])
def api_host_pair():
    data = request.get_json(silent=True) or {}
    code = (data.get("code") or "").strip()
    row = one("SELECT * FROM host_pair WHERE used=0 ORDER BY id DESC LIMIT 1")
    if not row or row["code"] != code:
        return jsonify({"error": "配对码不正确或已失效"})
    if row["expires"] < now_str()[:16]:
        return jsonify({"error": "配对码已过期"})
    x("UPDATE host_pair SET used=1 WHERE id=?", (row["id"],))
    return jsonify({"token": row["token"]})


def _host_token_ok():
    tok = request.args.get("token", "") or request.headers.get("X-Host-Token", "")
    if not tok and request.method == "POST":
        data = request.get_json(silent=True) or {}
        tok = str(data.get("token") or "")
    return bool(one("SELECT id FROM host_pair WHERE token=? AND used=1", (tok,))), tok


@bp.route("/api/host/tasks")
def api_host_tasks():
    ok, _ = _host_token_ok()
    if not ok:
        return jsonify({"error": "unauthorized"}), 401
    rows = q("SELECT id,op,path,content,created FROM host_tasks WHERE status='pending' ORDER BY id LIMIT 5")
    return jsonify({"tasks": [dict(r) for r in rows],
                    "allow_dirs": [d for d in (sget("host_allow") or "").split(";") if d]})


@bp.route("/api/host/report", methods=["POST"])
def api_host_report():
    data = request.get_json(silent=True) or {}
    ok, _ = _host_token_ok()
    if not ok:
        return jsonify({"error": "unauthorized"}), 401
    tid = data.get("id")
    if not one("SELECT id FROM host_tasks WHERE id=?", (tid,)):
        return jsonify({"error": "task not found"})
    if data.get("ok"):
        x("UPDATE host_tasks SET status='done', result=?, finished=? WHERE id=?",
          (str(data.get("result") or "")[:20000], now_str(), tid))
    else:
        x("UPDATE host_tasks SET status='error', err=?, finished=? WHERE id=?",
          (str(data.get("error") or "")[:1000], now_str(), tid))
    return jsonify({"ok": True})


@bp.route("/host/agent.py")
def host_agent_download():
    path = os.path.join(os.path.dirname(_DB), "host_agent.py")
    if not os.path.isfile(path):
        abort(404)
    return send_file(path, as_attachment=True, download_name="host_agent.py")

# ---------------------------------------------------------------------------
# 七、专项抓取：指令 + 模板驱动
# ---------------------------------------------------------------------------
DEFAULT_GD_PROMPT = (
    "你是一个资深爬虫专家。现在有一个网站：广东省投资项目在线审批监管平台（https://tzxm.gd.gov.cn/index.html）。"
    "请到「公示信息 → 办理结果公示 → 项目备案 → 备案公开」页面，找出备案通过日期在 20261027 ~ 20261029 之间的记录，"
    "按我上传的模板整理成台账。最后一列只对总投资 5 亿元以上的项目生成摘要："
    "建设单位拟建设项目名称，项目规模及内容，项目地位于项目所在地（精确到区），总投资XX亿元。"
)

# 模板列 → 数据字段 的模糊匹配规则（顺序敏感，先匹配到的先用）
COL_RULES = [
    ("code", ["备案项目编号", "项目编号", "编号", "项目代码"]),
    ("name", ["项目名称", "名称"]),
    ("area", ["项目所在地", "所在地", "地市", "区域", "地区"]),
    ("money", ["项目总投资", "总投资", "投资额", "金额"]),
    ("scale", ["项目规模及内容", "规模及内容", "建设规模", "规模", "内容"]),
    ("dept", ["建设单位", "项目单位", "企业名称", "申报单位", "单位"]),
    ("gov", ["备案机关", "核准机关", "审批机关", "备案机构", "机关"]),
    ("apply", ["备案申报日期", "申报日期", "申报时间", "申报"]),
    ("finish", ["复核通过日期", "通过日期", "备案日期", "通过时间", "日期"]),
    ("years", ["项目起止年限", "起止年限", "建设年限", "年限"]),
    ("state", ["项目当前状态", "当前状态", "状态"]),
    ("summary", ["摘要", "小结", "5亿元以上", "亿元以上"]),
]

FIELD_OF_HEADER = {}


# 字段 → 标准列名（没上传模板时，按指令里提到的列输出）
FIELD_HEADER = {
    "code": "项目编号", "name": "项目名称", "area": "项目所在地", "money": "项目总投资",
    "scale": "项目规模及内容", "dept": "建设单位", "gov": "备案机关",
    "apply": "申报日期", "finish": "通过日期", "years": "项目起止年限",
    "state": "项目当前状态", "summary": "摘要",
}


def _cols_from_text(text):
    """从指令里识别用户要输出哪些列；没提就返回空（由调用方回退标准列）。"""
    t = text or ""
    found = []
    for field, keys in COL_RULES:
        pos = -1
        for k in keys:
            i = t.find(k)
            if i >= 0 and (pos < 0 or i < pos):
                pos = i
        if pos >= 0:
            found.append((pos, field))
    found.sort()
    cols, seen = [], set()
    for _, f in found:
        if f not in seen:
            seen.add(f)
            cols.append(FIELD_HEADER.get(f, ""))
    return cols


URL_RE = re.compile(r"https?://[^\s，。；、）)\"'<>【】\]]+")


def _urls_in(text):
    """从指令里抠出网址（去掉结尾标点）。"""
    out = []
    for u in URL_RE.findall(text or ""):
        u = u.rstrip(".,;:!?，。；、）)]】\"'")
        if u and u not in out:
            out.append(u)
    return out


def _save_adapter(spec, name=""):
    """保存或更新适配器（平台适配页与自动识别共用）。"""
    spec = ad.normalize_spec(spec)
    nm = (name or spec.get("name") or spec.get("host") or "未命名").strip()[:60]
    payload = json.dumps(spec, ensure_ascii=False)
    old = one("SELECT id FROM adapters WHERE host=? AND list_url=?",
              (spec.get("host") or "", spec.get("list_url") or ""))
    if old:
        x("UPDATE adapters SET name=?, spec=?, updated=? WHERE id=?",
          (nm, payload, now_str(), old["id"]))
        return old["id"], nm
    aid = x("INSERT INTO adapters (name,host,list_url,spec,enabled,created,updated) "
            "VALUES (?,?,?,?,1,?,?)",
            (nm, spec.get("host") or "", spec.get("list_url") or "", payload,
             now_str(), now_str()))
    return aid, nm


def _gd_setting(k, d=""):
    return sget("gd_" + k, d)


def _gd_state():
    last_end = _gd_setting("last_end")
    hol = [x.strip() for x in _gd_setting("holidays").replace("，", ",").split(",") if x.strip()]
    wrk = [x.strip() for x in _gd_setting("workdays").replace("，", ",").split(",") if x.strip()]
    return last_end, hol, wrk


def _gd_default():
    y = datetime.date.today() - datetime.timedelta(days=1)
    return gd._s(y - datetime.timedelta(days=2)), gd._s(y)


def _gd_window():
    last_end, hol, wrk = _gd_state()
    yesterday = datetime.date.today() - datetime.timedelta(days=1)
    if last_end:
        start = gd._d(last_end) + datetime.timedelta(days=1)
    else:
        start = gd.shift_workday(yesterday, -2, hol, wrk)
    if start > yesterday:
        return None, None
    if (yesterday - start).days > 14:
        start = yesterday - datetime.timedelta(days=14)
    return gd._s(start), gd._s(yesterday)


# ---------------------------------------------------------------------------
# 指令解析：模型优先，失败回落正则
# ---------------------------------------------------------------------------
def _regex_plan(text):
    txt = text or ""
    dates = re.findall(r"(20\d{2})[-/.年]?(\d{1,2})[-/.月]?(\d{1,2})", txt)
    norm = []
    for y, m, d in dates:
        try:
            norm.append(datetime.date(int(y), int(m), int(d)).strftime("%Y%m%d"))
        except ValueError:
            continue
    start = min(norm) if norm else ""
    end = max(norm) if norm else ""
    if len(norm) >= 2 and start == end:
        start, end = min(norm), max(norm)
    big = 5.0
    m = re.search(r"(\d+(?:\.\d+)?)\s*亿元以上", txt)
    if m:
        big = float(m.group(1))
    adapter = "gd_tzxm" if any(k in txt for k in ("tzxm", "广东", "备案", "投资项目")) else ""
    return {"adapter": adapter, "start": start, "end": end,
            "min_amount_wan": int(big * 10000), "only_big": ("只" in txt and "以上" in txt),
            "how": "规则解析"}


def _llm_plan(cfg, text):
    prompt = (
        "把用户的爬取需求解析成 JSON，只输出 JSON，不要解释。字段：\n"
        '{"adapter":"gd_tzxm 或 generic","start":"YYYYMMDD","end":"YYYYMMDD",'
        '"min_amount_wan":整数,"only_big":true/false,"reason":"一句话"}\n'
        "已知适配器：gd_tzxm = 广东省投资项目在线审批监管平台(tzxm.gd.gov.cn) 备案公开；"
        "其它站点一律 generic。日期没写就留空字符串。金额门槛默认 50000（万元）。"
        "注意：若用户只是规定『摘要列只对 X 亿元以上项目生成』，属于摘要规则，only_big 必须是 false；"
        "只有用户明确说『只要/只保留/仅输出 X 亿元以上』时才是 true。"
    )
    txt, tokens, ms = _agent_chat(cfg, [{"role": "system", "content": prompt},
                                        {"role": "user", "content": text[:3000]}], timeout=60)
    m = re.search(r"\{.*\}", txt, re.S)
    data = json.loads(m.group(0)) if m else {}
    data["how"] = "模型解析（%s tokens）" % tokens
    return data


def _gd_plan(text):
    cfg = _agent_cfg()
    plan = None
    if cfg.get("api_key"):
        try:
            plan = _llm_plan(cfg, text)
        except Exception as e:  # noqa: BLE001
            plan = None
            fallback_note = "模型解析失败（%s），改用规则解析" % e
        else:
            fallback_note = ""
    else:
        fallback_note = "未配置模型 Key，使用规则解析"

    base = _regex_plan(text)
    plan = plan or {}
    out = {
        "adapter": plan.get("adapter") or base["adapter"],
        "start": str(plan.get("start") or base["start"] or ""),
        "end": str(plan.get("end") or base["end"] or ""),
        "min_amount_wan": int(plan.get("min_amount_wan") or base["min_amount_wan"]),
        "only_big": bool(re.search(r"(只保留|仅保留|只要|仅要|仅输出|只输出)", text)),
        "how": plan.get("how") or base["how"],
        "note": fallback_note,
    }
    out["start"] = re.sub(r"\D", "", out["start"])[:8]
    out["end"] = re.sub(r"\D", "", out["end"])[:8]
    if not out["start"] or not out["end"]:
        ws, we = _gd_default()
        out["start"] = out["start"] or ws
        out["end"] = out["end"] or we
    if out["start"] > out["end"]:
        out["start"], out["end"] = out["end"], out["start"]

    # 指令里出现了「已保存适配器」的域名 → 用适配器抓
    row, spec = _match_adapter(text)
    if row:
        out["adapter"] = "spec"
        out["adapter_id"] = row["id"]
        out["adapter_name"] = row["name"]
        out["spec"] = spec
    return out


# ---------------------------------------------------------------------------
# 模板读取与列映射
# ---------------------------------------------------------------------------
def _read_template(path):
    """返回模板表头（第一行）。支持 xlsx / csv。"""
    low = path.lower()
    if low.endswith(".csv"):
        raw = open(path, "rb").read()
        for enc in ("utf-8-sig", "gb18030", "utf-8"):
            try:
                text = raw.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        else:
            raise RuntimeError("CSV 编码无法识别")
        rows = list(csv.reader(io.StringIO(text)))
        return [str(c).strip() for c in rows[0]] if rows else []
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    first = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())
    return [str(c).strip() if c is not None else "" for c in first]


def _map_columns(headers):
    """模板列 → 字段。返回 {列序号: 字段}"""
    mapping, used = {}, set()
    for i, h in enumerate(headers):
        hh = re.sub(r"\s", "", h or "")
        if not hh:
            continue
        for field, keys in COL_RULES:
            if field in used:
                continue
            if any(k in hh for k in keys):
                mapping[i] = field
                used.add(field)
                break
    return mapping


def _build_row(fields, mapping, headers, min_wan):
    out = []
    for i in range(len(headers)):
        f = mapping.get(i)
        out.append(fields.get(f, "") if f else "")
    return out


def _fields_of(r):
    """抓取行 → 字段字典。兼容 13 列列表与适配器引擎的字典。"""
    if isinstance(r, dict):
        return r
    return {"code": r[0], "name": r[1], "area": r[2], "money": r[3], "scale": r[4],
            "dept": r[5], "gov": r[7], "apply": r[8], "finish": r[9],
            "years": r[10], "state": r[11], "summary": r[12]}


def _wan_of(r):
    if isinstance(r, dict):
        return float(r.get("_wan") or 0)
    return float(re.sub(r"[^\d.]", "", str(r[3] if len(r) > 3 else "")) or 0)


def _write_output(rows, headers, mapping, path, min_wan):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    wb = Workbook()
    ws = wb.active
    ws.title = "台账"
    ws.append(headers)

    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for c in ws[1]:
        c.font = Font(name="宋体", size=11, bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.fill = PatternFill("solid", fgColor="EAF0FD")
        c.border = border
    ws.row_dimensions[1].height = 43

    for r in rows:
        ws.append(_build_row(_fields_of(r), mapping, headers, min_wan))

    for r in ws.iter_rows(min_row=2):
        for c in r:
            c.font = Font(name="宋体", size=10)
            c.alignment = Alignment(vertical="center", wrap_text=True)
            c.border = border

    for i, h in enumerate(headers, start=1):
        letter = ws.cell(row=1, column=i).column_letter
        if i == 1:
            ws.column_dimensions[letter].width = 22
        elif any(k in h for k in ("规模", "内容", "摘要")):
            ws.column_dimensions[letter].width = 60
        elif any(k in h for k in ("名称", "单位", "机关")):
            ws.column_dimensions[letter].width = 30
        else:
            ws.column_dimensions[letter].width = 15
    ws.freeze_panes = "A2"
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    wb.save(path)
    return path


# ---------------------------------------------------------------------------
# 执行
# ---------------------------------------------------------------------------
def _gd_run_bg(run_id, plan, headers, mapping, template_path):
    try:
        # 没有适配器时：先自动探测站点并落库
        if plan.get("auto_url") and not plan.get("spec"):
            x("UPDATE gd_runs SET status='probing' WHERE id=?", (run_id,))
            res = ad.probe(plan["auto_url"], plan.get("hint") or "")
            spec = res["spec"]
            aid, nm = _save_adapter(spec, spec.get("name") or "")
            plan["spec"] = spec
            plan["adapter"] = "spec"
            plan["adapter_id"] = aid
            plan["adapter_name"] = "%s（自动创建）" % nm
            plan["auto_logs"] = res.get("logs", [])[-4:]
            x("UPDATE gd_runs SET plan=?, status='running' WHERE id=?",
              (json.dumps(plan, ensure_ascii=False), run_id))

        def cb(page, st):
            x("UPDATE gd_runs SET pages=?, count=? WHERE id=?", (page, st["kept"], run_id))

        min_wan = int(plan.get("min_amount_wan") or 50000)
        if plan.get("spec"):
            rows, st = ad.run_spec(plan["spec"], plan["start"], plan["end"],
                                   on_progress=cb, min_wan=min_wan)
        else:
            rows, st = gd.crawl(plan["start"], plan["end"], on_progress=cb)

        if plan.get("only_big"):
            rows = [r for r in rows if _wan_of(r) >= min_wan]
        path = os.path.join(_FILE_DIR, "gd_%s_%s_%d.xlsx" % (plan["start"], plan["end"], run_id))
        _write_output(rows, headers, mapping, path, min_wan)
        big = sum(1 for r in rows if _wan_of(r) >= min_wan)
        x("UPDATE gd_runs SET status='done', pages=?, count=?, big=?, file=?, finished=? WHERE id=?",
          (st["pages"], len(rows), big, path, now_str(), run_id))
        sset("gd_last_end", plan["end"])
        sset("gd_last_run", now_str())
    except Exception as e:  # noqa: BLE001
        traceback.print_exc()
        msg = "%s: %s" % (type(e).__name__, e)
        if plan.get("auto_url") and not plan.get("spec"):
            msg = "自动识别站点失败 —— " + str(e)
        x("UPDATE gd_runs SET status='error', err=?, finished=? WHERE id=?",
          (msg, now_str(), run_id))


def _gd_tick():
    if _gd_setting("enabled", "0") != "1":
        return
    if _gd_setting("last_run", "")[:10] == datetime.date.today().strftime("%Y-%m-%d"):
        return
    _, hol, wrk = _gd_state()
    today = datetime.date.today()
    if not gd.is_workday(today, hol, wrk):
        return
    if datetime.datetime.now().strftime("%H:%M") < _gd_setting("time", "08:30"):
        return
    start, end = _gd_window()
    if not start or one("SELECT id FROM gd_runs WHERE status IN ('running','probing')"):
        return
    headers = gd.HEADER
    mapping = _map_columns(headers)
    plan = {"adapter": "gd_tzxm", "start": start, "end": end,
            "min_amount_wan": 50000, "only_big": False, "how": "定时续跑"}
    rid = x("INSERT INTO gd_runs (start,end,status,instruction,template,headers,plan,created) "
            "VALUES (?,?,?,?,?,?,?,?)",
            (start, end, "running", "（定时续跑）", "", json.dumps(headers, ensure_ascii=False),
             json.dumps(plan, ensure_ascii=False), now_str()))
    sset("gd_last_run", now_str())
    threading.Thread(target=_gd_run_bg, args=(rid, plan, headers, mapping, ""), daemon=True).start()


def _gd_html(flash):
    runs = q("SELECT * FROM gd_runs ORDER BY id DESC LIMIT 8")
    _ld = one("SELECT id FROM gd_runs WHERE status='done' AND file IS NOT NULL "
              "AND file <> '' ORDER BY id DESC")
    last_done = _ld["id"] if _ld else 0
    running = bool(one("SELECT id FROM gd_runs WHERE status IN ('running','probing')"))
    last_end, _, _ = _gd_state()
    ws, we = _gd_window()
    enabled = _gd_setting("enabled", "0") == "1"
    prompt = _gd_setting("prompt") or DEFAULT_GD_PROMPT

    def st_pill(s):
        m = {"probing": ("p3", "识别站点中"), "running": ("p1", "执行中"),
             "done": ("p2", "已完成"), "error": ("p4", "失败")}
        c, t = m.get(s, ("p6", s or "-"))
        return "<span class='pill %s'>%s</span>" % (c, t)

    rows_html = "".join(
        "<tr><td class='cfg-num'>%d</td><td class='mini'>%s</td><td>%s</td><td>%s</td>"
        "<td class='cfg-num'>%d</td><td class='mini'>%s</td><td>%s</td></tr>"
        % (r["id"], esc((r["instruction"] or "")[:26]) + ("…" if len(r["instruction"] or "") > 26 else ""),
           "%s ~ %s" % (esc(r["start"]), esc(r["end"])), st_pill(r["status"]),
           r["count"] or 0, esc(r["created"] or ""),
           ("<a href='/gd/download/%d'>下载</a>" % r["id"]) if r["file"] and os.path.isfile(r["file"])
           else ("<span class='mini'>%s</span>" % esc((r["err"] or "")[:36])))
        for r in runs) or "<tr><td colspan='7' class='mini'>还没有执行记录</td></tr>"

    body = """
    <div class="card"><h2>专项抓取 <span class="mini">写指令 + 传模板，智能体解析后执行</span></h2>
      %s
      <div class="cfg-hint" style="line-height:2">
        <b>指令里直接写网址</b>就行：系统会自己识别站点、自动生成适配器（约 5-30 秒），然后开始抓，
        抓完在下面「抓取回来的信息」里直接显示结果。<br>
        已抓过的站点会记住，下次不再重复识别。<br>
        模板决定输出哪些列，列名按关键字自动对应。
      </div>

      <form method="post" action="/gd/run" enctype="multipart/form-data" data-ajax="1" id="gdForm">
        <label class="cfg-mini" style="margin-top:12px">抓取指令</label>
        <textarea name="instruction" rows="6" style="font-family:Microsoft YaHei;line-height:1.9">%s</textarea>
        <div class="cfg-bar" style="margin-top:10px">
          <span class="cfg-mini">模板文件（xlsx / csv，列名决定输出结构）</span>
          <input type="file" name="template" accept=".xlsx,.xls,.csv" style="max-width:280px">
          <a class="btn btn3 btn-sm" href="/gd/template">⬇ 下载标准模板</a>
          <span class="cfg-mini" style="margin-left:auto">留空则按指令里提到的列输出，没提用标准 13 列</span>
        </div>
        <div class="cfg-bar" style="margin-top:10px">
          <button class="btn" type="submit">▶ 开始执行</button>
          <span class="cfg-mini">上次抓到：<b>%s</b>　续跑窗口：<b>%s</b></span>
        </div>
      </form>

      <div class="cfg-prog"><i id="gdBar" style="width:%s"></i></div>
      <div class="cfg-mini" id="gdStatus">%s</div>
      <div id="planBox" style="display:none">
        <div class="cfg-sub">解析结果</div>
        <div class="cfg-pre" id="planText" style="padding:12px 14px;font-size:12.5px"></div>
      </div>

      <div id="pvBox" style="display:none;margin-top:14px">
        <div class="cfg-sub" style="margin-bottom:6px">抓取回来的信息
          <span class="cfg-mini" id="pvMeta" style="font-weight:400"></span></div>
        <textarea id="pvText" readonly rows="12"
          style="width:100%%;font-family:Consolas,Microsoft YaHei,monospace;font-size:12px;
          line-height:1.7;background:#fbfcfe;border:1px solid #e5e9f0;border-radius:8px;padding:10px 12px"></textarea>
        <div class="cfg-bar" style="margin-top:6px">
          <button type="button" class="btn btn3 btn-sm" onclick="copyPv()">复制全部</button>
          <span class="cfg-mini">制表符分隔，可直接粘到 Excel；完整数据点下方执行记录的「下载」</span>
        </div>
      </div>
    </div>

    <div class="card"><h2>定时续跑 <span class="mini">工作日执行，周末节假日顺延</span></h2>
      <form method="post" action="/gd/sched">
        <div class="row">
          <div><label>启用自动抓取</label>
            <select name="enabled"><option value="1" %s>启用</option><option value="0" %s>停用</option></select></div>
          <div><label>每天执行时间</label><input type="time" name="time" value="%s"></div>
          <div style="flex:2"><label>节假日（停抓，逗号分隔）</label>
            <input name="holidays" value="%s" placeholder="如 20261001,20261002"></div>
          <div style="flex:2"><label>调休上班日（周末也抓）</label>
            <input name="workdays" value="%s" placeholder="如 20261011"></div>
        </div>
        <div class="cfg-hint" style="margin-top:8px">
          自动任务固定按「广东省备案公开 + 标准 13 列」跑；窗口＝上次抓到日 +1 ~ 昨天，漏跑自动补齐。
        </div>
        <div style="margin-top:12px"><button class="btn btn2">保存设置</button></div>
      </form>
    </div>

    <div class="card"><h2>执行记录</h2>
      <div class="tbl-wrap"><table class="cfg">
        <tr><th>ID</th><th>指令</th><th>区间</th><th>状态</th><th>条数</th><th>时间</th><th>文件</th></tr>
        %s</table></div>
    </div>

    <script>
    (function(){
      var F=document.getElementById('gdForm');
      if(!F) return;
      F.addEventListener('submit',function(ev){
        ev.preventDefault();
        var btn=F.querySelector('button[type=submit]');
        btn.disabled=true; btn.textContent='解析指令中…';
        document.getElementById('gdStatus').textContent='正在解析指令并抓取，请稍候（1 天约 1 分钟）…';
        document.getElementById('gdBar').style.width='10%%';
        fetch('/gd/run',{method:'POST',body:new FormData(F)}).then(function(r){return r.json()}).then(function(d){
          btn.disabled=false; btn.textContent='▶ 开始执行';
          if(d.error){document.getElementById('gdStatus').textContent='失败：'+d.error;
            document.getElementById('gdBar').style.width='0';return;}
          if(d.plan){showPlan(d.plan);}
          poll(d.id);
        }).catch(function(e){
          btn.disabled=false; btn.textContent='▶ 开始执行';
          document.getElementById('gdStatus').textContent='提交失败：'+e;
        });
      });
      function showPlan(p){
        document.getElementById('planBox').style.display='';
        document.getElementById('planText').textContent=
          '适配器：'+(p.adapter_name||p.adapter||'未识别')+'\\n区间：'+p.start+' ~ '+p.end+
          '\\n金额门槛：'+(p.min_amount_wan/10000)+' 亿元'+(p.only_big?'（只保留达标）':'（全量保留，摘要仅达标）')+
          '\\n模板列：'+(p.headers||[]).join(' ｜ ')+
          '\\n解析方式：'+p.how+(p.note?('（'+p.note+'）'):'');
      }
      function poll(id){
        fetch('/gd/run/'+id+'/json').then(function(r){return r.json()}).then(function(d){
          if(d.status==='probing'){
            document.getElementById('gdStatus').textContent=
              '正在识别站点、自动创建适配器…（约 5-30 秒，成功后开始抓取）';
            document.getElementById('gdBar').style.width='12%%';
            setTimeout(function(){poll(id)},1500);
            return;
          }
          document.getElementById('gdStatus').textContent=
            '任务 #'+d.id+' '+d.status+'：已扫 '+d.pages+' 页，命中 '+d.count+' 条';
          document.getElementById('gdBar').style.width=(d.status==='done'?'100%%':Math.min(95,d.pages)+'%%');
          if(d.status==='running'){setTimeout(function(){poll(id)},1500);}
          else if(d.status==='error'){document.getElementById('gdStatus').textContent='失败：'+d.err;}
          else{loadPv(id);}
        });
      }
      function loadPv(id){
        fetch('/gd/run/'+id+'/preview?limit=50').then(function(r){return r.json()}).then(function(d){
          if(!d.ok){return;}
          document.getElementById('pvBox').style.display='';
          document.getElementById('pvMeta').textContent=
            '　共 '+d.total+' 条，下面显示前 '+d.shown+' 条';
          document.getElementById('pvText').value=d.text;
        });
      }
      function copyPv(){
        var t=document.getElementById('pvText'); if(!t) return;
        t.select(); t.setSelectionRange(0,999999);
        try{document.execCommand('copy');}catch(e){}
        if(navigator.clipboard){navigator.clipboard.writeText(t.value);}
        document.getElementById('pvMeta').textContent='　已复制到剪贴板';
      }
      if(%s){poll(%s);}
      else if(%s){loadPv(%s);}
    })();
    </script>
    """ % (flash, esc(prompt),
           esc(last_end or "—"), ("%s ~ %s" % (ws, we)) if ws else "已抓到昨天",
           "8" if running else "0", "处理中…" if running else "尚未执行",
           "selected" if enabled else "", "selected" if not enabled else "",
           esc(_gd_setting("time", "08:30")),
           esc(_gd_setting("holidays")), esc(_gd_setting("workdays")),
           rows_html,
           "true" if running else "false",
           (one("SELECT id FROM gd_runs WHERE status IN ('running','probing')")["id"]
            if running else 0),
           "true" if last_done else "false", (last_done or 0))
    return body


@bp.route("/gd/run", methods=["POST"])
def gd_run():
    if one("SELECT id FROM gd_runs WHERE status IN ('running','probing')"):
        return jsonify({"error": "已有任务在跑，请等它结束"})

    instruction = (request.form.get("instruction") or "").strip()
    if len(instruction) < 10:
        return jsonify({"error": "请写清抓取指令（网址、日期范围、输出要求）"})

    # 模板
    f = request.files.get("template")
    template_path, headers = "", []
    if f and f.filename:
        ext = os.path.splitext(f.filename)[1].lower()
        if ext not in (".xlsx", ".xls", ".csv"):
            return jsonify({"error": "模板只支持 xlsx / csv"})
        template_path = _save_upload(f)
        try:
            headers = _read_template(template_path)
        except Exception as e:  # noqa: BLE001
            return jsonify({"error": "模板读取失败：%s" % e})
        headers = [h or "" for h in headers]
        if not any(headers):
            return jsonify({"error": "模板第一行没有列名"})
    else:
        headers = _cols_from_text(instruction) or list(gd.HEADER)

    plan = _gd_plan(instruction)

    # 没匹配到已保存的适配器时：从指令里找网址，交给后台自动识别
    if plan["adapter"] != "gd_tzxm" and not plan.get("spec"):
        urls = _urls_in(instruction)
        if urls:
            plan["adapter"] = "auto"
            plan["auto_url"] = urls[0]
            plan["adapter_name"] = "自动识别中…"
        else:
            return jsonify({"error": "指令里没写网址。直接写成「抓 https://… 2026-10-27 到 10-29 的记录，"
                                     "按模板输出」，系统会自己识别站点。"})

    sset("gd_prompt", instruction)
    mapping = _map_columns(headers)
    plan["headers"] = headers
    plan["mapping"] = {str(k): v for k, v in mapping.items()}

    rid = x("INSERT INTO gd_runs (start,end,status,instruction,template,headers,plan,created) "
            "VALUES (?,?,?,?,?,?,?,?)",
            (plan["start"], plan["end"], "running", instruction[:1000], template_path,
             json.dumps(headers, ensure_ascii=False), json.dumps(plan, ensure_ascii=False), now_str()))
    threading.Thread(target=_gd_run_bg, args=(rid, plan, headers, mapping, template_path),
                     daemon=True).start()
    return jsonify({"id": rid, "plan": plan})


@bp.route("/gd/status")
def gd_status():
    r = one("SELECT * FROM gd_runs ORDER BY id DESC LIMIT 1")
    if not r:
        return jsonify({})
    return jsonify({"id": r["id"], "status": r["status"], "pages": r["pages"] or 0,
                    "count": r["count"] or 0, "big": r["big"] or 0, "err": r["err"] or ""})


@bp.route("/gd/run/<int:rid>/json")
def gd_run_json(rid):
    r = one("SELECT * FROM gd_runs WHERE id=?", (rid,))
    if not r:
        return jsonify({"error": "not found"})
    return jsonify({"id": r["id"], "status": r["status"], "pages": r["pages"] or 0,
                    "count": r["count"] or 0, "big": r["big"] or 0, "err": r["err"] or ""})


@bp.route("/gd/run/<int:rid>/preview")
def gd_preview(rid):
    """把结果文件前 N 行读出来，给页面文本框显示。"""
    r = one("SELECT * FROM gd_runs WHERE id=?", (rid,))
    if not r or not r["file"] or not os.path.isfile(r["file"]):
        return jsonify({"ok": False, "error": "还没有结果文件"})
    try:
        limit = max(1, min(int(request.args.get("limit", 50)), 500))
    except (TypeError, ValueError):
        limit = 50
    try:
        import openpyxl
        wb = openpyxl.load_workbook(r["file"], read_only=True, data_only=True)
        ws = wb.active
        lines, total = [], 0
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            cells = ["" if c is None else str(c).replace("\n", " ").replace("\t", " ")
                     for c in row]
            if i == 0:
                lines.append("\t".join(cells))
                continue
            total += 1
            if total <= limit:
                lines.append("\t".join(cells))
        wb.close()
    except Exception as e:  # noqa: BLE001
        return jsonify({"ok": False, "error": "读取结果失败：%s" % e})
    return jsonify({"ok": True, "text": "\n".join(lines), "total": total,
                    "shown": min(total, limit)})


@bp.route("/gd/download/<int:rid>")
def gd_download(rid):
    r = one("SELECT * FROM gd_runs WHERE id=?", (rid,))
    if not r or not r["file"] or not os.path.isfile(r["file"]):
        abort(404)
    name = "备案台账_%s_%s.xlsx" % (r["start"], r["end"])
    return send_file(r["file"], as_attachment=True, download_name=name)


@bp.route("/gd/template")
def gd_template():
    path = os.path.join(_FILE_DIR, "固贷备案台账模板.xlsx")
    if not os.path.isfile(path):
        gd.write_template(path)
    return send_file(path, as_attachment=True, download_name="固贷备案台账模板.xlsx")


@bp.route("/gd/sched", methods=["POST"])
def gd_sched():
    sset("gd_enabled", "1" if request.form.get("enabled") == "1" else "0")
    sset("gd_time", request.form.get("time") or "08:30")
    sset("gd_holidays", (request.form.get("holidays") or "").strip())
    sset("gd_workdays", (request.form.get("workdays") or "").strip())
    return _back("ok:定时设置已保存", "gd")

# ---------------------------------------------------------------------------
# 八、平台适配：给网址，自动生成抓取适配器
# ---------------------------------------------------------------------------
def _adapt_rows():
    return q("SELECT * FROM adapters ORDER BY id DESC")


def _match_adapter(text):
    """指令里提到某个已保存适配器的域名/名称，就用它。"""
    t = (text or "").lower()
    for r in _adapt_rows():
        if not r["enabled"]:
            continue
        host = (r["host"] or "").lower()
        if host and host in t:
            try:
                return r, json.loads(r["spec"])
            except ValueError:
                continue
        if r["name"] and r["name"] in (text or ""):
            try:
                return r, json.loads(r["spec"])
            except ValueError:
                continue
    return None, None


def _adapt_html(flash):
    rows = _adapt_rows()
    trs = ""
    for r in rows:
        try:
            spec = json.loads(r["spec"])
        except ValueError:
            spec = {}
        fm = spec.get("field_map") or {}
        trs += (
            "<tr><td><b>" + esc(r["name"]) + "</b>"
            "<div class='mini'>" + esc(r["host"]) + "</div></td>"
            "<td class='mini'>" + esc((spec.get("list_url") or "")[:58]) + "</td>"
            "<td class='cfg-num'>" + str(len(fm)) + "</td>"
            "<td class='mini'>" + esc(spec.get("method") or "-") + " · "
            + esc(spec.get("page_field") or "-") + "</td>"
            "<td class='mini'>" + esc(r["updated"] or "") + "</td>"
            "<td><a class='c-ops' href='javascript:loadSpec(" + str(r["id"]) + ")'>载入</a>"
            "<a class='c-ops del' href='/adapt/delete/" + str(r["id"]) + "' "
            "onclick=\"return confirm('删除这个适配器？')\">删除</a></td></tr>"
        )
    if not trs:
        trs = "<tr><td colspan='6' class='mini'>还没有适配器。上面填一个要抓的网址，点「探测接口」。</td></tr>"

    tips = (
        "<div class='cfg-hint' style='margin-top:10px;line-height:1.9'>"
        "探测先找接口：打开页面 → 从页面和脚本里提候选接口 → 并发试，看哪个返回列表数据，"
        "再确认翻页参数和字段对应。<br>"
        "找不到接口时（服务端渲染的站点），自动改成<b>解析页面列表</b>：按 &lt;li&gt;/&lt;tr&gt; 切块，"
        "只保留「有标题链接 + 有日期」的块，并顺着「下一页」翻页。<br>"
        "<b>处理不了的</b>：接口要签名、要登录、要 JS 渲染才有内容。这类会明确提示失败，"
        "可以把已知接口地址填进来再探测。"
        "</div>"
    )

    return (
        "<div class='card'><h2>平台适配 <span class='mini'>给一个网址，自动识别怎么抓</span></h2>"
        + flash +
        "<form id='adaptForm' data-ajax='1' style='margin-top:4px'>"
        "<div class='cfg-bar'>"
        "<input id='aUrl' placeholder='粘贴要抓的列表页网址，例如 https://tzxm.gd.gov.cn/index.html' "
        "style='flex:1;min-width:320px'>"
        "<input id='aHint' placeholder='补充说明（可选）：比如「备案公开」「按发布时间倒序」' style='width:260px'>"
        "<button type='button' class='btn btn2 btn-sm' onclick='doProbe()'>探测接口</button>"
        "</div></form>"
        "<div id='prBar' class='cfg-prog' style='display:none'><i></i></div>"
        "<div id='prLog' class='cfg-log' style=\"display:none;max-height:170px;"
        "white-space:pre-wrap;font-size:11.5px;line-height:1.75\"></div>"
        "<div id='prRes' style='display:none'>"
        "  <div class='cfg-bar' style='margin-top:12px'>"
        "    <b style='font-size:13px;color:#12224a'>识别结果</b>"
        "    <span id='prMeta' class='cfg-mini'></span>"
        "    <span class='sp'></span>"
        "    <button type='button' class='btn btn3 btn-sm' onclick='doTest()'>试抓一页</button>"
        "    <button type='button' class='btn btn2 btn-sm' onclick='doSave()'>保存适配器</button>"
        "    <button type='button' class='btn btn3 btn-sm' onclick='toggleSpec()'>手工调整</button>"
        "  </div>"
        "  <textarea id='specBox' rows='10' style='display:none;font-family:Consolas,monospace;"
        "font-size:12px;line-height:1.7'></textarea>"
        "  <div id='pvWrap' style='margin-top:10px'></div>"
        "</div>"
        + tips + "</div>"

        "<div class='card'><h2>已保存的适配器 <span class='mini'>共 " + str(len(rows)) + " 个</span></h2>"
        "<div class='tbl-wrap'><table class='cfg'>"
        "<tr><th>名称</th><th>接口</th><th>字段</th><th>方法 / 翻页</th><th>更新时间</th><th>操作</th></tr>"
        + trs + "</table></div>"
        "<div class='cfg-hint' style='margin-top:10px'>保存后到「专项抓取」里，指令中写上这个网站的域名，"
        "系统就会用对应适配器抓取；模板仍然按你上传的列来输出。</div></div>"

        "<script>"
        "var LAST=null;"
        "function pb(p){var b=document.getElementById('prBar');b.style.display='';"
        "b.firstElementChild.style.width=(p<0?100:p)+'%';}"
        "function lg(t){var d=document.getElementById('prLog');d.style.display='';"
        "d.textContent=t;}"
        "function toggleSpec(){var t=document.getElementById('specBox');"
        "t.style.display=(t.style.display==='none'?'':'none');}"
        "function doProbe(){"
        " var u=document.getElementById('aUrl').value.trim();"
        " if(!u){alert('请先填网址');return;}"
        " document.getElementById('prRes').style.display='none';"
        " document.getElementById('prLog').textContent='正在打开页面、寻找接口…（约 10-30 秒）';"
        " document.getElementById('prLog').style.display='';pb(30);"
        " fetch('/adapt/probe',{method:'POST',headers:{'Content-Type':'application/json'},"
        "  body:JSON.stringify({url:u,hint:document.getElementById('aHint').value})})"
        " .then(function(r){return r.json()}).then(function(d){"
        "  if(!d.ok){pb(-1);lg('探测失败：'+(d.error||'未知原因'));return;}"
        "  LAST=d.spec;"
        "  document.getElementById('prLog').textContent=(d.logs||[]).join('\\n')"
        "    +((d.warnings&&d.warnings.length)?('\\n\\n注意：'+d.warnings.join('；')):'');"
        "  document.getElementById('prRes').style.display='';"
        "  document.getElementById('specBox').value=JSON.stringify(d.spec,null,2);"
        "  document.getElementById('prMeta').textContent='置信度 '+(d.spec.confidence||'-')"
        "    +' · 字段 '+Object.keys(d.spec.field_map||{}).length+' 个';"
        "  pb(100);renderPreview(d.preview_rows||d.preview||[]);"
        " }).catch(function(e){pb(-1);lg('请求出错：'+e);});"
        "}"
        "function specOf(){var t=document.getElementById('specBox');"
        " if(t.style.display!=='none'&&t.value.trim()){try{return JSON.parse(t.value)}catch(e){alert('JSON 格式有误：'+e);return null}}"
        " return LAST;}"
        "function doTest(){var s=specOf();if(!s)return;"
        " lg('试抓第一页…');"
        " fetch('/adapt/test',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({spec:s})})"
        " .then(function(r){return r.json()}).then(function(d){"
        "  if(!d.ok){lg('试抓失败：'+d.error);return;}"
        "  lg('第一页返回 '+d.count+' 条，下面是前 5 条解析结果');renderPreview(d.rows||[]);"
        " });}"
        "function doSave(){var s=specOf();if(!s)return;"
        " fetch('/adapt/save',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({spec:s})})"
        " .then(function(r){return r.json()}).then(function(d){"
        "  if(!d.ok){lg('保存失败：'+d.error);return;}"
        "  lg('已保存，2 秒后刷新…');setTimeout(function(){location.href='/agent?tab=adapt&msg=ok:'"
        "+encodeURIComponent('适配器「'+(d.name||'')+'」已保存')},900);"
        " });}"
        "function loadSpec(id){fetch('/adapt/'+id+'/json').then(function(r){return r.json()}).then(function(d){"
        "  if(!d.ok){alert(d.error);return;}"
        "  LAST=d.spec;document.getElementById('prRes').style.display='';"
        "  document.getElementById('specBox').value=JSON.stringify(d.spec,null,2);"
        "  document.getElementById('specBox').style.display='';"
        "  document.getElementById('prMeta').textContent='已载入「'+d.name+'」，可试抓或修改后保存';"
        "  document.getElementById('aUrl').value=d.spec.page_url||'';"
        "  document.getElementById('prLog').style.display='none';window.scrollTo(0,0);"
        " });}"
                "function renderPreview(rows){"
        " if(!rows||!rows.length){document.getElementById('pvWrap').innerHTML='';return;}"
        " var cols=[['编号','code'],['名称','name'],['所在地','area'],['总投资','money'],['建设单位','dept'],['通过日期','finish'],['摘要','summary']];"
        " var use=cols.filter(function(c){return rows.some(function(r){var v=r[c[1]];return v!==undefined&&v!==null&&String(v)!=='';})});"
        " var out='<div class=tbl-wrap><table class=cfg><tr>';"
        " use.forEach(function(c){out+='<th>'+c[0]+'</th>'});"
        " out+='</tr>';"
        " rows.forEach(function(r){out+='<tr>';use.forEach(function(c){"
        "   var v=r[c[1]];v=(v===undefined||v===null)?'':String(v);"
        "   var cls=(c[1]==='name')?'':((c[1]==='money')?'cfg-num':'mini');"
        "   out+='<td'+(cls?(' class='+cls):'')+(c[1]==='finish'?' style=white-space:nowrap':'')+'>'+(v||'—')+'</td>';});"
        "  out+='</tr>';});"
        " document.getElementById('pvWrap').innerHTML=out+'</table></div>';}"
        "</script>"
    )


@bp.route("/adapt/probe", methods=["POST"])
def adapt_probe():
    data = request.get_json(silent=True) or {}
    url = (data.get("url") or "").strip()
    if not url:
        return jsonify({"ok": False, "error": "请填网址"})
    try:
        res = ad.probe(url, data.get("hint") or "")
    except Exception as e:  # noqa: BLE001
        traceback.print_exc()
        return jsonify({"ok": False, "error": "%s: %s" % (type(e).__name__, e)})
    return jsonify(res)


@bp.route("/adapt/test", methods=["POST"])
def adapt_test():
    data = request.get_json(silent=True) or {}
    try:
        spec = ad.normalize_spec(data.get("spec") or {})
    except Exception as e:  # noqa: BLE001
        return jsonify({"ok": False, "error": "spec 有问题：%s" % e})
    return jsonify(ad.test_spec(spec))


@bp.route("/adapt/save", methods=["POST"])
def adapt_save():
    data = request.get_json(silent=True) or {}
    try:
        spec = ad.normalize_spec(data.get("spec") or {})
    except Exception as e:  # noqa: BLE001
        return jsonify({"ok": False, "error": "spec 有问题：%s" % e})
    aid, name = _save_adapter(spec, str(data.get("name") or ""))
    return jsonify({"ok": True, "id": aid, "name": name})


@bp.route("/adapt/<int:aid>/json")
def adapt_json(aid):
    r = one("SELECT * FROM adapters WHERE id=?", (aid,))
    if not r:
        return jsonify({"ok": False, "error": "不存在"})
    try:
        spec = json.loads(r["spec"])
    except ValueError:
        return jsonify({"ok": False, "error": "保存的 spec 不是合法 JSON"})
    return jsonify({"ok": True, "name": r["name"], "spec": spec})


@bp.route("/adapt/delete/<int:aid>")
def adapt_delete(aid):
    x("DELETE FROM adapters WHERE id=?", (aid,))
    return _back("ok:已删除适配器", "adapt")

