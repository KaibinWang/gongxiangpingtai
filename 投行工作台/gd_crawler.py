# -*- coding: utf-8 -*-
"""
广东省投资项目在线审批监管平台 · 备案公开抓取

数据源：POST /tzxmspweb/api/publicityInformation/selectByPageBANew
（列表接口本身已含详情页全部字段，详情接口被风控封禁，不要调用）

特点：
- 单线程 + 随机间隔 + 失败退避，避免触发风控
- 倒序翻页，遇到 finishDate < 起始日即停（当页处理完再停）
- 输出 13 列，与「【自动爬取】固贷备案台账模板.xlsx」一致
- 最后一列只对总投资 ≥5 亿元的项目生成摘要
"""

from __future__ import annotations

import datetime
import json
import os
import random
import re
import ssl
import time
import urllib.error
import urllib.request

API = "https://tzxm.gd.gov.cn/tzxmspweb/api/publicityInformation/selectByPageBANew"
REFERER = "https://tzxm.gd.gov.cn/PublicityInformation/PublicityHandlingResultsList.html"
UAS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36 Edg/120.0",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Safari/605.1.15",
]

HEADER = ["备案项目编号", "项目名称", "项目所在地", "项目总投资", "项目规模及内容", "建设单位", "",
          "备案机关", "备案申报日期", "复核通过日期", "项目起止年限", "项目当前状态", "5亿元以上项目摘要"]

BIG_LINE = 50000.0          # 5 亿元 = 50000 万元


class RiskControl(Exception):
    """接口返回风控提示。"""


def _ctx():
    c = ssl.create_default_context()
    c.check_hostname = False
    c.verify_mode = ssl.CERT_NONE
    return c


_CTX = _ctx()


def fetch_page(page_number, page_size=15, timeout=30):
    body = json.dumps({"flag": "1", "nameOrCode": "", "city": "",
                       "pageSize": page_size, "pageNumber": page_number}).encode()
    req = urllib.request.Request(API, data=body, method="POST", headers={
        "Content-Type": "application/json;charset=utf-8",
        "X-Requested-With": "XMLHttpRequest",
        "Referer": REFERER,
        "Origin": "https://tzxm.gd.gov.cn",
        "Accept": "application/json, text/javascript, */*; q=0.01",
        "User-Agent": random.choice(UAS),
    })
    with urllib.request.urlopen(req, timeout=timeout, context=_CTX) as r:
        raw = r.read().decode("utf-8", "replace")
    try:
        d = json.loads(raw)
    except json.JSONDecodeError:
        raise RiskControl("返回不是 JSON，可能被 WAF 拦截：%s" % raw[:120])
    data = d.get("data")
    if isinstance(data, str):
        raise RiskControl(data)
    if not isinstance(data, dict):
        raise RiskControl("data 结构异常：%s" % str(data)[:120])
    return data


def fetch_page_retry(page_number, tries=4):
    delay = 2.0
    last = None
    for i in range(tries):
        try:
            return fetch_page(page_number)
        except Exception as e:  # noqa: BLE001
            last = e
            time.sleep(delay)
            delay *= 2
    raise last if last else RuntimeError("抓取失败")


# ---------------------------------------------------------------------------
# 字段整理
# ---------------------------------------------------------------------------
def ymd(v):
    """2026-10-30 / 2026/10/30 10:00:00 → '20261030'"""
    s = re.sub(r"[^\d]", "", str(v or "")[:10])
    return s[:8]


def money_text(v):
    try:
        n = float(v)
    except (TypeError, ValueError):
        return ""
    if n <= 0:
        return ""
    if abs(n - round(n)) < 0.01:
        return "%d万元" % int(round(n))
    return "%.2f万元" % n


def area_text(r):
    """项目所在地收敛到「市 + 区/县」。"""
    raw = (r.get("areaDetailName") or "").strip()
    if not raw:
        raw = (r.get("admdivName") or "").strip()
    m = re.match(r"^(.{2,8}?市)(.*)$", raw)
    if not m:
        return raw
    city, rest = m.group(1), m.group(2)
    # 直筒子市不下设区
    if city in ("东莞市", "中山市", "儋州市", "嘉峪关市"):
        return city
    # 特殊功能区映射
    for key, name in (("大亚湾", "大亚湾区"), ("松山湖", "松山湖"), ("横琴", "横琴粤澳深度合作区")):
        if rest.startswith(key):
            return city + name
    m2 = re.match(r"^(.{1,4}?(?:自治县|区|县|旗))", rest)
    if m2:
        return city + m2.group(1)
    return city + rest[:3] if rest else city


def summary_text(r, area):
    """5 亿元以上项目的摘要句。"""
    try:
        big = float(r.get("totalMoney") or 0) / 10000.0
    except (TypeError, ValueError):
        return ""
    dept = (r.get("projectDept") or "").strip()
    name = (r.get("projectName") or "").strip()
    scale = re.sub(r"\s+", "", (r.get("scaleContent") or "")).rstrip("。；;")
    yi = ("%.2f" % big).rstrip("0").rstrip(".")
    return "%s拟建设%s，%s，项目地位于%s，总投资%s亿元。" % (dept, name, scale, area, yi)


def to_row(r):
    area = area_text(r)
    total = r.get("totalMoney")
    try:
        big = float(total or 0) >= BIG_LINE
    except (TypeError, ValueError):
        big = False
    years = ""
    if r.get("projectStarttimeS") or r.get("projectEndtimeS"):
        years = "%s ~ %s" % (r.get("projectStarttimeS") or "", r.get("projectEndtimeS") or "")
    return [
        r.get("projectCode") or "",
        r.get("projectName") or "",
        area,
        money_text(total),
        (r.get("scaleContent") or "").strip(),
        r.get("projectDept") or "",
        "",
        r.get("approveUnitName") or r.get("fullName") or "",
        (str(r.get("applyTime") or "")[:10]).replace("/", "-"),
        (str(r.get("finishDate") or "")[:10]).replace("/", "-"),
        years,
        r.get("stateFlagName") or "",
        summary_text(r, area) if big else "",
    ]


# ---------------------------------------------------------------------------
# 抓取主流程
# ---------------------------------------------------------------------------
def crawl(start, end, on_progress=None, max_pages=600, sleep_range=(0.8, 1.8)):
    """
    start / end: 'YYYYMMDD'
    返回 (rows, stats)；rows 已按 finishDate 倒序。
    """
    rows, seen = [], set()
    page = 1
    stop = False
    stats = {"pages": 0, "fetched": 0, "kept": 0, "skipped_before": 0}
    while not stop and page <= max_pages:
        data = fetch_page_retry(page)
        items = data.get("list") or []
        if not items:
            break
        for r in items:
            stats["fetched"] += 1
            fd = ymd(r.get("finishDate"))
            if fd and fd < start:
                stats["skipped_before"] += 1
                stop = True
                continue
            if fd and start <= fd <= end:
                key = r.get("baId") or r.get("id") or r.get("projectCode")
                if key in seen:
                    continue
                seen.add(key)
                rows.append(to_row(r))
                stats["kept"] += 1
        stats["pages"] = page
        if on_progress:
            on_progress(page, stats)
        if stop:
            break
        page += 1
        time.sleep(random.uniform(*sleep_range))
    rows.sort(key=lambda x: x[9] or "", reverse=True)
    return rows, stats


# ---------------------------------------------------------------------------
# 写 Excel（与模板 13 列一致）
# ---------------------------------------------------------------------------
def write_xlsx(rows, path, title="备案公开"):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    wb = Workbook()
    ws = wb.active
    ws.title = title
    ws.append(HEADER)

    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    head_fill = PatternFill("solid", fgColor="EAF0FD")

    for c in ws[1]:
        c.font = Font(name="宋体", size=11, bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.fill = head_fill
        c.border = border
    ws.row_dimensions[1].height = 43

    for row in rows:
        ws.append(row)

    for r in ws.iter_rows(min_row=2):
        for c in r:
            c.font = Font(name="宋体", size=10)
            c.alignment = Alignment(vertical="center", wrap_text=True)
            c.border = border

    widths = [22, 34, 18, 14, 60, 30, 3, 26, 14, 14, 16, 14, 70]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w
    ws.freeze_panes = "A2"

    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    wb.save(path)
    return path


def write_template(path):
    return write_xlsx([], path, title="备案公开")


# ---------------------------------------------------------------------------
# 工作日与续跑窗口
# ---------------------------------------------------------------------------
def _d(s):
    return datetime.date(int(s[:4]), int(s[4:6]), int(s[6:8]))


def _s(d):
    return d.strftime("%Y%m%d")


def is_workday(day, holidays=(), workdays=()):
    key = _s(day)
    if key in set(workdays):
        return True
    if key in set(holidays):
        return False
    return day.weekday() < 5


def shift_workday(day, step, holidays=(), workdays=()):
    cur, moved = day, 0
    while moved < abs(step):
        cur += datetime.timedelta(days=1 if step > 0 else -1)
        if is_workday(cur, holidays, workdays):
            moved += 1
    return cur


def next_window(last_end, days=2, holidays=(), workdays=(), today=None):
    """从 last_end 之后取 days 个工作日，返回 (start, end)。"""
    today = today or datetime.date.today()
    cur = _d(last_end) if last_end else shift_workday(today, -days - 1, holidays, workdays)
    start = shift_workday(cur, 1, holidays, workdays)
    end = shift_workday(start, days - 1, holidays, workdays)
    return _s(start), _s(end)
