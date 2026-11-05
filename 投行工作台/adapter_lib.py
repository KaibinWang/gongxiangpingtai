# -*- coding: utf-8 -*-
"""
平台探测器 + 通用适配器引擎
===========================

给一个网址，自动找出「哪个接口在返回列表数据、分页参数叫什么、字段怎么对应」，
产出一份可保存、可复用的 adapter spec（JSON）。

之后抓取不再写死代码，而是由 spec 驱动：翻页 → 取列表 → 按日期过滤 → 套模板输出。

使用：
    import adapter_lib as ad
    res = ad.probe("https://tzxm.gd.gov.cn/index.html", "备案公开")
    spec = res["spec"]
    rows, stats = ad.run_spec(spec, "20261027", "20261029")
"""

import json
import random
import re
import ssl
import time
import urllib.parse
import urllib.request
from concurrent.futures import ThreadPoolExecutor, as_completed

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

DATE_KEYS = ("finish", "pass", "approv", "issued", "public", "publish", "audit", "end")
MONEY_KEYS = ("money", "amount", "invest", "total", "capital", "price")
NAME_KEYS = ("name", "title", "proname", "projname", "subject")


class ProbeError(Exception):
    pass


# ---------------------------------------------------------------------------
# 基础请求
# ---------------------------------------------------------------------------
def _ctx():
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    return ctx


def _req(url, method="GET", headers=None, data=None, timeout=15):
    h = {
        "User-Agent": UA,
        "Accept": "application/json, text/plain, */*",
        "Accept-Language": "zh-CN,zh;q=0.9",
        "X-Requested-With": "XMLHttpRequest",
    }
    h.update(headers or {})
    if data is not None and not isinstance(data, bytes):
        data = json.dumps(data, ensure_ascii=False).encode("utf-8")
        h.setdefault("Content-Type", "application/json;charset=utf-8")
    req = urllib.request.Request(url, data=data, method=method, headers=h)
    try:
        with urllib.request.urlopen(req, timeout=timeout, context=_ctx()) as r:
            raw = r.read()
            ctype = r.headers.get("Content-Type", "")
    except urllib.error.HTTPError as e:
        raw = e.read() or b""
        ctype = e.headers.get("Content-Type", "") if e.headers else ""
        if e.code >= 400 and not raw:
            raise ProbeError("HTTP %s" % e.code)
    except Exception as e:  # noqa: BLE001
        raise ProbeError("%s: %s" % (type(e).__name__, e))

    for enc in ("utf-8", "gb18030"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        text = raw.decode("utf-8", "replace")
    return text, ctype


def _try_json(text):
    t = (text or "").strip()
    if not t or t[0] not in "[{":
        return None
    try:
        return json.loads(t)
    except ValueError:
        return None


# ---------------------------------------------------------------------------
# 找列表
# ---------------------------------------------------------------------------
def _walk(obj, path=""):
    """深度遍历，产出 (点分路径, 值)。"""
    if isinstance(obj, dict):
        for k, v in obj.items():
            p = "%s.%s" % (path, k) if path else str(k)
            yield p, v
            yield from _walk(v, p)
    elif isinstance(obj, list):
        for i, v in enumerate(obj[:5]):
            p = "%s[%d]" % (path, i) if path else "[%d]" % i
            yield from _walk(v, p)


def _pick_list(payload):
    """从 JSON 里挑「最长的一批对象」当列表。返回 (路径, 列表)。"""
    best_path, best = "", []
    for path, val in _walk(payload):
        if not isinstance(val, list) or not val:
            continue
        dicts = [x for x in val if isinstance(x, dict)]
        if len(dicts) < max(1, len(val) * 0.8):
            continue
        if len(dicts) > len(best):
            best_path, best = path, dicts
    return best_path, best


def _path_get(obj, path):
    if not path:
        return obj
    cur = obj
    for part in path.replace("[", ".").replace("]", "").split("."):
        if not part:
            continue
        if isinstance(cur, list):
            try:
                cur = cur[int(part)]
            except (ValueError, IndexError):
                return None
        elif isinstance(cur, dict):
            if part not in cur:
                return None
            cur = cur[part]
        else:
            return None
    return cur


def _numbers(obj, keys=("total", "count", "pages", "page")):
    """找同级的分页数字。"""
    out = {}
    if isinstance(obj, dict):
        for k, v in obj.items():
            if isinstance(v, (int, float)) and any(t in str(k).lower() for t in keys):
                out[k] = int(v)
    return out


# ---------------------------------------------------------------------------
# 抓页面与候选接口
# ---------------------------------------------------------------------------
def _html(url):
    text, ctype = _req(url, headers={"Accept": "text/html,*/*"})
    if "json" in ctype:
        raise ProbeError("这个地址直接返回了 JSON，请在网址里填页面地址")
    m = re.search(r"<title[^>]*>(.*?)</title>", text, re.S | re.I)
    title = _text_of(m.group(1)) if m else ""
    title = (title.replace("&nbsp;", " ").replace("&amp;", "&")
                  .replace("&lt;", "<").replace("&gt;", ">").strip())
    return text, title


def _api_candidates(html, base_url, extra_js=()):
    """从页面 HTML / JS 里找候选接口。"""
    blobs = [html]
    for js in extra_js[:4]:
        try:
            t, _ = _req(js, headers={"Accept": "*/*"})
            if len(t) < 3_000_000:
                blobs.append(t)
        except ProbeError:
            continue

    keys = ("api", "select", "list", "query", "search", "page", "json", "data", "servlet")
    bad = (".css", ".png", ".jpg", ".gif", ".svg", ".ico", ".woff", ".ttf", ".map",
           ".html", ".htm", ".pdf", ".js")
    found, seen = [], set()
    for blob in blobs:
        for m in re.finditer(r"""["'`]((?:https?://|//)[^"'`\s]{6,180}|/[A-Za-z0-9_\-./]{4,120})["'`]""", blob):
            u = m.group(1).strip()
            if u.lower().endswith(bad) or " " in u:
                continue
            low = u.lower()
            if not any(k in low for k in keys):
                continue
            full = urllib.parse.urljoin(base_url, u)
            if full in seen:
                continue
            seen.add(full)
            found.append(full)

    # 打分：越像「列表接口」越靠前
    def score(u):
        low = u.lower()
        s = 0
        for k, w in (("select", 5), ("list", 4), ("search", 4), ("query", 3),
                     ("page", 3), ("api", 2), ("json", 2), ("data", 1)):
            if k in low:
                s += w
        if u.count("/") <= 3:
            s -= 2
        if "/api/" in low:
            s += 6
        if any(x in low for x in ("upload", "login", "logout", "help")) or "getbackurl" in low:
            s -= 7
        return -s

    found.sort(key=score)
    return found[:14]


def _js_files(html, base_url):
    out = []
    for m in re.finditer(r"""<script[^>]+src=["']([^"']+)["']""", html, re.I):
        u = urllib.parse.urljoin(base_url, m.group(1))
        if u.lower().endswith(".js") and u not in out:
            out.append(u)
    return out


# ---------------------------------------------------------------------------
# 试接口
# ---------------------------------------------------------------------------
def _looks_like_risk(payload, text):
    if isinstance(payload, str):
        return "禁止" in payload or "拦截" in payload or "验证" in payload
    return ("禁止" in text[:400] and "list" not in text)


def _attempt(url, method, headers, body):
    text, ctype = _req(url, method=method, headers=headers, data=body, timeout=15)
    payload = _try_json(text)
    if payload is None:
        return None
    if _looks_like_risk(payload, text):
        raise ProbeError("触发风控")
    path, items = _pick_list(payload)
    if not items:
        return None
    return {"payload": payload, "list_path": path, "items": items, "ctype": ctype}


def _make_hit(payload, text, url, method, body, hdr):
    if _looks_like_risk(payload, text):
        return None
    path, items = _pick_list(payload)
    if not items:
        return None
    return {"payload": payload, "list_path": path, "items": items,
            "method": method, "url": url, "body": body or {},
            "headers": hdr, "page_field": "", "size_field": ""}


POST_BODIES = (
    {"flag": "1", "pageNumber": 1, "pageSize": 15},
    {"pageNumber": 1, "pageSize": 15},
    {"page": 1, "size": 15},
)


def _url_variants(u):
    """很多站点生产环境会加 context-path，页面里写 /api/x，真实地址是 /tzxmspweb/api/x。"""
    out = [u]
    parts = urllib.parse.urlparse(u)
    if "/api/" in parts.path:
        for prefix in ("/tzxmspweb", "/prod-api", "/gateway", "/zcapi"):
            v = urllib.parse.urlunparse(parts._replace(path=prefix + parts.path))
            if v not in out:
                out.append(v)
    return out


def _sniff_one(u, referer):
    hdr = {"Referer": referer, "Origin": "{0.scheme}://{0.netloc}".format(urllib.parse.urlparse(referer))}
    try:
        text, _ = _req(u, headers=dict(hdr, Accept="application/json, */*"), timeout=8)
    except ProbeError:
        text = ""
    payload = _try_json(text) if text else None
    if payload is not None:
        hit = _make_hit(payload, text, u, "GET", None, hdr)
        if hit:
            return hit
    for body in POST_BODIES:
        try:
            text, _ = _req(u, method="POST", headers=dict(hdr), data=body, timeout=8)
        except ProbeError:
            continue
        payload = _try_json(text)
        if payload is None:
            continue
        hit = _make_hit(payload, text, u, "POST", body, hdr)
        if hit:
            return hit
    return None


def _sniff(cand, referer):
    for u in _url_variants(cand):
        hit = _sniff_one(u, referer)
        if hit:
            return hit
    return None


def _diff_page(hit, referer):
    """确认翻页参数：请求第 2 页，看首条记录是否变化。最多试 8 次。"""
    base = {"list_url": hit["url"], "method": hit["method"],
            "headers": hit["headers"], "body": hit.get("body") or {},
            "page_field": "", "size_field": "", "size": 15}
    sig0 = _sig(hit["items"][0])
    tried = 0

    def attempt(pk, sk):
        nonlocal tried
        if tried >= 9:
            return None
        tried += 1
        spec2 = dict(base, page_field=pk, size_field=sk)
        try:
            payload = _request_page(spec2, 2, referer)
        except Exception:  # noqa: BLE001
            return None
        items = _path_get(payload, hit["list_path"]) or []
        if isinstance(items, list) and items and _sig(items[0]) != sig0:
            return spec2
        return None

    # 先固定每页条数、只换页码字段（最省请求）
    for pk in ("pageNum", "pageNumber", "page", "pageNo", "pageIndex", "current", "pn"):
        got = attempt(pk, "pageSize")
        if got:
            return got, True
    # 再试条数字段
    for sk in ("size", "limit", "rows", "pageCount", "count"):
        for pk in ("pageNum", "pageNumber", "page"):
            got = attempt(pk, sk)
            if got:
                return got, True
    return base, False


def _sig(rec):
    if not isinstance(rec, dict):
        return str(rec)[:60]
    for k in ("id", "baId", "projectCode", "code", "name", "projectName", "title"):
        if k in rec:
            return "%s=%s" % (k, rec[k])
    return json.dumps(rec, ensure_ascii=False, sort_keys=True)[:80]


def _request_page(spec, page, referer=""):
    url = spec["list_url"]
    method = spec.get("method", "GET")
    headers = dict(spec.get("headers") or {})
    if referer:
        headers.setdefault("Referer", referer)
    body = dict(spec.get("body") or {})
    pk, sk, size = spec.get("page_field"), spec.get("size_field"), spec.get("size") or 15
    if method == "GET":
        qs = dict(spec.get("query") or {})
        if pk:
            qs[pk] = page
        if sk:
            qs[sk] = size
        text, _ = _req(url + ("&" if "?" in url else "?") + urllib.parse.urlencode(qs),
                       headers=headers, timeout=20)
    else:
        if pk:
            body[pk] = page
        if sk:
            body[sk] = size
        text, _ = _req(url, method="POST", headers=headers, data=body, timeout=20)
    payload = _try_json(text)
    if payload is None:
        raise ProbeError("返回值不是 JSON（可能被风控或需要登录）")
    if _looks_like_risk(payload, text):
        raise ProbeError("触发风控，稍后再试")
    return payload


# ---------------------------------------------------------------------------
# 字段推断
# ---------------------------------------------------------------------------
FIELD_HINTS = [
    ("code", ["projectcode", "pcode", "code", "number", "no", "sn", "regno"]),
    ("name", ["projectname", "proname", "projname", "name", "title", "subject"]),
    ("area", ["areadetailname", "areaname", "area", "admdiv", "region", "city", "district",
              "county", "address", "place", "location"]),
    ("money", ["totalmoney", "totalamount", "amount", "money", "invest", "capital", "price",
               "budget", "fund", "cost", "value", "contract"]),
    ("scale", ["scalecontent", "content", "scale", "desc", "description", "intro", "scope", "remark"]),
    ("dept", ["projectdept", "buildunit", "company", "enterprise", "dept", "unit", "org",
              "applicant", "tenderer", "buyer", "purchaser", "owner", "client", "bidder"]),
    ("gov", ["approveunit", "approveorg", "gov", "office", "authority", "admin",
             "agency", "platform", "center"]),
    ("apply", ["applytime", "applydate", "declaredate", "declare", "apply", "submit"]),
    ("finish", ["finishdate", "passtime", "passtime", "approvetime", "approvedate", "issuedate",
                "publicdate", "publishtime", "auditdate", "enddate", "finish"]),
    ("years", ["starttime", "endtime", "years", "period", "duration"]),
    ("state", ["stateflagname", "state", "status", "flag", "statustext"]),
]


def _sample(rec, key):
    v = rec.get(key)
    if isinstance(v, (dict, list)):
        return ""
    return "" if v is None else str(v)


def _is_date(v):
    s = str(v or "").strip()
    return bool(re.match(r"^(20\d{2})[-/.年]?(\d{1,2})[-/.月]?(\d{1,2})", s))


def _infer_fields(items):
    rec = items[0]
    keys = [k for k in rec.keys() if not isinstance(rec[k], (dict, list))]
    low = {k: k.lower().replace("_", "") for k in keys}

    mapping, used = {}, set()

    def _rank(k):
        """优先取「名称/文本」字段，别抓 xxxId 这种外键。"""
        lk = low[k]
        score = 0
        if lk.endswith("id") or lk.endswith("code") or lk.endswith("uuid"):
            score -= 5
        if lk.endswith("name") or lk.endswith("text") or lk.endswith("title"):
            score += 3
        try:
            float(str(rec[k]).replace(",", ""))
            score += 2          # 数值型（金额很可能是数字）
        except (TypeError, ValueError):
            pass
        if any(t in lk for t in ("money", "amount", "budget", "invest", "fund", "price")):
            score += 2
        return score

    def take(field, cands):
        if field in mapping:
            return
        pool = [k for k in keys if k not in used]

        def pick(match):
            best, best_s = "", -99
            for k in pool:
                if not match(k):
                    continue
                sc = _rank(k)
                if sc > best_s or (sc == best_s and len(low[k]) < len(low[best])):
                    best, best_s = k, sc
            return best

        for c in cands:                      # 完全相等
            k = pick(lambda x, c=c: low[x] == c)
            if k:
                mapping[field] = k
                used.add(k)
                return
        for c in cands:                      # 包含
            k = pick(lambda x, c=c: c in low[x])
            if k:
                mapping[field] = k
                used.add(k)
                return

    # 日期字段先定，避免被 "end"/"start" 之类的词抢走
    date_keys = [k for k in keys if any(t in low[k] for t in DATE_KEYS) and
                 any(_is_date(rec[k]) for _ in [0])]
    date_keys = [k for k in date_keys if _is_date(rec[k])]
    if date_keys:
        finish = next((k for k in date_keys if any(t in low[k] for t in ("finish", "pass", "approv", "issued", "public", "publish"))), date_keys[0])
        mapping["finish"] = finish
        used.add(finish)

    for field in ("code", "name", "area", "money", "scale", "dept", "gov", "apply", "years", "state"):
        if field == "apply":
            cands = ["applytime", "applydate", "declare", "apply", "submit"]
            take(field, cands)
        elif field == "gov":
            take(field, dict(FIELD_HINTS)["gov"])
        else:
            take(field, dict(FIELD_HINTS)[field])

    # 起止年限：两个字段拼起来
    st = next((k for k in keys if any(t in low[k] for t in ("starttime", "startdate", "begindate"))
               and k not in used), "")
    et = next((k for k in keys if any(t in low[k] for t in ("endtime", "enddate", "finishdate"))
               and k not in used and k != st), "")
    if st or et:
        mapping["years"] = [x for x in (st, et) if x]

    # id 字段（去重用）
    id_field = ""
    for k in ("baId", "id", "ID", "projectCode", "code", "uuid"):
        if k in rec:
            id_field = k
            break
    if not id_field:                     # 兜底：xxxId / noticeId / rowKey
        cand = [k for k in keys if low[k].endswith("id") or low[k].endswith("key")]
        id_field = cand[0] if cand else ""

    return mapping, id_field


def _money_text(v):
    try:
        n = float(str(v).replace(",", ""))
    except (TypeError, ValueError):
        return str(v or "")
    if n <= 0:
        return ""
    if abs(n - round(n)) < 0.01:
        return "%d万元" % int(round(n))
    return "%.2f万元" % n


def _clean_area(v):
    s = re.sub(r"\s+", "", str(v or ""))
    m = re.match(r"^(.{2,8}?市)(.*)$", s)
    if not m:
        return s
    city, rest = m.group(1), m.group(2)
    if city in ("东莞市", "中山市", "儋州市", "嘉峪关市"):
        return city
    m2 = re.match(r"^(.{1,5}?(?:自治县|区|县|旗))", rest)
    return city + m2.group(1) if m2 else (city + rest[:3] if rest else city)


def _date_text(v):
    s = re.sub(r"\D", "", str(v or ""))
    if len(s) >= 8:
        return "%s-%s-%s" % (s[:4], s[4:6], s[6:8])
    return str(v or "")[:10]


def ymd(v):
    s = re.sub(r"\D", "", str(v or ""))
    return s[:8] if len(s) >= 8 else ""


# ---------------------------------------------------------------------------
# HTML 列表模式（服务端渲染的站点，列表直接写在页面里）
# ---------------------------------------------------------------------------
TAG_RE = r"<%s[^>]*>(.*?)</%s>"
A_RE = re.compile(r"""<a[^>]+href=["']([^"']+)["'][^>]*>(.*?)</a>""", re.S | re.I)
NEXT_RE = re.compile(r"""<a[^>]+href=["']([^"']+)["'][^>]*>\s*(?:下一页|下页|next|&gt;|>)\s*</a>""", re.I)
DATE_RE = re.compile(r"(20\d{2})\s*[-./年]\s*(\d{1,2})\s*[-./月]\s*(\d{1,2})")


def _text_of(frag):
    t = re.sub(r"(?is)<(script|style)[^>]*>.*?</\1>", " ", frag)
    t = re.sub(r"(?s)<[^>]+>", "", t)
    return re.sub(r"\s+", " ", t).strip()


def _date_in(frag):
    m = DATE_RE.search(frag)
    if not m:
        return ""
    y, mo, d = m.groups()
    try:
        return "%04d-%02d-%02d" % (int(y), int(mo), int(d))
    except ValueError:
        return ""


def parse_html_list(html, base_url):
    """
    从页面里抽「标题 + 链接 + 日期」。
    按 <li>/<tr> 切块，要求块里同时有链接和日期，这样能自动跳过导航、页脚。
    """
    rows, seen = [], set()
    for tag in ("li", "tr", "dd"):
        blocks = re.findall(TAG_RE % (tag, tag), html, re.S)
        if len(blocks) < 8:
            continue
        got = []
        for b in blocks:
            m = A_RE.search(b)
            if not m:
                continue
            href, raw = m.group(1).strip(), m.group(2)
            title = _text_of(raw)
            if not title or not (6 <= len(title) <= 160):
                continue
            if href.lower().startswith(("javascript:", "#", "mailto:")):
                continue
            d = _date_in(b)
            if not d:
                continue
            full = urllib.parse.urljoin(base_url, href)
            if full in seen:
                continue
            seen.add(full)
            got.append({"name": title, "url": full, "finish": d})
        if len(got) > len(rows):
            rows = got
    return rows


def _html_next_tpl(html, base_url, page_url):
    """找「下一页」链接，推导页码模板。"""
    for m in NEXT_RE.finditer(html):
        href = urllib.parse.urljoin(base_url, m.group(1))
        mm = re.match(r"^(.*?)(\d+)(\.s?html?)$", href)
        if mm and int(mm.group(2)) <= 3:
            return mm.group(1) + "{n}" + mm.group(3)
    mm = re.match(r"^(.*?)(\d*)(\.s?html?)$", page_url)
    if mm and mm.group(2) == "":
        base = page_url[: -len(mm.group(3))].rstrip("_")
        return base + "_{n}" + mm.group(3)
    return ""


def render_dom(url, wait_ms=6000):
    """用系统浏览器（Edge/Chrome）真实渲染页面，等 JS 加载完再取 DOM。类似 RPA。"""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        raise ProbeError("未安装 playwright（pip install playwright）")

    errs = []
    for ch in ("msedge", "chrome", None):
        try:
            with sync_playwright() as p:
                kw = {"headless": True,
                      "args": ["--no-first-run", "--no-default-browser-check", "--disable-gpu"]}
                if ch:
                    kw["channel"] = ch
                browser = p.chromium.launch(**kw)
                page = browser.new_page(user_agent=UA)
                page.goto(url, timeout=30000, wait_until="domcontentloaded")
                try:
                    page.wait_for_load_state("networkidle", timeout=wait_ms)
                except Exception:  # noqa: BLE001
                    page.wait_for_timeout(2000)
                html = page.content()
                browser.close()
                return html
        except Exception as e:  # noqa: BLE001
            errs.append("%s: %s" % (ch or "chromium", str(e)[:80]))
            continue
    raise ProbeError("浏览器渲染失败：" + " | ".join(errs)[:220])


def probe_html_list(url, hint="", html=None, mode="html"):
    """页面解析：从静态 HTML 或浏览器渲染后的 DOM 里抽「标题+链接+日期」列表。"""
    if html is None:
        html, title = _html(url)
    else:
        title = ""
        m = re.search(r"<title[^>]*>(.*?)</title>", html, re.S | re.I)
        if m:
            title = _text_of(m.group(1))
    rows = parse_html_list(html, url)
    if len(rows) < 8:
        raise ProbeError("页面里没找到成规模的列表（少于 8 条），可能登录后可见")
    tpl = _html_next_tpl(html, url, url)
    netloc = urllib.parse.urlparse(url).netloc
    spec = {
        "mode": mode,
        "key": "auto_%s_%s" % (mode, re.sub(r"[^a-z0-9]+", "_", netloc.lower()).strip("_")),
        "name": title or netloc,
        "host": netloc,
        "page_url": url,
        "list_url": url,
        "method": "GET",
        "list_path": "",
        "page_url_tpl": tpl,
        "field_map": {"name": "__title", "finish": "__date", "url": "__url"},
        "date_field": "finish",
        "id_field": "url",
        "hint": hint or "",
        "confidence": (0.6 if not tpl else 0.75) if mode == "html" else (0.7 if not tpl else 0.8),
        "created": time.strftime("%Y-%m-%d %H:%M"),
        "notes": ["浏览器渲染 + 页面解析（无官方接口）"] if mode == "rpa"
                 else ["按页面 HTML 结构解析（无 JSON 接口）"],
    }
    logs = [
        ("浏览器渲染成功，DOM %d 字" % len(html)) if mode == "rpa"
        else ("打开页面成功（%d 字，标题：%s）" % (len(html), title or "无")),
        "页面里识出 %d 条「标题 + 链接 + 日期」" % len(rows),
        ("翻页模板：" + tpl) if tpl else "没找到「下一页」链接，只能抓第一页",
    ]
    preview = [dict(x, code="", area="", money="", dept="", scale="", state="", summary="")
               for x in rows[:5]]
    return {"ok": True, "spec": spec, "preview": rows[:3], "preview_rows": preview,
            "logs": logs, "warnings": [] if tpl else ["未识别到翻页，只能抓第一页"],
            "title": title}


def _html_page(spec, page):
    if page == 1 or not spec.get("page_url_tpl"):
        url = spec["list_url"]
    else:
        url = spec["page_url_tpl"].format(n=page)
    if spec.get("mode") == "rpa":
        text = render_dom(url)
    else:
        text, _ = _req(url, headers={"Accept": "text/html,*/*",
                                     "Referer": spec.get("page_url") or url}, timeout=20)
    return parse_html_list(text, url)


# ---------------------------------------------------------------------------
# 探测主流程
# ---------------------------------------------------------------------------
def probe(url, hint="", referer_hint="", max_candidates=None):
    """
    输入网址 → 输出 adapter spec + 预览数据。
    返回 {"ok":True,"spec":...,"preview":[...],"logs":[...],"warnings":[...]}
    """
    logs, warnings = [], []
    if not re.match(r"^https?://", url or ""):
        url = "http://" + (url or "")
    base = urllib.parse.urlparse(url)
    host = base.netloc

    html, title = _html(url)
    logs.append("打开页面成功（%d 字，标题：%s）" % (len(html), title or "无"))

    js = _js_files(html, url)
    logs.append("发现 %d 个脚本文件" % len(js))

    cands = _api_candidates(html, url, js)
    logs.append("提取到 %d 个候选接口" % len(cands))

    hit = None
    hits = []
    pool = [c for c in (cands if max_candidates is None else cands[:max_candidates]) if c != url][:9]
    deadline = time.time() + 55
    for i in range(0, len(pool), 3):
        if time.time() > deadline:
            warnings.append("探测超时，只试了前 %d 个候选接口" % i)
            break
        batch = pool[i:i + 3]
        with ThreadPoolExecutor(max_workers=3) as ex:
            futs = [ex.submit(_sniff, c, url) for c in batch]
            for fu in as_completed(futs):
                try:
                    r = fu.result()
                except Exception:  # noqa: BLE001
                    r = None
                if r:
                    hits.append(r)

    # 多个接口都能出列表时，选「字段最全」的那个（列表接口通常字段最多）
    best_score = -1
    for r in hits:
        fm, _id = _infer_fields(r["items"])
        score = len(fm) * 10 + min(len(r["items"]), 30)
        r["_score"] = score
        r["_fields"] = fm
        if score > best_score:
            best_score, hit = score, r
    if hit:
        logs.append("接口命中：%s（%s，%d 条记录 / %d 个字段）"
                    % (hit["url"].split("?")[0], hit["method"],
                       len(hit["items"]), len(hit["_fields"])))
        if len(hits) > 1:
            logs.append("另有 %d 个接口也能出列表，已选字段最全的一个" % (len(hits) - 1))

    if not hit:
        # 兜底 1：服务端渲染，列表直接写在 HTML 里
        try:
            res = probe_html_list(url, hint)
            res["logs"] = logs + ["没找到 JSON 接口，改用页面列表解析"] + res["logs"]
            return res
        except ProbeError as e_html:
            # 兜底 2：JS 动态渲染，用浏览器（RPA）加载后再解析
            try:
                dom = render_dom(url)
            except ProbeError as e_rpa:
                raise ProbeError("没找到可抓的列表（静态页面和浏览器渲染都失败）：%s；%s。"
                                 "可以换「列表页」网址，或把已知接口地址填进来。" % (e_html, e_rpa))
            res = probe_html_list(url, hint, html=dom, mode="rpa")
            res["logs"] = logs + ["静态 HTML 没有列表，改用浏览器渲染（RPA）"] + res["logs"]
            return res

    items = hit["items"]
    logs.append("第一页拿到 %d 条记录" % len(items))

    spec, paging_ok = _diff_page(hit, url)
    if paging_ok:
        logs.append("翻页参数确认：%s / %s" % (spec["page_field"], spec["size_field"]))
    else:
        warnings.append("翻页参数没能确认（可能站点只返回第一页，或每页条数被服务端固定）")

    mapping, id_field = _infer_fields(items)
    logs.append("识别字段 %d 个：%s" % (len(mapping), "、".join(sorted(mapping))))

    # 总数
    total = {}
    payload = hit["payload"]
    node = _path_get(payload, spec.get("list_path", "").rsplit(".", 1)[0]) if spec.get("list_path") else payload
    total = _numbers(node if isinstance(node, dict) else payload)

    spec.update({
        "key": "auto_%s" % re.sub(r"[^a-z0-9]+", "_", host.lower()).strip("_"),
        "name": title or host,
        "host": host,
        "page_url": url,
        "hint": hint or "",
        "list_path": hit["list_path"],
        "field_map": mapping,
        "id_field": id_field,
        "date_field": mapping.get("finish") or "",
        "total": total,
        "confidence": round(min(0.95, 0.45 + 0.05 * len(mapping) + (0.15 if paging_ok else 0)), 2),
        "created": time.strftime("%Y-%m-%d %H:%M"),
        "notes": logs[-2:],
    })
    if not mapping.get("finish"):
        warnings.append("没识别到「通过/发布日期」字段，按日期过滤会退化为不过滤，请在 spec 里手工指定 date_field")

    preview = [_row_of(x, spec, 50000) for x in items[:5] if isinstance(x, dict)]
    return {"ok": True, "spec": spec, "preview": items[:3], "preview_rows": preview,
            "logs": logs, "warnings": warnings, "title": title}


# ---------------------------------------------------------------------------
# 通用执行
# ---------------------------------------------------------------------------
HTML_ALIAS = {"__title": "name", "__date": "finish", "__url": "url"}


def _get_val(rec, spec, field):
    k = (spec.get("field_map") or {}).get(field)
    if not k:
        return ""
    if isinstance(k, str) and k in HTML_ALIAS:
        return rec.get(HTML_ALIAS[k])
    if isinstance(k, list):
        vals = [str(rec.get(x) or "").strip() for x in k]
        return " ~ ".join([v for v in vals if v])
    return rec.get(k)


def _row_of(rec, spec, min_wan):
    area = _clean_area(_get_val(rec, spec, "area"))
    money_raw = _get_val(rec, spec, "money")
    money = _money_text(money_raw)
    try:
        wan = float(str(money_raw).replace(",", "")) if str(money_raw).strip() else 0
    except (TypeError, ValueError):
        wan = 0
    name = str(_get_val(rec, spec, "name") or "").strip()
    dept = str(_get_val(rec, spec, "dept") or "").strip()
    scale = re.sub(r"\s+", "", str(_get_val(rec, spec, "scale") or "")).rstrip("。；;")
    summary = ""
    if wan >= min_wan and (name or dept):
        yi = ("%.2f" % (wan / 10000.0)).rstrip("0").rstrip(".")
        head = ("%s拟建设%s" % (dept, name)) if dept else ("拟建设%s" % name)
        parts = [head]
        if scale:
            parts.append(scale)
        if area:
            parts.append("项目地位于%s" % area)
        parts.append("总投资%s亿元" % yi)
        summary = "，".join(parts) + "。"
    years = str(_get_val(rec, spec, "years") or "")
    return {
        "code": str(_get_val(rec, spec, "code") or ""),
        "name": name,
        "area": area,
        "money": money,
        "scale": str(_get_val(rec, spec, "scale") or "").strip(),
        "dept": dept,
        "gov": str(_get_val(rec, spec, "gov") or ""),
        "apply": _date_text(_get_val(rec, spec, "apply")),
        "finish": _date_text(_get_val(rec, spec, "finish")),
        "years": years,
        "state": str(_get_val(rec, spec, "state") or ""),
        "summary": summary,
        "_wan": wan,
    }


def run_spec(spec, start, end, on_progress=None, max_pages=400, sleep_range=(0.8, 1.6),
             min_wan=50000):
    """按 spec 抓取。start/end 为 YYYYMMDD。返回 (dict 行列表, stats)。"""
    rows, seen = [], set()
    stats = {"pages": 0, "fetched": 0, "kept": 0, "skipped_before": 0}
    list_path = spec.get("list_path", "")
    date_field = spec.get("date_field") or ""
    page, stop = 1, False
    if not spec.get("page_field") and spec.get("mode") not in ("html", "rpa"):
        max_pages = 1        # 翻页参数没确认，只抓第一页，避免重复拉同一页
    if spec.get("mode") in ("html", "rpa") and not spec.get("page_url_tpl"):
        max_pages = 1        # 没找到「下一页」，只能抓第一页

    html_mode = spec.get("mode") in ("html", "rpa")
    while not stop and page <= max_pages:
        if html_mode:
            try:
                items = _html_page(spec, page)
            except ProbeError as e:
                if page == 1:
                    raise
                stats["error"] = str(e)
                break
            if not items:
                break
        else:
            try:
                payload = _request_page(spec, page, spec.get("page_url", ""))
            except ProbeError as e:
                if page == 1:
                    raise
                stats["error"] = str(e)
                break
            items = _path_get(payload, list_path) if list_path else payload
            if not isinstance(items, list) or not items:
                break

        for rec in items:
            if not isinstance(rec, dict):
                continue
            stats["fetched"] += 1
            d = ymd(rec.get(date_field)) if date_field else ""
            if d and d < start:
                stats["skipped_before"] += 1
                stop = True
                continue
            if date_field and d and not (start <= d <= end):
                continue
            if not date_field and page == 1 and not rows:
                pass  # 无日期字段：不过滤，靠 max_pages 控制
            key = str(rec.get(spec.get("id_field") or "", "")) or _sig(rec)
            if key in seen:
                continue
            seen.add(key)
            rows.append(_row_of(rec, spec, min_wan))
            stats["kept"] += 1

        stats["pages"] = page
        if on_progress:
            on_progress(page, stats)
        if stop:
            break
        page += 1
        time.sleep(random.uniform(*sleep_range))

    rows.sort(key=lambda r: r.get("finish") or "", reverse=True)
    return rows, stats


def test_spec(spec, pages=1):
    """只抓第一页，给页面做「试抓」。"""
    if spec.get("mode") in ("html", "rpa"):
        try:
            items = _html_page(spec, 1)
        except ProbeError as e:
            return {"ok": False, "error": str(e)}
        rows = [_row_of(x, spec, 50000) for x in items][:5]
        return {"ok": True, "count": len(items), "rows": rows, "total": {}}

    try:
        payload = _request_page(spec, 1, spec.get("page_url", ""))
    except ProbeError as e:
        return {"ok": False, "error": str(e)}
    items = _path_get(payload, spec.get("list_path", "")) if spec.get("list_path") else payload
    if not isinstance(items, list):
        return {"ok": False, "error": "按 list_path=%s 取不到列表，请检查路径" % spec.get("list_path")}
    rows = [_row_of(x, spec, 50000) for x in items if isinstance(x, dict)][:5]
    total = _numbers(payload)
    return {"ok": True, "count": len(items), "rows": rows, "total": total}


def normalize_spec(raw):
    """把用户编辑过的 spec 收拾干净。"""
    if isinstance(raw, str):
        raw = json.loads(raw)
    spec = dict(raw or {})
    spec["list_url"] = (spec.get("list_url") or "").strip()
    if not spec["list_url"]:
        raise ValueError("list_url 不能为空")
    spec["method"] = (spec.get("method") or "GET").upper()
    spec["list_path"] = (spec.get("list_path") or "").strip()
    spec["size"] = int(spec.get("size") or 15)
    spec["mode"] = spec.get("mode") or "json"
    spec["page_url_tpl"] = spec.get("page_url_tpl") or ""
    spec.setdefault("headers", {})
    spec.setdefault("body", {})
    fm = spec.get("field_map") or {}
    spec["field_map"] = {k: v for k, v in fm.items() if v}
    for k, v in list(spec["field_map"].items()):
        if isinstance(v, list):
            spec["field_map"][k] = [x for x in v if x]
    if spec["method"] == "GET" and spec.get("body"):
        spec["query"] = spec.pop("body")
    return spec
