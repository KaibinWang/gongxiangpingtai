# -*- coding: utf-8 -*-
"""
投行 & 科技金融 数智化工作台（Flask 多模块可运行版）
====================================================
模块：
  1. 首页看板      —— 商机/匹配/素材库汇总
  2. 商机管理      —— 商机台账 CRUD，多人共享（局域网）
  3. 跨表智能匹配  —— 替代手工 VLOOKUP：上传两张表，选关键列，精确+模糊匹配，导出结果
  4. 素材库        —— 扫描本地材料文件夹，自动解析 DOCX 标题结构（模板骨架）
  5. 初稿生成      —— 选一份历史材料做结构参照 + 商机数据概况，生成供修改的初稿（Markdown/文本）

技术栈：Flask + SQLite + openpyxl，单文件，无前端构建依赖。
运行：  python app.py  →  浏览器访问 http://127.0.0.1:5000
      （团队共享：同网段成员访问 http://本机IP:5000）
"""
import os
import re
import io
import csv
import uuid
import json
import time
import threading
import sqlite3
import zipfile
import difflib
import datetime
import urllib.parse
import random
import string
import re as _re_mail
import smtplib
from email.mime.text import MIMEText
from email.header import Header
from io import BytesIO

from flask import Flask, request, render_template_string, redirect, url_for, send_file, jsonify, session
from werkzeug.utils import secure_filename
from werkzeug.security import generate_password_hash, check_password_hash
import secrets
import openpyxl
from docx import Document
from docx.shared import Pt, RGBColor
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml.ns import qn

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# 数据库路径；可用环境变量 WB_DB_PATH 指向副本，供自动化验证使用，避免误伤生产数据
DB_PATH = os.environ.get("WB_DB_PATH") or os.path.join(BASE_DIR, "workbench.db")
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
RESULT_DIR = os.path.join(BASE_DIR, "results")
for d in (UPLOAD_DIR, RESULT_DIR):
    os.makedirs(d, exist_ok=True)

DEFAULT_MATERIAL_ROOT = r"D:\科技金融" if os.path.isdir(r"D:\科技金融") else ""

STAGES = ["初步接洽", "尽职调查", "方案提交", "商务谈判", "签约交割", "投后管理"]
STATUSES = ["进行中", "已成交", "已暂停", "已流失"]
EXP_STATUS = ["未接触", "已接触", "已开户", "已授信", "已投放"]
GOV_STAGES = ["信息收集", "已对接", "方案洽谈", "已储备", "已落地"]

# 商机类型（导入模板与台账共用）
DEAL_TYPES = ["资本市场商机", "增存商机", "并购商机", "非房固贷商机"]
DEAL_SUBTYPES = {
    "资本市场商机": ["回购增持", "可转债", "定向增发", "IPO", "其他"],
    "增存商机": ["一级市场股权融资款", "募集资金归集", "其他"],
    "并购商机": ["发行股份购买资产", "现金收购", "其他"],
    "非房固贷商机": ["发改委备案项目", "省市重点项目", "其他"],
}
DEAL_TYPE_PILL = {"资本市场商机": "p1", "增存商机": "p2", "并购商机": "p5", "非房固贷商机": "p3"}
DEAL_TYPE_KEY = {"资本市场商机": "cap", "增存商机": "deposit", "并购商机": "ma", "非房固贷商机": "loan"}
DEAL_KEY_TYPE = {v: k for k, v in DEAL_TYPE_KEY.items()}

# 导入模板列定义：(Excel 表头, 数据库字段, 是否必填, 说明)
IMPORT_COLUMNS = [
    ("商机类型", "deal_type", True, "四选一：资本市场商机 / 增存商机 / 并购商机 / 非房固贷商机"),
    ("细分类型", "sub_type", False, "回购增持、一级市场股权融资款、现金收购、发改委备案项目 等"),
    ("商机名称", "name", True, "一句话说清做什么，如「某公司科创债承销」"),
    ("客户/主体", "client", True, "上市公司 / 融资企业 / 并购竞买方 / 项目建设单位"),
    ("交易对手/标的", "counterparty", False, "并购标的方；其他类型可留空"),
    ("所属行业", "sector", False, "半导体 / 生物医药 / 高端制造 等"),
    ("金额（万元）", "amount", False, "单笔金额；区间则填下限"),
    ("金额上限（万元）", "amount2", False, "区间上限，如回购 1-2 亿元：金额填 10000、上限填 20000"),
    ("交易/业务方式", "method", False, "回购、发行股份+现金、B 轮融资、固定资产贷款 等"),
    ("股权比例(%)", "ratio", False, "并购交易标的股权比例，填数字，如 55"),
    ("推进阶段", "stage", False, "默认「初步接洽」，可选：初步接洽/尽职调查/方案提交/商务谈判/签约交割/投后管理"),
    ("状态", "status", False, "默认「进行中」，可选：进行中/已成交/已暂停/已流失"),
    ("负责人", "owner", False, "本行客户经理"),
    ("共享成员", "teammate", False, "多人用逗号分隔"),
    ("主办机构", "org", False, "如「东莞分行」「天河智慧城支行」"),
    ("信息来源", "source", False, "分行报送 / 上市公司公告 / 发改委备案 / 园区推送"),
    ("公告/备案日期", "notice_date", False, "格式 2026-10-19"),
    ("所在地", "location", False, "项目或企业所在区/市"),
    ("商机日期", "deal_date", False, "商机登记日期，留空按导入当天"),
    ("营销切入点", "entry", False, "下一步动作，如「以科技金融特色产品切入股权融资款」"),
    ("备注", "note", False, "其他补充说明"),
]
IMPORT_HEADERS = [c[0] for c in IMPORT_COLUMNS]
DEAL_EXTRA_FIELDS = ["deal_type", "sub_type", "counterparty", "method", "amount2",
                     "ratio", "location", "org", "source", "notice_date", "entry", "deal_date"]

app = Flask(__name__)
app.secret_key = "local-workbench"          # 启动时会被 init_db() 生成的随机密钥替换
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024
app.config["PERMANENT_SESSION_LIFETIME"] = datetime.timedelta(days=7)
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"

# 默认账号（首次启动自动创建，登录后强制改密码）
DEFAULT_USER = ("admin", "th2026", "管理员")

# 固贷商机：经营单位 → 对接人（产品经理）分工表
GUDAI_DIVISION = {
    "东莞分行": "叶琴", "珠海分行": "胡俊霖", "佛山分行": "欧丽娜", "江门分行": "叶琴",
    "惠州分行": "李典睛", "中山分行": "欧丽娜", "肇庆分行": "胡俊霖", "茂名分行": "欧丽娜",
    "阳江分行": "胡俊霖", "汕头分行": "胡俊霖", "湛江分行": "李典睛",
    "北京路支行": "欧丽娜", "分行营业部": "叶琴", "花园支行": "李典睛", "海珠支行": "胡俊霖",
    "合创中心支行": "胡俊霖", "开发区支行": "叶琴", "科技园支行": "胡俊霖", "越秀支行": "胡俊霖",
    "荔湾支行": "欧丽娜", "南沙分行": "叶琴", "天河支行": "欧丽娜", "横琴分行": "胡俊霖",
    "白云新城支行": "李典睛", "国际大厦支行": "欧丽娜", "环市支行": "欧丽娜", "公司事业部": "叶琴",
    "番禺支行": "李典睛", "珠江新城支行": "叶琴", "白云支行": "李典睛", "高德支行": "胡俊霖",
    "黄埔支行": "胡俊霖", "琶洲支行": "欧丽娜", "滨江东支行": "叶琴", "新城国际支行": "叶琴",
    "花城支行": "叶琴", "宝岗大道支行": "欧丽娜", "国防大厦支行": "李典睛", "北秀支行": "李典睛",
    "增城支行": "胡俊霖", "市桥支行": "李典睛", "花都支行": "李典睛", "万博支行": "欧丽娜",
    "凤凰城支行": "叶琴", "汉溪支行": "叶琴", "东风支行": "胡俊霖", "科教城支行": "李典睛",
    "天河智慧城支行": "胡俊霖", "华润天合支行": "欧丽娜", "开元支行": "欧丽娜",
}

# S70 贷款模版：表头（中文）↔ 数据库列名（50 列；「贷后管理人员行员号」在源表出现两次，
# 第二列改名为 贷后管理人员行员号2）
S70_HEADERS = [
    "数据日期", "借据号", "客户号", "客户名称", "贷后管理人员行员号", "贷后管理人员",
    "经营单位", "二级支行", "授信品种", "折亿元", "借据余额折人民币", "借据发放日期",
    "借据到期日期", "会计产品名称", "核心经办机构编号", "一级分行机构号", "一级分行机构名称",
    "贷后管理经办机构名称", "科目号", "产品编号", "币种代码", "信贷国民经济行业分类代码",
    "投向行业代码", "企业客户类型代码", "企业规模代码", "担保方式代码", "计财担保方式代码",
    "借款企业统一社会信用代码", "借据状态代码", "业务品种代码", "业务品种名称",
    "会计核心产品代码", "绿色信贷标识", "相关批复流水号", "批复对应主体客户号",
    "批复对应主体名称", "高新技术企业", "科技型中小企业", "专精特新中小企业",
    "专精特新小巨人企业", "国家技术创新示范企业", "制造业单项冠军企业", "国家企业技术中心",
    "创新型中小企业", "贷后管理一级分行", "贷后管理人员行员号2", "穿透前客户名称",
    "五级分类代码", "借据金额折人民币", "执行利率值"]
S70_COLS = [
    "data_date", "loan_no", "cust_no", "cust_name", "mgr_id", "mgr_name", "unit",
    "sub_branch", "credit_type", "amount_yi", "balance_rmb", "loan_start", "loan_end",
    "acct_prod", "core_org_no", "branch1_no", "branch1_name", "post_mgr_org",
    "subject_no", "prod_no", "currency", "industry_credit", "industry_target",
    "corp_type", "corp_scale", "guaranty", "guaranty_fin", "credit_code", "loan_status",
    "biz_no", "biz_name", "acct_core", "green", "approval_no", "appr_cust_no",
    "appr_cust_name", "high_tech", "tech_sme", "zjtx_sme", "zjtx_giant", "tech_demo",
    "champion", "tech_center", "innov_sme", "post_branch1", "mgr_id2",
    "pre_cust_name", "five_class", "loan_amt_rmb", "rate"]
S70_REAL_COLS = {"amount_yi", "balance_rmb", "loan_amt_rmb", "rate"}

# 列表/详情展示时脱敏的敏感字段（借据号、客户号、行员号、批复流水号、会计产品名称与代码）
S70_MASK_COLS = {"loan_no", "cust_no", "mgr_id", "mgr_id2",
                 "approval_no", "acct_prod", "acct_core"}


def s70_mask(v):
    """敏感字段打码：≥9 位保留前 4 后 4，5-8 位保留前 2 后 2，≤4 位全掩"""
    s = "" if v is None else str(v).strip()
    if not s:
        return ""
    n = len(s)
    if n <= 2:
        return "*" * n
    if n <= 8:
        keep = max(1, n // 4)
        return s[:keep] + "*" * (n - keep * 2) + s[-keep:]
    return s[:4] + "*" * (n - 8) + s[-4:]

# S70 列表默认隐藏的列（目前不隐藏任何列；如需隐藏，把列名加进此集合即可，
# 编辑、详情、导出始终保留全部字段）
S70_HIDDEN_COLS = set()

# 新增/编辑表单分组（全部 50 列都可填）
S70_FORM_GROUPS = [
    ("基本信息", ["data_date", "loan_no", "cust_no", "cust_name", "mgr_id", "mgr_name",
                  "unit", "sub_branch", "credit_type", "amount_yi", "balance_rmb",
                  "loan_start", "loan_end", "acct_prod"]),
    ("机构与科目", ["core_org_no", "branch1_no", "branch1_name", "post_mgr_org",
                    "subject_no", "prod_no", "currency"]),
    ("行业分类与企业属性", ["industry_credit", "industry_target", "corp_type", "corp_scale",
                            "guaranty", "guaranty_fin", "credit_code", "loan_status"]),
    ("业务品种与批复", ["biz_no", "biz_name", "acct_core", "green", "approval_no",
                        "appr_cust_no", "appr_cust_name"]),
    ("企业资质标识", ["high_tech", "tech_sme", "zjtx_sme", "zjtx_giant", "tech_demo",
                      "champion", "tech_center", "innov_sme"]),
    ("其他", ["post_branch1", "mgr_id2", "pre_cust_name", "five_class",
              "loan_amt_rmb", "rate"]),
]
_S70_GROUP_COLS = [c for _, cols in S70_FORM_GROUPS for c in cols]
assert sorted(_S70_GROUP_COLS) == sorted(S70_COLS), "S70 表单分组必须覆盖全部 50 列"


def s70_num(v):
    """数值显示：整数不带 .0 / 科学计数法，小数去掉浮点尾差（0.30000000000000004 → 0.3）"""
    if v is None or v == "":
        return ""
    try:
        f = float(v)
    except (TypeError, ValueError):
        return str(v)
    if f.is_integer():
        return str(int(f))
    return ("%.10f" % f).rstrip("0").rstrip(".")


def cn_date_to_iso(s):
    """2026年10月31日 -> 2026-10-31（给 <input type=date> 用）"""
    m = re.match(r"\s*(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日", str(s or ""))
    return "%s-%02d-%02d" % (m.group(1), int(m.group(2)), int(m.group(3))) if m else ""


def iso_date_to_cn(s):
    """2026-10-31 -> 2026年10月31日（与库中 data_date 不补零格式一致）"""
    m = re.match(r"\s*(\d{4})-(\d{1,2})-(\d{1,2})", str(s or ""))
    return "%d年%d月%d日" % (int(m.group(1)), int(m.group(2)), int(m.group(3))) if m else ""


# S70 导入时需归一化的日期列（data_date 存中文格式，借据起止日期存 ISO）
S70_DATE_COLS = {"data_date", "loan_start", "loan_end"}
XL_EPOCH = datetime.datetime(1899, 12, 30)   # Excel 1900 日期系统基准


def norm_date(v, cn=True):
    """把日期归一化：支持 datetime / Excel 序列号(46326) / 中文 / ISO / 纯数字字符串。
    无法识别时原样返回；无效日期（如 9月31日）也原样保留，不擅自改数。"""
    if v is None or v == "":
        return ""
    d = None
    if isinstance(v, datetime.datetime):
        d = v.date()
    elif isinstance(v, datetime.date):
        d = v
    elif isinstance(v, (int, float)):
        n = int(v)
        if 20000 <= n <= 80000:            # 1954-08-01 ~ 2119-01-05 区间视为 Excel 序列号
            d = (XL_EPOCH + datetime.timedelta(days=n)).date()
    else:
        s = str(v).strip()
        m = re.match(r"\s*(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日?", s)
        if m:
            try:
                d = datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            except ValueError:
                d = None
        else:
            m = re.match(r"\s*(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})", s)
            if m:
                try:
                    d = datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
                except ValueError:
                    d = None
            elif s.isdigit() and 20000 <= int(s) <= 80000:
                d = (XL_EPOCH + datetime.timedelta(days=int(s))).date()
    if d is None:
        return str(v).strip()
    return ("%d年%d月%d日" % (d.year, d.month, d.day)) if cn else d.strftime("%Y-%m-%d")


# ----------------------------------------------------------------------------
# 数据库
# ----------------------------------------------------------------------------
def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_db()
    conn.executescript("""
    CREATE TABLE IF NOT EXISTS deals (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL, client TEXT, sector TEXT,
        stage TEXT DEFAULT '初步接洽', amount REAL DEFAULT 0,
        owner TEXT, teammate TEXT, status TEXT DEFAULT '进行中',
        note TEXT, created TEXT, updated TEXT);
    CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT);
    CREATE TABLE IF NOT EXISTS materials (
        id INTEGER PRIMARY KEY AUTOINCREMENT, path TEXT UNIQUE,
        name TEXT, rel TEXT, kind TEXT, size REAL, mtime TEXT, cat TEXT);
    CREATE TABLE IF NOT EXISTS match_jobs (
        id INTEGER PRIMARY KEY AUTOINCREMENT, ts TEXT,
        file_a TEXT, file_b TEXT, headers_a TEXT, headers_b TEXT);
    CREATE TABLE IF NOT EXISTS match_runs (
        id INTEGER PRIMARY KEY AUTOINCREMENT, ts TEXT, file_a TEXT, file_b TEXT,
        col_a TEXT, col_b TEXT, total INTEGER, exact INTEGER, fuzzy INTEGER,
        unmatched INTEGER, result_path TEXT);
    CREATE TABLE IF NOT EXISTS expansion_clients (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL, sector TEXT, region TEXT, tag TEXT,
        source TEXT, status TEXT DEFAULT '未接触', owner TEXT,
        note TEXT, created TEXT, updated TEXT);
    CREATE TABLE IF NOT EXISTS gov_projects (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL, region TEXT, sector TEXT,
        invest REAL DEFAULT 0, demand REAL DEFAULT 0,
        stage TEXT DEFAULT '信息收集', source TEXT,
        owner TEXT, note TEXT, created TEXT, updated TEXT);
    CREATE TABLE IF NOT EXISTS gov_subs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        keyword TEXT NOT NULL, typ TEXT DEFAULT '行业', created TEXT);
    CREATE TABLE IF NOT EXISTS scan_logs (
        id INTEGER PRIMARY KEY AUTOINCREMENT, ts TEXT,
        trigger TEXT, added INTEGER, updated INTEGER, removed INTEGER, total INTEGER);
    CREATE TABLE IF NOT EXISTS pipeline_items (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        category TEXT NOT NULL DEFAULT 'bond',
        push_date TEXT, unit TEXT, name TEXT NOT NULL,
        amount REAL DEFAULT 0, content TEXT, manager TEXT, note TEXT,
        created TEXT, updated TEXT);
    CREATE TABLE IF NOT EXISTS roles (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE NOT NULL,
        menus TEXT DEFAULT '[]',
        created TEXT);
    CREATE TABLE IF NOT EXISTS gudai_deals (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        push_date TEXT, unit TEXT, name TEXT, region TEXT,
        amount REAL DEFAULT 0, content TEXT, manager TEXT,
        note TEXT, approval_amount REAL,
        created TEXT, updated TEXT);
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        pwd_hash TEXT NOT NULL,
        name TEXT,
        email TEXT,
        role TEXT DEFAULT 'user',
        must_change INTEGER DEFAULT 0,
        created TEXT, last_login TEXT);
    """)
    # 兼容旧库：users 补 email 列
    _ucols = [r[1] for r in conn.execute("PRAGMA table_info(users)").fetchall()]
    if "email" not in _ucols:
        conn.execute("ALTER TABLE users ADD COLUMN email TEXT")
    # 初始化角色（RBAC）：管理员=全部菜单，其余按业务线绑定
    default_roles = {
        "管理员": "ALL",
        "科技金融": ["home", "tech"],
        "综合融资": ["home", "gov"],
        "债券金融": ["home", "bond"],
        "资本市场": ["home", "capmarket"],
        "游客": ["home"],
    }
    for rname, rmenus in default_roles.items():
        if not conn.execute("SELECT 1 FROM roles WHERE name=?", (rname,)).fetchone():
            conn.execute("INSERT INTO roles (name, menus, created) VALUES (?,?,?)",
                         (rname, json.dumps(rmenus),
                          datetime.datetime.now().strftime("%Y-%m-%d %H:%M")))
    if not conn.execute("SELECT 1 FROM settings WHERE key='material_root'").fetchone():
        conn.execute("INSERT INTO settings VALUES ('material_root', ?)", (DEFAULT_MATERIAL_ROOT,))
    if not conn.execute("SELECT 1 FROM settings WHERE key='auto_scan_on'").fetchone():
        conn.execute("INSERT INTO settings VALUES ('auto_scan_on', '1')")
    if not conn.execute("SELECT 1 FROM settings WHERE key='auto_scan_time'").fetchone():
        conn.execute("INSERT INTO settings VALUES ('auto_scan_time', '08:30')")
    if not conn.execute("SELECT 1 FROM settings WHERE key='secret_key'").fetchone():
        conn.execute("INSERT INTO settings VALUES ('secret_key', ?)", (secrets.token_hex(32),))
    if not conn.execute("SELECT 1 FROM settings WHERE key='gudai_division'").fetchone():
        conn.execute("INSERT INTO settings VALUES ('gudai_division', ?)",
                     (json.dumps(GUDAI_DIVISION, ensure_ascii=False),))
    # S70 贷款表（按模版 50 列动态建表）
    cols_sql = ", ".join(
        "%s %s" % (c, "REAL" if c in S70_REAL_COLS else "TEXT") for c in S70_COLS)
    conn.execute("CREATE TABLE IF NOT EXISTS s70_loans "
                 "(id INTEGER PRIMARY KEY AUTOINCREMENT, %s, created TEXT)" % cols_sql)

    # 首次启动创建默认管理员，并要求登录后修改密码
    if not conn.execute("SELECT 1 FROM users LIMIT 1").fetchone():
        u, pw, nm = DEFAULT_USER
        conn.execute(
            "INSERT INTO users (username, pwd_hash, name, role, must_change, created) "
            "VALUES (?,?,?,?,?,?)",
            (u, generate_password_hash(pw), nm, "admin", 1,
             datetime.datetime.now().strftime("%Y-%m-%d %H:%M")))
    conn.commit()
    conn.close()


def esc_html(x):
    return str(x if x is not None else "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def get_setting(key, default=""):
    conn = get_db()
    row = conn.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    conn.close()
    return row["value"] if row else default


def set_setting(key, value):
    conn = get_db()
    conn.execute("INSERT OR REPLACE INTO settings VALUES (?,?)", (key, value))
    conn.commit()
    conn.close()


# ----------------------------------------------------------------------------
# 登录鉴权
# ----------------------------------------------------------------------------
init_db()


def migrate_db():
    """老库平滑升级：补齐商机新增字段"""
    conn = get_db()
    cols = {r["name"] for r in conn.execute("PRAGMA table_info(deals)")}
    for col in DEAL_EXTRA_FIELDS:
        if col not in cols:
            conn.execute("ALTER TABLE deals ADD COLUMN %s %s"
                         % (col, "REAL" if col in ("amount2", "ratio") else "TEXT"))
    conn.commit()
    conn.close()


migrate_db()
app.secret_key = get_setting("secret_key") or app.secret_key      # 每次重启不变，会话不失效

_LOGIN_FAILS = {}          # ip -> [失败次数, 解封时间戳]
MAX_FAILS, LOCK_SECONDS = 5, 30


def current_user():
    uid = session.get("uid")
    if not uid:
        return None
    conn = get_db()
    row = conn.execute("SELECT * FROM users WHERE id=?", (uid,)).fetchone()
    conn.close()
    return row


def _now():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M")


PUBLIC_PATHS = ("/login", "/register", "/register/send", "/favicon.ico")

# 一级菜单路由 → 菜单 key，用于访问控制
MENU_BY_SEG = {"deals": "deals", "tech": "tech", "gov": "gov", "bond": "bond",
               "capmarket": "capmarket", "matcher": "matcher", "library": "library",
               "draft": "draft", "users": "users", "config": "config", "agent": "agent"}


@app.before_request
def _require_login():
    if request.path in PUBLIC_PATHS:
        return None
    me = current_user()
    if me is None:
        nxt = request.path if request.method == "GET" else "/"
        return redirect(url_for("login", next=nxt))
    # 默认密码未修改时，先强制改密码
    if me["must_change"] and request.path not in ("/password", "/logout"):
        return redirect("/password")
    # 菜单级访问控制：用户只能访问其角色已绑定的模块
    seg = request.path.split("/")[1] if request.path.startswith("/") else ""
    if seg in MENU_BY_SEG and MENU_BY_SEG[seg] not in user_menu_keys(me):
        return page('<div class="err">无权限访问该模块，请联系管理员调整角色菜单绑定。</div>', "")
    return None


LOGIN_HTML = r"""
<!DOCTYPE html>
<html lang="zh"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>登录 · 投行部数智化工作台</title>
<style>
*{box-sizing:border-box;margin:0;padding:0;font-family:'Microsoft YaHei',sans-serif}
body{min-height:100vh;display:flex;align-items:center;justify-content:center;
  background:linear-gradient(135deg,#12224a 0%,#1d3a75 55%,#2456c8 100%);font-size:13px;color:#1c2333}
.box{width:392px;background:#fff;border-radius:16px;padding:34px 34px 28px;
  box-shadow:0 24px 60px rgba(6,15,40,.35)}
.logo{display:flex;align-items:center;gap:10px;margin-bottom:6px}
.logo i{width:34px;height:34px;border-radius:9px;background:#2456c8;color:#fff;font-style:normal;
  display:flex;align-items:center;justify-content:center;font-weight:700;font-size:16px}
h1{font-size:17px;color:#12224a;letter-spacing:.5px}
.sub{color:#8b95ab;font-size:11.5px;margin:2px 0 22px 44px}
label{display:block;font-size:12px;color:#5b6478;margin:14px 0 6px}
input{padding:10px 12px;border:1px solid #dde3ec;border-radius:8px;font-size:13.5px;width:100%;background:#fff}
input:focus{outline:none;border-color:#2456c8;box-shadow:0 0 0 3px rgba(36,86,200,.12)}
button{width:100%;margin-top:20px;padding:11px;border:none;border-radius:8px;background:#2456c8;
  color:#fff;font-size:14px;cursor:pointer;letter-spacing:1px}
button:hover{background:#1d47a8}
.err{background:#fdecea;border:1px solid #f5c9c6;color:#a12722;border-radius:8px;
  padding:9px 12px;margin-bottom:6px;font-size:12.5px}
.row{display:flex;justify-content:space-between;align-items:center;margin-top:14px;font-size:12px;color:#5b6478}
.row input{width:auto}
.foot{margin-top:22px;padding-top:14px;border-top:1px solid #eef1f6;color:#a9b1c2;font-size:11px;line-height:1.8}
</style><script>function tg(id){var e=document.getElementById(id);if(e)e.classList.toggle("open")}function delDeal(id){if(confirm("确认删除该商机？"))location.href="/deals/delete/"+id}
function openDeal(){document.getElementById("dealModal").classList.add("on")}
function closeDeal(){document.getElementById("dealModal").classList.remove("on")}
function syncExport(aid,typ,fid,tid){var a=document.getElementById(aid),f=document.getElementById(fid),t=document.getElementById(tid);
if(!a||!f||!t)return;var q="/deals/export?typ="+encodeURIComponent(typ||"")+"&dfrom="+encodeURIComponent(f.value||"")+"&dto="+encodeURIComponent(t.value||"");
a.href=q;a.title="导出 "+(f.value||"不限")+" 至 "+(t.value||"不限");}</script></head><body>
<form class="box" method="post">
  <div class="logo"><i>投</i><h1>投行部 · 数智化工作台</h1></div>
  <div class="sub">Investment Banking Digital Workbench</div>
  {% if err %}<div class="err">{{ err }}</div>{% endif %}
  <label>用户名</label>
  <input name="username" autocomplete="username" autofocus required>
  <label>密码</label>
  <input name="password" type="password" autocomplete="current-password" required>
  <div class="row">
    <label style="margin:0;display:flex;align-items:center;gap:6px">
      <input type="checkbox" name="remember" value="1"> 7 天内免登录
    </label>
    <span>{{ now }}</span>
  </div>
  <input type="hidden" name="next" value="{{ next }}">
  <button type="submit">登 录</button>
  <div class="foot">首次使用默认账号 admin / th2026，登录后会要求修改密码。<br>
    还没有账号？<a href="/register" style="color:#2456c8;text-decoration:none">注册账号</a></div>
</form>
</body></html>
"""


@app.route("/login", methods=["GET", "POST"])
def login():
    if current_user():
        return redirect("/")
    err = ""
    next_url = request.values.get("next", "") or "/"
    if not next_url.startswith("/") or next_url.startswith("/login"):
        next_url = "/"

    if request.method == "POST":
        ip = (request.headers.get("X-Forwarded-For", "") or request.remote_addr or "").split(",")[0].strip()
        cnt, until = _LOGIN_FAILS.get(ip, [0, 0.0])
        now_ts = time.time()
        if until > now_ts:
            err = "尝试次数过多，请 %d 秒后再试" % int(until - now_ts + 1)
        else:
            username = (request.form.get("username") or "").strip()
            pwd = request.form.get("password") or ""
            conn = get_db()
            row = conn.execute("SELECT * FROM users WHERE username=?", (username,)).fetchone()
            if row and check_password_hash(row["pwd_hash"], pwd):
                conn.execute("UPDATE users SET last_login=? WHERE id=?", (_now(), row["id"]))
                conn.commit()
                conn.close()
                session.clear()
                session["uid"] = row["id"]
                session.permanent = bool(request.form.get("remember"))
                _LOGIN_FAILS.pop(ip, None)
                return redirect("/password" if row["must_change"] else next_url)
            conn.close()
            cnt += 1
            _LOGIN_FAILS[ip] = [cnt, now_ts + LOCK_SECONDS if cnt >= MAX_FAILS else 0.0]
            err = ("连续失败 %d 次，请 %d 秒后再试" % (cnt, LOCK_SECONDS)) if cnt >= MAX_FAILS \
                else "用户名或密码不正确"

    return render_template_string(LOGIN_HTML, err=err, next=next_url, now=_now())


# ---------------------------------------------------------------------------
# 注册（邮箱验证码；未配置 SMTP 时验证码直接在页面提示，便于内网使用）
# ---------------------------------------------------------------------------
_REG_CODES = {}          # email -> {"code": "123456", "exp": ts, "sent": ts}
REG_CODE_TTL = 300       # 验证码有效期 5 分钟
REG_RESEND_GAP = 60      # 同邮箱重发间隔（秒）


def _send_mail(to_addr, subject, body):
    """按 settings 里的 SMTP 配置发信；未配置或发送失败返回 False"""
    host = get_setting("smtp_host", "").strip()
    user = get_setting("smtp_user", "").strip()
    pwd = get_setting("smtp_pass", "").strip()
    if not (host and user and pwd):
        return False
    try:
        port = int(get_setting("smtp_port", "465") or 465)
        from_addr = get_setting("smtp_from", "").strip() or user
        msg = MIMEText(body, "plain", "utf-8")
        msg["Subject"] = Header(subject, "utf-8")
        msg["From"] = from_addr
        msg["To"] = to_addr
        if port == 465:
            srv = smtplib.SMTP_SSL(host, port, timeout=12)
        else:
            srv = smtplib.SMTP(host, port, timeout=12)
            try:
                srv.starttls()
            except smtplib.SMTPException:
                pass
        srv.login(user, pwd)
        srv.sendmail(from_addr, [to_addr], msg.as_string())
        srv.quit()
        return True
    except Exception:
        return False


def _gen_pwd(n=8):
    alphabet = string.ascii_letters + string.digits
    return "".join(random.choice(alphabet) for _ in range(n))


def _mail_ok(addr):
    return bool(_re_mail.match(r"^[\w.+-]+@[\w-]+(\.[\w-]+)+$", addr or ""))


REGISTER_HTML = r"""
<!DOCTYPE html>
<html lang="zh"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>注册 · 投行部数智化工作台</title>
<style>
*{box-sizing:border-box;margin:0;padding:0;font-family:'Microsoft YaHei',sans-serif}
body{min-height:100vh;display:flex;align-items:center;justify-content:center;
  background:linear-gradient(135deg,#12224a 0%,#1d3a75 55%,#2456c8 100%);font-size:13px;color:#1c2333}
.box{width:420px;background:#fff;border-radius:16px;padding:34px 34px 26px;box-shadow:0 24px 60px rgba(6,15,40,.35)}
.logo{display:flex;align-items:center;gap:10px;margin-bottom:6px}
.logo i{width:34px;height:34px;border-radius:9px;background:#2456c8;color:#fff;font-style:normal;
  display:flex;align-items:center;justify-content:center;font-weight:700;font-size:16px}
h1{font-size:17px;color:#12224a}
.sub{color:#8b95ab;font-size:11.5px;margin:2px 0 18px 44px}
label{display:block;font-size:12px;color:#5b6478;margin:14px 0 6px}
input{padding:10px 12px;border:1px solid #dde3ec;border-radius:8px;font-size:13.5px;width:100%}
input:focus{outline:none;border-color:#2456c8;box-shadow:0 0 0 3px rgba(36,86,200,.12)}
.coderow{display:flex;gap:8px}
.coderow input{flex:1}
.coderow button{width:120px;margin:0;padding:10px;font-size:12.5px;border-radius:8px;
  background:#eaf0fd;color:#2456c8;border:1px solid #c7d6f5;cursor:pointer;letter-spacing:0}
.coderow button:hover{background:#dce7fb}
.coderow button:disabled{opacity:.65;cursor:default}
button{width:100%;margin-top:20px;padding:11px;border:none;border-radius:8px;background:#2456c8;
  color:#fff;font-size:14px;cursor:pointer;letter-spacing:1px}
button:hover{background:#1d47a8}
.err{background:#fdecea;border:1px solid #f5c9c6;color:#a12722;border-radius:8px;padding:9px 12px;
  margin-bottom:6px;font-size:12.5px}
.tip{font-size:12px;color:#1a7f4d;margin-top:8px;min-height:16px;line-height:1.7}
.done{background:#eef8f1;border:1px solid #bfe3cc;border-radius:10px;padding:16px;margin-bottom:6px}
.done b.pwd{display:inline-block;margin:8px 0;padding:6px 14px;background:#fff;border:1px dashed #7bbd92;
  border-radius:8px;font-size:20px;letter-spacing:2px;color:#12224a}
.foot{margin-top:18px;padding-top:14px;border-top:1px solid #eef1f6;text-align:center;font-size:12px;color:#8b95ab}
.foot a{color:#2456c8;text-decoration:none}
</style></head><body>
<form class="box" method="post" autocomplete="off">
  <div class="logo"><i>投</i><h1>注册账号</h1></div>
  <div class="sub">注册后为「游客」角色，初始密码仅显示一次</div>
  {% if err %}<div class="err">{{ err }}</div>{% endif %}
  {% if done %}
  <div class="done">
    <div style="font-size:13.5px;color:#1a7f4d">✅ 注册成功！您的初始密码：</div>
    <b class="pwd">{{ done }}</b>
    <div style="font-size:12px;color:#5b6478;line-height:1.8">
      请妥善保存该密码，登录后需立即修改。账号角色为「游客」，如需更多菜单权限请联系管理员。
    </div>
  </div>
  <a href="/login" style="display:block"><button type="button">去登录</button></a>
  {% else %}
  <label>用户名</label>
  <input name="username" value="{{ username }}" placeholder="3-20 位字母、数字或下划线" required>
  <label>邮箱</label>
  <input name="email" id="email" type="email" value="{{ email }}" placeholder="用于接收验证码" required>
  <label>验证码</label>
  <div class="coderow">
    <input name="code" maxlength="6" placeholder="6 位数字" required>
    <button type="button" id="btnCode" onclick="sendCode()">获取验证码</button>
  </div>
  <div class="tip" id="tip"></div>
  <button type="submit">注 册</button>
  <div class="foot"><a href="/login">已有账号？去登录</a></div>
  {% endif %}
  <div style="text-align:center;font-size:11px;color:#a9b1c2;margin-top:12px">{{ now }}</div>
</form>
<script>
function sendCode(){
  var e=document.getElementById('email').value.trim();
  var tip=document.getElementById('tip'), btn=document.getElementById('btnCode');
  if(!e){tip.style.color='#a12722';tip.textContent='请先填写邮箱';return}
  btn.disabled=true; btn.textContent='发送中…';
  fetch('/register/send',{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},
    body:'email='+encodeURIComponent(e)})
  .then(function(r){return r.json()}).then(function(d){
    tip.style.color = d.ok ? '#1a7f4d' : '#a12722';
    tip.textContent = d.msg || '';
    if(d.ok){
      var n=60; btn.textContent=n+' 秒后重发';
      var t=setInterval(function(){n--; btn.textContent=n+' 秒后重发';
        if(n<=0){clearInterval(t); btn.disabled=false; btn.textContent='重新获取'}},1000);
    } else { btn.disabled=false; btn.textContent='获取验证码'; }
  }).catch(function(){tip.style.color='#a12722';tip.textContent='网络异常，请重试';
    btn.disabled=false; btn.textContent='获取验证码';});
}
</script>
</body></html>
"""


@app.route("/register", methods=["GET", "POST"])
def register():
    if current_user():
        return redirect("/")
    err = ""
    done_pwd = ""
    form = {"username": "", "email": ""}
    if request.method == "POST":
        username = (request.form.get("username") or "").strip()
        email = (request.form.get("email") or "").strip().lower()
        code = (request.form.get("code") or "").strip()
        form = {"username": username, "email": email}
        rec = _REG_CODES.get(email)
        now_ts = time.time()
        if not _re_mail.match(r"^[A-Za-z0-9_]{3,20}$", username):
            err = "用户名需为 3-20 位字母、数字或下划线"
        elif not _mail_ok(email):
            err = "邮箱格式不正确"
        elif not rec or rec["exp"] < now_ts:
            err = "验证码已失效，请重新获取"
        elif rec["code"] != code:
            err = "验证码不正确"
        else:
            conn = get_db()
            try:
                if conn.execute("SELECT 1 FROM users WHERE username=?", (username,)).fetchone():
                    err = "用户名已被占用"
                elif conn.execute("SELECT 1 FROM users WHERE lower(email)=?", (email,)).fetchone():
                    err = "该邮箱已注册"
                else:
                    pwd = _gen_pwd()
                    conn.execute(
                        "INSERT INTO users (username, pwd_hash, name, email, role, must_change, created) "
                        "VALUES (?,?,?,?,?,?,?)",
                        (username, generate_password_hash(pwd), "", email, "游客", 1, _now()))
                    conn.commit()
                    done_pwd = pwd
                    _REG_CODES.pop(email, None)
            finally:
                conn.close()
    return render_template_string(REGISTER_HTML, err=err, done=done_pwd,
                                  username=form["username"], email=form["email"], now=_now())


@app.route("/register/send", methods=["POST"])
def register_send():
    data = request.get_json(silent=True) or request.form
    email = (data.get("email") or "").strip().lower()
    if not _mail_ok(email):
        return jsonify({"ok": False, "msg": "邮箱格式不正确"})
    now_ts = time.time()
    rec = _REG_CODES.get(email)
    if rec and now_ts - rec.get("sent", 0) < REG_RESEND_GAP:
        return jsonify({"ok": False,
                        "msg": "请 %d 秒后再试" % int(REG_RESEND_GAP - (now_ts - rec["sent"]))})
    code = "%06d" % random.randint(0, 999999)
    _REG_CODES[email] = {"code": code, "exp": now_ts + REG_CODE_TTL, "sent": now_ts}
    sent = _send_mail(email, "投行部数智化工作台 · 注册验证码",
                      "您的注册验证码是：%s\n有效期 5 分钟。如非本人操作请忽略。" % code)
    if sent:
        return jsonify({"ok": True, "msg": "验证码已发送至 %s，请查收（5 分钟内有效）" % email})
    return jsonify({"ok": True,
                    "msg": "未配置邮件服务，本次验证码：%s（5 分钟内有效）" % code})


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

@app.route("/password", methods=["GET", "POST"])
def password():
    me = current_user()
    err = msg = ""
    if request.method == "POST":
        old = request.form.get("old") or ""
        new = request.form.get("new") or ""
        rep = request.form.get("rep") or ""
        if not check_password_hash(me["pwd_hash"], old):
            err = "原密码不正确"
        elif len(new) < 6:
            err = "新密码至少 6 位"
        elif new == old:
            err = "新密码不能与原密码相同"
        elif new != rep:
            err = "两次输入的新密码不一致"
        else:
            conn = get_db()
            conn.execute("UPDATE users SET pwd_hash=?, must_change=0 WHERE id=?",
                         (generate_password_hash(new), me["id"]))
            conn.commit()
            conn.close()
            return redirect("/?msg=pwd")
    body = """
    <div class="card" style="max-width:520px">
      <h2>修改密码</h2>
      %s
      %s
      <form method="post" class="frm">
        <label>原密码</label><input type="password" name="old" required>
        <label>新密码（至少 6 位）</label><input type="password" name="new" required>
        <label>确认新密码</label><input type="password" name="rep" required>
        <div style="margin-top:18px"><button type="submit">保存新密码</button>
          <a class="btn btn3" href="/" style="margin-left:8px">返回首页</a></div>
      </form>
    </div>""" % (
        '<div class="err">%s</div>' % err if err else "",
        '<div class="err">当前仍在使用初始密码，请先修改后再使用工作台。</div>'
        if me["must_change"] else "")
    return page(body, "")


def _admin_required():
    me = current_user()
    if not is_admin_user(me):
        return page('<div class="err">仅管理员可访问该页面。</div>', "")
    return None


def _role_label(role):
    if role in ("admin", "管理员"):
        return "管理员"
    if role == "user":
        return "普通用户"
    return role or "普通用户"


@app.route("/users")
def users():
    deny = _admin_required()
    if deny:
        return deny
    conn = get_db()
    roles = conn.execute("SELECT * FROM roles ORDER BY id").fetchall()
    if request.args.get("tab") == "roles":
        return _roles_page(conn, roles)

    rows = conn.execute("SELECT * FROM users ORDER BY id").fetchall()
    conn.close()
    msg = (request.args.get("msg") or "").strip()
    banner = alert_js = ""
    if msg:
        safe = esc_html(msg)
        banner = "<div class='ok'>%s</div>" % safe
        alert_js = (
            "<div class='modal on' id='msgModal' "
            "onclick=\"if(event.target===this)closeMsg()\">"
            "<div class='modal-box' style='width:360px'>"
            "<div class='modal-h'><b>提示</b>"
            "<a class='modal-x' href='javascript:closeMsg()'>×</a></div>"
            "<p style='padding:4px 2px 16px;font-size:14px'>%s</p>"
            "<div style='text-align:right'>"
            "<button type='button' onclick='closeMsg()'>确定</button></div>"
            "</div></div>"
            "<script>function closeMsg(){var m=document.getElementById('msgModal');"
            "if(m)m.classList.remove('on')}</script>" % safe)
    edit_u = None
    if request.args.get("edit"):
        for u in rows:
            if str(u["id"]) == request.args["edit"]:
                edit_u = u
                break
    trs = "".join(
        "<tr><td><b>%s</b></td><td>%s</td><td>%s</td><td>%s</td><td class='mini'>%s</td>"
        "<td class='mini'>%s</td><td style='white-space:nowrap'><a href='/users?edit=%d'>编辑</a>"
        "<a href='/users/reset/%d' style='margin-left:8px'>重置密码</a>"
        "%s</td></tr>" % (
            u["username"], u["name"] or "", u["email"] or "—", _role_label(u["role"]),
            u["created"] or "", u["last_login"] or "—", u["id"], u["id"],
            "" if u["id"] == current_user()["id"] else
            " <a href='/users/delete/%d' style='color:#d0342c;margin-left:8px'>删除</a>" % u["id"])
        for u in rows)
    role_opts = "".join("<option value='%s'>%s</option>" % (r["name"], r["name"]) for r in roles)
    role_opts2 = "".join("<option value='%s'%s>%s</option>" % (
        r["name"], " selected" if edit_u and r["name"] == edit_u["role"] else "", r["name"])
        for r in roles)
    edit_display = "block" if edit_u else "none"
    body = """
    {banner}
    <div class="card"><div class="tabs">
      <a class="tab on" href="/users">👤 用户列表</a>
      <a class="tab" href="/users?tab=roles">🔑 角色管理</a>
    </div></div>
    <div class="card"><h2>用户列表 <span class="mini">（仅管理员可见）</span>
      <span style="margin-left:auto"><a class="btn btn2 btn-sm" id="btnAddUser"
        href="javascript:showBox('addUserCard','btnAddUser')">＋ 新增用户</a></span></h2>
      <div class="tbl-wrap"><table>
        <tr><th>用户名</th><th>姓名</th><th>邮箱</th><th>角色</th><th>创建时间</th><th>最近登录</th><th>操作</th></tr>
        {trs}</table></div>
    </div>
    <div class="card" id="addUserCard" style="display:none"><h2>新增用户</h2>
      <form method="post" action="/users/add" class="frm">
        <div class="row">
          <div><label>用户名</label><input name="username" required placeholder="登录名，字母数字"></div>
          <div><label>姓名</label><input name="name" placeholder="张三"></div>
          <div><label>邮箱</label><input name="email" type="email" placeholder="name@example.com"></div>
          <div><label>初始密码</label><input name="password" required value="th2026"></div>
          <div><label>角色</label><select name="role">{role_opts}</select></div>
        </div>
        <div style="margin-top:16px"><button type="submit">添加用户</button>
          <a class="btn btn3" href="javascript:hideBox('addUserCard','btnAddUser')"
            style="margin-left:8px">取消</a></div>
      </form>
    </div>
    <div class="card" id="editUserCard" style="display:{edit_display}"><h2>{edit_title}</h2>
      <form method="post" action="/users/save" class="frm">
        <input type="hidden" name="id" value="{edit_id}">
        <div class="row">
          <div><label>用户名（不可修改）</label><input value="{edit_username}" disabled></div>
          <div><label>姓名</label><input name="name" value="{edit_name}"></div>
          <div><label>邮箱</label><input name="email" type="email" value="{edit_email}"></div>
          <div><label>角色</label><select name="role">{role_opts2}</select></div>
        </div>
        <div style="margin-top:16px"><button type="submit">保存修改</button>
          <a class="btn btn3" href="/users" style="margin-left:8px">取消</a></div>
      </form>
    </div>
    {alert_js}""".format(banner=banner, trs=trs, role_opts=role_opts, alert_js=alert_js,
                         role_opts2=role_opts2, edit_display=edit_display,
                         edit_title=("编辑用户 · %s" % edit_u["username"]) if edit_u else "编辑用户",
                         edit_id=edit_u["id"] if edit_u else "",
                         edit_username=edit_u["username"] if edit_u else "",
                         edit_name=(edit_u["name"] or "") if edit_u else "",
                         edit_email=(edit_u["email"] or "") if edit_u else "")
    return page(body, "users")


def _roles_page(conn, roles):
    edit = None
    if request.args.get("edit"):
        for r in roles:
            if str(r["id"]) == request.args["edit"]:
                edit = r
                break
        if edit and edit["name"] == "管理员":
            conn.close()
            return redirect("/users?tab=roles")
    conn.close()

    label_map = {m[0]: m[3] for m in MENUS}

    def menus_of(r):
        try:
            data = json.loads(r["menus"] or "[]")
        except (ValueError, TypeError):
            data = []
        if isinstance(data, str) and data == "ALL":
            return "全部菜单"
        if not data:
            return "（未绑定）"
        return "、".join(label_map.get(k, k) for k in data if k in label_map)

    trs = "".join(
        "<tr><td><b>%s</b></td><td>%s</td><td class='mini'>%s</td><td>%s</td></tr>" % (
            r["name"], menus_of(r), r["created"] or "",
            ("—" if r["name"] == "管理员" else
             "<a href='/users?tab=roles&edit=%d'>编辑</a> "
             "<a href='/roles/delete/%d' onclick=\"return confirm('确认删除该角色？')\">删除</a>"
             % (r["id"], r["id"])))
        for r in roles)

    sel = set()
    if edit:
        try:
            d = json.loads(edit["menus"] or "[]")
            sel = set(d) if isinstance(d, list) else set()
        except (ValueError, TypeError):
            sel = set()
        form_title = "编辑角色 · %s" % edit["name"]
        name_val, rid = edit["name"], edit["id"]
    else:
        form_title, name_val, rid = "新增角色", "", ""

    menu_checks = "".join(
        "<label style='display:inline-flex;align-items:center;gap:5px;margin:0 16px 10px 0'>"
        "<input type='checkbox' name='menus' value='%s'%s> %s</label>"
        % (k, " checked" if k in sel else "", lb)
        for (k, _, _, lb) in MENUS if k != "home")

    form_display = "block" if edit else "none"
    add_btn = ("" if edit else
               "<span style='margin-left:auto'><a class='btn btn2 btn-sm' id='btnAddRole' "
               "href=\"javascript:showBox('roleFormCard','btnAddRole')\">＋ 新增角色</a></span>")
    cancel_btn = ("<a class='btn btn3' href='/users?tab=roles' style='margin-left:8px'>取消</a>"
                  if edit else
                  "<a class='btn btn3' href=\"javascript:hideBox('roleFormCard','btnAddRole')\" "
                  "style='margin-left:8px'>取消</a>")

    body = """
    <div class="card"><div class="tabs">
      <a class="tab" href="/users">👤 用户列表</a>
      <a class="tab on" href="/users?tab=roles">🔑 角色管理</a>
    </div></div>
    <div class="card"><h2>角色列表{add_btn}</h2>
      <div class="tbl-wrap"><table>
        <tr><th>角色名</th><th>绑定菜单</th><th>创建时间</th><th>操作</th></tr>
        {trs}</table></div>
    </div>
    <div class="card" id="roleFormCard" style="display:{form_display}"><h2>{form_title}</h2>
      <form method="post" action="/roles/save" class="frm">
        <input type="hidden" name="id" value="{rid}">
        <div class="row"><div style="flex:1">
          <label>角色名 *</label><input name="name" required value="{name_val}" placeholder="如 科技金融"></div></div>
        <label style="margin-top:12px">绑定菜单（首页看板默认包含，管理员默认全部）</label>
        <div style="margin-top:8px">{menu_checks}</div>
        <div style="margin-top:16px"><button type="submit">保存角色</button>{cancel_btn}</div>
      </form>
    </div>""".format(add_btn=add_btn, trs=trs, form_display=form_display,
                   form_title=form_title, rid=rid, name_val=name_val,
                   menu_checks=menu_checks, cancel_btn=cancel_btn)
    return page(body, "users")


@app.route("/users/add", methods=["POST"])
def users_add():
    deny = _admin_required()
    if deny:
        return deny
    username = (request.form.get("username") or "").strip()
    pwd = request.form.get("password") or "th2026"
    role = (request.form.get("role") or "user").strip()
    if not username:
        return redirect("/users")
    conn = get_db()
    try:
        conn.execute(
            "INSERT INTO users (username, pwd_hash, name, email, role, must_change, created) "
            "VALUES (?,?,?,?,?,?,?)",
            (username, generate_password_hash(pwd), (request.form.get("name") or "").strip(),
             (request.form.get("email") or "").strip(), role, 1, _now()))
        conn.commit()
    except sqlite3.IntegrityError:
        pass          # 用户名重复，忽略
    conn.close()
    return redirect("/users")


@app.route("/users/save", methods=["POST"])
def users_save():
    deny = _admin_required()
    if deny:
        return deny
    f = request.form
    uid = (f.get("id") or "").strip()
    if not uid:
        return redirect("/users")
    conn = get_db()
    conn.execute("UPDATE users SET name=?, email=?, role=? WHERE id=?",
                 ((f.get("name") or "").strip(), (f.get("email") or "").strip(),
                  (f.get("role") or "游客").strip(), uid))
    conn.commit()
    conn.close()
    return redirect("/users?msg=用户信息已更新")


@app.route("/users/reset/<int:uid>")
def users_reset(uid):
    deny = _admin_required()
    if deny:
        return deny
    conn = get_db()
    conn.execute("UPDATE users SET pwd_hash=?, must_change=1 WHERE id=?",
                 (generate_password_hash("th2026"), uid))
    conn.commit()
    conn.close()
    return redirect("/users?msg=密码已重置")


@app.route("/users/delete/<int:uid>")
def users_delete(uid):
    deny = _admin_required()
    if deny:
        return deny
    if uid == current_user()["id"]:
        return redirect("/users")
    conn = get_db()
    n_admin = conn.execute(
        "SELECT COUNT(*) c FROM users WHERE role IN ('admin','管理员')").fetchone()["c"]
    target = conn.execute("SELECT * FROM users WHERE id=?", (uid,)).fetchone()
    if target and not (is_admin_user(target) and n_admin <= 1):
        conn.execute("DELETE FROM users WHERE id=?", (uid,))
        conn.commit()
    conn.close()
    return redirect("/users")


@app.route("/roles/save", methods=["POST"])
def roles_save():
    deny = _admin_required()
    if deny:
        return deny
    name = (request.form.get("name") or "").strip()
    rid = (request.form.get("id") or "").strip()
    menus = request.form.getlist("menus")
    if not name:
        return redirect("/users?tab=roles")
    conn = get_db()
    old_name = None
    if rid:
        cur = conn.execute("SELECT name FROM roles WHERE id=?", (rid,)).fetchone()
        if cur is None or cur["name"] == "管理员":
            conn.close()
            return redirect("/users?tab=roles")
        old_name = cur["name"]
    menus_json = json.dumps(["home"] + [m for m in menus if m in MENU_KEYS and m != "home"])
    if rid:
        if old_name != name:
            dup = conn.execute("SELECT 1 FROM roles WHERE name=? AND id!=?",
                               (name, rid)).fetchone()
            if dup:
                name = old_name     # 新名字已被占用，保持原名
        conn.execute("UPDATE roles SET name=?, menus=? WHERE id=?", (name, menus_json, rid))
        if old_name and old_name != name:
            conn.execute("UPDATE users SET role=? WHERE role=?", (name, old_name))
    else:
        exist = conn.execute("SELECT name FROM roles WHERE name=?", (name,)).fetchone()
        if exist:
            if name == "管理员":
                conn.close()
                return redirect("/users?tab=roles")
            conn.execute("UPDATE roles SET menus=? WHERE name=?", (menus_json, name))
        else:
            conn.execute("INSERT INTO roles (name, menus, created) VALUES (?,?,?)",
                         (name, menus_json, datetime.datetime.now().strftime("%Y-%m-%d %H:%M")))
    conn.commit()
    conn.close()
    return redirect("/users?tab=roles")


@app.route("/roles/delete/<int:rid>")
def roles_delete(rid):
    deny = _admin_required()
    if deny:
        return deny
    conn = get_db()
    cur = conn.execute("SELECT name FROM roles WHERE id=?", (rid,)).fetchone()
    if cur and cur["name"] != "管理员":
        conn.execute("UPDATE users SET role='user' WHERE role=?", (cur["name"],))
        conn.execute("DELETE FROM roles WHERE id=?", (rid,))
        conn.commit()
    conn.close()
    return redirect("/users?tab=roles")


# ----------------------------------------------------------------------------
# 通用工具：表格读取 / 关键字归一化 / DOCX 结构解析
# ----------------------------------------------------------------------------
def read_table(path):
    """读取 xlsx / csv → (headers, rows[[...]])"""
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        rows = [[("" if c is None else str(c)).strip() for c in r]
                for r in ws.iter_rows(values_only=True)]
        wb.close()
    elif ext == ".csv":
        data = open(path, "rb").read()
        try:
            text = data.decode("utf-8-sig")
        except UnicodeDecodeError:
            text = data.decode("gbk", "ignore")
        rows = [[c.strip() for c in r] for r in csv.reader(text.splitlines()) if any(c.strip() for c in r)]
    else:
        raise ValueError("仅支持 .xlsx / .csv")
    if not rows:
        return [], []
    header, data = rows[0], rows[1:]
    width = max(len(r) for r in rows)
    header += ["(空列%d)" % (i + 1) for i in range(len(header), width)]
    header = ["%d·%s" % (i + 1, h) if h == "" else h for i, h in enumerate(header)]
    data = [r + [""] * (width - len(r)) for r in data]
    return header, data


def norm_key(s):
    """关键字归一化：全角→半角、去空格、统一括号、小写"""
    if s is None:
        return ""
    out = []
    for ch in str(s):
        o = ord(ch)
        if 0xFF01 <= o <= 0xFF5E:
            ch = chr(o - 0xFEE0)
        out.append(ch)
    s = "".join(out)
    return s.replace("（", "(").replace("）", ")").replace(" ", "").lower()


def docx_outline(path, maxn=30):
    """解析 DOCX 标题结构（大纲级别 / 标题样式）"""
    try:
        xml = zipfile.ZipFile(path).read("word/document.xml").decode("utf-8", "ignore")
    except Exception:
        return []
    heads = []
    for m in re.finditer(r"<w:p[ >].*?</w:p>", xml, re.S):
        para = m.group(0)
        lm = re.search(r'<w:outlineLvl w:val="([0-8])"', para)
        if lm:
            lvl = int(lm.group(1))
        else:
            sm = re.search(r'<w:pStyle w:val="(?:Heading|标题)([1-3])"', para)
            if not sm:
                continue
            lvl = int(sm.group(1)) - 1
        if lvl > 2:
            continue
        t = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", para)).strip()
        if t and len(t) < 80:
            heads.append({"lvl": lvl, "t": t})
        if len(heads) >= maxn:
            break
    return heads


def docx_content(path):
    """提取 docx 全文（按文档顺序：标题/正文/表格），用于初稿生成
    返回 [(kind, payload), ...]，kind ∈ h1..h4 / p / table
    标题判定：① 大纲级别 outlineLvl ② 样式名含 heading/标题+数字 ③ 中文序号段（一、/（一））"""
    try:
        z = zipfile.ZipFile(path)
        xml = z.read("word/document.xml").decode("utf-8", "ignore")
    except Exception:
        return []
    items = []
    for m in re.finditer(r"<w:(p|tbl)[ >].*?</w:\1>", xml, re.S):
        tag, frag = m.group(1), m.group(0)
        if tag == "tbl":
            rows = []
            for tr in re.finditer(r"<w:tr[ >].*?</w:tr>", frag, re.S):
                cells = []
                for tc in re.finditer(r"<w:tc[ >].*?</w:tc>", tr.group(0), re.S):
                    cells.append("".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", tc.group(0))).strip())
                if cells:
                    rows.append(cells)
            if rows:
                items.append(("table", rows))
            continue
        text = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", frag)).strip()
        if not text:
            continue
        lvl = None
        mo = re.search(r'<w:outlineLvl w:val="([0-8])"', frag)
        if mo and len(text) < 60:  # 长段落即使带大纲级别也视为正文（银行材料正文常挂编号大纲）
            lvl = int(mo.group(1)) + 1
        else:
            ms = re.search(r'<w:pStyle w:val="([^"]*)"', frag)
            if ms:
                mm = re.search(r"heading\s*([1-9])", ms.group(1), re.I) or \
                     re.match(r"^([1-9])$", ms.group(1))
                if mm and (len(text) < 60 or mo):  # 样式为标题，或该段确有大纲级别
                    lvl = int(mm.group(1))
        if lvl is None and len(text) < 50:  # 中文序号段兜底
            if re.match(r"^[一二三四五六七八九十]+、", text):
                lvl = 1
            elif re.match(r"^[（(][一二三四五六七八九十]+[）)]", text):
                lvl = 2
        items.append(("h%d" % min(lvl, 4), text) if lvl else ("p", text))
    return items


def has_data(text):
    """段落中含数字/百分比/金额 → 视为需人工核对的当期数据段"""
    return bool(re.search(r"[0-9０-９]\s*[%％]|(?<![0-9])[0-9][0-9,.，]{2,}(?![0-9])|亿|万元|亿元", text))


def scan_materials(root, trigger="手动"):
    """扫描素材库目录 → 写入 materials 表，并记录本次增/改/删差异"""
    conn = get_db()
    old = {r["rel"]: dict(r) for r in conn.execute("SELECT * FROM materials").fetchall()}
    old_keys = set(old)
    new_items, new_keys = [], set()
    if root and os.path.isdir(root):
        for dirpath, dirnames, filenames in os.walk(root):
            dirnames[:] = [d for d in dirnames if not d.startswith(".")]
            for fn in filenames:
                if fn.startswith("~$") or os.path.splitext(fn)[1].lower() in (".7z", ".zip", ".rar"):
                    continue
                p = os.path.join(dirpath, fn)
                ext = os.path.splitext(fn)[1].lower().lstrip(".")
                kind = {"docx": "docx", "doc": "doc", "pptx": "pptx", "pdf": "pdf",
                        "xlsx": "xlsx", "xls": "xls", "csv": "csv"}.get(ext, "other")
                rel = os.path.relpath(p, root)
                if "年度总结" in rel or "总结及策略" in rel:
                    cat = "年度总结及计划"
                elif rel.startswith("投行材料"):
                    cat = "投行条线汇报"
                else:
                    cat = "综合汇报与上会"
                st = os.stat(p)
                new_items.append((p, fn, rel, kind, round(st.st_size / 1024, 1),
                                  datetime.datetime.fromtimestamp(st.st_mtime).strftime("%Y-%m-%d"), cat))
                new_keys.add(rel)
    added = len(new_keys - old_keys)
    removed = len(old_keys - new_keys)
    new_by_rel = {x[2]: x for x in new_items}
    updated = sum(1 for k in (new_keys & old_keys)
                  if new_by_rel[k][4] != old[k]["size"] or new_by_rel[k][5] != old[k]["mtime"])
    # 按 path upsert：保持素材 id 稳定（初稿引用、书签不失效）
    for row in new_items:
        conn.execute("""INSERT INTO materials (path,name,rel,kind,size,mtime,cat) VALUES (?,?,?,?,?,?,?)
                        ON CONFLICT(path) DO UPDATE SET name=excluded.name, rel=excluded.rel,
                        kind=excluded.kind, size=excluded.size, mtime=excluded.mtime, cat=excluded.cat""", row)
    if new_items:
        conn.execute("DELETE FROM materials WHERE path NOT IN (%s)"
                     % ",".join("?" * len(new_items)), [r[0] for r in new_items])
    else:
        conn.execute("DELETE FROM materials")
    if trigger != "启动" or (added + updated + removed) > 0:
        conn.execute("INSERT INTO scan_logs (ts,trigger,added,updated,removed,total) VALUES (?,?,?,?,?,?)",
                     (datetime.datetime.now().strftime("%Y-%m-%d %H:%M"), trigger,
                      added, updated, removed, len(new_items)))
    conn.commit()
    conn.close()
    return len(new_items)


def auto_scan_worker():
    """后台线程：每日到点自动重扫素材库（程序运行期间生效）"""
    while True:
        time.sleep(60)
        try:
            if get_setting("auto_scan_on", "1") != "1":
                continue
            now = datetime.datetime.now()
            hhmm = get_setting("auto_scan_time", "08:30")
            try:
                tgt_h, tgt_m = int(hhmm[:2]), int(hhmm[3:5])
            except Exception:
                tgt_h, tgt_m = 8, 30
            today = now.strftime("%Y-%m-%d")
            if get_setting("last_auto_scan_date", "") == today:
                continue
            if (now.hour, now.minute) < (tgt_h, tgt_m):
                continue
            root = get_setting("material_root", "")
            if root and os.path.isdir(root):
                scan_materials(root, trigger="定时")
                set_setting("last_auto_scan_date", today)
                print("[auto-scan] %s 已完成每日素材库扫描" % now.strftime("%H:%M"))
        except Exception as e:
            print("[auto-scan] error:", e)


# ----------------------------------------------------------------------------
# 页面模板
# ----------------------------------------------------------------------------
BASE = r"""
<!DOCTYPE html>
<html lang="zh"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>投行 · 数智化工作台</title><style>
*{box-sizing:border-box;margin:0;padding:0;font-family:'Microsoft YaHei',sans-serif}
body{background:#f4f6fa;color:#1c2333;font-size:13px;display:flex;flex-direction:column;min-height:100vh}
a{text-decoration:none;color:#2456c8}
.top{background:#12224a;color:#fff;padding:14px 22px;display:flex;justify-content:space-between;align-items:center}
.top h1{font-size:17px}.top .sub{color:#8ea1cc;font-size:11px;margin-top:2px}
.wrap{display:flex;flex:1 1 auto;min-height:0}
.main-col{flex:1 1 auto;display:flex;flex-direction:column;min-width:0}
.compliance{display:flex;gap:9px;align-items:flex-start;margin:0;
  padding:6px 22px 6px 19px;background:#fdf4e0;border-bottom:1px solid #e8d5a8;
  border-left:4px solid #c8963e;font-size:13px;line-height:1.55;color:#9c4a12;font-weight:700}
.compliance .c-ico{flex-shrink:0;width:16px;height:16px;border-radius:50%;background:#c8963e;color:#fff;
  font-size:11px;font-weight:700;display:flex;align-items:center;justify-content:center;margin-top:2px;
  box-shadow:0 0 0 3px rgba(200,150,62,.18)}
.compliance b{color:#fff;background:#b5822e;font-weight:700;font-size:12px;letter-spacing:.5px;
  padding:1px 7px 2px;border-radius:4px;margin-right:5px;display:inline-block;vertical-align:1px}
.compliance .hl{color:#9c4a12;font-weight:700}
.compliance .c-txt{letter-spacing:.15px;font-weight:700;color:#9c4a12}
.side{width:170px;background:#12224a;padding:10px 8px;flex-shrink:0}
.side a{display:block;color:#c9d4ee;padding:10px 12px;border-radius:8px;font-size:13px;margin-bottom:2px}
.side a:hover{background:rgba(255,255,255,.08)}
.side a.on{background:#2456c8;color:#fff;font-weight:600}
.side a.sub{padding:6px 12px 6px 24px;font-size:12px;color:#9fb0d6;border-radius:6px;margin-bottom:1px}
.side a.sub:hover{background:rgba(255,255,255,.06);color:#fff}
.side a.sub.on{background:rgba(36,86,200,.45);color:#fff;font-weight:600}
.main{flex:1 1 auto;padding:8px 22px 16px;min-width:0}
.card{background:#fff;border:1px solid #e5e9f0;border-radius:12px;padding:16px 18px;margin-bottom:12px}
h2{font-size:15px;margin-bottom:12px}
.tbl-wrap{width:100%;max-width:100%;overflow-x:auto;overflow-y:visible;padding-bottom:4px}
.tbl-wrap::-webkit-scrollbar{height:10px;width:10px}
.tbl-wrap::-webkit-scrollbar-thumb{background:#c7d6f5;border-radius:5px}
.tbl-wrap::-webkit-scrollbar-thumb:hover{background:#9db8ea}
.tbl-wrap::-webkit-scrollbar-track{background:#f2f5fa;border-radius:5px}
table{width:100%;border-collapse:collapse;font-size:12.5px}
th{background:#f7f9fc;color:#5b6478;text-align:left;padding:9px 10px;border-bottom:1px solid #e5e9f0;white-space:nowrap}
td{padding:9px 10px;border-bottom:1px solid #eef1f6;vertical-align:top}
tr:hover td{background:#f8faff}
input,select,textarea{padding:7px 10px;border:1px solid #dde3ec;border-radius:7px;font-size:13px;width:100%;background:#fff}
button,.btn{padding:8px 16px;border:none;border-radius:7px;background:#2456c8;color:#fff;font-size:13px;cursor:pointer}
.btn2{background:#eaf0fd;color:#2456c8;border:1px solid #c7d6f5}
.btn3{background:#fff;color:#5b6478;border:1px solid #dde3ec}
.kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:14px}
.kpi{background:#fff;border:1px solid #e5e9f0;border-radius:12px;padding:14px 16px}
.kpi .t{color:#5b6478;font-size:12px}.kpi .v{font-size:22px;font-weight:700;margin-top:6px}
.kpi .v small{font-size:12px;font-weight:400;color:#5b6478}
.pill{display:inline-block;padding:2px 9px;border-radius:99px;font-size:11px;font-weight:600}
.p1{background:#eaf0fd;color:#2456c8}.p2{background:#e8f6ee;color:#1a7f4d}
.p3{background:#fdf4e3;color:#b97a12}.p4{background:#fdecea;color:#d0342c}
.p5{background:#f0ecfb;color:#6d4fc2}.p6{background:#f0f1f5;color:#5b6478}
.frm label{display:block;font-size:12px;color:#5b6478;margin:10px 0 4px}
.row{display:flex;gap:14px;flex-wrap:wrap}.row>div{flex:1;min-width:160px}
.mini{font-size:11.5px;color:#5b6478}
.err{background:#fdecea;border:1px solid #f5c9c6;color:#a12722;border-radius:8px;padding:10px 14px;margin-bottom:12px}
.ok{background:#e8f6ee;border:1px solid #bfe5cd;color:#14683d;border-radius:8px;padding:10px 14px;margin-bottom:12px}
.up{color:#d0342c;font-weight:700}
pre.draft{white-space:pre-wrap;font-size:13px;line-height:1.9;background:#fbfcfe;border:1px solid #e5e9f0;border-radius:8px;padding:16px;max-height:480px;overflow:auto}
.grp{margin-bottom:18px}
.grp-h{display:flex;align-items:center;gap:10px;margin:0 0 10px;padding-bottom:8px;border-bottom:1px dashed #e5e9f0}
.grp-h .t{font-weight:700;font-size:14px;color:#12224a}
.grp-h .n{font-size:11.5px;color:#8b95ab}
.deal{background:#fff;border:1px solid #e8ecf3;border-radius:10px;padding:12px 14px;margin-bottom:10px}
.deal:hover{border-color:#c7d6f5;box-shadow:0 2px 10px rgba(36,86,200,.06)}
.deal-h{display:flex;align-items:baseline;gap:10px;flex-wrap:wrap}
.deal-h b{font-size:13.5px;color:#12224a}
.deal-meta{font-size:11.5px;color:#5b6478;margin-top:6px;line-height:1.9}
.deal-meta em{font-style:normal;color:#8b95ab;margin-right:4px}
.deal-t{font-size:12.5px;color:#3d4657;line-height:1.85;margin-top:8px;padding:8px 10px;
  background:#f8fafd;border-left:3px solid #c7d6f5;border-radius:0 6px 6px 0;
  max-height:72px;overflow:hidden}
.deal-t.open{max-height:none}
.deal-t .fill{color:#2456c8}
.deal-f{display:flex;align-items:center;gap:10px;margin-top:8px;font-size:11.5px;color:#8b95ab}
.deal-f a{font-size:11.5px}
.two-col td{vertical-align:top}
.two-col .td-type{width:168px;padding-right:14px}
.c-name{font-weight:700;font-size:13px;color:#12224a;display:flex;align-items:center;gap:8px;flex-wrap:wrap}
.c-ops{margin-left:8px;font-weight:400;font-size:11.5px;color:#2456c8;white-space:nowrap}
.c-ops.del{color:#d0342c;margin-left:10px}
.c-text{font-size:12.5px;color:#3d4657;line-height:1.85;margin-top:6px}
.c-meta{font-size:11.5px;color:#8b95ab;margin-top:6px;line-height:1.8}
.modal{position:fixed;inset:0;background:rgba(12,20,40,.45);display:none;align-items:center;
  justify-content:center;z-index:400;padding:20px}
.modal.on{display:flex}
.modal-box{width:880px;max-width:94vw;max-height:88vh;overflow:auto;background:#fff;
  border-radius:14px;padding:18px 20px 16px;box-shadow:0 24px 60px rgba(6,15,40,.35)}
.modal-h{display:flex;align-items:center;justify-content:space-between;margin-bottom:6px;
  padding-bottom:10px;border-bottom:1px solid #eef1f6}
.modal-h b{font-size:14.5px;color:#12224a}
.modal-x{font-size:20px;color:#8b95ab;line-height:1;text-decoration:none}
.modal-x:hover{color:#d0342c}
.pager{display:flex;align-items:center;gap:10px;margin-top:12px}
.tabs{display:flex;gap:2px;border-bottom:1px solid #e5e9f0;flex-wrap:wrap;margin-bottom:2px}
.tab{padding:9px 15px;font-size:13px;color:#5b6478;border:1px solid transparent;border-bottom:none;
  border-radius:8px 8px 0 0;display:flex;align-items:center;gap:6px;margin-bottom:-1px;white-space:nowrap}
.tab:hover{background:#f7f9fc;color:#2456c8}
.tab.on{background:#fff;border-color:#e5e9f0;color:#2456c8;font-weight:700;box-shadow:inset 0 2px 0 #2456c8}
.tab span{font-size:11px;background:#eef1f7;color:#5b6478;border-radius:99px;padding:1px 7px;line-height:1.6}
.tab.on span{background:#2456c8;color:#fff}
.one-col td{padding:12px 14px}
.c-date{font-size:11.5px;color:#8b95ab;font-weight:400;margin-left:2px}
.c-name>.c-ops:first-of-type{margin-left:auto}
.deal-head{display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:12px}
.detail-title{font-size:16px;font-weight:700;color:#12224a;margin:6px 0 12px;padding-left:12px;
  border-left:4px solid #2456c8}
.detail-text{background:#fbfcfe;border:1px solid #e8ecf3;border-radius:10px;
  padding:20px 22px;font-size:14px;line-height:2.1;color:#2b3444;white-space:pre-wrap;
  letter-spacing:.2px;min-height:120px}
.edit-area{width:100%;padding:16px;border:1px solid #dde3ec;border-radius:10px;font-size:14px;
  line-height:2;font-family:'Microsoft YaHei',sans-serif;background:#fbfcfe;resize:vertical}
.edit-area:focus{outline:none;border-color:#2456c8;box-shadow:0 0 0 3px rgba(36,86,200,.12)}
table.kv th{background:#f7f9fc;color:#5b6478;font-weight:600;font-size:12.5px}
table.kv td{font-size:12.5px}
.mod-head{margin:16px 0 2px;padding:12px 14px;background:#f8fafd;border:1px solid #e8ecf3;border-radius:10px}
.mod-t{font-size:15px;font-weight:700;color:#12224a}
.mod-d{font-size:12px;color:#8b95ab;margin-top:4px}
.mod-s{font-size:12px;color:#5b6478;margin-top:8px}
</style><script>function tg(id){var e=document.getElementById(id);if(e)e.classList.toggle("open")}function delDeal(id){if(confirm("确认删除该商机？"))location.href="/deals/delete/"+id}
function openDeal(){document.getElementById("dealModal").classList.add("on")}
function closeDeal(){document.getElementById("dealModal").classList.remove("on")}
function syncExport(aid,typ,fid,tid){var a=document.getElementById(aid),f=document.getElementById(fid),t=document.getElementById(tid);
if(!a||!f||!t)return;var q="/deals/export?typ="+encodeURIComponent(typ||"")+"&dfrom="+encodeURIComponent(f.value||"")+"&dto="+encodeURIComponent(t.value||"");
a.href=q;a.title="导出 "+(f.value||"不限")+" 至 "+(t.value||"不限");}
function showBox(id,btnId){var e=document.getElementById(id);if(e){e.style.display='block'}
if(btnId){var b=document.getElementById(btnId);if(b)b.style.display='none'}}
function hideBox(id,btnId){var e=document.getElementById(id);if(e){e.style.display='none'}
if(btnId){var b=document.getElementById(btnId);if(b)b.style.display=''}}</script></head><body>
<div class="top"><div><h1>投行部 · 数智化工作台</h1>
</div>
<div style="font-size:12px;color:#8ea1cc">{{ now }} · {{ user_name }}
  <a href="/password" style="color:#c9d4ee;margin-left:10px">修改密码</a>
  <a href="/logout" style="color:#c9d4ee;margin-left:8px">退出</a></div></div>
<div class="wrap">
<nav class="side">
  {% for m in nav_items %}
  <a href="{{ m.url }}" class="{{ 'on' if act==m.key }}">{{ m.icon }} {{ m.label }}</a>
  {% endfor %}
</nav>
<div class="main-col">
<div class="compliance">
  <span class="c-ico">!</span>
  <div class="c-txt"><b>数据合规提示</b>本系统不与行内业务系统对接，
    <span class="hl">不读取、不存储任何行内业务数据</span>；客户信息由部门管理员及客户经理自行维护，
    商机与风险信息来自公开渠道采集，仅供营销辅助参考，
    <span class="hl">不构成授信决策依据</span>，涉及关键决策请回行内权威渠道核实。
    <span class="hl">本平台为非涉密平台；严禁传输、处理涉密内容。</span></div>
</div>
<div class="main">{{ body|safe }}</div></div></div></body></html>
"""


# ----------------------------------------------------------------------------
# RBAC：菜单定义与角色权限
# ----------------------------------------------------------------------------
MENUS = [
    ("home", "/", "📊", "首页看板"),
    ("deals", "/deals", "💡", "商机管理"),
    ("tech", "/tech", "🔬", "科技金融"),
    ("gov", "/gov", "🏛", "综合融资"),
    ("bond", "/bond", "📜", "债券融资"),
    ("capmarket", "/capmarket", "📈", "资本市场"),
    ("matcher", "/matcher", "🔗", "跨表匹配"),
    ("library", "/library", "📚", "素材库"),
    ("draft", "/draft", "📝", "初稿生成"),
    ("users", "/users", "👤", "用户管理"),
    ("config", "/config", "⚙", "配置管理"),
    ("agent", "/agent", "🤖", "智能体配置"),
]
MENU_KEYS = [m[0] for m in MENUS]
# 未绑定角色（或历史 user 角色）的兜底可见菜单：全部业务菜单，不含管理菜单
BUSINESS_KEYS = ["home", "deals", "tech", "gov", "bond", "capmarket",
                 "matcher", "library", "draft"]


def is_admin_user(me):
    return bool(me and me["role"] in ("admin", "管理员"))


def user_menu_keys(me):
    """返回用户可见菜单 key 集合。管理员=全部；绑定角色=角色菜单；无角色=业务菜单兜底。"""
    if not me:
        return set()
    if is_admin_user(me):
        return set(MENU_KEYS)
    conn = get_db()
    row = conn.execute("SELECT menus FROM roles WHERE name=?", (me["role"],)).fetchone()
    conn.close()
    if not row:
        return set(BUSINESS_KEYS)
    try:
        data = json.loads(row["menus"] or "[]")
    except (ValueError, TypeError):
        return set(BUSINESS_KEYS)
    if isinstance(data, str) and data == "ALL":
        return set(MENU_KEYS)
    if isinstance(data, list):
        return {k for k in data if k in MENU_KEYS} | {"home"}
    return set(BUSINESS_KEYS)


def page(body, act):
    me = current_user()
    visible = user_menu_keys(me)
    nav_items = [{"key": k, "url": u, "icon": ic, "label": lb}
                 for (k, u, ic, lb) in MENUS if k in visible]
    return render_template_string(BASE, body=body, act=act, req=request,
                                  now=datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
                                  user_name=(me["name"] or me["username"]) if me else "",
                                  is_admin=is_admin_user(me),
                                  nav_items=nav_items)


# ----------------------------------------------------------------------------
# 配置管理 / 智能体配置 / 宿主机授权（实现在 wb_config.py）
# ----------------------------------------------------------------------------
import wb_config

wb_config.init(DB_PATH, page, os.path.join(BASE_DIR, "uploads", "cfg"))
app.register_blueprint(wb_config.bp)
PUBLIC_PATHS = tuple(PUBLIC_PATHS) + ("/api/host/tasks", "/api/host/report", "/api/host/pair")


# ----------------------------------------------------------------------------
# 1. 首页看板
# ----------------------------------------------------------------------------
@app.route("/")
def home():
    conn = get_db()
    n_deal = conn.execute("SELECT COUNT(*) c FROM deals").fetchone()["c"]
    n_ing = conn.execute("SELECT COUNT(*) c FROM deals WHERE status='进行中'").fetchone()["c"]
    amt = conn.execute("SELECT COALESCE(SUM(amount),0) s FROM deals WHERE status='进行中'").fetchone()["s"]
    runs = conn.execute("SELECT * FROM match_runs ORDER BY id DESC LIMIT 8").fetchall()
    n_mat = conn.execute("SELECT COUNT(*) c FROM materials").fetchone()["c"]
    conn.close()
    rows = "".join(
        "<tr><td>%s</td><td class='mini'>%s ↔ %s（%s / %s）</td>"
        "<td>%d</td><td><span class='pill p2'>精确 %d</span> <span class='pill p3'>模糊 %d</span> "
        "<span class='pill p4'>未匹配 %d</span></td><td><a href='/matcher/result/%d'>下载</a></td></tr>"
        % (r["ts"], r["file_a"], r["file_b"], r["col_a"], r["col_b"],
           r["total"], r["exact"], r["fuzzy"], r["unmatched"], r["id"])
        for r in runs) or "<tr><td class='mini' colspan=5>暂无记录</td></tr>"
    banner = ('<div class="ok">密码已更新，请牢记新密码。如需新增同事账号，'
              '可在左侧「用户管理」中添加。</div>'
              if request.args.get("msg") == "pwd" else "")
    body = """
    {banner}
    <div class="kpis">
      <div class="kpi"><div class="t">商机总数</div><div class="v">{n_deal} <small>条</small></div>
        <div class="mini">进行中 {n_ing} 条</div></div>
      <div class="kpi"><div class="t">在途商机金额（进行中）</div><div class="v">{amt:,.0f} <small>万元</small></div></div>
      <div class="kpi"><div class="t">素材库文件</div><div class="v">{n_mat} <small>个</small></div>
        <div class="mini"><a href="/library/scan">↻ 重新扫描</a></div></div>
      <div class="kpi"><div class="t">累计匹配任务</div><div class="v">{n_run} <small>次</small></div>
        <div class="mini">替代手工 VLOOKUP</div></div>
    </div>
    <div class="card"><h2>最近匹配任务</h2><table>
      <tr><th>时间</th><th>文件</th><th>总行</th><th>结果</th><th></th></tr>{rows}</table></div>
    <div class="card"><h2>使用指引</h2>
    <p class="mini" style="line-height:2">
    ① <b>商机管理</b>：团队共用台账，按阶段推进、可筛选导出；<br>
    ② <b>跨表匹配</b>：上传任意两张 xlsx/csv，选关键列（如统一社会信用代码/企业名称），自动完成精确+模糊匹配，替代每周手工 VLOOKUP；<br>
    ③ <b>素材库</b>：扫描本地材料文件夹，自动解析 Word 材料的标题结构（模板骨架）；<br>
    ④ <b>初稿生成</b>：选历史材料作结构参照，结合商机数据，一键生成供修改的初稿。</p></div>
    """.format(n_deal=n_deal, n_ing=n_ing, amt=amt or 0, n_mat=n_mat,
               n_run=len(runs), rows=rows, banner=banner)
    return page(body, "home")


# ----------------------------------------------------------------------------
# 2. 商机管理
# ----------------------------------------------------------------------------
@app.route("/deals")
def deals():
    q = request.args.get("q", "").strip()
    stage = request.args.get("stage", "")
    status = request.args.get("status", "")
    typ_raw = request.args.get("typ", "").strip()
    typ = DEAL_KEY_TYPE.get(typ_raw, typ_raw)
    typ_key = DEAL_TYPE_KEY.get(typ, typ)
    msg = request.args.get("msg", "")
    dfrom = (request.args.get("dfrom") or "").strip()
    dto = (request.args.get("dto") or "").strip()
    try:
        cur_page = max(1, int(request.args.get("page", "1") or "1"))
    except (TypeError, ValueError):
        cur_page = 1

    edit = None
    if request.args.get("edit"):
        conn = get_db()
        edit = conn.execute("SELECT * FROM deals WHERE id=?", (request.args["edit"],)).fetchone()
        conn.close()

    sql, args = "SELECT * FROM deals WHERE 1=1", []
    if q:
        sql += (" AND (name LIKE ? OR client LIKE ? OR owner LIKE ? OR sector LIKE ?"
                " OR counterparty LIKE ? OR note LIKE ?)")
        args += ["%%%s%%" % q] * 6
    if stage:
        sql += " AND stage=?"
        args.append(stage)
    if status:
        sql += " AND status=?"
        args.append(status)
    if typ:
        sql += " AND deal_type=?"
        args.append(typ)
    if dfrom:
        sql += " AND COALESCE(deal_date, substr(created,1,10)) >= ?"
        args.append(dfrom)
    if dto:
        sql += " AND COALESCE(deal_date, substr(created,1,10)) <= ?"
        args.append(dto)
    sql += " ORDER BY updated DESC"

    conn = get_db()
    rows = conn.execute(sql, args).fetchall()
    if typ:
        stat = conn.execute("SELECT stage, COUNT(*) c, COALESCE(SUM(amount),0) s FROM deals"
                            " WHERE deal_type=? GROUP BY stage", (typ,)).fetchall()
        tsum = conn.execute("SELECT COUNT(*) c, COALESCE(SUM(amount),0) s FROM deals WHERE deal_type=?",
                            (typ,)).fetchone()
    else:
        stat = conn.execute("SELECT stage, COUNT(*) c, COALESCE(SUM(amount),0) s FROM deals GROUP BY stage").fetchall()
        tsum = conn.execute("SELECT COUNT(*) c, COALESCE(SUM(amount),0) s FROM deals").fetchone()
    typ_stat = conn.execute("SELECT COALESCE(deal_type,'') t, COUNT(*) c FROM deals GROUP BY deal_type").fetchall()
    conn.close()
    tmap = {r["t"]: r["c"] for r in typ_stat}

    def pill_stage(s):
        i = STAGES.index(s) if s in STAGES else 0
        return ["p1", "p5", "p3", "p1", "p2", "p6"][i]

    def money_range(r):
        a, b = r["amount"] or 0, r["amount2"] or 0
        if b and a:
            return "%s ~ %s 万" % (format(a, ",.0f"), format(b, ",.0f"))
        if a:
            return "%s 万" % format(a, ",.0f")
        return "—"

    def type_cell(r):
        t = r["deal_type"] or ""
        if not t:
            return "<span class='pill p6'>未分类</span>"
        return "<span class='pill %s'>%s</span>" % (DEAL_TYPE_PILL.get(t, "p6"), t)

    def _esc(x):
        return str(x or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    def _date_of(r):
        return (r["deal_date"] or (r["created"] or "")[:10] or "")

    def _meta(r):
        bits = []
        if r["client"]:
            bits.append(_esc(r["client"]))
        if r["counterparty"]:
            bits.append("标的 " + _esc(r["counterparty"]))
        bits.append(money_range(r))
        if r["org"]:
            bits.append(_esc(r["org"]))
        for k in ("notice_date", "location"):
            if r[k]:
                bits.append(_esc(r[k]))
        return " ｜ ".join(bits)

    def _row2(r):
        """全部视角：商机类别 / 具体内容"""
        t = r["deal_type"] or "未分类"
        text = (r["note"] or "").strip() or (r["name"] or "")
        return (
            "<tr><td class='td-type'><span class='pill %s'>%s</span>%s"
            "<div class='mini'>%s · %s</div></td>"
            "<td><div class='c-name'><span>%s</span><span class='c-date'>%s</span>"
            "<a class='c-ops' href='/deals/view/%d'>查看详情</a>"
            "<a class='c-ops' href='/deals/%d/edit'>编辑</a>"
            "<a class='c-ops del' href='javascript:delDeal(%d)'>删除</a></div>"
            "<div class='c-text'>%s</div></td></tr>"
        ) % (
            DEAL_TYPE_PILL.get(t, "p6"), _esc(t),
            ("<div class='mini'>%s</div>" % _esc(r["sub_type"])) if r["sub_type"] else "",
            _esc(r["stage"]), _esc(r["status"]),
            _esc(r["name"]), _date_of(r), r["id"], r["id"], r["id"], _esc(text))

    def _row1(r, seq=0):
        """模块视角：序号 + 具体内容"""
        text = (r["note"] or "").strip() or (r["name"] or "")
        seq_cell = ("<td class='mini' style='text-align:center;color:#8b95ab'>%d</td>" % seq) if seq else ""
        return (
            "<tr>%s<td><div class='c-name'><span>%s</span><span class='c-date'>%s</span>"
            "<a class='c-ops' href='/deals/view/%d'>查看详情</a>"
            "<a class='c-ops' href='/deals/%d/edit'>编辑</a>"
            "<a class='c-ops del' href='javascript:delDeal(%d)'>删除</a></div>"
            "<div class='c-text'>%s</div></td></tr>"
        ) % (seq_cell, _esc(r["name"]), _date_of(r), r["id"], r["id"], r["id"], _esc(text))

    PER = 10
    order = {t: i for i, t in enumerate(DEAL_TYPES)}
    rows_sorted = sorted(rows, key=lambda r: (order.get(r["deal_type"] or "", 99), -(r["amount"] or 0)))
    empty = "<tr><td class='mini' colspan='2'>暂无商机，可在下方新增，或用「粘贴导入 / 文件导入」批量录入。</td></tr>"

    # 全部视角：分页（每页 10 条，独立页码 page2）
    _total2 = len(rows_sorted)
    _pages2 = max(1, (_total2 + PER - 1) // PER)
    try:
        cur_page2 = max(1, int(request.args.get("page2", "1") or "1"))
    except (TypeError, ValueError):
        cur_page2 = 1
    cur_page2 = min(cur_page2, _pages2)
    cards = "".join(
        _row2(r) for r in rows_sorted[(cur_page2 - 1) * PER: cur_page2 * PER]) or empty

    # 模块视图：按日期倒序 + 分页（每页 10 条）+ 序号
    _mrows = sorted(rows, key=lambda r: (_date_of(r) or "0000-00-00"), reverse=True)
    _total = len(_mrows)
    _pages = max(1, (_total + PER - 1) // PER)
    cur_page = min(cur_page, _pages)
    _slice = _mrows[(cur_page - 1) * PER: cur_page * PER]
    cards_mod = "".join(_row1(r, (cur_page - 1) * PER + i + 1) for i, r in enumerate(_slice)) or empty

    def _q(**extra):
        prm = [("typ", typ_key), ("dfrom", dfrom), ("dto", dto), ("q", q),
               ("stage", stage), ("status", status)] + list(extra.items())
        return urllib.parse.urlencode([(k, v) for k, v in prm if v])

    def _pager(cp, pages, total, page_key):
        if pages <= 1:
            return ""
        prev = ("<a class='btn btn3 btn-sm' href='/deals?%s'>上一页</a>" % _q(**{page_key: cp - 1})
                if cp > 1 else
                "<span class='btn btn3 btn-sm' style='opacity:.4;pointer-events:none'>上一页</span>")
        nxt = ("<a class='btn btn3 btn-sm' href='/deals?%s'>下一页</a>" % _q(**{page_key: cp + 1})
               if cp < pages else
               "<span class='btn btn3 btn-sm' style='opacity:.4;pointer-events:none'>下一页</span>")
        return ("<div class='pager'>%s<span class='mini'>第 %d / %d 页 · 共 %d 条</span>%s</div>"
                % (prev, cp, pages, total, nxt))

    pager = _pager(cur_page, _pages, _total, "page")
    pager2 = _pager(cur_page2, _pages2, _total2, "page2")


    funnel = "".join(
        "<div class='kpi' style='padding:10px 14px'><div class='t'>%s</div>"
        "<div class='v' style='font-size:17px'>%d <small>个 / %s 万</small></div></div>"
        % (s, r["c"], format(r["s"], ",.0f")) for s, r in
        [(s, next((x for x in stat if x["stage"] == s), {"c": 0, "s": 0})) for s in STAGES])

    d = edit or {"id": "", "name": "", "client": "", "sector": "", "stage": "初步接洽",
                 "amount": "", "owner": "", "teammate": "", "status": "进行中", "note": ""}
    g = lambda k: (d[k] if (k in d.keys() and d[k] is not None) else "")
    opts = lambda vals, cur: "".join("<option %s>%s</option>" % ("selected" if v == cur else "", v) for v in vals)

    TYPE_ALIAS = {"非房固贷商机": "固贷商机"}
    TYPE_DESC = {
        "资本市场商机": "回购增持、可转债、定向增发、IPO 等资本市场业务机会",
        "增存商机": "一级市场股权融资款、募集资金归集等带来的存款沉淀机会",
        "并购商机": "发行股份购买资产、现金收购等并购交易，对应并购贷款与配套服务",
        "非房固贷商机": "发改委已备案的非房地产固定资产投资项目，对应项目贷款",
    }
    type_chips = (
        "<a class='tab %s' href='/deals'>全部商机<span>%d</span></a>"
        % ("on" if not typ else "", sum(tmap.values()))
        + "".join(
            "<a class='tab %s' href='/deals?typ=%s'>%s<span>%d</span></a>"
            % ("on" if typ == t else "", DEAL_TYPE_KEY.get(t, t), TYPE_ALIAS.get(t, t), tmap.get(t, 0))
            for t in DEAL_TYPES))

    funnel_title = ("%s · 阶段分布" % TYPE_ALIAS.get(typ, typ)) if typ else "全部商机 · 阶段分布"
    mod_head = ""          # 模块说明卡已按需求去掉

    banner = ""
    if msg.startswith("import:"):
        parts = msg.split(":")[1].split(",")
        if len(parts) >= 3:
            banner = ("<div class='ok'>导入完成：新增 %s 条，跳过 %s 条，失败 %s 条。</div>"
                      % (parts[0], parts[1], parts[2]))

    if typ:
        list_card = """
    <div class="card"><h2>商机内容 <span class="mini">共 %d 条</span>
        <span style="margin-left:auto;display:flex;gap:8px;align-items:center;flex-wrap:wrap">
          <form method="get" style="display:flex;gap:6px;align-items:center;margin:0">
            <input type="hidden" name="typ" value="%s">
            <input type="date" name="dfrom" id="dfM" value="%s" style="width:142px" title="开始日期"
              onchange="syncExport('expM','%s','dfM','dtM')" oninput="syncExport('expM','%s','dfM','dtM')">
            <span class="mini">至</span>
            <input type="date" name="dto" id="dtM" value="%s" style="width:142px" title="结束日期"
              onchange="syncExport('expM','%s','dfM','dtM')" oninput="syncExport('expM','%s','dfM','dtM')">
            <button class="btn3 btn-sm" style="padding:7px 12px">查询</button>
          </form>
          <a class="btn btn2 btn-sm" id="expM" href="/deals/export?typ=%s&dfrom=%s&dto=%s"
            onclick="syncExport('expM','%s','dfM','dtM')">⬇ 导出商机内容</a>
        </span></h2>
      <div class="tbl-wrap"><table class="one-col">
        <tr><th style="width:54px">序号</th><th>具体内容</th></tr>
        %s</table></div>
      %s
      <script>syncExport('expM','%s','dfM','dtM');</script></div>""" % (
            len(rows), typ_key, dfrom, typ_key, typ_key, dto, typ_key, typ_key,
            typ_key, dfrom, dto, typ_key, cards_mod, pager, typ_key)
    else:
        list_card = """
    <div class="card"><h2>商机台账 <span class="mini">共 %d 条</span>
        <span style="margin-left:auto;display:flex;gap:8px">
          <a class="btn btn2 btn-sm" href="javascript:openDeal()">＋ 添加商机</a>
          <a class="btn btn2 btn-sm" href="/deals/paste">📋 粘贴导入</a>
          <a class="btn btn3 btn-sm" href="/deals/import">⬆ 文件导入</a>
          <a class="btn btn3 btn-sm" href="/deals/template">⬇ 下载模板</a></span></h2>
      <form method="get" style="display:flex;gap:8px;margin-bottom:12px;flex-wrap:wrap">
        <input type="hidden" name="typ" value="%s">
        <input name="q" value="%s" placeholder="搜索 名称/客户/标的/行业/负责人" style="flex:1;min-width:180px">
        <select name="stage" style="width:120px"><option value="">全部阶段</option>%s</select>
        <select name="status" style="width:110px"><option value="">全部状态</option>%s</select>
        <button class="btn2">筛选</button><a class="btn3 btn" style="text-align:center" href="/deals">重置</a>
      </form>
      <div class="tbl-wrap"><table class="two-col">
        <tr><th>商机类别</th><th>具体内容</th></tr>
        %s</table></div>%s</div>""" % (len(rows), typ_key, q, opts(STAGES, stage or ""),
                                     opts(STATUSES, status or ""), cards, pager2)

    body = """
    {banner}
    <div class="card"><div class="tabs">{type_chips}</div>{mod_head}</div>
    {list_card}

    <div class="modal{modal_on}" id="dealModal" onclick="if(event.target===this)closeDeal()">
     <div class="modal-box">
      <div class="modal-h"><b>{ttl}</b><a class="modal-x" href="javascript:closeDeal()">×</a></div>
      <form method="post" action="/deals/save" class="frm">
      <input type="hidden" name="id" value="{id}">
      <div class="row">
        <div><label>商机类型</label><select name="deal_type"><option value="">未分类</option>{tps}</select></div>
        <div><label>细分类型</label><input name="sub_type" value="{sub_type}" placeholder="如 回购增持 / 一级市场股权融资款"></div>
        <div style="flex:2"><label>商机名称 *</label><input name="name" required value="{name}"></div>
        <div><label>日期</label><input name="deal_date" type="date" value="{deal_date}"></div>
      </div>
      <div class="row">
        <div><label>客户 / 主体 *</label><input name="client" value="{client}"></div>
        <div><label>交易对手 / 标的</label><input name="counterparty" value="{counterparty}"></div>
        <div><label>所属行业</label><input name="sector" value="{sector}"></div>
      </div>
      <div class="row">
        <div><label>金额（万元）</label><input name="amount" type="number" step="any" value="{amount}"></div>
        <div><label>金额上限（万元，区间用）</label><input name="amount2" type="number" step="any" value="{amount2}"></div>
        <div><label>交易 / 业务方式</label><input name="method" value="{method}" placeholder="回购 / 发行股份+现金 / B 轮"></div>
        <div><label>股权比例(%)</label><input name="ratio" type="number" step="any" value="{ratio}"></div>
      </div>
      <div class="row">
        <div><label>推进阶段</label><select name="stage">{stg2}</select></div>
        <div><label>状态</label><select name="status">{sts2}</select></div>
        <div><label>负责人</label><input name="owner" value="{owner}"></div>
        <div><label>共享成员</label><input name="teammate" value="{teammate}" placeholder="逗号分隔"></div>
      </div>
      <div class="row">
        <div><label>主办机构</label><input name="org" value="{org}" placeholder="如 东莞分行 / 天河智慧城支行"></div>
        <div><label>信息来源</label><input name="source" value="{source}" placeholder="分行报送 / 公告 / 发改委备案"></div>
        <div><label>公告 / 备案日期</label><input name="notice_date" type="date" value="{notice_date}"></div>
        <div><label>所在地</label><input name="location" value="{location}"></div>
      </div>
      <label>营销切入点 / 下一步</label><input name="entry" value="{entry}">
      <label>备注</label><textarea name="note" rows="2">{note}</textarea>
      <div style="margin-top:14px"><button>保存商机</button>
        <a class="btn btn3" href="javascript:closeDeal()" style="margin-left:8px">取消</a></div></form>
     </div></div>
    """.format(banner=banner, funnel=funnel, type_chips=type_chips, list_card=list_card,
               funnel_title=funnel_title, mod_head=mod_head,
               q=q, typ=typ_key, typ_key=typ_key, stg=opts(STAGES, stage or ""),
               sts=opts(STATUSES, status or ""), cards=cards,
               modal_on=" on" if d["id"] else "",
               ttl="编辑商机 #%s" % d["id"] if d["id"] else "新增商机",
               id=d["id"], name=d["name"], client=d["client"] or "", sector=d["sector"] or "",
               amount=d["amount"] if d["amount"] is not None else "",
               stg2=opts(STAGES, d["stage"]), owner=d["owner"] or "",
               teammate=d["teammate"] or "", sts2=opts(STATUSES, d["status"]), note=d["note"] or "",
               tps=opts(DEAL_TYPES, g("deal_type") or typ), sub_type=g("sub_type"),
               counterparty=g("counterparty"), method=g("method"),
               amount2=g("amount2"), ratio=g("ratio"),
               org=g("org"), source=g("source"), notice_date=g("notice_date"),
               location=g("location"), entry=g("entry"),
               deal_date=g("deal_date") or datetime.date.today().strftime("%Y-%m-%d"))
    return page(body, "deals")


@app.route("/deals/save", methods=["POST"])
def deals_save():
    f = request.form
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

    def val(k):
        return (f.get(k) or "").strip()

    def fnum(k):
        v = val(k)
        if not v:
            return None
        try:
            return float(v)
        except ValueError:
            return _num(v)

    data = {
        "name": val("name"),
        "client": val("client"),
        "sector": val("sector"),
        "stage": val("stage") or "初步接洽",
        "status": val("status") or "进行中",
        "amount": fnum("amount") or 0,
        "owner": val("owner"),
        "teammate": val("teammate"),
        "note": val("note"),
        "deal_type": val("deal_type"),
        "sub_type": val("sub_type"),
        "counterparty": val("counterparty"),
        "method": val("method"),
        "amount2": fnum("amount2"),
        "ratio": fnum("ratio"),
        "location": val("location"),
        "org": val("org"),
        "source": val("source"),
        "notice_date": _date(val("notice_date")) if val("notice_date") else "",
        "entry": val("entry"),
        "deal_date": _date(val("deal_date")) if val("deal_date")
        else datetime.date.today().strftime("%Y-%m-%d"),
    }
    conn = get_db()
    if f.get("id"):
        conn.execute("UPDATE deals SET %s, updated=? WHERE id=?"
                     % ",".join("%s=?" % k for k in DEAL_FIELDS),
                     [data[k] for k in DEAL_FIELDS] + [now, f["id"]])
    else:
        conn.execute("INSERT INTO deals (%s, created, updated) VALUES (%s,?,?)"
                     % (",".join(DEAL_FIELDS), ",".join(["?"] * len(DEAL_FIELDS))),
                     [data[k] for k in DEAL_FIELDS] + [now, now])
    conn.commit()
    conn.close()
    return redirect(url_for("deals"))


@app.route("/deals/view/<int:did>")
def deals_view(did):
    conn = get_db()
    r = conn.execute("SELECT * FROM deals WHERE id=?", (did,)).fetchone()
    conn.close()
    if not r:
        return page('<div class="err">该商机不存在或已删除。</div>', "deals")

    def esc(x):
        return str(x if x is not None else "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    t = r["deal_type"] or "未分类"
    stage_cls = {"进行中": "p1", "已成交": "p2", "已暂停": "p3", "已流失": "p4"}.get(r["status"], "p6")
    back = "/deals" + (("?typ=" + DEAL_TYPE_KEY.get(t, "")) if r["deal_type"] else "")
    text = esc(r["note"]) or "<span class='mini'>（暂无内容，点右侧「编辑」补充）</span>"

    body = (
        "<div class='card'>"
        "<div class='deal-head'>"
        "<span class='pill " + DEAL_TYPE_PILL.get(t, "p6") + "'>" + esc(t) + "</span>"
        "<span class='pill " + stage_cls + "'>" + esc(r["status"]) + "</span>"
        "<span class='mini'>" + esc(r["deal_date"] or (r["created"] or "")[:10]) + "</span>"
        "<span style='margin-left:auto;display:flex;gap:8px'>"
        "<a class='btn btn2 btn-sm' href='/deals/" + str(did) + "/edit'>编辑</a>"
        "<a class='btn btn3 btn-sm' href='javascript:delDeal(" + str(did) + ")'>删除</a>"
        "<a class='btn btn3 btn-sm' href='" + back + "'>返回列表</a>"
        "</span></div>"
        "<div class='detail-title'>" + esc(r["name"]) + "</div>"
        "<div class='detail-text'>" + text + "</div>"
        "</div>"
    )
    return page(body, "deals")


@app.route("/deals/<int:did>/edit", methods=["GET", "POST"])
def deals_edit(did):
    conn = get_db()
    r = conn.execute("SELECT * FROM deals WHERE id=?", (did,)).fetchone()
    if not r:
        conn.close()
        return page('<div class="err">该商机不存在或已删除。</div>', "deals")

    if request.method == "POST":
        note = (request.form.get("note") or "").strip()
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        # 名称为空时，用内容首句兜底
        name = r["name"]
        if not name and note:
            name = re.split(r"[。；;\n]", note)[0][:30]
        conn.execute("UPDATE deals SET note=?, name=?, updated=? WHERE id=?", (note, name, now, did))
        conn.commit()
        conn.close()
        return redirect("/deals/view/%d" % did)

    conn.close()
    t = r["deal_type"] or "未分类"
    body = (
        "<div class='card' style='max-width:920px'>"
        "<div class='deal-head'>"
        "<span class='pill " + DEAL_TYPE_PILL.get(t, "p6") + "'>" + esc_html(t) + "</span>"
        "<span class='mini'>" + esc_html(r["deal_date"] or (r["created"] or "")[:10]) + "</span>"
        "<span style='margin-left:auto' class='mini'>只编辑这段文字内容，分类和日期不变</span></div>"
        "<form method='post'>"
        "<textarea name='note' rows='16' class='edit-area' placeholder='在此填写商机内容'>"
        + esc_html(r["note"] or "") + "</textarea>"
        "<div style='margin-top:14px;display:flex;gap:8px'>"
        "<button>保存</button>"
        "<a class='btn btn3' href='/deals/view/" + str(did) + "'>取消</a>"
        "</div></form></div>"
    )
    return page(body, "deals")


@app.route("/deals/export")
def deals_export():
    typ_raw = request.args.get("typ", "").strip()
    typ = DEAL_KEY_TYPE.get(typ_raw, typ_raw)
    dfrom = (request.args.get("dfrom") or "").strip()
    dto = (request.args.get("dto") or "").strip()

    sql, args = "SELECT * FROM deals WHERE 1=1", []
    if typ:
        sql += " AND deal_type=?"
        args.append(typ)
    if dfrom:
        sql += " AND COALESCE(deal_date, substr(created,1,10)) >= ?"
        args.append(dfrom)
    if dto:
        sql += " AND COALESCE(deal_date, substr(created,1,10)) <= ?"
        args.append(dto)
    sql += " ORDER BY COALESCE(deal_date, substr(created,1,10)) DESC, id DESC"

    conn = get_db()
    rows = conn.execute(sql, args).fetchall()
    conn.close()

    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "商机内容"
    head = ["商机分类", "日期", "内容"]
    ws.append(head)
    fill = PatternFill("solid", fgColor="12224A")
    thin = Side(style="thin", color="D9DEE8")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for c in ws[1]:
        c.fill = fill
        c.font = Font(color="FFFFFF", bold=True, size=10.5)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border
    # 按分类排序（同类内保持 SQL 的日期倒序，sorted 是稳定排序）
    type_order = {t: i for i, t in enumerate(DEAL_TYPES)}
    rows = sorted(rows, key=lambda r: type_order.get(r["deal_type"] or "", 99))
    for r in rows:
        ws.append([
            r["deal_type"] or "未分类",
            r["deal_date"] or (r["created"] or "")[:10],
            (r["note"] or "").strip() or (r["name"] or ""),
        ])
    widths = [22, 14, 110]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(vertical="top", wrap_text=True)
        row[1].alignment = Alignment(vertical="top", horizontal="center")
        row[1].number_format = "@"
        row[0].alignment = Alignment(vertical="top", horizontal="center")
    ws.freeze_panes = "A2"

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    label = typ or "全部"
    fname = "商机内容_%s_%s.xlsx" % (label, datetime.datetime.now().strftime("%Y%m%d_%H%M"))
    return send_file(bio, as_attachment=True, download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/deals/delete/<int:did>")
def deals_delete(did):
    conn = get_db()
    conn.execute("DELETE FROM deals WHERE id=?", (did,))
    conn.commit(); conn.close()
    return redirect(url_for("deals"))


# ----------------------------------------------------------------------------
# 3. 跨表智能匹配（替代 VLOOKUP）
# ----------------------------------------------------------------------------
MATCHER_FORM = """
<div class="card"><h2>跨表智能匹配 · 替代手工 VLOOKUP</h2>
<form method="post" action="/matcher/upload" enctype="multipart/form-data" class="frm">
  <div class="row">
    <div><label>表 A（如：业务系统导出名单）</label><input type="file" name="fa" required accept=".xlsx,.csv"></div>
    <div><label>表 B（如：台账 / 报表底稿）</label><input type="file" name="fb" required accept=".xlsx,.csv"></div>
  </div>
  <div style="margin-top:14px"><button>上传并选择关键列</button></div>
</form>
<p class="mini" style="margin-top:10px">支持 xlsx / csv；关键列建议用「统一社会信用代码」，企业名称会自动做全角半角、空格、括号归一化后再精确匹配，未命中的用相似度模糊匹配兜底。</p></div>
{hist}
"""


@app.route("/matcher")
def matcher():
    conn = get_db()
    runs = conn.execute("SELECT * FROM match_runs ORDER BY id DESC LIMIT 8").fetchall()
    conn.close()
    hist = "<div class='card'><h2>历史任务</h2><table><tr><th>时间</th><th>文件</th><th>结果</th><th></th></tr>" + \
        "".join("<tr><td>%s</td><td class='mini'>%s ↔ %s</td>"
                "<td><span class='pill p2'>精确 %d</span> <span class='pill p3'>模糊 %d</span> "
                "<span class='pill p4'>未匹配 %d</span></td>"
                "<td><a href='/matcher/result/%d'>下载</a></td></tr>"
                % (r["ts"], r["file_a"], r["file_b"], r["exact"], r["fuzzy"], r["unmatched"], r["id"])
                for r in runs) + "</table></div>" or ""
    return page(MATCHER_FORM.format(hist=hist), "matcher")


@app.route("/matcher/upload", methods=["POST"])
def matcher_upload():
    fa, fb = request.files.get("fa"), request.files.get("fb")
    if not fa or not fb:
        return page("<div class='err'>请同时选择两个文件</div>" + MATCHER_FORM.format(hist=""), "matcher")
    try:
        pa = os.path.join(UPLOAD_DIR, secure_filename(fa.filename) or "a.xlsx")
        pb = os.path.join(UPLOAD_DIR, secure_filename(fb.filename) or "b.xlsx")
        fa.save(pa); fb.save(pb)
        ha, _ = read_table(pa)
        hb, _ = read_table(pb)
    except Exception as e:
        return page("<div class='err'>读取失败：%s</div>" % e, "matcher")
    conn = get_db()
    cur = conn.execute("""INSERT INTO match_jobs (ts,file_a,file_b,headers_a,headers_b) VALUES (?,?,?,?,?)""",
                       (datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
                        os.path.basename(pa), os.path.basename(pb),
                        json.dumps(ha, ensure_ascii=False), json.dumps(hb, ensure_ascii=False)))
    jid = cur.lastrowid
    conn.commit(); conn.close()
    return redirect("/matcher/job/%d" % jid)


@app.route("/matcher/job/<int:jid>")
def matcher_job(jid):
    conn = get_db()
    j = conn.execute("SELECT * FROM match_jobs WHERE id=?", (jid,)).fetchone()
    conn.close()
    if not j:
        return redirect("/matcher")
    ha, hb = json.loads(j["headers_a"]), json.loads(j["headers_b"])
    opt = lambda hs: "".join("<option>%s</option>" % h for h in hs)
    body = """
    <div class="card"><h2>第 2 步：选择关键列</h2>
    <p class="mini">表 A：<b>{fa}</b>（{na} 列） · 表 B：<b>{fb}</b>（{nb} 列）</p>
    <form method="post" action="/matcher/run/{jid}" class="frm">
      <div class="row" style="margin-top:10px">
        <div><label>表 A 关键列（用于匹配的字段）</label><select name="ca">{opta}</select></div>
        <div><label>表 B 关键列</label><select name="cb">{optb}</select></div>
        <div style="max-width:160px"><label>模糊匹配阈值（0.80~0.99）</label>
          <input name="th" value="0.85" type="number" step="0.01" min="0.5" max="1"></div>
      </div>
      <div style="margin-top:14px"><button>开始匹配</button>
      <a class="btn btn3" style="margin-left:8px" href="/matcher">返回重选文件</a></div></form></div>
    """.format(fa=j["file_a"], fb=j["file_b"], na=len(ha), nb=len(hb),
               jid=jid, opta=opt(ha), optb=opt(hb))
    return page(body, "matcher")


@app.route("/matcher/run/<int:jid>", methods=["POST"])
def matcher_run(jid):
    conn = get_db()
    j = conn.execute("SELECT * FROM match_jobs WHERE id=?", (jid,)).fetchone()
    conn.close()
    if not j:
        return redirect("/matcher")
    ca, cb = request.form["ca"], request.form["cb"]
    th = float(request.form.get("th") or 0.85)
    ia = int(ca.split("·")[0]) - 1 if "·" in ca else ha_idx(ca)
    ib = int(cb.split("·")[0]) - 1 if "·" in cb else ha_idx(cb)
    _, rows_a = read_table(os.path.join(UPLOAD_DIR, j["file_a"]))
    _, rows_b = read_table(os.path.join(UPLOAD_DIR, j["file_b"]))
    ia = min(ia, max((len(r) for r in rows_a), default=0) - 1) if rows_a else 0
    ib = min(ib, max((len(r) for r in rows_b), default=0) - 1) if rows_b else 0

    # 索引表 B：归一化 key → 行号（含完全重复键取首条）
    index = {}
    for y, r in enumerate(rows_b):
        k = norm_key(r[ib]) if ib < len(r) else ""
        if k and k not in index:
            index[k] = y
    keys_b = list(index.keys())

    results, n_exact, n_fuzzy = [], 0, 0
    for x, ra in enumerate(rows_a):
        ka = norm_key(ra[ia]) if ia < len(ra) else ""
        if not ka:
            results.append((x, "", 0.0, None, "空关键值")); continue
        if ka in index:
            y = index[ka]
            results.append((x, rows_b[y][ib], 1.0, y, "精确匹配")); n_exact += 1
        else:
            best, bs = None, 0.0
            for k in keys_b:
                if abs(len(k) - len(ka)) > 6:
                    continue
                s = difflib.SequenceMatcher(None, ka, k).ratio()
                if s > bs:
                    bs, best = s, k
            if best is not None and bs >= th:
                y = index[best]
                results.append((x, rows_b[y][ib], round(bs, 2), y, "模糊匹配")); n_fuzzy += 1
            else:
                results.append((x, "", round(bs if best else 0, 2), None, "未匹配"))

    # 输出 Excel：A 全部列 + 匹配状态 + B 的前 5 列
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "匹配结果"
    ha, _ = read_table(os.path.join(UPLOAD_DIR, j["file_a"]))
    hb, _ = read_table(os.path.join(UPLOAD_DIR, j["file_b"]))
    ws.append(ha + ["匹配状态", "匹配到的B关键值", "相似度"] + hb[:5])
    for x, kb, sc, y, status in results:
        ra = rows_a[x] + [status, kb, sc] + (rows_b[y][:5] if y is not None else [""] * 5)
        ws.append(ra)
    fname = "match_result_%s.xlsx" % datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    fpath = os.path.join(RESULT_DIR, fname)
    wb.save(fpath)

    n_un = len([r for r in results if r[4] == "未匹配"])
    conn = get_db()
    cur = conn.execute("""INSERT INTO match_runs (ts,file_a,file_b,col_a,col_b,total,exact,fuzzy,unmatched,result_path)
                          VALUES (?,?,?,?,?,?,?,?,?,?)""",
                       (datetime.datetime.now().strftime("%Y-%m-%d %H:%M"), j["file_a"], j["file_b"],
                        ca, cb, len(results), n_exact, n_fuzzy, n_un, fname))
    rid = cur.lastrowid
    conn.commit(); conn.close()

    preview = "".join(
        "<tr><td class='mini'>%s</td><td>%s</td><td>%s</td><td>%s</td>"
        "<td><span class='pill %s'>%s</span></td></tr>"
        % (rows_a[r[0]][ia] or "(空)", r[1] or "—", r[2] if r[2] else "—",
           (rows_b[r[3]][ib] if r[3] is not None and ib < len(rows_b[r[3]]) else "—") if r[4] != "未匹配" else "—",
           {"精确匹配": "p2", "模糊匹配": "p3", "未匹配": "p4", "空关键值": "p6"}[r[4]], r[4])
        for r in results[:60])
    body = """
    <div class="card"><h2>匹配完成</h2>
    <div class="kpis" style="grid-template-columns:repeat(4,1fr)">
      <div class="kpi"><div class="t">表 A 总行数</div><div class="v">{total}</div></div>
      <div class="kpi"><div class="t">精确匹配</div><div class="v" style="color:#1a7f4d">{ex}</div></div>
      <div class="kpi"><div class="t">模糊匹配 ≥{th}</div><div class="v" style="color:#b97a12">{fz}</div></div>
      <div class="kpi"><div class="t">未匹配</div><div class="v" style="color:#d0342c">{un}</div></div>
    </div>
    <a class="btn" href="/matcher/result/{rid}">⬇ 下载完整结果（Excel）</a>
    <a class="btn btn3" style="margin-left:8px" href="/matcher">再来一次</a></div>
    <div class="card"><h2>预览（前 60 行）</h2><table>
    <tr><th>A 关键值</th><th>归一化后命中</th><th>相似度</th><th>B 关键值</th><th>状态</th></tr>{pv}</table>
    <p class="mini" style="margin-top:8px">说明：模糊匹配结果建议人工复核（尤其是 0.85~0.92 区间）；未匹配项可能是新户或关键字段缺失，可导出后单独处理。</p></div>
    """.format(total=len(results), ex=n_exact, fz=n_fuzzy, un=n_un, th=th, rid=rid, pv=preview)
    return page(body, "matcher")


def ha_idx(colname):
    return 0


@app.route("/matcher/result/<int:rid>")
def matcher_result(rid):
    conn = get_db()
    r = conn.execute("SELECT result_path FROM match_runs WHERE id=?", (rid,)).fetchone()
    conn.close()
    if not r or not os.path.isfile(os.path.join(RESULT_DIR, r["result_path"])):
        return redirect("/matcher")
    return send_file(os.path.join(RESULT_DIR, r["result_path"]), as_attachment=True,
                     download_name=r["result_path"])


# ----------------------------------------------------------------------------
# 3.5 科技金融（拓户名单 + S70 匹配）
# ----------------------------------------------------------------------------
@app.route("/tech")
def tech():
    q = request.args.get("q", "").strip()
    status = request.args.get("status", "")
    edit = None
    if request.args.get("edit"):
        conn = get_db()
        edit = conn.execute("SELECT * FROM expansion_clients WHERE id=?",
                             (request.args["edit"],)).fetchone()
        conn.close()
    sql, args = "SELECT * FROM expansion_clients WHERE 1=1", []
    if q:
        sql += " AND (name LIKE ? OR sector LIKE ? OR region LIKE ? OR owner LIKE ?)"
        args += ["%%%s%%" % q] * 4
    if status:
        sql += " AND status=?"; args.append(status)
    sql += " ORDER BY updated DESC"

    _today = datetime.date.today()
    s70_date = request.args.get("s70date") or ("%d年%d月%d日" % (_today.year, _today.month, _today.day))
    year_start = float(get_setting("s70_year_start", "0") or 0)
    ys_normal = float(get_setting("s70_year_start_normal", "0") or 0)
    ys_fuf = float(get_setting("s70_year_start_fuf", "0") or 0)
    ys_disc = float(get_setting("s70_year_start_disc", "0") or 0)
    q_date = request.args.get("q_date", "").strip()          # 日期组件传 ISO：2026-10-31
    q_date_cn = iso_date_to_cn(q_date)                        # 库中格式：2026年10月31日
    q_cust = request.args.get("q_cust", "").strip()
    q_unit = request.args.get("q_unit", "").strip()
    q_credit = request.args.get("q_credit", "").strip()
    try:
        cur_page = max(1, int(request.args.get("page", "1")))
    except ValueError:
        cur_page = 1
    PER = 10

    cond, cargs = "1=1", []
    if q_date_cn:
        cond += " AND data_date = ?"; cargs.append(q_date_cn)
    if q_cust:
        cond += " AND cust_name LIKE ?"; cargs.append("%%%s%%" % q_cust)
    if q_unit:
        cond += " AND unit LIKE ?"; cargs.append("%%%s%%" % q_unit)
    if q_credit:
        cond += " AND credit_type=?"; cargs.append(q_credit)

    conn = get_db()
    rows = conn.execute(sql, args).fetchall()
    s70_dates = [r["data_date"] for r in conn.execute(
        "SELECT DISTINCT data_date FROM s70_loans ORDER BY data_date DESC")]

    def s70_sum(credit=None):
        qq = "SELECT COALESCE(SUM(amount_yi),0) s FROM s70_loans WHERE data_date=?"
        aa = [s70_date]
        if credit:
            qq += " AND credit_type=?"; aa.append(credit)
        return conn.execute(qq, aa).fetchone()["s"]

    s70_total = s70_sum()
    s70_normal = s70_sum("一般性贷款")
    s70_fuf = s70_sum("福费廷")
    s70_disc = s70_sum("贴现")
    s70_total_cnt = conn.execute("SELECT COUNT(*) c FROM s70_loans WHERE " + cond,
                                 cargs).fetchone()["c"]
    pages = max(1, (s70_total_cnt + PER - 1) // PER)
    cur_page = min(cur_page, pages)
    s70_rows = conn.execute(
        "SELECT * FROM s70_loans WHERE %s ORDER BY data_date DESC, id DESC LIMIT ? OFFSET ?" % cond,
        cargs + [PER, (cur_page - 1) * PER]).fetchall()
    s70_edit = None
    if request.args.get("sid"):
        s70_edit = conn.execute("SELECT * FROM s70_loans WHERE id=?",
                                (request.args["sid"],)).fetchone()
    s70_detail = None
    if request.args.get("sdetail"):
        s70_detail = conn.execute("SELECT * FROM s70_loans WHERE id=?",
                                  (request.args["sdetail"],)).fetchone()
    conn.close()
    s70_diff = s70_total - year_start
    diff_color = "#d0342c" if s70_diff >= 0 else "#1a7f4d"
    # 各分类「较年初」：逻辑与 S70 总余额一致（当前余额 - 该分类年初余额）
    d_normal, d_fuf, d_disc = s70_normal - ys_normal, s70_fuf - ys_fuf, s70_disc - ys_disc
    c_normal = "#d0342c" if d_normal >= 0 else "#1a7f4d"
    c_fuf = "#d0342c" if d_fuf >= 0 else "#1a7f4d"
    c_disc = "#d0342c" if d_disc >= 0 else "#1a7f4d"

    def s70_q(**kw):
        # 始终带上 s70date（看板数据日期），保证查询/分页不改变看板统计口径
        p = {"s70date": s70_date, "q_date": q_date, "q_cust": q_cust,
             "q_unit": q_unit, "q_credit": q_credit}
        p.update(kw)
        return urllib.parse.urlencode({k: v for k, v in p.items() if v})

    s70_export_url = "/tech/s70/export" + (("?" + s70_q()) if s70_q() else "")
    s70_prev = "/tech" + (("?" + s70_q(page=cur_page - 1)) if cur_page > 1 else "")
    s70_next = "/tech" + (("?" + s70_q(page=cur_page + 1)) if cur_page < pages else "")

    msg = (request.args.get("msg") or "").strip()
    banner = ""
    if msg.startswith("s70import:"):
        parts = msg.split(":")[1].split(",")
        banner = "<div class='ok'>S70 导入完成：新增 %s 条。</div>" % parts[0]

    trs = "".join(
        "<tr><td><b>%s</b><div class='mini'>%s · %s</div></td>"
        "<td>%s</td><td><span class='pill %s'>%s</span></td><td>%s</td>"
        "<td class='mini'>%s</td><td><a href='/tech?edit=%d&%s'>编辑</a> "
        "<a href='/tech/delete/%d' onclick=\"return confirm('确认删除？')\">删除</a></td></tr>"
        % (r["name"], r["sector"] or "", r["region"] or "",
           r["tag"] or "—",
           {"未接触": "p6", "已接触": "p1", "已开户": "p3", "已授信": "p5", "已投放": "p2"}.get(r["status"], "p6"),
           r["status"], r["owner"] or "—", r["source"] or "",
           r["id"], request.query_string.decode(), r["id"])
        for r in rows) or "<tr><td class='mini' colspan=7>暂无拓户名单，点击右上角「＋ 新增拓户」开始</td></tr>"

    show_cols = [c for c in S70_COLS if c not in S70_HIDDEN_COLS]
    show_headers = [h for h, c in zip(S70_HEADERS, S70_COLS) if c not in S70_HIDDEN_COLS]
    s70_ths = "".join("<th>%s</th>" % h for h in show_headers)
    s70_trs = "".join(
        "<tr><td><input type='checkbox' name='ids' value='%d' form='s70BatchForm' "
        "class='s70-chk'></td>%s<td style='white-space:nowrap'><a href='/tech?sdetail=%d'>详情</a> "
        "<a href='/tech?sid=%d&%s' style='margin-left:6px'>编辑</a> "
        "<a href='/tech/s70/delete/%d' onclick=\"return confirm('确认删除？')\" "
        "style='color:#d0342c;margin-left:6px'>删除</a></td></tr>"
        % (r["id"], "".join("<td>%s</td>" % (
            s70_mask(r[c]) if c in S70_MASK_COLS
            else (s70_num(r[c]) if c in S70_REAL_COLS else (r[c] or "")))
            for c in show_cols),
           r["id"], r["id"], s70_q(), r["id"])
        for r in s70_rows) or ("<tr><td class='mini' colspan=%d>暂无 S70 贷款数据，点击右上角「＋ 新增」或导入</td></tr>"
                               % (len(show_cols) + 2))

    s70_detail_html = ""
    if s70_detail:
        items = "".join(
            "<div style='padding:4px 0;border-bottom:1px dashed #eef1f6'>"
            "<span class='mini' style='display:inline-block;width:170px;color:#8b95ab'>%s</span>"
            "<b>%s</b></div>"
            % (h, (s70_mask(s70_detail[c]) if c in S70_MASK_COLS
                   else (s70_num(s70_detail[c]) if c in S70_REAL_COLS
                         else (s70_detail[c] or "—"))))
            for h, c in zip(S70_HEADERS, S70_COLS))
        s70_detail_html = """
    <div class="card"><h2>贷款详情 #%s
      <span style="margin-left:auto"><a class="btn btn3 btn-sm" href="/tech">关闭</a></span></h2>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:0 28px;font-size:12.5px">%s</div>
    </div>""" % (s70_detail["id"], items)

    s70_pager = ""
    if pages > 1:
        s70_pager = ("<div class='pager'><a href='%s'>上一页</a>"
                     "<span>第 %d / %d 页 · 共 %d 条</span><a href='%s'>下一页</a></div>"
                     % (s70_prev, cur_page, pages, s70_total_cnt, s70_next))
    s70_pager = s70_pager or ("<div class='pager'><span>共 %d 条</span></div>" % s70_total_cnt)

    credit_opts = "".join("<option %s>%s</option>" % ("selected" if v == q_credit else "", v)
                          for v in ("一般性贷款", "福费廷", "贴现"))
    # 查询「数据日期」用日期组件（可选任意一天），不允许手填文本
    q_date_iso = cn_date_to_iso(q_date) or q_date

    e = s70_edit

    def ev(k):
        if not e:
            return ""
        v = e[k]
        if v is None:
            return ""
        if k in S70_REAL_COLS:
            return s70_num(v)
        return v

    s70_form_display = "block" if s70_edit else "none"
    s70_add_btn = ("" if s70_edit else
                   "<a class='btn btn2 btn-sm' id='btnAddS70' "
                   "href=\"javascript:showBox('s70FormCard','btnAddS70')\">＋ 新增</a>")
    s70_cancel_btn = ("<a class='btn btn3' href='/tech' style='margin-left:8px'>取消</a>" if s70_edit else
                      "<a class='btn btn3' href=\"javascript:hideBox('s70FormCard','btnAddS70')\" "
                      "style='margin-left:8px'>取消</a>")

    h2c_rev = dict(zip(S70_COLS, S70_HEADERS))

    def _field(c):
        v = str(ev(c) if ev(c) != "" else "")
        v = v.replace("&", "&amp;").replace('"', "&quot;").replace("<", "&lt;")
        num = "type='number' step='any' " if c in S70_REAL_COLS else ""
        req = "required " if c == "data_date" else ""
        return ("<div><label>%s</label><input name='%s' %s%svalue=\"%s\"></div>"
                % (h2c_rev[c], c, num, req, v))

    s70_groups_html = "".join(
        "<div class='mini' style='margin:14px 0 6px;font-weight:600;color:#5b6478'>%s</div>"
        "<div class='row'>%s</div>" % (gname, "".join(_field(c) for c in cols))
        for gname, cols in S70_FORM_GROUPS)

    if s70_dates and s70_date not in s70_dates:
        s70_dates.insert(0, s70_date)
    date_opts = "".join("<option %s>%s</option>" % ("selected" if v == s70_date else "", v)
                        for v in (s70_dates or [s70_date]))

    d = edit or {"id": "", "name": "", "sector": "", "region": "", "tag": "",
                 "source": "", "status": "未接触", "owner": "", "note": ""}
    opts = lambda vals, cur: "".join("<option %s>%s</option>" % ("selected" if v == cur else "", v) for v in vals)
    form_display = "block" if edit else "none"
    add_btn = ("" if edit else
               "<span style='margin-left:auto'><a class='btn btn2 btn-sm' id='btnAddTech' "
               "href=\"javascript:showBox('techFormCard','btnAddTech')\">＋ 新增拓户</a></span>")
    cancel_btn = ("<a class='btn btn3' href='/tech' style='margin-left:8px'>取消</a>" if edit else
                  "<a class='btn btn3' href=\"javascript:hideBox('techFormCard','btnAddTech')\" "
                  "style='margin-left:8px'>取消</a>")
    body = """
    {banner}
    <div class="card"><h2>S70 贷款总览</h2>
      <form method="post" action="/tech/s70/yearstart" style="display:flex;gap:14px;align-items:flex-end;flex-wrap:wrap">
        <div><label>年初余额（亿元）· S70 总</label>
          <input name="year_start" type="number" step="any" value="{year_start_txt}" style="width:112px"></div>
        <div><label>一般性贷款年初</label>
          <input name="year_start_normal" type="number" step="any" value="{ys_normal_txt}" style="width:112px"></div>
        <div><label>福费廷年初</label>
          <input name="year_start_fuf" type="number" step="any" value="{ys_fuf_txt}" style="width:112px"></div>
        <div><label>贴现年初</label>
          <input name="year_start_disc" type="number" step="any" value="{ys_disc_txt}" style="width:112px"></div>
        <div><label>数据日期</label><select name="s70date" style="width:170px">{date_opts}</select></div>
        <div style="align-self:end"><button>保存 / 刷新</button></div>
      </form>
      <div class="row" style="margin-top:14px">
        <div class="kpi" style="flex:1"><div class="t">S70 总余额（亿元）</div>
          <div class="v" style="font-size:20px">{s70_total:.4f}</div>
          <div class="mini">较年初：<b style="color:{diff_color}">{s70_diff:+.4f}</b></div></div>
        <div class="kpi" style="flex:1"><div class="t">一般性贷款总余额（亿元）</div>
          <div class="v" style="font-size:20px">{s70_normal:.4f}</div>
          <div class="mini">较年初：<b style="color:{c_normal}">{d_normal:+.4f}</b></div></div>
        <div class="kpi" style="flex:1"><div class="t">福费廷总余额（亿元）</div>
          <div class="v" style="font-size:20px">{s70_fuf:.4f}</div>
          <div class="mini">较年初：<b style="color:{c_fuf}">{d_fuf:+.4f}</b></div></div>
        <div class="kpi" style="flex:1"><div class="t">贴现总余额（亿元）</div>
          <div class="v" style="font-size:20px">{s70_disc:.4f}</div>
          <div class="mini">较年初：<b style="color:{c_disc}">{d_disc:+.4f}</b></div></div>
      </div>
    </div>
    {s70_detail_html}
    <div class="card"><h2>S70 贷款明细</h2>
      <div style="display:flex;gap:8px;margin-bottom:10px;justify-content:flex-end;align-items:center;flex-wrap:wrap">
        <button type="button" class="btn3 btn-sm" onclick="s70BatchDel()"
          style="color:#d0342c;border-color:#f0c4c0">🗑 批量删除</button>
        <a class="btn btn3 btn-sm" href="{s70_export_url}">⬇ 导出</a>
        <form method="post" action="/tech/s70/import" enctype="multipart/form-data"
          style="display:flex;gap:6px;align-items:center;margin:0">
          <input type="file" name="file" accept=".xlsx,.xlsm" required style="width:170px;font-size:11.5px">
          <button class="btn2 btn-sm">⬆ 导入</button>
        </form>
        {s70_add_btn}
      </div>
      <form method="get" style="display:flex;gap:8px;margin-bottom:12px;flex-wrap:wrap">
        <input type="hidden" name="s70date" value="{s70_date}">
        <div style="display:flex;gap:4px;align-items:center">
          <input type="date" name="q_date" id="qDateInput" value="{q_date_iso}"
            style="width:152px" title="选择数据日期">
          <button type="button" class="btn3 btn-sm" onclick="clearQDate()" title="清空日期"
            style="padding:7px 10px;line-height:1">✕</button>
        </div>
        <input name="q_cust" value="{q_cust}" placeholder="客户名称" style="flex:1;min-width:140px">
        <input name="q_unit" value="{q_unit}" placeholder="经营单位" style="width:120px">
        <select name="q_credit" style="width:135px">
          <option value="">全部授信品种</option>{credit_opts}</select>
        <button class="btn2">查询</button><a class="btn3 btn" style="text-align:center" href="/tech">重置</a>
      </form>
      <div class="tbl-wrap"><table style="white-space:nowrap">
        <tr><th style="width:34px"><input type="checkbox" id="s70ChkAll"
          onclick="s70ToggleAll(this)" title="全选本页"></th>{s70_ths}<th>操作</th></tr>
        {s70_trs}</table></div>
      {s70_pager}
      <form id="s70BatchForm" method="post" action="/tech/s70/batch_delete"></form>
      <script>
      function clearQDate(){{
        var el=document.getElementById('qDateInput');
        if(!el)return;
        el.value='';
        el.form.submit();
      }}
      function s70ToggleAll(el){{
        document.querySelectorAll("input[name='ids']").forEach(function(c){{c.checked=el.checked}});
      }}
      function s70BatchDel(){{
        var boxes=document.querySelectorAll("input[name='ids']:checked");
        if(!boxes.length){{alert('请先勾选要删除的记录');return}}
        if(confirm('确认删除选中的 '+boxes.length+' 条记录？删除后不可恢复。')){{
          document.getElementById('s70BatchForm').submit();
        }}
      }}
      </script>
    </div>
    <div class="card" id="s70FormCard" style="display:{s70_form_display}"><h2>{s70_form_ttl}</h2>
    <form method="post" action="/tech/s70/save" class="frm">
      <input type="hidden" name="id" value="{s70_id}">
      {s70_groups_html}
      <div style="margin-top:14px"><button>保存</button>{s70_cancel_btn}</div></form></div>
    <div class="card"><h2>S70 / 各类报表匹配</h2>
    <p class="mini" style="line-height:2">科技金融室每周的 S70、拓户、各类需要 VLOOKUP 的报表，统一用<b>跨表匹配</b>完成：
    上传「系统导出名单」和「台账底稿」两张表，选关键列（建议统一社会信用代码），自动完成<b>精确匹配（全角半角/空格/括号自动归一化）+ 模糊匹配兜底</b>，
    模糊命中项进入人工复核，结果一键导出 Excel。
    <br><a class="btn" style="margin-top:8px" href="/matcher">🔗 打开跨表匹配</a></p></div>
    """.format(banner=banner, year_start=year_start, date_opts=date_opts,
               year_start_txt=s70_num(year_start) or "0",
               s70_date=s70_date,
               s70_total=s70_total, s70_diff=s70_diff, diff_color=diff_color,
               ys_normal_txt=s70_num(ys_normal) or "0",
               ys_fuf_txt=s70_num(ys_fuf) or "0",
               ys_disc_txt=s70_num(ys_disc) or "0",
               d_normal=d_normal, d_fuf=d_fuf, d_disc=d_disc,
               c_normal=c_normal, c_fuf=c_fuf, c_disc=c_disc,
               s70_normal=s70_normal, s70_fuf=s70_fuf, s70_disc=s70_disc,
               s70_ths=s70_ths, s70_trs=s70_trs, s70_pager=s70_pager,
               s70_detail_html=s70_detail_html,
               s70_export_url=s70_export_url, s70_add_btn=s70_add_btn,
               q_date=q_date, q_cust=q_cust, q_unit=q_unit, credit_opts=credit_opts,
               q_date_iso=q_date_iso,
               s70_form_display=s70_form_display, s70_groups_html=s70_groups_html,
               s70_form_ttl=("编辑 S70 贷款 #%s" % e["id"]) if e else "新增 S70 贷款",
               s70_id=e["id"] if e else "",
               e_data_date=ev("data_date"), e_loan_no=ev("loan_no"),
               e_cust_no=ev("cust_no"), e_credit_type=ev("credit_type"),
               e_cust_name=ev("cust_name"), e_unit=ev("unit"),
               e_sub_branch=ev("sub_branch"), e_mgr_name=ev("mgr_name"),
               e_mgr_id=ev("mgr_id"), e_acct_prod=ev("acct_prod"),
               e_amount_yi=ev("amount_yi"), e_balance_rmb=ev("balance_rmb"),
               e_loan_start=ev("loan_start"), e_loan_end=ev("loan_end"),
               s70_cancel_btn=s70_cancel_btn,
               q=q, sts=opts([""] + EXP_STATUS, status),
               trs=trs, ttl="编辑拓户名单 #%s" % d["id"] if d["id"] else "新增拓户名单",
               id=d["id"], name=d["name"], sector=d["sector"] or "", region=d["region"] or "",
               tag=d["tag"] or "", source=d["source"] or "", owner=d["owner"] or "",
               sts2=opts(EXP_STATUS, d["status"]), note=d["note"] or "",
               form_display=form_display, add_btn=add_btn, cancel_btn=cancel_btn)
    return page(body, "tech")


@app.route("/tech/save", methods=["POST"])
def tech_save():
    f = request.form
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    conn = get_db()
    if f.get("id"):
        conn.execute("""UPDATE expansion_clients SET name=?,sector=?,region=?,tag=?,source=?,
                        status=?,owner=?,note=?,updated=? WHERE id=?""",
                     (f["name"], f.get("sector"), f.get("region"), f.get("tag"), f.get("source"),
                      f.get("status"), f.get("owner"), f.get("note"), now, f["id"]))
    else:
        conn.execute("""INSERT INTO expansion_clients (name,sector,region,tag,source,status,owner,note,created,updated)
                        VALUES (?,?,?,?,?,?,?,?,?,?)""",
                     (f["name"], f.get("sector"), f.get("region"), f.get("tag"), f.get("source"),
                      f.get("status"), f.get("owner"), f.get("note"), now, now))
    conn.commit(); conn.close()
    return redirect(url_for("tech"))


@app.route("/tech/delete/<int:did>")
def tech_delete(did):
    conn = get_db()
    conn.execute("DELETE FROM expansion_clients WHERE id=?", (did,))
    conn.commit(); conn.close()
    return redirect(url_for("tech"))


@app.route("/tech/s70/import", methods=["POST"])
def tech_s70_import():
    f = request.files.get("file")
    if not f or not f.filename:
        return redirect("/tech")
    try:
        wb = openpyxl.load_workbook(f, data_only=True)
    except Exception:
        return page('<div class="err">无法读取文件，请上传 .xlsx 格式的 S70 贷款模版。</div>'
                    '<a class="btn" href="/tech">返回</a>', "tech")
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if len(rows) < 2:
        return redirect("/tech")
    headers = [str(h).strip() if h is not None else "" for h in rows[0]]
    h2c = dict(zip(S70_HEADERS, S70_COLS))
    idx = {}
    seen = {}
    for i, h in enumerate(headers):
        if not h:
            continue
        if h == "贷后管理人员行员号" and "贷后管理人员行员号" in seen:
            idx[S70_COLS.index("mgr_id2")] = i
        elif h in h2c:
            idx[S70_COLS.index(h2c[h])] = i
        seen[h] = i
    j_date = S70_COLS.index("data_date")
    j_no = S70_COLS.index("loan_no")
    now = _now()
    n_new = n_skip = 0
    conn = get_db()
    for row in rows[1:]:
        if not any(c is not None for c in row):
            continue

        def gv(j, row=row):
            i = idx.get(j)
            if i is None or i >= len(row) or row[i] is None or row[i] == "":
                return ""
            v = row[i]
            col = S70_COLS[j]
            if col in S70_DATE_COLS:
                # 兼容 Excel 序列号(46326)、datetime、中文、ISO 各种写法
                return norm_date(v, cn=(col == "data_date"))
            if isinstance(v, datetime.datetime):
                return v.strftime("%Y-%m-%d")
            if isinstance(v, float) and v.is_integer():
                return str(int(v))
            return str(v).strip()

        def gn(j, row=row):
            i = idx.get(j)
            if i is None or i >= len(row) or row[i] in (None, ""):
                return None
            try:
                return float(row[i])
            except (TypeError, ValueError):
                return None

        vals = [gn(S70_COLS.index(c)) if c in S70_REAL_COLS else gv(S70_COLS.index(c))
                for c in S70_COLS]
        conn.execute("INSERT INTO s70_loans (%s, created) VALUES (%s)"
                     % (",".join(S70_COLS), ",".join(["?"] * (len(S70_COLS) + 1))),
                     vals + [now])
        n_new += 1
    conn.commit()
    conn.close()
    return redirect("/tech?msg=s70import:%d,0" % n_new)


@app.route("/tech/s70/yearstart", methods=["POST"])
def tech_s70_yearstart():
    """顶部「保存 / 刷新」一次保存：S70 总 + 一般性 + 福费廷 + 贴现 四个年初余额"""
    def _f(name):
        try:
            return float((request.form.get(name) or "").strip() or 0)
        except ValueError:
            return 0.0

    for form_name, setting_key in (("year_start", "s70_year_start"),
                                   ("year_start_normal", "s70_year_start_normal"),
                                   ("year_start_fuf", "s70_year_start_fuf"),
                                   ("year_start_disc", "s70_year_start_disc")):
        if form_name in request.form:          # 只更新提交上来的字段
            set_setting(setting_key, str(_f(form_name)))
    s70date = (request.form.get("s70date") or "").strip()
    return redirect("/tech?s70date=" + urllib.parse.quote(s70date))


@app.route("/tech/s70/delete/<int:sid>")
def tech_s70_delete(sid):
    conn = get_db()
    conn.execute("DELETE FROM s70_loans WHERE id=?", (sid,))
    conn.commit(); conn.close()
    return redirect("/tech")


@app.route("/tech/s70/batch_delete", methods=["POST"])
def tech_s70_batch_delete():
    ids = [x for x in request.form.getlist("ids") if x.isdigit()]
    if ids:
        conn = get_db()
        conn.execute("DELETE FROM s70_loans WHERE id IN (%s)" % ",".join("?" * len(ids)), ids)
        conn.commit()
        conn.close()
    return redirect("/tech")


S70_EDIT_COLS = S70_COLS  # 新增/编辑覆盖全部 50 列


@app.route("/tech/s70/save", methods=["POST"])
def tech_s70_save():
    f = request.form
    now = _now()

    def s(k):
        return (f.get(k) or "").strip()

    def n(k):
        v = (f.get(k) or "").strip()
        try:
            return float(v) if v else None
        except ValueError:
            return None

    def sv(c):
        if c in S70_REAL_COLS:
            return n(c)
        if c in S70_DATE_COLS:
            # 手动录入的 2026-10-31 / 46326 等统一成库里格式
            return norm_date(s(c), cn=(c == "data_date"))
        return s(c)

    vals = [sv(c) for c in S70_EDIT_COLS]
    conn = get_db()
    if f.get("id"):
        sets = ",".join("%s=?" % c for c in S70_EDIT_COLS)
        conn.execute("UPDATE s70_loans SET %s WHERE id=?" % sets, vals + [f["id"]])
    else:
        conn.execute("INSERT INTO s70_loans (%s, created) VALUES (%s)"
                     % (",".join(S70_EDIT_COLS), ",".join(["?"] * (len(S70_EDIT_COLS) + 1))),
                     vals + [now])
    conn.commit()
    conn.close()
    return redirect("/tech")


@app.route("/tech/s70/export")
def tech_s70_export():
    q_date_cn = iso_date_to_cn(request.args.get("q_date", "").strip())
    q_cust = request.args.get("q_cust", "").strip()
    q_unit = request.args.get("q_unit", "").strip()
    q_credit = request.args.get("q_credit", "").strip()
    cond, cargs = "1=1", []
    if q_date_cn:
        cond += " AND data_date = ?"; cargs.append(q_date_cn)
    if q_cust:
        cond += " AND cust_name LIKE ?"; cargs.append("%%%s%%" % q_cust)
    if q_unit:
        cond += " AND unit LIKE ?"; cargs.append("%%%s%%" % q_unit)
    if q_credit:
        cond += " AND credit_type=?"; cargs.append(q_credit)
    conn = get_db()
    rows = conn.execute("SELECT * FROM s70_loans WHERE %s ORDER BY data_date DESC, id DESC" % cond,
                        cargs).fetchall()
    conn.close()

    from openpyxl.styles import Font, PatternFill, Alignment
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S70贷款"
    ws.append(S70_HEADERS)
    fill = PatternFill("solid", fgColor="12224A")
    for c in ws[1]:
        c.fill = fill
        c.font = Font(color="FFFFFF", bold=True, size=10.5)
        c.alignment = Alignment(horizontal="center", vertical="center")
    for r in rows:
        ws.append([(s70_num(r[c]) if c in S70_REAL_COLS else (r[c] or "")) for c in S70_COLS])
    for i in range(1, len(S70_HEADERS) + 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = 16
    ws.column_dimensions[openpyxl.utils.get_column_letter(S70_COLS.index("cust_name") + 1)].width = 26
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(vertical="top", wrap_text=True)
    ws.freeze_panes = "A2"
    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    fname = "S70贷款_%s.xlsx" % datetime.datetime.now().strftime("%Y%m%d_%H%M")
    return send_file(bio, as_attachment=True, download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ----------------------------------------------------------------------------
# 3.7 债券融资 / 资本市场（项目推进台账）
# ----------------------------------------------------------------------------
PIPELINE_CATS = {
    "bond": ("📜 债券融资", "债券承销 / 投资 / 交易等项目推进台账"),
    "capmarket": ("📈 资本市场", "IPO / 再融资 / 并购重组等项目推进台账"),
}


def _pipeline(cat):
    title, desc = PIPELINE_CATS[cat]
    q = request.args.get("q", "").strip()
    conn = get_db()
    edit = None
    if request.args.get("edit"):
        edit = conn.execute("SELECT * FROM pipeline_items WHERE id=? AND category=?",
                            (request.args["edit"], cat)).fetchone()
    sql, args = "SELECT * FROM pipeline_items WHERE category=?", [cat]
    if q:
        sql += (" AND (name LIKE ? OR unit LIKE ? OR content LIKE ?"
                " OR manager LIKE ? OR note LIKE ?)")
        args += ["%%%s%%" % q] * 5
    sql += " ORDER BY id"
    rows = conn.execute(sql, args).fetchall()
    conn.close()

    trs = []
    for i, r in enumerate(rows):
        amt = ("%g" % r["amount"]) if r["amount"] else ""
        trs.append(
            "<tr><td>%d</td><td>%s</td><td>%s</td><td><b>%s</b></td><td>%s</td>"
            "<td>%s</td><td>%s</td><td>%s</td>"
            "<td><a href='/%s?edit=%d'>编辑</a> <a href='/pipeline/delete/%d' "
            "onclick=\"return confirm('确认删除？')\">删除</a></td></tr>"
            % (i + 1, r["push_date"] or "", r["unit"] or "", r["name"], amt,
               r["content"] or "", r["manager"] or "", r["note"] or "",
               cat, r["id"], r["id"]))
    trs_html = ("".join(trs)
                or "<tr><td class='mini' colspan=9>暂无项目，点击右上角「＋ 新增项目」开始</td></tr>")

    d = edit or {"id": "", "push_date": "", "unit": "", "name": "",
                 "amount": "", "content": "", "manager": "", "note": ""}
    form_display = "block" if edit else "none"
    add_btn = ("" if edit else
               "<span style='margin-left:auto'><a class='btn btn2 btn-sm' id='btnAddPipe' "
               "href=\"javascript:showBox('pipeFormCard','btnAddPipe')\">＋ 新增项目</a></span>")
    cancel_btn = ("<a class='btn btn3' href='/{cat}' style='margin-left:8px'>取消</a>".format(cat=cat) if edit else
                  "<a class='btn btn3' href=\"javascript:hideBox('pipeFormCard','btnAddPipe')\" "
                  "style='margin-left:8px'>取消</a>")
    body = """
    <div class="card"><h2>{title} · 项目推进台账 <span class="mini">共 {n} 条</span>{add_btn}</h2>
    <p class="mini" style="margin-bottom:10px">{desc}</p>
    <form method="get" style="display:flex;gap:8px;margin-bottom:12px">
      <input name="q" value="{q}" placeholder="搜索 项目/经营单位/商机内容" style="flex:1;min-width:180px">
      <button class="btn2">搜索</button><a class="btn3 btn" style="text-align:center" href="/{cat}">重置</a>
    </form>
    <div class="tbl-wrap"><table>
      <tr><th style="width:44px">序号</th><th style="width:92px">推进时间</th><th>经营单位</th>
      <th>项目名字</th><th style="width:110px">项目金额（亿元）</th><th>商机内容</th>
      <th>产品经理/分管经营单位</th><th>营销情况及难点</th><th style="width:96px">操作</th></tr>
      {trs}</table></div></div>
    <div class="card" id="pipeFormCard" style="display:{form_display}"><h2>{form_ttl}</h2>
    <form method="post" action="/pipeline/save" class="frm">
      <input type="hidden" name="id" value="{id}">
      <input type="hidden" name="cat" value="{cat}">
      <div class="row">
        <div><label>推进时间</label><input name="push_date" value="{push_date}" placeholder="如 2026-07"></div>
        <div><label>经营单位</label><input name="unit" value="{unit}"></div>
        <div style="flex:2"><label>项目名字 *</label><input name="name" required value="{name}"></div>
        <div><label>项目金额（亿元）</label><input name="amount" type="number" step="any" value="{amount}"></div>
      </div>
      <label>商机内容</label><input name="content" value="{content}">
      <div class="row">
        <div><label>产品经理/分管经营单位</label><input name="manager" value="{manager}"></div>
        <div style="flex:2"><label>营销情况及难点</label><textarea name="note" rows="2">{note}</textarea></div>
      </div>
      <div style="margin-top:12px"><button>保存</button>{cancel_btn}</div></form></div>
    """.format(title=title, desc=desc, n=len(rows), q=q, cat=cat, trs=trs_html,
               form_ttl=("编辑 #%s" % d["id"]) if d["id"] else "＋ 新增项目",
               id=d["id"], push_date=d["push_date"] or "", unit=d["unit"] or "",
               name=d["name"], amount=(d["amount"] if d["amount"] else ""),
               content=d["content"] or "", manager=d["manager"] or "",
               note=d["note"] or "",
               form_display=form_display, add_btn=add_btn, cancel_btn=cancel_btn)
    return page(body, cat)


@app.route("/bond")
def bond():
    return _pipeline("bond")


@app.route("/capmarket")
def capmarket():
    return _pipeline("capmarket")


@app.route("/pipeline/save", methods=["POST"])
def pipeline_save():
    f = request.form
    cat = f.get("cat", "bond")
    if cat not in PIPELINE_CATS:
        cat = "bond"
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    try:
        amount = float(f.get("amount", "") or 0)
    except ValueError:
        amount = 0
    vals = (f.get("push_date", "").strip(), f.get("unit", "").strip(),
            f.get("name", "").strip() or "未命名项目", amount,
            f.get("content", "").strip(), f.get("manager", "").strip(),
            f.get("note", "").strip())
    conn = get_db()
    if f.get("id"):
        conn.execute("""UPDATE pipeline_items SET push_date=?, unit=?, name=?, amount=?,
                        content=?, manager=?, note=?, updated=? WHERE id=?""",
                     vals + (now, f["id"]))
    else:
        conn.execute("""INSERT INTO pipeline_items
                        (category, push_date, unit, name, amount, content, manager, note,
                         created, updated) VALUES (?,?,?,?,?,?,?,?,?,?)""",
                     (cat,) + vals + (now, now))
    conn.commit()
    conn.close()
    return redirect("/" + cat)


@app.route("/pipeline/delete/<int:pid>")
def pipeline_delete(pid):
    conn = get_db()
    row = conn.execute("SELECT category FROM pipeline_items WHERE id=?", (pid,)).fetchone()
    conn.execute("DELETE FROM pipeline_items WHERE id=?", (pid,))
    conn.commit()
    conn.close()
    return redirect("/" + (row["category"] if row else "bond"))


# ----------------------------------------------------------------------------
# 3.6 综合融资（发改委项目储备库 + 关键词订阅）
# ----------------------------------------------------------------------------
@app.route("/gov")
def gov():
    unit = request.args.get("unit", "").strip()
    name = request.args.get("name", "").strip()
    region = request.args.get("region", "").strip()
    push_date = request.args.get("push_date", "").strip()
    try:
        cur_page = max(1, int(request.args.get("page", "1")))
    except ValueError:
        cur_page = 1
    PER = 10

    cond, args = "1=1", []
    if unit:
        cond += " AND unit LIKE ?"; args.append("%%%s%%" % unit)
    if name:
        cond += " AND name LIKE ?"; args.append("%%%s%%" % name)
    if region:
        cond += " AND region LIKE ?"; args.append("%%%s%%" % region)
    if push_date:
        cond += " AND push_date LIKE ?"; args.append("%%%s%%" % push_date)

    conn = get_db()
    total = conn.execute("SELECT COUNT(*) c FROM gudai_deals WHERE " + cond, args).fetchone()["c"]
    pages = max(1, (total + PER - 1) // PER)
    cur_page = min(cur_page, pages)
    gd_rows = conn.execute(
        "SELECT * FROM gudai_deals WHERE %s ORDER BY id LIMIT ? OFFSET ?" % cond,
        args + [PER, (cur_page - 1) * PER]).fetchall()
    gd_edit = None
    if request.args.get("gid"):
        gd_edit = conn.execute("SELECT * FROM gudai_deals WHERE id=?",
                               (request.args["gid"],)).fetchone()
    conn.close()

    def _q(**kw):
        p = {"unit": unit, "name": name, "region": region, "push_date": push_date}
        p.update(kw)
        return urllib.parse.urlencode({k: v for k, v in p.items() if v})

    export_qs = _q()
    export_url = "/gov/gudai/export" + ("?" + export_qs if export_qs else "")
    prev_url = "/gov" + ("?" + _q(page=cur_page - 1) if cur_page > 1 else "")
    next_url = "/gov" + ("?" + _q(page=cur_page + 1) if cur_page < pages else "")

    gd_trs = "".join(
        "<tr><td>%d</td><td>%s</td><td>%s</td><td><b>%s</b></td><td>%s</td><td>%s</td>"
        "<td>%s</td><td>%s</td><td>%s</td><td>%s</td>"
        "<td><a href='/gov?gid=%d&%s'>编辑</a> "
        "<a href='/gov/gudai/delete/%d' onclick=\"return confirm('确认删除？')\">删除</a></td></tr>"
        % ((cur_page - 1) * PER + i + 1, g["push_date"] or "—", g["unit"] or "—", g["name"] or "—",
           g["region"] or "—", ("%g" % g["amount"]) if g["amount"] else "—",
           g["content"] or "—", g["manager"] or "—", g["note"] or "—",
           ("%g" % g["approval_amount"]) if g["approval_amount"] else "—",
           g["id"], _q(), g["id"])
        for i, g in enumerate(gd_rows)) or ("<tr><td class='mini' colspan=11>"
        "暂无固贷商机，点击右上角「＋ 新增项目」或到「商机管理 → 粘贴导入」导入</td></tr>")

    pager = ""
    if pages > 1:
        pager = ("<div class='pager'><a href='%s'>上一页</a>"
                 "<span>第 %d / %d 页 · 共 %d 条</span><a href='%s'>下一页</a></div>"
                 % (prev_url, cur_page, pages, total, next_url))
    pager = pager or ("<div class='pager'><span>共 %d 条</span></div>" % total)

    d = gd_edit or {"id": "", "push_date": "", "unit": "", "name": "", "region": "",
                    "amount": "", "content": "", "manager": "", "note": "",
                    "approval_amount": ""}
    form_display = "block" if gd_edit else "none"
    add_btn = ("" if gd_edit else
               "<a class='btn btn2 btn-sm' id='btnAddGudai' "
               "href=\"javascript:showBox('gudaiFormCard','btnAddGudai')\">＋ 新增项目</a>")
    cancel_btn = ("<a class='btn btn3' href='/gov' style='margin-left:8px'>取消</a>" if gd_edit else
                  "<a class='btn btn3' href=\"javascript:hideBox('gudaiFormCard','btnAddGudai')\" "
                  "style='margin-left:8px'>取消</a>")
    body = """
    <div class="card"><h2>固贷商机台账
      <span style="margin-left:auto;display:flex;gap:8px;align-items:center">
        <a class="btn btn3 btn-sm" href="{export_url}">⬇ 导出</a>{add_btn}</span></h2>
    <form method="get" style="display:flex;gap:8px;margin-bottom:12px;flex-wrap:wrap">
      <input name="unit" value="{unit}" placeholder="经营单位" style="width:120px">
      <input name="name" value="{name}" placeholder="项目" style="flex:1;min-width:160px">
      <input name="region" value="{region}" placeholder="区域" style="width:110px">
      <input name="push_date" value="{push_date}" placeholder="推送日期 如10月29日" style="width:150px">
      <button class="btn2">查询</button><a class="btn3 btn" style="text-align:center" href="/gov">重置</a>
    </form>
    <div class="tbl-wrap"><table>
      <tr><th style="width:40px">序号</th><th style="width:78px">推送时间</th><th>经营单位</th>
      <th>项目名字</th><th>区域</th><th style="width:100px">项目金额（亿元）</th><th>商机情况</th>
      <th>产品经理/分配经营单位</th><th>营销情况及难点</th><th style="width:100px">报批金额（亿元）</th><th>操作</th></tr>
      {gd_trs}</table></div>
    {pager}</div>
    <div class="card" id="gudaiFormCard" style="display:{form_display}"><h2>{form_ttl}</h2>
    <form method="post" action="/gov/gudai/save" class="frm">
      <input type="hidden" name="id" value="{id}">
      <div class="row">
        <div><label>推送时间</label><input name="push_date" value="{push_date_val}" placeholder="如 10月29日"></div>
        <div><label>经营单位</label><input name="unit" value="{unit_val}" placeholder="如 海珠支行"></div>
        <div style="flex:2"><label>项目名字 *</label><input name="name" required value="{name_val}"></div>
        <div><label>区域</label><input name="region" value="{region_val}" placeholder="如 东莞市"></div>
      </div>
      <div class="row">
        <div><label>项目金额（亿元）</label><input name="amount" type="number" step="any" value="{amount}"></div>
        <div><label>产品经理/分配经营单位</label><input name="manager" value="{manager}" placeholder="留空则按经营单位自动匹配"></div>
        <div><label>报批金额（亿元）</label><input name="approval_amount" type="number" step="any" value="{approval_amount}"></div>
      </div>
      <label>商机情况</label><textarea name="content" rows="3">{content}</textarea>
      <label>营销情况及难点</label><textarea name="note" rows="2">{note}</textarea>
      <div style="margin-top:12px"><button>保存</button>{cancel_btn}</div></form></div>
    """.format(export_url=export_url, add_btn=add_btn, unit=unit, name=name,
               region=region, push_date=push_date, gd_trs=gd_trs, pager=pager,
               form_display=form_display,
               form_ttl=("编辑固贷商机 #%s" % d["id"]) if d["id"] else "新增固贷商机",
               id=d["id"], push_date_val=d["push_date"] or "", unit_val=d["unit"] or "",
               name_val=d["name"], region_val=d["region"] or "",
               amount=(d["amount"] if d["amount"] else ""),
               manager=d["manager"] or "",
               approval_amount=(d["approval_amount"] if d["approval_amount"] else ""),
               content=d["content"] or "", note=d["note"] or "",
               cancel_btn=cancel_btn)
    return page(body, "gov")


@app.route("/gov/gudai/save", methods=["POST"])
def gov_gudai_save():
    f = request.form
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

    def _num(k):
        v = (f.get(k) or "").strip()
        try:
            return float(v) if v else None
        except ValueError:
            return None

    unit = (f.get("unit") or "").strip()
    manager = (f.get("manager") or "").strip() or GUDAI_DIVISION.get(unit, "")
    vals = ((f.get("push_date") or "").strip(), unit,
            (f.get("name") or "").strip() or "未命名项目",
            (f.get("region") or "").strip(), _num("amount"),
            (f.get("content") or "").strip(), manager,
            (f.get("note") or "").strip(), _num("approval_amount"))
    conn = get_db()
    if f.get("id"):
        conn.execute("""UPDATE gudai_deals SET push_date=?, unit=?, name=?, region=?, amount=?,
                        content=?, manager=?, note=?, approval_amount=?, updated=? WHERE id=?""",
                     vals + (now, f["id"]))
    else:
        conn.execute("""INSERT INTO gudai_deals (push_date, unit, name, region, amount, content,
                        manager, note, approval_amount, created, updated) VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
                     vals + (now, now))
    conn.commit()
    conn.close()
    return redirect("/gov")


@app.route("/gov/gudai/delete/<int:gid>")
def gov_gudai_delete(gid):
    conn = get_db()
    conn.execute("DELETE FROM gudai_deals WHERE id=?", (gid,))
    conn.commit()
    conn.close()
    return redirect("/gov")


@app.route("/gov/gudai/export")
def gov_gudai_export():
    unit = request.args.get("unit", "").strip()
    name = request.args.get("name", "").strip()
    region = request.args.get("region", "").strip()
    push_date = request.args.get("push_date", "").strip()
    cond, args = "1=1", []
    if unit:
        cond += " AND unit LIKE ?"; args.append("%%%s%%" % unit)
    if name:
        cond += " AND name LIKE ?"; args.append("%%%s%%" % name)
    if region:
        cond += " AND region LIKE ?"; args.append("%%%s%%" % region)
    if push_date:
        cond += " AND push_date LIKE ?"; args.append("%%%s%%" % push_date)
    conn = get_db()
    rows = conn.execute("SELECT * FROM gudai_deals WHERE %s ORDER BY id" % cond, args).fetchall()
    conn.close()

    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "固贷商机台账"
    head = ["序号", "推送时间", "经营单位", "项目名字", "区域", "项目金额（亿元）",
            "商机情况", "产品经理/分配经营单位", "营销情况及难点", "报批金额（亿元）"]
    ws.append(head)
    fill = PatternFill("solid", fgColor="12224A")
    thin = Side(style="thin", color="D9DEE8")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for c in ws[1]:
        c.fill = fill
        c.font = Font(color="FFFFFF", bold=True, size=10.5)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border
    for i, g in enumerate(rows, start=1):
        ws.append([i, g["push_date"] or "", g["unit"] or "", g["name"] or "",
                   g["region"] or "", g["amount"] if g["amount"] else "",
                   g["content"] or "", g["manager"] or "", g["note"] or "",
                   g["approval_amount"] if g["approval_amount"] else ""])
    widths = [6, 12, 14, 30, 12, 14, 60, 18, 30, 14]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(vertical="top", wrap_text=True)
    ws.freeze_panes = "A2"
    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    fname = "固贷商机台账_%s.xlsx" % datetime.datetime.now().strftime("%Y%m%d_%H%M")
    return send_file(bio, as_attachment=True, download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/gov/save", methods=["POST"])
def gov_save():
    f = request.form
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    conn = get_db()
    if f.get("id"):
        conn.execute("""UPDATE gov_projects SET name=?,region=?,sector=?,invest=?,demand=?,
                        stage=?,source=?,owner=?,note=?,updated=? WHERE id=?""",
                     (f["name"], f.get("region"), f.get("sector"),
                      float(f.get("invest") or 0), float(f.get("demand") or 0),
                      f.get("stage"), f.get("source"), f.get("owner"), f.get("note"), now, f["id"]))
    else:
        conn.execute("""INSERT INTO gov_projects (name,region,sector,invest,demand,stage,source,owner,note,created,updated)
                        VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
                     (f["name"], f.get("region"), f.get("sector"),
                      float(f.get("invest") or 0), float(f.get("demand") or 0),
                      f.get("stage"), f.get("source"), f.get("owner"), f.get("note"), now, now))
    conn.commit(); conn.close()
    return redirect(url_for("gov"))


@app.route("/gov/delete/<int:did>")
def gov_delete(did):
    conn = get_db()
    conn.execute("DELETE FROM gov_projects WHERE id=?", (did,))
    conn.commit(); conn.close()
    return redirect(url_for("gov"))


@app.route("/gov/sub", methods=["POST"])
def gov_sub():
    kw = request.form.get("keyword", "").strip()
    if kw:
        conn = get_db()
        conn.execute("INSERT INTO gov_subs (keyword,typ,created) VALUES (?,?,?)",
                     (kw, request.form.get("typ", "行业"),
                      datetime.datetime.now().strftime("%Y-%m-%d")))
        conn.commit(); conn.close()
    return redirect(url_for("gov"))


@app.route("/gov/unsub/<int:sid>")
def gov_unsub(sid):
    conn = get_db()
    conn.execute("DELETE FROM gov_subs WHERE id=?", (sid,))
    conn.commit(); conn.close()
    return redirect(url_for("gov"))


# ----------------------------------------------------------------------------
# 4. 素材库
# ----------------------------------------------------------------------------
@app.route("/library")
def library():
    root = get_setting("material_root", "")
    q = request.args.get("q", "").strip()
    cat = request.args.get("cat", "")
    auto_on = get_setting("auto_scan_on", "1") == "1"
    auto_time = get_setting("auto_scan_time", "08:30")
    conn = get_db()
    sql, args = "SELECT * FROM materials WHERE 1=1", []
    if q:
        sql += " AND (name LIKE ? OR rel LIKE ?)"
        args += ["%%%s%%" % q] * 2
    if cat:
        sql += " AND cat=?"
        args.append(cat)
    sql += " ORDER BY rel"
    mats = conn.execute(sql, args).fetchall()
    logs = conn.execute("SELECT * FROM scan_logs ORDER BY id DESC LIMIT 6").fetchall()
    conn.close()
    cats = ["年度总结及计划", "投行条线汇报", "综合汇报与上会"]
    trs = "".join(
        "<tr><td>%s</td><td><b>%s</b><div class='mini' style='font-size:10.5px'>%s</div></td>"
        "<td><span class='pill %s'>%s</span></td><td class='mini'>%s</td><td class='mini'>%s</td>"
        "<td>%s</td></tr>"
        % (m["kind"].upper(),
           m["name"], os.path.join(root, m["rel"]),
           {"年度总结及计划": "p5", "投行条线汇报": "p1", "综合汇报与上会": "p3"}.get(m["cat"], "p6"),
           m["cat"], "%s KB" % int(m["size"]) if m["size"] < 1024 else "%.1f MB" % (m["size"] / 1024),
           m["mtime"],
           "<a href='/library/outline/%d'>查看结构</a>" % m["id"] if m["kind"] == "docx" else "—")
        for m in mats) or "<tr><td class='mini' colspan=6>素材库为空，请先扫描</td></tr>"
    log_trs = "".join(
        "<tr><td class='mini'>%s</td><td><span class='pill %s'>%s</span></td>"
        "<td class='mini'>新增 <b style='color:#1a7f4d'>+%d</b> ／ 更新 <b style='color:#b97a12'>↻%d</b> ／ 移除 <b style='color:#d0342c'>−%d</b></td>"
        "<td class='mini'>共 %d 个文件</td></tr>"
        % (l["ts"], "p1" if l["trigger"] == "定时" else "p6", l["trigger"] or "手动",
           l["added"] or 0, l["updated"] or 0, l["removed"] or 0, l["total"] or 0)
        for l in logs) or "<tr><td class='mini' colspan=4>暂无扫描记录</td></tr>"
    last_auto = get_setting("last_auto_scan_date", "") or "尚未运行"
    body = """
    <div class="card"><h2>素材库（本地材料索引）</h2>
    <form method="post" action="/library/root" class="frm" style="margin-bottom:10px">
      <label>材料文件夹路径（局域网共享盘亦可）</label>
      <div class="row">
        <input name="root" value="{root}" placeholder="例如 D:\\科技金融">
        <button style="max-width:120px">保存路径</button>
      </div></form>
    <form method="post" action="/library/autoscan" style="margin-bottom:6px;display:flex;gap:10px;align-items:center;flex-wrap:wrap;background:#f7f9fc;border:1px solid #e5e9f0;border-radius:8px;padding:10px 12px">
      <b style="font-size:12.5px">⏰ 每日自动扫描</b>
      <label style="display:flex;align-items:center;gap:5px;font-size:12.5px;cursor:pointer">
        <input type="checkbox" name="on" value="1" {ck} style="width:auto"> 启用（程序运行期间，每天到点自动重扫并更新索引）
      </label>
      <label style="font-size:12.5px">时间 <input type="time" name="time" value="{tm}" style="width:100px;padding:5px 8px"></label>
      <button class="btn2" style="max-width:100px">保存设置</button>
      <span class="mini">上次定时扫描：{last}</span>
    </form>
    <form method="post" action="/library/scan" style="margin-bottom:12px">
      <button>↻ 立即扫描并更新索引</button></form>
    <form method="get" style="display:flex;gap:8px;margin-bottom:12px;flex-wrap:wrap">
      <input name="q" value="{q}" placeholder="搜索材料名称" style="flex:1;min-width:180px">
      <select name="cat" style="width:150px"><option value="">全部分类</option>{cs}</select>
      <button class="btn2">筛选</button></form>
    <table><tr><th>类型</th><th>材料（路径）</th><th>分类</th><th>大小</th><th>更新</th><th>结构</th></tr>{trs}</table>
    <p class="mini" style="margin-top:8px">「查看结构」＝解析 Word 文档的标题大纲，供初稿生成时作为模板骨架。文件本身不会被修改或上传。</p></div>

    <div class="card"><h2>扫描记录（最近 6 次）</h2>
    <table><tr><th>时间</th><th>方式</th><th>本次变化</th><th>索引结果</th></tr>{log_trs}</table>
    <p class="mini" style="margin-top:8px">「新增/更新/移除」为本次扫描相对上一次索引的差异——新放进文件夹的材料次日自动进入索引，删除的材料自动出索引。</p></div>
    """.format(root=root, q=q,
               ck="checked" if auto_on else "", tm=auto_time,
               last=last_auto,
               cs="".join("<option %s>%s</option>" % ("selected" if c == cat else "", c) for c in cats),
               trs=trs, log_trs=log_trs)
    return page(body, "library")


@app.route("/library/autoscan", methods=["POST"])
def library_autoscan():
    set_setting("auto_scan_on", "1" if request.form.get("on") == "1" else "0")
    t = request.form.get("time", "08:30").strip()
    if re.match(r"^\d{1,2}:\d{2}$", t):
        set_setting("auto_scan_time", t)
    return redirect("/library")


@app.route("/library/root", methods=["POST"])
def library_root():
    root = request.form.get("root", "").strip()
    if root and not os.path.isdir(root):
        return page("<div class='err'>路径不存在：%s</div>" % root, "library")
    set_setting("material_root", root)
    scan_materials(root)
    return redirect("/library")


@app.route("/library/scan", methods=["POST", "GET"])
def library_scan():
    root = get_setting("material_root", "")
    if not root or not os.path.isdir(root):
        return redirect("/library")
    n = scan_materials(root)
    conn = get_db()
    total = conn.execute("SELECT COUNT(*) c FROM materials").fetchone()["c"]
    conn.close()
    return redirect("/library")


@app.route("/library/outline/<int:mid>")
def library_outline(mid):
    conn = get_db()
    m = conn.execute("SELECT * FROM materials WHERE id=?", (mid,)).fetchone()
    conn.close()
    root = get_setting("material_root", "")
    items = docx_outline(os.path.join(root, m["rel"])) if m else []
    ol = "".join("<div style='padding-left:%dpx;font-size:12.5px;line-height:2;color:#41507a'>%s %s</div>"
                 % (h["lvl"] * 20 + 8, ["▸", "·", "  "][h["lvl"]], h["t"]) for h in items)
    body = """
    <div class="card"><h2>《%s》标题结构（模板骨架）</h2>
    <div style="border-left:3px solid #2456c8;padding:6px 12px;background:#fbfcff;border-radius:8px">%s</div>
    <p class="mini" style="margin-top:10px">此结构将用于初稿生成的章节参照。
    <a class="btn" style="margin-top:10px" href="/draft?mid=%d">→ 用此结构生成初稿</a>
    <a class="btn btn3" style="margin-left:8px" href="/library">返回素材库</a></p></div>
    """ % (m["name"], ol or "<span class='mini'>未解析到标题样式（可能用了手动加粗而非标题样式），可在 Word 中为标题应用样式后重新扫描。</span>", mid)
    return page(body, "library")


# ----------------------------------------------------------------------------
# 5. 初稿生成（参照原稿全文格式与框架）
# ----------------------------------------------------------------------------
def data_snapshot():
    """从工作台台账自动取当期数据速览文字（预览与导出共用）"""
    conn = get_db()
    n_all = conn.execute("SELECT COUNT(*) c FROM deals").fetchone()["c"]
    stat = conn.execute("""SELECT stage, COUNT(*) c, COALESCE(SUM(amount),0) s FROM deals
                           GROUP BY stage ORDER BY s DESC""").fetchall()
    recent = conn.execute("""SELECT name, client, stage, amount, owner FROM deals
                             ORDER BY updated DESC LIMIT 5""").fetchall()
    n_exp = conn.execute("SELECT COUNT(*) c FROM expansion_clients").fetchone()["c"]
    exp_stat = conn.execute("SELECT status, COUNT(*) c FROM expansion_clients GROUP BY status").fetchall()
    n_gov = conn.execute("SELECT COUNT(*) c FROM gov_projects").fetchone()["c"]
    gov_sum = conn.execute("SELECT COALESCE(SUM(demand),0) s FROM gov_projects").fetchone()["s"]
    conn.close()
    lines = ["截至 %s 台账数据：" % datetime.datetime.now().strftime("%Y-%m-%d")]
    if n_all:
        lines.append("一、在库商机 %d 个，金额合计 %s 万元。其中：" % (
            n_all, f"{sum(r['s'] for r in stat):,.0f}"))
        lines.append("　　" + "；".join("%s %d 个（%s 万元）" % (r["stage"], r["c"], f"{r['s']:,.0f}")
                                     for r in stat) + "。")
        if recent:
            lines.append("二、近期重点商机：")
            lines.extend("　　- %s（%s，%s，%s 万元，负责人 %s）"
                         % (r["name"], r["client"] or "—", r["stage"],
                            f"{r['amount']:,.0f}", r["owner"] or "—") for r in recent)
    if n_exp:
        lines.append("三、科技金融拓户：累计名单 %d 家，%s。" % (
            n_exp, "、".join("%s %d 家" % (r["status"], r["c"]) for r in exp_stat)))
    if n_gov:
        lines.append("四、综合融资储备项目 %d 个，在途融资需求合计 %s 万元。" % (n_gov, f"{gov_sum:,.0f}"))
    return lines if len(lines) > 1 else ["（台账暂无数据，请先在商机管理/科技金融/综合融资模块录入）"]


@app.route("/draft")
def draft():
    mid = request.args.get("mid", "")
    conn = get_db()
    root = get_setting("material_root", "")
    mats = conn.execute("SELECT * FROM materials WHERE kind='docx' ORDER BY rel").fetchall()
    conn.close()
    cands = [m for m in mats if re.search(r"总结|计划|汇报|报告|分析", m["name"])] or mats
    target = next((m for m in mats if str(m["id"]) == mid), None) or \
             (max(cands, key=lambda m: m["mtime"]) if cands else None)
    if not target:
        return page("<div class='card'><h2>初稿生成</h2><p class='mini'>素材库暂无 DOCX 材料，请先在素材库扫描。</p></div>", "draft")

    items = docx_content(os.path.join(root, target["rel"]))
    snap = data_snapshot()

    # ---- 渲染预览 ----
    n_h = sum(1 for k, _ in items if k.startswith("h"))
    n_p = sum(1 for k, _ in items if k == "p")
    n_t = sum(1 for k, _ in items if k == "table")
    n_data = sum(1 for k, t in items if k == "p" and has_data(t))

    html = []
    for k, v in items:
        if k.startswith("h"):
            lvl = min(int(k[1]), 4)
            html.append("<div style='font-weight:700;font-size:%dpx;margin:%dpx 0 %dpx;color:#12224a'>%s</div>"
                        % (19 - lvl * 2, 16, 6, v))
        elif k == "table":
            trs = "".join("<tr>%s</tr>" % "".join("<td class='mini' style='padding:5px 8px'>%s</td>" % c
                                                  for c in row) for row in v)
            html.append("<table style='margin:8px 0;border:1px solid #e5e9f0'>%s</table>" % trs)
        else:
            flag = has_data(v)
            html.append("<div style='line-height:1.9;margin:7px 0;%s'>%s%s</div>" % (
                "background:#fdf6e3;border-left:3px solid #e6a23c;padding:6px 10px;border-radius:4px" if flag else "",
                "<span class='pill p3' style='margin-right:8px' title='含当期数据，请核对更新'>⚠ 待更新</span>" if flag else "",
                v))
    content_html = "".join(html) or "<p class='mini'>未能解析该文档正文（可能为扫描件或加密文档）。</p>"
    snap_html = "".join("<div style='line-height:1.9'>%s</div>" % s for s in snap)

    opts = "".join("<option value='%d' %s>%s</option>"
                   % (m["id"], "selected" if target and m["id"] == target["id"] else "", m["name"]) for m in mats)
    body = """
    <div class="card"><h2>初稿生成（参照原稿格式与框架）</h2>
    <form method="get" class="frm"><div class="row">
      <div><label>参照原稿（结构与正文均取自该材料）</label>
        <select name="mid" onchange="this.form.submit()">{opts}</select></div>
    </div></form>
    <div style="display:flex;gap:8px;margin:10px 0;flex-wrap:wrap;align-items:center">
      <form method="post" action="/draft/download"><input type="hidden" name="mid" value="{mid}">
        <button>⬇ 导出 Word 初稿（.docx）</button></form>
      <span class="mini">原稿共解析：标题 {nh} 个 ｜ 正文段落 {np} 段（其中 {nd} 段含数据，已标黄待更新）｜ 表格 {nt} 个</span>
    </div>
    <div style="background:#f0f6ff;border:1px solid #c7d6f5;border-radius:8px;padding:10px 14px;margin-bottom:14px">
      <b style="font-size:12.5px">当期数据速览（自动取自工作台台账，可整段复制进正文）</b>
      {snap}</div>
    <p class="mini" style="margin-bottom:10px">以下初稿<b>沿用原稿全文</b>：结构、标题层级、正文、表格全部保留；
    含数字/金额的段落以黄色标出，提示替换为当期数据；文首已附当期数据速览。定稿责任仍在撰写人。</p>
    <div style="border:1px solid #e5e9f0;border-radius:8px;padding:14px 18px;background:#fff">{content}</div>
    </div>
    """.format(opts=opts, mid=target["id"], nh=n_h, np=n_p, nd=n_data, nt=n_t,
               snap=snap_html, content=content_html)
    return page(body, "draft")


@app.route("/draft/download", methods=["POST"])
def draft_download():
    mid = request.form.get("mid")
    conn = get_db()
    m = conn.execute("SELECT * FROM materials WHERE id=?", (mid,)).fetchone()
    conn.close()
    root = get_setting("material_root", "")
    items = docx_content(os.path.join(root, m["rel"])) if m else []
    snap = data_snapshot()

    doc = Document()
    # 全局中文字体
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(12)
    normal._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "仿宋")

    t = doc.add_paragraph()
    r = t.add_run("%s（初稿）" % ((m["name"].rsplit(".", 1)[0] if m else "材料") + "——当期版"))
    r.bold = True
    r.font.size = Pt(16)
    info = doc.add_paragraph("生成时间：%s ｜ 结构与正文参照原稿：%s ｜ 含数据段落已标黄，请核对更新后使用"
                             % (datetime.datetime.now().strftime("%Y-%m-%d %H:%M"), m["name"] if m else "—"))
    info.runs[0].font.size = Pt(9)
    info.runs[0].font.color.rgb = RGBColor(0x88, 0x88, 0x88)

    doc.add_heading("当期数据速览（自动生成，取自工作台台账）", level=1)
    for s in snap:
        doc.add_paragraph(s)

    for k, v in items:
        if k.startswith("h"):
            doc.add_heading(v, level=min(int(k[1]), 4))
        elif k == "table":
            tb = doc.add_table(rows=len(v), cols=len(v[0]))
            tb.style = "Table Grid"
            for i, row in enumerate(v):
                for j, cell in enumerate(row):
                    if j < len(v[0]):
                        c = tb.rows[i].cells[j]
                        c.text = cell
                        if i == 0:
                            for run in c.paragraphs[0].runs:
                                run.bold = True
        else:
            p = doc.add_paragraph(v)
            if has_data(v):
                p.add_run("　【待更新】").font.highlight_color = WD_COLOR_INDEX.YELLOW
                for run in p.runs:
                    run.font.highlight_color = WD_COLOR_INDEX.YELLOW

    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return send_file(bio, as_attachment=True,
                     download_name="初稿_%s_%s.docx" % (m["name"][:20] if m else "material",
                                                        datetime.datetime.now().strftime("%m%d_%H%M")),
                     mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


# ----------------------------------------------------------------------------
# 2.1 商机批量导入（Excel / CSV）
# ----------------------------------------------------------------------------
DEAL_FIELDS = ["name", "client", "sector", "stage", "status", "amount", "owner", "teammate", "note",
               "deal_type", "sub_type", "counterparty", "method", "amount2", "ratio",
               "location", "org", "source", "notice_date", "entry", "deal_date"]
IMPORT_DIR = os.path.join(UPLOAD_DIR, "imports")
TEMPLATE_PATH = os.path.join(BASE_DIR, "商机导入模板.xlsx")
os.makedirs(IMPORT_DIR, exist_ok=True)

# 模板示例（来自实际报送数据，导入前请删除）
TEMPLATE_SAMPLES = [
    ["资本市场商机", "回购增持", "生益科技股份回购", "广东生益科技股份有限公司", "", "电子材料",
     10000, 20000, "股份回购", "", "初步接洽", "进行中", "", "", "东莞分行",
     "分行报送", "", "东莞市", "营销回购专用账户及回购资金托管", "拟回购 1-2 亿元"],
    ["增存商机", "一级市场股权融资款", "深度医疗 B 轮股权融资款承接", "广州深度医疗器械科技有限公司",
     "幂方健康基金、晨晖创投等", "医疗器械", 10000, "", "B 轮股权融资", "",
     "初步接洽", "进行中", "", "", "广州分行", "上市公司公告", "2026-10-19", "广州市黄埔区",
     "以科技金融特色产品切入，营销股权融资款归集", "2023 年成立，旗舰产品 FUNZAR 手术机器人"],
    ["并购商机", "发行股份购买资产", "博盈特焊收购罗克莱 55% 股权", "博盈特焊技术股份有限公司",
     "苏州罗克莱海洋装备股份有限公司", "高端装备", "", "", "发行股份+现金", 55,
     "初步接洽", "进行中", "", "", "江门分行", "上市公司公告", "", "苏州市",
     "并购贷款 + 配套募集资金承接", "竞买方主办行江门分行，标的主办行苏州分行"],
    ["非房固贷商机", "发改委备案项目", "粤港澳大湾区电算协同一体化智算中心（一期）",
     "广东舜通大数据产业有限公司", "", "大数据/算力", 247500, "", "固定资产贷款", "",
     "初步接洽", "进行中", "", "", "韶关分行", "发改委备案", "", "韶关市韶关新区",
     "以项目贷切入，配套算力中心建设融资", "总投资 24.75 亿元，发改委已备案"],
]


def _norm_header(h):
    return re.sub(r"[\s\*＊:：()（）]", "", str(h or ""))


H2F = {}
for _h, _f, _req, _desc in IMPORT_COLUMNS:
    H2F[_norm_header(_h)] = _f
H2F[_norm_header("金额(万元)")] = "amount"
H2F[_norm_header("金额上限(万元)")] = "amount2"
H2F[_norm_header("股权比例")] = "ratio"


def build_deal_template():
    """生成导入模板：商机数据（表头+示例+下拉）+ 填写说明"""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.worksheet.datavalidation import DataValidation

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "商机数据"

    head_fill = PatternFill("solid", fgColor="12224A")
    head_font = Font(color="FFFFFF", bold=True, size=10.5)
    thin = Side(style="thin", color="D9DEE8")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for i, (h, f, req, desc) in enumerate(IMPORT_COLUMNS, start=1):
        c = ws.cell(row=1, column=i, value=h + ("*" if req else ""))
        c.fill, c.font = head_fill, head_font
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border

    date_idx = [c[1] for c in IMPORT_COLUMNS].index("deal_date")
    samples = []
    for row in TEMPLATE_SAMPLES:
        row = list(row)
        if len(row) == len(IMPORT_COLUMNS) - 1:
            row.insert(date_idx, "")
        samples.append(row)

    for r, row in enumerate(samples, start=2):
        for i, v in enumerate(row, start=1):
            c = ws.cell(row=r, column=i, value=v)
            c.border = border
            c.alignment = Alignment(vertical="top", wrap_text=True)
            if r <= 5:
                c.font = Font(color="8B95AB", italic=True, size=10)

    widths = [15, 18, 30, 28, 26, 12, 12, 14, 18, 10, 12, 10, 10, 12, 14, 14, 13, 14, 13, 30, 26]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 26
    ws.freeze_panes = "A2"

    # 下拉校验
    def add_dv(col, options, prompt):
        dv = DataValidation(type="list", formula1='"%s"' % ",".join(options),
                            allow_blank=True, showDropDown=False)
        dv.error = "请从下拉列表中选择"
        dv.promptTitle, dv.prompt = "可选值", prompt
        ws.add_data_validation(dv)
        dv.add("%s2:%s800" % (col, col))

    add_dv("A", DEAL_TYPES, "、".join(DEAL_TYPES))
    add_dv("K", STAGES, "、".join(STAGES))
    add_dv("L", STATUSES, "、".join(STATUSES))

    # 说明页
    ws2 = wb.create_sheet("填写说明")
    ws2.column_dimensions["A"].width = 20
    ws2.column_dimensions["B"].width = 10
    ws2.column_dimensions["C"].width = 78
    ws2["A1"] = "商机导入模板 · 填写说明"
    ws2["A1"].font = Font(bold=True, size=14, color="12224A")
    ws2["A2"] = "填好「商机数据」表后，回到工作台「商机管理 → 批量导入」上传即可。示例行（灰色斜体 4 行）请先删除。"
    ws2["A2"].font = Font(color="8A6D1F", size=10)

    for i, txt in enumerate(["字段", "必填", "填写说明"], start=1):
        c = ws2.cell(row=4, column=i, value=txt)
        c.fill, c.font = head_fill, head_font
        c.alignment = Alignment(horizontal="center")

    r = 5
    for h, f, req, desc in IMPORT_COLUMNS:
        ws2.cell(row=r, column=1, value=h).font = Font(bold=True, size=10)
        ws2.cell(row=r, column=2, value="是" if req else "").alignment = Alignment(horizontal="center")
        ws2.cell(row=r, column=3, value=desc).alignment = Alignment(wrap_text=True, vertical="top")
        r += 1

    r += 1
    ws2.cell(row=r, column=1, value="四类商机怎么填").font = Font(bold=True, size=12, color="12224A")
    tips = [
        ("资本市场商机", "客户填上市公司；细分填「回购增持 / 可转债 / 定向增发 / IPO」；"
                   "回购区间 1-2 亿元 → 金额 10000、金额上限 20000；交易方式填「股份回购」。"),
        ("增存商机", "客户填融资企业全称；细分填「一级市场股权融资款」；交易方式填轮次（如 B 轮）；"
                 "交易对手填投资方；营销切入点写怎么切。「金额约亿元」按 10000 万元填。"),
        ("并购商机", "客户填竞买方（上市公司），交易对手填标的方；股权比例填数字（55）；"
                 "交易方式填「发行股份+现金 / 现金收购」；主办机构填竞买方主办行，备注写标的主办行/注册地。"),
        ("非房固贷商机", "客户填项目建设单位；商机名称填项目全称；金额填总投资（万元）；"
                   "信息来源填「发改委备案」；所在地填到区；营销切入点写融资方案。"),
    ]
    for name, txt in tips:
        r += 1
        ws2.cell(row=r, column=1, value=name).font = Font(bold=True, size=10)
        c = ws2.cell(row=r, column=3, value=txt)
        c.alignment = Alignment(wrap_text=True, vertical="top")

    r += 2
    ws2.cell(row=r, column=1, value="导入规则").font = Font(bold=True, size=12, color="12224A")
    rules = [
        "必填：商机类型、商机名称、客户/主体。其余可留空，系统给默认值（阶段=初步接洽，状态=进行中）。",
        "重复判定：商机名称 + 客户 完全相同的，导入时会被识别为重复，可选择跳过或覆盖更新。",
        "金额一律填万元，纯数字，不要带「万元」「亿」等文字（带也能识别，但建议纯数字）。",
        "日期格式：2026-10-19 或 2026/10/19 均可。",
        "一次最多导入 2000 行，单文件不超过 50MB。",
    ]
    for t in rules:
        r += 1
        c = ws2.cell(row=r, column=3, value="· " + t)
        c.alignment = Alignment(wrap_text=True, vertical="top")

    os.makedirs(os.path.dirname(TEMPLATE_PATH), exist_ok=True)
    wb.save(TEMPLATE_PATH)
    return TEMPLATE_PATH


@app.route("/deals/template")
def deals_template():
    if not os.path.isfile(TEMPLATE_PATH):
        build_deal_template()
    return send_file(TEMPLATE_PATH, as_attachment=True, download_name="商机导入模板.xlsx")


def _num(v):
    """把 2,000 / 2亿元 / 10000万元 / 10000 统一成万元数值"""
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "").replace("，", "")
    mult = 1
    if "亿" in s:
        mult = 10000
        s = s.replace("亿元", "").replace("亿", "")
    s = s.replace("万元", "").replace("万", "").replace("元", "")
    m = re.search(r"-?\d+(\.\d+)?", s)
    if not m:
        return None
    try:
        return float(m.group()) * mult
    except ValueError:
        return None


def _date(v):
    if v is None or v == "":
        return ""
    if isinstance(v, datetime.datetime):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, datetime.date):
        return v.strftime("%Y-%m-%d")
    s = str(v).strip().replace("/", "-").replace(".", "-")
    m = re.match(r"(\d{4})-?(\d{1,2})-?(\d{1,2})", s)
    if m:
        return "%04d-%02d-%02d" % (int(m.group(1)), int(m.group(2)), int(m.group(3)))
    m = re.match(r"^(\d{4})(\d{2})(\d{2})$", s)
    if m:
        return "%s-%s-%s" % (m.group(1), m.group(2), m.group(3))
    return ""


def _read_upload(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb["商机数据"] if "商机数据" in wb.sheetnames else wb[wb.sheetnames[0]]
        return [list(r) for r in ws.iter_rows(values_only=True)]
    if ext == ".csv":
        raw = open(path, "rb").read()
        text = None
        for enc in ("utf-8-sig", "gb18030", "utf-8"):
            try:
                text = raw.decode(enc)
                break
            except Exception:
                continue
        return [list(r) for r in csv.reader(io.StringIO(text or ""))]
    return []


def _parse_deals(rows):
    """解析并校验，返回 (records, summary)"""
    if not rows:
        return [], {"total": 0, "new": 0, "dup": 0, "error": 0}

    # 定位表头：前 5 行里能匹配到 3 个以上已知列名的行
    hidx, mapping = None, {}
    for i, r in enumerate(rows[:5]):
        m = {}
        for j, cell in enumerate(r):
            key = _norm_header(cell)
            if key in H2F:
                m[j] = H2F[key]
        if len(m) >= 3:
            hidx, mapping = i, m
            break
    if hidx is None:
        return [], {"total": 0, "new": 0, "dup": 0, "error": 1, "fatal": "表头没识别出来，请用下载的模板填写"}

    conn = get_db()
    exist = {}
    for r in conn.execute("SELECT id, name, client FROM deals"):
        exist[(_norm_header(r["name"]), _norm_header(r["client"]))] = r["id"]
    conn.close()

    records = []
    seen = {}
    for idx, r in enumerate(rows[hidx + 1:], start=hidx + 2):
        if not any(str(c).strip() for c in r if c is not None):
            continue
        rec = {f: "" for f in DEAL_FIELDS}
        raw = {}
        for j, field in mapping.items():
            val = r[j] if j < len(r) else ""
            raw[field] = val
            if field in ("amount", "amount2", "ratio"):
                rec[field] = _num(val)
            elif field == "notice_date":
                rec[field] = _date(val)
            else:
                rec[field] = str(val).strip() if val is not None else ""

        errs, warns = [], []
        if not rec["deal_type"]:
            errs.append("缺少「商机类型」")
        elif rec["deal_type"] not in DEAL_TYPES:
            errs.append("商机类型「%s」不在可选范围" % rec["deal_type"])
        if not rec["name"]:
            errs.append("缺少「商机名称」")
        if not rec["client"]:
            errs.append("缺少「客户/主体」")
        if not rec["stage"]:
            rec["stage"] = "初步接洽"
        elif rec["stage"] not in STAGES:
            warns.append("阶段「%s」不认识，已按初步接洽" % rec["stage"])
            rec["stage"] = "初步接洽"
        if not rec["status"]:
            rec["status"] = "进行中"
        elif rec["status"] not in STATUSES:
            warns.append("状态「%s」不认识，已按进行中" % rec["status"])
            rec["status"] = "进行中"
        if rec["amount"] is None:
            rec["amount"] = 0
        if rec["sub_type"] and rec["deal_type"] in DEAL_SUBTYPES:
            if rec["sub_type"] not in DEAL_SUBTYPES[rec["deal_type"]]:
                warns.append("细分类型「%s」不在该类型的标准选项里（不阻止导入）" % rec["sub_type"])

        key = (_norm_header(rec["name"]), _norm_header(rec["client"]))
        dup_id = exist.get(key)
        if not errs and key in seen:
            warns.append("与本表第 %d 行重复" % seen[key])
            dup_id = "in-file"
        elif not errs:
            seen[key] = idx
        status = "error" if errs else ("dup" if dup_id else "new")
        records.append({"row": idx, "data": rec, "status": status,
                        "dup_id": dup_id, "msg": "；".join(errs + warns), "raw": raw})

    summary = {
        "total": len(records),
        "new": sum(1 for x in records if x["status"] == "new"),
        "dup": sum(1 for x in records if x["status"] == "dup"),
        "error": sum(1 for x in records if x["status"] == "error"),
    }
    return records, summary


def _import_preview(pending, token, source, back="/deals/import", back_text="重新上传"):
    """导入预览页：Excel 导入与粘贴导入共用"""
    records = pending.get("records", [])
    summary = pending.get("summary", {})
    tag = {"new": ("p2", "新增"), "dup": ("p3", "重复"), "error": ("p4", "错误")}
    shown = records[:80]

    def _esc2(x):
        return str(x or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    def _preview_card(x):
        d = x["data"]
        t = d.get("deal_type") or "未分类"
        bits = []
        if d.get("client"):
            bits.append(_esc2(d["client"]))
        if d.get("counterparty"):
            bits.append("标的 " + _esc2(d["counterparty"]))
        if d.get("amount"):
            bits.append(format(d["amount"], ",.0f") + " 万")
        if d.get("ratio"):
            bits.append(("%g" % d["ratio"]) + "%")
        if d.get("org"):
            bits.append(_esc2(d["org"]))
        if d.get("notice_date"):
            bits.append(_esc2(d["notice_date"]))
        if d.get("location"):
            bits.append(_esc2(d["location"]))
        raw = (x.get("raw") or {}).get("原文") or d.get("note") or ""
        return (
            "<tr><td class='td-type'><span class='pill %s'>%s</span>%s"
            "<div style='margin-top:6px'><span class='pill %s'>%s</span></div></td>"
            "<td><div class='c-name'>%s</div>"
            "<div class='c-text'>%s</div>"
            "<div class='c-meta'>%s</div>%s</td></tr>"
        ) % (
            DEAL_TYPE_PILL.get(t, "p6"), _esc2(t),
            ("<div class='mini'>%s</div>" % _esc2(d.get("sub_type"))) if d.get("sub_type") else "",
            tag.get(x["status"], ("p6", ""))[0], tag.get(x["status"], ("p6", ""))[1],
            _esc2(d.get("name")),
            _esc2(raw) if raw else "<span class='mini'>（无原文）</span>",
            " ｜ ".join(bits) or "未识别到结构化字段",
            ("<div class='c-meta' style='color:#a12722'>%s</div>" % _esc2(x["msg"])) if x["msg"] else "")

    order = {t: i for i, t in enumerate(DEAL_TYPES)}
    shown.sort(key=lambda x: order.get(x["data"].get("deal_type") or "", 99))
    trs = "".join(_preview_card(x) for x in shown)

    more = ""
    if len(records) > len(shown):
        more = "<p class='mini'>仅预览前 80 条，共 %d 条，全部都会按规则处理。</p>" % len(records)

    body = """
    <div class="card"><h2>导入预览 <span class="mini">%s · %s</span></h2>
      <div class="kpis">
        <div class="kpi"><div class="t">数据行</div><div class="v">%d</div></div>
        <div class="kpi"><div class="t">可新增</div><div class="v" style="color:#1a7f4d">%d</div></div>
        <div class="kpi"><div class="t">重复</div><div class="v" style="color:#b97a12">%d</div></div>
        <div class="kpi"><div class="t">有错误</div><div class="v" style="color:#d0342c">%d</div></div>
      </div>
      %s
      <form method="post" action="/deals/import/commit">
        <input type="hidden" name="token" value="%s">
        <label style="display:flex;gap:8px;align-items:center;margin:6px 0">
          <input type="radio" name="mode" value="skip" checked style="width:auto">
          重复的跳过（只新增不重复的）</label>
        <label style="display:flex;gap:8px;align-items:center;margin:6px 0">
          <input type="radio" name="mode" value="update" style="width:auto">
          重复的覆盖更新（用新内容覆盖台账原记录）</label>
        <div style="margin-top:16px">
          <button>确认导入</button>
          <a class="btn btn3" href="%s" style="margin-left:8px">%s</a>
          <a class="btn btn3" href="/deals/import/errors/%s" style="margin-left:8px">⬇ 下载错误行</a>
          <a class="btn btn3" href="/deals" style="margin-left:8px">返回台账</a>
        </div>
      </form>
    </div>
    <div class="card"><h2>解析结果 <span class="mini">商机类别 / 具体内容</span></h2>
      <div class="tbl-wrap"><table class="two-col">
        <tr><th>商机类别</th><th>具体内容</th></tr>
        %s</table></div>%s</div>
    """ % (source, pending.get("ts", ""), summary.get("total", 0), summary.get("new", 0),
           summary.get("dup", 0), summary.get("error", 0),
           '<div class="err">有 %d 条存在问题，导入时会自动跳过。</div>' % summary.get("error", 0)
           if summary.get("error") else "", token, back, back_text, token, trs, more)
    return page(body, "deals")


PASTE_DEMO = open(os.path.join(BASE_DIR, "粘贴导入示例.txt"), encoding="utf-8").read() \
    if os.path.isfile(os.path.join(BASE_DIR, "粘贴导入示例.txt")) else ""


@app.route("/deals/paste", methods=["GET", "POST"])
def deals_paste():
    if request.method == "POST":
        text = (request.form.get("content") or "").strip()
        if not text:
            return page('<div class="err">请先粘贴商机内容。</div><a class="btn" href="/deals/paste">返回</a>', "deals")
        records, summary = parse_paste_text(text)
        if summary.get("fatal"):
            return page('<div class="err">%s</div><a class="btn" href="/deals/paste">重新粘贴</a>'
                        % summary["fatal"], "deals")
        token = uuid.uuid4().hex[:12]
        pending = {"file": "粘贴文本（%d 字）" % len(text),
                   "ts": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
                   "records": records, "summary": summary, "mode": "paste"}
        with open(os.path.join(IMPORT_DIR, token + ".json"), "w", encoding="utf-8") as fp:
            json.dump(pending, fp, ensure_ascii=False)
        return _import_preview(pending, token, "粘贴导入", "/deals/paste", "重新粘贴")

    demo = request.args.get("demo")
    content = PASTE_DEMO if demo else ""
    q = (request.args.get("q") or "").strip()

    # 查询：先查已导入/已存在的商机
    qres = []
    if q:
        conn = get_db()
        qres = conn.execute(
            "SELECT * FROM deals WHERE name LIKE ? OR client LIKE ? OR counterparty LIKE ?"
            " OR note LIKE ? OR org LIKE ? ORDER BY updated DESC LIMIT 30",
            ["%%%s%%" % q] * 5).fetchall()
        conn.close()

    qhtml = ""
    if q:
        rows = "".join(
            "<tr><td><b>%s</b><div class='mini'>%s</div></td><td>%s</td><td>%s</td>"
            "<td>%s</td><td class='mini'>%s</td></tr>"
            % (r["name"], (r["deal_type"] or "未分类"), (r["client"] or "—"),
               ("%s 万" % format(r["amount"], ",.0f")) if r["amount"] else "—",
               (r["org"] or "—"), (r["updated"] or ""))
            for r in qres) or "<tr><td class='mini' colspan=5>没有找到匹配的商机</td></tr>"
        qhtml = """
        <div class="card"><h2>查询结果 <span class="mini">「%s」共 %d 条</span></h2>
          <div class="tbl-wrap"><table>
            <tr><th>商机 / 类型</th><th>客户</th><th>金额</th><th>主办机构</th><th>更新时间</th></tr>
            %s</table></div>
          <div style="margin-top:12px"><a class="btn btn3 btn-sm" href="/deals?q=%s">在台账中查看</a></div>
        </div>""" % (q, len(qres), rows, q)

    body = """
    <div class="card"><h2>粘贴导入 <span class="mini">自动识别商机类型与字段</span></h2>
      <div class="note">把群里/邮件里的商机清单整段粘进来即可。支持【资本市场商机】【增存商机】这类分节，
        也支持 1. 2. 3. 编号；系统会自动判断类型，抽取客户、金额、股权比例、主办机构、日期、所在地等字段。
        解析后先预览，确认无误再入库。</div>
      <form method="post">
        <textarea name="content" rows="14" style="width:100%%;font-size:13px;line-height:1.9;
          padding:12px;border:1px solid #dde3ec;border-radius:8px;font-family:'Microsoft YaHei',sans-serif"
          placeholder="示例：&#10;【资本市场商机】&#10;1.东莞分行客户XX科技拟回购1-2亿元；&#10;【并购商机】&#10;1、XX股份有限公司拟以支付现金方式收购YY有限公司100%%股权，交易价格不超过5.5亿元。">%s</textarea>
        <div style="margin-top:14px;display:flex;gap:8px;flex-wrap:wrap">
          <button>解析并预览</button>
          <a class="btn btn3" href="/deals/paste?demo=1">填入示例</a>
          <a class="btn btn3" href="/deals/paste">清空</a>
          <a class="btn btn3" href="/deals">返回台账</a>
        </div>
      </form>
      <div class="mini" style="margin-top:14px;line-height:1.9">
        识别规则：类型看关键词（回购/增持→资本市场；并购/收购→并购；融资/轮次→增存；项目/备案/总投资→非房固贷），
        有【】标题时以标题为准；金额自动换算成万元（1-2 亿 → 10000~20000）。
        识别不到客户名称的条目标红，可下载错误行手工补。
      </div>
    </div>

    <div class="card"><h2>查询已有商机 <span class="mini">导入前先查重</span></h2>
      <form method="get" style="display:flex;gap:8px;flex-wrap:wrap">
        <input name="q" value="%s" placeholder="输入企业名 / 商机名 / 标的 / 主办机构" style="flex:1;min-width:220px">
        <button class="btn2">查询</button>
      </form>
    </div>
    %s""" % (content, q, qhtml)
    return page(body, "deals")


@app.route("/deals/import", methods=["GET", "POST"])
def deals_import():
    if request.method == "GET":
        body = """
        <div class="card"><h2>批量导入商机</h2>
          <div class="note">支持 Excel（.xlsx）与 CSV。请先下载模板，按「商机数据」表填写，
            示例行请删除。导入前会先预览校验结果，确认后才会写入台账。</div>
          <div class="row" style="align-items:flex-end">
            <div style="flex:2">
              <form method="post" enctype="multipart/form-data">
                <label>选择文件</label>
                <input type="file" name="file" accept=".xlsx,.xlsm,.csv" required>
                <div style="margin-top:14px"><button>上传并校验</button>
                  <a class="btn btn3" href="/deals/template" style="margin-left:8px">⬇ 下载导入模板</a>
                  <a class="btn btn3" href="/deals" style="margin-left:8px">返回台账</a></div>
              </form>
            </div>
            <div style="flex:1">
              <div class="kpi"><div class="t">四类商机</div>
                <div style="font-size:12.5px;line-height:2;margin-top:6px">
                  ① 资本市场商机（回购增持等）<br>② 增存商机（一级市场股权融资款）<br>
                  ③ 并购商机<br>④ 固贷商机
                </div></div>
            </div>
          </div>
        </div>"""
        return page(body, "deals")

    f = request.files.get("file")
    if not f or not f.filename:
        return redirect("/deals/import")
    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in (".xlsx", ".xlsm", ".csv"):
        return page('<div class="err">只支持 .xlsx / .xlsm / .csv 文件</div>', "deals")

    token = uuid.uuid4().hex[:12]
    src_path = os.path.join(IMPORT_DIR, token + ext)
    f.save(src_path)
    rows = _read_upload(src_path)
    records, summary = _parse_deals(rows)

    pending = {"file": f.filename, "ts": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
               "records": records, "summary": summary}
    with open(os.path.join(IMPORT_DIR, token + ".json"), "w", encoding="utf-8") as fp:
        json.dump(pending, fp, ensure_ascii=False)
    try:
        os.remove(src_path)
    except OSError:
        pass          # 某些受控环境不允许删除文件，保留也不影响

    if summary.get("fatal"):
        return page('<div class="err">%s</div><a class="btn" href="/deals/import">重新上传</a>'
                    % summary["fatal"], "deals")

    tag = {"new": ("p2", "新增"), "dup": ("p3", "重复"), "error": ("p4", "错误")}
    shown = records[:80]
    trs = "".join(
        "<tr><td class='mini'>%d</td><td>%s</td><td>%s</td><td>%s</td><td>%s</td>"
        "<td><span class='pill %s'>%s</span></td><td class='mini'>%s</td></tr>"
        % (x["row"], x["data"].get("deal_type", ""), x["data"].get("name", "") or "—",
           x["data"].get("client", "") or "—",
           ("%s 万" % format(x["data"]["amount"], ",.0f")) if x["data"].get("amount") else "—",
           tag[x["status"]][0], tag[x["status"]][1], x["msg"] or "")
        for x in shown)

    more = ""
    if len(records) > len(shown):
        more = "<p class='mini'>仅预览前 80 行，共 %d 行，全部行都会按规则处理。</p>" % len(records)

    body = """
    <div class="card"><h2>导入预览 <span class="mini">%s · %s</span></h2>
      <div class="kpis">
        <div class="kpi"><div class="t">数据行</div><div class="v">%d</div></div>
        <div class="kpi"><div class="t">可新增</div><div class="v" style="color:#1a7f4d">%d</div></div>
        <div class="kpi"><div class="t">重复</div><div class="v" style="color:#b97a12">%d</div></div>
        <div class="kpi"><div class="t">有错误</div><div class="v" style="color:#d0342c">%d</div></div>
      </div>
      %s
      <form method="post" action="/deals/import/commit">
        <input type="hidden" name="token" value="%s">
        <label style="display:flex;gap:8px;align-items:center;margin:6px 0">
          <input type="radio" name="mode" value="skip" checked style="width:auto">
          重复的跳过（只新增不重复的）</label>
        <label style="display:flex;gap:8px;align-items:center;margin:6px 0">
          <input type="radio" name="mode" value="update" style="width:auto">
          重复的覆盖更新（用表里的内容覆盖台账原记录）</label>
        <div style="margin-top:16px">
          <button>确认导入</button>
          <a class="btn btn3" href="/deals/import" style="margin-left:8px">重新上传</a>
          <a class="btn btn3" href="/deals/import/errors/%s" style="margin-left:8px">⬇ 下载错误行</a>
        </div>
      </form>
    </div>
    <div class="card"><h2>明细</h2><div class="tbl-wrap"><table>
      <tr><th>行号</th><th>类型</th><th>商机名称</th><th>客户/主体</th><th>金额</th><th>状态</th><th>说明</th></tr>
      %s</table></div>%s</div>
    """ % (pending["file"], pending["ts"], summary["total"], summary["new"], summary["dup"],
           summary["error"],
           '<div class="err">有 %d 行存在错误，这些行会在导入时自动跳过。</div>' % summary["error"]
           if summary["error"] else "", token, token, trs, more)
    return _import_preview(pending, token, "文件：" + pending["file"])


@app.route("/deals/import/commit", methods=["POST"])
def deals_import_commit():
    token = request.form.get("token", "")
    mode = request.form.get("mode", "skip")
    path = os.path.join(IMPORT_DIR, token + ".json")
    if not os.path.isfile(path):
        return page('<div class="err">导入会话已失效，请重新上传。</div>', "deals")
    with open(path, encoding="utf-8") as fp:
        pending = json.load(fp)
    records = pending.get("records", [])

    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    cols = DEAL_FIELDS + ["created", "updated"]
    placeholders = ",".join(["?"] * len(cols))
    conn = get_db()
    n_new = n_skip = n_err = 0
    for x in records:
        if x["status"] == "error":
            n_err += 1
            continue
        if x["status"] == "dup" and (mode == "skip" or not x.get("dup_id")
                                     or x.get("dup_id") == "in-file"):
            n_skip += 1
            continue
        d = x["data"]
        if not d.get("deal_date"):
            d["deal_date"] = datetime.date.today().strftime("%Y-%m-%d")
        vals = [d.get(f) if d.get(f) != "" else None for f in DEAL_FIELDS]
        vals = [0 if (f == "amount" and v is None) else v
                for f, v in zip(DEAL_FIELDS, vals)]
        if x["status"] == "dup" and mode == "update" and x.get("dup_id"):
            sets = ",".join("%s=?" % f for f in DEAL_FIELDS)
            conn.execute("UPDATE deals SET %s, updated=? WHERE id=?" % sets,
                         vals + [now, x["dup_id"]])
            n_new += 0
        else:
            conn.execute("INSERT INTO deals (%s) VALUES (%s)" % (",".join(cols), placeholders),
                         vals + [now, now])
            n_new += 1
            # 固贷商机：同步写入固贷台账（供综合融资菜单展示）
            g = d.get("_gudai")
            if d.get("deal_type") == "非房固贷商机" and isinstance(g, dict):
                conn.execute(
                    "INSERT INTO gudai_deals (push_date, unit, name, region, amount, content, "
                    "manager, note, approval_amount, created, updated) VALUES (?,?,?,?,?,?,?,?,?,?,?)",
                    (g.get("push_date") or "", g.get("unit") or "", g.get("name") or "",
                     g.get("region") or "", g.get("amount"), g.get("content") or "",
                     g.get("manager") or "", g.get("note") or "", g.get("approval_amount"),
                     now, now))
    conn.commit()
    conn.close()
    try:
        os.remove(path)
    except OSError:
        pass
    return redirect("/deals?msg=import:%d,%d,%d" % (n_new, n_skip, n_err))


@app.route("/deals/import/errors/<token>")
def deals_import_errors(token):
    path = os.path.join(IMPORT_DIR, token + ".json")
    if not os.path.isfile(path):
        return page('<div class="err">导入会话已失效。</div>', "deals")
    with open(path, encoding="utf-8") as fp:
        pending = json.load(fp)
    bad = [x for x in pending.get("records", []) if x["status"] != "new"]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "待处理行"
    ws.append(["行号", "状态", "说明"] + IMPORT_HEADERS)
    for x in bad:
        d = x["data"]
        ws.append([x["row"], "重复" if x["status"] == "dup" else "错误", x["msg"]] +
                  [d.get(f, "") for _h, f, _r, _d in IMPORT_COLUMNS])
    for i in range(1, len(IMPORT_HEADERS) + 4):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = 18
    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return send_file(bio, as_attachment=True, download_name="导入问题行.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ----------------------------------------------------------------------------
# 2.2 粘贴文本导入：自动分类 + 字段抽取
# ----------------------------------------------------------------------------
TYPE_KEYWORDS = [
    ("并购商机", ["并购", "收购", "购买", "股权转让", "要约收购", "竞买方"]),
    ("非房固贷商机", ["固贷", "固定资产", "项目贷", "总投资", "发改委", "拟建设"]),
    ("增存商机", ["股权融资", "增存", "投资方", "B轮", "A轮", "C轮", "融资款", "募集资金"]),
    ("资本市场商机", ["回购", "增持", "定增", "定向增发", "可转债", "IPO"]),
]

SECTOR_HINTS = [
    (["光掩模", "半导体", "芯片", "集成电路", "晶圆"], "半导体"),
    (["医疗器械", "手术机器人", "医药", "生物", "CDMO", "创新药"], "生物医药/医疗器械"),
    (["大数据", "算力", "智算", "数据中心", "电算"], "大数据/算力"),
    (["新材料", "化工", "改性塑料"], "化工新材料"),
    (["覆铜板", "线路板", "电子材料", "电子"], "电子材料"),
    (["装备", "海洋工程", "机械", "机电"], "高端装备"),
    (["新能源", "光伏", "储能", "锂电"], "新能源"),
]

CN_NUM = "①②③④⑤⑥⑦⑧⑨⑩"
ORG_STOP = ["竞买方主办行为", "标的主办行为", "主办行为", "主办行", "客户", "位于", "由", "为", "在", "系"]


def _flat(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()


def _guess_type(text):
    for dtype, kws in TYPE_KEYWORDS:
        if any(k in text for k in kws):
            return dtype
    return ""


def _type_from_header(h):
    h = h or ""
    if "资本" in h or "回购" in h or "增持" in h:
        return "资本市场商机", ""
    if "增存" in h:
        return "增存商机", ("一级市场股权融资款" if "股权融资" in h else "")
    if "并购" in h:
        return "并购商机", ""
    if "固贷" in h or "发改委" in h or "建设" in h:
        return "非房固贷商机", ("发改委备案项目" if "备案" in h else "")
    return "", ""


def _orgs(t):
    out = []
    for m in re.finditer(r"[\u4e00-\u9fa5]{2,6}(?:分行|支行)", t):
        s = m.group(0)
        changed = True
        while changed:
            changed = False
            for p in ORG_STOP:
                if s.startswith(p) and len(s) - len(p) >= 3:
                    s, changed = s[len(p):], True
        if len(s) >= 4 and s not in out:
            out.append(s)
    return out


COMPANY_VERBS = ["拟通过", "拟以", "拟", "将通过", "通过", "以", "及", "发行股份及支付现金",
                 "发行股份", "支付现金", "现金", "方式", "购买", "收购", "标的为", "标的公司", "向", "其", "的"]


def _trim_company(name):
    s = name
    changed = True
    while changed:
        changed = False
        for v in COMPANY_VERBS:
            if s.startswith(v) and len(s) - len(v) >= 4:
                s, changed = s[len(v):], True
    return s


def _companies(t):
    """企业全称（去掉分行/支行前缀、动词前缀；不允许跨括号）"""
    t2 = re.sub(r"[一-龥]{2,10}(?:分行|支行)(?:客户)?", " ", t)
    found = re.findall(r"[一-龥A-Za-z0-9·]{4,40}?(?:股份有限公司|有限责任公司|有限公司|集团)", t2)
    out = []
    for f in found:
        f = _trim_company(re.sub(r"[（(][^）)]*[）)]", "", f).strip())
        if len(f) >= 4 and f not in out:
            out.append(f)
    return out


def _fallback_company(t):
    """没有「XX有限公司」后缀时，取动词前的主体名（如「生益科技拟回购」→ 生益科技）"""
    s = re.sub(r"^.*?(?:分行|支行)客户", "", t).strip()
    m = re.match(r"([\u4e00-\u9fa5A-Za-z0-9·]{2,15}?)(?=拟|计划|董事长|将|于|公告|回购|增持|收购|购买|建设|投建)", s)
    if m:
        return m.group(1)
    m = re.match(r"([\u4e00-\u9fa5A-Za-z0-9·]{2,15}?)(?:回购|增持)", s)
    return m.group(1) if m else ""


def _nickname(t):
    m = re.search(r"简称[“\"「]([^”\"」]{2,10})[”\"」]", t)
    return m.group(1) if m else ""


def _short(name):
    if not name:
        return ""
    s = re.sub(r"(股份有限公司|有限责任公司|有限公司|集团)$", "", name)
    s = re.sub(r"^(广州市|广州|深圳市|深圳|东莞市|东莞|佛山市|佛山|珠海市|珠海|惠州市|惠州)", "", s)
    return s or name


def _money(t):
    m = re.search(r"(\d+(?:\.\d+)?)\s*[-~至到]\s*(\d+(?:\.\d+)?)\s*亿", t)
    if m:
        return round(float(m.group(1)) * 10000, 2), round(float(m.group(2)) * 10000, 2)
    if re.search(r"约\s*亿元", t):
        return 10000.0, None
    m = re.search(r"(\d+(?:\.\d+)?)\s*亿", t)
    if m:
        return round(float(m.group(1)) * 10000, 2), None
    m = re.search(r"(\d+(?:\.\d+)?)\s*万", t)
    if m:
        return round(float(m.group(1)), 2), None
    return None, None


def _ratio(t):
    m = re.search(r"(\d+(?:\.\d+)?)\s*%\s*(?:的)?股权", t)
    if not m:
        m = re.search(r"(\d+(?:\.\d+)?)\s*%", t)
    return float(m.group(1)) if m else None


def _notice_date(t):
    m = re.search(r"(\d{1,2})\s*月\s*(\d{1,2})\s*日", t)
    if m:
        return "%d-%02d-%02d" % (datetime.datetime.now().year, int(m.group(1)), int(m.group(2)))
    m = re.search(r"(\d{4})[-/年](\d{1,2})[-/月](\d{1,2})", t)
    if m:
        return "%04d-%02d-%02d" % (int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return ""


def _location(t):
    m = re.search(r"(?:项目地?位于|位于|地处|注册地在|坐落在)([\u4e00-\u9fa5]{2,15}?)(?=[，。；,;]|$)", t)
    if m:
        return m.group(1)
    m = re.search(r"([\u4e00-\u9fa5]{2,6}市[\u4e00-\u9fa5]{0,10}(?:新区|区|县))", t)
    return m.group(1) if m else ""


def _sector(t):
    for kws, name in SECTOR_HINTS:
        if any(k in t for k in kws):
            return name
    return ""


def _method(t, dtype):
    if dtype == "资本市场商机":
        if "回购" in t:
            return "股份回购"
        if "增持" in t:
            return "股东增持"
        if "可转债" in t:
            return "可转债"
        if "定增" in t:
            return "定向增发"
    if dtype == "增存商机":
        m = re.search(r"((?:Pre-?IPO|A\+?|B\+?|C\+?|D|天使|战略)\s*轮)", t, re.I)
        return (m.group(1).replace(" ", "") + "股权融资") if m else "股权融资"
    if dtype == "并购商机":
        if "发行股份" in t and "现金" in t:
            return "发行股份+现金"
        if "发行股份" in t:
            return "发行股份购买资产"
        if "现金" in t:
            return "现金收购"
        return "股权收购"
    if dtype == "非房固贷商机":
        return "固定资产贷款"
    return ""


def _sub_from_method(dtype, method):
    if dtype == "并购商机":
        if method == "现金收购":
            return "现金收购"
        if method in ("发行股份+现金", "发行股份购买资产"):
            return "发行股份购买资产"
    return ""


def _project_name(t):
    m = re.search(r"拟建设(.{4,60}?项目(?:[（(][^）)]*[）)])?)", t)
    if m:
        return m.group(1)
    m = re.search(r"(?:建设|投建)(.{4,50}?(?:项目|工程))", t)
    return m.group(1) if m else ""


def _build_name(dtype, rec, comps, t):
    nick = _nickname(t)
    if dtype == "资本市场商机":
        base = nick or _short(rec["client"]) or (_short(comps[0]) if comps else "")
        act = "股份回购" if "回购" in t else ("股东增持" if "增持" in t else "")
        return (base + act) if base and act else ""
    if dtype == "增存商机":
        base = nick or _short(rec["client"]) or (_short(comps[0]) if comps else "")
        m = re.search(r"((?:Pre-?IPO|A\+?|B\+?|C\+?|D|天使|战略)\s*轮)", t, re.I)
        rnd = m.group(1).replace(" ", "") if m else ""
        return (base + rnd + "股权融资款承接") if base else ""
    if dtype == "并购商机":
        buyer = _short(comps[0]) if comps else ""
        seller = _short(comps[1]) if len(comps) > 1 else ""
        r = rec.get("ratio")
        rtxt = (("%g" % r) + "%") if r else ""
        if buyer and seller:
            return "%s收购%s%s股权" % (buyer, seller, rtxt)
        if buyer:
            return buyer + "并购项目"
        return ""
    if dtype == "非房固贷商机":
        pn = _project_name(t)
        if pn:
            return pn
        return (_short(comps[0]) + "固定资产贷款项目") if comps else ""
    return ""


def _segment(text):
    lines = [x.rstrip() for x in str(text or "").replace("\r", "").split("\n")]
    sections, cur = [], None
    for raw in lines:
        line = _flat(raw)
        m = re.search(r"【([^】]{2,40})】", line)
        if m:
            dtype, sub = _type_from_header(m.group(1))
            cur = {"type": dtype, "sub": sub, "head": m.group(1), "lines": []}
            sections.append(cur)
            rest = line[m.end():].strip(" 　:：")
            if rest:
                cur["lines"].append(rest)
            continue
        if not line:
            continue
        if cur is None:
            cur = {"type": "", "sub": "", "head": "", "lines": []}
            sections.append(cur)
        cur["lines"].append(line)
    return sections


def _split_items(lines):
    items, buf = [], []
    pat = re.compile(r"^(?:\d{1,2}\s*[、.．)）]|[%s])\s*" % CN_NUM)
    if not any(pat.match(x) for x in lines):
        return [" ".join(lines)] if lines else []
    for line in lines:
        if pat.match(line):
            if buf:
                items.append(" ".join(buf))
            buf = [pat.sub("", line).strip()]
        else:
            buf.append(line)
    if buf:
        items.append(" ".join(buf))
    return items


def parse_gudai_fields(text, push_date=""):
    """按固贷商机规则解析 9 个结构化字段（推送时间/经营单位/项目名字/区域/
    项目金额/商机情况/产品经理/营销难点/报批金额）"""
    g = {"push_date": push_date or "", "unit": "", "name": "", "region": "",
         "amount": None, "content": text, "manager": "",
         "note": "", "approval_amount": None}
    # 推送时间：优先从文本里括号 X月X日 提取
    m = re.search(r"[（(](\d{1,2})月(\d{1,2})日[）)]", text)
    if m:
        g["push_date"] = "%s月%s日" % (m.group(1), m.group(2))
    # 经营单位：「客户」之前的字段
    m = re.search(r"([^\s，。；、:：]{2,15})客户", text)
    if m:
        g["unit"] = m.group(1).strip()
    # 项目名字：第一个「拟建设」之后，到「拟建设含」或逗号/句号之前
    m = re.search(r"拟建设([^，。；、\n]{2,80}?)(?:，拟建设含|拟建设含)", text)
    if not m:
        m = re.search(r"拟建设([^，。；、\n]{2,80}?)[，。；]", text)
    if m:
        g["name"] = m.group(1).strip()
    # 区域：「项目地位于」之后，取到「市」
    m = re.search(r"项目地位于([^，。；、\n]{0,30}?市)", text)
    if m:
        g["region"] = m.group(1).strip()
    # 项目金额（亿元）：「总投资」之后的数字
    m = re.search(r"总投资\s*([\d.]+)\s*亿", text)
    if m:
        try:
            g["amount"] = float(m.group(1))
        except ValueError:
            pass
    # 产品经理：经营单位 → 分工表匹配
    if g["unit"]:
        g["manager"] = GUDAI_DIVISION.get(g["unit"], "")
    return g


def parse_paste_text(text):
    """把粘贴的文字解析成导入记录（结构与 Excel 导入一致）"""
    sections = _segment(text)
    records = []

    conn = get_db()
    exist = {}
    for r in conn.execute("SELECT id, name, client FROM deals"):
        exist[(_norm_header(r["name"]), _norm_header(r["client"]))] = r["id"]
    conn.close()

    seen = {}
    rowno = 0
    cur_push_date = ""
    for sec in sections:
        dtype, sub, head = sec["type"], sec["sub"], sec["head"]
        body_lines = []
        for line in sec["lines"]:
            m = re.match(r"^(回购增持|回购|增持|定向增发|可转债|股权融资|并购|收购|固定资产贷款|固贷)\s*[:：]?$", line)
            if m and not body_lines:
                sub = sub or ("回购增持" if "回购" in m.group(1) else m.group(1))
                continue
            # 推送日期行：如（10月29日），作为后续固贷商机的推送时间
            dm = re.search(r"[（(](\d{1,2})月(\d{1,2})日[）)]", line)
            if dm:
                cur_push_date = "%s月%s日" % (dm.group(1), dm.group(2))
            body_lines.append(line)

        for item in _split_items(body_lines):
            item = _flat(item).strip("；;。 ")
            if len(item) < 6:
                continue
            rowno += 1
            t = item
            comps = _companies(t)
            rec = {f: "" for f in DEAL_FIELDS}
            rec["deal_type"] = dtype or _guess_type(t)
            rec["client"] = comps[0] if comps else _fallback_company(t)
            if not comps and rec["client"]:
                comps = [rec["client"]]
            if rec["deal_type"] == "并购商机" and len(comps) > 1:
                rec["counterparty"] = comps[1]
            elif rec["deal_type"] == "增存商机":
                inv = re.search(r"投资方为([^，。；]{2,40})", t)
                if inv:
                    rec["counterparty"] = inv.group(1)

            amt, amt2 = _money(t)
            rec["amount"] = amt
            rec["amount2"] = amt2
            rec["ratio"] = _ratio(t)
            rec["sector"] = _sector(t)
            rec["method"] = _method(t, rec["deal_type"])
            rec["sub_type"] = sub or _sub_from_method(rec["deal_type"], rec["method"]) \
                or (DEAL_SUBTYPES.get(rec["deal_type"], [""])[0] if rec["deal_type"] else "")
            rec["notice_date"] = _notice_date(t)
            rec["location"] = _location(t)
            orgs = _orgs(t)
            rec["org"] = orgs[0] if orgs else ""
            rec["stage"] = "初步接洽"
            rec["status"] = "进行中"
            if "备案" in t or "备案" in head:
                rec["source"] = "发改委备案"
            elif "公告" in t:
                rec["source"] = "上市公司公告"
            elif orgs:
                rec["source"] = "分行报送"
            rec["name"] = _build_name(rec["deal_type"], rec, comps, t)
            if not rec["name"]:
                rec["name"] = t[:28]
            entry = re.search(r"(请[^。]{6,60}。?|积极[^。]{6,60}。?)", t)
            rec["entry"] = entry.group(1) if entry else ""
            note = item
            if len(orgs) > 1:
                note += "（相关机构：%s）" % "、".join(orgs[1:])
            rec["note"] = note
            if rec["deal_type"] == "非房固贷商机":
                rec["_gudai"] = parse_gudai_fields(t, cur_push_date)

            errs = []
            if not rec["deal_type"]:
                errs.append("无法判断商机类型")
            if not rec["client"]:
                errs.append("未能识别客户/主体名称")

            key = (_norm_header(rec["name"]), _norm_header(rec["client"]))
            dup_id = exist.get(key)
            if not errs and key in seen:
                dup_id = "in-file"
                errs.append("与第 %d 条重复" % seen[key])
            elif not errs and dup_id:
                errs.append("台账中已存在同名同客户记录")

            if not rec["deal_type"] or not rec["client"]:
                status = "error"
            elif dup_id:
                status = "dup"
            else:
                status = "new"
            seen.setdefault(key, rowno)
            records.append({"row": rowno, "data": rec, "status": status,
                            "dup_id": dup_id if dup_id != "in-file" else "in-file",
                            "msg": "；".join(errs), "raw": {"原文": item}})

    summary = {
        "total": len(records),
        "new": sum(1 for x in records if x["status"] == "new"),
        "dup": sum(1 for x in records if x["status"] == "dup"),
        "error": sum(1 for x in records if x["status"] == "error"),
    }
    if not records:
        summary["fatal"] = "没解析出内容，请确认粘贴的是商机清单（含【】分节或 1. 2. 编号）"
    return records, summary


# ----------------------------------------------------------------------------
if __name__ == "__main__":
    init_db()
    if get_setting("material_root") and os.path.isdir(get_setting("material_root")):
        scan_materials(get_setting("material_root"), trigger="启动")
    threading.Thread(target=auto_scan_worker, daemon=True).start()
    print("投行数智化工作台已启动 → http://127.0.0.1:5000")
    print("每日自动扫描：已启用（%s），线程已启动" % get_setting("auto_scan_time", "08:30"))
    app.run(host="0.0.0.0", port=int(os.environ.get("WB_PORT", "5000")), debug=False)
