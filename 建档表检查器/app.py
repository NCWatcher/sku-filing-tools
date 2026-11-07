# -*- coding: utf-8 -*-
"""
建档表检查器 v1.0
读一份「拼多多 17 列建档表」，对照编码库（可选）逐行体检，把错格挑出来并出报告。

检查它，不生成它 —— 生成用 工具-SKU建档生成\\SKU建档生成器.exe。

命令行（出问题时可以这样跑，不用开界面）：
    python app.py --check 建档.xlsx [--lib 款式编码文档.xlsx] [--out 报告.xlsx] [--json]
"""
import copy as _copy
import datetime
import glob
import json
import os
import re
import subprocess
import sys
import traceback

import openpyxl
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

APP_NAME = '建档表检查器'
VER = 'v2.0'
GAP_DEFAULT = 10

# v3.1 的默认位置（没有就用空，让用户自己选）
DEFAULT_LIB = r'款式编码文档-模板.xlsx'
DEFAULT_ARCHIVE = r'建档库'

# ── 17 列模板（与生成器完全一致） ────────────────────────────────────────────
HEAD = ['组合商品编码', '组合装商品标签', '组合款式编码', '组合商品名称', '组合商品简称',
        '组合商品实体编码', '虚拟分类', '组合颜色规格', '禁止库存同步', '商品编码',
        '数量', '应占售价', '组合成本价', '是否支持留言备注换货', '图片', '品牌', None]
HEAD_NAME = ['组合商品编码', '组合装商品标签', '组合款式编码', '组合商品名称', '组合商品简称',
             '组合商品实体编码', '虚拟分类', '组合颜色规格', '禁止库存同步', '商品编码',
             '数量', '应占售价', '组合成本价', '是否支持留言备注换货', '图片', '品牌', '（Q列·无表头）']
RED_HEAD = {1, 3, 10, 11, 12}
WIDTH = {'A': 39.375, 'B': 23.875, 'C': 39.375, 'D': 9.625, 'F': 12.625,
         'G': 6.625, 'H': 8.125, 'J': 21.625, 'K': 3.625, 'L': 6.625,
         'M': 8.125, 'N': 15.625, 'O': 3.625, 'Q': 9.0}
SZ_LETTER = ['XS', 'S', 'M', 'L', 'XL', '2XL', '3XL', '4XL', '5XL', '6XL']
HEIGHTS = ['80', '90', '100', '110', '120', '130', '140', '150', '160', '170', '180']

# ── 三种编码族（正则与生成器一字不差） ──────────────────────────────────────
PAT = re.compile(r'^([A-Z]+)(\d{3,7})([A-Z]*)T(\d)S(\d)(.+)$')
HEI = re.compile(r'^([A-Z][A-Z0-9]*?)(\d{3})$')
VEST = re.compile(r'^([A-Z]+)(\d{3,5})([A-Z]{2,6}?)(XXXL|XXL|XL|XS|S|M|L)$')
CJK = re.compile(r'[\u4e00-\u9fa5]')

ERR, WARN, INFO, OK = '错误', '警告', '提示', '通过'
LEVEL_ORDER = {ERR: 0, WARN: 1, INFO: 2, OK: 3}
FILL_OF = {ERR: 'FFC7CE', WARN: 'FFF2CC', INFO: 'DDEBF7'}
FONT_OF = {ERR: '9C0006', WARN: '9C6500', INFO: '1F4E78'}

Q_TIP = ('内裤型 Q 列应为 =LEFT(_xlfn.TEXTAFTER(J,"T"),1)*K，'
         '身高型/背心型（编码里没有 T 段）应直接写 1')


# ══════════════════════════════════════════════════════════════════════════
#  基础解析
# ══════════════════════════════════════════════════════════════════════════
def norm_size(s):
    s = str(s).strip().upper()
    m = re.match(r'^(\d+)XL$', s)
    if m:
        n = int(m.group(1))
        return 'XL' if n == 1 else '%dXL' % n
    m = re.match(r'^(X+)L$', s)
    if m:
        n = len(m.group(1))
        return 'XL' if n == 1 else '%dXL' % n
    return s


def size_key(s):
    n = norm_size(s)
    if n in SZ_LETTER:
        return (0, SZ_LETTER.index(n))
    if n.isdigit() and n in HEIGHTS:
        return (1, int(n))
    return (2, 0)


def looks_like_size(t):
    u = str(t).strip().upper()
    if u in ('M', 'S', 'XS') or u in HEIGHTS:
        return True
    return bool(re.match(r'^(\d+)?X*L$', u))


def parse_code(code):
    """返回 (货号, 尺码, 族)；族 = inner / height / vest / None"""
    code = str(code).strip()
    m = PAT.match(code)
    if m:
        return m.group(2), norm_size(m.group(6)), 'inner'
    m = HEI.match(code)
    if m and m.group(2) in HEIGHTS:
        return m.group(1), m.group(2), 'height'
    m = VEST.match(code)
    if m:
        return m.group(2), norm_size(m.group(4)), 'vest'
    return None, None, None


def t_count(code):
    m = PAT.match(code or '')
    return int(m.group(4)) if m else 0


def q_risk(code):
    """内裤型编码里 Q 公式会不会算歪，返回原因或 None"""
    if not code:
        return '没有编码'
    i = code.find('T')
    if i < 0:
        return '编码里没有 T 段'
    if not code[i + 1:i + 2].isdigit():
        return '编码里有第二个 T（花色码以 T 结尾），TEXTAFTER 会取到字母'
    if not re.match(r'^T\dS\d', code[i:]):
        return 'T 段后面不是 S+数字'
    return None


def fmt_size(s):
    return norm_size(s)


def norm_formula(f):
    return re.sub(r'\s+', '', str(f or '')).upper().replace('$', '')


def text(v):
    return '' if v is None else str(v).strip()


# ══════════════════════════════════════════════════════════════════════════
#  结果收集
# ══════════════════════════════════════════════════════════════════════════
class Report:
    def __init__(self):
        self.items = []

    def add(self, level, row, col, item, msg, cur='', advice=''):
        self.items.append({'level': level, 'row': row, 'col': col, 'item': item,
                           'msg': msg, 'cur': '' if cur is None else str(cur),
                           'advice': advice})

    def count(self, level):
        return sum(1 for x in self.items if x['level'] == level)

    def sorted(self):
        """错误在前、警告其次、提示最后；同一级别按行号从小到大"""
        def key(x):
            try:
                rk = int(x['row'])
            except (TypeError, ValueError):
                rk = 999999
            return (LEVEL_ORDER.get(x['level'], 9), rk, str(x['col']))
        return sorted(self.items, key=key)


# ══════════════════════════════════════════════════════════════════════════
#  编码库
# ══════════════════════════════════════════════════════════════════════════
def find_header(ws, max_scan=15):
    for r in range(1, min(max_scan, ws.max_row) + 1):
        row = {}
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if v is not None and text(v):
                row[text(v)] = c
        if '商品编码' in row:
            return r, row
    return None, None


def spec_flavor(v):
    """规格串「2XL;1条装;170/74A」里那一段中文花色名"""
    if v is None:
        return ''
    for seg in str(v).split(';'):
        seg = seg.strip()
        if seg and re.fullmatch(r'[\u4e00-\u9fa5]{1,6}', seg):
            return seg
    return ''


class Library:
    """编码库：code -> 信息；key(货号,花色,尺码) -> code

    两种格式都认：
      · ①编码库 表格版（表头里有「商品编码」）
      · 旧的多列粘贴版（组标题 + 一串编码、列错位），生成器也认这种

    花色名支持带标记的写法：库里叫「粉色-EO」，表里写「粉色」「EO-粉色」
    「粉色-EO」都要能对上（EO = 抗菌真空包装，带不带是两个 SKU）。
    """

    def __init__(self, path):
        self.path = path
        self.codes = {}       # code -> dict(no, nm, sz, fam, row, cat)
        self.bykey = {}       # (no, nm, sz) -> code
        self.by_no = {}       # no -> set(nm)
        self.names_of = {}    # no -> [nm, ...]  按编码库出现顺序
        self.sizes_of = {}    # (no, nm) -> set(sz)
        self.issues = []
        self.ok = False
        self.fmt = ''
        self.dup_count = 0   # 库里重复的编码，静默跳过，不报
        self.load()

    # ── 载入 ──
    def load(self):
        wb = openpyxl.load_workbook(self.path, data_only=True)
        order = []
        for ws in wb.worksheets:
            name = ws.title
            if '编码库' in name:
                order.insert(0, ws)
            elif '编码' in name or '款式' in name:
                order.append(ws)
        for ws in wb.worksheets:
            if ws not in order:
                order.append(ws)

        for ws in order:
            hr, cols = find_header(ws)
            if hr:
                self._load_table(ws, hr, cols)
                if self.codes:
                    self.fmt = '①编码库（%s）' % ws.title
                    return
        for ws in order:
            if self._load_paste(ws):
                self.fmt = '旧版粘贴块（%s）' % ws.title
                return
        self.issues.append({'level': ERR, 'row': 0, 'col': '编码库', 'item': '编码库读不了',
                            'msg': '这个文件里没读出任何商品编码', 'cur': os.path.basename(self.path),
                            'advice': '确认选的是「款式编码文档」，里面要有编码'})

    def _register(self, code, no, nm, sz, fam, row, cat='', quiet=False):
        if code in self.codes:
            # 编码库里有重复行是常事（同一个编码抄了两遍），不影响建档，直接跳过不报
            self.dup_count += 1
            return False
        key = (no, nm, sz)
        if key in self.bykey and self.bykey[key] != code:
            if not quiet:
                self.issues.append({'level': WARN, 'row': row, 'col': '编码库', 'item': '库里冲突',
                                    'msg': '%s-%s-%s 已经有 %s，这行又写了 %s'
                                           % (no, nm or '空版', sz, self.bykey[key], code),
                                    'cur': code, 'advice': '同一个货号+花色+尺码挂了两条不同编码，核对哪条对'})
            return False
        self.bykey[key] = code
        self.codes[code] = {'no': no, 'nm': nm, 'sz': sz, 'fam': fam, 'row': row, 'cat': cat}
        lst = self.names_of.setdefault(no, [])
        if nm and nm not in lst:
            lst.append(nm)
        self.by_no.setdefault(no, set()).add(nm)
        self.sizes_of.setdefault((no, nm), set()).add(sz)
        return True

    def _load_table(self, ws, hr, cols):
        c_code = cols['商品编码']
        c_name, c_no = cols.get('花色'), cols.get('货号')
        c_sz, c_pat = cols.get('尺码'), cols.get('花色码')
        c_cat = cols.get('品类')
        for r in range(hr + 1, ws.max_row + 1):
            code = text(ws.cell(r, c_code).value)
            if not code:
                continue
            no, sz, fam = parse_code(code)
            if not fam:
                self.issues.append({'level': ERR, 'row': r, 'col': '编码库', 'item': '编码格式',
                                    'msg': '这条编码不是内裤型/身高型/背心型，建档表里没法用',
                                    'cur': code, 'advice': '核对是否漏字或多了符号'})
                continue
            if code != code.upper():
                self.issues.append({'level': WARN, 'row': r, 'col': '编码库', 'item': '编码大小写',
                                    'msg': '编码里有小写字母，平台会当成另一个商品',
                                    'cur': code, 'advice': '改成大写'})
            nm = text(ws.cell(r, c_name).value) if c_name else ''
            if c_no:
                v = text(ws.cell(r, c_no).value)
                if v and v != no:
                    self.issues.append({'level': WARN, 'row': r, 'col': '编码库', 'item': '货号不一致',
                                        'msg': '货号填的是 %s，编码里解析出来是 %s' % (v, no),
                                        'cur': code, 'advice': '以编码为准，把货号改成 %s' % no})
            if c_sz:
                v = text(ws.cell(r, c_sz).value)
                if v and norm_size(v) != sz:
                    self.issues.append({'level': WARN, 'row': r, 'col': '编码库', 'item': '尺码不一致',
                                        'msg': '尺码填的是 %s，编码里是 %s' % (v, sz),
                                        'cur': code, 'advice': '以编码为准，把尺码改成 %s' % sz})
            if c_pat and fam == 'inner':
                v = text(ws.cell(r, c_pat).value)
                m = PAT.match(code)
                real = m.group(3) if m else ''
                if v and v != real:
                    self.issues.append({'level': INFO, 'row': r, 'col': '编码库', 'item': '花色码不一致',
                                        'msg': '花色码填 %s，编码里是 %s' % (v, real or '（空）'),
                                        'cur': code, 'advice': '不影响生成，改不改都行'})
            self._register(code, no, nm, sz, fam, r,
                           text(ws.cell(r, c_cat).value) if c_cat else '')

    def _load_paste(self, ws):
        """旧格式：题目块 + 组标题 + 多列并排的编码（父子行错位），返回读到的编码数"""
        max_r, max_c = ws.max_row, min(ws.max_column, 30)
        if max_r < 2 or max_c < 1:
            return 0
        grid = [[text(ws.cell(r, c).value) for c in range(1, max_c + 1)]
                for r in range(1, max_r + 1)]

        info = {}
        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                s = grid[r - 1][c - 1]
                if not s:
                    continue
                no, sz, fam = parse_code(s)
                if fam:
                    info[(r, c)] = (no, sz, fam)
        if not info:
            return 0

        hdrs = {}
        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                s = grid[r - 1][c - 1]
                if not s or (r, c) in info:
                    continue
                tgt = None
                for cc in range(c + 1, max_c + 1):
                    if (r, cc) in info:
                        tgt = info[(r, cc)]
                        break
                if tgt is None:
                    for rr in range(r + 1, max_r + 1):
                        if (rr, c) in info:
                            tgt = info[(rr, c)]
                            break
                if tgt is None:
                    continue
                no = tgt[0]
                if '-' in s:
                    name = s.split('-', 1)[1].strip()
                else:
                    i = s.find(no)
                    if i < 0:
                        continue
                    name = s[i + len(no):].strip().lstrip('-').strip()
                if name:
                    hdrs[(r, c)] = name

        row_hdr, col_hdr = {}, {}
        for (r, c), nm in hdrs.items():
            row_hdr.setdefault(r, []).append((c, nm))
            col_hdr.setdefault(c, []).append((r, nm))

        def resolve_hdr(r, c):
            best = None
            for cc, nm in row_hdr.get(r, []):
                if cc < c and (best is None or cc > best[0]):
                    best = (cc, nm)
            if best is not None:
                return best[1]
            lst = [(rr, nm) for rr, nm in col_hdr.get(c, []) if rr <= r]
            if lst:
                return max(lst)[1]
            for cc in range(c - 1, 0, -1):
                lst = [(rr, nm) for rr, nm in col_hdr.get(cc, []) if rr <= r]
                if lst:
                    return max(lst)[1]
            return ''

        got = []
        for (r, c), (no, sz, fam) in sorted(info.items()):
            code = grid[r - 1][c - 1]
            if fam == 'inner':
                nm = resolve_hdr(r, c)
            else:
                nm = ''
                for cc in range(1, max_c + 1):
                    v = grid[r - 1][cc - 1]
                    if v and v != code and ';' in v:
                        f = spec_flavor(v)
                        if f:
                            nm = f
                            break
                if not nm:
                    nm = resolve_hdr(r, c)
            got.append((code, no, nm, sz, fam, r))

        # 同一款只有一种花色时，把空花色的行统一过去（避免同一款拆成两个键）
        names = {}
        for code, no, nm, sz, fam, r in got:
            if fam in ('height', 'vest'):
                names.setdefault(no, set()).add(nm)
        fix = {}
        for no, nms in names.items():
            real = [x for x in nms if x]
            if len(real) == 1:
                fix[no] = real[0]
        for code, no, nm, sz, fam, r in got:
            if not nm and no in fix:
                nm = fix[no]
            self._register(code, no, nm, sz, fam, r, quiet=True)
        return len(got)

    def get(self, code):
        return self.codes.get(code)

    def resolve(self, no, color, marker=None):
        """把 A 列里的花色写法对到编码库里的规范名字

        A 列里标记写在谁身上，谁就是带标记的那版：
          「39120-EO-粉色+黄色」→ 粉色-EO、黄色（原版）
        所以顺序是：显式标记优先 → 原样 → 换标记位置 → 库里唯一带标记的那版。
        都对不上就原样返回（后面会被判成「库里查不到」）。
        """
        names = self.names_of.get(no, [])
        color = text(color)
        if not color:
            return ''
        if marker:
            for c in ('%s-%s' % (color, marker), '%s-%s' % (marker, color)):
                if c in names:
                    return c
        if color in names:
            return color
        cands = []
        ps = [x for x in color.split('-') if x]
        if len(ps) >= 2 and _is_marker(ps[0]):
            cands.append('-'.join(ps[1:] + [ps[0]]))
        if len(ps) >= 2 and _is_marker(ps[-1]):
            cands.append('-'.join([ps[-1]] + ps[:-1]))
        for c in cands:
            if c in names:
                return c
        if len(ps) == 1:
            hits = [n for n in names if split_marker(n)[1] == color and split_marker(n)[0]]
            if len(hits) == 1:
                return hits[0]
        return color

    def judge_nm(self, no, nm, sz):
        """判断「货号+花色+尺码」在库里有没有；花色写错时给出提示"""
        if (no, nm, sz) in self.bykey:
            return None
        names = [x for x in self.names_of.get(no, []) if x]
        if no not in self.names_of and no not in self.by_no:
            return '%s 这个货号在编码库里没有' % no
        if nm and nm not in names:
            return '%s 在库里没有「%s」这个花色；现有：%s' % (no, nm, '、'.join(names) or '（只有空版）')
        if (no, nm) in self.sizes_of and sz not in self.sizes_of[(no, nm)]:
            have = '、'.join(sorted(self.sizes_of[(no, nm)], key=size_key))
            return '%s-%s 在库里没有 %s 码；现有：%s' % (no, nm or '空版', sz, have)
        return None

COMP_TOK = re.compile(r'^([A-Za-z]*\d{3,})(?:-(.+))?$')


def _is_marker(tok):
    """短英文标记（EO 之类），不是花色名的一部分"""
    return bool(re.fullmatch(r'[A-Za-z]{1,8}', tok or ''))


def split_marker(part):
    """'EO-粉色' / '粉色-EO' / '粉色' → (标记, 花色名)"""
    part = text(part)
    if not part:
        return None, ''
    ps = [x for x in part.split('-') if x]
    if len(ps) >= 2 and _is_marker(ps[0]):
        return ps[0], '-'.join(ps[1:])
    if len(ps) >= 2 and _is_marker(ps[-1]):
        return ps[-1], '-'.join(ps[:-1])
    return None, part


def parse_a(a):
    """A 列 → (尺码, [(货号, 花色原文)], 去掉尺码的组合部分)

    破折号后面那截是花色名的一部分（如「粉色-EO」）。
    只有第一个破折号前面像货号（字母 + 3 位以上数字）才算新货号，
    所以「39120-EO-粉色+黄色」里的「黄色」仍然是同一个货号的花色。
    """
    s = text(a)
    size, head = None, s
    m = re.search(r'-([A-Za-z0-9]+)$', s)
    if m and looks_like_size(m.group(1)):
        size = m.group(1)
        head = s[:m.start()]
    comps, cur = [], None
    for tok in head.split('+'):
        t = tok.strip()
        if not t:
            continue
        m = COMP_TOK.match(t)
        if m:
            cur = m.group(1)
            comps.append((cur, (m.group(2) or '').strip()))
        elif cur is not None:
            comps.append((cur, t))
        else:
            comps.append(('', t))
    return size, comps, head


def canonize(lib, comps, force=None):
    """把 A 列花色名对到编码库里的规范写法

    返回 ([(货号, 规范花色)], {(货号, 规范花色): {'raw': 原文, 'mk': 显式标记}})

    force 给的是 {(货号): 标记}：这个货号在 A 列里声明过 EO，
    整组的每个花色都必须是带 EO 的那一版（严格模式）。
    """
    out, note = [], {}
    force = force or {}
    for no, part in comps:
        mk, color = split_marker(part)
        if mk:
            nm = lib.resolve(no, color, mk) if lib else color
        elif no in force:
            want = '%s-%s' % (color, force[no])
            got = lib.resolve(no, color, force[no]) if lib else color
            # 库里没有这版就故意留个查不到的名字，后面会报「库里没有这个花色」
            nm = got if (not lib or got != color) else want
        else:
            nm = lib.resolve(no, color) if lib else color
        key = (no, nm)
        out.append(key)
        note[key] = {'raw': part, 'mk': mk}
    return out, note


def uniq(seq):
    out = []
    for x in seq:
        if x not in out:
            out.append(x)
    return out


# ══════════════════════════════════════════════════════════════════════════
#  读表
# ══════════════════════════════════════════════════════════════════════════
class Table:
    def __init__(self, path):
        self.path = path
        self.name = os.path.basename(path)
        self.wb = openpyxl.load_workbook(path)
        self.ws = None
        for ws in self.wb.worksheets:
            for r in range(1, 11):
                for c in range(1, 5):
                    if text(ws.cell(r, c).value) == '组合商品编码':
                        self.ws = ws
                        self.hr = r
                        break
                if self.ws:
                    break
            if self.ws:
                break
        self.rows = []       # [{'r':n, 'vals':[...17]}]
        self.labels = {}
        if self.ws is None:
            return
        ws = self.ws
        for c in range(1, 18):
            self.labels[c] = text(ws.cell(self.hr, c).value)
        last = self.hr
        for r in range(self.hr + 1, ws.max_row + 1):
            vals = [ws.cell(r, c).value for c in range(1, 18)]
            if any(v is not None and text(v) != '' for v in vals):
                last = r
                self.rows.append({'r': r, 'vals': vals})
        self.last_row = last

    def cell(self, r, c):
        return self.ws.cell(r, c)


# ══════════════════════════════════════════════════════════════════════════
#  检查主体
# ══════════════════════════════════════════════════════════════════════════
def check(table_path, lib_path=None, gap=GAP_DEFAULT, style=True, strict_eo=True):
    rep = Report()
    meta = {'file': os.path.basename(table_path), 'lib': '', 'rows': 0, 'groups': 0,
            'codes': 0, 'strict_eo': bool(strict_eo),
            'time': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

    lib = None
    if lib_path and os.path.exists(lib_path):
        lib = Library(lib_path)
        meta['lib'] = os.path.basename(lib_path)
        # 编码库自身的问题：同一类最多列 5 条，避免刷屏
        shown = {}
        for it in lib.issues:
            shown[it['item']] = shown.get(it['item'], 0) + 1
            if shown[it['item']] <= 5:
                rep.add(it['level'], ('库%d' % it['row']) if it['row'] else 0,
                        it['col'], it['item'], it['msg'], it['cur'], it['advice'])
        for item, n in shown.items():
            if n > 5:
                rep.add(INFO, 0, '编码库', item,
                        '「%s」这类问题编码库里有 %d 处，上面只列了前 5 处' % (item, n),
                        '', '打开编码库自己过一遍')

    t = Table(table_path)
    if t.ws is None:
        rep.add(ERR, 0, '-', '表格结构', '这张表里找不到「组合商品编码」表头，不是 17 列建档表',
                meta['file'], '确认选的是建档表，不是编码库')
        return {'report': rep, 'meta': meta, 'table': t}

    ws = t.ws
    # ── 一、表头 ────────────────────────────────────────────────────────
    hr = t.hr
    if hr > 1:
        rep.add(WARN, hr, '-', '表头位置', '表头在第 %d 行，不是第 1 行' % hr, '',
                '平台导入一般要求第 1 行是表头，建议整表上移')
    for i in range(1, 18):
        want = HEAD_NAME[i - 1]
        got = text(ws.cell(hr, i).value)
        if i == 17:
            if got:
                rep.add(WARN, hr, 'Q', '表头', 'Q 列（条数）不该有表头，现在写的是「%s」' % got, got,
                        '清空 Q 列的表头')
            continue
        if i not in (5, 6, 7, 8, 9, 13, 14, 15, 16):
            if got != want:
                rep.add(ERR, hr, get_column_letter(i), '表头',
                        '第 %d 列表头应为「%s」，现在是「%s」' % (i, want, got or '空'),
                        got, '照模板改回「%s」' % want)
    for c in range(18, ws.max_column + 1):
        v = text(ws.cell(hr, c).value)
        if v:
            rep.add(ERR, hr, get_column_letter(c), '多余列',
                    '第 %d 列多出一个表头「%s」，平台只要 17 列' % (c, v), v, '删掉这一列')

    if style:
        _check_head_style(t, rep, hr)

    # 数据行
    data = t.rows
    meta['rows'] = len(data)
    if not data:
        rep.add(ERR, 0, '-', '没有数据', '表头下面一行数据都没有', '', '')
        return {'report': rep, 'meta': meta, 'table': t}

    # ── 二、分组 ────────────────────────────────────────────────────────
    groups, gaps = [], []
    prev_r = None
    for d in data:
        a = text(d['vals'][0])
        if not a:
            rep.add(ERR, d['r'], 'A', '空行', 'A 列是空的，但这一行有内容', '', '清空整行或补上组合编码')
            continue
        if not groups or (prev_r is not None and d['r'] != prev_r + 1):
            groups.append([])
        groups[-1].append(d)
        prev_r = d['r']
    # 组间空行
    for i in range(1, len(groups)):
        blank = groups[i][0]['r'] - groups[i - 1][-1]['r'] - 1
        gaps.append(blank)
        if blank != gap:
            rep.add(INFO, groups[i - 1][-1]['r'], '-', '组间空行',
                    '第 %d 组和第 %d 组之间空了 %d 行，设定是 %d 行'
                    % (i, i + 1, blank, gap), blank, '不影响公式，只是版式')
    meta['groups'] = len(groups)

    tail = []
    for r in range(t.last_row + 1, ws.max_row + 1):
        if any(ws.cell(r, c).value is not None for c in range(1, 18)):
            tail.append(r)
    if tail:
        rep.add(WARN, tail[0], '-', '多余内容', '最后一组下面还有内容（第 %d 行起）' % tail[0], '',
                '平台导入会把它们当成数据')

    used_codes = {}
    a_prefix = {}
    b_tags = {}

    for gi, g in enumerate(groups, 1):
        _check_group(t, g, gi, rep, lib, gap, style, used_codes, a_prefix, b_tags, strict_eo)

    # ── 三、全表 ────────────────────────────────────────────────────────
    if len(b_tags) > 1:
        rep.add(WARN, 0, 'B', '标签不统一',
                '整表出现了 %d 种标签开头：%s'
                % (len(b_tags), '、'.join(sorted(b_tags))), '',
                '一张建档表通常只用一种标签（如绿色标）')
    if lib:
        unused = [c for c in lib.codes if c not in used_codes]
        if unused and len(unused) < 200:
            rep.add(INFO, 0, '-', '库里未用',
                    '编码库里有 %d 个编码这次没用到（不影响上传）' % len(unused),
                    '、'.join(sorted(unused)[:6]) + ('…' if len(unused) > 6 else ''), '')

    meta['codes'] = len(used_codes)

    if rep.count(ERR) == 0:
        rep.add(OK, 0, '-', '总体', '没有发现错误，可以上传', '', '')
    return {'report': rep, 'meta': meta, 'table': t}


def _check_head_style(t, rep, hr):
    ws = t.ws
    h = ws.row_dimensions[hr].height
    if h and abs(h - 31) > 0.6:
        rep.add(INFO, hr, '-', '表头行高', '表头行高是 %.1f，模板是 31' % h, '%.1f' % h, '')
    for i in range(1, 18):
        if i == 17:
            continue  # Q 列表头本来就是空的、宋体 11，不参与表头样式比对
        c = ws.cell(hr, i)
        col = get_column_letter(i)
        fn = c.font.name
        fs = c.font.size
        if fn and fn != '微软雅黑':
            rep.add(INFO, hr, col, '表头字体', '表头字体是 %s，模板是微软雅黑' % fn, fn, '')
        if fs and abs(float(fs) - 9) > 0.01:
            rep.add(INFO, hr, col, '表头字号', '表头字号是 %s，模板是 9' % fs, fs, '')
        color = c.font.color.rgb if (c.font.color and c.font.color.rgb) else None
        is_red = bool(color) and str(color).upper().endswith('FF0000')
        if i in RED_HEAD and not is_red:
            rep.add(INFO, hr, col, '表头颜色', '这列表头应该是红字', color or '黑', '')
        if i not in RED_HEAD and i != 17 and is_red:
            rep.add(INFO, hr, col, '表头颜色', '这列表头多涂了红色，模板是黑字', color, '')
    for col, w in WIDTH.items():
        if col in ('A', 'C'):
            continue
        cw = ws.column_dimensions[col].width
        if cw and abs(cw - w) > 0.6:
            rep.add(INFO, hr, col, '列宽', '%s 列宽是 %.1f，模板是 %s' % (col, cw, w), '%.1f' % cw, '')


def _comp_label(no, nm):
    return '%s-%s' % (no, nm) if nm else '%s（空版）' % no


def _check_group(t, g, gi, rep, lib, gap, style, used_codes, a_prefix, b_tags, strict_eo=True):
    r0 = g[0]['r']
    # 尺码 → 行
    by_size = {}
    for d in g:
        sz = parse_a(text(d['vals'][0]))[0] or ''
        by_size.setdefault(norm_size(sz), []).append(d)

    # ── A 列：同一尺码内必须完全一样 ────────────────────────────────
    for szk, ds in by_size.items():
        first = text(ds[0]['vals'][0])
        for d in ds[1:]:
            v = text(d['vals'][0])
            if v != first:
                rep.add(ERR, d['r'], 'A', 'A 列不一致',
                        '同一个尺码（%s）里 A 列必须一字不差：第一行是「%s」，这行是「%s」'
                        % (szk or '?', first, v), v,
                        '同一个尺码的所有行填同一个组合编码')

    # ── C 列必须和 A 列一字不差 ────────────────────────────────────
    for d in g:
        a_v, c_v = text(d['vals'][0]), text(d['vals'][2])
        if c_v != a_v:
            rep.add(ERR, d['r'], 'C', 'A≠C',
                    '组合款式编码必须和组合商品编码一字不差', 'A=%s ｜ C=%s' % (a_v, c_v),
                    '把 C 列直接复制成 A 列')

    # ── D 列：组合商品名称不能空，同组只能一个 ──────────────────────
    d_vals = set(text(d['vals'][3]) for d in g)
    for d in g:
        if not text(d['vals'][3]):
            rep.add(ERR, d['r'], 'D', '名称空', '组合商品名称（D 列）是空的', '',
                    '填品类名，如 儿童内裤 / 少女背心 / 家居服套装')
    if len(d_vals) > 1:
        rep.add(ERR, r0, 'D', '组内名称不一致',
                '同一个组合里 D 列出现了 %d 个不同的名称：%s'
                % (len(d_vals), '、'.join(sorted(x or '（空）' for x in d_vals))), '',
                '同组应该只有一个品类名')

    # ── 组合前缀（去掉尺码）全组一致 ────────────────────────────────
    a0 = text(g[0]['vals'][0])
    size0, comps_raw, head = parse_a(a0)
    # 严格模式：A 列里声明了 EO 的货号，整组每个花色都要带 EO
    force = {}
    if strict_eo:
        for no, part in comps_raw:
            mk, _color = split_marker(part)
            if mk:
                force.setdefault(no, mk)
    comps, comp_note = canonize(lib, comps_raw, force)
    if not size0:
        rep.add(ERR, r0, 'A', 'A 列格式',
                'A 列末尾没有尺码（应以「-尺码」结尾）', a0, '示例：95001-天青+95129-藤紫-M')
    if not comps:
        rep.add(ERR, r0, 'A', 'A 列格式', 'A 列里读不出款号', a0, '示例：95001-天青+95129-藤紫-M')
        return
    for d in g:
        sz, cs, hd = parse_a(text(d['vals'][0]))
        if hd and hd != head:
            rep.add(ERR, d['r'], 'A', '组合不一致',
                    '同一组里组合部分不一样：本组是「%s」，这行是「%s」' % (head, hd),
                    text(d['vals'][0]), '同一组合的所有尺码，组合部分必须相同')

    dup_comps = [c for c in uniq(comps) if comps.count(c) > 1]
    if dup_comps:
        rep.add(INFO, r0, 'A', '同款重复',
                '组合里同一款号写了两次：%s（2 件同款时正常）'
                % '、'.join(_comp_label(n, s) for n, s in dup_comps), '', '')

    if head in a_prefix:
        rep.add(WARN, r0, 'A', '组合重复',
                '这个组合（%s）在前面第 %d 行也出现过，被拆成了两段' % (head, a_prefix[head]),
                head, '同一组合的尺码最好连着排')
    else:
        a_prefix[head] = r0

    order = [norm_size(parse_a(text(d['vals'][0]))[0] or '') for d in g]
    uniq_order = uniq(order)
    if uniq_order != sorted(uniq_order, key=size_key):
        rep.add(INFO, r0, 'A', '尺码顺序',
                '尺码不是从小到大：%s' % '、'.join(uniq_order), '',
                '平台不校验顺序，但不好核对')

    # ── B 列 ────────────────────────────────────────────────────────
    _check_labels(t, g, rep, b_tags)

    # ── 逐行 ────────────────────────────────────────────────────────
    fixed_q = []          # Q 写成固定数字但数值正确的行，组内汇总成一条警告
    for d in g:
        r = d['r']
        vals = d['vals']
        sz_a = parse_a(text(vals[0]))[0]
        code = text(vals[9])
        k, l, q = vals[10], vals[11], vals[16]

        # K 列
        if k is None or text(k) == '':
            rep.add(ERR, r, 'K', '数量空', '数量（K 列）是空的', '', '填 1')
        else:
            try:
                if float(k) != 1:
                    rep.add(ERR, r, 'K', '数量不是1', '数量应该是 1，现在是 %s' % k, k, '改成 1')
            except (TypeError, ValueError):
                rep.add(ERR, r, 'K', '数量不是数字', '数量应该是 1，现在是「%s」' % k, k, '改成 1')

        # L 列公式
        want_l = '=ROUND(Q%d/SUMIFS(Q:Q,A:A,A%d),4)' % (r, r)
        if norm_formula(l) != norm_formula(want_l):
            rep.add(ERR, r, 'L', 'L 公式',
                    '应占售价公式不对，应该是 =ROUND(Q%d/SUMIFS(Q:Q,A:A,A%d),4)' % (r, r),
                    l, '整列套同一个公式')

        # J 列
        if not code:
            rep.add(ERR, r, 'J', '缺编码',
                    '商品编码（J 列）是空的，这行传上去会报错', '',
                    '编码库里没有这个花色/尺码，补库后重新生成')
            if q is not None and text(q) != '':
                rep.add(WARN, r, 'Q', 'Q 无编码', 'J 列空着，Q 列却填了值', q, '')
            continue

        used_codes[code] = used_codes.get(code, 0) + 1
        if code != code.upper():
            rep.add(WARN, r, 'J', '编码大小写', '编码不是全大写，平台会当成另一个商品', code, '改成大写')

        no, sz, fam = parse_code(code)
        if not fam:
            rep.add(ERR, r, 'J', '编码格式',
                    '这个编码认不出来，不属于内裤型 / 身高型 / 背心型', code,
                    '核对是不是漏字、多了符号，或编码库里没有这条')
            continue

        if sz_a and norm_size(sz_a) != norm_size(sz):
            rep.add(ERR, r, 'J', '尺码不符',
                    'A 列这行是 %s 码，J 列编码却是 %s 码' % (sz_a, sz), code,
                    '这行编码填错了，会发错货')
        comp_nos = [n for n, _ in comps]
        if no not in comp_nos:
            rep.add(ERR, r, 'J', '货号不符',
                    '编码属于 %s，但 A 列这个组合里没有 %s' % (no, no), code,
                    'A 列组合和 J 列编码对不上')

        if lib:
            if lib.get(code) is None:
                rep.add(ERR, r, 'J', '库里查不到',
                        '编码库里没有这个编码', code,
                        '可能是手改过、拼错了，或编码库还没更新')

        _check_q(t, r, code, fam, q, k, rep, fixed_q)
        _check_cols_empty(t, r, rep)   # E–I、M–P 留空是数据问题，关掉样式检查也要查

        if style:
            _check_row_style(t, r, rep)

    # ── Q 列写成固定数字（数值没写错）→ 组内汇总一条警告 ──────────────
    if fixed_q:
        r0q = fixed_q[0]
        rep.add(WARN, r0q, 'Q', 'Q 写死数字',
                '本组有 %d 行的 Q 列写成了固定数字，数字是对的；按规范内裤型这里应该是公式，'
                '以后条数一改它不会跟着变' % len(fixed_q), '',
                '想稳一点就整列套公式 =LEFT(_xlfn.TEXTAFTER(J%d,"T"),1)*K%d' % (r0q, r0q))

    # ── 按尺码核对编码是否齐全 ──────────────────────────────────────
    if lib:
        _check_expected_codes(t, by_size, comps, rep, lib, comp_note, force)
    else:
        _check_sizes_by_no(t, by_size, comps, rep)

    # ── 应占售价合计 / 条数 ─────────────────────────────────────────
    _check_l_price(t, g, rep, comps, lib)
    _check_category(t, g, rep, lib)


def _check_labels(t, g, rep, b_tags):
    """B 列：标签,尺码<分隔符>N条装"""
    seps = set()
    b_by_size = {}
    for d in g:
        b = text(d['vals'][1])
        r = d['r']
        sz_a = parse_a(text(d['vals'][0]))[0]
        if not b:
            rep.add(ERR, r, 'B', '标签空', '组合装商品标签（B 列）是空的', '',
                    '格式：绿色标,M，2条装')
            continue
        m = re.match(r'^([^,，]+)[,，]([^,，]+?)([,，])(.+)$', b)
        if not m:
            rep.add(ERR, r, 'B', '标签格式',
                    '标签格式不对，应该是「标签,尺码，N条装」', b, '例：绿色标,M，2条装')
            continue
        tag, bsz, sep, mid = m.group(1), m.group(2), m.group(3), m.group(4)
        seps.add(sep)
        b_tags[tag] = b_tags.get(tag, 0) + 1
        if sz_a and norm_size(bsz) != norm_size(sz_a):
            rep.add(ERR, r, 'B', '标签尺码不符',
                    'B 列尺码写的是 %s，A 列是 %s' % (bsz, sz_a), b,
                    'B 列的尺码要和 A 列末尾一致')
        if not looks_like_size(bsz):
            rep.add(ERR, r, 'B', '标签尺码', 'B 列里这个尺码「%s」认不出来' % bsz, b, '')
        if not mid.strip():
            rep.add(ERR, r, 'B', '标签条数', 'B 列后半段（N条装）是空的', b, '')
        b_by_size.setdefault(norm_size(bsz), set()).add(b)
    for sz, vals in b_by_size.items():
        if len(vals) > 1:
            rep.add(ERR, g[0]['r'], 'B', '标签不统一',
                    '%s 码这一组的标签有 %d 种写法：%s'
                    % (sz, len(vals), ' ｜ '.join(sorted(vals))), '',
                    '同一尺码所有行的标签应完全一样')
    if len(seps) > 1:
        rep.add(WARN, g[0]['r'], 'B', '分隔符不统一',
                '同一个组合里中文逗号「，」和英文逗号「,」混着用', '',
                '少女文胸/背心用英文逗号，其余用中文逗号，整表统一')


def _check_q(t, r, code, fam, q, k, rep, fixed_q):
    want_inner = '=LEFT(_xlfn.TEXTAFTER(J%d,"T"),1)*K%d' % (r, r)
    if fam in ('height', 'vest'):
        if text(q) == '':
            rep.add(ERR, r, 'Q', 'Q 空', '身高型/背心型编码没有 T 段，Q 列应直接写 1', '', Q_TIP)
        else:
            try:
                if float(q) != 1:
                    rep.add(ERR, r, 'Q', 'Q 应为1',
                            '身高型/背心型 Q 列应该是 1，现在是 %s' % q, q, Q_TIP)
            except (TypeError, ValueError):
                rep.add(ERR, r, 'Q', 'Q 应为1',
                        '身高型/背心型 Q 列应该是 1，现在是「%s」' % q, q,
                        '这种编码没有 T 段，套公式会变成 #VALUE!')
        return
    if norm_formula(q) != norm_formula(want_inner):
        s = text(q)
        if re.fullmatch(r'\d+(\.\d+)?', s):
            # 写死数字：等于正确条数就当警告（组内汇总一条），不等于才是错
            if abs(float(s) - t_count(code)) < 1e-9:
                fixed_q.append(r)
            else:
                rep.add(ERR, r, 'Q', 'Q 值不对',
                        '内裤型这一行的条数应该是 %d，Q 列现在写的是 %s'
                        % (t_count(code), s), q,
                        'Q = 编码里 T 后面那个数字；按规范这里应该用公式 ' + want_inner)
        else:
            rep.add(ERR, r, 'Q', 'Q 公式',
                    '内裤型 Q 列公式不对，应该是 ' + want_inner, q, Q_TIP)
    risk = q_risk(code)
    if risk:
        rep.add(WARN, r, 'Q', 'Q 需手改',
                '%s —— 这一格的公式会算出 #VALUE!，整组应占售价都会跟着错' % risk,
                code, '把这一格的 Q 手工改成正确条数 %d' % t_count(code))


def _check_row_style(t, r, rep):
    ws = t.ws
    for i in (1, 2, 3, 4, 11, 12, 17):
        c = ws.cell(r, i)
        col = get_column_letter(i)
        if c.font.name and c.font.name != '微软雅黑':
            rep.add(INFO, r, col, '字体', '%s 列字体是 %s，模板是微软雅黑' % (col, c.font.name),
                    c.font.name, '')
            break
        if c.font.size and abs(float(c.font.size) - 9) > 0.01:
            rep.add(INFO, r, col, '字号', '%s 列字号是 %s，模板是 9' % (col, c.font.size),
                    c.font.size, '')
            break
    jc = ws.cell(r, 10)
    if jc.font.name and jc.font.name != '宋体':
        rep.add(INFO, r, 'J', '字体', 'J 列字体是 %s，模板是宋体 11' % jc.font.name,
                jc.font.name, '')
    if jc.font.size and abs(float(jc.font.size) - 11) > 0.01:
        rep.add(INFO, r, 'J', '字号', 'J 列字号是 %s，模板是 11' % jc.font.size,
                jc.font.size, '')


def _check_cols_empty(t, r, rep):
    ws = t.ws
    for i in list(range(5, 10)) + list(range(13, 17)):
        v = ws.cell(r, i).value
        if v is not None and text(v) != '':
            rep.add(WARN, r, get_column_letter(i), '该留空',
                    '%s 列（%s）本该留空，填了「%s」'
                    % (get_column_letter(i), HEAD_NAME[i - 1], v), v, '清空这一列')


def _variant_codes(lib, no, nm, szk):
    """同一个花色的「带标记版 / 不带标记版」两个编码都可能对

    A 列只把 EO 这种标记写一次，后面的花色看不出来带不带，
    所以这类行不能硬判成错。
    """
    mk, color = split_marker(nm)
    names = lib.names_of.get(no, [])
    out = set()
    for cand in names:
        mk2, color2 = split_marker(cand)
        if color2 == color:
            code = lib.bykey.get((no, cand, szk))
            if code:
                out.add(code)
    return out


def _check_expected_codes(t, by_size, comps, rep, lib, meta, force=None):
    """有编码库时：每个尺码应该正好是「组合里每个款号 × 库里的编码」这一套，且顺序要对上"""
    for szk, ds in by_size.items():
        rows = list(ds)
        if not szk:
            continue
        exp = [(no, nm, lib.bykey.get((no, nm, szk))) for no, nm in comps]
        if not any(c for _n, _m, c in exp):
            continue
        if len(rows) != len(exp):
            rep.add(ERR, rows[0]['r'], 'J', '行数不对',
                    '%s 码应该是 %d 行（组合里 %d 个款号/花色），实际 %d 行'
                    % (szk, len(exp), len(exp), len(rows)), '',
                    '少的补行、多的删行；一行一个款号')
        for i, (no, nm, code) in enumerate(exp):
            if code is None:
                if (no, nm) in lib.sizes_of:
                    have = '、'.join(sorted(lib.sizes_of[(no, nm)], key=size_key))
                    rep.add(ERR, rows[0]['r'], 'J', '库里没有这个码',
                            '%s 码：库里 %s 没有这个尺码（现有 %s）'
                            % (szk, _comp_label(no, nm), have), '',
                            '补编码库，或把这个尺码从这个组合里去掉')
                elif nm and no in lib.names_of and nm not in lib.names_of[no]:
                    rep.add(ERR, rows[0]['r'], 'J', '库里没有这个花色',
                            '%s 码：库里 %s 没有「%s」这个花色（现有：%s）'
                            % (szk, no, nm, '、'.join(
                                [x for x in lib.names_of.get(no, []) if x][:8])), '',
                            'A 列的花色名要照编码库写（带 EO 的要写全）')
                continue
            if i >= len(rows):
                continue
            got = text(rows[i]['vals'][9])
            if not got or got == code:
                continue
            info = lib.get(got)
            if info is None:
                continue          # 逐行已经报过「库里查不到」
            accept = {code}
            if not (force and no in force):
                # 没声明标记的组，同一个花色的两种版本都算对；
                # 声明了 EO 的组只认带 EO 的那一版
                accept |= _variant_codes(lib, no, nm, szk)
            if got in accept:
                continue
            got_label = _comp_label(info['no'], info['nm'])
            same_base = (info['no'] == no
                         and split_marker(info['nm'])[1] == split_marker(nm)[1])
            if same_base:
                if force and no in force:
                    why = ('A 列这个组合写了「%s」标记，整组都必须用带 %s 的编码'
                           % (force[no], force[no]))
                else:
                    why = '带不带 EO 是两个不同 SKU，按 A 列的写法选对应的编码'
                rep.add(ERR, rows[i]['r'], 'J', '版本不对',
                        '%s 码这一行填的是 %s，应该是 %s（%s）'
                        % (szk, got, code, why), got, why)
                continue
            other = [j for j, (n2, m2, c2) in enumerate(exp)
                     if info['no'] == n2 and split_marker(info['nm'])[1] == split_marker(m2)[1]]
            if other and other[0] != i:
                rep.add(ERR, rows[i]['r'], 'J', '花色顺序错',
                        '%s 码第 %d 行应该是 %s（%s），现在填的是 %s 的编码 %s'
                        % (szk, i + 1, _comp_label(no, nm), code, got_label, got), got,
                        'A 列花色顺序要和行顺序一一对应，错排会导致库存分裂')
            else:
                rep.add(ERR, rows[i]['r'], 'J', '花色不符',
                        '%s 码这一行应该是 %s（%s），填的却是 %s 的编码 %s'
                        % (szk, _comp_label(no, nm), code, got_label, got), got,
                        '按 A 列的花色重新查编码')


def _check_category(t, g, rep, lib):
    """D 列品类名要和编码库里的品类一致"""
    if not lib:
        return
    bad = {}
    for d in g:
        code = text(d['vals'][9])
        cat = text(d['vals'][3])
        if not code or not cat:
            continue
        info = lib.get(code)
        if info and info.get('cat') and info['cat'] != cat:
            bad.setdefault((cat, info['cat']), []).append(d['r'])
    for (cat, libcat), rows in bad.items():
        rep.add(WARN, rows[0], 'D', '品类不一致',
                'D 列写的是「%s」，编码库里这些编码的品类是「%s」（%d 行对不上）'
                % (cat, libcat, len(rows)), cat,
                '按编码库的品类名改，别自己另起名字')


def _check_sizes_by_no(t, by_size, comps, rep):
    """没有编码库时：只能按货号数行，够不够"""
    comp_nos = {}
    for no, nm in comps:
        comp_nos[no] = comp_nos.get(no, 0) + 1
    for szk, ds in by_size.items():
        counts, empties = {}, 0
        for d in ds:
            code = text(d['vals'][9])
            if not code:
                empties += 1
                continue
            no = parse_code(code)[0]
            counts[no] = counts.get(no, 0) + 1
            if no not in comp_nos:
                rep.add(ERR, d['r'], 'J', '多余行',
                        '%s 码里出现 %s，但 A 列组合里没有这个款号' % (szk or '?', no), '',
                        '')
        if empties:
            continue  # 缺的已经在「缺编码」里说过
        for no, want in comp_nos.items():
            got = counts.get(no, 0)
            if got < want:
                rep.add(ERR, ds[0]['r'], 'J', '行数不足',
                        '%s 码里 %s 应该有 %d 行，实际只有 %d 行' % (szk or '?', no, want, got),
                        '', '这个尺码少建了编码')
        codes = [text(d['vals'][9]) for d in ds if text(d['vals'][9])]
        for code in uniq(codes):
            n = codes.count(code)
            if n <= 1:
                continue
            no = parse_code(code)[0]
            cap = comp_nos.get(no, 1)
            rep.add(INFO if n <= cap else ERR, ds[0]['r'], 'J', '编码重复',
                    '%s 码里 %s 出现了 %d 次' % (szk or '?', code, n), code,
                    '同款多件时正常；否则是重复建行')


def _check_l_price(t, g, rep, comps, lib):
    """按 Q 重算应占售价，看整组合计是不是 1；顺带核 N条装"""
    qs, risky = [], False
    for d in g:
        code = text(d['vals'][9])
        if not code:
            qs.append(None)
            continue
        no, sz, fam = parse_code(code)
        if fam != 'inner':
            qs.append(1)
            continue
        if q_risk(code):
            risky = True
        qs.append(t_count(code))
    if any(q is None for q in qs):
        rep.add(WARN, g[0]['r'], 'L', '无法核对',
                '这一组有行缺编码，应占售价算不出去', '', '补编码后再核')
        return
    total = sum(qs)
    if total <= 0:
        return
    if risky:
        rep.add(WARN, g[0]['r'], 'L', '合计会失真',
                '这一组里有 Q 会算成 #VALUE!，整组应占售价合计不是 1', '',
                '把标黄的 Q 改成正确条数')
        return
    if abs(sum(q / total for q in qs) - 1) > 0.0005:
        rep.add(ERR, g[0]['r'], 'L', '合计不为1',
                '按 Q 重算，这一组应占售价合计不是 1', '', '')

    codes = [text(d['vals'][9]) for d in g if text(d['vals'][9])]
    if not codes or not all(parse_code(c)[2] == 'inner' for c in codes):
        return
    b = text(g[0]['vals'][1])
    m = re.match(r'^([^,，]+)[,，]([^,，]+?)[,，](.+)$', b)
    if not m:
        return
    mid = m.group(3)
    mm = re.search(r'(\d+)\s*(条|件|套)', mid)
    if not mm:
        return
    # N 条装 = 每个款号编码里 T 的数字之和（同款只算一次，与尺码多少无关）
    t_by_no = {}
    for d in g:
        code = text(d['vals'][9])
        if not code:
            continue
        cno, _csz, cfam = parse_code(code)
        if cfam == 'inner':
            t_by_no.setdefault(cno, set()).add(t_count(code))
    t_of_comp, ambiguous = [], False
    for no, nm in comps:
        ts = set()
        for d in g:
            code = text(d['vals'][9])
            if not code:
                continue
            cno, csz, cfam = parse_code(code)
            if cfam != 'inner' or cno != no:
                continue
            if lib and nm:
                info = lib.get(code)
                if info and info['nm'] and info['nm'] != nm:
                    continue
            ts.add(t_count(code))
        if len(ts) == 1:
            t_of_comp.append(ts.pop())
        elif len(ts) > 1:
            ambiguous = True
        else:
            fallback = t_by_no.get(no, set())
            if len(fallback) == 1:
                t_of_comp.append(next(iter(fallback)))
            elif len(fallback) > 1:
                ambiguous = True
    n = int(mm.group(1))
    want = sum(t_of_comp)
    if ambiguous or not t_of_comp:
        rep.add(INFO, g[0]['r'], 'B', '条数需人工确认',
                '同一个货号不同花色的 T 值不一样，工具算不准条数，请自己核对一遍', '',
                '')
        return
    if n != want:
        rep.add(ERR, g[0]['r'], 'B', '条数不符',
                '标签写 %d%s装，但这一组各款号 T 合计是 %d 条' % (n, mm.group(2), want),
                b, '条数 = 各款号编码 T 后面数字之和，不是行数也不是款数')
    elif mid.strip() != '%d%s装' % (n, mm.group(2)):
        rep.add(INFO, g[0]['r'], 'B', '条数写法',
                '标签后半段是「%s」，习惯写法是「%d%s装」' % (mid.strip(), n, mm.group(2)),
                b, '')


# ══════════════════════════════════════════════════════════════════════════
#  出报告
# ══════════════════════════════════════════════════════════════════════════
thin = Side(style='thin', color='BFBFBF')
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)


def write_report(out_path, res):
    rep, meta, t = res['report'], res['meta'], res['table']
    wb = Workbook()
    ws = wb.active
    ws.title = '①检查结果'
    ws.column_dimensions['A'].width = 70
    ws.column_dimensions['B'].width = 34

    title_font = Font(name='微软雅黑', size=15, bold=True, color='FF1F3864')
    lab = Font(name='微软雅黑', size=10, bold=True)
    val = Font(name='微软雅黑', size=10)
    ws['A1'] = '建档表检查报告'
    ws['A1'].font = title_font
    ws.merge_cells('A1:B1')

    rows = [('被检查文件', meta['file']), ('编码库', meta['lib'] or '（未提供，只查结构/公式）'),
            ('检查时间', meta['time']),
            ('数据行 / 组合数', '%d 行 / %d 组' % (meta['rows'], meta['groups'])),
            ('商品编码数', '%d 个' % meta['codes']),
            ('错误', '%d 处' % rep.count(ERR)), ('警告', '%d 处' % rep.count(WARN)),
            ('提示', '%d 处' % rep.count(INFO)),
            ('结论', '可以上传' if rep.count(ERR) == 0 else '先改完 %d 处错误再上传' % rep.count(ERR))]
    r = 3
    for k, v in rows:
        ws.cell(r, 1, k).font = lab
        c = ws.cell(r, 2, v)
        c.font = val
        if k == '错误' and rep.count(ERR):
            c.font = Font(name='微软雅黑', size=10, bold=True, color='FF9C0006')
        if k == '结论':
            c.font = Font(name='微软雅黑', size=10, bold=True,
                          color='FF9C0006' if rep.count(ERR) else 'FF006100')
        r += 1

    r += 1
    heads = ['级别', '行号', '列', '检查项', '说明', '单元格现值', '怎么改']
    widths = [7, 8, 6, 14, 58, 30, 34]
    for i, (h, w) in enumerate(zip(heads, widths), 1):
        c = ws.cell(r, i, h)
        c.font = Font(name='微软雅黑', size=9, bold=True, color='FFFFFFFF')
        c.fill = PatternFill('solid', fgColor='FF1F3864')
        c.alignment = Alignment(horizontal='left', vertical='center')
        c.border = BORDER
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.column_dimensions['A'].width = 7
    r += 1
    head_row = r - 1
    for it in rep.sorted():
        if it['level'] == OK:
            continue
        vals = [it['level'], it['row'] or '', it['col'], it['item'],
                it['msg'], it['cur'], it['advice']]
        for i, v in enumerate(vals, 1):
            c = ws.cell(r, i, v)
            c.font = Font(name='微软雅黑', size=9,
                          color='FF' + FONT_OF.get(it['level'], '000000'))
            c.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
            c.border = BORDER
        ws.cell(r, 1).fill = PatternFill('solid', fgColor='FF' + FILL_OF.get(it['level'], 'FFFFFF'))
        r += 1
    ws.freeze_panes = 'A%d' % (head_row + 1)

    # ② 标注版
    if t.ws is not None:
        src = t.ws
        dst = wb.create_sheet('②问题标注')
        for col, dim in src.column_dimensions.items():
            if dim.width:
                dst.column_dimensions[col].width = dim.width
        for rr, dim in src.row_dimensions.items():
            if dim.height:
                dst.row_dimensions[rr].height = dim.height
        for row in src.iter_rows(min_row=1, max_row=min(src.max_row, 2000), max_col=17):
            for c in row:
                d = dst.cell(c.row, c.column, c.value)
                d.font = _copy.copy(c.font)
                d.fill = _copy.copy(c.fill)
                d.border = _copy.copy(c.border)
                d.alignment = _copy.copy(c.alignment)
                d.number_format = c.number_format
        for it in rep.sorted():
            if it['level'] == OK or not it['row'] or it['col'] in ('-', '编码库'):
                continue
            try:
                col = openpyxl.utils.column_index_from_string(it['col'])
            except ValueError:
                continue
            cell = dst.cell(it['row'], col)
            cell.fill = PatternFill('solid', fgColor='FF' + FILL_OF.get(it['level'], 'FFFFFF'))
            old = cell.comment.text if cell.comment else ''
            cell.comment = Comment(('%s【%s】%s' % (old + '\n' if old else '', it['item'], it['msg']))
                                   [:800], '建档表检查器', width=300, height=120)
        dst.freeze_panes = 'A%d' % (t.hr + 1)

    wb.save(out_path)
    return out_path


# ══════════════════════════════════════════════════════════════════════════
#  命令行
# ══════════════════════════════════════════════════════════════════════════
def cli(argv):
    import argparse
    ap = argparse.ArgumentParser(description='检查拼多多 17 列建档表')
    ap.add_argument('--check', required=True, help='要检查的建档表 xlsx')
    ap.add_argument('--lib', default='', help='款式编码文档（①编码库）xlsx，可选')
    ap.add_argument('--out', default='', help='报告输出路径，可选')
    ap.add_argument('--gap', type=int, default=GAP_DEFAULT, help='组间空行，默认 10')
    ap.add_argument('--no-style', action='store_true', help='不检查字体、字号、列宽')
    ap.add_argument('--loose-eo', action='store_true',
                    help='不强制整组都是 EO（默认强制）')
    ap.add_argument('--json', action='store_true', help='以 JSON 输出结果')
    a = ap.parse_args(argv)
    res = check(a.check, a.lib or None, a.gap, not a.no_style, not a.loose_eo)
    rep = res['report']
    lines = []
    if a.json:
        lines.append(json.dumps({'meta': res['meta'], 'items': rep.sorted()},
                                ensure_ascii=False, indent=1))
    else:
        for it in rep.sorted():
            lines.append('[%s] %s%s %s：%s'
                         % (it['level'], row_label(it['row']),
                            it['col'], it['item'], it['msg']))
        lines.append('—— 错误 %d / 警告 %d / 提示 %d'
                     % (rep.count(ERR), rep.count(WARN), rep.count(INFO)))
    if a.out:
        write_report(a.out, res)
        lines.append('报告：%s' % a.out)
    _emit(lines, a.check)
    return 0 if rep.count(ERR) == 0 else 1


def row_label(r):
    """行号显示：正常是「第12行」，编码库里的问题显示「库757」，整表级的不显示"""
    if r is None or r == '' or str(r) in ('0', ''):
        return ''
    s = str(r)
    return ('第%s行 ' % s) if s.isdigit() else (s + ' ')


def _emit(lines, anchor):
    """命令行输出：打包成窗口程序后没有真的控制台，所以额外写一份日志到表格旁边"""
    text_out = '\n'.join(lines) + '\n'
    wrote = False
    if sys.stdout is not None:
        try:
            sys.stdout.write(text_out)
            sys.stdout.flush()
            wrote = True
        except Exception:
            pass
    if wrote and not getattr(sys, 'frozen', False):
        return
    try:
        stem = os.path.splitext(os.path.abspath(anchor))[0]
        with open(stem + '.检查日志.txt', 'w', encoding='utf-8') as fh:
            fh.write(text_out)
    except Exception:
        pass


# ══════════════════════════════════════════════════════════════════════════
#  界面
# ══════════════════════════════════════════════════════════════════════════
def cfg_path():
    base = os.environ.get('APPDATA') or os.path.expanduser('~')
    d = os.path.join(base, APP_NAME)
    try:
        os.makedirs(d, exist_ok=True)
    except Exception:
        return None
    return os.path.join(d, 'config.json')


def load_cfg():
    p = cfg_path()
    if p and os.path.exists(p):
        try:
            return json.load(open(p, encoding='utf-8'))
        except Exception:
            pass
    return {}


def save_cfg(cfg):
    p = cfg_path()
    if not p:
        return
    try:
        json.dump(cfg, open(p, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    except Exception:
        pass


def base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def desktop_dir():
    """桌面真实路径（不同电脑桌面位置可能不同，用系统 API 取）"""
    try:
        import ctypes
        buf = ctypes.create_unicode_buffer(1024)
        # CSIDL_DESKTOPDIRECTORY = 0x0010
        if ctypes.windll.shell32.SHGetFolderPathW(None, 0x0010, None, 0, buf) == 0:
            return buf.value
    except Exception:
        pass
    return os.path.join(os.path.expanduser('~'), 'Desktop')


def write_text_report(out_path, res):
    """纯文本报告：给不想开 Excel 的时候用"""
    rep, meta = res['report'], res['meta']
    lines = ['建档表检查报告', '=' * 60,
             '被检查文件：%s' % meta['file'],
             '编码库    ：%s' % (meta['lib'] or '（未提供）'),
             '检查时间  ：%s' % meta['time'],
             '数据行/组合：%d 行 / %d 组' % (meta['rows'], meta['groups']),
             '错误 %d ｜ 警告 %d ｜ 提示 %d'
             % (rep.count(ERR), rep.count(WARN), rep.count(INFO)),
             '结论：%s' % ('可以上传' if rep.count(ERR) == 0
                         else '先改完 %d 处错误再上传' % rep.count(ERR)),
             '=' * 60]
    for it in rep.sorted():
        if it['level'] == OK:
            continue
        lines.append('[%s] %s%s %s' % (it['level'], row_label(it['row']), it['col'], it['item']))
        lines.append('    %s' % it['msg'])
        if it['cur']:
            lines.append('    现值：%s' % it['cur'])
        if it['advice']:
            lines.append('    怎么改：%s' % it['advice'])
    with open(out_path, 'w', encoding='utf-8') as fh:
        fh.write('\n'.join(lines) + '\n')
    return out_path


def guess_output_dir():
    """找生成器的「输出」文件夹：先看自己旁边，再看同级目录里谁的 输出 最新"""
    base = base_dir()
    direct = [os.path.join(base, '输出'),
              os.path.join(base, '..', '输出'),
              os.path.join(base, '..', '工具-SKU建档生成', '输出')]
    for cand in direct:
        cand = os.path.abspath(cand)
        if os.path.isdir(cand):
            return cand
    parent = os.path.dirname(base)
    best, best_t = base, 0.0
    try:
        for name in os.listdir(parent):
            sub = os.path.join(parent, name, '输出')
            if os.path.isdir(sub):
                t = os.path.getmtime(sub)
                if t > best_t:
                    best, best_t = sub, t
    except Exception:
        pass
    return best


class App:
    def __init__(self, root):
        self.root = root
        root.title('%s  %s' % (APP_NAME, VER))
        root.geometry('980x740')
        root.minsize(900, 640)
        cfg = load_cfg()
        self.cfg = cfg
        self.res = None
        F = ('微软雅黑', 9)
        FB = ('微软雅黑', 10, 'bold')
        FT = ('微软雅黑', 15, 'bold')

        tk.Label(root, text='建档表检查器', font=FT, fg='#1F3864').pack(anchor='w', padx=16, pady=(14, 2))
        tk.Label(root, text='上传前先跑一遍：17 列结构、A/C 一致、B 列标签、J 编码、K/L/Q 公式，一次挑出来。',
                 font=('微软雅黑', 9), fg='#555').pack(anchor='w', padx=16, pady=(0, 10))

        f1 = tk.Frame(root)
        f1.pack(fill='x', padx=16)
        tk.Label(f1, text='① 建档表', font=FB, fg='#1F3864').pack(side='left')
        self.doc_var = tk.StringVar(value=cfg.get('doc', ''))
        tk.Entry(f1, textvariable=self.doc_var, font=F).pack(side='left', fill='x', expand=True, padx=8)
        tk.Button(f1, text='选择…', font=F, command=self.pick_doc).pack(side='left')

        f2 = tk.Frame(root)
        f2.pack(fill='x', padx=16, pady=(8, 0))
        tk.Label(f2, text='② 编码库', font=FB, fg='#1F3864').pack(side='left')
        default_lib = cfg.get('lib', '')
        if not default_lib and os.path.exists(DEFAULT_LIB):
            default_lib = DEFAULT_LIB
        self.lib_var = tk.StringVar(value=default_lib)
        tk.Entry(f2, textvariable=self.lib_var, font=F).pack(side='left', fill='x', expand=True, padx=8)
        tk.Button(f2, text='选择…', font=F, command=self.pick_lib).pack(side='left')
        tk.Button(f2, text='清空', font=F, command=lambda: self.lib_var.set('')).pack(side='left', padx=(6, 0))

        f3 = tk.Frame(root)
        f3.pack(fill='x', padx=16, pady=(8, 0))
        tk.Label(f3, text='③ 参数', font=FB, fg='#1F3864').pack(side='left')
        tk.Label(f3, text='组间空行', font=F).pack(side='left', padx=(8, 2))
        self.gap_var = tk.StringVar(value=str(cfg.get('gap', GAP_DEFAULT)))
        tk.Spinbox(f3, from_=0, to=60, width=4, textvariable=self.gap_var, font=F).pack(side='left')
        self.style_var = tk.BooleanVar(value=cfg.get('style', True))
        tk.Checkbutton(f3, text='检查字体/字号/列宽', variable=self.style_var,
                       font=F).pack(side='left', padx=(10, 0))
        self.eo_var = tk.BooleanVar(value=cfg.get('strict_eo', True))
        tk.Checkbutton(f3, text='EO 整组严格（写了 EO 整组都得带 EO）',
                       variable=self.eo_var, font=F).pack(side='left', padx=(10, 0))
        tk.Label(f3, text='（编码库可选：不选只查结构和公式；选了才查编码是否存在、花色对不对）',
                 font=('微软雅黑', 8), fg='#777').pack(side='left', padx=(12, 0))

        f4 = tk.Frame(root)
        f4.pack(fill='x', padx=16, pady=(12, 6))
        self.btn_run = tk.Button(f4, text='开始检查', font=('微软雅黑', 10, 'bold'),
                                 bg='#1F3864', fg='white', width=12, command=self.run)
        self.btn_run.pack(side='left')
        self.btn_out = tk.Button(f4, text='导出检查报告', font=F, width=12, state='disabled',
                                 command=self.export)
        self.btn_out.pack(side='left', padx=(8, 0))
        tk.Button(f4, text='打开文件夹', font=F, width=10, command=self.open_dir).pack(side='left', padx=(8, 0))
        tk.Button(f4, text='创建桌面快捷方式', font=F, width=16,
                  command=self.make_shortcut).pack(side='left', padx=(8, 0))
        self.stat = tk.Label(f4, text='还没检查', font=F, fg='#555')
        self.stat.pack(side='left', padx=(16, 0))

        f5 = tk.Frame(root)
        f5.pack(fill='x', padx=16)
        tk.Label(f5, text='筛选', font=F).pack(side='left')
        self.filt_var = tk.StringVar(value='全部')
        cb = ttk.Combobox(f5, textvariable=self.filt_var, font=F, width=8, state='readonly',
                          values=['全部', ERR, WARN, INFO])
        cb.pack(side='left', padx=(6, 0))
        cb.bind('<<ComboboxSelected>>', lambda e: self.refill())
        tk.Label(f5, text='（双击一行看完整说明）', font=('微软雅黑', 8), fg='#777').pack(side='left', padx=(10, 0))

        box = tk.Frame(root)
        box.pack(fill='both', expand=True, padx=16, pady=(6, 12))
        cols = ('行号', '列', '级别', '检查项', '说明')
        self.tree = ttk.Treeview(box, columns=cols, show='headings', height=20)
        for c, w, anchor in (('行号', 60, 'center'), ('列', 55, 'center'), ('级别', 60, 'center'),
                             ('检查项', 150, 'w'), ('说明', 600, 'w')):
            self.tree.heading(c, text=c)
            self.tree.column(c, width=w, anchor=anchor, stretch=(c == '说明'))
        vs = ttk.Scrollbar(box, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=vs.set)
        self.tree.pack(side='left', fill='both', expand=True)
        vs.pack(side='right', fill='y')
        self.tree.tag_configure(ERR, foreground='#9C0006', background='#FDE7E9')
        self.tree.tag_configure(WARN, foreground='#9C6500', background='#FFF9E0')
        self.tree.tag_configure(INFO, foreground='#1F4E78')
        self.tree.tag_configure(OK, foreground='#006100')
        self.tree.bind('<Double-1>', self.show_detail)
        self.items_all = []

    # ── 选文件 ──
    def pick_doc(self):
        p = filedialog.askopenfilename(title='选择要检查的建档表',
                                       initialdir=self._start_dir(self.doc_var.get()),
                                       filetypes=[('Excel 表格', '*.xlsx *.xlsm'), ('全部', '*.*')])
        if p:
            self.doc_var.set(p)
            self.stat.config(text='已选建档表', fg='#555')

    def pick_lib(self):
        p = filedialog.askopenfilename(title='选择款式编码文档（①编码库）',
                                       initialdir=self._start_dir(self.lib_var.get()),
                                       filetypes=[('Excel 表格', '*.xlsx *.xlsm'), ('全部', '*.*')])
        if p:
            self.lib_var.set(p)

    def _start_dir(self, cur):
        if cur and os.path.dirname(cur) and os.path.isdir(os.path.dirname(cur)):
            return os.path.dirname(cur)
        for cand in (DEFAULT_ARCHIVE, guess_output_dir()):
            if cand and os.path.isdir(cand):
                return cand
        return base_dir()

    # ── 检查 ──
    def run(self):
        doc = self.doc_var.get().strip().strip('"')
        if not doc:
            messagebox.showwarning('还没选文件', '请先在上面的「① 建档表」里选一个 xlsx。')
            return
        if not os.path.exists(doc):
            messagebox.showerror('文件不存在', '找不到这个文件：\n%s' % doc)
            return
        lib = self.lib_var.get().strip().strip('"')
        if lib and not os.path.exists(lib):
            messagebox.showerror('文件不存在', '找不到编码库：\n%s' % lib)
            return
        try:
            gap = int(self.gap_var.get())
        except ValueError:
            gap = GAP_DEFAULT
        self.root.config(cursor='watch')
        self.stat.config(text='检查中…', fg='#555')
        self.root.update_idletasks()
        try:
            self.res = check(doc, lib or None, gap, bool(self.style_var.get()),
                             bool(self.eo_var.get()))
        except PermissionError:
            self.root.config(cursor='')
            self.stat.config(text='文件被占用', fg='#9C0006')
            messagebox.showerror('文件被占用',
                                 '这个文件现在读不了，多半正被 WPS / Excel 打开着。\n\n'
                                 '先保存并关闭它，再点一次「开始检查」。')
            return
        except Exception as e:  # noqa: BLE001
            self.root.config(cursor='')
            messagebox.showerror('检查出错了', '%s\n\n%s' % (e, traceback.format_exc(limit=3)))
            self.stat.config(text='出错', fg='#9C0006')
            return
        self.root.config(cursor='')
        rep = self.res['report']
        self.items_all = rep.sorted()
        self.refill()
        e, w, i = rep.count(ERR), rep.count(WARN), rep.count(INFO)
        msg = '%d 组 / %d 行 ｜ 错误 %d ｜ 警告 %d ｜ 提示 %d' % (
            self.res['meta']['groups'], self.res['meta']['rows'], e, w, i)
        self.stat.config(text=msg, fg=('#9C0006' if e else ('#9C6500' if w else '#006100')))
        self.btn_out.config(state='normal')
        try:
            save_cfg({'doc': doc, 'lib': lib, 'gap': gap,
                      'style': bool(self.style_var.get()),
                      'strict_eo': bool(self.eo_var.get())})
        except Exception:
            pass
        # 通过时不弹窗：底部那一行绿字就够了，弹窗还要多点一下

    def refill(self):
        self.tree.delete(*self.tree.get_children())
        f = self.filt_var.get()
        for it in self.items_all:
            if f != '全部' and it['level'] != f:
                continue
            self.tree.insert('', 'end', values=(it['row'] or '', it['col'], it['level'],
                                                it['item'], it['msg']),
                             tags=(it['level'],))

    def show_detail(self, _evt=None):
        sel = self.tree.selection()
        if not sel:
            return
        vals = self.tree.item(sel[0], 'values')
        row, col, level, item, msg = vals
        for it in self.items_all:
            if it['item'] == item and str(it['row'] or '') == str(row) and it['col'] == col:
                detail = '级别：%s\n位置：%s%s\n检查项：%s\n\n%s' % (
                    level, ('第 %s 行 ' % row) if row else '', col, item, it['msg'])
                if it['cur']:
                    detail += '\n\n单元格现值：%s' % it['cur']
                if it['advice']:
                    detail += '\n\n怎么改：%s' % it['advice']
                messagebox.showinfo('问题详情', detail)
                return

    # ── 导出 ──
    def export(self):
        if not self.res:
            return
        doc = self.doc_var.get().strip().strip('"')
        d = os.path.dirname(doc) or base_dir()
        stem = os.path.splitext(os.path.basename(doc))[0]
        stamp = datetime.datetime.now().strftime('%m%d-%H%M')
        p = filedialog.asksaveasfilename(
            title='保存检查报告', initialdir=d,
            initialfile='检查报告-%s-%s.xlsx' % (stem, stamp),
            defaultextension='.xlsx',
            filetypes=[('Excel 表格（带标色）', '*.xlsx'), ('纯文本', '*.txt')])
        if not p:
            return
        try:
            if p.lower().endswith('.txt'):
                write_text_report(p, self.res)
            else:
                write_report(p, self.res)
        except Exception as e:  # noqa: BLE001
            messagebox.showerror('导出失败', '%s\n\n%s' % (e, traceback.format_exc(limit=3)))
            return
        if messagebox.askyesno('报告已生成', '已保存：\n%s\n\n要现在打开吗？' % p):
            self._open(p)

    def make_shortcut(self):
        """在桌面建快捷方式（桌面位置因电脑而异，用系统 API 取真实路径）"""
        exe = sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(__file__)
        if not getattr(sys, 'frozen', False):
            messagebox.showinfo('还没打包',
                                '源码模式不建快捷方式。\n打包成 exe 之后再点这个按钮。')
            return
        desktop = desktop_dir()
        if not desktop or not os.path.isdir(desktop):
            messagebox.showerror('找不到桌面', '取不到桌面目录，请手动把 exe 发送到桌面。')
            return
        lnk = os.path.join(desktop, '%s.lnk' % APP_NAME)
        ps = ('$s=(New-Object -ComObject WScript.Shell).CreateShortcut("%s");'
              '$s.TargetPath="%s";$s.WorkingDirectory="%s";$s.Description="检查拼多多 17 列建档表";'
              '$s.Save()' % (lnk, exe, os.path.dirname(exe)))
        try:
            subprocess.run(['powershell', '-NoProfile', '-Command', ps],
                           capture_output=True, timeout=30,
                           creationflags=0x08000000)  # CREATE_NO_WINDOW
            if os.path.exists(lnk):
                messagebox.showinfo('好了', '桌面快捷方式已创建：\n%s' % lnk)
            else:
                messagebox.showerror('没建成', '快捷方式没写出来，请手动把 exe 发送到桌面。')
        except Exception as e:  # noqa: BLE001
            messagebox.showerror('没建成', '%s\n\n请手动把 exe 发送到桌面。' % e)

    def open_dir(self):
        doc = self.doc_var.get().strip().strip('"')
        d = os.path.dirname(doc) if doc else guess_output_dir()
        if d and os.path.isdir(d):
            self._open(d)

    def _open(self, path):
        try:
            os.startfile(path)  # noqa: S606
        except Exception:
            messagebox.showinfo('打开失败', '请手动打开：\n%s' % path)


def crash_log(err_text):
    """打包成窗口程序后，未捕获的异常会弹系统对话框卡住；所以自己落一份日志"""
    try:
        base = base_dir()
        p = os.path.join(base, '检查器-错误日志.txt')
        with open(p, 'a', encoding='utf-8') as fh:
            fh.write('\n===== %s =====\n%s\n'
                     % (datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'), err_text))
        return p
    except Exception:
        return ''


def main():
    cli_mode = len(sys.argv) > 1 and any(x.split('=')[0] == '--check' for x in sys.argv[1:])
    try:
        if cli_mode:
            sys.exit(cli(sys.argv[1:]))
        root = tk.Tk()
        App(root)
        root.mainloop()
    except SystemExit:
        raise
    except Exception:  # noqa: BLE001
        p = crash_log(traceback.format_exc())
        if cli_mode:
            # 命令行模式下绝不能弹窗：打包成窗口程序后弹窗没人点，会一直挂着
            sys.exit(2)
        try:
            messagebox.showerror('程序出错了',
                                 '%s\n\n详细情况写在：\n%s' % (traceback.format_exc(limit=2), p))
        except Exception:
            pass
        sys.exit(2)


if __name__ == '__main__':
    main()
