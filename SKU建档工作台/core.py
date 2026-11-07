# -*- coding: utf-8 -*-
"""
SKU 建档工作台 v3.0
三个页签：
  ① 编码库 —— 粘贴聚水潭编码，实时识别，一键写入「①编码库」
  ② 生成   —— 写 SKU 组合，参数，生成预览
  ③ 预览   —— 表格里直接改，确认无误再导出 Excel
"""
import os
import re
import sys
import json
import shutil
import tempfile
import datetime
import subprocess
import traceback

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

APP_NAME = 'SKU建档工作台'
VER = 'v3.2'
GAP_ROWS = 10

COLS = ['组合商品编码', '组合装商品标签', '组合款式编码', '组合商品名称', '组合商品简称',
        '组合商品实体编码', '虚拟分类', '组合颜色规格', '禁止库存同步', '商品编码',
        '数量', '应占售价', '组合成本价', '是否支持留言备注换货', '图片', '品牌', None]
RED_HEAD = {1, 3, 10, 11, 12}
WIDTH = {'A': 39.375, 'B': 23.875, 'C': 39.375, 'D': 9.625, 'F': 12.625,
         'G': 6.625, 'H': 8.125, 'J': 21.625, 'K': 3.625, 'L': 6.625,
         'M': 8.125, 'N': 15.625, 'O': 3.625, 'Q': 9.0}
SZ_LETTER = ['XS', 'S', 'M', 'L', 'XL', '2XL', '3XL', '4XL', '5XL', '6XL']
HEIGHTS = ['80', '90', '100', '110', '120', '130', '140', '150', '160', '170', '180']

thin = Side(style='thin', color='000000')
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
F_HEAD = Font(name='微软雅黑', size=9, color='FF000000')
F_HEAD_RED = Font(name='微软雅黑', size=9, color='FFFF0000')
F_BODY = Font(name='微软雅黑', size=9)
F_CODE = Font(name='宋体', size=11)
F_MISS = Font(name='微软雅黑', size=9, color='C00000', bold=True)
FILL_MISS = PatternFill('solid', fgColor='FFC7CE')
FILL_CHECK = PatternFill('solid', fgColor='FFF2CC')
F_CHECK = Font(name='微软雅黑', size=9, color='C00000', bold=True)
AL_HDR = Alignment(horizontal='left', vertical='center', wrap_text=True)
AL_BODY = Alignment(horizontal='left', vertical='center')

PAT = re.compile(r'^([A-Z]+)(\d{3,7})([A-Z]*)T(\d)S(\d)(.+)$')
HEI = re.compile(r'^([A-Z][A-Z0-9]*?)(\d{3})$')
VEST = re.compile(r'^([A-Z]+)(\d{3,5})([A-Z]{2,6}?)(XXXL|XXL|XL|XS|S|M|L)$')
CJK = re.compile(r'[\u4e00-\u9fa5]')

Q_NOTE = ('当子商品编码出现两个"T"时，或空版款号、德绒套装等，需手动核查计算。\n'
          '本格由公式 =LEFT(TEXTAFTER(J,"T"),1)*K 得出，取的是第一个 T 后面的字符；\n'
          '编码里若有第二个 T（或花色码以 T 结尾），这里会是 #VALUE!，请手工改为正确条数。')


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


def q_risk(code):
    if not code:
        return '无编码'
    i = code.find('T')
    if i < 0:
        return '编码里没有 T 段'
    if not code[i + 1:i + 2].isdigit():
        return '编码出现两个 T（花色码含 T）'
    if not re.match(r'^T\dS\d', code[i:]):
        return 'T 段后不是 S+数字'
    return None


def t_count(code):
    m = PAT.match(code or '')
    return int(m.group(4)) if m else 0


def spec_flavor(v):
    if v is None:
        return ''
    for seg in str(v).split(';'):
        seg = seg.strip()
        if seg and re.fullmatch(r'[\u4e00-\u9fa5]{1,6}', seg):
            return seg
    return ''


def find_header(ws, max_scan=15):
    for r in range(1, min(max_scan, ws.max_row) + 1):
        row = {}
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if v is not None and str(v).strip():
                row[str(v).strip()] = c
        if '商品编码' in row:
            return r, row
    return None, None


def parse_lib(ws):
    hr, cols = find_header(ws)
    if not hr:
        return None, None, None, ['①编码库 里找不到「商品编码」这一列的表头']
    book, groups, nonstd, issues = {}, [], [], []
    c_code = cols['商品编码']
    c_name, c_no, c_sz = cols.get('花色'), cols.get('货号'), cols.get('尺码')
    for r in range(hr + 1, ws.max_row + 1):
        raw = ws.cell(r, c_code).value
        if raw is None:
            continue
        code = str(raw).strip()
        if not code:
            continue
        no, sz, fam = parse_code(code)
        if not fam:
            nonstd.append('第%d行 %s' % (r, code))
            continue
        nm = str(ws.cell(r, c_name).value).strip() if c_name and ws.cell(r, c_name).value else ''
        if c_no:
            v = ws.cell(r, c_no).value
            if v is not None and str(v).strip() not in ('', no):
                issues.append('第%d行 货号填的是 %s，但编码里是 %s' % (r, str(v).strip(), no))
        if c_sz:
            v = ws.cell(r, c_sz).value
            if v is not None and str(v).strip() and norm_size(str(v).strip()) != sz:
                issues.append('第%d行 尺码填的是 %s，但编码里是 %s' % (r, str(v).strip(), sz))
        key = (no, nm, sz)
        if key in book and book[key] != code:
            issues.append('冲突：%s-%s-%s 已有 %s，第%d行又写了 %s'
                          % (no, nm or '空版', sz, book[key], r, code))
            continue
        book[key] = code
        if (no, nm) not in groups:
            groups.append((no, nm))
    return book, groups, nonstd, issues


def parse_paste(ws):
    max_r, max_c = ws.max_row, min(ws.max_column, 30)
    grid = [[('' if ws.cell(r, c).value is None else str(ws.cell(r, c).value).strip())
             for c in range(1, max_c + 1)] for r in range(1, max_r + 1)]
    return parse_grid(grid)


def parse_grid(grid):
    max_r = len(grid)
    max_c = len(grid[0]) if max_r else 0
    info = {}
    for r in range(1, max_r + 1):
        for c in range(1, max_c + 1):
            s = grid[r - 1][c - 1]
            if not s:
                continue
            no, sz, fam = parse_code(s)
            if fam:
                info[(r, c)] = (no, sz, fam, s)

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
            hdrs[(r, c)] = name

    row_hdr, col_hdr = {}, {}
    for (r, c), nm in hdrs.items():
        row_hdr.setdefault(r, []).append((c, nm))
        col_hdr.setdefault(c, []).append((r, nm))

    def resolve(r, c):
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
        return None

    def row_cat(r):
        for c in range(1, max_c + 1):
            s = grid[r - 1][c - 1]
            if not s or (r, c) in info or (r, c) in hdrs:
                continue
            if ';' in s or not CJK.search(s) or len(s) > 12:
                continue
            if re.fullmatch(r'[\d.\s]+', s):
                continue
            return s
        return None

    book, groups, nonstd, cats = {}, [], [], {}
    for (r, c), (no, sz, fam, code) in sorted(info.items()):
        if fam == 'inner':
            nm = resolve(r, c)
            if nm is None:
                nonstd.append('第%d行 %s（找不到所属花色）' % (r, code))
                continue
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
                nm = resolve(r, c) or ''
        key = (no, nm, sz)
        if key in book and book[key] != code:
            nonstd.append('第%d行 与已有编码冲突：%s' % (r, code))
            continue
        book[key] = code
        ct = row_cat(r)
        if ct:
            cats[key] = ct
        if (no, nm) not in groups:
            groups.append((no, nm))

    fam_names = {}
    for (no, nm, sz), code in list(book.items()):
        if parse_code(code)[2] in ('height', 'vest'):
            fam_names.setdefault(no, set()).add(nm)
    for no, names in fam_names.items():
        real = [x for x in names if x]
        if len(real) == 1:
            target = real[0]
            for (n2, nm2, sz2), code in list(book.items()):
                if n2 == no and nm2 == '':
                    del book[(n2, nm2, sz2)]
                    book[(n2, target, sz2)] = code
                    if (n2, nm2, sz2) in cats:
                        cats[(n2, target, sz2)] = cats.pop((n2, nm2, sz2))
            groups = list(dict.fromkeys(
                (a, target if (a == no and b == '') else b) for a, b in groups))
    return book, groups, nonstd, cats


def parse_text_paste(text):
    """把编辑器里粘的文本切成网格再解析（制表符分列，空行也保留）"""
    lines = text.replace('\r\n', '\n').replace('\r', '\n').split('\n')
    grid = [[c.strip() for c in ln.split('\t')] for ln in lines]
    width = max((len(r) for r in grid), default=0)
    if not width:
        return {}, [], [], {}
    grid = [r + [''] * (width - len(r)) for r in grid]
    return parse_grid(grid)


def load_doc(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    if '①编码库' in wb.sheetnames:
        book, groups, nonstd, issues = parse_lib(wb['①编码库'])
        return book, groups, nonstd, issues, '编码库'
    book, groups, nonstd, cats = parse_paste(wb.worksheets[0])
    return book, groups, nonstd, [], '旧格式(%s)' % wb.worksheets[0].title


def refresh_summary(wb, lws):
    if '③款号总览' not in wb.sheetnames:
        return
    hr, cols = find_header(lws)
    if not hr:
        return
    c_code, c_name, c_cat = cols['商品编码'], cols.get('花色'), cols.get('品类')
    agg = {}
    for r in range(hr + 1, lws.max_row + 1):
        v = lws.cell(r, c_code).value
        if v is None:
            continue
        code = str(v).strip()
        no, sz, fam = parse_code(code)
        if not fam:
            continue
        nm = str(lws.cell(r, c_name).value).strip() if c_name and lws.cell(r, c_name).value else ''
        cat = str(lws.cell(r, c_cat).value).strip() if c_cat and lws.cell(r, c_cat).value else ''
        a = agg.setdefault((cat, no, nm), {'pat': set(), 'sz': set(), 'fam': fam, 'n': 0})
        if fam == 'inner':
            a['pat'].add(PAT.match(code).group(3))
        a['sz'].add(sz)
        a['n'] += 1
    s = wb['③款号总览']
    for r in range(5, max(s.max_row, 5) + 1):
        for c in range(1, 9):
            s.cell(r, c).value = None
    r = 5
    for (cat, no, nm), a in sorted(agg.items(),
                                   key=lambda kv: (kv[0][0] == '', kv[0][0], kv[0][1], kv[0][2])):
        szs = sorted(a['sz'], key=size_key)
        label = {'inner': '内裤码', 'height': '身高码', 'vest': '背心码'}[a['fam']]
        vals = [cat, no, nm or '（空版）', '、'.join(sorted(x for x in a['pat'] if x)),
                (szs[0] + ' ~ ' + szs[-1]) if szs else '', a['n'], label, None]
        for i, v in enumerate(vals, 1):
            c = s.cell(r, i, v)
            c.font = Font(name='微软雅黑', size=9)
            c.border = BORDER
            c.alignment = Alignment(horizontal='left', vertical='center')
        r += 1


def write_library(path, pbook, pcats, pcat=''):
    """把解析出来的编码写进「①编码库」，返回统计"""
    wb = openpyxl.load_workbook(path)
    if '①编码库' not in wb.sheetnames:
        raise ValueError('这份文档没有「①编码库」这张表。')
    lws = wb['①编码库']
    hr, cols = find_header(lws)
    if not hr:
        raise ValueError('①编码库 里找不到「商品编码」表头。')
    c_code = cols['商品编码']
    c_name, c_cat = cols.get('花色'), cols.get('品类')

    existing, keymap = {}, {}
    for r in range(hr + 1, lws.max_row + 1):
        v = lws.cell(r, c_code).value
        if v is None:
            continue
        v = str(v).strip()
        existing[v] = r
        no, sz, fam = parse_code(v)
        if fam:
            nm0 = str(lws.cell(r, c_name).value).strip() if c_name and lws.cell(r, c_name).value else ''
            keymap[(no, nm0, sz)] = v

    added, dup, filled, conflicts = 0, 0, 0, []
    last = lws.max_row
    for (no, nm, sz), code in sorted(pbook.items()):
        cat = pcats.get((no, nm, sz)) or pcat or None
        if code in existing:
            dup += 1
            if cat and c_cat:
                er = existing[code]
                if not lws.cell(er, c_cat).value:
                    lws.cell(er, c_cat, cat)
                    lws.cell(er, c_cat).font = Font(name='微软雅黑', size=9)
                    lws.cell(er, c_cat).border = BORDER
                    lws.cell(er, c_cat).alignment = Alignment(horizontal='left', vertical='center')
                    filled += 1
            continue
        k = (no, nm, sz)
        if k in keymap and keymap[k] != code:
            conflicts.append('%s-%s-%s：库里是 %s，新粘的是 %s'
                             % (no, nm or '空版', sz, keymap[k], code))
            continue
        last += 1
        m = PAT.match(code)
        vals = {'品类': cat, '货号': no, '花色': nm or None, '尺码': sz,
                '商品编码': code, '花色码': (m.group(3) or None) if m else None}
        for name, val in vals.items():
            if name in cols:
                lws.cell(last, cols[name], val)
        for c in range(1, lws.max_column + 1):
            cc = lws.cell(last, c)
            cc.font = Font(name='宋体' if c == c_code else '微软雅黑',
                           size=10 if c == c_code else 9)
            cc.border = BORDER
            cc.alignment = Alignment(horizontal='left', vertical='center')
        added += 1

    if added or filled:
        refresh_summary(wb, lws)
        wb.save(path)
    return {'added': added, 'dup': dup, 'filled': filled, 'conflicts': conflicts,
            'total': len(existing) + added}


def _is_marker(tok):
    return bool(re.fullmatch(r'[A-Za-z]{1,8}', tok or ''))


def parse_combo(text):
    """→ (raw, n, mid_txt, sizes)
       raw: [(货号, 花色部分)]，花色部分可能含 '-'，也可能是 '' / 'ALLIN' / 'EO-ALLIN'"""
    parts = [p.strip() for p in re.split(r'/', text) if p.strip()]
    if len(parts) < 3:
        raise ValueError('格式应为：货号-花色+花色 / N条装 / 尺码1,尺码2\n这行是：%s' % text)
    head = parts[0]
    m_tail = re.search(r'-(\d+X*L|X+L|[MS]|\d{3})\s*$', head, re.I)
    if m_tail and looks_like_size(m_tail.group(1)):
        head = head[:m_tail.start()]

    raw, cur = [], None
    for tok in head.split('+'):
        t = tok.strip()
        if not t:
            continue
        if looks_like_size(t) and not re.match(r'^[A-Za-z]*\d{3,}', t):
            raise ValueError('组合「%s」里把尺码「%s」写进组合部分了。\n\n'
                             '尺码只写在第二个斜杠后面：\n　　%s / %s / %s'
                             % (text, t, head.replace('+' + t, '').replace(t + '+', ''),
                                parts[1], parts[2]))
        m = re.match(r'^([A-Za-z]*\d{3,})(?:-(.+))?$', t)
        if m:
            cur = m.group(1)
            raw.append((cur, (m.group(2) or '').strip()))
            continue
        if cur is None:
            raise ValueError('这一行第一段看不懂：%s\n应写成「货号-花色」或纯「货号」' % text)
        raw.append((cur, t))

    mid = parts[1]
    m = re.search(r'(\d+)\s*(条|件|套)?装?', mid)
    if m:
        n = int(m.group(1))
        unit = m.group(2) or '条'
        mid_txt = '%d%s装' % (n, unit)
    else:
        n, mid_txt = None, mid
    sizes = []
    for s_ in re.split(r'[,，、\s]+', parts[2]):
        u = s_.strip().upper()
        if u and u not in sizes:
            sizes.append(u)
    sizes.sort(key=size_key)
    return raw, n, mid_txt, sizes


def build_rows(codebook, combos_text, product_name, tag_name, gap=GAP_ROWS, sep='，'):
    lines = [l.strip() for l in combos_text.splitlines()
             if l.strip() and not l.strip().startswith('#')]
    if not lines:
        raise ValueError('没有填组合。')

    by_no_size, names_of = {}, {}
    for (no, nm, s), c in codebook.items():
        by_no_size.setdefault((no, s), set()).add(c)
        lst = names_of.setdefault(no, [])
        if nm not in lst:
            lst.append(nm)          # 保持编码库里的出现顺序

    def variants(no, part):
        """返回可能的花色名（精确优先，再试标记挪位）"""
        lst = names_of.get(no, [])
        part = (part or '').strip()
        if part in lst:
            return part
        ps = [x for x in part.split('-') if x]
        # 两头都写了标记（EO-粉色-EO）→ 掐掉两头再找
        if len(ps) >= 3 and _is_marker(ps[0]) and ps[0] == ps[-1]:
            return variants(no, '-'.join(ps[1:-1]))
        if len(ps) >= 2 and _is_marker(ps[0]):
            alt = '-'.join(ps[1:] + [ps[0]])
            if alt in lst:
                return alt
        if len(ps) >= 2 and _is_marker(ps[-1]):
            alt = '-'.join([ps[-1]] + ps[:-1])
            if alt in lst:
                return alt
        return None

    def lookup(no, nm, s):
        if nm:
            return codebook.get((no, nm, norm_size(s)))
        codes = by_no_size.get((no, norm_size(s)), set())
        return next(iter(codes)) if len(codes) == 1 else None

    def label(no, nm):
        return '%s-%s' % (no, nm) if nm else '%s（空版）' % no

    def hint(no, nm):
        if no not in names_of:
            return '%s 这个款式在编码文档里没有' % no
        ns = [x for x in names_of[no] if x]
        if not ns:
            return '%s 只有空版，组合里写「%s」即可' % (no, no)
        if nm and nm not in ns:
            more = '…' if len(ns) > 12 else ''
            return '%s 现有花色：%s%s' % (no, '、'.join(ns[:12]), more)
        return ''

    rows, warns, check_cells, miss_cells = [], [], [], []
    for idx, line in enumerate(lines):
        raw, n, mid_txt, sizes = parse_combo(line)
        # ---- 展开成真正的组件 ----
        style, marker = 'plain', None
        for _no, part in raw:
            if not part:
                continue
            ps = part.split('-')
            if len(ps) >= 2 and _is_marker(ps[0]):
                style, marker = 'prefix', ps[0]
                break
            if len(ps) >= 2 and _is_marker(ps[-1]):
                style, marker = 'suffix', ps[-1]
                break
        if style == 'plain':
            for _no, part in raw:
                if part:
                    for p_ in part.split('-'):
                        if _is_marker(p_):
                            marker = p_
                            break
                if marker:
                    break

        comps, missing_parts = [], []
        for no, part in raw:
            if not part:
                comps.append((no, ''))
                continue
            ps = part.split('-')
            if ps and ps[-1].upper() == 'ALLIN':
                filters = [x for x in ps[:-1] if x]
                lst = [x for x in names_of.get(no, []) if x]
                if not lst:
                    missing_parts.append((no, part))
                    continue
                if filters:
                    lst = [x for x in lst
                           if all(any(f.upper() == q.upper() for q in x.split('-'))
                                  for f in filters)]
                elif marker:
                    # 这一行声明了 EO，ALLIN 就只取带 EO 的那些颜色
                    lst = [x for x in lst if marker in x.split('-')]
                if not lst:
                    missing_parts.append((no, part))
                    continue
                for x in lst:
                    if (no, x) not in comps:
                        comps.append((no, x))
                continue
            nm = variants(no, part)
            if nm is None and marker \
                    and not any(_is_marker(q) for q in part.split('-')):
                nm = variants(no, part + '-' + marker)
            # 写了标记（EO）就整组都套上：一个组合里不能一半 EO 一半原版
            if nm is not None and marker and marker not in nm.split('-'):
                forced = variants(no, nm + '-' + marker) or variants(no, marker + '-' + nm)
                if forced:
                    nm = forced
                elif not any(_is_marker(q) for q in nm.split('-')):
                    warns.append('%s-%s 在编码库里没有带 %s 的版本，这一组只能按不带 %s 的建'
                                 % (no, nm, marker, marker))
            if nm is None:
                comps.append((no, part))
            elif (no, nm) not in comps:
                comps.append((no, nm))

        miss = [c for c in comps if not any(lookup(c[0], c[1], s_) for s_ in sizes)]
        if len(miss) > 2:
            detail = []
            for c in miss:
                h = hint(*c)
                detail.append('  %s%s' % (label(*c), ('　← ' + h) if h else ''))
            raise ValueError('组合「%s」有 %d 个款号/花色查不到，超过上限 2：\n\n%s\n\n'
                             '（提示：花色名里的破折号要照编码库写，比如「粉色-EO」；'
                             '写「EO-粉色」也能认。）'
                             % (line, len(miss), '\n'.join(detail)))
        if missing_parts:
            for no, part in missing_parts:
                warns.append('ALLIN 展开失败：%s-%s 在编码库里没有匹配的花色（%s 现有：%s）'
                             % (no, part, no, '、'.join([x for x in names_of.get(no, []) if x][:10])))

        tot = 0
        for no, nm in comps:
            c = None
            for s_ in sizes:
                c = lookup(no, nm, s_)
                if c:
                    break
            tot += t_count(c)
        if miss:
            for c in miss:
                h = hint(*c)
                warns.append('查不到 %s%s' % (label(*c), ('　← ' + h) if h else ''))
        else:
            avail = set(norm_size(s_) for s_ in sizes)
            for no, nm in comps:
                avail &= set(norm_size(s_) for s_ in sizes if lookup(no, nm, s_))
            gone = [s_ for s_ in sizes if norm_size(s_) not in avail]
            if gone:
                warns.append('「%s」编码文档里没有 %s，这些行已留空标红'
                             % (line[:22] + '…', '、'.join(gone)))
        if n is not None and tot and tot != n:
            warns.append('「%s」文案写 %d，但编码里 T 合计是 %d 条'
                         % (line[:22] + '…', n, tot))

        if idx > 0 and gap > 0:
            for _ in range(gap):
                rows.append({'blank': True})

        # ---- A 列：按输入风格渲染 ----
        def show(no, nm, first):
            if not nm:
                return no if first else ''
            if style == 'prefix' and marker:
                base = '-'.join(p_ for p_ in nm.split('-')
                                if p_.upper() != marker.upper())
                if first:
                    return '%s-%s-%s' % (no, marker, base)
                return base
            if first:
                return '%s-%s' % (no, nm)
            return nm

        a_base, prev = '', None
        for i_, (no, nm) in enumerate(comps):
            tok = show(no, nm, no != prev)
            a_base += tok + '+'
            prev = no
        a_base = a_base[:-1]

        for s_ in sizes:
            a = a_base + '-' + s_
            lb = '%s,%s%s%s' % (tag_name, norm_size(s_), sep, mid_txt)
            for no, nm in comps:
                code = lookup(no, nm, s_)
                fam = parse_code(code)[2] if code else None
                q = 1 if fam in ('height', 'vest') else \
                    '=LEFT(_xlfn.TEXTAFTER(J{c},"T"),1)*K{c}'
                risk = q_risk(code) if (code and fam == 'inner') else None
                rows.append({'blank': False, 'gid': idx, 'a': a, 'b': lb, 'c': a,
                             'd': product_name, 'j': code or '', 'k': 1, 'q': q,
                             'kind': fam, 'miss': code is None, 'risk': risk})
                if code is None:
                    miss_cells.append('%s（%s 码）' % (label(no, nm), s_))
                elif risk:
                    check_cells.append('%s：%s' % (code, risk))
    return {'rows': rows, 'warns': warns, 'miss': miss_cells, 'check': check_cells}


def export_xlsx(rows, path):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Sheet1'
    for col, w in WIDTH.items():
        ws.column_dimensions[col].width = w
    ws.row_dimensions[1].height = 31
    for i, h in enumerate(COLS, 1):
        c = ws.cell(1, i, h)
        c.font = F_HEAD_RED if i in RED_HEAD else F_HEAD
        c.alignment = AL_HDR
        c.border = BORDER
    ws.cell(1, 17).border = Border()
    ws.cell(1, 17).font = Font(name='宋体', size=11)
    ws.cell(1, 17).alignment = Alignment()

    # 组内 Q 合计，用于算应占售价的显示值
    groups, cur = {}, None
    for rw in rows:
        if rw['blank']:
            continue
        cur = rw['a']
        groups.setdefault(cur, []).append(rw)

    r = 2
    max_a = 0
    for rw in rows:
        if rw['blank']:
            r += 1
            continue
        a = rw['a']
        max_a = max(max_a, sum(2 if ord(ch) > 127 else 1 for ch in a))
        ws.cell(r, 1, a)
        ws.cell(r, 2, rw['b'])
        ws.cell(r, 3, a)
        ws.cell(r, 4, rw['d'])
        ws.cell(r, 11, rw['k'])
        ws.cell(r, 12, '=ROUND(Q%d/SUMIFS(Q:Q,A:A,A%d),4)' % (r, r))
        q = rw['q']
        if isinstance(q, str):
            q = q.replace('{c}', str(r))
        ws.cell(r, 17, q)
        for i in range(1, 18):
            if i in (13, 14, 15, 16):
                ws.cell(r, i).font = Font(name='宋体', size=11)
                continue
            cc = ws.cell(r, i)
            cc.font = F_BODY
            cc.alignment = AL_BODY
            cc.border = BORDER
        jc = ws.cell(r, 10, rw['j'] or None)
        jc.font = F_CODE
        jc.border = Border()
        jc.alignment = Alignment(vertical='center')
        if rw['miss'] or not rw['j']:
            jc.fill = FILL_MISS
            jc.font = F_MISS
            jc.border = BORDER
        elif rw['risk']:
            qc = ws.cell(r, 17)
            qc.fill = FILL_CHECK
            qc.font = F_CHECK
            ts = ''
            m = PAT.match(rw['j'])
            if m:
                ts = 'T%sS%s，正确条数应为 %d' % (m.group(4), m.group(5), t_count(rw['j']))
            qc.comment = Comment(Q_NOTE + ('\n\n本行编码末尾是 %s。' % ts if ts else ''),
                                 'SKU建档工作台', width=320, height=190)
        r += 1
    aw = max(39.375, min(max_a * 1.15 + 2, 100))
    ws.column_dimensions['A'].width = aw
    ws.column_dimensions['C'].width = aw
    wb.save(path)


def cfg_path():
    d = os.path.join(os.environ.get('APPDATA') or os.path.expanduser('~'), APP_NAME)
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
    try:
        out = subprocess.run(
            ['powershell', '-NoProfile', '-Command', "[Environment]::GetFolderPath('Desktop')"],
            capture_output=True, text=True, timeout=15,
            creationflags=0x08000000)
        d = (out.stdout or '').strip()
        if d and os.path.isdir(d):
            return d
    except Exception:
        pass
    for name in ('Desktop', '桌面'):
        d = os.path.join(os.path.expanduser('~'), name)
        if os.path.isdir(d):
            return d
    return os.path.expanduser('~')
