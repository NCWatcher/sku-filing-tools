# -*- coding: utf-8 -*-
"""SKU 建档工作台 —— 界面层（依赖 core.py）"""
import os
import re
import sys
import json
import datetime
import subprocess

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import core as C


class App:
    def __init__(self, root):
        self.root = root
        root.title('%s  %s' % (C.APP_NAME, C.VER))
        root.geometry('1000x760')
        root.minsize(900, 680)
        self.cfg = cfg = C.load_cfg()
        self.rows = []
        self.last_out = None
        F = ('微软雅黑', 9)
        FB = ('微软雅黑', 10, 'bold')

        nb = ttk.Notebook(root)
        nb.pack(fill='both', expand=True, padx=8, pady=(8, 0))
        self.nb = nb
        self.t1 = tk.Frame(nb)
        self.t2 = tk.Frame(nb)
        self.t3 = tk.Frame(nb)
        nb.add(self.t1, text='  ① 编码库  ')
        nb.add(self.t2, text='  ② 生成  ')
        nb.add(self.t3, text='  ③ 预览 / 导出  ')

        # ---------- 通用：编码文档路径 ----------
        f = tk.Frame(root)
        f.pack(fill='x', padx=12, pady=(8, 0))
        tk.Label(f, text='编码文档', font=FB, fg='#1F3864').pack(side='left')
        default_doc = os.path.join(C.base_dir(), '款式编码文档-模板.xlsx')
        self.doc_var = tk.StringVar(value=cfg.get('doc') or
                                    (default_doc if os.path.exists(default_doc) else ''))
        tk.Entry(f, textvariable=self.doc_var, font=F, state='readonly').pack(
            side='left', fill='x', expand=True, padx=8)
        tk.Button(f, text='选择…', font=F, command=self.pick_doc).pack(side='left')
        tk.Button(f, text='桌面快捷方式', font=F, command=self.make_shortcut).pack(side='left', padx=(8, 0))
        self.doc_info = tk.Label(root, text='', font=('微软雅黑', 8), fg='#888', anchor='w')
        self.doc_info.pack(fill='x', padx=12)

        self.status = tk.Label(root, text='就绪。', font=('微软雅黑', 9), fg='#333',
                               anchor='w', justify='left', wraplength=960)
        self.status.pack(fill='x', padx=12, pady=(4, 10))

        self._build_tab1(F)
        self._build_tab2(F)
        self._build_tab3(F)
        self.check_doc()

    # ============ Tab1 编码库 ============
    def _build_tab1(self, F):
        t = self.t1
        tk.Label(t, text='从聚水潭「按款式」框选复制，直接粘到下面（Ctrl+V），粘完会自动识别。',
                 font=('微软雅黑', 9), fg='#555').pack(anchor='w', padx=12, pady=(10, 2))
        tk.Label(t, text='格式随意：可以多列并排、可以带组名、可以是空版、可以一次粘多个品类。',
                 font=('微软雅黑', 8), fg='#888').pack(anchor='w', padx=12)

        wrap = tk.Frame(t)
        wrap.pack(fill='both', expand=True, padx=12, pady=(6, 0))
        self.paste_txt = tk.Text(wrap, font=('Consolas', 9), undo=True, relief='solid', bd=1)
        sb = tk.Scrollbar(wrap, command=self.paste_txt.yview)
        self.paste_txt.configure(yscrollcommand=sb.set)
        self.paste_txt.pack(side='left', fill='both', expand=True)
        sb.pack(side='right', fill='y')
        self.paste_txt.bind('<<Modified>>', self._on_paste_change)

        self.paste_info = tk.Label(t, text='（还没粘内容）', font=('微软雅黑', 9),
                                   fg='#555', anchor='w', justify='left')
        self.paste_info.pack(fill='x', padx=12, pady=(6, 0))

        f = tk.Frame(t)
        f.pack(fill='x', padx=12, pady=(6, 8))
        self.btn_write = tk.Button(f, text='写入编码库', font=('微软雅黑', 10, 'bold'),
                                   bg='#1F3864', fg='white', relief='flat',
                                   padx=16, pady=5, command=self.write_lib)
        self.btn_write.pack(side='left')
        tk.Button(f, text='清空', font=F, relief='flat', padx=12, pady=5,
                  command=lambda: self.paste_txt.delete('1.0', 'end')).pack(side='left', padx=8)
        tk.Button(f, text='刷新库统计', font=F, relief='flat', padx=12, pady=5,
                  command=self.check_doc).pack(side='left')
        self.lib_info = tk.Label(t, text='', font=('微软雅黑', 9), fg='#006100', anchor='w')
        self.lib_info.pack(fill='x', padx=12, pady=(0, 8))
        self._paste_job = None
        self._parsed = None

    def _on_paste_change(self, ev=None):
        if not self.paste_txt.edit_modified():
            return
        self.paste_txt.edit_modified(False)
        if self._paste_job:
            self.root.after_cancel(self._paste_job)
        self._paste_job = self.root.after(400, self.parse_paste_live)

    def parse_paste_live(self):
        self._paste_job = None
        text = self.paste_txt.get('1.0', 'end')
        if not text.strip():
            self.paste_info.config(text='（还没粘内容）', fg='#555')
            self._parsed = None
            return
        try:
            book, groups, nonstd, cats = C.parse_text_paste(text)
        except Exception as ex:
            self.paste_info.config(text='识别失败：%s' % ex, fg='#C00000')
            self._parsed = None
            return
        self._parsed = (book, cats)
        cats_cnt = {}
        for v in cats.values():
            cats_cnt[v] = cats_cnt.get(v, 0) + 1
        cat_txt = '、'.join('%s %d 条' % (k, v) for k, v in cats_cnt.items()) or '未识别到品类'
        self.paste_info.config(
            text='识别到 %d 个款、%d 个编码　｜　%s　%s' % (
                len(groups), len(book), cat_txt,
                ('　（%d 条看不懂已跳过）' % len(nonstd)) if nonstd else ''),
            fg='#006100' if book else '#C00000')

    def write_lib(self):
        p = self.doc_var.get()
        if not p or not os.path.exists(p):
            messagebox.showwarning(C.APP_NAME, '请先选择编码文档。')
            return
        if not self._parsed or not self._parsed[0]:
            messagebox.showwarning(C.APP_NAME, '上面还没有识别到编码，先粘贴再试。')
            return
        book, cats = self._parsed
        try:
            r = C.write_library(p, book, cats)
        except PermissionError:
            messagebox.showerror(C.APP_NAME, '编码文档正被 WPS/Excel 打开着，\n'
                                             '请先保存并关闭它，再点写入。')
            return
        except Exception as ex:
            messagebox.showerror(C.APP_NAME, '写入失败：%s' % ex)
            return
        msg = '新增 %d 条，已存在 %d 条，补齐品类 %d 条，库内共 %d 条。' % (
            r['added'], r['dup'], r['filled'], r['total'])
        if r['conflicts']:
            msg += '\n\n⚠ 冲突 %d 条（同一货号+花色+尺码有两个编码），未写入：\n%s' % (
                len(r['conflicts']), '\n'.join('  ' + x for x in r['conflicts'][:6]))
        self.lib_info.config(text=msg, fg='#006100' if not r['conflicts'] else '#C00000')
        self.check_doc()
        if r['added'] or r['filled']:
            self.paste_txt.delete('1.0', 'end')
            self._parsed = None

    # ============ Tab2 生成 ============
    def _build_tab2(self, F):
        t = self.t2
        tk.Label(t, text='一行一个组合，格式：货号-花色+花色 / N条装 / 尺码1,尺码2',
                 font=('微软雅黑', 9), fg='#555').pack(anchor='w', padx=12, pady=(10, 2))
        tk.Label(t, text='3901-藤紫+星轨 / 2条装 / S,M,L,XL,2XL　　2620-深灰条纹 / 1件装 / M,L\n'
                         'MX088+MX089 / 保暖两件套 / 110,120,130,140,150,160　　'
                         '79005+79002 / 4条装 / M,L,XL',
                 font=('微软雅黑', 8), fg='#888', justify='left').pack(anchor='w', padx=12)

        wrap = tk.Frame(t)
        wrap.pack(fill='both', expand=True, padx=12, pady=(6, 0))
        self.combo_txt = tk.Text(wrap, font=('Consolas', 10), undo=True, relief='solid', bd=1)
        sb = tk.Scrollbar(wrap, command=self.combo_txt.yview)
        self.combo_txt.configure(yscrollcommand=sb.set)
        self.combo_txt.pack(side='left', fill='both', expand=True)
        sb.pack(side='right', fill='y')
        self.combo_txt.insert('1.0', self.cfg.get('combos') or
                              '3901-藤紫+星轨 / 2条装 / S,M,L,XL,2XL')

        f = tk.Frame(t)
        f.pack(fill='x', padx=12, pady=(8, 0))
        tk.Label(f, text='组合商品名称', font=F).grid(row=0, column=0, sticky='w')
        self.name_var = tk.StringVar(value=self.cfg.get('name', '儿童内裤'))
        tk.Entry(f, textvariable=self.name_var, font=F, width=14).grid(row=0, column=1, padx=(6, 16))
        tk.Label(f, text='标签', font=F).grid(row=0, column=2, sticky='w')
        self.tag_var = tk.StringVar(value=self.cfg.get('tag', '绿色标'))
        tk.Entry(f, textvariable=self.tag_var, font=F, width=8).grid(row=0, column=3, padx=(6, 16))
        tk.Label(f, text='B列逗号', font=F).grid(row=0, column=4, sticky='w')
        self.sep_var = tk.StringVar(value=self.cfg.get('sep', '中文，'))
        ttk.Combobox(f, textvariable=self.sep_var, font=F, width=7, state='readonly',
                     values=['中文，', '英文,']).grid(row=0, column=5, padx=(6, 16))
        tk.Label(f, text='组间空行', font=F).grid(row=0, column=6, sticky='w')
        self.gap_var = tk.StringVar(value=str(self.cfg.get('gap', C.GAP_ROWS)))
        tk.Spinbox(f, from_=0, to=30, textvariable=self.gap_var, font=F,
                   width=4).grid(row=0, column=7, padx=6)

        f2 = tk.Frame(t)
        f2.pack(fill='x', padx=12, pady=(12, 10))
        tk.Button(f2, text='生成预览', font=('微软雅黑', 11, 'bold'), bg='#1F3864',
                  fg='white', relief='flat', padx=20, pady=6,
                  command=self.do_build).pack(side='left')
        tk.Button(f2, text='清空组合', font=F, relief='flat', padx=12, pady=6,
                  command=lambda: self.combo_txt.delete('1.0', 'end')).pack(side='left', padx=8)

    def do_build(self):
        p = self.doc_var.get()
        if not p or not os.path.exists(p):
            messagebox.showwarning(C.APP_NAME, '请先选择编码文档。')
            return
        try:
            codebook, _, _, issues, _ = C.load_doc(p)
        except Exception as ex:
            messagebox.showerror(C.APP_NAME, '读编码文档失败：%s' % ex)
            return
        if not codebook:
            messagebox.showerror(C.APP_NAME, '编码文档里没读出任何编码。')
            return
        if issues:
            messagebox.showerror(C.APP_NAME, '编码库有 %d 处问题，先修好再生成：\n\n%s'
                                 % (len(issues), '\n'.join('  ' + x for x in issues[:8])))
            return
        try:
            gap = int(self.gap_var.get())
        except Exception:
            gap = C.GAP_ROWS
        sep = '，' if self.sep_var.get().startswith('中文') else ','
        try:
            res = C.build_rows(codebook, self.combo_txt.get('1.0', 'end'),
                               self.name_var.get().strip() or '儿童内裤',
                               self.tag_var.get().strip() or '绿色标', gap, sep)
        except Exception as ex:
            messagebox.showerror(C.APP_NAME, str(ex))
            return
        self.rows = res['rows']
        self.build_warns = res['warns']
        self.build_miss = res['miss']
        self.build_check = res['check']
        C.save_cfg({'doc': p, 'combos': self.combo_txt.get('1.0', 'end'),
                    'name': self.name_var.get(), 'tag': self.tag_var.get(),
                    'gap': gap, 'sep': self.sep_var.get(),
                    'out': self.cfg.get('out', '')})
        self.refresh_table()
        self.nb.select(self.t3)
        self.status.config(text='预览已生成，可直接在表格里改，确认后导出。', fg='#006100')

    # ============ Tab3 预览 ============
    def _build_tab3(self, F):
        t = self.t3
        top = tk.Frame(t)
        top.pack(fill='x', padx=12, pady=(8, 0))
        self.stat = tk.Label(top, text='还没有预览。先在「② 生成」里点生成预览。',
                             font=('微软雅黑', 9), fg='#555', anchor='w', justify='left')
        self.stat.pack(side='left')

        cols = ('a', 'b', 'd', 'j', 'k', 'q', 'l')
        heads = ('组合商品编码', '组合装商品标签', '组合商品名称', '商品编码',
                 '数量', '条数Q', '应占售价')
        widths = (300, 150, 90, 190, 50, 60, 70)
        wrap = tk.Frame(t)
        wrap.pack(fill='both', expand=True, padx=12, pady=(6, 0))
        self.tree = ttk.Treeview(wrap, columns=cols, show='headings', height=18)
        for c, h, w in zip(cols, heads, widths):
            self.tree.heading(c, text=h)
            self.tree.column(c, width=w, anchor='w')
        vs = ttk.Scrollbar(wrap, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=vs.set)
        self.tree.pack(side='left', fill='both', expand=True)
        vs.pack(side='right', fill='y')
        self.tree.tag_configure('miss', background='#FFC7CE')
        self.tree.tag_configure('check', background='#FFF2CC')
        self.tree.tag_configure('gap', background='#F2F2F2')
        self.tree.bind('<Double-1>', self.on_edit)

        f = tk.Frame(t)
        f.pack(fill='x', padx=12, pady=(8, 0))
        tk.Label(f, text='输出到', font=F).pack(side='left')
        self.out_var = tk.StringVar(value=self.cfg.get('out') or
                                    os.path.join(C.base_dir(), '输出'))
        tk.Entry(f, textvariable=self.out_var, font=F, state='readonly').pack(
            side='left', fill='x', expand=True, padx=8)
        tk.Button(f, text='更改…', font=F, command=self.pick_out).pack(side='left')
        tk.Button(f, text='默认', font=F, command=self.reset_out).pack(side='left', padx=(6, 0))

        f2 = tk.Frame(t)
        f2.pack(fill='x', padx=12, pady=(8, 4))
        tk.Label(f2, text='文件名', font=F).pack(side='left')
        self.fname_var = tk.StringVar(value=self.cfg.get('fname', ''))
        tk.Entry(f2, textvariable=self.fname_var, font=F).pack(
            side='left', fill='x', expand=True, padx=8)
        tk.Label(f2, text='留空=自动命名', font=('微软雅黑', 8), fg='#888').pack(side='left')

        f3 = tk.Frame(t)
        f3.pack(fill='x', padx=12, pady=(6, 10))
        tk.Button(f3, text='导出 Excel', font=('微软雅黑', 11, 'bold'), bg='#1F3864',
                  fg='white', relief='flat', padx=20, pady=6,
                  command=self.do_export).pack(side='left')
        tk.Button(f3, text='删除选中行', font=F, relief='flat', padx=12, pady=6,
                  command=self.del_row).pack(side='left', padx=8)
        tk.Button(f3, text='打开所在文件夹', font=F, relief='flat', padx=12, pady=6,
                  command=self.open_out).pack(side='left')
        tk.Label(f3, text='（双击任意格可修改）', font=('微软雅黑', 8),
                 fg='#888').pack(side='left', padx=10)

    def refresh_table(self):
        self.tree.delete(*self.tree.get_children())
        self._iid = {}
        idx = 0
        for n, rw in enumerate(self.rows):
            if rw['blank']:
                iid = self.tree.insert('', 'end', values=('', '', '', '', '', '', ''),
                                       tags=('gap',))
                self._iid[iid] = n
                continue
            idx += 1
            tags = ()
            if rw['miss']:
                tags = ('miss',)
            elif rw['risk']:
                tags = ('check',)
            iid = self.tree.insert('', 'end', tags=tags, values=(
                rw['a'], rw['b'], rw['d'], rw['j'], rw['k'],
                rw['q'] if isinstance(rw['q'], int) else '公式',
                self._lval(rw)))
            self._iid[iid] = n
        n_groups = len(set(r.get('gid') for r in self.rows if not r['blank']))
        msg = '共 %d 组、%d 行' % (n_groups, idx)
        if self.build_miss:
            msg += '　⚠ %d 格没查到编码（红）' % len(self.build_miss)
        if self.build_check:
            msg += '　⚠ %d 行 Q 需手动核查（黄）' % len(self.build_check)
        if self.build_warns:
            msg += '　· %d 条提示' % len(self.build_warns)
        self.stat.config(text=msg, fg='#C00000' if (self.build_miss or self.build_check) else '#006100')
        if self.build_warns:
            self.status.config(text='提示：\n' + '\n'.join('  ' + w for w in self.build_warns[:5]),
                               fg='#886600')

    def _lval(self, rw):
        s = [r for r in self.rows if not r['blank'] and r['a'] == rw['a']]
        tot = 0
        for r in s:
            q = r['q']
            if isinstance(q, int):
                tot += q
            else:
                m = re.search(r'T(\d)S', r['j'] or '')
                tot += int(m.group(1)) if m else 0
        q = rw['q']
        if not isinstance(q, int):
            m = re.search(r'T(\d)S', rw['j'] or '')
            q = int(m.group(1)) if m else 0
        if not tot:
            return ''
        return round(q / tot, 4)

    def on_edit(self, ev):
        iid = self.tree.identify_row(ev.y)
        col = self.tree.identify_column(ev.x)
        if not iid or iid not in self._iid:
            return
        n = self._iid[iid]
        rw = self.rows[n]
        if rw['blank']:
            return
        keys = {'#2': 'b', '#3': 'd', '#4': 'j', '#6': 'q'}
        key = keys.get(col)
        if not key:
            return
        top = tk.Toplevel(self.root)
        top.title('修改')
        top.geometry('420x120')
        top.transient(self.root)
        tk.Label(top, text={'b': '组合装商品标签', 'd': '组合商品名称',
                            'j': '商品编码', 'q': '条数 Q'}[key],
                 font=('微软雅黑', 9)).pack(anchor='w', padx=10, pady=(10, 2))
        var = tk.StringVar(value=str(rw[key]))
        e = tk.Entry(top, textvariable=var, font=('Consolas', 10))
        e.pack(fill='x', padx=10)
        e.focus_set()

        def ok(_=None):
            v = var.get().strip()
            if key == 'q':
                try:
                    v = int(v)
                except Exception:
                    messagebox.showwarning(C.APP_NAME, '条数要填数字', parent=top)
                    return
            rw[key] = v
            if key == 'j':
                rw['miss'] = not bool(v)
                fam = C.parse_code(v)[2] if v else None
                rw['kind'] = fam
                rw['risk'] = C.q_risk(v) if fam == 'inner' else None
                if fam in ('height', 'vest'):
                    rw['q'] = 1
            top.destroy()
            self.refresh_table()
        e.bind('<Return>', ok)
        tk.Button(top, text='确定', font=('微软雅黑', 9), command=ok).pack(pady=8)

    def del_row(self):
        sel = [i for i in self.tree.selection() if i in self._iid]
        if not sel:
            messagebox.showinfo(C.APP_NAME, '先在表格里选一行。')
            return
        idxs = sorted((self._iid[i] for i in sel), reverse=True)
        for n in idxs:
            del self.rows[n]
        self.refresh_table()

    # ---------- 输出相关 ----------
    def pick_out(self):
        d = filedialog.askdirectory(title='选择生成文件的保存位置',
                                    initialdir=self.out_var.get() or C.base_dir())
        if d:
            self.out_var.set(os.path.normpath(d))

    def reset_out(self):
        self.out_var.set(os.path.join(C.base_dir(), '输出'))

    def open_out(self):
        d = self.out_var.get() or os.path.join(C.base_dir(), '输出')
        try:
            os.makedirs(d, exist_ok=True)
            os.startfile(d)
        except Exception as ex:
            messagebox.showwarning(C.APP_NAME, '打不开：%s' % ex)

    def do_export(self):
        if not self.rows:
            messagebox.showwarning(C.APP_NAME, '还没有预览，先在「② 生成」里生成。')
            return
        stamp = datetime.datetime.now().strftime('%m%d-%H%M')
        n_groups = len(set(r.get('gid') for r in self.rows if not r['blank']))
        name = (self.fname_var.get() or '').strip()
        if name:
            for ch in '\\/:*?"<>|':
                name = name.replace(ch, '_')
            if not name.lower().endswith('.xlsx'):
                name += '.xlsx'
            fname = name
        else:
            fname = '建档-%d组-%s.xlsx' % (n_groups, stamp)

        dirs = []
        for d in [self.out_var.get(), os.path.join(C.base_dir(), '输出'),
                  os.path.join(os.path.expanduser('~'), 'Documents')]:
            if d and d not in dirs:
                dirs.append(d)
        errs, out = [], None
        for d in dirs:
            try:
                os.makedirs(d, exist_ok=True)
                stem, ext = os.path.splitext(fname)
                p = os.path.join(d, fname)
                i = 1
                while os.path.exists(p):
                    i += 1
                    p = os.path.join(d, '%s-%d%s' % (stem, i, ext))
                C.export_xlsx(self.rows, p)
                out = p
                break
            except Exception as ex:
                errs.append('%s（%s）' % (d, ex))
        if not out:
            messagebox.showerror(C.APP_NAME, '所有目录都存不进去：\n\n' + '\n'.join(errs))
            return
        self.last_out = out
        C.save_cfg({'doc': self.doc_var.get(), 'combos': self.combo_txt.get('1.0', 'end'),
                    'name': self.name_var.get(), 'tag': self.tag_var.get(),
                    'gap': self.gap_var.get(), 'sep': self.sep_var.get(),
                    'out': self.out_var.get(), 'fname': self.fname_var.get()})
        msg = '已导出：\n%s' % out
        if errs:
            msg += '\n\n⚠ 你指定的文件夹存不进去，已改存到上面这个位置。'
        self.status.config(text=msg, fg='#006100')
        if os.environ.get('SKU_NOOPEN') != '1':
            try:
                os.startfile(out)
            except Exception:
                pass

    # ---------- 文档 / 快捷方式 ----------
    def pick_doc(self):
        p = filedialog.askopenfilename(title='选择款式编码文档',
                                       filetypes=[('Excel 文件', '*.xlsx *.xlsm'), ('所有文件', '*.*')])
        if p:
            self.doc_var.set(p)
            self.check_doc()

    def check_doc(self):
        p = self.doc_var.get()
        if not p or not os.path.exists(p):
            self.doc_info.config(text='（未选择编码文档）', fg='#C00000')
            return
        try:
            book, groups, nonstd, issues, mode = C.load_doc(p)
            nos = sorted(set(g[0] for g in groups))
            self.doc_info.config(
                text='[%s] %d 款 × %d 个编码　%s%s' % (
                    mode, len(groups), len(book), '、'.join(nos[:12]) + ('…' if len(nos) > 12 else ''),
                    ('　⚠ %d 处不一致' % len(issues)) if issues else ''),
                fg='#C00000' if issues else '#006100')
        except Exception as ex:
            self.doc_info.config(text='读取失败：%s' % ex, fg='#C00000')

    def make_shortcut(self):
        exe = sys.executable
        desk = C.desktop_dir()
        lnk = os.path.join(desk, 'SKU建档工作台.lnk')
        ps = ('$s=(New-Object -ComObject WScript.Shell).CreateShortcut("%s");'
              '$s.TargetPath="%s";$s.WorkingDirectory="%s";'
              '$s.IconLocation="%s,0";$s.Save()' % (lnk, exe, C.base_dir(), exe))
        try:
            subprocess.run(['powershell', '-NoProfile', '-Command', ps],
                           capture_output=True, timeout=20, creationflags=0x08000000)
            if os.path.exists(lnk):
                self.status.config(text='桌面快捷方式已创建：%s' % lnk, fg='#006100')
            else:
                self.status.config(text='快捷方式没创建成功（可能桌面目录不可写）。\n'
                                        '可以手工：右键 exe → 发送到 → 桌面快捷方式。', fg='#C00000')
        except Exception as ex:
            self.status.config(text='创建快捷方式失败：%s' % ex, fg='#C00000')


def ensure_admin():
    if not getattr(sys, 'frozen', False):
        return
    try:
        import ctypes
        if ctypes.windll.shell32.IsUserAnAdmin():
            return
        args = ' '.join('"%s"' % a for a in sys.argv[1:])
        rc = ctypes.windll.shell32.ShellExecuteW(None, 'runas', sys.executable, args, None, 1)
        if rc > 32:
            sys.exit(0)
    except Exception:
        pass


def main():
    if '--selftest' in sys.argv:
        print('frozen:', getattr(sys, 'frozen', False))
        print('exe dir:', C.base_dir())
        r = tk.Tk()
        print('tk ok', tk.TkVersion)
        r.destroy()
        print('core ok, sizes', len(C.SZ_LETTER))
        print('selftest PASS')
        return
    ensure_admin()
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == '__main__':
    try:
        main()
    except Exception:
        try:
            from tkinter import messagebox as mb
            mb.showerror(C.APP_NAME, traceback.format_exc())
        except Exception:
            pass
