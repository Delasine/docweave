# -*- coding: utf-8 -*-
"""docweave.gui —— 图形界面（打包成 exe 的入口）。

界面把五种能力都做成独立页签：
  合并 / 文表对账 / 引用体检 / 口径漂移 / 批量处理 + 工具
所有耗时操作都在后台线程里跑，结果回填到表格与日志，不阻塞界面。
"""
from __future__ import annotations

import json
import os
import queue
import subprocess
import sys
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from . import ops, writers as W
from .model import read_document

APP_TITLE = "docweave · 文档合并与一致性审计"
APP_VERSION = "1.0.0"

DOC_TYPES = [("Word 文档", "*.docx"), ("全部文件", "*.*")]
DATA_TYPES = [("Excel / CSV", "*.xlsx *.xlsm *.csv *.tsv"), ("全部文件", "*.*")]
ANY_TYPES = [("支持的文件", "*.docx *.xlsx *.xlsm *.csv *.tsv *.md *.txt"),
             ("全部文件", "*.*")]
JSON_TYPES = [("JSON 清单", "*.json"), ("全部文件", "*.*")]
MD_TYPES = [("Markdown", "*.md"), ("全部文件", "*.*")]
DOCX_TYPES = [("Word 文档", "*.docx"), ("全部文件", "*.*")]
XLSX_TYPES = [("Excel 工作簿", "*.xlsx"), ("全部文件", "*.*")]


def resource(rel: str) -> str:
    """定位随包/随 exe 附带的资源。"""
    base = getattr(sys, "_MEIPASS", None)
    if base:
        p = os.path.join(base, rel)
        if os.path.exists(p):
            return p
    here = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(here, rel)


def open_path(path: str):
    if not path or not os.path.exists(path):
        messagebox.showwarning(APP_TITLE, f"文件不存在：\n{path}")
        return
    try:
        os.startfile(path)                                   # noqa: S606
    except Exception:
        subprocess.Popen(["cmd", "/c", "start", "", path], shell=False)


def open_folder(path: str):
    d = path if os.path.isdir(path) else os.path.dirname(path)
    if not d or not os.path.isdir(d):
        messagebox.showwarning(APP_TITLE, f"目录不存在：\n{d}")
        return
    try:
        os.startfile(d)                                      # noqa: S606
    except Exception:
        subprocess.Popen(["explorer", d])


# --------------------------------------------------------------------------
# 组件
# --------------------------------------------------------------------------

class FileRow(ttk.Frame):
    """一行「标签 + 输入框 + 浏览…」。mode: open / save / dir"""

    def __init__(self, master, label, mode="open", types=None, default_ext="",
                 width=62, initial="", label_width=12):
        super().__init__(master)
        self.mode, self.types, self.default_ext = mode, types or ANY_TYPES, default_ext
        self.label_text = label
        self.var = tk.StringVar(value=initial)
        ttk.Label(self, text=label, width=label_width, anchor="e").pack(side="left")
        ttk.Entry(self, textvariable=self.var, width=width).pack(
            side="left", fill="x", expand=True, padx=4)
        ttk.Button(self, text="浏览…", width=8, command=self._browse).pack(side="left")

    def _browse(self):
        if self.mode == "open":
            p = filedialog.askopenfilename(title=self.label_text, filetypes=self.types)
        elif self.mode == "dir":
            p = filedialog.askdirectory(title=self.label_text)
        else:
            p = filedialog.asksaveasfilename(title=self.label_text,
                                             filetypes=self.types,
                                             defaultextension=self.default_ext)
        if p:
            if (self.mode == "save" and self.default_ext
                    and not p.lower().endswith(self.default_ext)):
                p += self.default_ext
            self.var.set(os.path.normpath(p))

    def get(self) -> str:
        return self.var.get().strip()

    def set(self, v):
        self.var.set(v or "")

    def suggest(self, path: str, suffix: str):
        """按输入文件自动推导产物路径（仅当输出框为空时）。"""
        if self.get() or not path:
            return
        self.set(os.path.splitext(path)[0] + suffix)


def tip(parent, text):
    lbl = ttk.Label(parent, text=text, foreground="#646a73", wraplength=980,
                    justify="left")
    return lbl


# --------------------------------------------------------------------------
# 主窗口
# --------------------------------------------------------------------------

class App(ttk.Frame):
    def __init__(self, master: tk.Tk):
        super().__init__(master, padding=10)
        self.master = master
        self.q: queue.Queue = queue.Queue()
        self.busy = False
        self.last_outputs: list = []
        self.last_html: str = ""
        self.pack(fill="both", expand=True)
        self._build_header()
        self._build_tabs()
        self._build_results()
        self._build_status()
        self.master.after(120, self._poll)
        self.log("欢迎使用 docweave。第一次使用建议：工具页 →「生成演示数据」，"
                 "然后各页签点「载入演示数据」即可一键体验。")

    # -- 头部 ---------------------------------------------------------------
    def _build_header(self):
        bar = ttk.Frame(self)
        bar.pack(fill="x", pady=(0, 8))
        ttk.Label(bar, text="docweave", font=("Microsoft YaHei UI", 16, "bold")
                  ).pack(side="left")
        ttk.Label(bar, text="  文档语义合并 · 文表对账 · 引用体检 · 口径漂移 · 批量台账",
                  foreground="#646a73").pack(side="left")
        ttk.Button(bar, text="打开输出目录", command=self._open_last_dir).pack(side="right")
        ttk.Button(bar, text="帮助", command=lambda: self.nb.select(5)).pack(
            side="right", padx=6)

    # -- 页签 ---------------------------------------------------------------
    def _build_tabs(self):
        self.split = ttk.Panedwindow(self, orient="vertical")
        self.split.pack(fill="both", expand=True)
        self.nb = ttk.Notebook(self.split)
        self.split.add(self.nb, weight=3)
        self._tab_merge()
        self._tab_audit()
        self._tab_refs()
        self._tab_drift()
        self._tab_batch()
        self._tab_tools()

    def _page(self, title):
        f = ttk.Frame(self.nb, padding=12)
        self.nb.add(f, text=title)
        return f

    # 合并
    def _tab_merge(self):
        f = self._page("① 合并")
        tip(f, "把两个/三个文档合并成一份。给了「基线」就是三方合并（只有双方改到同一处才算冲突）；"
               "不给基线就是两个不同文档的对比整合。冲突默认在正文里同时保留双方并编号。"
               ).pack(anchor="w", pady=(0, 8))
        self.m_base = FileRow(f, "基线文档", "open", DOC_TYPES)
        self.m_ours = FileRow(f, "我方文档", "open", DOC_TYPES)
        self.m_theirs = FileRow(f, "对方文档", "open", DOC_TYPES)
        self.m_out = FileRow(f, "输出文件", "save", DOCX_TYPES, ".docx")
        self.m_rep = FileRow(f, "冲突报告", "save", MD_TYPES, ".md")
        self.m_json = FileRow(f, "机读结果", "save", JSON_TYPES, ".json")
        for r in (self.m_base, self.m_ours, self.m_theirs, self.m_out,
                  self.m_rep, self.m_json):
            r.pack(fill="x", pady=2)
        opt = ttk.LabelFrame(f, text="选项", padding=8)
        opt.pack(fill="x", pady=8)
        ttk.Label(opt, text="冲突处置").grid(row=0, column=0, sticky="e", padx=4)
        self.m_conf = ttk.Combobox(opt, width=18, state="readonly",
                                   values=["保留双方并标记", "全部按我方定稿", "全部按对方定稿"])
        self.m_conf.current(0)
        self.m_conf.grid(row=0, column=1, sticky="w")
        ttk.Label(opt, text="表格主键列").grid(row=0, column=2, sticky="e", padx=(18, 4))
        self.m_key = ttk.Entry(opt, width=8)
        self.m_key.insert(0, "0")
        self.m_key.grid(row=0, column=3, sticky="w")
        ttk.Label(opt, text="（第几列作行主键，none=不做单元格级合并）",
                  foreground="#646a73").grid(row=0, column=4, sticky="w", padx=6)
        ttk.Label(opt, text="版式模板").grid(row=1, column=0, sticky="e", padx=4, pady=6)
        self.m_tpl = ttk.Combobox(opt, width=18, state="readonly",
                                  values=["自动（优先我方）", "基线", "我方", "对方"])
        self.m_tpl.current(0)
        self.m_tpl.grid(row=1, column=1, sticky="w", pady=6)
        self.m_appendix = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt, text="在产物末尾附加「溯源附录」",
                        variable=self.m_appendix).grid(row=1, column=2, columnspan=2,
                                                       sticky="w", padx=(18, 4))
        btns = ttk.Frame(f)
        btns.pack(fill="x", pady=6)
        ttk.Button(btns, text="载入演示数据", command=self._demo_merge).pack(side="left")
        ttk.Button(btns, text="▶ 开始合并", command=self._run_merge).pack(
            side="left", padx=8)
        self.m_auto = tk.BooleanVar(value=True)
        ttk.Checkbutton(btns, text="自动填充产物路径", variable=self.m_auto).pack(side="left")
        ttk.Button(btns, text="刷新建议路径",
                   command=lambda: self._suggest_merge(force=True)).pack(side="left", padx=8)

    # 对账
    def _tab_audit(self):
        f = self._page("② 文表对账")
        tip(f, "把正文里的数字、日期、编号与 Excel/CSV 里的数据逐项核对，"
               "自动做单位换算（元/万元、%/‰、天/个月）与标签模糊匹配。"
               ).pack(anchor="w", pady=(0, 8))
        self.a_doc = FileRow(f, "正文文档", "open", ANY_TYPES)
        self.a_data = FileRow(f, "数据文件", "open", DATA_TYPES)
        self.a_out = FileRow(f, "对账报告", "save", DOCX_TYPES, ".docx")
        self.a_rep = FileRow(f, "Markdown", "save", MD_TYPES, ".md")
        self.a_json = FileRow(f, "机读结果", "save", JSON_TYPES, ".json")
        for r in (self.a_doc, self.a_data, self.a_out, self.a_rep, self.a_json):
            r.pack(fill="x", pady=2)
        opt = ttk.LabelFrame(f, text="选项", padding=8)
        opt.pack(fill="x", pady=8)
        ttk.Label(opt, text="数值容差").grid(row=0, column=0, sticky="e", padx=4)
        self.a_tol = ttk.Entry(opt, width=10)
        self.a_tol.insert(0, "0.005")
        self.a_tol.grid(row=0, column=1, sticky="w")
        ttk.Label(opt, text="（0.005 = 相差 0.5% 以内算一致）", foreground="#646a73"
                  ).grid(row=0, column=2, sticky="w", padx=6)
        self.a_recalc = tk.BooleanVar(value=False)
        self.a_showok = tk.BooleanVar(value=False)
        self.a_embed = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt, text="先用内置 LibreOffice 重算公式", variable=self.a_recalc
                        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=4)
        ttk.Checkbutton(opt, text="同时列出核对一致的项目", variable=self.a_showok
                        ).grid(row=1, column=2, sticky="w", padx=6)
        ttk.Checkbutton(opt, text="报告里附上数据表原文", variable=self.a_embed
                        ).grid(row=2, column=0, columnspan=2, sticky="w")
        btns = ttk.Frame(f)
        btns.pack(fill="x", pady=6)
        ttk.Button(btns, text="载入演示数据", command=self._demo_audit).pack(side="left")
        ttk.Button(btns, text="▶ 开始对账", command=self._run_audit).pack(side="left", padx=8)

    # 体检
    def _tab_refs(self):
        f = self._page("③ 引用体检")
        tip(f, "检查交叉引用与编号：悬空引用（见 3.2 节 / 图 5）、编号重复、编号断档、模糊引用。"
               ).pack(anchor="w", pady=(0, 8))
        self.r_doc = FileRow(f, "待检文档", "open", DOC_TYPES)
        self.r_out = FileRow(f, "体检报告", "save", DOCX_TYPES, ".docx")
        self.r_rep = FileRow(f, "Markdown", "save", MD_TYPES, ".md")
        self.r_json = FileRow(f, "机读结果", "save", JSON_TYPES, ".json")
        for r in (self.r_doc, self.r_out, self.r_rep, self.r_json):
            r.pack(fill="x", pady=2)
        btns = ttk.Frame(f)
        btns.pack(fill="x", pady=10)
        ttk.Button(btns, text="载入演示数据", command=self._demo_refs).pack(side="left")
        ttk.Button(btns, text="▶ 开始体检", command=self._run_refs).pack(side="left", padx=8)

    # 漂移
    def _tab_drift(self):
        f = self._page("④ 口径漂移")
        tip(f, "按顺序放入同一份文档的多个版本，找出指标在各版本间的变化，"
               "并检测同一份文档内部自相矛盾（同一个指标写了两个值）。只放一个文件就只做矛盾检查。"
               ).pack(anchor="w", pady=(0, 8))
        body = ttk.Frame(f)
        body.pack(fill="both", expand=True)
        left = ttk.Frame(body)
        left.pack(side="left", fill="both", expand=True)
        self.d_list = tk.Listbox(left, height=8, selectmode="extended",
                                 font=("Microsoft YaHei UI", 9))
        sb = ttk.Scrollbar(left, orient="vertical", command=self.d_list.yview)
        self.d_list.configure(yscrollcommand=sb.set)
        self.d_list.pack(side="left", fill="both", expand=True)
        sb.pack(side="left", fill="y")
        side = ttk.Frame(body)
        side.pack(side="left", fill="y", padx=8)
        ttk.Button(side, text="添加文件…", command=self._drift_add_files, width=12
                   ).pack(pady=2)
        ttk.Button(side, text="添加目录…", command=self._drift_add_dir, width=12
                   ).pack(pady=2)
        ttk.Button(side, text="移除选中", command=self._drift_remove, width=12
                   ).pack(pady=2)
        ttk.Button(side, text="清空", command=lambda: self.d_list.delete(0, "end"),
                   width=12).pack(pady=2)
        self.d_out = FileRow(f, "漂移报告", "save", DOCX_TYPES, ".docx")
        self.d_rep = FileRow(f, "Markdown", "save", MD_TYPES, ".md")
        self.d_json = FileRow(f, "机读结果", "save", JSON_TYPES, ".json")
        for r in (self.d_out, self.d_rep, self.d_json):
            r.pack(fill="x", pady=2)
        btns = ttk.Frame(f)
        btns.pack(fill="x", pady=6)
        ttk.Button(btns, text="载入演示数据", command=self._demo_drift).pack(side="left")
        ttk.Button(btns, text="▶ 开始分析", command=self._run_drift).pack(side="left", padx=8)

    # 批量
    def _tab_batch(self):
        f = self._page("⑤ 批量处理")
        tip(f, "先选一个目录 →「扫描并生成清单」（会自动发现 *_base/_ours/_theirs 组合、"
               "正文+数据表配对、版本序列），再点「开始批处理」。没变化的文件会按台账自动跳过；"
               "也可直接「扫描并增量处理」，一次完成变化记录、任务识别和处理。"
               ).pack(anchor="w", pady=(0, 8))
        self.b_dir = FileRow(f, "待处理目录", "dir")
        self.b_man = FileRow(f, "批量清单", "save", JSON_TYPES, ".json")
        self.b_out = FileRow(f, "输出目录", "dir")
        for r in (self.b_dir, self.b_man, self.b_out):
            r.pack(fill="x", pady=2)
        opt = ttk.LabelFrame(f, text="选项", padding=8)
        opt.pack(fill="x", pady=8)
        self.b_force = tk.BooleanVar(value=False)
        ttk.Checkbutton(opt, text="忽略台账，全部重跑", variable=self.b_force
                        ).grid(row=0, column=0, sticky="w")
        ttk.Label(opt, text="（默认按内容指纹跳过没有变化的文件）", foreground="#646a73"
                  ).grid(row=0, column=1, sticky="w", padx=8)
        btns = ttk.Frame(f)
        btns.pack(fill="x", pady=6)
        ttk.Button(btns, text="扫描并增量处理", command=self._scan_and_process).pack(side="left")
        ttk.Button(btns, text="生成清单预览", command=self._plan).pack(side="left", padx=6)
        ttk.Button(btns, text="按清单处理", command=self._run_batch).pack(side="left", padx=6)
        ttk.Button(btns, text="打开 HTML 报告", command=self._open_html).pack(side="left")
        ttk.Button(btns, text="载入演示数据", command=self._demo_batch).pack(
            side="left", padx=8)

    # 工具
    def _tab_tools(self):
        f = self._page("⑥ 工具 / 帮助")
        demo = ttk.LabelFrame(f, text="演示数据", padding=8)
        demo.pack(fill="x")
        self.t_demo = FileRow(demo, "生成到目录", "dir")
        self.t_demo.pack(fill="x")
        row = ttk.Frame(demo)
        row.pack(fill="x", pady=6)
        ttk.Button(row, text="生成演示数据", command=self._gen_demo).pack(side="left")
        ttk.Label(row, text="会生成三方合同稿、文表对账、引用体检、版本漂移、表格合并五组样例",
                  foreground="#646a73").pack(side="left", padx=8)

        led = ttk.LabelFrame(f, text="增量台账", padding=8)
        led.pack(fill="x", pady=8)
        self.t_led = FileRow(led, "台账根目录", "dir")
        self.t_led.pack(fill="x")
        row2 = ttk.Frame(led)
        row2.pack(fill="x", pady=6)
        ttk.Button(row2, text="查看台账", command=self._show_ledger).pack(side="left")
        ttk.Button(row2, text="清除台账", command=self._reset_ledger).pack(side="left", padx=8)

        rc = ttk.LabelFrame(f, text="Excel 公式重算", padding=8)
        rc.pack(fill="x")
        self.t_xlsx = FileRow(rc, "工作簿", "open", XLSX_TYPES)
        self.t_xlsx.pack(fill="x")
        row3 = ttk.Frame(rc)
        row3.pack(fill="x", pady=6)
        ttk.Button(row3, text="重算公式并另存", command=self._recalc).pack(side="left")
        ttk.Label(row3, text="Excel 缓存值可能过期，对账前重算更可靠",
                  foreground="#646a73").pack(side="left", padx=8)

        help_box = ttk.LabelFrame(f, text="使用说明与图例", padding=8)
        help_box.pack(fill="both", expand=True, pady=8)
        txt = tk.Text(help_box, height=12, wrap="word", relief="flat",
                      background="#fafbfc")
        txt.pack(fill="both", expand=True)
        txt.insert("1.0", HELP_TEXT)
        txt.configure(state="disabled")

    # -- 结果区 -------------------------------------------------------------
    def _build_results(self):
        box = ttk.LabelFrame(self.split, text="运行结果", padding=6)
        self.split.add(box, weight=2)
        self.summary = ttk.Label(box, text="尚未运行", foreground="#1f2329",
                                 wraplength=1000, justify="left")
        self.summary.pack(anchor="w", pady=(0, 4))
        pane = ttk.Panedwindow(box, orient="vertical")
        pane.pack(fill="both", expand=True)
        top = ttk.Frame(pane)
        self.tree = ttk.Treeview(top, show="headings", height=6)
        vs = ttk.Scrollbar(top, orient="vertical", command=self.tree.yview)
        hs = ttk.Scrollbar(top, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vs.pack(side="right", fill="y")
        hs.pack(side="bottom", fill="x")
        pane.add(top, weight=3)
        low = ttk.Frame(pane)
        ttk.Label(low, text="运行日志", foreground="#646a73").pack(anchor="w")
        self.logbox = tk.Text(low, height=6, wrap="word", relief="flat",
                              background="#f7f8fa", font=("Consolas", 9))
        ls = ttk.Scrollbar(low, orient="vertical", command=self.logbox.yview)
        self.logbox.configure(yscrollcommand=ls.set)
        self.logbox.pack(side="left", fill="both", expand=True)
        ls.pack(side="right", fill="y")
        pane.add(low, weight=1)
        self._set_columns([("提示", 980)])
        self.tree.insert("", "end", values=(
            "运行后这里会列出冲突 / 差异 / 问题明细；选中一行可点下方「打开选中的产物」。",))
        bar = ttk.Frame(box)
        bar.pack(fill="x", pady=(4, 0))
        ttk.Button(bar, text="打开选中的产物", command=self._open_selected).pack(side="left")
        ttk.Button(bar, text="打开产物所在目录",
                   command=lambda: open_folder(self.last_outputs[0])
                   if self.last_outputs else None).pack(side="left", padx=6)
        ttk.Button(bar, text="清空结果", command=self._clear_results).pack(side="left")

    def _build_status(self):
        self.status = ttk.Label(self, text="就绪", foreground="#1f7a3d")
        self.status.pack(anchor="w", pady=(6, 0))

    # -- 基础能力 -----------------------------------------------------------
    def log(self, msg: str):
        self.q.put(("log", msg))

    def set_status(self, text, color="#1f2329"):
        self.status.configure(text=text, foreground=color)

    def _poll(self):
        try:
            while True:
                kind, payload = self.q.get_nowait()
                if kind == "log":
                    self.logbox.insert("end", str(payload) + "\n")
                    self.logbox.see("end")
                elif kind == "done":
                    cb, result = payload
                    self.busy = False
                    self.set_status("完成", "#1f7a3d")
                    try:
                        if cb:
                            cb(result)
                    except Exception:
                        self.log("渲染结果出错：\n" + traceback.format_exc())
                elif kind == "error":
                    self.busy = False
                    exc, tb = payload
                    self.set_status("失败", "#c00000")
                    self.log("运行失败：" + str(exc))
                    self.log(tb)
                    messagebox.showerror(APP_TITLE, f"运行失败：\n{exc}")
        except queue.Empty:
            pass
        self.master.after(120, self._poll)

    def start(self, title: str, fn, on_done=None):
        if self.busy:
            messagebox.showinfo(APP_TITLE, "还有任务在运行，请稍候。")
            return
        self.busy = True
        self.set_status("运行中：" + title, "#b26a00")
        self._set_columns([("信息", 900)])
        self._clear_tree()
        self.log(f"── {title} ──")

        def worker():
            try:
                r = fn()
                self.q.put(("done", (on_done, r)))
            except Exception as exc:                        # noqa: BLE001
                self.q.put(("error", (exc, traceback.format_exc())))

        threading.Thread(target=worker, daemon=True).start()

    def _clear_tree(self):
        for i in self.tree.get_children():
            self.tree.delete(i)

    def _clear_results(self):
        self._clear_tree()
        self.summary.configure(text="尚未运行")
        self.last_outputs = []

    def _set_columns(self, cols):
        self.tree.configure(columns=[c[0] for c in cols])
        for name, width in cols:
            self.tree.heading(name, text=name)
            self.tree.column(name, width=width, anchor="w", stretch=False)

    def fill_tree(self, cols, rows):
        self._set_columns(cols)
        self._clear_tree()
        for r in rows:
            self.tree.insert("", "end", values=r)

    def show_summary(self, stats: dict, extra: str = ""):
        parts = [f"{k}：{v}" for k, v in stats.items()]
        text = "　｜　".join(parts)
        if extra:
            text = extra + ("\n" + text if text else "")
        self.summary.configure(text=text or "无统计信息")

    def remember(self, *paths):
        for p in paths:
            if p and os.path.exists(p):
                self.last_outputs.append(p)
        d = os.path.dirname(self.last_outputs[0]) if self.last_outputs else None
        if d:
            self.set_status(f"完成，产物在 {d}", "#1f7a3d")

    def _open_selected(self):
        sel = self.tree.selection()
        target = None
        if sel:
            vals = self.tree.item(sel[0], "values")
            for v in vals:
                if isinstance(v, str) and os.path.exists(v):
                    target = v
                    break
        if target:
            open_path(target)
        elif self.last_outputs:
            open_path(self.last_outputs[0])
        else:
            messagebox.showinfo(APP_TITLE, "还没有产物。")

    def _open_last_dir(self):
        if self.last_outputs:
            open_folder(self.last_outputs[0])
        else:
            open_path(os.getcwd())

    def _open_html(self):
        if self.last_html and os.path.exists(self.last_html):
            open_path(self.last_html)
        else:
            messagebox.showinfo(APP_TITLE, "还没有生成 HTML 报告，请先跑一次批处理。")

    # -- 演示数据 -----------------------------------------------------------
    def _demo_dir(self) -> str:
        return os.path.join(os.path.expanduser("~"), "docweave_demo")

    def _gen_demo(self):
        out = self.t_demo.get() or self._demo_dir()
        self.t_demo.set(out)

        def job():
            files = samples_files(out)
            self.log(f"演示数据已生成到 {out}（{len(files)} 个文件）")
            return {"out": out, "files": files}

        self.start("生成演示数据", job, self._after_demo)

    def _after_demo(self, r):
        self.show_summary({"生成文件数": len(r["files"])}, f"演示数据目录：{r['out']}")
        self.fill_tree([("文件", 380), ("完整路径", 600)],
                       [(f, os.path.join(r["out"], f)) for f in r["files"]])
        open_folder(r["out"])

    def _demo_merge(self):
        d = self._ensure_demo()
        self.m_base.set(os.path.join(d, "contract_base.docx"))
        self.m_ours.set(os.path.join(d, "contract_ours.docx"))
        self.m_theirs.set(os.path.join(d, "contract_theirs.docx"))
        out = os.path.join(os.path.expanduser("~"), "docweave_out")
        self.m_out.set(os.path.join(out, "contract_merged.docx"))
        self.m_rep.set(os.path.join(out, "contract_conflicts.md"))
        self.m_json.set(os.path.join(out, "contract_conflicts.json"))
        self.log("已载入演示数据（三方合同合并）。")

    def _demo_audit(self):
        d = self._ensure_demo()
        self.a_doc.set(os.path.join(d, "audit_doc.docx"))
        self.a_data.set(os.path.join(d, "audit_data.xlsx"))
        out = os.path.join(os.path.expanduser("~"), "docweave_out")
        self.a_out.set(os.path.join(out, "audit_report.docx"))
        self.a_rep.set(os.path.join(out, "audit_report.md"))
        self.a_json.set(os.path.join(out, "audit_findings.json"))
        self.log("已载入演示数据（文表对账）。")

    def _demo_refs(self):
        d = self._ensure_demo()
        self.r_doc.set(os.path.join(d, "refs_doc.docx"))
        out = os.path.join(os.path.expanduser("~"), "docweave_out")
        self.r_out.set(os.path.join(out, "refs_report.docx"))
        self.r_rep.set(os.path.join(out, "refs_report.md"))
        self.r_json.set(os.path.join(out, "refs_findings.json"))
        self.log("已载入演示数据（引用体检）。")

    def _demo_drift(self):
        d = self._ensure_demo()
        self.d_list.delete(0, "end")
        for i in (1, 2, 3):
            self.d_list.insert("end", os.path.join(d, f"spec_v{i}.docx"))
        out = os.path.join(os.path.expanduser("~"), "docweave_out")
        self.d_out.set(os.path.join(out, "drift_report.docx"))
        self.d_rep.set(os.path.join(out, "drift_report.md"))
        self.d_json.set(os.path.join(out, "drift_result.json"))
        self.log("已载入演示数据（3 个版本口径漂移）。")

    def _demo_batch(self):
        d = self._ensure_demo()
        self.b_dir.set(d)
        out = os.path.join(os.path.expanduser("~"), "docweave_out")
        self.b_man.set(os.path.join(out, "manifest.json"))
        self.b_out.set(out)
        self.log("已载入演示数据（批量处理）。接下来点「① 扫描并生成清单」。")

    def _ensure_demo(self) -> str:
        d = self._demo_dir()
        if not os.path.exists(os.path.join(d, "contract_base.docx")):
            files = samples_files(d)
            self.log(f"已生成演示数据到 {d}（{len(files)} 个文件）")
        return d

    # -- 各功能入口 ---------------------------------------------------------
    def _suggest_merge(self, force=False):
        if not self.m_auto.get() and not force:
            return
        src = self.m_ours.get() or self.m_theirs.get()
        if not src:
            return
        if force:
            self.m_out.set("")
            self.m_rep.set("")
            self.m_json.set("")
        base = os.path.splitext(src)[0]
        self.m_out.suggest(src, "_merged.docx")
        self.m_rep.suggest(src, "_conflicts.md")
        self.m_json.suggest(src, "_conflicts.json")
        if not self.m_out.get():
            self.m_out.set(base + "_merged.docx")

    def _run_merge(self):
        base = self.m_base.get() or None
        ours, theirs = self.m_ours.get(), self.m_theirs.get()
        if not (ours and theirs):
            messagebox.showwarning(APP_TITLE, "至少需要选择「我方文档」和「对方文档」。")
            return
        for p in (base, ours, theirs):
            if p and not os.path.exists(p):
                messagebox.showwarning(APP_TITLE, f"文件不存在：\n{p}")
                return
        self._suggest_merge()
        out = self.m_out.get()
        if not out:
            messagebox.showwarning(APP_TITLE, "请指定输出文件。")
            return
        os.makedirs(os.path.dirname(os.path.abspath(out)) or ".", exist_ok=True)
        conf = {0: "mark", 1: "ours", 2: "theirs"}[self.m_conf.current()]
        key_raw = self.m_key.get().strip().lower()
        key = None if key_raw in ("none", "无", "") else int(key_raw or 0)
        tpl = {0: None, 1: "base", 2: "ours", 3: "theirs"}[self.m_tpl.current()]
        rep = self.m_rep.get() or None
        js = self.m_json.get() or None
        appendix = bool(self.m_appendix.get())

        def job():
            self.log(f"基线：{base or '（无，走两方对比整合）'}")
            self.log(f"我方：{ours}")
            self.log(f"对方：{theirs}")
            return ops.run_merge(base, ours, theirs, out, report=rep,
                                 on_conflict=conf, table_key=key,
                                 with_appendix=appendix, prefer_template=tpl,
                                 sidecars=js)

        self.start("合并文档", job, self._after_merge)

    def _after_merge(self, r):
        from .labels import CONFLICT, zh

        st = dict(r.get("stats") or {})
        self.show_summary(st, f"输出：{r.get('output')}")
        rows = []
        for c in r.get("conflict_list", []):
            rows.append((c.get("id"), zh(CONFLICT, c.get("kind")), c.get("path"),
                         c.get("ours"), c.get("theirs")))
        self.fill_tree([("编号", 70), ("类型", 110), ("位置", 260),
                        ("我方", 320), ("对方", 320)], rows)
        if not rows:
            self.tree.insert("", "end", values=("—", "无冲突", "", "", ""))
        for c in r.get("conflict_list", []):
            self.log(f"冲突 {c.get('id')} [{zh(CONFLICT, c.get('kind'))}] "
                     f"{c.get('path')}")
        self.remember(r.get("output"), r.get("report"))

    def _run_audit(self):
        doc, data = self.a_doc.get(), self.a_data.get()
        if not (doc and data):
            messagebox.showwarning(APP_TITLE, "请选择正文文档和数据文件。")
            return
        self.a_out.suggest(data, "_audit.docx")
        self.a_rep.suggest(data, "_audit.md")
        self.a_json.suggest(data, "_audit.json")
        out = self.a_out.get() or None
        rep = self.a_rep.get() or None
        js = self.a_json.get() or None
        try:
            tol = float(self.a_tol.get())
        except ValueError:
            messagebox.showwarning(APP_TITLE, "容差必须是数字，例如 0.005。")
            return
        if out:
            os.makedirs(os.path.dirname(os.path.abspath(out)) or ".", exist_ok=True)
        recalc, showok, embed = (bool(self.a_recalc.get()), bool(self.a_showok.get()),
                                 bool(self.a_embed.get()))

        def job():
            if recalc:
                self.log("正在用内置 LibreOffice 重算公式 ……")
            return ops.run_audit(doc, data, out=out, report=rep, tol=tol,
                                 embed=embed, recalc=recalc, sidecars=js,
                                 show_ok=showok)

        self.start("文表对账", job, self._after_audit)

    def _after_audit(self, r):
        self.show_summary(r.get("stats") or {}, f"报告：{r.get('output') or '（仅日志）'}")
        rows = [(f.get("severity"), f.get("code"), f.get("label"), f.get("message"))
                for f in r.get("findings", [])]
        self.fill_tree([("严重度", 70), ("类型", 130), ("项目", 160), ("说明", 620)], rows)
        for row in rows:
            self.log(f"[{row[0]}] {row[3]}")
        if not rows:
            self.log("未发现差异。")
        self.remember(r.get("output"), r.get("report"))

    def _run_refs(self):
        doc = self.r_doc.get()
        if not doc:
            messagebox.showwarning(APP_TITLE, "请选择待检文档。")
            return
        self.r_out.suggest(doc, "_refs.docx")
        self.r_rep.suggest(doc, "_refs.md")
        self.r_json.suggest(doc, "_refs.json")
        out, rep, js = self.r_out.get() or None, self.r_rep.get() or None, \
            self.r_json.get() or None
        if out:
            os.makedirs(os.path.dirname(os.path.abspath(out)) or ".", exist_ok=True)
        self.start("引用体检", lambda: ops.run_refs(doc, out=out, report=rep,
                                                    sidecars=js),
                   self._after_refs)

    def _after_refs(self, r):
        self.show_summary(r.get("stats") or {}, f"报告：{r.get('output')}")
        rows = [(f.get("severity"), f.get("code"), f.get("label"), f.get("message"))
                for f in r.get("findings", [])]
        self.fill_tree([("严重度", 70), ("类型", 130), ("定位", 110), ("说明", 670)], rows)
        for row in rows:
            self.log(f"[{row[0]}] {row[3]}")
        self.remember(r.get("output"), r.get("report"))

    def _drift_add_files(self):
        ps = filedialog.askopenfilenames(title="选择同一文档的多个版本", filetypes=ANY_TYPES)
        for p in ps:
            self.d_list.insert("end", os.path.normpath(p))

    def _drift_add_dir(self):
        d = filedialog.askdirectory(title="选择包含多个版本的目录")
        if not d:
            return
        import glob as _g
        for pat in ("*.docx", "*.xlsx", "*.csv", "*.md", "*.txt"):
            for f in sorted(_g.glob(os.path.join(d, pat))):
                if not os.path.basename(f).startswith("~$"):
                    self.d_list.insert("end", os.path.normpath(f))

    def _drift_remove(self):
        for i in reversed(self.d_list.curselection()):
            self.d_list.delete(i)

    def _run_drift(self):
        files = list(self.d_list.get(0, "end"))
        if not files:
            messagebox.showwarning(APP_TITLE, "请至少添加一个文件。")
            return
        self.d_out.suggest(files[0], "_drift.docx")
        self.d_rep.suggest(files[0], "_drift.md")
        self.d_json.suggest(files[0], "_drift.json")
        out, rep, js = self.d_out.get() or None, self.d_rep.get() or None, \
            self.d_json.get() or None
        try:
            tol = float(self.a_tol.get())
        except ValueError:
            tol = 0.005
        if out:
            os.makedirs(os.path.dirname(os.path.abspath(out)) or ".", exist_ok=True)

        def job():
            self.log(f"参与文件 {len(files)} 个")
            return ops.run_drift(files, out=out, report=rep, tol=tol, sidecars=js)

        self.start("口径漂移分析", job, self._after_drift)

    def _after_drift(self, r):
        self.show_summary(r.get("stats") or {}, f"报告：{r.get('output')}")
        rows = [(f.get("severity"), f.get("code"), f.get("label"), f.get("message"))
                for f in r.get("findings", [])]
        self.fill_tree([("严重度", 70), ("类型", 130), ("指标", 160), ("说明", 620)], rows)
        for row in rows:
            self.log(f"[{row[0]}] {row[3]}")
        if not rows:
            self.log("未发现漂移或矛盾。")
        self.remember(r.get("output"), r.get("report"))

    def _plan(self):
        folder = self.b_dir.get()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning(APP_TITLE, "请选择要扫描的目录。")
            return
        out = self.b_out.get() or os.path.join(folder, "out")
        self.b_out.set(out)
        man = self.b_man.get() or os.path.join(out, "manifest.json")
        self.b_man.set(man)

        def job():
            r = ops.build_manifest(folder, outdir=out, out_path=man)
            self.log(f"清单已生成：{r['path']}（{len(r['jobs'])} 个任务）")
            for j in r["jobs"]:
                self.log(f"  · {j.get('type')} {j.get('name', '')}")
            return r

        self.start("扫描目录生成清单", job, self._after_plan)

    def _after_plan(self, r):
        self.show_summary({"发现任务数": len(r["jobs"])}, f"清单：{r['path']}")
        self.fill_tree([("任务名", 220), ("类型", 110), ("输入", 420), ("产物", 340)],
                       [(j.get("name", ""), j.get("type", ""),
                         j.get("doc") or j.get("dir") or j.get("files") or "",
                         j.get("out") or j.get("outdir") or "") for j in r["jobs"]])
        self.remember(r["path"])

    def _run_batch(self):
        man = self.b_man.get()
        if not man or not os.path.exists(man):
            messagebox.showwarning(APP_TITLE, "请先选择或生成批量清单。")
            return
        outdir = self.b_out.get() or None
        force = bool(self.b_force.get())

        def job():
            return ops.run_batch(man, outdir=outdir, force=force,
                                 ledger_root=outdir or os.path.dirname(man),
                                 log=self.log)

        self.start("批量处理", job, self._after_batch)

    def _scan_and_process(self):
        folder = self.b_dir.get()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning(APP_TITLE, "请选择要扫描的目录。")
            return
        outdir = self.b_out.get() or os.path.join(folder, "out")
        self.b_out.set(outdir)
        force = bool(self.b_force.get())

        def job():
            self.log(f"扫描并增量处理目录：{folder}")
            return ops.run_folder_incremental(
                folder, outdir=outdir, force=force, log=self.log)

        self.start("扫描并增量处理", job, self._after_folder_incremental)

    def _after_folder_incremental(self, r):
        self.b_man.set(r["manifest"])
        self._after_batch(r)

    def _after_batch(self, r):
        self.last_html = r.get("html", "")
        self.show_summary(r.get("summary") or {},
                          f"HTML 报告：{self.last_html}")
        rows = []
        for t in r.get("results", []):
            st = t.get("status")
            if st == "fail":
                concl = t.get("error", "")
            elif st == "skip":
                concl = "输入未变化，沿用上次结果"
            else:
                bits = [f"冲突 {t['conflicts']} 处"] if t.get("conflicts") else []
                for k in ("数值冲突", "悬空引用", "重复编号", "编号断档",
                          "文档内自相矛盾", "跨版本变化"):
                    v = (t.get("stats") or {}).get(k)
                    if v:
                        bits.append(f"{k} {v}")
                concl = "；".join(bits) or "无异常"
            rows.append((t.get("name"), t.get("kind"), st, concl, t.get("output") or ""))
        self.fill_tree([("任务", 190), ("类型", 90), ("状态", 70), ("结论", 330),
                        ("产物路径", 400)], rows)
        self.remember(self.last_html)

    def _show_ledger(self):
        root = self.t_led.get() or self.b_out.get() or os.getcwd()
        self.t_led.set(root)
        led = ops.Ledger(root)
        self.log(led.summary())
        rows = [(v.get("kind"), v.get("status"), os.path.basename(v.get("output") or ""),
                 v.get("at"), v.get("output"))
                for v in led.data.get("entries", {}).values()]
        self.show_summary({"台账任务数": len(rows)}, f"台账：{led.path}")
        self.fill_tree([("类型", 90), ("状态", 70), ("产物", 260), ("时间", 150),
                        ("完整路径", 430)], rows)

    def _reset_ledger(self):
        root = self.t_led.get() or self.b_out.get() or os.getcwd()
        led = ops.Ledger(root)
        if os.path.exists(led.path) and messagebox.askyesno(
                APP_TITLE, f"确定清除台账？\n{led.path}\n（清除后下次批量会全部重跑）"):
            os.remove(led.path)
            self.log(f"已清除台账：{led.path}")

    def _recalc(self):
        p = self.t_xlsx.get()
        if not p or not os.path.exists(p):
            messagebox.showwarning(APP_TITLE, "请选择要重算的 Excel 工作簿。")
            return
        outdir = os.path.join(os.path.dirname(p), "recalculated")

        def job():
            self.log("正在调用内置 LibreOffice 重算公式 ……")
            return {"out": ops.recalc_workbook(p, outdir=outdir)}

        self.start("公式重算", job, lambda r: (
            self.show_summary({"输出": r["out"]}), open_folder(r["out"]),
            self.remember(r["out"])))


def samples_files(out: str) -> list:
    from .samples import build_all
    return build_all(out, log=lambda m: None)


# --------------------------------------------------------------------------
# 帮助文本
# --------------------------------------------------------------------------

HELP_TEXT = """【五个页签做什么】
① 合并       两个/三个文档合成一份。给基线=三方合并（只有双方改到同一处才冲突）；不给基线=两个不同文档的对比整合。
             冲突默认在正文里同时保留双方：黄色=我方，青色=对方，绿色=取自对方的表格单元格；
             文末附「溯源附录」：冲突清单 / 表格行级处理 / 段落搬移 / 逐块来源。
② 文表对账   正文里的金额、比例、工期、日期、编号 与 Excel/CSV 逐项核对，自动做单位换算（元·万元·亿元、%·‰、天·个月·年）
             与标签模糊匹配；报告会把正文清单、数据表清单和数据表原文整合进同一份 Word。
③ 引用体检   悬空引用（见 3.2 节 / 图 5 / 附件 2）、编号重复、编号断档、模糊引用（见上文）。
④ 口径漂移   同一指标在多个版本间的取值变化；以及同一份文档里同一个指标写了两个值（自相矛盾）。
             只放一个文件就只做矛盾检查。
⑤ 批量处理   选目录 →「扫描并生成清单」自动发现 *_base/_ours/_theirs 组合、正文+数据表配对、版本序列，
             再「开始批处理」。没变化的文件按内容指纹自动跳过，产物含 HTML 汇总报告与增量台账。

【产物颜色图例】黄色=我方内容；青色=对方内容；绿色=取自对方的表格单元格或冲突标注；
                红色文字=冲突提示与编号（C1、C2…）。

【常见问题】
· 提示「旧版二进制格式」 → 该文件是 .doc/.xls，请先用 Word/Excel 另存为 .docx/.xlsx。
· 提示「输入文档含修订痕迹」 → 处理的是当前显示内容，建议先在 Word 里接受/拒绝修订。
· 表格没做单元格级合并 → 需要两侧列数一致且没有合并单元格；否则会整表保留双方并标为结构冲突。
· Excel 数字不新 → 勾选「先用内置 LibreOffice 重算公式」，或先在 Excel 里打开保存一次。
· 合并结果里某段位置不对 → 看报告的「段落搬移」章节，工具会把搬移识别出来而不是当成删除+新增。

【命令行等价用法】菜单里的每个功能都有对应命令，便于接进自动化流水线：
  docweave merge 基线.docx 我方.docx 对方.docx -o 合并.docx
  docweave audit 合同.docx 指标表.xlsx -o 对账.docx
  docweave refs 方案.docx -o 体检.docx
  docweave drift v1.docx v2.docx v3.docx -o 漂移.docx
  docweave plan ./文档 && docweave batch ./文档/out/manifest.json
"""


# --------------------------------------------------------------------------
# 自检（打包后的 exe 用这个模式验证）
# --------------------------------------------------------------------------

def selftest(outdir: str = None) -> int:
    """不弹窗地跑一遍全流程，结果写到 _selftest.json。用于验证 exe 可用。"""
    from .samples import build_all

    out = os.path.abspath(outdir or os.path.join(os.getcwd(), "docweave_selftest"))
    fix = os.path.join(out, "fixtures")
    res = {"workdir": out, "steps": [], "ok": False}

    def step(name, fn):
        try:
            info = fn()
            res["steps"].append({"step": name, "ok": True, "info": info})
            return info
        except Exception as exc:                            # noqa: BLE001
            res["steps"].append({"step": name, "ok": False,
                                 "error": f"{exc}",
                                 "trace": traceback.format_exc()})
            return None

    os.makedirs(out, exist_ok=True)
    step("生成演示数据", lambda: {"files": len(build_all(fix))})
    m = step("三方合并", lambda: ops.run_merge(
        os.path.join(fix, "contract_base.docx"),
        os.path.join(fix, "contract_ours.docx"),
        os.path.join(fix, "contract_theirs.docx"),
        os.path.join(out, "contract_merged.docx"),
        report=os.path.join(out, "contract_conflicts.md"),
        sidecars=os.path.join(out, "contract_conflicts.json")))
    a = step("文表对账", lambda: ops.run_audit(
        os.path.join(fix, "audit_doc.docx"),
        os.path.join(fix, "audit_data.xlsx"),
        out=os.path.join(out, "audit_report.docx"),
        report=os.path.join(out, "audit_report.md"),
        sidecars=os.path.join(out, "audit_findings.json")))
    r = step("引用体检", lambda: ops.run_refs(
        os.path.join(fix, "refs_doc.docx"),
        out=os.path.join(out, "refs_report.docx"),
        report=os.path.join(out, "refs_report.md")))
    d = step("口径漂移", lambda: ops.run_drift(
        [os.path.join(fix, f"spec_v{i}.docx") for i in (1, 2, 3)],
        out=os.path.join(out, "drift_report.docx"),
        report=os.path.join(out, "drift_report.md")))
    man = os.path.join(out, "manifest.json")
    step("生成批量清单", lambda: ops.build_manifest(fix, outdir=out, out_path=man))
    b = step("批量处理", lambda: ops.run_batch(man, outdir=out, force=True, quiet=True))

    checks = {
        "合并冲突数": (m or {}).get("conflicts", 0),
        "对账数值冲突": ((a or {}).get("stats") or {}).get("数值冲突", 0),
        "体检悬空引用": ((r or {}).get("stats") or {}).get("悬空引用", 0),
        "漂移自相矛盾": ((d or {}).get("stats") or {}).get("文档内自相矛盾", 0),
        "批量成功": ((b or {}).get("summary") or {}).get("成功", 0),
    }
    res["checks"] = checks
    produced = ["contract_merged.docx", "audit_report.docx", "refs_report.docx",
                "drift_report.docx", "manifest.json"]
    res["artifacts"] = {f: os.path.exists(os.path.join(out, f)) for f in produced}
    res["html_exists"] = os.path.exists(os.path.join(out, "batch_report.html"))
    res["ok"] = (all(s["ok"] for s in res["steps"])
                 and all(res["artifacts"].values())
                 and checks["合并冲突数"] >= 1
                 and checks["对账数值冲突"] >= 1
                 and checks["体检悬空引用"] >= 1
                 and checks["漂移自相矛盾"] >= 1
                 and checks["批量成功"] >= 5)
    with open(os.path.join(out, "_selftest.json"), "w", encoding="utf-8") as f:
        json.dump(res, f, ensure_ascii=False, indent=1)
    return 0 if res["ok"] else 1


def _pump(root, cond, timeout=180.0):
    """在无人工干预的情况下驱动 Tk 事件循环，直到条件满足或超时。"""
    import time

    t0 = time.time()
    while time.time() - t0 < timeout:
        root.update()
        if cond():
            return True
        time.sleep(0.05)
    return False


def uitest(workdir: str = None) -> int:
    """界面级自检：真的建出所有页签，逐个页签跑一遍，检查结果是否回填。

    不显示主窗口（withdraw），因此可以在无人值守时运行。
    """
    import time

    from .samples import build_all

    work = os.path.abspath(workdir or os.path.join(os.getcwd(), "docweave_uitest"))
    fix = os.path.join(work, "fixtures")
    out = os.path.join(work, "out")
    os.makedirs(fix, exist_ok=True)
    os.makedirs(out, exist_ok=True)
    if not os.path.exists(os.path.join(fix, "contract_base.docx")):
        build_all(fix)

    res = {"workdir": work, "tabs": [], "ok": False}
    root = tk.Tk()
    root.withdraw()
    _tune_fonts()
    try:
        app = App(root)

        def note(tab, ok, detail):
            res["tabs"].append({"tab": tab, "ok": bool(ok), "detail": detail})

        # ① 合并
        app.m_base.set(os.path.join(fix, "contract_base.docx"))
        app.m_ours.set(os.path.join(fix, "contract_ours.docx"))
        app.m_theirs.set(os.path.join(fix, "contract_theirs.docx"))
        app.m_out.set(os.path.join(out, "merged.docx"))
        app._run_merge()
        _pump(root, lambda: not app.busy)
        rows = app.tree.get_children()
        note("合并", len(rows) >= 1 and "冲突" in app.summary.cget("text")
             or len(rows) >= 1, f"结果行 {len(rows)}；{app.summary.cget('text')[:60]}")

        # ② 文表对账
        app.a_doc.set(os.path.join(fix, "audit_doc.docx"))
        app.a_data.set(os.path.join(fix, "audit_data.xlsx"))
        app.a_out.set(os.path.join(out, "audit.docx"))
        app._run_audit()
        _pump(root, lambda: not app.busy)
        rows_base = len(app.tree.get_children())
        note("文表对账", rows_base >= 3, f"发现 {rows_base} 条")

        # 勾选框联动：勾上「同时列出核对一致的项目」应当多出几条
        app.a_showok.set(True)
        app._run_audit()
        _pump(root, lambda: not app.busy)
        rows_ok = len(app.tree.get_children())
        app.a_showok.set(False)
        note("对账选项联动", rows_ok > rows_base, f"{rows_base} 条 → 勾选后 {rows_ok} 条")

        # ③ 引用体检
        app.r_doc.set(os.path.join(fix, "refs_doc.docx"))
        app.r_out.set(os.path.join(out, "refs.docx"))
        app._run_refs()
        _pump(root, lambda: not app.busy)
        rows = app.tree.get_children()
        note("引用体检", len(rows) >= 5, f"发现 {len(rows)} 条")

        # ④ 口径漂移
        app.d_list.delete(0, "end")
        for i in (1, 2, 3):
            app.d_list.insert("end", os.path.join(fix, f"spec_v{i}.docx"))
        app.d_out.set(os.path.join(out, "drift.docx"))
        app._run_drift()
        _pump(root, lambda: not app.busy)
        rows = app.tree.get_children()
        note("口径漂移", len(rows) >= 4, f"发现 {len(rows)} 条")

        # ⑤ 批量处理
        app.b_dir.set(fix)
        app.b_man.set(os.path.join(out, "manifest.json"))
        app.b_out.set(out)
        app._plan()
        _pump(root, lambda: not app.busy)
        plan_rows = len(app.tree.get_children())
        app.b_force.set(True)
        app._run_batch()
        _pump(root, lambda: not app.busy, timeout=300)
        rows = app.tree.get_children()
        note("批量处理", plan_rows >= 3 and len(rows) >= 5,
             f"清单 {plan_rows} 个任务，执行 {len(rows)} 条")

        # ⑥ 工具页：台账与帮助
        app.t_led.set(out)
        app._show_ledger()
        led_rows = len(app.tree.get_children())
        note("台账", led_rows >= 5, f"台账 {led_rows} 条")

        res["ok"] = all(t["ok"] for t in res["tabs"])
    except Exception:                                       # noqa: BLE001
        res["error"] = traceback.format_exc()
    finally:
        try:
            root.destroy()
        except Exception:
            pass
    with open(os.path.join(work, "_uitest.json"), "w", encoding="utf-8") as f:
        json.dump(res, f, ensure_ascii=False, indent=1)
    return 0 if res["ok"] else 1


# --------------------------------------------------------------------------

def _tune_fonts():
    try:
        import tkinter.font as tkfont
        for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont"):
            try:
                f = tkfont.nametofont(name)
                f.configure(family="Microsoft YaHei UI", size=10)
            except Exception:
                pass
    except Exception:
        pass


def main(argv=None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)
    # 窗口模式下没有 stdout，任何 print 都会炸，先兜住
    if sys.stdout is None or sys.stderr is None:
        devnull = open(os.devnull, "w", encoding="utf-8")
        sys.stdout = sys.stdout or devnull
        sys.stderr = sys.stderr or devnull
    if "--selftest" in argv:
        i = argv.index("--selftest")
        target = argv[i + 1] if len(argv) > i + 1 and not argv[i + 1].startswith("-") else None
        return selftest(target)
    if "--uitest" in argv:
        i = argv.index("--uitest")
        target = argv[i + 1] if len(argv) > i + 1 and not argv[i + 1].startswith("-") else None
        return uitest(target)
    if "--version" in argv:
        print(APP_VERSION)
        return 0
    demo_mode = "--demo-mode" in argv

    root = tk.Tk()
    root.title(APP_TITLE)
    _tune_fonts()
    try:
        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")
        style.configure("Treeview", rowheight=24)
        style.configure("Treeview.Heading", font=("Microsoft YaHei UI", 10, "bold"))
    except Exception:
        pass
    try:
        ico = resource("assets/docweave.ico")
        if os.path.exists(ico):
            root.iconbitmap(ico)
    except Exception:
        pass
    w, h = 1140, 880
    root.geometry(f"{w}x{h}")
    root.minsize(960, 700)
    try:
        root.update_idletasks()
        x = (root.winfo_screenwidth() - w) // 2
        y = max(0, (root.winfo_screenheight() - h) // 2 - 20)
        root.geometry(f"+{x}+{y}")
    except Exception:
        pass
    App(root)
    app = None
    for child in root.winfo_children():
        if isinstance(child, App):
            app = child
            break

    def _report(exc, val, tb):
        txt = "".join(traceback.format_exception(exc, val, tb))
        try:
            if app is not None:
                app.log("界面回调异常（已捕获，未崩溃）：\n" + txt)
                app.set_status("发生错误，详见日志", "#c00000")
            else:
                sys.stderr.write(txt)
        except Exception:
            pass

    root.report_callback_exception = _report
    if demo_mode:
        def _auto():
            try:
                app._demo_merge()
                app._run_merge()
            except Exception:
                app.log("演示模式启动失败：\n" + traceback.format_exc())

        root.after(1500, _auto)
    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
