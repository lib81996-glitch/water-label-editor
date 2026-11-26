# -*- coding: utf-8 -*-
"""选打印机的小弹窗。

为什么做成弹窗：主界面那格只留「现在用哪台」，列表挪进来 ——
一台机器上常挂着七八台打印机（虚拟 PDF/Fax、OneNote、共享机），平铺非常占地方，
而平时只关心一件事：这一张往哪台打。

排序和主界面一致：用过的次数多的在前，次数相同默认打印机优先，再按名字（见 eng.rank_printers）。
"""
import tkinter as tk
from tkinter import ttk

import label_engine as eng

FONT = "Microsoft YaHei"


class PrinterPicker:
    """一次选择；结果放在 self.result（取消为 None）"""

    # 端口列要放得下 PORTPROMPT:（11 个半角，约 110px），窄了就被切成 PORTPROM
    COLS = (("name", "打印机", 240, "w"),
            ("driver", "驱动", 180, "w"),
            ("port", "端口", 114, "w"),
            ("use", "用过", 52, "center"))

    def __init__(self, parent, infos=None, current="", usage=None, ranked=None):
        self._parent = parent
        self.current = current or ""
        self.usage = dict(eng.printer_usage() if usage is None else usage)
        self.result = None
        self.infos = []
        self._build()

        if ranked is not None:            # 测试里直接给排好的顺序
            self.infos = list(ranked)
            self._fill()
        else:
            self.reload(infos)

    # ---------------- 界面 ----------------
    def _build(self):
        w = tk.Toplevel(self._parent)
        self.win = w
        w.title("选择打印机")
        w.minsize(560, 320)
        w.transient(self._parent)
        try:
            px, py = self._parent.winfo_rootx(), self._parent.winfo_rooty()
            pw, ph = self._parent.winfo_width(), self._parent.winfo_height()
            w.geometry("660x380%+d%+d" % (px + max(0, (pw - 660) // 2),
                                          py + max(0, (ph - 380) // 3)))
        except tk.TclError:
            w.geometry("660x380")

        tip = ttk.Frame(w, padding=(10, 8, 10, 0))
        tip.pack(fill=tk.X)
        ttk.Label(tip, text="双击或选中后点「用这台」；按 Esc 取消",
                  foreground="#888").pack(side=tk.LEFT)

        box = ttk.Frame(w, padding=(10, 6, 10, 0))
        box.pack(fill=tk.BOTH, expand=True)
        self.tree = ttk.Treeview(box, columns=[c[0] for c in self.COLS],
                                 show="headings", selectmode="browse", height=9)
        for key, txt, wid, anch in self.COLS:
            self.tree.heading(key, text=txt)
            self.tree.column(key, width=wid, anchor=anch,
                             stretch=(key in ("name", "driver")))
        sb = ttk.Scrollbar(box, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.bind("<Double-1>", lambda e: self._ok())
        self.tree.tag_configure("cur", background="#E8F3EC")

        bar = ttk.Frame(w, padding=(10, 8, 10, 10))
        bar.pack(fill=tk.X)
        ttk.Button(bar, text="刷新", command=self.reload).pack(side=tk.LEFT)
        ttk.Button(bar, text="用这台", command=self._ok).pack(side=tk.RIGHT)
        ttk.Button(bar, text="取消", command=self._cancel).pack(side=tk.RIGHT, padx=(0, 6))
        self.lbl_n = ttk.Label(bar, text="", foreground="#888")
        self.lbl_n.pack(side=tk.LEFT, padx=(10, 0))

        w.protocol("WM_DELETE_WINDOW", self._cancel)
        w.bind("<Escape>", lambda e: self._cancel())
        w.bind("<Return>", lambda e: self._ok())

    # ---------------- 数据 ----------------
    def reload(self, infos=None):
        """重新枚举；保留原来选中的那台

        只有真的重扫（infos=None）才重读使用次数 —— 传进来一份现成列表时
        沿用手上这份计数，测试里注入的假数据才不会被冲掉。
        """
        if infos is None:
            infos = eng.list_printer_info()
            self.usage = dict(eng.printer_usage())
        self.infos = list(infos)
        self._fill()

    def _fill(self):
        sel = self.tree.selection()
        keep = sel[0] if sel else self.current
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        # 和主界面同一套规则：次数降序 → 默认机 → 名字
        order = sorted(self.infos, key=lambda d: (-self.usage.get(d["name"], 0),
                                                  not d.get("is_default"),
                                                  d["name"].lower()))
        for info in order:
            name = info["name"]
            n = self.usage.get(name, 0)
            mark = "  ●" if name == self.current else ""
            self.tree.insert("", "end", iid=name,
                             values=(name + mark, info.get("driver", ""),
                                     info.get("port", ""), n or ""),
                             tags=("cur",) if name == self.current else ())
        if keep and self.tree.exists(keep):
            self.tree.selection_set(keep)
            self.tree.see(keep)
        elif self.tree.get_children():
            first = self.tree.get_children()[0]
            self.tree.selection_set(first)
            self.tree.see(first)
        self.lbl_n.config(text="共 %d 台" % len(self.infos))

    # ---------------- 结果 ----------------
    def _ok(self):
        sel = self.tree.selection()
        if sel:
            self.result = sel[0]
        self._close()

    def _cancel(self):
        self.result = None
        self._close()

    def _close(self):
        try:
            self.win.grab_release()
        except tk.TclError:
            pass
        self.win.destroy()


def choose(parent, current=""):
    """打开弹窗选一台；返回打印机名，取消返回 None"""
    dlg = PrinterPicker(parent, current=current)
    try:
        dlg.win.grab_set()
        dlg.win.focus_set()
    except tk.TclError:
        pass
    parent.wait_window(dlg.win)
    return dlg.result
