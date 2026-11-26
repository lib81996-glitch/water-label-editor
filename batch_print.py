# -*- coding: utf-8 -*-
"""批量逐步打印窗口。

为什么做成独立窗口而不是第三个标签页：
  · 现有「打印队列」「打印预览」都是独立窗口，批量打印属同一类辅助面板；
  · 右侧 Notebook 那两个页签都是「当前这个项目」的两个切面，而批量是 N 个文件，
    塞进去会把「当前项目」的语义搅乱；
  · 打印过程中主窗口还能接着改后面的文件，两窗并排互不干扰；
  · 关掉窗口队列照跑 —— 标签页没有「关掉」这个动作。

怎么做到「打完一个停下，点继续打下一个」：
  不新开线程。任务照旧塞进主窗口已有的打印队列，只是把 qstate 置成 paused ——
  队列线程取不到下一个任务就自然停在那里；点「继续」把 qstate 置回 running
  并叫醒它。这套 paused 逻辑队列里本来就有（终止后自动暂停用的就是它）。
"""
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter import font as tkfont

import label_engine as eng

EXT = ".wslabel"
SAME_AS_FILE = "（按每个文件各自的设置）"


class BatchPrintWindow:
    """一次只允许存在一个；由 App.open_batch() 负责创建和复用。"""

    COLS = (("idx", "#", 40, "center"),
            ("name", "文件", 200, "w"),
            ("printer", "打印机", 156, "w"),
            ("mode", "方式", 56, "center"),
            ("copies", "份数", 68, "center"),
            ("status", "状态", 186, "w"))

    def __init__(self, app):
        self.app = app
        self.rows = []            # [{path, proj, copies, task, iid, _vals}]
        self.lastdir = ""
        self._editor = None       # 点份数格时临时浮出来的小输入框
        self._edit_iid = None
        self._top = None          # 置顶状态的当前值，避免每次都调 attributes
        self._tags = None         # 每行当前的行底色标签，没变就别重设

        w = tk.Toplevel(app.root)
        self.win = w
        w.title("批量打印")
        # 最小尺寸按「顶部那一行不打架」定：宽度再小「等继续时置顶」就和「清空」叠一起了；
        # 高度要够放「份数」行和底部动作行，否则一缩就被啃掉（实测踩过）。
        w.minsize(860, 500)
        # 开在主窗口中间偏上，而不是屏幕左上角 —— 否则一开就压住预览和内容区
        try:
            px, py = app.root.winfo_rootx(), app.root.winfo_rooty()
            pw, ph = app.root.winfo_width(), app.root.winfo_height()
            w.geometry("880x600%+d%+d" % (px + max(0, (pw - 880) // 2),
                                          py + max(0, (ph - 600) // 3)))
        except tk.TclError:
            w.geometry("880x600")
        w.protocol("WM_DELETE_WINDOW", self.close)
        self._build()
        self.refresh(force=True)

    # ---------------- 界面 ----------------
    def _build(self):
        w = self.win

        st = ttk.Style()
        # 只改字色，不动 padding —— vista 主题被覆盖 padding 后文字会被裁
        st.configure("Danger.TButton", foreground="#C62828")
        top = ttk.Frame(w, padding=(10, 8, 10, 0))
        top.pack(fill=tk.X)
        ttk.Label(top, text="打印机").pack(side=tk.LEFT)
        self.v_printer = tk.StringVar(value=SAME_AS_FILE)
        self.cb_printer = ttk.Combobox(top, textvariable=self.v_printer, width=26,
                                       state="readonly",
                                       values=[SAME_AS_FILE] + eng.list_printers())
        self.cb_printer.pack(side=tk.LEFT, padx=(6, 0))
        self.cb_printer.bind("<<ComboboxSelected>>", lambda e: self._apply_printer())
        self.v_top = tk.BooleanVar(value=True)
        ttk.Checkbutton(top, text="等「继续」时置顶", variable=self.v_top).pack(
            side=tk.LEFT, padx=(12, 0))

        # 「添加文件」是这一屏唯一的正事，给实底；「清空」是破坏性操作，
        # 红字 + 拉开，别让手滑点到。
        # 空隙开在「移除选中」右侧而不是「清空」右侧 —— 清空和移除选中同属删除类，
        # 挨在一起才读得成一组；空隙的位置决定它跟谁分离。
        tk.Button(top, text="＋ 添加文件", command=self.add_files,
                  bg="#1F6F45", fg="white", activebackground="#185236",
                  activeforeground="white", relief="flat", bd=0, cursor="hand2",
                  font=("Microsoft YaHei UI", 9, "bold"), padx=16, pady=5).pack(side=tk.RIGHT)
        self.btn_del = ttk.Button(top, text="移除选中",                command=self.remove_selected)
        self.btn_del.pack(side=tk.RIGHT, padx=(0, 24))
        self.btn_clear = ttk.Button(top, text="清空", style="Danger.TButton",
                                    command=self.clear_all)
        self.btn_clear.pack(side=tk.RIGHT, padx=(0, 4))

        box = ttk.Frame(w, padding=(10, 8, 10, 0))
        self.tree = ttk.Treeview(box, columns=[c[0] for c in self.COLS],
                                 show="headings", selectmode="extended", height=12)
        for key, txt, wid, anch in self.COLS:
            self.tree.heading(key, text=txt)
            self.tree.column(key, width=wid, anchor=anch,
                             stretch=(key == "name"))
        # 份数那一列要看清数字、也好点：行高放大一点，
        # 就地输入框用大一号的字体（ttk.Treeview 没有单元格控件，只能浮个 Spinbox 上去）
        style = ttk.Style()
        self._edit_font = tkfont.nametofont("TkDefaultFont").copy()
        self._edit_font.configure(size=self._edit_font.cget("size") + 2)
        style.configure("Batch.Treeview", rowheight=30)
        style.configure("Big.TSpinbox", font=self._edit_font)
        self.tree.configure(style="Batch.Treeview")
        sb = ttk.Scrollbar(box, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        # 点一下就直接输入，不用双击（壮哥：不用双击直接输入数值）
        self.tree.bind("<ButtonRelease-1>", self._on_cell_click)
        self.tree.bind("<Double-1>", self._on_double)
        self.tree.bind("<Delete>", lambda e: self.remove_selected())
        self.tree.tag_configure("next", background="#FFF6E5")
        self.tree.tag_configure("fail", background="#FFEBEE")

        opt = ttk.Frame(w, padding=(10, 8, 10, 0))
        self.v_step = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt, text="打完一个先停下，等我点「继续」",
                        variable=self.v_step,
                        command=self._on_step_toggle).pack(side=tk.LEFT)
        ttk.Label(opt, text="份数").pack(side=tk.LEFT, padx=(18, 4))
        self.v_allcopies = tk.StringVar(value="1")
        ttk.Spinbox(opt, from_=0, to=999, width=4, justify="center",
                    textvariable=self.v_allcopies).pack(side=tk.LEFT)
        ttk.Button(opt, text="全部应用", command=self.apply_all_copies).pack(side=tk.LEFT, padx=(8, 0))
        # 打错打印机重来时，前面已经打好的不能再出一遍纸；
        # 但「整批再打一份」是正常需求，所以给个开关，别替用户决定。
        self.v_skip = tk.BooleanVar(value=True)
        ttk.Checkbutton(opt, text="跳过已打完的",
                        variable=self.v_skip).pack(side=tk.LEFT, padx=(14, 0))
        ttk.Label(opt, text="份数填 0 = 跳过这个文件",
                  foreground="#888").pack(side=tk.LEFT, padx=(10, 0))

        row = ttk.Frame(w, padding=(10, 8, 10, 10))
        self.lbl_prog = ttk.Label(row, text="还没添加文件", foreground="#666")
        self.lbl_prog.pack(side=tk.LEFT)
        self.btn_stop = ttk.Button(row, text="终止当前", style="Danger.TButton",
                                   command=self.stop)
        self.btn_stop.pack(side=tk.RIGHT)
        self.btn_resume = ttk.Button(row, text="▶ 继 续",                   command=self.resume)
        self.btn_resume.pack(side=tk.RIGHT, padx=(0, 8))
        self.btn_start = tk.Button(row, text="开 始", command=self.start,
                                   bg="#1F6F45", fg="white", activebackground="#185236",
                                   activeforeground="white", relief="flat", bd=0,
                                   cursor="hand2", font=("Microsoft YaHei UI", 9, "bold"),
                                   disabledforeground="#BFD4C8",
                                   padx=22, pady=7)
        self.btn_start.pack(side=tk.RIGHT, padx=(0, 8))
        # 这一行要和进度文字、三个按钮共享宽度，字多了会把提示挤没（实测踩过）
        ttk.Label(row, text="点份数格直接改",
                  foreground="#aaa").pack(side=tk.RIGHT, padx=(0, 14))

        # 摆位顺序：先把底部两行贴底放好，最后再把表格塞进剩下的空间。
        # 反过来的话（表格先 pack + expand），窗口一矮就轮到「排在最后的」被啃掉，
        # 表现就是底部那一行按钮被窗口下沿切掉 —— 壮哥碰到过。
        row.pack(side=tk.BOTTOM, fill=tk.X)
        opt.pack(side=tk.BOTTOM, fill=tk.X)
        box.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

    # ---------------- 文件管理 ----------------
    def add_files(self):
        paths = filedialog.askopenfilenames(
            parent=self.win,
            title="选择要批量打印的水洗标文件",
            initialdir=self.lastdir or os.path.dirname(os.path.abspath(__file__)),
            filetypes=[("水洗标项目", "*" + EXT), ("全部文件", "*.*")])
        if not paths:
            return
        self.lastdir = os.path.dirname(paths[0])
        bad = []
        for p in paths:
            try:
                proj = eng.normalize(eng.load_project(p))
            except Exception as e:
                bad.append("%s：%s" % (os.path.basename(p), e))
                continue
            self.rows.append({"path": p, "proj": proj,
                              "copies": self._all_copies(),
                              "task": None, "iid": None, "_vals": None,
                              "_mode": None})
        self._rebuild()
        if bad:
            messagebox.showwarning("有文件没读进来",
                                   "\n".join(bad[:8]) + ("\n…" if len(bad) > 8 else ""),
                                   parent=self.win)

    def remove_selected(self):
        sel = list(self.tree.selection())
        if not sel:
            return
        losing = [self.rows[int(i)] for i in sel if i.isdigit()]
        self._drop_queued(losing)
        self.rows = [r for i, r in enumerate(self.rows) if str(i) not in sel]
        self._rebuild()

    def clear_all(self):
        self._drop_queued(self.rows)
        self.rows = []
        self._rebuild()

    def apply_all_copies(self):
        n = self._all_copies()
        for i in range(len(self.rows)):
            self._set_copies(i, n)
        self.refresh(force=True)

    def _set_copies(self, i, n):
        """改一行的份数；若这一行已经排进队但还没开打，把队里那份也改掉。

        不改的话表格显示 3、队里还按 1 打，用户看到的就是「表里表外不一致」。
        正在打的那一份不动 —— 张数在开工那一刻就定了，中途改会打出一半。
        """
        if not (0 <= i < len(self.rows)):
            return
        r = self.rows[i]
        r["copies"] = n
        r["_vals"] = None
        with self.app.qlock:
            for t in self.app.tasks:
                if t["id"] == r["task"] and t["status"] == "排队中":
                    t["copies"] = n
                    t["prog"] = "0/%d" % n

    def _all_copies(self):
        try:
            return max(0, min(999, int(str(self.v_allcopies.get()).strip())))
        except (TypeError, ValueError):
            return 1

    def _drop_queued(self, rows):
        """把这几行还排在队里、没开打的任务撤掉（已经打的不动）"""
        ids = {r["task"] for r in rows if r["task"]}
        if not ids:
            return
        with self.app.qlock:
            self.app.tasks = [t for t in self.app.tasks
                              if not (t["id"] in ids and t["status"] == "排队中")]

    # ---------------- 打印器覆盖 ----------------
    def _apply_printer(self):
        """换打印机：表格里的通道/文字全失效重算，还没打的结果一并作废。

        为什么连「排队中」的也要撤：队里的任务在入队那一刻就把打印机抄下来了
        （快照），换了机器它还是会往旧机器打。留着的话表格显示新机器、实际打旧机器，
        比直接说「这些还没打」更坑人。打印中和已完成的当然不动。
        """
        reset, stale = 0, []
        for r in self.rows:
            r["_mode"] = None        # 换了打印机，通道可能跟着变（佳博才走直发）
            r["_vals"] = None
            st = self._task_status(r["task"])
            if st == "排队中":
                stale.append(r)
                reset += 1
            elif st == "失败":
                r["task"] = None     # 上一台打失败的结果，换了机器就不再算数
                reset += 1
        self._drop_queued(stale)
        for r in stale:
            r["task"] = None
        self.refresh(force=True)
        msg = "打印机已改为：%s" % self.v_printer.get()
        if reset:
            msg += "（%d 个未打完的已重置，点「开始」用新打印机打）" % reset
        self.app.status.config(text=msg)

    def _printer_of(self, r):
        name = self.v_printer.get()
        if name and name != SAME_AS_FILE:
            return name
        return r["proj"].get("printer") or ""

    def _mode_of(self, r):
        """打印通道（TSPL 直发 / 图片）—— 算一次就存住。

        resolve_mode() 遇到非佳博的热敏机要枚举一遍打印机驱动（实测 0.21ms），
        而 refresh 每 250ms 要过一遍所有行；60 行就是每秒 240 次，白烧。
        打印机换了或行变了才重算，见 _apply_printer() 里的失效。
        """
        if r.get("_mode") is None:
            try:
                r["_mode"] = ("图片" if eng.resolve_mode(r["proj"]) == "image"
                              else "TSPL")
            except Exception:
                r["_mode"] = "图片"
        return r["_mode"]

    # ---------------- 表格 ----------------
    def _rebuild(self):
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        self._tags = None
        for i, r in enumerate(self.rows):
            r["iid"] = str(i)
            r["_vals"] = None
            self.tree.insert("", "end", iid=r["iid"],
                             values=("%d" % (i + 1), os.path.basename(r["path"]),
                                     self._printer_of(r), self._mode_of(r),
                                     r["copies"], self._status_text(
                                         r, None, "idle", False)))
        self.refresh(force=True)

    def _on_cell_click(self, event):
        """点一下「份数」格子就地输入 —— 不用双击。

        ttk.Treeview 没有「单元格控件」，只能在那一格上浮一个 Spinbox：
        点一次就浮出来、数字全选好，直接敲数字即可覆盖；回车或点别处提交，
        Esc 取消，上下小箭头加减。点别的格子会先把上一个收起来。
        """
        iid = self.tree.identify_row(event.y)
        if not iid:
            self._close_editor()
            return None
        if not iid.isdigit() or int(iid) >= len(self.rows):
            return "break"
        # 不用 identify_column() 判「点的是不是份数列」：实测它和真正画出来的
        # 格子边界差十几个像素（x=528 明明落在份数格里，它回报 #4 模式列），
        # 靠它会出现「点格子左边一点没反应」。按 bbox 矩形判，和眼睛看到的一致。
        box = self.tree.bbox(iid, "#5")
        if not box or not (box[0] <= event.x <= box[0] + box[2]):
            self._close_editor()
            return None                  # 不是份数列：交回默认处理（选中/滚动照常）
        if self._editor is not None and self._edit_iid == iid:
            return "break"               # 已经在编这一格，别重建（会吞掉刚敲的字）
        x, y, w, h = box
        self._close_editor()
        var = tk.StringVar(value=str(self.rows[int(iid)]["copies"]))
        sp = ttk.Spinbox(self.tree, from_=0, to=999, textvariable=var,
                         justify="center", style="Big.TSpinbox")
        sp.place(x=x, y=y, width=w, height=h)
        self._editor = sp
        self._edit_iid = iid
        sp.focus_set()
        sp.selection_range(0, "end")

        def commit(_=None):
            self._commit_copies(int(iid), var.get())
            self._close_editor()

        sp.bind("<Return>", commit)
        sp.bind("<KP_Enter>", commit)
        sp.bind("<FocusOut>", commit)
        sp.bind("<Escape>", lambda e: self._close_editor())
        return "break"

    # 双击也走同一条路（老习惯保留；单击已经能直接输入了）
    def _on_double(self, event):
        return self._on_cell_click(event)

    def _close_editor(self):
        self._edit_iid = None
        if self._editor is not None:
            try:
                self._editor.destroy()
            except tk.TclError:
                pass
            self._editor = None

    def _commit_copies(self, idx, raw):
        if not (0 <= idx < len(self.rows)):
            return
        try:
            n = int(str(raw).strip())
        except (TypeError, ValueError):
            n = 1
        self._set_copies(idx, max(0, min(999, n)))
        self.refresh(force=True)

    # ---------------- 状态刷新 ----------------
    def _status_text(self, r, t, state, is_next):
        if t is None:
            return "未开始"
        st = t["status"]
        if st == "排队中":
            # 只有真的在「单步等人点继续」时才算等待，否则如实说排队中 ——
            # 失败停住的时候也挂着「等待继续」会让人以为点一下就能接着打
            if state == "paused" and is_next and self.v_step.get():
                return "等待继续"
            return "排队中"
        if st == "打印中":
            return ("打印中 %s" % t.get("prog", "")).strip()
        if st == "失败":
            err = (t.get("err") or "").strip()
            if len(err) > 11:
                err = err[:11] + "…"
            return "失败：%s" % err
        return st

    def refresh(self, force=False):
        if not self.win.winfo_exists():
            return
        app = self.app
        with app.qlock:
            state = app.qstate
            tasks = {t["id"]: dict(t) for t in app.tasks}

        nxt = None
        for r in self.rows:
            t = tasks.get(r["task"])
            if t and t["status"] == "排队中":
                nxt = r
                break

        # 行底色：失败标浅红、「下一个要打的」标浅黄 —— 扫一眼就知道该看哪行
        tags = {}
        for r in self.rows:
            t = tasks.get(r["task"])
            tags[r["iid"]] = ("fail" if t and t["status"] == "失败"
                              else "next" if r is nxt else ())
        if tags != self._tags:
            for iid, tg in tags.items():
                if iid and self.tree.exists(iid):
                    self.tree.item(iid, tags=tg)
            self._tags = tags

        done = sheets = total = 0
        printing = any(t["status"] == "打印中" for t in tasks.values())
        for i, r in enumerate(self.rows):
            t = tasks.get(r["task"])
            vals = ("%d" % (i + 1), os.path.basename(r["path"]),
                    self._printer_of(r), self._mode_of(r), r["copies"],
                    self._status_text(r, t, state, r is nxt))
            if force or vals != r["_vals"]:
                if r["iid"] and self.tree.exists(r["iid"]):
                    self.tree.item(r["iid"], values=vals)
                r["_vals"] = vals
            if t and t["status"] == "已完成":
                done += 1
                sheets += int(t["copies"])
            if r["copies"] > 0:
                total += int(r["copies"])

        if not self.rows:
            self.lbl_prog.config(text="还没添加文件", foreground="#666")
        elif done and not nxt and not printing:
            self.lbl_prog.config(text="已全部打完：%d 个文件 / %d 张"
                                 % (done, sheets), foreground="#1b5e20")
        else:
            self.lbl_prog.config(text="已完成 %d / 共 %d 个文件　计 %d 张"
                                 % (done, len(self.rows), total), foreground="#666")

        self.btn_start.config(state=tk.DISABLED if printing else tk.NORMAL)
        self.btn_resume.config(state=tk.NORMAL if (state == "paused" and nxt)
                               else tk.DISABLED)
        self.btn_stop.config(state=tk.NORMAL if printing else tk.DISABLED)

        want_top = bool(self.v_top.get() and state == "paused" and nxt)
        if want_top != self._top:
            try:
                self.win.attributes("-topmost", want_top)
                if want_top:
                    self.win.lift()
                self._top = want_top
            except tk.TclError:
                pass

    # ---------------- 打印控制 ----------------
    def _on_step_toggle(self):
        self.app.batch_step = bool(self.v_step.get())
        if not self.app.batch_step and self.app.qstate == "paused":
            if any(t["status"] == "排队中" for t in self.app.tasks):
                self.app.q_resume()      # 取消单步 = 立刻把剩下的打完
                self.app.status.config(text="已取消单步，剩余任务自动打完")

    def start(self):
        app = self.app
        if not self.rows:
            messagebox.showinfo("没有文件", "先点「＋ 添加文件」选几个水洗标。",
                                parent=self.win)
            return
        # 只有「正在打」才拦。之前是按「排队中/打印中」一起拦，
        # 结果上一个文件打失败后剩下的还挂在队里排队，这里就永远拒绝开始 ——
        # 用户换好打印机却怎么也打不了（实测踩过）。
        if any(self._task_status(r["task"]) == "打印中" for r in self.rows if r["task"]):
            messagebox.showinfo(
                "正在打印",
                "当前这个还没打完。等它打完，或先点「终止当前」。",
                parent=self.win)
            return

        # 上一批剩下的排队任务：先撤掉。它们带着旧打印机（就是错的那台），
        # 留着会跟新任务一起打 —— 既重复又继续失败。
        stale = [r for r in self.rows if self._task_status(r["task"]) == "排队中"]
        self._drop_queued(stale)

        # 已经打完的文件默认不再重排，否则改完打印机一点开始，前面打好的又出一遍纸。
        # 勾掉「跳过已打完的」才会整批重打；整批都打完过时也按重打处理（否则按钮像坏了）。
        pool = [r for r in self.rows if r["copies"] > 0]
        if not pool:
            messagebox.showinfo("没有要打的",
                                "所有文件的份数都是 0。改一下份数再点开始。",
                                parent=self.win)
            return
        done_ids = {id(r) for r in self.rows
                    if self._task_status(r["task"]) == "已完成"}
        todo = pool
        if self.v_skip.get():
            todo = [r for r in pool if id(r) not in done_ids]
            if not todo:
                todo = pool             # 全都打完过 → 整批重打一遍

        no_printer = []
        app.batch_step = bool(self.v_step.get())

        # 整批里有没有「内容比纸长」的：批量是没人逐张盯的，打出来才发现就是一叠废纸。
        # 逐个问太烦，汇总成一句话问一次。
        over = []
        for r in (todo if self.v_skip.get() else pool):
            n = eng.page_overflow_mm(r["proj"])
            if n > 0:
                over.append("%s（超 %.1f mm）" % (r["name"], n))
        if over and not messagebox.askyesno(
                "有内容超出标签",
                "%d 个文件的内容比标签长，下半截会被裁掉：\n\n%s\n\n仍然开始打印吗？"
                % (len(over), "\n".join(over[:5]) + ("\n…" if len(over) > 5 else "")),
                parent=self.win):
            return
        # 先把队列按住，等全部任务排好再放开。
        # 否则第一条可能在下面的循环还没跑完时就被打完了，
        # 「后面还有活」判断成 False，单步暂停就漏掉了（实测踩过）。
        with app.qlock:
            app.qstate = "paused"
        for r in stale:
            r["task"] = None             # 队里已经撤掉了，这条线的结果作废
            r["_vals"] = None
        for r in todo:
            r["task"] = None
            r["_vals"] = None
        # 注意：已完成的行不动它的 task —— 留着才显示得出「已完成」，
        # 一并清掉会变成「未开始」，看不出这些是刚打过的。
        for r in todo:
            proj = dict(r["proj"])
            proj["printer"] = self._printer_of(r)
            if not proj["printer"]:
                no_printer.append(os.path.basename(r["path"]))
                continue
            proj["fabrics"] = list(r["proj"].get("fabrics") or [])
            t = app.enqueue_snapshot(proj, r["copies"])
            r["task"] = t["id"]
        if no_printer:
            messagebox.showwarning(
                "有文件没有打印机",
                "这几个文件里没记打印机，也在上面选一台：\n" + "\n".join(no_printer[:6]),
                parent=self.win)
        app.q_resume()               # 上一批可能停在 paused，这里明确放开
        app.status.config(text="批量打印已开始：%d 个文件%s"
                          % (len(todo), "（单步）" if app.batch_step else "（连续）"))
        self.refresh(force=True)

    def resume(self):
        self.app.batch_step = bool(self.v_step.get())
        self.app.q_resume()
        self.refresh(force=True)

    def stop(self):
        self.app.q_abort()

    def _task_status(self, tid):
        with self.app.qlock:
            for t in self.app.tasks:
                if t["id"] == tid:
                    return t["status"]
        return None

    def _next_index(self):
        """下一个还排在队里的行号；没有就返回 0"""
        with self.app.qlock:
            tasks = {t["id"]: t for t in self.app.tasks}
        for i, r in enumerate(self.rows):
            t = tasks.get(r["task"])
            if t and t["status"] == "排队中":
                return i
        return 0

    def close(self):
        paused_left = (self.app.qstate == "paused"
                       and any(t["status"] == "排队中" for t in self.app.tasks))
        self._close_editor()
        self.win.destroy()
        self.app.batch_win = None
        if paused_left:
            self.app.status.config(
                text="批量打印停在第 %d 个 —— 点状态栏的「继续批量」接着打"
                     % (self._next_index() + 1))
