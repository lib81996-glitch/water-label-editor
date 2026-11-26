# -*- coding: utf-8 -*-
"""
========================================================================
水洗标签编辑器 v1.12
========================================================================
简版标签软件 —— 只做水洗标这一件事。

能做什么：
  · 改纸宽 / 纸长（任意 mm），实时预览
  · 填号型、成分（多行）
  · 底部洗涤说明：标题 / 英文副标题 / 注意事项 都是文字，图标单独嵌入
  · 调字号、行距、左边距、起始高度
  · 保存成 .wslabel 单文件，下次双击继续改
  · 一键打印到佳博（TSPL 直发，不走任何驱动渲染）

不能做什么（故意的）：
  · 拖拽排版、条形码、二维码 —— 水洗标用不上
========================================================================
"""
import ctypes
import json
import math
import os
import re
import sys
import threading
import time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)
os.chdir(BASE_DIR)          # 锁定工作目录，避免从别处启动找不到模块
import label_engine as eng
import ddl_import as ddl     # 从 DLabel 的 .ddl 里取号型 / 成分
import batch_print as bprint  # 批量逐步打印窗口（独立弹窗）
import printer_pick as ppick  # 选择打印机的弹窗
import icon_pick as ipick      # 选洗涤图标的弹窗
# caption_bar.py 已停用：自绘标题栏会接管非客户区，把四边缩放吃掉，
# 换到别的电脑还容易超屏。保留文件只为留档，不再 import。

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)   # 高 DPI 清晰
except Exception:
    pass

EXT = ".wslabel"
FONT = "Microsoft YaHei"
# 窗口标题：名字 + 署名 + 文件名。署名放在文件名前面 ——
# 文件名一长，标题栏和任务栏都会从右边截，署名在前才不会被截掉。
APP_NAME = "水洗标签编辑器"
APP_SIGN = "By Yaoshi_Yellow"

# 自动存盘：草稿写在程序目录下的 .autosave，正常保存/退出会删掉，
# 只有崩溃、断电、任务管理器强杀才会留下 —— 下次启动时提供恢复。
# ★ 打包成单文件后 __file__ 指向解压出来的临时目录（退出即删），
#   草稿写在那儿等于没写 —— 所以这里认「exe 所在目录」，和 printer_usage.json 同一套。
APP_DIR = (os.path.dirname(os.path.abspath(sys.executable))
           if getattr(sys, "frozen", False) else BASE_DIR)
AUTOSAVE_DIR = os.path.join(APP_DIR, ".autosave")
AUTOSAVE_MS = 30000          # 30 秒一次，只写有改动的标签
AUTOSAVE_MAX = 30            # 草稿最多留几份：按标签名一对一写，同名覆盖，
                             # 超出这个数就把最旧的（按修改时间）删掉。
                             # 正常存盘/正常退出都会删，只有崩溃、断电、
                             # 任务管理器强杀才会真正留下来。留个上限是为了
                             # 万一长期这么用，.autosave 别无限长成一堆垃圾。
UNDO_LIMIT = 60              # 撤销栈上限（每步是一份 JSON 快照，约 2KB/份）

# 图标笔画粗细：名字 → 二值化阈值（阈值越高，笔画越粗）
TH_NAME = {"细": 128, "中": 170, "粗": 205}
TH_BY_VALUE = {v: k for k, v in TH_NAME.items()}

# 按尺码生成表格里的候选尺码。顺序就是打印出来的顺序（不是勾选顺序）：
# 尺码表得按大小排，勾完再回头改顺序没人受得了。
SIZE_ORDER = ["XS", "S", "M", "L", "XL", "2XL", "3XL", "4XL", "5XL", "6XL", "均码"]

# 打印通道：界面显示名 → 项目里存的值
# 自动加长纸：纸底留多少。取 6mm 略大于「纸底保留线」的 5mm，
# 免得算出来的长度正好卡在边线上，溢出判据因为浮点误差来回跳。
AUTO_TAIL_MM = 6.0

MODE_NAME = {"自动": "auto", "TSPL 直发": "tspl", "图片模式": "image"}
MODE_BY_VALUE = {v: k for k, v in MODE_NAME.items()}

# 纸张形态（跟着每个文件存）：连续纸没有缝隙，长度只能由文件尺寸决定
MEDIA_NAME = {"间隙纸": "gap", "连续纸": "continuous"}
MEDIA_BY_VALUE = {v: k for k, v in MEDIA_NAME.items()}

# 洗涤块的位置模式：界面显示名 → 项目里存的值
ALIGN_NAME = {"紧跟正文": "flow", "垂直居中": "center", "贴纸底": "bottom"}
ALIGN_BY_VALUE = {v: k for k, v in ALIGN_NAME.items()}

# 日期那一行的水平对齐（和上面的「洗涤块位置」是两回事，别混用）
DATE_ALIGN_NAME = {"居中": "center", "靠左": "left", "靠右": "right"}
DATE_ALIGN_BY_VALUE = {v: k for k, v in DATE_ALIGN_NAME.items()}


class FlowRow(ttk.Frame):
    """按可用宽度自动折行的按钮条。

    为什么需要它：底部按钮栏排成一行要 **1218 px**，而窗口允许拖到 900 px ——
    拖窄以后右边几个按钮就被容器裁掉了（实测 1000 px 时「保存」少 39 px、
    「打开」整块不见）。两种解法里：
      · 把最小宽度钉到 1218：小屏电脑上窗口就比屏幕还宽，又回到「显示不全」；
      · 让它自己折行：宽度够时还是一行（和原来一模一样），不够就折成两行。
    选后者 —— 「无论怎么拖都完整」的唯一诚实做法是不让任何控件被裁。

    实现上用 grid 而不是 pack：grid 可以把同一个控件重新指派到新的行列，
    pack 做不到（换布局管理器必须先 forget 再换，容易抖）。
    """

    def __init__(self, master, gap=8, **kw):
        super().__init__(master, **kw)
        self._gap = gap
        self._items = []          # [(widget, 前间距)]
        self._last_w = -1
        self._rows = 1
        self.bind("<Configure>", self._on_configure)

    def add(self, w, pad=0):
        """把一个控件加入流（顺序即显示顺序）"""
        self._items.append((w, pad))
        return w

    def _on_configure(self, ev):
        # 宽度没变就别重排：重排会改高度 -> 又触发 Configure -> 无限抖
        if abs(ev.width - self._last_w) < 2:
            return
        self._last_w = ev.width
        self.reflow(max(1, ev.width))

    def reflow(self, avail):
        """把控件铺到 avail 像素宽里，超了换行。返回用了几行。"""
        if not self._items:
            return 0
        avail = max(1, int(avail))
        row = 0
        col = 0
        used = 0
        for w, pad in self._items:
            need = w.winfo_reqwidth() + pad
            if col and used + need > avail:
                row += 1
                col = 0
                used = 0
            w.grid(row=row, column=col, sticky="w",
                   padx=(pad, self._gap), pady=(0, 3))
            used += need + self._gap
            col += 1
        self._rows = row + 1
        return self._rows

    def needed_width(self):
        """全部挤在一行需要多宽（用来判断该不该折行）"""
        return sum(w.winfo_reqwidth() + pad + self._gap for w, pad in self._items)


class App:
    def __init__(self, root, path=None):
        self.root = root
        self.path = path
        # ---- 多项目标签页：一个 tab = 一个文件项目 ----
        self.tabs = [{"proj": eng.normalize(eng.load_project(path) if path else {}),
                      "path": path, "dirty": False}]
        self.cur = 0
        self.proj = self.tabs[0]["proj"]      # 当前标签的项目（视图）
        self._building = True
        self._photos = []       # 防止 ImageTk 对象被回收
        self._scroll_canvases = []   # 两个表单页的滚动 Canvas，尺寸变化时重算 scrollregion
        self._scroll_bars = {}       # {canvas名: 对应滚动条}，同步显隐用
        # ---- 打印队列（多任务）----
        self.tasks = []          # [{id,name,proj,copies,printer,mode,status,prog,err}]
        self.task_seq = 0
        self.qstate = "idle"     # idle / running / paused
        self.q_paused_by = ""    # 暂停原因：abort（终止）/ fail（打不出来）/ step（单步）
        self.abort_flag = False
        self.qlock = threading.Lock()
        # 事件驱动：没有任务时线程一直睡着，入队/点「继续」才叫醒它。
        # 原来是 while True + sleep(0.25) 空转，队列空着也每秒醒 4 次。
        self.qevt = threading.Event()
        self.qstop = threading.Event()
        self._qstat_last = None             # 队列状态文字的上次值，避免每 250ms 白刷一次
        self.qwin = None
        # ---- 批量逐步打印（独立窗口，见 batch_print.py）----
        self.batch_win = None
        self.batch_step = False     # True = 一个任务打完就暂停，等用户点「继续」
        self.btn_bresume = None     # 主界面上的「继续批量」入口（单步暂停时出现）
        # ---- 撤销 / 重做（快照式，旁路功能，不参与渲染与打印）----
        self._undo_stack = []               # [(tab_index, 改动前的 json)]
        self._redo_stack = []
        self._undo_pending = None           # 连续敲字合并成一步
        self._undo_job = None
        self._autosave_job = None
        self._build_ui()
        self._sync_to_ui()
        self._building = False
        # 窗口尺寸一变（含被 fit 收进小屏），表单滚动条立刻跟上
        root.bind("<Configure>", lambda e: self._refresh_scroll_areas(), add="+")
        threading.Thread(target=self._queue_worker, daemon=True).start()
        root.after(120, self.redraw)
        root.after(250, self._poll_queue)
        # 撤销/重做快捷键（Text 控件没开 undo，不会被抢）
        self.root.bind_all("<Control-z>", self._on_undo_key)
        self.root.bind_all("<Control-y>", self._on_redo_key)
        self.root.bind_all("<Control-Z>", self._on_redo_key)   # Ctrl+Shift+Z 也重做
        self._schedule_autosave()
        root.after(400, self._offer_recovery)   # 上次没正常退出的草稿，启动后问一次

    # ---------------- UI ----------------
    def _build_ui(self):
        r = self.root
        # 用系统原生标题栏（不再自绘灰条）：v1.8 为了改颜色接管了非客户区，
        # 结果四边缩放被吃掉、换电脑还超屏。回到原生，颜色跟随系统主题，
        # 拖拽缩放 / 双击最大化 / 贴边吸附全部由 Windows 自己负责。
        r.title(APP_NAME + "  " + APP_SIGN)
        # 不写死 geometry —— 交给 Tk 按控件实际需要自适应，
        # 这样系统缩放不是 100% 时也不会把窗口压小、截断界面。
        r.minsize(900, 680)

        # ===== 项目标签页（多文件，可点击跳转）=====
        self.tabstrip = ttk.Frame(r, padding=(10, 5, 10, 0))
        self.tabstrip.pack(side=tk.TOP, fill=tk.X)
        self._refresh_tabs()

        # ===== 顶部：纸张 =====
        top = ttk.LabelFrame(r, text=" 纸张 ", padding=8)
        top.pack(side=tk.TOP, fill=tk.X, padx=10, pady=(10, 4))

        self.v_w = tk.StringVar()
        self.v_h = tk.StringVar()
        self.v_gap = tk.StringVar()
        self.v_dpi = tk.StringVar()

        def pair(parent, col, label, var, unit):
            ttk.Label(parent, text=label).grid(row=0, column=col, sticky="e", padx=(10, 4))
            e = ttk.Entry(parent, textvariable=var, width=8, justify="right")
            e.grid(row=0, column=col + 1)
            ttk.Label(parent, text=unit).grid(row=0, column=col + 2, sticky="w", padx=(4, 0))
            return e

        pair(top, 0, "纸宽", self.v_w, "mm")
        self.ent_h = pair(top, 3, "纸长", self.v_h, "mm")
        self.ent_gap = pair(top, 6, "间隙", self.v_gap, "mm")
        pair(top, 9, "分辨率", self.v_dpi, "dpi")
        # 成分多敲几个空行就能把纸顶穿，勾上这个就不用自己算该给多长。
        # 只往长里长，不会比手设的短；取消勾选就回到手设值。
        # 默认勾上：这是最常见的翻车点（多写一行成分，底下洗涤块就被挤出纸），
        # 让人自己去打勾，等于让每个人先踩一次坑。
        self.v_autoh = tk.BooleanVar(value=True)
        ttk.Checkbutton(top, text="内容超长时自动加长纸",
                        variable=self.v_autoh,
                        command=self._on_autoh).grid(row=0, column=12, padx=(14, 0))

        # 纸张形态：决定「这张纸多长」由谁说话。
        #   间隙纸 —— 每张之间有缝，机器靠缝传感器定位，长度基本由纸决定；
        #   连续纸 —— 整卷连着没有缝，机器找不到缝就只能按指令/纸长走，
        #             所以长度必须严格等于「纸长」，间隙那一格也就没意义了。
        # 这个值跟着文件存（每个 .wslabel 各带自己的）—— 壮哥的实际情况
        # 就是有的文件打连续纸、有的打间隙纸。
        #
        # 为什么挤在同一行末尾、不另起一行：单开一行只多 30px 看着没事，
        # 实际把内容请求高度顶到 1003，底部按钮栏被整条挤出屏幕 ——
        # 这类"多加一行"是最容易犯、也最难当场看出来的错。
        ttk.Label(top, text="纸张").grid(row=0, column=13, padx=(16, 4))
        self.v_media = tk.StringVar(value="间隙纸")
        self.cb_media = ttk.Combobox(top, textvariable=self.v_media, width=8,
                                     state="readonly", values=list(MEDIA_NAME))
        self.cb_media.grid(row=0, column=14)
        self.lbl_media_hint = ttk.Label(top, text="", foreground="#888")
        self.lbl_media_hint.grid(row=0, column=15, sticky="w", padx=(8, 0))

        # ===== 主体 =====
        main = ttk.Frame(r)
        main.pack(fill=tk.BOTH, expand=True, padx=10, pady=4)

        # 左：预览
        left = ttk.LabelFrame(main, text=" 预览（等比缩放，非实际大小） ", padding=4)
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.canvas = tk.Canvas(left, bg="#eceff1", highlightthickness=0,
                                width=560, height=660)
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.canvas.bind("<Configure>", lambda e: self.redraw())

        # 右：内容（分两个标签页，避免面板太长）
        right = ttk.LabelFrame(main, text=" 内容 ", padding=6)
        right.pack(side=tk.RIGHT, fill=tk.Y, padx=(8, 0))

        nb = ttk.Notebook(right)
        nb.pack(fill=tk.BOTH, expand=True)
        self.nb = nb

        # 右上角工具组（贴最右）：表格工具 / 自检页。
        # 1.9.1 调整：原来「表格工具」是单独 place 在页签右边界外侧，
        # 窗口一窄就压住「底部洗涤说明」页签文字；「查看 TSPL 指令」占着
        # 工具组第一个位置。按壮哥的图换成：
        #   · 表格工具 → 并入工具组，占原 TSPL 的位置（不再单独算坐标）；
        #   · 查看 TSPL 指令 → 移到底部按钮栏（批量打印与保存之间）。
        # 两个按钮都进同一个 pack 流，宽度自动算，永远不会互相压。
        tools = ttk.Frame(right)
        # 贴 right 右边缘。滚动条显示时要往左让 22px，由
        # _layout_right_buttons() 按滚动条实际显隐动态调整 x。
        tools.place(relx=1.0, y=6, x=-4, anchor="ne")
        # 一个入口管两件事：没有表格就新建，已经有了就改 —— 所以叫「表格工具」，
        # 不叫「插入」（插进去之后想再改就没地方点了）。
        self.btn_table = tk.Button(tools, text="表格工具",
                                   font=(FONT, 10, "bold"),
                                   bg="#3C8C4E", fg="white",
                                   activebackground="#46A05A", activeforeground="white",
                                   relief="flat", bd=0, highlightthickness=0,
                                   padx=12, pady=3, cursor="hand2",
                                   command=self.open_table_editor)
        self.btn_table.pack(side=tk.LEFT)
        self.btn_table.bind("<Enter>",
                            lambda e: self.btn_table.config(bg="#46A05A"))
        self.btn_table.bind("<Leave>",
                            lambda e: self.btn_table.config(bg="#3C8C4E"))
        self.btn_selftest = ttk.Button(tools, text="自检页", width=8,
                                       command=self.do_selftest)
        self.btn_selftest.pack(side=tk.LEFT, padx=(6, 0))
        self._right_tools = tools
        nb.bind("<Configure>", lambda e: self._layout_right_buttons())
        self.root.after(30, self._layout_right_buttons)

        # 两个表单页都要能滚 —— 窗口被强行收进小屏后，下面几行会被裁掉。
        # 技巧：nb.add 的是「外层容器」，但把变量名 t1/t2 指向里面的内容 frame，
        # 于是下面几十处 ttk.Label(t1,...) 一行都不用改，自动就进了滚动区。
        t1 = self._scroll_page(nb, "  标签内容  ")
        t2 = self._scroll_page(nb, "  底部洗涤说明  ")

        # ---------- 标签页 1 ----------
        self.v_brand = tk.BooleanVar(value=True)
        ttk.Checkbutton(t1, text="品牌：AOGZ（参与打印）",
                        variable=self.v_brand).pack(anchor="w", pady=(0, 8))

        ttk.Label(t1, text="号型 / 第一行").pack(anchor="w")
        self.v_title = tk.StringVar()
        ttk.Entry(t1, textvariable=self.v_title).pack(fill=tk.X, pady=(2, 10))

        fab_head = ttk.Frame(t1)
        fab_head.pack(fill=tk.X)
        ttk.Label(fab_head, text="成分（一行一个，超出右边距自动换行）").pack(side=tk.LEFT)
        # 空行分段：默认关，老文件照旧。开了以后成分框里敲的空行会在标签上
        # 真的空出一行 —— 想把「面料 / 里料 / 填充物」分开时用得上。
        self.v_keepblank = tk.BooleanVar(value=False)
        ttk.Checkbutton(fab_head, text="保留空行分段",
                        variable=self.v_keepblank).pack(side=tk.RIGHT)
        self.txt_fabric = tk.Text(t1, height=7, wrap="none", font=(FONT, 10),
                                  relief="solid", borderwidth=1)
        self.txt_fabric.pack(fill=tk.X, pady=(2, 10))

        # 产地单独一行：塞进成分框也能打，但那一行会跟着成分一起挤、一起缩进，
        # 想跟成分拉开距离就没地方调了。留空 = 不占位，老文件不受影响。
        orow = ttk.Frame(t1)
        orow.pack(fill=tk.X)
        # 先摆右边的勾选框：它用 side=RIGHT 先占位，否则输入框 expand 会把它挤出窗口
        self.v_center = tk.BooleanVar(value=False)
        ttk.Checkbutton(orow, text="正文居中", variable=self.v_center).pack(
            side=tk.RIGHT, padx=(10, 0))
        ttk.Label(orow, text="产地").pack(side=tk.LEFT)
        self.v_origin = tk.StringVar()
        ttk.Entry(orow, textvariable=self.v_origin).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(6, 0))
        ttk.Label(t1, text="产地留空就不打印；「正文居中」对品牌/号型/成分/产地一起生效",
                  foreground="#888", justify="left").pack(anchor="w", pady=(2, 6))

        # 生产日期：英文月简称 + 日（NOV 18），正文下面单独一行居中。
        # 月份从 12 个里选、日期手填，字号可调 —— 都是存下来的值，
        # 不是「每次现取当天」，补打同批货日期才不会被改掉。
        drow = ttk.Frame(t1)
        drow.pack(fill=tk.X)
        self.v_date = tk.BooleanVar(value=True)
        self.v_datemon = tk.StringVar()
        self.v_dateday = tk.StringVar()
        self.v_datefont = tk.StringVar()
        ttk.Checkbutton(drow, text="印生产日期", variable=self.v_date,
                        command=self._sync_date_row).pack(side=tk.LEFT)
        self._cb_datemon = ttk.Combobox(
            drow, textvariable=self.v_datemon, width=7, state="readonly",
            values=list(eng.MONTH_CHOICES))
        self._cb_datemon.pack(side=tk.LEFT, padx=(6, 0))
        self._ent_dateday = ttk.Entry(drow, textvariable=self.v_dateday, width=3)
        self._ent_dateday.pack(side=tk.LEFT, padx=(4, 0))
        ttk.Label(drow, text="字号").pack(side=tk.LEFT, padx=(10, 2))
        self._ent_datefont = ttk.Entry(drow, textvariable=self.v_datefont, width=3)
        self._ent_datefont.pack(side=tk.LEFT)
        ttk.Label(drow, text="dot").pack(side=tk.LEFT, padx=(2, 0))
        self._btn_today = ttk.Button(drow, text="今天", width=5,
                                     command=self._date_today)
        self._btn_today.pack(side=tk.LEFT, padx=(8, 0))

        # 手动自定义内容：填了就印这行字（替代 NOV 24），留空照旧用自动日期。
        # 单独起一行 —— 上面那排已经排满，硬塞进去会把「今天」挤出可视区。
        crow = ttk.Frame(t1)
        crow.pack(fill=tk.X, pady=(6, 0))
        ttk.Label(crow, text="自定义").pack(side=tk.LEFT)
        self.v_datecustom = tk.StringVar()
        self._ent_datecustom = ttk.Entry(crow, textvariable=self.v_datecustom,
                                         justify="center")
        self._ent_datecustom.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(6, 0))
        ttk.Label(crow, text="对齐").pack(side=tk.LEFT, padx=(10, 4))
        self.v_datealign = tk.StringVar(value="居中")
        self._cb_datealign = ttk.Combobox(
            crow, textvariable=self.v_datealign, width=6, state="readonly",
            values=["居中", "靠左", "靠右"])
        self._cb_datealign.pack(side=tk.LEFT)
        ttk.Label(t1, text="自定义填了就印你填的字，留空则印上面的日期；"
                           "默认居中，太长会自动缩字号；字号 8–48 dot（12≈1.5mm）",
                  foreground="#888", justify="left").pack(anchor="w", pady=(2, 8))

        grid = ttk.Frame(t1)
        grid.pack(fill=tk.X)

        self.v_left = tk.StringVar()
        # v_right 在下面「右边距」那一行就地创建（它和左边距同一行，控件位置靠后）
        self.v_starty = tk.StringVar()
        self.v_font = tk.StringVar()
        self.v_lg = tk.StringVar()
        self.v_indent = tk.StringVar()
        self.v_div = tk.BooleanVar(value=False)
        self.v_divw = tk.StringVar(value="20.0")
        self.v_divg = tk.StringVar(value="1.25")

        def row(rr, label, var, unit, values=None, width=10, editable=False):
            ttk.Label(grid, text=label, width=8, anchor="e").grid(row=rr, column=0, padx=(0, 6), pady=3)
            if values:
                w = ttk.Combobox(grid, textvariable=var, values=values, width=7,
                                 state=("normal" if editable else "readonly"))
            else:
                w = ttk.Entry(grid, textvariable=var, width=width, justify="right")
            w.grid(row=rr, column=1, sticky="w")
            ttk.Label(grid, text=unit, foreground="#666").grid(row=rr, column=2, sticky="w", padx=(4, 0))
            return w

        row(0, "左边距", self.v_left, "mm")
        # 右边距与左边距同一行
        ttk.Label(grid, text="右边距", width=8, anchor="e").grid(row=0, column=3, padx=(18, 6), pady=3)
        self.v_right = tk.StringVar()
        ttk.Entry(grid, textvariable=self.v_right, width=7, justify="right").grid(row=0, column=4, sticky="w")
        ttk.Label(grid, text="mm", foreground="#666").grid(row=0, column=5, sticky="w", padx=(4, 0))
        row(1, "起始高度", self.v_starty, "mm")
        row(2, "字体大小", self.v_font, "dot",
            ["12", "16", "20", "24", "28", "32", "40"], editable=True)
        row(3, "行间距", self.v_lg, "dot")

        # 字体：只列常见中文字体 + 浏览任意 ttf。选标签字体不需要看见系统里那几百个。
        self.v_fontsel = tk.StringVar()
        self._font_map = {"自动（雅黑 / 宋体 / 黑体）": ""}
        for _name, _path in eng.FONT_CHOICES[1:]:
            if eng.font_available(_path):
                self._font_map[_name] = _path
        # 接在「正文缩进」（row 4）后面 —— 那一行是后加的，字体不能抢它的格子
        ttk.Label(grid, text="字体", width=8, anchor="e").grid(
            row=5, column=0, padx=(0, 6), pady=3)
        self.cb_font = ttk.Combobox(grid, textvariable=self.v_fontsel,
                                    values=list(self._font_map), width=16,
                                    state="readonly")
        self.cb_font.grid(row=5, column=1, columnspan=2, sticky="w")
        ttk.Button(grid, text="浏览…", width=7,
                   command=self._pick_font).grid(row=5, column=3, sticky="w")
        # 只有「文件里的字体这台电脑没有」时才出字，平时不占地方
        self.lbl_fontwarn = ttk.Label(grid, text="", foreground="#B26A00",
                                      justify="left", wraplength=520)
        self.lbl_fontwarn.grid(row=6, column=0, columnspan=6, sticky="w",
                               pady=(2, 0))
        # 底版里号型块比左边距再靠右 3.7mm，而「洗涤方式」按整张纸居中 ——
        # 一个左边距管不了两件事，所以正文单独给一个缩进
        row(4, "正文缩进", self.v_indent, "mm")

        drow = ttk.Frame(t1)
        drow.pack(fill=tk.X, pady=(6, 0))
        ttk.Checkbutton(drow, text="分隔线", variable=self.v_div).pack(side=tk.LEFT)
        ttk.Label(drow, text="长", foreground="#666").pack(side=tk.LEFT, padx=(10, 2))
        ttk.Entry(drow, textvariable=self.v_divw, width=5,
                  justify="right").pack(side=tk.LEFT)
        ttk.Label(drow, text="mm　距正文", foreground="#666").pack(side=tk.LEFT, padx=(2, 2))
        ttk.Entry(drow, textvariable=self.v_divg, width=5,
                  justify="right").pack(side=tk.LEFT)
        ttk.Label(drow, text="mm", foreground="#666").pack(side=tk.LEFT, padx=(2, 0))

        ttk.Separator(t1).pack(fill=tk.X, pady=8)
        self.v_toprule = tk.BooleanVar(value=True)
        ttk.Checkbutton(t1, text="顶部装饰线（居中 5 × 0.4 mm，实色）",
                        variable=self.v_toprule).pack(anchor="w")

        # ---------- 标签页 2 ----------
        self.v_bottom = tk.BooleanVar(value=True)
        ttk.Checkbutton(t2, text="打印底部洗涤说明", variable=self.v_bottom).pack(anchor="w", pady=(0, 8))

        # 这三段文字各有各的字号：大标题通常比正文粗一点、副标题是英文可以小一号，
        # 捆成一个「总字号」就只能三者一起变，反而不好使。所以每段自带一个输入框。
        self.v_btfont = tk.StringVar()
        self.v_bsfont = tk.StringVar()
        self.v_bcfont = tk.StringVar()
        self.v_bfootfont = tk.StringVar()

        def head_row(text, var, first=False):
            """一行：左边是段落名，右边是这一段的字号（手动输入 dot）"""
            hf = ttk.Frame(t2)
            hf.pack(fill=tk.X, pady=(0 if first else 8, 0))
            ttk.Label(hf, text=text).pack(side=tk.LEFT)
            ttk.Label(hf, text="dot", foreground="#666").pack(side=tk.RIGHT)
            ent = ttk.Entry(hf, textvariable=var, width=4, justify="right")
            ent.pack(side=tk.RIGHT, padx=(0, 4))
            ttk.Label(hf, text="字号", foreground="#666").pack(side=tk.RIGHT, padx=(0, 4))
            return ent

        head_row("大标题", self.v_btfont, first=True)
        self.v_btitle = tk.StringVar()
        ttk.Entry(t2, textvariable=self.v_btitle).pack(fill=tk.X, pady=(2, 0))

        head_row("英文副标题", self.v_bsfont)
        self.v_bsub = tk.StringVar()
        ttk.Entry(t2, textvariable=self.v_bsub).pack(fill=tk.X, pady=(2, 0))

        head_row("注意事项（一行一句，居中）", self.v_bcfont)
        self.txt_bcaution = tk.Text(t2, height=5, wrap="none", font=(FONT, 10),
                                    relief="solid", borderwidth=1)
        self.txt_bcaution.pack(fill=tk.X, pady=(2, 2))

        # 页脚一行（MADE IN CHINA）：字号独立，不然它只能跟注意事项同字号
        head_row("页脚（如 MADE IN CHINA）", self.v_bfootfont)
        self.v_bfoot = tk.StringVar()
        ttk.Entry(t2, textvariable=self.v_bfoot).pack(fill=tk.X, pady=(2, 2))
        # 说实话比事后解释便宜：非 12/16/24/32 的字号打印机会当图片发过去，
        # 效果一模一样，只是数据量大一点、慢一点 —— 不写清楚会以为是 bug
        ttk.Label(t2, text="字号 12 / 16 / 24 / 32 走打印机内置字库；\n"
                           "其它数值自动转位图，打出来一模一样，只是数据大一点",
                  foreground="#888", justify="left").pack(anchor="w", pady=(5, 8))

        g2 = ttk.Frame(t2)
        g2.pack(fill=tk.X)

        def brow(rr, label, var, unit, values=None):
            lb = ttk.Label(g2, text=label, width=8, anchor="e")
            lb.grid(row=rr, column=0, padx=(0, 6), pady=3)
            if values:
                w = ttk.Combobox(g2, textvariable=var, values=values, width=6, state="readonly")
            else:
                w = ttk.Entry(g2, textvariable=var, width=7, justify="right")
            w.grid(row=rr, column=1, sticky="w")
            un = ttk.Label(g2, text=unit, foreground="#666")
            un.grid(row=rr, column=2, sticky="w", padx=(4, 0))
            return lb, un, w

        self.v_bgap = tk.StringVar()
        self._lbl_bgap, self._unit_bgap, self._ent_bgap = brow(
            0, "上方留白", self.v_bgap, "mm（内容底 → 洗涤方式顶）")
        # 位置模式：flow 紧跟正文 / center 垂直居中 / bottom 贴纸底
        self.v_balign = tk.StringVar(value=ALIGN_BY_VALUE["flow"])
        self._lbl_balign, _, self._cb_balign = brow(
            1, "位置", self.v_balign, "", values=list(ALIGN_NAME.values()))
        self._lbl_balign.config(width=8)
        self.v_iconth = tk.StringVar()
        brow(2, "笔画粗细", self.v_iconth, "中 = 约 2 dot，打出来实心",
             values=list(TH_NAME))

        # 一行只读状态：直接告诉用户「现在这两个数分别是多少、由谁说了算」
        self.lbl_bspan = ttk.Label(t2, text="", foreground="#888", justify="left")
        self.lbl_bspan.pack(anchor="w", pady=(6, 0))
        # 字号手输以后这行会变长（可能同时报「挤到纸底」和「比纸宽」），
        # 给个随宽度走的折行，免得它把右栏整个撑宽、把左边预览挤扁
        t2.bind("<Configure>",
                lambda e: self.lbl_bspan.configure(
                    wraplength=max(220, e.width - 24)))

        # 两个留白联动：动哪个，另一个立刻回填成真实结果（否则其中一个必然是死键）
        self._bsync = False
        self.v_bgap.trace_add("write", lambda *_: self._on_bgap_edit())

        # 图标：从「一个说明文字」改成真能选 —— 图标是有顺序的一串图，顺序错了意思就变了
        irow = ttk.Frame(t2)
        irow.pack(fill=tk.X, pady=(10, 0))
        ttk.Button(irow, text="选洗涤图标…", command=self._pick_icons).pack(side=tk.LEFT)
        self.lbl_icons = ttk.Label(irow, text="", foreground="#888")
        self.lbl_icons.pack(side=tk.LEFT, padx=(8, 0))
        ttk.Label(t2, text="顺序即打印顺序；图片会复制进 assets/icons/，换电脑跟着走",
                  foreground="#aaa").pack(anchor="w", pady=(2, 0))

        # 图标槽位：固定大小的格子，图片只在格子里等比缩小，位置永远不动。
        # 为什么不按「铺满纸宽」缩放：那样图标多大取决于放了几张、图本身多宽 ——
        # 少放一个，剩下的全跟着放大，版式每改一次就跳一次。槽位固定就不跳了。
        hrow = ttk.Frame(t2)
        hrow.pack(fill=tk.X, pady=(8, 0))
        ttk.Label(hrow, text="槽位").pack(side=tk.LEFT)
        self.v_islot = tk.StringVar()
        ttk.Entry(hrow, textvariable=self.v_islot, width=5).pack(
            side=tk.LEFT, padx=(4, 2))
        ttk.Label(hrow, text="mm", foreground="#888").pack(side=tk.LEFT)
        ttk.Label(hrow, text="间距").pack(side=tk.LEFT, padx=(10, 0))
        self.v_isgap = tk.StringVar()
        ttk.Entry(hrow, textvariable=self.v_isgap, width=5).pack(
            side=tk.LEFT, padx=(4, 2))
        ttk.Label(hrow, text="mm", foreground="#888").pack(side=tk.LEFT)
        ttk.Label(hrow, text="行距").pack(side=tk.LEFT, padx=(10, 0))
        self.v_isrow = tk.StringVar()
        ttk.Entry(hrow, textvariable=self.v_isrow, width=5).pack(
            side=tk.LEFT, padx=(4, 2))
        ttk.Label(hrow, text="mm", foreground="#888").pack(side=tk.LEFT)

        self.lbl_iconfit = ttk.Label(t2, text="", foreground="#888",
                                     justify="left")
        self.lbl_iconfit.pack(anchor="w", pady=(3, 0))
        ttk.Label(t2, text="一排最多 5 个，多了自动分排（6→3+3，7→3+4，8→4+4）；"
                           "图片只在槽内等比缩小，不拉伸、不放大",
                  foreground="#aaa").pack(anchor="w", pady=(1, 0))

        # ---------- 打印机列表（可刷新）+ 打印方式 ----------
        self.v_printer = tk.StringVar()
        # 打印位移 / 方向：主界面不给控件，由「打印预览」里调；放在这里是为了跟其它
        # 输入一样走 trace → _collect 的同一条通路，不另起一套状态。
        self.v_offx = tk.StringVar(value="0.00")
        self.v_offy = tk.StringVar(value="0.00")
        self.v_direction = tk.StringVar(value="0")
        pbox = ttk.LabelFrame(right, text=" 打印机 ", padding=4)
        pbox.pack(fill=tk.X, pady=(8, 0))

        # 平时只关心「这一张往哪台打」，所以不铺列表，只显示当前那台 + 一个切换入口。
        # 列表挪进弹出窗口（printer_pick.py），八台打印机挤一屏实在没必要。
        prow = ttk.Frame(pbox)
        prow.pack(fill=tk.X)
        self.lbl_printer = tk.Label(prow, text="（没选打印机）", anchor="w",
                                    bg="#F7F8FA", fg="#1F2329", relief="solid",
                                    bd=1, padx=6, pady=3, cursor="hand2")
        self.lbl_printer.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.lbl_printer.bind("<Button-1>", lambda e: self.choose_printer())
        ttk.Button(prow, text="切换…", command=self.choose_printer).pack(
            side=tk.LEFT, padx=(4, 0))
        self.lbl_printer_info = ttk.Label(pbox, text="", foreground="#888")
        self.lbl_printer_info.pack(anchor="w", pady=(3, 0))

        prow2 = ttk.Frame(pbox)
        prow2.pack(fill=tk.X, pady=(4, 0))
        ttk.Label(prow2, text="打印方式").pack(side=tk.LEFT, padx=(0, 4))
        self.v_mode = tk.StringVar()
        ttk.Combobox(prow2, textvariable=self.v_mode, width=10, state="readonly",
                     values=list(MODE_NAME)).pack(side=tk.LEFT)

        ttk.Label(pbox, text="别的打印机打出黑方块？切成「图片模式」即可",
                  foreground="#888").pack(anchor="w", pady=(4, 0))
        self.lbl_channel = ttk.Label(pbox, text="", foreground="#2f6f4f")
        self.lbl_channel.pack(anchor="w", pady=(2, 0))

        # ===== 状态区：操作反馈 / 队列状态 / 当前打印机 =====
        # 从按钮流里挪出来 —— 原来它夹在「批量打印」和「保存」中间，
        # 把动作区劈成两段，视线要跳过它才能接上。
        # 挪出后不另占一行，而是贴到底部那行的右端（与任务标签同行）。
        foot = ttk.Frame(r, padding=(12, 0, 12, 2))
        # before=main：让底部栏**先**拿它需要的高度，再让中间区域吃剩下的。
        # 默认顺序（main 先、foot 后）在窗口不够高时会把 foot 整个挤没 ——
        # 按钮栏凭空消失，用户只会以为程序坏了。1.10 加「纸张」那一行后
        # 内容请求高度涨到 1003，这个坑立刻现形。
        foot.pack(side=tk.BOTTOM, fill=tk.X, before=main)
        self._foot = foot
        self.lbl_ctx = ttk.Label(foot, text="", foreground="#767d84")
        self.lbl_ctx.pack(side=tk.RIGHT)
        # 环境检查入口：换电脑/换系统之后，「能不能跑」比「怎么排」更急。
        # 平时它只是一个小灰字，查出问题时变红 —— 不占地方，但随时点得到。
        self.lbl_env = ttk.Label(foot, text="环境检查", foreground="#8a94a0",
                                 cursor="hand2")
        self.lbl_env.pack(side=tk.RIGHT, padx=(0, 14))
        self.lbl_env.bind("<Button-1>", lambda e: self.check_env(verbose=True))
        self.qstat = ttk.Label(foot, text="队列空闲", foreground="#666")
        self.qstat.pack(side=tk.RIGHT, padx=(0, 14))
        self.status = ttk.Label(foot, text="就绪", foreground="#666")
        self.status.pack(side=tk.RIGHT, padx=(0, 14))
        self.root.after(80, self._update_status_ctx)

        # ===== 底部按钮（方案 A：只改颜色 / 高度 / 间距）=====
        bar = ttk.Frame(r, padding=(10, 2, 10, 8))
        # 同样 before=main：按钮栏比中间区域更要紧，窗口再矮也得露出来
        bar.pack(side=tk.BOTTOM, fill=tk.X, before=main)
        self.bar = bar
        _st = ttk.Style()      # 只建样式不覆盖 padding —— 一覆盖 vista 主题的字就被裁

        # 右侧文件组（保存 / 另存为 / 导入模板 / 打开）固定靠右，不参与折行。
        # anchor="n"：左边折成两行时它贴着上沿，不会掉到中间。
        grp_file = ttk.Frame(bar)
        grp_file.pack(side=tk.RIGHT, anchor="n")
        self._bar_file_grp = grp_file

        # 左侧动作区走 FlowRow：宽度够就是一行（和原来一模一样），
        # 拖窄了自动折行，永远不会有按钮被裁掉。
        flow = FlowRow(bar)
        flow.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self._bar_flow = flow

        flow.add(ttk.Label(flow, text="张数"))
        self.v_copies = tk.StringVar(value="1")
        flow.add(ttk.Spinbox(flow, from_=1, to=999, textvariable=self.v_copies,
                             width=4, justify="center"), pad=6)
        flow.add(ttk.Button(flow, text="打印预览", command=self.open_print_preview))
        flow.add(tk.Button(flow, text="打 印", command=self.do_print,
                           bg="#1F6F45", fg="white", activebackground="#185236",
                           activeforeground="white", relief="flat", bd=0, cursor="hand2",
                           font=("Microsoft YaHei UI", 9, "bold"), padx=18, pady=4), pad=4)
        flow.add(ttk.Button(flow, text="撤销", width=6, command=self.undo))
        flow.add(ttk.Button(flow, text="打印队列", command=self.open_queue))
        flow.add(ttk.Button(flow, text="批量打印", command=self.open_batch))
        # 查看 TSPL 指令：从右上角工具组挪到底部栏，落在「批量打印」与
        # 「保存」之间的空档（壮哥在截图里标的位置）。它是偶尔查一次的
        # 调试入口，放动作区末尾比挤在页签上更顺手，也给右上工具组腾位置。
        self.btn_tspl = flow.add(
            ttk.Button(flow, text="查看 TSPL 指令", command=self.show_tspl), pad=6)

        ttk.Button(grp_file, text="保存",
                   command=lambda: self.do_save(False)).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(grp_file, text="另存为",
                   command=lambda: self.do_save(True)).pack(side=tk.LEFT, padx=(0, 8))
        self.btn_import = ttk.Button(grp_file, text="导入模板", command=self.do_import_ddl)
        self.btn_import.pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(grp_file, text="打开", command=self.do_open).pack(side=tk.LEFT)

        # 首次布局：Configure 事件要等窗口映射后才来，先按当前宽度排一次，
        # 免得启动瞬间底栏是空的（一闪）。
        self._bar_last_w = -1
        self._bar_stacked = False
        bar.bind("<Configure>", lambda e: self._reflow_bar())
        self._bar_flow.after_idle(self._reflow_bar)

        # ===== 任务条（多任务跳转，位于按钮栏上方）=====
        # 位置理由：紧挨「打印 / 打印队列」动作区，又不挤占编辑面板；
        # 一个任务一枚彩色标签，像浏览器标签一样横排，状态一眼可辨。
        self.taskbar = ttk.Frame(foot)
        self.taskbar.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Label(self.taskbar, text="任务:").pack(side=tk.LEFT)
        self.taskstrip = ttk.Frame(self.taskbar)
        self.taskstrip.pack(side=tk.LEFT, padx=(6, 0))
        # 提示语缩一半：它和状态挤同一行，太长会把任务标签压掉
        self.lbl_taskhint = ttk.Label(self.taskbar, text="单击跳转 · 右键载入",
                                      foreground="#aaa")
        self.lbl_taskhint.pack(side=tk.LEFT, padx=(14, 0))
        # 窄的时候先让这句提示让位 —— 它是操作说明，不是信息本身；
        # 让位以后任务标签和右边那几个状态才有地方待着，不会被挤没。
        foot.bind("<Configure>", self._on_foot_resize)

        # 输入监听
        for v in (self.v_w, self.v_h, self.v_gap, self.v_dpi, self.v_title,
                  self.v_left, self.v_right, self.v_starty, self.v_font, self.v_lg, self.v_printer,
                  self.v_toprule, self.v_brand, self.v_bottom, self.v_btitle, self.v_bsub,
                  self.v_iconth, self.v_bgap, self.v_mode,
                  self.v_btfont, self.v_bsfont, self.v_bcfont,
                  self.v_offx, self.v_offy, self.v_direction,
                  self.v_keepblank, self.v_origin, self.v_center, self.v_fontsel,
                  self.v_date, self.v_datemon, self.v_dateday, self.v_datefont,
                  self.v_datecustom, self.v_datealign, self.v_media,
                  self.v_bfoot, self.v_bfootfont, self.v_balign,
                  self.v_islot, self.v_isgap, self.v_isrow,
                  self.v_indent, self.v_div, self.v_divw, self.v_divg):
            v.trace_add("write", self._collect)
        # 自定义日期一填，月份/日就该灰掉；清空再放回来
        self.v_datecustom.trace_add("write", lambda *_: self._sync_date_row())
        # 底栏右端那句「用哪台机 · 什么通道 · 多大纸」跟着这三样实时变
        for _v in (self.v_w, self.v_h, self.v_mode, self.v_printer):
            _v.trace_add("write", lambda *_: self._update_status_ctx())
        self.txt_fabric.bind("<KeyRelease>", self._collect)
        self.txt_bcaution.bind("<KeyRelease>", self._collect)
        self.v_mode.trace_add("write", lambda *_: self._update_channel_hint())
        # 纸张形态一改：间隙那格跟着灰/亮，底栏尺寸说明也要重算
        self.v_media.trace_add("write", lambda *_: self._sync_media_row())
        self.refresh_printers()
        # 关整窗必须接管：默认行为是直接销毁，多个标签的未保存改动会无声丢掉
        self.root.protocol("WM_DELETE_WINDOW", self._on_app_close)
        # 启动后静默自检一次（窗口先画出来再查，免得开窗就被弹窗挡住）
        self.root.after(900, lambda: self.check_env(False))

    # ---------------- 底栏自适应 ----------------
    def _reflow_bar(self):
        """按按钮栏实际宽度决定：一行并排 / 动作区一行 + 文件组单独一行

        为什么不是简单地让动作区自己折行：那样排出来是
            [打印队列] …… [保存][另存为][导入模板][打开]     ← 第一行
            [批量打印][查看 TSPL 指令]                        ← 第二行
        读取顺序被切开（批量打印跑到"打开"后面去了），看着像排错了。
        真正想要的是：动作区用满整行，文件组落到下一行靠右 ——
            [张数][打印预览][打印][撤销][打印队列][批量打印][查看 TSPL 指令]
                                          [保存][另存为][导入模板][打开]
        这样两行都是"从左往右读"，而且总高度一样（都是两行）。
        """
        flow = getattr(self, "_bar_flow", None)
        grp = getattr(self, "_bar_file_grp", None)
        if flow is None or grp is None or not flow.winfo_exists():
            return
        full = self.bar.winfo_width()
        if full <= 1:
            return
        # 宽度没变就不重排：重排会改高度 → 又触发 Configure → 抖成死循环
        if abs(full - self._bar_last_w) < 2:
            return
        self._bar_last_w = full
        grp_w = grp.winfo_reqwidth() + 18
        need = flow.needed_width()

        if need + grp_w <= full:
            if getattr(self, "_bar_stacked", False):
                grp.pack_configure(side=tk.RIGHT, anchor="n", fill="none")
                self._bar_stacked = False
            flow.reflow(full - grp_w)
        else:
            if not getattr(self, "_bar_stacked", False):
                grp.pack_configure(side=tk.BOTTOM, anchor="e", fill="none")
                self._bar_stacked = True
            flow.reflow(full)
        # 给测试/自检看的总行数：动作区行数 + 文件组是否另占一行
        self._bar_total_rows = getattr(flow, "_rows", 1) + (1 if self._bar_stacked else 0)

    def _on_foot_resize(self, ev):
        """状态行变窄时，先收起那句操作提示 —— 它是说明，不是信息本身。

        不做这件事的话：状态行一行要 1115 px，窗口一窄，pack 只能把任务条
        挤成 0 宽（任务标签凭空消失），而右边那几个状态还占着位置。
        收提示、留信息，比"两边都挤掉一点"更符合实际用法。
        """
        hint = getattr(self, "lbl_taskhint", None)
        if hint is None or not hint.winfo_exists():
            return
        narrow = ev.width < 1040
        shown = bool(hint.winfo_ismapped())
        try:
            if narrow and shown:
                hint.pack_forget()
            elif (not narrow) and (not shown):
                hint.pack(side=tk.LEFT, padx=(14, 0), after=self.taskstrip)
        except tk.TclError:
            pass

    # ---------------- 打印机 ----------------
    def refresh_printers(self):
        """重新读打印机，把「现在用哪台」同步到界面

        只认两个来源：这一格已经选中的 > 软件记忆里「你上次亲手切换」的那台。
        **没有第三级兜底** —— 以前选不到就自动挑「排序第一台」，而排序第一台在冷启动
        （一次都没打过）时正好是 Windows 的系统默认机，等于让一个你没选过的东西替你
        决定往哪台打。现在选不到就明说「没选打印机」，打之前拦住，而不是悄悄打错机器。
        """
        infos = eng.list_printer_info()
        self._printer_infos = infos
        names = [i["name"] for i in infos]
        cur = self.v_printer.get()
        if cur not in names:
            cur = ""                      # 上一台在用、现在没了 → 清空，绝不换一台顶上
        mem = eng.app_default_printer()
        if not cur and mem in names:
            cur = mem
        if cur != self.v_printer.get():
            self.v_printer.set(cur)
        # 记忆里那台在这台电脑上找不到（换机器 / 改名 / USB 没插）→ 界面要说出来
        self._printer_missing = mem if (mem and mem not in names) else ""
        self._sync_printer_widgets()
        if not infos:
            self.status.config(text="未检测到打印机，检查USB后点「切换…」重新读一次")
        self._update_channel_hint()

    def _sync_printer_widgets(self):
        """把当前打印机画到那一行：名字 + 驱动·端口·用过几次"""
        name = self.v_printer.get()
        self.lbl_printer.config(text=name or "（没选打印机）",
                                fg="#1F2329" if name else "#D93026")
        info = next((i for i in getattr(self, "_printer_infos", [])
                     if i["name"] == name), None)
        if info is None:
            self.lbl_printer_info.config(
                text=self._no_printer_hint(),
                foreground="#D93026" if getattr(self, "_printer_missing", "") else "#888")
            return
        bits = [b for b in (info.get("driver"), info.get("port")) if b]
        if info.get("is_default"):
            bits.append("系统默认")
        n = eng.printer_usage().get(name, 0)
        if n:
            bits.append("用过 %d 次" % n)
        self.lbl_printer_info.config(text=" · ".join(bits) or "—", foreground="#888")

    def _no_printer_hint(self):
        """没选中打印机时，下面那行说什么 —— 三种情况给三句话，别让人猜"""
        miss = getattr(self, "_printer_missing", "")
        if miss:
            return "上次用的「%s」现在不在这台电脑上，点「切换…」重新选" % miss
        if eng.use_doc_foreign():
            return "这台电脑上还没选过打印机，点「切换…」选一台"
        return "点「切换…」选一台打印机"

    def choose_printer(self):
        """弹出式选择打印机；选中的那台记成软件默认，下次打开直接用它"""
        infos = eng.list_printer_info()
        if not infos:
            messagebox.showwarning("没有打印机",
                                   "系统里没找到打印机，检查USB连接后再试。")
            return
        self.root.config(cursor="watch")
        try:
            self._printer_infos = infos
            pick = ppick.choose(self.root, self.v_printer.get())
        finally:
            self.root.config(cursor="")
        if not pick or pick == self.v_printer.get():
            self.refresh_printers()
            return
        self.set_printer(pick)

    def set_printer(self, name, remember=True):
        """换打印机：写进当前项目 + （可选）记成软件默认 + 刷新显示"""
        if name and name != self.v_printer.get():
            self.v_printer.set(name)          # trace → _collect，项目标脏并写进快照
        if remember and name:
            eng.set_app_default_printer(name)
        self.refresh_printers()
        self.status.config(text="打印机：%s" % (name or "（未选择）"))

    def _update_channel_hint(self):
        """显示当前打印机会走哪条通道（自动模式下）"""
        name = self.v_printer.get()
        # 只拿一份「假如用这台」的副本去判定通道，不写回项目 ——
        # 写回的话，刚打开文件、还没选打印机时就会把文件里的旧打印机抹掉
        probe = dict(self.proj)
        probe["printer"] = name
        probe["print_mode"] = MODE_NAME.get(self.v_mode.get(), "auto")
        mode = eng.resolve_mode(probe)
        if not name:
            # 没选打印机的时候，这一行留着说「会走图片模式」没意义，还会让人以为已经选好了
            self.lbl_channel.config(text="当前：还没选打印机", foreground="#D93026")
            return
        self.lbl_channel.config(text="当前：%s → %s" % (
            name, "图片模式（通用，不会黑方块）" if mode == "image" else "TSPL 直发"),
            foreground="#2f6f4f")

    # ---------------- 数据同步 ----------------
    def _sync_to_ui(self):
        """界面 ← 项目。自带 _building 保护：回填输入框会触发 trace，
        不挡住的话会把用户刚设好的锚边/间距覆盖掉（踩过这个坑）。"""
        prev = self._building
        self._building = True
        try:
            self._sync_to_ui_impl()
        finally:
            self._building = prev
        # 自动纸长的文件里可能存着一个旧长度（上次存盘后内容又改过、或文件来自
        # 更老的版本），回填完立刻按当前内容重算 —— 否则界面停在一个会打出残次品
        # 的值上，用户看不出异常，直到纸上少了半截才发现。
        try:
            if self.proj.get("auto_paper"):
                old_h = float(self.proj["paper"]["h"])
                self._apply_auto_paper(self.proj)
                if abs(float(self.proj["paper"]["h"]) - old_h) > 1e-6:
                    self.redraw()
                    self._sync_bottom_span()
        except (tk.TclError, KeyError, ValueError):
            pass

    def _sync_to_ui_impl(self):
        p = self.proj
        b = p.get("bottom") or {}
        self.v_w.set(str(p["paper"]["w"]))
        # %g 而不是 str：自动算出来的长度是整数，别显示成「112.0」再跳成「112」
        self.v_h.set("%g" % float(p["paper"]["h"]))
        self.v_autoh.set(bool(p.get("auto_paper", False)))
        try:
            self.ent_h.configure(state="disabled" if p.get("auto_paper") else "normal")
        except (tk.TclError, AttributeError):
            pass
        self.v_gap.set(str(p["paper"]["gap"]))
        self.v_dpi.set(str(p["paper"]["dpi"]))
        self.v_media.set(MEDIA_BY_VALUE.get(
            str(p["paper"].get("media") or "gap"), "间隙纸"))
        self._sync_media_row()
        self.v_title.set(p.get("title", ""))
        self.v_left.set(str(p.get("left_mm", 2.0)))
        self.v_right.set(str(p.get("right_mm", 2.0)))
        self.v_starty.set(str(p.get("start_y_mm", 20.0)))
        self.v_font.set(str(p.get("font_dots", 24)))
        self.v_lg.set(str(p.get("line_gap", 4)))
        # 打印机是「这台电脑接着哪台机器」的机器级设置，**只跟软件记忆走，不跟文件走**：
        # 记忆里那台就是上次你亲手切到的那台。文件里的 printer 字段仍然照写，
        # 但只给批量窗口的「跟文件一样」用 —— 主界面不拿它当兜底，
        # 否则在店里用条码打印做的文件，回到家打开还会去找那台机器。
        self.v_printer.set(eng.app_default_printer())
        self.v_mode.set(MODE_BY_VALUE.get(p.get("print_mode", "auto"), "自动"))
        self.v_offx.set("%.2f" % float(p.get("off_x_mm", 0.0) or 0.0))
        self.v_offy.set("%.2f" % float(p.get("off_y_mm", 0.0) or 0.0))
        self.v_direction.set(str(int(p.get("direction", 0) or 0)))
        self.refresh_printers()
        self.v_toprule.set(bool(p.get("top_rule", {}).get("on", True)))
        self.v_brand.set(bool(p.get("brand", {}).get("on", True)))
        self.v_bottom.set(bool(b.get("on", True)))
        self.v_bgap.set(str(b.get("gap_above_mm", 5.0)))
        self._sync_bottom_labels()
        self._sync_bottom_span()          # 按真实版式回填「实际」值
        self.v_btitle.set(b.get("title", ""))
        self.v_bsub.set(b.get("subtitle", ""))
        self.v_btfont.set(str(int(b.get("title_font", 16) or 16)))
        self.v_bsfont.set(str(int(b.get("sub_font", 16) or 16)))
        self.v_bcfont.set(str(int(b.get("body_font", 16) or 16)))
        self.v_iconth.set(TH_BY_VALUE.get(int(b.get("icon_th", 170)), "中"))
        self.v_islot.set("%g" % float(b.get("icon_slot_mm", 5.0) or 5.0))
        self.v_isgap.set("%g" % float(b.get("icon_slot_gap_mm", 0.6) or 0.0))
        self.v_isrow.set("%g" % float(b.get("icon_row_gap_mm", 1.0) or 1.0))
        self.v_bfoot.set(b.get("footer", ""))
        self.v_bfootfont.set(str(int(b.get("footer_font", 16) or 16)))
        self.v_balign.set(ALIGN_BY_VALUE.get(str(b.get("align", "flow")), "紧跟正文"))
        self.v_origin.set(p.get("origin", ""))
        self.v_center.set(str(p.get("text_align", "left")).lower() == "center")
        self.v_date.set(bool(p.get("date_on", False)))
        _lt = time.localtime()
        # 注意：_i() 收的是 StringVar（内部要 .get()），这里是项目里的裸值，直接转
        try:
            _m = int(p.get("date_month", 0) or 0)
        except (TypeError, ValueError):
            _m = 0
        try:
            _d = int(p.get("date_day", 0) or 0)
        except (TypeError, ValueError):
            _d = 0
        self.v_datemon.set(eng.MONTH_CHOICES[(_m - 1) if 1 <= _m <= 12 else (_lt.tm_mon - 1)])
        self.v_dateday.set(str(_d if 1 <= _d <= 31 else _lt.tm_mday))
        self.v_datefont.set(str(int(p.get("date_font", 12) or 12)))
        self.v_datecustom.set(str(p.get("date_custom", "") or ""))
        self.v_datealign.set(DATE_ALIGN_BY_VALUE.get(
            str(p.get("date_align", "center") or "center"), "居中"))
        self._sync_date_row()          # 没勾印日期时把这一行灰掉
        # 字体：项目里存的路径 → 下拉显示名；不在候选里就当「自定义：文件名」补进去
        _ff = str(p.get("font_file", "") or "")
        _lbl = next((n for n, v in self._font_map.items() if v == _ff), None)
        _warn = ""
        if _lbl is None and _ff:
            _base = os.path.basename(eng.font_expand_path(_ff))
            if eng.font_available(_ff):
                _lbl = "自定义：" + _base
            else:
                # 这台电脑没这个字体：文件是别人电脑上存的。
                # 下拉里明说「本机没有」，但**仍然把路径挂回去**——不然一存盘，
                # 别人精心挑的字体就被这台机器悄悄改掉了。
                _lbl = "自动（缺 %s）" % _base
                _warn = "本机没装 %s，正用雅黑出图。要还原就把它装进系统字体，或点「浏览…」重选" % _base
            self._font_map[_lbl] = _ff
            self.cb_font.configure(values=list(self._font_map))
        self.v_fontsel.set(_lbl or "自动（雅黑 / 宋体 / 黑体）")
        _lw = getattr(self, "lbl_fontwarn", None)
        if _lw is not None and _lw.winfo_exists():
            _lw.config(text=_warn)
        self.v_indent.set("%g" % float(p.get("body_indent_mm", 0.0) or 0.0))
        dv = p.get("divider") or {}
        self.v_div.set(bool(dv.get("on")))
        self.v_divw.set("%g" % float(dv.get("w_mm", 20.0) or 20.0))
        self.v_divg.set("%g" % float(dv.get("gap_mm", 1.25) or 1.25))
        self.v_keepblank.set(bool(p.get("fabric_keep_blank", False)))
        self.txt_fabric.delete("1.0", tk.END)
        self.txt_fabric.insert("1.0", "\n".join(p.get("fabrics", [])))
        self.txt_bcaution.delete("1.0", tk.END)
        self.txt_bcaution.insert("1.0", "\n".join(b.get("caution", [])))
        self._sync_icon_label()
        self._sync_iconfit()

    @staticmethod
    def _f(var, default):
        try:
            return float(str(var.get()).strip())
        except Exception:
            return default

    @staticmethod
    def _i(var, default):
        try:
            return int(float(str(var.get()).strip()))
        except Exception:
            return default

    def _sync_media_row(self):
        """纸张形态变化后收拾界面：连续纸没有缝，「间隙」那一格该灰掉

        为什么灰掉而不是留着可填：连续纸上根本没有缝可以找，填了数字也不会被用到。
        留着能填，只会让人以为「改了间隙就改了长度」，越调越歪。
        """
        cont = MEDIA_NAME.get(self.v_media.get(), "gap") == "continuous"
        try:
            self.ent_gap.configure(state="disabled" if cont else "normal")
        except (tk.TclError, AttributeError):
            pass
        if cont:
            # 提示要短：这一行已经挤了六组控件，长一句就把窗口撑宽
            hint = "无缝隙 · 长度按纸长"
        else:
            hint = "靠缝定位"
        try:
            self.lbl_media_hint.config(text=hint)
        except (tk.TclError, AttributeError):
            pass
        try:
            self._update_paper_line()
        except (tk.TclError, AttributeError):
            pass

    def _on_autoh(self):
        """勾/取消「自动加长纸」：纸长输入框跟着变灰或恢复"""
        on = bool(self.v_autoh.get())
        if not on:
            # 关掉的那一瞬间，输入框里还是自动算出来的长值（比如 114）。
            # 得先把基准值还回去，否则 _collect 会把这个 114 当成「用户新设的手工纸长」，
            # 于是取消勾选回不到 75，而且以后再勾上基准也变成 114 了。
            base = (self.proj.get("paper") or {}).get("h_base")
            if base:
                self.v_h.set("%g" % float(base))
                self.proj["paper"]["h"] = float(base)
        try:
            self.ent_h.configure(state="disabled" if on else "normal")
        except tk.TclError:
            pass
        self._collect()

    def _apply_auto_paper(self, p):
        """按当前内容核一遍纸长。**只碰** paper.h / paper.h_base / 纸长输入框。

        单独拎出来是为了「刚打开文件」时能只重算纸长：那条路上如果直接调
        _collect，会顺带把打印机等字段按界面现状反写回项目，等于悄悄改文件。

        规则：只加长不缩短 —— 内容缩回去时回到手设的基准长度，
        而不是越用越长或越用越短；纸长写回输入框只为显示，
        trace 会再进来一次，算出同一个值就停（不会振荡）。
        """
        if p.get("auto_paper"):
            if not p["paper"].get("h_base"):
                # 第一次勾上：把手设值记成基准（写进文件，下次打开还在）
                p["paper"]["h_base"] = max(1.0, self._f(self.v_h, 75))
            base = float(p["paper"]["h_base"])
            need = eng.page_bottom_mm(p) + AUTO_TAIL_MM
            h_auto = max(base, float(math.ceil(need)))
            p["paper"]["h"] = h_auto
            if self.v_h.get() != ("%g" % h_auto):
                self.v_h.set("%g" % h_auto)
        else:
            # 关掉之后输入框就是唯一权威：用户改多少就是多少，
            # 并把它记成新的基准 —— 下次再勾上自动加长，从新基准往上长。
            # 旧写法是「无条件回填 h_base」，导致取消勾选后手输 120 会被弹回 75。
            manual = max(1.0, self._f(self.v_h, 75))
            p["paper"]["h"] = manual
            p["paper"]["h_base"] = manual

    def _collect(self, *_, dirty=True):
        """把界面内容收进 self.proj。用户改动（dirty=True）才打未保存标记；
        程序性收集（切标签/新建/保存/看指令）不标记。
        注意：只有内容真的变了才标脏 —— 初始化时自动选打印机等动作会触发
        StringVar 的 trace，但内容没变，不该把标签标成「未保存」。"""
        if self._building or getattr(self, "_bsync", False):
            # _bsync：正在把「实际值」回填到输入框，这时不能反写回项目，
            # 否则实际值会覆盖用户设定的预留量，溢出判断变成自己比自己（永远不报或乱报）
            return
        p = self.proj
        before = json.dumps(p, sort_keys=True, ensure_ascii=False) if dirty else None
        p["paper"]["w"] = max(1.0, self._f(self.v_w, 35))
        p["paper"]["h"] = max(1.0, self._f(self.v_h, 75))
        p["paper"]["gap"] = self._f(self.v_gap, 2)
        p["paper"]["dpi"] = self._i(self.v_dpi, 203)
        p["paper"]["media"] = MEDIA_NAME.get(self.v_media.get(), "gap")
        p["title"] = self.v_title.get()
        p["left_mm"] = self._f(self.v_left, 2)
        p["right_mm"] = self._f(self.v_right, 2)
        p["start_y_mm"] = self._f(self.v_starty, 20)
        p["font_dots"] = max(8, min(72, self._i(self.v_font, 24)))
        p["line_gap"] = self._i(self.v_lg, 4)
        # 只有在面板上真的选了才写回项目：否则打开一个文件（还没选打印机）随便改一个字，
        # 就会把文件里存着的打印机抹成空 —— 那是批量窗口「跟文件一样」要用的
        if self.v_printer.get():
            p["printer"] = self.v_printer.get()
        p["print_mode"] = MODE_NAME.get(self.v_mode.get(), "auto")
        # 打印位移 / 方向 —— 只由打印预览窗口驱动
        p["off_x_mm"] = max(-50.0, min(50.0, self._f(self.v_offx, 0.0)))
        p["off_y_mm"] = max(-50.0, min(50.0, self._f(self.v_offy, 0.0)))
        try:
            _ang = int(float(self.v_direction.get() or 0)) % 360
        except Exception:
            _ang = 0
        p["direction"] = _ang if _ang in (0, 90, 180, 270) else 0
        p["top_rule"]["on"] = bool(self.v_toprule.get())
        p["brand"]["on"] = bool(self.v_brand.get())
        # 空行全部留在数据里：开关只管「渲不渲染」，不该顺手把用户敲的空行删掉。
        # 只砍掉末尾的空条目 —— 那是 Text 结尾换行带出来的，不是用户敲的。
        _fab = self.txt_fabric.get("1.0", "end").splitlines()
        while _fab and not _fab[-1].strip():
            _fab.pop()
        p["fabrics"] = _fab
        p["fabric_keep_blank"] = bool(self.v_keepblank.get())
        p["origin"] = self.v_origin.get()
        p["text_align"] = "center" if self.v_center.get() else "left"
        p["date_on"] = bool(self.v_date.get())
        # 月份/日都存成值：选好哪个月哪一天，纸上就是哪个月哪一天，跨天补打也不变。
        # 下拉显示的是「01 JAN」这种带数字注释的，解析交给 engine.month_of，
        # 它同时认老的「NOV」和纯数字「11」，老文件读进来不会变成 0。
        p["date_month"] = eng.month_of(self.v_datemon.get())
        _day = self._i(self.v_dateday, 0)
        p["date_day"] = max(1, min(31, _day or time.localtime().tm_mday))
        p["date_font"] = max(8, min(48, self._i(self.v_datefont, 12)))
        p["date_custom"] = str(self.v_datecustom.get())
        p["date_align"] = DATE_ALIGN_NAME.get(self.v_datealign.get(), "center")
        # 字体：下拉里选的是显示名，存的是路径（系统字体目录缩成 $WIN，换电脑还能找）
        _fsel = self.v_fontsel.get()
        p["font_file"] = self._font_map.get(_fsel, "")
        _lw = getattr(self, "lbl_fontwarn", None)
        if _lw is not None and _lw.winfo_exists() and "缺 " not in _fsel:
            _lw.config(text="")          # 选了别的字体，缺字提示就该消失
        p["body_indent_mm"] = max(0.0, min(30.0, self._f(self.v_indent, 0.0)))
        dv = p.setdefault("divider", {})
        dv["on"] = bool(self.v_div.get())
        dv["w_mm"] = max(1.0, min(35.0, self._f(self.v_divw, 20.0)))
        dv["gap_mm"] = max(0.0, min(20.0, self._f(self.v_divg, 1.25)))
        dv.setdefault("h_mm", 0.2)

        b = p.setdefault("bottom", {})
        b["on"] = bool(self.v_bottom.get())
        # 上方留白是区块唯一的定位来源，输入框里显示的就是设定值，直接读。
        # 有表格时它就是「距洗涤方式」，存在 table 里（表格对话框读同一个键）。
        if eng.table_on(p):
            p.setdefault("table", {})["to_bottom_mm"] = max(0.0, self._f(self.v_bgap, 4.0))
        else:
            b["gap_above_mm"] = max(0.0, self._f(self.v_bgap, 5.0))
        b.setdefault("margin_mm", 5.0)       # 纸底保留线，固定 5mm，不再暴露成可调项
        b["title"] = self.v_btitle.get()
        b["subtitle"] = self.v_bsub.get()
        # 字号：8~72 夹一下（太小打不出笔画、太大一屏装不下），和表格小字同一套口径
        b["title_font"] = max(8, min(72, self._i(self.v_btfont, 16)))
        b["sub_font"] = max(8, min(72, self._i(self.v_bsfont, 16)))
        b["body_font"] = max(8, min(72, self._i(self.v_bcfont, 16)))
        b["icon_th"] = TH_NAME.get(self.v_iconth.get(), 170)
        # 槽位尺寸：夹到能用又不离谱的区间（1~40mm 槽、0~20mm 间距行距）
        b["icon_slot_mm"] = max(1.0, min(40.0, self._f(self.v_islot, 5.0)))
        b["icon_slot_gap_mm"] = max(0.0, min(20.0, self._f(self.v_isgap, 0.6)))
        b["icon_row_gap_mm"] = max(0.0, min(20.0, self._f(self.v_isrow, 1.0)))
        b["caution"] = [s for s in self.txt_bcaution.get("1.0", "end").splitlines() if s.strip()]
        b["footer"] = self.v_bfoot.get()
        b["footer_font"] = max(8, min(72, self._i(self.v_bfootfont, 16)))
        b["align"] = ALIGN_NAME.get(self.v_balign.get(), "flow")
        # 位置模式一改，「上方留白」这一格该不该灰、单位说明写什么，都要跟着变
        self._sync_bottom_labels()

        # ===== 自动纸长 =====
        # 放在最后算：它要看「正文 + 表格 + 洗涤块」全都落定之后的底部。
        p["auto_paper"] = bool(self.v_autoh.get())
        self._apply_auto_paper(p)

        # 有改动 → 当前标签打上未保存标记（只在从干净变脏时刷新一次，避免每次按键重建标签条）
        if dirty and json.dumps(p, sort_keys=True, ensure_ascii=False) != before:
            self._note_undo(before)          # 改动前的样子入撤销栈（连续输入会合并）
            t = self.tabs[self.cur]
            if not t["dirty"]:
                t["dirty"] = True
                self._refresh_tabs()
        self.redraw()
        # 纸长一变，纸底还剩多少就变了 —— 状态行必须跟着走，否则显示的是旧纸长的数
        self._sync_bottom_span()
        self._sync_iconfit()

    # ---------------- 上下留白联动 ----------------
    def _gap_is_table(self):
        """表格打开时，「上方留白」其实等于「表格底 → 洗涤方式顶」，值存在 table 里"""
        return eng.table_on(self.proj)

    def _sync_bottom_labels(self):
        """按表格开关切换这一行的名字和单位说明，避免一个框两种含义却看不出来

        位置模式选了「居中 / 贴纸底」之后，这一格就从设定值变成了结果 ——
        灰掉并写明由谁接管，否则用户改了没反应会以为坏了。"""
        align = ALIGN_NAME.get(self.v_balign.get(), "flow")
        ent = getattr(self, "_ent_bgap", None)
        if align != "flow":
            self._lbl_bgap.config(text="上方留白")
            self._unit_bgap.config(text="mm（由位置模式接管，改这里不起作用）")
            if ent is not None:
                ent.state(["disabled"])
            return
        if ent is not None:
            ent.state(["!disabled"])
        if self._gap_is_table():
            self._lbl_bgap.config(text="距洗涤方式")
            self._unit_bgap.config(text="mm（表格底 → 洗涤方式顶）")
        else:
            self._lbl_bgap.config(text="上方留白")
            self._unit_bgap.config(text="mm（内容底 → 洗涤方式顶）")

    def _sync_date_row(self):
        """日期这一行：不印就整行灰掉，印了才让改。

        自定义内容填了之后，月份/日就不再参与打印（date_text 优先取自定义），
        把那两个控件一起灰掉，免得用户改了半天发现纸上没变化。
        """
        on = bool(self.v_date.get())
        custom = bool(str(self.v_datecustom.get()).strip()) if hasattr(
            self, "v_datecustom") else False
        for w in (self._cb_datemon, self._ent_dateday):
            try:
                w.state(["!disabled"] if (on and not custom) else ["disabled"])
            except tk.TclError:
                pass
        for w in (self._ent_datefont, self._btn_today,
                  getattr(self, "_ent_datecustom", None),
                  getattr(self, "_cb_datealign", None)):
            if w is None:
                continue
            try:
                w.state(["!disabled"] if on else ["disabled"])
            except tk.TclError:
                pass

    def _date_today(self):
        """「今天」——把月份和日一次性拉回当天（字号不动，那是排版选择）"""
        lt = time.localtime()
        self.v_datemon.set(eng.MONTH_CHOICES[lt.tm_mon - 1])
        self.v_dateday.set(str(lt.tm_mday))

    def _pick_font(self):
        """选一个字体文件。先试加载一次 —— 选到坏文件当场说，别等出图才炸"""
        path = filedialog.askopenfilename(
            parent=self.root, title="选字体文件",
            filetypes=[("字体文件", "*.ttf *.ttc *.otf"), ("全部文件", "*.*")])
        if not path:
            return
        store = eng.font_store_path(path)
        if not eng.font_ok(store):
            messagebox.showerror(
                "这个字体用不了",
                "文件读不出字形，可能是坏文件或不支持的格式：\n%s\n\n"
                "换一个 .ttf / .ttc 试试。" % path, parent=self.root)
            return
        # 能加载 ≠ 有中文。Arial / Consolas / Tahoma 都能正常加载，
        # 但标签上的「面料」「洗涤」会全变成方框 —— 出图前拦住，别等打印才发现。
        if not eng.font_cjk_ok(store):
            if not messagebox.askyesno(
                    "这个字体没有中文",
                    "文件能打开，但里面没有中文字形。\n"
                    "用它打印，所有中文会变成方块（□□□）。\n\n"
                    "只用英文的标签可以继续，否则请换「微软雅黑」「黑体」"
                    "这类中文字体。\n\n仍要使用吗？", parent=self.root,
                    icon="warning", default="no"):
                return
        name = "自定义：" + os.path.basename(path)
        self._font_map[name] = store
        self.cb_font.configure(values=list(self._font_map))
        self.v_fontsel.set(name)

    def _pick_icons(self):
        """选洗涤图标：顺序即打印顺序；预览按标签当前的槽位参数画"""
        b = self.proj.get("bottom") or {}
        cur = list(b.get("icons") or [])
        res = ipick.choose(self.root, cur, cfg=b)
        if res is None:
            return
        if not res:
            if not messagebox.askyesno(
                    "一个图标都没有",
                    "这样这一行不会打印任何图标。确定吗？", parent=self.root):
                return
        self.proj.setdefault("bottom", {})["icons"] = list(res)
        self._mark_dirty()
        self.redraw()
        self._sync_icon_label()
        self._sync_iconfit()

    def _sync_iconfit(self):
        """把「几个图标、怎么分排、实际多大」算给用户看。

        为什么要这一行：分排规则（6→3+3、7→3+4）和纸宽够不够，
        光看预览只能凭感觉；写出来才能一眼确认没排错。"""
        lbl = getattr(self, "lbl_iconfit", None)
        if lbl is None or not lbl.winfo_exists():
            return
        g = eng.icon_grid_mm(self.proj)
        if not g["n"]:
            lbl.config(text="还没放图标", foreground="#888")
            return
        rows = " + ".join(str(k) for k in g["rows"])
        if g["shrunk"]:
            nom = float((self.proj.get("bottom") or {}).get("icon_slot_mm", 5.0) or 5.0)
            lbl.config(
                text="⚠ 纸太窄：一排 %d 个 %.1fmm 槽位要 %.1fmm，可用只有 %.1fmm，"
                     "已自动压到 %.1fmm\n%d 个 → %s　整块高 %.1fmm"
                     % (max(g["rows"]), nom, g["need_nominal"], g["region"],
                        g["slot"], g["n"], rows, g["block_h"]),
                foreground="#c62828")
            return
        lbl.config(
            text="%d 个 → %s　槽位 %.1fmm　整块高 %.1fmm"
                 % (g["n"], rows, g["slot"], g["block_h"]),
            foreground="#888")

    def _sync_icon_label(self):
        lbl = getattr(self, "lbl_icons", None)
        if lbl is None or not lbl.winfo_exists():
            return
        icons = list((self.proj.get("bottom") or {}).get("icons") or [])
        miss = [p for p in icons if not os.path.exists(eng.asset_path(p))]
        txt = "当前 %d 个" % len(icons)
        if miss:
            txt += " · ⚠ %d 个文件找不到" % len(miss)
        lbl.config(text=txt, foreground="#c62828" if miss else "#888")

    def _refresh_gap_row(self, p=None):
        """表格开关或表格参数变了之后，刷新「上方留白」这一行（名字、值、实际值）"""
        if (p if p is not None else self.proj) is not self.proj:
            return
        self._sync_bottom_labels()
        self._sync_bottom_span()

    def _sync_bottom_span(self):
        """把真实结果回填到输入框（必要时），并刷新状态行。

        为什么必须回填：区块只有一个位置，两个留白不可能同时自由，
        所以模型的真实值永远是「最后被编辑的那个说了算」。
        但回填要克制 —— 只在「输入的数字和模型算的不一致」时才改，
        典型就是负数被夹成 0。用户正打到「1.」「abc」这种半截状态时不动它，
        否则每敲一个字符就把光标位置重写一遍，框里的字会跳。
        """
        if self._bsync:
            return
        p = self.proj
        b = p.get("bottom") or {}
        if not b.get("on", True):
            self.lbl_bspan.config(text="洗涤方式已关闭")
            return
        top, gap, tail = eng.bottom_span_mm(p)
        raw = str(self.v_bgap.get()).strip()
        try:
            typed = float(raw)
        except (TypeError, ValueError):
            typed = None            # 半截输入 / 空 / 字母，交给模型走默认值，不回写
        # 表格打开时，输入框显示的是表格那个键的真实值（两者是同一个量）
        if typed is not None and abs(typed - gap) > 0.05:
            self._bsync = True
            try:
                self.v_bgap.set("%.1f" % gap)
            finally:
                self._bsync = False
        warn = ""
        if eng.content_overflow(p):
            warn = "　⚠ 已挤到纸底，建议纸长 %.1f mm" % eng.suggested_h_mm(p)
        if tail >= 0:
            span = "纸底还剩 %.1f mm" % tail
        else:
            span = "已超出纸底 %.1f mm" % (-tail)
        wide = self._bottom_wide(p)
        if wide:
            # 字号手输以后，「一行够不够宽」不再是有保障的事：超了就该当场说，
            # 而不是等打出来发现右边被切掉（纸宽固定，字宽不会自己缩）
            warn += "　⚠ %s比纸宽（字号 %.0f 太大）" % (wide[0], wide[1])
        self.lbl_bspan.config(text="区块顶 %.1f mm · %s%s" % (top, span, warn),
                              foreground=("#C62828" if warn else "#888"))

    def _bottom_wide(self, p):
        """底部区块里有没有一行横着放不下。返回 (段落名, 字号) 或 None。

        只报第一处 —— 一起报会把这行字撑爆，反而不如改一处看一眼。
        """
        b = p.get("bottom") or {}
        if not b.get("on", True):
            return None
        try:
            dpi = int(p["paper"].get("dpi", 203))
            region = eng.mm2dot(max(4.0, float(p["paper"]["w"])
                                    - float(p.get("left_mm", 2.0))
                                    - float(p.get("right_mm", 2.0))), dpi)
        except (KeyError, TypeError, ValueError):
            return None
        rows = [("大标题", b.get("title", ""), b.get("title_font", 16)),
                ("英文副标题", b.get("subtitle", ""), b.get("sub_font", 16))]
        for i, s in enumerate(b.get("caution") or [], 1):
            rows.append(("注意事项第 %d 行" % i, s, b.get("body_font", 16)))
        for name, s, f in rows:
            s = str(s).strip()
            if not s:
                continue
            try:
                if eng.txt_w_dots(s, int(f)) > region:
                    return (name, int(f))
            except (TypeError, ValueError):
                continue
        return None

    def _on_bgap_edit(self):
        """用户改上方留白：值由 _collect 那条 trace 写进项目并重绘，这里只刷新状态行"""
        if self._building or self._bsync:
            return
        self._sync_bottom_span()

    # ---------------- 预览 ----------------
    def redraw(self):
        c = self.canvas
        c.delete("all")
        self._photos = []
        W, H = c.winfo_width(), c.winfo_height()
        if W < 30 or H < 30:
            return
        p = self.proj
        w_mm = float(p["paper"]["w"])
        h_mm = float(p["paper"]["h"])
        dpi = float(p["paper"]["dpi"]) or 203.0
        mm_px = dpi / 25.4            # 1mm = 多少 dot

        pad = 26
        scale = min((W - 2 * pad) / w_mm, (H - 2 * pad) / h_mm)
        pw, ph = w_mm * scale, h_mm * scale
        x0, y0 = (W - pw) / 2, (H - ph) / 2

        # 纸张
        c.create_rectangle(x0, y0, x0 + pw, y0 + ph, fill="white", outline="#9aa0a6", width=1)

        # ①②③④ 统一绘制清单 —— 与 TSPL 直发、PIL 出图共用同一份，
        # 三处版式永远一致；画布只是把这些坐标换算成屏幕像素。
        self._paint_ops(c, eng.build_ops(p), x0, y0, scale, dpi, w_mm)

        b = p.get("bottom") or {}
        over = b.get("on", True) and eng.content_overflow(p)

        # 标注
        c.create_text(x0, y0 - 10, text="%g × %g mm   %d dpi" % (w_mm, h_mm, int(dpi)),
                      anchor="sw", fill="#555", font=(FONT, -11))
        sug = eng.suggested_h_mm(p)
        # 纸底这一行只留一行字：内容挤到保留线时整行变红，把「怎么回事 + 怎么办」写在一起。
        # 以前另起一行红字提示，纸一长就压到这一行上，反而看不清。
        tip = "可打印 %d × %d dot" % (eng.mm2dot(w_mm, dpi), eng.mm2dot(h_mm, dpi))
        color = "#888"
        if b.get("on", True):
            if p.get("auto_paper"):
                # 纸长本来就是程序按内容算好的，再报一个「建议值」只会让人
                # 以为哪里偏了（而且纸长框是灰的，改也改不了）
                tip += "　纸长由内容决定（底部留 %g mm 余量）" % AUTO_TAIL_MM
            elif over:
                tip = "⚠ 超过纸底保留线　" + tip + "　建议纸长 %.1f mm" % sug
                color = "#d32f2f"
            elif str(b.get("align", "flow")) in ("bottom", "center"):
                # 贴底/居中时，「建议纸长」是按「紧跟正文」算的 —— 照着改会算错，
                # 纸一缩短洗涤块就被挤回正文上。这儿说实话：位置由模式接管。
                tip += "　底部留白 %.1f mm（位置：%s）" % (
                    eng.bottom_span_mm(p)[2],
                    ALIGN_BY_VALUE.get(str(b.get("align", "flow")), "紧跟正文"))
            else:
                tip += "　建议纸长 %.1f mm%s" % (
                    sug, "（当前偏长，底部空白大）" if sug < h_mm - 0.5 else "")
        c.create_text(x0 + pw, y0 + ph + 10, text=tip,
                      anchor="ne", fill=color, font=(FONT, -10))

    def _paint_ops(self, c, ops, x0, y0, scale, dpi, w_mm):
        """把引擎的绘制清单画到画布上：dot → mm → 屏幕像素。
        几何全部取自引擎，界面和打印不会各画一套。"""
        dpmm = 25.4 / dpi                       # 1 dot = 多少 mm
        for op in ops:
            k = op["kind"]
            if k == "rect" and op["mode"] == "center":
                rw = op["w_mm"]
                rx = x0 + (w_mm - rw) / 2.0 * scale
                ry = y0 + op["y_mm"] * scale
                rh = max(1.2, op["h_mm"] * scale)
                c.create_rectangle(rx, ry, rx + rw * scale, ry + rh,
                                   fill="#000000", outline="")
            elif k == "rect":
                ox = x0 + op["x"] * dpmm * scale
                oy = y0 + (op.get("base_mm", 0.0) + op["y"] * dpmm) * scale
                bw = max(1.0, op["w"] * dpmm * scale)
                bh = max(1.0, op["h"] * dpmm * scale)
                c.create_rectangle(ox, oy, ox + bw, oy + bh,
                                   fill="#000000", outline="")
            elif k == "text" and op["mode"] == "left":
                fpx = max(7, int(round(op["font"] * dpmm * scale)))
                c.create_text(x0 + op["x_mm"] * scale, y0 + op["y_mm"] * scale,
                              text=op["text"], anchor="nw", fill="#111111",
                              font=(FONT, -fpx))
            elif k == "text" and op["mode"] == "mid":
                fpx = max(6, int(round(op["font"] * dpmm * scale)))
                cx = x0 + (op["x"] + eng.txt_w_dots(op["text"], op["font"]) / 2.0) \
                    * dpmm * scale
                cy = y0 + (op["base_mm"] + op["y"] * dpmm) * scale
                c.create_text(cx, cy, text=op["text"], anchor="n",
                              fill="#111111", font=(FONT, -fpx))
            elif k == "text":                       # cell：按字形外框居中
                # 用 anchor="center" 而不是 "n" —— 后者按字体行框顶部对齐，
                # 行框比字形高一大截，文字会被推向格子底部
                fpx = max(6, int(round(op["font"] * dpmm * scale)))
                tw = eng.txt_w_dots(op["text"], op["font"]) * dpmm * scale
                f_px = op["font"] * dpmm * scale
                ox = x0 + op["x"] * dpmm * scale
                oy = y0 + (op["base_mm"] + op["y"] * dpmm) * scale
                c.create_text(ox + tw / 2.0, oy + f_px / 2.0,
                              text=op["text"], anchor="center",
                              fill="#111111", font=(FONT, -fpx))
            else:                                   # image
                try:
                    from PIL import Image, ImageTk
                    img, wd, hd = eng.prep_icon(op["path"], op["w"], op["th"])
                    if img is None:
                        continue
                    iw = max(2, int(round(wd * dpmm * scale)))
                    ih = max(2, int(round(hd * dpmm * scale)))
                    thumb = img.convert("L").resize((iw, ih),
                                                    Image.Resampling.NEAREST)
                    photo = ImageTk.PhotoImage(thumb)
                    self._photos.append(photo)
                    c.create_image(
                        x0 + op["x"] * dpmm * scale,
                        y0 + (op["base_mm"] + op["y"] * dpmm) * scale,
                        image=photo, anchor="nw")
                except Exception:
                    pass

    def _make_scroll_cmd(self, canvas, vbar):
        """给 Canvas 的 yscrollcommand：转发给滚动条，但内容没超出时藏起来。

        Tk 会频繁调这个回调（每次滚动、每次内容变化），所以必须记住
        「当前显示与否」，只在状态真的翻转时 pack/pack_forget，否则疯狂重排。
        """
        def _cmd(*frac):
            if len(frac) == 2:
                top, bottom = float(frac[0]), float(frac[1])
            else:
                top, bottom = 0.0, 1.0
            need = (bottom - top) < 0.999
            self._sync_scrollbar(canvas, vbar, need)
            vbar.set(top, bottom)
        return _cmd

    def _sync_scrollbar(self, canvas, vbar, need=None):
        """按内容是否超出，显示 / 隐藏滚动条。"""
        try:
            if not canvas.winfo_exists():
                return
            if need is None:
                sr = canvas.cget("scrollregion")
                if not sr:
                    need = False
                else:
                    x1, y1, x2, y2 = [float(v) for v in sr.split()]
                    need = (y2 - y1) > canvas.winfo_height() + 1
            shown = bool(vbar.winfo_ismapped())
            if need and not shown:
                vbar.pack(side=tk.RIGHT, fill=tk.Y)
                self._after_sb_visibility()
            elif not need and shown:
                vbar.pack_forget()
                self._after_sb_visibility()
        except (tk.TclError, ValueError):
            pass

    def _after_sb_visibility(self):
        """滚动条显隐后：重排页签按钮（避开滚动条）+ 重算滚动区域。

        用 after_idle 而不是直接调 —— 直接调时 pack 还没完成布局，
        winfo_width() 拿到旧值，按钮让位宽度会算错。
        """
        try:
            self.root.after_idle(self._layout_right_buttons)
            self.root.after_idle(self._refresh_scroll_areas)
        except tk.TclError:
            pass

    def _refresh_scroll_areas(self):
        """窗口尺寸变化后重算所有表单页的滚动区域 —— 滚动条该出现就出现。

        只在真的有滚动 Canvas 时干活，且用 after 合并短时间内的多次
        <Configure>，避免拖动窗口时每帧重算几十个控件。
        """
        if not getattr(self, "_scroll_canvases", None):
            return
        if getattr(self, "_scroll_job", None):
            return
        self._scroll_job = self.root.after(80, self._do_refresh_scroll)

    def _do_refresh_scroll(self):
        self._scroll_job = None
        for cv in list(getattr(self, "_scroll_canvases", [])):
            try:
                if not cv.winfo_exists():
                    continue
                cv.configure(scrollregion=cv.bbox("all"))
                # 找到这个 Canvas 配对的滚动条，同步显隐
                vb = getattr(self, "_scroll_bars", {}).get(str(cv))
                if vb is not None:
                    self._sync_scrollbar(cv, vb)
            except tk.TclError:
                pass

    def _scroll_page(self, nb, title):
        """造一个带垂直滚动条的表单页，返回**内层内容 frame**。

        为什么不能直接给 Notebook 加滚动：Notebook 的页面就是普通 frame，
        内容多出来时它自己不会滚。经典做法是 Canvas + Scrollbar + 内层 frame：
          · 外层 outer 作为 Notebook 真正的页（nb.add 的对象）；
          · 内层 inner 才是装控件的地方，函数把它返回给调用方当 t1/t2 用。

        三个必须处理好的细节（少一个都会出怪毛病）：
          ① scrollregion 要在内容变化时重算，否则滚动条长度不对、滚不到底；
          ② inner 宽度要跟着 Canvas 变，否则 fill=X 的输入框只占内容自然宽度；
          ③ 滚轮只在鼠标停在本页时生效 —— 不能 bind_all，那会连左边预览
             和别的弹窗一起滚。用 enter/leave 记账，再按平台绑不同滚轮事件。
        """
        outer = ttk.Frame(nb)
        canvas = tk.Canvas(outer, highlightthickness=0, bd=0)
        vbar = ttk.Scrollbar(outer, orient=tk.VERTICAL, command=canvas.yview)
        inner = ttk.Frame(canvas, padding=8)
        win_id = canvas.create_window((0, 0), window=inner, anchor="nw")

        # ★ 滚动条不常驻：内容没超出时不 pack，省下 17px 宽度给表单，
        #   也避免它压住右上角那排工具按钮。只有真的超长了才显示。
        canvas.configure(yscrollcommand=self._make_scroll_cmd(canvas, vbar))
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        def _on_inner(e):
            self._sync_scrollbar(canvas, vbar)

        def _on_canvas(e):
            # 内层宽度跟随可视宽度，输入框 / 表格才能 fill=X 铺满
            if e.width > 20:
                canvas.itemconfig(win_id, width=e.width)
            self._sync_scrollbar(canvas, vbar)

        inner.bind("<Configure>", _on_inner)
        canvas.bind("<Configure>", _on_canvas)
        inner.after(30, lambda: self._sync_scrollbar(canvas, vbar))

        # 滚轮：进页面才接管，离开就还回去。Windows 用 <MouseWheel>；
        # 顺手支持 X11 的 4/5（在这个项目用不到，但跨平台不踩坑）。
        state = {"in": False}

        def _wheel(e):
            if not state["in"]:
                return
            d = -1 if e.delta > 0 else 1
            canvas.yview_scroll(d, "units")
            return "break"

        def _wheel_x11(e):
            if not state["in"]:
                return
            canvas.yview_scroll(-1 if e.num == 4 else 1, "units")
            return "break"

        for w in (outer, canvas, inner):
            w.bind("<Enter>", lambda e: state.__setitem__("in", True), add="+")
            w.bind("<Leave>", lambda e: state.__setitem__("in", False), add="+")
        canvas.bind("<MouseWheel>", _wheel, add="+")
        inner.bind("<MouseWheel>", _wheel, add="+")
        canvas.bind("<Button-4>", _wheel_x11, add="+")
        canvas.bind("<Button-5>", _wheel_x11, add="+")
        inner.bind("<Button-4>", _wheel_x11, add="+")
        inner.bind("<Button-5>", _wheel_x11, add="+")

        self._scroll_canvases.append(canvas)
        if not hasattr(self, "_scroll_bars"):
            self._scroll_bars = {}
        self._scroll_bars[str(canvas)] = vbar
        nb.add(outer, text=title)
        return inner

    # ---------------- 动作 ----------------
    def open_print_preview(self):
        """打开打印预览。单例：重复点就把已开的窗口提到前面并重绘"""
        self._collect(dirty=False)
        w = getattr(self, "_pv_win", None)
        if w is not None and w.winfo_exists():
            w.deiconify()
            w.lift()
            w.focus_force()
            w._render()
            return
        PrintPreview(self)

    def _layout_right_buttons(self):
        """让右上角工具组避开滚动条。

        1.9.1 起「表格工具」和「自检页」都在同一个 tools 里 pack，
        不再各自算坐标 —— 宽度由 pack 自己算，永远不会互相压页签。
        这里只剩一件事：滚动条显示时把整个工具组往左让出它的宽度，
        否则滚动条会盖住工具组最右那个按钮。
        """
        tools = getattr(self, "_right_tools", None)
        if tools is None:
            return
        bars = getattr(self, "_scroll_bars", {})
        sb = next(iter(bars.values()), None)
        sbw = sb.winfo_width() + 2 if (sb is not None and sb.winfo_ismapped()) else 0
        try:
            tools.place_configure(relx=1.0, x=-(4 + sbw), y=6, anchor="ne")
        except tk.TclError:
            pass


    # ---------------- 运行环境检查 ----------------
    def env_problems(self):
        """「换到别人电脑上会不会用不了」这几件事，能查的都查一遍。

        为什么值得单独做一个检查：这几种毛病**都不报错**，只是悄悄不出东西 ——
        没中文字体 → 整张标签空白；目录不可写 → 存盘失败 + 崩溃恢复失效且无声；
        图标文件丢了 → 那一行空一块。用户遇到只会觉得「这软件坏了」。"""
        bad = []
        # ① 中文字体：没有的话渲染直接返回 None，一个字都不画，也不抛异常
        try:
            f = eng.pick_ttf()
        except Exception:
            f = None
        if not f:
            bad.append("找不到任何中文字体：标签会打成空白（不是方块，是一个字都没有）。\n"
                       "       解决：装个中文语言包，或把 .ttf 字体拷进系统字体目录。")
        # ② 程序目录可写：保存、自动恢复、复制图标都靠它
        try:
            tp = os.path.join(APP_DIR, ".wtest_%d" % os.getpid())
            with open(tp, "w", encoding="utf-8") as fh:
                fh.write("ok")
            os.remove(tp)
        except Exception:
            bad.append("程序所在的文件夹不能写：保存、自动恢复、复制图标都会失败。\n"
                       "       现在的位置：%s\n"
                       "       解决：把整个文件夹挪到桌面或 D/E 盘（别放 Program Files）。"
                       % APP_DIR)
        # ③ 图标文件在不在（项目里记的是相对路径，assets 被挪走就断）
        icons = list((self.proj.get("bottom") or {}).get("icons") or [])
        miss = [x for x in icons if not os.path.exists(eng.asset_path(x))]
        if miss:
            bad.append("有 %d 个洗涤图标文件找不到（共 %d 个）：那一行会缺图。\n"
                       "       解决：把图标补回 assets\\icons\\，或用「选洗涤图标…」重选。"
                       % (len(miss), len(icons)))
        return bad

    def check_env(self, verbose=False):
        """verbose=False 是启动时静默跑：没问题就一句话都不说，别天天弹窗；
        verbose=True 是用户点右下角「环境检查」，好坏都给个回执。"""
        try:
            bad = self.env_problems()
        except Exception as e:
            if verbose:
                messagebox.showinfo("环境检查", "检查本身出错了：%s" % e,
                                    parent=self.root)
            return
        try:
            self.lbl_env.config(
                text=("环境 %d 项待处理" % len(bad)) if bad else "环境检查",
                foreground="#c62828" if bad else "#8a94a0")
        except Exception:
            pass
        if not bad:
            if verbose:
                messagebox.showinfo(
                    "环境正常",
                    "这台电脑跑得动：\n"
                    "  · 中文字体：%s\n"
                    "  · 程序目录可写：是\n"
                    "  · 洗涤图标：齐全"
                    % os.path.basename(eng.pick_ttf() or "-"),
                    parent=self.root)
            return
        messagebox.showwarning(
            "这台电脑的环境有点问题",
            "程序能打开，但下面这些会影响出图和保存：\n\n"
            + "\n\n".join(bad)
            + "\n\n（右下角「环境检查」可以随时再查一遍）",
            parent=self.root)

    def _update_status_ctx(self):
        """状态条右边那句「用哪台机 / 什么通道 / 多大纸」"""
        try:
            p = self.proj
            paper = p.get("paper") or {}
            name = (self.v_printer.get() or "").strip() or "未选打印机"
            mode = "图片模式" if eng.resolve_mode(p) == "image" else "TSPL 直发"
            # 纸的形态必须写在状态条上：连续纸和间隙纸打出来的行为完全不同，
            # 光看「35 × 75 mm」分不出这份文件到底按哪种打。
            media = MEDIA_BY_VALUE.get(str(paper.get("media") or "gap"), "间隙纸")
            self.lbl_ctx.config(text="%s　·　%s　·　%s　·　%g × %g mm"
                                % (name, mode, media, paper.get("w", 35), paper.get("h", 75)))
        except Exception:
            pass

    def do_selftest(self):
        """打一张自检页：换纸、换机、怀疑打歪/字糊时，先打这一张再上真标签。

        内容是量出来的，不是看出来的 —— 外框四角 L 标看偏移，1mm 刻度尺量比例，
        25mm 标准线直接拿尺子比，字号阶梯看最小几号还认得出，灰阶看黑度够不够。
        """
        self._collect(dirty=False)
        printer = (self.v_printer.get() or "").strip()
        if not printer:
            messagebox.showwarning(
                "还没选打印机",
                "先在「打印机」那一栏选一台，再打自检页。",
                parent=self.root)
            return
        p = dict(self.proj)
        paper = p.get("paper") or {}
        w = float(paper.get("w") or 35.0)
        h = float(paper.get("h") or 75.0)
        gap = float(paper.get("gap") or 2.0)
        dpi = int(paper.get("dpi") or 203)
        mode = eng.resolve_mode(p)
        chan = "TSPL 直发" if mode == "tspl" else "图片模式"
        if not messagebox.askokcancel(
                "打一张自检页",
                "打印机：%s\n通道：%s\n纸型：%g × %g mm\n\n"
                "会真的出一张纸。上面有：\n"
                "  · 外框 + 四角 L 标 —— 看有没有打歪、缩放\n"
                "  · 1mm 刻度尺 + 25mm 标准线 —— 拿尺子直接比\n"
                "  · 字号阶梯 3.0→1.0mm —— 看最小几号还认得出\n"
                "  · 灰阶块 —— 看黑度、有没有糊成一片"
                % (printer, chan, w, h),
                parent=self.root):
            return
        self.status.config(text="正在打自检页…", foreground="#666")
        info = {"printer": printer,
                "paper": "%g×%gmm" % (w, h),
                "dpi": "%d dpi" % dpi,
                "mode": chan,
                "app": "WaterLabel"}
        self._st_result = None
        self._st_pending = True

        def work():
            # 注意：不能在这里调 root.after() —— 工作线程碰 Tcl 解释器会抛
            # RuntimeError: main thread is not in main loop。跟打印队列一样，
            # 线程只把结果放在属性里，由主线程的 _poll_queue 来取。
            try:
                n, used = eng.print_selftest(printer, mode, w, h, gap, info)
                # 状态行别写字节数：那一行一长就把整个窗口撑宽（实测 1300→1700）
                self._st_result = (True, used)
            except Exception as e:            # noqa: BLE001 —— 打不出去的原因五花八门
                self._st_result = (False, str(e))

        threading.Thread(target=work, daemon=True).start()

    def _take_selftest_result(self):
        """主线程取自检页结果；没有就返回 None"""
        if not getattr(self, "_st_pending", False):
            return None
        res = getattr(self, "_st_result", None)
        if res is None:
            return None
        self._st_pending = False
        self._st_result = None
        return res

    def _selftest_done(self, res):
        good, txt = res
        try:
            self.status.config(
                text=("自检页已发送：" + txt) if good else ("自检页没打出去：" + txt),
                foreground="#2f6f4f" if good else "#c00")
        except tk.TclError:
            return
        if not good:
            messagebox.showerror("自检页没打出去", txt, parent=self.root)

    def do_print(self):
        """把当前内容作为一个任务加入队列，不打断正在打印的任务"""
        self._collect(dirty=False)
        # 以面板上「现在选的是哪台」为准，不看项目里存着的那台 ——
        # 面板没选就是没选，绝不能拿文件里的旧打印机替用户打出去
        name = self.v_printer.get()
        p = json.loads(json.dumps(self.proj))
        p["printer"] = name
        if not name:
            messagebox.showwarning("没有打印机",
                                   "还没选打印机。点右上「打印机」那格或「切换…」选一台。")
            return
        if not any(str(f).strip() for f in p.get("fabrics", [])) and not p.get("title"):
            messagebox.showwarning("内容为空", "号型和成分都是空的，没什么可打。")
            return
        # 最后一道闸：内容超纸就打不全，而且打出来才发现（整批废纸）。
        # 这里必须问一句 —— 提示藏在小字里等于没提示。
        over = eng.page_overflow_mm(p)
        side = eng.page_side_overflow_mm(p)
        if over > 0 or side > 0:
            bits = []
            if over > 0:
                bits.append("内容比标签长 %.1f mm，下半截会被裁掉" % over)
            if side > 0:
                bits.append("正文比可印宽度多 %.1f mm，右边会被裁掉" % side)
            if not messagebox.askyesno(
                    "内容超出标签",
                    "；\n".join(bits) + "\n\n"
                    "建议：纸长改到 %.1f mm，或减少成分行数 / 缩小字号 / 改窄边距。\n\n"
                    "仍然要打这一张吗？" % eng.suggested_h_mm(p)):
                return
        # 图片模式下先确认驱动真的会按这份文件的尺寸设纸。
        # 不确认就开打，等一整批打完才发现长度不对，纸就废了 ——
        # 这是「每个文件尺寸都不一样」最容易翻车的地方。
        if eng.resolve_mode(p) == "image":
            try:
                rep = eng.check_paper_apply(name, p["paper"]["w"], p["paper"]["h"])
            except Exception:
                rep = None
            if rep and rep.get("devmode") and not rep.get("ok"):
                gw, gh = rep["got"]
                if not messagebox.askyesno(
                        "驱动不接受这个尺寸",
                        "这份文件要 %.1f × %.1f mm，但驱动只肯给 %.1f × %.1f mm。\n\n"
                        "这么打会左右偏、下半截被裁。两个办法：\n"
                        "  1) 打印方式切「TSPL 直发」—— 按指令里的 SIZE 走，不受驱动纸型限制\n"
                        "  2) 把纸宽/纸长改成驱动支持的值\n\n"
                        "仍然要打这一张吗？"
                        % (p["paper"]["w"], p["paper"]["h"], gw, gh)):
                    return
        try:
            copies = int(str(self.v_copies.get()).strip())
        except Exception:
            copies = 1
        copies = max(1, min(999, copies))
        self.v_copies.set(str(copies))
        self.enqueue_snapshot(p, copies)

    def enqueue_snapshot(self, proj_snapshot, copies=1):
        """把一份项目快照排进打印队列。打印预览走这条，主界面也走这条"""
        p = json.loads(json.dumps(proj_snapshot))
        mode = eng.resolve_mode(p)
        # 记一次使用：打印机列表按这个次数排序，常用的那台自动浮到最上面。
        # 放在这里而不是「在列表里点一下」—— 点开看看不算用，否则每次启动自动选中的那台会被刷成第一名。
        eng.note_printer_use(p.get("printer", ""))
        with self.qlock:
            self.task_seq += 1
            task = {
                "id": self.task_seq,
                "name": "%s #%d" % (p.get("name") or "水洗标", self.task_seq),
                "proj": p,                            # 快照：后面改界面不影响已入队任务
                "copies": copies,
                "printer": p.get("printer", ""),
                "mode": mode,
                "status": "排队中",
                "prog": "0/%d" % copies,
                "err": "",
            }
            self.tasks.append(task)
            # 只有「闲着」才自动开工。暂停中（失败停的 / 终止停的 / 单步停的）
            # 绝不因为来了新任务就自己跑起来 —— 用户点了张单张打印，
            # 结果上一批打失败的十几个跟着一起出纸，那就是一叠废纸。
            if self.qstate == "idle":
                self.qstate = "running"
            still_paused = self.qstate == "paused"
            pend = sum(1 for t in self.tasks if t["status"] == "排队中")
        self.qevt.set()                     # 叫醒队列线程（它平时在睡）
        chan = "图片模式" if mode == "image" else "TSPL 直发"
        if still_paused:
            self.status.config(text="已加入队列：%d 张（%s）｜ 队列暂停中，"
                                    "剩 %d 个 —— 点「继续」" % (copies, chan, pend))
        else:
            self.status.config(text="已加入队列：%d 张（%s）" % (copies, chan))
        return task

    # ---------------- 打印队列 ----------------
    def _take_next_task(self):
        """取一个待打任务并标记「打印中」；没有就返回 None"""
        with self.qlock:
            if self.qstate != "running":
                return None
            for t in self.tasks:
                if t["status"] == "排队中":
                    t["status"] = "打印中"
                    return t
        return None

    def _queue_worker(self):
        """后台线程：一次打一个任务，打完取下一个；终止后暂停等「继续」。

        事件驱动 —— 没活时阻塞在 qevt.wait()，不占 CPU，也不做空转轮询。
        """
        while not self.qstop.is_set():
            if not self.qevt.wait(0.5):
                continue                    # 超时醒来只是为了看一眼 qstop
            self.qevt.clear()
            while not self.qstop.is_set():
                task = self._take_next_task()
                if task is None:
                    break
                self.abort_flag = False
                ok, err = self._run_task(task)
                with self.qlock:
                    if ok:
                        task["status"] = "已完成"
                        task["prog"] = "%d/%d" % (task["copies"], task["copies"])
                    elif err == "ABORT":
                        task["status"] = "已终止"
                    else:
                        task["status"] = "失败"
                        task["err"] = err
                    if self.abort_flag and self.qstate == "running":
                        self.qstate = "paused"      # 终止后自动暂停，等用户点「继续」
                        self.q_paused_by = "abort"
                    elif not ok and self.qstate == "running" and any(
                            x["status"] == "排队中" for x in self.tasks):
                        # 失败就地停住，不再往下打。
                        # 打不出来基本都是环境问题（打印机选错、没开机、缺纸、端口被占），
                        # 接着往下打只会把剩下的一起打成失败 —— 越打越乱，还得逐个重来。
                        self.qstate = "paused"
                        self.q_paused_by = "fail"
                    elif (self.batch_step and self.qstate == "running"
                          and any(x["status"] == "排队中" for x in self.tasks)):
                        # 批量单步：这个打完了，后面还有活 —— 停下来等人点「继续」。
                        # 不新开线程，就靠 paused 让 _take_next_task() 取不到下一个。
                        self.qstate = "paused"
                        self.q_paused_by = "step"

    def _run_task(self, task):
        """执行并跟踪一个打印任务。"""
        project = task["proj"]
        copies = int(task["copies"])
        sent = 0

        def cancelled():
            return self.abort_flag

        def submitted(count):
            nonlocal sent
            sent += count
            task["prog"] = f"{sent}/{copies}"

        if cancelled():
            return False, "ABORT"
        if not (project.get("printer") or "").strip():
            return False, "没选打印机（点「切换…」选一台）"

        try:
            if task["mode"] == "image":
                eng.gdi_print(project, copies, cancelled=cancelled, submitted=submitted)
                report = eng.last_paper_report()
                if report.get("devmode") and not report.get("ok"):
                    w, h = report["want"]
                    dw, dh = report["got"]
                    task["warn"] = f"纸尺寸没被驱动接受：要 {w:.1f}×{h:.1f}，实际 {dw:.1f}×{dh:.1f} mm"
                elif report.get("feed_note"):
                    task["warn"] = report["feed_note"]
            else:
                eng.print_tspl_copies(project, copies, doc_name=task["name"],
                                      cancelled=cancelled, submitted=submitted)
            return True, ""
        except eng.PrintCancelled:
            eng.purge_printer(task["printer"])
            return False, "ABORT"
        except Exception as error:
            return False, f"已提交 {sent}/{copies} 张。{error}" if sent else str(error)

    def _poll_queue(self):
        """每 250ms 刷新主窗口状态和队列窗口"""
        try:
            st = self._take_selftest_result()
            if st is not None:
                self._selftest_done(st)
            with self.qlock:
                pend = sum(1 for t in self.tasks if t["status"] == "排队中")
                cur = [t for t in self.tasks if t["status"] == "打印中"]
                last = None                      # 最近一个已了结的任务，用来看上一步成没成
                for t in reversed(self.tasks):
                    if t["status"] in ("已完成", "已终止", "失败"):
                        last = t
                        break
            fail_last = last is not None and last["status"] == "失败"
            # 打完了但纸尺寸没被驱动接受 —— 不算失败，可成品一定是歪的，
            # 必须在这一行说出来，否则「已完成」会让人以为没问题
            warn_last = (last.get("warn") if last is not None else "") or ""
            aborting = bool(self.abort_flag and cur)     # 已经点了终止，还在等驱动返回
            if cur:
                t = cur[0]
                if aborting:
                    # 驱动卡住时终止要等这次调用返回才生效（实测卡 3 秒就等 3 秒）。
                    # 不写一句「正在终止」，用户会以为按钮坏了，接着猛点或直接强杀。
                    txt = "正在终止 %s —— 等打印机返回…" % t["name"]
                else:
                    txt = "正在打印 %s（%s）｜等待 %d" % (t["name"], t["prog"], pend)
            elif pend:
                if self.qstate != "paused":
                    txt = "排队中 %d 个任务" % pend
                elif self.q_paused_by == "fail":
                    # 打不出来停下的 —— 必须说清楚「是失败停的」，
                    # 否则用户以为只是普通暂停，直接点继续又白打一次。
                    txt = "有任务打印失败，队列已停 ｜ 剩 %d 个" % pend
                elif self.batch_step:
                    # 文字要短：这一行和「查看 TSPL」挤在一起，太长会把窗口撑宽
                    txt = "批量暂停 · 待继续 %d 个" % pend
                else:
                    txt = "队列已暂停，等待 %d 个任务 —— 点「打印队列」继续" % pend
            else:
                # 队列空了也要能看出「上一次是打失败了」，否则只显示「队列空闲」，
                # 用户以为打完了，纸其实没出来。
                if fail_last:
                    txt = "上次打印失败：%s" % (last.get("err") or "")[:10]
                elif warn_last:
                    txt = "⚠ " + warn_last
                else:
                    txt = "队列空闲"
            fg = "#c00" if (aborting or (self.qstate == "paused" and pend)
                            or (not pend and (fail_last or warn_last))) else "#666"
            if (txt, fg) != self._qstat_last:      # 文字没变就别碰控件，省掉无谓重绘
                self.qstat.config(text=txt, foreground=fg)
                self._qstat_last = (txt, fg)
            if self.qwin is not None and self.qwin.winfo_exists():
                self._refresh_qtree()
            self._refresh_taskstrip()
            self._sync_batch_resume()
            if self.batch_win is not None and self.batch_win.win.winfo_exists():
                self.batch_win.refresh()
        finally:
            self.root.after(250, self._poll_queue)

    def _sync_batch_resume(self):
        """队列停下时，在主界面放一个「继续」入口。

        理由：队列窗口／批量窗口可能被盖住或最小化，看不到「在等你点继续」
        就会以为软件卡死了。失败停下也要给入口 —— 否则修好打印机都找不到出口。
        """
        with self.qlock:
            need = bool(self.qstate == "paused"
                        and any(t["status"] == "排队中" for t in self.tasks))
            why = self.q_paused_by
            pend = sum(1 for t in self.tasks if t["status"] == "排队中")
        label = "▶ 继续批量" if self.batch_step and why == "step" else "▶ 继续打印"
        if need and self.btn_bresume is None:
            self.btn_bresume = ttk.Button(self.bar, text=label,
                                          command=self.q_resume)
            self.btn_bresume.pack(side=tk.LEFT, padx=(6, 0))
            # 这里不再往状态行写「等继续 / 剩几个」。
            # 原来写了一句「批量单步：等继续（剩 N 个）」，可左边 qstat 已经在说
            # 「批量暂停 · 待继续 N 个」，右边还排着「▶ 继续批量」按钮 ——
            # 三者讲同一件事。状态行以后只管「我刚做了什么」，队列归 qstat，各说各的。
        elif not need and self.btn_bresume is not None:
            self.btn_bresume.destroy()
            self.btn_bresume = None
            self._qstat_last = None          # 控件少了一个，状态文字重新算一次

    def open_batch(self):
        """打开批量逐步打印窗口（只允许一个，再点就抬起来）"""
        if self.batch_win is not None and self.batch_win.win.winfo_exists():
            self.batch_win.win.deiconify()
            self.batch_win.win.lift()
            return
        self.batch_win = bprint.BatchPrintWindow(self)

    def open_queue(self):
        if self.qwin is not None and self.qwin.winfo_exists():
            self.qwin.deiconify()
            self.qwin.lift()
            return
        w = tk.Toplevel(self.root)
        w.title("打印队列")
        w.geometry("800x400")
        self.qwin = w

        cols = ("id", "name", "printer", "mode", "copies", "status", "err")
        self.qtree = ttk.Treeview(w, columns=cols, show="headings", height=12,
                                  selectmode="browse")
        for key, txt, wid, anch in (("id", "#", 34, "center"),
                                     ("name", "任务", 130, "w"),
                                     ("printer", "打印机", 150, "w"),
                                     ("mode", "方式", 60, "center"),
                                     ("copies", "张数", 50, "center"),
                                     ("status", "状态", 130, "center"),
                                     ("err", "备注", 190, "w")):
            self.qtree.heading(key, text=txt)
            self.qtree.column(key, width=wid, anchor=anch)
        self.qtree.pack(fill=tk.BOTH, expand=True, padx=8, pady=(8, 4))

        qb = ttk.Frame(w, padding=(8, 4, 8, 8))
        qb.pack(fill=tk.X)
        self.btn_resume = ttk.Button(qb, text="▶ 继 续", command=self.q_resume)
        self.btn_resume.pack(side=tk.LEFT)
        self.btn_abort = ttk.Button(qb, text="⛔ 终止当前", command=self.q_abort)
        self.btn_abort.pack(side=tk.LEFT, padx=(8, 0))
        self.btn_del = ttk.Button(qb, text="删除选中", command=self.q_del)
        self.btn_del.pack(side=tk.LEFT, padx=(8, 0))
        self.btn_clean = ttk.Button(qb, text="清空已结束", command=self.q_clean)
        self.btn_clean.pack(side=tk.RIGHT)
        ttk.Label(qb, text="终止会自动暂停队列；「继续」从下一个任务接着打",
                  foreground="#888").pack(side=tk.LEFT, padx=(14, 0))
        self._refresh_qtree()

    def _refresh_qtree(self):
        """重建队列表格（行数很少，重建最简单可靠），并更新按钮可用性"""
        w = self.qwin
        if w is None or not w.winfo_exists():
            return
        if not (hasattr(self, "qtree") and hasattr(self, "btn_resume")):
            return   # 窗口刚建、控件还没建完的那一瞬间
        keep = self.qtree.selection()
        for iid in self.qtree.get_children():
            self.qtree.delete(iid)
        with self.qlock:
            rows = [(t["id"], t["name"], t["printer"],
                     "图片" if t["mode"] == "image" else "TSPL",
                     t["copies"],
                     t["status"] + ("（%s）" % t["prog"] if t["status"] == "打印中" else ""),
                     (t["err"] or "")[:60]) for t in self.tasks]
            printing = any(t["status"] == "打印中" for t in self.tasks)
            paused = self.qstate == "paused"
            has_done = any(t["status"] in ("已完成", "已终止", "失败") for t in self.tasks)
            sel_q = any(t["status"] == "排队中" for t in self.tasks)
        for r in rows:
            self.qtree.insert("", "end", iid=str(r[0]), values=r)
        for iid in keep:
            if self.qtree.exists(iid):
                self.qtree.selection_set(iid)
        self.btn_resume.config(state=tk.NORMAL if paused else tk.DISABLED)
        self.btn_abort.config(state=tk.NORMAL if printing else tk.DISABLED)
        self.btn_del.config(state=tk.NORMAL if sel_q else tk.DISABLED)
        self.btn_clean.config(state=tk.NORMAL if has_done else tk.DISABLED)

    def q_resume(self):
        with self.qlock:
            self.qstate = "running"
            self.q_paused_by = ""
        self.qevt.set()                     # 叫醒线程继续取任务

    def q_abort(self):
        self.abort_flag = True

    def q_del(self):
        sel = self.qtree.selection()
        if not sel:
            return
        tid = int(sel[0])
        with self.qlock:
            for t in list(self.tasks):
                if t["id"] == tid and t["status"] == "排队中":
                    self.tasks.remove(t)
                    break

    def q_clean(self):
        with self.qlock:
            self.tasks = [t for t in self.tasks
                          if t["status"] not in ("已完成", "已终止", "失败")]

    # ---------------- 主界面任务条 ----------------
    TASK_STYLE = {
        "排队中": ("#eceff1", "#455a64"),
        "打印中": ("#bbdefb", "#0d47a1"),
        "已完成": ("#c8e6c9", "#1b5e20"),
        "已终止": ("#ffe0b2", "#e65100"),
        "失败":   ("#ffcdd2", "#b71c1c"),
    }

    def _refresh_taskstrip(self):
        """任务条只在内容变化时重建，避免每 250ms 闪烁"""
        with self.qlock:
            tasks = list(self.tasks)
            sig = "|".join("%s:%s:%s" % (t["id"], t["status"], t["prog"])
                           for t in tasks) + "@" + self.qstate
        if sig == getattr(self, "_task_sig", None):
            return
        self._task_sig = sig
        for c in self.taskstrip.winfo_children():
            c.destroy()
        if not tasks:
            tk.Label(self.taskstrip, text="暂无任务 —— 点「打 印」把当前内容加入队列",
                     fg="#999").pack(side=tk.LEFT)
            return
        # 最多显示最近 6 个：这一行右边还排着状态 + 队列 + 打印机三段，
        # 一排标签按 130px 一个算，6 个刚好，再多就把右边挤出去了。
        show = tasks[-6:]
        if len(tasks) > 10:
            more = tk.Label(self.taskstrip, text="…共 %d 个" % len(tasks),
                            fg="#666", cursor="hand2")
            more.pack(side=tk.LEFT, padx=(0, 6))
            more.bind("<Button-1>", lambda e: self._jump_task(tasks[0]["id"]))
        for t in show:
            bg, fg = self.TASK_STYLE.get(t["status"], ("#eceff1", "#333333"))
            txt = t["name"]      # name 里已经带了「#序号」，别再叠一层
            if t["status"] == "打印中":
                txt += " %s" % t["prog"]
            elif t["status"] in ("已终止", "失败") and t["err"]:
                txt += "（%s）" % t["err"][:12]
            chip = tk.Label(self.taskstrip, text=txt, bg=bg, fg=fg,
                            padx=8, pady=2, cursor="hand2")
            chip.pack(side=tk.LEFT, padx=(0, 6))
            tid = t["id"]
            chip.bind("<Button-1>", lambda e, i=tid: self._jump_task(i))
            chip.bind("<Button-3>", lambda e, i=tid: self._load_task(i))

    def _jump_task(self, tid):
        """单击：打开/聚焦队列窗口并选中该任务"""
        self.open_queue()
        iid = str(tid)
        if self.qtree.exists(iid):
            self.qtree.selection_set(iid)
            self.qtree.see(iid)

    def _load_task(self, tid):
        """右键：把任务快照载入标签页，方便改了重打（不影响当前编辑内容）

        同一个任务只开一个标签：已经开过就直接切过去。
        以前每次都新建，连点几下就变成一串同名标签，看着像在「不断复制文件」。
        """
        with self.qlock:
            t = next((x for x in self.tasks if x["id"] == tid), None)
        if t is None:
            return
        for i, tab in enumerate(self.tabs):
            if tab.get("from_task") == tid:
                if i != self.cur:
                    self._switch_tab(i)
                self.status.config(text="任务 #%d 的标签已在前台" % tid)
                return
        proj = eng.normalize(json.loads(json.dumps(t["proj"])))
        self._new_tab(proj, None, "任务 #%d 已在新标签页打开" % tid, from_task=tid)

    def show_tspl(self):
        self._collect(dirty=False)
        win = tk.Toplevel(self.root)
        win.title("TSPL 指令")
        win.geometry("620x460")
        t = tk.Text(win, wrap="none", font=("Consolas", 9))
        t.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        try:
            txt = eng.build_tspl(self.proj)
        except Exception as e:
            # 90°/270° 时指令集里没有对应方向，引擎会主动抛错。
            # 这里必须接住，否则窗口是空白 + 控制台报错，用户完全不知道发生了什么。
            t.insert("1.0", "当前设置无法生成 TSPL 指令。\n\n原因：%s\n\n"
                            "说明：TSPL 指令集只有 0° 和 180° 两个方向，"
                            "90°/270° 横打只能走「图片模式」。\n"
                            "把方向改回 0°，或在打印方式里选「图片模式」，即可正常打印。"
                            % e)
            t.config(state="disabled")
            return
        # BITMAP 的十六进制太长，折叠显示，方便看坐标
        out = []
        for ln in txt.splitlines():
            if ln.startswith("BITMAP"):
                head, _, data = ln.rpartition(",")
                out.append("%s,<%d 字节十六进制已折叠>" % (head, len(data) // 2))
            else:
                out.append(ln)
        t.insert("1.0", "\n".join(out))
        t.config(state="disabled")

    # ---------------- 表格 ----------------
    def _mark_dirty(self):
        """把当前标签标成有未保存改动"""
        t = self.tabs[self.cur]
        if not t["dirty"]:
            t["dirty"] = True
            self._refresh_tabs()

    def open_table_editor(self):
        """表格编辑窗口：加行 / 改单元格文字 / 调字号行高，改动立刻反映到预览"""
        self._collect(dirty=False)
        t = self.proj.setdefault("table", {})
        # 打开前到底有没有「真表格」：默认项目自带一行空壳（on=False），
        # 那一行不算 —— 否则第一次点就显示「修改（1 行）」，明明还没建过。
        _rows_prev = t.get("rows") or []
        _used = bool(t.get("on")) or any(
            any(str(c).strip() for c in (r.get("cells") or [])) for r in _rows_prev)
        _had = len(_rows_prev) if _used else 0
        t.setdefault("rows", [{"cols": 4, "font": 16, "cells": ["", "", "", ""]}])
        t.setdefault("row_h_mm", 4.0)
        t.setdefault("gap_top_mm", 2.0)
        t.setdefault("to_bottom_mm", 4.0)
        t.setdefault("border_dot", 2)
        t.setdefault("caption", "充绒克数(g)")
        t.setdefault("caption_font", 16)
        t["on"] = True
        self._mark_dirty()
        self._refresh_gap_row()
        self.redraw()

        w = getattr(self, "_twin", None)
        if w is not None and w.winfo_exists():
            w.lift()
            w.focus_force()
            return

        win = tk.Toplevel(self.root)
        self._twin = win
        self._tproj = self.proj
        win.title("表格工具  —  " + ("修改（%d 行）" % _had if _had else "新建"))
        win.transient(self.root)

        # ---- 参数行 ----
        par = ttk.Frame(win, padding=(12, 10, 12, 2))
        par.pack(fill=tk.X)
        self.v_rowh = tk.StringVar(value=str(t["row_h_mm"]))
        self.v_gapt = tk.StringVar(value=str(t["gap_top_mm"]))
        self.v_gapb = tk.StringVar(value=str(t.get("to_bottom_mm", 4.0)))
        self.v_bwid = tk.StringVar(value=str(t["border_dot"]))
        self.v_cap = tk.StringVar(value=str(t.get("caption", "")))
        self.v_capf = tk.StringVar(value=str(t.get("caption_font", 16)))

        def param(pr, col, label, var, unit, values=None):
            ttk.Label(par, text=label).grid(row=pr, column=col * 3,
                                            sticky="e", padx=(0 if col == 0 else 14, 4),
                                            pady=(0 if pr == 0 else 8, 0))
            if values:
                wd = ttk.Combobox(par, textvariable=var, values=values,
                                  width=4, state="normal")   # 可下拉、也可直接输入
            else:
                wd = ttk.Entry(par, textvariable=var, width=6, justify="right")
            wd.grid(row=pr, column=col * 3 + 1, sticky="w",
                    pady=(0 if pr == 0 else 8, 0))
            ttk.Label(par, text=unit, foreground="#666").grid(
                row=pr, column=col * 3 + 2, sticky="w", padx=(3, 0),
                pady=(0 if pr == 0 else 8, 0))
            var.trace_add("write", lambda *_: self._apply_table_params())

        param(0, 0, "行高", self.v_rowh, "mm")
        param(0, 1, "上留白", self.v_gapt, "mm")
        param(0, 2, "距洗涤方式", self.v_gapb, "mm")     # 表格底边 → 洗涤方式标题的间距
        param(0, 3, "框线", self.v_bwid, "dot", ["1", "2", "3"])

        # 第二行：表格上方的小字（如「充绒表（g）」）+ 它的字号
        # 做成可下拉也可手输：常用的几个点一下就填上，别的照样手打。
        # 预设只留充绒这一路的 —— 「衣长/胸围/肩宽/袖长/下摆」是围度表的叫法，
        # 现在围度表直接用「按尺码生成 + 数值行」，那几条列出来只是干扰。
        ttk.Label(par, text="表格上方小字").grid(row=1, column=0, columnspan=2,
                                                sticky="e", padx=(0, 4), pady=(8, 0))
        ttk.Combobox(par, textvariable=self.v_cap, width=16, state="normal",
                     values=["充绒量(g)", "充绒表(g)", "充绒克数(g)"]).grid(
            row=1, column=2, columnspan=3, sticky="w", pady=(8, 0))
        ttk.Label(par, text="字号").grid(row=1, column=5, sticky="e",
                                       padx=(14, 4), pady=(8, 0))
        ttk.Entry(par, textvariable=self.v_capf, width=5, justify="right").grid(
            row=1, column=6, sticky="w", pady=(8, 0))
        ttk.Label(par, text="dot", foreground="#666").grid(
            row=1, column=7, sticky="w", padx=(3, 0), pady=(8, 0))
        self.v_cap.trace_add("write", lambda *_: self._apply_table_params())
        self.v_capf.trace_add("write", lambda *_: self._apply_table_params())

        ttk.Label(win, text="每格文字可直接改；字号任意输入（TSPL 非标准档位会自动转位图打印）",
                  foreground="#777").pack(anchor="w", padx=12, pady=(4, 0))

        # ---- 按尺码生成（尺码行 + 空白行，多组时共用一套列宽网格）----
        sbox = ttk.LabelFrame(win, text=" 按尺码生成表格 ", padding=(10, 6))
        sbox.pack(fill=tk.X, padx=12, pady=(8, 0))
        self.v_sizes = {}
        picks = [str(s) for s in (t.get("size_picks") or [])]
        for i, s in enumerate(SIZE_ORDER):
            v = tk.BooleanVar(value=(s in picks))
            self.v_sizes[s] = v
            # 两行摆，一行 11 个会把窗口撑到 600 多像素
            ttk.Checkbutton(sbox, text=s, variable=v,
                            command=self._size_hint).grid(
                row=i // 6, column=i % 6, sticky="w", padx=(0, 6), pady=1)

        self.lbl_sizehint = ttk.Label(sbox, text="", foreground="#666")
        self.lbl_sizehint.grid(row=2, column=0, columnspan=6, sticky="w",
                               pady=(6, 2))

        # 尺码行与数值行分开给高度：尺码两三个字符 3mm 就够；
        # 数值行以前默认 4.5（留给手写），充绒表这种「克数直接打印」的用不着那么高，
        # 3.5 就够而且更省纸 —— 想手写照样能改回 4.5。
        self.v_szh1 = tk.StringVar(value="3.0")
        self.v_szh2 = tk.StringVar(value="3.5")

        # 数值行的一次性录入：充绒表每个尺码的克数都不一样，一格一格点太费事，
        # 直接照打印顺序把一串数敲进来（空格 / 逗号 / 顿号都认）
        valrow = ttk.Frame(sbox)
        valrow.grid(row=3, column=0, columnspan=6, sticky="ew", pady=(2, 0))
        ttk.Label(valrow, text="数值").pack(side=tk.LEFT)
        self.v_vals = tk.StringVar()
        ttk.Entry(valrow, textvariable=self.v_vals, width=28).pack(
            side=tk.LEFT, padx=(6, 0))
        ttk.Label(valrow, text="例：120 135 150 170（按尺码顺序，空格隔开）",
                  foreground="#888").pack(side=tk.LEFT, padx=(6, 0))
        self.v_vals.trace_add("write", lambda *_: self._size_hint())

        hrow = ttk.Frame(sbox)
        hrow.grid(row=4, column=0, columnspan=6, sticky="ew", pady=(4, 0))
        ttk.Label(hrow, text="行高").pack(side=tk.LEFT)
        ttk.Label(hrow, text="尺码", foreground="#666").pack(side=tk.LEFT, padx=(6, 0))
        ttk.Entry(hrow, textvariable=self.v_szh1, width=4,
                  justify="right").pack(side=tk.LEFT, padx=(4, 0))
        ttk.Label(hrow, text="mm　数值", foreground="#666").pack(side=tk.LEFT, padx=(2, 0))
        ttk.Entry(hrow, textvariable=self.v_szh2, width=4,
                  justify="right").pack(side=tk.LEFT, padx=(4, 0))
        ttk.Label(hrow, text="mm", foreground="#666").pack(side=tk.LEFT, padx=(4, 0))
        self.btn_gen = ttk.Button(hrow, text="生成表格", command=self._gen_size_table)
        self.btn_gen.pack(side=tk.RIGHT)
        self.v_szh1.trace_add("write", lambda *_: self._size_hint())
        self.v_szh2.trace_add("write", lambda *_: self._size_hint())

        ttk.Label(sbox, text="生成的表格就是普通行 —— 每格照样能改字、改字号，也能再手动加行",
                  foreground="#888").grid(row=5, column=0, columnspan=6,
                                          sticky="w", pady=(6, 0))

        # ---- 可滚动的行区 ----
        wrap = ttk.Frame(win, padding=(12, 6))
        wrap.pack(fill=tk.BOTH, expand=True)
        cvs = tk.Canvas(wrap, highlightthickness=0, height=230)
        sb = ttk.Scrollbar(wrap, orient="vertical", command=cvs.yview)
        inner = ttk.Frame(cvs)
        cvs.create_window((0, 0), window=inner, anchor="nw")
        cvs.configure(yscrollcommand=sb.set)
        cvs.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        inner.bind("<Configure>", lambda e: cvs.configure(scrollregion=cvs.bbox("all")))
        self._tinner = inner

        # ---- 底部按钮 ----
        bar = ttk.Frame(win, padding=(12, 6, 12, 10))
        bar.pack(fill=tk.X)
        ttk.Button(bar, text="＋ 一行三列",
                   command=lambda: self._add_table_row(3)).pack(side=tk.LEFT)
        ttk.Button(bar, text="＋ 一行四列",
                   command=lambda: self._add_table_row(4)).pack(side=tk.LEFT, padx=6)
        ttk.Button(bar, text="＋ 一行五列",
                   command=lambda: self._add_table_row(5)).pack(side=tk.LEFT)
        ttk.Button(bar, text="删除表格",
                   command=self._drop_table).pack(side=tk.LEFT, padx=(20, 0))
        ttk.Button(bar, text="关闭", command=win.destroy).pack(side=tk.RIGHT)
        self.lbl_tableinfo = ttk.Label(bar, text="", foreground="#555")
        self.lbl_tableinfo.pack(side=tk.LEFT, padx=14)

        self._rebuild_table_rows()
        self._size_hint()
        win.update_idletasks()
        win.geometry("+%d+%d" % (self.root.winfo_rootx() + 80,
                                 self.root.winfo_rooty() + 130))

    def _size_selected(self):
        """勾了哪些尺码 —— 按 SIZE_ORDER 的顺序返回，与勾选先后无关"""
        vs = getattr(self, "v_sizes", {}) or {}
        return [s for s in SIZE_ORDER if vs.get(s) is not None and vs[s].get()]

    def _parse_vals(self):
        """把「120 135 150 170」拆成一串。空格、逗号、顿号、斜杠、分号都当分隔符 ——
        从 Excel 里复制过来通常是制表符隔开，从聊天里粘过来常是空格或顿号，
        一律认，省得为了个分隔符回头改。"""
        try:
            s = str(self.v_vals.get())
        except (AttributeError, tk.TclError):
            return []
        for ch in (",", "，", "、", "/", "／", ";", "；", "|", "\t", "\n"):
            s = s.replace(ch, " ")
        return [x for x in s.split() if x]

    def _size_hint(self):
        """一行说完：勾了几个、排成几+几、几行、总高多少。

        刻意写得短 —— 这块地方被别的控件挤着，长句子会顶出去。
        「每格多宽」「数值对没对上」这类能从勾选框和输入框自己看出来的，就不说了。"""
        lbl = getattr(self, "lbl_sizehint", None)
        btn = getattr(self, "btn_gen", None)
        picks = self._size_selected()
        if btn is not None and btn.winfo_exists():
            btn.state(["!disabled"] if picks else ["disabled"])
        if lbl is None or not lbl.winfo_exists():
            return
        if not picks:
            # 一个都没勾时整行收起来 —— 留个空标签会在勾选框和下面输入框之间
            # 留出一条莫名其妙的空白
            lbl.config(text="")
            lbl.grid_remove()
            return
        lbl.grid()
        plan, grid = eng.size_table_plan(len(picks))
        p = getattr(self, "_tproj", None) or self.proj
        # 总高 = 排数 × (尺码行高 + 数值行高) + 上方小字那一行
        # 小字也算高度：漏掉它提示会比实际少 2mm，用户按提示判断放不下就判断错了
        try:
            hh = (float(self.v_szh1.get()) + float(self.v_szh2.get())) * len(plan)
            t = p.get("table") or {}
            if str(t.get("caption", "")).strip():
                dpi = float(p["paper"].get("dpi", 203)) or 203.0
                hh += float(t.get("caption_font", 16)) * 25.4 / dpi \
                    + float(t.get("caption_gap_mm", 0.5))
            hh_txt = " · %.1f mm" % hh
        except (TypeError, ValueError, AttributeError, KeyError):
            hh_txt = ""
        # 数值给多了是真错（多出来的会被丢掉），这种必须说
        vals = self._parse_vals()
        bad = "　⚠ 数值多 %d 个" % (len(vals) - len(picks)) \
            if len(vals) > len(picks) else ""
        lbl.config(text="%d 个 → %s · %d 行%s%s"
                        % (len(picks), "+".join(str(x) for x in plan),
                           len(plan) * 2, hh_txt, bad),
                   foreground="#B26A00" if bad else "#2e7d32")

    def _table_has_content(self):
        """表格里有没有已经填过字 —— 有就不许一声不响地覆盖掉"""
        for r in ((getattr(self, "_tproj", None) or self.proj).get("table") or {}).get("rows") or []:
            if any(str(c).strip() for c in (r.get("cells") or [])):
                return True
        return False

    def _gen_size_table(self):
        """把勾选的尺码摆成「尺码行 + 数值行」的表格（排法见 eng.size_table_plan）"""
        picks = self._size_selected()
        if not picks:
            messagebox.showinfo("还没勾尺码",
                                "先在左边勾几个尺码，再点生成。", parent=self._twin)
            return
        if self._table_has_content():
            if not messagebox.askyesno(
                    "会覆盖现有表格",
                    "表格里已经有内容了，生成会把它整个替换掉。\n\n继续吗？",
                    parent=self._twin):
                return
        try:
            h_title = max(2.0, min(20.0, float(self.v_szh1.get())))
        except (TypeError, ValueError, AttributeError):
            h_title = 3.0
        try:
            h_blank = max(2.0, min(20.0, float(self.v_szh2.get())))
        except (TypeError, ValueError, AttributeError):
            h_blank = 3.5
        rows = eng.size_table_rows(picks, self._parse_vals(), h_title, h_blank)
        t = self._tproj.setdefault("table", {})
        t["rows"] = rows
        t["size_picks"] = picks          # 记住勾选，下次打开还在（但不会自动重排）
        t["on"] = True
        self._mark_dirty()
        self._rebuild_table_rows()
        if self._tproj is self.proj:
            self._refresh_gap_row(self.proj)
            self.redraw()
        self.status.config(text="已按 %d 个尺码生成 %d 行表格" % (len(picks), len(rows)))

    def _rebuild_table_rows(self):
        inner = getattr(self, "_tinner", None)
        if inner is None or not inner.winfo_exists():
            return
        for w in inner.winfo_children():
            w.destroy()
        tbl = (self._tproj.get("table") or {})
        rows = tbl.get("rows") or []
        if not rows:
            ttk.Label(inner, text="还没有行 —— 点下面「＋ 一行三列 / 四列 / 五列」",
                      foreground="#777").pack(pady=12)
            self._table_info()
            return
        for i, r in enumerate(rows):
            f = ttk.Frame(inner)
            f.pack(fill=tk.X, pady=2)
            n = int(r.get("cols", 4))
            g = int(r.get("grid", n))
            # 有网格时说明「几格 / 网格列数」，否则会让人以为少画了两列
            head = ("第 %d 行 · %d 列" % (i + 1, n) if g == n
                    else "第 %d 行 · %d 格(%d列网格)" % (i + 1, n, g))
            ttk.Label(f, text=head, width=20).pack(side=tk.LEFT)
            fv = tk.StringVar(value=str(r.get("font", 16)))
            ttk.Combobox(f, textvariable=fv, values=["8", "12", "16", "20", "24", "32", "40"],
                         width=4, state="normal").pack(side=tk.LEFT, padx=(0, 8))
            fv.trace_add("write", lambda *_, i=i, v=fv: self._set_row_font(i, v.get()))
            # 每行自己的高度：尺码行给 3mm、空白行给 4.5mm，这是真实吊牌的做法
            hv = tk.StringVar(
                value="%g" % float(r.get("row_h_mm", tbl.get("row_h_mm", 4.5))))
            ttk.Label(f, text="高", foreground="#666").pack(side=tk.LEFT)
            ttk.Entry(f, textvariable=hv, width=4,
                      justify="right").pack(side=tk.LEFT, padx=(2, 8))
            hv.trace_add("write", lambda *_, i=i, v=hv: self._set_row_h(i, v.get()))
            cells = list(r.get("cells") or []) + [""] * n
            for j in range(n):
                sv = tk.StringVar(value=cells[j])
                # 列越多格子越窄，否则 5 列会把对话框撑得比屏幕还宽
                ew = 18 if n == 3 else (13 if n == 4 else 10)
                e = ttk.Entry(f, textvariable=sv, width=ew)
                e.pack(side=tk.LEFT, padx=1)
                sv.trace_add("write", lambda *_, i=i, j=j, v=sv: self._set_cell(i, j, v.get()))
            ttk.Button(f, text="✕", width=3,
                       command=lambda i=i: self._del_table_row(i)).pack(side=tk.LEFT, padx=(6, 0))
        self._table_info()

    def _set_cell(self, i, j, val):
        rows = (self._tproj.get("table") or {}).get("rows") or []
        if i >= len(rows):
            return
        cells = rows[i].setdefault("cells", [])
        while len(cells) <= j:
            cells.append("")
        cells[j] = val
        self._mark_dirty()
        if self._tproj is self.proj:
            self.redraw()

    def _set_row_h(self, i, val):
        """改某一行的行高。写不进去（半截输入 / 字母）就保持原样，等打完再说"""
        rows = (self._tproj.get("table") or {}).get("rows") or []
        if i >= len(rows):
            return
        try:
            h = float(str(val).strip())
        except (TypeError, ValueError):
            return
        h = max(2.0, min(20.0, h))
        if abs(h - float(rows[i].get("row_h_mm", -1) or -1)) < 1e-9:
            return
        rows[i]["row_h_mm"] = h
        self._mark_dirty()
        if self._tproj is self.proj:
            self._refresh_gap_row(self.proj)
            self.redraw()
        self._table_info()

    def _set_row_font(self, i, val):
        rows = (self._tproj.get("table") or {}).get("rows") or []
        if i >= len(rows):
            return
        try:
            rows[i]["font"] = int(val)
        except Exception:
            return
        self._mark_dirty()
        if self._tproj is self.proj:
            self.redraw()

    def _add_table_row(self, cols):
        t = self._tproj.setdefault("table", {})
        t.setdefault("rows", []).append(
            {"cols": cols, "font": 16, "cells": [""] * cols})
        self._mark_dirty()
        self._rebuild_table_rows()
        if self._tproj is self.proj:
            self.redraw()

    def _del_table_row(self, i):
        rows = (self._tproj.get("table") or {}).get("rows") or []
        if i < len(rows):
            rows.pop(i)
        self._mark_dirty()
        self._rebuild_table_rows()
        if self._tproj is self.proj:
            self.redraw()

    def _drop_table(self):
        p = getattr(self, "_tproj", None) or self.proj
        (p.get("table") or {})["on"] = False
        self._mark_dirty()
        w = getattr(self, "_twin", None)
        if w is not None and w.winfo_exists():
            w.destroy()
        if p is self.proj:
            self._refresh_gap_row(p)
            self.redraw()
        self.status.config(text="已删除表格")

    def _apply_table_params(self):
        p = getattr(self, "_tproj", None) or self.proj
        t = p.setdefault("table", {})
        for key, var, cast in (("row_h_mm", self.v_rowh, float),
                               ("gap_top_mm", self.v_gapt, float),
                               ("to_bottom_mm", self.v_gapb, float),
                               ("border_dot", self.v_bwid, int)):
            try:
                t[key] = cast(float(str(var.get()).strip()))
            except Exception:
                pass
        # 表格上方小字 + 字号（字号任意输入，越界自动夹到 8~72）
        t["caption"] = str(self.v_cap.get())
        try:
            t["caption_font"] = max(8, min(72, int(round(float(str(self.v_capf.get()).strip())))))
        except Exception:
            pass
        self._mark_dirty()
        if p is self.proj:
            self._refresh_gap_row(p)
            self.redraw()
        self._table_info()

    def _table_info(self):
        lbl = getattr(self, "lbl_tableinfo", None)
        if lbl is None or not lbl.winfo_exists():
            return
        p = getattr(self, "_tproj", None) or self.proj
        rows = eng.table_rows(p)
        if not rows:
            lbl.config(text="表格为空", foreground="#555")
            return
        h_mm = float(p["paper"]["h"])
        b = p.get("bottom") or {}
        block_h = eng.bottom_height_mm(p) if b.get("on", True) else 0.0
        limit = h_mm - float(b.get("margin_mm", 5.0)) - block_h   # 洗涤方式顶边最多到这
        gap = eng.table_gap_below(p)
        need = (eng.table_height_mm(p)
                + float((p.get("table") or {}).get("gap_top_mm", 2.0)) + gap)
        avail = limit - eng.text_bottom_mm(p)
        if need > avail + 1e-6:
            lbl.config(text="%d 行 · 超 %.1f mm，洗涤方式会被挤出纸外"
                            % (len(rows), need - avail),
                       foreground="#d32f2f")
        else:
            lbl.config(text="%d 行 · 表格底 %.1f ＋间距 %.1f → 顶 %.1f / 限 %.1f"
                            % (len(rows), eng.table_bottom_mm(p), gap,
                               eng.bottom_top_mm(p), limit),
                       foreground="#2e7d32")

    # ---------------- 多项目标签页 ----------------
    def _tab_disp(self, tab):
        """标签上的显示名：有文件用文件名，没文件用项目名/未命名；未保存加 ●

        不再截断 —— 名字长了交给 _wrap_cjk 换行，见 _refresh_tabs。
        """
        if tab["path"]:
            name = os.path.basename(tab["path"])
            if name.lower().endswith(EXT):
                name = name[:-len(EXT)]
        else:
            name = (tab["proj"].get("name") or "").strip() or "未命名"
        return ("● " if tab["dirty"] else "") + name

    def _apply_tab_title(self):
        """窗口标题跟随当前标签"""
        tab = self.tabs[self.cur]
        if tab["path"]:
            name = os.path.basename(tab["path"])
        else:
            name = (tab["proj"].get("name") or "").strip() or "未命名"
        self.root.title("%s  %s  —  %s%s" % (
            APP_NAME, APP_SIGN, name, "（未保存）" if tab["dirty"] else ""))

    # ---------- 标签条尺寸 ----------
    # 宽度预算写死，不用实时量 strip.winfo_width()。
    # 理由：strip 是 fill=X 的，按钮变宽会把窗口撑大，窗口变宽又能算出更大字号 ——
    # 正反馈会一路把窗口顶出屏幕。所以用固定预算，字号只跟「分到每个标签的槽宽」走。
    TAB_BUDGET = 1600        # 标签区可用宽度（px，对应窗口约 1668 宽）
    TAB_RESERVE = 230        # 留给「＋ 新建 / ✕ 关闭」和间距
    TAB_PER_LINE = 8.0       # 每行最多 8 个汉字（半角字符算 0.5 个）
    TAB_SLOT = (130.0, 320.0)   # 单个标签的槽宽上下限
    TAB_FS = (12, 24)        # 字号上下限
    TAB_MAX_LINES = 4        # 最多 4 行（32 个字），再长就截断
    # 一个汉字实际占多少像素 / 每点（兜底值，真实值由 _calibrate_tab_unit 现量）。
    # 不能按「字号 = 像素」算，那样算出的字会比槽宽大三分之一，按钮横向溢出。
    TAB_PX_PER_UNIT = 1.3333

    @classmethod
    def _calibrate_tab_unit(cls):
        """实测「1 个汉字占多少像素 / 每点」，覆盖 TAB_PX_PER_UNIT 的兜底值。

        为什么现量：同一台机器，进程里 tk scaling 可能是 1.333 也可能是 1.667
        （取决于当前进程有没有声明 DPI 感知）。写死系数就会让字号差四分之一 ——
        偏大时 8 个汉字撑出按钮、折行被迫提前；偏小时字发虚、留白过大。
        """
        try:
            import tkinter.font as tkfont
            probe = tkfont.Font(family=FONT, size=20)
            v = probe.measure("幽") / 20.0
            if 0.5 < v < 4.0:            # 合理范围才采纳，防异常值把字号算飞
                cls.TAB_PX_PER_UNIT = v
        except Exception:
            pass                          # 没有 Tk 环境（纯逻辑测试）就用兜底值
        return cls.TAB_PX_PER_UNIT

    @staticmethod
    def _wrap_cjk(s, per=8.0, max_lines=4, measure=None, max_px=None):
        """按「汉字=1、半角=0.5」折行，每行不超过 per 个单位。

        再给一个像素上限 max_px 时，两者取更紧的那个 —— 因为半角字符（AOGZ-MY2065-）
        虽然只算 0.5 个单位，实际比半个汉字宽，光按单位折照样会撑出按钮。

        为什么自己折而不用 wraplength：Tk 的自动折行按空格断词，
        「AOGZ-MY2065-XL码」这种没有空格的名字它根本不折，会直接溢出。
        """
        lines, cur, w = [], "", 0.0
        for ch in s:
            cw = 0.5 if ord(ch) <= 0x2E80 else 1.0
            if cur and (w + cw > per + 1e-9
                        or (max_px and measure and measure(cur + ch) > max_px)):
                lines.append(cur)
                cur, w = "", 0.0
            cur += ch
            w += cw
        if cur:
            lines.append(cur)
        if len(lines) > max_lines:
            lines = lines[:max_lines]
            lines[-1] = lines[-1][:-1] + "…"
        return "\n".join(lines)

    def _tab_metrics(self):
        """算出每个标签的槽宽、左右内边距、字号（粗算，随后由探针按钮修正）。

        字号 = 槽宽 / (每行字数 × 单字宽度) —— 也就是「一行 8 个字正好填满按钮」，
        所以按钮分到多宽，字就按比例多大。

        取整用 floor 不用 round：round 会偶尔进一位，让 8 个字比按钮内宽多出一点点，
        折行守卫就会提前断行，看起来像「只显示了 6 个字」。
        """
        n = max(1, len(self.tabs))
        lo, hi = self.TAB_SLOT
        slot = min(hi, max(lo, (self.TAB_BUDGET - self.TAB_RESERVE) / float(n)))
        px = 18 if slot >= 220 else 12
        f_lo, f_hi = self.TAB_FS
        fs = int((slot - 2 * px) / (self.TAB_PER_LINE * self.TAB_PX_PER_UNIT))
        return slot, px, max(f_lo, min(f_hi, fs))

    def _fit_tab_font(self, parent, slot, px, fs_guess):
        """用一颗探针按钮实测「8 个汉字 + 按钮自带内衬」能不能装进槽宽。

        为什么量而不算：tk.Button 除了 padx/pady，还会自带一圈内衬，这圈内衬随 DPI 变
        （本机 1.333 缩放下 18px、1.65 缩放下 25px），写死就会差几像素把 8 个字挤出按钮，
        折行守卫只好提前断到 6 个字 —— 看起来就像「没做到一排 8 个」。

        返回 (字号, 文本可用宽度)。
        """
        import tkinter.font as tkfont
        eight = "幽" * int(self.TAB_PER_LINE)
        f_lo, f_hi = self.TAB_FS
        fs = max(f_lo, min(f_hi, int(fs_guess)))
        cache = getattr(self, "_tab_fit", None)
        if cache is None:
            cache = self._tab_fit = {}
        key = (int(slot), int(px))
        if key in cache:
            return cache[key]
        best = (fs, max(40, int(slot - 2 * px)))
        while fs >= f_lo:
            tw = tkfont.Font(family=FONT, size=fs).measure(eight)
            p = tk.Button(parent, text=eight, font=(FONT, fs), padx=px, pady=1,
                          relief="flat", bd=0, highlightthickness=0)
            p.update_idletasks()
            need = p.winfo_reqwidth()      # 8 个字 + 内衬，按钮实际要占的宽度
            p.destroy()
            best = (fs, max(40, int(slot - (need - tw))))   # 扣掉内衬，剩下给文字
            if need <= slot:
                break
            fs -= 1
        cache[key] = best
        return best


    def _refresh_tabs(self):
        """重建顶部标签条：一颗按钮一个项目 + ＋ 新建 + ✕ 关闭（多于一页才显示）

        按钮上的文字完整显示、按每行 8 个汉字折行；按钮分到的宽度越宽，字号越大。
        """
        strip = getattr(self, "tabstrip", None)
        if strip is None or not strip.winfo_exists():
            return
        for w in strip.winfo_children():
            w.destroy()
        st = ttk.Style()
        _ = st   # 保留引用，便于以后改主题
        self._calibrate_tab_unit()
        slot, px, fs_guess = self._tab_metrics()
        fs, max_px = self._fit_tab_font(strip, slot, px, fs_guess)
        # 选中态用绿色（配色取自参考图 #3C8C4E）；ttk 主题按钮改不了底色，
        # 所以标签用经典 tk.Button（可设背景色）
        CUR_BG, CUR_BG_HOVER, CUR_BG_ACT, CUR_FG = "#3C8C4E", "#46A05A", "#31743F", "#FFFFFF"
        IDLE_BG, IDLE_BG_HOVER, IDLE_BG_ACT, IDLE_FG = "#E7E7E7", "#DBDBDB", "#D0D0D0", "#5A5A5A"

        try:
            import tkinter.font as tkfont
            f_norm = tkfont.Font(family=FONT, size=fs, weight="normal")
            f_bold = tkfont.Font(family=FONT, size=fs, weight="bold")
        except Exception:
            f_norm = f_bold = None

        def mk(text, fontsize, bold, cur, cmd, pad_x, pad_y, mini=False):
            b = tk.Button(strip, text=text,
                          font=(FONT, fontsize, "bold" if bold else "normal"),
                          bg=(CUR_BG if cur else IDLE_BG),
                          fg=(CUR_FG if cur else IDLE_FG),
                          activebackground=(CUR_BG_ACT if cur else IDLE_BG_ACT),
                          activeforeground=(CUR_FG if cur else "#333333"),
                          relief="flat", bd=0, highlightthickness=0,
                          justify="left",          # 多行时左对齐，长短行不会看着歪
                          padx=pad_x, pady=pad_y, cursor="hand2", command=cmd)
            b.bind("<Enter>", lambda e: b.config(bg=(CUR_BG_HOVER if cur else IDLE_BG_HOVER)))
            b.bind("<Leave>", lambda e: b.config(bg=(CUR_BG if cur else IDLE_BG)))
            return b

        for i, tab in enumerate(self.tabs):
            cur = (i == self.cur)
            # 「▶」标当前页、「●」标未保存，都算进第一行的 8 个字里，不额外撑宽
            raw = ("▶ " if cur else "") + self._tab_disp(tab)
            fnt = f_bold if cur else f_norm
            meas = (lambda t, f=fnt: f.measure(t)) if fnt is not None else None
            txt = self._wrap_cjk(raw, self.TAB_PER_LINE, self.TAB_MAX_LINES,
                                 measure=meas, max_px=max_px)
            mk(txt, fs, cur, cur, lambda i=i: self._switch_tab(i), px, 2).pack(
                side=tk.LEFT, fill=tk.Y, padx=(0, 6), pady=(0, 2))
        # 「＋ 新建 / ✕ 关闭」不跟着标签拉高，竖直居中
        mk("＋ 新建", 14, False, False, self._new_tab, 10, 3).pack(
            side=tk.LEFT, padx=(12, 4), pady=(0, 2))
        if len(self.tabs) > 1:
            mk("✕ 关闭", 14, False, False, self._close_tab, 10, 3).pack(
                side=tk.LEFT, padx=(2, 0), pady=(0, 2))
        self._apply_tab_title()

    def _switch_tab(self, i):
        """点击切换：先把当前界面内容收进旧标签，再载入新标签"""
        if i == self.cur or i < 0 or i >= len(self.tabs):
            return
        self._collect(dirty=False)
        self.cur = i
        self._after_tab_change()

    def _after_tab_change(self):
        """切换后：proj/path 指向新标签，同步 UI + 预览 + 标签条"""
        # 表格编辑窗口绑定的是旧项目 —— 切标签后关掉，避免改错项目
        w = getattr(self, "_twin", None)
        if w is not None and w.winfo_exists() and getattr(self, "_tproj", None) is not self.tabs[self.cur]["proj"]:
            w.destroy()
        tab = self.tabs[self.cur]
        self.proj = tab["proj"]
        self.path = tab["path"]
        self._building = True
        self._sync_to_ui()
        self._building = False
        self.redraw()
        self._refresh_tabs()

    def _new_tab(self, proj=None, path=None, status_note=None, from_task=None):
        """新建一个标签页；可带项目（用于打开任务快照），否则给空白默认项目

        from_task：这个标签是从哪个打印任务载入的。记下来是为了同一个任务只开一个标签 ——
        见 _load_task()。
        """
        self._collect(dirty=False)           # 当前 UI 收进旧标签
        if proj is None:
            proj = eng.normalize({})
            dirty = False
        else:
            dirty = True                     # 快照内容未保存
        self.tabs.append({"proj": proj, "path": path, "dirty": dirty,
                          "from_task": from_task})
        self.cur = len(self.tabs) - 1
        self._after_tab_change()
        self.status.config(text=status_note or "已新建标签页（未命名）")

    def _save_tab(self, tab, as_new=False):
        """保存指定标签；成功返回 True，取消/失败返回 False"""
        path = tab["path"]
        if as_new or not path:
            name = (tab["proj"].get("name") or "水洗标") + EXT
            path = filedialog.asksaveasfilename(
                title="保存项目", defaultextension=EXT,
                initialfile=name, filetypes=[("水洗标项目", "*" + EXT)])
            if not path:
                return False
        try:
            eng.save_project(tab["proj"], path)
        except Exception as e:
            messagebox.showerror("保存失败", str(e))
            return False
        tab["path"] = path
        tab["dirty"] = False
        tab["from_task"] = None      # 存成文件了就不再是「某任务的快照」，别再复用这个标签
        self._drop_autosave(tab)          # 真存好了，草稿就没用了
        self._refresh_tabs()
        self.status.config(text="已保存：" + os.path.basename(path))
        return True

    def _on_app_close(self):
        """关整窗：逐个处理未保存的标签，一个都不许无声丢掉。

        返回三态：是 = 全部保存后退出 / 否 = 不保存直接退出 / 取消 = 回去继续编辑。
        只有「保存失败」或「保存对话框被取消」才中止退出 —— 那种情况下丢文件才是真丢。
        """
        self._collect(dirty=False)
        dirty = [t for t in self.tabs if t["dirty"]]
        if dirty:
            names = "\n".join("　· " + self._tab_disp(t).replace("● ", "")
                              for t in dirty)
            r = messagebox.askyesnocancel(
                "退出前确认",
                "有 %d 个标签还没保存：\n\n%s\n\n"
                "是 = 全部保存后退出\n否 = 不保存，直接退出\n取消 = 回去继续编辑"
                % (len(dirty), names))
            if r is None:
                return
            if r:
                for t in dirty:
                    if not self._save_tab(t):
                        self.status.config(text="保存未完成，已取消退出")
                        return
            else:
                for t in dirty:           # 明确选了「不保存」→ 草稿也一并清掉，别下次又弹
                    self._drop_autosave(t)
        self.qstop.set()          # 通知队列线程退出（它是 daemon，这里只是让它干净点）
        self.qevt.set()
        self.root.destroy()

    # ---------------- 撤销 / 重做 ----------------
    def _note_undo(self, before_json):
        """记一步撤销。连续敲字合并：700ms 内的连击只保留最初那份快照"""
        if self._undo_pending is None:
            self._undo_pending = (self.cur, before_json)
        if self._undo_job:
            try:
                self.root.after_cancel(self._undo_job)
            except Exception:
                pass
        self._undo_job = self.root.after(700, self._commit_undo)

    def _commit_undo(self):
        self._undo_job = None
        if self._undo_pending is None:
            return
        item = self._undo_pending
        self._undo_pending = None
        if self._undo_stack and self._undo_stack[-1][1] == item[1]:
            return                          # 和上一步一模一样，不重复记
        self._undo_stack.append(item)
        if len(self._undo_stack) > UNDO_LIMIT:
            self._undo_stack.pop(0)
        self._redo_stack.clear()

    def _on_undo_key(self, _e=None):
        self.undo()
        return "break"

    def _on_redo_key(self, _e=None):
        self.redo()
        return "break"

    def undo(self):
        if self._undo_job:
            try:
                self.root.after_cancel(self._undo_job)
            except Exception:
                pass
            self._undo_job = None
        self._commit_undo()
        if not self._undo_stack:
            self.status.config(text="没有可撤销的改动")
            return
        idx, snap = self._undo_stack.pop()
        self._redo_stack.append(
            (idx, json.dumps(self.proj, sort_keys=True, ensure_ascii=False)))
        self._apply_snapshot(idx, snap)
        self.status.config(text="已撤销（还剩 %d 步）" % len(self._undo_stack))

    def redo(self):
        if not self._redo_stack:
            self.status.config(text="没有可重做的改动")
            return
        idx, snap = self._redo_stack.pop()
        self._undo_stack.append(
            (idx, json.dumps(self.proj, sort_keys=True, ensure_ascii=False)))
        self._apply_snapshot(idx, snap)
        self.status.config(text="已重做（还剩 %d 步）" % len(self._redo_stack))

    def _apply_snapshot(self, idx, snap_json):
        """把一份快照灌回标签。全程 _building=True，防止控件回写把快照冲掉"""
        self.cur = max(0, min(int(idx), len(self.tabs) - 1))
        tab = self.tabs[self.cur]
        try:
            tab["proj"] = eng.normalize(json.loads(snap_json))
        except Exception:
            return
        tab["dirty"] = True
        self.proj = tab["proj"]
        self._tproj = self.proj
        self._building = True
        try:
            self._sync_to_ui()
        finally:
            self._building = False
        self._refresh_tabs()
        self.redraw()
        self._refresh_table_dialog()

    def _refresh_table_dialog(self):
        """撤销后让表格窗口跟上（它持有一份自己的 StringVar）"""
        w = getattr(self, "_twin", None)
        if w is None or not w.winfo_exists():
            return
        t = self.proj.get("table") or {}
        self._building = True
        try:
            self.v_rowh.set(str(t.get("row_h_mm", 4.0)))
            self.v_gapt.set(str(t.get("gap_top_mm", 2.0)))
            self.v_gapb.set(str(t.get("to_bottom_mm", 4.0)))
            self.v_bwid.set(str(t.get("border_dot", 2)))
            self.v_cap.set(str(t.get("caption", "")))
            self.v_capf.set(str(t.get("caption_font", 16)))
        finally:
            self._building = False
        self._rebuild_table_rows()
        self._table_info()

    # ---------------- 自动存盘 ----------------
    def _autosave_key(self, tab):
        src = tab.get("path") or ""
        if src:
            base = os.path.splitext(os.path.basename(src))[0]
        else:
            base = tab["proj"].get("name") or "未命名"
        safe = re.sub(r'[\\/:*?"<>|\s]+', "_", base).strip("_") or "未命名"
        return os.path.join(AUTOSAVE_DIR, safe + ".autosave")

    def _schedule_autosave(self):
        self._autosave_job = self.root.after(AUTOSAVE_MS, self._autosave_tick)

    def _autosave_tick(self):
        try:
            self._autosave_once()
        except Exception:
            pass
        finally:
            self._schedule_autosave()

    def _autosave_once(self):
        """只写有未保存改动的标签；先写 .tmp 再原子改名，避免崩在写一半留下坏文件"""
        if self._building:
            return
        dirty = [t for t in self.tabs if t["dirty"]]
        if not dirty:
            return
        self._collect(dirty=False)          # 把正在输入、还没提交的内容也收进来
        try:
            os.makedirs(AUTOSAVE_DIR, exist_ok=True)
        except OSError:
            return
        for t in dirty:
            try:
                dst = self._autosave_key(t)
                data = {"proj": eng.normalize(t["proj"]), "path": t["path"],
                        "saved_at": time.strftime("%Y-%m-%d %H:%M:%S")}
                with open(dst + ".tmp", "w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False)
                os.replace(dst + ".tmp", dst)
            except Exception:
                pass
        self._prune_autosaves()

    def _prune_autosaves(self):
        """草稿超过 AUTOSAVE_MAX 份就删最旧的（按修改时间）。"""
        try:
            items = []
            for f in os.listdir(AUTOSAVE_DIR):
                if not f.endswith(".autosave"):
                    continue
                full = os.path.join(AUTOSAVE_DIR, f)
                try:
                    items.append((os.path.getmtime(full), full))
                except OSError:
                    pass
            if len(items) <= AUTOSAVE_MAX:
                return
            items.sort()
            for _, full in items[:len(items) - AUTOSAVE_MAX]:
                try:
                    os.remove(full)
                except OSError:
                    pass
        except OSError:
            pass

    def _drop_autosave(self, tab):
        try:
            p = self._autosave_key(tab)
            if os.path.exists(p):
                os.remove(p)
        except Exception:
            pass

    def _offer_recovery(self):
        """启动时若发现草稿，问一次要不要恢复"""
        try:
            files = sorted(f for f in os.listdir(AUTOSAVE_DIR) if f.endswith(".autosave"))
        except OSError:
            return
        if not files:
            return
        shown = "\n".join("　· " + f[:-9] for f in files[:6])
        more = "" if len(files) <= 6 else "\n　… 还有 %d 份" % (len(files) - 6)
        r = messagebox.askyesno(
            "发现未保存草稿",
            "上次没有正常退出，找到 %d 份草稿：\n\n%s%s\n\n要载入吗？"
            "（选「否」会删掉它们）" % (len(files), shown, more))
        if not r:
            for f in files:
                try:
                    os.remove(os.path.join(AUTOSAVE_DIR, f))
                except OSError:
                    pass
            return
        loaded = 0
        for f in files:
            full = os.path.join(AUTOSAVE_DIR, f)
            try:
                with open(full, encoding="utf-8") as fp:
                    data = json.load(fp)
                proj = eng.normalize(data.get("proj") or {})
            except Exception:
                continue
            self.tabs.append({"proj": proj, "path": data.get("path"), "dirty": True})
            loaded += 1
        if loaded:
            self.cur = len(self.tabs) - 1
            self._after_tab_change()
            self.status.config(text="已恢复 %d 份草稿（仍为未保存状态）" % loaded)

    def _close_tab(self):
        """关闭当前标签：有改动先问 —— 是=保存并关 / 否=不保存关 / 取消=不动"""
        if len(self.tabs) <= 1:
            messagebox.showinfo("关闭标签", "只剩最后一个标签页了，不能关闭。")
            return
        tab = self.tabs[self.cur]
        if tab["dirty"]:
            r = messagebox.askyesnocancel(
                "关闭标签",
                "「%s」有未保存的修改。\n\n是 = 保存并关闭\n否 = 不保存，直接关闭" % self._tab_disp(tab))
            if r is None:
                return
            if r and not self._save_tab(tab):
                return                        # 保存失败/取消 → 不关
            if not r:
                # 明确选了「不保存」→ 草稿也一并清掉。
                # 以前漏了这一句：关标签选「否」草稿留在 .autosave 里，
                # 下次启动又弹「发现未保存草稿」，看着像程序没记住选择。
                self._drop_autosave(tab)
        self.tabs.pop(self.cur)
        self.cur = min(self.cur, len(self.tabs) - 1)
        self._after_tab_change()
        self.status.config(text="已关闭标签页")

    def do_save(self, as_new=False):
        self._collect(dirty=False)
        self._save_tab(self.tabs[self.cur], as_new)

    def do_open(self):
        """打开文件 → 每个文件各开一个新标签页（支持多选），并跳到最后打开的那个"""
        paths = filedialog.askopenfilenames(
            title="打开项目（可多选，每个文件一个标签页）",
            filetypes=[("水洗标项目", "*" + EXT), ("所有文件", "*.*")])
        if not paths:
            return
        self._collect(dirty=False)           # 当前 UI 收进旧标签
        ok = 0
        for path in paths:
            try:
                proj = eng.load_project(path)
            except Exception as e:
                messagebox.showerror("打开失败", "%s\n%s" % (path, e))
                continue
            self.tabs.append({"proj": proj, "path": path, "dirty": False})
            self.cur = len(self.tabs) - 1
            ok += 1
        if ok:
            self._after_tab_change()
            self.status.config(text="已打开 %d 个文件" % ok)

    # ---------------- 导入 DLabel 老模板（.ddl）----------------
    # 为什么这条路走得通：.ddl 是 DLabel 自己的 XML，文字就存在节点的
    # sourcetext / value 属性里，纯解析就能取，不用 OCR、不用调 DLabel 接口。
    def do_import_ddl(self):
        """从 .ddl 里取号型、成分和尺码表，填进一个**新**标签页。

        四个刻意选择：
          · 不改当前标签 —— 号型和成分正是你在校对的东西，悄悄覆盖比多开一页危险
          · 先弹确认框 —— 602 个历史模板里 601 个能取到成分、597 个能取到号型；
            取不到的要让你当场补，而不是打完才发现是空的
          · 只取文字，不取排版 —— 版式仍由本程序统一，免得两套排版打架
          · 尺码表只认「一整行都是尺码 + 正下方格数相同的数值行」，认不出就不生成，
            绝不猜 —— 猜错一格就要手工返工整张表
        """
        path = filedialog.askopenfilename(
            title="选择 DLabel 标签模板（.ddl）",
            filetypes=[("DLabel 标签模板", "*.ddl"), ("所有文件", "*.*")])
        if not path:
            return
        try:
            res = ddl.parse_ddl(path)
        except ddl.DdlError as e:
            messagebox.showerror("读不了这个文件",
                                 "%s\n\n%s" % (os.path.basename(path), e))
            return
        self._ddl_confirm(res)

    @staticmethod
    def _has_label(line):
        """这一行是不是自带标签（「面料: …」「里料成分: …」）"""
        return re.match(r"^\s*[\u4e00-\u9fa5A-Za-z]{0,10}[:：]", line) is not None

    def _ddl_confirm(self, res):
        """导入确认框：号型、成分都可当场改，再填进新标签页"""
        name = os.path.basename(res["file"])
        paper = res.get("paper")

        win = tk.Toplevel(self.root)
        win.title("导入模板文案")
        win.transient(self.root)
        win.resizable(False, False)

        # 模板名可能很长（600 多个历史文件名里什么长度都有），让它换行而不是把
        # 整个对话框撑成屏幕宽
        ttk.Label(win, text="来源：%s" % name, font=(FONT, 11, "bold"),
                  wraplength=420, justify="left").pack(anchor="w", padx=14,
                                                       pady=(12, 2))
        ttk.Label(win, text=("这个模板的纸张是 %g × %g mm" % (paper["w"], paper["h"])
                             if paper else "没读到纸张尺寸"),
                  foreground="#777").pack(anchor="w", padx=14)

        box = ttk.Frame(win, padding=(14, 8))
        box.pack(fill=tk.BOTH, expand=True)
        ttk.Label(box, text="号型（第一行）").grid(row=0, column=0, sticky="e", pady=4)
        v_size = tk.StringVar(value=res["size"])
        ttk.Entry(box, textvariable=v_size, width=36).grid(
            row=0, column=1, sticky="w", padx=(6, 0), pady=4)
        ttk.Label(box, text="成分（一行一个）").grid(
            row=1, column=0, sticky="ne", pady=4)
        txt = tk.Text(box, width=36, height=6, wrap="none")
        txt.grid(row=1, column=1, sticky="w", padx=(6, 0), pady=4)
        txt.insert("1.0", "\n".join(res["fabrics"]))

        # 模板里认出了「尺码行 + 数值行」就把表格一起生成出来 —— 这正是老标签
        # 里最费手的一段（4~5 格克数要一格一格敲）。默认勾上，因为认出来的
        # 内容就在下面的灰字里明摆着，认错了当场取消勾选即可。
        sizes = res.get("sizes") or []
        tvals = res.get("values") or []
        v_tbl = tk.BooleanVar(value=False)
        row_paper = 2
        if len(sizes) >= 2:
            v_tbl.set(True)
            trow = ttk.Frame(box)
            trow.grid(row=2, column=1, sticky="w", padx=(6, 0), pady=(6, 0))
            ttk.Checkbutton(trow, text="同时生成尺码表", variable=v_tbl).pack(anchor="w")
            # 成对显示「M 147g · L 154g」而不是两串并列 —— 两串的话得自己数
            # 第 3 个克数对第 3 个尺码，一眼看不出对错就没法核对
            if tvals and len(tvals) == len(sizes):
                detail = " · ".join("%s %s" % (s, v) for s, v in zip(sizes, tvals))
            else:
                detail = " / ".join(str(s) for s in sizes) + "（克数留空手填）"
            ttk.Label(trow, text=detail, foreground="#666", wraplength=330,
                      justify="left").pack(anchor="w", padx=(22, 0), pady=(1, 0))
            row_paper = 3

        v_paper = tk.BooleanVar(value=False)
        if paper:
            ttk.Checkbutton(box, text="同时把纸张改成 %g × %g mm" % (paper["w"], paper["h"]),
                            variable=v_paper).grid(row=row_paper, column=1, sticky="w",
                                                   padx=(6, 0), pady=(4, 0))

        # 取不到的字段直接说清楚，别让人以为已经填好了
        if res["notes"]:
            ttk.Label(win, text="⚠ " + "；".join(res["notes"]),
                      foreground="#C25E00", wraplength=430,
                      justify="left").pack(anchor="w", padx=14, pady=(6, 0))

        btns = ttk.Frame(win, padding=(14, 10))
        btns.pack(fill=tk.X)

        def do_fill(*_):
            size = v_size.get().strip()
            lines = [s.strip() for s in txt.get("1.0", "end").splitlines()
                     if s.strip()]
            proj = eng.normalize({})
            proj["name"] = os.path.splitext(name)[0].strip() or "导入模板"
            # 号型留空时写出一个空壳「号型: 」，比留着默认的 175/92A 安全 ——
            # 后者的后果是打出一张没人发现的错标签
            proj["title"] = ("号型: " + size) if size else "号型: "
            proj["fabrics"] = lines
            # 成分行自带标签（面料/里料/胆料）时不能再加「成分:」前缀，否则双标签
            proj["fabric_head"] = "" if any(self._has_label(s) for s in lines) else "成分:"
            made = 0
            if v_tbl.get() and len(sizes) >= 2:
                t = proj.setdefault("table", {})
                t["rows"] = eng.size_table_rows(sizes, tvals)
                # 只记预设里有的尺码，勾选框下次才点得回来（F 这种老写法不记）
                t["size_picks"] = [s for s in sizes if s in SIZE_ORDER]
                t["on"] = True
                made = len(t["rows"])
            if v_paper.get() and paper:
                proj["paper"]["w"] = paper["w"]
                proj["paper"]["h"] = paper["h"]
            win.destroy()
            self._new_tab(proj, None,
                          "已从 %s 导入号型、成分%s"
                          % (name, ("和 %d 行尺码表" % made) if made else ""))

        ttk.Button(btns, text="取消", command=win.destroy).pack(side=tk.RIGHT)
        ttk.Button(btns, text="填入新标签页", command=do_fill).pack(
            side=tk.RIGHT, padx=(0, 8))
        win.bind("<Escape>", lambda e: win.destroy())
        win.bind("<Return>", do_fill)
        # 摆在主窗中间，别跑到屏幕角落
        win.update_idletasks()
        self.root.update_idletasks()
        x = self.root.winfo_rootx() + (self.root.winfo_width() - win.winfo_width()) // 2
        y = self.root.winfo_rooty() + (self.root.winfo_height() - win.winfo_height()) // 3
        win.geometry("+%d+%d" % (max(0, x), max(0, y)))
        try:
            win.grab_set()
        except tk.TclError:
            pass
        win.focus_force()


PREVIEW_BLUE = "#1677FF"
PREVIEW_BLUE_D = "#0E5FD8"


class PrintPreview(tk.Toplevel):
    """打印预览：左边「打印设置」，右边「纸型状态 + 成品预览」。

    三条硬承诺：
      · 预览图由 render_label 出，与图片模式打出来的像素完全一致（几 ms 级，打字都跟得上）
      · 纸型状态读的是「驱动当前默认纸型」，读不到就说读不到，不编数字
      · 超出可印区的部分盖红色斜纹 —— 这是唯一能提前发现「换台机器就打歪/裁掉」的办法
    """

    def __init__(self, app):
        super().__init__(app.root)
        self.app = app
        app._pv_win = self
        self.title("打印预览")
        self.configure(bg="#FFFFFF")
        self.transient(app.root)
        self._img = None
        self._photo = None
        self._paint_key = None      # (图像对象, 宽, 高) —— 尺寸没变就不重新缩放
        self._job = None
        self._cfg_job = None
        self._paper = None
        self._fit_ov = (0.0, 0.0)
        self._setup_proc = None
        self._poll_job = None
        self.zoom = 1.0             # 1.0 = 适应窗口
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        bgw, bgr = "#FFFFFF", "#F4F5F7"

        body = tk.Frame(self, bg=bgw)
        body.pack(fill=tk.BOTH, expand=True)

        left = tk.Frame(body, bg=bgw, width=336)
        left.pack(side=tk.LEFT, fill=tk.Y)
        left.pack_propagate(False)
        tk.Frame(body, bg="#E5E6EB", width=1).pack(side=tk.LEFT, fill=tk.Y)
        right = tk.Frame(body, bg=bgr)
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # ================= 左：打印设置 =================
        lf = tk.Frame(left, bg=bgw, padx=20, pady=16)
        lf.pack(fill=tk.BOTH, expand=True)
        tk.Label(lf, text="打印设置", bg=bgw, fg="#1D2129",
                 font=(FONT, 15, "bold")).pack(anchor="w")

        # 打印机 + 设置
        prow = tk.Frame(lf, bg=bgw)
        prow.pack(fill=tk.X, pady=(16, 6))
        tk.Label(prow, text="打印机", bg=bgw, fg="#4E5969",
                 font=(FONT, 10)).pack(side=tk.LEFT)
        self.lnk_setup = tk.Label(prow, text="设置", bg=bgw, fg=PREVIEW_BLUE,
                                  font=(FONT, 10, "underline"), cursor="hand2")
        self.lnk_setup.pack(side=tk.RIGHT)
        self.lnk_setup.bind("<Button-1>", self._on_setup)
        self.lnk_setup.bind("<Enter>", lambda e: self.lnk_setup.config(fg=PREVIEW_BLUE_D))
        self.lnk_setup.bind("<Leave>", lambda e: self.lnk_setup.config(fg=PREVIEW_BLUE))

        self.v_printer = tk.StringVar(value=app.v_printer.get())
        names = [i["name"] for i in eng.list_printer_info()]
        # 主界面没选就保持空白，不替用户挑一台 —— 预览里「自己冒出来一台」
        # 会让人以为已经选好了，点打印就打出去
        self.cb = ttk.Combobox(lf, textvariable=self.v_printer, values=names,
                               state="readonly", font=(FONT, 10))
        self.cb.pack(fill=tk.X, ipady=2)
        self.cb.bind("<<ComboboxSelected>>", self._on_printer)

        # 打印方向
        tk.Label(lf, text="打印方向", bg=bgw, fg="#4E5969",
                 font=(FONT, 10)).pack(anchor="w", pady=(16, 6))
        drow = tk.Frame(lf, bg=bgw)
        drow.pack(fill=tk.X)
        self.dir_btns = {}
        for ang, txt in ((0, "A  0°"), (90, "▷  90°"), (180, "▽  180°")):
            b = tk.Button(drow, text=txt, font=(FONT, 10), bd=0, relief="flat",
                          highlightthickness=1, takefocus=0, padx=10, pady=5,
                          cursor="hand2", command=lambda a=ang: self._set_dir(a))
            b.pack(side=tk.LEFT, padx=(0, 6))
            self.dir_btns[ang] = b

        # 打印数量
        tk.Label(lf, text="打印数量", bg=bgw, fg="#4E5969",
                 font=(FONT, 10)).pack(anchor="w", pady=(16, 6))
        self.v_copies = tk.StringVar(value=app.v_copies.get())
        ttk.Spinbox(lf, from_=1, to=999, increment=1, textvariable=self.v_copies,
                    font=(FONT, 10), justify="left").pack(fill=tk.X, ipady=2)
        self.v_copies.trace_add("write", lambda *_: self._on_copies())

        # 偏移（初值取自项目，别用 0 覆盖用户之前存过的值）
        self._num_row(lf, "水平偏移", "v_offx", app.v_offx.get() or "0.00")
        self._num_row(lf, "垂直偏移", "v_offy", app.v_offy.get() or "0.00")

        # ================= 左下：打印按钮 =================
        foot = tk.Frame(left, bg=bgw, padx=20)
        foot.pack(side=tk.BOTTOM, fill=tk.X, pady=(0, 16))
        self.lbl_msg = tk.Label(foot, text="", bg=bgw, fg="#86909C",
                                font=(FONT, 9), wraplength=290, justify="left")
        self.lbl_msg.pack(anchor="w", pady=(0, 8))
        self.btn_print = tk.Button(foot, text="打印(1)", bg=PREVIEW_BLUE, fg="#FFFFFF",
                                   activebackground=PREVIEW_BLUE_D, activeforeground="#FFFFFF",
                                   font=(FONT, 12, "bold"), bd=0, relief="flat",
                                   pady=10, cursor="hand2", command=self._on_print)
        self.btn_print.pack(fill=tk.X)

        # ================= 右：标签信息 + 纸型 + 预览 =================
        rf = tk.Frame(right, bg=bgr, padx=22, pady=16)
        rf.pack(fill=tk.BOTH, expand=True)
        tk.Label(rf, text="标签信息", bg=bgr, fg="#1D2129",
                 font=(FONT, 14, "bold")).pack(anchor="w")
        self.lbl_info = tk.Label(rf, text="", bg=bgr, fg="#4E5969",
                                 font=(FONT, 10), justify="left")
        self.lbl_info.pack(anchor="w", pady=(6, 0))

        self.lbl_paper = tk.Label(rf, text="", bg=bgr, fg="#4E5969", font=(FONT, 10),
                                  justify="left", wraplength=620, anchor="w")
        self.lbl_paper.pack(fill=tk.X, pady=(10, 0))

        hr = tk.Frame(rf, bg=bgr)
        hr.pack(fill=tk.X, pady=(14, 6))
        tk.Label(hr, text="标签预览", bg=bgr, fg="#1D2129",
                 font=(FONT, 14, "bold")).pack(side=tk.LEFT)
        # pack 顺序决定右对齐后的左右次序：先 pack 的最靠右
        self.lbl_zoom = tk.Label(hr, text="适应窗口", bg=bgr, fg="#86909C", font=(FONT, 9))
        self.lbl_zoom.pack(side=tk.RIGHT, padx=(8, 0))
        for txt, cmd in (("⤢", self._fit), ("－", lambda: self._zoom(1 / 1.25)),
                         ("＋", lambda: self._zoom(1.25))):
            tk.Button(hr, text=txt, font=(FONT, 11), bd=0, relief="flat",
                      bg=bgr, fg="#4E5969", activebackground="#E5E6EB",
                      cursor="hand2", padx=8, pady=1, command=cmd).pack(side=tk.RIGHT, padx=(4, 0))

        self.cv = tk.Canvas(rf, bg="#EDEFF2", highlightthickness=0)
        self.cv.pack(fill=tk.BOTH, expand=True)
        self.cv.bind("<Configure>", lambda e: self._schedule_paint())
        self.cv.bind("<MouseWheel>", self._on_wheel)

        # ---- 初始化 ----
        self._sync_dir_buttons()
        self._on_copies()
        self._probe_paper()
        self._render()
        w, h = 1000, 722
        self.update_idletasks()
        rx, ry = app.root.winfo_rootx(), app.root.winfo_rooty()
        rw, rh = app.root.winfo_width(), app.root.winfo_height()
        self.geometry("%dx%d+%d+%d" % (w, h, rx + max(0, (rw - w) // 2),
                                       ry + max(0, (rh - h) // 3)))
        self.minsize(880, 620)
        self.after(80, self._redraw_canvas)

    # ---------------- 小控件工厂 ----------------
    def _num_row(self, parent, label, attr, default):
        bgw = "#FFFFFF"
        tk.Label(parent, text=label, bg=bgw, fg="#4E5969",
                 font=(FONT, 10)).pack(anchor="w", pady=(14, 6))
        row = tk.Frame(parent, bg=bgw)
        row.pack(fill=tk.X)
        var = tk.StringVar(value=default)
        try:
            sp = ttk.Spinbox(row, from_=-50.0, to=50.0, increment=0.1,
                             textvariable=var, font=(FONT, 10), format="%.2f")
        except Exception:                    # 个别 Tk 版本不认 format
            sp = ttk.Spinbox(row, from_=-50.0, to=50.0, increment=0.1,
                             textvariable=var, font=(FONT, 10))
        sp.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=2)
        tk.Label(row, text="mm", bg=bgw, fg="#86909C",
                 font=(FONT, 10)).pack(side=tk.LEFT, padx=(8, 2))
        setattr(self, attr, var)
        var.trace_add("write", lambda *_: self._on_offset())
        return var

    # ---------------- 参数变化 ----------------
    def _sync_dir_buttons(self):
        try:
            ang = int(float(self.app.v_direction.get() or 0)) % 360
        except Exception:
            ang = 0
        for a, b in self.dir_btns.items():
            if a == ang:
                b.config(bg="#EAF3FF", fg=PREVIEW_BLUE, activebackground="#DCEBFF",
                         activeforeground=PREVIEW_BLUE, highlightbackground=PREVIEW_BLUE)
            else:
                b.config(bg="#FAFAFA", fg="#4E5969", activebackground="#F0F1F3",
                         activeforeground="#4E5969", highlightbackground="#D9DCE1")

    def _set_dir(self, ang):
        self.app.v_direction.set(str(ang))
        self._sync_dir_buttons()
        self._schedule_render()

    def _on_printer(self, _e=None):
        name = self.v_printer.get()
        # 在预览里选也算「亲手切换」：主界面同步、记成软件默认，下次打开就是它
        self.app.set_printer(name)
        eng.clear_paper_cache(name)
        self._probe_paper()
        self._schedule_render()

    def _on_copies(self):
        try:
            n = max(1, min(999, int(float(str(self.v_copies.get()).strip()))))
        except Exception:
            n = 1
        self.btn_print.config(text="打印(%d)" % n)

    def _on_offset(self):
        self.app.v_offx.set(self.v_offx.get())
        self.app.v_offy.set(self.v_offy.get())
        self._schedule_render()

    # ---------------- 纸型探测 ----------------
    def _probe_paper(self, refresh=False):
        self._paper = eng.printer_paper_info(self.v_printer.get(), refresh=refresh)
        self._update_paper_line()

    def _update_paper_line(self):
        """底栏「这台机 · 这张纸」到底对不对

        1.10 之后判断口径变了：不再看「驱动当前挂着什么纸」，而是
        真的建一次画布、把纸型设成这份文件的尺寸、再读回来 ——
        驱动接受就绿，不接受就红。因为这才是真正会发生的事。
        """
        p = self.app.proj
        w, h = eng.label_size_mm(p)
        printer = self.v_printer.get()
        media = str((p.get("paper") or {}).get("media") or "gap")
        # 以面板上选的那台为准：通道要按「这台机 + 这个文件的打印方式」一起判
        probe = dict(p or {})
        probe["printer"] = printer
        mode = eng.resolve_mode(probe)
        if not printer:
            self.lbl_paper.config(text="?  还没选打印机", fg="#D93026")
            self._fit_ov = (0.0, 0.0)
            return
        if mode == "image":
            # 缓存放在实例上：纸宽/纸长每改一个字都会走到这里，
            # 每次都新建画布太慢，而同一个尺寸问第二遍答案不会变。
            cache = getattr(self, "_apply_cache", None)
            if cache is None:
                cache = self._apply_cache = {}
            r = eng.check_paper_apply(printer, w, h, cache=cache)
            gw, gh = r["got"]
            if r["ok"]:
                txt = ("✓  图片模式会按 %.1f × %.1f mm 设自定义纸（驱动已确认接受）" % (w, h))
                fg = "#1F9D55"
            elif not r.get("devmode"):
                txt = ("?  %s —— 只能靠驱动自己的纸型设置，建议切「TSPL 直发」" % r["why"])
                fg = "#86909C"
            else:
                txt = ("✕  %s —— 这么打会歪/被裁，建议切「TSPL 直发」或改纸宽纸长" % r["why"])
                fg = "#D93026"
            self._fit_ov = (0.0, 0.0)
        else:
            gap = float((p.get("paper") or {}).get("gap") or 2.0)
            gtxt = "GAP 0（连续纸没有缝）" if media == "continuous" \
                else "GAP %g mm（靠缝定位）" % gap
            txt = ("✓  TSPL 直发：SIZE %.1f × %.1f mm，%s —— 每份按自己的尺寸发指令" % (w, h, gtxt))
            fg = "#1F9D55"
            self._fit_ov = (0.0, 0.0)
        self.lbl_paper.config(text=txt, fg=fg)

    # ---------------- 渲染 ----------------
    def _schedule_render(self):
        if self._job:
            self.after_cancel(self._job)
        self._job = self.after(160, self._render)

    def _render(self):
        self._job = None
        if not self.winfo_exists():
            return
        try:
            self._img = eng.render_label(self.app.proj, dpi=300)
        except Exception as e:
            self.lbl_paper.config(text="预览渲染失败：%s" % e, fg="#D93026")
            return
        self._paint_key = None
        self._update_info()
        self._update_paper_line()
        self._redraw_canvas()

    def _update_info(self):
        p = self.app.proj
        w, h = eng.label_size_mm(p)
        info = self._paper
        rows = ["名称：%s" % (p.get("name") or "（未命名）"),
                "尺寸：%.1f × %.1f mm ｜ 成品 %.0f × %.0f px @300dpi" %
                (float(p["paper"]["w"]), float(p["paper"]["h"]),
                 self._img.size[0] if self._img else 0,
                 self._img.size[1] if self._img else 0),
                "驱动纸型：%s" % ("%.1f × %.1f mm（这台机器现在还挂着这张纸，"
                              "打印时本程序会按上面「尺寸」临时改掉，不用手动设）" %
                              (info["phys_w"], info["phys_h"])
                              if info else "读不到")]
        self.lbl_info.config(text="\n".join(rows))

    # ---------------- 画布 ----------------
    def _schedule_paint(self):
        if self._cfg_job:
            self.after_cancel(self._cfg_job)
        self._cfg_job = self.after(70, self._redraw_canvas)

    def _redraw_canvas(self):
        self._cfg_job = None
        if self._img is None or not self.winfo_exists():
            return
        cw, ch = self.cv.winfo_width(), self.cv.winfo_height()
        if cw < 60 or ch < 60:
            return
        iw, ih = self._img.size
        pad = 26
        fit = min((cw - 2 * pad) / float(iw), (ch - 2 * pad) / float(ih))
        k = max(0.05, fit * self.zoom)
        dw, dh = max(1, int(round(iw * k))), max(1, int(round(ih * k)))

        key = (id(self._img), dw, dh)
        if key != self._paint_key:                 # 尺寸没变就不重新缩放（拖窗口时才省得掉帧）
            from PIL import Image, ImageTk
            self._photo = ImageTk.PhotoImage(self._img.resize((dw, dh), Image.LANCZOS))
            self._paint_key = key

        self.cv.delete("all")
        x0, y0 = (cw - dw) // 2, (ch - dh) // 2
        self.cv.create_rectangle(x0 - 1, y0 - 1, x0 + dw, y0 + dh,
                                 outline="#C9CDD4", fill="#FFFFFF")
        self.cv.create_image(x0, y0, image=self._photo, anchor="nw")

        # 超限区域：红色斜纹（stipple 是 Tk 原生网点，不额外耗性能）
        w_mm, h_mm = eng.label_size_mm(self.app.proj)
        ox, oy = self._fit_ov
        if ox > 0:
            xf = x0 + int(round((w_mm - ox) / w_mm * dw))
            self.cv.create_rectangle(xf, y0, x0 + dw, y0 + dh,
                                     fill="#E85C4A", stipple="gray25", outline="")
            self.cv.create_line(xf, y0, xf, y0 + dh, fill="#D93026",
                                dash=(5, 3), width=2)
        if oy > 0:
            yf = y0 + int(round((h_mm - oy) / h_mm * dh))
            self.cv.create_rectangle(x0, yf, x0 + dw, y0 + dh,
                                     fill="#E85C4A", stipple="gray25", outline="")
            self.cv.create_line(x0, yf, x0 + dw, yf, fill="#D93026",
                                dash=(5, 3), width=2)

    def _zoom(self, f):
        self.zoom = max(0.4, min(4.0, self.zoom * f))
        self.lbl_zoom.config(text="%d%%" % round(self.zoom * 100))
        self._redraw_canvas()

    def _fit(self):
        self.zoom = 1.0
        self.lbl_zoom.config(text="适应窗口")
        self._redraw_canvas()

    def _on_wheel(self, e):
        self._zoom(1.12 if e.delta > 0 else 1 / 1.12)
        return "break"

    # ---------------- 设置（改纸型）----------------
    def _on_setup(self, _e=None):
        name = self.v_printer.get()
        if not name:
            return
        self._setup_proc = eng.open_printer_setup(name)
        if self._setup_proc is None:
            self.lbl_msg.config(text="打不开打印首选项，请到「设置→打印机」里手动改纸型",
                                fg="#D93026")
            return
        self.lbl_msg.config(text="已打开打印首选项（一般用不着改纸型了 —— 打印时本程序会按每份文件"
                                 "自己的尺寸临时设纸；这里主要用于改浓度/速度/打印方向）…",
                            fg="#4E5969")
        self._poll_setup()

    def _poll_setup(self):
        self._poll_job = None
        proc = self._setup_proc
        if proc is None:
            return
        if proc.poll() is None:                    # 对话框还开着，等它
            self._poll_job = self.after(400, self._poll_setup)
            return
        eng.clear_paper_cache(self.v_printer.get())
        self._probe_paper(refresh=True)
        self._update_info()
        self._redraw_canvas()
        lv, ox, oy = eng.paper_fit(*eng.label_size_mm(self.app.proj), info=self._paper)
        self.lbl_msg.config(
            text=("已重读纸型：%s" % ("匹配" if lv == "ok" else
                                     "仍然不符，请确认改的是这台机器、并且点了确定")),
            fg="#1F9D55" if lv == "ok" else "#D97706")

    # ---------------- 打印 ----------------
    def _on_print(self):
        self.app._collect(dirty=False)
        p = json.loads(json.dumps(self.app.proj))
        p["printer"] = self.v_printer.get()
        if not p.get("printer"):
            self.lbl_msg.config(text="还没选打印机 —— 在上面「打印机」下拉里选一台",
                                fg="#D93026")
            return
        try:
            copies = max(1, min(999, int(float(str(self.v_copies.get()).strip()))))
        except Exception:
            copies = 1
        task = self.app.enqueue_snapshot(p, copies)
        ch = ("图片模式" if task["mode"] == "image" else "TSPL 直发")
        self.lbl_msg.config(text="已加入队列：%d 张（%s）—— 进度看主窗口「队列」" % (copies, ch),
                            fg="#1F9D55")

    # ---------------- 关闭 ----------------
    def _on_close(self):
        for attr in ("_job", "_cfg_job", "_poll_job"):
            j = getattr(self, attr, None)
            if j:
                try:
                    self.after_cancel(j)
                except Exception:
                    pass
                setattr(self, attr, None)
        self.app._pv_win = None
        self.destroy()


def fit_window_to_screen(root):
    """启动后把窗口收进屏幕工作区，独立于标题栏实现。

    之前这活儿挂在自绘标题栏对象上，撤掉标题栏就没了 —— 但「换台小屏
    电脑窗口超屏」是真实存在的，所以单独留一个函数，谁都能调。
    只收进屏幕，不硬压内容；屏幕比 minsize 还小时先把下限降下来，
    至少保证左上角够得着。
    """
    import ctypes.wintypes as wt

    try:
        root.update_idletasks()
        work = wt.RECT()
        u = ctypes.windll.user32
        ok = u.SystemParametersInfoW(0x0030, 0, ctypes.byref(work), 0)
        if ok and work.right > work.left and work.bottom > work.top:
            left, top = work.left, work.top
            maxw, maxh = work.right - work.left, work.bottom - work.top
        else:
            left, top = 0, 0
            maxw, maxh = root.winfo_screenwidth(), root.winfo_screenheight()

        x, y = root.winfo_x(), root.winfo_y()
        w, h = root.winfo_width(), root.winfo_height()
        if w <= 1 or h <= 1:
            w, h = root.winfo_reqwidth(), root.winfo_reqheight()
        neww, newh = min(w, maxw), min(h, maxh)
        minw, minh = root.minsize()
        if neww < minw or newh < minh:
            root.minsize(min(neww, minw), min(newh, minh))
        newx = max(left, min(x, left + maxw - neww))
        newy = max(top, min(y, top + maxh - newh))
        if (neww, newh, newx, newy) != (w, h, x, y):
            root.geometry("%dx%d+%d+%d" % (neww, newh, newx, newy))
            return True
    except (tk.TclError, OSError, ValueError):
        pass
    return False


def main():
    path = None
    if len(sys.argv) > 1 and sys.argv[1].lower().endswith(EXT):
        path = sys.argv[1]
    root = tk.Tk()
    App(root, path)
    # 布局算完后把窗口收进屏幕：小屏 / 高缩放的电脑上，按内容自适应的
    # 尺寸可能超过屏幕。原生标题栏下这道保护依然值得留着。
    # 收两次：高 DPI 机器上字体度量会晚一点；两次幂等，已在屏内就直接返回。
    root.after(120, lambda: fit_window_to_screen(root))
    root.after(600, lambda: fit_window_to_screen(root))
    root.mainloop()


# ---------------- 打印分辨率和纸张预览 ----------------
_update_paper_line_with_driver = PrintPreview._update_paper_line
_update_info_with_driver = PrintPreview._update_info


def _update_paper_line(self):
    profile = eng.bitmap_printer_profile(self.v_printer.get())
    if not profile:
        return _update_paper_line_with_driver(self)
    width, length = eng.label_size_mm(self.app.proj)
    paper = self.app.proj["paper"]
    if paper.get("media") == "continuous":
        detail = "连续纸 · 多份按纸长相接"
    else:
        detail = f"间隙纸 · 传感器间隙 {float(paper.get('gap', 2)):g} mm"
    self._fit_ov = (0.0, 0.0)
    self.lbl_paper.config(text=f"{profile[0]} · {width:g} × {length:g} mm · {detail}", fg="#1F9D55")


def _update_info(self):
    _update_info_with_driver(self)
    profile = eng.bitmap_printer_profile(self.v_printer.get())
    if profile:
        lines = self.lbl_info.cget("text").splitlines()
        if lines and lines[-1].startswith("驱动纸型："):
            lines[-1] = f"打印分辨率：{profile[1]} dpi · 走纸按模板尺寸和纸张类型设置"
            self.lbl_info.config(text="\n".join(lines))


PrintPreview._update_paper_line = _update_paper_line
PrintPreview._update_info = _update_info

_build_ui_with_template_dpi = App._build_ui
_redraw_with_template_dpi = App.redraw
_update_channel_hint_with_template_dpi = App._update_channel_hint


def _refresh_print_dpi(self, *_):
    if not hasattr(self, "_paper_dpi_entry"):
        return
    profile = eng.bitmap_printer_profile(self.v_printer.get())
    dpi = profile[1] if profile else None
    if dpi == self._effective_print_dpi:
        return
    self._effective_print_dpi = dpi
    if dpi is None:
        self._paper_dpi_caption.configure(text="模板分辨率")
        self._paper_dpi_entry.configure(textvariable=self.v_dpi, state="normal")
    else:
        self._paper_dpi_caption.configure(text="打印分辨率")
        self._print_dpi_display.set(str(dpi))
        self._paper_dpi_entry.configure(textvariable=self._print_dpi_display, state="readonly")


def _build_ui(self):
    _build_ui_with_template_dpi(self)
    paper_frame = self.ent_gap.master
    self._paper_dpi_entry = next(
        widget for widget in paper_frame.winfo_children()
        if isinstance(widget, ttk.Entry) and str(widget.cget("textvariable")) == str(self.v_dpi)
    )
    entry_column = int(self._paper_dpi_entry.grid_info()["column"])
    self._paper_dpi_caption = paper_frame.grid_slaves(row=0, column=entry_column - 1)[0]
    # Keep the original variable bound to the template's design units.
    self._print_dpi_display = tk.StringVar(master=self.root)
    self._effective_print_dpi = 0
    self.v_printer.trace_add("write", self._refresh_print_dpi)
    self._refresh_print_dpi()


def _update_channel_hint(self):
    _update_channel_hint_with_template_dpi(self)
    self._refresh_print_dpi()


def _redraw(self):
    _redraw_with_template_dpi(self)
    paper = self.proj["paper"]
    width = float(paper["w"])
    length = float(paper["h"])
    template_dpi = float(paper["dpi"] or 203)
    old_caption = f"{width:g} × {length:g} mm   {int(template_dpi)} dpi"
    dpi = getattr(self, "_effective_print_dpi", None)
    caption = f"{width:g} × {length:g} mm   打印 {dpi} dpi" if dpi else (
        f"{width:g} × {length:g} mm   模板 {int(template_dpi)} dpi"
    )
    if dpi:
        printed_width, printed_length = eng.label_size_mm(self.proj)
        template_size = f"可打印 {eng.mm2dot(width, template_dpi)} × {eng.mm2dot(length, template_dpi)} dot"
        actual_size = f"打印画布 {eng.mm2dot(printed_width, dpi)} × {eng.mm2dot(printed_length, dpi)} dot"
    for item in self.canvas.find_all():
        if self.canvas.type(item) != "text":
            continue
        text = self.canvas.itemcget(item, "text")
        if text == old_caption:
            self.canvas.itemconfigure(item, text=caption)
        elif dpi and template_size in text:
            self.canvas.itemconfigure(item, text=text.replace(template_size, actual_size, 1))



App._refresh_print_dpi = _refresh_print_dpi
App._build_ui = _build_ui
App._update_channel_hint = _update_channel_hint
App.redraw = _redraw


# 应用主窗口、对话框与预览布局。
from PIL import Image, ImageTk

UI_GREEN = "#176849"
UI_GREEN_HOVER = "#12553d"
UI_INK = "#1c2730"
UI_MUTED = "#55616d"
UI_BORDER = "#d6dde3"
UI_CANVAS = "#edf1f4"
UI_WHITE = "#ffffff"
UI_ACTIVE_TAB = "#dceee5"
UI_WARNING_BG = "#fcebea"
UI_WARNING_INK = "#973a34"
UI_FONT = ("Microsoft YaHei UI", 10)
UI_HEAD = ("Microsoft YaHei UI", 12, "bold")
UI_TITLE = ("Microsoft YaHei UI", 16, "bold")
UI_TABLE_ROW_FONT = 24
PREVIEW_BLUE = UI_GREEN
PREVIEW_BLUE_D = UI_GREEN_HOVER


def _ui_theme(root):
    style = ttk.Style(root)
    style.theme_use("clam")
    root.configure(background=UI_WHITE)
    root.option_add("*TCombobox*Listbox.font", UI_FONT)
    style.configure(".", font=UI_FONT, foreground=UI_INK)
    for name in ("TFrame", "TLabel", "TCheckbutton", "TRadiobutton", "TLabelframe",
                 "TLabelframe.Label", "TNotebook"):
        style.configure(name, background=UI_WHITE)
    style.configure("TLabel", font=UI_FONT)
    style.configure("Quiet.TLabel", foreground=UI_MUTED)
    style.configure("Heading.TLabel", font=UI_HEAD)
    style.configure("Title.TLabel", font=UI_TITLE)
    style.configure("Canvas.TFrame", background=UI_CANVAS)
    style.configure("Canvas.TLabel", background=UI_CANVAS, font=UI_HEAD)
    style.configure("CanvasQuiet.TLabel", background=UI_CANVAS, foreground=UI_MUTED)
    style.configure("PreviewStatus.TLabel", background="#e7edf0", foreground=UI_MUTED, padding=(14, 6))
    style.configure("Warning.PreviewStatus.TLabel", background=UI_WARNING_BG, foreground=UI_WARNING_INK)
    style.configure("Document.TFrame", background="#f4f6f8")
    style.configure("Panel.TLabelframe", background=UI_WHITE, borderwidth=0, relief="flat", labelmargins=0)
    style.layout("Panel.TLabelframe.Label", [])
    style.configure("TSeparator", background=UI_BORDER)
    style.configure("TCheckbutton", indicatorsize=14, indicatormargin=(1, 1, 6, 1),
                    indicatorforeground=UI_GREEN, upperbordercolor="#75858c", lowerbordercolor="#75858c")
    for name in ("TEntry", "TCombobox", "TSpinbox"):
        style.configure(name, fieldbackground=UI_WHITE, background="#f4f6f8",
                        bordercolor=UI_BORDER, lightcolor=UI_WHITE, darkcolor=UI_WHITE,
                        padding=(7, 5), insertcolor=UI_INK)
        style.map(name, fieldbackground=[("disabled", "#f2f4f6"), ("readonly", "#f5f7f9")],
                  foreground=[("disabled", "#77828e"), ("readonly", UI_INK)],
                  bordercolor=[("focus", UI_GREEN)])
    style.configure("TCombobox", arrowsize=15, arrowcolor=UI_MUTED)
    style.configure("TSpinbox", arrowsize=11)
    for name, bg, fg in (("TButton", UI_WHITE, UI_INK), ("Primary.TButton", UI_GREEN, UI_WHITE),
                         ("Tab.TButton", UI_WHITE, UI_MUTED)):
        style.configure(name, background=bg, foreground=fg, bordercolor=UI_BORDER,
                        lightcolor=UI_BORDER if name == "TButton" else bg,
                        darkcolor=UI_BORDER if name == "TButton" else bg,
                        borderwidth=1,
                        relief="solid" if name in ("TButton", "Primary.TButton") else "flat",
                        padding=(10, 6))
        style.map(name, background=[("disabled", "#eef1f3"), ("active", UI_GREEN_HOVER if name == "Primary.TButton" else "#eaf3ef")],
                  foreground=[("disabled", "#77828e")], relief=[("pressed", "flat")])
    outline = "#729787"
    style.configure("Accent.TButton", background="#f8fbf9", foreground=UI_GREEN,
                    bordercolor=outline, lightcolor=outline, darkcolor=outline,
                    borderwidth=1, relief="solid", font=(*UI_FONT, "bold"))
    style.map("Accent.TButton", background=[("disabled", "#eef1f3"), ("pressed", "#bedfcd"),
                                             ("active", "#eaf3ef")])
    style.configure("Current.Tab.TButton", background=UI_ACTIVE_TAB, foreground=UI_GREEN,
                    lightcolor=UI_ACTIVE_TAB, darkcolor=UI_ACTIVE_TAB, font=(*UI_FONT, "bold"))
    style.map("Current.Tab.TButton", background=[("active", UI_ACTIVE_TAB)])
    style.configure("Table.TEntry", padding=(6, 3))
    style.configure("Table.TCombobox", padding=(6, 3))
    style.configure("Table.TButton", padding=(8, 3))
    for prefix, color in (("Row", UI_WHITE), ("AlternateRow", "#f2f6f8")):
        style.configure(f"{prefix}.TFrame", background=color)
        style.configure(f"{prefix}.TLabel", background=color)
    style.configure("TMenubutton", background=UI_WHITE, foreground=UI_INK,
                    bordercolor=UI_BORDER, lightcolor=UI_BORDER, darkcolor=UI_BORDER,
                    relief="solid", borderwidth=1, padding=(9, 5))
    style.configure("Primary.TButton", font=UI_HEAD, bordercolor=UI_GREEN, padding=(17, 7))
    style.configure("Treeview", background=UI_WHITE, fieldbackground=UI_WHITE,
                    bordercolor=UI_BORDER, rowheight=34, font=UI_FONT)
    style.configure("Treeview.Heading", background="#f1f4f6", foreground=UI_INK,
                    relief="flat", font=("Microsoft YaHei UI", 10, "bold"), padding=(8, 7))
    style.map("Treeview", background=[("selected", "#e6f2ec")], foreground=[("selected", UI_GREEN)])
    style.configure("Router.TNotebook", borderwidth=0, padding=0)
    style.layout("Router.TNotebook.Tab", [])
    side_scrollbar = "Vertical.TScrollbar"
    layout = style.layout(side_scrollbar)
    if layout[0][0] != "padding":
        style.layout(side_scrollbar, [("padding", {"sticky": "nswe", "children": layout})])
    # Clam measures a 10 px arrow as an 11 px control. A 20 px arrow
    # and one pixel of outer padding produce the intended 22 px width.
    style.configure(side_scrollbar, background="#e7ebef", troughcolor=UI_WHITE,
                    arrowsize=20, padding=(0, 0, 1, 0))


def _ui_button(parent, title, command, primary=False, **kwargs):
    kwargs.setdefault("style", "Primary.TButton" if primary else "TButton")
    return ttk.Button(parent, text=title, command=command, **kwargs)


def _ui_heading(parent, title):
    ttk.Label(parent, text=title, style="Heading.TLabel").pack(anchor="w", pady=(0, 8))


def _ui_field(parent, title, variable, unit="", choices=None):
    row = ttk.Frame(parent)
    row.pack(fill="x", pady=(0, 7))
    row.columnconfigure(1, weight=1)
    label = ttk.Label(row, text=title, width=11, anchor="w")
    label.grid(row=0, column=0, sticky="w", padx=(0, 8))
    if choices is None:
        widget = ttk.Entry(row, textvariable=variable, width=10)
    else:
        widget = ttk.Combobox(row, textvariable=variable, values=choices,
                              state="readonly", width=10)
    widget.grid(row=0, column=1, sticky="ew")
    ttk.Label(row, text=unit, width=4, style="Quiet.TLabel").grid(row=0, column=2, padx=(8, 0), sticky="w")
    return label, widget


def _ui_walk(root):
    for widget in root.winfo_children():
        if not isinstance(widget, tk.Toplevel):
            yield widget
            yield from _ui_walk(widget)


def _ui_reserve_side_scrollbar(scrollbar):
    if str(scrollbar.cget("orient")) != "vertical" or scrollbar.winfo_manager() != "pack":
        return
    packing = scrollbar.pack_info()
    for sibling in packing["in"].pack_slaves():
        if sibling is scrollbar:
            return
        sibling_packing = sibling.pack_info()
        if sibling_packing["side"] in ("left", "right") and sibling_packing["expand"]:
            scrollbar.pack_configure(before=sibling)
            return


def _ui_polish(root):
    for widget in _ui_walk(root):
        if isinstance(widget, ttk.Scrollbar):
            _ui_reserve_side_scrollbar(widget)
        elif isinstance(widget, ttk.Treeview):
            name = widget.cget("style") or "Treeview"
            ttk.Style(root).configure(name, rowheight=34)
        elif isinstance(widget, tk.Text):
            widget.configure(background=UI_WHITE, foreground=UI_INK, insertbackground=UI_INK,
                             font=UI_FONT, relief="flat", borderwidth=0, highlightthickness=1,
                             highlightbackground=UI_BORDER, highlightcolor=UI_GREEN, padx=7, pady=5)
        elif isinstance(widget, tk.Listbox):
            widget.configure(background=UI_WHITE, foreground=UI_INK, font=UI_FONT,
                             relief="flat", borderwidth=0, highlightthickness=1,
                             highlightbackground=UI_BORDER, highlightcolor=UI_GREEN,
                             selectbackground="#e6f2ec", selectforeground=UI_GREEN)
        elif isinstance(widget, tk.Button):
            old = str(widget.cget("background")).lower()
            strong = old in ("#3c8c4e", "#1f6f45", "#126446", "#176849")
            bg = UI_GREEN if strong else UI_WHITE
            fg = UI_WHITE if strong else UI_INK
            widget.configure(background=bg, foreground=fg, activeforeground=fg,
                             activebackground=UI_GREEN_HOVER if strong else "#eaf3ef",
                             font=UI_FONT, borderwidth=0, relief="flat", padx=11, pady=5)
        elif isinstance(widget, ttk.Checkbutton):
            widget.configure(cursor="hand2")


def _ui_panel(parent):
    # Keep long controls reachable in a short window.
    canvas = tk.Canvas(parent, background=UI_WHITE, borderwidth=0, highlightthickness=0)
    scroll = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=scroll.set)
    scroll.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)
    panel = ttk.Frame(canvas, padding=(20, 14, 20, 14))
    item = canvas.create_window(0, 0, anchor="nw", window=panel)
    panel.bind("<Configure>", lambda _: canvas.configure(scrollregion=canvas.bbox("all")), add="+")
    canvas.bind("<Configure>", lambda event: canvas.itemconfigure(item, width=event.width), add="+")
    canvas.bind("<MouseWheel>", lambda event: canvas.yview_scroll(-int(event.delta / 120), "units"), add="+")
    return panel


def _ui_divider(parent, pady=(5, 12)):
    ttk.Separator(parent, orient="horizontal").pack(fill="x", pady=pady)


def _ui_show_page(self, page):
    self.nb.select(self._ui_pages[page])


def _ui_sync_navigation(self, _event=None):
    if not hasattr(self, "_ui_pages"):
        return
    current = self.nb.select()
    group = next((key for key, widget in self._ui_pages.items() if str(widget) == current), "内容")
    if group == "洗涤说明":
        group = "内容"
    for key, button in self._ui_nav.items():
        active = key == group
        idle_style = "Accent.TButton" if key == "纸张与版式" else "Tab.TButton"
        button.configure(style="Current.Tab.TButton" if active else idle_style)
        button.state(["selected" if active else "!selected"])
        self._ui_nav_lines[key].configure(background=UI_GREEN if active else UI_WHITE)


def _ui_new_label(self):
    defaults = eng.normalize({})
    dialog = tk.Toplevel(self.root, background=UI_WHITE)
    dialog.transient(self.root)
    dialog.title("新建标签")
    dialog.resizable(False, False)
    name = tk.StringVar(dialog, value=defaults["name"])
    width = tk.StringVar(dialog, value=f'{defaults["paper"]["w"]:g}')
    height = tk.StringVar(dialog, value=f'{defaults["paper"]["h"]:g}')
    auto = tk.BooleanVar(dialog, value=defaults["auto_paper"])
    body = ttk.Frame(dialog, padding=(22, 16, 22, 18))
    body.pack(fill="both", expand=True)
    _ui_heading(body, "新建标签")
    _ui_divider(body)
    columns = ttk.Frame(body)
    columns.pack(fill="both", expand=True)
    form = ttk.Frame(columns)
    form.pack(side="left", fill="both", expand=True, padx=(0, 18))
    _, name_entry = _ui_field(form, "标签名称", name)
    _ui_field(form, "纸宽", width, "mm")
    _ui_field(form, "纸长", height, "mm")
    ttk.Label(form, text="连续纸 · 300 dpi · TSPL 直发", style="Quiet.TLabel").pack(anchor="w", pady=(5, 7))
    ttk.Checkbutton(form, text="内容超长时自动加长纸", variable=auto).pack(anchor="w")
    warning = ttk.Label(form, foreground="#a13932", wraplength=275)
    warning.pack(anchor="w", pady=(5, 0))
    cv = tk.Canvas(columns, width=184, height=230, bg=UI_CANVAS, highlightthickness=0)
    cv.pack(side="right", fill="y")

    def dimensions():
        try:
            w, h = float(width.get()), float(height.get())
            if not (math.isfinite(w) and math.isfinite(h) and w >= 1 and h >= 1):
                return None
            return w, h
        except ValueError:
            return None

    def draw(*_):
        cv.delete("all")
        size = dimensions()
        if size:
            w, h = size
            scale = min(100 / w, 162 / h)
            pw, ph = w * scale, h * scale
            x, y = (165 - pw) / 2, (247 - ph) / 2
            cv.create_rectangle(x, y, x + pw, y + ph, fill=UI_WHITE, outline=UI_BORDER)
            cv.create_text(83, y - 14, text=f"{w:g} mm", fill=UI_MUTED, font=UI_FONT)
            cv.create_text(min(177, x + pw + 18), y + ph / 2, text=f"{h:g} mm", angle=90,
                           fill=UI_MUTED, font=UI_FONT)

    def create():
        size = dimensions()
        if size is None:
            warning.configure(text="纸宽和纸长都需要填写至少 1 mm 的数字。")
            return
        self._new_tab()
        self.proj["name"] = name.get().strip() or defaults["name"]
        self.proj["paper"].update(w=size[0], h=size[1], h_base=size[1])
        self.proj["auto_paper"] = auto.get()
        self._sync_to_ui()
        self._refresh_tabs()
        self.redraw()
        dialog.destroy()

    for value in (width, height):
        value.trace_add("write", draw)
    draw()
    _ui_divider(body, (14, 10))
    footer = ttk.Frame(body)
    footer.pack(fill="x")
    _ui_button(footer, "新建", create, primary=True).pack(side="right")
    _ui_button(footer, "取消", dialog.destroy).pack(side="right", padx=8)
    dialog.bind("<Return>", lambda _: create())
    dialog.bind("<Escape>", lambda _: dialog.destroy())
    dialog.update_idletasks()
    _ui_center(dialog, self.root)
    dialog.grab_set()
    name_entry.focus_set()
    name_entry.select_range(0, "end")


def _ui_center(window, parent):
    window.update_idletasks()
    width = max(window.winfo_reqwidth(), window.winfo_width())
    height = max(window.winfo_reqheight(), window.winfo_height())
    x = max(0, parent.winfo_rootx() + (parent.winfo_width() - width) // 2)
    y = max(0, parent.winfo_rooty() + (parent.winfo_height() - height) // 2)
    window.geometry(f"+{x}+{y}")


_ui_existing_build = App._build_ui
_ui_existing_tabs = App._refresh_tabs
_ui_existing_redraw = App.redraw
_ui_existing_status = App._update_status_ctx


def _ui_tabs(self):
    _ui_existing_tabs(self)
    for item in self.tabstrip.winfo_children():
        if isinstance(item, tk.Button):
            text = item.cget("text")
            if text.strip() == "＋ 新建":
                item.pack_forget()
                continue
            chosen = text.lstrip().startswith("▶")
            item.configure(text=text.removeprefix("▶").strip(), font=UI_FONT, borderwidth=0,
                           relief="flat", background=UI_WHITE if chosen else "#f4f6f8",
                           foreground=UI_GREEN if chosen else UI_MUTED,
                           activebackground="#e7f3ed", activeforeground=UI_GREEN, padx=12, pady=3)
    self.tabstrip.configure(style="Document.TFrame")


def _ui_paper_page(self, wrap, original_font_values):
    panel = _ui_panel(wrap)
    _ui_heading(panel, "纸张")
    _ui_field(panel, "纸宽", self.v_w, "mm")
    _, self.ent_h = _ui_field(panel, "纸长", self.v_h, "mm")
    _, self.cb_media = _ui_field(panel, "纸张类型", self.v_media, choices=list(MEDIA_NAME))
    self.cb_media.bind("<<ComboboxSelected>>", lambda _: self._sync_media_row(), add="+")
    _, self.ent_gap = _ui_field(panel, "间隙", self.v_gap, "mm")
    self._paper_dpi_caption, self._paper_dpi_entry = _ui_field(panel, "模板分辨率", self.v_dpi, "dpi")
    ttk.Checkbutton(panel, text="内容超长时自动加长纸", variable=self.v_autoh,
                    command=self._on_autoh).pack(anchor="w", pady=(0, 5))
    self.lbl_media_hint = ttk.Label(panel, style="Quiet.TLabel")
    self.lbl_media_hint.pack(anchor="w")
    _ui_divider(panel)
    _ui_heading(panel, "版式")
    _ui_field(panel, "左边距", self.v_left, "mm")
    _ui_field(panel, "右边距", self.v_right, "mm")
    _ui_field(panel, "上方留白", self.v_starty, "mm")
    _ui_field(panel, "正文缩进", self.v_indent, "mm")
    _ui_field(panel, "正文字号", self.v_font, "dot")
    _ui_field(panel, "行间距", self.v_lg, "dot")
    _, self.cb_font = _ui_field(panel, "字体", self.v_fontsel, choices=original_font_values)
    _ui_button(panel, "浏览字体…", self._pick_font).pack(anchor="e")
    self.lbl_fontwarn = ttk.Label(panel, style="Quiet.TLabel", wraplength=325)
    self.lbl_fontwarn.pack(fill="x", pady=(3, 5))
    ttk.Checkbutton(panel, text="显示顶部装饰线", variable=self.v_toprule).pack(anchor="w", pady=3)
    ttk.Checkbutton(panel, text="显示正文分隔线", variable=self.v_div).pack(anchor="w", pady=3)
    _ui_field(panel, "分隔线宽度", self.v_divw, "mm")
    _ui_field(panel, "距正文", self.v_divg, "mm")


def _ui_print_page(self, wrap):
    panel = _ui_panel(wrap)
    _ui_heading(panel, "打印设置")
    device = ttk.Frame(panel)
    device.pack(fill="x", pady=(0, 8))
    ttk.Label(device, text="打印机", width=13).pack(side="left", padx=(0, 10))
    self.lbl_printer = tk.Label(device, bg="#f3f6f7", fg=UI_INK, relief="flat",
                                padx=6, pady=6, anchor="w", font=UI_FONT)
    _ui_button(device, "切换…", self.choose_printer).pack(side="right", padx=(8, 0))
    self.lbl_printer.pack(side="left", fill="x", expand=True)
    self.lbl_printer.bind("<Button-1>", lambda _: self.choose_printer())
    self.lbl_printer_info = ttk.Label(panel, style="Quiet.TLabel", wraplength=340)
    self.lbl_printer_info.pack(anchor="w", pady=(0, 8))
    self._ui_print_dpi = tk.StringVar(panel)
    _, dpi = _ui_field(panel, "打印分辨率", self._ui_print_dpi, "dpi")
    dpi.configure(state="readonly")
    _ui_field(panel, "打印方式", self.v_mode, choices=list(MODE_NAME))
    direction = ttk.Frame(panel)
    direction.pack(fill="x", pady=(0, 8))
    ttk.Label(direction, text="打印方向", width=13).pack(side="left", padx=(0, 10))
    for angle in (0, 90, 180, 270):
        ttk.Radiobutton(direction, text=f"{angle}°", value=str(angle),
                        variable=self.v_direction, command=lambda: self._collect()).pack(side="left", padx=(0, 9))
    _ui_field(panel, "水平偏移", self.v_offx, "mm")
    _ui_field(panel, "垂直偏移", self.v_offy, "mm")
    self.lbl_channel = ttk.Label(panel, style="Quiet.TLabel", wraplength=340)
    self.lbl_channel.pack(fill="x", pady=3)
    _ui_divider(panel)
    _ui_heading(panel, "打印工具")
    tools = ttk.Frame(panel)
    tools.pack(fill="x")
    tools.columnconfigure((0, 1), weight=1, uniform="print-tools")
    actions = (("批量打印", self.open_batch), ("打印队列", self.open_queue),
               ("自检页", self.do_selftest), ("TSPL 指令", self.show_tspl))
    for index, (label, action) in enumerate(actions):
        _ui_button(tools, label, action).grid(row=index // 2, column=index % 2,
                                             sticky="ew", padx=(0, 5), pady=(0, 5))


def _ui_content(self, content, wash):
    form = self.txt_fabric.master
    for child in form.winfo_children():
        if isinstance(child, ttk.Label):
            text = child.cget("text")
            if text == "号型 / 第一行":
                child.configure(text="号型")
            elif text.startswith(("产地留空", "自定义填了")):
                child.pack_forget()
    spacing_widgets = (self.cb_font.master,)
    for widget in spacing_widgets:
        widget.pack_forget()
    for child in form.winfo_children():
        if isinstance(child, ttk.Separator):
            child.pack_forget()
        if isinstance(child, ttk.Checkbutton) and child.cget("text").startswith("顶部装饰线"):
            child.pack_forget()
        if isinstance(child, ttk.Frame):
            if any(isinstance(w, ttk.Checkbutton) and str(w.cget("variable")) == str(self.v_div)
                   for w in child.winfo_children()):
                child.pack_forget()
            for inner in child.winfo_children():
                if isinstance(inner, ttk.Label) and inner.cget("text").startswith("成分（"):
                    inner.configure(text="成分")
    self.txt_fabric.configure(height=4)
    form.configure(padding=(18, 13, 18, 16))
    optional = [
        widget for widget in form.pack_slaves()
        if widget in (self._cb_datemon.master, self._ent_datecustom.master)
        or (isinstance(widget, ttk.Checkbutton) and str(widget.cget("variable")) == str(self.v_brand))
    ]
    optional_packing = [(widget, widget.pack_info()) for widget in optional]
    for widget in optional:
        widget.pack_forget()
    _ui_divider(form)
    action = ttk.Frame(form)
    action.pack(fill="x")
    _ui_button(action, "洗涤说明  ›", lambda: self._ui_show_page("洗涤说明"), width=-7).pack(side="left")
    _ui_button(action, "表格工具", self.open_table_editor, style="Accent.TButton", width=-7).pack(side="left", padx=6)
    _ui_button(action, "导入模板", self.do_import_ddl, style="Accent.TButton", width=-7).pack(side="left")

    def toggle_optional():
        if optional[0].winfo_manager():
            for widget in optional:
                widget.pack_forget()
            more.configure(text="更多内容设置  ›")
        else:
            for widget, stored in optional_packing:
                options = {key: value for key, value in stored.items() if key != "in"}
                widget.pack(**options)
            more.configure(text="收起内容设置")

    more = _ui_button(form, "更多内容设置  ›", toggle_optional)
    more.pack(anchor="w", pady=(9, 5))
    toggle_optional()
    wash_form = self.txt_bcaution.master
    wash_form.configure(padding=(18, 13, 18, 16))
    first = next(iter(wash_form.pack_slaves()))
    _ui_button(wash_form, "‹ 返回内容", lambda: self._ui_show_page("内容")).pack(
        anchor="w", before=first, pady=(0, 8))
    for child in wash_form.winfo_children():
        if isinstance(child, ttk.Label) and child.cget("text").startswith(("字号 12", "顺序即打印", "一排最多")):
            child.pack_forget()
    for pane in (content, wash):
        for widget in pane.winfo_children():
            if isinstance(widget, tk.Canvas):
                widget.configure(background=UI_WHITE, borderwidth=0, highlightthickness=0)


def _ui_build(self):
    _ui_theme(self.root)
    _ui_existing_build(self)
    root = self.root
    root.title(APP_NAME)
    root.minsize(840, 580)
    original_paper = self.ent_gap.master
    original_printer = self.lbl_printer.master.master
    original_font_values = self.cb_font.cget("values")
    left = self.canvas.master
    body = left.master
    right = self.nb.master
    content, wash = (self.nb.nametowidget(name) for name in self.nb.tabs())
    self._ui_pages = {"内容": content, "洗涤说明": wash}
    self._ui_nav = {}
    self._ui_nav_lines = {}
    original_paper.pack_forget()
    original_printer.pack_forget()
    self._right_tools.place_forget()
    self.bar.pack_forget()
    self._foot.pack_forget()
    self.tabstrip.pack_forget()
    body.pack_forget()
    header = ttk.Frame(root, padding=(17, 6))
    header.pack(fill="x")
    icon_path = eng.asset_path("waterlabel-app.ico")
    with Image.open(icon_path) as image:
        self._ui_app_mark = ImageTk.PhotoImage(image.resize((29, 29), Image.Resampling.LANCZOS), master=root)
    root.iconbitmap(default=icon_path)
    root.iconbitmap(icon_path)
    ttk.Label(header, image=self._ui_app_mark, padding=0).pack(side="left", padx=(0, 9))
    ttk.Label(header, text=APP_NAME, style="Title.TLabel").pack(side="left")
    menu_button = ttk.Menubutton(header, text="更多", direction="below")
    menu = tk.Menu(menu_button, tearoff=False, font=UI_FONT)
    for title, command in (("另存为…", lambda: self.do_save(as_new=True)),
                           ("导入模板", self.do_import_ddl), ("撤销", self.undo),
                           ("重做", self.redo), ("环境检查", lambda: self.check_env(verbose=True)),
                           ("自检页", self.do_selftest), ("查看 TSPL 指令", self.show_tspl)):
        menu.add_command(label=title, command=command)
    menu_button.configure(menu=menu)
    menu_button.pack(side="right", padx=(7, 0))
    for title, command in (("保存", self.do_save), ("打开", self.do_open), ("新建", self._ui_new_label)):
        _ui_button(header, title, command).pack(side="right", padx=(7, 0))
    self.tabstrip.pack(fill="x")
    _ui_divider(root, (0, 0))
    footer = ttk.Frame(root, padding=(16, 9))
    footer.pack(side="bottom", fill="x")
    self._ui_footer = ttk.Label(footer, style="Quiet.TLabel")
    self._ui_footer.pack(side="left", fill="x", expand=True, anchor="w")
    _ui_button(footer, "打印", self.do_print, primary=True).pack(side="right")
    _ui_button(footer, "打印预览", self.open_print_preview).pack(side="right", padx=8)
    ttk.Spinbox(footer, from_=1, to=999, width=5, textvariable=self.v_copies,
                justify="center").pack(side="right", padx=(5, 0))
    ttk.Label(footer, text="份数", style="Quiet.TLabel").pack(side="right")
    ttk.Separator(root, orient="horizontal").pack(side="bottom", fill="x")
    self._foot.configure(padding=(12, 3))
    self.lbl_ctx.pack_forget()
    self.lbl_env.pack_forget()
    self.lbl_taskhint.pack_forget()
    self._foot.pack(side="bottom", fill="x")
    left.pack_forget()
    right.pack_forget()
    body.pack(fill="both", expand=True)
    body.rowconfigure(0, weight=1)
    body.columnconfigure(0, weight=12, uniform="main")
    body.columnconfigure(1, weight=10, uniform="main")
    left.configure(text="", padding=0, style="Panel.TLabelframe")
    right.configure(text="", padding=0, style="Panel.TLabelframe")
    left.grid(row=0, column=0, sticky="nsew")
    right.grid(row=0, column=1, sticky="nsew", padx=(1, 0))
    self.canvas.configure(background=UI_CANVAS, highlightthickness=0)
    canvas_top = ttk.Frame(left, style="Canvas.TFrame", padding=(14, 7))
    canvas_top.pack(fill="x", before=self.canvas)
    ttk.Label(canvas_top, text="标签预览", style="Canvas.TLabel").pack(side="left")
    self._ui_dimensions = ttk.Label(canvas_top, style="CanvasQuiet.TLabel")
    self._ui_dimensions.pack(side="left", padx=11)
    self._ui_zoom = tk.StringVar(root, value="适应窗口")
    zoom = ttk.Combobox(canvas_top, width=10, state="readonly", textvariable=self._ui_zoom,
                        values=("适应窗口", "75%", "100%", "125%", "150%"))
    zoom.pack(side="right")
    zoom.bind("<<ComboboxSelected>>", lambda _: self.redraw())
    self._ui_preview_status = ttk.Label(left, style="PreviewStatus.TLabel")
    self._ui_preview_status.pack(side="bottom", fill="x", before=self.canvas)
    nav = ttk.Frame(right, padding=(12, 5, 12, 0))
    nav.pack(fill="x", before=self.nb)
    for title in ("内容", "纸张与版式", "打印"):
        slot = ttk.Frame(nav)
        slot.pack(side="left", padx=(0, 6))
        style = "Accent.TButton" if title == "纸张与版式" else "Tab.TButton"
        button = ttk.Button(slot, text=title, style=style, width=-8,
                            command=lambda value=title: self._ui_show_page(value))
        button.pack(fill="x")
        self._ui_nav[title] = button
        underline = tk.Frame(slot, height=3, background=UI_WHITE)
        underline.pack(fill="x")
        self._ui_nav_lines[title] = underline
    line = ttk.Separator(right)
    line.pack(fill="x", before=self.nb)
    self.nb.configure(style="Router.TNotebook")
    self._ui_content(content, wash)
    for title, make in (("纸张与版式", lambda page: self._ui_paper_page(page, original_font_values)),
                        ("打印", self._ui_print_page)):
        page = ttk.Frame(self.nb)
        self._ui_pages[title] = page
        self.nb.add(page, text=title)
        make(page)
    self.nb.bind("<<NotebookTabChanged>>", self._ui_sync_navigation, add="+")
    self._ui_show_page("内容")
    self._ui_sync_navigation()
    self._effective_print_dpi = 0
    self._refresh_print_dpi()
    self.refresh_printers()
    self._sync_media_row()
    self._update_status_ctx()
    self._refresh_tabs()
    _ui_polish(root)


def _ui_status(self):
    _ui_existing_status(self)
    if not hasattr(self, "_ui_print_dpi"):
        return
    paper = self.proj["paper"]
    name = self.v_printer.get().strip()
    profile = eng.bitmap_printer_profile(name)
    dpi = profile[1] if profile else paper.get("dpi", 300)
    media = MEDIA_BY_VALUE.get(paper.get("media"), "间隙纸")
    mode = "TSPL 直发" if eng.resolve_mode(self.proj) == "tspl" else "图片模式"
    shown_name = profile[0] if profile else (name or "未选打印机")
    self._ui_print_dpi.set(f"{dpi:g}" if name else "—")
    self._ui_footer.configure(text=f"{shown_name}  ·  {media}  ·  {mode}")
    self._ui_dimensions.configure(text=f'{paper["w"]:g} × {paper["h"]:g} mm')


def _ui_draw(self):
    if not hasattr(self, "_ui_zoom"):
        return _ui_existing_redraw(self)
    paper = self.proj["paper"]
    canvas = self.canvas
    cw, ch = canvas.winfo_width(), canvas.winfo_height()
    if cw < 40 or ch < 50:
        return
    self._update_status_ctx()
    dpi = getattr(self, "_effective_print_dpi", None)
    dpi_text = f"打印 {dpi:g} dpi" if dpi else f'模板 {paper["dpi"]:g} dpi'
    length_over = eng.page_overflow_mm(self.proj)
    width_over = eng.page_side_overflow_mm(self.proj)
    warnings = []
    if length_over > .05:
        warnings.append(f"内容超长 {length_over:.1f} mm")
    if width_over > .05:
        warnings.append(f"内容超宽 {width_over:.1f} mm")
    self._ui_preview_status.configure(text="  ·  ".join([dpi_text, *warnings]),
                                      style="Warning.PreviewStatus.TLabel" if warnings else "PreviewStatus.TLabel")
    key = json.dumps(self.proj, sort_keys=True), dpi
    try:
        if getattr(self, "_ui_image_key", None) != key:
            self._ui_label_image = eng.render_label(self.proj, dpi=min(300, dpi or int(paper["dpi"])))
            self._ui_image_key = key
        img = self._ui_label_image
        fit = min((cw - 46) / img.width, (ch - 28) / img.height)
        zoom = {"适应窗口": 1, "75%": .75, "100%": 1, "125%": 1.25, "150%": 1.5}[self._ui_zoom.get()]
        nw, nh = max(1, round(img.width * fit * zoom)), max(1, round(img.height * fit * zoom))
        self._ui_canvas_photo = ImageTk.PhotoImage(img.resize((nw, nh), Image.Resampling.LANCZOS))
    except (OSError, ValueError, OverflowError):
        return _ui_existing_redraw(self)
    x, y = (cw - nw) / 2, (ch - nh) / 2
    canvas.delete("all")
    canvas.create_rectangle(x + 2, y + 3, x + nw + 2, y + nh + 3, fill="#dfe5e9", outline="")
    canvas.create_image(x, y, anchor="nw", image=self._ui_canvas_photo)
    canvas.create_rectangle(x, y, x + nw, y + nh, outline="#cbd4db")


_ui_sync_legacy_scrollbar = App._sync_scrollbar
_ui_build_legacy_printer_picker = ppick.PrinterPicker._build


def _ui_sync_scrollbar(self, canvas, vbar, need=None):
    _ui_sync_legacy_scrollbar(self, canvas, vbar, need)
    if vbar.winfo_exists():
        _ui_reserve_side_scrollbar(vbar)


def _ui_build_printer_picker(self):
    _ui_build_legacy_printer_picker(self)
    for widget in self.tree.master.pack_slaves():
        if isinstance(widget, ttk.Scrollbar):
            _ui_reserve_side_scrollbar(widget)


App._build_ui = _ui_build
App._refresh_tabs = _ui_tabs
App._update_status_ctx = _ui_status
App._ui_show_page = _ui_show_page
App._ui_sync_navigation = _ui_sync_navigation
App._ui_paper_page = _ui_paper_page
App._ui_print_page = _ui_print_page
App._ui_content = _ui_content
App._ui_new_label = _ui_new_label
App._sync_scrollbar = _ui_sync_scrollbar
App.redraw = _ui_draw
App._layout_right_buttons = lambda self: None
App._reflow_bar = lambda self: None
ppick.PrinterPicker._build = _ui_build_printer_picker


def _ui_dialog_header(window, title, detail=""):
    if getattr(window, "_ui_header", None):
        return
    window.configure(background=UI_WHITE)
    header = ttk.Frame(window, padding=(18, 12, 18, 8))
    first = next(iter(window.pack_slaves()), None)
    header.pack(fill="x", before=first) if first else header.pack(fill="x")
    title_row = ttk.Frame(header)
    title_row.pack(fill="x")
    ttk.Label(title_row, text=title, style="Title.TLabel").pack(side="left")
    if detail:
        ttk.Label(title_row, text=detail, style="Quiet.TLabel").pack(side="right")
    _ui_divider(header, (8, 0))
    window._ui_header = header


def _ui_skin_dialog(window, title, primary=(), headings=()):
    _ui_dialog_header(window, title)
    _ui_polish(window)
    for widget in _ui_walk(window):
        if isinstance(widget, tk.Canvas):
            widget.configure(background=UI_CANVAS, highlightthickness=0)
        if isinstance(widget, ttk.Label) and widget.cget("text") in headings:
            widget.configure(style="Heading.TLabel")
        if isinstance(widget, (ttk.Button, tk.Button)):
            label = widget.cget("text").strip()
            if label in primary:
                if isinstance(widget, ttk.Button):
                    widget.configure(style="Primary.TButton")
                else:
                    widget.configure(bg=UI_GREEN, fg=UI_WHITE, activebackground=UI_GREEN_HOVER,
                                     activeforeground=UI_WHITE, padx=16, pady=6)


def _ui_render_card(canvas, image, width, height, *, zoom=1):
    scale = zoom * min(max(1, width - 18) / image.width, max(1, height - 22) / image.height)
    nw, nh = max(1, round(image.width * scale)), max(1, round(image.height * scale))
    x, y = (width - nw) / 2, (height - nh) / 2
    canvas._ui_photo = ImageTk.PhotoImage(image.resize((nw, nh), Image.Resampling.LANCZOS))
    canvas.create_image(x, y, image=canvas._ui_photo, anchor="nw")
    canvas.create_rectangle(x, y, x + nw, y + nh, outline=UI_BORDER)
    return x, y, nw, nh


_ui_preview_base = PrintPreview.__init__
_ui_preview_draw_base = PrintPreview._redraw_canvas
_ui_preview_copies_base = PrintPreview._on_copies
_ui_preview_info_base = PrintPreview._update_info


def _ui_preview_init(self, app):
    _ui_preview_base(self, app)
    _ui_dialog_header(self, "打印预览")
    left = self.cv.master.master
    right = self.btn_print.master.master
    outer = left.master
    separator = next(item for item in outer.pack_slaves() if item not in (left, right))
    for pane in (left, right, separator):
        pane.pack_forget()
    right.configure(width=300)
    right.pack(side="right", fill="y")
    separator.configure(background=UI_BORDER)
    separator.pack(side="right", fill="y")
    left.pack(side="left", fill="both", expand=True)
    for widget in _ui_walk(self):
        if isinstance(widget, (tk.Frame, tk.Label)) and widget.cget("background") in ("#F4F5F7", "#EDEFF2"):
            widget.configure(background=UI_CANVAS)
    _ui_polish(self)
    for button in self.dir_btns.values():
        button.configure(activebackground="#e7f3ed")
    self.btn_print.configure(background=UI_GREEN, activebackground=UI_GREEN_HOVER,
                             foreground=UI_WHITE, activeforeground=UI_WHITE, padx=16, pady=8)
    _ui_button(self.btn_print.master, "返回", self._on_close).pack(fill="x", pady=(6, 0))
    self.cv.configure(background=UI_CANVAS, highlightthickness=0)
    self._ui_roll_preview = True
    self._update_info()
    self._schedule_paint()


def _ui_preview_count(self):
    try:
        return min(999, max(1, int(float(self.v_copies.get()))))
    except (ValueError, OverflowError):
        return 1


def _ui_preview_info(self):
    _ui_preview_info_base(self)
    paper = self.app.proj["paper"]
    width, length = eng.label_size_mm(self.app.proj)
    profile = eng.bitmap_printer_profile(self.v_printer.get())
    dpi = profile[1] if profile else paper.get("dpi", 300)
    media = MEDIA_BY_VALUE.get(paper.get("media"), "间隙纸")
    count = _ui_preview_count(self)
    name = self.app.proj.get("name") or "水洗标"
    pitch = length if paper.get("media") == "continuous" else length + float(paper.get("gap", 0))
    self.lbl_info.configure(text=f"{name}  ·  {width:g} × {length:g} mm\n"
                                 f"{dpi:g} dpi  ·  {media}  ·  {count} 张"
                                 f"  ·  理论走纸 {count * pitch:g} mm")


def _ui_preview_copies(self):
    _ui_preview_copies_base(self)
    if getattr(self, "_ui_roll_preview", False):
        self._update_info()
        self._schedule_paint()


def _ui_preview_draw(self):
    actual_count = _ui_preview_count(self)
    if (not getattr(self, "_ui_roll_preview", False) or self._img is None or actual_count == 1
            or self.app.proj["paper"].get("media") != "continuous"):
        return _ui_preview_draw_base(self)
    if self._cfg_job is not None:
        self.after_cancel(self._cfg_job)
        self._cfg_job = None
    canvas = self.cv
    cw, ch = canvas.winfo_width(), canvas.winfo_height()
    if min(cw, ch) < 50:
        return
    count = min(actual_count, 3)
    image = self._img
    img = Image.new("L", (image.width, image.height * count), 255)
    for index in range(count):
        img.paste(image, (0, index * image.height))
    canvas.delete("all")
    self._paint_key = None
    x, y, width, height = _ui_render_card(canvas, img, cw - 36, ch - 18, zoom=self.zoom)
    for index in range(1, count):
        yy = y + height * index / count
        canvas.create_line(x, yy, x + width, yy, dash=(4, 4), fill="#788792")
    paper_length = eng.label_size_mm(self.app.proj)[1]
    canvas.create_text(x + width + 12, y + height / (2 * count), anchor="w",
                       text=f"{paper_length:g} mm", font=UI_FONT, fill=UI_MUTED)
    if actual_count > count:
        canvas.create_text(cw / 2, ch - 9, text=f"显示前 {count} 张 · 共 {actual_count} 张",
                           font=UI_FONT, fill=UI_MUTED)


PrintPreview.__init__ = _ui_preview_init
PrintPreview._redraw_canvas = _ui_preview_draw
PrintPreview._on_copies = _ui_preview_copies
PrintPreview._update_info = _ui_preview_info


_ui_queue_base = App.open_queue
_ui_batch_base = bprint.BatchPrintWindow._build


def _ui_batch_build(self):
    _ui_batch_base(self)
    _ui_skin_dialog(self.win, "批量打印", primary=("开 始", "开始", "▶ 继 续"))
    self.btn_start.configure(text="开始打印")
    self.btn_resume.configure(text="继续")
    try:
        self.win.minsize(min(805, self.win.winfo_screenwidth() - 40), 480)
    except tk.TclError:
        pass


def _ui_queue(self):
    result = _ui_queue_base(self)
    window = self.qwin
    if not window or not window.winfo_exists():
        return result
    if getattr(window, "_ui_header", None):
        return result
    _ui_skin_dialog(window, "打印队列", primary=("▶ 继 续",))
    controls = next((frame for frame in window.winfo_children()
                     if any(isinstance(w, ttk.Button) and "继续" in w.cget("text").replace(" ", "")
                            for w in frame.winfo_children())), None)
    if controls:
        for widget in controls.winfo_children():
            if isinstance(widget, ttk.Label) and "终止会" in widget.cget("text"):
                widget.pack_forget()
            elif isinstance(widget, ttk.Button):
                if "继 续" in widget.cget("text"):
                    widget.configure(text="继续")
                elif "终止当前" in widget.cget("text"):
                    widget.configure(text="终止当前")

        def pause_after_current():
            with self.qlock:
                if self.qstate == "running":
                    self.qstate = "paused"
                    self.q_paused_by = "manual"
            self._refresh_qtree()

        _ui_button(controls, "暂停后续", pause_after_current).pack(side="left", padx=5)
    return result


App.open_queue = _ui_queue
bprint.BatchPrintWindow._build = _ui_batch_build


def _ui_selftest(self):
    self._collect(dirty=False)
    project = json.loads(json.dumps(self.proj))
    printer = self.v_printer.get().strip()
    if not printer:
        messagebox.showwarning("还没选打印机", "请先在「打印」页选择一台打印机。", parent=self.root)
        return
    if getattr(self, "_st_pending", False):
        self.status.configure(text="自检页正在提交…")
        return
    paper = project["paper"]
    width, height = float(paper["w"]), float(paper["h"])
    profile = eng.bitmap_printer_profile(printer)
    dpi = profile[1] if profile else int(paper["dpi"])
    mode = eng.resolve_mode(project)
    media = paper.get("media", "gap")
    channel = "TSPL 直发" if mode == "tspl" else "图片模式"
    info = {"printer": printer, "dpi": dpi, "mode": channel, "app": APP_NAME,
            "paper": f"{width:g} × {height:g} mm", "time": time.strftime("%Y-%m-%d %H:%M")}
    dialog = tk.Toplevel(self.root, background=UI_WHITE)
    dialog.transient(self.root)
    dialog.title("自检页")
    _ui_dialog_header(dialog, "自检页")
    body = ttk.Frame(dialog, padding=(18, 8, 18, 10))
    body.pack(fill="both", expand=True)
    form = ttk.Frame(body)
    form.pack(side="left", fill="x", expand=True, anchor="n", padx=(0, 16))
    for label, value in (("打印机", printer), ("打印分辨率", f"{dpi} dpi"),
                         ("标签尺寸", info["paper"]), ("纸张类型", MEDIA_BY_VALUE[media])):
        row = ttk.Frame(form)
        row.pack(fill="x", pady=5)
        ttk.Label(row, text=label, width=12).pack(side="left")
        ttk.Label(row, text=value, foreground=UI_GREEN, wraplength=205).pack(side="left", fill="x")
    ttk.Label(form, text="量取 25 mm 标尺，并核对四角标记。",
              style="Quiet.TLabel", wraplength=260).pack(anchor="w", pady=(12, 0))
    canvas = tk.Canvas(body, width=196, height=277, background=UI_CANVAS, highlightthickness=0)
    canvas.pack(side="right", fill="both")
    try:
        image = eng.render_selftest(width, height, min(300, dpi), info)
        canvas.bind("<Configure>", lambda event: _ui_draw_test_card(canvas, image, event.width, event.height))
    except (OSError, ValueError, OverflowError):
        canvas.create_text(94, 126, text="预览暂不可用", font=UI_FONT, fill=UI_MUTED)

    def submit():
        dialog.destroy()
        self.status.configure(text="正在提交自检页…", foreground=UI_MUTED)
        self._st_result = None
        self._st_pending = True

        def work():
            try:
                _, channel_used = eng.print_selftest(printer, mode, width, height, float(paper["gap"]), info, media)
                self._st_result = True, channel_used
            except Exception as error:
                self._st_result = False, str(error)

        threading.Thread(target=work, daemon=True).start()

    footer = ttk.Frame(dialog, padding=(18, 5, 18, 15))
    footer.pack(fill="x")
    _ui_button(footer, "打印自检页", submit, primary=True).pack(side="right")
    _ui_button(footer, "关闭", dialog.destroy).pack(side="right", padx=8)
    dialog.bind("<Escape>", lambda _: dialog.destroy())
    _ui_center(dialog, self.root)
    dialog.grab_set()


def _ui_draw_test_card(canvas, image, width, height):
    canvas.delete("all")
    _ui_render_card(canvas, image, width, height)


def _ui_tspl(self):
    self._collect(dirty=False)
    dialog = tk.Toplevel(self.root, background=UI_WHITE)
    dialog.transient(self.root)
    dialog.title("TSPL 指令")
    dialog.geometry("680x465")
    dialog.minsize(470, 325)
    profile = eng.bitmap_printer_profile(self.v_printer.get())
    paper = self.proj["paper"]
    dpi = profile[1] if profile else paper.get("dpi", 300)
    device = profile[0] if profile else self.v_printer.get() or "未选打印机"
    media = MEDIA_BY_VALUE.get(paper.get("media"), "间隙纸")
    _ui_dialog_header(dialog, "TSPL 指令", f"{device} · {dpi:g} dpi · {media}")
    body = ttk.Frame(dialog, padding=(18, 4, 18, 6))
    body.pack(fill="both", expand=True)
    code = tk.Text(body, wrap="none", undo=False, font=("Consolas", 11), borderwidth=1,
                   relief="solid", background="#f4f6f8", foreground=UI_INK, padx=10, pady=9)
    scroll = ttk.Scrollbar(body, orient="vertical", command=code.yview)
    code.configure(yscrollcommand=scroll.set)
    scroll.pack(side="right", fill="y")
    code.pack(side="left", fill="both", expand=True)
    full = ""
    try:
        full = eng.build_tspl(self.proj)
        display = []
        for line in full.splitlines():
            if line.startswith("BITMAP "):
                header, _, data = line.rpartition(",")
                display.append(f"{header}, … [{len(data) // 2} 字节]")
            else:
                display.append(line)
        code.insert("1.0", "\n".join(display))
    except Exception as error:
        code.insert("1.0", str(error))
    code.configure(state="disabled")
    footer = ttk.Frame(dialog, padding=(18, 7, 18, 13))
    footer.pack(fill="x")
    caption = ttk.Label(footer, text="位图数据已折叠。" if full else "", style="Quiet.TLabel")
    caption.pack(side="left")

    def copy():
        dialog.clipboard_clear()
        dialog.clipboard_append(full)
        caption.configure(text="已复制完整 TSPL 指令")

    _ui_button(footer, "复制指令", copy, primary=True, state="normal" if full else "disabled").pack(side="right")
    _ui_button(footer, "关闭", dialog.destroy).pack(side="right", padx=8)
    dialog.bind("<Escape>", lambda _: dialog.destroy())
    _ui_center(dialog, self.root)


_ui_icons_build_base = ipick.IconPicker._build
_ui_icons_fill_base = ipick.IconPicker._fill_cur


def _ui_icons_fill(self, keep=None):
    # Use the original chooser's selection and order as the source of truth.
    _ui_icons_fill_base(self, keep=keep)
    if hasattr(self, "_ui_icon_cards"):
        self._ui_refresh_icon_cards()


def _ui_icon_image(self, path, max_size):
    key = (path, max_size)
    if key not in self._ui_icon_pictures:
        try:
            with Image.open(eng.asset_path(path)) as raw:
                picture = raw.convert("RGBA")
            picture.thumbnail((max_size, max_size), Image.Resampling.LANCZOS)
            self._ui_icon_pictures[key] = ImageTk.PhotoImage(picture)
        except (ValueError, OSError):
            self._ui_icon_pictures[key] = None
    return self._ui_icon_pictures[key]


def _ui_refresh_icon_cards(self):
    available = ipick.builtin_pool()
    active = self._sel_cur()
    for group in (self._ui_icon_cards, self._ui_chosen_cards):
        for widget in group.winfo_children():
            widget.destroy()
    for index, path in enumerate(available):
        # builtin_pool returns (display name, relative filename) on some versions.
        icon_path = path[-1] if isinstance(path, (list, tuple)) else path
        picture = _ui_icon_image(self, icon_path, 49)
        card = tk.Button(self._ui_icon_cards, image=picture or "",
                         text="" if picture else os.path.basename(icon_path), wraplength=84,
                         font=UI_FONT, background="#eaf3ee" if icon_path in self.chosen else UI_WHITE,
                         activebackground="#e1efe8", relief="flat", bd=0, highlightthickness=1,
                         highlightbackground=UI_GREEN if icon_path in self.chosen else UI_BORDER,
                         height=56 if picture else 2,
                         command=lambda selected=index: self._ui_add_icon_card(selected))
        card.grid(row=index // 5, column=index % 5, sticky="nsew", padx=4, pady=4, ipady=3)
    for idx, path in enumerate(self.chosen):
        picture = _ui_icon_image(self, path, 34)
        card = tk.Button(self._ui_chosen_cards, image=picture or "",
                         text="" if picture else str(idx + 1), font=UI_FONT,
                         background="#eaf3ee" if active == idx else UI_WHITE,
                         activebackground="#e1efe8", relief="flat", bd=0, highlightthickness=1,
                         highlightbackground=UI_GREEN if active == idx else UI_BORDER,
                         command=lambda selected=idx: self._ui_pick_icon_card(selected))
        card.grid(row=idx // 10, column=idx % 10, padx=(0, 5), pady=3, ipadx=5, ipady=4)


def _ui_add_icon_card(self, index):
    self.lst_pool.selection_clear(0, "end")
    self.lst_pool.selection_set(index)
    self._add_sel()


def _ui_pick_icon_card(self, index):
    self.lst_cur.selection_clear(0, "end")
    self.lst_cur.selection_set(index)
    self._ui_refresh_icon_cards()


def _ui_icons_build(self):
    _ui_icons_build_base(self)
    for child in self.win.pack_slaves():
        child.pack_forget()
    _ui_dialog_header(self.win, "选择洗涤图标")
    body = ttk.Frame(self.win, padding=(18, 6, 18, 14))
    body.pack(fill="both", expand=True)
    body_bottom = ttk.Frame(body)
    body_bottom.pack(side="bottom", fill="x")
    footer = ttk.Frame(body_bottom)
    footer.pack(side="bottom", fill="x", pady=(7, 0))
    _ui_button(footer, "确定", self._ok, primary=True).pack(side="right")
    _ui_button(footer, "取消", self._cancel).pack(side="right", padx=8)
    _ui_button(footer, "导入图片…", self._add_file).pack(side="left")
    _ui_button(footer, "恢复默认", self._reset).pack(side="left", padx=6)
    _ui_divider(body_bottom, (9, 1))
    _ui_heading(body_bottom, "已选图标")
    self._ui_chosen_cards = ttk.Frame(body_bottom)
    self._ui_chosen_cards.pack(fill="x", pady=(0, 5))
    order = ttk.Frame(body_bottom)
    order.pack(fill="x")
    for text, action in (("上移", lambda: self._move(-1)), ("下移", lambda: self._move(1)),
                         ("移除", self._remove), ("清空所选", self._clear)):
        _ui_button(order, text, action).pack(side="left", padx=(0, 6))
    self.lbl_n = ttk.Label(order, style="Quiet.TLabel")
    self.lbl_n.pack(side="right")
    _ui_heading(body, "可用图标")
    viewport = ttk.Frame(body)
    viewport.pack(fill="both", expand=True)
    self._ui_icon_cards = _ui_panel(viewport)
    self._ui_icon_cards.configure(padding=0)
    for col in range(5):
        self._ui_icon_cards.columnconfigure(col, weight=1, uniform="icons")
    self._ui_icon_pictures = {}
    self._fill_cur(keep=0)
    self.win.minsize(590, 435)
    self.win.geometry("720x560")
    _ui_center(self.win, self._parent)


App.do_selftest = _ui_selftest
App.show_tspl = _ui_tspl
ipick.IconPicker._build = _ui_icons_build
ipick.IconPicker._fill_cur = _ui_icons_fill
ipick.IconPicker._ui_refresh_icon_cards = _ui_refresh_icon_cards
ipick.IconPicker._ui_add_icon_card = _ui_add_icon_card
ipick.IconPicker._ui_pick_icon_card = _ui_pick_icon_card


# Preserve both original editors and their callbacks; only lay out their existing controls.
_ui_legacy_table = App.open_table_editor
_ui_legacy_import = App._ddl_confirm
_ui_legacy_rebuild_table_rows = App._rebuild_table_rows


def _ui_fit_legacy_dialog(window, anchor, minimum_width, minimum_height):
    window.update_idletasks()
    width = min(window.winfo_screenwidth() - 30, max(window.winfo_reqwidth(), minimum_width))
    height = min(window.winfo_screenheight() - 85, max(window.winfo_reqheight(), minimum_height))
    window.geometry(f"{width}x{height}")
    _ui_center(window, anchor)


def _ui_pin_table_footer(self, window, settings):
    info = self.lbl_tableinfo
    footer = info.master
    actions = [child for child in footer.pack_slaves() if isinstance(child, ttk.Button)]
    for child in footer.pack_slaves():
        child.pack_forget()

    footer.pack_forget()
    footer.configure(padding=(12, 5, 12, 8))
    footer.pack(side="bottom", fill="x", before=settings)
    ttk.Separator(footer, orient="horizontal").grid(row=0, column=0, columnspan=6,
                                                       sticky="ew", pady=(0, 6))
    footer.columnconfigure(4, weight=1)
    for column, button in enumerate(actions[:-1]):
        button.grid(row=1, column=column, sticky="w", padx=(0, 7))
    actions[-1].grid(row=1, column=5, sticky="e")
    info.grid(row=2, column=0, columnspan=6, sticky="w", pady=(5, 0))

    window.update_idletasks()
    width, height = window.minsize()
    width = max(width, footer.winfo_reqwidth())
    info.configure(wraplength=width - 24)
    # Preserve a usable row viewport before resizing the outer window.
    rows = self._tinner.master.master
    fixed_height = window.winfo_reqheight() - rows.winfo_reqheight()
    if not self.lbl_sizehint.winfo_manager():
        fixed_height += self.lbl_sizehint.winfo_reqheight() + 8
    height = max(height, min(window.winfo_screenheight() - 90, fixed_height + 54))
    window.minsize(width, height)


def _ui_rebuild_table_rows(self):
    _ui_legacy_rebuild_table_rows(self)
    inner = getattr(self, "_tinner", None)
    if inner is None or not inner.winfo_exists():
        return
    rows = (child for child in inner.winfo_children() if isinstance(child, ttk.Frame))
    for index, row in enumerate(rows):
        number, font, height_label, height, *cells, remove = row.pack_slaves()
        widgets = (number, font, height_label, height, *cells, remove)
        for widget in widgets:
            widget.pack_forget()

        style = "AlternateRow" if index % 2 else "Row"
        row.configure(style=f"{style}.TFrame", padding=(8, 3))
        row.pack_configure(pady=0)
        number.configure(style=f"{style}.TLabel")
        height_label.configure(text="行高", style=f"{style}.TLabel", foreground=UI_MUTED)
        font.configure(style="Table.TCombobox")
        height.configure(style="Table.TEntry")
        for column, widget in enumerate(widgets):
            widget.grid(row=0, column=column, sticky="ew", padx=(0, 6))
        for column, cell in enumerate(cells, start=4):
            cell.configure(style="Table.TEntry", width=6)
            row.columnconfigure(column, weight=1, uniform="cells")
        remove.configure(style="Table.TButton")
        remove.grid_configure(padx=(0, 0))


def _ui_add_table_row(self, cols):
    table = self._tproj.setdefault("table", {})
    table.setdefault("rows", []).append({"cols": cols, "font": UI_TABLE_ROW_FONT, "cells": [""] * cols})
    self._mark_dirty()
    self._rebuild_table_rows()
    if self._tproj is self.proj:
        self.redraw()


def _ui_gen_size_table(self):
    sizes = self._size_selected()
    if not sizes:
        messagebox.showinfo("还没勾尺码", "先在左边勾几个尺码，再点生成。", parent=self._twin)
        return
    if self._table_has_content() and not messagebox.askyesno(
        "会覆盖现有表格",
        "表格里已经有内容了，生成会把它整个替换掉。\n\n继续吗？",
        parent=self._twin,
    ):
        return

    size_height = max(2.0, min(20.0, self._f(self.v_szh1, 3.0)))
    value_height = max(2.0, min(20.0, self._f(self.v_szh2, 3.5)))
    rows = eng.size_table_rows(sizes, self._parse_vals(), size_height, value_height, font=UI_TABLE_ROW_FONT)
    table = self._tproj.setdefault("table", {})
    table.update(rows=rows, size_picks=sizes, on=True)
    self._mark_dirty()
    self._rebuild_table_rows()
    if self._tproj is self.proj:
        self._refresh_gap_row(self.proj)
        self.redraw()
    self.status.config(text=f"已按 {len(sizes)} 个尺码生成 {len(rows)} 行表格")


def _ui_style_legacy_table(self):
    answer = _ui_legacy_table(self)
    window = self._twin
    if not window or not window.winfo_exists() or getattr(window, "_ui_header", None):
        return answer
    settings = window.pack_slaves()[0]
    settings.configure(padding=(16, 5, 16, 3))
    for field in settings.grid_slaves():
        if isinstance(field, (ttk.Entry, ttk.Combobox)):
            field.grid_configure(pady=1)
            field.configure(style="Table.TCombobox" if isinstance(field, ttk.Combobox) else "Table.TEntry")

    generator = self.btn_gen.master.master
    generator.configure(padding=(9, 4, 9, 5))
    generator.pack_configure(pady=(5, 0))
    for field in generator.grid_slaves():
        if isinstance(field, ttk.Label) and field.cget("text").startswith("生成的表格就是普通行"):
            field.grid_remove()
        elif isinstance(field, ttk.Frame):
            field.grid_configure(pady=(2, 0))
    for field in _ui_walk(generator):
        if isinstance(field, ttk.Entry):
            field.configure(style="Table.TEntry")
    self.btn_gen.configure(style="Table.TButton")

    _ui_skin_dialog(window, "表格工具", primary=("关闭",))
    window._ui_header.configure(padding=(18, 7, 18, 5))
    canvas = self._tinner.master
    canvas.configure(height=270)
    row_item = canvas.find_all()[0]
    canvas.itemconfigure(row_item, width=canvas.winfo_width())
    canvas.bind("<Configure>", lambda event: canvas.itemconfigure(row_item, width=event.width), add="+")
    self._ui_pin_table_footer(window, settings)
    _ui_fit_legacy_dialog(window, self.root, 1020, 585)
    fit_window_to_screen(window)
    return answer


def _ui_style_legacy_import(self, data):
    existing = set(self.root.winfo_children())
    answer = _ui_legacy_import(self, data)
    window = next((child for child in self.root.winfo_children()
                   if isinstance(child, tk.Toplevel) and child not in existing), None)
    if window is None:
        return answer
    for child in window.pack_slaves():
        if isinstance(child, ttk.Label):
            child.pack_configure(anchor="w", padx=17)
    _ui_skin_dialog(window, "导入模板", primary=("填入新标签页",))
    _ui_fit_legacy_dialog(window, self.root, 565, 390)
    return answer


App._ui_pin_table_footer = _ui_pin_table_footer
App._rebuild_table_rows = _ui_rebuild_table_rows
App._add_table_row = _ui_add_table_row
App._gen_size_table = _ui_gen_size_table
App.open_table_editor = _ui_style_legacy_table
App._ddl_confirm = _ui_style_legacy_import

if __name__ == "__main__":
    main()
