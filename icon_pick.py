# -*- coding: utf-8 -*-
"""选洗涤图标的弹窗。

为什么单独开窗口：图标不是「一个文件路径」，而是**一串有顺序的图**——
手洗 / 不可漂白 / 阴凉晾干 / 低温熨 / 干洗，顺序错了意思就变了。
所以要有「池 → 顺序列表 → 上下移」这套动作，主界面那一个小格子装不下。

图从哪里来：
  · 内置池 = assets/icons/ 下现成的 png（换成自己的一套直接覆盖同名文件即可）
  · 「添加图片…」把任意 png/jpg 复制进 assets/icons/，随程序一起走，不会因为
    源文件被删/被挪就失效
"""
import os
import shutil
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import label_engine as eng

FONT = "Microsoft YaHei"
DEFAULT_ICONS = ["assets/icons/icon%d.png" % i for i in range(1, 6)]


def icons_dir():
    """图标目录（跟着 exe/脚本走，可写）"""
    return os.path.join(eng._app_write_dir(), "assets", "icons")


def builtin_pool():
    """池子里有哪些图标：目录里现成的 png/jpg，按文件名排。
    排除 source_strip.png —— 那是原图素材，不是图标。"""
    d = icons_dir()
    out = []
    if os.path.isdir(d):
        for f in sorted(os.listdir(d)):
            low = f.lower()
            if low.endswith((".png", ".jpg", ".jpeg", ".bmp")) and "source_strip" not in low:
                out.append("assets/icons/" + f)
    return out


class IconPicker:
    """一次选择；结果放在 self.result（取消为 None），是相对路径列表

    cfg 传标签当前的槽位参数（slot / gap / row_gap / max_cols），
    预览按同一套比例画 —— 不然弹窗里看着一排，打出来是两排。"""

    def __init__(self, parent, current=None, cfg=None):
        self._parent = parent
        self.result = None
        self.chosen = [str(p) for p in (current if current is not None
                                        else DEFAULT_ICONS) if str(p).strip()]
        self.cfg = dict(cfg or {})
        self._pics = []          # 防 ImageTk 被回收
        self._build()

    # ---------------- 界面 ----------------
    def _build(self):
        w = tk.Toplevel(self._parent)
        self.win = w
        w.title("选洗涤图标")
        w.minsize(720, 520)
        w.transient(self._parent)
        try:
            px, py = self._parent.winfo_rootx(), self._parent.winfo_rooty()
            pw, ph = self._parent.winfo_width(), self._parent.winfo_height()
            # 高度给到 600：两排图标时预览会占掉 150px 左右，
            # 还按 500 开窗的话「确定 / 取消」会被顶到窗口外面去。
            w.geometry("780x600%+d%+d" % (px + max(0, (pw - 780) // 2),
                                          py + max(0, (ph - 600) // 3)))
        except tk.TclError:
            w.geometry("780x600")

        ttk.Label(w, text="左边点一下 → 加到右边；右边的顺序就是打印顺序（从左到右、从上到下）",
                  foreground="#888").pack(anchor="w", padx=12, pady=(10, 6))

        body = ttk.Frame(w, padding=(12, 0))
        body.pack(fill=tk.BOTH, expand=True)
        body.columnconfigure(0, weight=1, uniform="c")
        body.columnconfigure(1, weight=1, uniform="c")
        body.rowconfigure(1, weight=1)

        ttk.Label(body, text="图标池").grid(row=0, column=0, sticky="w")
        ttk.Label(body, text="标签上的图标").grid(row=0, column=1, sticky="w", padx=(12, 0))

        left = ttk.Frame(body)
        left.grid(row=1, column=0, sticky="nsew")
        self.lst_pool = tk.Listbox(left, activestyle="dotbox", exportselection=False,
                                   font=(FONT, 10))
        sb1 = ttk.Scrollbar(left, orient="vertical", command=self.lst_pool.yview)
        self.lst_pool.configure(yscrollcommand=sb1.set)
        self.lst_pool.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb1.pack(side=tk.LEFT, fill=tk.Y)
        self.lst_pool.bind("<Double-Button-1>", lambda e: self._add_sel())
        btns = ttk.Frame(left)
        btns.pack(side=tk.LEFT, fill=tk.Y, padx=(6, 0))
        ttk.Button(btns, text="加入 →", width=8, command=self._add_sel).pack(pady=(0, 6))
        ttk.Button(btns, text="添加图片…", width=8, command=self._add_file).pack()

        right = ttk.Frame(body)
        right.grid(row=1, column=1, sticky="nsew", padx=(12, 0))
        self.lst_cur = tk.Listbox(right, activestyle="dotbox", exportselection=False,
                                  font=(FONT, 10))
        sb2 = ttk.Scrollbar(right, orient="vertical", command=self.lst_cur.yview)
        self.lst_cur.configure(yscrollcommand=sb2.set)
        self.lst_cur.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb2.pack(side=tk.LEFT, fill=tk.Y)
        btns2 = ttk.Frame(right)
        btns2.pack(side=tk.LEFT, fill=tk.Y, padx=(6, 0))
        ttk.Button(btns2, text="上移", width=8,
                   command=lambda: self._move(-1)).pack(pady=(0, 6))
        ttk.Button(btns2, text="下移", width=8,
                   command=lambda: self._move(1)).pack(pady=(0, 6))
        ttk.Button(btns2, text="移除", width=8,
                   command=self._remove).pack(pady=(0, 6))
        ttk.Button(btns2, text="清空", width=8,
                   command=self._clear).pack()

        # 预览：按标签上的槽位网格画 —— 固定大小的格子、一排最多 5 个、
        # 多了分排并居中。所见即所印，包括「排成几排」这件事。
        pf = ttk.LabelFrame(w, text=" 预览（按当前顺序） ", padding=6)
        pf.pack(fill=tk.X, padx=12, pady=(10, 0))
        self.cv = tk.Canvas(pf, height=96, bg="#FBFBFC", highlightthickness=0)
        self.cv.pack(fill=tk.X)
        # 窗口一变宽就重画：不然图标是按初始宽度排的，拉宽后全挤在左边
        self.cv.bind("<Configure>", lambda e: self._draw_preview())

        foot = ttk.Frame(w, padding=(12, 8))
        foot.pack(fill=tk.X)
        self.lbl_n = ttk.Label(foot, text="", foreground="#888")
        self.lbl_n.pack(side=tk.LEFT)
        ttk.Button(foot, text="确定", width=9, command=self._ok).pack(side=tk.RIGHT)
        ttk.Button(foot, text="取消", width=9, command=self._cancel).pack(side=tk.RIGHT, padx=(0, 6))
        ttk.Button(foot, text="恢复默认五个", command=self._reset).pack(side=tk.RIGHT, padx=(0, 12))

        w.bind("<Escape>", lambda e: self._cancel())
        w.protocol("WM_DELETE_WINDOW", self._cancel)
        self._fill()

    # ---------------- 数据 ----------------
    def _fill(self):
        self.lst_pool.delete(0, tk.END)
        for rel in builtin_pool():
            # 只显示文件名：把 assets/icons/ 前缀也写进去的话，列表一半宽度都被
            # 重复的路径吃掉，反而看不见真正的文件名
            self.lst_pool.insert(tk.END, os.path.basename(rel))
        if self.lst_pool.size():
            self.lst_pool.selection_set(0)
        self._fill_cur()

    def _fill_cur(self, keep=None):
        self.lst_cur.delete(0, tk.END)
        for i, rel in enumerate(self.chosen):
            name = os.path.basename(rel)
            miss = "" if os.path.exists(eng.asset_path(rel)) else "   ⚠ 文件不存在"
            self.lst_cur.insert(tk.END, "%d. %s%s" % (i + 1, name, miss))
        if keep is not None and 0 <= keep < self.lst_cur.size():
            self.lst_cur.selection_clear(0, tk.END)
            self.lst_cur.selection_set(keep)
            self.lst_cur.see(keep)
        # 「排成几排」写在这里而不是画在画布里：画布多留一行字就要多 16px 高，
        # 两排图标时窗口会顶到屏幕边，那行字反被窗口边框裁掉。
        txt = "共 %d 个图标" % len(self.chosen)
        if self.chosen:
            try:
                _, _, _, mc = self._slot_cfg()
                txt += " · 排 %s" % " + ".join(str(k) for k in
                                              eng.icon_rows(len(self.chosen), mc))
            except Exception:
                pass
        self.lbl_n.config(text=txt)
        self._draw_preview()

    def _sel_cur(self):
        s = self.lst_cur.curselection()
        return s[0] if s else -1

    def _sel_pool(self):
        s = self.lst_pool.curselection()
        return s[0] if s else -1

    # ---------------- 动作 ----------------
    def _add_sel(self):
        i = self._sel_pool()
        pool = builtin_pool()
        if i < 0 or i >= len(pool):
            return
        self.chosen.append(pool[i])
        self._fill_cur(keep=len(self.chosen) - 1)

    def _add_file(self):
        paths = filedialog.askopenfilenames(
            parent=self.win, title="选图标图片",
            filetypes=[("图片", "*.png *.jpg *.jpeg *.bmp"), ("全部文件", "*.*")])
        if not paths:
            return
        d = icons_dir()
        try:
            os.makedirs(d, exist_ok=True)
        except OSError as e:
            messagebox.showerror("添加失败", "图标目录不可写：\n%s\n\n%s" % (d, e),
                                 parent=self.win)
            return
        for src in paths:
            base = os.path.basename(src)
            dst = os.path.join(d, "custom_" + base)
            try:
                # 复制而不是引用原路径：原图被挪走/删掉之后标签不该突然少个图标
                if not os.path.exists(dst):
                    shutil.copy2(src, dst)
            except OSError as e:
                messagebox.showerror("添加失败", "复制失败：\n%s\n\n%s" % (src, e),
                                     parent=self.win)
                continue
            self.chosen.append("assets/icons/custom_" + base)
        self._fill()
        self._fill_cur(keep=len(self.chosen) - 1)

    def _move(self, d):
        i = self._sel_cur()
        j = i + d
        if i < 0 or not (0 <= j < len(self.chosen)):
            return
        self.chosen[i], self.chosen[j] = self.chosen[j], self.chosen[i]
        self._fill_cur(keep=j)

    def _remove(self):
        i = self._sel_cur()
        if i < 0:
            return
        self.chosen.pop(i)
        self._fill_cur(keep=min(i, len(self.chosen) - 1))

    def _clear(self):
        self.chosen = []
        self._fill_cur()

    def _reset(self):
        self.chosen = list(DEFAULT_ICONS)
        self._fill_cur(keep=0)

    # ---------------- 预览 ----------------
    def _slot_cfg(self):
        """槽位参数（mm）；没传就用引擎默认值"""
        d = eng.DEFAULT_PROJECT["bottom"]
        c = self.cfg or {}
        def num(k, dflt):
            try:
                return float(c.get(k, dflt) or dflt)
            except Exception:
                return float(dflt)
        slot = max(1.0, num("icon_slot_mm", d.get("icon_slot_mm", 5.0)))
        gap = max(0.0, num("icon_slot_gap_mm", d.get("icon_slot_gap_mm", 0.6)))
        rgap = max(0.0, num("icon_row_gap_mm", d.get("icon_row_gap_mm", 1.0)))
        try:
            mc = max(1, int(float(c.get("icon_max_cols", d.get("icon_max_cols", 5)) or 5)))
        except Exception:
            mc = 5
        return slot, gap, rgap, mc

    def _draw_preview(self):
        c = self.cv
        c.delete("all")
        self._pics = []
        if not self.chosen:
            c.configure(height=96)
            c.create_text(12, 48, text="（没有图标 —— 这一行不会打印）",
                          anchor="w", fill="#999", font=(FONT, 9))
            return
        try:
            from PIL import Image, ImageTk
        except ImportError:
            c.create_text(12, 48, text="（预览需要 Pillow，打印不受影响）",
                          anchor="w", fill="#999", font=(FONT, 9))
            return
        imgs = []
        for rel in self.chosen:
            try:
                imgs.append(Image.open(eng.asset_path(rel)))
            except (OSError, ValueError):
                continue
        if not imgs:
            c.create_text(12, 48, text="（选中的图标文件都读不到）", anchor="w",
                          fill="#c62828", font=(FONT, 9))
            return

        slot_mm, gap_mm, rgap_mm, mc = self._slot_cfg()
        plan = eng.icon_rows(len(imgs), mc)
        nrow = max(1, len(plan))

        # 可用宽度要问 Canvas 实际有多宽。刚创建时 winfo_width() 返回 1，
        # 拿它算就会按 200px 排，等窗口真铺开、图标全挤在左边 —— 所以绑了 <Configure>。
        cw = c.winfo_width()
        if cw <= 1:
            cw = 720
        pad = 12
        # 槽位在屏幕上的边长：先按「最宽的一排要塞进画布」定上限，
        # 再按排数压（三排以上不能把预览撑到吃掉按钮区），最后压一个 56px 的顶
        widest = max(plan) if plan else 1
        by_w = (cw - pad * 2 - (widest - 1) * 4) / float(widest)
        by_h = 150.0 / nrow
        slot_px = max(10, int(min(56, by_w, by_h)))
        gap_px = int(round(slot_px * gap_mm / slot_mm))
        rgap_px = int(round(slot_px * rgap_mm / slot_mm))
        total_h = nrow * slot_px + (nrow - 1) * rgap_px
        c.configure(height=max(76, total_h + pad * 2))

        i = 0
        for r, cnt in enumerate(plan):
            seg = imgs[i:i + cnt]
            i += cnt
            row_w = len(seg) * slot_px + max(0, len(seg) - 1) * gap_px
            x = max(pad, (cw - row_w) // 2)          # 居中，跟打印时一样
            y = pad + r * (slot_px + rgap_px)
            for im in seg:
                # 槽位框：虚线画出来，让人看见「格子是固定的，图只是塞进去」
                c.create_rectangle(x, y, x + slot_px, y + slot_px,
                                   outline="#DCDFE6", dash=(2, 2))
                k = min(slot_px / float(im.width), slot_px / float(im.height), 1.0)
                nw = max(1, int(round(im.width * k)))
                nh = max(1, int(round(im.height * k)))
                th = im.resize((nw, nh), Image.LANCZOS)
                ph = ImageTk.PhotoImage(th)
                self._pics.append(ph)
                c.create_image(x + (slot_px - nw) // 2, y + (slot_px - nh) // 2,
                               image=ph, anchor="nw")
                x += slot_px + gap_px

        rows_txt = " + ".join(str(k) for k in plan)
        self._rows_txt = rows_txt         # 给底下那行「共 N 个 · 3 + 4」用

    # ---------------- 结果 ----------------
    def _ok(self):
        self.result = list(self.chosen)
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


def choose(parent, current=None, cfg=None):
    """打开选图标窗口；返回相对路径列表，取消返回 None"""
    dlg = IconPicker(parent, current=current, cfg=cfg)
    try:
        dlg.win.grab_set()
        dlg.win.wait_window()
    except tk.TclError:
        pass
    return dlg.result
