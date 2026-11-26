# -*- coding: utf-8 -*-
"""
========================================================================
标签引擎 label_engine.py
========================================================================
职责：把「一个项目的数据」翻译成 TSPL 指令，并发给打印机。
      不涉及任何界面，界面层调用这里。

一个项目（.wslabel 文件）就是一份 JSON：
{
  "name": "水洗标",
  "paper":  { "w": 35, "h": 75, "gap": 2, "dpi": 203, "media": "gap" },
  "printer": "Gprinter GP-1834T",
  "left_mm": 3.0,          # 左边距
  "start_y_mm": 22.0,      # 文字块起始高度位置
  "font_dots": 24,         # 字库点阵：16 / 24 / 32
  "line_gap": 4,           # 行间距（dot）
  "title": "号型: 175/92A（XL）",
  "fabric_head": "成分:",
  "fabrics": ["65%棉", "35%聚酯纤维"]
}
========================================================================
"""
import ctypes
import atexit
import re
import json
import os
import sys
import time
from ctypes import wintypes

from PIL import Image

# ---------------- 进程内缓存 ----------------
# 为什么要缓存：主预览每敲一个键就重算一遍版面，而 bottom_ops 为了算图标高度
# 会把 PNG 从磁盘读出来解码；一次重绘里 bottom_ops 被调用 4 次、图标解码 25 次。
# 实测重绘 26.6ms → 7.1ms，渲染结果逐像素不变。
# 所有图像缓存的键都带 (mtime, size)，替换同名图片后立刻失效，不会读到旧图。
_ICON_CACHE = {}          # (文件身份, 宽dot, 阈值) → (PIL 图, 宽, 高)
_IMG_SIZE_CACHE = {}      # 文件身份 → (宽, 高)
_BOTTOM_OPS_CACHE = {}    # bottom/paper/边距 的签名 → (ops, 高度dot)


# ---------------- Win32 RAW 打印（纯 ctypes，零第三方依赖） ----------------
class _DOC_INFO_1(ctypes.Structure):
    _fields_ = [("pDocName", wintypes.LPWSTR),
                ("pOutputFile", wintypes.LPWSTR),
                ("pDatatype", wintypes.LPWSTR)]


# ---------------- 打印机使用次数（决定列表排序）----------------
# 为什么记次数：一台机器上常挂着一堆打印机（虚拟的 PDF/Fax、共享机、标签机），
# 平铺按字母排的话，天天用的那台常被挤到看不见的地方。按用过的次数排，常用的自然浮上来。
# 为什么落在程序目录：这个工具是绿色便携的，配置跟着程序走，拷到另一台机器还能用。
PRINTER_USE_FILE = "printer_usage.json"


def _app_write_dir():
    """可写目录：打包后是 exe 所在目录，开发时是源码目录"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def _use_path():
    return os.path.join(_app_write_dir(), PRINTER_USE_FILE)


USE_FILE_VERSION = 2


def _machine_tag():
    """这台电脑的标识

    配置落在程序目录，整个文件夹拷到另一台电脑时记忆会跟着走 —— 可「上次用条码打印」
    说的是上一台机器的事，那台打印机在这台机器上多半根本不存在。
    取不到机器名时返回空串：空串不参与隔离，免得莫名其妙把记忆清掉。
    """
    try:
        return (os.environ.get("COMPUTERNAME") or "").strip().upper()
    except Exception:
        return ""


def _as_bucket(raw):
    """把一桶记忆规整成 {"default", "counts"}；脏值直接丢，不让它把排序带歪"""
    default, counts = "", {}
    if isinstance(raw, dict):
        default = str(raw.get("default") or "").strip()
        src = raw.get("counts")
        if isinstance(src, dict):
            for k, v in src.items():
                if not str(k).strip():
                    continue
                try:
                    counts[str(k)] = int(v)
                except (TypeError, ValueError):
                    continue
    return {"default": default, "counts": counts}


def _read_use_file():
    """整份配置：{"version": 2, "machines": {电脑名: {"default","counts"}}}

    兼容两种老格式（都当成本机的，老用户升级不该丢记录）：
      · {"default": 名, "counts": {...}}   —— 加机器标识之前的那版
      · {名字: 次数}                        —— 最早只有计数表的那版
    """
    fresh = {"version": USE_FILE_VERSION, "machines": {}}
    try:
        with open(_use_path(), encoding="utf-8") as f:
            d = json.load(f)
    except Exception:
        return fresh
    if not isinstance(d, dict):
        return fresh
    if isinstance(d.get("machines"), dict):              # 新格式：按电脑分桶
        return {"version": USE_FILE_VERSION, "machines": dict(d["machines"])}
    if "counts" in d or "default" in d:                  # 老格式一
        return {"version": USE_FILE_VERSION, "machines": {"": _as_bucket(d)}}
    return {"version": USE_FILE_VERSION,                   # 老格式二：整份就是计数表
            "machines": {"": _as_bucket({"counts": d})}}


def _read_use_doc_file():
    """本机的记忆：{"machine": 电脑名, "default": 上次切换的那台, "counts": {名字: 次数}}

    只取本机那一桶 —— 别的电脑的记录留在文件里，但不采用。
    老格式没有电脑名，任何一台机器都可以接手（升级不丢记录）。
    """
    here = _machine_tag()
    machines = _read_use_file()["machines"]
    if here and here in machines:
        bucket = _as_bucket(machines[here])
    elif "" in machines:
        bucket = _as_bucket(machines[""])
    else:
        bucket = {"default": "", "counts": {}}
    return {"machine": here, "default": bucket["default"], "counts": bucket["counts"]}


def _write_use_doc_file(doc):
    """只写本机那一桶，别的电脑的记录原样留着 —— 这个文件夹是拷来拷去的"""
    here = doc.get("machine") or _machine_tag()
    machines = dict(_read_use_file()["machines"])
    machines.pop("", None)                # 老格式那桶已并入本机名下，别留两份
    machines[here] = {"default": doc.get("default", ""),
                      "counts": dict(doc.get("counts") or {})}
    path = _use_path()
    try:
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump({"version": USE_FILE_VERSION, "machines": machines},
                      f, ensure_ascii=False, indent=2)
        os.replace(tmp, path)            # 先写临时文件再换，避免写一半断电留下坏 json
    except Exception:
        pass


# 进程内缓存。原本每次「记一次使用」都要读一遍文件再写一遍文件，实测 5.7ms/次 ——
# 批量打 200 个文件光记数就花 1.1 秒，1000 个要 5.7 秒，全是磁盘往返。
# 现在读一次放内存，计数攒够 1.5 秒才落盘；「切换打印机」这种真该立刻记住的，
# 走 force 马上写。程序退出时 atexit 再兜一次。
_USE_MEM = {"path": None, "doc": None, "dirty": False, "stamp": 0.0, "sig": None}


def _use_sig(path):
    """文件的「指纹」：mtime + 大小。用来发现文件被外面改过"""
    try:
        st = os.stat(path)
        return (st.st_mtime_ns, st.st_size)
    except OSError:
        return None


def _use_mem():
    path = _use_path()
    sig = _use_sig(path)
    if (_USE_MEM["doc"] is None or _USE_MEM["path"] != path
            or _USE_MEM["sig"] != sig):
        # 文件变过（换电脑拷来的、手工改的、被别的实例写过的）→ 重新读，以文件为准。
        # 这里**不**把自己没落盘的计数写回去：那样等于拿内存缓存盖掉别人的修改，
        # 连「文件被改坏了要看空表」都做不到。丢几个计数只影响排序，无关紧要。
        _USE_MEM["doc"] = _read_use_doc_file()
        _USE_MEM["path"] = path
        _USE_MEM["sig"] = sig
        _USE_MEM["dirty"] = False
        _USE_MEM["stamp"] = 0.0
    return _USE_MEM["doc"]


def flush_printer_use(force=True):
    """把内存里的使用记录落盘。force=False 时 1.5 秒内的写请求合并成一次"""
    if not _USE_MEM["dirty"] or _USE_MEM["doc"] is None:
        return
    if not force and time.time() - _USE_MEM["stamp"] < 1.5:
        return
    _write_use_doc_file(_USE_MEM["doc"])
    _USE_MEM["dirty"] = False
    _USE_MEM["stamp"] = time.time()
    _USE_MEM["sig"] = _use_sig(_USE_MEM["path"])   # 自己刚写的，别当外部改动


def _read_use_doc():
    return _use_mem()


def _write_use_doc(doc):
    _USE_MEM["doc"] = doc
    _USE_MEM["path"] = _use_path()
    _USE_MEM["dirty"] = True
    flush_printer_use(force=True)


def forget_printer_use_memory():
    """测试/重置用：丢掉内存缓存，下次读会重新读文件"""
    _USE_MEM.update(path=None, doc=None, dirty=False, stamp=0.0, sig=None)


atexit.register(lambda: flush_printer_use(force=True))   # 关程序前把没落盘的计数写掉



def use_doc_foreign():
    """记忆是不是别的电脑留下的？（是的话界面得说明一句，别让人以为软件没记住）"""
    here = _machine_tag()
    machines = _read_use_file()["machines"]
    if not here or "" in machines:        # 老格式 / 取不到机器名 → 当本机
        return False
    if here in machines:
        b = _as_bucket(machines[here])
        if b["default"] or b["counts"]:
            return False                  # 本机已经有记录了，不算外来
    for k, v in machines.items():
        if k == here:
            continue
        b = _as_bucket(v)
        if b["default"] or b["counts"]:
            return True
    return False


def printer_usage():
    """{打印机名: 用过几次}；文件坏了就当空表，不让它影响启动"""
    return dict(_read_use_doc()["counts"])


def app_default_printer():
    """上次「亲手切换」到的那台（软件级记忆，下次打开直接用它）"""
    return _read_use_doc()["default"]


def set_app_default_printer(name):
    """记住「以后就用这台」；只在用户主动切换时调，不动使用次数"""
    name = (name or "").strip()
    doc = _read_use_doc()
    doc["default"] = name
    _write_use_doc(doc)
    return name


def note_printer_use(name, step=1):
    """记一次「用这台打印机打东西」，返回更新后的计数表

    计数挂在真正入队打印那一刻，而不是在列表里点一下 —— 点开看一眼不算使用，
    否则每次启动自动选中的那台会被刷成最高分。

    只动计数，**不动 default**：默认打印机是「你亲手切换」的结果，不是「最后打过哪台」。
    否则批量窗口用「跟文件一样」跑一轮，主界面的默认就被悄悄改掉，而你压根没切换过。
    """
    name = (name or "").strip()
    if not name:
        return printer_usage()
    doc = _use_mem()
    counts = dict(doc["counts"])
    counts[name] = counts.get(name, 0) + step
    doc["counts"] = counts
    _USE_MEM["dirty"] = True
    flush_printer_use(force=False)        # 攒 1.5 秒一次，别每张纸都写盘
    return counts


def rank_printers(infos):
    """按使用次数降序；次数相同的，默认打印机优先，再按名字 —— 稳定且可预期"""
    use = printer_usage()
    return sorted(infos, key=lambda d: (-use.get(d["name"], 0),
                                        not d["is_default"],
                                        d["name"].lower()))


def list_printers():
    """枚举本机打印机名（读注册表，比 EnumPrinters 可靠），按用过的次数排"""
    import winreg
    names = []
    try:
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                             r"SYSTEM\CurrentControlSet\Control\Print\Printers")
    except OSError:
        return names
    try:
        i = 0
        while True:
            try:
                names.append(winreg.EnumKey(key, i))
            except OSError:
                break
            i += 1
    finally:
        winreg.CloseKey(key)
    # 注册表键名有时带 ",0" 这类后缀，真实打印机名没有
    out = []
    for n in names:
        clean = re.sub(r",\d+$", "", n)
        if clean not in out:
            out.append(clean)
    use = printer_usage()
    return sorted(out, key=lambda n: (-use.get(n, 0), n.lower()))


def list_printer_info():
    """打印机列表：名称 / 驱动 / 端口 / 是否默认。给「打印列表」面板用"""
    import winreg
    default = ""
    try:
        k = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                           r"Software\Microsoft\Windows NT\CurrentVersion\Windows")
        default = winreg.QueryValueEx(k, "Device")[0].split(",")[0]
        winreg.CloseKey(k)
    except OSError:
        pass

    out = []
    try:
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                             r"SYSTEM\CurrentControlSet\Control\Print\Printers")
    except OSError:
        return out
    try:
        i = 0
        while True:
            try:
                raw = winreg.EnumKey(key, i)
            except OSError:
                break
            i += 1
            name = re.sub(r",\d+$", "", raw)
            driver = port = ""
            try:
                sk = winreg.OpenKey(key, raw)
                try:
                    driver = winreg.QueryValueEx(sk, "Printer Driver")[0]
                except OSError:
                    pass
                try:
                    port = winreg.QueryValueEx(sk, "Port")[0]
                except OSError:
                    pass
                winreg.CloseKey(sk)
            except OSError:
                pass
            out.append({"name": name, "driver": driver, "port": port,
                        "is_default": name == default,
                        "tspl": looks_like_tspl(name + " " + driver)})
    finally:
        winreg.CloseKey(key)
    return rank_printers(out)


# ---------------- 打印机的纸型 / 可印区探测 ----------------
# 能读到什么：驱动当前默认纸型、可印区、物理尺寸（都是驱动自报的数字）
# 读不到什么：机器里此刻装的是哪卷纸 —— Windows 没有这个接口，标签机基本也没有纸张传感器。
#            所以这里只能做「一致性校验」：驱动报的纸型 vs 文件尺寸，对不上就提示，
#            真正把纸型设对，还得用户点「设置」进驱动首选项改一次。
_PAPER_CACHE = {}


def printer_paper_info(name, refresh=False):
    """读目标打印机的驱动纸型 / 可印区，单位 mm。读不到返回 None（脱机、虚拟机等）"""
    if not name:
        return None
    if not refresh and name in _PAPER_CACHE:
        return _PAPER_CACHE[name]

    gdi = ctypes.WinDLL("gdi32", use_last_error=True)
    gdi.CreateDCW.restype = wintypes.HDC
    gdi.CreateDCW.argtypes = [wintypes.LPCWSTR] * 3 + [ctypes.c_void_p]
    gdi.GetDeviceCaps.argtypes = [wintypes.HDC, ctypes.c_int]
    gdi.GetDeviceCaps.restype = ctypes.c_int
    gdi.DeleteDC.argtypes = [wintypes.HDC]

    info = None
    hdc = gdi.CreateDCW("WINSPOOL", name, None, None)
    if hdc:
        try:
            dx = gdi.GetDeviceCaps(hdc, 88) or 203        # LOGPIXELSX
            dy = gdi.GetDeviceCaps(hdc, 90) or dx         # LOGPIXELSY
            pw = gdi.GetDeviceCaps(hdc, 110)              # PHYSICALWIDTH
            ph = gdi.GetDeviceCaps(hdc, 111)              # PHYSICALHEIGHT
            rw = gdi.GetDeviceCaps(hdc, 8)                # HORZRES
            rh = gdi.GetDeviceCaps(hdc, 10)               # VERTRES
            px = gdi.GetDeviceCaps(hdc, 112)              # PHYSICALOFFSETX
            py = gdi.GetDeviceCaps(hdc, 113)              # PHYSICALOFFSETY
            if pw > 0 and ph > 0:
                info = {
                    "name": name, "dpi": dx, "dpi_y": dy,
                    "phys_w": round(pw / dx * 25.4, 2),
                    "phys_h": round(ph / dy * 25.4, 2),
                    "print_w": round(rw / dx * 25.4, 2),
                    "print_h": round(rh / dy * 25.4, 2),
                    "margin_x": round(px / dx * 25.4, 2),
                    "margin_y": round(py / dy * 25.4, 2),
                    "ok": True,
                }
        finally:
            gdi.DeleteDC(hdc)
    _PAPER_CACHE[name] = info
    return info


def clear_paper_cache(name=None):
    """纸型改了之后要重新探测，先把缓存清掉"""
    if name is None:
        _PAPER_CACHE.clear()
    else:
        _PAPER_CACHE.pop(name, None)


def paper_fit(w_mm, h_mm, info, tol=0.2):
    """文件尺寸 vs 可印区。返回 (level, over_x, over_y)
       level: ok 适配 / warn 驱动纸型远大于标签（八成纸型没设对） / over 放不下 / unknown 读不到
    """
    if not info or not info.get("ok"):
        return "unknown", 0.0, 0.0
    pw, ph = info["print_w"], info["print_h"]
    over_x = max(0.0, w_mm - pw)
    over_y = max(0.0, h_mm - ph)
    if over_x > tol or over_y > tol:
        return "over", round(over_x, 2), round(over_y, 2)
    # 可印区比标签大出一倍还多，说明驱动还挂在 A4 之类的默认纸上
    if pw > max(w_mm * 2.0, w_mm + 20.0) and ph > max(h_mm * 2.0, h_mm + 20.0):
        return "warn", 0.0, 0.0
    return "ok", 0.0, 0.0


def open_printer_setup(name):
    """弹出打印机的「打印首选项」（改纸型就在这里改）。返回子进程，可 poll 等它关闭"""
    import subprocess
    try:
        return subprocess.Popen(["rundll32", "printui.dll,PrintUIEntry", "/e", "/n", name])
    except Exception:
        return None


def label_size_mm(proj):
    """成品尺寸（mm）：90°/270° 时宽高互换，预览和校纸都用它"""
    proj = normalize(proj)
    w = float(proj["paper"].get("w", 35))
    h = float(proj["paper"].get("h", 75))
    if int(proj.get("direction", 0) or 0) % 360 in (90, 270):
        return h, w
    return w, h


def looks_like_tspl(name):
    """按名字/驱动名猜是不是 TSPL 热敏机 —— 仅用于列表里的提示"""
    n = (name or "").lower()
    keys = ["gprinter", "gp-", "gp ", "tsc", "tspl", "佳博",
            "hprt", "汉印", "得力", "deli", "xprinter", "芯烨",
            "godex", "postek", "博思得", "argox", "zebra"]
    return any(k in n for k in keys)


def is_proven_tspl(name):
    """
    「自动」通道只信任实测通过的佳博 GP 系列。
    其它热敏机（得力/汉印…）很多没有中文 TSSxx.BF2 字库，直发会打黑方块，
    所以自动模式下一律走图片模式；要用直发可以手动选「TSPL 直发」。
    """
    n = (name or "").lower()
    return ("gprinter" in n) or ("gp-" in n) or n.startswith("gp ")


def resolve_mode(proj):
    """决定走哪条打印通道"""
    ang = int(proj.get("direction", 0) or 0) % 360
    if ang in (90, 270):
        return "image"          # TSPL 指令集只有 0°/180°，横打一律走图片模式
    mode = str(proj.get("print_mode") or "auto").lower()
    if mode in ("tspl", "image"):
        return mode
    name = proj.get("printer", "")
    if is_proven_tspl(name):
        return "tspl"
    # 队列名可能是「条码打印」这种别名，驱动名才是真实机型，补查一次
    try:
        for info in list_printer_info():
            if info["name"] == name and is_proven_tspl(info.get("driver", "")):
                return "tspl"
    except Exception:
        pass
    return "image"


def raw_print(printer_name, payload: bytes, doc_name="WaterLabel"):
    """把原始字节直投打印机队列（Datatype=RAW）"""
    spool = ctypes.WinDLL("winspool.drv", use_last_error=True)
    spool.OpenPrinterW.argtypes = [wintypes.LPWSTR, ctypes.POINTER(wintypes.HANDLE), ctypes.c_void_p]
    spool.OpenPrinterW.restype = wintypes.BOOL
    spool.StartDocPrinterW.argtypes = [wintypes.HANDLE, wintypes.DWORD, ctypes.c_void_p]
    spool.StartDocPrinterW.restype = wintypes.DWORD
    spool.StartPagePrinter.argtypes = [wintypes.HANDLE]
    spool.StartPagePrinter.restype = wintypes.BOOL
    spool.WritePrinter.argtypes = [wintypes.HANDLE, ctypes.c_void_p, wintypes.DWORD,
                                   ctypes.POINTER(wintypes.DWORD)]
    spool.WritePrinter.restype = wintypes.BOOL
    spool.EndPagePrinter.argtypes = [wintypes.HANDLE]
    spool.EndDocPrinter.argtypes = [wintypes.HANDLE]
    spool.ClosePrinter.argtypes = [wintypes.HANDLE]

    h = wintypes.HANDLE()
    if not spool.OpenPrinterW(printer_name, ctypes.byref(h), None):
        raise RuntimeError("打不开打印机「%s」（错误码 %d）\n请检查名字是否与系统里完全一致。"
                           % (printer_name, ctypes.get_last_error()))
    try:
        doc = _DOC_INFO_1(doc_name, None, "RAW")
        if not spool.StartDocPrinterW(h, 1, ctypes.byref(doc)):
            raise RuntimeError("StartDocPrinter 失败（错误码 %d）" % ctypes.get_last_error())
        spool.StartPagePrinter(h)
        written = wintypes.DWORD(0)
        buf = ctypes.create_string_buffer(payload)
        ok = spool.WritePrinter(h, buf, len(payload), ctypes.byref(written))
        spool.EndPagePrinter(h)
        spool.EndDocPrinter(h)
        if not ok:
            raise RuntimeError("WritePrinter 失败（错误码 %d）" % ctypes.get_last_error())
        return written.value
    finally:
        spool.ClosePrinter(h)


# ---------------- 项目数据结构 ----------------
DEFAULT_PROJECT = {
    "name": "水洗标",
    # media：纸的形态。
    #   gap        = 间隙纸（每张之间有缝/黑标，机器靠缝定位）——老文件默认这个
    #   continuous = 连续纸（整卷连着，没有缝）——长度完全由本次 SIZE/纸长决定
    "paper": {"w": 35.0, "h": 75.0, "gap": 2.0, "dpi": 203, "media": "gap"},
    "printer": "Gprinter GP-1834T",
    # 打印通道：auto = 按打印机名自动判断；tspl = 直发指令；image = 图片模式（任何打印机）
    "print_mode": "auto",
    # 打印位移（mm，可负）：模拟「标签在纸上的实际落点」，图片模式烘进像素、TSPL 走 REFERENCE
    "off_x_mm": 0.0,
    "off_y_mm": 0.0,
    # 打印方向：0 / 90 / 180。90° 只能走图片模式（TSPL 指令集原生只有 0° 和 180°）
    "direction": 0,
    "left_mm": 2.0,          # 左侧留白 2mm（固定规则）
    "right_mm": 2.0,         # 右侧留白 2mm（换行和底部区块都按它收窄）
    # 正文缩进：品牌/号型/成分这一块整体右移多少 mm。
    # 为什么要它：底版里号型块从 5.74mm 起（比左边距靠右 3.7mm），
    # 而「洗涤方式」是按整张纸居中 —— 一个 left_mm 管不了两件事。
    "body_indent_mm": 0.0,
    # 分隔线：正文（含表格）底 + gap 处画一条实心横线，底版里号型块和洗涤方式之间就有一条
    "divider": {"on": False, "gap_mm": 1.25, "w_mm": 20.0, "h_mm": 0.2},
    "start_y_mm": 20.0,      # 顶部留白 2cm（固定规则）
    "font_dots": 24,
    "line_gap": 4,
    "title": "号型: 175/92A（XL）",
    "fabric_head": "成分:",
    "fabrics": ["65%棉", "35%聚酯纤维"],
    # 产地：成分下面单独一行（「产地：中国」）。留空就不占位。
    # 为什么单列一个字段：塞进成分框里也能打出来，但那一行会跟着成分一起换行缩进，
    # 而且以后想做「产地加粗/单独间距」就没有落脚点。
    "origin": "",
    # 正文对齐：left = 全部顶左边距（老行为）；center = 每行按自己的宽度居中
    "text_align": "left",
    # 生产日期：英文月简称 + 日（如 NOV 18），单独一行居中，压在正文下面。
    # 为什么不让人直接打进「成分/产地」框里：
    #   ① 它是落款性质的一行，要独立居中、字号可调，塞进正文行里做不到；
    #   ② 单独成字段才能做月份下拉和日期校验。
    "date_on": True,
    # 月份 1~12、日 1~31。0 = 没选过，取当天（新建文件时就是今天）。
    # 存的是**值**不是规则：选好的月份日期就是纸上印的，补打同批货不会漂移。
    "date_month": 0,
    "date_day": 0,
    "date_font": 12,
    # 手动自定义内容：留空 = 照旧印「NOV 24」；填了 = 这一行就印你填的字。
    # 和 date_month/day 并存，互不覆盖 —— 想回到自动日期，清空这个框即可。
    "date_custom": "",
    # 这一行的水平对齐：left / center / right，默认居中。
    "date_align": "center",
    # 【已停用】v1.4 的「锁定日期字符串」。留着只为读回那半天存下的老文件。
    "date_lock": "",
    # 字体：空 = 自动（雅黑 → 宋体 → 黑体里挑第一个存在的）。
    # 系统字体目录下的路径会存成 $WIN\Fonts\xxx.ttf，换电脑还能找到。
    "font_file": "",
    # 品牌行：号型上面的固定内容，可开关是否参与打印
    "brand": {"on": True, "text": "品牌：AOGZ"},
    # 顶部装饰线：居中 5mm × 0.4mm，实色（BAR 命令，不经过任何缩放）
    "top_rule": {"on": True, "w_mm": 5.0, "h_mm": 0.4, "y_mm": 2.0},
    # 底部固定区块：文字走 TSPL 原生 TEXT，图标走独立 BITMAP
    "bottom": {
        "on": True,
        "title": "洗 涤 方 式",
        "subtitle": "CARE INSTRUCTION",
        "title_font": 16,
        "sub_font": 16,
        "body_font": 16,
        "caution": [
            "注意事项:按衣物颜色深浅",
            "分开洗涤，防止染色。",
            "按衣物材质用不同的洗涤剂",
            "和清洗方式。",
        ],
        "icons": [
            "assets/icons/icon1.png",
            "assets/icons/icon2.png",
            "assets/icons/icon3.png",
            "assets/icons/icon4.png",
            "assets/icons/icon5.png",
        ],
        "icon_w_dot": 0,         # 【已停用】旧逻辑：单个图标宽度上限（dot）
        "icon_gap_ratio": 0.064, # 【已停用】旧逻辑：间隙 / 图标总宽
        "icon_th": 170,          # 二值化阈值：越大笔画越粗（128 细 / 170 中 / 205 粗）
        "icon_h_max_mm": 0,      # 【已停用】旧逻辑：图标行高度上限
        # ---- 图标槽位（现役）----
        # 为什么不再「一排铺满纸宽」：那样图标大小取决于图片本身的大小和数量，
        # 换一张图、或者少放一个，剩下的全跟着放大 —— 版式每改一次就跳一次。
        # 改成固定槽位后，槽的大小只由这里决定，图片只在槽里等比缩放、位置永远不动。
        "icon_slot_mm": 5.0,     # 一个槽位的边长（正方形槽）；图片不放大，超过就等比缩进来
        "icon_slot_gap_mm": 0.6, # 同一排里槽与槽之间的横向间隙
        "icon_row_gap_mm": 1.0,  # 两排之间的行距
        "icon_max_cols": 5,      # 一排最多几个槽；超了就分排（6→3+3，7→3+4，8→4+4）
        "gap_mm": 1.0,          # 标题↔副标题、副标题↔图标 的间距
        "gap2_mm": 1.2,         # 图标↔注意事项 的间距
        "line_gap_mm": 0.3,     # 注意事项行距
        "gap_above_mm": 5.0,    # 上方留白：内容底 → 「洗涤方式」顶（唯一的位置来源）
        "margin_mm": 5.0,       # 纸底保留线：区块最多长到这里，越线就提示「挤到纸底」
        # 页脚一行（如 MADE IN CHINA）：画在注意事项下面，字号独立。
        # 为什么不让人塞进「注意事项」——那一行会跟注意事项同字号同行距，
        # 想让它大一点、离得远一点就没办法了。
        "footer": "",
        "footer_font": 16,
        # 区块位置：
        #   flow   = 紧跟正文（老行为，由「上方留白」决定）
        #   center = 在「内容底」和「纸底可用」之间垂直居中
        #   bottom = 贴纸底（底边对齐纸底保留线）
        # 为什么不干脆把「上方留白」改成贴底：一个区块只有一个位置。
        # 贴底之后上方留白就成了结果而不是设定，再留着输入框就是骗人。
        "align": "flow",
    },
    # 表格区：正文（成分）下方、洗涤说明上方的空白区
    # 两种固定行型 —— 一排三 / 一排四；行数不限；单元格文本可改、字号可调、默认居中
    "table": {
        "on": False,
        "caption": "充绒克数(g)",  # 表格上方一排小字（居中）；本机主要做充绒表，默认就给这个
        "caption_font": 16,       # 小字字号（dot，可任意 8~72）
        "caption_gap_mm": 0.5,    # 小字与表格之间的间距
        "gap_top_mm": 2.0,      # 表格与上方文字之间的留白（绿区上留白）
        "to_bottom_mm": 4.0,    # 表格底边到「洗涤方式」区块顶边的间距（可手动输入，真生效）
        "gap_bottom_mm": 2.0,   # 旧字段，仅旧文件兼容用，已不参与排版
        "row_h_mm": 4.0,        # 行高（35×75 纸上绿区约 15.3mm，4.0 可放 2 行）
        "border_dot": 2,        # 框线粗细（dot）
        "rows": [
            {"cols": 4, "font": 16, "cells": ["", "", "", ""]},
        ],
    },
}


def asset_path(rel):
    """
    把项目里的相对路径解析成绝对路径。
    优先用「exe/脚本所在目录」下的资源 —— 这样你直接替换 assets 里的图片就能生效；
    找不到再回退到打包时内嵌的资源。
    """
    if os.path.isabs(rel):
        return rel
    if getattr(sys, "frozen", False):
        base = os.path.dirname(os.path.abspath(sys.executable))
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    p = os.path.join(base, rel)
    if os.path.exists(p):
        return p
    return os.path.join(getattr(sys, "_MEIPASS", base), rel)


def _file_key(path):
    """文件身份 = 绝对路径 + 修改时间 + 大小。
    带 mtime 是为了「直接替换同名图片就换图案」这个用法能继续生效：文件一动，缓存自动失效。"""
    try:
        st = os.stat(path)
        return (os.path.abspath(path), st.st_mtime_ns, st.st_size)
    except OSError:
        return None


def _cache_put(store, key, value, limit):
    if len(store) >= limit:
        store.clear()                     # 简单粗暴的整表清空，比 LRU 更省代码且够用
    store[key] = value


def img_size(path):
    """裁掉白边后的 (宽, 高) 像素；失败返回 (0, 0)"""
    key = _file_key(path)
    if key is not None:
        hit = _IMG_SIZE_CACHE.get(key)
        if hit is not None:
            return hit
    try:
        from PIL import Image, ImageChops
        img = Image.open(path).convert("L")
        bg = Image.new("L", img.size, 255)
        bbox = ImageChops.difference(img, bg).getbbox()
        if bbox:
            r = (bbox[2] - bbox[0], bbox[3] - bbox[1])
        else:
            r = img.size
    except Exception:
        return 0, 0
    if key is not None:
        _cache_put(_IMG_SIZE_CACHE, key, r, 128)
    return r


def prep_icon(path, width_dots, thresh=170):
    """
    图标预处理：裁白边 → 灰度降采样 → 按阈值二值化。
    返回 (PIL 1-bit 图, 宽dot, 高dot)；失败返回 (None, 0, 0)。

    三个坑都踩过了：
    ① 重采样必须在灰度（L）模式下做。在 1-bit 上 resize，PIL 会顺手抖动，
       细线被打散成网点，打出来就是一片虚点。
    ② convert("1") 默认开 Floyd-Steinberg 抖动，必须显式传 dither=NONE。
    ③ 细线在热敏纸上会断。但**不要用膨胀（MinFilter）加粗** —— 线条一粗，
       洗衣盆里的手指、三角内的小三角就糊成一团。正确做法是抬高二值化阈值，
       把抗锯齿的过渡像素也算进黑里，笔画变成约 2 dot，细节还留着。

    带缓存：主预览每敲一个键都会重算版面，同一张图会被解码十几次。
    缓存键含文件修改时间，替换同名图片后立刻生效。
    """
    fk = _file_key(path)
    ck = (fk, int(width_dots), int(thresh)) if fk else None
    if ck is not None:
        hit = _ICON_CACHE.get(ck)
        if hit is not None:
            return hit
    try:
        from PIL import Image, ImageChops
        img = Image.open(path).convert("L")
        bg = Image.new("L", img.size, 255)
        bbox = ImageChops.difference(img, bg).getbbox()
        if bbox:
            img = img.crop(bbox)
        w0, h0 = img.size
        if w0 <= 0 or h0 <= 0:
            return None, 0, 0
        wd = max(1, int(width_dots))
        hd = max(1, int(round(wd * h0 / float(w0))))
        try:
            resample = Image.Resampling.LANCZOS
            dither_none = Image.Dither.NONE
        except AttributeError:
            resample = Image.LANCZOS
            dither_none = Image.NONE
        img = img.resize((wd, hd), resample)
        img = img.point(lambda x: 0 if x < int(thresh) else 255)
        img = img.convert("1", dither=dither_none)
        if ck is not None:
            _cache_put(_ICON_CACHE, ck, (img, wd, hd), 192)
        return img, wd, hd
    except Exception:
        return None, 0, 0


def text_bitmap_image(text, font_px, thresh=170):
    """
    把一段文字用 TTF 渲染成 1-bit 位图（宽高贴合文字外框）。
    为什么要它：TSPL 内置中文字库只有 8/12/16/24/32 五档，
      用户输入 20 这类自定义字号时 TEXT 命令无法精确出字，
      改走 BITMAP —— 直发和图片模式打出来完全一致。
    返回 PIL 1-bit 图；失败返回 None。
    """
    try:
        from PIL import Image, ImageDraw, ImageFont
        ttf = pick_ttf()
        if not ttf:
            return None
        f = ImageFont.truetype(ttf, max(6, int(round(font_px))))
        bbox = f.getbbox(text)
        w = max(1, bbox[2] - bbox[0])
        h = max(1, bbox[3] - bbox[1])
        img = Image.new("L", (w, h), 255)
        ImageDraw.Draw(img).text((-bbox[0], -bbox[1]), text, font=f, fill=0)
        img = img.point(lambda x: 0 if x < int(thresh) else 255)
        try:
            dn = Image.Dither.NONE
        except AttributeError:
            dn = Image.NONE
        return img.convert("1", dither=dn)
    except Exception:
        return None


def text_bitmap_with_pos(text, font_px, thresh=170):
    """
    同 text_bitmap_image，额外返回墨迹相对「行顶左端」的偏移 (dx, dy)。
    为什么要偏移：位图是紧贴墨迹裁的，直接放在行顶会偏高；
      图片模式和预览都按「行顶 + 字体上升部」排版，位图得跟随同样的偏移才对齐。
    """
    from PIL import ImageFont
    img = text_bitmap_image(text, font_px, thresh)
    if img is None:
        return None, 0, 0
    try:
        f = ImageFont.truetype(pick_ttf(), max(6, int(round(font_px))))
        bb = f.getbbox(text)
        return img, bb[0], bb[1]
    except Exception:
        return img, 0, 0


def bitmap_payload(img):
    """PIL 1-bit 图 → (TSPL BITMAP 的 hex 数据, 每行字节数)"""
    if img is None:
        return "", 0
    raw = img.tobytes()
    # PIL 里 1=白 0=黑；TSPL 里 1=打印(黑)，所以要取反
    inv = bytes(b ^ 0xFF for b in raw)
    return inv.hex().upper(), (img.size[0] + 7) // 8


def image_to_tspl_bitmap(path, width_dots, threshold=160):
    """图片 → TSPL BITMAP 数据（保留给旧的「整块图片」用法）"""
    img, wd, hd = prep_icon(path, width_dots)
    hexdata, wb = bitmap_payload(img)
    return hexdata, wb, wd, hd


def txt_w_dots(s, font_dots):
    """算文字在 TSSxx.BF2 字库下占多少 dot：全角 = font_dots，半角 = 一半"""
    w = 0
    for ch in s:
        w += font_dots if ord(ch) > 127 else max(1, font_dots // 2)
    return w


_TTF_W_CACHE = {}


def ttf_w_dots(s, font_dots):
    """用系统字体量真实宽度（dot）—— **排版折行和横向溢出检测必须用它**

    为什么不能只用 txt_w_dots：那是 TSPL 内置字库的估算（半角 = 字号一半），
    可屏幕上和图片模式用的是 TTF，实测半角比 font/2 宽 5~10%。
    按估算折行，最后一行实际会顶出右边距 —— 压测里「超长号型 / 长英文」
    横向画到 x=279（纸宽 280），而估算认为还有 2mm 余量，一声不响。
    """
    ttf = pick_ttf()
    if not ttf:
        return txt_w_dots(s, font_dots)
    key = (s, font_dots)
    hit = _TTF_W_CACHE.get(key)
    if hit is not None:
        return hit
    try:
        from PIL import ImageFont
        f = _TTF_W_CACHE.get(("__font__", font_dots))
        if f is None:
            f = ImageFont.truetype(ttf, max(6, int(font_dots)))
            _cache_put(_TTF_W_CACHE, ("__font__", font_dots), f, 16)
        w = int(round(f.getlength(s)))
    except Exception:
        w = txt_w_dots(s, font_dots)
    _cache_put(_TTF_W_CACHE, key, w, 400)
    return w


def icon_rows(n, max_cols=5):
    """n 个图标怎么分排 —— 排数尽量少，每排数量尽量平均，多的排靠后。

    实测对照（壮哥给的样例全对上了）：
        1~5 → [n]          一排到底
        6   → [3, 3]       3×2 两排
        7   → [3, 4]       一排 3 + 一排 4
        8   → [4, 4]       4×2 两排
        9   → [4, 5]   10 → [5, 5]   11 → [3, 4, 4]   12 → [4, 4, 4]
    为什么余数给后面的排：第一排永远不满，视觉重心在下面一排，
    跟标签「上轻下重」的读法一致；也让 7 个落在 3+4 而不是 4+3。
    """
    try:
        n = int(n)
    except Exception:
        return []
    if n <= 0:
        return []
    try:
        mc = int(max_cols) or 5
    except Exception:
        mc = 5
    mc = max(1, mc)
    rows = (n + mc - 1) // mc
    base, rem = divmod(n, rows)
    return [base] * (rows - rem) + [base + 1] * rem


def icon_grid_mm(proj: dict) -> dict:
    """图标槽位网格的实际尺寸（mm）。预览、打印、界面提示都问这一份。

    返回：
        rows       [3, 4]            每排几个
        n          总数
        slot       实际槽位边长（纸太窄时会等比缩小）
        gap        实际横向间隙
        row_gap    实际行距
        row_w      [w1, w2]          每排的总宽（含间隙）
        block_h    整块高度
        region     可用宽
        need       按「名义槽位」算出来的最宽一排
        shrunk     True = 纸太窄，槽位被压小了
    """
    proj = normalize(proj)
    b = proj.get("bottom") or {}
    paper = proj.get("paper") or {}
    w_mm = float(paper.get("w", 35) or 35)
    left_mm = float(proj.get("left_mm", 2.0) or 0.0)
    right_mm = float(proj.get("right_mm", 2.0) or 0.0)
    region = max(4.0, w_mm - left_mm - right_mm)

    n = len([p for p in (b.get("icons") or []) if str(p).strip()])
    rows = icon_rows(n, int(b.get("icon_max_cols", 5) or 5))
    slot = max(1.0, min(40.0, float(b.get("icon_slot_mm", 5.0) or 5.0)))
    gap = max(0.0, min(20.0, float(b.get("icon_slot_gap_mm", 0.6) or 0.0)))
    row_gap = max(0.0, min(20.0, float(b.get("icon_row_gap_mm", 1.0) or 0.0)))

    def width_of(k, s, g):
        return k * s + (k - 1) * g if k > 0 else 0.0

    need = max([width_of(k, slot, gap) for k in rows] or [0.0])
    need_nominal = need          # 名义值：纸太窄要压的时候，界面得能说出「本来要多少」
    shrunk = need > region + 1e-6
    if shrunk and need > 0:
        # 纸太窄，一排根本放不下 —— 连槽位带间隙一起等比压，别让它越出纸边
        f = region / need
        slot *= f
        gap *= f
        need = region
    return {
        "rows": rows, "n": n, "slot": slot, "gap": gap, "row_gap": row_gap,
        "row_w": [width_of(k, slot, gap) for k in rows],
        "block_h": (len(rows) * slot + max(0, len(rows) - 1) * row_gap) if rows else 0.0,
        "region": region, "need": need, "need_nominal": need_nominal,
        "shrunk": shrunk,
    }


def bottom_ops(proj: dict):
    """底部区块绘制清单（带记忆化）。

    同一个重绘周期里这份清单会被问 4 次（自己一次 + bottom_top_mm +
    content_overflow + suggested_h_mm 各一次），而算它需要解码每个图标。
    这里按「bottom 内容 + 纸张 + 左右边距」做签名缓存，返回副本防调用方改动。"""
    proj = normalize(proj)
    b = proj.get("bottom") or {}
    if not b.get("on", True):
        return [], 0
    try:
        # 日期那几项在项目顶层、不在 bottom 里，必须单独进签名 ——
        # 否则洗涤块内容没动、只勾了「印生产日期」，签名不变会直接吃旧缓存，
        # 界面预览看着像「勾了没反应」。
        key = (json.dumps(b, sort_keys=True, ensure_ascii=False),
               json.dumps(proj.get("paper"), sort_keys=True),
               proj.get("left_mm"), proj.get("right_mm"),
               bool(proj.get("date_on")),
               proj.get("date_lock"), proj.get("date_month"),
               proj.get("date_day"), proj.get("date_font"),
               # 自定义内容 / 对齐也是顶层的，漏进签名就会吃旧缓存：
               # 改了对齐却还印在原来那个 x 上（实测三种对齐全返回居中值）。
               proj.get("date_custom"), proj.get("date_align"))
    except (TypeError, ValueError):
        return _bottom_ops_impl(proj)
    hit = _BOTTOM_OPS_CACHE.get(key)
    if hit is None:
        hit = _bottom_ops_impl(proj)
        _cache_put(_BOTTOM_OPS_CACHE, key, hit, 24)
    ops, h = hit
    return [dict(o) for o in ops], h


def _bottom_ops_impl(proj: dict):
    """
    底部固定区块的绘制清单 —— 文字是 TSPL 原生文本，图标是独立位图。

    返回 (ops, total_h_dot)
    ops 元素：
      {"kind":"text",   "x":..,"y":..,"text":..,"font":16/24/32}
      {"kind":"bitmap", "x":..,"y":..,"path":..,"w":..,"h":..}
    x 以纸张左边为 0，y 以「区块顶端」为 0，单位都是 dot。
    预览和打印都调用这一份，保证两边一致。
    """
    proj = normalize(proj)
    b = proj.get("bottom") or {}
    if not b.get("on", True):
        return [], 0
    set_font(proj.get("font_file"))          # 区块文字居中按字宽算，同样要认字体

    paper = proj["paper"]
    dpi = int(paper.get("dpi", 203))
    w_mm = float(paper.get("w", 35))
    left_mm = float(proj.get("left_mm", 2.0))
    right_mm = float(proj.get("right_mm", 2.0))
    x0 = mm2dot(left_mm, dpi)
    # 可用宽度 = 纸宽 - 左右两侧留白
    region = max(16, mm2dot(max(4.0, w_mm - left_mm - right_mm), dpi))

    f_title = int(b.get("title_font", 24))
    f_sub = int(b.get("sub_font", 16))
    f_body = int(b.get("body_font", 16))
    gap1 = float(b.get("gap_mm", 1.0))
    gap2 = float(b.get("gap2_mm", 1.2))
    lgap = float(b.get("line_gap_mm", 0.3))
    icon_th = int(b.get("icon_th", 170))

    ops = []
    y = 0

    def add_text(s, f, gap_mm=0.0, align="center"):
        nonlocal y
        y += mm2dot(gap_mm, dpi)
        tw = txt_w_dots(s, f)
        if align == "left":
            x = x0
        elif align == "right":
            x = x0 + max(0, region - tw)
        else:                                   # 默认 / 老行为：居中
            x = x0 + max(0, (region - tw) // 2)
        ops.append({"kind": "text", "font": f, "text": s, "x": x, "y": y})
        y += f

    title = str(b.get("title", "")).strip()
    if title:
        add_text(title, f_title)

    sub = str(b.get("subtitle", "")).strip()
    if sub:
        add_text(sub, f_sub, gap1)

    icons = [str(p) for p in (b.get("icons") or []) if str(p).strip()]
    if icons:
        # 副标题→图标 的间距单独给一个键：底版里这段是 2.8mm，跟标题与副标题之间的
        # 0.8mm 差得远，共用一个 gap 就永远对不上。没写就退回 gap1（老行为）。
        y += mm2dot(float(b.get("icon_gap_mm", gap1)), dpi)
        raw = []
        for rel in icons:
            p = asset_path(rel)
            if not os.path.exists(p):
                continue
            w0, h0 = img_size(p)
            if w0 > 0 and h0 > 0:
                raw.append((p, w0, h0))
        if raw:
            g = icon_grid_mm(proj)
            slot_dot = mm2dot(g["slot"], dpi)
            gap_dot = mm2dot(g["gap"], dpi)
            row_gap_dot = mm2dot(g["row_gap"], dpi)
            # 分排按「实际能画出来的张数」算，不能按列表长度 —— 缺图的那些要占位
            plan = icon_rows(len(raw), int(b.get("icon_max_cols", 5) or 5))
            i = 0
            for r, cnt in enumerate(plan):
                seg = raw[i:i + cnt]
                i += cnt
                if not seg:
                    continue
                # 这一排的起点：整块居中。为什么不是左对齐 ——
                # 两排数量不一样（3+4）时左对齐会歪，居中才像「同一套图标」
                row_w = len(seg) * slot_dot + max(0, len(seg) - 1) * gap_dot
                xc = x0 + max(0, (region - row_w) // 2)
                for (p, w0, h0) in seg:
                    # 槽位是正方形固定框，图片只在里面「等比缩进来」，绝不放大：
                    # 放大等于用更少的像素硬撑，边缘全是锯齿；也绝不拉伸，
                    # 图标一旦变形（洗衣盆被压扁）读起来就不对了。
                    k = min(slot_dot / float(w0), slot_dot / float(h0), 1.0)
                    wd = max(1, int(round(w0 * k)))
                    # 按宽取整后高度可能刚好越过槽一两个 dot（竖长图尤其明显：
                    # 60×900 取整到宽 3，高就变成 45，比 5mm 的槽还高）。
                    # 越了就按高度反推宽度，向下取整 —— 宁可少半个 dot，不许出格。
                    if int(round(wd * h0 / float(w0))) > slot_dot:
                        wd = max(1, int(slot_dot * w0 / float(h0)))
                    img, wd2, hd = prep_icon(p, wd, icon_th)
                    if img is not None:
                        ops.append({"kind": "bitmap", "path": p, "w": wd2, "h": hd,
                                    "th": icon_th,
                                    "x": xc + max(0, (slot_dot - wd2) // 2),
                                    "y": y + r * (slot_dot + row_gap_dot)
                                    + max(0, (slot_dot - hd) // 2)})
                    xc += slot_dot + gap_dot
            # 整块高度按「排数 × 槽位」算，不按最高的一张图算：
            # 槽位固定，图标矮一点也不该让下一排往上跑，否则换个图版式又跳了
            y += len(plan) * slot_dot + max(0, len(plan) - 1) * row_gap_dot

    caut = [str(s).strip() for s in (b.get("caution") or []) if str(s).strip()]
    if caut:
        y += mm2dot(gap2, dpi)
        for s in caut:
            add_text(s, f_body)
            y += mm2dot(lgap, dpi)
        y -= mm2dot(lgap, dpi)                        # 去掉最后一行的行距

    # 页脚（MADE IN CHINA 这种）：字号独立，间距也独立，跟在注意事项后面
    footer = str(b.get("footer", "")).strip()
    if footer:
        f_foot = max(8, min(72, int(b.get("footer_font", 16))))
        add_text(footer, f_foot, float(b.get("footer_gap_mm", 1.2)))

    # 生产日期（NOV 18）：整张标签最底下一行，居中。
    # 为什么挂在洗涤块末尾而不是正文下面：水洗标最底部本来就是「落款」的位置
    # （产地 / 制造地 / 日期是同一类信息），贴底排版时那块空白不用白不用，
    # 正文也就不必为它让位。字号独立，和页脚一样走 add_text 的居中。
    if proj.get("date_on"):
        s_date = date_text(proj)
        f_date = max(8, min(48, int(proj.get("date_font", 12) or 12)))
        # 自定义内容可能很长（批次号、中文落款…），按可印宽度自动缩字号，
        # 别让它横着捅出纸边。自动日期「NOV 24」短，永远缩不到，行为不变。
        while f_date > 8 and txt_w_dots(s_date, f_date) > region:
            f_date -= 1
        add_text(s_date, f_date,
                 float(b.get("date_gap_mm", 1.2) or 0.0),
                 align=str(proj.get("date_align", "center") or "center"))

    return ops, y


def bottom_height_mm(proj: dict):
    """底部区块总高度（mm），给预览做防重叠检查用"""
    proj = normalize(proj)
    dpi = float(proj["paper"].get("dpi", 203)) or 203.0
    _, h = bottom_ops(proj)
    return h * 25.4 / dpi


def _normalize_saved_project(proj: dict) -> dict:
    """补全缺省字段，保证旧文件也能打开（嵌套字典做合并，不整体替换）"""
    p = json.loads(json.dumps(DEFAULT_PROJECT))
    proj = proj or {}
    for k, v in proj.items():
        if k in ("paper", "top_rule", "bottom", "bottom_image", "brand", "table") and isinstance(v, dict):
            p.setdefault(k, {}).update(v)
        else:
            p[k] = v
    # 没选打印机就是没选：别默默继承 DEFAULT_PROJECT 里写死的那台机器，
    # 否则一个没选过打印机的项目会被打到一台可能根本不存在的打印机上
    p["printer"] = str(proj.get("printer") or "").strip()
    # 旧项目只有整块图片 bottom_image，把它开关迁移到新的 bottom 结构
    if "bottom" not in proj and isinstance(proj.get("bottom_image"), dict):
        p["bottom"]["on"] = bool(proj["bottom_image"].get("on", True))
    if not isinstance(p.get("fabrics"), list):
        p["fabrics"] = str(p.get("fabrics", "")).splitlines()
    # 成分里的空行要不要真的空出一行（分段用）。默认关：老文件照旧不留空档
    p["fabric_keep_blank"] = bool(p.get("fabric_keep_blank", False))
    # 纸长要不要跟着内容长。默认开：成分多敲几行就把纸顶穿，是最常见的翻车现场，
    # 让它自己长比让用户每次手算纸长靠谱。（老文件里存过这个键的，仍以自己的值为准）
    p["auto_paper"] = bool(p.get("auto_paper", True))
    # 纸张形态：只认两个值，坏值一律当间隙纸（老文件没有这个键，走这里）
    _media = str((p.get("paper") or {}).get("media") or "gap").lower()
    p.setdefault("paper", {})["media"] = _media if _media in ("gap", "continuous") else "gap"
    # 打印位移 / 方向：统一夹到合法区间，坏值不让它把版面打飞
    def _num(v, d=0.0):
        try:
            return float(v)
        except Exception:
            return d
    p["off_x_mm"] = max(-50.0, min(50.0, _num(p.get("off_x_mm"), 0.0)))
    p["off_y_mm"] = max(-50.0, min(50.0, _num(p.get("off_y_mm"), 0.0)))
    try:
        ang = int(float(p.get("direction") or 0)) % 360
    except Exception:
        ang = 0
    p["direction"] = ang if ang in (0, 90, 180, 270) else 0

    # 图标槽位：夹到能用又不离谱的区间。旧文件没有这几个键，走上面的默认值；
    # 有了就按它排版 —— 老文件里的 icon_h_max_mm / icon_gap_ratio 一律不再参与。
    bt = p.setdefault("bottom", {})
    try:
        bt["icon_slot_mm"] = max(1.0, min(40.0, float(bt.get("icon_slot_mm", 5.0) or 5.0)))
    except Exception:
        bt["icon_slot_mm"] = 5.0
    try:
        bt["icon_slot_gap_mm"] = max(0.0, min(20.0, float(bt.get("icon_slot_gap_mm", 0.6) or 0.0)))
    except Exception:
        bt["icon_slot_gap_mm"] = 0.6
    try:
        bt["icon_row_gap_mm"] = max(0.0, min(20.0, float(bt.get("icon_row_gap_mm", 1.0) or 0.0)))
    except Exception:
        bt["icon_row_gap_mm"] = 1.0
    try:
        bt["icon_max_cols"] = max(1, min(8, int(float(bt.get("icon_max_cols", 5) or 5))))
    except Exception:
        bt["icon_max_cols"] = 5
    return p


# ---------------- TSPL 生成 ----------------
FONT_ALIAS = {12: "TSS12.BF2", 16: "TSS16.BF2", 24: "TSS24.BF2", 32: "TSS32.BF2"}


def mm2dot(mm, dpi):
    return int(round(mm * dpi / 25.4))


def layout(proj: dict):
    """
    排版核心 —— 返回 [(x_mm, y_mm, text), ...]
    预览画布和 TSPL 指令都调用这一份，保证「屏幕上看到的」= 「打出来的」。

    行为：
      · 品牌（可开关）在最上面，然后号型、成分
      · 成分不再做固定缩进对齐 —— 所有行一律顶左边距
      · 任何一行超出右边距就自动折行（按全角/半角字宽算，逐字贪心断行）
    """
    proj = normalize(proj)
    # 字宽要跟项目指定的字体走：这里不设的话，先调 layout 的调用方（状态行、
    # 侧溢出检查）会拿上一个字体的宽度去算，换字体后那几个数字就是错的。
    set_font(proj.get("font_file"))
    dpi = float(proj["paper"].get("dpi", 203))
    w_mm = float(proj["paper"].get("w", 35))
    font_dots = int(proj.get("font_dots", 24))
    font_mm = font_dots * 25.4 / dpi
    gap_mm = int(proj.get("line_gap", 4)) * 25.4 / dpi
    line_h = font_mm + gap_mm

    x0 = float(proj.get("left_mm", 2.0)) + max(0.0, float(proj.get("body_indent_mm", 0.0)))
    # 可用宽度 = 纸宽 - 左右两侧留白（缩进从可用宽度里扣，否则右边会被顶出去）
    region_mm = max(5.0, w_mm - x0 - float(proj.get("right_mm", 2.0)))

    def wrap(s):
        """逐字贪心断行：放不下就换行，保证永远不超出右边距"""
        out, cur = [], ""
        for ch in s:
            if cur and (ttf_w_dots(cur + ch, font_dots) * 25.4 / dpi) > region_mm:
                out.append(cur)
                cur = ch
            else:
                cur += ch
        if cur:
            out.append(cur)
        return out

    y = float(proj.get("start_y_mm", 20.0))
    out = []
    align = str(proj.get("text_align", "left") or "left").lower()

    def emit(s):
        nonlocal y
        # Tab 在点阵字体里量不出宽度（实测渲染宽度为 0，敲了等于没敲），
        # 按 4 个空格折一下 —— 用户按 Tab 缩进和敲 4 个空格，效果应该一样。
        s = str(s).replace("\t", "    ")
        for seg in wrap(s):
            # 纯空格段不占行。行首空格超出行宽时会被 wrap 拆出「一整行空格」，
            # 那一行在纸上什么也没有，却把下面所有内容往下推 —— 白占高度。
            if not seg.strip():
                continue
            x = x0
            if align == "center":
                # 居中按「这一行自己的宽度」算，不是按最长那行算 ——
                # 否则短行会跟长行共用同一条中轴以外的位置，看着像没对齐。
                tw_mm = txt_w_dots(seg, font_dots) * 25.4 / dpi
                x = x0 + max(0.0, (region_mm - tw_mm) / 2.0)
            out.append((x, y, seg))
            y += line_h

    br = proj.get("brand") or {}
    if br.get("on", True):
        # 不 strip：前导空格是排版手艺 —— 想在标签上把这一行往右挪，
        # 敲空格就该挪，别替用户把空格吃掉。（只拿 strip 后的值判断「是不是空的」）
        btxt = str(br.get("text", ""))
        if btxt.strip():
            emit(btxt)

    title = str(proj.get("title", ""))
    if title.strip():
        emit(title)

    head = str(proj.get("fabric_head", "成分:"))
    # rstrip 而不是 strip：行尾空格看不见，留着只会让换行判定莫名超宽；
    # 行首空格看得见，是要保留的排版信息。
    keep_blank = bool(proj.get("fabric_keep_blank", False))
    fabrics = [str(f).rstrip() for f in proj.get("fabrics", [])]
    # 末尾的空条目永远不占位：Text 结尾那个换行会带出一个空串，
    # 开了「保留空行」也不能因此每张标签底下凭空多一行。
    while fabrics and not fabrics[-1].strip():
        fabrics.pop()
    first_done = False
    for f in fabrics:
        if not f.strip():
            # 空行：开了开关才真的空出一行（分段用），否则当误输入丢掉
            if keep_blank:
                y += line_h
            continue
        # 首行带「成分:」前缀。前缀留空时直接输出原行 ——
        # 从老模板导入多段成分（面料/胆料/填充物自带标签）时要把前缀清空，
        # 否则会出现「成分: 面料: 100%锦纶」这种双标签；这时候也不能顺手把
        # 用户加的行首空格 strip 掉。
        # 「首行」指第一条非空的，前面若有空行，前缀不该落到那一行上。
        if not first_done and head.strip():
            emit("%s %s" % (head, f))
        else:
            emit(f)
        first_done = True

    # 产地：成分之后单独一行。空着就完全不占位 ——
    # 不能因为多了个字段就让所有老文件底下凭空多一行空白。
    origin = str(proj.get("origin", "")).replace("\t", "    ")
    if origin.strip():
        emit(origin)

    return out


# 英文月份简称。只用三个字母：水洗标上这一行是小字，全称占不下也没人这么排。
MONTH_EN = ("JAN", "FEB", "MAR", "APR", "MAY", "JUN",
            "JUL", "AUG", "SEP", "OCT", "NOV", "DEC")

# 界面下拉里的显示名：前面带月份数字，方便不懂英文的人对照着选（01 JAN … 12 DEC）。
# 数字**只给人看**，打印出去永远是 MONTH_EN 里的三字母 —— 一行小字里塞不下 "11"。
MONTH_CHOICES = tuple("%02d %s" % (i + 1, m) for i, m in enumerate(MONTH_EN))


def month_of(text):
    """把下拉里选中的那一项解析回 1~12。

    认三种写法：'01 JAN'（现在的下拉）、'NOV'（老文件/老版本）、'11'
    （有人直接填数字）。认不出返回 0，调用方按「取当天」处理。"""
    s = str(text or "").strip().upper()
    if not s:
        return 0
    tail = s.split()[-1]
    if tail in MONTH_EN:
        return MONTH_EN.index(tail) + 1
    if s.isdigit() and 1 <= int(s) <= 12:
        return int(s)
    return 0


def date_text(proj: dict, today=None) -> str:
    """生产日期文案：英文月简称 + 日，如 NOV 18。

    月份 / 日期都是**存在项目里的值**（界面上选、手输），不是每次现取当天 ——
    这样补打同一批货，纸上日期不会漂。（v1.4 那半天存的 date_lock 字符串仍然认。）
    date_month / date_day 为 0 表示没选过，取当天；today 只给测试用。
    """
    p = normalize(proj)
    # 自定义内容优先：填了就原样印，不跟月份/日拼在一起。
    # 留空才走自动日期，所以老文件（没这个键）行为完全不变。
    custom = str(p.get("date_custom", "") or "").strip()
    if custom:
        return custom
    lock = str(p.get("date_lock", "") or "").strip()
    if lock:
        return lock
    if today is None:
        today = time.localtime()
    try:
        m = int(p.get("date_month", 0) or 0)
    except (TypeError, ValueError):
        m = 0
    try:
        d = int(p.get("date_day", 0) or 0)
    except (TypeError, ValueError):
        d = 0
    if m < 1 or m > 12:
        m = today.tm_mon
    if d < 1:
        d = today.tm_mday
    d = max(1, min(31, d))
    return "%s %d" % (MONTH_EN[m - 1], d)


def date_place(proj: dict, lines=None):
    """日期落在哪 —— 返回 (x_mm, y_mm, font_dots) 或 None。

    【仅用于「洗涤说明块被关掉」这一种情况】：正常路径下日期是洗涤块的最后
    一行，由 _bottom_ops_impl 里的 add_text 负责居中，不走这里。
    关掉洗涤块以后块高为 0，没地方挂，就贴纸底保留线往上让一行。
    """
    proj = normalize(proj)
    if not proj.get("date_on"):
        return None
    b = proj.get("bottom") or {}
    if b.get("on", True):
        return None
    dpi = float(proj["paper"].get("dpi", 203))
    d2mm = 25.4 / dpi
    w_mm = float(proj["paper"].get("w", 35))
    h_mm = float(proj["paper"].get("h", 75))
    s_date = date_text(proj)
    f = max(8, min(48, int(proj.get("date_font", 12) or 12)))
    avail = max(1.0, w_mm - float(proj.get("left_mm", 2.0))
                - float(proj.get("right_mm", 2.0)))
    while f > 8 and ttf_w_dots(s_date, f) * d2mm > avail:
        f -= 1
    tw = ttf_w_dots(s_date, f) * d2mm
    y = h_mm - float(b.get("margin_mm", 5.0)) - f * d2mm
    align = str(proj.get("date_align", "center") or "center")
    if align == "left":
        x = float(proj.get("left_mm", 2.0))
    elif align == "right":
        x = max(0.0, w_mm - float(proj.get("right_mm", 2.0)) - tw)
    else:
        x = max(0.0, (w_mm - tw) / 2.0)
    return (x, max(0.0, y), f)


def text_bottom_mm(proj: dict) -> float:
    """文字区（品牌/号型/成分）最底部的 y（mm），底部区块用它来定位"""
    proj = normalize(proj)
    lines = layout(proj)
    if not lines:
        return float(proj.get("start_y_mm", 20.0))
    dpi = float(proj["paper"].get("dpi", 203))
    font_mm = int(proj.get("font_dots", 24)) * 25.4 / dpi
    return max(ly for (_, ly, _) in lines) + font_mm


def table_rows(proj: dict):
    """规范化后的表格行：[{cols, grid, font, cells}]，字库只认 12/16/24/32

    grid = 这一行铺在几列网格上（默认等于 cols）。
    为什么需要它：7 个尺码要排成「一排 4 + 一排 3」，如果下排按自己的 3 列等分，
    格子会比上排宽、竖线从上到下全断，看着是歪的。把两行的 grid 都定成 4，
    上排填 4 格、下排填 3 格 + 右下一个空格，列宽统一、竖线贯通。
    """
    proj = normalize(proj)
    t = proj.get("table") or {}
    out = []
    for r in (t.get("rows") or []):
        try:
            # 列数不再写死 3/4：1 列（只勾了「均码」这一个尺码）、3 列（尺码 M/L/XL）、
            # 4 列（带颜色）、5 列（尺码表最常见）都是真实需求。
            # 上限 8 留住，免得脏数据把每一格压成几个像素。
            cols = max(1, min(8, int(r.get("cols", 4))))
        except Exception:
            cols = 4
        try:
            grid = max(cols, min(8, int(r.get("grid", cols))))
        except Exception:
            grid = cols
        cells = [str(c) for c in (r.get("cells") or [])][:cols]
        cells += [""] * (cols - len(cells))
        try:
            f = int(round(float(r.get("font", 16))))
        except Exception:
            f = 16
        f = max(8, min(72, f))          # 自定义字号：8~72 dot（203dpi 下约 1~9mm 高）
        # 每行可以有自己行高：尺码行 3mm 就够，下面的空白行要给手写留 4.5~5mm。
        # 没写或写坏了就跟着表格全局行高走（老文件、手动加的行都是这种）。
        def _num(v, d):
            try:
                return float(v)
            except (TypeError, ValueError):
                return d
        base_h = _num(t.get("row_h_mm"), 4.5)
        rh = _num(r.get("row_h_mm"), base_h)
        rh = max(2.0, min(20.0, rh))
        out.append({"cols": cols, "grid": grid, "font": f, "row_h_mm": rh,
                    "cells": cells})
    return out


def size_table_plan(n: int):
    """尺码 → 排法。返回 (每排格数列表, 网格列数)。

    规则：5 个及以下一排铺满；6 个以上每排最多 4 个，尽量均分、多的排前面。
      3→[3]  4→[4]  5→[5]  6→[3,3]  7→[4,3]  8→[4,4]
      9→[3,3,3]  10→[4,3,3]  11→[4,4,3]  12→[4,4,4]
    """
    n = int(n)
    if n <= 0:
        return [], 0
    if n <= 5:
        return [n], n
    rows_n = (n + 3) // 4                 # ceil(n/4)
    base, extra = divmod(n, rows_n)       # 多出来的摊到前几排
    plan = [base + 1] * extra + [base] * (rows_n - extra)
    return plan, max(plan)


def size_table_rows(sizes, values=None, h_title=3.0, h_blank=3.5, font=16):
    """一串尺码（+ 一串数值）→ 表格行。

    走的是普通表格行，所以生成出来照样能改字、调行高、单独删行；
    多组共用同一个 grid，4+3 这种上下排的列宽一致、竖线贯通，
    不足的那一组右下角留空格子（画线不填字）。

    values 按打印顺序一段一段分下去：给少了后面的格子留空，给多了直接截掉
    （调用方负责在界面上提醒，别在这里默默吞掉用户输入）。
    """
    sizes = [str(s) for s in (sizes or [])]
    vals = [str(v) for v in (values or [])]
    plan, grid = size_table_plan(len(sizes))
    rows = []
    i = k = 0
    for g in plan:
        grp = sizes[i:i + g]
        i += g
        seg = vals[k:k + g]
        k += g
        blank = [seg[j] if j < len(seg) else "" for j in range(g)]
        rows.append({"cols": g, "grid": grid, "font": font,
                     "row_h_mm": h_title, "cells": list(grp)})
        rows.append({"cols": g, "grid": grid, "font": font,
                     "row_h_mm": h_blank, "cells": blank})
    return rows


def table_on(proj: dict) -> bool:
    proj = normalize(proj)
    return bool((proj.get("table") or {}).get("on")) and len(table_rows(proj)) > 0


def table_ops(proj: dict):
    """
    表格绘制清单。返回 (ops, height_dot)
      ops: {"kind":"bar","x","y","w","h"} / {"kind":"text","x","y","text","font"}
         x = 纸张绝对 dot；y = 相对表格顶端 dot
    规则：宽度铺满左右边距之间；按 grid 网格等分（默认 grid = 该行列数）；
          cells 只填前 cols 格，网格剩余的格子留空但线照画（尺码 4+3 靠这个对齐）；
          每行高度取该行自己的 row_h_mm，没写就跟着表格全局行高；
          单元格文本水平+垂直居中；
          四周边框 + 行线 + 列线全部实色（BAR 命令，不经过驱动渲染）。
    预览、TSPL、图片模式都调用这一份，三边一致。
    """
    proj = normalize(proj)
    t = proj.get("table") or {}
    rows = table_rows(proj)
    if not rows or not t.get("on"):
        return [], 0

    dpi = int(proj["paper"].get("dpi", 203))
    w_mm = float(proj["paper"].get("w", 35))
    left_mm = float(proj.get("left_mm", 2.0))
    right_mm = float(proj.get("right_mm", 2.0))
    x0 = mm2dot(left_mm, dpi)
    region = max(8, mm2dot(max(4.0, w_mm - left_mm - right_mm), dpi))
    bw = max(1, int(t.get("border_dot", 2)))

    ops = []
    y = 0

    # 表格上方小标题（居中，字号可任意）
    cap = str(t.get("caption", "")).strip()
    if cap:
        cap_f = max(8, min(72, int(t.get("caption_font", 16))))
        tw = txt_w_dots(cap, cap_f)
        ops.append({"kind": "caption", "x": x0 + max(0, (region - tw) // 2),
                    "y": 0, "text": cap, "font": cap_f,
                    "cell_x": x0, "cell_w": region})
        y += cap_f + mm2dot(float(t.get("caption_gap_mm", 0.5)), dpi)

    for r in rows:
        cols = r["cols"]
        grid = r.get("grid", cols)
        row_h = max(8, mm2dot(r.get("row_h_mm", 4.5), dpi))     # 这一行自己的高度
        # 网格列数决定竖线密度；cols 只决定这一行填几格（多余的格子留空）
        edges = [x0 + (region * i) // grid for i in range(grid)] + [x0 + region]
        ops.append({"kind": "bar", "x": x0, "y": y, "w": region, "h": bw})   # 行顶线
        for i in range(grid + 1):                                            # 竖线
            vx = edges[i] if i < grid else (x0 + region - bw)
            ops.append({"kind": "bar", "x": vx, "y": y, "w": bw, "h": row_h})
        for i, txt in enumerate(r["cells"]):                                 # 单元格文字
            s = str(txt).strip()
            if not s:
                continue
            cw = edges[i + 1] - edges[i]
            # 文字太长就逐步降字号（每次 2 dot，最低 8），保证不出格
            avail_w = cw - 2 * bw - 2
            f = r["font"]
            while f > 8 and txt_w_dots(s, f) > avail_w:
                f -= 2
            tw = txt_w_dots(s, f)
            tx = edges[i] + max(bw + 1, (cw - tw) // 2)
            ty = y + max(bw, (row_h - f) // 2)
            ops.append({"kind": "text", "x": tx, "y": ty, "text": s, "font": f,
                        "cell_x": edges[i], "cell_w": cw})
        y += row_h
    ops.append({"kind": "bar", "x": x0, "y": max(0, y - bw), "w": region, "h": bw})  # 底线
    return ops, y


def table_height_mm(proj: dict) -> float:
    _, h = table_ops(proj)
    dpi = float(normalize(proj)["paper"].get("dpi", 203)) or 203.0
    return h * 25.4 / dpi


def table_top_mm(proj: dict) -> float:
    """表格顶端 = 正文最后一行的底 + 上留白（默认 2mm）"""
    proj = normalize(proj)
    t = proj.get("table") or {}
    return text_bottom_mm(proj) + float(t.get("gap_top_mm", 2.0))


def table_bottom_mm(proj: dict) -> float:
    proj = normalize(proj)
    if not table_on(proj):
        return text_bottom_mm(proj)
    return table_top_mm(proj) + table_height_mm(proj)


def content_bottom_mm(proj: dict) -> float:
    """正文 + 表格 的最底部 —— 底部洗涤说明块据此避让"""
    proj = normalize(proj)
    if table_on(proj):
        return table_bottom_mm(proj)
    return text_bottom_mm(proj)


def table_gap_below(proj: dict) -> float:
    """表格底边到下方区块顶边的间距，默认 4mm —— 由「距洗涤方式」这个框直接决定"""
    proj = normalize(proj)
    if not table_on(proj):
        return 0.0
    t = proj.get("table") or {}
    if "to_bottom_mm" in t:
        try:
            return max(0.0, float(t["to_bottom_mm"]))
        except (TypeError, ValueError):
            return 4.0
    return max(0.0, float(t.get("gap_bottom_mm", 4.0)))


def gap_before_bottom_mm(proj: dict) -> float:
    """正文/表格底边 → 洗涤方式区块顶边 的间距。三个函数都走这一个口径，不许各算各的"""
    proj = normalize(proj)
    if table_on(proj):
        return table_gap_below(proj)
    return float((proj.get("bottom") or {}).get("gap_above_mm", 5.0))


def bottom_top_mm(proj: dict) -> float:
    """
    底部区块顶端的 y（mm）。

    三种位置模式（bottom.align）：
      flow   = 内容底 + 上方留白（默认，老行为）
      center = 在「内容底 + 上方留白」和「纸底可用 - 区块高」之间取中点
      bottom = 贴纸底：纸长 - 纸底保留 - 区块高

    为什么默认还是 flow：一个区块只有一个位置。贴底/居中之后，「上方留白」就从
    设定值变成了结果，再留输入框就是骗人 —— 所以只在用户明确选了模式时才接管，
    并且界面会把那一格灰掉并说明「由位置模式接管」。
    """
    proj = normalize(proj)
    b = proj.get("bottom") or {}
    flow_top = content_bottom_mm(proj) + gap_before_bottom_mm(proj)
    if not b.get("on", True):
        return flow_top
    mode = str(b.get("align", "flow") or "flow").lower()
    if mode not in ("bottom", "center"):
        return flow_top
    dpi = float(proj["paper"].get("dpi", 203)) or 203.0
    h_mm = float(proj["paper"].get("h", 75))
    _, block_h = bottom_ops(proj)
    block_mm = block_h * 25.4 / dpi
    margin = float(b.get("margin_mm", 5.0))
    # 底线只认「正文底」——不再用「上方留白」当底线。
    # 为什么：这两个模式一开，那一格就是灰的、写了也不该生效；而且它常被联动
    # 回填成「当前实际值」，一旦当成底线，贴底就会差 0.03mm 之差被顶起来报警。
    floor = content_bottom_mm(proj)
    inner = h_mm - margin - block_mm                   # 贴底时区块顶该在的位置
    if mode == "bottom":
        return max(floor, inner)
    # center：正文字底和「贴底位置」之间的中点。比 floor 还高就退回 floor，
    # 否则区块会往上顶进正文里。
    return max(floor, (floor + inner) / 2.0)


def bottom_span_mm(proj: dict):
    """返回 (区块顶 y, 实际上方留白, 纸底剩余空白)，单位 mm。
    界面用它显示状态行：上方留白是设定值，纸底余量是结果，帮用户判断纸长够不够。"""
    proj = normalize(proj)
    h_mm = float(proj["paper"].get("h", 75))
    dpi = float(proj["paper"].get("dpi", 203))
    _, block_h = bottom_ops(proj)
    block_h_mm = block_h * 25.4 / dpi
    top = bottom_top_mm(proj)
    return top, top - content_bottom_mm(proj), h_mm - (top + block_h_mm)


def page_bottom_mm(proj: dict) -> float:
    """整页内容实际占到的底部（mm）—— 正文 / 表格 / 洗涤块里最低的那个"""
    proj = normalize(proj)
    dpi = float(proj["paper"].get("dpi", 203))
    bottom = content_bottom_mm(proj)
    b = proj.get("bottom") or {}
    if b.get("on", True):
        _, block_h = bottom_ops(proj)
        bottom = max(bottom, bottom_top_mm(proj) + block_h * 25.4 / dpi)
    return bottom


def page_overflow_mm(proj: dict) -> float:
    """内容超出纸底多少 mm（>0 = 会被裁掉）。这是「会不会打不全」的唯一判据。

    和 content_overflow() 不是一回事：那个只看洗涤块有没有挤进纸底留白，
    洗涤块一关就永远返回 False —— 可正文和表格照样能把内容顶出纸外。
    压测里 6 行成分 / 72 号字 / 5 行表格都能顶穿，旧判据一声不响。
    """
    proj = normalize(proj)
    h_mm = float(proj["paper"].get("h", 75))
    return max(0.0, round(page_bottom_mm(proj) - h_mm, 2))


def page_side_overflow_mm(proj: dict) -> float:
    """正文横向超出可印区多少 mm（>0 = 左右会被裁）。

    竖向有 page_overflow_mm，横向以前完全没管：一个超长号型或者长英文，
    宽度直接顶出纸边，软件一声不响。宽度按 txt_w_dots 算，和 TSPL 直发同一个口径。
    """
    proj = normalize(proj)
    dpi = float(proj["paper"].get("dpi", 203))
    w_mm = float(proj["paper"].get("w", 35))
    limit = w_mm - float(proj.get("right_mm", 2.0))
    d2mm = 25.4 / dpi
    fd = int(proj.get("font_dots", 24))
    worst = 0.0
    for (x_mm, _y, s) in layout(proj):
        worst = max(worst, x_mm + ttf_w_dots(s, fd) * d2mm - limit)
    return max(0.0, round(worst, 2))


def page_side_margin_mm(proj: dict) -> float:
    """左右留白里较窄的那一边（mm）。小于 1mm 时，标签机的物理不可打区会切掉字"""
    proj = normalize(proj)
    return max(0.0, min(float(proj.get("left_mm", 2.0)), float(proj.get("right_mm", 2.0))))


def content_overflow(proj: dict) -> bool:
    """洗涤方式区块的底边是否越过纸底保留线（默认 5mm）——用来提示纸长不够"""
    proj = normalize(proj)
    if not (proj.get("bottom") or {}).get("on", True):
        return False
    h_mm = float(proj["paper"].get("h", 75))
    dpi = float(proj["paper"].get("dpi", 203))
    b = proj.get("bottom") or {}
    _, block_h = bottom_ops(proj)
    block_bottom = bottom_top_mm(proj) + block_h * 25.4 / dpi
    return block_bottom > h_mm - float(b.get("margin_mm", 5.0)) + 1e-6


def suggested_h_mm(proj: dict) -> float:
    """建议纸长：内容 + 表格 + 上方留白 + 区块 + 底部留白，正好让间距都成立"""
    proj = normalize(proj)
    dpi = float(proj["paper"].get("dpi", 203))
    b = proj.get("bottom") or {}
    _, block_h = bottom_ops(proj)
    if not b.get("on", True):
        block_h = 0
    return (content_bottom_mm(proj)
            + gap_before_bottom_mm(proj)
            + block_h * 25.4 / dpi + float(b.get("margin_mm", 5.0)))


# ---------------- 统一绘制清单 ----------------
# 以前同一个版式在三处各写一遍：屏幕 Canvas、PIL 位图、TSPL 指令。
# 加一个元素要改三遍，而且三遍的取整方式还不一样，容易对不齐。
# 现在只在这里生成一份 ops，三个渲染器各自消费：
#   _draw_pil_ops   → 图片模式 / 打印预览
#   _emit_tspl_ops  → TSPL 直发
#   LabelEditor._paint_ops → 屏幕预览
#
# 坐标有两套表示，是故意的：
#   mm  —— 正文、装饰线天然以毫米排版，渲染时按目标 dpi 换算（mm2dot）
#   dot —— 表格、底部区块本来就以设计 dpi 的点整数算好，直接用，避免两次取整漂 1 点
# 每个 op 自带模式，渲染器只按模式走同一套公式，不做二次推导。

def _op_rect_center(w_mm, y_mm, h_mm):
    """居中的水平实心条（顶部装饰线）：x 由渲染器按各自 dpi 取整数中点算"""
    return {"kind": "rect", "mode": "center",
            "w_mm": w_mm, "y_mm": y_mm, "h_mm": h_mm}


def _op_rect_abs(x, y, w, h, base_mm=0.0):
    """绝对位置实心矩形（表格框线）：x/y/w/h 是 dot，y 相对 base_mm 基准线"""
    return {"kind": "rect", "mode": "abs", "x": x, "y": y, "w": w, "h": h,
            "base_mm": base_mm}


def _op_text_left(x_mm, y_mm, text, font):
    return {"kind": "text", "mode": "left", "x_mm": x_mm, "y_mm": y_mm,
            "text": text, "font": int(font)}


def _op_text_mid(base_mm, x, y, text, font):
    """水平居中、顶端对齐（底部洗涤说明），x/y 是区块内 dot"""
    return {"kind": "text", "mode": "mid", "base_mm": base_mm,
            "x": x, "y": y, "text": text, "font": int(font)}


def _op_text_cell(base_mm, x, y, text, font, cell_x, cell_w):
    """单元格：按字形墨迹在格子里居中，x 是预先算好的左起点（TSPL 用）"""
    return {"kind": "text", "mode": "cell", "base_mm": base_mm,
            "x": x, "y": y, "text": text, "font": int(font),
            "cell_x": cell_x, "cell_w": cell_w}


def _op_image(base_mm, x, y, w, h, path, th):
    return {"kind": "image", "base_mm": base_mm, "x": x, "y": y,
            "w": w, "h": h, "path": path, "th": th}


def build_ops(proj: dict):
    """把整个标签摊成一份绘制清单（顺序即叠放顺序，先画的在下）"""
    proj = normalize(proj)
    # 字体在这里统一生效：三条渲染路径（预览 / 位图 / TSPL）都从 build_ops 取清单，
    # 在这里设一次，就保证了「量宽度用的字体」和「画字用的字体」是同一个。
    set_font(proj.get("font_file"))
    paper = proj["paper"]
    dpi = int(paper.get("dpi", 203))
    w_mm = float(paper.get("w", 35))
    ops = []

    # ① 顶部装饰线
    tr = proj.get("top_rule") or {}
    if tr.get("on", True):
        rw = float(tr.get("w_mm", 5.0))
        rh = float(tr.get("h_mm", 0.4))
        ops.append(_op_rect_center(rw, float(tr.get("y_mm", 2.0)), rh))

    # ② 正文（品牌 / 号型 / 成分）
    fd = int(proj.get("font_dots", 24))
    _body = layout(proj)
    for (lx, ly, s) in _body:
        ops.append(_op_text_left(lx, ly, s, fd))

    # ②.2 生产日期：正常情况由底部洗涤块的最后一行负责（见 _bottom_ops_impl）。
    # 只有「洗涤块被关掉」时才走这里 —— 否则关掉洗涤说明，日期会跟着一起消失。
    _dp = date_place(proj, _body)
    if _dp:
        _dx, _dy, _df = _dp
        ops.append(_op_text_left(_dx, _dy, date_text(proj), _df))

    # ②.5 分隔线：画在「上文内容底 + gap」处，所以成分多一行它就自动跟着下移
    dv = proj.get("divider") or {}
    if dv.get("on"):
        d_dpi = float(paper.get("dpi", 203)) or 203.0
        dy_mm = content_bottom_mm(proj) + max(0.0, float(dv.get("gap_mm", 1.25)))
        dw = mm2dot(max(0.5, float(dv.get("w_mm", 20.0))), int(d_dpi))
        dh = max(1, mm2dot(max(0.05, float(dv.get("h_mm", 0.2))), int(d_dpi)))
        ops.append(_op_rect_abs(mm2dot(max(0.0, (w_mm - dw * 25.4 / d_dpi) / 2.0),
                                       int(d_dpi)),
                                mm2dot(dy_mm, int(d_dpi)), dw, dh, 0.0))

    # ③ 表格
    if table_on(proj):
        t_base = max(0.0, table_top_mm(proj))
        t_ops, _th = table_ops(proj)
        for op in t_ops:
            if op["kind"] == "bar":
                ops.append(_op_rect_abs(op["x"], op["y"], op["w"], op["h"],
                                        t_base))
            else:
                ops.append(_op_text_cell(t_base, op["x"], op["y"], op["text"],
                                         op["font"], op.get("cell_x", op["x"]),
                                         op.get("cell_w", 0)))

    # ④ 底部洗涤说明区块
    b = proj.get("bottom") or {}
    if b.get("on", True):
        b_ops, _bh = bottom_ops(proj)
        if b_ops:
            b_base = max(0.0, bottom_top_mm(proj))
            for op in b_ops:
                if op["kind"] == "text":
                    ops.append(_op_text_mid(b_base, op["x"], op["y"],
                                            op["text"], op["font"]))
                else:
                    ops.append(_op_image(b_base, op["x"], op["y"], op["w"],
                                         op["h"], op["path"],
                                         op.get("th", 170)))
    return ops


def _draw_pil_ops(ops, img, draw, dpi, s, font_of):
    """把绘制清单画到 PIL 图上。s = 目标 dpi / 设计 dpi"""
    W = img.size[0]
    for op in ops:
        k = op["kind"]
        if k == "rect":
            if op["mode"] == "center":
                rw = max(1, mm2dot(op["w_mm"], dpi))
                rh = max(1, mm2dot(op["h_mm"], dpi))
                rx = max(0, (W - rw) // 2)
                ry = mm2dot(op["y_mm"], dpi)
                draw.rectangle([rx, ry, rx + rw - 1, ry + rh - 1], fill=0)
            else:
                ox = int(round(op["x"] * s))
                oy = (mm2dot(op.get("base_mm", 0.0), dpi)
                      + int(round(op["y"] * s)))
                bw = max(1, int(round(op["w"] * s)))
                bh = max(1, int(round(op["h"] * s)))
                draw.rectangle([ox, oy, ox + bw - 1, oy + bh - 1], fill=0)
        elif k == "text":
            f = op["font"]
            ft = font_of(f * s)
            if op["mode"] == "left":
                draw.text((mm2dot(op["x_mm"], dpi), mm2dot(op["y_mm"], dpi)),
                          op["text"], font=ft, fill=0)
            elif op["mode"] == "mid":
                ox = int(round(op["x"] * s))
                oy = (mm2dot(op["base_mm"], dpi)
                      + int(round(op["y"] * s)))
                cx = ox + int(round(txt_w_dots(op["text"], f) * s / 2.0))
                draw.text((cx, oy), op["text"], font=ft, fill=0, anchor="mt")
            else:                                  # cell
                # 按「字形墨迹」居中：PIL 的 m 锚点按字体行框算，
                # 中英混排时字形实际位置会偏，所以直接拿 getbbox 的墨迹范围对齐
                try:
                    bb = ft.getbbox(op["text"])
                except Exception:
                    bb = (0, 0, 1, f * s)
                cell_cx = (op["cell_x"] + op["cell_w"] / 2.0) * s
                cy_target = (mm2dot(op["base_mm"], dpi)
                             + int(round(op["y"] * s)) + f * s / 2.0)
                # getbbox 的坐标相对「行顶」（anchor=la），据此反推绘制原点
                pen_x = cell_cx - (bb[0] + bb[2]) / 2.0
                pen_y = cy_target - (bb[1] + bb[3]) / 2.0
                draw.text((pen_x, pen_y), op["text"], font=ft, fill=0,
                          anchor="la")
        else:                                      # image
            ox = int(round(op["x"] * s))
            oy = mm2dot(op["base_mm"], dpi) + int(round(op["y"] * s))
            wd = max(1, int(round(op["w"] * s)))
            ic, _iw, _ih = prep_icon(op["path"], wd, op["th"])
            if ic is not None:
                img.paste(ic.convert("L"), (ox, oy))


def _emit_tspl_ops(ops, proj):
    """把绘制清单翻成 TSPL 指令。
    标准字号走打印机内置字库（TEXT），非标准字号转位图（BITMAP）——
    后者保证「任意字号」在直发通道上也精确，代价是数据量变大。"""
    paper = proj["paper"]
    dpi = int(paper.get("dpi", 203))
    w_dot = mm2dot(float(paper.get("w", 35)), dpi)
    out = []
    for op in ops:
        k = op["kind"]
        if k == "rect":
            if op["mode"] == "center":
                rw = max(1, mm2dot(op["w_mm"], dpi))
                rh = max(1, mm2dot(op["h_mm"], dpi))
                rx = max(0, (w_dot - rw) // 2)
                ry = mm2dot(op["y_mm"], dpi)
                out.append("BAR %d,%d,%d,%d" % (rx, ry, rw, rh))
            else:
                out.append("BAR %d,%d,%d,%d"
                           % (op["x"],
                              mm2dot(op.get("base_mm", 0.0), dpi) + op["y"],
                              op["w"], op["h"]))
            continue

        if k == "image":
            img, _wd, hd = prep_icon(op["path"], op["w"], op["th"])
            if img is None:
                continue
            hexdata, wb = bitmap_payload(img)
            if hexdata:
                out.append("BITMAP %d,%d,%d,%d,0,%s"
                           % (op["x"], mm2dot(op["base_mm"], dpi) + op["y"],
                              wb, hd, hexdata))
            continue

        # ---- 文字 ----
        f = op["font"]
        fnt = FONT_ALIAS.get(f)
        base = mm2dot(op["base_mm"], dpi) if "base_mm" in op else 0
        if op["mode"] == "left":
            x0, y0 = mm2dot(op["x_mm"], dpi), mm2dot(op["y_mm"], dpi)
        else:
            x0, y0 = op["x"], base + op["y"]

        if fnt:                                    # 内置字库直出
            out.append('TEXT %d,%d,"%s",0,1,1,"%s"'
                       % (x0, y0, fnt, op["text"].replace('"', "'")))
            continue

        if op["mode"] == "cell":                   # 位图：按格子居中
            img = text_bitmap_image(op["text"], f)
            if img is None:
                continue
            hexdata, wb = bitmap_payload(img)
            if not hexdata:
                continue
            bw_px, bh_px = img.size
            bx = op["cell_x"] + max(0, (op["cell_w"] - bw_px) // 2)
            by = base + op["y"] + max(0, (f - bh_px) // 2)
            out.append("BITMAP %d,%d,%d,%d,0,%s" % (bx, by, wb, bh_px, hexdata))
        else:                                      # 位图：按原位对齐
            img, dx, dy = text_bitmap_with_pos(op["text"], f)
            if img is None:
                continue
            hexdata, wb = bitmap_payload(img)
            if not hexdata:
                continue
            out.append("BITMAP %d,%d,%d,%d,0,%s"
                       % (x0 + dx, y0 + dy, wb, img.size[1], hexdata))
    return out




def purge_printer(printer_name):
    """
    尽力清空打印机驱动队列（终止后防止残留指令继续打）。
    需要管理员权限才可能成功，失败就静默放弃 —— 只是兜底动作。
    """
    try:
        spool = ctypes.WinDLL("winspool.drv", use_last_error=True)
        spool.OpenPrinterW.argtypes = [wintypes.LPWSTR, ctypes.POINTER(wintypes.HANDLE),
                                       ctypes.c_void_p]
        spool.OpenPrinterW.restype = wintypes.BOOL
        spool.SetPrinterW.argtypes = [wintypes.HANDLE, wintypes.DWORD, ctypes.c_void_p,
                                      wintypes.DWORD]
        spool.SetPrinterW.restype = wintypes.BOOL
        spool.ClosePrinter.argtypes = [wintypes.HANDLE]
        h = wintypes.HANDLE()
        if spool.OpenPrinterW(printer_name, ctypes.byref(h), None):
            try:
                spool.SetPrinterW(h, 0, None, 3)      # PRINTER_CONTROL_PURGE
            finally:
                spool.ClosePrinter(h)
        return True
    except Exception:
        return False




def _require_printer(proj):
    """取目标打印机名；没有就直接报错，绝不让 Windows 拿「系统默认打印机」顶上

    以前这里写的是 `proj.get("printer") or DEFAULT_PROJECT["printer"]` ——
    一句 `CreateDCW("")` 的语义是「用系统默认打印机」，也就是你压根没选过，
    它也会自己找一台打出去，还可能是台 A4 办公机而不是标签机。
    """
    name = str(proj.get("printer") or "").strip()
    if not name:
        raise ValueError("没选打印机（点「切换…」选一台）")
    return name




# ---------------- 图片模式：整页渲染成黑白图，走 Windows 驱动 ----------------
# 为什么需要这条通道：TSPL 直发要求「固件懂 TSPL」+「有中文 TSSxx.BF2 字库」，
# 缺一个就会打出黑方块。图片模式不依赖打印机任何字库和指令集，任何打印机都能打。


def win_fonts_dir():
    """Windows 字体目录。

    **不能写死 C:\\Windows\\Fonts** —— 系统装在 D 盘、或者是精简版/英文版
    Windows（字体在别处或压根没有），写死就会挑不到字体，
    而挑不到字体的后果是**整张标签一个字都不画**（不是变方块，是空白），
    并且完全不报错。这条路径必须跟着环境走。"""
    w = (os.environ.get("WINDIR") or os.environ.get("SystemRoot")
         or os.environ.get("windir") or r"C:\Windows")
    return os.path.join(w, "Fonts")


# 自动模式下的候选，按「常见程度」排：雅黑 → 雅黑粗 → 黑体 → 宋体 → 等线/楷体/幼圆
_TTF_NAMES = ("msyh.ttc", "msyhbd.ttc", "simhei.ttf", "simsun.ttc",
              "Deng.ttf", "simkai.ttf", "simyou.ttf", "msyhl.ttc")
TTF_CANDIDATES = [os.path.join(win_fonts_dir(), n) for n in _TTF_NAMES]

# 扫字体目录的结果缓存（见 scan_cjk_font）
_FONT_SCAN = {"done": False, "path": None}


def scan_cjk_font():
    """候选名一个都不在时的兜底：扫整个字体目录，挑第一个真有中文字形的。

    什么时候会走到这：英文版 Windows、精简版系统、字体被改名（Noto Sans SC 之类）。
    这时常规候选全部落空，如果不兜底就是「标签上一个字都没有」。
    扫描要真加载几十上百个字体，慢，所以结果缓存；正常情况下永远不走这里。"""
    if _FONT_SCAN["done"]:
        return _FONT_SCAN["path"]
    _FONT_SCAN["done"] = True
    d = win_fonts_dir()
    try:
        names = sorted(os.listdir(d))
    except OSError:
        return None
    for n in names:
        if not n.lower().endswith((".ttf", ".ttc")):
            continue
        p = os.path.join(d, n)
        try:
            if font_cjk_ok(p):
                _FONT_SCAN["path"] = p
                return p
        except Exception:
            continue
    return None


def has_cjk_font():
    """这台电脑上到底有没有能打中文的字体 —— 启动自检用"""
    return bool(pick_ttf())


def font_expand_path(s):
    """把存起来的字体路径展开成真实路径。

    存的时候会把 Windows 字体目录缩成 $WIN\\Fonts\\xxx.ttf ——
    这样同一份文件拷到另一台电脑（Windows 装在别的盘）还能找到字体。"""
    s = str(s or "").strip()
    if not s:
        return ""
    if s.startswith("$WIN"):
        win = os.environ.get("WINDIR", r"C:\Windows")
        return os.path.join(win, s[4:].lstrip("\\/"))
    return s


def font_store_path(p):
    """反过来的方向：绝对路径 → 存储形式（系统字体目录用 $WIN 代替）"""
    p = str(p or "").strip()
    if not p:
        return ""
    p = os.path.abspath(p)
    win = os.environ.get("WINDIR", r"C:\Windows")
    base = os.path.join(win, "Fonts")
    try:
        if os.path.normcase(os.path.dirname(p)) == os.path.normcase(base):
            return "$WIN\\Fonts\\" + os.path.basename(p)
    except (ValueError, OSError):
        pass
    return p


# 当前渲染用的字体。None = 按 TTF_CANDIDATES 自动挑。
# 为什么用「全局当前」而不是给每个函数加参数：字宽函数 ttf_w_dots() 被排版、
# 换行、居中等几十处调用，全是无参的。与其改几十个签名，不如每次渲染前设一次。
_FONT_CUR = {"raw": None, "path": None}


def set_font(path):
    """切换渲染字体，返回实际生效的字体路径（None = 一个都用不了）。

    必须清字宽缓存：缓存是按「字号」当 key 的，换字体后同一字号的宽度全变了，
    不清就会用旧字体的宽度去排版 —— 画出来和算出来的对不上，字会挤或出格。

    同一个值重复调用直接返回：layout() / bottom_ops() 里也会调它（见下），
    重绘一次要进来好几次，每次都 expand + os.path.exists 没必要。"""
    raw = str(path or "")
    if raw == _FONT_CUR["raw"]:
        return pick_ttf()
    real = font_expand_path(raw)
    real = real if (real and os.path.exists(real)) else None
    _FONT_CUR["raw"] = raw
    if real != _FONT_CUR["path"]:
        _FONT_CUR["path"] = real
        _TTF_W_CACHE.clear()
    return pick_ttf()


def current_font():
    """当前生效的字体路径（自动模式返回实际挑中的那个）"""
    return pick_ttf()


# 界面「字体」下拉的候选。路径写成 $WIN 形式，换电脑（Windows 装在别的盘）也认。
# 只列常见中文字体：列全系统几百个字体对选标签字体没有帮助，反而更难找。
FONT_CHOICES = [
    ("自动（雅黑 / 宋体 / 黑体）", ""),
    ("微软雅黑", r"$WIN\Fonts\msyh.ttc"),
    ("黑体", r"$WIN\Fonts\simhei.ttf"),
    ("宋体", r"$WIN\Fonts\simsun.ttc"),
    ("等线", r"$WIN\Fonts\Deng.ttf"),
    ("楷体", r"$WIN\Fonts\simkai.ttf"),
    ("幼圆", r"$WIN\Fonts\simyou.ttf"),
]


def font_available(path):
    """这个字体在这台电脑上到底有没有 —— 分发到别的机器后用来提示"""
    p = font_expand_path(path)
    return bool(p) and os.path.exists(p)


def font_cjk_ok(path):
    """这个字体里到底有没有中文字形。

    做法：拿一个汉字和一个私用区码位（U+E123，正常字体都没有）分别渲染一次，
    两张点阵**完全一样**就说明汉字也落到了 .notdef —— 即这个字体没有中文，
    选中它中文全会变成方块/空白。纯英文字体（如 Arial）就是这种情况。"""
    p = font_expand_path(path)
    if not p or not os.path.exists(p):
        return False
    try:
        from PIL import ImageFont, Image
        f = ImageFont.truetype(p, 24)
        def mask(ch):
            im = Image.new("L", (32, 32), 0)
            from PIL import ImageDraw
            ImageDraw.Draw(im).text((2, 2), ch, font=f, fill=255)
            return im.tobytes()
        return mask("国") != mask("\ue123")
    except Exception:
        return False


def font_ok(path):
    """不只是「文件存在」，还得真能被解析。

    用户可能选到一个 .ttf 后缀的坏文件、或者把 ttc 里的某一片当字体用，
    真到出图时才炸就麻烦了 —— 选的时候先试加载一次。"""
    p = font_expand_path(path)
    if not p or not os.path.exists(p):
        return False
    try:
        from PIL import ImageFont
        ImageFont.truetype(p, 20)
        return True
    except Exception:
        return False


def pick_ttf():
    cur = _FONT_CUR["path"]
    if cur and os.path.exists(cur):
        return cur
    for p in TTF_CANDIDATES:
        if os.path.exists(p):
            return p
    return scan_cjk_font()      # 常规候选全落空：英文版/精简版系统的兜底


def render_label(proj: dict, dpi=None):
    """
    把整张标签画成一张黑白图（L 模式，只有 0/255）。
    坐标完全复用 build_ops()，只是换算到目标 dpi，所以和 TSPL 走同一套版式。
    """
    from PIL import Image, ImageDraw, ImageFont
    proj = normalize(proj)
    set_font(proj.get("font_file"))
    paper = proj["paper"]
    base_dpi = float(paper.get("dpi", 203))
    dpi = int(dpi or base_dpi)
    s = dpi / base_dpi                       # 目标分辨率相对设计分辨率的倍率
    w_mm = float(paper.get("w", 35))
    h_mm = float(paper.get("h", 75))
    W, H = mm2dot(w_mm, dpi), mm2dot(h_mm, dpi)
    img = Image.new("L", (W, H), 255)
    d = ImageDraw.Draw(img)

    ttf = pick_ttf()
    fcache = {}

    def font(dots):
        key = max(6, int(round(dots)))
        if key not in fcache:
            fcache[key] = (ImageFont.truetype(ttf, key) if ttf
                           else ImageFont.load_default())
        return fcache[key]

    # ①②③④ 统一绘制清单 —— 与 TSPL 直发、画布预览共用同一份，
    # 三处版式永远一致，改一处三处同时生效。
    _draw_pil_ops(build_ops(proj), img, d, dpi, s, font)

    # ⑤ 打印方向 + 打印位移
    #    先转方向、再施加位移：位移的语义是「成品在纸上的落点」，不该跟着旋转变成竖向。
    #    负偏移时内容越界被裁掉，正是纸上真实发生的事。预览与图片模式共用这一张图。
    _ang = int(proj.get("direction", 0) or 0) % 360
    if _ang == 90:
        img = img.rotate(-90, expand=True)       # 顺时针
    elif _ang == 180:
        img = img.rotate(180, expand=False)
    elif _ang == 270:
        img = img.rotate(90, expand=True)

    ox = mm2dot(float(proj.get("off_x_mm", 0.0)), dpi)
    oy = mm2dot(float(proj.get("off_y_mm", 0.0)), dpi)
    if ox or oy:
        cw, ch = img.size
        shifted = Image.new("L", (cw, ch), 255)
        shifted.paste(img, (ox, oy))
        img = shifted

    # TrueType 渲染带灰阶抗锯齿，直接发给打印机会被驱动半色调化（又变网点）。
    # 最后统一阈值化：纯黑纯白，笔画还略加粗，热敏纸上更实。
    img = img.point(lambda x: 0 if x < 170 else 255)
    return img


class _DOCINFO(ctypes.Structure):
    _fields_ = [("cbSize", wintypes.DWORD), ("lpszDocName", wintypes.LPWSTR),
                ("lpszOutput", wintypes.LPWSTR), ("lpszDatatype", wintypes.LPWSTR),
                ("fwType", wintypes.DWORD)]


class _BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [("biSize", wintypes.DWORD), ("biWidth", ctypes.c_long),
                ("biHeight", ctypes.c_long), ("biPlanes", wintypes.WORD),
                ("biBitCount", wintypes.WORD), ("biCompression", wintypes.DWORD),
                ("biSizeImage", wintypes.DWORD), ("biXPelsPerMeter", ctypes.c_long),
                ("biYPelsPerMeter", ctypes.c_long), ("biClrUsed", wintypes.DWORD),
                ("biClrImportant", wintypes.DWORD)]


class _BITMAPINFO(ctypes.Structure):
    _fields_ = [("bmiHeader", _BITMAPINFOHEADER), ("bmiColors", wintypes.DWORD * 3)]


def _dib24(img):
    """PIL 图 → 24bit 自下而上 DIB（GDI 要求 BGR、底朝上、行按 4 字节对齐）"""
    from PIL import Image
    g = img.convert("RGB").transpose(Image.FLIP_TOP_BOTTOM)
    w, h = g.size
    raw = g.tobytes()
    bgr = bytearray(raw)
    bgr[0::3], bgr[2::3] = raw[2::3], raw[0::3]        # RGB → BGR
    stride = ((w * 3 + 3) // 4) * 4
    if stride == w * 3:
        return bytes(bgr)
    out = bytearray(stride * h)
    rowlen = w * 3
    for y in range(h):
        out[y * stride:y * stride + rowlen] = bgr[y * rowlen:(y + 1) * rowlen]
    return bytes(out)


# ---------------- 按文件尺寸设驱动纸型（连续纸 / 任意长度） ----------------
# 为什么非做不可：图片模式建打印画布时，如果 DEVMODE 传 None，驱动就用它自己
# **上次选中的纸型**。而标签机驱动的纸型是一张固定清单 —— 实测这台 HPRT N31
# 报上来 59 项，35mm 宽只有 20/30/40/50/60/65/70，**没有 75**。
# 清单里没有 35×75，驱动就顶成最接近的一项，于是不管打开哪份文件，
# 打出来都是同一个长度。文件里存的 w/h 只进入了"画多大一张图"，
# 从来没进入"这张纸该多长"。
#
# 这台驱动的第 1 项就是 Custom(76×130)，id=256 —— 有自定义纸槽，长宽都能改。
# 所以正确做法是：每次打印前，把驱动当前 DEVMODE 复制一份，改成
# dmPaperSize=DMPAPER_USER + dmPaperWidth/Length=文件尺寸，再用它建画布。

DM_OUT_BUFFER = 0x2
DM_PAPERSIZE = 0x2
DM_PAPERLENGTH = 0x4
DM_PAPERWIDTH = 0x8
DMPAPER_USER = 256


class _DEVMODEW(ctypes.Structure):
    """DEVMODEW（打印机用那一支 union）。

    只用到纸型相关的几个字段，但整段必须按真实布局声明 ——
    DocumentProperties 填进来的是驱动原始字节，少一个字段，
    后面的 dmPaperWidth / dmFormName 就全部错位，改出来的纸型是垃圾值。
    """
    _fields_ = [
        ("dmDeviceName", wintypes.WCHAR * 32),
        ("dmSpecVersion", wintypes.WORD),
        ("dmDriverVersion", wintypes.WORD),
        ("dmSize", wintypes.WORD),
        ("dmDriverExtra", wintypes.WORD),
        ("dmFields", wintypes.DWORD),
        # union 第一支：打印机纸型
        ("dmOrientation", ctypes.c_short),
        ("dmPaperSize", ctypes.c_short),
        ("dmPaperLength", ctypes.c_short),
        ("dmPaperWidth", ctypes.c_short),
        ("dmScale", ctypes.c_short),
        ("dmCopies", ctypes.c_short),
        ("dmDefaultSource", ctypes.c_short),
        ("dmPrintQuality", ctypes.c_short),
        # union 结束
        ("dmColor", ctypes.c_short),
        ("dmDuplex", ctypes.c_short),
        ("dmYResolution", ctypes.c_short),
        ("dmTTOption", ctypes.c_short),
        ("dmCollate", ctypes.c_short),
        ("dmFormName", wintypes.WCHAR * 32),
        ("dmLogPixels", wintypes.WORD),
        ("dmBitsPerPel", wintypes.DWORD),
        ("dmPelsWidth", wintypes.DWORD),
        ("dmPelsHeight", wintypes.DWORD),
        ("dmNup", wintypes.DWORD),
        ("dmDisplayFrequency", wintypes.DWORD),
        ("dmICMMethod", wintypes.DWORD),
        ("dmICMIntent", wintypes.DWORD),
        ("dmMediaType", wintypes.DWORD),
        ("dmDitherType", wintypes.DWORD),
        ("dmReserved1", wintypes.DWORD),
        ("dmReserved2", wintypes.DWORD),
        ("dmPanningWidth", wintypes.DWORD),
        ("dmPanningHeight", wintypes.DWORD),
    ]


def _devmode_raw(printer):
    """向驱动要一份当前 DEVMODE 原始字节（含它自己的私有区 dmDriverExtra）"""
    spool = ctypes.WinDLL("winspool.drv", use_last_error=True)
    spool.DocumentPropertiesW.restype = ctypes.c_long
    spool.DocumentPropertiesW.argtypes = [wintypes.HWND, wintypes.HANDLE,
                                          wintypes.LPCWSTR, ctypes.c_void_p,
                                          ctypes.c_void_p, wintypes.DWORD]
    n = spool.DocumentPropertiesW(None, None, printer, None, None, 0)
    if n <= ctypes.sizeof(_DEVMODEW):
        return None                     # 驱动不给 DEVMODE（虚拟打印机常见）
    buf = ctypes.create_string_buffer(n)
    if spool.DocumentPropertiesW(None, None, printer, buf, None,
                                 DM_OUT_BUFFER) < 0:
        return None
    return buf


def devmode_for_paper(printer, w_mm, h_mm):
    """复制驱动当前 DEVMODE，把纸型改成自定义 w×h；拿不到就返回 None"""
    buf = _devmode_raw(printer)
    if buf is None:
        return None
    try:
        dm = _DEVMODEW.from_buffer(buf)
        dm.dmFields |= (DM_PAPERSIZE | DM_PAPERWIDTH | DM_PAPERLENGTH)
        dm.dmPaperSize = DMPAPER_USER
        dm.dmPaperWidth = max(1, int(round(float(w_mm) * 10)))    # 单位 0.1mm
        dm.dmPaperLength = max(1, int(round(float(h_mm) * 10)))
        return buf
    except Exception:
        return None


def _dc_paper_mm(hdc, dx, dy):
    """问画布：你实际的纸有多大（mm）。驱动不认自定义纸时这里会露出原形"""
    gdi = ctypes.WinDLL("gdi32", use_last_error=True)
    gdi.GetDeviceCaps.argtypes = [wintypes.HDC, ctypes.c_int]
    gdi.GetDeviceCaps.restype = ctypes.c_int
    pw = gdi.GetDeviceCaps(hdc, 110)          # PHYSICALWIDTH
    ph = gdi.GetDeviceCaps(hdc, 111)          # PHYSICALHEIGHT
    if pw <= 0 or ph <= 0:
        return 0.0, 0.0
    return pw / float(dx) * 25.4, ph / float(dy) * 25.4


def _check_paper_with_driver(printer, w_mm, h_mm, cache=None):
    """预检：驱动到底接不接受「自定义 w×h」这张纸（不打印，只建一次画布）

    返回 dict：{"ok": bool, "want": (w,h), "got": (w,h), "why": str}
    ok=False 表示这台驱动的图片模式不认自定义纸长 —— 必须提前告警，
    否则用户会打完一整批才发现长度全错。
    """
    key = None
    if cache is not None:
        key = (printer, round(float(w_mm), 1), round(float(h_mm), 1))
        if key in cache:
            return cache[key]

    res = {"ok": False, "want": (float(w_mm), float(h_mm)),
           "got": (0.0, 0.0), "why": ""}
    dm = devmode_for_paper(printer, w_mm, h_mm)
    if dm is None:
        res["why"] = "这台驱动不提供纸型信息（虚拟打印机/共享机常见），无法按文件尺寸设纸"
        if cache is not None:
            cache[key] = res
        return res

    gdi = ctypes.WinDLL("gdi32", use_last_error=True)
    gdi.CreateDCW.restype = wintypes.HDC
    gdi.CreateDCW.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR,
                              wintypes.LPCWSTR, ctypes.c_void_p]
    gdi.DeleteDC.argtypes = [wintypes.HDC]
    gdi.GetDeviceCaps.argtypes = [wintypes.HDC, ctypes.c_int]
    hdc = gdi.CreateDCW("WINSPOOL", printer, None, ctypes.byref(dm))
    if not hdc:
        res["why"] = "打不开打印机画布"
        if cache is not None:
            cache[key] = res
        return res
    try:
        dx = gdi.GetDeviceCaps(hdc, 88) or 203
        dy = gdi.GetDeviceCaps(hdc, 90) or dx
        got_w, got_h = _dc_paper_mm(hdc, dx, dy)
        res["got"] = (round(got_w, 1), round(got_h, 1))
        ok_w = abs(got_w - float(w_mm)) <= 0.8
        ok_h = abs(got_h - float(h_mm)) <= 0.8
        res["ok"] = bool(ok_w and ok_h)
        if not res["ok"]:
            res["why"] = ("驱动没接受自定义纸长：要 %.1f×%.1f，实际给的是 %.1f×%.1f"
                          % (w_mm, h_mm, got_w, got_h))
    finally:
        gdi.DeleteDC(hdc)
    if cache is not None:
        cache[key] = res
    return res


_LAST_PAPER_REPORT = {}


def last_paper_report():
    """上一次图片模式打印时，驱动实际接受的纸尺寸（给界面报错用）"""
    return dict(_LAST_PAPER_REPORT)




def _open_printer_dc(printer, devmode=None):
    """建打印机 DC 并问出真实 dpi；打不开就抛，让上层提示得清楚

    devmode 不为空时，画布就按这份 DEVMODE 里的纸型建 ——
    这是「每个文件按自己尺寸打」的关键，少了它驱动只会用上次留下的固定纸型。
    """
    gdi = ctypes.WinDLL("gdi32", use_last_error=True)
    gdi.CreateDCW.restype = wintypes.HDC
    gdi.CreateDCW.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR,
                              wintypes.LPCWSTR, ctypes.c_void_p]
    gdi.GetDeviceCaps.argtypes = [wintypes.HDC, ctypes.c_int]
    gdi.DeleteDC.argtypes = [wintypes.HDC]
    hdc = gdi.CreateDCW("WINSPOOL", printer, None,
                        ctypes.byref(devmode) if devmode is not None else None)
    if not hdc:
        raise RuntimeError("打不开打印机「%s」（创建画布失败）" % printer)
    dx = gdi.GetDeviceCaps(hdc, 88) or 203          # LOGPIXELSX
    dy = gdi.GetDeviceCaps(hdc, 90) or dx           # LOGPIXELSY
    return hdc, dx, dy




# ---------------- 自检卡：一张纸判断「驱动 / 纸型 / 通道」哪一层出问题 ----------------
# 为什么需要：打歪、发虚、一边缺角、黑方块，这几种毛病在打出来的成品上长得都差不多，
# 光看字猜不出原因。自检卡把「能自证的东西」全放一页：外框看有没有被裁切、
# 25mm 标准线拿尺子一量就知道驱动有没有缩放、字号阶梯看最小可读字、
# 细线看有没有断针、灰阶看浓度、反白看有没有糊。30 秒定位到层，不用瞎试。
_BAYER4 = ((0, 8, 2, 10), (12, 4, 14, 6), (3, 11, 1, 9), (15, 7, 13, 5))


def _swatch(d, x0, y0, x1, y1, ratio):
    """用 4×4 Bayer 抖动画一块「灰度」—— 热敏只有黑白，靠点密度骗眼睛"""
    thr = ratio * 16
    for yy in range(y0, y1):
        row = _BAYER4[(yy - y0) % 4]
        for xx in range(x0, x1):
            if row[(xx - x0) % 4] < thr:
                d.point((xx, yy), fill=0)


def render_selftest(w_mm=35.0, h_mm=75.0, dpi=203, info=None, *, mode_label="图片模式"):
    """画一张自检卡（L 模式位图）。info: {"printer","paper","dpi","mode","time","app"}"""
    from PIL import Image, ImageDraw, ImageFont
    info = info or {}
    dpm = dpi / 25.4

    def X(mm):
        return int(round(mm * dpm))

    W, H = X(w_mm), X(h_mm)
    img = Image.new("L", (W, H), 255)
    d = ImageDraw.Draw(img)
    ttf = pick_ttf()
    cache = {}

    def font(mm):
        px = max(6, int(round(mm * dpm)))
        if px not in cache:
            cache[px] = ImageFont.truetype(ttf, px) if ttf else ImageFont.load_default()
        return cache[px]

    def text(x_mm, y_mm, s, mm, anchor="la"):
        f = font(mm)
        d.text((X(x_mm), X(y_mm)), s, font=f, fill=0, anchor=anchor)

    def ctext(y_mm, s, mm):
        f = font(mm)
        bb = d.textbbox((0, 0), s, font=f)
        d.text(((W - (bb[2] - bb[0])) // 2 - bb[0], X(y_mm)), s, font=f, fill=0)

    # ① 外框 + 四角 L 标：缺边 / 偏移 / 缩放，一眼就看出来
    pad, lw = X(1.0), max(1, X(0.3))
    d.rectangle([pad, pad, W - 1 - pad, H - 1 - pad], outline=0, width=lw)
    L = X(3.5)
    for cx, cy, sx, sy in ((pad, pad, 1, 1), (W - 1 - pad, pad, -1, 1),
                           (pad, H - 1 - pad, 1, -1), (W - 1 - pad, H - 1 - pad, -1, -1)):
        d.line([cx, cy, cx + sx * L, cy], fill=0, width=lw)
        d.line([cx, cy, cx, cy + sy * L], fill=0, width=lw)

    # ② 标题
    ctext(2.0, "自检卡", 4.4)
    ctext(7.4, f"{mode_label} · {w_mm:g}×{h_mm:g} mm", 1.8)

    # ③ 横刻度尺：每 1mm 短刻度、每 5mm 长刻度 + 数字（量驱动有没有缩放）
    base_y = X(11.6)
    d.line([X(2.0), base_y, X(33.0), base_y], fill=0, width=1)
    for i in range(0, 32):
        x = X(2.0 + i)
        long_ = (i % 5 == 0)
        d.line([x, base_y, x, base_y - X(1.9 if long_ else 1.0)], fill=0, width=1)
        if long_ and i <= 25:                    # 30 贴边会被裁，只标到 25
            d.text((x - X(0.6), base_y + X(0.5)), str(i), font=font(1.5), fill=0)

    # ④ 25mm 标准线：拿尺子量这条线，是判断「驱动缩放 / 纸型不符」最直接的证据
    sy = X(15.4)
    d.line([X(4.0), sy, X(29.0), sy], fill=0, width=max(1, X(0.5)))
    text(4.0, 16.0, "↑ 这条线量出来应该是 25.0 mm", 1.6)

    # ⑤ 环境信息
    d.line([X(2.0), X(18.8), X(33.0), X(18.8)], fill=0, width=1)
    rows = [("打印机", str(info.get("printer") or "—")),
            ("纸型", str(info.get("paper") or "%g×%gmm" % (w_mm, h_mm))),
            ("分辨率", "%d dpi（驱动报的）" % dpi),
            ("时间", str(info.get("time") or ""))]
    y = 19.4
    for k, v in rows:
        text(2.4, y, k, 1.7)
        text(9.0, y, v[:22], 1.7)
        y += 2.5

    # ⑥ 字号阶梯：最小可读字是「这台机器 + 这卷纸」的真实能力
    text(2.4, 29.8, "字号阶梯（右边是字高）", 1.7)
    y = 31.6
    for mm in (3.0, 2.5, 2.0, 1.6, 1.3, 1.0):
        text(2.6, y, "永和 175/92A", mm)
        text(26.0, y + 0.2, "%.1fmm" % mm, 1.5)
        y += 2.9

    # ⑦ 细线组：断针 / 毛刺 / 断线
    text(2.4, 49.2, "细线 / 斜线 / 竖线", 1.5)
    y0 = X(51.4)
    d.line([X(3.0), y0, X(28.0), y0], fill=0, width=1)
    d.line([X(3.0), y0 + X(1.0), X(28.0), y0 + X(1.0)], fill=0, width=max(1, X(0.5)))
    d.line([X(3.0), y0, X(3.0), y0 + X(3.4)], fill=0, width=1)
    d.line([X(4.0), y0, X(6.4), y0 + X(2.4)], fill=0, width=1)
    for i in range(6):                          # 1px 疏密线：断针会断在某一条上
        x = X(8.0 + i * 1.2)
        d.line([x, y0, x, y0 + X(3.2)], fill=0, width=1)

    # ⑧ 灰阶：只有黑白也能看出浓度差（密度不同，眼睛当灰度看）
    y1, y2 = X(56.2), X(60.0)
    labels = (1.0, 0.8, 0.6, 0.4, 0.2, 0.08)
    bw = X(4.4)
    for i, r in enumerate(labels):
        x0 = X(2.6) + i * (bw + X(0.5))
        _swatch(d, x0, y1, x0 + bw, y2, r)
        d.rectangle([x0, y1, x0 + bw, y2], outline=0)
        d.text((x0 + X(0.2), y2 + X(0.3)), "%d%%" % int(r * 100),
               font=font(1.3), fill=0)

    # ⑨ 反白：黑底白字，浓度不够会糊成一片
    bx0, by0 = X(2.6), X(62.6)
    d.rectangle([bx0, by0, X(32.4), by0 + X(4.6)], fill=0)
    d.text(((bx0 + X(32.4)) // 2, by0 + X(2.3)), "反白 清晰吗", font=font(2.4),
           fill=255, anchor="mm")

    # ⑩ 底部判定提示
    text(2.4, 68.0, "边框四边齐全 → 纸型/偏移对", 1.4)
    text(2.4, 69.8, "标准线=25.0mm → 驱动没缩放", 1.4)
    text(2.4, 71.6, "灰阶能分开 6 档 → 浓度够", 1.4)
    return img




def _print_selftest_with_driver(printer, mode="image", w_mm=35.0, h_mm=75.0, gap_mm=2.0,
                   info=None, media="gap"):
    """打一张自检卡；返回 (发出去的字节数, 用了哪条通道)"""
    printer = (printer or "").strip()
    if not printer:
        raise ValueError("没选打印机（点「切换…」选一台）")
    info = dict(info or {})
    info.setdefault("printer", printer)
    info.setdefault("time", time.strftime("%Y-%m-%d %H:%M"))

    if mode == "tspl":
        tspl = build_selftest_tspl(w_mm, h_mm, gap_mm, media=media)
        return raw_print(printer, encode_tspl(tspl), doc_name="自检卡 TSPL"), "TSPL 直发"

    # 自检卡也要按这次要验的尺寸建画布，否则它打出来还是驱动上次那张纸，
    # 就失去了「拿尺子量一下就知道驱动有没有缩放」的意义
    dm = devmode_for_paper(printer, w_mm, h_mm)
    hdc, dx, dy = _open_printer_dc(printer, dm)
    gdi = ctypes.WinDLL("gdi32", use_last_error=True)
    gdi.DeleteDC.argtypes = [wintypes.HDC]
    try:
        info.setdefault("dpi", dx)
        img = render_selftest(w_mm, h_mm, dx, info)
        return _gdi_send(hdc, img, dx, dy, 1, "自检卡 图片模式"), "图片模式"
    finally:
        gdi.DeleteDC(hdc)


# ---------------- 项目文件存读 ----------------
def save_project(proj: dict, path: str):
    data = normalize(proj)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def load_project(path: str) -> dict:
    with open(path, "r", encoding="utf-8-sig") as f:
        return normalize(json.load(f))


# ---------------- 打印队列与 Windows / TSPL 通道 ----------------
def _tspl_bitmap_page(image, w_mm, h_mm, gap_mm, media):
    width, height = image.size
    width_bytes = (width + 7) // 8
    bitmap = image.convert("1", dither=Image.Dither.NONE)
    if width_bytes * 8 != width:
        padded = Image.new("1", (width_bytes * 8, height), 1)
        padded.paste(bitmap, (0, 0))
        bitmap = padded
    gap = "GAP 0 mm,0 mm" if str(media).lower() == "continuous" else f"GAP {gap_mm:g} mm,0 mm"
    return "\r\n".join((
        f"SIZE {w_mm:g} mm,{h_mm:g} mm",
        gap,
        "DIRECTION 1",
        "REFERENCE 0,0",
        "CLS",
        f"BITMAP 0,0,{width_bytes},{height},0,{bitmap.tobytes().hex().upper()}",
        "PRINT 1,1",
        "",
    ))


def _build_tspl_default(proj):
    proj = normalize(proj)
    if int(proj.get("direction") or 0) % 360 in (90, 270):
        raise RuntimeError("TSPL 直发不支持 90° 方向，请把打印方式切到「图片模式」")
    paper = proj["paper"]
    return _tspl_bitmap_page(render_label(proj), float(paper["w"]), float(paper["h"]),
                            float(paper["gap"]), paper["media"])


def encode_tspl(tspl: str) -> bytes:
    payload = bytearray()
    for line in tspl.splitlines():
        if line.startswith("BITMAP "):
            fields = line.split(",", 5)
            bitmap = bytes.fromhex(fields[5])
            if len(bitmap) != int(fields[2]) * int(fields[3]):
                raise ValueError("TSPL 位图长度与指令中的尺寸不一致")
            payload.extend((",".join(fields[:5]) + ",").encode("ascii"))
            payload.extend(bitmap)
        else:
            payload.extend(line.encode("gbk", errors="replace"))
        payload.extend(b"\r\n")
    return bytes(payload)


class PrintCancelled(Exception):
    """打印任务在提交前或提交过程中被终止。"""


def _check_cancelled(cancelled):
    if cancelled is not None and cancelled():
        raise PrintCancelled()


def _print_tspl_copies_default(proj, copies, doc_name=None, cancelled=None, submitted=None):
    _check_cancelled(cancelled)
    printer = _require_printer(proj)
    payload = encode_tspl(build_tspl(proj))
    name = doc_name or proj.get("name", "WaterLabel")
    result = None
    for _ in range(copies):
        _check_cancelled(cancelled)
        result = raw_print(printer, payload, doc_name=name)
        if submitted is not None:
            submitted(1)
    _check_cancelled(cancelled)
    return result


def _gdi_send(hdc, img, dx, dy, copies=1, doc_name="WaterLabel", cancelled=None):
    """送出 1:1 位图，取消或失败时终止当前 Windows 打印作业。"""
    gdi = ctypes.WinDLL("gdi32", use_last_error=True)
    for function in (gdi.StartPage, gdi.EndPage, gdi.EndDoc, gdi.AbortDoc):
        function.argtypes = [wintypes.HDC]
        function.restype = ctypes.c_int
    gdi.StartDocW.argtypes = [wintypes.HDC, ctypes.c_void_p]
    gdi.StartDocW.restype = ctypes.c_int
    gdi.StretchDIBits.argtypes = [wintypes.HDC] + [ctypes.c_int] * 8 + [
        ctypes.c_void_p, ctypes.c_void_p, wintypes.UINT, wintypes.DWORD]
    gdi.StretchDIBits.restype = ctypes.c_int

    raw = _dib24(img)
    w, h = img.size
    bi = _BITMAPINFO()
    hdr = bi.bmiHeader
    hdr.biSize = ctypes.sizeof(_BITMAPINFOHEADER)
    hdr.biWidth, hdr.biHeight = w, h
    hdr.biPlanes, hdr.biBitCount = 1, 24
    hdr.biCompression = 0
    hdr.biSizeImage = len(raw)
    hdr.biXPelsPerMeter = int(dx / 25.4 * 1000)
    hdr.biYPelsPerMeter = int(dy / 25.4 * 1000)
    buf = ctypes.create_string_buffer(raw)
    info = _DOCINFO(ctypes.sizeof(_DOCINFO), doc_name, None, None, 0)

    def require_success(result, operation):
        if result <= 0:
            raise RuntimeError(f"{operation} 失败（错误码 {ctypes.get_last_error()}）")

    _check_cancelled(cancelled)
    require_success(gdi.StartDocW(hdc, ctypes.byref(info)), "StartDoc")
    try:
        for _ in range(max(1, int(copies))):
            _check_cancelled(cancelled)
            require_success(gdi.StartPage(hdc), "StartPage")
            require_success(gdi.StretchDIBits(
                hdc, 0, 0, w, h, 0, 0, w, h, buf, ctypes.byref(bi), 0, 0x00CC0020
            ), "StretchDIBits")
            require_success(gdi.EndPage(hdc), "EndPage")
        _check_cancelled(cancelled)
        require_success(gdi.EndDoc(hdc), "EndDoc")
    except BaseException:
        gdi.AbortDoc(hdc)
        raise
    return len(raw)


def _gdi_submit_pages(printer, proj, w_mm, h_mm, pages, cancelled):
    global _LAST_PAPER_REPORT
    _check_cancelled(cancelled)
    dm = devmode_for_paper(printer, w_mm, h_mm)
    hdc, dx, dy = _open_printer_dc(printer, dm)
    gdi = ctypes.WinDLL("gdi32", use_last_error=True)
    gdi.DeleteDC.argtypes = [wintypes.HDC]
    try:
        got_w, got_h = _dc_paper_mm(hdc, dx, dy)
        _LAST_PAPER_REPORT = {
            "printer": printer,
            "want": (w_mm, h_mm),
            "got": (round(got_w, 2), round(got_h, 2)),
            "ok": abs(got_w - w_mm) <= 0.8 and abs(got_h - h_mm) <= 0.8,
            "devmode": dm is not None,
            "media": proj["paper"]["media"],
        }
        image = render_label(proj, dpi=dx)
        _gdi_send(hdc, image, dx, dy, pages, proj.get("name", "WaterLabel"), cancelled=cancelled)
    finally:
        gdi.DeleteDC(hdc)


def _gdi_try_continuous_page(printer, proj, w_mm, h_mm, copies, cancelled):
    _check_cancelled(cancelled)
    total_mm = h_mm * copies
    dm = devmode_for_paper(printer, w_mm, total_mm)
    if dm is None:
        return False
    try:
        hdc, dx, dy = _open_printer_dc(printer, dm)
    except RuntimeError:
        return False
    gdi = ctypes.WinDLL("gdi32", use_last_error=True)
    gdi.DeleteDC.argtypes = [wintypes.HDC]
    try:
        got_w, got_h = _dc_paper_mm(hdc, dx, dy)
        if abs(got_w - w_mm) > 0.8 or abs(got_h - total_mm) > 0.8:
            return False
        gdi.GetDeviceCaps.argtypes = [wintypes.HDC, ctypes.c_int]
        gdi.GetDeviceCaps.restype = ctypes.c_int
        max_width = gdi.GetDeviceCaps(hdc, 8)
        max_height = gdi.GetDeviceCaps(hdc, 10)
        if max_width <= 0 or max_height <= 0:
            return False

        label = render_label(proj, dpi=dx)
        image = Image.new("L", (label.width, mm2dot(total_mm, dy)), 255)
        for index in range(copies):
            top = mm2dot(index * h_mm, dy)
            height = mm2dot((index + 1) * h_mm, dy) - top
            cell = label if label.height == height else label.resize((label.width, height), Image.Resampling.NEAREST)
            image.paste(cell, (0, top))
        bounds = image.point(lambda pixel: 255 - pixel).getbbox()
        if bounds and (bounds[2] > max_width or bounds[3] > max_height):
            return False
        _gdi_send(hdc, image, dx, dy, 1, proj.get("name", "WaterLabel"), cancelled=cancelled)
        return True
    finally:
        gdi.DeleteDC(hdc)


def _gdi_print_with_driver(proj, copies=1, cancelled=None, submitted=None):
    global _LAST_PAPER_REPORT
    proj = normalize(proj)
    printer = _require_printer(proj)
    paper = proj["paper"]
    w_mm, h_mm = float(paper["w"]), float(paper["h"])
    copies = max(1, int(copies))
    remaining = copies
    documents = combined = 0
    continuous = paper["media"] == "continuous" and copies > 1
    _LAST_PAPER_REPORT = {}

    if continuous and int(proj.get("direction") or 0) % 360 in (0, 180) and h_mm > 0:
        limit = min(remaining, int(3276.7 // h_mm))
        while remaining > 1 and limit > 1:
            count = min(remaining, limit)
            if not _gdi_try_continuous_page(printer, proj, w_mm, h_mm, count, cancelled):
                limit = (count + 1) // 2
                continue
            remaining -= count
            combined += count
            documents += 1
            if submitted is not None:
                submitted(count)
    if remaining:
        _gdi_submit_pages(printer, proj, w_mm, h_mm, remaining, cancelled)
        documents += 1
        if submitted is not None:
            submitted(remaining)
    if continuous and combined < copies - 1:
        _LAST_PAPER_REPORT["feed_note"] = "部分或全部标签在同一作业内逐页打印，分页走纸由驱动控制"
    elif continuous and documents > 1:
        _LAST_PAPER_REPORT["feed_note"] = "纸卷按驱动支持的长度分段打印"
    _check_cancelled(cancelled)


def _render_tspl_selftest(w_mm, h_mm, dpi, info=None):
    return render_selftest(w_mm, h_mm, dpi, info, mode_label="TSPL 直发")


def build_selftest_tspl(w_mm, h_mm, gap_mm=2.0, dpi=203, media="gap", info=None):
    image = _render_tspl_selftest(w_mm, h_mm, dpi, info)
    return _tspl_bitmap_page(image, w_mm, h_mm, gap_mm, media)


def _print_selftest_on_channel(printer, mode="image", w_mm=35.0, h_mm=75.0, gap_mm=2.0,
                               info=None, media="gap"):
    if mode != "tspl":
        return _print_selftest_with_driver(printer, mode, w_mm, h_mm, gap_mm, info, media)
    printer = (printer or "").strip()
    if not printer:
        raise ValueError("没选打印机（点「切换…」选一台）")
    info = dict(info or {})
    info.setdefault("printer", printer)
    info.setdefault("time", time.strftime("%Y-%m-%d %H:%M"))
    dpi = int(str(info.get("dpi") or 203).split()[0])
    tspl = build_selftest_tspl(w_mm, h_mm, gap_mm, dpi, media, info=info)
    return raw_print(printer, encode_tspl(tspl), doc_name="自检卡 TSPL"), "TSPL 直发"


def print_project(proj: dict, copies=1):
    project = normalize(proj)
    if resolve_mode(project) == "image":
        return gdi_print(project, copies)
    return print_tspl_copies(project, max(1, int(copies)))


# ---------------- 打印机适配与新项目默认值 ----------------
def normalize(proj):
    if proj:
        return _normalize_saved_project(proj)

    project = _normalize_saved_project({
        "paper": {"media": "continuous", "dpi": 300},
        "print_mode": "tspl",
        "auto_paper": False,
        "divider": {"on": True},
    })
    # Convert only the built-in dot-based design values for a new label.
    # Saved labels keep their original design DPI and dot values.
    scale = project["paper"]["dpi"] / DEFAULT_PROJECT["paper"]["dpi"]
    for section, fields in (
        (project, ("font_dots", "line_gap", "date_font")),
        (project["bottom"], ("title_font", "sub_font", "body_font", "footer_font", "icon_w_dot")),
        (project["table"], ("caption_font", "border_dot")),
    ):
        for field in fields:
            section[field] = round(section[field] * scale)
    for row in project["table"]["rows"]:
        row["font"] = round(row["font"] * scale)
    return project


def bitmap_printer_profile(printer):
    name = str(printer or "").strip().casefold()
    if not name:
        return None
    identifiers = [name]
    for info in list_printer_info():
        if info["name"].casefold() == name:
            identifiers.append((info.get("driver") or "").strip().casefold())
            break
    for identifier in identifiers:
        if re.search(r"(?<![a-z0-9])gp[\s-]*1834t(?![a-z0-9])", identifier):
            return ("佳博 GP-1834T", 300, 1000)
        if identifier == "hprt n31":
            return ("汉印 N31", 203, 300)
    return None


def _tspl_page(image, width, length, paper, profile, driver_image=False):
    commands = _tspl_bitmap_page(image, width, length, float(paper["gap"]), paper["media"])
    commands = commands.replace(
        "REFERENCE 0,0\r\n", "REFERENCE 0,0\r\nOFFSET 0 mm\r\nSET TEAR OFF\r\n", 1
    )
    if driver_image and profile[1] == 203:
        commands = commands.replace("DIRECTION 1\r\n", "DIRECTION 0,0\r\n", 1)
    return commands


def build_tspl(proj):
    proj = normalize(proj)
    profile = bitmap_printer_profile(proj.get("printer"))
    if not profile:
        return _build_tspl_default(proj)
    width, length = label_size_mm(proj)
    paper = proj["paper"]
    driver_image = resolve_mode(proj) == "image"
    image = render_label(proj, dpi=profile[1])
    return _tspl_page(image, width, length, paper, profile, driver_image)


def _continuous_bitmap(image, length, count, dots_per_mm):
    if count == 1:
        return image
    canvas = Image.new("L", (image.width, round(length * count * dots_per_mm)), 255)
    for index in range(count):
        start = round(length * index * dots_per_mm)
        end = round(length * (index + 1) * dots_per_mm)
        cell = image.crop((0, 0, image.width, end - start)) if image.height > end - start else image
        canvas.paste(cell, (0, start))
    return canvas


def _send_bitmap_copies(proj, copies, doc_name, cancelled, submitted, profile, driver_image=False):
    _check_cancelled(cancelled)
    printer = _require_printer(proj)
    paper = proj["paper"]
    w_mm, h_mm = label_size_mm(proj)
    dpi, max_length = profile[1:]
    tile = render_label(proj, dpi=dpi)
    dots_per_mm = 8 if dpi == 203 else dpi / 25.4
    # One PRINT per segment also works on printers that ignore repeated PRINT commands.
    continuous = paper["media"] == "continuous"
    segment_size = max(1, int(max_length / h_mm)) if continuous else 1
    remaining = int(copies)
    result = None
    name = doc_name or proj.get("name") or "WaterLabel"
    payloads = {}
    while remaining > 0:
        _check_cancelled(cancelled)
        count = min(segment_size, remaining)
        if count not in payloads:
            image = _continuous_bitmap(tile, h_mm, count, dots_per_mm)
            commands = _tspl_page(image, w_mm, h_mm * count, paper, profile, driver_image)
            payloads[count] = encode_tspl(commands)
        _check_cancelled(cancelled)
        result = raw_print(printer, payloads[count], doc_name=name)
        remaining -= count
        if submitted is not None:
            submitted(count)
    _check_cancelled(cancelled)
    return result


def print_tspl_copies(proj, copies=1, doc_name=None, cancelled=None, submitted=None):
    proj = normalize(proj)
    profile = bitmap_printer_profile(proj.get("printer"))
    if not profile:
        return _print_tspl_copies_default(proj, copies, doc_name, cancelled, submitted)
    return _send_bitmap_copies(proj, copies, doc_name, cancelled, submitted, profile)


def gdi_print(proj, copies=1, cancelled=None, submitted=None):
    global _LAST_PAPER_REPORT
    proj = normalize(proj)
    printer = _require_printer(proj)
    profile = bitmap_printer_profile(printer)
    if not profile:
        return _gdi_print_with_driver(proj, copies, cancelled=cancelled, submitted=submitted)

    size = label_size_mm(proj)
    _LAST_PAPER_REPORT = {"printer": printer, "want": size, "got": size, "ok": True,
                          "devmode": False}
    _send_bitmap_copies(proj, max(1, int(copies)), None, cancelled, submitted, profile, driver_image=True)


def check_paper_apply(printer, w_mm, h_mm, cache=None):
    if bitmap_printer_profile(printer):
        size = (float(w_mm), float(h_mm))
        return {"want": size, "got": size, "ok": True, "why": ""}
    return _check_paper_with_driver(printer, w_mm, h_mm, cache=cache)


def print_selftest(printer, mode="image", w_mm=35.0, h_mm=75.0, gap_mm=2.0, info=None, media="gap"):
    profile = bitmap_printer_profile(printer)
    if not profile:
        return _print_selftest_on_channel(printer, mode, w_mm, h_mm, gap_mm, info, media)
    details = dict(info or {})
    details.setdefault("printer", printer)
    details.setdefault("time", time.strftime("%Y-%m-%d %H:%M"))
    dpi = profile[1]
    details["dpi"] = dpi
    image = _render_tspl_selftest(w_mm, h_mm, dpi, details)
    paper = {"gap": gap_mm, "media": media}
    commands = _tspl_page(image, w_mm, h_mm, paper, profile, driver_image=(mode == "image"))
    return raw_print(printer, encode_tspl(commands), doc_name="WaterLabel 自检页"), (
        "图片模式" if mode == "image" else "TSPL 直发"
    )
