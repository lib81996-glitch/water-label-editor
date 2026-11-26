# -*- coding: utf-8 -*-
"""从 DLabel 的 .ddl 标签文件里读出「号型 / 成分 / 纸张尺寸」。

为什么能读：.ddl 装的是 DLabel 自己的 XML（UTF-8 文本），文字全部放在
节点的 sourcetext / value 属性里，多行用真正的换行符连。它不是压缩包、
不是加密二进制，所以不用 OCR，也不用去调 DLabel 的接口。

实测包含历史标签和本工具生成的模板：
    602/602 都能解码、都能通过 XML 解析；
    成分取到 602/602，号型取到 563/602。
    剩下 39 个是模板里本来就没有号型这一行（棒球帽、渔夫帽、无尺码款，
    只有一个「规格：可调节 56cm-60cm」），这种留空比瞎猜一个安全。

历史上有三种写法，少认一种覆盖率就掉到 56%，所以三层都要吃住：
    ① 号型: 175/84A（L）＋ 成分: 38.6%棉（老款主流）
    ② 面料 / 胆料 / 填充物（羽绒服那种多段成分）
    ③ 标签和内容拆成两个节点：「成分：」一个、 「95%棉5%其他」另一个

表格（充绒表）为什么不能靠线条还原：
    599 个历史模板里只有 28 个真画了横竖线（4.7%），而且线跟文字对不上 ——
    R2611羽绒服 竖线 3 根（x=10.7/18.2/25.0），文字却是 4 列。所以格子只能
    从**文字**里切：一个文本框里某一行按「2 个以上空格」切开后每格都像尺码，
    那就是尺码行；正下方 10mm 内、格数相同的另一个文本框就是数值行。
    实测能成对认出 42/599（7.0%，覆盖 13 个款），只认尺码行 11 个，
    剩下 546 个是单尺码老标签 —— 本来就没有表可导，不算漏。
"""

import os
import re
import xml.etree.ElementTree as ET

# 文本可能放在这两个属性里，而且同一段文案两处各存一份，靠去重消掉
TEXT_ATTRS = ("sourcetext", "value")

# 成分侧的标签词。第一组是主流写法，第二组是从 602 个历史模板里统计出来的
# 真实变体（各出现 4~12 次）：内里 / 内胆 / 里布 / 主料 / 配料 / 外层，
# 以及被当成标签用的材质名 麂皮绒 / 毛绒。
COMP_KEYS = ("成分|面料|里料|胆料|填充物|内里|内胆|里布|内衬|主料|配料|外层"
             "|麂皮绒|毛绒")
# 前面最多 10 个字（覆盖「面料基布成分」「大身填充物」「领部位填充物」
# 「针织部分面料」这种长标签），后面可带一个字母或数字（「面料A」「面料1」）。
# 注意这里**不含**号型/规格 —— 那两个是另一件事，混进来会把号型当成分。
LABEL_RE = re.compile(
    r"^\s*([\u4e00-\u9fa5A-Za-z]{0,10}?(?:%s)[A-Za-z0-9]?)\s*[:：]\s*(.*)$" % COMP_KEYS)
# 「任何像标签的行」：2~10 个汉字/字母 + 冒号。用来**截断**续行 ——
# 「成分: 100%棉 / 执行标准: GB/T 2660」这种，别把执行标准吞进成分里。
ANY_LABEL_RE = re.compile(r"^\s*([\u4e00-\u9fa5A-Za-z]{1,10})\s*[:：]")
# 认不出词、但形状对（标签 + 成分值）时用这套兜底提取
LABEL_RE_FALLBACK = re.compile(
    r"^\s*([\u4e00-\u9fa5A-Za-z]{1,10})\s*[:：]\s*(.*)$")
# 号型侧：独立一套，标签只认「号型」「规格」
SIZE_LBL_RE = re.compile(r"^\s*(号\s*型|规\s*格)\s*[:：]\s*(.*)$")
SIZE_RE = SIZE_LBL_RE
PCT_RE = re.compile(r"^\s*[\d.]+\s*%")
INDENT_RE = re.compile(r"^[ \u3000]")
# 从「规格」里取的号型只认短、纯符号的串：175/84A（L）、XL、L码……
# 「可调节 56cm-60cm」这种不是号型，必须挡掉（挡不掉就会被当成成分的标题）
SIZE_OK_RE = re.compile(r"^[\dA-Za-z/／\-（）()码\s]{1,16}$|^(均码|无尺码|通用)$")
FN_SIZE_RE = re.compile(r"(?:^|[-_ ])([0-9A-Za-z]{1,4}码|均码|F码)(?=$|[-_ ])")

# 洗涤说明那几段也带「洗涤」字样，别被当成成分抓进去
STOP = ("洗涤", "注意", "方式", "CARE", "INSTRUCTION")

NO_PREFIX_LABELS = ("成分", "面料成分")      # 这两个去掉前缀，只留百分比

# ---- 表格（尺码行 + 数值行）----
# 一格算不算尺码。故意不从 label_engine.SIZE_ORDER 引 —— 老标签里 XXL/XXXL/F
# 这些写法不在本工具的预设里，但确实是尺码，认出来才能拼回表。
SIZE_CELL_RE = re.compile(
    r"^(?:XS|S|M|L|XL|XXL|2XL|3XL|XXXL|4XL|5XL|6XL|7XL|均码|F)$", re.I)
# 同一件事的不同写法，统一成本工具的规范（XXL = 2XL）。F 不动 ——
# 「F」印在标签上是有意义的，换成「均码」反而是改人家的字。
SIZE_ALIAS = {"XXL": "2XL", "XXXL": "3XL"}
# 列与列之间至少 2 个空格（半角或全角）—— 1 个空格太松，「品 牌」这种
# 手工对齐的字会被切开，2 个以上才基本是刻意的分列
COL_SPLIT_RE = re.compile(r"[\u3000 ]{2,}")


class DdlError(Exception):
    """文件打不开、或不是 DLabel 的 XML 标签文件"""


def _texts(root):
    """按文档顺序收集所有文案节点，相同的只留一份"""
    out, seen = [], set()
    for el in root.iter():
        for k, v in el.attrib.items():
            if k.lower() in TEXT_ATTRS and v and v.strip():
                if v in seen:
                    continue
                seen.add(v)
                out.append(v)
    return out


def _noise(s):
    return (not s) or len(s) > 24 or any(k in s for k in STOP)


def _size(vals, filename):
    """返回 (号型, 来源说明 "正文"/"文件名"/"")。

    分三层，越往后越宽松，最后才是文件名：
      ① 正文里明确的「号型: xxx」
      ② 标签和内容拆成两个节点的（一个节点只写「号型：」）
      ③ 只有「规格: xxx」且长得像号型（「可调节 56cm-60cm」不算）
    """
    for v in vals:                                  # ①
        for line in v.split("\n"):
            m = SIZE_LBL_RE.match(line)
            if m:
                s = m.group(2).strip()
                if not _noise(s):
                    return s, "正文"
    for i, v in enumerate(vals):                    # ②
        if i + 1 >= len(vals):
            break
        if ("号型" in v or "规格" in v) and v.strip().endswith((":", "：")):
            s = vals[i + 1].strip().split("\n")[0].strip()
            if not _noise(s) and SIZE_OK_RE.match(s):
                return s, "正文"
    for v in vals:                                  # ③
        for line in v.split("\n"):
            m = SIZE_LBL_RE.match(line)
            if m:
                s = m.group(2).strip()
                if SIZE_OK_RE.match(s):
                    return s, "正文"
    stem = os.path.splitext(os.path.basename(filename))[0]   # ④ 文件名兜底
    m = FN_SIZE_RE.search(stem)
    if m:
        return m.group(1), "文件名"
    return "", ""


def _comp_like(t):
    """这一行像不像成分内容（有百分比 / 括号备注 / 「100%锦纶」这种）"""
    return bool(PCT_RE.match(t) or t.startswith(("(", "（")) or "%" in t)


def _fabrics(vals):
    """返回成分行列表，面料 / 里料 / 内胆 / 胆料 / 填充物 全部收，一个都不漏。

    三个要点：
      · **跨文本块收集**：老模板常把「面料+里料」放一个节点、「填充物」放另一个节点
        （实测 A159 就是这样），看到第一个块就 return 会把填充物丢掉。
      · 同块内多组标签（面料 / 胆料 / 填充物）也要全部收下。
      · 带标签的多段成分保留前缀（「胆料: 100%聚酯纤维」只剩百分数就分不清是
        哪一层了）；纯「成分:」的只留百分比，界面上那几行本来就是只写百分比。
    """
    got = []
    for i, v in enumerate(vals):
        lines = v.split("\n")
        for j, line in enumerate(lines):
            m = LABEL_RE.match(line)
            if not m:
                # 标签词没见过的（「针织：100%棉」「网格：100%聚酯纤维」这种），
                # 只要冒号后面长得像成分就认 —— 认词表永远补不全，认形状才收得住
                g = ANY_LABEL_RE.match(line)
                if g and _comp_like(line[g.end():]) \
                        and not any(k in g.group(0) for k in STOP):
                    m = LABEL_RE_FALLBACK.match(line)
                if not m:
                    continue
            label = m.group(1).replace(" ", "")
            first = m.group(2).strip()
            if any(k in label for k in STOP):
                continue
            used = j                       # 已经被这条标签吃掉的行下标
            if first:
                got.append(first if label in NO_PREFIX_LABELS
                           else "%s: %s" % (label, first))
            else:
                # 标签后面是空的，常见的两种摆法：
                #   ① 内容就在下一行（「面料：」换行「72%棉」）—— 同段内往下看
                #   ② 整段只有一个光秃秃的「成分：」，内容在下一个文本节点里
                nxt = lines[j + 1].strip() if j + 1 < len(lines) else ""
                more_here = any(x.strip() for x in lines[j + 1:])
                if nxt and not ANY_LABEL_RE.match(nxt):
                    got.append(nxt if label in NO_PREFIX_LABELS
                               else "%s: %s" % (label, nxt))
                    used = j + 1           # 这一行已被标签领走，别再当续行收一遍
                elif not more_here:
                    # 只有「整段就一个空标签」时才往邻居节点找，
                    # 否则容易把别的区块（洗涤方式之类）抓进来
                    for v2 in vals[i + 1:i + 3]:
                        cand = [x.strip() for x in v2.split("\n") if x.strip()]
                        if cand and len(cand) <= 3 and all(_comp_like(x) for x in cand):
                            got.extend(cand)
                            break
            # 紧跟的续行：缩进过的、以百分比开头的、括号备注的，都算这一段成分；
            # 遇到任何「像标签」的行（含执行标准 / 安全类别）就断，别吞别人的东西
            for s in lines[used + 1:]:
                t = s.strip()
                if not t or ANY_LABEL_RE.match(t):
                    break
                if INDENT_RE.match(s) or _comp_like(t):
                    got.append(t)
                else:
                    break
    if got:
        return got
    # 最后一层：整段就是「95%棉5%其他」这种裸百分比，一个标签都没有
    for v in vals:
        hits = [s.strip() for s in v.split("\n")
                if PCT_RE.match(s) and not SIZE_LBL_RE.match(s)]
        if hits:
            return hits
    return []


def _paper(root):
    """纸张尺寸：<paper w="35" h="75" .../>。读不到就返回 None，不猜。"""
    for el in root.iter():
        if el.tag.lower() != "paper":
            continue
        try:
            w = float(el.attrib.get("w", ""))
            h = float(el.attrib.get("h", ""))
        except (TypeError, ValueError):
            return None
        if w > 0 and h > 0:
            return {"w": w, "h": h}
    return None


def _num(el, key):
    try:
        return float(el.attrib.get(key, 0))
    except (TypeError, ValueError):
        return 0.0


def _boxes(root):
    """所有**文本框**（itemtype=5）及其位置，按文档顺序。

    只要文本框：线条（itemtype=1）和图标（10）里没有字，表格的格子只能从
    文字里切出来，拿不到坐标就没法判断「谁在谁下面」。
    坐标单位是 mm，t 往下增大 —— 所以「正下方」= t 更大且差值在 10mm 内。
    """
    out = []
    for el in root.iter("drawobj"):
        if el.attrib.get("itemtype") != "5":
            continue
        txt = ""
        for fx in el.iter("text"):
            t = fx.attrib.get("sourcetext") or fx.attrib.get("value") or ""
            if t.strip():
                txt = t
                break
        if not txt.strip():
            continue
        out.append({"t": _num(el, "t"), "l": _num(el, "l"),
                    "rows": txt.split("\n")})
    return out


def _cols(line):
    """按 2 个以上空格切一行 → 各列文字"""
    return [c for c in COL_SPLIT_RE.split(line.strip()) if c.strip()]


def _size_row(box):
    """这个文本框里有没有一整行都是尺码，有就返回那一行的各格"""
    for line in box["rows"]:
        cs = _cols(line)
        if len(cs) >= 2 and all(SIZE_CELL_RE.match(c) for c in cs):
            return cs
    return None


def _canon(s):
    s = s.strip()
    return SIZE_ALIAS.get(s.upper(), s)


def _size_table(boxes):
    """返回 (尺码列表, 数值列表)。认不出就 ([], [])。

    两个判据缺一不可，宁可漏也不能错 —— 错一个格子就要手工返工整张表：
      · 尺码行：**同一个文本框内**一整行切出来每格都像尺码。不能跨框拼，
        否则「品牌: AOGZ」里的 AOGZ 也会被凑成一格。
      · 数值行：正下方 10mm 内、**格数相同**的另一个文本框，且内容不是尺码。
        格数相同这一条很关键：R2611 的数值行给的是 147g/154g/161g/169g，
        4 格对 4 格，一一对上；格数一乱就说明不是同一张表，直接放弃。
    """
    hit = None
    for b in boxes:
        cs = _size_row(b)
        if cs:
            hit = (b, cs)
            break
    if not hit:
        return [], []
    b, cs = hit
    vals = []
    for o in boxes:
        if o is b:
            continue
        dt = o["t"] - b["t"]
        if not (0 < dt <= 10):
            continue
        rows = [x for x in o["rows"] if x.strip()]
        if not rows:
            continue
        cv = _cols(rows[0])
        if len(cv) == len(cs) and not any(SIZE_CELL_RE.match(c) for c in cv):
            vals = [c.strip() for c in cv]
            break
    return [_canon(c) for c in cs], vals


def parse_ddl(path):
    """读一个 .ddl，返回
    {"size": 号型文本, "size_from": "正文"/"文件名", "fabrics": [成分行…],
     "paper": {"w","h"} 或 None, "sizes": [尺码…], "values": [数值…],
     "file": 路径, "notes": [给人看的提示…]}

    sizes / values 是表格的「尺码行 + 数值行」，认不出就是两个空列表
    （单尺码老标签本来没有表，这不是失败）。
    读不出来才抛 DdlError；字段缺失只写进 notes，不算失败。
    """
    try:
        raw = open(path, "rb").read()
    except OSError as e:
        raise DdlError("文件打不开：%s" % e)
    if not raw:
        raise DdlError("文件是空的")
    try:
        root = ET.fromstring(raw)          # 交给 ET 按文件自带声明解码（UTF-8）
    except ET.ParseError:
        # 少数老文件没写声明，退回 GBK 再试一次
        try:
            root = ET.fromstring(raw.decode("gbk", "replace").encode("utf-8"))
        except Exception:
            raise DdlError("这不是 DLabel 的 .ddl 标签文件（读不出 XML 结构）")

    vals = _texts(root)
    if not vals:
        raise DdlError("文件里没有任何文字，可能是个空标签")

    size, size_from = _size(vals, path)
    fabrics = _fabrics(vals)
    sizes, tvals = _size_table(_boxes(root))
    notes = []
    if not size:
        notes.append("模板里没有号型这一行（帽子 / 无尺码款常见），请自己填")
    elif size_from == "文件名":
        notes.append("号型是从文件名里取的，建议核对")
    if not fabrics:
        notes.append("没找到成分，请自己填")

    return {"size": size, "size_from": size_from, "fabrics": fabrics,
            "paper": _paper(root), "sizes": sizes, "values": tvals,
            "file": path, "notes": notes}
