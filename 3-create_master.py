"""
create_master.py
================
根据 test.json 数据，以 PPT模板.pptx 中母版0（slideMaster2）为基础，
生成一个新的 slideMaster，右上角章节信息完全由 JSON 驱动。

逻辑对标 index.html：
  - level1Chapters  → 一级章节 tab（删除旧文本框，按坐标槽位重新生成；活跃者由"组合26"背景块高亮）
  - level2Chapters  → 二级章节 bar（删除旧圆角矩形，重新生成；活跃者加粗+深色字体）
  - level3Chapters  → 三级章节 bar（同 level2 样式，可选）
  - tplName         → 母版名称

用法：
    python create_master.py
    python create_master.py --template tpl/PPT模板.pptx --input test.json --output output/out.pptx
"""

import json
import os
import re
import zipfile
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from lxml import etree


# ==================== XML 命名空间 ====================
PML_NS   = "http://schemas.openxmlformats.org/presentationml/2006/main"
DML_NS   = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS     = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
RELS_NS  = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS    = "http://schemas.openxmlformats.org/package/2006/content-types"

NSMAP = {"p": PML_NS, "a": DML_NS, "r": R_NS}


# ==================== 母版0 的精确几何参数（从模板提取）====================
# 四个章节 tab 文本框的 (x, y, cx, cy) — 从左到右，单位：EMU
CHAPTER_SLOTS: List[Tuple[int, int, int, int]] = [
    (6765290,  6350, 1468755, 380365),   # 文本框 27（槽位 0）
    (8162290,  6350, 1468120, 368300),   # 文本框 28（槽位 1）
    (9515475, 19050, 1300480, 368300),   # 文本框 29（槽位 2）
    (10815955,19050, 1300480, 368300),   # 文本框 30（槽位 3）
]

# 组合26（活跃章节背景块）相对于其对应槽位 x 的偏移量
# 计算：组合26.x - 文本框27.x = 6639560 - 6765290 = -125730
DECORATOR_X_OFFSET = 6639560 - 6765290  # = -125730

# 二级章节 bar 几何
# 距页面顶部 1.19cm, 距右侧 0.41cm
SECTION_BAR_Y      = int(1.19 * 360000)   # 428400 EMU
SECTION_BAR_HEIGHT = int(0.71 * 360000)   # 255600 EMU (文本框高度 0.71cm)
SECTION_BAR_RIGHT  = 12192000 - int(0.41 * 360000)  # 页面宽度(33.867cm=12192000) - 0.41cm

# 三级章节 bar，放在二级 bar 下方（间距 0.1cm）
SUBSECTION_GAP     = int(0.1 * 360000)    # 36000 EMU
SUBSECTION_BAR_Y   = SECTION_BAR_Y + SECTION_BAR_HEIGHT + SUBSECTION_GAP
SUBSECTION_BAR_HEIGHT = SECTION_BAR_HEIGHT  # 三级与二级同样式同高度

# 可支持的最大章节 slot 数（模板固定 4 个槽位）
MAX_SLOTS = len(CHAPTER_SLOTS)

# 源母版文件（PPT模板.pptx 中的 slideMaster2）
SOURCE_MASTER_PATH = "ppt/slideMasters/slideMaster2.xml"


# ==================== 工具函数 ====================

def _read_json(path: Path) -> Dict:
    """读取 JSON 文件，返回字典"""
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def _max_index(paths: List[str], pattern: str) -> int:
    """从文件路径列表中提取最大的数字索引"""
    max_i = 0
    rx = re.compile(pattern)
    for p in paths:
        m = rx.search(p)
        if m:
            max_i = max(max_i, int(m.group(1)))
    return max_i


def _next_rid(rels_tree: etree._Element) -> str:
    """计算 rels 文件中的下一个 rId"""
    max_rid = 0
    for rel in rels_tree.findall(f"{{{RELS_NS}}}Relationship"):
        rid = rel.get("Id", "")
        if rid.startswith("rId"):
            try:
                max_rid = max(max_rid, int(rid[3:]))
            except ValueError:
                pass
    return f"rId{max_rid + 1}"


def _ensure_ct_override(ct_tree: etree._Element, part_name: str, content_type: str) -> None:
    """确保 [Content_Types].xml 中存在对应的 Override 条目"""
    for ov in ct_tree.findall(f"{{{CT_NS}}}Override"):
        if ov.get("PartName") == part_name:
            return
    ov = etree.SubElement(ct_tree, f"{{{CT_NS}}}Override")
    ov.set("PartName", part_name)
    ov.set("ContentType", content_type)


def _append_pres_master(pres_tree: etree._Element,
                        pres_rels_tree: etree._Element,
                        new_master_target: str) -> None:
    """在 presentation.xml 和 rels 中注册新母版"""
    rid = _next_rid(pres_rels_tree)

    # 添加 rels 条目
    rel = etree.SubElement(pres_rels_tree, f"{{{RELS_NS}}}Relationship")
    rel.set("Id", rid)
    rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster")
    rel.set("Target", new_master_target)

    # 添加 sldMasterId 条目
    lst = pres_tree.find(f"{{{PML_NS}}}sldMasterIdLst")
    if lst is None:
        lst = etree.SubElement(pres_tree, f"{{{PML_NS}}}sldMasterIdLst")

    max_id = 2147483648
    for m in lst.findall(f"{{{PML_NS}}}sldMasterId"):
        try:
            max_id = max(max_id, int(m.get("id", "0")))
        except ValueError:
            pass

    new_el = etree.SubElement(lst, f"{{{PML_NS}}}sldMasterId")
    new_el.set("id", str(max_id + 1))
    new_el.set(f"{{{R_NS}}}id", rid)


# ==================== 颜色构建 ====================

def _make_active_color() -> etree._Element:
    """活跃项文字颜色（tx1 lumMod=75000 lumOff=25000，深色）"""
    fill = etree.Element(f"{{{DML_NS}}}solidFill")
    clr  = etree.SubElement(fill, f"{{{DML_NS}}}schemeClr", val="tx1")
    etree.SubElement(clr, f"{{{DML_NS}}}lumMod", val="75000")
    etree.SubElement(clr, f"{{{DML_NS}}}lumOff", val="25000")
    return fill


def _make_normal_color() -> etree._Element:
    """非活跃项文字颜色（bg1 lumMod=75000，浅色）"""
    fill = etree.Element(f"{{{DML_NS}}}solidFill")
    clr  = etree.SubElement(fill, f"{{{DML_NS}}}schemeClr", val="bg1")
    etree.SubElement(clr, f"{{{DML_NS}}}lumMod", val="75000")
    return fill


def _make_tab_color() -> etree._Element:
    """一级章节 tab 文字颜色（bg1 lumMod=75000，浅色，活跃由背景图区分）"""
    return _make_normal_color()


# ==================== 一级章节：生成方案 ====================

def _build_chapter_sp(zh_text: str,
                      en_text: str,
                      x: int, y: int,
                      cx: int, cy: int,
                      sp_id: int,
                      sp_name: str,
                      en_sz: int = 600,
                      active: bool = False) -> etree._Element:
    """
    生成一个一级章节 Tab 文本框（两段：中文 + 英文）。
    - 中文段：sz=1200，b=1，微软雅黑，浅色（bg1 lumMod=75000）
    - 英文段：sz=600，b=0，Microsoft YaHei Light，同色
    - 禁止换行（wrap=none），居中对齐
    活跃/非活跃样式完全由"组合26"背景图决定，文字颜色统一（不区分活跃）。
    """
    A = DML_NS
    P = PML_NS

    def _fill() -> etree._Element:
        """统一的 tab 文字颜色 (活跃时纯白)"""
        if active:
            fill = etree.Element(f"{{{DML_NS}}}solidFill")
            etree.SubElement(fill, f"{{{DML_NS}}}srgbClr", val="FFFFFF")
            return fill
        return _make_tab_color()

    sp = etree.Element(f"{{{P}}}sp")

    # nvSpPr
    nvSpPr = etree.SubElement(sp, f"{{{P}}}nvSpPr")
    etree.SubElement(nvSpPr, f"{{{P}}}cNvPr", id=str(sp_id), name=sp_name)
    etree.SubElement(nvSpPr, f"{{{P}}}cNvSpPr", txBox="1")
    etree.SubElement(nvSpPr, f"{{{P}}}nvPr", userDrawn="1")

    # spPr
    spPr = etree.SubElement(sp, f"{{{P}}}spPr")
    xfrm = etree.SubElement(spPr, f"{{{A}}}xfrm")
    etree.SubElement(xfrm, f"{{{A}}}off", x=str(x), y=str(y))
    etree.SubElement(xfrm, f"{{{A}}}ext", cx=str(cx), cy=str(cy))
    prstGeom = etree.SubElement(spPr, f"{{{A}}}prstGeom", prst="rect")
    etree.SubElement(prstGeom, f"{{{A}}}avLst")
    etree.SubElement(spPr, f"{{{A}}}noFill")

    # txBody
    txBody = etree.SubElement(sp, f"{{{P}}}txBody")
    bodyPr = etree.SubElement(txBody, f"{{{A}}}bodyPr",
                              wrap="none", rtlCol="0",
                              anchor="t", tIns="64800", bIns="0", lIns="0", rIns="0")
    etree.SubElement(bodyPr, f"{{{A}}}spAutoFit")
    etree.SubElement(txBody, f"{{{A}}}lstStyle")

    # ---- 段落1：中文章节名 ----
    p1 = etree.SubElement(txBody, f"{{{A}}}p")
    etree.SubElement(p1, f"{{{A}}}pPr", algn="ctr")

    r1 = etree.SubElement(p1, f"{{{A}}}r")
    rPr1 = etree.SubElement(r1, f"{{{A}}}rPr",
                            kumimoji="1", lang="zh-CN", altLang="en-US",
                            sz="1200", b="1", i="0", dirty="0")
    rPr1.append(_fill())
    etree.SubElement(rPr1, f"{{{A}}}latin",
                     typeface="微软雅黑",
                     panose="020B0503020204020204",
                     pitchFamily="34", charset="-122")
    etree.SubElement(rPr1, f"{{{A}}}ea",
                     typeface="微软雅黑",
                     panose="020B0503020204020204",
                     pitchFamily="34", charset="-122")
    t1 = etree.SubElement(r1, f"{{{A}}}t")
    t1.text = zh_text

    endPr1 = etree.SubElement(p1, f"{{{A}}}endParaRPr",
                               kumimoji="1", lang="en-US", altLang="zh-CN",
                               sz="1200", b="1", i="0", dirty="0")
    endPr1.append(_fill())
    etree.SubElement(endPr1, f"{{{A}}}latin",
                     typeface="微软雅黑",
                     panose="020B0503020204020204",
                     pitchFamily="34", charset="-122")
    etree.SubElement(endPr1, f"{{{A}}}ea",
                     typeface="微软雅黑",
                     panose="020B0503020204020204",
                     pitchFamily="34", charset="-122")

    # ---- 段落2：英文章节名 ----
    p2 = etree.SubElement(txBody, f"{{{A}}}p")
    etree.SubElement(p2, f"{{{A}}}pPr", algn="ctr")
    
    r2 = etree.SubElement(p2, f"{{{A}}}r")
    rPr2 = etree.SubElement(r2, f"{{{A}}}rPr",
                            kumimoji="1", lang="en-GB", altLang="zh-CN",
                            sz=str(en_sz), b="0", i="0", dirty="0")
    rPr2.append(_fill())
    etree.SubElement(rPr2, f"{{{A}}}latin",
                     typeface="Microsoft YaHei Light",
                     panose="020B0503020204020204",
                     pitchFamily="34", charset="-122")
    etree.SubElement(rPr2, f"{{{A}}}ea",
                     typeface="Microsoft YaHei Light",
                     panose="020B0503020204020204",
                     pitchFamily="34", charset="-122")
    t2 = etree.SubElement(r2, f"{{{A}}}t")
    t2.text = en_text  # 为空时段落高度约为 0，不占版面

    endPr2 = etree.SubElement(p2, f"{{{A}}}endParaRPr",
                               kumimoji="1", lang="zh-CN", altLang="en-US",
                               sz=str(en_sz), b="0", i="0", dirty="0")
    endPr2.append(_fill())
    etree.SubElement(endPr2, f"{{{A}}}latin",
                     typeface="Microsoft YaHei Light",
                     panose="020B0503020204020204",
                     pitchFamily="34", charset="-122")
    etree.SubElement(endPr2, f"{{{A}}}ea",
                     typeface="Microsoft YaHei Light",
                     panose="020B0503020204020204",
                     pitchFamily="34", charset="-122")

    return sp


def _estimate_text_width_exact(text: str) -> int:
    """通过字符匹配精确估算宽度 (微软雅黑 12pt)"""
    total = 0
    for ch in text:
        # 中文汉字以及全角标点
        if '\u4e00' <= ch <= '\u9fff' or ch in '，。；：！？（）【】《》“”：':
            total += 134000 # 微调
        elif ch.isspace():
            total += 67000  # 空格
        elif ch.isupper():
            total += 86000
        elif ch.isdigit():
            total += 74000
        elif ch.islower():
            total += 70000
        else:
            total += 67000
    return total

def _estimate_tab_width(zh_text: str) -> int:
    """按字符数动态计算一级章节 Tab 的宽度 (EMU)
       宽度 = 字数 * 0.43 + 1.79 (单位 cm)
    """
    L_zh = len(zh_text)
    return int((L_zh * 0.43 + 1.79) * 360000)

def _rebuild_level1_tabs(spTree: etree._Element,
                         level1: List[Dict],
                         active_idx: int) -> None:
    """
    生成方案：
    ① 删除模板中的原有章节文本框
    ② **靠右对齐**：计算总宽度，从右边锚点往左排布
    ③ 生成新 Tab 文本框 (支持多个)
    ④ 移动并 resize 组合 26（活跃背景块）
    """
    # ① 删除旧的章节文本框
    old_names = {"文本框 27", "文本框 28", "文本框 29", "文本框 30"}
    to_remove = []
    for sp in spTree.findall(f"{{{PML_NS}}}sp"):
        cNvPr = sp.find(f".//{{{PML_NS}}}cNvPr")
        if cNvPr is not None and cNvPr.get("name") in old_names:
            to_remove.append(sp)
    for sp in to_remove:
        spTree.remove(sp)

    if not level1:
        return

    # 🌟 自适应计算全局英文字号 (5-7pt) 🌟
    target_pt = 6 
    fits_7 = True
    fits_6 = True
    for ch in level1:
         w_zh = _estimate_text_width_exact(ch.get("zh", ""))
         # 估算高度比例进行 X 预估
         w_en_7 = _estimate_text_width_exact(ch.get("en", "")) * 7 // 12
         w_en_6 = _estimate_text_width_exact(ch.get("en", "")) * 6 // 12
         if w_en_7 > w_zh: fits_7 = False
         if w_en_6 > w_zh: fits_6 = False
    
    if fits_7: target_pt = 7
    elif fits_6: target_pt = 6
    else: target_pt = 5
    en_sz_val = target_pt * 100

    # ② 动态计算坐标槽位（从右往左锚定）
    widths = [_estimate_tab_width(ch.get("zh", "")) for ch in level1]
    tab_gap = 0 # 取消间距，统一使用文本框水平排列
    total_slots_width = sum(widths)
    
    RIGHT_ANCHOR = 12116435 # 原始槽位 3 的右边界
    current_x = RIGHT_ANCHOR - total_slots_width # 起始 x
    
    y = 0 # 文本框与页面顶部对齐
    cy = int(0.8 * 360000) # 文本框高度 = 0.8cm = 288000

    slots = []
    for cx in widths:
        slots.append((current_x, y, cx, cy))
        current_x += cx

    # ③ 为每个有数据的槽位生成新文本框 (全部渲染)
    if not slots:
        return

    active_slot = min(active_idx, len(level1) - 1) if level1 else 0

    for i in range(len(level1)):
        ch = level1[i]
        x, y, cx, cy = slots[i]
        sp = _build_chapter_sp(
            zh_text=ch.get("zh", ""),
            en_text=ch.get("en", ""),
            x=x, y=y, cx=cx, cy=cy,
            sp_id=200 + i,
            sp_name=f"ch_tab_{i}",
            en_sz=en_sz_val,
            active=(i == active_slot)
        )
        spTree.append(sp)

    # ④ 移除旧“组合 26” 并进行图片无缝拼接 🖼️
    for grpSp in spTree.findall(f"{{{PML_NS}}}grpSp"):
        cNvPr = grpSp.find(f".//{{{PML_NS}}}cNvPr")
        if cNvPr is not None and "组合 26" in cNvPr.get("name", ""):
            spTree.remove(grpSp)
            break

    active_slot = min(active_idx, len(level1) - 1) if level1 else 0
    slot_x, _, slot_cx, _ = slots[active_slot]
    L_zh = len(level1[active_slot].get("zh", ""))

    # 常量坐标比例计算
    W_L = 104601   # 418407 / 4
    W_R = 107369   # 418407 / 4 * (155/151)
    PIC_H = 104601 # 418407 / 4
    
    # 背景图中段色块宽高配置
    MID_H = int(1.12 * 360000) # 1.12cm
    mid_cx = int((L_zh * 0.43 + 0.75) * 360000) # 中间色块宽度 = 字数*0.43+0.75
    
    OVERLAP = 10800 # 极微量重叠(约0.03cm)，消除缝隙但不产生可见交叠

    # 色块与文本框顶面对齐且水平居中对齐在文本框内
    anchor_y = 0 
    mid_x = slot_x + (slot_cx - mid_cx) // 2

    # 左边色块出现微小空隙，额外向中间靠拢 0.005cm (1800 EMU)
    left_x = mid_x - W_L + OVERLAP + 1800
    right_x = mid_x + mid_cx - OVERLAP

    def _create_pic_node(p_id, name, r_id, x, y, cx, cy):
         pic = etree.Element(f"{{{PML_NS}}}pic")
         nvPicPr = etree.SubElement(pic, f"{{{PML_NS}}}nvPicPr")
         etree.SubElement(nvPicPr, f"{{{PML_NS}}}cNvPr", id=str(p_id), name=name)
         etree.SubElement(nvPicPr, f"{{{PML_NS}}}cNvPicPr")
         etree.SubElement(nvPicPr, f"{{{PML_NS}}}nvPr")
         
         blipFill = etree.SubElement(pic, f"{{{PML_NS}}}blipFill")
         blip = etree.SubElement(blipFill, f"{{{DML_NS}}}blip")
         blip.set("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed", r_id)
         stretch = etree.SubElement(blipFill, f"{{{DML_NS}}}stretch")
         etree.SubElement(stretch, f"{{{DML_NS}}}fillRect")
         
         spPr = etree.SubElement(pic, f"{{{PML_NS}}}spPr")
         xfrm = etree.SubElement(spPr, f"{{{DML_NS}}}xfrm")
         etree.SubElement(xfrm, f"{{{DML_NS}}}off", x=str(x), y=str(y))
         etree.SubElement(xfrm, f"{{{DML_NS}}}ext", cx=str(cx), cy=str(cy))
         prstGeom = etree.SubElement(spPr, f"{{{DML_NS}}}prstGeom", prst="rect")
         etree.SubElement(prstGeom, f"{{{DML_NS}}}avLst")
         return pic

    # ⚠️ 倒序insert(2, ...) 确保置于底层 (在 nvGrpSpPr, grpSpPr 之后)
    spTree.insert(2, _create_pic_node(1003, "bg_right", "rId6", right_x, anchor_y, W_R, PIC_H))
    spTree.insert(2, _create_pic_node(1002, "bg_mid", "rId5", mid_x, anchor_y, mid_cx, MID_H))
    spTree.insert(2, _create_pic_node(1001, "bg_left", "rId4", left_x, anchor_y, W_L, PIC_H))


# ==================== 二/三级章节 bar：生成方案 ====================

def _make_bar_run(parent: etree._Element,
                  text: str,
                  is_active: bool) -> None:
    """
    在 parent（a:p）下追加一个文字 run。
    活跃项：加粗 + 深色（tx1 lumMod=75000 lumOff=25000）
    非活跃：不加粗 + 浅色（bg1 lumMod=75000）
    """
    r = etree.SubElement(parent, f"{{{DML_NS}}}r")
    rPr = etree.SubElement(r, f"{{{DML_NS}}}rPr",
                           kumimoji="1", lang="zh-CN", altLang="en-US",
                           sz="1200",
                           b="1" if is_active else "0",
                           i="0", dirty="0")
    rPr.append(_make_active_color() if is_active else _make_normal_color())
    font = "微软雅黑" if is_active else "Microsoft YaHei Light"
    etree.SubElement(rPr, f"{{{DML_NS}}}latin",
                     typeface=font,
                     panose="020B0503020204020204",
                     pitchFamily="34", charset="-122")
    etree.SubElement(rPr, f"{{{DML_NS}}}ea",
                     typeface=font,
                     panose="020B0503020204020204",
                     pitchFamily="34", charset="-122")
    t = etree.SubElement(r, f"{{{DML_NS}}}t")
    t.text = text


def _estimate_bar_width(items: List[Dict]) -> int:
    """根据章节列表估算 bar 色块宽度（EMU）
    公式：色块宽度 = 标题总字数 * 0.43 + (标题数 - 1) * 0.7 + 2.2  (cm)
    """
    n = len(items)
    if n == 0:
        return int(2.2 * 360000)
    total_chars = sum(len(item.get("zh", "")) for item in items)
    width_cm = total_chars * 0.43 + (n - 1) * 0.7 + 2.2
    return int(width_cm * 360000)


def _build_bar_sp(items: List[Dict],
                  active_idx: int,
                  y: int,
                  height: int,
                  sp_id: int,
                  sp_name: str) -> etree._Element:
    """
    生成章节 bar 文本框。
    所有章节名称在同一段落内，以 6 个空格分隔，居中对齐。
    宽度由公式计算，右对齐到 SECTION_BAR_RIGHT。
    上下左右边距为 0。
    """
    # 根据新公式计算宽度，右边缘对齐 SECTION_BAR_RIGHT
    bar_cx = _estimate_bar_width(items)
    bar_x  = SECTION_BAR_RIGHT - bar_cx  # 右对齐

    sp = etree.Element(f"{{{PML_NS}}}sp")

    # nvSpPr
    nvSpPr = etree.SubElement(sp, f"{{{PML_NS}}}nvSpPr")
    etree.SubElement(nvSpPr, f"{{{PML_NS}}}cNvPr",
                     id=str(sp_id), name=sp_name)
    etree.SubElement(nvSpPr, f"{{{PML_NS}}}cNvSpPr", txBox="1")
    etree.SubElement(nvSpPr, f"{{{PML_NS}}}nvPr", userDrawn="1")

    # spPr
    spPr = etree.SubElement(sp, f"{{{PML_NS}}}spPr")
    xfrm = etree.SubElement(spPr, f"{{{DML_NS}}}xfrm")
    etree.SubElement(xfrm, f"{{{DML_NS}}}off", x=str(bar_x), y=str(y))
    etree.SubElement(xfrm, f"{{{DML_NS}}}ext", cx=str(bar_cx), cy=str(height))

    # 圆角矩形
    prstGeom = etree.SubElement(spPr, f"{{{DML_NS}}}prstGeom", prst="roundRect")
    avLst = etree.SubElement(prstGeom, f"{{{DML_NS}}}avLst")
    etree.SubElement(avLst, f"{{{DML_NS}}}gd", name="adj", fmla="val 50000")

    # 背景填充 F5F5F9
    solidFill = etree.SubElement(spPr, f"{{{DML_NS}}}solidFill")
    etree.SubElement(solidFill, f"{{{DML_NS}}}srgbClr", val="F5F5F9")

    # 边框
    ln = etree.SubElement(spPr, f"{{{DML_NS}}}ln", w="3175")
    lnFill = etree.SubElement(ln, f"{{{DML_NS}}}solidFill")
    lnClr = etree.SubElement(lnFill, f"{{{DML_NS}}}schemeClr", val="bg1")
    etree.SubElement(lnClr, f"{{{DML_NS}}}lumMod", val="65000")

    # txBody
    txBody = etree.SubElement(sp, f"{{{PML_NS}}}txBody")
    bodyPr = etree.SubElement(txBody, f"{{{DML_NS}}}bodyPr",
                              wrap="none",   # 禁止换行
                              anchor="ctr",  # 垂直居中
                              lIns="0", tIns="0", rIns="0", bIns="0",
                              rtlCol="0")
    etree.SubElement(bodyPr, f"{{{DML_NS}}}spAutoFit")
    etree.SubElement(txBody, f"{{{DML_NS}}}lstStyle")

    # 段落：居中对齐
    p = etree.SubElement(txBody, f"{{{DML_NS}}}p")
    etree.SubElement(p, f"{{{DML_NS}}}pPr", algn="ctr")

    # 各章节项之间以 6 个空格分隔，不加头尾空格
    SEP = "      "
    for i, sec in enumerate(items):
        if i > 0:
            _make_bar_run(p, SEP, False)  # 间隔符永远非活跃色
        _make_bar_run(p, sec.get("zh", ""), i == active_idx)

    # endParaRPr
    endPr = etree.SubElement(p, f"{{{DML_NS}}}endParaRPr",
                             kumimoji="1", lang="zh-CN", altLang="en-US",
                             sz="1200", dirty="0")
    endPr.append(_make_normal_color())
    etree.SubElement(endPr, f"{{{DML_NS}}}latin",
                     typeface="Microsoft YaHei Light",
                     panose="020B0503020204020204",
                     pitchFamily="34", charset="-122")
    etree.SubElement(endPr, f"{{{DML_NS}}}ea",
                     typeface="Microsoft YaHei Light",
                     panose="020B0503020204020204",
                     pitchFamily="34", charset="-122")

    return sp


def _rebuild_level2_bar(spTree: etree._Element,
                        level2: List[Dict],
                        active_l2: int) -> None:
    """
    生成方案：删除旧的文本框37，按 level2 数据重新生成二级章节 bar。
    若无 level2 数据则只删旧 bar，不新建。
    """
    # 删除旧 bar（按原始名称匹配）
    to_remove = []
    for sp in spTree.findall(f"{{{PML_NS}}}sp"):
        cNvPr = sp.find(f".//{{{PML_NS}}}cNvPr")
        if cNvPr is not None and cNvPr.get("name") == "文本框 37":
            to_remove.append(sp)
    for sp in to_remove:
        spTree.remove(sp)

    if not level2:
        return

    bar_sp = _build_bar_sp(
        items=level2,
        active_idx=active_l2,
        y=SECTION_BAR_Y,
        height=SECTION_BAR_HEIGHT,
        sp_id=300,
        sp_name="lv2_bar",
    )
    spTree.append(bar_sp)


def _rebuild_level3_bar(spTree: etree._Element,
                        level3: List[Dict],
                        active_l3: int,
                        has_level2: bool) -> None:
    """
    生成方案：删除旧三级 bar（如有），按 level3 数据重新生成三级章节 bar。
    若无 level3 数据则只删旧 bar，不新建。
    位置：若有 level2 则位于其下方，否则与 level2 同位置。
    """
    # 删除旧三级 bar
    to_remove = []
    for sp in spTree.findall(f"{{{PML_NS}}}sp"):
        cNvPr = sp.find(f".//{{{PML_NS}}}cNvPr")
        if cNvPr is not None and cNvPr.get("name") == "lv3_bar":
            to_remove.append(sp)
    for sp in to_remove:
        spTree.remove(sp)

    if not level3:
        return

    y = SUBSECTION_BAR_Y if has_level2 else SECTION_BAR_Y
    bar_sp = _build_bar_sp(
        items=level3,
        active_idx=active_l3,
        y=y,
        height=SUBSECTION_BAR_HEIGHT,
        sp_id=301,
        sp_name="lv3_bar",
    )
    spTree.append(bar_sp)


# ==================== 主函数 ====================

def create_master(template_pptx: Path,
                  input_jsons: List[Path],
                  output_pptx: Path) -> None:
    """
    修改后的复合母版逻辑：
    1. 按照 masterName 聚合 JSON 数据。
    2. 对每个聚合组：生成 1 个 master，在 master 内画 1 级、2 级章节。
    3. 针对该组下的每条具体数据（含独特的 3 级章节及 layoutName）：
       生成专属的 layout 版面，并在 layout 内画 3 级章节。
    """
    if not input_jsons:
         raise ValueError("至少需要提供一个输入 JSON 文件")

    # 分组聚合 JSON 数据
    master_groups = {} # masterName -> list of data dicts
    for json_path in input_jsons:
        data = _read_json(json_path)
        m_name = data.get("masterName", "default")
        if m_name not in master_groups:
            master_groups[m_name] = []
        master_groups[m_name].append(data)

    with zipfile.ZipFile(template_pptx, "r") as zin:
        # 扫描现有文件状态
        existing_masters = [
            p for p in zin.namelist()
            if p.startswith("ppt/slideMasters/slideMaster") and p.endswith(".xml")
        ]
        existing_layouts = [
            p for p in zin.namelist()
            if p.startswith("ppt/slideLayouts/slideLayout") and p.endswith(".xml")
        ]
        next_master_idx = _max_index(existing_masters, r"slideMaster(\d+)\.xml$") + 1
        next_layout_idx = _max_index(existing_layouts, r"slideLayout(\d+)\.xml$") + 1

        # 读取累加核心 XML
        pres_tree      = etree.fromstring(zin.read("ppt/presentation.xml"))
        pres_rels_tree = etree.fromstring(zin.read("ppt/_rels/presentation.xml.rels"))
        ct_tree        = etree.fromstring(zin.read("[Content_Types].xml"))
        theme_tree_base = etree.fromstring(zin.read("ppt/theme/theme2.xml"))

        # 动态提取源母版的 logo 图片引用 (避免硬编码)
        src_master_rels_path = SOURCE_MASTER_PATH.replace("slideMasters/", "slideMasters/_rels/") + ".rels"
        src_rels_tree = etree.fromstring(zin.read(src_master_rels_path))
        logo_image_target = "../media/image1.png"  # 默认值
        for rel in src_rels_tree.findall(f"{{{RELS_NS}}}Relationship"):
            if rel.get("Type", "").endswith("/image"):
                logo_image_target = rel.get("Target", logo_image_target)
                break  # 取第一个 image 类型即为 logo

        masters_to_write = []
        layouts_to_write = []

        curr_master_idx = next_master_idx
        curr_layout_idx = next_layout_idx

        for master_name, group_data in master_groups.items():
            # 取第一项数据作为 Master 层级的渲染依据 (1、2 级选项卡一致)
            first_data = group_data[0]
            
            l1_ch = first_data.get("level1Chapter", {})
            level1    = l1_ch.get("list", [])
            active_l1 = int(l1_ch.get("activeIndex", 0))

            l2_ch = first_data.get("level2Chapter", {})
            level2    = l2_ch.get("list", [])
            active_l2 = int(l2_ch.get("activeIndex", 0))

            # 边界容错
            if not level1:
                print(f"⚠️ 警告: 母版 {master_name} 缺少 level1 数据，跳过。")
                continue
            active_l1 = max(0, min(active_l1, len(level1) - 1))
            if level2:
                active_l2 = max(0, min(active_l2, len(level2) - 1))

            # --- 1. 生成 SlideMaster 及专属 Theme ---
            master_tree = etree.fromstring(zin.read(SOURCE_MASTER_PATH))
            cSld = master_tree.find(f"{{{PML_NS}}}cSld")
            if cSld is not None:
                cSld.set("name", master_name) # masterName -> 母版名
                
            import copy
            theme_tree = copy.deepcopy(theme_tree_base)
            theme_tree.set("name", master_name) # PPTUI 实际上优先展示主题名
                
            spTree_m = master_tree.find(f".//{{{PML_NS}}}spTree")
            if spTree_m is None:
                raise RuntimeError("未找到 spTree")

            _rebuild_level1_tabs(spTree_m, level1, active_l1)
            _rebuild_level2_bar(spTree_m, level2, active_l2)
            # 3级菜单不再在母版中渲染！

            # 清空原有的 sldLayoutIdLst
            sldLayoutIdLst = master_tree.find(f"{{{PML_NS}}}sldLayoutIdLst")
            if sldLayoutIdLst is not None:
                for ch in list(sldLayoutIdLst):
                    sldLayoutIdLst.remove(ch)
            else:
                sldLayoutIdLst = etree.SubElement(master_tree, f"{{{PML_NS}}}sldLayoutIdLst")

            # 动态加载并解析源母版的 rels，继承其原始的 rId 绑定关系，避免与 XML 文件的 pic 引用冲突
            src_m_rels_path = SOURCE_MASTER_PATH.replace("slideMasters/", "slideMasters/_rels/") + ".rels"
            src_m_rels_tree = etree.fromstring(zin.read(src_m_rels_path))
            logo_rId = "rId3"  # 默认兜底
            theme_rId = "rId2" # 默认兜底
            for rel in src_m_rels_tree.findall(f"{{{RELS_NS}}}Relationship"):
                r_type = rel.get("Type", "")
                if r_type.endswith("/image"):
                    logo_rId = rel.get("Id")
                elif r_type.endswith("/theme"):
                    theme_rId = rel.get("Id")

            # 聚合 Master 的 Relationships (指向绝对覆盖路径 logo.png 和专属 theme)
            m_rels_xml_parts = [
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n',
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
                f'<Relationship Id="{logo_rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="{logo_image_target}"/>',
                f'<Relationship Id="{theme_rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme{curr_master_idx}.xml"/>',
                '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/bg_left.png"/>',
                '<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/bg_mid.png"/>',
                '<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/bg_right.png"/>'
            ]

            # 注册母版到全局 PPT/CT
            _append_pres_master(pres_tree, pres_rels_tree, f"slideMasters/slideMaster{curr_master_idx}.xml")
            _ensure_ct_override(ct_tree, f"/ppt/slideMasters/slideMaster{curr_master_idx}.xml", 
                                "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml")
            _ensure_ct_override(ct_tree, f"/ppt/theme/theme{curr_master_idx}.xml", 
                                "application/vnd.openxmlformats-officedocument.theme+xml")

            # --- 2. 遍历同组内的子 JSON，为不同的 3 级配置生成独立的 Layout ---
            for lay_idx, data in enumerate(group_data):
                layout_name = data.get("layoutName", "default")
                
                # 注册此版面
                _ensure_ct_override(ct_tree, f"/ppt/slideLayouts/slideLayout{curr_layout_idx}.xml", 
                                    "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml")

                # 构建一个纯净的基础版面骨架
                l_xml_str = (
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                    '<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
                    ' xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
                    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
                    ' type="blank" preserve="1">'
                    f'<p:cSld name="{layout_name}">'
                    '<p:spTree>'
                    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
                    '<p:grpSpPr/>'
                    '</p:spTree>'
                    '</p:cSld>'
                    '</p:sldLayout>'
                )
                
                l_tree = etree.fromstring(l_xml_str.encode("utf-8"))
                l_spTree = l_tree.find(f".//{{{PML_NS}}}spTree")

                # 在 此版面 的 spTree 图层上渲染专属的 3 级 bar
                l3_ch = data.get("level3Chapter", {})
                level3    = l3_ch.get("list", [])
                active_l3 = int(l3_ch.get("activeIndex", 0))
                if level3:
                    active_l3 = max(0, min(active_l3, len(level3) - 1))
                    _rebuild_level3_bar(l_spTree, level3, active_l3, has_level2=bool(level2))

                # 版面 -> 母版 的回溯指针文件
                l_rels_xml = (
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    f'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"'
                    f' Target="../slideMasters/slideMaster{curr_master_idx}.xml"/>'
                    '</Relationships>'
                ).encode("utf-8")

                layouts_to_write.append({
                    "l_idx": curr_layout_idx,
                    "l_tree": l_tree,
                    "l_rels": l_rels_xml
                })

                # 向当前母版树的 sldLayoutIdLst 回写子版面挂靠
                rId_for_layout = f"rId{lay_idx + 10}" # 为错开基础占用，自 10 计
                new_li = etree.SubElement(sldLayoutIdLst, f"{{{PML_NS}}}sldLayoutId")
                new_li.set("id", str(2147483648 + curr_layout_idx))
                new_li.set(f"{{{R_NS}}}id", rId_for_layout)

                # 向母版的 rels 回写子版面指针
                m_rels_xml_parts.append(f'<Relationship Id="{rId_for_layout}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout{curr_layout_idx}.xml"/>')

                curr_layout_idx += 1

            m_rels_xml_parts.append('</Relationships>')
            
            masters_to_write.append({
                "m_idx": curr_master_idx,
                "m_tree": master_tree,
                "m_rels": "".join(m_rels_xml_parts).encode("utf-8"),
                "theme_tree": theme_tree
            })

            curr_master_idx += 1

        # 批量打包写出最终层
        output_pptx.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(output_pptx, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            overwrite_set = {
                "ppt/presentation.xml",
                "ppt/_rels/presentation.xml.rels",
                "[Content_Types].xml",
            }
            # 复制静态部分
            for item in zin.infolist():
                if item.filename not in overwrite_set:
                    zout.writestr(item, zin.read(item.filename))

            # 🌟 写入背景素材 🖼️
            current_script_dir = os.path.dirname(os.path.abspath(__file__))
            img_map = {
                "ppt/media/bg_left.png": os.path.join(current_script_dir, "tpl", "img", "左侧尖角.png"),
                "ppt/media/bg_mid.png": os.path.join(current_script_dir, "tpl", "img", "中间色块.png"),
                "ppt/media/bg_right.png": os.path.join(current_script_dir, "tpl", "img", "右侧尖角.png")
            }
            for target_p, source_p in img_map.items():
                if os.path.exists(source_p):
                    with open(source_p, 'rb') as f_img:
                        zout.writestr(target_p, f_img.read())

            # 写入覆盖修改后的累加核心文件
            zout.writestr("ppt/presentation.xml", etree.tostring(pres_tree, xml_declaration=True, encoding="UTF-8", standalone=True))
            zout.writestr("ppt/_rels/presentation.xml.rels", etree.tostring(pres_rels_tree, xml_declaration=True, encoding="UTF-8", standalone=True))
            zout.writestr("[Content_Types].xml", etree.tostring(ct_tree, xml_declaration=True, encoding="UTF-8", standalone=True))

            # --- 追加所有的动态 Master ---
            for m_item in masters_to_write:
                m_idx = m_item["m_idx"]
                m_xml_bytes = etree.tostring(m_item["m_tree"], xml_declaration=True, encoding="UTF-8", standalone=True)
                theme_xml_bytes = etree.tostring(m_item["theme_tree"], xml_declaration=True, encoding="UTF-8", standalone=True)
                zout.writestr(f"ppt/slideMasters/slideMaster{m_idx}.xml", m_xml_bytes)
                zout.writestr(f"ppt/slideMasters/_rels/slideMaster{m_idx}.xml.rels", m_item["m_rels"])
                zout.writestr(f"ppt/theme/theme{m_idx}.xml", theme_xml_bytes)
            
            # --- 追加所有的动态 Layout ---
            for l_item in layouts_to_write:
                l_idx = l_item["l_idx"]
                l_xml_bytes = etree.tostring(l_item["l_tree"], xml_declaration=True, encoding="UTF-8", standalone=True)
                zout.writestr(f"ppt/slideLayouts/slideLayout{l_idx}.xml", l_xml_bytes)
                zout.writestr(f"ppt/slideLayouts/_rels/slideLayout{l_idx}.xml.rels", l_item["l_rels"])

    print(f"✅ 已成功将 [{len(input_jsons)}] 份 JSON 配置聚合生成：合并输出 [{len(masters_to_write)}] 个主母版 和 [{len(layouts_to_write)}] 个下挂版面，打入包 -> {output_pptx}")


def main() -> None:
    import argparse
    parser = argparse.ArgumentParser(description="自动读取目录下所有 JSON 配置文件合并生成一个 PPT 的母版库")
    parser.add_argument("--template", default="tpl/PPT模板.pptx", help="输入 PPTX 模板路径")
    parser.add_argument("--input", default="output/master", help="包含 JSON 文件的目录")
    parser.add_argument("--output", default="output/master.pptx", help="输出 PPTX 路径")
    args = parser.parse_args()

    input_dir = Path(args.input)
    if not input_dir.exists() or not input_dir.is_dir():
        print(f"❌ 目录存在或不是文件夹: {args.input}")
        return

    import re
    def _sort_key(p):
        numbers = re.findall(r'\d+', p.name)
        return [int(n) for n in numbers]

    jsons = sorted(input_dir.glob("*.json"), key=_sort_key)
    if not jsons:
        print(f"❌ 目录 {args.input} 下未找到任何 .json 文件")
        return

    print(f"🔍 自动扫描到 {len(jsons)} 个 JSON 配置文件，准备批量合流。")
    create_master(
        Path(args.template),
        jsons,
        Path(args.output),
    )


if __name__ == "__main__":
    main()
