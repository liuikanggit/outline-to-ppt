#!/usr/bin/env python3
"""
第6步骤：生成 PPT 正文页（章节页 + 内容页）
根据 mubu_parsed_structure.json 的大纲层次结构，递归加入幻灯片。
内容页使用事先做好的、带联动导航条的 SlideMaster / Layout。
"""

import json
import os
import math
from pathlib import Path
from urllib.request import urlretrieve
from lxml import etree
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE

# ----------------- Open XML 命名空间 -----------------
NSMAP = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
}

# ----------------- 预设 物理坐标 (从原模版提取) -----------------
TITLE_L = 676800  # 1.88cm
TITLE_T = 1033200 # 2.87cm
TITLE_W = 6096000
TITLE_H = 368300

BODY_L = 676800   # 1.88cm
BODY_T = 1440000  # 4cm
BODY_W = 10838400 # 12192000 - 676800 * 2
BODY_H = 4818000

def download_image_local(uri):
    """下载图片缓存到本地"""
    if not uri: return None
    if uri.startswith('http'): url = uri
    else: url = f'https://mubu.com/{uri}'
    temp_dir = Path('/tmp/outline_images')
    temp_dir.mkdir(parents=True, exist_ok=True)
    local_filename = temp_dir / os.path.basename(uri)
    if local_filename.exists(): return str(local_filename)
    try: 
        urlretrieve(url, str(local_filename))
        return str(local_filename)
    except Exception: return None

# ----------------- 1. 自动分页与字号算法 -----------------
def paginate_contents_custom(contents):
    """
    智能防打散分页算法：支持降级字号（14pt）和标点智能切割
    """
    def get_lines(txt, level=0):
        max_char = max(1, 48 - (level + 1) * 2)
        return math.ceil(len(txt) / max_char) if txt else 1

    def get_lines_at_14pt(txt, level=0):
        max_char = max(1, 60 - (level + 1) * 2)
        return math.ceil(len(txt) / max_char) if txt else 1

    def build_tree(item, depth):
        txt = item.get('text', '')
        has_bullet = item.get('hasBullet', False)
        node = {
            'text': txt,
            'level': depth,
            'hasBullet': has_bullet or (depth >= 1),
            'lines': get_lines(txt, depth),
            'lines_14pt': get_lines_at_14pt(txt, depth),
            'children': []
        }
        for sub in item.get('subContent', []):
            node['children'].append(build_tree(sub, depth + 1))
        node['block_lines'] = node['lines'] + sum(c['block_lines'] for c in node['children'])
        node['block_lines_14pt'] = node['lines_14pt'] + sum(c['block_lines_14pt'] for c in node['children'])
        return node
    
    forest = []
    for c in contents:
        if c.get('type') == 'text':
            info = c.get('text', {})
            info_list = info if isinstance(info, list) else [info]
            for node_data in info_list:
                if node_data.get('text', ''):
                    forest.append(build_tree(node_data, 0))

    if not forest: return []
    total_lines_16pt = sum(n['block_lines'] for n in forest)
    total_lines_14pt = sum(n['block_lines_14pt'] for n in forest)

    def flatten_tree(nodes):
        flat = []
        def dfs(n):
            flat.append(n); 
            for ch in n['children']: dfs(ch)
        for n in nodes: dfs(n)
        return flat

    print(f"    [分页预测] 16pt行数: {total_lines_16pt} | 14pt行数: {total_lines_14pt}", end="")
    if total_lines_16pt <= 13:
        print(" -> 16pt 装得下")
        return [(flatten_tree(forest), 1600)]
    if total_lines_14pt <= 15:
        print(" -> 14pt 装得下")
        return [(flatten_tree(forest), 1400)]

    print(" -> 触发强制分页 (16pt)")
    pages = []; curr_page = []; curr_lines = 0; MAX_LINES = 13

    def commit_page():
        nonlocal curr_page, curr_lines
        if curr_page:
            pages.append((curr_page, 1600))
            curr_page = []; curr_lines = 0

    def split_text_by_punct(txt, level):
        import re
        def chunk_by_delimiters(text, delimiter_pattern):
            parts = re.split(delimiter_pattern, text)
            parts = [p for p in parts if p]
            res = []; current = ''
            for p in parts:
                if get_lines(current + p, level) <= MAX_LINES: current += p
                else:
                    if current: res.append(current)
                    current = p
            if current: res.append(current)
            return res
        primary = chunk_by_delimiters(txt, r'(?<=[。；！？])')
        final = []
        for c in primary:
            if get_lines(c, level) > MAX_LINES: final.extend(chunk_by_delimiters(c, r'(?<=[，,、])'))
            else: final.append(c)
        return final

    def pack_node(node):
        nonlocal curr_lines
        if curr_lines + node['block_lines'] <= MAX_LINES:
            def dfs_add(n):
                nonlocal curr_lines
                curr_page.append(n); curr_lines += n['lines']
                for ch in n['children']: dfs_add(ch)
            dfs_add(node)
            return
        if node['block_lines'] <= MAX_LINES:
            commit_page()
            def dfs_add2(n):
                nonlocal curr_lines
                curr_page.append(n); curr_lines += n['lines']
                for ch in n['children']: dfs_add2(ch)
            dfs_add2(node)
            return
        if not node['children'] and node['lines'] > MAX_LINES:
            chunks = split_text_by_punct(node['text'], node['level'])
            for ct in chunks:
                cl = get_lines(ct, node['level'])
                if curr_lines + cl > MAX_LINES: commit_page()
                curr_page.append({'text':ct, 'level':node['level'], 'hasBullet':node['hasBullet'], 'lines':cl})
                curr_lines += cl
            return
        if curr_lines + node['lines'] > MAX_LINES: commit_page()
        curr_page.append(node); curr_lines += node['lines']
        for ch in node['children']: pack_node(ch)

    for n in forest: pack_node(n)
    commit_page()
    return pages

# ----------------- 2. 章节页生成 -----------------
def create_chapter_slide(prs, template_slide, chapter):
    new_slide = prs.slides.add_slide(template_slide.slide_layout)
    for ph in list(new_slide.placeholders):
        sp = ph._element
        sp.getparent().remove(sp)
    spTree_new = new_slide.shapes._spTree
    for sp in template_slide.shapes._spTree:
        tag = sp.tag.split('}')[-1]
        if tag == 'sp':
            new_sp = etree.fromstring(etree.tostring(sp))
            spTree_new.append(new_sp)
            txBody = new_sp.find(f'.//{{{NSMAP["p"]}}}txBody')
            if txBody is not None:
                paras = txBody.findall(f'.//{{{NSMAP["a"]}}}p')
                if len(paras) >= 1:
                    r_nodes = paras[0].findall(f'.//{{{NSMAP["a"]}}}r')
                    if r_nodes:
                        t = r_nodes[0].find(f'{{{NSMAP["a"]}}}t')
                        if t is not None: t.text = chapter.get('chapterName', '')
                        for extra_r in r_nodes[1:]: paras[0].remove(extra_r)
                if len(paras) >= 2:
                    r_nodes = paras[1].findall(f'.//{{{NSMAP["a"]}}}r')
                    if r_nodes:
                        t = r_nodes[0].find(f'{{{NSMAP["a"]}}}t')
                        if t is not None: t.text = chapter.get('chapterEnName', '')
                        rPr = r_nodes[0].find(f'{{{NSMAP["a"]}}}rPr')
                        if rPr is None: rPr = etree.SubElement(r_nodes[0], f'{{{NSMAP["a"]}}}rPr')
                        rPr.set('sz', '1800')
                        for c in list(rPr): 
                            if c.tag.endswith('solidFill'): rPr.remove(c)
                        sf = etree.Element(f'{{{NSMAP["a"]}}}solidFill')
                        etree.SubElement(sf, f'{{{NSMAP["a"]}}}srgbClr', val='808080')
                        rPr.insert(0, sf)
                        for extra_r in r_nodes[1:]: paras[1].remove(extra_r)
    return new_slide

# ----------------- 3. 内容页生成 -----------------
def write_title_box(slide, page_title, page_note=""):
    title_box = slide.shapes.add_textbox(TITLE_L, TITLE_T, TITLE_W, TITLE_H)
    title_box.text_frame.auto_size = MSO_AUTO_SIZE.NONE
    t_p = title_box.text_frame.paragraphs[0]
    
    # 第一行：主标题
    r = t_p.add_run()
    r.text = page_title
    rPr = r._r.get_or_add_rPr()
    rPr.set('b', '1')
    try:
        etree.SubElement(rPr, f'{{{NSMAP["a"]}}}latin', typeface='微软雅黑')
        etree.SubElement(rPr, f'{{{NSMAP["a"]}}}ea', typeface='微软雅黑')
    except Exception:
        pass
        
    # 第二行：赘述(如有)
    if page_note:
        r_note = t_p.add_run()
        r_note.text = '\n' + page_note
        r_note.font.name = '微软雅黑 Light'
        r_note.font.size = Pt(11)
        r_note.font.color.rgb = RGBColor(178, 178, 178)
        r_note.font.bold = False

    return title_box

def create_content_slide_optimized(prs, content_layout, page_title, items, font_sz_val, page_note=""):
    new_slide = prs.slides.add_slide(content_layout)
    write_title_box(new_slide, page_title, page_note)
    
    body_box = new_slide.shapes.add_textbox(BODY_L, BODY_T, BODY_W, BODY_H)
    body_box.text_frame.auto_size = MSO_AUTO_SIZE.NONE
    body_box.text_frame.word_wrap = True
    txBody = body_box._element.find('.//' + f'{{{NSMAP["p"]}}}txBody')
    if txBody is None:
        txBody = etree.SubElement(body_box._element, f'{{{NSMAP["p"]}}}txBody')
        etree.SubElement(txBody, f'{{{NSMAP["a"]}}}bodyPr')
        etree.SubElement(txBody, f'{{{NSMAP["a"]}}}lstStyle')
    for p_old in txBody.findall(f'.//{{{NSMAP["a"]}}}p'): txBody.remove(p_old)

    for item in items:
        txt = item.get('text', '')
        level = item.get('level', 0)
        if txt:
            para = etree.SubElement(txBody, f'{{{NSMAP["a"]}}}p')
            pPr = etree.SubElement(para, f'{{{NSMAP["a"]}}}pPr')
            bullet_styles = [(342900, -342900, '●'), (742950, -342900, '●'), (1143000, -342900, '●'), (1543050, -342900, '●'), (1943100, -342900, '●')]
            lvl_idx = min(level, len(bullet_styles) - 1)
            ln = etree.SubElement(pPr, f'{{{NSMAP["a"]}}}lnSpc')
            etree.SubElement(ln, f'{{{NSMAP["a"]}}}spcPct', val='150000')
            if item.get('hasBullet', False):
                marL, ind, char = bullet_styles[lvl_idx]
                pPr.set('marL', str(marL)); pPr.set('indent', str(ind))
                etree.SubElement(pPr, f'{{{NSMAP["a"]}}}buSzPct', val='140000')
                etree.SubElement(pPr, f'{{{NSMAP["a"]}}}buFont', typeface='Arial')
                etree.SubElement(pPr, f'{{{NSMAP["a"]}}}buChar', char=char)
            elif level >= 1:
                pPr.set('marL', str(bullet_styles[lvl_idx][0]))
            
            defRPr = etree.SubElement(pPr, f'{{{NSMAP["a"]}}}defRPr')
            etree.SubElement(defRPr, f'{{{NSMAP["a"]}}}latin', typeface='Microsoft YaHei Light')
            etree.SubElement(defRPr, f'{{{NSMAP["a"]}}}ea', typeface='Microsoft YaHei Light')
            run = etree.SubElement(para, f'{{{NSMAP["a"]}}}r')
            rPr = etree.SubElement(run, f'{{{NSMAP["a"]}}}rPr', sz=str(font_sz_val))
            etree.SubElement(etree.SubElement(rPr, f'{{{NSMAP["a"]}}}solidFill'), f'{{{NSMAP["a"]}}}schemeClr', val='tx1')
            etree.SubElement(rPr, f'{{{NSMAP["a"]}}}latin', typeface='Microsoft YaHei Light')
            etree.SubElement(rPr, f'{{{NSMAP["a"]}}}ea', typeface='Microsoft YaHei Light')
            etree.SubElement(run, f'{{{NSMAP["a"]}}}t').text = txt
    return new_slide

def append_images_to_slide(slide, img_info, prs):
    uri = img_info.get('uri')
    img_path = download_image_local(uri)
    if not img_path: return False
    try:
        MAX_W = 10800000 # 30cm
        MAX_H = 4680000  # 13cm
        TOP_MARGIN = 1440000   # 4cm
        BOTTOM_MARGIN = 576000 # 1.6cm
        AVAILABLE_H = prs.slide_height - TOP_MARGIN - BOTTOM_MARGIN

        w_orig = img_info.get('w', 1470)
        h_orig = img_info.get('h', 786) or 1

        # 锁定长宽比（Aspect Ratio）缩放到宽30cm、高13cm约束盒子内
        fw, fh = MAX_W, int(MAX_W * (h_orig / w_orig))
        if fh > MAX_H:
            fh, fw = MAX_H, int(MAX_H * (w_orig / h_orig))
            
        # 水平居中
        left = (prs.slide_width - fw) // 2
        # 垂直区间内居中
        top = TOP_MARGIN + max(0, (AVAILABLE_H - fh) // 2)
        
        slide.shapes.add_picture(img_path, left, top, width=fw, height=fh)
        return True
    except Exception: return False

def set_cell_border(cell, color_rgb="A6A6A6", width_pt=0.25):
    tcPr = cell._tc.get_or_add_tcPr()
    w_emu = int(width_pt * 12700)
    
    # 严格的 OpenXML Schema: lnL, lnR, lnT, lnB 必须排在 solidFill 之前！
    # 先清理可能游离在末尾的旧线条节点
    for side in ['lnL', 'lnR', 'lnT', 'lnB']:
        ln = tcPr.find(f'{{{NSMAP["a"]}}}{side}')
        if ln is not None: tcPr.remove(ln)
        
    # 在 tcPr 的头部(前4个位置)顺序插入边框，符合微软严格的 XSD 校验标准
    insert_idx = 0
    for side in ['lnL', 'lnR', 'lnT', 'lnB']:
        ln = etree.Element(f'{{{NSMAP["a"]}}}{side}')
        ln.set('w', str(w_emu)); ln.set('cap', 'flat'); ln.set('cmpd', 'sng'); ln.set('algn', 'ctr')
        sf = etree.SubElement(ln, f'{{{NSMAP["a"]}}}solidFill')
        clr = etree.SubElement(sf, f'{{{NSMAP["a"]}}}srgbClr')
        clr.set('val', color_rgb)
        dash = etree.SubElement(ln, f'{{{NSMAP["a"]}}}prstDash')
        dash.set('val', 'solid')
        
        tcPr.insert(insert_idx, ln)
        insert_idx += 1

def add_table_to_slide(slide, data, prs):
    rows = data.get('rows', [])
    if not rows: return
    nr, nc = len(rows), len(rows[0])
    TW = 10800000 # 30cm
    TOP_MARGIN = 1440000 # 4cm
    BOTTOM_MARGIN = 576000 # 1.6cm
    MAX_TABLE_H = prs.slide_height - TOP_MARGIN - BOTTOM_MARGIN
    BASE_H = 450000 # 1.25cm
    
    total_h = nr * BASE_H
    actual_h = BASE_H if total_h <= MAX_TABLE_H else (MAX_TABLE_H // nr)
    total_h = min(total_h, MAX_TABLE_H)
    
    left = (prs.slide_width - TW) // 2
    top = TOP_MARGIN + max(0, (MAX_TABLE_H - total_h) // 2)

    col_lens = [max(1, sum(2 if ord(c) > 255 else 1 for c in str(r[i]))) for i in range(nc) for r in rows] # This logic needs fix
    col_max = [1] * nc
    for r in rows:
        for i, v in enumerate(r):
            col_max[i] = max(col_max[i], sum(2 if ord(c) > 255 else 1 for c in str(v)))
    sum_l = sum(col_max)
    cws = [int(TW * (l / sum_l)) for l in col_max]

    shape = slide.shapes.add_table(nr, nc, left, top, TW, total_h)
    table = shape.table
    for i, w in enumerate(cws): table.columns[i].width = w
    for r_idx, r_data in enumerate(rows):
        table.rows[r_idx].height = actual_h
        for c_idx, val in enumerate(r_data):
            cell = table.cell(r_idx, c_idx)
            cell.text = str(val)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = cell.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
            run = p.runs[0]; run.font.size = Pt(12)
            if r_idx == 0:
                cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor(69, 82, 103)
                run.font.name, run.font.bold, run.font.color.rgb = '微软雅黑', True, RGBColor(255, 255, 255)
            elif c_idx == 0:
                cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor(231, 246, 252)
                run.font.name, run.font.bold, run.font.color.rgb = '微软雅黑', True, RGBColor(0,0,0)
            else:
                cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor(255, 255, 255)
                run.font.name, run.font.bold, run.font.color.rgb = '微软雅黑 Light', False, RGBColor(0, 0, 0)
            set_cell_border(cell)

# ----------------- 4. 路由逻辑 -----------------
def get_layout_by_names(prs, m_name, l_name, default):
    for m in prs.slide_masters:
        mcs = m._element.find(f'{{{NSMAP["p"]}}}cSld')
        mn = mcs.get('name', '') if mcs is not None else ''
        if mn == m_name:
            for l in m.slide_layouts:
                lcs = l._element.find(f'{{{NSMAP["p"]}}}cSld')
                ln = lcs.get('name', '') if lcs is not None else ''
                if ln == l_name or (ln == "default" and not l_name): return l
    return default

def main():
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--input_json", default="output/mubu_parsed_structure.json")
    parser.add_argument("--input_pptx", default="output/master_with_toc.pptx")
    parser.add_argument("--output_pptx", default="output/final_presentation.pptx")
    args = parser.parse_args()
    if not os.path.exists(args.input_json) or not os.path.exists(args.input_pptx): return
    with open(args.input_json, "r", encoding="utf-8") as f: data = json.load(f)
    prs = Presentation(args.input_pptx)
    added_slides = []

    def expand(chapter, ancestor_names):
        cn = chapter.get('chapterName', '')
        note = chapter.get('note', '')
        level = chapter.get('level', 1)
        names = ancestor_names + [cn]
        
        tm, tl = "", ""
        if level == 1: tm, tl = names[0], "default"
        elif level == 2: tm, tl = f"{names[0]}-{names[1]}", "default"
        elif level >= 3: tm, tl = f"{names[0]}-{names[1]}", names[2]

        if level == 1: added_slides.append(create_chapter_slide(prs, prs.slides[2], chapter))
        
        contents = chapter.get('content', [])
        if contents:
            lay = get_layout_by_names(prs, tm, tl, prs.slides[3].slide_layout)
            # Text
            txts = [c for c in contents if c.get('type') == 'text']
            if txts:
                for pg, sz in paginate_contents_custom(txts):
                    added_slides.append(create_content_slide_optimized(prs, lay, cn, pg, sz, note))
            # Image
            for img in [c for c in contents if c.get('type') == 'image']:
                s = prs.slides.add_slide(lay)
                write_title_box(s, cn, note)
                append_images_to_slide(s, img.get('image', {}), prs)
                added_slides.append(s)
            # Table
            for tbl in [c for c in contents if c.get('type') == 'table']:
                s = prs.slides.add_slide(lay)
                write_title_box(s, cn, note)
                add_table_to_slide(s, tbl.get('table', {}), prs)
                added_slides.append(s)

        for sub in chapter.get('subChapter', []): expand(sub, names)

    for ch in data.get("chapters", []): expand(ch, [])

    # =========================================================
    # 清理模板残留页面 + 重排幻灯片顺序
    # =========================================================
    # 模板结构（5页）:
    #   [0] 封面  [1] 目录  [2] 章节样例  [3] 内容样例  [4] 结尾页
    # 目标: 删除 [2] 和 [3]，保留 [0][1] + 动态生成页 + [4]结尾
    # =========================================================
    
    R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    PML_NS_BODY = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
    
    sld_id_list = prs.slides._sldIdLst
    ids = list(sld_id_list)
    
    if len(ids) > 4:
        # 需要物理删除的模板样例页：索引 2 和 3
        template_sample_ids = [ids[2], ids[3]]
        
        # 第一步：从 python-pptx 底层彻底删除这两个 slide 的 Part 和 Relationship
        # python-pptx 1.0.2 保存时会根据 parts 关系链动态重建 [Content_Types].xml，
        # 所以只需断开关系链即可，无需手动操作 Content_Types
        for sld_id_elem in template_sample_ids:
            rId = sld_id_elem.get(f'{{{R_NS}}}id')
            if rId and rId in prs.part.rels:
                rel = prs.part.rels[rId]
                slide_part = rel.target_part
                
                # 删除 slide 自身的所有子关系（如 notesSlide、slideLayout 引用等）
                # 这一步确保 slide part 不再被任何关系链引用，
                # 从而彻底从 package 的 iter_parts 遍历中消失
                for sub_rId in list(slide_part.rels.keys()):
                    slide_part.drop_rel(sub_rId)
                
                # 从 presentation part 中移除该关系
                prs.part.drop_rel(rId)
            
            # 从 sldIdLst 中移除该条目
            sld_id_list.remove(sld_id_elem)
        
        # 第二步：重排剩余幻灯片顺序 —— Cover, TOC, 动态页..., 结尾页
        remaining_ids = list(sld_id_list)
        # 现在 remaining_ids 应该是: [cover, toc, end, ...dynamic_slides...]
        # cover = remaining[0], toc = remaining[1], end = remaining[2], dynamic = remaining[3:]
        if len(remaining_ids) > 2:
            cover = remaining_ids[0]
            toc = remaining_ids[1]
            end = remaining_ids[2]  # 删除了 [2][3] 后，原来的 [4] 变成了 [2]
            dynamic_slides = remaining_ids[3:]
            
            for i in remaining_ids:
                sld_id_list.remove(i)
            
            sld_id_list.append(cover)
            sld_id_list.append(toc)
            for i in dynamic_slides:
                sld_id_list.append(i)
            sld_id_list.append(end)
    
    # 确保所有幻灯片的 grpSpPr 都拥有 xfrm 以满足 Office 严格的 XSD 校验标准
    for slide in prs.slides:
        spTree = slide.shapes._spTree
        grpSpPr = spTree.find('{http://schemas.openxmlformats.org/presentationml/2006/main}grpSpPr')
        if grpSpPr is not None and len(grpSpPr) == 0:
            xfrm = etree.fromstring('<a:xfrm xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm>')
            grpSpPr.append(xfrm)

    prs.save(args.output_pptx)
    print(f"✅ 生成完成: {args.output_pptx}")

if __name__ == "__main__": main()
