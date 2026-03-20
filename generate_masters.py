"""
根据 input.md 的大纲结构，以母版0为模板，生成对应数量的幻灯片母版。
每个小节对应一个母版，命名为 母版1-1-1、母版1-1-2 等。
当前章/节/小节会高亮显示。
"""
import copy
import re
from pptx import Presentation
from pptx.util import Emu
from lxml import etree

# ==================== 常量定义 ====================
# XML 命名空间
NSMAP = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}

# 中文数字映射
CN_NUMBERS = ['一', '二', '三', '四', '五', '六', '七', '八', '九', '十']
# 英文序数映射
EN_NUMBERS = ['One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine', 'Ten']

# 母版0在 slide_masters 中的索引（0-indexed）
SOURCE_MASTER_INDEX = 1

# 章节文本框的位置参数（从母版0中提取的）
CHAPTER_BOX_Y = 19172
CHAPTER_BOX_WIDTH = 1163782
CHAPTER_BOX_HEIGHT = 369332

# 原始8章的间距: 每章 1062188 EMU（含间隔）
CHAPTER_BOX_STEP = 1062188  # 章节中心间距

# 原始右边缘参考值
# 第八章 right = 10952790 + 1163782 = 12116572
RIGHT_MARGIN = 12116572  # 章节行右边缘
SECTION_RIGHT_MARGIN = 12049637  # 节行右边缘
SUBSECTION_RIGHT_MARGIN = 12049760  # 小节行右边缘

# 节行参数
SECTION_BOX_Y = 427063
SECTION_BOX_HEIGHT = 259675

# 小节行参数
SUBSECTION_BOX_Y = 728980
SUBSECTION_BOX_HEIGHT = 243840

# 幻灯片宽度
SLIDE_WIDTH = 12192000  # 33.87 cm

# 顶行的装饰组合形状参数（组合26 - 高亮章节的背景）
DECORATOR_GROUP_WIDTH = 997398
DECORATOR_GROUP_HEIGHT = 400620
# 装饰组合相对于对齐章节的偏移
# 原始: 第一章 x=3517476, 组合26 x=3602038, 差值 = 84562
DECORATOR_OFFSET_X = 84562


# ==================== 解析数据 ====================
def parse_input(filepath):
    """通过 mubu_parser 加载幕布数据"""
    from mubu_parser import load_mubu_data
    res = load_mubu_data(filepath)
    # 兼容处理：在 generate_masters 里只关心树形 outline
    if isinstance(res, dict):
        return res.get('outline', [])
    return res


# ==================== XML 构建辅助函数 ====================
def make_solid_fill_highlight():
    """创建高亮颜色的填充 XML（深灰色，用于当前项）"""
    fill = etree.SubElement(etree.Element('dummy'), f'{{{NSMAP["a"]}}}solidFill')
    clr = etree.SubElement(fill, f'{{{NSMAP["a"]}}}schemeClr', val='tx1')
    etree.SubElement(clr, f'{{{NSMAP["a"]}}}lumMod', val='75000')
    etree.SubElement(clr, f'{{{NSMAP["a"]}}}lumOff', val='25000')
    return fill


def make_solid_fill_normal():
    """创建非高亮颜色的填充 XML（浅灰色，用于非当前项）"""
    fill = etree.SubElement(etree.Element('dummy'), f'{{{NSMAP["a"]}}}solidFill')
    clr = etree.SubElement(fill, f'{{{NSMAP["a"]}}}schemeClr', val='bg1')
    etree.SubElement(clr, f'{{{NSMAP["a"]}}}lumMod', val='75000')
    return fill


def make_solid_fill_chapter_highlight():
    """创建章节高亮颜色（纯白色，用于当前章）"""
    fill = etree.SubElement(etree.Element('dummy'), f'{{{NSMAP["a"]}}}solidFill')
    etree.SubElement(fill, f'{{{NSMAP["a"]}}}schemeClr', val='bg1')
    return fill


def make_solid_fill_chapter_normal():
    """创建章节非高亮颜色（浅灰色，用于非当前章）"""
    fill = etree.SubElement(etree.Element('dummy'), f'{{{NSMAP["a"]}}}solidFill')
    clr = etree.SubElement(fill, f'{{{NSMAP["a"]}}}schemeClr', val='bg1')
    etree.SubElement(clr, f'{{{NSMAP["a"]}}}lumMod', val='75000')
    return fill


def set_run_fill(rpr_elem, fill_elem):
    """设置 run 属性元素的填充颜色"""
    # 先移除已有的 solidFill
    for old_fill in rpr_elem.findall(f'{{{NSMAP["a"]}}}solidFill'):
        rpr_elem.remove(old_fill)
    # 插入新的 fill（放在最前面）
    rpr_elem.insert(0, copy.deepcopy(fill_elem))


def create_chapter_textbox(chapter_idx, chapter_name, is_highlight, x_pos, sp_id):
    """创建一个章节文本框 XML 元素"""
    a = NSMAP['a']
    p = NSMAP['p']

    sp = etree.Element(f'{{{p}}}sp')

    # nvSpPr
    nvSpPr = etree.SubElement(sp, f'{{{p}}}nvSpPr')
    cNvPr = etree.SubElement(nvSpPr, f'{{{p}}}cNvPr',
                             id=str(sp_id), name=f'文本框 {sp_id - 1}')
    cNvSpPr = etree.SubElement(nvSpPr, f'{{{p}}}cNvSpPr', txBox='1')
    nvPr = etree.SubElement(nvSpPr, f'{{{p}}}nvPr', userDrawn='1')

    # spPr
    spPr = etree.SubElement(sp, f'{{{p}}}spPr')
    xfrm = etree.SubElement(spPr, f'{{{a}}}xfrm')
    etree.SubElement(xfrm, f'{{{a}}}off', x=str(x_pos), y=str(CHAPTER_BOX_Y))
    etree.SubElement(xfrm, f'{{{a}}}ext', cx=str(CHAPTER_BOX_WIDTH), cy=str(CHAPTER_BOX_HEIGHT))
    prstGeom = etree.SubElement(spPr, f'{{{a}}}prstGeom', prst='rect')
    etree.SubElement(prstGeom, f'{{{a}}}avLst')
    etree.SubElement(spPr, f'{{{a}}}noFill')

    # txBody
    txBody = etree.SubElement(sp, f'{{{p}}}txBody')
    bodyPr = etree.SubElement(txBody, f'{{{a}}}bodyPr', wrap='square', rtlCol='0')
    etree.SubElement(bodyPr, f'{{{a}}}spAutoFit')
    etree.SubElement(txBody, f'{{{a}}}lstStyle')

    # 颜色选择
    if is_highlight:
        fill = make_solid_fill_chapter_highlight()
    else:
        fill = make_solid_fill_chapter_normal()

    cn_num = CN_NUMBERS[chapter_idx] if chapter_idx < len(CN_NUMBERS) else str(chapter_idx + 1)
    en_num = EN_NUMBERS[chapter_idx] if chapter_idx < len(EN_NUMBERS) else str(chapter_idx + 1)

    # 第一段：中文章名
    p1 = etree.SubElement(txBody, f'{{{a}}}p')
    pPr1 = etree.SubElement(p1, f'{{{a}}}pPr', algn='ctr')

    r1 = etree.SubElement(p1, f'{{{a}}}r')
    rPr1 = etree.SubElement(r1, f'{{{a}}}rPr', kumimoji='1', lang='zh-CN',
                            altLang='en-US', sz='1200', b='1', i='0', dirty='0')
    rPr1.append(copy.deepcopy(fill))
    latin1 = etree.SubElement(rPr1, f'{{{a}}}latin', typeface='微软雅黑',
                              panose='020B0503020204020204', pitchFamily='34', charset='-122')
    ea1 = etree.SubElement(rPr1, f'{{{a}}}ea', typeface='微软雅黑',
                           panose='020B0503020204020204', pitchFamily='34', charset='-122')
    r1_t = etree.SubElement(r1, f'{{{a}}}t')
    r1_t.text = chapter_name

    endPr1 = etree.SubElement(p1, f'{{{a}}}endParaRPr', kumimoji='1', lang='en-US',
                              altLang='zh-CN', sz='1200', b='1', i='0', dirty='0')
    endPr1.append(copy.deepcopy(fill))
    etree.SubElement(endPr1, f'{{{a}}}latin', typeface='微软雅黑',
                     panose='020B0503020204020204', pitchFamily='34', charset='-122')
    etree.SubElement(endPr1, f'{{{a}}}ea', typeface='微软雅黑',
                     panose='020B0503020204020204', pitchFamily='34', charset='-122')

    # 第二段：英文章名
    p2 = etree.SubElement(txBody, f'{{{a}}}p')
    pPr2 = etree.SubElement(p2, f'{{{a}}}pPr', algn='ctr')

    r2 = etree.SubElement(p2, f'{{{a}}}r')
    rPr2 = etree.SubElement(r2, f'{{{a}}}rPr', kumimoji='1', lang='en-GB',
                            altLang='zh-CN', sz='600', b='0', i='0', dirty='0')
    rPr2.append(copy.deepcopy(fill))
    etree.SubElement(rPr2, f'{{{a}}}latin', typeface='Microsoft YaHei Light',
                     panose='020B0503020204020204', pitchFamily='34', charset='-122')
    etree.SubElement(rPr2, f'{{{a}}}ea', typeface='Microsoft YaHei Light',
                     panose='020B0503020204020204', pitchFamily='34', charset='-122')
    r2_t = etree.SubElement(r2, f'{{{a}}}t')
    r2_t.text = ""

    endPr2 = etree.SubElement(p2, f'{{{a}}}endParaRPr', kumimoji='1', lang='zh-CN',
                              altLang='en-US', sz='600', b='0', i='0', dirty='0')
    endPr2.append(copy.deepcopy(fill))
    etree.SubElement(endPr2, f'{{{a}}}latin', typeface='Microsoft YaHei Light',
                     panose='020B0503020204020204', pitchFamily='34', charset='-122')
    etree.SubElement(endPr2, f'{{{a}}}ea', typeface='Microsoft YaHei Light',
                     panose='020B0503020204020204', pitchFamily='34', charset='-122')

    return sp


def create_section_bar(sections, highlight_section_idx, x_pos, width, sp_id):
    """创建节行文本框（圆角矩形带文本）"""
    a = NSMAP['a']
    p_ns = NSMAP['p']

    sp = etree.Element(f'{{{p_ns}}}sp')

    # nvSpPr
    nvSpPr = etree.SubElement(sp, f'{{{p_ns}}}nvSpPr')
    etree.SubElement(nvSpPr, f'{{{p_ns}}}cNvPr', id=str(sp_id), name=f'文本框 {sp_id - 1}')
    etree.SubElement(nvSpPr, f'{{{p_ns}}}cNvSpPr', txBox='1')
    etree.SubElement(nvSpPr, f'{{{p_ns}}}nvPr', userDrawn='1')

    # spPr
    spPr = etree.SubElement(sp, f'{{{p_ns}}}spPr')
    xfrm = etree.SubElement(spPr, f'{{{a}}}xfrm')
    etree.SubElement(xfrm, f'{{{a}}}off', x=str(x_pos), y=str(SECTION_BOX_Y))
    etree.SubElement(xfrm, f'{{{a}}}ext', cx=str(width), cy=str(SECTION_BOX_HEIGHT))
    prstGeom = etree.SubElement(spPr, f'{{{a}}}prstGeom', prst='roundRect')
    avLst = etree.SubElement(prstGeom, f'{{{a}}}avLst')
    etree.SubElement(avLst, f'{{{a}}}gd', name='adj', fmla='val 50000')
    # 背景填充
    solidFill = etree.SubElement(spPr, f'{{{a}}}solidFill')
    etree.SubElement(solidFill, f'{{{a}}}srgbClr', val='F5F5F9')
    # 边框
    ln = etree.SubElement(spPr, f'{{{a}}}ln', w='3175')
    lnFill = etree.SubElement(ln, f'{{{a}}}solidFill')
    lnClr = etree.SubElement(lnFill, f'{{{a}}}schemeClr', val='bg1')
    etree.SubElement(lnClr, f'{{{a}}}lumMod', val='65000')

    # txBody
    txBody = etree.SubElement(sp, f'{{{p_ns}}}txBody')
    bodyPr = etree.SubElement(txBody, f'{{{a}}}bodyPr', wrap='square',
                              lIns='0', tIns='0', rIns='0', bIns='0', rtlCol='0')
    etree.SubElement(bodyPr, f'{{{a}}}spAutoFit')
    etree.SubElement(txBody, f'{{{a}}}lstStyle')

    # 内容段落
    para = etree.SubElement(txBody, f'{{{a}}}p')
    etree.SubElement(para, f'{{{a}}}pPr', algn='ctr')

    separator = '      '  # 6个空格的间隔

    for i in range(len(sections)):
        is_current = (i == highlight_section_idx)
        section_text = f'第{i + 1}节'

        if is_current:
            fill = make_solid_fill_highlight()
            font_name = '微软雅黑'
            bold = '1'
        else:
            fill = make_solid_fill_normal()
            font_name = 'Microsoft YaHei Light'
            bold = '0'

        if i > 0:
            # 在前一项后添加间隔符
            sep_r = etree.SubElement(para, f'{{{a}}}r')
            sep_rPr = etree.SubElement(sep_r, f'{{{a}}}rPr', kumimoji='1',
                                       lang='zh-CN', altLang='en-US', sz='1200',
                                       dirty='0')
            if is_current:
                sep_rPr.set('b', bold)
            sep_rPr.append(copy.deepcopy(fill))
            etree.SubElement(sep_rPr, f'{{{a}}}latin', typeface=font_name,
                             panose='020B0503020204020204', pitchFamily='34', charset='-122')
            etree.SubElement(sep_rPr, f'{{{a}}}ea', typeface=font_name,
                             panose='020B0503020204020204', pitchFamily='34', charset='-122')
            sep_t = etree.SubElement(sep_r, f'{{{a}}}t')
            sep_t.text = separator

        # "第" 字
        r_pre = etree.SubElement(para, f'{{{a}}}r')
        rPr_pre = etree.SubElement(r_pre, f'{{{a}}}rPr', kumimoji='1',
                                    lang='zh-CN', altLang='en-US', sz='1200',
                                    dirty='0')
        if is_current:
            rPr_pre.set('b', bold)
        rPr_pre.append(copy.deepcopy(fill))
        etree.SubElement(rPr_pre, f'{{{a}}}latin', typeface=font_name,
                         panose='020B0503020204020204', pitchFamily='34', charset='-122')
        etree.SubElement(rPr_pre, f'{{{a}}}ea', typeface=font_name,
                         panose='020B0503020204020204', pitchFamily='34', charset='-122')
        t_pre = etree.SubElement(r_pre, f'{{{a}}}t')
        t_pre.text = '第'

        # 数字（英文 run）
        r_num = etree.SubElement(para, f'{{{a}}}r')
        rPr_num = etree.SubElement(r_num, f'{{{a}}}rPr', kumimoji='1',
                                    lang='en-US', altLang='zh-CN', sz='1200',
                                    dirty='0')
        if is_current:
            rPr_num.set('b', bold)
        rPr_num.append(copy.deepcopy(fill))
        etree.SubElement(rPr_num, f'{{{a}}}latin', typeface=font_name,
                         panose='020B0503020204020204', pitchFamily='34', charset='-122')
        etree.SubElement(rPr_num, f'{{{a}}}ea', typeface=font_name,
                         panose='020B0503020204020204', pitchFamily='34', charset='-122')
        t_num = etree.SubElement(r_num, f'{{{a}}}t')
        t_num.text = str(i + 1)

        # "节" 字
        r_suf = etree.SubElement(para, f'{{{a}}}r')
        rPr_suf = etree.SubElement(r_suf, f'{{{a}}}rPr', kumimoji='1',
                                    lang='zh-CN', altLang='en-US', sz='1200',
                                    dirty='0')
        if is_current:
            rPr_suf.set('b', bold)
        rPr_suf.append(copy.deepcopy(fill))
        etree.SubElement(rPr_suf, f'{{{a}}}latin', typeface=font_name,
                         panose='020B0503020204020204', pitchFamily='34', charset='-122')
        etree.SubElement(rPr_suf, f'{{{a}}}ea', typeface=font_name,
                         panose='020B0503020204020204', pitchFamily='34', charset='-122')
        t_suf = etree.SubElement(r_suf, f'{{{a}}}t')
        t_suf.text = '节'

    # endParaRPr
    endPr = etree.SubElement(para, f'{{{a}}}endParaRPr', kumimoji='1',
                             lang='zh-CN', altLang='en-US', sz='1200', dirty='0')
    endPr.append(copy.deepcopy(make_solid_fill_normal()))
    etree.SubElement(endPr, f'{{{a}}}latin', typeface='Microsoft YaHei Light',
                     panose='020B0503020204020204', pitchFamily='34', charset='-122')
    etree.SubElement(endPr, f'{{{a}}}ea', typeface='Microsoft YaHei Light',
                     panose='020B0503020204020204', pitchFamily='34', charset='-122')

    return sp


def create_subsection_bar(subsections, highlight_subsection_idx, x_pos, width, sp_id):
    """创建小节行文本框（圆角矩形带文本）"""
    a = NSMAP['a']
    p_ns = NSMAP['p']

    sp = etree.Element(f'{{{p_ns}}}sp')

    # nvSpPr
    nvSpPr = etree.SubElement(sp, f'{{{p_ns}}}nvSpPr')
    etree.SubElement(nvSpPr, f'{{{p_ns}}}cNvPr', id=str(sp_id), name=f'文本框 {sp_id - 1}')
    etree.SubElement(nvSpPr, f'{{{p_ns}}}cNvSpPr', txBox='1')
    etree.SubElement(nvSpPr, f'{{{p_ns}}}nvPr', userDrawn='1')

    # spPr
    spPr = etree.SubElement(sp, f'{{{p_ns}}}spPr')
    xfrm = etree.SubElement(spPr, f'{{{a}}}xfrm')
    etree.SubElement(xfrm, f'{{{a}}}off', x=str(x_pos), y=str(SUBSECTION_BOX_Y))
    etree.SubElement(xfrm, f'{{{a}}}ext', cx=str(width), cy=str(SUBSECTION_BOX_HEIGHT))
    prstGeom = etree.SubElement(spPr, f'{{{a}}}prstGeom', prst='roundRect')
    avLst = etree.SubElement(prstGeom, f'{{{a}}}avLst')
    etree.SubElement(avLst, f'{{{a}}}gd', name='adj', fmla='val 50000')
    solidFill = etree.SubElement(spPr, f'{{{a}}}solidFill')
    etree.SubElement(solidFill, f'{{{a}}}srgbClr', val='F5F5F9')
    ln = etree.SubElement(spPr, f'{{{a}}}ln', w='3175')
    lnFill = etree.SubElement(ln, f'{{{a}}}solidFill')
    lnClr = etree.SubElement(lnFill, f'{{{a}}}schemeClr', val='bg1')
    etree.SubElement(lnClr, f'{{{a}}}lumMod', val='65000')

    # txBody
    txBody = etree.SubElement(sp, f'{{{p_ns}}}txBody')
    bodyPr = etree.SubElement(txBody, f'{{{a}}}bodyPr', wrap='square',
                              lIns='0', tIns='0', rIns='0', bIns='0', rtlCol='0')
    etree.SubElement(bodyPr, f'{{{a}}}noAutofit')
    etree.SubElement(txBody, f'{{{a}}}lstStyle')

    # 内容段落
    para = etree.SubElement(txBody, f'{{{a}}}p')
    etree.SubElement(para, f'{{{a}}}pPr', algn='ctr')

    separator = '      '  # 6个空格的间隔

    for i in range(len(subsections)):
        is_current = (i == highlight_subsection_idx)
        subsection_text = f'第{i + 1}小节'

        if is_current:
            fill = make_solid_fill_highlight()
            font_name = '微软雅黑'
            bold = '1'
            kern = '1200'
        else:
            fill = make_solid_fill_normal()
            font_name = 'Microsoft YaHei Light'
            bold = '0'
            kern = '1200'

        if i > 0:
            # 间隔符
            sep_r = etree.SubElement(para, f'{{{a}}}r')
            sep_rPr = etree.SubElement(sep_r, f'{{{a}}}rPr', kumimoji='1',
                                       lang='zh-CN', altLang='en-US', sz='1200',
                                       dirty='0')
            sep_rPr.append(copy.deepcopy(fill))
            etree.SubElement(sep_rPr, f'{{{a}}}latin', typeface=font_name,
                             panose='020B0503020204020204', pitchFamily='34', charset='-122')
            etree.SubElement(sep_rPr, f'{{{a}}}ea', typeface=font_name,
                             panose='020B0503020204020204', pitchFamily='34', charset='-122')
            sep_t = etree.SubElement(sep_r, f'{{{a}}}t')
            sep_t.text = separator

        # "第X小节" 作为一个 run
        r = etree.SubElement(para, f'{{{a}}}r')
        rPr = etree.SubElement(r, f'{{{a}}}rPr', kumimoji='1',
                                lang='zh-CN', altLang='en-US', sz='1200',
                                kern=kern, dirty='0')
        if is_current:
            rPr.set('b', bold)
        rPr.append(copy.deepcopy(fill))
        etree.SubElement(rPr, f'{{{a}}}latin', typeface=font_name,
                         panose='020B0503020204020204', pitchFamily='34', charset='-122')
        etree.SubElement(rPr, f'{{{a}}}ea', typeface=font_name,
                         panose='020B0503020204020204', pitchFamily='34', charset='-122')
        etree.SubElement(rPr, f'{{{a}}}cs', typeface='+mn-cs')
        t = etree.SubElement(r, f'{{{a}}}t')
        t.text = f'第{i + 1}小节'

    # endParaRPr
    endPr = etree.SubElement(para, f'{{{a}}}endParaRPr', kumimoji='1',
                             lang='zh-CN', altLang='en-US', sz='1200', dirty='0')
    endPr.append(copy.deepcopy(make_solid_fill_normal()))
    etree.SubElement(endPr, f'{{{a}}}latin', typeface='Microsoft YaHei Light',
                     panose='020B0503020204020204', pitchFamily='34', charset='-122')
    etree.SubElement(endPr, f'{{{a}}}ea', typeface='Microsoft YaHei Light',
                     panose='020B0503020204020204', pitchFamily='34', charset='-122')

    return sp


def calculate_chapter_positions(num_chapters):
    """根据章节数量计算各章文本框的 x 位置（右对齐）"""
    # 保持原始间距 CHAPTER_BOX_STEP = 1062188
    # 最后一章的右边缘 = RIGHT_MARGIN
    # 最后一章的 x = RIGHT_MARGIN - CHAPTER_BOX_WIDTH
    last_chapter_x = RIGHT_MARGIN - CHAPTER_BOX_WIDTH

    positions = []
    for i in range(num_chapters):
        # 从右往左排列
        x = last_chapter_x - (num_chapters - 1 - i) * CHAPTER_BOX_STEP
        positions.append(x)

    return positions


def calculate_section_bar_params(num_sections):
    """根据节的数量计算节行的紧凑位置和宽度（右对齐）"""
    # 原始模板：10个节占宽 6919853 → 每节约 691985
    # 但我们要紧凑包裹，所以按每节的实际文字宽度计算
    # "第X节" = 3个字符 × 约150000 EMU + 分隔符6空格 × 约75000
    # 每个节项含分隔符约: 600000 EMU
    # 不含分隔符的节项: 450000 EMU
    per_section = 600000  # 每项(含半个左右分隔符)的宽度
    padding = 150000  # 左右各 padding
    bar_width = num_sections * per_section + padding
    bar_x = SECTION_RIGHT_MARGIN - bar_width
    return bar_x, bar_width


def calculate_subsection_bar_params(num_subsections):
    """根据小节数量计算小节行的紧凑位置和宽度（右对齐）"""
    # "第X小节" = 4个字符 × 约150000 + 分隔符
    per_subsection = 750000  # 每项(含半个左右分隔符)的宽度
    padding = 150000
    bar_width = num_subsections * per_subsection + padding
    bar_x = SUBSECTION_RIGHT_MARGIN - bar_width
    return bar_x, bar_width


# ==================== 母版复制和修改 ====================
def create_master_for_subsection(prs, source_master, outline, ch_idx, sec_idx, sub_idx):
    """为指定的小节创建一个新的母版

    基于 source_master 的 XML 结构复制，然后修改：
    - 章节文本框的数量和高亮
    - 节行的内容和高亮
    - 小节行的内容和高亮
    """
    a = NSMAP['a']
    p_ns = NSMAP['p']

    # 深拷贝源母版的 XML
    master_xml = copy.deepcopy(source_master._element)

    # 设置母版名称（在 cSld 元素上设置 name 属性，WPS 会显示此名称）
    master_name = f'母版{ch_idx + 1}-{sec_idx + 1}-{sub_idx + 1}'
    cSld = master_xml.find(f'{{{p_ns}}}cSld')
    if cSld is not None:
        cSld.set('name', master_name)

    # 获取 spTree（形状树）
    spTree = master_xml.find(f'.//{{{p_ns}}}spTree')

    # ============ 识别并删除旧的章节文本框、节行、小节行 ============
    shapes_to_remove = []
    for sp in list(spTree):
        # 只处理 sp 元素（形状）
        if sp.tag != f'{{{p_ns}}}sp':
            continue

        # 查找文本框名称
        cNvPr = sp.find(f'.//{{{p_ns}}}cNvPr')
        if cNvPr is None:
            continue
        name = cNvPr.get('name', '')

        # 识别章节文本框（文本框 12 ~ 文本框 30，包含 "第X章" 文本）
        text_content = ''
        for t_elem in sp.findall(f'.//{{{a}}}t'):
            if t_elem.text:
                text_content += t_elem.text

        if '章' in text_content and 'Chapter' in text_content:
            shapes_to_remove.append(sp)
        elif text_content.startswith('第') and '节' in text_content and '小节' not in text_content:
            # 节行
            shapes_to_remove.append(sp)
        elif '小节' in text_content:
            # 小节行
            shapes_to_remove.append(sp)

    for sp in shapes_to_remove:
        spTree.remove(sp)

    # ============ 计算布局参数 ============
    num_chapters = len(outline)
    chapter_positions = calculate_chapter_positions(num_chapters)

    # ============ 更新装饰组合（组合26）的位置 - 移到高亮章下方 ============
    highlight_chapter_x = chapter_positions[ch_idx]
    decorator_x = highlight_chapter_x + DECORATOR_OFFSET_X

    for grpSp in spTree.findall(f'{{{p_ns}}}grpSp'):
        cNvPr = grpSp.find(f'.//{{{p_ns}}}cNvPr')
        if cNvPr is not None and '组合 26' in cNvPr.get('name', ''):
            grpSpPr = grpSp.find(f'{{{p_ns}}}grpSpPr')
            if grpSpPr is not None:
                xfrm = grpSpPr.find(f'{{{a}}}xfrm')
                if xfrm is not None:
                    off = xfrm.find(f'{{{a}}}off')
                    if off is not None:
                        off.set('x', str(decorator_x))

    # ============ 添加新的章节文本框 ============
    sp_id_counter = 50  # 从一个不冲突的 ID 开始

    for i in range(num_chapters):
        is_highlight = (i == ch_idx)
        chapter_sp = create_chapter_textbox(i, outline[i]['name'], is_highlight,
                                             chapter_positions[i], sp_id_counter)
        spTree.append(chapter_sp)
        sp_id_counter += 1

    # ============ 添加节行（紧凑右对齐） ============
    current_chapter = outline[ch_idx]
    num_sections = len(current_chapter['sections'])

    sec_x, sec_width = calculate_section_bar_params(num_sections)

    section_bar = create_section_bar(
        current_chapter['sections'], sec_idx, sec_x, sec_width, sp_id_counter)
    spTree.append(section_bar)
    sp_id_counter += 1

    # ============ 添加小节行（紧凑右对齐） ============
    current_section = current_chapter['sections'][sec_idx]
    num_subsections = len(current_section['subsections'])

    sub_x, sub_width = calculate_subsection_bar_params(num_subsections)

    subsection_bar = create_subsection_bar(
        current_section['subsections'], sub_idx, sub_x, sub_width, sp_id_counter)
    spTree.append(subsection_bar)
    sp_id_counter += 1

    return master_xml


def main():
    import os
    import shutil
    import zipfile

    base_dir = os.path.dirname(os.path.abspath(__file__))
    input_path = os.path.join(base_dir, 'mubu_response.json')
    template_path = os.path.join(base_dir, 'tpl', 'PPT模板.pptx')
    output_path = os.path.join(base_dir, 'tpl', 'PPT模板_generated.pptx')

    outline = parse_input(input_path)
    print(f"解析大纲完成：{len(outline)} 个章节")
    for i, chapter in enumerate(outline):
        print(f"  第{CN_NUMBERS[i]}章: {len(chapter['sections'])} 节")
        for j, section in enumerate(chapter['sections']):
            print(f"    第{j + 1}节: {len(section['subsections'])} 小节")

    # 用 python-pptx 打开模板（只用来生成 XML，不用来保存）
    prs = Presentation(template_path)
    source_master = prs.slide_masters[SOURCE_MASTER_INDEX]
    source_master_part = source_master.part

    total_masters = sum(
        len(section['subsections'])
        for chapter in outline
        for section in chapter['sections']
    )
    print(f"\n将生成 {total_masters} 个母版")

    # 读取原始母版的 rels 文件来获取关系映射
    with zipfile.ZipFile(template_path, 'r') as zin:
        source_rels_xml = zin.read(
            f'ppt/slideMasters/_rels/slideMaster{SOURCE_MASTER_INDEX + 1}.xml.rels').decode('utf-8')

    # 找出原始模板中已有多少个 slideMaster 和 slideLayout
    with zipfile.ZipFile(template_path, 'r') as zin:
        existing_masters = [n for n in zin.namelist()
                           if n.startswith('ppt/slideMasters/slideMaster') and n.endswith('.xml')]
        existing_layouts = [n for n in zin.namelist()
                           if n.startswith('ppt/slideLayouts/slideLayout') and n.endswith('.xml')]
    num_existing_masters = len(existing_masters)
    num_existing_layouts = len(existing_layouts)
    print(f"原始模板中有 {num_existing_masters} 个母版, {num_existing_layouts} 个布局")

    # 读取原始 presentation.xml 和 presentation.xml.rels 以及 [Content_Types].xml
    with zipfile.ZipFile(template_path, 'r') as zin:
        pres_xml = zin.read('ppt/presentation.xml')
        pres_rels_xml = zin.read('ppt/_rels/presentation.xml.rels')
        content_types_xml = zin.read('[Content_Types].xml')

    # 解析这些 XML
    pres_tree = etree.fromstring(pres_xml)
    pres_rels_tree = etree.fromstring(pres_rels_xml)
    ct_tree = etree.fromstring(content_types_xml)

    RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
    CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'

    # 找到已有的最大 rId
    max_rId = 0
    for rel_elem in pres_rels_tree:
        rid = rel_elem.get('Id', '')
        if rid.startswith('rId'):
            num = int(rid[3:])
            if num > max_rId:
                max_rId = num

    # 找到已有的最大 sldMasterId 和 sldLayoutId
    sldMasterIdLst = pres_tree.find(f'{{{NSMAP["p"]}}}sldMasterIdLst')
    max_master_id = 0
    max_layout_id = 0
    if sldMasterIdLst is not None:
        for elem in sldMasterIdLst:
            mid = int(elem.get('id', '0'))
            if mid > max_master_id:
                max_master_id = mid
            # 扫描该母版下的所有 layoutId
    if max_master_id == 0:
        max_master_id = 2147483647

    # 扫描已有的 sldLayoutId
    for master_id_lst in pres_tree.findall(f'.//{{{NSMAP["p"]}}}sldLayoutId'):
        lid = int(master_id_lst.get('id', '0'))
        if lid > max_layout_id:
            max_layout_id = lid
    if max_layout_id == 0:
        max_layout_id = 2147483900

    # 准备新文件：先复制原始模板
    shutil.copy2(template_path, output_path)

    # 生成所有新母版并写入 ZIP
    master_count = 0
    # (type, path, xml_bytes)
    new_files = []

    for ch_idx, chapter in enumerate(outline):
        for sec_idx, section in enumerate(chapter['sections']):
            for sub_idx, subsection in enumerate(section['subsections']):
                master_name = f'母版{ch_idx + 1}-{sec_idx + 1}-{sub_idx + 1}'
                print(f"  生成: {master_name}")

                # 生成母版 XML
                new_master_xml = create_master_for_subsection(
                    prs, source_master, outline, ch_idx, sec_idx, sub_idx)

                # ---- 修改母版 XML 中的 sldLayoutIdLst，只引用新布局 ----
                p_ns = NSMAP['p']
                sldLayoutIdLst_elem = new_master_xml.find(f'{{{p_ns}}}sldLayoutIdLst')
                if sldLayoutIdLst_elem is not None:
                    # 清除所有旧的布局引用
                    for child in list(sldLayoutIdLst_elem):
                        sldLayoutIdLst_elem.remove(child)
                    # 添加新的布局引用（rId1 指向新布局）
                    new_layout_id = max_layout_id + master_count + 1
                    sldLayoutId_elem = etree.SubElement(sldLayoutIdLst_elem,
                                                         f'{{{p_ns}}}sldLayoutId')
                    sldLayoutId_elem.set('id', str(new_layout_id))
                    sldLayoutId_elem.set(f'{{{NSMAP["r"]}}}id', 'rId1')

                # 序列化母版 XML
                master_xml_bytes = etree.tostring(new_master_xml, xml_declaration=True,
                                                   encoding='UTF-8', standalone=True)

                new_master_idx = num_existing_masters + master_count + 1
                new_layout_idx = num_existing_layouts + master_count + 1

                # 母版文件
                new_files.append((
                    f'ppt/slideMasters/slideMaster{new_master_idx}.xml',
                    master_xml_bytes
                ))

                # ---- 创建新布局 XML（命名为 master_name）----
                layout_xml = (
                    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                    f'<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
                    f' xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
                    f' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
                    f' type="blank" preserve="1">'
                    f'<p:cSld name="{master_name}">'
                    f'<p:spTree>'
                    f'<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
                    f'<p:grpSpPr/>'
                    f'</p:spTree>'
                    f'</p:cSld>'
                    f'</p:sldLayout>'
                ).encode('utf-8')
                new_files.append((
                    f'ppt/slideLayouts/slideLayout{new_layout_idx}.xml',
                    layout_xml
                ))

                # ---- 母版的 .rels 文件（rId1=新布局, rId3=image, rId4=theme）----
                master_rels_xml = (
                    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                    f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    f'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"'
                    f' Target="../slideLayouts/slideLayout{new_layout_idx}.xml"/>'
                    f'<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"'
                    f' Target="../media/image1.png"/>'
                    f'<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"'
                    f' Target="../theme/theme2.xml"/>'
                    f'</Relationships>'
                ).encode('utf-8')
                new_files.append((
                    f'ppt/slideMasters/_rels/slideMaster{new_master_idx}.xml.rels',
                    master_rels_xml
                ))

                # ---- 布局的 .rels 文件（指回母版）----
                layout_rels_xml = (
                    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                    f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    f'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"'
                    f' Target="../slideMasters/slideMaster{new_master_idx}.xml"/>'
                    f'</Relationships>'
                ).encode('utf-8')
                new_files.append((
                    f'ppt/slideLayouts/_rels/slideLayout{new_layout_idx}.xml.rels',
                    layout_rels_xml
                ))

                # ---- 更新 presentation.xml.rels ----
                new_rId = f'rId{max_rId + master_count + 1}'
                rel_elem = etree.SubElement(pres_rels_tree, f'{{{RELS_NS}}}Relationship')
                rel_elem.set('Id', new_rId)
                rel_elem.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster')
                rel_elem.set('Target', f'slideMasters/slideMaster{new_master_idx}.xml')

                # ---- 更新 presentation.xml sldMasterIdLst ----
                if sldMasterIdLst is None:
                    sldMasterIdLst = etree.SubElement(pres_tree, f'{{{NSMAP["p"]}}}sldMasterIdLst')

                new_mid = max_master_id + master_count + 1
                sldMasterId = etree.SubElement(sldMasterIdLst, f'{{{NSMAP["p"]}}}sldMasterId')
                sldMasterId.set('id', str(new_mid))
                sldMasterId.set(f'{{{NSMAP["r"]}}}id', new_rId)

                # ---- 更新 [Content_Types].xml ----
                # 母版
                override = etree.SubElement(ct_tree, f'{{{CT_NS}}}Override')
                override.set('PartName', f'/ppt/slideMasters/slideMaster{new_master_idx}.xml')
                override.set('ContentType',
                             'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml')
                # 布局
                override2 = etree.SubElement(ct_tree, f'{{{CT_NS}}}Override')
                override2.set('PartName', f'/ppt/slideLayouts/slideLayout{new_layout_idx}.xml')
                override2.set('ContentType',
                              'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml')

                master_count += 1

    # 写入所有新文件到 ZIP
    with zipfile.ZipFile(output_path, 'a') as zout:
        for path, data in new_files:
            zout.writestr(path, data)

    # 更新已有的文件（presentation.xml, rels, content_types）
    tmp_path = os.path.join(base_dir, 'tpl', 'PPT模板_tmp.pptx')

    updated_pres_xml = etree.tostring(pres_tree, xml_declaration=True,
                                       encoding='UTF-8', standalone=True)
    updated_pres_rels = etree.tostring(pres_rels_tree, xml_declaration=True,
                                        encoding='UTF-8', standalone=True)
    updated_ct_xml = etree.tostring(ct_tree, xml_declaration=True,
                                     encoding='UTF-8', standalone=True)

    with zipfile.ZipFile(output_path, 'r') as zin:
        with zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == 'ppt/presentation.xml':
                    zout.writestr(item, updated_pres_xml)
                elif item.filename == 'ppt/_rels/presentation.xml.rels':
                    zout.writestr(item, updated_pres_rels)
                elif item.filename == '[Content_Types].xml':
                    zout.writestr(item, updated_ct_xml)
                else:
                    zout.writestr(item, zin.read(item.filename))

    os.replace(tmp_path, output_path)

    print(f"\n保存成功: {output_path}")
    print(f"共生成 {master_count} 个新母版")


if __name__ == '__main__':
    main()

