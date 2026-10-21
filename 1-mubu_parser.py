import json
import os
import re
import html
from html.parser import HTMLParser

# ----------------- 基础工具函数 (复用) -----------------
def format_node(node):
    """去除幕布节点文本中的 HTML 标签"""
    text = node.get('text', '')
    text = re.sub(r'<[^>]+>', '', text)
    text = html.unescape(text)
    return text.strip()

def format_note(node):
    """提取节点的附注(note)并清理 HTML"""
    text = node.get('note', '')
    text = re.sub(r'<[^>]+>', '', text)
    text = html.unescape(text)
    return text.strip()

class HTMLTableParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.rows = []
        self.current_row = []
        self.current_cell = ""
        self.in_cell = False
        self.is_header = False
        self.headers = []

    def handle_starttag(self, tag, attrs):
        if tag == 'tr':
            self.current_row = []
        elif tag in ['td', 'th']:
            self.in_cell = True
            self.current_cell = ""
            if tag == 'th': self.is_header = True

    def handle_data(self, data):
        if self.in_cell:
            self.current_cell += data

    def handle_endtag(self, tag):
        if tag in ['td', 'th']:
            self.in_cell = False
            self.current_row.append(self.current_cell.strip())
            if self.is_header:
                self.is_header = False
        elif tag == 'tr':
            if self.current_row:
                self.rows.append(self.current_row)

def parse_bilingual(text):
    """精确切分中文与英文"""
    text = text.strip()
    if text.endswith(')') or text.endswith('）'):
        idx_half = text.rfind('(')
        idx_full = text.rfind('（')
        idx = max(idx_half, idx_full)
        if idx > 0:
            cn = text[:idx].strip()
            en = text[idx+1:-1].strip()
            return cn, en
    return text, ""

# ----------------- v2 新增递归工具 -----------------
def build_sub_content(node):
    """
    递归构建子内容列表，支持文字、表格、图片混合。
    返回列表中每项要么是文字节点（hasBullet=True），
    要么是表格/图片标记（__type='table' / 'image'），供上层拆分。
    """
    sub_list = []
    for child in node.get('children', []):
        raw_text = child.get('text', '')
        has_table = '<table' in raw_text
        if has_table:
            parser = HTMLTableParser()
            parser.feed(raw_text)
            if parser.rows:
                sub_list.append({'__type': 'table', 'rows': parser.rows})
        else:
            text = format_node(child)
            if text.startswith('*'):
                text = text[1:].strip()
            cn, _ = parse_bilingual(text)
            # 图片挂在该节点文字之后、子节点之前，保持幕布中的视觉顺序
            images = [
                {'__type': 'image', 'image': {'id': img.get('id', ''), 'uri': img.get('uri', '')}}
                for img in child.get('images', []) if img.get('uri')
            ]
            if cn:
                sub_list.append({
                    'text': cn,
                    'hasBullet': True,
                    'subContent': images + build_sub_content(child)
                })
            elif images:
                # 纯图片节点（无文字）：图片和其子内容直接提升到当前层
                sub_list.extend(images)
                sub_list.extend(build_sub_content(child))
    return sub_list


def flush_text_group(chapter, title, title_en, index_path, text_group):
    """将积累的文字节点写入 chapter content（内部工具函数）"""
    if not text_group:
        return
    # 过滤掉 subContent 中的表格/图片标记，只保留纯文字子项
    def clean_sub(items):
        return [
            {**item, 'subContent': clean_sub(item.get('subContent', []))}
            for item in items if not item.get('__type')
        ]
    cleaned = [
        {**item, 'subContent': clean_sub(item.get('subContent', []))}
        for item in text_group
    ]
    if chapter['content'] and chapter['content'][-1]['type'] == 'text':
        chapter['content'][-1]['text'].extend(cleaned)
    else:
        chapter['content'].append({
            'title': title,
            'titleEn': title_en,
            'index': index_path,
            'type': 'text',
            'text': cleaned
        })


def split_item_by_blocks(item):
    """
    递归拆分一个文字项：表格/图片可能嵌在 subContent 的任意深度（如 4、5 级以后），
    按原始顺序返回分段列表，每段为 ('text', 文字项)、('table', rows) 或 ('image', 图片信息)。
    表格之后剩余的文字段会保留祖先文字作为上下文，保证续页语义完整。
    """
    segments = []
    pending_sub = []
    emitted_text = False

    def flush():
        nonlocal pending_sub, emitted_text
        # 表格前没有新文字、且本项文字已输出过时，不再重复输出空壳标题
        if pending_sub or not emitted_text:
            segments.append(('text', {**item, 'subContent': pending_sub}))
            emitted_text = True
        pending_sub = []

    for child in item.get('subContent', []):
        if child.get('__type') == 'table':
            flush()
            segments.append(('table', child['rows']))
            continue
        if child.get('__type') == 'image':
            flush()
            segments.append(('image', child['image']))
            continue
        for seg_type, seg in split_item_by_blocks(child):
            if seg_type == 'text':
                pending_sub.append(seg)
            else:
                flush()
                segments.append((seg_type, seg))

    if pending_sub or not emitted_text:
        flush()
    return segments


def explode_sub_content(chapter, title, title_en, index_path, top_text_item):
    """
    将一个顶层文字项（含混合 subContent）拆分为：
    若干文字块 + 表格块 + 图片块，按原始顺序写入 chapter content。
    表格/图片可嵌在任意深度，由 split_item_by_blocks 递归拆出。
    """
    for seg_type, seg in split_item_by_blocks(top_text_item):
        if seg_type == 'text':
            flush_text_group(chapter, title, title_en, index_path, [seg])
        elif seg_type == 'image':
            chapter['content'].append({
                'title': title,
                'titleEn': title_en,
                'index': index_path,
                'type': 'image',
                'image': seg
            })
        else:
            chapter['content'].append({
                'title': title,
                'titleEn': title_en,
                'index': index_path,
                'type': 'table',
                'table': {'rows': seg}
            })

def recursive_parse_v2(node, level, code_prefix, parent_cn, parent_en, index_path):
    """
    根据用户 TS 结构递归解析 Mubu 节点
    node: 当前节点
    level: 当前层级 1, 2, 3...
    code_prefix: 编号前缀, 如 '1' 或 '1.2'
    index_path: 数组形式的索引路径，如 [0, 1]
    """
    text = format_node(node)
    # 处理当前节点自身的星号（如果它作为子章节进入）
    if text.startswith('*'):
        text = text[1:].strip()
        
    cn, en = parse_bilingual(text)
    if not en: 
        en = parent_en # 继承上级英文名(如有需要)
    note = format_note(node)

    chapter = {
        "code": code_prefix,
        "level": level,
        "chapterName": cn,
        "chapterEnName": en,
        "note": note,
        "content": [],
        "subChapter": []
    }

    children = node.get('children', [])
    for i, child in enumerate(children):
        raw_text = child.get('text', '')
        c_text = format_node(child)
        
        # 核心逻辑：检测并处理强制内容标记 '*'
        force_content = False
        if c_text.startswith('*'):
            force_content = True
            c_text = c_text[1:].strip()
            
        c_cn, c_en = parse_bilingual(c_text)
        
        c_images = child.get('images', [])
        c_has_image = len(c_images) > 0
        
        # [NEW] 检测 HTML 表格
        has_table = '<table' in raw_text
        
        c_code = f"{code_prefix}.{i+1}"
        # 当前子节点的索引路径
        c_index_path = index_path + [i]

        # 核心判定规则更新：增加 force_content 和 has_table
        # 表格节点一定是内容页面
        is_content = force_content or has_table or c_has_image or (len(c_text) > 16) or (level + 1 >= 4) or (not child.get('children', []))

        if is_content:
            if has_table:
                # 解析 HTML 表格
                parser = HTMLTableParser()
                parser.feed(raw_text)
                if parser.rows:
                    chapter["content"].append({
                        "title": cn,
                        "titleEn": en,
                        "index": c_index_path,
                        "type": "table",
                        "table": {
                            "rows": parser.rows
                        }
                    })
                    # 处理完表格后，不再处理图片或文字，直接跳到下一个子节点
                    continue

            if c_has_image:
                # A. 图片内容: 根据 TS 结构, imageContent 在 content 里
                for img in c_images:
                    chapter["content"].append({
                        "title": cn, # 上级中文名
                        "titleEn": en,
                        "index": c_index_path,
                        "type": "image",
                        "image": {
                            "id": img.get("id", ""),
                            "uri": img.get("uri", "")
                        }
                    })
            elif c_cn:
                # B. 文字内容：递归子列表可能含表格，需拆分
                new_text_item = {
                    "text": c_cn,
                    "hasBullet": False,
                    "subContent": build_sub_content(child)
                }
                explode_sub_content(chapter, cn, en, c_index_path, new_text_item)
        else:
            # B. 判定为子章节, 递归
            sub_ch = recursive_parse_v2(child, level + 1, c_code, cn, en, c_index_path)
            chapter["subChapter"].append(sub_ch)

    return chapter


def load_mubu_data_v2(filepath):
    """主入口: 依照 Course 模型读取幕布数据并导出"""
    if not os.path.exists(filepath):
        print(f"找不到文件: {filepath}")
        return None

    with open(filepath, 'r', encoding='utf-8') as f:
        mubu_data = json.load(f)

    data_obj = mubu_data.get('data', {})
    summary = data_obj.get('summary', '')
    course_name, course_en_name = parse_bilingual(summary)

    # 找到 nodes
    content = data_obj.get('content', {})
    definition_str = content.get('definition', '{}')
    definition = json.loads(definition_str)
    nodes = definition.get('nodes', [])

    chapters = []
    for i, top_node in enumerate(nodes):
        # 强制将顶层一级节点作为大章 Chapter 排布
        # 顶层节点 index 表示为 [i]
        ch = recursive_parse_v2(top_node, 1, str(i + 1), course_name, course_en_name, [i])
        chapters.append(ch)

    course = {
        "courseName": course_name,
        "courseEnName": course_en_name,
        "chapters": chapters
    }
    return course

def main():
    base_dir = os.getcwd()
    input_path = os.path.join(base_dir, 'output','mubu_response.json')
    output_path = os.path.join(base_dir, 'output', 'mubu_parsed_structure.json')

    print("开始生成 v2版 级联大纲结构...")
    course_data = load_mubu_data_v2(input_path)
    
    if course_data:
        try:
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(course_data, f, ensure_ascii=False, indent=2)
            print(f"成功！v2版 新大纲 JSON 存盘至: {output_path}")
        except Exception as e:
            print(f"保存失败: {e}")

if __name__ == '__main__':
    main()
