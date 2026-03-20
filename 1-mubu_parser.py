import json
import os
import re
import html

# ----------------- 基础工具函数 (复用) -----------------
def format_node(node):
    """去除幕布节点文本中的 HTML 标签"""
    text = node.get('text', '')
    # 去除包含 <span> 标签等 HTML
    text = re.sub(r'<[^>]+>', '', text)
    # 解码 HTML 实体, 如 &nbsp;
    text = html.unescape(text)
    return text.strip()

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
    """递归构建 TextItem: { text: '...', hasBullet: True, subContent: [] }"""
    sub_list = []
    for child in node.get('children', []):
        text = format_node(child)
        cn, _ = parse_bilingual(text)
        if cn: # 排除纯空格节点
            sub_list.append({
                "text": cn,
                "hasBullet": True,
                "subContent": build_sub_content(child)
            })
    return sub_list

def recursive_parse_v2(node, level, code_prefix, parent_cn, parent_en, index_path):
    """
    根据用户 TS 结构递归解析 Mubu 节点
    node: 当前节点
    level: 当前层级 1, 2, 3...
    code_prefix: 编号前缀, 如 '1' 或 '1.2'
    index_path: 数组形式的索引路径，如 [0, 1]
    """
    text = format_node(node)
    cn, en = parse_bilingual(text)
    if not en: 
        en = parent_en # 继承上级英文名(如有需要)

    chapter = {
        "code": code_prefix,
        "level": level,
        "chapterName": cn,
        "chapterEnName": en,
        "content": [],
        "subChapter": []
    }

    children = node.get('children', [])
    for i, child in enumerate(children):
        c_text = format_node(child)
        c_cn, c_en = parse_bilingual(c_text)
        
        c_images = child.get('images', [])
        c_has_image = len(c_images) > 0
        
        c_code = f"{code_prefix}.{i+1}"
        # 当前子节点的索引路径
        c_index_path = index_path + [i]

        # 核心判定规则：任何节点只要 【带有图片】 或 【文字长度 > 9】 或 【包含冒号】 或 【当前章节层级 >= 3】, 均判定为本章节的 Content 
        is_content = c_has_image or (len(c_cn) > 9) or (":" in c_cn) or ("：" in c_cn) or (level >= 3)

        if is_content:
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
            # 如果它同时包含文字 (虽然 Mubu 里图片节点往往是文字节点的附属, 但也要防空)
            if c_cn:
                new_text_item = {
                    "text": c_cn,
                    "hasBullet": False,
                    "subContent": build_sub_content(child) # 子文本后代沉降
                }
                
                # 检查最后一个 content 是否是 text 类型，如果是，则聚合
                if chapter["content"] and chapter["content"][-1]["type"] == "text":
                    chapter["content"][-1]["text"].append(new_text_item)
                else:
                    chapter["content"].append({
                        "title": cn,
                        "titleEn": en,
                        "index": c_index_path,
                        "type": "text",
                        "text": [new_text_item]
                    })
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
