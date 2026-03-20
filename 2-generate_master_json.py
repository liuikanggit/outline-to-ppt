import json
import os

def load_data(path):
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)

def generate_master_json():
    base_dir = os.getcwd()
    input_path = os.path.join(base_dir, 'output', 'mubu_parsed_structure.json')
    output_dir = os.path.join(base_dir, 'output', 'master')
    
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # 清理旧文件以便重新生成
    for f in os.listdir(output_dir):
        if f.endswith('.json'):
            os.remove(os.path.join(output_dir, f))

    data = load_data(input_path)
    level1_chapters = data.get("chapters", [])
    
    # 提前构建 Level 1 导航列表 (带 en)
    l1_list = []
    for c in level1_chapters:
        l1_list.append({
            "zh": c.get("chapterName", ""),
            "en": c.get("chapterEnName", "")
        })

    def dfs(node, level, ancestors_lists, ancestors_indices):
        """
        node: 当前 Chapter 节点
        level: 当前层级 (1, 2, 3)
        ancestors_lists: [l1_list, l2_list, l3_list]
        ancestors_indices: [i1, i2, i3]
        """
        code = node.get("code", "")
        content = node.get("content", [])
        
        # 只要该章节下有内容，就生成一个 JSON 母版
        if content:
            master_info = {
                "masterName": f"母版-{code}", 
                "level1Chapter": {
                    "list": l1_list,
                    "activeIndex": ancestors_indices[0] if len(ancestors_indices) > 0 else 0
                }
            }
            if level >= 2:
                master_info["level2Chapter"] = {
                    "list": ancestors_lists[1],
                    "activeIndex": ancestors_indices[1]
                }
            if level >= 3:
                master_info["level3Chapter"] = {
                    "list": ancestors_lists[2],
                    "activeIndex": ancestors_indices[2]
                }
                
            out_file = os.path.join(output_dir, f"{code}.json")
            with open(out_file, 'w', encoding='utf-8') as f:
                json.dump(master_info, f, ensure_ascii=False, indent=2)
            # print(f"✅ 生成 {out_file}")

        # 递归子章节
        sub_chapters = node.get("subChapter", [])
        # 构建当前层级的 list 供下级使用 (按 ts 结构，仅需 zh)
        current_list = [{"zh": sc.get("chapterName", "")} for sc in sub_chapters]
        
        for idx, sub_node in enumerate(sub_chapters):
            next_lists = ancestors_lists + [current_list]
            next_indices = ancestors_indices + [idx]
            dfs(sub_node, level + 1, next_lists, next_indices)

    # 遍历顶层 Level 1
    for i, ch in enumerate(level1_chapters):
        dfs(ch, 1, [l1_list], [i])

    print(f"✅ 所有 Master JSON 已导出至 {output_dir}")

def main():
    generate_master_json()

if __name__ == "__main__":
    main()
