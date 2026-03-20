#!/usr/bin/env python3
"""
run_all.py
==========
一键联动 0~6 步骤，全自动化生成最终 PPT！
"""
import subprocess
import sys
import os

def run_script(cmd_list):
    print(f"\n🚀 正在执行: {' '.join(cmd_list)}")
    try:
        # 直接输出到控制台，方便看进度
        res = subprocess.run(cmd_list, capture_output=False, text=True)
        if res.returncode != 0:
            print(f"❌ 脚本执行失败: {' '.join(cmd_list)}")
            sys.exit(1)
    except Exception as e:
        print(f"❌ 运行报错: {e}")
        sys.exit(1)

def main():
    # 1. 获取 shareId (支持传参)
    share_id = "9bmCYuBg_R"  # 默认值
    if len(sys.argv) > 1:
        share_id = sys.argv[1]
    
    print(f"=== 🌟 一键生成 PPT 自动化流程 (ShareId: {share_id}) ===")

    current_dir = os.path.dirname(os.path.abspath(__file__))
    python_bin = os.path.join(current_dir, "venv", "bin", "python")
    
    scripts = [
        ["0-fetch_mubu.py", share_id],
        ["1-mubu_parser.py"],
        ["2-generate_master_json.py"],
        ["3-create_master.py"],
        ["4-generate_cover.py"],
        ["5-generate_toc.py"],
        ["6-generate_body.py"]
    ]

    for item in scripts:
        script_name = item[0]
        args = item[1:]
        cmd = [python_bin, script_name] + args
        run_script(cmd)

    print("\n✅ 所有步骤执行完毕！最终成品参见: output/final_presentation.pptx")

if __name__ == "__main__":
    main()
