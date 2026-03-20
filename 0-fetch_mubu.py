import json
import urllib.request
import os

def fetch_mubu_data(share_id: str, password: str = "", output_path: str = "mubu_response.json") -> bool:
    """
    通过 Mubu API 获取分享文档的内容
    """
    url = "https://api2.mubu.com/v3/api/document/share/get"
    headers = {
        "Content-Type": "application/json",
        "version": "3.0.0"
    }
    data = {
        "shareId": share_id,
        "password": password
    }
    
    req = urllib.request.Request(
        url, 
        data=json.dumps(data).encode('utf-8'), 
        headers=headers, 
        method='POST'
    )
    
    try:
        with urllib.request.urlopen(req, timeout=10) as response:
            if response.status == 200:
                resp_body = response.read().decode('utf-8')
                resp_json = json.loads(resp_body)
                
                if resp_json.get('code') == 0:
                    with open(output_path, 'w', encoding='utf-8') as f:
                        json.dump(resp_json, f, ensure_ascii=False, indent=2)
                    print(f"✅ 成功获取数据，并保存到: {output_path}")
                    return True
                else:
                    print(f"❌ 接口返回错误: {resp_json.get('msg')}")
                    return False
            else:
                print(f"❌ HTTP 请求失败，状态码: {response.status}")
                return False
    except Exception as e:
        print(f"❌ 请求发生异常: {e}")
        return False

def main():
    import sys
    base_dir = os.getcwd()
    output_file = os.path.join(base_dir, "output/mubu_response.json")
    
    # 确保输出目录存在，防止 open() 报错
    os.makedirs(os.path.dirname(output_file), exist_ok=True)

    
    # 默认 shareId
    target_share_id = "9bmCYuBg_R"
    if len(sys.argv) > 1:
        target_share_id = sys.argv[1]
    
    print(f"开始抓取 shareId: {target_share_id} ...")
    success = fetch_mubu_data(target_share_id, "", output_file)
    if success:
        print("提取文档流程结束，下一步可调用 mubu_parser.py 转换数据结构。")

if __name__ == "__main__":
    main()
