import zipfile
import re
from lxml import etree

nsmap = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
}

def inspect():
    path = "output/out_3.2.4.pptx"
    try:
        with zipfile.ZipFile(path, "r") as zin:
            masters = [p for p in zin.namelist() if "slideMasters/slideMaster" in p]
            masters.sort()
            if not masters:
                print("No slide masters found.")
                return

            last_master = masters[-1]
            print(f"Reading from: {last_master}")
            
            xml = zin.read(last_master)
            root = etree.fromstring(xml)

            for sp in root.findall(".//p:sp", nsmap):
                cNvPr = sp.find(".//p:cNvPr", nsmap)
                name = cNvPr.get("name") if cNvPr is not None else ""
                
                if name in ["lv2_bar", "lv3_bar"]:
                    print(f"\n--- {name} ---")
                    
                    # 获取 xfrm
                    xfrm = sp.find(".//a:xfrm", nsmap)
                    if xfrm is not None:
                        off = xfrm.find("a:off", nsmap)
                        ext = xfrm.find("a:ext", nsmap)
                        print(f"Position: x={off.get('x')}, y={off.get('y')}")
                        print(f"Size: cx={ext.get('cx')}, cy={ext.get('cy')}")
                    
                    # 获取 文字内容
                    text_runs = []
                    for t in sp.findall(".//a:t", nsmap):
                        if t.text:
                            text_runs.append(t.text)
                    full_text = "".join(text_runs)
                    print(f"Text String: {full_text}")
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    inspect()
