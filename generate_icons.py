import os
import subprocess
import sys

def create_ico(src, dest):
    print("Creating .ico for Windows...")
    try:
        from PIL import Image
    except ImportError:
        print("Pillow not found. Installing Pillow...")
        subprocess.run([sys.executable, "-m", "pip", "install", "Pillow"], check=True)
        from PIL import Image

    img = Image.open(src)
    # 尽可能包含所有的全尺寸图标
    sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
    img.save(dest, format='ICO', sizes=sizes)
    print(f"✅ 生成 .ico 成功: {dest}")

def create_icns(src, dest):
    print("Creating .icns for Mac...")
    from PIL import Image
    iconset_dir = "app_icon.iconset"
    
    # 清理掉之前失败留下的残留
    import shutil
    if os.path.exists(iconset_dir):
        shutil.rmtree(iconset_dir)
    os.makedirs(iconset_dir, exist_ok=True)
    
    img = Image.open(src)
    sizes = [
        (16, "icon_16x16.png"), (32, "icon_16x16@2x.png"),
        (32, "icon_32x32.png"), (64, "icon_32x32@2x.png"),
        (128, "icon_128x128.png"), (256, "icon_128x128@2x.png"),
        (256, "icon_256x256.png"), (512, "icon_256x256@2x.png"),
        (512, "icon_512x512.png"), (1024, "icon_512x512@2x.png")
    ]
    
    for size, name in sizes:
        out_path = os.path.join(iconset_dir, name)
        # 用 Pillow 进行高质量抗锯齿缩放，防止 iconutil 报错
        resized_img = img.resize((size, size), Image.Resampling.LANCZOS)
        # 强制转换为 RGBA，消除一切 profile 阻断
        if resized_img.mode != 'RGBA':
            resized_img = resized_img.convert('RGBA')
        resized_img.save(out_path, "PNG")
        
    # 调用 iconutil 合并为 .icns
    subprocess.run(["iconutil", "-c", "icns", iconset_dir, "-o", dest], check=True)
    
    # 清理中间文件夹
    shutil.rmtree(iconset_dir)
    print(f"✅ 生成 .icns 成功: {dest}")


if __name__ == "__main__":
    src_img = "app_icon_source.png"
    if not os.path.exists(src_img):
        print(f"❌ 找不到源文件: {src_img}")
        sys.exit(1)
        
    create_ico(src_img, "app_icon.ico")
    create_icns(src_img, "app_icon.icns")
