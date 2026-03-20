#!/usr/bin/env python3
"""
gui_app.py
==========
大纲生成 PPT 的 PyWebView 图形界面入口。
"""
import webview
import subprocess
import os
import sys
import json
import threading
import pptx  # 解决打包后找不到 pptx 模块的问题


# 全局窗体引用，用于调用 evaluate_js
window = None

class Api:
    def __init__(self, current_dir):
        import sys
        
        # 1. 静态资源目录（HTML & 脚本本身）: 打包后在 _MEIPASS，开发时在当前目录
        self.asset_dir = getattr(sys, '_MEIPASS', current_dir)
        
        # 2. 物理输出目录（用户运行程序的真实位置）
        if getattr(sys, 'frozen', False):
            # 如果是 Mac .app 内部运行，需要向上爬 3 级才能拿到外部包含文件夹
            if sys.platform == 'darwin':
                # dist/大纲生成PPT.app/Contents/MacOS/大纲生成PPT
                self.work_dir = os.path.abspath(os.path.join(os.path.dirname(sys.executable), '..', '..', '..'))
            else:
                self.work_dir = os.path.dirname(sys.executable)
        else:
            self.work_dir = current_dir

        print(f"📦 静态资源目录 (asset_dir): {self.asset_dir}")
        print(f"📂 工作输出目录 (work_dir): {self.work_dir}")

    def generate_ppt(self, share_id):
        """前端点击生成按钮时触发"""
        print(f"收到请求，开始处理 ShareID: {share_id}")
        
        # 启动后台耗时任务线程
        import threading
        threading.Thread(target=self._run_pipeline, args=(share_id,), daemon=True).start()
        
        # 立即给前端返回成功接收状态，避免阻塞前端UI
        return {"status": "success", "message": "任务已启动"}

    def _run_pipeline(self, share_id):
        """后台线程执行 0~6 步骤脚本并清理后置文件"""
        scripts = [
            "0-fetch_mubu",
            "1-mubu_parser",
            "2-generate_master_json",
            "3-create_master",
            "4-generate_cover",
            "5-generate_toc",
            "6-generate_body"
        ]
        
        import io
        import sys
        import json
        import importlib

        class LogRedirector:
            def __init__(self, js_call):
                self.js_call = js_call
            def write(self, s):
                if s and s.strip():
                    self.js_call(f'appendLog({json.dumps(s.strip())})')
            def flush(self):
                pass

        old_stdout = sys.stdout
        old_stderr = sys.stderr
        sys.stdout = LogRedirector(self._js_call)
        sys.stderr = sys.stdout

        # 确保能搜寻到静态资源目录（_MEIPASS 等）下的脚本
        if self.asset_dir not in sys.path:
            sys.path.insert(0, self.asset_dir)

        try:
            for script_name in scripts:
                self._js_call(f'appendLog(">>> 正在执行模块: {script_name}")')
                
                # 模拟 sys.argv 参数以适配各脚本内部加载或 argparse
                if script_name == "0-fetch_mubu":
                    sys.argv = [script_name + ".py", share_id]
                elif script_name == "3-create_master":
                    template_path = os.path.join(self.asset_dir, "tpl", "PPT模板.pptx")
                    sys.argv = [script_name + ".py", "--template", template_path]
                else:
                    sys.argv = [script_name + ".py"]

                # ⚠️ 动态修正脚本中原本可能强行找 __file__ 获取 CWD 的习惯，

                # 其实更好的方式是让 0-6 里面也能适配。我们先在调用前，设置好os.chdir
                os.chdir(self.work_dir)

                # 动态导入模块
                module = importlib.import_module(script_name)
                # 使用 reload 强制让模块代码全量再次执行一遍、刷新全局变量
                importlib.reload(module)

                # 如果模块有 main() 入口，调用它以触发执行
                if hasattr(module, 'main'):
                     module.main()

            # 全部完成，执行清理
            self._cleanup_temp_files(share_id)
            
            self._js_call('appendLog("✅ 所有步骤执行完毕！准备进入结果页...")')
            self._js_call('taskCompleted(true, "成功")')
            
        except Exception as e:
            self._js_call(f'appendLog("❌ 产生异常: {str(e)}")')
            self._js_call('taskCompleted(false, "异常导致中断")')
        finally:
            sys.stdout = old_stdout
            sys.stderr = old_stderr

    def _cleanup_temp_files(self, share_id):
        """清理过程中间文件，将最终PPT移动到单独的 share_id 文件夹，其余全删"""
        self._js_call('appendLog("🧹 开始清理系统缓存与整理输出...")')
        import shutil
        out_dir = os.path.join(self.work_dir, "output")
        final_dir = os.path.join(out_dir, share_id)
        
        # 1. 创建 share_id 文件夹
        if not os.path.exists(final_dir):
            os.makedirs(final_dir, exist_ok=True)
            
        # 2. 将成功生成的 final_presentation.pptx 移动到该文件夹下
        source_ppt = os.path.join(out_dir, "final_presentation.pptx")
        dest_ppt = os.path.join(final_dir, "final_presentation.pptx")
        if os.path.exists(source_ppt):
            shutil.move(source_ppt, dest_ppt)
            self._js_call(f'appendLog("  - 📁 最终 PPT 妥善存放在: output/{share_id}/final_presentation.pptx")')
            
        # 3. 清理掉原来的零碎中间文件
        targets_to_delete = [
            "mubu_response.json",
            "mubu_parsed_structure.json",
            "master.pptx",
            "master_with_cover.pptx",
            "master_with_toc.pptx"
        ]
        
        for t in targets_to_delete:
            p = os.path.join(out_dir, t)
            if os.path.exists(p):
                os.remove(p)

        # 清除 master 图片集和 json
        master_dir = os.path.join(out_dir, "master")
        if os.path.exists(master_dir):
            shutil.rmtree(master_dir)
            
        self._js_call('appendLog("✨ 清理完毕！存放文件夹已保持最简纯净状态。")')

    def _js_call(self, js_string):
        """安全向前端推送 JS 脚本"""
        global window
        if window:
            window.evaluate_js(js_string)

    def open_output_folder(self, share_id):
        """打开输出目录的对应文件夹"""
        output_dir = os.path.join(self.work_dir, "output", share_id)
        if not os.path.exists(output_dir):
             output_dir = os.path.join(self.work_dir, "output")
             
        if sys.platform == "darwin":
            os.system(f'open "{output_dir}"')
        elif sys.platform == "win32":
             os.startfile(output_dir)
        else:
            subprocess.run(["xdg-open", output_dir])

    def open_final_ppt(self, share_id):
        """打开最终生成的 PPT"""
        ppt_path = os.path.join(self.work_dir, "output", share_id, "final_presentation.pptx")
        if os.path.exists(ppt_path):
            if sys.platform == "darwin":
                os.system(f'open "{ppt_path}"')
            elif sys.platform == "win32":
                os.startfile(ppt_path)
            else:
                subprocess.run(["xdg-open", ppt_path])
        else:
            self._js_call('appendLog("⚠️ 未找到最终 PPT 文件，请先点击生成。")')

def main():
    global window
    current_dir = os.path.dirname(os.path.abspath(__file__))
    api = Api(current_dir)

    # 载入面板 HTML
    html_path = os.path.join(current_dir, "tpl", "gui", "gui_index.html")

    window = webview.create_window(
        title='大纲一键生成 PPT 工具',
        url=f'file://{html_path}' if sys.platform != 'win32' else html_path,
        js_api=api,
        width=850,
        height=700,
        resizable=True
    )

    webview.start(debug=False)

if __name__ == '__main__':
    main()
