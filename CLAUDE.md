# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目概述

将幕布（Mubu）大纲文档一键转换为定制样式的 PowerPoint 演示文稿。核心是一条 0~6 编号的顺序脚本流水线，配合 python-pptx / lxml 直接操作 OOXML 生成带联动导航条的 PPT；`gui_app.py` 用 PyWebView 提供本地图形界面，最终打包为 Windows exe（Nuitka，CI 构建）和 Mac app（PyInstaller，本地构建）。

## 常用命令

```bash
# 激活虚拟环境（脚本运行依赖 python-pptx / requests / pywebview / lxml）
source venv/bin/activate

# 一键跑完整流水线（0~6步），可传 shareId
python run_all.py <shareId>

# 单独运行某一步（每步都会读写 output/ 下的中间产物）
python 0-fetch_mubu.py <shareId>          # 拉取幕布分享文档 -> output/mubu_response.json
python 1-mubu_parser.py                    # 解析为章节结构 -> output/mubu_parsed_structure.json
python 2-generate_master_json.py           # 按章节生成母版信息 JSON -> output/master/*.json
python 3-create_master.py [--template tpl/PPT模板.pptx] [--input ...] [--output ...]  # 生成带导航条的 slideMaster -> output/master.pptx
python 4-generate_cover.py [--input_json ...] [--input_pptx ...] [--output_pptx ...]  # 生成封面页 -> output/master_with_cover.pptx
python 5-generate_toc.py                   # 生成总目录页 -> output/master_with_toc.pptx
python 6-generate_body.py                  # 生成章节页+内容页（自动分页）-> output/final_presentation.pptx

# 启动图形界面（开发模式，实际生产环境走打包后的 app/exe）
python gui_app.py

# Mac 本地打包（PyInstaller，产出 dist/大纲生成PPT.app 与 .zip）
./build_mac.sh
```

Windows exe 由 GitHub Actions（`.github/workflows/build-exe.yml`）在 push 到 master/main 时用 Nuitka 自动构建，产物作为 workflow artifact 上传，无需本地手动跑。

没有测试套件；`test_paths.py` 是一次性调试脚本，用于探测打包后（PyInstaller/Nuitka frozen 模式）`sys._MEIPASS`、`sys.executable`、`os.getcwd()` 等路径变量的实际值，验证路径修复逻辑用。

## 架构：流水线数据流

每一步都读写 `output/` 目录下的固定文件名作为交接点，脚本之间没有直接函数调用，全靠磁盘中间文件耦合：

```
0-fetch_mubu.py       shareId (CLI 参数) → output/mubu_response.json（幕布原始 API 响应）
1-mubu_parser.py       output/mubu_response.json → output/mubu_parsed_structure.json（递归解析出的 Course/Chapter 结构）
2-generate_master_json.py  output/mubu_parsed_structure.json → output/master/<code>.json（每个含内容的章节生成一份母版元信息）
3-create_master.py     tpl/PPT模板.pptx + output/master/*.json → output/master.pptx（按 JSON 驱动生成新 slideMaster，含右上角章节导航条）
4-generate_cover.py    output/master.pptx + mubu_parsed_structure.json → output/master_with_cover.pptx（替换第1页封面文字）
5-generate_toc.py      output/master_with_cover.pptx → output/master_with_toc.pptx（生成第2页总目录）
6-generate_body.py     output/master_with_toc.pptx → output/final_presentation.pptx（递归生成章节页与内容页，含自动分页/降字号/图片下载）
```

`gui_app.py` 中的 `Api._run_pipeline` 用 `importlib.import_module` + `reload` 动态串联这 0~6 个脚本（而不是走 subprocess），通过修改 `sys.argv` 模拟命令行参数，并把每一步的 `main()` 依次调用；执行前会 `os.chdir(self.work_dir)` 保证脚本内部基于 `os.getcwd()` 的相对路径（`output/`、`tpl/`）始终正确。

### 关键设计点

- **幕布大纲解析规则**（`1-mubu_parser.py`）：节点是否作为"内容页"而非"子章节"递归，由 `is_content` 判定（强制标记 `*` 前缀 / 含 HTML `<table>` / 含图片 / 文字长度 >16 / 层级达到4级 / 无子节点，任一命中即为内容）。文字节点中间嵌入的表格会被 `explode_sub_content` 按原始顺序拆分为独立的 text/table 内容块，保证表格与文字穿插时页面顺序正确。中英文标题统一用 `parse_bilingual` 按末尾括号切分（如 `中文名(English Name)`）。
- **母版生成**（`3-create_master.py`）：不是用 python-pptx 的高层 API 拼页面，而是直接用 `lxml` 操作 OOXML XML 树（`slideMaster2.xml` 等），按硬编码的 EMU 坐标槽位（`CHAPTER_SLOTS`、`SECTION_BAR_Y` 等，均从模板文件反推而来）重新生成一级/二级/三级章节导航条文本框，实现导航条随当前页面章节位置高亮联动。修改导航条样式/坐标需要同步理解 `tpl/PPT模板.pptx` 模板文件的原始结构。
- **正文自动分页**（`6-generate_body.py`）：`paginate_contents_custom` 先构建内容树估算行数（`get_lines` 用 16pt 假设，`get_lines_at_14pt` 用于超长时降级到 14pt 重新估算能否塞进单页），再据此决定分页与字号，避免同一段落被硬切到两页。
- **打包路径兼容**：`gui_app.py` 中 `Api.__init__` 需要处理三种运行形态的资源目录差异：普通 Python 运行、PyInstaller frozen 模式、以及 Mac `.app` 内部运行时因 macOS AppTranslocation 只读隔离区导致的写入失败（此时会退化为写到 `~/Downloads`）。改动路径逻辑前建议先跑 `test_paths.py` 或参考其历史输出理解各变量取值。
- **版本号单一来源**：`version.py` 的 `__version__`，GUI 标题栏与关于信息均从这里读取。

## 目录说明

- `tpl/PPT模板.pptx`：唯一的 PPT 设计模板底稿，`3-create_master.py` 从中提取 `slideMaster2` 并按 JSON 重排导航条。
- `tpl/gui/`：PyWebView 前端页面（HTML/CSS/JS），负责用户输入 shareId、展示流水线日志（`appendLog`）、任务完成通知（`taskCompleted`）。
- `tpl/img/`：封面/目录页用到的数字图（`2个字.png` ~ `9个字.png`）与 `logo.png`。
- `output/`：流水线中间产物与最终 PPT（`.gitignore` 排除），GUI 完成后会清理中间文件，仅保留 `output/<shareId>/<课程名>.pptx`。
