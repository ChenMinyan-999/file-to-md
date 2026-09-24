# 标准文档批量转 Markdown

基于 Python、PyMuPDF、MinerU 的标准/法律文档批量转换工具。

## 功能

- 递归扫描源目录，保留子目录结构
- 每个源文件单独生成一个 Markdown，不合并
- 数字原生 PDF：PyMuPDF / pymupdf4llm 文本提取
- 扫描件 PDF：MinerU 本地 managed 解析，支持表格、版面、OCR、公式
- DOCX：Pandoc / Mammoth / python-docx 多级回退
- 公式输出为 LaTeX：`$...$`、`$$...$$`
- 生成 QA 报告：
  - `qa_report.csv`
  - `review_queue.csv`
  - `review_index.html`

## 目录

```text
convert_to_md.py              主程序
requirements.txt              基础依赖
requirements_mineru.txt       MinerU 依赖
README.md                     本文件
README_转换说明.md             详细说明
部署与运行流程.md              部署与每次运行命令
run_converter.bat             一键转换
install_mineru.bat            安装 MinerU
setup_mineru_local.bat        下载模型并启用本地 managed 模式
```

## 快速使用

详细部署步骤见：

- [部署与运行流程.md](部署与运行流程.md)
- [README_转换说明.md](README_转换说明.md)

最简转换命令：

```powershell
python convert_to_md.py `
  --source "D:\path\to\标准2" `
  --output "D:\path\to\标准2_md" `
  --pdf-engine auto `
  --mineru-tier standard `
  --overwrite
```

## 注意

- 不要提交 `.venv/`
- 不要提交源 PDF / DOC / DOCX
- 不要提交转换后的 Markdown 和 `_assets`
- 仓库中只保留程序、配置和说明文档
- 具体忽略规则见 `.gitignore`

## 许可

未指定，如需开源请补充 LICENSE。
