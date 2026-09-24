# 标准文档批量转 Markdown 工具

本工具把下面目录中的所有文件（含子文件夹）转换成 Markdown，并保留原目录结构：

```text
D:\chemical_design\文档\标准2(1)\标准2
```

输出目录默认为：

```text
D:\chemical_design\文档\标准2(1)\标准2_md
```

> 结论先说：**数字原生 PDF / DOCX 可以做到内容不经过 OCR，因此不会产生 OCR 错字；扫描件 OCR 无法保证 100% 无误。**  
> 本工具不会“假装”扫描件转换成功：所有 OCR 结果都会标成“需人工复核”，并生成逐页原图和复核清单。

---

## 一、文件说明

| 文件 | 作用 |
|---|---|
| `convert_to_md.py` | 主程序 |
| `requirements.txt` | 基础依赖 |
| `requirements_basic_ocr.txt` | RapidOCR 基础 OCR 兜底依赖 |
| `requirements_mineru.txt` | MinerU 依赖（扫描件高精度解析） |
| `run_converter.bat` | 一键安装基础依赖并转换（推荐） |
| `run_converter_basic_ocr.bat` | 没有 MinerU 时用 RapidOCR 兜底（结果必须人工复核） |
| `install_mineru.bat` | 安装/更新 MinerU |
| `标准2_md\_qa\qa_report.csv` | 每个文件的转换报告 |
| `标准2_md\_qa\review_queue.csv` | 需要人工复核的文件清单 |
| `标准2_md\_qa\review_index.html` | 可视化复核页面 |

---

## 二、环境准备

### 1. Windows

建议安装以下软件：

1. **Python 3.10 / 3.11 / 3.12（64 位）**  
   安装时勾选 `Add Python to PATH`。

2. **Pandoc（推荐，处理 DOCX 效果最好）**  
   有管理员权限时可用 winget：

   ```powershell
   winget install --id JohnMacFarlane.Pandoc -e
   ```

   没有 Pandoc 时程序会自动改用 Mammoth 或 python-docx。

3. **LibreOffice（只有旧版 `.doc` 需要）**  
   当前目录主要是 `.pdf` 和 `.docx`，所以通常不是必须的。  
   如果以后有 `.doc`，可安装 LibreOffice，或安装 Microsoft Word + `pywin32`。

4. **MinerU（扫描件强烈推荐）**  
   当前目录中明确有扫描件，例如：

   ```text
   GB 30871-2022 《危险化学品企业特殊作业安全规范》扫描件.pdf
   ```

   请先运行：

   ```bat
   install_mineru.bat
   ```

   MinerU 首次运行会下载模型，体积较大；CPU 可以运行，NVIDIA 显卡会更快。  
   官方项目：<https://github.com/opendatalab/MinerU>

---

## 三、运行

### 方式 A：推荐流程

双击：

```text
run_converter.bat
```

它会：

1. 创建 `.venv` 虚拟环境；
2. 安装 `requirements.txt`；
3. 调用 `convert_to_md.py`；
4. 所有 PDF 优先交给 MinerU 解析（支持公式、表格、版面识别）；
5. 生成 QA 报告。

### 方式 B：先用基础 OCR 跑一遍

如果暂时不想安装体积较大的 MinerU，可双击：

```text
run_converter_basic_ocr.bat
```

该方式使用 RapidOCR，输出到 `标准2_md_basic_ocr`，所有扫描件结果会标记为 `REVIEW_OCR`。  
它只适合“先有结果再核对”，不能作为最终准确版本。  
之后安装 MinerU 时，推荐重新运行 `run_converter.bat`，输出到独立的 `标准2_md`，不要让基础 OCR 结果和最终结果混在一起。

### 方式 C：命令行

```bat
cd /d "D:\chemical_design\文档\标准2(1)"
call .venv\Scripts\activate.bat
python convert_to_md.py ^
  --source "D:\chemical_design\文档\标准2(1)\标准2" ^
  --output "D:\chemical_design\文档\标准2(1)\标准2_md" ^
  --pdf-engine auto ^
  --strict
```

常用参数：

| 参数 | 说明 |
|---|---|
| `--pdf-engine auto` | 优先 MinerU（公式/表格/版面）；MinerU 不可用时数字版用 pymupdf4llm |
| `--pdf-engine pymupdf` | 只做数字文本提取，速度快，但公式可能无法识别 |
| `--pdf-engine mineru` | 所有 PDF 都走 MinerU |
| `--mineru-tier standard/advanced` | 指定 MinerU 档位；公式多时建议 `standard` 或 `advanced` |
| `--allow-basic-ocr` | 没有 MinerU 时允许 RapidOCR 兜底 |
| `--overwrite` | 覆盖已有输出 |
| `--no-resume` | 不跳过已有输出 |
| `--limit 3` | 只处理前 3 个文件，用于试跑 |
| `--dry-run` | 只列出文件，不转换 |
| `--list-engines` | 查看本机检测到的引擎 |
| `--strict` | 有需复核或失败文件时返回非 0，便于验收 |

---

## 四、准确性保障机制

### PDF

- `--pdf-engine auto` 时优先用 MinerU，公式会输出为 `$...$`、`$$...$$` 的 LaTeX；
- 如果只是想快速提取数字文本、不需要公式识别，可用 `--pdf-engine pymupdf`；
- 转换后把 Markdown 与原始文本层做“字符 2-gram 覆盖率”回比；
- 覆盖率低于 97% 的文件进入 `review_queue.csv`；
- 发现 `�`、`□`、私用区字符等异常字符的文件也会进入复核队列；
- MinerU 输出中的公式、表格、图片会尽量保留在 Markdown 中。

### DOCX

- 优先 Pandoc → GitHub Flavored Markdown；
- 自动用 python-docx 回比正文与表格文本；
- 没有 Pandoc 时回退 Mammoth / python-docx，并在报告中注明。

### 扫描件 PDF

- 优先 MinerU 做版面、表格、公式和 OCR；
- MinerU 结果依然标记 `REVIEW_OCR`，因为 OCR 没有 100% 保证；
- 没有 MinerU 时，默认拒绝输出扫描件，避免产生看似准确、实际有错字的 Markdown；
- 只有加 `--allow-basic-ocr` 才用 RapidOCR，且每页原图保存在 `*_assets` 目录中，方便对照。

### 为什么不能承诺“绝对无错字”

如果扫描件本身图像模糊、印章遮挡、字体特殊、公式/上下标密集，任何 OCR 都可能出错。  
本工具的目标是：

1. 数字文档：**不引入 OCR 错字**；
2. 扫描文档：用最好的开源方案解析，并**把所有不确定项暴露出来**；
3. 通过 `review_index.html` 让人工只复核真正需要复核的文件，而不是全部重看。

---

## 五、输出目录结构

例如源文件：

```text
标准2\危化\GB 15603-2022《危险化学品仓库储存通则》.pdf
```

输出为：

```text
标准2_md\危化\GB 15603-2022《危险化学品仓库储存通则》.md
标准2_md\危化\GB 15603-2022《危险化学品仓库储存通则》_assets\
标准2_md\_qa\qa_report.csv
标准2_md\_qa\review_queue.csv
标准2_md\_qa\review_index.html
标准2_md\_qa\convert_to_md.log
```

图片会放在对应的 `*_assets` 目录中，Markdown 中使用相对路径引用。

---

## 六、人工复核建议

1. 打开 `标准2_md\_qa\review_index.html`；
2. 所有红色行都需要复核；
3. 对照“原文件”和“Markdown”；
4. 扫描件重点检查：标准编号、数字、单位、上下标、表格行列、负号/短横线、容易混淆的汉字；
5. 复核完成后，可把 `review_queue.csv` 中该文件删除或标记；

---

## 七、MinerU 4.x 本地公式识别配置

如果 `mineru server status` 显示：

```text
本地 (disabled)  健康=否
远程             健康=是
```

说明解析请求可能走了远程 `https://mineru.net/api`，而不是本地公式识别模型。  
建议改为本地 managed 模式：

```powershell
# 1. 启用本地 managed parse server，标准档支持公式识别
mineru config set parse_server.local.mode managed

# 2. 中国大陆网络建议使用 ModelScope 下载模型
$env:MINERU_MODEL_SOURCE = "modelscope"

# 3. 重启 MinerU 服务，让本地 parse server 读取新配置和模型源
mineru server stop
mineru server start

# 4. 反复检查，直到本地状态健康、档位为 standard
mineru server status
```

`mineru server status` 里本地 Parse Server 健康后，再用：

```powershell
python convert_to_md.py --pdf-engine auto --mineru-tier standard --overwrite
```

MinerU 识别出的公式会以 LaTeX 形式写入 Markdown：

```text
行内公式：$...$
行间公式：$$...$$
```

6. 如果某个扫描件 MinerU 仍然错误较多，建议把该 PDF 单独用更高 DPI 重新 OCR，或使用专业版 OCR/人工校对。
