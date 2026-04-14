# PDF Extractor 操作手册（TLG）

## 1. 这是什么工具

这个工具用来把 PDF 里的表格提取出来，并保存成 Excel 文件（`.xlsx`）。

你可以处理两类 PDF：

- **文字型 PDF**（可以直接选中文字）
- **扫描型 PDF**（本质是图片）

输出结果说明：

- 每个识别到的表格会放在一个独立工作表里
- 同时会生成一个 `_summary` 工作表，记录每个表格来自第几页、使用了哪种识别方式

最适合的场景：

- 发票、报表、清单这类有明显表格结构的 PDF

不擅长的场景：

- 非表格排版（纯段落）
- 图片质量很差、倾斜严重、遮挡严重的扫描件
- 非常复杂的跨页表格或手写内容

完整 CLI 示例（按需删减参数）：

```bash
python pdf_table_extractor.py "./pdf_folder" --output-dir "./outcome_folder" --pages "1-3,5" --mode img2table --prefer both --ocr-lang "chi_sim+eng" --ocr-lang-auto --borderless --img2table-min-confidence 35 --excel-style-mode basic --row-compact --row-compact-empty-ratio 0.8 --row-compact-header-rows 5 --min-rows 2 --min-cols 2 --min-filled-ratio 0.15 --accuracy-threshold 50 --verbose
```

> 上面是覆盖大多数参数的一条完整版命令。README 未提到的命令可忽略，这里仅作完整参考。

---

## 2. 使用前说明

> **先读完这部分再运行。**

1. **不要改脚本文件名和路径**：
   - `pdf_table_extractor.py`
   - `requirements.txt`
   - `training/` 目录下所有脚本

2. 不要自行修改高级参数（例如 `--min-rows`、`--min-cols`、`--min-filled-ratio`、`--accuracy-threshold`）。
   - 需要调整时，请直接联系 Haochen。

3. 任何文件替换前先备份：
   - 原始 PDF
   - 旧版输出 Excel

4. 如果终端提示缺少依赖或 OCR 引擎：
   - 不要手动删库
   - 先按本手册“常见报错”处理
   - 仍无法解决就联系 Haochen

---

## 3. 你只需要知道的文件和目录

- `pdf_table_extractor.py`
  - 主程序入口，**可以运行，不要编辑**。

- `requirements.txt`
  - 依赖清单，安装环境会用到，**不要编辑**。

- `README.md`
  - 项目说明文档，**可以阅读，不建议随意改**。

- 你的输入 PDF 所在目录
  - 可自行新建，例如 `mypdf/`。

---

## 4. 运行前准备

### 4.1 确认 Python 已安装

在 PowerShell 输入：

```bash
python --version
```

如果看到类似 `Python 3.x.x`，说明已安装。若报错（例如找不到 python），请先安装 Python。

### 4.2 进入项目目录

```bash
cd /workspace/pdf_extractor_excel_ver2
```

### 4.3 安装依赖

首次使用（或环境重装后）执行：

```bash
pip install -r requirements.txt
```

### 4.4 检查 OCR 引擎（扫描 PDF 需要）

如果要处理扫描件（图片型 PDF），还需要安装 Tesseract。

检查命令：

```bash
tesseract --version
```

- 能显示版本号：可用
- 报错找不到命令：说明未安装或 PATH 未配置好，可先执行：

```bash
winget install --id UB-Mannheim.TesseractOCR
```

### 4.5 查看命令帮助

```bash
python pdf_table_extractor.py --help
```

---

## 5. 最常用的操作

### 命令 1：单个 PDF 提取

**用途**：提取一个 PDF 的表格到 Excel。  
**什么时候用**：一次只处理 1 个文件。

```bash
python pdf_table_extractor.py "你的PDF所在路径"
```

示例：

```bash
python pdf_table_extractor.py "./invoice.pdf"
```

运行后：

- 成功时会看到类似：`[OK] invoice.pdf: saved X table(s) to ...`

结果位置：

- 默认保存在同目录，文件名类似：`invoice_tables.xlsx`

---

### 命令 2：单个 PDF + 指定输出文件名

**用途**：把结果保存为你指定的 Excel 名称。

```bash
python pdf_table_extractor.py "你的PDF路径" -o "你的输出xlsx路径"
```

示例：

```bash
python pdf_table_extractor.py "./invoice.pdf" -o "./result/haochen.xlsx"
```

运行后：

- 成功时会显示：`[OK] ... saved ... to 你指定的路径`

---

### 命令 3：批量提取（整个文件夹）

**用途**：一次处理一个文件夹里的多个 PDF。  
**什么时候用**：每天需要处理一批文件。

```bash
python pdf_table_extractor.py "PDF文件夹路径"
```

示例：

```bash
python pdf_table_extractor.py "./mypdfs"
```

参数说明：

- 输入必须是目录，目录内放 `.pdf` 文件

运行后：

- 每个 PDF 都会输出一行 `[OK]` 或 `[FAILED]`

结果位置：

- 默认在输入目录下新建：`extracted_tables/`

---

### 命令 4：批量提取 + 指定输出目录

**用途**：将批量结果统一保存到指定目录。  
**什么时候用**：希望结果集中放在一个固定位置。

```bash
python pdf_table_extractor.py "PDF文件夹路径" --output-dir "输出目录"
```

示例：

```bash
python pdf_table_extractor.py "./daily_pdfs" --output-dir "./haochen"
```

参数说明：

- `--output-dir` 可用于单文件或批量模式

---

## 6. 其他常用命令

### 6.1 只提取指定页

```bash
python pdf_table_extractor.py "你的PDF路径" --pages "页码规则"
```

示例：

```bash
python pdf_table_extractor.py "./report.pdf" --pages "1-3,5"
```

### 6.2 切换提取模式（文字型 / 扫描型）

```bash
python pdf_table_extractor.py "你的PDF路径" --mode auto
```

- 自动选择（可选模式：先尝试文字方式，必要时走 OCR）
- **默认模式是 `camelot`（文字型 PDF 优先）**

```bash
python pdf_table_extractor.py "你的PDF路径" --mode camelot
```

- 文字型提取

```bash
python pdf_table_extractor.py "你的PDF路径" --mode pdfplumber
```

- 另一种文字型引擎

```bash
python pdf_table_extractor.py "你的PDF路径" --mode img2table
```

- 扫描件提取
