\# PDF Extractor 操作手册



\## =============================1. 这是什么工具=============================

这个工具用来把 PDF 里的表格提取出来，保存成 Excel 文件。  

你可以处理两类 PDF：

\- //文字型 PDF//（可以直接选中文字）

\- //扫描型 PDF//（本质是图片）



输出是 `.xlsx` 文件。每个识别到的表格会放在一个工作表里，同时会有一个 `\_summary` 工作表，记录每个表格来自第几页、用了哪种识别方式。



它最适合：

\- 发票、报表、清单这类“有明显表格结构”的 PDF。



它不擅长：

\- 非表格排版（纯段落）-- WORKING ON THIS

\- 图片质量很差、倾斜严重、遮挡严重的扫描件

\- 非常复杂的跨页表格或手写内容  



One-line-for-all:

`python pdf\_table\_extractor.py "./pdf\_folder" --output-dir "./outcome\_folder" --pages "1-3,5" --mode img2table --prefer both --ocr-lang "chi\_sim+eng" --ocr-lang-auto --borderless --img2table-min-confidence 35 --excel-style-mode basic --row-compact --row-compact-empty-ratio 0.8 --row-compact-header-rows 5 --min-rows 2 --min-cols 2 --min-filled-ratio 0.15 --accuracy-threshold 50 --verbose`



以上是所有CLI都涉及到了的一条完整版命令，可以根据需求删减。README下面没提到的命令可以直接无视，这里只是提供一个完整的CLI



\## =============================2. 使用前说明=============================

> //先读完这部分再运行。//



1\. //不要改脚本文件名和路径//：

&#x20;  - `pdf\_table\_extractor.py`

&#x20;  - `requirements.txt`

&#x20;  - `training/` 目录下所有脚本



2\. 不要自己改高级参数（例如 `--min-rows`、`--min-cols`、`--min-filled-ratio`、`--accuracy-threshold`）。  

&#x20;  需要调整时，直接联系 Haochen。



3\. 任何文件替换前，先备份：

&#x20;  - 你的原始 PDF

&#x20;  - 旧版输出 Excel



4\. 如果终端提示缺少依赖或 OCR 引擎：

&#x20;  - 不要手动删库；

&#x20;  - 按本手册“常见报错”处理；

&#x20;  - 还不行就联系Haochen。



\##============================= 4. 你只需要知道的文件和目录=============================

\- `pdf\_table\_extractor.py`  

&#x20; 主程序入口。//可以运行，不要编辑//。



\- `requirements.txt`  

&#x20; 依赖清单。安装环境会用到。//不要编辑//。



\- `README.md`  

&#x20; 这份操作说明。//可以阅读，不建议随意改//。



\- 你的输入 PDF 所在目录（你自己新建即可，eg `mypdf`）  



\## =============================5. 运行前准备=============================

\### 5.1 先确认电脑里有 Python

在powershell输入：

`python --version`

如果看到类似 `Python 3.x.x`，说明已安装。  

如果报错（例如找不到 python），先安装 Python 再继续。



\### 5.2 进入项目目录

先打开终端，再进入项目文件夹，例如：

`cd /workspace/pdf\_extractor\_excel\_ver2`



\### 5.3 安装依赖

首次使用（或环境重装后）在powershell执行：

`pip install -r requirements.txt`



\### 5.4 检查 OCR 引擎（处理扫描 PDF 时需要）

如果你要处理扫描件（图片型 PDF），还需要系统里有 Tesseract。  

检查命令：

`tesseract --version`

\- 能显示版本号：可用。  

\- 报错找不到命令：说明没装或没配好 PATH，需要先安装/配置：

`winget install --id UB-Mannheim.TesseractOCR`



\### 5.5 查看可用命令帮助

`python pdf\_table\_extractor.py --help`



\## =============================6. 最常用的操作=============================



**### 命令 1：单个 PDF 提取**

//用途//  

提取一个 PDF 的表格到 Excel。



//什么时候用//  

一次只处理 1 个文件。



//命令//

`python pdf\_table\_extractor.py "你的PDF所在路径"`



示例：

`python pdf\_table\_extractor.py "./invoice.pdf"`



//运行后会看到什么//

\- 成功时看到类似：`\[OK] invoice.pdf: saved X table(s) to ...`



//结果保存在哪里//

\- 默认保存在同目录，文件名类似：`invoice\_tables.xlsx`



\---



**### 命令 2：单个 PDF + 指定输出文件名**

//用途//  

把结果保存为你指定的 Excel 名称。



//用途//  

把结果保存为你指定的 Excel 名称。



//命令//

`python pdf\_table\_extractor.py "你的PDF路径" -o "你的输出xlsx路径"`



示例：

python pdf\_table\_extractor.py "./invoice.pdf" -o "./result/haochen.xlsx"



//运行后会看到什么//

\- 成功时会显示 `\[OK] ... saved ... to 你指定的路径`



\---



**### 命令 3：批量提取（整个文件夹）**

//用途//  

一次处理一个文件夹里的多个 PDF。



//什么时候用//  

每天要处理一批文件。



//命令//

`python pdf\_table\_extractor.py "PDF文件夹路径"`



示例：

`python pdf\_table\_extractor.py "./mypdfs"`



//参数说明//

\- 输入必须是目录，目录里放 `.pdf` 文件。



//运行后会看到什么//

\- 每个 PDF 都会出现一行 `\[OK]` 或 `\[FAILED]`。



//结果保存在哪里//

\- 默认在输入目录下新建：`extracted\_tables/`



\---



**### 命令 4：批量提取 + 指定输出目录**

//用途//  

批量结果统一保存到指定目录。



//什么时候用//  

你希望把结果集中放到一个固定位置。



//命令//

`python pdf\_table\_extractor.py "PDF文件夹路径" --output-dir "输出目录"`



示例：

`python pdf\_table\_extractor.py "./daily\_pdfs" --output-dir "./haochen"`



//参数说明//

\- `--output-dir` 可用于单文件或批量模式。



**###一些其他命令**

//只提取指定页//

`python pdf\_table\_extractor.py "你的PDF路径" --pages"页码规则"`

示例：

`python pdf\_table\_extractor.py "./report.pdf" --pages "1-3,5"`



//切换提取模式（文字型 / 扫描型）//

`python pdf\_table\_extractor.py "你的PDF路径" --mode auto`  #自动选择（默认，先用文字方式，必要时走 OCR）

`python pdf\_table\_extractor.py "你的PDF路径" --mode camelot` #text-based

`python pdf\_table\_extractor.py "你的PDF路径" --mode pdfplumber` #another text-based engine

`python pdf\_table\_extractor.py "你的PDF路径" --mode img2table` #scanned pdf





