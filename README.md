# Excel 文件子表清理和内容校验脚本说明

这个目录提供两个 Python 脚本，用于批量处理 Excel 文件：

- `delete_extra_sheets.py`：删除多余工作表，只保留名称包含指定关键字的工作表。
- `verify_remote_sheet.py`：比对处理前后保留的工作表内容，确认内容没有被修改。

当前支持 `.xls` 和 `.xlsx`。默认保留关键字为 `远程漏洞`。默认输入、输出目录均基于脚本所在目录。

## 更新日志

### 2026-06-14

- 扩展支持文件类型：从仅支持 `.xls` 升级为同时支持 `.xls` 和 `.xlsx`。
- 统一清理与校验规则：保留所有名称包含关键字的工作表，删除所有不匹配的工作表；如果匹配到多个工作表，会全部保留。
- 新增公共模块 `excel_common.py`，集中维护文件收集、关键字校验、Excel COM 初始化、工作表匹配和 CSV 日志写入逻辑。
- 新增 `--dry-run` 参数，可预览每个文件将保留和删除的工作表，不生成处理后的 Excel 文件。
- 增加覆盖保护：默认不覆盖 `output` 目录中的同名文件，需要覆盖时必须显式添加 `--overwrite`。
- 增强处理日志：记录运行时间、扩展名、关键字、dry-run 状态、覆盖状态、原始工作表数量、匹配数量、删除数量、工作表列表、输出路径和异常原因。
- 增强比对日志：记录原始工作表、期望保留工作表、实际输出工作表、工作表数量和差异详情。
- 新增 `requirements.txt`，用于声明 Windows 环境下的 `pywin32` 依赖。
- 优化校验脚本：当没有同名文件需要比对时，不再额外启动 Excel COM。

## 目录结构

建议保持如下结构：

```text
脚本目录
├── delete_extra_sheets.py
├── verify_remote_sheet.py
├── excel_common.py
├── README.md
├── excel
│   └── 待处理的 .xls 或 .xlsx 文件
└── output
    ├── 处理后的 .xls 或 .xlsx 文件
    ├── 处理日志_*.csv
    └── 比对日志_*.csv
```

如果 `excel` 或 `output` 目录不存在，可以手动创建。脚本会自动创建 `output` 目录。

## 环境要求

脚本使用本机 Excel COM 自动化处理 Excel 文件，适用于 Windows 环境。

需要：

- Windows 系统
- 已安装 Microsoft Excel
- 已安装 Python
- Python 已安装 `pywin32`

安装依赖：

```powershell
python -m pip install pywin32
```

## 统一匹配规则

清理脚本和校验脚本共用同一条规则：

- 保留所有名称包含关键字的工作表。
- 删除所有名称不包含关键字的工作表。
- 如果一个文件中匹配到多个工作表，会全部保留。
- 如果一个文件中没有匹配工作表，清理脚本会跳过该文件并写入日志。

默认关键字为：

```text
远程漏洞
```

## 子表清理脚本

默认用法：

```powershell
python .\delete_extra_sheets.py
```

默认行为：

- 从脚本目录下的 `excel` 目录读取 `.xls` 和 `.xlsx` 文件。
- 只保留工作表名称包含 `远程漏洞` 的工作表。
- 删除其他工作表。
- 将处理后的同名文件保存到 `output` 目录。
- 原始文件不会被修改。
- 默认不会覆盖 `output` 目录中已有的同名文件。
- 在 `output` 目录生成 `处理日志_*.csv` 处理日志。

预演模式，不保存输出文件：

```powershell
python .\delete_extra_sheets.py --dry-run
```

允许覆盖输出目录中的同名文件：

```powershell
python .\delete_extra_sheets.py --overwrite
```

指定目录：

```powershell
python .\delete_extra_sheets.py --source "D:\input_excel" --output "D:\cleaned_excel"
```

处理单个文件：

```powershell
python .\delete_extra_sheets.py --source "D:\input_excel\sample.xlsx" --output "D:\cleaned_excel"
```

指定保留关键字：

```powershell
python .\delete_extra_sheets.py --keyword "远程漏洞"
```

也可以保留其他名称的工作表，例如：

```powershell
python .\delete_extra_sheets.py --keyword "Remote Vulnerabilities"
```

## 内容校验脚本

处理完成后，可以运行校验脚本确认保留工作表内容是否一致：

```powershell
python .\verify_remote_sheet.py
```

默认会比对：

- `excel` 目录中的原始 `.xls` 和 `.xlsx` 文件
- `output` 目录中的处理后 `.xls` 和 `.xlsx` 文件

校验内容包括：

- 原始文件和输出文件是否按文件名一一对应。
- 输出文件是否只包含原始文件中名称匹配关键字的工作表。
- 保留工作表的 `UsedRange` 范围是否一致。
- 保留工作表的单元格值是否一致。
- 能正常读取公式时，也会比对公式是否一致。

指定目录和关键字：

```powershell
python .\verify_remote_sheet.py --source "D:\input_excel" --output "D:\cleaned_excel" --keyword "远程漏洞"
```

校验完成后会在输出目录生成：

```text
比对日志_*.csv
```

日志中的 `status` 含义：

- `一致`：保留工作表内容一致，且输出文件的工作表列表符合清理规则。
- `异常`：存在差异，例如缺少输出文件、输出文件多余、工作表列表不符合规则、单元格值不同等。
- `失败`：文件打开或读取失败，需要查看 `message` 字段。

## 日志增强

处理日志会记录：

- 运行时间
- 文件扩展名
- 是否 dry-run
- 是否允许覆盖
- 使用的关键字
- 原始工作表数量
- 匹配工作表数量
- 删除工作表数量
- 原始工作表列表
- 保留工作表列表
- 删除工作表列表
- 输出文件路径
- 异常或跳过原因

比对日志会记录：

- 运行时间
- 文件扩展名
- 使用的关键字
- 原始工作表数量
- 输出工作表数量
- 期望保留的工作表列表
- 实际输出的工作表列表
- 差异详情

CSV 日志使用 `utf-8-sig` 编码，方便直接用 Excel 打开查看中文。

## 注意事项

- 当前脚本支持 `.xls` 和 `.xlsx`，不处理 `.xlsm`、`.csv` 等其他格式。
- Excel 临时文件，例如 `~$xxx.xlsx`，会被自动跳过。
- 输出文件与原始文件同名，但保存位置不同。
- 默认不会覆盖输出目录中的旧结果；需要覆盖时请显式添加 `--overwrite`。
- 建议正式处理前先运行 `--dry-run` 预览将保留和删除的工作表。
- 不建议在脚本运行时手动打开正在处理的 Excel 文件。
- 分享脚本时建议只分享 `.py` 文件和 `README.md`，不要附带个人数据文件、处理结果或日志文件。

## 常见问题

### 缺少 win32com 或 pywin32

说明当前 Python 环境不能调用 Excel COM。

处理方式：

```powershell
python -m pip install pywin32
```

### Excel.Application 启动失败

可能原因：

- 本机未安装 Microsoft Excel。
- Excel 后台进程异常。
- 文件正被其他程序占用。

可以先关闭所有 Excel 窗口，再到任务管理器中确认没有残留的 `EXCEL.EXE` 进程，然后重新运行脚本。

### 输出文件大小变化

这是正常情况。删除子表并由 Excel 重新保存后，文件体积可能变化，不代表保留工作表内容被修改。

是否修改内容，应以 `verify_remote_sheet.py` 的比对结果为准。

### 为什么使用 Excel COM

`.xls` 是旧版 Excel 格式，`.xlsx` 也可能包含样式、合并单元格、公式和其他复杂信息。当前脚本通过本机 Excel 打开并另存文件，可以尽量保持原文件格式稳定。
