# CSV 与 Excel 完整流程（第五阶段 5B）

本项目使用虚构模拟订单学习。5B 把已经完成的读取、字段映射、严格校验、指标计算和 Excel 报表连接起来，没有新增指标、GUI、数据库或第三方依赖。

## 1. 运行命令

在 VS Code 中打开本项目目录，使用已有虚拟环境。无需激活环境或更改 PowerShell 执行策略。

沿用原来的 CSV 使用方式：

```powershell
.\.venv\Scripts\python.exe -B -X utf8 report_metrics.py
.\.venv\Scripts\python.exe -B -X utf8 report_excel.py
```

默认仍读取 `data/sample_orders.csv`。指标命令只打印结果；报表命令默认在 `output` 下生成新的三工作表报表，显示完整路径。不需要对 CSV 确认字段映射。

指定自己新建的模拟 CSV：

```powershell
.\.venv\Scripts\python.exe -B -X utf8 report_excel.py "C:\你的临时目录\模拟订单.csv" --output "output\模拟CSV报表.xlsx"
```

读取自己新建的普通 Excel 订单：

```powershell
.\.venv\Scripts\python.exe -B -X utf8 report_metrics.py "C:\你的临时目录\模拟订单.xlsx"
.\.venv\Scripts\python.exe -B -X utf8 report_excel.py "C:\你的临时目录\模拟订单.xlsx" --output "output\模拟Excel报表.xlsx"
```

将示例中的输入路径替换为实际模拟文件。`--output` 必须是尚不存在的 `.xlsx` 路径；重复运行同一个指定名字会拒绝覆盖。省略 `--output` 则沿用日期、时间及随机编号生成新名称。输入与输出不能是同一个已存在的文件。

**已有运营报表是汇总结果，不能作为订单模板重新导入。** 输入仍需具备 5A 定义的全部 14 个订单明细字段。

## 2. Excel 交互与失败结果

1. 读取文件一次，显示内存快照的 SHA256 和工作表列表。
2. 明确选择工作表序号和表头原始行号；不默认使用活动页。
3. 查看每列候选。唯一候选可以回车接受；额外列必须明确映射或输入 `-` 忽略；重复或缺失字段不会被静默修复。
4. 对“实付金额”等有歧义的字段核实人民币元、明细分摊及业务含义，必要时输入 `明细`。
5. 查看最终映射，输入 `确认` 后才调用原有整份业务校验器。
6. 只有全部数据通过，指标命令才打印正式指标，报表命令才开始构建工作簿。

选择过程中输入 `q`、最终不输入 `确认`、交互中断或没有标准输入时，操作取消，返回码 `2`。没有跳过确认的参数。脚本 API 遇到取消会抛出 `InputCancelled`，不会用空列表代替数据继续计算。

Excel 读取、映射或校验失败返回码 `1`，显示问题位置和原因，不计算部分合格数据。普通报表保存失败也返回 `1`，不打印“报表已生成”。程序遇到未知异常会明确报告异常类型并失败退出；API 的异常链仍保留，不自动修复或重试。

源文件、原有报表均不修改。输出仍经过原有内存序列化回读、临时文件写入与刷盘、Windows 无覆盖发布流程。突然强制终止或断电仍可能留下 `.tmp`；发布完成后终端中断也可能已有完整文件，不能仅凭缺少成功消息断定文件不存在。

## 3. 为什么没有重复计算规则

```text
report_metrics.py / report_excel.py
  → metrics.calculate_metrics(path)
    → order_loader.load_orders(path)
      CSV：原 load_and_validate() → validate_rows()
      XLSX：inspect_and_validate() → 原 5A 读取、映射、确认 → validate_selection() → validate_rows()
    → 原有指标公式（整份校验通过后）
  → 原有控制台展示 / excel_report 的三工作表和安全保存
```

`inspect_excel.py` 只是把原交互过程提取为可调用函数；单独运行它仍仅检查数据，不计算指标或生成报表。`field_mapping.py`、`excel_reader.py`、CSV 解析与业务校验器均未改动。

`ValidatedOrders` 保存校验后的不可变明细与来源；`OrderSource` 保存文件名、格式，以及 Excel 工作表、表头行、快照 SHA256。它们是内部数据传递结构，不是数据库，也不是防范恶意 Python 调用者的权限系统。

`MetricsReport` 附带来源信息；来源不参与业务指标值的相等比较，所以同一批数据通过 CSV 和 Excel 得到相同指标，同时仍可分别检查各自来源。没有添加新的金额或退款公式，金额继续使用整数分，客单价继续使用精确分数，仅展示时四舍五入。

## 4. 报表内容与来源

保留三个工作表：

- **运营总览**：原有九项计数、金额与客单价，人民币元、两位小数。
- **商品销量排行**：按净成交件数降序，并列按商品编号字符串升序，不重新排序或换用其他公式。
- **数据校验摘要**：保留原有计数、口径和生成时间，追加“输入文件格式”。Excel 另记录实际工作表、表头原始行号、输入快照 SHA256 和快照说明。

CSV 原有摘要中的“输入 CSV 文件名”和时间单元格位置保留；Excel 使用“输入 Excel 文件名”。文件名、工作表名等仍强制写为文本，即使以 `=` 开头也不作为公式执行。

若确认前磁盘上的 Excel 被另一个程序替换，本轮仍使用已展示 SHA256 对应的内存快照，报表记录该快照。它不证明后来替换的新文件也通过了校验；需要重新运行并确认才能检查新版本。程序只读取一次，不因导出再次询问或重新读取源 Excel。

模拟样本的验收数字保持不变：80 条明细、40 个唯一订单；支付 3724.57 元、退款 900.86 元、净额 2823.71 元；支付订单 32、净有效订单 28；支付客单价 116.39 元、净客单价 100.85 元；净成交件数 116。

## 5. 自动测试和环境

```powershell
.\.venv\Scripts\python.exe -B -X utf8 -m unittest discover -v
.\.venv\Scripts\python.exe -B -X utf8 -m unittest -v test_order_pipeline
```

原有 205 项保留，新增 36 项集成测试，总计 241 项。新测试全部在自动清理的系统临时目录生成模拟输入副本和输出，不在 `data` 或正式 `output` 留下报表。测试包括完整真实子进程交互、CSV/Excel 指标一致、三工作表、来源、排序、取消、字段缺失和歧义、非法金额、日期时区与尾随内容、公式、同名保护、导出失败、临时文件清理、统一校验入口调用及原文件哈希不变。

Python 3.13.2、openpyxl 3.1.5、et-xmlfile 2.0.0，均沿用项目现有 `.venv`。没有安装新依赖。CSV 校验和指标路径仍能只用标准库运行；Excel 输入和 Excel 导出需要现有 openpyxl。

## 6. 本轮 GitHub 简要调查

核查日期：2026-10-18。通过 GitHub 公共接口查看仓库状态、固定提交、完整许可、依赖及实际格式适配源码，不仅依据 README。没有运行候选项目或其测试，不能据此声称 Windows/Python 3.13 的实际兼容性、安全性或全部测试结果。

| 候选 | 已核实许可、维护与依赖 | 复用判断 |
| --- | --- | --- |
| [jazzband/tablib](https://github.com/jazzband/tablib) | MIT；未归档；核查提交 `966876c76a6378557c1951badc1665b4d1375875`，最近提交 2026-10-06；Python ≥3.10，xlsx 可选依赖 openpyxl ≥2.6.0 | 有清晰格式适配层，但默认使用活动工作表和 `data_only=True`，不符合本项目的明确选择与公式拒绝要求；已有 openpyxl 足够，无需额外封装库 |
| [django-import-export/django-import-export](https://github.com/django-import-export/django-import-export) | BSD-2-Clause；未归档；核查提交 `e54d04a206b2dfdf4a3c426920ee22d7c090aa3b`，最近提交北京时间 2026-10-16；依赖 Django ≥5.2、Tablib ≥3.7.0、diff-match-patch | 有导入/导出适配接口，但引入 Web 框架与数据库相关层成本过高；XLSX 入口同样默认活动表和公式缓存，不直接采用 |

证据链接：[Tablib 许可](https://github.com/jazzband/tablib/blob/966876c76a6378557c1951badc1665b4d1375875/LICENSE)、[依赖](https://github.com/jazzband/tablib/blob/966876c76a6378557c1951badc1665b4d1375875/pyproject.toml)、[XLSX 实现](https://github.com/jazzband/tablib/blob/966876c76a6378557c1951badc1665b4d1375875/src/tablib/formats/_xlsx.py)；[django-import-export 许可](https://github.com/django-import-export/django-import-export/blob/e54d04a206b2dfdf4a3c426920ee22d7c090aa3b/LICENSE)、[依赖](https://github.com/django-import-export/django-import-export/blob/e54d04a206b2dfdf4a3c426920ee22d7c090aa3b/pyproject.toml)、[格式实现](https://github.com/django-import-export/django-import-export/blob/e54d04a206b2dfdf4a3c426920ee22d7c090aa3b/import_export/formats/base_formats.py)。

许可证在遵守保留声明等条件下允许相应复用，但本轮只参考格式适配的通用思路，没有复制候选代码或引入框架。复用的是本项目已验证的模块和已安装 openpyxl。

## 7. 已知限制

- Excel 仍只有普通 `.xlsx`、单工作表、单行表头；每次需要交互，没有批处理免确认、GUI、映射保存或自动平台适配。
- 原有严格字段、状态、渠道、商品价格一致性及退款件数规则不变，不能直接保证接入任意真实电商平台。
- 日期安全检查仍覆盖整本工作簿；未选择页的非法或不支持的原生 ISO 日期也会拒绝。数值日期仍有毫秒解析限制，不能恢复 Excel 已丢失的编号、金额或时间精度。
- 未全面验证恶意 XML、压缩包资源耗尽、超大数据及所有办公软件版本；仅处理自己创建、来源可信的模拟文件。
- 导出发布仍仅支持 Windows，并受原有 Excel 15 位数值精度保护约束。生成的是静态报表，筛选表格不会重新计算总览。
- 没有新增图表、AI 分析、平台连接、数据库或自动提交。完成开发后等待验收，不自动进入下一阶段。
