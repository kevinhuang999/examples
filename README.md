# 实验室自动化示例 / Lab Automation Examples

[中文](#中文) | [English](#english)

---

<a id="中文"></a>
# 中文

配合 CSDN 系列文章的**可运行代码**。每个目录是一个独立示例，拷走就能跑，不依赖本仓库其他部分。

| 目录 | 内容 | 依赖 |
| --- | --- | --- |
| [`20260909/`](20260909/) | 按 Excel 模板批量生成检验报告（保留合并单元格、命名区域、结论公式、签字栏） | openpyxl |
| [`20260910/`](20260910/) | 按 Word 模板批量生成检验报告书（模板只改三处，保留页眉页脚与表格样式） | docxtpl |
| [`20260911/`](20260911/) | 多台仪器导出的 Excel 自动合并汇总（表头行自动对齐、批号前导零保全、文字结果不混进均值） | pandas |
| [`20260912/`](20260912/) | 仪器导出的检验数据批量生成 PDF 检验报告（中文字体注册、跨页表头重复、单元格特殊字符转义） | reportlab |
| [`20260913/`](20260913/) | 稳定性考察数据按批次×时间点自动汇总（pivot 转宽表、检项按药典体例排行、超标格标红、空值留白） | pandas |
| [`20260914/`](20260914/) | 一个样品一份 Word 检验报告（模板占位符跨 run 替换、明细表按检项数动态长行、文件名非法字符处理） | python-docx |
| [`20260915/`](20260915/) | 中心实验室样本接收登记表（扫码明细与温控记录按运单号对齐、温度区间两头判、异常分三档结论并整行刷色） | openpyxl |
| [`20260916/`](20260916/) | 中心实验室采血 Kit 配置单自动生成（方案用量 × 各中心例数算套数、按 FEFO 排批次、按套展开标签清单、缺料与冲突单独成页） | openpyxl |
| [`20260917/`](20260917/) | 多中心检验报告按中心批量分发（各中心命名模板改名、按受试者分包、算 TAT、编号与原件与撞名三类异常单独成页） | openpyxl |
| [`20260918/`](20260918/) | 样本流转交接记录串成台账（按环节规则串链、缺环/时间倒挂/同环节重复/停留超时分别挑出、抄丢前导零的编号单列） | openpyxl |
| [`20260919/`](20260919/) | 按访视窗与检项清单拼访视时点报告（提前到访也判超窗、未到访与未到期分档、按访视点清单判缺项、方案外项目不丢） | openpyxl |
| [`20260920/`](20260920/) | 一份数据按各中心不同的模板出报告（按字段标签找格子不写死坐标、标签归一化、单位制按中心换算并同步改单位列、模板不要的字段不硬塞） | openpyxl |
| [`20260921/`](20260921/) | 存量纸质样本接收单批量转线上（字段按来源分三类、纸面手填数量与明细逐单核对、挂起单一份不出、线上单号按天+中心编流水） | openpyxl |
| [`20260922/`](20260922/) | 中心实验室 TAT 超时提醒（按起算口径选起点、扣停表时段、工作日口径跨调休逐格算、分五档出提醒名单与催办文本） | openpyxl |
| [`20260923/`](20260923/) | 第三方检测委托单批量生成（受理登记表按日期+客户归堆、日期与客户名归一化后分组、六类缺项整单挂起不占号、按标签反查模板格子出单） | openpyxl |
| [`20260924/`](20260924/) | 第三方检测报价单批量汇总与历史价比对（价目表长表按方法+生效期取价、项目名归一化与别名表匹配、折扣按能报项数定档、挂起项单列不硬报） | openpyxl |
| [`20260925/`](20260925/) | 第三方检测报告按客户命名规则批量改名分发（命名模板占位符改名、按客户拆包只复制不动原件、撞名/未登记收件人/原件缺失/待审核/已作废五类挂起、发放登记表按报告×收件人展开） | openpyxl |
| [`20260926/`](20260926/) | 送检样品登记与编号自动生成（编号按日期+客户代码+类别码分段发号、同客户当天第二批接着排不重开、一行按数量逐件展开为 -1…-N、编不出号整行挂起且不占号、台账+标签清单同源） | openpyxl |
| [`20260927/`](20260927/) | 第三方检测报告批量用印（按报告类型查用章规则逐枚落章、首页/末页/右侧骑缝三种落位、骑缝章一半放到页外并按骑缝页序阶梯下移、章图透明掩码不能丢、五类不够格的整份挂起、台账按报告×章登记） | openpyxl + pypdf |
| [`20260928/`](20260928/) | 检测机构开票申请与月度关账对账（申请单对报告台账/已开票台账/客户规则表逐单判定：抬头税号、报告未出、必须PO、金额不符、重复开票、框架合同不含税额度；作废票不占额度也不挡重开；出可开票清单+挂起清单+按客户的关账对账差额） | openpyxl |
| [`20260929/`](20260929/) | 多客户检测数据汇总报表自动生成（一套明细按各客户的报表口径各出一页：分组维度与表头组名由客户口径表驱动、合格判定按检测项目各自的标准且支持「<」报告值与文字结论、复检重出取最新而作废件不算最新、未出报告与判不了标准的单独计数不进合格率分母） | openpyxl |
| [`20260930/`](20260930/) | 第三方检测检测周期进度跟踪台账（承诺周期按项目取最长、分包件走分包周期、从收样日按工作日倒推销出承诺交期且跨假期含调休、当前环节按业务顺序还原不按时间戳排、剩余 0 归临近超期、缺日期/缺周期/缺环节/时间倒挂整单挂起、另出催办清单） | openpyxl |
| [`20261001/`](20261001/) | 填 Excel 报表模板时公式不丢（值列照写、公式列一个值都不写：预置公式按行数用 Translator 平移延展、多余预置行清值留样式否则空行被算成合格、汇总 COUNTIF 范围按实际行数重写、回读校验读的是公式字符串） | openpyxl |
| [`20261002/`](20261002/) | 按规则表把 Excel 台账批量导出成 PDF（列宽按字符数换算后等比缩到一页宽、跨页每页重复表头、页脚「第 X 页 / 共 Y 页」两遍法、纵向合并的单元格值下填、文件名占位符清洗与重名加序号、只有表头的空白页与文件/工作表缺失分别挂起，源工作簿 sha256 前后比对证明一行未改） | openpyxl + reportlab |
| [`20261003/`](20261003/) | 色谱仪器导出的峰表文件批量汇总成 Excel（按列名反查峰表、靠「总计」行非数字认终点、按规则表剔溶剂峰再算面积归一、`<1000` 按半值估而 `n.d.` 整条挂起、同编号多次进样不合并只算平行样偏差、GBK 编码回退） | openpyxl |
| [`20261004/`](20261004/) | 用口径表把导出的结果明细直接出成月度质量报表（中心字典/项目口径/报表定义三张表驱动：单位与小数位按口径换算、只有「已出报告」进统计、结果分纯数值/只给边界/文字结论三种形态而只有纯数值进均值、合格率分母只算判过定的、明细区填不满清值留边框、未纳入记录单独成页并写清原因） | openpyxl |
| [`20261005/`](20261005/) | 把一个文件夹里各仪器导出的结果文件批量解析成一张汇总表（每种文件的样子写成一张格式表：后缀 + 编码 + 分隔符 + 表头特征列 + 列映射 + 判空列；按特征列名找表头不写死行号、编码依次试 utf-8-sig/gbk/utf-8、按判空列切掉表尾合计与说明行、结果分纯数值/只给边界/文字结论且单位可从结果里拆、重复导出按内容判重而同样品结果不一致的两条都挂起、认不出的与只有表头的分别写清原因） | openpyxl |
| [`20261006/`](20261006/) | 仪器导出的结果文件按统一结构自动入库（主键 = 样品编号 + 检测项目 + 单位；判定顺序＝不存在 → 同日同值（无变化）→ 同日不同值（冲突，两条都不当有效值）→ 日期更晚（更新，旧值搬进历史表）→ 日期更早（跳过）；用 ON CONFLICT DO UPDATE 而不是 INSERT OR REPLACE，首次入库时间不被重置；结果分纯数值/低于下限/文字结论/空四态分开存，也算不出全空的均值；批次一个事务，整批重放新增与更新都为 0） | 标准库 sqlite3 |
| [`20261008/`](20261008/) | ELN 分次导出的实验记录按一张字段映射表整合成汇总表（列名按别名归一、表头行靠"命中已知字段名≥2"自动定位、GBK/UTF-8 回退、按列名不按下标读；前导零与「未测」原样保留、空值不填补；不认识的列、同名字段、非数值全部记进「字段映射留痕」；认不出的文件与缺主键的行挂起） | pandas + openpyxl |
| [`tools/windows-health-check/`](tools/windows-health-check/) | Windows 电脑体检与安全优化批处理脚本（一次查完 CPU / 内存 / 磁盘 / 网络 / 启动项 / 防火墙并出报告；一键做可逆清理：临时文件、DNS 刷新、广告与推荐关闭、SSD TRIM；改注册表前自动备份、不可逆操作单独确认且默认「否」；含 GBK 编码、`%` 转义、`Run32` 禁用标记、中文输入法吃回车四类实测坑的记录） | 无（Windows 自带命令 + PowerShell） |

## 怎么用

进入任一目录，看它自己的 `README.md`，然后：

```bash
python run.py
```

脚本会自己切到所在目录、缺依赖自动装、跑完自动回读校验并打印 PASS/FAIL。Windows 上也可直接双击 `run.py`。

## 每个示例的目录结构

```
YYYYMMDD/
├── README.md           说明与排错（先看这个）
├── run.py              一键脚本
├── 01_raw_data/        原始数据（模拟仪器导出）
├── 02_output/          跑完的成品（脚本自动生成）
└── source/
    ├── templates/      报告模板
    ├── run_log.txt     运行日志（每次追加）
    ├── build_fixtures.py  重新生成模板与数据
    └── requirements.txt
```

## 说明

- 示例数据为**故意构造的边界值**（空值、前导零、极端条数），用来暴露真实踩过的坑。数据与模板均为虚构，请勿用于生产。
- 所有代码零硬编码路径，Windows / macOS / Linux 通用。
- 每个示例的 `README.md` 里记了当时实测发现的问题与修法，比代码本身更值得看。

---

<a id="english"></a>
# English

**Runnable code** that accompanies the CSDN article series. Each directory is a standalone example — copy it out and it runs, with no dependency on the rest of this repo.

| Directory | Contents | Dependencies |
| --- | --- | --- |
| [`20260909/`](20260909/) | Batch-generate test reports from an Excel template (preserves merged cells, named ranges, conclusion formulas, signature block) | openpyxl |
| [`20260910/`](20260910/) | Batch-generate inspection report documents from a Word template (only three spots change; headers, footers, and table styles preserved) | docxtpl |
| [`20260911/`](20260911/) | Auto-merge Excel exports from multiple instruments (header row auto-alignment, preserving leading zeros in batch numbers, keeping text results out of the mean) | pandas |
| [`20260912/`](20260912/) | Batch-generate PDF test reports from instrument-exported data (Chinese font registration, repeating header rows across pages, escaping special characters in cells) | reportlab |
| [`20260913/`](20260913/) | Auto-summarize stability-study data by batch × time point (pivot to wide format, items ordered per pharmacopoeia convention, out-of-spec cells highlighted red, blanks left empty) | pandas |
| [`20260914/`](20260914/) | One Word test report per sample (cross-run placeholder replacement, detail table rows generated by item count, illegal filename character handling) | python-docx |
| [`20260915/`](20260915/) | Central-lab sample receipt log (scan details aligned with temperature records by waybill number, both ends of the temperature range checked, anomalies split into three conclusion tiers with full-row coloring) | openpyxl |
| [`20260916/`](20260916/) | Auto-generate blood-draw Kit configuration sheets for a central lab (protocol usage × per-site enrollment to compute kit counts, FEFO batch ordering, label list expanded per kit, shortages and conflicts on a separate page) | openpyxl |
| [`20260917/`](20260917/) | Batch-distribute multi-site test reports to each site (per-site naming templates, per-subject packaging, TAT calculation, numbering/original/naming-collision exceptions on a separate page) | openpyxl |
| [`20260918/`](20260918/) | Chain sample handover records into a ledger (chain by stage rules; missing links, time inversions, duplicate stages, and overlong stays each picked out; IDs that lost leading zeros listed separately) | openpyxl |
| [`20260919/`](20260919/) | Assemble visit-window reports from visit windows and item lists (early visits still count as out-of-window, not-visited vs. not-yet-due split, missing items flagged against the visit checklist, off-protocol items retained) | openpyxl |
| [`20260920/`](20260920/) | Produce reports from one dataset using each site's own template (locate cells by field labels instead of hard-coded coordinates, label normalization, unit conversion per site with the unit column updated, no forcing fields the template doesn't want) | openpyxl |
| [`20260921/`](20260921/) | Convert legacy paper sample receipts to digital in bulk (fields classified by source, handwritten quantities checked line by line against details, held-back receipts produce no output, online IDs numbered by day + site) | openpyxl |
| [`20260922/`](20260922/) | Central-lab TAT overdue reminders (start point chosen per the counting rule, clock-stop periods deducted, working-day counting across adjusted holidays cell by cell, reminder lists and follow-up text in five tiers) | openpyxl |
| [`20260923/`](20260923/) | Batch-generate third-party testing work orders (intake register grouped by date + client, dates and client names normalized before grouping, six categories of missing items put on hold without consuming numbers, template cells located by label) | openpyxl |
| [`20260924/`](20260924/) | Batch-summarize third-party testing quotes and compare against historical prices (price list in long format keyed by method + effective period, item names normalized and matched via alias table, discounts tiered by quotable item count, held items listed separately rather than quoted arbitrarily) | openpyxl |
| [`20260925/`](20260925/) | Batch-rename and distribute third-party test reports per client naming rules (placeholder-based renaming, per-client packaging that copies without touching originals; five hold categories: name collision, unregistered recipient, missing original, pending review, voided; distribution register expanded by report × recipient) | openpyxl |
| [`20260926/`](20260926/) | Auto-register submitted samples and generate IDs (segmented numbering by date + client code + category code, a second same-day batch for the same client continues the sequence, one row expanded per unit as -1…-N, rows that can't be numbered are held without consuming numbers, ledger and label list from the same source) | openpyxl |
| [`20260927/`](20260927/) | Batch-stamp third-party test reports (stamp rules looked up by report type, three placements: first page / last page / right-edge cross-page seal; cross-page seals half outside the page and stepped down by page order, transparent mask preserved, five categories of ineligible reports held whole, ledger by report × stamp) | openpyxl + pypdf |
| [`20260928/`](20260928/) | Testing-lab invoicing requests and monthly closing reconciliation (each request checked against report ledger / invoiced ledger / client rule table: invoice title and tax ID, report not yet issued, PO required, amount mismatch, duplicate invoicing, framework-contract tax-exclusive allowance; voided invoices neither consume allowance nor block reissue; outputs a billable list, a hold list, and per-client closing differences) | openpyxl |
| [`20260929/`](20260929/) | Auto-generate multi-client test data summary reports (one dataset rendered per client's report spec: grouping dimensions and header group names driven by the client spec table, pass/fail judged against each item's own criterion including "<" reported values and text conclusions, retests take the latest while voided ones don't count as latest, not-yet-issued and unjudgeable results counted separately and excluded from the pass-rate denominator) | openpyxl |
| [`20260930/`](20260930/) | Third-party testing turnaround progress ledger (promised cycle takes the longest per item, subcontracted items use the subcontract cycle, committed dates back-calculated from receipt date by working days including adjusted holidays, current stage restored by business order rather than timestamps, zero remaining marks near-overdue, missing date/cycle/stage or time inversions hold the whole order, plus a follow-up list) | openpyxl |
| [`20261001/`](20261001/) | Fill Excel report templates without losing formulas (value columns written, formula columns get no values: preset formulas extended by row count via Translator, surplus preset rows cleared but styled so blank rows aren't counted as passes, summary COUNTIF ranges rewritten to actual row count, read-back verifies the formula string) | openpyxl |
| [`20261002/`](20261002/) | Batch-export Excel ledgers to PDF per a rule table (column widths scaled to fit one page wide, headers repeated on every page, "Page X of Y" footer via the two-pass method, vertically merged cell values filled down, filename placeholders sanitized and duplicates suffixed, header-only blank pages and missing files/worksheets held separately; source workbook sha256 compared before and after to prove nothing changed) | openpyxl + reportlab |
| [`20261003/`](20261003/) | Batch-summarize chromatograph peak-table exports into Excel (peak tables located by column names, end detected by a non-numeric "Total" row, solvent peaks removed per a rule table before area normalization, `<1000` estimated at half value while `n.d.` holds the entire record, repeated injections of the same ID not merged — only parallel-sample deviation computed, GBK encoding fallback) | openpyxl |
| [`20261004/`](20261004/) | Turn exported results directly into monthly quality reports via spec tables (driven by three tables — site dictionary, item spec, report definition: units and decimals converted per spec, only "report issued" records counted, results split into pure numeric / boundary-only / text conclusion with only pure numerics entering the mean, pass-rate denominator counts only judged results, detail area cleared but bordered when not filled, excluded records on a separate page with reasons) | openpyxl |
| [`20261005/`](20261005/) | Batch-parse a folder of instrument exports into one summary table (each file type described in a format table: extension + encoding + delimiter + header feature columns + column mapping + null-check column; header located by feature column names rather than a fixed row number, encodings tried as utf-8-sig/gbk/utf-8 in order, trailing totals and notes trimmed by the null-check column, results split into pure numeric / boundary-only / text with unit extractable from the result, duplicate exports detected by content while two same-sample records with differing results are both held, unrecognized files and header-only files reported separately with reasons) | openpyxl |
| [`20261006/`](20261006/) | Auto-ingest instrument result files into a database with a uniform structure (primary key = sample ID + test item + unit; decision order: not exists → same day same value (no change) → same day different value (conflict, neither treated as valid) → later date (update, old value moved to history) → earlier date (skip); uses ON CONFLICT DO UPDATE rather than INSERT OR REPLACE so the first-ingest timestamp is not reset; results stored in four distinct states — pure numeric / below limit / text conclusion / empty — and an all-empty mean is never computed; one transaction per batch, replaying a whole batch yields zero inserts and zero updates) | standard library `sqlite3` |
| [`20261008/`](20261008/) | Consolidate ELN records exported in batches into a summary table via a field mapping table (column names normalized through aliases, header row auto-located by "≥2 known field names matched", GBK/UTF-8 fallback, read by column name not index; leading zeros and "未测" (not tested) preserved verbatim, blanks not filled; unknown columns, duplicate field names, and non-numeric values all recorded in a "field mapping audit trail"; unrecognized files and rows missing a primary key are held) | pandas + openpyxl |
| [`tools/windows-health-check/`](tools/windows-health-check/) | Windows PC health-check and security-optimization batch script (checks CPU / memory / disk / network / startup items / firewall in one pass and produces a report; one-click reversible cleanup: temp files, DNS flush, ad and recommendation opt-out, SSD TRIM; registry backed up before changes, irreversible operations confirmed separately and defaulting to "No"; documents four field-tested pitfalls — GBK encoding, `%` escaping, the `Run32` disable marker, and Chinese IMEs swallowing the Enter key) | None (built-in Windows commands + PowerShell) |

## How to Use

Enter any directory, read its own `README.md`, then:

```bash
python run.py
```

The script changes into its own directory, installs missing dependencies, reads back and verifies after running, and prints PASS/FAIL. On Windows you can also just double-click `run.py`.

## Per-example Directory Structure

```
YYYYMMDD/
├── README.md           Notes and troubleshooting (read this first)
├── run.py              One-click script
├── 01_raw_data/        Raw data (simulated instrument exports)
├── 02_output/          Generated output (created by the script)
└── source/
    ├── templates/      Report templates
    ├── run_log.txt     Run log (appended each run)
    ├── build_fixtures.py  Regenerate templates and data
    └── requirements.txt
```

## Notes

- The sample data consists of **deliberately constructed edge cases** (blanks, leading zeros, extreme row counts) intended to expose real pitfalls encountered in practice. Data and templates are fictional — do not use them in production.
- All code has zero hard-coded paths and works on Windows / macOS / Linux.
- Each example's `README.md` records the issues found during actual testing and how they were fixed — often more valuable than the code itself.
