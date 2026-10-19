# 示例：SOP 受控文档批量渲染（模板 + 台账 → 一批 docx）

一张台账（文件编号 / 名称 / 版本 / 状态 / 生效日期 / 修订人 / 修订内容）配一份带占位符的模板，
一次跑出所有「现行版」SOP 的 Word 文件：页眉版本号、页脚编号、正文交叉引用、文末修订历史表
都按台账自动对齐，不再一份份手改。

## 目录结构

```
20261019/
├── README.md                  ← 先看这个
├── run.py                     ← 跑主线（你要碰的）
├── 01_raw_data/
│   └── SOP台账.xlsx           ← 输入：台账（你要碰的）
├── 02_output/                 ← 输出：生成的 docx（跑完才有，不进 git）
└── source/
    ├── templates/
    │   └── SOP模板.docx        ← 输入：受控文档模板（你要改的就是它）
    ├── build_fixtures.py      ← 重造模板与台账
    ├── requirements.txt
    └── run_log.txt            ← 运行日志（不进 git）
```

01 → 02 就是操作顺序：`01_raw_data` 进，`02_output` 出。

## 怎么跑

```bash
python run.py --no-pause
```

Windows 直接双击 `run.py` 也行（结尾会停下来等回车，否则窗口一闪就关）。
缺依赖会自动用清华源装 python-docx 与 openpyxl；手动装：

```bash
pip install -i https://pypi.tuna.tsinghua.edu.cn/simple -r source/requirements.txt
```

想重造模板和台账：`python source/build_fixtures.py`。

## 换成自己的数据

1. 把 `source/templates/SOP模板.docx` 换成你们的模板，占位符保留原样：
   `{{DOC_NO}}` `{{TITLE}}` `{{VER}}` `{{EFF}}` `{{OWNER}}`，引用别处文件写 `{{SOP:编号}}`。
2. 台账表头与 `run.py` 顶部 `COL` 对齐（改表头就改 `COL`，别去改后面的代码）。
3. 台账里同一编号允许多行历史版本，`状态` 列写 `现行` 的那行才会被生成。

## 故意造的边界

| 边界 | 想验证什么 |
| --- | --- |
| SOP-QC-003 有 V01/V02/V03 三行，只有 V03 是现行 | 历史版本不能混进输出 |
| SOP-QC-007 有 V9 与 V10 两行 | 版本号必须按数字比；按字符串比 `"V9" > "V10"` 会挑错 |
| SOP-QA-011 全部作废 | 无现行版应跳过并报出来，而不是崩掉 |
| 模板引用了台账里没有的 SOP-QA-030 | 未登记的引用应标成 `[待确认:编号]`，不能静默留空 |
| SOP-QC-021 只有一条台账行 | 只有一份时循环不出错 |

## 实测发现的问题

- **版本号按字符串比会挑错现行版**：`max()` 直接比字符串时 `"V9" > "V10"`，纯化水那份会拿 V9 当现行。
  改成先抠数字转 `int` 再比（`ver_key()`）。
- **页眉页脚不在 `doc.paragraphs` 里**：只遍历正文，页眉的版本号一个都换不掉，生成出来的文件看着成功、
  实际全是 `{{VER}}`。必须单独遍历 `section.header` / `section.footer`。
- **占位符会被 Word 拆成多个 run**：手工编辑过的模板里，一段话常被拆成好几段 run（改过格式就会拆），
  逐个 run 替换必然漏。改成整段取文本、替换完写回第一个 run。
- **`table.add_row()` 不继承上一行的字体字号**：只保证格子数量对。本例模板表格用文档默认字体，所以看不出来；
  换了自定义样式的表格要自己设一遍。

## 验证环境

- Python 3.13，python-docx 1.2 / openpyxl 3.1，Windows。
- 11 项回读断言全 PASS：生成 6 份、跳过 1 个编号（SOP-QA-011）、待确认引用 1 处（SOP-QA-030）。
- 已知限制：只处理 docx 正文段落、表格、页眉页脚里的占位符，不处理文本框与图表标题；
  页码域（PAGE）依赖 Word 打开时更新，本例没做强制更新。

## 排错

| 现象 | 原因 | 怎么办 |
| --- | --- | --- |
| `FileNotFoundError: SOP模板.docx` | 模板没生成或路径被挪走 | 先跑 `python source/build_fixtures.py` |
| 生成的文档里还有 `{{VER}}` | 占位符被拆在多个 run 里，或写成了 `{{ Ver }}` | 占位符保持全大写、无空格、单独一段最稳 |
| 一份都没生成 | 台账 `状态` 列不是「现行」两字 | 改成与 `run.py` 的 `CURRENT` 一致 |
| 引用的文件展开失败 | 编号在台账里不存在，或不在现行版里 | 看输出里 `[待确认:编号]` 标记，补台账 |
| 中文文件名乱码（Linux/macOS） | 系统 locale 非 UTF-8 | 设 `LANG=zh_CN.UTF-8` 再跑 |
