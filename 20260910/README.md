# 示例：按 Word 模板批量生成检验报告书

配套长文：《Python 批量生成 Word 报告保留模板样式》（CSDN 发布后补链接）
`run.py` 里的核心代码与文章正文**逐字一致**，跑通即等于文章代码跑通。

## 一、怎么跑

本文件跟 `run.py` 并排放在最外层，看完接着跑就行。一条命令，Windows / macOS / Linux 通吃：

```bash
python run.py
```

Windows 上也可以**直接双击 `run.py`**：脚本会自己切到所在目录、缺依赖就自动装、
跑完停住窗口等你按回车。不想停就加 `--no-pause`。

依赖：docxtpl + python-docx + jinja2，脚本会自动装。想先手动装：

```bash
pip install -i https://pypi.tuna.tsinghua.edu.cn/simple docxtpl python-docx jinja2
```

跑完看 `02_output/`，两份报告书，回读校验 5 项全绿打印 PASS。
日志写进 `source/run_log.txt`（UTF-8，**每次追加**，历史留着不覆盖）。
每段开头先记「运行主机」（主机名 + 本机 IP，用来分辨是哪台机器跑的），再记「开始执行」时间和解释器。
完成行在结束时立即写下、不等按回车，跑完直接关窗口日志里也有完成时间和判定；
若某条只有开始没有完成，下次运行会自动补一句「上一条没有留下完成时间」。

## 二、目录结构

外层只放你要碰的四样东西（01 → 02 就是操作顺序），其余全收进 `source/`：

```
20260910/
├── README.md                 ← 本文件，先看这个
├── run.py                    ← 再跑这个
├── 01_raw_data/              ← 你的原始数据放这
│   └── 检验结果.csv            utf-8-sig 带 BOM，跟 LIMS 导出一致
├── 02_output/                ← 跑完的成品在这（脚本自动生成）
└── source/                   ← 模板、日志、维护脚本；日常不用碰
    ├── templates/
    │   └── 检验报告书模板.docx 模拟 QA 在 Word 里排好的模板
    ├── run_log.txt           运行日志（脚本自动生成，每次追加）
    ├── build_fixtures.py     重新生成模板和 CSV（改了模板结构后重跑）
    └── requirements.txt      依赖清单
```

整个目录直接拷给别人就能用：脚本用 `Path(__file__).parent` 定位，没有硬编码路径。
`build_fixtures.py` 只在需要重造数据时跑（`python source/build_fixtures.py`）。

## 三、验证到什么

两份样品，明细条数不同（4 条 / 3 条），回读 5 项全部 OK：

| 校验项 | 结果 |
| --- | --- |
| 表格行数 = 表头 1 + 明细 N | OK |
| 页眉里的 `{{ 报告编号 }}` 已替换 | OK |
| 抬头三个占位符已填 | OK |
| `单个杂质 <0.5% & 总量 ≤2.0%` 原样保留 | OK |
| 明细行字号 9.0pt（模板样式带出） | OK |

## 四、三个实测结论

1. 明细表循环必须「for 一行 / 内容一行 / endfor 一行」三段式。for 和 endfor 挤在同一行的首格和末格，**整张表会被静默清空**（表格还在，行数 0，不报错）。
2. 数据里的 `<` 和 `&` 会被 XML 吞掉。`<0.5%` → `0.5%`，`A&B` → `A`。必须传 `Environment(autoescape=True)` 给 `render()`，用 `|e` 反而丢字符。
3. 明细为空不报错，表格只剩表头行，看上去像漏做检项。建议提前判空。

## 五、给别人用的注意事项

| 项 | 说明 |
| --- | --- |
| 依赖 | docxtpl + python-docx + jinja2，脚本会自动装 |
| Python 版本 | 3.9 以上（代码用了 `list[dict]` 注解，3.8 会报 TypeError） |
| 路径 | 全相对定位，拷到哪都能跑 |
| 编码 | CSV 是 utf-8-sig，Excel 直接打开中文不乱码 |
| 日志 | `source/run_log.txt` 每次追加，先记运行主机（主机名 + IP），再记开始/完成时间和用时；关窗口也不丢完成行 |
| 要改的地方 | 模板换成 QA 自己的 Word 模板；`01_raw_data/检验结果.csv` 换成真实数据 |
| 已知限制 | 明细为空时表格只剩表头行，不报错，容易被误认为漏做检项 |

## 六、排错

| 现象 | 原因 | 怎么办 |
| --- | --- | --- |
| `02_output/` 是空的 | 脚本没跑到生成那步 | 看 `source/run_log.txt`，或命令行跑一遍看报错 |
| 自动装依赖失败 | 无写入权限 / 网络不通 | 手动：`python -m pip install --user -i https://pypi.tuna.tsinghua.edu.cn/simple docxtpl` |
| 控制台中文乱码 | cmd 代码页问题 | 看 `source/run_log.txt` |
| 表格里一行明细都没有 | for / endfor 没分成两行 | 见"三个实测结论"第 1 条 |
| `<0.5%` 变成 `0.5%` | 没传 autoescape | `doc.render(ctx, Environment(autoescape=True))` |
