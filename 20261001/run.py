# -*- coding: utf-8 -*-
"""实验室「检测结果报表」自动填表：值照写，公式不丢。

读 01_raw_data/ 里的两张表——仪器软件导出的结果明细、受检单位清单，
按受检单位归堆，套 source/templates/ 里的报表模板出报表：
值列照写，公式列不写死值——模板预置的公式按实际行数延展或清空，
汇总公式的统计范围跟着实际明细行数重写。
跑完自动回读校验（读的是公式字符串，不是 Excel 算出来的值），打印 PASS/FAIL。

用法：python run.py      （Windows 双击 run.py 也行；加 --no-pause 跳过结尾回车）
"""

import csv
import socket
import subprocess
import sys
from copy import copy
from datetime import date, datetime
from pathlib import Path

BASE = Path(__file__).resolve().parent
RAW_DIR = BASE / "01_raw_data"
OUT_DIR = BASE / "02_output"
SRC_DIR = BASE / "source"
TPL_FILE = SRC_DIR / "templates" / "检测结果报表模板.xlsx"
LOG_FILE = SRC_DIR / "run_log.txt"

RESULT_FILE = RAW_DIR / "仪器导出结果.csv"
UNIT_FILE = RAW_DIR / "受检单位清单.csv"
HANG_FILE = OUT_DIR / "挂起清单.csv"

BATCH = "20260930"          # 本批检测日期：抬头与文件名都用它，写死常量、不取当天
SHEET_NAME = "检测结果报表"
UNIT_KEY = "受检单位"
COL_SAMPLE = "样品编号"
COL_ITEM = "检测项目"
COL_AMOUNT = "含量(mg/kg)"
COL_JUDGE = "判定"
VALUE_COLS = ("称样量(g)", "定容体积(mL)", "稀释倍数", "仪器读数(mg/L)", "限值(mg/kg)")
HANG_COLS = [COL_SAMPLE, UNIT_KEY, COL_ITEM, "仪器读数(mg/L)", "挂起原因"]


def ensure_openpyxl():
    """缺依赖自动装（走清华源），装不上再退到 --user。"""
    try:
        import openpyxl  # noqa: F401
        return
    except ImportError:
        print("缺少 openpyxl，正在用清华源安装 …")
    mirror = ["-i", "https://pypi.tuna.tsinghua.edu.cn/simple"]
    for extra in ([], ["--user"]):
        if subprocess.call([sys.executable, "-m", "pip", "install", "openpyxl"] + extra + mirror) == 0:
            return
    raise SystemExit("openpyxl 装不上，请手动执行：pip install openpyxl")


def machine_tag():
    """运行主机标识：主机名 + 本机对外网卡 IP，用来分辨是哪台机器跑的。"""
    host = socket.gethostname()
    ip = "未知"
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        sock.connect(("8.8.8.8", 80))
        ip = sock.getsockname()[0]
        sock.close()
    except OSError:
        try:
            ip = socket.gethostbyname(host)
        except OSError:
            ip = "未知"
    return f"{host}（{ip}）"


def log_open():
    """开一段追加日志：先记运行主机，再记开始时间与解释器。"""
    LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
    fh = open(LOG_FILE, "a", encoding="utf-8")
    fh.write(f"\n运行主机：{machine_tag()}\n")
    fh.write(f"开始执行：{datetime.now():%Y-%m-%d %H:%M:%S}  "
             f"Python {sys.version.split()[0]}（{Path(sys.executable).name}）\n")
    return fh


def norm(text):
    """比较用的归一化：去掉所有空白（含全角空格 U+3000），统一成字符串。"""
    return "".join(str(text or "").split())


def to_float(text):
    """洗数字：洗不干净返回 None。写「<0.01」这类检出限表示法的读数，不许当 0.01 用。"""
    t = str(text or "").strip().replace(",", "")
    if not t:
        return None
    try:
        return float(t)
    except ValueError:
        return None


def read_table(path, key_col):
    """读 CSV；key_col 是这张表「靠它判空」的那一列——每张表不一样，必须当参数传。"""
    with open(path, newline="", encoding="utf-8-sig") as fh:
        return [r for r in csv.DictReader(fh) if str(r.get(key_col) or "").strip()]


def read_source():
    """读仪器导出的结果明细，按受检单位归堆；算不出含量的行挂起，不猜不兜。"""
    units, hangs = {}, []
    for r in read_table(RESULT_FILE, COL_SAMPLE):
        sample = (r.get(COL_SAMPLE) or "").strip()
        unit = norm(r.get(UNIT_KEY))
        item_name = (r.get(COL_ITEM) or "").strip()
        weigh = to_float(r.get("称样量(g)"))
        vol = to_float(r.get("定容体积(mL)"))
        dil = to_float(r.get("稀释倍数"))
        read = to_float(r.get("仪器读数(mg/L)"))
        limit = to_float(r.get("限值(mg/kg)"))

        # 判定顺序＝业务顺序：读数 → 称样量 → 定容 → 稀释倍数 → 限值
        reason = ""
        if read is None:
            reason = "仪器读数缺失或不是数字（检出限表示法不能直接参与计算）"
        elif weigh is None or weigh <= 0:
            reason = "称样量缺失或不是正数"
        elif vol is None:
            reason = "定容体积缺失"
        elif dil is None:
            reason = "稀释倍数缺失（不默认按 1 算）"
        elif limit is None:
            reason = "限值缺失，判定列算不出来"
        if reason:
            hangs.append({COL_SAMPLE: sample, UNIT_KEY: unit, COL_ITEM: item_name,
                          "仪器读数(mg/L)": (r.get("仪器读数(mg/L)") or "").strip(), "挂起原因": reason})
            continue

        amount = round(read * vol * dil / (weigh * 1000), 3)
        units.setdefault(unit, []).append({
            COL_SAMPLE: sample, COL_ITEM: item_name, "称样量(g)": weigh, "定容体积(mL)": vol,
            "稀释倍数": dil, "仪器读数(mg/L)": read, "限值(mg/kg)": limit,
            COL_AMOUNT: amount, COL_JUDGE: "合格" if amount <= limit else "不合格",
        })
    return units, hangs


def read_unit_plan():
    """读受检单位清单：拿本批报告编号与简称；清单里没有的单位不出报表，防文件名乱编。"""
    plan = {}
    for r in read_table(UNIT_FILE, UNIT_KEY):
        plan[norm(r.get(UNIT_KEY))] = {"报告编号": (r.get("本批报告编号") or "").strip(),
                                       "简称": norm(r.get("受检单位简称"))}
    return plan


def find_layout(ws):
    """按栏目名反查表头行与列号：模板挪了位置也不怕，写死坐标必漂。"""
    head_row = None
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        if any(c.value == COL_SAMPLE for c in row):
            head_row = row[0].row
            break
    if head_row is None:
        raise SystemExit(f"模板里找不到表头「{COL_SAMPLE}」，请检查模板")
    cols = {str(c.value).strip(): c.column for c in ws[head_row] if c.value}
    return head_row, cols


def label_value(ws, label):
    """按标签反查它右边那一格（抬头格），返回 (行, 列)。"""
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        for c in row:
            if str(c.value or "").strip() == label:
                return c.row, c.column + 1
    raise SystemExit(f"模板里找不到标签「{label}」")


def detail_rows(ws, head_row, amount_col):
    """表头下面、含量列写着公式的连续行就是模板预置明细行；靠公式认，不数行数。"""
    rows, r = [], head_row + 1
    while r <= ws.max_row:
        v = ws.cell(row=r, column=amount_col).value
        if isinstance(v, str) and v.startswith("="):
            rows.append(r)
        elif rows:
            break
        r += 1
    return rows


def find_summary(ws, judge_col):
    """汇总行靠「判定列里那个带 COUNTIF 的公式」认，不靠行号猜。"""
    for r in range(1, ws.max_row + 1):
        v = ws.cell(row=r, column=judge_col).value
        if isinstance(v, str) and "COUNTIF" in v:
            return r
    raise SystemExit("模板里找不到汇总行（判定列应有 COUNTIF 公式）")


def fit_detail(ws, preset, count, amount_col, judge_col):
    """把明细区调成 count 行：不够就插行拷样式，多了就清空值（样式留着，边框不掉）。"""
    from openpyxl.formula.translate import Translator
    from openpyxl.utils import get_column_letter

    last = preset[-1]
    if count > len(preset):
        extra = count - len(preset)
        ws.insert_rows(last + 1, extra)          # 插行：openpyxl 只搬格子，不动公式里的引用
        for r in range(last + 1, last + 1 + extra):
            for col in range(1, judge_col + 1):
                ws.cell(row=r, column=col)._style = copy(ws.cell(row=last, column=col)._style)

    rows = list(range(preset[0], preset[0] + count))
    src = preset[0]
    for r in rows:                               # 公式按行号平移：相对引用才会指到本行
        for col in (amount_col, judge_col):
            letter = get_column_letter(col)
            formula = ws.cell(row=src, column=col).value
            ws.cell(row=r, column=col,
                    value=Translator(formula, origin=f"{letter}{src}").translate_formula(f"{letter}{r}"))

    for r in preset[count:]:                     # 多出来的行：清值留样式，空行绝不会被算进合格率
        for col in range(1, judge_col + 1):
            ws.cell(row=r, column=col).value = None
    return rows


def retarget_summary(ws, rows, preset, judge_col):
    """汇总公式的统计范围不会自己跟着行数变，得按实际明细行重写。"""
    from openpyxl.utils import get_column_letter

    letter = get_column_letter(judge_col)
    old_ref = f"{letter}{preset[0]}:{letter}{preset[-1]}"
    new_ref = f"{letter}{rows[0]}:{letter}{rows[-1]}"
    cell = ws.cell(row=find_summary(ws, judge_col), column=judge_col)
    if old_ref in str(cell.value):
        cell.value = str(cell.value).replace(old_ref, new_ref)
    return new_ref


def fill_report(ws, unit, items, report_no):
    """把一家单位的明细填进模板副本：值列照写，公式列让它自己延展。"""
    head_row, cols = find_layout(ws)
    amount_col, judge_col = cols[COL_AMOUNT], cols[COL_JUDGE]
    preset = detail_rows(ws, head_row, amount_col)
    rows = fit_detail(ws, preset, len(items), amount_col, judge_col)

    for label, value in (("报告编号", report_no), (UNIT_KEY, unit), ("检测日期", batch_date())):
        r, c = label_value(ws, label)
        cell = ws.cell(row=r, column=c, value=value)
        if label == "检测日期":
            cell.number_format = "yyyy-mm-dd"     # 不设格式，Excel 里会显示成 46265

    for i, item in enumerate(items, start=1):
        r = rows[i - 1]
        ws.cell(row=r, column=1, value=i)
        ws.cell(row=r, column=cols[COL_SAMPLE], value=item[COL_SAMPLE]).number_format = "@"
        ws.cell(row=r, column=cols[COL_ITEM], value=item[COL_ITEM])
        for name in VALUE_COLS:
            ws.cell(row=r, column=cols[name], value=item[name])

    retarget_summary(ws, rows, preset, judge_col)
    return rows


def batch_date():
    return date(int(BATCH[:4]), int(BATCH[4:6]), int(BATCH[6:]))


def build_one(unit, items, plan):
    """一家单位一份报表：从模板重新 load 一次，模板本身一个格子都不动。"""
    from openpyxl import load_workbook

    wb = load_workbook(TPL_FILE)      # 不加 data_only：加了公式会被读成缓存值，保存后整列公式消失
    ws = wb[SHEET_NAME]
    rows = fill_report(ws, unit, items, plan["报告编号"])
    out = OUT_DIR / f"检测结果报表_{plan['简称']}_{BATCH}.xlsx"
    wb.save(out)
    wb.close()
    return out, rows


def write_hangs(hangs):
    with open(HANG_FILE, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=HANG_COLS)
        w.writeheader()
        w.writerows(hangs)


def verify():
    """回读两份报表核对：公式列读的是公式字符串（不打开 data_only），不是 Excel 算出来的值。"""
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter

    units, hangs = read_source()
    plan = read_unit_plan()
    checks = []

    def want(caption, got, exp):
        good = got == exp
        checks.append(good)
        print(f"{'OK  ' if good else 'FAIL'} {caption}：{got!r}" + ("" if good else f"（期望 {exp!r}）"))

    want("挂起条数", len(hangs), 4)
    want("有数据的受检单位数", len(units), 2)
    want("丰味有效行数", len(units["上海丰味食品有限公司"]), 7)
    want("清源有效行数", len(units["苏州清源饮料有限公司"]), 2)

    for unit, exp_rows, exp_rate, exp_blank in (
        ("上海丰味食品有限公司", 7, 0.8571, 0),
        ("苏州清源饮料有限公司", 2, 1.0, 2),
    ):
        items = units[unit]
        short = plan[unit]["简称"]
        out = OUT_DIR / f"检测结果报表_{short}_{BATCH}.xlsx"
        want(f"{short} 报表文件已生成", out.exists(), True)
        wb = load_workbook(out)
        ws = wb[SHEET_NAME]
        head_row, cols = find_layout(ws)
        amount_col, judge_col = cols[COL_AMOUNT], cols[COL_JUDGE]
        letter = get_column_letter(judge_col)
        rows = detail_rows(ws, head_row, amount_col)
        sum_row = find_summary(ws, judge_col)

        want(f"{short} 明细行数", len(rows), exp_rows)
        want(f"{short} 表头行", head_row, 6)
        want(f"{short} 含量列全是公式、没被写成值",
             all(str(ws.cell(row=r, column=amount_col).value).startswith("=") for r in rows), True)
        want(f"{short} 判定公式都指向本行",
             all(f"H{r}<=" in str(ws.cell(row=r, column=judge_col).value) for r in rows), True)
        want(f"{short} 汇总行紧跟在明细下面", sum_row, max(rows) + exp_blank + 1)
        want(f"{short} 汇总范围随行数重写",
             f"{letter}{rows[0]}:{letter}{rows[-1]}" in str(ws.cell(row=sum_row, column=judge_col).value), True)

        blanks = list(range(max(rows) + 1, sum_row))
        want(f"{short} 多余预置行数", len(blanks), exp_blank)
        want(f"{short} 多余行已清空（值+公式）",
             all(ws.cell(row=r, column=amount_col).value is None
                 and ws.cell(row=r, column=judge_col).value is None for r in blanks), True)
        want(f"{short} 多余行边框还在",
             all(ws.cell(row=r, column=amount_col).border.left.style == "thin" for r in blanks), True)
        want(f"{short} 明细行全带边框",
             all(ws.cell(row=r, column=amount_col).border.left.style == "thin" for r in rows), True)

        r_no, c_no = label_value(ws, "报告编号")
        r_u, c_u = label_value(ws, UNIT_KEY)
        r_d, c_d = label_value(ws, "检测日期")
        want(f"{short} 报告编号", ws.cell(row=r_no, column=c_no).value, plan[unit]["报告编号"])
        want(f"{short} 受检单位（全角空格已归一）", ws.cell(row=r_u, column=c_u).value, unit)
        want(f"{short} 检测日期格式", ws.cell(row=r_d, column=c_d).number_format, "yyyy-mm-dd")
        want(f"{short} 检测日期值", ws.cell(row=r_d, column=c_d).value.date(), batch_date())
        want(f"{short} 样品编号保持文本（前导零没丢）",
             (ws.cell(row=rows[0], column=cols[COL_SAMPLE]).value,
              ws.cell(row=rows[0], column=cols[COL_SAMPLE]).number_format),
             (items[0][COL_SAMPLE], "@"))
        want(f"{short} 合格率期望（Python 按同一口径算）",
             round(sum(1 for it in items if it[COL_JUDGE] == "合格") / len(items), 4), exp_rate)
        wb.close()

    want("挂起行都不进报表（说明行也没被当数据）",
         sum(len(v) for v in units.values()) + len(hangs), 13)
    want("挂起原因齐全", all(h["挂起原因"] for h in hangs), True)

    ok = all(checks)
    print(f"\n判定：{'PASS' if ok else 'FAIL'}（{sum(checks)}/{len(checks)} 项通过）")
    return ok


def main():
    ensure_openpyxl()
    fh = log_open()
    started = datetime.now()
    ok = False
    try:
        OUT_DIR.mkdir(parents=True, exist_ok=True)
        units, hangs = read_source()
        plan = read_unit_plan()
        for unit, items in units.items():
            if unit not in plan:
                for it in items:
                    it["挂起原因"] = "受检单位不在清单里，报告编号与报表文件名都拿不到"
                hangs.extend(items)
                continue
            out, rows = build_one(unit, items, plan[unit])
            print(f"{plan[unit]['简称']}：{len(items)} 行 → {out.name}（明细行 {rows[0]}-{rows[-1]}）")
        write_hangs(hangs)
        print(f"挂起 {len(hangs)} 行 → {HANG_FILE.name}")
        ok = verify()
    finally:
        fh.write(f"执行完成：{datetime.now():%Y-%m-%d %H:%M:%S}  "
                 f"用时 {(datetime.now() - started).total_seconds():.1f}s  判定 {'PASS' if ok else 'FAIL'}\n")
        fh.close()
    if "--no-pause" not in sys.argv:
        input("\n按回车键退出 …")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
