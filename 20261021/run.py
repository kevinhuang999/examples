# -*- coding: utf-8 -*-
"""实验室 Excel 模板版本盘点：一堆散落的模板 -> 版本清单 + 差异明细 + 受控模板。

跑法：python run.py            （Windows 也可以直接双击）
      python run.py --no-pause

文件里另外带了三样工程外壳，正文不逐行展开：缺依赖自动装（清华源）、
运行日志追加写、跑完回读断言打印 PASS/FAIL。
"""

import re
import socket
import subprocess
import sys
import time
import unicodedata
from copy import copy
from datetime import datetime
from pathlib import Path

BASE = Path(__file__).resolve().parent
RAW_DIR = BASE / "01_raw_data"
OUT_DIR = BASE / "02_output"
LIST_XLSX = OUT_DIR / "模板版本清单.xlsx"
CLEAN_XLSX = OUT_DIR / "统一后模板_检验记录.xlsx"
LOG = BASE / "source" / "run_log.txt"

# 只有文件名带「模板」的才进盘点：同一个共享盘里还躺着台账、进度表，混进来会把版本数算飞
NAME_KEY = "模板"
# 表头不一定在第一行，靠这几个列名找它在第几行
HEADER_KEYS = ("样品编号", "检测项目", "检测结果", "单位")
# 表头最多往前看几行（标题行 + 填写说明行，现实里最多压两三行）
HEADER_SCAN = 10
# 版本号写在标题行里，形如「版本号：V2.1」，用正则抓
VER_RE = re.compile(r"V\s*(\d+(?:\.\d+)?)", re.I)
# 文件名里带这些词，容易让人误以为是最新版
HINT_WORDS = ("最终", "最新", "新")
# 只认这两种，.xls 是旧格式，openpyxl 打不开
SKIP_SUFFIX = (".xlsx", ".xlsm")


def ensure_deps():
    try:
        import openpyxl  # noqa: F401
        return
    except ImportError:
        pass
    for extra in ([], ["--user"]):
        cmd = [sys.executable, "-m", "pip", "install",
               "-i", "https://pypi.tuna.tsinghua.edu.cn/simple", "openpyxl"] + extra
        if subprocess.call(cmd) == 0:
            return
    raise SystemExit("依赖安装失败，请手动执行：pip install openpyxl")


def host_label():
    name = socket.gethostname()
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        sock.connect(("8.8.8.8", 80))
        ip = sock.getsockname()[0]
        sock.close()
    except OSError:
        try:
            ip = socket.gethostbyname(name)
        except OSError:
            ip = "未知"
    return f"{name}（{ip}）"


def log_head():
    LOG.parent.mkdir(parents=True, exist_ok=True)
    with LOG.open("a", encoding="utf-8") as fh:
        fh.write(f"运行主机：{host_label()}\n")
        fh.write(f"开始执行：{datetime.now():%Y-%m-%d %H:%M:%S}\n")
        fh.write(f"Python：{sys.version.split()[0]}（{Path(sys.executable).name}）\n")


def log_tail(seconds, verdict):
    with LOG.open("a", encoding="utf-8") as fh:
        fh.write(f"执行完成：{datetime.now():%Y-%m-%d %H:%M:%S}  "
                 f"用时 {seconds:.1f}s  判定 {verdict}\n")


def norm(value):
    """列名归一：全角转半角、去掉所有空白、转小写，专门用来比列名同不同。"""
    text = unicodedata.normalize("NFKC", str(value if value is not None else ""))
    return re.sub(r"\s+", "", text).lower()


def find_header(rows):
    """表头在第几行？返回 0 基行号；命中列名最多的那一行才算表头。"""
    keys = {norm(k) for k in HEADER_KEYS}
    best, best_hits = -1, 0
    for i, row in enumerate(rows[:HEADER_SCAN]):
        hits = len(keys & {norm(c) for c in row})
        if hits > best_hits:
            best, best_hits = i, hits
    return best if best_hits >= 2 else -1


def read_version(rows, head):
    """在表头上方几行里找「V2.1」这种版本号，找不到就返回 None。"""
    for row in rows[:head + 1]:
        for cell in row:
            if isinstance(cell, str):
                hit = VER_RE.search(cell)
                if hit:
                    return f"V{hit.group(1)}"
    return None


def ver_key(ver):
    """把 V2.2 变成 (2, 2) —— 直接按字符串比会把 V2.10 排到 V2.9 前面。"""
    return tuple(int(x) for x in ver[1:].split(".")) if ver else ()


def probe(path):
    """读一个文件，取出版本号、表头行、列名、可见工作表数，算结构指纹。"""
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    visible = [ws.title for ws in wb.worksheets if ws.sheet_state == "visible"]
    hidden = [ws.title for ws in wb.worksheets if ws.sheet_state != "visible"]
    sheet = wb[visible[0]]
    rows = [list(r) for r in sheet.iter_rows(max_row=HEADER_SCAN + 3, values_only=True)]
    wb.close()

    head = find_header(rows)
    if head < 0:
        raise ValueError(f"前 {HEADER_SCAN} 行里找不到表头，列名可能被改过")
    cols = [c for c in rows[head] if norm(c)]
    return {
        "path": path,
        "name": path.name,
        "sub": path.parent.relative_to(RAW_DIR).as_posix(),
        "ver": read_version(rows, head),
        "header_row": head + 1,
        "cols": cols,
        "visible": visible,
        "hidden": hidden,
        "mtime": datetime.fromtimestamp(path.stat().st_mtime),
        # 指纹 = 可见页签个数 + 列名集合（排序后取，列序不算差异）
        "sig": f"{len(visible)}|{'/'.join(sorted(norm(c) for c in cols))}",
    }


def scan():
    """走一遍模板目录：该跳过的跳过，该报错的报错，别让一个坏文件带崩整轮。"""
    taken, skipped = [], []
    for path in sorted(RAW_DIR.rglob("*")):
        if not path.is_file():
            continue
        name = path.name
        if name.startswith("~$"):
            skipped.append((name, "Excel 打开时留下的锁文件，不是模板"))
            continue
        if path.suffix.lower() == ".xls":
            skipped.append((name, "旧格式 .xls，先另存为 .xlsx 再盘点"))
            continue
        if path.suffix.lower() not in SKIP_SUFFIX:
            skipped.append((name, "不是 Excel 文件"))
            continue
        if NAME_KEY not in path.stem:
            skipped.append((name, "文件名不含「模板」，不在这次盘点范围"))
            continue
        try:
            taken.append(probe(path))
        except Exception as exc:                      # 坏文件只记账，不中断
            skipped.append((name, f"打不开：{exc}"))
    return taken, skipped


def pick_current(files):
    """组内挑现行版：先比版本号，版本号一样再比修改时间，最后按文件名。"""
    return sorted(files, key=lambda f: (ver_key(f["ver"]), f["mtime"], f["name"]),
                  reverse=True)


def diff_cols(base_cols, other_cols):
    """两组列名的差：多出来的、缺掉的，回原样列名方便人看。"""
    base = {norm(c): c for c in base_cols}
    other = {norm(c): c for c in other_cols}
    extra = [other[k] for k in other if k not in base]
    missing = [base[k] for k in base if k not in other]
    return extra, missing


def group_files(taken):
    """按指纹归堆，堆按文件数从多到少编号 G1、G2……"""
    buckets = {}
    for item in taken:
        buckets.setdefault(item["sig"], []).append(item)
    groups = []
    for sig, files in buckets.items():
        ordered = pick_current(files)
        groups.append({"sig": sig, "files": ordered, "current": ordered[0]})
    groups.sort(key=lambda g: (len(g["files"]), ver_key(g["current"]["ver"])), reverse=True)
    for idx, group in enumerate(groups, 1):
        group["no"] = f"G{idx}"
    return groups


def write_list(groups, skipped, max_ver):
    """写《模板版本清单.xlsx》：版本清单 / 文件明细 / 差异明细 / 跳过清单四张表。"""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill

    wb = Workbook()
    head_font = Font(bold=True, color="FFFFFF")
    head_fill = PatternFill("solid", fgColor="4472C4")

    def put(ws, header, rows, widths):
        ws.append(header)
        for cell in ws[1]:
            cell.font, cell.fill = head_font, head_fill
            cell.alignment = Alignment(horizontal="center")
        for row in rows:
            ws.append(row)
        for idx, width in enumerate(widths, 1):
            ws.column_dimensions[ws.cell(row=1, column=idx).column_letter].width = width
        ws.freeze_panes = "A2"

    # ① 版本清单：一组一行
    ws = wb.active
    ws.title = "版本清单"
    rows = []
    for g in groups:
        cur = g["current"]
        vers = "/".join(sorted({f["ver"] or "无版本号" for f in g["files"]}))
        note = "现行版已定位"
        if len(g["files"]) == 1 and g["no"] != "G1":
            note = "单独一版，版式与主版不同，要不要合并需人工定"
        if any("待确认" in flag for flag in (f["flag"] for f in g["files"])):
            note += "；组内有文件没版本号"
        rows.append([g["no"], len(g["files"]), len(cur["cols"]), len(cur["visible"]),
                     cur["name"], cur["ver"] or "无版本号",
                     cur["mtime"].strftime("%Y-%m-%d"), vers, note])
    put(ws, ["组号", "文件数", "列数", "可见工作表", "现行候选文件", "版本号",
             "最近修改", "组内版本", "备注"], rows,
        [8, 8, 8, 12, 30, 10, 12, 16, 46])

    # ② 文件明细：一个文件一行
    ws = wb.create_sheet("文件明细")
    rows = []
    for g in groups:
        for pos, f in enumerate(g["files"]):
            if f["ver"] is None:
                verdict = "待确认（无版本号）"
            elif pos == 0 and g["no"] == "G1":
                verdict = "现行"                # 只有文件最多的那组，首位才算定下来的现行版
            elif pos == 0:
                verdict = "待并版（组内唯一）"
            else:
                verdict = "旧版"
            rows.append([g["no"], f["name"], f["sub"], f["ver"] or "",
                         f["header_row"], len(f["cols"]), len(f["visible"]),
                         f["mtime"].strftime("%Y-%m-%d %H:%M"), verdict, f["flag"]])
    put(ws, ["组号", "文件名", "所在目录", "版本号", "表头行", "列数",
             "可见工作表", "修改时间", "判定", "说明"], rows,
        [8, 30, 12, 10, 8, 8, 12, 18, 20, 52])

    # ③ 差异明细：组内其他文件 vs 组内现行；各组现行 vs 全局基准
    ws = wb.create_sheet("差异明细")
    base = groups[0]["current"]
    rows = []
    for g in groups:
        cur = g["current"]
        for f in g["files"][1:]:
            extra, missing = diff_cols(cur["cols"], f["cols"])
            sheet_gap = len(f["visible"]) - len(cur["visible"])
            rows.append([g["no"], f["name"], cur["name"], "、".join(extra) or "—",
                         "、".join(missing) or "—",
                         f"{sheet_gap:+d}" if sheet_gap else "0"])
        if g["no"] != "G1":
            extra, missing = diff_cols(base["cols"], cur["cols"])
            sheet_gap = len(cur["visible"]) - len(base["visible"])
            rows.append([g["no"], cur["name"], f"{base['name']}（G1）",
                         "、".join(extra) or "—", "、".join(missing) or "—",
                         f"{sheet_gap:+d}" if sheet_gap else "0"])
    put(ws, ["组号", "对比文件", "基准文件", "多出的列", "缺失的列", "工作表数差"], rows,
        [8, 30, 30, 24, 24, 12])

    # ④ 跳过清单：跳了谁、为什么
    ws = wb.create_sheet("跳过清单")
    put(ws, ["文件", "原因"], [list(row) for row in skipped], [36, 46])

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    wb.save(LIST_XLSX)


def write_clean(baseline):
    """照现行版的表头，另存一份干净的受控模板：只留标题行 + 表头 + 一行空位。"""
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font
    src = load_workbook(baseline["path"])
    wb = Workbook()
    wb.remove(wb.active)
    for name in baseline["visible"]:
        s_ws = src[name]
        head = find_header([list(r) for r in s_ws.iter_rows(
            max_row=HEADER_SCAN, values_only=True)])
        if head < 0:                                  # 汇总页没有样品编号，按主表的行位取
            head = baseline["header_row"] - 1
        d_ws = wb.create_sheet(name)
        d_ws.cell(row=1, column=1,
                  value=f"{name}（受控版 {baseline['ver']}，生效 "
                        f"{datetime.now():%Y-%m-%d}）").font = Font(bold=True, size=12)
        for idx in range(1, s_ws.max_column + 1):
            s_cell = s_ws.cell(row=head + 1, column=idx)
            d_cell = d_ws.cell(row=2, column=idx, value=s_cell.value)
            if s_cell.has_style:
                d_cell.font = copy(s_cell.font)
                d_cell.fill = copy(s_cell.fill)
                d_cell.border = copy(s_cell.border)
                d_cell.alignment = copy(s_cell.alignment)
            d_ws.column_dimensions[d_cell.column_letter].width = \
                s_ws.column_dimensions[s_cell.column_letter].width
        d_ws.freeze_panes = "A3"
    src.close()
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    wb.save(CLEAN_XLSX)


def verify(groups, skipped):
    """跑完把产出重新打开核一遍，逐项打印 OK，最后给总判定。"""
    from openpyxl import load_workbook
    checks = []
    detail = load_workbook(LIST_XLSX)["文件明细"]
    rows = [list(r) for r in detail.iter_rows(min_row=2, values_only=True)]
    checks.append(("纳入盘点 7 个文件", len(rows) == 7))
    checks.append(("归成 3 组", len(groups) == 3))
    checks.append(("最大一组 5 份", len(groups[0]["files"]) == 5))
    checks.append(("现行候选是 V2.2 那份（不是修改时间最新的副本）",
                   groups[0]["current"]["name"] == "检验记录模板_2025新.xlsx"))
    checks.append(("跳过 4 个", len(skipped) == 4))
    checks.append(("无版本号 1 份", sum(1 for f in groups[0]["files"] if not f["ver"]) == 1))
    checks.append(("命名与版本号矛盾 1 处",
                   sum(1 for g in groups for f in g["files"] if "最终" in f["flag"]) == 1))
    diff = load_workbook(LIST_XLSX)["差异明细"]
    drows = [list(r) for r in diff.iter_rows(min_row=2, values_only=True)]
    b_row = [r for r in drows if r[1].startswith("检验记录模板(最终版)")]
    checks.append(("14 列那版多出 2 列", bool(b_row) and len(b_row[0][3].split("、")) == 2))
    c_row = [r for r in drows if r[1].startswith("检验记录模板（旧）")]
    checks.append(("旧版缺 3 列", bool(c_row) and len(c_row[0][4].split("、")) == 3))
    clean = load_workbook(CLEAN_XLSX)
    ws = clean[clean.sheetnames[0]]
    cols = [c.value for c in ws[2] if c.value]
    checks.append(("受控模板表头 12 列", len(cols) == 12))
    checks.append(("受控模板标题带 V2.2", "V2.2" in str(ws["A1"].value)))
    checks.append(("受控模板只有干净表头，没有示例数据", ws.max_row == 2))
    checks.append(("损坏文件只进跳过清单", any("损坏" in s[0] for s in skipped)))
    checks.append(("隐藏页不算结构（可见仍 2 个）",
                   len(groups[0]["current"]["visible"]) == 2))
    checks.append(("表头压到第 3 行那份也归进主版",
                   any(f["header_row"] == 3 for f in groups[0]["files"])))

    fails = 0
    for name, ok in checks:
        print(f"  {'OK  ' if ok else 'FAIL'} {name}")
        fails += 0 if ok else 1
    return fails


def main():
    ensure_deps()
    log_head()
    t0 = time.time()

    taken, skipped = scan()
    groups = group_files(taken)
    max_ver = max((f["ver"] for f in taken if f["ver"]), key=ver_key, default=None)
    for group in groups:
        for f in group["files"]:
            flags = []
            if any(word in f["name"] for word in HINT_WORDS) and ver_key(f["ver"]) < ver_key(max_ver):
                flags.append(f"文件名带「最终/新」字样，但版本号 {f['ver']} 低于全盘最高 {max_ver}")
            if f["hidden"]:
                flags.append(f"另有 {len(f['hidden'])} 个隐藏页未计入结构")
            if f["header_row"] != 2:
                flags.append(f"表头在第 {f['header_row']} 行")
            f["flag"] = "；".join(flags)

    write_list(groups, skipped, max_ver)
    write_clean(groups[0]["current"])

    print(f"扫描 {len(taken) + len(skipped)} 个文件：纳入盘点 {len(taken)} 个，跳过 {len(skipped)} 个")
    for group in groups:
        cur = group["current"]
        print(f"  {group['no']}  {len(group['files'])} 份  现行候选 {cur['name']}"
              f"（{cur['ver'] or '无版本号'}，{len(cur['visible'])} 个可见工作表，{len(cur['cols'])} 列）")

    fails = verify(groups, skipped)
    verdict = "PASS" if fails == 0 else f"FAIL（{fails} 项）"
    print(f"结果：{verdict}  用时 {time.time() - t0:.1f}s")

    log_tail(time.time() - t0, verdict)
    if "--no-pause" not in sys.argv:
        input("\n按回车关闭窗口……")


if __name__ == "__main__":
    main()
