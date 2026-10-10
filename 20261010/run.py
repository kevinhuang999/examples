# -*- coding: utf-8 -*-
"""仪器导出结果文件 → 按样品号归堆、匹配、留痕，出一张可入库的清单表 —— 可运行示例

场景：实验室的色谱工作站（CDS）按批次导出结果文件，命名规则不一致——有的文件名
带实验室样品号，有的只有时间戳；样品号要靠文件里的「进样瓶位」按一张对照表反查。
本示例遍历导出目录，把每个结果行的样品号认出来，按样品归堆，另外出一份留痕表
和一份挂起清单（匹配不上 / 同一份文件混多个样品 / 同一样品多条结果 / 结果非数值）。

用法：
    python run.py              # 跑完暂停（双击也能用）
    python run.py --no-pause   # 不暂停（脚本 / CI 里跑）
"""
import argparse
import csv
import io
import re
import socket
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

BASE = Path(__file__).resolve().parent
RAW_DIR = BASE / "01_raw_data"
OUT_DIR = BASE / "02_output"
SRC_DIR = BASE / "source"
TPL_DIR = SRC_DIR / "templates"
MAP_FILE = TPL_DIR / "样品对照表.csv"
LOG_FILE = SRC_DIR / "run_log.txt"

DATA_SUFFIX = {".csv", ".xlsx", ".xls"}
ID_COL = "进样瓶位"            # 仪器标识所在的列（文件名认不出来时靠它反查样品号）
FILE_ID_RE = re.compile(r"(\d{6,})")   # 文件名里 6 位以上的数字段，视为候选样品号
OUT_NAME = "仪器结果归堆表.xlsx"
NUM_COLS = ["保留时间", "峰面积", "峰高", "结果", "含量(%)"]


def ensure_deps():
    """缺依赖自动装（清华源）；已装则直接返回。"""
    need = []
    for mod, pkg in (("pandas", "pandas"), ("openpyxl", "openpyxl")):
        try:
            __import__(mod)
        except ImportError:
            need.append(pkg)
    if not need:
        return
    print("缺少依赖：", ", ".join(need), "，正在从清华源安装……")
    cmd = [sys.executable, "-m", "pip", "install", "-i",
           "https://pypi.tuna.tsinghua.edu.cn/simple", *need]
    if subprocess.call(cmd) != 0:
        subprocess.call(cmd + ["--user"])


def norm(name):
    """列名归一：去空白、全角括号转半角、英文转小写。

    不同仪器/不同导出模板对同一列的写法会飘（`进样瓶位` / `进样瓶位 ` / `Vial`），
    匹配前先归一，而不是要求导出的列名一尘不染。
    """
    s = str(name) if name is not None else ""
    s = s.replace("（", "(").replace("）", ")")
    s = re.sub(r"[\s\u00a0\u3000]+", "", s)
    return s.lower()


def load_sample_map(path):
    """读样品对照表，返回 (仪器标识 → 样品号, 仪器标识 → 检测项目)。

    对照表是业务资产：写在表里、不写死在代码里——仪器换了、编号规则改了，改表不改代码。
    """
    with path.open(encoding="utf-8-sig", newline="") as f:
        rows = [r for r in csv.reader(f) if any(str(c).strip() for c in r)]
    id2no, id2item = {}, {}
    for r in rows[1:]:
        ident = str(r[0]).strip()
        no = str(r[1]).strip() if len(r) > 1 else ""
        if not ident or not no:
            continue
        id2no[norm(ident)] = no
        id2item[norm(ident)] = str(r[2]).strip() if len(r) > 2 else ""
    return id2no, id2item


def read_text(path):
    """按 utf-8-sig → gbk → utf-8 依次试；GBK 文件用 utf-8 读会直接抛异常，不是读出乱码。"""
    last = None
    for enc in ("utf-8-sig", "gbk", "utf-8"):
        try:
            return path.read_text(encoding=enc), enc
        except UnicodeDecodeError as e:
            last = e
    raise last


def read_matrix(path):
    """把一个导出文件读成二维表（不做任何解释），并返回实际用的编码。"""
    if path.suffix.lower() == ".csv":
        text, enc = read_text(path)
        return list(csv.reader(io.StringIO(text))), enc
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    return rows, "xlsx"


def cell(cells, i):
    return cells[i] if (i is not None and i < len(cells)) else ""


def normalize_value(v):
    """单元格 → 去首尾空白的字符串；空值一律返回 ''。"""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return re.sub(r"^[\s\u00a0\u3000]+|[\s\u00a0\u3000]+$", "", str(v))


def find_header(rows):
    """找表头行：第一个「非空单元格 >= 2」的行。

    不能写死 header=0——CDS 导出常在表头上压几行导出信息；
    也不能靠行号猜，模板改一版行号就全错。
    """
    for i, r in enumerate(rows):
        if sum(1 for c in r if str(c or "").strip()) >= 2:
            return i
    return -1


def pick_columns(headers):
    """按归一后的列名，把「进样瓶位 / 组分 / 各项结果」的列下标挑出来。

    认不出的列一律丢掉（不猜），但每个结果列都按名字进表，方便入库时对齐字段。
    """
    idx = {}
    for i, h in enumerate(headers):
        key = norm(h)
        if not key:
            continue
        if key in ("进样瓶位", "瓶位", "vial", "vialposition"):
            idx[ID_COL] = i
        elif key in ("组分", "component", "名称", "峰名称", "化合物", "检测项目"):
            idx["组分"] = i
        elif key in NUM_COLS:
            idx.setdefault(key, i)
    return idx


def sample_no_from_name(fname, id2no):
    """先从文件名里找样品号：文件名里的数字段直接命中对照表的样品号即采用。

    文件名认不出来不算错——只是这一条走「按瓶位反查」的路子，不终止解析。
    """
    known = set(id2no.values())
    for m in FILE_ID_RE.findall(fname):
        if m in known:
            return m
    return ""


def classify(rows, hdr, cols, fname, id2no, id2item, name_no):
    """切数据行，逐行认样品号，返回 (结果行列表, 挂起列表)。

    认号优先级：文件名里的样品号 > 瓶位反查。
    文件名已经认出来了，这一份文件整批就是同一个样品，逐行不再挑瓶位——
    老仪器导出的表干脆没有瓶位列，硬要求它有，整份文件就白白挂起了。
    """
    data, suspended = [], []
    for k, r in enumerate(rows[hdr + 1:]):
        cells = [normalize_value(c) for c in r]
        if not any(cells):
            continue
        pos = k + hdr + 2               # Excel 行号（1 起算，表头占 hdr+1 行）
        ident = cell(cells, cols.get(ID_COL)).strip()
        no = name_no or id2no.get(norm(ident), "")
        if not no:
            suspended.append((fname, f"第 {pos} 行", f"样品号匹配不上（瓶位「{ident}」不在对照表）"))
            continue
        row = {"样品号": no, "检测项目": id2item.get(norm(ident), ""),
               "仪器标识": ident or fname, "来源文件": fname}
        for name, i in cols.items():
            if name in (ID_COL,):
                continue
            row[name] = cell(cells, i)
        data.append(row)
    return data, suspended


def collect(id2no, id2item):
    """遍历 01_raw_data，逐个文件读、认样品号、归堆。"""
    data, trace, suspended, files = [], [], [], []
    for path in sorted(RAW_DIR.iterdir()):
        if path.suffix.lower() not in DATA_SUFFIX:
            continue                   # 说明文件之类不参与
        files.append(path.name)
        rows, enc = read_matrix(path)
        hdr = find_header(rows)
        if hdr < 0:
            suspended.append((path.name, "整个文件", "找不到表头（非空单元格不足 2 个）"))
            continue
        cols = pick_columns(rows[hdr])
        name_no = sample_no_from_name(path.name, id2no)
        # 两条路都走不通才挂起整份文件：文件名认不出、且表里也没有能反查的瓶位列
        if not name_no and ID_COL not in cols:
            suspended.append((path.name, "整个文件",
                              f"文件名认不出样品号，表里也没有「{ID_COL}」列"))
            continue
        got, susp = classify(rows, hdr, cols, path.name, id2no, id2item, name_no)
        trace.append((path.name, enc, hdr + 1, len(got), len(susp),
                      "文件名认号" if name_no else "按瓶位反查"))
        data.extend(got)
        suspended.extend(susp)
    return data, trace, suspended, files


def flag_dupes(data):
    """同一样品号 + 同一检测项目出现多行时标记出来（复检 / 平行样），但不擅自取数。"""
    from collections import Counter
    counter = Counter((r["样品号"], r.get("组分", "")) for r in data)
    for r in data:
        key = (r["样品号"], r.get("组分", ""))
        if counter[key] > 1:
            r["重复标记"] = f"同一样品同一组分 {counter[key]} 条"
        else:
            r["重复标记"] = ""
    return data


def write_outputs(data, trace, suspended):
    """归堆表 + 认号留痕 + 挂起清单，三份都落在 02_output/。"""
    import pandas as pd
    from openpyxl import load_workbook
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    cols = ["样品号", "检测项目", "仪器标识", "组分",
            "保留时间", "峰面积", "峰高", "结果", "含量(%)", "重复标记", "来源文件"]
    df = pd.DataFrame([{c: r.get(c, "") for c in cols} for r in data], columns=cols)
    out = OUT_DIR / OUT_NAME
    df.to_excel(out, index=False, sheet_name="归堆结果")
    wb = load_workbook(out)
    ws = wb["归堆结果"]
    for j, name in enumerate(cols, 1):      # 文本列设成文本格式，前导零才不会被吃掉
        if name in NUM_COLS:
            continue
        for i in range(2, ws.max_row + 1):
            ws.cell(row=i, column=j).number_format = "@"
    wb.save(out)
    pd.DataFrame(trace, columns=["来源文件", "编码", "表头行", "认出结果行", "挂起行", "认号方式"]).to_csv(
        OUT_DIR / "认号留痕.csv", index=False, encoding="utf-8-sig")
    pd.DataFrame(suspended, columns=["来源文件", "位置", "原因"]).to_csv(
        OUT_DIR / "挂起清单.csv", index=False, encoding="utf-8-sig")
    return out


def verify(data, trace, suspended):
    """生成后回读：逐项断言，打印 OK / FAIL，返回总判定。"""
    import pandas as pd
    from openpyxl import load_workbook
    out = OUT_DIR / OUT_NAME
    df = pd.read_excel(out, dtype=str) if out.exists() else pd.DataFrame()
    checks = [("归堆表已生成", out.exists() and out.stat().st_size > 0)]
    if not df.empty:
        checks.append(("归堆行数 = 12", len(df) == 12))
        checks.append(("样品号列前导零保留（0000182 仍是文本）",
                       "0000182" in set(df["样品号"].astype(str))))
        checks.append(("文件名认号命中（0000190，表里无瓶位列）",
                       "0000190" in set(df["样品号"].astype(str))))
        checks.append(("按瓶位反查命中（0000183 来自时间戳文件）",
                       "0000183" in set(df["样品号"].astype(str))))
        checks.append(("一份文件混两样品已拆开",
                       set(df[df["来源文件"] == "CDS_SEQ_20261010_A.csv"]["样品号"]) == {"0000184", "0000185"}))
        checks.append(("「未检出」原样保留", "未检出" in set(df["结果"].astype(str))))
        checks.append(("「<0.05」原样保留", "<0.05" in set(df["结果"].astype(str))))
        checks.append(("重复标记命中 2 行（复检 0000187）",
                       len(df[df["重复标记"].astype(str).str.contains("2 条", na=False)]) == 2))
    checks.append(("挂起 = 1 条（V99-Z9 匹配不上）", len(suspended) == 1
                   and suspended[0][0] == "CDS_20261010_unknown.csv"))
    checks.append(("认号留痕 6 条（说明文件不计）", len(trace) == 6))
    if out.exists():
        ws = load_workbook(out)["归堆结果"]
        checks.append(("样品号列已设文本格式", ws.cell(row=2, column=1).number_format == "@"))
    ok = True
    for name, passed in checks:
        print(("  OK   " if passed else "  FAIL ") + name)
        ok = ok and bool(passed)
    return ok


def host_info():
    """主机名 + 本机对外 IP，用来分辨这段日志是哪台机器跑的。"""
    name = socket.gethostname()
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
    except OSError:
        try:
            ip = socket.gethostbyname(name)
        except OSError:
            ip = "未知"
    return name, ip


def append_log(text):
    """日志追加不覆盖；每段开头记运行主机，方便分辨机器。"""
    with LOG_FILE.open("a", encoding="utf-8") as f:
        f.write(text)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--no-pause", action="store_true", help="跑完不暂停")
    args = ap.parse_args()

    ensure_deps()
    if not MAP_FILE.exists():
        print("找不到样品对照表：", MAP_FILE)
        print("先跑 source/build_fixtures.py 重造模板与数据。")
        return 1
    SRC_DIR.mkdir(parents=True, exist_ok=True)

    t0 = time.time()
    host, ip = host_info()
    append_log(f"\n运行主机：{host}（{ip}）\n开始执行：{datetime.now():%Y-%m-%d %H:%M:%S}\n"
               f"Python {sys.version.split()[0]}（{Path(sys.executable).name}）\n")

    print("=" * 62)
    print("仪器导出结果文件 → 按样品号归堆")
    print("=" * 62)
    id2no, id2item = load_sample_map(MAP_FILE)
    print(f"样品对照表：{len(id2no)} 条仪器标识 → 样品号")

    data, trace, suspended, files = collect(id2no, id2item)
    data = flag_dupes(data)
    print(f"扫描文件 {len(files)} 个 → 归堆 {len(data)} 行 / 挂起 {len(suspended)} 条")

    out = write_outputs(data, trace, suspended)
    print(f"已生成：{out.name}")
    ok = verify(data, trace, suspended)

    used = time.time() - t0
    verdict = "PASS" if ok else "FAIL"
    print(f"\n总判定：{verdict}（用时 {used:.1f}s）")
    append_log(f"执行完成：{datetime.now():%Y-%m-%d %H:%M:%S}  用时 {used:.1f}s  判定 {verdict}\n")
    if not args.no_pause:
        input("\n按回车键退出……")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
