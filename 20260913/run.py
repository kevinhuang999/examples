# -*- coding: utf-8 -*-
"""引流长文《稳定性考察数据按批号时间点手工汇总，几十张透视表翻到眼花》配套脚本。

核心代码与正文逐字一致，末尾多了自动回读校验，方便确认汇总表里该有的都在。
跨平台：Windows / macOS / Linux 都是一条命令，缺依赖会自动装。
    python run.py
    python run.py --no-pause   # 不暂停（CI 或脚本里用）

打包给别人时，把整个目录拷走即可，路径全部相对定位，无硬编码。
"""
import csv
import socket
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

# 控制台编码随环境变，只加容错，不强行改编码。
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(errors="replace")
    except Exception:
        pass

MIRROR = "https://pypi.tuna.tsinghua.edu.cn/simple"
NO_PAUSE = "--no-pause" in sys.argv


def _pause(msg=""):
    """Windows 双击 .py 时留住窗口，避免一闪而过。"""
    if msg:
        print(msg)
    if NO_PAUSE:
        return
    try:
        if sys.platform == "win32" and sys.stdin and sys.stdin.isatty():
            input("按回车退出")
    except Exception:
        pass


def _ensure(*modules_pkgs):
    """缺哪个装哪个。先直接装，失败再退到 --user（无管理员权限的机器）。"""
    for mod, pkg in modules_pkgs:
        try:
            __import__(mod)
            continue
        except ImportError:
            pass
        print(f"缺 {pkg}，正在用清华源安装（解释器: {sys.executable}）...")
        base = [sys.executable, "-m", "pip", "install", "-i", MIRROR, pkg]
        attempts = [base]
        if sys.prefix == sys.base_prefix:      # 不在虚拟环境里，普通安装可能没权限
            attempts.append(base[:4] + ["--user"] + base[4:])
        for cmd in attempts:
            subprocess.run(cmd, check=False)
            try:
                __import__(mod)
                break
            except ImportError:
                continue
        else:
            _pause(f"自动安装失败。手动执行：\n  {sys.executable} -m pip install -i {MIRROR} {pkg}")
            sys.exit(1)


_ensure(("pandas", "pandas"), ("openpyxl", "openpyxl"))

import pandas as pd

HERE = Path(__file__).resolve().parent
RAW_FILE = HERE / "01_raw_data" / "稳定性考察_仪器导出.csv"      # 仪器导出的明细（长表）
PLAN_FILE = HERE / "01_raw_data" / "考察方案_时间点顺序.csv"     # 时间点的考察先后
OUT_FILE = HERE / "02_output" / "稳定性考察汇总.xlsx"            # 一个批次一个页签
LOG_DIR = HERE / "source"                                       # 运行日志（run_log.txt 就落在 source/ 下）

ID_COLS = ["批号", "时间点", "检项"]
VAL_COL = "检验结果"
NUMERIC_ITEMS = ["有关物质", "含量测定", "干燥失重", "重金属", "炽灼残渣", "微生物限度"]
PASS_WORDS = ("符合", "合格", "通过")

# 检项在报告里的固定顺序（药典体例：性状、鉴别、检查项、含量）。
# 不写死这个顺序，pivot 出来的行按拼音排，pH 值会跑到最前面，报告没法看。
ITEM_ORDER = ["性状", "鉴别", "溶解度", "溶液颜色", "澄清度", "pH 值",
              "有关物质", "干燥失重", "炽灼残渣", "重金属", "微生物限度", "含量测定"]

# 数字类检项的判定界限。含量测定是区间（98.0%~102.0%），其余是上限——
# 只用一张「上限表」套不住区间项，含量 103% 会被判成合格，这是实跑踩出来的。
STD_LIMIT = {"有关物质": (None, 0.5), "干燥失重": (None, 2.0), "重金属": (None, 20.0),
             "炽灼残渣": (None, 0.1), "微生物限度": (None, 1000.0), "含量测定": (98.0, 102.0)}


def read_source(path: Path) -> pd.DataFrame:
    """读仪器导出的明细，丢掉表尾的空行、合计行、说明行。"""
    rows = []
    with open(path, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            batch = (row.get("批号") or "").strip()
            point = (row.get("时间点") or "").strip()
            item = (row.get("检项") or "").strip()
            if not (batch and point and item):
                continue                     # 空行、表尾「合计」、说明行都不是数据
            rows.append({
                "批号": batch, "时间点": point, "检项": item,
                "标准规定": (row.get("标准规定") or "").strip(),
                "检验结果": (row.get("检验结果") or "").strip(),
                "单位": (row.get("单位") or "").strip(),
            })
    # dtype=str 全程按文本读：批号 00903 的前导零一旦被猜成数字就补不回来了
    df = pd.DataFrame(rows, columns=ID_COLS + ["标准规定", VAL_COL, "单位"], dtype=str)
    return df


def read_time_order(path: Path) -> list:
    """读考察方案里时间点的先后顺序——「0 天 / 1 个月 / 3 个月」按字典序排是错的。"""
    order = []
    with open(path, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            point = (row.get("时间点") or "").strip()
            if point and point not in order:
                order.append(point)
    return order


def pivot_one(df: pd.DataFrame, batch: str, order: list) -> pd.DataFrame:
    """把一个批次的明细透成宽表：行是检项，列是时间点，格子里是检验结果。"""
    sub = df[df["批号"] == batch]
    wide = pd.pivot_table(
        sub,
        index="检项",
        columns="时间点",
        values=VAL_COL,
        aggfunc="first",        # 同格子本该只有一个值，重复取样用 first，别用 sum 把数字加两遍
        dropna=False,
        observed=False,
    )
    wide = wide.reindex(columns=[p for p in order if p in wide.columns])
    wide = wide.reindex(index=[i for i in ITEM_ORDER if i in wide.index])   # 按报告体例排行
    std = sub.drop_duplicates("检项").set_index("检项")
    wide = wide.reset_index()                       # 检项从索引变回第一列
    # 标准规定和单位是每个检项固定的属性，贴着检项放，读者不用左右拉
    wide.insert(1, "标准规定", std["标准规定"].reindex(wide["检项"]).fillna("").to_numpy())
    wide.insert(2, "单位", std["单位"].reindex(wide["检项"]).fillna("").to_numpy())
    # 空结果留空格子，别让 NaN 三个字母出现在报告里
    wide = wide.fillna("")
    return wide


def find_over(wide: pd.DataFrame) -> list:
    """扫一遍宽表，挑出数字越界的格子，返回 (检项, 时间点) 列表。

    界限写成一个元组 (下限, 上限)，哪头是 None 就只管另一头——
    含量测定这类区间项只判上限会把 103% 放过去。
    """
    over = []
    for _, row in wide.iterrows():
        item = row["检项"]
        if item not in STD_LIMIT:
            continue
        low, high = STD_LIMIT[item]
        for point in wide.columns[3:]:              # 前三列是检项/标准规定/单位
            v = str(row[point]).strip()
            if not v:
                continue
            try:
                num = float(v)
            except ValueError:
                continue                            # 文字结果（符合/白色粉末）不走这条
            if (low is not None and num < low) or (high is not None and num > high):
                over.append((item, point))
    return over


def write_output(df: pd.DataFrame, order: list, out_file: Path) -> dict:
    """一个批次一个页签，写进同一个 xlsx；页签名 = 批号，超标的格子刷红。"""
    from openpyxl import load_workbook
    from openpyxl.styles import PatternFill

    out_file.parent.mkdir(parents=True, exist_ok=True)
    made, over_map = {}, {}
    with pd.ExcelWriter(out_file, engine="openpyxl") as writer:
        for batch in df["批号"].drop_duplicates():
            wide = pivot_one(df, batch, order)
            wide.to_excel(writer, sheet_name=str(batch), index=False)
            made[batch] = wide
            over_map[str(batch)] = find_over(wide)

    # pandas 只能写值，刷颜色得再开一次工作簿
    red = PatternFill("solid", fgColor="FFC7CE")
    wb = load_workbook(out_file)
    for batch, cells in over_map.items():
        ws = wb[batch]
        head = {ws.cell(row=1, column=c).value: c for c in range(1, ws.max_column + 1)}
        for item, point in cells:
            for r in range(2, ws.max_row + 1):
                if ws.cell(row=r, column=head["检项"]).value == item:
                    ws.cell(row=r, column=head[point]).fill = red
                    break
    wb.save(out_file)
    return made, over_map


def main() -> bool:
    df = read_source(RAW_FILE)
    order = read_time_order(PLAN_FILE)
    made, over_map = write_output(df, order, OUT_FILE)
    print(f"读到 {len(df)} 行明细，{df['批号'].nunique()} 个批次，"
          f"时间点顺序 {' / '.join(order)}")
    for batch, wide in made.items():
        print(f"  {batch}  {len(wide)} 个检项 × {len(wide.columns) - 2} 个时间点")

    # ---------- 以下为校验，正文里没有 ----------
    print("\n===== 回读校验 =====")
    back = pd.read_excel(OUT_FILE, sheet_name=None, dtype=str, keep_default_na=False)
    long_batch = [b for b, w in made.items() if len(w) == 12][0]
    check = [
        ("5 个批次都落了页签", list(back.keys()) == list(made.keys()), list(back.keys())),
        ("批号前导零没丢", "00903" in back, [k for k in back if "903" in k]),
        ("时间点按考察顺序排，不是字典序",
         list(back["240905"].columns[3:]) == order,
         list(back["240905"].columns[3:])),
        ("检项做成了行、时间点做成了列",
         list(back["241012"].columns[:3]) == ["检项", "标准规定", "单位"],
         list(back["241012"].columns[:3])),
        ("少的那个时间点没有凭空多一列（241012 只有 0 天）",
         list(back["241012"].columns[3:]) == ["0 天"],
         list(back["241012"].columns[3:])),
        ("12 个检项批次一行不少", len(back[long_batch]) == 12, len(back[long_batch])),
        ("检项按报告体例排，不是拼音序",
         list(back[long_batch]["检项"])[:3] == ["性状", "鉴别", "溶解度"],
         list(back[long_batch]["检项"])[:3]),
        ("空结果没被填成 0 或 NaN",
         back["240906"].isna().to_numpy().sum() == 0
         and (back["240906"] == "").to_numpy().sum() >= 1,
         f"空格子 {(back['240906'] == '').to_numpy().sum()} 处，NaN {int(back['240906'].isna().to_numpy().sum())} 处"),
        ("超标格被刷红（241115 有关物质 6 个月）",
         over_map["241115"] == [("有关物质", "6 个月")], over_map["241115"]),
        ("标准规定保留在检项旁边", "≤ 0.5%" in back["240905"].to_numpy().__str__(), "≤ 0.5%"),
        ("超标那格是 0.63", "0.63" in back["241115"].to_numpy().__str__(), "0.63"),
    ]
    ok = True
    for name, passed, val in check:
        print(f"  [{'OK ' if passed else 'FAIL'}] {name}  ->  {val}")
        ok &= passed

    print("\n总判定:", "PASS" if ok else "FAIL")
    return bool(ok)


class _Tee:
    """同时写控制台和日志文件，万一控制台中文乱码还有日志可看。"""

    def __init__(self, *streams):
        self.streams = streams

    def write(self, s):
        for st in self.streams:
            try:
                st.write(s)
            except Exception:
                pass

    def flush(self):
        for st in self.streams:
            try:
                st.flush()
            except Exception:
                pass


def _note_unfinished():
    """上一条执行若被手动关窗口打断，日志里会缺完成行。开跑前补一句说明。"""
    LOG_FILE = LOG_DIR / "run_log.txt"
    if not LOG_FILE.exists():
        return
    tail = ""
    with open(LOG_FILE, encoding="utf-8", errors="replace") as f:
        for ln in f:
            if ln.strip():
                tail = ln.strip()
    if tail and not tail.startswith("执行完成"):
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write("（上一条没有留下完成时间，多半是窗口被手动关闭或中途强制中断）\n")


def _host_info() -> str:
    """主机名 + 本机 IP，用来分辨这段日志是哪台机器跑出来的。"""
    try:
        host = socket.gethostname()
    except Exception:
        host = "unknown"
    try:
        # UDP 不会真的发包，只按路由表取本机对外网卡的地址
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as s:
            s.connect(("8.8.8.8", 80))
            ip = s.getsockname()[0]
    except Exception:
        try:
            ip = socket.gethostbyname(host)
        except Exception:
            ip = "未知"
    return f"{host}（{ip}）"


def _run():
    """一次执行 = 追加一段：开始时间 + 过程 + 完成时间。历史日志不覆盖。"""
    LINE = "=" * 60
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    _note_unfinished()
    _log = open(LOG_DIR / "run_log.txt", "a", encoding="utf-8")   # 追加，别用 w
    _log.write(f"\n{LINE}\n")
    _log.write(f"运行主机：{_host_info()}\n")
    _log.write(f"开始执行：{datetime.now():%Y-%m-%d %H:%M:%S}\n")
    _log.write(f"Python   ：{sys.version.split()[0]}  （{Path(sys.executable).name}）\n")
    _log.flush()

    _orig = sys.stdout
    sys.stdout = _Tee(sys.stdout, _log)
    t0, ok, tail = time.time(), False, ""
    try:
        try:
            ok = main()
        except KeyboardInterrupt:
            tail = "\n被手动中断（Ctrl+C）。"
        except Exception as e:
            import traceback
            traceback.print_exc()
            tail = f"\n跑失败了：{type(e).__name__}: {e}"
        else:
            tail = "\n跑完了。结果在 02_output/，运行记录已写入 source/run_log.txt。"
        if tail:
            print(tail)
        # 完成行必须写在暂停之前：双击后关窗口就丢不了
        print(f"执行完成：{datetime.now():%Y-%m-%d %H:%M:%S}"
              f"  用时 {time.time() - t0:.1f}s  判定 {'PASS' if ok else 'FAIL'}")
    finally:
        sys.stdout.flush()
        _log.close()
        sys.stdout = _orig

    _pause()          # 等回车放在日志收尾之后
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    _run()
