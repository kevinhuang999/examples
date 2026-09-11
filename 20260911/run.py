# -*- coding: utf-8 -*-
"""引流长文《仪器导出的 Excel 要一个一个复制汇总，怎么用 Python 自动合并》配套脚本。

核心代码与正文逐字一致，末尾多了自动回读校验，方便确认合并结果对不对。
跨平台：Windows / macOS / Linux 都是一条命令，缺依赖会自动装。
    python run.py
    python run.py --no-pause   # 不暂停（CI 或脚本里用）

打包给别人时，把整个目录拷走即可，路径全部相对定位，无硬编码。
"""
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


_ensure(("openpyxl", "openpyxl"), ("pandas", "pandas"))

import pandas as pd

HERE = Path(__file__).resolve().parent
RAW_DIR = HERE / "01_raw_data"                           # 各仪器导出的 xlsx
OUT_DIR = HERE / "02_output"                             # 合并结果
LOG_DIR = HERE / "source"                                # 运行日志（run_log.txt 就落在 source/ 下）

SUMMARY_SHEET = "合并明细"
TOTAL_SHEET = "按项目汇总"


def read_source(raw_dir: Path) -> list:
    """找到目录里全部 xlsx，逐个读进 DataFrame 并记下来源文件。"""
    files = sorted(raw_dir.glob("*.xlsx"))
    if not files:
        raise FileNotFoundError(f"{raw_dir.name}/ 下没有 xlsx，先把仪器导出文件放进来")
    frames = []
    for f in files:
        df = pd.read_excel(f, header=None)      # 表头在哪行未知，先当纯数据读
        frames.append(df)
    return list(zip(files, frames))


def align_headers(pairs: list) -> list:
    """各文件表头行不同，逐个找『样品编号』所在行，把它上面的行全丢掉。"""
    aligned = []
    for f, df in pairs:
        col0 = df[0].astype(str)
        hit = col0[col0 == "样品编号"].index
        if hit.empty:
            raise ValueError(f"{f.name} 里找不到表头行（第 1 列没有『样品编号』）")
        header_idx = hit[0]
        # 批号按文本读：这是字符串标识不是数字，前导零不能丢
        df = pd.read_excel(f, header=header_idx, dtype={"批号": str})
        df["来源文件"] = f.name
        aligned.append(df)
    return aligned


def merge_and_summarize(aligned: list) -> tuple:
    """全部纵向拼成一张总表；检验结果列先转数值再按检验项目求平均。"""
    merged = pd.concat(aligned, ignore_index=True)
    numeric = pd.to_numeric(merged["检验结果"], errors="coerce")   # '符合' 会变 NaN
    merged["数值结果"] = numeric
    summary = (merged.dropna(subset=["数值结果"])
                     .groupby("检验项目", as_index=False)["数值结果"]
                     .mean())
    return merged, summary


def write_output(merged, summary, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    out = out_dir / "检验结果_汇总.xlsx"
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        merged.to_excel(writer, sheet_name=SUMMARY_SHEET, index=False)
        summary.to_excel(writer, sheet_name=TOTAL_SHEET, index=False)
    return out


def main() -> bool:
    pairs = read_source(RAW_DIR)
    aligned = align_headers(pairs)
    merged, summary = merge_and_summarize(aligned)
    out = write_output(merged, summary, OUT_DIR)
    print(f"已合并 {len(aligned)} 个文件，共 {len(merged)} 行明细")
    print(f"汇总表已生成 {out.name}，数值检项 {len(summary)} 项")

    # ---------- 以下为校验，正文里没有 ----------
    print("\n===== 回读校验 =====")
    book = pd.read_excel(out, sheet_name=None)
    detail = book[SUMMARY_SHEET]
    total = book[TOTAL_SHEET]
    row_003 = detail[detail["批号"].astype(str) == "00903"]
    yb = total[total["检验项目"] == "含量"]["数值结果"].iloc[0]
    check = [
        ("合并行数 = 8（合计行 0 进来）", len(detail) == 8, len(detail)),
        ("三份文件都有来源列", detail["来源文件"].nunique() == 3, sorted(detail["来源文件"].unique())),
        ("批号 00903 前导零还在", len(row_003) == 2 and row_003["批号"].iloc[0] == "00903",
         row_003["批号"].tolist()),
        ("'符合' 没混进平均值（汇总 6 项）", len(total) == 6, total["检验项目"].tolist()),
        ("含量均值 = (0.982+0.975)/2 = 0.9785", abs(yb - 0.9785) < 1e-9, yb),
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
