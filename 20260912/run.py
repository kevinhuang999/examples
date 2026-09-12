# -*- coding: utf-8 -*-
"""引流长文《检验报告一份份另存为 PDF 太慢，Python 批量生成能行吗》配套脚本。

核心代码与正文逐字一致，末尾多了自动回读校验，方便确认 PDF 里该有的都在。
跨平台：Windows / macOS / Linux 都是一条命令，缺依赖会自动装。
    python run.py
    python run.py --no-pause   # 不暂停（CI 或脚本里用）

打包给别人时，把整个目录拷走即可，路径全部相对定位，无硬编码。
"""
import csv
import re
import socket
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path
from xml.sax.saxutils import escape

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


_ensure(("reportlab", "reportlab"), ("pypdf", "pypdf"))

from pypdf import PdfReader
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

HERE = Path(__file__).resolve().parent
RAW_FILE = HERE / "01_raw_data" / "检验结果_仪器导出.csv"   # 仪器导出的明细
OUT_DIR = HERE / "02_output"                               # 一份样品一个 PDF
LOG_DIR = HERE / "source"                                  # 运行日志（run_log.txt 就落在 source/ 下）

FIELDS = ["样品编号", "样品名称", "批号", "检验项目", "标准规定",
          "检验结果", "单位", "单项结论", "检测人", "检测日期"]
DETAIL_COLS = ["检验项目", "标准规定", "检验结果", "单位", "单项结论"]

# 两张表共用一个总宽，左右边框才对得齐（A4 默认版心约 159mm，别超）
TABLE_W = 158 * mm
INFO_W = [22 * mm, 57 * mm, 22 * mm, 57 * mm]        # 抬头表：4 列加起来 = TABLE_W
DETAIL_W = [40 * mm, 40 * mm, 30 * mm, 24 * mm, 24 * mm]   # 明细表：5 列加起来 = TABLE_W

# 各系统的中文字体位置不一样，逐个试着注册，谁在就用谁
FONT_CANDIDATES = [
    ("C:/Windows/Fonts/msyh.ttc", 0),
    ("C:/Windows/Fonts/simhei.ttf", 0),
    ("/System/Library/Fonts/PingFang.ttc", 0),
    ("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc", 0),
    ("/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc", 0),
]


def setup_font() -> str:
    """注册一个能写汉字的中文字体，返回注册后的字体名。"""
    for path, index in FONT_CANDIDATES:
        if Path(path).exists():
            try:
                pdfmetrics.registerFont(TTFont("CJK", path, subfontIndex=index))
                return "CJK"
            except Exception:
                continue
    raise FileNotFoundError("找不到可用的中文字体文件，装一个再跑：" + "、".join(p for p, _ in FONT_CANDIDATES))


def read_source(path: Path) -> list:
    """逐行读仪器导出的 CSV，只留下真正的检验数据行。"""
    rows = []
    with open(path, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            sample = (row.get("样品编号") or "").strip()
            item = (row.get("检验项目") or "").strip()
            date = (row.get("检测日期") or "").strip()
            if not (sample and item and date):   # 空行、表尾「合计」行、说明行都不是数据
                continue
            rows.append({k: (row.get(k) or "").strip() for k in FIELDS})
    return rows


def group_by_sample(rows: list) -> dict:
    """按样品编号归堆，一个样品一份报告。"""
    groups = {}
    for r in rows:
        groups.setdefault(r["样品编号"], []).append(r)
    return groups


def make_story(sample_id: str, items: list, font: str) -> list:
    """把一个样品的明细拼成报告内容（抬头 + 明细表 + 结论），还没落成文件。"""
    title = ParagraphStyle("title", fontName=font, fontSize=16, leading=22, alignment=1)
    label = ParagraphStyle("label", fontName=font, fontSize=9, leading=13, textColor=colors.HexColor("#555555"))
    body = ParagraphStyle("body", fontName=font, fontSize=9, leading=13)
    head = items[0]
    story = [
        Paragraph("检 验 报 告", title),
        Spacer(1, 6 * mm),
        Table([["样品名称", head["样品名称"], "批号", head["批号"]],
               ["检测人", head["检测人"], "检测日期", head["检测日期"]]],
              colWidths=INFO_W,
              style=TableStyle([
                  ("FONTNAME", (0, 0), (-1, -1), font),
                  ("FONTSIZE", (0, 0), (-1, -1), 9),
                  ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#999999")),
                  ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#F2F2F2")),
                  ("BACKGROUND", (2, 0), (2, -1), colors.HexColor("#F2F2F2")),
                  ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
              ])),
        Spacer(1, 5 * mm),
    ]
    data = [DETAIL_COLS] + [[item[c] for c in DETAIL_COLS] for item in items]
    data = [[Paragraph(escape(cell), body) for cell in row] for row in data]
    table = Table(data, colWidths=DETAIL_W, repeatRows=1)
    table.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), font),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#999999")),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E8EEF7")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (2, 1), (4, -1), "CENTER"),
    ]))
    story.append(table)
    bad = [item["检验项目"] for item in items if item["单项结论"] != "符合"]
    conclusion = "结论：符合规定。" if not bad else "结论：不符合规定（" + "、".join(bad) + "）。"
    story += [Spacer(1, 5 * mm), Paragraph(conclusion, body), Paragraph(f"样品编号：{sample_id}", label)]
    return story


def write_output(groups: dict, out_dir: Path, font: str) -> list:
    """一个样品生成一份 PDF，存到 02_output/。"""
    out_dir.mkdir(parents=True, exist_ok=True)
    made = []
    for sample_id, items in groups.items():
        safe = re.sub(r'[\\/:*?"<>|]', "_", sample_id)     # 编号里带斜杠，不换掉写不出文件
        out = out_dir / f"检验报告_{safe}.pdf"
        doc = SimpleDocTemplate(str(out), pagesize=A4, title=f"检验报告 {sample_id}")
        doc.build(make_story(sample_id, items, font))
        made.append(out)
    return made


def main() -> bool:
    font = setup_font()
    rows = read_source(RAW_FILE)
    groups = group_by_sample(rows)
    made = write_output(groups, OUT_DIR, font)
    print(f"读到 {len(rows)} 行明细，归并成 {len(groups)} 份报告")
    for pdf in made:
        print(f"  {pdf.name}  {pdf.stat().st_size // 1024} KB")

    # ---------- 以下为校验，正文里没有 ----------
    print("\n===== 回读校验 =====")
    pages, texts = {}, {}
    for pdf in made:
        reader = PdfReader(str(pdf))
        pages[pdf.name] = len(reader.pages)
        texts[pdf.name] = [pg.extract_text() or "" for pg in reader.pages]
    t_all = "\n".join(t for pgs in texts.values() for t in pgs)
    full = [n for n in texts if "2026-0912-004" in n][0]
    check = [
        ("生成 4 份 PDF", len(made) == 4, [p.name for p in made]),
        ("带斜杠的编号落成了文件名", any("S_2026_0912_003" in n for n in texts), "S_2026_0912_003"),
        ("汉字没变成方块（标题文字可提取）", "检 验 报 告" in t_all, "检 验 报 告" in t_all),
        ("34 个检项的报告跨了页", pages[full] >= 2, f"{pages[full]} 页"),
        ("第二页的表头也重复了", "检验项目" in texts[full][1] if pages[full] >= 2 else False,
         [c for c in DETAIL_COLS if pages[full] >= 2 and c in texts[full][1]]),
        ("标准规定里的 <0.5 原样保留", "<0.5 EU/ml" in t_all, "<0.5 EU/ml" in t_all),
        ("样品名里的 & 没被改字", "PW&WFI" in t_all, [l for l in t_all.splitlines() if "PW" in l]),
        ("空的结果没显示成 None", "None" not in t_all, f"None 出现 {t_all.count('None')} 次"),
        ("表格数对得上（4+2+3+34）", sum(len(g) for g in groups.values()) == 43,
         [len(g) for g in groups.values()]),
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
