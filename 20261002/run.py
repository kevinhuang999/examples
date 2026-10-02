# -*- coding: utf-8 -*-
"""批量导出 Excel 到 PDF：按规则表逐张工作表导出，源工作簿一行都不改。

用法：
    python run.py              # 跑完暂停，等回车
    python run.py --no-pause   # 无人值守用

做的事：读 01_raw_data/导出规则.csv → 打开源工作簿（只读）→ 按规则排版成 PDF →
写进 02_output/ → 顺手把每份的页数/字节数登记成一本台账 → 回读校验。
"""
import subprocess
import sys


def ensure(*pkgs):
    """缺依赖就自动装（本机/客户机都可能没有），走清华源。"""
    for mod, pkg in pkgs:
        try:
            __import__(mod)
        except ImportError:
            print(f"缺少 {pkg}，正在用清华源安装 ...")
            subprocess.check_call([sys.executable, "-m", "pip", "install",
                                   "-i", "https://pypi.tuna.tsinghua.edu.cn/simple", pkg])


ensure(("openpyxl", "openpyxl"), ("reportlab", "reportlab"), ("pypdf", "pypdf"))

import csv
import hashlib
import os
import re
import shutil
import socket
import time
from datetime import date, datetime
from functools import partial
from pathlib import Path
from xml.sax.saxutils import escape

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT

# ---------------------------------------------------------------------------
# ① 配置路径与字段：只改这一区，下面的函数不用动
# ---------------------------------------------------------------------------
BASE = Path(__file__).resolve().parent          # 零硬编码：脚本在哪，包根就在哪
RAW = BASE / "01_raw_data"                      # 源工作簿 + 导出规则
OUT = BASE / "02_output"                        # 导出的 PDF + 导出台账
RULES = RAW / "导出规则.csv"                     # 一行一个导出任务
LOG = BASE / "source" / "run_log.txt"           # 追加写，不覆盖

PAGE_SIZES = {"纵向": A4, "横向": landscape(A4)}
MARGIN_L = MARGIN_R = 18 * mm
MARGIN_T = 15 * mm
MARGIN_B = 20 * mm                              # 留出页脚的位置
DEFAULT_COL_W = 8.43                            # Excel 默认列宽（字符数）
KEY_COL = 1                                     # 判空列：第一列留空 = 数据区结束
PLACEHOLDER = re.compile(r"\{源文件\}|\{工作表\}|\{导出日\}")
ILLEGAL = re.compile(r'[\\/:*?"<>|\r\n\t]+')
FALLBACK_NAME = "{源文件}_{工作表}"

FONT_NAME = "LabFont"
FONT_CANDIDATES = [
    r"C:\Windows\Fonts\simhei.ttf",
    r"C:\Windows\Fonts\msyh.ttc",
    r"C:\Windows\Fonts\simsun.ttc",
    "/System/Library/Fonts/PingFang.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
]


def register_font():
    """中文字体必须显式注册，否则整页是方块，而且不报错。"""
    for path in FONT_CANDIDATES:
        if Path(path).exists():
            try:
                pdfmetrics.registerFont(TTFont(FONT_NAME, path))
                return path
            except Exception:
                continue
    raise RuntimeError("没找到可用的中文字体，请在 FONT_CANDIDATES 里补一个本机字体路径")


FONT_PATH = register_font()
S_HEAD = ParagraphStyle("head", fontName=FONT_NAME, fontSize=8.5, leading=11, alignment=TA_CENTER)
S_CELL = ParagraphStyle("cell", fontName=FONT_NAME, fontSize=8, leading=10.5, alignment=TA_LEFT)
S_NUM = ParagraphStyle("num", fontName=FONT_NAME, fontSize=8, leading=10.5, alignment=TA_RIGHT)
S_NOTE = ParagraphStyle("note", fontName=FONT_NAME, fontSize=8.5, leading=12,
                        textColor=colors.HexColor("#595959"))
GRID_COLOR = colors.HexColor("#9DB2CE")
HEAD_BG = colors.HexColor("#DCE6F1")


class NumberedCanvas(canvas.Canvas):
    """两遍法：先把每页状态存起来，攒到最后知道总页数，再回填「第 X 页 / 共 Y 页」。"""

    def __init__(self, *args, keep_footer=True, **kwargs):
        super().__init__(*args, **kwargs)
        self._keep = keep_footer
        self._states = []

    def showPage(self):
        self._states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        total = len(self._states)
        for state in self._states:
            self.__dict__.update(state)
            if self._keep:
                self.setFont(FONT_NAME, 8)
                self.setFillColor(colors.HexColor("#666666"))
                self.drawRightString(self._pagesize[0] - MARGIN_R, 9 * mm,
                                     f"第 {self._pageNumber} 页 / 共 {total} 页")
            super().showPage()
        super().save()


# ---------------------------------------------------------------------------
# ② 读数据（归堆）：规则表 → 任务清单；工作表 → 表头 + 数据 + 表外说明
# ---------------------------------------------------------------------------
def norm(value):
    """全角空格不算空白，strip() 去不掉，必须显式压掉。"""
    return re.sub(r"\s+", "", str(value if value is not None else ""))


def decimals_of(fmt):
    """从 Excel 数字格式里抠出小数位：0.0000 → 4，General → 0。"""
    if "." not in fmt:
        return 0
    tail = re.split(r"[^0#]", fmt.split(".")[-1])[0]
    return len(tail)


def cell_text(cell):
    """把一个单元格变成"该显示成什么样"的字符串，并说明它是不是数字（决定左对齐还是右对齐）。"""
    value = cell.value
    if value is None or (isinstance(value, str) and value.strip() == ""):
        return "", False
    fmt = (cell.number_format or "").strip()
    if isinstance(value, bool):
        return ("是" if value else "否"), False
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d %H:%M" if (value.hour or value.minute) else "%Y-%m-%d"), False
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d"), False
    if isinstance(value, (int, float)):
        if fmt == "@":                                   # 文本格式的数字：原样输出，别自作聪明
            return str(value), False
        if "%" in fmt:
            return f"{value * 100:.{decimals_of(fmt)}f}%", True
        if "#,##0" in fmt:
            return f"{value:,.{decimals_of(fmt)}f}", True
        if fmt in ("", "General"):
            return (str(int(value)) if float(value).is_integer() else f"{value:g}"), True
        return f"{value:.{decimals_of(fmt)}f}", True
    return str(value).strip(), False


def read_rules(path):
    """读规则表。缺列不报错，用默认值兜；「文件」列留空的行直接跳过。"""
    if not path.exists():
        raise FileNotFoundError(f"找不到导出规则表：{path}")
    tasks = []
    with open(path, encoding="utf-8-sig", newline="") as f:      # 带 BOM 的中文 CSV 要 utf-8-sig
        for raw in csv.DictReader(f):
            row = {(k or "").strip(): (v or "").strip() for k, v in raw.items() if k}
            if not row.get("文件"):
                continue
            try:
                head_rows = int(row.get("表头行数") or 1)
            except ValueError:
                head_rows = 1
            tasks.append({
                "文件": row["文件"],
                "工作表": row.get("工作表") or "",
                "方向": row.get("方向") or "纵向",
                "表头行数": max(1, head_rows),
                "命名规则": row.get("命名规则") or FALLBACK_NAME,
                "页码": (row.get("页码") or "是") == "是",
                "接收方": row.get("接收方") or "",
            })
    return tasks


def read_table(ws, head_rows):
    """把一张工作表读成 表头 / 数据 / 表外说明 三块。

    纵向合并的单元格值向下填满（这样跨页时那一格不会没有值）；
    横向合并只在左上角留值，其余留空——PDF 里跨列合并遇到分页会出问题，宁可少一个单元格。
    数据区 = 表头下面第一列不为空的行；第一列一旦为空，剩下的非空内容归入表外说明。
    """
    last_row, last_col = ws.max_row, ws.max_column
    txt = [[cell_text(ws.cell(r, c))[0] for c in range(1, last_col + 1)] for r in range(1, last_row + 1)]
    isnum = [[cell_text(ws.cell(r, c))[1] for c in range(1, last_col + 1)] for r in range(1, last_row + 1)]
    for rng in ws.merged_cells.ranges:
        for r in range(rng.min_row, min(rng.max_row, last_row) + 1):
            for c in range(rng.min_col, min(rng.max_col, last_col) + 1):
                if r == rng.min_row or c != rng.min_col:
                    continue                                # 横向合并：只留左上角
                txt[r - 1][c - 1] = txt[rng.min_row - 1][rng.min_col - 1]
                isnum[r - 1][c - 1] = isnum[rng.min_row - 1][rng.min_col - 1]
    heads = [row for row in txt[:head_rows]]
    body, note, started = [], "", False
    for i in range(head_rows, last_row):
        if norm(txt[i][KEY_COL - 1]):
            body.append([(txt[i][c], isnum[i][c]) for c in range(len(txt[i]))])
            started = True
        elif started:                                        # 数据区已结束，后面非空的都算表外说明
            tail = [t for t in txt[i] if norm(t)]
            if tail:
                note = (note + " " + " ".join(tail)).strip()
    return heads, body, note


def fit_columns(widths, avail_pt):
    """Excel 列宽（字符数）→ PDF 点宽；总宽超一页就整体等比缩小，只缩不放。"""
    raw = [(w * 7 + 5) * 0.75 for w in widths]              # 字符数 → 像素 → 点，Excel 默认字体口径
    total = sum(raw)
    scale = min(1.0, avail_pt / total) if total else 1.0
    return [round(x * scale, 2) for x in raw], round(scale, 4)


# ---------------------------------------------------------------------------
# ③ 填 / 拼数据：表格样式、页脚、整张 PDF
# ---------------------------------------------------------------------------
def write_pdf(path, heads, body, note, col_pt, page_size, head_rows, keep_footer):
    doc = SimpleDocTemplate(str(path), pagesize=page_size, leftMargin=MARGIN_L, rightMargin=MARGIN_R,
                            topMargin=MARGIN_T, bottomMargin=MARGIN_B,
                            title=path.stem, author="", creator="")
    data = [[Paragraph(escape(t), S_HEAD) for t in row] for row in heads]
    for row in body:
        data.append([Paragraph(escape(t), S_NUM if is_num else S_CELL) for t, is_num in row])
    table = Table(data, colWidths=col_pt, repeatRows=head_rows, hAlign="LEFT")
    table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.4, GRID_COLOR),
        ("BACKGROUND", (0, 0), (-1, head_rows - 1), HEAD_BG),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    story = [table]
    if note:
        story += [Spacer(1, 6), Paragraph(escape(note), S_NOTE)]
    doc.build(story, canvasmaker=partial(NumberedCanvas, keep_footer=keep_footer))


def safe_name(name):
    """文件名里的路径分隔符、通配符一律压成 -：PDF 归档件要能平铺在一个目录里发出去。"""
    cleaned = ILLEGAL.sub("-", name).strip(" -.")
    return cleaned or "未命名"


def unique_name(name, used):
    """重名不覆盖：第二份起加 _2、_3，两份都留着。"""
    if name.lower() not in used:
        used.add(name.lower())
        return name, ""
    stem, suffix = os.path.splitext(name)
    i = 2
    while f"{stem}_{i}{suffix}".lower() in used:
        i += 1
    final = f"{stem}_{i}{suffix}"
    used.add(final.lower())
    return final, f"与已有文件重名，已改名 {final}"


# ---------------------------------------------------------------------------
# ④ 输出（保存、循环）：一条规则一张 PDF，跑完写台账
# ---------------------------------------------------------------------------
def export_one(task, used, today):
    """跑一条规则：四条挂起理由都在这儿分开判，挂起的不产文件、但照样登记。"""
    src = RAW / task["文件"]
    blank = {"源文件": task["文件"], "工作表": task["工作表"], "方向": task["方向"],
             "缩放比": "", "输出文件": "", "页数": "", "大小(KB)": "", "接收方": task["接收方"]}
    if not src.exists():
        return None, {**blank, "状态": "挂起", "备注": "找不到文件，检查规则表里的文件名"}
    book = load_workbook(src, data_only=True)                # 只取值；脚本从不 save 回去
    if task["工作表"] not in book.sheetnames:
        return None, {**blank, "状态": "挂起", "备注": f"找不到工作表（现有：{'、'.join(book.sheetnames)}）"}
    ws = book[task["工作表"]]
    if task["方向"] not in PAGE_SIZES:
        return None, {**blank, "状态": "挂起", "备注": "方向只认「纵向 / 横向」"}
    heads, body, note = read_table(ws, task["表头行数"])
    if not body:
        return None, {**blank, "状态": "挂起", "备注": "没有数据行（只有表头），不出空 PDF"}
    widths = [(ws.column_dimensions[get_column_letter(c)].width or DEFAULT_COL_W)
              for c in range(1, len(heads[0]) + 1)]
    page_size = PAGE_SIZES[task["方向"]]
    col_pt, scale = fit_columns(widths, page_size[0] - MARGIN_L - MARGIN_R)

    name = PLACEHOLDER.sub(lambda m: {"{源文件}": src.stem, "{工作表}": task["工作表"],
                                      "{导出日}": today}[m.group(0)], task["命名规则"])
    name, renamed = unique_name(safe_name(name) + ".pdf", used)
    pdf_path = OUT / name
    write_pdf(pdf_path, heads, body, note, col_pt, page_size,
              task["表头行数"], task["页码"])
    page_count, size = pdf_pages(pdf_path), pdf_path.stat().st_size

    notes = [x for x in [renamed] if x]
    if scale < 1:
        notes.append(f"列宽超一页，已缩放到 {scale:.0%}")
    if note:
        notes.append("末尾附了表外说明 1 处")
    if page_count > 1:
        notes.append(f"跨 {page_count} 页，表头每页重复")
    return name, {**blank, "缩放比": scale, "输出文件": name,
                  "页数": page_count, "大小(KB)": round(size / 1024, 1),
                  "状态": "已导出", "备注": "；".join(notes)}


def pdf_pages(path):
    from pypdf import PdfReader
    return len(PdfReader(str(path)).pages)


def export_all(today, log):
    """按规则表逐条跑，返回 台账行 + 源文件指纹（用来证明源工作簿没被改过）。"""
    before = file_fingerprint()
    tasks = read_rules(RULES)
    used, rows = set(), []
    for i, task in enumerate(tasks, 1):
        _, row = export_one(task, used, today)
        row["序号"] = i
        rows.append(row)
        log(f"[{i:>2}/{len(tasks)}] {row['状态']}  {task['文件']} → {task['工作表']}"
            f"  {row['输出文件'] or row['备注']}")
    return tasks, rows, (before, file_fingerprint())


def write_ledger(rows):
    """导出台账：一份 PDF 一行，挂起的也留一行——复核时要能一眼看出谁没出成、为什么。"""
    book = Workbook()
    ws = book.active
    ws.title = "导出台账"
    head = ["序号", "源文件", "工作表", "方向", "缩放比", "输出文件", "页数", "大小(KB)", "接收方", "状态", "备注"]
    thin = Side(style="thin", color="8EA9DB")
    for c, name in enumerate(head, 1):
        cell = ws.cell(1, c, name)
        cell.font = Font(name="微软雅黑", size=10, bold=True)
        cell.fill = PatternFill("solid", fgColor="DCE6F1")
        cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    done_fill = PatternFill("solid", fgColor="E2EFDA")
    hold_fill = PatternFill("solid", fgColor="FCE4D6")
    for i, row in enumerate(rows):
        r = i + 2
        for c, name in enumerate(head, 1):
            cell = ws.cell(r, c, row.get(name, ""))
            cell.font = Font(name="微软雅黑", size=10)
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
            cell.alignment = Alignment(horizontal="left" if name == "备注" else "center", vertical="center")
            if name == "状态":
                cell.fill = done_fill if row[name] == "已导出" else hold_fill
    for c, w in enumerate([6, 26, 14, 7, 8, 40, 7, 9, 10, 9, 46], 1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = "A2"
    path = OUT / "导出台账.xlsx"
    book.save(path)
    return path


# ---------------------------------------------------------------------------
# 回读校验：不看过程看结果
# ---------------------------------------------------------------------------
def flat(text):
    return re.sub(r"\s+", "", text or "")


def read_pdfs():
    from pypdf import PdfReader
    out = {}
    for p in sorted(OUT.glob("*.pdf")):
        reader = PdfReader(str(p))
        out[p.name] = (len(reader.pages), [flat(page.extract_text()) for page in reader.pages],
                       p.stat().st_size)
    return out


def file_fingerprint():
    """源工作簿的 sha256：跑完对一次，证明导出过程没碰过原始记录。"""
    out = {}
    for p in sorted(RAW.glob("*.xlsx")):
        out[p.name] = hashlib.sha256(p.read_bytes()).hexdigest()
    return out


def verify(tasks, rows, fingerprints, today):
    results = []

    def check(name, ok):
        results.append(bool(ok))
        print(("  OK   " if ok else "  FAIL ") + name)

    done = [r for r in rows if r["状态"] == "已导出"]
    hold = [r for r in rows if r["状态"] == "挂起"]
    pdfs = read_pdfs()
    def one(sheet):
        return next((r for r in done if r["工作表"] == sheet), None)
    src1, src2 = "检验记录台账_2026Q3", "稳定性考察记录_2026Q3"
    rel, rg1, rg2, stb, rmk = (one("放行台账"), one("试剂消耗台账"), None, one("第 3 批"), one("备注说明"))
    rg2 = next((r for r in done if r["输出文件"].endswith("_2.pdf")), None)
    exp_rel = f"{src1}_放行台账_{today}.pdf"
    exp_rg1 = f"{src1}_试剂消耗台账.pdf"
    exp_stb = f"{src2}_第 3 批_{today}.pdf"
    exp_rmk = f"{src2}-备注说明.pdf"

    check("规则表读出 8 条任务", len(tasks) == 8)
    check("已导出 5 份", len(done) == 5)
    check("挂起 3 条", len(hold) == 3)
    check("挂起：工作表不存在（报废台账）",
          any(norm(r["备注"]).find("找不到工作表") >= 0 for r in hold))
    check("挂起：只有表头没有数据行（培养基空白页）",
          any(norm(r["备注"]).find("没有数据行") >= 0 for r in hold))
    check("挂起：文件缺失（试剂出入库_2026Q3.xlsx）",
          any(norm(r["备注"]).find("找不到文件") >= 0 for r in hold))
    check("输出目录里正好 5 份 PDF", len(pdfs) == 5)

    check(f"放行台账文件名带导出日 {exp_rel}", rel and rel["输出文件"] == exp_rel)
    pages, texts, size = pdfs.get(exp_rel, (0, [], 0))
    check("放行台账 跨 2-3 页（60 行装不下一页）", 2 <= pages <= 3)
    check("放行台账 第 1 页有表头「批号」", "批号" in texts[0])
    check("放行台账 末页仍有表头「批号」（每页重复）", "批号" in texts[-1])
    check("放行台账 第 1 页页脚「第 1 页」", "第1页" in texts[0])
    check(f"放行台账 末页页脚「第 {pages} 页」", f"第{pages}页" in texts[-1] and "共" in texts[-1])
    check("放行台账 首行批号前导零保住 0000710", "0000710" in texts[0])
    check("放行台账 末行批号 0000769 也在", "0000769" in texts[-1])
    check("放行台账 文字结论「未检出」原样进 PDF", "未检出" in texts[0])
    check("放行台账 判定「不合格」原样进 PDF", "不合格" in texts[0])
    check("放行台账 表外说明没丢", any("低于检出限" in t for t in texts))
    check("放行台账 窄表不缩放（缩放比 = 1.0）", rel and rel["缩放比"] == 1.0)
    check("放行台账 台账备注写了表外说明", rel and "表外说明" in rel["备注"])

    check("试剂消耗台账 文件名无后缀", rg1 and rg1["输出文件"] == exp_rg1)
    check("试剂消耗台账 列宽超一页已缩放（0.75-0.90）",
          rg1 and 0.75 <= float(rg1["缩放比"]) <= 0.90)
    check("试剂消耗台账 只有 1 页", pdfs.get(exp_rg1, (0,))[0] == 1)
    check("同表第二份（纵向）重名自动加 _2", rg2 and rg2["输出文件"] == f"{src1}_试剂消耗台账_2.pdf")
    check("重名那份台账里写明了改名", rg2 and "重名" in rg2["备注"])

    check("备注说明 命名规则里的 / 已压成 -", rmk and rmk["输出文件"] == exp_rmk)
    rmk_pages, rmk_texts, _ = pdfs.get(exp_rmk, (0, [], 0))
    check("备注说明 只有 1 页", rmk_pages == 1)
    check("备注说明 规则写了「页码=否」→ 真的没有页脚", rmk_texts and "第1页" not in rmk_texts[0])
    check("备注说明 长文本里的 & 没被吃掉", rmk_texts and "&" in rmk_texts[0])

    stb_pages, stb_texts, _ = pdfs.get(exp_stb, (0, [], 0))
    check("第 3 批 只有 1 页", stb_pages == 1)
    check("第 3 批 纵向合并的考察点已向下填满（0 月 出现 2 次）",
          stb_texts and stb_texts[0].count("0月") >= 2)

    check("源工作簿跑完没被改动（sha256 前后一致）", fingerprints[0] == fingerprints[1])
    check("每份 PDF 都不小于 2 KB", all(v[2] > 2048 for v in pdfs.values()))
    check("导出台账 11 列都在", (OUT / "导出台账.xlsx").exists())

    verdict = "PASS" if all(results) else "FAIL"
    print(f"\n  回读校验：{sum(results)}/{len(results)} 项通过 → {verdict}")
    return verdict, sum(results), len(results)


# ---------------------------------------------------------------------------
# 入口：工程外壳（日志 / 暂停）都在这里，业务逻辑全在上面
# ---------------------------------------------------------------------------
def host_line():
    name = socket.gethostname()
    ip = "未知"
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
    except Exception:
        try:
            ip = socket.gethostbyname(name)
        except Exception:
            ip = "未知"
    return f"运行主机：{name}（{ip}）"


def main():
    started = time.time()
    LOG.parent.mkdir(parents=True, exist_ok=True)
    log = lambda msg: open(LOG, "a", encoding="utf-8").write(msg + "\n")
    log(host_line())
    log(f"开始执行：{datetime.now():%Y-%m-%d %H:%M:%S}  Python {sys.version.split()[0]}（{sys.executable}）")

    shutil.rmtree(OUT, ignore_errors=True)          # 重跑先清空，免得上一版文件被当成这一版
    OUT.mkdir(parents=True, exist_ok=True)
    today = date.today().strftime("%Y%m%d")
    tasks, rows, fingerprints = export_all(today, log)
    ledger = write_ledger(rows)
    print(f"  导出 {sum(1 for r in rows if r['状态'] == '已导出')} 份，"
          f"挂起 {sum(1 for r in rows if r['状态'] == '挂起')} 条，台账：{ledger.name}")
    verdict, passed, total = verify(tasks, rows, fingerprints, today)

    used = time.time() - started
    log(f"执行完成：{datetime.now():%Y-%m-%d %H:%M:%S}  用时 {used:.1f}s  判定 {verdict}")
    print(f"  用时 {used:.1f}s  判定 {verdict}")
    return verdict


if __name__ == "__main__":
    result = main()
    if "--no-pause" not in sys.argv:
        try:
            input("按回车退出 ...")
        except EOFError:
            pass
    sys.exit(0 if result == "PASS" else 1)
