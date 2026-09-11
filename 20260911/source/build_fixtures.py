# -*- coding: utf-8 -*-
"""造示例数据：模拟三家仪器各自导出的检验结果 Excel 台账。

运行一次即可，会在示例根目录的 01_raw_data/ 下生成三个 xlsx 文件。
真实使用时不需要这一步——这些文件由仪器工作站软件导出。
三个文件故意造的不一致：表头位置不同（第 1 行 / 第 2 行 / 第 3 行）、
批号带前导零、末尾带手工合计行、检查项少到只有 1 行。
"""
import subprocess
import sys
from pathlib import Path

MIRROR = "https://pypi.tuna.tsinghua.edu.cn/simple"

try:
    import openpyxl
    from openpyxl.styles import Font
except ImportError:
    print(f"缺 openpyxl，正在用清华源安装（解释器: {sys.executable}）...")
    subprocess.run([sys.executable, "-m", "pip", "install", "-i", MIRROR, "openpyxl"],
                   check=False)
    import openpyxl
    from openpyxl.styles import Font

BASE = Path(__file__).resolve().parent.parent   # 示例根目录（本脚本在 source/ 里）
RAW = BASE / "01_raw_data"
RAW.mkdir(parents=True, exist_ok=True)

HEADERS = ["样品编号", "样品名称", "批号", "检验项目", "标准规定", "检验结果", "单项结论"]
TITLE_FONT = Font(bold=True)
HEADER_FONT = Font(bold=True)


def _write(path: Path, rows: list, header_row: int, title: str = "") -> None:
    """header_row 从 1 数起；title 写在表头上方时 header_row=2。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    if title:
        ws.cell(row=1, column=1, value=title).font = TITLE_FONT
    for j, name in enumerate(HEADERS, start=1):
        c = ws.cell(row=header_row, column=j, value=name)
        c.font = HEADER_FONT
    for i, row in enumerate(rows, start=header_row + 1):
        for j, val in enumerate(row, start=1):
            ws.cell(row=i, column=j, value=val)
    # 批号一列按文本处理，防止前导零被 Excel 吃掉
    for i in range(header_row + 1, header_row + 1 + len(rows)):
        ws.cell(row=i, column=3).number_format = "@"
    wb.save(path)


# ---------- 仪器 A：HPLC 工作站，表头在第 1 行，末尾带手工合计行 ----------
hplc = [
    ["YB-2026-0901-001", "阿莫西林胶囊", "A0901", "含量", "90.0%~110.0%", 0.982, "符合"],
    ["YB-2026-0901-001", "阿莫西林胶囊", "A0901", "有关物质", "单个杂质 ≤0.5%", 0.21, "符合"],
    ["YB-2026-0901-001", "阿莫西林胶囊", "A0901", "水分", "不得过 12.0%", 0.087, "符合"],
    ["YB-2026-0901-002", "阿莫西林胶囊", "A0902", "含量", "90.0%~110.0%", 0.975, "符合"],
    ["YB-2026-0901-002", "阿莫西林胶囊", "A0902", "水分", "不得过 12.0%", 0.092, "符合"],
]
# ---------- 仪器 B：GC 工作站，标题占了第 1 行、表头在第 2 行，批号带前导零 ----------
gc = [
    ["YB-2026-0901-003", "乙醇（药用辅料）", "00903", "残留溶剂", "不得过 5000 ppm", 312, "符合"],
    ["YB-2026-0901-003", "乙醇（药用辅料）", "00903", "密度", "0.981~1.005", 0.992, "符合"],
]
# ---------- 仪器 C：酶标仪，标题+空行，表头在第 3 行，只有 1 行数据 ----------
elisa = [
    ["YB-2026-0901-004", "β-内酰胺酶", "00904", "效价", "≥90%", 0.94, "符合"],
]

_write(RAW / "检验结果_HPLC工作站.xlsx", hplc, header_row=1)
_write(RAW / "检验结果_GC工作站.xlsx", gc, header_row=2, title="残留溶剂检验数据导出")
_write(RAW / "检验结果_酶标仪.xlsx", elisa, header_row=3, title="酶标仪数据导出（勿删）")

print("已生成 01_raw_data/ 下三个 xlsx（表头位置 1/2/3 行各不相同，批号带前导零）")
