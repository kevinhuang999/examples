#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""第三方检测 · 检测报告用印：把当天送进用印室的报告逐份盖好章。

一句话：盖章这件事难的不是"把图贴上去"，是**每份报告要盖哪几枚章、盖在哪一页、
盖之前这份报告够不够格**——这些规则在每家机构都不一样，而系统方案默认你已经在系统里出报告了。

用法：
    python run.py              # 跑完停住等回车（Windows 双击也行）
    python run.py --no-pause   # 跑完直接退出
"""

import csv
import io
import re
import socket
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

# ----------------------------------------------------------------- ① 配置区（跑之前主要改这里）

ROOT = Path(__file__).resolve().parent
RAW_DIR = ROOT / "01_raw_data"
OUT_DIR = ROOT / "02_output"
TPL_DIR = ROOT / "source" / "templates"
LOG_FILE = ROOT / "source" / "run_log.txt"

REQ_CSV = RAW_DIR / "盖章申请单.csv"
RULE_CSV = RAW_DIR / "用章规则表.csv"
STAMP_DIR = RAW_DIR / "印章图"
PDF_DIR = RAW_DIR / "报告原件"
TPL_XLSX = TPL_DIR / "盖章台账模板.xlsx"

LEDGER_XLSX = OUT_DIR / "盖章台账.xlsx"
STAMPED_DIR = OUT_DIR / "已盖章报告"

# 申请单判空列用「申请日期」（每行都会有），规则表用「报告类型」。
# 换一张表先问一句"靠哪一列判空"——判空列挑错，整表会被静默读空。
REQ_HEAD = "申请日期"
RULE_HEAD = "报告类型"

# 只有「已审签」的报告能进用印室；其它状态一律挂起，不猜、不替人放行。
READY_STATUS = "已审签"

# 章图在页面上的尺寸（正方形，单位 pt）与距页边的留白。
# 章图跟着页面尺寸走：A4 纵向和横向一样按「右下角往里缩 MARGIN」定位。
STAMP_W = 110
MARGIN = 56
BOTTOM_OFFSET = 16          # 右下角那两枚章再往上抬一点，免得压住页脚那行小字

# 盖章的人：用印是专职岗，谁把报告送过来都要过这一个人手，所以写成常量，不从数据里取。
OPERATOR = "用印室"

LEADING_ZERO = re.compile(r"^\d")
BLANK_RE = re.compile(r"[\s\u3000]+")   # 全角空格 strip() 不管它，要显式replace


# ----------------------------------------------------------------- 装依赖 / 日志


def ensure_deps():
    """缺 openpyxl / pypdf 就自己装（走清华源）。"""
    need = []
    for mod, pkg in (("openpyxl", "openpyxl"), ("pypdf", "pypdf")):
        try:
            __import__(mod)
        except ImportError:
            need.append(pkg)
    if not need:
        return
    pkgs = ["-i", "https://pypi.tuna.tsinghua.edu.cn/simple", *need]
    for extra in ([], ["--user"]):
        r = subprocess.run([sys.executable, "-m", "pip", "install", *pkgs, *extra],
                           capture_output=True, text=True)
        if r.returncode == 0:
            return
    print("依赖安装失败，请手动执行：\n"
          "pip install openpyxl pypdf -i https://pypi.tuna.tsinghua.edu.cn/simple")
    sys.exit(1)


def host_line():
    """运行主机：主机名 + 本机对外网卡地址，用来分辨这段日志是哪台机器跑出来的。"""
    try:
        name = socket.gethostname()
    except Exception:
        name = "未知"
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


def log(text):
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(text + "\n")


# ----------------------------------------------------------------- ② 读数据


def norm(value):
    """比较用的键：去空白（含全角空格）。显示照原样，只压比较用的那一份。"""
    return BLANK_RE.sub("", str(value if value is not None else ""))


def read_csv_rows(path, head_col):
    """读一张 CSV 成 dict 列表，全部按文本取。

    **一律不转类型**——报告编号 0000122 一转 int 就变 122，台账和报告封面上写的都是带前导零的那个。
    表尾的说明行靠 `head_col` 那一列为空跳过；所以说明文字别写进那一列。
    """
    rows = []
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        for raw in csv.DictReader(f):
            if not raw or not norm(raw.get(head_col, "")):
                continue
            rows.append({k: ("" if v is None else str(v)) for k, v in raw.items()})
    return rows


def read_requests():
    """申请单：每行一份报告，行序就是用印室收到的顺序，台账跟着这个顺序走。"""
    return read_csv_rows(REQ_CSV, REQ_HEAD)


def read_rules():
    """用章规则表：报告类型 → 要盖的章（含落位与骑缝起始页），一张类型可以配多行。"""
    rules = {}
    for row in read_csv_rows(RULE_CSV, RULE_HEAD):
        kind = norm(row["报告类型"])
        rules.setdefault(kind, []).append({
            "章名": row["章名"].strip(),
            "落位": row["落位"].strip(),
            "骑缝起始页": row["骑缝起始页"].strip(),
        })
    return rules


def read_pdf_info(path):
    """读原件的页数与带旋转的页码。

    带旋转的页（/Rotate 非 0）落位坐标要跟显示方向对齐，脚本不猜——整份挂起，交给人处理。
    """
    from pypdf import PdfReader

    reader = PdfReader(str(path))
    pages = len(reader.pages)
    rotated = []
    for i, page in enumerate(reader.pages, 1):
        if int(page.get("/Rotate", 0) or 0) % 360 != 0:
            rotated.append(i)
    return pages, rotated


# ----------------------------------------------------------------- ③ 判定与排章


def stamp_xy(place, width, height, seq, total):
    """算章图左下角坐标（PDF 原点在左下角）。

    · 首页右下 / 末页右下：按页边留白定位，页面是横是竖都不影响。
    · 右侧骑缝：x 取「页宽 − 半个章宽」，**章图一半被放到页面外**——PDF 只画页面内的部分，
      天然就是半个章，不用去裁图；y 按页码阶梯下移，装订起来侧面才是完整的一枚。
    """
    if place in ("首页右下", "末页右下"):
        return width - MARGIN - STAMP_W, MARGIN + BOTTOM_OFFSET
    x = width - STAMP_W / 2
    top = height - MARGIN - STAMP_W
    step = (top - MARGIN) / max(total - 1, 1)
    return x, top - seq * step


def make_plan(req, rules, used, row_no):
    """把一行申请排成一份用印计划；排不出来就返回挂起原因。

    判定顺序 = 业务顺序：先看这份报告**够不够格**（审签状态），再看**有没有用章规则**，
    再看**章图和原件齐不齐**，再看**页数够不够骑缝**，最后才看**是不是有人已经报过它**。
    顺序反了，挂起理由会指向错误的补料动作——拿着"原件缺失"去补资料，补完发现这单根本没审签。
    """
    no = req["报告编号"].strip()
    kind = norm(req["报告类型"])
    if not no:
        return None, "报告编号空，台账上没法标识"

    status = norm(req["审签状态"])
    if status != READY_STATUS:
        return None, f"审签状态是「{req['审签状态'].strip()}」，用印室只收 {READY_STATUS} 的报告"

    if kind not in rules:
        return None, f"用章规则表里没有「{req['报告类型'].strip()}」这个类型，先定规则"

    missing = [r["章名"] for r in rules[kind] if not (STAMP_DIR / f"{r['章名']}.png").exists()]
    if missing:
        return None, "章图没有：" + "、".join(missing) + "（去印章图目录补上）"

    pdf = PDF_DIR / req["报告文件"].strip()
    if not pdf.exists():
        return None, f"报告原件没到：{req['报告文件'].strip()}"

    pages, rotated = read_pdf_info(pdf)
    if rotated:
        return None, f"第 {rotated[0]} 页带旋转，落位要人工核，先别盖"

    if no in used:
        return None, "重复申请：这份报告本次已经用印，第二次不再盖"

    plan = {
        "行号": row_no,
        "报告编号": no,
        "客户名称": req["客户名称"].strip(),
        "报告类型": req["报告类型"].strip(),
        "原件": pdf,
        "页数": pages,
        "章": [],
    }
    for rule in rules[kind]:
        place = rule["落位"]
        if place == "首页右下":
            codes = "1"
        elif place == "末页右下":
            codes = str(pages)
        else:
            start = int(rule["骑缝起始页"] or 1)
            if start > pages:
                return None, f"骑缝章要从第 {start} 页起，这份只有 {pages} 页，页数对不上"
            codes = f"{start}-{pages}"
        plan["章"].append({"章名": rule["章名"], "落位": place, "页码": codes})
    return plan, None


def build_page_map(plan):
    """章 → 页：把「一张章」展开成「要落在哪几页」，顺手算好它在骑缝序列里的第几格。

    台账里记的是页码区间（如 2-6），落图要逐页落；骑缝章的垂直位置按**骑缝页序**排，
    所以 seq 从骑缝起始页起数 0、total 是骑缝页数——用整份报告的页数去算，最后几页会挤在一起。
    """
    page_map = {}
    for stamp in plan["章"]:
        start, _, end = stamp["页码"].partition("-")
        start, end = int(start), int(end or start)
        for p in range(start, end + 1):
            page_map.setdefault(p, []).append({**stamp, "seq": p - start,
                                               "total": end - start + 1})
    return page_map


# ----------------------------------------------------------------- ④ 盖章与输出


def build_overlay(width, height, stamps_on_page):
    """给一页生成一张只有章的透明覆盖页。

    为什么要单独生成一张"覆盖页"再合过去：PDF 里合页只认 PageObject，
    PNG 图片得先由 reportlab 画成一页，才能叠到原件上——`drawImage(mask="auto")` 保住 alpha，
    否则章图会带一块白底，压住报告上的字。
    """
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfgen import canvas
    from pypdf import PdfReader

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(width, height))
    for stamp in stamps_on_page:
        x, y = stamp_xy(stamp["落位"], width, height, stamp["seq"], stamp["total"])
        c.drawImage(ImageReader(str(STAMP_DIR / f"{stamp['章名']}.png")), x, y,
                    width=STAMP_W, height=STAMP_W, mask="auto")
    c.save()
    buf.seek(0)
    return PdfReader(buf).pages[0]


def stamp_pdf(plan):
    """逐页合章，输出盖章件。原件一张不动——它是归档凭证，盖过的另存一份。"""
    from pypdf import PdfReader, PdfWriter

    reader = PdfReader(str(plan["原件"]))
    writer = PdfWriter()
    page_map = build_page_map(plan)
    for i, page in enumerate(reader.pages, 1):
        stamps = page_map.get(i)
        if stamps:
            w = float(page.mediabox.width)
            h = float(page.mediabox.height)
            page.merge_page(build_overlay(w, h, stamps))
        writer.add_page(page)
    dest = STAMPED_DIR / f"已盖章_{plan['报告编号']}.pdf"
    with open(dest, "wb") as f:
        writer.write(f)
    return dest


def write_ledger(rows, holds):
    """台账按「报告 × 章」登记：审核来问"这份报告盖了哪几枚章"，一行行对得上。"""
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment, Border, Font, Side

    wb = load_workbook(TPL_XLSX)
    ws = wb["盖章台账"]
    thin = Side(style="thin", color="D0D0D0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in rows:
        ws.append(row)
        for c in ws[ws.max_row]:
            c.border = border
            c.font = Font(name="微软雅黑", size=10)
            if c.column == 1:
                c.number_format = "@"
                c.alignment = Alignment(horizontal="left")
    hs = wb["挂起清单"]
    for row in holds:
        hs.append(row)
        for c in hs[hs.max_row]:
            c.font = Font(name="微软雅黑", size=10)
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    wb.save(LEDGER_XLSX)


# ----------------------------------------------------------------- 回读校验


def verify(rows, holds, plans, expect):
    from openpyxl import load_workbook
    from pypdf import PdfReader

    checks = []

    def ok(name, cond):
        checks.append((name, bool(cond)))

    wb = load_workbook(LEDGER_XLSX)
    ws = wb["盖章台账"]
    read = []
    for r in range(2, ws.max_row + 1):
        no = ws.cell(row=r, column=1).value
        if not no:
            continue
        read.append({
            "报告编号": str(no),
            "客户名称": str(ws.cell(row=r, column=2).value),
            "章名": str(ws.cell(row=r, column=4).value),
            "页码": str(ws.cell(row=r, column=6).value),
            "结论": str(ws.cell(row=r, column=9).value),
        })
    hs = wb["挂起清单"]
    hold_read = []
    for r in range(2, hs.max_row + 1):
        no = hs.cell(row=r, column=1).value
        if no:
            hold_read.append((str(no), str(hs.cell(row=r, column=4).value)))

    ok(f"台账 {expect['rows']} 行（报告 × 章 展开）", len(read) == expect["rows"])
    ok("台账里每行结论都是「已盖章」", all(x["结论"] == "已盖章" for x in read))
    ok(f"盖章件 {expect['stamped']} 份", len(list(STAMPED_DIR.glob("*.pdf"))) == expect["stamped"])
    ok(f"挂起 {expect['holds']} 行", len(hold_read) == expect["holds"])

    # 前导零：老系统编号 0000122 一路带到底，没被转成 122
    ok("报告编号 0000122 的前导零保住了", any(x["报告编号"] == "0000122" for x in read))
    ok("盖章件文件名也带着前导零", (STAMPED_DIR / "已盖章_0000122.pdf").exists())

    # 骑缝页码 = 起始页到末页；0000122 是 4 页、规则从第 2 页起
    ok("0000122 的骑缝页码是 2-4", any(x["报告编号"] == "0000122" and x["章名"] == "骑缝章"
                                        and x["页码"] == "2-4" for x in read))
    ok("首页章页码写 1", all(x["页码"] == "1" for x in read if x["章名"] == "检测专用章"))
    ok("末页章页码 = 该报告页数（JC2026-1001 是 6）",
       any(x["报告编号"] == "JC2026-1001" and x["章名"] == "资质认定章" and x["页码"] == "6" for x in read))

    # 章的组合按规则表走，不是每份都三枚
    ok("内部质控报告只盖 1 枚章", sum(1 for x in read if x["报告编号"] == "JC2026-1004") == 1)
    ok("加急委托报告盖 2 枚章", sum(1 for x in read if x["报告编号"] == "JC2026-1005") == 2)
    ok("委托检测报告盖 3 枚章", sum(1 for x in read if x["报告编号"] == "JC2026-1002") == 3)

    # 不够格的一份都不进台账
    nos = {x["报告编号"] for x in read}
    ok("审核中的不在台账里", "JC2026-1007" not in nos)
    ok("已作废的不在台账里", "JC2026-1008" not in nos)
    ok("原件没到的不在台账里", "JC2026-1009" not in nos)
    ok("规则表里没有的类型不在台账里", "JC2026-1011" not in nos)
    ok("页数不够骑缝的不在台账里", "JC2026-1013" not in nos)

    # 重复申请：这份报告只盖了一次，第二次只挂起
    ok("重复申请的报告只盖一次（JC2026-1001 仍是 3 行）",
       sum(1 for x in read if x["报告编号"] == "JC2026-1001") == 3)
    ok("重复申请那条挂起了", any(no == "JC2026-1001" and "重复申请" in r for no, r in hold_read))
    ok("章图缺失那条挂起、且指名是哪枚章",
       any("涉外认证章" in r for _, r in hold_read))
    ok("页数不够那条挂起、理由里带页数",
       any(no == "JC2026-1013" and "只有 1 页" in r for no, r in hold_read))
    ok("全角空格的客户名归一化后照样进台账", any("金鼎" in x["客户名称"] for x in read))

    # 骑缝章的 seq 从 0 起：第一页骑缝贴在最上面一格
    xy_first = stamp_xy("右侧骑缝", 595.28, 841.89, 0, 5)
    xy_last = stamp_xy("右侧骑缝", 595.28, 841.89, 4, 5)
    ok("骑缝章第一格在最上面、最后一格在最下面", xy_first[1] > xy_last[1])
    ok("骑缝章的 x 在页面右边缘（半个章在页外）", xy_first[0] > 595.28 - STAMP_W)
    ok("右下角两枚章的坐标一样（同一位置、不同章）",
       stamp_xy("首页右下", 595.28, 841.89, 0, 1) == stamp_xy("末页右下", 595.28, 841.89, 0, 1))

    # 盖章件：页数与原件一致；章图带透明遮罩（白底会压住报告上的字）
    src_pages = len(PdfReader(str(plans[0]["原件"])).pages)
    out = PdfReader(str(STAMPED_DIR / f"已盖章_{plans[0]['报告编号']}.pdf"))
    ok("盖章件页数 = 原件页数（不是把章另起一页）", len(out.pages) == src_pages)
    def masked_on(page):
        """数一页里带透明掩码的图。透明底丢了的话，章的空白部分会变成实色块压住报告的字。"""
        xobj = page["/Resources"]["/XObject"]
        n = 0
        for k in xobj.keys():
            obj = xobj[k]
            obj = obj.get_object() if hasattr(obj, "get_object") else obj
            if "/SMask" in obj:
                n += 1
        return n

    ok("首页章图带透明掩码（alpha 没丢）", masked_on(out.pages[0]) >= 1)
    ok("末页章图也带透明掩码", masked_on(out.pages[-1]) >= 1)
    ok("原件一个字节没动（还是原件目录里的原样）",
       all(p["原件"].parent == PDF_DIR and p["原件"].exists() for p in plans))

    # 台账顺序跟着申请单行序，不是按编号字符串排（字符串排会把 JC2026-1002 排到 0000122 前）
    order = [x["报告编号"] for x in read]
    ok("台账按申请单顺序排，不是按编号字符串排", order != sorted(order))
    ok("第一个进台账的是申请单第一行 JC2026-1001", order[0] == "JC2026-1001")

    print()
    for name, good in checks:
        print(f"  {'OK  ' if good else 'FAIL'} {name}")
    passed = sum(1 for _, g in checks if g)
    print(f"\n回读校验：{passed}/{len(checks)} 项通过")
    return passed == len(checks), len(checks), passed


# ----------------------------------------------------------------- 主线


def main():
    ensure_deps()
    t0 = time.time()
    LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
    log(host_line())
    log(f"开始执行：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    log(f"Python：{sys.version.split()[0]}（{sys.executable}）")

    if STAMPED_DIR.exists():
        import shutil
        shutil.rmtree(STAMPED_DIR)      # 重跑先清空，否则上一版的件会被当成这一版发出去
    STAMPED_DIR.mkdir(parents=True, exist_ok=True)

    requests = read_requests()
    rules = read_rules()
    plans, holds = [], []
    used = set()
    at = datetime.now().strftime("%Y-%m-%d %H:%M")
    for row_no, req in enumerate(requests, 1):
        plan, reason = make_plan(req, rules, used, row_no)
        if plan is None:
            holds.append([req["报告编号"].strip(), req["客户名称"].strip(),
                          req["报告类型"].strip(), reason])
            continue
        plans.append(plan)
        used.add(plan["报告编号"])

    rows = []
    for plan in plans:
        stamped = stamp_pdf(plan)
        for stamp in plan["章"]:
            rows.append([plan["报告编号"], plan["客户名称"], plan["报告类型"], stamp["章名"],
                         stamp["落位"], stamp["页码"], at, OPERATOR, "已盖章"])
        log(f"已盖章：{plan['报告编号']}（{plan['页数']} 页 / {len(plan['章'])} 枚）→ {stamped.name}")

    write_ledger(rows, holds)
    expect = {"rows": 18, "stamped": 7, "holds": 7}
    ok, total, passed = verify(rows, holds, plans, expect)

    used_s = time.time() - t0
    print(f"\n申请单 {len(requests)} 行 → 盖章 {len(plans)} 份 / 台账 {len(rows)} 行 → 挂起 {len(holds)} 行")
    print(f"结果目录：{OUT_DIR}")
    log(f"申请单 {len(requests)} 行 → 盖章 {len(plans)} 份 / 台账 {len(rows)} 行 → 挂起 {len(holds)} 行")
    log(f"执行完成：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}  用时 {used_s:.1f}s  "
        f"判定 {'PASS' if ok else 'FAIL'}（{passed}/{total}）")
    print(f"判定：{'PASS' if ok else 'FAIL'}（{passed}/{total} 项）")

    if "--no-pause" not in sys.argv:
        input("\n按回车退出…")


if __name__ == "__main__":
    main()
