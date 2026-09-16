# -*- coding: utf-8 -*-
"""引流长文《一份方案改一次，几百套采血 Kit 的配置单就全得重抄》配套脚本。

把三份导出文件拼成一份 Kit 配置单：
    方案规格（一行 = 某模板某访视点一种管的「每例」用量）
  + 中心入组计划（一行 = 一个中心用哪个模板、计划几例）
  + 物料批次台账（一行 = 一个批次，一个管型挂好几个批次）
→ 算每个中心每个访视点要组几套 Kit、每样管按 FEFO 从库存领多少支、每支管子贴什么标签；
  料不够的、名字对不上的、中心没启动的，一并列出来交给人去核，脚本不替人拍板。

核心代码与正文逐字一致，末尾多了自动回读校验，方便确认该标出来的都标出来了。
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
from datetime import date, datetime, timedelta
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


_ensure(("openpyxl", "openpyxl"))

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

HERE = Path(__file__).resolve().parent
SPEC_FILE = HERE / "01_raw_data" / "研究规格_方案导出.csv"           # 一行一种管，给的是每例用量
PLAN_FILE = HERE / "01_raw_data" / "中心入组计划_项目组导出.csv"      # 一个中心一行
STOCK_FILE = HERE / "01_raw_data" / "物料批次_仓库导出.csv"          # 一个管型挂好几个批次
OUT_FILE = HERE / "02_output" / "Kit配置清单.xlsx"                  # 组装清单 + 领料明细 + 标签 + 异常
LOG_DIR = HERE / "source"                                         # 运行日志（run_log.txt 就落在 source/ 下）

KEY_COL = "中心编号"
SPEC_COLS = ["访视模板", "访视编号", "访视点", "管型", "每例管数", "采血量mL", "处理要求"]
PLAN_COLS = ["中心编号", "中心名称", "国家", "访视模板", "计划例数", "启动日期"]
STOCK_COLS = ["管型", "批次号", "效期", "库存支数"]

# 组装日。物料效期要盖过「组装日 + 中心放着备用的 3 个月 + 采样窗口 1 个月」，
# 只跟今天比是不行的：料发到中心还得放几个月才用得上。
BUILD_DATE = date(2026, 9, 16)
SHELF_DAYS = 120

# 方案上写的是检测用途，仓库里按自己录的规格名记料，两套叫法必须显式对上。
# 不做模糊匹配：凝血管有 3.2% 和 3.8%、冻存管有内旋和外旋，猜错一次整批数据作废。
TUBE_ALIAS = {
    "EDTA抗凝管": "紫帽 EDTA-K2 6mL",
    "促凝管": "黄帽 促凝 5mL",
    "冻存管": "冻存管 2mL 外旋",
    "尿杯": "尿杯 100mL 带盖",
    "柠檬酸钠凝血管": "蓝帽 柠檬酸钠 3.2%",
}

UNIT_COLS = ["中心编号", "中心名称", "访视点", "套数", "计划每套管数", "实际每套管数",
             "内容物支数", "说明"]
ITEM_COLS = ["中心编号", "中心名称", "访视点", "管型", "仓库规格", "每例管数", "套数",
             "需用支数", "批次号", "效期", "领用支数"]
LABEL_COLS = ["标签编号", "中心编号", "中心名称", "访视点", "套号", "管位", "管型",
              "采血量mL", "处理要求"]
ALERT_COLS = ["中心编号", "中心名称", "访视点", "类型", "说明"]

HEAD_FILL = PatternFill("solid", fgColor="D9E1F2")
BAD_FILL = PatternFill("solid", fgColor="FFC7CE")     # 要人拍板的（缺料、模板对不上）

# 异常的类型词。写得具体一点，接的人一眼知道该找谁。
KIND_DUP = "方案重复行"
KIND_SHELF = "批次被跳过"
KIND_NO_TUBE = "管型没对上"
KIND_NO_STOCK = "物料表里没这个规格"
KIND_SHORT = "库存不足"
KIND_NOT_STARTED = "中心未启动"
KIND_NO_TEMPLATE = "访视模板对不上"


def to_int(text: str) -> int:
    """例数、支数写成空、"待定"、带千分位的，一律当 0 处理，别让整批算不出来。"""
    try:
        return int(str(text).strip().replace(",", ""))
    except (TypeError, ValueError):
        return 0


def to_date(text: str):
    """把效期文本转成日期；转不动的返回 None——宁可当没效期，也不能当没过期。"""
    try:
        return datetime.strptime(str(text).strip(), "%Y-%m-%d").date()
    except ValueError:
        return None


def read_spec(path: Path):
    """读方案规格，一行一个访视点的一种管，给的是每例用量。

    方案修订留下的重复行不能默默取后一条：先读到的那条才算，同时把冲突记出来，
    到底该听哪条得回去看方案，脚本不替人拍板。
    表尾的合计行、说明行落在没有访视点的那几行，直接丢。
    """
    rows, seen, conflicts = [], {}, []
    with open(path, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            if not (row.get("访视点") or "").strip():
                continue                            # 空行、合计行、说明行都落在这儿
            item = {c: (row.get(c) or "").strip() for c in SPEC_COLS}
            item["每例管数"] = to_int(item["每例管数"])
            key = (item["访视模板"], item["访视编号"], item["访视点"], item["管型"])
            if key in seen:
                conflicts.append({"模板": item["访视模板"], "访视点": item["访视点"],
                                  "管型": item["管型"], "用到的": seen[key]["每例管数"],
                                  "被忽略的": item["每例管数"]})
                continue
            seen[key] = item
            rows.append(item)
    return rows, conflicts


def read_plan(path: Path) -> list:
    """读中心入组计划，一个中心一行。中心编号带前导零，全程当文本，不能被当成数字。"""
    out = []
    with open(path, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            code = (row.get(KEY_COL) or "").strip()
            if not code:
                continue                            # 表尾的合计行、说明行都落在这儿
            out.append({c: (row.get(c) or "").strip() for c in PLAN_COLS})
    return out


def read_stock(path: Path, build_date: date):
    """读物料批次台账，按管型归堆，并按效期从近到远排——发料讲 FEFO，先出效期近的。

    效期盖不到「组装日 + 货架期」的批次直接不进池子：只看今天没过期就发出去，
    料到了中心可能已经用不了了。被筛掉的批次也留个记录，让人知道仓库里还有这批料。
    """
    cutoff = build_date + timedelta(days=SHELF_DAYS)
    pool, dropped, known = {}, [], set()
    with open(path, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            item = (row.get("管型") or "").strip()
            if not item:
                continue
            expire = to_date((row.get("效期") or "").strip())
            stock = to_int(row.get("库存支数"))
            known.add(item)
            rec = {"管型": item, "批次号": (row.get("批次号") or "").strip(), "效期": expire,
                   "原库存支数": stock, "剩余支数": stock}
            if expire is None or expire < cutoff:
                dropped.append({**rec, "最晚可用日": cutoff})
                continue
            pool.setdefault(item, []).append(rec)
    for batches in pool.values():
        batches.sort(key=lambda b: (b["效期"], b["批次号"]))
    return pool, dropped, known


def allocate(need: int, batches: list):
    """按 FEFO 从库存里领料。先算总账再动料，不够就整项退回，一支都不扣。

    扣了一半最难受：Kit 组不起来（缺一样就不成包），料还被占着，后面的中心也领不到。
    """
    if need <= 0:
        return []
    if sum(b["剩余支数"] for b in batches) < need:
        return None
    parts, left = [], need
    for b in batches:
        take = min(left, b["剩余支数"])
        if take <= 0:
            continue
        b["剩余支数"] -= take
        left -= take
        parts.append({"批次号": b["批次号"], "效期": b["效期"], "领用支数": take})
        if left == 0:
            break
    return parts


def visits_in_order(rows: list) -> list:
    """访视点按方案里的先后排，不按字典序——随访D1 排到筛选期前面就乱套了。"""
    out = []
    for r in rows:
        if r["访视点"] not in out:
            out.append(r["访视点"])
    return out


def alert(code: str, name: str, visit: str, kind: str, note: str) -> dict:
    """一条待办：哪家中心、哪个访视点、什么问题、下一步找谁。"""
    return {"中心编号": code, "中心名称": name, "访视点": visit, "类型": kind, "说明": note}


def plan_center(center: dict, spec_rows: list, pool: dict, known: set) -> dict:
    """算一个中心要组几套 Kit、每样管领多少支。

    套数 = 计划例数——一个受试者来一个访视点领一套，套数不乘访视点数；
    领用支数 = 套数 × 每例管数，这一项才乘。这个乘错一次，标签就多印一整套。
    """
    code, name = center[KEY_COL], center["中心名称"]
    out = {"units": [], "items": [], "labels": [], "alerts": []}
    n = to_int(center["计划例数"])
    if n <= 0:
        out["alerts"].append(alert(code, name, "", KIND_NOT_STARTED,
                                   f"计划例数 = {center['计划例数'] or '空'}，本次不出 Kit"))
        return out
    visits = [r for r in spec_rows if r["访视模板"] == center["访视模板"]]
    if not visits:
        out["alerts"].append(alert(code, name, "", KIND_NO_TEMPLATE,
                                   f"方案规格里没有「{center['访视模板']}」这个模板"))
        return out

    for visit in visits_in_order(visits):
        rows = [r for r in visits if r["访视点"] == visit]
        visit_code = rows[0]["访视编号"]
        planned = sum(r["每例管数"] for r in rows)
        buildable, notes, unit_items = [], [], []

        for r in rows:                              # 先排料，排不上的这一项先不出标签
            item = TUBE_ALIAS.get(r["管型"])
            need = n * r["每例管数"]
            if item is None:
                notes.append(f"{r['管型']}没对应物料")
                out["alerts"].append(alert(code, name, visit, KIND_NO_TUBE,
                                           f"方案里的「{r['管型']}」物料表里没有对应规格，本次没排料"))
                continue
            if item not in known:
                notes.append(f"{r['管型']}物料表里没有")
                out["alerts"].append(alert(code, name, visit, KIND_NO_STOCK,
                                           f"物料表里查不到「{item}」，本次没排料"))
                continue
            batches = pool.get(item, [])
            parts = allocate(need, batches)
            if parts is None:
                have = sum(b["剩余支数"] for b in batches)
                notes.append(f"{r['管型']}缺 {need - have} 支")
                out["alerts"].append(alert(code, name, visit, KIND_SHORT,
                                           f"{item} 只剩 {have} 支，这一项要 {need} 支，"
                                           f"缺 {need - have} 支，等料齐再组"))
                continue
            buildable.append({"spec": r, "item": item, "need": need, "parts": parts})
            for p in parts:
                unit_items.append({KEY_COL: code, "中心名称": name, "访视点": visit,
                                   "管型": r["管型"], "仓库规格": item, "每例管数": r["每例管数"],
                                   "套数": n, "需用支数": need, "批次号": p["批次号"],
                                   "效期": p["效期"].isoformat(), "领用支数": p["领用支数"]})

        # 标签按套发：一套 Kit 一个受试者来一次领一包，包里的管子连着编号，
        # 现场拆包时才能对上是第几个人的第几管。领料能按管型合并，标签不能。
        for seq in range(1, n + 1):
            pos = 0
            for b in buildable:
                for _ in range(b["spec"]["每例管数"]):
                    pos += 1
                    out["labels"].append({
                        "标签编号": f"{code}-{visit_code}-{seq:03d}-{pos:02d}",
                        KEY_COL: code, "中心名称": name, "访视点": visit, "套号": seq,
                        "管位": pos, "管型": b["spec"]["管型"],
                        "采血量mL": b["spec"]["采血量mL"], "处理要求": b["spec"]["处理要求"]})

        built = sum(b["spec"]["每例管数"] for b in buildable)
        out["units"].append({KEY_COL: code, "中心名称": name, "访视点": visit, "套数": n,
                             "计划每套管数": planned, "实际每套管数": built,
                             "内容物支数": n * built, "说明": "；".join(notes)})
        out["items"] += unit_items
    return out


def write_workbook(units: list, items: list, labels: list, alerts: list, out_path: Path) -> Path:
    """四个页签写进一个 xlsx。

    组装清单给组装岗照着组，领料明细给库房照单发料，标签清单直接导去打标机，
    异常与提醒给项目组——要拍板、要催料、要找方案的事都在最后一页。
    """
    wb = Workbook()
    sheets = [("Kit 组装清单", UNIT_COLS, units), ("Kit 领料明细", ITEM_COLS, items),
              ("标签打印清单", LABEL_COLS, labels), ("异常与提醒", ALERT_COLS, alerts)]
    for idx, (title, cols, rows) in enumerate(sheets):
        ws = wb.active if idx == 0 else wb.create_sheet()
        ws.title = title
        ws.append(cols)
        for col in range(1, len(cols) + 1):
            head = ws.cell(row=1, column=col)
            head.font, head.fill = Font(bold=True), HEAD_FILL
            head.alignment = Alignment(horizontal="center")
        for row in rows:
            ws.append([row.get(c, "") for c in cols])
        if title == "异常与提醒":                    # 待办页整行刷色，扫一眼就知道有几条
            for r in range(2, ws.max_row + 1):
                ws.cell(row=r, column=1).fill = BAD_FILL
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{ws.max_row}"
        for i, head in enumerate(cols, 1):
            ws.column_dimensions[get_column_letter(i)].width = max(10, len(head) * 2.2)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out_path))
    return out_path


def main() -> bool:
    spec, conflicts = read_spec(SPEC_FILE)
    plan = read_plan(PLAN_FILE)
    pool, dropped, known = read_stock(STOCK_FILE, BUILD_DATE)
    print(f"方案 {len(spec)} 行、中心 {len(plan)} 个、可排料的物料规格 {len(pool)} 种")

    alerts = [alert("", "", c["访视点"], KIND_DUP,
                    f"「{c['模板']}」{c['访视点']}的「{c['管型']}」有两行，"
                    f"本次按 {c['用到的']} 支算，另一行写的是 {c['被忽略的']} 支，回去核方案")
              for c in conflicts]
    alerts += [alert("", "", "", KIND_SHELF,
                     f"「{d['管型']}」批次 {d['批次号']} 效期 {d['效期']} 盖不到 "
                     f"{d['最晚可用日']}，本次没排它")
               for d in dropped]

    units, items, labels = [], [], []
    for center in plan:
        got = plan_center(center, spec, pool, known)
        units += got["units"]
        items += got["items"]
        labels += got["labels"]
        alerts += got["alerts"]

    out = write_workbook(units, items, labels, alerts, OUT_FILE)
    print(f"  Kit 套数：{sum(u['套数'] for u in units)} 套，覆盖 {len({u[KEY_COL] for u in units})} 个中心")
    print(f"  领用支数：{sum(i['领用支数'] for i in items)} 支，分 {len(items)} 个批次行")
    print(f"  标签：{len(labels)} 张；异常与提醒：{len(alerts)} 条")
    print(f"Kit 配置清单已生成：{out.name}")

    # ---------- 以下为校验，正文里没有 ----------
    print("\n===== 回读校验 =====")
    from openpyxl import load_workbook
    book = load_workbook(str(out))
    sh_unit, sh_item, sh_label, sh_alert = (book["Kit 组装清单"], book["Kit 领料明细"],
                                            book["标签打印清单"], book["异常与提醒"])

    def col(sheet, name):
        idx = [c.value for c in sheet[1]].index(name) + 1
        return [sheet.cell(row=i, column=idx).value for i in range(2, sheet.max_row + 1)]

    def rows_of(sheet):
        names = [c.value for c in sheet[1]]
        return [dict(zip(names, [sheet.cell(row=i, column=j + 1).value
                                 for j in range(len(names))]))
                for i in range(2, sheet.max_row + 1)]

    u_rows, i_rows, l_rows, a_rows = rows_of(sh_unit), rows_of(sh_item), rows_of(sh_label), rows_of(sh_alert)
    live = {c for c in col(sh_unit, KEY_COL)}                       # 真出了 Kit 的中心
    label_pos = {}
    for r in l_rows:
        label_pos.setdefault((r[KEY_COL], r["访视点"], r["套号"]), []).append(r["管位"])
    shelf = [r for r in a_rows if r["类型"] == KIND_SHELF]
    short_note = [r for r in a_rows if r["类型"] == KIND_SHORT]
    by_batch = {}
    for r in i_rows:
        by_batch[r["批次号"]] = by_batch.get(r["批次号"], 0) + r["领用支数"]

    checks = [
        ("组装清单 = 4 个有效中心 × 3 个访视点，没启动和模板对不上的中心不出行",
         sh_unit.max_row == 13 and live == {"01001", "01002", "01004", "01005"},
         f"{sh_unit.max_row - 1} 行 / 中心 {sorted(live)}"),
        ("套数 = 计划例数，不乘以访视点数",
         all(r["套数"] == {"01001": 12, "01002": 8, "01004": 5, "01005": 3}[r[KEY_COL]]
             for r in u_rows),
         f"01001 三个访视点都是 12 套：{[r['套数'] for r in u_rows if r[KEY_COL] == '01001']}"),
        ("中心编号前导零没丢（01001 原样在表里）",
         "01001" in col(sh_item, KEY_COL) and all(len(c) == 5 for c in live), "01001"),
        ("需用支数 = 套数 × 每例管数（01001 基线促凝管 = 12 × 2）",
         any(r[KEY_COL] == "01001" and r["访视点"] == "基线" and r["管型"] == "促凝管"
             and r["每例管数"] == 2 and r["需用支数"] == 24 for r in i_rows),
         24),
        ("促凝管按 FEFO 拆两批：B2509B 60 支 + B2601A 41 支",
         by_batch.get("B2509B") == 60 and by_batch.get("B2601A") == 41,
         f"B2509B {by_batch.get('B2509B')} 支 / B2601A {by_batch.get('B2601A')} 支"),
        ("效期盖不到采样窗口的批次一支都没排出去",
         "B2302C" not in by_batch and "B2301D" not in by_batch and len(shelf) == 2,
         "B2302C / B2301D 都没排，各留一条提醒"),
        ("规格相近的干扰料没被误领（内旋冻存管、3.8% 凝血管）",
         not any(r["仓库规格"] in ("冻存管 2mL 内旋", "蓝帽 柠檬酸钠 3.8%") for r in i_rows),
         "规格名对不上就不猜"),
        ("冻存管库存不足被标出来，差 4 支，成都中心那项没排料",
         len(short_note) == 1 and "缺 4 支" in short_note[0]["说明"]
         and sum(r["领用支数"] for r in i_rows if r["管型"] == "冻存管") == 24,
         short_note[0]["说明"]),
        ("每支管子一张标签：标签数 = 领用支数合计",
         len(l_rows) == sum(r["领用支数"] for r in i_rows) == 254,
         f"{len(l_rows)} 张 / 领用 {sum(r['领用支数'] for r in i_rows)} 支"),
        ("标签编号唯一，同套内管位从 1 连到「实际每套管数」",
         len({r["标签编号"] for r in l_rows}) == len(l_rows)
         and all(sorted(v) == list(range(1, len(v) + 1)) for v in label_pos.values()),
         f"{len(l_rows)} 个编号 / {len(label_pos)} 套"),
        ("标签按套展开：01001 筛选期 = 12 套 × 3 支 = 36 张",
         sum(1 for r in l_rows if r[KEY_COL] == "01001" and r["访视点"] == "筛选期") == 36,
         36),
        ("方案里没有对应物料的管型（血沉管）只进异常、不排料不贴标签",
         sum(1 for r in a_rows if r["类型"] == KIND_NO_TUBE) == 2
         and not any(r["管型"] == "血沉管" for r in i_rows)
         and not any(r["管型"] == "血沉管" for r in l_rows),
         "2 条异常"),
        ("方案修订留下的重复行被记出来，且每例管数取先读到的那条",
         len(conflicts) == 1 and conflicts[0]["管型"] == "促凝管"
         and conflicts[0]["用到的"] == 2 and conflicts[0]["被忽略的"] == 1
         and sum(1 for r in a_rows if r["类型"] == KIND_DUP) == 1,
         f"按 {conflicts[0]['用到的']} 支算，另一行 {conflicts[0]['被忽略的']} 支"),
        ("例数为 0 的中心进异常清单，且没有领料和标签",
         "01003" not in live and any(r[KEY_COL] == "01003" and r["类型"] == KIND_NOT_STARTED
                                     for r in a_rows),
         "01003 不出 Kit"),
        ("访视模板在方案里查不到的中心也进异常清单",
         any(r[KEY_COL] == "01006" and r["类型"] == KIND_NO_TEMPLATE for r in a_rows),
         "01006 模板「加强」对不上"),
        ("表尾的合计行、说明行没被当成数据（方案 17 行、中心 6 个）",
         len(spec) == 17 and len(plan) == 6, f"{len(spec)} 行 / {len(plan)} 个中心"),
        ("库存是真扣的：每个批次排出去的支数都不超过原库存",
         all(by_batch.get(b["批次号"], 0) <= b["原库存支数"]
             for batches in pool.values() for b in batches),
         "按批次核过一遍"),
        ("异常与提醒 8 条：重复行 1 + 批次跳过 2 + 未启动 1 + 模板对不上 1 + 管型没对上 2 + 缺料 1",
         len(a_rows) == 8,
         f"{len(a_rows)} 条"),
    ]
    ok = True
    for name, passed, val in checks:
        print(f"  [{'OK ' if passed else 'FAIL'}] {name}  ->  {val}")
        ok &= bool(passed)

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
