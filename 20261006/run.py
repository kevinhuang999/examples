# -*- coding: utf-8 -*-
"""仪器导出结果 → 自动入库（SQLite 单文件库）。

跑法：
    python run.py              # 双击也行
    python run.py --no-pause   # 跑完不等回车

做什么：
    把 01_raw_data/ 下两批仪器导出文件按统一结构写进 02_output/lab_results.db。
    重复导入不重复入库（幂等）；复检重出保留旧版本、不覆盖；同一天同一个键出两个不同
    结果时，两条都不许当有效值用。

工程外壳（正文里不逐行讲）：跑完自动回读断言打印 PASS/FAIL；日志追加写 source/run_log.txt，
每段开头记运行主机（主机名 + 本机 IP）。本示例只用标准库，唯一要求是 SQLite >= 3.24
（ON CONFLICT DO UPDATE 要用），启动时会检查。
"""

from __future__ import annotations

import argparse
import csv
import socket
import sqlite3
import sys
import time
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
RAW_DIR = ROOT / "01_raw_data"
OUT_DIR = ROOT / "02_output"
SRC_DIR = ROOT / "source"
LOG_PATH = SRC_DIR / "run_log.txt"
DB_PATH = OUT_DIR / "lab_results.db"
LEDGER_PATH = OUT_DIR / "入库台账.csv"

# ---------------------------------------------------------- ① 配置：路径与字段

BATCHES = [
    ("B1-20261005", RAW_DIR / "仪器导出_20261005"),
    ("B2-20261006", RAW_DIR / "仪器导出_20261006"),
    # 同一批文件再跑一遍：用来证明重复导入是幂等的（新增、更新都应为 0）
    ("B2-20261006-重放", RAW_DIR / "仪器导出_20261006"),
]

# 一种导出格式一条记录。新仪器来了就在这里加一行，不要改判定逻辑。
FORMATS = [
    {
        "名称": "仪器导出_逗号分隔",
        "后缀": (".csv",),
        "编码": ("utf-8-sig", "gbk"),
        "分隔符": ",",
        "必需列": ("样品编号", "检测项目", "结果", "单位", "检测日期", "仪器"),
        "判空列": "检测项目",
        "列映射": {
            "样品编号": "样品编号",
            "检测项目": "检测项目",
            "结果原值": "结果",
            "单位": "单位",
            "检测日期": "检测日期",
            "仪器": "仪器",
        },
    },
    {
        "名称": "仪器导出_制表符分隔",
        "后缀": (".txt",),
        "编码": ("gbk", "utf-8-sig"),
        "分隔符": "\t",
        "必需列": ("样品编号", "检测项目", "结果", "单位", "检测日期", "仪器"),
        "判空列": "检测项目",
        "列映射": {
            "样品编号": "样品编号",
            "检测项目": "检测项目",
            "结果原值": "结果",
            "单位": "单位",
            "检测日期": "检测日期",
            "仪器": "仪器",
        },
    },
]

# 入库主键：这三列一起决定"是不是同一条结果"。不含来源文件、不含仪器。
KEY_COLS = ("样品编号", "检测项目", "单位")

STATUS_NUM = "数值"
STATUS_LIMIT = "低于下限"
STATUS_TEXT = "文字结论"
STATUS_EMPTY = "空"

FLAG_OK = "正常"
FLAG_CONFLICT = "冲突待确认"

# 入库存量动作
ACT_NEW, ACT_SAME, ACT_UPDATE = "新增", "无变化", "更新"
ACT_OLD, ACT_CONFLICT = "旧版本跳过", "冲突"

RESULTS_COLS = (
    "样品编号", "检测项目", "单位", "结果原值", "结果数值", "结果状态",
    "检测日期", "仪器", "来源文件", "入库批次", "版本", "记录状态",
    "首次入库时间", "最后更新",
)

DDL = """
CREATE TABLE IF NOT EXISTS results (
    样品编号 TEXT NOT NULL,
    检测项目 TEXT NOT NULL,
    单位     TEXT NOT NULL,
    结果原值 TEXT,
    结果数值 REAL,
    结果状态 TEXT NOT NULL,
    检测日期 TEXT NOT NULL,
    仪器     TEXT,
    来源文件 TEXT,
    入库批次 TEXT,
    版本     INTEGER NOT NULL DEFAULT 1,
    记录状态 TEXT NOT NULL DEFAULT '正常',
    首次入库时间 TEXT NOT NULL,
    最后更新   TEXT NOT NULL,
    PRIMARY KEY (样品编号, 检测项目, 单位)
);
CREATE TABLE IF NOT EXISTS result_history (
    样品编号 TEXT NOT NULL,
    检测项目 TEXT NOT NULL,
    单位     TEXT NOT NULL,
    结果原值 TEXT,
    结果数值 REAL,
    结果状态 TEXT,
    检测日期 TEXT,
    仪器     TEXT,
    版本     INTEGER,
    首次入库时间 TEXT,
    最后更新   TEXT,
    替换时间   TEXT NOT NULL,
    替换批次   TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS ingest_pending (
    入库批次 TEXT NOT NULL,
    来源文件 TEXT NOT NULL,
    行号     INTEGER,
    样品编号 TEXT,
    检测项目 TEXT,
    原因     TEXT NOT NULL,
    记录时间 TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS ingest_log (
    批次 TEXT PRIMARY KEY,
    开始时间 TEXT, 结束时间 TEXT, 用时秒 REAL,
    文件数 INTEGER, 读出行数 INTEGER,
    新增 INTEGER, 更新 INTEGER, 无变化 INTEGER,
    旧版本跳过 INTEGER, 冲突 INTEGER, 挂起 INTEGER,
    判定 TEXT
);
"""


def now_str() -> str:
    return time.strftime("%Y-%m-%d %H:%M:%S")


# ------------------------------------------------------------------ 运行日志

def host_tag() -> str:
    """运行主机：主机名（本机 IP）。IP 取不到就写"未知"。"""
    try:
        name = socket.gethostname()
    except OSError:
        name = "未知"
    ip = "未知"
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        try:
            s.connect(("8.8.8.8", 80))
            ip = s.getsockname()[0]
        finally:
            s.close()
    except OSError:
        try:
            ip = socket.gethostbyname(socket.gethostname())
        except OSError:
            ip = "未知"
    return f"{name}（{ip}）"


def log_open() -> None:
    SRC_DIR.mkdir(parents=True, exist_ok=True)
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(f"运行主机：{host_tag()}\n")
        f.write(f"开始执行：{now_str()}  Python {sys.version.split()[0]}（{sys.executable}）\n")


def log_close(started: float, verdict: str) -> None:
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(f"执行完成：{now_str()}  用时 {time.time() - started:.1f}s  判定 {verdict}\n\n")


# ---------------------------------------------------------------- ② 读数据

def read_text(path: Path, encodings) -> str:
    """按候选编码依次试。GBK 文件用 utf-8 读会直接抛 UnicodeDecodeError，不是乱码。"""
    last = None
    for enc in encodings:
        try:
            return path.read_text(encoding=enc)
        except UnicodeDecodeError as exc:
            last = exc
    raise last


def pick_format(path: Path):
    """先看后缀，再用表头特征列确认——不靠文件名、不靠表头行号。"""
    lines = None
    for fmt in FORMATS:
        if path.suffix.lower() not in fmt["后缀"]:
            continue
        if lines is None:
            lines = read_text(path, fmt["编码"]).splitlines()
        for raw in lines[:10]:
            cells = [c.strip() for c in raw.split(fmt["分隔符"])]
            if all(col in cells for col in fmt["必需列"]):
                return fmt, lines
    return None, None


def read_export(path: Path, fmt, lines) -> list[dict]:
    """切出表头行之后的数据行，按判空列丢掉表尾说明行。"""
    header_at = None
    for i, raw in enumerate(lines):
        cells = [c.strip() for c in raw.split(fmt["分隔符"])]
        if all(col in cells for col in fmt["必需列"]):
            header_at = i
            break
    if header_at is None:
        return []
    rows = list(csv.reader(lines[header_at:], delimiter=fmt["分隔符"]))
    header = [c.strip() for c in rows[0]]
    out = []
    for line_no, cells in enumerate(rows[1:], start=header_at + 2):
        rec = {h: (cells[i] if i < len(cells) else "") for i, h in enumerate(header)}
        if not rec.get(fmt["判空列"], "").strip():
            continue                      # 表尾说明行 / 空行：判空列一空就到头
        item = {dst: rec.get(src, "") for dst, src in fmt["列映射"].items()}
        item["_来源文件"] = path.name
        item["_行号"] = line_no
        out.append(item)
    return out


# ------------------------------------------------------- ③ 归一：结果与键

def norm(text: str) -> str:
    """去掉所有空白，含全角空格 U+3000（strip 不管它）。"""
    return "".join(str(text or "").split())


def parse_result(raw: str):
    """结果原值 → (数值, 状态)。三种形态分开存，谁也不许顶替谁。"""
    text = norm(raw)
    if not text:
        return None, STATUS_EMPTY
    if text[0] in "<≤≦":
        return None, STATUS_LIMIT
    try:
        return float(text), STATUS_NUM
    except ValueError:
        return None, STATUS_TEXT


def parse_date(raw: str):
    """2026-10-05 / 2026/10/05 / 20261005 都认；认不出返回 None，不兜今天。"""
    text = norm(raw)
    for pattern in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y%m%d"):
        try:
            return datetime.strptime(text, pattern).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return None


def normalize_row(row: dict):
    """键补不齐就退回挂起，不让半条记录进库。"""
    item = {
        "样品编号": str(row.get("样品编号", "")).strip(),
        "检测项目": norm(row.get("检测项目", "")),
        "单位": norm(row.get("单位", "")) or "-",
        "结果原值": str(row.get("结果原值", "")).strip(),
        "检测日期": parse_date(row.get("检测日期", "")),
        "仪器": str(row.get("仪器", "")).strip(),
        "来源文件": row.get("_来源文件", ""),
        "_行号": row.get("_行号"),
    }
    if not item["样品编号"]:
        return None, "样品编号为空", item
    if not item["检测项目"]:
        return None, "检测项目为空", item
    if item["检测日期"] is None:
        return None, "检测日期认不出", item
    item["结果数值"], item["结果状态"] = parse_result(item["结果原值"])
    return item, None, item


# ------------------------------------------------------- ④ 入库：五个分支

def find_existing(conn, rec):
    sql = "SELECT * FROM results WHERE 样品编号=? AND 检测项目=? AND 单位=?"
    return conn.execute(sql, tuple(rec[c] for c in KEY_COLS)).fetchone()


def decide(existing, rec) -> str:
    """判定顺序＝业务顺序：没有 → 同日同值 → 同日不同值（冲突）→ 更晚（更新）→ 更早（跳过）。"""
    if existing is None:
        return ACT_NEW
    if existing["检测日期"] == rec["检测日期"]:
        return ACT_SAME if existing["结果原值"] == rec["结果原值"] else ACT_CONFLICT
    if rec["检测日期"] > existing["检测日期"]:
        return ACT_UPDATE
    return ACT_OLD


def insert_new(conn, rec, batch: str, stamp: str) -> None:
    conn.execute(
        f"INSERT INTO results ({','.join(RESULTS_COLS)}) VALUES ({','.join('?' * len(RESULTS_COLS))})",
        (rec["样品编号"], rec["检测项目"], rec["单位"], rec["结果原值"], rec["结果数值"],
         rec["结果状态"], rec["检测日期"], rec["仪器"], rec["来源文件"], batch, 1, FLAG_OK,
         stamp, stamp),
    )


def update_version(conn, old, rec, batch: str, stamp: str) -> None:
    """先把旧行整条搬进 history，再原地更新——落库不是覆盖，是新增一个版本。

    注意这里用的是 ON CONFLICT DO UPDATE，不能图省事写 INSERT OR REPLACE：
    后者会删掉旧行再插新行，rowid 变了、首次入库时间被清成现在，留痕也一起丢。
    """
    conn.execute(
        "INSERT INTO result_history (样品编号,检测项目,单位,结果原值,结果数值,结果状态,"
        "检测日期,仪器,版本,首次入库时间,最后更新,替换时间,替换批次) "
        "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)",
        (old["样品编号"], old["检测项目"], old["单位"], old["结果原值"], old["结果数值"],
         old["结果状态"], old["检测日期"], old["仪器"], old["版本"], old["首次入库时间"],
         old["最后更新"], stamp, batch),
    )
    conn.execute(
        "INSERT INTO results (样品编号,检测项目,单位,结果原值,结果数值,结果状态,检测日期,"
        "仪器,来源文件,入库批次,版本,记录状态,首次入库时间,最后更新) "
        "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?) "
        "ON CONFLICT(样品编号,检测项目,单位) DO UPDATE SET "
        "结果原值=excluded.结果原值, 结果数值=excluded.结果数值, 结果状态=excluded.结果状态, "
        "检测日期=excluded.检测日期, 仪器=excluded.仪器, 来源文件=excluded.来源文件, "
        "入库批次=excluded.入库批次, 版本=excluded.版本, 记录状态=excluded.记录状态, "
        "最后更新=excluded.最后更新",
        (rec["样品编号"], rec["检测项目"], rec["单位"], rec["结果原值"], rec["结果数值"],
         rec["结果状态"], rec["检测日期"], rec["仪器"], rec["来源文件"], batch,
         old["版本"] + 1, FLAG_OK, old["首次入库时间"], stamp),
    )


def mark_conflict(conn, existing, rec, batch: str, stamp: str) -> None:
    """同日同键两个不同结果：库里的标"冲突待确认"，新来的进挂起，两条都不当有效值。"""
    conn.execute(
        "UPDATE results SET 记录状态=?, 最后更新=? WHERE 样品编号=? AND 检测项目=? AND 单位=?",
        (FLAG_CONFLICT, stamp, existing["样品编号"], existing["检测项目"], existing["单位"]),
    )


def add_pending(conn, batch: str, rec: dict, reason: str, stamp: str) -> None:
    conn.execute(
        "INSERT INTO ingest_pending (入库批次,来源文件,行号,样品编号,检测项目,原因,记录时间) "
        "VALUES (?,?,?,?,?,?,?)",
        (batch, rec.get("来源文件", ""), rec.get("_行号"), rec.get("样品编号", ""),
         rec.get("检测项目", ""), reason, stamp),
    )


def run_batch(conn, batch: str, folder: Path) -> dict:
    """一个批次＝一个事务。解析不了的行进挂起，不算失败；日志与数据同事务，一起提交。"""
    started = time.time()
    stats = {"文件数": 0, "读出行数": 0, ACT_NEW: 0, ACT_UPDATE: 0, ACT_SAME: 0,
             ACT_OLD: 0, ACT_CONFLICT: 0, "挂起": 0}
    stamp = now_str()
    with conn:                                    # 整批一起提交，出错整批回滚
        for path in sorted(folder.iterdir()):
            if not path.is_file():
                continue
            fmt, lines = pick_format(path)
            if fmt is None:
                stats["挂起"] += 1
                add_pending(conn, batch, {"来源文件": path.name}, "文件格式认不出", stamp)
                continue
            stats["文件数"] += 1
            for row in read_export(path, fmt, lines):
                stats["读出行数"] += 1
                rec, reason, item = normalize_row(row)
                if rec is None:
                    stats["挂起"] += 1
                    add_pending(conn, batch, item, reason, stamp)
                    continue
                existing = find_existing(conn, rec)
                action = decide(existing, rec)
                stats[action] += 1
                if action == ACT_NEW:
                    insert_new(conn, rec, batch, stamp)
                elif action == ACT_UPDATE:
                    update_version(conn, existing, rec, batch, stamp)
                elif action == ACT_CONFLICT:
                    mark_conflict(conn, existing, rec, batch, stamp)
                    add_pending(conn, batch, rec, "与库中记录结果冲突", stamp)
        verdict = "PASS" if (stats[ACT_NEW] or stats[ACT_UPDATE] or stats[ACT_SAME]) else "FAIL"
        conn.execute(
            "INSERT INTO ingest_log (批次,开始时间,结束时间,用时秒,文件数,读出行数,新增,更新,"
            "无变化,旧版本跳过,冲突,挂起,判定) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)",
            (batch, stamp, now_str(), round(time.time() - started, 2), stats["文件数"],
             stats["读出行数"], stats[ACT_NEW], stats[ACT_UPDATE], stats[ACT_SAME],
             stats[ACT_OLD], stats[ACT_CONFLICT], stats["挂起"], verdict),
        )
    return stats


def write_ledger(conn) -> None:
    rows = conn.execute(
        "SELECT 批次,读出行数,新增,更新,无变化,旧版本跳过,冲突,挂起,判定 FROM ingest_log ORDER BY 开始时间"
    ).fetchall()
    with open(LEDGER_PATH, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["批次", "读出行数", "新增", "更新", "无变化", "旧版本跳过", "冲突", "挂起", "判定"])
        for r in rows:
            w.writerow(list(r))


# ---------------------------------------------------------------- 回读校验

def verify(conn) -> list[tuple[str, bool, str]]:
    checks = []

    def add(name, ok, detail=""):
        checks.append((name, bool(ok), detail))

    def one(sql, args=()):
        return conn.execute(sql, args).fetchone()[0]

    def row(key):
        return conn.execute(
            "SELECT * FROM results WHERE 样品编号=? AND 检测项目=? AND 单位=?", key
        ).fetchone()

    add("results 共 13 行", one("SELECT COUNT(*) FROM results") == 13,
        f"实际 {one('SELECT COUNT(*) FROM results')}")
    add("ingest_log 共 3 批", one("SELECT COUNT(*) FROM ingest_log") == 3)
    add("result_history 共 3 条旧版本", one("SELECT COUNT(*) FROM result_history") == 3,
        f"实际 {one('SELECT COUNT(*) FROM result_history')}")
    add("ingest_pending 共 7 行", one("SELECT COUNT(*) FROM ingest_pending") == 7,
        f"实际 {one('SELECT COUNT(*) FROM ingest_pending')}")

    b1 = conn.execute("SELECT 读出行数,新增,更新,挂起 FROM ingest_log WHERE 批次='B1-20261005'").fetchone()
    add("B1：读出 12 行、新增 11、挂起 1",
        tuple(b1) == (12, 11, 0, 1), f"实际 {tuple(b1)}")
    b2 = conn.execute(
        "SELECT 读出行数,新增,更新,无变化,旧版本跳过,冲突,挂起 FROM ingest_log WHERE 批次='B2-20261006'"
    ).fetchone()
    add("B2：读出 9 行、新增 2、更新 3、旧版本跳过 1、冲突 1、挂起 2",
        tuple(b2) == (9, 2, 3, 0, 1, 1, 2), f"实际 {tuple(b2)}")
    b3 = conn.execute(
        "SELECT 新增,更新,无变化 FROM ingest_log WHERE 批次='B2-20261006-重放'"
    ).fetchone()
    add("重放批次：新增 0、更新 0、无变化 5（幂等）", tuple(b3) == (0, 0, 5), f"实际 {tuple(b3)}")

    add("前导零样品编号在库", row(("0000081", "含量(%)", "%")) is not None)
    add("样品编号未被剥成 81", one("SELECT COUNT(*) FROM results WHERE 样品编号='81'") == 0)

    r83 = row(("0000083", "有关物质(%)", "%"))
    add("0000083 有关物质更新到 0.032、版本 2",
        r83 and norm(r83["结果原值"]) == "0.032" and r83["版本"] == 2,
        f"实际 {r83['结果原值'] if r83 else None} / v{r83['版本'] if r83 else None}")
    h83 = conn.execute(
        "SELECT 结果原值,结果数值,结果状态 FROM result_history WHERE 样品编号='0000083'"
    ).fetchone()
    add("history 留下旧的 <0.01（数值为空、状态=低于下限）",
        h83 and norm(h83["结果原值"]) == "<0.01" and h83["结果数值"] is None
        and h83["结果状态"] == STATUS_LIMIT, f"实际 {tuple(h83) if h83 else None}")

    r82 = row(("0000082", "含量(%)", "%"))
    add("0000082 含量取更晚的 99.35（旧 99.10 进 history）",
        r82 and norm(r82["结果原值"]) == "99.35"
        and one("SELECT COUNT(*) FROM result_history WHERE 样品编号='0000082' AND 结果原值='99.10'") == 1)
    add("更新不重置首次入库时间（history 与主表一致）",
        one("SELECT COUNT(*) FROM result_history h JOIN results r "
            "ON r.样品编号=h.样品编号 AND r.检测项目=h.检测项目 AND r.单位=h.单位 "
            "WHERE r.首次入库时间=h.首次入库时间") == 3)
    add("0000081 残留溶剂更新到 0.033、版本 2",
        (row(("0000081", "残留溶剂(%)", "%")) or {"版本": 0})["版本"] == 2)

    r84 = row(("0000084", "有关物质(%)", "%"))
    add("冲突行标记冲突待确认且保留旧值",
        r84 and r84["记录状态"] == FLAG_CONFLICT and norm(r84["结果原值"]) == "未检出",
        f"实际 {r84['记录状态'] if r84 else None}/{r84['结果原值'] if r84 else None}")
    add("冲突的新结果没进库",
        one("SELECT COUNT(*) FROM results WHERE 结果原值='0.028'") == 0)

    r85 = row(("0000085", "有关物质(%)", "%"))
    add("结果格留空 → 状态=空、无数值",
        r85 and r85["结果状态"] == STATUS_EMPTY and r85["结果数值"] is None)
    add("未检出 → 状态=文字结论、无数值",
        (row(("0000084", "有关物质(%)", "%")) or {"结果状态": ""})["结果状态"] in (STATUS_TEXT,))

    r91 = row(("0000091", "含量(%)", "-"))
    add("单位留空归一成 - 后入库", r91 is not None and abs(r91["结果数值"] - 98.0) < 1e-9)
    add("斜杠日期归一成 2026-10-05",
        (row(("0000088", "水分(%)", "%")) or {"检测日期": ""})["检测日期"] == "2026-10-05")

    dup = one("SELECT COUNT(*) FROM (SELECT 样品编号,检测项目,单位 FROM results "
              "GROUP BY 样品编号,检测项目,单位 HAVING COUNT(*)>1)")
    add("主键无重复", dup == 0, f"重复 {dup} 条")
    avg_hl = one("SELECT AVG(结果数值) FROM results WHERE 检测项目='含量(%)'")
    add("含量均值只由 4 条数值算出（98.4425）",
        avg_hl is not None and abs(avg_hl - 98.4425) < 1e-6, f"实际 {avg_hl}")
    avg_yg = one("SELECT AVG(结果数值) FROM results WHERE 检测项目='有关物质(%)'")
    add("有关物质均值=0.032（未检出/空/冲突都不计入）",
        avg_yg is not None and abs(avg_yg - 0.032) < 1e-9, f"实际 {avg_yg}")

    reasons = dict(conn.execute("SELECT 原因, COUNT(*) FROM ingest_pending GROUP BY 原因").fetchall())
    add("挂起原因分布：编号空 3 / 日期认不出 2 / 冲突 2",
        reasons == {"样品编号为空": 3, "检测日期认不出": 2, "与库中记录结果冲突": 2},
        f"实际 {reasons}")
    add("三批判定全 PASS", one("SELECT COUNT(*) FROM ingest_log WHERE 判定='PASS'") == 3)
    add("每行都留了来源文件与首次入库时间",
        one("SELECT COUNT(*) FROM results WHERE COALESCE(来源文件,'')='' OR COALESCE(首次入库时间,'')=''") == 0)
    add("库文件已落盘且非空", DB_PATH.exists() and DB_PATH.stat().st_size > 0)
    add("入库台账.csv 已生成", LEDGER_PATH.exists() and LEDGER_PATH.stat().st_size > 0)
    return checks


def main() -> int:
    started = time.time()
    parser = argparse.ArgumentParser()
    parser.add_argument("--no-pause", action="store_true")
    args = parser.parse_args()

    if sqlite3.sqlite_version_info < (3, 24):
        print(f"SQLite {sqlite3.sqlite_version} 太旧，ON CONFLICT DO UPDATE 需要 3.24+")
        return 2

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    SRC_DIR.mkdir(parents=True, exist_ok=True)
    log_open()

    if DB_PATH.exists():
        DB_PATH.unlink()                      # 每次重跑重建库，保证结果可复现
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.executescript(DDL)
    print(f"SQLite {sqlite3.sqlite_version}  库：{DB_PATH.name}\n")

    for batch, folder in BATCHES:
        stats = run_batch(conn, batch, folder)
        print(f"[{batch}] 文件 {stats['文件数']} 个，读出 {stats['读出行数']} 行 → "
              f"新增 {stats[ACT_NEW]}、更新 {stats[ACT_UPDATE]}、无变化 {stats[ACT_SAME]}、"
              f"旧版本跳过 {stats[ACT_OLD]}、冲突 {stats[ACT_CONFLICT]}、挂起 {stats['挂起']}")
    write_ledger(conn)
    print(f"\n入库台账：{LEDGER_PATH.name}")
    print(f"结果明细：{conn.execute('SELECT COUNT(*) FROM results').fetchone()[0]} 行 / "
          f"历史版本 {conn.execute('SELECT COUNT(*) FROM result_history').fetchone()[0]} 条 / "
          f"挂起 {conn.execute('SELECT COUNT(*) FROM ingest_pending').fetchone()[0]} 条\n")

    checks = verify(conn)
    for name, ok, detail in checks:
        print(f"[{'OK ' if ok else 'NG '}] {name}" + (f"  → {detail}" if detail and not ok else ""))
    bad = [c for c in checks if not c[1]]
    verdict = "PASS" if not bad else "FAIL"
    print(f"\n回读校验：{len(checks) - len(bad)}/{len(checks)} 项通过  总判定 {verdict}")
    conn.close()

    log_close(started, verdict)
    if not args.no_pause:
        input("按回车结束……")
    return 0 if verdict == "PASS" else 1


if __name__ == "__main__":
    sys.exit(main())
