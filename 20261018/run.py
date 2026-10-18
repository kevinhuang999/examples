"""实验室文件命名规范落地：把文件夹里五花八门的旧文件名，按规则表批量改成规范名。

规范格式：八位日期_样品编号_检项.ext
用法：python run.py            # 跑完停在窗口里
      python run.py --no-pause # 跑完直接退出（脚本调用）
"""

import csv
import re
import socket
import sys
import time
from datetime import datetime
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
WORK_DIR = BASE_DIR / "01_raw_data" / "待整理"
OUT_DIR = BASE_DIR / "02_output"
LOG_PATH = BASE_DIR / "source" / "run_log.txt"

PREVIEW_FILE = OUT_DIR / "改名预览.csv"
RENAME_LOG = OUT_DIR / "rename_log.csv"

# 规范名长什么样：八位日期_样品编号_检项.扩展名，匹配上就算已规范
STD_PATTERN = re.compile(r"^\d{8}_[A-Za-z]\d+_.+\.\w+$")

# 四种状态：只动「改名」，其余三种一律原地保留给人看
ST_OK = "改名"
ST_STD = "跳过_已规范"
ST_UNKNOWN = "挂起_未识别"
ST_DUP = "冲突_新名重复"

# 规则表：实验室里常见的几种旧命名习惯，一条正则一种；来新花样就在这里加一行
RULES = [
    ("日期_样品_检项", re.compile(
        r"(?P<date>[\d年月日.\-]+)_(?:样品)?(?P<sample>[A-Za-z]\d+)_(?P<item>[^ ]+?)\.(?P<ext>\w+)$")),
    ("英文_样品_检项_日期", re.compile(
        r"sample\s+(?P<sample>[A-Za-z]\d+)\s+(?P<item>[A-Za-z]+)\s+(?P<date>[\d.\-/]+)\.(?P<ext>\w+)$")),
    ("检项_日期_样品_副本", re.compile(
        r"(?P<item>[^ ]+?)\s+(?P<date>[\d.]+)\s+(?:样品)?(?P<sample>[A-Za-z]\d+)\s*\(\d+\)\.(?P<ext>\w+)$")),
]

# 检项别名：仪器吐的英文缩写归到统一中文名
ITEM_ALIAS = {"CBC": "血常规", "GLU": "血糖"}


def host_label():
    """日志里记下是哪台机器跑的：主机名 + 本机对外网卡地址。"""
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


def write_log(line):
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def normalize_date(text):
    """2026-3-5 / 2026.3.5 / 2026年3月5日 → 20260305，日期口径先统一再进文件名。"""
    m = re.search(r"(\d{4})[年.\-/]?(\d{1,2})[月.\-/]?(\d{1,2})日?", text)
    return f"{int(m.group(1)):04d}{int(m.group(2)):02d}{int(m.group(3)):02d}"


def match_rule(filename):
    """拿规则表逐条试，谁先匹配上算谁的；都匹配不上返回 None，挂起给人看。"""
    for rule_name, pattern in RULES:
        m = pattern.match(filename)
        if m:
            fields = m.groupdict()
            fields["item"] = ITEM_ALIAS.get(fields["item"], fields["item"])
            fields["date"] = normalize_date(fields["date"])
            return rule_name, fields
    return None


def make_new_name(fields):
    """拼规范名：八位日期_样品编号_检项.扩展名。"""
    return f"{fields['date']}_{fields['sample']}_{fields['item']}.{fields['ext']}"


def plan_renames():
    """扫目录 → 逐个识别 → 拼新名 → 查重，产出整张预览表。这一步不动任何文件。"""
    plan = []
    claimed = set()  # 本轮已认走的规范名，后来者撞上就是冲突
    for path in sorted(WORK_DIR.iterdir()):
        if not path.is_file():
            continue
        if STD_PATTERN.match(path.name):
            plan.append({"旧名": path.name, "新名": path.name, "状态": ST_STD, "规则": "-"})
            continue
        hit = match_rule(path.name)
        if hit is None:
            plan.append({"旧名": path.name, "新名": "", "状态": ST_UNKNOWN, "规则": "-"})
            continue
        rule_name, fields = hit
        new_name = make_new_name(fields)
        if new_name in claimed or (WORK_DIR / new_name).exists():
            plan.append({"旧名": path.name, "新名": new_name, "状态": ST_DUP, "规则": rule_name})
            continue
        claimed.add(new_name)
        plan.append({"旧名": path.name, "新名": new_name, "状态": ST_OK, "规则": rule_name})
    return plan


def write_preview(plan):
    OUT_DIR.mkdir(exist_ok=True)
    with open(PREVIEW_FILE, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=["旧名", "新名", "状态", "规则"])
        writer.writeheader()
        writer.writerows(plan)


def apply_plan(plan):
    """只动状态为「改名」的文件，动过的每一笔都写进 rename_log.csv，随时可回溯。"""
    done = 0
    new_file = not RENAME_LOG.exists()
    with open(RENAME_LOG, "a", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        if new_file:
            writer.writerow(["执行时间", "旧名", "新名"])
        for row in plan:
            if row["状态"] != ST_OK:
                continue
            (WORK_DIR / row["旧名"]).rename(WORK_DIR / row["新名"])
            stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            writer.writerow([stamp, row["旧名"], row["新名"]])
            done += 1
    return done


def verify(plan):
    checks = []

    def check(name, ok):
        checks.append((name, ok))

    by_status = {}
    for row in plan:
        by_status[row["状态"]] = by_status.get(row["状态"], 0) + 1
    check("预览表覆盖全部 6 个文件", len(plan) == 6)
    check("改名 3 个", by_status.get(ST_OK, 0) == 3)
    check("已规范跳过 1 个", by_status.get(ST_STD, 0) == 1)
    check("未识别挂起 1 个", by_status.get(ST_UNKNOWN, 0) == 1)
    check("重名冲突挂起 1 个", by_status.get(ST_DUP, 0) == 1)

    names = {p.name for p in WORK_DIR.iterdir() if p.is_file()}
    check("规范名文件 4 个（含原有 1 个）", sum(1 for n in names if STD_PATTERN.match(n)) == 4)
    check("挂起文件原地未动 2 个",
          "结果最终版改2.xlsx" in names and "血常规 2026.3.6 S002 (2).csv" in names)
    check("三个新名都已落位",
          all(n in names for n in ("20260305_S001_血常规.csv",
                                   "20260306_S002_血常规.csv",
                                   "20260309_S004_尿常规.csv")))

    with open(RENAME_LOG, encoding="utf-8-sig") as f:
        log_rows = list(csv.reader(f))
    check("rename_log 记了 3 笔", len(log_rows) - 1 == 3)

    ok_names = [r["新名"] for r in plan if r["状态"] == ST_OK]
    check("新名互不重复", len(set(ok_names)) == len(ok_names))

    for name, ok in checks:
        print(("OK   " if ok else "FAIL ") + name)
    return all(ok for _, ok in checks)


def main():
    started = time.time()
    write_log(f"运行主机：{host_label()}")
    write_log(f"开始执行：{datetime.now():%Y-%m-%d %H:%M:%S}  "
              f"Python {sys.version.split()[0]}（{Path(sys.executable).name}）")

    plan = plan_renames()
    write_preview(plan)
    done = apply_plan(plan)

    def count(status):
        return sum(1 for r in plan if r["状态"] == status)

    print(f"扫描 {len(plan)} 个文件：改名 {done}、已规范跳过 {count(ST_STD)}、"
          f"未识别挂起 {count(ST_UNKNOWN)}、重名冲突 {count(ST_DUP)}")
    print(f"预览表：{PREVIEW_FILE}")
    print(f"改名日志：{RENAME_LOG}")
    print(f"文件目录：{WORK_DIR}")

    passed = verify(plan)
    cost = time.time() - started
    print(f"判定：{'PASS' if passed else 'FAIL'}（用时 {cost:.1f}s）")
    write_log(f"执行完成：{datetime.now():%Y-%m-%d %H:%M:%S}  用时 {cost:.1f}s  "
              f"判定 {'PASS' if passed else 'FAIL'}")

    if "--no-pause" not in sys.argv:
        input("按回车结束……")


if __name__ == "__main__":
    main()
