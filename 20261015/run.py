#!/usr/bin/env python3
"""实验室日报周报自动生成：三张台账 → 日报一份 + 周报一份。

跑法：
    python run.py                # Windows 上双击 run.py 也行
    python run.py --no-pause     # 跑完不暂停（脚本调用时用）

输入：01_raw_data/ 下的三张台账（样品接收、检测完成、异常）
输出：02_output/实验室日报.xlsx、02_output/实验室周报.xlsx
"""
import csv
import socket
import subprocess
import sys
import time
from datetime import date, datetime, timedelta
from pathlib import Path

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font
except ImportError:                     # 缺依赖自动装，走国内源
    subprocess.check_call([sys.executable, "-m", "pip", "install",
                           "-i", "https://pypi.tuna.tsinghua.edu.cn/simple", "openpyxl"])
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font

# ---------- 顶部的路径与字段常量区 ----------
BASE = Path(__file__).resolve().parent
RAW_DIR = BASE / "01_raw_data"
OUT_DIR = BASE / "02_output"
SRC_DIR = BASE / "source"
LOG_FILE = SRC_DIR / "run_log.txt"

RECV_FILE = RAW_DIR / "样品接收台账.csv"
DONE_FILE = RAW_DIR / "检测完成台账.csv"
ABN_FILE = RAW_DIR / "异常台账.csv"
DAILY_NAME = "实验室日报.xlsx"
WEEKLY_NAME = "实验室周报.xlsx"

# 列名按导出文件里的表头取，不按第几列——模板一加列，按序号读就整体错位
RECV_DATE, RECV_ORDER = "接收日期", "委托单号"
RECV_SAMPLE, RECV_QTY = "样品编号", "样品数"
DONE_DATE, DONE_SAMPLE = "完成日期", "样品编号"
DONE_ITEM, DONE_CONC = "检测项", "结论"
TAT_COL = "TAT承诺天数"
ABN_DATE, ABN_SAMPLE, ABN_TYPE, ABN_STATE = "发生日期", "样品编号", "异常类型", "处理状态"
EMPTY = ""                              # 数字列没填时写空，不写 0

WEEKDAYS = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]


def read_ledger(path):
    """读一张台账。用 utf-8-sig——导出文件的 BOM 不剥掉，第一个列名会带上隐藏字符。"""
    with open(path, encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def to_date(v):
    """把台账里的日期写法归一到 date，认不出来返回 None。

    同一个 10 月 12 日，三张导出文件里能写出四种字符串：`2026/10/12`、
    `2026-10-12 09:20:31`、`2026年10月12日`、`26.10.13`。不归一直按字符串
    分组，同一天会被拆成四天，日报看着就是断档的。
    """
    s = str(v or "").strip()
    if not s:
        return None
    s = s.split(" ")[0].split("T")[0]                 # 带时间戳的先切掉时间
    s = s.replace("年", "-").replace("月", "-").replace("日", "")
    s = s.replace("/", "-").replace(".", "-")
    for fmt in ("%Y-%m-%d", "%y-%m-%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def to_int(v):
    """数量列：空、空格、"—" 都不是 0，是没填。区分不开，均值一算就被拉低。"""
    s = str(v or "").strip()
    if not s:
        return None
    try:
        return int(float(s))
    except ValueError:
        return None


def tat_days(recv, sample, done_day, promise):
    """实际周转天数 = 完成日 - 接收日。接收台账里没这个样品、承诺天数没填，返回 None。

    返回 None 时不许当 0 天算进平均——那会让「平均周转」好看得离谱，
    真出问题反而查不出来。
    """
    if sample not in recv or promise is None:
        return None
    return (done_day - recv[sample]["day"]).days


def load_recv():
    """接收台账 → {样品编号: {day, order, qty}}，另带两份待核清单。

    同一样品登记多次只留最早那次：补录、复检会把同一样品再写一遍，
    按行数算接收量会虚高，取最后一次还会把 TAT 算长（把补录那天当成收货日）。
    """
    out, dups, dropped = {}, [], []
    for r in read_ledger(RECV_FILE):
        sample = str(r.get(RECV_SAMPLE) or "").strip()
        day = to_date(r.get(RECV_DATE))
        if day is None:
            dropped.append(("接收台账", sample, str(r.get(RECV_DATE) or "").strip()))
            continue
        if sample in out:
            dups.append(("接收台账", sample,
                         f"{r.get(RECV_ORDER, '')} 与先登记的那条重复，只算一次"))
            out[sample]["day"] = min(out[sample]["day"], day)
            continue
        out[sample] = {"day": day, "order": str(r.get(RECV_ORDER) or "").strip(),
                       "qty": to_int(r.get(RECV_QTY))}
    return out, dups, dropped


def load_done(recv):
    """完成台账 → 已完成的项 + 在办项 + 两份待核清单。

    完成日期为空不是「当天没干活」，是「还没完成」：丢掉它在办一条不剩，
    当成 0 又会让完成数看着没变。所以空值和认不出的值分两堆，别混。
    """
    rows, doing, dups, dropped, seen = [], [], [], [], set()
    for r in read_ledger(DONE_FILE):
        sample = str(r.get(DONE_SAMPLE) or "").strip()
        item = str(r.get(DONE_ITEM) or "").strip()
        raw = str(r.get(DONE_DATE) or "").strip()
        promise = to_int(r.get(TAT_COL))
        day = to_date(raw)
        if day is None:
            if raw:
                dropped.append(("完成台账", sample, raw))
            else:
                doing.append({"sample": sample, "item": item})
            continue
        if (sample, item, day) in seen:       # 同一天同一样品同一检项写两遍：只算一次
            dups.append(("完成台账", sample, f"{item} 在 {day} 重复登记，只算一次"))
            continue
        seen.add((sample, item, day))
        tat = tat_days(recv, sample, day, promise)
        rows.append({"day": day, "sample": sample, "item": item,
                     "concl": str(r.get(DONE_CONC) or "").strip(),
                     "tat": tat, "promise": promise,
                     "over": tat is not None and promise is not None and tat > promise})
    return rows, doing, dups, dropped


def last_done(done):
    """样品 → 最后一次完成日（同一样品多个检项，取最晚那天）。"""
    out = {}
    for r in done:
        out[r["sample"]] = max(out.get(r["sample"], r["day"]), r["day"])
    return out


def backlog(recv, last, day):
    """截至 day 当天的在办样品数：已接收，但完成日晚于 day 或还没完成的。"""
    return sum(1 for s, info in recv.items()
               if info["day"] <= day and (s not in last or last[s] > day))


def day_list(d1, d2):
    """连续的日期列表：中间没记录的那天也要留一行。

    跳过空档日，日报看着天天满勤；补 0 又等于断言「那天什么都没发生」。
    留空才对——空着才看得出是没记录，不是没干活。
    """
    return [d1 + timedelta(days=i) for i in range((d2 - d1).days + 1)]


def daily_rows(days, recv, done, abn, last):
    """按天出一行汇总。空档日也出一行，指标留空、不补 0。"""
    rows = []
    for day in days:
        r_day = [i for i in recv.values() if i["day"] == day]
        d_day = [r for r in done if r["day"] == day]
        a_day = [r for r in abn if r["day"] == day]
        qty = [i["qty"] for i in r_day if i["qty"] is not None]
        note = []
        if not (r_day or d_day or a_day):
            note.append("三张台账当天都没有记录")
        if any(i["qty"] is None for i in r_day):
            note.append("有样品数未填")
        rows.append({
            "day": day,
            "recv_n": len(r_day) or None,
            "qty": sum(qty) if qty else None,
            "done_n": len(d_day) or None,
            "over_n": sum(1 for r in d_day if r["over"]) or None,
            "abn_n": len(a_day) or None,
            "backlog": backlog(recv, last, day),
            "note": "；".join(note),
        })
    return rows


def week_key(day):
    """ISO 周界：周一是一周第一天。返回 (年, 周号)。

    别用「最近 7 天」代替——那个窗口跨两个自然周，跟「上周」比的时候
    起点还每天往后挪，比出来的差值没有意义。
    """
    return day.isocalendar()[:2]


def period_sum(recv, done, abn, d1, d2):
    """一段日期内的四个数。对比用的都是这一段，不是「整周」。"""
    d_rows = [r for r in done if d1 <= r["day"] <= d2]
    return {
        "接收样品": sum(1 for i in recv.values() if d1 <= i["day"] <= d2),
        "完成检项": len(d_rows),
        "超 TAT": sum(1 for r in d_rows if r["over"]),
        "异常": sum(1 for r in abn if d1 <= r["day"] <= d2),
    }


def same_period(days):
    """本周区间与上周同期区间。

    周中出周报，本周才走了 4 天；拿 4 天去对上一整周 7 天，年年都是
    「断崖式下滑」——那是口径的问题，不是业绩的问题。所以上周切到同一星期几。
    """
    this_start = next(d for d in days if week_key(d) == week_key(days[-1]))
    this_end = days[-1]
    last_start = this_start - timedelta(days=7)
    return this_start, this_end, last_start, last_start + (this_end - this_start)


def week_range(key):
    """ISO 周的完整区间：周一到周日。区间写整周，别写「这一周里出现过的日子」——
    后者看着像这周只有 5 天，读者会以为漏了周末。"""
    start = date.fromisocalendar(key[0], key[1], 1)
    return start, start + timedelta(days=6)


def weekly_rows(recv, done, abn):
    """按周归总，同时给出本周与上周同期的对比。"""
    buckets = {}
    for info in recv.values():
        buckets.setdefault(week_key(info["day"]), {"recv": 0, "done": 0, "over": 0, "abn": 0})
        buckets[week_key(info["day"])]["recv"] += 1
    for r in done:
        b = buckets.setdefault(week_key(r["day"]), {"recv": 0, "done": 0, "over": 0, "abn": 0})
        b["done"] += 1
        b["over"] += 1 if r["over"] else 0
    for r in abn:
        b = buckets.setdefault(week_key(r["day"]), {"recv": 0, "done": 0, "over": 0, "abn": 0})
        b["abn"] += 1

    weeks = []
    for key, b in sorted(buckets.items()):
        start, end = week_range(key)
        weeks.append([f"{key[0]}-W{key[1]:02d}", f"{start} ~ {end}",
                      b["recv"], b["done"], b["over"], b["abn"]])

    days = day_list(min(i["day"] for i in recv.values()), max(i["day"] for i in recv.values()))
    d1, d2, l1, l2 = same_period(days)
    this_p = period_sum(recv, done, abn, d1, d2)
    last_p = period_sum(recv, done, abn, l1, l2)
    compare = [[k, this_p[k], last_p[k], this_p[k] - last_p[k]] for k in this_p]
    return weeks, compare, (d1, d2, l1, l2)


def detail_rows(days, recv, done, abn):
    """逐日明细长表：汇总对不上时，得能一眼回到那几行，不能只交一个数字。"""
    out = []
    for day in days:
        for s, info in sorted(recv.items()):
            if info["day"] == day:
                qty = EMPTY if info["qty"] is None else info["qty"]
                out.append([str(day), "接收", s, info["order"], f"{qty} 件", EMPTY])
        for r in done:
            if r["day"] == day:
                out.append([str(day), "完成", r["sample"], r["item"], r["concl"],
                            "超 TAT" if r["over"] else EMPTY])
        for r in abn:
            if r["day"] == day:
                out.append([str(day), "异常", r["sample"], r["type"], r["state"], EMPTY])
    return out


def gap_rows(daily, dups, dropped, doing):
    """该有人核一遍的，全汇成一张清单：空档日、认不出的日期、重复登记、在办项。"""
    out = []
    for row in daily:
        if not row["note"]:
            continue
        kind = "空档日" if "都没有记录" in row["note"] else "待核"
        out.append([kind, row["day"], row["note"], EMPTY])
    for src, key, raw in dropped:
        out.append(["日期认不出", None, f"{src} 里写的「{raw}」，没归进任何一天", key])
    for src, key, note in dups:
        out.append(["重复登记", None, f"{src}：{key} {note}", key])
    for r in doing:
        out.append(["在办", None, f"{r['sample']} 的 {r['item']} 还没完成", r["sample"]])
    return out


def _sheet(ws, header, rows, widths):
    ws.append(header)
    for c in ws[1]:
        c.font = Font(bold=True)
    for r in rows:
        ws.append([EMPTY if v is None else v for v in r])
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w
    ws.freeze_panes = "A2"


def write_daily(daily, detail):
    wb = Workbook()
    ws = wb.active
    ws.title = "日报总览"
    _sheet(ws, ["日期", "星期", "接收样品", "接收件数", "完成检项", "超 TAT", "异常", "在办", "备注"],
           [[r["day"], WEEKDAYS[r["day"].weekday()], r["recv_n"], r["qty"], r["done_n"],
             r["over_n"], r["abn_n"], r["backlog"], r["note"]] for r in daily],
           [12, 7, 10, 10, 10, 8, 7, 7, 34])
    _sheet(wb.create_sheet("逐日明细"), ["日期", "类型", "样品编号", "单号/检项", "结论/状态", "备注"],
           detail, [12, 7, 12, 16, 12, 12])
    wb.save(OUT_DIR / DAILY_NAME)


def write_weekly(weeks, compare, span, gaps):
    wb = Workbook()
    ws = wb.active
    ws.title = "按周汇总"
    _sheet(ws, ["周次", "周区间", "接收样品", "完成检项", "超 TAT", "异常"], weeks, [12, 26, 10, 10, 8, 7])
    d1, d2, l1, l2 = span
    _sheet(wb.create_sheet("本周对上周同期"),
           ["指标", f"本周 {d1} ~ {d2}", f"上周同期 {l1} ~ {l2}", "差异"], compare, [14, 22, 24, 10])
    _sheet(wb.create_sheet("缺报与待核清单"), ["类型", "日期", "说明", "关联单号/样品"],
           gaps, [12, 12, 52, 14])
    wb.save(OUT_DIR / WEEKLY_NAME)


def verify(daily, weeks, compare, gaps, span, detail):
    """生成之后重新打开两份成品回读——内存里对不算对。"""
    oks = []
    wb = load_workbook(OUT_DIR / DAILY_NAME)
    ws = wb["日报总览"]
    rows = {str(ws.cell(row=i + 2, column=1).value)[:10]:
            [ws.cell(row=i + 2, column=c).value for c in range(1, 10)]
            for i in range(ws.max_row - 1)}
    oks.append(("日报行数 = 连续天数（含空档日）", len(rows) == len(daily) == 11))
    oks.append(("周末空档日留行、指标留空不补 0",
                rows["2026-10-10"][2] in (None, EMPTY)
                and rows["2026-10-10"][8] == "三张台账当天都没有记录"))
    oks.append(("带时间戳的日期归到同一天（10-12 接收 2 个样品）", rows["2026-10-12"][2] == 2))
    oks.append(("重复登记的样品只算一次（10-12 共 24 件）", rows["2026-10-12"][3] == 24))
    oks.append(("完成日期为空算在办、不算当天完成（10-15 在办 3 个）", rows["2026-10-15"][7] == 3))
    oks.append(("超 TAT 共 4 条（按接收日+承诺天数判）", sum(r["over_n"] or 0 for r in daily) == 4))
    oks.append(("逐日明细行数 = 明细数据行数", wb["逐日明细"].max_row - 1 == len(detail)))

    wb2 = load_workbook(OUT_DIR / WEEKLY_NAME)
    d1, d2, l1, l2 = span
    oks.append(("周报按 ISO 周出两行（W41 / W42）", wb2["按周汇总"].max_row - 1 == len(weeks) == 2))
    oks.append(("本周区间从周一开始", d1.weekday() == 0))
    oks.append(("上周同期与本周等长（不是整周对 4 天）", (d2 - d1).days == (l2 - l1).days))
    oks.append(("对比表列头写出区间",
                str(wb2["本周对上周同期"].cell(row=1, column=2).value) == f"本周 {d1} ~ {d2}"))
    oks.append(("缺报与待核清单条数一致", wb2["缺报与待核清单"].max_row - 1 == len(gaps)))
    return oks


def host_ip():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("8.8.8.8", 80))                    # 不真发包，只为拿到本机对外网卡的地址
        return s.getsockname()[0]
    except OSError:
        try:
            return socket.gethostbyname(socket.gethostname())
        except OSError:
            return "未知"
    finally:
        s.close()


def log(line):
    with open(LOG_FILE, "a", encoding="utf-8") as f:   # 追加不覆盖，每次跑留一段
        f.write(line + "\n")


def main():
    t0 = time.time()
    log(f"运行主机：{socket.gethostname()}（{host_ip()}）")
    log(f"开始执行：{datetime.now():%Y-%m-%d %H:%M:%S}  Python {sys.version.split()[0]} / "
        f"{Path(sys.executable).name}")
    OUT_DIR.mkdir(exist_ok=True)

    recv, dups_r, drop_r = load_recv()
    done, doing, dups_d, drop_d = load_done(recv)
    abn, drop_a = [], []
    for r in read_ledger(ABN_FILE):
        day = to_date(r.get(ABN_DATE))
        sample = str(r.get(ABN_SAMPLE) or "").strip()
        if day is None:                       # 异常台账里日期认不出的行也不许悄悄丢
            drop_a.append(("异常台账", sample, str(r.get(ABN_DATE) or "").strip()))
            continue
        abn.append({"day": day, "sample": sample,
                    "type": str(r.get(ABN_TYPE) or "").strip(),
                    "state": str(r.get(ABN_STATE) or "").strip()})
    last = last_done(done)
    days = day_list(min(i["day"] for i in recv.values()), max(i["day"] for i in recv.values()))
    daily = daily_rows(days, recv, done, abn, last)
    weeks, compare, span = weekly_rows(recv, done, abn)
    gaps = gap_rows(daily, dups_r + dups_d, drop_r + drop_d + drop_a, doing)
    detail = detail_rows(days, recv, done, abn)
    write_daily(daily, detail)
    write_weekly(weeks, compare, span, gaps)

    oks = verify(daily, weeks, compare, gaps, span, detail)
    for name, ok in oks:
        print(("OK   " if ok else "FAIL ") + name)
    all_ok = all(ok for _, ok in oks)
    print(f"\n接收样品 {len(recv)} 个 / 完成检项 {len(done)} 条 / 在办 {len(doing)} 条 / "
          f"待核 {len(gaps)} 条")
    print("总判定：" + ("PASS" if all_ok else "FAIL"))
    log(f"执行完成：{datetime.now():%Y-%m-%d %H:%M:%S}  用时 {time.time() - t0:.1f}s  "
        f"判定 {'PASS' if all_ok else 'FAIL'}")

    if "--no-pause" not in sys.argv and sys.stdin and sys.stdin.isatty():
        input("按回车键结束...")


if __name__ == "__main__":
    main()
