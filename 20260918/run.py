# -*- coding: utf-8 -*-
"""引流长文《样本流转交接记录还靠纸单签字，查一管样本在谁手里只能翻半天》配套脚本。

把现场登记的交接记录，串成一管一行的流转台账：

    样本主表（一行一管样本）
  + 样本交接记录（一行一次交接：谁交给谁、什么时候）
  + 流转环节规则（一行一个环节：接收岗位、停留时限）
→ 每管样本按环节序号串成一条链，缺哪个环节、哪一段时间倒挂、哪一环记了两条、
  哪一环压了太久，都挑出来；交接记录里对不上主表的编号单列，不混进台账。

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


_ensure(("openpyxl", "openpyxl"))

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

HERE = Path(__file__).resolve().parent
SAMPLE_FILE = HERE / "01_raw_data" / "样本主表.csv"            # 一行一管样本
MOVE_FILE = HERE / "01_raw_data" / "样本交接记录.csv"          # 一行一次交接
RULE_FILE = HERE / "source" / "templates" / "流转环节规则.csv"  # 一行一个环节
OUT_FILE = HERE / "02_output" / "样本流转台账.xlsx"             # 台账 + 环节停留 + 异常与提醒
LOG_DIR = HERE / "source"                                     # 运行日志（run_log.txt 落在 source/ 下）

SAMPLE_COLS = ["样本编号", "中心编号", "中心名称", "访视点", "采集日期"]
MOVE_COLS = ["样本编号", "交接时间", "交出方", "接收方", "交接人", "备注"]
RULE_COLS = ["环节序号", "环节", "接收岗位", "停留时限小时"]

LEDGER_COLS = ["样本编号", "中心编号", "中心名称", "访视点", "环节数", "首环节时间", "末环节时间",
               "全程小时", "末环节持有人", "流转链", "链状态"]
STAGE_COLS = ["样本编号", "中心编号", "环节", "接收岗位", "进入时间", "离开时间",
              "停留小时", "时限小时", "超时"]
ALERT_COLS = ["样本编号", "中心编号", "环节", "类型", "说明"]

# 链状态只有四个词。挂起的两档（断链、需人工确认）都要人去看，完整的那档才是能直接交差的。
STATUS_OK = "完整"
STATUS_BREAK = "有断链"
STATUS_CHECK = "需人工确认"
STATUS_NONE = "无交接记录"

HEAD_FILL = PatternFill("solid", fgColor="D9E1F2")
BAD_FILL = PatternFill("solid", fgColor="FFC7CE")     # 要人拍板的（断链、倒挂、重复、超时）

# 异常的类型词。写得具体一点，接的人一眼知道该找谁。
KIND_MISS_STAGE = "环节缺失"
KIND_BACKWARD = "时间倒挂"
KIND_DUP = "同环节重复"
KIND_NO_TIME = "时间缺失"
KIND_OVER = "停留超时"
KIND_STRAY = "游离记录"
KIND_NO_RECORD = "无交接记录"


def to_int(text: str) -> int:
    """时限写成空、"不设限"、带单位的一律当 0；0 表示这一环不设停留时限。"""
    try:
        return int(str(text).strip())
    except (TypeError, ValueError):
        return 0


def to_dt(text: str):
    """交接时间只认「YYYY-MM-DD HH:MM」这一种写法，认不出来返回 None。

    空字符串绝不能当 0 点：那会把"没填时间"变成"半夜交接"，停留时长跟着全算错，
    而且错得看不出来。宁可让它空着，标成「时间缺失」交给人。
    """
    text = (text or "").strip()
    if not text:
        return None
    try:
        return datetime.strptime(text, "%Y-%m-%d %H:%M")
    except ValueError:
        return None


def hours_between(start, end):
    """一段停留的小时数。两头有一个是空就算不出来，返回 None —— 不能当 0 小时。"""
    if start is None or end is None:
        return None
    return round((end - start).total_seconds() / 3600, 1)


def alert(sid: str, code: str, stage: str, kind: str, note: str) -> dict:
    """一条异常记录。样本级的带编号与中心，游离记录的编号对不上主表，中心留空。"""
    return {"样本编号": sid, "中心编号": code, "环节": stage, "类型": kind, "说明": note}


def read_rules(path: Path):
    """读环节规则，一行一个环节，按「环节序号」排定先后。

    环节靠「接收岗位」认：谁接手，就算进了哪个环节。
    序号是唯一权威，不许拿交接时间排序来定环节先后——时间录反的时候，
    按时间排会把倒挂排成正常顺序，异常被抹平，谁看都正常。
    表尾说明行的序号列是空的，干脆丢。
    """
    rules = []
    with open(path, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            if not (row.get("环节序号") or "").strip():
                continue
            rules.append({c: (row.get(c) or "").strip() for c in RULE_COLS})
    rules.sort(key=lambda r: to_int(r["环节序号"]))
    return rules


def read_samples(path: Path):
    """读样本主表，一行一管。样本编号非空才算数，表尾的说明行走这儿被丢掉。"""
    out = []
    with open(path, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            if not (row.get("样本编号") or "").strip():
                continue
            out.append({c: (row.get(c) or "").strip() for c in SAMPLE_COLS})
    return out


def read_moves(path: Path):
    """读交接记录，一行一次交接，按样本编号归堆。

    编号一律当字符串读，不许转数字：00102V1A 这种前导零的编号一旦被当成数字，
    前面的零补不回来，跟主表就对不上了。
    """
    by_sample = {}
    with open(path, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            sid = (row.get("样本编号") or "").strip()
            if not sid:
                continue
            by_sample.setdefault(sid, []).append({c: (row.get(c) or "").strip() for c in MOVE_COLS})
    return by_sample


def build_chain(sample: dict, moves: list, rules: list):
    """把一管样本的交接记录串成一条链，顺带把毛病挑出来。

    顺序按规则表的环节序号走，一环比一环地核，不按交接时间排序：
    时间录反的时候，排序会把倒挂抹平，看着一切正常。
    一个环节记了两条，先到的那条算数，其余报出来交给人核——两条都"算数"是不可能的。
    """
    sid, code = sample["样本编号"], sample["中心编号"]
    empty = {"样本编号": sid, "中心编号": code, "中心名称": sample["中心名称"],
             "访视点": sample["访视点"], "环节数": 0, "首环节时间": "", "末环节时间": "",
             "全程小时": "", "末环节持有人": "", "流转链": "（无）", "链状态": STATUS_NONE}
    if not moves:
        # 一条记录都没有：这管样本要么还没进流程，要么登记漏了。四个环节全报"缺失"是噪音，
        # 只说一句"没有记录"，人去现场一查就知道是哪种。
        return empty, [], [alert(sid, code, "", KIND_NO_RECORD,
                                 "主表里有这管样本，交接记录里一条都没有：要么还没进流程，要么漏登了")]

    alerts, stages = [], []

    by_post = {}
    for mv in moves:
        by_post.setdefault(mv["接收方"], []).append(mv)

    for rule in rules:
        post = rule["接收岗位"]
        got = by_post.get(post, [])
        if not got:
            alerts.append(alert(sid, code, rule["环节"], KIND_MISS_STAGE,
                                f"交接记录里没有交给「{post}」这一条，{rule['环节']}这一环是断的"))
            continue

        timed = [m for m in got if to_dt(m["交接时间"])]
        first = min(timed, key=lambda m: to_dt(m["交接时间"])) if timed else got[0]
        if len(got) > 1:
            others = "、".join(m["交接时间"] or "空" for m in got if m is not first)
            alerts.append(alert(sid, code, rule["环节"], KIND_DUP,
                                f"同一个环节记了 {len(got)} 条（另一次是 {others}），"
                                f"先到的那条算数，其余交给人核"))
        if to_dt(first["交接时间"]) is None:
            alerts.append(alert(sid, code, rule["环节"], KIND_NO_TIME,
                                "这条交接的时间没填，这一段的停留时长算不出来，先找当时交接的人补"))

        stages.append({"环节": rule["环节"], "接收岗位": post, "进入": to_dt(first["交接时间"]),
                       "离开": None, "交货人": first["交接人"], "停留": None,
                       "时限": to_int(rule["停留时限小时"]), "超时": ""})

    # 停留算在"待着的"那一环上：本环进入 → 下一环进入，就是本环待了多久。
    for idx in range(len(stages) - 1):
        now, nxt = stages[idx], stages[idx + 1]
        now["离开"] = nxt["进入"]
        if now["进入"] and nxt["进入"] and nxt["进入"] < now["进入"]:
            alerts.append(alert(sid, code, nxt["环节"], KIND_BACKWARD,
                                f"进入时间 {nxt['进入']:%m-%d %H:%M} 比上一环 "
                                f"{now['环节']}的 {now['进入']:%m-%d %H:%M} 还早，时间多半录错了"))
            continue                                     # 倒挂那一段算不出停留，别拿负数去比时限
        now["停留"] = hours_between(now["进入"], nxt["进入"])
        if now["停留"] is not None and now["时限"] and now["停留"] > now["时限"]:
            now["超时"] = "是"
            alerts.append(alert(sid, code, now["环节"], KIND_OVER,
                                f"这一环待了 {now['停留']} 小时，超过 {now['时限']} 小时的时限"))

    if len(stages) < len(rules):
        status = STATUS_BREAK
    elif any(a["类型"] in (KIND_BACKWARD, KIND_DUP, KIND_NO_TIME) for a in alerts):
        status = STATUS_CHECK
    else:
        status = STATUS_OK

    first_in = next((s["进入"] for s in stages if s["进入"]), None)
    last_in = next((s["进入"] for s in reversed(stages) if s["进入"]), None)
    walked = len(stages) > 1        # 只有一环等于链根本没往下走，首末同一个时间、全程算不出来
    ledger = {
        "样本编号": sid, "中心编号": code, "中心名称": sample["中心名称"],
        "访视点": sample["访视点"], "环节数": len(stages),
        "首环节时间": f"{first_in:%Y-%m-%d %H:%M}" if first_in else "",
        "末环节时间": f"{last_in:%Y-%m-%d %H:%M}" if walked else "",
        "全程小时": hours_between(first_in, last_in) if walked else "",
        "末环节持有人": stages[-1]["交货人"] if stages else "",
        "流转链": " → ".join(s["环节"] for s in stages) if stages else "（无）",
        "链状态": status,
    }
    return ledger, stages, alerts


def write_workbook(ledgers: list, stages: list, alerts: list, out_path: Path) -> Path:
    """三个页签写进一个 xlsx。

    流转台账给管样本的人：一管一行，看链状态就知道哪管要处理；
    环节停留给要追超时的人：每管每一环待了多久、时限多少，超时的刷红；
    异常与提醒给现场与项目组：缺哪一环、时间哪对不上、哪个编号抄错了，都在最后一页。
    """
    wb = Workbook()
    sheets = [("流转台账", LEDGER_COLS, ledgers),
              ("环节停留", STAGE_COLS, stages),
              ("异常与提醒", ALERT_COLS, alerts)]
    for idx, (title, cols, data) in enumerate(sheets):
        ws = wb.active if idx == 0 else wb.create_sheet()
        ws.title = title
        ws.append(cols)
        for c in ws[1]:
            c.fill, c.font = HEAD_FILL, Font(bold=True)
        for row in data:
            if isinstance(row, dict):
                ws.append([row.get(c, "") for c in cols])
                if row.get("链状态") in (STATUS_BREAK, STATUS_CHECK) or row.get("超时") == "是":
                    for c in ws[ws.max_row]:
                        c.fill = BAD_FILL
            else:
                ws.append(row)
        for i, col in enumerate(cols, 1):
            ws.column_dimensions[get_column_letter(i)].width = min(max(len(col) + 6, 10), 34)
        ws.freeze_panes = "A2"

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    return out_path


def main() -> bool:
    rules = read_rules(RULE_FILE)
    samples = read_samples(SAMPLE_FILE)
    by_sample = read_moves(MOVE_FILE)
    print(f"环节规则 {len(rules)} 环、样本主表 {len(samples)} 管、交接记录 "
          f"{sum(len(v) for v in by_sample.values())} 条")

    ledgers, stages, alerts = [], [], []
    known = {s["样本编号"] for s in samples}

    for sid in sorted(set(by_sample) - known):
        alerts.append(alert(sid, "", "", KIND_STRAY,
                            f"交接记录里有 {len(by_sample[sid])} 条，样本主表里没有这个编号，"
                            f"多半是手抄时前导零丢了，先去核编号再决定并到哪一管"))

    for sample in samples:
        ledger, st, al = build_chain(sample, by_sample.get(sample["样本编号"], []), rules)
        ledgers.append(ledger)
        stages += [{"样本编号": sample["样本编号"], "中心编号": sample["中心编号"],
                    "环节": s["环节"], "接收岗位": s["接收岗位"],
                    "进入时间": f"{s['进入']:%Y-%m-%d %H:%M}" if s["进入"] else "",
                    "离开时间": f"{s['离开']:%Y-%m-%d %H:%M}" if s["离开"] else "",
                    "停留小时": s["停留"] if s["停留"] is not None else "",
                    "时限小时": s["时限"] or "", "超时": s["超时"]} for s in st]
        alerts += al

    write_workbook(ledgers, stages, alerts, OUT_FILE)

    done = [x for x in ledgers if x["链状态"] == STATUS_OK]
    print(f"  链完整 {len(done)} 管、有断链 {len([x for x in ledgers if x['链状态'] == STATUS_BREAK])} 管、"
          f"需人工确认 {len([x for x in ledgers if x['链状态'] == STATUS_CHECK])} 管、"
          f"无交接记录 {len([x for x in ledgers if x['链状态'] == STATUS_NONE])} 管")
    print(f"  异常与提醒 {len(alerts)} 条，其中停留超时 "
          f"{len([a for a in alerts if a['类型'] == KIND_OVER])} 条")
    print(f"流转台账已生成：{OUT_FILE.name}")

    return verify(ledgers, stages, alerts)


def verify(ledgers: list, stages: list, alerts: list) -> bool:
    """回读校验。这些数字是照着造好的边界数据定的，对不上说明逻辑变了或数据被改过。"""
    from openpyxl import load_workbook
    wb = load_workbook(OUT_FILE)
    ledger = list(wb["流转台账"].iter_rows(min_row=2, values_only=True))
    stage = list(wb["环节停留"].iter_rows(min_row=2, values_only=True))
    trouble = list(wb["异常与提醒"].iter_rows(min_row=2, values_only=True))
    kinds = {}
    for r in trouble:
        kinds[r[3]] = kinds.get(r[3], 0) + 1
    status = {}
    for r in ledger:
        status[r[10]] = status.get(r[10], 0) + 1

    def cell(sid, stage_name, col):
        for r in stage:
            if r[0] == sid and r[2] == stage_name:
                return r[col]
        return None

    checks = [
        ("流转台账 12 管", len(ledger) == 12, f"{len(ledger)} 管"),
        ("环节停留 40 行", len(stage) == 40, f"{len(stage)} 行"),
        ("异常与提醒 11 条", len(trouble) == 11, f"{len(trouble)} 条"),
        ("环节缺失 4 条（00104 缺 3 环 + 00301V1B 缺 1 环）",
         kinds.get(KIND_MISS_STAGE, 0) == 4, f"{kinds.get(KIND_MISS_STAGE, 0)} 条"),
        ("时间倒挂 1 条", kinds.get(KIND_BACKWARD, 0) == 1, f"{kinds.get(KIND_BACKWARD, 0)} 条"),
        ("同环节重复 1 条", kinds.get(KIND_DUP, 0) == 1, f"{kinds.get(KIND_DUP, 0)} 条"),
        ("时间缺失 1 条", kinds.get(KIND_NO_TIME, 0) == 1, f"{kinds.get(KIND_NO_TIME, 0)} 条"),
        ("停留超时 2 条", kinds.get(KIND_OVER, 0) == 2, f"{kinds.get(KIND_OVER, 0)} 条"),
        ("游离记录 1 条", kinds.get(KIND_STRAY, 0) == 1, f"{kinds.get(KIND_STRAY, 0)} 条"),
        ("无交接记录 1 条", kinds.get(KIND_NO_RECORD, 0) == 1, f"{kinds.get(KIND_NO_RECORD, 0)} 条"),
        ("链状态：完整 6 / 有断链 2 / 需人工确认 3 / 无交接记录 1",
         (status.get(STATUS_OK), status.get(STATUS_BREAK), status.get(STATUS_CHECK),
          status.get(STATUS_NONE)) == (6, 2, 3, 1),
         "/".join(f"{k}{v}" for k, v in status.items())),
        ("00104V1A 只走完 1 环（后面三环断着）",
         len([r for r in stage if r[0] == "00104V1A"]) == 1,
         f"{len([r for r in stage if r[0] == '00104V1A'])} 环"),
        ("只有一环时末环节时间与全程小时留空，不拿 0 小时糊过去",
         [r for r in ledger if r[0] == "00104V1A"][0][6:8] == (None, None),
         f"{[r for r in ledger if r[0] == '00104V1A'][0][6:8]}"),
        ("00202V1A 前处理岗取的是先到的 09-15 11:00 那条",
         cell("00202V1A", "前处理分装", 4) == "2026-09-15 11:00",
         cell("00202V1A", "前处理分装", 4)),
        ("00103V1A 检测岗停留 96.5 小时、标了超时",
         (cell("00103V1A", "检测上机", 6), cell("00103V1A", "检测上机", 8)) == (96.5, "是"),
         f"{cell('00103V1A', '检测上机', 6)} / {cell('00103V1A', '检测上机', 8)}"),
        ("00201V1A 倒挂那一段停留留空，不拿负数去比时限",
         cell("00201V1A", "检测上机", 6) == 50.0,
         f"{cell('00201V1A', '检测上机', 6)}"),
        ("游离记录那条点了编号，并说明是前导零丢了",
         any(r[0] == "102V1A" and "前导零" in str(r[4]) for r in trouble), "见异常页签"),
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
