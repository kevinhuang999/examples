# -*- coding: utf-8 -*-
"""引流长文《中心实验室样本接收登记表，一批几百管靠手工核到半夜》配套脚本。

把扫码枪导出的接收明细 + 温控记录仪导出的运输温度拼成一份样本接收登记表：
一管一管判温度、判时限、判外观、判访视点在不在方案里，异常的整行刷色并单独拎成一页。
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
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

HERE = Path(__file__).resolve().parent
RECV_FILE = HERE / "01_raw_data" / "样本接收明细_扫码枪导出.csv"    # 一管一行，扫码枪吐出来的
TEMP_FILE = HERE / "01_raw_data" / "温控记录_记录仪导出.csv"        # 一趟车一行，记录仪导出的
OUT_FILE = HERE / "02_output" / "样本接收登记表.xlsx"              # 登记总表 + 异常清单 + 中心汇总
LOG_DIR = HERE / "source"                                        # 运行日志（run_log.txt 就落在 source/ 下）

KEY_COL = "样本编号"
RECV_COLS = ["运单号", "样本编号", "受试者编号", "中心编号", "中心名称",
             "访视点", "样本类型", "采样日期", "接收日期", "外观备注"]
TEMP_COLS = ["运单号", "最高温度", "最低温度", "干冰余量"]

# 接收规则照 SOP 抄：超了就不能算合格接收，得先跟中心确认这一管还能不能用。
# 值 = (送达时限小时, 运输温度下限℃, 运输温度上限℃)；下限写 None 表示只管上限（干冰越冷越好）。
SAMPLE_RULES = {
    "全血EDTA": (24, 2.0, 8.0),
    "血清":     (48, 2.0, 8.0),
    "尿液":     (48, 15.0, 25.0),
    "唾液干冰": (72, None, -60.0),
}

# 方案里定死的访视点。不在这张清单里的，登记时就得标出来问一句——多半是中心填错了。
PLAN_VISITS = ["筛选期", "基线", "第1周期D1", "第1周期D8", "第2周期D1", "第2周期D8", "末次访视"]

REG_COLS = ["运单号", "样本编号", "受试者编号", "中心编号", "中心名称", "访视点", "样本类型",
            "采样日期", "接收日期", "运输温度", "外观备注", "接收结论", "异常说明"]

# 三种结论分开刷色：要跟中心确认的刷淡红，记录偏差的刷淡黄，合格的保留白底
BAD_FILL = PatternFill("solid", fgColor="FFC7CE")
WARN_FILL = PatternFill("solid", fgColor="FFEB9C")
HEAD_FILL = PatternFill("solid", fgColor="D9E1F2")

# 结论分三档，来源是接收 SOP 的三级处理（记录备案 / 隔离待评估 / 拒收启动偏差）。
# 全文只用这两档加合格：要中心确认的、记录偏差的。
NEED_CONFIRM, NOTE_ONLY, PASSED = "待中心确认", "有偏差接收", "合格接收"


def read_samples(path: Path) -> list:
    """读扫码枪导出的接收明细，一行一管。

    表尾经常挂着合计行、说明行，不是数据：样本编号一空就丢掉。
    """
    out = []
    with open(path, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            sid = (row.get(KEY_COL) or "").strip()
            if not sid:
                continue                       # 空行、表尾的合计行/说明行都落在这儿
            out.append({c: (row.get(c) or "").strip() for c in RECV_COLS})
    return out


def read_temp(path: Path) -> dict:
    """读温控记录仪导出的运输温度，按运单号归堆：一趟车一行，一趟车管着好几管样本。

    温度存成文本，判的时候再转数字——记录仪偶尔吐空值或 "NA"，
    转不动就得当"没记录"，不能当成 0℃：0℃ 对要求 2~8℃ 的样本是误判。
    """
    out = {}
    with open(path, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            way = (row.get("运单号") or "").strip()
            if not way:
                continue
            out[way] = {c: (row.get(c) or "").strip() for c in TEMP_COLS}
    return out


def to_float(text: str):
    """把温度文本转成数字；空值、NA 这类转不动的返回 None，别当成 0。"""
    try:
        return float(str(text).strip())
    except (TypeError, ValueError):
        return None


def hours_between(start_text: str, end_text: str):
    """采样到接收隔了几小时。时间缺一个就返回 None，交给人工看，不能默认合格。"""
    for fmt in ("%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try:
            start = datetime.strptime(start_text.strip(), fmt)
            end = datetime.strptime(end_text.strip(), fmt)
        except ValueError:
            continue
        return (end - start).total_seconds() / 3600
    return None


def temp_text(temp: dict) -> str:
    """把一趟车的最高/最低温写成一行字，挂到这一管的登记行上。"""
    high, low = to_float(temp.get("最高温度")), to_float(temp.get("最低温度"))
    return f"{low:g}~{high:g}℃" if None not in (high, low) else "无记录"


def check_row(row: dict, temp: dict):
    """判一管样本能不能合格接收，返回 (接收结论, 异常说明)。

    温度按运单判——记录仪只给整趟车的最高/最低温，一管超温，同一趟车的样本都算数，
    这是冷链的基本逻辑；时限按样本类型的时限卡；外观和访视点只记录偏差，不影响接收。
    """
    reasons, severe = [], False
    rule = SAMPLE_RULES.get(row["样本类型"])
    if rule is None:
        reasons.append(f"样本类型「{row['样本类型']}」不在接收规则里")
        severe = True
        limit_h, low, high = None, None, None
    else:
        limit_h, low, high = rule

    t_high, t_low = to_float(temp.get("最高温度")), to_float(temp.get("最低温度"))
    if t_high is None and t_low is None:
        reasons.append("无温控记录")            # 记录仪没导、运单号打错，都落这儿
        severe = True
    else:
        if high is not None and t_high is not None and t_high > high:
            reasons.append(f"运输温度超上限（{t_high:g}℃ > {high:g}℃）")
            severe = True
        if low is not None and t_low is not None and t_low < low:
            reasons.append(f"运输温度低于下限（{t_low:g}℃ < {low:g}℃）")
            severe = True

    hours = hours_between(row["采样日期"], row["接收日期"])
    if limit_h is not None:
        if hours is None:
            reasons.append("采样/接收时间缺失，判不了时限")
            severe = True
        elif hours > limit_h:
            reasons.append(f"送达超时（{hours:.0f}h > {limit_h}h）")
            severe = True

    if row["外观备注"]:
        reasons.append(f"外观异常（{row['外观备注']}）")

    if row["访视点"] not in PLAN_VISITS:
        reasons.append(f"访视点不在方案内（{row['访视点']}）")

    if not reasons:
        return PASSED, ""
    return (NEED_CONFIRM if severe else NOTE_ONLY), "；".join(reasons)


def write_register(rows: list, out_path: Path) -> Path:
    """登记总表 + 异常样本清单 + 中心汇总，三个页签写进一个 xlsx。

    接收岗不是看表格，是看颜色：整行刷色的那几管单独拎出来，
    剩下白底的直接入库，不用一行一行读结论列。
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "登记总表"
    ws.append(REG_COLS)
    for col in range(1, len(REG_COLS) + 1):
        head = ws.cell(row=1, column=col)
        head.font, head.fill = Font(bold=True), HEAD_FILL
        head.alignment = Alignment(horizontal="center")

    for row in rows:
        ws.append([row[c] for c in REG_COLS])
        if row["接收结论"] == PASSED:
            continue
        fill = BAD_FILL if row["接收结论"] == NEED_CONFIRM else WARN_FILL
        for col in range(1, len(REG_COLS) + 1):
            ws.cell(row=ws.max_row, column=col).fill = fill
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(REG_COLS))}{ws.max_row}"

    # 第二页：只留异常，接收岗直接拿这页去发邮件找中心确认
    bad = [r for r in rows if r["接收结论"] != PASSED]
    ws2 = wb.create_sheet("异常样本清单")
    ws2.append(["样本编号", "中心编号", "中心名称", "访视点", "样本类型",
                "到达日期", "运单号", "接收结论", "异常说明"])
    for row in bad:
        ws2.append([row["样本编号"], row["中心编号"], row["中心名称"], row["访视点"], row["样本类型"],
                    row["接收日期"], row["运单号"], row["接收结论"], row["异常说明"]])
    ws2.freeze_panes = "A2"

    # 第三页：按中心汇总，给项目组报数用
    ws3 = wb.create_sheet("中心汇总")
    ws3.append(["中心编号", "中心名称", "接收管数", PASSED, NEED_CONFIRM, NOTE_ONLY, "合格率"])
    for center in sorted({r["中心编号"] for r in rows}):
        group = [r for r in rows if r["中心编号"] == center]
        ok = sum(1 for r in group if r["接收结论"] == PASSED)
        ws3.append([center, group[0]["中心名称"], len(group), ok,
                    sum(1 for r in group if r["接收结论"] == NEED_CONFIRM),
                    sum(1 for r in group if r["接收结论"] == NOTE_ONLY),
                    round(ok / len(group), 3)])

    for sheet in (ws, ws2, ws3):              # 列宽按表头给够，中文列名别挤成竖条
        for idx, head in enumerate(sheet[1], 1):
            sheet.column_dimensions[get_column_letter(idx)].width = max(10, len(str(head.value)) * 2.2)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out_path))
    return out_path


def build_rows(samples: list, temp: dict) -> list:
    """把温控记录挂到每一管上，逐管判定，拼成登记行。"""
    rows = []
    for s in samples:
        t = temp.get(s["运单号"], {})
        verdict, why = check_row(s, t)
        rows.append({**s, "运输温度": temp_text(t), "接收结论": verdict, "异常说明": why})
    return rows


def main() -> bool:
    samples = read_samples(RECV_FILE)
    temp = read_temp(TEMP_FILE)
    print(f"读到 {len(samples)} 管样本、{len(temp)} 趟运单的温控记录")

    rows = build_rows(samples, temp)
    out = write_register(rows, OUT_FILE)
    for verdict in (PASSED, NEED_CONFIRM, NOTE_ONLY):
        print(f"  {verdict}：{sum(1 for r in rows if r['接收结论'] == verdict)} 管")
    print(f"登记表已生成：{out.name}")

    # ---------- 以下为校验，正文里没有 ----------
    print("\n===== 回读校验 =====")
    from openpyxl import load_workbook
    book = load_workbook(str(out))
    reg, bad_sheet, sum_sheet = book["登记总表"], book["异常样本清单"], book["中心汇总"]

    def col(sheet, name):
        idx = [c.value for c in sheet[1]].index(name) + 1
        return [sheet.cell(row=i, column=idx).value for i in range(2, sheet.max_row + 1)]

    def find(sid):
        return next(r for r in rows if r["样本编号"] == sid)

    ids = [r["样本编号"] for r in rows]
    hot = [r for r in rows if r["运单号"] == "T260915002"]        # 4.0~11.8，冷藏超上限
    cold = [r for r in rows if r["运单号"] == "T260915003"]       # 0.4~1.2，低于下限疑似冻结
    dry = [r for r in rows if r["运单号"] == "T260915004"]        # 干冰 -72.5~-68.0，正常
    late = [r for r in rows if "送达超时" in r["异常说明"]]
    note = [r for r in rows if r["接收结论"] == NOTE_ONLY]
    look = [r for r in rows if "外观异常" in r["异常说明"]]
    lead0 = [r for r in rows if r["样本编号"].startswith("01001")][0]
    no_temp = [r for r in rows if "无温控记录" in r["异常说明"]]
    unknown = [r for r in rows if "不在接收规则里" in r["异常说明"]]
    bad_ids = [r["样本编号"] for r in rows if r["接收结论"] != PASSED]

    sum_rows = [{"管数": sum_sheet.cell(row=i, column=3).value,
                 "合格": sum_sheet.cell(row=i, column=4).value,
                 "合格率": sum_sheet.cell(row=i, column=7).value}
                for i in range(2, sum_sheet.max_row + 1)]
    checks = [
        ("登记总表行数 = 样本数 + 表头，表尾脏行没被算成样本",
         reg.max_row == len(samples) + 1, f"{reg.max_row} 行 / {len(samples)} 管"),
        ("中心编号前导零没丢（01001 原样在表里）",
         lead0["中心编号"] == "01001" and "01001" in col(reg, "中心编号"),
         lead0["样本编号"]),
        ("冷藏超上限那趟车，6 管全部标出来",
         len(hot) == 6 and all(r["接收结论"] == NEED_CONFIRM for r in hot),
         f"{len(hot)} 管，第一管：{hot[0]['异常说明']}"),
        ("低于下限（疑似冻结）也判出来，不是只卡上限",
         len(cold) == 5 and all("低于下限" in r["异常说明"] for r in cold),
         f"{len(cold)} 管，第一管：{cold[0]['异常说明']}"),
        ("干冰趟只卡上限：-68℃ 那批不判异常",
         len(dry) == 4 and all(r["接收结论"] == PASSED for r in dry),
         f"{len(dry)} 管全部合格"),
        ("干冰不足那趟（最高 -52℃）全判待确认",
         all(r["接收结论"] == NEED_CONFIRM for r in rows if r["运单号"] == "T260915005"),
         find("01003001-SAL01")["异常说明"]),
        ("送达超时按样本类型时限判（全血 24h）",
         len(late) == 1 and late[0]["样本类型"] == "全血EDTA" and "47h" in late[0]["异常说明"],
         late[0]["异常说明"]),
        ("外观异常只记偏差，不影响接收",
         len(look) == 2 and all(r["接收结论"] == NOTE_ONLY for r in look)
         and len(note) == 3,
         "；".join(r["异常说明"] for r in look)),
        ("访视点不在方案里的被标出来",
         any("访视点不在方案内" in r["异常说明"] for r in rows),
         find("01001006-EDTA01")["异常说明"]),
        ("样本类型不认识的不猜，直接待确认",
         len(unknown) == 1 and unknown[0]["接收结论"] == NEED_CONFIRM, unknown[0]["样本类型"]),
        ("运单号查不到温控记录的，判待确认不是判合格",
         len(no_temp) == 2 and all(r["接收结论"] == NEED_CONFIRM for r in no_temp),
         f"{len(no_temp)} 管"),
        ("异常清单行数 = 非合格管数，且结论列只有异常那两种",
         bad_sheet.max_row == len(bad_ids) + 1
         and set(col(bad_sheet, "接收结论")) == {NEED_CONFIRM, NOTE_ONLY},
         f"{bad_sheet.max_row - 1} 行 / 非合格 {len(bad_ids)} 管"),
        ("登记总表里异常行整行刷了色，合格行没刷",
         all(reg.cell(row=i, column=1).fill.patternType == "solid"
             for i in range(2, reg.max_row + 1) if reg.cell(row=i, column=2).value in bad_ids)
         and all(reg.cell(row=i, column=1).fill.patternType is None
                 for i in range(2, reg.max_row + 1) if reg.cell(row=i, column=2).value not in bad_ids),
         f"异常 {len(bad_ids)} 行"),
        ("中心汇总的合格率按中心分开算",
         sum(r["管数"] for r in sum_rows) == len(rows)
         and all(abs(r["合格率"] - round(r["合格"] / r["管数"], 3)) < 1e-9 for r in sum_rows),
         f"{sum_sheet.max_row - 1} 个中心"),
        ("每管样本只在登记总表出现一次",
         len(ids) == len(set(ids)) and reg.max_row == len(ids) + 1, f"{len(set(ids))} 个编号"),
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
