"""重造 20261018 示例的输入数据：往 01_raw_data/待整理/ 里铺 6 个命名乱象文件。

覆盖的边界：中文式日期、英文命名带空格、(2) 副本、检项英文缩写、已规范名、认不出的旧名。
跑法：python build_fixtures.py        （先清空再重建 01_raw_data/待整理/）
"""

import shutil
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
WORK_DIR = BASE_DIR / "01_raw_data" / "待整理"

# (文件名, 内容首行)，内容只是占位，重点在名字
FILES = [
    ("2026-3-5_样品S001_血常规.csv", "样品编号,检项,结果\nS001,血常规,4.2\n"),
    ("sample S002 CBC 20260306.csv", "sample,item,result\nS002,CBC,4.5\n"),
    ("血常规 2026.3.6 S002 (2).csv", "sample,item,result\nS002,CBC,4.7\n"),
    ("20260308_S003_肝功能.csv", "样品编号,检项,结果\nS003,肝功能,31\n"),
    ("2026年3月9日_样品S004_尿常规.csv", "样品编号,检项,结果\nS004,尿常规,阴性\n"),
    ("结果最终版改2.xlsx", "占位\n"),
]


def main():
    if WORK_DIR.exists():
        shutil.rmtree(WORK_DIR)
    WORK_DIR.mkdir(parents=True)
    # 重建环境时把输出也清掉，保证每次演示都从干净状态开始
    out_dir = BASE_DIR / "02_output"
    if out_dir.exists():
        shutil.rmtree(out_dir)
    for name, content in FILES:
        (WORK_DIR / name).write_text(content, encoding="utf-8")
    print(f"已重建 {WORK_DIR}（{len(FILES)} 个文件）")


if __name__ == "__main__":
    main()
