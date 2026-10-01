#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""课表空闲时段汇总器

把「每人一张课表识别结果 JSON」汇总成值班排班用的《空闲时间表》xlsx。

用法:
    python scripts/build_roster.py --json-dir sample_timetables/roster_json
                                   [--out 输出目录] [--name 文件名]
                                   [--template assets/空闲时间表模板.xlsx]

铁律：复制模板生成新文件，绝不修改模板原件。
"""

import argparse
import json
import re
import sys
from copy import copy
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:  # pragma: no cover
    pass

DAYS = ["一", "二", "三", "四", "五"]
PERIODS = ["一", "二", "三", "四"]
CN_NUM = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "日": 7, "天": 7}
ARAB = {"0": 0, "1": 1, "2": 2, "3": 3, "4": 4, "5": 5, "6": 6, "7": 7, "8": 8, "9": 9}


# ---------------------------------------------------------------- 标签归一化

def norm_day(label):
    """'周一' / '星期一' / '1' / 'Mon' -> '一'..'五'"""
    if label is None:
        return None
    s = str(label).strip()
    for ch in s:
        if ch in CN_NUM and 1 <= CN_NUM[ch] <= 5:
            return DAYS[CN_NUM[ch] - 1]
    m = re.search(r"[1-5]", s)
    if m:
        return DAYS[int(m.group()) - 1]
    return None


def norm_period(label):
    """'第一大节' / '第2大节' / '1-2节' / '一' -> '一'..'四'"""
    if label is None:
        return None
    s = str(label).strip()
    m = re.search(r"第\s*([一二三四1-4])", s)
    if m:
        return _to_period(m.group(1))
    m = re.search(r"([1-4])\s*[-—~]\s*[1-9]", s) or re.search(r"([1-4])", s)
    if m:
        return _to_period(m.group(1))
    for ch in s:
        if ch in ("一", "二", "三", "四"):
            return ch
    return None


def _to_period(tok):
    if tok in ("一", "二", "三", "四"):
        return tok
    if tok.isdigit() and 1 <= int(tok) <= 4:
        return PERIODS[int(tok) - 1]
    return None


def natural_key(path):
    """文件名自然排序：001_... 排在 002_... 之前，10 排在 9 之后"""
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", path.name)]


# ---------------------------------------------------------------- 读取 JSON

def load_people(json_dir, verbose=True):
    files = sorted([p for p in Path(json_dir).glob("*.json")], key=natural_key)
    if not files:
        raise SystemExit(f"[错误] {json_dir} 下没有找到任何 .json 识别结果")

    people, problems = [], []
    for fp in files:
        try:
            data = json.loads(fp.read_text(encoding="utf-8"))
        except Exception as e:
            problems.append(f"{fp.name}: JSON 解析失败 {e}")
            continue

        name = str(data.get("name", "")).strip()
        if not name:
            problems.append(f"{fp.name}: 缺少 name 字段，已跳过")
            continue

        free = {}
        for day, periods in (data.get("free") or {}).items():
            d = norm_day(day)
            if d is None:
                problems.append(f"{fp.name}: 无法识别的星期「{day}」，已忽略")
                continue
            clean = []
            for p in periods or []:
                np_ = norm_period(p)
                if np_ is None:
                    problems.append(f"{fp.name}: 无法识别的大节「{p}」，已忽略")
                    continue
                if np_ not in clean:
                    clean.append(np_)
            free[d] = sorted(clean, key=lambda x: PERIODS.index(x))

        for d in DAYS:
            free.setdefault(d, [])

        uncertain = []
        for it in data.get("uncertain") or []:
            uncertain.append({
                "day": norm_day(it.get("day")) or str(it.get("day", "")),
                "period": norm_period(it.get("period")) or str(it.get("period", "")),
                "reason": str(it.get("reason", "")).strip(),
            })

        people.append({
            "name": name,
            "class": str(data.get("class", "")).strip(),
            "free": free,
            "uncertain": uncertain,
            "source": str(data.get("source_image", fp.name)),
            "region_source": str(data.get("region_source", "")).strip(),
            "notes": str(data.get("notes", "")).strip(),
            "json_file": fp.name,
        })

    if verbose and problems:
        print("[校验提示]")
        for p in problems:
            print("  -", p)
    return people


# ---------------------------------------------------------------- 汇总

def build_grid(people):
    grid = {(d, p): [] for d in DAYS for p in PERIODS}
    for person in people:
        for d in DAYS:
            for p in person["free"].get(d, []):
                grid[(d, p)].append(person["name"])
    return grid


def detect_same_name(people):
    """同名不同班 -> 存疑记录"""
    seen, out = {}, []
    for person in people:
        key = person["name"]
        if key in seen and seen[key] != person["class"]:
            out.append({
                "name": key,
                "class": f"{seen[key]} / {person['class']}",
                "day": "",
                "period": "",
                "reason": "同名：出现于多个班级，请确认是否同一人",
                "source": person["source"],
            })
        seen.setdefault(key, person["class"])
    return out


def collect_uncertain(people):
    rows = []
    for person in people:
        for it in person["uncertain"]:
            rows.append({
                "name": person["name"],
                "class": person["class"],
                "day": it["day"],
                "period": it["period"],
                "reason": it["reason"],
                "source": person["source"],
            })
    return rows


def free_text(person):
    parts = []
    for d in DAYS:
        if person["free"].get(d):
            parts.append(f"周{d}[{''.join(person['free'][d])}]")
    return "；".join(parts) if parts else "（无空闲）"


# ---------------------------------------------------------------- 写 xlsx

def locate_labels(ws):
    """返回 (列映射 {day: col}, 行映射 {period: row})"""
    col_map, row_map = {}, {}
    for col in range(2, ws.max_column + 1):
        d = norm_day(ws.cell(row=1, column=col).value)
        if d and d not in col_map:
            col_map[d] = col
    for row in range(2, ws.max_row + 1):
        p = norm_period(ws.cell(row=row, column=1).value)
        if p and p not in row_map:
            row_map[p] = row
    if not col_map or not row_map:
        raise SystemExit("[错误] 模板 Sheet1 未识别到星期表头或大节行标签，请检查模板结构")
    return col_map, row_map


def write_workbook(template, people, grid, uncertains, out_path):
    wb = openpyxl.load_workbook(template)
    ws = wb["Sheet1"] if "Sheet1" in wb.sheetnames else wb.worksheets[0]
    ws.title = "空闲时间表"
    col_map, row_map = locate_labels(ws)

    # 清空数据区（保留样式与表头标签）
    for d, col in col_map.items():
        for p, row in row_map.items():
            ws.cell(row=row, column=col).value = None

    # 填姓名
    for d, col in col_map.items():
        for p, row in row_map.items():
            names = grid.get((d, p), [])
            cell = ws.cell(row=row, column=col)
            cell.value = " ".join(names) if names else None
            cell.alignment = Alignment(wrap_text=True, vertical="center",
                                       horizontal=getattr(cell.alignment, "horizontal", None) or "left")

    # 列宽与自动行高
    for d, col in col_map.items():
        ws.column_dimensions[ws.cell(row=1, column=col).column_letter].width = 42
    for p, row in row_map.items():
        ws.row_dimensions[row].height = None

    # 存疑复核 sheet
    ws2 = wb["Sheet2"] if "Sheet2" in wb.sheetnames else wb.create_sheet()
    ws2.title = "存疑复核"
    _write_table(ws2, ["姓名", "班级", "星期", "大节", "原因", "原图"],
                 [[u["name"], u["class"], f"周{u['day']}" if u["day"] else "",
                   f"第{u['period']}大节" if u["period"] else "", u["reason"], u["source"]]
                  for u in uncertains], widths=[12, 14, 8, 12, 46, 16])

    # 逐人明细 sheet
    ws3 = wb["Sheet3"] if "Sheet3" in wb.sheetnames else wb.create_sheet()
    ws3.title = "逐人明细"
    _write_table(ws3, ["序号", "班级", "姓名", "空闲时段", "存疑数", "区域判定", "原图"],
                 [[i, p["class"], p["name"], free_text(p), len(p["uncertain"]),
                   p["region_source"], p["source"]]
                  for i, p in enumerate(people, 1)],
                 widths=[6, 14, 10, 40, 8, 12, 16])

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)


def _write_table(ws, headers, rows, widths):
    ws.delete_rows(1, ws.max_row)
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[cell.column_letter].width = widths[c - 1]
    ws.freeze_panes = "A2"
    for r, row in enumerate(rows, 2):
        for c, v in enumerate(row, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.alignment = Alignment(wrap_text=(c in (5, 4)), vertical="center")


# ---------------------------------------------------------------- main

def main():
    here = Path(__file__).resolve().parent
    root = here.parent
    ap = argparse.ArgumentParser(description="课表空闲时段汇总器")
    ap.add_argument("--json-dir", required=True, help="识别结果 JSON 所在目录")
    ap.add_argument("--out", default=None, help="输出目录，默认为 JSON 目录的上级")
    ap.add_argument("--name", default=None, help="输出文件名，默认 空闲时间表_YYYYMMDD.xlsx")
    ap.add_argument("--template", default=str(root / "assets" / "空闲时间表模板.xlsx"),
                    help="模板 xlsx（只读，绝不修改）")
    args = ap.parse_args()

    json_dir = Path(args.json_dir)
    template = Path(args.template)
    if not template.exists():
        raise SystemExit(f"[错误] 找不到模板：{template}")

    people = load_people(json_dir)
    grid = build_grid(people)
    uncertains = collect_uncertain(people) + detect_same_name(people)

    out_dir = Path(args.out) if args.out else json_dir.parent
    out_name = args.name or f"空闲时间表_{datetime.now():%Y%m%d}.xlsx"
    out_path = out_dir / out_name

    write_workbook(template, people, grid, uncertains, out_path)

    # 控制台摘要
    print(f"[完成] 识别人数：{len(people)}")
    print(f"[完成] 存疑条目：{len(uncertains)}")
    print("各时段空闲人数（行=大节，列=星期）：")
    print("      " + "  ".join(f"周{d}" for d in DAYS))
    for p in PERIODS:
        print(f"第{p}大节 " + "  ".join(f"{len(grid[(d, p)]):>3}" for d in DAYS))
    print(f"[输出] {out_path}")
    if uncertains:
        print("[存疑清单]")
        for u in uncertains:
            loc = f"周{u['day']}第{u['period']}大节" if u["day"] else "—"
            print(f"  - {u['name']}（{u['class']}）{loc}：{u['reason']}")


if __name__ == "__main__":
    main()
