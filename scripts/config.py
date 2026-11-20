# -*- coding: utf-8 -*-
"""参数装载层：高德 key、输入模版解析、拜访频次规范化。

key 从环境变量 AMAP_KEY 读取（惰性，调用时读，不在 import 时读），
绝不硬编码、绝不进 Excel、绝不在日志回显。
"""

import csv
import os

import openpyxl


def get_amap_key():
    """惰性读取高德 Web 服务 key。缺失则报错退出。"""
    key = os.environ.get("AMAP_KEY", "").strip()
    if not key:
        raise SystemExit("缺少高德 key：请先 export AMAP_KEY=你的key，或 run.py 传 --amap-key")
    return key


# 拜访频次别名 → 规范值
_FREQ_ALIASES = {
    "一周一次": "一周一次",
    "每周": "一周一次",
    "周": "一周一次",
    "两周一次": "两周一次",
    "两周": "两周一次",
    "四周一次": "四周一次",
    "四周": "四周一次",
    "一月一次": "四周一次",
    "": "一周一次",
}


def norm_freq(v):
    """规范化拜访频次；未知值告警并回落「一周一次」。"""
    if v is None:
        return "一周一次"
    s = str(v).strip()
    if s in _FREQ_ALIASES:
        return _FREQ_ALIASES[s]
    print(f"  ⚠ 未知拜访频次「{s}」，按「一周一次」处理")
    return "一周一次"


# 输入模版表头别名
_NAME_ALIASES = ("店名", "售点名称", "名称", "门店", "店铺")
_ADDR_ALIASES = ("详细地址", "地址", "门店地址", "位置")
_FREQ_COL_ALIASES = ("拜访频次", "频次", "拜访频率")
_NOTE_ALIASES = ("备注", "说明")


def _pick_col(headers, aliases):
    for a in aliases:
        for i, h in enumerate(headers):
            if h is not None and str(h).strip() == a:
                return i
    return None


def load_input(path):
    """读输入模版（xlsx 单 sheet 首行表头，或 csv），返回 [{店名, 地址, 拜访频次, 备注}]。"""
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active
        all_rows = list(ws.iter_rows(values_only=True))
    else:
        with open(path, encoding="utf-8-sig", newline="") as f:
            all_rows = list(csv.reader(f))

    if not all_rows:
        return []

    headers = [str(h).strip() if h is not None else "" for h in all_rows[0]]
    ic_name = _pick_col(headers, _NAME_ALIASES)
    ic_addr = _pick_col(headers, _ADDR_ALIASES)
    ic_freq = _pick_col(headers, _FREQ_COL_ALIASES)
    ic_note = _pick_col(headers, _NOTE_ALIASES)
    if ic_name is None:
        raise SystemExit(f"输入文件缺少「店名/售点名称」列，实际表头：{headers}")

    rows = []
    for line in all_rows[1:]:
        name = line[ic_name] if ic_name < len(line) else None
        if name is None or not str(name).strip():
            continue
        addr = line[ic_addr] if ic_addr is not None and ic_addr < len(line) else ""
        freq = line[ic_freq] if ic_freq is not None and ic_freq < len(line) else ""
        note = line[ic_note] if ic_note is not None and ic_note < len(line) else ""
        rows.append({
            "店名": str(name).strip(),
            "地址": str(addr).strip() if addr else "",
            "拜访频次": norm_freq(freq),
            "备注": str(note).strip() if note else "",
        })
    return rows
