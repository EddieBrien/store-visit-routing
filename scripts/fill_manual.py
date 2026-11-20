# -*- coding: utf-8 -*-
"""把用户补好的「待人工标注清单」回填到标点结果里（回填阶段）。

规则：
  - 用户在「状态」列补了正确地址的 → 按该地址编码；
  - 状态仍是 no_match/空 的 → 按「原地址」（备用地址）编码兜底；
  - 状态写「先空着」的 → 留空，不填。

读 <out_dir>/待人工标注清单.xlsx（用户已编辑），回写 <out_dir>/poi_cache.csv。
"""

import csv
import json
import os
import time
import urllib.parse
import urllib.request

import openpyxl

import config
import poi_geocode as pg

BLANK_FLAGS = ("", "no_match", "ambiguous", "error")


def geocode(address, city, district):
    """地址 → GCJ-02 坐标（锁定指定区）。"""
    addr = (address or "").strip()
    if not addr:
        return "", "", ""
    if district:
        d_short = district.rstrip("区市县")
        c_short = city.rstrip("市")
        if d_short and d_short not in addr and c_short not in addr:
            addr = district + addr
    url = (
        "https://restapi.amap.com/v3/geocode/geo?address={addr}&city={city}&key={key}&output=JSON"
    ).format(addr=urllib.parse.quote(addr), city=urllib.parse.quote(city), key=pg.AMAP_KEY)
    with urllib.request.urlopen(url, timeout=20) as r:
        j = json.loads(r.read().decode("utf-8"))
    geos = j.get("geocodes", []) or []
    if geos:
        loc = geos[0].get("location", "")
        if isinstance(loc, str) and "," in loc:
            lng, lat = loc.split(",")
            return lng, lat, geos[0].get("formatted_address", "")
    return "", "", ""


def load_manual(out_dir):
    """读用户编辑后的 待人工标注清单.xlsx（列：店名=1、状态=3、原地址=4）。"""
    path = os.path.join(out_dir, "待人工标注清单.xlsx")
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    manual = []
    for r in range(2, ws.max_row + 1):
        name = ws.cell(r, 1).value
        if name is None or not str(name).strip():
            continue
        manual.append({
            "店名": str(name).strip(),
            "状态": str(ws.cell(r, 3).value or "").strip(),
            "原地址": str(ws.cell(r, 4).value or "").strip(),
        })
    return manual


def _blank_row(name, note, freq, addr):
    return {"店名": name, "备注": note, "拜访频次": freq, "地址": addr,
            "关键词": "", "状态": "no_match",
            "经度": "", "纬度": "", "POI名称": "", "POI地址": "", "POI_ID": ""}


def main(input_path, out_dir, city, district):
    pg.load_key()
    stores = config.load_input(input_path)

    cache = {}
    cache_path = os.path.join(out_dir, "poi_cache.csv")
    try:
        with open(cache_path, encoding="utf-8-sig") as f:
            for row in csv.DictReader(f):
                cache[row.get("店名", "")] = row
    except FileNotFoundError:
        pass

    manual = load_manual(out_dir)
    manual_by_name = {m["店名"]: m for m in manual}

    # 预处理输入：店名清洗 + 备注合并
    entries = []
    for s in stores:
        clean, embedded_note = pg.strip_notes(s["店名"])
        entries.append({
            "店名": clean,
            "备注": pg.merge_note(embedded_note, s["备注"]),
            "拜访频次": s["拜访频次"],
            "地址": s["地址"],
        })

    results = []
    geocoded = 0
    still_blank = []
    for e in entries:
        name, note, freq, addr = e["店名"], e["备注"], e["拜访频次"], e["地址"]
        cr = cache.get(name)
        if cr and cr.get("状态") in ("exact", "base"):
            cr = dict(cr)
            cr["备注"] = note
            cr["拜访频次"] = freq
            cr["地址"] = addr
            results.append(cr)
            continue

        m = manual_by_name.get(name)
        renamed = name
        if m is None:
            # 用户可能改了店名（如「金鑫百货商店」→「金鑫商店」），按字符重叠匹配
            best, bestscore = None, 0
            for mn, mm in manual_by_name.items():
                sc = len(set(mn) & set(name))
                if sc > bestscore:
                    bestscore, best = sc, (mn, mm)
            if best and bestscore >= 2:
                renamed, m = best

        if m is None:
            still_blank.append(name)
            results.append(_blank_row(name, note, freq, addr))
            continue

        status = m["状态"]
        if status == "先空着":
            still_blank.append(renamed)
            row = _blank_row(renamed, note, freq, addr)
            row["状态"] = "manual_blank"
            results.append(row)
            continue

        address = m["原地址"] if status in BLANK_FLAGS else status
        lng, lat, _fmt = geocode(address, city, district)
        if lng and lat:
            geocoded += 1
            results.append({"店名": renamed, "备注": note, "拜访频次": freq, "地址": addr,
                            "关键词": "", "状态": "geocode",
                            "经度": lng, "纬度": lat,
                            "POI名称": "", "POI地址": address, "POI_ID": ""})
            print(f"  ✓ {renamed} ← 「{address}」 → {lng},{lat}")
        else:
            still_blank.append(renamed)
            row = _blank_row(renamed, note, freq, addr)
            row["状态"] = "manual_blank"
            row["POI地址"] = address
            results.append(row)
            print(f"  ✗ {renamed} 「{address}」 编码失败，留空")
        time.sleep(0.3)

    # 区界校验：以 POI 结果定大致范围，兜底编码越界的告警
    poi_pts = [(float(r["经度"]), float(r["纬度"])) for r in results
               if r["状态"] in ("exact", "base") and r.get("经度")]
    if poi_pts:
        mn_lng = min(p[0] for p in poi_pts) - 0.01
        mx_lng = max(p[0] for p in poi_pts) + 0.01
        mn_lat = min(p[1] for p in poi_pts) - 0.01
        mx_lat = max(p[1] for p in poi_pts) + 0.01
        for r in results:
            if r["状态"] == "geocode" and r.get("经度"):
                lng, lat = float(r["经度"]), float(r["纬度"])
                if not (mn_lng <= lng <= mx_lng and mn_lat <= lat <= mx_lat):
                    print(f"  ⚠ 越界告警：{r['店名']} {lng},{lat} 疑似跑出{district or '目标区'}，请核对")

    # 回写 poi_cache.csv
    with open(cache_path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=pg.CACHE_FIELDS)
        w.writeheader()
        for r in results:
            w.writerow({k: r.get(k, "") for k in pg.CACHE_FIELDS})

    filled = sum(1 for r in results if r["状态"] in ("exact", "base", "geocode") and r.get("经度"))
    print("\n=== 回填完成 ===")
    print(f"共 {len(results)} 家：POI精确 {sum(1 for r in results if r['状态'] in ('exact','base'))} "
          f"/ 地址兜底 {geocoded} / 留空 {len(still_blank)}")
    if still_blank:
        print("仍留空的店：", "、".join(still_blank))
    return len(still_blank)


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 4:
        raise SystemExit("用法：python3 fill_manual.py <input> <out_dir> <city> [district]")
    main(sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4] if len(sys.argv) > 4 else "")
