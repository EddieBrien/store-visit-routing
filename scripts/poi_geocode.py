# -*- coding: utf-8 -*-
"""按店名精确 POI 定位（标点阶段）。

方法：店名先去掉业务代码后缀（-BH/-DL/-AH/-CH/-CL 等），
按清理后的店名调高德 place/text（city=用户给的市, citylimit=true），
过滤 adname==区，按「完全同名 > 去括号同名」匹配取真实门店坐标。

产出：
  - <out_dir>/poi_cache.csv            POI 搜索结果缓存（统一 11 字段）
  - <out_dir>/待人工标注清单.xlsx       搜不到/歧义店（供人工回填）

坐标体系全程 GCJ-02，不转 WGS-84。
"""

import csv
import json
import os
import re
import time
import urllib.parse
import urllib.request

from openpyxl import Workbook

import config

BASE = os.path.dirname(os.path.abspath(__file__))
ASSET_TEMPLATE = os.path.join(BASE, "..", "assets", "导入模板.xlsx")

AMAP_KEY = None
CITY = ""
DISTRICT = ""
OUT_DIR = ""
SLEEP = 0.35  # 限流间隔（秒）；高德个人 key 3 QPS


def load_key():
    """设置全局 AMAP_KEY（惰性读 env）。函数名保留，fill_manual / pai_xian 会调用。"""
    global AMAP_KEY
    AMAP_KEY = config.get_amap_key()


# 店名末尾的人工备注（用户手打的提示，不算店名一部分）
NOTE_PATTERNS = ["没货不打算进", "定位错误没找到", "没找到"]


def strip_notes(name):
    """去掉店名里的人工备注，返回 (干净店名, 备注)。"""
    note = ""
    for pat in NOTE_PATTERNS:
        if pat in name:
            note = (note + " " + pat).strip()
        name = name.replace(f"（{pat}）", "").replace(f"({pat})", "").replace(pat, "")
    name = re.sub(r"[（(]\s*[)）]", "", name).strip()
    return name, note


def merge_note(embedded, template):
    parts = [p for p in (embedded, template) if p]
    return " ".join(parts)


def simplify(name):
    """把店名里的异体/繁体字转成高德常用简体，提高命中。"""
    table = {"億": "亿", "栢": "柏", "荘": "庄", "麺": "面"}
    return "".join(table.get(c, c) for c in name)


def clean_keyword(name):
    """去掉店名末尾的业务代码后缀，得到更接近真实 POI 名的搜索词。

    例：家里饭铁锅炖-BH → 家里饭铁锅炖
        金鹏超市(世纪路东段)－BH－腰部高产 → 金鹏超市(世纪路东段)
    注意：只删业务代码后缀，不动分店括号。
    """
    n = simplify(name.strip())
    n = re.sub(r'[-－—]\s*[A-Z]{1,3}\s*$', '', n)
    n = re.sub(r'[－—]\s*[^()（）]{1,12}\s*$', '', n)
    n = re.sub(r'(?<!\d)\d\s*$', '', n)
    return n.strip()


def normalize(name):
    """去掉括号后缀，用于「去括号同名」比较。"""
    return re.sub(r"[（(].*?[)）]", "", name).strip()


def s(v):
    """把高德返回的字段安全转成字符串（可能为 [] / None / 非字符串）。"""
    if v is None:
        return ""
    if isinstance(v, str):
        return v.strip()
    if isinstance(v, (list, tuple)):
        return ""
    return str(v)


def poi_search(keyword):
    url = (
        "https://restapi.amap.com/v3/place/text"
        "?keywords={kw}&city={city}&citylimit=true&offset=10&key={key}&output=JSON"
    ).format(
        kw=urllib.parse.quote(keyword),
        city=urllib.parse.quote(CITY),
        key=AMAP_KEY,
    )
    with urllib.request.urlopen(url, timeout=20) as resp:
        return json.loads(resp.read().decode("utf-8"))


def match(keyword, pois):
    """在 POI 列表中挑最匹配的一个。返回 (status, poi|None)。

    status: exact 完全同名 / base 去括号同名(唯一) / ambiguous 去括号同名(多个) / no_match 无
    """
    if not pois:
        return "no_match", None
    if DISTRICT:
        pois = [p for p in pois if p.get("adname") == DISTRICT]
    if not pois:
        return "no_match", None

    exact = [p for p in pois if s(p.get("name")) == keyword]
    if exact:
        return "exact", exact[0]

    base = [p for p in pois if normalize(s(p.get("name"))) == normalize(keyword)]
    if len(base) == 1:
        return "base", base[0]
    if len(base) > 1:
        return "ambiguous", base[0]
    return "no_match", None


def suggest(keyword):
    """对没搜到的店，试几个关键词拿最佳候选 POI（或 None）。

    只作「候选」给用户人工确认，不自动填入。
    """
    candidates = [keyword]
    base = normalize(keyword)
    if base and base != keyword:
        candidates.append(base)
    for tail in ("超市", "商店", "便利店", "便利"):
        if keyword.endswith(tail):
            candidates.append(keyword[:-len(tail)])
    for kw in dict.fromkeys(candidates):
        try:
            j = poi_search(kw)
        except Exception:
            continue
        pois = j.get("pois", []) or []
        if DISTRICT:
            pois = [p for p in pois if p.get("adname") == DISTRICT]
        for p in pois:
            pn = s(p.get("name"))
            if len(set(pn) & set(keyword)) >= 2:
                loc = s(p.get("location"))
                if "," in loc:
                    lng, lat = loc.split(",")
                    return {"名称": pn, "地址": s(p.get("address")), "经度": lng, "纬度": lat}
        time.sleep(SLEEP)
    return None


CACHE_FIELDS = ["店名", "备注", "拜访频次", "地址", "关键词", "状态",
                "经度", "纬度", "POI名称", "POI地址", "POI_ID"]


def load_cache(out_dir):
    cache = {}
    try:
        with open(os.path.join(out_dir, "poi_cache.csv"), encoding="utf-8-sig") as f:
            for row in csv.DictReader(f):
                cache[row.get("店名", "")] = row
    except FileNotFoundError:
        pass
    return cache


def save_cache(out_dir, rows):
    with open(os.path.join(out_dir, "poi_cache.csv"), "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=CACHE_FIELDS)
        w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k, "") for k in CACHE_FIELDS})


def nav_url(name, lng, lat):
    """高德导航一键链接（GCJ-02）。"""
    to = f"{lng},{lat},{urllib.parse.quote(name)}"
    return f"https://uri.amap.com/navigation?to={to}&mode=car&coordinate=gaode"


def main(input_path, out_dir, city, district):
    """标点阶段。返回待人工店数。"""
    global CITY, DISTRICT, OUT_DIR
    CITY, DISTRICT, OUT_DIR = city, district, out_dir
    load_key()
    os.makedirs(out_dir, exist_ok=True)

    stores = config.load_input(input_path)
    if not stores:
        raise SystemExit("输入模版里没有读到任何门店，请检查文件内容。")

    cache = load_cache(out_dir)
    print(f"清单 {len(stores)} 家，城市={CITY}，区={DISTRICT or '（不限）'}")

    results = []
    stat = {"exact": 0, "base": 0, "ambiguous": 0, "no_match": 0, "error": 0}

    for i, raw in enumerate(stores, 1):
        clean, embedded_note = strip_notes(raw["店名"])
        note = merge_note(embedded_note, raw["备注"])
        keyword = clean_keyword(clean)

        if clean in cache and cache[clean].get("状态") in ("exact", "base"):
            row = dict(cache[clean])
            row["关键词"] = keyword
            row["备注"] = note
            row["拜访频次"] = raw["拜访频次"]
            row["地址"] = raw["地址"]
        else:
            try:
                j = poi_search(keyword)
                pois = j.get("pois", []) or []
            except Exception as e:
                print(f"  [{i}/{len(stores)}] {clean} 请求失败({type(e).__name__})")
                pois = None

            if pois is None:
                row = {"店名": clean, "备注": note, "拜访频次": raw["拜访频次"], "地址": raw["地址"],
                       "关键词": keyword, "状态": "error",
                       "经度": "", "纬度": "", "POI名称": "", "POI地址": "", "POI_ID": ""}
            else:
                status, poi = match(keyword, pois)
                if poi:
                    loc = s(poi.get("location"))
                    if "," in loc:
                        lng, lat = loc.split(",")
                    else:
                        lng, lat, status = "", "", "no_match"
                    row = {"店名": clean, "备注": note, "拜访频次": raw["拜访频次"], "地址": raw["地址"],
                           "关键词": keyword, "状态": status,
                           "经度": lng, "纬度": lat,
                           "POI名称": s(poi.get("name")),
                           "POI地址": s(poi.get("address")),
                           "POI_ID": s(poi.get("id"))}
                else:
                    row = {"店名": clean, "备注": note, "拜访频次": raw["拜访频次"], "地址": raw["地址"],
                           "关键词": keyword, "状态": status,
                           "经度": "", "纬度": "", "POI名称": "", "POI地址": "", "POI_ID": ""}
            time.sleep(SLEEP)

        results.append(row)
        stat[row["状态"]] = stat.get(row["状态"], 0) + 1
        if i % 20 == 0 or i == len(stores):
            print(f"  [{i}/{len(stores)}] exact={stat.get('exact', 0)} "
                  f"base={stat.get('base', 0)} ambiguous={stat.get('ambiguous', 0)} "
                  f"no_match={stat.get('no_match', 0)} error={stat.get('error', 0)}")

    save_cache(out_dir, results)

    manual = [r for r in results if r["状态"] in ("no_match", "ambiguous", "error")]
    print(f"\n为 {len(manual)} 家待人工店找候选 POI …")
    for r in manual:
        cand = suggest(r["关键词"])
        r["候选"] = cand
        time.sleep(SLEEP)

    _write_manual_list(out_dir, manual)

    print("\n=== 汇总 ===")
    print(f"总数 {len(results)}：exact {stat.get('exact', 0)} / base {stat.get('base', 0)} "
          f"/ ambiguous {stat.get('ambiguous', 0)} / no_match {stat.get('no_match', 0)} "
          f"/ error {stat.get('error', 0)}")
    if stat.get("exact", 0) + stat.get("base", 0) == 0 and results:
        print("⚠ 没有一家店精确命中，很可能城市/区填错了，请核对 --city / --district。")
    if manual:
        print(f"\n待人工标注 {len(manual)} 家（见 待人工标注清单.xlsx）：")
        for r in manual:
            print(f"  - {r['店名']} [{r['状态']}]")
    return len(manual)


def _write_manual_list(out_dir, manual):
    headers = ["店名", "拜访频次", "状态", "原地址", "候选POI", "候选地址",
               "候选经度", "候选纬度", "备注"]
    with open(os.path.join(out_dir, "待人工标注清单.csv"), "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(headers)
        for r in manual:
            c = r.get("候选") or {}
            w.writerow([r["店名"], r["拜访频次"], r["状态"], r["地址"],
                        c.get("名称", ""), c.get("地址", ""), c.get("经度", ""),
                        c.get("纬度", ""), r.get("备注", "")])

    wb = Workbook()
    ws = wb.active
    ws.title = "待人工标注"
    ws.append(headers)
    for r in manual:
        c = r.get("候选") or {}
        ws.append([r["店名"], r["拜访频次"], r["状态"], r["地址"],
                   c.get("名称", ""), c.get("地址", ""), c.get("经度", ""),
                   c.get("纬度", ""), r.get("备注", "")])
    wb.save(os.path.join(out_dir, "待人工标注清单.xlsx"))


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 4:
        raise SystemExit("用法：python3 poi_geocode.py <input> <out_dir> <city> [district]")
    main(sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4] if len(sys.argv) > 4 else "")
