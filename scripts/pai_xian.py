# -*- coding: utf-8 -*-
"""排线阶段：读统一 poi_cache.csv → 均衡分 5 圈（周一~周五）→ 组内最近邻 →
按拜访频次分配第几周 → 产出三件套（跑店排线.xlsx / 导入模板_已填.xlsx / 路线.kml）。

坐标体系全程 GCJ-02，不转 WGS-84。
"""

import csv
import html
import math
import os

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font

import poi_geocode as pg

BASE = os.path.dirname(os.path.abspath(__file__))
K = 5
DAYS = ["周一", "周二", "周三", "周四", "周五"]
COLORS = ["#7E97FF", "#12E3AA", "#FFBF00", "#FF983F", "#FF6968"]
OK = ("exact", "base", "geocode")


# ---------------- 分组算法（本地计算，无高德调用） ----------------
def project(pts):
    lat0 = sum(p["lat"] for p in pts) / len(pts)
    cos = math.cos(math.radians(lat0))
    for p in pts:
        p["x"] = p["lng"] * 111320.0 * cos
        p["y"] = p["lat"] * 110540.0
    return pts


def d2(a, b):
    return (a["x"] - b["x"]) ** 2 + (a["y"] - b["y"]) ** 2


def init_centroids(pts, k):
    idx = [0]
    while len(idx) < k:
        best_i, best_d = -1, -1.0
        for i, p in enumerate(pts):
            d = min(d2(p, pts[j]) for j in idx)
            if d > best_d:
                best_d, best_i = d, i
        idx.append(best_i)
    return [(pts[i]["x"], pts[i]["y"]) for i in idx]


def min_cost_assignment(cost, caps):
    n, k = len(cost), len(caps)
    S, T = 0, n + k + 1

    def pid(i):
        return 1 + i

    def cid(j):
        return 1 + n + j

    N = n + k + 2
    graph = [[] for _ in range(N)]

    def add(u, v, cap, c):
        graph[u].append([v, cap, c, len(graph[v])])
        graph[v].append([u, 0, -c, len(graph[u]) - 1])

    for i in range(n):
        add(S, pid(i), 1, 0)
        for j in range(k):
            add(pid(i), cid(j), 1, cost[i][j])
    for j in range(k):
        add(cid(j), T, caps[j], 0)

    flow = 0
    while flow < n:
        dist = [float("inf")] * N
        inq = [False] * N
        prevv = [-1] * N
        preve = [-1] * N
        dist[S] = 0.0
        q = [S]
        inq[S] = True
        while q:
            u = q.pop(0)
            inq[u] = False
            for ei, e in enumerate(graph[u]):
                v, cap, c, _ = e
                if cap > 0 and dist[v] > dist[u] + c + 1e-9:
                    dist[v] = dist[u] + c
                    prevv[v] = u
                    preve[v] = ei
                    if not inq[v]:
                        q.append(v)
                        inq[v] = True
        if dist[T] == float("inf"):
            break
        d = float("inf")
        v = T
        while v != S:
            d = min(d, graph[prevv[v]][preve[v]][1])
            v = prevv[v]
        v = T
        while v != S:
            e = graph[prevv[v]][preve[v]]
            e[1] -= d
            graph[v][e[3]][1] += d
            v = prevv[v]
        flow += d

    assign = [-1] * n
    for i in range(n):
        u = pid(i)
        for e in graph[u]:
            v, cap, _, _ = e
            if n + 1 <= v <= n + k and cap == 0:
                assign[i] = v - n - 1
    return assign


def balanced_kmeans(pts, k, targets, iters=40):
    cents = init_centroids(pts, k)
    best = None
    for _ in range(iters):
        cost = [[int(round((pts[i]["x"] - cents[j][0]) ** 2 + (pts[i]["y"] - cents[j][1]) ** 2))
                 for j in range(k)] for i in range(len(pts))]
        assign = min_cost_assignment(cost, targets)
        clusters = [[] for _ in range(k)]
        for i, j in enumerate(assign):
            clusters[j].append(i)
        new = []
        for cl in clusters:
            new.append((sum(pts[i]["x"] for i in cl) / len(cl),
                        sum(pts[i]["y"] for i in cl) / len(cl)))
        tot = sum((pts[i]["x"] - new[j][0]) ** 2 + (pts[i]["y"] - new[j][1]) ** 2
                  for j, cl in enumerate(clusters) for i in cl)
        if best is None or tot < best[0]:
            best = (tot, clusters)
        if max(abs(new[j][0] - cents[j][0]) + abs(new[j][1] - cents[j][1])
               for j in range(k)) < 1e-6:
            break
        cents = new
    return best[1]


def order_route(pts, indices):
    remaining = set(indices)
    cur = min(remaining, key=lambda i: pts[i]["lng"])
    remaining.remove(cur)
    route = [cur]
    while remaining:
        nxt = min(remaining, key=lambda i: d2(pts[i], pts[cur]))
        route.append(nxt)
        remaining.remove(nxt)
        cur = nxt
    return route


def freq_weight(freq):
    """周访=4 两周=2 四周=1（一个月四周的工作量权重）。"""
    return {"一周一次": 4, "两周一次": 2, "四周一次": 1}.get(freq, 4)


def rebalance(pts, clusters):
    """把重区里的店挪到轻区，让每天工作量（按周加权）尽量均衡。"""
    for _ in range(500):
        w = [sum(freq_weight(pts[i]["freq"]) for i in cl) for cl in clusters]
        mx = max(range(K), key=lambda j: w[j])
        mn = min(range(K), key=lambda j: w[j])
        gap = w[mx] - w[mn]
        if gap <= 2:
            break
        cx = sum(pts[i]["x"] for i in clusters[mn]) / len(clusters[mn])
        cy = sum(pts[i]["y"] for i in clusters[mn]) / len(clusters[mn])
        cands = sorted(clusters[mx], key=lambda idx: (
            freq_weight(pts[idx]["freq"]),
            (pts[idx]["x"] - cx) ** 2 + (pts[idx]["y"] - cy) ** 2))
        moved = False
        for i in cands:
            if w[mx] - freq_weight(pts[i]["freq"]) >= w[mn]:
                clusters[mx].remove(i)
                clusters[mn].append(i)
                moved = True
                break
        if not moved:
            break
    return clusters


# ---------------- 排线主流程 ----------------
def load_poi_cache(out_dir):
    rows = []
    with open(os.path.join(out_dir, "poi_cache.csv"), encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            rows.append(r)
    return rows


def route_stage(out_dir):
    rows = load_poi_cache(out_dir)

    # 去重：同一店名只保留首行
    seen = set()
    uniq = []
    for r in rows:
        nm = (r.get("店名") or "").strip()
        if nm in seen:
            continue
        seen.add(nm)
        uniq.append(r)
    rows = uniq

    def _has_coord(r):
        return r.get("状态") in OK and (r.get("经度") or "").strip() and (r.get("纬度") or "").strip()

    placed = [r for r in rows if _has_coord(r)]
    blank = [r for r in rows if not _has_coord(r)]
    print(f"去重后 {len(rows)} 家：确定坐标 {len(placed)} 家，待人工 {len(blank)} 家")

    if len(placed) < K:
        print(f"错误：有坐标门店不足 {K} 家（当前 {len(placed)}），无法分 5 圈。请先人工补坐标后重跑。")
        return None, None, blank

    pts = []
    for r in placed:
        pts.append({
            "name": r["店名"],
            "freq": r.get("拜访频次") or "一周一次",
            "lng": float(r["经度"]),
            "lat": float(r["纬度"]),
            "addr": (r.get("POI地址") or "").strip() or (r.get("地址") or "").strip(),
        })
    project(pts)
    n = len(pts)

    base, rem = divmod(n, K)
    targets = [base + (1 if j < rem else 0) for j in range(K)]
    clusters = balanced_kmeans(pts, K, targets)
    clusters = rebalance(pts, clusters)

    # 按质心经度从西到东 → 周一~周五
    order = sorted(range(K), key=lambda j: sum(pts[i]["lng"] for i in clusters[j]) / len(clusters[j]))
    day_routes = [order_route(pts, clusters[j]) for j in order]
    for d, route in enumerate(day_routes):
        for k, i in enumerate(route, 1):
            pts[i]["day"] = DAYS[d]
            pts[i]["seq"] = k

    # 周次均衡分配
    cnt = {d: {w: 0 for w in range(1, 5)} for d in DAYS}
    for p in pts:
        if p["freq"] == "一周一次":
            for w in range(1, 5):
                cnt[p["day"]][w] += 1
            p["week"] = "每周"
    for p in pts:
        if p["freq"] == "两周一次":
            w1 = cnt[p["day"]][1] + cnt[p["day"]][3]
            w2 = cnt[p["day"]][2] + cnt[p["day"]][4]
            p["week"] = "第1周" if w1 <= w2 else "第2周"
            for w in (1, 3) if p["week"] == "第1周" else (2, 4):
                cnt[p["day"]][w] += 1
    for p in pts:
        if p["freq"] == "四周一次":
            best_w = min(range(1, 5), key=lambda w: cnt[p["day"]][w])
            p["week"] = f"第{best_w}周"
            cnt[p["day"]][best_w] += 1

    print("\n每天家数（按周）：")
    for d in DAYS:
        line = [f"第{w}周={cnt[d][w]}" for w in range(1, 5)]
        print(f"  {d}：{line}")

    write_template(out_dir, day_routes, pts, blank)
    write_kml(out_dir, day_routes, pts, blank)
    write_schedule(out_dir, day_routes, pts, blank, cnt)
    write_manual_list(out_dir, blank)
    return day_routes, pts, blank


# ---------------- 产出三件套 ----------------
def write_template(out_dir, day_routes, pts, blank):
    src = openpyxl.load_workbook(pg.ASSET_TEMPLATE)
    ws = src["标记位置"]
    ws.delete_rows(2, ws.max_row - 1)
    for d, route in enumerate(day_routes):
        color, folder = COLORS[d], DAYS[d]
        for i in route:
            p = pts[i]
            ws.append([p["name"], p["lng"], p["lat"], p["addr"], color, None, None, None, folder])
    for b in blank:
        ws.append([b["店名"], "", "", b.get("POI地址") or b.get("地址") or "", None, None, None, None, None])
    src.save(os.path.join(out_dir, "导入模板_已填.xlsx"))
    print(f"  已写 导入模板_已填.xlsx（5 色分圈，{len(pts)} 家带坐标）")


def write_kml(out_dir, day_routes, pts, blank):
    lines = ['<?xml version="1.0" encoding="UTF-8"?>',
             '<kml xmlns="http://www.opengis.net/kml/2.2">',
             '<Document><name>跑店门店（5天分组·含频次）</name>']
    for d, route in enumerate(day_routes):
        lines.append(f'<Folder><name>{DAYS[d]}（{len(route)}家）</name>')
        for k, i in enumerate(route, 1):
            p = pts[i]
            lines.append(f'<Placemark><name>{k}. {html.escape(p["name"])}</name>')
            lines.append(f'<Point><coordinates>{p["lng"]},{p["lat"]},0</coordinates></Point></Placemark>')
        lines.append('</Folder>')
    if blank:
        lines.append('<Folder><name>待补坐标</name>')
        for b in blank:
            lines.append(f'<Placemark><name>{html.escape(b["店名"])}</name></Placemark>')
        lines.append('</Folder>')
    lines += ['</Document>', '</kml>']
    with open(os.path.join(out_dir, "路线.kml"), "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print("  已写 路线.kml")


def write_schedule(out_dir, day_routes, pts, blank, cnt):
    wb = Workbook()
    ws = wb.active
    ws.title = "排线表"
    headers = ["星期几", "序号", "店名", "拜访频次", "第几周", "地址"]
    for c, h in enumerate(headers, 1):
        ws.cell(1, c, h).font = Font(bold=True)

    link_font = Font(color="0071E3", underline="single")
    r = 2
    for d, route in enumerate(day_routes):
        for i in route:
            p = pts[i]
            ws.cell(r, 1, p["day"])
            ws.cell(r, 2, p["seq"])
            nc = ws.cell(r, 3, p["name"])
            nc.hyperlink = pg.nav_url(p["name"], p["lng"], p["lat"])
            nc.font = link_font
            ws.cell(r, 4, p["freq"])
            ws.cell(r, 5, p["week"] if p["week"] else "每周")
            ws.cell(r, 6, p["addr"])
            r += 1
    for b in blank:
        ws.cell(r, 1, "待补坐标")
        ws.cell(r, 3, b["店名"])
        ws.cell(r, 4, b.get("拜访频次") or "")
        ws.cell(r, 6, b.get("POI地址") or b.get("地址") or "")
        r += 1

    for col, w in zip("ABCDEF", (8, 6, 34, 12, 10, 46)):
        ws.column_dimensions[col].width = w

    ws2 = wb.create_sheet("每日家数")
    ws2.cell(1, 1, "星期几").font = Font(bold=True)
    for w in range(1, 5):
        ws2.cell(1, 1 + w, f"第{w}周").font = Font(bold=True)
    for ri, d in enumerate(DAYS, 2):
        ws2.cell(ri, 1, d)
        for w in range(1, 5):
            ws2.cell(ri, 1 + w, cnt[d][w])
    ws2.cell(2 + K, 1, "合计")
    for w in range(1, 5):
        ws2.cell(2 + K, 1 + w, sum(cnt[d][w] for d in DAYS))

    wb.save(os.path.join(out_dir, "跑店排线.xlsx"))
    print("  已写 跑店排线.xlsx（周一到周五 + 频次 + 第几周 + 每日家数）")


def write_manual_list(out_dir, blank):
    wb = Workbook()
    ws = wb.active
    ws.title = "待补坐标"
    headers = ["店名", "拜访频次", "状态", "候选POI", "候选经度", "候选纬度", "说明"]
    for c, h in enumerate(headers, 1):
        ws.cell(1, c, h).font = Font(bold=True)
    for b in blank:
        ws.append([b["店名"], b.get("拜访频次") or "", b.get("状态") or "",
                   b.get("POI名称") or None, b.get("经度") or None, b.get("纬度") or None,
                   "高德搜不到，请补地址/坐标"])
    for col, w in zip("ABCDEFG", (34, 12, 10, 22, 12, 12, 26)):
        ws.column_dimensions[col].width = w
    wb.save(os.path.join(out_dir, "待补坐标清单.xlsx"))
    print(f"  已写 待补坐标清单.xlsx（{len(blank)} 家待补）")


if __name__ == "__main__":
    import sys
    out = sys.argv[1] if len(sys.argv) > 1 else "output"
    route_stage(out)
