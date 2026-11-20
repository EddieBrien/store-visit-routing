# -*- coding: utf-8 -*-
"""排线编排器（唯一入口）。

geocode（POI 标点）→ [人工回填 checkpoint] → backfill（回填）→ route（分圈排线产出三件套）。

用法：
  python3 scripts/run.py --input <xlsx|csv> --city <市> [--district <区>]
        [--outdir <目录=./output>] [--amap-key <key>] [--stage all|geocode|backfill|route]
"""

import argparse
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)


def load_env_file(path):
    """极简 .env 解析（KEY=VALUE 行），不引 python-dotenv。"""
    if not os.path.exists(path):
        return {}
    d = {}
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            d[k.strip()] = v.strip().strip('"').strip("'")
    return d


def resolve_key(cli_key):
    if cli_key:
        return cli_key.strip()
    env = os.environ.get("AMAP_KEY", "").strip()
    if env:
        return env
    d = load_env_file(os.path.join(HERE, "..", ".env"))
    return d.get("AMAP_KEY", "").strip()


def _write_marker(marker, val):
    with open(marker, "w", encoding="utf-8") as f:
        f.write(val)


def _read_marker(marker):
    if os.path.exists(marker):
        return open(marker, encoding="utf-8").read().strip()
    return ""


def main():
    ap = argparse.ArgumentParser(description="门店拜访排线")
    ap.add_argument("--input", required=True, help="输入模版路径（xlsx 或 csv）")
    ap.add_argument("--city", required=True, help="城市，如 大连市")
    ap.add_argument("--district", default="", help="行政区，如 普兰店区；可留空（不做区界过滤）")
    ap.add_argument("--outdir", default="output", help="输出目录（默认 ./output）")
    ap.add_argument("--amap-key", default="", help="高德 Web 服务 key（优先于环境变量）")
    ap.add_argument("--stage", default="all", choices=["all", "geocode", "backfill", "route"])
    args = ap.parse_args()

    import poi_geocode as pg
    import fill_manual as fm
    import pai_xian as px

    outdir = os.path.abspath(args.outdir)
    os.makedirs(outdir, exist_ok=True)
    marker = os.path.join(outdir, ".pipeline_stage")

    stage = args.stage
    if stage == "all":
        done = _read_marker(marker)
        if done == "geocoded":
            stage = "backfill"
        elif done == "backfilled":
            stage = "route"
        elif done == "done":
            print("上次已全部完成（done）。如需重跑，请删掉 .pipeline_stage 或显式指定 --stage。")
            return
        else:
            stage = "geocode"

    # 仅 geocode / backfill 需要高德 key（route 纯本地计算）
    if stage in ("geocode", "backfill"):
        key = resolve_key(args.amap_key)
        if not key:
            print("错误：缺少高德 key。请 export AMAP_KEY=xxx 或传 --amap-key xxx")
            sys.exit(1)
        os.environ["AMAP_KEY"] = key

    if stage == "geocode":
        n_manual = pg.main(args.input, outdir, args.city, args.district)
        _write_marker(marker, "geocoded")
        if n_manual > 0:
            print(f"\n有 {n_manual} 家待人工标注。请编辑：{os.path.join(outdir, '待人工标注清单.xlsx')}")
            print("在第 3 列「状态」填正确地址（按地址编码）或「先空着」（留空），保存后重跑：")
            print(f"  python3 scripts/run.py --input \"{args.input}\" --city \"{args.city}\""
                  f"{(' --district ' + args.district) if args.district else ''} --stage backfill")
            return
        stage = "route"

    if stage == "backfill":
        fm.main(args.input, outdir, args.city, args.district)
        _write_marker(marker, "backfilled")
        stage = "route"

    if stage == "route":
        px.route_stage(outdir)
        _write_marker(marker, "done")
        print("\n全部完成。三件套已输出到：", outdir)


if __name__ == "__main__":
    main()
