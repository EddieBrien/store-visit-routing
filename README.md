# store-visit-routing · 门店拜访排线 skill

把一份门店清单（店名 + 详细地址 + 可选拜访频次）变成「精确坐标 + 5 天均衡分圈 + 带导航的排线表」。跑店从「凭经验找店」变成「照排线点导航」。

## 一、安装（同事拿到这份仓库后）

任选其一：

```bash
# A. 直接 clone 到 skills 目录
git clone <仓库地址> ~/.claude/skills/store-visit-routing

# B. clone 到别处，再软链
git clone <仓库地址> ~/store-visit-routing
ln -s ~/store-visit-routing ~/.claude/skills/store-visit-routing
```

装好后，在 Claude Code 里说「帮我排线」或「门店分圈」，并给出模版文件，即可触发。

## 二、准备

1. **装依赖**（只依赖 openpyxl）：
   ```bash
   pip3 install -r requirements.txt
   ```
2. **申请高德 key**：lbs.amap.com → 控制台 → 创建应用 → 添加 Key → 服务平台选 **「Web 服务」**（这是服务端 key，别提交到公开仓库、别写进前端页面）。
3. **准备输入模版**：见 `references/input-template.md`（`店名` + `详细地址` + 可选 `拜访频次`）。

## 三、命令行跑法（不通过 Claude Code 也能用）

```bash
export AMAP_KEY=你的key
python3 scripts/run.py --input 我的门店.xlsx --city 大连市 --district 普兰店区 --outdir output
```

第一步标点后若有搜不到的门店，会停下来让你编辑 `output/待人工标注清单.xlsx` 的「状态」列，然后：

```bash
python3 scripts/run.py --input 我的门店.xlsx --city 大连市 --district 普兰店区 --outdir output --stage backfill
```

## 四、产出三件套（在 `--outdir` 目录）

- `跑店排线.xlsx` —— 排线表（点店名跳高德导航）+ 每日家数周次矩阵
- `导入模板_已填.xlsx` —— 导入高德云图 / 小程序，5 色分天
- `路线.kml` —— 5 个文件夹、每家一个点

## 五、注意事项

- 坐标体系是 **GCJ-02**（高德原生），全程不转 WGS-84。
- 高德 key 是个人密钥，只在你自己机器上通过环境变量使用；**不要**把它写进模版 Excel、不要 commit、不要发群里。
- `assets/导入模板.xlsx` 是高德官方模板，脚本只读它、输出到 `--outdir`，不会覆盖它。
