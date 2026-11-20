# store-visit-routing · 门店拜访排线 skill

一句话：把一份门店清单变成「精确坐标 + 5 天均衡分圈 + 带导航的排线表」。跑店从「凭经验找店」变成「照排线点导航」。

---

## 快速上手

### 你需要准备两样东西

**① 门店清单**（Excel 或 CSV，三列）

| 店名 | 详细地址 | 拜访频次 |
|---|---|---|
| 金鹏超市(世纪路东段) | 普兰店区世纪路东段32号 | 两周一次 |

- 店名、详细地址必填
- 拜访频次选填，只写三种：**一周一次 / 两周一次 / 四周一次**（不填默认一周一次）

**② 高德 key（免费）**

- 打开 lbs.amap.com → 控制台 → 创建应用 → 添加 Key → 服务平台选 **「Web 服务」**
- ⚠️ key 只自己用，别发群里、别写进 Excel

### 怎么跑

第一次先装一下（一次性，前提：已装 Claude Code）：

```bash
git clone https://github.com/EddieBrien/store-visit-routing.git ~/.claude/skills/store-visit-routing
pip3 install -r ~/.claude/skills/store-visit-routing/requirements.txt
```

然后跟 Claude Code 说一句 **「帮我排线」**。它会问你要三样：门店清单文件、城市和区（如 `大连市` / `普兰店区`）、高德 key。

### 你会拿到三样

- `跑店排线.xlsx` —— 周一到周五每天走哪些店、顺序、第几周，点店名直接跳高德导航
- `导入模板_已填.xlsx` —— 导进高德云图 / 小程序，5 种颜色分天
- `路线.kml` —— 地图路线文件

### 两个注意

- 有搜不到坐标的店时，脚本会停下来给你一个清单：在「状态」列填对地址（或写「先空着」）再继续即可。
- 换一批店、换一个片区，重跑一次就有新线，方法不用改。

---

## 进阶：命令行直跑（不通过 Claude Code 也能用）

```bash
export AMAP_KEY=你的key
python3 scripts/run.py --input 我的门店.xlsx --city 大连市 --district 普兰店区 --outdir output
```

第一步标点后若有搜不到的门店，会停下来让你编辑 `output/待人工标注清单.xlsx` 的「状态」列，然后：

```bash
python3 scripts/run.py --input 我的门店.xlsx --city 大连市 --district 普兰店区 --outdir output --stage backfill
```

输入模版列名 / 别名 / 人工回填怎么填，见 `references/input-template.md`。

---

## 注意事项（技术）

- 坐标体系是 **GCJ-02**（高德原生），全程不转 WGS-84。
- 高德 key 是个人密钥，只在你自己机器上通过环境变量使用；**不要**把它写进模版 Excel、不要 commit、不要发群里。
- `assets/导入模板.xlsx` 是高德官方模板，脚本只读它、输出到 `--outdir`，不会覆盖它。
