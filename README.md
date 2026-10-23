# 全天候策略项目

研究股票、黄金、其他商品和国债的风险预算配置，构建期货版全天候策略。

## 当前阶段

- 已实现：同花顺 iFinD 加权期货收盘价下载、历史风险估计。
- 已准备：品种信息及9套参数下的逐期调仓安排。
- 待实现：组合优化、独立回测模块。
- `main.py` 目前仍为 PyCharm 示例，尚未接入统一运行流程。

## 项目目录

```text
全天候策略项目/
├── main.py
├── scripts/
│   ├── __init__.py
│   ├── 1.数据下载/
│   │   ├── __init__.py
│   │   └── 1.下载期货数据.py
│   ├── 2.组合优化/
│   │   ├── __init__.py
│   │   ├── 1.风险估计.py
│   │   └── 2.组合优化.py       # 目前为空
│   └── 3.回测/
│       └── __init__.py        # 预留
└── data/
    ├── futures/
    │   ├── stocks/           # IF、IC、IM
    │   ├── bonds/            # TS、TF、T、TL
    │   ├── commodities/      # AU、SC、CU、RB、M
    │   ├── close_price.csv
    │   └── download_manifest.csv
    └── prep/
        └── 期货回测安排.xlsx
```

## 环境与依赖

使用 Python 3.9 或以上版本。以下命令均在项目根目录执行。

风险估计使用 NumPy、pandas，读取调仓 xlsx 还需要 openpyxl：

```bash
./.venv/bin/python -m pip install numpy pandas openpyxl
```

下载程序仅使用 Python 标准库，不依赖以上三个包。

## 数据与调仓安排

- 数据来源：同花顺 iFinD HTTP API，加权期货日收盘价。
- 现有行情总表区间：2013-09-06至2026-09-30；未上市品种从实际可用日期开始。
- 回测设定从2015-01-01开始；此前行情只用于历史风险估计。
- 首个信号日：2014-12-31；首次执行日：2015-01-05收盘。
- 回看窗口：63、126、252个日收益，分别需要64、127、253个收盘价。
- 调仓频率：每月、每季度、每半年，共9套安排。
- `data/prep/期货回测安排.xlsx` 的“品种数据”记录名称与代码等信息，“调仓安排”记录每期信号日、执行日、参与品种和类内比例。
- 同类品种按名义敞口等分。新品种何时纳入直接沿用现有安排，风险估计程序不重新制定纳入规则。

行情总表保留上市前空值，不填零，不向前或向后填充。沪铜上市时间仅精确到 `1993-03`，不将月初当作精确上市日。

## 运行风险估计

默认计算2014-12-31、63日窗口、每月方案：

```bash
./.venv/bin/python "scripts/2.组合优化/1.风险估计.py"
```

例如，查看同一信号日的252日窗口、每季度方案：

```bash
./.venv/bin/python "scripts/2.组合优化/1.风险估计.py" \
  --signal-date 2014-12-31 --lookback 252 --frequency 每季度
```

| 参数 | 默认值或要求 |
|---|---|
| `--prices` | `data/futures/close_price.csv` |
| `--schedule` | `data/prep/期货回测安排.xlsx` |
| `--signal-date` | `2014-12-31`，须为行情中的交易日 |
| `--lookback` | `63`，表示收益观测数 |
| `--frequency` | `每月`；另可选 `每季度`、`每半年` |

日期、窗口和频率的组合必须在现有调仓表中有唯一匹配记录。程序打印样本价格区间、四类资产年化协方差矩阵和年化波动率，运行时只读输入，不生成结果文件。

计算口径：

1. 按当期名单选择品种，截取截至信号日的最近 L+1 个共同价格。
2. 计算简单日收益：`当日收盘价 / 上日收盘价 - 1`。
3. 整个窗口固定使用当期名单及类内等分比例，合成四类资产日收益。
4. 固定类别顺序为“股票、黄金、其他商品、国债”，以 `252 × 样本协方差（ddof=1）` 年化。
5. 检查共同窗口完整性、有限正价格、比例及矩阵有效性。允许奇异半正定矩阵，要求国债方差为正；不自动填补数据、延长窗口凑数或修补矩阵。

对外函数：

| 函数 | 用途 |
|---|---|
| `load_prices(path)` | 读取日期索引的收盘价表，保留上市前空值 |
| `load_schedule(path)` | 读取现有xlsx，返回每期日期、品种代码及类内比例 |
| `estimate_risk(prices, signal_date, lookback, allocations)` | 计算单个信号日、单个窗口的风险，返回结果字典 |

`allocations` 格式为 `{类别: {完整行情代码: 类内比例}}`。结果包含带类别标签的协方差、波动率、类别日收益、品种比例、样本日期及数量、矩阵诊断。

行情与安排可各读一次，再逐期调用计算函数。调仓频率只负责选择安排；同一日期、窗口、名单和比例的风险估计可复用。风险预算、波动率参数和敞口约束属于后续组合优化，不参与风险估计。

当前目录与文件名带数字及点号，尚未接入常规包导入；现阶段使用上述直接执行方式，后续统一入口再处理导入。

## 风险预算

按“股票、黄金、其他商品、国债”的顺序，风险预算为 `1/3、1/6、1/6、1/3`。类内等分的是名义敞口，类别间的敞口须由后续组合优化结合协方差求解。

## 下载行情

当前脚本配置为2013-09-06至2026-09-30。先离线查看品种配置，此命令不读取凭证、不联网：

```bash
./.venv/bin/python "scripts/1.数据下载/1.下载期货数据.py" --list
```

下载脚本移动后，内部项目根目录定位尚未调整，当前默认查找 `scripts/.env`、输出到 `scripts/data/futures/`。以下命令使用环境变量提供令牌，并显式指定项目根目录下的输出位置；项目根目录的 `.env` 目前不会被该脚本自动读取。

在自己的终端配置实际令牌，替换占位文本：

```bash
export IFIND_REFRESH_TOKEN="your_refresh_token"
```

先核验各品种最早和最近各14个自然日的小样本：

```bash
./.venv/bin/python "scripts/1.数据下载/1.下载期货数据.py" \
  --probe --output data/futures
```

核验结果写入 `data/futures_probe/probe_report.csv` 和 `probe_samples.csv`，不覆盖正式行情。首次使用时，需对照 iFinD 终端核对加权序列及价格口径。

核验后下载完整区间：

```bash
./.venv/bin/python "scripts/1.数据下载/1.下载期货数据.py" \
  --start 2013-09-06 --end 2026-09-30 --output data/futures
```

`--start`、`--end` 可修改日期；将顶部 `END_DATE` 留空才默认上海时间的昨天。程序保留接口交易日期，不另行移动夜盘日期。

只下载部分品种时，显式使用单独目录，例如：

```bash
./.venv/bin/python "scripts/1.数据下载/1.下载期货数据.py" \
  --codes IF8888.CFE TF8888.CFE --output data/futures_subset
```

## 行情文件说明

完整下载生成14个 UTF-8-SIG 编码的CSV：12个分类存放的单品种文件，加上 `close_price.csv` 和 `download_manifest.csv`。当前程序不生成 `data_issues.csv`。

| 文件 | 内容 |
|---|---|
| `stocks/*.csv`、`bonds/*.csv`、`commodities/*.csv` | 日期、代码、名称、资产类别、收盘价；上市前及非交易日不补行 |
| `close_price.csv` | 日期×完整行情代码宽表；按接口返回日期并集对齐，无价格时为空 |
| `download_manifest.csv` | 品种信息、请求与实际区间、行数、状态、对齐缺口、非正价格数量、来源及下载时间等19列 |

完整总表的品种顺序为：

```text
IF8888.CFE、IC8888.CFE、IM8888.CFE、TS8888.CFE、TF8888.CFE、T8888.CFE、TL8888.CFE、AU8888.SHF、SC8888.SHF、CU8888.SHF、RB8888.SHF、M8888.DCE
```

部分下载也生成所选品种的总表和清单。行情列顺序与风险估计的四类顺序不同，风险估计按代码和类别标签匹配。

上市时间用于裁剪请求起点。下载程序检查较长空段、异常两端缺口、缺失价格及冲突重复日期；清单另列对齐缺口和非正价格数量。非正价格会保留，风险估计使用到它们时会报错。

所有行情请求成功后才开始保存。每个文件使用临时文件再替换，整个目录不是一次性原子替换。加权价格是供应商合成序列，研究收益不直接等同于可交易合约组合的实际收益。

## 凭证与版本管理

下载 HTTP 接口只使用 `IFIND_REFRESH_TOKEN`，环境变量优先；`IFIND_USERNAME` 和 `IFIND_PASSWORD` 不是当前 HTTP 下载的必填项。真实凭证不写入代码、README、日志或提交记录。

`.env`、虚拟环境、PyCharm配置、缓存、运行结果及日志保留在本机。当前Git已跟踪12个品种CSV、行情总表、下载清单及准备xlsx。

`data/` 仍在 `.gitignore` 中：已跟踪文件的修改可以正常提交，新增未跟踪数据文件会被忽略；需要纳入新数据文件时，明确指定该文件使用 `git add -f -- 路径`。

## 已有验证

风险估计已用现有642期调仓安排逐期对照独立计算结果，并通过缺失数据、非法日期、错误分配、未来数据隔离及奇异矩阵等17项检查。组合优化及回测尚未实现，当前没有净值或绩效结果。

## 口径参考

- [iFinD HTTP 部署与鉴权](https://quantapi.51ifind.com/gwstatic/static/ds_web/quantapi-web/help-center/deploy.html)
- [iFinD 历史行情接口手册](https://quantapi.51ifind.com/gwstatic/static/ds_web/quantapi-web/help-center/manual.html)
- [iFinD 权限说明](https://quantapi.51ifind.com/gwstatic/static/ds_web/quantapi-web/help-center.html)
- [中金所上市品种资料](https://www.cffex.com.cn/u/cms/www/202405/31111942nrka.pdf)
- [山金期货原始研报](https://shanjinqh.com/uploadfiles/2024/04/20240401173444403.pdf)
- [上期所《铜铝期货三十年》](https://www.shfe.com.cn/upload/20231027/1698388688886.pdf)
