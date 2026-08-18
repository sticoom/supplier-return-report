# 每月供应商质量退货金额统计工具

把品质部每月手工统计各供应商质量退货金额的流程（Excel 透视 + VLOOKUP，约 17 步）自动化：
上传 4 个月度导出文件 → 自动计算 → 产出与手工版式一致的结算报告 xlsx + SQLite 历史累计，
可通过命令行 CLI 或 FastAPI 网页使用。

计算规则以 [`dev-pipeline/supplier-return-report/requirement.md`](dev-pipeline/supplier-return-report/requirement.md) 第 3 节为准（已逐条拍板，不得改动）。

## 功能

- 读 4 个月度文件：退货(FBA)订单导出、退货(FBM)订单导出、DLM退货统计SKU导出、采购入库单
- 质量退货筛选（DEFECTIVE / MISSING_PARTS / QUALITY_UNACCEPTABLE，FBM 带 `CR-` 前缀）按 SKU×原因透视
- 不分一供/二供：按上传采购入库单（≤报告月月末）各供应商交货数量占比分摊质量退货费用；无交货数据时归 DLM 默认供应商并标注人工复核
- 单价取「报告月月末及之前最近一次入库的不含税单价」，缺价 SKU 金额留空并进「数据校验」
- 考核系数：V3/V4 版协议按批次合格率分档（×0 ~ ×1.2）；未签/V2 = 1.0；本月无验货 = 0
- 应扣金额 < 200（乘系数前判断）的供应商进「低于200清单」，按季度累计、新季度重新起算
- 产出单个 xlsx：供应商费用明细 / 各供应商结算清单（复刻手工版式）/ 低于200清单 / 季度累计 / 买家备注复核清单 / 供应商批次合格率矩阵 / 数据校验

## 安装

Python 3.10+，在项目根目录：

```bash
pip install -r requirements-dev.txt   # 开发（含 pytest/httpx）
pip install -r requirements.txt       # 仅运行
```

## 测试

```bash
python -m pytest -v                   # 单测 + 合成数据端到端（默认不含真实数据集成测试）
python -m pytest -m integration -v    # 真实数据集成测试（须本机存在真实样例文件，见下）
```

- 所有常规测试的 fixture 均为 openpyxl 代码生成的小型合成 xlsx/csv（`tests/fixtures_build.py`），
  不提交任何二进制数据文件。
- `tests/test_integration_real.py` 分两层：
  - 合成全量 E2E（默认跑）：真实 CLI 子进程全量跑合成 7 月数据出报告，再做「人工抽查 3 个供应商」
    的自动化等价——验货率与验货文件独立重算交叉校验、SKU 件数×单价=金额、交货拆分行守恒。
  - 真实数据 E2E（`marker=integration`）：读 `D:/Downloads` 7 月真实文件 + 本地参考数据。
    `pyproject.toml` 的 `addopts = -m "not integration"` 保证默认运行自动跳过（避免普通回归
    触碰 123MB 验货报表），真实文件缺失时也会 skip。
- 真实数据集成测试依赖的本机文件（只读，不入库）：
  `D:/Downloads/退货(FBA)订单导出-*.xlsx`、`D:/Downloads/退货(FBM)订单导出-*.xlsx`、
  `D:/Downloads/DLM退货统计SKU导出-*.xlsx`、`C:/Users/13676/Desktop/飞书下载/采购入库单_*.xlsx`、
  `data/2026年验货数据报表.xlsx`（sheet「26年验货原始数据」）、`data/供应商协议签订记录.csv`

## CLI 运行（全量）

```bash
python -m engine.pipeline --month 2026-07 \
  --fba     "D:/Downloads/退货(FBA)订单导出-947873100663549952.xlsx" \
  --fbm     "D:/Downloads/退货(FBM)订单导出-947873572875407360.xlsx" \
  --dlm     "D:/Downloads/DLM退货统计SKU导出-2026-08-17.xlsx" \
  --inbound "C:/Users/13676/Desktop/飞书下载/采购入库单_202601-07.xlsx" \
  --inspection "data/2026年验货数据报表.xlsx" \
  --agreements "data/供应商协议签订记录.csv" \
  --out data/reports
```

- `--inspection` / `--agreements` 可选：传入即更新参考库（SQLite，立即生效，只替换传入的部分）
- `--db` 默认 `data/app.db`，`--out` 默认 `data/reports`
- 成功后 stdout 输出 RunSummary JSON（供应商数 / 低于200数 / 复核条数 / 校验条数 / 报告路径）
- 生成真实报告后建议人工抽查 3 个供应商：① 有验货的供应商对「当月检验合格率」= 验货文件
  合格批数÷总批数；② 任一 SKU 行 件数×单价=金额，件数与 FBA/FBM 原始订单筛选求和一致；
  ③ 交货拆分行 note=「按交货比例分摊」且件数为小数、两行件数之和等于原始件数
  （`tests/test_integration_real.py` 已把这三步做成自动化断言，真实数据跑 `-m integration` 即等价执行）

## Web 本地运行

```bash
python -m uvicorn web.main:app --port 8300
```

浏览器打开 <http://127.0.0.1:8300/>。环境变量 `SRR_DATA_DIR` 可指定数据根目录（默认项目 `data/`）。

## 部署到阿里云 ECS

目标：`root@120.25.100.51:/opt/supplier-return-report`，uvicorn 监听 `0.0.0.0:8300`，
**参照服务器 `/opt/fund-dashboard` 既有 systemd 模式**（`systemctl` 管服务，不用 nohup/pkill）。
部署属「修改生产配置」：**以下触碰服务器的命令，执行前必须征得用户同意**（dev-pipeline 部署门）。

部署产物由 `deploy/service.py` 单一来源生成（`tests/test_deploy.py` 守护）：

```bash
python -m deploy.service unit   # 打印 systemd unit 文件内容
python -m deploy.service plan   # 打印部署步骤命令清单（同步 + 安装）
```

首次部署（每条命令执行前征得用户同意）：

```bash
# 1) 同步代码（data/ 与 tests/ 不上传；参考库在服务器网页「更新参考库」或 CLI 重新导入）
ssh root@120.25.100.51 "mkdir -p /opt/supplier-return-report"
scp -r engine web deploy root@120.25.100.51:/opt/supplier-return-report/
scp requirements.txt pyproject.toml root@120.25.100.51:/opt/supplier-return-report/

# 2) 服务器上装依赖、落盘 unit、systemd 接管（在 /opt/supplier-return-report 内执行）
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/python -m deploy.service unit > /etc/systemd/system/supplier-return.service
systemctl daemon-reload
systemctl enable --now supplier-return
```

验收与运维：

- 网页验收：<http://120.25.100.51:8300/>（选月 → 上传 → 生成 → 下载）；
  命令行探活 `python -c "from deploy.service import wait_healthy; print(wait_healthy('http://120.25.100.51:8300/'))"`
- 安全组需放行 8300 端口（参照 fund-dashboard 的 8000 放行方式，在阿里云控制台操作）
- 日志：`journalctl -u supplier-return -f`
- 更新代码：重新 `scp` 后 `ssh root@120.25.100.51 "cd /opt/supplier-return-report && .venv/bin/pip install -r requirements.txt && systemctl restart supplier-return"`
- 服务器数据目录：`/opt/supplier-return-report/data/`（uploads/reports/reference/app.db，同本地布局）

## 每月操作步骤（网页）

1. 选择报告月（YYYY-MM）
2. 上传 4 个月度文件（FBA / FBM / DLM / 采购入库单，按文件名关键字自动识别类型）
3. 可选：「更新参考库」上传新版验货数据报表（读「26年验货原始数据」sheet）或协议签订记录（csv/xlsx），立即生效
4. 点「生成报告」→ 展示校验摘要（供应商数 / 缺价 SKU 数 / 需复核条数）→ 下载 xlsx
5. 历史页可按月重新下载；「本季度 <200 累计」视图查看季度累计

## 数据目录说明

| 路径 | 内容 |
|---|---|
| `data/uploads/{YYYY-MM}/` | 该月上传的原始月度文件 |
| `data/reports/{YYYY-MM}/` | 产出的 `YYYY年M月供应商质量退货金额汇总表.xlsx` |
| `data/ref/` | 网页上传的参考库引导文件（验货 / 协议） |
| `data/app.db` | SQLite：月度供应商结果、SKU 明细、参考数据表、上传记录 |

`data/` 整体被 .gitignore 排除。同月重跑为替换语义（upsert），不会重复累计。

## 季度累计规则

- 按应扣金额（乘系数前）判断：< 200 进「低于200清单」，≥ 200 进正式月度结算清单
- 每次上传一个月数据，该月 <200 供应商的应扣/应承担金额即累加进季度累计表
- 累计按季度分组（1-3 月 / 4-6 月 / 7-9 月 / 10-12 月），新季度重新起算
- 「季度累计」sheet = 逐月明细行 + 季度小计行（SQLite 历史 + 本月结果合成）

## 已知边界

- **参考库需人工更新**：验货数据报表与协议签订记录不会自动拉取（飞书 wiki 自动拉取列入后续增强），
  有新版本时通过网页「更新参考库」或 CLI `--inspection` / `--agreements` 重新导入
- **买家备注需人工复核**：带备注的质量退货订单全部计入金额，同时在「买家备注复核清单」sheet 列出，
  由品质部人工复核后在 Excel 中手动调整
- **采购入库单导出范围要覆盖报告月**：样例文件仅覆盖 07-31~08-04 时，大量 SKU 会进「缺单价」
  校验项——这是数据覆盖问题不是缺陷，真实使用请导出更大时间范围
- **口径差异提示**：领星全原因退货合计 vs DLM 退货量合计差异率 > 5% 时在「数据校验」提示检查导出期间
- **真实数据集成测试依赖本机文件**：换机器或缺文件时自动 skip
- **部署到阿里云 ECS**：参照 `/opt/fund-dashboard` 模式放 `/opt/supplier-return-report`、uvicorn 跑
  8300 端口；部署属「修改生产配置」，执行前必须征得用户同意。步骤见上文「部署到阿里云 ECS」节，
  命令清单可由 `python -m deploy.service plan` 生成

## 项目结构

```
engine/          loaders.py 输入解析 / rules.py 纯函数规则 / report.py 版式输出 / pipeline.py 编排+SQLite+CLI / models.py 契约数据类
web/             main.py FastAPI + static/index.html 原生 JS 单页
deploy/          service.py 部署产物生成（systemd unit / 部署命令 / 探活）
tests/           fixtures_build.py 合成文件工厂 + 各层测试 + test_integration_real.py 端到端
data/            运行时目录（gitignore）：uploads/ reports/ ref/ app.db
```
