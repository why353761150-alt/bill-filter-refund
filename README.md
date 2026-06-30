# 信用卡账单退款记录清理工具

> 个人中国银行信用卡账单自动化处理：PDF → Excel → 退款清理 → 跨账期分离 → 消费分析
>
> 每月跑一次，七步搞定；中间三次人工介入（删候选、标备注、填类别），其余全自动。

## 这是什么

一个本地跑的 Python 工具，把"中国银行信用卡电子账单"从 PDF 一路处理到"分类消费报表"。

整个流程围绕"**账期**"组织（`YYYY-MM`），每个月一个独立目录，互不干扰。流水线有 7 步，其中 3 步需要你打开 Excel 改几笔数据然后保存——这就是"**人工介入**"，不复杂，每次一分钟左右。

设计上有一个核心原则：**步骤"是否完成"完全由产物文件决定**。没有隐藏的 `.done` 标记文件，没有需要单独维护的状态数据库。文件在 = 跑过了；文件不在 = 没跑过；想重跑？删文件即可。

## 三种跑法

### 1. 双击 `run_all.bat`（最常用）

第一次运行会自动：
1. 检查 Python
2. 创建虚拟环境 `.venv/`
3. 装依赖（用清华镜像，约 1-3 分钟）
4. 检查 `.env`（没有就跳过自动下载）
5. 启动流水线，**在 3 个介入点暂停等你**

之后每次跑都会自动跳过已完成的步骤。

### 2. VSCode 调试（按 F5）

双击 `start_vscode.bat` 打开项目，按 `F5`，在调试面板选"一键全流程"。

### 3. 命令行

```bash
.venv\Scripts\activate
python -m bill_pipeline.cli all --interactive    # 全流程
python -m bill_pipeline.cli status --period 2026-06  # 看进度
python -m bill_pipeline.cli reset --period 2026-06   # 清空某个账期
```

## 七步流水线

| 步骤 | 命令 | 自动 / 介入 | 产物 | 说明 |
|------|------|------------|------|------|
| 1. 抓取 | `fetch` | 自动 | `data/input/<账期>/*.pdf` | 从 QQ 邮箱拉账单；本地有 PDF 就跳过 |
| 2. 解析 | `parse` | 自动 | `data/working/<账期>/中国银行.xlsx` | PDF → Excel |
| 3. 退款候选 | `match-refunds` | 自动 | `*_退款候选清单.xlsx` | 找出可疑的退款行 |
| 4. 清理 | `clean` | **介入 ①** | `*_清理结果.xlsx` | 删掉不是退款的行 |
| 5. 跨账期分离 | `split-cross` | **介入 ②** | `中国银行_最终账单.xlsx` | 标"上期还款"等备注 |
| 6. 准备输入 | `prepare-input` | 自动 | `分析账单.xlsx` | 提取支出列待你分类 |
| 7. 分析 | `analyze` | **介入 ③** | `账单分析报告_<账期>.xlsx` | 用 DeepSeek 填类别后生成 |

## 三次人工介入详解

**介入 ①：删掉不是退款的行**（约 30 秒）

打开 `*_退款候选清单.xlsx`，看"退款候选"sheet。把明显不是退款的行（误判的）整行删掉，保存。
> 程序会用黄底提醒"这条金额很大，注意看"。

**介入 ②：标注跨账期交易**（约 1-2 分钟）

打开 `*_清理结果.xlsx`，看"剔除退款后账单"sheet。对每一条带黄底（金额 > 200）或备注为空的行，**在"备注"列**填入以下之一：

| 备注关键字 | 含义 |
|-----------|------|
| `上期还款` | 本期还了上期账单的钱 |
| `上期账单退款，已抵扣上期账单还款` | 上期退的款当时抵了上期还款 |
| `上期账单退款，已抵扣本期账单还款` | 上期退的款抵了本期要还的钱 |

不填就视为"本期消费"，照常计入分析。

> 这一步数据有错的话，最终账单的"本月应还"校验会变红——重看一遍是哪个行标错了。

**介入 ③：消费分类**（约 1-2 分钟）

打开 `分析账单.xlsx`，用 DeepSeek 给每笔消费填"交易类别"列（旅行/娱乐/网购/家居/通勤/通信费/物业/水费/电费/燃气费/待定）。填完保存，跑 analyze 出分类汇总报告。

> 参考提示词见文末。

## 凭据安全

项目**绝不**在代码里写邮箱密码。自动下载需要 `.env`：

```bash
copy .env.example .env
# 然后用文本编辑器打开 .env 填写：
#   BOC_EMAIL=你的QQ邮箱@qq.com
#   BOC_AUTH_CODE=你的QQ邮箱授权码
```

`.env` 在 `.gitignore` 里，不会被提交。占位符值会被识别为"未配置"，不会报错，直接跳过下载。

**不配置也能用**——把 PDF 手动放到 `data\input\<账期>\` 即可。

## 想重跑某一步？删它的产物

```bash
# 1) 重新跑 跨账期分离
del data\output\2026-06\中国银行_最终账单.xlsx
python -m bill_pipeline.cli split-cross --period 2026-06

# 2) 重置整个账期（清空所有产物，目录结构保留）
python -m bill_pipeline.cli reset --period 2026-06
```

**原理**：流水线的状态完全由"产物文件存不存在"决定：

| 步骤 | 产物 |
|------|------|
| fetch | `data/input/<账期>/*.pdf` |
| parse | `data/working/<账期>/中国银行.xlsx` |
| match_refunds | `data/working/<账期>/*_退款候选清单.xlsx` |
| clean | `data/working/<账期>/*_清理结果.xlsx` |
| split_cross | `data/output/<账期>/中国银行_最终账单.xlsx` |
| prepare_input | `data/working/<账期>/分析账单.xlsx` |
| analyze | `data/output/<账期>/账单分析报告_*.xlsx` |

`status` 命令会列出每个步骤对应的产物文件名，缺哪个一目了然。

## 项目结构

```
g:\project\bill\
├── README.md                       # 本文档
├── requirements.txt                # Python 依赖
├── .env.example                    # 邮箱配置模板
├── .gitignore
├── start_vscode.bat                # 双击开 VSCode
├── run_all.bat                     # 双击跑全流程
├── bill.code-workspace             # VSCode 工作区
│
├── bill_pipeline/                  # 核心代码包
│   ├── config.py                   # 路径、列名、配置
│   ├── excel_utils.py              # Excel 工具（合计行、标黄等）
│   ├── email_fetcher.py            # 邮箱下载
│   ├── pdf_parser.py               # PDF 解析
│   ├── refund_matcher.py           # 退款候选 + 清理
│   ├── cross_period.py             # 跨账期分离 + 校验
│   ├── analyzer.py                 # 账单分析 + 图表
│   ├── pipeline.py                 # 流水线编排 + 状态机
│   └── cli.py                      # 命令行入口
│
├── data/
│   ├── input/<YYYY-MM>/            # 原始账单（PDF + 解析后 xlsx）
│   ├── working/<YYYY-MM>/          # 人工标注中间产物
│   └── output/<YYYY-MM>/           # 最终产物
│
└── .vscode/launch.json             # VSCode 调试配置
```

## 关键设计

### 1. 账期隔离

每个月一个目录，`data/{input,working,output}/2026-06/` 是 6 月的全部数据。处理 4 月的账单时不会影响 5 月。

### 2. 产物驱动

状态判断零开销，零隐藏文件。代码里也写得很直白：

```python
# pipeline.py
def _is_done_by_product(step, period):
    products = _product_for(step, period)
    return len(products) > 0 and all(p.exists() for p in products)
```

`run_all.bat` 也是同样的逻辑——检查 `.deps_installed` 标记是为了避免每次都跑 `pip install`。

### 3. 人工介入的判定

介入点用 mtime 比较判断"用户改没改"：

```python
# 上游产物里最早一个的 mtime 作为基准
baseline = min(p.stat().st_mtime for p in upstream_products)
# 用户改过（产物 mtime 比基准晚）→ 放行
if product.stat().st_mtime > baseline:
    return True
# 否则暂停等用户
```

### 4. 跨账期校验

`split_cross` 跑完会自动校验：

- **上期欠款校验**：上期还款 + 上期账单退款(抵上期) = 上期欠款余额
- **本期欠款校验**：最终账单支出 - 上期账单退款(抵本期) = 本期欠款余额

差值 > 0.01 时"汇总"sheet 里会标红，控制台打 ✗。同时在"最终账单"sheet 末尾加一行 **本月应还 = 支出合计 - 存入合计**（红底白字），打开就看到。

### 5. 凭据安全

不写死在代码里 → `.env` → `.gitignore`。占位符值视为"未配置"，自动跳过下载。

## 消费分类提示词（介入 ③ 用）

把 `分析账单.xlsx` 内容发给 DeepSeek，附上：

```
这是我的账单，我要对每笔消费进行分类，旅行，娱乐，网购，家居，
通勤，通信费，物业，水费、电费、燃气费、你帮我按照以上分类对
消费进行分类，如果拿不准的就空着，最后按照原表格的形式，我
可以直接复制的方式输出给我。

如果摘要中明确有"旅行"、"机票"、"酒店"，或者全是英文，等，归为旅行。
如果摘要中明确有"电影"、"娱乐"、"酒吧"、"美团"等，归为娱乐。
如果摘要中明确有"淘宝"、"京东"、"拼多多"等，归为网购。
如果摘要中明确有"公交"、"地铁"、"巴士"、"打车"、"哈罗"等，归为通勤。
如果摘要中明确有"手机充值"、"话费"等，归为通信费。
如果摘要中明确有"物业"，归为物业。
如果摘要中明确有"水费"、"电费"、"燃气费"，分别对应。
"生活缴费"标注为待定。
```

## 常见问题

**Q: 跑全流程时中途打断了怎么办？**
A: 重跑 `run_all.bat`，已完成的步骤会自动跳过。

**Q: 某一步出错了想重跑？**
A: 删那一步的产物文件，再跑流水线。详见上面"想重跑某一步"。

**Q: 介入点不暂停、直接退出了？**
A: 非交互模式下遇到人工介入点会直接退出。加 `--interactive` 参数：
```bash
python -m bill_pipeline.cli all --interactive
```
或者直接用 `run_all.bat`（默认带 `--interactive`）。

**Q: 大额支出阈值 200 元能改吗？**
A: 在 `.env` 中加 `HIGHLIGHT_THRESHOLD=500`。

**Q: 换电脑怎么迁移？**
A: 整个项目目录拷过去即可，所有路径都是相对的。`run_all.bat` 首次运行会重新建虚拟环境、装依赖。

**Q: 想支持其他银行怎么办？**
A: 仿照 `pdf_parser.py` 写一个 `xxx_bank_parser.py`，再在 `pipeline.py` 注册新步骤。

## 故障排查

| 现象 | 原因 | 解决 |
|------|------|------|
| `ModuleNotFoundError: No module named 'pdfplumber'` | 依赖没装 | 删 `.deps_installed` 标记后重跑 `run_all.bat` |
| `找不到 PDF` | PDF 没放进 input 目录 | 把 PDF 放到 `data/input/<账期>/` |
| `未配置邮箱` | .env 缺失或内容是占位符 | 编辑 `.env` 填入真实凭据 |
| 跑全流程直接退出 | 非交互模式遇到人工检查点 | 用 `run_all.bat` 或加 `--interactive` |
| Excel 打开是乱码 | 系统编码问题 | 用 UTF-8 编码或升级 Excel |

## 开发说明

- Python 3.10+
- 依赖见 `requirements.txt`
- 所有路径基于 `PROJECT_ROOT`，无硬编码绝对路径
- 每个模块可独立运行（除 `pipeline.py` 外）
