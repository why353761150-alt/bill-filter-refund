# 中国银行信用卡账单处理流水线

> 月度信用卡账单处理自动化：PDF → Excel → 退款清理 → 跨账期分离 → 消费分析

## 快速开始

### 1. 打开项目

双击 `start_vscode.bat`，会在当前目录打开 VSCode。

### 2. 安装依赖（首次）

VSCode 中按 `` Ctrl+` `` 打开终端，执行：

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

或者直接双击 `run_all.bat`，首次运行会自动创建虚拟环境并安装依赖。

### 3. （可选）配置邮箱自动下载

如果希望自动从邮箱下载账单 PDF：

```bash
copy .env.example .env
```

然后用文本编辑器打开 `.env`，填写：

```ini
BOC_EMAIL=你的QQ邮箱@qq.com
BOC_AUTH_CODE=你的QQ邮箱授权码
```

> 授权码获取：QQ 邮箱 → 设置 → 账户 → POP3/IMAP 服务 → 开启 → 生成授权码

**不配置也不影响使用**，只是不能自动下载，需要手动把 PDF 放入 `data\input\<账期>\`。

### 4. 跑全流程

**方式 A：双击 `run_all.bat`**

- 第一次：会依次跑所有步骤，**在 3 个人工介入点暂停**等你处理
- 后续：自动跳过已完成的步骤

**方式 B：VSCode 中按 `F5`**

在调试面板选择「一键全流程」→ 按 F5

**方式 C：终端命令行**

```bash
.venv\Scripts\activate
python -m bill_pipeline.cli all --interactive
```

### 5. 跑单个步骤

```bash
python -m bill_pipeline.cli fetch            # 下载账单
python -m bill_pipeline.cli parse            # PDF → Excel
python -m bill_pipeline.cli match-refunds    # 生成退款候选
python -m bill_pipeline.cli clean            # 清理账单
python -m bill_pipeline.cli split-cross      # 分离跨账期
python -m bill_pipeline.cli prepare-input    # 准备分析输入
python -m bill_pipeline.cli analyze          # 生成分析报告
python -m bill_pipeline.cli status           # 查看状态
python -m bill_pipeline.cli reset            # 重置某账期状态
```

## 处理流程

```
┌─────────────────────────────────────────────────────────────┐
│ 1. 抓取账单    (自动)   PDF → data/input/<账期>/            │
│ 2. 解析 PDF    (自动)   → 中国银行.xlsx                      │
├─────────────────────────────────────────────────────────────┤
│ 3. 退款候选    (自动)   → _退款候选清单.xlsx                │
│    ┌───────────────────────────────────┐                    │
│    │ ⏸️ 人工介入 ①：打开候选清单，     │ ← 黄行需筛选      │
│    │   删除不是退款的行                │                    │
│    └───────────────────────────────────┘                    │
│ 4. 清理账单    (自动)   → _清理结果.xlsx                    │
│    ┌───────────────────────────────────┐                    │
│    │ ⏸️ 人工介入 ②：在"备注"列填写     │                    │
│    │   跨账期标注                      │                    │
│    └───────────────────────────────────┘                    │
│ 5. 跨账期分离  (自动)   → 中国银行_最终账单.xlsx            │
│    ↳ 大于 200 元的支出自动标黄                              │
├─────────────────────────────────────────────────────────────┤
│ 6. 准备输入    (自动)   → 分析账单.xlsx                     │
│    ┌───────────────────────────────────┐                    │
│    │ ⏸️ 人工介入 ③：用 DeepSeek 分类   │                    │
│    │   填"交易类别"列                  │                    │
│    └───────────────────────────────────┘                    │
│ 7. 分析报告    (自动)   → 账单分析报告_<账期>.xlsx          │
└─────────────────────────────────────────────────────────────┘
```

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
│   ├── cross_period.py             # 跨账期分离
│   ├── analyzer.py                 # 账单分析
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

每个月独立目录，互不干扰：

- `data/input/2026-04/` 是 4 月原始账单
- `data/output/2026-05/` 是 5 月最终账单
- **状态完全由产物决定**：没有隐藏的 `.done` 标记文件

### 2. 产物驱动的状态机

每个步骤"做了什么" = "产出了什么文件"。

| 步骤 | 产物文件 | 含义 |
|------|----------|------|
| fetch | `data/input/<账期>/*.pdf` | 账单 PDF |
| parse | `data/working/<账期>/中国银行.xlsx` | 解析后的原始账单 |
| match_refunds | `data/working/<账期>/*_退款候选清单.xlsx` | 退款候选清单 |
| clean | `data/working/<账期>/*_清理结果.xlsx` | 剔除退款后的账单 |
| split_cross | `data/output/<账期>/中国银行_最终账单.xlsx` | 分离跨账期后的最终账单 |
| prepare_input | `data/working/<账期>/分析账单.xlsx` | 准备分析的人工分类清单 |
| analyze | `data/output/<账期>/账单分析报告_*.xlsx` | 分类汇总报告 |

**判断规则**：
- 步骤的产物已存在 → 该步骤"已完成"，自动跳过
- 产物不存在 → 该步骤未跑，执行它
- 人工介入点：跑某一步前，对比"上游产物 mtime"和"本步骤产物 mtime"，如果产物 mtime 更新 → 用户已处理过，放行；否则提示用户去改

### 3. 用户重跑某个步骤

直接把它的产物文件删掉，再跑流水线即可——**不需要任何 .done 标记**：

```bash
# 重跑 split_cross 步骤
del data\output\2026-04\中国银行_最终账单.xlsx
python -m bill_pipeline.cli split-cross --period 2026-04

# 重置整个账期（清空所有产物）
python -m bill_pipeline.cli reset --period 2026-04
```

### 4. 幂等性

每个步骤都"安全可重跑"：
- 已完成的步骤自动跳过
- 手工保存过的文件不会被代码覆盖（基于 mtime 比较）

### 5. 人工介入检测

通过比较"代码生成时间"和"文件修改时间"判断人工是否处理过：

```python
# pipeline.py 里的判定逻辑
baseline = min(p.stat().st_mtime for p in upstream_products)  # 上游产物最早 mtime
if product.stat().st_mtime > baseline:
    # 用户改过了，放行
    return True
# 否则暂停等用户
```

### 6. 凭据安全

- **绝不**把邮箱密码写入代码
- 通过 `.env` 读取，`.env` 在 `.gitignore` 中
- 占位符值会被识别为"未配置"，自动跳过下载

## 消费分类提示词（用于人工介入 ③）

把 `分析账单.xlsx` 的内容发给 DeepSeek，附上以下提示词：

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
A: 重新跑 `run_all.bat`，已完成步骤会自动跳过。

**Q: 某一步出错了想重跑？**
A: `python -m bill_pipeline.cli reset --period 2026-04` 然后再跑。

**Q: 大额支出阈值 200 元能改吗？**
A: 在 `.env` 中加一行 `HIGHLIGHT_THRESHOLD=500`。

**Q: 换电脑怎么迁移？**
A: 整个项目目录拷过去即可，无需改任何路径配置（已全用相对路径）。

**Q: 想支持其他银行怎么办？**
A: 仿照 `pdf_parser.py` 写一个 `xxx_bank_parser.py`，再在 `pipeline.py` 注册新步骤。

## 故障排查

| 现象 | 原因 | 解决 |
|------|------|------|
| `ModuleNotFoundError: No module named 'pdfplumber'` | 依赖没装 | `pip install -r requirements.txt` |
| `找不到 PDF` | PDF 没放进 input 目录 | 把 PDF 放到 `data/input/<账期>/` |
| `未配置邮箱` | .env 缺失或内容是占位符 | 编辑 `.env` 填入真实凭据 |
| 跑全流程直接退出 | 非交互模式遇到人工检查点 | 加 `--interactive` 参数 |
| Excel 打开是乱码 | 系统编码问题 | 用 UTF-8 编码的 CSV 或升级 Excel |

## 开发说明

- Python 3.10+
- 依赖见 `requirements.txt`
- 所有路径基于 `PROJECT_ROOT`，无硬编码绝对路径
- 每个模块可独立运行（除 `pipeline.py` 外）
