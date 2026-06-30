"""
全局配置：路径、列名、汇总键、账期识别。

设计原则：
- 所有路径基于 PROJECT_ROOT 相对解析，不硬编码绝对路径
- 所有列名集中定义，避免散落
- 账期(period) 是一切状态隔离的边界
"""
from __future__ import annotations

import os
import re
import logging
from pathlib import Path
from dataclasses import dataclass
from dotenv import load_dotenv

# ============================================================
# 路径常量
# ============================================================
PROJECT_ROOT = Path(__file__).resolve().parent.parent

# 加载 .env（不存在也不报错）
load_dotenv(PROJECT_ROOT / ".env")

DATA_DIR = PROJECT_ROOT / "data"
DATA_INPUT = DATA_DIR / "input"
DATA_WORKING = DATA_DIR / "working"
DATA_OUTPUT = DATA_DIR / "output"

# 确保目录存在
for d in (DATA_INPUT, DATA_WORKING, DATA_OUTPUT):
    d.mkdir(parents=True, exist_ok=True)

# ============================================================
# 统一数据模型（解决前后不一致问题）
# ============================================================
# 账单明细 7 列：所有"清理后账单"统一使用这套列
TRANSACTION_COLUMNS = [
    "交易日",          # Transaction Date
    "银行记账日",      # Posting Date
    "卡号后四位",      # Last Four Digits
    "交易描述",        # Description
    "存入",            # Deposit
    "支出",            # Expenditure
    "备注",            # 人工填：跨账期/退款标注
]

# 原始账单附加的摘要列（仅 PDF 解析后的中间产物使用）
SUMMARY_COLUMNS = [
    "上期欠款余额",
    "本期支出金额",
    "本期存入金额",
    "本期欠款余额",
]

# 最终账单使用的列名（用于分析账单）
ANALYSIS_COLUMNS = [
    "交易日",
    "交易摘要",
    "交易金额（RMB）",
    "交易类别",     # 人工用 DeepSeek 填
]

# 跨账期备注关键字
CROSS_PERIOD_NOTES = {
    "LAST_REPAYMENT": "上期还款",
    "LAST_REFUND_OFFSET_LAST": "上期账单退款，已抵扣上期账单还款",
    "LAST_REFUND_OFFSET_THIS": "上期账单退款，已抵扣本期账单还款",
}

# ============================================================
# 业务参数（可被环境变量覆盖）
# ============================================================
HIGHLIGHT_THRESHOLD = float(os.getenv("HIGHLIGHT_THRESHOLD", "200"))  # 大额支出标黄阈值
TEMPLATE_AMOUNT = 0.01  # 金额匹配容差

# 原始账单文件名（在 input 目录的相对名）
RAW_BILL_NAME = "中国银行.xlsx"
REFUND_CANDIDATE_SUFFIX = "_退款候选清单.xlsx"
CLEAN_RESULT_SUFFIX = "_清理结果.xlsx"
FINAL_BILL_NAME = "中国银行_最终账单.xlsx"
ANALYSIS_INPUT_NAME = "分析账单.xlsx"
ANALYSIS_REPORT_PREFIX = "账单分析报告_"

# 账期正则（匹配 PDF 文件名中的"2026年04月"或"2026-04"等）
PERIOD_PATTERNS = [
    re.compile(r"(\d{4})年(\d{1,2})月"),
    re.compile(r"(\d{4})[-_](\d{1,2})"),
]

# ============================================================
# 日志
# ============================================================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("bill")


# ============================================================
# 配置工具
# ============================================================
@dataclass
class EmailConfig:
    email: str
    password: str
    imap_server: str
    download_dir: Path


def get_email_config() -> EmailConfig | None:
    """
    从环境变量读取邮箱配置。
    任一关键字段缺失则返回 None，调用方应优雅降级。
    """
    email = os.getenv("BOC_EMAIL", "").strip()
    code = os.getenv("BOC_AUTH_CODE", "").strip()
    server = os.getenv("BOC_IMAP_SERVER", "imap.qq.com").strip()

    # 占位符视为未配置
    placeholders = {"your_email@qq.com", "your_authorization_code_here"}
    if not email or not code or email in placeholders or code in placeholders:
        return None

    download_dir = Path(os.getenv("BOC_DOWNLOAD_DIR", str(DATA_INPUT))).resolve()
    download_dir.mkdir(parents=True, exist_ok=True)

    return EmailConfig(
        email=email,
        password=code,
        imap_server=server,
        download_dir=download_dir,
    )


def identify_period(text: str) -> str | None:
    """
    从文件名/路径中识别账期，返回 YYYY-MM 格式。
    无法识别返回 None。
    """
    if not text:
        return None
    for pat in PERIOD_PATTERNS:
        m = pat.search(text)
        if m:
            year, month = m.group(1), int(m.group(2))
            if 1 <= month <= 12:
                return f"{year}-{month:02d}"
    return None


def current_period() -> str:
    """当前月作为默认账期"""
    import datetime
    now = datetime.date.today()
    return f"{now.year}-{now.month:02d}"


def period_dirs(period: str) -> dict[str, Path]:
    """获取某账期下的所有目录（自动创建）"""
    dirs = {
        "input": DATA_INPUT / period,
        "working": DATA_WORKING / period,
        "output": DATA_OUTPUT / period,
    }
    for d in dirs.values():
        d.mkdir(parents=True, exist_ok=True)
    return dirs
