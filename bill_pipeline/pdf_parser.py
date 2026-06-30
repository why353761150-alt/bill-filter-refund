"""
中国银行信用卡电子账单 PDF → 标准 Excel。

输入：data/input/<账期>/<账单>.PDF
输出：data/input/<账期>/中国银行.xlsx（11 列：7 列交易 + 4 列摘要）
"""
from __future__ import annotations

import re
from pathlib import Path
import pandas as pd
import pdfplumber

from .config import (
    TRANSACTION_COLUMNS,
    SUMMARY_COLUMNS,
    RAW_BILL_NAME,
    log,
    period_dirs,
    identify_period,
    current_period,
)
from .excel_utils import parse_money, normalize_money_column


# 中行账单典型列名
TEMPLATE_HEADERS = [
    "交易日 Transaction Date",
    "银行记账日 Posting Date",
    "卡号后四位 Last Four Digits of Card Number",
    "交易描述 Description",
    "存入 Deposit",
    "支出 Expenditure",
]


def _extract_summary(first_page) -> dict[str, float]:
    """从首页提取摘要信息（上期/本期余额、支出、存入）"""
    summary = {k: None for k in SUMMARY_COLUMNS}
    tables = first_page.extract_tables()

    for table in tables:
        table_str = str(table)
        if "本期支出金额" not in table_str and "Purchase" not in table_str:
            continue

        for row in table:
            row_clean = [str(x).replace("\n", "") if x else "" for x in row]
            if "人民币" in row_clean[0] or "RMB" in row_clean[0]:
                try:
                    summary["上期欠款余额"] = parse_money(row_clean[1])
                    summary["本期支出金额"] = parse_money(row_clean[2])
                    summary["本期存入金额"] = parse_money(row_clean[3])
                    summary["本期欠款余额"] = parse_money(row_clean[4])
                    log.info(f"📊 摘要: {summary}")
                    return summary
                except Exception as e:
                    log.warning(f"摘要解析部分失败: {e}")
    return summary


def _extract_transactions(pdf) -> list[list[str]]:
    """提取所有页面的交易明细"""
    transactions = []
    is_transaction_table = False
    headers: list[str] = []

    for page in pdf.pages:
        for table in page.extract_tables() or []:
            if not table:
                continue
            first_row_str = "".join(str(x) for x in table[0] if x)

            if "交易日" in first_row_str or "Transaction Date" in first_row_str:
                is_transaction_table = True
                headers = [str(h).replace("\n", " ") for h in table[0]]
                transactions.extend(table[1:])
            elif is_transaction_table:
                # 跨页续表
                if len(table[0]) >= 5 and "积分" not in str(table[0]):
                    transactions.extend(table)
    return transactions, headers


def _build_dataframe(transactions: list[list[str]],
                     headers: list[str]) -> pd.DataFrame:
    """从原始行构建标准 DataFrame"""
    if not transactions:
        return pd.DataFrame(columns=TRANSACTION_COLUMNS + SUMMARY_COLUMNS)

    df = pd.DataFrame(transactions)

    # 尝试匹配表头
    if len(df.columns) == len(headers):
        df.columns = headers
    elif len(df.columns) >= 6:
        df = df.iloc[:, :6]
        df.columns = TEMPLATE_HEADERS
    else:
        log.warning(f"列数异常 ({len(df.columns)}), 强制取前 6 列")
        df = df.iloc[:, :6] if df.shape[1] >= 6 else df
        df.columns = TEMPLATE_HEADERS[: df.shape[1]]

    # 清洗换行符
    df = df.replace(r"\n", "", regex=True)

    # 过滤无效行：交易日必须含数字
    first_col = df.columns[0]
    df = df[df[first_col].astype(str).str.contains(r"\d", na=False)]

    # 构建标准 7 列
    result = pd.DataFrame()
    result["交易日"] = df.iloc[:, 0].astype(str).str.strip()
    result["银行记账日"] = df.iloc[:, 1].astype(str).str.strip() if df.shape[1] > 1 else ""
    result["卡号后四位"] = df.iloc[:, 2].astype(str).str.strip() if df.shape[1] > 2 else ""
    result["交易描述"] = df.iloc[:, 3].astype(str).str.strip() if df.shape[1] > 3 else ""
    result["存入"] = normalize_money_column(df.iloc[:, 4]) if df.shape[1] > 4 else 0.0
    result["支出"] = normalize_money_column(df.iloc[:, 5]) if df.shape[1] > 5 else 0.0
    result["备注"] = ""

    return result


def parse_pdf(pdf_path: Path | str) -> Path:
    """
    解析单个 PDF 账单，输出标准化 Excel。

    Returns:
        输出 Excel 路径
    """
    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"找不到 PDF: {pdf_path}")

    # 识别账期，决定输出目录
    period = identify_period(pdf_path.name) or current_period()
    dirs = period_dirs(period)
    output_path = dirs["input"] / RAW_BILL_NAME

    log.info(f"📄 解析 PDF: {pdf_path.name}")
    log.info(f"📅 账期: {period}, 文件大小: {pdf_path.stat().st_size} 字节")

    with pdfplumber.open(pdf_path) as pdf:
        log.info(f"   共 {len(pdf.pages)} 页")
        summary = _extract_summary(pdf.pages[0])
        transactions, headers = _extract_transactions(pdf)

    if not transactions:
        log.warning("⚠️  PDF 中未找到交易明细")
        return output_path

    df = _build_dataframe(transactions, headers)
    log.info(f"   提取到 {len(df)} 条交易记录")

    # 填充摘要列（仅第 1 行）
    for key, val in summary.items():
        if val is not None and not df.empty:
            df.loc[0, key] = val

    # 补全缺失的摘要列
    for col in SUMMARY_COLUMNS:
        if col not in df.columns:
            df[col] = None

    # 重新排列：标准 7 列 + 4 列摘要
    final_columns = TRANSACTION_COLUMNS + SUMMARY_COLUMNS
    for col in final_columns:
        if col not in df.columns:
            df[col] = ""
    df = df[final_columns]

    df.to_excel(output_path, sheet_name="原始账单", index=False)
    log.info(f"✅ 已保存: {output_path}")
    return output_path
