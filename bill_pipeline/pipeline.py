"""
流水线编排：状态机 + 3 次人工介入检查点。

设计要点：
- **状态完全由产物决定**：不做"已执行/未执行"的标记文件，而是看产物文件是否存在。
  - 步骤 A 的产物缺失 → 跑 A
  - 步骤 A 的产物存在 → 跳过 A（除非在检查点需要"用户改过"）
- 人工介入点：检测产物文件被修改时间是否晚于其上一步的产物
- 用户视角："看产物有没有，要就用，不要就重跑"——和文件管理器一致
"""
from __future__ import annotations

import time
from pathlib import Path
from typing import Callable

from .config import (
    RAW_BILL_NAME,
    REFUND_CANDIDATE_SUFFIX,
    CLEAN_RESULT_SUFFIX,
    FINAL_BILL_NAME,
    ANALYSIS_INPUT_NAME,
    ANALYSIS_REPORT_PREFIX,
    period_dirs,
    identify_period,
    current_period,
    log,
)


# ============================================================
# 步骤注册表
# ============================================================
STEPS = [
    "fetch",          # 1. 下载账单
    "parse",          # 2. PDF 解析
    "match_refunds",  # 3. 生成退款候选（人工 ①）
    "clean",          # 4. 应用退款清单（人工 ②：备注列）
    "split_cross",    # 5. 分离跨账期
    "prepare_input",  # 6. 准备分析输入（人工 ③：交易类别）
    "analyze",        # 7. 生成分析报告
]


# ============================================================
# 产物映射：每个步骤"产出"什么文件
#
# 设计原则：这是流水线唯一的真相来源。
# 步骤是否已完成 = 它的产物文件是否存在。
# 没有任何 .done 标记文件。
# ============================================================
def _product_for(step: str, period: str) -> list[Path]:
    """
    返回某步骤的所有产物文件路径。
    跑前检查"是否都存在"；存在 → 已完成；否则 → 跑。
    """
    dirs = period_dirs(period)
    working = dirs["working"]
    output = dirs["output"]
    input_dir = dirs["input"]

    if step == "fetch":
        # 产出一个 PDF（去重，大小写不敏感）
        seen = set()
        unique = []
        for p in list(input_dir.glob("*.pdf")) + list(input_dir.glob("*.PDF")):
            if p.name.lower() not in seen:
                seen.add(p.name.lower())
                unique.append(p)
        return unique

    if step == "parse":
        # 产出原始账单 xlsx
        f = working / RAW_BILL_NAME
        return [f] if f.exists() else []

    if step == "match_refunds":
        # 产出退款候选清单
        return list(working.glob(f"*{REFUND_CANDIDATE_SUFFIX}"))

    if step == "clean":
        # 产出清理结果
        return list(working.glob(f"*{CLEAN_RESULT_SUFFIX}"))

    if step == "split_cross":
        # 产出最终账单
        f = output / FINAL_BILL_NAME
        return [f] if f.exists() else []

    if step == "prepare_input":
        # 产出分析账单
        f = working / ANALYSIS_INPUT_NAME
        return [f] if f.exists() else []

    if step == "analyze":
        # 产出分析报告 xlsx
        return list(output.glob(f"{ANALYSIS_REPORT_PREFIX}*.xlsx"))

    return []


def _is_done_by_product(step: str, period: str) -> bool:
    """步骤是否完成 = 它的产物文件是否已存在"""
    products = _product_for(step, period)
    return len(products) > 0 and all(p.exists() for p in products)


# ============================================================
# 检查点：执行本步骤之前需要等用户改完的上游产物
#
# 判定逻辑改为对比"上游产物 vs 上上游产物"，
# 不再依赖任何 .done 文件。
# ============================================================
def _wait_for_upstream_modification(period: str, current_step: str,
                                    upstream_step: str, product: Path,
                                    instruction: str, interactive: bool) -> bool:
    """
    在执行 current_step 之前，等用户修改完上游产物。

    判定：product mtime > 上游步骤的产物 mtime
    （即：用户改过文件 = 这个文件比它的输入文件还新）

    返回 True = 用户已修改（或产物不存在），False = 未修改
    """
    if not product.exists():
        return True  # 产物都不存在，没法判断，视为通过

    # 找到上游步骤的"最早期产物"作为基准时间
    upstream_products = _product_for(upstream_step, period)
    if not upstream_products:
        return True  # 上游产物都没有，基准时间用现在

    # 取上游产物中最早的一个 mtime 作为"代码生成完成时间"
    baseline = min(p.stat().st_mtime for p in upstream_products if p.exists())

    if product.stat().st_mtime > baseline:
        return True

    log.info("=" * 50)
    log.info(f"⏸️  {current_step} 执行前需要你修改文件")
    log.info(f"   打开: {product}")
    log.info(f"   {instruction}")
    log.info("   改完保存后，按 Enter 继续...")
    log.info("=" * 50)

    if not interactive:
        return False  # 非交互模式 → 让流水线退出

    try:
        input()
        return True
    except EOFError:
        log.warning("非交互模式，跳过等待")
        return False


def _run_step(period: str, step: str, runner: Callable[[], Path | list[Path]],
              interactive: bool = False) -> Path | None:
    """
    通用步骤执行器（产物驱动）。

    逻辑：
    1. 产物已存在 → 跳过
    2. 检查点：在执行前，等用户改完上游产物
    3. 执行 → 产物由 runner 写出，本函数不写任何标记

    Args:
        interactive: 是否在人工介入点 input() 等待
    """
    if _is_done_by_product(step, period):
        log.info(f"⏭️  {step} 产物已存在，跳过")
        return None

    # 检查点：在执行本步骤前，等用户改完上游产物
    if step in CHECKPOINT_BEFORE:
        cfg = CHECKPOINT_BEFORE[step]
        for product in cfg["product"](period):
            if not _wait_for_upstream_modification(
                period=period,
                current_step=step,
                upstream_step=cfg["upstream"],
                product=product,
                instruction=cfg["instruction"],
                interactive=interactive,
            ):
                log.info(f"⏸️  {step} 之前需要你修改产物，退出流水线")
                return WAITING_FOR_HUMAN

    log.info(f"▶️  开始执行: {step}")
    t0 = time.time()
    try:
        result = runner(period)
        if result is None:
            log.warning(f"⏭️  {step} 跳过（无可用输入）")
            return None
        outputs = result if isinstance(result, list) else [result]
        log.info(f"✅ {step} 完成 ({time.time() - t0:.1f}s)，产物: {[p.name for p in outputs]}")
        return result
    except Exception as e:
        log.error(f"❌ {step} 失败: {e}")
        import traceback
        traceback.print_exc()
        return None


# 特殊返回值：表示步骤已完成但需要用户改产物才能继续
WAITING_FOR_HUMAN = object()


# ============================================================
# 各步骤的具体实现
# ============================================================
def _find_input_pdf(period: str) -> Path | None:
    """在 data/input/<period>/ 中找 PDF"""
    d = period_dirs(period)["input"]
    pdfs = list(d.glob("*.pdf")) + list(d.glob("*.PDF"))
    if not pdfs:
        # 兜底：直接在 data/input/ 根目录找
        pdfs = list((Path(__file__).parent.parent / "data" / "input").glob("*.pdf"))
    return pdfs[0] if pdfs else None


def _find_raw_bill(period: str) -> Path | None:
    """在 data/working/<period>/ 中找原始账单"""
    d = period_dirs(period)["working"]
    f = d / RAW_BILL_NAME
    if f.exists():
        return f
    # 兜底：input 目录
    f2 = period_dirs(period)["input"] / RAW_BILL_NAME
    return f2 if f2.exists() else None


def step_fetch(period: str) -> Path | None:
    """下载账单。如果 input 目录已有 PDF 则跳过。"""
    existing_pdf = _find_input_pdf(period)
    if existing_pdf:
        log.info(f"⏭️  本地已有 PDF: {existing_pdf.name}，跳过下载")
        return existing_pdf
    from .email_fetcher import fetch_bill
    return fetch_bill()


def step_parse(period: str) -> Path | None:
    from .pdf_parser import parse_pdf
    pdf = _find_input_pdf(period)
    if not pdf:
        log.warning(f"💡 账期 {period} 未找到 PDF，跳过")
        return None
    return parse_pdf(pdf)


def step_match_refunds(period: str) -> Path | None:
    from .refund_matcher import match_refunds
    bill = _find_raw_bill(period)
    if not bill:
        log.warning(f"💡 账期 {period} 未找到原始账单，跳过")
        return None
    return match_refunds(bill)


def step_clean(period: str) -> Path | None:
    from .refund_matcher import apply_refunds
    bill = _find_raw_bill(period)
    dirs = period_dirs(period)
    candidates = list(dirs["working"].glob(f"*{REFUND_CANDIDATE_SUFFIX}"))
    if not bill or not candidates:
        log.warning(f"💡 账期 {period} 缺少账单或退款清单，跳过")
        return None
    return apply_refunds(bill, candidates[0])


def step_split_cross(period: str) -> Path | None:
    from .cross_period import split_cross_period
    dirs = period_dirs(period)
    cleaned = list(dirs["working"].glob(f"*{CLEAN_RESULT_SUFFIX}"))
    if not cleaned:
        log.warning(f"💡 账期 {period} 未找到清理结果，跳过")
        return None
    return split_cross_period(cleaned[0])


def step_prepare_input(period: str) -> Path | None:
    """把最终账单的支出列复制到 working/分析账单.xlsx（人工填类别）"""
    from .excel_utils import ensure_transaction_columns
    import pandas as pd
    from openpyxl import load_workbook

    dirs = period_dirs(period)
    final = dirs["output"] / FINAL_BILL_NAME
    target = dirs["working"] / ANALYSIS_INPUT_NAME

    if not final.exists():
        log.warning(f"💡 未找到最终账单: {final}")
        return None

    df = pd.read_excel(final, sheet_name="最终账单")
    # 只保留支出 > 0 的记录
    expenses = df[df["支出"] > 0].copy()
    # 分析账单不需要"备注"列（已经在跨账期分离时处理过了）
    expenses = expenses[["交易日", "交易描述", "支出"]]
    expenses.columns = ["交易日", "交易摘要", "交易金额（RMB）"]
    expenses["交易类别"] = ""  # 人工填

    # 如果已存在且人工已修改（mtime 比最终账单晚），不覆盖
    if target.exists() and target.stat().st_mtime > final.stat().st_mtime:
        log.info(f"   分析账单已有人工修改，跳过覆盖")
        return target

    expenses.to_excel(target, sheet_name="分析账单", index=False)
    log.info(f"📝 已准备分析输入: {target}")
    log.info(f"   请手工填'交易类别'列（参考 analyzer.py 顶部提示词）")
    return target


def step_analyze(period: str) -> Path | None:
    from .analyzer import analyze
    dirs = period_dirs(period)
    ai = dirs["working"] / ANALYSIS_INPUT_NAME
    if not ai.exists():
        log.warning(f"💡 未找到分析账单: {ai}")
        return None
    return analyze(ai)


# ============================================================
# 流水线编排
# ============================================================
STEP_RUNNERS = {
    "fetch": step_fetch,
    "parse": step_parse,
    "match_refunds": step_match_refunds,
    "clean": step_clean,
    "split_cross": step_split_cross,
    "prepare_input": step_prepare_input,
    "analyze": step_analyze,
}

# 检查点：执行本步骤之前需要等用户改完的上游产物
# 键 = 当前步骤名（在它跑之前要等）
# 值 = 字典 {
#   "upstream": 哪个步骤的产物要被改（即上游步骤）,
#   "product": 产物的 getter（返回要检查的文件路径列表）,
#   "instruction": 提示用户改什么
# }
CHECKPOINT_BEFORE = {
    "clean": {
        "upstream": "match_refunds",
        "product": lambda p: list(period_dirs(p)["working"]
                                   .glob(f"*{REFUND_CANDIDATE_SUFFIX}")),
        "instruction": "在'退款候选'sheet 删除不是退款的行（黄行尤其要仔细看），然后保存",
    },
    "split_cross": {
        "upstream": "clean",
        "product": lambda p: list(period_dirs(p)["working"]
                                   .glob(f"*{CLEAN_RESULT_SUFFIX}")),
        "instruction": "在'剔除退款后账单'sheet 的'备注'列填写跨账期标注（'上期还款' / '上期账单退款，已抵扣上期账单还款' / '上期账单退款，已抵扣本期账单还款'），然后保存",
    },
    "analyze": {
        "upstream": "prepare_input",
        "product": lambda p: [period_dirs(p)["working"] / ANALYSIS_INPUT_NAME],
        "instruction": "用 DeepSeek 给每笔消费分类，填入'交易类别'列，然后保存",
    },
}


def run_pipeline(period: str | None = None, interactive: bool = False,
                 start_from: str | None = None,
                 only: str | None = None) -> dict:
    """
    串联执行所有步骤（产物驱动）。

    Args:
        period: 账期 YYYY-MM，默认当前月
        interactive: 是否在人工介入点 input() 等待
        start_from: 从某个步骤开始（用于跳过已完成部分）
        only: 只跑某个步骤
    """
    period = period or current_period()
    log.info(f"🚀 启动流水线: 账期 = {period}")
    log.info(f"   交互模式: {interactive}")
    period_dirs(period)

    results: dict[str, Path | None] = {}

    steps_to_run = STEPS
    if only:
        if only not in STEPS:
            raise ValueError(f"未知步骤: {only}, 可选: {STEPS}")
        steps_to_run = [only]
    elif start_from and start_from in STEPS:
        steps_to_run = STEPS[STEPS.index(start_from):]

    for step in steps_to_run:
        runner = STEP_RUNNERS[step]
        result = _run_step(period, step, runner, interactive=interactive)
        results[step] = result

        # 检查点未完成 → 退出流水线
        if result is WAITING_FOR_HUMAN:
            log.info(
                f"⏸️  {step} 之前需要你修改产物，"
                f"改完后跑 'python -m bill_pipeline.cli {step}' 或 'cli.py all --from {step}' 继续"
            )
            break

    log.info("=" * 50)
    log.info("📋 流水线结果")
    log.info("=" * 50)
    for step, result in results.items():
        if result is WAITING_FOR_HUMAN:
            status = "⏸️"
        elif result is not None:
            status = "✅"
        elif _is_done_by_product(step, period):
            status = "✅"  # 产物已存在（被跳过）
        elif step in CHECKPOINT_BEFORE:
            status = "⏸️"
        else:
            status = "❌"
        log.info(f"  {status} {step}: {result}")
    return results


def show_status(period: str | None = None) -> dict:
    """显示某个账期各步骤的完成状态（按产物判断）"""
    period = period or current_period()
    status = {}
    log.info(f"📊 账期 {period} 状态（按产物判断）:")
    for step in STEPS:
        products = _product_for(step, period)
        done = bool(products) and all(p.exists() for p in products)
        status[step] = "✓" if done else "○"
        if products:
            product_names = [p.name for p in products]
        else:
            product_names = ["(无产物)"]
        log.info(f"  {status[step]} {step:<16}  →  {', '.join(product_names)}")
    return status


def reset_pipeline(period: str | None = None):
    """重置某账期：删除所有步骤的产物文件（无需碰任何 .done 标记，因为根本没有）"""
    import shutil
    period = period or current_period()
    dirs = period_dirs(period)

    log.info(f"⚠️  即将删除账期 {period} 的所有产物:")
    # 删 input (PDF)、working (中间产物)、output (最终账单+分析报告)
    for sub in ("input", "working", "output"):
        d = dirs[sub]
        if d.exists():
            for f in d.iterdir():
                if f.is_file():
                    log.info(f"   - {f.relative_to(dirs[sub].parent)}")
                    f.unlink()
            # 目录保留，文件清空
    log.info(f"🗑️  已清空账期 {period} 的所有产物文件")
    log.info(f"   目录结构保留（input/working/output/），下次跑会重新生成")
