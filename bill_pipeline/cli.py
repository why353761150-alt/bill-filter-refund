"""
命令行入口：python -m bill_pipeline.cli <command> [options]

可用命令：
    fetch           # 1. 下载账单
    parse           # 2. PDF 解析
    match-refunds   # 3. 生成退款候选
    clean           # 4. 应用退款清单
    split-cross     # 5. 分离跨账期
    prepare-input   # 6. 准备分析输入
    analyze         # 7. 生成分析报告
    all             # 一键全流程
    status          # 查看状态（按产物判断）
    reset           # 清空某账期的所有产物文件

状态判断：完全基于产物文件，不使用任何隐藏的 .done 标记。
- 想重跑某一步？删掉它的产物文件，再跑流水线。
- 想重置整个账期？用 'reset' 命令。
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .config import identify_period, current_period, log, DATA_INPUT


COMMAND_MAP = {
    "fetch": "fetch",
    "parse": "parse",
    "match-refunds": "match_refunds",
    "clean": "clean",
    "split-cross": "split_cross",
    "prepare-input": "prepare_input",
    "analyze": "analyze",
    "all": None,  # 特殊处理
    "status": None,  # 特殊处理
    "reset": None,  # 特殊处理
}


def _guess_period(args_period: str | None) -> str:
    """从 --period / 当前月 / input 目录文件名中识别账期"""
    if args_period:
        return args_period
    # 从 input 目录的 PDF/xlsx 文件名识别
    if DATA_INPUT.exists():
        for f in DATA_INPUT.rglob("*"):
            if f.is_file():
                p = identify_period(f.name)
                if p:
                    return p
    return current_period()


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="bill",
        description="中国银行信用卡账单处理流水线",
    )
    parser.add_argument("command", choices=list(COMMAND_MAP.keys()),
                        help="要执行的命令")
    parser.add_argument("--period", "-p", help="账期 YYYY-MM（默认自动识别）")
    parser.add_argument("--interactive", "-i", action="store_true",
                        help="在人工介入点暂停等待（仅 all 命令）")
    parser.add_argument("--from", dest="start_from", help="从某个步骤开始")

    args = parser.parse_args(argv)

    period = _guess_period(args.period)
    log.info(f"账期: {period}")

    if args.command == "status":
        from .pipeline import show_status
        show_status(period)
        return 0

    if args.command == "reset":
        from .pipeline import reset_pipeline
        reset_pipeline(period)
        return 0

    if args.command == "all":
        from .pipeline import run_pipeline
        run_pipeline(period=period, interactive=args.interactive,
                     start_from=getattr(args, "start_from", None))
        return 0

    # 单步执行
    step = COMMAND_MAP[args.command]
    from .pipeline import STEP_RUNNERS, run_pipeline
    run_pipeline(period=period, only=step)
    return 0


if __name__ == "__main__":
    sys.exit(main())
