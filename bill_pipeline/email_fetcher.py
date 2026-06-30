"""
中国银行信用卡账单自动下载（QQ 邮箱 IMAP）。

凭据从 .env 读取，未配置时优雅降级（返回 None，由调用方决定如何处理）。
"""
from __future__ import annotations

import datetime
import imaplib
import email
from email.header import decode_header
from email.utils import parsedate_to_datetime
from pathlib import Path

from .config import get_email_config, identify_period, log, period_dirs, current_period


def fetch_bill() -> Path | None:
    """
    下载当月中国银行信用卡电子账单 PDF。

    Returns:
        下载成功的 PDF 路径；若未配置邮箱 / 网络失败 / 未找到账单，返回 None。
    """
    cfg = get_email_config()
    if cfg is None:
        log.info("⏭️  未配置 BOC_EMAIL/BOC_AUTH_CODE，跳过自动下载")
        log.info("    请将 PDF 账单手动放入 data/input/<账期>/ 后重跑")
        return None

    log.info(f"🔌 连接邮箱 {cfg.email} ...")
    try:
        mail = imaplib.IMAP4_SSL(cfg.imap_server)
        mail.login(cfg.email, cfg.password)
        mail.select("inbox")
    except Exception as e:
        log.error(f"❌ 邮箱登录失败: {e}")
        return None

    try:
        today = datetime.date.today()
        first_day = today.replace(day=1).strftime("%d-%b-%Y")
        status, messages = mail.search(None, f'(SINCE "{first_day}")')
        if status != "OK":
            log.error("❌ 搜索邮件失败")
            return None

        email_ids = messages[0].split()
        bill_found = False
        downloaded = None

        # 倒序遍历：从最新邮件开始
        for e_id in reversed(email_ids):
            status, msg_data = mail.fetch(e_id, "(RFC822)")
            for response_part in msg_data:
                if not isinstance(response_part, tuple):
                    continue
                msg = email.message_from_bytes(response_part[1])

                # 解码主题
                subject, _ = decode_header(msg["Subject"])[0]
                if isinstance(subject, bytes):
                    subject = subject.decode("utf-8", errors="ignore")

                if "中国银行" in subject and "账单" in subject:
                    log.info(f"🔎 找到账单邮件: {subject}")

                    # 识别账期
                    period = identify_period(subject) or current_period()
                    dirs = period_dirs(period)

                    for part in msg.walk():
                        if part.get_content_maintype() == "multipart":
                            continue
                        if part.get("Content-Disposition") is None:
                            continue
                        filename = part.get_filename()
                        if not filename:
                            continue
                        fn, _ = decode_header(filename)[0]
                        if isinstance(fn, bytes):
                            fn = fn.decode("utf-8", errors="ignore")

                        if not fn.lower().endswith(".pdf"):
                            continue

                        target = dirs["input"] / fn
                        target.parent.mkdir(parents=True, exist_ok=True)
                        with open(target, "wb") as f:
                            f.write(part.get_payload(decode=True))
                        log.info(f"✅ 已下载: {target}")
                        downloaded = target
                        bill_found = True
                        break
                if bill_found:
                    break
            if bill_found:
                break

        if not bill_found:
            log.warning(f"💡 当月邮箱中未找到中国银行账单邮件")

        mail.logout()
        return downloaded

    except Exception as e:
        log.error(f"❌ 下载过程出错: {e}")
        return None
