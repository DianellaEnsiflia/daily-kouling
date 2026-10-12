"""
每日口诀邮件 - 供 GitHub Actions 云端每天21:30运行
即使本地关机也能发
"""
import os
import smtplib
try:
    import xlrd
except ImportError:  # 没有 xlrd 也能跑（走内置 fallback 口诀）
    xlrd = None
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

EMAIL_ADDRESS = os.environ.get("EMAIL_ADDRESS", "457016761@qq.com")
EMAIL_PASSWORD = os.environ.get("EMAIL_PASSWORD")  # 从云端密钥读取，不再写死
TO_EMAIL = os.environ.get("TO_EMAIL", "457016761@qq.com")
SMTP_SERVER = "smtp.qq.com"
SMTP_PORT = 465

XLS_PATH = os.path.join(os.path.dirname(__file__), "十大领域背诵口诀第3版教（下载后再打开）.xls")
# 如果文件在 E盘，也尝试
ALT_XLS = r"E:\资料\系统集成\十大领域背诵口诀第3版教（下载后再打开）.xls"

DOMAINS = ["整合管理","范围管理","进度管理","成本管理","质量管理","资源管理","沟通管理","风险管理","采购管理","干系人管理"]

BEISONG_PATH = os.path.join(os.path.dirname(__file__), "背诵清单.md")
# 十大领域过程口诀（随仓库走，云端也能读到；不再依赖本机 XLS）
KOULING_MD = os.path.join(os.path.dirname(__file__), "十大领域口诀.md")


def load_kouling_sections():
    """从《十大领域口诀.md》按 ## 标题切成 领域名 -> 正文"""
    if not os.path.exists(KOULING_MD):
        return {}
    secs, title, lines = {}, None, []
    for line in open(KOULING_MD, encoding="utf-8").read().splitlines():
        if line.startswith("## "):
            if title:
                secs[title] = "\n".join(lines).strip()
            title, lines = line[3:].strip(), []
        elif title is not None:
            lines.append(line)
    if title:
        secs[title] = "\n".join(lines).strip()
    return secs


def load_beisong_cards():
    """从《背诵清单》里切成一张张"今日背诵卡"（按 ## / ### 标题分段）。
    ★ 或 🎯 标记的重点卡会被放两次，出现频率翻倍。"""
    if not os.path.exists(BEISONG_PATH):
        return []
    text = open(BEISONG_PATH, encoding="utf-8").read()
    cards, title, lines = [], None, []
    for line in text.splitlines():
        if line.startswith("### "):
            if title and "\n".join(lines).strip():
                cards.append((title, "\n".join(lines).strip()))
            title, lines = line[4:].strip(), []
        elif line.startswith("## "):
            if title and "\n".join(lines).strip():
                cards.append((title, "\n".join(lines).strip()))
            title, lines = line[3:].strip(), []
        elif line.startswith("# "):
            continue
        elif title is not None:
            lines.append(line)
    if title and "\n".join(lines).strip():
        cards.append((title, "\n".join(lines).strip()))
    weighted = []
    for t, c in cards:
        weighted.append((t, c))
        if "★" in t or "🎯" in t:  # ★=老标记，🎯=新标记
            weighted.append((t, c))  # 重点卡加倍
    return weighted


def get_kouling_text(day_index):
    domain = DOMAINS[day_index % len(DOMAINS)]
    # 首选：仓库内置《十大领域口诀.md》（云端可用，内容最全）
    for title, body in load_kouling_sections().items():
        if domain in title and body:
            return domain, body
    try:
        path = XLS_PATH if os.path.exists(XLS_PATH) else ALT_XLS
        if xlrd is not None and os.path.exists(path):
            wb = xlrd.open_workbook(path)
            # 尝试找到对应sheet
            sheet_names = wb.sheet_names()
            # 第一页是整合，第二页可能是范围等，简化取对应索引
            idx = min(day_index % len(sheet_names), len(sheet_names)-1)
            sh = wb.sheet_by_index(idx)
            lines = []
            for r in range(min(15, sh.nrows)):
                vals = [str(sh.cell(r,c).value).strip() for c in range(min(5, sh.ncols))]
                if any(vals):
                    lines.append(" | ".join([v for v in vals if v]))
            content = "\n".join(lines[:10])
            if content:
                return domain, content
    except Exception as e:
        print(f"read xls fail: {e}")
    # 回退：返回口诀文字
    fallback = {
        0: "整合管理：立管议事组资产；专家人数议日程。（制定章程7过程）",
        1: "范围管理：规划范围-收集需求-定义范围-创建WBS-确认范围-控制范围",
        2: "进度管理：规划-定义活动-排序-估算-制定进度-控制进度",
        3: "成本管理：规划成本-估算-制定预算-控制成本",
        4: "质量管理：规划质量-管理质量-控制质量",
        5: "资源管理：规划-估算-获取-建设团队-管理团队-控制资源",
        6: "沟通管理：规划沟通-管理沟通-监督沟通",
        7: "风险管理：规划-识别-定性-定量-规划应对-实施应对-监督风险",
        8: "采购管理：规划采购-实施采购-控制采购",
        9: "干系人管理：识别-规划参与-管理参与-监督参与",
    }
    return domain, fallback.get(day_index % 10, domain)

def esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def build_html(domain, text, card=None):
    today = datetime.now().strftime("%Y年%m月%d日")
    day_idx = datetime.now().timetuple().tm_yday % 10
    card_html = ""
    if card:
        c_title, c_body = card
        card_html = f"""
<div style="padding:20px;border-top:2px dashed #ddd;">
<div style="font-weight:bold;color:#764ba2;font-size:16px;margin-bottom:8px;">🎯 今日背诵卡：{esc(c_title)}</div>
<div style="white-space:pre-wrap;line-height:1.7;background:#faf8ff;border-left:4px solid #764ba2;padding:12px;border-radius:4px;font-size:14px;">{esc(c_body)}</div>
</div>
"""
    return f"""
<html><body style="font-family:Microsoft YaHei;padding:20px;background:#f5f5f5;">
<div style="max-width:600px;margin:0 auto;background:white;border-radius:10px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.1);">
<div style="background:linear-gradient(135deg,#667eea,#764ba2);padding:20px;text-align:center;color:white;">
<h2 style="margin:0;">📚 每日口诀 - {domain}</h2>
<p style="margin:6px 0 0 0;opacity:0.9;">{today} · 十大领域 {day_idx+1}/10</p>
</div>
<div style="padding:20px;white-space:pre-wrap;line-height:1.6;">{esc(text)}</div>
{card_html}
<div style="padding:0 20px 20px 20px;color:#999;font-size:12px;border-top:1px solid #eee;margin-top:10px;">
<p>睡前5分钟，顺一遍口诀 + 今日背诵卡。电脑关机也能收到（云端发送）。</p>
</div>
</div>
</body></html>
"""

def send(domain, html, subject_suffix=""):
    if not EMAIL_PASSWORD:
        print("Missing EMAIL_PASSWORD env")
        return False
    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"📚 每日口诀 {domain}{subject_suffix} - {datetime.now().strftime('%m/%d')}"
    msg["From"] = EMAIL_ADDRESS
    msg["To"] = TO_EMAIL
    msg.attach(MIMEText("请用HTML查看", "plain", "utf-8"))
    msg.attach(MIMEText(html, "html", "utf-8"))
    server = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT)
    server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
    server.send_message(msg)
    server.quit()
    print(f"Sent {domain} to {TO_EMAIL}")
    return True

if __name__ == "__main__":
    idx = datetime.now().timetuple().tm_yday % 10
    # 也支持按日期轮换，避免每天相同
    domain, text = get_kouling_text(idx)
    cards = load_beisong_cards()
    card = cards[datetime.now().timetuple().tm_yday % len(cards)] if cards else None
    if card:
        print(f"今日背诵卡: {card[0]}（共{len(cards)}张卡轮换）")
    html = build_html(domain, text, card)
    send(domain, html, subject_suffix=f" + 背诵卡：{card[0][:12]}" if card else "")
