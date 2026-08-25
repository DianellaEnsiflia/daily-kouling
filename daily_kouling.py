"""
每日口诀邮件 - 供 GitHub Actions 云端每天21:30运行
即使本地关机也能发
"""
import os
import smtplib
import xlrd
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

def get_kouling_text(day_index):
    domain = DOMAINS[day_index % len(DOMAINS)]
    # 简化：先返回整合管理示例，完整版可按需扩展读取xls
    try:
        path = XLS_PATH if os.path.exists(XLS_PATH) else ALT_XLS
        if os.path.exists(path):
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

def build_html(domain, text):
    today = datetime.now().strftime("%Y年%m月%d日")
    day_idx = datetime.now().timetuple().tm_yday % 10
    return f"""
<html><body style="font-family:Microsoft YaHei;padding:20px;background:#f5f5f5;">
<div style="max-width:600px;margin:0 auto;background:white;border-radius:10px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.1);">
<div style="background:linear-gradient(135deg,#667eea,#764ba2);padding:20px;text-align:center;color:white;">
<h2 style="margin:0;">📚 每日口诀 - {domain}</h2>
<p style="margin:6px 0 0 0;opacity:0.9;">{today} · 十大领域 {day_idx+1}/10</p>
</div>
<div style="padding:20px;white-space:pre-wrap;line-height:1.6;">{text}</div>
<div style="padding:0 20px 20px 20px;color:#999;font-size:12px;border-top:1px solid #eee;margin-top:10px;">
<p>睡前5分钟，顺一遍口诀。电脑关机也能收到（云端发送）。</p>
</div>
</div>
</body></html>
"""

def send(domain, html):
    if not EMAIL_PASSWORD:
        print("Missing EMAIL_PASSWORD env")
        return False
    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"📚 每日口诀 {domain} - {datetime.now().strftime('%m/%d')}"
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
    html = build_html(domain, text)
    send(domain, html)
