"""把 docs/index.html 报告发到邮箱。

邮件正文 = 报告「执行摘要」表格（富文本 HTML + 精简纯文本双版本），
完整报告作为附件：整份报告 HTML 不适合塞进邮件正文，也不适合拍平成纯文本。
配置写在下面的 MAIL 常量里（同名环境变量可覆盖，便于 Actions 用 Secrets）。
"""
import html as html_lib
import os
import re
import smtplib
import ssl
import sys
from datetime import datetime, timedelta, timezone
from email.header import Header
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from html.parser import HTMLParser

HERE = os.path.dirname(os.path.abspath(__file__))
REPORT = os.path.join(HERE, "docs", "index.html")
CST = timezone(timedelta(hours=8))

# ---- 邮箱配置：非敏感项写死，授权码只从环境变量 MAIL_PASS 读 ----
# 云端：GitHub 仓库 Settings → Secrets → Actions → 新建 MAIL_PASS，
#       workflow 已用 ${{ secrets.MAIL_PASS }} 注入（日志自动打码）。
# 本地：终端里临时设 $env:MAIL_PASS='授权码' 再运行；不配则跳过发信。
# 任何情况下都不要把授权码写进会被提交的文件。
SEND_EMAIL = True
# 用 `or 默认值` 而非 get(k, 默认值)：GitHub Actions 里未配置 Secrets 时，
# env 会被设成空字符串，get(k, default) 会拿到 "" 而误判。
MAIL = {
    "host": os.environ.get("MAIL_HOST") or "smtp.qq.com",
    "port": int(os.environ.get("MAIL_PORT") or 465),
    "user": os.environ.get("MAIL_USER") or "783109755@qq.com",
    "pass": os.environ.get("MAIL_PASS") or "",
    "to":   os.environ.get("MAIL_TO") or (
        "809080777@qq.com,1279885368@qq.com,120957262@qq.com,eliyan0919@qq.com,"
        "1013562640@qq.com"),
}
# -----------------------------------------------------------------


class _SummaryTable(HTMLParser):
    """抓「执行摘要」章节后的第一张表（维度 / 今日读数 / 倾向）。"""

    def __init__(self):
        super().__init__()
        self.rows = []
        self._h2 = False
        self._h2text = []
        self._after = False
        self._in_table = False
        self._done = False
        self._row = None
        self._cell = None

    def handle_starttag(self, tag, attrs):
        if tag == "h2":
            self._h2 = True
            self._h2text = []
        elif (self._after and tag == "table" and not self._in_table
              and not self._done):
            self._in_table = True
        elif self._in_table and tag == "tr":
            self._row = []
        elif self._in_table and tag in ("td", "th"):
            self._cell = []

    def handle_data(self, data):
        if self._h2:
            self._h2text.append(data)
        elif self._cell is not None:
            t = data.strip()
            if t:
                self._cell.append(t)

    def handle_endtag(self, tag):
        if tag == "h2":
            if "执行摘要" in "".join(self._h2text):
                self._after = True
            self._h2 = False
        elif tag in ("td", "th") and self._cell is not None:
            if self._row is not None:
                self._row.append("".join(self._cell).strip())
            self._cell = None
        elif tag == "tr" and self._row is not None:
            if any(self._row):
                self.rows.append(self._row)
            self._row = None
        elif tag == "table" and self._in_table:
            self._in_table = False
            self._done = True


def extract_summary(html):
    p = _SummaryTable()
    try:
        p.feed(html)
    except Exception:
        pass
    return p.rows


def _plain(fragment):
    """去掉内联标签、还原 HTML 实体。"""
    return html_lib.unescape(re.sub(r"<[^>]+>", "", fragment)).strip()


def extract_highlights(html):
    """抓报告里的「今日要点（30 秒版）」5 条 + 「一句话结论」整段。"""
    start = html.find("<p><strong>今日要点")
    end = html.find("<hr/>", start) if start >= 0 else -1
    block = html[start:end] if start >= 0 and end > start else ""
    points = [_plain(x) for x in re.findall(
        r"<p class='li'>(.*?)</p>", block, re.S)]
    m = re.search(
        r"<blockquote><strong>一句话结论</strong>(.*?)</blockquote>",
        block, re.S)
    conclusion = _plain(m.group(1)) if m else ""
    if conclusion.startswith("："):
        conclusion = conclusion[1:].strip()
    return points, conclusion


def build_text(rows, day, points, conclusion):
    """纯文本兜底版：要点 → 一句话结论 → 摘要，一行一条不堆表格。"""
    lines = ["ETF 日报 {}".format(day), ""]
    if points:
        lines.append("【今日要点（30 秒版）】")
        lines.extend(points)
        lines.append("")
    if conclusion:
        lines += ["【一句话结论】", conclusion, ""]
    lines.append("【执行摘要】")
    for r in rows[1:] if rows and len(rows[0]) >= 2 else rows:
        if len(r) >= 3:
            lines.append("· {}：{}（{}）".format(r[0], r[1], r[2]))
        elif len(r) >= 2:
            lines.append("· {}：{}".format(r[0], r[1]))
    lines += ["", "完整报告（18 章、9 张图表）见附件 "
              "etf_report_{}.html，浏览器打开即可。".format(day)]
    return "\n".join(lines)


def build_html(rows, day, points, conclusion):
    """邮件富文本版：要点 → 一句话结论 → 摘要表，全内联样式。"""
    esc = html_lib.escape
    th = ("padding:8px 10px;border:1px solid #dfe3e8;background:#f2f6fb;"
          "text-align:left;font-weight:600;color:#25476b")
    td = "padding:8px 10px;border:1px solid #e6e9ee;color:#333"
    head = rows[0] if rows else ["维度", "今日读数", "倾向"]
    body = []
    for i, r in enumerate(rows[1:] if len(rows) > 1 else []):
        cells = (r + ["", "", ""])[:3]
        bg = "#ffffff" if i % 2 == 0 else "#fafbfd"
        body.append(
            "<tr style='background:{}'><td style='{};font-weight:600'>{}</td>"
            "<td style='{}'>{}</td><td style='{}'>{}</td></tr>".format(
                bg, td, esc(cells[0]), td, esc(cells[1]), td, esc(cells[2])))
    thead = "<tr>{}</tr>".format("".join(
        "<th style='{}'>{}</th>".format(th, esc(c)) for c in head[:3]))

    sec = []
    if points:
        sec.append(
            "<h3 style='margin:20px 0 8px;color:#1d3f66;font-size:15px'>"
            "今日要点（30 秒版）</h3>"
            "<div style='background:#f7f9fc;border:1px solid #e6ebf2;"
            "border-radius:6px;padding:10px 14px;font-size:13px'>{}</div>".format(
                "".join("<p style='margin:4px 0'>{}</p>".format(esc(p))
                        for p in points)))
    if conclusion:
        sec.append(
            "<h3 style='margin:20px 0 8px;color:#1d3f66;font-size:15px'>"
            "一句话结论</h3>"
            "<div style='border-left:3px solid #3b7dd8;background:#f2f7ff;"
            "padding:10px 14px;font-size:13px;color:#2c3e50'>{}</div>".format(
                esc(conclusion)))
    sec.append(
        "<h3 style='margin:20px 0 8px;color:#1d3f66;font-size:15px'>"
        "执行摘要</h3>"
        "<table style='border-collapse:collapse;width:100%;font-size:13px'>"
        "{}{}</table>".format(thead, "".join(body)))

    return (
        "<meta charset='utf-8'>"
        "<div style='font:14px/1.7 -apple-system,BlinkMacSystemFont,"
        "'Microsoft YaHei',sans-serif;color:#222;max-width:760px'>"
        "<h2 style='margin:0 0 6px;color:#1d3f66;font-size:20px'>"
        "ETF 日报 {}</h2>"
        "<p style='margin:0 0 4px;color:#6b7684;font-size:13px'>"
        "详细数据与分析见附件</p>"
        "{}"
        "<p style='margin:18px 0 0;color:#6b7684;font-size:13px'>"
        "完整报告（含 18 章、9 张图表）见附件 <b>etf_report_{}.html</b>，"
        "下载后用浏览器打开即可。</p></div>").format(
            esc(day), "".join(sec), esc(day))


def main(argv):
    if not SEND_EMAIL:
        print("SEND_EMAIL = False，跳过发信")
        return 0
    to = [x.strip() for x in MAIL["to"].replace(";", ",").split(",")
          if x.strip()]
    if not (MAIL["host"] and MAIL["user"] and MAIL["pass"] and to):
        print("邮箱配置不完整（MAIL.host/user/pass/to），跳过发送")
        return 0
    if not os.path.exists(REPORT):
        print("报告不存在：{}".format(REPORT), file=sys.stderr)
        return 1

    with open(REPORT, encoding="utf-8") as f:
        html = f.read()
    day = datetime.fromtimestamp(os.path.getmtime(REPORT), CST).strftime(
        "%Y-%m-%d")
    subject = "ETF 日报 {}".format(day)
    rows = extract_summary(html)[:20]     # 防御：只要执行摘要表
    points, conclusion = extract_highlights(html)
    text = build_text(rows, day, points, conclusion)
    html_body = build_html(rows, day, points, conclusion)

    if "--dry-run" in argv:
        print("收件人:", ", ".join(to))
        print("主题:", subject)
        print("摘要表: {} 行 | 要点: {} 条 | 结论: {} 字".format(
            len(rows), len(points), len(conclusion)))
        print("--- 纯文本正文 ---")
        print(text)
        print("--- HTML 正文: {} 字符 ---".format(len(html_body)))
        print("--- 附件大小: {:.0f} KB ---".format(
            len(html.encode("utf-8")) / 1024))
        return 0

    msg = MIMEMultipart("mixed")
    msg["Subject"] = Header(subject, "utf-8")
    msg["From"] = MAIL["user"]
    msg["To"] = ", ".join(to)
    alt = MIMEMultipart("alternative")
    alt.attach(MIMEText(text, "plain", "utf-8"))
    alt.attach(MIMEText(html_body, "html", "utf-8"))
    msg.attach(alt)
    att = MIMEText(html, "html", "utf-8")
    att.add_header("Content-Disposition", "attachment",
                   filename="etf_report_{}.html".format(day))
    msg.attach(att)

    try:
        if MAIL["port"] == 465:
            server = smtplib.SMTP_SSL(
                MAIL["host"], MAIL["port"],
                context=ssl.create_default_context(), timeout=30)
        else:
            server = smtplib.SMTP(MAIL["host"], MAIL["port"], timeout=30)
            server.starttls(context=ssl.create_default_context())
        with server:
            server.login(MAIL["user"], MAIL["pass"])
            server.sendmail(MAIL["user"], to, msg.as_string())
    except Exception as exc:
        print("发送失败：{}: {}".format(type(exc).__name__, exc),
              file=sys.stderr)
        return 1
    print("已发送：{} → {} 个收件人（摘要 {} 行）".format(
        subject, len(to), len(rows)))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
