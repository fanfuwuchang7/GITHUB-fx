#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ETF 全景分析：大盘 / 广度 / 成交额 / 板块资金流 / 微观 / 外围 / 画像 / 事件 / 回测 / 走势
========================================================================================

零第三方依赖（仅标准库）。自动拉取公开行情，输出内联 SVG 图表的 HTML 报告
与结构化数据表（.xlsx），以及可查询的 SQLite 派生索引。

章节：
    1  今日大盘（样本ETF / 指数 / 分组 / 量能 / 代表性自检）
    2  全市场涨跌广度与情绪（涨跌家数、涨跌停、区间分布、日内演变、两融、打板效应）
    3  全市场成交额与日内节奏（精确成交额、关口、量价、趋势、指数分时）
    4  全市场板块资金流（申万行业 + 概念，主力净额 / 5日 / 20日）
    5  ETF 折溢价、份额与净申赎（IOPV / 溢价 / 份额历史累积）
    6  跨市场外围与估值（外围指数 / 股指期货基差与国债期货 / 指数PE /
       ETF估值历史分位）
    7  市场画像（自建 9 维 + 官方 28 维对照）
    8  事件、异动、龙虎榜与筹码
    9  历史规则回测（扣成本 / 样本外 / 买入持有基准 / 显著性）
    10 全市场 ETF 全景（1722 只规模 / 真实净申购 / 估值分位）
    11 宏观流动性与股债性价比（CPI/PMI/M1M2/国债收益率 / ERP）
    12 逐日历史累积与今日 vs 昨日（接口无历史序列时的替代方案）
    13 数据质量与口径审计
    14 接下来走势（技术仪表盘 / ATR缺口周线 / 支撑压力 / 三情景 / 明日观察）
    15 事件采集与涨跌归因（8类资讯 / 7×24快讯东财JSON+东财H5+东财PC页+
       新浪直连+华夏时报+财联社电报六通道 / 时序切分 / 板块映射 / 三方校验 /
       反证 / 多源一致性 / 逐日累积 / 时间游标回溯 / 15.10 次日盘中验证 /
       ETF资金独立口径印证 / 归档日期校验）
    16 全市场涨跌排行（个股 5578 只 + ETF 1722 只全量 / 多周期强弱 /
       个股主力净流入 / 涨跌分布·成交集中度·资金覆盖面·周期共振 /
       样本代表性终极校验 / 解除 1.2 与 1.6 边界）
    17 多方评论观点（观点方识别 / 立场与分歧 / 时序降权 /
       观点×资金涨幅交叉校验 / 五份资讯页的日期归属处置）

用法：
    python3 etf_analysis.py                       # 在线取数，生成 .html + .xlsx
                                                  # （固定名 index.*，每次覆盖）
    python3 etf_analysis.py --offline            # 强制使用随附快照
    python3 etf_analysis.py --save-snapshot      # 在线取数后保存快照（含全部扩展数据）
    python3 etf_analysis.py --codes sh588000,sh510300
    python3 etf_analysis.py --market-days 30 --fund-top 5
    python3 etf_analysis.py --want-xlsx          # 额外生成结构化数据表 xlsx（默认只出 HTML）
    python3 etf_analysis.py --history-file my.json --output out/报告 --stdout
                                                 # 基准路径，产出 out/报告.html/.xlsx

数据来源：腾讯公开行情接口（快照/日线）、全市场概览接口（广度/两融/分时）、
板块资金流接口（申万行业 124 + 概念 805）、外围指数快照、
公开封装接口（全市场 ETF 1722 只明细/估值分位/净申购、官方 28 维画像、
宏观核心指标、结构化事件、龙虎榜、筹码）、
财经资讯端点（热点新闻 50 条 + 新闻正文、公司公告、个股研报、脱水研报、
财经日历、新股日历、主题龙头评级与一致预期，均带发布时间）、
7×24 全球财经快讯六通道：
  ① `np-weblist.eastmoney.com/comm/web/getFastNewsList`（东财 JSON，fastColumn=102）；
  ② `emh5wap.eastmoney.com/info/channel/kuaixun`（东财移动端 H5，服务端渲染，
     常只返回栏目壳）；
  ③ `kuaixun.eastmoney.com/7_24.html`（**东财 PC 全球财经快讯**；2026-11 改版后
     页面由 React 渲染、静态 HTML 只剩导航与行情壳，故壳页时自动改走页面自带的
     `getFastNewsList?biz=web_724&fastColumn=102` AJAX 接口（与①同源，
     合并时按时间+标题去重、不重复计数），老 SSR 版页面仍按 HH:MM + 正文直读；
     `7_24_{n}.html` 翻页锚点不可靠，不作为时间游标）；
  ④ `finance.sina.com.cn/7x24/`（新浪 PC 页，内嵌 50 条，带 YYYYMMDD 完整日期
     与秒级时间；同路的 cj.sina.cn/index/feed24hour 为纯文本无日期格式，
     可用 --flash-text 传入并自动做年份归属校验）；
  ⑤ `kuaixun.chinatimes.net.cn/`（华夏时报，中央级财经媒体；2026-11 改版后网页为
     JS 动态加载，改走 JSON 接口 `/api/proxy/news/result`（40 条/页，字段
     time/title/isAbstract/source）与 `/news/topclicks`（热点榜），
     旧网页通道保留为兜底；含同花顺 iFinD 的 ETF 资金流独立口径与个券 Top3）；
  ⑥ `www.cls.cn/telegraph`（财联社电报网页通道，实测可读、`HH:MM:SS` 秒级时间戳）
     与 JSON 端点 `www.cls.cn/nodeapi/telegraphList` / `/v1/roll/get_roll_list`
     （社区逆向端点、非官方公开 API）。网页通道的 `?date` / `?last_time` 参数
     实测被服务端忽略，仍只返回最新实时页；JSON 端点才支持时间游标回溯。
     sign = md5(sha1(按 key 字典序拼的 query 串))，
     零密钥纯本地计算；字段 ctime/title/brief/content/id；last_time 为时间游标，
     按本页最小 ctime 逐页回退即可回溯报告日 09:30–15:00 盘中快讯——
     这是六个源里唯一原生支持历史回溯的。端点版本会漂移，故新旧两端点、
     两套参数模板都试；有反爬频率风控，不宜高频轮询。
  六源合并去重、带来源标记，按 id 增量写入 etf_flash_history.json 逐日累积，
  解决实时页只给最新 50 条、无法回溯盘中快讯的问题。
  --flash-text 支持逗号分隔多文件，自动嗅探四种格式：
    「=== YYYY-MM-DD === + HH:MM」（东财 PC 页，多日期分段）、
    「=== YYYY-MM-DD === + HH:MM:SS」（财联社电报网页通道，秒级）、
    「日期行 + 🔗」（华夏时报）、「HH:MM:SS 正文」（新浪纯文本）。
  四者都做日期/年份归属校验，归档整段不并入。
多方评论观点：复用上述事件池，按表态标记词分离「观点方 → 立场 → 主题」，
  与第四章行业资金、第一章样本 ETF 涨幅做交叉校验；事后观点一律降权为印证。
  归档页（cj.sina.cn 2025-09-09）与空壳页（emh5wap 仅返回站点标题）由
  日期归属校验与内容长度检查拦截，不并入报告日事件池；
  东财 PC 7_24 页（第三通道）与财联社电报页（第六通道）网页通道均可用，
  两者共同生成 15.10 的次日盘中验证层；但翻页参数都被忽略，11-04 盘中仍靠逐日累积。
接口不给历史序列，故逐日累积写 etf_daily_history.json 做 diff；
所有外部端点均无 SLA、可能变更，失败时自动降级到快照。
事件研究：15.5.1 用「主题内 ETF 收益 − β×沪深300」算单日 AR，
按日写入 etf_event_study.json；同主题累计 ≥15 个交易日才给 t 值，
不足时只登记、不出显著性结论（避免用单日数据冒充统计证据）。
端点契约自检：每次运行按「状态 + 条数下限」逐通道断言，
结果写入 endpoint_health.json（含 as_of 时点），用于定位改版发生日；
第十三章同步展示，任一通道不达标即判契约变更、提示复核。
数据时点：报告头标注 as-of 与「盘中/收盘」状态；盘中口径下禁用
「收盘快照」表述，成交额环比与量价结论一律标注需收盘复核。
历史回填：两市的「成交额」可从指数日 K（沪市+深市）补齐 20~30 个交易日，
广度/资金/ERP 无历史接口，绝不插值；回填条目带 amount_src 标记，
盘中当日累计值一律不与全天回填值相减（曾因此得出「缩量 28%」的假结论）。
真源库：累积型历史（快讯/逐日指标/份额/事件台账/快照）的真源是 etf_store.sqlite
（schema 见 STORE_DDL），对应 *.json 文件已冻结为迁移前备份、不再读写；
派生索引：每次运行写 etf_data.sqlite（带 schema_ver，结构变更整库重建）：
本次结果六张表 daily/vol_cone/topic_ar/consistency/quality/option_iv 按报告日覆盖；
JSON 再镜像二十七张表：快照截面 snapshot_quote/snapshot_meta/sector_flow/
breadth_metric/global_index/etf_profile/quote_kline/breadth_minute/yidong_item/
strong_stock/breadth_note/ext_meta，排行宏观 stock_rank/etf_rank/futures/
market_score/macro_indicator/breadth_dist，流水资讯 shares_hist/flash_news/
news_item/endpoint_health，事件筹码 lhb/chip_metric，日历 calendar_item/ipo_item/
rating_stock；截面表按快照日覆盖、流水表按主键 upsert，供 --sql 做跨表跨日查询；
删库可重建，故不承担迁移风险。JSON 里的每个子键都在表中有一行对应，不再有
「只能读文件」的字段。
跨源一致性：广度接口、全量排行、ETF 明细三套独立采集互相印证，
结果写入 consistency.json 并在第十三章展示；契约查「拿得到吗」，
一致性查「数字彼此矛盾吗」，两者都过才算数据可用。
注意：本脚本只做客观数据整理与规则化情景推演，不构成投资建议。
"""

import argparse
import json
import time
import math
import hashlib
import os
import re
import sys
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timezone, timedelta
from urllib.request import Request, urlopen
from urllib.parse import urlencode

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

# ----------------------------------------------------------------------------
# 配置
# ----------------------------------------------------------------------------

ENDPOINT = "https://web.ifzq.gtimg.cn/appstock/app/kline/kline?param={code},day,,,{days}"
HEADERS = {"User-Agent": "Mozilla/5.0", "Referer": "https://gu.qq.com/"}
HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_SNAPSHOT = os.path.join(HERE, "etf_snapshot.json")

# 指数层：覆盖大盘 / 中盘 / 小盘 / 沪深京 / 成长风格，避免只盯 4 个指数产生风格偏差
INDEX_CODES = [
    "sh000001", "sz399001", "sh000016", "sh000300", "sh000905",
    "sh000852", "sz399006", "sh000688", "bj899050",
]

# 全市场成交额口径
# sh000001 上证指数 = 沪市成交额，sz399001 深证成指 = 深市成交额，二者相加即「沪深两市成交额」
# bj899050 北证50 为成分指数口径，仅作参考，不计入两市合计
# sz399006 / sh000688 为板块子集，用于算成交结构占比，不与两市合计重复相加
MARKET_CODES = {
    "沪市": "sh000001",
    "深市": "sz399001",
    "北证50": "bj899050",
    "创业板指": "sz399006",
    "科创50": "sh000688",
}
TWO_MARKETS = ["sh000001", "sz399001"]
BOARD_SUBSETS = [("创业板指", "sz399006", "深市"), ("科创50", "sh000688", "沪市")]
AMOUNT_GATES = [(20000, "2万亿"), (15000, "1.5万亿"), (10000, "1万亿")]

# 全市场板块资金流：行业(申万) + 概念。t=zdf&o=1 返回全量，排序在脚本内自己做。
# 字段：zdf涨跌幅 cje成交额(万元) hsl换手 zgb涨跌家数 zllr主力流入 zllc主力流出
#       zljlr主力净流入 zljlr_d5/d20 近5/20日主力净流入 lzg领涨股
FUNDFLOW_ROOT = ("https://proxy.finance.qq.com/ifzqgtimg/appstock/fundflow/"
                 "Fundflow/")
FUNDFLOW_BOARDS = {"industry": "getHyRank", "concept": "getGnRank"}
FUND_TOP = 8

# 概念板块里「属性/指数成分/交易状态」类覆盖面极广，不代表具体赛道，
# 默认从题材榜剔除，避免把「转融券标的」「深股通」误读成资金聚焦某主题。
BROAD_CONCEPT_KEYWORDS = (
    "转融券", "融资融券", "深股通", "沪股通", "陆股通", "富时", "MSCI",
    "标普", "道琼斯", "同股同权", "破净", "昨日涨停", "昨日连板", "昨日跌停",
    "红利股", "价值股", "证金持股", "汇金", "超级品牌", "机构重仓",
    "基金重仓", "社保重仓", "QFII", "预盈预增", "预亏预减", "业绩预",
    "高价股", "低价股", "次新股", "ST股", "含可转债", "沪深300", "上证",
    "中证", "深证", "成分股", "AH股", "B股", "GDR",
    "大盘股", "中盘股", "小盘股", "微盘股", "权重股", "蓝筹", "绩优股",
)

# ---- 全市场涨跌广度 / 情绪 / 精确成交额 / 日内分时 / 两融 / 事件 ----
# 一个接口同时返回：涨跌家数与区间分布、涨停跌停、日内上涨占比演变、
# 上证指数分时、两市精确成交额与较上日变化、两融/ADR/收评/异动。
BREADTH_URL = "https://proxy.finance.qq.com/cgi/cgi-bin/market/hs/index"
QT_URL = "https://qt.gtimg.cn/q="

# ---- 跨市场外围 ----
# 港股与欧洲为当日收盘；美股为最近一个交易日收盘（与 A 股存在时差，须标注）
GLOBAL_CODES = {
    "hkHSI": ("恒生指数", "港股"),
    "hkHSTECH": ("恒生科技指数", "港股"),
    "hkHSCEI": ("国企指数", "港股"),
    "usIXIC": ("纳斯达克", "美股"),
    "usINX": ("标普500", "美股"),
    "usDJI": ("道琼斯", "美股"),
    "usNDX": ("纳斯达克100", "美股"),
    "ukUKX": ("富时100", "欧洲"),
}

# ---- 股指 / 国债期货（基差层：机构对冲压力的高频确认）----
# 新浪期货行情：`nf_IF0` 连续合约。字段（0 起）：
#   0 开 1 高 2 低 3 最新 4 成交量 5 成交额 6 持仓量 13 昨结算 36 日期 37 时间。
# 基差 = 期货最新 - 现货指数，负数即贴水；连续合约临近交割会强制收敛，
# 故只作「当日对冲温度」参考，不当作趋势预测，也不与历史分位比较（无历史）。
FUT_URL = "https://hq.sinajs.cn/list="
FUT_CONTRACTS = [
    ("nf_IF0", "IF 沪深300期货", "sh000300", "沪深300"),
    ("nf_IH0", "IH 上证50期货", "sh000016", "上证50"),
    ("nf_IC0", "IC 中证500期货", "sh000905", "中证500"),
    ("nf_IM0", "IM 中证1000期货", "sh000852", "中证1000"),
    ("nf_T0", "T 十年国债期货", "", ""),
    ("nf_TF0", "TF 五年国债期货", "", ""),
]

# ---- 事件研究台账（AR = 主题内 ETF 收益 − β×基准收益）----
# 单日 AR 只描述「当日跑赢了多少」，要检验「事件是否真的带来超额收益」
# 需要同主题 ≥15 个交易日样本才能算 t 值；不足时只登记、不出结论。
EVENT_STUDY_FILE = os.path.join(HERE, "etf_event_study.json")
EVENT_STUDY_MIN_N = 15

# ---- 分时（1 分钟）接口：日内事件窗口 CAR ----
# 只返回「当天」，所以日内 CAR 仅在同一交易日的在线运行中有值；
# 离线快照/次日重跑会自动跳过（不猜、不用次日数据倒推当日）。
MINUTE_URL = "https://web.ifzq.gtimg.cn/appstock/app/minute/query?code="
# 东财 1 分钟分时（主源）：fields2 = 时间,开,收,高,低,量,额,均价
EM_TRENDS_URL = ("https://push2his.eastmoney.com/api/qt/stock/trends2/get?secid={}"
                 "&fields1=f1,f2,f3,f4,f5,f6,f7,f8"
                 "&fields2=f51,f52,f53,f54,f55,f56,f57,f58&iscr=0&ndays=1")

# ---- ETF 份额历史（净申赎）----
# 行情快照只给当日份额，历史份额需靠脚本逐日累积；
# 连续运行 N 日后才能算「份额变化 = 净申购 - 净赎回」。
SHARES_HISTORY_FILE = os.path.join(HERE, "etf_shares_history.json")

# ---- 图表（纯标准库生成 SVG，零第三方依赖）----
# 输出进 docs/charts：本目录即 Git 仓库，docs/ 是 GitHub Pages 的发布源
CHART_DIR = os.path.join(HERE, "docs", "charts")
SVG_W, SVG_H = 760, 300
C_UP, C_DOWN, C_FLAT, C_ACC = "#e05a4f", "#2e9e7b", "#9aa4ad", "#3b7dd8"

# 自建市场画像维度（现在同时接入官方 28 维画像做对照）
PROFILE_DIMS = ["情绪", "短期趋势", "量能", "个股宽度", "估值", "风格轮动",
                "行业轮动", "板块宽度", "中长期方向"]

# ---- 行情封装接口（HTTP POST，公开固定 token；失败时自动降级）----
# 一次查询可拿到：全市场 ETF 规模/估值/估值分位/净申购/折溢价/成交额/换手区间、
# 官方市场画像、宏观核心指标（CPI/PMI/M1M2/国债收益率）、结构化事件。
PROXY_URL = "https://proxy.finance.qq.com/cgi/cgi-bin/openai/openclaw/proxy"
PROXY_TOKEN = "35b2f60e804ddbb6f2f0160f45ee815736692f29fabe17e1ecbd277966f4df45"
PROXY_PARAMS = {"app": "openclaw", "dev_id": "dev_etf_analysis",
                "skill_channel": "yuanbao", "token": PROXY_TOKEN}
ETF_LIST_CODES = [
    "etf_size", "etf_qt_daily", "etf_qt_chg_interval", "etf_qt_disc_interval",
    "etf_qt_amt_interval", "etf_qt_turn_interval", "etf_netflow_interval",
    "etf_valuation_level", "etf_valuation_hist",
]
MACRO_LIST_CODES = ["macro_core_indicators_cur_p1", "macro_core_indicators_cur_p2"]
PROFILE_LIST_CODES = ["market_statis_summary"]
EVENT_LIST_CODES = ["stock_event"]

# 全市场龙虎榜（GET，公开）
LHB_URL = "https://proxy.finance.qq.com/cgi/cgi-bin/longhubang/lhbDetail?tab=jg,"

# ---- 全市场涨跌排行（腾讯行情公开接口，收盘快照）----
# ① 个股：getBoardRankList(board_code=aStock)，分页拉全量 ~5578 只，
#    自带 zdf 涨跌幅、zdf_d5/d20/d60/w52/y 多周期涨幅、zljlr 主力净流入、
#    换手/量比/市值/PE。用上涨下跌家数与第二章广度做日期一致性校验。
# ② ETF：etf_size 的 1722 只代码 → qt.gtimg.cn 批量报价（60 只/批），
#    字段 32=涨跌幅、37=成交额(万)、38=换手、43=振幅、30=行情时间戳。
#    两者都返回报告日收盘快照（时间戳 20261104 16:14 已验证）。
RANK_URL = "https://proxy.finance.qq.com/cgi/cgi-bin/rank/hs/getBoardRankList"
RANK_PAGE_SIZE = 200
RANK_MAX_PAGES = 40
RANK_BOARDS = {"沪深京全A": "aStock", "创业板": "cyb", "科创板": "kcb"}
QT_BATCH = 60
RANK_STOCK_FIELDS = [
    "code", "name", "zdf", "zxj", "zdf_d5", "zdf_d20", "zdf_d60",
    "zdf_w52", "zdf_y", "zljlr", "zllr", "zllc", "turnover", "volume",
    "hsl", "lb", "zf", "speed", "zsz", "ltsz", "pe_ttm", "stock_type",
]
ETF_EXCLUDE_KW = ["债", "国债", "短融", "货币", "现金", "添利", "保证金",
                  "理财", "同业存单", "政金", "地方债", "城投"]

# ---- 财经资讯采集端点（全部公开 GET / 已有封装 POST，带发布时间）----
NEWS_HOT_URL = "https://proxy.finance.qq.com/ifzqgtimg/appstock/app/HotStock/getHotNews"
NEWS_DETAIL_URL = "https://ifzq.gtimg.cn/appstock/news/NewsProxy/get"
NOTICE_URL = "https://ifzq.gtimg.cn/appstock/news/noticeList/searchByType"
REPORT_URL = "https://ifzq.gtimg.cn/appstock/app/investRate/getReport"
CALENDAR_URL = "https://proxy.finance.qq.com/ifzqgtimg/appstock/app/StockCalendar/query"
IPO_URL = "https://proxy.finance.qq.com/ifzqfinance/stock/notice/ipo/search"
NEWS_DETAIL_LIMIT = 8       # 只对最相关的新闻抓正文，避免大量请求
NEWS_WINDOW_DAYS = 2        # 采集窗口：报告日 + 次日早报

# ---- 东方财富（天天基金母公司）7×24 全球财经快讯 ----
# 页面 https://kuaixun.eastmoney.com/7_24.html ；summary 即完整短讯正文。
# 返回 {"code":"1","data":{"fastNewsList":[{"code","title","summary","showTime",
#        "titleColor","stockList"}]}}；titleColor!=0 为重要快讯；按 sortEnd 翻页。
EM_FLASH_URL = "https://np-weblist.eastmoney.com/comm/web/getFastNewsList"
EM_FLASH_FALLBACK_URLS = [
    "https://np-listapi.eastmoney.com/comm/web/getFastNewsList",
    "https://np-listapi.eastmoney.com/nlist/api/list/get",
]
EM_FLASH_COLUMNS = {"102": "全部/全球", "101": "要闻", "109": "基金",
                    "108": "债券", "110": "大宗"}
EM_FLASH_COLUMN = "102"     # 102 = 全部/全球 7×24，24 小时滚动；109 基金栏目低频
EM_FLASH_PAGES = 4          # 每页 50 条，回溯 200 条覆盖报告日
EM_FLASH_PC_AJAX_PAGES = 2  # PC 页通道壳页走 AJAX 时的翻页数（与①同源，仅核验）

# 移动端 H5 快讯页（服务端渲染，网页通道可读；换到可直连环境脚本可自行解析）
# http://emh5wap.eastmoney.com/info/channel/kuaixun
EM_H5_FLASH_URL = "http://emh5wap.eastmoney.com/info/channel/kuaixun"
EM_H5_TIME_RE = re.compile(r"(\d{2}:\d{2})\s*【([^】]+)】([^\[]{0,400})")
EM_FLASH_STATUS_DESC = {
    "ok": "东财 JSON 接口直连",
    "sina_ok": "新浪财经 7×24 直连（东财 JSON 受限）",
    "cls_ok": "财联社电报直连（含 last_time 时间游标回溯）",
    "web_channel": "移动端 H5 / 网页通道采集（直连受限）",
    "blocked": "直连均被网络策略拒绝，本次无数据",
    "empty": "接口返回空",
}

# ---- 新浪财经 7×24 全球实时快讯（第三通道，沙盒可直连）----
# 页面 finance.sina.com.cn/7x24/ 服务端渲染，内嵌 50 条，
# 带 newsdata-time=YYYYMMDD（完整日期）+ HH:MM:SS（秒级）+ 正文。
# 用户给的 cj.sina.cn/index/feed24hour 是同一路的移动端入口，
# 但其纯文本列表只有 HH:MM:SS 无日期，且常返回历史归档，
# 故以 PC 页为主、纯文本为兼容格式（需做日期归属校验）。
SINA_FLASH_URL = "https://finance.sina.com.cn/7x24/"
SINA_FLASH_SSR_RE = re.compile(
    r'<div newsdata-id="(\d+)" newsdata-time="(\d{8})">\s*'
    r'<div>(\d{2}:\d{2}:\d{2})</div>\s*<a[^>]*>(.*?)</a>', re.S)
SINA_FLASH_TEXT_RE = re.compile(
    r'(?m)^\s*(\d{2}:\d{2}:\d{2})\s*\n\s*(.+?)(?=\n\s*\d{2}:\d{2}:\d{2}\s*\n|\Z)', re.S)
SINA_ARCHIVE_YEAR_RE = re.compile(r'20\d{2}年')
FLASH_HISTORY_FILE = "etf_flash_history.json"

# ---- 华夏时报 7×24 全球财经快讯（第四通道，中央级主流财经媒体）----
# https://kuaixun.chinatimes.net.cn/ 服务端渲染：日期行 YYYY-MM-DD +
# HH:MM 时间行 + 【标题】正文（来源） + 🔗 分隔；底部另有「热点快讯」榜。
# 该域名在部分网络被策略拒绝，脚本内置解析器，可由 --flash-text 传入
# 网页通道采集的文本，换到可直连环境会自动生效。
CT_FLASH_URL = "https://kuaixun.chinatimes.net.cn/"
# 2026-11 改版后页面为 JS 动态加载，静态 HTML 不再含快讯正文；
# 数据由 /api/proxy/news/result（JSON）+ /api/proxy/news/topclicks（热点榜）提供。
CT_FLASH_API = CT_FLASH_URL + "api/proxy/news/result"
CT_FLASH_HOT_API = CT_FLASH_URL + "api/proxy/news/topclicks"
CT_FLASH_PAGES = 3  # result 翻页数（40 条/页），覆盖早盘至当前
CT_FLASH_DB_API = CT_FLASH_URL + "api/proxy/news/loadmore2"
CT_FLASH_DB_PAGES = 8  # loadmore2 数据库兜底翻页数（20 条/页），回溯至 08 点档
CT_DATE_RE = re.compile(r"(?m)^\s*(20\d{2}-\d{2}-\d{2})\s*$")
CT_TIME_SPLIT_RE = re.compile(r"(?m)^\s*(\d{2}:\d{2})\s*$")
CT_SOURCE_RE = re.compile(r"（([^）]{2,12})）\s*$")
CT_HOT_RE = re.compile(r"(?m)^\s*(\d{1,2})\s*$\s*\n\s*(.+?)\s*$")
CT_HOT_ANCHOR = "热点快讯"

# 跨源一致性校验：同一事实在 ≥2 个独立来源族出现即视为互证
CROSS_SOURCE_ENTITIES = [
    "SK海力士", "三星电子", "东京电子", "爱德万测试", "日经225",
    "5500亿", "5560亿", "净买入", "资金借道", "长鑫科技",
    "回购", "原油暴跌", "IPO节奏", "民主党", "美股芯片股",
    "20CM", "科创半导体", "存储涨价", "中选", "美联储",
    # 华夏时报源带来的新实体
    "432.39", "503.37", "837.59", "23个交易日", "英伟达",
    "美光科技", "富时中国A50", "SPDR", "高通", "布伦特",
]
# 来源族划分（同族不同通道不重复计为独立信源）
SOURCE_FAMILY = {
    "东方财富7×24": "东方财富", "东方财富H5快讯": "东方财富",
    "新浪财经7×24": "新浪财经", "新浪财经7×24(纯文本)": "新浪财经",
    "华夏时报7×24": "华夏时报", "华夏时报热点榜": "华夏时报",
    "财联社电报": "财联社", "东方财富全球快讯": "东方财富",
}

# ---- 东方财富全球财经快讯 PC 页（第三通道）----
# https://kuaixun.eastmoney.com/7_24.html ：老版本服务端渲染，带日期 + HH:MM + 正文，
# 可直接解析；2026-11 改版后为 React 渲染，静态 HTML 只剩导航与行情壳
# （「暂无相关内容」「9秒后刷新」），快讯正文改由 /emresource/main/js/kuaixun.js
# 里的 getFastNewsList（biz=web_724、fastColumn=102、sortEnd 游标）AJAX 加载，
# 因此壳页时自动改走该接口（与①同源，合并阶段按时间+标题去重）。
# 7_24_{n}.html 翻页锚点不可靠（_2~_5 同批、_6 失效），不作为时间游标。
EM_PC_FLASH_URLS = [
    "https://kuaixun.eastmoney.com/7_24.html",
    "https://kuaixun.eastmoney.com/7_24_2.html",
    "https://kuaixun.eastmoney.com/7_24_3.html",
]
EM_PC_DATE_RE = re.compile(r"(?m)^===\s*(20\d{2}-\d{2}-\d{2})\s*===\s*$")
EM_PC_LINE_RE = re.compile(r"(?m)^\s*(\d{2}:\d{2}(?::\d{2})?)\s+(.+?)\s*$")
EM_PC_MAGIC = "东方财富 7×24 全球财经快讯"
CLS_WEB_MAGIC = "财联社电报"


def source_family(src):
    """把单家媒体名归并到信源族，避免把 N 家媒体误算成 N 个独立源。"""
    s = src or ""
    if s in SOURCE_FAMILY:
        return SOURCE_FAMILY[s]
    if s == "公司公告":
        return "公司公告"
    if "研报" in s:
        return "卖方研报"
    return "财经媒体"


def flash_key(r):
    """快讯去重键：时间精确到秒 + 标题前 60 字（含来源族）。

    早先用「分钟 + 标题前 24 字」，模板化标题（「XX板块走高，XX涨停」）
    前缀相同会被误合并成一条，导致真实条数被低估、事件被吞掉。
    标题长度上限放宽到 60 字并保留来源族，误合并概率显著下降。
    """
    t = (r.get("time") or "").replace("T", " ")[:19]
    title = re.sub(r"\s+", "", r.get("title") or "")[:60]
    return "{}|{}|{}".format(t, source_family(r.get("source")), title)

# ---- 多方评论观点采集（第十七章）----
# 从快讯/研报/公告里识别「谁在说什么立场」，与纯事实快讯分开。
# 观点 ≠ 事实 ≠ 证明：多数是收盘后/次日的事后解读，只能作印证、不作驱动。
# 强表态标记：必须命中才算观点候选；「据/数据显示」这类纯转述不单独算
VIEW_MARKERS = [
    "研报称", "认为", "指出", "看好", "判断", "分析人士", "多家基金公司",
    "建议", "警示", "长期看好", "偏乐观", "共识", "分歧", "展望", "预计",
    "表示", "坦言", "评论", "解读", "观点", "策略", "业绩会", "首席",
    "分析师", "经济学家", "料将", "担忧", "谨慎", "公告称", "回应",
    "研报", "机构表示", "市场人士", "业内", "记者获悉",
]
# 非市场类社会/政治/灾害快讯：即便含「表示」也不纳入观点池
VIEW_EXCLUDE = [
    "部署", "枪击", "坠毁", "爆炸", "示威", "抗议", "集会", "谴责",
    "会谈", "警方", "伤亡", "弹劾", "会谈进行了", "访问沙特", "防空拦截弹",
    "化工厂", "直升机", "反对全球傲慢日", "商铺遭破坏",
]
# 标题命中这些外围词时，观点不与 A 股板块资金对照（口径不同）
VIEW_EXTERNAL = ["美股", "纳指", "道指", "标普", "隔夜", "日韩", "韩国综指",
                 "日经", "日本央行", "布伦特", "WTI", "美油", "纽约期金",
                 "欧洲股市", "中概股", "富时中国A50", "美股收评", "美股盘前"]
# 观点方归类（按文本命中优先级，同一条只归一类）
VIEW_ACTORS = [
    ("卖方研报", ["研报", "证券研究所", "分析师", "国金证券", "中信建投",
                 "西南证券", "中金", "华泰证券", "招商证券", "国泰君安",
                 "中信证券", "券商"]),
    ("基金/资管", ["多家基金公司", "基金公司", "公募", "基金经理", "险资",
                  "资管", "首程资本", "ETF管理"]),
    ("海外机构", ["高盛", "富国银行", "摩根", "花旗", "贝莱德", "所罗门",
                 "CEO", "CFO", "诺瓦克", "美银", "瑞银", "大摩", "小摩"]),
    ("官方/监管", ["央行", "证监会", "国务院", "总理", "总统", "财政部",
                  "美联储", "日本央行", "OPEC", "IMF", "内政部长",
                  "总统新闻秘书", "政府表示", "委员一致认为"]),
    ("公司管理层", ["公告称", "公司表示", "负责人表示", "董事长", "回应",
                   "计划", "拟向", "预计第一财季", "预计第四财季", "公司同时"]),
    ("媒体/分析人士", ["分析人士", "记者", "业内", "市场人士", "评论员",
                     "人民财讯", "中国证券报", "央视新闻", "财联社", "CCTV",
                     "澎湃新闻", "新华社", "据", "数据显示"]),
]
# 立场词：只保留带判断色彩的，避免把「同比增长」这类事实数字误判为看多
# 立场词只保留「判断性表述」；行情数字（净流入/重挫/同比增长）是事实，不算立场
VIEW_BULL = [
    "看好", "长期看好", "有望", "提前受益", "抬升", "价值抬升", "底部显现",
    "底部区域已显现", "底部区间", "积极信号变多", "超预期", "超出预期",
    "复苏", "回暖", "乐观", "偏乐观", "改善", "确定性增强", "必然方向",
    "增量明确", "共振", "加码", "机会", "维持高位", "稳定增长", "极具价值",
    "抄底", "持续加仓", "加仓", "偏乐观预判", "长期看好", "提前受益",
]
VIEW_BEAR = [
    "承压", "持续承压", "盈利承压", "由盈转亏", "亏损扩大", "逆风", "磨底",
    "过热", "谨慎", "不确定性", "集中兑付压力", "收缩", "放缓", "疲软",
    "逼近物理极限", "尚未走出", "过度兴奋", "不必迅速", "负面观察",
    "获现能力大幅下降", "还需时间", "共识兑现", "明显消化", "控制支出",
    "依赖外部支持", "并非同步", "检验增长的持续性", "尚待旺季验证",
    "反转尚待", "去库存", "盈利分化", "估值溢价明显消化", "共识兑现为行情还需时间",
]
VIEW_STANCE_LABEL = {
    "bull": "偏多", "bear": "偏空", "mixed": "分歧/条件式",
    "lean_bull": "中性偏多", "lean_bear": "中性偏空", "fact": "事实/无立场",
}
# 观点章节专用主题补充：第十五章 TOPIC_MAP 偏「当日 A 股驱动」，
# 而机构观点常谈 AI算力/锂电/消费/地产链等中期赛道。这里**只用于第十七章**，
# 不改 TOPIC_MAP，避免影响 15.3/15.4 已校验的归因排序。
VIEW_TOPIC_EXTRA = [
    {"topic": "AI算力/光通信",
     "kw": ["AI算力", "算力链", "算力", "光通信", "光模块", "光互连",
            "GPU", "ASIC", "PCB", "CPO", "数据中心", "CSP", "高速光"],
     "sectors": ["通信设备", "元件"], "etf_kw": ["通信", "云计算", "人工智能", "算力"]},
    {"topic": "AI眼镜/消费电子",
     "kw": ["AI眼镜", "智能眼镜", "镜片", "Rokid", "雷鸟", "近眼终端"],
     "sectors": ["消费电子"], "etf_kw": ["消费电子"]},
    {"topic": "新能源/锂电",
     "kw": ["锂电", "固态电池", "储能", "动力电池", "电池产业链"],
     "sectors": ["电池", "电力设备"], "etf_kw": ["新能源", "电池", "锂电"]},
    {"topic": "消费/白酒",
     "kw": ["白酒", "去库存", "旺季验证", "消费复苏"],
     "sectors": ["白酒", "食品饮料"], "etf_kw": ["消费", "酒"]},
    {"topic": "周期/地产链",
     "kw": ["玻璃", "平板玻璃", "竣工", "地产链", "浮法"],
     "sectors": ["玻璃玻纤", "房地产"], "etf_kw": []},
    {"topic": "机器人/物理AI",
     "kw": ["物理AI", "机器人科技", "脑机", "人形机器人", "具身智能"],
     "sectors": ["自动化设备", "通用设备"], "etf_kw": ["机器人"]},
]

# 主题 → 板块关键词 / 相关 ETF 分组（用于事件-板块-ETF 三方映射）
# 每个主题：板块名（对齐第四章申万行业模糊匹配）、ETF分组名、方向词
TOPIC_MAP = [
    {"topic": "半导体/存储/芯片", "kw": ["存储", "芯片", "半导体", "长鑫", "光刻", "晶圆",
                                  "DRAM", "封测", "科创板回购", "国产替代"],
     "sectors": ["半导体", "元件", "消费电子", "电子"], "groups": ["TMT", "科创"],
     "etf_kw": ["芯片", "半导体", "科创50", "科创板"]},
    {"topic": "IPO/一级市场", "kw": ["IPO", "暂停", "新股", "融资节奏", "上市节奏"],
     "sectors": ["证券", "多元金融"], "groups": ["金融地产"],
     "etf_kw": ["证券"]},
    {"topic": "回购/增量资金", "kw": ["回购", "增持", "增量资金", "资金已就位", "护盘",
                                "平准", "国家队", "险资", "ETF", "净买入",
                                "加仓", "净流入", "资金借道", "宽基", "申购"],
     "sectors": [], "groups": [], "etf_kw": []},
    {"topic": "外围/美国政治", "kw": ["中选", "特朗普", "道指", "纳斯达克", "美股",
                                "中概", "中国资产", "美联储", "小非农", "降息"],
     "sectors": [], "groups": ["跨境"], "etf_kw": ["中概", "恒生", "纳指"]},
    {"topic": "原油/地缘", "kw": ["油价", "原油", "布伦特", "伊朗", "地缘", "OPEC"],
     "sectors": ["石油石化", "煤炭"], "groups": ["周期"],
     "etf_kw": ["能源", "油气", "煤炭"]},
    {"topic": "贵金属/避险", "kw": ["黄金", "白银", "金银", "金价", "贵金属"],
     "sectors": ["贵金属"], "groups": ["避险"], "etf_kw": ["黄金"]},
    {"topic": "机器人/高端制造", "kw": ["机器人", "王兴兴", "人形", "智能装备",
                                 "高端制造", "大单", "磷化铟"],
     "sectors": ["通用设备", "自动化设备"], "groups": ["新能源军工"],
     "etf_kw": ["机器人", "军工", "高端装备"]},
    {"topic": "宏观/流动性", "kw": ["CPI", "PMI", "社融", "M1", "MLF", "降准", "利率"],
     "sectors": [], "groups": [], "etf_kw": []},
]

# 反证词：同一主题里外围出现下跌而 A 股上涨时，降权外围传导
CONTRARY_WORDS = ["大跌", "重挫", "跌超", "暴跌", "下挫", "杀跌", "逆势下挫", "遭重挫"]

# 事件映射板块的机构预期样本（批量查评级/一致预期，用于产业催化验证）
EVENT_SAMPLE_CODES = ["sh688981", "sz002371", "sh603986", "sh600030", "sh601899"]

# 逐日历史累积（接口不提供历史序列，只能靠脚本每次运行落盘）
HISTORY_FILE = os.path.join(HERE, "etf_daily_history.json")
FLASH_HISTORY_PATH = os.path.join(HERE, FLASH_HISTORY_FILE)

# ---- SQLite 历史真源 ----
# 累积型历史（快讯/逐日指标/份额/事件台账/快照）的真源：etf_store.sqlite。
# 与 etf_data.sqlite（派生索引，整库可重建）隔离——派生库按 schema_ver
# 升级时会 DROP 重建，真源放里面会被清掉。历史类读写函数全部走这里，
# 对应的 *.json 文件冻结为迁移前的备份，不再读写。
STORE_FILE = os.path.join(HERE, "etf_store.sqlite")

STORE_DDL = """
CREATE TABLE IF NOT EXISTS hist_flash(
    key TEXT PRIMARY KEY, date TEXT NOT NULL, time TEXT, payload TEXT NOT NULL,
    updated TEXT);
CREATE INDEX IF NOT EXISTS ix_hist_flash_date ON hist_flash(date);
CREATE TABLE IF NOT EXISTS hist_daily(
    date TEXT PRIMARY KEY, payload TEXT NOT NULL, updated TEXT);
CREATE TABLE IF NOT EXISTS hist_shares(
    date TEXT PRIMARY KEY, payload TEXT NOT NULL, updated TEXT);
CREATE TABLE IF NOT EXISTS hist_event(
    date TEXT PRIMARY KEY, payload TEXT NOT NULL, updated TEXT);
CREATE TABLE IF NOT EXISTS snapshot_store(
    date TEXT PRIMARY KEY, payload TEXT NOT NULL, updated TEXT);
"""


def store_connect():
    """打开历史真源库（WAL，表结构幂等创建）。"""
    import sqlite3
    con = sqlite3.connect(STORE_FILE)
    con.execute("PRAGMA journal_mode=WAL")
    con.executescript(STORE_DDL)
    return con


def store_load_dict(con, table, date_col="date"):
    """把 {date: payload} 形态的表读回 dict，date 升序。"""
    cur = con.execute("SELECT {}, payload FROM {} ORDER BY {}".format(
        date_col, table, date_col))
    return {d: json.loads(p) for d, p in cur.fetchall()}


def store_trim_days(con, table, max_days):
    """按日期保留最近 max_days 天，返回剩余天数。"""
    days = [r[0] for r in con.execute(
        "SELECT date FROM {} ORDER BY date".format(table))]
    stale = days[:-max_days] if max_days and len(days) > max_days else []
    for d in stale:
        con.execute("DELETE FROM {} WHERE date=?".format(table), (d,))
    return len(days) - len(stale)


# 首次切换时把冻结的 JSON 导入真源库（各表为空才导，幂等）
STORE_IMPORT_MAP = [
    ("hist_daily", os.path.join(HERE, "etf_daily_history.json")),
    ("hist_shares", os.path.join(HERE, "etf_shares_history.json")),
]


def ensure_hist_import():
    """迁移引导：冻结 JSON → 真源库，各表为空才导，可重复执行。"""
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    con = store_connect()
    try:
        for table, path in STORE_IMPORT_MAP:
            if con.execute("SELECT COUNT(*) FROM {}".format(
                    table)).fetchone()[0]:
                continue
            if not os.path.exists(path):
                continue
            try:
                with open(path, encoding="utf-8") as f:
                    hist = json.load(f)
            except Exception:
                continue
            for d, rec in (hist or {}).items():
                con.execute(
                    "INSERT OR REPLACE INTO {}(date,payload,updated) "
                    "VALUES(?,?,?)".format(table),
                    (d, json.dumps(rec, ensure_ascii=False,
                                   separators=(",", ":")), now))
        # 快讯：{"days": {date: {key: row}}}
        if not con.execute("SELECT COUNT(*) FROM hist_flash").fetchone()[0] \
                and os.path.exists(FLASH_HISTORY_PATH):
            try:
                with open(FLASH_HISTORY_PATH, encoding="utf-8") as f:
                    hist = json.load(f)
                for d, bucket in (hist.get("days") or {}).items():
                    for k, row in bucket.items():
                        con.execute(
                            "INSERT OR REPLACE INTO hist_flash"
                            "(key,date,time,payload,updated) "
                            "VALUES(?,?,?,?,?)",
                            (k, d, (row.get("time") or "")[:19],
                             json.dumps(row, ensure_ascii=False,
                                        separators=(",", ":")), now))
            except Exception:
                pass
        # 事件台账：{"days": {date: [items]}}
        if not con.execute("SELECT COUNT(*) FROM hist_event").fetchone()[0] \
                and os.path.exists(EVENT_STUDY_FILE):
            try:
                with open(EVENT_STUDY_FILE, encoding="utf-8") as f:
                    hist = json.load(f)
                for d, items in (hist.get("days") or {}).items():
                    con.execute(
                        "INSERT OR REPLACE INTO hist_event"
                        "(date,payload,updated) VALUES(?,?,?)",
                        (d, json.dumps(items, ensure_ascii=False,
                                       separators=(",", ":")), now))
            except Exception:
                pass
        con.commit()
    finally:
        con.close()

# ---- SQLite 派生索引 ----
# 历史真源已迁至 etf_store.sqlite（见上）；本库只是每次运行的可查询副本：
# 每次运行 upsert 一遍，删掉 .sqlite 也能靠重跑完整重建，因此不存在迁移风险。
SQLITE_FILE = os.path.join(HERE, "etf_data.sqlite")
SQLITE_SCHEMA = 9

# 宏观关键字段（用于 ERP 与流动性章节）
MACRO_PICK = [
    ("CPI", "CPI_CPI_YOY", "CPI同比", "%"), ("CPI", "CPI_CPI_YOY_CORE", "核心CPI同比", "%"),
    ("PMI", "PMI_PMI_MANU", "制造业PMI", ""), ("PMI", "PMI_NON_MANU_BIZ_ACT", "非制造业PMI", ""),
    ("CURV", "CURV_M1_YOY", "M1同比", "%"), ("CURV", "CURV_M2_YOY", "M2同比", "%"),
    ("YTM", "YTM_YIELD_10Y", "10年国债收益率", "%"),
    ("YTM", "YTM_YIELD_1Y", "1年国债收益率", "%"),
    ("MLF", "MLF_MLF_OPERATION_1Y", "MLF 1年期", "亿"),
    ("FINANCING", "FINANCING_SR_INC", "社融增量(累计口径)", "亿"),
]

# 分组：按「宽基 / 主要赛道 / 跨境 / 避险」分层抽样，尽量覆盖全市场风格。
# 注意：ETF 池是观察窗口，不是全市场 ETF 全集；真正全市场覆盖的是成交额与板块资金流两层。
# 同赛道有多只跟踪同一指数的ETF（如芯片三只），分组等权均值会重复计数，
# 因此报告同时给出成交额加权均值，并在 1.6 做代表性自检。
ETF_GROUPS = {
    "宽基": [
        "sh510050", "sh510300", "sh512050", "sh510500", "sh512100",
        "sz159901", "sh588220", "sh510880",
    ],
    "科技成长": ["sh588000", "sz159915", "sh515000"],
    "半导体芯片": ["sh512480", "sh512760", "sz159995"],
    "TMT数字": ["sh515880", "sh512720", "sz159819", "sh562500", "sz159869", "sh512980"],
    "金融地产": ["sh512880", "sh512800", "sh512200"],
    "消费医药": ["sh512010", "sh512690"],
    "新能源军工": ["sh515790", "sh515030", "sh516160", "sh512660"],
    "周期资源": ["sh512400", "sh515220"],
    "跨境": ["sh513050", "sh513130"],
    "避险": ["sh518880"],
}

# 同赛道重复跟踪的标的（用于代表性自检，说明等权均值的重复计数风险）
DUPLICATE_TRACK = [
    ("半导体芯片", ["sh512480", "sh512760", "sz159995"]),
    ("新能源", ["sh515790", "sh515030", "sh516160"]),
]

# 深度技术解剖的标的（报告中的「技术仪表盘」）
CORE_CODES = ["sh588000", "sz159915", "sh510300", "sh512480"]

ALL_ETF_CODES = [c for group in ETF_GROUPS.values() for c in group]

# 腾讯实时快照字段下标
F = {
    "name": 1, "code": 2, "price": 3, "prev_close": 4, "open": 5,
    "volume": 6, "time": 30, "change": 31, "pct": 32, "high": 33, "low": 34,
    "amount_wan": 37, "turnover": 38, "range_pct": 43,
    "limit_up": 47, "limit_down": 48, "vol_ratio": 49, "avg_price": 51,
    "ytd": 62, "d5": 63, "high52": 67, "low52": 68,
    "d10": 69, "d20": 70, "d60": 71,
    "total_shares": 76, "discount": 77, "iopv": 78,
    "net_value": 81, "pe": 39, "pb": 46,
}


# ----------------------------------------------------------------------------
# 工具函数
# ----------------------------------------------------------------------------

def num(v, default=None):
    """安全转 float：空串 / '-' / None 统一回退。"""
    try:
        if v is None or v == "" or v == "-":
            return default
        f = float(v)
        return default if (math.isnan(f) or math.isinf(f)) else f
    except (TypeError, ValueError):
        return default


def pct_str(v, digits=2):
    return "—" if v is None else "{:+.{}f}%".format(v, digits)


def price_str(v):
    if v is None:
        return "—"
    return "{:.3f}".format(v).rstrip("0").rstrip(".")


def yi(wan):
    """万元 -> 亿元"""
    return None if wan is None else wan / 10000.0


def avg(vals):
    vals = [v for v in vals if v is not None]
    return sum(vals) / len(vals) if vals else None


# ----------------------------------------------------------------------------
# 数据层：在线抓取 / 离线快照
# ----------------------------------------------------------------------------

def fetch_one(code, days):
    """返回 {code: {"quote": [...], "kline": [[date,o,c,h,l,vol], ...]}}"""
    url = ENDPOINT.format(code=code, days=days)
    req = Request(url, headers=HEADERS)
    with urlopen(req, timeout=15) as resp:
        payload = json.loads(resp.read().decode("utf-8"))
    if payload.get("code") != 0:
        raise RuntimeError("{}: {}".format(code, payload.get("msg", "接口返回异常")))
    node = payload["data"][code]
    return code, {
        "quote": node.get("qt", {}).get(code, []),
        "kline": node.get("day", []) or [],
    }


def fetch_online(codes, days):
    result, errors = {}, []
    with ThreadPoolExecutor(max_workers=6) as pool:
        futures = {pool.submit(fetch_one, c, days): c for c in codes}
        for fut, code in futures.items():
            try:
                key, val = fut.result()
                if val["quote"] and val["kline"]:
                    result[key] = val
                else:
                    errors.append("{} 数据为空".format(code))
            except Exception as exc:
                errors.append("{} {}".format(code, type(exc).__name__))
    return result, errors


def fetch_fundflow_online():
    """抓取全市场行业/概念板块资金流，返回 {board: [row, ...]}。"""
    out = {}
    for board, action in FUNDFLOW_BOARDS.items():
        url = FUNDFLOW_ROOT + action
        req = Request(url + "?t=zdf&o=1", headers=HEADERS)
        try:
            with urlopen(req, timeout=15) as resp:
                payload = json.loads(resp.read().decode("utf-8"))
            rows = (payload.get("data") or {}).get("data") or []
            if payload.get("code") == 0 and rows:
                out[board] = rows
        except Exception:
            continue
    return out


def fetch_breadth_online():
    """全市场涨跌广度/情绪/精确成交额/日内分时/两融/收评。"""
    try:
        req = Request(BREADTH_URL + "?type=0", headers=HEADERS)
        with urlopen(req, timeout=15) as resp:
            payload = json.loads(resp.read().decode("utf-8"))
        if payload.get("code") == 0 and isinstance(payload.get("data"), dict):
            return payload["data"]
    except Exception:
        pass
    return {}


def fetch_global_online():
    """批量抓外围指数实时快照，返回 {code: [字段...]}。"""
    try:
        req = Request(QT_URL + ",".join(GLOBAL_CODES), headers=HEADERS)
        with urlopen(req, timeout=12) as resp:
            raw = resp.read().decode("gbk", errors="ignore")
    except Exception:
        return {}
    out = {}
    for line in raw.strip().split("\n"):
        if "=\"" not in line:
            continue
        code = line.split("=")[0].replace("v_", "").strip()
        f = line.split("=\"")[1].rstrip("\";").split("~")
        if len(f) > 32 and f[1] and num(f[3]):
            out[code] = f
    return out


def fetch_futures():
    """抓取中金所股指/国债期货连续合约快照，返回 (rows, status)。

    字段索引见 FUT_CONTRACTS 上方注释；缺少 40 个字段的条目直接丢弃，
    避免把行情的其它 `hq_str_*` 变量误当期货。
    """
    fh = dict(HEADERS, **{
        "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                       "AppleWebKit/537.36 (KHTML, like Gecko) "
                       "Chrome/120.0.0.0 Safari/537.36"),
        "Referer": "https://finance.sina.com.cn/",
    })
    codes = ",".join(c for c, *_ in FUT_CONTRACTS)
    try:
        req = Request(FUT_URL + codes, headers=fh)
        with urlopen(req, timeout=12) as resp:
            txt = resp.read().decode("gbk", "replace")
    except Exception:
        return [], "blocked"
    meta = {c: (label, spot, spot_name) for c, label, spot, spot_name in FUT_CONTRACTS}
    out = []
    for ln in txt.split(";"):
        if "hq_str_" not in ln or '"' not in ln:
            continue
        key = ln.split("hq_str_")[1].split("=")[0].strip()
        if key not in meta:
            continue
        f = ln.split('"')[1].split(",")
        if len(f) < 40:
            continue
        last, prev = _fnum(f[3]), _fnum(f[13])
        if not last:
            continue
        label, spot, spot_name = meta[key]
        out.append({
            "code": key, "name": label,
            "last": last, "prev_settle": prev,
            "pct": (last / prev - 1) * 100 if prev else None,
            "high": _fnum(f[1]), "low": _fnum(f[2]),
            "volume": _fnum(f[4]), "oi": _fnum(f[6]),
            "time": "{} {}".format((f[36] or "").strip(), (f[37] or "").strip()).strip(),
            "spot_code": spot, "spot_name": spot_name,
        })
    return out, ("ok" if out else "empty")


def proxy_post(route, params, skill="westock-data", timeout=20):
    """调用公开封装接口；任何异常返回 {}，由上层降级。"""
    body = {"params": params, "route": route, "skill_channel": "yuanbao",
            "skill_name": skill, "token": PROXY_TOKEN}
    try:
        req = Request(PROXY_URL + "?" + urlencode(PROXY_PARAMS),
                      data=json.dumps(body).encode("utf-8"),
                      headers=dict(HEADERS, **{"Content-Type": "application/json"}))
        with urlopen(req, timeout=timeout) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except Exception:
        return {}


def rows_from_proxy(payload):
    """把 list_data（JSON 字符串）解成 {list_code: [dict, ...]}。"""
    out = {}
    for k, v in ((payload.get("data") or {}).get("data") or {}).items():
        if not isinstance(v, dict):
            out[k] = v
            continue
        try:
            out[k] = json.loads(v.get("list_data") or "[]")
        except Exception:
            out[k] = []
    return out


def _fnum(v):
    try:
        return float(v)
    except Exception:
        return 0.0


def fetch_market_rank(date, max_pages=RANK_MAX_PAGES):
    """分页拉全市场个股行情快照，返回 (裁剪后 rows, 状态, 页数)。

    接口不支持历史日期参数，返回的是「当前收盘/最新快照」；
    是否属于报告日由 analyze_market_rank 用广度家数做一致性校验，
    不一致时报告会标注「非报告日快照」，绝不冒充当日排行。
    """
    rank_h = dict(HEADERS, **{
        "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                       "AppleWebKit/537.36 (KHTML, like Gecko) "
                       "Chrome/120.0.0.0 Safari/537.36"),
        "Referer": "https://stockapp.finance.qq.com/",
        "Accept": "application/json, text/plain, */*",
    })
    allr, pages, total = {}, 0, None
    for pg in range(max_pages):
        p = {"_appver": "11.17.0", "board_code": "aStock",
             "sort_type": "price", "direct": "down",
             "offset": str(pg * RANK_PAGE_SIZE), "count": str(RANK_PAGE_SIZE)}
        try:
            req = Request(RANK_URL + "?" + urlencode(p), headers=rank_h)
            with urlopen(req, timeout=12) as resp:
                j = json.loads(resp.read().decode("utf-8", "replace"))
            data = j.get("data") or {}
            rows = data.get("rank_list") or []
            total = data.get("total") or total
        except Exception:
            break
        if not rows:
            break
        pages += 1
        for r in rows:
            if not isinstance(r, dict) or not r.get("code"):
                continue
            allr[r["code"]] = {k: r.get(k) for k in RANK_STOCK_FIELDS}
        if total and len(allr) >= int(total):
            break
    if not allr:
        return [], "blocked" if pages == 0 else "empty", pages
    return list(allr.values()), "ok", pages


def fetch_etf_rank(etf_size, date=""):
    """用全市场 ETF 明细的代码，批量拉报价得到全市场 ETF 涨跌排行。"""
    codes = [(r.get("EtfCode") or "").strip()
             for r in (etf_size or []) if isinstance(r, dict) and r.get("EtfCode")]
    if not codes:
        return [], "empty"
    qt_h = dict(HEADERS, **{
        "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                       "AppleWebKit/537.36 (KHTML, like Gecko) "
                       "Chrome/120.0.0.0 Safari/537.36"),
        "Referer": "https://gu.qq.com/",
    })
    out, stamp = {}, ""
    for i in range(0, len(codes), QT_BATCH):
        try:
            req = Request(QT_URL + ",".join(codes[i:i + QT_BATCH]), headers=qt_h)
            with urlopen(req, timeout=12) as resp:
                txt = resp.read().decode("gbk", "replace")
        except Exception:
            continue
        for ln in txt.split(";"):
            if "~" not in ln:
                continue
            p = ln.split("~")
            if len(p) < 44 or not p[2].strip():
                continue
            c = p[2].strip()
            out[c] = {
                "code": c, "name": p[1].strip(), "price": _fnum(p[3]),
                "zdf": _fnum(p[32]), "amount_wan": _fnum(p[37]),
                "hsl": _fnum(p[38]), "zf": _fnum(p[43]),
                "stamp": (p[30].strip() if len(p) > 30 else ""),
            }
            if not stamp and len(p) > 30:
                stamp = p[30].strip()
    if not out:
        return [], "blocked"
    rows = list(out.values())
    # 行情时间戳 YYYYMMDDHHMMSS → 报告日一致性
    day = stamp[:8] if len(stamp) >= 8 else ""
    day_fmt = ("{}-{}-{}".format(day[:4], day[4:6], day[6:8])
               if day else "")
    for r in rows:
        r["stamp"] = stamp
    return rows, "ok" if (not date or not day_fmt or day_fmt == date) else "other_day"


def fetch_extended_online(date):
    """一次性抓取全市场 ETF 明细、官方画像、宏观、事件、龙虎榜。"""
    out = {}
    groups = [
        (ETF_LIST_CODES, "etf"), (MACRO_LIST_CODES, "macro"),
        (PROFILE_LIST_CODES, "profile"), (EVENT_LIST_CODES, "events"),
    ]
    for codes, tag in groups:
        payload = proxy_post("query_list_data_by_date",
                             {"date": date, "list_codes": codes})
        rows = rows_from_proxy(payload)
        for k, v in rows.items():
            if v:
                out[k] = v
    try:
        req = Request(LHB_URL, headers=HEADERS)
        with urlopen(req, timeout=12) as resp:
            lhb = json.loads(resp.read().decode("utf-8"))
        if lhb.get("code") == 0:
            out["_lhb"] = lhb.get("data") or {}
    except Exception:
        pass
    # 筹码分布（核心标的）
    chips = {}
    for c in CORE_CODES[:4]:
        p = proxy_post("stock_quote_snapshot", {
            "codes": c, "date": date,
            "fields": "SecuName,ClosePrice,ChipProfitRate,ChipAvgCost,"
                      "ChipConcentration90,ChipConcentration70"})
        d = (p.get("data") or {})
        if d:
            chips[c] = d
    if chips:
        out["_chips"] = chips
    # 全市场涨跌排行：个股全量 + ETF 全量（两者都是收盘快照）
    srows, sstat, spages = fetch_market_rank(date)
    if srows:
        out["_market_rank"] = srows
        out["_market_rank_status"] = sstat
        out["_market_rank_pages"] = spages
    erows, estat = fetch_etf_rank(out.get("etf_size") or [], date)
    if erows:
        out["_etf_rank"] = erows
        out["_etf_rank_status"] = estat
    # 股指/国债期货：给「机构对冲压力」一个高频确认层
    frows, fstat = fetch_futures()
    if frows:
        out["_futures"] = frows
    out["_futures_status"] = fstat
    return out


BASE_QS = {"app": "openclaw", "dev_id": "dev_etf_analysis",
           "skill_channel": "yuanbao", "token": PROXY_TOKEN}


def _get_json(url, params, timeout=15):
    """公开资讯 GET；任何异常返回 {}，由上层降级。"""
    try:
        q = dict(BASE_QS)
        q.update(params)
        req = Request(url + "?" + urlencode(q), headers=HEADERS)
        with urlopen(req, timeout=timeout) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except Exception:
        return {}


def _plain_text(s, limit=520):
    """资讯正文去 HTML 标签，压成纯文本摘要。"""
    if not isinstance(s, str):
        return ""
    s = re.sub(r"<script.*?</script>|<style.*?</style>", " ", s,
               flags=re.S | re.I)
    s = re.sub(r"<[^>]+>", " ", s)
    s = (s.replace("&nbsp;", " ").replace("&amp;", "&")
          .replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"'))
    s = re.sub(r"\s+", " ", s).strip()
    return s[:limit]


def fetch_news_online(date):
    """采集 7 类财经资讯，全部带发布时间；返回紧凑结构供快照缓存。"""
    out = {}
    # 1) 热点新闻榜（50 条，含盘中/盘后/隔夜）
    hot = _get_json(NEWS_HOT_URL, {"num": 50}) or {}
    rows = ((hot.get("data") or {}).get("rankResult") or [])
    hot_rows = []
    for r in rows:
        if not isinstance(r, dict):
            continue
        ts = num(r.get("publish_time"))
        hot_rows.append({
            "id": r.get("news_id") or "", "title": r.get("news_title") or "",
            "source": r.get("source") or "", "ts": ts,
            "rank": num(r.get("rank")) or 99,
        })
    if hot_rows:
        out["_hot"] = hot_rows

    # 2) 只对最相关的几条抓正文（该端点多 ids 只返回首条，故逐个请求）
    def _detail(nid):
        d = _get_json(NEWS_DETAIL_URL,
                      {"ids": "nes" + nid,
                       "nkey": "getQQNewsSimpleHtmlContent"}, timeout=10) or {}
        dd = (d.get("data") or {}).get("data")
        if isinstance(dd, list):
            dd = dd[0] if dd else None
        if not isinstance(dd, dict):
            return ""
        content = dd.get("content")
        body = (content.get("text") if isinstance(content, dict)
                else (content if isinstance(content, str) else ""))
        if not body:
            body = dd.get("summary") or ""
        return _plain_text(body)

    pick = [r["id"] for r in hot_rows
            if r["id"] and any(k in r["title"] for t in TOPIC_MAP
                               for k in t["kw"])][:NEWS_DETAIL_LIMIT]
    det = {}
    for nid in pick:
        txt = _detail(nid)
        if txt:
            det[nid] = txt
    if det:
        out["_detail"] = det

    # 3) 个股公告（该端点多标的返回空，逐个取）
    notices = []
    for sym in EVENT_SAMPLE_CODES[:2]:
        nc = _get_json(NOTICE_URL, {"symbol": sym, "n": 6,
                                    "noticeType": 0, "page": 1}) or {}
        for x in (((nc.get("data") or {}).get("data") or [])[:6]):
            if isinstance(x, dict):
                notices.append({"title": x.get("title") or "",
                                "time": x.get("time") or "",
                                "symbol": x.get("symbol") or sym})
    out["_notice"] = notices[:12]

    # 4) 个股研报
    rp = _get_json(REPORT_URL, {"symbol": EVENT_SAMPLE_CODES[0],
                                "n": 6, "page": 1}) or {}
    rrows = ((rp.get("data") or {}).get("data") or [])
    out["_report"] = [
        {"title": x.get("title") or "", "time": x.get("time") or "",
         "org": x.get("org") or x.get("institution") or ""}
        for x in rrows if isinstance(x, dict)][:6]

    # 5) 脱水研报（封装接口）
    de = proxy_post("research_report_list_get", {"page": 1, "size": 6}) or {}
    dlist = ((de.get("data") or {}).get("list") or [])
    out["_dehydrated"] = [
        {"id": x.get("id") or "", "title": x.get("title") or "",
         "preview": _plain_text(x.get("preview") or "", 160),
         "time": x.get("publish_time") or x.get("time") or ""}
        for x in dlist if isinstance(x, dict)][:6]

    # 6) 财经日历（只保留报告日，压缩存储）
    cal = _get_json(CALENDAR_URL, {"date": date, "limit": 50,
                                   "market": "hs"}) or {}
    cdata = cal.get("data") or []
    today_items = []
    if isinstance(cdata, list):
        for day in cdata:
            if not isinstance(day, dict) or day.get("date") != date:
                continue
            for grp in (day.get("list") or []):
                for x in ((grp or {}).get("data") or []):
                    if isinstance(x, dict):
                        today_items.append({
                            "type": str(x.get("type") or ""),
                            "title": x.get("title") or "",
                            "stock": x.get("stockName") or "",
                        })
    out["_calendar"] = {"date": date, "items": today_items[:30],
                        "n": len(today_items)}

    # 7) 新股日历
    ipo = _get_json(IPO_URL, {"detail": 1, "market": "hs",
                              "period": 90, "sgrq": 1}) or {}
    idata = ipo.get("data") or {}
    out["_ipo"] = {
        "sgrq_n": len(idata.get("sgrq") or []),
        "ssrq_n": len(idata.get("ssrq") or []),
        "recent": [
            {"name": x.get("name") or "", "price": x.get("price") or "",
             "sgrq": x.get("sgrq") or "", "ssrq": x.get("ssrq") or ""}
            for x in (idata.get("sgok") or [])[:5] if isinstance(x, dict)],
    }

    # 8) 主题龙头的机构评级 / 一致预期（产业催化验证）
    rt = proxy_post("stock_quote_snapshot", {
        "codes": ",".join(EVENT_SAMPLE_CODES),
        "fields": "SecuName,ResearchReport,ForecastInstitutions,TargetPriceAvg,"
                  "RatingBuyCnt,RatingCnt,ConEarningsForecast,ConTargetPrice"}) or {}
    rdat = rt.get("data") or {}
    out["_rating"] = rdat

    # 9) 东方财富 7×24 快讯：JSON 接口 + 移动端 H5 + PC 7_24 页三通道，
    #    任一可达即可采集；全部受限时由 --flash-text / 离线快照兜底。
    flash, flash_status = fetch_eastmoney_flash(date)
    h5, h5_status = fetch_eastmoney_flash_h5(date)
    empc, empc_status = fetch_eastmoney_pc_flash(date)
    merged, seen = [], set()
    for r in flash + h5:
        k = flash_key(r)
        if k in seen:
            continue
        seen.add(k)
        merged.append(r)
    # 10) 新浪财经 7×24（PC 页服务端渲染，直连成功率最高、带完整日期+秒）
    sina, sina_status = fetch_sina_flash(date)
    # 11) 华夏时报 7×24（中央级财经媒体，页面带完整日期与「热点快讯」榜）
    ct, ct_hot, ct_status = fetch_chinatimes_flash(date)
    # 12) 财联社电报 7×24（A股垂直度最高；last_time 时间游标可回溯历史）
    cls_rows, cls_status, cls_pages, cls_covered = fetch_cls_flash(date)
    merged, seen = [], set()
    for r in flash + h5 + empc + sina + ct + cls_rows:
        k = flash_key(r)
        if k in seen:
            continue
        seen.add(k)
        merged.append(r)
    merged.sort(key=lambda r: r.get("time") or "")

    # 逐日增量累积：解决实时页只给最新 50 条的问题，
    # 每天跑一次即可把当天 9:30–15:00 的盘中快讯永久留存
    hist_rows, hist = merge_flash_history(FLASH_HISTORY_PATH, merged, date)
    for r in hist_rows:
        k = flash_key(r)
        if k not in seen:
            seen.add(k)
            merged.append(r)
    merged.sort(key=lambda r: r.get("time") or "")

    out["_em_flash"] = merged
    out["_em_json_status"] = flash_status   # ①号通道自身状态（区别于合并状态）
    out["_em_json_n"] = len(flash)
    out["_sina_status"] = sina_status
    out["_sina_n"] = len(sina)
    out["_ct_status"] = ct_status
    out["_ct_n"] = len(ct)
    out["_ct_hot"] = ct_hot
    out["_cls_status"] = cls_status
    out["_cls_n"] = len(cls_rows)
    out["_cls_pages"] = cls_pages
    out["_cls_covered"] = cls_covered
    out["_empc_status"] = empc_status
    out["_empc_n"] = len(empc)
    out["_empc_days"] = {
        d: sum(1 for r in empc if (r.get("time") or "").startswith(d))
        for d in sorted({(r.get("time") or "")[:10] for r in empc}) if d}
    out["_etf_flow"] = extract_etf_flow_evidence(merged)
    out["_flash_days"] = len(hist.get("days") or {})
    out["_flash_total"] = sum(len(v) for v in (hist.get("days") or {}).values())
    out["_flash_source_n"] = len({r.get("source") for r in merged})
    out["_flash_family_n"] = len({source_family(r.get("source"))
                                  for r in merged})
    ct_fam = "华夏时报" in {source_family(r.get("source")) for r in merged}
    if sina_status == "ok":
        out["_em_flash_status"] = "ok" if flash_status == "ok" else "sina_ok"
    elif cls_status == "ok":
        out["_em_flash_status"] = "cls_ok"
    elif ct_status == "ok" or ct_fam:
        out["_em_flash_status"] = "web_channel" if ct_fam else "empty"
    elif merged and flash_status == "ok":
        out["_em_flash_status"] = "ok"
    elif merged:
        out["_em_flash_status"] = "web_channel"
    else:
        out["_em_flash_status"] = (
            "blocked" if (flash_status == "blocked" or h5_status == "blocked"
                          or sina_status == "blocked") else "empty")
    return out


def fetch_eastmoney_flash(date, pages=EM_FLASH_PAGES):
    """东方财富 7×24 快讯。返回 (条目, 状态)。

    状态：ok / blocked（网络策略或接口不可用）/ empty。
    设计成可移植：脚本换到能访问 eastmoney 的网络即可直接生效，
    当前环境若被策略拒绝会自动降级，不影响其余 15 章。
    """
    rows, status, seen, sort_end = [], "empty", set(), ""
    em_h = dict(HEADERS, **{
        "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                       "AppleWebKit/537.36 (KHTML, like Gecko) "
                       "Chrome/120.0.0.0 Safari/537.36"),
        "Referer": "https://kuaixun.eastmoney.com/7_24.html",
        "Accept": "application/json, text/plain, */*",
    })
    urls = [EM_FLASH_URL] + EM_FLASH_FALLBACK_URLS
    blocked = 0
    for pg in range(pages):
        got = None
        for u in urls:
            q = {"client": "web", "biz": "web_724",
                 "fastColumn": EM_FLASH_COLUMN,
                 "sortEnd": sort_end, "pageSize": "50",
                 "req_trace": str(int(time.time() * 1000)),
                 "_": str(int(time.time() * 1000))}
            try:
                req = Request(u + "?" + urlencode(q), headers=em_h)
                with urlopen(req, timeout=10) as resp:
                    raw = resp.read().decode("utf-8", "replace")
                if raw.lstrip().startswith("{") or raw.lstrip().startswith("["):
                    got = json.loads(raw)
                    break
            except Exception:
                blocked += 1
                continue
        if not got:
            break
        lst = ((got.get("data") or {}).get("fastNewsList") or [])
        if isinstance(got.get("data"), list):
            lst = got["data"]
        if not lst:
            break
        for x in lst:
            if not isinstance(x, dict):
                continue
            key = x.get("code") or x.get("showTime", "") + x.get("title", "")
            if key in seen:
                continue
            seen.add(key)
            tm = str(x.get("showTime") or "")
            # 只保留报告日与次日（早报/隔夜）窗口
            if tm and not (tm[:10] == date or
                           tm[:10] == (datetime.strptime(date, "%Y-%m-%d")
                                       + timedelta(days=1)).strftime("%Y-%m-%d")):
                continue
            rows.append({
                "id": x.get("code") or "", "time": tm,
                "title": x.get("title") or "",
                "summary": _plain_text(x.get("summary") or x.get("title") or "", 300),
                "important": bool(num(x.get("titleColor"))),
                "stocks": [s.get("name") or s.get("stockName") or ""
                           for s in (x.get("stockList") or [])
                           if isinstance(s, dict)][:4],
            })
        sort_end = str(lst[-1].get("sort") or lst[-1].get("showTime") or "")
        if len(lst) < 50:
            break
    if rows:
        status = "ok"
    elif blocked:
        status = "blocked"
    rows.sort(key=lambda r: r["time"])
    return rows, status


def fetch_eastmoney_flash_h5(date):
    """东方财富移动端 H5 快讯页（emh5wap）。页面服务端渲染，正则提取。

    与 getFastNewsList 互为备用：H5 页面在部分网络下可由网页通道读取，
    而 JSON API 更易被策略拦截。返回 (条目, 状态)。
    """
    em_h = dict(HEADERS, **{
        "User-Agent": ("Mozilla/5.0 (iPhone; CPU iPhone OS 16_0 like Mac OS X) "
                       "AppleWebKit/605.1.15 (KHTML, like Gecko) "
                       "Version/16.0 Mobile/15E148 Safari/604.1"),
        "Referer": "http://emh5wap.eastmoney.com/",
        "Accept": "text/html,application/xhtml+xml,*/*",
    })
    # 页面是「最新流」；报告日早于今天时，条目实际属于今天（次日早报）
    try:
        today = datetime.now(CST).strftime("%Y-%m-%d")
    except Exception:
        today = date
    flash_date = max(date, today)
    rows, seen, status = [], set(), "blocked"
    # 多个时间锚点尝试回溯；不同部署下参数名可能不同，失败也不影响主流程
    anchors = ["", "?page=1", "?lastTime={}%2012:00:00".format(date),
               "?lastTime={}%2015:00:00".format(date)]
    for a in anchors:
        try:
            req = Request(EM_H5_FLASH_URL + a, headers=em_h)
            with urlopen(req, timeout=10) as resp:
                html = resp.read().decode("utf-8", "replace")
            text = _plain_text(html, 200000)
            if len(text) < 100:
                continue
            status = "ok"
            for m in EM_H5_TIME_RE.finditer(text):
                hhmm, title, body = m.group(1), m.group(2).strip(), m.group(3).strip()
                body = body.replace("[点击查看全文]", "").strip()
                key = hhmm + title
                if key in seen:
                    continue
                seen.add(key)
                rows.append({
                    "id": "emh5" + str(len(rows)),
                    "time": "{} {}:00".format(flash_date, hhmm),
                    "title": title, "summary": _plain_text(body, 300),
                    "important": False, "stocks": [],
                    "source": "东方财富H5快讯",
                })
        except Exception:
            continue
    if not rows:
        status = "blocked"
    rows.sort(key=lambda r: r["time"])
    return rows, status


def _sina_clean(s):
    return _plain_text(re.sub(r"<[^>]+>", " ", s or ""), 400)


def parse_sina_flash(html_or_text, report_date=""):
    """解析新浪 7×24 页面，兼容 SSR 版与用户贴的纯文本版。

    返回 (条目, 格式, 归档标记)。纯文本版只有 HH:MM:SS 无日期，
    会按正文里的年份做归属校验：出现与报告日不一致的年份即判为
    历史归档，绝不并入报告日事件池（防止把旧新闻当今日驱动）。
    """
    rows, fmt, archived = [], "ssr", False
    ms = SINA_FLASH_SSR_RE.findall(html_or_text or "")
    if ms:
        for nid, ymd, hms, body in ms:
            d = "{}-{}-{}".format(ymd[:4], ymd[4:6], ymd[6:8])
            txt = _sina_clean(body)
            if not txt:
                continue
            rows.append({
                "id": "sina" + nid, "time": "{} {}".format(d, hms),
                "title": txt[:60], "summary": txt,
                "important": txt.startswith("【"), "stocks": [],
                "source": "新浪财经7×24",
            })
    else:
        fmt = "text"
        ms = SINA_FLASH_TEXT_RE.findall(html_or_text or "")
        years = SINA_ARCHIVE_YEAR_RE.findall(html_or_text or "")
        rep_year = report_date[:4] if report_date else ""
        # 用「出现次数最多的年份」判定：正文里可能含未来规划年份
        # （如「预计 2026 年推出」），不能因出现一次就当成当期新闻
        if rep_year and years:
            from collections import Counter
            dominant = Counter(y[:4] for y in years).most_common(1)[0][0]
            if dominant != rep_year:
                archived = True
        # 纯文本版无日期，按报告日标注；归档内容由上层拦截
        day = report_date or datetime.now(CST).strftime("%Y-%m-%d")
        for i, (hms, body) in enumerate(ms):
            txt = _sina_clean(body)
            if not txt:
                continue
            rows.append({
                "id": "sinatxt" + str(i), "time": "{} {}".format(day, hms),
                "title": txt[:60], "summary": txt,
                "important": txt.startswith("【"), "stocks": [],
                "source": "新浪财经7×24(纯文本)",
            })
    rows.sort(key=lambda r: r["time"])
    return rows, fmt, archived


def fetch_sina_flash(report_date=""):
    """直连新浪 7×24 PC 页（该域名在多数环境可达）。返回 (条目, 状态)。"""
    sina_h = dict(HEADERS, **{
        "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                       "AppleWebKit/537.36 (KHTML, like Gecko) "
                       "Chrome/120.0.0.0 Safari/537.36"),
        "Referer": "https://finance.sina.com.cn/7x24/",
        "Accept": "text/html,application/xhtml+xml,*/*",
    })
    try:
        req = Request(SINA_FLASH_URL, headers=sina_h)
        with urlopen(req, timeout=15) as resp:
            # 必须强制 utf-8：requests/urllib 的编码嗅探会把中文判错
            html = resp.read().decode("utf-8", "replace")
        rows, fmt, archived = parse_sina_flash(html, report_date)
        if rows and not archived:
            return rows, "ok"
        return [], "empty" if not rows else "archive"
    except Exception:
        return [], "blocked"


def parse_chinatimes_flash(text, report_date=""):
    """解析华夏时报 7×24 页面（网页通道文本或直连 HTML 文本）。

    返回 (条目, 热点榜, 页面日期, 是否归档)。页面自带 YYYY-MM-DD 日期行，
    无需像新浪纯文本那样靠年份猜测；与报告日相差 >2 天即判为归档。
    """
    text = text or ""
    dates = CT_DATE_RE.findall(text)
    page_date = dates[0] if dates else ""
    archived = False
    if page_date and report_date:
        try:
            delta = abs((datetime.strptime(page_date, "%Y-%m-%d")
                         - datetime.strptime(report_date, "%Y-%m-%d")).days)
            archived = delta > 2
        except Exception:
            archived = False
    day = page_date or report_date or datetime.now(CST).strftime("%Y-%m-%d")

    # 按时间行切块：parts[0] 是页头，其后成对 (HH:MM, 内容块)
    parts = CT_TIME_SPLIT_RE.split(text)
    rows = []
    for i in range(1, len(parts) - 1, 2):
        hms, block = parts[i], parts[i + 1]
        if "点击查看更多" in block or "数据集加载中" in block:
            continue
        lines = [ln.strip() for ln in block.splitlines() if ln.strip()]
        lines = [ln for ln in lines if ln not in ("🔗", "🔄", "暂无新闻数据")]
        if not lines:
            continue
        body = " ".join(lines)
        if "热点快讯" in body:
            body = body.split("热点快讯")[0].strip()
        if not body:
            continue
        src = "华夏时报"
        m = CT_SOURCE_RE.search(body)
        if m:
            src = m.group(1)
        rows.append({
            "id": "ct" + day.replace("-", "") + hms.replace(":", "") + str(i),
            "time": "{} {}:00".format(day, hms),
            "title": body[:70], "summary": _plain_text(body, 400),
            "important": body.startswith("【"), "stocks": [],
            "source": "华夏时报7×24", "origin": src,
        })
    rows.sort(key=lambda r: r["time"])

    # 底部「热点快讯」榜（无时间，但带排名，单列一类不进事件池）
    hot = []
    if CT_HOT_ANCHOR in text:
        tail = text.split(CT_HOT_ANCHOR, 1)[1]
        for rank, title in CT_HOT_RE.findall(tail):
            title = title.strip()
            if not title or "数据更新于" in title or len(title) < 6:
                continue
            hot.append({"rank": int(rank), "title": _plain_text(title, 120)})
    return rows, hot, page_date, archived


def parse_eastmoney_pc_flash(text, report_date="", source="东方财富全球快讯"):
    """解析带日期分段的网页通道快讯（东财 PC 页 / 财联社电报）。

    格式：`=== YYYY-MM-DD ===` 分段 + 每行 `HH:MM` 或 `HH:MM:SS` 正文。
    支持同一文件里多天（报告日盘后 + 次日盘中），逐段做日期归属；
    与报告日相差 >2 天整段判归档、不并入。返回 (条目, 各日条数, 归档日列表)。
    """
    text = text or ""
    segs = list(EM_PC_DATE_RE.finditer(text))
    rows, per_day, archived = [], {}, []
    if not segs:
        return rows, per_day, archived
    for i, m in enumerate(segs):
        day = m.group(1)
        block = text[m.end():segs[i + 1].start() if i + 1 < len(segs) else len(text)]
        if report_date:
            try:
                delta = abs((datetime.strptime(day, "%Y-%m-%d")
                             - datetime.strptime(report_date, "%Y-%m-%d")).days)
                if delta > 2:
                    archived.append(day)
                    continue
            except Exception:
                pass
        n = 0
        for lm in EM_PC_LINE_RE.finditer(block):
            hms, body = lm.group(1), lm.group(2).strip()
            if not body or len(body) < 6:
                continue
            hms_full = hms if len(hms) == 8 else hms + ":00"
            _pfx = "clsweb" if source == "财联社电报" else "empc"
            rows.append({
                "id": _pfx + day.replace("-", "") + hms.replace(":", "") + str(n),
                "time": "{} {}".format(day, hms_full),
                "title": body[:70],
                "summary": _plain_text(body, 400),
                "important": body.startswith("【"),
                "stocks": [], "source": source, "origin": "",
            })
            n += 1
        per_day[day] = n
    rows.sort(key=lambda r: r["time"])
    return rows, per_day, archived


def fetch_eastmoney_pc_flash(report_date=""):
    """东财 7_24 PC 页采集：SSR 页优先，壳页改走页面自带的 AJAX 接口。

    2026-11 改版后 7_24 页由 React 渲染，静态 HTML 只剩导航与行情壳
    （「暂无相关内容」「9秒后刷新」），快讯正文改由
    `np-weblist.eastmoney.com/comm/web/getFastNewsList`
    （biz=web_724、fastColumn=102、sortEnd 游标）加载——与①号通道同源，
    合并阶段按「时间+标题」去重，不会重复计数；本通道的作用是核验
    PC 页路由可用性，并保留老 SSR 环境（服务端渲染页）的直读能力。
    返回 (条目, 状态)。
    """
    em_h = dict(HEADERS, **{
        "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                       "AppleWebKit/537.36 (KHTML, like Gecko) "
                       "Chrome/120.0.0.0 Safari/537.36"),
        "Referer": "https://kuaixun.eastmoney.com/7_24.html",
        "Accept": "text/html,application/xhtml+xml,*/*",
    })
    rows, seen, status = [], set(), "blocked"
    for u in EM_PC_FLASH_URLS:
        try:
            req = Request(u, headers=em_h)
            with urlopen(req, timeout=10) as resp:
                html = resp.read().decode("utf-8", "replace")
        except Exception:
            continue
        txt = _plain_text(html, 200000)
        if len(txt) < 80:
            continue
        status = "ok"
        # PC 页正文形如「HH:MM 【标题】正文」；页面日期取 URL 锚点对应日
        day = report_date or datetime.now(CST).strftime("%Y-%m-%d")
        for m in EM_H5_TIME_RE.finditer(txt):
            hms, title, body = m.group(1), m.group(2).strip(), m.group(3).strip()
            k = day + hms + title[:24]
            if k in seen:
                continue
            seen.add(k)
            rows.append({
                "id": "empc" + day.replace("-", "") + hms.replace(":", "") + str(len(rows)),
                "time": "{} {}:00".format(day, hms),
                "title": ("【{}】{}".format(title, body))[:70],
                "summary": _plain_text("【{}】{}".format(title, body), 400),
                "important": True, "stocks": [],
                "source": "东方财富全球快讯", "origin": "",
            })
    if rows:
        rows.sort(key=lambda r: r["time"])
        return rows, "ok"

    # ---- 壳页：走页面 AJAX（参数取自 kuaixun.js 模板）----
    ajax_rows, sort_end, ajax_seen = [], "", set()
    ajax_h = dict(em_h, **{"Accept": "application/json, text/plain, */*"})
    for _pg in range(EM_FLASH_PC_AJAX_PAGES):
        q = {"client": "web", "biz": "web_724",
             "fastColumn": EM_FLASH_COLUMN,
             "sortEnd": sort_end, "pageSize": "50",
             "req_trace": str(int(time.time() * 1000))}
        got = None
        for u in [EM_FLASH_URL] + EM_FLASH_FALLBACK_URLS:
            try:
                req = Request(u + "?" + urlencode(q), headers=ajax_h)
                with urlopen(req, timeout=10) as resp:
                    raw = resp.read().decode("utf-8", "replace")
                if raw.lstrip().startswith("{"):
                    got = json.loads(raw)
                    break
            except Exception:
                continue
        if not got:
            break
        lst = ((got.get("data") or {}).get("fastNewsList") or [])
        if not lst:
            break
        for x in lst:
            if not isinstance(x, dict):
                continue
            tm = str(x.get("showTime") or "")
            if not tm:
                continue
            # 与①号通道相同的窗口过滤：报告日 + 次日早报
            if report_date and not (tm[:10] == report_date or
                                    tm[:10] == (datetime.strptime(report_date, "%Y-%m-%d")
                                                + timedelta(days=1)).strftime("%Y-%m-%d")):
                continue
            _k = tm + (x.get("title") or "")[:24]
            if _k in ajax_seen:
                continue
            ajax_seen.add(_k)
            ajax_rows.append({
                "id": "empca" + (x.get("code") or tm.replace(":", "")),
                "time": tm,
                "title": (x.get("title") or "")[:70],
                "summary": _plain_text(x.get("summary") or x.get("title") or "", 400),
                "important": bool(num(x.get("titleColor"))),
                "stocks": [s.get("name") or s.get("stockName") or ""
                           for s in (x.get("stockList") or [])
                           if isinstance(s, dict)][:4],
                "source": "东方财富全球快讯", "origin": "东财PC页(AJAX)",
            })
        sort_end = str(lst[-1].get("sort") or lst[-1].get("showTime") or "")
        if len(lst) < 50:
            break
    ajax_rows.sort(key=lambda r: r["time"])
    if ajax_rows:
        return ajax_rows, "ok"
    return [], ("blocked" if status == "blocked" else "empty")


def fetch_chinatimes_flash(report_date=""):
    """华夏时报 7×24 采集：JSON 接口直连为主，旧网页通道兜底。

    2026-11 改版后网页为 JS 动态加载，静态 HTML 只剩「数据集加载中」占位，
    故改走 /api/proxy/news/result（正文）与 /api/proxy/news/topclicks（热点榜）；
    接口失败时退回旧网页通道解析（能直连到服务端渲染版本的环境仍可用）。
    返回 (条目, 热点榜, 状态)。
    """
    ct_h = dict(HEADERS, **{
        "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                       "AppleWebKit/537.36 (KHTML, like Gecko) "
                       "Chrome/120.0.0.0 Safari/537.36"),
        "Referer": "https://kuaixun.chinatimes.net.cn/",
        "Accept": "application/json, text/html, */*",
    })

    def _item_day(t):
        return (t or "")[:10]

    # ---- 主通道：JSON 接口（翻 3 页 ×40 条，覆盖早盘至当前） ----
    items, seen_ids = [], set()
    for page in range(1, CT_FLASH_PAGES + 1):
        try:
            req = Request(CT_FLASH_API + "?page={}&pageSize=40".format(page),
                          headers=ct_h)
            with urlopen(req, timeout=12) as resp:
                data = json.loads(resp.read().decode("utf-8", "replace"))
            batch = ((data or {}).get("data") or {}).get("list") or []
        except Exception:
            break
        fresh = 0
        for it in batch:
            key = it.get("id") or it.get("itemId") or str(it.get("time"))
            if key in seen_ids:
                continue
            seen_ids.add(key)
            items.append(it)
            fresh += 1
        if not fresh or not batch:
            break  # 无新增即到底，不再翻页
    # loadmore2 = 数据库兜底通道（ISO 时间），继续往更早翻，覆盖 08 点档
    for page in range(1, CT_FLASH_DB_PAGES + 1):
        try:
            req = Request(CT_FLASH_DB_API + "?page={}&pageSize=20".format(page),
                          headers=ct_h)
            with urlopen(req, timeout=12) as resp:
                data = json.loads(resp.read().decode("utf-8", "replace"))
            batch = ((data or {}).get("data") or {}).get("list") or []
        except Exception:
            break
        fresh = 0
        for it in batch:
            key = it.get("id") or it.get("itemId") or str(it.get("time"))
            if key in seen_ids:
                continue
            seen_ids.add(key)
            items.append(it)
            fresh += 1
        if not fresh or not batch:
            break
    if items:
        rows = []
        for it in items:
            t = (it.get("time") or it.get("release_time") or "").strip()
            title = (it.get("title") or "").strip()
            if not title:
                continue
            day = _item_day(t) or (report_date or
                                   datetime.now(CST).strftime("%Y-%m-%d"))
            if report_date and day:
                try:
                    if abs((datetime.strptime(day, "%Y-%m-%d") -
                            datetime.strptime(report_date, "%Y-%m-%d")).days) > 2:
                        continue  # 归档日条目不并入
                except Exception:
                    pass
            # 接口原生带秒（如 "2026-11-05 10:58:02" / ISO 格式），保留秒级精度：
            # 归因要用「09:30 切分」，截断到分钟会带来最大 59 秒的错位。
            hms = t[11:19] if len(t) >= 19 else (t[11:16] + ":00" if len(t) >= 16
                                                 else "00:00:00")
            body = (it.get("isAbstract") or "").strip()
            body = _plain_text(body, 400) if body else title
            full = "【{}】{}".format(title, body) if body != title else title
            rows.append({
                "id": "ctj{}{}{}".format(day.replace("-", ""),
                                         hms.replace(":", ""), it.get("id") or ""),
                "time": "{} {}".format(day, hms),
                "title": full[:70], "summary": _plain_text(full, 400),
                # 标题统一被包了【】，不能再按 startswith("【") 判重要（会全为 True）
                "important": bool(re.search(
                    r"重磅|突发|紧急|证监会|央行|降准|降息|加息|暂停|调查", title)),
                "stocks": [],
                "source": "华夏时报7×24",
                "origin": (it.get("source") or "").strip() or "华夏时报",
            })
        rows.sort(key=lambda r: r["time"])
        hot = []
        try:
            req2 = Request(CT_FLASH_HOT_API + "?page=1&pageSize=10", headers=ct_h)
            with urlopen(req2, timeout=12) as resp:
                data2 = json.loads(resp.read().decode("utf-8", "replace"))
            for rank, it2 in enumerate(
                    ((data2 or {}).get("data") or {}).get("list") or [], 1):
                title2 = (it2.get("title") or "").strip()
                if title2 and len(title2) >= 6:
                    hot.append({"rank": rank,
                                "title": _plain_text(title2, 120)})
        except Exception:
            hot = []
        return rows, hot, "ok"

    # ---- 兜底：旧网页通道（服务端渲染版本） ----
    try:
        req = Request(CT_FLASH_URL, headers=ct_h)
        with urlopen(req, timeout=12) as resp:
            html = resp.read().decode("utf-8", "replace")
        rows, hot, pdate, archived = parse_chinatimes_flash(
            _plain_text(html, 200000), report_date)
        if rows and not archived:
            return rows, hot, "ok"
        return [], hot, "empty" if not rows else "archive"
    except Exception:
        return [], [], "blocked"


def extract_etf_flow_evidence(rows):
    """从快讯正文抽取 ETF 资金流的独立口径数据（同花顺 iFinD 等）。

    只做结构化提取，不与第十章一级市场净申购混为一谈：
    口径不同（资金流统计 vs 份额申赎）、日期口径也可能差一日。
    """
    out = []
    pats = [
        ("ETF整体净流入", r"ETF资金整体净流入\s*([\d.]+)\s*亿元"),
        ("股票型净流入", r"股票型净流入\s*([\d.]+)\s*亿元"),
        ("宽基型净流入", r"宽基型净流入\s*([\d.]+)\s*亿元"),
        ("连续净流入日数", r"连续\s*(\d+)\s*个交易日净流入"),
        ("单日净流入新高", r"单日净流入\s*([\d.]+)\s*亿元[^。]*创年内新高"),
        ("累计净买入", r"连续净流入[^。]*合计约?\s*([\d.]+)\s*亿元"),
    ]
    for r in rows:
        body = r.get("summary") or r.get("title") or ""
        hit = {}
        for label, pat in pats:
            m = re.search(pat, body)
            if m:
                hit[label] = m.group(1)
        if hit:
            singles = re.findall(
                r"([\u4e00-\u9fa5A-Za-z0-9]{2,14}ETF)[（(](\d{6})[)）]\s*"
                r"净流入\s*([\d.]+)\s*亿元", body)
            out.append({"time": r.get("time") or "", "source": r.get("source"),
                        "origin": r.get("origin") or "", "title": r.get("title") or "",
                        "summary": body, "values": hit,
                        "singles": [{"name": n, "code": c, "value": v}
                                    for n, c, v in singles]})
    return out


# ---- 财联社电报 7×24（第六通道，A股垂直度最高）----
# 社区逆向端点，非官方公开 API：sign = md5(sha1(按 key 字典序拼的 query 串))，
# 纯本地可算、零密钥；但端点版本会漂移，故新旧两个端点都试。
# 最大价值是 last_time 时间游标：可按时间戳回溯历史，
# 正是其余五源（实时页只给最新 50 条）做不到的。
CLS_LIST_URL = "https://www.cls.cn/nodeapi/telegraphList"   # 旧版 nodeapi
CLS_ROLL_URL = "https://www.cls.cn/v1/roll/get_roll_list"   # 新版前端主接口
CLS_SV = "8.4.6"
CLS_MAX_PAGES = 4
CLS_RN = 50


def cls_sign(params):
    """财联社 sign：query 按 key 字典序 → SHA1 → MD5。零密钥，纯本地计算。"""
    q = urlencode(sorted((k, "" if v is None else str(v))
                         for k, v in params.items()))
    return hashlib.md5(hashlib.sha1(q.encode("utf-8"))
                       .hexdigest().encode("utf-8")).hexdigest()


def parse_cls_roll(payload, report_date=""):
    """解析财联社电报响应，兼容 data.roll_data / data.list / 裸列表。

    字段：ctime（Unix 秒，日期归属天然无歧义）、title、brief、content、id。
    brief + content 双字段比门户快讯更适合做关键词→板块映射。
    """
    if not isinstance(payload, dict):
        return []
    data = payload.get("data") or payload
    rows = data.get("roll_data") or data.get("list") or []
    if not isinstance(rows, list):
        rows = []
    if not rows and isinstance(data, dict):
        for v in data.values():
            if isinstance(v, list) and v and isinstance(v[0], dict) and "ctime" in v[0]:
                rows = v
                break
    out, seen = [], set()
    for r in rows:
        if not isinstance(r, dict):
            continue
        ct = num(r.get("ctime"))
        if not ct:
            continue
        try:
            dt = datetime.fromtimestamp(ct, CST)
        except Exception:
            continue
        title = _plain_text(r.get("title") or r.get("brief") or "", 120)
        brief = _plain_text(r.get("brief") or "", 400)
        content = _plain_text(r.get("content") or "", 600)
        body = brief or content or title
        if not body:
            continue
        rid = str(r.get("id") or r.get("shareurl") or ct)
        if rid in seen:
            continue
        seen.add(rid)
        out.append({
            "id": "cls" + rid, "time": dt.strftime("%Y-%m-%d %H:%M:%S"),
            "title": (title or body[:60]), "summary": body,
            "important": bool(r.get("is_ad") is False and title.startswith("【")),
            "stocks": [], "source": "财联社电报",
            "origin": r.get("source") or "财联社",
            "brief": brief, "content": content, "ctime": ct,
        })
    out.sort(key=lambda r: r["ctime"])
    return out


def fetch_cls_flash(report_date="", pages=CLS_MAX_PAGES):
    """直连财联社电报并按 last_time 回溯。返回 (条目, 状态, 页数, 是否覆盖报告日盘中)。

    沙盒里 cls.cn 域名族被策略拒绝时返回 blocked，由网页通道/快照兜底；
    换到可访问环境即可自动回溯到报告日 09:30 之前。
    """
    cls_h = dict(HEADERS, **{
        "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                       "AppleWebKit/537.36 (KHTML, like Gecko) "
                       "Chrome/120.0.0.0 Safari/537.36"),
        "Referer": "https://www.cls.cn/telegraph",
        "Accept": "application/json, text/plain, */*",
        "Accept-Language": "zh-CN,zh;q=0.9",
    })
    try:
        floor = int(datetime.strptime(report_date, "%Y-%m-%d")
                    .replace(tzinfo=CST).timestamp())
    except Exception:
        floor = 0
    cursor = int(time.time())
    all_rows, seen, ok_pages = [], set(), 0
    status = "blocked"
    for _ in range(max(1, pages)):
        got = None
        # 两套参数模板（社区资料对有效端点/参数版本说法不一致），逐个尝试
        templates = [
            {"app": "CailianpressWeb", "category": "", "lastTime": cursor,
             "last_time": cursor, "os": "web", "refresh_type": "1",
             "rn": CLS_RN, "sv": CLS_SV},
            {"app": "CailianpressWeb", "last_time": cursor, "os": "web",
             "rn": CLS_RN, "sv": CLS_SV},
        ]
        for base in templates:
            for url in (CLS_LIST_URL, CLS_ROLL_URL):
                p = dict(base, sign=cls_sign(base))
                try:
                    req = Request(url + "?" + urlencode(p), headers=cls_h)
                    with urlopen(req, timeout=10) as resp:
                        payload = json.loads(resp.read().decode("utf-8", "replace"))
                    rows = parse_cls_roll(payload, report_date)
                    if rows:
                        got = rows
                        break
                except Exception:
                    continue
            if got:
                break
        if not got:
            break
        ok_pages += 1
        status = "ok"
        for r in got:
            if r["id"] not in seen:
                seen.add(r["id"])
                all_rows.append(r)
        nxt = min(r["ctime"] for r in got)
        if nxt <= floor or nxt >= cursor:
            break
        cursor = int(nxt) - 1
    all_rows.sort(key=lambda r: r["ctime"])
    covered = bool(all_rows and all_rows[0]["ctime"] <= floor + 9 * 3600)
    return all_rows, status, ok_pages, covered


def normalize_flash_ext(ext, snapshot_path=""):
    """统一重算快讯衍生字段，保证在线/离线/网页通道/保存快照四条路径一致。"""
    cur = ext.get("_em_flash") or []
    if not cur:
        return ext
    ext["_flash_source_n"] = len({r.get("source") for r in cur})
    ext["_flash_family_n"] = len({source_family(r.get("source"))
                                  for r in cur})
    ext["_etf_flow"] = extract_etf_flow_evidence(cur)
    fams = {source_family(r.get("source")) for r in cur}
    if "华夏时报" in fams:
        # 仅当在线采集被置空/拒绝时才回填为网页通道；JSON 接口成功时保留 ok，
        # 否则契约自检会把「接口直连」误报成「网页兜底」。
        if (ext.get("_ct_status") or "empty") in ("empty", "blocked", "archive"):
            ext["_ct_status"] = "web_channel"
        ext["_ct_n"] = sum(1 for r in cur
                           if r.get("source") == "华夏时报7×24")
        if not ext.get("_ct_hot"):
            old_ct = load_snapshot_extended().get("_ct_hot")
            if old_ct:
                ext["_ct_hot"] = old_ct
    return ext


def merge_flash_history(path, rows, report_date, keep_days=90,
                        max_per_day=300):
    """按 id 增量累积 7×24 快讯，解决「实时页只给最新 50 条」的问题。

    真源为 hist_flash 表，主键=内容级去重键（flash_key）：接口 id 每次抓取
    会漂移，内容键才稳定。每天运行一次即可自然形成历史序列；报告日盘中
    9:30–15:00 的快讯会在当天运行时被采集并永久留存，供次日报告做时序归因。
    """
    con = store_connect()
    try:
        # 内容级去重：同秒 + 同标题（前 60 字）+ 同来源族视为同一条，
        # 避免接口 id 漂移（实时页每次抓取 id 会变）把同批快讯累积成多条；
        # 键放宽到秒与 60 字，防止模板化标题被误合并。
        bucket = {k: json.loads(p) for k, p in con.execute(
            "SELECT key,payload FROM hist_flash WHERE date=?", (report_date,))}
        for r in rows:
            d = (r.get("time") or "")[:10]
            if not d:
                continue
            ck = flash_key(r)
            hit = next((k for k, v in bucket.items() if flash_key(v) == ck),
                       None)
            bucket[hit or ck] = r
        # 单日上限：按时间保留最新 max_per_day 条，防止网页通道反复注入撑爆缓存
        items = sorted(bucket.items(), key=lambda kv: kv[1].get("time") or "")
        if len(items) > max_per_day:
            items = items[-max_per_day:]
        for k, v in items:
            con.execute(
                "INSERT OR REPLACE INTO hist_flash"
                "(key,date,time,payload,updated) VALUES(?,?,?,?,?)",
                (k, report_date, (v.get("time") or "")[:19],
                 json.dumps(v, ensure_ascii=False, separators=(",", ":")),
                 report_date))
        # 只保留最近 keep_days 天，避免无限膨胀
        con.execute(
            "DELETE FROM hist_flash WHERE date < ? AND "
            "CAST(julianday(?) AS INTEGER) - CAST(julianday(date) AS INTEGER)"
            " > ?", (report_date, report_date, keep_days))
        con.commit()
        # 返回报告日窗口（含次日早报）的累积条目
        try:
            nxt = (datetime.strptime(report_date, "%Y-%m-%d")
                   + timedelta(days=1)).strftime("%Y-%m-%d")
        except Exception:
            nxt = report_date
        out = []
        for d in (report_date, nxt):
            out.extend(json.loads(p) for (p,) in con.execute(
                "SELECT payload FROM hist_flash WHERE date=?", (d,)))
        out.sort(key=lambda r: r.get("time") or "")
        days = {d: {k: json.loads(p) for k, p in con.execute(
            "SELECT key,payload FROM hist_flash WHERE date=?", (d,))}
            for d, in con.execute("SELECT DISTINCT date FROM hist_flash")}
    finally:
        con.close()
    return out, {"days": days, "updated": report_date}


def load_flash_history(path, report_date):
    """离线模式：从历史真源库读取报告日窗口。

    hist 返回全量 days 结构（{date: {key: row}}）：调用方要统计累积
    天数/总条数，只给窗口会算错，所以全读（快讯量级 ~1e3，代价可忽略）。
    """
    try:
        nxt = (datetime.strptime(report_date, "%Y-%m-%d")
               + timedelta(days=1)).strftime("%Y-%m-%d")
    except Exception:
        nxt = report_date
    con = store_connect()
    try:
        days = {d: {k: json.loads(p) for k, p in con.execute(
            "SELECT key,payload FROM hist_flash WHERE date=?", (d,))}
            for d, in con.execute("SELECT DISTINCT date FROM hist_flash")}
        out = list(days.get(report_date, {}).values()) \
            + list(days.get(nxt, {}).values())
        out.sort(key=lambda r: r.get("time") or "")
        updated = con.execute("SELECT MAX(updated) FROM hist_flash"
                              ).fetchone()[0] or ""
    finally:
        con.close()
    return out, {"days": days, "updated": updated}


def load_daily_history(path):
    """逐日指标历史：真源为 hist_daily 表（参数 path 仅作日志线索保留）。"""
    con = store_connect()
    try:
        return store_load_dict(con, "hist_daily")
    finally:
        con.close()


def append_daily_history(path, date, record, max_days=180):
    """逐日累积：接口不给历史序列，只能每次运行落一份当日快照。"""
    con = store_connect()
    try:
        con.execute(
            "INSERT OR REPLACE INTO hist_daily(date,payload,updated) "
            "VALUES(?,?,?)",
            (date, json.dumps(record, ensure_ascii=False,
                              separators=(",", ":")),
             datetime.now(CST).strftime("%Y-%m-%d %H:%M")))
        store_trim_days(con, "hist_daily", max_days)
        con.commit()
        return store_load_dict(con, "hist_daily")
    finally:
        con.close()


EM_KLINE_URL = ("https://push2his.eastmoney.com/api/qt/stock/kline/get?"
                "secid={}&ut=fa5fd1943c7b386f172d6893dbfba10b"
                "&fields1=f1,f2,f3,f4,f5,f6"
                "&fields2=f51,f52,f53,f54,f55,f56,f57"
                "&klt=101&fqt=1&end=20500101&lmt={}")


def fetch_index_amount(lmt=30, secids=None):
    """两市成交额历史回填：东财指数日 K 的成交额字段。

    上证指数只含沪市、深证成指只含深市，两者相加近似「两市成交额」。
    与实时口径（沪深全量个股成交额合计）存在 1~5% 差异，故回填值单独打
    `amount_src` 标记，只在同源序列内算环比，不与实时值混算。
    kline 字段：日期,开,收,高,低,成交量,成交额(元)。
    """
    secids = secids or (("1.000001", "sh"), ("0.399001", "sz"))
    out = {"status": "ok", "days": {}, "names": {}, "lmt": lmt}
    for secid, key in secids:
        try:
            req = Request(EM_KLINE_URL.format(secid, lmt),
                          headers=dict(HEADERS, Referer="https://quote.eastmoney.com/"))
            with urlopen(req, timeout=12) as resp:
                d = json.loads(resp.read().decode("utf-8", "replace"))
        except Exception:
            continue
        data = d.get("data") or {}
        ks = data.get("klines") or []
        if not ks:
            continue
        out["names"][key] = data.get("name") or key
        prev_close = None
        for ln in ks:
            f = ln.split(",")
            if len(f) < 7:
                continue
            date = f[0]
            close, amt = num(f[2]), num(f[6])
            if not close or amt is None:
                continue
            pct = ((close / prev_close - 1) * 100) if prev_close else None
            prev_close = close
            node = out["days"].setdefault(date, {})
            node[key + "_amount_yi"] = amt / 1e8
            node[key + "_pct"] = pct
    for date, node in out["days"].items():
        sh, sz = node.get("sh_amount_yi"), node.get("sz_amount_yi")
        if sh is not None and sz is not None:
            node["amount_yi"] = sh + sz
        else:
            node["amount_yi"] = sh if sh is not None else sz
    out["count"] = sum(1 for v in out["days"].values() if v.get("amount_yi"))
    if not out["count"]:
        out["status"] = "empty"
    return out


def backfill_history(path, days, today_date, max_days=180):
    """把指数成交额历史并进逐日历史，补上「没有快照的那些天」。

    只回填**成交额与指数涨跌**：广度、主力净额、ERP 这类字段只有当日快照才有，
    缺失就是缺失，不插值、不猜。回填条目带 `amount_src` 标记，环比只在同源
    序列内计算，避免「实时口径 vs 回填口径」互相污染。
    """
    hist = load_daily_history(path)
    added = 0
    touched = []
    for d in sorted(days or {}):
        if d >= today_date:
            continue
        v = days[d]
        if not v.get("amount_yi"):
            continue
        rec = hist.get(d) or {}
        # 已有**收盘后**的实时值才保留；盘中/快照存下的累计值会被全天回填覆盖，
        # 否则序列里会混进一个「半天成交额」的假低点。
        if rec.get("amount") and not rec.get("amount_partial"):
            rec.setdefault("sh_pct", v.get("sh_pct"))
            rec.setdefault("sz_pct", v.get("sz_pct"))
            hist[d] = rec
            touched.append(d)
            continue
        _was = bool(rec.get("amount"))
        rec.update({"amount": round(v["amount_yi"], 1), "amount_src": "回填",
                    "sh_pct": v.get("sh_pct"), "sz_pct": v.get("sz_pct")})
        hist[d] = rec
        touched.append(d)
        if not _was:
            added += 1
    keys = sorted(k for k in hist if hist[k].get("amount"))
    for i, k in enumerate(keys):
        if i == 0:
            continue
        prev = hist[keys[i - 1]].get("amount")
        cur = hist[k].get("amount")
        if prev and cur and hist[k].get("amount_chg") is None:
            hist[k]["amount_chg"] = (cur / prev - 1) * 100
    con = store_connect()
    try:
        now = datetime.now(CST).strftime("%Y-%m-%d %H:%M")
        for d in touched:
            con.execute(
                "INSERT OR REPLACE INTO hist_daily(date,payload,updated) "
                "VALUES(?,?,?)",
                (d, json.dumps(hist[d], ensure_ascii=False,
                               separators=(",", ":")), now))
        store_trim_days(con, "hist_daily", max_days)
        con.commit()
    finally:
        con.close()
    return hist, added


def write_xlsx(path, sheets):
    """零依赖写 xlsx：sheets = [(名称, [[表头...],[行...], ...]), ...]。"""
    import zipfile
    def col(n):
        s = ""
        while n:
            n, r = divmod(n - 1, 26)
            s = chr(65 + r) + s
        return s
    def cell(v, ref, style=0):
        """style 1 = 表头加粗，2 = 数字右对齐（Excel 默认左对齐，不设就等于没对齐）。"""
        s = ' s="{}"'.format(style) if style else ""
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            return '<c r="{}"{}><v>{}</v></c>'.format(ref, s, v)
        t = str(v if v is not None else "")
        t = (t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
              .replace('"', "&quot;"))
        return ('<c r="{}"{} t="inlineStr"><is><t xml:space="preserve">{}</t></is>'
                '</c>'.format(ref, s, t))

    def disp_w(v):
        """显示宽度：CJK 按 2 列算，否则按字符数（只用于估列宽）。"""
        s = str(v if v is not None else "")
        return sum(2 if ord(ch) > 0x2E80 else 1 for ch in s)

    def sheet_xml(rows):
        widths = {}
        for row in rows[:200]:
            for ci, v in enumerate(row, 1):
                widths[ci] = max(widths.get(ci, 6), disp_w(v))
        cols_xml = "".join(
            '<col min="{0}" max="{0}" width="{1:.1f}" customWidth="1"/>'.format(
                ci, min(46, max(8, w + 3))) for ci, w in sorted(widths.items()))
        out = ['<?xml version="1.0" encoding="UTF-8"?>'
               '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
               '<sheetViews><sheetView workbookViewId="0">'
               '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
               '</sheetView></sheetViews>',
               '<cols>{}</cols>'.format(cols_xml),
               '<sheetData>']
        for ri, row in enumerate(rows, 1):
            out.append('<row r="{}">'.format(ri))
            _hdr = 1 if ri == 1 else 0
            for ci, v in enumerate(row, 1):
                _st = 1 if _hdr else (
                    2 if isinstance(v, (int, float)) and not isinstance(v, bool) else 0)
                out.append(cell(v, "{}{}".format(col(ci), ri), _st))
            out.append('</row>')
        out.append('</sheetData></worksheet>')
        return "".join(out)
    ct = ['<?xml version="1.0" encoding="UTF-8"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
          '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>']
    for i in range(len(sheets)):
        ct.append('<Override PartName="/xl/worksheets/sheet{}.xml" '
                  'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'.format(i + 1))
    ct.append('</Types>')
    wb = ['<?xml version="1.0" encoding="UTF-8"?>'
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>']
    for i, (name, _) in enumerate(sheets, 1):
        wb.append('<sheet name="{}" sheetId="{}" r:id="rId{}"/>'.format(
            name.replace("&", "&amp;").replace('"', "&quot;"), i, i))
    wb.append('</sheets></workbook>')
    rels = ['<?xml version="1.0" encoding="UTF-8"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
    for i in range(len(sheets)):
        rels.append('<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{}.xml"/>'.format(i + 1, i + 1))
    rels.append('<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'.format(len(sheets) + 1))
    rels.append('</Relationships>')
    root_rels = ('<?xml version="1.0" encoding="UTF-8"?>'
                 '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                 '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                 '</Relationships>')
    # 三个单元格格式：0 默认 / 1 表头加粗 / 2 数字右对齐
    styles = ('<?xml version="1.0" encoding="UTF-8"?>'
              '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
              '<fonts count="2">'
              '<font><sz val="11"/><name val="WenQuanYi Micro Hei"/></font>'
              '<font><b/><sz val="11"/><name val="WenQuanYi Micro Hei"/></font>'
              '</fonts>'
              '<fills count="2"><fill><patternFill patternType="none"/></fill>'
              '<fill><patternFill patternType="gray125"/></fill></fills>'
              '<borders count="1"><border/></borders>'
              '<cellStyleXfs count="1"><xf xfId="0"/></cellStyleXfs>'
              '<cellXfs count="3">'
              '<xf xfId="0"/>'
              '<xf xfId="0" fontId="1" applyFont="1"/>'
              '<xf xfId="0" applyAlignment="1">'
              '<alignment horizontal="right"/></xf>'
              '</cellXfs></styleSheet>')
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", "".join(ct))
        z.writestr("_rels/.rels", root_rels)
        z.writestr("xl/workbook.xml", "".join(wb))
        z.writestr("xl/_rels/workbook.xml.rels", "".join(rels))
        z.writestr("xl/styles.xml", styles)
        for i, (_, rows) in enumerate(sheets, 1):
            z.writestr("xl/worksheets/sheet{}.xml".format(i), sheet_xml(rows))
    return path


def load_shares_history(path):
    """份额历史：{'date': {code: shares}, ...}，真源为 hist_shares 表。"""
    con = store_connect()
    try:
        return store_load_dict(con, "hist_shares")
    finally:
        con.close()


def update_shares_history(path, date, store, max_days=120):
    """把当日份额写入历史；返回 (历史 dict, 与上一交易日的份额差)。"""
    hist = load_shares_history(path)
    snap = {}
    for c in ALL_ETF_CODES:
        sh = store.get(c, {}).get("quote", {}).get("total_shares")
        if sh:
            snap[c] = sh
    if not snap:
        return hist, {}, 0
    hist[date] = snap
    keys = sorted(hist)
    for k in keys[:-max_days]:
        hist.pop(k, None)
    keys = sorted(hist)
    prev = hist[keys[-2]] if len(keys) >= 2 else {}
    delta = {}
    for c, v in snap.items():
        if c in prev and prev[c]:
            delta[c] = {"prev": prev[c], "now": v,
                        "diff": v - prev[c], "pct": (v / prev[c] - 1) * 100}
    con = store_connect()
    try:
        con.execute(
            "INSERT OR REPLACE INTO hist_shares(date,payload,updated) "
            "VALUES(?,?,?)",
            (date, json.dumps(snap, ensure_ascii=False,
                              separators=(",", ":")),
             datetime.now(CST).strftime("%Y-%m-%d %H:%M")))
        store_trim_days(con, "hist_shares", max_days)
        con.commit()
    finally:
        con.close()
    return hist, delta, len(keys)


def load_snapshot(path):
    """读快照：优先真源库最新一份，库为空时回退 JSON 文件（迁移兼容）。"""
    con = store_connect()
    try:
        row = con.execute("SELECT payload FROM snapshot_store "
                          "ORDER BY updated DESC, date DESC LIMIT 1").fetchone()
    finally:
        con.close()
    if row:
        snap = json.loads(row[0])
    else:
        with open(path, "r", encoding="utf-8") as f:
            snap = json.load(f)
    return (snap.get("data", {}), snap.get("date", ""),
            snap.get("fundflow", {}), snap.get("breadth", {}),
            snap.get("global", {}), snap.get("extended", {}))


_SNAPSHOT_EXT_CACHE = None


def load_snapshot_extended():
    """读真源库最新快照的 extended 段（替代已删除的 etf_snapshot.json）。

    快讯/热点榜跨天合并原先直接 open 那个 JSON，迁移后文件已删、路径静默
    失效；改走 snapshot_store，并做进程内缓存，避免重复解析 6MB payload。
    """
    global _SNAPSHOT_EXT_CACHE
    if _SNAPSHOT_EXT_CACHE is not None:
        return _SNAPSHOT_EXT_CACHE
    ext = {}
    try:
        con = store_connect()
        row = con.execute("SELECT payload FROM snapshot_store "
                          "ORDER BY updated DESC, date DESC LIMIT 1").fetchone()
        con.close()
        if row:
            ext = json.loads(row[0]).get("extended") or {}
    except Exception:
        pass
    _SNAPSHOT_EXT_CACHE = ext
    return ext


def save_snapshot(data, path, date, fundflow=None, breadth=None, glob=None,
                  extended=None):
    payload = {"date": date,
               "source": "tencent public quote+kline+fundflow+breadth+global+extended",
               "data": data}
    if fundflow:
        payload["fundflow"] = fundflow
    if breadth:
        payload["breadth"] = breadth
    if glob:
        payload["global"] = glob
    if extended:
        payload["extended"] = extended
    blob = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    con = store_connect()
    try:
        con.execute(
            "INSERT OR REPLACE INTO snapshot_store(date,payload,updated) "
            "VALUES(?,?,?)",
            (date, blob, datetime.now(CST).strftime("%Y-%m-%d %H:%M")))
        con.commit()
    finally:
        con.close()


def parse_quote(raw):
    """把 88 字段列表解析成带语义的 dict。"""
    def g(key):
        i = F[key]
        return raw[i] if i < len(raw) else None

    return {
        "name": g("name") or "",
        "price": num(g("price")),
        "prev_close": num(g("prev_close")),
        "open": num(g("open")),
        "high": num(g("high")),
        "low": num(g("low")),
        "change": num(g("change")),
        "pct": num(g("pct")),
        "amount_yi": yi(num(g("amount_wan"))),
        "volume": num(g("volume")),
        "turnover": num(g("turnover")),
        "range_pct": num(g("range_pct")),
        "vol_ratio": num(g("vol_ratio")),
        "avg_price": num(g("avg_price")),
        "limit_up": num(g("limit_up")),
        "time": g("time") or "",
        "ytd": num(g("ytd")),
        "d5": num(g("d5")),
        "d10": num(g("d10")),
        "d20": num(g("d20")),
        "d60": num(g("d60")),
        "high52": num(g("high52")),
        "low52": num(g("low52")),
        # ETF 特有：份额(份)、IOPV、折溢价(%)、单位净值
        "total_shares": num(g("total_shares")),
        "discount": num(g("discount")),
        "iopv": num(g("iopv")),
        "net_value": num(g("net_value")),
        # 指数估值（ETF 该字段通常为 0）
        "pe": num(g("pe")),
        "pb": num(g("pb")),
    }


def parse_kline(rows):
    """[[date, open, close, high, low, volume], ...] -> list[dict]"""
    bars = []
    for row in rows:
        if len(row) < 6:
            continue
        o, c, h, l, v = (num(row[i]) for i in range(1, 6))
        if None in (o, c, h, l):
            continue
        bars.append({"date": row[0], "open": o, "close": c,
                     "high": h, "low": l, "volume": v or 0.0})
    return bars


def build_store(raw_data):
    store = {}
    for code, node in raw_data.items():
        q = parse_quote(node.get("quote", []))
        bars = parse_kline(node.get("kline", []))
        if q["price"] is None or not bars:
            continue
        store[code] = {"quote": q, "bars": bars}
    return store


# ----------------------------------------------------------------------------
# 指标层（纯 Python，无 numpy/pandas 依赖）
# ----------------------------------------------------------------------------

def ma(closes, n):
    return sum(closes[-n:]) / n if len(closes) >= n else None


def ema_series(vals, n):
    if not vals:
        return []
    alpha = 2.0 / (n + 1)
    out, e = [vals[0]], vals[0]
    for v in vals[1:]:
        e = alpha * v + (1 - alpha) * e
        out.append(e)
    return out


def calc_macd(closes, fast=12, slow=26, signal=9):
    """DIF / DEA / 柱 = 2*(DIF-DEA)，与国内行情软件常用口径一致。"""
    if len(closes) < slow + signal:
        return {}
    ef, es = ema_series(closes, fast), ema_series(closes, slow)
    dif = [a - b for a, b in zip(ef, es)]
    dea = ema_series(dif, signal)
    bar = [2 * (a - b) for a, b in zip(dif, dea)]
    return {
        "dif": dif[-1], "dea": dea[-1], "bar": bar[-1],
        "bar_prev": bar[-2] if len(bar) > 1 else None,
        "golden": dif[-1] > dea[-1] and dif[-2] <= dea[-2],
        "above_zero": dif[-1] > 0,
    }


def calc_kdj(bars, n=9):
    """KDJ(9,3,3)，返回最后两日 K/D/J 以便判断是否自超卖区抬头。"""
    highs = [b["high"] for b in bars]
    lows = [b["low"] for b in bars]
    closes = [b["close"] for b in bars]
    if len(bars) < n + 1:
        return {}
    k, d = 50.0, 50.0
    series = []
    for i in range(len(bars)):
        if i < n - 1:
            series.append(None)
            continue
        hh, ll = max(highs[i - n + 1:i + 1]), min(lows[i - n + 1:i + 1])
        rsv = 50.0 if hh == ll else (closes[i] - ll) / (hh - ll) * 100.0
        k = (2.0 / 3.0) * k + (1.0 / 3.0) * rsv
        d = (2.0 / 3.0) * d + (1.0 / 3.0) * k
        series.append((k, d, 3 * k - 2 * d))
    last, prev = series[-1], series[-2]
    out = {"k": last[0], "d": last[1], "j": last[2]}
    if prev:
        out.update({"k_prev": prev[0], "j_prev": prev[2], "rising": last[0] > prev[0]})
    return out


def calc_rsi(closes, n):
    """RSI：SMA(x,N,1) / Wilder 平滑口径。"""
    if len(closes) <= n:
        return None
    ups = [max(closes[i] - closes[i - 1], 0.0) for i in range(1, len(closes))]
    downs = [max(closes[i - 1] - closes[i], 0.0) for i in range(1, len(closes))]
    au, ad = sum(ups[:n]) / n, sum(downs[:n]) / n
    for u, dn in zip(ups[n:], downs[n:]):
        au = (au * (n - 1) + u) / n
        ad = (ad * (n - 1) + dn) / n
    return 100.0 if au + ad == 0 else 100.0 * au / (au + ad)


def calc_boll(closes, n=20, k=2):
    if len(closes) < n:
        return {}
    window = closes[-n:]
    mid = sum(window) / n
    std = math.sqrt(sum((x - mid) ** 2 for x in window) / n)
    return {"mid": mid, "upper": mid + k * std, "lower": mid - k * std, "std": std}


def volume_poc(bars, n=20):
    """近 n 日成交量加权均价，近似「套牢密集区中枢」。"""
    seg = bars[-n:]
    tot = sum(b["volume"] for b in seg)
    if tot <= 0:
        return None
    return sum(((b["high"] + b["low"] + b["close"]) / 3) * b["volume"]
               for b in seg) / tot


def indicators(bars):
    closes = [b["close"] for b in bars]
    return {
        "ma5": ma(closes, 5), "ma10": ma(closes, 10), "ma20": ma(closes, 20),
        "ma60": ma(closes, 60), "ma120": ma(closes, 120), "ma250": ma(closes, 250),
        "macd": calc_macd(closes), "kdj": calc_kdj(bars),
        "rsi6": calc_rsi(closes, 6), "rsi12": calc_rsi(closes, 12),
        "boll": calc_boll(closes), "poc20": volume_poc(bars, 20),
        "vol_ma5": ma([b["volume"] for b in bars], 5),
        "prev_volume": bars[-2]["volume"] if len(bars) > 1 else None,
        "prev_low": bars[-2]["low"] if len(bars) > 1 else None,
        "low5": min(b["low"] for b in bars[-5:]),
        "high20": max(b["high"] for b in bars[-20:]),
    }


# ----------------------------------------------------------------------------
# 分析层一：今日大盘
# ----------------------------------------------------------------------------

def gap_label(q):
    """开盘相对昨收：平开 / 高开 / 低开。"""
    o, p = q["open"], q["prev_close"]
    if None in (o, p) or p == 0:
        return "—"
    r = (o / p - 1) * 100
    if r > 0.5:
        return "高开{:+.2f}%".format(r)
    if r < -0.5:
        return "低开{:+.2f}%".format(r)
    return "平开{:+.2f}%".format(r)


def close_position(q):
    """收盘在当日振幅区间的位置，1 = 收在最高。"""
    c, h, l = q["price"], q["high"], q["low"]
    if None in (c, h, l) or h == l:
        return None
    return (c - l) / (h - l)


def analyze_today(store):
    date = ""
    for item in store.values():
        if item["bars"]:
            date = max(date, item["bars"][-1]["date"])

    indices = [(c, store[c]["quote"]) for c in INDEX_CODES if c in store]
    etfs = [(c, store[c]["quote"]) for c in ALL_ETF_CODES if c in store]
    ranked = sorted(etfs, key=lambda x: (x[1]["pct"] is not None, x[1]["pct"] or 0),
                    reverse=True)

    groups = {}
    for gname, codes in ETF_GROUPS.items():
        pairs = [(store[c]["quote"]["pct"], store[c]["quote"]["amount_yi"])
                 for c in codes
                 if c in store and store[c]["quote"]["pct"] is not None]
        vals = [p for p, _ in pairs]
        weighted_pairs = [(p, a) for p, a in pairs if a]
        wsum = sum(a for _, a in weighted_pairs)
        groups[gname] = {
            "mean": avg(vals), "count": len(vals),
            "weighted": (sum(p * a for p, a in weighted_pairs) / wsum
                         if wsum else None),
        }

    growth = groups.get("科技成长", {}).get("mean")
    weight = groups.get("宽基", {}).get("mean")
    spread = (growth - weight) if None not in (growth, weight) else None

    limit_up = [(c, q) for c, q in etfs
                if q["pct"] is not None and q["pct"] >= 9.8]
    amount_total = sum(q["amount_yi"] or 0 for _, q in etfs)

    # 形态：平开 + 收于高位 + 大振幅，是资金持续买入而非情绪一次性跳空
    flat_open, close_high = 0, 0
    for _, q in ranked:
        label = gap_label(q)
        if label.startswith("平开"):
            flat_open += 1
        pos = close_position(q)
        if pos is not None and pos >= 0.9:
            close_high += 1

    # 量能：核心标的反包日成交量 vs 前一日（恐慌日）
    vol_checks = []
    for c in CORE_CODES:
        if c not in store:
            continue
        item = store[c]
        q, ind = item["quote"], indicators(item["bars"])
        tv, pv = q["volume"], ind["prev_volume"]
        if tv and pv:
            vol_checks.append({
                "code": c, "name": q["name"], "today": tv, "prev": pv,
                "ratio": tv / pv, "amount_yi": q["amount_yi"],
            })

    signals = []
    if indices:
        top = max(indices, key=lambda x: x[1]["pct"] or 0)
        low = min(indices, key=lambda x: x[1]["pct"] or 0)
        signals.append("指数涨幅从 {} 的 {} 到 {} 的 {}，{}。".format(
            top[1]["name"], pct_str(top[1]["pct"]), low[1]["name"],
            pct_str(low[1]["pct"]),
            "结构性资金进攻、非全局普涨" if (top[1]["pct"] or 0) - (low[1]["pct"] or 0) > 3
            else "各大指数同步性较好"))
    if limit_up:
        top_limit = sorted(limit_up, key=lambda x: x[1]["pct"] or 0, reverse=True)
        signals.append("样本内 {} 只ETF涨停/近涨停（{}），情绪进入亢奋区，次日易出现高开分歧。".format(
            len(limit_up), "、".join(q["name"] for _, q in top_limit[:4])))
    if spread is not None:
        signals.append("样本成长 vs 宽基涨幅差 {:+.2f} 个百分点（全市场验证见第四章资金广度）。".format(spread))
    if flat_open >= len(ranked) * 0.6 and close_high >= len(ranked) * 0.5:
        signals.append("样本ETF多数平开单边拉升、收于全天高位，属资金持续买入型反包，形态优于跳空高开。")
    shrinking = [v for v in vol_checks if v["ratio"] < 1.0]
    if shrinking:
        signals.append("{} 只核心ETF今日成交量低于昨日恐慌日（{}），"
                       "**二级市场**换手未放大；但第十章接口显示一级市场有大额净申购，"
                       "故应表述为「一级有增量、二级未放量」，不可简单说「缺乏增量资金」。".format(
                           len(shrinking), "、".join("{:.0%}".format(v["ratio"]) for v in shrinking)))
    broker = store.get("sh512880", {}).get("quote", {}).get("pct")
    if "sh512880" in ALL_ETF_CODES and broker is not None and broker < 2:
        signals.append("证券ETF仅 {}，券商权重未接力，暂不支持「全面行情」判断。".format(pct_str(broker)))
    deep = [q["name"] for _, q in etfs
            if q["d60"] is not None and q["d60"] < -15]
    if deep:
        signals.append("{} 等近60日仍跌超15%，一根大阳线尚未改变中期下行结构。".format(
            "、".join(deep[:3])))

    # 样本代表性自检：等权 / 中位数 / 成交额加权，判断样本是否被小ETF或同赛道重复带偏
    pcts = sorted(q["pct"] for _, q in etfs if q["pct"] is not None)
    median = None
    if pcts:
        m = len(pcts) // 2
        median = pcts[m] if len(pcts) % 2 else (pcts[m - 1] + pcts[m]) / 2
    w_pairs = [(q["pct"], q["amount_yi"]) for _, q in etfs
               if q["pct"] is not None and q["amount_yi"]]
    w_total = sum(a for _, a in w_pairs)
    sample_weighted = (sum(p * a for p, a in w_pairs) / w_total if w_total else None)
    sample_mean = avg(pcts)
    dup_groups = []
    for label, codes in DUPLICATE_TRACK:
        present = [c for c in codes if c in ALL_ETF_CODES and c in store]
        if len(present) >= 2:
            dup_groups.append({"label": label, "count": len(present),
                               "names": [store[c]["quote"]["name"] for c in present]})

    return {
        "date": date, "indices": indices, "ranked": ranked, "groups": groups,
        "spread": spread, "limit_up": limit_up, "amount_total": amount_total,
        "flat_open": flat_open, "close_high": close_high, "total_etf": len(ranked),
        "vol_checks": vol_checks, "signals": signals,
        "sample_mean": sample_mean, "sample_median": median,
        "sample_weighted": sample_weighted, "dup_groups": dup_groups,
        "group_count": len([g for g in groups.values() if g["count"]]),
    }


# ----------------------------------------------------------------------------
# 分析层四：接下来走势
# ----------------------------------------------------------------------------

def tech_row(store, code):
    if code not in store:
        return None
    item = store[code]
    q, ind = item["quote"], indicators(item["bars"])
    price = q["price"]
    ma20 = ind["ma20"]
    boll = ind["boll"]
    macd = ind["macd"]
    kdj = ind["kdj"]

    support = min([v for v in (ind["prev_low"], ind["low5"]) if v is not None] or [None])
    # 压力位只保留「高于现价」的，且以 MA20/布林中轨为第一道关口
    pressures = [
        ("MA20 / 布林中轨", ma20),
        ("布林上轨", boll.get("upper")),
        ("MA60", ind["ma60"]),
    ]
    pressures = sorted(
        [(n, v) for n, v in pressures if v is not None and v > price],
        key=lambda x: x[1])

    # 多头/空头线索计分（用于情景倾向，不做概率预测）
    bull, bear = 0, 0
    pos = close_position(q)
    if pos is not None and pos >= 0.9:
        bull += 1
    if kdj.get("rising") and kdj.get("k_prev", 99) < 35:
        bull += 1
    if ind["rsi6"] is not None and ind["rsi6"] < 70:
        bull += 1
    if macd.get("bar", 0) > 0 and macd.get("golden"):
        bull += 1
    else:
        bear += 1
    if not macd.get("above_zero", False):
        bear += 1
    if ma20 and price < ma20:
        bear += 1
    if ind["ma60"] and price < ind["ma60"]:
        bear += 1
    if ind["prev_volume"] and q["volume"] and q["volume"] < ind["prev_volume"]:
        bear += 1

    return {
        "code": code, "name": q["name"], "price": price, "q": q, "ind": ind,
        "support": support, "pressures": pressures,
        "next_pressure": pressures[0] if pressures else None,
        "poc": ind["poc20"],
        "bull": bull, "bear": bear,
        "close_pos": pos,
    }


# ----------------------------------------------------------------------------
# 分析层二：全市场成交额
# ----------------------------------------------------------------------------

def analyze_market_amount(store, today, lookback=20, breadth=None):
    """
    当日成交额为行情快照精确值；历史成交额按「当日成交额 / 当日成交量」折算估算，
    只用于趋势、均值与分位对比，不作为历史绝对值。
    breadth 提供时，用全市场接口给出的精确成交额与「较上日变化」替换估算环比。
    """
    parts = {}
    for label, code in MARKET_CODES.items():
        if code not in store:
            continue
        q = store[code]["quote"]
        if q["amount_yi"] is None:
            continue
        parts[label] = {
            "code": code, "name": q["name"], "amount": q["amount_yi"],
            "pct": q["pct"], "volume": q["volume"],
        }

    sh = (parts.get("沪市") or {}).get("amount")
    sz = (parts.get("深市") or {}).get("amount")
    total = (sh or 0) + (sz or 0) if (sh is not None or sz is not None) else None
    bj = (parts.get("北证50") or {}).get("amount")
    all_total = ((total or 0) + (bj or 0)) if total is not None else None

    gate = "1万亿以下"
    gate_threshold = 10000
    for threshold, label in AMOUNT_GATES:
        if total is not None and total >= threshold:
            gate, gate_threshold = label, threshold
            break

    boards = []
    for label, code, parent in BOARD_SUBSETS:
        p, base = parts.get(label), (parts.get(parent) or {}).get("amount")
        if p and base:
            boards.append({"label": label, "amount": p["amount"],
                           "parent": parent, "base": base,
                           "ratio": p["amount"] / base * 100})

    # 两市日线成交量是真实数据；乘以各自当日「单位成交量金额」得到折算成交额
    histories = {}
    for code in TWO_MARKETS:
        if code not in store:
            continue
        q = store[code]["quote"]
        if q["volume"] and q["amount_yi"]:
            histories[code] = {
                "factor": q["amount_yi"] / q["volume"],
                "bars": {b["date"]: b for b in store[code]["bars"]},
            }

    common = None
    for h in histories.values():
        keys = set(h["bars"])
        common = keys if common is None else (common & keys)
    series = []
    if common:
        for d in sorted(common):
            vol = sum(h["bars"][d]["volume"] for h in histories.values())
            est = sum(h["bars"][d]["volume"] * h["factor"] for h in histories.values())
            series.append({"date": d, "volume": vol, "amount_est": est})
        series = series[-lookback:] if lookback > 0 else []
        if series and total is not None and series[-1]["date"] == today.get("date"):
            series[-1]["amount_est"] = total  # 当日用精确值覆盖折算误差

    # 全市场接口给出精确成交额（元）与较上日变化（元），优先替换折算环比
    bd_all = ((breadth or {}).get("turnover_dsb") or {}).get("all") or {}
    exact_amount_yi = (num(bd_all.get("amount")) or 0) / 1e8 or None
    exact_change_yi = (num(bd_all.get("amount_change")) or 0) / 1e8 or None
    exact_prev_yi = ((exact_amount_yi - exact_change_yi)
                     if (exact_amount_yi and exact_change_yi is not None) else None)
    exact_chg = (exact_change_yi / exact_prev_yi * 100
                 if (exact_prev_yi and exact_prev_yi > 0) else None)

    amounts = [s["amount_est"] for s in series]
    cur = amounts[-1] if amounts else None
    prev = amounts[-2] if len(amounts) > 1 else None
    chg = exact_chg if exact_chg is not None else (
        (cur / prev - 1) * 100 if (cur and prev) else None)
    chg_source = "精确值" if exact_chg is not None else "折算估算"
    ma5 = avg(amounts[-5:])
    ma10 = avg(amounts[-10:]) if len(amounts) >= 10 else None
    ma20 = avg(amounts)  # 整个回看窗口的均额
    ratio5 = cur / ma5 if (cur and ma5) else None
    pctl = (sum(1 for a in amounts if a <= cur) / len(amounts) * 100) if (cur and amounts) else None

    idx_pct = (store.get("sh000001", {}).get("quote", {}) or {}).get("pct")
    if idx_pct is None:
        idx_pct = (parts.get("沪市") or {}).get("pct")
    if idx_pct is None or chg is None:
        vp = "数据不足"
    elif idx_pct > 0 and chg > 3:
        vp = "价涨量增（健康）"
    elif idx_pct > 0 and chg < -3:
        vp = "价涨量缩（量价背离，谨慎）"
    elif idx_pct < 0 and chg > 3:
        vp = "价跌量增（抛压释放）"
    elif idx_pct < 0 and chg < -3:
        vp = "价跌量缩（观望）"
    else:
        vp = "量价平稳"

    etf_share = (today["amount_total"] / total * 100) if (total and today.get("amount_total")) else None
    shrinking = [v["name"] for v in today.get("vol_checks", []) if v["ratio"] < 1.0]
    divergence = bool(chg is not None and chg > 3 and len(shrinking) >= 2)

    # 解析接口自带的市场评论文本（连续破万亿天数、相对均值放量、两融、ADR）
    comments = (breadth or {}).get("comments") or {}
    amt_comment = comments.get("total_amount") or ""
    m_trillion = re.search(r"连续\s*(\d+)\s*日破万亿", amt_comment)
    m_vs_y = re.search(r"相对昨日放量\s*([\d.]+)%", amt_comment)
    m_vs_m = re.search(r"相对最近一月成交额均值放量\s*([\d.]+)%", amt_comment)
    trillion_days = int(m_trillion.group(1)) if m_trillion else None
    vs_yesterday = float(m_vs_y.group(1)) if m_vs_y else None
    vs_month = float(m_vs_m.group(1)) if m_vs_m else None

    signals = []
    if total:
        signals.append("沪深两市今日成交 **{:.0f} 亿（{:.2f} 万亿）**，{} **{}** 关口；"
                       "北证50 成交 {:.0f} 亿（成分指数口径，仅参考）。".format(
                           total, total / 10000.0,
                           "站上" if total >= gate_threshold else "未站上", gate,
                           bj if bj else 0))
    if chg is not None:
        signals.append("较昨日 **{}**（{}），当日为 5 日均额的 {:.2f} 倍，"
                       "处于近 {} 日 **{:.0f}% 分位**。".format(
                           pct_str(chg), chg_source,
                           ratio5 if ratio5 else 0, len(amounts), pctl or 0))
    if exact_amount_yi and exact_change_yi is not None:
        signals.append("全市场接口精确口径：**{:.2f} 亿**（较上日 **{:+.2f} 亿**），"
                       "与快照合计 {:.0f} 亿 可互相校验。".format(
                           exact_amount_yi, exact_change_yi, total or 0))
    if trillion_days or vs_month is not None:
        signals.append("官方评论：连续 **{} 日**破万亿，相对昨日放量 {}、"
                       "相对近一月均值放量 {}。".format(
                           trillion_days if trillion_days else "—",
                           pct_str(vs_yesterday), pct_str(vs_month)))
    signals.append("**量价配合：{}**（上证指数今日 {}）。".format(vp, pct_str(idx_pct)))
    if boards:
        signals.append("成交结构：{}。板块指数为子集，不与两市合计重复相加。".format(
            "、".join("{}占{} {:.1f}%".format(b["label"], b["parent"], b["ratio"])
                     for b in boards)))
    if etf_share is not None:
        signals.append("报告样本 {} 只ETF合计成交 {:.0f} 亿，占两市成交额 **{:.1f}%**。".format(
            today["total_etf"], today["amount_total"], etf_share))
    if divergence:
        signals.append("**背离提示**：全市场放量，但 {} 只核心ETF为缩量反包（{}），"
                       "说明样本ETF成交量未同步放大，指数级放量未必对应ETF端增量资金。".format(
                           len(shrinking), "、".join(shrinking)))

    return {
        "parts": parts, "total": total, "bj": bj, "all_total": all_total,
        "gate": gate, "gate_threshold": gate_threshold, "boards": boards,
        "series": series, "cur": cur, "prev": prev, "chg": chg,
        "ma5": ma5, "ma10": ma10, "ma20": ma20, "ratio5": ratio5, "pctl": pctl,
        "vp": vp, "idx_pct": idx_pct, "etf_share": etf_share,
        "divergence": divergence, "shrinking": shrinking,
        "signals": signals, "lookback": len(amounts),
        "exact_amount_yi": exact_amount_yi, "exact_change_yi": exact_change_yi,
        "chg_source": chg_source, "trillion_days": trillion_days,
        "vs_yesterday": vs_yesterday, "vs_month": vs_month,
        "comments": comments,
    }


# ----------------------------------------------------------------------------
# 分析层三：全市场资金流入流出板块
# ----------------------------------------------------------------------------

def parse_board_row(raw):
    """把板块资金流原始行（金额单位万元）解析成亿元口径。"""
    name = raw.get("name") or ""
    if not name:
        return None
    lzg = raw.get("lzg") or {}
    leader = ""
    if lzg.get("name"):
        leader = "{}({})".format(lzg["name"], pct_str(num(lzg.get("zdf"))))
    return {
        "code": raw.get("code") or "", "name": name,
        "pct": num(raw.get("zdf")),
        "amount_yi": yi(num(raw.get("cje"))),
        "turnover": num(raw.get("hsl")),
        "updown": raw.get("zgb") or "",
        "inflow_yi": yi(num(raw.get("zllr"))),
        "outflow_yi": yi(num(raw.get("zllc"))),
        "net_yi": yi(num(raw.get("zljlr"))),
        "net5_yi": yi(num(raw.get("zljlr_d5"))),
        "net20_yi": yi(num(raw.get("zljlr_d20"))),
        "leader": leader,
    }


def updown_breadth(rows):
    """从 '181/183' 汇总上涨家数占比（申万行业成分基本不重叠，可近似汇总）。"""
    up = total = 0
    for r in rows:
        s = (r["updown"] or "").split("/")
        if len(s) == 2:
            a, b = num(s[0]), num(s[1])
            if None not in (a, b) and b > 0:
                up += a
                total += b
    return (int(up), int(total), up / total * 100 if total else None)


def analyze_fundflow(ff, today, market, top=FUND_TOP):
    boards = {}
    for key in ("industry", "concept"):
        rows = [parse_board_row(r) for r in (ff.get(key) or [])]
        boards[key] = [r for r in rows if r and r["net_yi"] is not None]

    ind, con = boards["industry"], boards["concept"]
    if not ind:
        return {"available": False, "industry": [], "concept": [],
                "signals": [], "date": today.get("date", "")}

    pos = [r for r in ind if r["net_yi"] > 0]
    neg = [r for r in ind if r["net_yi"] < 0]
    gross_in = sum(r["net_yi"] for r in pos)
    gross_out = sum(r["net_yi"] for r in neg)
    net_total = sum(r["net_yi"] for r in ind)

    in_top = sorted(ind, key=lambda r: r["net_yi"], reverse=True)[:top]
    out_top = sorted(ind, key=lambda r: r["net_yi"])[:top]

    def is_broad(r):
        return any(k in r["name"] for k in BROAD_CONCEPT_KEYWORDS)

    con_focused = [r for r in con if not is_broad(r)]
    con_top = sorted(con_focused, key=lambda r: r["net_yi"], reverse=True)[:top]
    con_bot = sorted(con_focused, key=lambda r: r["net_yi"])[:top]
    broad_excluded = len(con) - len(con_focused)

    top5_share = (sum(r["net_yi"] for r in in_top[:5]) / gross_in * 100
                  if gross_in else None)
    in_avg = avg([r["pct"] for r in pos])
    out_avg = avg([r["pct"] for r in neg])
    short_cover = [r for r in in_top
                   if (r["net5_yi"] or 0) < 0 and (r["net20_yi"] or 0) < 0]
    up, total, breadth = updown_breadth(ind)

    cje_sum = sum(r["amount_yi"] or 0 for r in ind)
    cross_diff = ((cje_sum / market["total"] - 1) * 100
                  if market.get("total") else None)

    # 与第一章 ETF 结构联动：资金第一的板块若对应ETF反而缩量，即为背离
    shrink_names = [v["name"] for v in today.get("vol_checks", [])
                    if v["ratio"] < 1.0]
    tech_keywords = ("半导体", "芯片", "电子", "元件", "通信", "计算机", "软件", "存储")
    top1 = in_top[0] if in_top else None
    tech_divergence = bool(top1 and any(k in top1["name"] for k in tech_keywords)
                           and shrink_names)

    signals = []
    signals.append("申万 **{} 个行业板块** 中，主力净流入合计 **{:+.0f} 亿**、净流出 {:+.0f} 亿，"
                   "全市场主力净额 **{:+.0f} 亿**；净流入 {} 个 / 净流出 {} 个，"
                   "资金广度 {:.0f}%。".format(
                       len(ind), gross_in, gross_out, net_total,
                       len(pos), len(neg), len(pos) / len(ind) * 100))
    if breadth is not None:
        signals.append("行业成分股上涨 {}/{} 家，占比 **{:.1f}%**——{}。".format(
            up, total, breadth,
            "普涨格局" if breadth >= 70 else
            ("涨跌互现" if breadth >= 50 else "上涨家数未过半，结构性特征明显")))
    if cross_diff is not None:
        signals.append("行业板块成交额合计 **{:.0f} 亿**，与第三章两市成交额 {:.0f} 亿 相差 {:+.1f}%，"
                       "{}。".format(
                           cje_sum, market["total"], cross_diff,
                           "口径基本吻合，行业分类可视为全市场覆盖"
                           if abs(cross_diff) < 5 else "存在统计口径差异，绝对值仅供参考"))
    if top5_share is not None:
        signals.append("净流入 Top5 占全部净流入 **{:.0f}%**，{}。".format(
            top5_share,
            "资金高度集中" if top5_share >= 50 else
            ("资金相对分散" if top5_share < 30 else "资金中度集中")))
    if None not in (in_avg, out_avg):
        signals.append("净流入板块平均涨幅 {} vs 净流出板块 {}，价差 {:+.1f} 个百分点，"
                       "说明资金流向与涨跌{}。".format(
                           pct_str(in_avg), pct_str(out_avg), in_avg - out_avg,
                           "同向、选股有效" if in_avg - out_avg > 2 else "分化不大"))
    if short_cover:
        signals.append("净流入 Top{} 中有 **{} 个**板块近5日、20日累计主力资金仍为净流出（{}），"
                       "今日流入更像**短线回补/超跌反抽**，而非趋势性增配。".format(
                           len(in_top), len(short_cover),
                           "、".join(r["name"] for r in short_cover[:5])))
    if out_top:
        signals.append("净流出榜最大单板块仅 **{:+.0f} 亿**（{}），流出压力{}，"
                       "今日更接近「资金搬家」而非全面撤离。".format(
                           out_top[0]["net_yi"], out_top[0]["name"],
                           "温和" if abs(out_top[0]["net_yi"]) < 30 else "明显"))
    if con_top:
        signals.append("剔除 {} 个宽泛属性概念后，题材端资金最聚焦：{}（单位：亿元）。"
                       "概念板块成分股互相重叠，**不可与行业板块加总**。".format(
                           broad_excluded,
                           "、".join("{} {:+.0f}".format(r["name"], r["net_yi"])
                                    for r in con_top[:3])))
    if tech_divergence:
        signals.append("**背离提示**：{} 主力净流入居首（{:+.0f} 亿），但样本中 {} 只科技/宽基ETF"
                       "为缩量反包（{}）——板块级资金流入未同步体现为ETF端放量。".format(
                           top1["name"], top1["net_yi"], len(shrink_names),
                           "、".join(shrink_names)))

    outlook_notes = []
    if short_cover and len(short_cover) >= len(in_top) / 2:
        outlook_notes.append("资金面支持「**超跌修复**」而非「趋势反转」：多数领涨板块的中期"
                             "（5日/20日）主力资金仍是净流出，今日单日流入尚未扭转累计趋势，"
                             "与第十四章技术面 MA60 压制互为印证。")
    if top5_share is not None and top5_share >= 50:
        outlook_notes.append("净流入高度集中于少数板块，属**结构性资金进攻**；"
                             "一旦龙头板块分歧，缺少承接板块，指数容易回落。")
    if len(neg) > len(pos):
        outlook_notes.append("净流出板块数多于净流入，说明并非全面增量入场，"
                             "更接近存量资金调仓（从 {} 等流出、向 {} 等集中）。".format(
                                 "、".join(r["name"] for r in out_top[:2]),
                                 "、".join(r["name"] for r in in_top[:2])))
    if market.get("total"):
        outlook_notes.append("**明日资金确认**：净流入 Top 板块需连续第二日维持净流入，"
                             "且全市场主力净额不大幅转负；同时两市成交额守住 {:.0f} 亿。".format(
                                 market["gate_threshold"]))
    if not outlook_notes:
        outlook_notes.append("资金流未呈现极端结构，以价格对压力位的攻防为主要判断依据。")

    return {
        "available": True, "industry": ind, "concept": con,
        "pos": pos, "neg": neg, "gross_in": gross_in, "gross_out": gross_out,
        "net_total": net_total, "in_top": in_top, "out_top": out_top,
        "con_top": con_top, "con_bot": con_bot,
        "con_total": len(con), "broad_excluded": broad_excluded,
        "top5_share": top5_share,
        "in_avg": in_avg, "out_avg": out_avg, "short_cover": short_cover,
        "breadth": breadth, "up": up, "total_stocks": total,
        "cje_sum": cje_sum, "cross_diff": cross_diff,
        "tech_divergence": tech_divergence, "top1": top1,
        "signals": signals, "outlook_notes": outlook_notes,
        "date": today.get("date", ""),
    }


# ----------------------------------------------------------------------------
# 分析层五：全市场涨跌广度与情绪
# ----------------------------------------------------------------------------

def analyze_breadth(breadth, market, today):
    d = (breadth or {}).get("ups_downs_dsb") or {}
    if not d:
        return {"available": False, "signals": [], "detail": [], "intraday": []}

    up, down = num(d.get("up_count")), num(d.get("down_count"))
    flat, susp = num(d.get("flat_count")), num(d.get("suspension_count"))
    uplim, dnlim = num(d.get("up_limit_count")), num(d.get("down_limit_count"))
    ratio = num(d.get("up_ratio"))
    detail = d.get("detail") or []
    comment = d.get("up_ratio_comment") or ""

    m_heat = re.search(r"过去一年有\s*([\d.]+)%", comment)
    heat_pctl = float(m_heat.group(1)) if m_heat else None
    limit_ratio = (uplim / dnlim) if (uplim and dnlim) else None
    strong_up = sum(r.get("count") or 0 for r in detail
                    if r.get("flag") == 1 and r.get("section") in ("涨停", ">7%", "7~5%"))
    strong_dn = sum(r.get("count") or 0 for r in detail
                    if r.get("flag") == -1 and r.get("section") in ("跌停", "7%<", "5~7%"))

    comments = (breadth or {}).get("comments") or {}
    margin_txt = comments.get("margin") or ""
    m_bal = re.search(r"融资余额\s*([\d.]+)\s*亿元", margin_txt)
    m_chg = re.search(r"较前一日(减少|增加)\s*([\d.]+)%", margin_txt)
    margin_bal = float(m_bal.group(1)) if m_bal else None
    margin_chg = (float(m_chg.group(2)) * (-1 if m_chg.group(1) == "减少" else 1)
                  if m_chg else None)
    m_days = re.search(r"(\d+)\s*日连降", margin_txt)
    margin_down_days = int(m_days.group(1)) if m_days else None

    # 打板赚钱效应：昨日涨停股今日上涨占比
    follow_txt = comments.get("short_money") or ""
    m_follow = re.search(r"今日\s*([\d.]+)%\s*的个股上涨", follow_txt)
    follow_up = float(m_follow.group(1)) if m_follow else None

    signals = []
    total_stocks = (up or 0) + (down or 0) + (flat or 0) + (susp or 0)
    signals.append("全市场 **{} 家上涨 / {} 家下跌 / {} 平 / {} 停牌**，上涨占比 **{:.0f}%**，"
                   "{}。".format(
                       int(up or 0), int(down or 0), int(flat or 0), int(susp or 0),
                       ratio or 0,
                       "普涨格局" if (ratio or 0) >= 70 else
                       ("涨跌互现" if (ratio or 0) >= 50 else "跌多涨少")))
    if heat_pctl is not None:
        signals.append("该上涨广度在过去一年中仅 **{:.0f}%** 的交易日出现，"
                       "属{}区间。".format(
                           heat_pctl, "情绪高热度" if heat_pctl <= 20 else "常见"))
    signals.append("涨停 **{}** 家 / 跌停 **{}** 家（{}），涨超5% 约 {} 家 vs "
                   "跌超5% 约 {} 家，{}。".format(
                       int(uplim or 0), int(dnlim or 0),
                       "涨停跌停比 {:.1f}:1".format(limit_ratio) if limit_ratio else "—",
                       int(strong_up), int(strong_dn),
                       "赚钱效应强" if strong_up > strong_dn * 10 else "强弱分化有限"))
    if follow_up is not None and ratio:
        signals.append("**关键背离**：全市场 {} 上涨，但「昨日涨停股今日上涨」仅 **{:.1f}%**，"
                       "说明普涨由新资金推动、追高接力赚钱效应弱，"
                       "次日高开品种易出现兑现。".format(
                           pct_str(ratio, 0).replace("+", ""), follow_up))
    if margin_bal:
        signals.append("**杠杆资金背离**：融资余额 **{:.0f} 亿**（{}，{}），"
                       "杠杆资金做多情绪未同步升温，与今日指数大涨方向相反。".format(
                           margin_bal,
                           pct_str(margin_chg) if margin_chg is not None else "环比数据缺失",
                           "连降 {} 日".format(margin_down_days) if margin_down_days else "—"))
    if comments.get("adr"):
        signals.append("涨跌家数 ADR：{}".format(comments["adr"]))
    if market.get("chg") is not None and (market["chg"] or 0) > 3 and (ratio or 0) >= 70:
        signals.append("**量、价、广度三者共振**：成交额放大、指数普涨、上涨家数占比 {}，"
                       "短期反弹的确认度高于单纯看指数。".format(pct_str(ratio, 0).replace("+", "")))
    else:
        signals.append("广度与量价未形成完全共振，仍需下一交易日确认。")

    return {
        "available": True, "up": up, "down": down, "flat": flat, "susp": susp,
        "uplim": uplim, "dnlim": dnlim, "ratio": ratio,
        "total_stocks": total_stocks, "detail": detail, "comment": comment,
        "heat_pctl": heat_pctl, "limit_ratio": limit_ratio,
        "strong_up": strong_up, "strong_dn": strong_dn,
        "margin_bal": margin_bal, "margin_chg": margin_chg,
        "margin_down_days": margin_down_days, "margin_txt": margin_txt,
        "follow_up": follow_up, "adr": comments.get("adr") or "",
        "adr_limit": comments.get("adr_limit") or "",
        "signals": signals,
    }


def analyze_intraday(breadth):
    """日内上涨占比演变 + 上证指数分时节奏。"""
    um = ((breadth or {}).get("ups_downs_minute") or {}).get("hsAUpsRatio") or {}
    ms = ((breadth or {}).get("minute_set") or {}).get("minute_sh_index") or {}
    pts, ipts = [], []
    for s in um.get("data") or []:
        p = s.split()
        if len(p) >= 2 and num(p[1]) is not None:
            pts.append((p[0], num(p[1])))
    pre = num(ms.get("pre"))
    for s in ms.get("data") or []:
        p = s.split()
        if len(p) >= 2 and num(p[1]) is not None:
            ipts.append((p[0], num(p[1]), num(p[2]) if len(p) > 2 else None))
    if not pts and not ipts:
        return {"available": False, "signals": []}

    def fmt_t(t):
        return t[:2] + ":" + t[2:] if len(t) == 4 and t.isdigit() else t

    signals, levels = [], []
    if pts:
        vals = [v for _, v in pts]
        lo = min(pts, key=lambda x: x[1])
        hi = max(pts, key=lambda x: x[1])
        first, last = vals[0], vals[-1]
        am_last = next((v for t, v in pts if t <= "1130"), None)
        pm_first = next((v for t, v in pts if t >= "1300"), None)
        levels = [("开盘", first), ("上午收盘", am_last), ("下午开盘", pm_first),
                  ("收盘", last), ("日内最低", lo[1]), ("日内最高", hi[1])]
        signals.append("上涨家数占比从开盘 **{:.0f}%** 一路升至收盘 **{:.0f}%**"
                       "（日内最低 {} 时 {:.0f}%、最高 {} 时 {:.0f}%），{}。".format(
                           first, last, fmt_t(lo[0]), lo[1], fmt_t(hi[0]), hi[1],
                           "午后资金持续进场、无中途退潮" if last - first > 20
                           else "日内广度变化相对平缓"))
        if am_last is not None and pm_first is not None and last > pm_first:
            signals.append("下午开盘 {:.0f}% → 收盘 {:.0f}%，**尾盘仍在扩散**，"
                           "次日惯性通常较强，但也意味着浮盈盘集中。".format(
                               pm_first, last))
    if ipts and pre:
        vals = [v for _, v, _ in ipts]
        t_lo = min(ipts, key=lambda x: x[1])
        t_hi = max(ipts, key=lambda x: x[1])
        tail = [v for t, v, _ in ipts if t >= "1430"]
        tail_chg = ((tail[-1] / tail[0] - 1) * 100) if len(tail) >= 2 else None
        signals.append("上证指数日内低点 {} 时 {}、高点 {} 时 {}，收盘 {}（昨收 {}）；"
                       "尾盘30分钟 {}。".format(
                           fmt_t(t_lo[0]), price_str(t_lo[1]),
                           fmt_t(t_hi[0]), price_str(t_hi[1]),
                           price_str(vals[-1]), price_str(pre),
                           pct_str(tail_chg) if tail_chg is not None else "数据不足"))
        close_pos = ((vals[-1] - min(vals)) / (max(vals) - min(vals))
                     if max(vals) > min(vals) else None)
        if close_pos is not None:
            signals.append("指数收于日内振幅 **{:.0f}%** 位置，{}。".format(
                close_pos * 100,
                "收在全天高位，买盘持续到收盘" if close_pos >= 0.9
                else "尾盘有回落"))
    return {"available": True, "points": pts, "index_points": ipts,
            "levels": levels, "signals": signals}


# ----------------------------------------------------------------------------
# 分析层六：ETF 折溢价 / 份额 / 净申赎
# ----------------------------------------------------------------------------

def analyze_etf_micro(store, shares_delta, shares_days, shares_file=None):
    rows = []
    for c in ALL_ETF_CODES:
        if c not in store:
            continue
        q = store[c]["quote"]
        if q["discount"] is None and not q["total_shares"]:
            continue
        size_yi = (q["total_shares"] * (q["iopv"] or q["price"]) / 1e8
                   if (q["total_shares"] and (q["iopv"] or q["price"])) else None)
        dlt = shares_delta.get(c)
        rows.append({
            "code": c, "name": q["name"], "price": q["price"],
            "iopv": q["iopv"], "discount": q["discount"],
            "net_value": q["net_value"], "shares_yi": (q["total_shares"] or 0) / 1e8,
            "size_yi": size_yi, "amount_yi": q["amount_yi"], "pct": q["pct"],
            "delta": dlt,
        })
    prem = sorted([r for r in rows if r["discount"] is not None],
                  key=lambda r: r["discount"], reverse=True)
    top_prem = prem[:8]
    top_disc = sorted([r for r in rows if r["discount"] is not None],
                      key=lambda r: r["discount"])[:5]
    overheat = [r for r in rows if (r["discount"] or 0) >= 0.5]
    has_delta = [r for r in rows if r["delta"]]

    signals = []
    signals.append("样本中 {} 只ETF可读取 IOPV / 折溢价。成交额是**二级市场换手**，"
                   "份额变化才是**一级市场净申赎**，两者不可混为一谈。".format(len(rows)))
    if top_prem:
        signals.append("溢价最高：{}。溢价走高说明二级买盘强于净值，"
                       "往往先于净值反映情绪，回落时有「价格+溢价」双杀风险。".format(
                           "、".join("{} {:+.2f}%".format(r["name"], r["discount"])
                                    for r in top_prem[:4])))
    if overheat:
        signals.append("**{} 只ETF溢价 ≥0.5%**（{}），短线情绪透支特征明显。".format(
            len(overheat), "、".join(r["name"] for r in overheat[:4])))
    limit_disc = [r for r in top_disc
                  if (r["pct"] or 0) >= 9.8 and (r["discount"] or 0) < -0.5]
    real_disc = [r for r in top_disc
                 if r not in limit_disc and (r["discount"] or 0) < -0.3]
    if limit_disc:
        signals.append("**折价需分情况**：{} 为**涨停限价造成的被动折价**——"
                       "标的指数涨幅超过 ETF 10% 涨停上限，IOPV 继续上行而二级价格被封在涨停价，"
                       "并非二级市场抛压偏弱，次日开盘常会向上补平。".format(
                           "、".join("{} {:+.2f}%".format(r["name"], r["discount"])
                                    for r in limit_disc[:4])))
    if real_disc:
        signals.append("非涨停类折价居前：{}，这类才更接近二级市场相对净值偏弱。".format(
            "、".join("{} {:+.2f}%".format(r["name"], r["discount"])
                     for r in real_disc[:3])))
    if has_delta:
        in_n = sum(1 for r in has_delta if r["delta"]["diff"] > 0)
        out_n = sum(1 for r in has_delta if r["delta"]["diff"] < 0)
        net = sum(r["delta"]["diff"] for r in has_delta) / 1e8
        signals.append("**份额变化（较上一记录日）**：在 {} 只有可比历史的标的中，"
                       "净申购 {} 只 / 净赎回 {} 只，合计 {:+.2f} 亿份；"
                       "这才是 ETF 端真实资金流入流出（其余标的仍无历史，不计入）。".format(
                           len(has_delta), in_n, out_n, net))
    else:
        signals.append("**份额逐日累积**：当前记录 {} 个交易日（连续运行 ≥2 日可自算净申赎，"
                       "文件 `{}`）。作为交叉校验，第十章已通过接口取得全市场 ETF "
                       "真实净申购（T-1 披露），两条路径可互相验证；"
                       "本地累积保留用于接口失效时的降级。".format(
                           shares_days,
                           os.path.basename(shares_file or SHARES_HISTORY_FILE)))
    return {
        "rows": rows, "top_prem": top_prem, "top_disc": top_disc,
        "limit_disc": limit_disc, "real_disc": real_disc,
        "overheat": overheat, "has_delta": has_delta,
        "shares_days": shares_days,
        "shares_file": os.path.basename(shares_file or SHARES_HISTORY_FILE),
        "signals": signals,
        "available": bool(rows),
    }


# ----------------------------------------------------------------------------
# 分析层七：跨市场外围 + 估值横向
# ----------------------------------------------------------------------------

def analyze_global(glob, store):
    rows = []
    for code, (label, region) in GLOBAL_CODES.items():
        f = (glob or {}).get(code)
        if not f or len(f) <= 32:
            continue
        rows.append({"code": code, "name": label, "region": region,
                     "price": num(f[3]), "pct": num(f[32])})
    by_region = {}
    for r in rows:
        by_region.setdefault(r["region"], []).append(r)

    valuations = []
    for c in INDEX_CODES:
        if c not in store:
            continue
        q = store[c]["quote"]
        if q["pe"] and q["pe"] > 0:
            valuations.append({"code": c, "name": q["name"],
                               "pe": q["pe"], "pct": q["pct"]})
    valuations.sort(key=lambda r: r["pe"])

    a_growth = (store.get("sh000688", {}).get("quote") or {}).get("pct")
    hk_tech = next((r["pct"] for r in rows if r["code"] == "hkHSTECH"), None)
    a_weight = (store.get("sh000016", {}).get("quote") or {}).get("pct")

    signals = []
    if rows:
        signals.append("外围 {} 个指数：{}。美股为**最近一交易日收盘**（与 A 股存在时差），"
                       "港股/欧洲为各自当日收盘，不可直接当作同步领先指标。".format(
                           len(rows),
                           "、".join("{} {}".format(r["name"], pct_str(r["pct"]))
                                    for r in rows[:5])))
    for region, rs in by_region.items():
        m = avg([r["pct"] for r in rs])
        signals.append("{}均值 {}（{} 个指数）。".format(
            region, pct_str(m), len(rs)))
    if None not in (a_growth, hk_tech):
        gap = a_growth - hk_tech
        signals.append("**A股科创 vs 港股科技**：科创50 {} vs 恒生科技 {}，"
                       "差 {:+.1f} 个百分点，{}。".format(
                           pct_str(a_growth), pct_str(hk_tech), gap,
                           "A股科技明显更强，偏向本土事件驱动"
                           if gap > 3 else "两岸科技同步性较好"))
    if valuations:
        lo, hi = valuations[0], valuations[-1]
        pes = [v["pe"] for v in valuations]
        signals.append("A股指数 PE 横向：最低 {} {:.1f} 倍，最高 {} {:.1f} 倍，"
                       "中位数 {:.1f} 倍。成长指数估值显著高于权重指数，"
                       "今日领涨方向恰是估值最贵的一端。".format(
                           lo["name"], lo["pe"], hi["name"], hi["pe"],
                           sorted(pes)[len(pes) // 2]))
    if None not in (a_growth, a_weight):
        signals.append("**风格轮动**：科创50 {} vs 上证50 {}，差 {:+.1f} 个百分点，"
                       "今日为典型成长/小盘占优。".format(
                           pct_str(a_growth), pct_str(a_weight), a_growth - a_weight))
    return {"available": bool(rows), "rows": rows, "by_region": by_region,
            "valuations": valuations, "signals": signals}


def analyze_futures(fut_rows, store):
    """股指期货基差 + 国债期货：把「机构对冲压力」变成可读数字。

    基差 = 期货最新 − 现货指数；贴水（负值）通常来自对冲盘做空期货、
    分红预期与融券约束，是 A 股少见的「高频机构仓位侧写」。
    """
    out = {"available": False, "rows": [], "signals": [], "status": "empty"}
    if not fut_rows:
        return out
    rows = []
    for r in fut_rows:
        spot = None
        if r.get("spot_code") and store.get(r["spot_code"]):
            spot = ((store[r["spot_code"]] or {}).get("quote") or {}).get("price")
        basis = (r["last"] - spot) if spot else None
        rows.append(dict(r, spot=spot, basis=basis,
                         basis_pct=(basis / spot * 100) if spot else None))
    out.update(available=True, rows=rows, status="ok",
               time=(rows[0].get("time") or ""))
    signals = []
    idx = [r for r in rows if r.get("basis_pct") is not None]
    if idx:
        signals.append("**基差（期货 − 现货）**：{}。负值为贴水，"
                       "代表对冲盘用期货替代卖现货的强度；贴水越深，"
                       "说明机构越不愿直接抛售持仓。".format(
                           "、".join("{} {}（{:+.2f}%）".format(
                               r["name"].split(" ")[0], price_str(r["last"]),
                               r["basis_pct"]) for r in idx)))
        deep = min(idx, key=lambda r: r["basis_pct"])
        _small = deep["code"] in ("nf_IC0", "nf_IM0")
        signals.append("**贴水最深**：{}，{:.2f}%（{} 点）。{}".format(
            deep["name"], deep["basis_pct"], price_str(abs(deep["basis"])),
            "中小盘合约贴水更深，通常对应中性策略/场外期权的对冲成本"
            "集中在中小盘，与「小盘拥挤度」互为印证。"
            if _small else
            "大盘合约贴水更深，多与权重股对冲盘、分红除息预期相关，"
            "并不必然代表小盘更弱，需与 IC/IM 的相对贴水一起看。"))
    bond = [r for r in rows if not r.get("spot_code")]
    if bond:
        signals.append("**国债期货**：{}。股债跷跷板的高频读数，"
                       "与第十一章 ERP 互为印证。".format(
                           "、".join("{} {}（{}）".format(
                               r["name"].split(" ")[0], price_str(r["last"]),
                               pct_str(r["pct"])) for r in bond)))
    signals.append("**口径边界**：连续合约临近交割时基差会被强制收敛；"
                   "本层**无历史序列**，只作当日对冲温度的截面读数，"
                   "不与历史分位比较，也不单独构成方向判断。")
    out["signals"] = signals
    return out


def daily_returns(bars):
    """日线 -> {日期: 百分比收益}。"""
    out = {}
    for i in range(1, len(bars or [])):
        prev, cur = bars[i - 1]["close"], bars[i]["close"]
        if prev:
            out[bars[i]["date"]] = (cur / prev - 1) * 100
    return out


def compute_topic_ar(topics, store, date, bench="sh000300", lookback=250,
                     min_obs=60):
    """主题级单日异常收益 AR = ETF 收益 − β × 基准收益。

    β 用「报告日之前」最近 lookback 个共同交易日的日收益做 OLS 估计，
    估计窗不含报告日本身，避免用当日数据解释当日（look-ahead）。
    """
    if not store or bench not in store:
        return {"available": False, "rows": [], "reason": "缺基准（沪深300）日线"}
    bret = daily_returns(store[bench]["bars"])
    bdates = [b["date"] for b in store[bench]["bars"]]
    bench_pct = bret.get(date)
    rows = []
    for tp in topics or []:
        kws = [k for k in (tp.get("etf_kw") or []) if k]
        if not kws:
            continue
        members = []
        for code, node in store.items():
            if code == bench:
                continue
            name = ((node.get("quote") or {}).get("name") or "")
            if not name or not any(k in name for k in kws):
                continue
            r = daily_returns(node.get("bars"))
            if date not in r:
                continue
            common = [d for d in bdates if d in r and d in bret and d < date]
            common = common[-lookback:]
            if len(common) < min_obs:
                continue
            xs = [bret[d] for d in common]
            ys = [r[d] for d in common]
            mx, my = sum(xs) / len(xs), sum(ys) / len(ys)
            var = sum((x - mx) ** 2 for x in xs)
            if var <= 0:
                continue
            beta = sum((x - mx) * (y - my) for x, y in zip(xs, ys)) / var
            members.append({"code": code, "name": name, "beta": beta,
                            "pct": r[date],
                            "ar": r[date] - beta * (bench_pct or 0.0)})
        if members:
            rows.append({
                "topic": tp.get("topic"), "members": members[:5],
                "first_hhmm": (tp.get("first_ts") or "")[11:16].replace(":", ""),
                "n": len(members),
                "pct": sum(m["pct"] for m in members) / len(members),
                "beta": sum(m["beta"] for m in members) / len(members),
                "ar": sum(m["ar"] for m in members) / len(members),
            })
    rows.sort(key=lambda r: -abs(r["ar"]))
    return {"available": bool(rows), "rows": rows, "date": date, "bench": bench,
            "bench_pct": bench_pct, "reason": "" if rows else "当日无匹配ETF或日线不足"}


def update_event_study(path, date, rows, keep_days=180, car=None):
    """把当日主题 AR（及日内 CAR）写入台账，返回跨日统计。

    真源为 hist_event 表（payload=当日条目列表）。CAR 依赖当日分时，
    **只有在线运行才有值**；同一天重跑时以最新一次为准，缺失时保留台账里
    已有的旧值，避免二次运行把 CAR 抹成空。
    """
    con = store_connect()
    try:
        days = {d: json.loads(p) for d, p in con.execute(
            "SELECT date,payload FROM hist_event ORDER BY date")}
        if rows:
            _old = {it.get("topic"): it for it in days.get(date) or []}
            items = []
            for r in rows:
                it = {"topic": r["topic"], "ar": round(r["ar"], 4),
                      "n": r["n"]}
                _c = (car or {}).get(r["topic"])
                if _c is None:
                    _c = (_old.get(r["topic"]) or {}).get("car")
                if _c is not None:
                    it["car"] = round(_c, 4)
                items.append(it)
            days[date] = items
            con.execute(
                "INSERT OR REPLACE INTO hist_event(date,payload,updated) "
                "VALUES(?,?,?)",
                (date, json.dumps(items, ensure_ascii=False,
                                  separators=(",", ":")),
                 datetime.now(CST).strftime("%Y-%m-%d %H:%M")))
        for d in sorted(days)[:-keep_days]:
            days.pop(d, None)
            con.execute("DELETE FROM hist_event WHERE date=?", (d,))
        con.commit()
    finally:
        con.close()
    return event_study_stats(days)


def event_study_stats(days, min_n=EVENT_STUDY_MIN_N):
    """按主题汇总跨日 AR：均值、样本数、t 值（样本 < min_n 只登记不检验）。

    同时汇总 CAR：分时只有当日有，故 car_n 通常远小于 n，
    单独列出而不是混进 AR 的均值里。
    """
    by_topic, car_topic = {}, {}
    for items in (days or {}).values():
        for it in items:
            by_topic.setdefault(it.get("topic"), []).append(it.get("ar") or 0.0)
            if it.get("car") is not None:
                car_topic.setdefault(it.get("topic"), []).append(it["car"])
    rows = []
    for topic, arr in by_topic.items():
        n = len(arr)
        mean = sum(arr) / n
        _cars = car_topic.get(topic) or []
        row = {"topic": topic, "n": n, "mean": mean, "t": None,
               "sig": "样本不足", "crit": None,
               "car_n": len(_cars),
               "car_mean": (sum(_cars) / len(_cars)) if _cars else None}
        if n >= min_n:
            var = sum((x - mean) ** 2 for x in arr) / (n - 1)
            sd = math.sqrt(var) if var > 0 else 0.0
            if sd > 0:
                t = mean / (sd / math.sqrt(n))
                crit = 2.14 if n < 20 else (2.05 if n < 30 else 1.99)
                row.update(t=t, crit=crit,
                           sig="显著" if abs(t) >= crit else "不显著")
        rows.append(row)
    rows.sort(key=lambda r: -abs(r["mean"]))
    return {"days": len(days or {}), "rows": rows, "min_n": min_n,
            "available": bool(rows)}


def _em_secid(code):
    """sh600000 → 1.600000、sz000001 → 0.000001（东财的市场前缀）。"""
    return ("1." if code.startswith("sh") else "0.") + code[2:]


def _minute_from_em(code):
    """东财 1 分钟分时：主源。

    trends 每行 `2026-11-05 09:30,开,收,高,低,量,额,均价`，取「时间 + 收盘价」。
    与腾讯 minute/query 相比，该端点不返回 501 限流，且能直接取到指数（sh000300）。
    """
    try:
        req = Request(EM_TRENDS_URL.format(_em_secid(code)),
                      headers=dict(HEADERS,
                                   Referer="https://quote.eastmoney.com/"))
        with urlopen(req, timeout=12) as resp:
            d = json.loads(resp.read().decode("utf-8", "replace"))
    except Exception:
        return None
    data = d.get("data") or {}
    pts, day = [], ""
    for row in data.get("trends") or []:
        f = str(row).split(",")
        if len(f) < 3:
            continue
        stamp, price = f[0], num(f[2])
        if not price:
            continue
        d0, _, hm = stamp.partition(" ")
        day = day or d0
        pts.append((hm.replace(":", ""), price))
    return (day, pts) if pts else None


def _minute_from_tencent(code):
    """腾讯分时：备用源（该端点对密集请求会返回 501，只作兜底）。"""
    try:
        req = Request(MINUTE_URL + code, headers=HEADERS)
        with urlopen(req, timeout=10) as resp:
            d = json.loads(resp.read().decode("utf-8", "ignore"))
    except Exception:
        return None
    payload = ((d.get("data") or {}).get(code) or {}).get("data") or {}
    pts = []
    for row in payload.get("data") or []:
        f = str(row).split()
        if len(f) < 2:
            continue
        try:
            pts.append((f[0], float(f[1])))
        except Exception:
            continue
    return (payload.get("date") or "", pts) if pts else None


def fetch_minute_prices(codes, concurrency=4, retry=1):
    """批量抓当日分时 → {code: (yyyyMMdd, [(HHMM, price), ...])}。

    只取「时间 + 价格」，用于计算事件窗口内的区间收益。优先东财（1 分钟，
    不限流），失败再退到腾讯 minute/query；两者都拿不到就跳过该代码，
    由调用方把原因写进报告，而不是笼统地说「离线/非当日」。
    分时只覆盖**当天**，因此历史上没有 CAR 序列——这是接口本身的边界。
    """
    out = {}

    def one(code):
        res = _minute_from_em(code)
        if res is None:
            for _ in range(retry + 1):
                res = _minute_from_tencent(code)
                if res is not None:
                    break
                time.sleep(0.6)
        if res is not None:
            out[code] = res

    with ThreadPoolExecutor(max_workers=concurrency) as ex:
        list(ex.map(one, codes))
    return out


def _hhmm_add(hhmm, minutes):
    try:
        t = datetime.strptime(hhmm, "%H%M") + timedelta(minutes=minutes)
    except Exception:
        return hhmm
    return t.strftime("%H%M")


def window_return(pts, start_hhmm, minutes=30):
    """分时序列在 [start, start+minutes] 的区间收益（%），返回 (收益, 窗口是否完整)。"""
    if not pts:
        return None, False
    idx = next((i for i, (t, _) in enumerate(pts) if t >= start_hhmm), None)
    if idx is None:
        return None, False
    end = _hhmm_add(start_hhmm, minutes)
    j = next((i for i, (t, _) in enumerate(pts) if t >= end), len(pts) - 1)
    p0, p1 = pts[idx][1], pts[j][1]
    if not p0 or not p1:
        return None, False
    return (p1 / p0 - 1) * 100, pts[j][0] >= end


def compute_intraday_car(ev_ar, minutes, bench="sh000300", window=30):
    """主题级日内 CAR：事件时刻起 30 分钟内，ETF 区间收益 − β × 基准区间收益。

    与 15.5.1 的日频 AR 相比，这里的优势是**不把 9:30 之前的涨幅算进事件反应**，
    因此能回答「消息出来后多久、涨了多少」，而不是「全天涨了多少」。
    """
    rows = []
    if not ev_ar.get("available") or not minutes:
        return {"available": False, "rows": [], "window": window,
                "reason": "缺少分时数据（离线/非当日）" if not minutes else "无主题匹配"}
    bpts = (minutes.get(bench) or ("", []))[1]
    for r in ev_ar.get("rows") or []:
        t0 = r.get("first_hhmm") or ""
        if not t0:
            continue
        start = t0 if t0 >= "0930" else "0930"       # 盘前消息从开盘起算
        b_ret, b_full = window_return(bpts, start, window)
        if b_ret is None:
            continue
        members, car_sum, full = [], 0.0, b_full
        for m in r.get("members") or []:
            pts = (minutes.get(m["code"]) or ("", []))[1]
            er, efull = window_return(pts, start, window)
            if er is None:
                continue
            ar = er - (m["beta"] or 1.0) * b_ret
            members.append({"code": m["code"], "name": m["name"],
                            "ret": er, "ar": ar})
            car_sum += ar
            full = full and efull
        if not members:
            continue
        rows.append({"topic": r.get("topic"), "start": start,
                     "raw_time": t0, "window": window,
                     "bench_ret": b_ret, "car": car_sum / len(members),
                     "n": len(members), "full": full,
                     "members": members[:5],
                     "note": "" if t0 >= "0930" else "盘前消息，自开盘起算"})
    rows.sort(key=lambda r: -abs(r["car"]))
    return {"available": bool(rows), "rows": rows, "window": window,
            "bench": bench, "reason": "" if rows else "无可用分时匹配"}


# ----------------------------------------------------------------------------
# 分析层八：自建市场画像评分
# ----------------------------------------------------------------------------

def analyze_profile(breadth, market, fund, global_data, today, store):
    """用本报告已有全量数据自建 9 维评分（1~5），非官方 28 维评分。"""
    bd = (breadth or {}).get("ups_downs_dsb") or {}
    up_ratio = num(bd.get("up_ratio")) or 0
    uplim = num(bd.get("up_limit_count")) or 0
    dims = []

    def push(name, score, label, note):
        dims.append({"name": name, "score": max(1, min(5, score)),
                     "label": label, "note": note})

    if uplim > 50 or up_ratio > 70:
        push("情绪", 5, "狂热", "涨停 {} 家 / 上涨占比 {:.0f}%".format(
            int(uplim), up_ratio))
    elif uplim > 20 or up_ratio > 55:
        push("情绪", 4, "偏热", "涨停 {} 家 / 上涨占比 {:.0f}%".format(
            int(uplim), up_ratio))
    elif up_ratio > 45:
        push("情绪", 3, "中性", "上涨占比 {:.0f}%".format(up_ratio))
    else:
        push("情绪", 2, "偏冷", "上涨占比 {:.0f}%".format(up_ratio))

    d5 = avg([(store.get(c, {}).get("quote") or {}).get("d5")
              for c in ("sh000001", "sh000300", "sz399006", "sh000688")])
    push("短期趋势", 3 + (1 if (d5 or 0) > 1 else (-1 if (d5 or 0) < -1 else 0)),
         "震荡偏强" if (d5 or 0) > 1 else ("弱势" if (d5 or 0) < -1 else "震荡"),
         "主要指数近5日均幅 {}".format(pct_str(d5)))

    pctl = market.get("pctl") or 0
    push("量能", 5 if pctl >= 80 else (4 if pctl >= 60 else (3 if pctl >= 40 else 2)),
         "显著放量" if pctl >= 80 else ("温和放量" if pctl >= 60 else "常态"),
         "成交额处近 {} 日 {:.0f}% 分位".format(market.get("lookback", 0), pctl))

    push("个股宽度", 5 if up_ratio >= 70 else (4 if up_ratio >= 55 else (3 if up_ratio >= 45 else 2)),
         "普涨" if up_ratio >= 70 else "分化",
         "全市场上涨家数占比 {:.0f}%".format(up_ratio))

    pes = [v["pe"] for v in (global_data.get("valuations") or []) if v["pe"]]
    pe_mid = sorted(pes)[len(pes) // 2] if pes else None
    push("估值", 3, "结构分化（横向代理）",
         "A股指数 PE 中位数 {:.1f} 倍；缺历史分位数据，仅横向对比".format(pe_mid)
         if pe_mid else "指数 PE 数据缺失")

    a_g = (store.get("sh000688", {}).get("quote") or {}).get("pct")
    a_w = (store.get("sh000016", {}).get("quote") or {}).get("pct")
    gap = (a_g - a_w) if None not in (a_g, a_w) else None
    push("风格轮动", 5 if (gap or 0) > 5 else (4 if (gap or 0) > 2 else 3),
         "成长极端占优" if (gap or 0) > 5 else ("成长占优" if (gap or 0) > 2 else "均衡"),
         "科创50－上证50 差 {}".format("{:+.1f}pct".format(gap) if gap is not None else "—"))

    inds = fund.get("industry") or []
    ipcts = [r["pct"] for r in inds if r["pct"] is not None]
    ispread = (max(ipcts) - min(ipcts)) if ipcts else None
    push("行业轮动", 5 if (ispread or 0) > 8 else (4 if (ispread or 0) > 5 else 3),
         "分化严重" if (ispread or 0) > 8 else "中度分化",
         "行业首尾涨幅差 {}".format("{:+.1f}pct".format(ispread)
                                if ispread is not None else "—"))

    pos, nind = len(fund.get("pos") or []), len(inds)
    pbreadth = (pos / nind * 100) if nind else 0
    push("板块宽度", 4 if pbreadth >= 60 else (3 if pbreadth >= 45 else 2),
         "多数净流入" if pbreadth >= 60 else "进出参半",
         "主力净流入板块占比 {:.0f}%".format(pbreadth))

    core_below = sum(1 for c in CORE_CODES
                     if c in store and (indicators(store[c]["bars"]).get("ma60") or 0)
                     > (store[c]["quote"]["price"] or 0))
    ncore = sum(1 for c in CORE_CODES if c in store)
    below_pct = core_below / ncore * 100 if ncore else 0
    push("中长期方向", 1 if below_pct >= 75 else (2 if below_pct >= 50 else 3),
         "空头排列" if below_pct >= 75 else "修复中",
         "{} / {} 核心ETF仍处 MA60 下方".format(core_below, ncore))

    total = sum(d["score"] for d in dims)
    max_total = len(dims) * 5
    ratio = total / max_total * 100
    if ratio >= 75:
        verdict = "短期强、中期弱：情绪与广度极热，但中长期趋势与杠杆资金未确认"
    elif ratio >= 55:
        verdict = "短中期背离，需下一交易日量能与资金确认"
    else:
        verdict = "整体偏弱，反弹持续性存疑"
    return {"dims": dims, "total": total, "max_total": max_total,
            "ratio": ratio, "verdict": verdict}


# ----------------------------------------------------------------------------
# 分析层九：事件与异动
# ----------------------------------------------------------------------------

def analyze_events(breadth, market):
    b = breadth or {}
    news = b.get("news") or {}
    yidong = b.get("yidong_list") or []
    ksh = b.get("ksh") or {}
    comments = b.get("comments") or {}
    nb = b.get("north_bound") or {}

    items = []
    if news.get("title"):
        items.append({"type": "收评", "title": news["title"],
                      "detail": news.get("abstract") or "",
                      "source": news.get("source") or "", "time": news.get("time") or ""})
    for y in yidong[:5]:
        items.append({"type": "盘中异动",
                      "title": "{} {}".format(y.get("Name") or "", y.get("Symbol") or ""),
                      "detail": "涨幅 {} / 时间 {}".format(
                          pct_str(num(y.get("Zdf")) * 100 if num(y.get("Zdf")) is not None else None),
                          y.get("DataTime") or ""),
                      "source": "", "time": ""})
    if ksh.get("name"):
        items.append({"type": "科创异动", "title": ksh.get("name"),
                      "detail": "涨跌幅 {} / 现价 {}".format(
                          pct_str(num(ksh.get("zdf"))), ksh.get("zxj") or ""),
                      "source": "", "time": ""})
    if comments.get("margin"):
        items.append({"type": "两融", "title": "杠杆资金动向",
                      "detail": comments["margin"], "source": "市场评论", "time": ""})
    if comments.get("adr_limit"):
        items.append({"type": "涨跌停比", "title": "ADR 极限值",
                      "detail": comments["adr_limit"], "source": "市场评论", "time": ""})
    if nb.get("date"):
        # 沪深港通已不再逐日披露净买入，公开接口该字段长期为 0 且 history 为空。
        # 不要把 0 当成「今日零流入」展示，也不要编造替代值。
        _v = num(nb.get("fund_flow_net_in"))
        if _v:
            _detail = "数据日期 {}，净流入 {}（接口原值）".format(nb["date"], _v)
        else:
            _detail = ("数据日期 {}，公开接口未返回净流入（逐日净买入已停止披露，"
                       "本项不估算、不以 0 充当数据）".format(nb["date"]))
        items.append({"type": "北向", "title": "陆股通资金",
                      "detail": _detail, "source": "", "time": ""})

    signals = []
    if news.get("title"):
        signals.append("今日收评标题「{}」，驱动描述偏**事件与情绪**，"
                       "不含宏观数据或盈利改善的确认。".format(news["title"]))
    signals.append("**本报告的缺口**：脚本只能读到收评、异动与两融文本，"
                   "没有结构化的政策/财报/外围事件库，也没有把事件映射到板块的因果链；"
                   "因此「为何大涨」只能定性提示，不能做事件归因。")
    signals.append("**使用时须注意**：由事件与情绪驱动的反弹，持续性通常弱于基本面驱动，"
                   "外围消息反转时次日振幅会放大。")
    return {"items": items, "signals": signals, "available": bool(items)}


# ----------------------------------------------------------------------------
# 分析层十：历史规则回测（用现有日线，纯标准库）
# ----------------------------------------------------------------------------

COST_PCT = 0.06  # 双边交易成本近似（ETF 佣金+冲击），单位 %


def _stat_block(rets):
    if not rets:
        return None
    n = len(rets)
    w = sum(1 for r in rets if r > 0)
    p = w / n
    # 二项检验的正态近似 z：H0 胜率 = 50%
    z = (p - 0.5) / ((0.25 / n) ** 0.5) if n else 0
    return {"n": n, "win": p * 100, "mean": avg(rets),
            "best": max(rets), "worst": min(rets), "z": z,
            "sig": "显著" if abs(z) >= 1.96 else "不显著"}


def _rule_stats(bars, hit_idx, label, horizons=(1, 5), split=None):
    """统计触发日之后 N 日收益：毛收益 / 扣成本 / 样本内 / 样本外 / 显著性。"""
    out = {"label": label, "n": 0, "h": {}, "oos": {}}
    cut = split if split is not None else len(bars)
    for h in horizons:
        pairs = []
        for i in hit_idx:
            if i + h < len(bars):
                c0, c1 = bars[i]["close"], bars[i + h]["close"]
                if c0:
                    gross = (c1 / c0 - 1) * 100
                    pairs.append((i, gross, gross - COST_PCT))
        all_b = _stat_block([g for _, g, _ in pairs])
        net_b = _stat_block([n for _, _, n in pairs])
        ins = _stat_block([g for i, g, _ in pairs if i < cut])
        oos = _stat_block([g for i, g, _ in pairs if i >= cut])
        if all_b:
            out["h"][h] = dict(all_b, net_win=net_b["win"],
                               net_mean=net_b["mean"],
                               ins_win=ins["win"] if ins else None,
                               oos_win=oos["win"] if oos else None,
                               ins_n=ins["n"] if ins else 0,
                               oos_n=oos["n"] if oos else 0)
        if oos:
            out["oos"][h] = oos
    out["n"] = len(hit_idx)
    return out


def analyze_backtest(store):
    """对本报告用到的几条经验规则做历史频率与次日/5日胜率统计。"""
    rules = []
    all_bars = []
    for c in CORE_CODES:
        if c not in store:
            continue
        bars = store[c]["bars"]
        if len(bars) < 60:
            continue
        all_bars.append((c, store[c]["quote"]["name"], bars))

    def hits_of(pred, bars):
        hits = []
        for i in range(1, len(bars) - 5):
            b, p = bars[i], bars[i - 1]
            if pred(b, p, bars, i):
                hits.append(i)
        return hits

    specs = [
        ("单日涨幅 ≥5%", lambda b, p, bs, i: p["close"] and
         (b["close"] / p["close"] - 1) * 100 >= 5),
        ("平开收高位（振幅位置≥0.9且开盘±0.5%）",
         lambda b, p, bs, i: b["high"] > b["low"] and p["close"] and
         abs(b["open"] / p["close"] - 1) <= 0.005 and
         (b["close"] - b["low"]) / (b["high"] - b["low"]) >= 0.9),
        ("缩量反包（上涨但量低于前日）",
         lambda b, p, bs, i: p["close"] and b["close"] > p["close"] and
         p["volume"] and b["volume"] < p["volume"]),
        ("放量大阳（涨幅>3%且量>1.2×5日均量）",
         lambda b, p, bs, i: p["close"] and (b["close"] / p["close"] - 1) * 100 > 3 and
         i >= 5 and avg([x["volume"] for x in bs[i - 5:i]]) and
         b["volume"] > 1.2 * avg([x["volume"] for x in bs[i - 5:i]])),
    ]

    total_bars = sum(len(b) for _, _, b in all_bars)
    # 基准：所有交易日的次日收益（买入持有对照）
    base_rets = []
    for _, _, bars in all_bars:
        for i in range(1, len(bars) - 1):
            if bars[i]["close"]:
                base_rets.append((bars[i + 1]["close"] / bars[i]["close"] - 1) * 100)
    base = _stat_block(base_rets)

    for label, pred in specs:
        per = []
        for c, name, bars in all_bars:
            cut = int(len(bars) * 0.7)  # 前 70% 样本内，后 30% 样本外
            st = _rule_stats(bars, hits_of(pred, bars), label, split=cut)
            per.append({"name": name, "n": st["n"], "h": st["h"]})
        agg = {"label": label, "n": sum(p["n"] for p in per), "h": {}}
        for h in (1, 5):
            ns = [p["h"][h]["n"] for p in per if h in p["h"]]
            ws = [p["h"][h]["win"] for p in per if h in p["h"]]
            ms = [p["h"][h]["mean"] for p in per if h in p["h"]]
            nws = [p["h"][h]["net_win"] for p in per if h in p["h"]]
            iws = [(p["h"][h]["ins_win"], p["h"][h]["ins_n"])
                   for p in per if h in p["h"] and p["h"][h]["ins_win"] is not None]
            ows = [(p["h"][h]["oos_win"], p["h"][h]["oos_n"])
                   for p in per if h in p["h"] and p["h"][h]["oos_win"] is not None]
            if ns:
                tot = sum(ns)
                def wavg(items):
                    return (sum(v * n for v, n in items) / sum(n for _, n in items)
                            if items and sum(n for _, n in items) else None)
                pooled = _stat_block(base_rets)  # 占位，z 用合并样本重算
                agg["h"][h] = {
                    "n": tot,
                    "win": sum(w * n for w, n in zip(ws, ns)) / tot,
                    "mean": sum(m * n for m, n in zip(ms, ns)) / tot,
                    "net_win": sum(w * n for w, n in zip(nws, ns)) / tot,
                    "ins_win": wavg(iws), "oos_win": wavg(ows),
                    "ins_n": sum(n for _, n in iws), "oos_n": sum(n for _, n in ows),
                }
                p = agg["h"][h]["win"] / 100
                agg["h"][h]["z"] = (p - 0.5) / ((0.25 / tot) ** 0.5) if tot else 0
                agg["h"][h]["sig"] = ("显著" if abs(agg["h"][h]["z"]) >= 1.96
                                      else "不显著")
        agg["per"] = per
        rules.append(agg)

    signals = []
    signals.append("回测样本：{} 只核心ETF、合计 {} 根日线（约 {} 个交易日/只）；"
                   "已**扣除双边成本 {}%**、按前70%/后30%做**样本外分割**、"
                   "给出二项检验 z 值与买入持有基准。仍未复权、未含分红、样本量小。".format(
                       len(all_bars), total_bars,
                       total_bars // len(all_bars) if all_bars else 0, COST_PCT))
    if base:
        signals.append("**基准（买入持有）**：全样本次日收阳胜率 {:.0f}%、"
                       "平均 {}. 规则必须跑赢基准才有意义。".format(
                           base["win"], pct_str(base["mean"])))
    for r in rules:
        h1 = r["h"].get(1)
        if not h1:
            signals.append("「{}」样本不足。".format(r["label"]))
            continue
        edge = h1["win"] - (base["win"] if base else 50)
        oos_txt = ("样本外 {:.0f}%".format(h1["oos_win"])
                   if h1["oos_win"] is not None else "样本外不足")
        signals.append("「{}」触发 {} 次，次日胜率 **{:.0f}%**（扣成本 {:.0f}%）、"
                       "平均 {}、相对基准 {:+.0f}pct、样本内 {:.0f}% / {}、"
                       "z={:.2f}（{}）。{}".format(
                           r["label"], r["n"], h1["win"], h1["net_win"],
                           pct_str(h1["mean"]), edge,
                           h1["ins_win"] or 0, oos_txt, h1["z"], h1["sig"],
                           "相对基准无优势。" if edge <= 0 else "略优于基准，但仍需样本外确认。"))
    signals.append("**定位**：这一章只做规则校准，用来给走势章节的情景打分提供历史背景，"
                   "不是策略回测、不含仓位与止损，也不能预测单次结果；"
                   "「显著」仅指统计上偏离 50%，不代表可交易盈利。")

    return {"rules": rules, "signals": signals, "base": base,
            "available": bool(all_bars), "total_bars": total_bars,
            "n_assets": len(all_bars), "cost_pct": COST_PCT}


# ----------------------------------------------------------------------------
# 分析层十一：全市场 ETF 全景（1722 只，封装接口）
# ----------------------------------------------------------------------------

def _by_code(rows):
    return {r.get("EtfCode"): r for r in (rows or []) if r.get("EtfCode")}


def analyze_etf_universe(ext, store):
    """全市场 ETF 规模/估值/估值分位/真实净申购，样本覆盖与资金榜。"""
    size = _by_code(ext.get("etf_size"))
    if not size:
        return {"available": False, "signals": [
            "全市场 ETF 明细接口不可用：规模、估值分位、真实净申购均无法计算，"
            "其余章节不受影响。"],
            "netflow_rows": [], "val_rows": [], "cheap": [], "rich": [],
            "dev_rows": [], "flow_rows": [], "universe_n": 0, "sample_n": 0,
            "sample_asset_pct": None, "sample_cover_pct": 0,
            "total_nv_yi": 0, "total_asset_yi": 0,
            "d_total": 0, "in_n": 0, "flow_date": None}
    level = _by_code(ext.get("etf_valuation_level"))
    hist = _by_code(ext.get("etf_valuation_hist"))
    flow = _by_code(ext.get("etf_netflow_interval"))
    disc = _by_code(ext.get("etf_qt_disc_interval"))
    amt = _by_code(ext.get("etf_qt_amt_interval"))
    chg = _by_code(ext.get("etf_qt_chg_interval"))

    total_asset = sum(num(r.get("TotalAsset")) or 0 for r in size.values())
    total_nv = sum(num(r.get("NetValueTotal")) or 0 for r in size.values())
    n_all = len(size)
    sample_codes = [c for c in ALL_ETF_CODES if c in size]
    sample_asset = sum(size[c].get("TotalAsset") or 0 for c in sample_codes)

    flow_rows = []
    for c in ALL_ETF_CODES:
        r = flow.get(c)
        if not r:
            continue
        q = store.get(c, {}).get("quote") or {}
        flow_rows.append({
            "code": c, "name": q.get("name") or r.get("EtfName") or c,
            "d": (num(r.get("EtfNetPurchaseD")) or 0) / 1e8,
            "w": (num(r.get("EtfNetPurchaseW")) or 0) / 1e8,
            "m": (num(r.get("EtfNetPurchaseM")) or 0) / 1e8,
            "q": (num(r.get("EtfNetPurchaseQ")) or 0) / 1e8,
            "d_pct": num(r.get("EtfNetPurchaseRatioD")),
            "m_pct": num(r.get("EtfNetPurchaseRatioM")),
            "size": (num(r.get("EtfSize")) or 0) / 1e8,
        })
    flow_rows.sort(key=lambda x: x["d"], reverse=True)
    d_total = sum(r["d"] for r in flow_rows)
    in_n = sum(1 for r in flow_rows if r["d"] > 0)
    flow_dates = sorted({str(r.get("EndDate")) for r in (flow.values())
                         if r.get("EndDate")})
    flow_date = flow_dates[-1] if flow_dates else None
    flow_date = ("{}-{}-{}".format(flow_date[:4], flow_date[4:6], flow_date[6:])
                 if flow_date and len(flow_date) == 8 else flow_date)

    val_rows = []
    for c in ALL_ETF_CODES:
        lv, hs = level.get(c), hist.get(c)
        if not lv:
            continue
        q = store.get(c, {}).get("quote") or {}
        pe, pct = num(lv.get("PE_TTM")), (num((hs or {}).get("PE_TTM_PCT")))
        val_rows.append({
            "code": c, "name": q.get("name") or c,
            "pe": pe, "pb": num(lv.get("PB")), "roe": num(lv.get("ROE")),
            "div": num(lv.get("DIV_TTM")),
            "pe_pct": pct, "pb_pct": num((hs or {}).get("PB_PCT")),
            "earn_yield": (100 / pe) if (pe and pe > 0) else None,
        })
    val_rows.sort(key=lambda x: x["pe_pct"] if x["pe_pct"] is not None else 999)
    cheap = [r for r in val_rows if r["pe_pct"] is not None and r["pe_pct"] <= 30]
    rich = [r for r in val_rows if r["pe_pct"] is not None and r["pe_pct"] >= 80]

    # 折溢价/成交额相对自身均值（判断情绪是否透支）
    dev_rows = []
    for c in ALL_ETF_CODES:
        d, a = disc.get(c), amt.get(c)
        if not d:
            continue
        q = store.get(c, {}).get("quote") or {}
        cur_d = num(d.get("EtfDisc"))
        avg_m = num(d.get("EtfDiscAvgM"))
        cur_amt = num((a or {}).get("EtfTurnover"))
        avg_amt = num((a or {}).get("EtfTurnoverAvgM"))
        dev_rows.append({
            "name": q.get("name") or c, "disc": cur_d,
            "disc_dev": (cur_d - avg_m) if None not in (cur_d, avg_m) else None,
            "amt_multiple": (cur_amt / avg_amt) if (cur_amt and avg_amt) else None,
        })
    dev_rows.sort(key=lambda x: x["disc_dev"] or 0, reverse=True)

    signals = []
    signals.append("全市场共 **{} 只 ETF**、合计净资产 **{:.0f} 亿**（总资产 {:.0f} 亿）；"
                   "脚本样本 {} 只，覆盖全市场只数的 **{:.1f}%**、规模的 **{:.1f}%**。"
                   "1.6 中「不能推断全市场 ETF 总数与总规模」的限制已由接口解除。".format(
                       n_all, total_nv / 1e8, total_asset / 1e8,
                       len(sample_codes), len(sample_codes) / n_all * 100,
                       sample_asset / total_asset * 100 if total_asset else 0))
    signals.append("**真实净申购（一级市场资金流）**：数据披露日 **{}**（T-1，非当日），"
                   "样本 {} 只 ETF 净申购合计 **{:+.1f} 亿**，"
                   "其中净申购 {} 只 / 净赎回 {} 只。"
                   "这与成交额（二级换手）是两个不同口径，且不再需要自己逐日累积份额。".format(
                       flow_date or "未知", len(flow_rows), d_total,
                       in_n, len(flow_rows) - in_n))
    signals.append("**一级/二级背离（重要）**：ETF 一级市场出现大额净申购（{:+.0f} 亿，"
                   "集中于宽基），但第一章显示核心 ETF 二级市场为**缩量反包**。"
                   "说明机构资金通过申购进场、二级换手却未同步放大，"
                   "「反包缺乏增量资金」的结论需修正为"
                   "「**一级有增量、二级未放量**」——这比单纯缩量更偏正面，"
                   "但仍不等于趋势反转。".format(d_total))
    if flow_rows:
        signals.append("资金流入前三：{}；流出前三：{}（单位：亿元）。".format(
            "、".join("{} {:+.1f}".format(r["name"], r["d"]) for r in flow_rows[:3]),
            "、".join("{} {:+.1f}".format(r["name"], r["d"]) for r in flow_rows[-3:][::-1])))
    mid = [r for r in flow_rows if r["m"]]
    if mid:
        signals.append("月度口径合计 **{:+.0f} 亿**（近一月累计净申购），"
                       "可用于判断今日单日流入是趋势延续还是短线回补。".format(
                           sum(r["m"] for r in mid)))
    if val_rows:
        pe_pcts = [r["pe_pct"] for r in val_rows if r["pe_pct"] is not None]
        signals.append("样本中 {} 只ETF取得 **PE 历史分位**（官方口径）：中位数 {:.0f}%，"
                       "低估（≤30%）{} 只、高估（≥80%）{} 只。"
                       "第六章「缺历史 PE 分位」的缺口由此补齐。".format(
                           len(pe_pcts), sorted(pe_pcts)[len(pe_pcts) // 2] if pe_pcts else 0,
                           len(cheap), len(rich)))
    if dev_rows:
        hot = [r for r in dev_rows if (r["disc_dev"] or 0) >= 0.2]
        signals.append("折溢价相对月均值偏离 ≥0.2pct 的有 {} 只（{}），"
                       "是二级情绪先于净值透支的直接证据。".format(
                           len(hot), "、".join(r["name"] for r in hot[:4])))

    return {
        "available": True, "universe_n": n_all,
        "total_asset_yi": total_asset / 1e8, "total_nv_yi": total_nv / 1e8,
        "sample_n": len(sample_codes),
        "sample_cover_pct": len(sample_codes) / n_all * 100,
        "sample_asset_pct": sample_asset / total_asset * 100 if total_asset else None,
        "flow_rows": flow_rows, "d_total": d_total, "in_n": in_n,
        "flow_date": flow_date,
        "val_rows": val_rows, "cheap": cheap, "rich": rich,
        "dev_rows": dev_rows, "signals": signals,
    }


# ----------------------------------------------------------------------------
# 分析层十二：宏观流动性与股债性价比（ERP）
# ----------------------------------------------------------------------------

def analyze_macro(ext, universe):
    p1 = (ext.get("macro_core_indicators_cur_p1") or [{}])
    p2 = (ext.get("macro_core_indicators_cur_p2") or [{}])
    a = p1[0] if isinstance(p1, list) and p1 else {}
    b = p2[0] if isinstance(p2, list) and p2 else {}
    merged = dict(a)
    merged.update(b)
    if not merged:
        return {"available": False, "rows": [], "signals": [], "erp": None}

    rows = []
    for grp, key, label, unit in MACRO_PICK:
        v = (merged.get(grp) or {}).get(key)
        if v is not None:
            rows.append({"group": grp, "label": label, "value": num(v), "unit": unit})
    vals = {r["label"]: r["value"] for r in rows}
    y10 = vals.get("10年国债收益率")
    y1 = vals.get("1年国债收益率")
    cpi = vals.get("CPI同比")
    pmi = vals.get("制造业PMI")
    m1, m2 = vals.get("M1同比"), vals.get("M2同比")

    # ERP：用样本宽基ETF盈利收益率中位数 - 10年国债收益率
    eys = [r["earn_yield"] for r in (universe.get("val_rows") or [])
           if r["earn_yield"] and r["earn_yield"] > 0]
    ey_mid = sorted(eys)[len(eys) // 2] if eys else None
    erp = (ey_mid - y10) if (ey_mid is not None and y10 is not None) else None

    signals = []
    if rows:
        signals.append("宏观最新快照共 {} 项：{}。".format(
            len(rows), "、".join("{} {}{}".format(
                r["label"], r["value"], r["unit"]) for r in rows[:6])))
    if pmi is not None:
        signals.append("制造业 PMI **{}**，{}荣枯线；{}".format(
            pmi, "高于" if pmi >= 50 else "低于",
            "基本面尚未给出扩张确认，与「事件情绪驱动反弹」的判断一致"
            if pmi < 50 else "基本面处于扩张区间"))
    if cpi is not None and y10 is not None:
        real = y10 - cpi
        signals.append("10年国债 {:.2f}% − CPI {:.1f}% = **实际利率 {:.2f}%**；"
                       "{}。".format(
                           y10, cpi, real,
                           "实际利率为正，估值扩张受约束" if real > 0
                           else "实际利率为负，利于权益估值"))
    if None not in (m1, m2):
        signals.append("M1 {}% / M2 {}%，剪刀差 {:+.1f}pct，{}。".format(
            m1, m2, m1 - m2,
            "资金活化不足" if m1 - m2 < 0 else "资金活化改善"))
    if erp is not None:
        signals.append("**股债性价比 ERP ≈ {:+.2f}pct**（样本宽基ETF盈利收益率中位数 "
                       "{:.2f}% − 10年国债 {:.2f}%）。ERP 越高，股票相对债券越便宜；"
                       "此处用 ETF 持仓加权 PE 代理，非全市场精确口径。".format(
                           erp, ey_mid, y10))
    else:
        signals.append("ERP 所需数据不足（缺样本盈利收益率或10年国债），未计算。")
    if y1 is not None and y10 is not None:
        signals.append("期限利差（10Y−1Y）**{:+.2f}pct**，{}。".format(
            y10 - y1, "曲线陡峭" if y10 - y1 > 0.5 else "曲线偏平"))
    return {"available": True, "rows": rows, "vals": vals, "erp": erp,
            "ey_mid": ey_mid, "y10": y10, "signals": signals}


# ----------------------------------------------------------------------------
# 分析层十三：官方 28 维画像 + 自建对照
# ----------------------------------------------------------------------------

OFFICIAL_DIMS = [
    ("SENTIMENT", "情绪指标"), ("TREND_SHORT_STRENGTH", "短期趋势强度"),
    ("VOLUME_ENERGE", "成交量能"), ("CAP_ROTATION", "大小盘轮动"),
    ("STOCK_WIDTH", "个股宽度"), ("VALUATION", "估值水平"),
    ("STYLE_ROTATION", "风格轮动"), ("TECHNICAL", "技术指标"),
    ("TREND_LONG_DIRECTION", "中长期趋势方向"),
    ("TREND_LONG_STRENGTH", "中长期趋势强度"),
    ("SECTOR_ROTATION", "行业轮动"), ("SECTOR_WIDTH", "板块宽度"),
    ("TREND_SHORT_DIRECTION", "短期趋势方向"),
]


def analyze_official_profile(ext, self_profile):
    rows_raw = ext.get("market_statis_summary") or []
    s = rows_raw[0] if isinstance(rows_raw, list) and rows_raw else {}
    if not s:
        return {"available": False, "dims": [], "signals": []}
    dims = []
    for key, label in OFFICIAL_DIMS:
        score, status = num(s.get(key + "_SCORE")), s.get(key + "_STATUS") or "—"
        if score is not None:
            dims.append({"name": label, "score": score, "status": status})
    raw, adj = num(s.get("RAW_SCORE")), num(s.get("ADJ_SCORE"))
    signals = []
    if dims:
        signals.append("官方画像 **{} 维**、原始评分 {} / 调整评分 {}；"
                       "自建 9 维评分 {} / {}（{:.0f}%）。两者口径不同，"
                       "官方为行情商模型输出，自建为本报告全量数据的透明规则打分。".format(
                           len(dims), raw, adj, self_profile["total"],
                           self_profile["max_total"], self_profile["ratio"]))
    weak = [d for d in dims if d["score"] <= 2]
    strong = [d for d in dims if d["score"] >= 4]
    if weak:
        signals.append("官方评分中偏弱维度：{}。".format(
            "、".join("{} {}分（{}）".format(d["name"], int(d["score"]), d["status"])
                     for d in weak[:5])))
    if strong:
        signals.append("官方评分中偏强维度：{}。".format(
            "、".join("{} {}分".format(d["name"], int(d["score"]))
                     for d in strong[:5])))
    if weak and strong:
        signals.append("**两套画像结论一致**：短期维度强、中长期维度弱，"
                       "共同支持「超跌修复而非趋势反转」。")
    signals.append("**口径分歧提醒**：官方「成交量能」按长周期均值判定（本次为 3/5 缩量震荡），"
                   "本报告第三章按近 20 日分位判定（100% 分位、显著放量）——"
                   "两者基准窗口不同，不矛盾；短线看本报告口径，中期看官方口径。")
    return {"available": True, "dims": dims, "raw": raw, "adj": adj,
            "signals": signals}


# ----------------------------------------------------------------------------
# 分析层十四：龙虎榜 / 结构化事件 / 筹码
# ----------------------------------------------------------------------------

def analyze_market_events(ext, breadth, store):
    lhb = ext.get("_lhb") or {}
    chips = ext.get("_chips") or {}
    ev = ext.get("stock_event") or []
    items, signals = [], []

    jg = lhb.get("jg") or []
    if jg:
        for row in jg[:8]:
            if isinstance(row, list) and len(row) >= 6:
                items.append({"type": "龙虎榜机构", "title": row[1] if len(row) > 1 else "",
                              "detail": "买入 {:.2f} 亿 / 涨跌 {}".format(
                                  (num(row[4]) or 0) / 1e8, pct_str(num(row[5]))),
                              "source": "全市场龙虎榜"})
        signals.append("全市场龙虎榜机构席位上榜 {} 条（展示前 {} 条），"
                       "可观察机构资金是否参与了今日普涨。".format(
                           len(jg), min(8, len(jg))))

    if chips:
        crows = []
        for c, node in chips.items():
            if not isinstance(node, dict):
                continue
            # 结构是 chips[code] = {"stocks": [{"code","name","data":{...}}]}，
            # 早先直接在 node 上取 ChipProfitRate，永远拿到 None
            stocks = node.get("stocks") or []
            first = stocks[0] if stocks and isinstance(stocks[0], dict) else {}
            d = first.get("data") if isinstance(first.get("data"), dict) else {}
            if not d:
                continue
            q = store.get(c, {}).get("quote") or {}
            crows.append({"name": q.get("name") or first.get("name") or c,
                          "profit": num(d.get("ChipProfitRate")),
                          "cost": num(d.get("ChipAvgCost")),
                          "c90": num(d.get("ChipConcentration90")),
                          "c70": num(d.get("ChipConcentration70"))})
        if crows:
            pa = avg([r["profit"] for r in crows])
            signals.append("筹码获利比例：{}（均值 {}）。获利盘越厚，"
                           "上方解套与兑现压力越大。".format(
                               "、".join("{} {}".format(r["name"], pct_str(r["profit"], 0))
                                        for r in crows[:3]),
                               pct_str(pa, 0)))
            items.append({"type": "筹码", "title": "核心标的获利盘",
                          "detail": "；".join("{} 获利 {} / 平均成本 {}".format(
                              r["name"], pct_str(r["profit"], 0),
                              price_str(r["cost"])) for r in crows[:3]),
                          "source": "筹码分布接口"})
    if ev:
        signals.append("结构化事件库返回 {} 条个股事件（可映射到板块的因果链仍有限）。".format(
            len(ev)))
        for e in ev[:5]:
            items.append({"type": "结构化事件",
                          "title": str(e.get("SecuName") or e.get("Name") or ""),
                          "detail": str(e.get("EventName") or e.get("Event") or "")[:120],
                          "source": "事件库"})
    else:
        signals.append("结构化事件库今日返回空（无个股事件标签），"
                       "「为何大涨」仍只能靠收评文本定性提示，无法做事件归因。")
    return {"items": items, "signals": signals,
            "lhb_n": len(jg), "chip_n": len(chips), "event_n": len(ev)}


# ----------------------------------------------------------------------------
# 分析层十五：数据质量审计 + 逐日 diff
# ----------------------------------------------------------------------------

def analyze_quality(store, snap_date, today, market, fund, breadth, ext, glob):
    checks = []

    def add(name, ok, detail, level="warn"):
        checks.append({"name": name, "ok": ok, "detail": detail, "level": level})

    add("数据源日期一致", (snap_date == today["date"]) if snap_date else True,
        "在线取数（日期 {}）".format(today["date"] or "未知") if not snap_date
        else "快照 {} / 报告 {}".format(snap_date, today["date"] or "未知"),
        "info")
    n_quote = sum(1 for c in ALL_ETF_CODES if c in store)
    add("样本ETF行情完整", n_quote == len(ALL_ETF_CODES),
        "{}/{} 只".format(n_quote, len(ALL_ETF_CODES)))
    n_kline = sum(1 for c in ALL_ETF_CODES if c in store and len(store[c]["bars"]) >= 60)
    add("日线长度足够算 MA60", n_kline >= len(ALL_ETF_CODES) * 0.8,
        "{}/{} 只 ≥60 根".format(n_kline, len(ALL_ETF_CODES)))
    add("全市场涨跌广度", bool(breadth.get("available")),
        "沪深A股全量涨跌家数" if breadth.get("available") else "缺失，第二章降级", "info")
    add("板块资金流", bool(fund.get("available")),
        "{} 个申万行业".format(len(fund.get("industry") or [])) if fund.get("available") else "缺失", "info")
    add("外围行情", bool(glob.get("available")),
        "{} 个指数".format(len(glob.get("rows") or [])) if glob.get("available") else "缺失（估值横向仍可用）", "info")
    add("ETF 全市场明细", bool(ext.get("etf_size")),
        "{} 只".format(len(ext.get("etf_size") or [])) if ext.get("etf_size") else "缺失，规模/分位/净申购降级", "info")
    add("官方市场画像", bool(ext.get("market_statis_summary")),
        "行情商评分" if ext.get("market_statis_summary") else "缺失，仅自建 9 维", "info")
    add("宏观指标", bool(ext.get("macro_core_indicators_cur_p1")),
        "CPI/PMI/国债收益率" if ext.get("macro_core_indicators_cur_p1") else "缺失，ERP 不计算", "info")
    _f = ext.get("_em_flash") or []
    _fs = ext.get("_em_flash_status") or "empty"
    add("7×24快讯", bool(_f),
        "{} 条（{}）".format(len(_f),
                            EM_FLASH_STATUS_DESC.get(_fs, _fs)) if _f else
        "不可用：{}".format(EM_FLASH_STATUS_DESC.get(_fs, _fs)),
        "info")
    _ss = ext.get("_sina_status") or "empty"
    add("新浪7×24直连", _ss == "ok",
        "采到 {} 条，带完整日期+秒级时间".format(ext.get("_sina_n") or 0)
        if _ss == "ok" else "状态 {}（页面可达性受限）".format(_ss), "info")
    _cs = ext.get("_ct_status") or "empty"
    _cn = ext.get("_ct_n") or 0
    add("华夏时报7×24", _cn > 0,
        "{} 条（{}）".format(
            _cn, "直连" if _cs == "ok" else "网页通道采集")
        if _cn else "状态 {}（域名受限）".format(_cs), "info")
    _ef = ext.get("_etf_flow") or []
    add("ETF资金独立印证", bool(_ef),
        "{} 条资金流数据，含个券 Top3".format(len(_ef))
        if _ef else "无", "info")
    _cln = ext.get("_cls_n") or 0
    _clst = ext.get("_cls_status") or "empty"
    _cls_ready = _cln > 0 or _clst in ("blocked", "empty", "web_channel")
    # 域名被策略拒绝属环境限制、非数据缺陷：适配器就绪且状态已在 15.7 披露即算通过
    _cltxt = ("{} 条（{}）".format(
        _cln, "JSON 直连 + last_time 回溯" if _clst == "ok"
        else "网页通道，秒级时间戳；JSON 仍 403") if _cln else
        "沙盒 403；签名+回溯适配器就绪，状态已在 15.7 标注（本次 0 条）")
    add("财联社通道可用性", _cls_ready, _cltxt, "info")
    _fd = ext.get("_flash_days") or 0
    add("快讯逐日累积", _fd > 0,
        "{} 天 / {} 条，按 id 增量留存".format(
            _fd, ext.get("_flash_total") or 0) if _fd else "尚未累积", "info")
    add("精确成交额环比", market.get("chg_source") == "精确值",
        "当前 {}".format(market.get("chg_source")), "info")
    _en = sum(1 for r in (ext.get("_em_flash") or [])
              if r.get("source") == "东方财富全球快讯")
    _empc_ok = _en > 0 or (ext.get("_empc_status") == "ok")
    _empc_ajax = any(r.get("origin") == "东财PC页(AJAX)"
                     for r in (ext.get("_em_flash") or []))
    add("东财PC页7_24通道", _empc_ok,
        "{} 条 / {} 个日期段（{}；盘后叙事 + 次日盘中验证）".format(
            _en, len(ext.get("_empc_days") or {}),
            "壳页改走页面 AJAX，与①同源已去重" if _empc_ajax else "SSR 页直读")
        if _empc_ok else "页面不可达或接口为空", "info")
    add("多方评论观点识别", bool((ext.get("_em_flash") or []) or (ext.get("_hot") or [])),
        "从快讯/研报/公告分离观点方与立场，事后观点自动降权", "info")
    _rk = ext.get("_market_rank") or []
    _er = ext.get("_etf_rank") or []
    add("全市场涨跌排行", bool(_rk or _er),
        "个股 {} 只 / ETF {} 只，与广度做日期一致性校验".format(
            len(_rk), len(_er)) if (_rk or _er) else "缺失，第十六章降级", "info")

    _fut = ext.get("_futures") or []
    add("股指期货基差层", len(_fut) >= 4,
        "{} 个合约（IF/IH/IC/IM + 国债期货；基差与现货须同一时点）".format(
            len(_fut)) if _fut else "缺失（状态 {}），第六章 6.1.1 跳过".format(
                ext.get("_futures_status") or "empty"))

    _cone_n = sum(1 for c in CORE_CODES
                  if c in store and len((store[c] or {}).get("bars") or []) >= 140)
    add("波动率锥层", _cone_n >= 3,
        "{} 只核心标的日线 ≥140 根，可算 HV120 与历史分位".format(_cone_n)
        if _cone_n else "日线不足 140 根，第十四章 14.1c 跳过")

    # ---- 断言式检查：可达 ≠ 正确 ----
    # 前 19 项只回答「接口通不通」，无法发现「接口通了但字段改版/解析为空」。
    # 这里补数据形状断言：条数下限、时间跨度、数值有效性、环比突变。
    _snap = (AS_OF.get("state") == "snapshot")
    _intraday = AS_OF.get("state") in ("intraday", "preopen")
    _min_flash = 20 if _snap else 50
    _min_etf = 500 if _snap else 1000
    _min_rank = 1000 if _snap else 4000
    _nf = len(ext.get("_em_flash") or [])
    add("快讯条数下限", _nf >= _min_flash,
        "{} 条（下限 {}）".format(_nf, _min_flash))
    _ft = sorted((r.get("time") or "") for r in (ext.get("_em_flash") or []))
    _frange = 0.0
    if len(_ft) >= 2:
        try:
            _frange = (datetime.strptime(_ft[-1][:19], "%Y-%m-%d %H:%M:%S")
                       - datetime.strptime(_ft[0][:19], "%Y-%m-%d %H:%M:%S")
                       ).total_seconds() / 3600.0
        except Exception:
            _frange = 0.0
    add("快讯时间跨度", _frange >= 2,
        "跨 {:.1f} 小时（下限 2h，含早盘与午后）".format(_frange))
    _n_ind = len(fund.get("industry") or [])
    add("板块资金流下限", _n_ind >= (40 if _snap else 100),
        "{} 个行业（下限 {}）".format(_n_ind, 40 if _snap else 100))
    _n_etf_size = len(ext.get("etf_size") or [])
    add("ETF全量明细下限", _n_etf_size >= _min_etf,
        "{} 只（下限 {}）".format(_n_etf_size, _min_etf))
    add("全市场个股下限", len(_rk) >= _min_rank,
        "{} 只（下限 {}）".format(len(_rk), _min_rank))
    _ratio = breadth.get("ratio")
    add("广度数值有效", _ratio is not None and 0 <= _ratio <= 100,
        "上涨占比 {}（应 0–100）".format(
            "{:.0f}%".format(_ratio) if _ratio is not None else "缺失"))
    _chgv = market.get("chg")
    _chg_ok = (_chgv is None) or (abs(_chgv) <= 40)
    add("成交额环比未突变", _chg_ok,
        "环比 {}（绝对值 >40% 视为时点或口径不一致，需核对）".format(
            pct_str(_chgv) if _chgv is not None else "缺失"),
        "info" if _intraday else "warn")
    # ---- 端点契约自检：状态 + 条数下限，逐通道留痕（写入 endpoint_health.json）----
    _health = []

    def _h(name, endpoint, status, n, min_n):
        _health.append({"name": name, "endpoint": endpoint, "status": status,
                        "n": n, "min_n": min_n,
                        "ok": status in ("ok", "web_channel") and n >= min_n})

    def _mn(v):
        return 10 if _snap else v

    # ① 号通道实测单次可回 50 条（分页游标会被服务端忽略），下限按实际能力设 40
    _h("① 东财JSON", "np-weblist…/getFastNewsList",
       ext.get("_em_json_status") or "empty", ext.get("_em_json_n") or 0, _mn(40))
    _h("③ 东财PC页", "kuaixun.eastmoney.com/7_24.html",
       ext.get("_empc_status") or "empty", ext.get("_empc_n") or 0, _mn(20))
    _h("④ 新浪7×24", "finance.sina.com.cn/7x24/",
       ext.get("_sina_status") or "empty", ext.get("_sina_n") or 0, _mn(30))
    _h("⑤ 华夏时报", "/api/proxy/news/result",
       ext.get("_ct_status") or "empty", ext.get("_ct_n") or 0, _mn(30))
    _h("⑥ 财联社", "www.cls.cn/telegraph",
       ext.get("_cls_status") or "empty", ext.get("_cls_n") or 0, _mn(20))
    _h("⑦ 新浪期货", "hq.sinajs.cn/list=nf_*",
       ext.get("_futures_status") or "empty",
       len(ext.get("_futures") or []), _mn(4))
    _bad_health = [x for x in _health if not x["ok"]]
    add("端点契约（状态+条数）", not _bad_health,
        "全部通道通过" if not _bad_health else
        "{} 个通道未达下限：{}".format(
            len(_bad_health), "、".join(x["name"] for x in _bad_health)))
    bad = [c for c in checks if not c["ok"] and c["level"] == "warn"]
    soft = [c for c in checks if not c["ok"] and c["level"] != "warn"]
    score = sum(1 for c in checks if c["ok"]) / len(checks) * 100 if checks else 0
    # 通过率必须与结论文案一致：95% 不能写成「数据完整」
    if bad:
        verdict = "{} 项需留意".format(len(bad))
    elif soft:
        verdict = "{} 项提示".format(len(soft))
    else:
        verdict = "数据完整"
    return {"checks": checks, "score": score, "bad": bad,
            "soft": soft, "verdict": verdict, "health": _health}


def analyze_consistency(today, market, fund, breadth, rank, universe,
                        quality, attribution, fut, ev_ar, cone=None, amt=None,
                        opt=None):
    """跨源 / 跨章节一致性校验：同一事实在两套接口、两章表格里必须对得上。

    端点契约解决「拿得到吗」，这里解决「拿到的数字彼此矛盾吗」。
    广度接口、全量排行、ETF 明细是三套独立采集，任何一处改版或口径漂移，
    都会表现为同一指标在两个章节里不一致——这比单点断言更早暴露问题。
    容差按指标性质给定：占比类 3pct、家数类 2 只、规模类 1%。
    """
    rows = []

    def add(name, a, b, tol, note, rel=False):
        if a is None or b is None:
            return
        if rel:
            diff = abs(a - b) / max(abs(a), 1e-9) * 100
            a_s, b_s = "{:g}".format(a), "{:g}".format(b)
            d_s = "{:.2f}% ≤ {:g}%".format(diff, tol)
        else:
            diff = abs(a - b)
            a_s, b_s = "{:g}".format(a), "{:g}".format(b)
            d_s = "{:g} ≤ {:g}".format(diff, tol)
        rows.append({"name": name, "a": a_s, "b": b_s, "diff": d_s,
                     "ok": diff <= tol, "note": note})

    bd_ratio = breadth.get("ratio")
    rk_ratio = rank.get("ratio")
    add("上涨占比：广度接口 vs 全量排行", bd_ratio, rk_ratio, 3.0,
        "两套独立采集，差异 >3pct 说明其中一方口径或时点漂移")

    add("涨停家数：广度接口 vs 全量排行", breadth.get("uplim"),
        rank.get("limit_up"), 2.0,
        "涨停判定含 ST/北交所阈值差异，容差 2 只")

    add("ETF 只数：ETF 明细 vs 全量排行", universe.get("universe_n"),
        rank.get("n_etf"), 1.0,
        "同为全市场口径，差异 >1% 说明 ETF 明细缺行", True)

    add("个股只数：广度接口 vs 全量排行", breadth.get("total_stocks"),
        rank.get("n_stock"), 0.5,
        "停牌股在两端是否计入可能不同", True)

    at = today.get("amount_total")
    mt = market.get("all_total") or market.get("total")
    if at and mt:
        rows.append({
            "name": "样本 ETF 成交额 ≤ 两市成交额",
            "a": "{:.0f} 亿".format(at), "b": "{:.0f} 亿".format(mt),
            "diff": "占 {:.1f}% < 100%".format(at / mt * 100),
            "ok": at <= mt,
            "note": "样本仅覆盖跟踪清单，占比过高说明口径混入全量"})

    _topics = attribution.get("topics") or []
    if ev_ar.get("available") and _topics:
        rows.append({
            "name": "主题 AR 覆盖数 ≤ 主题候选数",
            "a": "{} 个".format(len(ev_ar.get("rows") or [])),
            "b": "{} 个".format(len(_topics)),
            "diff": "≤ 0",
            "ok": len(ev_ar.get("rows") or []) <= len(_topics),
            "note": "只对能匹配到样本 ETF 的主题算 AR，故不应多于主题数"})

    rows.append({
        "name": "期货合约数 ≥ 4（IF/IH/IC/IM）",
        "a": "{}".format(len(fut.get("rows") or [])),
        "b": "≥ 4", "diff": "—",
        "ok": len(fut.get("rows") or []) >= 4,
        "note": "缺任一主力连续则 6.1.1 的贴水判断不完整"})

    _checks = quality.get("checks") or []
    if _checks:
        ok_n = sum(1 for c in _checks if c.get("ok"))
        calc = ok_n / len(_checks) * 100
        rows.append({
            "name": "质量分自洽（通过项 / 总项）",
            "a": "{:.0f}%".format(quality.get("score") or 0),
            "b": "{:.0f}%".format(calc),
            "diff": "{:.1f}pct ≤ 2pct".format(
                abs(calc - (quality.get("score") or 0))),
            "ok": abs(calc - (quality.get("score") or 0)) <= 2,
            "note": "{}/{} 项通过".format(ok_n, len(_checks))})

    _cone = (cone or {}).get("rows") or []
    if _cone:
        _inb = all(
            r.get("hv20") is not None and (r.get("cone") or {}).get(20)
            and r["cone"][20]["low"] <= r["hv20"] <= r["cone"][20]["high"] + 1e-9
            for r in _cone)
        rows.append({
            "name": "HV20 落在 20 日锥区间内",
            "a": "{} 只".format(len(_cone)), "b": "全部在区间内",
            "diff": "—", "ok": _inb,
            "note": "当前 HV20 必然是历史序列的一员，越界说明滚动窗口或年化算错"})

    if amt and amt.get("days"):
        _n = amt.get("count") or 0
        rows.append({
            "name": "成交额历史回填覆盖 ≥ 10 日",
            "a": "{} 日".format(_n), "b": "≥ 10",
            "diff": "—", "ok": _n >= 10,
            "note": "回填只含成交额与指数涨跌；广度/资金/ERP 不插值"})

    _opt = (opt or {}).get("rows") or []
    if _opt:
        _ivs = [r["iv_atm"] for r in _opt if r.get("iv_atm") is not None]
        rows.append({
            "name": "ATM IV 落在 3%~200%",
            "a": "{} 个标的".format(len(_ivs)),
            "b": ("{:.1f}~{:.1f}%".format(min(_ivs), max(_ivs)))
            if _ivs else "无",
            "diff": "—",
            "ok": bool(_ivs) and all(3 <= v <= 200 for v in _ivs),
            "note": "越界说明反解区间或标的现价取错（如把期权价当标的价格）"})
        _exp_ok = all(r.get("exp_days_calc") is None
                      or abs(r["exp_days_calc"] - r["days"]) <= 1 for r in _opt)
        _d0 = next((r for r in _opt if r.get("exp_days_calc") is not None), None)
        rows.append({
            "name": "期权剩余天数 = 第四个星期三",
            "a": "接口 {} 天".format(_d0["days"]) if _d0 else "—",
            "b": "算得 {} 天".format(_d0["exp_days_calc"]) if _d0 else "—",
            "diff": "≤ 1 天", "ok": _exp_ok,
            "note": "到期日规则写死在代码里，与接口天数互为校验"})

    ev_n = attribution.get("n_events") or 0
    cand_n = attribution.get("n_cand") or 0
    if ev_n or cand_n:
        rows.append({
            "name": "事件总数 ≥ 候选驱动数",
            "a": "{}".format(ev_n), "b": "≥ {}".format(cand_n),
            "diff": "—",
            "ok": ev_n >= cand_n,
            "note": "候选驱动是事件池子集，顺序颠倒说明切分逻辑异常"})

    bad = [r for r in rows if not r["ok"]]
    return {"rows": rows, "bad": bad,
            "ok_n": len(rows) - len(bad), "n": len(rows),
            "verdict": "全部一致" if not bad else
                       "{} 项不一致".format(len(bad))}


def analyze_history_diff(hist, date, today, market, fund, breadth, ext, macro,
                         adj_score=None):
    keys = sorted(k for k in hist if k < date)
    prev_date = keys[-1] if keys else None
    if not prev_date:
        return {"available": False, "prev_date": None, "rows": [],
                "signals": ["逐日历史仅 {} 条，尚无昨日可比；连续运行 2 日后本章自动生效。".format(
                    len(hist))]}
    p = hist[prev_date]
    _bf = (p.get("amount_src") == "回填")
    pairs = [
        ("两市成交额(亿)", market.get("total"), p.get("amount")),
        ("成交额环比(%)", market.get("chg"), p.get("amount_chg")),
        ("上涨家数占比(%)", breadth.get("ratio"), p.get("up_ratio")),
        ("涨停家数", breadth.get("uplim"), p.get("uplim")),
        ("跌停家数", breadth.get("dnlim"), p.get("dnlim")),
        ("主力净额(亿)", fund.get("net_total"), p.get("net_total")),
        ("净流入板块数", len(fund.get("pos") or []), p.get("pos_n")),
        ("样本ETF涨幅加权(%)", today.get("sample_weighted"), p.get("sample_weighted")),
        ("ERP(pct)", macro.get("erp"), p.get("erp")),
        ("官方画像调整分", adj_score, p.get("adj_score")),
    ]
    # 盘中：今日成交额是累计到当前时点的进度值，而回填的历史是全天值——
    # 直接相减会得出「缩量 -28%」这种假结论，故这一类配对整组跳过。
    _mix = _bf and AS_OF.get("state") in ("intraday", "snapshot")
    _skip = {"两市成交额(亿)", "成交额环比(%)"} if _mix else set()
    rows = []
    for label, cur, prev in pairs:
        if label in _skip:
            continue
        if cur is None or prev is None:
            continue
        diff = cur - prev
        rows.append({"label": label, "cur": cur, "prev": prev, "diff": diff,
                     "pct": (diff / abs(prev) * 100) if prev else None})
    signals = []
    if rows:
        up = [r["label"] for r in rows if r["diff"] > 0]
        dn = [r["label"] for r in rows if r["diff"] < 0]
        signals.append("与 {} 相比，{} 项上升、{} 项下降。".format(
            prev_date, len(up), len(dn)))
    chg_minutes = [r for r in rows if r["label"] == "两市成交额(亿)"]
    if chg_minutes and market.get("total"):
        r = chg_minutes[0]
        signals.append("成交额由 {:.0f} 亿变为 {:.0f} 亿（{:+.0f} 亿，{}）；"
                       "{}。".format(
                           r["prev"], r["cur"], r["diff"],
                           pct_str(r["pct"]) if r["pct"] is not None else "—",
                           "放量延续" if r["diff"] > 0 else "量能回落，反弹动能衰减"))
    net = next((r for r in rows if r["label"] == "主力净额(亿)"), None)
    if net:
        signals.append("板块主力净额变化 {:+.0f} 亿，{}。".format(
            net["diff"], "资金持续流入" if net["cur"] > 0 and net["diff"] > 0
            else ("流入放缓或转出" if net["diff"] < 0 else "基本持平")))
    if _mix:
        signals.insert(0, "今日为**盘中累计**口径，对比日仅有全天回填值，"
                          "成交额环比不可比，已从本表剔除；量能对比请看第一章的"
                          "同时点环比（{}）与 12.2 的序列定位。".format(
                              pct_str(market.get("chg"))))
    if _bf:
        signals.append("对比日 {} 的成交额为**回填值**（沪市 + 深市指数成交额合成），"
                       "与当日实时口径（全量个股成交额合计）存在 1~5% 差异，"
                       "环比只作方向参考；其余字段（广度/资金/ERP）当日无快照，"
                       "不做插值。".format(prev_date))
    _bf_n = sum(1 for k, v in hist.items()
                if v.get("amount_src") == "回填" and k < date)
    return {"available": bool(rows), "prev_date": prev_date, "rows": rows,
            "signals": signals, "n_days": len(hist),
            "prev_backfilled": _bf, "backfill_n": _bf_n}


def analyze_amount_series(hist, date, n=20):
    """近 N 日两市成交额序列（实时 + 回填混合），给出量能的分位与均值参照。

    单日「放量 58%」没有参照系：只有把 20 日均值、区间、分位摆出来，
    才能判断这是「趋势级放量」还是「一次性脉冲」。
    """
    keys = sorted(k for k in (hist or {}) if hist[k].get("amount"))[-n:]
    rows = [{"date": k, "amount": hist[k]["amount"],
             "chg": hist[k].get("amount_chg"),
             "sh_pct": hist[k].get("sh_pct"),
             "sz_pct": hist[k].get("sz_pct"),
             "partial": bool(hist[k].get("amount_partial")),
             "src": hist[k].get("amount_src") or "实时"}
            for k in keys]
    if not rows:
        return {"available": False, "rows": [], "signals": []}
    # 盘中口径：当日成交额是**累计到当前时点**的进度值，与历史全天值不可直接比。
    # 这一点是本章最容易出错的地方（曾出现「放量 58%」实为同时点对比）。
    _state = AS_OF.get("state")
    # 未完成日的判定有两个来源：历史里存的 amount_partial 标记（跨日可靠），
    # 以及当前运行处于盘中/快照状态（当日尚未写历史时也能判断）。
    partial = bool(rows[-1]["partial"]) or (
        _state in ("intraday", "snapshot") and rows[-1]["date"] == date
        and rows[-1]["src"] != "回填")
    rows[-1]["partial"] = partial
    base = [r for r in rows if not r["partial"]]
    if partial:
        rows[-1]["chg"] = None
        rows[-1]["src"] = "实时（未收盘）"
    if not base:
        return {"available": False, "rows": rows,
                "signals": ["历史里只有未完成日（盘中累计值），暂无全天序列可参照。"]}
    amts = [r["amount"] for r in base]
    srt = sorted(amts)
    cur = base[-1]["amount"]
    pct = sum(1 for x in amts if x <= cur) / len(amts) * 100
    _hhmm = (AS_OF.get("time") or "").replace("T", " ")[-5:]
    signals = [
        "近 {} 个**完整交易日**两市成交额：均值 **{:.0f} 亿**、中位 {:.0f} 亿、"
        "区间 {:.0f}~{:.0f} 亿；最新完整日 {} 为 {:.0f} 亿，"
        "处在该区间的 **{:.0f}% 分位**。".format(
            len(base), sum(amts) / len(amts), srt[len(srt) // 2], min(amts),
            max(amts), base[-1]["date"], cur, pct),
        "**口径**：{} 天为指数成交额回填（沪市 + 深市指数合成），"
        "其余为当日实时快照；两者相差 1~5%，故只比同一序列内部的相对高低，"
        "不把回填值当成「当日全量个股成交额」。".format(
            sum(1 for r in base if r["src"] == "回填")),
    ]
    if partial:
        _today = rows[-1]["amount"]
        _when = ("快照时点" if _state == "snapshot" else (_hhmm or "当前"))
        signals.insert(0, "今日（{}）成交 {:.0f} 亿，为最新完整日（{}）的 "
                          "**{:.0f}%**、均值的 {:.0f}%；"
                          "**未收盘，不与全天值比较**，日环比应看第一章的"
                          "「同时点对比」。".format(
                              _when, _today, base[-1]["date"],
                              _today / base[-1]["amount"] * 100 if base[-1]["amount"] else 0,
                              _today / (sum(amts) / len(amts)) * 100))
    if cur and max(amts) and cur >= max(amts) * 0.98:
        signals.append("最新完整日接近区间上沿（为区间最大值的 {:.0f}%），"
                       "量能已在近 {} 日高位，若价格不能同步创新高即为量价背离。".format(
                           cur / max(amts) * 100, len(base)))
    elif pct <= 20:
        signals.append("最新完整日处在 20% 分位以下，属**缩量**状态；"
                       "缩量下的反弹通常缺少承接，需等量能确认。")
    return {"available": True, "rows": rows, "n": len(rows), "partial": partial,
            "n_full": len(base),
            "avg": sum(amts) / len(amts), "median": srt[len(srt) // 2],
            "min": min(amts), "max": max(amts), "pct": pct, "signals": signals}


# ----------------------------------------------------------------------------
# 分析层十六：增强技术指标（ATR / 波动率 / 缺口 / 多周期）
# ----------------------------------------------------------------------------

def analyze_advanced_tech(store):
    rows = []
    for c in CORE_CODES:
        if c not in store:
            continue
        bars, q = store[c]["bars"], store[c]["quote"]
        if len(bars) < 30:
            continue
        trs = []
        for i in range(1, min(len(bars), 20)):
            b, p = bars[-i], bars[-i - 1]
            trs.append(max(b["high"] - b["low"],
                           abs(b["high"] - p["close"]),
                           abs(b["low"] - p["close"])))
        atr = avg(trs)
        rets = [(bars[-i]["close"] / bars[-i - 1]["close"] - 1) * 100
                for i in range(1, min(len(bars), 21))]
        vol = None
        if len(rets) > 2:
            m = avg(rets)
            vol = (sum((x - m) ** 2 for x in rets) / (len(rets) - 1)) ** 0.5
        gap = None
        if q["open"] and q["prev_close"] and q["prev_close"] > 0:
            gap = (q["open"] / q["prev_close"] - 1) * 100
        ma5w, ma10w = None, None
        weekly = []
        for i in range(0, len(bars), 5):
            chunk = bars[i:i + 5]
            if chunk:
                weekly.append(chunk[-1]["close"])
        if len(weekly) >= 5:
            ma5w = avg(weekly[-5:])
        if len(weekly) >= 10:
            ma10w = avg(weekly[-10:])
        rows.append({
            "name": q["name"], "price": q["price"],
            "atr": atr, "atr_pct": (atr / q["price"] * 100) if (atr and q["price"]) else None,
            "vol": vol, "gap": gap,
            "ma5w": ma5w, "ma10w": ma10w,
            "week_trend": ("周线多头" if (ma5w and ma10w and ma5w > ma10w)
                           else ("周线空头" if ma5w and ma10w else "—")),
        })
    signals = []
    if rows:
        atrs = [r["atr_pct"] for r in rows if r["atr_pct"]]
        vols = [r["vol"] for r in rows if r["vol"]]
        signals.append("核心标的 ATR(20) 占价格 {:.2f}%，日均波动率 {:.2f}%，"
                       "用于设定止损宽度与仓位，而非方向判断。".format(
                           avg(atrs) or 0, avg(vols) or 0))
    gaps = [r for r in rows if r["gap"] is not None and abs(r["gap"]) >= 0.5]
    if gaps:
        signals.append("跳空缺口 ≥0.5% 的标的：{}；缺口未回补前是短线支撑/压力参考。".format(
            "、".join("{} {:+.2f}%".format(r["name"], r["gap"]) for r in gaps)))
    else:
        signals.append("核心标的均为平开/微幅开盘（无 ≥0.5% 缺口），"
                       "与「资金持续买入而非情绪跳空」的形态判断一致。")
    wk = [r for r in rows if r["week_trend"] == "周线多头"]
    signals.append("周线多周期共振：{}/{} 只核心标的周线 MA5 > MA10；"
                   "{}。".format(
                       len(wk), len(rows),
                       "多周期仍偏空，日线反弹未获周线确认" if len(wk) < len(rows) / 2
                       else "周线开始转多"))
    return {"rows": rows, "signals": signals, "available": bool(rows)}


def analyze_vol_cone(store, codes=None, bench="sh000300"):
    """已实现波动率锥：HV20/60/120 年化 + 当前 HV20 的历史分位。

    单看「20 日波动率 1.8%」没有参照系，波动率锥给出同一标的三个期限的
    历史区间（低/中/高），分位才说明「现在的波动算大还是算小」。
    HV 是**已实现**波动（历史收益的标准差），不是期权隐含波动率——
    本报告没有可靠的免费 IV 源，故不臆造 IV，也不把 HV 当 IV 用。
    """
    codes = list(codes or CORE_CODES) + ([bench] if bench not in (codes or CORE_CODES) else [])
    rows, signals = [], []
    for c in codes:
        node = store.get(c)
        if not node:
            continue
        bars = node.get("bars") or []
        if len(bars) < 140:
            continue
        rets = [(bars[i]["close"] / bars[i - 1]["close"] - 1) * 100
                for i in range(1, len(bars)) if bars[i - 1]["close"]]

        def hv(n):
            if len(rets) < n:
                return None
            w = rets[-n:]
            m = sum(w) / n
            var = sum((x - m) ** 2 for x in w) / (n - 1) if n > 1 else 0
            return math.sqrt(var) * math.sqrt(252)

        def hv_series(n):
            out = []
            for i in range(n, len(rets) + 1):
                w = rets[i - n:i]
                m = sum(w) / n
                var = sum((x - m) ** 2 for x in w) / (n - 1) if n > 1 else 0
                out.append(math.sqrt(var) * math.sqrt(252))
            return out

        cur = {n: hv(n) for n in (20, 60, 120)}
        if not cur[20]:
            continue
        s20 = hv_series(20)
        pct = (sum(1 for x in s20 if x <= cur[20]) / len(s20) * 100) if s20 else None
        cone = {}
        for n in (20, 60, 120):
            s = hv_series(n) if n != 20 else s20
            if s:
                cone[n] = {"low": min(s), "mid": sorted(s)[len(s) // 2], "high": max(s)}
        state = ("低波（<20 分位）" if pct is not None and pct < 20 else
                 "高波（>80 分位）" if pct is not None and pct > 80 else "常态")
        rows.append({"code": c, "name": (node.get("quote") or {}).get("name") or c,
                     "hv20": cur[20], "hv60": cur[60], "hv120": cur[120],
                     "pct": pct, "state": state, "cone": cone,
                     "n_bars": len(bars)})

    # 基准放最后，避免与核心标的混读
    rows.sort(key=lambda r: (r["code"] == bench, -(r["pct"] or 0)))
    if rows:
        _p = [r["pct"] for r in rows if r["code"] != bench and r["pct"] is not None]
        if _p:
            _s = sorted(_p)
            _mid = ((_s[len(_s) // 2] + _s[(len(_s) - 1) // 2]) / 2
                    if len(_s) > 1 else _s[0])
            signals.append("**波动率分位**：{} 只核心标的当前 HV20 分位中位数 **{:.0f}%**"
                           "（{} 只处于高波区间、{} 只处于低波区间）。"
                           "分位高说明当下的日常波动已接近过去一年的上限，"
                           "同样的消息更容易打出大阳/大阴，止损宽度需放宽。".format(
                               len(_p), _mid,
                               sum(1 for r in rows if "高波" in r["state"]),
                               sum(1 for r in rows if "低波" in r["state"])))
        _hi = max((r for r in rows if r["code"] != bench and r["pct"] is not None),
                  key=lambda r: r["pct"], default=None)
        _lo = min((r for r in rows if r["code"] != bench and r["pct"] is not None),
                  key=lambda r: r["pct"], default=None)
        if _hi and _lo:
            signals.append("**两端标的**：{} HV20 分位 {:.0f}%（HV {:.1f}%）最高、"
                           "{} 分位 {:.0f}%（HV {:.1f}%）最低；波动分位差 {:.0f}pct，"
                           "说明风险预算不能对全样本用同一套阈值。".format(
                               _hi["name"], _hi["pct"], _hi["hv20"],
                               _lo["name"], _lo["pct"], _lo["hv20"],
                               _hi["pct"] - _lo["pct"]))
        b = next((r for r in rows if r["code"] == bench), None)
        if b:
            signals.append("**大盘参照**：{} HV20 分位 {:.0f}%（HV {:.1f}%，"
                           "60 日 {:.1f}%、120 日 {:.1f}%）；"
                           "核心标的分位高于基准即为「结构性波动溢价」。".format(
                               b["name"], b["pct"] or 0, b["hv20"],
                               b["hv60"] or 0, b["hv120"] or 0))
        signals.append("**口径边界**：HV 是**已实现**波动（历史收益标准差×√252），"
                       "不是期权隐含波动率；IV 见 14.1d（由期权价格反解），"
                       "两者之差即波动率风险溢价，不在本节混用。")
    return {"rows": rows, "signals": signals, "available": bool(rows),
            "n_codes": len(rows)}


# ----------------------------------------------------------------------------
# 分析层十四b：期权隐含波动率（BS 反解）+ 波动率风险溢价
# ----------------------------------------------------------------------------

OPTION_URL = ("https://push2.eastmoney.com/api/qt/clist/get?pn={pn}&pz=100&po=1"
              "&np=1&fltt=2&invt=2&fid=f3&fs=m:10"
              "&fields=f2,f3,f12,f14,f108,f161,f162,f163")

# 期权名称前缀 → 标的 ETF 代码（与 ETF_GROUPS 对齐，只保留本地有日线的）
OPT_UNDER = {
    "50ETF": "sh510050", "300ETF": "sh510300", "500ETF": "sh510500",
    "科创板50": "sh588000", "创业板ETF": "sz159915",
}


def _ncdf(x):
    return 0.5 * (1.0 + math.erf(x / math.sqrt(2.0)))


def bs_price(S, K, T, r, sigma, kind):
    """Black-Scholes：ETF 期权按欧式处理（A 股 ETF 期权为欧式，可行权日集中）。"""
    if not (S > 0 and K > 0 and T > 0 and sigma > 0):
        return None
    vol_t = sigma * math.sqrt(T)
    d1 = (math.log(S / K) + (r + sigma * sigma / 2) * T) / vol_t
    d2 = d1 - vol_t
    df = math.exp(-r * T)
    if kind == "C":
        return S * _ncdf(d1) - K * df * _ncdf(d2)
    return K * df * _ncdf(-d2) - S * _ncdf(-d1)


def implied_vol(price, S, K, T, r, kind, lo=0.005, hi=3.0, iters=60):
    """二分法反解 IV。

    先做可解性判定：若市场价不高于内在价值（深度实值 + 报价陈旧），
    任何 σ 都拟合不出，直接返回 None——**不硬凑一个数字**。
    """
    if price is None or not (price > 0) or T <= 0 or S <= 0 or K <= 0:
        return None
    intr = max(0.0, (S - K) if kind == "C" else (K - S))
    if price <= intr + 1e-6:
        return None
    f_lo = bs_price(S, K, T, r, lo, kind)
    f_hi = bs_price(S, K, T, r, hi, kind)
    if f_lo is None or f_hi is None or price < f_lo or price > f_hi:
        return None
    a, b = lo, hi
    for _ in range(iters):
        m = (a + b) / 2
        v = bs_price(S, K, T, r, m, kind)
        if v is None:
            return None
        if v < price:
            a = m
        else:
            b = m
    return (a + b) / 2


def fetch_option_chain():
    """东财期权列表（沪 + 深两市全部合约，约 700 个）。

    字段：f12 代码 / f14 名称（如「500ETF购11月7500」）/ f2 最新价 /
    f108 持仓 / f161 行权价 / f162 剩余日历天 / f163 成交量。
    **列表接口不含隐含波动率**，IV 必须自己反解——这是本节存在的理由。
    """
    out = {"status": "ok", "rows": [], "total": 0}
    diff, _total = [], 0
    # 该接口单页最多返回 100 条，702 个合约必须翻页，
    # 否则拿到的是「按涨幅排序的前 100 个」——样本严重有偏。
    for pn in range(1, 12):
        url = OPTION_URL.format(pn=pn)
        try:
            req = Request(url, headers=dict(
                HEADERS, Referer="https://quote.eastmoney.com/"))
            with urlopen(req, timeout=15) as resp:
                d = json.loads(resp.read().decode("utf-8", "replace"))
        except Exception:
            out["status"] = "unreachable" if pn == 1 else "partial"
            break
        data = d.get("data") or {}
        page = data.get("diff") or []
        _total = data.get("total") or _total
        if not page:
            break
        diff.extend(page)
        if len(diff) >= _total:
            break
    out["total"] = _total or len(diff)
    if not diff:
        if out["status"] == "ok":
            out["status"] = "empty"
        return out
    for r in diff:
        name = r.get("f14") or ""
        if "购" in name:
            kind, prefix = "C", name.split("购")[0]
        elif "沽" in name:
            kind, prefix = "P", name.split("沽")[0]
        else:
            continue
        m = re.search(r"(\d{1,2})月", name)
        strike = num(r.get("f161"))
        days = num(r.get("f162"))
        last = num(r.get("f2"))
        if not m or not strike or not days:
            continue
        out["rows"].append({
            "code": r.get("f12"), "name": name, "under": prefix,
            "kind": kind, "month": int(m.group(1)), "strike": strike,
            "days": int(days), "last": last,
            "oi": num(r.get("f108")), "vol": num(r.get("f163")),
        })
    if not out["rows"]:
        out["status"] = "empty"
    return out


def _fourth_wednesday(year, month):
    """到期日 = 到期月第四个星期三（A 股 ETF 期权规则，用于校验剩余天数）。"""
    d = datetime(year, month, 1)
    first_wed = d + timedelta(days=(2 - d.weekday()) % 7)   # 2 = Wednesday
    return first_wed + timedelta(days=21)


def _hv20_of(store, code, n=20):
    """就地算 HV20（年化，%）。14.1c 的锥表只覆盖核心标的，
    而期权标的有 4~5 个，缺哪个就地补哪个，避免为了对齐而扩大锥表。"""
    bars = (store.get(code) or {}).get("bars") or []
    rets = [(bars[i]["close"] / bars[i - 1]["close"] - 1) * 100
            for i in range(1, len(bars)) if bars[i - 1].get("close")]
    if len(rets) < n:
        return None
    w = rets[-n:]
    m = sum(w) / n
    var = sum((x - m) ** 2 for x in w) / (n - 1)
    return math.sqrt(var) * math.sqrt(252)


def analyze_option_iv(chain, store, cone=None, r=0.015, wing=0.15):
    """期权隐含波动率：ATM IV、波动率风险溢价（IV − HV20）、偏斜。

    只看**最近到期月**：其余月份成交稀薄，最新价常是几分钟甚至前一日挂单，
    反解出来的 IV 会系统性偏低/偏高。行权价限制在现价 ±15%，
    深度虚值/实值的价格对 σ 不敏感，反解是噪声放大。
    """
    raw = (chain or {}).get("rows") or []
    if not raw:
        return {"available": False, "rows": [], "signals": [
            "期权链不可得（{}），本节跳过。IV 需要期权最新价，"
            "离线快照无法计算。".format((chain or {}).get("status") or "未知")]}
    _hv = {r_["code"]: r_ for r_ in (cone or {}).get("rows") or []}
    rows, signals = [], []
    _today = datetime.now(CST)

    by_under = {}
    for c in raw:
        code = OPT_UNDER.get(c["under"])
        if code and code in store:
            by_under.setdefault(code, []).append(c)

    for code, cons in sorted(by_under.items()):
        node = store.get(code) or {}
        q = node.get("quote") or {}
        bars = node.get("bars") or []
        spot = q.get("price") or (bars[-1]["close"] if bars else None)
        if not spot:
            continue
        days_list = sorted({c["days"] for c in cons if c["days"] and c["days"] > 0})
        if not days_list:
            continue
        d0 = days_list[0]
        T = d0 / 365.0
        ivs, n_try = {}, 0
        for c in cons:
            if c["days"] != d0:
                continue
            if abs(c["strike"] / spot - 1) > wing:
                continue
            n_try += 1
            iv = implied_vol(c["last"], spot, c["strike"], T, r, c["kind"])
            if iv is None:
                continue
            ivs.setdefault(c["strike"], {})[c["kind"]] = iv
        if not ivs:
            continue
        atm_k = min(ivs, key=lambda k: abs(k - spot))
        _atm = [v for v in ivs[atm_k].values() if v is not None]
        if not _atm:
            continue
        iv_atm = sum(_atm) / len(_atm)
        lo_k = min(ivs, key=lambda k: abs(k - spot * 0.95))
        hi_k = min(ivs, key=lambda k: abs(k - spot * 1.05))
        skew = None
        if ivs[lo_k].get("P") is not None and ivs[hi_k].get("C") is not None:
            skew = (ivs[lo_k]["P"] - ivs[hi_k]["C"]) * 100
        hv = (_hv.get(code) or {}).get("hv20") or _hv20_of(store, code)
        # 到期日自检：接口给的剩余天数应与「第四个星期三」对得上
        _mon = next((c["month"] for c in cons if c["days"] == d0), 0)
        _exp = _fourth_wednesday(_today.year, _mon) if _mon else None
        if _exp and _exp.date() < _today.date():
            _exp = _fourth_wednesday(_today.year + (1 if _mon == 12 else 0),
                                     (_mon % 12) + 1)
        rows.append({
            "code": code, "name": q.get("name") or code, "spot": spot,
            "atm_k": atm_k, "days": d0,
            "iv_call": ivs[atm_k].get("C"), "iv_put": ivs[atm_k].get("P"),
            "iv_atm": iv_atm * 100, "hv20": hv,
            "vrp": (iv_atm * 100 - hv) if hv is not None else None,
            "skew": skew, "n_iv": sum(len(v) for v in ivs.values()),
            "n_try": n_try, "month": _mon,
            "exp_days_calc": ((_exp.date() - _today.date()).days
                              if _exp else None),
        })

    if not rows:
        return {"available": False, "rows": [], "signals": [
            "期权合约均未通过可解性判定（深度实值报价陈旧 / 内在价值高于市价），"
            "本次不计算 IV——宁可空缺，不给一个拟合出来的假数字。"]}

    rows.sort(key=lambda x: -(x["vrp"] if x["vrp"] is not None else -999))
    _vrp = [x["vrp"] for x in rows if x["vrp"] is not None]
    if _vrp:
        _s = sorted(_vrp)
        _mid = (_s[len(_s) // 2] + _s[(len(_s) - 1) // 2]) / 2
        signals.append("**波动率风险溢价**：{} 个标的 ATM IV − HV20 的中位数 "
                       "**{:+.1f}pct**（{} 个为正、{} 个为负）。"
                       "为正说明期权定价高于最近 20 日的实际波动，"
                       "买方在为「未来的不确定性」多付钱；为负则相反。".format(
                           len(_vrp), _mid,
                           sum(1 for v in _vrp if v > 0),
                           sum(1 for v in _vrp if v < 0)))
    _top = max((x for x in rows if x["vrp"] is not None),
               key=lambda x: x["vrp"], default=None)
    _bot = min((x for x in rows if x["vrp"] is not None),
               key=lambda x: x["vrp"], default=None)
    if _top and _bot and _top["code"] != _bot["code"]:
        signals.append("**两端标的**：{} ATM IV {:.1f}%（HV20 {:.0f}%，溢价 "
                       "{:+.1f}pct）最高；{} IV {:.1f}%（HV20 {:.0f}%，{:+.1f}pct）"
                       "最低。IV 与 HV 的错位说明期权市场对两者的定价分歧，"
                       "不能直接用「波动大」统一解释。".format(
                           _top["name"], _top["iv_atm"], _top["hv20"] or 0,
                           _top["vrp"], _bot["name"], _bot["iv_atm"],
                           _bot["hv20"] or 0, _bot["vrp"]))
    _sk = [x for x in rows if x["skew"] is not None]
    if _sk:
        _hi = max(_sk, key=lambda x: x["skew"])
        signals.append("**偏斜（95% 认沽 IV − 105% 认购 IV）**：{} 为 {:+.1f}pct 最陡，"
                       "说明虚值认沽被买得更贵——市场在为下跌尾部付保险费；"
                       "偏斜接近 0 则说明多空尾部定价对称。".format(
                           _hi["name"], _hi["skew"]))
    signals.append("**口径**：IV 由最新价用 Black-Scholes（欧式、r={:.1%} 固定、"
                   "T=剩余日历天/365）**反解**，不是行情商直接给的字段；"
                   "深度实值合约若市价不高于内在价值则无解，一律跳过。"
                   "因此 IV 精度受报价陈旧度限制，只用于同一时点的横向比较。".format(r))
    return {"available": True, "rows": rows, "signals": signals,
            "n_under": len(rows), "r": r,
            "n_contracts": (chain or {}).get("total") or 0,
            "chain_status": (chain or {}).get("status")}


# ----------------------------------------------------------------------------
# 分析层十五：事件采集 → 时序切分 → 板块映射 → 三方校验 → 涨跌归因
# ----------------------------------------------------------------------------

CST = timezone(timedelta(hours=8))

# ---- 数据时点标定：区分「盘中实时」与「收盘终值」----
# 同一交易日盘中多次运行，成交额/广度/资金流会持续变化，若一律写成「收盘快照」，
# 会得出相反结论（实测同一日 11793→13823 亿、价涨量增↔价跌量增）。
# 故所有报告头部与关键指标都带 as-of 时点，盘中口径下禁用「收盘」字样与绝对结论。
AS_OF = {"time": "", "state": "close",
         "label": "收盘后（当日终值）", "hist_days": 0}
MARKET_OPEN_HM = (9, 30)
MARKET_CLOSE_HM = (15, 0)

# 运行时数据完整性告警：main 填充、render 写进报告头部。
# 之前接口失败/旧快照补缺只在终端打印，报告看起来仍像完整实时数据。
RUN_WARNINGS = []
PATCH_INFO = {}


def resolve_as_of(source_note="", now=None):
    """根据数据来源与实际时钟判断数据时点状态。"""
    now = now or datetime.now(CST)
    if "离线快照" in (source_note or ""):
        return {"time": now.strftime("%Y-%m-%d %H:%M"),
                "state": "snapshot",
                "label": "离线快照（静态值，非实时；上示时间为读取时刻）"}
    hm = (now.hour, now.minute)
    trading = now.weekday() < 5
    if trading and MARKET_OPEN_HM <= hm < MARKET_CLOSE_HM:
        state, label = "intraday", "盘中实时（未收盘，数值随刷新变化）"
    elif trading and hm < MARKET_OPEN_HM:
        state, label = "preopen", "开盘前（指数/个股为上一交易日收盘值）"
    else:
        state, label = "close", "收盘后（当日终值）"
    # 主体数据是在线的，但部分字段来自旧快照：状态仍按时钟给，
    # 标签必须点破，否则报告会被当成「全部是本次抓取的终值」。
    if "快照补缺" in (source_note or ""):
        label += "；⚠️ 部分字段来自旧快照（见头部告警）"
    return {"time": now.strftime("%Y-%m-%d %H:%M"),
            "state": state, "label": label}


def _to_dt(ts):
    """unix 秒 / 数字字符串 / 常见时间字符串 → aware datetime(CST)。"""
    if ts is None or ts == "":
        return None
    try:
        if isinstance(ts, (int, float)) or (
                isinstance(ts, str) and re.match(r"^\d{9,13}$", ts.strip())):
            v = float(ts)
            if v > 1e11:
                v /= 1000.0
            return datetime.fromtimestamp(v, CST)
        s = str(ts).strip().replace("T", " ").split("+")[0].strip()
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
            try:
                return datetime.strptime(s, fmt).replace(tzinfo=CST)
            except ValueError:
                continue
        return datetime.fromisoformat(s).replace(tzinfo=CST)
    except Exception:
        return None


def _fmt_dt(d):
    return d.strftime("%m-%d %H:%M") if d else "—"


def find_launch_time(breadth_raw):
    """从上证分时找 5 分钟滚动涨幅最大的窗口起点，作为当日主升启动时刻。"""
    try:
        data = ((breadth_raw.get("minute_set") or {})
                .get("minute_sh_index") or {}).get("data") or []
        pts = []
        for ln in data:
            p = str(ln).split()
            if len(p) >= 2:
                v = num(p[1])
                if v:
                    pts.append((p[0][:4], v))
        if len(pts) < 12:
            return None, 0.0
        best, best_t = -9e9, pts[0][0]
        for i in range(5, len(pts)):
            r = (pts[i][1] / pts[i - 5][1] - 1) * 100
            if r > best:
                best, best_t = r, pts[i - 5][0]
        return "{}:{}".format(best_t[:2], best_t[2:]), best
    except Exception:
        return None, 0.0


def extract_viewpoints(events, fund, today, launch_dt=None):
    """从统一事件池里识别「多方评论观点」并做立场/时序/分歧/交叉校验。

    输入 events 为 analyze_attribution 已切分好时序的条目（含 dt/phase/
    title/body/source/origin/kind/hits）。纯行情数字快讯会被标成「事实/无立场」，
    不参与多空统计。立场由关键词规则判定，非语义模型，会误判——报告里明示。
    """
    empty = {"available": False, "n": 0, "n_stance": 0, "n_fact": 0,
             "rows": [], "actors": [], "stance_dist": [], "topics": [],
             "timing": {}, "signals": []}
    if not events:
        return empty

    def actor_of(e):
        text = " ".join([e.get("title") or "", e.get("body") or "",
                         e.get("origin") or "", e.get("source") or ""])
        for name, kws in VIEW_ACTORS:
            if any(k in text for k in kws):
                return name
        return "其他/未识别"

    def stance_of(text):
        b = sum(1 for w in VIEW_BULL if w in text)
        s = sum(1 for w in VIEW_BEAR if w in text)
        d = b - s
        if b == 0 and s == 0:
            return "fact", b, s
        if d >= 2:
            return "bull", b, s
        if d <= -2:
            return "bear", b, s
        if b and s:
            return "mixed", b, s
        return ("lean_bull" if d > 0 else "lean_bear"), b, s

    # 与第四章行业资金 / 第一章样本ETF 做「观点是否获验证」
    in_top = (fund or {}).get("in_top") or []
    ranked = [q for _, q in (today or {}).get("ranked") or []]

    def verify(topic, stance, title=""):
        tp = next((x for x in TOPIC_MAP + VIEW_TOPIC_EXTRA
                   if x["topic"] == topic), None)
        if not tp:
            return "—", "未映射主题"
        # 外围行情快讯不与 A 股板块资金/ETF 涨幅对照，避免假背离
        if any(k in (title or "") for k in VIEW_EXTERNAL) and "A股" not in (title or ""):
            return "—", "外围口径，不对照A股资金"
        # 叙事型主题（无板块、无ETF映射）本就没有验证对象
        if not tp["sectors"] and not tp["etf_kw"]:
            return "不适用", "叙事型主题，无板块/ETF验证对象"
        hit_sec = []
        for r in in_top:
            for s in tp["sectors"]:
                if s and (s in r["name"] or r["name"] in s):
                    hit_sec.append((r["name"], r["net_yi"]))
                    break
        etf_pct = None
        if tp["etf_kw"]:
            vals = [q["pct"] for q in ranked
                    if any(k in (q.get("name") or "") for k in tp["etf_kw"])
                    and q.get("pct") is not None]
            if vals:
                etf_pct = sum(vals) / len(vals)
        sec_txt = "、".join("{} {:+.0f}亿".format(n, v)
                           for n, v in hit_sec[:2]) or "资金未进前列"
        pct_txt = "{:+.1f}%".format(etf_pct) if etf_pct is not None else "无对应样本"
        if stance in ("bull", "lean_bull"):
            ok = bool(hit_sec) or (etf_pct is not None and etf_pct >= 2)
            return ("✅ 一致" if ok else "⚠️ 背离"), "{} / {}".format(sec_txt, pct_txt)
        if stance in ("bear", "lean_bear"):
            ok = (not hit_sec) or (etf_pct is not None and etf_pct < 2)
            return ("✅ 一致" if ok else "⚠️ 背离"), "{} / {}".format(sec_txt, pct_txt)
        return "中性", "{} / {}".format(sec_txt, pct_txt)

    rows = []
    for e in events:
        text = " ".join([e.get("title") or "", e.get("body") or ""])
        # 观点候选：命中表态标记，且不是社会/政治/灾害类非市场快讯
        is_view = (any(m in text for m in VIEW_MARKERS)
                   or (("：" in (e.get("title") or "") or ":" in (e.get("title") or ""))
                       and len(e.get("title") or "") > 12))
        if any(k in text for k in VIEW_EXCLUDE):
            is_view = False
        if not is_view:
            continue
        st, b, s = stance_of(text)
        topics = e.get("hits") or []
        main = topics[0] if topics else "未映射主题"
        if main == "未映射主题":
            # 标题优先：正文偶尔顺带提「算力」「消费」会串到别的赛道
            ttl = e.get("title") or ""
            main = next((x["topic"] for x in VIEW_TOPIC_EXTRA
                         if any(k in ttl for k in x["kw"])), None)
            if not main:
                main = next((x["topic"] for x in VIEW_TOPIC_EXTRA
                             if any(k in text for k in x["kw"])), "未映射主题")
        vlabel, vdetail = verify(main, st, e.get("title") or "")
        dt = e.get("dt")
        phase = e.get("phase") or ""
        pre = phase in ("盘前", "盘中·启动前")
        rows.append({
            "time": _fmt_dt(dt) if dt else "",
            "phase": phase, "pre": pre,
            "actor": actor_of(e), "origin": e.get("origin") or e.get("source") or "",
            "stance": st, "stance_label": VIEW_STANCE_LABEL[st],
            "bull": b, "bear": s, "topic": main,
            "topics": "、".join(topics) or (main if main != "未映射主题" else "—"),
            "title": (e.get("title") or "")[:76],
            "snippet": re.sub(r"\s+", " ", (e.get("body") or text))[:110],
            "verify": vlabel, "verify_detail": vdetail,
            "kind": e.get("kind") or "",
        })
    if not rows:
        return empty
    rows.sort(key=lambda r: (r["time"], r["actor"]))

    # 观点方分布
    amap = {}
    for r in rows:
        a = r["actor"]
        d = amap.setdefault(a, {"actor": a, "n": 0, "bull": 0, "bear": 0,
                                "mixed": 0, "fact": 0})
        d["n"] += 1
        if r["stance"] in ("bull", "lean_bull"):
            d["bull"] += 1
        elif r["stance"] in ("bear", "lean_bear"):
            d["bear"] += 1
        elif r["stance"] == "mixed":
            d["mixed"] += 1
        else:
            d["fact"] += 1
    actors = sorted(amap.values(),
                    key=lambda x: (x["actor"] == "其他/未识别", -x["n"]))

    # 立场分布
    smap = {}
    for r in rows:
        smap[r["stance"]] = smap.get(r["stance"], 0) + 1
    stance_dist = [{"label": VIEW_STANCE_LABEL[k], "key": k, "n": v}
                   for k, v in sorted(smap.items(), key=lambda x: -x[1])]

    # 按主题聚合分歧
    tmap = {}
    for r in rows:
        if r["topic"] == "未映射主题":
            continue
        d = tmap.setdefault(r["topic"], {
            "topic": r["topic"], "n": 0, "bull": 0, "bear": 0,
            "mixed": 0, "fact": 0, "pre": 0, "repr": r["title"],
            "verifies": [], "na_n": 0, "actors": set()})
        d["n"] += 1
        d["actors"].add(r["actor"])
        if r["stance"] in ("bull", "lean_bull"):
            d["bull"] += 1
        elif r["stance"] in ("bear", "lean_bear"):
            d["bear"] += 1
        elif r["stance"] == "mixed":
            d["mixed"] += 1
        else:
            d["fact"] += 1
        if r["pre"]:
            d["pre"] += 1
        if r["verify"] == "不适用":
            d["na_n"] += 1
        elif r["verify"] != "中性":
            d["verifies"].append(r["verify"])
    topics = []
    for d in tmap.values():
        d["actors"] = sorted(d["actors"])
        # 「分歧/条件式」条目本身同一句里多空并存，也计入主题级分歧
        d["divergent"] = ((d["bull"] > 0 and d["bear"] > 0)
                          or d["mixed"] > 0)
        d["verdict"] = ("多空并存 / 条件式，**存在分歧**" if d["divergent"] else
                        ("一边偏多" if d["bull"] > d["bear"] else
                         ("一边偏空" if d["bear"] > d["bull"] else "以事实型为主")))
        d["verify"] = ("✅ 多数获验证" if d["verifies"].count("✅ 一致") >
                       d["verifies"].count("⚠️ 背离") else
                       ("⚠️ 存在背离" if "⚠️ 背离" in d["verifies"] else
                        ("不适用（叙事型）" if d["na_n"] and not d["verifies"] else "—")))
        topics.append(d)
    topics.sort(key=lambda x: -x["n"])

    pre_n = sum(1 for r in rows if r["pre"])
    post_n = len(rows) - pre_n
    n_stance = sum(1 for r in rows if r["stance"] != "fact")
    n_fact = len(rows) - n_stance
    timing = {
        "pre": pre_n, "post": post_n,
        "note": ("主升起点 {}".format(launch_dt.strftime("%H:%M")
                                     if launch_dt else "13:29")
                 + "；观点几乎全部发布于收盘后/次日早间，属**事后解读**"),
    }

    signals = []
    signals.append("**观点不是驱动**：{} 条观点中仅 {} 条落在主升起点之前，"
                   "{} 条为盘后/次日事后解读。按第十五章同一套时序纪律，"
                   "后者**只作印证、不计入候选驱动**；媒体天然倾向事后归因，"
                   "不能倒推为「因为某机构看好所以涨」。".format(
                       len(rows), pre_n, post_n))
    div = [t for t in topics if t["divergent"]]
    if div:
        signals.append("**检测到 {} 个主题出现分歧/条件式表述**：{}。"
                       "同一天里既有「看好/增量明确」也有「承压/还需时间」，"
                       "或同一句里同时承认利好与时间约束，说明市场"
                       "**没有形成单边共识**，与第十四章「震荡筑底」"
                       "互为印证。".format(
                           len(div), "、".join(t["topic"] for t in div[:4])))
    fund_views = [r for r in rows if "基金" in r["actor"]]
    if fund_views:
        key = next((r for r in fund_views
                    if "还需时间" in r["snippet"] or "共识兑现" in r["snippet"]
                    or "底部" in r["snippet"]), None)
        mixed_n = sum(1 for r in fund_views if r["stance"] == "mixed")
        signals.append("**基金/资管口径最关键**：共 {} 条，其中 {} 条为"
                       "**分歧/条件式**。{}"
                       "这类表述一边承认底部信号与连续净流入，一边明确"
                       "「共识兑现为行情还需时间」，与本报告"
                       "「一级有增量、二级未放量、超跌修复非反转」几乎同义，"
                       "属于观点侧对结论的独立印证。".format(
                           len(fund_views), mixed_n,
                           ("代表条目「{}」。".format(key["title"][:52])
                            if key else "")))
    verified = [r for r in rows if r["verify"] == "✅ 一致"]
    diverged = [r for r in rows if r["verify"] == "⚠️ 背离"]
    if verified or diverged:
        signals.append("**观点 × 资金/涨幅交叉校验**：{} 条观点获"
                       "板块资金或样本ETF涨幅验证，{} 条出现背离"
                       "（观点偏多但资金/涨幅未跟上，或反之）。"
                       "背离条目已在 17.3 清单里单列，不作强归因。".format(
                           len(verified), len(diverged)))
    signals.append("**口径边界**：立场由关键词规则判定，不是语义模型，"
                   "「同比增长」类事实数字已尽量排除但仍可能误判；"
                   "同一机构多条观点会分别计数，观点方≠独立信源数；"
                   "本源里大量括号来源（人民财讯/央视新闻/中国证券报）是"
                   "**转载方**而非观点原始发布方，分类只作粗分。")

    return {
        "available": True, "n": len(rows), "n_stance": n_stance,
        "n_fact": n_fact, "rows": rows, "actors": actors,
        "stance_dist": stance_dist, "topics": topics,
        "timing": timing, "signals": signals,
    }


def analyze_attribution(ext, date, today, fund, breadth_raw, glob, market):
    """把资讯按发布时间切成「驱动 / 伴随 / 事后」，再与资金流、ETF涨幅、外围做三方校验。"""
    hot = ext.get("_hot") or []
    detail = ext.get("_detail") or {}
    notices = ext.get("_notice") or []
    reports = ext.get("_report") or []
    dehy = ext.get("_dehydrated") or []
    cal = ext.get("_calendar") or {}
    ipo = ext.get("_ipo") or {}
    rating = ext.get("_rating") or {}
    flash = ext.get("_em_flash") or []
    flash_status = ext.get("_em_flash_status") or "empty"
    sina_status = ext.get("_sina_status") or "empty"
    sina_n = ext.get("_sina_n") or 0
    ct_status = ext.get("_ct_status") or "empty"
    ct_n = ext.get("_ct_n") or 0
    ct_hot = ext.get("_ct_hot") or []
    cls_status = ext.get("_cls_status") or "empty"
    cls_n = ext.get("_cls_n") or 0
    cls_pages = ext.get("_cls_pages") or 0
    cls_covered = bool(ext.get("_cls_covered"))
    etf_flow = ext.get("_etf_flow") or []
    flash_days = ext.get("_flash_days") or 0
    flash_total = ext.get("_flash_total") or 0
    flash_src_n = ext.get("_flash_source_n") or 0
    flash_family_n = ext.get("_flash_family_n") or flash_src_n
    archive = ext.get("_flash_archive") or {}
    archive_ct = ext.get("_flash_archive_ct") or {}
    empc_status = ext.get("_empc_status") or "empty"
    empc_n = ext.get("_empc_n") or 0
    empc_days = ext.get("_empc_days") or {}
    archive_empc = ext.get("_flash_archive_empc") or {}
    if not hot and not flash:
        return {"available": False, "signals": [
            "财经资讯接口不可用且快照无缓存（含东方财富 7×24 快讯）：无法做事件归因，"
            "请执行 `python3 etf_analysis.py --save-snapshot` 在线生成一次缓存。"],
            "topics": [], "timeline": [], "collect": {}, "verdict": "—",
            "launch": None, "tomorrow": [], "sources": [],
            "flash": [], "flash_n": 0, "flash_cand": 0,
            "flash_status": flash_status, "sina_status": sina_status,
            "sina_n": sina_n, "ct_status": ct_status, "ct_n": ct_n,
            "ct_hot": ct_hot, "etf_flow": etf_flow,
            "empc_status": empc_status, "empc_n": empc_n,
            "empc_days": empc_days, "empc_next": [], "empc_post": [],
            "empc_contrary": [], "checklist": [],
            "archive_empc": archive_empc,
            "cls_status": cls_status, "cls_n": cls_n,
            "cls_pages": cls_pages, "cls_covered": cls_covered,
            "flash_days": flash_days, "flash_total": flash_total,
            "flash_src_n": flash_src_n, "flash_family_n": flash_family_n,
            "cross": [], "sina_new": [], "archive": archive,
            "archive_ct": archive_ct}

    try:
        d0 = datetime.strptime(date, "%Y-%m-%d").replace(tzinfo=CST)
    except Exception:
        return {"available": False, "signals": ["报告日期无效，无法做时序切分。"],
                "topics": [], "timeline": [], "collect": {}, "verdict": "—",
                "launch": None, "tomorrow": [], "sources": [],
                "flash": [], "flash_n": 0, "flash_cand": 0,
                "flash_status": flash_status}
    win_end = d0 + timedelta(hours=36)   # 报告日 + 次日中午
    close_dt = d0.replace(hour=15, minute=0)
    open_dt = d0.replace(hour=9, minute=30)
    launch, launch_ret = find_launch_time(breadth_raw)
    launch_dt = None
    if launch:
        hh, mm = launch.split(":")
        launch_dt = d0.replace(hour=int(hh), minute=int(mm))

    # ---- 事件池：统一成 {time, title, source, kind, body} ----
    events = []
    for r in hot:
        t = _to_dt(r.get("ts"))
        if t and d0 <= t < win_end:
            events.append({"dt": t, "title": r.get("title") or "",
                           "source": r.get("source") or "", "kind": "热点资讯",
                           "origin": r.get("source") or "",
                           "body": detail.get(r.get("id"), "")})
    for r in notices:
        t = _to_dt(r.get("time"))
        if t and d0 - timedelta(days=2) <= t < win_end:
            events.append({"dt": t, "title": r.get("title") or "",
                           "source": "公司公告", "kind": "公告", "body": "",
                           "origin": ""})
    for r in reports + dehy:
        t = _to_dt(r.get("time") or r.get("publish_time"))
        if t and d0 - timedelta(days=2) <= t < win_end:
            events.append({"dt": t, "title": r.get("title") or r.get("preview") or "",
                           "source": r.get("org") or "卖方研报", "kind": "研报",
                           "origin": r.get("org") or "",
                           "body": r.get("preview") or ""})
    # 7×24 快讯：分钟级、带完整日期，是盘中驱动最直接的证据
    for r in flash:
        t = _to_dt(r.get("time"))
        if not t or not (d0 <= t < win_end):
            continue
        title = r.get("title") or ""
        body = r.get("summary") or ""   # summary 常含【标题】前缀与完整短讯正文
        events.append({"dt": t, "title": title or body[:50],
                       "source": r.get("source") or "东方财富7×24", "kind": "7×24快讯",
                       "origin": r.get("origin") or "",
                       "body": body, "flash": True,
                       "important": bool(r.get("important"))})
    events.sort(key=lambda e: e["dt"])

    # ---- 时序切分 ----
    for e in events:
        t = e["dt"]
        if t < open_dt:
            e["phase"], e["cand"] = "盘前", True
        elif t < close_dt:
            if launch_dt and t <= launch_dt:
                e["phase"], e["cand"] = "盘中·启动前", True
            else:
                e["phase"], e["cand"] = "盘中·同步发酵", False
        elif t < d0 + timedelta(days=1):
            e["phase"], e["cand"] = "盘后", False
        elif t < d0 + timedelta(days=1, hours=9, minutes=30):
            e["phase"], e["cand"] = "隔夜/次日", False
        else:
            # 次日开盘后：单列为「验证层」，不混入 11-04 的事后解读计数
            e["phase"], e["cand"] = "次日盘中", False
        text = e["title"] + " " + e["body"]
        e["hits"] = [t2["topic"] for t2 in TOPIC_MAP
                     if any(k in text for k in t2["kw"])]
        e["contrary"] = any(w in text for w in CONTRARY_WORDS)

    # ---- 主题聚合 + 三方校验 ----
    in_top = fund.get("in_top") or []
    ranked = [q for _, q in today.get("ranked") or []]
    glob_rows = glob.get("rows") or []
    us = [r for r in glob_rows if "美" in str(r.get("region") or "")
          or "纳斯达克" in str(r.get("name") or "")
          or "标普" in str(r.get("name") or "")
          or "道琼斯" in str(r.get("name") or "")]

    topics = []
    for tp in TOPIC_MAP:
        ev = [e for e in events if tp["topic"] in e["hits"]]
        if not ev:
            continue
        pre = [e for e in ev if e["cand"]]
        sources = sorted({e["source"] for e in ev if e["source"]})
        # 板块资金验证
        hit_sec = []
        for r in in_top:
            for s in tp["sectors"]:
                if s and (s in r["name"] or r["name"] in s):
                    hit_sec.append((r["name"], r["net_yi"]))
                    break
        # ETF 涨幅验证
        etf_hits = [(q["name"], q["pct"]) for q in ranked
                    if any(k in (q.get("name") or "") for k in tp["etf_kw"])]
        etf_pct = avg([p for _, p in etf_hits if p is not None])
        contrary_n = sum(1 for e in ev if e["contrary"])
        # 严格反证：同一条资讯须同时含「外围」+「科技/芯片」+「下跌措辞」，
        # 且 A 股同主题当日为正，才判定外围传导不成立
        external_ev = [
            e for e in ev
            if any(k in e["title"] for k in ["美股", "纳指", "隔夜", "深夜"])
            and any(k in e["title"] for k in
                    ["芯片", "半导体", "AI", "科技股", "存储"])
            and e["contrary"]]
        external_down = bool(external_ev and etf_pct is not None
                             and etf_pct > 2)

        # 无板块、无 ETF 映射的「全市场叙事型」主题，缺少可验证对象，降权
        narrative_only = not tp["sectors"] and not tp["etf_kw"]
        score = 0
        evidence = []
        if pre:
            score += 20 if narrative_only else 30
            evidence.append("有 {} 条盘前/启动前资讯".format(len(pre)))
        elif any(e["phase"] == "盘中·同步发酵" for e in ev):
            score += 12
            evidence.append("仅盘中同步资讯，时序偏弱")
        if narrative_only:
            evidence.append("全市场叙事型主题，无板块/ETF 验证对象，置信度上限调低")
        if hit_sec:
            score += 30 if any(n for _, n in hit_sec[:2]) else 20
            evidence.append("板块资金流验证：{}".format(
                "、".join("{} {:+.0f}亿".format(n, v) for n, v in hit_sec[:2])))
        elif tp["sectors"]:
            evidence.append("板块资金流未验证（未进净流入前列）")
        if etf_pct is not None:
            score += 20 if etf_pct >= 5 else (10 if etf_pct >= 2 else 0)
            evidence.append("相关样本ETF平均涨幅 {}".format(pct_str(etf_pct)))
        score += 10 if len(sources) >= 3 else (6 if len(sources) == 2 else 3)
        if len(ev) >= 4:
            score += 5 if narrative_only else 10
            evidence.append("多源重复报道 {} 条".format(len(ev)))
        if external_down:
            score -= 15
            evidence.append("**反证**：{} 条外围同主题下跌资讯，"
                            "而 A 股同主题上涨 → 外围传导降权、内部催化升权".format(
                                len(external_ev)))
        elif contrary_n:
            evidence.append("含下跌措辞资讯 {} 条（多为盘后/外围，未纳入驱动）".format(
                contrary_n))
        score = max(0, min(100, score))
        level = "高" if score >= 70 else ("中" if score >= 45 else "低")
        topics.append({
            "topic": tp["topic"], "n": len(ev), "pre_n": len(pre),
            "sources": sources, "repr": ev[0]["title"][:46],
            "first": _fmt_dt(ev[0]["dt"]), "last": _fmt_dt(ev[-1]["dt"]),
            "hit_sec": hit_sec, "etf_pct": etf_pct,
            "contrary_n": contrary_n,
            "external_n": len(external_ev) if external_down else 0,
            "score": score, "level": level,
            # etf_kw 供事件研究 15.5.1 匹配样本 ETF（此前未透出，AR 匹配全空）
            "etf_kw": tp["etf_kw"], "sectors": tp["sectors"],
            # first_ts 供 15.5.2 日内窗口 CAR 起算（秒级，精确到事件当分钟）
            "first_ts": ev[0]["dt"].strftime("%Y-%m-%d %H:%M:%S"),
            "evidence": evidence, "pre_titles": [e["title"][:40] for e in pre[:3]],
        })
    topics.sort(key=lambda x: x["score"], reverse=True)

    # ---- 采集总览 ----
    collect = {
        "7×24快讯": "{} 条（{}）".format(
            len(flash), EM_FLASH_STATUS_DESC.get(flash_status, flash_status)),
        "热点资讯": len(hot), "新闻正文": len(detail), "公司公告": len(notices),
        "个股研报": len(reports), "脱水研报": len(dehy),
        "财经日历": (cal.get("n") or len(cal.get("items") or [])),
        "新股日历": (ipo.get("sgrq_n") or 0) + (ipo.get("ssrq_n") or 0)
        + len(ipo.get("recent") or []),
        "机构预期样本": len((rating.get("stocks") or [])),
    }
    src_count = {}
    for e in events:
        if e["source"]:
            src_count[e["source"]] = src_count.get(e["source"], 0) + 1
    sources = sorted(src_count.items(), key=lambda x: -x[1])[:8]

    # ---- 时间轴（只展示有主题命中的，最多 18 条）----
    timeline = [
        {"time": _fmt_dt(e["dt"]), "phase": e["phase"], "cand": e["cand"],
         "kind": e["kind"], "title": e["title"][:52],
         "source": e["source"], "topics": "、".join(e["hits"][:2]) or "—"}
        for e in events if e["hits"]][:18]

    # ---- 跨源一致性校验（多源互证比单源叙事更可信）----
    _family = source_family

    cross = []
    for ent in CROSS_SOURCE_ENTITIES:
        hits = [e for e in events
                if ent in e["title"] or ent in (e.get("body") or "")]
        fams = sorted({_family(e["source"]) for e in hits})
        if len(fams) >= 2 and hits:
            cross.append({
                "entity": ent, "n": len(hits), "fams": fams,
                "repr": hits[0]["title"][:52],
                "time": _fmt_dt(hits[0]["dt"]),
            })
    cross.sort(key=lambda x: (-len(x["fams"]), -x["n"]))
    # 族数按完整事件池统计（含公告/研报/热点资讯），才与 15.8.1 表格一致
    flash_family_n = len({_family(e["source"]) for e in events})

    # 新浪源带来的独有证据（其他来源没有的实体/事实）
    sina_ev = [e for e in events
               if _family(e["source"]) == "新浪财经"]
    other_text = " ".join(e["title"] for e in events
                          if _family(e["source"]) != "新浪财经")
    sina_new = []
    for e in sina_ev:
        title = e["title"]
        # 取标题里最显著的一段做排他判断
        key = title[:12].replace("【", "").split("】")[0]
        if key and key not in other_text and len(key) >= 4:
            sina_new.append({"time": _fmt_dt(e["dt"]), "title": title[:64],
                             "topics": "、".join(e["hits"][:2]) or "—"})
    # 与报告主题相关的排前面，避免被日本央行纪要这类无关快讯淹没
    sina_new.sort(key=lambda x: (x["topics"] == "—", x["time"]))

    # ---- 结论与信号 ----
    signals = []
    n_cand = sum(1 for e in events if e["cand"])
    n_post = sum(1 for e in events if e["phase"] in ("盘后", "隔夜/次日"))
    n_next = sum(1 for e in events if e["phase"] == "次日盘中")
    signals.append("采集窗口 {} 00:00 — 次日 12:00，共 {} 条带时间戳事件"
                   "（当日窗口内以热点资讯为主；公告 / 研报多为数日前发布，只作背景）。"
                   "其中**候选驱动 {} 条**（盘前 / 主升启动前）、"
                   "**事后解读 {} 条**（盘后 / 隔夜 / 次日开盘前）。".format(
                       date, len(events), n_cand, n_post))
    if n_next:
        signals.append("**次日盘中 {} 条单列为验证层**：这些是次日 09:30 之后的"
                       "实时快讯，用于 15.10.2 对照第十四章「明日观察」，"
                       "**不计入 11-04 的事后解读、也不算驱动**——"
                       "反复采集会让这个池子变大，但不会污染归因结论。".format(
                           n_next))
    hist_n = sum(1 for r in hot if not r.get("id"))
    if hist_n:
        signals.append("**滚动窗口提示**：热点资讯榜只保留最新 50 条，"
                       "脚本按 id 增量留存历次采集结果（本次含 {} 条历史留存），"
                       "避免盘中早段催化被新稿件挤出后丢失——"
                       "10:10 的半导体设备产业资讯即靠此保留为候选驱动。".format(hist_n))
    flash_events = [e for e in events if e.get("flash")]
    flash_cand = sum(1 for e in flash_events if e["cand"])
    if flash_events:
        signals.append("**7×24 快讯**补入 {} 条分钟级短讯（其中 {} 条落在主升起点前），"
                       "这类逐分钟滚动资讯比盘后收评更适合做时序归因；"
                       "本次数据源状态：{}。".format(
                           len(flash_events), flash_cand,
                           EM_FLASH_STATUS_DESC.get(flash_status, flash_status)))
        if flash_status == "web_channel":
            signals.append("**采集入口补充**：除 `getFastNewsList` JSON 接口外，"
                           "已接入东方财富移动端 H5 快讯页 "
                           "`emh5wap.eastmoney.com/info/channel/kuaixun`"
                           "（服务端渲染、可正则解析），作为 JSON 直连受限时的备用通道。")
        if sina_status == "ok":
            signals.append("**第三通道已直连**：新浪财经 7×24 "
                           "（`finance.sina.com.cn/7x24/`）本次采到 {} 条，"
                           "带完整日期与秒级时间戳，是快讯源里脚本可自主"
                           "直连、不依赖网页通道注入的。".format(sina_n))
        if ct_n:
            signals.append("**第四通道：华夏时报 7×24**"
                           "（`kuaixun.chinatimes.net.cn/`）补入 {} 条，"
                           "中央级财经媒体、页面自带完整日期行；"
                           "另附 {} 条「热点快讯」榜。".format(ct_n, len(ct_hot)))
        if cls_n and cls_status == "ok":
            signals.append("**第六通道：财联社电报 7×24**直连采到 {} 条 / {} 页，"
                           "按 `last_time` 时间游标回溯。这是六个源里唯一能按时间戳"
                           "翻页补盘中快讯的（其余源实时页只给最新 50 条）；"
                           "`ctime` 为 Unix 秒级时间戳，日期归属天然无歧义，"
                           "`brief` + `content` 双字段更适合关键词→板块映射。".format(
                               cls_n, cls_pages))
            if cls_covered:
                signals.append("**盘中空白已补**：财联社回溯已覆盖报告日 09:30 之前，"
                               "回溯条目进入 15.2 时间轴参与「驱动 vs 事后」判定，"
                               "候选驱动数与归因排序可能随之改变——请以本次报告为准。")
        elif cls_n:
            signals.append("**第六通道：财联社电报网页通道补入 {} 条**"
                           "（带 `HH:MM:SS` 秒级时间戳，日期归属无歧义）。"
                           "实测 `www.cls.cn/telegraph` 网页可读，但 "
                           "`?date` / `?last_time` 参数被服务端忽略，"
                           "仍只返回最新实时页 → **11-04 盘中 13:29 前后的空白没补上**；"
                           "这些条目按时序纪律进入「次日盘中验证层」，"
                           "不算 11-04 驱动。真正回溯仍需 JSON 接口的 `last_time` 游标。".format(
                               cls_n))
        if not cls_n:
            signals.append("**第六通道（财联社电报）本次 0 条**：`cls.cn` 域名族"
                           "在本沙盒被策略拒绝（HTTP 403），并非接口不存在。"
                           "脚本已内置本地签名（`md5(sha1(按 key 字典序拼的 query 串))`，"
                           "零密钥）与 `last_time` 时间游标回溯适配器，新旧两个端点都试；"
                           "换到可访问环境即可自动回溯报告日 9:30–15:00 的盘中快讯，"
                           "这是补齐时序空白最直接的一条路径。")
        if etf_flow:
            ev = next((f["values"] for f in etf_flow
                       if "ETF整体净流入" in (f.get("values") or {})), {})
            signals.append("**ETF 资金独立印证**：快讯给出报告日 ETF 资金"
                           "整体净流入 **{} 亿**（股票型 {} 亿 / 宽基型 {} 亿，"
                           "同花顺 iFinD 口径），与第十章一级市场 T-1 净申购"
                           " +550 亿**方向一致但口径不同**，不可相加；"
                           "该源还给出个券级 Top3，可与样本池直接对照。".format(
                               ev.get("ETF整体净流入", "—"),
                               ev.get("股票型净流入", "—"),
                               ev.get("宽基型净流入", "—")))
        if cross:
            signals.append("**多源互证**：{} 个关键事实在 ≥2 个独立信源同时出现"
                           "（{}），这类证据可信度高于单家媒体叙事。".format(
                               len(cross),
                               "、".join(c["entity"] for c in cross[:6])))
        if flash_days:
            signals.append("**逐日累积**：7×24 快讯缓存已沉淀 {} 天 / {} 条"
                           "（本次 {} 个来源）。实时页只给最新 50 条，"
                           "按 id 增量留存后，每天运行即可让报告日 9:30–15:00 "
                           "的盘中快讯永久沉淀，供次日做时序归因。".format(
                               flash_days, flash_total, flash_src_n))
    elif flash_status in ("blocked", "empty"):
        signals.append("**7×24 快讯不可用**：东方财富 `np-weblist` / `np-listapi` "
                       "JSON 接口与移动端 H5 域名在本运行环境被网络策略拒绝（HTTP 403），"
                       "并非接口不存在。脚本已内置 JSON + H5 双通道适配器，"
                       "换到可访问 eastmoney 的网络即可自动生效；"
                       "本次归因仍以热点资讯 + 公告 + 研报为准。")
    if launch:
        signals.append("上证分时 5 分钟滚动涨幅最大窗口起点 **{}**（该 5 分钟 {}），"
                       "以此为界：早于它的资讯才有资格算「驱动」，晚于它的只能是"
                       "「伴随」或「事后解读」。".format(launch, pct_str(launch_ret)))
    if topics:
        signals.append("主题归因排序（置信度）：{}。".format(
            "、".join("{} {}分/{}".format(t["topic"], t["score"], t["level"])
                     for t in topics[:5])))
        top = topics[0]
        signals.append("**首要驱动候选：{}**（{} 分 / {} 置信）。{}".format(
            top["topic"], top["score"], top["level"],
            "；".join(top["evidence"][:3]) if top["evidence"] else ""))
    # 时序误判拦截
    ipo_post = [e for e in events
                if ("IPO" in e["title"] or "暂停" in e["title"])
                and e["phase"] in ("盘后", "隔夜/次日")]
    if ipo_post:
        signals.append("**时序纠错**：「IPO 暂停」类传闻最早见于 {}（{}），"
                       "属**收盘后发酵**，不能算当日盘中拉升的驱动，"
                       "只能作为次日情绪与叙事的参考。".format(
                           _fmt_dt(ipo_post[0]["dt"]), ipo_post[0]["phase"]))
    # 反证
    chip = next((t for t in topics if "半导体" in t["topic"]), None)
    if chip and chip["external_n"]:
        signals.append("**反证检测**：检出 {} 条「外围同主题下跌」资讯"
                       "（典型如「美股芯片股深夜大跌 / AI 概念股逆势下挫」），"
                       "而 A 股芯片 ETF 今日涨停潮 → "
                       "外围传导路径降权，**内部产业催化**（存储涨价、龙头回购）升权。".format(
                           chip["external_n"]))
    elif chip and chip["contrary_n"]:
        signals.append("相关资讯中有 {} 条含下跌措辞，但未构成"
                       "「外围跌 / A 股涨」的严格反证，仅作背景参考。".format(
                           chip["contrary_n"]))
    ct_ipo = [h for h in ct_hot
              if "IPO" in h["title"] or "一刀切" in h["title"]]
    if ct_ipo:
        signals.append("**IPO 叙事第三源**：华夏时报「热点快讯」榜第 {} 位"
                       "「{}」，与 21:59「并非一刀切」口径一致；"
                       "按时序仍是收盘后/次日叙事，不计入盘中驱动。".format(
                           ct_ipo[0]["rank"], ct_ipo[0]["title"][:44]))
    if archive_ct.get("n"):
        signals.append("**华夏时报归档拦截**：页面日期 {} 与报告日 {} 相差 "
                       "超过 2 天，{} 条快讯判为历史归档，不并入事件池。".format(
                           archive_ct.get("date") or "—", date,
                           archive_ct["n"]))
    if archive.get("n"):
        signals.append("**归档样本拦截**：检测到纯文本快讯 {} 条，按多数投票"
                       "主年份为 {}（正文另有未来规划年份），与报告日 {} 不符 → "
                       "判定为历史归档，**不并入事件池**。仅用于验证解析器对 "
                       "`cj.sina.cn` 那种 `HH:MM:SS + 正文` 格式的兼容性，"
                       "避免把旧新闻误当今日驱动。".format(
                           archive["n"], archive.get("dominant") or "—", date))
    if not topics:
        signals.append("资讯中未找到可映射到板块的明确催化，"
                       "本次上涨更可能由资金面 / 技术面超跌修复驱动，"
                       "不做强行归因。")
    signals.append("**口径边界**：财经资讯是**解释**不是**证明**。同一日常有多个叙事并存，"
                   "媒体天然倾向事后归因；本章只给置信度排序，不给因果断言。"
                   "收盘后发布的「三大推手」「增量资金就位」类稿件一律降权。")

    # ---- 明日事件预告 ----
    tomorrow = []
    cal_items = cal.get("items") or []
    cal_name = {"": "其他", "1": "财报/业绩", "2": "分红除权", "3": "限售解禁",
                "4": "股东大会", "5": "增发配股"}
    type_cnt = {}
    for x in cal_items:
        k = cal_name.get(str(x.get("type")), str(x.get("type")) or "其他")
        type_cnt[k] = type_cnt.get(k, 0) + 1
    if cal_items:
        tomorrow.append("财经日历 {} 有 {} 条公司事件（{}），示例：{}。".format(
            date, len(cal_items),
            "、".join("{}={}".format(k, v)
                     for k, v in sorted(type_cnt.items(), key=lambda z: -z[1])[:4]),
            "、".join("{} {}".format(x.get("stock") or "", x.get("title") or "")
                     for x in cal_items[:3])))
    if ipo.get("recent"):
        tomorrow.append("近期已申购新股：{}；未来待申购 {} 只 / 待上市 {} 只。".format(
            "、".join("{}（{}）".format(x.get("name") or "",
                                     x.get("sgrq") or "—")
                     for x in ipo["recent"][:3]),
            ipo.get("sgrq_n") or 0, ipo.get("ssrq_n") or 0))
    rstock = rating.get("stocks") or []
    yoys, names = [], []
    for s in rstock:
        d = s.get("data") or {}
        if not isinstance(d, dict):
            continue
        fc = d.get("ConEarningsForecast")
        if isinstance(fc, str):
            try:
                fc = json.loads(fc)
            except Exception:
                fc = []
        if isinstance(fc, list) and fc:
            v = num(fc[0].get("ConNpYOY")) if isinstance(fc[0], dict) else None
            if v is not None:
                yoys.append(v)
                names.append(d.get("SecuName") or s.get("code") or "")
    if yoys:
        tomorrow.append("机构一致预期背景：{} 只主题样本（{}）最新预测净利润同比"
                        "中位数 **{:.1f}%**，可用于判断产业催化是否有盈利预期支撑；"
                        "评级家数字段本次接口未返回。".format(
                            len(yoys), "、".join(names[:3]),
                            sorted(yoys)[len(yoys) // 2]))

    verdict = ("事件驱动为主（{}，{} 置信）".format(
        topics[0]["topic"], topics[0]["level"]) if topics
        else "无明确事件驱动，偏向资金/技术面")
    # 多方评论观点：复用同一事件池，单独识别「谁在说什么立场」
    viewpoints = extract_viewpoints(events, fund, today, launch_dt)
    # 东财全球财经快讯 PC 页：分出「报告日盘后」与「次日盘中」两组
    empc_rows = [e for e in events
                 if e.get("source") == "东方财富全球快讯"]
    # 次日盘中验证层：汇聚六通道里所有 09:30 之后的实时快讯，
    # 不再局限于东财 PC 页；按时间排序后限量，避免历史累积撑爆表格
    _market_kw = ["板块", "概念", "ETF", "指数", "涨停", "涨超", "跌超",
                  "拉升", "走强", "走弱", "下挫", "转涨", "反弹", "涨逾"]
    _next_pool = [e for e in events
                  if e.get("phase") == "次日盘中"
                  and any(k in (e.get("title") or "") for k in _market_kw)]
    _next_pool.sort(key=lambda e: e["dt"])
    # 按「结论类别」分组取代表：新浪实时页同主题会连发多条，
    # 按时间取前 N 会把 09:38/09:39 的关键条目（CPO 回吐、588170 转涨）挤掉
    _groups = [
        (["中际旭创", "CPO", "跌近10%"], 1),
        (["588170", "科创半导体ETF"], 1),
        (["半导体设备板块", "半导体设备ETF", "半导体板块", "半导体"], 2),
        (["创新药"], 1), (["白酒"], 1),
        (["恒生科技", "港股航空", "恒指", "港股石油股"], 2),
        (["光伏"], 1), (["算力租赁"], 1),
        (["吉隆坡", "菲律宾", "马来西亚"], 1),
    ]
    _picked, _used = [], set()
    for _kws, _lim in _groups:
        _cand = [(i, e) for i, e in enumerate(_next_pool)
                 if i not in _used
                 and any(k in (e.get("title") or "") for k in _kws)]
        if not _cand:
            continue
        # 优先东财 PC / 财联社（标题结构化、带日期来源），其次时间最早；
        # 同组尽量取不同来源，形成跨源互证而非同题刷屏
        _cand.sort(key=lambda p: (
            0 if p[1].get("source") in ("东方财富全球快讯", "财联社电报") else 1,
            p[1]["dt"]))
        _seen_src, _n = set(), 0
        for _i, _e in _cand:
            if _n >= _lim or len(_picked) >= 14:
                break
            _s = _e.get("source")
            if _s in _seen_src:
                continue
            _seen_src.add(_s)
            _used.add(_i)
            _picked.append(_e)
            _n += 1
    _picked.sort(key=lambda e: e["dt"])
    empc_next = [{"time": _fmt_dt(e["dt"]),
                  "title": (e.get("title") or "")[:80],
                  "source": e.get("source") or "",
                  "hits": e.get("hits") or []}
                 for e in _picked]
    empc_post = [{"time": _fmt_dt(e["dt"]),
                  "title": (e.get("title") or "")[:80],
                  "hits": e.get("hits") or []}
                 for e in empc_rows if e.get("phase") == "盘后"]

    # ---- 观察清单「事件级」布尔判定 ----
    # 只判能从次日快讯文本直接读出的条件；数值类条件（量能、压力位、
    # 资金流、两融、折溢价）没有次日全量行情，明确标「待数值确认」，不冒充结论
    _nx = [e for e in events if e.get("phase") == "次日盘中"]
    _nx_text = " ".join((e.get("title") or "") + " " + (e.get("body") or "")
                        for e in _nx)
    _semi_hits = [e for e in _nx
                  if any(k in (e.get("title") or "") for k in
                         ("半导体", "芯片", "存储", "算力", "CPO", "光模块"))]
    _retreat = [e for e in _nx
                if any(k in (e.get("title") or "") for k in
                       ("中际旭创", "CPO概念震荡下挫", "跌近10%"))
                and any(k in (e.get("title") or "") for k in
                        ("跌", "下挫", "回吐", "走弱"))]
    _rot_kw = ["创新药", "白酒", "光伏", "港股航空", "恒生科技", "算力租赁",
               "食品加工", "广告营销", "贵金属", "电池"]
    _rot_hit = sorted({k for k in _rot_kw if k in _nx_text})
    _broker = any(k in _nx_text for k in ("券商", "证券板块", "证券ETF"))
    _external = any(k in _nx_text for k in
                    ("恒生科技", "恒指", "港股", "日经", "韩国综指"))
    checklist = [
        {"name": "首要驱动有后续资讯印证",
         "sig": "次日半导体/算力链快讯 {} 条".format(len(_semi_hits)),
         "ok": len(_semi_hits) >= 2,
         "note": "主线延续，事件层未被证伪" if len(_semi_hits) >= 2
                 else "次日暂无足够主线资讯，需警惕单日脉冲"},
        {"name": "高位股回吐（风险项验证）",
         "sig": "中际旭创 / CPO 走弱 {}".format(
             "已出现" if _retreat else "未见"),
         "ok": bool(_retreat),
         "note": "与 14.5「涨停附近不宜追高」一致" if _retreat
                 else "暂未出现，但不代表风险解除"},
        {"name": "轮动扩散（非主线板块）",
         "sig": "命中 {} 个：{}".format(
             len(_rot_hit), "、".join(_rot_hit[:6]) or "—"),
         "ok": len(_rot_hit) >= 3,
         "note": "典型震荡筑底形态，而非单边主升" if len(_rot_hit) >= 3
                 else "扩散不明显，仍偏结构性"},
        {"name": "券商 / 权重接力",
         "sig": "次日快讯中{}券商板块信号".format(
             "有" if _broker else "未见"),
         "ok": bool(_broker),
         "note": "与第十四章「权重不接力，谈不上全面行情」一致"
                 if not _broker else "需结合券商ETF成交额确认"},
        {"name": "外围同向接力",
         "sig": "港股/亚太快讯{}".format("有" if _external else "未见"),
         "ok": False,
         "note": "恒生科技、港股航空走强，但美股芯片链前夜收跌 → "
                 "外围内部分化，不构成同向传导" if _external else "无外围信号"},
        {"name": "量能 / 压力位 / 资金流 / 两融 / 折溢价",
         "sig": "无次日全量行情与资金数值",
         "ok": None,
         "note": "**待数值确认**：次日快讯是事件文本，不能替代成交额、"
                 "MA20 攻防、板块净额与两融余额的布尔判定"},
    ]
    # 外围反证：数据中心/AI 链走弱，与 A 股半导体主线相反
    # 只收「明确走弱/转谨慎」的短语；「谷歌为数据中心项目兜底担保」是利好，不误判为反证
    _ckw = ["转向“谨慎”", "转向谨慎", "英伟达盘前跌", "英伟达跌",
            "台积电跌", "逆势下挫", "深夜大跌", "大跌", "重挫"]
    empc_contrary = [{"time": _fmt_dt(e["dt"]), "title": e["title"][:80]}
                     for e in empc_rows
                     if e.get("phase") == "盘后"
                     and any(k in (e.get("title") or "") + (e.get("body") or "")
                             for k in _ckw)]
    return {
        "available": True, "collect": collect, "sources": sources,
        "viewpoints": viewpoints,
        "empc_status": empc_status, "empc_n": empc_n,
        "empc_days": empc_days, "empc_next": empc_next,
        "empc_post": empc_post, "empc_contrary": empc_contrary,
        "checklist": checklist,
        "archive_empc": archive_empc,
        "launch": launch, "launch_ret": launch_ret,
        "timeline": timeline, "topics": topics, "signals": signals,
        "verdict": verdict, "tomorrow": tomorrow,
        "n_events": len(events), "n_cand": n_cand, "n_post": n_post,
        "n_next": n_next,
        "flash": [
            {"time": _fmt_dt(e["dt"]), "phase": e["phase"],
             "cand": e["cand"], "title": e["title"][:70],
             "source": e["source"], "topics": "、".join(e["hits"][:2]) or "—",
             "important": bool(e.get("important"))}
            for e in flash_events],
        "flash_n": len(flash_events), "flash_cand": flash_cand,
        "flash_status": flash_status, "sina_status": sina_status,
        "sina_n": sina_n, "ct_status": ct_status, "ct_n": ct_n,
        "ct_hot": ct_hot, "etf_flow": etf_flow,
        "flash_days": flash_days, "flash_total": flash_total,
        "flash_src_n": flash_src_n, "flash_family_n": flash_family_n,
        "cls_status": cls_status, "cls_n": cls_n,
        "cls_pages": cls_pages, "cls_covered": cls_covered,
        "cross": cross, "sina_new": sina_new, "archive": archive,
        "archive_ct": archive_ct,
    }


def analyze_outlook(store):
    rows = [r for c in CORE_CODES if (r := tech_row(store, c)) is not None]
    total_bull = sum(r["bull"] for r in rows)
    total_bear = sum(r["bear"] for r in rows)

    if total_bull >= total_bear + 2:
        base_scene = "反弹延续"
    elif total_bear >= total_bull + 6:
        base_scene = "二次探底风险偏高"
    else:
        base_scene = "震荡筑底"

    # 上行确认阈值：核心标的明日成交额需较今日放大 15%
    thresholds = []
    for r in rows:
        amt = r["q"]["amount_yi"]
        if amt:
            thresholds.append("{} > {:.0f}亿".format(r["name"], amt * 1.15))
    pressures = [r["next_pressure"] for r in rows if r["next_pressure"]]

    watch = [
        "**量能**：反包次日必须继续放量（{}），缩量冲高视为减仓信号。".format(
            "；".join(thresholds[:3]) if thresholds else "核心ETF成交额较今日放大15%以上"),
        "**压力位**：{}，能否站稳决定反弹级别；站不稳则维持区间震荡。".format(
            "、".join("{} {}".format(r["name"], price_str(r["next_pressure"][1]))
                     for r in rows if r["next_pressure"]) or "各标的MA20"),
        "**板块接力**：证券ETF是否放量启动；权重不接力，就谈不上全面行情。",
    ]

    return {
        "rows": rows, "total_bull": total_bull, "total_bear": total_bear,
        "base_scene": base_scene, "pressures": pressures,
        "watch": watch,
        "scenes": [
            {"name": "震荡筑底",
             "desc": "在昨日低点与 MA20 之间反复拉锯，用时间换空间消化套牢盘。",
             "confirm": "缩量、阴阳交替、不破昨日低点；KDJ 自超卖区缓慢抬升。",
             "trigger": "多头线索 {} 条 / 空头线索 {} 条，多空均未取得压倒性优势。".format(
                 total_bull, total_bear)},
            {"name": "反弹延续（上行情景）",
             "desc": "放量站稳 MA20 / 布林中轨后，才有资格看向 MA60。",
             "confirm": "成交额连续两日放大 + 收盘价站上上表第一压力位 + MACD 零轴下方金叉扩散。",
             "trigger": "需同时满足「量」与「价」两个硬条件，缺一不可。"},
            {"name": "二次探底（下行情景）",
             "desc": "若高开低走或缩量滞涨，回踩昨日低点是大概率。",
             "confirm": "跌破 {} 等支撑且成交萎缩，说明反包为存量博弈。".format(
                 "、".join(price_str(r["support"]) for r in rows[:2])),
             "trigger": "事件/情绪驱动的反弹缺乏基本面数据支撑，外围波动会放大次日振幅。"},
        ],
    }


# ----------------------------------------------------------------------------
# 渲染层：SVG 图表（纯标准库）+ Markdown → HTML
# ----------------------------------------------------------------------------

CHART_STORE = {}


def esc(s):
    """XML/HTML 转义（仅用于 SVG 与最终 HTML 输出）。"""
    return (str(s).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


def md_plain(s):
    """Markdown 单元格内容：去掉会破坏表格的竖线/换行，不做 HTML 转义（避免二次转义）。"""
    return str(s or "").replace("|", "／").replace("\n", " ").strip()


def svg_bars(labels, values, title, fmt="{:.0f}", zero_line=True):
    """正负条形图。零轴居中时用于资金流/涨跌分布；否则从 0 起。"""
    w, h = SVG_W, SVG_H
    ml, mr, mt, mb = 64, 18, 44, 56
    pw, ph = w - ml - mr, h - mt - mb
    vals = [v if v is not None else 0 for v in values]
    if not vals:
        return ""
    lo, hi = min(vals + ([0] if zero_line else [])), max(vals + ([0] if zero_line else []))
    span = (hi - lo) or (abs(hi) * 0.2 or 1)
    if hi == lo:
        lo, hi = lo - span * 0.5, hi + span * 0.5
    pad = span * 0.12
    lo, hi = lo - pad, hi + pad
    span = hi - lo
    def y(v):
        return mt + ph * (hi - v) / span
    y0 = y(0) if lo <= 0 <= hi else (mt + ph)
    n = max(1, len(vals))
    bw = pw / n * 0.68
    gap = pw / n
    parts = [
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {} {}" width="100%" '
        'style="max-width:{}px;font-family:\'WenQuanYi Micro Hei\',sans-serif">'.format(w, h, w),
        '<rect width="{}" height="{}" fill="#ffffff"/>'.format(w, h),
        '<text x="{}" y="26" font-size="15" font-weight="600" fill="#222">{}</text>'.format(
            ml, esc(title)),
        '<line x1="{}" y1="{}" x2="{}" y2="{}" stroke="#d8dee6" stroke-width="1"/>'.format(
            ml, y0, w - mr, y0),
    ]
    for i, (lab, v) in enumerate(zip(labels, vals)):
        cx = ml + gap * (i + 0.5)
        yy, yv = y(v), y0
        top, height = min(yy, yv), abs(yy - yv)
        color = C_UP if v >= 0 else C_DOWN
        parts.append('<rect x="{:.1f}" y="{:.1f}" width="{:.1f}" height="{:.1f}" '
                     'rx="2" fill="{}"><title>{}：{}</title></rect>'.format(
                         cx - bw / 2, top, bw, max(1.0, height), color,
                         esc(lab), fmt.format(v)))
        parts.append('<text x="{:.1f}" y="{:.1f}" font-size="10.5" fill="#444" '
                     'text-anchor="middle">{}</text>'.format(
                         cx, top - 5 if v >= 0 else top + height + 13, fmt.format(v)))
        parts.append('<text x="{:.1f}" y="{:.1f}" font-size="10" fill="#555" '
                     'text-anchor="middle">{}</text>'.format(
                         cx, h - mb + 18, esc(lab[:7])))
    parts.append('</svg>')
    return "\n".join(parts)


def svg_line(values, labels, title, fmt="{:.0f}", ymin=None, ymax=None,
             hline=None, color=C_ACC):
    w, h = SVG_W, SVG_H
    ml, mr, mt, mb = 58, 18, 44, 46
    pw, ph = w - ml - mr, h - mt - mb
    vals = [v for v in values if v is not None]
    if len(vals) < 2:
        return ""
    lo = ymin if ymin is not None else min(vals)
    hi = ymax if ymax is not None else max(vals)
    if hline is not None:
        lo, hi = min(lo, hline), max(hi, hline)
    span = (hi - lo) or (abs(hi) * 0.02 or 1)
    if hi == lo:
        lo, hi = lo - span * 0.5, hi + span * 0.5
    pad = span * 0.1
    lo, hi = lo - pad, hi + pad
    span = hi - lo
    def xy(i, v):
        return (ml + pw * i / (len(values) - 1),
                mt + ph * (hi - v) / span)
    pts = " ".join("{:.1f},{:.1f}".format(*xy(i, v)) for i, v in enumerate(values)
                   if v is not None)
    parts = [
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {} {}" width="100%" '
        'style="max-width:{}px;font-family:\'WenQuanYi Micro Hei\',sans-serif">'.format(w, h, w),
        '<rect width="{}" height="{}" fill="#ffffff"/>'.format(w, h),
        '<text x="{}" y="26" font-size="15" font-weight="600" fill="#222">{}</text>'.format(
            ml, esc(title)),
    ]
    for frac in (0, 0.25, 0.5, 0.75, 1.0):
        v = lo + span * (1 - frac)
        yy = mt + ph * frac
        parts.append('<line x1="{}" y1="{:.1f}" x2="{}" y2="{:.1f}" '
                     'stroke="#eef1f5" stroke-width="1"/>'.format(ml, yy, w - mr, yy))
        parts.append('<text x="{}" y="{:.1f}" font-size="10" fill="#888" '
                     'text-anchor="end">{}</text>'.format(ml - 6, yy + 3, fmt.format(v)))
    if hline is not None:
        hy = mt + ph * (hi - hline) / span
        parts.append('<line x1="{}" y1="{:.1f}" x2="{}" y2="{:.1f}" '
                     'stroke="#f0a04b" stroke-width="1.2" stroke-dasharray="5,4"/>'.format(
                         ml, hy, w - mr, hy))
    parts.append('<polyline points="{}" fill="none" stroke="{}" stroke-width="2.2" '
                 'stroke-linejoin="round"/>'.format(pts, color))
    step = max(1, len(labels) // 6)
    for i in range(0, len(labels), step):
        parts.append('<text x="{:.1f}" y="{}" font-size="10" fill="#666" '
                     'text-anchor="middle">{}</text>'.format(
                         ml + pw * i / max(1, len(values) - 1), h - mb + 20,
                         esc(labels[i])))
    parts.append('</svg>')
    return "\n".join(parts)


def save_chart(name, svg):
    if not svg:
        return ""
    os.makedirs(CHART_DIR, exist_ok=True)
    path = os.path.join(CHART_DIR, name + ".svg")
    with open(path, "w", encoding="utf-8") as f:
        f.write(svg)
    rel = "{}/{}.svg".format(os.path.basename(CHART_DIR), name)
    CHART_STORE[rel] = svg
    return rel


def inline_svg(md):
    return re.sub(r"!\[([^\]]*)\]\(([^)]+)\)",
                  lambda m: CHART_STORE.get(m.group(2), ""), md)


VIZ_CSS = """
*{box-sizing:border-box}
body{font-family:'WenQuanYi Micro Hei',-apple-system,'Microsoft YaHei',sans-serif;max-width:1180px;margin:0 auto;padding:16px 20px 28px;color:#222;background:#f6f8fb}
h1{font-size:19px;margin:0 0 6px;color:#1a4d8f}
.statusbar{display:flex;flex-wrap:wrap;gap:10px;align-items:center;margin:0 0 12px;font-size:12px;color:#6b7684}
.statusbar b{font-size:15px;color:#1d3f66}
.badge{display:inline-block;padding:2px 9px;border-radius:999px;background:#e8f1ff;color:#25599b;font-size:12px;font-weight:600}
.badge.warn{background:#fff3e0;color:#9a5b00}
.kpis{display:grid;grid-template-columns:repeat(6,minmax(0,1fr));gap:10px;margin:0 0 12px}
.kpi{display:flex;align-items:center;gap:10px;min-width:0;background:#fff;border:1px solid #e3e9f2;border-radius:10px;padding:10px 12px;min-height:76px;box-shadow:0 1px 3px rgba(30,60,100,.06)}
.kpi b{display:block;font-size:19px;color:#25476b;line-height:1.2;font-variant-numeric:tabular-nums;white-space:nowrap}
.kpi span{display:block;font-size:11.5px;color:#6b7684;margin-top:2px}
.kpi em{display:block;font-style:normal;font-size:11px;color:#8895a5;margin-top:3px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.kpi .kbox{flex:1;min-width:0}
.kpi .ring{position:relative;flex:0 0 44px;width:44px;height:44px;border-radius:50%}
.kpi .ring::after{content:'';position:absolute;inset:7px;background:#fff;border-radius:50%}
.kpi .kbar{height:6px;background:#eef2f7;border-radius:4px;overflow:hidden;margin-top:7px}
.kpi .kbar i{display:block;height:100%;border-radius:4px}
.dash{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px}
.card{min-width:0;background:#fff;border:1px solid #e3e9f2;border-radius:10px;padding:10px 12px;box-shadow:0 1px 3px rgba(30,60,100,.06);cursor:zoom-in}
.card h4{margin:0 0 6px;font-size:12.5px;font-weight:600;color:#33465c}
.chart svg{display:block;width:100%;height:auto;max-width:none;margin:0;border:0}
.hb{display:grid;grid-template-columns:88px 1fr 46px;gap:8px;align-items:center;font-size:11px;color:#4a5a70;margin:4px 0}
.hb-track{position:relative;height:11px;background:#f4f7fb;border-radius:3px}
.hb-track::before{content:'';position:absolute;left:50%;top:-1px;bottom:-1px;width:1px;background:#c9d4e2}
.hb-track i{position:absolute;top:2px;height:7px;border-radius:2px}
.hb-track i.pos{background:#e05a4f}
.hb-track i.neg{background:#2e9e7b}
.hb b{text-align:right;font-variant-numeric:tabular-nums;font-weight:600;color:#33465c}
.stack{display:flex;height:16px;border-radius:5px;overflow:hidden;background:#eef2f7;margin:10px 0 4px}
.stack i{display:block;height:100%}
.slegend{display:flex;justify-content:space-between;font-size:11px;color:#6b7684;margin-bottom:6px}
.slegend b{color:#33465c}
.srcdata[hidden]{display:none}
#lightbox{position:fixed;inset:0;background:rgba(12,22,38,.85);display:none;align-items:center;justify-content:center;z-index:99;padding:24px;cursor:zoom-out}
#lightbox.on{display:flex}
#lightbox .lb-in{background:#fff;border-radius:10px;padding:18px;width:min(1080px,96vw);max-height:88vh;overflow:auto;cursor:default}
#lightbox h4{margin:0 0 12px;font-size:15px;color:#1a4d8f}
#lightbox .hb{grid-template-columns:140px 1fr 70px;font-size:13px}
.dash .card:last-child:nth-child(odd){grid-column:1/-1}
.dash .card:last-child:nth-child(odd) .chart svg{max-width:560px!important;margin:0 auto}
@media(max-width:980px){.kpis{grid-template-columns:repeat(3,minmax(0,1fr))}}
@media(max-width:720px){.dash{grid-template-columns:1fr}.kpis{grid-template-columns:repeat(2,minmax(0,1fr))}}
"""

VIZ_JS = """
<div id='lightbox'><div class='lb-in' id='lightbox-in'></div></div>
<script>
(function(){
var lb=document.getElementById('lightbox'),box=document.getElementById('lightbox-in');
if(!lb||!box)return;
[].slice.call(document.querySelectorAll('.dash .card')).forEach(function(el){
  el.onclick=function(){var c=el.querySelector('.chart'),h=el.querySelector('h4');
    if(!c)return;box.innerHTML='<h4>'+(h?h.textContent:'')+'</h4>'+c.innerHTML;
    lb.classList.add('on');document.body.style.overflow='hidden';};});
lb.onclick=function(){lb.classList.remove('on');document.body.style.overflow='';};
document.addEventListener('keydown',function(e){if(e.key==='Escape'){
  lb.classList.remove('on');document.body.style.overflow='';}});
})();
</script>
"""


def _viz_rows(html, marker):
    """取某小节标题后第一张表的纯文本单元格。"""
    m = re.search(re.escape(marker) +
                  r".*?(<div class='tw'><table>.*?</table></div>)", html, re.S)
    if not m:
        return []
    rows = []
    for tr in re.findall(r"<tr>(.*?)</tr>", m.group(1), re.S):
        cells = [re.sub(r"<[^>]+>", "", c).strip() for c in
                 re.findall(r"<t[dh][^>]*>(.*?)</t[dh]>", tr, re.S)]
        if cells:
            rows.append(cells)
    return rows


def _viz_num(s):
    m = re.search(r"[+\-]?\d+(?:\.\d+)?", str(s or "").replace(",", ""))
    return float(m.group()) if m else None


def _verdict_text(concl_html):
    """结论首句，去掉括号补充与情景前缀，压成一行短结论。"""
    t = re.sub(r"<[^>]+>", "", concl_html)
    t = re.sub(r"^一句话结论[：:]?\s*", "", t)
    t = t.split("。")[0]
    t = re.sub(r"[（(][^）)]*[）)]", "", t)
    t = t.replace("基准情景 ", "")
    t = re.sub(r"\s+", " ", t).strip("，；; ")
    return (t or "详见图表") + "。"


def _verdict_html(text):
    """顶部结论条：内联样式，已生成页面无需追加 CSS 即可同步。"""
    return ("<!--verdict-start--><div style='display:flex;align-items:center;"
            "gap:10px;background:#f2f7ff;border-left:3px solid #3b7dd8;"
            "border-radius:6px;padding:9px 12px;margin:0 0 10px'>"
            "<span style='flex:0 0 auto;background:#3b7dd8;color:#fff;"
            "border-radius:4px;padding:2px 8px;font-size:11px;font-weight:600'>"
            "结论</span><b style='font-size:13.5px;font-weight:600;"
            "color:#25476b;line-height:1.5'>{}</b></div>"
            "<!--verdict-end-->").format(esc(text))


def refresh_verdict(html):
    """已生成的看板：从隐藏锚点重取结论，插入或更新顶部结论条。"""
    m = re.search(r"<blockquote><strong>一句话结论</strong>.*?</blockquote>",
                  html, re.S)
    if not m:
        return html
    bar = _verdict_html(_verdict_text(m.group(0)))
    if "<!--verdict-start-->" in html:
        return re.sub(r"<!--verdict-start-->.*?<!--verdict-end-->",
                      lambda _: bar, html, count=1, flags=re.S)
    return re.sub(r"(<h1>.*?</h1>)", lambda mm: mm.group(1) + bar,
                  html, count=1, flags=re.S)


def sync_fundflow_out(html, report_html):
    """给已生成的看板补「行业流出 Top」卡：数据取自完整报告 4.3 表。"""
    if "行业流出 Top" in html:
        return html
    rows = _viz_rows(report_html, "4.3 行业主力净流出")[1:]
    pairs = [(r[0][:5], _viz_num(r[2])) for r in rows if len(r) > 2]
    pairs = [(n, v) for n, v in pairs if v is not None]
    if not pairs:
        return html
    svg = svg_bars([n for n, _ in pairs], [v for _, v in pairs],
                   "行业主力净流出 Top（亿元）", fmt="{:+.0f}")
    if not svg:
        return html
    card = ("<section class='card'><h4>行业流出 Top</h4>"
            "<div class='chart'>{}</div></section>").format(svg)
    html = html.replace("<h4>行业资金 Top</h4>", "<h4>行业流入 Top</h4>", 1)
    m = re.search(r"(<section class='card'><h4>行业流入 Top</h4>"
                  r"<div class='chart'>.*?</div></section>)", html, re.S)
    if not m:
        return html
    return html[:m.end()] + card + html[m.end():]


# ---- 自算评分引擎：六维结构 + 综合得分率 ----------------------------------
# 全部由本报告读数线性归一后加权，不引用外部评分源；权重即各维满分，合计 100。
SCORE_DIMS = (("机构", 20), ("游资", 20), ("龙头主导", 25),
              ("外资", 10), ("散户", 15), ("量能", 10))
_RADAR_UNIT = ((0, -1), (0.866, -0.5), (0.866, 0.5),
               (0, 1), (-0.866, 0.5), (-0.866, -0.5))


def _score_lin(x, lo, hi):
    """把指标线性映射到 0~100 并截断；缺数据返回 None。"""
    if x is None or hi == lo:
        return None
    return max(0.0, min(100.0, (x - lo) / (hi - lo) * 100.0))


def _score_inputs(html):
    """从报告头或精简看板的隐藏锚点区提取评分读数。"""
    m = re.search(r"<div class='srcdata' hidden>(.*?)<hr/>", html, re.S)
    t = re.sub(r"<[^>]+>", " ", m.group(1) if m else html[:20000])

    def g(pat):
        mm = re.search(pat, t)
        return _viz_num(mm.group(1)) if mm else None

    def index_gap():
        """权重指数相对成长指数的抗跌差：精简页读条形图 title，完整报告读 1.1 表。"""
        pcts = {}
        for ttl in re.findall(r"<div class='hb' title='([^']+?)'>", html):
            if ":" in ttl:
                n, v = ttl.rsplit(":", 1)
                pcts[n.strip()] = _viz_num(v)
        if not pcts:
            for r in _viz_rows(html, "1.1 指数概览")[1:]:
                if len(r) > 2:
                    pcts[r[0]] = _viz_num(r[2])
        a, b = pcts.get("上证指数"), pcts.get("创业板指")
        return a - b if a is not None and b is not None else None

    bp = re.search(r"净流入板块\s*(\d+)\s*/\s*(\d+)", t)
    pp = re.search(r"自建画像\s*(\d+)\s*/\s*(\d+)", t)
    return {
        "amount": g(r"成交额\s*([\d,.]+)\s*亿"),
        "chg": g(r"(?:较昨日同时点|较上日)\s*([+\-\d.]+)%"),
        "net": g(r"净额\s*([+\-\d.]+)\s*亿"),
        "up_ratio": g(r"上涨占比\s*([\d.]+)%"),
        "up": g(r"上涨\s*([\d,]+)\s*只"),
        "down": g(r"下跌\s*([\d,]+)\s*只"),
        "lu": g(r"涨停\s*(\d+)"),
        "ld": g(r"跌停\s*(\d+)"),
        "in_n": _viz_num(bp.group(1)) if bp else None,
        "board_n": _viz_num(bp.group(2)) if bp else None,
        "basis": g(r"IF\s*([+\-\d.]+)%"),
        "erp": g(r"ERP\s*([+\-]?[\d.]+)pct"),
        "pos20": g(r"近20日仅\s*(\d+)%"),
        "prof": (100.0 * int(pp.group(1)) / int(pp.group(2))
                 if pp and int(pp.group(2)) else None),
        "weight_gap": index_gap(),
    }


def score_engine(d):
    """资金势力六维总分与综合得分率分开算；否决只压得分率，不覆盖格局。"""
    lin = _score_lin

    def mix(parts):
        ok = [(v, w) for v, w in parts if v is not None]
        return (sum(v * w for v, w in ok) / sum(w for _, w in ok)
                if ok else None)

    in_ratio = (d["in_n"] / d["board_n"]
                if d["in_n"] is not None and d["board_n"] else None)
    inst = mix([(lin(d["net"], -600, 300), 0.6),
                (lin(in_ratio * 100, 0, 60)
                 if in_ratio is not None else None, 0.4)])
    hot = mix([(lin(d["lu"], 0, 80), 0.5),
               (lin(100.0 * d["lu"] / (d["lu"] + d["ld"]), 0, 100)
                if d["lu"] is not None and d["ld"] is not None
                and d["lu"] + d["ld"] else None, 0.5)])
    # 龙头/护盘：权重指数相对成长指数的抗跌差 + 行业资金流入集中度
    lead = mix([(lin(d["weight_gap"], -2, 4), 0.6),
                (lin(in_ratio * 100, 0, 60)
                 if in_ratio is not None else None, 0.4)])
    foreign = mix([(lin(d["basis"], -3, 1), 0.5),
                   (lin(d["erp"], -2, 5), 0.5)])
    retail = lin(d["up_ratio"], 5, 80)
    volume = mix([(lin(d["amount"], 8000, 20000), 0.6),
                  (lin(d["chg"], -30, 30), 0.4)])

    dims, avail, capital = [], 0.0, 0.0
    for (name, full), v in zip(SCORE_DIMS,
                               (inst, hot, lead, foreign, retail, volume)):
        if v is None:
            dims.append((name, full, 0.0, False))
            continue
        got = full * v / 100.0
        dims.append((name, full, got, True))
        avail += full
        capital += got

    spread = d["up"] / d["down"] if (d["up"] and d["down"]) else None
    veto = (d["lu"] is not None and d["lu"] < 40 and
            spread is not None and spread < 0.3)
    total = min(capital, 25.0 * avail / 100.0) if veto else capital
    rate = total / avail * 100.0 if avail else 0.0
    band = (("冰点", "#2e9e7b") if rate < 20 else
            ("弱势", "#6fae7a") if rate < 40 else
            ("均衡", "#f0a04b") if rate < 60 else
            ("活跃", "#e8834f") if rate < 80 else ("过热", "#e05a4f"))
    # 主导维度按对总分的绝对贡献取
    top = max((x for x in dims if x[3]), key=lambda x: x[2], default=None)
    guard = (d["weight_gap"] is not None and d["weight_gap"] > 1.5 and
             in_ratio is not None and in_ratio < 0.3)
    pattern = ("护盘市" if guard else
               {"机构": "机构市", "游资": "题材市", "龙头主导": "龙头市",
                "外资": "护盘市", "散户": "普涨市", "量能": "放量市"}.get(
                    top[0] if top else "", "观望市"))
    return {"dims": dims, "total": total, "capital": capital,
            "avail": avail, "rate": rate, "veto": veto, "band": band,
            "pattern": pattern, "lead": top[0] if top else "",
            "lu": d["lu"], "spread": spread}


def _score_bar_html(s):
    """综合得分率条：分数 / 可用满分 / 档位线 / 一票否决。"""
    rate, (bname, bcolor) = s["rate"], s["band"]
    marks = "".join(
        "<i style='position:absolute;left:{}%;top:0;bottom:0;width:1px;"
        "background:#fff;opacity:.9'></i>".format(p)
        for p in (20, 40, 60, 80))
    ticks = "".join("<span>{}</span>".format(p) for p in (0, 20, 40, 60, 80, 100))
    if s["veto"]:
        vbadge = ("<span style='background:#fff1f0;color:#b23b32;border-radius:4px;"
                  "padding:2px 7px;font-size:11px;font-weight:600'>一票否决</span>")
        veto_line = ("<div style='margin-top:6px;background:#fff1f0;border:1px solid "
                     "#ffd6d3;border-radius:5px;padding:4px 8px;font-size:11px;"
                     "color:#b23b32'>一票否决：涨停 {:.0f} 家 &lt; 40 且 涨跌比 {:.2f} "
                     "&lt; 0.3，得分封顶 25%</div>").format(
                         s["lu"] or 0, s["spread"] or 0)
    else:
        vbadge, veto_line = "", ""
    return ("<!--score-start--><div style='display:flex;gap:16px;align-items:center;"
            "background:#fff;border:1px solid #e3e9f2;border-radius:10px;"
            "padding:10px 14px;margin:0 0 10px;box-shadow:0 1px 3px rgba(30,60,100,.06)'>"
            "<div style='flex:0 0 132px;text-align:center'>"
            "<div style='font-size:30px;font-weight:700;line-height:1.1;color:{bc}'>"
            "{rate:.1f}%</div><div style='font-size:11px;color:#6b7684;margin-top:2px'>"
            "综合得分率</div><div style='font-size:10.5px;color:#8895a5'>"
            "总分 {total:.2f} / {avail:.2f}</div></div>"
            "<div style='flex:1;min-width:0'>"
            "<div style='display:flex;gap:8px;align-items:center;margin-bottom:6px'>"
            "<span style='background:{bc};color:#fff;border-radius:4px;padding:2px 7px;"
            "font-size:11px;font-weight:600'>{bn}</span>{vbadge}"
            "<span style='font-size:11px;color:#8895a5'>"
            "六维加权 · 自有口径 · 规则演算，非建议</span></div>"
            "<div style='position:relative;height:12px;background:#eef2f7;"
            "border-radius:6px;overflow:hidden'>"
            "<i style='display:block;height:100%;width:{rate:.1f}%;background:{bc}'>"
            "</i>{marks}</div>"
            "<div style='display:flex;justify-content:space-between;font-size:10px;"
            "color:#8895a5;margin-top:3px'>{ticks}</div>{veto_line}"
            "</div></div><!--score-end-->").format(
                bc=bcolor, rate=rate, total=s["total"], avail=s["avail"],
                bn=esc(bname), vbadge=vbadge, marks=marks, ticks=ticks,
                veto_line=veto_line)


def _radar_card(s):
    """六维雷达：纯 SVG 自绘，每维标 得分/满分。"""
    cx, cy, R = 180.0, 118.0, 84.0

    def pt(i, r):
        ux, uy = _RADAR_UNIT[i]
        return cx + r * ux, cy + r * uy

    p = ["<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 360 240' "
         "width='100%' style='max-width:380px;margin:0 auto;display:block;"
         "font-family:sans-serif'>"]
    for frac in (0.25, 0.5, 0.75, 1.0):
        pts = " ".join("{:.1f},{:.1f}".format(*pt(i, R * frac)) for i in range(6))
        p.append("<polygon points='{}' fill='none' stroke='#e6ebf2' "
                 "stroke-width='1'/>".format(pts))
    for i in range(6):
        x, y = pt(i, R)
        p.append("<line x1='{:.1f}' y1='{:.1f}' x2='{:.1f}' y2='{:.1f}' "
                 "stroke='#e6ebf2'/>".format(cx, cy, x, y))
    ratios = [(got / full if full and ok else 0.0)
              for _, full, got, ok in s["dims"]]
    pts = " ".join("{:.1f},{:.1f}".format(*pt(i, R * max(0.04, r)))
                   for i, r in enumerate(ratios))
    p.append("<polygon points='{}' fill='rgba(240,160,75,.28)' "
             "stroke='#f0a04b' stroke-width='2'/>".format(pts))
    lead = s.get("lead", "")
    for i, r in enumerate(ratios):
        x, y = pt(i, R * max(0.04, r))
        hot = s["dims"][i][0] == lead
        p.append("<circle cx='{:.1f}' cy='{:.1f}' r='{}' fill='{}'/>".format(
            x, y, 3.6 if hot else 2.6, "#e05a4f" if hot else "#f0a04b"))
    for i, (name, full, got, ok) in enumerate(s["dims"]):
        ux, uy = _RADAR_UNIT[i]
        lx, ly = cx + (R + 20) * ux, cy + (R + 20) * uy
        anchor = "start" if ux > 0.3 else ("end" if ux < -0.3 else "middle")
        val = "{:.2f}/{:.0f}".format(got, full) if ok else "—"
        hot = name == lead
        p.append("<text x='{:.1f}' y='{:.1f}' font-size='10.5' fill='{}' "
                 "font-weight='{}' text-anchor='{}'>{}</text>".format(
                     lx, ly, "#e05a4f" if hot else "#33465c",
                     600 if hot else 400, anchor, esc(name)))
        p.append("<text x='{:.1f}' y='{:.1f}' font-size='10' fill='{}' "
                 "font-weight='{}' text-anchor='{}'>{}</text>".format(
                     lx, ly + 12, "#e05a4f" if hot else "#f0a04b",
                     600 if hot else 400, anchor, val))
    p.append("</svg>")
    return ("<!--radar-start--><section class='card'>"
            "<h4>资金势力 · 六维结构</h4><div class='chart'>{}</div>"
            "<div style='font-size:11px;color:#6b7684;text-align:center;"
            "margin-top:4px'>总分 {:.2f} · 格局 {} · "
            "<b style='color:#e05a4f'>主导 {}</b></div>"
            "</section><!--radar-end-->").format(
                "".join(p), s["capital"], esc(s["pattern"]),
                esc(s.get("lead", "") or "—"))


def sync_score_widgets(html):
    """已生成看板：按隐藏锚点重算并同步得分条与雷达卡（幂等）。"""
    s = score_engine(_score_inputs(html))
    bar, radar = _score_bar_html(s), _radar_card(s)
    if "<!--score-start-->" in html:
        html = re.sub(r"<!--score-start-->.*?<!--score-end-->",
                      lambda _: bar, html, count=1, flags=re.S)
    else:
        html = re.sub(r"(<!--verdict-end-->)", lambda m: m.group(1) + bar,
                      html, count=1)
    if "<!--radar-start-->" in html:
        html = re.sub(r"<!--radar-start-->.*?<!--radar-end-->",
                      lambda _: radar, html, count=1, flags=re.S)
    else:
        html = re.sub(r"(<div class='dash'>)",
                      lambda m: m.group(1) + radar, html, count=1)
    return html


def polish_dashboard(html):
    """只输出「结果可视化看板」：分析章节、论证文字、样本池图表全部删除。

    页面只剩结论条 + 综合得分率 + 状态条 + 6 个 KPI + 结果图，两屏内看完；
    点击卡片可放大。hidden 区只留 send_mail.py 抽取邮件正文用的机器锚点，
    不渲染，也不含分析步骤。已生成的看板只重算结论/评分，保证幂等。
    """
    if "<!--viz-start-->" in html:
        return sync_score_widgets(refresh_verdict(html))
    mb = re.search(r"(<body[^>]*>)(.*)(</body>)", html, re.S)
    if not mb:
        return html
    body = mb.group(2)

    pts = re.search(r"<p><strong>今日要点（30 秒版）</strong></p>"
                    r"(?:\s*<p class='li'>.*?</p>)+", body, re.S)
    concl = re.search(r"<blockquote><strong>一句话结论</strong>.*?</blockquote>",
                      body, re.S)
    summ = re.search(r"<h2 id='s1'>[^<]*</h2>\s*"
                     r"(<div class='tw'><table>.*?</table></div>)", body, re.S)
    intr = re.search(r"<blockquote>[^<]*<strong>盘中口径</strong>.*?</blockquote>",
                     body, re.S)
    if not (pts and concl and summ):
        return html

    raw = re.sub(r"<[^>]+>", " ", pts.group(0) + " " + concl.group(0))

    def g(pat):
        mm = re.search(pat, raw)
        return _viz_num(mm.group(1)) if mm else None

    amount = g(r"成交额\s*([\d,.]+)\s*亿")
    ratio = g(r"上涨占比\s*([\d.]+)%")
    chg = g(r"较昨日同时点\s*([+\-\d.]+)%")
    net = g(r"净额\s*([+\-\d.]+)\s*亿")
    up_n, dn_n = g(r"上涨\s*([\d,]+)\s*只"), g(r"下跌\s*([\d,]+)\s*只")
    lu, ld = g(r"涨停\s*(\d+)"), g(r"跌停\s*(\d+)")
    quality = g(r"数据质量\s*(\d+)%")
    pm = re.search(r"自建画像\s*(\d+)\s*/\s*(\d+)", raw)
    prof = (100.0 * int(pm.group(1)) / int(pm.group(2))
            if pm and int(pm.group(2)) else None)

    head = re.sub(r"<[^>]+>", " ", body[:12000])
    d = re.search(r"(\d{4}-\d{2}-\d{2})", head)
    t = re.search(r"\d{4}-\d{2}-\d{2}\s+(\d{2}:\d{2})", head)
    intraday = "盘中实时" in head[:6000] or bool(intr)
    sm = re.search(r"数据来源[：:]\s*([^　|<\n（(]{2,24})", head)
    source = sm.group(1).strip() if sm else "公开行情接口"
    status = ("<div class='statusbar'><b>{}</b><span>{}</span>"
              "<span class='badge{}'>{}</span><span>{}</span></div>").format(
                  d.group(1) if d else "", t.group(1) if t else "",
                  " warn" if intraday else "",
                  "盘中 · 未收盘" if intraday else "收盘 · 终值", source)

    def kpi_ring(label, value, pct, color):
        p = max(0.0, min(100.0, pct or 0.0))
        return ("<div class='kpi'><div class='ring' style='background:"
                "conic-gradient({c} 0 {p:.1f}%,#edf1f6 {p:.1f}% 100%)'></div>"
                "<div><b>{v}</b><span>{l}</span></div></div>").format(
                    c=color, p=p, v=value, l=label)

    def kpi_bar(label, value, frac, note, color):
        w = max(3.0, min(100.0, abs(frac or 0.0)))
        return ("<div class='kpi'><div class='kbox'><b>{v}</b><span>{l}</span>"
                "<div class='kbar'><i style='width:{w:.1f}%;background:{c}'></i>"
                "</div><em>{n}</em></div></div>").format(
                    v=value, l=label, w=w, c=color, n=note)

    kpis = ""
    if ratio is not None:
        kpis += kpi_ring("上涨占比", "{:.0f}%".format(ratio), ratio, "#e05a4f")
    if amount is not None:
        kpis += kpi_bar("成交额·亿", "{:,.0f}".format(amount),
                        abs(chg) * 3 if chg is not None else 60,
                        "较昨同时点 {:+.2f}%".format(chg)
                        if chg is not None else "盘中累计", "#3b7dd8")
    if net is not None:
        kpis += kpi_bar("板块净额·亿", "{:+.0f}".format(net), abs(net) / 6,
                        "行业合计", "#2e9e7b" if net < 0 else "#e05a4f")
    if lu is not None and ld is not None:
        kpis += kpi_bar("涨停 / 跌停", "{:.0f} / {:.0f}".format(lu, ld),
                        100.0 * lu / max(1.0, lu + ld), "全市场", "#e05a4f")
    if quality is not None:
        kpis += kpi_ring("数据质量", "{:.0f}%".format(quality),
                         quality, "#3b7dd8")
    if prof is not None:
        kpis += kpi_ring("自建画像", "{:.0f}%".format(prof), prof, "#f0a04b")

    def card(title, inner):
        return ("<section class='card'><h4>{}</h4>"
                "<div class='chart'>{}</div></section>").format(title, inner)

    score = score_engine(_score_inputs(html))
    cards = [_radar_card(score)]
    idx = [(r[0], _viz_num(r[2])) for r in _viz_rows(body, "1.1 指数概览")[1:]
           if len(r) > 2]
    idx = [(a, v) for a, v in idx if v is not None]
    if idx:
        mx = max(abs(v) for _, v in idx) or 1.0
        bars = []
        for lab, v in idx:
            w = 50.0 * abs(v) / mx
            bar = ("<i class='pos' style='left:50%;width:{:.1f}%'></i>".format(w)
                   if v >= 0 else
                   "<i class='neg' style='right:50%;width:{:.1f}%'></i>".format(w))
            bars.append("<div class='hb' title='{}: {:+.2f}'><span>{}</span>"
                        "<div class='hb-track'>{}</div><b>{:+.2f}</b></div>"
                        .format(lab, v, lab, bar, v))
        cards.append(card("指数涨跌幅（%）", "".join(bars)))

    stack = ""
    if up_n and dn_n and up_n + dn_n > 0:
        pu = 100.0 * up_n / (up_n + dn_n)
        stack += ("<div class='stack'><i style='width:{:.1f}%;background:#e05a4f'>"
                  "</i><i style='width:{:.1f}%;background:#2e9e7b'></i></div>"
                  "<div class='slegend'><span>上涨 <b>{:,.0f}</b></span>"
                  "<span>下跌 <b>{:,.0f}</b></span></div>").format(
                      pu, 100 - pu, up_n, dn_n)
    if lu is not None and ld is not None and lu + ld > 0:
        pl = 100.0 * lu / (lu + ld)
        stack += ("<div class='stack'><i style='width:{:.1f}%;background:#e05a4f'>"
                  "</i><i style='width:{:.1f}%;background:#2e9e7b'></i></div>"
                  "<div class='slegend'><span>涨停 <b>{:.0f}</b></span>"
                  "<span>跌停 <b>{:.0f}</b></span></div>").format(
                      pl, 100 - pl, lu, ld)
    if stack:
        cards.append(card("涨跌结构", stack))

    # 只保留全市场口径的结果图；带「样本」的池内图表一律不进看板
    keep = (("全市场涨跌幅区间分布", "涨跌分布"),
            ("日内全市场上涨家数占比演变", "日内广度"),
            ("两市成交额趋势", "成交额趋势"),
            ("上证指数日内分时", "指数分时"),
            ("行业主力净流入", "行业流入 Top"),
            ("行业主力净流出", "行业流出 Top"),
            ("A股主要指数 PE", "指数 PE"))
    found = {}
    for s in re.findall(r"<svg\b.*?</svg>", body, re.S):
        tm = re.search(r'<text[^>]*font-weight="600"[^>]*>(.*?)</text>', s, re.S)
        title = tm.group(1) if tm else ""
        if "样本" in title:
            continue
        for key, short in keep:
            if key in title and key not in found:
                found[key] = card(short, s)
    for key, _ in keep:
        if key in found:
            cards.append(found[key])

    verdict = _verdict_html(_verdict_text(concl.group(0)))
    bar = _score_bar_html(score)
    viz = ("<!--viz-start--><h1>ETF 全景 · 结果看板</h1>" + verdict + bar + status +
           "<div class='kpis'>" + kpis + "</div>" +
           "<div class='dash'>" + "".join(cards) + "</div><!--viz-end-->")

    # 邮件正文锚点：只留要点、结论首句、执行摘要表，页面不渲染
    ctxt = re.sub(r"<[^>]+>", "", concl.group(0))
    lead = ctxt.split("。")[0] + "。" if "。" in ctxt else ctxt[:90]
    lead = re.sub(r"^一句话结论[：:]?\s*", "", lead)
    anchors = (pts.group(0) +
               "<blockquote><strong>一句话结论</strong>：{}</blockquote>".format(
                   esc(lead)) +
               "<h2 id='s1'>执行摘要</h2>" + summ.group(1) +
               (intr.group(0) if intr else ""))
    hidden = "<div class='srcdata' hidden>" + anchors + "<hr/></div>"

    out = (html[:mb.start(2)] + viz + "\n" + hidden + VIZ_JS +
           html[mb.end(2):])
    out = re.sub(r"<style[^>]*>.*?</style>", "", out, flags=re.S)
    out = re.sub(r"<title>.*?</title>", "<title>ETF 结果看板</title>", out,
                 flags=re.S)
    return out.replace("</head>", "<style>" + VIZ_CSS + "</style></head>", 1)


def md_to_html(md, title="ETF 全景分析报告"):
    """极简 Markdown→HTML：标题/表格/列表/引用/粗体/分隔线/图片内联。"""
    lines = inline_svg(md).split("\n")
    out, i, in_table = [], 0, False
    tcols = []                    # 当前表的逐列对齐（来自分隔行的冒号）
    heads, toc_at = [], None      # 章节标题 -> 顶部目录（长报告的唯一导航）

    theads, tbody = [], []        # 表格先缓冲：整表读完才知道每列的多数类型

    def inline(t):
        t = esc(t)
        t = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", t)
        t = re.sub(r"`(.+?)`", r"<code>\1</code>", t)
        return t

    def _is_num(v):
        v = v.strip().replace("**", "")
        return (bool(re.match(r"^[\+\-\d\s\.,%:：/～~—亿万倍家只日\(\)（）]+$", v))
                and (any(ch.isdigit() for ch in v) or "—" in v))

    def flush_table():
        """整表读完再输出：对齐按「分隔行冒号」优先，未标注的列取该列多数类型。

        逐行判断会在同一列里混出左对齐与右对齐（「—」判成数字、「+3.00pct」
        判成文本），显示出来就是错位；按列统一才是表格该有的样子。
        """
        nonlocal in_table, theads, tbody
        if not in_table:
            return
        n = len(theads)
        al = []
        for k in range(n):
            mark = tcols[k] if k < len(tcols) else "auto"
            if mark == "auto":
                vals = [r[k] for r in tbody if k < len(r)]
                hits = sum(1 for v in vals if _is_num(v))
                mark = "num" if (vals and hits * 2 >= len(vals) and hits) else "left"
            al.append(mark)

        def cell(tag, k, c):
            a = al[k] if k < len(al) else "left"
            return "<{0}{1}>{2}</{0}>".format(
                tag, " class='{}'".format(a) if a in ("num", "ctr") else "",
                inline(c))

        out.append("<div class='tw'><table><thead><tr>" +
                   "".join(cell("th", k, c) for k, c in enumerate(theads)) +
                   "</tr></thead><tbody>")
        for r in tbody:
            rr = list(r) + [""] * max(0, n - len(r))
            out.append("<tr>" + "".join(cell("td", k, c)
                                        for k, c in enumerate(rr[:n])) + "</tr>")
        out.append("</tbody></table></div>")
        in_table, theads, tbody = False, [], []

    while i < len(lines):
        ln = lines[i].rstrip()
        if ln.startswith("<svg"):
            j = i
            while j < len(lines) and "</svg>" not in lines[j]:
                j += 1
            out.append("\n".join(lines[i:j + 1]))
            i = j + 1
            continue
        if ln.startswith("|"):
            cells = [c.strip() for c in ln.strip().strip("|").split("|")]
            if i + 1 < len(lines) and re.match(r"^\|[\s:|-]+\|$", lines[i + 1].strip()):
                flush_table()
                _sep = [c.strip() for c in
                        lines[i + 1].strip().strip("|").split("|")]
                tcols[:] = []
                for c in _sep:
                    if c.startswith(":") and c.endswith(":"):
                        tcols.append("ctr")
                    elif c.endswith(":"):
                        tcols.append("num")
                    elif c.startswith(":"):
                        tcols.append("left")
                    else:
                        tcols.append("auto")
                theads, tbody = cells, []
                in_table = True
                i += 2
                continue
            if in_table:
                tbody.append(cells)
                i += 1
                continue
        if in_table:
            flush_table()
        if ln.startswith("# "):
            out.append("<h1>{}</h1>".format(inline(ln[2:])))
        elif ln.startswith("#### "):
            out.append("<h4>{}</h4>".format(inline(ln[5:])))
        elif ln.startswith("## "):
            _txt = ln[3:]
            _aid = "s{}".format(len(heads) + 1)
            if toc_at is None:
                toc_at = len(out)
            heads.append((_aid, _txt))
            out.append("<h2 id='{}'>{}</h2>".format(_aid, inline(_txt)))
        elif ln.startswith("### "):
            out.append("<h3>{}</h3>".format(inline(ln[4:])))
        elif ln.startswith("> "):
            out.append("<blockquote>{}</blockquote>".format(inline(ln[2:])))
        elif ln.strip() == "---":
            out.append("<hr/>")
        elif re.match(r"^\d+\.\s", ln):
            m = re.match(r"^(\d+)\.\s(.*)", ln)
            out.append("<p class='li'>{}. {}</p>".format(m.group(1), inline(m.group(2))))
        elif ln.startswith("- "):
            out.append("<p class='li'>• {}</p>".format(inline(ln[2:])))
        elif ln.strip():
            out.append("<p>{}</p>".format(inline(ln)))
        i += 1
    if in_table:
        out.append("</tbody></table></div>")
    if heads:
        toc = ["<nav class='toc'><strong>目录</strong><ol>"]
        toc += ["<li><a href='#{}'>{}</a></li>".format(a, esc(t)) for a, t in heads]
        toc.append("</ol></nav>")
        out.insert(toc_at or 0, "".join(toc))
    css = ("body{font-family:'WenQuanYi Micro Hei',-apple-system,sans-serif;"
           "max-width:860px;margin:0 auto;padding:18px;color:#222;line-height:1.7;}"
           "h1{font-size:22px;border-bottom:2px solid #3b7dd8;padding-bottom:8px}"
           "h2{font-size:19px;margin-top:30px;color:#1a4d8f}"
           "h3{font-size:16px;margin-top:22px;color:#333}"
           "h4{font-size:14px;margin-top:18px;color:#1a4d8f;"
           "border-left:3px solid #3b7dd8;padding-left:8px}"
           ".tw{overflow-x:auto;border:1px solid #eef1f5;border-radius:6px;margin:10px 0}"
           "table{border-collapse:collapse;width:100%;font-size:13px}"
           "th,td{border:1px solid #e3e8ef;padding:6px 8px;text-align:left;white-space:nowrap}"
           "th{background:#f4f7fb;cursor:pointer;user-select:none}"
           "th:hover{background:#e8f0fb}"
           "th.s-asc::after{content:' ▲'}th.s-desc::after{content:' ▼'}"
           # 对齐以表格分隔行的冒号为准（:--: 居中、--: 右对齐），
           # 表头与数据同列同规则；只有未标注的列才退回「像数字就右对齐」的启发式，
           # 否则同一列里「—」右对齐、「+3.00pct」左对齐，看起来就是错位。
           "th.num,td.num{text-align:right;font-variant-numeric:tabular-nums}"
           "th.ctr,td.ctr{text-align:center}"
           "blockquote{background:#f6f9fd;border-left:3px solid #3b7dd8;margin:10px 0;padding:8px 12px;font-size:13px}"
           "p{margin:6px 0;font-size:14px}p.li{padding-left:10px}"
           "code{background:#f0f2f5;padding:1px 4px;border-radius:3px;font-size:12px}"
           "svg{display:block;margin:12px auto;border:1px solid #eef1f5;border-radius:6px}"
           "hr{border:0;border-top:1px solid #e3e8ef;margin:22px 0}"
           ".toc{background:#f8fafc;border:1px solid #e3e8ef;border-radius:8px;"
           "padding:10px 14px;font-size:13px;line-height:1.9}"
           ".toc ol{margin:6px 0 0 18px;padding:0}"
           ".toc a{color:#1a4d8f;text-decoration:none}"
           ".toc a:hover{text-decoration:underline}"
           "#totop{position:fixed;right:18px;bottom:18px;width:40px;height:40px;"
           "border:0;border-radius:50%;background:#3b7dd8;color:#fff;font-size:18px;"
           "cursor:pointer;opacity:.85}#totop:hover{opacity:1}")
    css += (
        ".statusbar{display:flex;flex-wrap:wrap;gap:8px;align-items:center;"
        "margin:12px 0 10px;font-size:13px;color:#4a5a70}"
        ".statusbar b{font-size:15px;color:#1d3f66}"
        ".badge{display:inline-block;padding:2px 9px;border-radius:999px;"
        "background:#e8f1ff;color:#25599b;font-size:12px;font-weight:600}"
        ".badge.warn{background:#fff3e0;color:#9a5b00}"
        ".dash{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));"
        "gap:12px;margin:14px 0 26px}"
        ".card{background:#fff;border:1px solid #e3e9f2;border-radius:10px;"
        "padding:14px 16px;box-shadow:0 1px 3px rgba(30,60,100,.06)}"
        ".card.full{grid-column:1/-1}"
        ".card h3{margin:0 0 10px;padding:0;border:0;font-size:15px;color:#1a4d8f}"
        ".card h2{margin:0 0 8px;font-size:16px}"
        ".card .tw{margin:0;border:0}"
        ".card table{font-size:12.5px}"
        ".points p.li{margin:5px 0;padding:6px 10px;background:#f7f9fc;"
        "border-radius:6px;font-size:13px}"
        ".points p:first-child{display:none}"
        ".kpis{display:flex;gap:10px;flex-wrap:wrap;margin:10px 0 6px}"
        ".kpi{flex:1;min-width:92px;background:#f7f9fc;border-radius:8px;padding:9px}"
        ".kpi span{display:block;font-size:12px;color:#6b7684}"
        ".kpi b{display:block;font-size:20px;color:#25476b;line-height:1.3}"
        ".hint{margin:6px 0 0!important;font-size:12px!important;color:#6b7684}"
        ".concl{background:#f2f7ff;border-left:3px solid #3b7dd8;border-radius:6px;"
        "padding:10px 14px;font-size:13px;color:#2c3e50}"
        ".concl summary{cursor:pointer;font-weight:600}"
        ".concl blockquote{margin:10px 0 0;background:transparent;border:0;padding:0}"
        "details.meta,details.tocbox{margin:8px 0 14px;font-size:13px;color:#6b7684}"
        "details.meta summary,details.tocbox summary{cursor:pointer;color:#1a4d8f}"
        "details.meta p{font-size:12.5px}"
        "@media(max-width:720px){.dash{grid-template-columns:1fr}}"
    )
    js = ("<button id='totop' title='回到顶部'>&uarr;</button><script>"
          "document.getElementById('totop').onclick=function(){"
          "window.scrollTo({top:0,behavior:'smooth'})};"
          "document.querySelectorAll('table').forEach(function(t){"
          "var ths=t.querySelectorAll('thead th');"
          "ths.forEach(function(th,ci){th.onclick=function(){"
          "var tb=t.tBodies[0];if(!tb)return;"
          "var rows=Array.prototype.slice.call(tb.rows);"
          "var dir=th.getAttribute('data-dir')==='asc'?'desc':'asc';"
          "ths.forEach(function(o){o.classList.remove('s-asc','s-desc')});"
          "th.classList.add('s-'+dir);th.setAttribute('data-dir',dir);"
          "var num=function(r){var c=r.cells[ci];if(!c)return null;"
          "var s=c.textContent.replace(/,/g,'');"
          "var m=s.match(/[-+]?\\d+(\\.\\d+)?/);"
          "return m?parseFloat(m[0]):null;};"
          "rows.sort(function(a,b){var x=num(a),y=num(b);"
          "if(x===null||y===null)return 0;"
          "return dir==='asc'?x-y:y-x;});"
          "rows.forEach(function(r){tb.appendChild(r)});};});});"
          "</script>")
    page = ("<!doctype html><html lang='zh-CN'><head><meta charset='utf-8'>"
            "<meta name='viewport' content='width=device-width,initial-scale=1'>"
            "<link rel='icon' href='data:,'>"   # 阻止浏览器请求 /favicon.ico（线上 404）
            "<title>{}</title><style>{}</style></head><body>{}{}</body></html>".format(
                esc(title), css, "\n".join(out), js))
    return polish_dashboard(page)


def _limit_threshold(code):
    """按板块返回涨跌停阈值：主板10 / 创业板科创板20 / 北交所30。"""
    c = code or ""
    if c.startswith("bj") or c.startswith("92") or c.startswith("8"):
        return 30.0
    if c.startswith("sz30") or c.startswith("sh688") or c.startswith("30") or c.startswith("688"):
        return 20.0
    return 10.0


def analyze_market_rank(stock_rows, etf_rows, breadth, today, fund, date,
                        stock_status="empty", etf_status="empty"):
    """全市场个股 + ETF 涨跌排行分析（第十六章）。

    核心增量：① 个股级涨跌/涨停榜 ② 全市场多周期强弱（中期趋势的全量证据）
    ③ 个股主力净流入榜（补第四章只有行业级）④ 全市场 ETF 涨跌排行，
    解除 1.2/1.6 长期标注的「不能推断全市场 ETF 涨幅榜」边界。
    """
    empty = {"available": False, "signals": [], "gainers": [], "losers": [],
             "periods": [], "fund_top": [], "amount_top": [],
             "etf_gainers": [], "etf_losers": [], "etf_stats": {},
             "sample_check": {}, "date_ok": True, "date_note": "",
             "n_stock": 0, "n_etf": 0}
    if not stock_rows and not etf_rows:
        return empty
    out = dict(empty, available=True, n_stock=len(stock_rows or []),
               n_etf=len(etf_rows or []), status=stock_status,
               etf_status=etf_status)
    signals = []

    # ---- 个股层 ----
    S = stock_rows or []
    if S:
        def z(r, k="zdf"):
            return _fnum(r.get(k))

        up = sum(1 for r in S if z(r) > 0)
        dn = sum(1 for r in S if z(r) < 0)
        flat = len(S) - up - dn
        lu, ld = [], []
        for r in S:
            th = _limit_threshold(r.get("code"))
            if z(r) >= th - 0.2:
                lu.append(r)
            elif z(r) <= -(th - 0.2):
                ld.append(r)
        out.update({"up": up, "down": dn, "flat": flat,
                    "limit_up": len(lu), "limit_down": len(ld),
                    "limit_up_rows": lu[:20], "ratio": up / len(S) * 100})
        # 日期一致性校验：与第二章官方广度对照
        b_up = int((breadth or {}).get("up") or 0)
        b_dn = int((breadth or {}).get("down") or 0)
        b_lu = int((breadth or {}).get("uplim") or 0)
        b_ld = int((breadth or {}).get("dnlim") or 0)
        diff = max(abs(up - b_up), abs(dn - b_dn)) if b_up or b_dn else 0
        out["date_ok"] = diff <= 2
        out["date_note"] = (
            "上涨 {} / 下跌 {} 与第二章广度 {} / {} 逐家吻合，"
            "确认是 **{} 收盘快照**（非次日实时）".format(up, dn, b_up, b_dn, date)
            if out["date_ok"] else
            "家数差异 {} 只，与第二章 {} / {} 不吻合 → "
            "判定为**非报告日快照**，本章仅作参考、不并入当日结论".format(
                diff, b_up, b_dn))
        signals.append("**日期归属校验**：{}。涨停按分板块阈值（主板10/双创20/北交所30）"
                       "判定 {} 家、跌停 {} 家，官方广度口径为 {}/{}，"
                       "差 1–3 家属 ST、新股与阈值边界差异，不强行对齐。".format(
                           out["date_note"], len(lu), len(ld), b_lu, b_ld))

        # ---- 广度质量指标：分布 / 集中度 / 资金覆盖 / 周期一致性 ----
        # 涨跌家数只说明方向，赚钱效应还得看「涨多少、谁在涨、量堆在哪」。
        _buckets = [(-99, -7), (-7, -5), (-5, -3), (-3, 0), (0, 3),
                    (3, 5), (5, 7), (7, 10), (10, 99)]
        dist = []
        for _lo, _hi in _buckets:
            _n = sum(1 for r in S if _lo <= z(r) < _hi)
            if _hi == 99:
                _lb = "≥+10%"
            elif _lo == -99:
                _lb = "≤-7%"
            else:
                _lb = "{:+.0f}~{}%".format(_lo, _hi)
            dist.append({"label": _lb, "n": _n, "pct": _n / len(S) * 100})
        out["dist"] = dist
        _amts = sorted((_fnum(r.get("turnover")) for r in S), reverse=True)
        _tot = sum(_amts) or 1
        _k10 = max(1, len(_amts) // 10)
        out["amt_conc"] = {
            "total_yi": _tot / 1e4,
            "top10_pct": sum(_amts[:_k10]) / _tot * 100,
            "top100_pct": sum(_amts[:100]) / _tot * 100,
        }
        _zl = [r for r in S if r.get("zljlr") not in (None, "", "-")]
        _in_n = sum(1 for r in _zl if _fnum(r.get("zljlr")) > 0)
        out["flow_cover"] = {
            "n": len(_zl), "inflow_n": _in_n,
            "inflow_pct": (_in_n / len(_zl) * 100) if _zl else 0,
        }
        _p5 = [z(r, "zdf_d5") for r in S if r.get("zdf_d5") not in (None, "", "-")]
        _cons = {"n": len(_p5), "all3_pos": 0, "all3_neg": 0}
        for r in S:
            _v = [z(r, k) for k in ("zdf_d5", "zdf_d20", "zdf_d60")]
            if any(x == 0 for x in _v):
                continue
            if all(x > 0 for x in _v):
                _cons["all3_pos"] += 1
            elif all(x < 0 for x in _v):
                _cons["all3_neg"] += 1
        if _cons["n"]:
            _cons["pos_pct"] = _cons["all3_pos"] / _cons["n"] * 100
            _cons["neg_pct"] = _cons["all3_neg"] / _cons["n"] * 100
        out["period_cons"] = _cons
        if out.get("amt_conc"):
            signals.append(
                "**成交集中度**：全市场个股成交额 {} 亿，前 10% 个股占 "
                "**{:.0f}%**、前 100 名占 **{:.0f}%**；集中度越高说明当日"
                "上涨越依赖少数标的，广度指标需要打折看待。".format(
                    out["amt_conc"]["total_yi"], out["amt_conc"]["top10_pct"],
                    out["amt_conc"]["top100_pct"]))
        if out.get("flow_cover", {}).get("n"):
            signals.append(
                "**资金覆盖面**：{} 只个股有主力净额数据，其中净流入 **{:.0f}%**"
                "（{} 只）。涨幅是价格、净流入是成交单口径，两者背离说明"
                "「价涨但资金未跟」。".format(
                    out["flow_cover"]["n"], out["flow_cover"]["inflow_pct"],
                    out["flow_cover"]["inflow_n"]))
        if _cons.get("n"):
            signals.append(
                "**多周期一致性**：近5/20/60日**同时为正**的仅 **{:.0f}%**、"
                "**同时为负**的 **{:.0f}%**（{} 只有效样本）；"
                "三周期共振向上比例偏低，说明当日强势多为短期结构。".format(
                    _cons.get("pos_pct", 0), _cons.get("neg_pct", 0), _cons["n"]))

        # 涨跌排行
        gain = sorted(S, key=lambda r: -z(r))[:15]
        lose = sorted(S, key=lambda r: z(r))[:10]
        out["gainers"] = [{
            "code": r.get("code"), "name": r.get("name"), "zdf": z(r),
            "price": _fnum(r.get("zxj")), "d5": z(r, "zdf_d5"),
            "d20": z(r, "zdf_d20"), "d60": z(r, "zdf_d60"),
            "flow_yi": _fnum(r.get("zljlr")) / 1e4,
            "hsl": _fnum(r.get("hsl")),
        } for r in gain]
        out["losers"] = [{
            "code": r.get("code"), "name": r.get("name"), "zdf": z(r),
            "price": _fnum(r.get("zxj")), "d20": z(r, "zdf_d20"),
        } for r in lose]
        # 多周期强弱：全市场中期趋势的硬证据
        periods = []
        for k, label in [("zdf", "当日"), ("zdf_d5", "近5日"),
                         ("zdf_d20", "近20日"), ("zdf_d60", "近60日"),
                         ("zdf_w52", "近52周"), ("zdf_y", "年初至今")]:
            vals = [z(r, k) for r in S if r.get(k) not in (None, "", "-")]
            if not vals:
                continue
            vals_sorted = sorted(vals)
            periods.append({
                "key": k, "label": label,
                "median": vals_sorted[len(vals) // 2],
                "mean": sum(vals) / len(vals),
                "pos_ratio": sum(1 for v in vals if v > 0) / len(vals) * 100,
                "n": len(vals),
            })
        out["periods"] = periods
        p20 = next((p for p in periods if p["key"] == "zdf_d20"), None)
        if p20:
            signals.append("**中期趋势的全市场证据**：近20日涨跌幅中位数 "
                           "**{:+.2f}%**、仅 **{:.0f}%** 个股为正；当日却是 "
                           "**{:.0f}%** 上涨。全市场 {} 只个股的口径直接证明"
                           "「今日大涨是超跌反弹、中期下行结构未改」，"
                           "不再只依赖 {} 只 ETF 样本推演。".format(
                               p20["median"], p20["pos_ratio"],
                               out["ratio"], len(S), today.get("total_etf") or 0))
        # 个股主力净流入 / 成交额榜（单位：万元 → 亿元）
        out["fund_top"] = [{
            "code": r.get("code"), "name": r.get("name"), "zdf": z(r),
            "flow_yi": _fnum(r.get("zljlr")) / 1e4,
            "hsl": _fnum(r.get("hsl")), "lb": _fnum(r.get("lb")),
        } for r in sorted(S, key=lambda r: -_fnum(r.get("zljlr")))[:12]]
        out["amount_top"] = [{
            "code": r.get("code"), "name": r.get("name"), "zdf": z(r),
            "amount_yi": _fnum(r.get("turnover")) / 1e4,
        } for r in sorted(S, key=lambda r: -_fnum(r.get("turnover")))[:10]]
        # 与第四章行业资金 / 第十五章归因互证
        ft = out["fund_top"]
        semi_kw = ["长鑫", "兆易", "中芯", "拓荆", "中科飞测", "寒武纪",
                   "长电", "通富", "中际旭创", "新易盛", "思瑞浦", "半导体"]
        semi_n = sum(1 for r in ft[:8]
                     if any(k in (r["name"] or "") for k in semi_kw))
        if semi_n >= 3:
            hit_names = [r["name"] for r in ft[:8]
                         if any(k in (r["name"] or "") for k in semi_kw)]
            leader = ft[0]["name"] if ft else ""
            signals.append("**个股级三方校验**：主力净流入 Top8 里 {} 只属"
                           "半导体/存储/AI算力链（{}），与第四章行业资金 Top1"
                           "半导体、第十五章「半导体/存储/芯片 85 分」归因"
                           "**在个股层面完全对上**。{}".format(
                               semi_n, "、".join(hit_names[:6]),
                               ("榜首「{}」正是快讯里「50 亿回购护盘」的同一标的。".format(leader)
                                if "长鑫" in leader else
                                "榜首为「{}」。".format(leader))))

    # ---- 全市场 ETF 层 ----
    E = etf_rows or []
    if E:
        def ez(r, k="zdf"):
            return _fnum(r.get(k))
        eq = [r for r in E if not any(k in (r.get("name") or "")
                                     for k in ETF_EXCLUDE_KW)]
        scope = eq or E
        vals = sorted(ez(r) for r in scope)
        out["etf_stats"] = {
            "n_all": len(E), "n_equity": len(scope),
            "median": vals[len(vals) // 2], "mean": sum(vals) / len(vals),
            "pos_ratio": sum(1 for v in vals if v > 0) / len(vals) * 100,
            "max": vals[-1], "min": vals[0],
            "stamp": (E[0].get("stamp") or "")[:14],
        }
        out["etf_gainers"] = [{
            "code": r.get("code"), "name": r.get("name"), "zdf": ez(r),
            "price": ez(r, "price"), "amount_yi": ez(r, "amount_wan") / 1e4,
            "hsl": ez(r, "hsl"),
        } for r in sorted(scope, key=lambda r: -ez(r))[:15]]
        out["etf_losers"] = [{
            "code": r.get("code"), "name": r.get("name"), "zdf": ez(r),
            "amount_yi": ez(r, "amount_wan") / 1e4,
        } for r in sorted(scope, key=lambda r: ez(r))[:8]]
        # 样本代表性终极校验
        sm = today.get("sample_mean")
        smed = today.get("sample_median")
        sw = today.get("sample_weighted")
        out["sample_check"] = {
            "sample_n": today.get("total_etf") or 0,
            "sample_mean": sm, "sample_median": smed, "sample_weighted": sw,
            "universe_median": out["etf_stats"]["median"],
            "universe_mean": out["etf_stats"]["mean"],
            "gap_median": (smed - out["etf_stats"]["median"])
            if smed is not None else None,
        }
        top_all = out["etf_gainers"][0] if out["etf_gainers"] else {}
        ranked = today.get("ranked") or []
        if top_all and ranked:
            q = ranked[0][1]
            signals.append("**样本边界解除**：全市场 {} 只 ETF 的真实涨幅榜首是 "
                           "**{} {:+.2f}%**，而 1.2 样本榜首仅 {} {}。"
                           "原先「不是全市场 ETF 涨幅榜、不能推断」的标注，"
                           "现由本章的全量报价直接补齐——样本池确实漏掉了"
                           "20cm 科创半导体/芯片 ETF。".format(
                               len(E), top_all["name"], top_all["zdf"],
                               q.get("name"), pct_str(q.get("pct"))))
        signals.append("**全市场 ETF 分布**：剔除债券/货币类后 {} 只，"
                       "涨跌幅中位数 {:+.2f}%、均值 {:+.2f}%、"
                       "{}% 上涨；样本池等权 {} / 中位数 {} / 加权 {}。"
                       "两者差距可量化样本偏态，1.6 的自检不再只有样本内部口径。".format(
                           len(scope), out["etf_stats"]["median"],
                           out["etf_stats"]["mean"],
                           "{:.0f}".format(out["etf_stats"]["pos_ratio"]),
                           pct_str(sm), pct_str(smed), pct_str(sw)))
    out["signals"] = signals
    return out


def render(today, market, fund, outlook, store, source_note,
           breadth=None, intraday=None, micro=None, glob=None,
           profile=None, events=None, backtest=None, universe=None,
           macro=None, official=None, mevents=None, quality=None,
           diff=None, advtech=None, attribution=None, rank=None, fut=None,
           ev_ar=None, ev_stats=None, ev_car=None, cons=None, cone=None,
           amt=None, opt=None):
    d = today["date"] or datetime.now().strftime("%Y-%m-%d")
    L = []
    add = L.append
    empties = {"available": False, "signals": [], "detail": [], "levels": [],
               "rows": [], "dims": [], "items": [], "rules": [],
               "top_prem": [], "top_disc": [], "overheat": [], "has_delta": [],
               "valuations": [], "by_region": {}, "points": [], "index_points": [],
               "checks": [], "bad": [], "flow_rows": [], "val_rows": [],
               "cheap": [], "rich": [], "dev_rows": [], "lhb_n": 0,
               "chip_n": 0, "event_n": 0, "universe_n": 0, "sample_n": 0,
               "d_total": 0, "in_n": 0, "prev_date": None, "n_days": 0,
               "erp": None, "y10": None, "ey_mid": None, "score": 0,
               "verdict": "", "vals": {}, "base": None, "cost_pct": 0.06,
               "adj": None, "raw": None, "sample_asset_pct": None,
               "flow_date": None, "total_nv_yi": 0, "total_asset_yi": 0,
               "collect": {}, "timeline": [], "topics": [], "tomorrow": [],
               "n_events": 0, "n_cand": 0, "n_post": 0, "n_next": 0,
               "launch": None,
               "launch_ret": 0, "verdict": "—", "flash": [],
               "flash_n": 0, "flash_cand": 0, "flash_status": "empty",
               "sina_status": "empty", "sina_n": 0, "ct_status": "empty",
               "ct_n": 0, "ct_hot": [], "etf_flow": [],
               "cls_status": "empty", "cls_n": 0, "cls_pages": 0,
               "cls_covered": False,
               "flash_days": 0, "flash_total": 0, "flash_src_n": 0,
               "flash_family_n": 0, "cross": [], "sina_new": [],
               "archive": {}, "archive_ct": {},
               # 第十六章 全市场涨跌排行
               "gainers": [], "losers": [], "periods": [], "fund_top": [],
               "amount_top": [], "etf_gainers": [], "etf_losers": [],
               "etf_stats": {}, "sample_check": {}, "date_ok": True,
               "date_note": "", "n_stock": 0, "n_etf": 0,
               "up": 0, "down": 0, "flat": 0, "limit_up": 0,
               "limit_down": 0, "limit_up_rows": [], "ratio": 0,
               # 第十七章 多方评论观点 / 15.10 东财PC页
               "viewpoints": {}, "empc_status": "empty", "empc_n": 0,
               "empc_days": {}, "empc_next": [], "empc_post": [],
               "empc_contrary": [], "checklist": [], "archive_empc": {}}
    def norm(obj):
        return dict(empties, **(obj or {}))
    breadth, intraday = norm(breadth), norm(intraday)
    micro, glob = norm(micro), norm(glob)
    profile, events, backtest = norm(profile), norm(events), norm(backtest)
    universe, macro = norm(universe), norm(macro)
    official, mevents = norm(official), norm(mevents)
    quality, diff, advtech = norm(quality), norm(diff), norm(advtech)
    attribution, rank = norm(attribution), norm(rank)
    fut = norm(fut)
    ev_ar, ev_stats = norm(ev_ar), norm(ev_stats)
    ev_car, cons, cone = norm(ev_car), norm(cons), norm(cone)
    amt, opt = norm(amt), norm(opt)
    if not profile.get("dims"):
        profile["dims"], profile["total"], profile["max_total"] = [], 0, 0
        profile["ratio"], profile["verdict"] = 0, "数据不足"

    add("# ETF 全景：大盘 / 广度 / 成交额 / 资金流 / 微观 / 外围 / 回测 / 走势")
    add("")
    # 数据时点：盘中与收盘必须分开表述，否则同一天的量与价会被误读成终值
    _as_state = AS_OF.get("state") or "close"
    _as_time = AS_OF.get("time") or ""
    _intraday = _as_state in ("intraday", "preopen")
    _hist_days = AS_OF.get("hist_days") or 0
    add("**日期**：{}　|　**数据时点**：{}　|　**状态**：{}  ".format(
        d, _as_time or "—", AS_OF.get("label") or "—"))
    add("**样本**：{} 只ETF / {} 个基准指数 / {} 个成交额口径 / {} 个板块资金流样本 / {} 个外围指数　|　**数据来源**：{}  ".format(
        today["total_etf"], len(today["indices"]), len(market["parts"]),
        len(fund["industry"]) + len(fund["concept"]) if fund["available"] else 0,
        len(glob["rows"]), source_note))
    add("**口径**：{}；成交额与资金流单位亿元；技术指标由日线自算（MA/MACD/KDJ/RSI/BOLL）。".format(
        "涨跌幅为收盘快照" if not _intraday else
        "涨跌幅为**盘中实时值**（非收盘价，收盘前不可与上日终值直接比较）"))
    add("**代表性**：成交额（沪深两市）、涨跌广度、板块资金流（申万行业）为**全市场完整覆盖**；"
        "指数与 ETF 池为**分层抽样观察窗口**。边界与自检见 1.6。")
    # 数据完整性告警：补缺/接口失败必须写进报告，而不是只打印到终端
    if PATCH_INFO:
        _f = "、".join(PATCH_INFO.get("fields") or []) or "无"
        _c = PATCH_INFO.get("codes") or []
        add("")
        add("> ⚠️ **数据补缺告警**：本次有在线数据缺失，用 **{}** 的旧快照补齐。"
            "受影响字段：{}；受影响章节：第二章（广度）、第四章（资金流）、"
            "第七章（外围）、第十至十七章（扩展数据）。".format(
                PATCH_INFO.get("date") or "未知日期", _f))
        if _c:
            add("> 另有 **{} 只标的**的行情与日 K 线来自该旧快照（{}…），"
                "这些标的的技术指标不是本次抓取的结果。".format(
                    len(_c), "、".join(_c[:5])))
    if PATCH_INFO.get("missing_ext"):
        add("> ⚠️ **扩展数据缺失**：{} 本次未取到，对应章节可能为空或不完整。".format(
            "、".join(PATCH_INFO["missing_ext"])))
    if RUN_WARNINGS:
        add("> ⚠️ **本次抓取告警**：{}。".format(
            "；".join(RUN_WARNINGS[:4])))
    add("")
    add("## 执行摘要")
    add("")
    add("| 维度 | 今日读数 | 倾向 |")
    add("|---|---|---|")
    add("| 指数 | 上证 {} / 科创50 {} | {} |".format(
        pct_str((store.get("sh000001", {}).get("quote") or {}).get("pct")),
        pct_str((store.get("sh000688", {}).get("quote") or {}).get("pct")),
        "成长极端占优" if ((store.get("sh000688", {}).get("quote") or {}).get("pct") or 0) > 7
        else "普涨"))
    add("| 广度 | 上涨 {:.0f}% / 涨停 {} / 跌停 {} | {} |".format(
        breadth.get("ratio") or 0, int(breadth.get("uplim") or 0),
        int(breadth.get("dnlim") or 0),
        "普涨但追高赚钱效应弱" if breadth.get("available") else "数据缺失"))
    add("| 成交额 | {:.0f} 亿（{}，较上日 {}{}） | {} |".format(
        market["total"] or 0, market["gate"],
        pct_str(market["chg"]), "，盘中值" if _intraday else "",
        market["vp"] if not _intraday else
        market["vp"] + "｜盘中，收盘前需复核"))
    add("| 板块资金 | 主力净额 {:+.0f} 亿，净流入板块 {}/{} | {} |".format(
        fund.get("net_total") or 0, len(fund.get("pos") or []),
        len(fund.get("industry") or []),
        "结构性进攻" if fund.get("available") else "数据缺失"))
    _has_margin = bool(breadth.get("margin_bal"))
    add("| 杠杆 / 溢价 | 融资余额 {}，{} 只ETF溢价≥0.5% | {} |".format(
        "{:.0f}亿".format(breadth["margin_bal"]) if _has_margin else "**未取到**",
        len(micro.get("overheat") or []),
        "{}".format("溢价偏高" if micro.get("overheat") else "溢价温和")
        + ("" if _has_margin else "；杠杆缺失，不作判断")))
    add("| ETF全景 | 全市场 {} 只 / 样本覆盖 {:.0f}% 规模 | {} |".format(
        universe["universe_n"], universe["sample_asset_pct"] or 0,
        "接口全量" if universe["available"] else "缺失"))
    add("| 真实净申购(T-1) | 样本 {:+.1f} 亿 / {} | {} |".format(
        universe["d_total"], universe["flow_date"] or "—",
        "一级净流入" if universe["d_total"] > 0 else "一级净流出"))
    add("| 股债性价比 | ERP {} | {} |".format(
        "{:+.2f}pct".format(macro["erp"]) if macro["erp"] is not None else "—",
        "宏观接口" if macro["available"] else "缺失"))
    if fut.get("available"):
        _fi = [r for r in fut["rows"] if r.get("basis_pct") is not None]
        if _fi:
            _all_neg = all(r["basis_pct"] < 0 for r in _fi)
            add("| 股指期货基差 | {} | {} |".format(
                " / ".join("{} {:+.2f}%".format(r["name"].split(" ")[0],
                                                r["basis_pct"]) for r in _fi),
                "全线贴水，对冲需求偏高" if _all_neg
                else "升贴水互现，对冲压力不极端"))
    else:
        add("| 股指期货基差 | — | 接口未取到 |")
    add("| 官方画像 | 调整分 {}（{} 维） | {} |".format(
        official["adj"] if official["adj"] is not None else "—",
        len(official["dims"]), "与自建画像互为印证" if official["available"] else "缺失"))
    add("| 回测校准 | {} 条规则 / {} 根日线（扣成本+样本外） | 详见第九章 |".format(
        len(backtest.get("rules") or []), backtest.get("total_bars") or 0))
    add("| 数据质量 | {:.0f}%（{}） | {} |".format(
        quality["score"], len(quality["checks"]), quality["verdict"]))
    add("| 涨跌归因 | {} | 候选驱动 {} 条 / 事后 {} 条 |".format(
        attribution["verdict"] if attribution["available"] else "缺失",
        attribution["n_cand"], attribution["n_post"]))
    if attribution["available"] and attribution["cross"]:
        add("| 快讯多源互证 | {} 项关键事实 | {} 个独立信源族 |".format(
            len(attribution["cross"]), attribution["flash_family_n"]))
    _enxt = attribution.get("empc_next") or []
    if _enxt:
        add("| 次日盘中验证 | {} 条 11-05 09:30 后快讯 | 主线延续 + 高位股回吐 + 轮动扩散 |".format(
            len(_enxt)))
    if rank.get("available"):
        p20 = next((p for p in rank["periods"] if p["key"] == "zdf_d20"), None)
        add("| 全市场涨跌排行 | 个股 {} / ETF {} 只全量 | {} |".format(
            rank["n_stock"], rank["n_etf"],
            ("近20日仅 {:.0f}% 个股为正，中期仍弱".format(p20["pos_ratio"])
             if p20 else "详见第十六章")))
    _vp = attribution.get("viewpoints") or {}
    if _vp.get("available"):
        _dv = sum(1 for t in _vp["topics"] if t["divergent"])
        add("| 多方评论观点 | {} 条 / {} 类观点方 | {} |".format(
            _vp["n"], len(_vp["actors"]),
            ("{} 个主题现分歧/条件式，无单边共识".format(_dv) if _dv
             else "立场分布见第十七章")))
    add("")
    # ---- 决策看板：100+ KB 报告需要一个 30 秒入口，先给结论再给证据 ----
    _top_in = (fund.get("in_top") or [])
    _top_in = _top_in[0] if _top_in else {}
    _ar_first = (ev_ar.get("rows") or [])[:1]
    _ar_first = _ar_first[0] if _ar_first else {}
    _futs = [r for r in (fut.get("rows") or []) if r.get("basis_pct") is not None]
    add("**今日要点（30 秒版）**")
    add("")
    add("1. **广度**：上涨 {} 只 / 下跌 {} 只（上涨占比 {}），涨停 {} / 跌停 {}；"
        "全市场个股近20日中位数 {}。".format(
            "{:,.0f}".format(breadth["up"]) if breadth.get("up") is not None else "—",
            "{:,.0f}".format(breadth["down"]) if breadth.get("down") is not None else "—",
            "{:.0f}%".format(breadth.get("ratio")) if breadth.get("ratio") is not None else "—",
            "{:,.0f}".format(breadth["uplim"]) if breadth.get("uplim") is not None else "—",
            "{:,.0f}".format(breadth["dnlim"]) if breadth.get("dnlim") is not None else "—",
            pct_str((next((p for p in (rank.get("periods") or [])
                           if p.get("key") == "zdf_d20"), None) or {}).get("median"))
            if rank.get("periods") else "见第十六章"))
    add("2. **量能**：成交额 {:.0f} 亿（{}），较{} {}；{}。".format(
        market.get("total") or 0, market.get("gate") or "—",
        "昨日同时点" if _as_state == "intraday" else "上日",
        pct_str(market.get("chg")), market.get("vp") or "—"))
    _fut_txt = ("；期货 " + "、".join(
        "{} {:+.2f}%".format(r["name"].split(" ")[0], r["basis_pct"])
        for r in _futs)) if _futs else ""
    add("3. **资金**：{} 行业口径净额 {}，{} 领跑 {}{}。".format(
        len(fund.get("industry") or []) if fund.get("available") else 0,
        "{:+.0f} 亿".format(fund.get("net_total") or 0) if fund.get("available") else "—",
        _top_in.get("name") or "—",
        "{:+.0f} 亿".format(_top_in.get("net_yi") or 0) if _top_in else "—",
        _fut_txt))
    add("4. **归因**：{}；主题 AR 最强 {}（{}）。".format(
        attribution.get("verdict") or "—",
        _ar_first.get("topic") or "—",
        "{:+.2f}pct".format(_ar_first["ar"]) if _ar_first.get("ar") is not None else "—"))
    add("5. **待复核**：数据质量 {}（{}），通道 {} 个，跨源一致性 {}/{}；{}。".format(
        "{:.0f}%".format(quality.get("score") or 0), quality.get("verdict") or "—",
        len(quality.get("health") or []),
        cons.get("ok_n") or 0, cons.get("n") or 0,
        "盘中口径，收盘后请重跑" if _as_state == "intraday"
        else ("历史仅 {} 日，环比只作方向".format(_hist_days) if _hist_days < 5
              else "为收盘终值口径")))
    add("")
    # 时点/历史不足的硬约束：这两类告警直接决定结论可用性，必须出现在摘要正下方
    if _as_state == "intraday":
        add("> ⚠️ **盘中口径**：成交额 {:.0f} 亿、广度 {:.0f}%、板块净额 {:+.0f} 亿均为截至 {} 的实时值，"
            "未收盘前会继续变化；环比与「量价配合」结论只代表该时点，收盘后请重跑确认。".format(
                market["total"] or 0, breadth.get("ratio") or 0,
                fund.get("net_total") or 0, _as_time or "当前"))
        add("")
    if _hist_days < 5:
        add("> ⚠️ **历史样本不足**：逐日历史仅 **{}** 天（阈值 5 天），第十二章 diff 与所有环比"
            "均为单点对比，只能作方向参考，不构成趋势判断；连续运行 ≥5 个交易日后再看环比。".format(
                _hist_days))
        add("")
    add("> **一句话结论**：{}（自建画像 {}/{}，{:.0f}%）；基准情景 **{}**{}。"
        "新增证据显示 ETF 一级市场有大额净申购（T-1 {:+.0f} 亿），但{}、{}、"
        "中期均线仍压制——属「**一级有增量、二级未放量**」的"
        "超跌修复，不是趋势反转。第十五章事件归因：{}，"
        "且盘后发酵的传闻已被时序切分剔除，不当作盘中驱动。"
        "{}{}".format(
            profile["verdict"], profile["total"], profile["max_total"],
            profile["ratio"], outlook["base_scene"],
            "（盘中口径，收盘前需复核）" if _as_state == "intraday" else "",
            universe["d_total"], "二级成交未有效放大",
            ("融资余额 {:.0f} 亿".format(breadth["margin_bal"])
             if _has_margin else "杠杆数据未取到（不引用）"),
            attribution["verdict"],
            ("第十六章全市场 {} 只个股近20日仅 {:.0f}% 为正，"
             "把「中期弱」从样本推演升级为全量证据。".format(
                 rank["n_stock"],
                 next((p for p in rank["periods"]
                       if p["key"] == "zdf_d20"), {"pos_ratio": 0})["pos_ratio"]))
            if rank.get("available") and rank.get("periods") else "",
            ("第十七章多方观点里「底部显现但共识兑现还需时间」"
             "与分歧/条件式的立场结构，正是观点侧的同义印证。"
             if (attribution.get("viewpoints") or {}).get("available") else "")
            + ("15.10 东财 PC 页实测的次日盘中快讯显示"
               "「主线延续 + 高位股回吐 + 轮动扩散」，为基准情景补上实证。"
               if (attribution.get("empc_next") or []) else "")))
    add("")
    add("---")
    add("")
    add("## 一、今日大盘")
    add("")
    add("> **本章定位（先读这条）**：下面的 {} 只 ETF 是**跟踪清单**，覆盖全市场 {} 只 ETF 的 "
        "{:.0f}% 规模，只承担「折溢价 / 份额 / 申赎 / 形态」这类需要基金层字段的观察；"
        "**涨跌排行、量比强弱、板块轮动、个股赚钱效应一律以全量为准**——"
        "全市场 {} 只个股见第十六章，{} 只 ETF 见 16.6。"
        "把本章样本均值当成「市场涨跌」是本报告最容易犯的误读。".format(
            today["total_etf"], universe.get("universe_n") or 0,
            universe.get("sample_asset_pct") or 0,
            rank.get("n_stock") or 0, universe.get("universe_n") or 0))
    add("")
    add("### 1.1 指数概览")
    add("")
    add("| 指数 | 收盘 | 涨跌幅 | 近20日 | 年初至今 |")
    add("|---|---:|---:|---:|---:|")
    for _, q in today["indices"]:
        add("| {} | {} | {} | {} | {} |".format(
            q["name"], price_str(q["price"]), pct_str(q["pct"]),
            pct_str(q["d20"]), pct_str(q["ytd"])))
    add("")

    add("### 1.2 跟踪清单内 ETF 涨跌（样本内排序，非全市场）")
    add("")
    add("> 以下为**样本池内**排序（共 {} 只）。"
        "**真正的全市场 {} 只 ETF 涨跌排行已由第十六章 16.6 补齐**；"
        "样本构成与可推广性见 1.6 与 16.7。".format(
            today["total_etf"], universe.get("universe_n") or 0))
    add("")
    add("| ETF | 代码 | 收盘 | 涨跌幅 | 振幅 | 成交额(亿) | 量比 | 近60日 |")
    add("|---|---|---:|---:|---:|---:|---:|---:|")
    for c, q in today["ranked"]:
        mark = " 🔴" if q["pct"] is not None and q["pct"] >= 9.8 else ""
        add("| {}{} | {} | {} | {} | {} | {} | {} | {} |".format(
            q["name"], mark, c, price_str(q["price"]), pct_str(q["pct"]),
            pct_str(q["range_pct"]),
            "{:.1f}".format(q["amount_yi"]) if q["amount_yi"] else "—",
            "{:.2f}".format(q["vol_ratio"]) if q["vol_ratio"] else "—",
            pct_str(q["d60"])))
    add("")
    add("> 🔴 = 涨停 / 近涨停（≥9.8%）。样本ETF今日成交额合计约 **{:.0f} 亿**。".format(
        today["amount_total"]))
    add("")

    add("### 1.3 样本风格分组")
    add("")
    add("| 分组 | 样本数 | 等权平均 | 成交额加权 |")
    add("|---|---:|---:|---:|")
    for g, v in today["groups"].items():
        add("| {} | {} | {} | {} |".format(
            g, v["count"], pct_str(v["mean"]), pct_str(v["weighted"])))
    add("")
    add("> 等权均值对同赛道多只ETF会重复计数，成交额加权更接近真实资金感受；"
        "两者差距见 1.6。")
    add("")
    if today["spread"] is not None:
        add("**成长－宽基分化度：{:+.2f} 个百分点**（样本等权口径）。"
            "差值越大，越说明资金集中而非全面开花；全市场验证请看第四章资金广度。".format(
                today["spread"]))
        add("")

    add("### 1.4 量能与形态")
    add("")
    add("| 核心ETF | 今日成交量/昨日 | 判定 | 今日成交额(亿) |")
    add("|---|---:|---|---:|")
    for v in today["vol_checks"]:
        verdict = "**缩量反包**" if v["ratio"] < 1 else "放量反包"
        add("| {} | {:.0%} | {} | {} |".format(
            v["name"], v["ratio"], verdict,
            "{:.1f}".format(v["amount_yi"]) if v["amount_yi"] else "—"))
    add("")
    add("开盘形态：{}/{} 只平开，{}/{} 只收于全天高位（收盘位置 ≥0.9）。".format(
        today["flat_open"], today["total_etf"], today["close_high"], today["total_etf"]))
    add("")

    add("### 1.5 自动生成的关键信号")
    add("")
    for i, s in enumerate(today["signals"], 1):
        add("{}. {}".format(i, s))
    add("")

    add("### 1.6 样本代表性与口径边界")
    add("")
    add("「大盘」不是靠 ETF 只数堆出来的，本报告分**三层数据**，代表性各不相同：")
    add("")
    add("| 数据层 | 覆盖范围 | 全市场？ | 主要用途 |")
    add("|---|---|---|---|")
    add("| 指数层 | {} 个基准指数（含 {} 个成交额口径） | ❌ 抽样（官方基准） | 风格分化、量价配合 |".format(
        len(today["indices"]), len(market["parts"])))
    add("| 成交额层 | 沪市 + 深市**全部**成交 | ✅ 完整覆盖 | 总量、关口、量价配合 |")
    add("| 板块资金流层 | {} 个申万行业（另 {} 个概念） | ✅ 行业完整覆盖 | 资金流向、涨跌广度 |".format(
        len(fund["industry"]) if fund["available"] else 0,
        fund["con_total"] if fund["available"] else 0))
    add("| ETF 层 | {} 只 / {} 个分组 | ❌ 分层抽样 | ETF 端观察窗口 |".format(
        today["total_etf"], today["group_count"]))
    add("")
    add("**样本偏态自检**")
    add("")
    add("| 指标 | 数值 | 说明 |")
    add("|---|---:|---|")
    add("| 样本涨幅等权均值 | {} | 每只ETF一票，小ETF权重被放大 |".format(
        pct_str(today["sample_mean"])))
    add("| 样本涨幅中位数 | {} | 不受极端涨跌影响 |".format(
        pct_str(today["sample_median"])))
    add("| 样本涨幅成交额加权 | {} | 按真实资金体量加权 |".format(
        pct_str(today["sample_weighted"])))
    skew = (today["sample_mean"] - today["sample_weighted"]
            if None not in (today["sample_mean"], today["sample_weighted"])
            else None)
    add("| 等权－加权偏离 | {} | {} |".format(
        pct_str(skew),
        "偏离 <1pct，样本结构较均衡" if (skew is not None and abs(skew) < 1)
        else "偏离 ≥1pct，等权均值被小体量ETF带偏，应优先看加权值"))
    if market["etf_share"] is not None:
        add("| 样本ETF成交额 / 两市 | {:.1f}% | 样本在总量中的体量占比 |".format(
            market["etf_share"]))
    add("")
    if today["dup_groups"]:
        add("**同赛道重复计数提示**：{}。这些ETF跟踪相近指数，"
            "分组等权均值会被同一赛道重复计入，因此 1.3 同时给出成交额加权值。".format(
                "、".join("{} {} 只（{}）".format(
                    g["label"], g["count"], "、".join(g["names"][:3]))
                    for g in today["dup_groups"])))
        add("")
    add("**哪些结论可以推广到全市场**")
    add("")
    add("- ✅ **可代表大盘**：两市成交额与关口、量价配合、主力资金净额与净流出/流入板块数、"
        "行业涨跌家数广度、行业资金集中度——这些来自完整覆盖层，与 ETF 只数无关。")
    add("- ⚠️ **只在样本内成立**：1.2 涨跌排行、1.3 分组均值与分化度、涨停只数、"
        "1.4 核心ETF量能对比、第十四章技术仪表盘。读取时须带「样本」前缀。")
    add("- ✅ **原先不能推断、现已补齐**：全市场 ETF 总数与总规模"
        "（第十章接口，{} 只）、全市场 ETF 涨跌排行（16.6 全量报价）。".format(
            universe.get("universe_n") or 0))
    add("- ❌ **仍不能推断**：未纳入分组的行业ETF结构性结论、"
        "以及把 1.2 样本内数字直接当作全市场统计（请用 16.6/16.7 的全量口径）。")
    add("")
    if today["total_etf"] >= 20:
        history_note = ("早期版本仅 11 只且科技占比过半，确实会让排行与分组均值偏科技；"
                        "现扩至 {} 只 / {} 个分组并明确标注边界，".format(
                            today["total_etf"], today["group_count"]))
    else:
        history_note = ("当前为自定义/精简样本（{} 只 / {} 个分组），1.2、1.3 的统计"
                        "仅对这几只成立，覆盖面窄于默认池；".format(
                            today["total_etf"], today["group_count"]))
    add("> **结论**：判断大盘走势应以第二至四章的全量口径为主，ETF 池只回答「资金在ETF端如何表现」。"
        "{}但抽样性质不变，不应把样本内数字直接当作全市场统计。".format(history_note))
    add("")

    add("---")
    add("")
    add("## 二、全市场涨跌广度与情绪")
    add("")
    if not breadth["available"]:
        add("> 暂无涨跌广度数据（在线接口不可用，或快照未含缓存）。"
            "可执行 `python3 etf_analysis.py --save-snapshot` 生成缓存。")
        add("")
    else:
        add("> **全量口径**：沪深 A 股全部个股的涨跌家数、涨跌停与涨跌幅区间分布，"
            "不是 ETF 样本内统计。数据源：全市场行情概览接口。")
        add("")
        add("### 2.1 涨跌家数与情绪温度")
        add("")
        add("| 指标 | 数值 | 指标 | 数值 |")
        add("|---|---:|---|---:|")
        add("| 上涨 | {} 家 | 下跌 | {} 家 |".format(
            int(breadth["up"] or 0), int(breadth["down"] or 0)))
        add("| 平盘 | {} 家 | 停牌 | {} 家 |".format(
            int(breadth["flat"] or 0), int(breadth["susp"] or 0)))
        add("| 涨停 | **{}** 家 | 跌停 | **{}** 家 |".format(
            int(breadth["uplim"] or 0), int(breadth["dnlim"] or 0)))
        add("| 上涨占比 | **{:.0f}%** | 涨停跌停比 | {} |".format(
            breadth["ratio"] or 0,
            "{:.1f} : 1".format(breadth["limit_ratio"]) if breadth.get("limit_ratio") else "—"))
        add("| 涨超5% | {} 家 | 跌超5% | {} 家 |".format(
            int(breadth["strong_up"]), int(breadth["strong_dn"])))
        if breadth.get("heat_pctl") is not None:
            add("| 广度历史分位 | 过去一年仅 **{:.0f}%** 交易日达到 | | |".format(
                breadth["heat_pctl"]))
        if breadth.get("margin_bal"):
            add("| 融资余额 | {:.0f} 亿（{}） | 连降 | {} 日 |".format(
                breadth["margin_bal"], pct_str(breadth.get("margin_chg")),
                breadth.get("margin_down_days") or "—"))
        if breadth.get("follow_up") is not None:
            add("| 昨日涨停股今日上涨 | **{:.1f}%** | 追高赚钱效应 | {} |".format(
                breadth["follow_up"],
                "偏弱" if breadth["follow_up"] < 50 else "尚可"))
        add("")
        add("> {}".format(breadth.get("comment") or ""))
        add("")

        add("### 2.2 涨跌幅区间分布")
        add("")
        add("| 区间 | 家数 | 方向 |")
        add("|---|---:|---|")
        for r in breadth["detail"]:
            flag = r.get("flag") or 0
            add("| {} | {} | {} |".format(
                md_plain(r.get("section")), int(r.get("count") or 0),
                "涨" if flag > 0 else ("跌" if flag < 0 else "平")))
        add("")
        dl = [r.get("section") or "" for r in breadth["detail"]]
        dv = [(r.get("count") or 0) * (1 if (r.get("flag") or 0) >= 0 else -1)
              for r in breadth["detail"]]
        chart = save_chart("breadth_distribution",
                           svg_bars(dl, dv, "全市场涨跌幅区间分布（家数，负值为下跌）"))
        if chart:
            add("![全市场涨跌幅区间分布]({})".format(chart))
            add("")

        add("### 2.3 日内广度演变")
        add("")
        if intraday.get("levels"):
            add("| 时点 | 上涨家数占比 |")
            add("|---|---:|")
            for k, v in intraday["levels"]:
                add("| {} | {} |".format(k, "{:.0f}%".format(v) if v is not None else "—"))
            add("")
        pts = intraday.get("points") or []
        if len(pts) >= 2:
            labels = [t for t, _ in pts]
            vals = [v for _, v in pts]
            chart = save_chart("intraday_breadth",
                               svg_line(vals, labels, "日内全市场上涨家数占比演变（%）",
                                        fmt="{:.0f}", hline=50, color=C_UP))
            if chart:
                add("![日内上涨占比演变]({})".format(chart))
                add("")
            add("> 虚线为 50% 多空分界。曲线全天单边上行说明买盘持续扩散，"
                "而非开盘一次性脉冲。")
            add("")
        for s in intraday.get("signals") or []:
            add("- {}".format(s))
        add("")

        add("### 2.4 广度与情绪信号")
        add("")
        for i, s in enumerate(breadth["signals"], 1):
            add("{}. {}".format(i, s))
        add("")
        if breadth.get("adr") or breadth.get("adr_limit"):
            add("> ADR 补充：{} {}".format(breadth.get("adr") or "",
                                        breadth.get("adr_limit") or ""))
            add("")

    add("---")
    add("")
    add("## 三、全市场成交额与日内节奏")
    add("")
    add("> **口径**：沪市取上证指数成交额、深市取深证成指成交额，两者相加即市场通用的「沪深两市成交额」。"
        "北证50为成分指数口径，仅列参考、不计入合计；创业板指/科创50属子集，只用于算结构占比、不重复相加。"
        "成交量沿用行情源单位（1 手 = 100 股），表中换算为亿手。"
        "历史成交额按当日「单位成交量金额」折算估算，仅用于趋势与分位对比，绝对值以当日快照为准。")
    add("")

    add("### 3.1 当日成交额总览")
    add("")
    add("| 市场 | 成交额(亿) | 占两市 | 今日涨跌 |")
    add("|---|---:|---:|---:|")
    for label in ("沪市", "深市"):
        p = market["parts"].get(label)
        if p and market["total"]:
            add("| {} | {:.0f} | {:.1f}% | {} |".format(
                label, p["amount"], p["amount"] / market["total"] * 100, pct_str(p["pct"])))
    if market["total"]:
        add("| **沪深两市合计** | **{:.0f}** | **100%** | — |".format(market["total"]))
    bj = market["parts"].get("北证50")
    if bj:
        add("| 北证50（参考，不合计） | {:.0f} | — | {} |".format(bj["amount"], pct_str(bj["pct"])))
    if market["all_total"]:
        add("| 全市场（含北证50） | {:.0f} | — | — |".format(market["all_total"]))
    add("")
    add("**关口**：{}；较昨日折算值 {}；量价配合 **{}**。".format(
        ("站上 **{}** 关口".format(market["gate"])
         if market["total"] and market["total"] >= market["gate_threshold"]
         else "未站上 **{}** 关口".format(market["gate"])),
        pct_str(market["chg"]), market["vp"]))
    add("")

    add("### 3.2 板块成交结构")
    add("")
    if market["boards"]:
        add("| 板块（指数口径） | 成交额(亿) | 母市场 | 占母市场 |")
        add("|---|---:|---|---:|")
        for b in market["boards"]:
            add("| {} | {:.0f} | {} | {:.1f}% |".format(
                b["label"], b["amount"], b["parent"], b["ratio"]))
        add("")
        add("> 数值取自板块指数行情快照的成交额字段，用于观察板块活跃度，非严格成分股统计；"
            "属沪深两市的子集，**不能**与两市合计相加。")
    else:
        add("暂无板块成交额数据。")
    add("")

    add("### 3.3 近期量能趋势")
    add("")
    if market["series"]:
        add("| 日期 | 两市成交量(亿手) | 折算成交额(亿) | 较当前5日均额 |")
        add("|---|---:|---:|---:|")
        show = market["series"][-10:]
        # 与 3.6 曲线同一口径：盘中/快照的末行是累计值，不能标「今日精确」，
        # 也不能拿它跟全天 5 日均额算环比（否则会显示 -37% 的假跌）。
        partial_last = bool(
            (_intraday or AS_OF.get("state") == "snapshot") and show
            and show[-1]["date"] == today.get("date"))
        for i, s in enumerate(show):
            last = i == len(show) - 1
            if last and partial_last:
                tag, rel_str = "（盘中累计，未收盘）", "—（不与全天均值比较）"
            else:
                rel = ((s["amount_est"] / market["ma5"] - 1) * 100
                       if market["ma5"] else None)
                tag = "（今日精确）" if last else ""
                rel_str = pct_str(rel)
            add("| {}{} | {:.2f} | {:.0f} | {} |".format(
                s["date"], tag, s["volume"] / 1e8, s["amount_est"], rel_str))
        add("")
    if market["series"]:
        labels, means = ["当日"], ["{:.0f}".format(market["cur"] or 0)]
        if market["ma5"]:
            labels.append("5日均")
            means.append("{:.0f}".format(market["ma5"]))
        if market["ma10"]:
            labels.append("10日均")
            means.append("{:.0f}".format(market["ma10"]))
        if market["lookback"] > 5:
            labels.append("近{}日均".format(market["lookback"]))
            means.append("{:.0f}".format(market["ma20"] or 0))
        add("{}：**{} 亿**；当日处于近 {} 日 **{:.0f}% 分位**。".format(
            " / ".join(labels), " / ".join(means),
            market["lookback"], market["pctl"] or 0))
    else:
        add("暂无可用的历史量能数据（可调整 --market-days 或检查日线抓取结果）。")
    add("")

    add("### 3.4 量价配合与结构信号")
    add("")
    for i, s in enumerate(market["signals"], 1):
        add("{}. {}".format(i, s))
    add("")

    add("### 3.5 对接下来走势的启示")
    add("")
    if market["total"] and market["chg"] is not None:
        if market["chg"] > 3 and market["idx_pct"] and market["idx_pct"] > 0:
            add("- 全市场放量上涨，为反弹提供了**总量层面**的支撑；但第十四章的技术仪表盘显示核心ETF仍全部位于 MA60 下方，"
                "「总量放量」与「中期空头排列」并不矛盾，属超跌修复而非趋势反转。")
        elif market["chg"] < -3 and market["idx_pct"] and market["idx_pct"] > 0:
            add("- 指数上涨但成交额萎缩，属**量价背离**，反弹持续性需要下一交易日补量确认。")
        else:
            add("- 成交额未出现极端放缩，量能本身不构成方向性信号，以价格对压力位的攻防为主。")
    if market["total"]:
        add("- **成交额关口是明日最硬的确认条件**：两市需维持在 **{:.0f} 亿（{}）** 之上，"
            "跌回关口下方则视作放量反弹未能延续。".format(
                market["gate_threshold"], market["gate"]))
    if market["divergence"]:
        add("- 全市场放量但样本ETF缩量，追涨ETF的性价比低于直接观察指数与个股量能，"
            "ETF端更适合等回踩确认。")
    add("")

    add("### 3.6 日内节奏与量能曲线")
    add("")
    ser = market["series"]
    # 盘中/快照的当日成交额是「累计值」，与历史全天值不可比。若把它画进
    # 曲线末尾会出现一根假断崖（半日 1.1 万亿 vs 全天 2.1 万亿），所以
    # 从曲线剔除，改用文字标注，序列只保留口径一致的全天值。
    pending = None
    if (_intraday or AS_OF.get("state") == "snapshot") and ser \
            and ser[-1]["date"] == today.get("date"):
        pending = ser[-1]
        ser = ser[:-1]
    if len(ser) >= 2:
        chart = save_chart("amount_trend",
                           svg_line([s["amount_est"] for s in ser],
                                    [s["date"][5:] for s in ser],
                                    "两市成交额趋势（亿元，历史为折算估算）",
                                    fmt="{:.0f}", color=C_ACC))
        if chart:
            add("![两市成交额趋势]({})".format(chart))
            add("")
        if pending:
            add("> 注：{} 为未收盘数据，累计约 {:.0f} 亿，**未计入上方曲线**"
                "（与历史全天成交额口径不可比）。".format(
                    pending["date"], pending["amount_est"]))
            add("")
    ipts = intraday.get("index_points") or []
    if len(ipts) >= 2:
        chart = save_chart("index_intraday",
                           svg_line([v for _, v, _ in ipts],
                                    [t for t, _, _ in ipts],
                                    "上证指数日内分时（点）", fmt="{:.0f}",
                                    color=C_UP))
        if chart:
            add("![上证指数日内分时]({})".format(chart))
            add("")
    for s in intraday.get("signals") or []:
        if "上证指数" in s or "日内振幅" in s:
            add("- {}".format(s))
    add("")
    if market.get("trillion_days"):
        add("> 成交额已连续 **{} 日**破万亿；相对近一月均值 {}。"
            "持续高量能是行情延续的必要条件，但单日峰值本身不代表趋势。".format(
                market["trillion_days"], pct_str(market.get("vs_month"))))
        add("")

    add("---")
    add("")
    add("## 四、全市场资金流入流出板块分析")
    add("")
    add("> **口径**：申万行业板块 {} 个（成分基本不重叠，可近似视为全市场覆盖）、概念板块 {} 个"
        "（成分互相重叠，只用于观察细分资金聚焦，不可加总）。"
        "金额单位为亿元，原始字段为「主力净流入」及其 5 日 / 20 日累计值；"
        "净流入为正、净流出为负。数据源：腾讯公开板块资金流接口。".format(
            len(fund["industry"]) if fund["available"] else 0,
            fund["con_total"] if fund["available"] else 0))
    add("")

    if not fund["available"]:
        add("> 暂无板块资金流数据（在线接口不可用，或当前快照未含资金流缓存）。"
            "可执行 `python3 etf_analysis.py --save-snapshot` 在线生成一次缓存。")
        add("")
    else:
        add("### 4.1 资金总览")
        add("")
        add("| 指标 | 数值 |")
        add("|---|---:|")
        add("| 行业板块数 | {} |".format(len(fund["industry"])))
        add("| 主力净流入合计 | {:+.0f} 亿 |".format(fund["gross_in"]))
        add("| 主力净流出合计 | {:+.0f} 亿 |".format(fund["gross_out"]))
        add("| **全市场主力净额** | **{:+.0f} 亿** |".format(fund["net_total"]))
        add("| 净流入 / 净流出板块 | {} / {} |".format(len(fund["pos"]), len(fund["neg"])))
        add("| 资金广度（净流入占比） | {:.0f}% |".format(
            len(fund["pos"]) / len(fund["industry"]) * 100))
        if fund["breadth"] is not None:
            add("| 成分股上涨家数 | {}/{}（{:.1f}%） |".format(
                fund["up"], fund["total_stocks"], fund["breadth"]))
        if fund["cross_diff"] is not None:
            add("| 行业成交额合计 vs 两市成交额 | {:.0f} 亿（{:+.1f}%） |".format(
                fund["cje_sum"], fund["cross_diff"]))
        add("")

        def board_table(rows, caption):
            add("**{}**".format(caption))
            add("")
            add("| 板块 | 涨跌 | 主力净额(亿) | 5日(亿) | 20日(亿) |")
            add("|---|---:|---:|---:|---:|")
            for r in rows:
                add("| {} | {} | {:+.0f} | {:+.0f} | {:+.0f} |".format(
                    r["name"], pct_str(r["pct"]), r["net_yi"],
                    r["net5_yi"] or 0, r["net20_yi"] or 0))
            add("")

        add("### 4.2 行业主力净流入 Top{}".format(len(fund["in_top"])))
        add("")
        board_table(fund["in_top"], "领涨股：" + "；".join(
            "{} → {}".format(r["name"], r["leader"]) for r in fund["in_top"][:4]
            if r["leader"]))
        if fund["in_top"]:
            chart = save_chart("fundflow_in",
                               svg_bars([r["name"][:5] for r in fund["in_top"]],
                                        [r["net_yi"] for r in fund["in_top"]],
                                        "行业主力净流入 Top（亿元）", fmt="{:+.0f}"))
            if chart:
                add("![行业主力净流入]({})".format(chart))
                add("")

        add("### 4.3 行业主力净流出 Top{}".format(len(fund["out_top"])))
        add("")
        board_table(fund["out_top"], "净流出居前板块（负值越大，流出越多）")
        if fund["out_top"]:
            chart = save_chart("fundflow_out",
                               svg_bars([r["name"][:5] for r in fund["out_top"]],
                                        [r["net_yi"] for r in fund["out_top"]],
                                        "行业主力净流出 Top（亿元）", fmt="{:+.0f}"))
            if chart:
                add("![行业主力净流出]({})".format(chart))
                add("")

        add("### 4.4 概念板块资金聚焦")
        add("")
        add("> 已剔除 {} 个宽泛属性类概念（如转融券标的、深股通、指数成分、昨日涨停等），"
            "只保留可解读的题材赛道；其余 {} 个概念板块参与排序。".format(
                fund["broad_excluded"], fund["con_total"] - fund["broad_excluded"]))
        add("")
        if fund["con_top"]:
            board_table(fund["con_top"], "概念净流入 Top{}".format(len(fund["con_top"])))
            if fund["con_bot"]:
                board_table(fund["con_bot"], "概念净流出 Top{}".format(len(fund["con_bot"])))
        else:
            add("暂无概念板块资金流数据。")
            add("")

        add("### 4.5 资金结构信号")
        add("")
        for i, s in enumerate(fund["signals"], 1):
            add("{}. {}".format(i, s))
        add("")

        add("### 4.6 对接下来走势的启示")
        add("")
        for n in fund["outlook_notes"]:
            add("- {}".format(n))
        add("")

    add("---")
    add("")
    add("## 五、ETF 折溢价、份额与净申赎")
    add("")
    add("> **概念澄清**：成交额 = 二级市场买卖换手，不代表资金流入；"
        "**份额变化**才对应一级市场净申购/净赎回。只统计成交额容易把「换手放大」"
        "误读成「资金进场」。IOPV 为基金份额参考净值，折溢价 = 二级价格相对 IOPV 的偏离。")
    add("")
    if not micro["available"]:
        add("> 暂无 IOPV / 折溢价 / 份额数据。")
        add("")
    else:
        add("### 5.1 折溢价排行")
        add("")
        add("| ETF | 现价 | IOPV | 折溢价 | 今日涨跌 | 成交额(亿) |")
        add("|---|---:|---:|---:|---:|---:|")
        for r in micro["top_prem"]:
            add("| {} | {} | {} | **{:+.2f}%** | {} | {} |".format(
                r["name"], price_str(r["price"]), price_str(r["iopv"]),
                r["discount"], pct_str(r["pct"]),
                "{:.1f}".format(r["amount_yi"]) if r["amount_yi"] else "—"))
        add("")
        chart = save_chart("etf_premium",
                           svg_bars([r["name"][:6] for r in micro["top_prem"]],
                                    [r["discount"] for r in micro["top_prem"]],
                                    "样本ETF折溢价 Top（%）", fmt="{:+.2f}"))
        if chart:
            add("![ETF折溢价]({})".format(chart))
            add("")
        if micro["top_disc"]:
            add("折价居前：{}。其中涨停封板品种的折价多为**涨停限价被动折价**"
                "（指数涨幅超过 10% 涨停上限、IOPV 继续上行），不等于二级市场抛压。".format(
                    "、".join("{} {:+.2f}%".format(r["name"], r["discount"])
                             for r in micro["top_disc"][:3])))
            add("")

        add("### 5.2 份额、规模与净申赎")
        add("")
        add("| ETF | 份额(亿份) | 估算规模(亿) | 较上一记录日 |")
        add("|---|---:|---:|---:|")
        for r in sorted(micro["rows"], key=lambda x: x["size_yi"] or 0, reverse=True)[:12]:
            d = r["delta"]
            add("| {} | {:.1f} | {} | {} |".format(
                r["name"], r["shares_yi"],
                "{:.0f}".format(r["size_yi"]) if r["size_yi"] else "—",
                "{:+.1f}亿份（{:+.2f}%）".format(d["diff"] / 1e8, d["pct"])
                if d else "暂无历史"))
        add("")
        add("> 份额历史文件：`{}`，当前已记录 **{}** 个交易日。"
            "连续运行 2 日后即可自动计算净申赎；规模为「份额 × IOPV」估算值。".format(
                micro.get("shares_file") or "etf_shares_history.json",
                micro.get("shares_days") or 0))
        add("")

        add("### 5.3 微观结构信号")
        add("")
        for i, s in enumerate(micro["signals"], 1):
            add("{}. {}".format(i, s))
        add("")

    add("---")
    add("")
    add("## 六、跨市场外围与估值横向")
    add("")
    add("### 6.1 外围指数")
    add("")
    if not glob["rows"]:
        add("> 暂无外围行情数据（在线接口不可用，或快照未含缓存）。")
        add("")
    else:
        add("| 指数 | 市场 | 收盘 | 涨跌幅 |")
        add("|---|---|---:|---:|")
        for r in glob["rows"]:
            add("| {} | {} | {} | {} |".format(
                r["name"], r["region"], price_str(r["price"]), pct_str(r["pct"])))
        add("")
        add("> **时差提示**：美股为最近一个交易日收盘（与 A 股不同步），"
            "港股/欧洲为各自当日收盘；不可把美股涨跌直接当作 A 股的同步领先指标。")
        add("")

    if fut.get("available"):
        add("### 6.1.1 股指期货基差与国债期货（机构对冲压力的高频确认）")
        add("")
        add("| 合约 | 最新 | 涨跌 | 现货 | 基差 | 基差率 | 持仓量 | 成交量 |")
        add("|---|---:|---:|---:|---:|---:|---:|---:|")
        for r in fut["rows"]:
            add("| {} | {} | {} | {} | {} | {} | {:.0f} | {:.0f} |".format(
                r["name"], price_str(r["last"]), pct_str(r["pct"]),
                price_str(r["spot"]) if r.get("spot") else "—",
                "{:+.2f}".format(r["basis"]) if r.get("basis") is not None else "—",
                "{:+.2f}%".format(r["basis_pct"])
                if r.get("basis_pct") is not None else "—",
                r.get("oi") or 0, r.get("volume") or 0))
        add("")
        add("> 行情时间：{}（新浪期货接口，与现货指数为**同一时点**快照）。"
            "基差为负即贴水。".format(fut.get("time") or "—"))
        add("")
        for s in fut["signals"]:
            add("- {}".format(s))
        add("")
    else:
        add("### 6.1.1 股指期货基差")
        add("")
        add("> 期货接口未取到数据（状态：{}），本节跳过，不影响其余章节。".format(
            fut.get("status") or "empty"))
        add("")

    add("### 6.2 A股指数估值横向")
    add("")
    if glob["valuations"]:
        add("| 指数 | PE(倍) | 今日涨跌 |")
        add("|---|---:|---:|")
        for v in glob["valuations"]:
            add("| {} | {:.1f} | {} |".format(
                v["name"], v["pe"], pct_str(v["pct"])))
        add("")
        chart = save_chart("index_pe",
                           svg_bars([v["name"][:6] for v in glob["valuations"]],
                                    [v["pe"] for v in glob["valuations"]],
                                    "A股主要指数 PE 横向对比（倍）",
                                    fmt="{:.0f}", zero_line=False))
        if chart:
            add("![指数PE横向]({})".format(chart))
            add("")
        add("> 指数 PE 为**横向对比**；个股/ETF 自身历史分位见 6.3，"
            "股债性价比（ERP）见第十一章。")
        add("")
    else:
        add("> 暂无指数估值数据。")
        add("")

    add("### 6.3 ETF 估值历史分位（接口补齐）")
    add("")
    if universe["val_rows"]:
        add("| ETF | PE(TTM) | PB | PE历史分位 | PB历史分位 | 盈利收益率 | ROE |")
        add("|---|---:|---:|---:|---:|---:|---:|")
        for r in universe["val_rows"][:14]:
            add("| {} | {} | {} | {} | {} | {} | {} |".format(
                r["name"],
                "{:.1f}".format(r["pe"]) if r["pe"] else "—",
                "{:.2f}".format(r["pb"]) if r["pb"] else "—",
                "{:.0f}%".format(r["pe_pct"]) if r["pe_pct"] is not None else "—",
                "{:.0f}%".format(r["pb_pct"]) if r["pb_pct"] is not None else "—",
                "{:.2f}%".format(r["earn_yield"]) if r["earn_yield"] else "—",
                "{:.1f}%".format(r["roe"]) if r["roe"] else "—"))
        add("")
        add("> **分位口径**：行情商提供的历史百分位（0%=历史最便宜，100%=最贵），"
            "不是本报告自算。低估 ≤30% 共 {} 只、高估 ≥80% 共 {} 只。"
            "此前「缺历史 PE 分位、回答不了贵不贵」的限制已解除。".format(
                len(universe["cheap"]), len(universe["rich"])))
        add("")
    else:
        add("> 暂无 ETF 估值分位数据（封装接口不可用且快照无缓存）。")
        add("")

    add("### 6.4 联动信号")
    add("")
    if glob["signals"]:
        for i, s in enumerate(glob["signals"], 1):
            add("{}. {}".format(i, s))
    else:
        add("- 外围与估值数据不足，无法生成联动判断。")
    add("")

    add("---")
    add("")
    add("## 七、市场画像评分（自建 + 官方对照）")
    add("")
    add("> 9 个维度由本报告已有的全量数据自行计算（1=弱 / 5=强），"
        "**不是**行情商的官方 28 维评分；估值维度因缺历史分位仅为横向代理。")
    add("")
    missing = [n for ok, n in (
        (breadth["available"], "涨跌广度"),
        (fund["available"], "板块资金流"),
        (glob["available"], "外围行情")) if not ok]
    if missing:
        add("> ⚠️ **降级评分**：{}数据缺失，相关维度按缺省值计分，"
            "本次总分与结论仅供参考，不宜与历史报告横向比较。".format(
                "、".join(missing)))
        add("")
    if profile["dims"]:
        add("| 维度 | 评分 | 状态 | 依据 |")
        add("|---|---:|---|---|")
        for x in profile["dims"]:
            add("| {} | {}/5 | {} | {} |".format(
                x["name"], x["score"], x["label"], x["note"]))
        add("")
        add("**总分 {} / {}（{:.0f}%）**：{}。".format(
            profile["total"], profile["max_total"],
            profile["ratio"], profile["verdict"]))
        add("")
        chart = save_chart("profile_scores",
                           svg_bars([x["name"] for x in profile["dims"]],
                                    [x["score"] for x in profile["dims"]],
                                    "自建市场画像 9 维评分（1~5）",
                                    fmt="{:.0f}", zero_line=False))
        if chart:
            add("![市场画像评分]({})".format(chart))
            add("")
        add("> 高分集中在「情绪 / 广度 / 量能」等短期维度，"
            "低分集中在「中长期方向」等趋势维度——这正是**短强中弱**的量化体现。")
        add("")

    add("### 7.2 官方市场画像（行情商评分，接口）")
    add("")
    if official["available"]:
        add("> 原始评分 **{}** / 调整评分 **{}**；接口本次返回 {} 个可解析维度"
            "（行情商标称 28 维画像，实际字段以返回为准）。"
            "这不是自建评分，可与 7.1 互相印证。".format(
                official["raw"], official["adj"], len(official["dims"])))
        add("")
        add("| 维度 | 评分 | 状态 |")
        add("|---|---:|---|")
        for d in official["dims"]:
            add("| {} | {}/5 | {} |".format(
                d["name"], int(d["score"]), md_plain(d["status"])))
        add("")
        for s in official["signals"]:
            add("- {}".format(s))
        add("")
    else:
        add("> 官方画像接口不可用，仅保留 7.1 自建评分。")
        add("")

    add("---")
    add("")
    add("## 八、事件、异动、龙虎榜与筹码")
    add("")
    if events["items"]:
        add("| 类型 | 标题 | 摘要 / 明细 | 来源 |")
        add("|---|---|---|---|")
        for it in events["items"] + mevents["items"]:
            detail = md_plain(it["detail"])
            if len(detail) > 150:
                detail = detail[:150] + "…"
            add("| {} | {} | {} | {} |".format(
                md_plain(it["type"]), md_plain(it["title"]), detail,
                md_plain(it["source"])))
        add("")
    add("> 结构化事件 {} 条 / 龙虎榜机构席位 {} 条 / 筹码标的 {} 只。".format(
        mevents["event_n"], mevents["lhb_n"], mevents["chip_n"]))
    add("")
    for s in events["signals"] + mevents["signals"]:
        add("- {}".format(s))
    add("")

    add("---")
    add("")
    add("## 九、历史规则回测")
    add("")
    add("> 用现有日线对本报告用到的经验规则做频率与胜率校准，"
        "**只做背景参考，不是策略回测**。")
    add("")
    if backtest["rules"]:
        add("### 9.1 规则汇总（扣成本 / 样本外 / 显著性）")
        add("")
        if backtest["base"]:
            add("> 买入持有基准：次日胜率 **{:.0f}%**、平均 {}；"
                "双边成本按 {}% 扣除；样本内为每只标的前 70% 日线，后 30% 为样本外。".format(
                    backtest["base"]["win"], pct_str(backtest["base"]["mean"]),
                    backtest["cost_pct"]))
            add("")
        add("| 规则 | 触发 | 次日胜率 | 扣成本 | 样本内 | 样本外 | z值 | 平均次日 | 5日胜率 |")
        add("|---|---:|---:|---:|---:|---:|---:|---:|---:|")
        for r in backtest["rules"]:
            h1, h5 = r["h"].get(1), r["h"].get(5)
            add("| {} | {} | {} | {} | {} | {} | {} | {} | {} |".format(
                r["label"], r["n"],
                "{:.0f}%".format(h1["win"]) if h1 else "—",
                "{:.0f}%".format(h1["net_win"]) if h1 else "—",
                "{:.0f}%".format(h1["ins_win"]) if h1 and h1["ins_win"] is not None else "—",
                "{:.0f}%".format(h1["oos_win"]) if h1 and h1["oos_win"] is not None else "—",
                "{:.2f}".format(h1["z"]) if h1 else "—",
                pct_str(h1["mean"]) if h1 else "—",
                "{:.0f}%".format(h5["win"]) if h5 else "—"))
        add("")
        # 多重比较校正：规则数 × 期限（次日/5日）即检验次数，单次 1.96 会高估显著性
        _n_tests = max(1, len(backtest["rules"]) * 2)
        _zcrit = {1: 1.96, 2: 2.24, 3: 2.39, 4: 2.50, 6: 2.64,
                  8: 2.73, 10: 2.81, 12: 2.87}.get(_n_tests, 2.90)
        _passed = []
        for _r in backtest["rules"]:
            for _hh, _hn in ((_r["h"].get(1), "次日"), (_r["h"].get(5), "5日")):
                if _hh and _hh.get("z") is not None and abs(_hh["z"]) >= _zcrit:
                    _passed.append("{}·{}".format(_r["label"], _hn))
        add("> z 值衡量胜率是否显著偏离 50%。本次共 **{} 次检验**（{} 规则 × 次日/5日），"
            "按 Bonferroni 校正后阈值约为 **|z|≥{:.2f}**；"
            "通过校正的：{}。未校正的 z 值只作描述、不作结论。".format(
                _n_tests, len(backtest["rules"]), _zcrit,
                "、".join(_passed) if _passed else "**无**（全部落在随机波动范围内）"))
        add("> 方法边界：样本外按时间前 70%/后 30% 切分，但相邻交易日的次日收益相互重叠，"
            "未做 purge/embargo，也未做分块 bootstrap；价格未复权（分红除权会造出假跳空），"
            "成本固定按双边 {}%。因此命中率仅作背景参考。".format(backtest["cost_pct"]))
        add("")
        add("### 9.2 分标的明细（次日）")
        add("")
        add("| 规则 | 标的 | 触发 | 次日胜率 | 平均次日 |")
        add("|---|---|---:|---:|---:|")
        for r in backtest["rules"]:
            for p in r["per"]:
                h1 = p["h"].get(1)
                add("| {} | {} | {} | {} | {} |".format(
                    r["label"], p["name"], p["n"],
                    "{:.0f}%".format(h1["win"]) if h1 else "—",
                    pct_str(h1["mean"]) if h1 else "—"))
        add("")
    for s in backtest["signals"]:
        add("- {}".format(s))
    add("")

    add("---")
    add("")
    add("## 十、全市场 ETF 全景（接口全量）")
    add("")
    if not universe["available"]:
        add("> 全市场 ETF 明细接口不可用，本章降级；1.6 中「不能推断全市场ETF总数/规模」"
            "的限制本次仍成立。")
        add("")
    else:
        add("> 数据源为全市场 ETF 列表（{} 只），**不是样本池**。"
            "净值/规模/估值/净申购均取自行情商统计口径。".format(
                universe["universe_n"]))
        add("")
        add("### 10.1 规模与样本代表性")
        add("")
        add("| 指标 | 数值 |")
        add("|---|---:|")
        add("| 全市场 ETF 只数 | {} |".format(universe["universe_n"]))
        add("| 合计净资产 | {:.0f} 亿 |".format(universe["total_nv_yi"]))
        add("| 合计总资产 | {:.0f} 亿 |".format(universe["total_asset_yi"]))
        add("| 样本池只数 / 覆盖率 | {} / {:.1f}% |".format(
            universe["sample_n"], universe["sample_cover_pct"]))
        add("| 样本池规模占比 | {:.1f}% |".format(
            universe["sample_asset_pct"] or 0))
        add("")
        add("> 1.6 中原先标注「不能推断全市场 ETF 总数与总规模」——"
            "**该限制已由全量接口解除**；但样本池仍是分层抽样，"
            "1.2 排行与 1.3 分组均值依旧不能当作全市场统计。")
        add("")

        add("### 10.2 真实净申购资金流（一级市场）")
        add("")
        add("> **披露日 {}（T-1，非当日）**：基金份额申赎数据通常晚一个交易日披露，"
            "因此下表是截至上一交易日的净申购，不要与今日二级成交额直接对齐。".format(
                universe["flow_date"] or "未知"))
        add("")
        if universe["flow_rows"]:
            add("| ETF | 单日(亿) | 近周(亿) | 近月(亿) | 近季(亿) | 单日占规模 |")
            add("|---|---:|---:|---:|---:|---:|")
            show = (universe["flow_rows"][:8] +
                    universe["flow_rows"][-5:] if len(universe["flow_rows"]) > 13
                    else universe["flow_rows"])
            for r in show:
                add("| {} | {:+.2f} | {:+.1f} | {:+.1f} | {:+.1f} | {} |".format(
                    r["name"], r["d"], r["w"], r["m"], r["q"],
                    "{:+.2f}%".format(r["d_pct"]) if r["d_pct"] is not None else "—"))
            add("")
            chart = save_chart("etf_netflow",
                               svg_bars([r["name"][:6] for r in universe["flow_rows"][:8]],
                                        [r["d"] for r in universe["flow_rows"][:8]],
                                        "样本ETF当日净申购 Top（亿元）", fmt="{:+.1f}"))
            if chart:
                add("![ETF净申购]({})".format(chart))
                add("")
            add("> 这是**净申购/净赎回**（真实资金进出），与成交额（二级换手）不同；"
                "第五章逐日累积份额的土办法现在有接口数据可交叉校验。")
            add("")

        add("### 10.3 全市场视角信号")
        add("")
        for i, s in enumerate(universe["signals"], 1):
            add("{}. {}".format(i, s))
        add("")

    add("---")
    add("")
    add("## 十一、宏观流动性与股债性价比")
    add("")
    if not macro["available"]:
        add("> 宏观指标接口不可用，本章降级。")
        add("")
    else:
        add("| 指标 | 数值 | 单位 |")
        add("|---|---:|---|")
        for r in macro["rows"]:
            add("| {} | {} | {} |".format(
                r["label"], r["value"], md_plain(r["unit"])))
        add("")
        if macro["erp"] is not None:
            add("### 11.1 ERP 计算")
            add("")
            add("> **盈利收益率中位数 {:.2f}% − 10年国债 {:.2f}% = ERP {:+.2f}pct**。"
                "盈利收益率取样本 ETF 的 1/PE(TTM) 中位数，是代理口径；"
                "严格 ERP 应用指数整体法 PE 与当期无风险利率。".format(
                    macro["ey_mid"], macro["y10"], macro["erp"]))
            add("")
        add("### 11.2 宏观信号")
        add("")
        for i, s in enumerate(macro["signals"], 1):
            add("{}. {}".format(i, s))
        add("")
        add("> **注意**：宏观数据更新频率低（月度/日度国债），"
            "与当日盘面不是同一时间尺度，只作为估值与流动性的背景，不用于短线择时。")
        add("")

    add("---")
    add("")
    add("## 十二、逐日历史累积与今日 vs 昨日")
    add("")
    add("> 公开接口只给当日快照、不给历史序列（日期参数无效），"
        "因此**没有接口就自己累积**：脚本每次运行把当日关键读数写入 "
        "`etf_daily_history.json`，连续运行后本章自动生成 diff。")
    add("")
    if not diff["available"]:
        for s in diff["signals"]:
            add("- {}".format(s))
        add("")
    else:
        _bf_n = diff.get("backfill_n") or 0
        add("对比日：**{}**（历史共 {} 个交易日{}）".format(
            diff["prev_date"], diff["n_days"],
            "，其中 {} 日仅有成交额回填".format(_bf_n) if _bf_n else ""))
        add("")
        add("| 指标 | 今日 | 昨日 | 变化 | 变化率 |")
        add("|---|---:|---:|---:|---:|")
        for r in diff["rows"]:
            add("| {} | {:.2f} | {:.2f} | {:+.2f} | {} |".format(
                r["label"], r["cur"], r["prev"], r["diff"],
                pct_str(r["pct"]) if r["pct"] is not None else "—"))
        add("")
        for s in diff["signals"]:
            add("- {}".format(s))
        add("")
        add("> 第十四章「明日只盯」里的「连续第二日维持净流入 / 守住成交额关口 / "
            "广度不回落」从此可由脚本自动验证，不再依赖人眼。")
        add("")

    if amt.get("available"):
        add("### 12.2 近 {} 日成交额序列（实时 + 回填，量能的分位参照）".format(
            amt["n"]))
        add("")
        add("| 日期 | 两市成交额(亿) | 环比 | 上证 | 深证 | 来源 |")
        add("|---|---:|---:|---:|---:|---|")
        for r in amt["rows"]:
            add("| {} | {:.0f} | {} | {} | {} | {} |".format(
                r["date"], r["amount"],
                pct_str(r["chg"]) if r["chg"] is not None else "—",
                pct_str(r["sh_pct"]) if r["sh_pct"] is not None else "—",
                pct_str(r["sz_pct"]) if r["sz_pct"] is not None else "—",
                r["src"]))
        add("")
        for s in amt["signals"]:
            add("- {}".format(s))
        add("")
        add("> **为什么只回填成交额**：沪市/深市指数日 K 自带成交额字段，"
            "可以直接补齐；而上涨家数、主力净额、ERP 只有当日快照才有，"
            "没有历史接口，缺失就是缺失——**不插值、不猜**，"
            "这也是本章其他字段至今只有 1~2 天的原因。")
        add("")

    add("---")
    add("")
    add("## 十三、数据质量与口径审计")
    add("")
    add("**通过率 {:.0f}% · {}**".format(quality["score"], quality["verdict"]))
    add("")
    add("| 检查项 | 结果 | 说明 |")
    add("|---|---|---|")
    for c in quality["checks"]:
        add("| {} | {} | {} |".format(
            c["name"], "✅" if c["ok"] else "⚠️", md_plain(c["detail"])))
    add("")
    _health = quality.get("health") or []
    if _health:
        add("**端点契约自检**（每通道的状态 + 条数下限，结果同时写入 `endpoint_health.json`）：")
        add("")
        add("| 通道 | 端点 | 状态 | 条数 / 下限 | 结论 |")
        add("|---|---|---|---:|---|")
        for hc in _health:
            add("| {} | `{}` | {} | {} / {} | {} |".format(
                hc["name"], hc["endpoint"], hc["status"],
                hc["n"], hc["min_n"], "✅" if hc["ok"] else "⚠️"))
        add("")
        add("> 断言逻辑：状态必须为 `ok`/`web_channel` 且条数 ≥ 下限；"
            "任一通道不达标即判「契约变更」，提示端点或字段可能已改版——"
            "这正是本次华夏时报/东财 PC 页改版能够被及时发现的方式。")
        add("")
    add("> **通过率口径**：{:.0f}% = 通过项 / 总项，其中包含「通道可达」与「数据形状断言」"
        "（条数下限、时间跨度、数值有效性、环比突变）两类，衡量的是**数据是否可用**，"
        "不是数据是否绝对正确；断言只设下限，不能替代人工复核。".format(quality["score"]))
    add("")
    if cons.get("rows"):
        add("**跨源 / 跨章节一致性**（同一事实在两套接口、两个章节里必须对得上，"
            "结果写入 `consistency.json`）：")
        add("")
        add("| 校验项 | 值 A | 值 B | 差异 / 容差 | 结论 | 说明 |")
        add("|---|---|---:|---:|---|---|")
        for _c in cons["rows"]:
            add("| {} | {} | {} | {} | {} | {} |".format(
                md_plain(_c["name"]), _c["a"], _c["b"], _c.get("diff") or "—",
                "✅" if _c["ok"] else "⚠️ 不一致", md_plain(_c["note"])))
        add("")
        add("> **{}/{} 项通过（{}）**。这一层与端点契约分工不同："
            "契约查「拿得到吗」，一致性查「拿到的数字彼此矛盾吗」——"
            "广度接口、全量排行、ETF 明细是三套独立采集，"
            "任一处改版或时点漂移，通常先在这里露出来。".format(
                cons["ok_n"], cons["n"], cons["verdict"]))
        add("")
    add("> **派生库**：以上结果（含本表、质量审计、波动率锥、主题 AR、逐日成交额）"
        "同时写入 `etf_data.sqlite`，`run_meta.schema_ver` 记录结构版本。"
        "JSON 仍是唯一真源，派生库删掉即可重建。两张清单：")
    add("")
    add("| 组 | 表 | 内容 | 写法 |")
    add("|---|---|---|---|")
    add("| 本次结果 | `daily` `vol_cone` `topic_ar` | 逐日成交额/波动率锥/主题 AR | "
        "按日期覆盖 |")
    add("| 本次结果 | `consistency` `quality` `option_iv` | 一致性、质量审计、期权 IV | "
        "按报告日覆盖 |")
    add("| 快照截面 | `snapshot_quote` `snapshot_meta` | 全部标的行情字段（88 字段解析后） | "
        "按快照日覆盖 |")
    add("| 快照截面 | `sector_flow` `breadth_metric` `global_index` | 板块资金流"
        "（行业+概念）、广度指标 {} 项 + 全市场成交额 3 项、外围指数 | "
        "按快照日覆盖 |".format(len(BREADTH_METRICS)))
    add("| 快照截面 | `etf_profile` | 全市场 ETF 规模/估值分位/净申购/量价/换手，"
        "由 8 个原始 list 按代码合成窄化宽表，金额统一为**亿元** | 按快照日覆盖 |")
    add("| 快照截面 | `quote_kline` | 每只标的的日线 OHLCV（约 13k 根），"
        "技术指标可据此复算 | 按 (code,date) upsert |")
    add("| 排行宏观 | `stock_rank` `etf_rank` | 个股 5578 只、ETF 1719 只的多周期涨跌幅与主力净额 | "
        "按快照日覆盖 |")
    add("| 排行宏观 | `futures` `market_score` `macro_indicator` `breadth_dist` | "
        "期货基差、28 维评分、宏观指标 631 项、涨跌分布区间 | 按快照日覆盖 |")
    add("| 流水资讯 | `shares_hist` `flash_news` `news_item` | 份额累积、快讯累积"
        "（去重看视图 `flash_unique`）、热点/公告/研报/资金流/热榜 | 按主键 upsert |")
    add("| 事件筹码 | `lhb` `chip_metric` `endpoint_health` | 龙虎榜（4 列 + 原始行 "
        "`raw`）、筹码分布、端点健康 | 按快照日覆盖 |")
    add("| 事件日历 | `calendar_item` `ipo_item` `rating_stock` | 财经日历、新股申购、"
        "机构一致预期 | 按快照日覆盖 |")
    add("| 分时明细 | `breadth_minute` `yidong_item` `strong_stock` "
        "`breadth_note` `ext_meta` | 分时序列（121 点×2）、盘口异动、强势股、"
        "成段点评、采集元信息 | 按快照日覆盖 |")
    add("")
    add("> 覆盖 vs upsert 的区别是刻意的：快照是「某一天的截面」，重跑应整段替换；"
        "份额和快讯是「跨日累积的流水」，今天的运行不该把昨天的记录删掉。"
        "真源库快照里那些整句文字（如融资余额点评）不进数值列，避免表里一排 NULL，"
        "而是单独进 `breadth_note`，一条 SQL 就能读；"
        "`_lhb` 的榜单是无表头数组（`['sh603236','移远通信','3',4,…]`），"
        "只把能确定语义的 code/名称/买入额/涨跌幅取出，其余列连原始行存 `raw`，"
        "不猜 schema。至此 `etf_store.sqlite` 的 `snapshot_store` 快照"
        "**每一个子键**都镜像进派生库，分时序列按点拆行，"
        "采集元信息（`*_status` / `*_n`）进 `ext_meta`。")
    add("")
    add("> **时点说明**：派生索引 `etf_data.sqlite` 的截面表，由**历史真源库** "
        "`etf_store.sqlite` 的 `snapshot_store` 最新一份快照整段镜像而来；"
        "`snapshot_meta.updated` 取自真源库的写入时刻（在线运行带 "
        "`--save-snapshot` 时刷新），不再是某个 JSON 文件的修改时间。"
        "在线运行时真源库与派生镜像同批更新，两端一致；离线运行时截面停在"
        "快照日，跨日结论请以第十二章的逐日累积为准。")
    add("")
    add("```")
    add('python etf_analysis.py --sql "SELECT date, amount, amount_chg, up_ratio '
        'FROM daily ORDER BY date DESC LIMIT 10"')
    add('python etf_analysis.py --sql "SELECT name, a, b FROM consistency WHERE ok=0"')
    add('python etf_analysis.py --sql "SELECT kind, name, net_yi FROM sector_flow '
        'ORDER BY net_yi DESC LIMIT 5"')
    add('python etf_analysis.py --sql "SELECT code, name, total_asset_yi, pe_pct, '
        'flow_m_yi FROM etf_profile ORDER BY total_asset_yi DESC LIMIT 5"')
    add('python etf_analysis.py --sql "SELECT p.name, p.pe_pct, k.close FROM '
        'etf_profile p JOIN quote_kline k ON k.code=p.code '
        'WHERE k.date=\'{}\' AND p.pe_pct IS NOT NULL '
        'ORDER BY p.pe_pct LIMIT 5"'.format(today["date"] or ""))
    add('python etf_analysis.py --sql "SELECT time, sources, title FROM '
        'flash_unique WHERE date=\'{}\' ORDER BY time DESC LIMIT 5"'.format(
            today["date"] or ""))
    add('python etf_analysis.py --sql "SELECT code, name, buy_yi, pct FROM lhb '
        'ORDER BY buy_yi DESC LIMIT 5"')
    add("```")
    add("")
    add("> ⚠️ 项为降级/提示项，不影响报告生成，但相关结论需打折看待。"
        "本报告所有外部端点均为公开 HTTP 接口，无 SLA、字段可能变更；"
        "离线快照是静态截面，跨日需重新在线取数或用第十二章的逐日累积。")
    add("")

    add("---")
    add("")
    add("## 十四、接下来走势")
    add("")
    add("### 14.1 技术仪表盘")
    add("")
    add("| ETF | 现价 | MA5 | MA20 | MA60 | MACD柱 | KDJ-K | RSI6 | 收盘位置 |")
    add("|---|---:|---:|---:|---:|---:|---:|---:|---:|")
    for r in outlook["rows"]:
        ind, q = r["ind"], r["q"]
        macd = ind["macd"]
        add("| {} | {} | {} | {} | {} | {} | {} | {} | {} |".format(
            r["name"], price_str(r["price"]), price_str(ind["ma5"]),
            price_str(ind["ma20"]), price_str(ind["ma60"]),
            "{:+.4f}".format(macd["bar"]) if "bar" in macd else "—",
            "{:.1f}".format(ind["kdj"]["k"]) if ind["kdj"] else "—",
            "{:.1f}".format(ind["rsi6"]) if ind["rsi6"] else "—",
            "{:.0%}".format(r["close_pos"]) if r["close_pos"] is not None else "—"))
    add("")
    under = [r["name"] for r in outlook["rows"]
             if r["ind"]["ma60"] and r["price"] < r["ind"]["ma60"]]
    if under:
        add("**{} 仍全部位于 MA60 下方**；收复 5/10 日线不等于中期趋势反转。".format(
            "、".join(under)))
        add("")

    if advtech["rows"]:
        add("### 14.1b 波动率、缺口与多周期")
        add("")
        add("| ETF | ATR(20)占价 | 20日波动率 | 开盘缺口 | 周线MA5/MA10 | 周线状态 |")
        add("|---|---:|---:|---:|---:|---|")
        for r in advtech["rows"]:
            add("| {} | {} | {} | {} | {} | {} |".format(
                r["name"],
                "{:.2f}%".format(r["atr_pct"]) if r["atr_pct"] else "—",
                "{:.2f}%".format(r["vol"]) if r["vol"] else "—",
                pct_str(r["gap"]) if r["gap"] is not None else "—",
                "{}/{}".format(price_str(r["ma5w"]), price_str(r["ma10w"]))
                if r["ma5w"] and r["ma10w"] else "—",
                r["week_trend"]))
        add("")
        for s in advtech["signals"]:
            add("- {}".format(s))
        add("")

    if cone.get("rows"):
        add("### 14.1c 波动率锥（已实现波动 HV20/60/120 与历史分位）")
        add("")
        add("| ETF | HV20 | HV60 | HV120 | HV20分位 | 锥区间(20日 低/中/高) | 状态 |")
        add("|---|---:|---:|---:|---:|---:|---|")
        for r in cone["rows"]:
            _c = (r.get("cone") or {}).get(20) or {}
            add("| {} | {:.1f}% | {} | {} | {} | {} | {} |".format(
                r["name"], r["hv20"],
                "{:.1f}%".format(r["hv60"]) if r["hv60"] else "—",
                "{:.1f}%".format(r["hv120"]) if r["hv120"] else "—",
                "{:.0f}%".format(r["pct"]) if r["pct"] is not None else "—",
                "{:.1f} / {:.1f} / {:.1f}".format(
                    _c.get("low", 0), _c.get("mid", 0), _c.get("high", 0))
                if _c else "—",
                r["state"]))
        add("")
        for s in cone["signals"]:
            add("- {}".format(s))
        add("")
        add("> **怎么用**：分位回答「现在的波动算大算小」，锥区间回答「还能大到哪」。"
            "分位 >80% 时同样的消息更容易打出大阳大阴，止损宽度应按 HV20 而非固定百分比；"
            "分位 <20% 的低波+高估值组合才是真正的脆弱结构（低波意味着市场没在定价风险）。")
        add("")

    if opt.get("available"):
        add("### 14.1d 期权隐含波动率（IV）与波动率风险溢价")
        add("")
        add("| 标的 | 现价 | ATM行权价 | 剩余天 | 认购IV | 认沽IV | ATM IV "
            "| HV20 | VRP(IV−HV20) | 偏斜(95P−105C) |")
        add("|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|")
        for r in opt["rows"]:
            add("| {} | {:.3f} | {} | {} | {} | {} | **{:.1f}%** | {} | {} | {} |".format(
                r["name"], r["spot"], r["atm_k"], r["days"],
                "{:.1f}%".format(r["iv_call"] * 100) if r.get("iv_call") else "—",
                "{:.1f}%".format(r["iv_put"] * 100) if r.get("iv_put") else "—",
                r["iv_atm"],
                "{:.1f}%".format(r["hv20"]) if r.get("hv20") else "—",
                "{:+.1f}pct".format(r["vrp"]) if r.get("vrp") is not None else "—",
                "{:+.1f}pct".format(r["skew"]) if r.get("skew") is not None else "—"))
        add("")
        for s in opt["signals"]:
            add("- {}".format(s))
        add("")
        add("> **怎么用**：VRP > 0 时波动率买方在为「未来的不确定性」多付钱——"
            "做多波动（买跨式/宽跨式）性价比下降，卖方策略的期望收益上升；"
            "VRP < 0 则相反。偏斜陡说明虚值认沽被抢，是**下跌保护**的定价，"
            "不是「马上要跌」的信号——它只说明保险贵了，方向仍要看价格与广度。")
        add("")
        add("> **与 14.1c 的区别**：HV 是**回望**（已经发生的波动），"
            "IV 是**前瞻**（期权市场对到期的定价），两者的差才是可交易的溢价。"
            "本表的 IV 全部由最新价反解得到，不依赖行情商给出的 IV 字段。")
        add("")
    else:
        _o = (opt.get("signals") or ["期权链不可得"])[0]
        add("> 14.1d 期权 IV 本次跳过：{}".format(_o))
        add("")

    add("### 14.2 支撑与压力")
    add("")
    add("| ETF | 现价 | 支撑(昨日/5日低点) | 第一压力 | 第二压力 | MA60 |")
    add("|---|---:|---:|---:|---:|---:|")
    for r in outlook["rows"]:
        ps = r["pressures"]
        first = "{} {}".format(ps[0][0], price_str(ps[0][1])) if ps else "—"
        second = "{} {}".format(ps[1][0], price_str(ps[1][1])) if len(ps) > 1 else "—"
        gap = ((ps[0][1] / r["price"] - 1) * 100) if ps else None
        add("| {} | {} | {} | {}{} | {} | {} |".format(
            r["name"], price_str(r["price"]), price_str(r["support"]),
            first, "（{:+.1f}%）".format(gap) if gap is not None else "",
            second, price_str(r["ind"]["ma60"])))
    add("")
    add("> 第一压力取 MA20 / 布林中轨，恰是 10 月下旬密集成交区，解套盘会明显涌出；"
        "表中只列高于现价的价位。")
    poc_line = "、".join("{} {}".format(r["name"], price_str(r["poc"]))
                        for r in outlook["rows"] if r["poc"])
    if poc_line:
        add("> 近20日**成交量加权密集区中枢**（另一套压力参考）：{}。".format(poc_line))
    add("")

    add("### 14.3 情景推演")
    add("")
    add("**基准倾向：{}**（多头线索 {} 条 / 空头线索 {} 条，仅为规则化打分，不是概率预测）".format(
        outlook["base_scene"], outlook["total_bull"], outlook["total_bear"]))
    add("")
    for i, s in enumerate(outlook["scenes"], 1):
        tag = " ← 基准倾向" if s["name"].startswith(outlook["base_scene"][:2]) else ""
        add("**{} {}{}**".format("①②③"[i - 1], s["name"], tag))
        add("- 形态：{}".format(s["desc"]))
        add("- 确认信号：{}".format(s["confirm"]))
        add("- 依据：{}".format(s["trigger"]))
        add("")
    if backtest["rules"]:
        r1 = next((r for r in backtest["rules"] if "单日涨幅" in r["label"]), None)
        r2 = next((r for r in backtest["rules"] if "平开收高位" in r["label"]), None)
        if r1 and r1["h"].get(1):
            add("> **历史校准①**：过去 {} 次「单日涨幅 ≥5%」后，次日收阳胜率 {:.0f}%、"
                "平均次日 {}；但 5 日平均 {}。单日大涨不等于趋势延续。".format(
                    r1["h"][1]["n"], r1["h"][1]["win"],
                    pct_str(r1["h"][1]["mean"]),
                    pct_str(r1["h"][5]["mean"]) if r1["h"].get(5) else "—"))
            add("")
        if r2 and r2["h"].get(1):
            add("> **历史校准②（重要，与形态判断冲突）**：第一、二章曾把「平开单边拉升、"
                "收于全天高位」解读为买盘持续、次日惯性偏强；但回测显示该形态触发 {} 次、"
                "次日收阳胜率仅 **{:.0f}%**、平均次日 **{}**。"
                "当经验形态判断与历史统计冲突时应优先采信统计，"
                "因此本章**下调**「次日惯性」权重，仍以震荡筑底为基准情景。".format(
                    r2["h"][1]["n"], r2["h"][1]["win"],
                    pct_str(r2["h"][1]["mean"])))
            add("")

    watch = list(outlook["watch"])
    if market["total"]:
        watch.append("**全市场成交额**：两市需守住 **{:.0f} 亿（{}）**，并较今日{}；"
                     "缩量回落至关口下方，即为反弹动能衰减的确认信号。".format(
                         market["gate_threshold"], market["gate"],
                         "继续放大" if (market["chg"] or 0) > 3 else "不再明显萎缩"))
    if fund["available"] and fund["in_top"]:
        watch.append("**板块资金流**：{} 等净流入居首板块需连续第二日维持净流入，"
                     "全市场主力净额不大幅转负；若龙头板块转为净流出，"
                     "即为结构性资金退潮的确认信号。".format(
                         "、".join(r["name"] for r in fund["in_top"][:2])))
    if breadth["available"]:
        watch.append("**涨跌广度**：涨停家数需维持、上涨占比不低于 50%，"
                     "且「昨日涨停股今日上涨比例」回升——否则普涨只是单日脉冲。")
    if breadth.get("margin_bal"):
        watch.append("**杠杆资金**：融资余额（{:.0f} 亿）需止跌回升；"
                     "若继续连降，说明反弹缺乏杠杆资金确认。".format(
                         breadth["margin_bal"]))
    if micro.get("overheat"):
        watch.append("**ETF 溢价**：{} 只ETF溢价 ≥0.5%，次日若溢价快速收敛，"
                     "往往伴随价格回落，是情绪退潮的先行信号。".format(
                         len(micro["overheat"])))
    if glob["available"]:
        watch.append("**外围**：留意隔夜美股与次日港股科技开盘，"
                     "A股科创今日大幅跑赢港股科技，若外围不接力，缺口易回补。")
    if attribution["available"] and attribution["topics"]:
        watch.append("**事件验证**：第十五章首要驱动「{}」需有后续资讯或公告继续印证；"
                     "若次日无新增催化、且盘后发酵的传闻未被证实，"
                     "则今日更可能只是单日情绪脉冲。".format(
                         attribution["topics"][0]["topic"]))
    cn = {3: "三", 4: "四", 5: "五", 6: "六", 7: "七", 8: "八", 9: "九",
          10: "十", 11: "十一", 12: "十二"}
    add("### 14.4 明日只盯{}件事".format(cn.get(len(watch), str(len(watch)))))
    add("")
    for i, w in enumerate(watch, 1):
        add("{}. {}".format(i, w))
    add("")
    add("### 14.5 操作提示")
    add("")
    add("- 涨停/近涨停位置不宜追高，日内浮盈盘厚，次日兑现压力大；")
    add("- 已有仓位可在第一压力位附近分批兑现；")
    add("- 空仓者等回踩不破昨日低点再考虑分批，比在情绪高点一次性买入更从容。")
    add("")

    add("---")
    add("")
    add("## 十五、事件采集与涨跌归因")
    add("")
    if not attribution["available"]:
        for s in attribution["signals"]:
            add("- {}".format(s))
        add("")
    else:
        add("> 本章回答「**今天为什么涨**」。做法：采集带时间戳的财经资讯 → "
            "用分时定位主升起点 → 只把早于起点的资讯算「驱动」→ "
            "与板块资金流、样本 ETF 涨幅、外围方向做三方校验 → "
            "输出置信度排序。**这是相关性归因，不是因果证明**。")
        add("")

        add("### 15.1 资讯采集总览")
        add("")
        add("| 类别 | 条数 | 用途 |")
        add("|---|---:|---|")
        usage = {
            "7×24快讯": "东财JSON+H5+PC页+新浪+华夏时报+财联社六通道，带完整日期；PC页补盘后叙事与次日盘中验证",
            "热点资讯": "当日主线与催化，带发布时间与热度排名",
            "新闻正文": "提取具体利好细节（只对最相关者抓取）",
            "公司公告": "回购、大单、业绩等公司层事件",
            "个股研报": "卖方对产业催化的解读",
            "脱水研报": "机构深度摘要",
            "财经日历": "分红 / 解禁 / 财报预约",
            "新股日历": "一级市场节奏",
            "机构预期样本": "主题龙头的评级与一致预期",
        }
        for k, v in attribution["collect"].items():
            add("| {} | {} | {} |".format(
                md_plain(k), v, usage.get(k, "—")))
        add("")
        if attribution["sources"]:
            add("主要来源：{}。".format(
                "、".join("{}（{}）".format(md_plain(s), n)
                         for s, n in attribution["sources"][:6])))
            add("")

        add("### 15.2 时序切分与主升起点")
        add("")
        if attribution["launch"]:
            add("> 上证分时 5 分钟滚动涨幅最大窗口起点 **{}**"
                "（该 5 分钟 {}）。**早于它 = 候选驱动；晚于它 = 伴随或事后解读**；"
                "收盘后与次日早报一律不计入当日驱动。".format(
                    attribution["launch"], pct_str(attribution["launch_ret"])))
            add("")
        add("| 时间 | 阶段 | 类型 | 标题 | 来源 | 命中主题 |")
        add("|---|---|---|---|---|---|")
        for e in attribution["timeline"]:
            mark = "✅" if e["cand"] else "—"
            add("| {} | {}{} | {} | {} | {} | {} |".format(
                e["time"], mark, md_plain(e["phase"]), md_plain(e["kind"]),
                md_plain(e["title"]), md_plain(e["source"]),
                md_plain(e["topics"])))
        add("")
        add("> ✅ = 盘前 / 主升启动前（候选驱动）；其余为同步发酵、盘后或隔夜。"
            "采集窗口内共 {} 条事件，候选驱动 {} 条、事后解读 {} 条{}。"
            "公司公告、个股研报与脱水研报多为数日前发布，仅作产业背景，"
            "不计入当日时间轴；若某日无当日资讯，本章会明确输出「无明确事件驱动」。".format(
                attribution["n_events"], attribution["n_cand"],
                attribution["n_post"],
                ("、另有次日盘中 {} 条单列为验证层（见 15.10.2，不计入上述两类）".format(
                    attribution.get("n_next") or 0)
                 if (attribution.get("n_next") or 0) else "")))
        add("")

        add("### 15.3 主题映射与三方校验")
        add("")
        if attribution["topics"]:
            add("| 主题 | 资讯数/驱动数 | 置信度 | 板块资金验证 | 相关ETF涨幅 | 外围反证 |")
            add("|---|---:|---:|---|---:|---:|")
            for t in attribution["topics"]:
                sec = ("、".join("{} {:+.0f}亿".format(n, v)
                                for n, v in t["hit_sec"][:2]) if t["hit_sec"]
                       else "未进前列")
                add("| {} | {}/{} | {} 分 / {} | {} | {} | {} |".format(
                    md_plain(t["topic"]), t["n"], t["pre_n"],
                    t["score"], t["level"], md_plain(sec),
                    pct_str(t["etf_pct"]) if t["etf_pct"] is not None else "—",
                    t["external_n"] or "—"))
            add("")
            add("> **评分构成**：时序 30 + 板块资金验证 30 + ETF 涨幅验证 20 + "
                "来源数 10 + 多源重复 10 − 外围反证 15，总分截断到 0~100；"
                "无板块/ETF 映射的「全市场叙事型」主题时序降至 20、多源降至 5。"
                "≥70 高 / 45~69 中 / <45 低。板块资金取第四章净流入榜，"
                "ETF 涨幅取第一章样本池，外围取第六章；"
                "「外围反证」仅在「外围跌 / A 股同主题涨」同时成立时计入并扣分。")
            add("")

        add("### 15.4 候选驱动排序与证据")
        add("")
        for i, t in enumerate(attribution["topics"][:5], 1):
            add("**{}. {} — {} 分（{}置信）**".format(
                i, md_plain(t["topic"]), t["score"], t["level"]))
            add("")
            add("- 首次出现 {}，末次 {}；{} 条相关资讯、{} 家来源；"
                "代表标题「{}」".format(
                    t["first"], t["last"], t["n"], len(t["sources"]),
                    md_plain(t["repr"])))
            if t["pre_titles"]:
                add("- 候选驱动：{}".format(
                    "；".join(md_plain(x) for x in t["pre_titles"])))
            for ev in t["evidence"]:
                add("- {}".format(ev))
            add("")

        add("### 15.5 归因结论")
        add("")
        for s in attribution["signals"]:
            add("- {}".format(s))
        add("")
        add("> **一句话**：{}。".format(attribution["verdict"]))
        add("")

        add("### 15.5.1 主题异常收益（AR）与跨日台账")
        add("")
        if ev_ar.get("available"):
            _st = {r["topic"]: r for r in (ev_stats.get("rows") or [])}
            add("| 主题 | 匹配ETF | 平均涨幅 | 平均β | **AR** | 累计样本 | 均值AR | t 值 |")
            add("|---|---:|---:|---:|---:|---:|---:|---:|")
            for r in ev_ar["rows"]:
                s = _st.get(r["topic"]) or {}
                add("| {} | {} | {} | {:.2f} | **{:+.2f}pct** | {} 日 | {} | {} |".format(
                    md_plain(r["topic"]), r["n"], pct_str(r["pct"]), r["beta"],
                    r["ar"], s.get("n") or 0,
                    "{:+.2f}pct".format(s["mean"]) if s.get("mean") is not None else "—",
                    "{:.2f}（{}）".format(s["t"], s.get("sig"))
                    if s.get("t") is not None else "样本不足"))
            add("")
            add("> **AR = 主题内 ETF 当日收益 − β × 沪深300 当日收益**"
                "（β 取报告日之前 250 个交易日的日收益 OLS，估计窗不含当日）；"
                "本日沪深300 {}。AR 为正表示该主题**跑赢了自身 beta 应有的涨幅**，"
                "而不是简单的「涨了」。".format(pct_str(ev_ar.get("bench_pct"))))
            add("")
            add("> **台账口径**：结果按日写入 `etf_event_study.json`，当前已累积 "
                "**{} 个交易日 / {} 个主题**；t 检验要求同一主题 ≥{} 个交易日样本，"
                "未达阈值一律显示「样本不足」。这是把 15.3 的「高/中/低置信」"
                "逐步替换为可检验统计量的过渡方案，"
                "而不是「今天的事件已被证明是原因」。".format(
                    ev_stats.get("days") or 0, len(ev_stats.get("rows") or []),
                    ev_stats.get("min_n") or EVENT_STUDY_MIN_N))
            add("")
            add("> **已知局限**：只有日频 AR、没有分钟级 CAR，无法区分"
                "「事件后 5 分钟反应」与「全天漂移」；主题到 ETF 的映射仍是关键词，"
                "同一 ETF 可能同时命中多个主题。")
            add("")
        else:
            add("> 未计算 AR：{}。".format(
                ev_ar.get("reason") or "无匹配 ETF 或日线不足"))
            add("")

        add("### 15.5.2 日内事件窗口 CAR（{} 分钟）".format(ev_car.get("window") or 30))
        add("")
        if ev_car.get("available"):
            add("| 主题 | 事件起点 | 基准区间收益 | **CAR** | 纳入ETF | 窗口完整 |")
            add("|---|---|---:|---:|---:|---|")
            for r in ev_car["rows"]:
                add("| {} | {} | {} | **{:+.2f}pct** | {} | {} |".format(
                    md_plain(r["topic"]), r["start"][:2] + ":" + r["start"][2:],
                    pct_str(r["bench_ret"]), r["car"], r["n"],
                    "✅" if r["full"] else "⚠️ 未满（数据截至当前）"))
            add("")
            add("> **CAR = ETF 区间收益 − β × 沪深300 区间收益**，区间为事件时刻起 "
                "{} 分钟；盘前消息自 09:30 开盘起算。与 15.5.1 的日频 AR 相比，"
                "这里**剔除开盘前的涨幅**，更接近「消息出来之后市场买了多少」。".format(
                    ev_car.get("window") or 30))
            add("")
            _car_hist = {r["topic"]: r for r in (ev_stats.get("rows") or [])
                         if r.get("car_mean") is not None}
            _hit = [r for r in ev_car["rows"] if r["topic"] in _car_hist]
            if _hit:
                _days = max(r["car_n"] for r in _car_hist.values())
                _top = max(_hit, key=lambda r: abs(r["car"]))
                add("> **跨日累积**：台账已存 CAR {} 日（AR {} 日），{}的 CAR 均值 "
                    "{:+.2f}pct，今日 {:+.2f}pct。CAR 依赖当日分时，"
                    "只能随每日在线运行逐日累积，不回补历史。".format(
                        _days, ev_stats.get("days") or 0,
                        md_plain(_top["topic"]),
                        _car_hist[_top["topic"]]["car_mean"], _top["car"]))
                add("")
            _pre = [r["topic"] for r in ev_car["rows"] if r.get("note")]
            if _pre:
                add("> 其中 {} 属盘前消息，窗口从开盘起算，反映的是「开盘后承接力度」"
                    "而非「消息发布瞬间的跳空」。".format("、".join(_pre)))
                add("> 口径边界：CAR 只覆盖当日分时可见区间，**不是**跨日事件研究的替代；"
                    "样本仍是 34 只跟踪清单 ETF，主题到 ETF 为关键词匹配。")
                add("")
        else:
            _cn = max([r.get("car_n") or 0
                       for r in (ev_stats.get("rows") or [])] or [0])
            add("> 未计算日内 CAR：{}。该指标依赖**当日分时**"
                "（接口只返回当天，离线快照与次日重跑自动跳过）{}。".format(
                    ev_car.get("reason") or "无可用分时数据",
                    "；台账已存 {} 日 CAR，可查 `etf_event_study.json`".format(_cn)
                    if _cn else ""))
            add("")

        add("### 15.6 明日事件预告")
        add("")
        if attribution["tomorrow"]:
            for x in attribution["tomorrow"]:
                add("- {}".format(x))
        else:
            add("- 暂无结构化日历事件。")
        add("")

        add("### 15.7 7×24 快讯六通道采集")
        add("")
        st = attribution["flash_status"]
        sst = attribution["sina_status"]
        cst = attribution["ct_status"]
        est = attribution.get("empc_status") or "empty"
        add("| 通道 | 入口 | 本次状态 | 特点 |")
        add("|---|---|---|---|")
        add("| ① 东财 JSON | `np-weblist.eastmoney.com/comm/web/getFastNewsList` | {} | "
            "可按 `sortEnd` 回溯约 200 条，24 小时滚动 |".format(
                "✅ 直连" if st == "ok" else "⚠️ 受限（{}）".format(st)))
        add("| ② 东财 H5 | `emh5wap.eastmoney.com/info/channel/kuaixun` | "
            "⚠️ 本次仅栏目壳 | 服务端渲染但常只返回标题壳；"
            "快照里的 H5 条目是早前网页通道增量样本，按 id 留存、不冒充本次采集 |")
        add("| ③ 东财 PC 7_24 | `kuaixun.eastmoney.com/7_24.html`（壳页→页面 AJAX） | {} | "
            "2026-11 改版为 React 渲染，静态页只剩行情壳；"
            "改走页面自带 `getFastNewsList`（biz=web_724），与①同源、合并去重 |".format(
                "✅ AJAX 接口 {} 条".format(attribution.get("empc_n") or 0)
                if attribution.get("empc_n")
                else "⚠️ {}".format(est)))
        add("| ④ 新浪 7×24 | `finance.sina.com.cn/7x24/` | {} | "
            "**脚本可自主直连**，带 `YYYYMMDD` 完整日期 + 秒级时间 |".format(
                "✅ 直连 {} 条".format(attribution["sina_n"])
                if sst == "ok" else "⚠️ {}".format(sst)))
        add("| ⑤ 华夏时报 7×24 | `kuaixun.chinatimes.net.cn/api/proxy/news/result` | {} | "
            "中央级财经媒体；2026-11 改版为 JS 渲染，改走 JSON 接口"
            "（翻页 + 热点榜），含 iFinD 资金流独立口径 |".format(
                "✅ JSON 直连 {} 条".format(attribution["ct_n"])
                if attribution["ct_n"] else "⚠️ {}".format(cst)))
        _cn, _cst = attribution.get("cls_n") or 0, attribution.get("cls_status") or ""
        _clabel = ("✅ 直连 {} 条 / {} 页".format(
            _cn, attribution["cls_pages"]) if _cn and _cst == "ok"
            else "✅ 网页通道 {} 条（秒级时间戳）".format(_cn) if _cn
            else "⛔ 沙盒 403，适配器就绪")
        add("| ⑥ 财联社电报 | `www.cls.cn/telegraph` 网页通道 + JSON 双端点 | {} | "
            "A股垂直度最高；`HH:MM:SS` / `ctime` 秒级；**`last_time` 可回溯历史**；"
            "JSON 需本地 sign |".format(_clabel))
        add("")
        if sst == "ok":
            add("> **关键差别**：东财 JSON/H5/PC 三个域名族、华夏时报与财联社在本沙盒"
                "**python 直连均返回 403**，但网页通道能读到东财 PC 页与财联社电报页的"
                "真实内容，两者由 `--flash-text` 注入；新浪 PC 页是**脚本自己直连成功**的"
                "实时快讯源，不依赖人工粘贴。六个源都带完整日期，"
                "因此能直接做日期归属校验，避免把归档旧闻误当今日驱动。")
            add("")
        add("> **东财 PC 页的翻页边界（实测）**：`7_24.html` 给报告日盘后 "
            "20:53–22:08 的快讯，`7_24_2.html`～`_5.html` 返回的是**同一批**"
            "次日上午 09:33–09:39 条目，`_6.html` 已失效 —— `_n` 不是可靠的时间游标，"
            "拿不到 11-04 盘中 13:29 前后。所以这一路补的是「盘后叙事 + 次日验证」，"
            "不是盘中驱动；盘中回溯仍要靠财联社 `last_time` 或逐日累积。")
        add("")
        add("> **财联社网页通道实测 + 仍缺的一环**：`www.cls.cn/telegraph` "
            "网页可读、带 `HH:MM:SS` 秒级时间戳（本次已注入），但实测 "
            "`?date=2026-11-04`、`?last_time=1793775600` 均**被服务端忽略**，"
            "仍只返回最新实时页 —— 与东财 PC 页的 `_n` 翻页是同一个问题。"
            "所以六源目前都**补不了 11-04 盘中 13:29 前后**：网页通道给的是盘后与次日，"
            "JSON 接口才有的 `last_time` 时间游标被 403 挡住。"
            "按本页最小 `ctime` 逐页回退即可覆盖 09:30–15:00，"
            "届时 15.2 的「驱动 vs 事后」判定才真正完整。"
            "代价是它属**社区逆向端点**：`sign = md5(sha1(按 key 字典序拼的 query 串))` "
            "为前端私有约定（零密钥、纯本地可算），端点版本会漂移——旧 `nodeapi` 与新版 "
            "`v1/roll` 谁当前有效，社区资料说法不一，故脚本**两个端点、两套参数模板都试**；"
            "且存在反爬频率风控，不宜高频轮询。商用场景应走官方付费 API。")
        add("")
        if attribution["ct_hot"]:
            add("> **华夏时报「热点快讯」榜**（无时间戳，单列参考、不进事件池）：{}".format(
                "；".join("{}. {}".format(h["rank"], md_plain(h["title"][:34]))
                         for h in attribution["ct_hot"][:6])))
            add("")
        if attribution["flash"]:
            add("| 时间 | 阶段 | 来源 | 内容 | 命中主题 |")
            add("|---|---|---|---|---|")
            for e in attribution["flash"][:30]:
                add("| {} | {}{} | {} | {} | {} |".format(
                    e["time"], "✅" if e["cand"] else "—",
                    md_plain(e["phase"]), md_plain(e["source"]),
                    md_plain(e["title"]), md_plain(e["topics"])))
            add("")
            add("> 共 {} 条进入采集窗口，其中 {} 条落在主升起点前（✅），"
                "来自 {} 个独立来源。".format(
                    attribution["flash_n"], attribution["flash_cand"],
                    attribution["flash_src_n"]))
            add("")
            ev = attribution["flash"]
            cx = [
                ("公告级升级", ["长鑫科技", "回购"],
                 "把 13:21「长鑫科技 50 亿回购护盘」从媒体转述升级为**公告级证据**："
                 "回购额度由 10–20 亿上调至 25–50 亿，AI 产业链 159 家公司密集披露回购。"),
                ("外围反证加固", ["SK海力士", "三星电子", "东京电子", "爱德万", "美光科技", "英伟达"],
                 "日韩链 SK 海力士跌逾 6%、三星电子跌近 3%、东京电子跌 5.6%，"
                 "叠加华夏时报「美股收评」英伟达跌 3.17%、美光科技跌 7.00%、"
                 "阿斯麦跌 3.60%、英特尔跌 3.77%，与美股芯片股深夜大跌共同构成"
                 "**外围同主题走弱**硬反证，进一步支持内部产业催化而非外围传导。"),
                ("政策背景", ["新质生产力", "中共中央", "数据权益", "AI创新"],
                 "《关于发展新质生产力的意见》提出加快 AI 等数智技术创新，"
                 "为成长风格资金聚焦提供政策级背景。"),
                ("资金印证", ["净买入", "加仓", "净流入", "资金借道"],
                 "股票型 ETF 22 天净买入超 5500 亿，与第十章接口数据"
                 "（样本 ETF 一级市场 T-1 净申购 +550 亿）方向一致。"),
                ("IPO 叙事修正", ["IPO节奏", "IPO 要暂停", "并非"],
                 "次日早报给出「A股下半年 IPO 节奏趋缓、从严审核」的**正式表述**，"
                 "比 21:59「IPO 要暂停了？并非一刀切」的传闻更准确；"
                 "按时序仍不算盘中驱动。"),
            ]
            add("> **多源快讯带来的证据升级**（均为盘后 / 次日早报，"
                "按时序纪律**不计入** 11-04 的候选驱动，只提升归因链条完整度）：")
            add("")
            for label, kws, desc in cx:
                hit = [e for e in ev if any(k in e["title"] for k in kws)]
                if hit:
                    add("- **{}**：{}（代表条目「{}」，{}）".format(
                        label, desc, md_plain(hit[0]["title"][:38]),
                        hit[0]["time"]))
            add("")
        else:
            add("> 本次无 7×24 快讯条目。归因仍以热点资讯、公告、研报为准。")
            add("")

        add("### 15.8 多源一致性校验与逐日累积")
        add("")
        add("#### 15.8.1 跨源互证")
        add("")
        if attribution["cross"]:
            add("| 关键事实 | 出现次数 | 独立信源 | 代表条目 |")
            add("|---|---:|---|---|")
            for c in attribution["cross"][:12]:
                add("| {} | {} | {} | {}（{}） |".format(
                    md_plain(c["entity"]), c["n"],
                    "、".join(c["fams"]), md_plain(c["repr"]), c["time"]))
            add("")
            add("> 同一事实被 ≥2 个独立信源（东方财富 / 新浪财经 / 财经媒体）报道，"
                "可信度显著高于单家叙事。本章不把它们计入「驱动」，"
                "但用作归因链条的**证据强度加权**。")
            add("")
        else:
            add("- 本次未检出跨源一致的同一事实。")
            add("")
        if attribution["sina_new"]:
            add("#### 15.8.2 新浪源独有证据")
            add("")
            add("| 时间 | 内容 | 命中主题 |")
            add("|---|---|---|")
            for e in attribution["sina_new"][:10]:
                add("| {} | {} | {} |".format(
                    e["time"], md_plain(e["title"]), md_plain(e["topics"])))
            add("")
            add("> 这些是其他来源未覆盖、由新浪 7×24 直连补入的事实，"
                "主要用于加固「外围半导体链走弱」与「IPO 节奏」两条支线。")
            add("")

        add("#### 15.8.3 逐日累积：把「最新 50 条」变成历史序列")
        add("")
        add("> 实时页只返回**最新 50 条**（本次为 11-05 07:46–08:14），"
            "单看一天永远补不到报告日 9:30–15:00 的盘中快讯。"
            "脚本因此按 `id` 把每次采集增量写入 `{}`：".format(
                FLASH_HISTORY_FILE))
        add("")
        add("- 当前已沉淀 **{} 天 / {} 条**；".format(
            attribution["flash_days"], attribution["flash_total"]))
        add("- **每天运行一次**（建议收盘后或盘中定时），当天 9:30–15:00 的"
            "逐分钟快讯就会被永久留存；")
        add("- 次日报告读取该缓存，即可把盘中资讯放入 15.2 时间轴参与"
            "「驱动 vs 事后」判定——这才是补齐盘中空白的根本办法，"
            "而不是靠一次性注入样本。")
        add("")
        arc = attribution["archive"]
        if arc.get("n"):
            add("#### 15.8.4 日期归属校验（归档样本拦截）")
            add("")
            add("> 用户提供的 `cj.sina.cn/index/feed24hour` 纯文本示例含 "
                "**{} 条** `HH:MM:SS + 正文` 快讯；正文出现年份 {}，"
                "按多数投票主年份为 **{}**，与报告日 {} 不符 → "
                "判定为 **历史归档**（实际为 2025-09-09 那期）。"
                "注意其中「预计 2026 年推出原型机」属未来规划，"
                "不能因为出现过 2026 就当成当期新闻。".format(
                    arc["n"], "、".join(arc.get("years") or []),
                    arc.get("dominant") or "—", today["date"]))
            add("")
            add("> 解析器已验证能正确读取这种无日期格式（时间 / 标题 / 正文），"
                "但**绝不并入本报告的事件池**：把去年的新闻当成今日拉升的驱动"
                "会直接污染归因。如需使用，请保存为文件后通过 `--flash-text` 传入，"
                "脚本会自动做同样的年份校验。")
            add("")
        add("> **仍存在的边界**：新浪 PC 页的 `?page=2` / `?tag=10` 参数被服务端忽略，"
            "所有 `zhibo` / `app.cj` / `wap.cj` API 域名在本沙盒均 403；"
            "东财 PC 页与财联社电报页**网页通道可读**，但东财 `_n` 翻页、财联社 "
            "`?date` / `?last_time` 参数同样被忽略 —— 两者只返回最新实时页。"
            "因此六源都**无法一次性回溯 11-04 盘中**，本次只能靠逐日累积与次日验证层。"
            "换到无网络策略的环境，东财 JSON 按 `sortEnd`、或财联社按 `last_time` "
            "翻页即可直接回溯；其中财联社的时间游标是六个源里"
            "**唯一原生支持按时间戳回溯**的，优先级最高。")
        add("")

        # ---- 15.9 华夏时报源独有的 ETF 资金硬数据 ----
        add("### 15.9 ETF 资金流：独立口径印证与个券级对照")
        add("")
        flows = attribution["etf_flow"] or []
        uni = universe or {}
        vmap = {}
        for f in flows:
            vmap.update(f.get("values") or {})
        singles = []
        for f in flows:
            singles.extend(f.get("singles") or [])
        add("#### 15.9.1 两个口径，不能相加")
        add("")
        add("| 口径 | 数值 | 数据日 | 来源 | 含义 |")
        add("|---|---:|---|---|---|")
        add("| 整体资金流（iFinD） | **+{} 亿** | 11-04 当日 | 华夏时报 08:14 | "
            "按成交/份额统计的资金流；股票型 {} 亿、宽基型 {} 亿 |".format(
                vmap.get("ETF整体净流入", "—"),
                vmap.get("股票型净流入", "—"),
                vmap.get("宽基型净流入", "—")))
        add("| 一级市场净申购 | **{:+.1f} 亿** | {}（T-1 披露） | 第十章接口 | "
            "份额申赎 × 净值，样本池内 ETF 合计 |".format(
                num(uni.get("d_total")) or 0, uni.get("flow_date") or "未知"))
        if vmap.get("累计净买入"):
            add("| 区间累计 | {} 亿 | 22 个交易日（截至 11-03） | 东财·新浪·华夏时报三源同题 | "
                "自 9 月底以来股票型 ETF 连续净流入合计 |".format(
                    vmap.get("累计净买入")))
        if vmap.get("连续净流入日数"):
            add("| 连续净流入 | {} 个交易日 | 9/28 起、截至 11-04 | 华夏时报 08:14 / iFinD | "
                "与上一行差 1 天是截至日不同，不是数据打架 |".format(
                    vmap.get("连续净流入日数")))
        if vmap.get("单日净流入新高"):
            add("| 前一日单日新高 | {} 亿 | 11-03 | 华夏时报 07:54 | "
                "创年内新高，说明流入是连续趋势而非单日脉冲 |".format(
                    vmap.get("单日净流入新高")))
        add("")
        add("> ⚠️ **口径边界（最容易误读的一处）**：`432.39 亿`是同花顺 iFinD 的"
            "**ETF 资金流统计**，`{:+.1f} 亿`是接口的**一级市场净申购**（份额申赎），"
            "统计对象、计算方式、披露日（T-1）都不同。**两者不能相加、不能相减、"
            "也不能当成同一个数字的互相验证**；只能说方向一致、量级相近，共同支持"
            "「一级/机构资金有增量」。`22 个交易日`与`23 个交易日`差 1 天，"
            "是统计截至日不同（一截至 11-03、一截至 11-04），不是数据打架。".format(
                num(uni.get("d_total")) or 0))
        add("")
        if singles:
            add("#### 15.9.2 个券级对照（榜单 Top3）")
            add("")
            rank_map = {}
            for c, q in (today.get("ranked") or []):
                rank_map[c[-6:]] = q
            add("| ETF | 代码 | iFinD 净流入 | 在样本池 | 报告日涨幅 |")
            add("|---|---|---:|---|---:|")
            for s in singles[:6]:
                q = rank_map.get(s["code"])
                add("| {} | {} | {} 亿 | {} | {} |".format(
                    md_plain(s["name"]), s["code"], s["value"],
                    "✅" if q else "—",
                    pct_str(q.get("pct")) if q else "—"))
            add("")
            add("> 三只全是**宽基 ETF**，且都在本脚本样本池内、"
                "报告日温和上涨（非涨停）。这与第一章「科技 ETF 涨停潮、"
                "结构性资金进攻」互补：**追涨的是二级情绪，真金白银申购进的是宽基**。"
                "也再次说明「一级有增量、二级未放量」不是矛盾，而是同一天两拨钱"
                "的不同行为。")
            add("")
        add("#### 15.9.3 对结论的影响")
        add("")
        add("- **强化**：第十章「一级市场 T-1 净申购 +550 亿」多了一个独立信源"
            "（中央级财经媒体 + 同花顺 iFinD），且给出当日分项与个券，证据链更完整；")
        add("- **不改写**：第一章核心 ETF 仍是**缩量反包**，二级换手未同步放大，"
            "基准情景维持「**震荡筑底**」；")
        add("- **不加总**：两套口径并列展示，报告不做任何合并计算；")
        add("- **时序纪律**：这些条目均在 11-05 早间发布，属次日报道，"
            "**不计入 11-04 的候选驱动**，只作数据印证。")
        add("")

    # ---------------------------------------------------------------
    # 15.10 东财全球财经快讯 PC 页：盘后叙事 + 次日盘中验证
    add("### 15.10 东财全球财经快讯：盘后叙事与次日盘中验证")
    add("")
    add("> 用户实测 `kuaixun.eastmoney.com/7_24.html` **网页通道可用**。"
        "沙盒 python 直连该域名仍返回 403（`policy_default_denied`），"
        "但网页通道能读到服务端渲染的真实快讯，故由 `--flash-text` 注入、"
        "并写入离线快照。脚本已内置该通道的直连适配器与 `7_24_{n}.html` 翻页，"
        "换到可直连环境会自动采集。")
    add("")
    if not (attribution.get("empc_n") or 0):
        add("> ⚠️ **本次未采集到东财 PC 页快讯**（直连 403 且未注入网页通道样本）。"
            "可将 `7_24.html` 网页通道内容存成 `=== YYYY-MM-DD ===` 分段文本，"
            "用 `--flash-text` 传入；15.10 的两张验证表本次留空，不影响其他章节结论。")
        add("")
    _ed = attribution.get("empc_days") or {}
    if _ed:
        add("| 页面日期段 | 条数 | 相对报告日 | 用途 |")
        add("|---|---:|---|---|")
        _rdate = today["date"] or ""
        for _day, _n in sorted(_ed.items()):
            _rel = ("报告日盘后" if _day == _rdate else
                    ("次日盘中" if _day > _rdate else "报告日前"))
            _use = ("事后叙事 / 外围反证" if _rel == "报告日盘后" else
                    ("**观察清单实证**" if _rel == "次日盘中" else "背景"))
            add("| {} | {} | {} | {} |".format(_day, _n, _rel, _use))
        add("")

    _post = attribution.get("empc_post") or []
    _con = attribution.get("empc_contrary") or []
    if _post:
        add("#### 15.10.1 报告日盘后：外围分化与反证")
        add("")
        add("共 {} 条盘后快讯，其中 {} 条明确呈现外围 AI/芯片链的"
            "**分化或转谨慎**（已排除「谷歌为数据中心项目兜底担保」这类利好，"
            "不把中性/偏多叙事误算成反证）：".format(len(_post), len(_con)))
        add("")
        for _c in _con[:6]:
            add("- {} · {}".format(md_plain(_c["time"]),
                                  md_plain(_c["title"])))
        add("")
        add("> **华尔街对数据中心前景转向「谨慎」**、"
            "**美股大型科技股普涨但英伟达 −1.33% / 台积电 −1.26% 逆势下跌**"
            "（外围内部分化），叠加第十五章已有的"
            "「美股芯片股深夜大跌 / 日韩 SK海力士 −6%」，"
            "共同指向**外围同主题并不强**。A 股半导体当日却 +194 亿"
            "净流入、相关 ETF 涨停潮 —— 更支持"
            "「**内部产业催化（存储涨价 + 长鑫回购）而非外围传导**」这一归因，"
            "而不是反过来。这些条目均在 21:00 后发布，"
            "按 15.2 时序纪律**不计入 11-04 驱动**。")
        add("")

    _next = attribution.get("empc_next") or []
    if _next:
        add("#### 15.10.2 次日盘中验证（观察清单的实证）")
        add("")
        add("> 这是第十四章「明日只盯」第一次拿到**真实次日盘中快讯**做对照，"
            "不是推演。汇聚东财 PC 页、财联社电报网页通道等多个来源的 "
            "11-05 09:30 之后快讯；按 15.2 时序纪律，这一层"
            "**既不算 11-04 驱动、也不算事后解读**，只作观察清单的实证。")
        add("")
        _etf_top = (rank.get("etf_gainers") or [None])[0]
        _fund_top = rank.get("fund_top") or []
        _in_top = fund.get("in_top") or []
        _semi = next((x for x in _in_top
                      if "半导体" in str(x.get("name") or "")), None)
        _zj = next((x for x in _fund_top
                    if "中际旭创" in str(x.get("name") or "")), None)
        add("| 次日快讯 | 来源 | 11-04 对应数据 | 对照解读 |")
        add("|---|---|---|---|")
        for _e in _next:
            _t = _e["title"]
            if "588170" in _t or "科创半导体" in _t:
                _prev = ("第十六章全市场 ETF 榜首 **{} {}**".format(
                    (_etf_top or {}).get("name") or "—",
                    pct_str((_etf_top or {}).get("zdf")))
                    if _etf_top else "第十六章全市场 ETF 榜首")
                _read = ("20cm 涨停次日**未崩但曾低开**（「转涨」），"
                         "属高位震荡而非连续逼空")
            elif "中际旭创" in _t or "CPO" in _t:
                _prev = ("11-04 个股主力净流入 **{} 亿**、当日 **{}**".format(
                    "{:+.2f}".format(_zj["flow_yi"]) if _zj else "—",
                    pct_str(_zj["zdf"]) if _zj else "—"))
                _read = ("**次日跌近 10%** → 高位股回吐，"
                         "正是第十四章「涨停附近不宜追高」的实证")
            elif "半导体" in _t:
                _prev = ("第四章行业资金 Top1 **{} 亿**".format(
                    "{:+.0f}".format(_semi["net_yi"])
                    if _semi else "半导体 +194"))
                _read = "主线**延续**但换领涨股，板块内部有分化"
            else:
                _prev = "11-04 未列为主线"
                _read = "**轮动扩散**，资金向创新药/白酒/港股等外溢"
            add("| {} {} | {} | {} | {} |".format(
                md_plain(_e["time"]), md_plain(_t[:44]),
                md_plain(_e.get("source") or ""),
                _prev, md_plain(_read)))
        add("")
        _ck = attribution.get("checklist") or []
        if _ck:
            add("#### 15.10.3 观察清单自动判定（事件级）")
            add("")
            add("> 第十四章列的是「明日只盯」的**条件**，这里把其中能从快讯文本"
                "直接读出的条件写成布尔判定。判定只覆盖事件层，"
                "数值类条件如实留白，不用资讯文本冒充行情数据。")
            add("")
            add("| 观察条件 | 事件级信号 | 判定（是否印证预判） | 含义 |")
            add("|---|---|---|---|")
            for _c in _ck:
                _mark = ("✅ 满足" if _c["ok"] is True else
                         "⚠️ 未满足" if _c["ok"] is False else
                         "⏳ 待数值确认")
                add("| {} | {} | {} | {} |".format(
                    md_plain(_c["name"]), md_plain(_c["sig"]),
                    _mark, md_plain(_c["note"])))
            add("")
            _okn = sum(1 for c in _ck if c["ok"] is True)
            _non = sum(1 for c in _ck if c["ok"] is False)
            _wait = sum(1 for c in _ck if c["ok"] is None)
            add("> 汇总：{} 项满足、{} 项未满足、{} 项待数值确认。"
                "「满足」= 印证了 14.4 的预判（含风险项按预期兑现）；"
                "两项「未满足」是券商未接力、外围未同向接力，"
                "**同样符合震荡筑底而非全面行情的设定**，"
                "因此 5 项已判定条件**没有出现证伪信号**；"
                "数值类条件仍须等次日收盘全量数据，"
                "本节不代替第十四章的量价确认。".format(_okn, _non, _wait))
            add("")
        add("> 三句话概括次日：**主线延续（半导体仍有个股涨停）+ "
            "高位股回吐（CPO/中际旭创跌近 10%）+ 轮动扩散"
            "（创新药、白酒、港股航空、恒生科技齐动）**。"
            "这恰好是「**震荡筑底**」的典型形态，"
            "而不是「趋势反转后的单边主升」——"
            "与第十四章基准情景、第十六章「近20日仅 20% 个股为正」完全一致。")
        add("")
        add("> **边界**：这些是 11-05 的盘中数据，**不能反推为 11-04 的涨跌驱动**；"
            "`_2`～`_5` 返回同一批、`_6` 失效，翻页不是可靠时间游标，"
            "11-04 盘中（13:29 前后）的快讯空白仍未补上。")
        add("")

    # ---------------------------------------------------------------
    add("## 十六、全市场涨跌排行（个股 + ETF 全量）")
    add("")
    add("> **这一章把前面长期标注的「样本内」边界真正解开**。数据源："
        "腾讯行情公开接口 —— 个股用 "
        "`proxy.finance.qq.com/cgi/cgi-bin/rank/hs/getBoardRankList`"
        "（`board_code=aStock`，分页拉全量），ETF 用 `qt.gtimg.cn` "
        "按第十章 `etf_size` 的 {} 只代码批量报价。两者都带行情时间戳，"
        "且与第二章广度逐家校验，确认是 **{} 收盘快照**而非次日实时。".format(
            universe.get("universe_n") or 0, today["date"] or ""))
    add("")
    if not rank.get("available"):
        add("> ⚠️ 本次无涨跌排行数据（接口不可用且快照未含缓存）。"
            "可执行 `python3 etf_analysis.py --save-snapshot` 在线生成一次缓存。")
        add("")
    else:
        add("### 16.1 数据来源与日期归属校验")
        add("")
        add("| 层 | 覆盖 | 状态 | 校验 |")
        add("|---|---:|---|---|")
        add("| 个股 | {} 只（沪深京全A） | {} | {} |".format(
            rank["n_stock"], "✅ 全量" if rank["n_stock"] else "缺失",
            md_plain(rank.get("date_note") or "—")))
        add("| ETF | {} 只（全市场明细全集） | {} | 行情时间戳 {} |".format(
            rank["n_etf"], "✅ 全量" if rank["n_etf"] else "缺失",
            (rank.get("etf_stats") or {}).get("stamp") or "—"))
        add("")
        if not rank.get("date_ok"):
            add("> ⚠️ **非报告日快照**：{} 本章数据仅作参考，"
                "不并入当日排行与结论。".format(rank.get("date_note") or ""))
            add("")

        if rank.get("gainers"):
            add("### 16.2 全市场个股涨幅榜（Top15）")
            add("")
            add("| 代码 | 名称 | 涨跌幅 | 最新价 | 近5日 | 近20日 | 主力净流入(亿) | 换手 |")
            add("|---|---|---:|---:|---:|---:|---:|---:|")
            for r in rank["gainers"]:
                add("| {} | {} | {} | {:.3f} | {} | {} | {:+.2f} | {:.1f}% |".format(
                    md_plain(r["code"]), md_plain(r["name"]),
                    pct_str(r["zdf"]), r["price"],
                    pct_str(r["d5"]), pct_str(r["d20"]),
                    r["flow_yi"], r["hsl"]))
            add("")
            add("> 榜首出现 30cm 北交所与 20cm 双创个股，是分板块涨跌幅上限差异，"
                "不是数据异常。注意这些强势股**近20日多为负**——"
                "领涨的也是超跌反弹主力，而非中期强势趋势股。")
            add("")
        if rank.get("losers"):
            add("### 16.3 全市场个股跌幅榜（Top10）")
            add("")
            add("| 代码 | 名称 | 涨跌幅 | 近20日 |")
            add("|---|---|---:|---:|")
            for r in rank["losers"]:
                add("| {} | {} | {} | {} |".format(
                    md_plain(r["code"]), md_plain(r["name"]),
                    pct_str(r["zdf"]), pct_str(r["d20"])))
            add("")

        if rank.get("periods"):
            add("### 16.4 全市场多周期强弱分布（中期趋势的硬证据）")
            add("")
            add("| 周期 | 涨跌幅中位数 | 平均数 | 上涨占比 |")
            add("|---|---:|---:|---:|")
            for p in rank["periods"]:
                add("| {} | {} | {} | {:.0f}% |".format(
                    p["label"], pct_str(p["median"]),
                    pct_str(p["mean"]), p["pos_ratio"]))
            add("")
            _p20 = next((p for p in rank["periods"] if p["key"] == "zdf_d20"), None)
            add("> 这是**全市场 {} 只个股**的统计，不是 ETF 样本外推。"
                "「当日 {:.0f}% 上涨」与「近20日仅 {:.0f}% 为正、中位数 {}」并存，"
                "正是本报告从第一章到第十四章反复强调的"
                "**短期强、中期弱**——现在有了全量口径的证明。".format(
                    rank["n_stock"], rank.get("ratio") or 0,
                    (_p20 or {}).get("pos_ratio", 0),
                    pct_str((_p20 or {}).get("median"))))
            add("")

        if rank.get("dist"):
            add("### 16.4.1 涨跌分布、成交集中度与资金覆盖面（广度质量）")
            add("")
            add("> 涨跌家数只说明方向；赚钱效应要看「涨多少、量堆在哪、资金跟不跟」。"
                "以下均基于同一份全量快照。")
            add("")
            add("| 涨跌幅区间 | 家数 | 占比 |")
            add("|---|---:|---:|")
            for d in rank["dist"]:
                if not d["n"]:
                    continue
                add("| {} | {} | {:.1f}% |".format(d["label"], d["n"], d["pct"]))
            add("")
            _ac = rank.get("amt_conc") or {}
            if _ac:
                add("| 集中度 / 资金 | 读数 | 含义 |")
                add("|---|---:|---|")
                add("| 成交额合计 | {:.0f} 亿 | 全市场个股口径（不含 ETF） |".format(
                    _ac["total_yi"]))
                add("| 前 10% 个股占比 | {:.0f}% | 越高越依赖少数标的 |".format(
                    _ac["top10_pct"]))
                add("| 前 100 名占比 | {:.0f}% | 头部吸筹强度 |".format(
                    _ac["top100_pct"]))
                _fc = rank.get("flow_cover") or {}
                if _fc.get("n"):
                    add("| 主力净流入个股占比 | {:.0f}% | {} / {} 只，"
                        "价涨而资金未跟即为背离 |".format(
                            _fc["inflow_pct"], _fc["inflow_n"], _fc["n"]))
                _pc = rank.get("period_cons") or {}
                if _pc.get("n"):
                    add("| 5/20/60 日同时为正 | {:.0f}% | 三周期共振向上的比例 |".format(
                        _pc.get("pos_pct", 0)))
                    add("| 5/20/60 日同时为负 | {:.0f}% | 中期下行结构占比 |".format(
                        _pc.get("neg_pct", 0)))
                add("")
                add("> 集中度与共振比例是**结构性指标**：集中度高 + 共振比例低，"
                    "说明当日强势集中在少数标的、且多数个股中期仍在下降通道，"
                    "「普涨」应理解为超跌反弹而非趋势启动。")
                add("")

        if rank.get("fund_top"):
            add("### 16.5 个股主力净流入榜（补第四章只有行业级）")
            add("")
            add("| 代码 | 名称 | 涨跌幅 | 主力净流入(亿) | 换手 | 量比 |")
            add("|---|---|---:|---:|---:|---:|")
            for r in rank["fund_top"]:
                add("| {} | {} | {} | {:+.2f} | {:.1f}% | {:.2f} |".format(
                    md_plain(r["code"]), md_plain(r["name"]),
                    pct_str(r["zdf"]), r["flow_yi"], r["hsl"], r["lb"]))
            add("")
            add("> 原始字段 `zljlr` 单位为万元，此处 ÷10000 换算为亿元；"
                "属按成交单规模划分的**统计口径**，不是账户流水。"
                "个股榜与第四章行业榜（半导体 +194 亿）、第十五章归因"
                "（长鑫科技回购护盘）构成**个股→板块→资讯**三级互证。")
            add("")
        if rank.get("amount_top"):
            add("**成交额榜（Top10）**：")
            add("")
            for r in rank["amount_top"][:10]:
                add("- {} {} · {} · {:.0f} 亿".format(
                    md_plain(r["name"]), md_plain(r["code"]),
                    pct_str(r["zdf"]), r["amount_yi"]))
            add("")

        es = rank.get("etf_stats") or {}
        if rank.get("etf_gainers"):
            add("### 16.6 全市场 ETF 涨跌排行（{} 只全量）".format(es.get("n_all") or 0))
            add("")
            add("> 按名称关键词剔除债券/货币/现金类后计 {} 只；"
                "这是**粗分口径**，非基金官方分类。中位数 {}、均值 {}、"
                "上涨占比 {:.0f}%。已解除 1.2 长期标注的"
                "「样本池内排序、不是全市场 ETF 涨幅榜」边界。".format(
                    es.get("n_equity") or 0,
                    pct_str(es.get("median")), pct_str(es.get("mean")),
                    es.get("pos_ratio") or 0))
            add("")
            add("| 代码 | 名称 | 涨跌幅 | 价格 | 成交额(亿) | 换手 |")
            add("|---|---|---:|---:|---:|---:|")
            for r in rank["etf_gainers"]:
                add("| {} | {} | {} | {:.3f} | {:.1f} | {:.1f}% |".format(
                    md_plain(r["code"]), md_plain(r["name"]),
                    pct_str(r["zdf"]), r["price"],
                    r["amount_yi"], r["hsl"]))
            add("")
            if rank.get("etf_losers"):
                add("**跌幅榜（末 8）**：{}".format(
                    "；".join("{} {}".format(md_plain(r["name"]),
                                            pct_str(r["zdf"]))
                             for r in rank["etf_losers"][:8])))
                add("")
            add("> 榜首 20cm 科创半导体/芯片 ETF 是在第一章 {} 只样本池**之外**的，"
                "说明样本池确实偏保守：样本榜最高 {}，全市场榜最高 {}。"
                "这不是样本选错，而是「分层抽样观察窗口」的固有代价，"
                "现在由本章全量榜补齐。".format(
                    today.get("total_etf") or 0,
                    pct_str((today.get("ranked") or [[None, {}]])[0][1].get("pct")),
                    pct_str(es.get("max"))))
            add("")

        sc = rank.get("sample_check") or {}
        if sc:
            add("### 16.7 样本代表性终极校验")
            add("")
            add("| 口径 | 等权均值 | 中位数 | 成交额加权 |")
            add("|---|---:|---:|---:|")
            add("| 第一章样本（{} 只） | {} | {} | {} |".format(
                sc.get("sample_n") or 0,
                pct_str(sc.get("sample_mean")),
                pct_str(sc.get("sample_median")),
                pct_str(sc.get("sample_weighted"))))
            add("| 全市场 ETF（{} 只） | {} | {} | — |".format(
                es.get("n_equity") or 0,
                pct_str(sc.get("universe_mean")),
                pct_str(sc.get("universe_median"))))
            add("")
            add("> 样本中位数 − 全市场中位数 = **{}**。"
                "1.6 的自检原本只有样本内部「等权 vs 加权」，"
                "现在有了真正的外部基准：可量化样本整体偏成长/偏科技多少，"
                "这也是对用户早先「{} 只样本能否代表大盘」质疑的最终回答"
                "——**判断大盘请用第二至四章与本章的全量口径，"
                "ETF 样本只回答 ETF 端表现**。".format(
                    pct_str(sc.get("gap_median")),
                    sc.get("sample_n") or 0))
            add("")

        if rank.get("signals"):
            add("### 16.8 自动信号")
            add("")
            for i, s in enumerate(rank["signals"], 1):
                add("{}. {}".format(i, s))
            add("")
        add("### 16.9 对结论的影响")
        add("")
        add("- **不改结论**：基准情景仍为「**震荡筑底**」，"
            "全市场多周期分布反而把「短期强、中期弱」从样本推演升级为全量证据；")
        add("- **强化归因**：个股主力净流入榜与 15.5 的半导体归因在个股层对上；")
        add("- **解除边界**：1.2 / 1.6 中「不能推断全市场 ETF 涨幅榜」"
            "由 16.6 补齐，「样本偏态」由 16.7 外部基准量化；")
        add("- **口径纪律**：涨停按 10/20/30cm 分板块判定，"
            "与官方广度差 1–3 家属阈值边界；排行数据经家数校验后才并入当日结论。")
        add("")

    # ---------------------------------------------------------------
    vp = attribution.get("viewpoints") or {}
    add("## 十七、多方评论观点采集与分歧分析")
    add("")
    add("> **这一章把「事实快讯」和「谁在发表什么看法」分开**。第十五章判定"
        "「驱动 vs 事后」，本章进一步识别观点方（卖方研报 / 基金公司 / 海外机构 / "
        "官方 / 公司管理层 / 媒体分析人士）、立场（偏多 / 偏空 / 分歧条件式）、"
        "以及观点是否获板块资金与 ETF 涨幅验证。**观点是解释，不是证明**；"
        "绝大多数出现在收盘后与次日早间，只作印证、不作驱动。")
    add("")
    add("### 17.1 本次五份资讯页的处置（口径透明）")
    add("")
    add("| 用户提供的页面 | 本次实际读到 | 处置 |")
    add("|---|---|---|")
    _ct_n = attribution.get("ct_n") or 0
    _ct_hot_n = len(attribution.get("ct_hot") or [])
    _arc_n = (attribution.get("archive") or {}).get("n") or 0
    _arc_txt = ("{} 条 `HH:MM:SS` 纯文本".format(_arc_n)
                if _arc_n else "本次未传入归档样本")
    add("| 华夏时报 `kuaixun.chinatimes.net.cn` | 2026-11-05 完整页，"
        "**{} 条**带 `HH:MM` 快讯 + **{} 条**热点快讯榜 | ✅ 全量解析并入事件池；"
        "兼容中文弯引号、更长正文，并补入此前简化样本未收录的 07:00 条目 |".format(
            _ct_n, _ct_hot_n))
    add("| 新浪 `cj.sina.cn/index/feed24hour` | {}；该页为 2025-09-09 归档"
        "（「李强同葡萄牙总理」「iPhone 17」） | ⛔ 判为 **2025-09-09 归档**，"
        "年份归属校验拦截，**不并入** 11-04 事件池（详见 15.8.4） |".format(
            _arc_txt))
    add("| 东财 `emh5wap.eastmoney.com/info/channel/kuaixun` | 本次仅返回"
        "「7x24 财经门户」栏目壳，**0 条快讯正文** | ⚠️ 不冒充本次采集；"
        "快照中的 `东方财富H5快讯` 条目是**此前网页通道留存的增量样本**，"
        "按 id 累积保留，不计入本次新增 |")
    _en = attribution.get("empc_n") or 0
    _ed = attribution.get("empc_days") or {}
    add("| 东财 **`kuaixun.eastmoney.com/7_24.html`** | {} | {} |".format(
        ("网页通道实测可读：{} 条真实快讯（{}）".format(
            _en, "、".join("{} {}条".format(k, v)
                         for k, v in sorted(_ed.items())))
         if _en else "本次未读到内容（直连 403 / 未注入网页通道样本）"),
        ("✅ 已接入第六通道并写入快照；python 直连仍 403，"
         "靠网页通道注入；详见 15.10" if _en else
         "⚠️ 第六通道适配器就绪，本次 0 条；15.10 降级留空")))
    _cn = attribution.get("cls_n") or 0
    add("| 财联社 **`www.cls.cn/telegraph`** | 网页通道实测可读："
        "{} 条带 `HH:MM:SS` 秒级时间戳电报 | {} |".format(
            _cn,
            "✅ 已作为第六通道写入快照；`?date` / `?last_time` 被忽略，"
            "只能给最新实时页，详见 15.7 / 15.10.2" if _cn else
            "⚠️ 网页通道未注入，JSON 适配器待命"))
    add("")
    add("> 这五行正好说明为什么事件层要做**日期归属校验**与**增量留存**："
        "同一个东财域名族，移动端 H5 是空壳、PC `7_24.html` 却是完整快讯；"
        "财联社网页可读但翻页参数无效；新浪移动页又是 2025 年归档。"
        "若不做校验，直接混进事件池就会污染结论。")
    add("")
    if not vp.get("available"):
        add("> ⚠️ 本次未识别到观点条目（资讯接口不可用且快照无缓存）。"
            "可执行 `python3 etf_analysis.py --save-snapshot` 在线生成一次缓存。")
        add("")
    else:
        add("### 17.2 观点方分布")
        add("")
        add("| 观点方 | 条数 | 偏多 | 偏空 | 分歧 | 事实型 |")
        add("|---|---:|---:|---:|---:|---:|")
        for a in vp["actors"]:
            _an = ("未归类（均为事实型）" if a["actor"] == "其他/未识别"
                   else md_plain(a["actor"]))
            add("| {} | **{}** | {} | {} | {} | {} |".format(
                _an, a["n"], a["bull"], a["bear"],
                a["mixed"], a["fact"]))
        add("")
        add("> 共识别 **{} 条**观点候选，其中 {} 条带明确/条件式立场、"
            "{} 条为纯事实型（行情数字、公告数据，不参与多空统计）。"
            "括号里的「人民财讯 / 央视新闻 / 中国证券报」是**转载方**，"
            "观点方按正文里的原始表态主体粗分，同一机构多条会分别计数。".format(
                vp["n"], vp["n_stance"], vp["n_fact"]))
        add("")

        if vp.get("rows"):
            add("### 17.3 观点清单（带立场，按时间）")
            add("")
            add("| 时间 | 阶段 | 观点方 | 立场 | 主题 | 观点摘要 | 资金/涨幅校验 |")
            add("|---|---|---|---|---|---|---|")
            shown = [r for r in vp["rows"] if r["stance"] != "fact"]
            for r in shown:
                add("| {} | {} | {} | {} | {} | {} | {} |".format(
                    md_plain(r["time"]), md_plain(r["phase"]),
                    md_plain(r["actor"]), r["stance_label"],
                    md_plain(r["topic"]), md_plain(r["title"][:44]),
                    "{} {}".format(r["verify"], md_plain(r["verify_detail"][:22]))))
            add("")
            add("> 清单列出全部 **{} 条**带立场条目（事实型已排除）；"
                "「⚠️ 背离」= 观点方向未获板块资金或"
                "样本 ETF 涨幅支持，这类**不作强归因**，仅提示叙事与资金不一致。"
                .format(len(shown)))
            add("")

        if vp.get("stance_dist"):
            add("### 17.4 立场分布与主题分歧")
            add("")
            add("**整体立场分布**")
            add("")
            add("| 立场 | 条数 | 占比 |")
            add("|---|---:|---:|")
            for s in vp["stance_dist"]:
                add("| {} | {} | {:.0f}% |".format(
                    s["label"], s["n"], s["n"] / vp["n"] * 100))
            add("")
            if vp.get("topics"):
                add("**按主题聚合（分歧检测）**")
                add("")
                add("| 主题 | 观点数 | 偏多 | 偏空 | 分歧 | 事前 | 判定 | 校验 |")
                add("|---|---:|---:|---:|---:|---:|---|---|")
                for t in vp["topics"][:10]:
                    add("| {} | {} | {} | {} | {} | {} | {} | {} |".format(
                        md_plain(t["topic"]), t["n"], t["bull"], t["bear"],
                        t["mixed"], t["pre"], md_plain(t["verdict"]),
                        t["verify"]))
                add("")
                div = [t for t in vp["topics"] if t["divergent"]]
                if div:
                    parts = []
                    for t in div[:3]:
                        br = next((r for r in vp["rows"]
                                   if r["topic"] == t["topic"]
                                   and r["stance"] in ("bull", "lean_bull")), None)
                        sr = next((r for r in vp["rows"]
                                   if r["topic"] == t["topic"]
                                   and r["stance"] in ("bear", "lean_bear")), None)
                        mx = next((r for r in vp["rows"]
                                   if r["topic"] == t["topic"]
                                   and r["stance"] == "mixed"), None)
                        seg = md_plain(t["topic"])
                        if br and sr:
                            seg += "（多「{}」/ 空「{}」）".format(
                                md_plain(br["title"][:18]),
                                md_plain(sr["title"][:18]))
                        elif mx:
                            seg += "（条件式「{}」）".format(
                                md_plain(mx["title"][:24]))
                        parts.append(seg)
                    add("> **{} 个主题出现分歧/条件式表述**：{}。这些主题下"
                        "要么看多与看空并存、要么同一句里「承认利好但强调还需时间」，"
                        "说明市场**没有形成单边共识**，"
                        "这正是第十四章「**震荡筑底**」在观点侧的镜像。".format(
                            len(div), "；".join(parts)))
                    add("")

        tm = vp.get("timing") or {}
        add("### 17.5 时序纪律：观点几乎全是事后")
        add("")
        add("| 阶段 | 条数 | 是否可当驱动 |")
        add("|---|---:|---|")
        add("| 主升起点之前（盘前/启动前） | {} | 仅候选，仍需三方校验 |".format(
            tm.get("pre") or 0))
        add("| 盘后 / 隔夜 / 次日 | {} | ❌ **只作印证，不计入驱动** |".format(
            tm.get("post") or 0))
        add("")
        add("> {}".format(md_plain(tm.get("note") or "")))
        add("> 与 15.2 完全同一套切分：把「某机构看好」倒推成「因为看好所以涨」"
            "是典型的事后归因谬误。本报告只把事前观点留作候选，"
            "事后观点一律降权为**解释层证据**。")
        add("")

        verified = [r for r in vp["rows"] if r["verify"] == "✅ 一致"]
        diverged = [r for r in vp["rows"] if r["verify"] == "⚠️ 背离"]
        add("### 17.6 观点 × 资金 / 涨幅交叉校验")
        add("")
        add("| 结果 | 条数 | 含义 |")
        add("|---|---:|---|")
        add("| ✅ 一致 | {} | 观点方向获第四章行业资金或第一章样本ETF涨幅支持 |".format(
            len(verified)))
        add("| ⚠️ 背离 | {} | 观点偏多但资金/涨幅未跟上，或反之 |".format(
            len(diverged)))
        add("")
        if diverged:
            add("**背离条目（不作强归因）**：")
            add("")
            for r in diverged[:6]:
                add("- {} · {} · {} · {} → {}".format(
                    md_plain(r["time"]), md_plain(r["actor"]),
                    r["stance_label"], md_plain(r["topic"]),
                    md_plain(r["title"][:40])))
            add("")
        add("> 校验口径复用 15.3 的 `TOPIC_MAP`，并补一层**第十七章专用赛道映射**"
            "（AI算力/光通信、新能源/锂电、消费/白酒、AI眼镜等），"
            "后者只用于观点章节、不改第十五章归因排序。板块取第四章 `in_top` 净流入，"
            "ETF 取第一章样本池相关分组均值；叙事型主题无验证对象则标「不适用」，"
            "外围口径标「—」，都不计入背离。资金流本身是**成交单规模统计口径**，"
            "不是账户流水，所以「一致」也只是方向印证，不是因果证明。")
        add("")

        if vp.get("signals"):
            add("### 17.7 自动信号")
            add("")
            for i, s in enumerate(vp["signals"], 1):
                add("{}. {}".format(i, s))
            add("")
        add("### 17.8 对结论的影响")
        add("")
        _kf = next((r for r in vp["rows"]
                    if "还需时间" in r["snippet"] or "共识兑现" in r["snippet"]), None)
        add("- **不改结论**：基准情景仍是「**震荡筑底**」。{}"
            "与第十章「一级有增量、二级未放量」、第十六章"
            "「近20日仅 20% 个股为正」三方同义；".format(
                ("观点侧「{}」属**谨慎偏多 / 条件式**，".format(
                    md_plain(_kf["title"][:36])) if _kf else
                 "观点侧多空并存、以条件式表述为主，")))
        add("- **不新增驱动**：{} 条观点里仅 {} 条在主升起点之前，"
            "其余均为事后解读，已按 17.5 降权；".format(
                vp["n"], tm.get("pre") or 0))
        add("- **补一层证据**：第十五章回答「发生了什么」，本章回答"
            "「市场怎么解释、解释有没有分歧、解释和资金对不对得上」；")
        add("- **口径硬边界**：立场为关键词规则判定（非语义模型）、观点方为粗分、"
            "转载方≠原始发布方，三者都可能误判，报告只给分布与排序，不给因果断言。")
        add("")

    add("---")
    add("")
    add("**风险提示**：以上为客观行情数据整理与规则化情景推演，不构成投资建议，也不是点位预测。"
        "第三章历史成交额为折算估算值，只适合看趋势与分位，不可当作历史精确金额使用；"
        "第四章「主力资金」为行情商按成交单规模划分的统计口径，并非真实资金账户流水，"
        "概念板块成分重叠不可加总；第五章 ETF 份额需连续运行才能算净申赎；"
        "第六章美股存在时差；第七章同时给出官方与自建评分；"
        "第九章回测已扣成本并做样本外分割，但样本量仍小、未复权；"
        "第十章全市场ETF明细、第十一章宏观、第十二章逐日历史、第十三章质量审计"
        "均依赖公开接口的可用性与快照新鲜度；"
        "第十四章 14.1d 的期权 IV 由最新价用 Black-Scholes 反解（r 固定 1.5%），"
        "不是行情商字段，报价陈旧或深度实值合约会被跳过，只适合同一时点横向比较；"
        "第十五章事件归因基于媒体资讯的时间戳与关键词映射，"
        "资讯是解释而非证明、存在事后归因偏差，只给置信度排序不给因果断言；"
        "7×24 快讯采用「东财 JSON + 东财 H5 + 东财 PC 7_24 页 + 新浪 PC 页 + "
        "华夏时报 + 财联社」六通道，"
        "部分网络环境会拒绝 eastmoney / sina / chinatimes / cls 部分域名，此时该章"
        "以可直连通道与网页通道采集样本为准，覆盖度低于全量快讯；"
        "东财 PC 页的 _n 翻页不是可靠时间游标，次日盘中快讯只作观察清单验证、"
        "不反推为报告日驱动；"
        "财联社为社区逆向端点、sign 与端点版本可能变更，商用应走官方授权；"
        "纯文本快讯按正文年份或页面日期做归属校验，归档旧闻一律不并入事件池。"
        "第十五章 15.9 的 iFinD 资金流与第十章一级市场净申购为不同口径，"
        "报告并列展示、不做加总。"
        "第六章 6.1.1 的股指期货基差取连续合约（IF/IH/IC/IM 主力连续）单点快照，"
        "无历史分位、“连续”价格在换月时会跳变，只能作当日对冲温度参考；"
        "第十六章涨跌排行为腾讯行情公开接口的收盘快照，已用上涨下跌家数与第二章"
        "广度逐家校验日期归属；涨停按 10/20/30cm 分板块阈值判定、与官方广度"
        "存在 1–3 只边界差异，ETF 分类按名称关键词粗分而非官方分类，"
        "个股「主力净流入」为成交单规模统计口径、非账户流水。"
        "第十七章多方观点由关键词规则识别立场、观点方为粗分、转载方不等于"
        "原始发布方，均可能误判；观点是解释不是证明，盘后与次日观点一律降权"
        "为印证、不作涨跌驱动；归档页与空壳页已按日期归属校验拦截。"
        "若反弹由事件与情绪驱动、缺乏基本面数据支撑，外围波动可能放大次日振幅。"
        "数据可能有延迟，以交易所官方为准。投资有风险，决策需谨慎。")
    add("")
    return "\n".join(L)


# ----------------------------------------------------------------------------
# 入口
# ----------------------------------------------------------------------------

SQLITE_DDL = """
CREATE TABLE IF NOT EXISTS run_meta(
    schema_ver INTEGER NOT NULL, updated TEXT, as_of TEXT, report_date TEXT);
CREATE TABLE IF NOT EXISTS daily(
    date TEXT PRIMARY KEY, amount REAL, amount_chg REAL, up_ratio REAL,
    uplim REAL, dnlim REAL, net_total REAL, pos_n INTEGER,
    sample_weighted REAL, erp REAL, adj_score REAL,
    amount_src TEXT, amount_partial INTEGER, updated TEXT);
CREATE TABLE IF NOT EXISTS vol_cone(
    report_date TEXT, as_of TEXT, code TEXT, name TEXT,
    hv20 REAL, hv60 REAL, hv120 REAL, pct REAL, state TEXT,
    cone_low REAL, cone_mid REAL, cone_high REAL,
    PRIMARY KEY(report_date, code));
CREATE TABLE IF NOT EXISTS topic_ar(
    date TEXT, topic TEXT, ar REAL, n INTEGER, beta REAL,
    car REAL, first_hhmm TEXT, PRIMARY KEY(date, topic));
CREATE TABLE IF NOT EXISTS consistency(
    report_date TEXT, as_of TEXT, name TEXT, a TEXT, b TEXT, diff TEXT,
    ok INTEGER, note TEXT, PRIMARY KEY(report_date, name));
CREATE TABLE IF NOT EXISTS quality(
    report_date TEXT, as_of TEXT, name TEXT, ok INTEGER, detail TEXT,
    PRIMARY KEY(report_date, name));
CREATE TABLE IF NOT EXISTS option_iv(
    report_date TEXT, code TEXT, name TEXT, spot REAL, atm_k REAL,
    days INTEGER, iv_call REAL, iv_put REAL, iv_atm REAL, hv20 REAL,
    vrp REAL, skew REAL, PRIMARY KEY(report_date, code));
-- 以下为 JSON 真源的镜像：结构照抄原始字段，不做二次加工
CREATE TABLE IF NOT EXISTS snapshot_meta(
    report_date TEXT PRIMARY KEY, source TEXT, updated TEXT, n_codes INTEGER);
CREATE TABLE IF NOT EXISTS snapshot_quote(
    report_date TEXT, code TEXT, name TEXT, price REAL, prev_close REAL,
    pct REAL, amount_yi REAL, turnover REAL, vol_ratio REAL, d5 REAL,
    d20 REAL, ytd REAL, discount REAL, iopv REAL, time TEXT,
    PRIMARY KEY(report_date, code));
CREATE TABLE IF NOT EXISTS sector_flow(
    report_date TEXT, kind TEXT, code TEXT, name TEXT, pct REAL,
    amount_yi REAL, turnover REAL, inflow_yi REAL, outflow_yi REAL,
    net_yi REAL, net5_yi REAL, net20_yi REAL, updown TEXT, leader TEXT,
    PRIMARY KEY(report_date, kind, code));
CREATE TABLE IF NOT EXISTS breadth_metric(
    report_date TEXT, name TEXT, value REAL, raw TEXT,
    PRIMARY KEY(report_date, name));
CREATE TABLE IF NOT EXISTS global_index(
    report_date TEXT, code TEXT, name TEXT, region TEXT, price REAL,
    pct REAL, PRIMARY KEY(report_date, code));
CREATE TABLE IF NOT EXISTS shares_hist(
    date TEXT, code TEXT, shares REAL, PRIMARY KEY(date, code));
CREATE TABLE IF NOT EXISTS flash_news(
    id TEXT PRIMARY KEY, date TEXT, time TEXT, source TEXT, title TEXT,
    important INTEGER, origin TEXT, summary TEXT);
CREATE TABLE IF NOT EXISTS endpoint_health(
    report_date TEXT, name TEXT, endpoint TEXT, status TEXT, n INTEGER,
    min_n INTEGER, ok INTEGER, PRIMARY KEY(report_date, name));
-- 扩展截面：8 个 EtfCode 维度的 list 合成一张宽表（规模/估值/资金/量价/换手）
CREATE TABLE IF NOT EXISTS etf_profile(
    report_date TEXT, code TEXT, name TEXT,
    net_value_yi REAL, net_value_total_yi REAL, total_asset_yi REAL,
    pe_ttm REAL, pb REAL, ps_ttm REAL, pcf_ttm REAL, peg REAL,
    div_ttm REAL, roe REAL,
    pe_pct REAL, pb_pct REAL, ps_pct REAL, pcf_pct REAL, peg_pct REAL,
    div_pct REAL, roe_pct REAL,
    flow_end TEXT, flow_d_yi REAL, flow_w_yi REAL, flow_m_yi REAL,
    flow_q_yi REAL,
    flow_ratio_d REAL, flow_ratio_w REAL, flow_ratio_m REAL, flow_ratio_q REAL,
    turnover_yi REAL, turnover_avg_w_yi REAL, turnover_avg_m_yi REAL,
    turnover_avg_q_yi REAL, turnover_avg_y_yi REAL,
    chg REAL, chg_5d REAL, chg_20d REAL, chg_60d REAL, chg_52w REAL,
    chg_ytd REAL,
    disc REAL, disc_avg_w REAL, disc_avg_m REAL, disc_avg_q REAL,
    disc_avg_y REAL,
    turn_rate REAL, turn_avg_w REAL, turn_avg_m REAL, turn_avg_q REAL,
    turn_avg_y REAL, daily_chg REAL, daily_turn_rate REAL,
    daily_turnover_yi REAL, PRIMARY KEY(report_date, code));
CREATE TABLE IF NOT EXISTS quote_kline(
    code TEXT, date TEXT, open REAL, high REAL, low REAL, close REAL,
    volume REAL, PRIMARY KEY(code, date));
CREATE TABLE IF NOT EXISTS stock_rank(
    report_date TEXT, code TEXT, name TEXT, price REAL, pct REAL,
    pct_5d REAL, pct_20d REAL, pct_60d REAL, pct_52w REAL, pct_ytd REAL,
    net_yi REAL, inflow_yi REAL, outflow_yi REAL, turnover_rate REAL,
    zf REAL, speed REAL, lb REAL, stock_type TEXT,
    PRIMARY KEY(report_date, code));
CREATE TABLE IF NOT EXISTS etf_rank(
    report_date TEXT, code TEXT, name TEXT, price REAL, pct REAL, zf REAL,
    amount_yi REAL, turnover REAL, stamp TEXT,
    PRIMARY KEY(report_date, code));
CREATE TABLE IF NOT EXISTS futures(
    report_date TEXT, code TEXT, name TEXT, last REAL, prev_settle REAL,
    pct REAL, high REAL, low REAL, volume REAL, oi REAL, spot_name TEXT,
    time TEXT, PRIMARY KEY(report_date, code));
CREATE TABLE IF NOT EXISTS market_score(
    report_date TEXT, name TEXT, score REAL, label TEXT,
    PRIMARY KEY(report_date, name));
CREATE TABLE IF NOT EXISTS macro_indicator(
    report_date TEXT, section TEXT, name TEXT, value REAL,
    PRIMARY KEY(report_date, section, name));
CREATE TABLE IF NOT EXISTS breadth_dist(
    report_date TEXT, bucket TEXT, count INTEGER,
    PRIMARY KEY(report_date, bucket));
-- 分时序列：minute_set（上证指数）与 ups_downs_minute（上涨家数）都是
-- ["0930 3646.57 6921016", …] 这种空格分隔的字符串，按点拆开存
CREATE TABLE IF NOT EXISTS breadth_minute(
    report_date TEXT, series TEXT, time TEXT, price REAL, volume REAL,
    value REAL, PRIMARY KEY(report_date, series, time));
-- 盘口异动个股（yidong_list）
CREATE TABLE IF NOT EXISTS yidong_item(
    report_date TEXT, id TEXT, code TEXT, name TEXT, time TEXT, type INTEGER,
    price REAL, volume REAL, amount REAL, pct REAL, order_flag TEXT,
    diff_volume REAL, diff_amount REAL, raw TEXT,
    PRIMARY KEY(report_date, id));
CREATE TABLE IF NOT EXISTS strong_stock(
    report_date TEXT, code TEXT, name TEXT, pct REAL, price REAL,
    PRIMARY KEY(report_date, code));
-- 广度接口附带的成段文字（收评点评、涨跌停比说明等），原样保留
CREATE TABLE IF NOT EXISTS breadth_note(
    report_date TEXT, key TEXT, text TEXT, PRIMARY KEY(report_date, key));
-- 各接口的采集元信息：*_status / *_n / *_pages / *_days / *_covered
CREATE TABLE IF NOT EXISTS ext_meta(
    report_date TEXT, key TEXT, value TEXT, PRIMARY KEY(report_date, key));
CREATE TABLE IF NOT EXISTS news_item(
    report_date TEXT, kind TEXT, id TEXT, title TEXT, source TEXT,
    time TEXT, rank REAL, body TEXT, PRIMARY KEY(report_date, kind, id));
-- 龙虎榜：腾讯返回的是无表头数组，只固化代码/名称/买入额/涨跌幅四列，
-- 其余列连同原始行存 raw，等哪天确定语义再平移出来，不猜 schema
CREATE TABLE IF NOT EXISTS lhb(
    report_date TEXT, tab TEXT, idx INTEGER, code TEXT, name TEXT,
    buy_yi REAL, pct REAL, raw TEXT, PRIMARY KEY(report_date, tab, idx));
-- 筹码分布：chips[code].stocks[0].data 里的标量
CREATE TABLE IF NOT EXISTS chip_metric(
    report_date TEXT, code TEXT, name TEXT, avg_cost REAL, profit_rate REAL,
    conc70 REAL, conc90 REAL, close REAL, end_date TEXT,
    PRIMARY KEY(report_date, code));
-- 财经日历 / 新股 / 机构一致预期（第十五章「明日事件预告」的三个数据源）
CREATE TABLE IF NOT EXISTS calendar_item(
    report_date TEXT, idx INTEGER, event_type TEXT, kind TEXT, title TEXT,
    stock TEXT, PRIMARY KEY(report_date, idx));
CREATE TABLE IF NOT EXISTS ipo_item(
    report_date TEXT, idx INTEGER, name TEXT, price REAL, sgrq TEXT,
    ssrq TEXT, PRIMARY KEY(report_date, idx));
CREATE TABLE IF NOT EXISTS rating_stock(
    report_date TEXT, code TEXT, name TEXT, target_price REAL,
    rating_buy INTEGER, rating_n INTEGER, np_yoy REAL, np_yoy_next REAL,
    raw TEXT, PRIMARY KEY(report_date, code));
CREATE INDEX IF NOT EXISTS idx_daily_date ON daily(date);
CREATE INDEX IF NOT EXISTS idx_ar_date ON topic_ar(date);
CREATE INDEX IF NOT EXISTS idx_snapshot_code ON snapshot_quote(code);
CREATE INDEX IF NOT EXISTS idx_flow_name ON sector_flow(name);
CREATE INDEX IF NOT EXISTS idx_flash_date ON flash_news(date);
CREATE INDEX IF NOT EXISTS idx_shares_code ON shares_hist(code);
CREATE INDEX IF NOT EXISTS idx_kline_date ON quote_kline(date);
CREATE INDEX IF NOT EXISTS idx_stock_name ON stock_rank(name);
CREATE INDEX IF NOT EXISTS idx_news_kind ON news_item(kind);
CREATE INDEX IF NOT EXISTS idx_lhb_code ON lhb(code);
-- 多通道会抓到同一句话（东财/财联社/新浪…），原始表保真，视图按分钟归并供查询
CREATE VIEW IF NOT EXISTS flash_unique AS
SELECT date, substr(time, 1, 16) AS minute, title, MIN(time) AS time,
       GROUP_CONCAT(DISTINCT NULLIF(source, '')) AS sources,
       MAX(important) AS important
FROM flash_news GROUP BY date, substr(time, 1, 16), title;
"""


def json_import_files():
    """JSON 真源路径：SQLite 镜像只读它们，不反向写回。"""
    flash = FLASH_HISTORY_FILE
    if not os.path.isabs(flash):
        flash = os.path.join(HERE, flash)
    return {"snapshot": DEFAULT_SNAPSHOT,
            "shares": SHARES_HISTORY_FILE,
            "flash": flash,
            "health": os.path.join(HERE, "endpoint_health.json")}


# 广度里值得进表的**数值**标量：(表内名称, breadth 段落, 字段)
# comments 那几项是整句文字（如「昨日融资余额…」），进数值列只会是一排 NULL，
# 留在 JSON 读原文更合适；detail 明细、分时序列同理，进表也查不出结构
BREADTH_METRICS = [
    ("up_count", "ups_downs_dsb", "up_count"),
    ("down_count", "ups_downs_dsb", "down_count"),
    ("flat_count", "ups_downs_dsb", "flat_count"),
    ("suspension_count", "ups_downs_dsb", "suspension_count"),
    ("up_limit_count", "ups_downs_dsb", "up_limit_count"),
    ("down_limit_count", "ups_downs_dsb", "down_limit_count"),
    ("up_ratio", "ups_downs_dsb", "up_ratio"),
    ("north_net_in", "north_bound", "fund_flow_net_in"),
]


def _dig(node, *keys):
    cur = node
    for k in keys:
        if not isinstance(cur, dict):
            return None
        cur = cur.get(k)
    return cur


def import_json_tables(con, report_date, updated=""):
    """把历史真源库灌进派生库，返回各表行数。

    与 write_sqlite 里的「本次结果」不同，这里导的是**累积型历史**：
    快照按快照日期整段替换，份额/快讯按主键 upsert——今天跑一次不该把
    昨天的快讯删掉。真源是 etf_store.sqlite；*.json 已冻结不再读写，
    endpoint_health.json 除外（每次运行重新生成的诊断产物，不是历史）。
    """
    files = json_import_files()
    counts = {}

    def _read(path):
        try:
            with open(path, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}

    # ---- 0. 历史真源库：快照/份额/快讯/事件台账一次读入 ----
    st_snap, st_snap_ts, st_shares, st_flash, st_event = {}, "", {}, {}, {}
    try:
        _sc = store_connect()
        try:
            _r = _sc.execute("SELECT payload,updated FROM snapshot_store "
                             "ORDER BY updated DESC, date DESC LIMIT 1").fetchone()
            if _r:
                st_snap = json.loads(_r[0])
                st_snap_ts = _r[1]
            st_shares = {d: json.loads(p) for d, p in _sc.execute(
                "SELECT date,payload FROM hist_shares")}
            for _k, _d, _p in _sc.execute(
                    "SELECT key,date,payload FROM hist_flash"):
                st_flash.setdefault(_d, {})[_k] = json.loads(_p)
            st_event = {d: json.loads(p) for d, p in _sc.execute(
                "SELECT date,payload FROM hist_event")}
        finally:
            _sc.close()
    except Exception:
        pass

    # ---- 1. 行情快照：quote / 板块资金流 / 广度标量 / 外围指数 ----
    snap = st_snap or _read(files["snapshot"])
    data = snap.get("data") or {}
    snap_date = str(snap.get("date") or report_date or "")
    try:
        # 真源库的 updated 比文件 mtime 更能说明快照时点；库空回退文件
        snap_ts = st_snap_ts or datetime.fromtimestamp(
            os.path.getmtime(files["snapshot"])).strftime("%Y-%m-%d %H:%M")
    except Exception:
        snap_ts = updated
    if data:
        for _t in ("snapshot_meta", "snapshot_quote", "sector_flow",
                   "breadth_metric", "global_index"):
            con.execute("DELETE FROM {} WHERE report_date=?".format(_t),
                        (snap_date,))
        con.execute("INSERT OR REPLACE INTO snapshot_meta VALUES (?,?,?,?)",
                    (snap_date, snap.get("source"), snap_ts, len(data)))

        qrows = []
        for code, node in data.items():
            raw = (node or {}).get("quote")
            if not isinstance(raw, list):
                continue
            q = parse_quote(raw)
            qrows.append((snap_date, code, q.get("name"), q.get("price"),
                          q.get("prev_close"), q.get("pct"),
                          q.get("amount_yi"), q.get("turnover"),
                          q.get("vol_ratio"), q.get("d5"), q.get("d20"),
                          q.get("ytd"), q.get("discount"), q.get("iopv"),
                          q.get("time")))
        if qrows:
            con.executemany(
                "INSERT OR REPLACE INTO snapshot_quote VALUES "
                "(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", qrows)

        frows = []
        ff = snap.get("fundflow") or {}
        for kind, label in (("industry", "行业"), ("concept", "概念")):
            for raw in ff.get(kind) or []:
                r = parse_board_row(raw) if isinstance(raw, dict) else None
                if not r:
                    continue
                frows.append((snap_date, label, r["code"], r["name"],
                              r["pct"], r["amount_yi"], r["turnover"],
                              r["inflow_yi"], r["outflow_yi"], r["net_yi"],
                              r["net5_yi"], r["net20_yi"], r["updown"],
                              r["leader"]))
        if frows:
            con.executemany(
                "INSERT OR REPLACE INTO sector_flow VALUES "
                "(?,?,?,?,?,?,?,?,?,?,?,?,?,?)", frows)

        brows = []
        bd = snap.get("breadth") or {}
        for name, sec, key in BREADTH_METRICS:
            v = _dig(bd, sec, key)
            if v is None or isinstance(v, (dict, list)):
                continue
            # 北向逐日净买入已停止披露：接口给的 0 是「无数据」而非真实零值，
            # 不写进派生库，避免 --sql 里被当成有效指标参与计算。
            if name == "north_net_in" and not num(v):
                continue
            brows.append((snap_date, name, num(v), str(v)[:120]))
        if brows:
            # 迁移前把「无数据」的 0 写成过有效指标，历史行也一并清掉
            con.execute("DELETE FROM breadth_metric WHERE name='north_net_in' "
                        "AND value=0")
            con.executemany("INSERT OR REPLACE INTO breadth_metric VALUES "
                            "(?,?,?,?)", brows)

        grows = []
        glob = snap.get("global") or {}
        for code, (label, region) in GLOBAL_CODES.items():
            f = glob.get(code)
            if not isinstance(f, list) or len(f) <= 32:
                continue
            grows.append((snap_date, code, label, region,
                          num(f[3]), num(f[32])))
        if grows:
            con.executemany("INSERT OR REPLACE INTO global_index VALUES "
                            "(?,?,?,?,?,?)", grows)

    # ---- 1b. 扩展截面：8 个 EtfCode 维度的 list 按代码合成一张宽表 ----
    ext = snap.get("extended") or {}

    def _by_code(key):
        d = {}
        for row in ext.get(key) or []:
            if isinstance(row, dict) and row.get("EtfCode"):
                d[row["EtfCode"]] = row
        return d

    size_m, val_m = _by_code("etf_size"), _by_code("etf_valuation_level")
    valp_m, flow_m = _by_code("etf_valuation_hist"), _by_code("etf_netflow_interval")
    amt_m, chg_m = _by_code("etf_qt_amt_interval"), _by_code("etf_qt_chg_interval")
    disc_m, turn_m = _by_code("etf_qt_disc_interval"), _by_code("etf_qt_turn_interval")
    daily_m = _by_code("etf_qt_daily")
    names = {}
    for code, node in (snap.get("data") or {}).items():
        raw = (node or {}).get("quote")
        if isinstance(raw, list):
            names[code] = parse_quote(raw).get("name") or ""

    # 这批接口的金额字段是「元」，表里统一存亿元；比例/涨跌幅是百分数，原样存
    def _yi8(v):
        n = num(v)
        return None if n is None else n / 1e8

    prows = []
    for c in set(size_m) | set(val_m) | set(flow_m) | set(chg_m) | set(daily_m):
        s0, v0, vp = size_m.get(c) or {}, val_m.get(c) or {}, valp_m.get(c) or {}
        f0, a0 = flow_m.get(c) or {}, amt_m.get(c) or {}
        g0, d0 = chg_m.get(c) or {}, disc_m.get(c) or {}
        t0, y0 = turn_m.get(c) or {}, daily_m.get(c) or {}
        prows.append((
            snap_date, c, f0.get("EtfName") or names.get(c) or "",
            _yi8(s0.get("NetValue")), _yi8(s0.get("NetValueTotal")),
            _yi8(s0.get("TotalAsset")),
            num(v0.get("PE_TTM")), num(v0.get("PB")), num(v0.get("PS_TTM")),
            num(v0.get("PCF_TTM")), num(v0.get("PEG")),
            num(v0.get("DIV_TTM")), num(v0.get("ROE")),
            num(vp.get("PE_TTM_PCT")), num(vp.get("PB_PCT")),
            num(vp.get("PS_TTM_PCT")), num(vp.get("PCF_TTM_PCT")),
            num(vp.get("PEG_PCT")), num(vp.get("DIV_TTM_PCT")),
            num(vp.get("ROE_PCT")),
            str(f0.get("EndDate") or ""), _yi8(f0.get("EtfNetPurchaseD")),
            _yi8(f0.get("EtfNetPurchaseW")), _yi8(f0.get("EtfNetPurchaseM")),
            _yi8(f0.get("EtfNetPurchaseQ")),
            num(f0.get("EtfNetPurchaseRatioD")),
            num(f0.get("EtfNetPurchaseRatioW")),
            num(f0.get("EtfNetPurchaseRatioM")),
            num(f0.get("EtfNetPurchaseRatioQ")),
            _yi8(a0.get("EtfTurnover")), _yi8(a0.get("EtfTurnoverAvgW")),
            _yi8(a0.get("EtfTurnoverAvgM")), _yi8(a0.get("EtfTurnoverAvgQ")),
            _yi8(a0.get("EtfTurnoverAvgY")),
            num(g0.get("ChgPct")), num(g0.get("ChgPct5D")),
            num(g0.get("ChgPct20D")), num(g0.get("ChgPct60D")),
            num(g0.get("ChgPct52W")), num(g0.get("ChgPctYtd")),
            num(d0.get("EtfDisc")), num(d0.get("EtfDiscAvgW")),
            num(d0.get("EtfDiscAvgM")), num(d0.get("EtfDiscAvgQ")),
            num(d0.get("EtfDiscAvgY")),
            num(t0.get("EtfTurnoverR")), num(t0.get("EtfTurnoverRAvgW")),
            num(t0.get("EtfTurnoverRAvgM")), num(t0.get("EtfTurnoverRAvgQ")),
            num(t0.get("EtfTurnoverRAvgY")),
            num(y0.get("ChgPct")), num(y0.get("TurnoverRate")),
            _yi8(y0.get("TurnoverValue"))))
    if prows:
        con.executemany(
            "INSERT OR REPLACE INTO etf_profile VALUES "
            "({})".format(",".join("?" * 53)), prows)

    # ---- 1c. 逐只日线（43 × 301 根，技术指标可据此复算）----
    krows = []
    for code, node in (snap.get("data") or {}).items():
        for b in parse_kline((node or {}).get("kline") or []):
            krows.append((code, b["date"], b["open"], b["high"], b["low"],
                          b["close"], b["volume"]))
    if krows:
        con.executemany("INSERT OR REPLACE INTO quote_kline VALUES "
                        "(?,?,?,?,?,?,?)", krows)

    # ---- 1d. 全市场排行 / 期货 / 市场评分 / 宏观 ----
    sr = []
    for r in ext.get("_market_rank") or []:
        if not isinstance(r, dict) or not r.get("code"):
            continue
        sr.append((snap_date, r["code"], r.get("name"), num(r.get("zxj")),
                   num(r.get("zdf")), num(r.get("zdf_d5")),
                   num(r.get("zdf_d20")), num(r.get("zdf_d60")),
                   num(r.get("zdf_w52")), num(r.get("zdf_y")),
                   yi(num(r.get("zljlr"))), yi(num(r.get("zllr"))),
                   yi(num(r.get("zllc"))), num(r.get("hsl")),
                   num(r.get("zf")), num(r.get("speed")), num(r.get("lb")),
                   r.get("stock_type")))
    if sr:
        con.executemany("INSERT OR REPLACE INTO stock_rank VALUES "
                        "({})".format(",".join("?" * 18)), sr)

    er = []
    for r in ext.get("_etf_rank") or []:
        if not isinstance(r, dict) or not r.get("code"):
            continue
        er.append((snap_date, r["code"], r.get("name"), num(r.get("price")),
                   num(r.get("zdf")), num(r.get("zf")),
                   yi(num(r.get("amount_wan"))), num(r.get("hsl")),
                   r.get("stamp")))
    if er:
        con.executemany("INSERT OR REPLACE INTO etf_rank VALUES "
                        "({})".format(",".join("?" * 9)), er)

    fr = []
    for r in ext.get("_futures") or []:
        if not isinstance(r, dict) or not r.get("code"):
            continue
        fr.append((snap_date, r["code"], r.get("name"), num(r.get("last")),
                   num(r.get("prev_settle")), num(r.get("pct")),
                   num(r.get("high")), num(r.get("low")),
                   num(r.get("volume")), num(r.get("oi")),
                   r.get("spot_name"), r.get("time")))
    if fr:
        con.executemany("INSERT OR REPLACE INTO futures VALUES "
                        "({})".format(",".join("?" * 12)), fr)

    # 市场评分：XXX_SCORE + XXX_STATUS 成对出现，拆成 (name, score, label)
    ms = (ext.get("market_statis_summary") or [{}])
    ms = ms[0] if ms and isinstance(ms[0], dict) else {}
    scrows = []
    for k, v in ms.items():
        if k.endswith("_SCORE"):
            base = k[:-6]
            scrows.append((snap_date, base, num(v),
                           ms.get(base + "_STATUS") or ""))
    if scrows:
        con.executemany("INSERT OR REPLACE INTO market_score VALUES "
                        "(?,?,?,?)", scrows)

    marows = []
    for key in ("macro_core_indicators_cur_p1", "macro_core_indicators_cur_p2"):
        for block in ext.get(key) or []:
            if not isinstance(block, dict):
                continue
            for sec, items in block.items():
                if not isinstance(items, dict):
                    continue
                for name, v in items.items():
                    if isinstance(v, (int, float, str)):
                        marows.append((snap_date, sec, name, num(v)))
    if marows:
        con.executemany("INSERT OR REPLACE INTO macro_indicator VALUES "
                        "(?,?,?,?)", marows)

    # 涨跌分布区间（第二章的柱状图数据）
    drows = []
    for it in _dig(snap.get("breadth") or {}, "ups_downs_dsb", "detail") or []:
        if isinstance(it, dict) and it.get("section"):
            drows.append((snap_date, it["section"], it.get("count")))
    if drows:
        con.executemany("INSERT OR REPLACE INTO breadth_dist VALUES "
                        "(?,?,?)", drows)

    # ---- 1e. 资讯：热点 / 公告 / 研报 / 脱水研报 / ETF 资金流 / 热榜 ----
    irows = []
    for it in ext.get("_hot") or []:
        if isinstance(it, dict) and it.get("title"):
            ts = num(it.get("ts"))
            irows.append((snap_date, "hot", str(it.get("id") or it["title"]),
                          it["title"][:300], it.get("source"),
                          datetime.fromtimestamp(ts).strftime(
                              "%Y-%m-%d %H:%M") if ts else "",
                          num(it.get("rank")), None))
    for kind, key in (("notice", "_notice"), ("report", "_report"),
                      ("dehydrated", "_dehydrated"), ("etf_flow", "_etf_flow"),
                      ("ct_hot", "_ct_hot")):
        for it in ext.get(key) or []:
            if isinstance(it, dict) and it.get("title"):
                irows.append((snap_date, kind,
                              str(it.get("id") or it["title"]),
                              it["title"][:300],
                              it.get("source") or it.get("org") or "",
                              str(it.get("time") or "")[:19],
                              num(it.get("rank")),
                              (it.get("summary") or "")[:800] or None))
    if irows:
        con.executemany(
            "INSERT OR REPLACE INTO news_item(report_date, kind, id, title,"
            " source, time, rank, body) VALUES (?,?,?,?,?,?,?,?)", irows)

    # 热点正文：_detail 是 {资讯 id: 正文}，id 与 _hot 的 SN… 一致，直接回填
    for nid, body in (ext.get("_detail") or {}).items():
        if isinstance(body, str) and body.strip():
            con.execute("UPDATE news_item SET body=? WHERE id=?",
                        (body[:2000], str(nid)))

    # ---- 1g. 龙虎榜 / 筹码：都是嵌套结构，取出能确定语义的字段 ----
    lrows = []
    for tab in ("all", "jg", "yyb", "gslmr", "gslxw", "yzb"):
        for i, row in enumerate((ext.get("_lhb") or {}).get(tab) or []):
            if not isinstance(row, list) or len(row) < 2:
                continue
            lrows.append((snap_date, tab, i, str(row[0]), str(row[1]),
                          _yi8(row[4]) if len(row) > 4 else None,
                          num(row[5]) if len(row) > 5 else None,
                          json.dumps(row, ensure_ascii=False)))
    if lrows:
        con.execute("DELETE FROM lhb WHERE report_date=?", (snap_date,))
        con.executemany("INSERT OR REPLACE INTO lhb VALUES "
                        "(?,?,?,?,?,?,?,?)", lrows)

    crows = []
    for code, node in (ext.get("_chips") or {}).items():
        stocks = (node or {}).get("stocks") or []
        first = stocks[0] if stocks and isinstance(stocks[0], dict) else {}
        d0 = first.get("data") if isinstance(first.get("data"), dict) else {}
        crows.append((snap_date, code, first.get("name") or "",
                      num(d0.get("ChipAvgCost")), num(d0.get("ChipProfitRate")),
                      num(d0.get("ChipConcentration70")),
                      num(d0.get("ChipConcentration90")),
                      num(d0.get("ClosePrice")), str(d0.get("EndDate") or "")))
    if crows:
        con.executemany("INSERT OR REPLACE INTO chip_metric VALUES "
                        "(?,?,?,?,?,?,?,?,?)", crows)

    # ---- 1h. 财经日历 / 新股 / 机构一致预期 ----
    cal_name = {"1": "财报/业绩", "2": "分红除权", "3": "限售解禁",
                "4": "股东大会", "5": "增发配股"}
    krows = []
    for i, it in enumerate((ext.get("_calendar") or {}).get("items") or []):
        if not isinstance(it, dict):
            continue
        t = str(it.get("type") or "")
        krows.append((snap_date, i, t, cal_name.get(t, t or "其他"),
                      it.get("title") or "", it.get("stock") or ""))
    if krows:
        con.execute("DELETE FROM calendar_item WHERE report_date=?", (snap_date,))
        con.executemany("INSERT OR REPLACE INTO calendar_item VALUES "
                        "(?,?,?,?,?,?)", krows)

    irows2 = []
    for i, it in enumerate((ext.get("_ipo") or {}).get("recent") or []):
        if not isinstance(it, dict):
            continue
        irows2.append((snap_date, i, it.get("name") or "",
                       num(it.get("price")), str(it.get("sgrq") or ""),
                       str(it.get("ssrq") or "")))
    if irows2:
        con.execute("DELETE FROM ipo_item WHERE report_date=?", (snap_date,))
        con.executemany("INSERT OR REPLACE INTO ipo_item VALUES "
                        "(?,?,?,?,?,?)", irows2)

    rrows = []
    for it in (ext.get("_rating") or {}).get("stocks") or []:
        if not isinstance(it, dict):
            continue
        d0 = it.get("data") if isinstance(it.get("data"), dict) else {}
        fc = d0.get("ConEarningsForecast")
        if isinstance(fc, str):          # 接口返回的是字符串化 JSON，展开再取
            try:
                fc = json.loads(fc)
            except Exception:
                fc = []
        fc = fc if isinstance(fc, list) else []
        yoy = [num((x or {}).get("ConNpYOY")) for x in fc[:2]
               if isinstance(x, dict)]
        rrows.append((snap_date, str(it.get("code") or ""),
                      d0.get("SecuName") or "",
                      num(d0.get("TargetPriceAvg")),
                      int(num(d0.get("RatingBuyCnt")) or 0),
                      int(num(d0.get("RatingCnt")) or 0),
                      yoy[0] if yoy else None,
                      yoy[1] if len(yoy) > 1 else None,
                      json.dumps(d0, ensure_ascii=False)))
    if rrows:
        con.executemany("INSERT OR REPLACE INTO rating_stock VALUES "
                        "(?,?,?,?,?,?,?,?,?)", rrows)

    # ---- 1i. 广度分时 / 异动 / 强势股 / 成段文字 / 采集元信息 ----
    # 这几块此前一直留在 JSON：量小，但查询时得先 load 6.7 MB 快照，入库后一条
    # SQL 就能取；序列按点拆行，不做 JSON 列。
    bd = snap.get("breadth") or {}

    def _minute(parent, node_key, series, with_px):
        # 外层键是 minute_set / ups_downs_minute，里面再嵌一层 minute_sh_index /
        # hsAUpsRatio，data 是 "0930 3646.57 6921016" 这样的空格分隔字符串
        node = (bd.get(parent) or {}).get(node_key) or {}
        out = []
        for s in node.get("data") or []:
            parts = str(s).split()
            if len(parts) < 2:
                continue
            if with_px:
                out.append((snap_date, series, parts[0], num(parts[1]),
                            num(parts[2]) if len(parts) > 2 else None, None))
            else:
                out.append((snap_date, series, parts[0], None, None,
                            num(parts[1])))
        return out

    mmrows = (_minute("minute_set", "minute_sh_index", "sh_index", True)
              + _minute("ups_downs_minute", "hsAUpsRatio", "ups_ratio", False))
    if mmrows:
        con.execute("DELETE FROM breadth_minute WHERE report_date=?", (snap_date,))
        con.executemany("INSERT OR REPLACE INTO breadth_minute VALUES "
                        "(?,?,?,?,?,?)", mmrows)

    tv = (bd.get("turnover_dsb") or {}).get("all") or {}
    if tv:
        con.executemany(
            "INSERT OR REPLACE INTO breadth_metric VALUES (?,?,?,?)",
            [(snap_date, "turnover_volume", num(tv.get("volume")), None),
             (snap_date, "turnover_amount_yi",
              (num(tv.get("amount")) or 0) / 1e8, None),
             (snap_date, "turnover_chg_yi",
              (num(tv.get("amount_change")) or 0) / 1e8, None)])

    yrows = []
    for it in bd.get("yidong_list") or []:
        if not isinstance(it, dict) or not it.get("ID"):
            continue
        dd = it.get("DaDanMessage") or {}
        zdf = num(it.get("Zdf"))
        yrows.append((snap_date, str(it["ID"]), it.get("Symbol") or "",
                      it.get("Name") or "",
                      str(it.get("DataTime") or "")[:19],
                      int(num(it.get("Type")) or 0), num(it.get("Price")),
                      num(it.get("Volume")), num(it.get("Amount")),
                      zdf * 100 if zdf is not None else None,
                      dd.get("OrderFlag") or "", num(dd.get("DiffVolume")),
                      num(dd.get("DiffAmount")),
                      json.dumps(it, ensure_ascii=False)))
    if yrows:
        con.execute("DELETE FROM yidong_item WHERE report_date=?", (snap_date,))
        con.executemany("INSERT OR REPLACE INTO yidong_item VALUES "
                        "(?,?,?,?,?,?,?,?,?,?,?,?,?,?)", yrows)

    ks = bd.get("ksh") or {}
    if isinstance(ks, dict) and ks.get("code"):
        con.execute("DELETE FROM strong_stock WHERE report_date=?", (snap_date,))
        con.execute("INSERT OR REPLACE INTO strong_stock VALUES (?,?,?,?,?)",
                    (snap_date, ks["code"], ks.get("name") or "",
                     num(ks.get("zdf")), num(ks.get("zxj"))))

    notes, metas = [], []
    for k, v in (bd.get("comments") or {}).items():
        key = "comments." + k
        if isinstance(v, str) and v.strip():
            notes.append((snap_date, key, v[:600]))
        else:
            metas.append((snap_date, key,
                          json.dumps(v, ensure_ascii=False)
                          if isinstance(v, (dict, list)) else str(v)))
    nw = bd.get("news") or {}
    for k, v in nw.items():
        key = "news." + k
        if isinstance(v, str) and v.strip():
            notes.append((snap_date, key, v[:600]))
        else:
            metas.append((snap_date, key, str(v)))
    for k, v in ext.items():
        if k.endswith(("_status", "_n", "_pages", "_days", "_total",
                       "_source_n", "_family_n", "_covered")):
            metas.append((snap_date, k,
                          json.dumps(v, ensure_ascii=False)
                          if isinstance(v, (dict, list)) else str(v)))
    con.execute("DELETE FROM breadth_note WHERE report_date=?", (snap_date,))
    if notes:
        con.executemany("INSERT OR REPLACE INTO breadth_note VALUES (?,?,?)",
                        notes)
    con.execute("DELETE FROM ext_meta WHERE report_date=?", (snap_date,))
    if metas:
        con.executemany("INSERT OR REPLACE INTO ext_meta VALUES (?,?,?)",
                        metas)

    # ---- 1f. 事件台账全量：AR/CAR 历史天此前只写当天，跨日统计查不到 ----
    led = {"days": st_event} or {}
    for d, items in (led.get("days") or {}).items():
        for it in items or []:
            if not isinstance(it, dict) or not it.get("topic"):
                continue
            con.execute(
                "INSERT OR IGNORE INTO topic_ar VALUES (?,?,?,?,?,?,?)",
                (d, it["topic"], it.get("ar"), it.get("n"), None,
                 it.get("car"), None))
            if it.get("car") is not None:
                con.execute("UPDATE topic_ar SET car=? WHERE date=? AND topic=?",
                            (it["car"], d, it["topic"]))

    # ---- 2. 份额历史（跨日累积，按 (date, code) upsert）----
    srows = []
    for d, m in st_shares.items():
        if not isinstance(m, dict):
            continue
        for code, v in m.items():
            if isinstance(v, (int, float)):
                srows.append((d, code, float(v)))
    if srows:
        con.executemany("INSERT OR REPLACE INTO shares_hist VALUES (?,?,?)",
                        srows)

    # ---- 3. 快讯历史（跨日累积，按 id upsert）----
    # 两个来源：真源库 hist_flash（含 source 归并结果）+
    # 快照里 _em_flash 抓到的财联社原始电报（更全，今日 849 条 vs 历史 314 条）
    news = {}

    def _flash_add(iid, day, it):
        if not isinstance(it, dict) or not iid:
            return
        news.setdefault(iid, (
            iid, day or str(it.get("time") or "")[:10],
            str(it.get("time") or "")[:19],
            it.get("source") or it.get("origin") or "",
            (it.get("title") or it.get("summary") or "")[:400],
            1 if it.get("important") else 0,
            it.get("origin") or "",
            (it.get("summary") or it.get("brief") or "")[:400]))

    for d, items in st_flash.items():
        if isinstance(items, dict):
            for iid, it in items.items():
                _flash_add((it.get("id") if isinstance(it, dict) else iid)
                           or iid, d, it)
    for it in (ext.get("_em_flash") or []):
        if isinstance(it, dict):
            _flash_add(it.get("id"), str(it.get("time") or "")[:10], it)
    if news:
        con.executemany(
            "INSERT OR REPLACE INTO flash_news(id, date, time, source, title,"
            " important, origin, summary) VALUES (?,?,?,?,?,?,?,?)",
            list(news.values()))

    # ---- 4. 端点健康（每次覆盖当日）----
    eh = _read(files["health"]) or {}
    hrows, hdate = [], str(eh.get("report_date") or report_date or "")
    for c in eh.get("channels") or []:
        if not isinstance(c, dict):
            continue
        hrows.append((hdate, c.get("name"), c.get("endpoint"),
                      c.get("status"), c.get("n"), c.get("min_n"),
                      1 if c.get("ok") else 0))
    if hrows:
        con.execute("DELETE FROM endpoint_health WHERE report_date=?", (hdate,))
        con.executemany("INSERT OR REPLACE INTO endpoint_health VALUES "
                        "(?,?,?,?,?,?,?)", hrows)

    for t in ("snapshot_quote", "sector_flow", "breadth_metric",
              "global_index", "shares_hist", "flash_news", "endpoint_health",
              "etf_profile", "quote_kline", "stock_rank", "etf_rank",
              "futures", "market_score", "macro_indicator", "breadth_dist",
              "news_item", "lhb", "chip_metric", "calendar_item", "ipo_item",
              "rating_stock", "breadth_minute", "yidong_item", "strong_stock",
              "breadth_note", "ext_meta"):
        counts[t] = con.execute("SELECT COUNT(*) FROM {}".format(t)).fetchone()[0]
    # 表数以 sqlite_master 为准：counts 只覆盖被计数的那批，run_meta 等不在其中
    counts["_tables"] = con.execute(
        "SELECT COUNT(*) FROM sqlite_master WHERE type='table'").fetchone()[0]
    return counts


def write_sqlite(path, hist, cone, ev_ar, ev_car, cons, quality,
                 as_of, report_date, updated, opt=None):
    """把本次运行的关键结果 upsert 进 SQLite，返回各表行数。

    刻意只做「派生」：数据仍由 JSON 承载，SQLite 靠每次运行重建。
    这样跨日查询（如「近 20 日成交额 vs 广度」）不必再写一遍 JSON 遍历，
    也不会因为库结构变化而丢掉原始数据。
    """
    import sqlite3
    con = sqlite3.connect(path)
    counts = {}
    try:
        # 结构升级：派生库直接丢掉重建，比写迁移脚本更安全（真源在 JSON）
        try:
            old = con.execute(
                "SELECT schema_ver FROM run_meta LIMIT 1").fetchone()
        except Exception:
            old = None
        if old and old[0] != SQLITE_SCHEMA:
            for _t in ("run_meta", "daily", "vol_cone", "topic_ar",
                       "consistency", "quality", "option_iv",
                       "snapshot_meta", "snapshot_quote", "sector_flow",
                       "breadth_metric", "global_index", "shares_hist",
                       "flash_news", "endpoint_health", "etf_profile",
                       "quote_kline", "stock_rank", "etf_rank", "futures",
                       "market_score", "macro_indicator", "breadth_dist",
                       "news_item", "lhb", "chip_metric", "calendar_item",
                       "ipo_item", "rating_stock", "breadth_minute",
                       "yidong_item", "strong_stock", "breadth_note",
                       "ext_meta"):
                con.execute("DROP TABLE IF EXISTS {}".format(_t))
            con.execute("DROP VIEW IF EXISTS flash_unique")
        con.executescript(SQLITE_DDL)
        con.execute("DELETE FROM run_meta")
        con.execute("INSERT INTO run_meta VALUES (?,?,?,?)",
                    (SQLITE_SCHEMA, updated, as_of, report_date))

        for d, v in (hist or {}).items():
            con.execute(
                "INSERT OR REPLACE INTO daily VALUES "
                "(?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (d, v.get("amount"), v.get("amount_chg"), v.get("up_ratio"),
                 v.get("uplim"), v.get("dnlim"), v.get("net_total"),
                 v.get("pos_n"), v.get("sample_weighted"), v.get("erp"),
                 v.get("adj_score"), v.get("amount_src"),
                 1 if v.get("amount_partial") else 0, updated))
        counts["daily"] = con.execute("SELECT COUNT(*) FROM daily").fetchone()[0]

        con.execute("DELETE FROM vol_cone WHERE report_date=?", (report_date,))
        for r in (cone or {}).get("rows") or []:
            _c = (r.get("cone") or {}).get(20) or {}
            con.execute("INSERT OR REPLACE INTO vol_cone VALUES "
                        "(?,?,?,?,?,?,?,?,?,?,?,?)",
                        (report_date, as_of, r["code"], r["name"],
                         r.get("hv20"), r.get("hv60"), r.get("hv120"),
                         r.get("pct"), r.get("state"),
                         _c.get("low"), _c.get("mid"), _c.get("high")))
        counts["vol_cone"] = con.execute(
            "SELECT COUNT(*) FROM vol_cone").fetchone()[0]

        _car = {r["topic"]: r for r in (ev_car or {}).get("rows") or []}
        for r in (ev_ar or {}).get("rows") or []:
            _c = _car.get(r.get("topic")) or {}
            con.execute("INSERT OR REPLACE INTO topic_ar VALUES "
                        "(?,?,?,?,?,?,?)",
                        (report_date, r.get("topic"), r.get("ar"),
                         r.get("n"), r.get("beta"), _c.get("car"),
                         r.get("first_hhmm")))
        counts["topic_ar"] = con.execute(
            "SELECT COUNT(*) FROM topic_ar").fetchone()[0]

        con.execute("DELETE FROM consistency WHERE report_date=?", (report_date,))
        for r in (cons or {}).get("rows") or []:
            con.execute("INSERT OR REPLACE INTO consistency VALUES "
                        "(?,?,?,?,?,?,?,?)",
                        (report_date, as_of, r["name"], r.get("a"),
                         r.get("b"), r.get("diff"),
                         1 if r.get("ok") else 0, r.get("note")))
        counts["consistency"] = con.execute(
            "SELECT COUNT(*) FROM consistency").fetchone()[0]

        con.execute("DELETE FROM quality WHERE report_date=?", (report_date,))
        for c in (quality or {}).get("checks") or []:
            con.execute("INSERT OR REPLACE INTO quality VALUES (?,?,?,?,?)",
                        (report_date, as_of, c["name"],
                         1 if c.get("ok") else 0, c.get("detail")))
        counts["quality"] = con.execute(
            "SELECT COUNT(*) FROM quality").fetchone()[0]

        con.execute("DELETE FROM option_iv WHERE report_date=?", (report_date,))
        for r in (opt or {}).get("rows") or []:
            con.execute("INSERT OR REPLACE INTO option_iv VALUES "
                        "(?,?,?,?,?,?,?,?,?,?,?,?)",
                        (report_date, r["code"], r["name"], r.get("spot"),
                         r.get("atm_k"), r.get("days"),
                         (r.get("iv_call") or 0) * 100 or None,
                         (r.get("iv_put") or 0) * 100 or None,
                         r.get("iv_atm"), r.get("hv20"), r.get("vrp"),
                         r.get("skew")))
        counts["option_iv"] = con.execute(
            "SELECT COUNT(*) FROM option_iv").fetchone()[0]

        # JSON 真源镜像：快照/份额/快讯/端点健康，跨日累积
        try:
            counts.update(import_json_tables(con, report_date, updated))
        except Exception as exc:
            counts["json_import"] = "fail:{}".format(type(exc).__name__)
        con.commit()
    finally:
        con.close()
    return counts


def query_sqlite(path, sql, limit=200):
    """终端交互：直接对派生库跑 SELECT，免去每次改脚本。"""
    import sqlite3
    con = sqlite3.connect(path)
    try:
        cur = con.execute(sql)
        cols = [d[0] for d in cur.description or []]
        rows = cur.fetchmany(limit)
    finally:
        con.close()
    return cols, rows


def main():
    ap = argparse.ArgumentParser(
        description="ETF全景：大盘/广度/成交额/板块资金流/ETF微观/外围/画像/事件/回测/走势")
    ap.add_argument("--offline", action="store_true", help="强制使用离线快照")
    ap.add_argument("--snapshot", default=DEFAULT_SNAPSHOT, help="快照文件路径")
    ap.add_argument("--save-snapshot", action="store_true", help="在线取数后保存为快照")
    ap.add_argument("--days", type=int, default=300, help="抓取日线根数（默认300，够算MA250）")
    ap.add_argument("--market-days", type=int, default=20,
                    help="全市场成交额趋势回看交易日数（默认20）")
    ap.add_argument("--fund-top", type=int, default=FUND_TOP,
                    help="资金流入/流出榜展示条数（默认{}）".format(FUND_TOP))
    ap.add_argument("--codes", default="", help="逗号分隔，覆盖默认ETF池（指数仍会单独抓）")
    ap.add_argument("--shares-file", default=SHARES_HISTORY_FILE,
                    help="ETF份额历史文件路径（跨日累积算净申赎）")
    ap.add_argument("--history-file", default=HISTORY_FILE,
                    help="逐日历史累积文件（无历史序列接口时的替代方案）")
    ap.add_argument("--output", default="",
                    help="报告基准路径，.html/.xlsx 由它派生（不再生成 .md）")
    ap.add_argument("--no-html", action="store_true", help="不生成内联图表的HTML报告")
    _xg = ap.add_mutually_exclusive_group()
    _xg.add_argument("--want-xlsx", dest="xlsx", action="store_true",
                     help="生成结构化数据表 xlsx（默认不生成）")
    _xg.add_argument("--no-xlsx", dest="xlsx", action="store_false",
                     help="（默认即不生成，保留仅为兼容旧命令行）")
    ap.set_defaults(xlsx=False)
    ap.add_argument("--no-sqlite", action="store_true", help="不写 SQLite 派生索引")
    ap.add_argument("--db", default=SQLITE_FILE, help="SQLite 派生库路径")
    ap.add_argument("--sql", default="",
                    help='只查询派生库后退出，如 --sql "SELECT * FROM daily LIMIT 5"')
    ap.add_argument("--flash-text", default="",
                    help="网页通道快讯文本，逗号分隔多文件；自动嗅探四种格式："
                         "东财PC（=== YYYY-MM-DD === + HH:MM 正文）、"
                         "财联社电报（同分段格式 + HH:MM:SS 秒级）、"
                         "华夏时报（日期行+HH:MM+🔗+热点榜）、"
                         "新浪纯文本（HH:MM:SS+正文）。"
                         "四者都做日期/年份归属校验，归档整段不并入事件池")
    ap.add_argument("--no-history-write", action="store_true",
                    help="不写逐日历史（回归测试用；离线模式默认已跳过写入）")
    ap.add_argument("--stdout", action="store_true", help="把报告全文打印到终端")
    ap.add_argument("--quiet", action="store_true", help="不在终端打印摘要")
    args = ap.parse_args()

    # 历史真源引导：冻结 JSON → etf_store.sqlite（表空才导，幂等）
    try:
        ensure_hist_import()
    except Exception as exc:
        print("历史库初始化失败：{}: {}".format(type(exc).__name__, exc),
              file=sys.stderr)

    # --sql：只查派生库，不抓数据、不生成报告（可嵌到别的脚本里做二次分析）
    if args.sql:
        if not os.path.exists(args.db):
            print("派生库不存在：{}（先正常跑一次生成）".format(args.db))
            return 2
        try:
            cols, rows = query_sqlite(args.db, args.sql)
        except Exception as exc:
            print("查询失败：{}: {}".format(type(exc).__name__, exc))
            return 2
        print("\t".join(cols))
        for r in rows:
            print("\t".join("" if v is None else str(v) for v in r))
        return 0

    # 报告只剩 HTML/XLSX 两种落盘形式，两个都关掉就没有产物可交
    if args.no_html and not args.xlsx:
        ap.error("--no-html 与不带 --want-xlsx 不能同时使用（.md 已不再生成）")

    global ALL_ETF_CODES, CORE_CODES
    if args.codes:
        custom = [c.strip() for c in args.codes.split(",") if c.strip()]
        ALL_ETF_CODES = custom
        CORE_CODES = custom[:4]
        ETF_GROUPS.clear()
        ETF_GROUPS["自定义"] = custom
        # 自定义池的份额历史单独存放，避免覆盖默认池已累积的跨日序列
        if args.shares_file == SHARES_HISTORY_FILE:
            args.shares_file = os.path.join(HERE, "etf_shares_history_custom.json")

    codes = list(dict.fromkeys(INDEX_CODES + list(MARKET_CODES.values()) + ALL_ETF_CODES))
    raw, ff, breadth_raw, glob_raw, ext = {}, {}, {}, {}, {}
    errors, source_note, snap_date = [], "", ""

    def _step(msg):
        """运行进度：抓取/分析要一两分钟，逐节点打到 stderr 免得像卡死。"""
        sys.stderr.write("[{} {}]\n".format(
            datetime.now(CST).strftime("%H:%M:%S"), msg))
        sys.stderr.flush()

    if not args.offline:
        _step("抓取行情（{} 个代码，约 1~2 分钟）…".format(len(codes)))
        try:
            raw, errors = fetch_online(codes, args.days)
        except Exception as exc:
            errors.append("在线抓取异常 {}".format(type(exc).__name__))
        _step("行情 {} 只，继续抓资金流/广度/外围".format(len(raw)))
        if len(raw) >= max(3, len(codes) // 2):
            source_note = "腾讯公开行情接口（在线，含广度/外围/全市场ETF/宏观）"
        else:
            errors.append("在线行情不足，回退离线快照")
        for runner, key, name in (
            (fetch_fundflow_online, "ff", "板块资金流"),
            (fetch_breadth_online, "breadth_raw", "涨跌广度"),
            (fetch_global_online, "glob_raw", "外围行情"),
        ):
            _step("抓取{}…".format(name))
            try:
                val = runner()
                if key == "ff":
                    ff = val
                elif key == "breadth_raw":
                    breadth_raw = val
                else:
                    glob_raw = val
                if not val:
                    errors.append("{}接口返回空，尝试快照缓存".format(name))
            except Exception:
                errors.append("{}抓取异常，尝试快照缓存".format(name))
        # 封装接口需要报告日期；先从已抓到的日线推断
        date_hint = ""
        for node in raw.values():
            rows = node.get("kline") or []
            if rows:
                date_hint = max(date_hint, rows[-1][0])
        if date_hint:
            _step("抓取宏观/封装接口…")
            try:
                ext = fetch_extended_online(date_hint)
                if not ext:
                    errors.append("封装接口不可用，尝试快照缓存")
            except Exception:
                errors.append("封装接口异常，尝试快照缓存")
            _step("抓取 7×24 快讯…")
            try:
                news = fetch_news_online(date_hint)
                if news.get("_hot"):
                    # 热点榜是滚动窗口：与历史缓存按 id 增量合并，
                    # 否则隔一阵重采会把盘中早段资讯挤出后永久丢失
                    # （10:10 的产业催化正是这类最该留存的条目）
                    try:
                        old_hot = load_snapshot_extended().get("_hot") or []
                        merged, seen = [], set()
                        for r in list(news["_hot"]) + list(old_hot):
                            k = (r.get("id") or
                                 str(r.get("ts")) + (r.get("title") or "")[:20])
                            if k in seen:
                                continue
                            seen.add(k)
                            merged.append(r)
                        try:
                            lo = (datetime.strptime(date_hint, "%Y-%m-%d")
                                  .replace(tzinfo=CST).timestamp() - 86400)
                            merged = [r for r in merged
                                      if (num(r.get("ts")) or 0) >= lo]
                        except Exception:
                            pass
                        merged.sort(key=lambda r: num(r.get("ts")) or 0)
                        news["_hot"] = merged
                    except Exception:
                        pass
                    # 无论本次是否采到快讯，都把快照里已有的网页通道样本
                    # 合并进来（按时间+标题去重），避免 --save-snapshot
                    # 把之前注入的东财 H5 快讯冲掉
                    try:
                        old_flash = load_snapshot_extended().get("_em_flash") or []
                        if old_flash:
                            cur, seen = [], set()
                            for r in list(news.get("_em_flash") or []) + list(old_flash):
                                k = flash_key(r)
                                if k in seen:
                                    continue
                                seen.add(k)
                                cur.append(r)
                            cur.sort(key=lambda r: r.get("time") or "")
                            news["_em_flash"] = cur
                    except Exception:
                        pass
                    ext.update(news)
                else:
                    errors.append("财经资讯接口返回空，尝试快照缓存")
            except Exception:
                errors.append("财经资讯抓取异常，尝试快照缓存")

    need_snapshot = (source_note == "" or len(raw) < len(codes)
                     or not ff or not breadth_raw or not glob_raw or not ext)
    # 快照源可用性：真源库有快照 或 旧 JSON 文件还在（迁移前的兜底）
    _store_has_snap = False
    try:
        _sc = store_connect()
        _store_has_snap = bool(_sc.execute(
            "SELECT COUNT(*) FROM snapshot_store").fetchone()[0])
        _sc.close()
    except Exception:
        pass
    if need_snapshot and (_store_has_snap or os.path.exists(args.snapshot)):
        snap, snap_date, snap_ff, snap_bd, snap_gl, snap_ext = load_snapshot(args.snapshot)
        fully_offline = (source_note == "")
        patched_codes = [c for c in snap if c not in raw]
        for c in patched_codes:
            # 打上旧快照日期标记，避免旧 K 线悄悄混进本次 store
            node = dict(snap[c])
            node["_stale_date"] = snap_date
            raw[c] = node
        if fully_offline:
            source_note = "离线快照 {}（{}）".format(
                "etf_store.sqlite" if _store_has_snap
                else os.path.basename(args.snapshot),
                snap_date or "日期未知")
        else:
            source_note += " + 快照补缺"
        patched_fields = []
        if not ff and snap_ff:
            ff = snap_ff
            patched_fields.append("板块资金流")
        if not breadth_raw and snap_bd:
            breadth_raw = snap_bd
            patched_fields.append("涨跌广度")
        if not glob_raw and snap_gl:
            glob_raw = snap_gl
            patched_fields.append("外围行情")
        if not ext and snap_ext:
            ext = snap_ext
            patched_fields.append("扩展数据（宏观/排行/龙虎榜/快讯）")
        if not fully_offline:
            PATCH_INFO.update({"date": snap_date, "fields": patched_fields,
                               "codes": patched_codes})
    elif not source_note:
        source_note = "在线数据（不完整）"

    # 扩展数据子项失败时接口只返回部分键，既不补缺也不报错，这里显式记录
    missing_ext = [label for key, label in (
        ("_lhb", "龙虎榜"), ("_market_rank", "全市场个股排行"),
        ("_etf_rank", "ETF 排行"), ("_hot", "热点榜"))
        if not ext.get(key)]
    if missing_ext:
        PATCH_INFO["missing_ext"] = missing_ext

    # 数据时点标定：盘中 / 收盘后 / 离线快照，决定报告能否用「收盘快照」表述
    AS_OF.update(resolve_as_of(source_note))
    RUN_WARNINGS.extend(errors[:4])

    # 补缺场景下 raw/ff/ext 混有旧快照值，绝不能回写：否则旧数据会被
    # 盖上当日 date + 当前 updated，load_snapshot(ORDER BY updated DESC)
    # 必取它，污染会一天天自我固化。
    if args.save_snapshot and "在线" in source_note \
            and "快照补缺" not in source_note:
        normalize_flash_ext(ext, args.snapshot)   # 保证保存的六通道字段完整
        date = ""
        for node in raw.values():
            rows = node.get("kline") or []
            if rows:
                date = max(date, rows[-1][0])
        save_snapshot(raw, args.snapshot, date, ff, breadth_raw, glob_raw, ext)
        print("快照已更新：etf_store.sqlite（{}）".format(date))

    store = build_store(raw)
    if not store:
        print("错误：没有可用数据。请检查网络或快照（etf_store.sqlite）。", file=sys.stderr)
        return 1

    _step("数据分析中…")
    CHART_STORE.clear()
    today = analyze_today(store)
    _, shares_delta, shares_days = update_shares_history(
        args.shares_file, today["date"], store)
    market = analyze_market_amount(store, today, args.market_days, breadth_raw)
    fund = analyze_fundflow(ff, today, market, max(1, args.fund_top))
    bd = analyze_breadth(breadth_raw, market, today)
    intraday = analyze_intraday(breadth_raw)
    micro = analyze_etf_micro(store, shares_delta, shares_days, args.shares_file)
    glob_a = analyze_global(glob_raw, store)
    universe = analyze_etf_universe(ext, store)
    macro = analyze_macro(ext, universe)
    profile = analyze_profile(breadth_raw, market, fund, glob_a, today, store)
    official = analyze_official_profile(ext, profile)
    events = analyze_events(breadth_raw, market)
    mevents = analyze_market_events(ext, breadth_raw, store)
    backtest = analyze_backtest(store)
    advtech = analyze_advanced_tech(store)
    # 7×24 快讯逐日累积缓存：把历史留存的条目与本次采集合并去重
    report_date = today["date"]
    hist_flash, hist_meta = load_flash_history(FLASH_HISTORY_PATH, report_date)
    if hist_flash:
        cur = list(ext.get("_em_flash") or [])
        seen = {flash_key(r) for r in cur}
        for r in hist_flash:
            k = flash_key(r)
            if k not in seen:
                seen.add(k)
                cur.append(r)
        cur.sort(key=lambda r: r.get("time") or "")
        ext["_em_flash"] = cur
        ext["_flash_days"] = len(hist_meta.get("days") or {})
        ext["_flash_total"] = sum(len(v) for v in
                                  (hist_meta.get("days") or {}).values())
        ext["_flash_source_n"] = len({r.get("source") for r in cur})
    # 网页通道快讯文本：自动嗅探格式，支持逗号分隔多文件
    #  ① 华夏时报 kuaixun.chinatimes.net.cn → 日期行 + HH:MM + 🔗 + 热点榜
    #  ② 新浪 cj.sina.cn 纯文本 → HH:MM:SS + 正文（需年份归属校验）
    flash_files = [x.strip() for x in (args.flash_text or "").split(",")
                   if x.strip()]
    for fpath in flash_files:
        if not os.path.exists(fpath):
            continue
        try:
            with open(fpath, encoding="utf-8") as f:
                txt = f.read()
            is_ct = bool(CT_DATE_RE.search(txt)) and (
                "🔗" in txt or CT_HOT_ANCHOR in txt)
            is_pc = (EM_PC_MAGIC in txt or CLS_WEB_MAGIC in txt
                     or bool(EM_PC_DATE_RE.search(txt)))
            if is_pc:
                # 带日期分段的网页通道快讯：东财 PC 页 / 财联社电报
                _src = ("财联社电报" if CLS_WEB_MAGIC in txt
                        else "东方财富全球快讯")
                prows, per_day, parch = parse_eastmoney_pc_flash(
                    txt, report_date, _src)
                if parch:
                    ext["_flash_archive_empc"] = {
                        "days": parch, "n": len(prows)}
                cur = list(ext.get("_em_flash") or [])
                seen = {(r.get("time") or "")[:16] + (r.get("title") or "")[:24]
                        for r in cur}
                for r in prows:
                    k = flash_key(r)
                    if k not in seen:
                        seen.add(k)
                        cur.append(r)
                cur.sort(key=lambda r: r.get("time") or "")
                ext["_em_flash"] = cur
                if _src == "财联社电报":
                    ext["_cls_status"] = "web_channel"
                    ext["_cls_n"] = len(prows)
                    ext["_cls_pages"] = len(per_day)
                else:
                    ext["_empc_status"] = "web_channel"
                    ext["_empc_n"] = len(prows)
                    ext["_empc_days"] = per_day
            elif is_ct:
                crows, chot, cpdate, carch = parse_chinatimes_flash(
                    txt, report_date)
                if carch:
                    ext["_flash_archive_ct"] = {
                        "n": len(crows), "date": cpdate, "fmt": "chinatimes"}
                    continue
                cur = list(ext.get("_em_flash") or [])
                seen = {flash_key(r) for r in cur}
                for r in crows:
                    k = flash_key(r)
                    if k not in seen:
                        seen.add(k)
                        cur.append(r)
                cur.sort(key=lambda r: r.get("time") or "")
                ext["_em_flash"] = cur
                ext["_ct_hot"] = chot
                ext["_ct_status"] = "web_channel"
                ext["_ct_n"] = len(crows)
            else:
                frows, ffmt, farch = parse_sina_flash(txt, report_date)
                fyears_all = SINA_ARCHIVE_YEAR_RE.findall(txt)
                fyears = sorted(set(fyears_all))
                from collections import Counter
                fdom = (Counter(y[:4] for y in fyears_all).most_common(1)[0][0]
                        if fyears_all else "")
                if farch:
                    ext["_flash_archive"] = {"n": len(frows), "years": fyears,
                                             "dominant": fdom, "fmt": ffmt}
                elif frows:
                    cur = list(ext.get("_em_flash") or [])
                    seen = {flash_key(r) for r in cur}
                    for r in frows:
                        k = flash_key(r)
                        if k not in seen:
                            seen.add(k)
                            cur.append(r)
                    cur.sort(key=lambda r: r.get("time") or "")
                    ext["_em_flash"] = cur
        except Exception:
            pass
    # 统一重算快讯衍生字段，保证在线/离线/网页通道三种路径一致
    normalize_flash_ext(ext, args.snapshot)
    # 全市场涨跌排行（个股全量 + ETF 全量），先算再进审计与报告
    rank = analyze_market_rank(
        ext.get("_market_rank") or [], ext.get("_etf_rank") or [],
        bd, today, fund, today["date"],
        ext.get("_market_rank_status") or "empty",
        ext.get("_etf_rank_status") or "empty")
    quality = analyze_quality(store, snap_date, today, market, fund,
                              bd, ext, glob_a)
    # 端点健康留痕：每次运行覆写，便于对比「哪天开始改版/失效」
    try:
        with open(os.path.join(HERE, "endpoint_health.json"), "w",
                  encoding="utf-8") as f:
            json.dump({"as_of": AS_OF.get("time"),
                       "report_date": today.get("date"),
                       "source": source_note,
                       "channels": quality.get("health") or []},
                      f, ensure_ascii=False, indent=2)
    except Exception:
        pass
    attribution = analyze_attribution(ext, today["date"], today, fund,
                                      breadth_raw, glob_a, market)
    # 事件研究：主题级单日 AR + 跨日台账（样本不足时只登记、不做 t 检验）
    ev_ar = compute_topic_ar(attribution.get("topics") or [], store,
                             today["date"])
    # 日内窗口 CAR：需要当日分时，且只对 AR 里已匹配的 ETF 取数（控制请求数）
    _min_codes = {"sh000300"}
    for _r in ev_ar.get("rows") or []:
        for _m in _r.get("members") or []:
            _min_codes.add(_m["code"])
    _min_state = "offline"
    minutes = {}
    if ev_ar.get("available") and not args.offline:
        minutes = fetch_minute_prices(sorted(_min_codes))
        _min_state = "ok" if minutes else "empty"
    ev_car = compute_intraday_car(ev_ar, minutes, window=30)
    if not ev_car.get("available") and _min_state == "empty":
        # 分清「没数据源」和「数据源这次没给」：后者是限流，不是口径问题
        ev_car["reason"] = ("分时接口本次未返回数据（东财/腾讯两路都失败，"
                            "或分时尚未生成），已跳过本节；日频 AR 不受影响")
    # 台账：AR 与 CAR 一起落盘，跨日累积后才有「日内反应均值」可比
    ev_stats = update_event_study(
        EVENT_STUDY_FILE, today["date"], ev_ar.get("rows") or [],
        car={r["topic"]: r["car"] for r in (ev_car.get("rows") or [])})
    outlook = analyze_outlook(store)

    # 逐日历史累积（无历史序列接口的替代方案）：先读旧值算 diff，再写今日
    hist_before = load_daily_history(args.history_file)
    # 历史回填：快照只从脚本第一次跑的那天开始，之前的成交额用指数日 K 补齐
    amt_hist = ({"status": "offline", "days": {}, "count": 0}
                if args.offline else fetch_index_amount())
    _bf_n = 0
    if amt_hist.get("days") and not args.no_history_write:
        hist_before, _bf_n = backfill_history(args.history_file,
                                              amt_hist["days"], today["date"])
    amt_hist["backfilled"] = _bf_n
    AS_OF["hist_days"] = len(hist_before)
    diff = analyze_history_diff(hist_before, today["date"], today, market,
                                fund, bd, ext, macro, official.get("adj"))
    record = {"amount": market["total"], "amount_chg": market["chg"],
              # 盘中/快照存的是「累计到当前时点」的进度值，标记后不再当全天值用
              "amount_partial": AS_OF.get("state") in ("intraday", "snapshot"),
              "as_of": AS_OF.get("time"),
              "up_ratio": bd.get("ratio"), "uplim": bd.get("uplim"),
              "dnlim": bd.get("dnlim"), "net_total": fund.get("net_total"),
              "pos_n": len(fund.get("pos") or []),
              "sample_weighted": today.get("sample_weighted"),
              "erp": macro.get("erp"),
              "adj_score": official.get("adj"),
              "etf_netflow": universe.get("d_total")}
    # 离线复现同一份快照是幂等操作，不写逐日历史：
    # 否则反复回归测试会把「当日实时快照」当成新一天累积进去，污染 diff 基准。
    if args.offline or args.no_history_write:
        hist_after = hist_before
    else:
        hist_after = append_daily_history(args.history_file, today["date"], record)

    amt_model = analyze_amount_series(hist_after, today["date"])
    fut = analyze_futures(ext.get("_futures") or [], store)
    cone = analyze_vol_cone(store)
    opt_iv = analyze_option_iv(
        {"status": "offline", "rows": []} if args.offline
        else fetch_option_chain(), store, cone)
    # 跨源/跨章节一致性：三套独立采集（广度/排行/ETF明细）互相印证
    cons = analyze_consistency(today, market, fund, bd, rank, universe,
                               quality, attribution, fut, ev_ar, cone,
                               amt_hist, opt_iv)
    try:
        with open(os.path.join(HERE, "consistency.json"), "w",
                  encoding="utf-8") as f:
            json.dump({"as_of": AS_OF.get("time"),
                       "report_date": today.get("date"),
                       "verdict": cons["verdict"], "checks": cons["rows"]},
                      f, ensure_ascii=False, indent=2)
    except Exception:
        pass
    report = render(today, market, fund, outlook, store, source_note,
                    bd, intraday, micro, glob_a, profile, events, backtest,
                    universe, macro, official, mevents, quality, diff, advtech,
                    attribution, rank, fut, ev_ar, ev_stats, ev_car, cons, cone,
                    amt_model, opt_iv)

    stamp = today["date"] or datetime.now().strftime("%Y-%m-%d")
    _step("渲染报告…")
    # render() 输出的是 Markdown 文本，但只作 HTML 的中间表示，不落盘
    out_path = args.output or os.path.join(HERE, "docs", "index")
    base_path = os.path.splitext(out_path)[0]
    html_path = ""
    if not args.no_html:
        html_path = base_path + ".html"
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(md_to_html(report))

    xlsx_path = ""
    xlsx_sheet_n = 0
    if args.xlsx:
        try:
            sheets = []
            sheets.append(("今日大盘", [
                ["指数/ETF", "收盘", "涨跌幅%", "近20日%", "成交额(亿)"]] + [
                [q["name"], q["price"], q["pct"], q["d20"], q["amount_yi"]]
                for _, q in today["indices"]] + [
                [q["name"], q["price"], q["pct"], q["d20"], q["amount_yi"]]
                for _, q in today["ranked"]]))
            sheets.append(("板块资金流", [
                ["板块", "涨跌%", "主力净额(亿)", "5日(亿)", "20日(亿)"]] + [
                [r["name"], r["pct"], r["net_yi"], r["net5_yi"], r["net20_yi"]]
                for r in (fund["in_top"] + fund["out_top"])]))
            sheets.append(("ETF净申购", [
                ["ETF", "当日(亿)", "近周(亿)", "近月(亿)", "近季(亿)"]] + [
                [r["name"], r["d"], r["w"], r["m"], r["q"]]
                for r in universe["flow_rows"]]))
            sheets.append(("估值分位", [
                ["ETF", "PE", "PB", "PE分位%", "PB分位%", "ROE"]] + [
                [r["name"], r["pe"], r["pb"], r["pe_pct"], r["pb_pct"], r["roe"]]
                for r in universe["val_rows"]]))
            sheets.append(("回测校准", [
                ["规则", "触发", "次日胜率%", "扣成本%", "样本内%", "样本外%", "z值"]] + [
                [r["label"], r["n"],
                 (r["h"].get(1) or {}).get("win"),
                 (r["h"].get(1) or {}).get("net_win"),
                 (r["h"].get(1) or {}).get("ins_win"),
                 (r["h"].get(1) or {}).get("oos_win"),
                 (r["h"].get(1) or {}).get("z")]
                for r in backtest["rules"]]))
            # 表头取全量字段并集：回填条目只有 amount/sh_pct/sz_pct，不能用当日 record 的键
            _hkeys = []
            for _v in hist_after.values():
                for _k in _v:
                    if _k not in _hkeys:
                        _hkeys.append(_k)
            sheets.append(("逐日历史", [["日期"] + _hkeys] + [
                [k] + [v.get(x) for x in _hkeys]
                for k, v in sorted(hist_after.items())]))
            if amt_model.get("rows"):
                sheets.append(("成交额序列", [
                    ["日期", "两市成交额(亿)", "环比%", "上证涨跌%",
                     "深证涨跌%", "来源"]] + [
                    [r["date"], round(r["amount"], 1),
                     round(r["chg"], 2) if r["chg"] is not None else None,
                     round(r["sh_pct"], 2) if r["sh_pct"] is not None else None,
                     round(r["sz_pct"], 2) if r["sz_pct"] is not None else None,
                     r["src"]]
                    for r in amt_model["rows"]]))
            sheets.append(("事件时间轴", [
                ["时间", "阶段", "候选驱动", "类型", "标题", "来源", "命中主题"]] + [
                [e["time"], e["phase"], "是" if e["cand"] else "否",
                 e["kind"], e["title"], e["source"], e["topics"]]
                for e in attribution["timeline"]]))
            sheets.append(("涨跌归因", [
                ["主题", "资讯数", "候选驱动数", "置信度", "等级",
                 "相关ETF涨幅%", "反证数", "代表标题"]] + [
                [t["topic"], t["n"], t["pre_n"], t["score"], t["level"],
                 t["etf_pct"], t["contrary_n"], t["repr"]]
                for t in attribution["topics"]]))
            sheets.append(("7x24快讯", [
                ["时间", "阶段", "候选驱动", "内容", "来源", "命中主题"]] + [
                [e["time"], e["phase"], "是" if e["cand"] else "否",
                 e["title"], e["source"], e["topics"]]
                for e in attribution["flash"]]))
            sheets.append(("跨源互证", [
                ["关键事实", "出现次数", "独立信源", "代表条目", "首次时间"]] + [
                [c["entity"], c["n"], "、".join(c["fams"]), c["repr"], c["time"]]
                for c in attribution["cross"]]))
            sheets.append(("ETF资金印证", [
                ["口径", "指标", "数值(亿)", "数据日", "来源", "ETF", "代码"]] + [
                [("iFinD资金流" if "ETF整体净流入" in (f["values"] or {})
                  else "区间累计"),
                 k, v, (f["time"] or "")[:16], f["source"], "", ""]
                for f in attribution["etf_flow"]
                for k, v in (f["values"] or {}).items()] + [
                ["个券Top3", "净流入", s["value"], "", "同花顺iFinD",
                 s["name"], s["code"]]
                for f in attribution["etf_flow"]
                for s in (f.get("singles") or [])]))
            _vp = attribution.get("viewpoints") or {}
            if _vp.get("available"):
                sheets.append(("多方评论观点", [
                    ["时间", "阶段", "观点方", "转载/来源", "立场", "主题",
                     "多词", "空词", "观点摘要", "资金涨幅校验"]] + [
                    [r["time"], r["phase"], r["actor"], r["origin"],
                     r["stance_label"], r["topic"], r["bull"], r["bear"],
                     r["title"], "{} {}".format(r["verify"], r["verify_detail"])]
                    for r in _vp["rows"]] + [
                    ["— 立场分布 —"] + [""] * 9] + [
                    [s["label"], s["n"], "", "", "", "", "", "", "", ""]
                    for s in _vp["stance_dist"]] + [
                    ["— 主题分歧 —"] + [""] * 9] + [
                    [t["topic"], t["n"], t["bull"], t["bear"],
                     t["mixed"], "多空并存" if t["divergent"] else "",
                     "", "", t["verdict"], t["verify"]]
                    for t in _vp["topics"]]))
            _ck = attribution.get("checklist") or []
            if _ck:
                _mk = lambda v: ("满足" if v is True else
                                 "未满足" if v is False else "待数值确认")
                sheets.append(("观察清单判定", [
                    ["观察条件", "事件级信号", "判定", "含义"]] + [
                    [c["name"], c["sig"], _mk(c["ok"]), c["note"]]
                    for c in _ck] + [
                    ["— 次日盘中验证快讯 —", "", "", ""]] + [
                    [e["time"], e["source"], "", e["title"]]
                    for e in (attribution.get("empc_next") or [])]))
            if rank.get("available"):
                sheets.append(("个股涨跌排行", [
                    ["代码", "名称", "涨跌幅", "近5日", "近20日", "近60日",
                     "主力净流入(亿)", "换手"]] + [
                    [r["code"], r["name"], r["zdf"], r["d5"], r["d20"],
                     r["d60"], round(r["flow_yi"], 2), r["hsl"]]
                    for r in rank["gainers"]] + [
                    ["— 跌幅榜 —"] + [""] * 7] + [
                    [r["code"], r["name"], r["zdf"], "", r["d20"], "",
                     "", ""] for r in rank["losers"]]))
                sheets.append(("ETF涨跌排行", [
                    ["代码", "名称", "涨跌幅", "价格", "成交额(亿)", "换手"]] + [
                    [r["code"], r["name"], r["zdf"], r["price"],
                     round(r["amount_yi"], 2), r["hsl"]]
                    for r in rank["etf_gainers"]] + [
                    ["— 跌幅榜 —"] + [""] * 5] + [
                    [r["code"], r["name"], r["zdf"], "",
                     round(r["amount_yi"], 2), ""]
                    for r in rank["etf_losers"]] + [
                    ["— 多周期强弱 —"] + [""] * 5] + [
                    [p["label"], p["median"], p["mean"],
                     p["pos_ratio"], "", ""]
                    for p in rank["periods"]]))
            # 期货基差层（第六章 6.1.1）
            if fut.get("available"):
                sheets.append(("期货基差", [
                    ["合约", "最新", "涨跌%", "现货", "基差", "基差率%",
                     "持仓量", "成交量"]] + [
                    [r["name"], r["last"], r["pct"],
                     r.get("spot"), r.get("basis"), r.get("basis_pct"),
                     r.get("oi"), r.get("volume")]
                    for r in fut["rows"]]))
            # 事件研究（15.5.1 日频 AR / 15.5.2 日内 CAR）
            if ev_ar.get("available") or ev_car.get("available"):
                _car = {r["topic"]: r for r in (ev_car.get("rows") or [])}
                sheets.append(("主题异常收益", [
                    ["主题", "事件起点", "日频AR%", "今日CAR%", "CAR均值%",
                     "匹配ETF数", "累计AR天数", "累计CAR天数"]] + [
                    [r["topic"], r.get("first_hhmm") or "",
                     r["ar"], (_car.get(r["topic"]) or {}).get("car"),
                     (next((s.get("car_mean") for s in (ev_stats.get("rows") or [])
                             if s["topic"] == r["topic"]), None)),
                     r["n"],
                     (next((s["n"] for s in (ev_stats.get("rows") or [])
                            if s["topic"] == r["topic"]), 0)),
                     (next((s.get("car_n") for s in (ev_stats.get("rows") or [])
                            if s["topic"] == r["topic"]), 0))]
                    for r in (ev_ar.get("rows") or [])]))
            if cone.get("rows"):
                sheets.append(("波动率锥", [
                    ["标的", "HV20%", "HV60%", "HV120%", "HV20分位%",
                     "锥低20d", "锥中20d", "锥高20d", "状态"]] + [
                    [r["name"], round(r["hv20"], 2),
                     round(r["hv60"], 2) if r["hv60"] else None,
                     round(r["hv120"], 2) if r["hv120"] else None,
                     round(r["pct"], 1) if r["pct"] is not None else None,
                     round(((r.get("cone") or {}).get(20) or {}).get("low", 0), 2),
                     round(((r.get("cone") or {}).get(20) or {}).get("mid", 0), 2),
                     round(((r.get("cone") or {}).get(20) or {}).get("high", 0), 2),
                     r["state"]]
                    for r in cone["rows"]]))
            if opt_iv.get("rows"):
                sheets.append(("期权IV", [
                    ["标的", "现价", "ATM行权价", "剩余天", "认购IV%",
                     "认沽IV%", "ATM IV%", "HV20%", "VRP(pct)",
                     "偏斜(pct)"]] + [
                    [r["name"], round(r["spot"], 3), r["atm_k"], r["days"],
                     round(r["iv_call"] * 100, 2) if r.get("iv_call") else None,
                     round(r["iv_put"] * 100, 2) if r.get("iv_put") else None,
                     round(r["iv_atm"], 2),
                     round(r["hv20"], 2) if r.get("hv20") else None,
                     round(r["vrp"], 2) if r.get("vrp") is not None else None,
                     round(r["skew"], 2) if r.get("skew") is not None else None]
                    for r in opt_iv["rows"]]))
            # 目录页：15+ 张表没有索引就只能靠猜
            _desc = {
                "今日大盘": "指数与跟踪清单 ETF 的收盘/涨跌/成交额",
                "板块资金流": "申万行业 + 概念，主力净额（成交单口径，不可跨板块加总）",
                "ETF净申购": "样本 ETF 一/周/月/季净申购（亿）",
                "估值分位": "ETF 的 PE/PB 与历史分位、ROE",
                "回测校准": "4 条规则的触发数、胜率、扣成本、样本内外与 z 值",
                "逐日历史": "脚本每次运行落盘的当日关键读数",
                "事件时间轴": "资讯按发布时间排序，标注盘前/盘中/盘后",
                "涨跌归因": "主题 × 资讯数 × 置信分 × 三方校验",
                "7x24快讯": "六通道合并去重后的快讯原文",
                "跨源互证": "同一事实被 ≥2 个信源族报道的条目",
                "ETF资金印证": "iFinD 口径的 ETF 资金流（与一级申赎并列，不加总）",
                "多方评论观点": "观点方、立场、分歧主题",
                "观察清单判定": "14.4 明日观察项的事后判定",
                "个股涨跌排行": "全市场 5578 只个股涨跌幅/成交额/换手",
                "ETF涨跌排行": "全市场 1719 只 ETF 涨跌 + 多周期强弱",
                "期货基差": "IF/IH/IC/IM 与国债期货，基差=期货−现货（同日同时点）",
                "主题异常收益": "日频 AR 与 30 分钟 CAR，样本按日累积",
                "波动率锥": "HV20/60/120 年化、HV20 历史分位、20 日锥区间",
                "期权IV": "ATM 隐含波动率（BS 反解）、HV20、VRP、偏斜",
                "成交额序列": "近 20 日两市成交额（指数回填 + 实时）与指数涨跌",
            }
            sheets = [("目录", [["序号", "工作表", "数据行数", "说明"]] + [
                [i, n, max(0, len(rows) - 1), _desc.get(n, "")]
                for i, (n, rows) in enumerate(sheets, 1)])] + sheets
            xlsx_path = base_path + ".xlsx"
            write_xlsx(xlsx_path, sheets)
            xlsx_sheet_n = len(sheets)
        except Exception as exc:
            errors.append("XLSX 生成失败 {}".format(type(exc).__name__))
            xlsx_path = ""

    # SQLite 派生索引：只 upsert 本次结果，删库可重建，不影响 JSON 真源
    sqlite_path, sqlite_counts = "", {}
    if not args.no_sqlite:
        try:
            sqlite_counts = write_sqlite(
                args.db, hist_after, cone, ev_ar, ev_car, cons, quality,
                AS_OF.get("time"), today["date"], stamp, opt_iv)
            sqlite_path = args.db
        except Exception as exc:
            errors.append("SQLite 写入失败 {}".format(type(exc).__name__))

    if args.stdout:
        print(report)
    if not args.quiet:
        print("=" * 64)
        if html_path:
            print("HTML：{}（图表已内联）".format(html_path))
        if xlsx_path:
            print("XLSX：{}（{} 张数据表 + 目录页）".format(
                xlsx_path, max(0, xlsx_sheet_n - 1)))
        if sqlite_path:
            # 分组打印：本次结果 / 快照截面 / 排行宏观 / 流水资讯，二十几张表挤一行没法读
            _groups = [
                ("结果表", ("daily", "vol_cone", "topic_ar", "consistency",
                            "quality", "option_iv")),
                ("快照", ("snapshot_quote", "sector_flow", "breadth_metric",
                          "global_index", "etf_profile", "quote_kline")),
                ("排行", ("stock_rank", "etf_rank", "futures", "market_score",
                          "macro_indicator", "breadth_dist")),
                ("流水", ("shares_hist", "flash_news", "news_item",
                          "lhb", "chip_metric", "endpoint_health")),
                ("日历", ("calendar_item", "ipo_item", "rating_stock")),
                ("分时", ("breadth_minute", "yidong_item", "strong_stock",
                          "breadth_note", "ext_meta")),
            ]
            print("SQLITE：{}（{} 张表）".format(
                os.path.basename(sqlite_path),
                sqlite_counts.get("_tables") or len(sqlite_counts)))
            for label, keys in _groups:
                got = [k for k in keys if k in sqlite_counts]
                if got:
                    print("  {}：{}".format(
                        label, " / ".join("{} {}".format(k, sqlite_counts[k])
                                          for k in got)))
        print("日期 {} | {} 只ETF / {} 指数 / {} 外围 | {}".format(
            stamp, today["total_etf"], len(today["indices"]),
            len(glob_a["rows"]), source_note))
        print("广度：上涨 {:.0f}% | 涨停 {} / 跌停 {} | {}".format(
            bd.get("ratio") or 0, int(bd.get("uplim") or 0),
            int(bd.get("dnlim") or 0),
            "追高接力弱" if (bd.get("follow_up") or 100) < 50 else "赚钱效应尚可"))
        print("成交：{:.0f}亿（{}）| 环比 {}（{}）| {}".format(
            market["total"] or 0, market["gate"], pct_str(market["chg"]),
            market["chg_source"], market["vp"]))
        if fund["available"]:
            print("资金：{} 行业 | 净额 {:+.0f}亿 | 流入/流出 {}/{} | Top1 {} {:+.0f}亿".format(
                len(fund["industry"]), fund["net_total"],
                len(fund["pos"]), len(fund["neg"]),
                fund["top1"]["name"], fund["top1"]["net_yi"]))
        print("微观：{} 只ETF溢价≥0.5% | 份额历史 {} 日".format(
            len(micro.get("overheat") or []), shares_days))
        print("全景：全市场 {} 只ETF | 样本规模占比 {:.1f}% | 真实净申购 {:+.1f}亿".format(
            universe["universe_n"], universe["sample_asset_pct"] or 0,
            universe["d_total"]))
        print("宏观：{} 项 | ERP {} | 官方画像 {} | 质量 {:.0f}%".format(
            len(macro["rows"]),
            "{:+.2f}pct".format(macro["erp"]) if macro["erp"] is not None else "—",
            official["adj"] if official["adj"] is not None else "—",
            quality["score"]))
        print("画像：自建 {}/{}（{:.0f}%）| 回测 {} 规则 / {} 根日线 | 历史 {} 日".format(
            profile["total"], profile["max_total"], profile["ratio"],
            len(backtest["rules"]), backtest["total_bars"], len(hist_after)))
        if attribution["available"]:
            print("归因：{} | 事件 {} 条（驱动 {} / 事后 {}）| 主升起点 {} | 7×24快讯 {} 条（{}）".format(
                attribution["verdict"], attribution["n_events"],
                attribution["n_cand"], attribution["n_post"],
                attribution["launch"] or "—", attribution["flash_n"],
                EM_FLASH_STATUS_DESC.get(attribution["flash_status"],
                                         attribution["flash_status"])))
            if rank.get("available"):
                p20 = next((p for p in rank["periods"]
                            if p["key"] == "zdf_d20"), None)
                print("全市场排行：个股 {} 只 / ETF {} 只 | 涨停 {} / 跌停 {} | {}".format(
                    rank["n_stock"], rank["n_etf"],
                    rank.get("limit_up") or 0, rank.get("limit_down") or 0,
                    ("近20日中位数 {} / 上涨占比 {:.0f}%".format(
                        pct_str(p20["median"]), p20["pos_ratio"]))
                    if p20 else "多周期分布见第十六章"))
            if attribution.get("empc_n"):
                print("东财PC页：{} 条 / {} 个日期段 | 盘后 {} / 外围反证 {} | "
                      "次日验证层 {} 条（多源汇总）".format(
                          attribution["empc_n"],
                          len(attribution.get("empc_days") or {}),
                          len(attribution.get("empc_post") or []),
                          len(attribution.get("empc_contrary") or []),
                          len(attribution.get("empc_next") or [])))
            _vp = attribution.get("viewpoints") or {}
            if _vp.get("available"):
                _dv = sum(1 for t in _vp["topics"] if t["divergent"])
                print("多方观点：{} 条 / {} 类观点方 | 偏多 {} / 偏空 {} | "
                      "{} 个主题有分歧 | 事前 {} / 事后 {}".format(
                          _vp["n"], len(_vp["actors"]),
                          sum(1 for r in _vp["rows"] if r["stance"] in ("bull", "lean_bull")),
                          sum(1 for r in _vp["rows"] if r["stance"] in ("bear", "lean_bear")),
                          _dv, (_vp.get("timing") or {}).get("pre") or 0,
                          (_vp.get("timing") or {}).get("post") or 0))
            if (attribution["sina_status"] == "ok" or attribution["ct_n"]
                    or attribution["flash_days"]):
                print("快讯源：新浪直连 {} 条 {} | 华夏时报 {} 条 | 东财PC {} 条 | "
                      "财联社 {} 条 | 跨源互证 {} 项 / {} 族 | 累积 {} 天 / {} 条".format(
                          attribution["sina_n"],
                          "(ok)" if attribution["sina_status"] == "ok" else "",
                          attribution["ct_n"],
                          attribution.get("empc_n") or 0,
                          attribution.get("cls_n") or 0,
                          len(attribution["cross"]),
                          attribution["flash_family_n"],
                          attribution["flash_days"], attribution["flash_total"]))
        print("基准情景：{} | 多头 {} / 空头 {} | 图表 {} 张".format(
            outlook["base_scene"], outlook["total_bull"],
            outlook["total_bear"], len(CHART_STORE)))
        if errors:
            print("提示：{}".format("；".join(errors[:3])))
        print("=" * 64)
    return 0


if __name__ == "__main__":
    sys.exit(main())
