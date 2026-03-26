import re
import datetime
import os
import time

import feedparser
import requests
from bs4 import BeautifulSoup
from deep_translator import GoogleTranslator
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ── 字体设置 ──────────────────────────────────────────────────────
EN_FONT = "Calibri"
CN_FONT = "Microsoft YaHei"  # Windows 打开时使用微软雅黑


def set_run_font(run, size_pt, bold=False, color=None, en_font=EN_FONT, cn_font=CN_FONT):
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.name = en_font
    if color:
        run.font.color.rgb = RGBColor(*color)
    # 设置东亚字体（中文）
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), en_font)
    rFonts.set(qn("w:hAnsi"), en_font)
    rFonts.set(qn("w:eastAsia"), cn_font)
    rFonts.set(qn("w:cs"), cn_font)


def set_para_spacing(para, before_pt=0, after_pt=4, line_rule=None):
    pPr = para._p.get_or_add_pPr()
    spacing = pPr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        pPr.append(spacing)
    spacing.set(qn("w:before"), str(int(before_pt * 20)))
    spacing.set(qn("w:after"), str(int(after_pt * 20)))
    if line_rule:
        spacing.set(qn("w:lineRule"), line_rule[0])
        spacing.set(qn("w:line"), line_rule[1])


# ── 新闻来源（RSS） ────────────────────────────────────────────────
FEEDS = {
    "🌍 国际政治 / World Politics": [
        "http://feeds.bbci.co.uk/news/world/rss.xml",
        "https://feeds.reuters.com/reuters/worldNews",
        "https://rss.nytimes.com/services/xml/rss/nyt/World.xml",
    ],
    "💰 经济财经 / Economy & Finance": [
        "https://feeds.reuters.com/reuters/businessNews",
        "https://feeds.bbci.co.uk/news/business/rss.xml",
        "https://rss.nytimes.com/services/xml/rss/nyt/Business.xml",
    ],
    "🔬 科技科学 / Technology & Science": [
        "https://feeds.bbci.co.uk/news/technology/rss.xml",
        "https://feeds.bbci.co.uk/news/science_and_environment/rss.xml",
        "https://rss.nytimes.com/services/xml/rss/nyt/Technology.xml",
    ],
    "🎭 文化社会 / Culture & Society": [
        "https://feeds.bbci.co.uk/news/entertainment_and_arts/rss.xml",
        "https://feeds.reuters.com/reuters/lifestyle",
        "https://rss.nytimes.com/services/xml/rss/nyt/Arts.xml",
    ],
    "🌿 气候环境 / Climate & Environment": [
        "https://www.theguardian.com/environment/rss",
        "https://feeds.bbci.co.uk/news/science_and_environment/rss.xml",
    ],
    "⚽ 体育 / Sports": [
        "https://feeds.bbci.co.uk/sport/rss.xml",
        "https://feeds.reuters.com/reuters/sportsNews",
    ],
}

ARTICLES_PER_SECTION = 5
SUMMARY_MAX_CHARS = 800
HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; WeeklyNewsBot/1.0)"}


# ── 抓取正文 ──────────────────────────────────────────────────────
def fetch_full_text(url, fallback_summary):
    try:
        resp = requests.get(url, headers=HEADERS, timeout=8)
        soup = BeautifulSoup(resp.text, "lxml")
        # 尝试常见正文选择器
        for selector in ["article", "[class*='article-body']", "[class*='story-body']",
                         "main", "[class*='content']"]:
            container = soup.select_one(selector)
            if container:
                paras = container.find_all("p")
                text = " ".join(p.get_text(strip=True) for p in paras if len(p.get_text(strip=True)) > 40)
                if len(text) > 200:
                    return text[:SUMMARY_MAX_CHARS]
    except Exception:
        pass
    return fallback_summary[:SUMMARY_MAX_CHARS] if fallback_summary else ""


def clean_html(text):
    text = re.sub(r"<[^>]+>", "", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def fetch_articles(urls, limit):
    articles = []
    seen_titles = set()
    for url in urls:
        if len(articles) >= limit:
            break
        try:
            feed = feedparser.parse(url)
            for entry in feed.entries:
                if len(articles) >= limit:
                    break
                title = clean_html(entry.get("title", "")).strip()
                if not title or title in seen_titles:
                    continue
                seen_titles.add(title)
                rss_summary = clean_html(entry.get("summary", entry.get("description", "")))
                link = entry.get("link", "")
                # 尝试抓取正文，失败则用 RSS 摘要
                full_text = fetch_full_text(link, rss_summary) if link else rss_summary[:SUMMARY_MAX_CHARS]
                articles.append({"title": title, "summary": full_text, "link": link})
                time.sleep(0.3)  # 避免请求过快
        except Exception as e:
            print(f"  [警告] 抓取 {url} 失败: {e}")
    return articles[:limit]


# ── 翻译 ──────────────────────────────────────────────────────────
def translate_to_chinese(text):
    if not text:
        return ""
    try:
        translator = GoogleTranslator(source="auto", target="zh-CN")
        chunks, result = [], []
        # 按句子分块，避免超出 5000 字符限制
        words = text.split(". ")
        chunk = ""
        for w in words:
            if len(chunk) + len(w) < 4000:
                chunk += w + ". "
            else:
                chunks.append(chunk.strip())
                chunk = w + ". "
        if chunk:
            chunks.append(chunk.strip())
        for c in chunks:
            result.append(translator.translate(c))
            time.sleep(0.2)
        return " ".join(result)
    except Exception as e:
        print(f"  [警告] 翻译失败: {e}")
        return ""


# ── 生成 Word 文档 ────────────────────────────────────────────────
def add_horizontal_rule(doc):
    p = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "4")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "CCCCCC")
    pBdr.append(bottom)
    pPr.append(pBdr)
    set_para_spacing(p, before_pt=0, after_pt=0)


def build_document(sections_data, week_str, date_str):
    doc = Document()

    # 页面边距
    for sec in doc.sections:
        sec.top_margin = Cm(2.5)
        sec.bottom_margin = Cm(2.5)
        sec.left_margin = Cm(3)
        sec.right_margin = Cm(3)

    # ── 大标题
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_para_spacing(title_p, before_pt=0, after_pt=6)
    r = title_p.add_run("每周全球时事新闻简报")
    set_run_font(r, size_pt=24, bold=True, color=(0x1F, 0x49, 0x7D))

    # 英文副标题
    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_para_spacing(sub_p, before_pt=0, after_pt=2)
    r2 = sub_p.add_run("Weekly Global News Briefing")
    set_run_font(r2, size_pt=13, color=(0x5A, 0x5A, 0x5A))

    # 日期行
    date_p = doc.add_paragraph()
    date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_para_spacing(date_p, before_pt=0, after_pt=16)
    r3 = date_p.add_run(week_str)
    set_run_font(r3, size_pt=10, color=(0x99, 0x99, 0x99))

    add_horizontal_rule(doc)

    for section_name, articles in sections_data.items():
        # 板块标题
        sec_p = doc.add_paragraph()
        set_para_spacing(sec_p, before_pt=18, after_pt=8)
        sec_run = sec_p.add_run(section_name)
        set_run_font(sec_run, size_pt=14, bold=True, color=(0x2E, 0x74, 0xB5))

        if not articles:
            np = doc.add_paragraph()
            r = np.add_run("（本周暂无相关新闻 / No articles available this week）")
            set_run_font(r, size_pt=10, color=(0x99, 0x99, 0x99))
            continue

        for i, art in enumerate(articles, 1):
            # 英文标题
            en_p = doc.add_paragraph()
            set_para_spacing(en_p, before_pt=10, after_pt=2)
            en_run = en_p.add_run(f"{i}.  {art['en_title']}")
            set_run_font(en_run, size_pt=11, bold=True, color=(0x1A, 0x1A, 0x1A))

            # 中文标题
            cn_p = doc.add_paragraph()
            set_para_spacing(cn_p, before_pt=0, after_pt=6)
            cn_run = cn_p.add_run(f"     {art['cn_title']}")
            set_run_font(cn_run, size_pt=11, bold=True, color=(0x2E, 0x74, 0xB5))

            # 英文正文
            if art["en_summary"]:
                en_body = doc.add_paragraph()
                en_body.paragraph_format.left_indent = Cm(0.8)
                set_para_spacing(en_body, before_pt=0, after_pt=4)
                r = en_body.add_run(art["en_summary"])
                set_run_font(r, size_pt=10, color=(0x33, 0x33, 0x33))

            # 中文正文
            if art["cn_summary"]:
                cn_body = doc.add_paragraph()
                cn_body.paragraph_format.left_indent = Cm(0.8)
                set_para_spacing(cn_body, before_pt=0, after_pt=10)
                r = cn_body.add_run(art["cn_summary"])
                set_run_font(r, size_pt=10, color=(0x33, 0x33, 0x33))

        add_horizontal_rule(doc)

    # 页脚
    footer_p = doc.add_paragraph()
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_para_spacing(footer_p, before_pt=12, after_pt=0)
    r = footer_p.add_run(
        f"来源：BBC News · Reuters · The New York Times · The Guardian  ·  生成日期：{date_str}"
    )
    set_run_font(r, size_pt=8, color=(0xAA, 0xAA, 0xAA))

    return doc


# ── 主流程 ────────────────────────────────────────────────────────
def main():
    today = datetime.date.today()
    week_str = today.strftime("%Y 年第 %W 周  ·  Week %W of %Y  ·  %B %d, %Y")
    date_str = today.strftime("%Y-%m-%d")
    filename = f"WeeklyNews_{date_str}.docx"
    output_dir = "reports"
    os.makedirs(output_dir, exist_ok=True)

    print(f"📰 开始抓取新闻 ({today}) ...")
    sections_data = {}

    for section_name, urls in FEEDS.items():
        print(f"\n  ▶ {section_name}")
        raw_articles = fetch_articles(urls, ARTICLES_PER_SECTION)
        processed = []
        for art in raw_articles:
            print(f"    翻译标题: {art['title'][:60]}...")
            cn_title = translate_to_chinese(art["title"])
            cn_summary = translate_to_chinese(art["summary"]) if art["summary"] else ""
            processed.append({
                "en_title": art["title"],
                "cn_title": cn_title,
                "en_summary": art["summary"],
                "cn_summary": cn_summary,
            })
        sections_data[section_name] = processed

    print(f"\n📄 生成 Word 文档...")
    doc = build_document(sections_data, week_str, date_str)
    output_path = os.path.join(output_dir, filename)
    doc.save(output_path)
    print(f"✅ 已保存: {output_path}")


if __name__ == "__main__":
    main()
