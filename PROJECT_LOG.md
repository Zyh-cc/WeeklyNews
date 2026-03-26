# WeeklyNews 项目日志

> 每次对话开始前，让 Claude 读取此文件以了解项目背景和进展。

---

## 项目概述

**仓库地址**：https://github.com/Zyh-cc/WeeklyNews
**本地路径**：`E:/GitHub/WeeklyNews`
**用途**：每周一 08:00（北京时间）自动抓取全球时事新闻 + 豆瓣书榜，生成中英双语 Word 文档，自动推送到 GitHub 仓库。
**运行方式**：GitHub Actions（无需本地运行，全云端自动化）
**手动触发**：`gh workflow run "Weekly News Generator" --repo Zyh-cc/WeeklyNews`
**拉取最新报告**：`cd E:/GitHub/WeeklyNews && git pull origin main`

---

## 技术栈

| 组件 | 说明 |
|------|------|
| `feedparser` | 解析 RSS 订阅源，抓取新闻条目 |
| `requests` + `beautifulsoup4` | 抓取新闻正文 / 豆瓣书籍详情 |
| `deep-translator` | 调用 Google 翻译，英文→中文 |
| `python-docx` | 生成 Word 文档，含字体/排版控制 |
| GitHub Actions | 定时调度（每周一 00:00 UTC = 08:00 北京时间） |

---

## 新闻来源

| 板块 | RSS 来源 |
|------|----------|
| 🌍 国际政治 | BBC World、Reuters World、NYT World |
| 💰 经济财经 | Reuters Business、BBC Business、NYT Business |
| 🔬 科技科学 | BBC Technology、BBC Science、NYT Technology |
| 🎭 文化社会 | BBC Entertainment、Reuters Lifestyle、NYT Arts |
| 🌿 气候环境 | The Guardian Environment、BBC Science |
| ⚽ 体育 | BBC Sport、Reuters Sports |

每板块抓取 **5 条**新闻，每条最多 **800 字**正文（优先抓原文页，失败降级为 RSS 摘要）。

---

## 书单来源

- **豆瓣每周书榜** Top 10
- 抓取地址：`https://book.douban.com/chart?subcat=G&time=week&type=rank`（多 URL 备用）
- 每本书含：书名、作者、分类标签、豆瓣评分、内容简介（最多 400 字）

---

## 文档格式规范

| 元素 | 字体 | 字号 | 颜色 |
|------|------|------|------|
| 文档大标题 | 微软雅黑 / Times New Roman | 24pt | 深蓝 #1F497D |
| 板块标题 | 同上 | 14pt | 蓝 #2E74B5 |
| 新闻英文标题 | Times New Roman | 13pt 粗体 | 近黑 #1A1A1A |
| 新闻中文标题 | 宋体 | 13pt 粗体 | 蓝 #2E74B5 |
| 新闻正文（英） | Times New Roman | 10pt | 深灰 #333333 |
| 新闻正文（中） | 宋体 | 10pt | 深灰 #333333 |
| 来源标注 | 宋体 / TNR | 8pt | 浅灰 #AAAAAA |
| 书名 | 宋体 / TNR | 14pt 粗体 | 近黑 #1A1A1A |
| 书籍元信息 | 宋体 / TNR | 9pt | 中灰 #666666 |
| 书籍简介 | 宋体 / TNR | 10pt | 深灰 #333333 |

- 页边距：上下 2.5cm，左右 3cm
- 板块间有分割线

---

## 开发历程

### 2026-03-26 · 初始搭建
- 创建本地仓库 `E:/GitHub/WeeklyNews`，连接 GitHub 远程仓库
- 编写 `scripts/generate_news.py`：RSS 抓取 + Google 翻译 + python-docx 生成
- 配置 GitHub Actions，每周一 00:00 UTC 自动运行
- 首次运行成功，生成 `reports/WeeklyNews_2026-03-26.docx`

### 2026-03-26 · 字体与内容优化
- 修复字体问题：明确设置东亚字体（`w:eastAsia`），解决 Linux 生成 / Windows 打开的字体错乱
- 弃用 `List Bullet` 样式，改用精确缩进 + 段落间距控制
- 正文抓取：通过 BeautifulSoup 抓取原文页（最多 800 字），失败降级 RSS 摘要
- 新增来源：NYT、The Guardian；新增板块：气候环境
- GitHub Actions 加入 `fonts-noto-cjk` 安装步骤

### 2026-03-26 · 书单板块上线
- 新增豆瓣每周书榜 Top 10 板块，附作者、分类、评分、简介
- 首次验证豆瓣抓取成功（未被 GitHub Actions IP 封锁）

### 2026-03-26 · 排版细化
- 英文字体改为 **Times New Roman**，中文字体改为**宋体**
- 新闻标题从 11pt 调大至 **13pt**，书名从 12pt 调大至 **14pt**
- 每条新闻末尾新增**来源标注**（灰色小字）
- 新增 `.gitignore`，避免 Word 临时文件（`~$*`）被误提交

---

## 已知问题 / 待优化项

- [ ] 豆瓣书单有时因反爬导致抓取失败，可考虑备用数据源或加代理
- [ ] 部分新闻来源（如 NYT）正文需登录，实际只能获取 RSS 摘要
- [ ] 翻译质量依赖 Google 翻译免费接口，专有名词偶有误译
- [ ] 文档暂无封面页，可考虑加入带 Logo 的封面设计
- [ ] 暂无邮件/微信推送，只能手动 `git pull` 拉取
- [ ] 书单简介来自豆瓣页面，部分书因页面结构差异导致简介为空

---

## 文件结构

```
E:/GitHub/WeeklyNews/
├── .github/
│   └── workflows/
│       └── weekly_news.yml     # GitHub Actions 定时任务
├── scripts/
│   └── generate_news.py        # 主脚本：抓取+翻译+生成Word
├── reports/
│   └── WeeklyNews_YYYY-MM-DD.docx  # 每周自动生成的报告
├── requirements.txt            # Python 依赖
├── .gitignore
├── README.md
└── PROJECT_LOG.md              # 本文件
```
