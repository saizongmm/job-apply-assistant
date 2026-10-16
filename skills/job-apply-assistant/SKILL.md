---
name: job-apply-assistant
description: >-
  通用求职投递助手。当用户提供自己的简历(CV)文件 + 目标岗位关键词(如 Business
  Development、Developer、Marketing、PR、Sales 等) + 国家/地区范围(如 中国大陆、中国香港、新加坡),
  自动完成:采集近期在招岗位(LinkedIn / jobsDB / Indeed)、生成带去重的总表(tracker)、
  按匹配度筛选 Top 岗位、基于用户真实 CV 用 AI 定制每份简历、并用用户本机已登录的浏览器
  逐批打开岗位便于手动投递。触发词:帮我找工作、投简历、job apply、根据我的简历生成、
  岗位采集、定制简历、自动打开招聘链接、一键打开投递。
version: 1.0.0
---

# 通用求职投递助手 (Job Apply Assistant)

把"为 Selina 做的 Marketing/PR 投递管线"变成一个**任何人都能用**的通用工具。
核心思路:用户上传自己的 CV → 指定岗位 + 国家 → 系统复用同一套采集/去重/定制/打开逻辑。

## 输入(用户必须提供)
1. **CV 文件**:`.docx` 或 `.txt`(PDF 尽力支持)。这是用户的真实底稿,**不要**用模板替换。
2. **目标岗位关键词**:如 `Business Development`、`Developer`、`Data Scientist`、`Marketing Manager`。
   可以多个(用英文逗号分隔)。
3. **国家/地区范围**:如 `CN`(中国大陆)、`HK`(中国香港)、`SG`(新加坡)、`UK`、`US`。
4. 可选:全职/兼职、时间窗口(默认近 7 天)、是否要生成 cover letter(默认不生成)。

## 工作流程(7 步)
> 每一步都用 `scripts/` 下的脚本,工作目录建议新建一个独立文件夹(例如 `~/apply_<岗位>_2026`),
> 把 CV 放进去,脚本默认读取同目录的 `application_tracker.xlsx` 与 `tailored_resumes/`。

### 步骤 0 — 准备
- 新建工作目录,把用户的 CV 复制进去(例如 `cv.docx`)。
- 确认已装依赖:`pip install openpyxl python-docx`(PDF 可选 `pip install pdfminer.six`)。

### 步骤 1 — 解析 CV
```bash
python3 scripts/parse_cv.py --in cv.docx --out profile.json
```
生成 `profile.json`(姓名/邮箱/电话/经历/教育 + 原文 `raw_text`)。**始终保留 `raw_text`**,
后续定制以它为准。如果解析不到某项(比如姓名),向用户确认。

### 步骤 2 — 采集岗位(用户授权下用 WebFetch)
对每个 `(岗位关键词, 国家)` 组合,用 WebFetch 抓 LinkedIn 搜索结果。
**关键:无登录采集会撞 1000+ 懒加载墙**,破解方法是给关键词加**行业限定词**把结果压到 <1000:
- URL 模板:`https://www.linkedin.com/jobs/search/?keywords=<KW>&location=<LOC>&f_TPR=r604800&sortBy=DD`
  (`f_TPR=r604800` = 近 7 天;`location` 用 `China` / `Hong Kong` / `Singapore` 等)。
- 紧凑提取 prompt(让模型只返回 TSV,5 列):
  `Output ONLY a tab-separated table, no headers, no markdown: JOB_TITLE <tab> COMPANY <tab> LOCATION <tab> DATE_POSTED <tab> URL. Use the full linkedin.com/jobs/view/NNNNN URL.`
- 若返回 "1,000+" 或单 JD 而非卡片列表 → 给 KW 加行业词(如 `Business Development SaaS` /
  `Developer Fintech` / `Marketing Manager Technology`)再抓。HK/SG 可加 jobsDB、Indeed。
- 把每次结果**追加**保存到 `harvest.tsv`(去重交给步骤 3)。

### 步骤 3 — 建/追加 tracker(去重)
```bash
python3 scripts/build_tracker.py --tsv harvest.tsv --tracker application_tracker.xlsx [--create]
```
- 自动备份原 tracker;按 LinkedIn job-id 去重(已存在 + 批内)。
- 用 `norm_city` 归一化城市、`infer_role` 推断级别(Director/Comms/Manager/IC);追加 `LIVE` 行。
- 表结构(12 列):`# | City | Company | Job Title | Seniority | Status | LinkedIn URL | Tailored Resume | Cover Letter | Applied? | Reply | Notes`。

### 步骤 4 — 评分 + 选 Top N
按匹配度打分(参考 `scripts/build_tracker.py` 里的 `score_row`):
岗位关键词命中 +2、城市命中(用户指定的国家)+1、级别(Director/IC 按用户目标)+、行业相关 +1;
**排除招聘中介**(Robert Walters / Michael Page / Hays 等)、**排除用户明确不要的行业**。
选出 Top 15~30(用户决定数量)。

### 步骤 5 — AI 定制简历(读 CV + JD)
对选中的每个岗位:
1. 若 JD 详情不足,用 WebFetch 抓取该职位页补全(职责/要求)。
2. **由你(LLM)基于 `profile.json` 的 `raw_text` + JD 生成定制内容**,写进 `tailored.json`:
   - `summary`:3~4 条针对该 JD 重写的专业摘要(突出最相关的经历/技能/成就)。
   - `experience`(可选):从用户经历里挑选/改写最相关的几条,bullets 已按 JD 重排。
   - `target`:`{title, company, location}`。
3. 生成 docx:
```bash
python3 scripts/make_docx.py --profile profile.json --tailored tailored.json --out tailored_resumes/<文件名>.docx
```
4. 把文件名写回 tracker 的 `Tailored Resume` 列(只写这一列 + Notes,**不要**动其它列)。
   **不生成 cover letter**(除非用户要求)。

### 步骤 6 — 用用户浏览器逐批打开(Path A,手动投递)
```bash
python3 scripts/open_jobs.py --auto-open --batch 15 --delay 2
```
- 用 `open -a "Google Chrome"` 打开链接 —— **用的是用户本机已登录的浏览器/真实 IP**,
  无 cookie 跨路径问题、无反爬限流。
- 脚本只打开"有定制简历 + 未投/未决定"的岗位;打开后标 `Opened`,重跑不重复。
- 用户在浏览器里手动点 "Apply" 并上传对应简历(清单里每条都附了简历链接)。
- 投完哪些,用户报 `#` 号,运行 `python3 scripts/open_jobs.py --done <#列表>` 标 `Applied?=Yes`;
  不投的 `--skip <#列表>` 标 `No`。

### 步骤 7 — 收尾
- 更新工作记录;提示剩余可投岗位数;可继续 `--auto-open` 开下一批。
- 把 `application_tracker.xlsx` 作为交付物呈现。

## 重要诚实声明(必须告诉用户)
- **本工具不做"无人值守自动提交"**。LinkedIn 强力反爬 + macOS 钥匙串 cookie 跨路径不可解密,
  Playwright 复用不了用户的登录态;且自动提交风险高、易封号。Path A 的"在你自己浏览器里点一下"
  是可靠且安全的做法。
- **无登录大规模采集会撞墙/限流**。小规模(每岗每国几十~上百条/周)用 WebFetch 可行;
  要规模化需用户自己的 LinkedIn 会话或付费数据源。
- 定制简历是**基于用户真实 CV 的改写/突出**,不是编造经历。务必忠实于用户提供的底稿。

## 脚本参考
- `parse_cv.py` — CV → 结构化 `profile.json`(docx/txt,PDF 尽力)。
- `build_tracker.py` — harvest.tsv → 去重追加到 tracker(可 `--create` 新建)。
- `make_docx.py` — `profile.json` + `tailored.json` → 定制 `.docx`(通用版,无硬编码姓名/经历)。
- `open_jobs.py` — Path A 打开/标记(通用,读任意 tracker)。

## 与原 Selina 管线的区别
| 项 | 原管线 | 通用 Skill |
|---|---|---|
| 底稿 CV | 写死在代码常量 | 用户上传文件 → `parse_cv.py` |
| 简历改写 | 只认 marketing 话术 | LLM 读 CV+JD 定制 |
| 岗位/国家 | 写死 China + Marketing | 用户参数:岗位关键词 + 国家 |
| 采集 | agent 手动 WebFetch | 同,但参数化 + 墙破解说明 |
| 打开/去重/tracker | 通用 | 通用(直接复用) |
