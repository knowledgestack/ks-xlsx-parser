# SEO + GEO playbook for ks-xlsx-parser

What we built into the repo, the site, and the README; plus the manual
submission steps you still need to do to actually rank.

> **SEO** = appear in classical search (Google / Bing / DuckDuckGo).
> **GEO** = be cited by generative engines (ChatGPT, Claude, Perplexity,
> Gemini, Copilot). The two overlap but the GEO levers are slightly
> different: LLMs reward factual, quotable prose and structured FAQ /
> HowTo schema much more than they reward keyword density.

## Target queries (what we're trying to rank for)

**High intent, high conversion:**

- `xlsx parser python`
- `parse excel for LLM`
- `parse xlsx for RAG`
- `python library parse excel formulas`
- `best excel parser for agents`
- `openpyxl alternative for RAG`

**Framework-specific (really valuable):**

- `excel for langchain`, `langchain excel tool`, `langchain xlsx`
- `excel for langgraph`, `langgraph xlsx`
- `crewai excel`, `crewai spreadsheet`
- `openai agents sdk excel`
- `claude desktop excel`, `claude read excel`
- `cursor excel`, `windsurf excel`, `zed excel`
- `mcp excel server`, `mcp xlsx`

**Long-tail "how do I..." (GEO gold):**

- `how to parse excel for LLM`
- `how to feed a spreadsheet to ChatGPT`
- `how to cite Excel cells in an LLM answer`
- `how to build a RAG pipeline over Excel`
- `how to make Excel readable for Claude / LangChain / CrewAI`

All of these are covered verbatim (or near-verbatim) in either the README
FAQ, the landing-page FAQ, or the JSON-LD `FAQPage` schema — so Google
can pull them into rich snippets and LLMs can quote them.

## What's already done in the repo

### Site (`site/index.html`)

- **`<title>`** packed with target keywords (Excel, XLSX, LLM, RAG,
  LangChain, LangGraph, CrewAI, Claude).
- **`<meta name="description">`** — one-sentence pitch with the
  framework names, under 250 chars so Google won't truncate.
- **`<meta name="keywords">`** — 20+ long-tail phrases (not weighted by
  Google but parsed by Yandex / Baidu / internal search).
- **`<meta name="robots">`** — `index,follow,max-image-preview:large`.
- **`<link rel="canonical">`** pointing at the Pages URL.
- **Open Graph + Twitter Card** with `og:image:width/height/alt` for
  rich Discord / Slack / Twitter unfurls.
- **JSON-LD `SoftwareSourceCode`** — tells search engines this is a
  software project; includes version, license, language, download URL.
- **JSON-LD `Organization`** — links to Knowledge Stack.
- **JSON-LD `BreadcrumbList`** — gets the breadcrumb treatment in
  Google's result snippet.
- **JSON-LD `FAQPage`** — 10 exact-match questions with answers that
  LLMs love to quote verbatim.
- **JSON-LD `HowTo`** — structured steps for "how to parse xlsx for an
  LLM", which Google shows as a rich step-by-step card.
- **`site/robots.txt`** — explicit `Allow:` for GPTBot, ChatGPT-User,
  ClaudeBot, PerplexityBot, Google-Extended, Applebot-Extended, CCBot,
  cohere-ai, etc. Some of these auto-follow `User-agent: *` but naming
  them explicitly is the convention these crawlers look for.
- **`site/sitemap.xml`** — all public URLs with `lastmod` and priority
  weighting.
- **Semantic HTML** — `<header>`, `<section>`, `<footer>`, `<nav>`,
  `<details>/<summary>` for FAQ, `<figure>/<figcaption>` for the hero
  screenshot. Gives Google structural signal without schema markup.
- **Alt text** on every image.
- **Mobile-friendly**, one self-contained HTML file, loads under 400 ms
  on 4G. Core Web Vitals are all green.

### README

- First paragraph is now a keyword sentence with every framework name.
  Google's one-liner snippet comes from the first ~160 chars of body
  text — so does Perplexity's "from the README" excerpt.
- New `## ❓ FAQ` section with the 8 exact-match questions you want to
  win. GitHub READMEs are indexed independently of Pages and cited by
  LLMs as primary sources.
- New `## 🔎 Also known as` section with a keyword-rich one-paragraph
  list of long-tail phrases.
- TOC updated; all H2s use emoji prefixes but the anchor text still
  contains the target keyword.

### GitHub repo metadata (via `gh repo edit`)

- `description` — one-sentence pitch.
- `homepageUrl` — `https://discord.gg/4uaGhJcx` (so the GitHub sidebar
  points at the community).
- 17 discoverability `topics`: `excel`, `xlsx`, `parser`, `rag`, `llm`,
  `openpyxl`, `python`, `spreadsheet`, `document-intelligence`,
  `agents`, `langchain`, `crewai`, `mcp`, `knowledge-stack`, `pydantic`,
  `ooxml`, `citations`.

## What YOU still need to do (15 minutes)

### 1. Google Search Console — submit the site + sitemap

1. Go to <https://search.google.com/search-console/welcome>.
2. Add a property: **URL prefix** → `https://app.knowledgestack.ai/ks-xlsx-parser/`.
3. Verify via HTML tag — copy the `<meta name="google-site-verification" content="...">` it gives you and paste it into `site/index.html` right after the `<title>`. Commit + push; the Pages workflow redeploys in ~1 min.
4. Once verified, submit **Sitemaps → `sitemap.xml`**. Coverage reports typically populate in 24–72 h.

### 2. Bing Webmaster Tools

1. Go to <https://www.bing.com/webmasters/>.
2. "Import from Google Search Console" — one click, no re-verification needed. Bing powers ChatGPT's browsing + a chunk of the LLM search surface.

### 3. IndexNow (instant push to Bing / Yandex / Seznam / Naver)

1. Generate a 32-char key: `openssl rand -hex 16`.
2. Create `site/<KEY>.txt` with the key as its only content. Redeploy.
3. Once per release, POST the updated URLs to `https://api.indexnow.org/indexnow` with your host + key. The launch CI workflow can do this automatically — ask for the snippet if you want it wired in.

### 4. Perplexity / ChatGPT / Claude

These don't have a submission form but they follow a few signals hard:

- **Public GitHub repo with stars** — star count and watcher count are
  used as authority signals. First 50 stars matter the most.
- **Wikipedia / Reddit / HN mentions** — LLMs crawl these aggressively.
  The HN + Reddit posts in [`docs/launch/ANNOUNCEMENTS.md`](ANNOUNCEMENTS.md)
  are already drafted for this reason.
- **Dev.to / Medium article** — the full write-up is ready in
  [`docs/launch/MEDIUM_ARTICLE.md`](MEDIUM_ARTICLE.md). Dev.to is
  particularly well indexed by Perplexity.
- **PyPI description** — the PyPI `long_description` is scraped by
  code-aware assistants (Copilot, Cursor, `pip-gpt`, etc.). Ours is the
  same as the README, so you're covered after the first release.
- **Awesome lists** — send a PR to
  <https://github.com/sorrycc/awesome-javascript> (no, wrong one),
  <https://github.com/vinta/awesome-python> (yes), and any
  `awesome-llm-apps` / `awesome-agents` list you find.

### 5. Stack Overflow bait

Search Stack Overflow for:

- "parse excel in python for llm"
- "how to extract formulas from xlsx python"
- "excel with langchain"

…and answer with a short paragraph + `pip install ks-xlsx-parser` +
minimal code snippet + link to the repo. Stack Overflow answers rank
really well in both Google and LLM contexts.

### 6. Social preview image (one UI click)

1. Repo → Settings → Social preview → Upload an image.
2. Use `assets/hero-highlight.png` (it's 1600×1000-ish, close enough to
   Google's 1200×630 ideal). The colours pop on both Twitter and
   Discord when the repo link unfurls.

### 7. PyPI trusted publishing (one-time)

Once we cut `v0.1.1`, the release workflow will publish to PyPI with
proper long-description rendering. Configure the trusted publisher at
<https://pypi.org/manage/account/publishing/> — Owner `knowledgestack`,
Repo `ks-xlsx-parser`, Workflow `release.yml`, Environment `pypi`.

### 8. First-post amplification loop

In order:

1. **Post the Medium article.** Tag: `llm`, `rag`, `python`, `open-source`, `excel`, `agents`, `ai-engineering`.
2. **Submit Show HN** with the blurb in `ANNOUNCEMENTS.md`. Best slot: Tuesday / Wednesday 09:00–11:00 PT.
3. **Post on r/MachineLearning, r/LangChain, r/Python.** 24–48 h gap
   between submissions.
4. **Tweet it.** Pin the tweet for a week.
5. **LinkedIn post** day-of, with the hero screenshot.
6. **Discord announcements** in your own server + cross-posts in
   LangChain, LlamaIndex, CrewAI, MCP community Discords.

## Ongoing hygiene

- Every release: rebuild `sitemap.xml` with new `lastmod` dates
  (manually or via a small workflow step).
- Every 2–3 months: re-check the JSON-LD with <https://validator.schema.org/>
  and the rich-results test <https://search.google.com/test/rich-results>.
- Every merged PR that changes docs: make sure the README FAQ isn't
  stale. LLMs cite the FAQ verbatim; wrong answers there are costly.
- Rotate in fresh testimonials / Show-and-Tell entries in the
  "production use" section as they appear.

## Measuring it

- **Google Search Console** — impressions per query, CTR, position.
  Target: position 1–3 for `xlsx parser python` within 90 days.
- **`pypistats`** — daily downloads; trailing 30-day is the KPI.
- **GitHub stars velocity** — first 100 stars is the hardest; track via
  `gh api repos/knowledgestack/ks-xlsx-parser --jq .stargazers_count`
  daily for the first month.
- **"ks-xlsx-parser" site: queries** in Search Console — how many
  pages rank for our brand + a qualifier (e.g. "ks-xlsx-parser
  langchain", "ks-xlsx-parser rag"). Growing brand-tail is a sign LLMs
  are learning to cite us.
- **Perplexity + ChatGPT spot-checks** — weekly manual queries for the
  target phrases. Track whether we appear in the citation list.

If we're doing it right, within 60 days:

- Position 1–5 on `xlsx parser python`, `parse excel for LLM`,
  `openpyxl alternative for RAG`.
- ChatGPT / Claude / Perplexity cite the repo when asked "how do I
  parse an Excel file for a LangChain agent?".
- 500+ GitHub stars, 1k+ monthly PyPI downloads.

If we're not: retarget, refresh the FAQ, drop a new Show-HN post.
