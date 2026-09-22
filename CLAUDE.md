# journals — project notes for Claude

## What this is

A personal news digest. `news_rss.py` downloads recent articles from ~10 publications' RSS feeds into
`news_articles.xlsx`; `app.py` is a tiny Flask website (http://127.0.0.1:5000) with one
"Check latest news" button that runs the fetch script and shows the articles sorted into tabs.
Articles can be marked read with the `+` on each card and rated Great / Normal / Hard / Bad in the "Read" view (button in the header);
that history lives in `reading_history.json` (keyed by article link, no accounts).
Sorting is done with **plain keyword rules only** in `classify.py` — no LLM, no API, no ML. Keep it that way.

## Files

| File | Role |
|---|---|
| `news_rss.py` | Fetches feeds → `news_articles.xlsx` (one sheet per publication, columns: Section, Title, Description, Link, Published (GMT+4), Author). Author comes from `<dc:creator>` / Atom `<author>`; FT, The Economist and Le Monde feeds have none. `FEEDS` dict at top. Runs standalone. |
| `classify.py` | Loads the xlsx, dedups, assigns one **topic** per article, adds **tags** (Educational, regions such as Asia), and groups **same-story** articles across outlets. All keyword lists are at the top. `python3 classify.py` prints counts. |
| `app.py` | Flask: `GET /` page, `POST /refresh` runs `news_rss.py`, `GET /articles` returns `classify.build_report()` as JSON. Reading history: `GET /history`, `POST /history/add` (article snapshot), `POST /history/rate` (`{link, rating}`), `POST /history/remove` (`{link}`); writes go through a lock + temp-file rename. Templates are cached (debug off) — restart after editing `index.html`. |
| `templates/index.html`, `static/style.css` | The page. Vanilla JS, no build step. Tabs / search / publication chips are all client-side filters over the one `/articles` payload. |
| `news_articles.xlsx` | Generated output, overwritten on every fetch. Git-tracked but treat as data, not code. |
| `reading_history.json` | Read articles + ratings, keyed by link. Stores a snapshot of each article (title, description, publication…) because the xlsx is overwritten each fetch and `article.id` is just a row index. Data, not code. |

Run: `pip3 install -r requirements.txt` then `python3 app.py`. Python is `/usr/bin/python3` (3.9) — avoid 3.10+ syntax (`match`, `X | Y` types).

## How classification works (so you can tune it correctly)

- **Text normalization** (`normalize`): lowercase, curly quotes → straight, `a.i.` → `ai`, `u.s.` → `us`, hyphens → spaces.
- **Keyword regex** (`keyword_regex`): whole-word match with optional endings `s/es/ed/ing/er/ers/ist/al/ic/ian` — so `election` also matches `elections`, `hack` matches `hackers`. Because of this, **do not add keywords that are prefixes of unrelated common words** (`ai` is safe thanks to word boundaries; `app` won't hit `apple`; but e.g. `art` would hit `artist`, which is intended, while `bar` would hit `barred`).
- **Topic** = highest score among `TOPICS`: title hits ×2 + description hits ×1 + `SECTION_BONUS` for the feed section. Score < `MIN_TOPIC_SCORE` (2) → "Other". Order of `TOPICS` does not matter.
- **Educational** is a tag judged on the **title only** (`EDUCATIONAL_PATTERNS`: question titles, "how to / why the / what is", "explained", "guide to", …). Titles containing review/podcast/cartoon/newsletter/live/obituary/recipe are excluded (`NOT_EDUCATIONAL`). Don't move this to descriptions — that was tried and it over-triggered massively (bare "how", "explain", "understand").
- **Regions** (`REGIONS`, currently Asia) are tags: true if the title hits, or the description has ≥2 hits, or the feed section is listed (e.g. Economist "Asia"/"China"). Add a region with one more dict entry.
- **Same story**: titles → stemmed tokens minus `STOP_WORDS`; two articles from *different* publications are related if they share ≥2 tokens with at least one *rare* one (in ≤`RARE_MAX`=6 titles), or ≥3 tokens. Groups are built by non-transitive "star" clustering (seed + its related articles, largest first, no overlap). Transitive union-find was tried and chained Fed + Bank of Japan + UK retail sales into one blob via `rate/raises/interest` — don't go back to it.

## Ongoing work: improving keywords from daily spreadsheets

The user will regularly hand over new `news_articles.xlsx` files (or point at the current one after a fetch) so the keyword rules get better over time. The workflow for that:

1. Run `python3 classify.py path/to/file.xlsx` for the counts, then inspect buckets. A quick way to see what's in a bucket / what's wrongly placed:
   ```bash
   python3 -c "
   import classify as c
   r = c.build_report('news_articles.xlsx')
   for a in r['articles']:
       if a['topic'] == 'Other': print(a['publication'], '|', a['title'], '|', c.score_topics(a))"
   ```
   Same pattern for `a['educational']`, `a['regions']`, or `r['stories']`.
2. Look for: articles in **Other** that clearly belong somewhere (missing keyword), articles in the **wrong topic** (over-broad keyword — remove it or make it multi-word), Educational false positives, junk same-story groups (add the shared filler word to `STOP_WORDS`).
3. Edit the lists at the top of `classify.py`, re-run, compare counts. Keep changes small and explain which articles motivated each keyword.
4. Words already tried and **removed on purpose** because they matched too much: `state`, `eu` (hit "EU-wide"), `leadership`, `economist` (the paper's name), bare `million/billion/trillion`, bare `died`/`death` in Culture (pulled in prison-death stories; obituary phrasings `dies aged`/`has died` stay), bare `"how "`/`explain`/`understand` for Educational.
5. Reasonable target on ~500 unique articles: Other ≲ 30, Educational ≈ 30–50, each same-story group genuinely about one event.

If the user sends an xlsx with a different layout, `load_articles` expects the columns above in that order (Author optional) and one sheet per publication; adapt there, not elsewhere.

## Style

Plain, boring Python and JS; keyword lists as literal lists; comments explain *why* a rule exists. No new dependencies beyond `flask` and `openpyxl` without asking.
