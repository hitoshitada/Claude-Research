# Search Grounding検索クエリテンプレート
# 変数: {date_from}, {date_to}, {topic_content}
# --- は各クエリの区切り

Find the latest news articles published between {date_from} and {date_to}
about the following topic: {topic_content}

Focus on major publications from the United States, United Kingdom, Germany,
France, and the Netherlands. Find 10 distinct, real articles.

For each article found, provide:
- The exact title of the article
- The publication/source name (e.g., Reuters, Bloomberg, Nature)
- The publication date
- The country of the source
- A detailed summary covering main findings, background, and implications
  (at least 5 sentences per article)

Format as a numbered list:
### Article N: [Article Title]
- **Source**: [Publication Name]
- **Country**: [Country]
- **Date**: [Date]
- **Summary**: [Detailed summary]

Only include real articles with verifiable sources.

---

Find recent news articles published in the last 7 days
(from {date_from} to {date_to}) about: {topic_content}

Focus specifically on sources from Japan, South Korea, and Taiwan.
Include articles from publications like Nikkei, Yonhap, KAIST, CNA,
SEMI Taiwan, and other major Asian media outlets.
Find 10 distinct, real articles.

For each article, provide:
- The exact title
- The publication/source name
- The publication date
- The country of the source
- A detailed summary (at least 5 sentences)

Format as a numbered list with ### Article N: [Title] headers.
Only include real articles with verifiable sources.

---

Find recent research papers, industry reports, and in-depth technical
articles published between {date_from} and {date_to} about: {topic_content}

Focus on specialized publications, academic journals, industry analysis,
and technical media. Prioritize sources like peer-reviewed journals,
industry newsletters, company press releases, and expert analysis.
Find 10 distinct articles from reputable sources worldwide.

For each article, provide:
- The exact title
- The publication/source name
- The publication date
- The country of the source
- A detailed summary covering methodology, key findings, and implications
  (at least 5 sentences)

Format as a numbered list with ### Article N: [Title] headers.
Only include real articles with verifiable sources.
