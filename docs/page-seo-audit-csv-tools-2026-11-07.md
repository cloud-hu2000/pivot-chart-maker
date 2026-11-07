# Page SEO audit: CSV analysis tools

Date: 2026-11-07, Asia/Hong_Kong. Scope: three new standard tool pages, tool directory and appended shared CSS. Source: local production build. Previous audit remains in `page-seo-audit.md`.

| URL | Search intent | Title | Description | H1 / H2 | Unique article word count |
|---|---|---:|---:|---|---:|
| `/count-unique-values-csv` | Count unique CSV values and frequencies | Count Unique Values in CSV Online (33) | 147 characters | Count Unique Values in CSV Online | 934 |
| `/remove-duplicate-rows-csv` | Remove CSV duplicates by selected columns | Remove Duplicate Rows in CSV Online (35) | 148 characters | Remove Duplicate Rows in CSV Online | 878 |
| `/unpivot-csv` | Convert wide CSV to long form | Unpivot CSV Online: Wide to Long (32) | 148 characters | Unpivot CSV Online: Wide to Long | 896 |

Word counts are whitespace-separated words in the seven unique article paragraphs in `src/csv-pages-content.json`, excluding navigation, shared material, tool controls, FAQ and footer. Titles and descriptions measured from final HTML after decoding entities. All pages have one H1 and a matching visible H2, unique descriptions, a raw-HTML CSV file input, textarea and delimiter selector, example result, contextual tool links, full tool on page and sitemap link. The CSV format reference [RFC 4180](https://www.rfc-editor.org/rfc/rfc4180) returned HTTP 200. The new paths appear in `public/sitemap.xml` and tools hub.

- SEO-01 through SEO-08: Pass for the three new pages. Examples derive from the actual calculations; input and output limits are shown in the tool.
- SEO-09: Desktop tool interactions pass on local preview. Mobile first-screen visibility: Pass in Lighthouse screenshots at 412 × 823 for all three pages; the upload input is visible.
- SEO-10: Pass for example, footer and sitemap. No customer testimonials asserted.
- SEO-11: Each page performs its own task; shared workspace handoff is N/A because these independent tools do not have a shared workspace. Users export a CSV to carry data into a different tool.
- SEO-12: Pass for crawlable links. Existing data-cleaning guide remains available; no new blog post was created.
- SEO-13: Lab LCP on local production build: 1062 ms / 1059 ms / 1063 ms, respectively, Lighthouse 12.8.2 simulated mobile, default throttling, headless Chrome, 2026-11-07. Field LCP unknown.
- SEO-14: Unknown. Three workflows were functionally tested, but neither field INP nor a timed lab interaction was obtained. Lighthouse TBT 0 ms is not INP.
- SEO-15: Lighthouse load CLS 0 for all three; interaction-time CLS and field CLS unknown.
- SEO-16: Existing homepage, three core landing pages, and five existing tools were compared against live on 2026-11-07 and their Title, Meta Description and H1 matched. First publication date and freeze expiration are not independently established. The three new pages have no prior TDH to preserve.

The 10 MB / 100,000 record / 1,000 column input limit and 300,000 unpivot output-row limit apply to these new tools. CSV data is parsed locally, exact values are case sensitive, blank records skipped, and result previews show up to 12 rows. A browser download of the unpivot example produced a CSV with the expected two rows; the other two modes share the same export code but were not independently downloaded.

Build: `npm run build` passed. `node scripts/test-csv-operations.mjs` passed. Local browser: each of the three tools was loaded, configured using a small pasted sample and produced the expected result. No production deploy was attempted. Overall release audit: **conditional**. On 2026-11-07, the user explicitly requested immediate publication with timed interaction and interaction CLS evidence still missing; these remain pending.
