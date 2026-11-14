# Page SEO audit — table tools

Date: 2026-11-14 (Asia/Hong_Kong); production Vite build for three new standard tool pages. Prior three CSV pages remain within TDH protection through 2026-12-05; their live and built Title, Description and H1 matched at this audit.

| Page and main keyword | Title / H1 (characters) | Description characters | Unique explanatory paragraph words | Keyword H2 |
|---|---|---:|---:|---|
| `/transpose-csv`: transpose CSV | Transpose CSV Online (20) | 150 | 1322 | Transpose CSV Online |
| `/merge-csv-files`: merge CSV files | Merge CSV Files Online (22) | 153 | 1185 | Merge CSV Files Online |
| `/excel-to-csv`: Excel to CSV | Excel to CSV Converter Online (29) | 143 | 1132 | Excel to CSV Converter Online |

The word figures are whitespace counts of the seven page-specific article paragraphs in the final content file; labels, menus, shared footer and FAQ excluded. Each page has one H1, one matching H2, original HTML input, a working tool, a reproducible example, FAQ, internal links, sitemap entry, and a relevant [RFC 4180](https://www.rfc-editor.org/rfc/rfc4180) link (verified 200 on 2026-11-07). Current link availability should be checked again if publishing is delayed.

SEO-01–08: Pass for intent distinction, metadata, article length, inputs and authority reference. SEO-09: Desktop local production preview interactions passed and Lighthouse mobile screenshots at 412 × 823 show input on first screen. SEO-10: Pass, example output and sitemap included. SEO-11: Independent utilities link to one another; transferring output is via CSV download rather than shared workspace (N/A). SEO-12: Crawlable tool links included; no new blog required. SEO-16: Existing protected TDH unchanged; new pages have no existing TDH.

Performance: Lighthouse 12.8.2, headless Chrome, simulated mobile/default throttling, local final production build, 2026-11-14. LCP: transpose 1063 ms, merge 1063 ms, Excel 1062 ms. Load CLS: 0 for each. Separately, browser Event Timing and Layout Shift observers were temporarily added to a local preview copy to observe actual small sample interactions; the original build files were restored afterwards. Observed max event duration: 24 ms (transpose), 16 ms (merge), 32 ms (Excel); CLS excluding recent-input shifts was 0 throughout each observed workflow. These are small local desktop samples, not field INP or field CLS. Field Core Web Vitals: Unknown; insufficient data for new pages.

Correctness: `npm run build`, `node scripts/test-table-transform.mjs`, and `node scripts/test-csv-operations.mjs` passed. Browser tested a single-column transpose, union of two files with reordered/missing headers, and selection of both sheets in a sample XLSX. Previous code was used for site layout; new algorithms and text were drafted using Cursor composer-2.5 and reviewed locally. No claim of competitor rankings or keyword search volume is made.

Release conclusion: required content and available lab interaction evidence completed; field metrics pending. Result depends on actual file sizes, browser and hardware. No previous first-publication dates were reset.
