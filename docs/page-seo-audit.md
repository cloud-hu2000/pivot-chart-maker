# Page SEO audit

Date: 2026-10-27 (Asia/Shanghai). Scope: three new core tool pages. Existing homepage Title, Description, and H1 were left unchanged.

Word count is whitespace-separated words in the page subtitle, section headings, paragraphs, example table, source note, and page FAQ, after link labels replace markdown links. Navigation, footer, buttons, and JSON-LD are excluded.

| Page | Role | Title chars | Description chars | H1 | Keyword H2 | Words | File input in raw HTML |
| --- | --- | --- | --- | --- | --- | --- | --- |
| `/online-pivot-table` | Core tool | 44 | 145 | Online Pivot Table | Build an online pivot table from the 12-row sales sample | 1552 | Yes |
| `/pivot-chart-without-excel` | Core tool | 44 | 146 | Pivot Chart Without Excel | Create a pivot chart without Excel from the traffic sample | 1713 | Yes |
| `/xlsx-to-pivot-chart` | Core tool | 46 | 144 | XLSX to Pivot Chart | Turn the sample workbook into an XLSX to pivot chart | 1735 | Yes |

Checked locally at `http://127.0.0.1:5174` on 2026-10-27:

- `/online-pivot-table` opens on the Pivot Table tab. The sales sample shows West 404,815 and grand total 1,194,082. Survey switches to 9 rows and total 396. Traffic stays on the table tab with Month grouped by month.
- `/pivot-chart-without-excel` opens on the Chart tab with the 8-row traffic sample and a line chart.
- `/xlsx-to-pivot-chart` reads `pivot-sales-sample.xlsx` from the Sales button, shows 12 rows, and keeps grand total 1,194,082. The chart stays visible.
- Homepage H1 remains Free Online Pivot Chart Maker. Its title is unchanged.
- At a 390px viewport, the three sample buttons wrap inside the page.

Authority links used in the body: Microsoft Create a PivotTable, Create a PivotChart, and File formats that are supported in Excel.

LCP, INP, and CLS: Unknown: awaiting field data. No lab trace was recorded in this pass. The browser check confirmed the tool rendered and the sample controls changed the summary.
