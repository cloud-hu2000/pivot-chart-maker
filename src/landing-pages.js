const MICROSOFT_PIVOT_TABLE = "https://support.microsoft.com/en-us/excel/get-started/create-a-pivottable-to-analyze-worksheet-data";
const MICROSOFT_PIVOT_CHART = "https://support.microsoft.com/en-us/excel/get-started/create-a-pivotchart";
const MICROSOFT_FILE_FORMATS = "https://support.microsoft.com/en-us/excel/file-formats-that-are-supported-in-excel";

export const SAMPLES = {
  sales: {
    id: "sales",
    button: "Sales, 12 rows",
    fileName: "pivot-sales-sample.csv",
    xlsxName: "pivot-sales-sample.xlsx",
    xlsxUrl: "/samples/pivot-sales-sample.xlsx",
    csvUrl: "/samples/pivot-sales-sample.csv",
    sheetName: "Sales",
    columns: ["Region", "Category", "Sub-Category", "Sales", "Quantity", "Order Date"],
    rows: [
      { Region: "East", Category: "Technology", "Sub-Category": "Phones", Sales: "152805", Quantity: "640", "Order Date": "2026-01-15" },
      { Region: "East", Category: "Office Supplies", "Sub-Category": "Binders", Sales: "108728", Quantity: "830", "Order Date": "2026-02-10" },
      { Region: "East", Category: "Furniture", "Sub-Category": "Chairs", Sales: "71654", Quantity: "220", "Order Date": "2026-03-04" },
      { Region: "West", Category: "Technology", "Sub-Category": "Accessories", Sales: "183656", Quantity: "710", "Order Date": "2026-01-28" },
      { Region: "West", Category: "Office Supplies", "Sub-Category": "Paper", Sales: "126438", Quantity: "930", "Order Date": "2026-02-16" },
      { Region: "West", Category: "Furniture", "Sub-Category": "Tables", Sales: "94721", Quantity: "260", "Order Date": "2026-03-21" },
      { Region: "Central", Category: "Technology", "Sub-Category": "Machines", Sales: "101734", Quantity: "410", "Order Date": "2026-01-09" },
      { Region: "Central", Category: "Office Supplies", "Sub-Category": "Storage", Sales: "81104", Quantity: "510", "Order Date": "2026-02-25" },
      { Region: "Central", Category: "Furniture", "Sub-Category": "Bookcases", Sales: "63812", Quantity: "185", "Order Date": "2026-04-11" },
      { Region: "South", Category: "Technology", "Sub-Category": "Copiers", Sales: "90432", Quantity: "340", "Order Date": "2026-01-20" },
      { Region: "South", Category: "Office Supplies", "Sub-Category": "Labels", Sales: "69735", Quantity: "470", "Order Date": "2026-02-03" },
      { Region: "South", Category: "Furniture", "Sub-Category": "Furnishings", Sales: "49263", Quantity: "170", "Order Date": "2026-03-14" },
    ],
    rowField: "Region",
    columnField: "Category",
    valueField: "Sales",
    aggregation: "sum",
    chart: "bar",
    intent: "compare",
  },
  survey: {
    id: "survey",
    button: "Survey, 9 rows",
    fileName: "pivot-survey-sample.csv",
    xlsxName: "pivot-survey-sample.xlsx",
    xlsxUrl: "/samples/pivot-survey-sample.xlsx",
    csvUrl: "/samples/pivot-survey-sample.csv",
    sheetName: "Survey",
    columns: ["Team", "Topic", "Responses"],
    rows: [
      { Team: "Support", Topic: "Onboarding", Responses: "86" },
      { Team: "Support", Topic: "Billing", Responses: "64" },
      { Team: "Support", Topic: "Product", Responses: "41" },
      { Team: "Sales", Topic: "Onboarding", Responses: "22" },
      { Team: "Sales", Topic: "Billing", Responses: "18" },
      { Team: "Sales", Topic: "Product", Responses: "35" },
      { Team: "Success", Topic: "Onboarding", Responses: "54" },
      { Team: "Success", Topic: "Billing", Responses: "29" },
      { Team: "Success", Topic: "Product", Responses: "47" },
    ],
    rowField: "Team",
    columnField: "Topic",
    valueField: "Responses",
    aggregation: "sum",
    chart: "bar",
    intent: "compare",
  },
  traffic: {
    id: "traffic",
    button: "Traffic, 8 rows",
    fileName: "pivot-traffic-sample.csv",
    xlsxName: "pivot-traffic-sample.xlsx",
    xlsxUrl: "/samples/pivot-traffic-sample.xlsx",
    csvUrl: "/samples/pivot-traffic-sample.csv",
    sheetName: "Traffic",
    columns: ["Month", "Channel", "Sessions"],
    rows: [
      { Month: "2026-01-01", Channel: "Organic", Sessions: "4200" },
      { Month: "2026-01-01", Channel: "Paid", Sessions: "1800" },
      { Month: "2026-02-01", Channel: "Organic", Sessions: "4550" },
      { Month: "2026-02-01", Channel: "Paid", Sessions: "2100" },
      { Month: "2026-03-01", Channel: "Organic", Sessions: "5100" },
      { Month: "2026-03-01", Channel: "Paid", Sessions: "1960" },
      { Month: "2026-04-01", Channel: "Organic", Sessions: "4980" },
      { Month: "2026-04-01", Channel: "Paid", Sessions: "2400" },
    ],
    rowField: "Month",
    columnField: "Channel",
    valueField: "Sessions",
    aggregation: "sum",
    chart: "line",
    intent: "compare",
  },
};

const salesExample = {
  caption: "Sum of Sales by Region and Category in the 12-row sales sample. Grand total is 1,194,082.",
  headers: ["Region", "Technology", "Office Supplies", "Furniture", "Total"],
  rows: [
    ["East", "152,805", "108,728", "71,654", "333,187"],
    ["West", "183,656", "126,438", "94,721", "404,815"],
    ["Central", "101,734", "81,104", "63,812", "246,650"],
    ["South", "90,432", "69,735", "49,263", "209,430"],
    ["Grand total", "528,627", "386,005", "279,450", "1,194,082"],
  ],
  note: "West is 404,815 of 1,194,082 sales, about 33.9 percent. Technology is 528,627, about 44.3 percent. These shares are calculated from the table; the tool does not add a percent-of-total column.",
};

const trafficExample = {
  caption: "Sum of Sessions by month and channel in the 8-row traffic sample. Grand total is 27,090.",
  headers: ["Month", "Organic", "Paid", "Total"],
  rows: [
    ["2026-01", "4,200", "1,800", "6,000"],
    ["2026-02", "4,550", "2,100", "6,650"],
    ["2026-03", "5,100", "1,960", "7,060"],
    ["2026-04", "4,980", "2,400", "7,380"],
    ["Grand total", "18,830", "8,260", "27,090"],
  ],
  note: "Organic is 18,830 of 27,090 sessions, about 69.5 percent. April's total, 7,380, is 320 sessions above March's 7,060. Subtract the exported totals yourself; there is no period-over-period aggregation.",
};

const surveyExample = {
  caption: "Sum of Responses by team and topic in the 9-row survey sample. Grand total is 396.",
  headers: ["Team", "Onboarding", "Billing", "Product", "Total"],
  rows: [
    ["Support", "86", "64", "41", "191"],
    ["Sales", "22", "18", "35", "75"],
    ["Success", "54", "29", "47", "130"],
    ["Grand total", "162", "111", "123", "396"],
  ],
  note: "Support accounts for 191 of 396 responses, about 48.2 percent. Onboarding is the largest topic at 162 responses. Count of Responses is also 9, because each source row has one response total.",
};

export const LANDING_PAGES = [
  {
    slug: "online-pivot-table",
    paths: ["/online-pivot-table", "/online-pivot-table.html"],
    title: "Online Pivot Table Maker | Free Browser Tool",
    description: "Create an online pivot table from CSV or XLSX in your browser. Group by date, filter rows, sum, average, or count, then export the checked table.",
    h1: "Online Pivot Table",
    subtitle: "Start with the grouped table. The same workspace can draw a bar, line, or pie chart of those totals after the numbers check out.",
    eyebrow: "Table first, chart second",
    toolName: "Online Pivot Table",
    cta: "Create a Pivot Table",
    defaultView: "table",
    defaultSample: "sales",
    samples: ["sales", "survey", "traffic"],
    keyword: "online pivot table",
    uploadLabel: "Upload a CSV or XLSX file",
    uploadHint: "The pivot table opens first. Switch to Chart only after the totals look right.",
    h2: "Build an online pivot table from the 12-row sales sample",
    example: salesExample,
    authority: `Microsoft documents the desktop version of this layout in [Create a PivotTable to analyze worksheet data](${MICROSOFT_PIVOT_TABLE}). This page uses the same row, column, and value idea, then calculates the summary in the browser instead of in Excel.`,
    related: [
      ["/pivot-chart-without-excel", "Create a pivot chart without Excel"],
      ["/xlsx-to-pivot-chart", "Turn an XLSX file into a pivot chart"],
      ["/pivot-table-vs-pivot-chart.html", "When to keep the table and when to chart it"],
      ["/csv-to-pivot-chart.html", "Prepare a CSV before you summarize it"],
      ["/sitemap.xml", "Sitemap"],
    ],
    faq: [
      { q: "What is an online pivot table here?", a: "It is a summary built in this browser from a CSV or XLSX file. You pick a row field, an optional column field, a value field, and Sum, Average, or Count. The sales sample produces a 4 by 3 table whose grand total is 1,194,082." },
      { q: "Does the chart replace the table?", a: "No. This page opens on the Pivot Table tab. The Chart tab draws a bar, line, or pie view of the same aggregation. Export the table as CSV or XLSX when you need the figures." },
      { q: "Can it calculate percent of total, median, or distinct count?", a: "No. Aggregations are Sum, Average, and Count. You can divide a cell by the grand total yourself. The sales table's Technology column is 528,627 of 1,194,082, which is about 44.3 percent, but that percentage is not a built-in field." },
      { q: "Where does the file go?", a: "The browser reads it locally. Refreshing the page clears the rows. A first XLSX export may request the Excel-writing script; that request is application code, not your spreadsheet." },
    ],
    sections: [
      {
        heading: "Build an online pivot table from the 12-row sales sample",
        paragraphs: [
          "An online pivot table answers a grouped question without scrolling every source row. This page opens that table first. The chart is available on the next tab, using the same fields, filters, and aggregation. Load the sales sample to see the layout before you replace it with your own CSV or XLSX file.",
          "The sample has 12 data rows and 6 columns: Region, Category, Sub-Category, Sales, Quantity, and Order Date. Dates run from 2026-01-09 through 2026-04-11. There is no title row, no blank row, and no subtotal row. That shape is what the parser expects: the first row supplies field names, and every later row is one record.",
          "Set Region as the row field, Category as the column field, and Sales as the value, with Sum. The online pivot table then has four row labels and three category columns. East is 333,187, West is 404,815, Central is 246,650, and South is 209,430. The grand total is 1,194,082. West is the largest region. You can check any cell against the source: West Technology is the single Accessories row, 183,656, because each region has only one row per category in this file.",
        ],
      },
      {
        heading: "Use the survey and traffic samples on the same table",
        paragraphs: [
          "Two more samples sit above the tool so you do not have to invent a file just to learn the controls. The survey sample has 9 rows and 3 columns: Team, Topic, and Responses. Sum of Responses with Team on rows and Topic on columns totals 396. Support is 191, Success is 130, and Sales is 75. Onboarding is 162, Billing is 111, and Product is 123. Support is about 48.2 percent of 396. Switch Aggregation to Count and the grand total becomes 9, one count per source row, because every Responses cell is filled.",
          "The traffic sample has 8 rows: Month, Channel, and Sessions. Months are stored as 2026-01-01, 2026-02-01, 2026-03-01, and 2026-04-01 so the date control can group them. With Month grouped by month, Channel across the columns, and Sum of Sessions, the four monthly totals are 6,000, 6,650, 7,060, and 7,380. The grand total is 27,090. Organic is 18,830 and Paid is 8,260. Stay on the Pivot Table tab while you read those totals. Open Chart afterward if you want the line.",
        ],
      },
      {
        heading: "Choose one aggregation and read the row count",
        paragraphs: [
          "Sum adds numeric values. Average divides that sum by the number of numeric rows in the group. Count counts non-blank value cells. On the sales sample, Average of Sales by Region, with the column field cleared, is 111,062.33 for East, 134,938.33 for West, 82,216.67 for Central, and 69,810 for South. Each of those averages uses 3 source rows. Count of Sales by Region is 3, 3, 3, and 3, and the grand count is 12.",
          "The status line above the result reports how many source rows are included. With no filter it should read 12 of 12 for sales, 9 of 9 for the survey, and 8 of 8 for traffic. If you add a filter, that fraction drops. The on-screen table lists at most 10 row groups. These samples have 4, 3, and 4 groups, so nothing is hidden. A larger file can have more groups than the screen shows; the CSV and XLSX exports include every group, not only the first 10.",
        ],
      },
      {
        heading: "Group dates, then filter the online pivot table",
        paragraphs: [
          "Order Date is detected as a date because the sample uses YYYY-MM-DD. The grouping list offers the original day, week, month, quarter, or year. Month labels for this file are 2026-01, 2026-02, 2026-03, and 2026-04. Sum of Sales by month is 528,627, then 386,005, then 215,638, then 63,812. Those four numbers add back to 1,194,082. Quarter grouping collapses January through March into 2026-Q1 and leaves April in 2026-Q2. Year grouping returns a single 2026 row. Week labels use an ISO week such as 2026-W02. The tool does not compute the change from one month to the next. April minus March is 63,812 minus 215,638, which is a drop of 151,826, and you calculate that from the exported column.",
          "Filters keep or drop dimension values before the aggregation. Add a Category filter and leave only Technology. The four Technology amounts are 152,805, 183,656, 101,734, and 90,432, and the filtered total is 528,627 from 4 of 12 rows. Add a Region filter for West only and the table shrinks to West's three category rows, totaling 404,815. Clear the filter pills to return to 12 of 12. A filter is a checkbox list, not a formula. If a value is missing from the list, it was blank or the field type did not match the text you expected.",
        ],
      },
      {
        heading: "Export the table, then decide whether the chart helps",
        paragraphs: [
          "Export CSV downloads the pivot matrix only: row labels, column labels, and the aggregated numbers. It does not include the 12 source rows. Export XLSX builds a workbook with two worksheets. Sheet1 Raw Data contains the source columns and rows. Sheet2 Pivot Chart contains the field names, the aggregation, the filtered row count, a chart image when the chart canvas is available, and the pivot values starting at row 33. Those pivot cells are stored numbers. They do not refresh when you edit Sheet1 inside Excel. [Create a PivotTable to analyze worksheet data](" + MICROSOFT_PIVOT_TABLE + ") explains Excel's refreshable PivotTable, which this export is not. If you need that object, use the raw sheet as the source inside Excel.",
          "The Chart tab can show a clustered bar, a line, or a pie of the same aggregation. Pie is a reasonable look only when you have one value column and a short list of slices. The sales cross-tab has three category columns, so a bar comparison is the clearer default. Export PNG from the Chart tab. On the Pivot Table tab the chart canvas is not displayed, so the PNG button is the wrong control for a table-only check. The vertical axis labels are scaled in thousands. Sales figures in the hundreds of thousands are readable that way; the survey's 396 responses are easier to audit in the table than on that axis.",
        ],
      },
      {
        heading: "Limits and a network check for this online pivot table",
        paragraphs: [
          "The file input accepts .csv and .xlsx up to 10 MB. Parsing stops above 200 rows, 30 columns, or 6,000 cells. The sales sample is 12 by 6, which is 72 cells, so it is far inside the cap. An XLSX workbook contributes only its first worksheet. A second sheet is ignored, even if the pivot you wanted lives there. Move that sheet to the front, or save it as its own workbook, before you upload it. CSV files use the first row as headers. A report title in row 1 becomes the field names and the real header becomes a data row.",
          "To check privacy, open the browser's network panel, refresh this page, and then load a sample or choose a file. You should see the page's scripts and styles, not a request whose body is the spreadsheet. Choosing a file uses the local file API. The first time you export XLSX, the page may fetch the Excel-writing library if it is not already in memory. That response is script, and the workbook is assembled on the device and saved through a download. Refresh the tab when you want the rows discarded. Nothing in this online pivot table writes your file to an account, because there is no account.",
        ],
      },
    ],
  },
  {
    slug: "pivot-chart-without-excel",
    paths: ["/pivot-chart-without-excel", "/pivot-chart-without-excel.html"],
    title: "Pivot Chart Without Excel | Free Online Tool",
    description: "Create a pivot chart without Excel. Upload CSV or XLSX, choose fields in the browser, and export a bar, line, or pie chart with the summary table.",
    h1: "Pivot Chart Without Excel",
    subtitle: "Create a pivot chart online from a CSV or XLSX file. The browser summarizes the rows and draws a bar, line, or pie chart without installing Excel.",
    eyebrow: "No desktop spreadsheet required",
    toolName: "Pivot Chart Without Excel",
    cta: "Create a Pivot Chart",
    defaultView: "chart",
    defaultSample: "traffic",
    samples: ["traffic", "sales", "survey"],
    keyword: "pivot chart without Excel",
    uploadLabel: "Upload a CSV or XLSX file",
    uploadHint: "Excel does not need to be installed. CSV and XLSX are both read in the browser.",
    h2: "Create a pivot chart without Excel from the traffic sample",
    example: trafficExample,
    authority: `Excel's own steps are Insert, then PivotChart, as described in [Create a PivotChart](${MICROSOFT_PIVOT_CHART}). Use that path when Excel is installed and you need a refreshable PivotChart. Use this page when you only have a browser and a CSV or XLSX file.`,
    related: [
      ["/online-pivot-table", "Build the online pivot table behind the chart"],
      ["/xlsx-to-pivot-chart", "Start from an XLSX workbook"],
      ["/how-to-create-pivot-chart.html", "How to create a pivot chart, step by step"],
      ["/pivot-chart-types-guide.html", "Choose a bar, line, or pie chart"],
      ["/sitemap.xml", "Sitemap"],
    ],
    faq: [
      { q: "Can I create a pivot chart online without Excel installed?", a: "Yes. This page runs in a current browser. Upload a CSV or XLSX file, or load the traffic sample. You get a chart plus the summary table. Excel can stay uninstalled." },
      { q: "Which chart types can I create?", a: "Bar, line, and pie. The traffic sample starts as a line because Month is a date. Sales starts as a bar. Pie fits a short, single-series share, such as one column of category totals, not a two-channel time series." },
      { q: "Is this the same object Excel creates?", a: "No. Excel can store a PivotChart that refreshes from a PivotCache. This page draws the chart in the browser and can export a PNG plus the calculated numbers. The XLSX download does not contain a refreshable PivotChart." },
      { q: "What should I check before exporting the PNG?", a: "Open the Pivot Table tab and confirm the totals. For the traffic sample, the four months are 6,000, 6,650, 7,060, and 7,380, and the grand total is 27,090. Then return to the Chart tab and export PNG." },
    ],
    sections: [
      {
        heading: "Create a pivot chart without Excel from the traffic sample",
        paragraphs: [
          "A pivot chart without Excel is a summarized chart made where the spreadsheet happens to be open: in the browser. You do not need a desktop license, a ribbon, or a PivotChart field pane. Upload a CSV or XLSX file, choose the row, the optional column, the value, and Sum, Average, or Count. This page starts on the Chart tab with the traffic sample already grouped.",
          "That sample has 8 rows and 3 columns: Month, Channel, and Sessions. Each month appears twice, once for Organic and once for Paid. The dates are 2026-01-01, 2026-02-01, 2026-03-01, and 2026-04-01. The date control is set to Month, so the axis labels are 2026-01, 2026-02, 2026-03, and 2026-04 rather than four identical day numbers. Channel stays as the series field. Sum of Sessions draws two lines. Organic moves 4,200, 4,550, 5,100, 4,980. Paid moves 1,800, 2,100, 1,960, 2,400.",
          "The table behind that chart is the part worth checking before you share the picture. Monthly totals are 6,000, 6,650, 7,060, and 7,380. The grand total is 27,090. Organic is 18,830, about 69.5 percent of sessions. Paid is 8,260. March Organic is the highest single cell at 5,100. April is the highest month at 7,380, which is 320 sessions above March. This pivot chart without Excel does not plot that 320-session difference for you. Read it off the totals, or add the subtraction in a spreadsheet after you export the matrix.",
        ],
      },
      {
        heading: "What you skip from the Excel PivotChart steps",
        paragraphs: [
          "[Create a PivotChart](" + MICROSOFT_PIVOT_CHART + ") tells Excel users to select a cell in a table, choose Insert, then PivotChart, pick a location, and then tick fields. That produces a chart tied to a PivotTable. Change the fields and both objects update. Refresh the PivotTable after the source grows and the chart follows. A pivot chart without Excel on this page does not create that pair of objects. It calculates one aggregation in memory and paints one Chart.js view.",
          "The practical difference shows up when the source changes. Here, you upload the new file or paste the new extract and rebuild the same field choices. You do not press Refresh on a PivotCache. That is a limitation, and it is also why the page can run on a machine where Excel is not installed. The export options match that model: PNG for the picture, CSV for the aggregated grid, and XLSX for the source rows plus the calculated grid and a static chart image. None of those files keeps refreshing after you leave the page.",
        ],
      },
      {
        heading: "Switch samples when the question is not a time series",
        paragraphs: [
          "The sales sample is the comparison chart. It has 12 rows. Put Region on the axis, Category in the series, and Sum of Sales in the values. West leads at 404,815. Technology leads the columns at 528,627. The grand total is 1,194,082. A bar chart matches that question because the regions are names, not dates. The line option still draws, but it implies an order the region names do not have. Leave the chart type on Bar for that sample.",
          "The survey sample is smaller: 9 rows, 396 responses, three teams and three topics. Support is 191, Success is 130, and Sales is 75. It is useful when you want to see Count versus Sum. Sum of Responses totals 396. Count of Responses totals 9. The chart axis on this page is labeled in thousands, so a few hundred responses look cramped. Use the survey sample to learn the field controls, then judge the teams from the pivot table. Use traffic or sales when you want a chart whose scale matches the axis.",
          "Chart intent has two settings, Compare and Share. Compare keeps a bar or line. Share can switch an automatic chart to pie when there is a single value column and between 2 and 12 groups. The traffic sample has two series, Organic and Paid, so it stays a line even if you press Share. Clear the Channel field, leave Month as the only dimension, and a pie of four monthly totals becomes available. Four slices that add to 27,090 are readable. A pie of many product names is not. The [chart type guide](/pivot-chart-types-guide.html) covers that choice in more detail.",
        ],
      },
      {
        heading: "Prepare the file that feeds the chart",
        paragraphs: [
          "The chart is only as specific as the rows you upload. Use one header row. Keep one record on each following row. Do not include a title, a blank spacer, or a total line that already sums the rows. A total line is added again by Sum, so the grand total would be too large. For the traffic shape, that mistake would mean an extra row labeled Total whose Sessions cell is 27,090, and the chart would plot that label beside the months.",
          "Dates need a shape the detector recognizes. YYYY-MM-DD works. A month name plus a four-digit year, such as Jan 2026, works. A day, a month name, and a year, such as 15 Jan 2026, works. A locale string such as 4/1/2026 is not detected as a date, so the Month grouping list will not appear until you set the field type to Date yourself. After you override the type, choose Month, Quarter, Year, Week, or Day. The traffic file uses the first of each month so Month and Day labels do not disagree about which day the row represents.",
          "Numbers can include commas. A currency symbol can be detected as currency. Percent signs are divided by 100. Text in the value column is skipped for Sum and Average. Count skips blank value cells and counts the rest. If the status says 0 rows can be used, the value field is the wrong column or the cells are labels. The traffic Sessions column is plain integers, so all 8 rows are usable.",
        ],
      },
      {
        heading: "Filters, limits, and the export you can hand to someone else",
        paragraphs: [
          "A filter changes both the chart and the table. On the traffic sample, filter Channel to Organic only. The line should keep four points: 4,200, 4,550, 5,100, and 4,980, totaling 18,830, and the status should read 4 of 8 rows. Filter to Paid only and the points are 1,800, 2,100, 1,960, and 2,400, totaling 8,260. If both filters are cleared, you are back to 8 of 8 and 27,090. The chart does not warn you that a filter is active beyond that row count, so read the count before you export.",
          "Limits are 10 MB, 200 rows, 30 columns, and 6,000 cells. The traffic sample is 8 by 3. A year of daily rows with several channels can cross 200 rows; aggregate it to months in the source, or split the file, before uploading. Only the first worksheet of an XLSX file is read. CSV is often the easier extract when you are creating a pivot chart without Excel on a shared computer, because you can save the sheet as CSV from another program and skip the workbook entirely. The [CSV to pivot chart guide](/csv-to-pivot-chart.html) covers header cleanup.",
          "Export PNG from the Chart tab after the table totals match the source. The image is the chart canvas, not a vector file. Export CSV when a colleague needs the 4 by 2 session grid without the picture. Export XLSX when they need the 8 source rows and the summary in one workbook. Tell them the summary will not refresh inside Excel. For a refreshable chart, the Microsoft article above is the path that actually creates one, and it requires Excel.",
        ],
      },
      {
        heading: "Confirm the page stayed on your machine",
        paragraphs: [
          "Creating a pivot chart without Excel should not mean uploading the workbook to an account you do not control. This page has no sign-in. The sample buttons copy rows that shipped with the site, or, on the XLSX page, fetch the sample workbook file itself. Your own file is chosen through the file input and parsed in the tab. Watch the network panel while you select a local file: the selection should not create a POST of that file. Refreshing discards the parsed rows.",
          "The first XLSX export can still cause a network request, because the Excel-writing code is loaded when that button is used. The response is a script chunk. The workbook bytes are built afterward in the browser and saved as pivot-chart-maker-export.xlsx. If you only need the PNG, you can skip that download. The PNG is produced from the chart already on the page.",
        ],
      },
    ],
  },
  {
    slug: "xlsx-to-pivot-chart",
    paths: ["/xlsx-to-pivot-chart", "/xlsx-to-pivot-chart.html"],
    title: "XLSX to Pivot Chart Online | Free Browser Tool",
    description: "Make an XLSX to pivot chart online from the first worksheet. Summarize rows locally in your browser, then export a PNG, CSV, or values workbook.",
    h1: "XLSX to Pivot Chart",
    subtitle: "Upload an Excel workbook and chart the first worksheet in the browser. CSV still works, but this page is built for .xlsx files.",
    eyebrow: "First worksheet only",
    toolName: "XLSX to Pivot Chart",
    cta: "Create a Chart from XLSX",
    defaultView: "chart",
    defaultSample: "sales",
    loadSampleAsXlsx: true,
    samples: ["sales", "traffic", "survey"],
    keyword: "XLSX to pivot chart",
    uploadLabel: "Upload an XLSX workbook",
    uploadHint: "The first worksheet is read. CSV is accepted if you already exported the sheet as text.",
    h2: "Turn the sample workbook into an XLSX to pivot chart",
    example: salesExample,
    authority: `An .xlsx file is the Office Open XML workbook Excel writes by default. Microsoft lists it among the workbooks Excel opens in [File formats that are supported in Excel](${MICROSOFT_FILE_FORMATS}). [Create a PivotChart](${MICROSOFT_PIVOT_CHART}) is the desktop command this page does not replace.`,
    related: [
      ["/online-pivot-table", "Review the same file as an online pivot table"],
      ["/pivot-chart-without-excel", "Create the chart without Excel installed"],
      ["/create-pivot-chart-in-excel.html", "Build a native PivotChart inside Excel"],
      ["/csv-to-excel", "Convert a CSV into an XLSX workbook first"],
      ["/sitemap.xml", "Sitemap"],
    ],
    faq: [
      { q: "Which sheet becomes the XLSX to pivot chart?", a: "The first worksheet only. Later sheets are not listed and are not combined. In the sample workbook the first sheet is named Sales and has 12 data rows. Move the sheet you want to the front before uploading a larger workbook." },
      { q: "Will the exported workbook still be a live Excel pivot chart?", a: "No. The download has a raw-data sheet and a sheet of calculated values plus a static chart image. Excel will not refresh those values when the source rows change. Build a native PivotChart in Excel if you need refresh." },
      { q: "Why did my date column stay as text?", a: "The detector looks for YYYY-MM-DD, a month name plus a year, or a day plus a month name plus a year. A cell that Excel displays as 1/15/2026 can arrive as that text and stay a text field. Set the type override to Date, or write the sample's ISO dates before you save the workbook." },
      { q: "What are the size limits?", a: "10 MB, 200 rows, 30 columns, and 6,000 cells, counted on the first worksheet after the header. The sales workbook is 12 rows and 6 columns. Files over those limits need a smaller extract." },
    ],
    sections: [
      {
        heading: "Turn the sample workbook into an XLSX to pivot chart",
        paragraphs: [
          "An XLSX to pivot chart starts from a workbook, not from a typed-in table. Click Sales, 12 rows on this page and the browser fetches pivot-sales-sample.xlsx, then reads it with the same parser your upload uses. The file has one worksheet, named Sales. Row 1 is the header. Rows 2 through 13 are the records. The columns are Region, Category, Sub-Category, Sales, Quantity, and Order Date. There is no second sheet to discover, which is the situation you want when you are testing the tool.",
          "The chart opens on Sum of Sales by Region and Category. East totals 333,187, West 404,815, Central 246,650, and South 209,430. The grand total is 1,194,082. Technology across regions is 528,627, Office Supplies is 386,005, and Furniture is 279,450. Those figures are the check on the parser. If your own workbook's first sheet is the same extract, you should land on the same cross-tab. If the totals differ, the header row, a hidden total row, or a different first sheet is the first place to look.",
          "You can download the [sales sample workbook](/samples/pivot-sales-sample.xlsx) and open it in Excel before uploading it again. The Order Date cells are stored as the text 2026-01-15 and similar ISO dates, not as Excel date serials. That is deliberate. The detector recognizes that text, shows the date grouping control, and can bucket the column by day, week, month, quarter, or year. A workbook you save from Excel after converting those cells to dates may display 1/15/2026 instead. Override the field type to Date if the grouping list does not appear.",
        ],
      },
      {
        heading: "What the XLSX reader keeps and what it ignores",
        paragraphs: [
          "The reader takes the first worksheet and turns its used range into rows. Formatting, column widths, colors, comments, defined names, and charts that already live in the file are not imported. Formulas come across as their displayed values when the file stored a cached result, which is normal for an .xlsx saved by Excel. A formula with no cached value can arrive blank. PivotTables that already exist in the workbook are not reused. This XLSX to pivot chart builds a new summary from the flat rows. If the workbook only contains a finished PivotTable and not the original transactions, upload that source range instead.",
          "[File formats that are supported in Excel](" + MICROSOFT_FILE_FORMATS + ") distinguishes .xlsx from older .xls files and from CSV text. This page accepts .xlsx and .csv. It does not accept .xls, .xlsb, or .ods. If you still have a binary .xls workbook, open it in a spreadsheet and save a copy as .xlsx, or save the sheet as CSV and use the CSV input. The [CSV to Excel tool](/csv-to-excel) can also wrap a CSV in a new .xlsx file, with values stored as text so leading zeros survive. That text storage is a different choice from this chart page, which parses numbers so it can sum them.",
          "Merged cells, a title in row 1, and subtotal rows are the usual reasons an XLSX to pivot chart looks wrong. Unmerge the headers so each column has one name. Delete the title so row 1 is Region, Category, and the rest. Delete subtotal rows so Sum does not add a number that is already a sum. Blank header cells are renamed Column 1, Column 2, and so on. Duplicate header names get a numeric suffix. The sample workbook does not need those repairs: all six headers are unique and filled.",
        ],
      },
      {
        heading: "Two more workbooks: traffic and survey",
        paragraphs: [
          "Traffic, 8 rows fetches pivot-traffic-sample.xlsx. Its first sheet has Month, Channel, and Sessions. Month is an ISO date, so the chart can be a line of monthly sessions. The monthly totals are 6,000, 6,650, 7,060, and 7,380. Organic plus Paid is 27,090. Use this workbook when the Excel file you care about is a time series and you want to see the date grouping control after a real .xlsx parse, not only after an in-memory sample.",
          "Survey, 9 rows fetches pivot-survey-sample.xlsx, with Team, Topic, and Responses. The cross-tab totals 396 responses. Support is 191, Success is 130, and Sales is 75. It is a clean small workbook for checking that the first sheet's header became the field list. The chart axis is labeled in thousands, and 396 is a small number on that scale, so confirm Survey in the pivot table. The sales and traffic workbooks are the ones that match the axis.",
          "Each sample button replaces the previous rows. The file name shown in the tool becomes the .xlsx name, and the row count should match the button: 12, 8, or 9. If a fetch fails, the page falls back to the same rows held in the script and shows that fallback in the status area. The totals stay comparable either way. Your own upload never uses that fallback. A local file that fails does so with the parser's error, and the previous sample remains until a file parses.",
        ],
      },
      {
        heading: "From the first worksheet to a chart you can export",
        paragraphs: [
          "After the sheet loads, check the detected types before you trust the picture. Sales and Quantity should be numbers. Order Date should be a date. Region, Category, and Sub-Category should be text. Then set the row, column, and value. The sample chart uses Region, Category, and Sum of Sales. Change the value to Quantity and the grand total becomes 5,675, which is 640+830+220+710+930+260+410+510+185+340+470+170. That second total is a useful sanity check that you are not still looking at Sales.",
          "Filters apply to the parsed rows, not to an Excel slicer. Filtering Category to Furniture leaves 71,654, 94,721, 63,812, and 49,263, which sum to 279,450, from 4 of 12 rows. The chart follows the filter. Clear it before you export a chart that is supposed to represent the whole sheet. The on-screen pivot table shows up to 10 row groups; this sample has 4 regions, so the screen and the export match. A workbook with more than 10 row groups still exports every group in the CSV and in the XLSX summary.",
          "Export PNG from the Chart tab once the 1,194,082 total is the total you expected. Export CSV if you only need the cross-tab. Export XLSX if you want the parsed rows and the summary together. The exported workbook is a new file, pivot-chart-maker-export.xlsx. It contains the raw rows and the calculated pivot values. It does not contain an Excel PivotChart, a PivotCache, or a Refresh button. [Create a PivotChart](" + MICROSOFT_PIVOT_CHART + ") remains the way to build that native object from Sheet1 Raw Data after you open the export in Excel.",
        ],
      },
      {
        heading: "Size, privacy, and the first-sheet rule",
        paragraphs: [
          "Stay under 10 MB, 200 data rows, 30 columns, and 6,000 cells on the sheet that is first in the workbook. A financial model with many formula columns can pass the row limit and still fail the column or cell limit. Copy the columns you intend to group onto a new first sheet, or save that range as a separate .xlsx, and upload the smaller file. The ignored sheets do not count toward the limit because they are not parsed, but they do count toward the 10 MB file size. A 9 MB workbook whose first sheet is tiny can still be too large to open here.",
          "The sample .xlsx files are requested from this site when you click a sample button. That request is the sample file, and its contents are the public rows listed above. Choosing your own workbook should not add a request that contains those private bytes. The file input reads the disk through the browser. Confirm that in the network panel: select a local .xlsx and watch for a new upload. Parsing and the chart update should proceed without one. The Excel-writing library may download the first time you export XLSX. That payload is code. The file you save is generated locally.",
          "When the workbook is wrong for a chart, fix the sheet rather than the picture. One header row, ISO dates or an explicit Date override, numeric measures, no subtotal rows, and the intended sheet in front. Then build the XLSX to pivot chart again. The [online pivot table](/online-pivot-table) shows the same engine with the table opened first, which is the better check when the workbook is messy. This page keeps the chart in front because the search that lands here is asking for the picture.",
        ],
      },
    ],
  },
];

export function getSample(id) {
  return SAMPLES[id] || null;
}

export function normalizePath(pathname) {
  const path = String(pathname || "/").split("?")[0].replace(/\/+$/, "") || "/";
  return path.toLowerCase();
}

export function getLandingPage(pathname) {
  const path = normalizePath(pathname);
  return LANDING_PAGES.find((page) => page.paths.includes(path)) || null;
}

function linkPattern() {
  return /\[([^\]]+)\]\((https?:\/\/[^)\s]+|\/[^)\s]+)\)/g;
}

export function inlineToHtml(value) {
  const source = String(value);
  let html = "";
  let last = 0;
  for (const match of source.matchAll(linkPattern())) {
    html += escapeHtml(source.slice(last, match.index));
    const href = match[2];
    const external = href.startsWith("http");
    html += `<a href="${escapeHtml(href)}"${external ? ' target="_blank" rel="noopener noreferrer"' : ""}>${escapeHtml(match[1])}</a>`;
    last = match.index + match[0].length;
  }
  html += escapeHtml(source.slice(last));
  return html;
}

export function linkedText(value) {
  return String(value).replace(linkPattern(), "$1");
}

function escapeHtml(value) {
  return String(value).replace(/[&<>"']/g, (character) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[character]);
}

export function pageWordCount(page) {
  const chunks = [
    page.subtitle,
    ...page.sections.flatMap((section) => [section.heading, ...section.paragraphs]),
    page.example.caption,
    ...page.example.headers,
    ...page.example.rows.flat(),
    page.example.note,
    page.authority,
    ...page.faq.flatMap((item) => [item.q, item.a]),
  ];
  return linkedText(chunks.join(" ")).split(/\s+/).filter(Boolean).length;
}

export function renderLandingDocument(page) {
  const origin = "https://www.pivotchartmaker.org";
  const canonical = `${origin}/${page.slug}`;
  const nav = [
    ["/", "Pivot Chart Maker"],
    ["/online-pivot-table", "Online Pivot Table"],
    ["/pivot-chart-without-excel", "Pivot Chart Without Excel"],
    ["/xlsx-to-pivot-chart", "XLSX to Pivot Chart"],
    ["/tools", "Data Tools"],
    ["/sitemap.xml", "Sitemap"],
  ]
    .filter(([href]) => href !== `/${page.slug}`)
    .map(([href, label]) => `<a href="${href}">${escapeHtml(label)}</a>`)
    .join("");
  const exampleRows = page.example.rows
    .map((row) => `<tr>${row.map((cell) => `<td>${escapeHtml(cell)}</td>`).join("")}</tr>`)
    .join("");
  const sections = page.sections
    .map(
      (section) =>
        `<section><h2>${escapeHtml(section.heading)}</h2>${section.paragraphs.map((paragraph) => `<p>${inlineToHtml(paragraph)}</p>`).join("")}</section>`
    )
    .join("");
  const faq = page.faq
    .map((item) => `<section><h3>${escapeHtml(item.q)}</h3><p>${inlineToHtml(item.a)}</p></section>`)
    .join("");
  const related = page.related.map(([href, label]) => `<li><a href="${href}">${escapeHtml(label)}</a></li>`).join("");
  const schema = {
    "@context": "https://schema.org",
    "@graph": [
      {
        "@type": "WebApplication",
        name: page.h1,
        url: canonical,
        applicationCategory: "BusinessApplication",
        operatingSystem: "Any",
        offers: { "@type": "Offer", price: "0", priceCurrency: "USD" },
      },
      {
        "@type": "FAQPage",
        mainEntity: page.faq.map((item) => ({
          "@type": "Question",
          name: item.q,
          acceptedAnswer: { "@type": "Answer", text: linkedText(item.a) },
        })),
      },
    ],
  };
  return `<!doctype html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>${escapeHtml(page.title)}</title>
  <meta name="description" content="${escapeHtml(page.description)}" />
  <meta name="robots" content="index, follow" />
  <link rel="canonical" href="${canonical}" />
  <link rel="icon" type="image/svg+xml" href="/favicon.svg" />
  <meta property="og:type" content="website" />
  <meta property="og:url" content="${canonical}" />
  <meta property="og:title" content="${escapeHtml(page.title)}" />
  <meta property="og:description" content="${escapeHtml(page.description)}" />
  <meta property="og:image" content="${origin}/og-image.svg" />
  <script type="application/ld+json">${JSON.stringify(schema)}</script>
</head>
<body>
  <div id="root">
    <header>
      <a href="/">Pivot Chart Maker</a>
      <nav aria-label="Primary navigation">${nav}</nav>
    </header>
    <main>
      <h1>${escapeHtml(page.h1)}</h1>
      <p>${escapeHtml(page.subtitle)}</p>
      <form>
        <label>${escapeHtml(page.uploadLabel)}
          <input type="file" name="spreadsheet" accept=".csv,.xlsx,text/csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" />
        </label>
      </form>
      <p>${escapeHtml(page.uploadHint)}</p>
      <article>
        ${sections}
        <section>
          <h2>Checked result</h2>
          <p>${escapeHtml(page.example.caption)}</p>
          <table>
            <thead><tr>${page.example.headers.map((header) => `<th>${escapeHtml(header)}</th>`).join("")}</tr></thead>
            <tbody>${exampleRows}</tbody>
          </table>
          <p>${escapeHtml(page.example.note)}</p>
        </section>
        <section>
          <h2>Source note</h2>
          <p>${inlineToHtml(page.authority)}</p>
        </section>
        <section>
          <h2>Questions about this page</h2>
          ${faq}
        </section>
        <section>
          <h2>Related pages</h2>
          <ul>${related}</ul>
        </section>
      </article>
    </main>
    <footer>
      <a href="/">Home</a>
      <a href="/tools">Data tools</a>
      <a href="/sitemap.xml">Sitemap</a>
    </footer>
  </div>
  <script type="module" src="/src/main.jsx"></script>
</body>
</html>
`;
}
