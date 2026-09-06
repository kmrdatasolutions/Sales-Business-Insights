# Sales Business Insights — Final Report Edition

A professional browser-based business intelligence dashboard that turns uploaded Excel, CSV, TXT and SQL data into executive analysis and downloadable reports.

## Included
- Excel / CSV / TXT / SQL upload
- Automatic metric and column detection
- Data-quality profiling
- Executive Overview
- Product & Customer analysis
- Regional & Operations analysis
- Business / Key Insights
- Growth, profit, margin and loss analysis
- Light / Dark theme
- Company name configuration
- Download Report menu
- Excel report export
- PDF executive report export
- PowerPoint presentation export
- Exports respect the currently selected filters

## Company branding
Click the **Settings** icon in the top bar and enter the company name. It is stored in the browser and used on exported reports.

## Report downloads
After uploading data, click **Report** in the top bar and choose Excel, PDF or PowerPoint.

## Run

```powershell
npm.cmd install
npm.cmd run dev
```

Then open the local URL shown by Vite, usually `http://localhost:5173/`.

## Notes on mixed data
The dashboard uses semantic aliases plus numeric/category/date fallbacks, so common variations such as Revenue/Amount/Turnover, Net Profit/Earnings, Qty/Units, State/Country/Region, Product/Item and Customer/Client can still be analyzed. When a metric cannot be confidently detected, the UI marks it as unavailable instead of inventing a value.
