# DataVision AI — Smart Dashboard Generator

A frontend MVP that lets users upload Excel (`.xlsx`, `.xls`) or CSV files and automatically creates an interactive analytics dashboard.

## Features
- Upload one or multiple Excel/CSV files together; rows are combined by matching column names (processed locally in the browser)
- Automatic column detection for numeric and categorical fields
- KPI cards for totals, profit, orders/records and averages
- Trend area chart, category donut chart and top performers
- Professional chart gallery with bar, line, pie, scatter, treemap, heatmap, waterfall and funnel charts
- Data cleaning workspace to detect/remove duplicates and empty columns, inspect missing values, and review/convert mixed data types
- Region/category filters and searchable data explorer
- Rule-based quick insights
- Ask questions about your data in the AI Chat (totals, averages and highest/lowest breakdowns)
- Export filtered data to Excel
- Print dashboard / Save as PDF using the browser print dialog
- Dark and light theme toggle
- Responsive layout

## Requirements
- Node.js 18 or newer
- npm

## Run in VS Code
1. Extract the ZIP file.
2. Open the `datavision-ai` folder in VS Code.
3. Open Terminal → New Terminal.
4. Run:

```bash
npm install
npm run dev
```

5. Open the local URL shown in the terminal (usually `http://localhost:5173`).

## Build for production

```bash
npm run build
npm run preview
```

## Notes
- The project includes sample sales data so the dashboard is visible immediately.
- Upload spreadsheets with column headers in the first row. You can select multiple CSV/XLSX/XLS files in one upload; rows are appended, columns are matched by name (case-insensitive), and missing columns are left blank. The first worksheet of each Excel file is used.
- The dashboard detects common column names such as Sales, Revenue, Profit, Orders, Date, Region, Category and Product.
- Data processing happens in the browser; files are not sent to a server.
- AI Insights and AI Chat in this starter project use local, rule-based analysis; they are not connected to an external AI model.
- AI Chat answers use the records currently included by the dashboard filters and the detected column names.
- Open Charts to explore all eight chart formats; choose a numeric measure and the charts update with the current dashboard filters.
- Open Data Cleaning to review issues and explicitly apply row removal, missing-value filling, empty-column removal, and type conversions. Invalid conversion values become blank; changes remain in the current browser session until exported.
- In this MVP, the first worksheet is loaded when a workbook is uploaded.
