# Sales Performance Dashboard (Advanced Excel)

An interactive sales analytics workbook built to demonstrate advanced Microsoft Excel
skills: formula-driven KPIs, PivotTables, slicers, and dynamic charts on a
1,800-row synthetic sales dataset (2024–2025, Irish regions).

## What's inside

- **`RawData` sheet** — 1,800 transaction records (Excel Table) covering Region,
  Category, Product, Salesperson, Units Sold, Unit Price, Unit Cost. Revenue, Cost,
  and Profit are calculated with formulas (`=Units*Price`, etc.), not hardcoded values.
- **`Summary` sheet** — KPI cards (Total Revenue, Cost, Profit, Units Sold, Avg Order
  Value) and breakdown tables (Revenue by Region, Revenue by Category, Monthly Revenue
  Trend) built with `SUMIFS`/`AVERAGE`, plus bar, pie, and line charts.
- **`PivotTable Dashboard` sheet** *(you add this in Excel — see below)* — a PivotTable
  with Region/Category/Salesperson slicers and a PivotChart.

## Skills demonstrated

`SUMIFS` · `AVERAGE` · Excel Tables · PivotTables · Slicers · PivotCharts ·
conditional formatting · dynamic charting · KPI dashboard design · data modeling

## How to extend it into a full PivotTable dashboard

The raw data and formula layer are done — add the PivotTable/slicer layer yourself
in Excel (this is the hands-on part worth practicing):

1. Open the workbook, click inside the `SalesData` table on `RawData`.
2. **Insert → PivotTable** → New Worksheet → rename the sheet `PivotTable Dashboard`.
3. Drag `Region` to Rows, `Category` to Columns, `Revenue` to Values (set to Sum).
4. **PivotTable Analyze → Insert Slicer** → tick `Region`, `Category`, `Salesperson`.
5. **PivotTable Analyze → PivotChart** → insert a clustered column chart from the
   same PivotTable.
6. Format the slicers and chart to match the KPI cards on the `Summary` sheet, then
   arrange everything on one screen for a clean dashboard view.
7. Save, and take a screenshot/GIF of the finished dashboard for this README.

## Dataset

Synthetic but realistic: 1,800 orders across 5 Irish regions (Dublin, Cork, Galway,
Limerick, Waterford), 4 product categories, and 8 salespeople, dated Jan 2024–Dec 2025.
Generated for demonstration purposes — not real sales data.

## Tools used

Microsoft Excel · openpyxl (dataset/formula generation)
