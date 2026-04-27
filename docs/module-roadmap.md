# Module Roadmap

---

## Released

### v1.0 — Core Platform + Annual Ingest
**Status:** ✅ Complete

- Home sheet with branded UI and module buttons
- Raw Data sheet — master booking table
- YoY Dashboard — KPIs across all years with delta column
- Monthly Trends — revenue and booking count by month
- Top Clients — 15 tracked clients with revenue heatmap
- Brand Mix — hotel brand portfolio with % share
- **Module 1: Annual Report Ingestion**
  - File picker dialog
  - Auto-detects Main sheet
  - Column mapping by header name (tolerates variations)
  - Booking ID deduplication
  - Alternating row styling
  - Full dashboard recalculate on import

---

## In development

### v1.1 — RFP Formatter
**Status:** 🔧 Active development  
**Target:** Q2 2024

**Module 2: RFP Formatter**
- Client selector from Raw Data
- Date range filter
- Auto-populate RFP fields:
  - Client name, program name
  - Property, city, country
  - Arrival / departure dates
  - Room nights, revenue target
- Generate formatted RFP sheet
- Export as PDF or new workbook
- Brand-consistent styling

---

## Planned

### v1.2 — Benchmark Report
**Status:** 📋 Planned  
**Target:** Q3 2024

**Module 3: Benchmark Report**
- Load HB portfolio average reference data
- Compare selected client KPIs vs benchmarks:
  - Revenue per booking
  - Room nights per program
  - Brand mix vs portfolio average
  - Seasonal pattern match
- Highlight over / under performance per metric
- One-page benchmark summary sheet
- Export as PDF for client presentations

---

### v2.0 — Budget Forecaster
**Status:** 💡 Concept  

- Use historical YoY trends to project next year targets
- Room night and revenue forecast by month
- Adjustable growth rate slider
- Exportable forecast sheet

---

### v2.1 — Property Scorecard
**Status:** 💡 Concept  

- Rate each property by repeat usage, revenue per room night
- Geographic spread analysis
- Preferred partner flag
- Sortable scorecard table

---

## Distribution model

Associates share `HB_Intelligence_Toolkit.xlsm` directly with clients.
Client data never leaves their machine — no cloud upload, no login required.

New module releases are delivered as updated `.bas` files.
Clients re-import in under 60 seconds and new buttons come alive on the Home sheet.