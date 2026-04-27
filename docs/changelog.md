# Changelog

All notable changes to the RVHB Intelligence Toolkit are documented here.
Format: `[vX.X] — YYYY-MM-DD`

---

## [v1.0] — 2024-04-27

### Added
- Project structure initialized under `RVHB-Intelligence/`
- Git repository with `.gitignore` for Python, Excel, and OS artifacts
- `HB_Intelligence_Toolkit.xlsm` — main workbook with 6 sheets:
  - 🏠 Home — branded landing sheet with module buttons and status line
  - 📥 Raw Data — master booking table, pre-loaded with 2022–2023 data
  - 📊 YoY Dashboard — 333 live SUMPRODUCT formulas, 4-year comparison
  - 📅 Monthly Trends — revenue and booking count by month per year
  - 👥 Top Clients — 15 tracked clients, revenue heatmap
  - 🏨 Brand Mix — 9 hotel brands, revenue and % share per year
- `HB_Core.bas` — shared constants, navigation, utilities
- `HB_Module1_Ingest.bas` — annual report ingestion with deduplication
- `HB_Module2_RFP.bas` — RFP formatter stub (v1.1)
- `HB_Module3_Benchmark.bas` — benchmark report stub (v1.2)
- `data/samples/HBMasterReport_2023.xls` — sample report for testing
- `docs/setup-guide.md` — developer build and import instructions
- `docs/module-roadmap.md` — full versioned feature roadmap
- `src/python/write_vba_sources.py` — utility to write all VBA files from script

### Architecture decisions
- VBA split into one `.bas` file per module for clean version control
- `HB_Core.bas` owns all shared constants — modules never hardcode values
- Python builds the `.xlsm` shell; VBA handles all runtime behavior
- Booking ID used as deduplication key — safe to re-run ingestion any time

---

## [v1.1] — TBD

- Module 2: RFP Formatter (active development)

---

## [v1.2] — TBD

- Module 3: Benchmark Report (planned)