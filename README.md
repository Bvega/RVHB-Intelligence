# RVHB Intelligence Toolkit

A modular Excel-based intelligence platform built for HelmsBriscoe associates.
Clients receive a single `.xlsm` file with button-driven features — no installs,
no browser, no SaaS. Just Excel.

---

## What it does

| Module | Status | Description |
|--------|--------|-------------|
| Module 1 — Annual Ingest | ✅ v1.0 | Import any HB Master Report and auto-update all dashboards |
| Module 2 — RFP Formatter | 🔧 v1.1 | Generate sourcing-ready RFPs from Raw Data |
| Module 3 — Benchmark | 📋 v1.2 | Compare client KPIs against HB portfolio averages |

---

## Project structure
RVHB-Intelligence/
├── dist/               → Files delivered to clients
├── src/
│   ├── python/         → Excel workbook builders
│   ├── vba/            → One .bas file per module
│   └── web/            → Browser-based updater tool
├── data/
│   └── samples/        → Sample reports for testing
├── docs/               → Setup guide, roadmap, changelog
└── releases/           → Version snapshots
---

## Quick start

### For associates (end users)
1. Receive `HB_Intelligence_Toolkit.xlsm` from your team
2. Open it in Excel — enable macros when prompted
3. Follow the one-time VBA setup on the Home sheet (30 seconds)
4. Click **Update with New Year's Report** each year

### For developers
See [`docs/setup-guide.md`](docs/setup-guide.md) for full build instructions.

---

## One-time VBA setup (end users)

1. Open `HB_Intelligence_Toolkit.xlsm`
2. Press `Alt + F11` to open the VBA editor
3. `File` → `Import File` → select `HB_Core.bas` first
4. Repeat for `HB_Module1_Ingest.bas`, `HB_Module2_RFP.bas`, `HB_Module3_Benchmark.bas`
5. Close the VBA editor
6. Right-click each button on the Home sheet → **Assign Macro** → type the macro name shown

> Import order matters: `HB_Core.bas` must always be imported first.

---

## Tech stack

- **Excel VBA** — macro engine, all module logic
- **Python + openpyxl** — workbook builder scripts
- **Git** — version control and release management

---

## Versioning

| Version | Description |
|---------|-------------|
| v1.0 | Home sheet, Raw Data, YoY Dashboard, Monthly Trends, Top Clients, Brand Mix, Module 1 Ingest |
| v1.1 | Module 2 — RFP Formatter |
| v1.2 | Module 3 — Benchmark Report |

---

## License

Internal use — HelmsBriscoe / RVHB Intelligence  
© 2024. Confidential.