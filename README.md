# Treasury Report AI

**Automated lender reporting package generator for Saudi corporate borrowers.**

Treasury Report AI eliminates the 3-5 day manual process of assembling monthly bank reporting packages. Built for Saudi mid-market companies managing multiple bank relationships.

## What It Does

- **Covenant Compliance Monitoring** — Auto-calculates DSCR, Debt/Equity, ICR, Current Ratio, Debt/EBITDA, and Net Worth covenants per bank-specific definitions
- **Multi-Bank Support** — Handles different covenant structures across SNB, Al Rajhi, SABB, SIDF, and other Saudi banks simultaneously
- **Islamic Finance Native** — Supports profit-rate based calculations, Murabaha facility structures, and SIDF development loan covenants
- **Traffic-Light Dashboard** — Interactive HTML dashboard with color-coded compliance status across all facilities
- **PDF Certificate Generation** — Professional covenant compliance certificates ready to send to your relationship manager
- **Early Warning System** — Flags covenants within 15% of breach threshold before they become problems

## Quick Start

```bash
pip install pandas openpyxl reportlab
python covenant_engine.py
```

Input your financial data via the Excel template (`sample_data.xlsx`), and the engine generates:
1. A PDF Covenant Compliance Certificate
2. An interactive HTML dashboard with utilization charts

## Sample Output

The prototype evaluates 11 covenants across 4 Saudi banks in under 1 second:
- **9 Compliant** (green)
- **2 Warnings** (amber) — SNB Current Ratio at 11.8% headroom, Al Rajhi Net Worth at 12.5% headroom
- **0 Breaches** (red)

## Project Structure

```
treasury-report-ai/
├── covenant_engine.py      # Core calculation engine (DSCR, leverage, ICR, etc.)
├── report_generator.py     # PDF covenant compliance certificate generator
├── dashboard.py            # Interactive HTML dashboard with charts
├── sample_data.xlsx        # Template with sample financial data
└── output/
    ├── Covenant_Compliance_Certificate.pdf
    └── dashboard.html
```

## Who This Is For

Saudi corporate treasury teams at companies with:
- 3+ bank lending relationships
- SAR 100M+ in total credit facilities
- Monthly covenant reporting obligations
- Mix of conventional and Islamic finance facilities

## Roadmap

- [ ] SAP/Oracle data import connectors
- [ ] Automated bank-specific PDF formatting
- [ ] Email scheduling to relationship managers
- [ ] Historical trend analysis and breach prediction
- [ ] Arabic language report generation

## About

Built by [Mohammed Abdul Gaffar, CFA](https://decoded.finance) — 4+ years managing SAR 1bn+ in credit facilities at a major Saudi industrial group. This tool automates the exact workflow I did manually every month.

Part of the [decoded.finance](https://decoded.finance) ecosystem.

## License

MIT
