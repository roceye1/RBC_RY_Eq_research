# Royal Bank of Canada (RY) — Independent Equity Research

*The presentation preview and downloadable decks will appear here after the binary deliverables are uploaded.*

**Independent equity research by Aamir Sohaib**  
**Research cut-off:** 5 November 2026  
**Security:** Royal Bank of Canada, NYSE/TSX: `RY`  
**Financial scope:** FY2025 reported results; Q3 FY2026 interim data; FY2026–FY2027 analyst estimates  
**Indicative conclusion:** **HOLD / WATCHLIST**, subject to the assumptions, limitations, and non-advisory disclaimer below.

## Executive summary

RBC combines scale, diversified earnings, strong capital and high profitability. In FY2025, it reported **C$66.605 billion revenue**, **C$20.369 billion net income**, and **16.3% ROE**. Its **13.5% Q3 FY2026 CET1 ratio** suggests a significant capital buffer. However, FY2025 impaired-loan and credit-loss ratios increased, and public forward P/E comparisons show a premium to selected Canadian and global banks.

An independently constructed, **illustrative—not a sell-side analyst—scenario analysis** uses FY2027 adjusted EPS assumptions, 14×–19× exit P/E values and an explicitly assumed US$0.723/C$ exchange rate. The resulting indicative share prices span **US$172 to US$261**, with a **~US$214 base case** compared with the **US$203.76 5-Nov-2026 close**. This is not a prediction or individualized investment recommendation; see [METHOD.md](METHOD.md).

## View the work

> **Publication status:** The analysis, methodology, references and CSV data are uploaded. The four PowerPoint/PDF deliverables and supporting source/image assets are not yet in this repository; links below will work once the remaining files are uploaded. See [UPLOAD_REMAINING.md](UPLOAD_REMAINING.md).

| Deliverable | File | Purpose |
|---|---|---|
| Full research report (PDF) | [23-slide PDF](deliverables/RBC_Independent_Equity_Research_Full.pdf) | Easy to read in a browser |
| Full editable presentation | [PowerPoint](deliverables/RBC_Independent_Equity_Research_Full.pptx) | Editable research deck |
| LinkedIn carousel | [8-page portrait PDF](deliverables/RBC_LinkedIn_8_Slide_Carousel.pdf) | Designed for LinkedIn document posts |
| LinkedIn editable slides | [PowerPoint](deliverables/RBC_LinkedIn_8_Slide_Carousel.pptx) | Editable carousel source |
| LinkedIn post copy | [LINKEDIN_POST.md](LINKEDIN_POST.md) | Ready to personalize and publish |
| Sources | [SOURCES.md](SOURCES.md) | Direct URLs and dates |
| Methodology | [METHOD.md](METHOD.md) | Definitions, scenarios, caveats |
| Data tables | [`data/`](data/) | Audit-friendly CSV exports |

## Research workflow

1. **Business and geography:** compare revenue sources, segments and country exposures.
2. **Reported accounting:** reconcile FY2024–FY2025 revenue and earnings, including PCL, expenses and taxes.
3. **Earnings quality and balance sheet:** evaluate ROE, CET1, efficiency, impaired loans and PCL ratios.
4. **Public-market context:** compare selected bank forward P/E multiples and the 1-/10-year price-only history of seven NYSE securities.
5. **Portfolio perspective:** quantify weekly price-return correlations across the same seven tickers.
6. **Forecasts:** distinguish FY2026E and FY2027E public consensus from RBC's reported results.
7. **Independent valuation:** run bear/base/bull sensitivity analyses before reaching a conditional conclusion.

## Core figures

| Item | Value | Basis |
|---|---:|---|
| FY2025 revenue | C$66.605B | RBC FY2025 annual report |
| FY2025 reported net income | C$20.369B | RBC FY2025 annual report |
| FY2025 ROE | 16.3% | RBC FY2025 annual report |
| FY2025 efficiency ratio | 54.9% | RBC FY2025 annual report |
| FY2025 gross impaired loans / related loans | 0.83% | RBC FY2025 annual report |
| FY2025 PCL / average net loans | 0.43% | RBC FY2025 annual report |
| Q3 FY2026 CET1 | 13.5% | RBC Q3 FY2026 release |
| FY2027E normalized/adjusted diluted EPS | C$17.91 | Yahoo Finance analyst consensus, snapshot as of research cut-off |
| Reference NYSE share price | US$203.76 | 5 November 2026 close |

## Important methodology choices

- **CAD is the default reporting currency.** USD is used for NYSE security returns and valuation; the conversion for forward EPS uses an explicitly **illustrative** factor of **US$0.723 per C$1**. Historical valuation conclusions are sensitive to foreign exchange.
- **FY2023 onward** is used for the reported revenue/net-income history in the revised deck to avoid treating the accounting break from **IFRS 17** as operational growth.
- **EPS forecasts** are not company-reported facts and may change; FY2026 results are scheduled for release on **3 December 2026**.
- **Five segments + Corporate Support** reconcile to consolidated revenue and income. Central items are not silently omitted.
- **Peer valuations are indicative.** Different banking structures, financial-reporting calendars, local GAAP treatments, currency exposures and vendor conventions limit comparability. The deck intentionally excludes PEG and EV/EBITDA comparisons.
- **Share-price charts are price-only,** split-adjusted USD data and **exclude dividends**; therefore they are not total-return comparisons.
- **The earnings-surprise slide is RBC-only** to avoid false cross-country equivalence from mixed consensus vendors/EPS definitions.
- The source deck's **two historical share-price chart graphics** were retained with permission from the author-supplied presentation. Their numerical end returns and the key correlations were independently checked against Alpaca data during the project review. Raw monthly/weekly price bars are *not* reproduced in this repository; recreating every plotted point requires authorized market-data access.

## Rebuild / customize

> The Python build script and its graphical dependencies are part of the downloadable publication package but have not yet been uploaded into this repository.

The slide-building source is `build_research.py` (Python / `python-pptx`). The chart assets are in `assets/`. For regeneration:

```bash
python -m pip install python-pptx
python build_research.py
```

This generates editable PowerPoint decks in `deliverables/`. To regenerate the PDFs locally, convert them with LibreOffice/PowerPoint or your preferred renderer. The generated PDF files are provided already and do not require a local build.

## Scope, licensing and disclaimer

This is an **independent educational case study**. It is not affiliated with, commissioned by, sponsored by or endorsed by Royal Bank of Canada, Bloomberg, Alpaca, Yahoo Finance or any other referenced organization. No investment-advisory relationship is created. The opinions are author-level interpretations, not guaranteed outcomes. Forward-looking estimates and scenario valuations are inherently uncertain; investors should consult original filings and current data before acting.

**Third-party material:** RBC filings, market-data facts, and references remain the property of their respective owners. The embedded historical performance graphics originated from the supplied presentation. Public posting does not grant a license to reuse third-party source material. No blanket open-source license is asserted for the full compilation; contact the author for reuse permissions.

## Repository status

This repository is the public home for the project. The PDFs, editable decks, research data, methodology, and supporting assets are maintained here. The original publishing scripts in the research package are for initializing a new repository and should not be run against this already-created repository.
