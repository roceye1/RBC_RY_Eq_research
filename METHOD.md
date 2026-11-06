# Calculation methods, assumptions and limitations

## Financial statements and bridges

Company reported figures are from RBC's FY2025 annual report, in **millions of CAD**:

| Item | FY2024 | FY2025 | FY2025 minus FY2024 |
|---|---:|---:|---:|
| Total revenue | 57,344 | 66,605 | +9,261 |
| Provisions for credit losses | 3,232 | 4,362 | +1,130 |
| Non-interest expense | 34,250 | 36,592 | +2,342 |
| Income tax | 3,622 | 5,282 | +1,660 |
| Net income | 16,240 | 20,369 | +4,129 |

Exact bridge (C$M): `16,240 + 9,261 - 1,130 - 2,342 - 1,660 = 20,369`.

**Growth calculation:** `(Current / Prior) − 1`. Thus revenue increases about `16.15%` and net income about `25.43%` in FY2025. All displayed bridge numbers round to C$0.01B. **Accounting comparability:** FY2023 data are IFRS-17 consistent with subsequent years, while a direct FY2022 comparison on the former insurance accounting basis is inappropriate without restatement.

**Segment reconciliation FY2025 (C$M):** Total revenue `19,854 Personal + 8,562 Commercial + 22,378 Wealth + 1,321 Insurance + 14,426 Capital Markets + 64 Corporate Support = 66,605`. Reported net income `7,105 + 3,020 + 4,289 + 828 + 5,393 − 266 = 20,369`. We do not conflate segment subtotals with the consolidated figures.

## Credit measures

- Gross impaired loans / related loans and acceptances: FY2025 0.83%, FY2024 0.59%.
- PCL / average net loans and acceptances: FY2025 0.43%, FY2024 0.35%.
- Q3 FY2026 PCL/loans ratio: 36 basis points (annualized quarterly measure as reported by RBC), and CET1 13.5%. **These quarterly ratios are not annual FY2025 ratios** and are clearly period-labeled.

## Price-only total horizon returns

For an NYSE-listed share with split-adjusted reference closes `P_start` and `P_end`, the **price return** is `(P_end / P_start − 1) × 100%`. These returns **exclude cash dividends and withholding taxes**; for foreign banks they include the USD share-price impact of FX. The 1- and 10-year chart graphics are sourced from the originally authored deck (Alpaca SIP) and the endpoint returns were recalculated. Do not call them 'total return'.

For weekly correlations, take the log change `r_t = ln(P_t / P_(t−1))`, compute Pearson correlation over 260 weekly changes for the original Nov 2021–Oct 2026 sample, and round to two decimals. Corporate-action adjustments use split-adjusted price observations, not dividends. The correlation values are descriptive of the sample and can change markedly.

## Consensus EPS and revenue (CAD)

| Metric | FY2025 actual | FY2026E | FY2027E |
|---|---:|---:|---:|
| Revenue, C$B | 66.605 | 72.42 | 76.44 |
| Adjusted/normalized diluted EPS, C$ | 14.43 | 16.42 | 17.91 |

`FY2026E revenue growth = 72.42 / 66.605 − 1 = 8.73%`  
`FY2027E revenue growth = 76.44 / 72.42 − 1 = 5.55%`  
`FY2026E adjusted EPS growth = 16.42 / 14.43 − 1 = 13.79%`  
`FY2027E adjusted EPS growth = 17.91 / 16.42 − 1 = 9.07%`.

The FY2025 *adjusted* EPS (C$14.43) is used in the forecast growth calculation, not reported FY2025 diluted EPS of C$14.07; mixing the two would misrepresent the growth rate. FY2026 was not yet reported at the 5-Nov-2026 research cut-off; RBC announced FY2026 results for 3 December 2026.

## Indicative valuation, not a broker price target

The method is a simple **12-month sensitivity** on FY2027 diluted adjusted EPS in CAD, a terminal earnings multiple, and an explicit illustrative FX conversion:

`Scenario USD price = FY2027 EPS in C$ × terminal P/E × USD per C$`  
`Modeled price return = Scenario USD price / US$203.76 − 1`

| Case | FY2027 EPS (CAD) | Terminal P/E | USD/CAD factor (USD per CAD) | Implied USD share price | Price return vs US$203.76 |
|---|---:|---:|---:|---:|---:|
| Bear | 17.00 | 14.0x | 0.723 | US$172.07 | −15.6% |
| Base | 17.91 | 16.5x | 0.723 | US$213.66 | +4.9% |
| Bull | 19.00 | 19.0x | 0.723 | US$261.00 | +28.1% |

**Sources and authorship:** FY2027E `C$17.91` is the consensus snapshot from Yahoo Finance. Every terminal multiple is **an author-chosen scenario assumption**; bear and bull EPS numbers are also author assumptions. This is a scenario model, not an optimized target or claim of expected price. It omits dividends, transaction costs, taxes, changing exchange rates, and paths of interim earnings. **Sensitivities are not probabilities** and do not guarantee future outcomes.

The indicated reference-price multiple based on FY2027E adjusted EPS is `US$203.76 / (C$17.91 × 0.723) ≈ 15.7x`. This is **not identical** to a dynamic vendor NTM forward P/E. Vendor multiples are placed separately on the peer-analysis slide.

## Interpretations and research integrity

- Forward P/E is generally more directly useful for a deposit-taking bank than EV/EBITDA. Price/book, tangible price/book and ROE are valid extensions, but must use matched period definitions; unsupported valuations are not inserted.
- Historical EPS-surprise observations use only one vendor and one bank to avoid false cross-bank comparison where definitions differ.
- Geographic revenue shares and country/office footprints are not interchangeable.
- The two historical price-chart images were originally author-supplied. Raw proprietary price bars are **not** distributed in this public-ready research pack.
- Any investment interpretation is based on publicly available information and explicitly identified modeling assumptions. It is educational, not individualized investment advice.