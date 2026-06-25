# Consumer Finance Market Intelligence Tracker

Scrapes Google News RSS for **70+ consumer finance companies** across 10 market segments covering credit cards and home loans. Automatically classifies articles into funding, partnerships, regulatory, product, M&A, disruptions, global expansion, and competitive intelligence events. Outputs a fully styled, Power BI-ready Excel file.

## Setup

```bash
pip install -r requirements.txt
python consumer_finance_tracker.py
```

## Output: `consumer_finance_intel.xlsx`

| Sheet | Contents |
|---|---|
| **Summary** | Overview stats, topic counts, segment breakdown |
| **Companies** | Full company list with segment, HQ, founded, website |
| **News** | All headlines from last 12 months, tagged by topic |
| **Financials** | Funding rounds, valuations, investors extracted from news |
| **Partnerships** | Partnership announcements with partner names |
| **Rates** | APR/rate mentions extracted from headlines by product type |
| **Rewards** | Rewards program announcements — cashback, points, miles, sign-up bonuses, and program changes |
| **Disruptions** | Negative news, outages, lawsuits, layoffs, data breaches, executive departures, credit losses, and reputational events |
| **Global Expansion** | International market entries, cross-border partnerships, foreign regulatory approvals, and overseas product launches |
| **Competition** | Competitive moves — new entrants, pricing wars, feature parity announcements, market share shifts, and head-to-head comparisons |

## Companies Tracked (70+)

### Segments

- **Issuers – Major Banks** (Chase, Bank of America, Citi, Wells Fargo, Capital One, Discover, American Express...)
- **Issuers – Fintech / Neobank Cards** (Apple Card/GS, Chime, Current, Varo, Dave, Step, Petal, Deserve...)
- **Credit Card Networks** (Visa, Mastercard, American Express, Discover/NACS...)
- **BNPL & Installment Credit** (Affirm, Klarna, Afterpay, Sezzle, Splitit, Zip, PayPal Pay Later...)
- **Credit Building & Underwriting** (Experian Boost, Credit Karma, Self Financial, Kikoff, Grow Credit, Perch...)
- **Mortgage Originators – Traditional** (Rocket Mortgage, UWM, loanDepot, Fairway, Guild Mortgage, PennyMac...)
- **Mortgage Originators – Fintech** (Better.com, Morty, Tomo, Flyhomes, Homie, Orchard...)
- **Mortgage Infrastructure & Tech** (Blend, Encompass/ICE, Maxwell, SimpleNexus/nCino, Polly, Optimal Blue...)
- **Mortgage Servicers & Secondary Market** (Mr. Cooper, Pennymac, Nationstar, Fannie Mae, Freddie Mac, Ginnie Mae...)
- **Regulatory Bodies** (CFPB, OCC, FHFA, HUD, FTC, Federal Reserve...)

## Power BI Setup

1. Open Power BI Desktop
2. **Get Data → Excel Workbook** → select `consumer_finance_intel.xlsx`
3. Load all 10 sheets
4. Recommended visuals:
   - **Bar chart**: News volume by Company (add Topic as legend)
   - **Pie/donut**: Articles by Segment
   - **Timeline/line chart**: Funding events over time (Date field)
   - **Table**: Financials sheet — Company, Round, Amount, Valuation, Investors
   - **Table**: Partnerships — Company, Partner, Headline + URL
   - **Table**: Rates — Company, Product Type, APR/Rate Mentioned, Date, URL
   - **Table**: Rewards — Company, Reward Type, Bonus/Rate Mentioned, Date, URL
   - **Table**: Disruptions — Company, Disruption Type, Severity, Date, URL
   - **Map visual**: Global Expansion — Country, Company, Entry Type, Date
   - **Matrix**: Competition — Company vs. Competitor, Event Type, Date
   - **Slicers**: Segment, Topic, Product Type (credit card vs. mortgage), Reward Type, Disruption Type, Region, Date range, Company

## Adding Companies

Edit the `COMPANIES` list in `consumer_finance_tracker.py`:

```python
{"name": "Your Company", "segment": "Fintech Cards", "hq": "City, Country", "founded": 2021, "website": "example.com"},
```

## Notes
- Pulls last **12 months** of news only
- Each run overwrites the file with fresh data
- Google News RSS is free and doesn't require an API key
- Rate/APR extraction is regex-based from headlines — verify important figures against source URLs
- Regulatory actions (CFPB enforcement, OCC guidance) are tagged separately for compliance monitoring
- The **Rates sheet** is unique to this tracker: it captures any APR, rate, or basis-point mention in headlines to help track competitive pricing signals over time
- The **Rewards sheet** tracks cashback percentages, points multipliers, sign-up bonuses, miles offers, and program changes — useful for competitive benchmarking across issuers
- The **Disruptions sheet** flags negative signals — lawsuits, breaches, outages, layoffs, executive exits, delinquency spikes, and regulatory enforcement actions — with a Severity tag (Low / Medium / High)
- The **Global Expansion sheet** tracks international moves: new country launches, cross-border partnerships, foreign license approvals, and overseas acquisitions — tagged by region
- The **Competition sheet** captures head-to-head competitive signals: new market entrants, feature copycat announcements, pricing undercuts, and analyst market share commentary
