# NYC Affordable Housing Lottery Tracker

A live tracker for NYC affordable-housing lotteries, built for a NYC DOE teacher commuting to
**Crotona Park East, Bronx**. Every apartment has to pass three independent tests:

1. **Lottery eligibility** — against the income minimum *and* maximum actually published in the advertisement, never inferred from the AMI label.
2. **Real-world affordability** — rent and utilities as a share of gross monthly income, against your own thresholds.
3. **Commute** — a real subway trip to Freeman St / 174 St, not straight-line distance.

## The web app (`index.html`)

Open it and it pulls **every currently open rental lottery live** from the NYC Housing Connect
public API, reads each advertisement's full unit / income / preference tables, and scores them.
Curated re-rentals that never appear on Housing Connect (NYC HDC and HPD's marketing agents) are
merged in from `data/manual-listings.json`.

No build step, no server, no API key. It is a single static HTML file.

* Baseline inputs are pre-filled, so it works without touching anything.
* Change salary, household size, rent ceilings, commute limits or geography and everything
  recalculates instantly.
* Inputs and application statuses are saved in your browser's local storage only — nothing is uploaded.

### Running it locally

```bash
python3 -m http.server 8777
```

Then open <http://127.0.0.1:8777>. (Opening the file directly with `file://` will not work — the
browser blocks the `fetch` calls for the local data files.)

## Data

| File | What it is |
|---|---|
| `data/commute.json` | 322 Manhattan and Brooklyn stations with the scheduled time to Freeman St / 174 St, built from the MTA GTFS weekday feed (06:30–09:00 departures) with a transfer-aware shortest path. Stations sharing a name are kept separate — there is no free transfer between the three different 125 St stations. |
| `data/manual-listings.json` | Re-rentals from NYC HDC and HPD's marketing-agent list, transcribed from the official advertisement PDFs. Dated — re-check against the linked advertisement before applying. |

## The spreadsheet

`scripts/refresh.py` rebuilds the 12-sheet Excel workbook from the same live data:

```bash
python3 scripts/refresh.py --dry-run     # report what changed
python3 scripts/refresh.py               # rebuild the workbook
```

It preserves everything you typed: the INPUTS sheet, application statuses, dates applied,
confirmation and log numbers, and the whole Application Tracker and New Listings sheets.
Lotteries that have closed drop out of the live feed and are reported as REMOVED.

Requires `openpyxl`. LibreOffice is used to recalculate formulas if it is installed.

## Key rules this encodes

* **Income is gross, before any deductions** — *"Gross income, before any expenses or deductions, is used to determine income eligibility, with the exception of self-employment income."* ([Marketing Handbook §5-4.A(3)](https://www.nyc.gov/assets/hpd/downloads/pdfs/services/marketing-handbook-8-21.pdf)) Take-home pay and commuting costs do not reduce it.
* **Siblings are one household, automatically** — "Immediate Family Member" expressly includes a sibling, so brothers qualify under §5-2(ii)(a) with no proof of financial interdependence.
* **10% Municipal Employee / Military Veteran preference** — since 15 Nov 2025, for applicants paid by the City of New York. ([Addendum §5-1(C)](https://www.nyc.gov/assets/hpd/downloads/pdfs/services/MH-addendum-MEMV-preference.pdf)) It does not appear on every lottery.
* **NYC residents are processed before non-residents**, including inside that preference.
* **Asset limit** = the four-person HUD income limit at the unit's AMI, whatever your household size. Retirement accounts are excluded.

## Sources

[NYC Housing Connect](https://housingconnect.nyc.gov/PublicWeb/search-lotteries) ·
[NYC HDC re-rentals](https://www.nychdc.com/find-re-rentals) ·
[HPD re-rental marketing agents](https://www.nyc.gov/site/hpd/services-and-information/find-affordable-housing-re-rentals.page) ·
[HPD Area Median Income](https://www.nyc.gov/site/hpd/services-and-information/area-median-income.page) ·
[MTA GTFS](https://www.mta.info/developers)
