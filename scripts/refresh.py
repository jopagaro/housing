#!/usr/bin/env python3
"""Refresh the housing workbook from live sources.

  python3 scripts/refresh.py                 # rebuild the workbook from live data
  python3 scripts/refresh.py --dry-run       # just report what changed
  python3 scripts/refresh.py --out other.xlsx

What it does
  1. Pulls every currently-open rental lottery from the NYC Housing Connect public API,
     including the full unit / income / preference tables from each advertisement.
  2. Adds the curated re-rentals in data/manual-listings.json (HDC and marketing agents,
     which are not on Housing Connect).
  3. Works out the nearest useful station and a real door-to-door commute for every
     listing from data/commute.json (built from the MTA GTFS schedule).
  4. Rebuilds the workbook, then copies your own entries back in: the INPUTS sheet,
     application statuses, dates applied, confirmation and log numbers, and the whole
     Application Tracker and New Listings sheets.
  5. Recalculates with LibreOffice when it is installed.

Listings that have closed or been withdrawn simply stop appearing in the live feed; they
are reported as REMOVED and dropped from the database, so the sheet never goes stale.
"""
import argparse, json, math, os, shutil, subprocess, sys, tempfile, datetime as dt
from urllib import request

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
API = "https://a806-housingconnectapi.nyc.gov/HPDPublicAPI/api"
AMI4 = 169600                      # 2026 four-person 100% AMI (HPD) - basis for asset limits
WALK_MPM = 80                      # metres walked per minute
WAIT = 4                           # average wait at the origin station
WALK_FROM_DEST = 7                 # station to workplace, matches the workbook default
DEFAULT_OUT = os.path.join(os.path.dirname(ROOT), "NYC_Affordable_Housing_System.xlsx")

SEARCH_BODY = {"UnitTypes": [], "NearbyPlaces": [], "NearbySubways": [], "Amenities": [],
               "Applied": None, "HPDUserId": None, "Boroughs": [], "Neighborhoods": [],
               "HouseholdSize": None, "Income": "", "HouseholdType": 1, "OwnerTypes": [],
               "PreferanceTypes": [], "LotteryTypes": [], "Min": None, "Max": None,
               "RentalSubsidy": None}


def _get(url, data=None):
    req = request.Request(url, method="POST" if data else "GET")
    req.add_header("Accept", "application/json")
    body = None
    if data is not None:
        req.add_header("Content-Type", "application/json")
        body = json.dumps(data).encode()
    with request.urlopen(req, body, timeout=60) as r:
        return json.loads(r.read().decode())


def haversine(a, b, c, d):
    R, t = 6371000, math.pi / 180
    x, y = (c - a) * t, (d - b) * t
    u = math.sin(x / 2) ** 2 + math.cos(a * t) * math.cos(c * t) * math.sin(y / 2) ** 2
    return 2 * R * math.asin(math.sqrt(u))


def transit_for(lat, lon, stations, max_walk_min=12):
    """Best commute plus every line within walking range (the corridor needs all of them)."""
    if lat is None or lon is None:
        return None
    near = []
    for s in stations:
        dist = haversine(lat, lon, s["lat"], s["lon"])
        if dist > 1400:
            continue
        walk = dist / WALK_MPM
        near.append({"s": s, "walk": walk, "dist": dist,
                     "total": walk + WAIT + s["min"] + WALK_FROM_DEST})
    if not near:
        return None
    near.sort(key=lambda x: x["total"])
    within = [x for x in near if x["dist"] <= max_walk_min * WALK_MPM] or near
    best = within[0]
    lines = set()
    for x in within:
        lines.update(x["s"]["lines"].split())
    return {"best": best, "lines": lines,
            "alts": [x["s"]["name"] + " (" + x["s"]["lines"] + ")" for x in within[1:3]]}


def corridor_of(boro, tr, zipcode=None):
    if zipcode == "10044":
        return "Not on Manhattan island"
    if boro == "Manhattan":
        return "Manhattan"
    if boro == "Brooklyn":
        L = tr["lines"] if tr else set()
        if {"2", "5"} & L:
            return "Brooklyn 2/5"
        if {"3", "4"} & L:
            return "Brooklyn 3/4"
        return "Brooklyn other"
    return boro or "Unknown"


BLANK = dict(side="n/a", cb="", pub=None, laundry="", elevator="", status="NEW",
             prefpct=None, util=0, utilinc="", amen="", notes="")


def fill_transit(row, stations):
    tr = transit_for(row.get("lat"), row.get("lon"), stations)
    if tr:
        b = tr["best"]; s = b["s"]
        row.update(station=s["name"], lines=s["lines"], walk=round(b["walk"]),
                   ride=round(s["min"]), transfers=s["transfers"], dest=s["dest"],
                   direct="YES" if s["transfers"] == 0 else "NO",
                   ttype="None" if s["transfers"] == 0 else
                         ("Same-platform" if s["transfers"] == 1 else "Walking transfer"),
                   tdetail=("No transfer - direct %s train" % s["route"]) if s["transfers"] == 0
                           else ("%d transfer(s); best boarding at %s" % (s["transfers"], s["name"])),
                   route="Walk %d min to %s, then %s toward %s" % (
                         round(b["walk"]), s["name"], s["route"], s["dest"]))
        if tr["alts"]:
            row["tdetail"] += ". Also walkable: " + "; ".join(tr["alts"])
    else:
        row.update(station="REVIEW - no coordinates", lines="", walk=None, ride=None,
                   transfers=1, dest="Freeman St", direct="NO", ttype="Easy same-station",
                   tdetail="", route="Confirm from the listing")
    row["corridor"] = corridor_of(row.get("boro"), tr, row.get("zip"))
    return row


def fetch_housing_connect(stations):
    search = _get(API + "/Lottery/SearchLotteries", SEARCH_BODY)
    rentals = search.get("rentals") or []
    rows, seen_lotteries = [], []
    for meta in rentals:
        lid = meta["lotteryId"]
        try:
            ad = _get(API + "/Lottery/GetLotteryAdvertisement?lotteryId=%s" % lid)
        except Exception as e:
            print("  ! could not load advertisement %s: %s" % (lid, e), file=sys.stderr)
            continue
        seen_lotteries.append((lid, ad.get("lotteryName")))
        b = (ad.get("lotteryBuildings") or [{}])[0]
        prefs = ad.get("lotterySetAsidePreferences") or []
        me = next((p for p in prefs if "municipal" in (p.get("name") or "").lower()), None)
        cb = next((p for p in prefs if "community board" in (p.get("name") or "").lower()), None)

        grouped = {}
        for u in ad.get("units") or []:
            k = (u.get("unitLayoutTypeName"), u.get("unitRegulatoryMechanismAmi"), u.get("actualRent"))
            g = grouped.setdefault(k, {"u": u, "inc": {}})
            for x in u.get("unitIncome") or []:
                g["inc"][x["houseHoldSize"]] = (x["minimumIncome"], x["maximumIncome"])

        for (layout, ami, rent), g in grouped.items():
            inc = g["inc"]
            sizes = sorted(inc)
            if not sizes:
                continue
            mins = [inc[s][0] for s in sizes if inc[s][0] is not None]
            def cap(n):
                if n in inc:
                    return inc[n][1]
                return "N/A" if n < min(sizes) else None
            layout_l = (layout or "").lower()
            beds = 0 if "studio" in layout_l else next(
                (int(c) for c in (layout or "") if c.isdigit()), 0)
            row = dict(BLANK)
            row.update(
                id="HC-%s" % lid, name=ad.get("lotteryName") or "",
                addr=", ".join(x for x in [b.get("address"), b.get("city"),
                                           b.get("state"), b.get("zip")] if x),
                zip=b.get("zip"), hood=meta.get("neighborhood") or "",
                boro=(meta.get("borough") or "").strip(),
                cb=cb["name"] if cb else "",
                program=ad.get("lotteryDescription") or "Housing Connect lottery",
                url="https://housingconnect.nyc.gov/PublicWeb/details/%s" % lid,
                appurl="https://housingconnect.nyc.gov/PublicWeb/details/%s" % lid,
                deadline=(ad.get("endDate") or "")[:10] or None,
                unit=layout, beds=beds, ami="%s%% AMI" % ami,
                minimum=min(mins) if mins else None,
                max1=cap(1), max2=cap(2), max3=cap(3),
                rent=rent,
                asset=(g["u"].get("maximumAssetCap") or 0) or round(AMI4 * (ami or 0) / 100),
                mepref="YES" if me else "NO",
                prefpct=(me["requirementAllocation"] / 100.0) if me else None,
                amen=", ".join(a.get("name", "") for a in (ad.get("amenities") or [])),
                lat=float(b["latitude"]) if b.get("latitude") else None,
                lon=float(b["longitude"]) if b.get("longitude") else None,
                laundry="Yes" if "laundry" in (str(ad.get("amenities")) or "").lower() else "",
                elevator="Yes" if "elevator" in (str(ad.get("amenities")) or "").lower() else "",
                notes="Live from the Housing Connect API. Preferences advertised: " +
                      (", ".join("%s %s%%" % (p["name"], p["requirementAllocation"]) for p in prefs)
                       or "none"),
                utilinc="See advertisement", util=0,
            )
            rows.append(fill_transit(row, stations))
    return rows, seen_lotteries


def load_manual(stations):
    path = os.path.join(ROOT, "data", "manual-listings.json")
    out = []
    for m in json.load(open(path)):
        row = dict(BLANK)
        row.update(m)
        row["minimum"] = m.get("min")
        row.setdefault("zip", None)
        out.append(fill_transit(row, stations))
    return out


# ---------------------------------------------------------------- user state
USER_COLS = ["Application status", "Date applied", "Confirmation number", "Log number"]
COPY_SHEETS = ["APPLICATION_TRACKER", "NEW_LISTINGS"]


def read_state(path):
    if not os.path.exists(path):
        return None
    from openpyxl import load_workbook
    wb = load_workbook(path)
    st = {"rows": {}, "inputs": {}, "sheets": {}}
    if "ACTIVE_OPPORTUNITIES" in wb.sheetnames:
        ws = wb["ACTIVE_OPPORTUNITIES"]
        hdr = {ws.cell(row=3, column=c).value: c for c in range(1, ws.max_column + 1)}
        need = [c for c in USER_COLS if c in hdr]
        for r in range(4, ws.max_row + 1):
            key = (ws.cell(row=r, column=hdr["Listing ID"]).value,
                   ws.cell(row=r, column=hdr["Unit type"]).value,
                   ws.cell(row=r, column=hdr["AMI tier"]).value)
            if not key[0]:
                continue
            vals = {c: ws.cell(row=r, column=hdr[c]).value for c in need}
            if any(v not in (None, "", "NEW") for v in vals.values()):
                st["rows"][key] = vals
    if "INPUTS" in wb.sheetnames:
        ws = wb["INPUTS"]
        for r in range(1, ws.max_row + 1):
            v = ws.cell(row=r, column=2).value
            if v is not None and not (isinstance(v, str) and v.startswith("=")):
                st["inputs"][r] = v
    for name in COPY_SHEETS:
        if name in wb.sheetnames:
            ws = wb[name]
            st["sheets"][name] = [[ws.cell(row=r, column=c).value
                                   for c in range(1, ws.max_column + 1)]
                                  for r in range(4, ws.max_row + 1)]
    return st


def apply_state(path, st):
    if not st:
        return 0
    from openpyxl import load_workbook
    wb = load_workbook(path)
    restored = 0
    ws = wb["ACTIVE_OPPORTUNITIES"]
    hdr = {ws.cell(row=3, column=c).value: c for c in range(1, ws.max_column + 1)}
    for r in range(4, ws.max_row + 1):
        key = (ws.cell(row=r, column=hdr["Listing ID"]).value,
               ws.cell(row=r, column=hdr["Unit type"]).value,
               ws.cell(row=r, column=hdr["AMI tier"]).value)
        if key in st["rows"]:
            for col, val in st["rows"][key].items():
                if col in hdr and val not in (None, ""):
                    ws.cell(row=r, column=hdr[col]).value = val
            restored += 1
    wi = wb["INPUTS"]
    for r, v in st["inputs"].items():
        cur = wi.cell(row=r, column=2).value
        if cur is None or not (isinstance(cur, str) and cur.startswith("=")):
            wi.cell(row=r, column=2).value = v
    for name, rows in st["sheets"].items():
        if name not in wb.sheetnames:
            continue
        ws2 = wb[name]
        for i, rowvals in enumerate(rows):
            for j, v in enumerate(rowvals):
                if v not in (None, ""):
                    ws2.cell(row=4 + i, column=1 + j).value = v
    wb.save(path)
    return restored


def recalc(path):
    soffice = shutil.which("soffice") or "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    if not os.path.exists(soffice):
        print("  (LibreOffice not found - formulas will compute when you open the file)")
        return
    with tempfile.TemporaryDirectory() as td:
        subprocess.run([soffice, "--headless", "--norestore",
                        "--convert-to", "xlsx", "--outdir", td, path],
                       capture_output=True, timeout=300)
        made = os.path.join(td, os.path.basename(path))
        if os.path.exists(made):
            shutil.copy(made, path)
            print("  recalculated with LibreOffice")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--out", default=DEFAULT_OUT)
    a = ap.parse_args()

    stations = json.load(open(os.path.join(ROOT, "data", "commute.json")))
    print("Fetching live Housing Connect lotteries...")
    live, lotteries = fetch_housing_connect(stations)
    print("  %d open rental lotteries -> %d unit/AMI rows" % (len(lotteries), len(live)))
    for lid, nm in lotteries:
        print("    %s  %s" % (lid, nm))
    manual = load_manual(stations)
    print("  %d curated re-rental rows" % len(manual))
    rows = live + manual

    prev = read_state(a.out)
    if prev is not None:
        old_ids = {k[0] for k in prev["rows"]}
        new_ids = {r["id"] for r in rows}
        gone = sorted(i for i in old_ids if i not in new_ids and i.startswith("HC-"))
        if gone:
            print("  REMOVED (no longer open): " + ", ".join(gone))

    if a.dry_run:
        print("\nDry run - nothing written.")
        return

    tmp = os.path.join(tempfile.gettempdir(), "listings_live.json")
    json.dump(rows, open(tmp, "w"), default=str)
    env = dict(os.environ, LISTINGS_JSON=tmp, WORKBOOK_OUT=a.out)
    r = subprocess.run([sys.executable, os.path.join(ROOT, "scripts", "build_workbook.py")],
                       cwd=os.path.join(ROOT, "scripts"), env=env,
                       capture_output=True, text=True)
    if r.returncode:
        print(r.stdout); print(r.stderr, file=sys.stderr); sys.exit(1)
    print(" ", r.stdout.strip())
    n = apply_state(a.out, prev)
    if n:
        print("  restored your entries on %d rows (plus INPUTS and the tracker sheets)" % n)
    recalc(a.out)
    print("Done -> %s" % a.out)


if __name__ == "__main__":
    main()
