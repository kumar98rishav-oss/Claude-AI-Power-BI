"""Meridian Medical Supply - training dataset generator.

A fictional US distributor of medical consumables: 6 product categories, 12 suppliers,
5 regional warehouses, 150 healthcare customers, 4 carriers, Jan 2024 - Dec 2025.

Writes clean CSVs to ./data and a deliberately dirty copy to ./data_quality_lab.
Deterministic: the same seed always produces the same files.

    python generate_dataset.py
"""
import csv, math, os, random
from collections import defaultdict
from datetime import date, timedelta

SEED = 20260101
rng = random.Random(SEED)
HERE = os.path.dirname(os.path.abspath(__file__))
OUT, LAB = os.path.join(HERE, "data"), os.path.join(HERE, "data_quality_lab")
START, END = date(2024, 1, 1), date(2025, 12, 31)

# --------------------------------------------------------------------------- dimensions
SUPPLIERS = [  # key, name, country, standard lead time (days), payment terms, tier
    (1, "Harbor Wound Solutions", "USA", 10, "Net 30", "Strategic"),
    (2, "Keystone Medical", "USA", 12, "Net 45", "Preferred"),
    (3, "Summit Surgical Co.", "USA", 14, "Net 30", "Strategic"),
    (4, "Cedar Lab Products", "Canada", 18, "Net 60", "Approved"),
    (5, "BrightCare Diagnostics", "USA", 9, "Net 30", "Strategic"),
    (6, "Pacific Rim Medical", "Malaysia", 28, "Net 60", "Preferred"),
    (7, "Evergreen PPE Ltd", "Vietnam", 30, "Net 60", "Preferred"),
    (8, "Meadow Safety Gear", "USA", 8, "Net 30", "Approved"),
    (9, "Atlas Mobility", "USA", 15, "Net 45", "Strategic"),
    (10, "Northgate Healthcare Supply", "Mexico", 20, "Net 45", "Approved"),
    (11, "Orion Infusion Systems", "USA", 12, "Net 30", "Strategic"),
    (12, "Lumen Diagnostics", "Germany", 25, "Net 60", "Approved"),
]
# category -> (subcategory per product, products, suppliers, price range, unit-qty range, cost ratio)
CATALOG = {
    "Wound Care": ([("Dressings", "Sterile Gauze Pads 4x4 (100 pk)"), ("Bandages", "Adhesive Bandages Assorted (100 pk)"), ("Dressings", "Foam Dressing 6x6 (10 pk)"), ("Dressings", "Hydrocolloid Dressing 4x4 (10 pk)"), ("Bandages", "Elastic Bandage 4 in (12 pk)"), ("Tapes", "Medical Tape 1 in (12 rolls)"), ("Cleansers", "Wound Cleanser Spray 8 oz"), ("Dressings", "Alginate Dressing 4x4 (10 pk)"), ("Dressings", "Non-Adherent Pads 3x4 (50 pk)"), ("Closures", "Skin Closure Strips (50 pk)")], (1, 2), (9, 48), (4, 40), 0.55),
    "Surgical Supplies": ([("Blades", "Surgical Scalpel No.10 (10 pk)"), ("Sutures", "Suture Kit Nylon 3-0 (12 pk)"), ("Drapes", "Sterile Drape 40x40 (20 pk)"), ("Sponges", "Surgical Sponge 4x4 (200 pk)"), ("Closures", "Skin Stapler 35W"), ("Instruments", "Disposable Instrument Tray"), ("Blades", "Surgical Blade No.15 (100 pk)"), ("Instruments", "Disposable Hemostat Forceps (25 pk)"), ("Accessories", "Surgical Skin Marker (10 pk)"), ("Accessories", "Specimen Container 4 oz (100 pk)")], (3, 2), (18, 140), (2, 20), 0.52),
    "Diagnostics": ([("Rapid Tests", "Rapid Flu A/B Test (25 pk)"), ("Rapid Tests", "COVID-19 Antigen Test (25 pk)"), ("Test Strips", "Glucose Test Strips (50 pk)"), ("Test Strips", "Urinalysis Strips (100 pk)"), ("Devices", "Digital Thermometer"), ("Devices", "Fingertip Pulse Oximeter"), ("Devices", "Adult Blood Pressure Cuff"), ("Rapid Tests", "Pregnancy Test Cassette (50 pk)"), ("Consumables", "Safety Lancets 28G (200 pk)"), ("Rapid Tests", "Strep A Rapid Test (25 pk)")], (5, 12), (14, 120), (2, 30), 0.50),
    "PPE": ([("Gloves", "Nitrile Exam Gloves M (100 pk)"), ("Gloves", "Nitrile Exam Gloves L (100 pk)"), ("Respiratory", "N95 Respirator (20 pk)"), ("Respiratory", "Surgical Mask Level 3 (50 pk)"), ("Apparel", "Isolation Gown (10 pk)"), ("Apparel", "Face Shield (24 pk)"), ("Apparel", "Shoe Covers (100 pk)"), ("Apparel", "Bouffant Cap (100 pk)"), ("Hygiene", "Hand Sanitizer 500 ml"), ("Hygiene", "Disinfectant Wipes (160 ct)")], (7, 8), (6, 45), (10, 120), 0.62),
    "Mobility Aids": ([("Walking Aids", "Folding Walker"), ("Wheelchairs", "Standard Wheelchair 18 in"), ("Walking Aids", "Quad Cane"), ("Walking Aids", "Forearm Crutches (pair)"), ("Walking Aids", "Rollator 4-Wheel"), ("Bathroom Safety", "Transfer Bench"), ("Bathroom Safety", "Raised Toilet Seat"), ("Bed Safety", "Bed Assist Rail"), ("Wheelchairs", "Knee Scooter"), ("Bathroom Safety", "Shower Chair")], (9, 10), (35, 260), (1, 6), 0.58),
    "Infusion": ([("IV Access", "IV Start Kit"), ("IV Access", "IV Catheter 22G (50 pk)"), ("IV Sets", "Infusion Pump Tubing Set (20 pk)"), ("Flushes", "Saline Flush Syringe 10 ml (30 pk)"), ("IV Sets", "Extension Set 7 in (50 pk)"), ("IV Sets", "Primary IV Set 60 drop (50 pk)"), ("Syringes", "Luer Lock Syringe 10 ml (100 pk)"), ("Needles", "Hypodermic Needle 21G (100 pk)"), ("Sharps", "Sharps Container 2 gal"), ("IV Access", "Latex-Free Tourniquet (25 pk)")], (11, 6), (12, 160), (2, 25), 0.60),
}
STORY_PRODUCT = "Infusion Pump Tubing Set (20 pk)"   # discount creep -> loss-making in 2025
STORY_SUPPLIER = 6                                  # Pacific Rim Medical: late from Jul 2025
STORY_CARRIER = 1                                   # Swift Freight: late deliveries in H2 2025
FLU_CATEGORIES = {"PPE", "Diagnostics"}

WAREHOUSES = [  # key, code, name, city, state, region, capacity units
    (1, "ATL", "Atlanta DC", "Atlanta", "GA", "South", 180000),
    (2, "DAL", "Dallas DC", "Dallas", "TX", "Central", 160000),
    (3, "CHI", "Chicago DC", "Chicago", "IL", "Midwest", 170000),
    (4, "EWR", "Newark DC", "Newark", "NJ", "Northeast", 200000),
    (5, "RNO", "Reno DC", "Reno", "NV", "West", 150000),
]
REGION_STATES = {"South": ["GA", "FL", "NC", "TN", "AL"], "Central": ["TX", "OK", "KS", "LA"], "Midwest": ["IL", "OH", "MI", "WI", "MN"], "Northeast": ["NJ", "NY", "PA", "MA", "CT"], "West": ["NV", "CA", "AZ", "OR", "WA"]}
CITIES = {"GA": ["Savannah", "Macon"], "FL": ["Tampa", "Orlando"], "NC": ["Raleigh", "Durham"], "TN": ["Nashville", "Knoxville"], "AL": ["Mobile", "Huntsville"], "TX": ["Austin", "Houston"], "OK": ["Tulsa", "Norman"], "KS": ["Wichita", "Topeka"], "LA": ["Baton Rouge", "Shreveport"], "IL": ["Peoria", "Springfield"], "OH": ["Columbus", "Dayton"], "MI": ["Lansing", "Grand Rapids"], "WI": ["Madison", "Green Bay"], "MN": ["Rochester", "Duluth"], "NJ": ["Trenton", "Camden"], "NY": ["Albany", "Buffalo"], "PA": ["Allentown", "Erie"], "MA": ["Worcester", "Lowell"], "CT": ["Hartford", "New Haven"], "NV": ["Las Vegas", "Henderson"], "CA": ["Fresno", "Sacramento"], "AZ": ["Tucson", "Mesa"], "OR": ["Salem", "Eugene"], "WA": ["Spokane", "Tacoma"]}
CARRIERS = [(1, "Swift Freight", "Ground", 3), (2, "BlueLine Express", "Express", 1), (3, "Continental LTL", "LTL", 4), (4, "Metro Courier", "Same Day", 0)]
PLACES = ["Lakeview", "Riverside", "Cedar Ridge", "Maple Grove", "Harbor Point", "Stonebridge", "Willow Creek", "Fairhaven", "Pine Valley", "Brookside", "Silver Lake", "Oak Hollow", "Granite Falls", "Meadowbrook", "Sunset Hills", "Clearwater", "Northfield", "Eastgate", "Westbrook", "Southport", "Bayview", "Highland Park", "Kingsley", "Ashford", "Bellmont", "Crestview", "Dover Plains", "Elmwood", "Foxborough", "Glenwood", "Hartwell", "Ironwood", "Juniper", "Kensington", "Linden", "Marlow", "Newbury", "Oakridge", "Prescott", "Quincy Hill", "Rosewood", "Sheffield", "Thornbury", "Upton", "Vale", "Whitmore", "Yorkshire", "Alder Bay", "Birchwood", "Copperfield"]
CUSTOMER_TYPES = [("Hospital", "Medical Center", 0.22), ("Clinic", "Family Clinic", 0.34), ("Pharmacy", "Pharmacy", 0.18), ("Nursing Home", "Care Home", 0.16), ("Surgery Center", "Surgery Center", 0.10)]


def daterange(a, b):
    d = a
    while d <= b:
        yield d
        d += timedelta(days=1)


def month_end(d):
    nxt = date(d.year + (d.month == 12), d.month % 12 + 1, 1)
    return nxt - timedelta(days=1)


def add_business_days(d, n):
    while n > 0:
        d += timedelta(days=1)
        if d.weekday() < 5:
            n -= 1
    return d


def write(path, header, rows):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(header)
        for r in rows:
            w.writerow(["" if v is None else (v.isoformat() if isinstance(v, date) else v) for v in r])


def main():
    # ---- Dim_Date (fiscal year starts 1 July: FY2025 = Jul 2024 - Jun 2025)
    dim_date = []
    for d in daterange(START, END):
        fy = d.year + (d.month >= 7)
        fm = (d.month - 7) % 12 + 1
        dim_date.append((d, d.year, f"Q{(d.month - 1) // 3 + 1}", (d.month - 1) // 3 + 1, d.month, d.strftime("%b"), d.strftime("%b %Y"), d.year * 100 + d.month,
                         d.isoweekday(), d.strftime("%a"), "Yes" if d.weekday() >= 5 else "No", f"FY{fy}", f"FQ{(fm - 1) // 3 + 1}", fm))

    # ---- Dim_Product
    products, pkey = [], 0
    for cat, (items, sups, (plo, phi), (qlo, qhi), cost_ratio) in CATALOG.items():
        for i, (sub, name) in enumerate(items):
            pkey += 1
            price = round(rng.uniform(plo, phi), 2)
            cost = round(price * rng.uniform(cost_ratio - 0.06, cost_ratio + 0.06), 2)
            if name == STORY_PRODUCT:
                cost = round(price * 0.74, 2)
            products.append(dict(key=pkey, sku=f"MS-{1000 + pkey}", name=name, cat=cat, sub=sub, supplier=sups[i % 2], cost=cost, price=price, qlo=qlo, qhi=qhi,
                                 status="Discontinued" if pkey in (17, 49) else "Active", pop=rng.uniform(0.5, 1.6)))

    # ---- Dim_Customer
    customers, used = [], set()
    region_wh = {w[5]: w[0] for w in WAREHOUSES}
    for ck in range(1, 151):
        t = rng.choices(CUSTOMER_TYPES, weights=[c[2] for c in CUSTOMER_TYPES])[0]
        region = rng.choices(list(REGION_STATES), weights=[24, 18, 20, 24, 14])[0]
        st = rng.choice(REGION_STATES[region])
        nm = f"{rng.choice(PLACES)} {t[1]}"
        for _ in range(40):
            if nm not in used:
                break
            nm = f"{rng.choice(PLACES)} {t[1]}"
        else:
            nm = f"{rng.choice(PLACES)} {t[1]} of {rng.choice(CITIES[st])}"
        used.add(nm)
        customers.append(dict(key=ck, name=nm, type=t[0], city=rng.choice(CITIES[st]), state=st, region=region, since=START - timedelta(days=rng.randint(60, 2500)), weight=rng.paretovariate(1.6) * (2.2 if t[0] == "Hospital" else 1)))

    # ---- Fact_Sales
    sales, so = [], 500000
    disc_levels = [0, 0, 0, 0.05, 0.05, 0.10, 0.15]
    cust_w = [c["weight"] for c in customers]
    by_cat = defaultdict(list)
    for p in products:
        by_cat[p["cat"]].append(p)
    season = [1.22, 1.14, 1.0, 0.93, 0.9, 0.88, 0.9, 0.95, 1.0, 1.07, 1.16, 1.28]
    for d in daterange(START, END):
        growth = 1 + 0.20 * ((d - START).days / 730)
        lam = (12.5 if d.weekday() < 5 else 2.5) * growth * season[d.month - 1]
        n_orders = sum(1 for _ in range(int(lam * 3)) if rng.random() < 1 / 3)
        for _ in range(n_orders):
            so += 1
            cust = rng.choices(customers, weights=cust_w)[0]
            wh = region_wh[cust["region"]] if rng.random() < 0.9 else rng.randint(1, 5)
            carrier = rng.choices(CARRIERS, weights=[50, 20, 18, 12] if cust["type"] != "Hospital" else [38, 34, 16, 12])[0]
            ship = add_business_days(d, rng.choice([0, 1, 1, 2])) if d.weekday() < 5 else add_business_days(d, 1)
            requested = ship + timedelta(days=max(carrier[3], 1) + 1)
            late_p = 0.08
            if carrier[0] == STORY_CARRIER and ship >= date(2025, 7, 1):
                late_p = 0.34
            delivered = ship + timedelta(days=carrier[3] + (rng.randint(1, 4) if rng.random() < late_p else rng.choice([0, 0, 0, 1]) * (carrier[3] > 1)))
            if ship > END:
                ship = None
            if delivered > END or ship is None:
                delivered = None
            weights = []
            for cat in CATALOG:
                w = 1.0
                if cat in FLU_CATEGORIES:
                    w *= 0.7 + 0.6 * season[d.month - 1] ** 2
                if cat == "Mobility Aids":
                    w *= 0.45 if cust["type"] in ("Nursing Home", "Pharmacy") else 0.12
                if cat == "Surgical Supplies" and cust["type"] in ("Hospital", "Surgery Center"):
                    w *= 1.8
                weights.append(w)
            lines = rng.choices([1, 2, 3, 4, 5], weights=[30, 30, 20, 12, 8])[0]
            chosen = set()
            for ln in range(1, lines + 1):
                cat = rng.choices(list(CATALOG), weights=weights)[0]
                p = rng.choices(by_cat[cat], weights=[x["pop"] for x in by_cat[cat]])[0]
                if p["key"] in chosen or (p["status"] == "Discontinued" and d >= date(2025, 4, 1)):
                    continue
                chosen.add(p["key"])
                qty = rng.randint(p["qlo"], p["qhi"]) * (2 if cust["type"] == "Hospital" else 1)
                price = round(p["price"] * (1.03 if d.year == 2025 else 1), 2)
                unit_cost = round(p["cost"] * (1.045 if d.year == 2025 else 1), 2)
                disc = rng.choice(disc_levels) + (0.05 if cust["type"] == "Hospital" and rng.random() < 0.4 else 0)
                if p["name"] == STORY_PRODUCT:
                    disc = (0.10 if d.year == 2024 else 0.22) + rng.choice([0, 0.05, 0.08]) + (0.06 if d >= date(2025, 7, 1) else 0)
                disc = round(min(disc, 0.40), 2)
                amount = round(qty * price * (1 - disc), 2)
                sales.append([so, len(chosen), d, requested, ship, delivered, cust["key"], p["key"], wh, carrier[0], qty, price, disc, amount, round(qty * unit_cost, 2)])
    sales.sort(key=lambda r: (r[2], r[0], r[1]))
    for i, r in enumerate(sales, 1):
        r.insert(0, i)  # SalesLineKey

    # ---- replenishment simulation -> Fact_Purchases + Fact_Inventory
    months = sorted({(d.year, d.month) for d in daterange(START, END)})
    demand = defaultdict(float)  # (prod, wh, (y, m)) -> qty
    for r in sales:
        demand[(r[8], r[9], (r[3].year, r[3].month))] += r[11]
    sup_lt = {s[0]: s[3] for s in SUPPLIERS}
    purchases, inventory, po_no = [], [], 700000
    po_groups = defaultdict(list)   # (supplier, wh, orderdate) -> lines
    receipts = defaultdict(float)   # (prod, wh, (y, m)) -> qty received
    open_qty = defaultdict(float)
    pending = []                    # (received_date, prod, wh, qty)
    onhand = {}
    for p in products:
        for w in WAREHOUSES:
            avg = sum(demand[(p["key"], w[0], m)] for m in months[:3]) / 3
            onhand[(p["key"], w[0])] = round(avg * 2.6)
    for mi, (y, m) in enumerate(months):
        first = date(y, m, 1)
        for p in products:
            if p["status"] == "Discontinued" and first >= date(2025, 3, 1):
                continue
            lt_days = sup_lt[p["supplier"]]
            for w in WAREHOUSES:
                k = (p["key"], w[0])
                hist = [demand[(p["key"], w[0], mm)] for mm in months[max(0, mi - 3):mi]] or [demand[(p["key"], w[0], months[0])]]
                fc = max(sum(hist) / len(hist), 0.5)
                safety = math.ceil(fc * 0.5)
                rop = math.ceil(fc * (lt_days / 30 + 0.5) + safety)
                position = onhand[k] + open_qty[k]
                if position < rop:
                    qty = max(int(math.ceil((rop + fc * 1.5 - position) / 10.0)) * 10, 10)
                    od = add_business_days(first - timedelta(days=1), rng.randint(1, 4))
                    promised = od + timedelta(days=lt_days)
                    late_p, late_rng = 0.07, (1, 5)
                    if p["supplier"] == STORY_SUPPLIER and od >= date(2025, 6, 15):
                        late_p, late_rng = 0.70, (14, 35)
                    received = promised + timedelta(days=(rng.randint(*late_rng) if rng.random() < late_p else -rng.choice([0, 0, 1, 2])))
                    rq = qty if not (p["supplier"] == 2 and rng.random() < 0.18) else int(qty * rng.uniform(0.7, 0.95))
                    if received > END:
                        received, rq = None, 0
                    po_groups[(p["supplier"], w[0], od)].append([p["key"], qty, rq, round(p["cost"] * (1.045 if od.year == 2025 else 1) * 0.98, 2), promised, received])
                    open_qty[k] += qty
                    if received:
                        pending.append((received, p["key"], w[0], rq, qty))
        # receipts landing this month
        me = month_end(first)
        for item in [x for x in pending if x[0] <= me]:
            pending.remove(item)
            _, pk, wk, rq, qty = item
            onhand[(pk, wk)] += rq
            open_qty[(pk, wk)] -= qty
        for p in products:
            for w in WAREHOUSES:
                k = (p["key"], w[0])
                dem = demand[(p["key"], w[0], (y, m))]
                onhand[k] = max(round(onhand[k] - dem), 0)
                hist = [demand[(p["key"], w[0], mm)] for mm in months[max(0, mi - 2):mi + 1]]
                fc = max(sum(hist) / len(hist), 0.5)
                safety = math.ceil(fc * 0.5)
                rop = math.ceil(fc * (sup_lt[p["supplier"]] / 30 + 0.5) + safety)
                inventory.append([me, p["key"], w[0], onhand[k], int(max(open_qty[k], 0)), safety, rop, round(onhand[k] * p["cost"], 2)])
    for (sup, wk, od), lines in sorted(po_groups.items(), key=lambda x: (x[0][2], x[0][0], x[0][1])):
        po_no += 1
        for ln, (pk, qty, rq, uc, promised, received) in enumerate(lines, 1):
            purchases.append([f"PO-{po_no}", ln, od, promised, received, sup, pk, wk, qty, rq if received else None, uc, round(qty * uc, 2)])

    # ---- write clean data
    write(f"{OUT}/Dim_Date.csv", ["Date", "Year", "Quarter", "QuarterNo", "MonthNo", "Month", "MonthYear", "YearMonth", "WeekdayNo", "Weekday", "IsWeekend", "FiscalYear", "FiscalQuarter", "FiscalMonthNo"], dim_date)
    write(f"{OUT}/Dim_Product.csv", ["ProductKey", "SKU", "ProductName", "Category", "Subcategory", "SupplierKey", "UnitCost", "ListPrice", "Status"], [[p["key"], p["sku"], p["name"], p["cat"], p["sub"], p["supplier"], p["cost"], p["price"], p["status"]] for p in products])
    write(f"{OUT}/Dim_Supplier.csv", ["SupplierKey", "SupplierName", "Country", "StandardLeadTimeDays", "PaymentTerms", "SupplierTier"], SUPPLIERS)
    write(f"{OUT}/Dim_Warehouse.csv", ["WarehouseKey", "WarehouseCode", "WarehouseName", "City", "State", "Region", "CapacityUnits"], WAREHOUSES)
    write(f"{OUT}/Dim_Customer.csv", ["CustomerKey", "CustomerName", "CustomerType", "City", "State", "Region", "CustomerSince"], [[c["key"], c["name"], c["type"], c["city"], c["state"], c["region"], c["since"]] for c in customers])
    write(f"{OUT}/Dim_Carrier.csv", ["CarrierKey", "CarrierName", "ServiceMode", "StandardTransitDays"], CARRIERS)
    sales_hdr = ["SalesLineKey", "SalesOrderID", "LineNo", "OrderDate", "RequestedDeliveryDate", "ShipDate", "DeliveryDate", "CustomerKey", "ProductKey", "WarehouseKey", "CarrierKey", "Quantity", "UnitPrice", "DiscountPct", "SalesAmount", "CostAmount"]
    write(f"{OUT}/Fact_Sales.csv", sales_hdr, sales)
    po_hdr = ["PONumber", "POLineNo", "OrderDate", "PromisedDate", "ReceivedDate", "SupplierKey", "ProductKey", "WarehouseKey", "QtyOrdered", "QtyReceived", "UnitCost", "LineCost"]
    write(f"{OUT}/Fact_Purchases.csv", po_hdr, purchases)
    write(f"{OUT}/Fact_Inventory.csv", ["SnapshotDate", "ProductKey", "WarehouseKey", "OnHandQty", "OnOrderQty", "SafetyStockQty", "ReorderPointQty", "InventoryValue"], inventory)

    # ---- data quality lab: the same data with documented, deliberate defects
    dirty = [list(r) for r in sales]
    N = len(sales)
    at = lambda f: int(N * f)
    nxt = max(r[0] for r in dirty)
    issues = []
    for i in (at(0.06), at(0.41), at(0.77)):                       # orphan product key
        nxt += 1; r = list(dirty[i]); r[0] = nxt; r[8] = 999; dirty.append(r)
    issues.append(("Fact_Sales", "Orphan key", "3 rows with ProductKey 999, which is not in Dim_Product", "LEFT JOIN Dim_Product; ProductKey 999"))
    for i in (at(0.21), at(0.58)):                             # exact duplicate lines
        dirty.append(list(dirty[i]))
    issues.append(("Fact_Sales", "Duplicate rows", "2 order lines appear twice with the same SalesLineKey", "GROUP BY SalesLineKey HAVING COUNT(*) > 1"))
    ng = at(0.45); dirty[ng][11] = -dirty[ng][11]; dirty[ng][14] = -dirty[ng][14]; dirty[ng][15] = -dirty[ng][15]
    issues.append(("Fact_Sales", "Negative quantity", f"SalesLineKey {dirty[ng][0]} has a negative Quantity with no return flag", "Quantity < 0"))
    for i in (at(0.015), at(0.33), at(0.69), at(0.88)):
        dirty[i][7] = None
    issues.append(("Fact_Sales", "Missing customer", "4 rows with a blank CustomerKey", "CustomerKey IS NULL"))
    tx = at(0.27); dirty[tx][13] = "5%"
    issues.append(("Fact_Sales", "Text in a numeric column", f"SalesLineKey {dirty[tx][0]} has DiscountPct stored as '5%'", "Power Query type-change error on DiscountPct"))
    for i in (at(0.12), at(0.12) + 1, at(0.12) + 2):
        d = dirty[i][3]; dirty[i][3] = d.strftime("%m/%d/%Y")
    issues.append(("Fact_Sales", "Mixed date formats", "3 OrderDate values written as MM/DD/YYYY instead of ISO", "Power Query date parse errors"))
    write(f"{LAB}/Fact_Sales.csv", sales_hdr, dirty)
    cust_dirty = [[c["key"], c["name"], c["type"], c["city"], c["state"], c["region"], c["since"]] for c in customers]
    cust_dirty[10][5] = cust_dirty[10][5].lower(); cust_dirty[25][5] = cust_dirty[25][5] + " "; cust_dirty[40][1] = "  " + cust_dirty[40][1]
    issues.append(("Dim_Customer", "Inconsistent text", "Region in lower case on one row, a trailing space on another, leading spaces in one CustomerName", "DISTINCT Region returns extra values"))
    cust_dirty.append(list(cust_dirty[60])); cust_dirty[-1][1] = cust_dirty[-1][1] + " Inc"
    issues.append(("Dim_Customer", "Duplicate key", f"CustomerKey {cust_dirty[60][0]} appears twice with different names, which breaks a one-to-many relationship", "GROUP BY CustomerKey HAVING COUNT(*) > 1"))
    write(f"{LAB}/Dim_Customer.csv", ["CustomerKey", "CustomerName", "CustomerType", "City", "State", "Region", "CustomerSince"], cust_dirty)
    write(f"{LAB}/ISSUES.csv", ["Table", "Issue", "Detail", "How to detect"], issues)

    # ---- summary for the docs
    rev = defaultdict(float)
    for r in sales:
        rev[r[3].year] += r[14]
    print("rows:", {"Fact_Sales": len(sales), "Fact_Purchases": len(purchases), "Fact_Inventory": len(inventory), "Dim_Customer": len(customers), "Dim_Product": len(products)})
    print("revenue by year:", {y: round(v) for y, v in rev.items()})


if __name__ == "__main__":
    main()
