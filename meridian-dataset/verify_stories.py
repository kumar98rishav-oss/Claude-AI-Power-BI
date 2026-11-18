"""Checks that the planted stories are present in ./data, and prints the figures a trainer can quote.

    python verify_stories.py
"""
import csv, os
from collections import Counter, defaultdict

D = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
read = lambda n: list(csv.DictReader(open(os.path.join(D, n), encoding="utf-8")))
S, P, I = read("Fact_Sales.csv"), read("Fact_Purchases.csv"), read("Fact_Inventory.csv")
prod = {r["ProductKey"]: r for r in read("Dim_Product.csv")}
cat_of = lambda k: prod[k]["Category"]
half = lambda d: d[:4] + (" H1" if int(d[5:7]) <= 6 else " H2")
pct = lambda a, b: round(a / b * 100, 1) if b else None

rev = defaultdict(float)
for r in S:
    rev[r["OrderDate"][:4]] += float(r["SalesAmount"])
print("1. Growth        revenue 2024 %s, 2025 %s, YoY %s%%" % (round(rev["2024"]), round(rev["2025"]), pct(rev["2025"] - rev["2024"], rev["2024"])))

m = defaultdict(float)
for r in S:
    if cat_of(r["ProductKey"]) in ("PPE", "Diagnostics"):
        m[int(r["OrderDate"][5:7])] += float(r["SalesAmount"])
winter, summer = sum(m[k] for k in (12, 1, 2)), sum(m[k] for k in (6, 7, 8))
print("2. Flu season    PPE + Diagnostics: Dec-Feb revenue is %sx Jun-Aug" % round(winter / summer, 2))

for y in ("2024", "2025"):
    rows = [r for r in S if r["OrderDate"][:4] == y and prod[r["ProductKey"]]["ProductName"].startswith("Infusion Pump")]
    a, c = sum(float(r["SalesAmount"]) for r in rows), sum(float(r["CostAmount"]) for r in rows)
    print("3. Loss-maker    Infusion Pump Tubing Set %s: margin %s%%, average discount %s%%" % (y, pct(a - c, a), round(sum(float(r["DiscountPct"]) for r in rows) / len(rows) * 100, 1)))

ot = defaultdict(lambda: [0, 0])
for r in P:
    if r["ReceivedDate"]:
        k = (r["SupplierKey"] == "6", half(r["OrderDate"]))
        ot[k][1] += 1; ot[k][0] += r["ReceivedDate"] <= r["PromisedDate"]
print("4. Late supplier Pacific Rim Medical on-time %%: %s | all others %s%%" % ({k[1]: pct(*v) for k, v in sorted(ot.items()) if k[0]}, pct(sum(v[0] for k, v in ot.items() if not k[0]), sum(v[1] for k, v in ot.items() if not k[0]))))

so = defaultdict(int)
for r in I:
    if r["OnHandQty"] == "0" and prod[r["ProductKey"]]["SupplierKey"] == "6":
        so[half(r["SnapshotDate"])] += 1
print("5. Stockouts     Pacific Rim products, month-end stockouts by half-year: %s" % dict(sorted(so.items())))

ct = defaultdict(lambda: [0, 0])
for r in S:
    if r["DeliveryDate"] and r["CarrierKey"] == "1":
        ct[half(r["ShipDate"])][1] += 1; ct[half(r["ShipDate"])][0] += r["DeliveryDate"] <= r["RequestedDeliveryDate"]
print("6. Late carrier  Swift Freight on-time delivery %%: %s" % {k: pct(*v) for k, v in sorted(ct.items())})

print("7. Blanks        open POs (no ReceivedDate): %d | undelivered order lines (no DeliveryDate): %d" % (sum(not r["ReceivedDate"] for r in P), sum(not r["DeliveryDate"] for r in S)))
