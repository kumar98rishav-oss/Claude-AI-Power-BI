"""Builds the MeridianSupplyChain PBIP: a TMDL semantic model over ./data and a three-page PBIR report.

    python build_pbip.py [--data-folder "C:\\path\\to\\data\\"]

The report folder was first created with `powerbi-report-author scaffold`; this script owns the
semantic model and the pages/visuals. IDs come from uuid5, so a rebuild produces identical files.
"""
import json, os, shutil, sys, uuid

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.join(HERE, "pbip")
NAME = "MeridianSupplyChain"
SM = os.path.join(ROOT, f"{NAME}.SemanticModel")
RP = os.path.join(ROOT, f"{NAME}.Report")
DATA = sys.argv[sys.argv.index("--data-folder") + 1] if "--data-folder" in sys.argv else os.path.join(HERE, "data") + "\\"
NS = uuid.UUID("6f1c2a3b-0d4e-4f5a-9b8c-7d6e5f4a3b2c")
tag = lambda *k: str(uuid.uuid5(NS, "/".join(k)))
hexid = lambda *k: uuid.uuid5(NS, "v/" + "/".join(k)).hex[:20]


def write(path, text):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8", newline="\n") as f:  # UTF-8 without BOM
        f.write(text)


def q(name):  # TMDL quoting for names with spaces or symbols
    return f"'{name}'" if not name.replace("_", "").isalnum() else name


# --------------------------------------------------------------------------- semantic model
I, D, S, DT = "int64", "double", "string", "dateTime"
M_TYPE = {I: "Int64.Type", D: "type number", S: "type text", DT: "type date"}
TABLES = {
    "Dim_Date": dict(desc="One row per day, Jan 2024 - Dec 2025. Fiscal year starts 1 July.", cols=[
        ("Date", DT, dict(key=True)), ("Year", I, {}), ("Quarter", S, dict(sort="QuarterNo")), ("QuarterNo", I, dict(hide=True)),
        ("MonthNo", I, dict(hide=True)), ("Month", S, dict(sort="MonthNo")), ("MonthYear", S, dict(sort="YearMonth")), ("YearMonth", I, dict(hide=True)),
        ("WeekdayNo", I, dict(hide=True)), ("Weekday", S, dict(sort="WeekdayNo")), ("IsWeekend", S, {}), ("FiscalYear", S, {}), ("FiscalQuarter", S, {}), ("FiscalMonthNo", I, dict(hide=True))],
        hier=[("Calendar", ["Year", "Quarter", "Month"])], date_table=True),
    "Dim_Product": dict(desc="60 products in 6 categories. Each product has one supplier.", cols=[
        ("ProductKey", I, dict(hide=True)), ("SKU", S, {}), ("ProductName", S, {}), ("Category", S, {}), ("Subcategory", S, {}),
        ("SupplierKey", I, dict(hide=True)), ("UnitCost", D, dict(fmt="\\$#,##0.00")), ("ListPrice", D, dict(fmt="\\$#,##0.00")), ("Status", S, {})],
        hier=[("Product Hierarchy", ["Category", "Subcategory", "ProductName"])]),
    "Dim_Supplier": dict(desc="12 suppliers. Reached from purchases through Dim_Product (snowflake).", cols=[
        ("SupplierKey", I, dict(hide=True)), ("SupplierName", S, {}), ("Country", S, {}), ("StandardLeadTimeDays", I, {}), ("PaymentTerms", S, {}), ("SupplierTier", S, {})]),
    "Dim_Warehouse": dict(desc="5 regional distribution centres.", cols=[
        ("WarehouseKey", I, dict(hide=True)), ("WarehouseCode", S, {}), ("WarehouseName", S, {}), ("City", S, {}), ("State", S, {}), ("Region", S, {}), ("CapacityUnits", I, dict(fmt="#,##0"))],
        hier=[("Network", ["Region", "WarehouseName"])]),
    "Dim_Customer": dict(desc="150 healthcare customers.", cols=[
        ("CustomerKey", I, dict(hide=True)), ("CustomerName", S, {}), ("CustomerType", S, {}), ("City", S, {}), ("State", S, {}), ("Region", S, {}), ("CustomerSince", DT, {})]),
    "Dim_Carrier": dict(desc="4 outbound carriers.", cols=[
        ("CarrierKey", I, dict(hide=True)), ("CarrierName", S, {}), ("ServiceMode", S, {}), ("StandardTransitDays", I, {})]),
    "Fact_Sales": dict(desc="One row per sales order line. Blank ShipDate/DeliveryDate = not yet shipped/delivered.", nulls=["ShipDate", "DeliveryDate"], cols=[
        ("SalesLineKey", I, dict(hide=True)), ("SalesOrderID", I, {}), ("LineNo", I, dict(hide=True)), ("OrderDate", DT, {}), ("RequestedDeliveryDate", DT, {}),
        ("ShipDate", DT, {}), ("DeliveryDate", DT, {}), ("CustomerKey", I, dict(hide=True)), ("ProductKey", I, dict(hide=True)), ("WarehouseKey", I, dict(hide=True)),
        ("CarrierKey", I, dict(hide=True)), ("Quantity", I, dict(hide=True)), ("UnitPrice", D, dict(hide=True)), ("DiscountPct", D, dict(hide=True, fmt="0%")),
        ("SalesAmount", D, dict(hide=True)), ("CostAmount", D, dict(hide=True))]),
    "Fact_Purchases": dict(desc="One row per purchase order line. Blank ReceivedDate = still open.", nulls=["ReceivedDate", "QtyReceived"], cols=[
        ("PONumber", S, {}), ("POLineNo", I, dict(hide=True)), ("OrderDate", DT, {}), ("PromisedDate", DT, {}), ("ReceivedDate", DT, {}),
        ("SupplierKey", I, dict(hide=True)), ("ProductKey", I, dict(hide=True)), ("WarehouseKey", I, dict(hide=True)),
        ("QtyOrdered", I, dict(hide=True)), ("QtyReceived", I, dict(hide=True)), ("UnitCost", D, dict(hide=True)), ("LineCost", D, dict(hide=True))],
        calc=[("Days Late", I, "IF ( ISBLANK ( Fact_Purchases[ReceivedDate] ), BLANK (), MAX ( 0, DATEDIFF ( Fact_Purchases[PromisedDate], Fact_Purchases[ReceivedDate], DAY ) ) )")]),
    "Fact_Inventory": dict(desc="Month-end stock snapshot per product and warehouse. Semi-additive: never sum across dates.", cols=[
        ("SnapshotDate", DT, {}), ("ProductKey", I, dict(hide=True)), ("WarehouseKey", I, dict(hide=True)), ("OnHandQty", I, dict(hide=True)),
        ("OnOrderQty", I, dict(hide=True)), ("SafetyStockQty", I, dict(hide=True)), ("ReorderPointQty", I, dict(hide=True)), ("InventoryValue", D, dict(hide=True))]),
}

LAST_SNAPSHOT = "LASTNONBLANK ( Dim_Date[Date], CALCULATE ( COUNTROWS ( Fact_Inventory ) ) )"
MEASURES = [  # folder, name, expression (str or list of lines), format, description
    ("01 Sales", "Revenue", "SUM ( Fact_Sales[SalesAmount] )", "\\$#,##0", "Net sales after discount."),
    ("01 Sales", "Cost", "SUM ( Fact_Sales[CostAmount] )", "\\$#,##0", None),
    ("01 Sales", "Gross Margin", "[Revenue] - [Cost]", "\\$#,##0", None),
    ("01 Sales", "Gross Margin %", "DIVIDE ( [Gross Margin], [Revenue] )", "0.0%", "Gross margin as a share of revenue."),
    ("01 Sales", "Units Sold", "SUM ( Fact_Sales[Quantity] )", "#,##0", None),
    ("01 Sales", "Orders", "DISTINCTCOUNT ( Fact_Sales[SalesOrderID] )", "#,##0", None),
    ("01 Sales", "Avg Order Value", "DIVIDE ( [Revenue], [Orders] )", "\\$#,##0", None),
    ("01 Sales", "Gross Sales", "SUMX ( Fact_Sales, Fact_Sales[Quantity] * Fact_Sales[UnitPrice] )", "\\$#,##0", "Revenue before discount."),
    ("01 Sales", "Avg Discount %", "DIVIDE ( [Gross Sales] - [Revenue], [Gross Sales] )", "0.0%", "Discount given, weighted by sales value."),
    ("02 Time Intelligence", "Revenue PY", "CALCULATE ( [Revenue], SAMEPERIODLASTYEAR ( Dim_Date[Date] ) )", "\\$#,##0", "Revenue for the same period last year."),
    ("02 Time Intelligence", "Revenue YoY %", "IF ( HASONEVALUE ( Dim_Date[Year] ), DIVIDE ( [Revenue] - [Revenue PY], [Revenue PY] ) )", "0.0%", "Blank unless one year is in context: across all years, last year is only part of the data."),
    ("02 Time Intelligence", "Revenue YTD", "TOTALYTD ( [Revenue], Dim_Date[Date] )", "\\$#,##0", "Calendar year to date."),
    ("02 Time Intelligence", "Revenue FYTD", 'TOTALYTD ( [Revenue], Dim_Date[Date], "6/30" )', "\\$#,##0", "Fiscal year to date. The fiscal year ends on 30 June."),
    ("02 Time Intelligence", "Revenue by Ship Date", "CALCULATE ( [Revenue], USERELATIONSHIP ( Fact_Sales[ShipDate], Dim_Date[Date] ) )", "\\$#,##0", "Uses the inactive ShipDate relationship."),
    ("03 Delivery", "Lines Delivered", "CALCULATE ( COUNTROWS ( Fact_Sales ), NOT ISBLANK ( Fact_Sales[DeliveryDate] ) )", "#,##0", None),
    ("03 Delivery", "On-Time Delivery %", [
        "VAR OnTime =", "    CALCULATE (", "        COUNTROWS ( Fact_Sales ),",
        "        FILTER ( Fact_Sales, NOT ISBLANK ( Fact_Sales[DeliveryDate] ) && Fact_Sales[DeliveryDate] <= Fact_Sales[RequestedDeliveryDate] )", "    )",
        "RETURN", "    DIVIDE ( OnTime, [Lines Delivered] )"], "0.0%", "Delivered lines that arrived on or before the requested date."),
    ("04 Procurement", "PO Lines", "COUNTROWS ( Fact_Purchases )", "#,##0", None),
    ("04 Procurement", "Purchase Spend", "SUM ( Fact_Purchases[LineCost] )", "\\$#,##0", None),
    ("04 Procurement", "Lines Received", "CALCULATE ( COUNTROWS ( Fact_Purchases ), NOT ISBLANK ( Fact_Purchases[ReceivedDate] ) )", "#,##0", None),
    ("04 Procurement", "Open PO Lines", "CALCULATE ( COUNTROWS ( Fact_Purchases ), ISBLANK ( Fact_Purchases[ReceivedDate] ) )", "#,##0", "Ordered but not yet received."),
    ("04 Procurement", "Supplier On-Time %", [
        "VAR OnTime =", "    CALCULATE (", "        COUNTROWS ( Fact_Purchases ),",
        "        FILTER ( Fact_Purchases, NOT ISBLANK ( Fact_Purchases[ReceivedDate] ) && Fact_Purchases[ReceivedDate] <= Fact_Purchases[PromisedDate] )", "    )",
        "RETURN", "    DIVIDE ( OnTime, [Lines Received] )"], "0.0%", "Received lines that arrived on or before the promised date."),
    ("04 Procurement", "In-Full %", [
        "VAR InFull =", "    CALCULATE (", "        COUNTROWS ( Fact_Purchases ),",
        "        FILTER ( Fact_Purchases, NOT ISBLANK ( Fact_Purchases[ReceivedDate] ) && Fact_Purchases[QtyReceived] >= Fact_Purchases[QtyOrdered] )", "    )",
        "RETURN", "    DIVIDE ( InFull, [Lines Received] )"], "0.0%", "Received lines delivered in the full quantity."),
    ("04 Procurement", "OTIF %", [
        "VAR Otif =", "    CALCULATE (", "        COUNTROWS ( Fact_Purchases ),", "        FILTER (", "            Fact_Purchases,",
        "            NOT ISBLANK ( Fact_Purchases[ReceivedDate] )", "                && Fact_Purchases[ReceivedDate] <= Fact_Purchases[PromisedDate]",
        "                && Fact_Purchases[QtyReceived] >= Fact_Purchases[QtyOrdered]", "        )", "    )",
        "RETURN", "    DIVIDE ( Otif, [Lines Received] )"], "0.0%", "On time and in full."),
    ("04 Procurement", "Avg Lead Time (Days)", "AVERAGEX ( FILTER ( Fact_Purchases, NOT ISBLANK ( Fact_Purchases[ReceivedDate] ) ), DATEDIFF ( Fact_Purchases[OrderDate], Fact_Purchases[ReceivedDate], DAY ) )", "0.0", None),
    ("04 Procurement", "Avg Days Late", "AVERAGE ( Fact_Purchases[Days Late] )", "0.0", "Uses the Days Late calculated column."),
    ("04 Procurement", "Spend by Received Date", "CALCULATE ( [Purchase Spend], USERELATIONSHIP ( Fact_Purchases[ReceivedDate], Dim_Date[Date] ) )", "\\$#,##0", "Uses the inactive ReceivedDate relationship."),
    ("05 Inventory", "On Hand Units", f"CALCULATE ( SUM ( Fact_Inventory[OnHandQty] ), {LAST_SNAPSHOT} )", "#,##0", "Stock at the last snapshot in the period. Semi-additive."),
    ("05 Inventory", "Inventory Value", f"CALCULATE ( SUM ( Fact_Inventory[InventoryValue] ), {LAST_SNAPSHOT} )", "\\$#,##0", "Stock value at the last snapshot in the period."),
    ("05 Inventory", "Stockouts", f"CALCULATE ( COUNTROWS ( Fact_Inventory ), Fact_Inventory[OnHandQty] = 0, {LAST_SNAPSHOT} )", "#,##0", "Product-warehouse pairs with zero stock at the last snapshot."),
    ("05 Inventory", "Below Safety Stock", [
        "CALCULATE (", "    COUNTROWS ( Fact_Inventory ),", "    FILTER ( Fact_Inventory, Fact_Inventory[OnHandQty] < Fact_Inventory[SafetyStockQty] ),",
        f"    {LAST_SNAPSHOT}", ")"], "#,##0", "Product-warehouse pairs below safety stock at the last snapshot."),
]
RELS = [  # name, from, to, active
    ("Sales_OrderDate", "Fact_Sales.OrderDate", "Dim_Date.Date", True), ("Sales_ShipDate", "Fact_Sales.ShipDate", "Dim_Date.Date", False),
    ("Sales_Customer", "Fact_Sales.CustomerKey", "Dim_Customer.CustomerKey", True), ("Sales_Product", "Fact_Sales.ProductKey", "Dim_Product.ProductKey", True),
    ("Sales_Warehouse", "Fact_Sales.WarehouseKey", "Dim_Warehouse.WarehouseKey", True), ("Sales_Carrier", "Fact_Sales.CarrierKey", "Dim_Carrier.CarrierKey", True),
    ("Purchases_OrderDate", "Fact_Purchases.OrderDate", "Dim_Date.Date", True), ("Purchases_PromisedDate", "Fact_Purchases.PromisedDate", "Dim_Date.Date", False),
    ("Purchases_ReceivedDate", "Fact_Purchases.ReceivedDate", "Dim_Date.Date", False), ("Purchases_Product", "Fact_Purchases.ProductKey", "Dim_Product.ProductKey", True),
    ("Purchases_Warehouse", "Fact_Purchases.WarehouseKey", "Dim_Warehouse.WarehouseKey", True),
    ("Inventory_SnapshotDate", "Fact_Inventory.SnapshotDate", "Dim_Date.Date", True), ("Inventory_Product", "Fact_Inventory.ProductKey", "Dim_Product.ProductKey", True),
    ("Inventory_Warehouse", "Fact_Inventory.WarehouseKey", "Dim_Warehouse.WarehouseKey", True),
    ("Product_Supplier", "Dim_Product.SupplierKey", "Dim_Supplier.SupplierKey", True),
]


def table_tmdl(name, t):
    out = [f"/// {t['desc']}", f"table {name}", f"\tlineageTag: {tag(name)}"]
    if t.get("date_table"):
        out.append("\tdataCategory: Time")
    out.append("")
    for col, dtype, o in t["cols"]:
        out.append(f"\tcolumn {q(col)}")
        out.append(f"\t\tdataType: {dtype}")
        if o.get("key"):
            out += ["\t\tisUnique", "\t\tisKey"]
        if o.get("hide"):
            out.append("\t\tisHidden")
        if dtype == DT:
            out.append("\t\tformatString: yyyy-mm-dd")
        elif o.get("fmt"):
            out.append(f"\t\tformatString: {o['fmt']}")
        elif dtype == I:
            out.append("\t\tformatString: 0")
        out += [f"\t\tlineageTag: {tag(name, col)}", "\t\tsummarizeBy: none", f"\t\tsourceColumn: {col}"]
        if o.get("sort"):
            out.append(f"\t\tsortByColumn: {o['sort']}")
        if dtype == DT:
            out += ["", "\t\tannotation UnderlyingDateTimeDataType = Date"]
        out.append("")
    for col, dtype, expr in t.get("calc", []):
        out += [f"\tcolumn {q(col)} = {expr}", f"\t\tdataType: {dtype}", "\t\tformatString: 0", f"\t\tlineageTag: {tag(name, col)}", "\t\tsummarizeBy: none", ""]
    for hname, levels in t.get("hier", []):
        out += [f"\thierarchy {q(hname)}", f"\t\tlineageTag: {tag(name, 'h', hname)}", ""]
        for lv in levels:
            out += [f"\t\tlevel {q(lv)}", f"\t\t\tlineageTag: {tag(name, 'h', hname, lv)}", f"\t\t\tcolumn: {q(lv)}", ""]
    types = ", ".join('{"%s", %s}' % (c, M_TYPE[d]) for c, d, _ in t["cols"])
    m = ["let", f'    Source = Csv.Document(File.Contents(DataFolder & "{name}.csv"), [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),',
         "    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),"]
    step = "Promoted"
    if t.get("nulls"):
        cols = ", ".join(f'"{c}"' for c in t["nulls"])
        m.append(f'    BlanksToNull = Table.ReplaceValue(Promoted, "", null, Replacer.ReplaceValue, {{{cols}}}),')
        step = "BlanksToNull"
    m += [f'    Typed = Table.TransformColumnTypes({step}, {{{types}}}, "en-US")', "in", "    Typed"]
    out += [f"\tpartition {name} = m", "\t\tmode: import", "\t\tsource ="] + ["\t\t\t\t" + l for l in m] + ["", "\tannotation PBI_ResultType = Table", ""]
    return "\n".join(out)


def measures_tmdl():
    out = ["/// Holds every measure. A convention: measures can live in any table.", "table _Measures", "\tisHidden", f"\tlineageTag: {tag('_Measures')}", ""]
    for folder, name, expr, fmt, desc in MEASURES:
        if desc:
            out.append(f"\t/// {desc}")
        if isinstance(expr, list):
            out.append(f"\tmeasure {q(name)} =")
            out += ["\t\t\t" + l for l in expr]
        else:
            out.append(f"\tmeasure {q(name)} = {expr}")
        out += [f"\t\tformatString: {fmt}", f"\t\tdisplayFolder: {folder}", f"\t\tlineageTag: {tag('_Measures', name)}", ""]
    out += ["\tcolumn m", "\t\tdataType: int64", "\t\tisHidden", f"\t\tlineageTag: {tag('_Measures', 'm')}", "\t\tsummarizeBy: none", "\t\tsourceColumn: m", "",
            "\tpartition _Measures = m", "\t\tmode: import", "\t\tsource = let Source = #table(type table [m = Int64.Type], {{1}}) in Source", ""]
    return "\n".join(out)


def build_model():
    if os.path.exists(SM):
        shutil.rmtree(SM)
    d = os.path.join(SM, "definition")
    order = list(TABLES) + ["_Measures"]
    write(os.path.join(SM, "definition.pbism"), json.dumps({"$schema": "https://developer.microsoft.com/json-schemas/fabric/item/semanticModel/definitionProperties/1.0.0/schema.json", "version": "4.2", "settings": {}}, indent=2))
    write(os.path.join(SM, ".platform"), json.dumps({"$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
                                                     "metadata": {"type": "SemanticModel", "displayName": NAME}, "config": {"version": "2.0", "logicalId": tag("semanticmodel")}}, indent=2))
    write(os.path.join(d, "database.tmdl"), "database\n\tcompatibilityLevel: 1606\n")
    write(os.path.join(d, "model.tmdl"), "\n".join([
        "model Model", "\tculture: en-US", "\tdefaultPowerBIDataSourceVersion: powerBI_V3", "\tdiscourageImplicitMeasures", "\tsourceQueryCulture: en-US",
        "\tdataAccessOptions", "\t\tlegacyRedirects", "\t\treturnErrorValuesAsNull", "",
        "annotation PBI_QueryOrder = " + json.dumps(["DataFolder"] + order).replace(" ", ""), "", "annotation __PBI_TimeIntelligenceEnabled = 0", "",
        'annotation PBI_ProTooling = ["DevMode"]', ""] + [f"ref table {q(t)}" for t in order] + ["", "ref role 'West Region'", "", "ref cultureInfo en-US", ""]))
    write(os.path.join(d, "expressions.tmdl"), "\n".join([
        "/// Folder that holds the CSV files, ending in a backslash. Change it once and every table follows.",
        f'expression DataFolder = "{DATA}" meta [IsParameterQuery = true, Type = "Text", IsParameterQueryRequired = true]',
        f"\tlineageTag: {tag('DataFolder')}", "", "\tannotation PBI_ResultType = Text", ""]))
    write(os.path.join(d, "cultures", "en-US.tmdl"), "cultureInfo en-US\n")
    write(os.path.join(d, "roles", "West Region.tmdl"), "\n".join([
        "/// Static row-level security: members see only the West region warehouse and its sales, purchases and stock.",
        "role 'West Region'", "\tmodelPermission: read", "", '\ttablePermission Dim_Warehouse = Dim_Warehouse[Region] = "West"', ""]))
    rels = []
    for n, f, t, active in RELS:
        rels += [f"relationship {n}"] + ([] if active else ["\tisActive: false"]) + [f"\tfromColumn: {f}", f"\ttoColumn: {t}", ""]
    write(os.path.join(d, "relationships.tmdl"), "\n".join(rels))
    for n, t in TABLES.items():
        write(os.path.join(d, "tables", f"{n}.tmdl"), table_tmdl(n, t))
    write(os.path.join(d, "tables", "_Measures.tmdl"), measures_tmdl())


# --------------------------------------------------------------------------- report
SCHEMA_VC = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.9.0/schema.json"
SCHEMA_PAGE = "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.1.0/schema.json"
lit = lambda v: {"expr": {"Literal": {"Value": v}}}
MEASURE_NAMES = {m[1] for m in MEASURES}


def field(ref):
    table, prop = ref.split(".", 1)
    kind = "Measure" if table == "_Measures" else "Column"
    return {kind: {"Expression": {"SourceRef": {"Entity": table}}, "Property": prop}}


def proj(ref):
    return {"field": field(ref), "queryRef": ref, "nativeQueryRef": ref.split(".", 1)[1]}


def visual(page, vid, vtype, pos, roles=None, title=None, sort=None, objects=None, z=0):
    x, y, w, h = pos
    v = {"visualType": vtype}
    if roles:
        v["query"] = {"queryState": {r: {"projections": [proj(f) for f in fs]} for r, fs in roles.items()}}
        if sort:
            ref, direction = sort
            v["query"]["sortDefinition"] = {"sort": [{"field": field(ref), "direction": direction}], "isDefaultSort": True}
    if objects:
        v["objects"] = objects
    vco = {}
    if title:
        vco["title"] = [{"properties": {"show": lit("true"), "text": lit(f"'{title}'")}}]
    if vco:
        v["visualContainerObjects"] = vco
    body = {"$schema": SCHEMA_VC, "name": vid, "position": {"x": x, "y": y, "z": z, "height": h, "width": w, "tabOrder": z}, "visual": v}
    return vid, body


def heading(page, text, sub):
    runs = [{"value": text, "textStyle": {"fontFamily": "Segoe UI Semibold", "fontSize": "20px", "color": "#14171D"}}]
    return visual(page, "title", "textbox", (24, 12, 640, 64), objects={"general": [{"properties": {"paragraphs": [
        {"textRuns": runs, "horizontalTextAlignment": "left"},
        {"textRuns": [{"value": sub, "textStyle": {"fontFamily": "Segoe UI", "fontSize": "12px", "color": "#5A6170"}}], "horizontalTextAlignment": "left"}]}}]})


def slicer(page, vid, ref, label, x, select=None):
    objects = {"data": [{"properties": {"mode": lit("'Dropdown'")}}],
               "header": [{"properties": {"show": lit("true"), "text": lit(f"'{label}'")}}]}
    if select is not None:  # pre-selected value, e.g. 2025 for the Year slicer
        table, col = ref.split(".", 1)
        cond = {"In": {"Expressions": [{"Column": {"Expression": {"SourceRef": {"Source": "s"}}, "Property": col}}], "Values": [[{"Literal": {"Value": select}}]]}}
        objects["general"] = [{"properties": {"filter": {"filter": {"Version": 2, "From": [{"Name": "s", "Entity": table, "Type": 0}], "Where": [{"Condition": cond}]}}}}]
    return visual(page, vid, "slicer", (x, 8, 184, 80), {"Values": [ref]}, objects=objects)


def matrix_objects():
    return {"columnHeaders": [{"properties": {"columnAdjustment": lit("'growToFit'"), "autoSizeColumnWidth": lit("true")}}]}


def card(page, vid, measures, pos):
    return visual(page, vid, "cardVisual", pos, {"Data": [f"_Measures.{m}" for m in measures]})


PAGES = [
    ("ExecutiveOverview", "Executive Overview", lambda p: [
        heading(p, "Meridian Medical Supply", "Executive overview · revenue, margin and delivery"),
        slicer(p, "slYear", "Dim_Date.Year", "Year", 704, "2025L"), slicer(p, "slRegion", "Dim_Warehouse.Region", "Region", 888), slicer(p, "slCategory", "Dim_Product.Category", "Category", 1072),
        card(p, "kpis", ["Revenue", "Gross Margin %", "Revenue YoY %", "On-Time Delivery %"], (24, 96, 1232, 112)),
        visual(p, "revTrend", "lineChart", (24, 224, 728, 232), {"Category": ["Dim_Date.MonthYear"], "Y": ["_Measures.Revenue", "_Measures.Revenue PY"]}, "Revenue by month, this year and last", ("Dim_Date.MonthYear", "Ascending")),
        visual(p, "revCategory", "clusteredBarChart", (768, 224, 488, 232), {"Category": ["Dim_Product.Category"], "Y": ["_Measures.Revenue"]}, "Revenue by category", ("_Measures.Revenue", "Descending")),
        visual(p, "marginProduct", "clusteredBarChart", (24, 472, 728, 232), {"Category": ["Dim_Product.ProductName"], "Y": ["_Measures.Gross Margin %"]}, "Gross margin % by product, lowest first", ("_Measures.Gross Margin %", "Ascending")),
        visual(p, "otdCarrier", "clusteredBarChart", (768, 472, 488, 232), {"Category": ["Dim_Carrier.CarrierName"], "Y": ["_Measures.On-Time Delivery %"]}, "On-time delivery % by carrier", ("_Measures.On-Time Delivery %", "Ascending")),
    ]),
    ("SupplierPerformance", "Supplier Performance", lambda p: [
        heading(p, "Supplier Performance", "Purchase orders · on time, in full, lead time"),
        slicer(p, "slYear", "Dim_Date.Year", "Year", 888, "2025L"), slicer(p, "slTier", "Dim_Supplier.SupplierTier", "Supplier tier", 1072),
        card(p, "kpis", ["Supplier On-Time %", "In-Full %", "OTIF %", "Open PO Lines"], (24, 96, 1232, 112)),
        visual(p, "otSupplier", "clusteredBarChart", (24, 224, 600, 232), {"Category": ["Dim_Supplier.SupplierName"], "Y": ["_Measures.Supplier On-Time %"]}, "Supplier on-time %, worst first", ("_Measures.Supplier On-Time %", "Ascending")),
        visual(p, "otTrend", "lineChart", (640, 224, 616, 232), {"Category": ["Dim_Date.MonthYear"], "Y": ["_Measures.Supplier On-Time %"]}, "Supplier on-time % by month", ("Dim_Date.MonthYear", "Ascending")),
        visual(p, "supplierTable", "pivotTable", (24, 472, 1232, 232), {"Rows": ["Dim_Supplier.SupplierName"], "Values": ["_Measures.PO Lines", "_Measures.Purchase Spend", "_Measures.Supplier On-Time %", "_Measures.In-Full %", "_Measures.Avg Lead Time (Days)", "_Measures.Avg Days Late"]}, "Supplier scorecard", objects=matrix_objects()),
    ]),
    ("InventoryHealth", "Inventory Health", lambda p: [
        heading(p, "Inventory Health", "Month-end stock · semi-additive measures"),
        slicer(p, "slRegion", "Dim_Warehouse.Region", "Region", 888), slicer(p, "slCategory", "Dim_Product.Category", "Category", 1072),
        card(p, "kpis", ["On Hand Units", "Inventory Value", "Stockouts", "Below Safety Stock"], (24, 96, 1232, 112)),
        visual(p, "valueTrend", "lineChart", (24, 224, 728, 232), {"Category": ["Dim_Date.MonthYear"], "Y": ["_Measures.Inventory Value"]}, "Inventory value at month end", ("Dim_Date.MonthYear", "Ascending")),
        visual(p, "stockoutsWh", "clusteredBarChart", (768, 224, 488, 232), {"Category": ["Dim_Warehouse.WarehouseName"], "Y": ["_Measures.Stockouts"]}, "Stockouts by warehouse, latest month", ("_Measures.Stockouts", "Descending")),
        visual(p, "productTable", "pivotTable", (24, 472, 1232, 232), {"Rows": ["Dim_Product.Category", "Dim_Product.ProductName"], "Values": ["_Measures.On Hand Units", "_Measures.Inventory Value", "_Measures.Stockouts", "_Measures.Below Safety Stock", "_Measures.Units Sold"]}, "Stock by product", objects=matrix_objects()),
    ]),
]


def build_report():
    pages_dir = os.path.join(RP, "definition", "pages")
    if os.path.exists(pages_dir):
        shutil.rmtree(pages_dir)
    for pid, display, make in PAGES:
        write(os.path.join(pages_dir, pid, "page.json"), json.dumps({"$schema": SCHEMA_PAGE, "name": pid, "displayName": display, "displayOption": "FitToPage", "height": 720, "width": 1280}, indent=2))
        for z, (vid, body) in enumerate(make(pid)):
            body["position"]["z"] = body["position"]["tabOrder"] = z * 1000
            for r in body["visual"].get("query", {}).get("queryState", {}).values():
                for pj in r["projections"]:
                    if "Measure" in pj["field"]:
                        assert pj["nativeQueryRef"] in MEASURE_NAMES, pj["nativeQueryRef"]
            write(os.path.join(pages_dir, pid, "visuals", vid, "visual.json"), json.dumps(body, indent=2, ensure_ascii=False))
    write(os.path.join(pages_dir, "pages.json"), json.dumps({"$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.1.0/schema.json",
                                                             "pageOrder": [p[0] for p in PAGES], "activePageName": PAGES[0][0]}, indent=2))


if __name__ == "__main__":
    build_model()
    build_report()
    write(os.path.join(ROOT, ".gitignore"), "**/.pbi/localSettings.json\n**/.pbi/cache.abf\n")
    print("built", ROOT)
