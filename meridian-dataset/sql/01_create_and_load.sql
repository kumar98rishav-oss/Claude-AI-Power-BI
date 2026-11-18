/* ============================================================================
   Meridian Medical Supply - create the database and load the clean CSVs.

   Run with sqlcmd (sets the folder the CSVs are in):
     sqlcmd -S localhost\SQLEXPRESS -E -C -i 01_create_and_load.sql -v DataFolder="C:\path\to\Meridian_SupplyChain\data"

   In SSMS: Query > SQLCMD Mode, then change the :setvar line below.
   Safe to re-run: it drops and recreates the tables in MeridianSupplyChain only.
   No foreign keys on purpose, so the orphan-key demos in later sessions can load.
   ========================================================================== */
:setvar DataFolder "C:\Users\risha\Downloads\ProgrammersIO_Training\Meridian_SupplyChain\data"
SET NOCOUNT ON;
IF DB_ID('MeridianSupplyChain') IS NULL CREATE DATABASE MeridianSupplyChain;
GO
USE MeridianSupplyChain;
GO
DROP TABLE IF EXISTS dbo.Fact_Sales, dbo.Fact_Purchases, dbo.Fact_Inventory,
    dbo.Dim_Date, dbo.Dim_Product, dbo.Dim_Supplier, dbo.Dim_Warehouse, dbo.Dim_Customer, dbo.Dim_Carrier;

CREATE TABLE dbo.Dim_Date (
    [Date] date NOT NULL PRIMARY KEY, [Year] int, Quarter varchar(2), QuarterNo int, MonthNo int, [Month] varchar(3),
    MonthYear varchar(8), YearMonth int, WeekdayNo int, Weekday varchar(3), IsWeekend varchar(3),
    FiscalYear varchar(6), FiscalQuarter varchar(3), FiscalMonthNo int);
CREATE TABLE dbo.Dim_Product (
    ProductKey int NOT NULL PRIMARY KEY, SKU varchar(10), ProductName varchar(60), Category varchar(30), Subcategory varchar(30),
    SupplierKey int, UnitCost decimal(10,2), ListPrice decimal(10,2), Status varchar(15));
CREATE TABLE dbo.Dim_Supplier (
    SupplierKey int NOT NULL PRIMARY KEY, SupplierName varchar(60), Country varchar(30), StandardLeadTimeDays int,
    PaymentTerms varchar(10), SupplierTier varchar(15));
CREATE TABLE dbo.Dim_Warehouse (
    WarehouseKey int NOT NULL PRIMARY KEY, WarehouseCode varchar(5), WarehouseName varchar(30), City varchar(30),
    State varchar(2), Region varchar(15), CapacityUnits int);
CREATE TABLE dbo.Dim_Customer (
    CustomerKey int NOT NULL PRIMARY KEY, CustomerName varchar(80), CustomerType varchar(20), City varchar(30),
    State varchar(2), Region varchar(15), CustomerSince date);
CREATE TABLE dbo.Dim_Carrier (
    CarrierKey int NOT NULL PRIMARY KEY, CarrierName varchar(30), ServiceMode varchar(15), StandardTransitDays int);
CREATE TABLE dbo.Fact_Sales (
    SalesLineKey int NOT NULL PRIMARY KEY, SalesOrderID int, [LineNo] int, OrderDate date, RequestedDeliveryDate date,
    ShipDate date NULL, DeliveryDate date NULL, CustomerKey int, ProductKey int, WarehouseKey int, CarrierKey int,
    Quantity int, UnitPrice decimal(10,2), DiscountPct decimal(5,2), SalesAmount decimal(12,2), CostAmount decimal(12,2));
CREATE TABLE dbo.Fact_Purchases (
    PONumber varchar(12) NOT NULL, POLineNo int NOT NULL, OrderDate date, PromisedDate date, ReceivedDate date NULL,
    SupplierKey int, ProductKey int, WarehouseKey int, QtyOrdered int, QtyReceived int NULL, UnitCost decimal(10,2), LineCost decimal(12,2),
    CONSTRAINT PK_Fact_Purchases PRIMARY KEY (PONumber, POLineNo));
CREATE TABLE dbo.Fact_Inventory (
    SnapshotDate date NOT NULL, ProductKey int NOT NULL, WarehouseKey int NOT NULL, OnHandQty int, OnOrderQty int,
    SafetyStockQty int, ReorderPointQty int, InventoryValue decimal(12,2),
    CONSTRAINT PK_Fact_Inventory PRIMARY KEY (SnapshotDate, ProductKey, WarehouseKey));
GO

DECLARE @t TABLE (name sysname);
INSERT @t VALUES ('Dim_Date'), ('Dim_Product'), ('Dim_Supplier'), ('Dim_Warehouse'), ('Dim_Customer'), ('Dim_Carrier'),
                 ('Fact_Sales'), ('Fact_Purchases'), ('Fact_Inventory');
DECLARE @n sysname, @sql nvarchar(max);
DECLARE c CURSOR LOCAL FAST_FORWARD FOR SELECT name FROM @t;
OPEN c; FETCH NEXT FROM c INTO @n;
WHILE @@FETCH_STATUS = 0
BEGIN
    SET @sql = N'BULK INSERT dbo.' + QUOTENAME(@n) + N' FROM ''$(DataFolder)\' + @n + N'.csv'' '
             + N'WITH (FORMAT = ''CSV'', FIRSTROW = 2, CODEPAGE = ''65001'', KEEPNULLS, TABLOCK);';
    EXEC (@sql);
    FETCH NEXT FROM c INTO @n;
END
CLOSE c; DEALLOCATE c;
GO

-- Proof: these counts and totals must match the Power BI model after a refresh.
SELECT 'Fact_Sales' AS [Table], COUNT(*) AS [Rows], SUM(SalesAmount) AS Total FROM dbo.Fact_Sales
UNION ALL SELECT 'Fact_Purchases', COUNT(*), SUM(LineCost) FROM dbo.Fact_Purchases
UNION ALL SELECT 'Fact_Inventory', COUNT(*), SUM(InventoryValue) FROM dbo.Fact_Inventory
UNION ALL SELECT 'Dim_Customer', COUNT(*), NULL FROM dbo.Dim_Customer
UNION ALL SELECT 'Dim_Product', COUNT(*), NULL FROM dbo.Dim_Product;
