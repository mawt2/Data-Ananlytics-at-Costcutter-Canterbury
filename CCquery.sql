/* Create department dictionary */
DROP TABLE IF EXISTS DepDictionary
CREATE TABLE DepDictionary
(Department INT, DepartmentNew INT, DepartmentNewName NVARCHAR (50))
GO

/* Set dictionary values */
INSERT INTO     DepDictionary VALUES (001,1,'Food Cupboard')
INSERT INTO     DepDictionary VALUES (002,1,'Food Cupboard')
INSERT INTO 	DepDictionary VALUES (003,1,'Food Cupboard')
INSERT INTO 	DepDictionary VALUES (004,1,'Food Cupboard')
INSERT INTO 	DepDictionary VALUES (005,1,'Food Cupboard')
INSERT INTO 	DepDictionary VALUES (006,1,'Food Cupboard')
INSERT INTO 	DepDictionary VALUES (007,1,'Food Cupboard')
INSERT INTO 	DepDictionary VALUES (008,1,'Food Cupboard')
INSERT INTO 	DepDictionary VALUES (009,1,'Food Cupboard')
INSERT INTO 	DepDictionary VALUES (010,1,'Food Cupboard')
INSERT INTO 	DepDictionary VALUES (011,1,'Food Cupboard')
INSERT INTO 	DepDictionary VALUES (012,2,'Biscuits & Cakes')
INSERT INTO 	DepDictionary VALUES (013,2,'Biscuits & Cakes')
INSERT INTO 	DepDictionary VALUES (014,2,'Biscuits & Cakes')
INSERT INTO 	DepDictionary VALUES (015,3,'Bakery')
INSERT INTO 	DepDictionary VALUES (016,1,'Food Cupboard')
INSERT INTO 	DepDictionary VALUES (017,4,'Dairy & Chilled')
INSERT INTO 	DepDictionary VALUES (018,5,'Soft Drinks')
INSERT INTO 	DepDictionary VALUES (019,6,'Confectionary')
INSERT INTO 	DepDictionary VALUES (020,7,'Crisps & Snacks')
INSERT INTO 	DepDictionary VALUES (021,7,'Crisps & Snacks')
INSERT INTO 	DepDictionary VALUES (022,8,'Household')
INSERT INTO 	DepDictionary VALUES (023,8,'Household')
INSERT INTO 	DepDictionary VALUES (024,8,'Household')
INSERT INTO 	DepDictionary VALUES (025,8,'Household')
INSERT INTO 	DepDictionary VALUES (026,8,'Household')
INSERT INTO 	DepDictionary VALUES (027,8,'Household')
INSERT INTO 	DepDictionary VALUES (028,8,'Household')
INSERT INTO 	DepDictionary VALUES (029,9,'Alcohol')
INSERT INTO 	DepDictionary VALUES (030,10,'Tobacco')
INSERT INTO 	DepDictionary VALUES (031,8,'Household')
INSERT INTO 	DepDictionary VALUES (032,8,'Household')
INSERT INTO 	DepDictionary VALUES (033,4,'Diary & Chilled')
INSERT INTO 	DepDictionary VALUES (034,8,'Household')
INSERT INTO 	DepDictionary VALUES (035,11,'Frozen')
INSERT INTO 	DepDictionary VALUES (036,11,'Frozen')
INSERT INTO 	DepDictionary VALUES (037,4,'Diary & Chilled')
INSERT INTO 	DepDictionary VALUES (045,12,'Lottery & PayPoint')
INSERT INTO 	DepDictionary VALUES (046,12,'Lottery & PayPoint')
INSERT INTO 	DepDictionary VALUES (048,8,'Household')
INSERT INTO 	DepDictionary VALUES (056,8,'Food Cupboard')
INSERT INTO 	DepDictionary VALUES (057,8,'Food Cupboard')
INSERT INTO 	DepDictionary VALUES (066,8,'Dairy & Chilled')
INSERT INTO 	DepDictionary VALUES (070,12,'Lottery & PayPoint')
INSERT INTO 	DepDictionary VALUES (071,10,'Tobacco')
GO

/* AVERAGE MARGIN BY DEPARTMENT */
/**********************************************************************/
/* Get updated department names */
DROP VIEW IF EXISTS    ProductsNewDep
GO
CREATE VIEW            ProductsNewDep AS
SELECT                 DepDictionary.DepartmentNew, DepDictionary.Department, tblProducts.Stocked,
	               tblProducts.SupplierCost, tblProducts.PackQuantity, tblProducts.CurrentSell,
	               tblProducts.WeeklySales, tblProducts.Description
FROM                   DepDictionary
INNER JOIN             tblProducts ON tblProducts.Department=DepDictionary.Department;
GO
/* Filtered table */
DROP TABLE IF EXISTS   ProductsNewDepPrelim
SELECT                 Department, CAST(DepartmentNew AS INT)[DepartmentNew], SupplierCost, 
                       (PackQuantity*CurrentSell) AS Revenue, Description, WeeklySales
INTO                   ProductsNewDepPrelim
FROM                   ProductsNewDep
WHERE                  Stocked = 1 AND DepartmentNew IS NOT NULL AND SupplierCost > 0
                       AND NOT(Description = 'LIPTON ICE TEA PEACH PM1') AND NOT(Description = 'DIET COKE PM1.75')
                       AND NOT(Description = 'MOUNTAIN DEW REGULAR PM1.15') AND NOT(Description = 'PEPSI REGULAR')  
GO
/* Results table  */ 
DROP TABLE IF EXISTS   DepAvMar
GO
SELECT                 DepartmentNew,
                       ROUND(AVG(((Revenue-SupplierCost)/NULLIF(Revenue,0))*100),1) AS AvMargin
INTO                   DepAvMar
FROM                   ProductsNewDepPrelim
GROUP BY               DepartmentNew
ORDER BY               DepartmentNew
GO

/* BEST SELLER BY DEPARTMENT */
/****************************************************************************/
/* Results table */
DROP VIEW IF EXISTS DepBestSellPrelim
GO
CREATE VIEW DepBestSellPrelim AS
SELECT   DepartmentNew, MAX(WeeklySales) AS WeeklySales
FROM 
     ProductsNewDepPrelim
GROUP BY
    DepartmentNew

GO


/* Get bestseller product names */
DROP Table IF EXISTS DepBestSell
GO
SELECT 
    ProductsNewDepPrelim.Description,
	DepBestSellPrelim.WeeklySales,
	DepBestSellPrelim.DepartmentNew
	
INTO
    DepBestSell

FROM
    ProductsNewDepPrelim
INNER JOIN DepBestSellPrelim ON DepBestSellPrelim.WeeklySales=ProductsNewDepPrelim.WeeklySales 
AND DepBestSellPrelim.DepartmentNew=ProductsNewDepPrelim.DepartmentNew

	GO

/* PROPORTION OF SALES BY DEPARTMENT */
/******************************************************************************/

/* Filtered table */
DROP TABLE IF EXISTS Sales
GO
SELECT 
    Department, PercOfSales
INTO
    Sales
FROM 
    dbo.Sheet1$
WHERE 
    Department IS NOT NULL AND PercOfSales IS NOT NULL
GO

/* Get updated department names */
DROP VIEW IF EXISTS SalesPrelim
GO
CREATE VIEW SalesPrelim AS
SELECT 
    DepDictionary.DepartmentNew,
	DepDictionary.Department,
	Sales.PercOfSales
FROM
    DepDictionary
INNER JOIN Sales ON Sales.Department=DepDictionary.Department;
GO

/* Results table */
DROP TABLE IF EXISTS SalesNewDep
GO
SELECT   
    DepartmentNew,
    ROUND(SUM(PercOfSales),1) AS PercOfSalesNewDep
INTO
    SalesNewDep
FROM 
     SalesPrelim
GROUP BY 
    DepartmentNew
ORDER BY 
	DepartmentNew
GO

