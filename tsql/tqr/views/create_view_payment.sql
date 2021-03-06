USE thankQ4_Reporter
GO

drop VIEW myView_payments
GO

-- create VIEW myView_payments AS
ALTER VIEW myView_payments AS
-- --------------------------------------------------------------
WITH
cte_sourcecode AS(
SELECT 
   SOURCECODE,SOURCETYPE
   ,IIF(SOURCETYPE LIKE 'Sponsorship', 'P',IIF(SOURCETYPE LIKE 'Merch%', 'M',IIF(SOURCETYPE LIKE 'Group', 'G',IIF(SOURCETYPE LIKE 'Bequest', 'B'
      ,IIF(ADDITIONALCODE3 IS NULL OR RTRIM(ADDITIONALCODE3) = '', 'U','C'))))) AS SOURCECATEGORY
   ,ADDITIONALCODE1 AS QBCODE,ADDITIONALCODE5 AS QBCLASS,ADDITIONALCODE3 AS CAMPAIGN
   ,DEFAULT_DES1=DESTINATIONCODE,DEFAULT_DES2=DESTINATIONCODE2
   ,ARCHIVE,EXCLUDEFROMDROPDOWN
   -- ,SOURCEDESCRIPTION,SOURCENOTES
   -- ,SOURCECODE_CREATEDBY=CREATEDBY,SOURCECODE_CREATED=CREATED
   -- ,SOURCECODE_MODIFIEDBY=MODIFIEDBY,SOURCECODE_MODIFIED=MODIFIED
FROM TBL_SOURCECODE)
-- --------------------------------------------------------------
,cte_batchitemsplit AS (
SELECT
   B1.SERIALNUMBER,B1.PAYMENTAMOUNT,B2.DATEOFPAYMENT
   ,ACTUAL_DES1=B1.DESTINATIONCODE,ACTUAL_DES2=B1.DESTINATIONCODE2
   ,B1.SOURCECODE2
   ,B6.*
   ,FY=IIF(MONTH(B2.DATEOFPAYMENT)<7,YEAR(B2.DATEOFPAYMENT),YEAR(B2.DATEOFPAYMENT)+1)
   ,FYMTH=IIF(MONTH(B2.DATEOFPAYMENT)<7,MONTH(B2.DATEOFPAYMENT)+6,MONTH(B2.DATEOFPAYMENT)-6)
   ,CY=YEAR(B2.DATEOFPAYMENT),CYMTH=MONTH(B2.DATEOFPAYMENT),DAY=DAY(B2.DATEOFPAYMENT)
   ,TRX_ID=CAST(DENSE_RANK() OVER(PARTITION BY B1.SERIALNUMBER ORDER BY B2.DATEOFPAYMENT ASC, CONCAT(B1.SERIALNUMBER,'-',B1.ADMITNAME,'-',B2.RECEIPTNO) ASC) AS INT)
   ,DATE_ID=CAST(DENSE_RANK() OVER(PARTITION BY B1.SERIALNUMBER ORDER BY B2.DATEOFPAYMENT ASC) AS INT)
   ,MTH_ID=CAST(DENSE_RANK() OVER(PARTITION BY B1.SERIALNUMBER ORDER BY YEAR(B2.DATEOFPAYMENT)*100+MONTH(B2.DATEOFPAYMENT) ASC) AS INT)
   ,FY_ID=CAST(DENSE_RANK() OVER(PARTITION BY B1.SERIALNUMBER ORDER BY IIF(MONTH(B2.DATEOFPAYMENT)<7,YEAR(B2.DATEOFPAYMENT),YEAR(B2.DATEOFPAYMENT)+1) ASC) AS INT) 
   ,B4.ADMITNAME,B2.REVERSED
   ,B4.APPROVED,B4.STAGE
   -- ,TRX_KEY=CONCAT(B1.SERIALNUMBER,'-',B1.ADMITNAME,'-',B2.RECEIPTNO)
FROM
   TBL_BATCHITEMSPLIT        B1
   LEFT JOIN TBL_BATCHITEM   B2 ON (B1.SERIALNUMBER = B2.SERIALNUMBER) AND (B1.RECEIPTNO = B2.RECEIPTNO) AND (B1.ADMITNAME = B2.ADMITNAME)
   LEFT JOIN TBL_BATCHHEADER B4 ON (B2.ADMITNAME = B4.ADMITNAME)
   LEFT JOIN cte_sourcecode  B6 ON (B1.SOURCECODE = B6.SOURCECODE)
WHERE 
   (B2.REVERSED IS NULL OR NOT(B2.REVERSED=1 OR B2.REVERSED=-1)) AND (B4.STAGE ='Batch Approved')  /*Only Approved Batches and excluding fully reversed Batchitems(like they never exist)*/
   -- AND CAST(B2.DATEOFPAYMENT AS DATE) <= CAST(@MYDATE AS DATE)
)
-- --------------------------------------------------------------
SELECT * FROM cte_batchitemsplit
