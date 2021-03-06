WITH
-- --------------------------------------------------------------
ones AS ( 
SELECT * FROM (VALUES (0),(1),(2),(3),(4),(5),(6),(7),(8),(9)) AS numbers(x) 
)
-- --------------------------------------------------------------
,generation_def AS(
SELECT *
FROM (VALUES
    (0, 1824, 'ANCIENT', '01.AC')
    ,(1825, 1844, 'EARLY COLONIAL', '02.EC')
    ,(1845, 1864, 'MID COLONIAL', '03.MC')
    ,(1865, 1884, 'LATE COLONIAL', '04.LC')
    ,(1885, 1904, 'HARD TIMERS', '05.HT')
    ,(1905, 1924, 'FEDERATION', '06.F')
    ,(1925, 1944, 'SILENT', '07.S')
    ,(1945, 1964, 'BABY BOOMERS', '08.BB')
    ,(1965, 1979, 'GENERATION X', '09.X')
    ,(1980, 1994, 'GENERATION Y', '10.Y')
    ,(1995, 2009, 'GENERATION Z', '11.Z')
    ,(2010, 9999, 'MILLENIALS', '12.M')  ) AS generation(y1,y2,gen,gen_abr) 
)
-- --------------------------------------------------------------
,generation AS (
SELECT CY=n.x, GEN=g.gen, GEN_ABR=g.gen_abr
FROM
    (SELECT x=1000*o1000.x + 100*o100.x + 10*o10.x + o1.x
    FROM ones o1, ones o10, ones o100, ones o1000 ) n
    LEFT JOIN generation_def g on(n.x>=g.y1 AND n.x<=g.y2)
WHERE n.x BETWEEN 1 AND YEAR(CURRENT_TIMESTAMP)
)
-- --------------------------------------------------------------
,cte_payments AS(
  SELECT
    B1.SERIALNUMBER,
    B2.DATEOFPAYMENT,
    B1.PAYMENTAMOUNT, B1.PAYMENTAMOUNTNETT, B1.GSTAMOUNT,
    B1.SOURCECODE,
    S1.SOURCETYPE,
    CASE WHEN LTRIM(RTRIM(B2.MANUALRECEIPTNO)) LIKE '[0-9][0-9]-[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9]'
    THEN LTRIM(RTRIM(B2.MANUALRECEIPTNO))
    ELSE
      CASE WHEN B2.NOTES IS NOT NULL AND B2.NOTES LIKE '%REX Order Number: [0-9][0-9]-[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9]%'
      THEN SUBSTRING(B2.NOTES, PATINDEX('%REX Order Number: [0-9][0-9]-[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9]%', B2.NOTES) + LEN('REX Order Number:') + 1, 11)
      ELSE NULL
      END
    END AS [REXORDERID],
    CASE WHEN B2.REVERSED IN (1, -1, 2)
    THEN 0
    ELSE DENSE_RANK() OVER(PARTITION BY CASE WHEN B2.REVERSED IN (1, -1, 2) THEN 0 ELSE -1 END ORDER BY CONCAT(CONVERT(VARCHAR(10), B2.DATEOFPAYMENT, 112), B2.SERIALNUMBER, B2.ADMITNAME, B2.RECEIPTNO) ASC)
    END AS [TRANSACTIONID],
    -- Note: For a merchandise order reference, PLEDGEID and EXTERNALREF seems to have the same reference to order id
    CASE WHEN B3.PLEDGEID IS NULL AND B1.EXTERNALREF IS NOT NULL
    THEN B1.EXTERNALREF
    ELSE 
      CASE WHEN B1.EXTERNALREF IS NULL AND B3.PLEDGEID IS NOT NULL
      THEN B3.PLEDGEID
      ELSE B3.PLEDGEID
      END 
    END AS [TQORDERID],
    B1.ADMITNAME,
    B2.REVERSED
    
  FROM
    TBL_BATCHITEMSPLIT            B1
    LEFT JOIN TBL_BATCHITEM       B2 ON (B1.SERIALNUMBER = B2.SERIALNUMBER) AND (B1.RECEIPTNO = B2.RECEIPTNO) AND (B1.ADMITNAME = B2.ADMITNAME)
    LEFT JOIN TBL_BATCHITEMPLEDGE B3 ON (B1.SERIALNUMBER = B3.SERIALNUMBER) AND (B1.RECEIPTNO = B3.RECEIPTNO) AND (B1.ADMITNAME = B3.ADMITNAME) AND (B1.LINEID = B3.LINEID)
    LEFT JOIN TBL_BATCHHEADER     B4 ON (B2.ADMITNAME = B4.ADMITNAME)
    LEFT JOIN Tbl_SOURCECODE      S1 ON (B1.SOURCECODE = S1.SOURCECODE)
  WHERE
    (B2.REVERSED IS NULL OR NOT(B2.REVERSED IN (1, -1)))
    AND (B4.STAGE ='Batch Approved')
    AND B2.DATEOFPAYMENT BETWEEN '20130101' AND CURRENT_TIMESTAMP
)
-- ---------------------------------------------------------------------------------
,cte_product as (
  SELECT PRODUCTID,
    SUBSTRING(PRODUCTID, PATINDEX('%[x][0-9]%', PRODUCTID)+1, 255) AS [PRODUCTID_CLEAN],
    PRODUCTTYPE,PRODUCTSUBTYPE,NAME,THEDESCRIPTION,COSTPRICE,SELLINGPRICE,GST,PRODUCTPICTURE
  FROM TBL_PRODUCT
)
-- --------------------------------------------------------------
,cte_order_detail as (
  SELECT ADMITNAME,ORDERTYPE,PRODUCTID,
    SUBSTRING(PRODUCTID, PATINDEX('%[x][0-9]%', PRODUCTID)+1, 255) AS [PRODUCTID_CLEAN],
    SOURCECODE,LINETOTAL,
    COSTPRICE,SELLINGPRICE,DISCOUNT,PERCENTDISCOUNT, ORDERED, DELIVERED
  FROM TBL_ORDERDETAIL
  WHERE ORDERTYPE = 'Customer'
)
-- --------------------------------------------------------------
,cte_order_header as (
  SELECT ADMITNAME,ORDERTYPE,ORDERDATE,SERIALNUMBER,SENDON,COMPLETED
  FROM TBL_ORDERHEADER
  WHERE ORDERTYPE = 'Customer'
)
-- ---------------------------------------------------------------------------------
,cte_orders as (
  SELECT 
    o1.ADMITNAME, o1.SERIALNUMBER, o1.ORDERTYPE, o1.ORDERDATE, o1.SENDON, o1.COMPLETED
    ,o2.SOURCECODE, o2.LINETOTAL, o2.COSTPRICE, o2.SELLINGPRICE, o2.DISCOUNT, o2.ORDERED, o2.DELIVERED
    ,o3.PRODUCTTYPE, o3.PRODUCTSUBTYPE, o3.NAME, o3.THEDESCRIPTION,  o3.GST
  FROM
    TBL_ORDERHEADER o1
    LEFT JOIN TBL_ORDERDETAIL o2 ON (o1.ADMITNAME = o2.ADMITNAME) 
    LEFT JOIN TBL_PRODUCT o3 ON (o2.PRODUCTID = o3.PRODUCTID)
  WHERE
    o1.ORDERTYPE = 'Customer' AND o2.ORDERTYPE = 'Customer'
)
-- ---------------------------------------------------------------------------------
,cte_rex_customerids as(
  -- BY default one contact can only have one REXID
  SELECT SERIALNUMBER
  , [PARAMETERVALUE] AS [LAST_REXID]
  , [REXIDS]
  , [THQIDS] AS [LAST_REXID_POOL]
  FROM (
    SELECT [SERIALNUMBER], [PARAMETERVALUE]
      , ROW_NUMBER() OVER(PARTITION BY [SERIALNUMBER] ORDER BY [CREATED] DESC) AS [PAT_SERIALNUMBER_ROW]
      , COUNT(PARAMETERVALUE) OVER(PARTITION BY [SERIALNUMBER]) AS [REXIDS]
      , COUNT(SERIALNUMBER) OVER(PARTITION BY [PARAMETERVALUE]) AS [THQIDS]
  FROM TBL_CONTACTPARAMETER
  WHERE [PARAMETERNAME] LIKE '%Customer%Number%'
  ) tmp
WHERE [PAT_SERIALNUMBER_ROW] = 1
)

-- ---------------------------------------------------------------------------------


select 
  o1.ADMITNAME AS [TQ_ORDER_ID]
  -- ,COUNT(DISTINCT t1.TRANSACTIONID) AS [TRANSACTIONIDS]

  ,SUM(o1.LINETOTAL) AS [LINE_TOTAL]
  
  ,SUM(t1.PAYMENTAMOUNT) AS [PMT_TOTAL]
  ,SUM(o1.LINETOTAL) - SUM(t1.PAYMENTAMOUNT) AS [DIFF]
  
  
from 
  cte_orders o1
  left join Tbl_BATCHITEMSPLIT t1 on (t1.EXTERNALREF = o1.ADMITNAME and t1.SOURCECODE = o1.SOURCECODE)
where 
  -- t1.SOURCETYPE like '%Merch%'
  -- and t1.TQORDERID is not null

   o1.ADMITNAME like 'OR0013935'

group by 
  o1.ADMITNAME
-- having 
--   COUNT(DISTINCT t1.TRANSACTIONID) = 1


-- order BY
--   COUNT(DISTINCT t1.TRANSACTIONID) DESC


-- ---------------------------------------------------------------------------------
-- OPTION(RECOMPILE)
-- ---------------------------------------------------------------------------------