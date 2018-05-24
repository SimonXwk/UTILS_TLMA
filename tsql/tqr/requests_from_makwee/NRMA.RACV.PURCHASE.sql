with
-- --------------------------------------------------------------
ones AS (
	SELECT * FROM (VALUES (0), (1), (2), (3), (4),(5), (6), (7), (8), (9)) AS numbers(x)
)
-- --------------------------------------------------------------
,generation_def AS (
   SELECT * FROM (
      VALUES
      (0,1824,'ANCIENT','01.AC')
      ,(1825,1844,'EARLY COLONIAL','02.EC')
      ,(1845,1864,'MID COLONIAL','03.MC')
      ,(1865,1884,'LATE COLONIAL','04.LC')
      ,(1885,1904,'HARD TIMERS','05.HT')
      ,(1905,1924,'FEDERATION','06.F')
      ,(1925,1944,'SILENT','07.S')
      ,(1945,1964,'BABY BOOMERS','08.BB')
      ,(1965,1979,'GENERATION X','09.X')
      ,(1980,1994,'GENERATION Y','10.Y')
      ,(1995,2009,'GENERATION Z','11.Z')
      ,(2010,9999,'MILLENIALS','12.M')
      ) AS generation(y1,y2,gen,gen_abr) 
)
-- --------------------------------------------------------------
,cte_generation AS(
SELECT cy=n.x,gen=g.gen,gen_abr=g.gen_abr
FROM 
   (SELECT x=1000*o1000.x + 100*o100.x + 10*o10.x + o1.x FROM ones o1, ones o10, ones o100, ones o1000 ) n 
   LEFT JOIN generation_def g on(n.x>=g.y1 AND n.x<=g.y2)
WHERE n.x BETWEEN 1 AND YEAR(CURRENT_TIMESTAMP)
)
-- --------------------------------------------------------------
,cte_merch_acquisition_customer as (
  SELECT SERIALNUMBER
  FROM TBL_CONTACT
  WHERE CONTACTTYPE <> 'ADDRESS' AND ( SOURCE LIKE '%NRMA%' OR SOURCE LIKE '%RACV%')
)
-- --------------------------------------------------------------
,cte_product as (
  SELECT PRODUCTID,NAME,PRODUCTTYPE,PRODUCTSUBTYPE,THEDESCRIPTION,PRODUCTPICTURE,SELLINGPRICE
    ,SUBSTRING(PRODUCTID, PATINDEX('%[x][0-9]%', PRODUCTID)+1, 255) AS CLEANID
  FROM TBL_PRODUCT
)
-- --------------------------------------------------------------
,cte_order_detail as (
  SELECT ADMITNAME,ORDERTYPE,PRODUCTID,SELLINGPRICE,PERCENTDISCOUNT
    ,SUBSTRING(PRODUCTID, PATINDEX('%[x][0-9]%', PRODUCTID)+1, 255) AS CLEANID
  FROM TBL_ORDERDETAIL
)
-- --------------------------------------------------------------
,cte_order_header as (
  SELECT ADMITNAME,SERIALNUMBER,ORDERTYPE,ORDERDATE
  FROM TBL_ORDERHEADER
  WHERE ORDERTYPE = 'Customer'
)
-- --------------------------------------------------------------
,bd_merchcustomer_firstday as (
  SELECT
    mcf1.SERIALNUMBER,MIN(ORDERDATE) AS FIRSTORDERDATE
  FROM 
    cte_merch_acquisition_customer mcf1
    LEFT JOIN cte_order_header mscf2 ON (mcf1.SERIALNUMBER = mscf2.SERIALNUMBER)
  GROUP BY mcf1.SERIALNUMBER
)
-- --------------------------------------------------------------
,bd_merchcustomer_firstday_products as (
  SELECT
    t1.SERIALNUMBER
    ,t5.SOURCE,t5.CONTACTTYPE,t5.PRIMARYCATEGORY,t5.POSTCODE
    ,STATE = IIF(UPPER(LTRIM(RTRIM(t5.ADDRESSLINE4))) IN ('VIC','NSW','SA','QLD','WA','TAS','ACT','NT'),UPPER(LTRIM(RTRIM(t5.ADDRESSLINE4))),IIF(RTRIM(t5.ADDRESSLINE4)='' OR t5.ADDRESSLINE4 IS NULL,'[DNF]','O/S'))
    ,GENERATION = IIF((t5.DATEOFBIRTH IS NULL OR RTRIM(t5.DATEOFBIRTH)=''),'[DNF]',t6.GEN)
    ,t5.CREATED,YEAR(t5.CREATED) + IIF(MONTH(t5.CREATED)<7,0,1) AS CREATEDFY
    ,t2.ORDERDATE,YEAR(t2.ORDERDATE) + IIF(MONTH(t2.ORDERDATE)<7,0,1) AS ORDERFY
    ,t2.ADMITNAME AS ORDERID
    ,t3.PRODUCTID,t3.SELLINGPRICE,t3.PERCENTDISCOUNT
    ,t4.PRODUCTSUBTYPE,t4.PRODUCTTYPE,t4.NAME,t4.SELLINGPRICE AS PRODUCTSELLINGPRICE
  FROM 
    bd_merchcustomer_firstday t1
    LEFT JOIN cte_order_header t2 ON (t1.SERIALNUMBER = t2.SERIALNUMBER) AND (t1.FIRSTORDERDATE = t2.ORDERDATE)
    LEFT JOIN cte_order_detail t3 ON (t2.ADMITNAME = t3.ADMITNAME)
    LEFT JOIN cte_product t4 ON (t3.CLEANID = t4.CLEANID)
    LEFT JOIN TBL_CONTACT t5 ON (t1.SERIALNUMBER = t5.SERIALNUMBER)
    LEFT JOIN cte_generation t6 ON (YEAR(t5.DATEOFBIRTH) = t6.cy)
)
-- --------------------------------------------------------------

select
  *
from bd_merchcustomer_firstday_products
order by SERIALNUMBER

-- --------------------------------------------------------------
-- select
--   IIF(NAME IS NULL OR RTRIM(NAME)='','[NO ORDER PLACED]',NAME) AS PRODUCTNAME
--   ,PRODUCTTYPE,PRODUCTSUBTYPE,PRODUCTSELLINGPRICE
--   ,AVG(SELLINGPRICE) AS AVGSELLINGPRICE
--   ,COUNT(PRODUCTID) AS SOLD
--   ,COUNT(DISTINCT SERIALNUMBER) AS UNIQUECUSTOMERS
-- from 
--   bd_merchcustomer_firstday_products
-- group by 
--   IIF(NAME IS NULL OR RTRIM(NAME)='','[NO ORDER PLACED]',NAME) 
--   ,PRODUCTSUBTYPE,PRODUCTTYPE,PRODUCTSELLINGPRICE
-- order by  
--   COUNT(DISTINCT SERIALNUMBER) desc


