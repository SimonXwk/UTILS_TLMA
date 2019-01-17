;
WITH
cte_payments AS(
  SELECT
    B1.SERIALNUMBER,
    B2.DATEOFPAYMENT,
    B1.PAYMENTAMOUNT,
    B1.SOURCECODE,
    S1.SOURCETYPE,
    CASE WHEN B2.REVERSED=2
    THEN 0
    ELSE DENSE_RANK() OVER(PARTITION BY CASE WHEN B2.REVERSED IN (1, -1, 2) THEN 0 ELSE -1 END ORDER BY CONCAT(CONVERT(VARCHAR(10), B2.DATEOFPAYMENT, 112), B2.SERIALNUMBER, B2.ADMITNAME, B2.RECEIPTNO) ASC)
    END AS [TRANSACTION ID],
    B1.EXTERNALREF
  FROM
    TBL_BATCHITEMSPLIT            B1
    LEFT JOIN TBL_BATCHITEM       B2 ON (B1.SERIALNUMBER = B2.SERIALNUMBER) AND (B1.RECEIPTNO = B2.RECEIPTNO) AND (B1.ADMITNAME = B2.ADMITNAME)
    LEFT JOIN TBL_BATCHITEMPLEDGE B3 ON (B1.SERIALNUMBER = B3.SERIALNUMBER) AND (B1.RECEIPTNO = B3.RECEIPTNO) AND (B1.ADMITNAME = B3.ADMITNAME) AND (B1.LINEID = B3.LINEID)
    LEFT JOIN TBL_BATCHHEADER     B4 ON (B2.ADMITNAME = B4.ADMITNAME)
    LEFT JOIN Tbl_SOURCECODE      S1 ON (B1.SOURCECODE = S1.SOURCECODE)
  WHERE
    (B2.REVERSED IS NULL OR NOT(B2.REVERSED IN (1, -1)))
    AND (B4.STAGE ='Batch Approved')
    AND B2.DATEOFPAYMENT BETWEEN '20120701' AND CURRENT_TIMESTAMP
)
-- ---------------------------------------------------------------------------------
,cte_product as (
  SELECT PRODUCTID,NAME,PRODUCTTYPE,PRODUCTSUBTYPE,THEDESCRIPTION,PRODUCTPICTURE,SELLINGPRICE
    ,SUBSTRING(PRODUCTID, PATINDEX('%[x][0-9]%', PRODUCTID)+1, 255) AS CLEANID
  FROM TBL_PRODUCT
)
-- --------------------------------------------------------------
,cte_order_detail as (
  SELECT ADMITNAME,ORDERTYPE,PRODUCTID,COSTPRICE,SELLINGPRICE,DISCOUNT,PERCENTDISCOUNT, ORDERED, DELIVERED, LINETOTAL
    ,SUBSTRING(PRODUCTID, PATINDEX('%[x][0-9]%', PRODUCTID)+1, 255) AS CLEANID
  FROM TBL_ORDERDETAIL
  WHERE ORDERTYPE = 'Customer'
)
-- --------------------------------------------------------------
,cte_order_header as (
  SELECT ADMITNAME,SERIALNUMBER,ORDERTYPE,ORDERDATE
  FROM TBL_ORDERHEADER
  WHERE ORDERTYPE = 'Customer'
)
-- ---------------------------------------------------------------------------------

select
  t1.SERIALNUMBER,
  CONCAT(
    CASE WHEN RTRIM(ISNULL(t2.TITLE,''))='' THEN '' ELSE RTRIM(t2.TITLE) + ' ' END
    ,CASE WHEN RTRIM(ISNULL(t2.FIRSTNAME,''))='' THEN '' ELSE RTRIM(t2.FIRSTNAME) + ' ' END
    ,CASE WHEN RTRIM(ISNULL(t2.OTHERINITIAL,''))='' THEN '' ELSE RTRIM(t2.OTHERINITIAL) + ' ' END
    ,CASE WHEN RTRIM(ISNULL(t2.KEYNAME,''))='' THEN '' ELSE RTRIM(t2.KEYNAME) END
  ) AS [FULL NAME],
  t1.DATEOFPAYMENT AS [TRANSACTION DATE],
  o1.ADMITNAME as [ORDER ID],
  -- t1.SOURCECODE,
  -- t1.SOURCETYPE,
  CASE WHEN o1.ADMITNAME IS NULL 
  THEN SUM(t1.PAYMENTAMOUNT)
  ELSE o2.LINETOTAL
  END AS [LINE TOTAL],
  t1.[TRANSACTION ID],
  o1.ORDERDATE as [ORDER DATE],
  o2.CLEANID as [PRODUCTID],
  o3.NAME as [PRODUCT NAME],
  o3.PRODUCTTYPE,
  o3.PRODUCTSUBTYPE,
  o2.COSTPRICE, o2.SELLINGPRICE, o2.DISCOUNT,
  o2.ORDERED,
  o2.DELIVERED
from 
  cte_payments t1
  left join tbl_contact t2 on (t1.SERIALNUMBER = t2.SERIALNUMBER)
  left join cte_order_header o1 on (t1.EXTERNALREF = o1.ADMITNAME)
  left join cte_order_detail o2 on(o1.ADMITNAME = o2.ADMITNAME)
  left join cte_product o3 on(o2.CLEANID = o3.CLEANID)
-- where 
--  t1.SOURCETYPE like 'merch%'
group by 
  t1.SERIALNUMBER, t1.DATEOFPAYMENT, t1.[TRANSACTION ID]
  -- , t1.SOURCECODE, t1.SOURCETYPE
  ,t2.TITLE, t2.FIRSTNAME, t2.OTHERINITIAL, t2.KEYNAME
  ,o1.ADMITNAME, o1.ORDERDATE
  , o2.CLEANID, o2.ORDERED, o2.DELIVERED,o2.SELLINGPRICE, o2.COSTPRICE, o2.SELLINGPRICE, o2.DISCOUNT, o2.LINETOTAL
  ,o3.NAME, o3.PRODUCTTYPE, o3.PRODUCTSUBTYPE
Having
  SUM(CASE WHEN t1.SOURCETYPE like '%Merchandise%' THEN t1.PAYMENTAMOUNT ELSE 0 END) > 0