with
cte_postcode as (
 SELECT C1.SERIALNUMBER
  ,POSTCODE = LTRIM(RTRIM(C1.POSTCODE))
  ,ISNUMBER = ISNUMERIC(LTRIM(RTRIM(C1.POSTCODE)))
 FROM TBL_CONTACT C1
 WHERE
  C1.CONTACTTYPE <> 'ADDRESS'
  AND (RTRIM(C1.POSTCODE) <> '' OR C1.POSTCODE IS NOT NULL)
)
-- --------------------------------------------------------------
,cte_payments AS (
  SELECT
     B1.SERIALNUMBER,SUM(B1.PAYMENTAMOUNT) AS PAYMENTAMOUNT
  FROM
     TBL_BATCHITEMSPLIT        B1
     LEFT JOIN TBL_BATCHITEM   B2 ON (B1.SERIALNUMBER = B2.SERIALNUMBER) AND (B1.RECEIPTNO = B2.RECEIPTNO) AND (B1.ADMITNAME = B2.ADMITNAME)
     LEFT JOIN TBL_BATCHHEADER B4 ON (B2.ADMITNAME = B4.ADMITNAME)
  WHERE 
     (B2.REVERSED IS NULL OR NOT(B2.REVERSED=1 OR B2.REVERSED=-1)) AND (B4.STAGE ='Batch Approved')  /*Only Approved Batches and excluding fully reversed Batchitems(like they never exist)*/
     -- AND CAST(B2.DATEOFPAYMENT AS DATE) <= CAST(@MYDATE AS DATE)
  GROUP BY B1.SERIALNUMBER
)
-- --------------------------------------------------------------
select count(*) AS COUNT,sum(T2.PAYMENTAMOUNT) as TOTAL, CURRENT_TIMESTAMP as STP 
from 
  cte_postcode T1
  left join cte_payments T2 on (T1.SERIALNUMBER = T2.SERIALNUMBER)
where ISNUMBER = 1
and ( CAST(POSTCODE AS INT) BETWEEN 2264 AND 2315 )
