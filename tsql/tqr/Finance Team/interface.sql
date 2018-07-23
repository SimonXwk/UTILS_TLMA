;
WITH 
-- ----------------------------------------------------------
cte_payments AS (

SELECT
  CAST(B4.APPROVED AS DATE) AS APPROVED
  , B4.ADMITNAME
  , B4.PAYINGINREFERENCE
  , B4.ACCOUNTREFERENCE
  , B1.PAYMENTAMOUNT
  , B1.PAYMENTAMOUNTNETT
  , S1.ADDITIONALCODE1
  , S1.ADDITIONALCODE5
  , S1.ADDITIONALCODE3

FROM
  Tbl_BATCHITEMSPLIT B1
  LEFT JOIN Tbl_BATCHITEM   B2 ON (B1.SERIALNUMBER = B2.SERIALNUMBER) AND (B1.RECEIPTNO = B2.RECEIPTNO) AND (B1.ADMITNAME = B2.ADMITNAME)
  LEFT JOIN TBL_BATCHHEADER B4 ON (B1.ADMITNAME = B4.ADMITNAME)
  LEFT JOIN Tbl_SOURCECODE  S1 ON (B1.SOURCECODE = S1.SOURCECODE)
WHERE
  B4.APPROVED IS NOT NULL
  AND B4.STAGE = 'Batch Approved'
  AND CAST(B4.APPROVED AS DATE) BETWEEN '20180701' AND '20180716'
)
-- ----------------------------------------------------------
-- ----------------------------------------------------------
,cte_batch_payment_range AS (
  select 1 as f
)
-- ----------------------------------------------------------
,cte_fix_batches AS (
SELECT 
  APPROVED, ADMITNAME, ADDITIONALCODE1, ADDITIONALCODE5, SUM(PAYMENTAMOUNTNETT) AS NET
FROM
  cte_payments
WHERE
  ADDITIONALCODE3 LIKE '%WinterSpring%' AND ADDITIONALCODE5 = '15'
GROUP BY
  APPROVED, ADMITNAME, ADDITIONALCODE1, ADDITIONALCODE5


)
-- ----------------------------------------------------------
select *
from 
  cte_fix_batches
  order by APPROVED, ADMITNAME