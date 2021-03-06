-- DECLARE
--   @SEG01 AS VARCHAR(25) = '00.New Contacts'
--   , @SEG02 AS VARCHAR(25) = '01.Pledge Givers'
--   , @SEG03 AS VARCHAR(25) = '02.Major Donors'
--   , @SEG04 AS VARCHAR(25) = '03.Premium Donors'
--   , @SEG05 AS VARCHAR(25) = '04.High Value Donors'
--   , @SEG06 AS VARCHAR(25) = '05.General Donors'
--   , @SEG07 AS VARCHAR(25) = '06.LYBUNT Donors'
--   , @SEG08 AS VARCHAR(25) = '07.SYBUNT Donors'
--   , @SEG09 AS VARCHAR(25) = '08.Groups'
--   , @SEG10 AS VARCHAR(25) = '09.Merchandise Only'
--   , @SEG11 AS VARCHAR(25) = '10.Never Given'
--   , @SEG12 AS VARCHAR(25) = '11.Third Party'
--   , @SEG13 AS VARCHAR(25) = '99.Deceased';

WITH
-- ------------------------------------------------------------------------
cte_payments AS(
SELECT B1.SERIALNUMBER, B1.SOURCECODE , B1.PAYMENTAMOUNT
  , B2.DATEOFPAYMENT
  , [FY] = YEAR(B2.DATEOFPAYMENT) + IIF(MONTH(B2.DATEOFPAYMENT)<7,0,1)
  , B2.RECEIPTNO ,B2.ADMITNAME, B2.REVERSED
  -- , S1.SOURCETYPE
  -- , [ACCOUNT] = S1.ADDITIONALCODE1
  -- , [CLASS] = S1.ADDITIONALCODE3
  -- , [CAMPAIGNCODE] = S1.ADDITIONALCODE3
  -- , [TRXID] = CONCAT(B2.SERIALNUMBER, B2.ADMITNAME, B2.RECEIPTNO)
  , [TRXID] = DENSE_RANK() OVER (ORDER BY B2.SERIALNUMBER, B2.ADMITNAME, B2.RECEIPTNO)
FROM
  dbo.TBL_BATCHITEMSPLIT B1
  LEFT JOIN dbo.TBL_BATCHITEM B2 ON (B2.SERIALNUMBER=B2.SERIALNUMBER AND B1.ADMITNAME=B2.ADMITNAME AND B1.RECEIPTNO=B2.RECEIPTNO)
  LEFT JOIN dbo.TBL_BATCHHEADER B4 ON (B1.ADMITNAME=B4.ADMITNAME)
  -- LEFT JOIN dbo.TBL_SOURCECODE S1 ON (B1.SOURCECODE=S1.SOURCECODE)
WHERE
  B4.STAGE='Batch Approved' AND (B2.REVERSED IS NULL OR (B2.REVERSED NOT IN (1,-1)))
)
-- ------------------------------------------------------------------------
, cte_segments AS (
SELECT SERIALNUMBER
, [FY] = RIGHT(PARAMETERNAME, 4)
, [SEGMENT] = CONCAT(PARAMETERVALUE,'.',PARAMETERNOTE)
FROM dbo.TBL_CONTACTPARAMETER
WHERE PARAMETERNAME LIKE 'FY____'
)
-- ------------------------------------------------------------------------
, cte_fy_active_donors_with_segment AS(
SELECT ac1.FY,ac2.SEGMENT
, [DONORS] = COUNT(DISTINCT ac1.SERIALNUMBER)
-- , [TRXS] = COUNT(DISTINCT ac1.TRXID)
FROM cte_payments ac1
  LEFT JOIN cte_segments ac2 ON (ac1.SERIALNUMBER = ac2.SERIALNUMBER and ac1.FY = ac2.FY)
WHERE ac1.REVERSED <> 2 OR ac1.REVERSED IS NULL
GROUP BY ac1.FY, ac2.SEGMENT
)
-- ------------------------------------------------------------------------
, cte_fy_active_donors_trx_with_segment AS(
SELECT ac1.FY,ac2.SEGMENT
, [TRXS] = COUNT(DISTINCT ac1.TRXID)
FROM cte_payments ac1
  LEFT JOIN cte_segments ac2 ON (ac1.SERIALNUMBER = ac2.SERIALNUMBER and ac1.FY = ac2.FY)
WHERE ac1.REVERSED <> 2 OR ac1.REVERSED IS NULL
GROUP BY ac1.FY, ac2.SEGMENT
)
-- ------------------------------------------------------------------------
, cte_fy_active_donors_value_with_segment AS(
SELECT ac1.FY,ac2.SEGMENT
, [TOTAL] = COUNT(DISTINCT ac1.PAYMENTAMOUNT)
FROM cte_payments ac1
  LEFT JOIN cte_segments ac2 ON (ac1.SERIALNUMBER = ac2.SERIALNUMBER and ac1.FY = ac2.FY)
GROUP BY ac1.FY, ac2.SEGMENT
)
-- ------------------------------------------------------------------------
, cte_fy_segments AS(
SELECT FY, SEGMENT, [CONTACTS] = COUNT(DISTINCT SERIALNUMBER)
FROM cte_segments
GROUP BY FY, SEGMENT
)
-- ------------------------------------------------------------------------

select
  -- [FY] = ISNULL(t1.FY, t2.FY)
  -- , [SEGMENT] = ISNULL(t1.SEGMENT,'SNF')
  -- , [CONTACTS] = ISNULL(t1.CONTACTS,0)
  -- , [DONORS] = ISNULL(t2.DONORS,0)
  -- -- , [TRXS] = t2.TRXS
  -- , [TOTAL] = t3.TOTAL

   [FY] = ISNULL(t1.FY, ts.FY)
  , [SEGMENT] = ISNULL(ts.SEGMENT,'SNF')
  , [CONTACTS] = ISNULL(ts.CONTACTS,0)
  , [DONORS] = ISNULL(t1.DONORS,0)
  , [TOTAL] =  ISNULL(t2.TOTAL,0)
  , [TRXS] =  ISNULL(t3.TRXS,0)
from

  -- cte_fy_segments t1
  -- left join cte_fy_active_donors_value_with_segment t3 on (t1.FY = t3.FY and t1.SEGMENT = t3.SEGMENT)
  -- full outer join cte_fy_active_donors_with_segment t2 on (t1.FY = t2.FY and t1.SEGMENT = t2.SEGMENT)
  cte_fy_active_donors_with_segment t1 
  left join cte_fy_active_donors_value_with_segment t2 on (t1.FY = t2.FY and t1.SEGMENT = t2.SEGMENT)
  left join cte_fy_active_donors_trx_with_segment t3 on (t1.FY = t3.FY and t1.SEGMENT = t3.SEGMENT)
  full outer join cte_fy_segments ts on (t1.FY = ts.FY and t1.SEGMENT = ts.SEGMENT)
order by
  t2.fy asc, t1.SEGMENT asc


-- select
--   *
--   , [SEG_TOTAL] = ISNULL([00.New Contacts],0)
--     + ISNULL([01.Pledge Givers],0)
--     + ISNULL([02.Major Donors],0)
--     + ISNULL([03.Premium Donors],0)
--     + ISNULL([04.High Value Donors],0)
--     + ISNULL([05.General Donors],0)
--     + ISNULL([06.LYBUNT Donors],0)
--     + ISNULL([07.SYBUNT Donors],0)
--     + ISNULL([08.Groups],0)
--     + ISNULL([09.Merchandise Only],0)
--     + ISNULL([10.Never Given],0)
--     + ISNULL([11.Third Party],0)
--     + ISNULL([99.Deceased],0)
-- from
--   (
--     select
--       t1.FY, t2.DONORS
--       , t1.SEGMENT
--       , CONTACTS = COUNT(DISTINCT t1.SERIALNUMBER)
--     from
--       cte_segments t1
--       left join cte_fy_active_donors t2 on (t1.FY = t2.FY)
--     group by t1.FY, t2.DONORS, t1.SEGMENT
--   ) datatable
--   pivot
--   (
--     SUM(CONTACTS)
--     FOR SEGMENT IN (
--       [00.New Contacts]
--       ,[01.Pledge Givers]
--       ,[02.Major Donors]
--       ,[03.Premium Donors]
--       ,[04.High Value Donors]
--       ,[05.General Donors]
--       ,[06.LYBUNT Donors]
--       ,[07.SYBUNT Donors]
--       ,[08.Groups]
--       ,[09.Merchandise Only]
--       ,[10.Never Given]
--       ,[11.Third Party]
--       ,[99.Deceased]
--       )
--   ) pivottable