WITH
-- ------------------------------------------------------------
cte_payments as (
SELECT
  B1.SERIALNUMBER
  ,B2.DATEOFPAYMENT , B1.PAYMENTAMOUNT
  ,[CAMPAIGN] = S1.ADDITIONALCODE3
FROM
  dbo.TBL_BATCHITEMSPLIT B1
  LEFT JOIN dbo.TBL_BATCHITEM B2 ON (B2.SERIALNUMBER=B2.SERIALNUMBER AND B1.ADMITNAME=B2.ADMITNAME AND B1.RECEIPTNO=B2.RECEIPTNO)
  LEFT JOIN dbo.TBL_BATCHHEADER B4 ON (B1.ADMITNAME=B4.ADMITNAME)
  LEFT JOIN dbo.TBL_SOURCECODE S1 ON (B1.SOURCECODE=S1.SOURCECODE)
 WHERE
  B4.STAGE='Batch Approved'
  AND (B2.REVERSED IS NULL OR (B2.REVERSED NOT IN (1,-1)))
)
-- ------------------------------------------------------------

-- ------------------------------------------------------------
select 
*
from 
(
  select 
    t1.SERIALNUMBER, t1.CAMPAIGN
    ,t2.CONTACTTYPE, t2.PRIMARYCATEGORY, [DNM] = IIF(t2.DONOTMAIL=-1,'YES' , NULL)
    , [TOTAL]=SUM(t1.PAYMENTAMOUNT)
  from 
    cte_payments t1
    left join TBL_CONTACT t2 on (t1.SERIALNUMBER = t2.SERIALNUMBER)
  where 
    t1.CAMPAIGN in (
      'Clearance Catalogue FY18', '19PP.M WinterSpring'
    )
  group by
    t1.SERIALNUMBER, t1.CAMPAIGN
    ,t2.CONTACTTYPE, t2.PRIMARYCATEGORY, t2.DONOTMAIL
) DATATABLE
PIVOT
(
	SUM([TOTAL])
	FOR CAMPAIGN IN ([Clearance Catalogue FY18], [19PP.M WinterSpring])
) PIVOTTABLE

