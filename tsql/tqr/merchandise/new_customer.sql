DECLARE @REX_RECEIPT_PATTERN varchar(100), @REX_RECEIPT_NOTE_PATTERN varchar(100), @REX_RECEIPT_PREFIX varchar(20);
SET @REX_RECEIPT_PREFIX = 'REX Order Number:'
SET @REX_RECEIPT_PATTERN = '[0-9][0-9]-[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9]';
SET @REX_RECEIPT_NOTE_PATTERN = '%' + @REX_RECEIPT_PREFIX + ' ' + @REX_RECEIPT_PATTERN + '%';
WITH
cte_payments as (
  SELECT
    B1.SERIALNUMBER, B2.DATEOFPAYMENT, B1.PAYMENTAMOUNT
    , [ISMERCHANDISE]=IIF(S1.SOURCETYPE LIKE 'MERCH%', -1, 0)
    , [ISMERCHSPONSORSHIP]=IIF(S1.SOURCETYPE LIKE 'MERCH%SPONSORSHIP%', -1, 0)
    , [ISPURCHASE]=IIF(S1.SOURCETYPE LIKE 'MERCH%PURCHASE' OR S1.SOURCETYPE LIKE 'MERCH%POSTAGE', -1, 0)
    , C1.CONTACTTYPE, C1.PRIMARYCATEGORY, C1.DONOTMAIL, C1.DONOTMAILREASON
    , C1.SORTKEYREF1, C1.SORTKEYREFREL1 , C1.SORTKEYREFREL2
    , C2.DONOTCALL
    , [MANUALRECEIPTNO] = LTRIM(RTRIM(ISNULL(B2.MANUALRECEIPTNO, '')))
  FROM
    TBL_BATCHITEMSPLIT        B1
    LEFT JOIN TBL_BATCHITEM   B2 ON (B1.SERIALNUMBER = B2.SERIALNUMBER) AND (B1.RECEIPTNO = B2.RECEIPTNO) AND (B1.ADMITNAME = B2.ADMITNAME)
    LEFT JOIN TBL_BATCHHEADER B4 ON (B2.ADMITNAME = B4.ADMITNAME)
    LEFT JOIN TBL_SOURCECODE  S1 ON (B1.SOURCECODE = S1.SOURCECODE)
    LEFT JOIN TBL_CONTACT     C1 ON (B1.SERIALNUMBER = C1.SERIALNUMBER)
    LEFT JOIN TBL_CONTACTATTRIBUTE C2 ON (B1.SERIALNUMBER = C2.SERIALNUMBER)
  WHERE
    (B2.REVERSED IS NULL OR NOT (B2.REVERSED=1 OR B2.REVERSED=-1)) -- Full reversal and re-entry should be excluded all time
    AND (B4.STAGE ='Batch Approved')
    AND (C1.CONTACTTYPE NOT LIKE 'ADDRESS')
)
-- --------------------------------------------------------------
,cte_first_date as (
  SELECT SERIALNUMBER, [FIRSTDATE]=MIN(DATEOFPAYMENT)
  FROM TBL_BATCHITEM
  WHERE (REVERSED IS NULL OR NOT (REVERSED=1 OR REVERSED=-1 OR REVERSED=2))
  GROUP BY SERIALNUMBER
)
  -- --------------------------------------------------------------
,cte_orders as (
  SELECT SERIALNUMBER, DATEOFPAYMENT
      ,[ORDERID] = IIF( NOT LTRIM(RTRIM(ISNULL(MANUALRECEIPTNO, ''))) LIKE @REX_RECEIPT_PATTERN,
        SUBSTRING(NOTES
          ,PATINDEX(@REX_RECEIPT_NOTE_PATTERN,NOTES)+LEN(@REX_RECEIPT_PREFIX)+1
          ,11)
        , MANUALRECEIPTNO)
  FROM TBL_BATCHITEM
  WHERE
    (REVERSED IS NULL OR NOT (REVERSED=1 OR REVERSED=-1 OR REVERSED=2))
    AND (MANUALRECEIPTNO LIKE '%'+@REX_RECEIPT_PATTERN+'%' OR NOTES LIKE @REX_RECEIPT_NOTE_PATTERN)

  GROUP BY SERIALNUMBER, MANUALRECEIPTNO, DATEOFPAYMENT, NOTES
)
-- --------------------------------------------------------------
 ,cte_onboard_last as (
  SELECT SERIALNUMBER
  , PARAMETERNAME = FIRST_VALUE(PARAMETERNAME) OVER (PARTITION BY SERIALNUMBER ORDER BY CREATED DESC)
  , PARAMETERVALUE = FIRST_VALUE(PARAMETERVALUE) OVER (PARTITION BY SERIALNUMBER ORDER BY CREATED DESC)
  , EFFECTIVEFROM = FIRST_VALUE(EFFECTIVEFROM) OVER (PARTITION BY SERIALNUMBER ORDER BY CREATED DESC)
  , EFFECTIVETO = FIRST_VALUE(EFFECTIVETO) OVER (PARTITION BY SERIALNUMBER ORDER BY CREATED DESC)
  , PARAMETERNOTE = FIRST_VALUE(PARAMETERNOTE) OVER (PARTITION BY SERIALNUMBER ORDER BY CREATED DESC)
  FROM Tbl_CONTACTPARAMETER
  WHERE PARAMETERNAME = 'Merch Onboarding'
  GROUP BY SERIALNUMBER, CREATED, PARAMETERNAME, PARAMETERVALUE, EFFECTIVEFROM, EFFECTIVETO, PARAMETERNOTE
)
-- --------------------------------------------------------------
select
  t1.SERIALNUMBER, t1.FIRSTDATE
  ,[FIRSTORDER]= MIN(t2.ORDERID)
  ,[FIRSTORDERS] = COUNT(DISTINCT t2.ORDERID)
  ,[TOTAL]=SUM(t3.PAYMENTAMOUNT)
  ,[MERCHANDISE_TOTAL] = SUM(IIF(t3.ISMERCHANDISE = -1, t3.PAYMENTAMOUNT,0))
  ,[MERCHANDISE_PURCHASE] = SUM(IIF(t3.ISPURCHASE = -1, t3.PAYMENTAMOUNT,0))
  ,[MERCHANDISE_PLEDGE] = SUM(IIF(t3.ISMERCHSPONSORSHIP = -1, t3.PAYMENTAMOUNT,0))
  ,t3.CONTACTTYPE, t3.PRIMARYCATEGORY, t3.DONOTMAIL, t3.DONOTMAILREASON, t3.DONOTCALL, t3.SORTKEYREF1, t3.SORTKEYREFREL1, t3.SORTKEYREFREL2
  ,[PARAMETERNAME]=MIN(t4.PARAMETERNAME)
  ,[PARAMETERVALUE]=MIN(t4.PARAMETERVALUE)
  ,[EFFECTIVEFROM]=MIN(t4.EFFECTIVEFROM)
  ,[EFFECTIVETO]=MIN(t4.EFFECTIVETO)
  ,[PARAMETERNOTE]=MIN(t4.PARAMETERNOTE)
  ,[THANKYOU_DUE] = DATEADD(dd, 7, t1.FIRSTDATE)
  ,[WELCOMPACK_DUE] = DATEADD(dd, 14, t1.FIRSTDATE)
  ,[FIRST_ORDER_SOURCE] = IIF(MIN(t3.MANUALRECEIPTNO)=MIN(t2.ORDERID),'Alternative Receipt Number','Notes' )
from
  cte_first_date t1
  left join cte_orders t2 on (t1.FIRSTDATE = t2.DATEOFPAYMENT and t1.SERIALNUMBER = t2.SERIALNUMBER)
  left join cte_payments t3 on (t1.FIRSTDATE = t3.DATEOFPAYMENT and t1.SERIALNUMBER = t3.SERIALNUMBER)
  left join cte_onboard_last t4 on (t1.SERIALNUMBER = t4.SERIALNUMBER)
where
  t1.FIRSTDATE between '20180701' and '20190630'
group by
  t1.SERIALNUMBER,t1.FIRSTDATE
  ,t3.CONTACTTYPE, t3.PRIMARYCATEGORY, t3.DONOTMAIL, t3.DONOTMAILREASON, t3.DONOTCALL, t3.SORTKEYREF1, t3.SORTKEYREFREL1 ,t3.SORTKEYREFREL2
having
  SUM(IIF(t3.ISMERCHANDISE = -1, t3.PAYMENTAMOUNT,0)) > 0
order by
  t1.FIRSTDATE desc