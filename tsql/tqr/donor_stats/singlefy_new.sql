WITH
cte_batch_item AS (
SELECT SERIALNUMBER, DATEOFPAYMENT, PAYMENTAMOUNT, REVERSED, ADMITNAME, RECEIPTNO, OWNERCODE
  ,TRXID = IIF(REVERSED=2, 0, DENSE_RANK() OVER (PARTITION BY SERIALNUMBER ORDER BY DATEOFPAYMENT ASC,REVERSED ASC, RECEIPTNO ASC))
FROM TBL_BATCHITEM
WHERE (REVERSED IS NULL OR NOT (REVERSED = -1 OR REVERSED =1))
)
-- ------------------------------------------------------------------------
, cte_segments AS (
SELECT SERIALNUMBER
, [SEGMENT] = CONCAT(PARAMETERVALUE,'.',PARAMETERNOTE)
FROM TBL_CONTACTPARAMETER
WHERE PARAMETERNAME LIKE 'FY2019'
)

-- ------------------------------------------------------------------------
, cte_batch_item_total_cfy AS (
SELECT *
  ,TRXIDDESC = IIF(REVERSED=2, 0, DENSE_RANK() OVER (PARTITION BY SERIALNUMBER ORDER BY DATEOFPAYMENT DESC,REVERSED ASC, RECEIPTNO DESC))
FROM cte_batch_item
WHERE DATEOFPAYMENT BETWEEN '2018/07/01' AND '2019/06/30' 
)
-- ------------------------------------------------------------------------
, cte_batch_item_total_ltd AS (
SELECT *
FROM cte_batch_item
WHERE DATEOFPAYMENT <= '2019/06/30' 
)
-- ------------------------------------------------------------------------
, cte_batch_item_total_lfy AS (
SELECT *
FROM cte_batch_item
WHERE DATEOFPAYMENT BETWEEN '2017/07/01' AND '2018/06/30' 
)
-- ------------------------------------------------------------------------
, cte_sfy_stats AS (
SELECT st.SERIALNUMBER
  -- > Shared calcs
  , [TRXS] = COUNT(IIF(st.REVERSED=2,NULL,st.SERIALNUMBER)) -- COUNT(ALL expression) evaluates expression for each row in a group, and returns the number of nonnull values.
  , [TOTAL] = SUM(st.PAYMENTAMOUNT)
  , [AVG] = SUM(st.PAYMENTAMOUNT)/COUNT(IIF(st.REVERSED=2,NULL,st.SERIALNUMBER))
  , [STDEV] = STDEV(st.PAYMENTAMOUNT)
  , [MIN] = MIN(IIF(st.REVERSED=2,null,st.PAYMENTAMOUNT))
  , [MAX] = MAX(IIF(st.REVERSED=2,null,st.PAYMENTAMOUNT))
  , [FIRST_DATE] =  MIN(IIF(st.REVERSED=2,null,st.DATEOFPAYMENT))
  , [LAST_DATE] =  MAX(IIF(st.REVERSED=2,null,st.DATEOFPAYMENT))
  -- > Unique Calcs
  , [LAST_GIFT] = SUM(IIF(TRXIDDESC=1,st.PAYMENTAMOUNT,0))
  , [LFYTRXS] = (
    SELECT COUNT(IIF(tmp.REVERSED=2,NULL,tmp.SERIALNUMBER))
    FROM cte_batch_item_total_lfy tmp
    WHERE (st.SERIALNUMBER = tmp.SERIALNUMBER)
  )
  , [LEGACY] = (
    SELECT SUM(tmp1.PAYMENTAMOUNT)
    FROM 
      TBL_BATCHITEMSPLIT tmp1 
      RIGHT JOIN cte_batch_item_total_cfy tmp2 ON (((tmp1.SERIALNUMBER = tmp2.SERIALNUMBER) AND (tmp1.RECEIPTNO = tmp2.RECEIPTNO) AND (tmp1.ADMITNAME = tmp2.ADMITNAME)))
      LEFT JOIN Tbl_SOURCECODE tmp3 ON (tmp1.SOURCECODE = tmp3.SOURCECODE)
    WHERE tmp3.SOURCETYPE = 'Bequest' AND (tmp1.SERIALNUMBER = st.SERIALNUMBER)
  )
FROM cte_batch_item_total_cfy st 
  -- LEFT JOIN cte_batch_item_total_lfy lst ON (st.SERIALNUMBER = lst.SERIALNUMBER)
GROUP BY st.SERIALNUMBER
)
-- ------------------------------------------------------------------------
, cte_ltd_stats AS (
SELECT st.SERIALNUMBER
  -- > Shared calcs
  , [TRXS] = COUNT(IIF(st.REVERSED=2,NULL,st.SERIALNUMBER)) -- COUNT(ALL expression) evaluates expression for each row in a group, and returns the number of nonnull values.
  , [TOTAL] = SUM(st.PAYMENTAMOUNT)
  , [AVG] = SUM(st.PAYMENTAMOUNT)/COUNT(IIF(st.REVERSED=2,NULL,st.SERIALNUMBER))
  , [STDEV] = STDEV(st.PAYMENTAMOUNT)
  , [MIN] = MIN(IIF(st.REVERSED=2,null,st.PAYMENTAMOUNT))
  , [MAX] = MAX(IIF(st.REVERSED=2,null,st.PAYMENTAMOUNT))
  , [FIRST_DATE] =  MIN(IIF(st.REVERSED=2,null,st.DATEOFPAYMENT))
  , [LAST_DATE] =  MAX(IIF(st.REVERSED=2,null,st.DATEOFPAYMENT))
  -- > Unique Calcs
  , [MODE] = (
    SELECT TOP 1 tmp.PAYMENTAMOUNT
    FROM cte_batch_item_total_ltd tmp
    WHERE (tmp.SERIALNUMBER = st.SERIALNUMBER) AND (tmp.REVERSED IS NULL OR NOT (tmp.REVERSED=2))
    GROUP BY tmp.PAYMENTAMOUNT
    ORDER BY COUNT(tmp.PAYMENTAMOUNT) DESC, tmp.PAYMENTAMOUNT ASC
  )
  , [MEDIAN] = ((
    (SELECT TOP 1 md.PAYMENTAMOUNT FROM
      (SELECT TOP 50 PERCENT tmp.PAYMENTAMOUNT
      FROM cte_batch_item_total_ltd tmp
      WHERE (tmp.REVERSED IS NULL OR NOT (tmp.REVERSED=2)) AND (tmp.SERIALNUMBER = st.SERIALNUMBER)
      ORDER BY tmp.PAYMENTAMOUNT ASC) md
    ORDER BY md.PAYMENTAMOUNT DESC
    ) + 
    (SELECT TOP 1 md.PAYMENTAMOUNT FROM
      (SELECT TOP 50 PERCENT tmp.PAYMENTAMOUNT
      FROM cte_batch_item_total_ltd tmp
      WHERE (tmp.REVERSED IS NULL OR NOT (tmp.REVERSED=2)) AND (tmp.SERIALNUMBER = st.SERIALNUMBER)
      ORDER BY tmp.PAYMENTAMOUNT DESC) md
    ORDER BY md.PAYMENTAMOUNT ASC
    ))/2
  )
  , [FIRST_GIFT] = SUM(IIF(TRXID=1,st.PAYMENTAMOUNT,0))
  , [LENGTH] = DATEDIFF(year,MIN(IIF(REVERSED=2,null,DATEOFPAYMENT)),MAX(IIF(REVERSED=2,null,DATEOFPAYMENT)))
FROM cte_batch_item_total_ltd st
GROUP BY st.SERIALNUMBER
)
-- ------------------------------------------------------------------------

select 
  c1.SERIALNUMBER
  ,c1.CONTACTTYPE as [Contacttype]
  ,c1.GENDER as [Gender]
  ,c1.ADDRESSLINE4 as [State]
  ,c1.COUNTRY as [Country]
  ,c1.DATEOFBIRTH   as [DateOfBirth]
  --> Number of Transactions
  ,st1.TRXS as [FYTotalNo]
  ,st2.TRXS as [LTDTotalNo] 
  --> Total Value
  ,st1.TOTAL  as [FYTotal]
  ,st2.TOTAL  as [LTDTotal] 
  --> Value Per Transaction
  ,st1.AVG  as [FYAve]
  ,st2.AVG  as [LTDAve]
  --> Standard Deviation
  ,st1.STDEV  as [FYStdev]
  ,st2.STDEV  as [LTDStdev]
  --> MODE and MEDIAN
  ,st2.MODE  as [LTDMode]
  ,st2.MEDIAN  as [LTDMedian]
  --> Smallest Gift Size
  ,st1.MIN  as [FYMinGift]
  ,st2.MIN  as [LTDMinGift]
  --> Largest Gift Size
  ,st1.MAX  as [FYMaxGift]
  ,st2.MAX  as [LTDMaxGift]
  --> First Transaction Date
  -- ,st1.FIRST_DATE   as [FYFirstGift]
  ,st2.FIRST_DATE as [LTDFirstDate]
  ,st2.FIRST_GIFT as [LTDFirstGift]

  --> First Transaction Date
  ,st1.LAST_DATE  as [FYLastDate]
  -- ,st2.LAST_DATE   as [LTDLastDate]
  ,st1.LAST_GIFT as [FYLastGift]

  --> Support Length
  ,st2.LENGTH as [LengthOfSupportYrs]
  --> Last FY Number of Transaction
  ,st1.LFYTRXS as [LFYTotalNo]

  ,c1.ANONYMOUS as [ANON]
  ,c1.DONOTMAIL as [DONOTMAIL]
  ,c1.DONOTMAILREASON as [DONOTMAILREASON]
  ,c1.DONOTMAILFROM as [DONOTMAILFROM]
  ,c1.DONOTMAILUNTIL as [DONOTMAILUNTIL]
  ,c1.POSTCODE as [POSTCODE]
  ,c1.Primarycategory as [Primarycategory]

  --> LEGACY Total
  ,st1.LEGACY   as [LegacyPmt]
  --> Segment
  ,c2.SEGMENT as [Segment]

from 
  cte_sfy_stats st1
  left join cte_ltd_stats st2 on (st1.SERIALNUMBER=st2.SERIALNUMBER)
  left join TBL_CONTACT c1 on (st1.SERIALNUMBER = c1.SERIALNUMBER)
  left join cte_segments c2 on (st1.SERIALNUMBER = c2.SERIALNUMBER) 
where
  c1.CONTACTTYPE <> 'ADDRESS' 
order by 
  c1.SerialNumber