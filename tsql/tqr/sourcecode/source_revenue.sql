WITH
cte_sourcecode AS (
  SELECT
      SOURCECODE, SOURCETYPE, DESTINATIONCODE AS DDES1, DESTINATIONCODE2  AS DDES2
      ,ADDITIONALCODE3 AS CAMPAIGN, ADDITIONALCODE1 AS ACCOUNT, ADDITIONALCODE5 AS CLASS
      ,IIF(SOURCETYPE LIKE 'Merch%', 'M',IIF(SOURCETYPE = 'Bequest', 'B', 'F')) AS STREAM
  FROM TBL_SOURCECODE
)
,cte_payment AS (
    SELECT
      B1.SERIALNUMBER,B1.PAYMENTAMOUNT,B2.DATEOFPAYMENT
      ,FULLNAME=CONCAT(
        IIF(RTRIM(ISNULL(B2.TITLE,''))='','',RTRIM(B2.TITLE)+' ')
        ,IIF(RTRIM(ISNULL(B2.FIRSTNAME,''))='','',RTRIM(B2.FIRSTNAME)+' '),B2.KEYNAME)
      ,ADES1=B1.DESTINATIONCODE,ADES2=B1.DESTINATIONCODE2
      ,B1.SOURCECODE2
      ,B6.*
      ,FY=IIF(MONTH(B2.DATEOFPAYMENT)<7,YEAR(B2.DATEOFPAYMENT),YEAR(B2.DATEOFPAYMENT)+1)
      ,FYMTH=IIF(MONTH(B2.DATEOFPAYMENT)<7,MONTH(B2.DATEOFPAYMENT)+6,MONTH(B2.DATEOFPAYMENT)-6)
      ,CY=YEAR(B2.DATEOFPAYMENT),CYMTH=MONTH(B2.DATEOFPAYMENT),DAY=DAY(B2.DATEOFPAYMENT)
      ,TRX_ID=CAST(DENSE_RANK() OVER(PARTITION BY B1.SERIALNUMBER ORDER BY B2.DATEOFPAYMENT ASC, CONCAT(B1.SERIALNUMBER,'-',B1.ADMITNAME,'-',B2.RECEIPTNO) ASC) AS INT)
      ,DATE_ID=CAST(DENSE_RANK() OVER(PARTITION BY B1.SERIALNUMBER ORDER BY B2.DATEOFPAYMENT ASC) AS INT)
      ,FY_ID=CAST(DENSE_RANK() OVER(PARTITION BY B1.SERIALNUMBER ORDER BY IIF(MONTH(B2.DATEOFPAYMENT)<7,YEAR(B2.DATEOFPAYMENT),YEAR(B2.DATEOFPAYMENT)+1) ASC) AS INT) 
      ,B4.ADMITNAME,B2.REVERSED,B4.APPROVED,B4.STAGE
  FROM
    TBL_BATCHITEMSPLIT        B1
    LEFT JOIN TBL_BATCHITEM   B2 ON (B1.SERIALNUMBER = B2.SERIALNUMBER) AND (B1.RECEIPTNO = B2.RECEIPTNO) AND (B1.ADMITNAME = B2.ADMITNAME)
    LEFT JOIN TBL_BATCHHEADER B4 ON (B2.ADMITNAME = B4.ADMITNAME)
    LEFT JOIN cte_sourcecode  B6 ON (B1.SOURCECODE = B6.SOURCECODE)
  WHERE
      (B2.REVERSED IS NULL OR NOT (B2.REVERSED=1 OR B2.REVERSED=-1))
)
----------------------------------------------------------------------------------------------------------------------------
select
  T1.DATEOFPAYMENT
  ,SUM(T1.PAYMENTAMOUNT) AS TOTAL
  ,T1.SOURCECODE, T1.ADES1, T1.ADES2
from
  cte_payment T1
where
  T1.SOURCECODE IN ('18EAS02','18FEB02')
group by 
  T1.DATEOFPAYMENT, T1.SOURCECODE , T1.ADES1, T1.ADES2
order by 
  T1.DATEOFPAYMENT


