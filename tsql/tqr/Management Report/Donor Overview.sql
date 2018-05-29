with
digit AS ( 
  select * from (values (0),(1),(2),(3),(4),(5),(6),(7),(8),(9)) as numbers(x) 
)
-- --------------------------------------------------------------
,generation_def AS(
  select *
  from (values
    (0, 1824, 'Ancient', '01.AC')
    ,(1825, 1844, 'Early Colonial', '02.EC')
    ,(1845, 1864, 'Mid Colonial', '03.MC')
    ,(1865, 1884, 'Late Colonial', '04.LC')
    ,(1885, 1904, 'Hard Timers', '05.HT')
    ,(1905, 1924, 'Federation', '06.F')
    ,(1925, 1944, 'Silent', '07.S')
    ,(1945, 1964, 'Baby Boomers', '08.BB')
    ,(1965, 1979, 'Generation X', '09.X')
    ,(1980, 1994, 'Generation Y', '10.Y')
    ,(1995, 2009, 'Generation Z', '11.Z')
    ,(2010, 9999, 'Millenials', '12.M') ) AS generation(y1,y2,gen,gen_abr) 
)
-- --------------------------------------------------------------
,generation AS (
  select GENERANTION_YEAR=n.x, GENERANTION=g.gen, GENERANTION_ABBRE=g.gen_abr
  from
    (select x=1000*o1000.x + 100*o100.x + 10*o10.x + o1.x
    from digit o1, digit o10, digit o100, digit o1000 ) n
      left join generation_def g on (n.x>=g.y1 and n.x<=g.y2)
  where n.x BETWEEN 1000 and YEAR(CURRENT_TIMESTAMP)
)
-- --------------------------------------------------------------
,cte_legit_contact as (
  select * 
    ,FULLNAME=CONCAT(
        IIF(RTRIM(ISNULL(TITLE,''))='','',RTRIM(TITLE)+' ')
        ,IIF(RTRIM(ISNULL(FIRSTNAME,''))='','',RTRIM(FIRSTNAME)+' ')
        ,IIF(RTRIM(ISNULL(OTHERINITIAL,''))='','',RTRIM(OTHERINITIAL)+' '),KEYNAME)
  from
    TBL_CONTACT 
    left join generation ON (YEAR(DATEOFBIRTH) = GENERANTION_YEAR)
  where CONTACTTYPE <> 'ADDRESS'
)
-- --------------------------------------------------------------
,cte_contact_first_date_trx as (
  select SERIALNUMBER
    ,min(DATEOFPAYMENT) as first_date_trx
    ,year(min(DATEOFPAYMENT))+IIF(month(min(DATEOFPAYMENT))<7,0,1) as first_fy_trx
  from Tbl_BATCHITEM
  where REVERSED is null or not (REVERSED=-1 or REVERSED=1 or REVERSED=2)
  group by SERIALNUMBER
)
-- --------------------------------------------------------------
,cte_payments as (
  select bsp.SERIALNUMBER,bsp.ADMITNAME,bsp.RECEIPTNO,bsp.PAYMENTAMOUNT,bsp.GSTAMOUNT,bsp.SOURCECODE,bsp.SOURCECODE2,bsp.DESTINATIONCODE,bsp.DESTINATIONCODE2
    ,sc.SOURCETYPE,sc.ADDITIONALCODE1 as QBCODE,sc.ADDITIONALCODE5 as QBCLASS,sc.ADDITIONALCODE3 as CAMPAIGN
    ,btm.DATEOFPAYMENT,btm.PAYMENTTYPE,btm.CREDITCARDTYPE,btm.MANUALRECEIPTNO,btm.REVERSED
    ,bpl.PLEDGEID,bpl.PLEDGELINENO
    ,bhr.APPROVED
    ,IIF(sc.SOURCETYPE like 'Merch%','Merch Platform','NonM Platform') as PLATFORM
    ,TRXID = dense_rank() over (partition by bsp.SERIALNUMBER order by btm.DATEOFPAYMENT asc, concat(bsp.ADMITNAME,bsp.RECEIPTNO) asc)
  from 
    Tbl_BATCHITEMSPLIT bsp
    left join Tbl_SOURCECODE sc on (bsp.SOURCECODE=sc.SOURCECODE)
    left join Tbl_BATCHITEMPLEDGE bpl on (bsp.ADMITNAME=bpl.ADMITNAME and bsp.SERIALNUMBER=bpl.SERIALNUMBER and bsp.RECEIPTNO=bpl.RECEIPTNO and bsp.LINEID=bpl.LINEID)
    left join Tbl_BATCHITEM btm on (bsp.ADMITNAME=btm.ADMITNAME and bsp.SERIALNUMBER=btm.SERIALNUMBER and bsp.RECEIPTNO=btm.RECEIPTNO)
    left join Tbl_BATCHHEADER bhr on (bsp.ADMITNAME=bhr.ADMITNAME)
  WHERE
    btm.REVERSED is null or not(btm.REVERSED=-1 or btm.REVERSED=1) and bhr.STAGE = 'Batch Approved'
)

-- *****************************************************************
select
  *
from 
  cte_payments
where 
  SERIALNUMBER = '0301167'
-- for xml path('')

