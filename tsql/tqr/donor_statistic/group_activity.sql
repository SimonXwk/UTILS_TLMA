select
  SerialNumber
  ,MonthOfPayment
  ,YEAR(MonthOfPayment) + CASE WHEN MONTH(MonthOfPayment) < 7 THEN 0 ELSE 1 END AS [FY]
  ,PaymentAmount
  ,PrimaryCategory
  , CASE WHEN TheState IS NULL OR RTRIM(TheState) = '' THEN 'MISSING'
    ELSE 
        CASE WHEN UPPER(LTRIM(RTRIM(TheState))) IN ('VIC','NSW','SA','QLD','WA','TAS','ACT','NT')
        THEN UPPER(LTRIM(RTRIM(TheState)))
        ELSE 'O/S'
        END
    END AS [TheState]
  ,case when MonthGapPrev is null then 'New' else case when MonthGapPrev<=12 then 'Continuing' else 'Reactivated'end end as Status
from
  (
  select M1.SerialNumber, M1.MonthOfPayment, M1.PaymentAmount, M1.PrimaryCategory, M1.TheState, DateDiff(m,M2.MonthOfPayment,M1.MonthOfPayment) as MonthGapPrev
  from 
    (
    Select SerialNumber, MonthOfPayment, sum(PaymentAmount) as PaymentAmount, PrimaryCategory, TheState, Rank() over (partition by SerialNumber order by MonthOfPayment desc) RankOrder
    from (
      select SerialNumber, DateAdd(d,-1*Day(DateOfPayment)+1,DateOfPayment) as MonthOfPayment, PrimaryCategory, TheState, PaymentAmount 
      from View_Payments
      where
        (reversed is null or reversed=0 or abs(reversed)=2)
        and PrimaryCategory in ('CC List','Group','Church')
        and Anonymous <> 'Yes'
        and (DateOfPayment between DATEFROMPARTS(IIF(MONTH(CURRENT_TIMESTAMP)<7,YEAR(CURRENT_TIMESTAMP),YEAR(CURRENT_TIMESTAMP)+1)-1-/*<FYOFFSET1>*/5/*</FYOFFSET1>*/,7,1) and CURRENT_TIMESTAMP)) P1
    group by SerialNumber, MonthOfPayment, PrimaryCategory, TheState) M1
    left join
    (
    Select SerialNumber, MonthOfPayment, Rank() over (partition by SerialNumber order by MonthOfPayment desc) RankOrder
    from (
    select SerialNumber, DateAdd(d,-1*Day(DateOfPayment)+1,DateOfPayment) as MonthOfPayment
    from View_Payments
    where 
      (reversed is null or reversed=0 or abs(reversed)=2)
      and PrimaryCategory in ('CC List','Group','Church')
      and Anonymous <> 'Yes'
      ) P2
    group by SerialNumber, MonthOfPayment) M2
      on M1.SerialNumber=M2.SerialNumber and M2.RankOrder=M1.RankOrder+1) R
union
select
  SerialNumber
  ,LapseDate
  ,YEAR(LapseDate) + CASE WHEN MONTH(LapseDate) < 7 THEN 0 ELSE 1 END AS [FY]
  ,Null as PaymentAmount,
  PrimaryCategory
  , CASE WHEN TheState IS NULL OR RTRIM(TheState) = '' THEN 'MISSING'
    ELSE 
        CASE WHEN UPPER(LTRIM(RTRIM(TheState))) IN ('VIC','NSW','SA','QLD','WA','TAS','ACT','NT')
        THEN UPPER(LTRIM(RTRIM(TheState)))
        ELSE 'O/S'
        END
    END AS [TheState]
  ,'Lapsed' as Status
from
 (
  select M1.SerialNumber, M1.MonthOfPayment, M1.PrimaryCategory, M1.TheState, DateDiff(m,M1.MonthOfPayment,M3.MonthOfPayment) as MonthGapNext, DateAdd(m,13,M1.MonthOfPayment) as LapseDate
  from (
      Select SerialNumber, MonthOfPayment, PrimaryCategory, TheState, Rank() over (partition by SerialNumber order by MonthOfPayment desc) RankOrder
      from (select SerialNumber, DateAdd(d,-1*Day(DateOfPayment)+1,DateOfPayment) as MonthOfPayment, PrimaryCategory, TheState from View_Payments
      where 
        (reversed is null or reversed=0 or abs(reversed)=2)
        and PrimaryCategory in ('CC List','Group','Church')
        and Anonymous <> 'Yes'
        and (DateOfPayment between DATEFROMPARTS(IIF(MONTH(CURRENT_TIMESTAMP)<7,YEAR(CURRENT_TIMESTAMP),YEAR(CURRENT_TIMESTAMP)+1)-1-/*<FYOFFSET1>*/5/*</FYOFFSET1>*/,7,1) and CURRENT_TIMESTAMP)) P1
  group by SerialNumber, MonthOfPayment, PrimaryCategory, TheState) M1
  left join
  (
  Select SerialNumber, MonthOfPayment, Rank() over (partition by SerialNumber order by MonthOfPayment desc) RankOrder
  from (
    select SerialNumber, DateAdd(d,-1*Day(DateOfPayment)+1,DateOfPayment) as MonthOfPayment  from View_Payments
    where 
      (reversed is null or reversed=0 or abs(reversed)=2)
      and PrimaryCategory in ('CC List','Group','Church')
      and Anonymous <> 'Yes'
      ) P3 
    group by SerialNumber, MonthOfPayment) M3
  on M1.SerialNumber=M3.SerialNumber and M3.RankOrder=M1.RankOrder-1) R
 where (MonthGapNext is null and GetDate() >= LapseDate) or MonthGapNext > 12



-- CASE WHEN ISNULL(BATCHITEM.ADDRESSLINE1,'')<>'' 
-- THEN BATCHITEM.ADDRESSLINE4 
-- ELSE 
-- CASE WHEN dbo.CONTACT.SERIALNUMBER IS NOT NULL 
-- THEN 
--     CASE WHEN ISNULL(dbo.CONTACT.DEFAULTADDRESS, '') = '' 
--     THEN dbo.CONTACT.ADDRESSLINE4 
--     ELSE dbo.CONTACTATTRIBUTE.DEFAULTADDRESSLINE4 
--     END 
-- ELSE
--     CASE WHEN ISNULL(dbo.CONTACTSARCHIVE.DEFAULTADDRESS,
--                    '') = '' 
--     THEN dbo.CONTACTSARCHIVE.ADDRESSLINE4 
--     ELSE dbo.CONTACTATTRIBUTESARCHIVE.DEFAULTADDRESSLINE4 
--     END
-- END 
-- END AS TheState, 