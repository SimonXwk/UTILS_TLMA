
with
-- --------------------------------------------------------------
ones AS (
 SELECT * FROM (VALUES (0), (1), (2), (3), (4),(5), (6), (7), (8), (9)) AS numbers(x)
)
-- --------------------------------------------------------------
,generation_def AS (
   SELECT * FROM (
      VALUES
      (0,1824,'Ancient','01.AC')
      ,(1825,1844,'Early Colonial','02.EC')
      ,(1845,1864,'Mid Colonial','03.MC')
      ,(1865,1884,'Late Colonial','04.LC')
      ,(1885,1904,'Hard Timers','05.HT')
      ,(1905,1924,'Federation','06.F')
      ,(1925,1944,'Silent','07.S')
      ,(1945,1964,'Baby Boomers','08.BB')
      ,(1965,1979,'Generation X','09.X')
      ,(1980,1994,'Generation Y','10.Y')
      ,(1995,2009,'Generation Z','11.Z')
      ,(2010,9999,'Millenials','12.M')
      ) AS generation(y1,y2,gen,gen_abr) 
)
-- --------------------------------------------------------------
,cte_generation AS(
SELECT cy=n.x,gen=g.gen,gen_abr=g.gen_abr
FROM 
   (SELECT x=1000*o1000.x + 100*o100.x + 10*o10.x + o1.x FROM ones o1, ones o10, ones o100, ones o1000 ) n 
   LEFT JOIN generation_def g on(n.x>=g.y1 AND n.x<=g.y2)
WHERE n.x BETWEEN 1 AND YEAR(CURRENT_TIMESTAMP)
)
-- --------------------------------------------------------------
,cte_rex_customerid AS(
SELECT SERIALNUMBER, PARAMETERNAME, PARAMETERVALUE
FROM TBL_CONTACTPARAMETER
WHERE PARAMETERNAME LIKE '%Customer%Number%'
)
-- --------------------------------------------------------------
,cte_tq_rex_id_count AS(
SELECT SERIALNUMBER,COUNT(PARAMETERVALUE) AS REXID_COUNT
FROM cte_rex_customerid
GROUP BY SERIALNUMBER
)
-- --------------------------------------------------------------
select
   t1.SERIALNUMBER
   ,r1.REXID_COUNT
   ,CAST(t1.PARAMETERVALUE AS VARCHAR(10)) AS CUSTOMERNUMBER
   ,PRIMARYCATEGORY = IIF(c1.PRIMARYCATEGORY IS NULL OR RTRIM(c1.PRIMARYCATEGORY )='','[DNF]',c1.PRIMARYCATEGORY)
   ,GENDER = IIF(c1.GENDER IS NULL OR RTRIM(c1.GENDER )='','[DNF]',c1.GENDER)
   ,CONTACTTYPE = IIF(c1.CONTACTTYPE IS NULL OR RTRIM(c1.CONTACTTYPE )='','[DNF]',c1.CONTACTTYPE)
   ,SOURCE = IIF(c1.SOURCE IS NULL OR RTRIM(c1.SOURCE )='','[DNF]',c1.SOURCE)
   ,STATE=CAST(IIF(UPPER(LTRIM(RTRIM(c1.ADDRESSLINE4))) IN ('VIC','NSW','SA','QLD','WA','TAS','ACT','NT'),UPPER(LTRIM(RTRIM(c1.ADDRESSLINE4))),IIF(RTRIM(c1.ADDRESSLINE4)='' OR c1.ADDRESSLINE4 IS NULL,'[DNF]','O/S')) AS VARCHAR(6))
   ,AGE = IIF(c1.DATEOFBIRTH IS NULL OR RTRIM(c1.DATEOFBIRTH )='',NULL,(DATEDIFF(day,c1.DATEOFBIRTH,CURRENT_TIMESTAMP)+1)/365)
   ,GENERATION=IIF((c1.DATEOFBIRTH IS NULL OR RTRIM(c1.DATEOFBIRTH)=''),'[DNF]',t3.GEN)
   ,NAME = CONCAT(
      IIF(RTRIM(c1.TITLE)='' OR c1.TITLE IS NULL,'',RTRIM(c1.TITLE)+' ')
      ,IIF(RTRIM(c1.FIRSTNAME)='' OR c1.FIRSTNAME IS NULL,'',RTRIM(c1.FIRSTNAME)+' ')
      ,IIF(RTRIM(c1.OTHERINITIAL)='' OR c1.OTHERINITIAL IS NULL,'',RTRIM(c1.OTHERINITIAL)+' '),c1.KEYNAME)
   ,DONOTCALL = IIF(c2.DONOTCALL=-1,'YES',NULL)
   ,DONOTEMAIL = IIF(c2.DONOTEMAIL=-1,'YES',NULL)
   ,DONOTMAIL = IIF(c1.DONOTMAIL=-1,'YES',NULL)
   ,MAJORDONOR = IIF(c1.MAJORDONOR = -1,'YES',NULL)
   ,CURRENT_TIMESTAMP AS CURRENTTIMESTAMP
from
   cte_rex_customerid t1
   left join cte_tq_rex_id_count r1 ON (t1.SERIALNUMBER = r1.SERIALNUMBER)
   left join TBL_CONTACT c1 ON (t1.SERIALNUMBER = c1.SERIALNUMBER)
   LEFT JOIN TBL_CONTACTATTRIBUTE c2 ON (c1.SERIALNUMBER = c2.SERIALNUMBER)
   left join cte_generation t3 ON (YEAR(c1.DATEOFBIRTH) = t3.cy)