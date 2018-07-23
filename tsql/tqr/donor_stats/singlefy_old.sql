-- /*<BOFY>*/'2018/07/01'/*</BOFY>*/
-- /*<EOFY>*/'2019/06/30'/*</EOFY>*/
-- 0
-- /*<BOLFY>*/'2017/07/01'/*</BOLFY>*/
-- /*<EOLFY>*/'2017/06/30'/*</EOLFY>*/

SELECT t1.SerialNumber, 

tbl_contact.Contacttype AS Contacttype,
tbl_contact.GENDER AS Gender,
tbl_contact.Addressline4 AS State,
tbl_contact.Country AS Country,
tbl_contact.DateOfBirth AS DateOfBirth,

/* FY Version of total payments number */
COUNT(t1.SerialNumber) AS FYTotalNo,

/* LTD Version of total payments number */
(SELECT  COUNT(t4.SerialNumber) 
FROM  tbl_BATCHITEM AS t4
WHERE  (t4.reversed is null or not (t4.reversed = -1 or t4.reversed =1 or t4.reversed =2)) AND (t1.SerialNumber=t4.SerialNumber) AND (t4.DateOfPayment <= /*<EOFY>*/'2019/06/30'/*</EOFY>*/)
) AS LTDTotalNo,

/* Total Gift Amount Received in selected FY year  */
(SELECT 
SUM(t17.PaymentAmount)
FROM tbl_BATCHITEM AS t17
WHERE (t17.reversed is null or not (t17.reversed = -1 or t17.reversed =1)) AND (t1.SerialNumber=t17.SerialNumber) AND (t17.DateOfPayment BETWEEN /*<BOFY>*/'2018/07/01'/*</BOFY>*/ AND /*<EOFY>*/'2019/06/30'/*</EOFY>*/) AND (t1.SerialNumber=t17.SerialNumber)
)AS FYTotal, 

/* Total Gift Amount Received up to selected FY year  */
(SELECT  SUM(t5.PaymentAmount) 
FROM  tbl_BATCHITEM AS t5
WHERE  (t5.reversed is null or not (t5.reversed = -1 or t5.reversed =1)) AND (t1.SerialNumber=t5.SerialNumber) AND (t5.DateOfPayment <= /*<EOFY>*/'2019/06/30'/*</EOFY>*/)
) AS LTDTotal,

/* Average of Gift Amount Received in selected FY year  */
AVG(t1.PaymentAmount) AS FYAve,

/* Average of Gift Amount Received up to selected FY year  */
(SELECT  AVG(t6.PaymentAmount) 
FROM  tbl_BATCHITEM AS t6
WHERE  (t6.reversed is null or not (t6.reversed = -1 or t6.reversed =1 or t6.reversed =2)) AND (t1.SerialNumber=t6.SerialNumber) AND (t6.DateOfPayment <= /*<EOFY>*/'2019/06/30'/*</EOFY>*/)
) AS LTDAve,

/* Standard Devation of Gift Amount Received in selected FY year  */
STDEV(t1.PaymentAmount) AS FYStdev,

/* Standard Devation of Gift Amount Received up to selected FY year  */
(SELECT  STDEV(t7.PaymentAmount) 
FROM  tbl_BATCHITEM AS t7
WHERE  (t7.reversed is null or not (t7.reversed = -1 or t7.reversed =1 or t7.reversed =2)) AND (t1.SerialNumber=t7.SerialNumber) AND (t7.DateOfPayment <= /*<EOFY>*/'2019/06/30'/*</EOFY>*/)
) AS LTDStdev,

/* Mode Gift Amount Received up to selected FY year  */
(SELECT  TOP 1 t11.PaymentAmount
FROM  tbl_BATCHITEM AS t11
WHERE  (t11.reversed is null or not (t11.reversed = -1 or t11.reversed =1 or t11.reversed =2)) AND (t1.SerialNumber=t11.SerialNumber) AND (t11.DateOfPayment <= /*<EOFY>*/'2019/06/30'/*</EOFY>*/)
GROUP BY t11.PaymentAmount,t11.SerialNumber
ORDER BY t11.SerialNumber ASC,COUNT(t11.PaymentAmount) DESC,t11.PaymentAmount ASC
) AS LTDMode,

/* Meadian Gift Amount Received up to selected FY year  */
(
(
(SELECT TOP 1 t13.PaymentAmount
FROM(
  SELECT Top 50 Percent  t12.PaymentAmount
  FROM tbl_BATCHITEM AS t12
  WHERE  (t12.reversed is null or not (t12.reversed = -1 or t12.reversed =1 or t12.reversed =2)) AND (t12.SerialNumber=t1.SerialNumber) AND (t12.DateOfPayment <= /*<EOFY>*/'2019/06/30'/*</EOFY>*/)
  ORDER BY t12.PaymentAmount ASC
           ) AS t13
ORDER BY t13.PaymentAmount DESC
)
+
(SELECT TOP 1 t15.PaymentAmount
FROM(
  SELECT Top 50 Percent  t14.PaymentAmount
  FROM tbl_BATCHITEM AS t14
  WHERE  (t14.reversed is null or not (t14.reversed = -1 or t14.reversed =1 or t14.reversed =2)) AND (t14.SerialNumber=t1.SerialNumber) AND (t14.DateOfPayment <= /*<EOFY>*/'2019/06/30'/*</EOFY>*/)
  ORDER BY t14.PaymentAmount DESC
  ) AS t15
ORDER BY t15.PaymentAmount ASC
)
)/2) AS LTDMedian,


/* Smallest Gift Amount Received in selected FY year  */
MIN(t1.PaymentAmount) AS FYMinGift,

/* Smallest Gift Amount Received up to selected FY year  */
(SELECT  MIN(t8.PaymentAmount) 
FROM  tbl_BATCHITEM AS t8
WHERE  (t8.reversed is null or not (t8.reversed = -1 or t8.reversed =1 or t8.reversed =2)) AND (t1.SerialNumber=t8.SerialNumber) AND (t8.DateOfPayment <= /*<EOFY>*/'2019/06/30'/*</EOFY>*/)
) AS LTDMinGift,

/* Largest Gift Amount Received in selected FY year  */
MAX(t1.PaymentAmount) AS FYMaxGift,

/* Largest Gift Amount Received up to selected FY year  */
(SELECT  MAX(t9.PaymentAmount) 
FROM  tbl_BATCHITEM AS t9
WHERE  (t9.reversed is null or not (t9.reversed = -1 or t9.reversed =1 or t9.reversed =2)) AND (t1.SerialNumber=t9.SerialNumber) AND (t9.DateOfPayment <= /*<EOFY>*/'2019/06/30'/*</EOFY>*/)
) AS LTDMaxGift,

/* Date of first gift received in selected in history */
(SELECT  MIN(t10.DateOfPayment) 
FROM  tbl_BATCHITEM AS t10
WHERE  (t10.reversed is null or not (t10.reversed = -1 or t10.reversed =1 or t10.reversed =2)) AND (t1.SerialNumber=t10.SerialNumber) AND (t10.DateOfPayment <= /*<EOFY>*/'2019/06/30'/*</EOFY>*/)
) AS LTDFirstDate,

/* $Amount of first gift received in selected in history */
(SELECT  top 1 t2.PaymentAmount
FROM  tbl_BATCHITEM AS t2
WHERE  (t2.reversed is null or not (t2.reversed = -1 or t2.reversed =1 or t2.reversed =2)) AND (t1.SerialNumber=t2.SerialNumber)
ORDER BY t2.SerialNumber ASC,t2.DateOfPayment ASC
) AS LTDFirstGift,

/* Date of last gift received in selected FY year */
MAX(t1.DateOfPayment) AS FYLastDate,

/* $Amount of last gift received in selected FY year */
(SELECT  top 1 t3.PaymentAmount
FROM  tbl_BATCHITEM AS t3
WHERE  (t3.reversed is null or not (t3.reversed = -1 or t3.reversed =1 or t3.reversed =2)) AND (t1.SerialNumber=t3.SerialNumber) AND (t3.DateOfPayment BETWEEN /*<BOFY>*/'2018/07/01'/*</BOFY>*/ AND /*<EOFY>*/'2019/06/30'/*</EOFY>*/)
ORDER BY t3.SerialNumber ASC,t3.DateOfPayment DESC
) AS FYLastGift,

/* How many days a donor is connected to us since begining */
CAST( MAX(t1.DateOfPayment)-
(SELECT  MIN(t10.DateOfPayment) 
FROM  tbl_BATCHITEM AS t10
WHERE  (t10.reversed is null or not (t10.reversed = -1 or t10.reversed =1 or t10.reversed =2)) AND (t1.SerialNumber=t10.SerialNumber) AND (t10.DateOfPayment <= /*<EOFY>*/'2019/06/30'/*</EOFY>*/)
) AS numeric )/365
AS LengthOfSupportYrs,

/* How many Gifts they give last FY */
(SELECT 
COUNT(t16.SerialNumber)
FROM tbl_BATCHITEM AS t16
WHERE (t16.reversed is null or not (t16.reversed = -1 or t16.reversed =1 or t16.reversed =2)) AND (t1.SerialNumber=t16.SerialNumber) AND (t16.DateOfPayment BETWEEN /*<BOLFY>*/'2017/07/01'/*</BOLFY>*/  AND /*<EOLFY>*/'2017/06/30'/*</EOLFY>*/)
GROUP BY t16.SerialNumber
)AS LFYTotalNo,

tbl_contact.ANONYMOUS AS ANON,
tbl_contact.DONOTMAIL AS DONOTMAIL,
tbl_contact.DONOTMAILREASON AS DONOTMAILREASON,
tbl_contact.DONOTMAILFROM AS DONOTMAILFROM,
tbl_contact.DONOTMAILUNTIL AS DONOTMAILUNTIL,
tbl_contact.POSTCODE,tbl_contact.Primarycategory,

(SELECT  sum(t17.PAYMENTAMOUNT) 
FROM  Tbl_BATCHITEMSPLIT AS t17 LEFT JOIN tbl_BATCHITEM AS t18 
ON ((t17.SERIALNUMBER = t18.SERIALNUMBER) AND (t17.RECEIPTNO = t18.RECEIPTNO) AND (t17.ADMITNAME = t18.ADMITNAME))
WHERE  
(t18.reversed is null or not (t18.reversed = -1 or t18.reversed =1)) AND (t1.SerialNumber=t18.SerialNumber) AND (Tbl_Contact.SerialNumber=t18.SerialNumber) AND
(t17.SOURCECODE like 'LEG%' ) AND (t18.DateOfPayment BETWEEN /*<BOFY>*/'2018/07/01'/*</BOFY>*/ AND /*<EOFY>*/'2019/06/30'/*</EOFY>*/)
)AS LegacyPmt,

(SELECT TOP 1 [SEGMENT] = CONCAT(PARAMETERVALUE,'.',PARAMETERNOTE)
FROM TBL_CONTACTPARAMETER tmp
WHERE PARAMETERNAME = /*<SEG>*/'FY2019'/*</SEG>*/ AND (tmp.SERIALNUMBER = t1.SERIALNUMBER)
)
AS Segment


FROM tbl_BATCHITEM AS t1 LEFT JOIN tbl_Contact ON tbl_Contact.SERIALNUMBER = t1.SERIALNUMBER
WHERE (t1.reversed is null or not (t1.reversed = -1 or t1.reversed =1 or t1.reversed =2)) AND (t1.DateOfPayment BETWEEN /*<BOFY>*/'2018/07/01'/*</BOFY>*/ AND /*<EOFY>*/'2019/06/30'/*</EOFY>*/)


GROUP BY Tbl_Contact.SerialNumber,t1.SerialNumber,tbl_contact.Contacttype,tbl_contact.GENDER,tbl_contact.Addressline4,tbl_contact.Country, tbl_contact.DateOfBirth,tbl_contact.ANONYMOUS,tbl_contact.DONOTMAIL,tbl_contact.DONOTMAILREASON ,tbl_contact.DONOTMAILFROM,tbl_contact.DONOTMAILUNTIL,tbl_contact.POSTCODE,tbl_contact.Primarycategory

ORDER BY t1.SerialNumber