SELECT t1.SerialNumber,

  tbl_contact.Contacttype AS Contacttype,
  tbl_contact.GENDER AS Gender,
  tbl_contact.Addressline4 AS State,
  tbl_contact.Country AS Country,
  tbl_contact.DateOfBirth AS DateOfBirth,

  /* FY Version of total payments number */
  COUNT(t1.SerialNumber) AS TotalNo,

  /* Total Gift Amount Received in selected FY year  */
  (SELECT
    SUM(t17.PaymentAmount)
  FROM tbl_BATCHITEM AS t17
  WHERE (t17.reversed is null) AND (t1.SerialNumber=t17.SerialNumber)
)AS Total,

  /* Average of Gift Amount Received in selected FY year  */
  AVG(t1.PaymentAmount) AS Ave,

  /* Standard Devation of Gift Amount Received in selected FY year  */
  STDEV(t1.PaymentAmount) AS Stdev,

  /* Mode Gift Amount Received up to selected FY year  */
  (SELECT TOP 1
    t11.PaymentAmount
  FROM tbl_BATCHITEM AS t11
  WHERE  (t11.reversed is null or not (t11.reversed = -1 or t11.reversed =1 or t11.reversed =2)) AND (t1.SerialNumber=t11.SerialNumber)
  GROUP BY t11.PaymentAmount,t11.SerialNumber
  ORDER BY t11.SerialNumber ASC,COUNT(t11.PaymentAmount) DESC,t11.PaymentAmount ASC
) AS Mode,

  /* Meadian Gift Amount Received up to selected FY year  */
  (
(
(SELECT TOP 1
    t13.PaymentAmount
  FROM(
  SELECT Top 50 Percent
      t12.PaymentAmount
    FROM tbl_BATCHITEM AS t12
    WHERE  (t12.reversed is null or not (t12.reversed = -1 or t12.reversed =1 or t12.reversed =2)) AND (t12.SerialNumber=t1.SerialNumber)
    ORDER BY t12.PaymentAmount ASC
           ) AS t13
  ORDER BY t13.PaymentAmount DESC
)
+
(SELECT TOP 1
    t15.PaymentAmount
  FROM(
  SELECT Top 50 Percent
      t14.PaymentAmount
    FROM tbl_BATCHITEM AS t14
    WHERE  (t14.reversed is null or not (t14.reversed = -1 or t14.reversed =1 or t14.reversed =2)) AND (t14.SerialNumber=t1.SerialNumber)
    ORDER BY t14.PaymentAmount DESC
  ) AS t15
  ORDER BY t15.PaymentAmount ASC
)
)/2) AS Median,

  /* Smallest Gift Amount Received in selected FY year  */
  MIN(t1.PaymentAmount) AS MinGift,

  /* Largest Gift Amount Received in selected FY year  */
  MAX(t1.PaymentAmount) AS MaxGift,

  /* Date of first gift received in selected in history */
  (SELECT MIN(t10.DateOfPayment)
  FROM tbl_BATCHITEM AS t10
  WHERE  (t10.reversed is null or not (t10.reversed = -1 or t10.reversed =1 or t10.reversed =2)) AND (t1.SerialNumber=t10.SerialNumber)
) AS FirstDate,

  /* $Amount of first gift received in selected in history */
  (SELECT top 1
    t2.PaymentAmount
  FROM tbl_BATCHITEM AS t2
  WHERE  (t2.reversed is null or not (t2.reversed = -1 or t2.reversed =1 or t2.reversed =2)) AND (t1.SerialNumber=t2.SerialNumber)
  ORDER BY t2.SerialNumber ASC,t2.DateOfPayment ASC
) AS FirstGift,

  /* Date of last gift received in selected FY year */
  MAX(t1.DateOfPayment) AS LastDate,

  /* $Amount of last gift received in selected FY year */
  (SELECT top 1
    t3.PaymentAmount
  FROM tbl_BATCHITEM AS t3
  WHERE  (t3.reversed is null or not (t3.reversed = -1 or t3.reversed =1 or t3.reversed =2)) AND (t1.SerialNumber=t3.SerialNumber)
  ORDER BY t3.SerialNumber ASC,t3.DateOfPayment DESC
) AS LastGift,

  /* How many days a donor is connected to us since begining */
  CAST( MAX(t1.DateOfPayment)-
(SELECT MIN(t10.DateOfPayment)
  FROM tbl_BATCHITEM AS t10
  WHERE  (t10.reversed is null or not (t10.reversed = -1 or t10.reversed =1 or t10.reversed =2)) AND (t1.SerialNumber=t10.SerialNumber)
) AS numeric )/365
AS LengthOfSupportYrs,

  tbl_contact.ANONYMOUS AS ANON,
  tbl_contact.DONOTMAIL AS DONOTMAIL,
  tbl_contact.DONOTMAILREASON AS DONOTMAILREASON,
  tbl_contact.DONOTMAILFROM AS DONOTMAILFROM,
  tbl_contact.DONOTMAILUNTIL AS DONOTMAILUNTIL,
  tbl_contact.Primarycategory,

  (SELECT sum(t15.PAYMENTAMOUNT)
  FROM Tbl_BATCHITEMSPLIT AS t15 LEFT JOIN tbl_BATCHITEM AS t16
    ON ((t15.SERIALNUMBER = t16.SERIALNUMBER) AND (t15.RECEIPTNO = t16.RECEIPTNO) AND (t15.ADMITNAME = t16.ADMITNAME))

  WHERE  
(t16.reversed is null or not (t16.reversed = -1 or t16.reversed =1)) AND (t1.SerialNumber=t16.SerialNumber) AND
    (Tbl_Contact.SerialNumber=t16.SerialNumber) AND
    (t15.SOURCECODE like 'LEG%' ) AND
    (
t16.DATEOFPAYMENT BETWEEN 
Cast(cast( (CASE WHEN  Month(GETDATE() )<7  THEN Year(GETDATE() )-11  ELSE Year(GETDATE() )-10  END) as VARCHAR(4)) +'/7/1' AS DATETIME)  
AND 
Cast(cast( (CASE WHEN  Month(GETDATE() )<7  THEN Year(GETDATE() )-1  ELSE Year(GETDATE() )  END) as VARCHAR(4)) +'/6/30'AS DATETIME) 
)
)AS LegacyPmt


FROM tbl_BATCHITEM AS t1 LEFT JOIN tbl_Contact ON tbl_Contact.SERIALNUMBER = t1.SERIALNUMBER
WHERE (t1.reversed is null or not (t1.reversed = -1 or t1.reversed =1 or t1.reversed =2)) AND (t1.DateOfPayment BETWEEN Cast(cast( (CASE WHEN  Month(GETDATE() )<7  THEN Year(GETDATE() )-11 ELSE Year(GETDATE() )-10  END) as VARCHAR(4)) +'/7/1'
AS DATETIME)  AND Cast(cast( (CASE WHEN  Month(GETDATE() )<7  THEN Year(GETDATE() )-1  ELSE Year(GETDATE() )  END) as VARCHAR(4)) +'/6/30'
AS DATETIME) )

GROUP BY t1.SerialNumber,tbl_contact.Contacttype,tbl_contact.GENDER,tbl_contact.Addressline4,tbl_contact.Country, tbl_contact.DateOfBirth,tbl_contact.ANONYMOUS,tbl_contact.DONOTMAIL,tbl_contact.DONOTMAILREASON ,tbl_contact.DONOTMAILFROM,tbl_contact.DONOTMAILUNTIL,tbl_contact.Primarycategory,Tbl_Contact.SerialNumber


ORDER BY t1.SerialNumber