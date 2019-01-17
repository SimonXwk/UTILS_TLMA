USE thankQ4_Reporter

SELECT * FROM sys.views
WHERE name LIKE 'SV_%' OR name LIKE 'View_%'
ORDER BY create_date DESC



-- select CAST(SERIALNUMBER AS VARCHAR(7)),PAYMENTAMOUNT,DATEOFPAYMENT,REVERSED,TRX_ID,TRX_ID2,TRX_ID3,lead(TRX_ID) OVER(ORDER BY TRX_ID )
-- from MyView_Payments
-- where SERIALNUMBER = '0303725'
-- order by DATEOFPAYMENT





-- SELECT
--   lag(v) OVER (ORDER BY v),
--   v, 
--   lead(v) OVER (ORDER BY v)
-- FROM (
--   VALUES (1), (2), (3), (4)
-- ) t(v)