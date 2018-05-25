-- Contains a row for each view object, with sys.objects.type = V.
SELECT
  *
FROM sys.views
ORDER BY create_date desc