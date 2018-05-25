SELECT
  database_id
  ,DB_NAME(database_id) AS [databases]
  ,create_date
FROM sys.databases