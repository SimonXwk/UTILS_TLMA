SELECT
  SCHEMA_NAME(o.schema_id) as schema_name
  ,o.type_desc
  ,o.name AS obj_name
  ,c.name AS col_name
  ,TYPE_NAME(c.system_type_id) AS sys_type
  ,TYPE_NAME(c.user_type_id) AS user_type
  ,c.max_length
  ,IIF(c.is_nullable=0,'False','') AS nullable
  ,c.precision
FROM sys.objects AS o
  JOIN sys.columns AS c ON o.object_id = c.object_id
WHERE o.name = 'TBL_BATCHITEMSPLIT' 