-- Contains a row for each object that is a trigger,
-- with a type of TR or TA. DML trigger names are schema-scoped and, therefore, are visible in sys.objects.
-- DDL trigger names are scoped by the parent entity and are only visible in this view.

-- The parent_class and name columns uniquely identify the trigger in the database.
SELECT 
  *
FROM
  sys.triggers