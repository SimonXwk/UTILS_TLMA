-- Returns a row for each column of an object that has columns, such as views or tables. The following is a list of object types that have columns:
-- Table-valued assembly functions (FT)
-- Inline table-valued SQL functions (IF)
-- Internal tables (IT)
-- System tables (S)
-- Table-valued SQL functions (TF)
-- User tables (U)
-- Views (V)

SELECT
  *
FROM sys.columns