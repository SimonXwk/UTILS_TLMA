
WITH
cte_campaigns AS (
  SELECT
    DISTINCT dbo.TBL_SOURCECODE.ADDITIONALCODE3 AS [CAMPAIGNCODE]
  FROM dbo.TBL_SOURCECODE
  WHERE 
    dbo.TBL_SOURCECODE.ADDITIONALCODE3 IS NOT NULL
    AND dbo.TBL_SOURCECODE.ADDITIONALCODE3 LIKE '%.%'
)
-- --------------------------------------------------
, cte_pledge_source_code AS (
  SELECT 
    dbo.Tbl_SOURCECODE.SOURCECODE
    , dbo.Tbl_SOURCECODE.ADDITIONALCODE3 AS [CAMPAIGNCODE]
  FROM dbo.Tbl_SOURCECODE
  WHERE dbo.TBL_SOURCECODE.SOURCETYPE LIKE '%Sponsorship'
)
-- --------------------------------------------------

select * from cte_pledge_source_code