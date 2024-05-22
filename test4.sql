WITH 
    profile_summary AS (
        SELECT subproject_id, subproject_name, product_supply_type
        FROM F16_AMET_PROFILE_SUMMARY
        WHERE subproject_id IN (SELECT DISTINCT prod_swid FROM F16_AMET_COST_ARTIFACT)
    ),
    awcp AS (
        SELECT prod_swid, NVL(SUM(HOURS), 0) AS CURRENT_AWCP 
        FROM F16_AMET_COST_ARTIFACT
        WHERE COST_TYPE = 'A1'
        GROUP BY prod_swid
    ),
    rework_acwp AS (
        SELECT prod_swid, NVL(SUM(HOURS), 0) AS CURRENT_REWORK_ACWP 
        FROM F16_AMET_COST_ARTIFACT
        WHERE COST_TYPE = 'A1' AND SUBSTR(activity_code, 1, 1) = 'R'
        GROUP BY prod_swid
    ),
    bcwp AS (
        SELECT prod_swid, NVL(SUM(HOURS), 0) AS CURRENT_BCWP 
        FROM F16_AMET_COST_ARTIFACT
        WHERE COST_TYPE IN ('B3', 'P3')
        GROUP BY prod_swid
    ),
    bcws AS (
        SELECT prod_swid, NVL(SUM(HOURS), 0) AS FINAL_BCWS 
        FROM F16_AMET_COST_ARTIFACT
        WHERE COST_TYPE IN ('B2', 'PB')
        GROUP BY prod_swid
    ),
    rework_bcws AS (
        SELECT prod_swid, NVL(SUM(HOURS), 0) AS FINAL_REWORK_BCWS 
        FROM F16_AMET_COST_ARTIFACT
        WHERE COST_TYPE IN ('B2', 'PB') AND SUBSTR(activity_code, 1, 1) = 'R'
        GROUP BY prod_swid
    ),
    ctc AS (
        SELECT prod_swid, NVL(SUM(HOURS), 0) AS CURRENT_CTC 
        FROM F16_AMET_COST_ARTIFACT
        WHERE COST_TYPE IN ('F1', 'PF') AND MONTH_END >= TO_DATE('3/01/2024', 'MM/DD/YYYY')
        GROUP BY prod_swid
    ),
    rework_ctc AS (
        SELECT prod_swid, NVL(SUM(HOURS), 0) AS CURRENT_REWORK_CTC 
        FROM F16_AMET_COST_ARTIFACT
        WHERE COST_TYPE IN ('F1', 'PF') AND SUBSTR(activity_code, 1, 1) = 'R' AND MONTH_END >= TO_DATE('3/01/2024', 'MM/DD/YYYY')
        GROUP BY prod_swid
    ),
    profile_data AS (
        SELECT 
            SUBPROJECT_NAME,
            MAX(CASE WHEN PROFILE_DATA_NAME = 'Product - Prod Line' THEN PROFILE_VALUE_STRING ELSE NULL END) AS PROD_LINE,
            MAX(CASE WHEN PROFILE_DATA_NAME = 'Product - SW Domain' THEN PROFILE_VALUE_STRING ELSE NULL END) AS "DOMAIN",
            MAX(CASE WHEN PROFILE_DATA_NAME = 'Product - SW Domain Subtype' THEN PROFILE_VALUE_STRING ELSE NULL END) AS SUBDOMAIN,
            MAX(CASE WHEN PROFILE_DATA_NAME = 'Product - Saftey Level' THEN PROFILE_VALUE_STRING ELSE NULL END) AS SAFETY_LEVEL,
            MAX(CASE WHEN PROFILE_DATA_NAME = 'Estimate - Labor Hours' THEN PROFILE_VALUE_INT ELSE NULL END) AS ESTIMATED_LABOR_HOURS,
            MAX(CASE WHEN PROFILE_DATA_NAME = 'Rollup - Organization Specific Usage Detail' THEN PROFILE_VALUE_STRING ELSE NULL END) AS ROLLUP_ORG_SPEC_USAGE_DETAIL,
            MAX(CASE WHEN PROFILE_DATA_NAME = 'Subproject Status' THEN PROFILE_VALUE_STRING ELSE NULL END) AS SUBPROJECT_STATUS,
            MAX(CASE WHEN PROFILE_DATA_NAME = 'Subproject Characterization Status Date' THEN PROFILE_VALUE_DATE ELSE NULL END) AS ACTUAL_RELEASE_DATE,
            MAX(CASE WHEN PROFILE_DATA_NAME = 'Release - Actual Release Date (Production Equivalent)' THEN PROFILE_VALUE_DATE ELSE NULL END) AS END_DATE,
            MAX(CASE WHEN PROFILE_DATA_NAME = 'Estimate - Project Start Date' THEN PROFILE_VALUE_DATE ELSE NULL END) AS PROJECT_START_DATE
        FROM F16_AMET_PROFILE_DATA
        GROUP BY SUBPROJECT_NAME
    )

SELECT 
    ps.subproject_id AS SUBPROJECT_ID, 
    ps.subproject_name, 
    ps.product_supply_type,
    pd.PROD_LINE, 
    pd."DOMAIN", 
    pd.SUBDOMAIN, 
    pd.SAFETY_LEVEL, 
    pd.ESTIMATED_LABOR_HOURS, 
    pd.ROLLUP_ORG_SPEC_USAGE_DETAIL, 
    pd.PROJECT_START_DATE,
    awcp.CURRENT_AWCP, 
    rework_acwp.CURRENT_REWORK_ACWP, 
    bcwp.CURRENT_BCWP, 
    bcws.FINAL_BCWS, 
    rework_bcws.FINAL_REWORK_BCWS, 
    ctc.CURRENT_CTC, 
    rework_ctc.CURRENT_REWORK_CTC,
    pd.SUBPROJECT_STATUS, 
    pd.ACTUAL_RELEASE_DATE, 
    pd.END_DATE
FROM profile_summary ps
LEFT JOIN awcp ON ps.subproject_id = awcp.prod_swid
LEFT JOIN rework_acwp ON ps.subproject_id = rework_acwp.prod_swid
LEFT JOIN bcwp ON ps.subproject_id = bcwp.prod_swid
LEFT JOIN bcws ON ps.subproject_id = bcws.prod_swid
LEFT JOIN rework_bcws ON ps.subproject_id = rework_bcws.prod_swid
LEFT JOIN ctc ON ps.subproject_id = ctc.prod_swid
LEFT JOIN rework_ctc ON ps.subproject_id = rework_ctc.prod_swid
LEFT JOIN profile_data pd ON ps.subproject_name = pd.SUBPROJECT_NAME
ORDER BY 1;
