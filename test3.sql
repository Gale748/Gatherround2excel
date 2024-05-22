WITH 
-- Subproject information
profile_summary AS (
    SELECT subproject_id, subproject_name, product_supply_type
    FROM F16_AMET_PROFILE_SUMMARY
    WHERE subproject_id IN (SELECT DISTINCT prod_swid FROM F16_AMET_COST_ARTIFACT)
),
-- Aggregated cost data
cost_data AS (
    SELECT 
        prod_swid,
        cost_type,
        activity_code,
        NVL(SUM(CASE WHEN COST_TYPE='A1' THEN HOURS ELSE 0 END), 0) AS current_awcp,
        NVL(SUM(CASE WHEN COST_TYPE='A1' AND SUBSTR(activity_code,1,1) IN ('R') THEN HOURS ELSE 0 END), 0) AS current_rework_acwp,
        NVL(SUM(CASE WHEN COST_TYPE IN ('B3','P3') THEN HOURS ELSE 0 END), 0) AS current_bcwp,
        NVL(SUM(CASE WHEN COST_TYPE IN ('B2','PB') THEN HOURS ELSE 0 END), 0) AS final_bcws,
        NVL(SUM(CASE WHEN COST_TYPE IN ('B2','PB') AND SUBSTR(activity_code,1,1) IN ('R') THEN HOURS ELSE 0 END), 0) AS final_rework_bcws,
        NVL(SUM(CASE WHEN COST_TYPE IN ('F1','PF') AND MONTH_END >= TO_DATE('3/01/2024','MM/DD/YYYY') THEN HOURS ELSE 0 END), 0) AS current_ctc,
        NVL(SUM(CASE WHEN COST_TYPE IN ('F1','PF') AND SUBSTR(activity_code,1,1) IN ('R') AND MONTH_END >= TO_DATE('3/01/2024','MM/DD/YYYY') THEN HOURS ELSE 0 END), 0) AS current_rework_ctc
    FROM F16_AMET_COST_ARTIFACT
    GROUP BY prod_swid, cost_type, activity_code
),
-- Profile data
profile_data AS (
    SELECT subproject_name,
           MAX(CASE WHEN profile_data_name = 'Product - Prod Line' THEN profile_value_string END) AS prod_line,
           MAX(CASE WHEN profile_data_name = 'Product - SW Domain' THEN profile_value_string END) AS domain,
           MAX(CASE WHEN profile_data_name = 'Product - SW Domain Subtype' THEN profile_value_string END) AS subdomain,
           MAX(CASE WHEN profile_data_name = 'Product - Saftey Level' THEN profile_value_string END) AS safety_level,
           MAX(CASE WHEN profile_data_name = 'Estimate - Labor Hours' THEN profile_value_int END) AS estimated_labor_hours,
           MAX(CASE WHEN profile_data_name = 'Rollup - Organization Specific Usage Detail' THEN profile_value_string END) AS rollup_org_spec_usage_detail,
           MAX(CASE WHEN profile_data_name = 'Subproject Status' THEN profile_value_string END) AS subproject_status,
           MAX(CASE WHEN profile_data_name = 'Subproject Characterization Status Date' THEN profile_value_date END) AS actual_release_date,
           MAX(CASE WHEN profile_data_name = 'Release - Actual Release Date (Production Equivalent)' THEN profile_value_date END) AS end_date,
           MAX(CASE WHEN profile_data_name = 'Estimate - Project Start Date' THEN profile_value_date END) AS project_start_date
    FROM F16_AMET_PROFILE_DATA
    GROUP BY subproject_name
)

SELECT 
    q0.subproject_id AS subproject_id,
    q0.subproject_name,
    q0.product_supply_type,
    pd.prod_line,
    pd.domain,
    pd.subdomain,
    pd.safety_level,
    pd.estimated_labor_hours,
    pd.rollup_org_spec_usage_detail,
    pd.project_start_date,
    NVL(cd.current_awcp, 0) AS current_awcp,
    NVL(cd.current_rework_acwp, 0) AS current_rework_acwp,
    NVL(cd.current_bcwp, 0) AS current_bcwp,
    NVL(cd.final_bcws, 0) AS final_bcws,
    NVL(cd.final_rework_bcws, 0) AS final_rework_bcws,
    NVL(cd.current_ctc, 0) AS current_ctc,
    NVL(cd.current_rework_ctc, 0) AS current_rework_ctc,
    cd.cost_type,
    cd.activity_code,
    pd.subproject_status,
    pd.actual_release_date,
    pd.end_date
FROM profile_summary q0
LEFT JOIN cost_data cd ON q0.subproject_id = cd.prod_swid
LEFT JOIN profile_data pd ON q0.subproject_name = pd.subproject_name
ORDER BY q0.subproject_id, cd.cost_type, cd.activity_code;
