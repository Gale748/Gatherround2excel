SELECT 
    q0.subproject_id AS SUBPROJECT_ID,
    q0.subproject_name,
    q1.activity_code,
    CASE 
        WHEN q1.cost_type = 'A1' THEN 'ACWP'
        WHEN q1.cost_type IN ('PF', 'F1') THEN 'BCWP'
        WHEN q1.cost_type IN ('PB', 'B2') THEN 'BCWS'
        WHEN q1.cost_type IN ('P3', 'B3') THEN 'CTC'
        ELSE 'OTHER'
    END AS COST_TYPE_NAME,
    COALESCE(SUM(hours), 0) AS HOURS
FROM 
    F16_AMET_PROFILE_SUMMARY q0
JOIN 
    F16_AMET_COST_ARTIFACT q1 ON q0.subproject_id = q1.prod_swid
WHERE
    q0.subproject_id IN (SELECT DISTINCT prod_swid FROM F16_AMET_COST_ARTIFACT)
GROUP BY
    q0.subproject_id,
    q0.subproject_name,
    q1.activity_code,
    CASE 
        WHEN q1.cost_type = 'A1' THEN 'ACWP'
        WHEN q1.cost_type IN ('PF', 'F1') THEN 'BCWP'
        WHEN q1.cost_type IN ('PB', 'B2') THEN 'BCWS'
        WHEN q1.cost_type IN ('P3', 'B3') THEN 'CTC'
        ELSE 'OTHER'
    END
ORDER BY
    q0.subproject_id,
    q1.activity_code,
    COST_TYPE_NAME;
