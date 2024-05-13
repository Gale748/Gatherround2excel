SELECT
    q0.subproject_id AS SUBPROJECT_ID,
    q0.subproject_name,
    COALESCE(SUM(CASE WHEN activity_code = 'PC' THEN hours ELSE 0 END), 0) AS ACT_PC,
    COALESCE(SUM(CASE WHEN activity_code = 'PJ' THEN hours ELSE 0 END), 0) AS ACT_PJ,
    COALESCE(SUM(CASE WHEN activity_code = 'RA' THEN hours ELSE 0 END), 0) AS ACT_RA,
    COALESCE(SUM(CASE WHEN activity_code = 'RC' THEN hours ELSE 0 END), 0) AS ACT_RC,
    COALESCE(SUM(CASE WHEN activity_code = 'RD' THEN hours ELSE 0 END), 0) AS ACT_RD,
    COALESCE(SUM(CASE WHEN activity_code = 'RR' THEN hours ELSE 0 END), 0) AS ACT_RR,
    COALESCE(SUM(CASE WHEN activity_code = 'RT' THEN hours ELSE 0 END), 0) AS ACT_RT,
    COALESCE(SUM(CASE WHEN activity_code = 'RU' THEN hours ELSE 0 END), 0) AS ACT_RU,
    COALESCE(SUM(CASE WHEN activity_code = 'WA' THEN hours ELSE 0 END), 0) AS ACT_WA,
    COALESCE(SUM(CASE WHEN activity_code = 'WC' THEN hours ELSE 0 END), 0) AS ACT_WC,
    COALESCE(SUM(CASE WHEN activity_code = 'WD' THEN hours ELSE 0 END), 0) AS ACT_WD,
    COALESCE(SUM(CASE WHEN activity_code = 'WR' THEN hours ELSE 0 END), 0) AS ACT_WR,
    COALESCE(SUM(CASE WHEN activity_code = 'WT' THEN hours ELSE 0 END), 0) AS ACT_WT,
    COALESCE(SUM(CASE WHEN activity_code = 'WU' THEN hours ELSE 0 END), 0) AS ACT_WU,
    COALESCE(SUM(hours), 0) AS GRAND_TOTAL
FROM
    F16_AMET_PROFILE_SUMMARY q0
JOIN
    F16_AMET_COST_ARTIFACT q1 ON q0.subproject_id = q1.prod_swid
WHERE
    q0.subproject_id IN (SELECT DISTINCT prod_swid FROM F16_AMET_COST_ARTIFACT)
    AND q1.COST_TYPE = 'A1'
GROUP BY
    q0.subproject_id,
    q0.subproject_name
ORDER BY
    q0.subproject_id;
