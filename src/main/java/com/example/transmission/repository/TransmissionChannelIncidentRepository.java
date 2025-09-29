package com.example.transmission.repository;

import com.example.transmission.domain.ProblemFinalEndDate;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;
import org.springframework.stereotype.Repository;

import java.time.LocalDate;
import java.util.List;
import java.util.Map;

@Repository
public interface TransmissionChannelIncidentRepository extends JpaRepository<ProblemFinalEndDate, Long> {

    @Query(value =
            "SELECT\n" +
                    "    DATE(dt.create_date) AS day,\n" +
                    "    p.target_name AS parentName,\n" +
                    "    rt.target_name AS incidentName,\n" +
                    "    rt.kpi AS kpi_progress,\n" +
                    "    COUNT(DISTINCT dt.problem_id) AS tong_xu_ly,\n" +
                    "    COUNT(DISTINCT CASE\n" +
                    "        WHEN COALESCE(gp.process_time_total_gnoc, gp.process_time_total)\n" +
                    "                 <= (CASE WHEN rpg.name = 'Sự cố Kênh truyền quốc tế' THEN 5 ELSE 3 END)\n" +
                    "            THEN dt.problem_id END\n" +
                    "    ) AS xu_li_trong_han_3h,\n" +
                    "    COUNT(DISTINCT CASE\n" +
                    "        WHEN COALESCE(gp.process_time_total_gnoc, gp.process_time_total) <= 24\n" +
                    "            THEN dt.problem_id END\n" +
                    "    ) AS xu_li_trong_han_24h,\n" +
                    "    COUNT(DISTINCT CASE\n" +
                    "        WHEN COALESCE(gp.process_time_total_gnoc, gp.process_time_total) <= 48\n" +
                    "            THEN dt.problem_id END\n" +
                    "    ) AS xu_li_trong_han_48h\n" +
                    "FROM d_vcoc_problem_group rpg\n" +
                    "     JOIN d_vcoc_problem_group rpgc ON rpg.prob_group_id = rpgc.parent_id\n" +
                    "     JOIN rp_target_mapping rtm ON rtm.parent_prob_group_id = rpg.prob_group_id AND rtm.prob_group_id = rpgc.prob_group_id\n" +
                    "     LEFT JOIN d_vcoc_problem_final_end_date dt ON rtm.prob_type_id = dt.prob_type_id\n" +
                    "           AND DATE(dt.create_date) BETWEEN DATE_SUB(:currentDate, INTERVAL DAY(:currentDate) - 1 DAY) AND :currentDate\n" +
                    "           AND dt.status <> '4'\n" +
                    "     LEFT JOIN d_vcoc_rpt_problem_gpdn_process gp ON dt.problem_id = gp.problem_id\n" +
                    "           AND (gp.cust_accept_level IS NULL OR gp.cust_accept_level NOT IN (\n" +
                    "               'Không Happy call KH – Đóng hủy KN theo yêu cầu CC',\n" +
                    "               'Không Happy call KH – Đóng trùng.')\n" +
                    "           )\n" +
                    "     JOIN rp_target rt ON rt.id = rtm.target_id\n" +
                    "     JOIN rp_target p ON rt.parent_id = p.id\n" +
                    "     JOIN rp_target_of_table rtot ON rt.id = rtot.target_id\n" +
                    "     JOIN rp_table_target_config rttc ON rtot.table_target_config_id = rttc.id\n" +
                    "     JOIN rp_report_config c ON c.complaint_group_id = rpg.prob_group_id\n" +
                    "     JOIN rp_report_inf i ON c.report_id = i.id\n" +
                    "WHERE i.report_name = 'Báo cáo kênh truyền ngày'\n" +
                    "  AND rttc.report_table_name = 'Tiến độ xử lí'\n" +
                    "GROUP BY DATE(dt.create_date), rt.target_name, rt.kpi, p.target_name\n" +
                    "ORDER BY day, parentName, incidentName",
            nativeQuery = true)
    List<Map<String, Object>> getOnTimeHandle(@Param("currentDate") LocalDate currentDate);

    @Query(value =
            "SELECT\n" +
                    "    d.parentName,\n" +
                    "    d.day,\n" +
                    "    d.kpiChild,\n" +
                    "    d.totalSubcriberChild,\n" +
                    "    d.kpiParent,\n" +
                    "    d.totalSubcriberParent,\n" +
                    "    d.incidentName,\n" +
                    "    d.totalPerDay\n" +
                    "FROM (\n" +
                    "         SELECT\n" +
                    "             DATE(dt.create_date) AS day,\n" +
                    "             rt.kpi AS kpiChild,\n" +
                    "             rt.total_subcriber AS totalSubcriberChild,\n" +
                    "             p.total_subcriber AS totalSubcriberParent,\n" +
                    "             p.kpi AS kpiParent,\n" +
                    "             p.target_name AS parentName,\n" +
                    "             rt.target_name AS incidentName,\n" +
                    "             COUNT(dt.problem_id) AS totalPerDay\n" +
                    "         FROM d_vcoc_problem_group rpg\n" +
                    "                  JOIN d_vcoc_problem_group rpgc ON rpg.prob_group_id = rpgc.parent_id\n" +
                    "                  JOIN rp_target_mapping rtm ON rtm.parent_prob_group_id = rpg.prob_group_id\n" +
                    "                       AND rtm.prob_group_id = rpgc.prob_group_id\n" +
                    "                  LEFT JOIN d_vcoc_problem_final_end_date dt ON rtm.prob_type_id = dt.prob_type_id\n" +
                    "                       AND DATE(dt.create_date) BETWEEN DATE_SUB(:currentDate, INTERVAL DAY(:currentDate) - 1 DAY)\n" +
                    "                                                    AND :currentDate\n" +
                    "                       AND dt.status <> '4'\n" +
                    "                  LEFT JOIN d_vcoc_rpt_problem_gpdn_process gp ON dt.problem_id = gp.problem_id\n" +
                    "                       AND (gp.cust_accept_level IS NULL OR gp.cust_accept_level NOT IN (\n" +
                    "                            'Không Happy call KH – Đóng hủy KN theo yêu cầu CC',\n" +
                    "                            'Không Happy call KH – Đóng trùng.')\n" +
                    "                       )\n" +
                    "                  JOIN rp_target rt ON rt.id = rtm.target_id\n" +
                    "                  JOIN rp_target p ON rt.parent_id = p.id\n" +
                    "                  JOIN rp_target_of_table rtot ON rt.id = rtot.target_id\n" +
                    "                  JOIN rp_table_target_config rttc ON rtot.table_target_config_id = rttc.id\n" +
                    "                  JOIN rp_report_config c ON c.complaint_group_id = rpg.prob_group_id\n" +
                    "                  JOIN rp_report_inf i ON c.report_id = i.id\n" +
                    "         WHERE i.report_name = 'Báo cáo kênh truyền ngày'\n" +
                    "           AND rttc.report_table_name = 'TỶ LỆ PAKH'\n" +
                    "         GROUP BY DATE(dt.create_date), rt.target_name, rt.order_no\n" +
                    "     ) d\n" +
                    "ORDER BY d.day, d.incidentName;\n",
            nativeQuery = true)
    List<Map<String, Object>> getCategorySummary(@Param("currentDate") LocalDate currentDate);

    @Query(value =
            "WITH base AS (\n" +
                    "    SELECT\n" +
                    "        dt.create_date AS day,\n" +
                    "        p.target_name AS parentName,\n" +
                    "        rt.target_name AS incidentName,\n" +
                    "        COUNT(dt.problem_id) AS numberOfSC\n" +
                    "    FROM d_vcoc_problem_group rpg\n" +
                    "             JOIN d_vcoc_problem_group rpgc\n" +
                    "                  ON rpg.prob_group_id = rpgc.parent_id\n" +
                    "             JOIN rp_target_mapping rtm\n" +
                    "                  ON rtm.parent_prob_group_id = rpg.prob_group_id\n" +
                    "                      AND rtm.prob_group_id = rpgc.prob_group_id\n" +
                    "             LEFT JOIN d_vcoc_problem_final_end_date dt\n" +
                    "                       ON rtm.prob_type_id = dt.prob_type_id\n" +
                    "                           AND DATE(dt.create_date) BETWEEN DATE_SUB(:currentDate, INTERVAL DAY(:currentDate) - 1 DAY)\n" +
                    "                              AND :currentDate\n" +
                    "                                       AND dt.status <> '4'"+
                    "             LEFT JOIN d_vcoc_rpt_problem_gpdn_process gp\n" +
                    "                       ON dt.problem_id = gp.problem_id\n" +
                    "                           AND (\n" +
                    "                              gp.cust_accept_level IS NULL\n" +
                    "                                  OR gp.cust_accept_level NOT IN (\n" +
                    "                                                                  'Không Happy call KH – Đóng hủy KN theo yêu cầu CC',\n" +
                    "                                                                  'Không Happy call KH – Đóng trùng.'\n" +
                    "                                  )\n" +
                    "                              )\n" +
                    "             JOIN rp_target rt\n" +
                    "                  ON rt.id = rtm.target_id\n" +
                    "             LEFT JOIN rp_target p\n" +
                    "                       ON rt.parent_id = p.id\n" +
                    "             JOIN rp_target_of_table rtot\n" +
                    "                  ON rt.id = rtot.target_id\n" +
                    "             JOIN rp_table_target_config rttc\n" +
                    "                  ON rtot.table_target_config_id = rttc.id\n" +
                    "             JOIN rp_report_config c\n" +
                    "                  ON c.complaint_group_id = rpg.prob_group_id\n" +
                    "             JOIN rp_report_inf i\n" +
                    "                  ON c.report_id = i.id\n" +
                    "    WHERE i.report_name = 'Báo cáo kênh truyền ngày'\n" +
                    "      AND rttc.report_table_name = 'PAKH THEO DỊCH VỤ'\n" +
                    "    GROUP BY\n" +
                    "        dt.create_date,\n" +
                    "        p.target_name,\n" +
                    "        rt.target_name\n" +
                    ")\n" +
                    "\n" +
                    "SELECT\n" +
                    "    day,\n" +
                    "    parentName,\n" +
                    "    incidentName,\n" +
                    "    numberOfSC,\n" +
                    "\n" +
                    "    SUM(CASE WHEN parentName = 'SC' THEN numberOfSC ELSE 0 END)\n" +
                    "        OVER (PARTITION BY day) AS total_SC,\n" +
                    "    SUM(CASE WHEN parentName = 'HT' THEN numberOfSC ELSE 0 END)\n" +
                    "        OVER (PARTITION BY day) AS total_HT,\n" +
                    "    ROUND(\n" +
                    "            SUM(CASE WHEN parentName = 'SC' THEN numberOfSC ELSE 0 END)\n" +
                    "                OVER () / DAY(:currentDate), 2\n" +
                    "    ) AS avg_SC,\n" +
                    "    ROUND(\n" +
                    "            SUM(CASE WHEN parentName = 'HT' THEN numberOfSC ELSE 0 END)\n" +
                    "                OVER () / DAY(:currentDate), 2\n" +
                    "    ) AS avg_HT,\n" +
                    "    ROUND(\n" +
                    "            SUM(numberOfSC) OVER () / DAY(:currentDate), 2\n" +
                    "    ) AS avg_totalAll,\n" +
                    "    ROUND(\n" +
                    "            SUM(numberOfSC) OVER (PARTITION BY incidentName)\n" +
                    "                / DAY(:currentDate), 2\n" +
                    "    ) AS avg_per_incident\n" +
                    "\n" +
                    "FROM base\n" +
                    "ORDER BY\n" +
                    "    day,\n" +
                    "    parentName,\n" +
                    "    incidentName;\n",
            nativeQuery = true)
    List<Map<String, Object>> getDailyStats(@Param("currentDate") LocalDate currentDate);

    @Query(value =
            "WITH total_sltb AS ( \n" +
                    "    SELECT SUM(t.total_subcriber) AS total_sltb\n" +
                    "    FROM rp_target t\n" +
                    "    WHERE t.parent_id IN (\n" +
                    "        SELECT id \n" +
                    "        FROM rp_target \n" +
                    "        WHERE target_name = 'Toàn quốc'\n" +
                    "    )\n" +
                    "),\n" +
                    "base AS ( \n" +
                    "    SELECT \n" +
                    "        DATE(dt.create_date) AS day, \n" +
                    "        lf.location_name AS province, \n" +
                    "        COUNT(DISTINCT dt.problem_id) AS slpa, \n" +
                    "        rt.total_subcriber AS sltb\n" +
                    "    FROM d_vcoc_problem_group rpg\n" +
                    "    JOIN d_vcoc_problem_group rpgc \n" +
                    "         ON rpg.prob_group_id = rpgc.parent_id\n" +
                    "    JOIN rp_target_mapping rtm \n" +
                    "         ON rtm.parent_prob_group_id = rpg.prob_group_id\n" +
                    "        AND rtm.prob_group_id = rpgc.prob_group_id\n" +
                    "    LEFT JOIN d_vcoc_problem_final_end_date dt \n" +
                    "         ON rtm.prob_type_id = dt.prob_type_id\n" +
                    "        AND dt.province = lf.location_code\n" +
                    "        AND DATE(dt.create_date) BETWEEN DATE_SUB(:currentDate, INTERVAL DAY(:currentDate) - 1 DAY) \n" +
                    "                                     AND :currentDate\n" +
                    "        AND dt.status <> '4'\n" +
                    "    LEFT JOIN d_vcoc_rpt_problem_gpdn_process gp \n" +
                    "         ON dt.problem_id = gp.problem_id\n" +
                    "        AND gp.cust_accept_level NOT IN (\n" +
                    "            'Không Happy call KH – Đóng hủy KN theo yêu cầu CC', \n" +
                    "            'Không Happy call KH – Đóng trùng.'\n" +
                    "        )\n" +
                    "    LEFT JOIN d_location_full lf \n" +
                    "         ON dt.province = lf.location_code\n" +
                    "    LEFT JOIN rp_target rt \n" +
                    "         ON rt.id = rtm.target_id\n" +
                    "        AND rt.parent_id IN (\n" +
                    "            SELECT id FROM rp_target WHERE target_name = 'Toàn quốc'\n" +
                    "        )\n" +
                    "    WHERE EXISTS (\n" +
                    "              SELECT 1\n" +
                    "              FROM rp_report_config c\n" +
                    "              JOIN rp_report_inf i \n" +
                    "                   ON c.report_id = i.id\n" +
                    "              WHERE c.complaint_group_id = rpg.prob_group_id\n" +
                    "                AND i.report_name = 'Báo cáo kênh truyền ngày'\n" +
                    "          )\n" +
                    "      AND EXISTS (\n" +
                    "              SELECT 1\n" +
                    "              FROM rp_target_of_table rtot\n" +
                    "              JOIN rp_table_target_config rttc \n" +
                    "                   ON rtot.table_target_config_id = rttc.id\n" +
                    "              WHERE rtot.target_id = rt.id\n" +
                    "                AND rttc.report_table_name = 'TỶ LỆ PAKH THEO TỈNH'\n" +
                    "          )\n" +
                    "    GROUP BY DATE(dt.create_date), lf.location_name, rt.total_subcriber\n" +
                    ") \n" +
                    "SELECT \n" +
                    "    day, \n" +
                    "    province, \n" +
                    "    slpa, \n" +
                    "    sltb, \n" +
                    "    CASE \n" +
                    "        WHEN sltb IS NULL OR sltb = 0 THEN NULL\n" +
                    "        ELSE ROUND(slpa * 1000.0 / sltb, 2) \n" +
                    "    END AS tlpa\n" +
                    "FROM base\n" +
                    "UNION ALL\n" +
                    "SELECT \n" +
                    "    day, \n" +
                    "    'Toàn quốc' AS province, \n" +
                    "    SUM(slpa), \n" +
                    "    (SELECT total_sltb FROM total_sltb), \n" +
                    "    ROUND(SUM(slpa) * 1000.0 / NULLIF((SELECT total_sltb FROM total_sltb), 0), 2) \n" +
                    "FROM base\n" +
                    "GROUP BY day\n" +
                    "ORDER BY day, province;\n",
            nativeQuery = true)
    List<Map<String, Object>> getComplaintRateAndTotalSubscribers(@Param("currentDate") LocalDate currentDate);

    @Query(value =
                    "WITH data AS (\n" +
                            "    SELECT\n" +
                            "        DATE(dt.create_date) AS ngay,\n" +
                            "        lf.location_name AS tinh,\n" +
                            "        COUNT(DISTINCT dt.problem_id) AS tong_sc_da_xu_ly, \n" +
                            "        COUNT(\n" +
                            "            DISTINCT CASE\n" +
                            "                WHEN (\n" +
                            "                    (gp.prob_group_child_name = 'Sự cố Kênh truyền quốc tế'\n" +
                            "                     AND COALESCE(gp.process_time_total_gnoc, gp.process_time_total) <= 5)\n" +
                            "                    OR\n" +
                            "                    (COALESCE(gp.process_time_total_gnoc, gp.process_time_total) <= 3)\n" +
                            "                )\n" +
                            "                THEN dt.problem_id\n" +
                            "            END\n" +
                            "        ) AS da_xu_ly_3h\n" +
                            "    FROM d_vcoc_problem_group rpg\n" +
                            "    JOIN d_vcoc_problem_group rpgc\n" +
                            "        ON rpg.prob_group_id = rpgc.parent_id\n" +
                            "    JOIN rp_target_mapping rtm\n" +
                            "        ON rtm.parent_prob_group_id = rpg.prob_group_id\n" +
                            "       AND rtm.prob_group_id = rpgc.prob_group_id\n" +
                            "    JOIN rp_target rt \n" +
                            "        ON rt.id = rtm.target_id\n" +
                            "    JOIN d_location_full lf \n" +
                            "        ON lf.location_name = rt.target_name\n" +
                            "    LEFT JOIN d_vcoc_problem_final_end_date dt\n" +
                            "        ON rtm.prob_type_id = dt.prob_type_id\n" +
                            "       AND dt.province = lf.location_code\n" +
                            "       AND DATE(dt.create_date) BETWEEN DATE_SUB(@currentDate, INTERVAL DAY(@currentDate) - 1 DAY)\n" +
                            "                                   AND @currentDate\n" +
                            "       AND dt.status <> '4'\n" +
                            "    LEFT JOIN d_vcoc_rpt_problem_gpdn_process gp\n" +
                            "        ON dt.problem_id = gp.problem_id\n" +
                            "       AND (gp.cust_accept_level IS NULL\n" +
                            "            OR gp.cust_accept_level NOT IN (\n" +
                            "                'Không Happy call KH – Đóng hủy KN theo yêu cầu CC',\n" +
                            "                'Không Happy call KH – Đóng trùng.'\n" +
                            "            ))\n" +
                            "    JOIN rp_target_of_table rtot \n" +
                            "        ON rt.id = rtot.target_id\n" +
                            "    JOIN rp_table_target_config rttc \n" +
                            "        ON rtot.table_target_config_id = rttc.id\n" +
                            "    JOIN rp_report_config c \n" +
                            "        ON c.complaint_group_id = rpg.prob_group_id\n" +
                            "    JOIN rp_report_inf i \n" +
                            "        ON c.report_id = i.id\n" +
                            "    WHERE i.report_name = 'Báo cáo kênh truyền ngày'\n" +
                            "      AND rttc.report_table_name = 'TLXL trong 3h theo tỉnh'\n" +
                            "    GROUP BY DATE(dt.create_date), lf.location_name\n" +
                            ")\n" +
                            "SELECT\n" +
                            "    ngay,\n" +
                            "    tinh,\n" +
                            "    tong_sc_da_xu_ly,\n" +
                            "    da_xu_ly_3h,\n" +
                            "    ROUND(da_xu_ly_3h * 100.0 / NULLIF(tong_sc_da_xu_ly, 0), 2) AS ty_le_3h\n" +
                            "FROM data\n" +
                            "\n" +
                            "UNION ALL\n" +
                            "\n" +
                            "SELECT\n" +
                            "    ngay,\n" +
                            "    'Toàn quốc' AS tinh,\n" +
                            "    SUM(tong_sc_da_xu_ly),\n" +
                            "    SUM(da_xu_ly_3h),\n" +
                            "    ROUND(SUM(da_xu_ly_3h) * 100.0 / NULLIF(SUM(tong_sc_da_xu_ly), 0), 2) AS ty_le_3h\n" +
                            "FROM data\n" +
                            "GROUP BY ngay\n" +
                            "ORDER BY ngay, tinh;\n", nativeQuery = true)
    List<Map<String, Object>> getHandleRate3h(@Param("currentDate") LocalDate currentDate);

    @Query(value =
            "WITH data AS (\n" +
                    "    SELECT\n" +
                    "        DATE(dt.create_date) AS ngay,\n" +
                    "        lf.location_name AS tinh,\n" +
                    "        COUNT(dt.problem_id) AS tong_sc_da_xu_ly,\n" +
                    "        SUM(\n" +
                    "            CASE\n" +
                    "                WHEN COALESCE(gp.process_time_total_gnoc, gp.process_time_total) <= 24\n" +
                    "                THEN 1 ELSE 0\n" +
                    "            END\n" +
                    "        ) AS da_xu_ly_24h\n" +
                    "    FROM d_vcoc_problem_group rpg\n" +
                    "    JOIN d_vcoc_problem_group rpgc\n" +
                    "        ON rpg.prob_group_id = rpgc.parent_id\n" +
                    "    JOIN rp_target_mapping rtm\n" +
                    "        ON rtm.parent_prob_group_id = rpg.prob_group_id\n" +
                    "       AND rtm.prob_group_id = rpgc.prob_group_id\n" +
                    "    JOIN rp_target rt \n" +
                    "        ON rt.id = rtm.target_id\n" +
                    "    JOIN d_location_full lf \n" +
                    "        ON lf.location_name = rt.target_name\n" +
                    "    LEFT JOIN d_vcoc_problem_final_end_date dt\n" +
                    "        ON rtm.prob_type_id = dt.prob_type_id\n" +
                    "       AND dt.province = lf.location_code\n" +
                    "       AND DATE(dt.create_date) BETWEEN DATE_SUB(:currentDate, INTERVAL DAY(:currentDate) - 1 DAY)\n" +
                    "                              AND :currentDate\n" +
                    "                                       AND dt.status <> '4'"+
                    "    LEFT JOIN d_vcoc_rpt_problem_gpdn_process gp\n" +
                    "        ON dt.problem_id = gp.problem_id\n" +
                    "       AND (gp.cust_accept_level IS NULL\n" +
                    "            OR gp.cust_accept_level NOT IN (\n" +
                    "                'Không Happy call KH – Đóng hủy KN theo yêu cầu CC',\n" +
                    "                'Không Happy call KH – Đóng trùng.'\n" +
                    "            ))\n" +
                    "    JOIN rp_target_of_table rtot \n" +
                    "        ON rt.id = rtot.target_id\n" +
                    "    JOIN rp_table_target_config rttc \n" +
                    "        ON rtot.table_target_config_id = rttc.id\n" +
                    "    JOIN rp_report_config c \n" +
                    "        ON c.complaint_group_id = rpg.prob_group_id\n" +
                    "    JOIN rp_report_inf i \n" +
                    "        ON c.report_id = i.id\n" +
                    "    WHERE i.report_name = 'Báo cáo kênh truyền ngày'\n" +
                    "      AND rttc.report_table_name = 'TLXL trong 24h theo tỉnh'\n" +
                    "    GROUP BY DATE(dt.create_date), lf.location_name\n" +
                    "),\n" +
                    "luy_ke AS (\n" +
                    "    SELECT\n" +
                    "        ngay,\n" +
                    "        tinh,\n" +
                    "        tong_sc_da_xu_ly,\n" +
                    "        da_xu_ly_24h,\n" +
                    "        SUM(da_xu_ly_24h) OVER (PARTITION BY tinh ORDER BY ngay) AS luy_ke_24h,\n" +
                    "        SUM(tong_sc_da_xu_ly) OVER (PARTITION BY tinh ORDER BY ngay) AS tong_sc_luy_ke\n" +
                    "    FROM data\n" +
                    ")\n" +
                    "SELECT\n" +
                    "    ngay,\n" +
                    "    tinh,\n" +
                    "    tong_sc_da_xu_ly,\n" +
                    "    da_xu_ly_24h,\n" +
                    "    ROUND(da_xu_ly_24h * 100.0 / NULLIF(tong_sc_da_xu_ly, 0), 2) AS ty_le_24h,\n" +
                    "    luy_ke_24h,\n" +
                    "    ROUND(luy_ke_24h * 100.0 / NULLIF(tong_sc_luy_ke, 0), 2) AS ty_le_luy_ke_24h,\n" +
                    "    tong_sc_luy_ke\n" +
                    "FROM luy_ke\n" +
                    "\n" +
                    "UNION ALL\n" +
                    "\n" +
                    "SELECT\n" +
                    "    ngay,\n" +
                    "    'Toàn quốc' AS tinh,\n" +
                    "    SUM(tong_sc_da_xu_ly),\n" +
                    "    SUM(da_xu_ly_24h),\n" +
                    "    ROUND(SUM(da_xu_ly_24h) * 100.0 / NULLIF(SUM(tong_sc_da_xu_ly), 0), 2) AS ty_le_24h,\n" +
                    "    SUM(SUM(da_xu_ly_24h)) OVER (ORDER BY ngay) AS luy_ke_24h,\n" +
                    "    ROUND(\n" +
                    "        SUM(SUM(da_xu_ly_24h)) OVER (ORDER BY ngay) * 100.0\n" +
                    "        / NULLIF(SUM(SUM(tong_sc_da_xu_ly)) OVER (ORDER BY ngay), 0),\n" +
                    "        2\n" +
                    "    ) AS ty_le_luy_ke_24h,\n" +
                    "    SUM(SUM(tong_sc_da_xu_ly)) OVER (ORDER BY ngay) AS tong_sc_luy_ke\n" +
                    "FROM data\n" +
                    "GROUP BY ngay\n" +
                    "ORDER BY ngay, tinh;\n", nativeQuery = true)
    List<Map<String, Object>> getHandleRate24h(@Param("currentDate") LocalDate currentDate);

    @Query(value =
            "WITH data AS (\n" +
                    "    SELECT\n" +
                    "        DATE(dt.create_date) AS ngay,\n" +
                    "        lf.location_name AS tinh,\n" +
                    "        COUNT(dt.problem_id) AS tong_sc_da_xu_ly,\n" +
                    "        SUM(\n" +
                    "            CASE\n" +
                    "                WHEN COALESCE(gp.process_time_total_gnoc, gp.process_time_total) <= 48\n" +
                    "                THEN 1 ELSE 0\n" +
                    "            END\n" +
                    "        ) AS da_xu_ly_48h\n" +
                    "    FROM d_vcoc_problem_group rpg\n" +
                    "    JOIN d_vcoc_problem_group rpgc\n" +
                    "        ON rpg.prob_group_id = rpgc.parent_id\n" +
                    "    JOIN rp_target_mapping rtm\n" +
                    "        ON rtm.parent_prob_group_id = rpg.prob_group_id\n" +
                    "       AND rtm.prob_group_id = rpgc.prob_group_id\n" +
                    "    JOIN rp_target rt \n" +
                    "        ON rt.id = rtm.target_id\n" +
                    "    JOIN d_location_full lf \n" +
                    "        ON lf.location_name = rt.target_name\n" +
                    "    LEFT JOIN d_vcoc_problem_final_end_date dt\n" +
                    "        ON rtm.prob_type_id = dt.prob_type_id\n" +
                    "       AND dt.province = lf.location_code\n" +
                    "       AND DATE(dt.create_date) BETWEEN DATE_SUB(:currentDate, INTERVAL DAY(:currentDate) - 1 DAY)\n" +
                    "                              AND :currentDate\n" +
                    "                                       AND dt.status <> '4'"+
                    "    LEFT JOIN d_vcoc_rpt_problem_gpdn_process gp\n" +
                    "        ON dt.problem_id = gp.problem_id\n" +
                    "       AND (gp.cust_accept_level IS NULL\n" +
                    "            OR gp.cust_accept_level NOT IN (\n" +
                    "                'Không Happy call KH – Đóng hủy KN theo yêu cầu CC',\n" +
                    "                'Không Happy call KH – Đóng trùng.'\n" +
                    "            ))\n" +
                    "    JOIN rp_target_of_table rtot \n" +
                    "        ON rt.id = rtot.target_id\n" +
                    "    JOIN rp_table_target_config rttc \n" +
                    "        ON rtot.table_target_config_id = rttc.id\n" +
                    "    JOIN rp_report_config c \n" +
                    "        ON c.complaint_group_id = rpg.prob_group_id\n" +
                    "    JOIN rp_report_inf i \n" +
                    "        ON c.report_id = i.id\n" +
                    "    WHERE i.report_name = 'Báo cáo kênh truyền ngày'\n" +
                    "      AND rttc.report_table_name = 'TLXL trong 48h theo tỉnh'\n" +
                    "    GROUP BY DATE(dt.create_date), lf.location_name\n" +
                    "),\n" +
                    "luy_ke AS (\n" +
                    "    SELECT\n" +
                    "        ngay,\n" +
                    "        tinh,\n" +
                    "        tong_sc_da_xu_ly,\n" +
                    "        da_xu_ly_48h,\n" +
                    "        SUM(da_xu_ly_48h) OVER (PARTITION BY tinh ORDER BY ngay) AS luy_ke_48h,\n" +
                    "        SUM(tong_sc_da_xu_ly) OVER (PARTITION BY tinh ORDER BY ngay) AS tong_sc_luy_ke\n" +
                    "    FROM data\n" +
                    ")\n" +
                    "SELECT\n" +
                    "    ngay,\n" +
                    "    tinh,\n" +
                    "    tong_sc_da_xu_ly,\n" +
                    "    da_xu_ly_48h,\n" +
                    "    ROUND(da_xu_ly_48h * 100.0 / NULLIF(tong_sc_da_xu_ly, 0), 2) AS ty_le_48h,\n" +
                    "    luy_ke_48h,\n" +
                    "    ROUND(luy_ke_48h * 100.0 / NULLIF(tong_sc_luy_ke, 0), 2) AS ty_le_luy_ke_48h,\n" +
                    "    tong_sc_luy_ke\n" +
                    "FROM luy_ke\n" +
                    "\n" +
                    "UNION ALL\n" +
                    "\n" +
                    "SELECT\n" +
                    "    ngay,\n" +
                    "    'Toàn quốc' AS tinh,\n" +
                    "    SUM(tong_sc_da_xu_ly),\n" +
                    "    SUM(da_xu_ly_48h),\n" +
                    "    ROUND(SUM(da_xu_ly_48h) * 100.0 / NULLIF(SUM(tong_sc_da_xu_ly), 0), 2) AS ty_le_48h,\n" +
                    "    SUM(SUM(da_xu_ly_48h)) OVER (ORDER BY ngay) AS luy_ke_48h,\n" +
                    "    ROUND(\n" +
                    "        SUM(SUM(da_xu_ly_48h)) OVER (ORDER BY ngay) * 100.0\n" +
                    "        / NULLIF(SUM(SUM(tong_sc_da_xu_ly)) OVER (ORDER BY ngay), 0),\n" +
                    "        2\n" +
                    "    ) AS ty_le_luy_ke_48h,\n" +
                    "    SUM(SUM(tong_sc_da_xu_ly)) OVER (ORDER BY ngay) AS tong_sc_luy_ke\n" +
                    "FROM data\n" +
                    "GROUP BY ngay\n" +
                    "ORDER BY ngay, tinh;\n",
            nativeQuery = true)
    List<Map<String, Object>> getHandleRate48h(@Param("currentDate") LocalDate currentDate);

    @Query(value =
            "WITH data AS ( " +
                    "    SELECT " +
                    "        DATE(dt.create_date) AS ngay, " +
                    "        rt.target_name AS tinh, " +
                    "        COUNT(DISTINCT dt.problem_id) AS tong_sc_da_xu_ly, " +
                    "        COUNT(DISTINCT CASE " +
                    "            WHEN ( " +
                    "                (gp.prob_group_child_name = 'Sự cố Kênh truyền quốc tế' " +
                    "                 AND COALESCE(gp.process_time_total_gnoc, gp.process_time_total) <= 5) " +
                    "                OR (COALESCE(gp.process_time_total_gnoc, gp.process_time_total) <= 3) " +
                    "            ) " +
                    "            THEN dt.problem_id END) AS da_xu_ly_3h_vip " +
                    "    FROM d_vcoc_problem_group rpg " +
                    "    JOIN d_vcoc_problem_group rpgc " +
                    "        ON rpg.prob_group_id = rpgc.parent_id " +
                    "    JOIN rp_target_mapping rtm " +
                    "        ON rtm.parent_prob_group_id = rpg.prob_group_id " +
                    "       AND rtm.prob_group_id = rpgc.prob_group_id " +
                    "    JOIN rp_target rt " +
                    "        ON rt.id = rtm.target_id " +
                    "    JOIN d_location_full lf " +
                    "        ON lf.location_name = rt.target_name " +
                    "    LEFT JOIN d_vcoc_problem_final_end_date dt " +
                    "        ON rtm.prob_type_id = dt.prob_type_id " +
                    "       AND dt.province = lf.location_code " +
                    "       AND DATE(dt.create_date) BETWEEN DATE_SUB(@currentDate, INTERVAL DAY(@currentDate) - 1 DAY) " +
                    "                                    AND @currentDate " +
                    "       AND dt.status <> '4' " +
                    "    JOIN d_vcoc_problem_att_value_final_diemtrongyeu dty " +
                    "        ON dt.problem_id = dty.problem_id " +
                    "       AND dty.att_value IN ('1', '2') " +
                    "    LEFT JOIN d_vcoc_rpt_problem_gpdn_process gp " +
                    "        ON dt.problem_id = gp.problem_id " +
                    "       AND (gp.cust_accept_level IS NULL OR gp.cust_accept_level NOT IN ( " +
                    "            'Không Happy call KH – Đóng hủy KN theo yêu cầu CC', " +
                    "            'Không Happy call KH – Đóng trùng.' " +
                    "        )) " +
                    "    JOIN rp_target_of_table rtot " +
                    "        ON rt.id = rtot.target_id " +
                    "    JOIN rp_table_target_config rttc " +
                    "        ON rtot.table_target_config_id = rttc.id " +
                    "    JOIN rp_report_config c " +
                    "        ON c.complaint_group_id = rpg.prob_group_id " +
                    "    JOIN rp_report_inf i " +
                    "        ON c.report_id = i.id " +
                    "    WHERE i.report_name = 'Báo cáo kênh truyền ngày' " +
                    "      AND rttc.report_table_name = 'TLXL trong 3h vip theo tỉnh' " +
                    "    GROUP BY DATE(dt.create_date), rt.target_name " +
                    ") " +
                    "SELECT ngay, tinh, tong_sc_da_xu_ly, da_xu_ly_3h_vip, " +
                    "       ROUND(da_xu_ly_3h_vip * 100.0 / NULLIF(tong_sc_da_xu_ly, 0), 2) AS ty_le_3h_vip " +
                    "FROM data " +
                    "UNION ALL " +
                    "SELECT ngay, 'Toàn quốc' AS tinh, SUM(tong_sc_da_xu_ly), SUM(da_xu_ly_3h_vip), " +
                    "       ROUND(SUM(da_xu_ly_3h_vip) * 100.0 / NULLIF(SUM(tong_sc_da_xu_ly), 0), 2) AS ty_le_3h_vip " +
                    "FROM data " +
                    "GROUP BY ngay " +
                    "ORDER BY ngay, tinh; ",
            nativeQuery = true)
    List<Map<String, Object>> getHandleRate3hVip(@Param("currentDate") LocalDate currentDate);

    @Query(value =
            "SELECT\n" +
                    "    DATE(dt.create_date) AS ngay,\n" +
                    "    SUM(CAST(COALESCE(NULLIF(rpt.process_time_total_gnoc, ''), NULLIF(rpt.process_time_total, '')) AS DECIMAL(15,2))) AS tong_thoi_gian_xu_ly,\n" +
                    "    COUNT(dt.problem_id) AS so_luong_phan_anh,\n" +
                    "    ROUND(SUM(CAST(COALESCE(NULLIF(rpt.process_time_total_gnoc, ''), NULLIF(rpt.process_time_total, '')) AS DECIMAL(15,2))) / NULLIF(COUNT(DISTINCT dt.problem_id), 0), 2) AS thoi_gian_tb_xu_ly\n" +
                    "FROM d_vcoc_problem_group g\n" +
                    "         JOIN d_vcoc_problem_group gc\n" +
                    "              ON g.prob_group_id = gc.parent_id\n" +
                    "         JOIN rp_target_mapping rtm\n" +
                    "              ON rtm.parent_prob_group_id = g.prob_group_id\n" +
                    "                  AND rtm.prob_group_id = gc.prob_group_id\n" +
                    "         JOIN rp_target rt\n" +
                    "              ON rt.id = rtm.target_id\n" +
                    "         JOIN d_location_full lf\n" +
                    "              ON lf.location_name = rt.target_name\n" +
                    "         LEFT JOIN d_vcoc_problem_final_end_date dt\n" +
                    "              ON dt.prob_type_id = rtm.prob_type_id\n" +
                    "                 AND dt.province = lf.location_code\n" +
                    "                 AND DATE(dt.create_date) BETWEEN DATE_FORMAT(:currentDate, '%Y-%m-01') AND :currentDate\n" +
                    "                                       AND dt.status <> '4'"+
                    "         LEFT JOIN d_vcoc_rpt_problem_gpdn_process rpt\n" +
                    "              ON dt.problem_id = rpt.problem_id\n" +
                    "                 AND (rpt.cust_accept_level IS NULL OR rpt.cust_accept_level NOT IN (\n" +
                    "                                                         'Không Happy call KH – Đóng hủy KN theo yêu cầu CC',\n" +
                    "                                                         'Không Happy call KH – Đóng trùng.'\n" +
                    "                           ))\n" +
                    "         JOIN rp_target_of_table rtot\n" +
                    "              ON rt.id = rtot.target_id\n" +
                    "         JOIN rp_table_target_config rttc\n" +
                    "              ON rtot.table_target_config_id = rttc.id\n" +
                    "         JOIN rp_report_config c\n" +
                    "              ON c.complaint_group_id = g.prob_group_id\n" +
                    "         JOIN rp_report_inf i\n" +
                    "              ON i.id = c.report_id\n" +
                    "WHERE i.report_name = 'Báo cáo kênh truyền ngày'\n" +
                    "  AND rttc.report_table_name = 'KPI VTS'\n" +
                    "GROUP BY DATE(dt.create_date)\n" +
                    "ORDER BY ngay;\n",
            nativeQuery = true)
    List<Map<String, Object>> getAvgHandleTime(@Param("currentDate") LocalDate currentDate);

    @Query(value =
            "SELECT\n" +
                    "    DATE(dt.create_date) AS ngay,\n" +
                    "    SUM(CASE\n" +
                    "            WHEN rpt.cust_accept_level IN ('KH đồng ý', 'KH Hài lòng', 'KH chấp nhận') THEN 1\n" +
                    "            ELSE 0\n" +
                    "        END) AS tu_so_hai_long,\n" +
                    "    SUM(CASE\n" +
                    "            WHEN rpt.cust_accept_level IN ('KH đồng ý', 'KH Hài lòng', 'KH chấp nhận', 'KH không hài lòng - NVKT Đóng ảo') THEN 1\n" +
                    "            ELSE 0\n" +
                    "        END) AS mau_so_phan_hoi,\n" +
                    "    ROUND(\n" +
                    "            SUM(CASE\n" +
                    "                    WHEN rpt.cust_accept_level IN ('KH đồng ý', 'KH Hài lòng', 'KH chấp nhận') THEN 1\n" +
                    "                    ELSE 0\n" +
                    "                END) * 100.0\n" +
                    "                / NULLIF(SUM(CASE\n" +
                    "                                 WHEN rpt.cust_accept_level IN ('KH đồng ý', 'KH Hài lòng', 'KH chấp nhận', 'KH không hài lòng - NVKT Đóng ảo') THEN 1\n" +
                    "                                 ELSE 0\n" +
                    "                END), 0), 2\n" +
                    "    ) AS ty_le_hai_long\n" +
                    "FROM d_vcoc_problem_group g\n" +
                    "         JOIN d_vcoc_problem_group gc ON g.prob_group_id = gc.parent_id\n" +
                    "         JOIN rp_target_mapping rtm ON rtm.parent_prob_group_id = g.prob_group_id\n" +
                    "    AND rtm.prob_group_id = gc.prob_group_id\n" +
                    "         LEFT JOIN d_vcoc_problem_final_end_date dt ON dt.prob_type_id = rtm.prob_type_id\n" +
                    "    AND DATE(dt.create_date) BETWEEN DATE_FORMAT(:currentDate, '%Y-%m-01') AND :currentDate\n" +
                    "                                       AND dt.status <> '4'"+
                    "         LEFT JOIN d_vcoc_rpt_problem_gpdn_process rpt ON dt.problem_id = rpt.problem_id\n" +
                    "         JOIN rp_target rt ON rt.id = rtm.target_id\n" +
                    "         JOIN rp_target_of_table rtot ON rt.id = rtot.target_id\n" +
                    "         JOIN rp_table_target_config rttc ON rtot.table_target_config_id = rttc.id\n" +
                    "         JOIN rp_report_config c ON c.complaint_group_id = g.prob_group_id\n" +
                    "         JOIN rp_report_inf i ON i.id = c.report_id\n" +
                    "WHERE i.report_name = 'Báo cáo kênh truyền ngày'\n" +
                    "  AND rttc.report_table_name = 'KPI VTS'\n" +
                    "GROUP BY DATE(dt.create_date)\n" +
                    "ORDER BY ngay;\n", nativeQuery = true)
    List<Map<String, Object>> getLevelSatisfy(@Param("currentDate") LocalDate currentDate);

    @Query(value =
            "WITH parentAgg AS (\n" +
                    "    SELECT\n" +
                    "        parentName,\n" +
                    "        day,\n" +
                    "        totalSubcriberParent,\n" +
                    "        SUM(totalPerDay) AS totalPerDayParent\n" +
                    "    FROM (\n" +
                    "             SELECT\n" +
                    "                 DATE(dt.create_date) AS day,\n" +
                    "                 p.total_subcriber AS totalSubcriberParent,\n" +
                    "                 p.target_name AS parentName,\n" +
                    "                 COUNT(DISTINCT dt.problem_id) AS totalPerDay\n" +
                    "             FROM d_vcoc_problem_group rpg\n" +
                    "                      JOIN d_vcoc_problem_group rpgc\n" +
                    "                           ON rpg.prob_group_id = rpgc.parent_id\n" +
                    "                      JOIN rp_target_mapping rtm\n" +
                    "                           ON rtm.parent_prob_group_id = rpg.prob_group_id\n" +
                    "                               AND rtm.prob_group_id = rpgc.prob_group_id\n" +
                    "                      LEFT JOIN d_vcoc_problem_final_end_date dt\n" +
                    "                                ON rtm.prob_type_id = dt.prob_type_id\n" +
                    "                                    AND DATE(dt.create_date) BETWEEN DATE_SUB(:currentDate, INTERVAL DAY(:currentDate) - 1 DAY)\n" +
                    "                                       AND :currentDate\n" +
                    "                                       AND dt.status <> '4'"+
                    "                      LEFT JOIN d_vcoc_rpt_problem_gpdn_process gp\n" +
                    "                                ON dt.problem_id = gp.problem_id\n" +
                    "                                    AND (\n" +
                    "                                       gp.cust_accept_level IS NULL\n" +
                    "                                           OR gp.cust_accept_level NOT IN (\n" +
                    "                                                                           'Không Happy call KH – Đóng hủy KN theo yêu cầu CC',\n" +
                    "                                                                           'Không Happy call KH – Đóng trùng.'\n" +
                    "                                           )\n" +
                    "                                       )\n" +
                    "                      JOIN rp_target rt ON rt.id = rtm.target_id\n" +
                    "                      JOIN rp_target p ON rt.parent_id = p.id\n" +
                    "                      JOIN rp_target_of_table rtot ON rt.id = rtot.target_id\n" +
                    "                      JOIN rp_table_target_config rttc ON rtot.table_target_config_id = rttc.id\n" +
                    "                      JOIN rp_report_config c ON c.complaint_group_id = rpg.prob_group_id\n" +
                    "                      JOIN rp_report_inf i ON c.report_id = i.id\n" +
                    "             WHERE i.report_name = 'Báo cáo kênh truyền ngày'\n" +
                    "               AND rttc.report_table_name = 'Biểu đồ xu thế PAKH dịch vụ kênh truyền'\n" +
                    "             GROUP BY DATE(dt.create_date), rt.target_name, rt.order_no, p.total_subcriber, p.target_name\n" +
                    "         ) base\n" +
                    "    GROUP BY parentName, day, totalSubcriberParent\n" +
                    ")\n" +
                    "\n" +
                    "SELECT\n" +
                    "    parentName,\n" +
                    "    day,\n" +
                    "    totalSubcriberParent,\n" +
                    "    totalPerDayParent,\n" +
                    "    (totalPerDayParent * 1000 / totalSubcriberParent) AS tlpaParent,\n" +
                    "    SUM(totalPerDayParent) OVER (\n" +
                    "        PARTITION BY parentName\n" +
                    "        ORDER BY day\n" +
                    "        ) AS luyKeTotalPerDayParent,\n" +
                    "    (SUM(totalPerDayParent) OVER (\n" +
                    "        PARTITION BY parentName\n" +
                    "        ORDER BY day\n" +
                    "        ) * 1000 / totalSubcriberParent) AS tlpaParentLuyKe\n" +
                    "\n" +
                    "FROM parentAgg\n" +
                    "ORDER BY day, parentName;", nativeQuery = true)
    List<Map<String, Object>> getLast8daysAndAvgMonth(@Param("currentDate") LocalDate currentDate);

    @Query(value =
            "SELECT\n" +
            "    'Sự cố' AS type,\n" +
            "    ROUND(AVG(CASE WHEN report_year = :year - 1 THEN avg_incident END), 2) AS avgLastYear,\n" +
            "    ROUND(MAX(CASE WHEN report_year = :year AND report_month = 1  THEN avg_incident END), 2) AS T1,\n" +
            "    ROUND(MAX(CASE WHEN report_year = :year AND report_month = 2  THEN avg_incident END), 2) AS T2,\n" +
            "    ROUND(MAX(CASE WHEN report_year = :year AND report_month = 3  THEN avg_incident END), 2) AS T3,\n" +
            "    ROUND(MAX(CASE WHEN report_year = :year AND report_month = 4  THEN avg_incident END), 2) AS T4,\n" +
            "    ROUND(MAX(CASE WHEN report_year = :year AND report_month = 5  THEN avg_incident END), 2) AS T5,\n" +
            "    ROUND(MAX(CASE WHEN report_year = :year AND report_month = 6  THEN avg_incident END), 2) AS T6,\n" +
            "    ROUND(MAX(CASE WHEN report_year = :year AND report_month = 7  THEN avg_incident END), 2) AS T7,\n" +
            "    ROUND(MAX(CASE WHEN report_year = :year AND report_month = 8  THEN avg_incident END), 2) AS T8,\n" +
            "    ROUND(MAX(CASE WHEN report_year = :year AND report_month = 9  THEN avg_incident END), 2) AS T9,\n" +
            "    ROUND(MAX(CASE WHEN report_year = :year AND report_month = 10 THEN avg_incident END), 2) AS T10,\n" +
            "    ROUND(MAX(CASE WHEN report_year = :year AND report_month = 11 THEN avg_incident END), 2) AS T11,\n" +
            "    ROUND(MAX(CASE WHEN report_year = :year AND report_month = 12 THEN avg_incident END), 2) AS T12\n" +
            "FROM rp_monthly_complaint_summary\n" +
            "\n" +
            "UNION ALL\n" +
            "\n" +
            "SELECT\n" +
            "    'TLPA/1000TB' AS type,\n" +
            "    ROUND(AVG(CASE WHEN report_year = :year - 1 THEN ratio_per_1000 END), 2) AS avgLastYear,\n" +
            "    ROUND(MAX(CASE WHEN report_year = :year AND report_month = 1  THEN ratio_per_1000 END), 2) AS T1,\n" +
            "    ROUND(MAX(CASE WHEN report_year = :year AND report_month = 2  THEN ratio_per_1000 END), 2) AS T2,\n" +
            "    ROUND(MAX(CASE WHEN report_year = :year AND report_month = 3  THEN ratio_per_1000 END), 2) AS T3,\n" +
            "    ROUND(MAX(CASE WHEN report_year = :year AND report_month = 4  THEN ratio_per_1000 END), 2) AS T4,\n" +
            "    ROUND(MAX(CASE WHEN report_year = :year AND report_month = 5  THEN ratio_per_1000 END), 2) AS T5,\n" +
            "    ROUND(MAX(CASE WHEN report_year = :year AND report_month = 6  THEN ratio_per_1000 END), 2) AS T6,\n" +
            "    ROUND(MAX(CASE WHEN report_year = :year AND report_month = 7  THEN ratio_per_1000 END), 2) AS T7,\n" +
            "    ROUND(MAX(CASE WHEN report_year = :year AND report_month = 8  THEN ratio_per_1000 END), 2) AS T8,\n" +
            "    ROUND(MAX(CASE WHEN report_year = :year AND report_month = 9  THEN ratio_per_1000 END), 2) AS T9,\n" +
            "    ROUND(MAX(CASE WHEN report_year = :year AND report_month = 10 THEN ratio_per_1000 END), 2) AS T10,\n" +
            "    ROUND(MAX(CASE WHEN report_year = :year AND report_month = 11 THEN ratio_per_1000 END), 2) AS T11,\n" +
            "    ROUND(MAX(CASE WHEN report_year = :year AND report_month = 12 THEN ratio_per_1000 END), 2) AS T12\n" +
            "FROM rp_monthly_complaint_summary;\n", nativeQuery = true)
    List<Map<String, Object>> getIncidentMonthlySummary(@Param("year") int year);   
}

