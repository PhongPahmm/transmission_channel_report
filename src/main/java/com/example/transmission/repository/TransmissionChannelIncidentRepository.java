package com.example.transmission.repository;

import com.example.transmission.domain.TransmissionChannelIncident;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.stereotype.Repository;

import java.util.List;
import java.util.Map;

@Repository
public interface TransmissionChannelIncidentRepository extends JpaRepository<TransmissionChannelIncident, Long> {
    @Query(value =
            "SELECT " +
                    "  CASE " +
                    "    WHEN LOWER(category) LIKE '%sự cố office wan%' " +
                    "      OR LOWER(category) LIKE '%sự cố metrowan%' " +
                    "    THEN 'Sự cố Office WAN' " +
                    "    ELSE category " +
                    "  END AS category_group, " +
                    "  COUNT(*) AS total_records " +
                    "FROM rp_transmission_channel_incidents " +
                    "WHERE received_date = CURDATE() - INTERVAL 7 DAY " +
                    "  AND complaint_group = '05. Báo hỏng DV Kênh truyền' " +
                    "  AND satisfaction_level NOT IN ( " +
                    "        'Không Happy call KH – Đóng hủy KN theo yêu cầu CC', " +
                    "        'Không Happy call KH – Đóng trùng.' " +
                    "      ) " +
                    "  AND LOWER(category) LIKE '%sự cố%' " +
                    "GROUP BY category_group",
            nativeQuery = true)
    List<Map<String, Object>> getCategorySummary();

    @Query(value = "SELECT category_group, luy_ke AS total_records, so_ngay, pa_ngay, ngay_check " +
            "FROM (SELECT " +
            "CASE " +
            "WHEN TRIM(LOWER(category)) IN ('sự cố office wan', 'sự cố metrowan') THEN 'Office WAN' " +
            "WHEN TRIM(LOWER(category)) IN ('sự cố hàng loạt leased line', 'sự cố leased line') THEN 'Leased Line' " +
            "ELSE TRIM(category) END AS category_group, " +
            "COUNT(*) AS luy_ke, " +
            "DAY(DATE_SUB(CURDATE(), INTERVAL 5 DAY)) AS so_ngay, " +
            "CASE WHEN DAY(DATE_SUB(CURDATE(), INTERVAL 5 DAY)) <= 0 THEN 0 " +
            "ELSE ROUND(COUNT(*) / DAY(DATE_SUB(CURDATE(), INTERVAL 5 DAY)), 0) END AS pa_ngay, " +
            "DATE_SUB(CURDATE(), INTERVAL 5 DAY) AS ngay_check " +
            "FROM rp_transmission_channel_incidents " +
            "WHERE complaint_group = '05. Báo hỏng DV Kênh truyền' " +
            "AND satisfaction_level NOT LIKE 'Không Happy call KH – Đóng hủy KN theo yêu cầu CC' " +
            "AND satisfaction_level NOT LIKE 'Không Happy call KH – Đóng trùng.' " +
            "AND received_date >= DATE_FORMAT(DATE_SUB(CURDATE(), INTERVAL 5 DAY), '%Y-%m-01') " +
            "AND received_date <= DATE_SUB(CURDATE(), INTERVAL 5 DAY) " +
            "AND LOWER(category) LIKE 'sự cố%' " +
            "GROUP BY " +
            "CASE " +
            "WHEN TRIM(LOWER(category)) IN ('sự cố office wan', 'sự cố metrowan') THEN 'Office WAN' " +
            "WHEN TRIM(LOWER(category)) IN ('sự cố hàng loạt leased line', 'sự cố leased line') THEN 'Leased Line' " +
            "ELSE TRIM(category) END) AS cg", nativeQuery = true)
    List<Map<String, Object>> getComplaintDay();

    @Query(value = "SELECT category, " +
            "SUM(CASE WHEN COALESCE(gnoc_processing_hours, total_processing_hours) <= " +
            "    CASE WHEN category = 'Sự cố Kênh truyền quốc tế' THEN 5 ELSE 3 END " +
            "    THEN 1 ELSE 0 END) AS xu_li_trong_han_3h, " +
            "SUM(CASE WHEN COALESCE(gnoc_processing_hours, total_processing_hours) <= 24 " +
            "    THEN 1 ELSE 0 END) AS xu_li_trong_han_24h, " +
            "SUM(CASE WHEN COALESCE(gnoc_processing_hours, total_processing_hours) <= 48 " +
            "    THEN 1 ELSE 0 END) AS xu_li_trong_han_48h " +
            "FROM rp_transmission_channel_incidents " +
            "WHERE received_date = CURDATE() - INTERVAL 7 DAY " +
            "  AND complaint_group = '05. Báo hỏng DV Kênh truyền' " +
            "  AND LOWER(category) LIKE '%sự cố%' " +
            "  AND COALESCE(satisfaction_level, '') NOT IN ( " +
            "        'Không Happy call KH – Đóng hủy KN theo yêu cầu CC', " +
            "        'Không Happy call KH – Đóng trùng.' " +
            "      ) " +
            "GROUP BY category " +
            "UNION ALL " +
            "SELECT 'TỔNG' AS category, " +
            "SUM(CASE WHEN COALESCE(gnoc_processing_hours, total_processing_hours) <= " +
            "    CASE WHEN category = 'Sự cố Kênh truyền quốc tế' THEN 5 ELSE 3 END " +
            "    THEN 1 ELSE 0 END) AS xu_li_trong_han_3h, " +
            "SUM(CASE WHEN COALESCE(gnoc_processing_hours, total_processing_hours) <= 24 " +
            "    THEN 1 ELSE 0 END) AS xu_li_trong_han_24h, " +
            "SUM(CASE WHEN COALESCE(gnoc_processing_hours, total_processing_hours) <= 48 " +
            "    THEN 1 ELSE 0 END) AS xu_li_trong_han_48h " +
            "FROM rp_transmission_channel_incidents " +
            "WHERE received_date = CURDATE() - INTERVAL 7 DAY " +
            "  AND complaint_group = '05. Báo hỏng DV Kênh truyền' " +
            "  AND LOWER(category) LIKE '%sự cố%' " +
            "  AND COALESCE(satisfaction_level, '') NOT IN ( " +
            "        'Không Happy call KH – Đóng hủy KN theo yêu cầu CC', " +
            "        'Không Happy call KH – Đóng trùng.' " +
            "      )", nativeQuery = true)
    List<Map<String, Object>> getOnTimeHandle();

    @Query(value =
            "SELECT category, COUNT(*) AS tong_xu_ly " +
                    "FROM rp_transmission_channel_incidents " +
                    "WHERE received_date = (CURDATE() - INTERVAL 7 DAY) " +
                    "  AND complaint_group = '05. Báo hỏng DV Kênh truyền' " +
                    "  AND LOWER(category) LIKE '%sự cố%' " +
                    "  AND COALESCE(satisfaction_level, '') NOT IN ( " +
                    "        'Không Happy call KH – Đóng hủy KN theo yêu cầu CC', " +
                    "        'Không Happy call KH – Đóng trùng.' " +
                    "      ) " +
                    "GROUP BY category " +
                    "UNION ALL " +
                    "SELECT 'TỔNG PHẢN ÁNH ĐÃ XỬ LÝ' AS category, COUNT(*) AS tong_xu_ly " +
                    "FROM rp_transmission_channel_incidents " +
                    "WHERE received_date = (CURDATE() - INTERVAL 7 DAY) " +
                    "  AND complaint_group = '05. Báo hỏng DV Kênh truyền' " +
                    "  AND LOWER(category) LIKE '%sự cố%' " +
                    "  AND COALESCE(satisfaction_level, '') NOT IN ( " +
                    "        'Không Happy call KH – Đóng hủy KN theo yêu cầu CC', " +
                    "        'Không Happy call KH – Đóng trùng.' " +
                    "      )",
            nativeQuery = true)
    List<Map<String, Object>> getTotalHandle();
    @Query(value =
            "WITH handle_data AS ( " +
                    "    SELECT category, " +
                    "           COUNT(*) AS tong_xu_ly, " +
                    "           SUM(CASE WHEN COALESCE(gnoc_processing_hours, total_processing_hours) <= " +
                    "                        CASE WHEN category = 'Sự cố Kênh truyền quốc tế' THEN 5 ELSE 3 END " +
                    "                    THEN 1 ELSE 0 END) AS xu_li_trong_han_3h, " +
                    "           SUM(CASE WHEN COALESCE(gnoc_processing_hours, total_processing_hours) <= 24 " +
                    "                    THEN 1 ELSE 0 END) AS xu_li_trong_han_24h, " +
                    "           SUM(CASE WHEN COALESCE(gnoc_processing_hours, total_processing_hours) <= 48 " +
                    "                    THEN 1 ELSE 0 END) AS xu_li_trong_han_48h " +
                    "    FROM rp_transmission_channel_incidents " +
                    "    WHERE received_date BETWEEN '2025-08-01' AND '2025-08-28' " +
                    "      AND complaint_group = '05. Báo hỏng DV Kênh truyền' " +
                    "      AND LOWER(category) LIKE '%sự cố%' " +
                    "      AND COALESCE(satisfaction_level, '') NOT IN ( " +
                    "            'Không Happy call KH – Đóng hủy KN theo yêu cầu CC', " +
                    "            'Không Happy call KH – Đóng trùng.' " +
                    "      ) " +
                    "    GROUP BY category " +
                    ") " +
                    "SELECT category, " +
                    "       tong_xu_ly, " +
                    "       xu_li_trong_han_3h, " +
                    "       xu_li_trong_han_24h, " +
                    "       xu_li_trong_han_48h " +
                    "FROM handle_data " +
                    "UNION ALL " +
                    "SELECT 'TỔNG PHẢN ÁNH ĐÃ XỬ LÝ', " +
                    "       SUM(tong_xu_ly), " +
                    "       SUM(xu_li_trong_han_3h), " +
                    "       SUM(xu_li_trong_han_24h), " +
                    "       SUM(xu_li_trong_han_48h) " +
                    "FROM handle_data",
            nativeQuery = true)
    List<Map<String, Object>> getResult();

    @Query(value = "SELECT DATE(received_date) AS ngay, " +
            "       CASE " +
            "           WHEN LOWER(category) LIKE '%hỗ trợ%' THEN 'HT' " +
            "           WHEN LOWER(category) LIKE '%sự cố%' " +
            "                AND (LOWER(category) LIKE '%office wan%' OR LOWER(category) LIKE '%metrowan%') " +
            "                THEN 'Office WAN' " +
            "           WHEN LOWER(category) LIKE '%leased line%' THEN 'Leased Line' " +
            "           WHEN LOWER(category) LIKE '%kênh trắng%' THEN 'Kênh trắng' " +
            "           WHEN LOWER(category) LIKE '%kênh truyền quốc tế%' THEN 'Kênh truyền quốc tế' " +
            "           ELSE category " +
            "       END AS cate_hien_thi, " +
            "       COUNT(*) AS so_luong " +
            "FROM rp_transmission_channel_incidents " +
            "WHERE received_date BETWEEN DATE_FORMAT(CURDATE() - INTERVAL 1 MONTH, '%Y-%m-01') " +
            "                        AND LAST_DAY(CURDATE() - INTERVAL 1 MONTH) " +
            "  AND complaint_group = '05. Báo hỏng DV Kênh truyền' " +
            "  AND COALESCE(satisfaction_level, '') NOT IN ( " +
            "        'Không Happy call KH – Đóng hủy KN theo yêu cầu CC', " +
            "        'Không Happy call KH – Đóng trùng.' " +
            "      ) " +
            "  AND (LOWER(category) LIKE '%sự cố%' OR LOWER(category) LIKE '%hỗ trợ%') " +
            "GROUP BY DATE(received_date), cate_hien_thi " +
            "UNION ALL " +
            "SELECT DATE(received_date) AS ngay, " +
            "       'Sự cố' AS cate_hien_thi, " +
            "       COUNT(*) AS so_luong " +
            "FROM rp_transmission_channel_incidents " +
            "WHERE received_date BETWEEN DATE_FORMAT(CURDATE() - INTERVAL 1 MONTH, '%Y-%m-01') " +
            "                        AND LAST_DAY(CURDATE() - INTERVAL 1 MONTH) " +
            "  AND complaint_group = '05. Báo hỏng DV Kênh truyền' " +
            "  AND COALESCE(satisfaction_level, '') NOT IN ( " +
            "        'Không Happy call KH – Đóng hủy KN theo yêu cầu CC', " +
            "        'Không Happy call KH – Đóng trùng.' " +
            "      ) " +
            "  AND LOWER(category) LIKE '%sự cố%' " +
            "GROUP BY DATE(received_date) " +
            "ORDER BY ngay, " +
            "         CASE WHEN cate_hien_thi = 'Sự cố' THEN 0 ELSE 1 END, " +
            "         so_luong DESC",
            nativeQuery = true)
    List<Map<String, Object>> getDailyStats();


}
