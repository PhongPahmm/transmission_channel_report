package com.example.transmission.service;

import com.example.transmission.domain.ServiceSubscriber;
import com.example.transmission.dto.ExcelDTO;
import com.example.transmission.repository.ProvinceSubscriberRepository;
import com.example.transmission.repository.ServiceSubscriberRepository;
import com.example.transmission.repository.TransmissionChannelIncidentRepository;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.*;
import java.math.BigDecimal;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.stream.Collectors;

@Slf4j
@Service
@RequiredArgsConstructor
public class TransmissionChannelIncidentService {

    private final TransmissionChannelIncidentRepository transmissionChannelIncidentRepository;
    private final ServiceSubscriberRepository serviceSubscriberRepository;
    private final ProvinceSubscriberRepository provinceSubscriberRepository;

    // Constants for Excel structure
    private static final String TEMPLATE_PATH = "/templates/template_output.xlsx";
    private static final String SHEET_NAME = "PL06-Kenh truyen";
    private static final String TOTAL = "TỔNG";
    private static final String TOTAL_PROCESSED_CATEGORY = "TỔNG PHẢN ÁNH ĐÃ XỬ LÝ";

    // Table 1 constants
    private static final int TABLE1_START_ROW = 7;
    private static final int TABLE1_END_ROW = 10;
    private static final int CATEGORY_COL = 1;
    private static final int KPI_COL = 2;
    private static final int SUBSCRIBERS_COL = 3;
    private static final int VALUE_COL = 4;
    private static final int COMP_DAY_COL = 8;
    private static final int PARENT_KPI_ROW = 6;

    // Table 2 constants
    private static final int TABLE2_START_ROW = 6;
    private static final int PROGRESS_KPI_COL = 17; // Column R
    private static final int ON_TIME_HANDLE_COL = 19; // Column T
    private static final int TOTAL_HANDLE_COL = 20; // Column U
    private static final int RESULT_COL = 22; // Column W

    // Table 3 constants (Daily Stats)
    private static final int TABLE3_START_ROW = 32; // B34 = row 33 (0-based)
    private static final int TABLE3_END_ROW = 37;   // B38 = row 37 (0-based)
    private static final int TABLE3_DATA_START_COL = 4; // Column E = index 4

    // Table 4 constants (Complaint Rate by Province)
    private static final int TABLE4_START_ROW = 44; // B45 = row 44 (0-based)
    private static final int TABLE4_END_ROW = 77;   // B78 = row 77 (0-based)
    private static final int TABLE4_DATA_START_COL = 6; // Column G = index 6
    private static final int TABLE4_TOTAL_ROW = 43;

    // Table 5 constants (Handle Rate by Province)
    private static final int TABLE5_START_ROW = 84; // B85 = row 84 (0-based)
    private static final int TABLE5_END_ROW = 117;  // B118 = row 117 (0-based)
    private static final int TABLE5_DATA_START_COL = 7; // Column H = index 7 (da_xu_li_3h)
    private static final int TABLE5_TOTAL_ROW = 83;     // Row 84 = row 83 (0-based)

    // Table 6 constants (Handle Rate by Province)
    private static final int TABLE6_START_ROW = 124;
    private static final int TABLE6_END_ROW = 157;
    private static final int TABLE6_DATA_START_COL = 7;
    private static final int TABLE6_TOTAL_ROW = 123;

    // Table 7 constants (Handle Rate by Province)
    private static final int TABLE7_START_ROW = 164;
    private static final int TABLE7_END_ROW = 197;
    private static final int TABLE7_DATA_START_COL = 7;
    private static final int TABLE7_TOTAL_ROW = 163;

    // Table 8 constants (Handle Rate by Province)
    private static final int TABLE8_START_ROW = 204;
    private static final int TABLE8_END_ROW = 237;
    private static final int TABLE8_DATA_START_COL = 7;
    private static final int TABLE8_TOTAL_ROW = 203;

    public byte[] exportExcelFile() {
        try (InputStream inputStream = getClass().getResourceAsStream(TEMPLATE_PATH);
             XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
             ByteArrayOutputStream out = new ByteArrayOutputStream()) {

            Sheet sheet = workbook.getSheet(SHEET_NAME);

            // Get all required data
            ExcelDTO excelData = fetchExcelData();

            // Fill tables
            fillTable1(sheet, excelData);
            fillTable2(sheet, excelData);
            fillTable3(sheet, excelData);
            fillTable4(sheet, excelData);
            fillTable5(sheet, excelData);
            fillTable6(sheet, excelData);
            fillTable7(sheet, excelData);
            fillTable8(sheet, excelData);

            workbook.setForceFormulaRecalculation(true);
            workbook.write(out);
            return out.toByteArray();

        } catch (IOException e) {
            log.error("Error while exporting Excel", e);
            throw new RuntimeException("Error while exporting Excel", e);
        }
    }

    private ExcelDTO fetchExcelData() {
        return ExcelDTO.builder()
                .categorySummaryResults(transmissionChannelIncidentRepository.getCategorySummary())
                .complaintDayResults(transmissionChannelIncidentRepository.getComplaintDay())
                .serviceSubscribers(serviceSubscriberRepository.findAll())
                .onTimeHandle(transmissionChannelIncidentRepository.getOnTimeHandle())
                .progressKpi(serviceSubscriberRepository.getProgressKpi())
                .totalHandle(transmissionChannelIncidentRepository.getTotalHandle())
                .results(transmissionChannelIncidentRepository.getResult())
                .dailyStats(transmissionChannelIncidentRepository.getDailyStats())
                .complaintRate(transmissionChannelIncidentRepository.getComplaintRateAndTotalSubscribers())
                .provinceSubscribers(provinceSubscriberRepository.getProvinceSubscribers())
                .handleRate3h(transmissionChannelIncidentRepository.getHandleRate3h())
                .handleRate24h(transmissionChannelIncidentRepository.getHandleRate24h())
                .handleRate48h(transmissionChannelIncidentRepository.getHandleRate48h())
                .handleRate3hVip(transmissionChannelIncidentRepository.getHandleRate3hVip())
                .build();
    }

    private void fillTable1(Sheet sheet, ExcelDTO data) {
        fillParentKpi(sheet, data.getServiceSubscribers());
        fillServiceSubscriberData(sheet, data.getServiceSubscribers());
        fillCategoryData(sheet, data.getCategorySummaryResults(), data.getComplaintDayResults());
    }

    private void fillParentKpi(Sheet sheet, List<ServiceSubscriber> serviceSubscribers) {
        if (serviceSubscribers.isEmpty()) return;

        ServiceSubscriber parent = serviceSubscribers.get(0);
        BigDecimal parentKpi = parent.getComplaintGroupKpi();

        if (parentKpi != null) {
            setCellValue(sheet, PARENT_KPI_ROW, KPI_COL, parentKpi.doubleValue());
        }
    }

    private void fillServiceSubscriberData(Sheet sheet, List<ServiceSubscriber> serviceSubscribers) {
        for (int i = 0; i < serviceSubscribers.size(); i++) {
            ServiceSubscriber ss = serviceSubscribers.get(i);
            log.debug("Processing service subscriber: {}", ss);

            int rowIndex = TABLE1_START_ROW + i;
            setCellValue(sheet, rowIndex, SUBSCRIBERS_COL, ss.getTotalSubscribers());
            setCellValue(sheet, rowIndex, KPI_COL, ss.getServiceKpi().doubleValue());
        }
    }

    private void fillCategoryData(Sheet sheet, List<Map<String, Object>> categorySummary,
                                  List<Map<String, Object>> complaintDayResults) {
        for (int rowIndex = TABLE1_START_ROW; rowIndex <= TABLE1_END_ROW; rowIndex++) {
            String templateCategory = getCellStringValue(sheet, rowIndex, CATEGORY_COL);
            if (templateCategory == null) continue;

            final int currentRowIndex = rowIndex; // Make effectively final for lambda

            // Fill category summary data
            findMatchingCategory(categorySummary, templateCategory, "category_group")
                    .ifPresent(data -> {
                        setCellValue(sheet, currentRowIndex, VALUE_COL, getNumberValue(data, "total_records"));
                    });

            // Fill complaint day data
            findMatchingCategory(complaintDayResults, templateCategory, "category_group")
                    .ifPresent(data -> {
                        setCellValue(sheet, currentRowIndex, COMP_DAY_COL, getNumberValue(data, "pa_ngay"));
                    });
        }
    }

    private void fillTable2(Sheet sheet, ExcelDTO data) {
        fillProgressKpiData(sheet, data.getProgressKpi());
        fillOnTimeHandleData(sheet, data.getOnTimeHandle());
        fillTotalHandleData(sheet, data.getTotalHandle());
        fillResultData(sheet, data.getResults());
    }

    private void fillTable3(Sheet sheet, ExcelDTO data) {
        fillDailyStats(sheet, data.getDailyStats());
    }

    private void fillTable4(Sheet sheet, ExcelDTO data) {
        fillComplaintRateAndSubscribersByProvince(sheet, data.getComplaintRate(), data.getProvinceSubscribers());
    }

    private void fillTable5(Sheet sheet, ExcelDTO data) {
        fillHandleRate3hByProvince(sheet, data.getHandleRate3h());
    }
    private void fillTable6(Sheet sheet, ExcelDTO data) {
        fillHandleRate24hByProvince(sheet, data.getHandleRate24h());
    }

    private void fillTable7(Sheet sheet, ExcelDTO data) {
        fillHandleRate48hByProvince(sheet, data.getHandleRate48h());
    }

    private void fillTable8(Sheet sheet, ExcelDTO data) {
        fillHandleRate3hVipByProvince(sheet, data.getHandleRate3hVip());
    }

    private void fillProgressKpiData(Sheet sheet, List<Map<String, Object>> progressKpi) {
        if (progressKpi == null || progressKpi.isEmpty()) return;

        Map<String, Object> data = progressKpi.get(0);
        fillTimeBasedData(sheet, data, PROGRESS_KPI_COL,
                "progress_3h", "progress_24h", "progress_48h", true);
    }

    private void fillOnTimeHandleData(Sheet sheet, List<Map<String, Object>> onTimeHandle) {
        findCategoryData(onTimeHandle, TOTAL, "category")
                .ifPresent(data -> {
                    fillTimeBasedData(sheet, data, ON_TIME_HANDLE_COL,
                            "xu_li_trong_han_3h", "xu_li_trong_han_24h", "xu_li_trong_han_48h", false);
                });
    }

    private void fillTotalHandleData(Sheet sheet, List<Map<String, Object>> totalHandle) {
        findCategoryData(totalHandle, TOTAL_PROCESSED_CATEGORY, "category")
                .ifPresent(data -> {
                    Double totalValue = getNumberValue(data, "tong_xu_ly");
                    if (totalValue != null) {
                        for (int i = 0; i < 3; i++) {
                            setCellValue(sheet, TABLE2_START_ROW + i, TOTAL_HANDLE_COL, totalValue);
                        }
                    }
                });
    }

    private void fillResultData(Sheet sheet, List<Map<String, Object>> results) {
        findCategoryData(results, TOTAL_PROCESSED_CATEGORY, "category")
                .ifPresent(data -> {
                    Double total = getNumberValue(data, "tong_xu_ly");
                    if (total != null && total > 0) {
                        Double val3h = getNumberValue(data, "xu_li_trong_han_3h");
                        Double val24h = getNumberValue(data, "xu_li_trong_han_24h");
                        Double val48h = getNumberValue(data, "xu_li_trong_han_48h");

                        setCellValue(sheet, TABLE2_START_ROW, RESULT_COL,
                                val3h != null ? val3h / total : null);
                        setCellValue(sheet, TABLE2_START_ROW + 1, RESULT_COL,
                                val24h != null ? val24h / total : null);
                        setCellValue(sheet, TABLE2_START_ROW + 2, RESULT_COL,
                                val48h != null ? val48h / total : null);
                    }
                });
    }

    private void fillTimeBasedData(Sheet sheet, Map<String, Object> data, int column,
                                   String key3h, String key24h, String key48h, boolean divideBy100) {
        Double val3h = getNumberValue(data, key3h);
        Double val24h = getNumberValue(data, key24h);
        Double val48h = getNumberValue(data, key48h);

        if (divideBy100) {
            val3h = val3h != null ? val3h / 100 : null;
            val24h = val24h != null ? val24h / 100 : null;
            val48h = val48h != null ? val48h / 100 : null;
        }

        setCellValue(sheet, TABLE2_START_ROW, column, val3h);
        setCellValue(sheet, TABLE2_START_ROW + 1, column, val24h);
        setCellValue(sheet, TABLE2_START_ROW + 2, column, val48h);
    }

    private void fillDailyStats(Sheet sheet, List<Map<String, Object>> dailyData) {
        if (dailyData == null || dailyData.isEmpty()) return;

        // Process each service row from B34 to B38 (rows 33-37 in 0-based indexing)
        for (int serviceRow = TABLE3_START_ROW; serviceRow <= TABLE3_END_ROW; serviceRow++) {
            String serviceName = getCellStringValue(sheet, serviceRow, CATEGORY_COL); // Column B
            if (serviceName == null) continue;

            final int currentServiceRow = serviceRow; // Make effectively final for lambda

            // Filter all data by cate_hien_thi matching service name
            List<Map<String, Object>> serviceData = dailyData.stream()
                    .filter(d -> {
                        Object cate = d.get("cate_hien_thi");
                        return cate != null &&
                                cate.toString().trim()
                                        .contains(serviceName);
                    })
                    .collect(Collectors.toList());

            for (int day = 1; day <= 31; day++) {
                int colIndex = TABLE3_DATA_START_COL + (day - 1); // E=4, F=5, etc.
                setCellValue(sheet, currentServiceRow, colIndex, 0.0);
            }
            // Fill data for each day
            serviceData.forEach(data -> {
                Object ngayObj = data.get("ngay");
                Object countObj = data.get("so_luong");

                if (ngayObj != null && countObj instanceof Number) {
                    // Determine column by day (assuming ngayObj is a Date)
                    int day = extractDayFromDate(ngayObj);
                    if (day > 0) {
                        int colIndex = TABLE3_DATA_START_COL + (day - 1); // E=4, F=5, etc.
                        setCellValue(sheet, currentServiceRow, colIndex, ((Number) countObj).doubleValue());
                    }
                }
            });
        }
    }

    private void fillComplaintRateAndSubscribersByProvince(Sheet sheet, List<Map<String, Object>> complaintData,
                                                           List<Map<String, Object>> provinceSubscribers) {
        if (provinceSubscribers == null || provinceSubscribers.isEmpty()) return;

        int currentDay = java.time.LocalDate.now().getDayOfMonth();

        // Create optimized maps once
        Map<String, List<Map<String, Object>>> dataByProvince = complaintData != null ?
                complaintData.stream().collect(Collectors.groupingBy(
                        data -> data.get("tinh") != null ? data.get("tinh").toString().trim() : "Unknown"
                )) : Map.of();

        Map<String, Double> provinceSubscriberMap = provinceSubscribers.stream()
                .collect(Collectors.toMap(
                        data -> data.get("province_name") != null ? data.get("province_name").toString().trim() : "Unknown",
                        data -> {
                            Object sltbObj = data.get("total_subscribers");
                            return (sltbObj instanceof Number) ? ((Number) sltbObj).doubleValue() : 0.0;
                        }
                ));

        // Fill sltb to column E for all province rows
        for (int provinceRow = TABLE4_START_ROW; provinceRow <= TABLE4_END_ROW; provinceRow++) {
            String provinceName = getCellStringValue(sheet, provinceRow, CATEGORY_COL);
            if (provinceName == null) continue;

            Double subscriberCount = provinceSubscriberMap.entrySet().stream()
                    .filter(entry -> {
                        String lowerKey = entry.getKey().toLowerCase();
                        String lowerName = provinceName.toLowerCase();
                        return lowerKey.contains(lowerName) || lowerName.contains(lowerKey);
                    })
                    .map(Map.Entry::getValue)
                    .findFirst()
                    .orElse(0.0);

            setCellValue(sheet, provinceRow, 4, subscriberCount); // Column E
        }

        // Fill total sltb to column E for TOTAL row
        Double totalSltb = provinceSubscriberMap.getOrDefault("TỔNG", 0.0);
        setCellValue(sheet, TABLE4_TOTAL_ROW, 4, totalSltb);

        // Process province rows with optimized logic
        for (int provinceRow = TABLE4_START_ROW; provinceRow <= TABLE4_END_ROW; provinceRow++) {
            String provinceName = getCellStringValue(sheet, provinceRow, CATEGORY_COL);
            if (provinceName == null) continue;

            // Find subscriber count (already calculated above, but need to recalculate for consistency)
            Double subscriberCount = provinceSubscriberMap.entrySet().stream()
                    .filter(entry -> {
                        String lowerKey = entry.getKey().toLowerCase();
                        String lowerName = provinceName.toLowerCase();
                        return lowerKey.contains(lowerName) || lowerName.contains(lowerKey);
                    })
                    .map(Map.Entry::getValue)
                    .findFirst()
                    .orElse(0.0);

            // Initialize default values for past days only
            for (int day = 1; day <= currentDay; day++) {
                int colIndex = TABLE4_DATA_START_COL + (day - 1) * 3;
                setCellValue(sheet, provinceRow, colIndex, 0.0); // slpa
                setCellValue(sheet, provinceRow, colIndex + 1, subscriberCount); // sltb
                setCellValue(sheet, provinceRow, colIndex + 2, 0.0); // tlpa
            }

            // Fill actual complaint data if available
            Optional<String> matchingKey = dataByProvince.keySet().stream()
                    .filter(key -> {
                        String lowerKey = key.toLowerCase();
                        String lowerName = provinceName.toLowerCase();
                        return lowerKey.contains(lowerName) || lowerName.contains(lowerKey);
                    })
                    .findFirst();

            if (matchingKey.isPresent()) {
                List<Map<String, Object>> provinceData = dataByProvince.get(matchingKey.get());
                int finalProvinceRow = provinceRow;
                provinceData.forEach(data -> {
                    Object receivedDateObj = data.get("received_date");
                    Object slpaObj = data.get("slpa");
                    Object tlpaObj = data.get("tlpa");

                    if (receivedDateObj != null && slpaObj instanceof Number) {
                        int day = extractDayFromDate(receivedDateObj);
                        if (day > 0 && day <= 31) {
                            int colIndex = TABLE4_DATA_START_COL + (day - 1) * 3;
                            setCellValue(sheet, finalProvinceRow, colIndex, ((Number) slpaObj).doubleValue());
                            if (tlpaObj instanceof Number) {
                                setCellValue(sheet, finalProvinceRow, colIndex + 2, ((Number) tlpaObj).doubleValue());
                            }
                        }
                    }
                });
                log.debug("Filled complaint rate data for province: {} with {} records", provinceName, provinceData.size());
            }
        }

        // Handle TOTAL row
        List<Map<String, Object>> totalData = dataByProvince.get("TỔNG");

        // Initialize TOTAL row with default values for past days
        for (int day = 1; day <= currentDay; day++) {
            int colIndex = TABLE4_DATA_START_COL + (day - 1) * 3;
            setCellValue(sheet, TABLE4_TOTAL_ROW, colIndex, 0.0); // slpa
            setCellValue(sheet, TABLE4_TOTAL_ROW, colIndex + 1, totalSltb); // sltb
            setCellValue(sheet, TABLE4_TOTAL_ROW, colIndex + 2, 0.0); // tlpa
        }

        // Fill actual TOTAL data if available
        if (totalData != null && !totalData.isEmpty()) {
            totalData.forEach(data -> {
                int day = extractDayFromDate(data.get("received_date"));
                if (day > 0 && day <= 31) {
                    int colIndex = TABLE4_DATA_START_COL + (day - 1) * 3;

                    Object slpa = data.get("slpa");
                    Object sltb = data.get("sltb");
                    Object tlpa = data.get("tlpa");

                    if (slpa instanceof Number) {
                        setCellValue(sheet, TABLE4_TOTAL_ROW, colIndex, ((Number) slpa).doubleValue());
                    }
                    Double sltbValue = (sltb instanceof Number) ? ((Number) sltb).doubleValue() : totalSltb;
                    setCellValue(sheet, TABLE4_TOTAL_ROW, colIndex + 1, sltbValue);
                    if (tlpa instanceof Number) {
                        setCellValue(sheet, TABLE4_TOTAL_ROW, colIndex + 2, ((Number) tlpa).doubleValue());
                    }
                }
            });
        }
    }
    private void fillHandleRate3hByProvince(Sheet sheet, List<Map<String, Object>> handleRateData) {
        if (handleRateData == null) handleRateData = List.of();

        int currentDay = java.time.LocalDate.now().getDayOfMonth();

        // Create optimized map once - group data by province
        Map<String, List<Map<String, Object>>> dataByProvince = handleRateData.stream()
                .collect(Collectors.groupingBy(
                        d -> d.get("tinh") != null ? d.get("tinh").toString().trim() : "unknown"
                ));

        // Process province rows with optimized logic
        for (int provinceRow = TABLE5_START_ROW; provinceRow <= TABLE5_END_ROW; provinceRow++) {
            String provinceName = getCellStringValue(sheet, provinceRow, CATEGORY_COL);
            if (provinceName == null) continue;

            // Initialize default values for past days, null for future days
            for (int day = 1; day <= 31; day++) {
                int colIndex = TABLE5_DATA_START_COL + (day - 1) * 3;
                if (day <= currentDay) {
                    setCellValue(sheet, provinceRow, colIndex, 0.0);     // da_xu_ly_3h
                    setCellValue(sheet, provinceRow, colIndex + 1, 0.0); // tong_sc_da_xu_ly
                    setCellValue(sheet, provinceRow, colIndex + 2, 0.0); // ty_le_3h
                } else {
                    setCellValue(sheet, provinceRow, colIndex, null);
                    setCellValue(sheet, provinceRow, colIndex + 1, null);
                    setCellValue(sheet, provinceRow, colIndex + 2, null);
                }
            }

            // Fill actual handle rate data if available
            Optional<String> matchingKey = dataByProvince.keySet().stream()
                    .filter(key -> {
                        String lowerKey = key.toLowerCase();
                        String lowerName = provinceName.toLowerCase();
                        return lowerKey.contains(lowerName) || lowerName.contains(lowerKey);
                    })
                    .findFirst();

            if (matchingKey.isPresent()) {
                List<Map<String, Object>> provinceData = dataByProvince.get(matchingKey.get());
                int finalProvinceRow = provinceRow;
                provinceData.forEach(data -> {
                    int day = extractDayFromDate(data.get("ngay"));
                    if (day > 0 && day <= 31) {
                        int colIndex = TABLE5_DATA_START_COL + (day - 1) * 3;
                        setCellValue(sheet, finalProvinceRow, colIndex, getNumberValue(data, "da_xu_ly_3h"));
                        setCellValue(sheet, finalProvinceRow, colIndex + 1, getNumberValue(data, "tong_sc_da_xu_ly"));
                        Double tyLe = getNumberValue(data, "ty_le_3h");
                        if (tyLe != null) {
                            setCellValue(sheet, finalProvinceRow, colIndex + 2, tyLe / 100);
                        }
                    }
                });
                log.debug("Filled handle rate data for province: {} with {} records", provinceName, provinceData.size());
            }
        }

        // Handle TOTAL row
        fillHandleRateTotalRow(sheet, dataByProvince, currentDay);
    }

    private void fillHandleRateTotalRow(Sheet sheet, Map<String, List<Map<String, Object>>> dataByProvince, int currentDay) {
        List<Map<String, Object>> totalData = dataByProvince.get("TỔNG");

        // Initialize TOTAL row with default values for past days, null for future days
        for (int day = 1; day <= 31; day++) {
            int colIndex = TABLE5_DATA_START_COL + (day - 1) * 3;
            if (day <= currentDay) {
                setCellValue(sheet, TABLE5_TOTAL_ROW, colIndex, 0.0);     // da_xu_ly_3h
                setCellValue(sheet, TABLE5_TOTAL_ROW, colIndex + 1, 0.0); // tong_sc_da_xu_ly
                setCellValue(sheet, TABLE5_TOTAL_ROW, colIndex + 2, 0.0); // ty_le_3h
            } else {
                setCellValue(sheet, TABLE5_TOTAL_ROW, colIndex, null);
                setCellValue(sheet, TABLE5_TOTAL_ROW, colIndex + 1, null);
                setCellValue(sheet, TABLE5_TOTAL_ROW, colIndex + 2, null);
            }
        }

        // Fill actual TOTAL data if available
        if (totalData != null && !totalData.isEmpty()) {
            totalData.forEach(data -> {
                int day = extractDayFromDate(data.get("ngay"));
                if (day > 0 && day <= 31) {
                    int colIndex = TABLE5_DATA_START_COL + (day - 1) * 3;
                    setCellValue(sheet, TABLE5_TOTAL_ROW, colIndex, getNumberValue(data, "da_xu_ly_3h"));
                    setCellValue(sheet, TABLE5_TOTAL_ROW, colIndex + 1, getNumberValue(data, "tong_sc_da_xu_ly"));
                    Double tyLe = getNumberValue(data, "ty_le_3h");
                    if (tyLe != null) {
                        setCellValue(sheet, TABLE5_TOTAL_ROW, colIndex + 2, tyLe / 100);
                    }
                }
            });
            log.debug("Filled handle rate total data with {} records", totalData.size());
        }
    }

    private void fillHandleRate24hByProvince(Sheet sheet, List<Map<String, Object>> handleRateData) {
        if (handleRateData == null) handleRateData = List.of();

        int currentDay = java.time.LocalDate.now().getDayOfMonth();

        // Create optimized map once - group data by province
        Map<String, List<Map<String, Object>>> dataByProvince = handleRateData.stream()
                .collect(Collectors.groupingBy(
                        d -> d.get("tinh") != null ? d.get("tinh").toString().trim() : "unknown"
                ));

        // Process province rows with optimized logic
        for (int provinceRow = TABLE6_START_ROW; provinceRow <= TABLE6_END_ROW; provinceRow++) {
            String provinceName = getCellStringValue(sheet, provinceRow, CATEGORY_COL);
            if (provinceName == null) continue;

            // Initialize default values for past days, null for future days
            for (int day = 1; day <= 31; day++) {
                int colIndex = TABLE6_DATA_START_COL + (day - 1) * 3;
                if (day <= currentDay) {
                    setCellValue(sheet, provinceRow, colIndex, 0.0);     // da_xu_ly_24h
                    setCellValue(sheet, provinceRow, colIndex + 1, 0.0); // tong_sc_da_xu_ly
                    setCellValue(sheet, provinceRow, colIndex + 2, 0.0); // ty_le_24h
                } else {
                    setCellValue(sheet, provinceRow, colIndex, null);
                    setCellValue(sheet, provinceRow, colIndex + 1, null);
                    setCellValue(sheet, provinceRow, colIndex + 2, null);
                }
            }

            // Fill actual handle rate data if available
            Optional<String> matchingKey = dataByProvince.keySet().stream()
                    .filter(key -> {
                        String lowerKey = key.toLowerCase();
                        String lowerName = provinceName.toLowerCase();
                        return lowerKey.contains(lowerName) || lowerName.contains(lowerKey);
                    })
                    .findFirst();

            if (matchingKey.isPresent()) {
                List<Map<String, Object>> provinceData = dataByProvince.get(matchingKey.get());
                int finalProvinceRow = provinceRow;
                provinceData.forEach(data -> {
                    int day = extractDayFromDate(data.get("ngay"));
                    if (day > 0 && day <= 31) {
                        int colIndex = TABLE6_DATA_START_COL + (day - 1) * 3;
                        setCellValue(sheet, finalProvinceRow, colIndex, getNumberValue(data, "da_xu_ly_24h"));
                        setCellValue(sheet, finalProvinceRow, colIndex + 1, getNumberValue(data, "tong_sc_da_xu_ly"));
                        Double tyLe = getNumberValue(data, "ty_le_24h");
                        if (tyLe != null) {
                            setCellValue(sheet, finalProvinceRow, colIndex + 2, tyLe / 100);
                        }
                    }
                });
                log.debug("Filled 24h handle rate data for province: {} with {} records", provinceName, provinceData.size());
            }
        }

        // Handle TOTAL row
        fillHandleRate24hTotalRow(sheet, dataByProvince, currentDay);
    }

    private void fillHandleRate24hTotalRow(Sheet sheet, Map<String, List<Map<String, Object>>> dataByProvince, int currentDay) {
        List<Map<String, Object>> totalData = dataByProvince.get("TỔNG");

        // Initialize TOTAL row with default values for past days, null for future days
        for (int day = 1; day <= 31; day++) {
            int colIndex = TABLE6_DATA_START_COL + (day - 1) * 3;
            if (day <= currentDay) {
                setCellValue(sheet, TABLE6_TOTAL_ROW, colIndex, 0.0);     // da_xu_ly_24h
                setCellValue(sheet, TABLE6_TOTAL_ROW, colIndex + 1, 0.0); // tong_sc_da_xu_ly
                setCellValue(sheet, TABLE6_TOTAL_ROW, colIndex + 2, 0.0); // ty_le_24h
            } else {
                setCellValue(sheet, TABLE6_TOTAL_ROW, colIndex, null);
                setCellValue(sheet, TABLE6_TOTAL_ROW, colIndex + 1, null);
                setCellValue(sheet, TABLE6_TOTAL_ROW, colIndex + 2, null);
            }
        }

        // Fill actual TOTAL data if available
        if (totalData != null && !totalData.isEmpty()) {
            totalData.forEach(data -> {
                int day = extractDayFromDate(data.get("ngay"));
                if (day > 0 && day <= 31) {
                    int colIndex = TABLE6_DATA_START_COL + (day - 1) * 3;
                    setCellValue(sheet, TABLE6_TOTAL_ROW, colIndex, getNumberValue(data, "da_xu_ly_24h"));
                    setCellValue(sheet, TABLE6_TOTAL_ROW, colIndex + 1, getNumberValue(data, "tong_sc_da_xu_ly"));
                    Double tyLe = getNumberValue(data, "ty_le_24h");
                    if (tyLe != null) {
                        setCellValue(sheet, TABLE6_TOTAL_ROW, colIndex + 2, tyLe / 100);
                    }
                }
            });
            log.debug("Filled 24h handle rate total data with {} records", totalData.size());
        }
    }

    private void fillHandleRate48hByProvince(Sheet sheet, List<Map<String, Object>> handleRateData) {
        if (handleRateData == null) handleRateData = List.of();

        int currentDay = java.time.LocalDate.now().getDayOfMonth();

        // Create optimized map once - group data by province
        Map<String, List<Map<String, Object>>> dataByProvince = handleRateData.stream()
                .collect(Collectors.groupingBy(
                        d -> d.get("tinh") != null ? d.get("tinh").toString().trim() : "unknown"
                ));

        // Process province rows with optimized logic
        for (int provinceRow = TABLE7_START_ROW; provinceRow <= TABLE7_END_ROW; provinceRow++) {
            String provinceName = getCellStringValue(sheet, provinceRow, CATEGORY_COL);
            if (provinceName == null) continue;

            // Initialize default values for past days, null for future days
            for (int day = 1; day <= 31; day++) {
                int colIndex = TABLE7_DATA_START_COL + (day - 1) * 3;
                if (day <= currentDay) {
                    setCellValue(sheet, provinceRow, colIndex, 0.0);     // da_xu_ly_48h
                    setCellValue(sheet, provinceRow, colIndex + 1, 0.0); // tong_sc_da_xu_ly
                    setCellValue(sheet, provinceRow, colIndex + 2, 0.0); // ty_le_48h
                } else {
                    setCellValue(sheet, provinceRow, colIndex, null);
                    setCellValue(sheet, provinceRow, colIndex + 1, null);
                    setCellValue(sheet, provinceRow, colIndex + 2, null);
                }
            }

            // Fill actual handle rate data if available
            Optional<String> matchingKey = dataByProvince.keySet().stream()
                    .filter(key -> {
                        String lowerKey = key.toLowerCase();
                        String lowerName = provinceName.toLowerCase();
                        return lowerKey.contains(lowerName) || lowerName.contains(lowerKey);
                    })
                    .findFirst();

            if (matchingKey.isPresent()) {
                List<Map<String, Object>> provinceData = dataByProvince.get(matchingKey.get());
                int finalProvinceRow = provinceRow;
                provinceData.forEach(data -> {
                    int day = extractDayFromDate(data.get("ngay"));
                    if (day > 0 && day <= 31) {
                        int colIndex = TABLE7_DATA_START_COL + (day - 1) * 3;
                        setCellValue(sheet, finalProvinceRow, colIndex, getNumberValue(data, "da_xu_ly_48h"));
                        setCellValue(sheet, finalProvinceRow, colIndex + 1, getNumberValue(data, "tong_sc_da_xu_ly"));
                        Double tyLe = getNumberValue(data, "ty_le_48h");
                        if (tyLe != null) {
                            setCellValue(sheet, finalProvinceRow, colIndex + 2, tyLe / 100);
                        }
                    }
                });
                log.debug("Filled 48h handle rate data for province: {} with {} records", provinceName, provinceData.size());
            }
        }

        // Handle TOTAL row
        fillHandleRate48hTotalRow(sheet, dataByProvince, currentDay);
    }

    private void fillHandleRate48hTotalRow(Sheet sheet, Map<String, List<Map<String, Object>>> dataByProvince, int currentDay) {
        List<Map<String, Object>> totalData = dataByProvince.get("TỔNG");

        // Initialize TOTAL row with default values for past days, null for future days
        for (int day = 1; day <= 31; day++) {
            int colIndex = TABLE7_DATA_START_COL + (day - 1) * 3;
            if (day <= currentDay) {
                setCellValue(sheet, TABLE7_TOTAL_ROW, colIndex, 0.0);     // da_xu_ly_48h
                setCellValue(sheet, TABLE7_TOTAL_ROW, colIndex + 1, 0.0); // tong_sc_da_xu_ly
                setCellValue(sheet, TABLE7_TOTAL_ROW, colIndex + 2, 0.0); // ty_le_48h
            } else {
                setCellValue(sheet, TABLE7_TOTAL_ROW, colIndex, null);
                setCellValue(sheet, TABLE7_TOTAL_ROW, colIndex + 1, null);
                setCellValue(sheet, TABLE7_TOTAL_ROW, colIndex + 2, null);
            }
        }

        // Fill actual TOTAL data if available
        if (totalData != null && !totalData.isEmpty()) {
            totalData.forEach(data -> {
                int day = extractDayFromDate(data.get("ngay"));
                if (day > 0 && day <= 31) {
                    int colIndex = TABLE7_DATA_START_COL + (day - 1) * 3;
                    setCellValue(sheet, TABLE7_TOTAL_ROW, colIndex, getNumberValue(data, "da_xu_ly_48h"));
                    setCellValue(sheet, TABLE7_TOTAL_ROW, colIndex + 1, getNumberValue(data, "tong_sc_da_xu_ly"));
                    Double tyLe = getNumberValue(data, "ty_le_48h");
                    if (tyLe != null) {
                        setCellValue(sheet, TABLE7_TOTAL_ROW, colIndex + 2, tyLe / 100);
                    }
                }
            });
            log.debug("Filled 48h handle rate total data with {} records", totalData.size());
        }
    }

    private void fillHandleRate3hVipByProvince(Sheet sheet, List<Map<String, Object>> handleRateData) {
        if (handleRateData == null) handleRateData = List.of();

        int currentDay = java.time.LocalDate.now().getDayOfMonth();

        // Create optimized map once - group data by province
        Map<String, List<Map<String, Object>>> dataByProvince = handleRateData.stream()
                .collect(Collectors.groupingBy(
                        d -> d.get("tinh") != null ? d.get("tinh").toString().trim() : "unknown"
                ));

        // Process province rows with optimized logic
        for (int provinceRow = TABLE8_START_ROW; provinceRow <= TABLE8_END_ROW; provinceRow++) {
            String provinceName = getCellStringValue(sheet, provinceRow, CATEGORY_COL);
            if (provinceName == null) continue;

            // Initialize default values for past days, null for future days
            for (int day = 1; day <= 31; day++) {
                int colIndex = TABLE8_DATA_START_COL + (day - 1) * 3;
                if (day <= currentDay) {
                    setCellValue(sheet, provinceRow, colIndex, 0.0);     // da_xu_ly_3h_vip
                    setCellValue(sheet, provinceRow, colIndex + 1, 0.0); // tong_sc_da_xu_ly_vip
                    setCellValue(sheet, provinceRow, colIndex + 2, 0.0); // ty_le_3h_vip
                } else {
                    setCellValue(sheet, provinceRow, colIndex, null);
                    setCellValue(sheet, provinceRow, colIndex + 1, null);
                    setCellValue(sheet, provinceRow, colIndex + 2, null);
                }
            }

            // Fill actual handle rate data if available
            Optional<String> matchingKey = dataByProvince.keySet().stream()
                    .filter(key -> {
                        String lowerKey = key.toLowerCase();
                        String lowerName = provinceName.toLowerCase();
                        return lowerKey.contains(lowerName) || lowerName.contains(lowerKey);
                    })
                    .findFirst();

            if (matchingKey.isPresent()) {
                List<Map<String, Object>> provinceData = dataByProvince.get(matchingKey.get());
                int finalProvinceRow = provinceRow;
                provinceData.forEach(data -> {
                    int day = extractDayFromDate(data.get("ngay"));
                    if (day > 0 && day <= 31) {
                        int colIndex = TABLE8_DATA_START_COL + (day - 1) * 3;
                        setCellValue(sheet, finalProvinceRow, colIndex, getNumberValue(data, "da_xu_ly_3h_vip"));
                        setCellValue(sheet, finalProvinceRow, colIndex + 1, getNumberValue(data, "tong_sc_da_xu_ly_vip"));
                        Double tyLe = getNumberValue(data, "ty_le_3h_vip");
                        if (tyLe != null) {
                            setCellValue(sheet, finalProvinceRow, colIndex + 2, tyLe / 100);
                        }
                    }
                });
                log.debug("Filled 3h VIP handle rate data for province: {} with {} records", provinceName, provinceData.size());
            }
        }

        // Handle TOTAL row
        fillHandleRate3hVipTotalRow(sheet, dataByProvince, currentDay);
    }

    private void fillHandleRate3hVipTotalRow(Sheet sheet, Map<String, List<Map<String, Object>>> dataByProvince, int currentDay) {
        List<Map<String, Object>> totalData = dataByProvince.get("TỔNG");

        // Initialize TOTAL row with default values for past days, null for future days
        for (int day = 1; day <= 31; day++) {
            int colIndex = TABLE8_DATA_START_COL + (day - 1) * 3;
            if (day <= currentDay) {
                setCellValue(sheet, TABLE8_TOTAL_ROW, colIndex, 0.0);     // da_xu_ly_3h_vip
                setCellValue(sheet, TABLE8_TOTAL_ROW, colIndex + 1, 0.0); // tong_sc_da_xu_ly_vip
                setCellValue(sheet, TABLE8_TOTAL_ROW, colIndex + 2, 0.0); // ty_le_3h_vip
            } else {
                setCellValue(sheet, TABLE8_TOTAL_ROW, colIndex, null);
                setCellValue(sheet, TABLE8_TOTAL_ROW, colIndex + 1, null);
                setCellValue(sheet, TABLE8_TOTAL_ROW, colIndex + 2, null);
            }
        }

        // Fill actual TOTAL data if available
        if (totalData != null && !totalData.isEmpty()) {
            totalData.forEach(data -> {
                int day = extractDayFromDate(data.get("ngay"));
                if (day > 0 && day <= 31) {
                    int colIndex = TABLE8_DATA_START_COL + (day - 1) * 3;
                    setCellValue(sheet, TABLE8_TOTAL_ROW, colIndex, getNumberValue(data, "da_xu_ly_3h_vip"));
                    setCellValue(sheet, TABLE8_TOTAL_ROW, colIndex + 1, getNumberValue(data, "tong_sc_da_xu_ly_vip"));
                    Double tyLe = getNumberValue(data, "ty_le_3h_vip");
                    if (tyLe != null) {
                        setCellValue(sheet, TABLE8_TOTAL_ROW, colIndex + 2, tyLe / 100);
                    }
                }
            });
            log.debug("Filled 3h VIP handle rate total data with {} records", totalData.size());
        }
    }

    private int extractDayFromDate(Object dateObj) {
        try {
            if (dateObj instanceof java.sql.Date) {
                return ((java.sql.Date) dateObj).toLocalDate().getDayOfMonth();
            } else if (dateObj instanceof java.util.Date) {
                return ((java.util.Date) dateObj).toInstant()
                        .atZone(java.time.ZoneId.systemDefault())
                        .toLocalDate()
                        .getDayOfMonth();
            } else if (dateObj instanceof java.time.LocalDate) {
                return ((java.time.LocalDate) dateObj).getDayOfMonth();
            }
        } catch (Exception e) {
            log.warn("Failed to extract day from date object: {}", dateObj, e);
        }
        return -1; // Invalid day
    }

    // Utility methods
    private Optional<Map<String, Object>> findMatchingCategory(List<Map<String, Object>> dataList,
                                                               String templateCategory, String categoryKey) {
        return dataList.stream()
                .filter(data -> {
                    Object categoryObj = data.get(categoryKey);
                    return categoryObj != null &&
                            categoryObj.toString().trim().toLowerCase()
                                    .contains(templateCategory.toLowerCase());
                })
                .findFirst();
    }

    private Optional<Map<String, Object>> findCategoryData(List<Map<String, Object>> dataList,
                                                           String categoryValue, String categoryKey) {
        if (dataList == null || dataList.isEmpty()) return Optional.empty();

        return dataList.stream()
                .filter(data -> {
                    Object categoryObj = data.get(categoryKey);
                    return categoryObj != null &&
                            categoryValue.equalsIgnoreCase(categoryObj.toString().trim());
                })
                .findFirst();
    }

    private String getCellStringValue(Sheet sheet, int rowIndex, int colIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) return null;

        Cell cell = row.getCell(colIndex);
        if (cell == null) return null;

        return cell.getStringCellValue().trim();
    }

    private void setCellValue(Sheet sheet, int rowIndex, int colIndex, Number value) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) row = sheet.createRow(rowIndex);

        Cell cell = row.getCell(colIndex);
        if (cell == null) cell = row.createCell(colIndex);

        if (value == null) {
            cell.setBlank();
        } else {
            cell.setCellValue(value.doubleValue());
        }
    }

    private Double getNumberValue(Map<String, Object> data, String key) {
        Object value = data.get(key);
        return value instanceof Number ? ((Number) value).doubleValue() : null;
    }
}