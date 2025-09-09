package com.example.transmission.service;

import com.example.transmission.domain.ServiceSubscriber;
import com.example.transmission.dto.ExcelDTO;
import com.example.transmission.repository.ProvinceSubscriberRepository;
import com.example.transmission.repository.ServiceSubscriberRepository;
import com.example.transmission.repository.TransmissionChannelIncidentRepository;
import com.spire.xls.Workbook;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import com.spire.xls.*;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
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
    private static final String TEMPLATE_PATH_INPUT = "/templates/template_input.xlsx";
    private static final String SHEET_NAME = "PL06-Kenh truyen";
    private static final String SHEET_NAME_2 = "KPI VTS";
    private static final String SHEET_NAME_IN = "Phụ lục BC ngay";
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
    private static final int TABLE5_KPI_ROW = 81; // G82 = row 81 (0-based)
    private static final int TABLE5_KPI_COL = 6;  // Column G = index 6

    // Table 6 constants (Handle Rate by Province)
    private static final int TABLE6_START_ROW = 124;
    private static final int TABLE6_KPI_ROW = 121;
    private static final int TABLE6_END_ROW = 157;
    private static final int TABLE6_DATA_START_COL = 7;
    private static final int TABLE6_TOTAL_ROW = 123;

    // Table 7 constants (Handle Rate by Province)
    private static final int TABLE7_START_ROW = 164;
    private static final int TABLE7_KPI_ROW = 161;
    private static final int TABLE7_END_ROW = 197;
    private static final int TABLE7_DATA_START_COL = 7;
    private static final int TABLE7_TOTAL_ROW = 163;

    // Table 8 constants (Handle Rate by Province)
    private static final int TABLE8_START_ROW = 204;
    private static final int TABLE8_KPI_ROW = 201;
    private static final int TABLE8_END_ROW = 237;
    private static final int TABLE8_DATA_START_COL = 7;
    private static final int TABLE8_TOTAL_ROW = 203;

    // Add these constants to your existing constants section
    private static final int SHEET2_KPI_3H_ROW = 4;    // C5 = row 4 (0-based)
    private static final int SHEET2_KPI_24H_ROW = 10;   // C11 = row 10 (0-based)
    private static final int SHEET2_KPI_48H_ROW = 16;   // C17 = row 16 (0-based)
    private static final int SHEET2_KPI_SATISFY_ROW = 45;
    private static final int SHEET2_DATA_START_COL = 2; // Column C = index 2
    private static final int SHEET2_DATA_END_COL = 36;

    private static final int SHEET2_3H_DA_XU_LI_ROW = 6;
    private static final int SHEET2_3H_TONG_SC_ROW = SHEET2_3H_DA_XU_LI_ROW + 1;
    private static final int SHEET2_24H_DA_XU_LI_ROW = 12;
    private static final int SHEET2_24H_TONG_SC_ROW = SHEET2_24H_DA_XU_LI_ROW + 1;
    private static final int SHEET2_48H_DA_XU_LI_ROW = 18;
    private static final int SHEET2_48H_TONG_SC_ROW = SHEET2_48H_DA_XU_LI_ROW + 1;
    public byte[] exportExcelFile() {
        try (InputStream inputStream = getClass().getResourceAsStream(TEMPLATE_PATH);
             XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
             ByteArrayOutputStream out = new ByteArrayOutputStream()) {

            Sheet sheet = workbook.getSheet(SHEET_NAME);
            Sheet sheet2 = workbook.getSheet(SHEET_NAME_2);

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
            fillSheet2KpiData(sheet2, excelData);
            fillSatisfactionLevel(sheet2, excelData);
            fillAvgHandleTime(sheet2, excelData);

            String inputChartPath = extractResourceToTemp(TEMPLATE_PATH_INPUT);
            BufferedImage chartImage = exportChartToImage(inputChartPath, SHEET_NAME_IN);
            insertChartImage(workbook, SHEET_NAME, chartImage);

            workbook.setForceFormulaRecalculation(true);

            workbook.write(out);
            return out.toByteArray();

        } catch (IOException e) {
            log.error("Error while exporting Excel", e);
            throw new RuntimeException("Error while exporting Excel", e);
        } catch (Exception e) {
            throw new RuntimeException(e);
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
                .satisfactionLevel(transmissionChannelIncidentRepository.getLevelSatisfy())
                .avgTimeHandle(transmissionChannelIncidentRepository.getAvgHandleTime())
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

            final int currentRowIndex = rowIndex;

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
        List<Map<String, Object>> handleRateData = data.getHandleRate3h();
        Double progress3h = null;
        if (data.getProgressKpi() != null && !data.getProgressKpi().isEmpty()) {
            Double val = getNumberValue(data.getProgressKpi().get(0), "progress_3h");
            if (val != null) progress3h = val / 100;
        }
        log.debug("Progress KPI 3h: {}", progress3h);

        fillHandleRate3hByProvince(sheet, handleRateData, progress3h);
        fillTable5Kpi(sheet, data.getProgressKpi());
    }

    private void fillTable6(Sheet sheet, ExcelDTO data) {
        List<Map<String, Object>> handleRateData = data.getHandleRate24h();
        Double progress24h = null;
        if (data.getProgressKpi() != null && !data.getProgressKpi().isEmpty()) {
            Double val = getNumberValue(data.getProgressKpi().get(0), "progress_24h");
            if (val != null) progress24h = val / 100;
        }
        fillHandleRate24hByProvince(sheet, handleRateData, progress24h);
        fillTable6Kpi(sheet, data.getProgressKpi());
    }

    private void fillTable7(Sheet sheet, ExcelDTO data) {
        List<Map<String, Object>> handleRateData = data.getHandleRate48h();
        Double progress48h = null;
        if (data.getProgressKpi() != null && !data.getProgressKpi().isEmpty()) {
            Double val = getNumberValue(data.getProgressKpi().get(0), "progress_48h");
            if (val != null) progress48h = val / 100;
        }
        fillHandleRate48hByProvince(sheet, handleRateData, progress48h);
        fillTable7Kpi(sheet, data.getProgressKpi());
    }

    private void fillTable8(Sheet sheet, ExcelDTO data) {
        List<Map<String, Object>> handleRateData = data.getHandleRate3hVip();
        Double progress3hVip = null;
        if (data.getProgressKpi() != null && !data.getProgressKpi().isEmpty()) {
            Double val = getNumberValue(data.getProgressKpi().get(0), "progress_3hVip");
            if (val != null) progress3hVip = val / 100;
        }
        fillHandleRate3hVipByProvince(sheet, handleRateData, progress3hVip);
        fillTable8Kpi(sheet, data.getProgressKpi());
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
        Double totalSltb = provinceSubscriberMap.getOrDefault(TOTAL, 0.0);
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
        List<Map<String, Object>> totalData = dataByProvince.get(TOTAL);

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

    private void fillTable5Kpi(Sheet sheet, List<Map<String, Object>> kpiData) {
        if (kpiData == null || kpiData.isEmpty()) return;

        Double progress3h = getNumberValue(kpiData.get(0), "progress_3h");
        if (progress3h != null) {
            String text = String.format("Đánh giá so với KPI (>=%.2f%%)", progress3h);
            setCellValueTable5(sheet, text);
            log.debug("Filled Table 5 KPI: {}", text);
        }
    }
    private void fillTable6Kpi(Sheet sheet, List<Map<String, Object>> kpiData) {
        if (kpiData == null || kpiData.isEmpty()) return;

        Double progress24h = getNumberValue(kpiData.get(0), "progress_24h");
        if (progress24h != null) {
            String text = String.format("Đánh giá so với KPI (>=%.2f%%)", progress24h);
            setCellValueTable6(sheet, text);
            log.debug("Filled Table 6 KPI: {}", text);
        }
    }
    private void fillTable7Kpi(Sheet sheet, List<Map<String, Object>> kpiData) {
        if (kpiData == null || kpiData.isEmpty()) return;

        Double progress48h = getNumberValue(kpiData.get(0), "progress_48h");
        if (progress48h != null) {
            String text = String.format("Đánh giá so với KPI (>=%.2f%%)", progress48h);
            setCellValueTable7(sheet, text);
            log.debug("Filled Table 7 KPI: {}", text);
        }
    }
    private void fillTable8Kpi(Sheet sheet, List<Map<String, Object>> kpiData) {
        if (kpiData == null || kpiData.isEmpty()) return;

        Double progress3hVip = getNumberValue(kpiData.get(0), "progress_3hVip");
        if (progress3hVip != null) {
            String text = String.format("Đánh giá so với KPI (>=%.2f%%)", progress3hVip);
            setCellValueTable8(sheet, text);
            log.debug("Filled Table 8 KPI: {}", text);
        }
    }

    private void fillHandleRate3hByProvince(Sheet sheet, List<Map<String, Object>> handleRateData, Double progress3h) {
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

            // Initialize cumulative columns D, E, F (assuming these are fixed columns)
            setCellValue(sheet, provinceRow, 3, 0.0);  // Column D (luy_ke_3h)
            setCellValue(sheet, provinceRow, 4, 0.0);  // Column E (tong_sc_luy_ke_3h)
            setCellValue(sheet, provinceRow, 5, 0.0);  // Column F (ty_le_luy_ke_3h)
            setCellValueString(sheet, provinceRow, 6, "Không đạt");  // Column G

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

                // Calculate and fill cumulative data for columns D, E, F
                fillCumulativeDataTable5(sheet, finalProvinceRow, provinceData, progress3h);
                log.debug("Filled handle rate data for province: {} with {} records", provinceName, provinceData.size());
            }
        }
        // Handle TOTAL row
        fillHandleRateTotalRow(sheet, dataByProvince, currentDay, progress3h);
    }
    private void fillCumulativeDataTable5(Sheet sheet, int row, List<Map<String, Object>> provinceData, Double progress3h) {
        if (provinceData == null || provinceData.isEmpty()) {
            setCellValue(sheet, row, 3, 0.0);  // Lũy kế 3h
            setCellValue(sheet, row, 4, 0.0);  // Tổng SC lũy kế
            setCellValue(sheet, row, 5, 0.0);  // Tỷ lệ lũy kế
            setCellValueString(sheet, row, 6, "N/A"); // Cột G KPI
            return;
        }

        double cumulativeDaXuLy3h = 0.0;
        double cumulativeTongSc = 0.0;

        for (Map<String, Object> data : provinceData) {
            cumulativeDaXuLy3h += getNumberValue(data, "da_xu_ly_3h") != null
                    ? getNumberValue(data, "da_xu_ly_3h") : 0.0;
            cumulativeTongSc += getNumberValue(data, "tong_sc_da_xu_ly") != null
                    ? getNumberValue(data, "tong_sc_da_xu_ly") : 0.0;
        }

        double cumulativePercent = cumulativeTongSc > 0
                ? (cumulativeDaXuLy3h / cumulativeTongSc)
                : 0.0;

        // Fill D, E, F
        setCellValue(sheet, row, 3, cumulativeDaXuLy3h);
        setCellValue(sheet, row, 4, cumulativeTongSc);
        setCellValue(sheet, row, 5, cumulativePercent);

        // Fill G (KPI đạt hay không)
        if (progress3h != null) {
            String kpiStatus = cumulativePercent >= progress3h ? "Đạt" : "Không đạt";
            setCellValueString(sheet, row, 6, kpiStatus);
        } else {
            setCellValueString(sheet, row, 6, "N/A");
        }
    }
    private void fillCumulativeDataTable6(Sheet sheet, int row, List<Map<String, Object>> provinceData, Double progress3h) {
        if (provinceData == null || provinceData.isEmpty()) {
            setCellValue(sheet, row, 3, 0.0);  // Lũy kế 3h
            setCellValue(sheet, row, 4, 0.0);  // Tổng SC lũy kế
            setCellValue(sheet, row, 5, 0.0);  // Tỷ lệ lũy kế
            setCellValueString(sheet, row, 6, "N/A"); // Cột G KPI
            return;
        }

        double cumulativeDaXuLy3h = 0.0;
        double cumulativeTongSc = 0.0;

        for (Map<String, Object> data : provinceData) {
            cumulativeDaXuLy3h += getNumberValue(data, "da_xu_ly_24h") != null
                    ? getNumberValue(data, "da_xu_ly_24h") : 0.0;
            cumulativeTongSc += getNumberValue(data, "tong_sc_da_xu_ly") != null
                    ? getNumberValue(data, "tong_sc_da_xu_ly") : 0.0;
        }

        double cumulativePercent = cumulativeTongSc > 0
                ? (cumulativeDaXuLy3h / cumulativeTongSc)
                : 0.0;

        // Fill D, E, F
        setCellValue(sheet, row, 3, cumulativeDaXuLy3h);
        setCellValue(sheet, row, 4, cumulativeTongSc);
        setCellValue(sheet, row, 5, cumulativePercent);

        // Fill G (KPI đạt hay không)
        if (progress3h != null) {
            String kpiStatus = cumulativePercent >= progress3h ? "Đạt" : "Không đạt";
            setCellValueString(sheet, row, 6, kpiStatus);
        } else {
            setCellValueString(sheet, row, 6, "N/A");
        }
    }

    private void fillHandleRateTotalRow(Sheet sheet, Map<String, List<Map<String, Object>>> dataByProvince, int currentDay, Double progress3h) {
        List<Map<String, Object>> totalData = dataByProvince.get(TOTAL);

        // Initialize TOTAL row with default values for past days, null for future days
        for (int day = 1; day <= 31; day++) {
            int colIndex = TABLE5_DATA_START_COL - 4 + (day - 1) * 3;
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

        double cumulativeDaXuLy3h = 0.0;
        double cumulativeTongSc = 0.0;

        if (totalData != null && !totalData.isEmpty()) {
            for (Map<String, Object> data : totalData) {
                cumulativeDaXuLy3h += getNumberValue(data, "da_xu_ly_3h") != null ? getNumberValue(data, "da_xu_ly_3h") : 0.0;
                cumulativeTongSc += getNumberValue(data, "tong_sc_da_xu_ly") != null ? getNumberValue(data, "tong_sc_da_xu_ly") : 0.0;
            }
            for (Map<String, Object> data : totalData) {
                int day = extractDayFromDate(data.get("ngay"));
                if (day > 0 && day <= 31) {
                    int colIndex = TABLE5_DATA_START_COL + (day - 1) * 3;
                    setCellValue(sheet, TABLE5_TOTAL_ROW, colIndex, getNumberValue(data, "da_xu_ly_3h"));
                    setCellValue(sheet, TABLE5_TOTAL_ROW, colIndex + 1, getNumberValue(data, "tong_sc_da_xu_ly"));
                    Double tyLe = getNumberValue(data, "ty_le_3h");
                    if (tyLe != null) setCellValue(sheet, TABLE5_TOTAL_ROW, colIndex + 2, tyLe / 100);
                }
            }
        }

        // Fill cumulative columns D, E, F
        setCellValue(sheet, TABLE5_TOTAL_ROW, 3, cumulativeDaXuLy3h);
        setCellValue(sheet, TABLE5_TOTAL_ROW, 4, cumulativeTongSc);
        double cumulativePercent = cumulativeTongSc > 0 ? (cumulativeDaXuLy3h / cumulativeTongSc) : 0.0;
        setCellValue(sheet, TABLE5_TOTAL_ROW, 5, cumulativePercent);

        // Fill KPI vào cột G
        if (progress3h != null) {
            String kpiStatus = cumulativePercent >= progress3h ? "Đạt" : "Không đạt";
            setCellValueString(sheet, TABLE5_TOTAL_ROW, 6, kpiStatus);
        } else {
            setCellValueString(sheet, TABLE5_TOTAL_ROW, 6, "N/A");
        }

        log.debug("Filled handle rate TOTAL row with cumulative KPI");
    }
    private void fillHandleRateTotalRowTable6(Sheet sheet, Map<String, List<Map<String, Object>>> dataByProvince, int currentDay, Double progress3h) {
        List<Map<String, Object>> totalData = dataByProvince.get(TOTAL);

        // Initialize TOTAL row with default values for past days, null for future days
        for (int day = 1; day <= 31; day++) {
            int colIndex = TABLE6_DATA_START_COL - 4 + (day - 1) * 3;
            if (day <= currentDay) {
                setCellValue(sheet, TABLE6_TOTAL_ROW, colIndex, 0.0);     // da_xu_ly_3h
                setCellValue(sheet, TABLE6_TOTAL_ROW, colIndex + 1, 0.0); // tong_sc_da_xu_ly
                setCellValue(sheet, TABLE6_TOTAL_ROW, colIndex + 2, 0.0); // ty_le_3h
            } else {
                setCellValue(sheet, TABLE6_TOTAL_ROW, colIndex, null);
                setCellValue(sheet, TABLE6_TOTAL_ROW, colIndex + 1, null);
                setCellValue(sheet, TABLE6_TOTAL_ROW, colIndex + 2, null);
            }
        }

        double cumulativeDaXuLy24h = 0.0;
        double cumulativeTongSc = 0.0;

        if (totalData != null && !totalData.isEmpty()) {
            for (Map<String, Object> data : totalData) {
                cumulativeDaXuLy24h += getNumberValue(data, "da_xu_ly_24h") != null ? getNumberValue(data, "da_xu_ly_24h") : 0.0;
                cumulativeTongSc += getNumberValue(data, "tong_sc_da_xu_ly") != null ? getNumberValue(data, "tong_sc_da_xu_ly") : 0.0;
            }
            for (Map<String, Object> data : totalData) {
                int day = extractDayFromDate(data.get("ngay"));
                if (day > 0 && day <= 31) {
                    int colIndex = TABLE6_DATA_START_COL + (day - 1) * 3;
                    setCellValue(sheet, TABLE6_TOTAL_ROW, colIndex, getNumberValue(data, "da_xu_ly_24h"));
                    setCellValue(sheet, TABLE6_TOTAL_ROW, colIndex + 1, getNumberValue(data, "tong_sc_da_xu_ly"));
                    Double tyLe = getNumberValue(data, "ty_le_24h");
                    if (tyLe != null) setCellValue(sheet, TABLE6_TOTAL_ROW, colIndex + 2, tyLe / 100);
                }
            }
        }

        // Fill cumulative columns D, E, F
        setCellValue(sheet, TABLE6_TOTAL_ROW, 3, cumulativeDaXuLy24h);
        setCellValue(sheet, TABLE6_TOTAL_ROW, 4, cumulativeTongSc);
        double cumulativePercent = cumulativeTongSc > 0 ? (cumulativeDaXuLy24h / cumulativeTongSc) : 0.0;
        setCellValue(sheet, TABLE6_TOTAL_ROW, 5, cumulativePercent);

        // Fill KPI vào cột G
        if (progress3h != null) {
            String kpiStatus = cumulativePercent >= progress3h ? "Đạt" : "Không đạt";
            setCellValueString(sheet, TABLE6_TOTAL_ROW, 6, kpiStatus);
        } else {
            setCellValueString(sheet, TABLE6_TOTAL_ROW, 6, "N/A");
        }
    }
    private void fillHandleRate24hByProvince(Sheet sheet, List<Map<String, Object>> handleRateData, Double progress24h) {
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
                    setCellValueString(sheet, provinceRow, 6, "Không đạt");  // Column G

                }
            }
            // Initialize cumulative columns D, E, F (assuming these are fixed columns)
            setCellValue(sheet, provinceRow, 3, 0.0);
            setCellValue(sheet, provinceRow, 4, 0.0);
            setCellValue(sheet, provinceRow, 5, 0.0);
            setCellValueString(sheet, provinceRow, 6, "Không đạt");  // Column G

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
                fillCumulativeDataTable6(sheet, finalProvinceRow, provinceData, progress24h);

                log.debug("Filled 24h handle rate data for province: {} with {} records", provinceName, provinceData.size());
            }
        }

        // Handle TOTAL row
        fillHandleRateTotalRowTable6(sheet, dataByProvince, currentDay, progress24h);
    }

    private void fillHandleRate48hByProvince(Sheet sheet, List<Map<String, Object>> handleRateData, Double progress48h) {
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
                fillCumulativeDataTable7(sheet, finalProvinceRow, provinceData, progress48h);

                log.debug("Filled 48h handle rate data for province: {} with {} records", provinceName, provinceData.size());
            }
        }

        // Handle TOTAL row
        fillHandleRateTotalRowTable7(sheet, dataByProvince, currentDay, progress48h);
    }
    private void fillCumulativeDataTable7(Sheet sheet, int row, List<Map<String, Object>> provinceData, Double progress48h) {
        if (provinceData == null || provinceData.isEmpty()) {
            setCellValue(sheet, row, 3, 0.0);  // Lũy kế h
            setCellValue(sheet, row, 4, 0.0);  // Tổng SC lũy kế
            setCellValue(sheet, row, 5, 0.0);  // Tỷ lệ lũy kế
            setCellValueString(sheet, row, 6, "N/A"); // Cột G KPI
            return;
        }

        double cumulativeDaXuLy48h = 0.0;
        double cumulativeTongSc = 0.0;

        for (Map<String, Object> data : provinceData) {
            cumulativeDaXuLy48h += getNumberValue(data, "da_xu_ly_48h") != null
                    ? getNumberValue(data, "da_xu_ly_48h") : 0.0;
            cumulativeTongSc += getNumberValue(data, "tong_sc_da_xu_ly") != null
                    ? getNumberValue(data, "tong_sc_da_xu_ly") : 0.0;
        }

        double cumulativePercent = cumulativeTongSc > 0
                ? (cumulativeDaXuLy48h / cumulativeTongSc)
                : 0.0;

        // Fill D, E, F
        setCellValue(sheet, row, 3, cumulativeDaXuLy48h);
        setCellValue(sheet, row, 4, cumulativeTongSc);
        setCellValue(sheet, row, 5, cumulativePercent);

        // Fill G (KPI đạt hay không)
        if (progress48h != null) {
            String kpiStatus = cumulativePercent >= progress48h ? "Đạt" : "Không đạt";
            setCellValueString(sheet, row, 6, kpiStatus);
        } else {
            setCellValueString(sheet, row, 6, "N/A");
        }
    }
    private void fillHandleRateTotalRowTable7(Sheet sheet, Map<String, List<Map<String, Object>>> dataByProvince, int currentDay, Double progress3h) {
        List<Map<String, Object>> totalData = dataByProvince.get(TOTAL);

        // Initialize TOTAL row with default values for past days, null for future days
        for (int day = 1; day <= 31; day++) {
            int colIndex = TABLE7_DATA_START_COL - 4 + (day - 1) * 3;
            if (day <= currentDay) {
                setCellValue(sheet, TABLE7_TOTAL_ROW, colIndex, 0.0);     // da_xu_ly_3h
                setCellValue(sheet, TABLE7_TOTAL_ROW, colIndex + 1, 0.0); // tong_sc_da_xu_ly
                setCellValue(sheet, TABLE7_TOTAL_ROW, colIndex + 2, 0.0); // ty_le_3h
            } else {
                setCellValue(sheet, TABLE7_TOTAL_ROW, colIndex, null);
                setCellValue(sheet, TABLE7_TOTAL_ROW, colIndex + 1, null);
                setCellValue(sheet, TABLE7_TOTAL_ROW, colIndex + 2, null);
            }
        }

        double cumulativeDaXuLy48h = 0.0;
        double cumulativeTongSc = 0.0;

        if (totalData != null && !totalData.isEmpty()) {
            for (Map<String, Object> data : totalData) {
                cumulativeDaXuLy48h += getNumberValue(data, "da_xu_ly_48h") != null ? getNumberValue(data, "da_xu_ly_48h") : 0.0;
                cumulativeTongSc += getNumberValue(data, "tong_sc_da_xu_ly") != null ? getNumberValue(data, "tong_sc_da_xu_ly") : 0.0;
            }
            for (Map<String, Object> data : totalData) {
                int day = extractDayFromDate(data.get("ngay"));
                if (day > 0 && day <= 31) {
                    int colIndex = TABLE7_DATA_START_COL + (day - 1) * 3;
                    setCellValue(sheet, TABLE7_TOTAL_ROW, colIndex, getNumberValue(data, "da_xu_ly_48h"));
                    setCellValue(sheet, TABLE7_TOTAL_ROW, colIndex + 1, getNumberValue(data, "tong_sc_da_xu_ly"));
                    Double tyLe = getNumberValue(data, "ty_le_48h");
                    if (tyLe != null) setCellValue(sheet, TABLE7_TOTAL_ROW, colIndex + 2, tyLe / 100);
                }
            }
        }

        // Fill cumulative columns D, E, F
        setCellValue(sheet, TABLE7_TOTAL_ROW, 3, cumulativeDaXuLy48h);
        setCellValue(sheet, TABLE7_TOTAL_ROW, 4, cumulativeTongSc);
        double cumulativePercent = cumulativeTongSc > 0 ? (cumulativeDaXuLy48h / cumulativeTongSc) : 0.0;
        setCellValue(sheet, TABLE7_TOTAL_ROW, 5, cumulativePercent);

        // Fill KPI vào cột G
        if (progress3h != null) {
            String kpiStatus = cumulativePercent >= progress3h ? "Đạt" : "Không đạt";
            setCellValueString(sheet, TABLE7_TOTAL_ROW, 6, kpiStatus);
        } else {
            setCellValueString(sheet, TABLE7_TOTAL_ROW, 6, "N/A");
        }
    }

    private void fillHandleRate3hVipByProvince(Sheet sheet, List<Map<String, Object>> handleRateData, Double progress3hVip) {
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
                    setCellValueString(sheet, provinceRow, 6, "Không đạt");  // Column G
                }
            }
            // Initialize cumulative columns D, E, F (assuming these are fixed columns)
            setCellValue(sheet, provinceRow, 3, 0.0);
            setCellValue(sheet, provinceRow, 4, 0.0);
            setCellValue(sheet, provinceRow, 5, 0.0);
            setCellValueString(sheet, provinceRow, 6, "Không đạt");  // Column G

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
                fillCumulativeDataTable8(sheet, finalProvinceRow, provinceData, progress3hVip);

                log.debug("Filled 3h VIP handle rate data for province: {} with {} records", provinceName, provinceData.size());
            }
        }

        // Handle TOTAL row
        fillHandleRateTotalRowTable8(sheet, dataByProvince, currentDay, progress3hVip);
    }
    private void fillCumulativeDataTable8(Sheet sheet, int row, List<Map<String, Object>> provinceData, Double progress3hVip) {
        if (provinceData == null || provinceData.isEmpty()) {
            setCellValue(sheet, row, 3, 0.0);  // Lũy kế h
            setCellValue(sheet, row, 4, 0.0);  // Tổng SC lũy kế
            setCellValue(sheet, row, 5, 0.0);  // Tỷ lệ lũy kế
            setCellValueString(sheet, row, 6, "N/A"); // Cột G KPI
            return;
        }

        double cumulativeDaXuLy3hVip = 0.0;
        double cumulativeTongSc = 0.0;

        for (Map<String, Object> data : provinceData) {
            cumulativeDaXuLy3hVip += getNumberValue(data, "da_xu_ly_3hVip") != null
                    ? getNumberValue(data, "da_xu_ly_3hVip") : 0.0;
            cumulativeTongSc += getNumberValue(data, "tong_sc_da_xu_ly") != null
                    ? getNumberValue(data, "tong_sc_da_xu_ly") : 0.0;
        }

        double cumulativePercent = cumulativeTongSc > 0
                ? (cumulativeDaXuLy3hVip / cumulativeTongSc)
                : 0.0;

        // Fill D, E, F
        setCellValue(sheet, row, 3, cumulativeDaXuLy3hVip);
        setCellValue(sheet, row, 4, cumulativeTongSc);
        setCellValue(sheet, row, 5, cumulativePercent);

        // Fill G (KPI đạt hay không)
        if (progress3hVip != null) {
            String kpiStatus = cumulativePercent >= progress3hVip ? "Đạt" : "Không đạt";
            setCellValueString(sheet, row, 6, kpiStatus);
        } else {
            setCellValueString(sheet, row, 6, "N/A");
        }
    }
    private void fillHandleRateTotalRowTable8(Sheet sheet, Map<String, List<Map<String, Object>>> dataByProvince, int currentDay, Double progress3h) {
        List<Map<String, Object>> totalData = dataByProvince.get(TOTAL);

        // Initialize TOTAL row with default values for past days, null for future days
        for (int day = 1; day <= 31; day++) {
            int colIndex = TABLE8_DATA_START_COL - 4 + (day - 1) * 3;
            if (day <= currentDay) {
                setCellValue(sheet, TABLE8_TOTAL_ROW, colIndex, 0.0);     // da_xu_ly_3h
                setCellValue(sheet, TABLE8_TOTAL_ROW, colIndex + 1, 0.0); // tong_sc_da_xu_ly
                setCellValue(sheet, TABLE8_TOTAL_ROW, colIndex + 2, 0.0); // ty_le_3h
            } else {
                setCellValue(sheet, TABLE8_TOTAL_ROW, colIndex, null);
                setCellValue(sheet, TABLE8_TOTAL_ROW, colIndex + 1, null);
                setCellValue(sheet, TABLE8_TOTAL_ROW, colIndex + 2, null);
            }
        }

        double cumulativeDaXuLy48h = 0.0;
        double cumulativeTongSc = 0.0;

        if (totalData != null && !totalData.isEmpty()) {
            for (Map<String, Object> data : totalData) {
                cumulativeDaXuLy48h += getNumberValue(data, "da_xu_ly_3hVip") != null ? getNumberValue(data, "da_xu_ly_3hVip") : 0.0;
                cumulativeTongSc += getNumberValue(data, "tong_sc_da_xu_ly") != null ? getNumberValue(data, "tong_sc_da_xu_ly") : 0.0;
            }
            for (Map<String, Object> data : totalData) {
                int day = extractDayFromDate(data.get("ngay"));
                if (day > 0 && day <= 31) {
                    int colIndex = TABLE7_DATA_START_COL + (day - 1) * 3;
                    setCellValue(sheet, TABLE8_TOTAL_ROW, colIndex, getNumberValue(data, "da_xu_ly_3hVip"));
                    setCellValue(sheet, TABLE8_TOTAL_ROW, colIndex + 1, getNumberValue(data, "tong_sc_da_xu_ly"));
                    Double tyLe = getNumberValue(data, "ty_le_3hVip");
                    if (tyLe != null) setCellValue(sheet, TABLE8_TOTAL_ROW, colIndex + 2, tyLe / 100);
                }
            }
        }

        // Fill cumulative columns D, E, F
        setCellValue(sheet, TABLE8_TOTAL_ROW, 3, cumulativeDaXuLy48h);
        setCellValue(sheet, TABLE8_TOTAL_ROW, 4, cumulativeTongSc);
        double cumulativePercent = cumulativeTongSc > 0 ? (cumulativeDaXuLy48h / cumulativeTongSc) : 0.0;
        setCellValue(sheet, TABLE8_TOTAL_ROW, 5, cumulativePercent);

        // Fill KPI vào cột G
        if (progress3h != null) {
            String kpiStatus = cumulativePercent >= progress3h ? "Đạt" : "Không đạt";
            setCellValueString(sheet, TABLE8_TOTAL_ROW, 6, kpiStatus);
        } else {
            setCellValueString(sheet, TABLE8_TOTAL_ROW, 6, "N/A");
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
    private void setCellValueString(Sheet sheet, int rowIndex, int colIndex, String value) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) row = sheet.createRow(rowIndex);

        Cell cell = row.getCell(colIndex);
        if (cell == null) cell = row.createCell(colIndex);

        if (value == null) {
            cell.setBlank();
        } else {
            cell.setCellValue(value);
        }
    }

    private void setCellValueTable5(Sheet sheet, String value) {
        Row r = sheet.getRow(TransmissionChannelIncidentService.TABLE5_KPI_ROW);
        if (r == null) r = sheet.createRow(TransmissionChannelIncidentService.TABLE5_KPI_ROW);

        Cell cell = r.getCell(TransmissionChannelIncidentService.TABLE5_KPI_COL);
        if (cell == null) cell = r.createCell(TransmissionChannelIncidentService.TABLE5_KPI_COL);

        cell.setCellValue(value != null ? value : "");
    }
    private void setCellValueTable6(Sheet sheet, String value) {
        Row r = sheet.getRow(TransmissionChannelIncidentService.TABLE6_KPI_ROW);
        if (r == null) r = sheet.createRow(TransmissionChannelIncidentService.TABLE6_KPI_ROW);

        Cell cell = r.getCell(TransmissionChannelIncidentService.TABLE5_KPI_COL);
        if (cell == null) cell = r.createCell(TransmissionChannelIncidentService.TABLE5_KPI_COL);

        cell.setCellValue(value != null ? value : "");
    }
    private void setCellValueTable7(Sheet sheet, String value) {
        Row r = sheet.getRow(TransmissionChannelIncidentService.TABLE7_KPI_ROW);
        if (r == null) r = sheet.createRow(TransmissionChannelIncidentService.TABLE7_KPI_ROW);

        Cell cell = r.getCell(TransmissionChannelIncidentService.TABLE5_KPI_COL);
        if (cell == null) cell = r.createCell(TransmissionChannelIncidentService.TABLE5_KPI_COL);

        cell.setCellValue(value != null ? value : "");
    }
    private void setCellValueTable8(Sheet sheet, String value) {
        Row r = sheet.getRow(TransmissionChannelIncidentService.TABLE8_KPI_ROW);
        if (r == null) r = sheet.createRow(TransmissionChannelIncidentService.TABLE8_KPI_ROW);

        Cell cell = r.getCell(TransmissionChannelIncidentService.TABLE5_KPI_COL);
        if (cell == null) cell = r.createCell(TransmissionChannelIncidentService.TABLE5_KPI_COL);

        cell.setCellValue(value != null ? value : "");
    }

    private Double getNumberValue(Map<String, Object> data, String key) {
        Object value = data.get(key);
        return value instanceof Number ? ((Number) value).doubleValue() : null;
    }

    public BufferedImage exportChartToImage(String inputPath, String sheetName) throws Exception {
        Workbook wbInput = new Workbook();
        wbInput.loadFromFile(inputPath);
        Worksheet sheetInput = wbInput.getWorksheets().get(sheetName);

        Chart chart = null;
        for (int i = 0; i < sheetInput.getCharts().size(); i++) {
            Chart c = sheetInput.getCharts().get(i);
            // Ví dụ tìm chart trong vùng F10:Q25
            if (c.getLeftColumn() >= 6 && c.getRightColumn() <= 17
                    && c.getTopRow() >= 10 && c.getBottomRow() <= 25) {
                chart = c;
                break;
            }
        }
        if (chart == null) throw new RuntimeException("Không tìm thấy chart trong block F10:Q25");

        return chart.saveToImage();
    }

    public void insertChartImage(XSSFWorkbook workbook, String sheetName, BufferedImage chartImage) throws Exception {
        XSSFSheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) throw new RuntimeException("Không tìm thấy sheet " + sheetName);

        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        ImageIO.write(chartImage, "png", baos);
        int pictureIdx = workbook.addPicture(baos.toByteArray(), org.apache.poi.ss.usermodel.Workbook.PICTURE_TYPE_PNG);

        CreationHelper helper = workbook.getCreationHelper();
        Drawing<?> drawing = sheet.createDrawingPatriarch();
        ClientAnchor anchor = helper.createClientAnchor();
        anchor.setCol1(12); // M
        anchor.setRow1(13); // 14

        drawing.createPicture(anchor, pictureIdx).resize();
    }

    private String extractResourceToTemp(String resourcePath) throws Exception {
        try (var in = getClass().getResourceAsStream(resourcePath)) {
            if (in == null) throw new RuntimeException("Không tìm thấy resource: " + resourcePath);
            var tempFile = java.io.File.createTempFile("excel_template_", ".xlsx");
            tempFile.deleteOnExit();
            try (var out = new java.io.FileOutputStream(tempFile)) {
                in.transferTo(out);
            }
            return tempFile.getAbsolutePath();
        }
    }

    private void fillSheet2KpiData(Sheet sheet2, ExcelDTO data) {
        if (sheet2 == null) {
            log.warn("Sheet 2 (KPI VTS) not found, skipping KPI data population");
            return;
        }

        int currentDay = java.time.LocalDate.now().getDayOfMonth();

        // Get KPI progress data
        List<Map<String, Object>> progressKpi = data.getProgressKpi();
        Double progress3h = null;
        Double progress24h = null;
        Double progress48h = null;
        Double satisfyLevel = null;

        if (progressKpi != null && !progressKpi.isEmpty()) {
            Map<String, Object> kpiData = progressKpi.get(0);
            progress3h = getNumberValue(kpiData, "progress_3h");
            progress24h = getNumberValue(kpiData, "progress_24h");
            progress48h = getNumberValue(kpiData, "progress_48h");
            satisfyLevel = getNumberValue(kpiData, "satisfyLevel");

            // Convert from percentage to decimal if needed
            if (progress3h != null) progress3h = progress3h / 100;
            if (progress24h != null) progress24h = progress24h / 100;
            if (progress48h != null) progress48h = progress48h / 100;
            if (satisfyLevel != null) satisfyLevel = satisfyLevel / 100;
        }

        // Fill KPI 3H section (rows C5, C7, C8)
        fillSheet2KpiRow(sheet2, SHEET2_KPI_3H_ROW, progress3h, currentDay, "3H KPI");
        fillSheet2ActualDataRow(sheet2, SHEET2_3H_DA_XU_LI_ROW, data.getHandleRate3h(), "da_xu_ly_3h", currentDay, "3H Đã xử lý");
        fillSheet2ActualDataRow(sheet2, SHEET2_3H_TONG_SC_ROW, data.getHandleRate3h(), "tong_sc_da_xu_ly", currentDay, "3H Tổng SC");

        // Fill KPI 24H section (rows C11, C13, C14)
        fillSheet2KpiRow(sheet2, SHEET2_KPI_24H_ROW, progress24h, currentDay, "24H KPI");
        fillSheet2ActualDataRow(sheet2, SHEET2_24H_DA_XU_LI_ROW, data.getHandleRate24h(), "da_xu_ly_24h", currentDay, "24H Đã xử lý");
        fillSheet2ActualDataRow(sheet2, SHEET2_24H_TONG_SC_ROW, data.getHandleRate24h(), "tong_sc_da_xu_ly", currentDay, "24H Tổng SC");

        // Fill KPI 48H section (rows C17, C19, C20)
        fillSheet2KpiRow(sheet2, SHEET2_KPI_48H_ROW, progress48h, currentDay, "48H KPI");
        fillSheet2ActualDataRow(sheet2, SHEET2_48H_DA_XU_LI_ROW, data.getHandleRate48h(), "da_xu_ly_48h", currentDay, "48H Đã xử lý");
        fillSheet2ActualDataRow(sheet2, SHEET2_48H_TONG_SC_ROW, data.getHandleRate48h(), "tong_sc_da_xu_ly", currentDay, "48H Tổng SC");
        fillSheet2KpiRow(sheet2, SHEET2_KPI_SATISFY_ROW, satisfyLevel, currentDay, "Satisfaction Level KPI");

        log.debug("Filled Sheet 2 complete data - 3H KPI: {}, 24H KPI: {}, 48H KPI: {}", progress3h, progress24h, progress48h);
    }

    private void fillSheet2KpiRow(Sheet sheet, int rowIndex, Double kpiValue, int currentDay, String kpiType) {
        if (kpiValue == null) {
            log.warn("KPI value is null for {}, skipping row {}", kpiType, rowIndex + 1);
            return;
        }

        // Fill KPI value for each day from C to AK (columns 2 to 36, representing days 1-35)
        for (int day = 1; day <= 35; day++) {
            int colIndex = SHEET2_DATA_START_COL + (day - 1); // C=2, D=3, ..., AK=36

            // Only fill data for past and current days, leave future days empty
            if (day <= currentDay) {
                setCellValue(sheet, rowIndex, colIndex, kpiValue);
            } else {
                setCellValue(sheet, rowIndex, colIndex, null);
            }
        }

        log.debug("Filled {} row {} with value {} for {} days", kpiType, rowIndex + 1, kpiValue, currentDay);
    }

    private void fillSheet2ActualDataRow(Sheet sheet, int rowIndex, List<Map<String, Object>> handleRateData,
                                         String dataKey, int currentDay, String dataType) {
        if (handleRateData == null || handleRateData.isEmpty()) {
            log.warn("Handle rate data is null or empty for {}, skipping row {}", dataType, rowIndex + 1);
            return;
        }

        // Group data by province, focusing on TOTAL data for Sheet2
        Map<String, List<Map<String, Object>>> dataByProvince = handleRateData.stream()
                .collect(Collectors.groupingBy(
                        d -> d.get("tinh") != null ? d.get("tinh").toString().trim() : "unknown"
                ));

        List<Map<String, Object>> totalData = dataByProvince.get(TOTAL);

        // Initialize all days with 0 for past days, null for future days
        for (int day = 1; day <= 35; day++) {
            int colIndex = SHEET2_DATA_START_COL + 1 + (day - 1); // C=2, D=3, ..., AK=36

            if (day <= currentDay) {
                setCellValue(sheet, rowIndex, colIndex, 0.0);
            } else {
                setCellValue(sheet, rowIndex, colIndex, null);
            }
        }

        // Fill actual data if available
        if (totalData != null && !totalData.isEmpty()) {
            totalData.forEach(data -> {
                int day = extractDayFromDate(data.get("ngay"));
                if (day > 0 && day <= 35 && day <= currentDay) {
                    int colIndex = SHEET2_DATA_START_COL+1 + (day - 1);
                    Double value = getNumberValue(data, dataKey);
                    if (value != null) {
                        setCellValue(sheet, rowIndex, colIndex, value);
                    }
                }
            });
            log.debug("Filled {} row {} with {} records", dataType, rowIndex + 1, totalData.size());
        } else {
            log.debug("No TOTAL data found for {}, using default values", dataType);
        }
    }
    private void fillSatisfactionLevel(Sheet sheet, ExcelDTO excelData) {
        List<Map<String, Object>> satisfactionData = excelData.getSatisfactionLevel();
        if (satisfactionData == null) satisfactionData = List.of();

        Row numeratorRow = sheet.getRow(47);   // Row 48 in Excel
        Row denominatorRow = sheet.getRow(48); // Row 49 in Excel

        if (numeratorRow == null) numeratorRow = sheet.createRow(47);
        if (denominatorRow == null) denominatorRow = sheet.createRow(48);

        int startCol = 3; // Column D

        // Group data by date
        Map<String, Map<String, Object>> dataByDate = satisfactionData.stream()
                .collect(Collectors.toMap(
                        r -> r.get("Ngày").toString(),
                        r -> r
                ));

        LocalDate firstDayPrevMonth = LocalDate.now().minusMonths(1).withDayOfMonth(1);
        int daysInMonth = firstDayPrevMonth.lengthOfMonth();
        int currentDay = LocalDate.now().getDayOfMonth();

        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd");

        for (int day = 1; day <= daysInMonth; day++) {
            int colIndex = startCol + (day - 1);

            String dateStr = firstDayPrevMonth.withDayOfMonth(day).format(formatter);
            Map<String, Object> record = dataByDate.get(dateStr);

            Double numerator = null;
            Double denominator = null;

            if (record != null) {
                Object numObj = record.get("Tử số - KH hài lòng");
                Object denObj = record.get("Mẫu số - Tổng KH có phản hồi");

                if (numObj instanceof Number) numerator = ((Number) numObj).doubleValue();
                if (denObj instanceof Number) denominator = ((Number) denObj).doubleValue();
            }

            // Nếu ngày <= ngày hiện tại thì fill dữ liệu, ngược lại để trống
            if (day <= currentDay) {
                Cell numeratorCell = numeratorRow.getCell(colIndex);
                if (numeratorCell == null) numeratorCell = numeratorRow.createCell(colIndex);
                numeratorCell.setCellValue(numerator != null ? numerator : 0);

                Cell denominatorCell = denominatorRow.getCell(colIndex);
                if (denominatorCell == null) denominatorCell = denominatorRow.createCell(colIndex);
                denominatorCell.setCellValue(denominator != null ? denominator : 0);
            } else {
                Cell numeratorCell = numeratorRow.getCell(colIndex);
                if (numeratorCell == null) numeratorCell = numeratorRow.createCell(colIndex);
                numeratorCell.setBlank();

                Cell denominatorCell = denominatorRow.getCell(colIndex);
                if (denominatorCell == null) denominatorCell = denominatorRow.createCell(colIndex);
                denominatorCell.setBlank();
            }
        }
    }

    private void fillAvgHandleTime(Sheet sheet, ExcelDTO excelData) {
        List<Map<String, Object>> avgHandleTimeData = excelData.getAvgTimeHandle();
        if (avgHandleTimeData == null) avgHandleTimeData = List.of();

        // Rows 32 and 33 (POI index starts at 0 -> row 31 và 32)
        Row totalTimeRow = sheet.getRow(31);   // D32
        Row totalCountRow = sheet.getRow(32);  // D33

        if (totalTimeRow == null) totalTimeRow = sheet.createRow(31);
        if (totalCountRow == null) totalCountRow = sheet.createRow(32);

        int startCol = 3; // D là cột số 3 (0-based index)

        // Map dữ liệu theo ngày để lookup
        Map<String, Map<String, Object>> dataByDate = avgHandleTimeData.stream()
                .collect(Collectors.toMap(
                        r -> r.get("Ngày").toString(),
                        r -> r
                ));

        LocalDate firstDayPrevMonth = LocalDate.now().minusMonths(1).withDayOfMonth(1);
        int daysInMonth = firstDayPrevMonth.lengthOfMonth();

        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd");

        for (int day = 1; day <= daysInMonth; day++) {
            int colIndex = startCol + (day - 1);

            String dateStr = firstDayPrevMonth.withDayOfMonth(day).format(formatter);
            Map<String, Object> record = dataByDate.get(dateStr);

            double totalHours = 0;
            double totalCount = 0;

            if (record != null) {
                Object hoursObj = record.get("Tổng thời gian xử lý (giờ)");
                Object countObj = record.get("Số lượng phản ánh");

                if (hoursObj instanceof Number) totalHours = ((Number) hoursObj).doubleValue();
                if (countObj instanceof Number) totalCount = ((Number) countObj).doubleValue();
            }

            // Fill total processing hours
            Cell totalHoursCell = totalTimeRow.getCell(colIndex);
            if (totalHoursCell == null) totalHoursCell = totalTimeRow.createCell(colIndex);
            totalHoursCell.setCellValue(totalHours);

            // Fill total count
            Cell totalCountCell = totalCountRow.getCell(colIndex);
            if (totalCountCell == null) totalCountCell = totalCountRow.createCell(colIndex);
            totalCountCell.setCellValue(totalCount);
        }
    }

}