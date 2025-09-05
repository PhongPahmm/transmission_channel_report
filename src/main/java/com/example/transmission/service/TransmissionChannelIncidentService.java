package com.example.transmission.service;

import com.example.transmission.domain.ServiceSubscriber;
import com.example.transmission.dto.ExcelDTO;
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
            fillTable3(sheet, excelData); // Add Table 3 processing

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
        if (value == null) return;

        Row row = sheet.getRow(rowIndex);
        if (row == null) row = sheet.createRow(rowIndex);

        Cell cell = row.getCell(colIndex);
        if (cell == null) cell = row.createCell(colIndex);

        cell.setCellValue(value.doubleValue());
    }

    private Double getNumberValue(Map<String, Object> data, String key) {
        Object value = data.get(key);
        return value instanceof Number ? ((Number) value).doubleValue() : null;
    }
}