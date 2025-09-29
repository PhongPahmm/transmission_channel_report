package com.example.transmission.service;

import com.example.transmission.dto.TransmissionChanelDTO;
import com.example.transmission.repository.TransmissionChannelIncidentRepository;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.*;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.stream.Collectors;

@Slf4j
@Service
@RequiredArgsConstructor
public class TransmissionChannelIncidentService  {

    private final TransmissionChannelIncidentRepository transmissionChannelIncidentRepository;

    // Constants for Excel structure
    private static final String TEMPLATE_PATH = "/templates/template_output.xlsx";
    private static final String SHEET_NAME = "PL06-Kenh truyen";
    private static final String SHEET_NAME_2 = "KPI VTS";

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
    private static final int TABLE3_SC_ROW = 32;        // Row 33 (label: Sự cố)
    private static final int TABLE3_START_ROW = 33;     // Row 34 (Office WAN)
    private static final int TABLE3_END_ROW = 36;       // Row 37 (Kênh truyền quốc tế)
    private static final int TABLE3_HT_ROW = 37;        // Row 38 (label: HT)
    private static final int TABLE3_DATA_START_COL = 4; // Column E = index 4
    private static final int TABLE3_AVG_COL = 3;
    private static final int TABLE3_TOTAL_ROW = 31;


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

    public byte[] exportExcelFile(LocalDate currentDate) {
        try (InputStream inputStream = getClass().getResourceAsStream(TEMPLATE_PATH)) {
            assert inputStream != null;
            try (XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
                 ByteArrayOutputStream out = new ByteArrayOutputStream()) {

                Sheet sheet = workbook.getSheet(SHEET_NAME);
                Sheet sheet2 = workbook.getSheet(SHEET_NAME_2);

                // Get all required data
                TransmissionChanelDTO excelData = fetchExcelData(currentDate);

                // Fill tables
                fillTable1(sheet, excelData, currentDate);
                fillTable2(sheet, excelData, currentDate);
                fillTable3(sheet, excelData, currentDate);
                fillTable4(sheet, excelData, currentDate);
                fillTable5(sheet, excelData, currentDate);
                fillTable6(sheet, excelData, currentDate);
                fillTable7(sheet, excelData, currentDate);
                fillTable8(sheet, excelData, currentDate);
                fillSheet2KpiData(sheet2, excelData, currentDate);
                fillSatisfactionLevel(sheet2, excelData, currentDate);
                fillAvgHandleTime(sheet2, excelData, currentDate);
                fillMonthlySummaryData(sheet, excelData, currentDate);

                workbook.setForceFormulaRecalculation(true);

                workbook.write(out);
                return out.toByteArray();

            }
        } catch (IOException e) {
            log.error("Error while exporting Excel", e);
            throw new RuntimeException("Error while exporting Excel", e);
        } catch (Exception e) {
            log.error("Unexpected error while exporting Excel", e);
            throw new RuntimeException("lỗi ", e);
        }
    }

    private TransmissionChanelDTO fetchExcelData(LocalDate currentDate) {
        return new TransmissionChanelDTO(
                transmissionChannelIncidentRepository.getCategorySummary(currentDate),
                transmissionChannelIncidentRepository.getOnTimeHandle(currentDate),
                transmissionChannelIncidentRepository.getDailyStats(currentDate),
                transmissionChannelIncidentRepository.getComplaintRateAndTotalSubscribers(currentDate),
                transmissionChannelIncidentRepository.getHandleRate3h(currentDate),
                transmissionChannelIncidentRepository.getHandleRate24h(currentDate),
                transmissionChannelIncidentRepository.getHandleRate48h(currentDate),
                transmissionChannelIncidentRepository.getHandleRate3hVip(currentDate),
                transmissionChannelIncidentRepository.getLevelSatisfy(currentDate),
                transmissionChannelIncidentRepository.getAvgHandleTime(currentDate),
                null,
                transmissionChannelIncidentRepository.getLast8daysAndAvgMonth(currentDate),
                transmissionChannelIncidentRepository.getIncidentMonthlySummary(currentDate.getYear())
        );
    }

    private void fillTable1(Sheet sheet, TransmissionChanelDTO data, LocalDate currentDate) {
        List<Map<String, Object>> categoryResults = data.getCategorySummaryResults();
        if (categoryResults == null || categoryResults.isEmpty()) return;

        // Lấy số ngày trong tháng từ currentDate
        int daysInMonth = currentDate.lengthOfMonth();

        // Nhóm theo incidentName để tính tổng totalPerDay
        Map<String, Integer> sumByIncident = categoryResults.stream()
                .collect(Collectors.groupingBy(
                        r -> Objects.toString(r.get("incidentName"), ""),
                        Collectors.summingInt(r -> ((Number) r.get("totalPerDay")).intValue())
                ));

        // Tính avgDay = tổng / daysInMonth
        Map<String, Double> avgByIncident = sumByIncident.entrySet().stream()
                .collect(Collectors.toMap(
                        Map.Entry::getKey,
                        e -> e.getValue() / (double) daysInMonth
                ));

        // Lọc dữ liệu cho ngày hiện tại
        List<Map<String, Object>> filteredResults = categoryResults.stream()
                .filter(Objects::nonNull)
                .filter(record -> {
                    LocalDate recordDate = extractLocalDate(record.get("day"));
                    return recordDate != null && recordDate.equals(currentDate);
                }).collect(Collectors.toList());

        if (filteredResults.isEmpty()) {
            System.out.println("Không có dữ liệu cho ngày: " + currentDate);
            return;
        }

        // Fill ngày lên tiêu đề
        Row titleRow = sheet.getRow(4);
        if (titleRow == null) titleRow = sheet.createRow(4);
        Cell titleCell = titleRow.getCell(3);
        if (titleCell == null) titleCell = titleRow.createCell(3);
        titleCell.setCellValue(currentDate.format(DateTimeFormatter.ofPattern("dd/MM/yyyy")));

        int rowIndex = TABLE1_START_ROW;

        for (Map<String, Object> rowData : filteredResults) {
            if (rowIndex > TABLE1_END_ROW) break;

            Row row = sheet.getRow(rowIndex);
            if (row == null) row = sheet.createRow(rowIndex);

            String incidentName = Objects.toString(rowData.get("incidentName"), "");

            // Category = incidentName
            setCellValue(row, CATEGORY_COL, incidentName);

            // KPI = kpiChild
            setCellValue(row, KPI_COL, rowData.get("kpiChild"));

            // Subscribers = totalSubcriberChild
            setCellValue(row, SUBSCRIBERS_COL, rowData.get("totalSubcriberChild"));

            // Value = totalPerDay
            setCellValue(row, VALUE_COL, rowData.get("totalPerDay"));

            // Comp day = avgDay tính từ code
            Double avgDay = avgByIncident.getOrDefault(incidentName, 0.0);
            setCellValue(row, COMP_DAY_COL, avgDay);

            rowIndex++;
        }

        // Fill parent KPI row
        Row parentRow = sheet.getRow(PARENT_KPI_ROW);
        if (parentRow == null) parentRow = sheet.createRow(PARENT_KPI_ROW);

        Optional<Map<String, Object>> firstRowOpt = filteredResults.stream().findFirst();
        if (firstRowOpt.isPresent()) {
            Map<String, Object> firstRow = firstRowOpt.get();
            setCellValue(parentRow, KPI_COL, firstRow.get("kpiParent"));
            setCellValue(parentRow, SUBSCRIBERS_COL, firstRow.get("totalSubcriberParent"));
        }
    }


    private void fillTable2(Sheet sheet, TransmissionChanelDTO data, LocalDate currentDate) {
        List<Map<String, Object>> raw = data.getOnTimeHandle();
        if (raw == null || raw.isEmpty()) return;
        List<Map<String, Object>> onTimeHandle = calculateCumulativeTable2(raw);
        Row titleRow = sheet.getRow(4);
        if (titleRow == null) {
            titleRow = sheet.createRow(4);
        }
        Cell titleCell = titleRow.getCell(18);
        if (titleCell == null) {
            titleCell = titleRow.createCell(18);
        }
        titleCell.setCellValue(currentDate.format(DateTimeFormatter.ofPattern("dd/MM/yyyy")));

        List<Map<String, Object>> filteredRecords = onTimeHandle.stream()
                .filter(Objects::nonNull)
                .filter(record -> {
                    LocalDate recordDate = extractLocalDate(record.get("day"));
                    return recordDate != null && recordDate.equals(currentDate);
                }).collect(Collectors.toList());


        if (filteredRecords.isEmpty()) {
            System.out.println(" Không có dữ liệu cho ngày: " + currentDate);
            return;
        }

        for (Map<String, Object> record : filteredRecords) {
            String incident = (String) record.get("incidentName");

            int rowOffset = -1;
            if ("Tỉ lệ sự cố xử lý trong 3h".equals(incident)) {
                rowOffset = 0;
            } else if ("Tỉ lệ sự cố xử lý trong 24h".equals(incident)) {
                rowOffset = 1;
            } else if ("Tỉ lệ sự cố xử lý trong 48h".equals(incident)) {
                rowOffset = 2;
            }

            if (rowOffset < 0) {
                System.out.println("Incident name không khớp: " + incident);
                continue;
            }

            // --- KPI Progress ---
            Double kpiProgress = getNumberValue(record, "kpi_progress");
            if (kpiProgress != null) {
                setCellValue(sheet, TABLE2_START_ROW + rowOffset, PROGRESS_KPI_COL, kpiProgress / 100);
            }

            // --- xử lý ---
            Double handle = getNumberValue(record, "xu_li_trong_han_3h");
            if (handle != null) {
                setCellValue(sheet, TABLE2_START_ROW + rowOffset, ON_TIME_HANDLE_COL, handle);
            }

            // --- Tổng xử lý ---
            Double tongXuLy = getNumberValue(record, "tong_xu_ly");
            if (tongXuLy != null) {
                setCellValue(sheet, TABLE2_START_ROW + rowOffset, TOTAL_HANDLE_COL, tongXuLy);
            }

            // --- Kết quả ---
            if (tongXuLy != null && tongXuLy > 0) {
                Double xuLyTrongHan = null;
                if (incident.contains("3h")) {
                    xuLyTrongHan = getNumberValue(record, "ty_le_3h_luy_ke");
                } else if (incident.contains("24h")) {
                    xuLyTrongHan = getNumberValue(record, "ty_le_24h_luy_ke");
                } else if (incident.contains("48h")) {
                    xuLyTrongHan = getNumberValue(record, "ty_le_48h_luy_ke");
                }
                if (xuLyTrongHan != null) {
                    setCellValue(sheet, TABLE2_START_ROW + rowOffset, RESULT_COL, xuLyTrongHan / 100);
                }
            }
        }
    }
    private List<Map<String, Object>> calculateCumulativeTable2(List<Map<String, Object>> rawData) {
        // Gom nhóm theo parentName + incidentName
        Map<String, List<Map<String, Object>>> grouped = rawData.stream()
                .collect(Collectors.groupingBy(r -> r.get("parentName") + "_" + r.get("incidentName")));

        List<Map<String, Object>> result = new ArrayList<>();

        for (List<Map<String, Object>> group : grouped.values()) {
            // Sắp xếp theo ngày
            group.sort(Comparator.comparing(r -> extractLocalDate(r.get("day"))));

            int cum3h = 0, cum24h = 0, cum48h = 0, cumTotal = 0;

            for (Map<String, Object> record : group) {
                int xuLy3h = getInt(record, "xu_li_trong_han_3h");
                int xuLy24h = getInt(record, "xu_li_trong_han_24h");
                int xuLy48h = getInt(record, "xu_li_trong_han_48h");
                int tong = getInt(record, "tong_xu_ly");

                cum3h += xuLy3h;
                cum24h += xuLy24h;
                cum48h += xuLy48h;
                cumTotal += tong;

                record.put("tong_3h_luy_ke", cum3h);
                record.put("tong_24h_luy_ke", cum24h);
                record.put("tong_48h_luy_ke", cum48h);
                record.put("tong_xu_ly_luy_ke", cumTotal);

                record.put("ty_le_3h_luy_ke", cumTotal > 0 ? round2((cum3h * 100.0) / cumTotal) : 0.0);
                record.put("ty_le_24h_luy_ke", cumTotal > 0 ? round2((cum24h * 100.0) / cumTotal) : 0.0);
                record.put("ty_le_48h_luy_ke", cumTotal > 0 ? round2((cum48h * 100.0) / cumTotal) : 0.0);

                result.add(record);
            }
        }
        return result;
    }

    private int getInt(Map<String, Object> map, String key) {
        Object val = map.get(key);
        if (val instanceof Number) {
            return ((Number) val).intValue();
        }
        return 0;
    }

    private double round2(double value) {
        return Math.round(value * 100.0) / 100.0;
    }


    private void fillTable3(Sheet sheet, TransmissionChanelDTO data, LocalDate currentDate) {
        fillG41Date(sheet, currentDate);
        // Fill header row 31 (C, D, E)
        fillTable3Header(sheet, currentDate);
        
        fillDailyStats(sheet, data.getDailyStats(), currentDate);
    }

    private void fillTable4(Sheet sheet, TransmissionChanelDTO data, LocalDate currentDate) {
        // Fill header row 42 (G)
        fillTable4Header(sheet, currentDate);
        
        // Fill date into G41

        
        fillComplaintRateByProvince(sheet, data.getComplaintRate(), currentDate);
    }

    private void fillTable5(Sheet sheet, TransmissionChanelDTO data, LocalDate currentDate) {
        // Fill header row 82 (H)
        fillTable5Header(sheet, currentDate);
        
        List<Map<String, Object>> handleRateData = data.getHandleRate3h();
        Double progress3h = null;

        if (data.getOnTimeHandle() != null && !data.getOnTimeHandle().isEmpty()) {
            Optional<Map<String, Object>> scKpi = data.getOnTimeHandle().stream()
                    .filter(row -> {
                        Object parent = row.get("incidentName");
                        return parent != null && parent.toString().equals("Tỉ lệ sự cố xử lý trong 3h");
                    })
                    .findFirst();

            if (scKpi.isPresent()) {
                Double val = getNumberValue(scKpi.get(), "kpi_progress");
                if (val != null) {
                    progress3h = val;
                }
            }
        }

        fillHandleRate3hByProvince(sheet, handleRateData, progress3h, currentDate);
        fillTable5Kpi(sheet, progress3h);
    }

    private void fillTable6(Sheet sheet, TransmissionChanelDTO data, LocalDate currentDate) {
        // Fill header row 122 (H)
        fillTable6Header(sheet, currentDate);
        
        List<Map<String, Object>> handleRateData = data.getHandleRate24h();
        Double progress24h = null;
        if (data.getOnTimeHandle() != null && !data.getOnTimeHandle().isEmpty()) {
            Optional<Map<String, Object>> scKpi = data.getOnTimeHandle().stream()
                    .filter(row -> {
                        Object parent = row.get("incidentName");
                        return parent != null && parent.toString().equals("Tỉ lệ sự cố xử lý trong 24h");
                    })
                    .findFirst();

            if (scKpi.isPresent()) {
                Double val = getNumberValue(scKpi.get(), "kpi_progress");
                if (val != null) {
                    progress24h = val;
                }
            }
        }
        fillHandleRate24hByProvince(sheet, handleRateData, progress24h, currentDate);
        fillTable6Kpi(sheet, progress24h);

    }

    private void fillTable7(Sheet sheet, TransmissionChanelDTO data, LocalDate currentDate) {
        // Fill header row 162 (H)
        fillTable7Header(sheet, currentDate);
        
        List<Map<String, Object>> handleRateData = data.getHandleRate48h();
        Double progress48h = null;
        if (data.getOnTimeHandle() != null && !data.getOnTimeHandle().isEmpty()) {
            Optional<Map<String, Object>> scKpi = data.getOnTimeHandle().stream()
                    .filter(row -> {
                        Object parent = row.get("incidentName");
                        return parent != null && parent.toString().equals("Tỉ lệ sự cố xử lý trong 48h");
                    })
                    .findFirst();

            if (scKpi.isPresent()) {
                Double val = getNumberValue(scKpi.get(), "kpi_progress");
                if (val != null) {
                    progress48h = val;
                }
            }
        }
        fillHandleRate48hByProvince(sheet, handleRateData, progress48h, currentDate);
        fillTable7Kpi(sheet, progress48h);
    }

    private void fillTable8(Sheet sheet, TransmissionChanelDTO data, LocalDate currentDate) {
        // Fill header row 202 (H)
        fillTable8Header(sheet, currentDate);
        
        List<Map<String, Object>> handleRateData = data.getHandleRate3hVip();
        Double progress3hVip = null;
        if (data.getOnTimeHandle() != null && !data.getOnTimeHandle().isEmpty()) {
            Optional<Map<String, Object>> scKpi = data.getOnTimeHandle().stream()
                    .filter(row -> {
                        Object parent = row.get("incidentName");
                        return parent != null && parent.toString().equals("Tỉ lệ sự cố xử lý trong 3h Vip");
                    })
                    .findFirst();

            if (scKpi.isPresent()) {
                Double val = getNumberValue(scKpi.get(), "kpi_progress");
                if (val != null) {
                    progress3hVip = val;
                }
            }
        }
        fillHandleRate3hVipByProvince(sheet, handleRateData, progress3hVip, currentDate);
        fillTable8Kpi(sheet, progress3hVip);
    }

    private void fillDailyStats(Sheet sheet, List<Map<String, Object>> dailyData, LocalDate currentDate) {
        if (dailyData == null || dailyData.isEmpty()) return;

        int maxDay = currentDate.getDayOfMonth();

        // Remove null records up-front for safety
        List<Map<String, Object>> safeDailyData = dailyData.stream()
                .filter(Objects::nonNull)
                .collect(Collectors.toList());

        // --- 1. Fill chi tiết theo incidentName (rows 34–37) ---
        for (int serviceRow = TABLE3_START_ROW; serviceRow <= TABLE3_END_ROW; serviceRow++) {
            String serviceName = getCellStringValue(sheet, serviceRow, CATEGORY_COL); // Cột B
            if (serviceName == null) continue;

            final int currentServiceRow = serviceRow;

            List<Map<String, Object>> serviceData = safeDailyData.stream()
                    .filter(d -> {
                        Object cate = d.get("incidentName");
                        Object parent = d.get("parentName");
                        return cate != null && cate.toString().trim().contains(serviceName) &&
                               parent != null && "SC".equals(parent.toString().trim());
                    })
                    .collect(Collectors.toList());

            // Reset cells: < currentDate = 0, > currentDate = null
            for (int day = 1; day <= 31; day++) {
                int colIndex = TABLE3_DATA_START_COL + (day - 1);
                if (day <= maxDay) {
                    setCellValue(sheet, currentServiceRow, colIndex, 0.0);
                } else {
                    setCellValue(sheet, currentServiceRow, colIndex, null);
                }
            }

            // Fill dữ liệu cho incidentName
            serviceData.forEach(data -> {
                Object ngayObj = data.get("day");
                Object countObj = data.get("numberOfSC");

                LocalDate date = extractLocalDate(ngayObj);
                if (date != null &&
                        date.getMonthValue() == currentDate.getMonthValue() &&
                        date.getYear() == currentDate.getYear() &&
                        countObj instanceof Number) {
                    int day = date.getDayOfMonth();
                    if (day > 0 && day <= maxDay) {
                        int colIndex = TABLE3_DATA_START_COL + (day - 1);
                        setCellValue(sheet, currentServiceRow, colIndex, ((Number) countObj).doubleValue());
                    }
                }
            });
            // --- Fill AVG cho từng incidentName ---
            Optional<Map<String, Object>> firstValid = serviceData.stream().filter(Objects::nonNull).findFirst();
            if (firstValid.isPresent()) {
                Object avgIncident = firstValid.get().get("avg_per_incident");
                if (avgIncident instanceof Number) {
                    setCellValue(sheet, currentServiceRow, TABLE3_AVG_COL, ((Number) avgIncident).doubleValue());
                }
            }
        }

        // --- 2. Fill tổng SC (row 33) & HT (row 38) ---
        for (int day = 1; day <= 31; day++) {
            int colIndex = TABLE3_DATA_START_COL + (day - 1);
            if (day <= maxDay) {
                setCellValue(sheet, TABLE3_SC_ROW, colIndex, 0.0);
                setCellValue(sheet, TABLE3_HT_ROW, colIndex, 0.0);
            } else {
                setCellValue(sheet, TABLE3_SC_ROW, colIndex, null);
                setCellValue(sheet, TABLE3_HT_ROW, colIndex, null);
            }
        }

        safeDailyData.forEach(data -> {
            Object ngayObj = data.get("day");
            Object scObj = data.get("total_SC");
            Object htObj = data.get("total_HT");

            LocalDate date = extractLocalDate(ngayObj);
            if (date != null &&
                    date.getMonthValue() == currentDate.getMonthValue() &&
                    date.getYear() == currentDate.getYear()) {
                int day = date.getDayOfMonth();
                if (day > 0 && day <= maxDay) {
                    int colIndex = TABLE3_DATA_START_COL + (day - 1);

                    if (scObj instanceof Number) {
                        setCellValue(sheet, TABLE3_SC_ROW, colIndex, ((Number) scObj).doubleValue());
                    }
                    if (htObj instanceof Number) {
                        setCellValue(sheet, TABLE3_HT_ROW, colIndex, ((Number) htObj).doubleValue());
                    }
                }
            }
        });
        // --- 3. Fill AVG SC, HT, TotalAll ---
        Optional<Map<String, Object>> firstRowOpt = safeDailyData.stream().filter(Objects::nonNull).findFirst();
        if (firstRowOpt.isPresent()) {
            Map<String, Object> firstRow = firstRowOpt.get();
            Object avgSC = firstRow.get("avg_SC");
            Object avgHT = firstRow.get("avg_HT");
            Object avgTotalAll = firstRow.get("avg_totalAll");

            if (avgSC instanceof Number) {
                setCellValue(sheet, TABLE3_SC_ROW, TABLE3_AVG_COL, ((Number) avgSC).doubleValue());
            }
            if (avgHT instanceof Number) {
                setCellValue(sheet, TABLE3_HT_ROW, TABLE3_AVG_COL, ((Number) avgHT).doubleValue());
            }
            if (avgTotalAll instanceof Number) {
                setCellValue(sheet, TABLE3_TOTAL_ROW, TABLE3_AVG_COL, ((Number) avgTotalAll).doubleValue());
            }
        }
    }

    private void fillComplaintRateByProvince(Sheet sheet, List<Map<String, Object>> complaintData, LocalDate currentDate) {
        if (complaintData == null || complaintData.isEmpty()) return;

        int currentDay = currentDate.getDayOfMonth();

        // Group data by province
        Map<String, List<Map<String, Object>>> dataByProvince = complaintData.stream()
                .collect(Collectors.groupingBy(
                        data -> data.get("province") != null ? data.get("province").toString().trim() : "Unknown"
                ));

        for (int provinceRow = TABLE4_START_ROW; provinceRow <= TABLE4_END_ROW; provinceRow++) {
            String provinceName = getCellStringValue(sheet, provinceRow, CATEGORY_COL);
            if (provinceName == null) continue;

            // Lấy data của tỉnh
            List<Map<String, Object>> provinceData = dataByProvince.entrySet().stream()
                    .filter(e -> {
                        String lowerKey = e.getKey().toLowerCase();
                        String lowerName = provinceName.toLowerCase();
                        return lowerKey.contains(lowerName) || lowerName.contains(lowerKey);
                    })
                    .map(Map.Entry::getValue)
                    .findFirst()
                    .orElse(Collections.emptyList());

            provinceData.sort(Comparator.comparingInt(d -> extractDayFromDate(d.get("day"))));

            Map<Integer, Map<String, Object>> dayMap = new HashMap<>();
            Double provinceSltb = null; // số thuê bao chung cho tỉnh này

            for (Map<String, Object> data : provinceData) {
                int day = extractDayFromDate(data.get("day"));
                if (day > 0 && day <= currentDay) {
                    dayMap.put(day, data);
                    // Cột E sẽ fill sltb đầu tiên có trong list
                    if (provinceSltb == null && data.get("sltb") instanceof Number) {
                        provinceSltb = ((Number) data.get("sltb")).doubleValue();
                    }
                }
            }

            // ==== FILL CỘT E (LŨY KẾ SLTB) ====
            if (provinceSltb != null) {
                setCellValue(sheet, provinceRow, TABLE4_DATA_START_COL - 2, provinceSltb);
            }

            Double lastKnownSltb = provinceSltb;

            // ==== FILL CÁC NGÀY ====
            for (int day = 1; day <= currentDay; day++) {
                int colIndex = TABLE4_DATA_START_COL + (day - 1) * 3;

                Map<String, Object> data = dayMap.get(day);
                double slpa = 0.0;
                Double sltb = null;
                Double tlpa = 0.0;

                if (data != null) {
                    slpa = data.get("slpa") instanceof Number ? ((Number) data.get("slpa")).doubleValue() : 0.0;
                    sltb = data.get("sltb") instanceof Number ? ((Number) data.get("sltb")).doubleValue() : null;

                    if (data.get("tlpa") instanceof Number) {
                        tlpa = ((Number) data.get("tlpa")).doubleValue();
                    }
                }

                if (sltb == null && lastKnownSltb != null) sltb = lastKnownSltb;
                if (sltb != null) lastKnownSltb = sltb;

                setCellValue(sheet, provinceRow, colIndex, slpa);
                if (sltb != null) setCellValue(sheet, provinceRow, colIndex + 1, sltb);

                setCellValue(sheet, provinceRow, colIndex + 2, tlpa);
            }

        }

        // ==== TOTAL ROW ====
        List<Map<String, Object>> totalData = dataByProvince.get("Toàn quốc");
        if (totalData != null && !totalData.isEmpty()) {
            totalData.sort(Comparator.comparingInt(d -> extractDayFromDate(d.get("day"))));

            Map<Integer, Map<String, Object>> totalDayMap = new HashMap<>();
            Double totalSltb = null;

            for (Map<String, Object> data : totalData) {
                int day = extractDayFromDate(data.get("day"));
                if (day > 0 && day <= currentDay) {
                    totalDayMap.put(day, data);
                    if (totalSltb == null && data.get("sltb") instanceof Number) {
                        totalSltb = ((Number) data.get("sltb")).doubleValue();
                    }
                }
            }

            // Fill cột E của TOTAL row
            if (totalSltb != null) {
                setCellValue(sheet, TABLE4_TOTAL_ROW, TABLE4_DATA_START_COL - 2, totalSltb);
            }

            Double lastKnownTotalSltb = totalSltb;
            for (int day = 1; day <= currentDay; day++) {
                int colIndex = TABLE4_DATA_START_COL + (day - 1) * 3;

                Map<String, Object> data = totalDayMap.get(day);
                double slpa = 0.0;
                Double sltb = null;
                Double tlpa = 0.0; // luôn mặc định 0

                if (data != null) {
                    slpa = data.get("slpa") instanceof Number ? ((Number) data.get("slpa")).doubleValue() : 0.0;
                    sltb = data.get("sltb") instanceof Number ? ((Number) data.get("sltb")).doubleValue() : null;
                    if (data.get("tlpa") instanceof Number) {
                        tlpa = ((Number) data.get("tlpa")).doubleValue();
                    }
                }

                if (sltb == null && lastKnownTotalSltb != null) sltb = lastKnownTotalSltb;
                if (sltb != null) lastKnownTotalSltb = sltb;

                setCellValue(sheet, TABLE4_TOTAL_ROW, colIndex, slpa);
                if (sltb != null) setCellValue(sheet, TABLE4_TOTAL_ROW, colIndex + 1, sltb);
                setCellValue(sheet, TABLE4_TOTAL_ROW, colIndex + 2, tlpa); // luôn fill
            }

        }
    }

    private void fillTable5Kpi(Sheet sheet, Double progress3h) {
        if (progress3h == null) return;

        String text = String.format("Đánh giá so với KPI (>=%.2f%%)", progress3h);
        setCellValueTable5(sheet, text);
        log.debug("Filled Table 5 KPI: {}", text);
    }
    private void fillTable6Kpi(Sheet sheet, Double progress24h) {
        if (progress24h == null) return;

        String text = String.format("Đánh giá so với KPI (>=%.2f%%)", progress24h);
        setCellValueTable6(sheet, text);
    }
    private void fillTable7Kpi(Sheet sheet, Double progress48h) {
        if (progress48h == null) return;

        String text = String.format("Đánh giá so với KPI (>=%.2f%%)", progress48h);
        setCellValueTable7(sheet, text);
    }
    private void fillTable8Kpi(Sheet sheet, Double progress3hVip) {
        if (progress3hVip == null) return;

        String text = String.format("Đánh giá so với KPI (>=%.2f%%)", progress3hVip);
        setCellValueTable8(sheet, text);
    }

    private void fillHandleRateTotalRow(Sheet sheet, Map<String, List<Map<String, Object>>> dataByProvince, int currentDay, Double progress3h) {
        List<Map<String, Object>> totalData = dataByProvince.get("Toàn quốc");

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
        List<Map<String, Object>> totalData = dataByProvince.get("Toàn quốc");

        // Initialize TOTAL row with default values for past days, null for future days
        for (int day = 1; day <= 31; day++) {
            int colIndex = TABLE6_DATA_START_COL - 4 + (day - 1) * 3;
            if (day <= currentDay) {
                setCellValue(sheet, TABLE6_TOTAL_ROW, colIndex, 0.0);
                setCellValue(sheet, TABLE6_TOTAL_ROW, colIndex + 1, 0.0);
                setCellValue(sheet, TABLE6_TOTAL_ROW, colIndex + 2, 0.0);
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
    private void fillHandleRate3hByProvince(Sheet sheet, List<Map<String, Object>> handleRateData, Double progress3h, LocalDate currentDay) {
        if (handleRateData == null) handleRateData = List.of();
        int curDay = currentDay.getDayOfMonth();
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
                if (day <= curDay) {
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
        fillHandleRateTotalRow(sheet, dataByProvince, curDay, progress3h);
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


    private void fillHandleRate24hByProvince(Sheet sheet, List<Map<String, Object>> handleRateData, Double progress24h, LocalDate currentDate) {
        if (handleRateData == null) handleRateData = List.of();

        int currentDay = currentDate.getDayOfMonth();

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
    private void fillHandleRate48hByProvince(Sheet sheet, List<Map<String, Object>> handleRateData, Double progress48h, LocalDate currentDate) {
        if (handleRateData == null) handleRateData = List.of();

        int currentDay = currentDate.getDayOfMonth();

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
        List<Map<String, Object>> totalData = dataByProvince.get("Toàn quốc");

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
    private void fillHandleRate3hVipByProvince(Sheet sheet, List<Map<String, Object>> handleRateData, Double progress3hVip, LocalDate currentDate) {
        if (handleRateData == null) handleRateData = List.of();

        int currentDay = currentDate.getDayOfMonth();

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
                        setCellValue(sheet, finalProvinceRow, colIndex + 1, getNumberValue(data, "tong_sc_da_xu_ly"));
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
            cumulativeDaXuLy3hVip += getNumberValue(data, "da_xu_ly_3h_vip") != null
                    ? getNumberValue(data, "da_xu_ly_3h_vip") : 0.0;
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
        List<Map<String, Object>> totalData = dataByProvince.get("Toàn quốc");

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
                cumulativeDaXuLy48h += getNumberValue(data, "da_xu_ly_3h_vip") != null ? getNumberValue(data, "da_xu_ly_3h_vip") : 0.0;
                cumulativeTongSc += getNumberValue(data, "tong_sc_da_xu_ly") != null ? getNumberValue(data, "tong_sc_da_xu_ly") : 0.0;
            }
            for (Map<String, Object> data : totalData) {
                int day = extractDayFromDate(data.get("ngay"));
                if (day > 0 && day <= 31) {
                    int colIndex = TABLE7_DATA_START_COL + (day - 1) * 3;
                    setCellValue(sheet, TABLE8_TOTAL_ROW, colIndex, getNumberValue(data, "da_xu_ly_3h_vip"));
                    setCellValue(sheet, TABLE8_TOTAL_ROW, colIndex + 1, getNumberValue(data, "tong_sc_da_xu_ly"));
                    Double tyLe = getNumberValue(data, "da_xu_ly_3h_vip");
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
        LocalDate ld = extractLocalDate(dateObj);
        return ld != null ? ld.getDayOfMonth() : -1;
    }

    private String getCellStringValue(Sheet sheet, int rowIndex, int colIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) return null;

        Cell cell = row.getCell(colIndex);
        if (cell == null) return null;

        return cell.getStringCellValue().trim();
    }

    private void setCellValue(Row row, int colIndex, Object value) {
        if (row == null) return;

        Cell cell = row.getCell(colIndex);
        if (cell == null) {
            cell = row.createCell(colIndex);
        }

        if (value == null) {
            cell.setBlank();
        } else if (value instanceof Number) {
            cell.setCellValue(((Number) value).doubleValue());
        } else {
            cell.setCellValue(value.toString());
        }
    }
    private void setCellValue(Sheet sheet, int rowIndex, int colIndex, Object value) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }

        Cell cell = row.getCell(colIndex);
        if (cell == null) {
            cell = row.createCell(colIndex);
        }

        if (value == null) {
            cell.setBlank();
        } else if (value instanceof Number) {
            cell.setCellValue(((Number) value).doubleValue());
        } else {
            cell.setCellValue(value.toString());
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

    private void fillSheet2KpiData(Sheet sheet2, TransmissionChanelDTO data, LocalDate currentDate) {
        if (sheet2 == null) {
            log.warn("Sheet 2 (KPI VTS) not found, skipping KPI data population");
            return;
        }

        int currentDay = currentDate.getDayOfMonth();

        // Initialize KPI values
        Double progress3h = null;
        Double progress24h = null;
        Double progress48h = null;
        Double satisfyLevel = null;

        // Get KPI progress data and extract each KPI type
        if (data.getOnTimeHandle() != null && !data.getOnTimeHandle().isEmpty()) {

            // Extract 3H KPI
            Optional<Map<String, Object>> kpi3h = data.getOnTimeHandle().stream()
                    .filter(row -> {
                        Object incidentName = row.get("incidentName");
                        return incidentName != null && incidentName.toString().equals("Tỉ lệ sự cố xử lý trong 3h");
                    })
                    .findFirst();

            if (kpi3h.isPresent()) {
                Double val = getNumberValue(kpi3h.get(), "kpi_progress");
                if (val != null) {
                    progress3h = val / 100; // Convert percentage to decimal
                }
            }

            // Extract 24H KPI
            Optional<Map<String, Object>> kpi24h = data.getOnTimeHandle().stream()
                    .filter(row -> {
                        Object incidentName = row.get("incidentName");
                        return incidentName != null && incidentName.toString().equals("Tỉ lệ sự cố xử lý trong 24h");
                    })
                    .findFirst();

            if (kpi24h.isPresent()) {
                Double val = getNumberValue(kpi24h.get(), "kpi_progress");
                if (val != null) {
                    progress24h = val / 100; // Convert percentage to decimal
                }
            }

            // Extract 48H KPI
            Optional<Map<String, Object>> kpi48h = data.getOnTimeHandle().stream()
                    .filter(row -> {
                        Object incidentName = row.get("incidentName");
                        return incidentName != null && incidentName.toString().equals("Tỉ lệ sự cố xử lý trong 48h");
                    })
                    .findFirst();

            if (kpi48h.isPresent()) {
                Double val = getNumberValue(kpi48h.get(), "kpi_progress");
                if (val != null) {
                    progress48h = val / 100; // Convert percentage to decimal
                }
            }

            // Extract Satisfaction Level KPI (nếu có tên khác, cần điều chỉnh)
            Optional<Map<String, Object>> kpiSatisfy = data.getOnTimeHandle().stream()
                    .filter(row -> {
                        Object incidentName = row.get("incidentName");
                        return incidentName != null &&
                                (incidentName.toString().contains("Mức độ hài lòng"));
                    })
                    .findFirst();

            if (kpiSatisfy.isPresent()) {
                Double val = getNumberValue(kpiSatisfy.get(), "kpi_progress");
                if (val != null) {
                    satisfyLevel = val / 100; // Convert percentage to decimal
                }
            }
        }

        // Fill KPI 3H section (rows C5, C7, C8)
        fillSheet2KpiRow(sheet2, SHEET2_KPI_3H_ROW, progress3h, currentDay, "Tỉ lệ sự cố xử lý trong 3h");
        fillSheet2ActualDataRow(sheet2, SHEET2_3H_DA_XU_LI_ROW, data.getHandleRate3h(), "da_xu_ly_3h", currentDate);
        fillSheet2ActualDataRow(sheet2, SHEET2_3H_TONG_SC_ROW, data.getHandleRate3h(), "tong_sc_da_xu_ly", currentDate);

        // Fill KPI 24H section (rows C11, C13, C14)
        fillSheet2KpiRow(sheet2, SHEET2_KPI_24H_ROW, progress24h, currentDay, "Tỉ lệ sự cố xử lý trong 24h");
        fillSheet2ActualDataRow(sheet2, SHEET2_24H_DA_XU_LI_ROW, data.getHandleRate24h(), "da_xu_ly_24h", currentDate);
        fillSheet2ActualDataRow(sheet2, SHEET2_24H_TONG_SC_ROW, data.getHandleRate24h(), "tong_sc_da_xu_ly", currentDate);

        // Fill KPI 48H section (rows C17, C19, C20)
        fillSheet2KpiRow(sheet2, SHEET2_KPI_48H_ROW, progress48h, currentDay, "Tỉ lệ sự cố xử lý trong 48h");
        fillSheet2ActualDataRow(sheet2, SHEET2_48H_DA_XU_LI_ROW, data.getHandleRate48h(), "da_xu_ly_48h", currentDate);
        fillSheet2ActualDataRow(sheet2, SHEET2_48H_TONG_SC_ROW, data.getHandleRate48h(), "tong_sc_da_xu_ly", currentDate);
        fillSheet2KpiRow(sheet2, SHEET2_KPI_SATISFY_ROW, satisfyLevel, currentDay, "Mức độ hài lòng");

    }
    private void fillSheet2KpiRow(Sheet sheet, int rowIndex, Double kpiValue, int currentDay, String kpiType) {
        if (kpiValue == null) {
            log.warn("KPI value is null for {}, skipping row {}", kpiType, rowIndex + 1);
            return;
        }

        for (int day = 1; day <= 35; day++) {
            int colIndex = SHEET2_DATA_START_COL - 1 + day ; // C=2, D=3, ..., AK=36 (KPI starts from column C)

            // Fill KPI up to day + 1 (include the next day), leave further days empty
            if (day <= currentDay + 1) {
                setCellValue(sheet, rowIndex, colIndex, kpiValue);
            } else {
                setCellValue(sheet, rowIndex, colIndex, null);
            }
        }

        log.debug("Filled {} row {} with value {} for {} days", kpiType, rowIndex + 1, kpiValue, currentDay);
    }

    private void fillSheet2ActualDataRow(Sheet sheet, int rowIndex, List<Map<String, Object>> handleRateData,
                                         String dataKey, LocalDate currentDate) {
        if (sheet == null) return;

        Row row = sheet.getRow(rowIndex);
        if (row == null) row = sheet.createRow(rowIndex);

        Map<String, List<Map<String, Object>>> dataByProvince = handleRateData != null
                ? handleRateData.stream().filter(Objects::nonNull).collect(Collectors.groupingBy(
                d -> d.get("tinh") != null ? d.get("tinh").toString().trim() : "unknown"
        ))
                : Map.of();

        List<Map<String, Object>> totalData = dataByProvince.getOrDefault("Toàn quốc", List.of());

        Map<LocalDate, Map<String, Object>> dataByDate = totalData.stream()
                .filter(Objects::nonNull)
                .collect(Collectors.toMap(
                        d -> {
                            Object ngay = d.get("ngay");
                            LocalDate parsed = extractLocalDate(ngay);
                            return parsed; // may be null
                        },
                        d -> d,
                        (existing, replacement) -> replacement,
                        LinkedHashMap::new
                ));
        // Remove any null keys caused by invalid/null dates
        dataByDate.remove(null);

        // Use the current date from the parameter
        LocalDate baseDate = currentDate;
        int currentDay = currentDate.getDayOfMonth();

        for (int day = 1; day <= 35; day++) {
            int colIndex = SHEET2_DATA_START_COL + day; // D=3, E=4, ..., AL=37 (data starts from column D)

            Cell cell = row.getCell(colIndex);
            if (cell == null) cell = row.createCell(colIndex);

            if (day <= currentDay) {
                LocalDate date = baseDate.withDayOfMonth(day); // trực tiếp LocalDate
                Map<String, Object> data = dataByDate.get(date);
                
                Double value = 0.0;
                if (data != null) {
                    value = getNumberValue(data, dataKey);
                    if (value == null) value = 0.0;
                }
                
                cell.setCellValue(value);
            } else {
                cell.setBlank();
            }
        }
    }

    private void fillSatisfactionLevel(Sheet sheet, TransmissionChanelDTO excelData, LocalDate currentDate) {
        List<Map<String, Object>> satisfactionData = excelData.getSatisfactionLevel();
        if (satisfactionData == null) satisfactionData = List.of();

        Row numeratorRow = sheet.getRow(47);   // Row 48 in Excel
        Row denominatorRow = sheet.getRow(48); // Row 49 in Excel

        if (numeratorRow == null) numeratorRow = sheet.createRow(47);
        if (denominatorRow == null) denominatorRow = sheet.createRow(48);

        int startCol = 3; // Column D

        // Group data by date using LocalDate, ignoring null/invalid dates
        Map<LocalDate, Map<String, Object>> dataByDate = satisfactionData.stream()
                .filter(Objects::nonNull)
                .collect(Collectors.toMap(
                        r -> extractLocalDate(r.get("ngay")),
                        r -> r,
                        (a, b) -> b,
                        LinkedHashMap::new
                ));
        dataByDate.remove(null);

        int currentDay = currentDate.getDayOfMonth();

        for (int day = 1; day <= currentDay; day++) {
            int colIndex = startCol + (day - 1);

            LocalDate dayDate = currentDate.withDayOfMonth(day);
            Map<String, Object> record = dataByDate.get(dayDate);

            Double numerator = null;
            Double denominator = null;

            if (record != null) {
                Object numObj = record.get("tu_so_hai_long");
                Object denObj = record.get("mau_so_phan_hoi");

                if (numObj instanceof Number) numerator = ((Number) numObj).doubleValue();
                if (denObj instanceof Number) denominator = ((Number) denObj).doubleValue();
            }

            Cell numeratorCell = numeratorRow.getCell(colIndex);
            if (numeratorCell == null) numeratorCell = numeratorRow.createCell(colIndex);
            numeratorCell.setCellValue(numerator != null ? numerator : 0);

            Cell denominatorCell = denominatorRow.getCell(colIndex);
            if (denominatorCell == null) denominatorCell = denominatorRow.createCell(colIndex);
            denominatorCell.setCellValue(denominator != null ? denominator : 0);
        }
    }

    private void fillAvgHandleTime(Sheet sheet, TransmissionChanelDTO excelData, LocalDate currentDate) {
        List<Map<String, Object>> avgHandleTimeData = excelData.getAvgTimeHandle();
        if (avgHandleTimeData == null) avgHandleTimeData = List.of();

        // Rows 32 and 33 (POI index starts at 0 -> row 31 và 32)
        Row totalTimeRow = sheet.getRow(31);   // D32
        Row totalCountRow = sheet.getRow(32);  // D33

        if (totalTimeRow == null) totalTimeRow = sheet.createRow(31);
        if (totalCountRow == null) totalCountRow = sheet.createRow(32);

        int startCol = 3; // D là cột số 3 (0-based index)

        // Map dữ liệu theo ngày để lookup using LocalDate, ignoring null/invalid dates
        Map<LocalDate, Map<String, Object>> dataByDate = avgHandleTimeData.stream()
                .filter(Objects::nonNull)
                .collect(Collectors.toMap(
                        r -> extractLocalDate(r.get("ngay")),
                        r -> r,
                        (a, b) -> b,
                        LinkedHashMap::new
                ));
        dataByDate.remove(null);

        int currentDay = currentDate.getDayOfMonth();
        for (int day = 1; day <= currentDay; day++) {
            int colIndex = startCol + (day - 1);
            LocalDate dayDate = currentDate.withDayOfMonth(day);
            Map<String, Object> record = dataByDate.get(dayDate);

            double totalHours = 0;
            double totalCount = 0;

            if (record != null) {
                Object hoursObj = record.get("tong_thoi_gian_xu_ly");
                Object countObj = record.get("so_luong_phan_anh");

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

    private void fillTable3Header(Sheet sheet, LocalDate currentDate) {
        Row headerRow = sheet.getRow(30); // Row 31 in Excel (0-based index)
        if (headerRow == null) headerRow = sheet.createRow(30);

        // Get current month and year
        int currentMonth = currentDate.getMonthValue();
        int currentYear = currentDate.getYear();
        
        // Format month name in Vietnamese
        String monthName = getVietnameseMonthName(currentMonth);
        
        // Column C (index 2): "Lũy kế T7.2025"
        Cell cellC = headerRow.getCell(2);
        if (cellC == null) cellC = headerRow.createCell(2);
        cellC.setCellValue("Lũy kế " + monthName + "." + currentYear);
        
        // Column D (index 3): "TB ngày T7.2025"
        Cell cellD = headerRow.getCell(3);
        if (cellD == null) cellD = headerRow.createCell(3);
        cellD.setCellValue("TB ngày " + monthName + "." + currentYear);
        
        // Column E (index 4): Use common function
        fillTableHeader(sheet, currentDate, 30, 4);
    }

    private String getVietnameseMonthName(int month) {
        String[] monthNames = {
            "", "T1", "T2", "T3", "T4", "T5", "T6",
            "T7", "T8", "T9", "T10", "T11", "T12"
        };
        return monthNames[month];
    }

    private void fillTable4Header(Sheet sheet, LocalDate currentDate) {
        fillTableHeader(sheet, currentDate, 41, 6); // Row 42, Column G
    }

    private void fillG41Date(Sheet sheet, LocalDate currentDate) {
        Row row = sheet.getRow(40); // Row 41 (0-based index)
        if (row == null) row = sheet.createRow(40);

        Cell cell = row.getCell(6); // Column G (0-based index)
        if (cell == null) cell = row.createCell(6);

        // Hiển thị chỉ số ngày
        int currentDay = currentDate.getDayOfMonth();
        cell.setCellValue(currentDay);
    }

    private void fillTable5Header(Sheet sheet, LocalDate currentDate) {
        fillTableHeader(sheet, currentDate, 81, 7); // Row 82, Column H
    }

    private void fillTable6Header(Sheet sheet, LocalDate currentDate) {
        fillTableHeader(sheet, currentDate, 121, 7); // Row 122, Column H
    }

    private void fillTable7Header(Sheet sheet, LocalDate currentDate) {
        fillTableHeader(sheet, currentDate, 161, 7); // Row 162, Column H
    }

    private void fillTable8Header(Sheet sheet, LocalDate currentDate) {
        fillTableHeader(sheet, currentDate, 201, 7); // Row 202, Column H
    }

    private void fillTableHeader(Sheet sheet, LocalDate currentDate, int rowIndex, int colIndex) {
        Row headerRow = sheet.getRow(rowIndex);
        if (headerRow == null) headerRow = sheet.createRow(rowIndex);

        // Get current month and year
        int currentMonth = currentDate.getMonthValue();
        int currentYear = currentDate.getYear();
        
        // Fill cell with "tháng/1/năm" format
        Cell cell = headerRow.getCell(colIndex);
        if (cell == null) cell = headerRow.createCell(colIndex);
        String firstDayOfMonth = currentMonth + "/1/" + currentYear;
        cell.setCellValue(firstDayOfMonth);
    }

    private void fillMonthlySummaryData(Sheet sheet, TransmissionChanelDTO data, LocalDate currentDate) {
        // Fill T1 to T12 data (D342 to O344) from getIncidentMonthlySummary
        fillMonthlySummaryT1ToT12(sheet, data.getIncidentMonthlySummary());
        
        // Fill 8 days data (P342 to W344) from getLast8daysAndAvgMonth
        fillLast8DaysData(sheet, data.getLast8daysAndAvgMonth(), currentDate);
    }

    private void fillMonthlySummaryT1ToT12(Sheet sheet, List<Map<String, Object>> monthlySummaryData) {
        if (monthlySummaryData == null || monthlySummaryData.isEmpty()) return;

        // T1 to T12 columns: D342 to O344 (columns 3 to 14, 0-based)
        int startCol = 3; // Column D
        int startRow = 341; // Row 342 (0-based)
        
        for (Map<String, Object> rowData : monthlySummaryData) {
            String type = (String) rowData.get("type");
            int rowIndex = startRow;
            
            // Determine which row to fill based on type
            if ("Sự cố".equals(type)) {
                rowIndex = startRow + 1; // Row 343
            } else if ("TLPA/1000TB".equals(type)) {
                rowIndex = startRow + 2; // Row 344
            } else {
                continue; // Skip unknown types
            }
            
            // Fill T1 to T12 data
            for (int month = 1; month <= 12; month++) {
                String columnKey = "T" + month;
                Object value = rowData.get(columnKey);
                if (value instanceof Number) {
                    setCellValue(sheet, rowIndex, startCol + month - 1, ((Number) value).doubleValue());
                }
            }
        }
    }

    private void fillLast8DaysData(Sheet sheet, List<Map<String, Object>> last8DaysData, LocalDate currentDate) {
        if (last8DaysData == null || last8DaysData.isEmpty()) return;

        // D342 to O344 (columns 3 to 14, 0-based) for months 1-12
        int monthlyStartCol = 3; // Column D
        // P342 to W344 (columns 15 to 22, 0-based) for 8 days
        int dailyStartCol = 15; // Column P
        int headerRow = 341; // Row 342 (0-based) - for headers
        int suCoRow = 342; // Row 343 (0-based) - for "Sự cố" data
        int tlpaRow = 343; // Row 344 (0-based) - for "TLPA/1000TB" data

        // Get data for current month only (for cumulative values)
        List<Map<String, Object>> currentMonthData = last8DaysData.stream()
                .filter(data -> {
                    Object dayObj = data.get("day");
                    if (dayObj == null) return false;
                    
                    LocalDate dataDate = extractLocalDate(dayObj);
                    if (dataDate == null) return false;
                    
                    return dataDate.getMonthValue() == currentDate.getMonthValue() && 
                           dataDate.getYear() == currentDate.getYear();
                })
                .sorted((a, b) -> {
                    Object dayA = a.get("day");
                    Object dayB = b.get("day");
                    if (dayA == null || dayB == null) return 0;
                    
                    LocalDate dateA = extractLocalDate(dayA);
                    LocalDate dateB = extractLocalDate(dayB);
                    if (dateA == null || dateB == null) return 0;
                    
                    return dateA.compareTo(dateB);
                })
                .collect(Collectors.toList());

        // Get the last 8 days of data, sorted by day
        List<Map<String, Object>> sortedData = last8DaysData.stream()
                .sorted((a, b) -> {
                    Object dayA = a.get("day");
                    Object dayB = b.get("day");
                    if (dayA == null || dayB == null) return 0;

                    LocalDate dateA = extractLocalDate(dayA);
                    LocalDate dateB = extractLocalDate(dayB);
                    if (dateA == null || dateB == null) return 0;

                    return dateA.compareTo(dateB);
                })
                .collect(Collectors.toList());

        // Fill headers (months) in row 342
        Row headerRowObj = sheet.getRow(headerRow);
        if (headerRowObj == null) headerRowObj = sheet.createRow(headerRow);

        // Fill "Sự cố" data in row 343
        Row suCoRowObj = sheet.getRow(suCoRow);
        if (suCoRowObj == null) suCoRowObj = sheet.createRow(suCoRow);

        // Fill "TLPA/1000TB" data in row 344
        Row tlpaRowObj = sheet.getRow(tlpaRow);
        if (tlpaRowObj == null) tlpaRowObj = sheet.createRow(tlpaRow);

        // Fill T1 to T12 headers (columns D to O)
        String[] monthHeaders = {"T1", "T2", "T3", "T4", "T5", "T6", "T7", "T8", "T9", "T10", "T11", "T12"};
        for (int monthIndex = 0; monthIndex < 12; monthIndex++) {
            int colIndex = monthlyStartCol + monthIndex;
            setCellValue(sheet, headerRow, colIndex, monthHeaders[monthIndex]);
        }

        // Fill cumulative data for current month (D-O)
        if (!currentMonthData.isEmpty()) {
            // Get the latest cumulative data for the current month
            Map<String, Object> latestData = currentMonthData.get(currentMonthData.size() - 1);
            int currentMonthCol = currentDate.getMonthValue() - 1; // Convert to 0-based index
            int colIndex = monthlyStartCol + currentMonthCol;
            
            // Fill "Sự cố" data (Row 343) with luyKeTotalPerDayParent
            Object luyKeTotalPerDayParent = latestData.get("luyKeTotalPerDayParent");
            if (luyKeTotalPerDayParent instanceof Number) {
                setCellValue(sheet, suCoRow, colIndex, ((Number) luyKeTotalPerDayParent).doubleValue());
            }
            
            // Fill "TLPA/1000TB" data (Row 344) with tlpaParentLuyKe
            Object tlpaParentLuyKe = latestData.get("tlpaParentLuyKe");
            if (tlpaParentLuyKe instanceof Number) {
                setCellValue(sheet, tlpaRow, colIndex, ((Number) tlpaParentLuyKe).doubleValue());
            }
        }

        // Generate dynamic headers for last 8 days from currentDate (W back to P)
        // W = currentDate, V = currentDate-1, ..., P = currentDate-7
        for (int offset = 0; offset < 8; offset++) {
            LocalDate headerDate = currentDate.minusDays(offset);
            int colIndex = dailyStartCol + (7 - offset);

            String dateStr = String.format("%02d/%02d", headerDate.getDayOfMonth(), headerDate.getMonthValue());
            setCellValue(sheet, headerRow, colIndex, dateStr);
        }

        // Fill data for available days (W back to P)
        int n = Math.min(8, sortedData.size());
        for (int i = 0; i < n; i++) {
            Map<String, Object> dayData = sortedData.get(sortedData.size() - 1 - i); // latest first
            int colIndex = dailyStartCol + (7 - i);

            Object totalPerDay = dayData.get("totalPerDayParent");
            if (totalPerDay instanceof Number) {
                setCellValue(sheet, suCoRow, colIndex, ((Number) totalPerDay).doubleValue());
            }

            Object tlpaParent = dayData.get("tlpaParent");
            if (tlpaParent instanceof Number) {
                setCellValue(sheet, tlpaRow, colIndex, ((Number) tlpaParent).doubleValue());
            }
        }
    }

    private LocalDate extractLocalDate(Object dateObj) {
        try {
            if (dateObj == null) return null;
            if (dateObj instanceof java.sql.Date) {
                return ((java.sql.Date) dateObj).toLocalDate();
            } else if (dateObj instanceof java.util.Date) {
                return ((java.util.Date) dateObj).toInstant()
                        .atZone(java.time.ZoneId.systemDefault())
                        .toLocalDate();
            } else if (dateObj instanceof java.time.LocalDate) {
                return (java.time.LocalDate) dateObj;
            } else if (dateObj instanceof CharSequence) {
                String raw = dateObj.toString().trim();
                if (raw.isEmpty()) return null;
                // Try common formats: yyyy-MM-dd, dd/MM/yyyy, yyyy/MM/dd
                DateTimeFormatter[] fmts = new DateTimeFormatter[] {
                        DateTimeFormatter.ISO_LOCAL_DATE,
                        DateTimeFormatter.ofPattern("dd/MM/yyyy"),
                        DateTimeFormatter.ofPattern("yyyy/MM/dd")
                };
                for (DateTimeFormatter f : fmts) {
                    try { return LocalDate.parse(raw, f); } catch (Exception ignore) {}
                }
                // Fallback: replace '-' with '/' and try dd/MM/yyyy if it looks like dd-MM-yyyy
                if (raw.matches("\\d{2}-\\d{2}-\\d{4}")) {
                    try { return LocalDate.parse(raw.replace('-', '/'), DateTimeFormatter.ofPattern("dd/MM/yyyy")); } catch (Exception ignore) {}
                }
            }
        } catch (Exception e) {
            log.warn("Failed to extract LocalDate from object: {}", dateObj, e);
        }
        return null;
    }

}