package com.example.transmission.dto;

import lombok.Data;

import java.util.List;
import java.util.Map;

@Data
public class TransmissionChanelDTO {
    private List<Map<String, Object>> categorySummaryResults;
    private List<Map<String, Object>> onTimeHandle;
    private List<Map<String, Object>> dailyStats;
    private List<Map<String, Object>> complaintRate;
    private List<Map<String, Object>> handleRate3h;
    private List<Map<String, Object>> handleRate24h;
    private List<Map<String, Object>> handleRate48h;
    private List<Map<String, Object>> handleRate3hVip;
    private List<Map<String, Object>> satisfactionLevel;
    private List<Map<String, Object>> avgTimeHandle;
    private List<Map<String, Object>> provinceSubscribers;
    private List<Map<String, Object>> last8daysAndAvgMonth;
    private List<Map<String, Object>> incidentMonthlySummary;

    // Default constructor
    public TransmissionChanelDTO() {
    }

    // All-args constructor
    public TransmissionChanelDTO(List<Map<String, Object>> categorySummaryResults,
                                List<Map<String, Object>> onTimeHandle,
                                List<Map<String, Object>> dailyStats,
                                List<Map<String, Object>> complaintRate,
                                List<Map<String, Object>> handleRate3h,
                                List<Map<String, Object>> handleRate24h,
                                List<Map<String, Object>> handleRate48h,
                                List<Map<String, Object>> handleRate3hVip,
                                List<Map<String, Object>> satisfactionLevel,
                                List<Map<String, Object>> avgTimeHandle,
                                List<Map<String, Object>> provinceSubscribers,
                                List<Map<String, Object>> last8daysAndAvgMonth,
                                List<Map<String, Object>> incidentMonthlySummary) {
        this.categorySummaryResults = categorySummaryResults;
        this.onTimeHandle = onTimeHandle;
        this.dailyStats = dailyStats;
        this.complaintRate = complaintRate;
        this.handleRate3h = handleRate3h;
        this.handleRate24h = handleRate24h;
        this.handleRate48h = handleRate48h;
        this.handleRate3hVip = handleRate3hVip;
        this.satisfactionLevel = satisfactionLevel;
        this.avgTimeHandle = avgTimeHandle;
        this.provinceSubscribers = provinceSubscribers;
        this.last8daysAndAvgMonth = last8daysAndAvgMonth;
        this.incidentMonthlySummary = incidentMonthlySummary;
    }
}
