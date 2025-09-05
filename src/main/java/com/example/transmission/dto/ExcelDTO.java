package com.example.transmission.dto;

import com.example.transmission.domain.ServiceSubscriber;
import lombok.Builder;
import lombok.Data;

import java.util.List;
import java.util.Map;

@Data
@Builder
public class ExcelDTO {
    private List<Map<String, Object>> categorySummaryResults;
    private List<Map<String, Object>> complaintDayResults;
    private List<ServiceSubscriber> serviceSubscribers;
    private List<Map<String, Object>> onTimeHandle;
    private List<Map<String, Object>> progressKpi;
    private List<Map<String, Object>> totalHandle;
    private List<Map<String, Object>> results;
    private List<Map<String, Object>> dailyStats;

}
