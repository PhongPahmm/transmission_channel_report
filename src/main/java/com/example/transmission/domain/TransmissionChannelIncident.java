package com.example.transmission.domain;

import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.AllArgsConstructor;

import javax.persistence.*;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;

@Entity
@Table(name = "rp_transmission_channel_incidents")
@Data
@NoArgsConstructor
@AllArgsConstructor
public class TransmissionChannelIncident {
    
    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;
    
    @Column(name = "received_date")
    private LocalDate receivedDate;
    
    @Column(name = "complaint_group")
    private String complaintGroup;
    
    @Column(name = "category")
    private String category;
    
    @Column(name = "complaint_type")
    private String complaintType;
    
    @Column(name = "complaint_content", columnDefinition = "TEXT")
    private String complaintContent;
    
    @Column(name = "complaint_status")
    private String complaintStatus;
    
    @Column(name = "processing_result", columnDefinition = "TEXT")
    private String processingResult;
    
    @Column(name = "satisfaction_level")
    private String satisfactionLevel;
    
    @Column(name = "total_processing_hours", precision = 5, scale = 2)
    private BigDecimal totalProcessingHours;
    
    @Column(name = "root_cause_lvl1")
    private String rootCauseLvl1;
    
    @Column(name = "root_cause_lvl2")
    private String rootCauseLvl2;
    
    @Column(name = "root_cause_lvl3")
    private String rootCauseLvl3;
    
    @Column(name = "progress")
    private String progress;
    
    @Column(name = "gnoc_processing_hours", precision = 5, scale = 2)
    private BigDecimal gnocProcessingHours;
    
    @Column(name = "spm_processing_hours", precision = 5, scale = 2)
    private BigDecimal spmProcessingHours;
    
    @Column(name = "gnoc_progress")
    private String gnocProgress;
    
    @Column(name = "gnoc_appointment_hours", precision = 5, scale = 2)
    private BigDecimal gnocAppointmentHours;
    
    @Column(name = "response_time")
    private LocalDateTime responseTime;
    
    @Column(name = "critical_channel")
    private String criticalChannel;
    
    @Column(name = "actual_processing_hours", precision = 5, scale = 2)
    private BigDecimal actualProcessingHours;
    
    @Column(name = "kv_region")
    private String kvRegion;
    
    @Column(name = "new_province")
    private String newProvince;
    
    @Column(name = "new_ward")
    private String newWard;
}
