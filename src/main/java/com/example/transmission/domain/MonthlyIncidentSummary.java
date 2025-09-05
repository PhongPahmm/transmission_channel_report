package com.example.transmission.domain;

import javax.persistence.*;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.AllArgsConstructor;

import java.math.BigDecimal;

@Entity
@Table(name = "rp_monthly_incidents_summary")
@Data
@NoArgsConstructor
@AllArgsConstructor
public class MonthlyIncidentSummary {
    
    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;
    
    @Column(name = "report_year", nullable = false)
    private Integer reportYear;
    
    @Column(name = "report_month", nullable = false)
    private Byte reportMonth;
    
    @Column(name = "avg_incidents", nullable = false, precision = 10, scale = 4)
    private BigDecimal avgIncidents;
    
    @Column(name = "ratio_per_1000", nullable = false, precision = 10, scale = 4)
    private BigDecimal ratioPer1000;
}
