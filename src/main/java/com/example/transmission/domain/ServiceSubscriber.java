package com.example.transmission.domain;

import javax.persistence.*;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.AllArgsConstructor;

import java.math.BigDecimal;

@Entity
@Table(name = "rp_service_subscribers")
@Data
@NoArgsConstructor
@AllArgsConstructor
public class ServiceSubscriber {
    
    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;

    @Column(name = "complaint_group", nullable = false, length = 100)
    private String complaintGroup;

    @Column(name = "service_name", nullable = false, length = 100)
    private String serviceName;
    
    @Column(name = "total_subscribers", nullable = false)
    private Long totalSubscribers;
    
    @Column(name = "service_kpi", nullable = false, precision = 5, scale = 2)
    private BigDecimal serviceKpi;

    @Column(name = "complaint_group_kpi", nullable = false, precision = 5, scale = 2)
    private BigDecimal complaintGroupKpi;
}
