package com.example.transmission.domain;

import lombok.Data;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.Table;
import java.math.BigDecimal;
import java.util.Date;

@Entity
@Table(name = "d_vcoc_rpt_problem_gpdn_process")
@Data
public class ProblemProcess {
    @Id
    private Long id;

    @Column(name = "prob_group_id")
    private Integer probGroupId;

    @Column(name = "problem_id")
    private Long problemId;

    @Column(name = "isdn_account")
    private String isdnAccount;

    @Column(name = "prob_group_name")
    private String probGroupName;

    @Column(name = "prob_group_child_name")
    private String probGroupChildName;

    @Column(name = "prob_type_name")
    private String probTypeName;

    @Column(name = "prob_status")
    private String probStatus;

    @Column(name = "cust_accept_level")
    private Integer custAcceptLevel;

    @Column(name = "process_time_total")
    private BigDecimal processTimeTotal;

    @Column(name = "process_current")
    private BigDecimal processCurrent;

    @Column(name = "process_time_total_gnoc")
    private BigDecimal processTimeTotalGnoc;

    @Column(name = "prd_id")
    private Integer prdId;

    @Column(name = "sync_date")
    private Date syncDate;

    @Column(name = "report_type")
    private String reportType;
}

