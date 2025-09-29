package com.example.transmission.domain;
import lombok.Data;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.Table;
import java.util.Date;

@Entity
@Table(name = "d_vcoc_problem_type")
@Data
public class ProblemType {
        @Id
        @Column(name = "prob_type_id")
        private Integer probTypeId;

        @Column(name = "prob_group_id")
        private Integer probGroupId;

        private String name;
        private String code;
        private String description;

        @Column(name = "problem_template")
        private String problemTemplate;

        private Integer status;

        @Column(name = "accept_source_required")
        private Integer acceptSourceRequired;

        private Integer iscoordinate;

        @Column(name = "create_by")
        private String createBy;

        @Column(name = "create_date")
        private Date createDate;

        @Column(name = "from_work_time")
        private String fromWorkTime;

        @Column(name = "to_work_time")
        private String toWorkTime;

        @Column(name = "config_accept_date")
        private Integer configAcceptDate;

        private Integer type;

        @Column(name = "time_process_type")
        private Integer timeProcessType;

        @Column(name = "evaluate_gnoc_type")
        private String evaluateGnocType;

        @Column(name = "problem_template_mbccs")
        private String problemTemplateMbccs;

        @Column(name = "check_add_forward_team")
        private Integer checkAddForwardTeam;

        private Integer passing;

        @Column(name = "show_mbccs")
        private Integer showMbccs;

        @Column(name = "customer_display_content")
        private String customerDisplayContent;

        @Column(name = "is_customer")
        private Integer isCustomer;

        @Column(name = "check_add_cooperation_time")
        private Integer checkAddCooperationTime;

        @Column(name = "show_system")
        private String showSystem;
}


