package com.example.transmission.domain;

import lombok.Data;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.Table;
import java.util.Date;

@Entity
@Data
@Table(name = "d_vcoc_problem_group")
public class ProblemGroup {
    @Id
    @Column(name = "prob_group_id")
    private Integer probGroupId;

    @Column(name = "is_process")
    private Integer isProcess;

    @Column(name = "not_req_cause_exp")
    private Integer notReqCauseExp;

    private String code;
    private String name;
    private Integer status;

    @Column(name = "create_date")
    private Date createDate;

    @Column(name = "last_update")
    private Date lastUpdate;

    private String description;

    @Column(name = "create_by")
    private String createBy;

    @Column(name = "from_work_time")
    private String fromWorkTime;

    @Column(name = "to_work_time")
    private String toWorkTime;

    @Column(name = "parent_id")
    private Integer parentId;

    @Column(name = "problem_level_id")
    private Integer problemLevelId;

    @Column(name = "show_mbccs")
    private Integer showMbccs;
}
