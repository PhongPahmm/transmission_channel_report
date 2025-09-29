package com.example.transmission.domain;

import lombok.Data;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.Table;
import java.time.LocalDateTime;

@Entity
@Data
@Table(name = "d_vcoc_problem_final_end_date")
public class ProblemFinalEndDate {
    @Id
    @Column(name = "problem_id")
    private Long problemId;

    @Column(name = "district")
    private String district;

    @Column(name = "precinct")
    private String precinct;

    @Column(name = "cause_id")
    private Integer causeId;

    @Column(name = "process_result_content", columnDefinition = "TEXT")
    private String processResultContent;

    @Column(name = "prob_priority_id")
    private Integer probPriorityId;

    @Column(name = "problem_level_id")
    private Integer problemLevelId;

    @Column(name = "prob_channel_id")
    private Integer probChannelId;

    @Column(name = "prob_type_id")
    private Integer probTypeId;

    @Column(name = "sat_level_id")
    private Integer satLevelId;

    @Column(name = "user_accept")
    private String userAccept;

    @Column(name = "shop_accept_id")
    private Integer shopAcceptId;

    @Column(name = "parent_problem_id")
    private Long parentProblemId;

    @Column(name = "last_process_time")
    private LocalDateTime lastProcessTime;

    @Column(name = "num_contact_cust")
    private Integer numContactCust;

    @Column(name = "end_date")
    private LocalDateTime endDate;

    @Column(name = "problem_content", columnDefinition = "TEXT")
    private String problemContent;

    @Column(name = "note", columnDefinition = "TEXT")
    private String note;

    @Column(name = "cust_limit_date")
    private LocalDateTime custLimitDate;

    @Column(name = "status")
    private String status;

    @Column(name = "res_date")
    private LocalDateTime resDate;

    @Column(name = "res_limit_date")
    private LocalDateTime resLimitDate;

    @Column(name = "province")
    private String province;

    @Column(name = "contact_number")
    private String contactNumber;

    @Column(name = "duplicate_id")
    private Long duplicateId;

    @Column(name = "complainer_name")
    private String complainerName;

    @Column(name = "complainer_phone")
    private String complainerPhone;

    @Column(name = "complainer_address")
    private String complainerAddress;

    @Column(name = "create_date")
    private LocalDateTime createDate;

    @Column(name = "prob_accept_type_id")
    private Integer probAcceptTypeId;

    @Column(name = "assign_status")
    private String assignStatus;

    @Column(name = "cooperation_exp_date")
    private LocalDateTime cooperationExpDate;

    @Column(name = "pre_result", columnDefinition = "TEXT")
    private String preResult;

    @Column(name = "complainer_email")
    private String complainerEmail;

    @Column(name = "customer_text", columnDefinition = "TEXT")
    private String customerText;

    @Column(name = "arise")
    private LocalDateTime arise;

    @Column(name = "shop_process_id")
    private Integer shopProcessId;

    @Column(name = "user_process")
    private String userProcess;

    @Column(name = "return_status")
    private String returnStatus;

    @Column(name = "return_reason", columnDefinition = "TEXT")
    private String returnReason;

    @Column(name = "processing_note", columnDefinition = "TEXT")
    private String processingNote;

    @Column(name = "responsible_party_id")
    private Integer responsiblePartyId;

    @Column(name = "re_comp_number")
    private String reCompNumber;

    @Column(name = "result_id")
    private Integer resultId;

    @Column(name = "start_processing_date")
    private LocalDateTime startProcessingDate;

    @Column(name = "result_content", columnDefinition = "TEXT")
    private String resultContent;

    @Column(name = "cooperate_type")
    private String cooperateType;

    @Column(name = "cooperate_status")
    private String cooperateStatus;

    @Column(name = "complainer_idno")
    private String complainerIdno;

    @Column(name = "processing_user")
    private String processingUser;

    @Column(name = "lock_date")
    private LocalDateTime lockDate;

    @Column(name = "agent_note", columnDefinition = "TEXT")
    private String agentNote;

    @Column(name = "suspend_time")
    private LocalDateTime suspendTime;

    @Column(name = "cust_appoint_date")
    private LocalDateTime custAppointDate;

    @Column(name = "last_user")
    private String lastUser;

    @Column(name = "last_shop_id")
    private Integer lastShopId;

    @Column(name = "cooperate_shop_id")
    private Integer cooperateShopId;

    @Column(name = "assign_date")
    private LocalDateTime assignDate;

    @Column(name = "assign_task_status")
    private String assignTaskStatus;

    @Column(name = "isdn")
    private String isdn;

    @Column(name = "is_pass_quantity")
    private Boolean isPassQuantity;

    @Column(name = "reason_quantity")
    private String reasonQuantity;

    @Column(name = "last_cooperate_date")
    private LocalDateTime lastCooperateDate;

    @Column(name = "cooperation_end_date")
    private LocalDateTime cooperationEndDate;
}
