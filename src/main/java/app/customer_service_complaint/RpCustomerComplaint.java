package app.customer_service_complaint;

import lombok.Getter;
import lombok.Setter;

import javax.persistence.*;
import java.math.BigDecimal;
import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalTime;

@Getter
@Setter
@Entity
@Table(name = "rp_customer_complaints", schema = "cos_report")
public class RpCustomerComplaint {
    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    @Column(name = "id", nullable = false)
    private Long id;

    @Column(name = "complaint_code", length = 50)
    private String complaintCode;

    @Column(name = "phone_number", length = 50)
    private String phoneNumber;

    @Column(name = "contact_count")
    private Integer contactCount;

    @Column(name = "complaint_receiver", length = 100)
    private String complaintReceiver;

    @Column(name = "received_date")
    private LocalDate receivedDate;

    @Column(name = "received_time")
    private LocalTime receivedTime;

    @Column(name = "complaint_group", length = 100)
    private String complaintGroup;

    @Column(name = "category", length = 200)
    private String category;

    @Column(name = "complaint_type", length = 200)
    private String complaintType;

    @Column(name = "received_method", length = 100)
    private String receivedMethod;

    @Column(name = "priority_level", length = 50)
    private String priorityLevel;

    @Lob
    @Column(name = "complaint_content")
    private String complaintContent;

    @Column(name = "customer_appointment_date")
    private LocalDate customerAppointmentDate;

    @Column(name = "processing_deadline")
    private LocalDate processingDeadline;

    @Lob
    @Column(name = "processing_content")
    private String processingContent;

    @Column(name = "processing_unit", length = 100)
    private String processingUnit;

    @Column(name = "processing_appointment_date")
    private LocalDate processingAppointmentDate;

    @Column(name = "complaint_status", length = 100)
    private String complaintStatus;

    @Column(name = "processing_result", length = 200)
    private String processingResult;

    @Column(name = "last_processor", length = 100)
    private String lastProcessor;

    @Column(name = "close_date")
    private LocalDate closeDate;

    @Column(name = "close_time")
    private LocalTime closeTime;

    @Column(name = "complaint_level", length = 50)
    private String complaintLevel;

    @Column(name = "error_generated", length = 200)
    private String errorGenerated;

    @Column(name = "customer_feedback_date")
    private LocalDate customerFeedbackDate;

    @Column(name = "customer_response_deadline")
    private LocalDate customerResponseDeadline;

    @Column(name = "customer_satisfaction_level", length = 100)
    private String customerSatisfactionLevel;

    @Column(name = "return_result", length = 200)
    private String returnResult;

    @Column(name = "return_reason", length = 200)
    private String returnReason;

    @Column(name = "duplicate_entry_id", length = 50)
    private String duplicateEntryId;

    @Column(name = "received_source", length = 100)
    private String receivedSource;

    @Column(name = "district", length = 100)
    private String district;

    @Column(name = "ward", length = 100)
    private String ward;

    @Column(name = "receiving_channel", length = 100)
    private String receivingChannel;

    @Column(name = "package_name", length = 100)
    private String packageName;

    @Column(name = "total_processing_time_hours", precision = 10, scale = 2)
    private BigDecimal totalProcessingTimeHours;

    @Column(name = "dslam_dlu", length = 100)
    private String dslamDlu;

    @Column(name = "error_cause_lvl1", length = 200)
    private String errorCauseLvl1;

    @Column(name = "error_cause_lvl2", length = 200)
    private String errorCauseLvl2;

    @Column(name = "error_cause_lvl3", length = 200)
    private String errorCauseLvl3;

    @Column(name = "error_cause_lvl4", length = 200)
    private String errorCauseLvl4;

    @Column(name = "contract_number", length = 100)
    private String contractNumber;

    @Lob
    @Column(name = "notes")
    private String notes;

    @Column(name = "assigner", length = 100)
    private String assigner;

    @Column(name = "assignee", length = 100)
    private String assignee;

    @Column(name = "progress", length = 100)
    private String progress;

    @Column(name = "overdue_time_hours", precision = 10, scale = 2)
    private BigDecimal overdueTimeHours;

    @Lob
    @Column(name = "preliminary_processing_content")
    private String preliminaryProcessingContent;

    @Column(name = "contact_number", length = 50)
    private String contactNumber;

    @Lob
    @Column(name = "additional_info")
    private String additionalInfo;

    @Lob
    @Column(name = "sms_1715")
    private String sms1715;

    @Column(name = "weak_point_code", length = 50)
    private String weakPointCode;

    @Column(name = "responsible_unit", length = 100)
    private String responsibleUnit;

    @Column(name = "field_staff_code", length = 50)
    private String fieldStaffCode;

    @Column(name = "service_type", length = 100)
    private String serviceType;

    @Column(name = "customer_object", length = 100)
    private String customerObject;

    @Column(name = "customer_object_detail", length = 200)
    private String customerObjectDetail;

    @Column(name = "subscriber_status", length = 100)
    private String subscriberStatus;

    @Column(name = "service_status", length = 100)
    private String serviceStatus;

    @Column(name = "overdue_cause_lvl1", length = 200)
    private String overdueCauseLvl1;

    @Column(name = "overdue_cause_lvl2", length = 200)
    private String overdueCauseLvl2;

    @Column(name = "overdue_department", length = 100)
    private String overdueDepartment;

    @Column(name = "ntms_processing_unit", length = 100)
    private String ntmsProcessingUnit;

    @Column(name = "ntms_processor", length = 100)
    private String ntmsProcessor;

    @Column(name = "appointment_count")
    private Integer appointmentCount;

    @Column(name = "successful_appointment_count")
    private Integer successfulAppointmentCount;

    @Column(name = "failed_appointment_count")
    private Integer failedAppointmentCount;

    @Column(name = "unevaluated_appointment_count")
    private Integer unevaluatedAppointmentCount;

    @Column(name = "appointment_user", length = 100)
    private String appointmentUser;

    @Column(name = "appointment_evaluation_result", length = 200)
    private String appointmentEvaluationResult;

    @Column(name = "aon_gpon", length = 50)
    private String aonGpon;

    @Column(name = "cooperating_unit_type", length = 100)
    private String cooperatingUnitType;

    @Lob
    @Column(name = "notes_detail")
    private String notesDetail;

    @Column(name = "subscriber_code", length = 50)
    private String subscriberCode;

    @Column(name = "total_gnoc_processing_time", precision = 10, scale = 2)
    private BigDecimal totalGnocProcessingTime;

    @Column(name = "satisfaction_level2", length = 100)
    private String satisfactionLevel2;

    @Column(name = "complaint_handling_evaluation", length = 200)
    private String complaintHandlingEvaluation;

    @Column(name = "ntms_estimated_completion_time")
    private Instant ntmsEstimatedCompletionTime;

    @Lob
    @Column(name = "sms_1715_content")
    private String sms1715Content;

    @Column(name = "cooperation_status", length = 100)
    private String cooperationStatus;

    @Column(name = "receiving_department", length = 100)
    private String receivingDepartment;

    @Column(name = "cooperating_unit", length = 100)
    private String cooperatingUnit;

    @Column(name = "technician_appointment_time")
    private Instant technicianAppointmentTime;

    @Column(name = "cooperation_result", length = 200)
    private String cooperationResult;

    @Column(name = "cooperation_fail_reason", length = 200)
    private String cooperationFailReason;

    @Column(name = "total_spm_processing_time", precision = 10, scale = 2)
    private BigDecimal totalSpmProcessingTime;

    @Column(name = "gnoc_processing_progress", length = 100)
    private String gnocProcessingProgress;

    @Column(name = "total_gnoc_appointment_time", precision = 10, scale = 2)
    private BigDecimal totalGnocAppointmentTime;

    @Column(name = "issue_completion_date")
    private LocalDate issueCompletionDate;

    @Column(name = "issue_completion_time")
    private LocalTime issueCompletionTime;

    @Column(name = "error_code", length = 50)
    private String errorCode;

    @Column(name = "ft_cd_reject_flag")
    private Boolean ftCdRejectFlag;

    @Column(name = "vip_customer_flag")
    private Boolean vipCustomerFlag;

    @Column(name = "critical_point_flag")
    private Boolean criticalPointFlag;

    @Column(name = "customer_classification", length = 100)
    private String customerClassification;

    @Column(name = "subscriber", length = 100)
    private String subscriber;

    @Column(name = "segment", length = 100)
    private String segment;

    @Column(name = "payment_method", length = 100)
    private String paymentMethod;

    @Column(name = "arpu_info", length = 100)
    private String arpuInfo;

    @Column(name = "actual_processing_time_hours", precision = 10, scale = 2)
    private BigDecimal actualProcessingTimeHours;

    @Column(name = "pa_increase_time", precision = 10, scale = 2)
    private BigDecimal paIncreaseTime;

    @Column(name = "province_new", length = 100)
    private String provinceNew;

    @Column(name = "commune_new", length = 100)
    private String communeNew;

}