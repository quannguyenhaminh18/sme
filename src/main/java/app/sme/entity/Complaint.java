package app.sme.entity;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import javax.persistence.*;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;

/**
 * Entity representing the complaints table
 */
@Entity
@Table(name = "complaints")
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class Complaint {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Integer id;

    @Column(name = "complaint_no", length = 50)
    private String complaintNo;

    @Column(name = "phone_number", length = 20)
    private String phoneNumber;

    @Column(name = "complainant_name", columnDefinition = "TEXT")
    private String complainantName;

    @Column(name = "complainant_address", columnDefinition = "TEXT")
    private String complainantAddress;

    @Column(name = "contact_count")
    private Integer contactCount;

    @Column(name = "receiver_name", columnDefinition = "TEXT")
    private String receiverName;

    @Column(name = "received_date")
    private LocalDate receivedDate;

    @Column(name = "received_time")
    private LocalTime receivedTime;

    @Column(name = "complaint_group", columnDefinition = "TEXT")
    private String complaintGroup;

    @Column(name = "complaint_type", columnDefinition = "TEXT")
    private String complaintType;

    @Column(name = "category")
    private String category;

    @Column(name = "reception_method", columnDefinition = "TEXT")
    private String receptionMethod;

    @Column(name = "priority_level", length = 50)
    private String priorityLevel;

    @Column(name = "complaint_content", columnDefinition = "TEXT")
    private String complaintContent;

    @Column(name = "latest_customer_appointment")
    private LocalDate latestCustomerAppointment;

    @Column(name = "processing_deadline")
    private LocalDate processingDeadline;

    @Column(name = "processing_content", columnDefinition = "TEXT")
    private String processingContent;

    @Column(name = "processing_unit", columnDefinition = "TEXT")
    private String processingUnit;

    @Column(name = "processing_appointment_date")
    private LocalDate processingAppointmentDate;

    @Column(name = "complaint_status", columnDefinition = "TEXT")
    private String complaintStatus;

    @Column(name = "processing_result", columnDefinition = "TEXT")
    private String processingResult;

    @Column(name = "final_processor", columnDefinition = "TEXT")
    private String finalProcessor;

    @Column(name = "closing_date")
    private LocalDate closingDate;

    @Column(name = "closing_time")
    private LocalTime closingTime;

    @Column(name = "complaint_level", length = 50)
    private String complaintLevel;

    @Column(name = "arising_error", columnDefinition = "TEXT")
    private String arisingError;

    @Column(name = "customer_feedback_date")
    private LocalDate customerFeedbackDate;

    @Column(name = "customer_feedback_appointment")
    private LocalDate customerFeedbackAppointment;

    @Column(name = "customer_satisfaction_level", length = 50)
    private String customerSatisfactionLevel;

    @Column(name = "return_result", columnDefinition = "TEXT")
    private String returnResult;

    @Column(name = "return_reason", columnDefinition = "TEXT")
    private String returnReason;

    @Column(name = "duplicate_entry_id", length = 50)
    private String duplicateEntryId;

    @Column(name = "reception_source", columnDefinition = "TEXT")
    private String receptionSource;

    @Column(columnDefinition = "TEXT")
    private String district;

    @Column(columnDefinition = "TEXT")
    private String ward;

    @Column(name = "reception_channel", columnDefinition = "TEXT")
    private String receptionChannel;

    @Column(name = "package_name", columnDefinition = "TEXT")
    private String packageName;

    @Column(name = "total_processing_time_hours", precision = 10, scale = 2)
    private BigDecimal totalProcessingTimeHours;

    @Column(name = "dslam_dlu", columnDefinition = "TEXT")
    private String dslamDlu;

    @Column(name = "level1_error_reason", columnDefinition = "TEXT")
    private String level1ErrorReason;

    @Column(name = "level2_error_reason", columnDefinition = "TEXT")
    private String level2ErrorReason;

    @Column(name = "level3_error_reason", columnDefinition = "TEXT")
    private String level3ErrorReason;

    @Column(name = "contract_number", length = 100)
    private String contractNumber;

    @Column(columnDefinition = "TEXT")
    private String notes;

    @Column(columnDefinition = "TEXT")
    private String assigner;

    @Column(columnDefinition = "TEXT")
    private String assignee;

    @Column(name = "progress_status", columnDefinition = "TEXT")
    private String progressStatus;

    @Column(name = "overdue_time_hours", precision = 10, scale = 2)
    private BigDecimal overdueTimeHours;

    @Column(name = "preliminary_processing_content", columnDefinition = "TEXT")
    private String preliminaryProcessingContent;

    @Column(name = "contact_number", length = 20)
    private String contactNumber;

    @Column(name = "additional_info", columnDefinition = "TEXT")
    private String additionalInfo;

    @Column(name = "sms_1715_content", columnDefinition = "TEXT")
    private String sms1715Content;

    @Column(name = "weak_point_code", columnDefinition = "TEXT")
    private String weakPointCode;

    @Column(name = "responsible_unit", columnDefinition = "TEXT")
    private String responsibleUnit;

    @Column(name = "local_staff_id", length = 50)
    private String localStaffId;

    @Column(name = "customer_type", columnDefinition = "TEXT")
    private String customerType;

    @Column(name = "customer_type_detail", columnDefinition = "TEXT")
    private String customerTypeDetail;

    @Column(name = "subscriber_status", columnDefinition = "TEXT")
    private String subscriberStatus;

    @Column(name = "subscriber_state", columnDefinition = "TEXT")
    private String subscriberState;

    @Column(name = "overdue_error_reason_lvl1", columnDefinition = "TEXT")
    private String overdueErrorReasonLvl1;

    @Column(name = "overdue_error_reason_lvl2", columnDefinition = "TEXT")
    private String overdueErrorReasonLvl2;

    @Column(name = "overdue_responsible_department", columnDefinition = "TEXT")
    private String overdueResponsibleDepartment;

    @Column(name = "ntms_processing_unit", columnDefinition = "TEXT")
    private String ntmsProcessingUnit;

    @Column(name = "ntms_processor", columnDefinition = "TEXT")
    private String ntmsProcessor;

    @Column(name = "appointment_count")
    private Integer appointmentCount;

    @Column(name = "appointment_success_count")
    private Integer appointmentSuccessCount;

    @Column(name = "appointment_fail_count")
    private Integer appointmentFailCount;

    @Column(name = "appointment_not_evaluated_count")
    private Integer appointmentNotEvaluatedCount;

    @Column(name = "appointment_user", columnDefinition = "TEXT")
    private String appointmentUser;

    @Column(name = "appointment_evaluation_result", columnDefinition = "TEXT")
    private String appointmentEvaluationResult;

    @Column(name = "aon_gpon", length = 50)
    private String aonGpon;

    @Column(name = "cooperation_unit_type", columnDefinition = "TEXT")
    private String cooperationUnitType;

    @Column(name = "note_content", columnDefinition = "TEXT")
    private String noteContent;

    @Column(name = "subscriber_code", length = 50)
    private String subscriberCode;

    @Column(name = "total_incident_processing_time_gnoc", precision = 10, scale = 2)
    private BigDecimal totalIncidentProcessingTimeGnoc;

    @Column(name = "customer_satisfaction_lvl2", length = 50)
    private String customerSatisfactionLvl2;

    @Column(name = "evaluation_when_receiving", columnDefinition = "TEXT")
    private String evaluationWhenReceiving;

    @Column(name = "ntms_expected_completion_time")
    private LocalDateTime ntmsExpectedCompletionTime;

    @Column(name = "sms_1715_message", columnDefinition = "TEXT")
    private String sms1715Message;

    @Column(name = "cooperation_status", columnDefinition = "TEXT")
    private String cooperationStatus;

    @Column(name = "receiving_department", columnDefinition = "TEXT")
    private String receivingDepartment;

    @Column(name = "cooperation_unit", columnDefinition = "TEXT")
    private String cooperationUnit;

    @Column(name = "appointment_processing_time_hours", precision = 10, scale = 2)
    private BigDecimal appointmentProcessingTimeHours;

    @Column(name = "cooperation_result", columnDefinition = "TEXT")
    private String cooperationResult;

    @Column(name = "cooperation_failure_reason", columnDefinition = "TEXT")
    private String cooperationFailureReason;

    @Column(name = "total_incident_processing_time_spm", precision = 10, scale = 2)
    private BigDecimal totalIncidentProcessingTimeSpm;

    @Column(name = "gnoc_processing_progress", columnDefinition = "TEXT")
    private String gnocProcessingProgress;

    @Column(name = "gnoc_total_appointment_time", precision = 10, scale = 2)
    private BigDecimal gnocTotalAppointmentTime;

    @Column(name = "incident_completion_date")
    private LocalDate incidentCompletionDate;

    @Column(name = "incident_completion_time")
    private LocalTime incidentCompletionTime;

    @Column(name = "error_code", columnDefinition = "TEXT")
    private String errorCode;

    @Column(name = "ft_cd_rejection", columnDefinition = "TEXT")
    private String ftCdRejection;

    @Column(name = "vip_customer_flag")
    private Boolean vipCustomerFlag;

    @Column(name = "critical_point_flag")
    private Boolean criticalPointFlag;

    @Column(name = "customer_classification", columnDefinition = "TEXT")
    private String customerClassification;

    @Column(name = "subscriber_number", length = 50)
    private String subscriberNumber;

    @Column(columnDefinition = "TEXT")
    private String segment;

    @Column(name = "payment_method", columnDefinition = "TEXT")
    private String paymentMethod;

    @Column(name = "arpu_info", columnDefinition = "TEXT")
    private String arpuInfo;

    @Column(name = "hot_pick_time")
    private LocalDateTime hotPickTime;

    @Column(name = "work_status", columnDefinition = "TEXT")
    private String workStatus;

    @Column(name = "error_description", columnDefinition = "TEXT")
    private String errorDescription;

    @Column(columnDefinition = "TEXT")
    private String cause;

    @Column(columnDefinition = "TEXT")
    private String solution;

    @Column(name = "detailed_result", columnDefinition = "TEXT")
    private String detailedResult;

    @Column(name = "response_time")
    private LocalDateTime responseTime;

    @Column(name = "ticket_code", length = 100)
    private String ticketCode;

    @Column(name = "wo_code", length = 50)
    private String woCode;

    @Column(name = "cell_id_gnoc", length = 50)
    private String cellIdGnoc;

    @Column(name = "processing_method", columnDefinition = "TEXT")
    private String processingMethod;

    @Column(name = "g_line", length = 50)
    private String gLine;

    @Column(name = "subscriber_cable_length", precision = 10, scale = 2)
    private BigDecimal subscriberCableLength;

    @Column(name = "project_code", length = 50)
    private String projectCode;

    @Column(name = "project_name", columnDefinition = "TEXT")
    private String projectName;

    @Column(name = "project_address", columnDefinition = "TEXT")
    private String projectAddress;

    @Column(name = "ticket_close_source", columnDefinition = "TEXT")
    private String ticketCloseSource;

    @Column(name = "cancel_reason", columnDefinition = "TEXT")
    private String cancelReason;

    @Column(name = "cancel_request_subscriber", columnDefinition = "TEXT")
    private String cancelRequestSubscriber;

    @Column(name = "total_actual_processing_time_hours", precision = 10, scale = 2)
    private BigDecimal totalActualProcessingTimeHours;

    @Column(name = "area_type", columnDefinition = "TEXT")
    private String areaType;

    @Column(name = "branch_explanation", columnDefinition = "TEXT")
    private String branchExplanation;

    @Column(name = "transaction_store", columnDefinition = "TEXT")
    private String transactionStore;

    @Column(columnDefinition = "TEXT")
    private String hamlet;

    @Column(name = "gnoc_first_completion_time")
    private LocalDateTime gnocFirstCompletionTime;

    @Column(name = "feedback_deadline")
    private LocalDate feedbackDeadline;

    @Column(name = "error_transfer_date")
    private LocalDate errorTransferDate;

    @Column(name = "technical_feedback_content", columnDefinition = "TEXT")
    private String technicalFeedbackContent;

    @Column(name = "customer_care_feedback_status", columnDefinition = "TEXT")
    private String customerCareFeedbackStatus;

    @Column(name = "penalty_amount", precision = 10, scale = 2)
    private BigDecimal penaltyAmount;

    @Column(name = "violating_staff_name", columnDefinition = "TEXT")
    private String violatingStaffName;

    @Column(name = "violating_staff_card", length = 50)
    private String violatingStaffCard;

    @Column(length = 255)
    private String email;

    @Column(name = "phone_contact", length = 20)
    private String phoneContact;

    @Column(name = "technical_feedback_receiver", columnDefinition = "TEXT")
    private String technicalFeedbackReceiver;

    @Column(columnDefinition = "TEXT")
    private String region;

    @Column(name = "happy_call_staff_name", columnDefinition = "TEXT")
    private String happyCallStaffName;

    @Column(name = "hpc_result", columnDefinition = "TEXT")
    private String hpcResult;

    @Column(name = "hpc_result_details", columnDefinition = "TEXT")
    private String hpcResultDetails;

    @Column(name = "contact_phone_number", length = 20)
    private String contactPhoneNumber;

    @Column(name = "solution_code", length = 50)
    private String solutionCode;
}