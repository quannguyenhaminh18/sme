package app.sme;

import lombok.*;
import lombok.experimental.FieldDefaults;

import java.math.BigDecimal;

@Getter
@Setter
@FieldDefaults(level = AccessLevel.PRIVATE)
@AllArgsConstructor
@NoArgsConstructor
@Builder

public class DocReportDTO {
    String name;
    BigDecimal kpi;
    Long totalSubscriber;

    Long totalComplaintDay;
    BigDecimal complaintRateDay;

    Long totalComplaintMonth;
    BigDecimal complaintRateMonth;

    Long totalLastMonth;
    BigDecimal complaintRateLastMonth;

    String kpiDay;
    String kpiMonth;
    BigDecimal compareRate;
}
