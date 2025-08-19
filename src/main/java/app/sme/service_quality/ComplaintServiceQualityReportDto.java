package app.sme.service_quality;

import lombok.*;
import lombok.experimental.FieldDefaults;

import java.math.BigDecimal;

@Getter
@Setter
@NoArgsConstructor
@FieldDefaults(level = AccessLevel.PRIVATE)
public class ComplaintServiceQualityReportDto {
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

    public ComplaintServiceQualityReportDto(IComplaintServiceQuality result) {
        this.name = result.getName();
        this.kpi = result.getKpi();
        this.totalSubscriber = result.getTotalSubscriber();
        this.totalComplaintDay = result.getTotalComplaintDay();
        this.totalComplaintMonth = result.getTotalComplaintMonth();
        this.totalLastMonth = result.getTotalLastMonth();
        this.kpiDay = result.getKpiDay();
        this.kpiMonth = result.getKpiMonth();
        this.compareRate = result.getCompareRate();
        this.complaintRateDay = result.getComplaintRateDay();
        this.complaintRateMonth = result.getComplaintRateMonth();
        this.complaintRateLastMonth = result.getComplaintRateLastMonth();
    }

}
