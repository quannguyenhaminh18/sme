package app.sme.service_quality;

import java.math.BigDecimal;

public interface IComplaintServiceQuality {
    String getName();

    BigDecimal getKpi();

    Long getTotalSubscriber();

    Long getTotalComplaintDay();

    BigDecimal getComplaintRateDay();

    Long getTotalComplaintMonth();

    BigDecimal getComplaintRateMonth();

    Long getTotalLastMonth();

    BigDecimal getComplaintRateLastMonth();

    String getKpiDay();

    String getKpiMonth();

    BigDecimal getCompareRate();

}
