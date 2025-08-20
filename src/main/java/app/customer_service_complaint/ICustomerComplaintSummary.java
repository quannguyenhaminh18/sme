package app.customer_service_complaint;

import java.math.BigDecimal;

public interface ICustomerComplaintSummary {

    String getParentName();

    String getServiceName();

    String getCategoryName();

    Long getLuyKe();

    BigDecimal getTrungBinhNgay();

    BigDecimal getTiTrong();
}