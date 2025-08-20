package app.sme;

public interface SMEChartProjection {
    String getCategory();
    Long getTotalComplaintsLastMonth();
    String getDimensionType();      // ví dụ: "MONTH", "DAY", "AVG_THIS_YEAR", ...
    String getDimensionValue();     // ví dụ: "2025-01", "2025-07-15"
    Long getTotalComplaints();
}
