package app.sme;

public interface SMEProjection {
    String getCategory();
    Integer getDayOfMonth();
    Long getCountYesterday();
    Long getTotalReceivedYesterday();
    Long getTotalOnTimeYesterday();
    Double getAvgProcessingTimeHours();
    Long getCooperationUnitCount();
    Long getTotalLast8Days();
    Long getTotalLast8Months();
    Long getTotalCurrentYear();
    Long getTotalLastYear();
    Long getTotalComplaintsPerDay();
    Long getTotalClosedPerDay();
    Long getTotalOnTimePerDay();
}