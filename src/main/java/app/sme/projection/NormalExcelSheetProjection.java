package app.sme.projection;

public interface NormalExcelSheetProjection {
    String getCategory();
    Integer getDayOfMonth();
    Long getCountYesterday();
    Long getTotalReceivedYesterday();
    Long getTotalOnTimeYesterday();
    Long getCooperationUnitCount();
    Long getTotalComplaintsPerDay();
    Long getTotalClosedPerDay();
    Long getTotalOnTimePerDay();
}