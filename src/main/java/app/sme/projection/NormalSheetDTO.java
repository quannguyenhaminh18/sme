package app.sme.projection;

public interface NormalSheetDTO {
    String getCategory();
    Long getTotalReceivedYesterday();
    Long getTotalErrorReceivedYesterday();
    Long getTotalErrorClosedYesterday();
    Long getTotalErrorClosedOntimeYesterday();
    Long getTotalCooperationUnitMonth();
    Long getTotalProcessingTimeMonth();
}
