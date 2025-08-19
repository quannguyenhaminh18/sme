package app.sme;

import app.sme.service_quality.ComplainStatisticRepository;
import app.sme.service_quality.ComplaintServiceQualityReportDto;
import app.sme.service_quality.ComplaintsStatistic;
import app.sme.service_quality.IComplaintServiceQuality;
import lombok.RequiredArgsConstructor;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

@Service
@RequiredArgsConstructor
public class ComplaintService {
    private final ComplaintRepository complaintRepository;
    private final ComplainStatisticRepository complainStatisticRepository;

    private Sheet getRequiredSheet(Workbook workbook, String sheetName) {
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            throw new IllegalArgumentException("Không tìm thấy sheet " + sheetName);
        }
        return sheet;
    }

    public byte[] exportExcelFile() {
        try (InputStream inputStream = getClass().getResourceAsStream("/templates/SME_Phu_luc.xlsx")) {
            assert inputStream != null;
            try (Workbook workbook = WorkbookFactory.create(inputStream)) {
                Sheet sheet = getRequiredSheet(workbook, "SL_ngay");

                // Lấy danh sách category từ cột B8 đến B39
                int startRow = 7; // B8
                int endRow = 38;  // B39
                int categoryCol = 1; // Cột B
                java.util.List<String> categories = new java.util.ArrayList<>();
                for (int i = startRow; i <= endRow; i++) {
                    Row row = sheet.getRow(i);
                    if (row != null) {
                        Cell cell = row.getCell(categoryCol);
                        if (cell != null && cell.getCellType() == CellType.STRING) {
                            categories.add(cell.getStringCellValue().trim());
                        } else {
                            categories.add("");
                        }
                    } else {
                        categories.add("");
                    }
                }

                // TEST: Override ngày hệ thống thành 2-8-2025 để kiểm tra logic xuất file
                LocalDate today = LocalDate.of(2025, 8, 3);
                LocalDate endDate = today.minusDays(1);
                int currentDay = endDate.getDayOfMonth();
                LocalDate firstDay = today.withDayOfMonth(1);
                java.sql.Date sqlStart = java.sql.Date.valueOf(firstDay);
                java.sql.Date sqlEnd = java.sql.Date.valueOf(endDate);

                ByteArrayOutputStream out = new ByteArrayOutputStream();
                workbook.write(out);
                return out.toByteArray();
            }
        } catch (IOException e) {
            throw new RuntimeException("Lỗi đọc/ghi file Excel", e);
        }
    }

    public byte[] exportServiceQualityDoc() {
        //Data table
        List<IComplaintServiceQuality> dataTable = complaintRepository.reportServiceQuality();
        List<ComplaintServiceQualityReportDto> convertList = dataTable.stream()
                .map(ComplaintServiceQualityReportDto::new)
                .collect(Collectors.toList());

        //Data chart
        List<ComplaintsStatistic> chartData = complainStatisticRepository.findAll();

        try (InputStream is = new ClassPathResource("templates/doc/bao_cao_ngay_cldv_sme.docx").getInputStream();
             OPCPackage pkg = OPCPackage.open(is);
             XWPFDocument doc = new XWPFDocument(pkg);
             ByteArrayOutputStream out = new ByteArrayOutputStream()) {

            var now = LocalDate.now();
            var date = now.getDayOfMonth();
            var month = now.getMonthValue();
            var year = now.getYear();
            var reportDate = now.format(DateTimeFormatter.ofPattern("dd/MM/yyyy"));

            Map<String, String> map = new HashMap<>();

            map.put("${date}", String.valueOf(date));
            map.put("${month}", String.valueOf(month));
            map.put("${year}", String.valueOf(year));
            map.put("${reportDate}", reportDate);

            replaceTextInDocument(doc, map);


            doc.write(out);
            return out.toByteArray();


        } catch (Exception e) {
            throw new RuntimeException("Lỗi fill template DOCX", e);
        }

    }

    private void replaceTextInDocument(XWPFDocument doc, Map<String, String> map) {
        // Paragraphs
        for (XWPFParagraph p : doc.getParagraphs()) {
            replaceTextInParagraph(p, map);
        }

        // Tables
        for (XWPFTable t : doc.getTables()) {
            for (XWPFTableRow r : t.getRows()) {
                for (XWPFTableCell c : r.getTableCells()) {
                    for (XWPFParagraph p : c.getParagraphs()) {
                        replaceTextInParagraph(p, map);
                    }
                }
            }

        }

        // Headers/Footers (nếu có)
        for (XWPFHeader header : doc.getHeaderList()) {
            for (XWPFParagraph p : header.getParagraphs()) replaceTextInParagraph(p, map);
            for (XWPFTable t : header.getTables()) {
                for (XWPFTableRow r : t.getRows())
                    for (XWPFTableCell c : r.getTableCells())
                        for (XWPFParagraph p : c.getParagraphs()) replaceTextInParagraph(p, map);
            }
        }

        for (XWPFFooter footer : doc.getFooterList()) {
            for (XWPFParagraph p : footer.getParagraphs()) replaceTextInParagraph(p, map);
        }
    }


    private void replaceTextInParagraph(XWPFParagraph paragraph, Map<String, String> map) {
        StringBuilder sb = new StringBuilder();
        for (XWPFRun run : paragraph.getRuns()) {
            sb.append(run.text());
        }
        String combined = sb.toString();
        boolean changed = false;
        for (Map.Entry<String, String> e : map.entrySet()) {
            if (combined.contains(e.getKey())) {
                combined = combined.replace(e.getKey(), e.getValue());
                changed = true;
            }
        }
        if (changed) {
            // clear old runs and set a single new run (giữ định dạng cơ bản)
            int runCount = paragraph.getRuns().size();
            for (int i = runCount - 1; i >= 0; i--) {
                paragraph.removeRun(i);
            }
            XWPFRun newRun = paragraph.createRun();
            newRun.setText(combined, 0);
        }
    }

}
