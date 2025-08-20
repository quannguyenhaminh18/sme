package app.sme;

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
import java.time.YearMonth;
import java.time.format.DateTimeFormatter;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

@Service
@RequiredArgsConstructor
public class ComplaintService {
    private final ComplaintRepository complaintRepository;
    private final CategoryRepository categoryRepository;

    private Sheet getRequiredSheet(Workbook workbook, String sheetName) {
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            throw new IllegalArgumentException("Không tìm thấy sheet " + sheetName);
        }
        return sheet;
    }

    public byte[] exportExcelFile() {
        try (InputStream inputStream = getClass().getResourceAsStream("/templates/SME_Phu_luc.xlsx")) {
            if (inputStream == null) {
                throw new IllegalArgumentException("Không tìm thấy file template SME_Phu_luc.xlsx");
            }

            try (Workbook workbook = WorkbookFactory.create(inputStream);
                 ByteArrayOutputStream out = new ByteArrayOutputStream()) {

                Sheet sheet = getRequiredSheet(workbook, "SL_ngay");

                // --- Lấy dữ liệu từ DB ---
                Map<String, Long> categoryTotals = categoryRepository.findAll().stream()
                        .collect(Collectors.toMap(Category::getName, Category::getTotalSubscriber));

                List<SLNgayProjection> summaries = complaintRepository.findCountsYesterday();

                Map<String, Long> yesterdayCounts = summaries.stream()
                        .collect(Collectors.toMap(
                                SLNgayProjection::getCategory,
                                s -> s.getCountYesterday() != null ? s.getCountYesterday().longValue() : 0L
                        ));

                Map<String, Long> totalThisMonthUntilYesterday = summaries.stream()
                        .collect(Collectors.toMap(
                                SLNgayProjection::getCategory,
                                s -> s.getCountThisMonth() != null ?
                                        s.getCountThisMonth().longValue() : 0L
                        ));


                // --- Fill dữ liệu C7–C38 ---
                for (int rowIndex = 6; rowIndex <= 37; rowIndex++) {
                    Row row = sheet.getRow(rowIndex);
                    if (row == null) continue;

                    Cell nameCell = row.getCell(1); // cột B
                    if (nameCell == null) continue;

                    String categoryName = nameCell.getStringCellValue();
                    if (categoryName == null || categoryName.isBlank()) continue;

                    // điền total_subscriber vào cột C
                    if (categoryTotals.containsKey(categoryName)) {
                        Cell subscriberCell = row.getCell(2);
                        if (subscriberCell == null) {
                            subscriberCell = row.createCell(2);
                        }
                        subscriberCell.setCellValue(categoryTotals.get(categoryName));
                    }

                    // điền total_until_yesterday vào cột D
                    if (totalThisMonthUntilYesterday.containsKey(categoryName)) {
                        Cell dCell = row.getCell(3); // cột D = index 3
                        if (dCell == null) {
                            dCell = row.createCell(3);
                        }
                        dCell.setCellValue(totalThisMonthUntilYesterday.get(categoryName));
                    }

                    // điền count_yesterday vào đúng cột ngày hôm qua
                    LocalDate yesterday = LocalDate.now().minusDays(1);
                    int colIndex = 7 + yesterday.getDayOfMonth() - 1; // H = index 7
                    if (yesterdayCounts.containsKey(categoryName)) {
                        Cell dataCell = row.getCell(colIndex);
                        if (dataCell == null) {
                            dataCell = row.createCell(colIndex);
                        }
                        dataCell.setCellValue(yesterdayCounts.get(categoryName));
                    }
                }

                // --- Update header H4 + C4 ---
                LocalDate today = LocalDate.now();
                int month = today.getMonthValue();
                int year = today.getYear();
                YearMonth yearMonth = YearMonth.of(year, month);
                int daysInMonth = yearMonth.lengthOfMonth();

                Row row4 = sheet.getRow(3);
                if (row4 == null) {
                    throw new IllegalStateException("Không tìm thấy row 4 trong sheet");
                }

                Cell h4Cell = row4.getCell(7);
                if (h4Cell == null) {
                    System.out.println("⚠️ H4 chưa tồn tại trong file template");
                    h4Cell = row4.createCell(7);
                }
                h4Cell.setCellValue("T" + month + "." + year);

                Cell c4Cell = row4.getCell(2);
                if (c4Cell == null) {
                    System.out.println("⚠️ C4 chưa tồn tại trong file template");
                    c4Cell = row4.createCell(2);
                }
                c4Cell.setCellValue("Luỹ kế T" + month + "." + year);


                // --- Update header ngày (H5–AL5) ---
                Row row5 = sheet.getRow(4);
                DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd/MM");
                int startCol = 7; // H
                int maxCol = 37;  // AL

                CellStyle baseStyle = null;
                Cell baseCell = row5.getCell(startCol);
                if (baseCell != null) baseStyle = baseCell.getCellStyle();

                for (int day = 1; day <= daysInMonth; day++) {
                    LocalDate date = LocalDate.of(year, month, day);
                    int colIndex = startCol + day - 1;
                    Cell cell = row5.getCell(colIndex);
                    if (cell == null) {
                        cell = row5.createCell(colIndex);
                        if (baseStyle != null) cell.setCellStyle(baseStyle);
                    }
                    cell.setCellValue(date.format(formatter));
                }

                Sheet sheet2 = this.getRequiredSheet(workbook, "KPI_ngay");
                Map<String, Integer> serviceRowMap = new HashMap<>();
                serviceRowMap.put("Dịch vụ CA", 7);
                serviceRowMap.put("Dịch vụ BHXH", 12);
                serviceRowMap.put("Dịch vụ Hóa đơn điện tử", 17);
                serviceRowMap.put("Dịch vụ vTracking", 22);
                Map<String, int[]> serviceTotalRowMap = new HashMap<>();
                serviceTotalRowMap.put("Dịch vụ CA", new int[]{28, 29});
                serviceTotalRowMap.put("Dịch vụ BHXH", new int[]{33, 34});
                serviceTotalRowMap.put("Dịch vụ Hóa đơn điện tử", new int[]{38, 39});
                serviceTotalRowMap.put("Dịch vụ vTracking", new int[]{43, 44});
                Map<String, Integer> normalizedServiceRowMap = new HashMap<>();

                for (Map.Entry<String, Integer> e : serviceRowMap.entrySet()) {
                    normalizedServiceRowMap.put(e.getKey().toLowerCase(), (Integer) e.getValue());
                }

                Map<String, int[]> normalizedServiceTotalRowMap = new HashMap<>();

                for (Map.Entry<String, int[]> e : serviceTotalRowMap.entrySet()) {
                    normalizedServiceTotalRowMap.put(e.getKey().toLowerCase(), e.getValue());
                }

                List<Object[]> stats = this.complaintRepository.findYesterdayCounts();
                Map<String, Integer> countMap = new HashMap<>();

                for (Object[] row : stats) {
                    if (row[0] != null && row[1] != null) {
                        String service = row[0].toString().trim().toLowerCase();
                        int countYesterday = ((Number) row[1]).intValue();
                        countMap.put(service, countYesterday);
                    }
                }

                for (Map.Entry<String, Integer> entry : normalizedServiceRowMap.entrySet()) {
                    String serviceName = entry.getKey();
                    int totalRowIdx2 = entry.getValue();
                    Row row = sheet2.getRow(totalRowIdx2);
                    if (row == null) {
                        row = sheet2.createRow(totalRowIdx2);
                    }

                    int todayInt = LocalDate.now().getDayOfMonth();
                    int colIdx = 4 + (todayInt - 2);
                    Cell cell = row.getCell(colIdx, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    double oldValue = cell.getCellType() == CellType.NUMERIC ? cell.getNumericCellValue() : (double) 0.0F;
                    int countYesterday = countMap.getOrDefault(serviceName, 0);
                    cell.setCellValue(oldValue + (double) countYesterday);
                }

                List<Object[]> totalStats = this.complaintRepository.findYesterdayOnTimeAndReceivedCounts();
                Map<String, Integer[]> totalMap = new HashMap<>();

                for (Object[] row : totalStats) {
                    if (row[0] != null && row[1] != null && row[2] != null) {
                        String service = row[0].toString().trim().toLowerCase();
                        int totalOnTime = ((Number) row[1]).intValue();
                        int totalReceived = ((Number) row[2]).intValue();
                        totalMap.put(service, new Integer[]{totalOnTime, totalReceived});
                    }
                }

                for (Map.Entry<String, int[]> entry : normalizedServiceTotalRowMap.entrySet()) {
                    String serviceName = entry.getKey();
                    int[] rows = entry.getValue();
                    Integer[] values = totalMap.getOrDefault(serviceName, new Integer[]{0, 0});
                    int totalOnTime = values[0];
                    int totalReceived = values[1];
                    int todayInt = LocalDate.now().getDayOfMonth();
                    int colIdx = 4 + (todayInt - 2);
                    Row rowOnTime = sheet2.getRow(rows[0]);
                    if (rowOnTime == null) {
                        rowOnTime = sheet2.createRow(rows[0]);
                    }

                    Cell cellOnTime = rowOnTime.getCell(colIdx, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    cellOnTime.setCellValue(totalOnTime);
                    Row rowReceived = sheet2.getRow(rows[1]);
                    if (rowReceived == null) {
                        rowReceived = sheet2.createRow(rows[1]);
                    }

                    Cell cellReceived = rowReceived.getCell(colIdx, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    cellReceived.setCellValue(totalReceived);
                }
                ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
                workbook.write(outputStream);
                return outputStream.toByteArray();
            }
        } catch (IOException e) {
            throw new RuntimeException("Lỗi đọc/ghi file Excel", e);
        }
    }

    public byte[] exportDoc() {
        //Data table
        List<DocxProjection> dataTable = complaintRepository.reportServiceQuality();
//        List<DocReportDTO> convertList = dataTable.stream()
//                .map(DocReportDTO::new)
//                .collect(Collectors.toList());

        //Data chart
//        List<ComplaintsStatistic> chartData = complainStatisticRepository.findAll();

        try (InputStream is = new ClassPathResource("templates/sme.docx").getInputStream();
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
