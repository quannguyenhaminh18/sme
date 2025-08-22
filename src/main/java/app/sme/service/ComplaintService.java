package app.sme.service;

import app.sme.projection.NormalSheetDTO;
import app.sme.entity.Category;
import app.sme.projection.ExcelChartSheetProjection;
import app.sme.repo.CategoryRepository;
import app.sme.repo.ComplaintRepository;
import app.sme.util.WordUtil;
import com.aspose.cells.ImageType;
import lombok.RequiredArgsConstructor;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Service;

import java.io.*;
import java.time.LocalDate;
import java.time.YearMonth;
import java.time.format.DateTimeFormatter;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.function.Function;

@Service
@RequiredArgsConstructor
public class ComplaintService {
    private final ComplaintRepository complaintRepository;
    private final CategoryRepository categoryRepository;
    private static final String TEMPLATES_XLSX_PATH = "/templates/SME.xlsx";
    private static final String OUTPUT_XLSX_PATH = "/app/sme/data/SME_output.xlsx";

    private Sheet getRequiredSheet(Workbook workbook, String sheetName) {
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            throw new IllegalArgumentException("Không tìm thấy sheet " + sheetName);
        }
        return sheet;
    }

    public byte[] exportExcelFile() {
        try {
            File outputFile = new File(OUTPUT_XLSX_PATH);
            System.out.println("OUTPUT FILE PATH: " + outputFile.getAbsolutePath());

            InputStream inputStream;
            if (!outputFile.exists()) {
                inputStream = getClass().getResourceAsStream(TEMPLATES_XLSX_PATH);
                if (inputStream == null) {
                    throw new IllegalArgumentException("Không tìm thấy file template SME.xlsx");
                }
                File outputDir = new File(outputFile.getParent());
                if (!outputDir.exists()) {
                    outputDir.mkdirs();
                }
                try (FileOutputStream fileOut = new FileOutputStream(outputFile)) {
                    byte[] buffer = new byte[1024];
                    int bytesRead;
                    while ((bytesRead = inputStream.read(buffer)) != -1) {
                        fileOut.write(buffer, 0, bytesRead);
                    }
                }
                inputStream.close();
            }
            inputStream = new FileInputStream(outputFile);

            try (Workbook workbook = WorkbookFactory.create(inputStream);
                 ByteArrayOutputStream out = new ByteArrayOutputStream()) {
                // --- Sheet 1: SL_ngay ---
                Sheet sheet1 = getRequiredSheet(workbook, "SL_ngay");

                // Update header H4 + C4
                LocalDate today = LocalDate.now();
                int month = today.getMonthValue();
                int year = today.getYear();
                Row row4 = sheet1.getRow(3);
                Cell h4Cell = row4.getCell(7, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                h4Cell.setCellValue("T" + month + "." + year);
                Cell cellC4 = row4.getCell(2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                cellC4.setCellValue("Luỹ kế T" + month + "." + year);

                // Update header ngày (H5 + h42)
                Row row5 = sheet1.getRow(4);
                Cell cell = row5.getCell(7, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                cell.setCellValue(LocalDate.of(year, month, 1));

                // Lấy danh sách category từ CategoryRepository
                List<Category> categories = categoryRepository.findAll();
                List<NormalSheetDTO> data = complaintRepository.findTotalView();

                // Tạo map để ánh xạ category với count_yesterday
                Map<String, Long> categoryCounts = new HashMap<>();
                for (NormalSheetDTO projection : data) {
                    Long countYesterday = projection.getTotalReceivedYesterday();
                    categoryCounts.put(projection.getCategory(), countYesterday);
                }
                System.out.println("Category Counts: " + categoryCounts);
                // Điền dữ liệu B7–B38 (category names) và H7–AL38 (count_yesterday cho ngày hôm qua)
                int rowIndex = 6; // Bắt đầu từ B7
                int yesterdayCol = 7 + (today.getDayOfMonth() - 2); // Cột cho ngày hôm qua
                for (Category category : categories) {
                    if (rowIndex > 37) break; // Giới hạn đến B38
                    Row row = sheet1.getRow(rowIndex);
                    // Điền count_yesterday vào cột tương ứng với ngày hôm qua
                    Cell cellYesterday = row.getCell(yesterdayCol, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    Long count = categoryCounts.getOrDefault(category.getName(), 0L);
                    cellYesterday.setCellValue(count != null ? count : 0);
                    rowIndex++;
                }

                // --- Sheet 2: KPI_ngay ---
                Sheet sheet2 = getRequiredSheet(workbook, "KPI_ngay");

                // Update header D3
                Row row3 = sheet2.getRow(2);
                Cell d3Cell = row3.getCell(3, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                d3Cell.setCellValue("T" + month + "." + year);

                // Update header ngày (E3–AI3)
                Row row3KPI = sheet2.getRow(2);
                Cell cellE3 = row3KPI.getCell(4, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                cellE3.setCellValue(LocalDate.of(year, month, 1));

                int todayColIndex = 4 + (today.getDayOfMonth() - 1); // E=4 (0-based)

                // Map: row index -> function lấy dữ liệu từ DTO
                Map<Integer, Function<NormalSheetDTO, Long>> rowMapping = new HashMap<>();

                rowMapping.put(7, dto -> "Dịch vụ CA".equals(dto.getCategory()) ? dto.getTotalErrorReceivedYesterday() : null);
                rowMapping.put(12, dto -> "Dịch vụ BHXH".equals(dto.getCategory()) ? dto.getTotalErrorReceivedYesterday() : null);
                rowMapping.put(17, dto -> "Dịch vụ Hóa đơn điện tử".equals(dto.getCategory()) ? dto.getTotalErrorReceivedYesterday() : null);
                rowMapping.put(22, dto -> "Dịch vụ vTracking".equals(dto.getCategory()) ? dto.getTotalErrorReceivedYesterday() : null);
                rowMapping.put(49, dto -> "Camera NĐ10".equals(dto.getCategory()) ? dto.getTotalErrorReceivedYesterday() : null);
                rowMapping.put(54, dto -> "MySign".equals(dto.getCategory()) ? dto.getTotalErrorReceivedYesterday() : null);
                rowMapping.put(59, dto -> "vESS".equals(dto.getCategory()) ? dto.getTotalErrorReceivedYesterday() : null);

                rowMapping.put(28, dto -> "Dịch vụ CA".equals(dto.getCategory()) ? dto.getTotalErrorClosedOntimeYesterday() : null);
                rowMapping.put(33, dto -> "Dịch vụ BHXH".equals(dto.getCategory()) ? dto.getTotalErrorClosedOntimeYesterday() : null);
                rowMapping.put(38, dto -> "Dịch vụ Hóa đơn điện tử".equals(dto.getCategory()) ? dto.getTotalErrorClosedOntimeYesterday() : null);
                rowMapping.put(43, dto -> "Dịch vụ vTracking".equals(dto.getCategory()) ? dto.getTotalErrorClosedOntimeYesterday() : null);
                rowMapping.put(65, dto -> "Camera NĐ10".equals(dto.getCategory()) ? dto.getTotalErrorClosedOntimeYesterday() : null);
                rowMapping.put(70, dto -> "MySign".equals(dto.getCategory()) ? dto.getTotalErrorClosedOntimeYesterday() : null);
                rowMapping.put(75, dto -> "vESS".equals(dto.getCategory()) ? dto.getTotalErrorClosedOntimeYesterday() : null);

                rowMapping.put(29, dto -> "Dịch vụ CA".equals(dto.getCategory()) ? dto.getTotalErrorClosedYesterday() : null);
                rowMapping.put(34, dto -> "Dịch vụ BHXH".equals(dto.getCategory()) ? dto.getTotalErrorClosedYesterday() : null);
                rowMapping.put(39, dto -> "Dịch vụ Hóa đơn điện tử".equals(dto.getCategory()) ? dto.getTotalErrorClosedYesterday() : null);
                rowMapping.put(44, dto -> "Dịch vụ vTracking".equals(dto.getCategory()) ? dto.getTotalErrorClosedYesterday() : null);
                rowMapping.put(66, dto -> "Camera NĐ10".equals(dto.getCategory()) ? dto.getTotalErrorClosedYesterday() : null);
                rowMapping.put(71, dto -> "MySign".equals(dto.getCategory()) ? dto.getTotalErrorClosedYesterday() : null);
                rowMapping.put(76, dto -> "vESS".equals(dto.getCategory()) ? dto.getTotalErrorClosedYesterday() : null);

                for (NormalSheetDTO dto : data) {
                    for (Map.Entry<Integer, Function<NormalSheetDTO, Long>> entry : rowMapping.entrySet()) {
                        Long value = entry.getValue().apply(dto);
                        if (value != null) {
                            int rowIdx = entry.getKey();
                            Row row = sheet2.getRow(rowIdx);
                            if (row == null) row = sheet2.createRow(rowIdx);
                            Cell target = row.getCell(todayColIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                            target.setCellValue(value);
                        }
                    }
                }

                // Số ngày từ đầu tháng đến hôm qua
                int daysFromStartToYesterday = today.getDayOfMonth() - 1;
                if (daysFromStartToYesterday <= 0) {
                    daysFromStartToYesterday = 1; // tránh chia 0 nếu hôm nay là ngày 1
                }

                int[] targetRows = {5, 10, 15, 20, 47, 52, 57}; // 0-based index

                for (int rowIdx : targetRows) {
                    Row targetRow = sheet2.getRow(rowIdx);
                    if (targetRow == null) targetRow = sheet2.createRow(rowIdx);
                    Cell targetCell = targetRow.getCell(3, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);

                    // Lấy D(row+2) và D(row+3)
                    Row rowPlus2 = sheet2.getRow(rowIdx + 2);
                    Row rowPlus3 = sheet2.getRow(rowIdx + 3);

                    double dPlus2 = 0;
                    double dPlus3 = 1; // tránh chia 0

                    if (rowPlus2 != null) {
                        Cell cPlus2 = rowPlus2.getCell(3);
                        if (cPlus2 != null && cPlus2.getCellType() == CellType.NUMERIC) {
                            dPlus2 = cPlus2.getNumericCellValue();
                        }
                    }
                    if (rowPlus3 != null) {
                        Cell cPlus3 = rowPlus3.getCell(3);
                        if (cPlus3 != null && cPlus3.getCellType() == CellType.NUMERIC) {
                            dPlus3 = cPlus3.getNumericCellValue();
                        }
                    }

                    // Tính giá trị
                    double value = dPlus2 * 10000.0 / dPlus3 / daysFromStartToYesterday;
                    if (targetCell.getCellType() == CellType.FORMULA) {
                        targetCell.setCellFormula(null);
                    }
                    targetCell.setCellValue(value);
                }

                // --- Sheet 3: Tong_hop ---

                Sheet sheet3 = getRequiredSheet(workbook, "Tong_hop");
                // Update header D5
                Row row5Sheet3 = sheet3.getRow(4);
                Cell D5Cell = row5Sheet3.getCell(3, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                D5Cell.setCellValue("T" + month + "." + year);
                // Update header  (D13)
                Row row13Sheet3 = sheet3.getRow(12);
                Cell D13Cell = row13Sheet3.getCell(3, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                D13Cell.setCellValue("T" + month + "." + year);

                // Update header c22
                Row row22Sheet3 = sheet3.getRow(21);
                Cell c22Cell = row22Sheet3.getCell(2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                c22Cell.setCellValue("T" + month + "." + year);

                // --- Sheet 4: TLXLTB_PH ---
                Sheet sheet4 = getRequiredSheet(workbook, "TLXLTB_PH");

                // Update ô C3
                Row row3Sheet4 = sheet4.getRow(2);
                Cell c3Cell = row3Sheet4.getCell(2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                c3Cell.setCellValue("T" + month + "." + year);

                int rowIndexSheet4 = 6; // F7
                for (Category category : categories) {
                    if (rowIndexSheet4 > 37) break; // dừng F38
                    Row row = sheet4.getRow(rowIndexSheet4);
                    if (row == null) row = sheet4.createRow(rowIndexSheet4);

                    Cell cellF = row.getCell(5, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK); // cột F
                    cellF.setCellFormula(null);
                    // Tìm DTO có category trùng
                    Optional<NormalSheetDTO> matchedDTO = data.stream()
                            .filter(dto -> dto.getCategory() != null && category.getName().equals(dto.getCategory()))
                            .findFirst();

                    Long cooperationUnitTotal = matchedDTO.map(NormalSheetDTO::getTotalCooperationUnitMonth).orElse(0L);

                    cellF.setCellValue(cooperationUnitTotal);

                    rowIndexSheet4++;
                }

                // --- Sheet 5: Bieu_Do ---
                Sheet sheet5 = getRequiredSheet(workbook, "Bieu_Do");

                // Update ô E3 (ngày hôm qua)
                Row row3Sheet5 = sheet5.getRow(2);
                Cell e3Cell = row3Sheet5.getCell(4, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                LocalDate yesterday = today.minusDays(1);
                e3Cell.setCellValue(yesterday.format(DateTimeFormatter.ofPattern("dd/MM/yyyy")));

                // Update ô H3
                Cell h3Cell = row3Sheet5.getCell(7, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                h3Cell.setCellValue("T" + month + "." + year);

                // Update ô K4
                Row row4Sheet5 = sheet5.getRow(3);
                Cell k4Cell = row4Sheet5.getCell(10, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                k4Cell.setCellValue("SS với T" + (month - 1) + "." + year);

                // --- Update header rows (Row 13 in sheet "Bieu_Do") ---
                Row headerRow = sheet5.getRow(12); // Hàng 13 (index 12)
                if (headerRow == null) headerRow = sheet5.createRow(12);

                Cell c13 = headerRow.getCell(2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                c13.setCellValue("TB " + (year - 1));

                Cell c14 = sheet5.getRow(13) != null ? sheet5.getRow(13).getCell(2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK)
                        : sheet5.createRow(13).createCell(2);
                c14.setCellValue("TB " + year);

                YearMonth currentYM = YearMonth.of(year, month);
                for (int i = 0; i < 8; i++) {
                    YearMonth ym = currentYM.minusMonths(7 - i); // Lấy từ 7 tháng trước đến tháng hiện tại
                    int colIndex = 4 + i; // E = index 4
                    Cell cellRow13Header = headerRow.getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    cellRow13Header.setCellValue("T" + ym.getMonthValue() + "." + ym.getYear());
                }

                for (int i = 0; i < 8; i++) {
                    LocalDate date = yesterday.minusDays(7 - i); // 7 ngày trước → hôm qua
                    int colIndex = 13 + i; // N = index 13
                    Cell cellRow13In8Day = headerRow.getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    cellRow13In8Day.setCellValue(date.format(DateTimeFormatter.ofPattern("dd/MM")));
                }

                long avgLastYear = complaintRepository.getComplaintSummary().stream()
                        .filter(p -> "AVG_LAST_YEAR".equals(p.getDimensionType()))
                        .map(ExcelChartSheetProjection::getTotalComplaints)
                        .findFirst()
                        .orElse(0L);

                Row row14 = sheet5.getRow(13); // Hàng 14 (index 13)
                Cell avgLastYearCell = row14.getCell(2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK); // C = index 2
                avgLastYearCell.setCellValue(avgLastYear);

                long avgThisYear = complaintRepository.getComplaintSummary().stream()
                        .filter(p -> "AVG_THIS_YEAR".equals(p.getDimensionType()))
                        .map(ExcelChartSheetProjection::getTotalComplaints)
                        .findFirst()
                        .orElse(0L);

                Cell avgThisYearCell = row14.getCell(3, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK); // D = index 3
                avgThisYearCell.setCellValue(avgThisYear);

                for (int i = 0; i < 8; i++) {
                    YearMonth ym = currentYM.minusMonths(7 - i);
                    String dimValue = ym.toString(); // "2025-01"
                    long value = complaintRepository.getComplaintSummary().stream()
                            .filter(p -> "MONTH".equals(p.getDimensionType()) && dimValue.equals(p.getDimensionValue()))
                            .map(ExcelChartSheetProjection::getTotalComplaints)
                            .findFirst()
                            .orElse(0L);

                    int colIndex = 4 + i; // E=4
                    Cell cellRow14 = row14.getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    cellRow14.setCellValue(value);
                }

                for (int i = 0; i < 8; i++) {
                    LocalDate date = yesterday.minusDays(7 - i);
                    String dimValue = date.toString(); // "2025-08-12"
                    long value = complaintRepository.getComplaintSummary().stream()
                            .filter(p -> "DAY".equals(p.getDimensionType()) && dimValue.equals(p.getDimensionValue()))
                            .map(ExcelChartSheetProjection::getTotalComplaints)
                            .findFirst()
                            .orElse(0L);

                    int colIndex = 13 + i; // N=13
                    Cell cellRow14Day = row14.getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    cellRow14Day.setCellValue(value);
                }

                // Ghi đè file vào OUTPUT_PATH
                try (FileOutputStream fileOut = new FileOutputStream(OUTPUT_XLSX_PATH)) {
                    workbook.write(fileOut);
                }

                // Xuất file dưới dạng byte array
                workbook.write(out);
                return out.toByteArray();
            }
        } catch (IOException e) {
            throw new RuntimeException("Lỗi đọc/ghi file Excel", e);
        }
    }

    public byte[] exportDoc() {
        try (InputStream is = new ClassPathResource("templates/SME.docx").getInputStream();
             OPCPackage pkg = OPCPackage.open(is);
             XWPFDocument doc = new XWPFDocument(pkg);
             ByteArrayOutputStream out = new ByteArrayOutputStream()) {

            // -------- Mở Excel --------
            File excelFile = new File(OUTPUT_XLSX_PATH);
            if (!excelFile.exists()) {
                throw new RuntimeException("File Excel không tồn tại: " + OUTPUT_XLSX_PATH);
            }

            Map<String, String> map = new HashMap<>();
            try (FileInputStream fis = new FileInputStream(excelFile);
                 Workbook workbook = WorkbookFactory.create(fis)) {

                Sheet sheet = workbook.getSheet("Bieu_Do");

                double numPA = sheet.getRow(13).getCell(20).getNumericCellValue(); // U14
                double numPALastWeek = sheet.getRow(13).getCell(13).getNumericCellValue(); // N14
                double numPALastMonth = sheet.getRow(13).getCell(10).getNumericCellValue(); // K14
                double numPACumulative = sheet.getRow(13).getCell(12).getNumericCellValue(); // L14
                double prevMonthTotal = sheet.getRow(13).getCell(11).getNumericCellValue(); // L13

                double pctChangeWeek = ((numPA - numPALastWeek) / Math.max(1, numPALastWeek)) * 100;
                double pctChangeMonth = ((numPACumulative - numPALastMonth) / Math.max(1, numPALastMonth)) * 100;

                pctChangeWeek = Math.round(pctChangeWeek);
                pctChangeMonth = Math.round(pctChangeMonth);

                LocalDate now = LocalDate.now();
                int dayOfMonth = now.getDayOfMonth();
                double avgPerDay = numPACumulative / dayOfMonth;

                YearMonth prevMonth = YearMonth.now().minusMonths(1);
                int daysPrevMonth = prevMonth.lengthOfMonth();
                double avgPerDayLastMonth = prevMonthTotal / daysPrevMonth;

                double pctChangeAvg = ((avgPerDay - avgPerDayLastMonth) / Math.max(1, avgPerDayLastMonth)) * 100;
                pctChangeAvg = Math.round(pctChangeAvg);

                int passedByDay = 0;
                int passedByMonth = 0;
                for (int r = 5; r <= 8; r++) { // 0-indexed
                    Cell cDay = sheet.getRow(r).getCell(6); // G
                    if (cDay != null && "Đạt".equals(cDay.getStringCellValue())) passedByDay++;
                    Cell cMonth = sheet.getRow(r).getCell(9); // J
                    if (cMonth != null && "Đạt".equals(cMonth.getStringCellValue())) passedByMonth++;
                }

                map.put("${num_PA}", String.valueOf((int) numPA));
                map.put("${pct_change_week}", (pctChangeWeek >= 0 ? "tăng " : "giảm ") + Math.abs((int) pctChangeWeek));
                map.put("${num_PA_last_week}", String.valueOf((int) numPALastWeek));
                map.put("${pct_change_month}", (pctChangeMonth >= 0 ? "tăng " : "giảm ") + Math.abs((int) pctChangeMonth));
                map.put("${num_PA_last_month}", String.valueOf((int) numPALastMonth));
                map.put("${passed_by_day}", String.valueOf(passedByDay));
                map.put("${num_PA_cumulative}", String.valueOf((int) numPACumulative));
                map.put("${avg_PA_per_day}", String.valueOf((int) avgPerDay));
                map.put("${avg_PA_per_day_last_month}", String.valueOf((int) avgPerDayLastMonth));
                map.put("${pct_change_avg}", (pctChangeAvg >= 0 ? "tăng " : "giảm ") + Math.abs((int) pctChangeAvg));
                map.put("${passed_by_month}", String.valueOf(passedByMonth));
                map.put("${closed_complaint_this_month}", "0");
                map.put("${closed_complaint_percent}", "0");

                LocalDate yesterday = LocalDate.now().minusDays(1);
                map.put("${day}", String.valueOf(yesterday.getDayOfMonth()));
                map.put("${month}", String.valueOf(now.getMonthValue()));
                map.put("${year}", String.valueOf(now.getYear()));
            }
            WordUtil.replaceTextInDocument(doc, map);
            com.aspose.cells.Workbook asposeWorkbook = new com.aspose.cells.Workbook(OUTPUT_XLSX_PATH);
            com.aspose.cells.Worksheet asposeSheet = asposeWorkbook.getWorksheets().get("Bieu_Do");

            com.aspose.cells.ImageOrPrintOptions options = new com.aspose.cells.ImageOrPrintOptions();
            options.setImageType(ImageType.PNG);
            com.aspose.cells.Chart chart = asposeSheet.getCharts().get(0);
            ByteArrayOutputStream chartOut = new ByteArrayOutputStream();
            chart.toImage(chartOut, options);
            byte[] chartBytes = chartOut.toByteArray();

            com.aspose.cells.Range range = asposeSheet.getCells().createRange("A3:L9");
            byte[] tableBytes = range.toImage(options);

            int picCount = 0;
            for (XWPFParagraph p : doc.getParagraphs()) {
                List<XWPFRun> runs = p.getRuns();
                for (int i = 0; i < runs.size(); i++) {
                    XWPFRun run = runs.get(i);
                    for (XWPFPicture pic : run.getEmbeddedPictures()) {
                        picCount++;
                        p.removeRun(i); // xóa run cũ

                        XWPFRun newRun = p.insertNewRun(i);
                        if (picCount == 1) {
                            // ảnh 1: bảng A3:L9
                            newRun.addPicture(
                                    new ByteArrayInputStream(tableBytes),
                                    Document.PICTURE_TYPE_PNG,
                                    "table.png",
                                    Units.toEMU(500),
                                    Units.toEMU(100)
                            );
                        } else {
                            // ảnh 2: chart
                            newRun.addPicture(
                                    new ByteArrayInputStream(chartBytes),
                                    Document.PICTURE_TYPE_PNG,
                                    "chart.png",
                                    Units.toEMU(450),
                                    Units.toEMU(300)
                            );
                        }

                        if (picCount == 2) break;
                    }
                    if (picCount == 2) break;
                }
                if (picCount == 2) break;
            }

            doc.write(out);
            return out.toByteArray();

        } catch (Exception e) {
            throw new RuntimeException("Lỗi fill template DOCX hoặc đọc Excel", e);
        }
    }
}
