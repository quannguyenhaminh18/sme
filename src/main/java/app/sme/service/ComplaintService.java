package app.sme.service;

import app.sme.entity.Category;
import app.sme.projection.ExcelChartSheetProjection;
import app.sme.projection.NormalExcelSheetProjection;
import app.sme.repo.CategoryRepository;
import app.sme.repo.ComplaintRepository;
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
            // Kiểm tra file đầu ra đã tồn tại chưa
            File outputFile = new File(OUTPUT_XLSX_PATH);
            InputStream inputStream;
            if (outputFile.exists()) {
                inputStream = new FileInputStream(outputFile);
            } else {
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
                inputStream = new FileInputStream(outputFile);
            }

            try (Workbook workbook = WorkbookFactory.create(inputStream);
                 ByteArrayOutputStream out = new ByteArrayOutputStream()) {

                // --- Sheet 1: SL_ngay ---
                Sheet sheet1 = getRequiredSheet(workbook, "SL_ngay");

                // Update header H4 + C4
                LocalDate today = LocalDate.now();
                int month = today.getMonthValue();
                int year = today.getYear();
                YearMonth yearMonth = YearMonth.of(year, month);
                int daysInMonth = yearMonth.lengthOfMonth();

                Row row4 = sheet1.getRow(3);
                if (row4 == null) {
                    throw new IllegalStateException("Không tìm thấy row 4 trong sheet SL_ngay");
                }

                Cell h4Cell = row4.getCell(7, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                h4Cell.setCellValue("T" + month + "." + year);

                Cell cellC4 = row4.getCell(2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                cellC4.setCellValue("Luỹ kế T" + month + "." + year);

                // Update header ngày (H5–AL5)
                Row row5 = sheet1.getRow(4);
                DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd/MM");
                int startColSL = 7; // H
                for (int day = 1; day <= daysInMonth; day++) {
                    LocalDate date = LocalDate.of(year, month, day);
                    int colIndex = startColSL + day - 1;
                    Cell cell = row5.getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    cell.setCellValue(date.format(formatter));
                }

                // Lấy danh sách category từ CategoryRepository
                List<Category> categories = categoryRepository.findAll();
                // Lấy dữ liệu từ view hiện tại và view tháng trước
                List<NormalExcelSheetProjection> data = complaintRepository.findAllSMEView();
                List<ExcelChartSheetProjection> previousMonthData = complaintRepository.findErrorComplaintsLastMonth();

                // Tạo map để ánh xạ category với count_yesterday
                Map<String, Long> categoryCounts = new HashMap<>();
                for (NormalExcelSheetProjection projection : data) {
                    if (projection.getCategory() == null) continue; // Bỏ qua hàng tổng hợp
                    Long countYesterday = projection.getCountYesterday();
                    categoryCounts.put(projection.getCategory(), countYesterday != null ? countYesterday : 0L);
                }

                // Điền dữ liệu B7–B38 (category names) và H7–AL38 (count_yesterday cho ngày hôm qua)
                int rowIndex = 6; // Bắt đầu từ B7
                int yesterdayCol = startColSL + (today.getDayOfMonth() - 2); // Cột cho ngày hôm qua
                for (Category category : categories) {
                    if (rowIndex > 37) break; // Giới hạn đến B38
                    Row row = sheet1.getRow(rowIndex);
                    if (row == null) row = sheet1.createRow(rowIndex);

                    // Điền category name vào cột B
                    Cell cellB = row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    cellB.setCellValue(category.getName());

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
                if (row3 == null) {
                    throw new IllegalStateException("Không tìm thấy row 3 trong sheet KPI_ngay");
                }
                Cell d3Cell = row3.getCell(3, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                d3Cell.setCellValue("T" + month + "." + year);

                // Update header ngày (E3–AI3)
                Row row3KPI = sheet2.getRow(2);
                int startColKPI = 4; // E
                for (int day = 1; day <= daysInMonth; day++) {
                    LocalDate date = LocalDate.of(year, month, day);
                    int colIndex = startColKPI + day - 1;
                    Cell cell = row3KPI.getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    cell.setCellValue(date.format(formatter));
                }

                // Tạo map để ánh xạ category với số khiếu nại theo ngày, tổng từ đầu tháng, số bản ghi đóng, và số bản ghi đúng hạn
                Map<String, Map<Integer, Long>> categoryDayCounts = new HashMap<>();
                Map<String, Long> categoryMonthlyTotals = new HashMap<>();
                Map<String, Map<Integer, Long>> categoryClosedDayCounts = new HashMap<>();
                Map<String, Long> categoryClosedMonthlyTotals = new HashMap<>();
                Map<String, Map<Integer, Long>> categoryOnTimeDayCounts = new HashMap<>();
                Map<String, Long> categoryOnTimeMonthlyTotals = new HashMap<>();
                Map<String, Long> categoryCooperationUnitTotals = new HashMap<>();
                for (NormalExcelSheetProjection projection : data) {
                    if (projection.getCategory() == null || projection.getDayOfMonth() == null) continue;
                    String category = projection.getCategory();
                    Integer day = projection.getDayOfMonth();
                    Long totalComplaints = projection.getTotalComplaintsPerDay();
                    Long countYesterday = projection.getCountYesterday();
                    Long totalClosed = projection.getTotalClosedPerDay();
                    Long totalReceivedYesterday = projection.getTotalReceivedYesterday();
                    Long totalOnTime = projection.getTotalOnTimePerDay();
                    Long totalOnTimeYesterday = projection.getTotalOnTimeYesterday();
                    Long cooperationUnitCount = projection.getCooperationUnitCount();

                    // Lưu số khiếu nại theo ngày
                    if (totalComplaints != null) {
                        categoryDayCounts.computeIfAbsent(category, k -> new HashMap<>()).put(day, totalComplaints);
                        // Tính tổng từ đầu tháng đến hôm qua
                        if (day <= today.getDayOfMonth() - 1) {
                            categoryMonthlyTotals.merge(category, totalComplaints, Long::sum);
                        }
                    }
                    // Lưu count_yesterday cho ngày hôm qua
                    if (countYesterday != null && day == today.getDayOfMonth() - 1) {
                        categoryDayCounts.computeIfAbsent(category, k -> new HashMap<>()).put(day, countYesterday);
                    }
                    // Lưu số bản ghi đóng theo ngày
                    if (totalClosed != null) {
                        categoryClosedDayCounts.computeIfAbsent(category, k -> new HashMap<>()).put(day, totalClosed);
                        // Tính tổng số bản ghi đóng từ đầu tháng đến hôm qua
                        if (day <= today.getDayOfMonth() - 1) {
                            categoryClosedMonthlyTotals.merge(category, totalClosed, Long::sum);
                        }
                    }
                    // Lưu total_received_yesterday cho ngày hôm qua
                    if (totalReceivedYesterday != null && day == today.getDayOfMonth() - 1) {
                        categoryClosedDayCounts.computeIfAbsent(category, k -> new HashMap<>()).put(day, totalReceivedYesterday);
                    }
                    // Lưu số bản ghi đúng hạn theo ngày
                    if (totalOnTime != null) {
                        categoryOnTimeDayCounts.computeIfAbsent(category, k -> new HashMap<>()).put(day, totalOnTime);
                        // Tính tổng số bản ghi đúng hạn từ đầu tháng đến hôm qua
                        if (day <= today.getDayOfMonth() - 1) {
                            categoryOnTimeMonthlyTotals.merge(category, totalOnTime, Long::sum);
                        }
                    }
                    // Lưu total_on_time_yesterday cho ngày hôm qua
                    if (totalOnTimeYesterday != null && day == today.getDayOfMonth() - 1) {
                        categoryOnTimeDayCounts.computeIfAbsent(category, k -> new HashMap<>()).put(day, totalOnTimeYesterday);
                    }
                    // Lưu tổng cooperation_unit_count từ đầu tháng đến hôm qua
                    if (cooperationUnitCount != null && day <= today.getDayOfMonth() - 1) {
                        categoryCooperationUnitTotals.merge(category, cooperationUnitCount, Long::sum);
                    }
                }

                // Định nghĩa các hàng và category tương ứng cho Sheet 2
                Map<Integer, String> rowToCategory = new HashMap<>();
                // Tổng khiếu nại
                rowToCategory.put(7, "Dịch vụ CA"); // Hàng 8
                rowToCategory.put(12, "Dịch vụ BHXH"); // Hàng 13
                rowToCategory.put(17, "Dịch vụ Hóa đơn điện tử"); // Hàng 18
                rowToCategory.put(22, "Dịch vụ vTracking"); // Hàng 23
                rowToCategory.put(49, "Camera NĐ10"); // Hàng 50
                rowToCategory.put(54, "MySign"); // Hàng 55
                rowToCategory.put(59, "vESS"); // Hàng 60
                // Bản ghi đúng hạn
                rowToCategory.put(28, "Dịch vụ CA"); // Hàng 29
                rowToCategory.put(33, "Dịch vụ BHXH"); // Hàng 34
                rowToCategory.put(38, "Dịch vụ Hóa đơn điện tử"); // Hàng 39
                rowToCategory.put(43, "Dịch vụ vTracking"); // Hàng 44
                rowToCategory.put(65, "Camera NĐ10"); // Hàng 66
                rowToCategory.put(70, "MySign"); // Hàng 71
                rowToCategory.put(75, "vESS"); // Hàng 76
                // Bản ghi đóng
                rowToCategory.put(29, "Dịch vụ CA"); // Hàng 30
                rowToCategory.put(34, "Dịch vụ BHXH"); // Hàng 35
                rowToCategory.put(39, "Dịch vụ Hóa đơn điện tử"); // Hàng 40
                rowToCategory.put(44, "Dịch vụ vTracking"); // Hàng 45
                rowToCategory.put(66, "Camera NĐ10"); // Hàng 67
                rowToCategory.put(71, "MySign"); // Hàng 72
                rowToCategory.put(76, "vESS"); // Hàng 77

                // Điền dữ liệu cho Sheet 2
                int yesterdayColKPI = startColKPI + (today.getDayOfMonth() - 2); // Cột cho ngày hôm qua
                for (Map.Entry<Integer, String> entry : rowToCategory.entrySet()) {
                    int rowNum = entry.getKey();
                    String category = entry.getValue();
                    Row row = sheet2.getRow(rowNum);
                    if (row == null) row = sheet2.createRow(rowNum);

                    // Điền tổng số (tổng khiếu nại, bản ghi đóng, hoặc bản ghi đúng hạn) từ đầu tháng đến hôm qua vào cột D
                    Cell cellD = row.getCell(3, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    Long total;
                    if (rowNum <= 22 || rowNum == 49 || rowNum == 54 || rowNum == 59) {
                        // Hàng 8, 13, 18, 23, 50, 55, 60: Tổng khiếu nại
                        total = categoryMonthlyTotals.getOrDefault(category, 0L);
                    } else if (rowNum <= 29 || rowNum == 33 || rowNum == 38 || rowNum == 43 || rowNum == 65 || rowNum == 70 || rowNum == 75) {
                        // Hàng 29, 34, 39, 44, 66, 71, 76: Tổng bản ghi đúng hạn
                        total = categoryOnTimeMonthlyTotals.getOrDefault(category, 0L);
                    } else {
                        // Hàng 30, 35, 40, 45, 67, 72, 77: Tổng bản ghi đóng
                        total = categoryClosedMonthlyTotals.getOrDefault(category, 0L);
                    }
                    cellD.setCellValue(total != null ? total : 0);

                    // Điền số liệu ngày hôm qua
                    Cell cellYesterday = row.getCell(yesterdayColKPI, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    Long count;
                    if (rowNum <= 22 || rowNum == 49 || rowNum == 54 || rowNum == 59) {
                        // Hàng 8, 13, 18, 23, 50, 55, 60: count_yesterday
                        count = categoryDayCounts.getOrDefault(category, new HashMap<>()).getOrDefault(today.getDayOfMonth() - 1, 0L);
                    } else if (rowNum <= 29 || rowNum == 33 || rowNum == 38 || rowNum == 43 || rowNum == 65 || rowNum == 70 || rowNum == 75) {
                        // Hàng 29, 34, 39, 44, 66, 71, 76: total_on_time_yesterday
                        count = categoryOnTimeDayCounts.getOrDefault(category, new HashMap<>()).getOrDefault(today.getDayOfMonth() - 1, 0L);
                    } else {
                        // Hàng 30, 35, 40, 45, 67, 72, 77: total_received_yesterday
                        count = categoryClosedDayCounts.getOrDefault(category, new HashMap<>()).getOrDefault(today.getDayOfMonth() - 1, 0L);
                    }
                    cellYesterday.setCellValue(count != null ? count : 0);
                }

                // --- Sheet 4: TLXLTB_PH ---
                Sheet sheet4 = getRequiredSheet(workbook, "TLXLTB_PH");

                // Update ô C3
                Row row3Sheet4 = sheet4.getRow(2);
                if (row3Sheet4 == null) {
                    throw new IllegalStateException("Không tìm thấy row 3 trong sheet TLXLTB_PH");
                }
                Cell c3Cell = row3Sheet4.getCell(2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                c3Cell.setCellValue("T" + month + "." + year);

                // Điền dữ liệu F7–F38 (tổng cooperation_unit_count từ đầu tháng đến hôm qua)
                int rowIndexSheet4 = 6; // Bắt đầu từ F7
                for (Category category : categories) {
                    if (rowIndexSheet4 > 37) break; // Giới hạn đến F38
                    Row row = sheet4.getRow(rowIndexSheet4);
                    if (row == null) row = sheet4.createRow(rowIndexSheet4);

                    // Điền tổng cooperation_unit_count vào cột F
                    Cell cellF = row.getCell(5, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    Long cooperationUnitTotal = categoryCooperationUnitTotals.getOrDefault(category.getName(), 0L);
                    cellF.setCellValue(cooperationUnitTotal != null ? cooperationUnitTotal : 0);
                    rowIndexSheet4++;
                }

                // --- Sheet 5: Bieu_Do ---
                Sheet sheet5 = getRequiredSheet(workbook, "Bieu_Do");

                // Update ô E3 (ngày hôm qua)
                Row row3Sheet5 = sheet5.getRow(2);
                if (row3Sheet5 == null) {
                    throw new IllegalStateException("Không tìm thấy row 3 trong sheet Bieu_Do");
                }
                Cell e3Cell = row3Sheet5.getCell(4, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                LocalDate yesterday = today.minusDays(1);
                e3Cell.setCellValue(yesterday.format(DateTimeFormatter.ofPattern("dd/MM/yyyy")));

                // Update ô H3
                Cell h3Cell = row3Sheet5.getCell(7, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                h3Cell.setCellValue("T" + month + "." + year);

                // Update ô K4
                Row row4Sheet5 = sheet5.getRow(3);
                if (row4Sheet5 == null) {
                    throw new IllegalStateException("Không tìm thấy row 4 trong sheet Bieu_Do");
                }
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
                    Cell cell = headerRow.getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    cell.setCellValue("T" + ym.getMonthValue() + "." + ym.getYear());
                }

                for (int i = 0; i < 8; i++) {
                    LocalDate date = yesterday.minusDays(7 - i); // 7 ngày trước → hôm qua
                    int colIndex = 13 + i; // N = index 13
                    Cell cell = headerRow.getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    cell.setCellValue(date.format(DateTimeFormatter.ofPattern("dd/MM")));
                }


                // Tạo map cho số phản ánh tháng trước và total_subscriber
                Map<String, Long> previousMonthTotals = new HashMap<>();
                for (ExcelChartSheetProjection projection : previousMonthData) {
                    if (projection.getCategory() == null) continue;
                    previousMonthTotals.put(projection.getCategory(), projection.getTotalComplaintsLastMonth());
                }

                Map<String, Long> categorySubscribers = new HashMap<>();
                for (Category category : categories) {
                    categorySubscribers.put(category.getName(), category.getTotalSubscriber());
                }

                // Định nghĩa các hàng và category cho Sheet 5
                Map<Integer, String> sheet5RowToCategory = new HashMap<>();
                sheet5RowToCategory.put(5, "Dịch vụ CA"); // Hàng 6
                sheet5RowToCategory.put(6, "Dịch vụ BHXH"); // Hàng 7
                sheet5RowToCategory.put(7, "Dịch vụ Hóa đơn điện tử"); // Hàng 8
                sheet5RowToCategory.put(8, "Dịch vụ vTracking"); // Hàng 9

                // Điền dữ liệu E6–E9 (count_yesterday) và K6–K9 (tỷ lệ)
                for (Map.Entry<Integer, String> entry : sheet5RowToCategory.entrySet()) {
                    int rowNum = entry.getKey();
                    String category = entry.getValue();
                    Row row = sheet5.getRow(rowNum);
                    if (row == null) row = sheet5.createRow(rowNum);

                    // Điền count_yesterday vào cột E
                    Cell cellE = row.getCell(4, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    Long countYesterday = categoryCounts.getOrDefault(category, 0L);
                    cellE.setCellValue(countYesterday != null ? countYesterday : 0);

                    // Điền tỷ lệ phản ánh/10k thuê bao vào cột K
                    Cell cellK = row.getCell(10, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    Long totalComplaintsLastMonth = previousMonthTotals.getOrDefault(category, 0L);
                    Long totalSubscriber = categorySubscribers.getOrDefault(category, 0L);
                    double ratio = (totalSubscriber != null && totalSubscriber > 0)
                            ? (totalComplaintsLastMonth.doubleValue() / totalSubscriber) * 10000
                            : 0.0;
                    cellK.setCellValue(ratio);
                }
                Row row14 = sheet5.getRow(13); // hàng 14 (Excel index bắt đầu từ 0)
                if (row14 == null) {
                    row14 = sheet5.createRow(13);
                }

                long avgLastYear = complaintRepository.getComplaintSummary().stream()
                        .filter(p -> "AVG_LAST_YEAR".equals(p.getDimensionType()))
                        .map(ExcelChartSheetProjection::getTotalComplaints)
                        .findFirst()
                        .orElse(0L);

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
                    Cell cell = row14.getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    cell.setCellValue(value);
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
                    Cell cell = row14.getCell(colIndex, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    cell.setCellValue(value);
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
                if (sheet == null) throw new RuntimeException("Không tìm thấy sheet Bieu_Do");

                // Lấy giá trị từ ô U14, N14, L14, K14...
                double numPA = sheet.getRow(13).getCell(20).getNumericCellValue(); // U14 (0-indexed)
                double numPALastWeek = sheet.getRow(13).getCell(13).getNumericCellValue(); // N14
                double numPALastMonth = 0; // theo yêu cầu
                double numPACumulative = sheet.getRow(13).getCell(12).getNumericCellValue(); // L14
                double prevMonthTotal = sheet.getRow(13).getCell(11).getNumericCellValue(); // L13

                double pctChangeWeek = ((numPA - numPALastWeek) / numPALastWeek) * 100;
                double pctChangeMonth = ((numPACumulative - numPALastMonth)) * 100;
                pctChangeWeek = Math.round(pctChangeWeek);
                pctChangeMonth = Math.round(pctChangeMonth);

                // Tính avg_PA_per_day
                LocalDate now = LocalDate.now();
                int dayOfMonth = now.getDayOfMonth();
                double avgPerDay = numPACumulative / Math.max(1, dayOfMonth);

                YearMonth prevMonth = YearMonth.now().minusMonths(1);
                int daysPrevMonth = prevMonth.lengthOfMonth();
                double avgPerDayLastMonth = prevMonthTotal / daysPrevMonth;

                double pctChangeAvg = ((avgPerDay - avgPerDayLastMonth) / Math.max(1, avgPerDayLastMonth)) * 100;
                pctChangeAvg = Math.round(pctChangeAvg);

                // Đếm "Đạt" từ G6:G9 và J6:J9
                int passedByDay = 0;
                int passedByMonth = 0;
                for (int r = 5; r <= 8; r++) { // 0-indexed
                    Cell cDay = sheet.getRow(r).getCell(6); // G
                    if (cDay != null && "Đạt".equals(cDay.getStringCellValue())) passedByDay++;
                    Cell cMonth = sheet.getRow(r).getCell(9); // J
                    if (cMonth != null && "Đạt".equals(cMonth.getStringCellValue())) passedByMonth++;
                }

                // Gán vào map placeholder
                map.put("${num_PA}", String.valueOf((int) numPA));
                map.put("${pct_change_week}", (pctChangeWeek >= 0 ? "tăng " : "giảm ") + Math.abs((int) pctChangeWeek));
                map.put("${num_PA_last_week}", String.valueOf((int) numPALastWeek));
                map.put("${pct_change_month}", (pctChangeMonth >= 0 ? "tăng " : "giảm ") + Math.abs((int) pctChangeMonth));
                map.put("${num_PA_last_month}", String.valueOf(0));
                map.put("${passed_by_day}", String.valueOf(passedByDay));
                map.put("${num_PA_cumulative}", String.valueOf((int) numPACumulative));
                map.put("${avg_PA_per_day}", String.valueOf((int) avgPerDay));
                map.put("${avg_PA_per_day_last_month}", String.valueOf((int) avgPerDayLastMonth));
                map.put("${pct_change_avg}", (pctChangeAvg >= 0 ? "tăng " : "giảm ") + Math.abs((int) pctChangeAvg));
                map.put("${passed_by_month}", String.valueOf(passedByMonth));
                map.put("${closed_complaint_this_month}", "0");
                map.put("${closed_complaint_percent}", "0");

                // Thêm placeholder ngày tháng
                map.put("${day}", String.valueOf(dayOfMonth));
                map.put("${month}", String.valueOf(now.getMonthValue()));
                map.put("${year}", String.valueOf(now.getYear()));

            }
            replaceTextInDocument(doc, map);
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
        for (XWPFRun run : paragraph.getRuns()) {
            String text = run.text();
            boolean changed = false;
            for (Map.Entry<String, String> e : map.entrySet()) {
                if (text.contains(e.getKey())) {
                    text = text.replace(e.getKey(), e.getValue());
                    changed = true;
                }
            }
            if (changed) {
                run.setText(text, 0); // ghi đè text mà không mất style
            }
        }
    }

}
