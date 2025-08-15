package app.sme;

import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDate;

@Service
@RequiredArgsConstructor
public class ComplaintService {
    private final ComplaintRepository complaintRepository;
    private final ComplaintCountRepository complaintCountRepository;

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

                // Lấy dữ liệu count từ repository
                java.util.List<ComplaintCount> counts = complaintCountRepository.findByDateRange(sqlStart, sqlEnd);
                // Map: category -> (date -> count)
                java.util.Map<String, java.util.Map<Integer, Integer>> dataMap = new java.util.HashMap<>();
                for (ComplaintCount cc : counts) {
                    String cat = cc.getCategory() != null ? cc.getCategory().trim() : "";
                    int d = cc.getReportDate().toLocalDate().getDayOfMonth();
                    dataMap.computeIfAbsent(cat, k -> new java.util.HashMap<>()).put(d, cc.getTotalComplaints());
                }

                // Ghi dữ liệu vào các ô từ H8 đến cột ứng với ngày hiện tại
                for (int rowIdx = startRow; rowIdx <= endRow; rowIdx++) {
                    String cat = categories.get(rowIdx - startRow);
                    Row row = sheet.getRow(rowIdx);
                    if (row == null) row = sheet.createRow(rowIdx);
                    for (int d = 1; d <= currentDay; d++) {
                        int colIdx = 7 + d - 1; // H là 7
                        Cell cell = row.getCell(colIdx, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                        int value = 0;
                        if (dataMap.containsKey(cat) && dataMap.get(cat).containsKey(d)) {
                            value = dataMap.get(cat).get(d);
                        }
                        cell.setCellValue(value);
                    }
                }

                // === Tính tổng từng ngày cho hàng tổng (row 6, tức là dòng 7 Excel) ===
                int totalRowIdx = 6; // Hàng tổng là dòng 7 (index 6)
                Row totalRow = sheet.getRow(totalRowIdx);
                if (totalRow == null) totalRow = sheet.createRow(totalRowIdx);
                for (int d = 1; d <= currentDay; d++) {
                    int colIdx = 7 + d - 1; // H là 7
                    int sum = 0;
                    for (int rowIdx = startRow; rowIdx <= endRow; rowIdx++) {
                        Row row = sheet.getRow(rowIdx);
                        if (row != null) {
                            Cell cell = row.getCell(colIdx);
                            if (cell != null && cell.getCellType() == CellType.NUMERIC) {
                                sum += (int) cell.getNumericCellValue();
                            }
                        }
                    }
                    Cell totalCell = totalRow.getCell(colIdx, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    totalCell.setCellValue(sum);
                }
                ByteArrayOutputStream out = new ByteArrayOutputStream();
                workbook.write(out);
                return out.toByteArray();
            }
        } catch (IOException e) {
            throw new RuntimeException("Lỗi đọc/ghi file Excel", e);
        }
    }

}
