package app.sme;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDate;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

@Service
@RequiredArgsConstructor
@Slf4j
public class ComplaintService {
    private final ComplaintRepository complaintRepository;
//    private final ComplaintCountRepository complaintCountRepository;
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
//                //sheet1
//                Sheet sheet = getRequiredSheet(workbook, "SL_ngay");
//
//                // Lấy danh sách category từ cột B8 đến B39
//                int startRow = 7; // B8
//                int endRow = 38;  // B39
//                int categoryCol = 1; // Cột B
//                java.util.List<String> categories = new java.util.ArrayList<>();
//                for (int i = startRow; i <= endRow; i++) {
//                    Row row = sheet.getRow(i);
//                    if (row != null) {
//                        Cell cell = row.getCell(categoryCol);
//                        if (cell != null && cell.getCellType() == CellType.STRING) {
//                            categories.add(cell.getStringCellValue().trim());
//                        } else {
//                            categories.add("");
//                        }
//                    } else {
//                        categories.add("");
//                    }
//                }
//
//                // TEST: Override ngày hệ thống thành 2-8-2025 để kiểm tra logic xuất file
//                LocalDate today = LocalDate.of(2025, 8, 3);
//                LocalDate endDate = today.minusDays(1);
//                int currentDay = endDate.getDayOfMonth();
//                LocalDate firstDay = today.withDayOfMonth(1);
//                java.sql.Date sqlStart = java.sql.Date.valueOf(firstDay);
//                java.sql.Date sqlEnd = java.sql.Date.valueOf(endDate);
//
//                // Lấy dữ liệu count từ repository
//                java.util.List<ComplaintCount> counts = complaintCountRepository.findByDateRange(sqlStart, sqlEnd);
//                // Map: category -> (date -> count)
//                java.util.Map<String, java.util.Map<Integer, Integer>> dataMap = new java.util.HashMap<>();
//                for (ComplaintCount cc : counts) {
//                    String cat = cc.getCategory() != null ? cc.getCategory().trim() : "";
//                    int d = cc.getReportDate().toLocalDate().getDayOfMonth();
//                    dataMap.computeIfAbsent(cat, k -> new java.util.HashMap<>()).put(d, cc.getTotalComplaints());
//                }
//
//                // Ghi dữ liệu vào các ô từ H8 đến cột ứng với ngày hiện tại
//                for (int rowIdx = startRow; rowIdx <= endRow; rowIdx++) {
//                    String cat = categories.get(rowIdx - startRow);
//                    Row row = sheet.getRow(rowIdx);
//                    if (row == null) row = sheet.createRow(rowIdx);
//                    for (int d = 1; d <= currentDay; d++) {
//                        int colIdx = 7 + d - 1; // H là 7
//                        Cell cell = row.getCell(colIdx, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
//                        int value = 0;
//                        if (dataMap.containsKey(cat) && dataMap.get(cat).containsKey(d)) {
//                            value = dataMap.get(cat).get(d);
//                        }
//                        cell.setCellValue(value);
//                    }
//                }
//
//                // === Tính tổng từng ngày cho hàng tổng (row 6, tức là dòng 7 Excel) ===
//                int totalRowIdx = 6; // Hàng tổng là dòng 7 (index 6)
//                Row totalRow = sheet.getRow(totalRowIdx);
//                if (totalRow == null) totalRow = sheet.createRow(totalRowIdx);
//                for (int d = 1; d <= currentDay; d++) {
//                    int colIdx = 7 + d - 1; // H là 7
//                    int sum = 0;
//                    for (int rowIdx = startRow; rowIdx <= endRow; rowIdx++) {
//                        Row row = sheet.getRow(rowIdx);
//                        if (row != null) {
//                            Cell cell = row.getCell(colIdx);
//                            if (cell != null && cell.getCellType() == CellType.NUMERIC) {
//                                sum += (int) cell.getNumericCellValue();
//                            }
//                        }
//                    }
//                    Cell totalCell = totalRow.getCell(colIdx, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
//                    totalCell.setCellValue(sum);
//                }

                Sheet sheet2 = getRequiredSheet(workbook, "KPIs_ngay");

                // Map service -> row index cho countYesterday
                Map<String, Integer> serviceRowMap = new HashMap<>();
                serviceRowMap.put("Dịch vụ CA", 7);
                serviceRowMap.put("Dịch vụ BHXH", 12);
                serviceRowMap.put("Dịch vụ Hóa đơn điện tử", 17);
                serviceRowMap.put("Dịch vụ vTracking", 22);

                // Map service -> vị trí row cho total on time / total received
                Map<String, int[]> serviceTotalRowMap = new HashMap<>();
                serviceTotalRowMap.put("Dịch vụ CA", new int[]{28, 29});
                serviceTotalRowMap.put("Dịch vụ BHXH", new int[]{33, 34});
                serviceTotalRowMap.put("Dịch vụ Hóa đơn điện tử", new int[]{38, 39});
                serviceTotalRowMap.put("Dịch vụ vTracking", new int[]{43, 44});

                // Normalize key -> lower case
                Map<String, Integer> normalizedServiceRowMap = new HashMap<>();
                for (Map.Entry<String, Integer> e : serviceRowMap.entrySet()) {
                    normalizedServiceRowMap.put(e.getKey().toLowerCase(), e.getValue());
                }

                Map<String, int[]> normalizedServiceTotalRowMap = new HashMap<>();
                for (Map.Entry<String, int[]> e : serviceTotalRowMap.entrySet()) {
                    normalizedServiceTotalRowMap.put(e.getKey().toLowerCase(), e.getValue());
                }

                // ================= LẤY DỮ LIỆU countYesterday =================
                List<Object[]> stats = complaintRepository.findYesterdayCounts();
                Map<String, Integer> countMap = new HashMap<>();
                for (Object[] row : stats) {
                    if (row[0] == null || row[1] == null) continue;
                    String service = row[0].toString().trim().toLowerCase();
                    int countYesterday = ((Number) row[1]).intValue();
                    countMap.put(service, countYesterday);
                }

                // Ghi countYesterday vào sheet2
                for (Map.Entry<String, Integer> entry : normalizedServiceRowMap.entrySet()) {
                    String serviceName = entry.getKey();
                    int totalRowIdx2 = entry.getValue();

                    Row row = sheet2.getRow(totalRowIdx2);
                    if (row == null) row = sheet2.createRow(totalRowIdx2);

                    int today = LocalDate.now().getDayOfMonth();
                    int colIdx = 4 + (today - 2);

                    Cell cell = row.getCell(colIdx, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    double oldValue = cell.getCellType() == CellType.NUMERIC ? cell.getNumericCellValue() : 0;
                    int countYesterday = countMap.getOrDefault(serviceName, 0);
                    cell.setCellValue(oldValue + countYesterday);
                }

                List<Object[]> totalStats = complaintRepository.findYesterdayOnTimeAndReceivedCounts();
                Map<String, Integer[]> totalMap = new HashMap<>();
                for (Object[] row : totalStats) {
                    if (row[0] == null || row[1] == null || row[2] == null) continue;
                    String service = row[0].toString().trim().toLowerCase();
                    int totalOnTime = ((Number) row[1]).intValue();
                    int totalReceived = ((Number) row[2]).intValue();
                    totalMap.put(service, new Integer[]{totalOnTime, totalReceived});
                }

                // Ghi total_on_time_yesterday và total_received_yesterday
                for (Map.Entry<String, int[]> entry : normalizedServiceTotalRowMap.entrySet()) {
                    String serviceName = entry.getKey();
                    int[] rows = entry.getValue(); // [rowOnTime, rowReceived]

                    Integer[] values = totalMap.getOrDefault(serviceName, new Integer[]{0, 0});
                    int totalOnTime = values[0];
                    int totalReceived = values[1];

                    int today = LocalDate.now().getDayOfMonth();
                    int colIdx = 4 + (today - 2);

                    // total on time
                    Row rowOnTime = sheet2.getRow(rows[0]);
                    if (rowOnTime == null) rowOnTime = sheet2.createRow(rows[0]);
                    Cell cellOnTime = rowOnTime.getCell(colIdx, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    cellOnTime.setCellValue(totalOnTime);

                    // total received
                    Row rowReceived = sheet2.getRow(rows[1]);
                    if (rowReceived == null) rowReceived = sheet2.createRow(rows[1]);
                    Cell cellReceived = rowReceived.getCell(colIdx, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    cellReceived.setCellValue(totalReceived);
                }

                ByteArrayOutputStream out = new ByteArrayOutputStream();
                workbook.write(out);
                System.out.println("Dữ liệu đã được ghi");
                return out.toByteArray();
            }
        } catch (IOException e) {
            throw new RuntimeException("Lỗi đọc/ghi file Excel", e);
        }
    }

}
