package app.customer_service_complaint;

import lombok.AccessLevel;
import lombok.RequiredArgsConstructor;
import lombok.experimental.FieldDefaults;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.time.LocalDate;
import java.time.YearMonth;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Function;
import java.util.stream.Collectors;

@Service
@FieldDefaults(level = AccessLevel.PRIVATE)
@Slf4j
@RequiredArgsConstructor
public class RpCustomerComplaintService {
    final RpCustomerComplaintRepository repository;
    static final int MONTH_ROW = 2;
    static final int MONTH_CELL = 3;
    static final int MONTH_CELL_DETAIL = 4;
    static final int PIVOT_FIRST_COL = 6;


    public byte[] exportExcelFile() {
        //Summary data
        List<ICustomerComplaintSummary> summaries = repository.reportCustomerComplaintSummary(null);

        //Pivot data
        List<ICustomerComplaintPivot> pivots = repository.reportCustomerComplaintPivot(null);

        ClassPathResource template = new ClassPathResource("templates/excel/bao_cao_phu_luc.xlsx");
        try (InputStream is = template.getInputStream();
             Workbook wb = new XSSFWorkbook(is);
             ByteArrayOutputStream out = new ByteArrayOutputStream()) {

            Sheet sheet1 = wb.getSheetAt(0);

            // Tính tháng hiện tại
            LocalDate now = LocalDate.now();
            YearMonth yearMonth = YearMonth.of(now.getYear(), now.getMonth());
            String monthTitle = "T" + now.getMonthValue() + "." + now.getYear();
            String monthTitleDetail = monthTitle + "_" + "Chi tiết ngày";

            //Ghi tiêu đề tháng hiện tại
            Row monthRow = sheet1.getRow(MONTH_ROW);
            Cell montCell = monthRow.getCell(MONTH_CELL);
            montCell.setCellValue(monthTitle);
            Cell montCellDetail = monthRow.getCell(MONTH_CELL_DETAIL);
            montCellDetail.setCellValue(monthTitleDetail);

            int daysInMonth = yearMonth.lengthOfMonth();


            wb.write(out);
            return out.toByteArray();
        } catch (Exception e) {
            throw new RuntimeException("Export Excel failed", e);
        }
    }
}
