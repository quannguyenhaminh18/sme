package app.sme.repo;

import app.sme.entity.Complaint;
import app.sme.projection.ExcelChartSheetProjection;
import app.sme.projection.NormalExcelSheetProjection;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;

import java.util.List;

public interface ComplaintRepository extends JpaRepository<Complaint, Long> {
    @Query(value = "SELECT * FROM vw_complaints_stats", nativeQuery = true)
    List<NormalExcelSheetProjection> findAllSMEView();

    @Query(value = "SELECT c.category AS category, " +
            "COUNT(0) AS totalComplaintsLastMonth " +
            "FROM complaints c " +
            "WHERE c.complaint_type LIKE '%lỗi%' " +
            "AND c.closing_date >= DATE_FORMAT(CURDATE(), '%Y-%m-01') - INTERVAL 1 MONTH " +
            "AND c.closing_date < DATE_FORMAT(CURDATE(), '%Y-%m-01') " +
            "GROUP BY c.category",
            nativeQuery = true)
    List<ExcelChartSheetProjection> findErrorComplaintsLastMonth();

    @Query(
            value = "SELECT dimension_type, dimension_value, total_complaints " +
                    "FROM vw_complaint_summary",
            nativeQuery = true
    )
    List<ExcelChartSheetProjection> getComplaintSummary();

}