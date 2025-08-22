package app.sme.repo;

import app.sme.projection.NormalSheetDTO;
import app.sme.entity.Complaint;
import app.sme.projection.DocxProjection;
import app.sme.projection.ExcelChartSheetProjection;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;

import java.util.List;

public interface ComplaintRepository extends JpaRepository<Complaint, Long> {
    @Query(value = "SELECT * FROM optimized_complaint_view", nativeQuery = true)
    List<NormalSheetDTO> findTotalView();

    @Query(value = "SELECT * from previous_month_complaint_view", nativeQuery = true)
    List<ExcelChartSheetProjection> findErrorComplaintsLastMonth();

    @Query(
            value = "SELECT dimension_type, dimension_value, total_complaints " +
                    "FROM vw_complaint_summary",
            nativeQuery = true
    )
    List<ExcelChartSheetProjection> getComplaintSummary();

    @Query(value = "SELECT * FROM previous_month_complaint_view", nativeQuery = true)
    List<DocxProjection> findAllPreviousMonth();

}