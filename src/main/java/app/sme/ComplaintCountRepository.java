package app.sme;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;
import java.sql.Date;
import java.util.List;

public interface ComplaintCountRepository extends JpaRepository<ComplaintCount, Long> {
    @Query("SELECT c FROM ComplaintCount c " +
            "WHERE c.reportDate BETWEEN :startDate AND :endDate " +
            "ORDER BY c.reportDate, c.category")
    List<ComplaintCount> findByDateRange(@Param("startDate") Date startDate,
                                         @Param("endDate") Date endDate);
}

