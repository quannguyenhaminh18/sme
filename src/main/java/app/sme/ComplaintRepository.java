package app.sme;

import org.springframework.data.repository.Repository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;

import java.sql.Date;
import java.util.List;

public interface ComplaintRepository extends Repository<Complaint, Long> {

    @Query("SELECT COUNT(c) FROM Complaint c WHERE LOWER(TRIM(c.category)) = LOWER(TRIM('Dịch vụ CA'))")
    Long countCategoryDichVuCA();

    @Query(value = "SELECT category, count_yesterday " +
            "FROM vw_complaints_stats ",
            nativeQuery = true)
    List<Object[]> findYesterdayCounts();

    @Query(value = "SELECT " +
            "c.category, " +
            "MAX(c.total_on_time_yesterday) AS total_on_time, " +
            "MAX(c.total_received_yesterday) AS total_received " +
            "FROM vw_complaints_stats c " +
            "GROUP BY c.category",
            nativeQuery = true)
    List<Object[]> findYesterdayOnTimeAndReceivedCounts();

}