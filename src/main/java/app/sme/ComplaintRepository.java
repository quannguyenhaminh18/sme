package app.sme;

import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.Repository;

import java.util.List;

public interface ComplaintRepository extends Repository<Complaint, Long> {

    @Query(value = "WITH cte AS (SELECT c.category                                      AS name,\n" +
            "                    c.kpi                                           AS kpi,\n" +
            "                    c.total_subscriber                              AS totalSubscriber,\n" +
            "\n" +
            "                    -- dữ liệu trong ngày\n" +
            "                    COALESCE(c.count_yesterday, 0)                  AS totalComplaintDay,\n" +
            "                    ROUND(\n" +
            "                            IF(c.count_yesterday IS NULL OR c.count_yesterday = 0,\n" +
            "                               0,\n" +
            "                               c.count_yesterday / c.total_subscriber * 10000\n" +
            "                            ),\n" +
            "                            2\n" +
            "                    )                                               AS complaintRateDay,\n" +
            "\n" +
            "                    -- tổng dữ liệu của tháng\n" +
            "                    COALESCE(c.total_this_month_until_yesterday, 0) AS totalComplaintMonth,\n" +
            "                    ROUND(\n" +
            "                            IF(c.total_this_month_until_yesterday IS NULL OR c.total_this_month_until_yesterday = 0,\n" +
            "                               0,\n" +
            "                               c.total_this_month_until_yesterday / c.total_subscriber * 10000\n" +
            "                            ),\n" +
            "                            2\n" +
            "                    )                                               AS complaintRateMonth,\n" +
            "                    -- tổng dữ liệu tháng trước\n" +
            "                    COALESCE(c.total_last_month, 0)                 AS totalLastMonth,\n" +
            "                    ROUND(\n" +
            "                            IF(c.total_last_month IS NULL OR c.total_last_month = 0,\n" +
            "                               0,\n" +
            "                               c.total_last_month / (c.total_subscriber - c.total_this_month_until_yesterday) * 10000\n" +
            "                            ),\n" +
            "                            2\n" +
            "                    )                                               AS complaintRateLastMonth\n" +
            "             FROM vw_complaint_summary_filtered c\n" +
            "             GROUP BY c.category)\n" +
            "SELECT *,\n" +
            "       IF(cte.complaintRateDay < cte.kpi, 'Đạt', 'Không đạt')                             AS kpiDay,\n" +
            "       IF(cte.complaintRateMonth < cte.kpi, 'Đạt', 'Không đạt')                           AS kpiMonth,\n" +
            "       (cte.complaintRateMonth - cte.complaintRateLastMonth) / cte.complaintRateLastMonth as compareRate\n" +
            "FROM cte;\n", nativeQuery = true)
    List<DocxProjection> reportServiceQuality();

    @Query(
            value = "SELECT category, count_yesterday FROM vw_complaints_stats ",
            nativeQuery = true
    )
    List<Object[]> findYesterdayCounts();

    @Query(value = "SELECT " +
            "c.category, " +
            "MAX(c.total_on_time_yesterday) AS total_on_time, " +
            "MAX(c.total_received_yesterday) AS total_received " +
            "FROM vw_complaints_stats c " +
            "GROUP BY c.category",
            nativeQuery = true)
    List<Object[]> findYesterdayOnTimeAndReceivedCounts();

    @Query(value =
            "SELECT " +
                    " c.category AS category, " +
                    " SUM(CASE WHEN c.received_date = CURDATE() - INTERVAL 1 DAY THEN 1 ELSE 0 END) AS countYesterday, " +
                    " SUM(CASE WHEN c.received_date BETWEEN DATE_FORMAT(CURDATE(), '%Y-%m-01') " +
                    "          AND CURDATE() - INTERVAL 1 DAY THEN 1 ELSE 0 END) AS countThisMonth " +
                    "FROM vw_complaint_summary_filtered c " +
                    "GROUP BY c.category",
            nativeQuery = true)
    List<SLNgayProjection> findCountsYesterday();
}