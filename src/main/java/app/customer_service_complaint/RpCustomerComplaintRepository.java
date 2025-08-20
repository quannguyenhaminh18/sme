package app.customer_service_complaint;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;
import org.springframework.stereotype.Repository;

import java.util.List;

@Repository
public interface RpCustomerComplaintRepository extends JpaRepository<RpCustomerComplaint, Long> {


    @Query(value = "with total_complaint as (select count(*) as total\n" +
            "                         from rp_customer_complaints\n" +
            "                         where received_date >= DATE_FORMAT(CURDATE(), '%Y-%m-01')\n" +
            "                           and received_date < CURDATE()\n" +
            "                           AND (:isCancel is null or (:isCancel = true and complaint_status = 'Hủy phản ánh')))\n" +
            "select parent.name                     as parentName,\n" +
            "       rsc.name                        as serviceName,\n" +
            "       c.name                          as categoryName,\n" +
            "       count(*)                        as luyKe,\n" +
            "       count(*) / (day(curdate()) - 1) as trungBinhNgay,\n" +
            "       count(*) / total                as tiTrong\n" +
            "from rp_customer_complaints rcc\n" +
            "\n" +
            "         join rp_complaint_customer_category c on rcc.category = c.name\n" +
            "         join rp_customer_service_complaints rsc on c.service_id = rsc.id\n" +
            "         join rp_customer_service_complaints parent on parent.id = rsc.parent_id\n" +
            "         cross join total_complaint t\n" +
            "where rcc.received_date >= DATE_FORMAT(CURDATE(), '%Y-%m-01')\n" +
            "  and rcc.received_date < CURDATE()\n" +
            "  AND (:isCancel is null or (:isCancel = true and complaint_status = 'Hủy phản ánh'))\n" +
            "group by parent.name, rsc.name\n" +
            "order by parent.order_no, rsc.order_no", nativeQuery = true)
    List<ICustomerComplaintSummary> reportCustomerComplaintSummary(@Param("isCancel") Boolean isCancel);


    @Query(value = "SELECT parent.name                                AS parentName,\n" +
            "       rsc.name                                   AS serviceName,\n" +
            "       c.name                                     AS categoryName,\n" +
            "       -- pivot 31 ngày\n" +
            "       SUM(IF(DAY(rcc.received_date) = 1, 1, 0))  AS d01,\n" +
            "       SUM(IF(DAY(rcc.received_date) = 2, 1, 0))  AS d02,\n" +
            "       SUM(IF(DAY(rcc.received_date) = 3, 1, 0))  AS d03,\n" +
            "       SUM(IF(DAY(rcc.received_date) = 4, 1, 0))  AS d04,\n" +
            "       SUM(IF(DAY(rcc.received_date) = 5, 1, 0))  AS d05,\n" +
            "       SUM(IF(DAY(rcc.received_date) = 6, 1, 0))  AS d06,\n" +
            "       SUM(IF(DAY(rcc.received_date) = 7, 1, 0))  AS d07,\n" +
            "       SUM(IF(DAY(rcc.received_date) = 8, 1, 0))  AS d08,\n" +
            "       SUM(IF(DAY(rcc.received_date) = 9, 1, 0))  AS d09,\n" +
            "       SUM(IF(DAY(rcc.received_date) = 10, 1, 0)) AS d10,\n" +
            "       SUM(IF(DAY(rcc.received_date) = 11, 1, 0)) AS d11,\n" +
            "       SUM(IF(DAY(rcc.received_date) = 12, 1, 0)) AS d12,\n" +
            "       SUM(IF(DAY(rcc.received_date) = 13, 1, 0)) AS d13,\n" +
            "       SUM(IF(DAY(rcc.received_date) = 14, 1, 0)) AS d14,\n" +
            "       SUM(IF(DAY(rcc.received_date) = 15, 1, 0)) AS d15,\n" +
            "       SUM(IF(DAY(rcc.received_date) = 16, 1, 0)) AS d16,\n" +
            "       SUM(IF(DAY(rcc.received_date) = 17, 1, 0)) AS d17,\n" +
            "       SUM(IF(DAY(rcc.received_date) = 18, 1, 0)) AS d18,\n" +
            "       SUM(IF(DAY(rcc.received_date) = 19, 1, 0)) AS d19,\n" +
            "       SUM(IF(DAY(rcc.received_date) = 20, 1, 0)) AS d20,\n" +
            "       SUM(IF(DAY(rcc.received_date) = 21, 1, 0)) AS d21,\n" +
            "       SUM(IF(DAY(rcc.received_date) = 22, 1, 0)) AS d22,\n" +
            "       SUM(IF(DAY(rcc.received_date) = 23, 1, 0)) AS d23,\n" +
            "       SUM(IF(DAY(rcc.received_date) = 24, 1, 0)) AS d24,\n" +
            "       SUM(IF(DAY(rcc.received_date) = 25, 1, 0)) AS d25,\n" +
            "       SUM(IF(DAY(rcc.received_date) = 26, 1, 0)) AS d26,\n" +
            "       SUM(IF(DAY(rcc.received_date) = 27, 1, 0)) AS d27,\n" +
            "       SUM(IF(DAY(rcc.received_date) = 28, 1, 0)) AS d28,\n" +
            "       SUM(IF(DAY(rcc.received_date) = 29, 1, 0)) AS d29,\n" +
            "       SUM(IF(DAY(rcc.received_date) = 30, 1, 0)) AS d30,\n" +
            "       SUM(IF(DAY(rcc.received_date) = 31, 1, 0)) AS d31\n" +
            "\n" +
            "FROM rp_customer_complaints rcc\n" +
            "         JOIN rp_complaint_customer_category c ON rcc.category = c.name\n" +
            "         JOIN rp_customer_service_complaints rsc ON c.service_id = rsc.id\n" +
            "         JOIN rp_customer_service_complaints parent ON parent.id = rsc.parent_id\n" +
            "WHERE rcc.received_date >= DATE_FORMAT(CURDATE(), '%Y-%m-01')\n" +
            "  AND rcc.received_date < CURDATE()\n" +
            "  AND (:isCancel IS NULL OR (:isCancel = TRUE AND complaint_status = 'Hủy phản ánh'))\n" +
            "GROUP BY parent.name, rsc.name, c.name\n" +
            "order by parent.order_no, rsc.order_no;\n", nativeQuery = true)
    List<ICustomerComplaintPivot> reportCustomerComplaintPivot(@Param("isCancel") Boolean isCancel);
}
