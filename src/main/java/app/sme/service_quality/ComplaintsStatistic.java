package app.sme.service_quality;

import lombok.Getter;
import lombok.Setter;
import org.hibernate.annotations.Immutable;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.Table;
import java.math.BigDecimal;

/**
 * Mapping for DB view
 */
@Getter
@Setter
@Entity
@Immutable
@Table(name = "complaints_statistics", schema = "sme_report")
public class ComplaintsStatistic {
    @Id
    @Column(name = "period", length = 10)
    private String period;

    @Column(name = "period_type", nullable = false, length = 5)
    private String periodType;

    @Column(name = "complaint_count", nullable = false)
    private Long complaintCount;

    @Column(name = "complaint_ratio", precision = 26, scale = 2)
    private BigDecimal complaintRatio;

}