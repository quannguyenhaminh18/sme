package app.sme.service_quality;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface ComplainStatisticRepository extends JpaRepository<ComplaintsStatistic, String> {
}
