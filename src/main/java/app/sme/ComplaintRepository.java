package app.sme;

import org.springframework.data.repository.Repository;
import org.springframework.data.jpa.repository.Query;

public interface ComplaintRepository extends Repository<Complaint, Long> {

    @Query("SELECT COUNT(c) FROM Complaint c WHERE LOWER(TRIM(c.category)) = LOWER(TRIM('Dịch vụ CA'))")
    Long countCategoryDichVuCA();
}