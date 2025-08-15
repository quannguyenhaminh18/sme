package app.sme;

import lombok.Getter;
import lombok.Setter;

import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.Table;

@Entity
@Getter
@Setter
@Table(name = "complaint_counts")
public class ComplaintCount {
    @Id
    private String id; // hoặc kết hợp composite key

    private java.sql.Date reportDate;
    private String category;
    private int totalComplaints;

}
