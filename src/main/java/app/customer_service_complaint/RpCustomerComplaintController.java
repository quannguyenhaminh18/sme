package app.customer_service_complaint;

import lombok.AccessLevel;
import lombok.RequiredArgsConstructor;
import lombok.experimental.FieldDefaults;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("/customer-service-complaint")
@RequiredArgsConstructor
@FieldDefaults(level = AccessLevel.PRIVATE)
public class RpCustomerComplaintController {
    final RpCustomerComplaintService rpCustomerComplaintService;

    @GetMapping("/export-excel")
    public ResponseEntity<byte[]> exportExcelFile() {
        byte[] file = rpCustomerComplaintService.exportExcelFile();

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=bao_cao_phu_luc.xlsx")
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .body(file);
    }
}
