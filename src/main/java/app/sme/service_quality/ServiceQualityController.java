package app.sme.service_quality;

import app.sme.ComplaintService;
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
@RequestMapping("/sme/service-quality")
@RequiredArgsConstructor
@FieldDefaults(level = AccessLevel.PRIVATE)
public class ServiceQualityController {
    final ComplaintService complaintService;

    @GetMapping("/export-doc")
    public ResponseEntity<byte[]> exportDoc() {
        byte[] file = complaintService.exportServiceQualityDoc();
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=bao-cao-sme.docx")
                .contentType(MediaType.parseMediaType(
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"))
                .body(file);

    }
}
