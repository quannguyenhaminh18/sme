package app.sme;

import lombok.RequiredArgsConstructor;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/sme/export")
@RequiredArgsConstructor
public class ComplaintController {
    private final ComplaintService complaintService;

    @GetMapping("excel")
    public ResponseEntity<byte[]> exportExcelFile() {
        byte[] updatedFile = complaintService.exportExcelFile();
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=updated.xlsx")
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .body(updatedFile);
    }

    @GetMapping("doc")
    public ResponseEntity<byte[]> exportDoc() {
        byte[] file = complaintService.exportDoc();
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=updated.docx")
                .contentType(MediaType.parseMediaType(
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"))
                .body(file);

    }
}
