package app.sme;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/sme/export")
public class ComplaintController {
    @Autowired
    private ComplaintService complaintService;

    @GetMapping
    public ResponseEntity<byte[]> exportExcelFile() {
        byte[] updatedFile = complaintService.exportExcelFile();

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=updated.xlsx")
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .body(updatedFile);
    }
}
