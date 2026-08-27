package com.apex.reconciliation_app.controller;

import com.apex.reconciliation_app.service.ExceptionsReportService;
import com.apex.reconciliation_app.service.ExportService;
import com.apex.reconciliation_app.service.RithumParserService;
import com.apex.reconciliation_app.service.WalmartParserService;
import dto.ExceptionRecord;
import lombok.RequiredArgsConstructor;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;


import java.util.List;

@RestController
@RequestMapping("/api/reconciliation")
@RequiredArgsConstructor
public class FileUploadController {

    private final RithumParserService rithumParserService;
    private final ExportService exportService;
    private final WalmartParserService walmartParserService;
    private final ExceptionsReportService exceptionsReportService;

    @PostMapping("/rithum")
    public ResponseEntity<?> uploadRithumFile(@RequestParam("file") MultipartFile file) {

        if (file.isEmpty()) {
            return ResponseEntity.status(HttpStatus.BAD_REQUEST).body("Please select a file to upload.");
        }

        try {
            rithumParserService.parseAndSaveInputStream(file.getInputStream());
            return ResponseEntity.ok("Rithum file uploaded and processed successfully!");
        } catch (Exception e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                    .body("Could not process the file: " + e.getMessage());
        }
    }

    @PostMapping("/walmart")
    public ResponseEntity<?> uploadWalmartFile(@RequestParam("file") MultipartFile file) {
        if (file.isEmpty()) {
            return ResponseEntity.status(HttpStatus.BAD_REQUEST).body("Please select a file to upload.");
        }

        try {
            List<ExceptionRecord> exceptions = walmartParserService.parseAndUpdate(file.getInputStream());

            // Scenario A: Perfect file
            if (exceptions.isEmpty()){
                return ResponseEntity.ok("Walmart file uploaded and processed successfully with 0 errors!");
            }

            // Scenario B: Errors found
            InputStreamResource resource = new InputStreamResource(exceptionsReportService.generateReport(exceptions));
            HttpHeaders headers = new HttpHeaders();
            headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=walmart_exceptions_report.xlsx");

            return ResponseEntity.ok()
                    .headers(headers)
                    .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                    .body(resource);

        } catch (Exception e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                    .body("Could not process the file: " + e.getMessage());
        }
    }

    @GetMapping("/export")
    public ResponseEntity<InputStreamResource> downloadReport() {

        InputStreamResource file = new InputStreamResource(exportService.exportToExcel());

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=reconciliation_master_report.xlsx")
                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                .body(file);
    }
}
