package com.apex.reconciliation_app.controller;

import com.apex.reconciliation_app.dto.MarketplaceParseResult;
import com.apex.reconciliation_app.model.AmazonRawTransaction;
import com.apex.reconciliation_app.model.AmazonSuspense;
import com.apex.reconciliation_app.model.WalmartRawTransaction;
import com.apex.reconciliation_app.model.WalmartSuspense;
import com.apex.reconciliation_app.service.*;
import lombok.RequiredArgsConstructor;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Date;

@RestController
@RequestMapping("/api/reconciliation")
@RequiredArgsConstructor
public class FileUploadController {

    private final RithumParserService rithumParserService;
    private final ExportService exportService;

    private final WalmartParserService walmartParserService;
    private final WalmartReportService walmartReportService;

    private final AmazonParserService amazonParserService;
    private final AmazonReportService amazonReportService;

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
            MarketplaceParseResult<WalmartSuspense, WalmartRawTransaction> result = walmartParserService.parseAndUpdate(file.getInputStream());

            InputStreamResource resource = new InputStreamResource(walmartReportService.generateReport(result));

            HttpHeaders headers = new HttpHeaders();
            headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=walmart_upload_receipt.xlsx");

            return ResponseEntity.ok()
                    .headers(headers)
                    .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                    .body(resource);

        } catch (Exception e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                    .body("Could not process the file: " + e.getMessage());
        }
    }

    @PostMapping("/amazon")
    public ResponseEntity<?> uploadAmazonFile(@RequestParam("file") MultipartFile file) {
        if (file.isEmpty()) {
            return ResponseEntity.status(HttpStatus.BAD_REQUEST).body("Please select a file to upload.");
        }

        try {
            MarketplaceParseResult<AmazonSuspense, AmazonRawTransaction> result = amazonParserService.parseAndUpdate(file.getInputStream());
            InputStreamResource resource = new InputStreamResource(amazonReportService.generateReport(result));

            HttpHeaders headers = new HttpHeaders();
            headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=amazon_upload_receipt.xlsx");

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
