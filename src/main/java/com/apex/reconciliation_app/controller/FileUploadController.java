package com.apex.reconciliation_app.controller;

import com.apex.reconciliation_app.dto.MarketplaceParseResult;
import com.apex.reconciliation_app.model.AmazonRawTransaction;
import com.apex.reconciliation_app.model.AmazonSuspense;
import com.apex.reconciliation_app.model.WalmartRawTransaction;
import com.apex.reconciliation_app.model.WalmartSuspense;
import com.apex.reconciliation_app.service.*;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.util.HashSet;
import java.util.Set;

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


    @PostMapping("/upload")
    public ResponseEntity<?> uploadUniversalFile(@RequestParam("file") MultipartFile file) {
        if (file.isEmpty()) {
            return ResponseEntity.status(HttpStatus.BAD_REQUEST).body("Please select a file to upload.");
        }

        try {
            byte[] fileBytes = file.getBytes();
            String fileType = detectFileType(fileBytes);
            switch (fileType) {
                case "WALMART" -> {
                    MarketplaceParseResult<WalmartSuspense, WalmartRawTransaction> result =
                            walmartParserService.parseAndUpdate(new ByteArrayInputStream(fileBytes));
                    return buildReceiptResponse(walmartReportService.generateReport(result), "walmart_upload_receipt.xlsx");
                }
                case "AMAZON" -> {
                    MarketplaceParseResult<AmazonSuspense, AmazonRawTransaction> result =
                            amazonParserService.parseAndUpdate(new ByteArrayInputStream(fileBytes));
                    return buildReceiptResponse(amazonReportService.generateReport(result), "amazon_upload_receipt.xlsx");
                }
                case "RITHUM" -> {
                    rithumParserService.parseAndSaveInputStream(new ByteArrayInputStream(fileBytes));
                    return ResponseEntity.ok("RITHUM: Master base data uploaded and initialized successfully.");
                }
                default -> {
                    return ResponseEntity.status(HttpStatus.BAD_REQUEST)
                            .body("ERROR: Unknown file format. Could not match headers to Rithum, Walmart, or Amazon signatures.");
                }
            }
        } catch (Exception e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                    .body("Could not process the file: " + e.getMessage());
        }
    }


    @GetMapping("/export")
    public ResponseEntity<InputStreamResource> downloadReport() {

        return buildReceiptResponse(exportService.exportToExcel(), "reconciliation_master_report.xlsx");
    }

    // Helper Methods
    private String detectFileType(byte[] fileBytes) throws Exception {
        try (InputStream is = new ByteArrayInputStream(fileBytes);
             Workbook workbook = new XSSFWorkbook(is)) {

            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);
            if (headerRow == null) return "UNKNOWN";

            Set<String> headers = new HashSet<>();
            for (Cell cell : headerRow) {
                if (cell != null && cell.getCellType() == CellType.STRING) {
                    headers.add(cell.getStringCellValue().trim().toLowerCase());
                }
            }

            if (headers.contains("transaction key") && headers.contains("partner item id")) {
                return "WALMART";
            }
            if (headers.contains("settlement id") && headers.contains("order id")) {
                return "AMAZON";
            }
            if (headers.contains("secondary site order id") && headers.contains("channeladvisor order id")){
                return "RITHUM";
            }

            return "UNKNOWN";
        }
    }

    private ResponseEntity<InputStreamResource> buildReceiptResponse(ByteArrayInputStream stream, String filename) {
        InputStreamResource resource = new InputStreamResource(stream);
        HttpHeaders headers = new HttpHeaders();
        headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + filename);

        return ResponseEntity.ok()
                .headers(headers)
                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                .body(resource);
    }
}
