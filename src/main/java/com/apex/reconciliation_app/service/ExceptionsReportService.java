package com.apex.reconciliation_app.service;

import dto.ExceptionRecord;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;

@Service
public class ExceptionsReportService {

    public ByteArrayInputStream generateReport(List<ExceptionRecord> exceptions) {

        try (Workbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream out = new ByteArrayOutputStream()) {

            Sheet sheet = workbook.createSheet("Exceptions Report");

            // --- CREATE AND STYLE HEADER ROW ---
            Row headerRow = sheet.createRow(0);
            String[] headers = {"Site Order ID",
                                "SKU",
                                "Amount",
                                "Transaction Description",
                                "Error Reason"};

            CellStyle headerStyle = workbook.createCellStyle();
            Font headerFont = workbook.createFont();
            headerFont.setBold(true);
            headerStyle.setFont(headerFont);

            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
                cell.setCellStyle(headerStyle);
            }

            // --- POPULATE DATA ROWS ---
            int rowIdx = 1;
            for (ExceptionRecord record : exceptions) {
                Row row = sheet.createRow(rowIdx++);

                row.createCell(0).setCellValue(record.siteOrderId());
                row.createCell(1).setCellValue(record.sku());
                row.createCell(2).setCellValue(record.amount());
                row.createCell(3).setCellValue(record.transactionDate() != null ? record.transactionDate() : "");
                row.createCell(4).setCellValue(record.errorReason());
            }

            // --- 3. AUTO-SIZE COLUMNS ---
            for (int i = 0; i < headers.length; i++) {
                sheet.autoSizeColumn(i);
            }

            // --- 4. WRITE TO STREAM ---
            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());

        } catch (IOException e) {
            throw new RuntimeException("Failed to generate Exceptions Report: " + e.getMessage());
        }
    }
}
