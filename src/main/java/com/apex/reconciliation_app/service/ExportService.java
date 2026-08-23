package com.apex.reconciliation_app.service;

import com.apex.reconciliation_app.model.ReconciliationRecord;
import com.apex.reconciliation_app.repository.ReconciliationRepository;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.List;

@Service
@RequiredArgsConstructor
public class ExportService {

    private final ReconciliationRepository repository;

    public ByteArrayInputStream exportToExcel() {
        // 1. Grab all the saved data from the db
        List<ReconciliationRecord> records = repository.findAll();

        // 2. Create a new workbook and sheet in memory
        try (Workbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream out = new ByteArrayOutputStream())
        {
            Sheet sheet = workbook.createSheet("Reconciliation Data");

            // 3. Create the header row
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("Composite ID");
            headerRow.createCell(1).setCellValue("Site Name");
            headerRow.createCell(2).setCellValue("SKU");
            headerRow.createCell(3).setCellValue("Site Order ID");
            headerRow.createCell(4).setCellValue("Order Date");
            headerRow.createCell(5).setCellValue("Account");
            headerRow.createCell(6).setCellValue("Salesperson");
            headerRow.createCell(7).setCellValue("Total Less Tax");
            headerRow.createCell(8).setCellValue("Total Seller Cost");


            // 4. Loop through the database records and populate the rows
            int rowIndex = 1;
            for (ReconciliationRecord record : records) {
                Row row = sheet.createRow(rowIndex++);

                // String cells
                row.createCell(0).setCellValue(record.getCompositeId() != null ? record.getCompositeId() : "");
                row.createCell(1).setCellValue(record.getSiteName() != null ? record.getSiteName() : "");
                row.createCell(2).setCellValue(record.getSku() != null ? record.getSku() : "");
                row.createCell(3).setCellValue(record.getSiteOrderId() != null ? record.getSiteOrderId() : "");
                row.createCell(4).setCellValue(record.getOrderDate() != null ? record.getOrderDate() : "");
                row.createCell(5).setCellValue(record.getAccount() != null ? record.getAccount() : "");
                row.createCell(6).setCellValue(record.getSalesperson() != null ? record.getSalesperson() : "");

                // Numeric cells
                row.createCell(7).setCellValue(record.getTotalLessTax() != null ? record.getTotalLessTax() : 0.0);
                row.createCell(8).setCellValue(record.getTotalSellerCost() != null ? record.getTotalSellerCost() : 0.0);
            }

            // 5. Write the workbook to our byte stream and return it.
            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());

        } catch (Exception e) {
            throw new RuntimeException("Failed to export data to Excel: " + e.getMessage());
        }

    }
}
