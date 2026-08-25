package com.apex.reconciliation_app.service;

import com.apex.reconciliation_app.model.ReconciliationRecord;
import com.apex.reconciliation_app.repository.ReconciliationRepository;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.InputStream;
import java.util.*;

@Service
@RequiredArgsConstructor
public class WalmartParserService {

    private final ReconciliationRepository repository;

    public void parseAndUpdate(InputStream inputStream) {
        try (Workbook workbook = new XSSFWorkbook(inputStream)) {
            Sheet sheet = workbook.getSheetAt(0);
            Map<String, ReconciliationRecord> recordsToUpdate = new HashMap<>();

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                // 1. Map based on Walmart's structure
                String transactionDesc = getStringValue(row.getCell(3)).trim().toUpperCase();   // Column D
                String siteOrderId = getStringValue(row.getCell(6)).trim().toUpperCase();       // Column G (Purchase Order #)
                double amount = getDoubleValue(row.getCell(8)) * -1;                            // Column I (Amount)
                String amountType = getStringValue(row.getCell(9)).trim().toUpperCase();        // Column J (Amount Type)
                String sku = getStringValue(row.getCell(14)).trim().toUpperCase();              // Column O (Partner Item Id)

                if (siteOrderId.isEmpty() || sku.isEmpty()) continue;
                String compositeId = siteOrderId + "-" + sku;

                // 2. Fetch record from map or pull from DB if we haven't seen it yet
                ReconciliationRecord record = recordsToUpdate.get(compositeId);
                if (record == null) {
                    Optional<ReconciliationRecord> dbRecord = repository.findById(compositeId);
                    if (dbRecord.isPresent()){
                        record = dbRecord.get();
                        recordsToUpdate.put(compositeId, record);
                    } else {
                        // If it doesn't exist in DB from rithum upload
                        continue;
                    }
                }

                // 3. Route the amount to the correct field based on ledger hierarchy
                switch (transactionDesc) {
                    case ("PURCHASE") -> {
                        switch (amountType) {
                            case ("PRODUCT PRICE") -> record.setSiteOrderAmount((record.getSiteOrderAmount() != null ? record.getSiteOrderAmount() : 0.0) + (amount * -1));
                            case ("COMMISSION ON PRODUCT") -> record.setSiteOrderFee((record.getSiteOrderFee() != null ? record.getSiteOrderFee() : 0.0) + amount);
                            case ("TOTAL WALMART FUNDED SAVINGS") -> record.setSiteOrderOtherFees1((record.getSiteOrderOtherFees1() != null ? record.getSiteOrderOtherFees1() : 0.0) + amount);
                        }
                    }
                    case ("RETURN REFUND") -> {
                        switch (amountType) {
                            case ("PRODUCT PRICE") -> record.setAmountRefunded((record.getAmountRefunded() != null ? record.getAmountRefunded() : 0.0) + amount);
                            case ("COMMISSION ON PRODUCT") -> record.setReturnFee1((record.getReturnFee1() != null ? record.getReturnFee1() : 0.0) + amount);
                            case ("TOTAL WALMART FUNDED SAVINGS") -> record.setReturnFee2((record.getReturnFee2() != null ? record.getReturnFee2() : 0.0) + amount);
                        }
                    }
                    case ("WALMART RETURN SHIPPING CHARGE") -> record.setReturnShipping((record.getReturnShipping() != null ? record.getReturnShipping() : 0.0) + amount);
                }
            }
            // 4. Save all the updated records back to the database.
            repository.saveAll(recordsToUpdate.values());
            System.out.println("Successfully aggregated and updated " + recordsToUpdate.size() + " Walmart records.");

        } catch (Exception e) {
            throw new RuntimeException("Failed to parse Walmart Excel file: " + e.getMessage());
        }
    }

    // --- SAFE DATA EXTRACTION HELPERS ---
    // TODO: refactor into a shared 'ExcelUtils' because this is repeated in other Service
    private String getStringValue(Cell cell) {
        if (cell == null) return "";
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue().trim();
            case NUMERIC -> {
                if (DateUtil.isCellDateFormatted(cell)) {
                    java.text.SimpleDateFormat dateFormat = new java.text.SimpleDateFormat("MM/dd/yyyy HH:mm");
                    yield dateFormat.format(cell.getDateCellValue());
                } else {
                    yield String.valueOf((long) cell.getNumericCellValue());
                }
            }
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            default -> "";
        };
    }

    private Double getDoubleValue(Cell cell) {
        if (cell == null) return 0.0;
        return switch (cell.getCellType()) {
            case NUMERIC -> cell.getNumericCellValue();
            case STRING -> {
                try {
                    yield Double.parseDouble(cell.getStringCellValue().replace("$", "").trim());
                } catch (NumberFormatException e) {
                    yield 0.0;
                }
            }
            default -> 0.0;
        };
    }
}
