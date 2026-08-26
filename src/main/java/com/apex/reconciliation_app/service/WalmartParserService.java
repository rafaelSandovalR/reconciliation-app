package com.apex.reconciliation_app.service;

import com.apex.reconciliation_app.enums.WalmartColumn;
import com.apex.reconciliation_app.model.ReconciliationRecord;
import com.apex.reconciliation_app.repository.ReconciliationRepository;
import com.apex.reconciliation_app.util.ExcelUtils;
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
                String transactionDesc = ExcelUtils.getStringValue(row.getCell(WalmartColumn.TRANSACTION_DESC.getIndex())).trim().toUpperCase();    // Column D (Transaction Description)
                String siteOrderId = ExcelUtils.getStringValue(row.getCell(WalmartColumn.PURCHASE_ORDER.getIndex())).trim().toUpperCase();          // Column G (Purchase Order #)
                double amount = ExcelUtils.getDoubleValue(row.getCell(WalmartColumn.AMOUNT.getIndex())) * -1;                                       // Column I (Amount) INVERTED
                String amountType = ExcelUtils.getStringValue(row.getCell(WalmartColumn.AMOUNT_TYPE.getIndex())).trim().toUpperCase();              // Column J (Amount Type)
                String sku = ExcelUtils.getStringValue(row.getCell(WalmartColumn.SKU.getIndex())).trim().toUpperCase();                             // Column O (Partner Item Id)

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
                            case ("PROMO CODE") -> record.setSiteOrderOtherFees2((record.getSiteOrderOtherFees2() != null ? record.getSiteOrderOtherFees2() : 0.0) + amount);
                            case ("OTHER TAX (FEES)") -> record.setSiteOrderOtherFees3((record.getSiteOrderOtherFees3() != null ? record.getSiteOrderOtherFees3() : 0.0) + amount);
                        }
                    }
                    case ("RETURN REFUND") -> {
                        record.setReturnStatus("Yes");

                        switch (amountType) {
                            case ("PRODUCT PRICE") -> record.setAmountRefunded((record.getAmountRefunded() != null ? record.getAmountRefunded() : 0.0) + amount);
                            case ("COMMISSION ON PRODUCT") -> record.setReturnFee1((record.getReturnFee1() != null ? record.getReturnFee1() : 0.0) + amount);
                            case ("TOTAL WALMART FUNDED SAVINGS") -> record.setReturnFee2((record.getReturnFee2() != null ? record.getReturnFee2() : 0.0) + amount);
                        }
                    }
                    case ("WALMART RETURN SHIPPING CHARGE") -> record.setReturnShipping((record.getReturnShipping() != null ? record.getReturnShipping() : 0.0) + amount);
                    case ("CUSTOMER RETURN REVERSAL") -> record.setReturnFee2((record.getReturnFee2() != null ? record.getReturnFee2() : 0.0) + amount);
                    case ("WALMART FAILED RETURN DELIVERY PROCESS CHARGE") -> record.setReturnFee3((record.getReturnFee3() != null ? record.getReturnFee3() : 0.0) + amount);
                }
            }

            // POST PROCESSING PASS
            for (ReconciliationRecord record : recordsToUpdate.values()) {
                if (record.getReturnFee1() != null && record.getReturnFee1() < 0) {
                    double originalCommission = record.getSiteOrderFee() != null ? record.getSiteOrderFee() : 0.0;
                    double unrefundedDifference = originalCommission + record.getReturnFee1();
                    record.setReturnFee1(unrefundedDifference);
                }
            }

            // 4. Save all the updated records back to the database.
            repository.saveAll(recordsToUpdate.values());
            System.out.println("Successfully aggregated and updated " + recordsToUpdate.size() + " Walmart records.");

        } catch (Exception e) {
            throw new RuntimeException("Failed to parse Walmart Excel file: " + e.getMessage());
        }
    }
}
