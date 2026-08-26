package com.apex.reconciliation_app.service;

import com.apex.reconciliation_app.enums.WalmartColumn;
import com.apex.reconciliation_app.exception.BucketOverflowException;
import com.apex.reconciliation_app.model.ReconciliationRecord;
import com.apex.reconciliation_app.repository.ReconciliationRepository;
import com.apex.reconciliation_app.util.ExcelUtils;
import dto.ExceptionRecord;
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

    public List<ExceptionRecord> parseAndUpdate(InputStream inputStream) {
        try (Workbook workbook = new XSSFWorkbook(inputStream)) {
            Sheet sheet = workbook.getSheetAt(0);

            // --- DYNAMIC HEADER MAPPING ---
            Map<String, Integer> headerMap = ExcelUtils.getHeaderMap(sheet.getRow(0));

            Integer typeIndex = headerMap.get(WalmartColumn.TRANSACTION_TYPE.getHeaderName().toUpperCase());
            Integer descIndex = headerMap.get(WalmartColumn.TRANSACTION_DESC.getHeaderName().toUpperCase());
            Integer poIndex = headerMap.get(WalmartColumn.PURCHASE_ORDER.getHeaderName().toUpperCase());
            Integer amountIndex = headerMap.get(WalmartColumn.AMOUNT.getHeaderName().toUpperCase());
            Integer amountTypeIndex = headerMap.get(WalmartColumn.AMOUNT_TYPE.getHeaderName().toUpperCase());
            Integer skuIndex = headerMap.get(WalmartColumn.SKU.getHeaderName().toUpperCase());

            if (descIndex == null || poIndex == null || amountIndex == null || amountTypeIndex == null || skuIndex == null) {
                throw new RuntimeException("Missing required columns in Walmart file!");
            }

            Map<String, ReconciliationRecord> recordsToUpdate = new HashMap<>();

            // Dead Letter Queue
            List<ExceptionRecord> failedRows = new ArrayList<>();

            // --- MAIN PROCESSING LOOP ---
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                // Extract Values
                String transactionType = ExcelUtils.getStringValue(row.getCell(typeIndex)).trim().toUpperCase();
                String transactionDesc = ExcelUtils.getStringValue(row.getCell(descIndex)).trim().toUpperCase();
                String siteOrderId = ExcelUtils.getStringValue(row.getCell(poIndex)).trim().toUpperCase();
                double amount = ExcelUtils.getDoubleValue(row.getCell(amountIndex)) * -1;
                String amountType = ExcelUtils.getStringValue(row.getCell(amountTypeIndex)).trim().toUpperCase();
                String sku = ExcelUtils.getStringValue(row.getCell(skuIndex)).trim().toUpperCase();

                if (siteOrderId.isEmpty() || sku.isEmpty()) continue;
                String compositeId = siteOrderId + "-" + sku;

                // Fetch record from map or pull from DB if we haven't seen it yet
                ReconciliationRecord record = recordsToUpdate.get(compositeId);
                if (record == null) {
                    Optional<ReconciliationRecord> dbRecord = repository.findById(compositeId);
                    if (dbRecord.isPresent()){
                        record = dbRecord.get();
                        recordsToUpdate.put(compositeId, record);
                    } else {
                        // FAULT TOLERANCE: Record the missing order and skip to the next row
                        failedRows.add(new ExceptionRecord(siteOrderId, sku, amount, transactionDesc, "Missing from Rithum base data"));
                        continue;
                    }
                }

                // Route the amount to the correct field based on ledger hierarchy
                switch (transactionType) {
                    case ("SALE") -> {
                        switch (amountType) {
                            case ("PRODUCT PRICE") -> record.setSiteOrderAmount((record.getSiteOrderAmount() != null ? record.getSiteOrderAmount() : 0.0) + (amount * -1));
                            case ("COMMISSION ON PRODUCT") -> record.setSiteOrderFee((record.getSiteOrderFee() != null ? record.getSiteOrderFee() : 0.0) + amount);
                            case ("TOTAL WALMART FUNDED SAVINGS") -> record.setSiteOrderOtherFees1((record.getSiteOrderOtherFees1() != null ? record.getSiteOrderOtherFees1() : 0.0) + amount);
                            case ("PROMO CODE") -> record.setSiteOrderOtherFees2((record.getSiteOrderOtherFees2() != null ? record.getSiteOrderOtherFees2() : 0.0) + amount);
                            case ("OTHER TAX (FEES)") -> record.setSiteOrderOtherFees3((record.getSiteOrderOtherFees3() != null ? record.getSiteOrderOtherFees3() : 0.0) + amount);
                        }
                    }
                    case ("REFUND") -> {

                        if ("RETURN REFUND".equals(transactionDesc) || "SELLER INITIATED RETURNS".equals(transactionDesc)){
                            record.setReturnStatus("Yes");
                        }

                        switch (amountType) {
                            case ("PRODUCT PRICE") -> record.setAmountRefunded((record.getAmountRefunded() != null ? record.getAmountRefunded() : 0.0) + amount);
                            case ("COMMISSION ON PRODUCT") -> record.setReturnFee1((record.getReturnFee1() != null ? record.getReturnFee1() : 0.0) + amount);
                            case ("TOTAL WALMART FUNDED SAVINGS") -> record.setReturnFee2((record.getReturnFee2() != null ? record.getReturnFee2() : 0.0) + amount);
                            case ("EXCESSREFUNDADJUSTMENT") -> record.setReturnFee3((record.getReturnFee3() != null ? record.getReturnFee3() : 0.0) + amount);
                        }
                    }
                }

                if ("WALMART RETURN SHIPPING CHARGE".equals(transactionDesc)) {
                    record.setReturnShipping((record.getReturnShipping() != null ? record.getReturnShipping() : 0.0) + amount);
                }

                if ("FEE/REIMBURSEMENT".equals(amountType)) {
                    try {
                        addDynamicReturnFee(record, amount);
                    } catch (BucketOverflowException e) {
                        // FAULT TOLERANCE: Record the overflow and continue processing the rest of the file
                        failedRows.add(new ExceptionRecord(siteOrderId, sku, amount, transactionDesc, "Bucket overflow: Too many fee lines"));
                    }
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

            // Save all the updated records back to the database.
            repository.saveAll(recordsToUpdate.values());
            System.out.println("Successfully aggregated and updated " + recordsToUpdate.size() + " Walmart records.");
            if (!failedRows.isEmpty()) {
                System.out.println("Encountered " + failedRows.size() + " problematic rows to be exported.");
            }
            return failedRows;

        } catch (Exception e) {
            throw new RuntimeException("Failed to parse Walmart Excel file: " + e.getMessage());
        }
    }

    private void addDynamicReturnFee(ReconciliationRecord record, double amount) {

        if (record.getReturnFee2() == null) {
            record.setReturnFee2(amount);
        } else if (record.getReturnFee3() == null) {
            record.setReturnFee3(amount);
        } else {
            throw new BucketOverflowException("No Empty ReturnFee columns remaining");
        }
    }
}
