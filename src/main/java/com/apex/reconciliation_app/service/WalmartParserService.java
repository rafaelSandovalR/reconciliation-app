package com.apex.reconciliation_app.service;

import com.apex.reconciliation_app.enums.WalmartColumn;
import com.apex.reconciliation_app.exception.BucketOverflowException;
import com.apex.reconciliation_app.model.ReconciliationRecord;
import com.apex.reconciliation_app.model.WalmartRawTransaction;
import com.apex.reconciliation_app.model.WalmartSuspense;
import com.apex.reconciliation_app.repository.ReconciliationRepository;
import com.apex.reconciliation_app.repository.WalmartRawTransactionRepository;
import com.apex.reconciliation_app.repository.WalmartSuspenseRepository;
import com.apex.reconciliation_app.util.ExcelUtils;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.InputStream;
import java.time.LocalDateTime;
import java.util.*;

@Service
@RequiredArgsConstructor
public class WalmartParserService {

    private final ReconciliationRepository repository;
    private final WalmartRawTransactionRepository auditRepository;
    private final WalmartSuspenseRepository suspenseRepository;

    public List<WalmartSuspense> parseAndUpdate(InputStream inputStream) {
        try (Workbook workbook = new XSSFWorkbook(inputStream)) {
            Sheet sheet = workbook.getSheetAt(0);

            // --- DYNAMIC HEADER MAPPING ---
            Map<String, Integer> stringHeaderMap = ExcelUtils.getHeaderMap(sheet.getRow(0));
            Map<WalmartColumn, Integer> headerMap = new EnumMap<>(WalmartColumn.class);

            for (WalmartColumn col : WalmartColumn.values()) {
                Integer index = stringHeaderMap.get(col.getHeaderName().toUpperCase());
                if (index != null) {
                    headerMap.put(col,index);
                }
            }

            // Verify core routing columns exist
            if (!headerMap.containsKey(WalmartColumn.TRANSACTION_TYPE) || !headerMap.containsKey(WalmartColumn.PURCHASE_ORDER) || !headerMap.containsKey(WalmartColumn.AMOUNT)) {
                throw new RuntimeException("Missing required columns in Walmart file!");
            }

            Map<String, ReconciliationRecord> recordsToUpdate = new HashMap<>();

            // Audit Trail and Suspense Queue
            List<WalmartRawTransaction> auditTrail = new ArrayList<>();
            List<WalmartSuspense> failedRows = new ArrayList<>();

            // --- MAIN PROCESSING LOOP ---
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                // Extract core routing variables
                String transactionType = ExcelUtils.getStringValue(row.getCell(headerMap.get(WalmartColumn.TRANSACTION_TYPE))).trim().toUpperCase();
                String transactionDesc = ExcelUtils.getStringValue(row.getCell(headerMap.get(WalmartColumn.TRANSACTION_DESC))).trim().toUpperCase();
                String purchaseOrder = ExcelUtils.getStringValue(row.getCell(headerMap.get(WalmartColumn.PURCHASE_ORDER))).trim().toUpperCase();
                double rawAmount = ExcelUtils.getDoubleValue(row.getCell(headerMap.get(WalmartColumn.AMOUNT)));
                double invertedAmount = rawAmount * -1; //
                String amountType = ExcelUtils.getStringValue(row.getCell(headerMap.get(WalmartColumn.AMOUNT_TYPE))).trim().toUpperCase();
                String sku = ExcelUtils.getStringValue(row.getCell(headerMap.get(WalmartColumn.SKU))).trim().toUpperCase();

                if (purchaseOrder.isEmpty() || sku.isEmpty()) continue;
                String compositeId = purchaseOrder + "-" + sku;

                // Fetch record from map or DB
                ReconciliationRecord record = recordsToUpdate.get(compositeId);
                if (record == null) {
                    Optional<ReconciliationRecord> dbRecord = repository.findById(compositeId);
                    if (dbRecord.isPresent()){
                        record = dbRecord.get();
                        recordsToUpdate.put(compositeId, record);
                    } else {
                        // FAULT TOLERANCE: Record the missing order to the Suspense Queue
                        failedRows.add(buildSuspenseRow(row, headerMap, "Missing from Rithum base data"));
                        continue;
                    }
                }

                // Route the amount to the correct field based on ledger hierarchy
                switch (transactionType) {
                    case ("SALE") -> {
                        switch (amountType) {
                            case ("PRODUCT PRICE") -> record.setSiteOrderAmount((record.getSiteOrderAmount() != null ? record.getSiteOrderAmount() : 0.0) + rawAmount);
                            case ("COMMISSION ON PRODUCT") -> record.setSiteOrderFee((record.getSiteOrderFee() != null ? record.getSiteOrderFee() : 0.0) + invertedAmount);
                            case ("TOTAL WALMART FUNDED SAVINGS") -> record.setSiteOrderOtherFees1((record.getSiteOrderOtherFees1() != null ? record.getSiteOrderOtherFees1() : 0.0) + invertedAmount);
                            case ("PROMO CODE") -> record.setSiteOrderOtherFees2((record.getSiteOrderOtherFees2() != null ? record.getSiteOrderOtherFees2() : 0.0) + invertedAmount);
                            case ("OTHER TAX (FEES)") -> record.setSiteOrderOtherFees3((record.getSiteOrderOtherFees3() != null ? record.getSiteOrderOtherFees3() : 0.0) + invertedAmount);
                        }
                    }
                    case ("REFUND") -> {

                        if ("RETURN REFUND".equals(transactionDesc) || "SELLER INITIATED RETURNS".equals(transactionDesc)){
                            record.setReturnStatus("Yes");
                        }

                        switch (amountType) {
                            case ("PRODUCT PRICE") -> record.setAmountRefunded((record.getAmountRefunded() != null ? record.getAmountRefunded() : 0.0) + invertedAmount);
                            case ("COMMISSION ON PRODUCT") -> record.setReturnFee1((record.getReturnFee1() != null ? record.getReturnFee1() : 0.0) + invertedAmount);
                            case ("TOTAL WALMART FUNDED SAVINGS") -> record.setReturnFee2((record.getReturnFee2() != null ? record.getReturnFee2() : 0.0) + invertedAmount);
                            case ("EXCESSREFUNDADJUSTMENT") -> record.setReturnFee3((record.getReturnFee3() != null ? record.getReturnFee3() : 0.0) + invertedAmount);
                        }
                    }
                }

                if ("WALMART RETURN SHIPPING CHARGE".equals(transactionDesc)) {
                    record.setReturnShipping((record.getReturnShipping() != null ? record.getReturnShipping() : 0.0) + invertedAmount);
                }

                if ("FEE/REIMBURSEMENT".equals(amountType)) {
                    try {
                        addDynamicReturnFee(record, invertedAmount);
                    } catch (BucketOverflowException e) {
                        // FAULT TOLERANCE: Record overflow to Suspense Queue
                        failedRows.add(buildSuspenseRow(row, headerMap, "Bucket overflow: Too many fee lines. Need to consolidate fees or create extra column."));
                        continue;
                    }
                }

                // If the row survived all the checks, it's successful and will be added to the audit trail
                auditTrail.add(buildAuditRow(row,headerMap));
            }

            // POST PROCESSING PASS
            for (ReconciliationRecord record : recordsToUpdate.values()) {
                if (record.getReturnFee1() != null && record.getReturnFee1() < 0) {
                    double originalCommission = record.getSiteOrderFee() != null ? record.getSiteOrderFee() : 0.0;
                    double commissionRefundDifference = originalCommission + record.getReturnFee1();
                    record.setReturnFee1(commissionRefundDifference);
                }
            }

            // Save all the updated records back to the database.
            repository.saveAll(recordsToUpdate.values());
            auditRepository.saveAll(auditTrail);
            suspenseRepository.saveAll(failedRows);

            System.out.println("Updated " + recordsToUpdate.size() + " Master Walmart records.");
            System.out.println("Saved " + auditTrail.size() + " Audit rows.");
            System.out.println("Saved " + failedRows.size() + " Suspense rows.");
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

    private WalmartRawTransaction buildAuditRow(Row row, Map<WalmartColumn, Integer> headerMap) {
        return WalmartRawTransaction.builder()
                .transactionKey(getStringSafe(row, headerMap, WalmartColumn.TRANSACTION_KEY))
                .transactionPostedTimestamp(getDateSafe(row, headerMap, WalmartColumn.TRANSACTION_POSTED_TIMESTAMP))
                .transactionType(getStringSafe(row, headerMap, WalmartColumn.TRANSACTION_TYPE))
                .transactionDesc(getStringSafe(row, headerMap, WalmartColumn.TRANSACTION_DESC))
                .customerOrder(getStringSafe(row, headerMap, WalmartColumn.CUSTOMER_ORDER))
                .customerOrderLine(getStringSafe(row, headerMap, WalmartColumn.CUSTOMER_ORDER_LINE))
                .purchaseOrder(getStringSafe(row, headerMap, WalmartColumn.PURCHASE_ORDER))
                .purchaseOrderLine(getStringSafe(row, headerMap, WalmartColumn.PURCHASE_ORDER_LINE))
                .amount(getDoubleSafe(row, headerMap, WalmartColumn.AMOUNT))
                .amountType(getStringSafe(row, headerMap, WalmartColumn.AMOUNT_TYPE))
                .shipQty(getDoubleSafe(row, headerMap, WalmartColumn.SHIP_QTY))
                .commissionRate(getDoubleSafe(row, headerMap, WalmartColumn.COMMISSION_RATE))
                .baseCommissionRate(getDoubleSafe(row, headerMap, WalmartColumn.BASE_COMMISSION_RATE))
                .transactionReasonDesc(getStringSafe(row, headerMap, WalmartColumn.TRANSACTION_REASON_DESC))
                .partnerItemId(getStringSafe(row, headerMap, WalmartColumn.SKU))
                .partnerGtIn(getStringSafe(row, headerMap, WalmartColumn.PARTNER_GTIN))
                .partnerItemName(getStringSafe(row, headerMap, WalmartColumn.PARTNER_ITEM_NAME))
                .productTaxCode(getStringSafe(row, headerMap, WalmartColumn.PRODUCT_TAX_CODE))
                .shipToState(getStringSafe(row, headerMap, WalmartColumn.SHIP_TO_STATE))
                .shipToCity(getStringSafe(row, headerMap, WalmartColumn.SHIP_TO_CITY))
                .shipToZipcode(getStringSafe(row, headerMap, WalmartColumn.SHIP_TO_ZIPCODE))
                .contractCategory(getStringSafe(row, headerMap, WalmartColumn.CONTRACT_CATEGORY))
                .productType(getStringSafe(row, headerMap, WalmartColumn.PRODUCT_TYPE))
                .commissionRule(getStringSafe(row, headerMap, WalmartColumn.COMMISSION_RULE))
                .shippingMethod(getStringSafe(row, headerMap, WalmartColumn.SHIPPING_METHOD))
                .fulfillmentType(getStringSafe(row, headerMap, WalmartColumn.FULFILLMENT_TYPE))
                .fulfillmentDetails(getStringSafe(row, headerMap, WalmartColumn.FULFILLMENT_DETAILS))
                .originalCommission(getDoubleSafe(row, headerMap, WalmartColumn.ORIGINAL_COMMISSION))
                .commissionIncentiveProgram(getDoubleSafe(row, headerMap, WalmartColumn.COMMISSION_INCENTIVE_PROGRAM))
                .commissionSaving(getDoubleSafe(row, headerMap, WalmartColumn.COMMISSION_SAVING))
                .customerPromoType(getStringSafe(row, headerMap, WalmartColumn.CUSTOMER_PROMO_TYPE))
                .totalWalmartFundedSavings(getDoubleSafe(row, headerMap, WalmartColumn.TOTAL_WALMART_FUNDED_SAVINGS))
                .campaignId(getStringSafe(row, headerMap, WalmartColumn.CAMPAIGN_ID))
                .itemCondition(getStringSafe(row, headerMap, WalmartColumn.ITEM_CONDITION))
                .originalCharge(getDoubleSafe(row, headerMap, WalmartColumn.ORIGINAL_CHARGE))
                .chargeSavings(getDoubleSafe(row, headerMap, WalmartColumn.CHARGE_SAVINGS))
                .incentiveProgramName(getStringSafe(row, headerMap, WalmartColumn.INCENTIVE_PROGRAM_NAME))
                .shipToCountry(getStringSafe(row, headerMap, WalmartColumn.SHIP_TO_COUNTRY))
                .build();
    }

    private WalmartSuspense buildSuspenseRow(Row row, Map<WalmartColumn, Integer> headerMap, String reason) {
        return WalmartSuspense.builder()
                .errorReason(reason)
                .transactionKey(getStringSafe(row, headerMap, WalmartColumn.TRANSACTION_KEY))
                .transactionPostedTimestamp(getDateSafe(row, headerMap, WalmartColumn.TRANSACTION_POSTED_TIMESTAMP))
                .transactionType(getStringSafe(row, headerMap, WalmartColumn.TRANSACTION_TYPE))
                .transactionDesc(getStringSafe(row, headerMap, WalmartColumn.TRANSACTION_DESC))
                .customerOrder(getStringSafe(row, headerMap, WalmartColumn.CUSTOMER_ORDER))
                .customerOrderLine(getStringSafe(row, headerMap, WalmartColumn.CUSTOMER_ORDER_LINE))
                .purchaseOrder(getStringSafe(row, headerMap, WalmartColumn.PURCHASE_ORDER))
                .purchaseOrderLine(getStringSafe(row, headerMap, WalmartColumn.PURCHASE_ORDER_LINE))
                .amount(getDoubleSafe(row, headerMap, WalmartColumn.AMOUNT))
                .amountType(getStringSafe(row, headerMap, WalmartColumn.AMOUNT_TYPE))
                .shipQty(getDoubleSafe(row, headerMap, WalmartColumn.SHIP_QTY))
                .commissionRate(getDoubleSafe(row, headerMap, WalmartColumn.COMMISSION_RATE))
                .baseCommissionRate(getDoubleSafe(row, headerMap, WalmartColumn.BASE_COMMISSION_RATE))
                .transactionReasonDesc(getStringSafe(row, headerMap, WalmartColumn.TRANSACTION_REASON_DESC))
                .partnerItemId(getStringSafe(row, headerMap, WalmartColumn.SKU))
                .partnerGtIn(getStringSafe(row, headerMap, WalmartColumn.PARTNER_GTIN))
                .partnerItemName(getStringSafe(row, headerMap, WalmartColumn.PARTNER_ITEM_NAME))
                .productTaxCode(getStringSafe(row, headerMap, WalmartColumn.PRODUCT_TAX_CODE))
                .shipToState(getStringSafe(row, headerMap, WalmartColumn.SHIP_TO_STATE))
                .shipToCity(getStringSafe(row, headerMap, WalmartColumn.SHIP_TO_CITY))
                .shipToZipcode(getStringSafe(row, headerMap, WalmartColumn.SHIP_TO_ZIPCODE))
                .contractCategory(getStringSafe(row, headerMap, WalmartColumn.CONTRACT_CATEGORY))
                .productType(getStringSafe(row, headerMap, WalmartColumn.PRODUCT_TYPE))
                .commissionRule(getStringSafe(row, headerMap, WalmartColumn.COMMISSION_RULE))
                .shippingMethod(getStringSafe(row, headerMap, WalmartColumn.SHIPPING_METHOD))
                .fulfillmentType(getStringSafe(row, headerMap, WalmartColumn.FULFILLMENT_TYPE))
                .fulfillmentDetails(getStringSafe(row, headerMap, WalmartColumn.FULFILLMENT_DETAILS))
                .originalCommission(getDoubleSafe(row, headerMap, WalmartColumn.ORIGINAL_COMMISSION))
                .commissionIncentiveProgram(getDoubleSafe(row, headerMap, WalmartColumn.COMMISSION_INCENTIVE_PROGRAM))
                .commissionSaving(getDoubleSafe(row, headerMap, WalmartColumn.COMMISSION_SAVING))
                .customerPromoType(getStringSafe(row, headerMap, WalmartColumn.CUSTOMER_PROMO_TYPE))
                .totalWalmartFundedSavings(getDoubleSafe(row, headerMap, WalmartColumn.TOTAL_WALMART_FUNDED_SAVINGS))
                .campaignId(getStringSafe(row, headerMap, WalmartColumn.CAMPAIGN_ID))
                .itemCondition(getStringSafe(row, headerMap, WalmartColumn.ITEM_CONDITION))
                .originalCharge(getDoubleSafe(row, headerMap, WalmartColumn.ORIGINAL_CHARGE))
                .chargeSavings(getDoubleSafe(row, headerMap, WalmartColumn.CHARGE_SAVINGS))
                .incentiveProgramName(getStringSafe(row, headerMap, WalmartColumn.INCENTIVE_PROGRAM_NAME))
                .shipToCountry(getStringSafe(row, headerMap, WalmartColumn.SHIP_TO_COUNTRY))
                .build();
    }

    // SAFE EXTRACTION UTILS
    private String getStringSafe(Row row, Map<WalmartColumn, Integer> headerMap, WalmartColumn col) {
        if (!headerMap.containsKey(col)) return null;
        return ExcelUtils.getStringValue(row.getCell(headerMap.get(col)));
    }

    private Double getDoubleSafe(Row row, Map<WalmartColumn, Integer> headerMap, WalmartColumn col) {
        if (!headerMap.containsKey(col)) return null;
        Cell cell = row.getCell(headerMap.get(col));
        if (cell == null || cell.getCellType() == CellType.BLANK) return null;
        return ExcelUtils.getDoubleValue(cell);
    }

    private LocalDateTime getDateSafe(Row row, Map<WalmartColumn, Integer> headerMap, WalmartColumn col) {
        if (!headerMap.containsKey(col)) return null;
        Cell cell = row.getCell(headerMap.get(col));
        if (cell != null && cell.getCellType() != CellType.BLANK && DateUtil.isCellDateFormatted(cell)) {
            return cell.getLocalDateTimeCellValue();
        }
        return null;
    }
}
