package com.apex.reconciliation_app.service;

import com.apex.reconciliation_app.dto.MarketplaceParseResult;
import com.apex.reconciliation_app.enums.AmazonColumn;
import com.apex.reconciliation_app.model.*;
import com.apex.reconciliation_app.repository.AmazonRawTransactionRepository;
import com.apex.reconciliation_app.repository.AmazonSuspenseRepository;
import com.apex.reconciliation_app.repository.ReconciliationRepository;
import com.apex.reconciliation_app.util.ExcelUtils;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.BeanUtils;
import org.springframework.stereotype.Service;

import java.io.InputStream;
import java.util.*;

@Service
@RequiredArgsConstructor
public class AmazonParserService {

    private final ReconciliationRepository repository;
    private final AmazonRawTransactionRepository auditRepository;
    private final AmazonSuspenseRepository suspenseRepository;

    public MarketplaceParseResult<AmazonSuspense, AmazonRawTransaction> parseAndUpdate(InputStream inputStream) {
        try (Workbook workbook = new XSSFWorkbook(inputStream)) {
            Sheet sheet = workbook.getSheetAt(0);
            Map<AmazonColumn, Integer> headerMap = ExcelUtils.buildHeaderMap(sheet.getRow(0), AmazonColumn.class);

            Map<String, ReconciliationRecord> recordsToUpdate = new HashMap<>();

            List<AmazonRawTransaction> auditTrail = new ArrayList<>();
            List<AmazonSuspense> actionableSuspense = new ArrayList<>();
            List<AmazonSuspense> errorSuspense = new ArrayList<>();
            Set<String> processedLineIds = new HashSet<>();

            // --- MAIN PROCESSING LOOP ---
                // ADD LOGIC
            for (int i = 1; i < sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                AmazonRawTransaction auditRow = buildAuditRow(row, headerMap);

                // Identifiers
                String orderId = auditRow.getOrderId() != null ? auditRow.getOrderId().trim() : "";
                String sku = auditRow.getSku() != null ? auditRow.getSku().trim() : "";
                if (orderId.isEmpty()) {
                    errorSuspense.add(buildSuspenseRow(auditRow, "Skipped: Missing Purchase Order(Non-order line item)"));
                    continue;
                };

                // Routing Variables
                String type = auditRow.getType() != null ? auditRow.getType().trim().toUpperCase() : "";
                String description = auditRow.getDescription() != null ? auditRow.getDescription().trim().toUpperCase() : "";

                double productSales = auditRow.getProductSales() != null ? auditRow.getProductSales() : 0.0;
                double shippingCredits = auditRow.getShippingCredits() != null ? auditRow.getShippingCredits() * -1 : 0.0;
                double giftWrapCredits = auditRow.getGiftWrapCredits() != null ? auditRow.getGiftWrapCredits() * -1 : 0.0;
                double regulatoryFee = auditRow.getRegulatoryFee() != null ? auditRow.getRegulatoryFee() * -1 : 0.0;
                double promotionalRebates = auditRow.getPromotionalRebates() != null ? auditRow.getPromotionalRebates() * -1 : 0.0;
                double sellingFees = auditRow.getSellingFees() != null ? auditRow.getSellingFees() * -1 : 0.0;
                double fbaFees = auditRow.getFbaFees() != null ? auditRow.getFbaFees() * -1 : 0.0;
                double otherTransactionFees = auditRow.getOtherTransactionFees() != null ? auditRow.getOtherTransactionFees() * -1 : 0.0;
                double other  = auditRow.getOther() != null ? auditRow.getOther() * -1 : 0.0;

                // Generate and set Composite Transaction ID
                String compositeTransactionId = String.format("%s-%s-%s-%s",
                        orderId, sku, type, description);
                auditRow.setCompositeTransactionId(compositeTransactionId);

                // Idempotency checks
                if (processedLineIds.contains(compositeTransactionId)) {
                    errorSuspense.add(buildSuspenseRow(auditRow, "Duplicate Record: Found multiple times in current upload"));
                    continue;
                }
                processedLineIds.add(compositeTransactionId);

                if (auditRepository.existsByCompositeTransactionId(compositeTransactionId)) {
                    errorSuspense.add(buildSuspenseRow(auditRow, "Duplicate Record: Already processed in a previous upload"));
                    continue;
                }

                // Fetch target records (Single for Items, Multi for Shipping)
                List<ReconciliationRecord> targetRecords = new ArrayList<>();

                if (!sku.isEmpty()) {
                    String compositeId = orderId + "-" + sku;
                    ReconciliationRecord record = recordsToUpdate.get(compositeId);
                    if (record == null) {
                        Optional<ReconciliationRecord> dbRecord = repository.findById(compositeId);
                        if (dbRecord.isPresent()) {
                            record = dbRecord.get();
                            recordsToUpdate.put(compositeId, record);
                            targetRecords.add(record);
                        }
                    } else {
                        targetRecords.add(record);
                    }
                } else {
                    // SKU is empty: fetch all records for this Order ID
                    List<ReconciliationRecord> dbRecords = repository.findBySiteOrderId(orderId);
                    for (ReconciliationRecord dbRec : dbRecords) {
                        ReconciliationRecord cached = recordsToUpdate.get(dbRec.getCompositeId());
                        if (cached != null) {
                            targetRecords.add(cached);
                        } else {
                            recordsToUpdate.put(dbRec.getCompositeId(), dbRec);
                            targetRecords.add(dbRec);
                        }
                    }
                }

                if (targetRecords.isEmpty()) {
                    actionableSuspense.add(buildSuspenseRow(auditRow, "Missing from Rithum base data"));
                }

                int recordCount = targetRecords.size();
                double splitOther  = other / recordCount;

                for (ReconciliationRecord record : targetRecords) {
                    // Route amount to correct field
                    switch (type) {
                        case "ORDER" ->  {
                            record.setSiteOrderAmount((record.getSiteOrderAmount() != null ? record.getSiteOrderAmount() : 0.0) + productSales);
                            record.setSiteOrderFee((record.getSiteOrderFee() != null ? record.getSiteOrderFee() : 0.0) + sellingFees);

                            if (shippingCredits != 0) record.addDynamicRegularFee(shippingCredits, "SHIPPING CREDITS");
                            if (giftWrapCredits != 0) record.addDynamicRegularFee(giftWrapCredits, "GIFT WRAP CREDITS");
                            if (regulatoryFee != 0) record.addDynamicRegularFee(regulatoryFee, "REGULATORY FEE");
                            if (promotionalRebates != 0) record.addDynamicRegularFee(promotionalRebates, "PROMOTIONAL REBATES");
                            if (fbaFees != 0) record.addDynamicRegularFee(fbaFees, "FBA FEE");
                            if (otherTransactionFees != 0) record.addDynamicRegularFee(otherTransactionFees, "OTHER TRANSACTION FEE");
                            if (other != 0) record.addDynamicRegularFee(other, "OTHER");
                        }
                        case "REFUND" -> {
                            record.setReturnStatus("Yes");
                            record.setAmountRefunded((record.getAmountRefunded() != null ? record.getAmountRefunded() : 0.0) + productSales);
                            record.setCommissionRefund((record.getCommissionRefund() != null ? record.getCommissionRefund() : 0.0) + sellingFees);

                            if (shippingCredits != 0) record.addDynamicReturnFee(shippingCredits, "SHIPPING CREDITS REFUND");
                            if (giftWrapCredits != 0) record.addDynamicReturnFee(giftWrapCredits, "GIFT WRAP CREDITS REFUND");
                            if (regulatoryFee != 0) record.addDynamicReturnFee(regulatoryFee, "REGULATORY FEE REFUND");
                            if (promotionalRebates != 0) record.addDynamicReturnFee(promotionalRebates, "PROMO REBATE REFUND");
                            if (fbaFees != 0) record.addDynamicReturnFee(fbaFees, "FBA FEE REFUND");
                            if (otherTransactionFees != 0) record.addDynamicReturnFee(otherTransactionFees, "OTHER FEE REFUND");
                            if (other != 0) record.addDynamicReturnFee(other, "OTHER REFUND");
                        }
                        case "SHIPPING SERVICES" -> {
                            switch (description) {
                                case "SHIPPING LABEL PURCHASED THROUGH AMAZON" ->
                                        record.setActualShippingCosts((record.getActualShippingCosts() != null ? record.getActualShippingCosts() : 0.0) + splitOther);

                                case "ADJUSTMENT" -> {
                                    // If the bucket already has money in it, this is a multiple!
                                    if (record.getShippingAdjustments() != null && record.getShippingAdjustments() != 0.0) {
                                        String note = "Multiple Shipping Adjs Combined, ";
                                        // Only add the note if we haven't already added it
                                        if (record.getNotes() == null || !record.getNotes().contains(note)) {
                                            record.setNotes(record.getNotes() == null ? note : record.getNotes().concat(note));
                                        }
                                    }
                                    record.setShippingAdjustments((record.getShippingAdjustments() != null ? record.getShippingAdjustments() : 0.0) + splitOther);
                                }

                                case "RETURNPOSTAGEBILLING" ->
                                        record.setReturnShipping((record.getReturnShipping() != null ? record.getReturnShipping() : 0.0) + splitOther);

                                case "SHIPPING LABEL REFUNDED THROUGH AMAZON" ->
                                        record.addDynamicReturnFee(splitOther, "SHIPPING LABEL REFUND");

                                default -> record.addDynamicRegularFee(splitOther, "UNMAPPED SHIPPING SERVICE: " + description);
                            }
                        }
                        default -> {
                            // Safety net: Catches any NEW order-level transaction types Amazon might invent in the future
                            actionableSuspense.add(buildSuspenseRow(auditRow, "Action Required: Unmapped Transaction Type (" + type + ") found for Order."));
                            continue;
                        }
                    }
                }
                // Add survivors to audit trail
                auditTrail.add(auditRow);
            }

            // POST PROCESSING
            for (ReconciliationRecord record : recordsToUpdate.values()) {
                record.calculateCommissionRefundDelta();
            }

            // SAVE UPDATED DATA
            repository.saveAll(recordsToUpdate.values());
            auditRepository.saveAll(auditTrail);
            suspenseRepository.saveAll(actionableSuspense);

            List<AmazonSuspense> allReceiptErrors = new ArrayList<>(actionableSuspense);
            allReceiptErrors.addAll(errorSuspense);

            System.out.println("Updated " + recordsToUpdate.size() + " Rithum Master Amazon records.");
            System.out.println("Processed " + (auditTrail.size() + allReceiptErrors.size()) + " Amazon Marketplace rows");
            System.out.println("Saved " + auditTrail.size() + " Audit rows.");
            System.out.println("Saved " + actionableSuspense.size() + " Actionable Suspense rows.");
            System.out.println("Skipped " + errorSuspense.size() + " Duplicate rows (Added to receipt only)");

            return new MarketplaceParseResult<>(allReceiptErrors, auditTrail);

        } catch (Exception e) {
            throw new RuntimeException("Failed to parse Walmart Excel file: " + e.getMessage());
        }
    }

    private AmazonRawTransaction buildAuditRow(Row row, Map<AmazonColumn, Integer> headerMap) {
        return AmazonRawTransaction.builder()
                .dateTime(ExcelUtils.getDateSafe(row, headerMap, AmazonColumn.DATE_TIME))
                .settlementId(ExcelUtils.getStringSafe(row, headerMap, AmazonColumn.SETTLEMENT_ID))
                .type(ExcelUtils.getStringSafe(row, headerMap, AmazonColumn.TYPE))
                .orderId(ExcelUtils.getStringSafe(row, headerMap, AmazonColumn.ORDER_ID))
                .sku(ExcelUtils.getStringSafe(row, headerMap, AmazonColumn.SKU))
                .description(ExcelUtils.getStringSafe(row, headerMap, AmazonColumn.DESCRIPTION))
                .quantity(ExcelUtils.getDoubleSafe(row, headerMap, AmazonColumn.QUANTITY))
                .marketplace(ExcelUtils.getStringSafe(row, headerMap, AmazonColumn.MARKETPLACE))
                .accountType(ExcelUtils.getStringSafe(row, headerMap, AmazonColumn.ACCOUNT_TYPE))
                .fulfillment(ExcelUtils.getStringSafe(row, headerMap, AmazonColumn.FULFILLMENT))
                .orderCity(ExcelUtils.getStringSafe(row, headerMap, AmazonColumn.ORDER_CITY))
                .orderState(ExcelUtils.getStringSafe(row, headerMap, AmazonColumn.ORDER_STATE))
                .orderPostal(ExcelUtils.getStringSafe(row, headerMap, AmazonColumn.ORDER_POSTAL))
                .taxCollectionModel(ExcelUtils.getStringSafe(row, headerMap, AmazonColumn.TAX_COLLECTION_MODEL))
                .productSales(ExcelUtils.getDoubleSafe(row, headerMap, AmazonColumn.PRODUCT_SALES))
                .productSalesTax(ExcelUtils.getDoubleSafe(row, headerMap, AmazonColumn.PRODUCT_SALES_TAX))
                .shippingCredits(ExcelUtils.getDoubleSafe(row, headerMap, AmazonColumn.SHIPPING_CREDITS))
                .shippingCreditsTax(ExcelUtils.getDoubleSafe(row, headerMap, AmazonColumn.SHIPPING_CREDITS_TAX))
                .giftWrapCredits(ExcelUtils.getDoubleSafe(row, headerMap, AmazonColumn.GIFT_WRAP_CREDITS))
                .giftWrapCreditTax(ExcelUtils.getDoubleSafe(row, headerMap, AmazonColumn.GIFTWRAP_CREDITS_TAX))
                .RegulatoryFee(ExcelUtils.getDoubleSafe(row, headerMap, AmazonColumn.REGULATORY_FEE))
                .taxOnRegulatoryFee(ExcelUtils.getDoubleSafe(row, headerMap, AmazonColumn.TAX_ON_REGULATORY_FEE))
                .promotionalRebates(ExcelUtils.getDoubleSafe(row, headerMap, AmazonColumn.PROMOTIONAL_REBATES))
                .promotionalRebatestax(ExcelUtils.getDoubleSafe(row, headerMap, AmazonColumn.PROMOTIONAL_REBATES_TAX))
                .marketplaceWithheldTax(ExcelUtils.getDoubleSafe(row, headerMap, AmazonColumn.MARKETPLACE_WITHHELD_TAX))
                .sellingFees(ExcelUtils.getDoubleSafe(row, headerMap, AmazonColumn.SELLING_FEES))
                .fbaFees(ExcelUtils.getDoubleSafe(row, headerMap, AmazonColumn.FBA_FEES))
                .otherTransactionFees(ExcelUtils.getDoubleSafe(row, headerMap, AmazonColumn.OTHER_TRANSACTION_FEES))
                .other(ExcelUtils.getDoubleSafe(row, headerMap, AmazonColumn.OTHER))
                .total(ExcelUtils.getDoubleSafe(row, headerMap, AmazonColumn.TOTAL))
                .transactionStatus(ExcelUtils.getStringSafe(row, headerMap, AmazonColumn.TRANSACTION_STATUS))
                .transactionReleaseDate(ExcelUtils.getDateSafe(row, headerMap, AmazonColumn.TRANSACTION_RELEASE_DATE))
                .build();
    }

    private AmazonSuspense buildSuspenseRow(AmazonRawTransaction auditRow, String reason) {
        AmazonSuspense suspenseRow = new AmazonSuspense();
        BeanUtils.copyProperties(auditRow, suspenseRow);
        suspenseRow.setErrorReason(reason);
        return suspenseRow;
    }

}
