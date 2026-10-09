package com.apex.reconciliation_app.service;

import com.apex.reconciliation_app.dto.MarketplaceParseResult;
import com.apex.reconciliation_app.enums.EbayColumn;
import com.apex.reconciliation_app.model.ebay.EbayRawTransaction;
import com.apex.reconciliation_app.model.ebay.EbaySuspense;
import com.apex.reconciliation_app.model.ReconciliationRecord;
import com.apex.reconciliation_app.repository.EbayRawTransactionRepository;
import com.apex.reconciliation_app.repository.EbaySuspenseRepository;
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
public class EbayParserService {

    private final ReconciliationRepository repository;
    private final EbayRawTransactionRepository auditRepository;
    private final EbaySuspenseRepository suspenseRepository;

    public MarketplaceParseResult<EbaySuspense, EbayRawTransaction> parseAndUpdate(InputStream inputStream) {
        try (Workbook workbook = new XSSFWorkbook(inputStream)) {
            Sheet sheet = workbook.getSheetAt(0);
            Map<EbayColumn, Integer> headerMap = ExcelUtils.buildHeaderMap(sheet.getRow(0), EbayColumn.class);
            Map<String, ReconciliationRecord> recordsToUpdate = new HashMap<>();
            List<ReconciliationRecord> recordsToDelete = new ArrayList<>();
            List<EbayRawTransaction> auditTrail = new ArrayList<>();
            List<EbaySuspense> actionableSuspense = new ArrayList<>();
            List<EbaySuspense> errorSuspense = new ArrayList<>();
            Set<String> processedLineIds = new HashSet<>();

            // PASS 1: Ingestion, Map Building & Summary Filtering
            List<EbayRawTransaction> itemLevelRows = new ArrayList<>();
            Map<String, String> itemToSkuMap = new HashMap<>();     // Key: OrderNumber-ItemID -> Anchored SKU
            Map<String, String> returnToSkuMap = new HashMap<>();   // Key: ReferenceID -> SKU

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                EbayRawTransaction auditRow = buildAuditRow(row, headerMap);
                String orderNumber = auditRow.getOrderNumber() != null ? auditRow.getOrderNumber().trim() : "";

                if (orderNumber.isEmpty()) {
                    errorSuspense.add(buildSuspenseRow(auditRow, "Missing Order Number(Non-order line item)"));
                    continue;
                }

                String type = cleanEbayString(auditRow.getType()).toUpperCase();
                String sku = cleanEbayString(auditRow.getCustomLabel());
                String itemId = cleanEbayString(auditRow.getItemId());
                String referenceId = cleanEbayString(auditRow.getReferenceId());

                // Filter Master Summary Rows
                if ("ORDER".equals(type) && itemId.isEmpty()) {
                    auditTrail.add(auditRow);
                    continue;
                }

                // Build anchor maps
                if ("ORDER".equals(type) && !sku.isEmpty()) {
                    String itemKey = orderNumber + "-" + itemId;
                    itemToSkuMap.putIfAbsent(itemKey, sku);
                }
                if ("REFUND".equals(type) && !sku.isEmpty() & !referenceId.isEmpty()) {
                    returnToSkuMap.put(referenceId, sku);
                }

                itemLevelRows.add(auditRow);
            }

            // PASS 2: Routing & Math
            for (EbayRawTransaction auditRow : itemLevelRows) {
                String orderNumber = cleanEbayString(auditRow.getOrderNumber());
                String sku = cleanEbayString(auditRow.getCustomLabel());
                String transactionId = cleanEbayString(auditRow.getTransactionId());
                String itemId = cleanEbayString(auditRow.getItemId());
                String referenceId = cleanEbayString(auditRow.getReferenceId());
                String type = cleanEbayString(auditRow.getType()).toUpperCase();
                String description = cleanEbayString(auditRow.getDescription()).toUpperCase();

                // Idempotency Check
                String compositeTransactionId = String.format("%s-%s-%s-%s-%s-%s-%s",
                        type, orderNumber, transactionId, itemId, sku, referenceId, description);
                auditRow.setCompositeTransactionId(compositeTransactionId);

                if (processedLineIds.contains(compositeTransactionId)) {
                    errorSuspense.add(buildSuspenseRow(auditRow, "Duplicate record in upload file, already processed"));
                    continue;
                }
                processedLineIds.add(compositeTransactionId);

                if (auditRepository.existsByCompositeTransactionId(compositeTransactionId)) {
                    errorSuspense.add(buildSuspenseRow(auditRow, "Already processed in a previous upload"));
                    continue;
                }

                if ("HOLD".equals(type) || "PAYMENT DISPUTE".equals(type) || "CLAIM".equals(type)) {
                    auditTrail.add(auditRow);
                    continue;
                }

                // Determine Target SKU using Maps
                String targetSku = sku;;
                if (targetSku.isEmpty()) {
                    if ("OTHER FEE".equals(type) && !itemId.isEmpty()) {
                        targetSku = itemToSkuMap.getOrDefault(orderNumber + "-" + itemId, "");
                    } else if ("SHIPPING LABEL".equals(type) && !referenceId.isEmpty()) {
                        targetSku = returnToSkuMap.getOrDefault(referenceId, "");
                    }
                }

                String compositeId = "";
                ReconciliationRecord record = null;

                // Try Primary Key (Cache -> DB) ONLY if we have a SKU
                if (!targetSku.isEmpty()) {
                    compositeId = orderNumber + "-" + targetSku;
                    auditRow.setCompositeId(compositeId);

                    record = recordsToUpdate.get(compositeId);
                    if (record == null) {
                        Optional<ReconciliationRecord> dbRecord = repository.findById(compositeId);
                        if (dbRecord.isPresent()) record = dbRecord.get();
                    }
                }

                // Cascading DB fallback (If pk failed or targetSku is STILL empty becuase it wasn't in the pass 1 map
                if (record == null) {
                    List<ReconciliationRecord> baseRecords = repository.findBySiteOrderId(orderNumber);

                    if (baseRecords.isEmpty()) {
                        actionableSuspense.add(buildSuspenseRow(auditRow, "Missing from Rithum base data."));
                        continue;
                    } else if (baseRecords.size() == 1) {
                        // Fallback 1: Single-item order -> Auto-heal
                        record = baseRecords.get(0);
                    } else {
                        // Fallback 2: Precise Match via eBay Site Listing ID
                        ReconciliationRecord listingIdMatch = baseRecords.stream()
                                .filter(r -> itemId.equals(r.getSiteListingId()))
                                .findFirst()
                                .orElse(null);

                        if (listingIdMatch != null) {
                            record = listingIdMatch;
                        } else {
                            // Could add third fallback: fuzzy title match
                            actionableSuspense.add(buildSuspenseRow(auditRow, "Ambigious match: Rithum data missing SKU, and all fallback matches failed."));
                            continue;
                        }
                    }

                    // Auto-heal Rithum Record (Only if we have a targetSku to heal it with
                    if (!targetSku.isEmpty() && !targetSku.equals(record.getSiteOrderItemId())) {
                        String oldCompositeId = record.getCompositeId();
                        ReconciliationRecord healedRecord = new ReconciliationRecord();
                        BeanUtils.copyProperties(record, healedRecord);

                        healedRecord.setSku(targetSku);
                        healedRecord.setCompositeId(compositeId);

                        recordsToDelete.add(record);
                        record = healedRecord;
                    }

                    auditRow.setCompositeId(record.getCompositeId());
                }

                recordsToUpdate.put(record.getCompositeId(), record);

                // Invert all of these
                double grossTransactionAmount = invert(auditRow.getGrossTransactionAmount());
                double itemSubtotal = zeroIfNull(auditRow.getItemSubtotal());
                double shippingAndHandling = zeroIfNull(auditRow.getShippingAndHandling());

                // Fallback for multi-item lines where GTA is "--" (0.0)
                if (grossTransactionAmount == 0 && itemSubtotal != 0) {
                    grossTransactionAmount = (itemSubtotal + shippingAndHandling) * -1;
                }

                double finalValueFeeVariable = invert(auditRow.getFinalValueFeeVariable());
                double finalValueFeeFixed = invert(auditRow.getFinalValueFeeFixed());
                double commission = finalValueFeeFixed + finalValueFeeVariable;
                double regulatoryOperatingFee = invert(auditRow.getRegulatoryOperatingFee());
                double veryHighItemNotAsDescribedFee = invert(auditRow.getVeryHighItemNotAsDescribedFee());
                double belowStandardPerformancefee = invert(auditRow.getBelowStandardPerformanceFee());
                double internationalFee = invert(auditRow.getInternationalFee());
                double charityDonation = invert(auditRow.getCharityDonation());
                double depositProcessingFee = invert(auditRow.getDepositProcessingFee());



                switch (type) {
                    case "ORDER" -> {
                        // Undo the invert ONLY for siteOrderAmount
                        record.setSiteOrderAmount(zeroIfNull(record.getSiteOrderAmount()) + grossTransactionAmount * -1);
                        record.setSiteOrderFee(zeroIfNull(record.getSiteOrderFee()) + commission);

                        if (regulatoryOperatingFee != 0) record.addDynamicRegularFee(regulatoryOperatingFee, "REGULATORY OPERATING FEE");
                        if (veryHighItemNotAsDescribedFee != 0) record.addDynamicRegularFee(veryHighItemNotAsDescribedFee, "VERY HIGH \"ITEM NOT AS DESCRIBED\" FEE");
                        if (belowStandardPerformancefee != 0) record.addDynamicRegularFee(belowStandardPerformancefee, "BELOW STANDARD PERFORMANCE FEE");
                        if (internationalFee != 0) record.addDynamicRegularFee(internationalFee, "INTERNATIONAL FEE");
                        if (charityDonation != 0) record.addDynamicRegularFee(charityDonation, "CHARITY DONATION");
                        if (depositProcessingFee != 0) record.addDynamicRegularFee(depositProcessingFee, "DEPOSIT PROCESSING FEE");

                    }
                    case "REFUND" -> {
                        if (referenceId.toUpperCase().contains("RETURN")){
                            record.setReturnStatus("Yes");
                        }
                        record.setAmountRefunded(zeroIfNull(record.getAmountRefunded()) + grossTransactionAmount);
                        record.setCommissionRefund(zeroIfNull(record.getCommissionRefund()) + commission);

                        if (regulatoryOperatingFee != 0) record.addDynamicReturnFee(regulatoryOperatingFee, "REGULATORY OPERATING FEE");
                        if (veryHighItemNotAsDescribedFee != 0) record.addDynamicReturnFee(veryHighItemNotAsDescribedFee, "VERY HIGH \"ITEM NOT AS DESCRIBED\" FEE");
                        if (belowStandardPerformancefee != 0) record.addDynamicReturnFee(belowStandardPerformancefee, "BELOW STANDARD PERFORMANCE FEE");
                        if (internationalFee != 0) record.addDynamicReturnFee(internationalFee, "INTERNATIONAL FEE");
                        if (charityDonation != 0) record.addDynamicReturnFee(charityDonation, "CHARITY DONATION");
                        if (depositProcessingFee != 0) record.addDynamicReturnFee(depositProcessingFee, "DEPOSIT PROCESSING FEE");
                    }
                    case "OTHER FEE" -> {
                        record.addDynamicRegularFee(grossTransactionAmount, description);
                    }
                    case "SHIPPING LABEL" -> {
                        record.setReturnShipping(zeroIfNull(record.getReturnShipping()) + grossTransactionAmount);
                    }
                    default -> {
                        actionableSuspense.add(buildSuspenseRow(auditRow, "Action Required: Unmapped Transaction Type (" + type + ") found for Order."));
                    }
                }

                auditTrail.add(auditRow);
            }

            // POST PROCESSING
            for (ReconciliationRecord record : recordsToUpdate.values()) {
                record.calculateCommissionRefundDelta();
            }

            // SAVE UPDATED DATA
            repository.deleteAll(recordsToDelete);
            repository.saveAll(recordsToUpdate.values());
            auditRepository.saveAll(auditTrail);
            suspenseRepository.saveAll(actionableSuspense);
            
            List<String> logs = List.of(
                "Updated " + recordsToUpdate.size() + " Rithum Master Ebay records.",
                "Processed " + (auditTrail.size() + errorSuspense.size() + actionableSuspense.size()) + " Ebay Marketplace rows",
                "Saved " + auditTrail.size() + " Audit rows.",
                "Saved " + actionableSuspense.size() + " Actionable Suspense rows.",
                "Skipped " + errorSuspense.size() + " error rows."
            );


            return new MarketplaceParseResult<>(errorSuspense, actionableSuspense, auditTrail, logs);

        } catch (Exception e) {
            throw new RuntimeException("Failed to parse Ebay Excel file: " + e.getMessage());
        }
    }

    // Null-safe math helpers

    private double invert(Double value) {
        return value != null ? value * -1.0 : 0.0;
    }

    private double zeroIfNull(Double value) {
        return value != null ? value: 0.0;
    }

    private String cleanEbayString(String value) {
        if (value == null) return "";
        String trimmed = value.trim();
        return  trimmed.equals("--") ? "" : trimmed;
    }

    private EbayRawTransaction buildAuditRow(Row row, Map<EbayColumn, Integer> headerMap) {
        return EbayRawTransaction.builder()
                .transactionCreationDate(ExcelUtils.getDateSafe(row, headerMap, EbayColumn.TRANSACTION_CREATION_DATE))
                .type(ExcelUtils.getStringSafe(row, headerMap, EbayColumn.TYPE))
                .orderNumber(ExcelUtils.getStringSafe(row, headerMap, EbayColumn.ORDER_NUMBER))
                .legacyOrderId(ExcelUtils.getStringSafe(row, headerMap, EbayColumn.LEGACY_ORDER_ID))
                .buyerUsername(ExcelUtils.getStringSafe(row, headerMap, EbayColumn.BUYER_USERNAME))
                .buyerName(ExcelUtils.getStringSafe(row, headerMap, EbayColumn.BUYER_NAME))
                .shipToCity(ExcelUtils.getStringSafe(row, headerMap, EbayColumn.SHIP_TO_CITY))
                .shipToProvinceRegionState(ExcelUtils.getStringSafe(row, headerMap, EbayColumn.SHIP_TO_PROVINCE_REGION_STATE))
                .shipToZip(ExcelUtils.getStringSafe(row, headerMap, EbayColumn.SHIP_TO_ZIP))
                .shipToCountry(ExcelUtils.getStringSafe(row, headerMap, EbayColumn.SHIP_TO_COUNTRY))
                .netAmount(ExcelUtils.getDoubleSafe(row, headerMap, EbayColumn.NET_AMOUNT))
                .payoutCurrency(ExcelUtils.getStringSafe(row, headerMap, EbayColumn.PAYOUT_CURRENCY))
                .payoutDate(ExcelUtils.getDateSafe(row, headerMap, EbayColumn.PAYOUT_DATE))
                .payoutId(ExcelUtils.getStringSafe(row, headerMap, EbayColumn.PAYOUT_ID))
                .payoutMethod(ExcelUtils.getStringSafe(row, headerMap, EbayColumn.PAYOUT_METHOD))
                .payoutStatus(ExcelUtils.getStringSafe(row, headerMap, EbayColumn.PAYOUT_STATUS))
                .reasonForHold(ExcelUtils.getStringSafe(row, headerMap, EbayColumn.REASON_FOR_HOLD))
                .itemId(ExcelUtils.getStringSafe(row, headerMap, EbayColumn.ITEM_ID))
                .transactionId(ExcelUtils.getStringSafe(row, headerMap, EbayColumn.TRANSACTION_ID))
                .itemTitle(ExcelUtils.getStringSafe(row, headerMap, EbayColumn.ITEM_TITLE))
                .customLabel(ExcelUtils.getStringSafe(row, headerMap, EbayColumn.CUSTOM_LABEL))
                .quantity(ExcelUtils.getDoubleSafe(row, headerMap, EbayColumn.QUANTITY))
                .itemSubtotal(ExcelUtils.getDoubleSafe(row, headerMap, EbayColumn.ITEM_SUBTOTAL))
                .shippingAndHandling(ExcelUtils.getDoubleSafe(row, headerMap, EbayColumn.SHIPPING_AND_HANDLING))
                .sellerCollectedTax(ExcelUtils.getDoubleSafe(row, headerMap, EbayColumn.SELLER_COLLECTED_TAX))
                .ebayCollectedTax(ExcelUtils.getDoubleSafe(row, headerMap, EbayColumn.EBAY_COLLECTED_TAX))
                .finalValueFeeFixed(ExcelUtils.getDoubleSafe(row, headerMap, EbayColumn.FINAL_VALUE_FEE_FIXED))
                .finalValueFeeVariable(ExcelUtils.getDoubleSafe(row, headerMap, EbayColumn.FINAL_VALUE_FEE_VARIABLE))
                .regulatoryOperatingFee(ExcelUtils.getDoubleSafe(row, headerMap, EbayColumn.REGULATORY_OPERATING_FEE))
                .veryHighItemNotAsDescribedFee(ExcelUtils.getDoubleSafe(row, headerMap, EbayColumn.VERY_HIGH_ITEM_NOT_AS_DESCRIBED_FEE))
                .belowStandardPerformanceFee(ExcelUtils.getDoubleSafe(row, headerMap, EbayColumn.BELOW_STANDARD_PERFORMANCE_FEE))
                .internationalFee(ExcelUtils.getDoubleSafe(row, headerMap, EbayColumn.INTERNATIONAL_FEE))
                .charityDonation(ExcelUtils.getDoubleSafe(row, headerMap, EbayColumn.CHARITY_DONATION))
                .depositProcessingFee(ExcelUtils.getDoubleSafe(row, headerMap, EbayColumn.DEPOSIT_PROCESSING_FEE))
                .grossTransactionAmount(ExcelUtils.getDoubleSafe(row, headerMap, EbayColumn.GROSS_TRANSACTION_AMOUNT))
                .transactionCurrency(ExcelUtils.getStringSafe(row, headerMap, EbayColumn.TRANSACTION_CURRENCY))
                .exchangeRate(ExcelUtils.getDoubleSafe(row, headerMap, EbayColumn.EXCHANGE_RATE))
                .referenceId(ExcelUtils.getStringSafe(row, headerMap, EbayColumn.REFERENCE_ID))
                .description(ExcelUtils.getStringSafe(row, headerMap, EbayColumn.DESCRIPTION))
                .build();
    }

    private EbaySuspense buildSuspenseRow(EbayRawTransaction auditRow, String reason) {
        EbaySuspense suspenseRow = new EbaySuspense();
        BeanUtils.copyProperties(auditRow, suspenseRow);
        suspenseRow.setErrorReason(reason);
        return suspenseRow;
    }
}
