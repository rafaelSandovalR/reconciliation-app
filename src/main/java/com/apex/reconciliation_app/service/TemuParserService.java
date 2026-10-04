package com.apex.reconciliation_app.service;

import com.apex.reconciliation_app.dto.MarketplaceParseResult;
import com.apex.reconciliation_app.enums.TemuColumn;
import com.apex.reconciliation_app.model.ReconciliationRecord;
import com.apex.reconciliation_app.model.temu.TemuRawTransaction;
import com.apex.reconciliation_app.model.temu.TemuSuspense;
import com.apex.reconciliation_app.repository.ReconciliationRepository;
import com.apex.reconciliation_app.repository.TemuRawTransactionRepository;
import com.apex.reconciliation_app.repository.TemuSuspenseRepository;
import com.apex.reconciliation_app.util.ExcelUtils;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jspecify.annotations.NonNull;
import org.springframework.beans.BeanUtils;
import org.springframework.stereotype.Service;

import java.io.InputStream;
import java.time.LocalDateTime;
import java.util.*;

@Service
@RequiredArgsConstructor
public class TemuParserService {

    private final ReconciliationRepository repository;
    private final TemuRawTransactionRepository auditRepository;
    private final TemuSuspenseRepository suspenseRepository;

    private static final java.time.format.DateTimeFormatter TEMU_DATE_FORMATTER =
            java.time.format.DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");

    public MarketplaceParseResult<TemuSuspense, TemuRawTransaction> parseAndUpdate(InputStream inputStream) {
        try (Workbook workbook = new XSSFWorkbook(inputStream)) {
            Sheet sheet = workbook.getSheetAt(0);
            Map<TemuColumn, Integer> headerMap = ExcelUtils.buildHeaderMap(sheet.getRow(0), TemuColumn.class);

            Map<String, ReconciliationRecord> recordsToUpdate = new HashMap<>();
            List<ReconciliationRecord> recordsToDelete = new ArrayList<>();
            Map<String, String> idTranslationMap = new HashMap<>();
            List<TemuRawTransaction> auditTrail = new ArrayList<>();
            List<TemuSuspense> actionableSuspense = new ArrayList<>();
            List<TemuSuspense> errorSuspense = new ArrayList<>();
            Set<String> processedLineIds = new HashSet<>();

            List<TemuRawTransaction> itemLevelRows = new ArrayList<>();
            List<TemuRawTransaction> orderLevelRows = new ArrayList<>();
            Map<String, List<String>> orderToItemIdMap = new HashMap<>();
            Map<String, List<String>> orderToReturnItemIdMap = new HashMap<>();

            // Pass 1: Read all rows, build anchors, build order-level and item-level lists
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                TemuRawTransaction auditRow = buildAuditRow(row, headerMap);
                String type = auditRow.getTransactionType() != null ? auditRow.getTransactionType().trim().toUpperCase() : "";
                String orderId = auditRow.getOrderId() != null ? auditRow.getOrderId().trim() : "";

                if (orderId.isEmpty()) {
                    errorSuspense.add(buildSuspenseRow(auditRow, "Missing Order ID (Non-order line item)"));
                    continue;
                }

                String compositeTransactionId = getCompositeTransactionId(auditRow, type, orderId);

                if (processedLineIds.contains(compositeTransactionId)) {
                    errorSuspense.add(buildSuspenseRow(auditRow, "Duplicate record in upload file, already processed"));
                    continue;
                }
                processedLineIds.add(compositeTransactionId);

                if (auditRepository.existsByCompositeTransactionId(compositeTransactionId)) {
                    errorSuspense.add(buildSuspenseRow(auditRow, "Already processed in a previous upload"));
                    continue;
                }
                auditRow.setCompositeTransactionId(compositeTransactionId);
                String itemId = auditRow.getOrderItemId() != null ? auditRow.getOrderItemId().trim() : "";

                if (itemId.isEmpty()) {
                    orderLevelRows.add(auditRow);
                    continue;
                } else {
                    itemLevelRows.add(auditRow);
                    auditRow.setCompositeId(orderId + "-" + itemId);
                }

                switch (type) {
                    case "ORDER PAYMENT" -> orderToItemIdMap.computeIfAbsent(orderId, k -> new ArrayList<>()).add(itemId);
                    case "REFUND" -> orderToReturnItemIdMap.computeIfAbsent(orderId, k -> new ArrayList<>()).add(itemId);
                }
            }

            // Pass 2: Item-level processing
            // TODO: Can pass 2 and pass one be merged?
            for (TemuRawTransaction row : itemLevelRows) {
                String type = row.getTransactionType() != null ? row.getTransactionType().trim().toUpperCase() : "";
                if (!"ORDER PAYMENT".equals(type) && !"REFUND".equals(type)) {
                    actionableSuspense.add(buildSuspenseRow(row, "Record not recognized. Contains SKU but is not Order or Refund row"));
                    continue;
                }

                String orderId = row.getOrderId() != null ? row.getOrderId().trim() : "";
                String itemId = row.getOrderItemId() != null ? row.getOrderItemId().trim() : "";
                String temuTitle = row.getSku() != null ? row.getSku().trim() : "";
                String compositeId = row.getCompositeId();
                ReconciliationRecord record = recordsToUpdate.get(compositeId);

                if (record == null) {
                    Optional<ReconciliationRecord> dbRecord = repository.findById(compositeId);
                    if (dbRecord.isPresent()) {
                        record = dbRecord.get();
                    } else {
                        List<ReconciliationRecord> baseRecords = repository.findBySiteOrderId(orderId);

                        if (baseRecords.isEmpty()) {
                            actionableSuspense.add(buildSuspenseRow(row, "Missing from Rithum base data"));
                            continue;
                        } else if (baseRecords.size() == 1) {
                            // Fallback 1: Single item order -> Auto-heal
                            record = baseRecords.get(0);
                        } else {
                            // Fallback 2: Multi-item order -> Fuzzy title match
                            List<ReconciliationRecord> matches = new ArrayList<>();
                            for (ReconciliationRecord br : baseRecords) {
                                if (br.getTitle() != null && br.getTitle().contains(temuTitle)) {
                                    matches.add(br);
                                }
                            }

                            if (matches.size() == 1) {
                                record = matches.get(0);
                            } else {
                                actionableSuspense.add(buildSuspenseRow(row, "Ambiguous match: Rithum data missing Item IDs, and title match failed."));
                                continue;
                            }
                        }
                        // Auto-heal rithum record
                        String oldCompositeId = record.getCompositeId();

                        ReconciliationRecord healedRecord = new ReconciliationRecord();
                        BeanUtils.copyProperties(record, healedRecord);

                        healedRecord.setSiteOrderItemId(itemId);
                        healedRecord.setCompositeId(compositeId);

                        idTranslationMap.put(oldCompositeId, compositeId);

                        recordsToDelete.add(record);
                        record = healedRecord;
                    }
                }

                // Enforce Cache Integrity
                recordsToUpdate.put(record.getCompositeId(), record);

                double retailPrice = zeroIfNull(row.getRetailPrice());
                double platformDiscount = invert(row.getPlatformDiscount());
                double sellerDiscount = invert(row.getSellerDiscount());
                double serviceFee = invert(row.getServiceFee());
                double platformIncentive = invert(row.getPlatformIncentive());
                double shipping = invert(row.getShipping());
                double platformIncentiveShipping = invert(row.getPlatformIncentiveShipping());
                double signOnDelivery = invert(row.getSignOnDelivery());
                double others = invert(row.getOthers());

                switch (type) {
                    case "ORDER PAYMENT" -> {
                        record.setSiteOrderAmount(zeroIfNull(record.getSiteOrderAmount()) + retailPrice);
                        record.setSiteOrderFee(zeroIfNull(record.getSiteOrderFee()) + serviceFee);

                        if (platformDiscount != 0) record.addDynamicRegularFee(platformDiscount, "PLATFORM DISCOUNT");
                        if (sellerDiscount != 0) record.addDynamicRegularFee(sellerDiscount, "SELLER DISCOUNT");
                        if (platformIncentive != 0) record.addDynamicRegularFee(platformIncentive, "PLATFORM INCENTIVE");
                        if (shipping != 0) record.addDynamicRegularFee(shipping, "SHIPPING");
                        if (platformIncentiveShipping != 0) record.addDynamicRegularFee(platformIncentiveShipping, "PLATFORM INCENTIVE SHIPPING");
                        if (signOnDelivery != 0) record.addDynamicRegularFee(signOnDelivery, "SIGN ON DELIVERY");
                        if (others != 0) record.addDynamicRegularFee(others, "OTHERS");
                    }
                    case "REFUND" -> {
                        if (zeroIfNull(row.getQuantity()) > 0) record.setReturnStatus("Yes");
                        record.setAmountRefunded(zeroIfNull(record.getAmountRefunded()) + (retailPrice * -1));
                        record.setCommissionRefund(zeroIfNull(record.getCommissionRefund()) + serviceFee);

                        if (platformDiscount != 0) record.addDynamicReturnFee(platformDiscount, "PLATFORM DISCOUNT");
                        if (sellerDiscount != 0) record.addDynamicReturnFee(sellerDiscount, "SELLER DISCOUNT");
                        if (platformIncentive != 0) record.addDynamicReturnFee(platformIncentive, "PLATFORM INCENTIVE");
                        if (shipping != 0) record.addDynamicReturnFee(shipping, "SHIPPING");
                        if (platformIncentiveShipping != 0) record.addDynamicReturnFee(platformIncentiveShipping, "PLATFORM INCENTIVE SHIPPING");
                        if (signOnDelivery != 0) record.addDynamicReturnFee(signOnDelivery, "SIGN ON DELIVERY");
                        if (others != 0) record.addDynamicReturnFee(others, "OTHERS");
                    }
                }

                auditTrail.add(row);
            }

            // Pass 3: Order-level fee distribution
            for (TemuRawTransaction row : orderLevelRows) {
                String type = row.getTransactionType() != null ? row.getTransactionType().trim().toUpperCase() : "";
                String orderId = row.getOrderId() != null ? row.getOrderId().trim() : "";
                boolean isReturnFee = type.contains("RETURN");

                List<ReconciliationRecord> validRecordsToUpdate = new ArrayList<>();

                if (!isReturnFee) {
                    // Standard shipping: Distribute across all items in the order, ignoring the missing item ids
                    List<ReconciliationRecord> baseRecords = repository.findBySiteOrderId(orderId);
                    if (baseRecords != null && !baseRecords.isEmpty()) {
                        for (ReconciliationRecord br : baseRecords) {
                            // Check if DB record was healed in pass 2
                            String activeId = idTranslationMap.getOrDefault(br.getCompositeId(), br.getCompositeId());
                            validRecordsToUpdate.add(recordsToUpdate.getOrDefault(activeId, br));
                        }
                    }
                } else {
                    // Return shipping: Requires strict item anchoring
                    List<String> itemIds = orderToReturnItemIdMap.get(orderId);
                    if (itemIds != null && !itemIds.isEmpty()) {
                        for (String itemId : itemIds) {
                            String compositeId = orderId + "-" + itemId;
                            ReconciliationRecord record = recordsToUpdate.get(compositeId);
                            if (record == null) {
                                Optional<ReconciliationRecord> dbRecord = repository.findById(compositeId);
                                if (dbRecord.isPresent()) record = dbRecord.get();
                            }
                            if (record != null) validRecordsToUpdate.add(record);
                        }
                    }

                    // Return fallback: If unanchored, we can only proceed safely if it is a 1-item order
                    if (validRecordsToUpdate.isEmpty()) {
                        List<ReconciliationRecord> baseRecords = repository.findBySiteOrderId(orderId);
                        if (baseRecords != null && baseRecords.size() == 1) {
                            ReconciliationRecord br = baseRecords.get(0);
                            String activeId = idTranslationMap.getOrDefault(br.getCompositeId(), br.getCompositeId());
                            validRecordsToUpdate.add(recordsToUpdate.getOrDefault(activeId, br));
                        }
                    }
                }

                if (validRecordsToUpdate.isEmpty()) {
                    if (!isReturnFee) {
                        actionableSuspense.add(buildSuspenseRow(row, "Missing from Rithum base data"));
                    } else {
                        // TODO: Future Suspense Scanner - Implement historical DB query to check if this order already has a processed REFUND row.
                        actionableSuspense.add(buildSuspenseRow(row, "Unanchored return fee (Multiple items in order, specific returned item unknown)"));
                    }
                    continue;
                }

                // Route or Distribute
                double others = invert(row.getOthers());
                double splitAmount = others / validRecordsToUpdate.size();

                for (ReconciliationRecord record : validRecordsToUpdate) {
                    switch (type) {
                        case "SHIPPING LABEL PURCHASE" -> {
                            record.setActualShippingCosts(zeroIfNull(record.getActualShippingCosts()) + splitAmount);
                        }
                        case "SHIPPING LABEL FOR RETURN PURCHASE" -> {
                            record.setReturnShipping(zeroIfNull(record.getReturnShipping()) + splitAmount);
                        }
                        default ->  {
                            // TODO: ROUTE SHIPPING LABEL ADJUSTMENT FEES TO SHIPPINGADJUSTMENT COLUMN
                            if (isReturnFee) record.addDynamicReturnFee(splitAmount, type);
                            else record.addDynamicRegularFee(splitAmount, type);
                        }
                    }
                    recordsToUpdate.put(record.getCompositeId(), record);
                }
                auditTrail.add(row);
            }


            for (ReconciliationRecord record : recordsToUpdate.values()) {
                record.calculateCommissionRefundDelta();
            }

            repository.deleteAll(recordsToDelete);
            repository.saveAll(recordsToUpdate.values());
            auditRepository.saveAll(auditTrail);
            suspenseRepository.saveAll(actionableSuspense);

            List<TemuSuspense> allReceiptErrors = new ArrayList<>(actionableSuspense);
            allReceiptErrors.addAll(errorSuspense);

            List<String> logs = List.of(
                    "Updated " + recordsToUpdate.size() + " Rithum Temu records.",
                    "Processed " + (auditTrail.size() + allReceiptErrors.size()) + " Temu marketplace rows",
                    "- Saved " + auditTrail.size() + " audit rows.",
                    "- Saved " + actionableSuspense.size() + " actionable suspense rows.",
                    "- Skipped " + errorSuspense.size() + " error rows (Added to receipt only)"
            );

            return new MarketplaceParseResult<>(allReceiptErrors, auditTrail, logs);
        } catch (Exception e) {
            throw new RuntimeException("Failed to parse Temu Excel file: " + e.getMessage());
        }
    }

    private @NonNull String getCompositeTransactionId(TemuRawTransaction auditRow, String type, String orderId) {
        String dateTime = auditRow.getDateTime() != null ? auditRow.getDateTime().toString() : "";
        String relatedId = auditRow.getRelatedId() != null ? auditRow.getRelatedId().trim() : "";
        String orderItemId = auditRow.getOrderItemId() != null ? auditRow.getOrderItemId().trim() : "";
        Double retailPrice = zeroIfNull(auditRow.getRetailPrice());

        return String.format("%s-%s-%s-%s-%s-%.2f",
                dateTime, type, relatedId, orderId, orderItemId, retailPrice);
    }

    private TemuRawTransaction buildAuditRow(Row row, Map<TemuColumn, Integer> headerMap) {
        return TemuRawTransaction.builder()
                .dateTime(parseTemuDate(row, headerMap))
                .transactionType(ExcelUtils.getStringSafe(row, headerMap, TemuColumn.TRANSACTION_TYPE))
                .relatedId(ExcelUtils.getStringSafe(row, headerMap, TemuColumn.RELATED_ID))
                .orderId(ExcelUtils.getStringSafe(row, headerMap, TemuColumn.ORDER_ID))
                .orderItemId(ExcelUtils.getStringSafe(row, headerMap, TemuColumn.ORDER_ITEM_ID))
                .sku(ExcelUtils.getStringSafe(row, headerMap, TemuColumn.SKU))
                .skuId(ExcelUtils.getStringSafe(row, headerMap, TemuColumn.SKU_ID))
                .quantity(ExcelUtils.getDoubleSafe(row, headerMap, TemuColumn.QUANTITY))
                .shipCity(ExcelUtils.getStringSafe(row, headerMap, TemuColumn.SHIP_CITY))
                .shipState(ExcelUtils.getStringSafe(row, headerMap, TemuColumn.SHIP_STATE))
                .retailPrice(ExcelUtils.getDoubleSafe(row, headerMap, TemuColumn.RETAIL_PRICE))
                .platformDiscount(ExcelUtils.getDoubleSafe(row, headerMap, TemuColumn.PLATFORM_DISCOUNT))
                .sellerDiscount(ExcelUtils.getDoubleSafe(row, headerMap, TemuColumn.SELLER_DISCOUNT))
                .serviceFee(ExcelUtils.getDoubleSafe(row, headerMap, TemuColumn.SERVICE_FEE))
                .serviceFeeTax(ExcelUtils.getDoubleSafe(row, headerMap, TemuColumn.SERVICE_FEE_TAX))
                .platformIncentive(ExcelUtils.getDoubleSafe(row, headerMap, TemuColumn.PLATFORM_INCENTIVE))
                .subtotal(ExcelUtils.getDoubleSafe(row, headerMap, TemuColumn.SUBTOTAL))
                .shipping(ExcelUtils.getDoubleSafe(row, headerMap, TemuColumn.SHIPPING))
                .platformIncentiveShipping(ExcelUtils.getDoubleSafe(row, headerMap, TemuColumn.PLATFORM_INCENTIVE_SHIPPING))
                .productTax(ExcelUtils.getDoubleSafe(row, headerMap, TemuColumn.PRODUCT_TAX))
                .shippingTax(ExcelUtils.getDoubleSafe(row, headerMap, TemuColumn.SHIPPING_TAX))
                .signOnDelivery(ExcelUtils.getDoubleSafe(row, headerMap, TemuColumn.SIGN_ON_DELIVERY))
                .signOnDeliveryTax(ExcelUtils.getDoubleSafe(row, headerMap, TemuColumn.SIGN_ON_DELIVERY_TAX))
                .marketplaceWithheldTax(ExcelUtils.getDoubleSafe(row, headerMap, TemuColumn.MARKETPLACE_WITHHELD_TAX))
                .others(ExcelUtils.getDoubleSafe(row, headerMap, TemuColumn.OTHERS))
                .total(ExcelUtils.getDoubleSafe(row, headerMap, TemuColumn.TOTAL))
                .currency(ExcelUtils.getStringSafe(row, headerMap, TemuColumn.CURRENCY))
                .build();
    }

    private TemuSuspense buildSuspenseRow(TemuRawTransaction auditRow, String reason) {
        TemuSuspense suspenseRow = new TemuSuspense();
        BeanUtils.copyProperties(auditRow, suspenseRow);
        suspenseRow.setErrorReason(reason);
        return suspenseRow;
    }

    private double invert(Double value) {
        return value != null ? value * -1.0 : 0.0;
    }

    private double zeroIfNull(Double value) {
        return value != null ? value: 0.0;
    }

    private LocalDateTime parseTemuDate(Row row, Map<TemuColumn, Integer> headerMap) {
        LocalDateTime nativeDate = ExcelUtils.getDateSafe(row, headerMap, TemuColumn.DATE_TIME);
        if (nativeDate != null) return nativeDate;

        String dateStr = ExcelUtils.getStringSafe(row, headerMap, TemuColumn.DATE_TIME);
        if (dateStr != null && !dateStr.trim().isEmpty()) {
            try {
                return LocalDateTime.parse(dateStr.trim(), TEMU_DATE_FORMATTER);
            } catch (Exception e) {
                // If it fails, we will catch it later in the suspense checks
            }
        }
        return null;
    }
}
