package com.apex.reconciliation_app.service;

import com.apex.reconciliation_app.dto.MarketplaceParseResult;
import com.apex.reconciliation_app.enums.TemuColumn;
import com.apex.reconciliation_app.model.ReconciliationRecord;
import com.apex.reconciliation_app.model.TemuRawTransaction;
import com.apex.reconciliation_app.model.TemuSuspense;
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

    public MarketplaceParseResult<TemuSuspense, TemuRawTransaction> parseAndUpdate(InputStream inputStream) {
        try (Workbook workbook = new XSSFWorkbook(inputStream)) {
            Sheet sheet = workbook.getSheetAt(0);
            Map<TemuColumn, Integer> headerMap = ExcelUtils.buildHeaderMap(sheet.getRow(0), TemuColumn.class);

            Map<String, ReconciliationRecord> recordsToUpdate = new HashMap<>();
            List<TemuRawTransaction> auditTrail = new ArrayList<>();
            List<TemuSuspense> actionableSuspense = new ArrayList<>();
            List<TemuSuspense> errorSuspense = new ArrayList<>();
            Set<String> processedLineIds = new HashSet<>();

            List<TemuRawTransaction> itemLevelRows = new ArrayList<>();
            List<TemuRawTransaction> orderLevelRows = new ArrayList<>();
            Map<String, List<String>> orderToSkuMap = new HashMap<>();

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
                if (auditRow.getSku() == null || auditRow.getSku().trim().isEmpty()) {
                    orderLevelRows.add(auditRow);
                } else {
                    itemLevelRows.add(auditRow);
                }

                if ("ORDER PAYMENT".equals(type)) {
                    orderToSkuMap.computeIfAbsent(auditRow.getOrderId(), k -> new ArrayList<>()).add(auditRow.getSku());
                }
            }

            // Pass 2: Item-level processing
            for (TemuRawTransaction row : itemLevelRows) {
                String type = row.getTransactionType() != null ? row.getTransactionType().trim().toUpperCase() : "";
                if (!"ORDER PAYMENT".equals(type) && !"REFUND".equals(type)) {
                    actionableSuspense.add(buildSuspenseRow(row, "Record not recognized. Contains SKU but is not Order or Refund row"));
                    continue;
                }

                String compositeId = row.getOrderId() + "-" + row.getSku();
                ReconciliationRecord record = recordsToUpdate.get(compositeId);
                if (record == null) {
                    Optional<ReconciliationRecord> dbRecord = repository.findById(compositeId);
                    if (dbRecord.isPresent()) {
                        record = dbRecord.get();
                        recordsToUpdate.put(compositeId, record);
                    } else {
                        actionableSuspense.add(buildSuspenseRow(row, "Missing from Rithum base data"));
                        continue;
                    }
                }

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
                        record.setReturnStatus("Yes");
                        record.setAmountRefunded(zeroIfNull(record.getAmountRefunded()) + (retailPrice * -1));
                        record.setCommissionRefund(zeroIfNull(record.getCommissionRefund()) + (serviceFee * -1));

                        if (platformDiscount != 0) record.addDynamicReturnFee(platformDiscount, "PLATFORM DISCOUNT");
                        if (sellerDiscount != 0) record.addDynamicReturnFee(sellerDiscount, "SELLER DISCOUNT");
                        if (platformIncentive != 0) record.addDynamicReturnFee(platformIncentive, "PLATFORM INCENTIVE");
                        if (shipping != 0) record.addDynamicReturnFee(shipping, "SHIPPING");
                        if (platformIncentiveShipping != 0) record.addDynamicReturnFee(platformIncentiveShipping, "PLATFORM INCENTIVE SHIPPING");
                        if (signOnDelivery != 0) record.addDynamicReturnFee(signOnDelivery, "SIGN ON DELIVERY");
                        if (others != 0) record.addDynamicReturnFee(others, "OTHERS");
                    }
                }
            }

            // Pass 3: Order-level fee distribution
            for (TemuRawTransaction row : orderLevelRows) {
                List<String> skus = orderToSkuMap.get(row.getOrderId());
                if (skus == null || skus.isEmpty()) {
                    actionableSuspense.add(buildSuspenseRow(row, "Unanchored order-level fee (No associated Order Payment found)"));
                    continue;
                }

                String type = row.getTransactionType() != null ? row.getTransactionType().trim().toUpperCase() : "";
                double others = invert(row.getOthers());
                double splitAmount = others / skus.size();

                for (String sku : skus) {
                    String compositeId = row.getOrderId() + "-" + sku;
                    ReconciliationRecord record = recordsToUpdate.get(compositeId);
                    if (record == null) {
                        Optional<ReconciliationRecord> dbRecord = repository.findById(compositeId);
                        if (dbRecord.isPresent()) {
                            record = dbRecord.get();
                            recordsToUpdate.put(compositeId, record);
                            switch (type) {
                                case "SHIPPING LABEL PURCHASE" -> {
                                    record.setActualShippingCosts(zeroIfNull(record.getActualShippingCosts()) + splitAmount);
                                }
                                case "SHIPPING LABEL FOR RETURN PURCHASE" -> {
                                    record.setReturnShipping(zeroIfNull(record.getReturnShipping()) + splitAmount);
                                }
                                default ->  {
                                    if (type.contains("RETURN")) {
                                        record.addDynamicReturnFee(splitAmount, type);
                                    } else {
                                        record.addDynamicRegularFee(splitAmount, type);
                                    }
                                }
                            }
                        } else {
                            actionableSuspense.add(buildSuspenseRow(row, "Missing from Rithum base data"));
                            continue;
                        }
                    }
                }
            }


            for (ReconciliationRecord record : recordsToUpdate.values()) {
                record.calculateCommissionRefundDelta();
            }

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

    private static @NonNull String getCompositeTransactionId(TemuRawTransaction auditRow, String type, String orderId) {
        LocalDateTime dateTime = auditRow.getDateTime() != null ? auditRow.getDateTime() : null;
        String relatedId = auditRow.getRelatedId() != null ? auditRow.getRelatedId().trim() : "";
        String orderItemId = auditRow.getOrderItemId() != null ? auditRow.getOrderItemId().trim() : "";
        String sku = auditRow.getSku() != null ? auditRow.getSku().trim() : "";
        String skuId = auditRow.getSkuId() != null ? auditRow.getSkuId().trim() : "";

        return String.format("%s-%s-%s-%s-%s-%s-%s",
                dateTime, type, relatedId, orderId, orderItemId, sku, skuId);
    }

    private TemuRawTransaction buildAuditRow(Row row, Map<TemuColumn, Integer> headerMap) {
        return TemuRawTransaction.builder()
                .dateTime(ExcelUtils.getDateSafe(row, headerMap, TemuColumn.DATE_TIME))
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
}
