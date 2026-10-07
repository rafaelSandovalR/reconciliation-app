package com.apex.reconciliation_app.service;

import com.apex.reconciliation_app.dto.MarketplaceParseResult;
import com.apex.reconciliation_app.enums.ShopSimonColumn;
import com.apex.reconciliation_app.model.ReconciliationRecord;
import com.apex.reconciliation_app.model.shopsimon.ShopSimonRawTransaction;
import com.apex.reconciliation_app.model.shopsimon.ShopSimonSuspense;
import com.apex.reconciliation_app.repository.ReconciliationRepository;
import com.apex.reconciliation_app.repository.ShopSimonRawTransactionRepository;
import com.apex.reconciliation_app.repository.ShopSimonSuspenseRepository;
import com.apex.reconciliation_app.util.ExcelUtils;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.BeanUtils;
import org.springframework.stereotype.Service;

import java.io.InputStream;
import java.time.LocalDateTime;
import java.util.*;

@Service
@RequiredArgsConstructor
public class ShopSimonParserService {

    private final ReconciliationRepository repository;
    private final ShopSimonRawTransactionRepository auditRepository;
    private final ShopSimonSuspenseRepository suspenseRepository;


    public MarketplaceParseResult<ShopSimonSuspense, ShopSimonRawTransaction> parseAndUpdate(InputStream inputStream) {
        try (Workbook workbook = new XSSFWorkbook(inputStream)) {

            Sheet sheet = workbook.getSheetAt(0);
            Map<ShopSimonColumn, Integer> headerMap = ExcelUtils.buildHeaderMap(sheet.getRow(0), ShopSimonColumn.class);
            Map<String, ReconciliationRecord> recordsToUpdate = new HashMap<>();
            List<ShopSimonRawTransaction> auditTrail = new ArrayList<>();
            List<ShopSimonSuspense> actionableSuspense = new ArrayList<>();
            List<ShopSimonSuspense> errorSuspense = new ArrayList<>();
            Set<String> processedLineIds = new HashSet<>();


            // PROCESSING LOGIC
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                ShopSimonRawTransaction auditRow = buildAuditRow(row, headerMap);
                String poNumber = auditRow.getPoNumber() != null ? auditRow.getPoNumber().trim() : "";
                String sku = auditRow.getSku() != null ? auditRow.getSku().trim() : "";


                if (poNumber.isEmpty() || sku.isEmpty()) {
                    errorSuspense.add(buildSuspenseRow(auditRow, "Missing Purchase Order or SKU (Non-order line item"));
                    continue;
                }
                String compositeId = poNumber + "-" + sku;
                auditRow.setCompositeId(compositeId);
                String type = auditRow.getType() != null ? auditRow.getType().trim().toUpperCase() : "";
                String dateTime = auditRow.getTransactionDate() != null ? auditRow.getTransactionDate().toString() : "";
                double totalConsumerPrice = auditRow.getTotalConsumerPrice() != null ? auditRow.getTotalConsumerPrice() : 0.0;

                String compositeTransactionId = String.format("%s-%s-%s-%s-%.2f",
                        dateTime, type, poNumber, sku, totalConsumerPrice);
                auditRow.setCompositeTransactionId(compositeTransactionId);

                // IDEMPOTENCY CHECKS
                if (processedLineIds.contains(compositeTransactionId)) {
                    errorSuspense.add(buildSuspenseRow(auditRow, "Duplicate record in upload file, already processed"));
                    continue;
                }
                processedLineIds.add(compositeTransactionId);

                if (auditRepository.existsByCompositeTransactionId(compositeTransactionId)) {
                    errorSuspense.add(buildSuspenseRow(auditRow, "Already processed in a previous upload"));
                    continue;
                }

                // FETCH FROM CACHE OR DB
                ReconciliationRecord record = recordsToUpdate.get(compositeTransactionId);
                if (record == null) {
                    Optional<ReconciliationRecord> dbRecord = repository.findById(compositeId);
                    if (dbRecord.isPresent()) {
                        record = dbRecord.get();
                        recordsToUpdate.put(compositeId, record);
                    } else {
                        actionableSuspense.add(buildSuspenseRow(auditRow, "Missing from Rithum base data"));
                        continue;
                    }
                }

                double totalCommissionAmount = auditRow.getTotalCommissionAmount() != null ? auditRow.getTotalCommissionAmount() : 0.0;

                switch (type) {
                    case "ORDER" -> {
                        record.setSiteOrderAmount((record.getSiteOrderAmount() != null ? record.getSiteOrderAmount() : 0.0) + totalConsumerPrice);
                        record.setSiteOrderFee((record.getSiteOrderFee() != null ? record.getSiteOrderFee() : 0.0) + totalCommissionAmount);
                    }
                    case "RETURN" -> {
                        record.setReturnStatus("Yes");
                        record.setAmountRefunded((record.getAmountRefunded() != null ? record.getAmountRefunded() : 0.0) + totalConsumerPrice);
                        record.setCommissionRefund((record.getCommissionRefund() != null ? record.getCommissionRefund() : 0.0) + totalCommissionAmount);
                    }
                }
                auditTrail.add(auditRow);
            }

            for (ReconciliationRecord record : recordsToUpdate.values()) {
                record.calculateCommissionRefundDelta();
            }

            repository.saveAll(recordsToUpdate.values());
            auditRepository.saveAll(auditTrail);
            suspenseRepository.saveAll(actionableSuspense);

            List<ShopSimonSuspense> allReceiptErrors = new ArrayList<>(actionableSuspense);
            allReceiptErrors.addAll(errorSuspense);

            List<String> logs = List.of(
                    "Updated " + recordsToUpdate.size() + " Rithum ShopSimon base records.",
                    "Processed " + (auditTrail.size() + allReceiptErrors.size()) + " ShopSimon marketplace rows",
                    "-> Saved " + auditTrail.size() + " audit rows.",
                    "-> Saved " + actionableSuspense.size() + " actionable suspense rows.",
                    "-> Skipped " + errorSuspense.size() + " error rows (Added to receipt only)"
            );

            return new MarketplaceParseResult<>(allReceiptErrors, auditTrail, logs);

        } catch (Exception e) {
            throw new RuntimeException("Failed to parse Shop Simon Excel file: " + e.getMessage());
        }
    }

    private ShopSimonRawTransaction buildAuditRow(Row row, Map<ShopSimonColumn, Integer> headerMap) {
        return ShopSimonRawTransaction.builder()
                .type(ExcelUtils.getStringSafe(row, headerMap, ShopSimonColumn.TYPE))
                .entry(ExcelUtils.getStringSafe(row, headerMap, ShopSimonColumn.ENTRY))
                .orderId(ExcelUtils.getStringSafe(row, headerMap, ShopSimonColumn.ORDER_ID))
                .partnerId(ExcelUtils.getStringSafe(row, headerMap, ShopSimonColumn.PARTNER_ID))
                .poNumber(ExcelUtils.getStringSafe(row, headerMap, ShopSimonColumn.PO_NUMBER))
                .consumerOrderNumber(ExcelUtils.getStringSafe(row, headerMap, ShopSimonColumn.CONSUMER_ORDER_NUMBER))
                .supplierOrderNumber(ExcelUtils.getStringSafe(row, headerMap, ShopSimonColumn.SUPPLIER_ORDER_NUMBER))
                .currencyCode(ExcelUtils.getStringSafe(row, headerMap, ShopSimonColumn.CURRENCY_CODE))
                .lineNumber(ExcelUtils.getDoubleSafe(row, headerMap, ShopSimonColumn.LINE_NUMBER))
                .sku(ExcelUtils.getStringSafe(row, headerMap, ShopSimonColumn.SKU))
                .partnerSku(ExcelUtils.getStringSafe(row, headerMap, ShopSimonColumn.PARTNER_SKU))
                .itemId(ExcelUtils.getStringSafe(row, headerMap, ShopSimonColumn.ITEM_ID))
                .transactionDate(parseShopSimonDate(row, headerMap, ShopSimonColumn.TRANSACTION_DATE))
                .quantityShipped(ExcelUtils.getDoubleSafe(row, headerMap, ShopSimonColumn.QUANTITY_SHIPPED))
                .amountShipped(ExcelUtils.getDoubleSafe(row, headerMap, ShopSimonColumn.AMOUNT_SHIPPED))
                .returnId(ExcelUtils.getStringSafe(row, headerMap, ShopSimonColumn.RETURN_ID))
                .quantityReturned(ExcelUtils.getDoubleSafe(row, headerMap, ShopSimonColumn.QUANTITY_RETURNED))
                .lineItemFulfillment(ExcelUtils.getStringSafe(row, headerMap, ShopSimonColumn.LINE_ITEM_FULFILLMENT))
                .consumerUnitPrice(ExcelUtils.getDoubleSafe(row, headerMap, ShopSimonColumn.CONSUMER_UNIT_PRICE))
                .consumerUnitPriceIncludingTax(ExcelUtils.getDoubleSafe(row, headerMap, ShopSimonColumn.CONSUMER_UNIT_PRICE_INCLUDING_TAX))
                .taxType(ExcelUtils.getStringSafe(row, headerMap, ShopSimonColumn.TAX_TYPE))
                .taxPercentage(ExcelUtils.getDoubleSafe(row, headerMap, ShopSimonColumn.TAX_PERCENTAGE))
                .unitTaxAmount(ExcelUtils.getDoubleSafe(row, headerMap, ShopSimonColumn.UNIT_TAX_AMOUNT))
                .totalTax(ExcelUtils.getDoubleSafe(row, headerMap, ShopSimonColumn.TOTAL_TAX))
                .consumerUnitPriceExcludingTax(ExcelUtils.getDoubleSafe(row, headerMap, ShopSimonColumn.CONSUMER_UNIT_PRICE_EXCLUDING_TAX))
                .commissionPercentage(ExcelUtils.getDoubleSafe(row, headerMap, ShopSimonColumn.COMMISSION_PERCENTAGE))
                .totalCommissionAmount(ExcelUtils.getDoubleSafe(row, headerMap, ShopSimonColumn.TOTAL_COMMISSION_AMOUNT))
                .adjustmentReason(ExcelUtils.getStringSafe(row, headerMap, ShopSimonColumn.ADJUSTMENT_REASON))
                .transactionAmount(ExcelUtils.getDoubleSafe(row, headerMap, ShopSimonColumn.TRANSACTION_AMOUNT))
                .totalAmountPaid(ExcelUtils.getDoubleSafe(row, headerMap, ShopSimonColumn.TOTAL_AMOUNT_PAID))
                .totalAmountOutstanding(ExcelUtils.getDoubleSafe(row, headerMap, ShopSimonColumn.TOTAL_AMOUNT_OUTSTANDING))
                .totalConsumerPrice(ExcelUtils.getDoubleSafe(row, headerMap, ShopSimonColumn.TOTAL_CONSUMER_PRICE))
                .trackingNumber(ExcelUtils.getStringSafe(row, headerMap, ShopSimonColumn.TRACKING_NUMBER))
                .invoiceId(ExcelUtils.getStringSafe(row, headerMap, ShopSimonColumn.INVOICE_ID))
                .supplierInvoiceNumber(ExcelUtils.getStringSafe(row, headerMap, ShopSimonColumn.SUPPLIER_INVOICE_NUMBER))
                .invoiceRecordType(ExcelUtils.getStringSafe(row, headerMap, ShopSimonColumn.INVOICE_RECORD_TYPE))
                .suborderId(ExcelUtils.getStringSafe(row, headerMap, ShopSimonColumn.SUBORDER_ID))
                .invoiceDate(parseShopSimonDate(row, headerMap, ShopSimonColumn.INVOICE_DATE))
                .quantityInvoiced(ExcelUtils.getDoubleSafe(row, headerMap, ShopSimonColumn.QUANTITY_INVOICED))
                .orderQuantity(ExcelUtils.getDoubleSafe(row, headerMap, ShopSimonColumn.ORDER_QUANTITY))
                .unitCost(ExcelUtils.getDoubleSafe(row, headerMap, ShopSimonColumn.UNIT_COST))
                .totalUnitCost(ExcelUtils.getDoubleSafe(row, headerMap, ShopSimonColumn.TOTAL_UNIT_COST))
                .freightAmount(ExcelUtils.getDoubleSafe(row, headerMap, ShopSimonColumn.FREIGHT_AMOUNT))
                .salesTaxAmount(ExcelUtils.getDoubleSafe(row, headerMap, ShopSimonColumn.SALES_TAX_AMOUNT))
                .charges(ExcelUtils.getDoubleSafe(row, headerMap, ShopSimonColumn.CHARGES))
                .credits(ExcelUtils.getDoubleSafe(row, headerMap, ShopSimonColumn.CREDITS))
                .subtotalExcludingLineItems(ExcelUtils.getDoubleSafe(row, headerMap, ShopSimonColumn.SUBTOTAL_EXCLUDING_LINE_ITEMS))
                .build();
    }

    private ShopSimonSuspense buildSuspenseRow(ShopSimonRawTransaction auditRow, String reason) {
        ShopSimonSuspense suspenseRow = new ShopSimonSuspense();
        BeanUtils.copyProperties(auditRow, suspenseRow);
        suspenseRow.setErrorReason(reason);
        return suspenseRow;
    }

    private LocalDateTime parseShopSimonDate(Row row, Map<ShopSimonColumn, Integer> headerMap, ShopSimonColumn column) {
        LocalDateTime nativeDate = ExcelUtils.getDateSafe(row, headerMap, column);
        if (nativeDate != null ) return nativeDate;

        String dateStr = ExcelUtils.getStringSafe(row, headerMap, column);
        if (dateStr != null && !dateStr.trim().isEmpty()) {
            try {
                return java.time.OffsetDateTime.parse(dateStr.trim()).toLocalDateTime();
            } catch (Exception e) {
                // TODO: Handle if missing
            }
        }
        return null;
    }
}
