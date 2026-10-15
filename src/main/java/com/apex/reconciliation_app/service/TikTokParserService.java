package com.apex.reconciliation_app.service;

import com.apex.reconciliation_app.dto.MarketplaceParseResult;
import com.apex.reconciliation_app.enums.TikTokColumn;
import com.apex.reconciliation_app.model.ReconciliationRecord;
import com.apex.reconciliation_app.model.tiktok.TikTokRawTransaction;
import com.apex.reconciliation_app.model.tiktok.TikTokSuspense;
import com.apex.reconciliation_app.repository.ReconciliationRepository;
import com.apex.reconciliation_app.repository.TikTokRawTransactionRepository;
import com.apex.reconciliation_app.repository.TikTokSuspenseRepository;
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
public class TikTokParserService {

    private final ReconciliationRepository repository;
    private final TikTokRawTransactionRepository auditRepository;
    private final TikTokSuspenseRepository suspenseRepository;

    public MarketplaceParseResult<TikTokSuspense, TikTokRawTransaction> parseAndUpdate(InputStream inputStream) {
        try (Workbook workbook = new XSSFWorkbook(inputStream)) {
            Sheet sheet = workbook.getSheetAt(0);
            Map<TikTokColumn, Integer> headerMap = ExcelUtils.buildHeaderMap(sheet.getRow(0), TikTokColumn.class);
            Map<String, ReconciliationRecord> recordsToUpdate = new HashMap<>();
            Map<String, String> tiktokToRithumIdMap = new HashMap<>();
            List<TikTokRawTransaction> auditTrail = new ArrayList<>();
            List<TikTokSuspense> actionableSuspense = new ArrayList<>();
            List<TikTokSuspense> errorSuspense = new ArrayList<>();
            Set<String> processedLineIds = new HashSet<>();

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                TikTokRawTransaction auditRow = buildAuditRow(row, headerMap);
                String type = cleanTikTokString(auditRow.getType());
                String orderId = cleanTikTokString(auditRow.getOrderAdjustmentId());
                String skuId = cleanTikTokString(auditRow.getSkuId());

                // Filter irrelevant rows
                if (!"Order".equalsIgnoreCase(type)) {
                    errorSuspense.add(buildSuspenseRow(auditRow, "Non-order line item"));
                    continue;
                }
                if (orderId.isEmpty()) {
                    errorSuspense.add(buildSuspenseRow(auditRow, "Missing Order ID"));
                }

                // Filter zero-value rows
                double totalSettlement = zeroIfNull(auditRow.getTotalSettlementAmount());
                double grossSales = zeroIfNull(auditRow.getGrossSales());
                double grossRefund = zeroIfNull(auditRow.getGrossSalesRefund());
                double shipping = zeroIfNull(auditRow.getShipping());
                double fees = zeroIfNull(auditRow.getFees());

                if (totalSettlement == 0 && grossSales == 0 && grossRefund == 0 && shipping == 0 && fees == 0) {
                    auditTrail.add(auditRow);
                    continue;
                }

                // Idempotency Check
                String statementDateStr = auditRow.getStatementDate() != null ? auditRow.getStatementDate().toString() : "";
                String compositeTransactionId = String.format("%s-%s-%s-%.2f",
                        statementDateStr, orderId, skuId, totalSettlement);
                auditRow.setCompositeTransactionId(compositeTransactionId);

                if (processedLineIds.contains(compositeTransactionId)) {
                    errorSuspense.add(buildSuspenseRow(auditRow, "Duplicate record in upload file"));
                    continue;
                }
                processedLineIds.add(compositeTransactionId);

                if (auditRepository.existsByCompositeTransactionId(compositeTransactionId)) {
                    errorSuspense.add(buildSuspenseRow(auditRow, "Already processed in previous upload."));
                    continue;
                }

                String constructedTitle = buildConstructedTitle(auditRow.getProductName(), auditRow.getSkuName());

                // Cascading database fallback & auto-heal
                ReconciliationRecord record = null;
                String tiktokKey = orderId + "-" + skuId;
                String rithumCompositeId = tiktokToRithumIdMap.get(tiktokKey);

                // Cache check
                if (rithumCompositeId != null) {
                    record = recordsToUpdate.get(rithumCompositeId);
                }

                // Db fallback
                if (record == null) {
                    List<ReconciliationRecord> baseRecords = repository.findBySiteOrderId(orderId);

                    if (baseRecords.isEmpty()) {
                        actionableSuspense.add(buildSuspenseRow(auditRow, "Missing from Rithum base data."));
                        continue;
                    } else if (baseRecords.size() == 1) {
                        record = baseRecords.get(0);
                    } else {
                        // Fuzzy title match for multi-item orders
                        List<ReconciliationRecord> titleMatches = new ArrayList<>();
                        for (ReconciliationRecord br : baseRecords) {
                            if (br.getTitle() != null && !constructedTitle.isEmpty() && br.getTitle().contains(constructedTitle)) {
                                titleMatches.add(br);
                            }
                        }
                        if (titleMatches.size() == 1) {
                            record = titleMatches.get(0);
                        } else {
                            actionableSuspense.add(buildSuspenseRow(auditRow, "Ambigious match: Fuzzy title match failed."));
                            continue;
                        }
                    }

                    // Protect against overwritign if the DB record was already pulled into the cache by previous row
                    if (recordsToUpdate.containsKey(record.getCompositeId())) {
                        record = recordsToUpdate.get(record.getCompositeId());
                    }

                    tiktokToRithumIdMap.put(tiktokKey, record.getCompositeId());
                }

                auditRow.setCompositeId(record.getCompositeId());
                recordsToUpdate.put(record.getCompositeId(), record);


                // Isolate all return-specific shipping
                double invertedReturnShipping = invert(auditRow.getReturnShippingFee());
                double invertedReturnShipReimbursement = invert(auditRow.getReturnShippingFeeReimbursement());
                double invertedReturnLabel = invert(auditRow.getReturnShippingLabelFee());
                double invertedReturnShippingPaidByCustomer = invert(auditRow.getReturnShippingFeePaidByCustomers());

                double totalReturnShipping = invertedReturnShipping + invertedReturnShipReimbursement + invertedReturnLabel + invertedReturnShippingPaidByCustomer;

                // Isolate all return-specific fees
                double refundAdminFee = invert(auditRow.getRefundAdministrationFee());
                double affiliateCommissionRefund = invert(auditRow.getAffiliateCommissionRefund());

                double totalReturnFees = refundAdminFee + affiliateCommissionRefund;

                // Calculate true forward costs
                double invertedShipping = invert(shipping);
                double invertedFees = invert(fees);

                double trueForwardShipping = invertedShipping - totalReturnShipping;
                double trueForwardFees = invertedFees - totalReturnFees;

                // Math Routing
                if (grossSales != 0) {
                    record.setSiteOrderAmount(zeroIfNull(record.getSiteOrderAmount()) + grossSales);
                    record.setActualShippingCosts(zeroIfNull(record.getActualShippingCosts()) + trueForwardShipping);
                    record.setSiteOrderFee(zeroIfNull(record.getSiteOrderFee()) + trueForwardFees);
                }

                if (grossRefund != 0) {
                    record.setReturnStatus("Yes");
                    record.setAmountRefunded(zeroIfNull(record.getAmountRefunded()) + (grossRefund * -1));

                    record.setReturnShipping(zeroIfNull(record.getReturnShipping()) + (invertedReturnShipping + invertedReturnLabel));
                    record.addDynamicReturnFee(invertedReturnShipReimbursement, "RETURN SHIPPING FEE REIMBURSEMENT");
                    record.addDynamicReturnFee(invertedReturnShippingPaidByCustomer, "RETURN SHIPPING PAID BY CUSTOMER");

                    record.addDynamicReturnFee(refundAdminFee, "REFUND ADMINISTRATION FEE");
                    record.addDynamicReturnFee(affiliateCommissionRefund, "AFFILIATE COMMISSION REFUND");

                    // Handle Refund-Only rows where commission is fully returned
                    if (grossSales == 0) {
                        // TODO: Check if this is correct
                        record.setReturnShipping(zeroIfNull(record.getReturnShipping()) + trueForwardShipping);
                        record.setCommissionRefund(zeroIfNull(record.getCommissionRefund()) + trueForwardFees);
                    }
                }
            }

            for (ReconciliationRecord record : recordsToUpdate.values()) {
                record.calculateCommissionRefundDelta();
            }

            repository.saveAll(recordsToUpdate.values());
            auditRepository.saveAll(auditTrail);
            suspenseRepository.saveAll(actionableSuspense);

            List<String> logs = List.of(
                    "Updated " + recordsToUpdate.size() + " Rithum TikTok records.",
                    "Processed " + (auditTrail.size() + errorSuspense.size() + actionableSuspense.size()) + " TikTok rows",
                    "Saved " + auditTrail.size() + " Audit rows.",
                    "Saved " + actionableSuspense.size() + " Actionable Suspense rows.",
                    "Skipped " + errorSuspense.size() + " error rows."
            );

            return new MarketplaceParseResult<>(errorSuspense, actionableSuspense, auditTrail, logs);
        } catch (Exception e) {
            throw new RuntimeException("Failed to parse TikTok Excel file: " + e.getMessage());
        }
    }

    private String buildConstructedTitle(String productName, String skuName) {
        String cleanProd = productName != null ? productName.replaceAll("\\s+", " ").trim() : "";
        String size = "";
        if (skuName != null) {
            size = skuName.split(",")[0].trim();
        }
        return (cleanProd + " Size " + size).trim();
    }

    private String cleanTikTokString(String value) {
        if (value == null) return "";
        String trimmed = value.trim();
        return trimmed.equals("/") ? "" : trimmed;
    }

    private double invert(Double value) {
        return value != null ? value * -1.0 : 0.0;
    }

    private double zeroIfNull(Double value) {
        return value != null ? value: 0.0;
    }

    private TikTokRawTransaction buildAuditRow(Row row, Map<TikTokColumn, Integer> headerMap) {
        return TikTokRawTransaction.builder()
                .statementDate(ExcelUtils.getDateSafe(row, headerMap, TikTokColumn.STATEMENT_DATE))
                .statementId(ExcelUtils.getStringSafe(row, headerMap, TikTokColumn.STATEMENT_ID))
                .paymentId(ExcelUtils.getStringSafe(row, headerMap, TikTokColumn.PAYMENT_ID))
                .status(ExcelUtils.getStringSafe(row, headerMap, TikTokColumn.STATUS))
                .currency(ExcelUtils.getStringSafe(row, headerMap, TikTokColumn.CURRENCY))
                .type(ExcelUtils.getStringSafe(row, headerMap, TikTokColumn.TYPE))
                .orderAdjustmentId(ExcelUtils.getStringSafe(row, headerMap, TikTokColumn.ORDER_ADJUSTMENT_ID))
                .skuId(ExcelUtils.getStringSafe(row, headerMap, TikTokColumn.SKU_ID))
                .quantity(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.QUANTITY))
                .productName(ExcelUtils.getStringSafe(row, headerMap, TikTokColumn.PRODUCT_NAME))
                .skuName(ExcelUtils.getStringSafe(row, headerMap, TikTokColumn.SKU_NAME))
                .orderCreatedDate(ExcelUtils.getDateSafe(row, headerMap, TikTokColumn.ORDER_CREATED_DATE))
                .orderShipmentDate(ExcelUtils.getDateSafe(row, headerMap, TikTokColumn.ORDER_SHIPMENT_DATE))
                .orderDeliveryDate(ExcelUtils.getDateSafe(row, headerMap, TikTokColumn.ORDER_DELIVERY_DATE))
                .totalSettlementAmount(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.TOTAL_SETTLEMENT_AMOUNT))
                .netSales(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.NET_SALES))
                .grossSales(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.GROSS_SALES))
                .grossSalesRefund(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.GROSS_SALES_REFUND))
                .sellerDiscount(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.SELLER_DISCOUNT))
                .sellerDiscountRefund(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.SELLER_DISCOUNT_REFUND))
                .shipping(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.SHIPPING))
                .tiktokShopShippingFee(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.TIKTOK_SHOP_SHIPPING_FEE))
                .fulfilledByTiktokShopShippingFee(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.FULFILLED_BY_TIKTOK_SHOP_SHIPPING_FEE))
                .signatureConfirmationServiceFee(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.SIGNATURE_CONFIRMATION_SERVICE_FEE))
                .shippingInsuranceFee(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.SHIPPING_INSURANCE_FEE))
                .customerPaidShippingFee(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.CUSTOMER_PAID_SHIPPING_FEE))
                .customerPaidShippingFeeRefund(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.CUSTOMER_PAID_SHIPPING_FEE_REFUND))
                .tiktokShopShippingIncentive(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.TIKTOK_SHOP_SHIPPING_INCENTIVE))
                .tiktokShopShippingIncentiveRefund(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.TIKTOK_SHOP_SHIPPING_INCENTIVE_REFUND))
                .shippingFeeSubsidy(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.SHIPPING_FEE_SUBSIDY))
                .returnShippingFee(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.RETURN_SHIPPING_FEE))
                .fbtFulfillmentFee(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.FBT_FULFILLMENT_FEE))
                .customerShippingFeeOffset(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.CUSTOMER_SHIPPING_FEE_OFFSET))
                .shippingFeeDiscount(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.SHIPPING_FEE_DISCOUNT))
                .returnShippingLabelFee(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.RETURN_SHIPPING_LABEL_FEE))
                .fbtFulfillmentFeeReimbursement(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.FBT_FULFILLMENT_FEE_REIMBURSEMENT))
                .returnShippingFeePaidByCustomers(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.RETURN_SHIPPING_FEE_PAID_BY_CUSTOMERS))
                .returnShippingFeeReimbursement(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.RETURN_SHIPPING_FEE_REIMBURSEMENT))
                .shippingAppServiceFee(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.SHIPPING_APP_SERVICE_FEE))
                .shippingProtectionServiceFee(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.SHIPPING_PROTECTION_SERVICE_FEE))
                .shippingProtectionReimbursement(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.SHIPPING_PROTECTION_REIMBURSEMENT))
                .fbtOverallMerchantSubsidy(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.FBT_OVERALL_MERCHANT_SUBSIDY))
                .fbtKeyMerchantSubsidy(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.FBT_KEY_MERCHANT_SUBSIDY))
                .fees(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.FEES))
                .transactionFee(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.TRANSACTION_FEE))
                .referralFee(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.REFERRAL_FEE))
                .refundAdministrationFee(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.REFUND_ADMINISTRATION_FEE))
                .affiliateCommission(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.AFFILIATE_COMMISSION))
                .affiliatePartnerCommission(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.AFFILIATE_PARTNER_COMMISSION))
                .affiliateShopAdsCommission(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.AFFILIATE_SHOP_ADS_COMMISSION))
                .tiktokShopPartnerCommission(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.TIKTOK_SHOP_PARTNER_COMMISSION))
                .coFundedPromotionSellerFunded(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.CO_FUNDED_PROMOTION_SELLER_FUNDED))
                .smartPromotionFeeTax(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.SMART_PROMOTION_FEE_TAX))
                .coFundedCreatorBonus(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.CO_FUNDED_CREATOR_BONUS))
                .affiliateCommissionDeposit(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.AFFILIATE_COMMISSION_DEPOSIT))
                .affiliateCommissionRefund(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.AFFILIATE_COMMISSION_REFUND))
                .affiliatePartnerShopAdsCommission(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.AFFILIATE_PARTNER_SHOP_ADS_COMMISSION))
                .campaignServiceFee(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.CAMPAIGN_SERVICE_FEE))
                .campaignResourceFee(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.CAMPAIGN_RESOURCE_FEE))
                .salesTaxOnReferralFees(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.SALES_TAX_ON_REFERRAL_FEES))
                .smartPromotionFee(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.SMART_PROMOTION_FEE))
                .managedServicePlanSalesTax(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.MANAGED_SERVICE_PLAN_SALES_TAX))
                .managedServicePlanPerOrderFee(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.MANAGED_SERVICE_PLAN_PER_ORDER_FEE))
                .coFundedPromotionCampaignPeriodFee(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.CO_FUNDED_PROMOTION_CAMPAIGN_PERIOD_FEE))
                .smartPromotionCampaignPeriodFee(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.SMART_PROMOTION_CAMPAIGN_PERIOD_FEE))
                .coFundedPromotionCampaignPeriodFeeSalesTax(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.CO_FUNDED_PROMOTION_CAMPAIGN_PERIOD_FEE_SALES_TAX))
                .smartPromotionCampaignPeriodFeeSalesTax(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.SMART_PROMOTION_CAMPAIGN_PERIOD_FEE_SALES_TAX))
                .gmvMaxCoupon(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.GMV_MAX_COUPON))
                .gmvMaxCouponSalesTax(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.GMV_MAX_COUPON_SALES_TAX))
                .gmvMaxAdFee(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.GMV_MAX_AD_FEE))
                .categoryLedCampaignFee(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.CATEGORY_LED_CAMPAIGN_FEE))
                .categoryLedCampaignFeeTax(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.CATEGORY_LED_CAMPAIGN_FEE_TAX))
                .brandAmplificationProgramCommissionFee(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.BRAND_AMPLIFICATION_PROGRAM_COMMISSION_FEE))
                .brandAmplificationProgramFeeTax(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.BRAND_AMPLIFICATION_PROGRAM_FEE_TAX))
                .brandCampaignFee(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.BRAND_CAMPAIGN_FEE))
                .brandCampaignFeeTax(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.BRAND_CAMPAIGN_FEE_TAX))
                .gmvMaxAdFeeSalesTax(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.GMV_MAX_AD_FEE_SALES_TAX))
                .shippingProtectionFeeTax(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.SHIPPING_PROTECTION_FEE_TAX))
                .adjustmentAmount(ExcelUtils.getDoubleSafe(row, headerMap, TikTokColumn.ADJUSTMENT_AMOUNT))
                .adjustmentReason(ExcelUtils.getStringSafe(row, headerMap, TikTokColumn.ADJUSTMENT_REASON))
                .relatedOrderId(ExcelUtils.getStringSafe(row, headerMap, TikTokColumn.RELATED_ORDER_ID))
                .build();
    }

    private TikTokSuspense buildSuspenseRow(TikTokRawTransaction auditRow, String reason) {
        TikTokSuspense suspenseRow = new TikTokSuspense();
        BeanUtils.copyProperties(auditRow, suspenseRow);
        suspenseRow.setErrorReason(reason);
        return suspenseRow;
    }
}
