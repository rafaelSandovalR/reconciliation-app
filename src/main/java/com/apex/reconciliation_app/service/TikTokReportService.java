package com.apex.reconciliation_app.service;

import com.apex.reconciliation_app.enums.TikTokColumn;
import com.apex.reconciliation_app.model.tiktok.TikTokRawTransaction;
import com.apex.reconciliation_app.model.tiktok.TikTokSuspense;
import com.apex.reconciliation_app.model.tiktok.TikTokTransactionData;
import org.apache.poi.ss.usermodel.Row;
import org.springframework.stereotype.Service;

import java.time.format.DateTimeFormatter;

@Service
public class TikTokReportService extends AbstractReportService<TikTokSuspense, TikTokRawTransaction, TikTokColumn> {

    private static final DateTimeFormatter TIKTOK_FORMATTER = DateTimeFormatter.ofPattern("yyyy/MM/dd");

    @Override
    protected Class<TikTokColumn> getColumnEnumClass() {
        return TikTokColumn.class;
    }

    @Override
    protected void writeSuspenseData(Row row, TikTokSuspense record) {
        int col = 0;
        setCellValue(row.createCell(col++), record.getErrorReason());
        setCellValue(row.createCell(col++), record.getCompositeId());
        setCellValue(row.createCell(col++), record.getCompositeTransactionId());
        writeCommonData(row, record, col);
    }

    @Override
    protected void writeAuditData(Row row, TikTokRawTransaction record) {
        int col = 0;
        setCellValue(row.createCell(col++), record.getCompositeId());
        setCellValue(row.createCell(col++), record.getCompositeTransactionId());
        writeCommonData(row, record, col);
    }

    private void writeCommonData(Row row, TikTokTransactionData record, int col) {
        setCellValue(row.createCell(col++), record.getStatementDate() != null ? record.getStatementDate().format(TIKTOK_FORMATTER) : null);
        setCellValue(row.createCell(col++), record.getStatementId());
        setCellValue(row.createCell(col++), record.getPaymentId());
        setCellValue(row.createCell(col++), record.getStatus());
        setCellValue(row.createCell(col++), record.getCurrency());
        setCellValue(row.createCell(col++), record.getType());
        setCellValue(row.createCell(col++), record.getOrderAdjustmentId());
        setCellValue(row.createCell(col++), record.getSkuId());
        setCellValue(row.createCell(col++), record.getQuantity());
        setCellValue(row.createCell(col++), record.getProductName());
        setCellValue(row.createCell(col++), record.getSkuName());
        setCellValue(row.createCell(col++), record.getOrderCreatedDate()!= null ? record.getOrderCreatedDate().format(TIKTOK_FORMATTER) : null);
        setCellValue(row.createCell(col++), record.getOrderShipmentDate()!= null ? record.getOrderShipmentDate().format(TIKTOK_FORMATTER) : null);
        setCellValue(row.createCell(col++), record.getOrderDeliveryDate()!= null ? record.getOrderDeliveryDate().format(TIKTOK_FORMATTER) : null);
        setCellValue(row.createCell(col++), record.getTotalSettlementAmount());
        setCellValue(row.createCell(col++), record.getNetSales());
        setCellValue(row.createCell(col++), record.getGrossSales());
        setCellValue(row.createCell(col++), record.getGrossSalesRefund());
        setCellValue(row.createCell(col++), record.getSellerDiscount());
        setCellValue(row.createCell(col++), record.getSellerDiscountRefund());
        setCellValue(row.createCell(col++), record.getShipping());
        setCellValue(row.createCell(col++), record.getTiktokShopShippingFee());
        setCellValue(row.createCell(col++), record.getFulfilledByTiktokShopShippingFee());
        setCellValue(row.createCell(col++), record.getSignatureConfirmationServiceFee());
        setCellValue(row.createCell(col++), record.getShippingInsuranceFee());
        setCellValue(row.createCell(col++), record.getCustomerPaidShippingFee());
        setCellValue(row.createCell(col++), record.getCustomerPaidShippingFeeRefund());
        setCellValue(row.createCell(col++), record.getTiktokShopShippingIncentive());
        setCellValue(row.createCell(col++), record.getTiktokShopShippingIncentiveRefund());
        setCellValue(row.createCell(col++), record.getShippingFeeSubsidy());
        setCellValue(row.createCell(col++), record.getReturnShippingFee());
        setCellValue(row.createCell(col++), record.getFbtFulfillmentFee());
        setCellValue(row.createCell(col++), record.getCustomerShippingFeeOffset());
        setCellValue(row.createCell(col++), record.getShippingFeeDiscount());
        setCellValue(row.createCell(col++), record.getReturnShippingLabelFee());
        setCellValue(row.createCell(col++), record.getFbtFulfillmentFeeReimbursement());
        setCellValue(row.createCell(col++), record.getReturnShippingFeePaidByCustomers());
        setCellValue(row.createCell(col++), record.getReturnShippingFeeReimbursement());
        setCellValue(row.createCell(col++), record.getShippingAppServiceFee());
        setCellValue(row.createCell(col++), record.getShippingProtectionServiceFee());
        setCellValue(row.createCell(col++), record.getShippingProtectionReimbursement());
        setCellValue(row.createCell(col++), record.getFbtOverallMerchantSubsidy());
        setCellValue(row.createCell(col++), record.getFbtKeyMerchantSubsidy());
        setCellValue(row.createCell(col++), record.getFees());
        setCellValue(row.createCell(col++), record.getTransactionFee());
        setCellValue(row.createCell(col++), record.getReferralFee());
        setCellValue(row.createCell(col++), record.getRefundAdministrationFee());
        setCellValue(row.createCell(col++), record.getAffiliateCommission());
        setCellValue(row.createCell(col++), record.getAffiliatePartnerCommission());
        setCellValue(row.createCell(col++), record.getAffiliateShopAdsCommission());
        setCellValue(row.createCell(col++), record.getTiktokShopPartnerCommission());
        setCellValue(row.createCell(col++), record.getCoFundedPromotionSellerFunded());
        setCellValue(row.createCell(col++), record.getSmartPromotionFeeTax());
        setCellValue(row.createCell(col++), record.getCoFundedCreatorBonus());
        setCellValue(row.createCell(col++), record.getAffiliateCommissionDeposit());
        setCellValue(row.createCell(col++), record.getAffiliateCommissionRefund());
        setCellValue(row.createCell(col++), record.getAffiliatePartnerShopAdsCommission());
        setCellValue(row.createCell(col++), record.getCampaignServiceFee());
        setCellValue(row.createCell(col++), record.getCampaignResourceFee());
        setCellValue(row.createCell(col++), record.getSalesTaxOnReferralFees());
        setCellValue(row.createCell(col++), record.getSmartPromotionFee());
        setCellValue(row.createCell(col++), record.getManagedServicePlanSalesTax());
        setCellValue(row.createCell(col++), record.getManagedServicePlanPerOrderFee());
        setCellValue(row.createCell(col++), record.getCoFundedPromotionCampaignPeriodFee());
        setCellValue(row.createCell(col++), record.getSmartPromotionCampaignPeriodFee());
        setCellValue(row.createCell(col++), record.getCoFundedPromotionCampaignPeriodFeeSalesTax());
        setCellValue(row.createCell(col++), record.getSmartPromotionCampaignPeriodFeeSalesTax());
        setCellValue(row.createCell(col++), record.getGmvMaxCoupon());
        setCellValue(row.createCell(col++), record.getGmvMaxCouponSalesTax());
        setCellValue(row.createCell(col++), record.getGmvMaxAdFee());
        setCellValue(row.createCell(col++), record.getCategoryLedCampaignFee());
        setCellValue(row.createCell(col++), record.getCategoryLedCampaignFeeTax());
        setCellValue(row.createCell(col++), record.getBrandAmplificationProgramCommissionFee());
        setCellValue(row.createCell(col++), record.getBrandAmplificationProgramFeeTax());
        setCellValue(row.createCell(col++), record.getBrandCampaignFee());
        setCellValue(row.createCell(col++), record.getBrandCampaignFeeTax());
        setCellValue(row.createCell(col++), record.getGmvMaxAdFeeSalesTax());
        setCellValue(row.createCell(col++), record.getShippingProtectionFeeTax());
        setCellValue(row.createCell(col++), record.getAdjustmentAmount());
        setCellValue(row.createCell(col++), record.getAdjustmentReason());
        setCellValue(row.createCell(col++), record.getRelatedOrderId());

    }
}
