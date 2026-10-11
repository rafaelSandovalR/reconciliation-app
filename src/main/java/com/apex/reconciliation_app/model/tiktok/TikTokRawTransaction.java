package com.apex.reconciliation_app.model.tiktok;

import jakarta.persistence.*;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.time.LocalDateTime;

@Entity
@Table(name = "tiktok_raw_transactions")
@Data
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class TikTokRawTransaction implements TikTokTransactionData {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;

    @Builder.Default
    private LocalDateTime importTimeStamp = LocalDateTime.now();

    @Column(unique = true)
    private String compositeTransactionId;

    private String compositeId;

    // Timestamps
    private LocalDateTime statementDate;
    private LocalDateTime orderCreatedDate;
    private LocalDateTime orderShipmentDate;
    private LocalDateTime orderDeliveryDate;

    // Identifiers
    private String statementId;
    private String paymentId;
    private String orderAdjustmentId;
    private String relatedOrderId;

    // Transaction Details
    private String status;
    private String currency;
    private String type;
    private String skuId;
    private Double quantity;
    private String productName;
    private String skuName;
    private String adjustmentReason;

    // Financial Metrics
    private Double totalSettlementAmount;
    private Double netSales;
    private Double grossSales;
    private Double grossSalesRefund;
    private Double sellerDiscount;
    private Double sellerDiscountRefund;
    private Double shipping;
    private Double tiktokShopShippingFee;
    private Double fulfilledByTiktokShopShippingFee;
    private Double signatureConfirmationServiceFee;
    private Double shippingInsuranceFee;
    private Double customerPaidShippingFee;
    private Double customerPaidShippingFeeRefund;
    private Double tiktokShopShippingIncentive;
    private Double tiktokShopShippingIncentiveRefund;
    private Double shippingFeeSubsidy;
    private Double returnShippingFee;
    private Double fbtFulfillmentFee;
    private Double customerShippingFeeOffset;
    private Double shippingFeeDiscount;
    private Double returnShippingLabelFee;
    private Double fbtFulfillmentFeeReimbursement;
    private Double returnShippingFeePaidByCustomers;
    private Double returnShippingFeeReimbursement;
    private Double shippingAppServiceFee;
    private Double shippingProtectionServiceFee;
    private Double shippingProtectionReimbursement;
    private Double fbtOverallMerchantSubsidy;
    private Double fbtKeyMerchantSubsidy;
    private Double fees;
    private Double transactionFee;
    private Double referralFee;
    private Double refundAdministrationFee;
    private Double affiliateCommission;
    private Double affiliatePartnerCommission;
    private Double affiliateShopAdsCommission;
    private Double tiktokShopPartnerCommission;
    private Double coFundedPromotionSellerFunded;
    private Double smartPromotionFeeTax;
    private Double coFundedCreatorBonus;
    private Double affiliateCommissionDeposit;
    private Double affiliateCommissionRefund;
    private Double affiliatePartnerShopAdsCommission;
    private Double campaignServiceFee;
    private Double campaignResourceFee;
    private Double salesTaxOnReferralFees;
    private Double smartPromotionFee;
    private Double managedServicePlanSalesTax;
    private Double managedServicePlanPerOrderFee;
    private Double coFundedPromotionCampaignPeriodFee;
    private Double smartPromotionCampaignPeriodFee;
    private Double coFundedPromotionCampaignPeriodFeeSalesTax;
    private Double smartPromotionCampaignPeriodFeeSalesTax;
    private Double gmvMaxCoupon;
    private Double gmvMaxCouponSalesTax;
    private Double gmvMaxAdFee;
    private Double categoryLedCampaignFee;
    private Double categoryLedCampaignFeeTax;
    private Double brandAmplificationProgramCommissionFee;
    private Double brandAmplificationProgramFeeTax;
    private Double brandCampaignFee;
    private Double brandCampaignFeeTax;
    private Double gmvMaxAdFeeSalesTax;
    private Double shippingProtectionFeeTax;
    private Double adjustmentAmount;
}
