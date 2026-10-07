package com.apex.reconciliation_app.model.walmart;

import java.time.LocalDateTime;

public interface WalmartTransactionData {
    String getTransactionKey();
    LocalDateTime getTransactionPostedTimestamp();
    String getTransactionType();
    String getTransactionDesc();
    String getCustomerOrder();
    String getCustomerOrderLine();
    String getPurchaseOrder();
    String getPurchaseOrderLine();
    Double getAmount();
    String getAmountType();
    Double getShipQty();
    Double getCommissionRate();
    Double getBaseCommissionRate();
    String getTransactionReasonDesc();
    String getPartnerItemId();
    String getPartnerGtIn();
    String getPartnerItemName();
    String getProductTaxCode();
    String getShipToState();
    String getShipToCity();
    String getShipToZipcode();
    String getContractCategory();
    String getProductType();
    String getCommissionRule();
    String getShippingMethod();
    String getFulfillmentType();
    String getFulfillmentDetails();
    Double getOriginalCommission();
    Double getCommissionIncentiveProgram();
    Double getCommissionSaving();
    String getCustomerPromoType();
    Double getTotalWalmartFundedSavings();
    String getCampaignId();
    String getItemCondition();
    Double getOriginalCharge();
    Double getChargeSavings();
    String getIncentiveProgramName();
    String getShipToCountry();
}
