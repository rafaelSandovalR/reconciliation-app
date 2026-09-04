package com.apex.reconciliation_app.model;

import java.time.LocalDateTime;

public interface AmazonTransactionData {

    LocalDateTime getDateTime();
    String getSettlementId();
    String getType();
    String getOrderId();
    String getSku();
    String getDescription();
    Double getQuantity();
    String getMarketplace();
    String getAccountType();
    String getFulfillment();
    String getOrderCity();
    String getOrderState();
    String getOrderPostal();
    String getTaxCollectionModel();
    Double getProductSales();
    Double getProductSalesTax();
    Double getShippingCredits();
    Double getShippingCreditsTax();
    Double getGiftWrapCredits();
    Double getGiftWrapCreditsTax();

    // Note: Adjusting capitalization to match standard Lombok getter generation based on your builder
    Double getRegulatoryFee();
    Double getTaxOnRegulatoryFee();
    Double getPromotionalRebates();
    Double getPromotionalRebatesTax();
    Double getMarketplaceWithheldTax();

    Double getSellingFees();
    Double getFbaFees();
    Double getOtherTransactionFees();
    Double getOther();
    Double getTotal();
    String getTransactionStatus();
    LocalDateTime getTransactionReleaseDate();
}