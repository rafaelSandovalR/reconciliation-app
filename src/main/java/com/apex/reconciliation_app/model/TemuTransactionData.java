package com.apex.reconciliation_app.model;

import java.time.LocalDateTime;

public interface TemuTransactionData {
    LocalDateTime getDateTime();
    String getTransactionType();
    String getRelatedId();
    String getOrderId();
    String getOrderItemId();
    String getSku();
    String getSkuId();
    Double getQuantity();
    String getShipCity();
    String getShipState();
    Double getRetailPrice();
    Double getPlatformDiscount();
    Double getSellerDiscount();
    Double getServiceFee();
    Double getServiceFeeTax();
    Double getPlatformIncentive();
    Double getSubtotal();
    Double getShipping();
    Double getPlatformIncentiveShipping();
    Double getProductTax();
    Double getShippingTax();
    Double getSignOnDelivery();
    Double getSignOnDeliveryTax();
    Double getMarketplaceWithheldTax();
    Double getOthers();
    Double getTotal();
    String getCurrency();
}
