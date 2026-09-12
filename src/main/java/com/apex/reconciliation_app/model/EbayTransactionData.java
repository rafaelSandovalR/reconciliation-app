package com.apex.reconciliation_app.model;

import java.time.LocalDateTime;

public interface EbayTransactionData {

    LocalDateTime getTransactionCreationDate();
    String getType();
    String getOrderNumber();
    String getLegacyOrderId();
    String getBuyerUsername();
    String getBuyerName();
    String getShipToCity();
    String getShipToProvinceRegionState();
    String getShipToZip();
    String getShipToCountry();
    Double getNetAmount();
    String getPayoutCurrency();
    LocalDateTime getPayoutDate();
    String getPayoutId();
    String getPayoutMethod();
    String getPayoutStatus();
    String getReasonForHold();
    String getItemId();
    String getTransactionId();
    String getItemTitle();
    String getCustomLabel();
    Double getQuantity();
    Double getItemSubtotal();
    Double getShippingAndHandling();
    Double getSellerCollectedTax();
    Double getEbayCollectedTax();
    Double getFinalValueFeeFixed();
    Double getFinalValueFeeVariable();
    Double getRegulatoryOperatingFee();
    Double getVeryHighItemNotAsDescribedFee();
    Double getBelowStandardPerformanceFee();
    Double getInternationalFee();
    Double getCharityDonation();
    Double getDepositProcessingFee();
    Double getGrossTransactionAmount();
    String getTransactionCurrency();
    Double getExchangeRate();
    String getReferenceId();
    String getDescription();


}
