package com.apex.reconciliation_app.enums;

import lombok.Getter;

@Getter
public enum EbayColumn implements ExcelColumn {

    TRANSACTION_CREATION_DATE("Transaction creation date"),
    TYPE("Type"),
    ORDER_NUMBER("Order number"),
    LEGACY_ORDER_ID("Legacy order ID"),
    BUYER_USERNAME("Buyer username"),
    BUYER_NAME("Buyer name"),
    SHIP_TO_CITY("Ship to city"),
    SHIP_TO_PROVINCE_REGION_STATE("Ship to province/region/state"),
    SHIP_TO_ZIP("Ship to zip"),
    SHIP_TO_COUNTRY("Ship to country"),
    NET_AMOUNT("Net amount"),
    PAYOUT_CURRENCY("Payout currency"),
    PAYOUT_DATE("Payout date"),
    PAYOUT_ID("Payout ID"),
    PAYOUT_METHOD("Payout method"),
    PAYOUT_STATUS("Payout status"),
    REASON_FOR_HOLD("Reason for hold"),
    ITEM_ID("Item ID"),
    TRANSACTION_ID("Transaction ID"),
    ITEM_TITLE("Item title"),
    CUSTOM_LABEL("Custom label"),
    QUANTITY("Quantity"),
    ITEM_SUBTOTAL("Item subtotal"),
    SHIPPING_AND_HANDLING("Shipping and handling"),
    SELLER_COLLECTED_TAX("Seller collected tax"),
    EBAY_COLLECTED_TAX("eBay collected tax"),
    FINAL_VALUE_FEE_FIXED("Final Value Fee - fixed"),
    FINAL_VALUE_FEE_VARIABLE("Final Value Fee - variable"),
    REGULATORY_OPERATING_FEE("Regulatory operating fee"),
    VERY_HIGH_ITEM_NOT_AS_DESCRIBED_FEE("Very high \"item not as described\" fee"),
    BELOW_STANDARD_PERFORMANCE_FEE("Below standard performance fee"),
    INTERNATIONAL_FEE("International fee"),
    CHARITY_DONATION("Charity donation"),
    DEPOSIT_PROCESSING_FEE("Deposit processing fee"),
    GROSS_TRANSACTION_AMOUNT("Gross transaction amount"),
    TRANSACTION_CURRENCY("Transaction currency"),
    EXCHANGE_RATE("Exchange rate"),
    REFERENCE_ID("Reference ID"),
    DESCRIPTION("Description");


    private final String headerName;

    EbayColumn(String headerName) { this.headerName = headerName; }
}
