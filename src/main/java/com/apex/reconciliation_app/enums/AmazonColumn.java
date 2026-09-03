package com.apex.reconciliation_app.enums;

import lombok.Getter;

@Getter
public enum AmazonColumn {
    DATE_TIME("date/time"),
    SETTLEMENT_ID("settlement id"),
    TYPE("type"),
    ORDER_ID("order id"),
    SKU("sku"),
    DESCRIPTION("description"),
    QUANTITY("quantity"),
    MARKETPLACE("marketplace"),
    ACCOUNT_TYPE("account type"),
    FULFILLMENT("fulfillment"),
    ORDER_CITY("order city"),
    ORDER_STATE("order state"),
    ORDER_POSTAL("order postal"),
    TAX_COLLECTION_MODEL("tax collection model"),
    PRODUCT_SALES("product sales"),
    PRODUCT_SALES_TAX("product sales tax"),
    SHIPPING_CREDITS("shipping credits"),
    SHIPPING_CREDITS_TAX("shipping credits tax"),
    GIFT_WRAP_CREDITS("gift wrap credits"),
    GIFTWRAP_CREDITS_TAX("giftwrap credits tax"),
    REGULATORY_FEE("Regulatory Fee"),
    TAX_ON_REGULATORY_FEE("Tax On Regulatory Fee"),
    PROMOTIONAL_REBATES("promotional rebates"),
    PROMOTIONAL_REBATES_TAX("promotional rebates tax"),
    MARKETPLACE_WITHHELD_TAX("marketplace withheld tax"),
    SELLING_FEES("selling fees"),
    FBA_FEES("fba fees"),
    OTHER_TRANSACTION_FEES("other transaction fees"),
    OTHER("other"),
    TOTAL("total"),
    TRANSACTION_STATUS("Transaction Status"),
    TRANSACTION_RELEASE_DATE("Transaction Release Date");

    private final String headerName;

    AmazonColumn(String headerName) { this.headerName = headerName; }
}
