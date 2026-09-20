package com.apex.reconciliation_app.enums;

import lombok.Getter;

@Getter
public enum TemuColumn implements ExcelColumn{
    DATE_TIME("Date/time"),
    TRANSACTION_TYPE("Transaction type"),
    RELATED_ID("Related ID"),
    ORDER_ID("Order ID"),
    ORDER_ITEM_ID("Order item ID"),
    SKU("SKU"),
    SKU_ID("SKU ID"),
    QUANTITY("Quantity"),
    SHIP_CITY("Ship city"),
    SHIP_STATE("Ship state"),
    RETAIL_PRICE("Retail price"),
    PLATFORM_DISCOUNT("Platform discount"),
    SELLER_DISCOUNT("Seller discount"),
    SERVICE_FEE("Service fee"),
    SERVICE_FEE_TAX("Service fee tax"),
    PLATFORM_INCENTIVE("Platform incentive"),
    SUBTOTAL("Subtotal"),
    SHIPPING("Shipping"),
    PLATFORM_INCENTIVE_SHIPPING("Platform incentive - Shipping"),
    PRODUCT_TAX("Product Tax"),
    SHIPPING_TAX("Shipping Tax"),
    SIGN_ON_DELIVERY("Sign on delivery"),
    SIGN_ON_DELIVERY_TAX("Sign on delivery tax"),
    MARKETPLACE_WITHHELD_TAX("Marketplace Withheld Tax"),
    OTHERS("Others"),
    TOTAL("Total"),
    CURRENCY("Currency");

    private final String headerName;

    TemuColumn(String headerName) {this.headerName = headerName; }

}
