package com.apex.reconciliation_app.enums;


import lombok.Getter;

@Getter
public enum WalmartColumn {
    TRANSACTION_KEY("Transaction Key"),
    TRANSACTION_POSTED_TIMESTAMP("Transaction Posted Timestamp"),
    TRANSACTION_TYPE("Transaction Type"),
    TRANSACTION_DESC("Transaction Description"),
    CUSTOMER_ORDER("Customer Order #"),
    CUSTOMER_ORDER_LINE("Customer Order line #"),
    PURCHASE_ORDER("Purchase Order #"),
    PURCHASE_ORDER_LINE("Purchase Order line #"),
    AMOUNT("Amount"),
    AMOUNT_TYPE("Amount Type"),
    SHIP_QTY("Ship Qty"),
    COMMISSION_RATE("Commission Rate"),
    BASE_COMMISSION_RATE("Base Commission Rate"),
    TRANSACTION_REASON_DESC("Transaction Reason Description"),
    SKU("Partner Item Id"),
    PARTNER_GTIN("Partner GTIN"),
    PARTNER_ITEM_NAME("Partner Item Name"),
    PRODUCT_TAX_CODE("Product Tax Code"),
    SHIP_TO_STATE("Ship to State"),
    SHIP_TO_CITY("Ship to City"),
    SHIP_TO_ZIPCODE("Ship to Zipcode"),
    CONTRACT_CATEGORY("Contract Category"),
    PRODUCT_TYPE("Product Type"),
    COMMISSION_RULE("Commission Rule"),
    SHIPPING_METHOD("Shipping Method"),
    FULFILLMENT_TYPE("Fulfillment Type"),
    FULFILLMENT_DETAILS("Fulfillment Details"),
    ORIGINAL_COMMISSION("Original Commission"),
    COMMISSION_INCENTIVE_PROGRAM("Commission Incentive Program"),
    COMMISSION_SAVING("Commission Saving"),
    CUSTOMER_PROMO_TYPE("Customer Promo Type"),
    TOTAL_WALMART_FUNDED_SAVINGS("Total Walmart Funded Savings Program"),
    CAMPAIGN_ID("Campaign Id"),
    ITEM_CONDITION("Item Condition"),
    ORIGINAL_CHARGE("Original charge"),
    CHARGE_SAVINGS("Charge Savings"),
    INCENTIVE_PROGRAM_NAME("Incentive Program Name"),
    SHIP_TO_COUNTRY("Ship to country");

    private final String headerName;

    WalmartColumn(String headerName) {
        this.headerName = headerName;
    }

}
