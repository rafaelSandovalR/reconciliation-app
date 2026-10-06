package com.apex.reconciliation_app.enums;

import lombok.Getter;

@Getter
public enum ShopSimonColumn implements ExcelColumn{
    TYPE("type"),
    ENTRY("entry"),
    ORDER_ID("order_id"),
    PARTNER_ID("partner_id"),
    PO_NUMBER("po_number"),
    CONSUMER_ORDER_NUMBER("consumer_order_number"),
    SUPPLIER_ORDER_NUMBER("supplier_order_number"),
    CURRENCY_CODE("currency_code"),
    LINE_NUMBER("line_number"),
    SKU("sku"),
    PARTNER_SKU("partner_sku"),
    ITEM_ID("item_id"),
    TRANSACTION_DATE("transaction_date"),
    QUANTITY_SHIPPED("quantity_shipped"),
    AMOUNT_SHIPPED("amount_shipped"),
    RETURN_ID("return_id"),
    QUANTITY_RETURNED("quantity_returned"),
    LINE_ITEM_FULFILLMENT("line_item_fulfillment"),
    CONSUMER_UNIT_PRICE("consumer_unit_price"),
    CONSUMER_UNIT_PRICE_INCLUDING_TAX("consumer_unit_price_including_tax"),
    TAX_TYPE("tax_type"),
    TAX_PERCENTAGE("tax_percentage"),
    UNIT_TAX_AMOUNT("unit_tax_amount"),
    TOTAL_TAX("total_tax"),
    CONSUMER_UNIT_PRICE_EXCLUDING_TAX("consumer_unit_price_excluding_tax"),
    COMMISSION_PERCENTAGE("commission_percentage"),
    TOTAL_COMMISSION_AMOUNT("total_commission_amount"),
    ADJUSTMENT_REASON("adjustment_reason"),
    TRANSACTION_AMOUNT("transaction_amount"),
    TOTAL_AMOUNT_PAID("total_amount_paid"),
    TOTAL_AMOUNT_OUTSTANDING("total_amount_outstanding"),
    TOTAL_CONSUMER_PRICE("total_consumer_price"),
    TRACKING_NUMBER("tracking_number"),
    INVOICE_ID("invoice_id"),
    SUPPLIER_INVOICE_NUMBER("supplier_invoice_number"),
    INVOICE_RECORD_TYPE("invoice_record_type"),
    SUBORDER_ID("suborder_id"),
    INVOICE_DATE("invoice_date"),
    QUANTITY_INVOICED("quantity_invoiced"),
    ORDER_QUANTITY("order_quantity"),
    UNIT_COST("unit_cost"),
    TOTAL_UNIT_COST("total_unit_cost"),
    FREIGHT_AMOUNT("freight_amount"),
    SALES_TAX_AMOUNT("sales_tax_amount"),
    CHARGES("charges"),
    CREDITS("credits"),
    SUBTOTAL_EXCLUDING_LINE_ITEMS("subtotal_excluding_line_items");

    private final String headerName;

    ShopSimonColumn(String headerName) { this.headerName = headerName; }
}
