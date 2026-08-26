package com.apex.reconciliation_app.enums;


import lombok.Getter;

@Getter
public enum WalmartColumn {
    TRANSACTION_TYPE("Transaction Type"),
    TRANSACTION_DESC("Transaction Description"),
    PURCHASE_ORDER("Purchase Order #"),
    AMOUNT("Amount"),
    AMOUNT_TYPE("Amount Type"),
    SKU("Partner Item Id");

    private final String headerName;

    WalmartColumn(String headerName) {
        this.headerName = headerName;
    }

}
