package com.apex.reconciliation_app.enums;


import lombok.Getter;

@Getter
public enum WalmartColumn {
    TRANSACTION_DESC("Transaction Description"),    // Column D
    PURCHASE_ORDER("Purchase Order #"),      // Column G
    AMOUNT("Amount"),              // Column I
    AMOUNT_TYPE("Amount Type"),         // Column J
    SKU("Partner Item Id");                // Column O

    private final String headerName;

    WalmartColumn(String headerName) {
        this.headerName = headerName;
    }

}
