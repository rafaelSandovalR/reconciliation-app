package com.apex.reconciliation_app.enums;

public enum WalmartColumn {
    TRANSACTION_DESC(3),    // Column D
    PURCHASE_ORDER(6),      // Column G
    AMOUNT(8),              // Column I
    AMOUNT_TYPE(9),         // Column J
    SKU(14);                // Column O

    private final int index;

    WalmartColumn(int index) {
        this.index = index;
    }

    public int getIndex() {
        return index;
    }
}
