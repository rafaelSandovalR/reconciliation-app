package com.apex.reconciliation_app.enums;

import lombok.Getter;

@Getter
public enum RithumColumn {
    SITE_NAME(1),
    SKU(3),
    ORDER_DATE(5),
    ACCOUNT(10),
    SITE_ORDER_ID(60),
    TOTAL_LESS_TAX(76),
    TOTAL_SELLER_COST(77),
    SITE_FEES(78),
    PAYPAL_FEES(79),
    CA_FEES(80),
    PICK_PACK_FEES(81),
    SHIPPING_COSTS_EST(82),
    PROFIT(83),
    SALESPERSON(85);

    private final int index;

    RithumColumn(int index) {
        this.index = index;
    }

}
