package com.apex.reconciliation_app.enums;

import com.apex.reconciliation_app.model.ReconciliationRecord;
import com.apex.reconciliation_app.util.ExcelUtils;
import lombok.Getter;
import org.apache.poi.ss.usermodel.Cell;

import java.util.function.BiConsumer;

@Getter
public enum RithumColumn {
    SITE_NAME("Site Name", (record, cell) -> record.setSiteName(ExcelUtils.getStringValue(cell))),
    SKU("SKU", (record, cell) -> record.setSku(ExcelUtils.getStringValue(cell))),
    ORDER_DATE("Order Date", (record, cell) -> record.setOrderDate(ExcelUtils.getStringValue(cell))),
    ACCOUNT("Account", (record, cell) -> record.setAccount(ExcelUtils.getStringValue(cell))),
    SITE_ORDER_ID("Site Order ID", (record, cell) -> record.setSiteOrderId(ExcelUtils.getStringValue(cell))),
    SALESPERSON("Salesperson", (record, cell) -> record.setSalesperson(ExcelUtils.getStringValue(cell))),

    TOTAL_LESS_TAX("totallesstax", (record, cell) -> record.setTotalLessTax(ExcelUtils.getDoubleValue(cell))),
    TOTAL_SELLER_COST("totalsellercost", (record, cell) -> record.setTotalSellerCost(ExcelUtils.getDoubleValue(cell))),
    SITE_FEES("SiteFees", (record, cell) -> record.setSiteFees(ExcelUtils.getDoubleValue(cell))),
    PAYPAL_FEES("PaypalFees", (record, cell) -> record.setPaypalFees(ExcelUtils.getDoubleValue(cell))),
    CA_FEES("CAFees", (record, cell) -> record.setCaFees(ExcelUtils.getDoubleValue(cell))),
    PICK_PACK_FEES("PickPackFees", (record, cell) -> record.setPickPackFees(ExcelUtils.getDoubleValue(cell))),
    SHIPPING_COSTS_EST("ShippingCostEST", (record, cell) -> record.setShippingCostsEst(ExcelUtils.getDoubleValue(cell))),
    PROFIT("profit", (record, cell) -> record.setRithumProfit(ExcelUtils.getDoubleValue(cell)));


    private final String headerName;
    private final BiConsumer<ReconciliationRecord, Cell> setter;

    RithumColumn(String headerName, BiConsumer<ReconciliationRecord, Cell> setter) {
        this.headerName = headerName;
        this.setter = setter;
    }

    public void applyTo(ReconciliationRecord record, Cell cell) {
        if (setter != null) {
            setter.accept(record, cell);
        }
    }

}
