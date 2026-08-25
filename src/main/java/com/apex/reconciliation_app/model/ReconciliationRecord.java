package com.apex.reconciliation_app.model;

import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import lombok.Data;


@Entity
@Data
public class ReconciliationRecord {

    @Id
    private String compositeId;

    // Rithum Base Data
    private String siteName;
    private String sku;
    private String siteOrderId;
    private String orderDate;
    private String account;
    private String salesperson;

    // Rithum Financials
    private Double totalLessTax;
    private Double totalSellerCost;
    private Double siteFees;
    private Double paypalFees;
    private Double caFees;
    private Double pickPackFees;
    private Double shippingCostsEst;
    private Double rithumProfit;

    // Marketplace Data (Initially Null)
    private Double siteOrderAmount;
    private Double siteOrderFee;
    private Double siteOrderOtherFees1;
    private Double siteOrderOtherFees2;
    private Double siteOrderOtherFees3;
    private Double siteOrderOtherFees4;

    private Double actualShippingCosts;
    private Double shippingAdjustments;

    private Double actualProfit;
    private Double profitDiff;

    private String returnStatus;
    private Double amountRefunded;
    private Double returnShipping;
    private Double returnFee1;
    private Double returnFee2;
    private Double returnFee3;
    private Double returnFee4;

    private String notes;
}
