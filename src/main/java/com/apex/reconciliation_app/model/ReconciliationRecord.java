package com.apex.reconciliation_app.model;

import com.apex.reconciliation_app.exception.BucketOverflowException;
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

    public void addDynamicRegularFee(double amount, String amountType) {
        if (this.siteOrderOtherFees1 == null || this.siteOrderOtherFees1 == 0) {
            this.siteOrderOtherFees1 = amount;
        } else if (this.siteOrderOtherFees2 == null || this.siteOrderOtherFees2 == 0) {
            this.siteOrderOtherFees2 = amount;
        } else if (this.siteOrderOtherFees3 == null || this.siteOrderOtherFees3 == 0) {
            this.siteOrderOtherFees3 = amount;
        } else if (this.siteOrderOtherFees4 == null || this.siteOrderOtherFees4 == 0) {
            this.siteOrderOtherFees4 = amount;
        } else {
            throw new BucketOverflowException("No Empty SiteOrderOtherFee columns remaining");
        }

        String note = "Fee: " + amountType + ", ";
        this.notes = (this.notes == null ? note : this.notes.concat(note));
    }

    public void addDynamicReturnFee(double amount, String amountType) {
        if (this.returnFee2 == null || this.returnFee2 == 0) {
            this.returnFee2 = amount;
        } else if (this.returnFee3 == null || this.returnFee3 == 0) {
            this.returnFee3 = amount;
        } else if (this.returnFee4 == null || this.returnFee4 == 0) {
            this.returnFee4 = amount;
        } else {
            throw new BucketOverflowException("No Empty ReturnFee columns remaining");
        }

        String note = "Return Fee: " + amountType + ", ";
        this.notes = (this.notes == null ? note : this.notes.concat(note));
    }
}
