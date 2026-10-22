package com.apex.reconciliation_app.model;

import com.apex.reconciliation_app.exception.BucketOverflowException;
import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import lombok.Data;

import java.time.LocalDateTime;


@Entity
@Data
public class ReconciliationRecord {

    @Id
    private String compositeId;

    private String secondarySiteOrderId;
    private String siteName;
    private String buyer;
    private String sku;
    private String title;
    private LocalDateTime orderDate;
    private LocalDateTime shipmentStatusDate;
    private String paymentType;
    private Double quantity;
    private Double orderTotal;
    private String account;
    private String siteListingId;
    private String channelAdvisorOrderId;
    private String buyerCreditCard;

    @Column(length = 4000)
    private String buyerFeedbackMessage;

    private String checkoutStatus;
    private String currency;
    private LocalDateTime deliverByDate;
    private String disputeStatus;
    private LocalDateTime estimatedShipDate;
    private String fulfillment;
    private String itemASIN;
    private String itemClassification;
    private String itemEAN;
    private String itemHarmonizedCode;
    private String itemISBN;
    private String itemMPN;
    private Double itemPromotion;
    private String itemPromotionCode;
    private String itemSecondChanceOffer;
    private Double itemSellerCost;
    private Double itemShippingPrice;
    private Double itemShippingPromotion;
    private Double itemShippingTax;
    private Double itemTax;
    private Double itemTotalPromotion;
    private String itemUPC;
    private Double itemWeight;
    private String merchantReferenceNumber;
    private LocalDateTime orderImportDate;
    private String orderTags;
    private String paymentStatus;
    private LocalDateTime paymentStatusDate;

    @Column(length = 4000)
    private String privateNotes;

    @Column(length = 4000)
    private String publicNotes;

    private String refundStatus;
    private String requestedShippingCarrierName;
    private String requestedShippingClassName;
    private String sellerLeftFeedback;
    private String shippingStatus;
    private String specialInstructions;
    private Double totalShippingPrice;
    private Double totalShippingTaxPrice;
    private Double totalTaxPrice;
    private Double unitPrice;
    private String warehouseLocation;
    private String buyerUserId;
    private String paymentPaypalAccountId;
    private String paymentTransactionId;
    private String sellerOrderId;
    private String siteOrderId;
    private String siteOrderItemId;
    private String sitePurchaseOrderId;
    private String shippingCompanyName;
    private String shippingFirstName;
    private String shippingLastName;
    private String shippingAddressLine1;
    private String shippingAddressLine2;
    private String shippingCity;
    private String shippingStateorProvince;
    private String shippingPostalCode;
    private String shippingCountry;
    private String shippingDayPhone;
    private String shippingEveningPhone;
    private Double orderPromotion;
    private String orderPromotionCode;
    private Double totalLessTax;
    private Double totalSellerCost;
    private Double siteFees;
    private Double paypalFees;
    private Double caFees;
    private Double pickPackFees;
    private Double shippingCostEst;
    private Double profit;
    private String orderNumber;
    private String salesperson;
    private String invoiceNumber;

    // Calculated Fields
    private Double siteOrderAmount;
    private Double siteOrderFee;
    private Double siteOrderOtherFees1;
    private Double siteOrderOtherFees2;
    private Double siteOrderOtherFees3;
    private Double siteOrderOtherFees4;
    private Double actualShippingCost;
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
    private Double commissionRefund; // Hidden Variable



    public void addDynamicRegularFee(double amount, String amountType) {
        if (amount == 0) return;
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
        if (amount == 0) return;
        if (this.returnFee1 == null || this.returnFee1 == 0) {
            this.returnFee1 = amount;
        } else if (this.returnFee2 == null || this.returnFee2 == 0) {
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

    public void calculateCommissionRefundDelta() {

        if (this.getSiteOrderFee() == null || this.getSiteOrderFee() == 0 ||
            this.getCommissionRefund() == null || this.getCommissionRefund() == 0) {
            return;
        }

        double originalCommission = this.getSiteOrderFee();
        double rawDelta = originalCommission + this.getCommissionRefund();
        double delta = Math.round(rawDelta * 100.0) / 100.0;

        if (delta != 0){
            addDynamicReturnFee(delta, "COMMISSION DELTA");
        }
    }
}
