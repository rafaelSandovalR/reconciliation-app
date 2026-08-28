package com.apex.reconciliation_app.model;


import jakarta.persistence.*;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.time.LocalDateTime;

@Entity
@Table(name = "walmart_raw_transactions")
@Data
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class WalmartRawTransaction {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;

    // Automatically records when row was ingested
    @Builder.Default
    private LocalDateTime importTimeStamp = LocalDateTime.now();

    // IDENTIFIERS
    private String transactionKey;
    private String purchaseOrder;
    private String purchaseOrderLine;
    private String customerOrder;
    private String customerOrderLine;
    private String partnerGtIn;
    private String partnerItemId; // SKU
    private String campaignId;

    // TIMESTAMPS
    private LocalDateTime transactionPostedTimestamp;

    // TRANSACTION DETAILS
    private String transactionType;
    private String transactionDesc;
    private String amountType;
    private String transactionReasonDesc;
    private String partnerItemName;
    private String productTaxCode;
    private String contractCategory;
    private String productType;
    private String commissionRule;
    private String shippingMethod;
    private String fulfillmentType;
    private String fulfillmentDetails;
    private String customerPromoType;
    private String incentiveProgramName;
    private String itemCondition;

    // FINANCIAL METRICS
    private Double amount;
    private Double shipQty;
    private Double commissionRate;
    private Double baseCommissionRate;
    private Double originalCommission;
    private Double commissionIncentiveProgram;
    private Double commissionSaving;
    private Double totalWalmartFundedSavings;
    private Double originalCharge;
    private Double chargeSavings;

    // SHIPPING INFO
    private String shipToState;
    private String shipToCity;
    private String shipToZipcode; // String to preserve leading zeros in zip codes
    private String shipToCountry;
}
