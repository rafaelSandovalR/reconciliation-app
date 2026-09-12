package com.apex.reconciliation_app.model;

import jakarta.persistence.*;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.time.LocalDateTime;

@Entity
@Table(name = "ebay_suspense_queue")
@Data
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class EbaySuspense implements EbayTransactionData{
    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;

    private String errorReason;

    @Builder.Default
    private LocalDateTime importTimeStamp = LocalDateTime.now();

    // Identifiers
    private String type;
    private String orderNumber;
    private String legacyOrderId;
    private String buyerUsername;
    private String buyerName;
    private String itemId;
    private String transactionId;

    // Timestamps
    private LocalDateTime transactionCreationDate;
    private LocalDateTime payoutDate;

    // Transaction Details
    private String payoutCurrency;
    private String payoutId;
    private String payoutMethod;
    private String payoutStatus;
    private String reasonForHold;
    private String itemTitle;
    private String customLabel;
    private String transactionCurrency;
    private String referenceId;
    private String description;

    private Double quantity;
    private Double exchangeRate;

    // Financial Metrics
    private Double netAmount;
    private Double itemSubtotal;
    private Double shippingAndHandling;
    private Double sellerCollectedTax;
    private Double ebayCollectedTax;
    private Double finalValueFeeFixed;
    private Double finalValueFeeVariable;
    private Double regulatoryOperatingFee;
    private Double veryHighItemNotAsDescribedFee;
    private Double belowStandardPerformanceFee;
    private Double internationalFee;
    private Double charityDonation;
    private Double depositProcessingFee;
    private Double grossTransactionAmount;

    // Shipping Info
    private String shipToCity;
    private String shipToProvinceRegionState;
    private String shipToZip;
    private String shipToCountry;
}
