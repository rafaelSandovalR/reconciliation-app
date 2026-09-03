package com.apex.reconciliation_app.model;

import jakarta.persistence.*;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.time.LocalDateTime;

@Entity
@Table(name = "amazon_raw_transactions")
@Data
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class AmazonRawTransaction {
    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;

    @Builder.Default
    private LocalDateTime importTimeStamp = LocalDateTime.now();

    @Column(unique = true)
    private String compositeTransactionId;

    // IDENTIFIERS
    private String settlementId;
    private String orderId;
    private String sku;


    // TIMESTAMPS
    private LocalDateTime dateTime;
    private LocalDateTime transactionReleaseDate;

    // TRANSACTION DETAILS
    private String type;
    private String description;
    private String marketplace;
    private String accountType;
    private String fulfillment;
    private String taxCollectionModel;
    private String transactionStatus;

    private Double quantity;

    // FINANCIAL METRICS
    private Double productSales;
    private Double productSalesTax;
    private Double shippingCredits;
    private Double shippingCreditsTax;
    private Double giftWrapCredits;
    private Double giftWrapCreditTax;
    private Double RegulatoryFee;
    private Double taxOnRegulatoryFee;
    private Double promotionalRebates;
    private Double promotionalRebatestax;
    private Double marketplaceWithheldTax;
    private Double sellingFees;
    private Double fbaFees;
    private Double otherTransactionFees;
    private Double other;
    private Double total;

    // SHIPPING INFO
    private String orderCity;
    private String orderState;
    private String orderPostal;
}
