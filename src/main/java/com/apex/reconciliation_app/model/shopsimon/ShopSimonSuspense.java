package com.apex.reconciliation_app.model.shopsimon;

import jakarta.persistence.*;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.time.LocalDateTime;

@Entity
@Table(name="shop_simon_suspense_queue")
@Data
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class ShopSimonSuspense implements ShopSimonTransactionData{

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;

    private String errorReason;

    @Builder.Default
    private LocalDateTime importTimeStamp = LocalDateTime.now();

    @Column(unique = true)
    private String compositeTransactionId;

    private String compositeId;

    // IDENTIFIERS
    private String type;
    private String entry;
    private String orderId;
    private String partnerId;
    private String poNumber;
    private String consumerOrderNumber;
    private String supplierOrderNumber;
    private String sku;
    private String partnerSku;
    private String itemId;
    private String returnId;

    // TIMESTAMPS
    private LocalDateTime transactionDate;
    private LocalDateTime invoiceDate;

    // TRANSACTION DETAILS
    private String currencyCode;
    private String lineItemFulfillment;
    private String taxType;
    private String adjustmentReason;
    private String trackingNumber;
    private Double lineNumber;
    private Double quantityShipped;
    private Double amountShipped;
    private Double quantityReturned;

    private String invoiceId;
    private String supplierInvoiceNumber;
    private String invoiceRecordType;
    private String suborderId;
    private Double quantityInvoiced;
    private Double orderQuantity;


    // FINANCIAL METRICS
    private Double consumerUnitPrice;
    private Double consumerUnitPriceIncludingTax;
    private Double taxPercentage;
    private Double unitTaxAmount;
    private Double totalTax;
    private Double consumerUnitPriceExcludingTax;
    private Double commissionPercentage;
    private Double totalCommissionAmount;
    private Double transactionAmount;
    private Double totalAmountPaid;
    private Double totalAmountOutstanding;
    private Double totalConsumerPrice;

    private Double unitCost;
    private Double totalUnitCost;
    private Double freightAmount;
    private Double salesTaxAmount;
    private Double charges;
    private Double credits;
    private Double subtotalExcludingLineItems;
}
