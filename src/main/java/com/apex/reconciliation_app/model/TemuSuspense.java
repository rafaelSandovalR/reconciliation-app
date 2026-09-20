package com.apex.reconciliation_app.model;

import jakarta.persistence.*;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.time.LocalDateTime;


@Entity
@Table(name="temu_suspense_queue")
@Data
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class TemuSuspense implements TemuTransactionData{

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;

    private String errorReason;

    @Builder.Default
    private LocalDateTime importTimeStamp = LocalDateTime.now();

    @Column(unique = true)
    private String compositeTransactionId; // TODO: Apply to other suspense entities

    // IDENTIFIERS
    private String sku;
    private String skuId;
    private String relatedId;
    private String orderId;
    private String orderItemId;

    // Timestamps
    private LocalDateTime dateTime;

    // TRANSACTION DETAILS
    private String transactionType;
    private String currency;
    private Double quantity;

    // SHIPPING INFO
    private String shipCity;
    private String shipState;

    // FINANCIAL METRICS
    private Double retailPrice;
    private Double platformDiscount;
    private Double sellerDiscount;
    private Double serviceFee;
    private Double serviceFeeTax;
    private Double platformIncentive;
    private Double subtotal;
    private Double shipping;
    private Double platformIncentiveShipping;
    private Double productTax;
    private Double shippingTax;
    private Double signOnDelivery;
    private Double signOnDeliveryTax;
    private Double marketplaceWithheldTax;
    private Double others;
    private Double total;
}
