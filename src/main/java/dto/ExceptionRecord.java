package dto;

public record ExceptionRecord(
        String siteOrderId,
        String sku,
        double amount,
        String transactionDate,
        String errorReason
) {
}
