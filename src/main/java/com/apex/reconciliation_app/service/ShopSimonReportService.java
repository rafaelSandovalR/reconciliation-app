package com.apex.reconciliation_app.service;

import com.apex.reconciliation_app.enums.ShopSimonColumn;
import com.apex.reconciliation_app.model.shopsimon.ShopSimonRawTransaction;
import com.apex.reconciliation_app.model.shopsimon.ShopSimonSuspense;
import com.apex.reconciliation_app.model.shopsimon.ShopSimonTransactionData;
import org.apache.poi.ss.usermodel.Row;
import org.springframework.stereotype.Service;
import java.time.format.DateTimeFormatter;

@Service
public class ShopSimonReportService  extends AbstractReportService<ShopSimonSuspense, ShopSimonRawTransaction, ShopSimonColumn> {

    private static final DateTimeFormatter SHOP_SIMON_FORMATTER =
            DateTimeFormatter.ofPattern("yyyy-MM-dd'T'HH:mm:ss");

    @Override
    protected Class<ShopSimonColumn> getColumnEnumClass() {
        return ShopSimonColumn.class;
    }

    @Override
    protected String getSuspenseErrorReason(ShopSimonSuspense suspenseRecord) {
        return suspenseRecord.getErrorReason();
    }

    @Override
    protected String getCompositeId(ShopSimonSuspense suspenseRecord) {
        return suspenseRecord.getCompositeId();
    }

    @Override
    protected void writeSuspenseData(Row row, ShopSimonSuspense record, int startingColIdx) {
        writeCommonData(row, record, startingColIdx);
    }

    @Override
    protected void writeAuditData(Row row, ShopSimonRawTransaction record, int startingColIdx) {
        writeCommonData(row, record, startingColIdx);
    }

    private void writeCommonData(Row row, ShopSimonTransactionData record, int col) {
        setCellValue(row.createCell(col++), record.getType());
        setCellValue(row.createCell(col++), record.getEntry());
        setCellValue(row.createCell(col++), record.getOrderId());
        setCellValue(row.createCell(col++), record.getPartnerId());
        setCellValue(row.createCell(col++), record.getPoNumber());
        setCellValue(row.createCell(col++), record.getConsumerOrderNumber());
        setCellValue(row.createCell(col++), record.getSupplierOrderNumber());
        setCellValue(row.createCell(col++), record.getCurrencyCode());
        setCellValue(row.createCell(col++), record.getLineNumber());
        setCellValue(row.createCell(col++), record.getSku());
        setCellValue(row.createCell(col++), record.getPartnerSku());
        setCellValue(row.createCell(col++), record.getItemId());
        setCellValue(row.createCell(col++), record.getTransactionDate() != null ? record.getTransactionDate().format(SHOP_SIMON_FORMATTER) : null);
        setCellValue(row.createCell(col++), record.getQuantityShipped());
        setCellValue(row.createCell(col++), record.getAmountShipped());
        setCellValue(row.createCell(col++), record.getReturnId());
        setCellValue(row.createCell(col++), record.getQuantityReturned());
        setCellValue(row.createCell(col++), record.getLineItemFulfillment());
        setCellValue(row.createCell(col++), record.getConsumerUnitPrice());
        setCellValue(row.createCell(col++), record.getConsumerUnitPriceIncludingTax());
        setCellValue(row.createCell(col++), record.getTaxType());
        setCellValue(row.createCell(col++), record.getTaxPercentage());
        setCellValue(row.createCell(col++), record.getUnitTaxAmount());
        setCellValue(row.createCell(col++), record.getTotalTax());
        setCellValue(row.createCell(col++), record.getConsumerUnitPriceExcludingTax());
        setCellValue(row.createCell(col++), record.getCommissionPercentage());
        setCellValue(row.createCell(col++), record.getTotalCommissionAmount());
        setCellValue(row.createCell(col++), record.getAdjustmentReason());
        setCellValue(row.createCell(col++), record.getTransactionAmount());
        setCellValue(row.createCell(col++), record.getTotalAmountPaid());
        setCellValue(row.createCell(col++), record.getTotalAmountOutstanding());
        setCellValue(row.createCell(col++), record.getTotalConsumerPrice());
        setCellValue(row.createCell(col++), record.getTrackingNumber());
        setCellValue(row.createCell(col++), record.getInvoiceId());
        setCellValue(row.createCell(col++), record.getSupplierInvoiceNumber());
        setCellValue(row.createCell(col++), record.getInvoiceRecordType());
        setCellValue(row.createCell(col++), record.getSuborderId());
        setCellValue(row.createCell(col++), record.getInvoiceDate() != null ? record.getInvoiceDate().format(SHOP_SIMON_FORMATTER) : null);
        setCellValue(row.createCell(col++), record.getQuantityInvoiced());
        setCellValue(row.createCell(col++), record.getOrderQuantity());
        setCellValue(row.createCell(col++), record.getUnitCost());
        setCellValue(row.createCell(col++), record.getTotalUnitCost());
        setCellValue(row.createCell(col++), record.getFreightAmount());
        setCellValue(row.createCell(col++), record.getSalesTaxAmount());
        setCellValue(row.createCell(col++), record.getCharges());
        setCellValue(row.createCell(col++), record.getCredits());
        setCellValue(row.createCell(col++), record.getSubtotalExcludingLineItems());

    }
}
