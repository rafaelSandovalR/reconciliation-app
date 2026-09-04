package com.apex.reconciliation_app.service;

import com.apex.reconciliation_app.enums.AmazonColumn;
import com.apex.reconciliation_app.model.AmazonRawTransaction;
import com.apex.reconciliation_app.model.AmazonSuspense;
import com.apex.reconciliation_app.model.AmazonTransactionData; // The shared interface
import org.apache.poi.ss.usermodel.Row;
import org.springframework.stereotype.Service;

@Service
public class AmazonReportService extends AbstractReportService<AmazonSuspense, AmazonRawTransaction, AmazonColumn> {

    @Override
    protected Class<AmazonColumn> getColumnEnumClass() {
        return AmazonColumn.class;
    }

    @Override
    protected String getSuspenseErrorReason(AmazonSuspense suspenseRecord) {
        return suspenseRecord.getErrorReason();
    }

    @Override
    protected void writeSuspenseData(Row row, AmazonSuspense record, int startingColIdx) {
        writeCommonData(row, record, startingColIdx); // Passes the interface
    }

    @Override
    protected void writeAuditData(Row row, AmazonRawTransaction record, int startingColIdx) {
        writeCommonData(row, record, startingColIdx); // Passes the interface
    }

    private void writeCommonData(Row row, AmazonTransactionData record, int col) {
        setCellValue(row.createCell(col++), record.getDateTime() != null ? record.getDateTime().format(formatter) : null);
        setCellValue(row.createCell(col++), record.getSettlementId());
        setCellValue(row.createCell(col++), record.getType());
        setCellValue(row.createCell(col++), record.getOrderId());
        setCellValue(row.createCell(col++), record.getSku());
        setCellValue(row.createCell(col++), record.getDescription());
        setCellValue(row.createCell(col++), record.getQuantity());
        setCellValue(row.createCell(col++), record.getMarketplace());
        setCellValue(row.createCell(col++), record.getAccountType());
        setCellValue(row.createCell(col++), record.getFulfillment());
        setCellValue(row.createCell(col++), record.getOrderCity());
        setCellValue(row.createCell(col++), record.getOrderState());
        setCellValue(row.createCell(col++), record.getOrderPostal());
        setCellValue(row.createCell(col++), record.getTaxCollectionModel());
        setCellValue(row.createCell(col++), record.getProductSales());
        setCellValue(row.createCell(col++), record.getProductSalesTax());
        setCellValue(row.createCell(col++), record.getShippingCredits());
        setCellValue(row.createCell(col++), record.getShippingCreditsTax());
        setCellValue(row.createCell(col++), record.getGiftWrapCredits());
        setCellValue(row.createCell(col++), record.getGiftWrapCreditsTax());
        setCellValue(row.createCell(col++), record.getRegulatoryFee());
        setCellValue(row.createCell(col++), record.getTaxOnRegulatoryFee());
        setCellValue(row.createCell(col++), record.getPromotionalRebates());
        setCellValue(row.createCell(col++), record.getPromotionalRebatesTax());
        setCellValue(row.createCell(col++), record.getMarketplaceWithheldTax());
        setCellValue(row.createCell(col++), record.getSellingFees());
        setCellValue(row.createCell(col++), record.getFbaFees());
        setCellValue(row.createCell(col++), record.getOtherTransactionFees());
        setCellValue(row.createCell(col++), record.getOther());
        setCellValue(row.createCell(col++), record.getTotal());
        setCellValue(row.createCell(col++), record.getTotal());
        setCellValue(row.createCell(col++), record.getTransactionStatus());
        setCellValue(row.createCell(col++), record.getTransactionReleaseDate() != null ? record.getTransactionReleaseDate().format(formatter) : null);
    }
}