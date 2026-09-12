package com.apex.reconciliation_app.service;

import com.apex.reconciliation_app.enums.EbayColumn;
import com.apex.reconciliation_app.model.EbayRawTransaction;
import com.apex.reconciliation_app.model.EbaySuspense;
import com.apex.reconciliation_app.model.EbayTransactionData;
import org.apache.poi.ss.usermodel.Row;
import org.springframework.stereotype.Service;

@Service
public class EbayReportService extends AbstractReportService<EbaySuspense, EbayRawTransaction, EbayColumn> {

    @Override
    protected Class<EbayColumn> getColumnEnumClass() {
        return EbayColumn.class;
    }

    @Override
    protected String getSuspenseErrorReason(EbaySuspense suspenseRecord) {
        return suspenseRecord.getErrorReason();
    }

    @Override
    protected void writeSuspenseData(Row row, EbaySuspense record, int startingColIdx) {
        writeCommonData(row, record, startingColIdx);
    }

    @Override
    protected void writeAuditData(Row row, EbayRawTransaction record, int startingColIdx) {
        writeCommonData(row, record, startingColIdx);
    }

    private void writeCommonData(Row row, EbayTransactionData record, int col) {
        setCellValue(row.createCell(col++), record.getTransactionCreationDate() != null ? record.getTransactionCreationDate().format(formatter) : null);
        setCellValue(row.createCell(col++), record.getType());
        setCellValue(row.createCell(col++), record.getOrderNumber());
        setCellValue(row.createCell(col++), record.getLegacyOrderId());
        setCellValue(row.createCell(col++), record.getBuyerUsername());
        setCellValue(row.createCell(col++), record.getBuyerName());
        setCellValue(row.createCell(col++), record.getShipToCity());
        setCellValue(row.createCell(col++), record.getShipToProvinceRegionState());
        setCellValue(row.createCell(col++), record.getShipToZip());
        setCellValue(row.createCell(col++), record.getShipToCountry());
        setCellValue(row.createCell(col++), record.getNetAmount());
        setCellValue(row.createCell(col++), record.getPayoutCurrency());
        setCellValue(row.createCell(col++), record.getPayoutDate() != null ? record.getPayoutDate().format(formatter) : null);
        setCellValue(row.createCell(col++), record.getPayoutId());
        setCellValue(row.createCell(col++), record.getPayoutMethod());
        setCellValue(row.createCell(col++), record.getPayoutStatus());
        setCellValue(row.createCell(col++), record.getReasonForHold());
        setCellValue(row.createCell(col++), record.getItemId());
        setCellValue(row.createCell(col++), record.getTransactionId());
        setCellValue(row.createCell(col++), record.getItemTitle());
        setCellValue(row.createCell(col++), record.getCustomLabel());
        setCellValue(row.createCell(col++), record.getQuantity());
        setCellValue(row.createCell(col++), record.getItemSubtotal());
        setCellValue(row.createCell(col++), record.getShippingAndHandling());
        setCellValue(row.createCell(col++), record.getSellerCollectedTax());
        setCellValue(row.createCell(col++), record.getEbayCollectedTax());
        setCellValue(row.createCell(col++), record.getFinalValueFeeFixed());
        setCellValue(row.createCell(col++), record.getFinalValueFeeVariable());
        setCellValue(row.createCell(col++), record.getRegulatoryOperatingFee());
        setCellValue(row.createCell(col++), record.getVeryHighItemNotAsDescribedFee());
        setCellValue(row.createCell(col++), record.getBelowStandardPerformanceFee());
        setCellValue(row.createCell(col++), record.getInternationalFee());
        setCellValue(row.createCell(col++), record.getCharityDonation());
        setCellValue(row.createCell(col++), record.getDepositProcessingFee());
        setCellValue(row.createCell(col++), record.getGrossTransactionAmount());
        setCellValue(row.createCell(col++), record.getTransactionCurrency());
        setCellValue(row.createCell(col++), record.getExchangeRate());
        setCellValue(row.createCell(col++), record.getReferenceId());
        setCellValue(row.createCell(col++), record.getDescription());
    }
}
