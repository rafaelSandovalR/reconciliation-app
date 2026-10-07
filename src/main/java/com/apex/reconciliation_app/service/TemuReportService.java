package com.apex.reconciliation_app.service;

import com.apex.reconciliation_app.enums.TemuColumn;
import com.apex.reconciliation_app.model.temu.TemuRawTransaction;
import com.apex.reconciliation_app.model.temu.TemuSuspense;
import com.apex.reconciliation_app.model.temu.TemuTransactionData;
import org.apache.poi.ss.usermodel.Row;
import org.springframework.stereotype.Service;

@Service
public class TemuReportService extends AbstractReportService<TemuSuspense, TemuRawTransaction, TemuColumn> {

    @Override
    protected Class<TemuColumn> getColumnEnumClass() { return TemuColumn.class; }


    @Override
    protected void writeSuspenseData(Row row, TemuSuspense record) {
        int col = 0;
        setCellValue(row.createCell(col++), record.getErrorReason());
        setCellValue(row.createCell(col++), record.getCompositeId());
        setCellValue(row.createCell(col++), record.getCompositeTransactionId());
        writeCommonData(row, record, col);
    }

    @Override
    protected void writeAuditData(Row row, TemuRawTransaction record) {
        int col = 0;
        setCellValue(row.createCell(col++), record.getCompositeId());
        setCellValue(row.createCell(col++), record.getCompositeTransactionId());
        writeCommonData(row, record, col);
    }

    private void writeCommonData(Row row, TemuTransactionData record, int col) {
        setCellValue(row.createCell(col++), record.getDateTime() != null ? record.getDateTime().format(formatter) : null);
        setCellValue(row.createCell(col++), record.getTransactionType());
        setCellValue(row.createCell(col++), record.getRelatedId());
        setCellValue(row.createCell(col++), record.getOrderId());
        setCellValue(row.createCell(col++), record.getOrderItemId());
        setCellValue(row.createCell(col++), record.getSku());
        setCellValue(row.createCell(col++), record.getSkuId());
        setCellValue(row.createCell(col++), record.getQuantity());
        setCellValue(row.createCell(col++), record.getShipCity());
        setCellValue(row.createCell(col++), record.getShipState());
        setCellValue(row.createCell(col++), record.getRetailPrice());
        setCellValue(row.createCell(col++), record.getPlatformDiscount());
        setCellValue(row.createCell(col++), record.getSellerDiscount());
        setCellValue(row.createCell(col++), record.getServiceFee());
        setCellValue(row.createCell(col++), record.getServiceFeeTax());
        setCellValue(row.createCell(col++), record.getPlatformIncentive());
        setCellValue(row.createCell(col++), record.getSubtotal());
        setCellValue(row.createCell(col++), record.getShipping());
        setCellValue(row.createCell(col++), record.getPlatformIncentiveShipping());
        setCellValue(row.createCell(col++), record.getProductTax());
        setCellValue(row.createCell(col++), record.getShippingTax());
        setCellValue(row.createCell(col++), record.getSignOnDelivery());
        setCellValue(row.createCell(col++), record.getSignOnDeliveryTax());
        setCellValue(row.createCell(col++), record.getMarketplaceWithheldTax());
        setCellValue(row.createCell(col++), record.getOthers());
        setCellValue(row.createCell(col++), record.getTotal());
        setCellValue(row.createCell(col++), record.getCurrency());
    }
}
