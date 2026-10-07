package com.apex.reconciliation_app.service;

import com.apex.reconciliation_app.dto.MarketplaceParseResult;
import com.apex.reconciliation_app.enums.WalmartColumn;
import com.apex.reconciliation_app.model.walmart.WalmartRawTransaction;
import com.apex.reconciliation_app.model.walmart.WalmartSuspense;
import com.apex.reconciliation_app.model.walmart.WalmartTransactionData;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.format.DateTimeFormatter;
import java.util.Date;
import java.util.List;

@Service
public class WalmartReportService extends AbstractReportService<WalmartSuspense, WalmartRawTransaction, WalmartColumn>{

    private static final DateTimeFormatter WALMART_FORMATTER =
            DateTimeFormatter.ofPattern("M/d/yyyy");

    @Override
    protected Class<WalmartColumn> getColumnEnumClass() { return WalmartColumn.class; }

    @Override
    protected void writeSuspenseData(Row row, WalmartSuspense record) {
        int col = 0;
        setCellValue(row.createCell(col++), record.getErrorReason());
        setCellValue(row.createCell(col++), record.getCompositeId());
        setCellValue(row.createCell(col++), record.getCompositeTransactionId());
        writeCommonData(row, record, col);
    }

    @Override
    protected void writeAuditData(Row row, WalmartRawTransaction record) {
        int col = 0;
        setCellValue(row.createCell(col++), record.getCompositeId());
        setCellValue(row.createCell(col++), record.getCompositeTransactionId());
        writeCommonData(row, record, col);
    }

    private void writeCommonData(Row row, WalmartTransactionData record, int col) {
        setCellValue(row.createCell(col++), record.getTransactionKey());
        setCellValue(row.createCell(col++), record.getTransactionPostedTimestamp() != null ? record.getTransactionPostedTimestamp().format(WALMART_FORMATTER) : null);
        setCellValue(row.createCell(col++), record.getTransactionType());
        setCellValue(row.createCell(col++), record.getTransactionDesc());
        setCellValue(row.createCell(col++), record.getCustomerOrder());
        setCellValue(row.createCell(col++), record.getCustomerOrderLine());
        setCellValue(row.createCell(col++), record.getPurchaseOrder());
        setCellValue(row.createCell(col++), record.getPurchaseOrderLine());
        setCellValue(row.createCell(col++), record.getAmount());
        setCellValue(row.createCell(col++), record.getAmountType());
        setCellValue(row.createCell(col++), record.getShipQty());
        setCellValue(row.createCell(col++), record.getCommissionRate());
        setCellValue(row.createCell(col++), record.getBaseCommissionRate());
        setCellValue(row.createCell(col++), record.getTransactionReasonDesc());
        setCellValue(row.createCell(col++), record.getPartnerItemId());
        setCellValue(row.createCell(col++), record.getPartnerGtIn());
        setCellValue(row.createCell(col++), record.getPartnerItemName());
        setCellValue(row.createCell(col++), record.getProductTaxCode());
        setCellValue(row.createCell(col++), record.getShipToState());
        setCellValue(row.createCell(col++), record.getShipToCity());
        setCellValue(row.createCell(col++), record.getShipToZipcode());
        setCellValue(row.createCell(col++), record.getContractCategory());
        setCellValue(row.createCell(col++), record.getProductType());
        setCellValue(row.createCell(col++), record.getCommissionRule());
        setCellValue(row.createCell(col++), record.getShippingMethod());
        setCellValue(row.createCell(col++), record.getFulfillmentType());
        setCellValue(row.createCell(col++), record.getFulfillmentDetails());
        setCellValue(row.createCell(col++), record.getOriginalCommission());
        setCellValue(row.createCell(col++), record.getCommissionIncentiveProgram());
        setCellValue(row.createCell(col++), record.getCommissionSaving());
        setCellValue(row.createCell(col++), record.getCustomerPromoType());
        setCellValue(row.createCell(col++), record.getTotalWalmartFundedSavings());
        setCellValue(row.createCell(col++), record.getCampaignId());
        setCellValue(row.createCell(col++), record.getItemCondition());
        setCellValue(row.createCell(col++), record.getOriginalCharge());
        setCellValue(row.createCell(col++), record.getChargeSavings());
        setCellValue(row.createCell(col++), record.getIncentiveProgramName());
        setCellValue(row.createCell(col++), record.getShipToCountry());
    }
}
