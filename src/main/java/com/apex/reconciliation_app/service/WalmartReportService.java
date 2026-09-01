package com.apex.reconciliation_app.service;

import com.apex.reconciliation_app.enums.WalmartColumn;
import com.apex.reconciliation_app.model.WalmartRawTransaction;
import com.apex.reconciliation_app.model.WalmartSuspense;
import com.apex.reconciliation_app.dto.WalmartParseResult;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.format.DateTimeFormatter;
import java.util.List;

@Service
public class WalmartReportService {

    public ByteArrayInputStream generateReport(WalmartParseResult result) {

        try (Workbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream out = new ByteArrayOutputStream()) {

            DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");
            CellStyle headerStyle = createHeaderStyle(workbook);

            // Sheet 1: Suspense Queue
            Sheet suspenseSheet = workbook.createSheet("Suspense Queue");
            buildSuspenseSheet(suspenseSheet, result.suspenseQueue(), headerStyle, formatter);

            // Sheet 2: Audit Trail
            Sheet auditSheet = workbook.createSheet("Audit Trail");
            buildAuditSheet(auditSheet, result.auditTrail(), headerStyle, formatter);

            // --- WRITE TO STREAM ---
            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());

        } catch (IOException e) {
            throw new RuntimeException("Failed to generate Upload Report: " + e.getMessage());
        }
    }

    // Sheet Generation Helpers
    private void buildSuspenseSheet(Sheet sheet, List<WalmartSuspense> data, CellStyle headerStyle, DateTimeFormatter formatter) {
        Row headerRow = sheet.createRow(0);
        int colIdx = 0;

        Cell errorHeader = headerRow.createCell(colIdx++);
        errorHeader.setCellValue("Error Reason");
        errorHeader.setCellStyle(headerStyle);

        for (WalmartColumn col : WalmartColumn.values()) {
            Cell cell = headerRow.createCell(colIdx++);
            cell.setCellValue(col.getHeaderName());
            cell.setCellStyle(headerStyle);
        }

        int rowIdx = 1;
        for (WalmartSuspense record : data) {
            Row row = sheet.createRow(rowIdx++);
            int col = 0;

            setCellValue(row.createCell(col++), record.getErrorReason());
            setCellValue(row.createCell(col++), record.getTransactionKey());
            setCellValue(row.createCell(col++), record.getTransactionPostedTimestamp() != null ? record.getTransactionPostedTimestamp().format(formatter) : null);
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

        for (int i = 0; i < 10; i++) sheet.autoSizeColumn(i);
    }

    private void buildAuditSheet(Sheet sheet, List<WalmartRawTransaction> data, CellStyle headerStyle, DateTimeFormatter formatter) {
        Row headerRow = sheet.createRow(0);
        int colIdx = 0;

        for (WalmartColumn col : WalmartColumn.values()) {
            Cell cell = headerRow.createCell(colIdx++);
            cell.setCellValue(col.getHeaderName());
            cell.setCellStyle(headerStyle);
        }

        int rowIdx = 1;
        for (WalmartRawTransaction record : data) {
            Row row = sheet.createRow(rowIdx++);
            int col = 0;

            setCellValue(row.createCell(col++), record.getTransactionKey());
            setCellValue(row.createCell(col++), record.getTransactionPostedTimestamp() != null ? record.getTransactionPostedTimestamp().format(formatter) : null);
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

        for (int i = 0; i < 9; i++) sheet.autoSizeColumn(i);
    }

    // --- CELL UTILITIES ---

    private CellStyle createHeaderStyle(Workbook workbook) {
        CellStyle headerStyle = workbook.createCellStyle();
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerStyle.setFont(headerFont);
        return headerStyle;
    }

    private void setCellValue(Cell cell, String value) {
        cell.setCellValue(value != null ? value : "");
    }

    private void setCellValue(Cell cell, Double value) {
        if (value != null) {
            cell.setCellValue(value);
        } else {
            cell.setCellValue("");
        }
    }
}
