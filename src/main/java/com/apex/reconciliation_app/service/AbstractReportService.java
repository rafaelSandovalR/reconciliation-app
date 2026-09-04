package com.apex.reconciliation_app.service;

import com.apex.reconciliation_app.dto.MarketplaceParseResult;
import com.apex.reconciliation_app.enums.ExcelColumn;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.format.DateTimeFormatter;
import java.util.List;

public abstract class AbstractReportService<S, A, E extends Enum<E> & ExcelColumn> {
    protected final DateTimeFormatter formatter = DateTimeFormatter.ofPattern("M/d/yyyy");

    protected abstract Class<E> getColumnEnumClass();
    protected abstract String getSuspenseErrorReason(S suspenseRecord);
    protected abstract void writeSuspenseData(Row row, S record, int startingColIdx);
    protected abstract void writeAuditData(Row row, A record, int startingColIdx);

    public ByteArrayInputStream generateReport(MarketplaceParseResult<S, A> result) {
        try (Workbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream out = new ByteArrayOutputStream()) {

            CellStyle headerStyle = createHeaderStyle(workbook);

            buildSuspenseSheet(workbook.createSheet("Suspense Queue"), result.suspenseQueue(), headerStyle);
            buildAuditSheet(workbook.createSheet("Audit Trail"), result.auditTrail(), headerStyle);

            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());

        } catch (IOException e) {
            throw new RuntimeException("Failed to generate Upload Report: " + e.getMessage());
        }
    }

    private void buildSuspenseSheet(Sheet sheet, List<S> data, CellStyle headerStyle) {
        Row headerRow = sheet.createRow(0);
        int colIdx = 0;

        Cell errorHeader = headerRow.createCell(colIdx++);
        errorHeader.setCellValue("Error Reason");
        errorHeader.setCellStyle(headerStyle);

        for (E col : getColumnEnumClass().getEnumConstants()) {
            Cell cell = headerRow.createCell(colIdx++);
            cell.setCellValue(col.getHeaderName());
            cell.setCellStyle(headerStyle);
        }

        int rowIdx = 1;
        for (S record : data) {
            Row row = sheet.createRow(rowIdx++);
            setCellValue(row.createCell(0), getSuspenseErrorReason(record));
            writeSuspenseData(row, record, 1); // Shifted by 1 for Error Reason
        }

        for (int i = 0; i < 10; i++) sheet.autoSizeColumn(i);
    }

    private void buildAuditSheet(Sheet sheet, List<A> data, CellStyle headerStyle) {
        Row headerRow = sheet.createRow(0);
        int colIdx = 0;

        for (E col : getColumnEnumClass().getEnumConstants()) {
            Cell cell = headerRow.createCell(colIdx++);
            cell.setCellValue(col.getHeaderName());
            cell.setCellStyle(headerStyle);
        }

        int rowIdx = 1;
        for (A record : data) {
            Row row = sheet.createRow(rowIdx++);
            writeAuditData(row, record, 0);
        }

        for (int i = 0; i < 9; i++) sheet.autoSizeColumn(i);
    }

    private CellStyle createHeaderStyle(Workbook workbook) {
        CellStyle headerStyle = workbook.createCellStyle();
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerStyle.setFont(headerFont);
        return headerStyle;
    }
    protected void setCellValue(Cell cell, String value) {
        cell.setCellValue(value != null ? value : "");
    }

    protected void setCellValue(Cell cell, Double value) {
        if (value != null) {
            cell.setCellValue(value);
        } else {
            cell.setCellValue("");
        }
    }
}
