package com.apex.reconciliation_app.service;

import com.apex.reconciliation_app.model.ReconciliationRecord;
import com.apex.reconciliation_app.repository.ReconciliationRepository;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.List;

@Service
@RequiredArgsConstructor
public class ExportService {

    private final ReconciliationRepository repository;

    public ByteArrayInputStream exportToExcel() {
        // 1. Grab all the saved data from the db
        List<ReconciliationRecord> records = repository.findAll();

        // 2. Create a new workbook and sheet in memory
        try (Workbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream out = new ByteArrayOutputStream())
        {
            Sheet sheet = workbook.createSheet("Reconciliation Data");
            int idx = 0;
            // 3. Create the header row
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(idx++).setCellValue("Composite ID");
            headerRow.createCell(idx++).setCellValue("Site Name");
            headerRow.createCell(idx++).setCellValue("SKU");
            headerRow.createCell(idx++).setCellValue("Site Order ID");
            headerRow.createCell(idx++).setCellValue("Site Order Item ID");
            headerRow.createCell(idx++).setCellValue("Order Date");
            headerRow.createCell(idx++).setCellValue("Account");
            headerRow.createCell(idx++).setCellValue("Salesperson");

            headerRow.createCell(idx++).setCellValue("Total Less Tax");
            headerRow.createCell(idx++).setCellValue("Total Seller Cost");
            headerRow.createCell(idx++).setCellValue("Site Fees");
            headerRow.createCell(idx++).setCellValue("PayPal Fees");
            headerRow.createCell(idx++).setCellValue("CA Fees");
            headerRow.createCell(idx++).setCellValue("Pickpack Fees");
            headerRow.createCell(idx++).setCellValue("Shipping Costs Est");
            headerRow.createCell(idx++).setCellValue("Rithum Profit");

            headerRow.createCell(idx++).setCellValue("SiteOrderAmount");
            headerRow.createCell(idx++).setCellValue("SiteOrderFee");
            headerRow.createCell(idx++).setCellValue("SiteOrderFees1");
            headerRow.createCell(idx++).setCellValue("SiteOrderFees2");
            headerRow.createCell(idx++).setCellValue("SiteOrderFees3");
            headerRow.createCell(idx++).setCellValue("SiteOrderFees4");

            headerRow.createCell(idx++).setCellValue("ActualShippingCosts");
            headerRow.createCell(idx++).setCellValue("ShippingAdjustments");

            headerRow.createCell(idx++).setCellValue("ActualProfit");
            headerRow.createCell(idx++).setCellValue("ProfitDiff");
            headerRow.createCell(idx++).setCellValue("Return");
            headerRow.createCell(idx++).setCellValue("AmountRefunded");
            headerRow.createCell(idx++).setCellValue("ReturnShipping");
            headerRow.createCell(idx++).setCellValue("ReturnFee1");
            headerRow.createCell(idx++).setCellValue("ReturnFee2");
            headerRow.createCell(idx++).setCellValue("ReturnFee3");
            headerRow.createCell(idx++).setCellValue("ReturnFee4");

            headerRow.createCell(idx++).setCellValue("Notes");


            // 4. Loop through the database records and populate the rows
            int rowIndex = 1;
            for (ReconciliationRecord record : records) {
                Row row = sheet.createRow(rowIndex++);
                int colIndex = 0;

                // String cells
                row.createCell(colIndex++).setCellValue(record.getCompositeId() != null ? record.getCompositeId() : "");
                row.createCell(colIndex++).setCellValue(record.getSiteName() != null ? record.getSiteName() : "");
                row.createCell(colIndex++).setCellValue(record.getSku() != null ? record.getSku() : "");
                row.createCell(colIndex++).setCellValue(record.getSiteOrderId() != null ? record.getSiteOrderId() : "");
                row.createCell(colIndex++).setCellValue(record.getSiteOrderItemId() != null ? record.getSiteOrderItemId() : "");
                row.createCell(colIndex++).setCellValue(record.getOrderDate() != null ? record.getOrderDate() : "");
                row.createCell(colIndex++).setCellValue(record.getAccount() != null ? record.getAccount() : "");
                row.createCell(colIndex++).setCellValue(record.getSalesperson() != null ? record.getSalesperson() : "");

                // Numeric cells
                row.createCell(colIndex++).setCellValue(record.getTotalLessTax() != null ? record.getTotalLessTax() : 0.0);
                row.createCell(colIndex++).setCellValue(record.getTotalSellerCost() != null ? record.getTotalSellerCost() : 0.0);
                row.createCell(colIndex++).setCellValue(record.getSiteFees() != null ? record.getSiteFees() : 0.0);
                row.createCell(colIndex++).setCellValue(record.getPaypalFees() != null ? record.getPaypalFees() : 0.0);
                row.createCell(colIndex++).setCellValue(record.getCaFees() != null ? record.getCaFees() : 0.0);
                row.createCell(colIndex++).setCellValue(record.getPickPackFees() != null ? record.getPickPackFees() : 0.0);
                row.createCell(colIndex++).setCellValue(record.getShippingCostsEst() != null ? record.getShippingCostsEst() : 0.0);
                row.createCell(colIndex++).setCellValue(record.getRithumProfit() != null ? record.getRithumProfit() : 0.0);

                row.createCell(colIndex++).setCellValue(record.getSiteOrderAmount() != null ? record.getSiteOrderAmount() : 0.0);
                row.createCell(colIndex++).setCellValue(record.getSiteOrderFee() != null ? record.getSiteOrderFee() : 0.0);
                row.createCell(colIndex++).setCellValue(record.getSiteOrderOtherFees1() != null ? record.getSiteOrderOtherFees1() : 0.0);
                row.createCell(colIndex++).setCellValue(record.getSiteOrderOtherFees2() != null ? record.getSiteOrderOtherFees2() : 0.0);
                row.createCell(colIndex++).setCellValue(record.getSiteOrderOtherFees3() != null ? record.getSiteOrderOtherFees3() : 0.0);
                row.createCell(colIndex++).setCellValue(record.getSiteOrderOtherFees4() != null ? record.getSiteOrderOtherFees4() : 0.0);

                row.createCell(colIndex++).setCellValue(record.getActualShippingCosts() != null ? record.getActualShippingCosts() : 0.0);
                row.createCell(colIndex++).setCellValue(record.getShippingAdjustments() != null ? record.getShippingAdjustments(): 0.0);

                row.createCell(colIndex++).setCellValue(record.getActualProfit() != null ? record.getActualProfit() : 0.0);
                row.createCell(colIndex++).setCellValue(record.getProfitDiff() != null ? record.getProfitDiff() : 0.0);
                row.createCell(colIndex++).setCellValue(record.getReturnStatus() != null ? record.getReturnStatus() : "");
                row.createCell(colIndex++).setCellValue(record.getAmountRefunded() != null ? record.getAmountRefunded() : 0.0);
                row.createCell(colIndex++).setCellValue(record.getReturnShipping() != null ? record.getReturnShipping() : 0.0);
                row.createCell(colIndex++).setCellValue(record.getReturnFee1() != null ? record.getReturnFee1() : 0.0);
                row.createCell(colIndex++).setCellValue(record.getReturnFee2() != null ? record.getReturnFee2() : 0.0);
                row.createCell(colIndex++).setCellValue(record.getReturnFee3() != null ? record.getReturnFee3() : 0.0);
                row.createCell(colIndex++).setCellValue(record.getReturnFee4() != null ? record.getReturnFee4() : 0.0);

                row.createCell(colIndex++).setCellValue(record.getNotes() != null ? record.getNotes() : "");

            }

            // 5. Write the workbook to our byte stream and return it.
            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());

        } catch (Exception e) {
            throw new RuntimeException("Failed to export data to Excel: " + e.getMessage());
        }

    }
}
