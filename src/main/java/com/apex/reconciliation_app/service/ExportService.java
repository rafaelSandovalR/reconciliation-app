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

            // 3. Create the header row
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("Composite ID");
            headerRow.createCell(1).setCellValue("Site Name");
            headerRow.createCell(2).setCellValue("SKU");
            headerRow.createCell(3).setCellValue("Site Order ID");
            headerRow.createCell(4).setCellValue("Order Date");
            headerRow.createCell(5).setCellValue("Account");
            headerRow.createCell(6).setCellValue("Salesperson");

            headerRow.createCell(7).setCellValue("Total Less Tax");
            headerRow.createCell(8).setCellValue("Total Seller Cost");
            headerRow.createCell(9).setCellValue("Site Fees");
            headerRow.createCell(10).setCellValue("PayPal Fees");
            headerRow.createCell(11).setCellValue("CA Fees");
            headerRow.createCell(12).setCellValue("Pickpack Fees");
            headerRow.createCell(13).setCellValue("Shipping Costs Est");
            headerRow.createCell(14).setCellValue("Rithum Profit");

            headerRow.createCell(15).setCellValue("SiteOrderAmount");
            headerRow.createCell(16).setCellValue("SiteOrderFee");
            headerRow.createCell(17).setCellValue("SiteOrderFees1");
            headerRow.createCell(18).setCellValue("SiteOrderFees2");
            headerRow.createCell(19).setCellValue("SiteOrderFees3");
            headerRow.createCell(20).setCellValue("SiteOrderFees4");

            headerRow.createCell(21).setCellValue("ActualShippingCosts");
            headerRow.createCell(22).setCellValue("ShippingAdjustments");

            headerRow.createCell(23).setCellValue("ActualProfit");
            headerRow.createCell(24).setCellValue("ProfitDiff");
            headerRow.createCell(25).setCellValue("Return");
            headerRow.createCell(26).setCellValue("AmountRefunded");
            headerRow.createCell(27).setCellValue("ReturnShipping");
            headerRow.createCell(28).setCellValue("ReturnFee1");
            headerRow.createCell(29).setCellValue("ReturnFee2");
            headerRow.createCell(30).setCellValue("ReturnFee3");
            headerRow.createCell(31).setCellValue("ReturnFee4");

            headerRow.createCell(32).setCellValue("Notes");


            // 4. Loop through the database records and populate the rows
            int rowIndex = 1;
            for (ReconciliationRecord record : records) {
                Row row = sheet.createRow(rowIndex++);

                // String cells
                row.createCell(0).setCellValue(record.getCompositeId() != null ? record.getCompositeId() : "");
                row.createCell(1).setCellValue(record.getSiteName() != null ? record.getSiteName() : "");
                row.createCell(2).setCellValue(record.getSku() != null ? record.getSku() : "");
                row.createCell(3).setCellValue(record.getSiteOrderId() != null ? record.getSiteOrderId() : "");
                row.createCell(4).setCellValue(record.getOrderDate() != null ? record.getOrderDate() : "");
                row.createCell(5).setCellValue(record.getAccount() != null ? record.getAccount() : "");
                row.createCell(6).setCellValue(record.getSalesperson() != null ? record.getSalesperson() : "");

                // Numeric cells
                row.createCell(7).setCellValue(record.getTotalLessTax() != null ? record.getTotalLessTax() : 0.0);
                row.createCell(8).setCellValue(record.getTotalSellerCost() != null ? record.getTotalSellerCost() : 0.0);
                row.createCell(9).setCellValue(record.getSiteFees() != null ? record.getSiteFees() : 0.0);
                row.createCell(10).setCellValue(record.getPaypalFees() != null ? record.getPaypalFees() : 0.0);
                row.createCell(11).setCellValue(record.getCaFees() != null ? record.getCaFees() : 0.0);
                row.createCell(12).setCellValue(record.getPickPackFees() != null ? record.getPickPackFees() : 0.0);
                row.createCell(13).setCellValue(record.getShippingCostsEst() != null ? record.getShippingCostsEst() : 0.0);
                row.createCell(14).setCellValue(record.getRithumProfit() != null ? record.getRithumProfit() : 0.0);

                row.createCell(15).setCellValue(record.getSiteOrderAmount() != null ? record.getSiteOrderAmount() : 0.0);
                row.createCell(16).setCellValue(record.getSiteOrderFee() != null ? record.getSiteOrderFee() : 0.0);
                row.createCell(17).setCellValue(record.getSiteOrderOtherFees1() != null ? record.getSiteOrderOtherFees1() : 0.0);
                row.createCell(18).setCellValue(record.getSiteOrderOtherFees2() != null ? record.getSiteOrderOtherFees2() : 0.0);
                row.createCell(19).setCellValue(record.getSiteOrderOtherFees3() != null ? record.getSiteOrderOtherFees3() : 0.0);
                row.createCell(20).setCellValue(record.getSiteOrderOtherFees4() != null ? record.getSiteOrderOtherFees4() : 0.0);

                row.createCell(21).setCellValue(record.getActualShippingCosts() != null ? record.getActualShippingCosts() : 0.0);
                row.createCell(22).setCellValue(record.getShippingAdjustments() != null ? record.getShippingAdjustments(): 0.0);

                row.createCell(23).setCellValue(record.getActualProfit() != null ? record.getActualProfit() : 0.0);
                row.createCell(24).setCellValue(record.getProfitDiff() != null ? record.getProfitDiff() : 0.0);
                row.createCell(25).setCellValue(record.getReturnStatus() != null ? record.getReturnStatus() : "");
                row.createCell(26).setCellValue(record.getAmountRefunded() != null ? record.getAmountRefunded() : 0.0);
                row.createCell(27).setCellValue(record.getReturnShipping() != null ? record.getReturnShipping() : 0.0);
                row.createCell(28).setCellValue(record.getReturnFee1() != null ? record.getReturnFee1() : 0.0);
                row.createCell(29).setCellValue(record.getReturnFee2() != null ? record.getReturnFee2() : 0.0);
                row.createCell(30).setCellValue(record.getReturnFee3() != null ? record.getReturnFee3() : 0.0);
                row.createCell(31).setCellValue(record.getReturnFee4() != null ? record.getReturnFee4() : 0.0);

                row.createCell(32).setCellValue(record.getNotes() != null ? record.getNotes() : "");

            }

            // 5. Write the workbook to our byte stream and return it.
            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());

        } catch (Exception e) {
            throw new RuntimeException("Failed to export data to Excel: " + e.getMessage());
        }

    }
}
