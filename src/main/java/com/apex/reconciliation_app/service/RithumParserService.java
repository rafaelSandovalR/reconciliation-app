package com.apex.reconciliation_app.service;


import com.apex.reconciliation_app.enums.RithumColumn;
import com.apex.reconciliation_app.model.ReconciliationRecord;
import com.apex.reconciliation_app.repository.ReconciliationRepository;
import com.apex.reconciliation_app.util.ExcelUtils;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

@Service
@RequiredArgsConstructor
public class RithumParserService {

    private final ReconciliationRepository repository;

    public void parseAndSaveInputStream(InputStream inputStream) {
        try (Workbook workbook = new XSSFWorkbook(inputStream)) {

            Sheet sheet = workbook.getSheetAt(0);
            List<ReconciliationRecord> records = new ArrayList<>();

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                // 1. Grab the critical matching data first
                String siteOrderId = ExcelUtils.getStringValue(row.getCell(RithumColumn.SITE_ORDER_ID.getIndex()));
                String sku = ExcelUtils.getStringValue(row.getCell(RithumColumn.SKU.getIndex()));

                if (siteOrderId.isEmpty() || sku.isEmpty()) continue;

                // 2. Create the record and generate our Composite ID
                ReconciliationRecord record = new ReconciliationRecord();
                record.setCompositeId(siteOrderId + "-" + sku);

                // 3. Fill in the Rithum base data
                record.setSiteName(ExcelUtils.getStringValue(row.getCell(RithumColumn.SITE_NAME.getIndex())));
                record.setSku(sku);
                record.setSiteOrderId(siteOrderId);
                record.setOrderDate(ExcelUtils.getStringValue(row.getCell(RithumColumn.ORDER_DATE.getIndex())));
                record.setAccount(ExcelUtils.getStringValue(row.getCell(RithumColumn.ACCOUNT.getIndex())));
                record.setSalesperson(ExcelUtils.getStringValue(row.getCell(RithumColumn.SALESPERSON.getIndex())));

                // 4. Fill in the financial data using the Double helper
                record.setTotalLessTax(ExcelUtils.getDoubleValue(row.getCell(RithumColumn.TOTAL_LESS_TAX.getIndex())));
                record.setTotalSellerCost(ExcelUtils.getDoubleValue(row.getCell(RithumColumn.TOTAL_SELLER_COST.getIndex())));
                record.setSiteFees(ExcelUtils.getDoubleValue(row.getCell(RithumColumn.SITE_FEES.getIndex())));
                record.setPaypalFees(ExcelUtils.getDoubleValue(row.getCell(RithumColumn.PAYPAL_FEES.getIndex())));
                record.setCaFees(ExcelUtils.getDoubleValue(row.getCell(RithumColumn.CA_FEES.getIndex())));
                record.setPickPackFees(ExcelUtils.getDoubleValue(row.getCell(RithumColumn.PICK_PACK_FEES.getIndex())));
                record.setShippingCostsEst(ExcelUtils.getDoubleValue(row.getCell(RithumColumn.SHIPPING_COSTS_EST.getIndex())));
                record.setRithumProfit(ExcelUtils.getDoubleValue(row.getCell(RithumColumn.PROFIT.getIndex())));

                records.add(record);
            }
            repository.saveAll(records);
            System.out.println("Successfully saved " + records.size() + " Rithum records.");

        } catch (Exception e) {
            throw new RuntimeException("Failed to parse Rithum Excel file:" + e.getMessage());
        }
    }
}
