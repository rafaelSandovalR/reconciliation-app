package com.apex.reconciliation_app.service;


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
                String siteOrderId = ExcelUtils.getStringValue(row.getCell(60));
                String sku = ExcelUtils.getStringValue(row.getCell(3));

                if (siteOrderId.isEmpty() || sku.isEmpty()) continue;

                // 2. Create the record and generate our Composite ID
                ReconciliationRecord record = new ReconciliationRecord();
                record.setCompositeId(siteOrderId + "-" + sku);

                // 3. Fill in the Rithum base data
                record.setSiteName(ExcelUtils.getStringValue(row.getCell(1)));
                record.setSku(sku);
                record.setSiteOrderId(siteOrderId);
                record.setOrderDate(ExcelUtils.getStringValue(row.getCell(5)));
                record.setAccount(ExcelUtils.getStringValue(row.getCell(10)));
                record.setSalesperson(ExcelUtils.getStringValue(row.getCell(85)));

                // 4. Fill in the financial data using the Double helper
                record.setTotalLessTax(ExcelUtils.getDoubleValue(row.getCell(76)));
                record.setTotalSellerCost(ExcelUtils.getDoubleValue(row.getCell(77)));
                record.setSiteFees(ExcelUtils.getDoubleValue(row.getCell(78)));
                record.setPaypalFees(ExcelUtils.getDoubleValue(row.getCell(79)));
                record.setCaFees(ExcelUtils.getDoubleValue(row.getCell(80)));
                record.setPickPackFees(ExcelUtils.getDoubleValue(row.getCell(81)));
                record.setShippingCostsEst(ExcelUtils.getDoubleValue(row.getCell(82)));
                record.setRithumProfit(ExcelUtils.getDoubleValue(row.getCell(83)));

                records.add(record);
            }
            repository.saveAll(records);
            System.out.println("Successfully saved " + records.size() + " Rithum records.");

        } catch (Exception e) {
            throw new RuntimeException("Failed to parse Rithum Excel file:" + e.getMessage());
        }
    }
}
