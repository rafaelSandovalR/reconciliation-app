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
import java.util.*;

@Service
@RequiredArgsConstructor
public class RithumParserService {

    private final ReconciliationRepository repository;

    public void parseAndSaveInputStream(InputStream inputStream) {
        try (Workbook workbook = new XSSFWorkbook(inputStream)) {
            Sheet sheet = workbook.getSheetAt(0);

            Map<String, Integer> headerMap = ExcelUtils.getHeaderMap(sheet.getRow(0));

            // --- Extract all indices automatically
            Map<RithumColumn, Integer> indexMap = new EnumMap<>(RithumColumn.class);
            for (RithumColumn column: RithumColumn.values()) {
                Integer idx = headerMap.get(column.getHeaderName().toUpperCase());
                if (idx != null) {
                    indexMap.put(column, idx);
                }
            }

            // Safety check
            if (!indexMap.containsKey(RithumColumn.SITE_ORDER_ID) || !indexMap.containsKey(RithumColumn.SKU)) {
                throw new RuntimeException("Missing critical matching columns (Site Order ID or SKU) in Rithum file!");
            }

            List<ReconciliationRecord> records = new ArrayList<>();

            // Track IDs within this specific upload to prevent same-file duplicates
            Set<String> processedIds = new HashSet<>();

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                ReconciliationRecord record = new ReconciliationRecord();

                // Dynamically fill record; Loop through columns and apply enum setter
                for (Map.Entry<RithumColumn, Integer> entry: indexMap.entrySet()) {
                    RithumColumn column = entry.getKey();
                    Cell cell = row.getCell(entry.getValue());
                    column.applyTo(record, cell);
                }

                if (record.getSiteOrderId() == null || record.getSiteOrderId().isEmpty() ||
                    record.getSku() == null || record.getSku().isEmpty()) {
                    continue;
                }

                String compositeId = record.getSiteOrderId() + "-" + record.getSku();
                record.setCompositeId(compositeId);

                // IDEMPOTENCY CHECK
                if (processedIds.contains(compositeId) || repository.existsById(compositeId)) {
                    continue;
                }

                processedIds.add(compositeId);
                records.add(record);
            }

            // Only save brand new records
            repository.saveAll(records);
            System.out.println("Successfully saved " + records.size() + " new Rithum records.");

        } catch (Exception e) {
            throw new RuntimeException("Failed to parse Rithum Excel file:" + e.getMessage());
        }
    }
}
