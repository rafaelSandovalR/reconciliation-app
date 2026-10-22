package com.apex.reconciliation_app.service;

import com.apex.reconciliation_app.enums.RithumColumn;
import com.apex.reconciliation_app.model.ReconciliationRecord;
import com.apex.reconciliation_app.repository.ReconciliationRepository;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
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

            // 3. Dynamically build header row from Enum
            Row headerRow = sheet.createRow(0);
            int headerColIdx = 0;
            for (RithumColumn column : RithumColumn.values()) {
                headerRow.createCell(headerColIdx++).setCellValue(column.getHeaderName());
            }

            // 4. Dynamically populate data rows from Enum getters
            int rowIdx = 1;
            for (ReconciliationRecord record : records) {
                Row row = sheet.createRow(rowIdx++);
                int colIdx = 0;

                for (RithumColumn column : RithumColumn.values()) {
                    Cell cell = row.createCell(colIdx++);
                    Object value = column.extractFrom(record);

                    if (value instanceof String) {
                        cell.setCellValue((String)value);
                    } else if (value instanceof Double) {
                        cell.setCellValue((Double) value);
                    } else if (value instanceof Integer) {
                        cell.setCellValue((Integer) value);
                    } else if (value instanceof LocalDateTime){
                        cell.setCellValue((LocalDateTime) value);
                    } else if (value != null) {
                        cell.setCellValue(value.toString());
                    } else {
                        cell.setCellValue("");
                    }
                }
            }

            // 5. Write the workbook to our byte stream and return it.
            workbook.write(out);
            return new ByteArrayInputStream(out.toByteArray());

        } catch (Exception e) {
            throw new RuntimeException("Failed to export data to Excel: " + e.getMessage());
        }

    }
}
