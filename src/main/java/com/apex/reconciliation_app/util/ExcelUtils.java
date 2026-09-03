package com.apex.reconciliation_app.util;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;

import java.time.LocalDateTime;
import java.util.HashMap;
import java.util.Map;

public class ExcelUtils {

    // Private constructor to prevent accidental instantiation
    private ExcelUtils() {
        throw new IllegalStateException("Utility class");
    }

    // --- SAFE DATA EXTRACTION HELPERS ---
    public static String getStringValue(Cell cell) {
        if (cell == null) return "";
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue().trim();
            case NUMERIC -> {
                if (DateUtil.isCellDateFormatted(cell)) {
                    java.text.SimpleDateFormat dateFormat = new java.text.SimpleDateFormat("MM/dd/yyyy HH:mm");
                    yield dateFormat.format(cell.getDateCellValue());
                } else {
                    yield String.valueOf((long) cell.getNumericCellValue());
                }
            }
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            default -> "";
        };
    }

    public static Double getDoubleValue(Cell cell) {
        if (cell == null) return 0.0;
        return switch (cell.getCellType()) {
            case NUMERIC -> cell.getNumericCellValue();
            case STRING -> {
                try {
                    yield Double.parseDouble(cell.getStringCellValue().replace("$", "").trim());
                } catch (NumberFormatException e) {
                    yield 0.0;
                }
            }
            default -> 0.0;
        };
    }

    public static Map<String, Integer> getHeaderMap(Row headerRow) {
        Map<String, Integer> headerMap = new HashMap<>();

        for (Cell cell: headerRow) {
            String headerName = getStringValue(cell).trim().toUpperCase();
            if (!headerName.isEmpty()) {
                headerMap.put(headerName, cell.getColumnIndex());
            }
        }

        return headerMap;
    }

    // SAFE EXTRACTION UTILS FOR ANY ENUM
    public static <E extends Enum<E>> String getStringSafe(Row row, Map<E, Integer> headerMap, E col) {
        if (!headerMap.containsKey(col)) return null;
        return getStringValue(row.getCell((headerMap.get(col))));
    }

    public static <E extends Enum<E>> Double getDoubleSafe(Row row, Map<E, Integer> headerMap, E col) {
        if (!headerMap.containsKey(col)) return null;
        Cell cell = row.getCell(headerMap.get(col));
        if (cell == null || cell.getCellType() == CellType.BLANK) return null;
        return getDoubleValue(cell);
    }

    public static <E extends Enum<E>>LocalDateTime getDateSafe(Row row, Map<E, Integer> headerMap, E col) {
        if(!headerMap.containsKey(col)) return null;
        Cell cell = row.getCell(headerMap.get(col));
        if (cell != null && cell.getCellType() != CellType.BLANK && DateUtil.isCellDateFormatted(cell)) {
            return cell.getLocalDateTimeCellValue();
        }
        return null;
    }
}
