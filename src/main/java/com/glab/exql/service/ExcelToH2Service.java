package com.glab.exql.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ClassPathResource;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Service;

import java.io.InputStream;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

@Service
public class ExcelToH2Service {

    private static final DateTimeFormatter FORMATTER = DateTimeFormatter.ofPattern("yyyy.MM.dd");

    private final JdbcTemplate jdbcTemplate;

    public ExcelToH2Service(JdbcTemplate jdbcTemplate) {
        this.jdbcTemplate = jdbcTemplate;
    }

    /**
     * Load Excel sheet into H2 table dynamically with type inference.
     *
     * @param fileName  Excel file in classpath (e.g., "data.xlsx")
     * @param sheetName Sheet name
     * @param tableName H2 table name
     */
    public void loadExcelSheet(String fileName, String sheetName, String tableName) throws Exception {
        try (InputStream is = new ClassPathResource(fileName).getInputStream();
             Workbook workbook = new XSSFWorkbook(is)) {

            Sheet sheet = workbook.getSheet(sheetName);
            if (sheet == null) throw new IllegalArgumentException("Sheet not found: " + sheetName);

            Row header = sheet.getRow(0);
            if (header == null) throw new IllegalStateException("Header row missing in Excel");

            // Column names normalized
            List<String> columNames = new ArrayList<>();
            for (Cell cell : header) {
                columNames.add(normalize(cell.getStringCellValue()));
            }

            // Type inference from first non-header row
            Row firstDataRow = sheet.getRow(1);
            if (firstDataRow == null) {
                throw new IllegalStateException("Excel sheet has no data rows");
            }

            List<String> columnTypes = new ArrayList<>();
            for (int i = 0; i < columNames.size(); i++) {
                Cell cell = firstDataRow.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                columnTypes.add(mapExcelTypeToH2Type(cell));
            }

            // Create table with inferred types
            jdbcTemplate.execute(buildCreateTableSql(tableName, columNames, columnTypes));

            // Build insert SQL
            String insertSql = buildInsertSql(tableName, columNames);

            // Insert all rows
            for (int i = 1; i <= getLastDataRow(sheet); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                Object[] values = new Object[columNames.size()];
                for (int c = 0; c < columNames.size(); c++) {
                    Cell cell = row.getCell(c, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    values[c] = readCellValue(cell);
                }

                jdbcTemplate.update(insertSql, values);
            }
        }
    }

    private int getLastDataRow(Sheet sheet) {

        for (int rowNum = sheet.getLastRowNum(); rowNum >= 0; rowNum--) {
            Row row = sheet.getRow(rowNum);
            if (row != null && rowHasData(row)) {
                return rowNum;
            }
        }
        return -1; // no data rows
    }

    private static boolean rowHasData(Row row) {
        if (row.getCell(1).getCellType() != CellType.BLANK) {
            return true;
        }
        return false;
    }

    private Object readCellValue(Cell cell) {
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> DateUtil.isCellDateFormatted(cell)
                    ? cell.getLocalDateTimeCellValue()
                    : cell.getNumericCellValue();
            case BOOLEAN -> cell.getBooleanCellValue();
            default -> null;
        };
    }

    /**
     * Infer SQL type from first row cell
     */
    private String mapExcelTypeToH2Type(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return "VARCHAR(255)";
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return "DATE";
                } else {
                    double val = cell.getNumericCellValue();
                    // If the number is an integer, use INTEGER
                    if (val % 1 == 0) {
                        return "INTEGER";
                    } else {
                        return "DECIMAL(19,4)";
                    }
                }
            case BOOLEAN:
                return "BOOLEAN";
            default:
                return "VARCHAR(255)";
        }
    }

    private String buildCreateTableSql(String tableName, List<String> columNames, List<String> types) {
        String cols = "";
        for (int i = 0; i < columNames.size(); i++) {
            String col = columNames.get(i);
            cols += col + " " + types.get(i);
            if ("id".equalsIgnoreCase(col)) cols += " PRIMARY KEY";
            if (i < columNames.size() - 1) cols += ",";
        }
        return "CREATE TABLE IF NOT EXISTS " + tableName + " (" + cols + ")";
    }

    private String buildInsertSql(String tableName, List<String> columns) {
        String placeholders = columns.stream().map(c -> "?").collect(Collectors.joining(","));
        return "INSERT INTO " + tableName + " (" + String.join(",", columns) + ") VALUES (" + placeholders + ")";
    }

    private String normalize(String colName) {
        return colName.trim().toLowerCase().replaceAll("[^a-z0-9_]", "_");
    }
}
