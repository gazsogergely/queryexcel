package com.glab.exql.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ClassPathResource;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Service;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

@Service
public class ExcelToH2Service {

    private final JdbcTemplate jdbcTemplate;

    public ExcelToH2Service(JdbcTemplate jdbcTemplate) {
        this.jdbcTemplate = jdbcTemplate;
    }

    public void loadExcelSheet(String fileName, String sheetName, String tableName) throws Exception {
        try (InputStream is = new ClassPathResource(fileName).getInputStream();
             Workbook workbook = new XSSFWorkbook(is)) {

            Sheet sheet = workbook.getSheet(sheetName);
            if (sheet == null) throw new IllegalArgumentException("Sheet not found: " + sheetName);

            Row header = sheet.getRow(0);
            if (header == null) throw new IllegalStateException("Header row missing in Excel");

            List<String> columns = new ArrayList<>();
            for (Cell cell : header) {
                columns.add(normalize(cell.getStringCellValue()));
            }

            jdbcTemplate.execute(buildCreateTableSql(tableName, columns));
            String insertSql = buildInsertSql(tableName, columns);

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                Object[] values = new Object[columns.size()];
                for (int c = 0; c < columns.size(); c++) {
                    Cell cell = row.getCell(c, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    values[c] = readCell(cell);
                }
                jdbcTemplate.update(insertSql, values);
            }
        }
    }

    private Object readCell(Cell cell) {
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> DateUtil.isCellDateFormatted(cell)
                    ? cell.getLocalDateTimeCellValue()
                    : cell.getNumericCellValue();
            case BOOLEAN -> cell.getBooleanCellValue();
            default -> null;
        };
    }

    private String buildCreateTableSql(String tableName, List<String> cols) {
        return "CREATE TABLE IF NOT EXISTS " + tableName + " (" +
                cols.stream().map(c -> c + " VARCHAR(255)").collect(Collectors.joining(",")) +
                ")";
    }

    private String buildInsertSql(String tableName, List<String> cols) {
        String placeholders = cols.stream().map(c -> "?").collect(Collectors.joining(","));
        return "INSERT INTO " + tableName + " (" + String.join(",", cols) + ") VALUES (" + placeholders + ")";
    }

    private String normalize(String colName) {
        return colName.trim().toLowerCase().replaceAll("[^a-z0-9_]", "_");
    }
}

