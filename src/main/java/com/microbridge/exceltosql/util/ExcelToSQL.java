package com.microbridge.exceltosql.util;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.jdbc.core.namedparam.MapSqlParameterSource;
import org.springframework.jdbc.core.namedparam.NamedParameterJdbcTemplate;
import org.springframework.jdbc.datasource.DriverManagerDataSource;
import org.springframework.stereotype.Component;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.file.Paths;
import java.sql.Types;
import java.time.LocalDate;
import java.util.Iterator;

@Component
public class ExcelToSQL {
    private final NamedParameterJdbcTemplate namedParameterJdbcTemplate;

    public ExcelToSQL() {
        DriverManagerDataSource dataSource = new DriverManagerDataSource();
        dataSource.setUrl("jdbc:postgresql://localhost/neco_db");
        dataSource.setUsername("postgres");
        dataSource.setPassword("");
        dataSource.setDriverClassName("org.postgresql.Driver");

        namedParameterJdbcTemplate = new NamedParameterJdbcTemplate(dataSource);
    }


    public FileInputStream fileInputStream(String filePath) {

        try {
            return new FileInputStream(Paths.get(filePath).toFile());
        } catch (FileNotFoundException fnfe) {
            System.err.println("The specified file not found");
            return null;
        }
    }


    public void writeExcelToSQL(String filePath) {

        try {

            StringBuilder insertQuery = composeQuery("centre", tableColumns());

            MapSqlParameterSource parameterSource = new MapSqlParameterSource();
            Workbook workbook = new XSSFWorkbook(fileInputStream(filePath));
            int sheetCount = workbook.getNumberOfSheets();

            int sheetStart = 0;
            int rowCou = 0;
            int cellCou = 0;

            while (sheetStart < sheetCount) {

                Sheet sheet = workbook.getSheetAt(sheetStart);

                Iterator<Row> rowIterator = sheet.rowIterator();

                if (sheetStart == 0) {
                    rowIterator.next();
                }

                while (rowIterator.hasNext()) {
                    Row row = rowIterator.next();
                    Iterator<Cell> cellIterator = row.cellIterator();

                    while (cellIterator.hasNext()) {

                        Cell cell = cellIterator.next();
                        cell.setCellType(CellType.STRING);

                        if ((cell.getColumnIndex() != 1)) {

                            if (cell.getColumnIndex() == 6) {
                                parameterSource.addValue(tableColumns()[cell.getColumnIndex()], LocalDate.now(), Types.TIMESTAMP);

                            } else {
                                parameterSource.addValue(tableColumns()[cell.getColumnIndex()], cell.getStringCellValue(), columnsTypes()[cell.getColumnIndex()]);

                            }

                        }

                    }

                    namedParameterJdbcTemplate.update(insertQuery.toString(), parameterSource);
                    cellCou++;
                }

                sheetStart++;
                System.out.println(cellCou);
            }

        } catch (IOException ioe) {

        }

    }


    public StringBuilder composeQuery(String table, String... columns) {
        StringBuilder insertQuery = new StringBuilder(String.format("INSERT INTO %s (", table));

        for (int column = 0; column < columns.length; column++) {
            if (column < columns.length - 1) {
                if (column != 1)
                    insertQuery.append(String.format("%s, ", columns[column]));
            } else {
                insertQuery.append(String.format("%s", columns[column]));
            }
        }

        insertQuery.append(String.format(")VALUES("));
        for (int column = 0; column < columns.length; column++) {
            if (column < columns.length - 1) {
                if (column != 1)
                    insertQuery.append(String.format(":%s, ", columns[column]));
            } else {
                insertQuery.append(String.format(":%s", columns[column]));
            }
        }

        insertQuery.append(")");
        return insertQuery;
    }

    public String[] tableColumns() {
        String[] columns = {"state_code", "state_name", "code", "name", "country_code", "created_by", "created_at", "active", "exam_type", "claimed"};
        return columns;
    }


    public int[] columnsTypes() {
        int[] types = {Types.VARCHAR, Types.VARCHAR, Types.VARCHAR, Types.VARCHAR, Types.VARCHAR, Types.INTEGER, Types.TIMESTAMP, Types.BOOLEAN, Types.VARCHAR, Types.BOOLEAN};
        return types;
    }

}
