package ru.mywayline.xlreport.js.api;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

class JsCellApiTest {

    @Test
    void setNumberFormatReusesCachedStyle() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet("S1");
            var row = sheet.createRow(0);
            var cell1 = row.createCell(0);
            var cell2 = row.createCell(1);
            CellStyleCache cache = new CellStyleCache();

            var api1 = new JsCellApi(cell1, cache);
            var api2 = new JsCellApi(cell2, cache);
            api1.setNumberFormat("#,##0.00");
            api2.setNumberFormat("#,##0.00");

            assertEquals(1, cache.createdCount());
            assertEquals(1, cache.hitCount());
        }
    }

    @Test
    void setNumberFormatSkipsRepeatOnSameCell() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet("S1");
            var cell = sheet.createRow(0).createCell(0);
            CellStyleCache cache = new CellStyleCache();
            var api = new JsCellApi(cell, cache);

            api.setNumberFormat("#,##0.00");
            api.setNumberFormat("#,##0.00");

            assertEquals(1, cache.createdCount());
            assertEquals(0, cache.hitCount());
        }
    }

    @Test
    void copyStyleFromSameWorkbookSharesStyle() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet("S1");
            var source = sheet.createRow(0).createCell(0);
            source.setCellValue("x");
            var target = sheet.createRow(1).createCell(0);
            CellStyleCache cache = new CellStyleCache();

            var sourceApi = new JsCellApi(source, cache);
            var targetApi = new JsCellApi(target, cache);
            targetApi.copyStyleFrom(sourceApi);

            assertEquals(source.getCellStyle().getIndex(), target.getCellStyle().getIndex());
            assertEquals(0, cache.createdCount());
        }
    }
}
