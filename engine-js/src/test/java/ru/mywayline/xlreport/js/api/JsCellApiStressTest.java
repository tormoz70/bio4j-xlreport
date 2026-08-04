package ru.mywayline.xlreport.js.api;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

class JsCellApiStressTest {

    @Test
    void formatsManyCellsWithoutExceedingStyleBudget() throws Exception {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet("S1");
            CellStyleCache cache = new CellStyleCache();
            var cell = sheet.createRow(0).createCell(0);
            var api = new JsCellApi(cell, cache);

            for (int i = 0; i < 1000; i++) {
                api.setNumberFormat("#,##0.00");
            }

            assertEquals(1, cache.createdCount());
            assertEquals(0, cache.hitCount());
        }
    }
}
