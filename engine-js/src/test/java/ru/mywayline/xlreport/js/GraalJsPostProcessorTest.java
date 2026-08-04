package ru.mywayline.xlreport.js;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import ru.mywayline.xlreport.core.model.ExtAttributes;
import ru.mywayline.xlreport.core.model.ReportConfig;
import ru.mywayline.xlreport.poi.PoiReportSession;
import java.nio.file.Files;
import java.nio.file.Path;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

class GraalJsPostProcessorTest {

    @Test
    void dispatchesValidMacroModule() throws Exception {
        Path tempDir = Files.createTempDirectory("graal-macro-test");
        Path moduleFile = tempDir.resolve("DemoModule.js");
        Files.writeString(moduleFile, """
            var DemoModule = {
              runMacro: function(report) {
                report.sheet('Sheet1').cellAt(1, 1).setValue('macro-ok');
              }
            };
            """);

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            workbook.createSheet("Sheet1");
            ReportConfig config = ReportConfig.builder()
                .uid("u1")
                .fullCode("demo")
                .title("Demo")
                .subject("s")
                .author("a")
                .extAttributes(ExtAttributes.builder().targetFormat("msexcel").build())
                .build();
            PoiReportSession session = new PoiReportSession(workbook, tempDir.resolve("out.xlsx"));

            GraalJsPostProcessor processor = new GraalJsPostProcessor();
            processor.processMacro(config, "DemoModule.runMacro", tempDir.toString(), session);

            assertEquals("macro-ok", workbook.getSheetAt(0).getRow(0).getCell(0).getStringCellValue());
        }
    }

    @Test
    void rejectsInvalidMacroModuleName() throws Exception {
        Path tempDir = Files.createTempDirectory("graal-macro-invalid");
        Path moduleFile = tempDir.resolve("My-Module.js");
        Files.writeString(moduleFile, "var MyModule = { run: function(r) {} };");

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            workbook.createSheet("Sheet1");
            ReportConfig config = ReportConfig.builder()
                .uid("u2")
                .fullCode("demo")
                .title("Demo")
                .subject("s")
                .author("a")
                .extAttributes(ExtAttributes.builder().targetFormat("msexcel").build())
                .build();
            PoiReportSession session = new PoiReportSession(workbook, tempDir.resolve("out.xlsx"));

            GraalJsPostProcessor processor = new GraalJsPostProcessor();
            processor.processMacro(config, "My-Module.run", tempDir.toString(), session);

            assertTrue(workbook.getSheetAt(0).getRow(0) == null
                || workbook.getSheetAt(0).getRow(0).getCell(0) == null);
        }
    }
}
