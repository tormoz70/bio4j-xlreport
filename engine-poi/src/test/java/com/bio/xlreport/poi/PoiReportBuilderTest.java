package com.bio.xlreport.poi;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.bio.xlreport.core.api.MapDataProvider;
import com.bio.xlreport.core.model.CompatibilityMode;
import com.bio.xlreport.core.model.DataSourceConfig;
import com.bio.xlreport.core.model.ExtAttributes;
import com.bio.xlreport.core.model.FieldDef;
import com.bio.xlreport.core.model.ReportConfig;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

class PoiReportBuilderTest {

    @Test
    void buildFromNamedRange() throws Exception {
        Path tempDir = Files.createTempDirectory("poi-builder-test");
        Path template = tempDir.resolve("template.xlsx");
        Path result = tempDir.resolve("result.xlsx");

        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            var sheet = wb.createSheet(WorkbookUtil.createSafeSheetName("Sheet1"));
            var titleRow = sheet.createRow(0);
            titleRow.createCell(0).setCellValue("#REPORT_TITLE#");
            var row = sheet.createRow(1);
            row.createCell(0).setCellValue("details");
            row.createCell(1).setCellValue("=sales_A");
            var name = wb.createName();
            name.setNameName("Sales");
            name.setRefersToFormula("'Sheet1'!A2:B2");
            try (OutputStream out = Files.newOutputStream(template)) {
                wb.write(out);
            }
        }

        var cfg = ReportConfig.builder()
            .uid("u1")
            .fullCode("demo")
            .title("DemoTitle")
            .subject("s")
            .author("a")
            .templatePath(template)
            .outputPath(result)
            .compatibilityMode(CompatibilityMode.STRICT)
            .extAttributes(ExtAttributes.builder().targetFormat("msexcel").build())
            .dataSource(
                DataSourceConfig.builder()
                    .rangeName("Sales")
                    .title("Sales")
                    .timeoutMinutes(1)
                    .maxRowsLimit(1000)
                    .field(FieldDef.builder().name("A").build())
                    .field(FieldDef.builder().name("B").build())
                    .build()
            )
            .build();

        var provider = new MapDataProvider(
            Map.of(
                "Sales",
                List.of(
                    Map.of("A", "R1", "B", "X1"),
                    Map.of("A", "R2", "B", "X2")
                )
            )
        );

        var builder = new PoiReportBuilder();
        try (var session = builder.build(cfg, provider)) {
            session.save();
        }

        assertTrue(Files.exists(result));
        try (XSSFWorkbook wb = new XSSFWorkbook(Files.newInputStream(result))) {
            var s = wb.getSheet("Sheet1");
            assertEquals("DemoTitle", s.getRow(0).getCell(0).getStringCellValue());
            assertEquals("R2", s.getRow(2).getCell(1).getStringCellValue());
        }
    }

    @Test
    void buildWithGroupMarkersAndTotals() throws Exception {
        Path tempDir = Files.createTempDirectory("poi-builder-group-test");
        Path template = tempDir.resolve("template-group.xlsx");
        Path result = tempDir.resolve("result-group.xlsx");

        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            var sheet = wb.createSheet("Sheet1");
            sheet.createRow(0).createCell(0).setCellValue("groupH=sales_dept");
            var details = sheet.createRow(1);
            details.createCell(0).setCellValue("details");
            details.createCell(1).setCellValue("=sales_amount");
            sheet.createRow(2).createCell(0).setCellValue("groupF=sales_dept");
            var totals = sheet.createRow(3);
            totals.createCell(0).setCellValue("totals");
            totals.createCell(1).setCellValue("sum(sales_amount)");

            var name = wb.createName();
            name.setNameName("Sales");
            name.setRefersToFormula("'Sheet1'!A1:B4");
            try (OutputStream out = Files.newOutputStream(template)) {
                wb.write(out);
            }
        }

        var cfg = ReportConfig.builder()
            .uid("u2")
            .fullCode("demo.group")
            .title("Demo")
            .subject("s")
            .author("a")
            .templatePath(template)
            .outputPath(result)
            .compatibilityMode(CompatibilityMode.STRICT)
            .extAttributes(ExtAttributes.builder().targetFormat("msexcel").build())
            .dataSource(
                DataSourceConfig.builder()
                    .rangeName("Sales")
                    .field(FieldDef.builder().name("dept").build())
                    .field(FieldDef.builder().name("amount").build())
                    .build()
            )
            .build();

        var provider = new MapDataProvider(
            Map.of(
                "Sales",
                List.of(
                    Map.of("dept", "A", "amount", 10),
                    Map.of("dept", "A", "amount", 20),
                    Map.of("dept", "B", "amount", 7)
                )
            )
        );

        var builder = new PoiReportBuilder();
        try (var session = builder.build(cfg, provider)) {
            session.save();
        }

        try (XSSFWorkbook wb = new XSSFWorkbook(Files.newInputStream(result))) {
            var s = wb.getSheet("Sheet1");
            var evaluator = wb.getCreationHelper().createFormulaEvaluator();
            assertEquals("A", s.getRow(0).getCell(0).getStringCellValue());
            assertEquals(30d, evaluator.evaluate(s.getRow(3).getCell(1)).getNumberValue(), 0.0001);
            assertEquals(7d, evaluator.evaluate(s.getRow(6).getCell(1)).getNumberValue(), 0.0001);
            assertEquals(37d, evaluator.evaluate(s.getRow(7).getCell(1)).getNumberValue(), 0.0001);
            // One grouped block should exist between first header and first footer.
            assertTrue(s.getRow(1).getOutlineLevel() > 0);
        }
    }

    @Test
    void suppressesGroupedDetailFieldWithUnderscoreName() throws Exception {
        Path tempDir = Files.createTempDirectory("poi-builder-group-underscore");
        Path template = tempDir.resolve("template-group-underscore.xlsx");
        Path result = tempDir.resolve("result-group-underscore.xlsx");

        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            var sheet = wb.createSheet("Sheet1");
            sheet.createRow(0).createCell(0).setCellValue("groupH=sales_region_name");
            var details = sheet.createRow(1);
            details.createCell(0).setCellValue("details");
            details.createCell(1).setCellValue("=sales_region_name");
            details.createCell(2).setCellValue("=sales_amount");
            sheet.createRow(2).createCell(0).setCellValue("groupF=sales_region_name");

            var name = wb.createName();
            name.setNameName("Sales");
            name.setRefersToFormula("'Sheet1'!A1:C3");
            try (OutputStream out = Files.newOutputStream(template)) {
                wb.write(out);
            }
        }

        var cfg = ReportConfig.builder()
            .uid("u7")
            .fullCode("demo.group.underscore")
            .title("Demo")
            .subject("s")
            .author("a")
            .templatePath(template)
            .outputPath(result)
            .compatibilityMode(CompatibilityMode.STRICT)
            .extAttributes(ExtAttributes.builder().targetFormat("msexcel").build())
            .dataSource(
                DataSourceConfig.builder()
                    .rangeName("Sales")
                    .field(FieldDef.builder().name("region_name").build())
                    .field(FieldDef.builder().name("amount").build())
                    .build()
            )
            .build();

        var provider = new MapDataProvider(
            Map.of(
                "Sales",
                List.of(
                    Map.of("region_name", "R1", "amount", 10),
                    Map.of("region_name", "R1", "amount", 20)
                )
            )
        );

        var builder = new PoiReportBuilder();
        try (var session = builder.build(cfg, provider)) {
            session.save();
        }

        try (XSSFWorkbook wb = new XSSFWorkbook(Files.newInputStream(result))) {
            var s = wb.getSheet("Sheet1");
            assertEquals("R1", s.getRow(0).getCell(1).getStringCellValue());
            assertEquals(CellType.BLANK, s.getRow(1).getCell(1).getCellType());
            assertEquals(CellType.BLANK, s.getRow(2).getCell(1).getCellType());
            assertEquals(10d, s.getRow(1).getCell(2).getNumericCellValue(), 0.0001);
            assertEquals(20d, s.getRow(2).getCell(2).getNumericCellValue(), 0.0001);
        }
    }

    @Test
    void runTotalUsesChildTotalRowsInRootFormula() throws Exception {
        Path tempDir = Files.createTempDirectory("poi-builder-runtotal");
        Path template = tempDir.resolve("template-runtotal.xlsx");
        Path result = tempDir.resolve("result-runtotal.xlsx");

        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            var sheet = wb.createSheet("Sheet1");
            sheet.createRow(0).createCell(0).setCellValue("groupH=sales_dept");
            var details = sheet.createRow(1);
            details.createCell(0).setCellValue("details");
            details.createCell(1).setCellValue("=sales_amount");
            sheet.createRow(2).createCell(0).setCellValue("groupF=sales_dept");
            var totals = sheet.createRow(3);
            totals.createCell(0).setCellValue("totals");
            totals.createCell(1).setCellValue("sum(sum)");

            var name = wb.createName();
            name.setNameName("Sales");
            name.setRefersToFormula("'Sheet1'!A1:B4");
            try (OutputStream out = Files.newOutputStream(template)) {
                wb.write(out);
            }
        }

        var cfg = ReportConfig.builder()
            .uid("u5")
            .fullCode("demo.runtotal")
            .title("Demo")
            .subject("s")
            .author("a")
            .templatePath(template)
            .outputPath(result)
            .compatibilityMode(CompatibilityMode.STRICT)
            .extAttributes(ExtAttributes.builder().targetFormat("msexcel").build())
            .dataSource(
                DataSourceConfig.builder()
                    .rangeName("Sales")
                    .field(FieldDef.builder().name("dept").build())
                    .field(FieldDef.builder().name("amount").build())
                    .build()
            )
            .build();

        var provider = new MapDataProvider(
            Map.of(
                "Sales",
                List.of(
                    Map.of("dept", "A", "amount", 10),
                    Map.of("dept", "A", "amount", 20),
                    Map.of("dept", "B", "amount", 7)
                )
            )
        );

        var builder = new PoiReportBuilder();
        try (var session = builder.build(cfg, provider)) {
            session.save();
        }

        try (XSSFWorkbook wb = new XSSFWorkbook(Files.newInputStream(result))) {
            var s = wb.getSheet("Sheet1");
            String rootFormula = s.getRow(7).getCell(1).getCellFormula().replace("$", "");
            // Root runtotal must aggregate group footer rows, not raw detail ranges.
            assertTrue(rootFormula.contains("B4"), "formula: " + rootFormula);
            assertTrue(rootFormula.contains("B7"), "formula: " + rootFormula);
        }
    }

    @Test
    void countTotalsUseRowsAndChildTotalSum() throws Exception {
        Path tempDir = Files.createTempDirectory("poi-builder-count-total");
        Path template = tempDir.resolve("template-count-total.xlsx");
        Path result = tempDir.resolve("result-count-total.xlsx");

        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            var sheet = wb.createSheet("Sheet1");
            sheet.createRow(0).createCell(0).setCellValue("groupH=sales_dept");
            var details = sheet.createRow(1);
            details.createCell(0).setCellValue("details");
            details.createCell(1).setCellValue("=sales_amount");
            sheet.createRow(2).createCell(0).setCellValue("groupF=sales_dept");
            var totals = sheet.createRow(3);
            totals.createCell(0).setCellValue("totals");
            totals.createCell(1).setCellValue("cnt");

            var name = wb.createName();
            name.setNameName("Sales");
            name.setRefersToFormula("'Sheet1'!A1:B4");
            try (OutputStream out = Files.newOutputStream(template)) {
                wb.write(out);
            }
        }

        var cfg = ReportConfig.builder()
            .uid("u6")
            .fullCode("demo.count")
            .title("Demo")
            .subject("s")
            .author("a")
            .templatePath(template)
            .outputPath(result)
            .compatibilityMode(CompatibilityMode.STRICT)
            .extAttributes(ExtAttributes.builder().targetFormat("msexcel").build())
            .dataSource(
                DataSourceConfig.builder()
                    .rangeName("Sales")
                    .field(FieldDef.builder().name("dept").build())
                    .field(FieldDef.builder().name("amount").build())
                    .build()
            )
            .build();

        var provider = new MapDataProvider(
            Map.of(
                "Sales",
                List.of(
                    Map.of("dept", "A", "amount", 10),
                    Map.of("dept", "A", "amount", 20),
                    Map.of("dept", "B", "amount", 7)
                )
            )
        );

        var builder = new PoiReportBuilder();
        try (var session = builder.build(cfg, provider)) {
            session.save();
        }

        try (XSSFWorkbook wb = new XSSFWorkbook(Files.newInputStream(result))) {
            var s = wb.getSheet("Sheet1");
            var evaluator = wb.getCreationHelper().createFormulaEvaluator();
            // Group totals count details in each group.
            assertEquals(2d, evaluator.evaluate(s.getRow(3).getCell(1)).getNumberValue(), 0.0001);
            assertEquals(1d, evaluator.evaluate(s.getRow(6).getCell(1)).getNumberValue(), 0.0001);
            // Root total counts via child totals sum.
            String rootFormula = s.getRow(7).getCell(1).getCellFormula().replace("$", "");
            assertTrue(rootFormula.contains("SUM("), "formula: " + rootFormula);
            assertEquals(3d, evaluator.evaluate(s.getRow(7).getCell(1)).getNumberValue(), 0.0001);
        }
    }

    @Test
    void buildSingleRowReplacesNamedPlaceholders() throws Exception {
        Path tempDir = Files.createTempDirectory("poi-builder-single-row");
        Path template = tempDir.resolve("template-single.xlsx");
        Path result = tempDir.resolve("result-single.xlsx");

        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            var sheet = wb.createSheet("Sheet1");
            sheet.createRow(0).createCell(0).setCellValue("Value: customer_name");
            sheet.createRow(1).createCell(0).setCellValue("(customer_status)");
            try (OutputStream out = Files.newOutputStream(template)) {
                wb.write(out);
            }
        }

        var cfg = ReportConfig.builder()
            .uid("u3")
            .fullCode("demo.single")
            .title("Demo")
            .subject("s")
            .author("a")
            .templatePath(template)
            .outputPath(result)
            .compatibilityMode(CompatibilityMode.STRICT)
            .extAttributes(ExtAttributes.builder().targetFormat("msexcel").build())
            .dataSource(
                DataSourceConfig.builder()
                    .rangeName("customer")
                    .singleRow(true)
                    .build()
            )
            .build();

        var provider = new MapDataProvider(
            Map.of(
                "customer",
                List.of(Map.of("name", "Alice", "status", "VIP"))
            )
        );

        var builder = new PoiReportBuilder();
        try (var session = builder.build(cfg, provider)) {
            session.save();
        }

        try (XSSFWorkbook wb = new XSSFWorkbook(Files.newInputStream(result))) {
            var s = wb.getSheet("Sheet1");
            assertEquals("Value: Alice", s.getRow(0).getCell(0).getStringCellValue());
            assertEquals("VIP", s.getRow(1).getCell(0).getStringCellValue());
        }
    }

    @Test
    void lenientModeAllowsTemplateWithoutDetailsMarker() throws Exception {
        Path tempDir = Files.createTempDirectory("poi-builder-lenient");
        Path template = tempDir.resolve("template-lenient.xlsx");
        Path result = tempDir.resolve("result-lenient.xlsx");

        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            var sheet = wb.createSheet("Sheet1");
            var row = sheet.createRow(0);
            row.createCell(0).setCellValue("=sales_A");
            var name = wb.createName();
            name.setNameName("Sales");
            name.setRefersToFormula("'Sheet1'!A1:A1");
            try (OutputStream out = Files.newOutputStream(template)) {
                wb.write(out);
            }
        }

        var cfg = ReportConfig.builder()
            .uid("u4")
            .fullCode("demo.lenient")
            .title("Demo")
            .subject("s")
            .author("a")
            .templatePath(template)
            .outputPath(result)
            .compatibilityMode(CompatibilityMode.LENIENT)
            .extAttributes(ExtAttributes.builder().targetFormat("msexcel").build())
            .dataSource(
                DataSourceConfig.builder()
                    .rangeName("Sales")
                    .field(FieldDef.builder().name("A").build())
                    .build()
            )
            .build();

        var provider = new MapDataProvider(
            Map.of("Sales", List.of(Map.of("A", "R1")))
        );

        var builder = new PoiReportBuilder();
        try (var session = builder.build(cfg, provider)) {
            session.save();
        }

        try (XSSFWorkbook wb = new XSSFWorkbook(Files.newInputStream(result))) {
            assertEquals("R1", wb.getSheet("Sheet1").getRow(0).getCell(0).getStringCellValue());
        }
    }
}
