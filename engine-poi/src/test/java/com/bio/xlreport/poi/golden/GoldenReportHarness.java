package com.bio.xlreport.poi.golden;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.bio.xlreport.core.ReportEngine;
import com.bio.xlreport.core.api.MapDataProvider;
import com.bio.xlreport.core.model.CompatibilityMode;
import com.bio.xlreport.core.parse.XmlReportConfigParser;
import com.bio.xlreport.poi.PoiReportBuilder;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.UUID;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public final class GoldenReportHarness {

    private GoldenReportHarness() {
    }

    public static Path run(GoldenReportCase testCase) throws Exception {
        Path tempDir = Files.createTempDirectory("golden-" + testCase.getCaseId() + "-");
        Path templatePath = tempDir.resolve("template-" + UUID.randomUUID() + ".xlsx");
        Path outputPath = testCase.getOutputFile() != null ? testCase.getOutputFile() : tempDir.resolve("result.xlsx");

        createTemplate(templatePath, testCase);
        String xml = testCase.getXmlConfig()
            .replace("${TEMPLATE_PATH}", templatePath.toString().replace("\\", "/"))
            .replace("${OUTPUT_PATH}", outputPath.toString().replace("\\", "/"));

        var engine = ReportEngine.of(new XmlReportConfigParser(), new PoiReportBuilder());
        CompatibilityMode mode = testCase.getCompatibilityMode() == null ? CompatibilityMode.STRICT : testCase.getCompatibilityMode();
        engine.buildFromXml(xml, new MapDataProvider(testCase.getDataByRange()), mode);

        verify(testCase, outputPath);
        return outputPath;
    }

    private static void createTemplate(Path templatePath, GoldenReportCase testCase) throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            testCase.getTemplateInitializer().accept(wb);
            try (OutputStream out = Files.newOutputStream(templatePath)) {
                wb.write(out);
            }
        }
    }

    private static void verify(GoldenReportCase testCase, Path outputPath) throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook(Files.newInputStream(outputPath))) {
            var sheet = wb.getSheet(testCase.getMainSheetName());
            assertNotNull(sheet, "Main sheet must exist");
            var evaluator = wb.getCreationHelper().createFormulaEvaluator();

            for (var exp : testCase.getCellExpectations()) {
                var ref = new CellReference(exp.getA1Ref());
                var row = sheet.getRow(ref.getRow());
                assertNotNull(row, "Row missing for " + exp.getA1Ref());
                var cell = row.getCell(ref.getCol());
                assertNotNull(cell, "Cell missing for " + exp.getA1Ref());

                if (exp.getExpectedString() != null) {
                    assertEquals(exp.getExpectedString(), cell.getStringCellValue(), exp.getA1Ref());
                }
                if (exp.getExpectedNumeric() != null) {
                    double value = (cell.getCellType() == CellType.FORMULA)
                        ? evaluator.evaluate(cell).getNumberValue()
                        : cell.getNumericCellValue();
                    assertEquals(exp.getExpectedNumeric(), value, 0.0001, exp.getA1Ref());
                }
                if (exp.getExpectedFormulaContains() != null) {
                    assertEquals(CellType.FORMULA, cell.getCellType(), "Expected formula at " + exp.getA1Ref());
                    assertTrue(
                        cell.getCellFormula().replace("$", "").contains(exp.getExpectedFormulaContains()),
                        "Formula mismatch at " + exp.getA1Ref() + ": " + cell.getCellFormula()
                    );
                }
            }

            for (var nr : testCase.getNamedRangeExpectations()) {
                var name = wb.getName(nr.getName());
                assertNotNull(name, "Named range not found: " + nr.getName());
                assertTrue(
                    name.getRefersToFormula().replace("$", "").contains(nr.getExpectedRefersToContains()),
                    "Named range mismatch: " + nr.getName() + " -> " + name.getRefersToFormula()
                );
            }
        }
    }
}
