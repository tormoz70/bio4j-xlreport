package com.bio.xlreport.poi.golden;

import static org.junit.jupiter.api.Assertions.assertNotNull;

import java.nio.file.Path;
import java.util.List;
import java.util.Map;
import org.junit.jupiter.api.Test;

class GoldenRegressionTest {

    @Test
    void golden_groupedTotalsAndNamedRange() throws Exception {
        String xml = """
            <reportDef full_code="golden.grouped" liveScripts="false">
              <adv_template>${TEMPLATE_PATH}</adv_template>
              <filename_fmt>${OUTPUT_PATH}</filename_fmt>
              <title>Golden report</title>
              <subject>regression</subject>
              <autor>test</autor>
              <dss>
                <ds range="Sales">
                  <fields>
                    <field name="dept" />
                    <field name="amount" />
                  </fields>
                </ds>
              </dss>
            </reportDef>
            """;

        GoldenReportCase testCase = GoldenReportCase.builder()
            .caseId("grouped-totals")
            .xmlConfig(xml)
            .mainSheetName("Sheet1")
            .templateInitializer(wb -> {
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
            })
            .dataForRange(
                "Sales",
                List.of(
                    Map.of("dept", "A", "amount", 10),
                    Map.of("dept", "A", "amount", 20),
                    Map.of("dept", "B", "amount", 7)
                )
            )
            .cellExpectation(GoldenCellExpectation.builder().a1Ref("A1").expectedString("A").build())
            .cellExpectation(GoldenCellExpectation.builder().a1Ref("B4").expectedNumeric(30d).build())
            .cellExpectation(GoldenCellExpectation.builder().a1Ref("B7").expectedNumeric(7d).build())
            .cellExpectation(GoldenCellExpectation.builder().a1Ref("B8").expectedNumeric(37d).expectedFormulaContains("B4").build())
            .namedRangeExpectation(
                GoldenNamedRangeExpectation.builder().name("Sales").expectedRefersToContains("A1:B8").build()
            )
            .build();

        Path result = GoldenReportHarness.run(testCase);
        assertNotNull(result);
    }
}
