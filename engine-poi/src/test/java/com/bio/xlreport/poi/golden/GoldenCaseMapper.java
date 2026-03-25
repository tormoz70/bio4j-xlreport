package com.bio.xlreport.poi.golden;

import com.bio.xlreport.core.model.CompatibilityMode;
import com.bio.xlreport.poi.golden.spec.GoldenCaseSpec;

public final class GoldenCaseMapper {
    private GoldenCaseMapper() {
    }

    public static GoldenReportCase toRuntimeCase(GoldenCaseSpec spec) {
        var builder = GoldenReportCase.builder()
            .caseId(spec.getCaseId())
            .xmlConfig(spec.getXmlConfig())
            .mainSheetName(spec.getMainSheetName())
            .compatibilityMode(parseMode(spec.getCompatibilityMode()))
            .templateInitializer(wb -> {
                var template = spec.getTemplate();
                var sheet = wb.createSheet(template.getSheetName());
                for (int r = 0; r < template.getRows().size(); r++) {
                    var row = sheet.createRow(r);
                    var cells = template.getRows().get(r);
                    for (int c = 0; c < cells.size(); c++) {
                        row.createCell(c).setCellValue(cells.get(c));
                    }
                }
                if (template.getNamedRanges() != null) {
                    for (var nr : template.getNamedRanges()) {
                        var name = wb.createName();
                        name.setNameName(nr.getName());
                        name.setRefersToFormula(nr.getRefersToFormula());
                    }
                }
            });

        if (spec.getDataByRange() != null) {
            spec.getDataByRange().forEach(builder::dataForRange);
        }
        if (spec.getExpectations() != null) {
            if (spec.getExpectations().getCells() != null) {
                spec.getExpectations().getCells().forEach(cell -> builder.cellExpectation(
                    GoldenCellExpectation.builder()
                        .a1Ref(cell.getA1Ref())
                        .expectedNumeric(cell.getExpectedNumeric())
                        .expectedString(cell.getExpectedString())
                        .expectedFormulaContains(cell.getExpectedFormulaContains())
                        .build()
                ));
            }
            if (spec.getExpectations().getNamedRanges() != null) {
                spec.getExpectations().getNamedRanges().forEach(nr -> builder.namedRangeExpectation(
                    GoldenNamedRangeExpectation.builder()
                        .name(nr.getName())
                        .expectedRefersToContains(nr.getExpectedRefersToContains())
                        .build()
                ));
            }
        }
        return builder.build();
    }

    private static CompatibilityMode parseMode(String value) {
        if (value == null || value.isBlank()) {
            return CompatibilityMode.STRICT;
        }
        if ("lenient".equalsIgnoreCase(value)) {
            return CompatibilityMode.LENIENT;
        }
        return CompatibilityMode.STRICT;
    }
}
