package com.bio.xlreport.poi.golden.spec;

import lombok.Data;

@Data
public class CellExpectationSpec {
    private String a1Ref;
    private Double expectedNumeric;
    private String expectedString;
    private String expectedFormulaContains;
}
