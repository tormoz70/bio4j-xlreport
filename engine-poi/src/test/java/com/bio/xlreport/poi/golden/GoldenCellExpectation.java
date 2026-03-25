package com.bio.xlreport.poi.golden;

import lombok.Builder;
import lombok.Value;

@Value
@Builder
public class GoldenCellExpectation {
    String a1Ref;
    Double expectedNumeric;
    String expectedString;
    String expectedFormulaContains;
}
