package com.bio.xlreport.poi.golden.spec;

import java.util.List;
import lombok.Data;

@Data
public class ExpectationsSpec {
    private List<CellExpectationSpec> cells;
    private List<NamedRangeExpectationSpec> namedRanges;
}
