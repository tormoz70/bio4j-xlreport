package com.bio.xlreport.poi.golden.spec;

import lombok.Data;

@Data
public class NamedRangeExpectationSpec {
    private String name;
    private String expectedRefersToContains;
}
