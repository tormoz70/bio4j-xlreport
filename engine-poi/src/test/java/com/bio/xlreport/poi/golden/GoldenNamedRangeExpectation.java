package com.bio.xlreport.poi.golden;

import lombok.Builder;
import lombok.Value;

@Value
@Builder
public class GoldenNamedRangeExpectation {
    String name;
    String expectedRefersToContains;
}
