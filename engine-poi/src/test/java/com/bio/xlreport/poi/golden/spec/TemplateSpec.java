package com.bio.xlreport.poi.golden.spec;

import java.util.List;
import lombok.Data;

@Data
public class TemplateSpec {
    private String sheetName;
    private List<List<String>> rows;
    private List<NamedRangeSpec> namedRanges;
}
