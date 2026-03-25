package com.bio.xlreport.poi.golden.spec;

import java.util.List;
import java.util.Map;
import lombok.Data;

@Data
public class GoldenCaseSpec {
    private String caseId;
    private String xmlConfig;
    private String mainSheetName;
    private String compatibilityMode;
    private TemplateSpec template;
    private Map<String, List<Map<String, Object>>> dataByRange;
    private ExpectationsSpec expectations;
}
