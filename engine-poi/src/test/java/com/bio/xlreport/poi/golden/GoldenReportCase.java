package com.bio.xlreport.poi.golden;

import com.bio.xlreport.core.model.CompatibilityMode;
import java.nio.file.Path;
import java.util.List;
import java.util.Map;
import java.util.function.Consumer;
import lombok.Builder;
import lombok.Singular;
import lombok.Value;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

@Value
@Builder
public class GoldenReportCase {
    String caseId;
    String xmlConfig;
    String mainSheetName;
    CompatibilityMode compatibilityMode;

    Consumer<XSSFWorkbook> templateInitializer;

    @Singular("dataForRange")
    Map<String, List<Map<String, Object>>> dataByRange;

    @Singular
    List<GoldenCellExpectation> cellExpectations;
    @Singular
    List<GoldenNamedRangeExpectation> namedRangeExpectations;

    Path outputFile;
}
