package com.bio.xlreport.poi.golden;

import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Path;
import org.junit.jupiter.api.TestFactory;
import org.junit.jupiter.api.DynamicTest;
import java.util.stream.Stream;

class GoldenCasesFromFilesTest {

    @TestFactory
    Stream<DynamicTest> runGoldenCasesFromResources() throws Exception {
        Path projectRoot = Path.of(System.getProperty("user.dir"));
        var specs = GoldenCaseRepository.loadAllFromProject(projectRoot);
        assertTrue(!specs.isEmpty(), "No golden case files found in engine-poi/src/test/resources/golden-cases");

        return specs.stream().map(spec -> DynamicTest.dynamicTest(
            "golden-file: " + spec.getCaseId(),
            () -> {
                var runtimeCase = GoldenCaseMapper.toRuntimeCase(spec);
                var out = GoldenReportHarness.run(runtimeCase);
                assertNotNull(out);
            }
        ));
    }
}
