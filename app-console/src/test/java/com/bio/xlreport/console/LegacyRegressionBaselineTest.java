package com.bio.xlreport.console;

import static org.junit.jupiter.api.Assertions.assertTrue;

import com.fasterxml.jackson.databind.ObjectMapper;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Duration;
import java.time.Instant;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import org.junit.jupiter.api.Test;

class LegacyRegressionBaselineTest {
    private static final ObjectMapper JSON = new ObjectMapper();

    @Test
    void generateMachineReadableBaselineDiffReport() throws Exception {
        Path projectRoot = Path.of(System.getProperty("user.dir")).toAbsolutePath().normalize();
        Path examplesDir = projectRoot.resolve("examples");
        if (!Files.exists(examplesDir)) {
            examplesDir = projectRoot.getParent().resolve("examples");
            projectRoot = projectRoot.getParent().toAbsolutePath().normalize();
        }
        assertTrue(Files.exists(examplesDir), "Examples dir not found: " + examplesDir);

        List<Path> baselineFiles = Files.list(examplesDir)
            .filter(p -> p.getFileName().toString().startsWith("legacy-baseline-") && p.getFileName().toString().endsWith(".json"))
            .sorted(Comparator.comparing(p -> p.getFileName().toString()))
            .toList();
        assertTrue(!baselineFiles.isEmpty(), "No baseline files found in: " + examplesDir);

        Path outDir = projectRoot.resolve("app-console").resolve("build").resolve("reports").resolve("legacy-regression");
        Files.createDirectories(outDir);

        for (Path baselinePath : baselineFiles) {
            BaselineSpec baseline = JSON.readValue(baselinePath.toFile(), BaselineSpec.class);
            runSingleBaseline(projectRoot, baseline, outDir);
        }
    }

    private static void runSingleBaseline(Path projectRoot, BaselineSpec baseline, Path outDir) throws Exception {
        boolean enabled = "true".equalsIgnoreCase(System.getenv(defaultString(baseline.enabledEnv, "ORACLE_REGRESSION_ENABLED")));
        String provider = defaultString(baseline.executionProvider, "oracle").toLowerCase();
        List<Map<String, Object>> rows = new ArrayList<>();
        for (String rel : baseline.reports) {
            Map<String, Object> one = new LinkedHashMap<>();
            one.put("report", rel);
            one.put("criteria", baseline.criteria);
            if (!enabled) {
                one.put("status", "SKIPPED");
                one.put("reason", "Set " + defaultString(baseline.enabledEnv, "ORACLE_REGRESSION_ENABLED") + "=true to run regression.");
                rows.add(one);
                continue;
            }
            if (!"oracle".equals(provider)) {
                one.put("status", "SKIPPED");
                one.put("reason", "Execution provider '" + provider + "' is not implemented in this harness.");
                rows.add(one);
                continue;
            }

            Path rpt = Path.of(baseline.sourceRoot).resolve(rel);
            Path out = projectRoot.resolve("out").resolve(rpt.getFileName().toString().replace(".xml", "-baseline.xlsx"));
            Instant started = Instant.now();
            try {
                XlReportConsoleMain.main(
                    new String[] {
                        "/rpt:" + rpt,
                        "/mode:lenient",
                        "/out:" + out,
                        "/dbUrl:" + envOrThrow("ORACLE_DB_URL"),
                        "/dbUser:" + envOrThrow("ORACLE_DB_USER"),
                        "/dbPassword:" + envOrThrow("ORACLE_DB_PASSWORD"),
                        "/dbDriver:" + envOrDefault("ORACLE_DB_DRIVER", "oracle.jdbc.OracleDriver"),
                        "/dbFetchSize:" + envOrDefault("ORACLE_DB_FETCH_SIZE", "1000"),
                        "/rptStopOnFinish:false"
                    }
                );
                long ms = Duration.between(started, Instant.now()).toMillis();
                one.put("status", "OK");
                one.put("durationMs", ms);
                one.put("slaMs", baseline.sla.getOrDefault(rpt.getFileName().toString(), baseline.sla.getOrDefault("default", 0L)));
                one.put("output", out.toString());
            } catch (Exception ex) {
                one.put("status", "FAILED");
                one.put("error", ex.getMessage());
                one.put("failureType", classifyFailure(ex.getMessage()));
            }
            rows.add(one);
        }
        Path diff = outDir.resolve(baseline.baselineId + "-diff.json");

        Map<String, Object> payload = new LinkedHashMap<>();
        payload.put("baselineId", baseline.baselineId);
        payload.put("sourceRoot", baseline.sourceRoot);
        payload.put("enabledEnv", defaultString(baseline.enabledEnv, "ORACLE_REGRESSION_ENABLED"));
        payload.put("executionProvider", provider);
        payload.put("regressionEnabled", enabled);
        payload.put("results", rows);
        JSON.writerWithDefaultPrettyPrinter().writeValue(diff.toFile(), payload);

        assertTrue(Files.exists(diff), "Diff report not generated: " + diff);
    }

    private static String envOrThrow(String key) {
        String value = System.getenv(key);
        if (value == null || value.isBlank()) {
            throw new IllegalArgumentException("Missing env variable: " + key);
        }
        return value;
    }

    private static String envOrDefault(String key, String def) {
        String value = System.getenv(key);
        return value == null || value.isBlank() ? def : value;
    }

    private static String classifyFailure(String msg) {
        if (msg == null || msg.isBlank()) {
            return "UNKNOWN";
        }
        String low = msg.toLowerCase();
        if (low.contains("parseDateTimeBestEffort".toLowerCase()) || low.contains("formatDateTime".toLowerCase())) {
            return "INCOMPATIBLE_SQL_DIALECT";
        }
        if (low.contains("file not found")) {
            return "MISSING_REPORT_FILE";
        }
        if (low.contains("ora-") || low.contains("sql")) {
            return "SQL_EXECUTION_ERROR";
        }
        return "RUNTIME_ERROR";
    }

    private static String defaultString(String value, String def) {
        if (value == null || value.isBlank()) {
            return def;
        }
        return value;
    }

    @SuppressWarnings("unused")
    private static class BaselineSpec {
        public String baselineId;
        public String sourceRoot;
        public String enabledEnv;
        public String executionProvider;
        public Map<String, Object> criteria;
        public Map<String, Long> sla;
        public List<String> reports;
    }
}
