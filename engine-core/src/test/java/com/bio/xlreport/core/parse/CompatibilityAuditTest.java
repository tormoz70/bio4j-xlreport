package com.bio.xlreport.core.parse;

import static org.junit.jupiter.api.Assertions.assertTrue;

import com.bio.xlreport.core.model.CompatibilityMode;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.OffsetDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.Comparator;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import org.junit.jupiter.api.Assumptions;
import org.junit.jupiter.api.Test;

class CompatibilityAuditTest {

    @Test
    void auditRealXmlReportsAndWriteSummary() throws Exception {
        Path reportsDir = Path.of("c:\\data\\tmp\\rrequest-to-migrate\\ekb-cabinet\\ekb-rpt\\rpts_v2\\01_givc");
        Assumptions.assumeTrue(Files.exists(reportsDir), "Real reports directory not found.");

        var parser = new XmlReportConfigParser();
        var xmlFiles = Files.walk(reportsDir)
            .filter(p -> Files.isRegularFile(p) && p.getFileName().toString().toLowerCase(Locale.ROOT).endsWith(".xml"))
            .sorted(Comparator.naturalOrder())
            .toList();

        int parsedOk = 0;
        int parsedFailed = 0;
        int totalDs = 0;
        int withMacros = 0;
        int withSqlScripts = 0;
        int withConnections = 0;
        int withEnvironment = 0;
        int withParams = 0;
        List<String> failures = new ArrayList<>();
        List<FileAudit> fileAudits = new ArrayList<>();

        for (Path xmlFile : xmlFiles) {
            String xml = Files.readString(xmlFile);
            String low = xml.toLowerCase(Locale.ROOT);
            if (low.contains("<macrobefore") || low.contains("<macroafter")) {
                withMacros++;
            }
            if (low.contains("<sqlscriptbefore") || low.contains("<sqlscriptafter") || low.contains("<sqlscripts")) {
                withSqlScripts++;
            }
            if (low.contains("<connections")) {
                withConnections++;
            }
            if (low.contains("<environment")) {
                withEnvironment++;
            }
            if (low.contains("<params")) {
                withParams++;
            }
            totalDs += countOccurrences(low, "<ds ");

            var actions = detectMigrationActions(low);
            fileAudits.add(new FileAudit(xmlFile, actions));

            try {
                parser.parse(xml, xmlFile, CompatibilityMode.LENIENT);
                parsedOk++;
            } catch (Exception ex) {
                parsedFailed++;
                if (failures.size() < 100) {
                    failures.add(xmlFile.getFileName() + " :: " + ex.getMessage());
                }
            }
        }

        Path outDir = Path.of(System.getProperty("user.dir"))
            .resolve("build")
            .resolve("reports")
            .resolve("compat-audit");
        Files.createDirectories(outDir);
        Path outFile = outDir.resolve("01_givc-compatibility-summary.md");
        Path actionsFile = outDir.resolve("01_givc-migration-actions.md");
        Path actionsJsonFile = outDir.resolve("01_givc-migration-actions.json");

        StringBuilder md = new StringBuilder();
        md.append("# Compatibility Audit: 01_givc\n\n");
        md.append("- Generated at: ").append(OffsetDateTime.now().format(DateTimeFormatter.ISO_OFFSET_DATE_TIME)).append("\n");
        md.append("- Directory: `").append(reportsDir).append("`\n");
        md.append("- XML files found: ").append(xmlFiles.size()).append("\n");
        md.append("- Parsed successfully (LENIENT): ").append(parsedOk).append("\n");
        md.append("- Parse failures: ").append(parsedFailed).append("\n");
        md.append("- Total `<ds>` entries: ").append(totalDs).append("\n");
        md.append("- Files with macros: ").append(withMacros).append("\n");
        md.append("- Files with SQL script hooks: ").append(withSqlScripts).append("\n");
        md.append("- Files with explicit `<connections>`: ").append(withConnections).append("\n");
        md.append("- Files with `<environment>`: ").append(withEnvironment).append("\n");
        md.append("- Files with `<params>`: ").append(withParams).append("\n\n");

        md.append("## Parse Failures (up to 100)\n\n");
        if (failures.isEmpty()) {
            md.append("- none\n");
        } else {
            for (String f : failures) {
                md.append("- ").append(f).append("\n");
            }
        }

        Files.writeString(outFile, md.toString());
        Files.writeString(actionsFile, buildActionsReport(reportsDir, fileAudits));
        Files.writeString(actionsJsonFile, buildActionsJsonReport(reportsDir, xmlFiles.size(), fileAudits));
        assertTrue(parsedOk > 0, "At least one real XML should parse.");
    }

    private String buildActionsReport(Path reportsDir, List<FileAudit> fileAudits) {
        StringBuilder md = new StringBuilder();
        md.append("# Migration Actions: 01_givc\n\n");
        md.append("- Generated at: ").append(OffsetDateTime.now().format(DateTimeFormatter.ISO_OFFSET_DATE_TIME)).append("\n");
        md.append("- Directory: `").append(reportsDir).append("`\n");
        md.append("- Files analyzed: ").append(fileAudits.size()).append("\n\n");

        md.append("## Action Categories\n\n");
        md.append("- `MIGRATE_MACRO_TO_JS` (`HIGH`): перенос `macroBefore/macroAfter` в JS post-scripts\n");
        md.append("- `MIGRATE_SQL_SCRIPT_HOOKS` (`HIGH`): перенос `sqlScriptBefore/sqlScriptAfter/sqlScripts` в pre/post pipeline\n");
        md.append("- `VERIFY_CONNECTION_MAPPING` (`HIGH`): сверка connectionName/dbtype/schema и Java DataProvider\n");
        md.append("- `VERIFY_ENV_VARIABLES` (`MEDIUM`): проверка миграции `environment/variable` в env vars\n");
        md.append("- `VERIFY_PARAM_SQL` (`HIGH`): проверка SQL-типов параметров и вычисления значений\n");
        md.append("- `VERIFY_LARGE_LIMITS` (`MEDIUM`): контроль `maxRowsLimit/timeoutMinutes` и performance профиля\n");
        md.append("- `VERIFY_LEAVE_GROUP_DATA` (`MEDIUM`): проверить поддержку `leaveGroupData` в шаблонной логике\n");
        md.append("- `VERIFY_MULTI_DS_TEMPLATE` (`MEDIUM`): проверить шаблон с множеством named ranges (`<ds>`)\n");
        md.append("- `NO_SPECIAL_ACTION` (`LOW`): специфичных действий не обнаружено\n\n");

        Map<String, Integer> actionCounts = countActions(fileAudits);
        md.append("## Action Summary\n\n");
        md.append("| Action | Priority | Files |\n");
        md.append("|---|---|---:|\n");
        for (String action : orderedActions()) {
            int count = actionCounts.getOrDefault(action, 0);
            md.append("| ").append(action).append(" | ").append(priorityOf(action)).append(" | ").append(count).append(" |\n");
        }
        md.append("\n");

        md.append("## Migration Templates\n\n");
        md.append("### JS Macro Hook Template\n\n");
        md.append("```javascript\n");
        md.append("// post-script template for migrated macroAfter/macroBefore\n");
        md.append("const sheet = report.sheet(\"Sheet1\");\n");
        md.append("// Example: set audit marker after report build\n");
        md.append("sheet.cell(\"A1\").setValue(\"Migrated by JS hook\");\n");
        md.append("// Example: collapse details block\n");
        md.append("sheet.groupRows(2, 200);\n");
        md.append("```\n\n");

        md.append("### Connection Mapping Template\n\n");
        md.append("```text\n");
        md.append("legacy connectionName -> java datasource key\n");
        md.append("cub7                 -> reporting.clickhouse.main\n");
        md.append("default              -> reporting.default\n");
        md.append("\n");
        md.append("Verify SQL dialect tags (oracle/postgres/clickhouse) and schema overrides.\n");
        md.append("```\n\n");

        md.append("## Per-file Recommendations\n\n");
        for (FileAudit a : fileAudits) {
            md.append("### ").append(a.xmlFile().getFileName()).append("\n");
            if (a.actions().isEmpty()) {
                md.append("- NO_SPECIAL_ACTION\n\n");
            } else {
                for (String action : a.actions()) {
                    md.append("- [").append(priorityOf(action)).append("] ").append(action).append("\n");
                }
                md.append("\n");
            }
        }
        return md.toString();
    }

    private String buildActionsJsonReport(Path reportsDir, int filesAnalyzed, List<FileAudit> fileAudits) {
        Map<String, Integer> actionCounts = countActions(fileAudits);
        StringBuilder json = new StringBuilder();
        json.append("{\n");
        json.append("  \"generatedAt\": \"").append(jsonEscape(OffsetDateTime.now().format(DateTimeFormatter.ISO_OFFSET_DATE_TIME))).append("\",\n");
        json.append("  \"directory\": \"").append(jsonEscape(reportsDir.toString())).append("\",\n");
        json.append("  \"filesAnalyzed\": ").append(filesAnalyzed).append(",\n");
        json.append("  \"actionsSummary\": [\n");
        var ordered = orderedActions();
        for (int i = 0; i < ordered.size(); i++) {
            String action = ordered.get(i);
            json.append("    {\n");
            json.append("      \"action\": \"").append(jsonEscape(action)).append("\",\n");
            json.append("      \"priority\": \"").append(jsonEscape(priorityOf(action))).append("\",\n");
            json.append("      \"files\": ").append(actionCounts.getOrDefault(action, 0)).append("\n");
            json.append("    }");
            if (i < ordered.size() - 1) {
                json.append(",");
            }
            json.append("\n");
        }
        json.append("  ],\n");
        json.append("  \"files\": [\n");
        for (int i = 0; i < fileAudits.size(); i++) {
            FileAudit a = fileAudits.get(i);
            json.append("    {\n");
            json.append("      \"fileName\": \"").append(jsonEscape(a.xmlFile().getFileName().toString())).append("\",\n");
            json.append("      \"relativePath\": \"").append(jsonEscape(a.xmlFile().toString())).append("\",\n");
            json.append("      \"actions\": [");
            for (int k = 0; k < a.actions().size(); k++) {
                json.append("\"").append(jsonEscape(a.actions().get(k))).append("\"");
                if (k < a.actions().size() - 1) {
                    json.append(", ");
                }
            }
            json.append("],\n");
            json.append("      \"priorities\": [");
            for (int k = 0; k < a.actions().size(); k++) {
                String pr = priorityOf(a.actions().get(k));
                json.append("\"").append(jsonEscape(pr)).append("\"");
                if (k < a.actions().size() - 1) {
                    json.append(", ");
                }
            }
            json.append("]\n");
            json.append("    }");
            if (i < fileAudits.size() - 1) {
                json.append(",");
            }
            json.append("\n");
        }
        json.append("  ]\n");
        json.append("}\n");
        return json.toString();
    }

    private List<String> detectMigrationActions(String lowXml) {
        LinkedHashSet<String> actions = new LinkedHashSet<>();
        if (lowXml.contains("<macrobefore") || lowXml.contains("<macroafter")) {
            actions.add("MIGRATE_MACRO_TO_JS");
        }
        if (lowXml.contains("<sqlscriptbefore") || lowXml.contains("<sqlscriptafter") || lowXml.contains("<sqlscripts")) {
            actions.add("MIGRATE_SQL_SCRIPT_HOOKS");
        }
        if (lowXml.contains("<connections") || lowXml.contains("connectionname=")) {
            actions.add("VERIFY_CONNECTION_MAPPING");
        }
        if (lowXml.contains("<environment")) {
            actions.add("VERIFY_ENV_VARIABLES");
        }
        if (lowXml.contains("<param") && lowXml.contains("type=\"sql\"")) {
            actions.add("VERIFY_PARAM_SQL");
        }
        if (lowXml.contains("maxrowslimit=") || lowXml.contains("timeoutminutes=")) {
            actions.add("VERIFY_LARGE_LIMITS");
        }
        if (lowXml.contains("leavegroupdata=")) {
            actions.add("VERIFY_LEAVE_GROUP_DATA");
        }
        if (countOccurrences(lowXml, "<ds ") > 1) {
            actions.add("VERIFY_MULTI_DS_TEMPLATE");
        }
        return new ArrayList<>(actions);
    }

    private Map<String, Integer> countActions(List<FileAudit> fileAudits) {
        Map<String, Integer> counts = new LinkedHashMap<>();
        for (String action : orderedActions()) {
            counts.put(action, 0);
        }
        for (FileAudit a : fileAudits) {
            if (a.actions().isEmpty()) {
                counts.computeIfPresent("NO_SPECIAL_ACTION", (k, v) -> v + 1);
                continue;
            }
            for (String action : a.actions()) {
                counts.compute(action, (k, v) -> v == null ? 1 : v + 1);
            }
        }
        return counts;
    }

    private List<String> orderedActions() {
        return List.of(
            "MIGRATE_MACRO_TO_JS",
            "MIGRATE_SQL_SCRIPT_HOOKS",
            "VERIFY_CONNECTION_MAPPING",
            "VERIFY_PARAM_SQL",
            "VERIFY_ENV_VARIABLES",
            "VERIFY_LARGE_LIMITS",
            "VERIFY_LEAVE_GROUP_DATA",
            "VERIFY_MULTI_DS_TEMPLATE",
            "NO_SPECIAL_ACTION"
        );
    }

    private String priorityOf(String action) {
        return switch (action) {
            case "MIGRATE_MACRO_TO_JS", "MIGRATE_SQL_SCRIPT_HOOKS", "VERIFY_CONNECTION_MAPPING", "VERIFY_PARAM_SQL" -> "HIGH";
            case "VERIFY_ENV_VARIABLES", "VERIFY_LARGE_LIMITS", "VERIFY_LEAVE_GROUP_DATA", "VERIFY_MULTI_DS_TEMPLATE" -> "MEDIUM";
            default -> "LOW";
        };
    }

    private static String jsonEscape(String value) {
        if (value == null) {
            return "";
        }
        return value
            .replace("\\", "\\\\")
            .replace("\"", "\\\"")
            .replace("\r", "\\r")
            .replace("\n", "\\n")
            .replace("\t", "\\t");
    }

    private static int countOccurrences(String s, String token) {
        int count = 0;
        int idx = 0;
        while ((idx = s.indexOf(token, idx)) >= 0) {
            count++;
            idx += token.length();
        }
        return count;
    }

    private record FileAudit(
        Path xmlFile,
        List<String> actions
    ) {
    }
}
