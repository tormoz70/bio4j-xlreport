package com.bio.xlreport.console;

import com.bio.xlreport.core.api.DataProvider;
import com.bio.xlreport.core.api.MapDataProvider;
import com.bio.xlreport.core.api.ReportSession;
import com.bio.xlreport.core.model.CompatibilityMode;
import com.bio.xlreport.core.model.PostScriptConfig;
import com.bio.xlreport.core.model.ReportConfig;
import com.bio.xlreport.core.parse.XmlReportConfigParser;
import com.bio.xlreport.js.GraalJsPostProcessor;
import com.bio.xlreport.poi.PoiReportBuilder;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Duration;
import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.NoSuchElementException;
import java.util.Locale;
import java.util.Scanner;

public final class XlReportConsoleMain {
    private static final ObjectMapper JSON = new ObjectMapper();
    private static final DateTimeFormatter DT_FMT = DateTimeFormatter.ofPattern("yyyy.MM.dd HH:mm:ss");

    private XlReportConsoleMain() {
    }

    public static void main(String[] args) throws Exception {
        if (args.length == 0) {
            printHelp();
            return;
        }
        if ("console".equalsIgnoreCase(args[0])) {
            runInteractive();
            return;
        }
        runBatch(args);
    }

    private static void runInteractive() throws Exception {
        try (Scanner scanner = new Scanner(System.in)) {
            System.out.println("XL Report Console Runner");
            System.out.print("Report XML path: ");
            String reportPath = scanner.nextLine();
            System.out.print("Data JSON path (optional): ");
            String dataPath = scanner.nextLine();
            System.out.print("Compatibility mode [strict|lenient] (default strict): ");
            String modeText = scanner.nextLine();

            String[] args = new String[] {
                "/rpt:" + reportPath,
                (dataPath == null || dataPath.isBlank()) ? "" : "/data:" + dataPath,
                "/mode:" + (modeText == null || modeText.isBlank() ? "strict" : modeText)
            };
            runBatch(args);
            System.out.println("Press ENTER to exit...");
            scanner.nextLine();
        }
    }

    private static void runBatch(String[] rawArgs) throws Exception {
        Instant t0 = Instant.now();
        ConsoleArgs args = ConsoleArgs.parse(rawArgs);
        if (args.getReportXml() == null) {
            throw new IllegalArgumentException("Report file must be provided: /rpt:<path>");
        }
        if (!Files.exists(args.getReportXml())) {
            throw new IllegalArgumentException("Report file not found: " + args.getReportXml());
        }

        Instant started = Instant.now();
        log("Report start: " + args.getReportXml().getFileName());
        log("Compatibility mode: " + args.getCompatibilityMode());
        if (!args.getReportParams().isEmpty()) {
            log("Params: " + args.getReportParams());
        }

        Instant stage = Instant.now();
        String xml = Files.readString(args.getReportXml());
        logStage("read report XML", stage);
        if (!args.getReportParams().isEmpty()) {
            xml = injectParamsIntoXml(xml, args.getReportParams());
        }

        XmlReportConfigParser parser = new XmlReportConfigParser();
        stage = Instant.now();
        ReportConfig config = parser.parse(xml, args.getReportXml(), modeOrDefault(args.getCompatibilityMode()));
        logStage("parse XML config", stage);
        Map<String, String> xmlParams = new LinkedHashMap<>(config.getParams() == null ? Map.of() : config.getParams());
        Map<String, String> runtimeParams = buildRuntimeParams(config, args);
        config = config.toBuilder().params(runtimeParams).build();
        if (args.getTemplateFile() != null) {
            config = config.toBuilder().templatePath(args.getTemplateFile()).build();
        }
        if (args.getOutputFile() != null) {
            config = config.toBuilder().outputPath(args.getOutputFile()).build();
        }
        config = alignOutputExtensionWithTemplate(config);
        if (isOracleMode(args)) {
            stage = Instant.now();
            Map<String, String> resolvedFromXml = resolveReportParamsFromXml(args, xmlParams, runtimeParams);
            logStage("resolve XML params via Oracle", stage);
            Map<String, String> placeholderParams = new LinkedHashMap<>(runtimeParams);
            // Values derived from XML params (including SQL) should override raw runtime values in template placeholders.
            placeholderParams.putAll(resolvedFromXml);
            config = config.toBuilder().params(placeholderParams).build();
        }
        if (!Files.exists(config.getTemplatePath())) {
            throw new IllegalArgumentException(
                "Template file not found: " + config.getTemplatePath() + ". " +
                    "Use /template:<path-to-xlsx> to specify explicit template."
            );
        }

        stage = Instant.now();
        DataProvider provider = createDataProvider(args, config, runtimeParams);
        logStage("create data provider", stage);
        stage = Instant.now();
        Path out = runBuildPipeline(config, provider);
        logStage("engine build (fetch + poi + post)", stage);

        Duration dur = Duration.between(started, Instant.now());
        checkPerformanceSla(args, dur, provider, config);
        log("Report done: " + out);
        log("Duration: " + formatDuration(dur));
        log("Total elapsed from args parse: " + formatDuration(Duration.between(t0, Instant.now())));

        if (args.isStopOnFinish()) {
            // In non-interactive launches (e.g., CI/cmd wrappers), stdin may be unavailable.
            if (System.console() != null) {
                System.out.println("Press ENTER to exit...");
                try (Scanner s = new Scanner(System.in)) {
                    s.nextLine();
                } catch (NoSuchElementException ignored) {
                    // Nothing to read from stdin; just exit normally.
                }
            }
        }
    }

    private static String injectParamsIntoXml(String xml, Map<String, String> params) {
        if (params.isEmpty()) {
            return xml;
        }
        if (xml.contains("<params>")) {
            // Keep existing params in XML as primary source.
            return xml;
        }
        StringBuilder block = new StringBuilder();
        block.append("<params>");
        params.forEach((k, v) -> block.append("<param name=\"").append(escapeXml(k)).append("\">").append(escapeXml(v)).append("</param>"));
        block.append("</params>");
        int insertPos = xml.indexOf("<dss>");
        if (insertPos < 0) {
            insertPos = xml.lastIndexOf("</report");
            if (insertPos < 0) {
                return xml;
            }
        }
        return xml.substring(0, insertPos) + block + xml.substring(insertPos);
    }

    private static CompatibilityMode modeOrDefault(CompatibilityMode mode) {
        return mode == null ? CompatibilityMode.STRICT : mode;
    }

    private static Map<String, String> buildRuntimeParams(ReportConfig config, ConsoleArgs args) {
        Map<String, String> cliParams = args.getReportParams();
        if (cliParams == null || cliParams.isEmpty()) {
            cliParams = new LinkedHashMap<>();
        }
        Map<String, String> merged = new LinkedHashMap<>();
        // Runtime bind parameters: CLI values and system params only.
        merged.putAll(cliParams);
        // Legacy/system runtime values used in many SQL reports.
        if (args.getUserRoles() != null && !args.getUserRoles().isBlank()) {
            merged.putIfAbsent("SYS_CURUSERROLES", args.getUserRoles());
            merged.putIfAbsent("sys_curuserroles", args.getUserRoles());
        }
        if (args.getUserOrgId() != null && !args.getUserOrgId().isBlank()) {
            merged.putIfAbsent("SYS_CURODEPUID", args.getUserOrgId());
            merged.putIfAbsent("sys_curodepuid", args.getUserOrgId());
            // Common alias in report SQLs.
            merged.putIfAbsent("org_id", args.getUserOrgId());
        }
        if (!merged.containsKey("cre_date")) {
            String fallbackDate = merged.getOrDefault("date_to", LocalDate.now().toString());
            merged.put("cre_date", fallbackDate);
        }
        return merged;
    }

    private static ReportConfig alignOutputExtensionWithTemplate(ReportConfig config) {
        String templateExt = extensionOf(config.getTemplatePath());
        String outputExt = extensionOf(config.getOutputPath());
        if ("xlsm".equals(templateExt) && !"xlsm".equals(outputExt)) {
            Path adjusted = replaceExtension(config.getOutputPath(), "xlsm");
            log("Output extension adjusted to .xlsm to match macro-enabled template: " + adjusted);
            return config.toBuilder().outputPath(adjusted).build();
        }
        return config;
    }

    private static String extensionOf(Path path) {
        if (path == null || path.getFileName() == null) {
            return "";
        }
        String name = path.getFileName().toString();
        int dot = name.lastIndexOf('.');
        if (dot < 0 || dot == name.length() - 1) {
            return "";
        }
        return name.substring(dot + 1).toLowerCase(Locale.ROOT);
    }

    private static Path replaceExtension(Path path, String extNoDot) {
        String name = path.getFileName().toString();
        int dot = name.lastIndexOf('.');
        String base = dot < 0 ? name : name.substring(0, dot);
        String newName = base + "." + extNoDot;
        Path parent = path.getParent();
        return parent == null ? Path.of(newName) : parent.resolve(newName);
    }

    private static Map<String, List<Map<String, Object>>> loadData(Path dataJson) throws Exception {
        if (dataJson == null || dataJson.toString().isBlank()) {
            return new LinkedHashMap<>();
        }
        if (!Files.exists(dataJson)) {
            throw new IllegalArgumentException("Data JSON file not found: " + dataJson);
        }
        return JSON.readValue(
            dataJson.toFile(),
            new TypeReference<Map<String, List<Map<String, Object>>>>() {
            }
        );
    }

    private static DataProvider createDataProvider(ConsoleArgs args, ReportConfig config, Map<String, String> runtimeParams) throws Exception {
        if (isOracleMode(args)) {
            DbProfiles profiles = resolveDbProfiles(args, config);
            if (profiles.byConnection().isEmpty()) {
                throw new IllegalArgumentException(
                    "Oracle mode requires /dbUrl,/dbUser,/dbPassword or /dbProfiles:<json> with at least one profile."
                );
            }
            log("Data source mode: ORACLE JDBC profiles=" + profiles.byConnection().keySet());
            return new RoutingJdbcDataProvider(profiles.defaultProfile(), profiles.byConnection());
        }
        Map<String, List<Map<String, Object>>> data = loadData(args.getDataJson());
        log("Data source mode: JSON");
        return new MapDataProvider(data);
    }

    private static Path runBuildPipeline(ReportConfig config, DataProvider provider) throws Exception {
        OracleJdbcDataProvider oracle = extractOracleProvider(provider);
        if (oracle != null) {
            executeSqlHooks(oracle, config.getPreSqlScripts(), "pre-sql");
        }
        PoiReportBuilder builder = new PoiReportBuilder();
        GraalJsPostProcessor jsPost = new GraalJsPostProcessor();
        try (ReportSession session = builder.build(config, provider)) {
            if (oracle != null) {
                executeSqlHooks(oracle, config.getPostSqlScripts(), "post-sql");
            }
            runLegacyMacroScripts(config, jsPost, session);
            for (PostScriptConfig script : config.getPostScripts()) {
                jsPost.process(config, script, session);
            }
            session.save();
            return session.outputPath();
        }
    }

    private static OracleJdbcDataProvider extractOracleProvider(DataProvider provider) {
        if (provider instanceof OracleJdbcDataProvider oracle) {
            return oracle;
        }
        if (provider instanceof RoutingJdbcDataProvider routed) {
            return routed.defaultProvider();
        }
        return null;
    }

    private static void executeSqlHooks(OracleJdbcDataProvider oracle, List<String> hooks, String stageName) throws Exception {
        if (hooks == null || hooks.isEmpty()) {
            return;
        }
        int idx = 1;
        for (String sql : hooks) {
            if (sql == null || sql.isBlank()) {
                continue;
            }
            oracle.executeSqlHook(stageName + "-" + idx, sql, 120);
            idx++;
        }
    }

    private static void runLegacyMacroScripts(
        ReportConfig config,
        GraalJsPostProcessor jsPost,
        ReportSession session
    ) throws Exception {
        runLegacyMacroScript(config.getLegacyMacroBefore(), "legacy-macro-before", config, jsPost, session);
        runLegacyMacroScript(config.getLegacyMacroAfter(), "legacy-macro-after", config, jsPost, session);
        runLegacyMacroScript(config.getLegacyAutostart(), "legacy-autostart", config, jsPost, session);
    }

    private static void runLegacyMacroScript(
        String macroName,
        String prefix,
        ReportConfig config,
        GraalJsPostProcessor jsPost,
        ReportSession session
    ) throws Exception {
        if (macroName == null || macroName.isBlank()) {
            return;
        }
        String script = """
            // Auto-migrated legacy macro hook.
            // legacy macro name: %s
            if (typeof report.applyLegacyMacro === 'function') {
              report.applyLegacyMacro('%s');
            }
            """.formatted(escapeJs(macroName), escapeJs(macroName));
        PostScriptConfig cfg = PostScriptConfig.builder()
            .name(prefix + "-" + macroName)
            .inlineScript(script)
            .timeoutMs(30_000L)
            .build();
        jsPost.process(config, cfg, session);
    }

    private static String escapeJs(String value) {
        return value.replace("\\", "\\\\").replace("'", "\\'");
    }

    private static void checkPerformanceSla(
        ConsoleArgs args,
        Duration duration,
        DataProvider provider,
        ReportConfig config
    ) throws Exception {
        if (args.getPerfSlaMs() <= 0) {
            return;
        }
        long elapsedMs = duration.toMillis();
        if (elapsedMs <= args.getPerfSlaMs()) {
            log("SLA OK: " + elapsedMs + " ms <= " + args.getPerfSlaMs() + " ms");
            return;
        }
        log("SLA VIOLATION: " + elapsedMs + " ms > " + args.getPerfSlaMs() + " ms");
        OracleJdbcDataProvider oracle = extractOracleProvider(provider);
        if (oracle == null || args.getDbaPackOut() == null) {
            return;
        }
        writeDbaPack(args.getDbaPackOut(), config, elapsedMs, args.getPerfSlaMs(), oracle.queryTracesSnapshot());
    }

    private static void writeDbaPack(
        Path outFile,
        ReportConfig config,
        long elapsedMs,
        long slaMs,
        List<OracleJdbcDataProvider.QueryTrace> traces
    ) throws Exception {
        Path parent = outFile.getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }
        StringBuilder sb = new StringBuilder();
        sb.append("# DBA Pack\n\n");
        sb.append("- report: ").append(config.getFullCode()).append("\n");
        sb.append("- elapsedMs: ").append(elapsedMs).append("\n");
        sb.append("- slaMs: ").append(slaMs).append("\n");
        sb.append("- generatedAt: ").append(LocalDateTime.now().format(DT_FMT)).append("\n\n");
        sb.append("## Queries\n\n");
        int idx = 1;
        for (var t : traces) {
            sb.append("### ").append(idx).append(". ").append(t.stage()).append("\n");
            sb.append("- elapsedMs: ").append(t.elapsedMs()).append("\n");
            sb.append("- bindCount: ").append(t.bindCount()).append("\n");
            sb.append("```sql\n").append(t.sql()).append("\n```\n\n");
            idx++;
        }
        Files.writeString(outFile, sb.toString());
        log("DBA pack written: " + outFile);
    }

    private static boolean isOracleMode(ConsoleArgs args) {
        if (args.getDbProfilesFile() != null) {
            return true;
        }
        return args.getDbUrl() != null && !args.getDbUrl().isBlank();
    }

    private static Map<String, String> resolveReportParamsFromXml(
        ConsoleArgs args,
        Map<String, String> xmlParams,
        Map<String, String> runtimeParams
    ) throws Exception {
        if (xmlParams.isEmpty()) {
            return Map.of();
        }
        OracleJdbcDataProvider resolver = resolveDbProfiles(args, null, runtimeParams).defaultProfile();
        Map<String, String> resolved = resolver.resolveReportParams(xmlParams);
        log("Resolved report params from XML definitions: " + resolved.size());
        return resolved;
    }

    private static DbProfiles resolveDbProfiles(ConsoleArgs args, ReportConfig config) throws Exception {
        return resolveDbProfiles(args, config, Map.of());
    }

    private static DbProfiles resolveDbProfiles(ConsoleArgs args, ReportConfig config, Map<String, String> runtimeParams) throws Exception {
        Map<String, OracleJdbcDataProvider> byConn = new LinkedHashMap<>();
        OracleJdbcDataProvider defaultProvider = null;

        if (args.getDbProfilesFile() != null) {
            Map<String, DbProfileSpec> specs = JSON.readValue(
                args.getDbProfilesFile().toFile(),
                new TypeReference<Map<String, DbProfileSpec>>() {
                }
            );
            for (var e : specs.entrySet()) {
                String key = normalizeConnectionName(e.getKey());
                DbProfileSpec spec = e.getValue();
                OracleJdbcDataProvider provider = new OracleJdbcDataProvider(
                    spec.dbUrl,
                    spec.dbUser,
                    spec.dbPassword,
                    spec.dbDriver == null || spec.dbDriver.isBlank() ? "oracle.jdbc.OracleDriver" : spec.dbDriver,
                    spec.dbFetchSize > 0 ? spec.dbFetchSize : 1_000,
                    args.getReportXml(),
                    runtimeParams
                );
                byConn.put(key, provider);
                if ("default".equals(key) && defaultProvider == null) {
                    defaultProvider = provider;
                }
            }
        }
        if (args.getDbUrl() != null && !args.getDbUrl().isBlank()) {
            if (args.getDbUser() == null || args.getDbUser().isBlank()) {
                throw new IllegalArgumentException("Oracle mode requires /dbUser:<user> when /dbUrl is provided.");
            }
            if (args.getDbPassword() == null) {
                throw new IllegalArgumentException("Oracle mode requires /dbPassword:<password> when /dbUrl is provided.");
            }
            OracleJdbcDataProvider cli = new OracleJdbcDataProvider(
                args.getDbUrl(),
                args.getDbUser(),
                args.getDbPassword(),
                args.getDbDriver(),
                args.getDbFetchSize(),
                args.getReportXml(),
                runtimeParams
            );
            byConn.putIfAbsent("default", cli);
            if (defaultProvider == null) {
                defaultProvider = cli;
            }
        }
        if (defaultProvider == null && !byConn.isEmpty()) {
            defaultProvider = byConn.values().iterator().next();
        }
        if (config != null && defaultProvider != null) {
            for (var ds : config.getDataSources()) {
                String key = normalizeConnectionName(ds.getConnectionName());
                byConn.putIfAbsent(key, defaultProvider);
            }
        }
        return new DbProfiles(defaultProvider, byConn);
    }

    private static String normalizeConnectionName(String value) {
        if (value == null || value.isBlank()) {
            return "default";
        }
        return value.trim().toLowerCase(Locale.ROOT);
    }

    private static String escapeXml(String s) {
        return s
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace("\"", "&quot;")
            .replace("'", "&apos;");
    }

    private static void printHelp() {
        System.out.println("XL Report Builder Console");
        System.out.println("Usage:");
        System.out.println("  gradlew :app-console:run --args=\"console\"");
        System.out.println("  gradlew :app-console:run --args=\"/rpt:<report.xml> [/template:<template.xlsx>] [/data:<data.json>] [/mode:strict|lenient] [/out:<file>] [/rptPrms:k=v;a=b] [/rptStopOnFinish:true|false]\"");
        System.out.println("  gradlew :app-console:run --args=\"/rpt:<report.xml> /dbUrl:<jdbc-url> /dbUser:<user> /dbPassword:<password> [/dbDriver:oracle.jdbc.OracleDriver] [/dbFetchSize:1000] [/dbProfiles:<profiles.json>] [/perfSlaMs:60000] [/dbaPackOut:C:/tmp/dba-pack.md] [/template:<template.xlsx>] [/mode:strict|lenient] [/out:<file>] [/rptPrms:k=v;a=b]\"");
        System.out.println();
        System.out.println("data.json format:");
        System.out.println("  { \"mRng\": [ { \"field1\": \"value\", \"field2\": 10 } ] }");
    }

    private static void log(String msg) {
        System.out.println(LocalDateTime.now().format(DT_FMT) + " - " + msg);
    }

    private static String formatDuration(Duration d) {
        long h = d.toHours();
        long m = d.toMinutesPart();
        long s = d.toSecondsPart();
        return String.format("%02d:%02d:%02d", h, m, s);
    }

    private static void logStage(String stageName, Instant startedAt) {
        Duration d = Duration.between(startedAt, Instant.now());
        log("Stage '" + stageName + "' took " + d.toMillis() + " ms");
    }

    private record DbProfileSpec(
        String dbUrl,
        String dbUser,
        String dbPassword,
        String dbDriver,
        int dbFetchSize
    ) {
    }

    private record DbProfiles(
        OracleJdbcDataProvider defaultProfile,
        Map<String, OracleJdbcDataProvider> byConnection
    ) {
    }
}
