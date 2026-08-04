package ru.mywayline.xlreport.js;

import ru.mywayline.xlreport.core.api.ReportBuildStats;
import ru.mywayline.xlreport.core.api.ReportPostProcessor;
import ru.mywayline.xlreport.core.api.ReportSession;
import ru.mywayline.xlreport.core.model.PostScriptConfig;
import ru.mywayline.xlreport.core.model.ReportConfig;
import ru.mywayline.xlreport.js.api.JsReportApi;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.regex.Pattern;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.graalvm.polyglot.Context;
import org.graalvm.polyglot.HostAccess;

/**
 * Executes JavaScript post-processing scripts against the generated workbook.
 *
 * <h3>Legacy VBA macro dispatch</h3>
 * When a script config has a {@code macroName} set (via
 * {@link #processMacro(ReportConfig, String, String, ReportSession)}), this processor:
 * <ol>
 *   <li>Parses {@code "Module1.m1"} → module={@code "Module1"}, function={@code "m1"}</li>
 *   <li>Loads {@code Module1.js} from the report's template directory (if it exists)</li>
 *   <li>Calls {@code Module1["m1"](report)} in the GraalJS context</li>
 * </ol>
 */
@Slf4j
public class GraalJsPostProcessor implements ReportPostProcessor {
    private static final Pattern IDENT = Pattern.compile("[A-Za-z_][A-Za-z0-9_]*");
    private static final String MACRO_DISPATCH_SCRIPT = """
        (function() {
          var mod = _macroModuleObj;
          var fn = _macroFunction;
          if (mod && typeof mod[fn] === 'function') {
            mod[fn](report);
          } else {
            report.warn('Macro function not found: ' + fn);
          }
        })();
        """;

    @Override
    public void process(ReportConfig reportConfig, PostScriptConfig scriptConfig, ReportSession session) throws Exception {
        if (!(session.documentHandle() instanceof XSSFWorkbook workbook)) {
            throw new IllegalArgumentException("Unsupported session document handle for JS post-processing.");
        }
        String script = resolveScript(scriptConfig);
        if (script == null || script.isBlank()) {
            return;
        }

        long timeoutMs = scriptConfig.getTimeoutMs() > 0 ? scriptConfig.getTimeoutMs() : 30_000L;
        JsReportApi reportApi = new JsReportApi(workbook, reportConfig.getParams(), session.buildStats());

        evalScript(reportApi, script, scriptConfig.getName(), timeoutMs, session.buildStats());
    }

    /**
     * Dispatches a legacy VBA macro call.
     *
     * @param macroName   e.g. "Module1.m1"
     * @param macroLibDir directory to search for the JS module file; typically the template dir
     */
    public void processMacro(
        ReportConfig reportConfig,
        String macroName,
        String macroLibDir,
        ReportSession session
    ) throws Exception {
        if (macroName == null || macroName.isBlank()) return;
        if (!(session.documentHandle() instanceof XSSFWorkbook workbook)) {
            throw new IllegalArgumentException("Unsupported session document handle for JS macro processing.");
        }

        // Parse "Module1.m1" → moduleName="Module1", functionName="m1"
        String moduleName;
        String functionName;
        int dot = macroName.indexOf('.');
        if (dot > 0) {
            moduleName   = macroName.substring(0, dot);
            functionName = macroName.substring(dot + 1);
        } else {
            moduleName   = null;
            functionName = macroName;
        }

        if (moduleName != null && !IDENT.matcher(moduleName).matches()) {
            log.warn("Invalid macro module name '{}': rejected", moduleName);
            return;
        }
        if (!IDENT.matcher(functionName).matches()) {
            log.warn("Invalid macro function name '{}': rejected", functionName);
            return;
        }

        // Load module JS file if present
        String moduleScript = null;
        if (moduleName != null && macroLibDir != null) {
            Path jsFile = Path.of(macroLibDir).resolve(moduleName + ".js");
            if (Files.exists(jsFile)) {
                moduleScript = Files.readString(jsFile);
                log.info("Loaded macro module file: {}", jsFile);
            } else {
                log.warn("Macro module file not found: {} — macro '{}' will not execute. " +
                    "Create the file and implement the function.", jsFile, macroName);
                return;
            }
        }

        if (moduleScript == null) {
            log.warn("Cannot dispatch macro '{}': no module file resolved. " +
                "Ensure the template directory contains {}.js", macroName,
                moduleName != null ? moduleName : "<module>");
            return;
        }

        String dispatchScript = "// Auto-dispatch: " + macroName + "\n" + MACRO_DISPATCH_SCRIPT;

        JsReportApi reportApi = new JsReportApi(workbook, reportConfig.getParams(), session.buildStats());
        evalMacroScript(reportApi, dispatchScript, moduleScript, moduleName, functionName, "macro-" + macroName, 60_000L, session.buildStats());
    }

    // ── Internal ──────────────────────────────────────────────────────────────

    private void evalScript(
        JsReportApi reportApi,
        String script,
        String scriptName,
        long timeoutMs,
        ReportBuildStats stats
    ) {
        evalScript(reportApi, script, scriptName, timeoutMs, stats, null, null, null);
    }

    private void evalMacroScript(
        JsReportApi reportApi,
        String dispatchScript,
        String moduleScript,
        String moduleName,
        String functionName,
        String scriptName,
        long timeoutMs,
        ReportBuildStats stats
    ) {
        evalScript(reportApi, dispatchScript, scriptName, timeoutMs, stats, moduleScript, moduleName, functionName);
    }

    private void evalScript(
        JsReportApi reportApi,
        String script,
        String scriptName,
        long timeoutMs,
        ReportBuildStats stats,
        String moduleScript,
        String moduleName,
        String functionName
    ) {
        HostAccess hostAccess = HostAccess.newBuilder(HostAccess.EXPLICIT)
            .allowArrayAccess(true)
            .build();
        long startedNs = System.nanoTime();
        try (Context context = Context.newBuilder("js")
            .allowIO(false)
            .allowHostAccess(hostAccess)
            .allowHostClassLookup(className -> false)
            .option("engine.WarnInterpreterOnly", "false")
            .build()) {

            context.getBindings("js").putMember("report", reportApi);
            context.getBindings("js").putMember("timeoutMs", timeoutMs);
            if (moduleScript != null) {
                context.eval("js", moduleScript);
                Object moduleObj = context.getBindings("js").getMember(moduleName);
                if (moduleObj == null) {
                    log.warn("Macro module '{}' was not defined by {}", moduleName, scriptName);
                    return;
                }
                context.getBindings("js").putMember("_macroModuleObj", moduleObj);
                context.getBindings("js").putMember("_macroFunction", functionName);
            }
            context.eval("js", script);
            log.debug("Executed JS script: {}", scriptName);
        } finally {
            reportApi.publishStyleCacheStats();
            if (stats != null) {
                stats.incrementExecutedScripts();
                stats.addScriptEvalMs((System.nanoTime() - startedNs) / 1_000_000L);
            }
        }
    }

    private String resolveScript(PostScriptConfig cfg) throws Exception {
        if (cfg.getInlineScript() != null && !cfg.getInlineScript().isBlank()) {
            return cfg.getInlineScript();
        }
        if (cfg.getScriptPath() != null && !cfg.getScriptPath().isBlank()) {
            return Files.readString(Path.of(cfg.getScriptPath()));
        }
        return null;
    }
}
