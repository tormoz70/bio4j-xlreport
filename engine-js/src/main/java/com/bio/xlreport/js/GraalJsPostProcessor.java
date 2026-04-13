package com.bio.xlreport.js;

import com.bio.xlreport.core.api.ReportPostProcessor;
import com.bio.xlreport.core.api.ReportSession;
import com.bio.xlreport.core.model.PostScriptConfig;
import com.bio.xlreport.core.model.ReportConfig;
import com.bio.xlreport.js.api.JsReportApi;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;
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
        JsReportApi reportApi = new JsReportApi(workbook, reportConfig.getParams());

        evalScript(reportApi, script, scriptConfig.getName(), timeoutMs);
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

        String safeModule   = moduleName;
        String safeFunction = functionName.replace("'", "\\'");
        String dispatchScript = moduleScript + "\n" +
            "// Auto-dispatch: " + macroName + "\n" +
            "(function() {\n" +
            "  var _mod = (typeof " + safeModule + " !== 'undefined') ? " + safeModule + " : null;\n" +
            "  if (_mod && typeof _mod['" + safeFunction + "'] === 'function') {\n" +
            "    _mod['" + safeFunction + "'](report);\n" +
            "  } else {\n" +
            "    report.warn('Macro function not found: " + macroName.replace("'", "\\'") + "');\n" +
            "  }\n" +
            "})();\n";

        JsReportApi reportApi = new JsReportApi(workbook, reportConfig.getParams());
        evalScript(reportApi, dispatchScript, "macro-" + macroName, 60_000L);
    }

    // ── Internal ──────────────────────────────────────────────────────────────

    private void evalScript(JsReportApi reportApi, String script, String scriptName, long timeoutMs) {
        HostAccess hostAccess = HostAccess.newBuilder(HostAccess.EXPLICIT)
            .allowArrayAccess(true)
            .build();
        try (Context context = Context.newBuilder("js")
            .allowIO(false)
            .allowHostAccess(hostAccess)
            .allowHostClassLookup(className -> false)
            .option("engine.WarnInterpreterOnly", "false")
            .build()) {

            context.getBindings("js").putMember("report", reportApi);
            context.getBindings("js").putMember("timeoutMs", timeoutMs);
            context.eval("js", script);
            log.debug("Executed JS script: {}", scriptName);
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
