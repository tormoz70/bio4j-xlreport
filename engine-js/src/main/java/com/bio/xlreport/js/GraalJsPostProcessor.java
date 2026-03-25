package com.bio.xlreport.js;

import com.bio.xlreport.core.api.ReportPostProcessor;
import com.bio.xlreport.core.api.ReportSession;
import com.bio.xlreport.core.model.PostScriptConfig;
import com.bio.xlreport.core.model.ReportConfig;
import com.bio.xlreport.js.api.JsReportApi;
import java.nio.file.Files;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.graalvm.polyglot.Context;
import org.graalvm.polyglot.HostAccess;

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
        JsReportApi reportApi = new JsReportApi(workbook);

        HostAccess hostAccess = HostAccess.newBuilder(HostAccess.EXPLICIT).build();
        try (Context context = Context.newBuilder("js")
            .allowIO(false)
            .allowHostAccess(hostAccess)
            .allowHostClassLookup(className -> false)
            .option("engine.WarnInterpreterOnly", "false")
            .build()) {

            context.getBindings("js").putMember("report", reportApi);
            context.getBindings("js").putMember("timeoutMs", timeoutMs);
            context.eval("js", script);
            log.debug("Executed JS post-script: {}", scriptConfig.getName());
        }
    }

    private String resolveScript(PostScriptConfig cfg) throws Exception {
        if (cfg.getInlineScript() != null && !cfg.getInlineScript().isBlank()) {
            return cfg.getInlineScript();
        }
        if (cfg.getScriptPath() != null && !cfg.getScriptPath().isBlank()) {
            return Files.readString(java.nio.file.Path.of(cfg.getScriptPath()));
        }
        return null;
    }
}
