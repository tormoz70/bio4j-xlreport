package ru.mywayline.xlreport.core.api;

import ru.mywayline.xlreport.core.model.PostScriptConfig;
import ru.mywayline.xlreport.core.model.ReportConfig;

public interface ReportPostProcessor {
    void process(ReportConfig reportConfig, PostScriptConfig scriptConfig, ReportSession session) throws Exception;
}
