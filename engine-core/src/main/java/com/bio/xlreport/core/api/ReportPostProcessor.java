package com.bio.xlreport.core.api;

import com.bio.xlreport.core.model.PostScriptConfig;
import com.bio.xlreport.core.model.ReportConfig;

public interface ReportPostProcessor {
    void process(ReportConfig reportConfig, PostScriptConfig scriptConfig, ReportSession session) throws Exception;
}
