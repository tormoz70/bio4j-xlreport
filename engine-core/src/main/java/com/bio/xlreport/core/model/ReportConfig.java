package com.bio.xlreport.core.model;

import java.nio.file.Path;
import java.util.List;
import java.util.Map;
import lombok.Builder;
import lombok.Singular;
import lombok.Value;

@Value
@Builder(toBuilder = true)
public class ReportConfig {
    String uid;
    String fullCode;
    String title;
    String subject;
    String author;

    Path templatePath;
    Path outputPath;

    CompatibilityMode compatibilityMode;
    boolean convertResultToPdf;

    ExtAttributes extAttributes;

    @Singular("param")
    Map<String, String> params;
    @Singular("env")
    Map<String, String> envVars;
    @Singular
    List<DataSourceConfig> dataSources;
    @Singular
    List<PostScriptConfig> postScripts;
}
