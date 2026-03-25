package com.bio.xlreport.core.model;

import lombok.Builder;
import lombok.Value;

@Value
@Builder
public class PostScriptConfig {
    String name;
    String scriptPath;
    String inlineScript;
    long timeoutMs;
}
