package com.bio.xlreport.core.model;

import lombok.Builder;
import lombok.Value;

@Value
@Builder
public class ExtAttributes {
    String targetFormat;
    boolean liveScripts;
    String userLogin;
    String userUID;
    String remoteIP;
    String shortCode;
    String workPath;
    String localPath;
    String pwdOpen;
    String pwdWrite;
}
