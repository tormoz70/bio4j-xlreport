package com.bio.xlreport.core.api;

import com.bio.xlreport.core.model.ReportConfig;

public interface ReportBuilder {
    ReportSession build(ReportConfig config, DataProvider dataProvider) throws Exception;
}
