package ru.mywayline.xlreport.core.api;

import ru.mywayline.xlreport.core.model.ReportConfig;

public interface ReportBuilder {
    ReportSession build(ReportConfig config, DataProvider dataProvider) throws Exception;
}
