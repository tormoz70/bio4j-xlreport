package ru.mywayline.xlreport.core;

import ru.mywayline.xlreport.core.api.DataProvider;
import ru.mywayline.xlreport.core.api.ReportBuilder;
import ru.mywayline.xlreport.core.api.ReportPostProcessor;
import ru.mywayline.xlreport.core.api.ReportSession;
import ru.mywayline.xlreport.core.model.CompatibilityMode;
import ru.mywayline.xlreport.core.model.ReportConfig;
import ru.mywayline.xlreport.core.parse.XmlReportConfigParser;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import lombok.RequiredArgsConstructor;

@RequiredArgsConstructor
public class ReportEngine {
    private final XmlReportConfigParser configParser;
    private final ReportBuilder reportBuilder;
    private final List<ReportPostProcessor> postProcessors;

    public static ReportEngine of(XmlReportConfigParser parser, ReportBuilder builder) {
        return new ReportEngine(parser, builder, new ArrayList<>());
    }

    public ReportEngine withPostProcessor(ReportPostProcessor postProcessor) {
        List<ReportPostProcessor> all = new ArrayList<>(postProcessors);
        all.add(postProcessor);
        return new ReportEngine(configParser, reportBuilder, all);
    }

    public Path buildFromXml(String xmlConfig, DataProvider dataProvider, CompatibilityMode mode) throws Exception {
        ReportConfig config = configParser.parse(xmlConfig, mode);
        return build(config, dataProvider);
    }

    public Path build(ReportConfig reportConfig, DataProvider dataProvider) throws Exception {
        try (ReportSession session = reportBuilder.build(reportConfig, dataProvider)) {
            for (var script : reportConfig.getPostScripts()) {
                for (var postProcessor : postProcessors) {
                    postProcessor.process(reportConfig, script, session);
                }
            }
            session.save();
            return session.outputPath();
        }
    }
}
