package com.bio.xlreport.core.api;

import java.nio.file.Path;

public interface ReportSession extends AutoCloseable {
    Path outputPath();

    Object documentHandle();

    void save() throws Exception;

    @Override
    void close() throws Exception;
}
