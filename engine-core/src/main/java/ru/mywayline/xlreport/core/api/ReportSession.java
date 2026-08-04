package ru.mywayline.xlreport.core.api;

import java.nio.file.Path;

public interface ReportSession extends AutoCloseable {
    Path outputPath();

    Object documentHandle();

    void save() throws Exception;

    /**
     * Optional build-time counters; returns null when the session does not collect stats.
     */
    default ReportBuildStats buildStats() {
        return null;
    }

    @Override
    void close() throws Exception;
}
