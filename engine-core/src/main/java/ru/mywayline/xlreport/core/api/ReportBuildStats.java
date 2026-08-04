package ru.mywayline.xlreport.core.api;

/**
 * Lightweight counters collected during report build and JS post-processing.
 */
public class ReportBuildStats {
    private int processedDataSources = 0;
    private int generatedRows = 0;
    private int executedScripts = 0;
    private long scriptEvalMs = 0;
    private int createdStyles = 0;
    private int styleCacheHits = 0;
    private long outputFileBytes = 0;

    public int processedDataSources() {
        return processedDataSources;
    }

    public int generatedRows() {
        return generatedRows;
    }

    public int executedScripts() {
        return executedScripts;
    }

    public long scriptEvalMs() {
        return scriptEvalMs;
    }

    public int createdStyles() {
        return createdStyles;
    }

    public int styleCacheHits() {
        return styleCacheHits;
    }

    public long outputFileBytes() {
        return outputFileBytes;
    }

    public void incrementDataSources() {
        processedDataSources++;
    }

    public void addGeneratedRows(int count) {
        if (count > 0) {
            generatedRows += count;
        }
    }

    public void incrementExecutedScripts() {
        executedScripts++;
    }

    public void addScriptEvalMs(long ms) {
        if (ms > 0) {
            scriptEvalMs += ms;
        }
    }

    public void addStyleCacheStats(int created, int hits) {
        if (created > 0) {
            createdStyles += created;
        }
        if (hits > 0) {
            styleCacheHits += hits;
        }
    }

    public void setOutputFileBytes(long bytes) {
        outputFileBytes = bytes;
    }

    public String summaryLine() {
        return "stats: dataSources=" + processedDataSources
            + ", rows=" + generatedRows
            + ", scripts=" + executedScripts
            + ", scriptMs=" + scriptEvalMs
            + ", stylesCreated=" + createdStyles
            + ", styleCacheHits=" + styleCacheHits
            + ", outputBytes=" + outputFileBytes;
    }
}
