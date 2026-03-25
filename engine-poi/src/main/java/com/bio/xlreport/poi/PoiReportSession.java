package com.bio.xlreport.poi;

import com.bio.xlreport.core.api.ReportSession;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.FileSystemException;
import java.nio.file.Path;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PoiReportSession implements ReportSession {
    private final XSSFWorkbook workbook;
    private Path outputPath;

    public PoiReportSession(XSSFWorkbook workbook, Path outputPath) {
        this.workbook = workbook;
        this.outputPath = outputPath;
    }

    @Override
    public Path outputPath() {
        return outputPath;
    }

    @Override
    public Object documentHandle() {
        return workbook;
    }

    @Override
    public void save() throws Exception {
        // Force Excel to recalculate formulas (totals, aggregates) on open.
        workbook.setForceFormulaRecalculation(true);
        if (outputPath != null && outputPath.getParent() != null) {
            Files.createDirectories(outputPath.getParent());
        }
        try {
            writeTo(outputPath);
        } catch (FileSystemException ex) {
            // Common on Windows when the target file is open in Excel.
            Path fallback = buildFallbackPath(outputPath);
            writeTo(fallback);
            outputPath = fallback;
        }
    }

    @Override
    public void close() throws Exception {
        workbook.close();
    }

    private void writeTo(Path target) throws Exception {
        try (OutputStream out = Files.newOutputStream(target)) {
            workbook.write(out);
        }
    }

    private Path buildFallbackPath(Path target) {
        String name = target.getFileName().toString();
        int dot = name.lastIndexOf('.');
        String base = dot > 0 ? name.substring(0, dot) : name;
        String ext = dot > 0 ? name.substring(dot) : "";
        int idx = 1;
        Path parent = target.getParent();
        Path candidate;
        do {
            candidate = (parent == null)
                ? Path.of(base + "-" + idx + ext)
                : parent.resolve(base + "-" + idx + ext);
            idx++;
        } while (Files.exists(candidate));
        return candidate;
    }
}
