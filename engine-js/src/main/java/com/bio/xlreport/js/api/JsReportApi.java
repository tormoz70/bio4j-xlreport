package com.bio.xlreport.js.api;

import java.util.Collections;
import java.util.Map;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.graalvm.polyglot.HostAccess;

@Slf4j
public class JsReportApi {
    private final XSSFWorkbook workbook;
    private final Map<String, String> params;

    public JsReportApi(XSSFWorkbook workbook) {
        this(workbook, Collections.emptyMap());
    }

    public JsReportApi(XSSFWorkbook workbook, Map<String, String> params) {
        this.workbook = workbook;
        this.params = params != null ? params : Collections.emptyMap();
    }

    // ── Sheet access ──────────────────────────────────────────────────────────

    /** Get sheet by name. Throws if not found. */
    @HostAccess.Export
    public JsSheetApi sheet(String name) {
        var sheet = workbook.getSheet(name);
        if (sheet == null) {
            throw new IllegalArgumentException("Sheet not found: " + name);
        }
        return new JsSheetApi(sheet);
    }

    /** Get sheet by 1-based index. */
    @HostAccess.Export
    public JsSheetApi sheetAt(int indexOneBased) {
        return new JsSheetApi(workbook.getSheetAt(indexOneBased - 1));
    }

    /** Names of all sheets in the workbook. */
    @HostAccess.Export
    public String[] sheetNames() {
        String[] names = new String[workbook.getNumberOfSheets()];
        for (int i = 0; i < names.length; i++) {
            names[i] = workbook.getSheetName(i);
        }
        return names;
    }

    /**
     * Create a new sheet at the end of the workbook.
     * If a sheet with that name already exists, returns the existing sheet.
     */
    @HostAccess.Export
    public JsSheetApi createSheet(String name) {
        XSSFSheet existing = workbook.getSheet(name);
        if (existing != null) {
            return new JsSheetApi(existing);
        }
        return new JsSheetApi(workbook.createSheet(name));
    }

    /**
     * Create a new sheet inserted at the given 1-based position.
     * If a sheet with that name already exists, returns the existing sheet.
     */
    @HostAccess.Export
    public JsSheetApi createSheetAt(String name, int positionOneBased) {
        XSSFSheet existing = workbook.getSheet(name);
        if (existing != null) {
            return new JsSheetApi(existing);
        }
        XSSFSheet sheet = workbook.createSheet(name);
        workbook.setSheetOrder(name, positionOneBased - 1);
        return new JsSheetApi(sheet);
    }

    /** Remove a sheet by name. No-op if the sheet does not exist. */
    @HostAccess.Export
    public void removeSheet(String name) {
        int idx = workbook.getSheetIndex(name);
        if (idx >= 0) {
            workbook.removeSheetAt(idx);
        }
    }

    // ── Named ranges ─────────────────────────────────────────────────────────

    /**
     * Returns a named range as a JsRangeApi object.
     * Returns null if the named range does not exist.
     */
    @HostAccess.Export
    public JsRangeApi namedRange(String rangeName) {
        var name = workbook.getName(rangeName);
        if (name == null) {
            return null;
        }
        String formula = name.getRefersToFormula();
        if (formula == null || formula.isBlank()) {
            return null;
        }
        try {
            AreaReference area = new AreaReference(formula, SpreadsheetVersion.EXCEL2007);
            CellReference first = area.getFirstCell();
            CellReference last = area.getLastCell();
            String sheetName = first.getSheetName();
            var sheet = sheetName != null ? workbook.getSheet(sheetName) : workbook.getSheetAt(0);
            if (sheet == null) return null;
            return new JsRangeApi(sheet, first.getRow(), first.getCol(),
                last.getRow(), last.getCol());
        } catch (Exception e) {
            log.warn("Cannot resolve named range '{}': {}", rangeName, e.getMessage());
            return null;
        }
    }

    /**
     * Returns the raw formula string of a named range (e.g. "Sheet1!$A$1:$Z$10").
     * Useful when you just need the address, not the cells.
     */
    @HostAccess.Export
    public String namedRangeFormula(String rangeName) {
        var name = workbook.getName(rangeName);
        return name != null ? name.getRefersToFormula() : null;
    }

    // ── Report parameters ────────────────────────────────────────────────────

    /**
     * Get a report parameter value by name (from the XML report definition params).
     * Returns null if the parameter is not defined.
     */
    @HostAccess.Export
    public String param(String name) {
        return params.get(name);
    }

    /**
     * Get all parameter names.
     */
    @HostAccess.Export
    public String[] paramNames() {
        return params.keySet().toArray(new String[0]);
    }

    // ── Logging ──────────────────────────────────────────────────────────────

    /** Log an informational message from the JS script. */
    @HostAccess.Export
    public void log(String message) {
        log.info("[JS] {}", message);
    }

    /** Log a warning from the JS script. */
    @HostAccess.Export
    public void warn(String message) {
        log.warn("[JS] {}", message);
    }

    /**
     * Legacy VBA macro migration bridge.
     * @deprecated Implement the macro logic directly in JavaScript instead.
     */
    @Deprecated
    @HostAccess.Export
    public boolean applyLegacyMacro(String macroName) {
        log.warn("[JS] applyLegacyMacro called for '{}' — implement the macro body in JS instead.", macroName);
        return macroName != null && !macroName.isBlank();
    }
}
