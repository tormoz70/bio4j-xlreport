package com.bio.xlreport.js.api;

import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.graalvm.polyglot.HostAccess;

public class JsReportApi {
    private final XSSFWorkbook workbook;

    public JsReportApi(XSSFWorkbook workbook) {
        this.workbook = workbook;
    }

    @HostAccess.Export
    public JsSheetApi sheet(String name) {
        var sheet = workbook.getSheet(name);
        if (sheet == null) {
            throw new IllegalArgumentException("Sheet not found: " + name);
        }
        return new JsSheetApi(sheet);
    }

    @HostAccess.Export
    public String namedRange(String rangeName) {
        Name name = workbook.getName(rangeName);
        if (name == null) {
            return null;
        }
        return name.getRefersToFormula();
    }
}
