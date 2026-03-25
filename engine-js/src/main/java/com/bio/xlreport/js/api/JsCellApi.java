package com.bio.xlreport.js.api;

import org.apache.poi.ss.usermodel.Cell;
import org.graalvm.polyglot.HostAccess;

public class JsCellApi {
    private final Cell cell;

    public JsCellApi(Cell cell) {
        this.cell = cell;
    }

    @HostAccess.Export
    public String value() {
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> Double.toString(cell.getNumericCellValue());
            case BOOLEAN -> Boolean.toString(cell.getBooleanCellValue());
            case FORMULA -> "=" + cell.getCellFormula();
            default -> "";
        };
    }

    @HostAccess.Export
    public void setValue(String value) {
        if (value != null && value.startsWith("=")) {
            cell.setCellFormula(value.substring(1));
        } else {
            cell.setCellValue(value == null ? "" : value);
        }
    }
}
