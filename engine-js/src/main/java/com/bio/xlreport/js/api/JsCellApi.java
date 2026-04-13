package com.bio.xlreport.js.api;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.graalvm.polyglot.HostAccess;

public class JsCellApi {
    private final Cell cell;

    public JsCellApi(Cell cell) {
        this.cell = cell;
    }

    /** Cell value as string. Formula cells return "=<formula>". */
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

    /** Cell value as double. Returns 0 for non-numeric cells. */
    @HostAccess.Export
    public double numericValue() {
        return switch (cell.getCellType()) {
            case NUMERIC -> cell.getNumericCellValue();
            case STRING -> {
                try {
                    yield Double.parseDouble(cell.getStringCellValue().trim());
                } catch (NumberFormatException e) {
                    yield 0.0;
                }
            }
            case BOOLEAN -> cell.getBooleanCellValue() ? 1.0 : 0.0;
            default -> 0.0;
        };
    }

    /** Set cell value. Strings starting with "=" are set as formulas. */
    @HostAccess.Export
    public void setValue(String value) {
        if (value != null && value.startsWith("=")) {
            cell.setCellFormula(value.substring(1));
        } else {
            cell.setCellValue(value == null ? "" : value);
        }
    }

    /** Set cell value as a number. */
    @HostAccess.Export
    public void setNumericValue(double value) {
        cell.setCellValue(value);
    }

    /** Set cell value as boolean. */
    @HostAccess.Export
    public void setBooleanValue(boolean value) {
        cell.setCellValue(value);
    }

    /** Set an Excel formula (without leading "="). */
    @HostAccess.Export
    public void setFormula(String formula) {
        cell.setCellFormula(formula);
    }

    /** Clear cell content (blank cell). */
    @HostAccess.Export
    public void clear() {
        cell.setBlank();
    }

    /**
     * Apply a number format string to this cell's style.
     * Examples: "m/d/yyyy", "#,##0.00", "#,##0.00$", "@" (text), "0.00%"
     */
    @HostAccess.Export
    public void setNumberFormat(String formatString) {
        var workbook = cell.getSheet().getWorkbook();
        var helper = workbook.getCreationHelper();
        var style = workbook.createCellStyle();
        style.cloneStyleFrom(cell.getCellStyle());
        style.setDataFormat(helper.createDataFormat().getFormat(formatString));
        cell.setCellStyle(style);
    }

    /**
     * Copy cell style (number format, font, fill, borders, alignment) from another cell.
     */
    @HostAccess.Export
    public void copyStyleFrom(JsCellApi source) {
        var workbook = cell.getSheet().getWorkbook();
        var newStyle = workbook.createCellStyle();
        newStyle.cloneStyleFrom(source.cell.getCellStyle());
        cell.setCellStyle(newStyle);
    }

    /** 1-based row index of this cell within the sheet. */
    @HostAccess.Export
    public int row() {
        return cell.getRowIndex() + 1;
    }

    /** 1-based column index of this cell within the sheet. */
    @HostAccess.Export
    public int col() {
        return cell.getColumnIndex() + 1;
    }

    /** A1-style address of this cell, e.g. "B5". */
    @HostAccess.Export
    public String address() {
        return colLetter(cell.getColumnIndex()) + (cell.getRowIndex() + 1);
    }

    /** Returns true if the cell is blank / empty. */
    @HostAccess.Export
    public boolean isEmpty() {
        return cell.getCellType() == CellType.BLANK ||
            (cell.getCellType() == CellType.STRING && cell.getStringCellValue().isEmpty());
    }

    private static String colLetter(int col0) {
        StringBuilder sb = new StringBuilder();
        int c = col0;
        do {
            sb.insert(0, (char) ('A' + c % 26));
            c = c / 26 - 1;
        } while (c >= 0);
        return sb.toString();
    }
}
