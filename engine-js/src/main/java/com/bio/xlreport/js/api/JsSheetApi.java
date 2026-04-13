package com.bio.xlreport.js.api;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.graalvm.polyglot.HostAccess;

public class JsSheetApi {
    private final Sheet sheet;

    public JsSheetApi(Sheet sheet) {
        this.sheet = sheet;
    }

    /** Sheet name. */
    @HostAccess.Export
    public String name() {
        return sheet.getSheetName();
    }

    /**
     * Returns cell by A1-style reference, e.g. "B5".
     * Creates the row/cell if they don't exist.
     */
    @HostAccess.Export
    public JsCellApi cell(String a1Ref) {
        CellReference ref = new CellReference(a1Ref);
        return cellAt(ref.getRow() + 1, ref.getCol() + 1);
    }

    /**
     * Returns cell at 1-based row/col position.
     * Creates the row/cell if they don't exist.
     */
    @HostAccess.Export
    public JsCellApi cellAt(int rowOneBased, int colOneBased) {
        int row0 = rowOneBased - 1;
        int col0 = colOneBased - 1;
        Row row = sheet.getRow(row0);
        if (row == null) {
            row = sheet.createRow(row0);
        }
        var cell = row.getCell(col0);
        if (cell == null) {
            cell = row.createCell(col0);
        }
        return new JsCellApi(cell);
    }

    /**
     * Returns a named range scoped to this sheet or workbook-level.
     * If the named range is not found, returns null.
     */
    @HostAccess.Export
    public JsRangeApi namedRange(String name) {
        var wb = (XSSFWorkbook) sheet.getWorkbook();
        var namedRange = wb.getName(name);
        if (namedRange == null) {
            return null;
        }
        return parseAreaReference(namedRange.getRefersToFormula());
    }

    /**
     * Returns a range by A1-style reference, e.g. "A1:Z100" or "Sheet1!A1:Z100".
     */
    @HostAccess.Export
    public JsRangeApi range(String a1Ref) {
        return parseAreaReference(a1Ref);
    }

    /** Number of the last row that has data (1-based). Returns 0 if sheet is empty. */
    @HostAccess.Export
    public int rowCount() {
        return sheet.getLastRowNum() + 1;
    }

    /**
     * Number of the last column with data in the given row (1-based, row is 1-based).
     * Returns 0 if the row doesn't exist.
     */
    @HostAccess.Export
    public int colCount(int rowOneBased) {
        Row row = sheet.getRow(rowOneBased - 1);
        if (row == null) return 0;
        return row.getLastCellNum(); // getLastCellNum is already 1-based
    }

    /**
     * Group rows for outline (collapse/expand). Row indices are 1-based.
     */
    @HostAccess.Export
    public void groupRows(int startRowOneBased, int endRowOneBased) {
        sheet.groupRow(startRowOneBased - 1, endRowOneBased - 1);
    }

    /** Hide a row (1-based). */
    @HostAccess.Export
    public void hideRow(int rowOneBased) {
        Row row = sheet.getRow(rowOneBased - 1);
        if (row == null) {
            row = sheet.createRow(rowOneBased - 1);
        }
        row.setZeroHeight(true);
    }

    /** Show a previously hidden row (1-based). */
    @HostAccess.Export
    public void showRow(int rowOneBased) {
        Row row = sheet.getRow(rowOneBased - 1);
        if (row != null) {
            row.setZeroHeight(false);
        }
    }

    /**
     * Set column width. widthChars is the number of characters (same as Excel UI).
     * col is 1-based.
     */
    @HostAccess.Export
    public void setColumnWidth(int colOneBased, double widthChars) {
        // POI uses units of 1/256th of a character width
        sheet.setColumnWidth(colOneBased - 1, (int) (widthChars * 256));
    }

    /** Auto-size a column to fit its content. col is 1-based. */
    @HostAccess.Export
    public void autoSizeColumn(int colOneBased) {
        sheet.autoSizeColumn(colOneBased - 1);
    }

    private JsRangeApi parseAreaReference(String formula) {
        if (formula == null || formula.isBlank()) return null;
        try {
            AreaReference area = new AreaReference(formula, SpreadsheetVersion.EXCEL2007);
            CellReference first = area.getFirstCell();
            CellReference last = area.getLastCell();
            // Resolve sheet — may be different from this.sheet if formula includes sheet name
            var wb = (XSSFWorkbook) sheet.getWorkbook();
            Sheet targetSheet = sheet;
            if (first.getSheetName() != null) {
                Sheet s = wb.getSheet(first.getSheetName());
                if (s != null) targetSheet = s;
            }
            return new JsRangeApi(targetSheet, first.getRow(), first.getCol(),
                last.getRow(), last.getCol());
        } catch (Exception e) {
            return null;
        }
    }
}
