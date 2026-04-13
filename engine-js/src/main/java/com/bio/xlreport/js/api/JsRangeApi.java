package com.bio.xlreport.js.api;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.graalvm.polyglot.HostAccess;

/**
 * Represents a rectangular range of cells within a sheet, exposed to JavaScript.
 * Row and column indices are 1-based in all public methods.
 */
public class JsRangeApi {
    private final Sheet sheet;
    private final int firstRow0; // 0-based
    private final int firstCol0; // 0-based
    private final int lastRow0;  // 0-based (inclusive)
    private final int lastCol0;  // 0-based (inclusive)

    public JsRangeApi(Sheet sheet, int firstRow0, int firstCol0, int lastRow0, int lastCol0) {
        this.sheet = sheet;
        this.firstRow0 = firstRow0;
        this.firstCol0 = firstCol0;
        this.lastRow0 = lastRow0;
        this.lastCol0 = lastCol0;
    }

    /** Number of rows in the range. */
    @HostAccess.Export
    public int rowCount() {
        return lastRow0 - firstRow0 + 1;
    }

    /** Number of columns in the range. */
    @HostAccess.Export
    public int colCount() {
        return lastCol0 - firstCol0 + 1;
    }

    /**
     * Returns the cell at the given 1-based position within the range.
     * e.g. cell(1,1) = top-left cell of the range.
     */
    @HostAccess.Export
    public JsCellApi cell(int rowOneBased, int colOneBased) {
        int absRow = firstRow0 + rowOneBased - 1;
        int absCol = firstCol0 + colOneBased - 1;
        Row row = sheet.getRow(absRow);
        if (row == null) {
            row = sheet.createRow(absRow);
        }
        var cell = row.getCell(absCol);
        if (cell == null) {
            cell = row.createCell(absCol);
        }
        return new JsCellApi(cell);
    }

    /** Absolute 1-based row index of the first row in this range. */
    @HostAccess.Export
    public int firstRow() {
        return firstRow0 + 1;
    }

    /** Absolute 1-based column index of the first column in this range. */
    @HostAccess.Export
    public int firstCol() {
        return firstCol0 + 1;
    }

    /** Absolute 1-based row index of the last row in this range. */
    @HostAccess.Export
    public int lastRow() {
        return lastRow0 + 1;
    }

    /** Absolute 1-based column index of the last column in this range. */
    @HostAccess.Export
    public int lastCol() {
        return lastCol0 + 1;
    }

    /** Sheet name this range belongs to. */
    @HostAccess.Export
    public String sheetName() {
        return sheet.getSheetName();
    }

    /** A1-style address of the range, e.g. "Sheet1!A3:Z100". */
    @HostAccess.Export
    public String address() {
        return sheet.getSheetName() + "!" +
            colLetter(firstCol0) + (firstRow0 + 1) + ":" +
            colLetter(lastCol0) + (lastRow0 + 1);
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
