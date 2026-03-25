package com.bio.xlreport.js.api;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.graalvm.polyglot.HostAccess;

public class JsSheetApi {
    private final Sheet sheet;

    public JsSheetApi(Sheet sheet) {
        this.sheet = sheet;
    }

    @HostAccess.Export
    public String name() {
        return sheet.getSheetName();
    }

    @HostAccess.Export
    public JsCellApi cell(String a1Ref) {
        CellReference ref = new CellReference(a1Ref);
        Row row = sheet.getRow(ref.getRow());
        if (row == null) {
            row = sheet.createRow(ref.getRow());
        }
        var cell = row.getCell(ref.getCol());
        if (cell == null) {
            cell = row.createCell(ref.getCol());
        }
        return new JsCellApi(cell);
    }

    @HostAccess.Export
    public void groupRows(int startRowOneBased, int endRowOneBased) {
        sheet.groupRow(startRowOneBased - 1, endRowOneBased - 1);
    }
}
