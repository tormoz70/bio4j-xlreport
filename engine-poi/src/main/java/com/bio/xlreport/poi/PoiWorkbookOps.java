package com.bio.xlreport.poi;

import java.util.Locale;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;

public final class PoiWorkbookOps {

    private PoiWorkbookOps() {
    }

    public static Name findName(Workbook wb, String name) {
        for (Name n : wb.getAllNames()) {
            if (n.getNameName().equalsIgnoreCase(name)) {
                return n;
            }
        }
        return null;
    }

    public static AreaReference resolveArea(Workbook wb, Name name) {
        String formula = name.getRefersToFormula();
        return new AreaReference(formula, SpreadsheetVersion.EXCEL2007);
    }

    public static void setValue(Cell cell, Object value) {
        if (value == null) {
            cell.setBlank();
            return;
        }
        if (value instanceof Number n) {
            cell.setCellValue(n.doubleValue());
            return;
        }
        if (value instanceof Boolean b) {
            cell.setCellValue(b);
            return;
        }
        String text = String.valueOf(value);
        if (text.startsWith("=")) {
            cell.setCellFormula(text.substring(1));
        } else {
            cell.setCellValue(text);
        }
    }

    public static void applyPlaceholder(Sheet sheet, String token, String value) {
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell.getCellType() == CellType.STRING) {
                    String current = cell.getStringCellValue();
                    if (current != null && current.contains(token)) {
                        cell.setCellValue(current.replace(token, value == null ? "" : value));
                    }
                }
            }
        }
    }

    public static String a1(int rowIdxZeroBased, int colIdxZeroBased) {
        return new CellReference(rowIdxZeroBased, colIdxZeroBased).formatAsString();
    }

    public static String rangeA1(
        String sheetName,
        int rowStartZeroBased,
        int rowEndZeroBased,
        int colStartZeroBased,
        int colEndZeroBased
    ) {
        String left = a1(rowStartZeroBased, colStartZeroBased);
        String right = a1(rowEndZeroBased, colEndZeroBased);
        return "'" + sheetName.replace("'", "''") + "'!" + left + ":" + right;
    }

    public static String safeSheetName(String sheetName) {
        return sheetName == null ? "" : sheetName.toLowerCase(Locale.ROOT);
    }
}
