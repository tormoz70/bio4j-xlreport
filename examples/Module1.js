/**
 * Module1.js — Migration of VBA Module1 macros to JavaScript.
 *
 * Place this file in the same directory as the report XML / XLSM template.
 * The engine loads it automatically when the XML contains:
 *   <macroAfter name="Module1.m1"/>
 *   <macroAfter name="Module1.buildAll"/>
 *
 * Each function receives a single argument: `report` (JsReportApi).
 *
 * API overview:
 *   report.sheet(name)            → JsSheetApi
 *   report.sheetAt(idx)           → JsSheetApi  (1-based)
 *   report.sheetNames()           → String[]
 *   report.createSheet(name)      → JsSheetApi
 *   report.removeSheet(name)
 *   report.namedRange(name)       → JsRangeApi  (null if not found)
 *   report.namedRangeFormula(name)→ String
 *   report.param(name)            → String
 *   report.log(msg) / report.warn(msg)
 *
 *   sheet.cell("B5")             → JsCellApi
 *   sheet.cellAt(row, col)       → JsCellApi    (1-based)
 *   sheet.namedRange(name)       → JsRangeApi
 *   sheet.range("A1:Z100")       → JsRangeApi
 *   sheet.rowCount()             → int
 *   sheet.groupRows(from, to)
 *   sheet.hideRow(row)
 *   sheet.setColumnWidth(col, chars)
 *
 *   range.rowCount() / range.colCount()
 *   range.cell(row, col)         → JsCellApi    (1-based within range)
 *   range.firstRow/firstCol/lastRow/lastCol
 *   range.address()              → "Sheet1!A3:Z100"
 *
 *   cell.value() / cell.numericValue()
 *   cell.setValue(str) / cell.setNumericValue(d) / cell.setFormula(f)
 *   cell.setNumberFormat(fmt)    e.g. "m/d/yyyy", "#,##0.00", "#,##0.00$"
 *   cell.copyStyleFrom(otherCell)
 *   cell.row() / cell.col() / cell.address()
 *   cell.isEmpty() / cell.clear()
 */

var Module1 = {};

// ---------------------------------------------------------------------------
// Module1.m1 — Was: create PivotTable from named range "mRng" on new sheet "Свод"
//
// VBA original built a real Excel PivotTable (not possible via Apache POI).
// This JS replacement builds a summary cross-table manually:
//   - reads data from named range "mRng" on sheet 1
//   - aggregates by "Регион" (col 1), "Город" (col 2), "Киносеть" (col 3)
//   - writes grouped summary to new sheet "Свод"
//
// Columns in mRng (based on VBA pivot field mapping):
//  1=Неделя, 2=Регион, 3=Город, 4=Киносеть, 5=Дата начала показа,
//  6=Дата окончания показа, 7=Дней в показе, 8=Сеансов, 9=Зрителей, 10=Gross
// ---------------------------------------------------------------------------
Module1.m1 = function(report) {
    report.log("Module1.m1: building summary sheet");

    var ws = report.sheetAt(1);
    var mRng = ws.namedRange("mRng");
    if (mRng === null) {
        report.warn("Module1.m1: named range 'mRng' not found, skipping");
        return;
    }

    // --- header cells from source sheet (rows 5-6, cols 3-4 in VBA) ---
    var hdr1c1 = ws.cellAt(5, 3).value();
    var hdr1c2 = ws.cellAt(5, 4).value();
    var hdr2c1 = ws.cellAt(6, 3).value();
    var hdr2c2 = ws.cellAt(6, 4).value();

    // --- aggregate by Регион/Город/Киносеть ---
    var rows = mRng.rowCount();
    var aggMap = {};   // key → {region, city, network, sessions, viewers, gross, daysMin, daysMax, dateStart, dateEnd}

    for (var r = 1; r <= rows; r++) {
        var region  = mRng.cell(r, 2).value();
        var city    = mRng.cell(r, 3).value();
        var network = mRng.cell(r, 4).value();
        if (!region && !city && !network) continue;

        var key = region + "|" + city + "|" + network;
        var days    = mRng.cell(r, 7).numericValue();
        var sess    = mRng.cell(r, 8).numericValue();
        var viewers = mRng.cell(r, 9).numericValue();
        var gross   = mRng.cell(r, 10).numericValue();

        if (!aggMap[key]) {
            aggMap[key] = {
                region: region, city: city, network: network,
                sessions: 0, viewers: 0, gross: 0, days: 0
            };
        }
        var agg = aggMap[key];
        agg.sessions += sess;
        agg.viewers  += viewers;
        agg.gross    += gross;
        agg.days     = Math.max(agg.days, days);
    }

    // --- write to new sheet ---
    var svod = report.createSheet("Свод");

    svod.cellAt(1, 1).setValue(hdr1c1);
    svod.cellAt(1, 2).setValue(hdr1c2);
    svod.cellAt(2, 1).setValue(hdr2c1);
    svod.cellAt(2, 2).setValue(hdr2c2);

    // column headers
    svod.cellAt(4, 1).setValue("Регион");
    svod.cellAt(4, 2).setValue("Город");
    svod.cellAt(4, 3).setValue("Киносеть");
    svod.cellAt(4, 4).setValue("Итого дней");
    svod.cellAt(4, 5).setValue("Итого сеансов");
    svod.cellAt(4, 6).setValue("Итого зрителей");
    svod.cellAt(4, 7).setValue("Итого Gross");

    var outputRow = 5;
    var keys = Object.keys(aggMap).sort();
    for (var i = 0; i < keys.length; i++) {
        var a = aggMap[keys[i]];
        svod.cellAt(outputRow, 1).setValue(a.region);
        svod.cellAt(outputRow, 2).setValue(a.city);
        svod.cellAt(outputRow, 3).setValue(a.network);
        svod.cellAt(outputRow, 4).setNumericValue(a.days);
        svod.cellAt(outputRow, 5).setNumericValue(a.sessions);
        svod.cellAt(outputRow, 6).setNumericValue(a.viewers);
        var grossCell = svod.cellAt(outputRow, 7);
        grossCell.setNumericValue(a.gross);
        grossCell.setNumberFormat("#,##0.00$");
        outputRow++;
    }

    report.log("Module1.m1: summary sheet 'Свод' written, " + (outputRow - 5) + " rows");
};

// ---------------------------------------------------------------------------
// Module1.buildAll — Was: build city×date×theatre matrix
//
// VBA original built a complex cross-table from mRngSumm and mRngDays.
// This JS replacement is a stub — implement the required aggregation logic.
// ---------------------------------------------------------------------------
Module1.buildAll = function(report) {
    report.log("Module1.buildAll: not yet implemented — add logic here");
    report.warn("Module1.buildAll requires a custom implementation. See Module1.js for guidance.");
};

// ---------------------------------------------------------------------------
// Module1.mGoto — Was: empty stub in VBA
// ---------------------------------------------------------------------------
Module1.mGoto = function(report) {
    // intentionally empty — the original VBA sub was also empty
};

// ---------------------------------------------------------------------------
// Module1.Macros1 — inspect if needed, currently a stub
// ---------------------------------------------------------------------------
Module1.Macros1 = function(report) {
    report.log("Module1.Macros1: not yet implemented");
};
