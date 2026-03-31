package com.bio.xlreport.poi;

import com.bio.xlreport.core.api.DataProvider;
import com.bio.xlreport.core.api.ReportBuilder;
import com.bio.xlreport.core.api.ReportSession;
import com.bio.xlreport.core.model.CompatibilityMode;
import com.bio.xlreport.core.model.DataSourceConfig;
import com.bio.xlreport.core.model.FieldDef;
import com.bio.xlreport.core.model.ReportConfig;
import com.bio.xlreport.core.model.SortDef;
import com.bio.xlreport.core.model.SortDirection;
import java.io.InputStream;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.Collections;
import java.util.LinkedHashSet;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

@Slf4j
public class PoiReportBuilder implements ReportBuilder {
    private static final Pattern PLACEHOLDER_PATTERN = Pattern.compile("\\w+_[\\w#$]+\\b");
    private static final Pattern GROUP_MARKER_PATTERN = Pattern.compile("(?i)^(group[HF])\\s*[_=]\\s*(.+)$");
    private static final Pattern TOTAL_DIRECTIVE_PATTERN = Pattern.compile("(?i)^(sum|cnt|avg|min|max)\\s*(?:\\(\\s*([^)]+)\\s*\\))?$");

    @Override
    public ReportSession build(ReportConfig config, DataProvider dataProvider) throws Exception {
        try (InputStream in = Files.newInputStream(config.getTemplatePath())) {
            XSSFWorkbook workbook = new XSSFWorkbook(in);

            for (DataSourceConfig ds : config.getDataSources()) {
                writeDataSource(workbook, ds, dataProvider.fetch(ds), config.getCompatibilityMode());
            }
            applyGlobalPlaceholders(workbook, config);

            return new PoiReportSession(workbook, config.getOutputPath());
        }
    }

    private void writeDataSource(
        XSSFWorkbook wb,
        DataSourceConfig ds,
        List<Map<String, Object>> rows,
        CompatibilityMode mode
    ) {
        if (rows == null) {
            rows = List.of();
        }
        rows = applySorts(rows, ds.getSorts());
        if (ds.isSingleRow()) {
            applySingleRowDataSource(wb, ds, rows);
            return;
        }
        Name name = PoiWorkbookOps.findName(wb, ds.getRangeName());
        if (name == null) {
            throw new IllegalArgumentException("Named range not found in template: " + ds.getRangeName());
        }

        AreaReference area = PoiWorkbookOps.resolveArea(wb, name);
        CellReference first = area.getFirstCell();
        CellReference last = area.getLastCell();
        Sheet sheet = null;
        if (first.getSheetName() != null) {
            sheet = wb.getSheet(first.getSheetName());
        } else if (name.getSheetIndex() >= 0) {
            sheet = wb.getSheetAt(name.getSheetIndex());
        }
        if (sheet == null) {
            throw new IllegalArgumentException("Sheet not found for named range: " + ds.getRangeName());
        }

        int startRow = first.getRow();
        int endRowTemplate = last.getRow();
        int startCol = Math.min(first.getCol(), last.getCol());
        int endCol = Math.max(first.getCol(), last.getCol());

        TemplateDef templateDef = parseTemplate(sheet, startRow, endRowTemplate, startCol, endCol, ds, mode);
        BuildResult buildResult = buildOutputRows(rows, templateDef);
        List<OutputRow> outputRows = buildResult.rows();
        int rowCount = Math.max(outputRows.size(), 1);

        if (rowCount > templateDef.templateHeight() && sheet.getLastRowNum() >= (endRowTemplate + 1)) {
            sheet.shiftRows(endRowTemplate + 1, sheet.getLastRowNum(), rowCount - templateDef.templateHeight(), true, false);
        }

        writeRows(sheet, startRow, startCol, endCol, templateDef, outputRows);
        applyGrouping(sheet, startRow, buildResult.groupSpans());
        int endRow = startRow + outputRows.size() - 1;
        String newRange = PoiWorkbookOps.rangeA1(sheet.getSheetName(), startRow, endRow, startCol, endCol);
        name.setRefersToFormula(newRange);
    }

    private List<Map<String, Object>> applySorts(List<Map<String, Object>> rows, List<SortDef> sorts) {
        if (rows == null || rows.size() <= 1 || sorts == null || sorts.isEmpty()) {
            return rows == null ? List.of() : rows;
        }
        List<Map<String, Object>> sorted = new ArrayList<>(rows);
        sorted.sort((a, b) -> compareBySorts(a, b, sorts));
        return sorted;
    }

    private int compareBySorts(Map<String, Object> a, Map<String, Object> b, List<SortDef> sorts) {
        for (SortDef sort : sorts) {
            if (sort == null || sort.getFieldName() == null || sort.getFieldName().isBlank()) {
                continue;
            }
            Object va = lookupIgnoreCase(a, sort.getFieldName());
            Object vb = lookupIgnoreCase(b, sort.getFieldName());
            int cmp = compareValues(va, vb);
            if (cmp == 0) {
                continue;
            }
            if (sort.getDirection() == SortDirection.DESC) {
                cmp = -cmp;
            }
            return cmp;
        }
        return 0;
    }

    private Object lookupIgnoreCase(Map<String, Object> row, String key) {
        if (row == null || key == null) {
            return null;
        }
        Object direct = row.get(key);
        if (direct != null) {
            return direct;
        }
        for (var e : row.entrySet()) {
            if (e.getKey() != null && e.getKey().equalsIgnoreCase(key)) {
                return e.getValue();
            }
        }
        return null;
    }

    @SuppressWarnings({ "rawtypes", "unchecked" })
    private int compareValues(Object a, Object b) {
        if (a == b) {
            return 0;
        }
        if (a == null) {
            return -1;
        }
        if (b == null) {
            return 1;
        }
        if (a instanceof Comparable ca && a.getClass().isAssignableFrom(b.getClass())) {
            return ca.compareTo(b);
        }
        return String.valueOf(a).compareToIgnoreCase(String.valueOf(b));
    }

    private void applySingleRowDataSource(XSSFWorkbook wb, DataSourceConfig ds, List<Map<String, Object>> rows) {
        Map<String, Object> firstRow = rows.isEmpty() ? Collections.emptyMap() : rows.getFirst();
        for (Sheet sheet : wb) {
            for (var entry : firstRow.entrySet()) {
                String tokenRaw = ds.getRangeName() + "_" + entry.getKey();
                String value = String.valueOf(entry.getValue() == null ? "" : entry.getValue());
                PoiWorkbookOps.applyPlaceholder(sheet, "(" + tokenRaw + ")", value);
                PoiWorkbookOps.applyPlaceholder(sheet, tokenRaw, value);
            }
        }
    }

    private BuildResult buildOutputRows(List<Map<String, Object>> rows, TemplateDef templateDef) {
        List<OutputRow> out = new ArrayList<>();
        List<GroupSpan> spans = new ArrayList<>();
        List<Integer> allDetailIndexes = new ArrayList<>();
        if (rows == null) {
            rows = List.of();
        }
        if (templateDef.groupFields().isEmpty()) {
            for (Map<String, Object> row : rows) {
                out.add(OutputRow.detail(row));
                allDetailIndexes.add(out.size() - 1);
            }
        } else {
            allDetailIndexes.addAll(emitGrouped(rows, templateDef, out, spans, 0));
        }
        if (templateDef.totalsTemplateRow() != null) {
            int scopeStart = out.isEmpty() ? 0 : 0;
            int scopeEnd = out.isEmpty() ? -1 : out.size() - 1;
            out.add(OutputRow.total(allDetailIndexes, null, null, true, 0, scopeStart, scopeEnd));
        }
        if (out.isEmpty()) {
            out.add(OutputRow.detail(Map.of()));
        }
        return new BuildResult(out, spans);
    }

    private List<Integer> emitGrouped(
        List<Map<String, Object>> rows,
        TemplateDef templateDef,
        List<OutputRow> out,
        List<GroupSpan> spans,
        int level
    ) {
        List<Integer> allDetailIndexes = new ArrayList<>();
        if (level >= templateDef.groupFields().size()) {
            for (Map<String, Object> row : rows) {
                out.add(OutputRow.detail(row));
                allDetailIndexes.add(out.size() - 1);
            }
            return allDetailIndexes;
        }
        String groupField = templateDef.groupFields().get(level);
        LinkedHashMap<Object, List<Map<String, Object>>> grouped = new LinkedHashMap<>();
        for (Map<String, Object> row : rows) {
            Object key = row.get(groupField);
            grouped.computeIfAbsent(key == null ? "<empty>" : key, k -> new ArrayList<>()).add(row);
        }
        for (var entry : grouped.entrySet()) {
            Object key = entry.getKey();
            List<Map<String, Object>> part = entry.getValue();
            int headerIdx = -1;
            if (templateDef.groupHeaderByField().containsKey(groupField)) {
                out.add(OutputRow.header(groupField, key));
                headerIdx = out.size() - 1;
            }
            int childScopeStart = out.size();
            List<Integer> groupDetailIndexes = emitGrouped(part, templateDef, out, spans, level + 1);
            allDetailIndexes.addAll(groupDetailIndexes);
            int childScopeEnd = out.isEmpty() ? -1 : out.size() - 1;
            int footerIdx = -1;
            if (templateDef.groupFooterByField().containsKey(groupField)) {
                out.add(OutputRow.total(groupDetailIndexes, groupField, key, false, level + 1, childScopeStart, childScopeEnd));
                footerIdx = out.size() - 1;
            }
            if (headerIdx >= 0 && footerIdx > headerIdx + 1) {
                spans.add(new GroupSpan(headerIdx + 1, footerIdx - 1, level + 1));
            }
        }
        return allDetailIndexes;
    }

    private void applyGrouping(Sheet sheet, int startRow, List<GroupSpan> spans) {
        for (GroupSpan span : spans) {
            int from = startRow + span.startOutputIdx();
            int to = startRow + span.endOutputIdx();
            if (span.level() <= 0) {
                continue;
            }
            if (to > from) {
                sheet.groupRow(from, to);
            }
        }
    }

    private void writeRows(
        Sheet sheet,
        int startRow,
        int startCol,
        int endCol,
        TemplateDef templateDef,
        List<OutputRow> outputRows
    ) {
        for (int i = 0; i < outputRows.size(); i++) {
            OutputRow out = outputRows.get(i);
            int targetRowIdx = startRow + i;
            Row target = sheet.getRow(targetRowIdx);
            if (target == null) {
                target = sheet.createRow(targetRowIdx);
            }
            boolean skipTemplateFormulas = out.kind() == RowKind.DETAIL;
            RowSnapshot styleSnapshot = selectTemplateSnapshot(templateDef, out);
            copyTemplateRow(styleSnapshot, target, startCol, endCol, skipTemplateFormulas);

            switch (out.kind()) {
                case DETAIL -> writeDetailRow(target, startCol, endCol, templateDef, out.detailValues());
                case GROUP_HEADER -> writeGroupHeaderRow(target, startCol, templateDef, out.groupField(), out.groupValue());
                case TOTAL -> writeTotalRow(target, startRow, startCol, endCol, templateDef, out, outputRows);
            }
            clearMarkerColumnIfTechnical(target, startCol, templateDef);
        }
    }

    private void clearMarkerColumnIfTechnical(Row target, int startCol, TemplateDef templateDef) {
        Cell marker = target.getCell(startCol);
        if (marker == null || marker.getCellType() != CellType.STRING) {
            return;
        }
        String text = marker.getStringCellValue();
        if (text == null) {
            return;
        }
        String m = text.trim();
        if (m.equalsIgnoreCase("details") || m.equalsIgnoreCase("totals") || parseGroupMarker(m) != null) {
            marker.setBlank();
        }
    }

    private RowSnapshot selectTemplateSnapshot(TemplateDef templateDef, OutputRow out) {
        return switch (out.kind()) {
            case DETAIL -> templateDef.detailsTemplateSnapshot();
            case GROUP_HEADER -> templateDef.groupHeaderSnapshotByField().get(out.groupField());
            case TOTAL -> {
                if (out.rootTotal()) {
                    yield templateDef.totalsTemplateSnapshot();
                }
                RowSnapshot footer = templateDef.groupFooterSnapshotByField().get(out.groupField());
                yield footer != null ? footer : templateDef.totalsTemplateSnapshot();
            }
        };
    }

    private void copyTemplateRow(RowSnapshot src, Row dst, int startCol, int endCol, boolean skipFormulas) {
        if (src == null) {
            return;
        }
        dst.setHeight(src.height());
        for (int c = startCol; c <= endCol; c++) {
            CellSnapshot srcCell = src.cellsByColumn().get(c);
            Cell dstCell = dst.getCell(c);
            if (dstCell == null) {
                dstCell = dst.createCell(c);
            }
            if (srcCell == null) {
                dstCell.setBlank();
                continue;
            }
            CellStyle style = srcCell.style();
            if (style != null) {
                dstCell.setCellStyle(style);
            }
            if (srcCell.cellType() == CellType.FORMULA) {
                if (skipFormulas) {
                    dstCell.setBlank();
                    continue;
                }
                dstCell.setCellFormula(srcCell.formula() == null ? "" : srcCell.formula());
            } else if (srcCell.cellType() == CellType.STRING) {
                dstCell.setCellValue(srcCell.stringValue() == null ? "" : srcCell.stringValue());
            } else if (srcCell.cellType() == CellType.NUMERIC && srcCell.numericValue() != null) {
                dstCell.setCellValue(srcCell.numericValue());
            } else if (srcCell.cellType() == CellType.BOOLEAN && srcCell.boolValue() != null) {
                dstCell.setCellValue(srcCell.boolValue());
            } else {
                dstCell.setBlank();
            }
        }
    }

    private void writeDetailRow(
        Row target,
        int startCol,
        int endCol,
        TemplateDef templateDef,
        Map<String, Object> rowValues
    ) {
        for (int c = startCol; c <= endCol; c++) {
            int idx = c - startCol;
            String field = templateDef.columnFields().get(idx);
            Cell cell = target.getCell(c);
            if (cell == null) {
                cell = target.createCell(c);
            }
            if (field != null) {
                if (shouldSuppressGroupFieldInDetail(field, templateDef)) {
                    cell.setBlank();
                } else {
                    PoiWorkbookOps.setValue(cell, rowValues.get(field));
                }
                continue;
            }
            String templateExpr = templateDef.detailExpressions().get(idx);
            if (templateExpr != null) {
                PoiWorkbookOps.setValue(cell, substitutePlaceholders(templateExpr, rowValues));
            }
        }
    }

    private boolean shouldSuppressGroupFieldInDetail(String field, TemplateDef templateDef) {
        if (templateDef.leaveGroupData() || field == null || field.isBlank()) {
            return false;
        }
        String norm = normalizeFieldName(field, templateDef);
        boolean hasGroupMarker = templateDef.groupHeaderByField().keySet().stream().anyMatch(gf -> gf.equalsIgnoreCase(norm))
            || templateDef.groupFooterByField().keySet().stream().anyMatch(gf -> gf.equalsIgnoreCase(norm));
        if (!hasGroupMarker) {
            return false;
        }
        return templateDef.groupFields().stream().anyMatch(gf -> gf.equalsIgnoreCase(norm));
    }

    private void writeGroupHeaderRow(Row target, int startCol, TemplateDef templateDef, String field, Object value) {
        Integer col = templateDef.fieldToColumn().get(field);
        int headerCol = col != null ? startCol + col : startCol;
        Cell cell = target.getCell(headerCol);
        if (cell == null) {
            cell = target.createCell(headerCol);
        }
        PoiWorkbookOps.setValue(cell, value);
    }

    private void writeTotalRow(
        Row target,
        int outputStartRow,
        int startCol,
        int endCol,
        TemplateDef templateDef,
        OutputRow totalRow,
        List<OutputRow> allRows
    ) {
        List<Integer> detailIndexes = totalRow.detailIndexes();
        boolean rootTotal = totalRow.rootTotal();
        Row directiveRow = rootTotal ? templateDef.totalsTemplateRow() : null;
        if (!rootTotal && !detailIndexes.isEmpty()) {
            // The group footer row style is already copied; totals directives are still read from totals row.
            directiveRow = templateDef.totalsTemplateRow();
        }
        if (directiveRow == null) {
            return;
        }
        for (int c = startCol; c <= endCol; c++) {
            int idx = c - startCol;
            String directive = templateDef.totalDirectives().get(idx);
            if (directive == null || directive.isBlank()) {
                continue;
            }
            Cell cell = target.getCell(c);
            if (cell == null) {
                cell = target.createCell(c);
            }
            if (!isSupportedTotalDirective(directive)) {
                // Keep literal labels from totals template row (e.g. "Итого").
                PoiWorkbookOps.setValue(cell, directive);
                continue;
            }
            Object value = computeTotalValue(
                directive,
                templateDef.columnFields().get(idx),
                detailIndexes,
                outputStartRow,
                c,
                startCol,
                templateDef,
                totalRow,
                allRows
            );
            if (value instanceof String s && s.startsWith("=")) {
                cell.setCellFormula(s.substring(1));
            } else {
                PoiWorkbookOps.setValue(cell, value);
            }
        }
    }

    private Object computeTotalValue(
        String directive,
        String field,
        List<Integer> detailIndexes,
        int outputStartRow,
        int currentCol,
        int startCol,
        TemplateDef templateDef,
        OutputRow totalRow,
        List<OutputRow> allRows
    ) {
        String d = directive.trim().toLowerCase();
        String aggField = field;
        boolean runTotalRequested = false;
        Matcher matcher = TOTAL_DIRECTIVE_PATTERN.matcher(d);
        if (matcher.matches()) {
            d = matcher.group(1).toLowerCase();
            String parsedField = matcher.group(2);
            if (parsedField != null && !parsedField.isBlank()) {
                String pf = parsedField.trim().toLowerCase();
                if ("sum".equals(pf) || "cnt".equals(pf)) {
                    runTotalRequested = true;
                } else {
                    aggField = normalizeFieldName(parsedField, templateDef);
                }
            }
        } else if ("sum(sum)".equals(d)) {
            d = "sum";
            runTotalRequested = true;
        } else if ("cnt(cnt)".equals(d)) {
            d = "cnt";
            runTotalRequested = true;
        }
        if (d.startsWith("=")) {
            return directive;
        }

        Integer fieldOffset = resolveFieldOffset(aggField, templateDef);
        int aggCol = runTotalRequested ? currentCol : (fieldOffset == null ? currentCol : startCol + fieldOffset);
        List<Integer> childTotalIndexes = findChildTotalIndexes(totalRow, allRows);
        boolean cntWithChildTotals = "cnt".equals(d) && !childTotalIndexes.isEmpty();
        List<Integer> sourceIndexes = detailIndexes;
        if ((runTotalRequested || cntWithChildTotals) && !childTotalIndexes.isEmpty()) {
            sourceIndexes = childTotalIndexes;
        }
        List<Integer> excelRows = sourceIndexes.stream().map(i -> outputStartRow + i + 1).sorted().toList();
        if (excelRows.isEmpty()) {
            return 0d;
        }
        String arg = buildFunctionArgs(excelRows, aggCol);
        if (arg.isBlank()) {
            return 0d;
        }
        if ("sum".equals(d)) {
            return "=SUM(" + arg + ")";
        }
        if ("cnt".equals(d)) {
            if (runTotalRequested || cntWithChildTotals) {
                return "=SUM(" + arg + ")";
            }
            return buildRowsCountFormula(excelRows, aggCol);
        }
        if ("avg".equals(d)) {
            return "=AVERAGE(" + arg + ")";
        }
        if ("min".equals(d)) {
            return "=MIN(" + arg + ")";
        }
        if ("max".equals(d)) {
            return "=MAX(" + arg + ")";
        }
        return "";
    }

    private boolean isSupportedTotalDirective(String directive) {
        String d = directive.trim().toLowerCase();
        if (d.startsWith("=")) {
            return true;
        }
        if ("sum(sum)".equals(d) || "cnt(cnt)".equals(d)) {
            return true;
        }
        return TOTAL_DIRECTIVE_PATTERN.matcher(d).matches();
    }

    private List<Integer> findChildTotalIndexes(OutputRow totalRow, List<OutputRow> allRows) {
        List<Integer> result = new ArrayList<>();
        if (totalRow.scopeEnd() < totalRow.scopeStart()) {
            return result;
        }
        for (int i = totalRow.scopeStart(); i <= totalRow.scopeEnd(); i++) {
            OutputRow row = allRows.get(i);
            if (row.kind() == RowKind.TOTAL && row.level() == totalRow.level() + 1) {
                result.add(i);
            }
        }
        return result;
    }

    private Integer resolveFieldOffset(String field, TemplateDef templateDef) {
        if (field == null || field.isBlank()) {
            return null;
        }
        String normalized = normalizeFieldName(field, templateDef);
        return templateDef.fieldToColumn().get(normalized);
    }

    private String buildFunctionArgs(List<Integer> excelRows1Based, int aggColZeroBased) {
        var ranges = buildCompressedRanges(excelRows1Based, aggColZeroBased);
        return String.join(",", ranges);
    }

    private List<String> buildCompressedRanges(List<Integer> excelRows1Based, int aggColZeroBased) {
        if (excelRows1Based.isEmpty()) {
            return List.of();
        }
        List<String> ranges = new ArrayList<>();
        int runStart = excelRows1Based.getFirst();
        int runPrev = runStart;
        for (int i = 1; i < excelRows1Based.size(); i++) {
            int cur = excelRows1Based.get(i);
            if (cur == runPrev + 1) {
                runPrev = cur;
                continue;
            }
            ranges.add(toA1Range(runStart, runPrev, aggColZeroBased));
            runStart = cur;
            runPrev = cur;
        }
        ranges.add(toA1Range(runStart, runPrev, aggColZeroBased));
        return ranges;
    }

    private String buildRowsCountFormula(List<Integer> excelRows1Based, int aggColZeroBased) {
        List<String> ranges = buildCompressedRanges(excelRows1Based, aggColZeroBased);
        if (ranges.isEmpty()) {
            return "0";
        }
        String expr = ranges.stream().map(r -> "ROWS(" + r + ")").reduce((a, b) -> a + "+" + b).orElse("0");
        return "=" + expr;
    }

    private String toA1Range(int startRow1Based, int endRow1Based, int colZeroBased) {
        String col = CellReference.convertNumToColString(colZeroBased);
        if (startRow1Based == endRow1Based) {
            return col + startRow1Based;
        }
        return col + startRow1Based + ":" + col + endRow1Based;
    }

    private String normalizeFieldName(String value, TemplateDef templateDef) {
        if (value == null) {
            return null;
        }
        String raw = value.trim();
        if (raw.isBlank()) {
            return raw;
        }
        String matchedRaw = findKnownFieldIgnoreCase(raw, templateDef);
        if (matchedRaw != null) {
            return matchedRaw;
        }
        int underscore = raw.indexOf('_');
        if (underscore >= 0 && underscore < raw.length() - 1) {
            String stripped = raw.substring(underscore + 1);
            String matchedStripped = findKnownFieldIgnoreCase(stripped, templateDef);
            if (matchedStripped != null) {
                return matchedStripped;
            }
        }
        return raw;
    }

    private String findKnownFieldIgnoreCase(String candidate, TemplateDef templateDef) {
        for (String known : templateDef.fieldToColumn().keySet()) {
            if (known.equalsIgnoreCase(candidate)) {
                return known;
            }
        }
        for (String known : templateDef.groupFields()) {
            if (known.equalsIgnoreCase(candidate)) {
                return known;
            }
        }
        for (String known : templateDef.groupHeaderByField().keySet()) {
            if (known.equalsIgnoreCase(candidate)) {
                return known;
            }
        }
        for (String known : templateDef.groupFooterByField().keySet()) {
            if (known.equalsIgnoreCase(candidate)) {
                return known;
            }
        }
        return null;
    }

    private String substitutePlaceholders(String value, Map<String, Object> rowValues) {
        Matcher m = PLACEHOLDER_PATTERN.matcher(value);
        StringBuilder out = new StringBuilder();
        while (m.find()) {
            String placeholder = m.group();
            int underscore = placeholder.indexOf('_');
            String field = underscore >= 0 ? placeholder.substring(underscore + 1) : placeholder;
            Object replacement = rowValues.get(field);
            m.appendReplacement(out, Matcher.quoteReplacement(String.valueOf(replacement == null ? "" : replacement)));
        }
        m.appendTail(out);
        return out.toString();
    }

    private TemplateDef parseTemplate(
        Sheet sheet,
        int startRow,
        int endRow,
        int startCol,
        int endCol,
        DataSourceConfig ds,
        CompatibilityMode mode
    ) {
        Row detailsRow = null;
        Row totalsRow = null;
        Map<String, Row> headers = new LinkedHashMap<>();
        Map<String, Row> footers = new LinkedHashMap<>();

        for (int r = startRow; r <= endRow; r++) {
            Row row = sheet.getRow(r);
            if (row == null) {
                continue;
            }
            String marker = cellAsText(row.getCell(startCol));
            if (marker == null) {
                continue;
            }
            String m = marker.trim();
            if ("details".equalsIgnoreCase(m)) {
                detailsRow = row;
            } else if ("totals".equalsIgnoreCase(m)) {
                totalsRow = row;
            } else {
                GroupMarker gm = parseGroupMarker(m);
                if (gm != null) {
                    if (gm.header()) {
                        headers.put(gm.fieldName(), row);
                    } else {
                        footers.put(gm.fieldName(), row);
                    }
                }
            }
        }
        if (detailsRow == null) {
            if (mode == CompatibilityMode.STRICT) {
                throw new IllegalArgumentException(
                    "Template region must contain 'details' marker row for range: " + ds.getRangeName()
                );
            }
            // Lenient mode fallback for legacy templates without explicit marker row.
            detailsRow = sheet.getRow(startRow);
        }

        List<String> colFields = new ArrayList<>();
        List<String> detailExpr = new ArrayList<>();
        List<String> totalDirectives = new ArrayList<>();
        Map<String, Integer> fieldToCol = new LinkedHashMap<>();

        for (int c = startCol; c <= endCol; c++) {
            String expr = cellAsExpression(detailsRow.getCell(c));
            if (c == startCol && "details".equalsIgnoreCase(cellAsText(detailsRow.getCell(c)))) {
                expr = "";
            }
            String field = detectColumnField(expr);
            colFields.add(field);
            detailExpr.add(expr);
            if (field != null && !fieldToCol.containsKey(field)) {
                fieldToCol.put(field, c - startCol);
            }
            String directive = totalsRow != null ? cellAsText(totalsRow.getCell(c)) : null;
            if (c == startCol && "totals".equalsIgnoreCase(directive)) {
                directive = "";
            }
            totalDirectives.add(directive == null ? "" : directive);
        }

        List<String> groupFields = resolveGroupFields(headers.keySet(), ds.getFields(), fieldToCol);
        Map<String, RowSnapshot> headerSnapshots = new LinkedHashMap<>();
        for (var e : headers.entrySet()) {
            headerSnapshots.put(e.getKey(), snapshotRow(e.getValue(), startCol, endCol));
        }
        Map<String, RowSnapshot> footerSnapshots = new LinkedHashMap<>();
        for (var e : footers.entrySet()) {
            footerSnapshots.put(e.getKey(), snapshotRow(e.getValue(), startCol, endCol));
        }
        return new TemplateDef(
            detailsRow,
            totalsRow,
            headers,
            footers,
            snapshotRow(detailsRow, startCol, endCol),
            snapshotRow(totalsRow, startCol, endCol),
            headerSnapshots,
            footerSnapshots,
            colFields,
            detailExpr,
            totalDirectives,
            groupFields,
            fieldToCol,
            ds.isLeaveGroupData(),
            endRow - startRow + 1
        );
    }

    private RowSnapshot snapshotRow(Row row, int startCol, int endCol) {
        if (row == null) {
            return null;
        }
        Map<Integer, CellSnapshot> cells = new LinkedHashMap<>();
        for (int c = startCol; c <= endCol; c++) {
            Cell cell = row.getCell(c);
            if (cell == null) {
                continue;
            }
            CellType type = cell.getCellType();
            String formula = null;
            String stringValue = null;
            Double numericValue = null;
            Boolean boolValue = null;
            if (type == CellType.FORMULA) {
                formula = cell.getCellFormula();
            } else if (type == CellType.STRING) {
                stringValue = cell.getStringCellValue();
            } else if (type == CellType.NUMERIC) {
                numericValue = cell.getNumericCellValue();
            } else if (type == CellType.BOOLEAN) {
                boolValue = cell.getBooleanCellValue();
            }
            cells.put(c, new CellSnapshot(cell.getCellStyle(), type, formula, stringValue, numericValue, boolValue));
        }
        return new RowSnapshot(row.getHeight(), cells);
    }

    private List<String> resolveGroupFields(
        java.util.Set<String> markerFields,
        List<FieldDef> fields,
        Map<String, Integer> fieldToCol
    ) {
        if (!markerFields.isEmpty()) {
            return new ArrayList<>(markerFields);
        }
        LinkedHashSet<String> guessed = new LinkedHashSet<>();
        if (fields != null) {
            for (FieldDef f : fields) {
                if (f.getName() != null && fieldToCol.containsKey(f.getName())) {
                    guessed.add(f.getName());
                }
            }
        }
        return new ArrayList<>(guessed);
    }

    private String cellAsExpression(Cell cell) {
        if (cell == null) {
            return "";
        }
        if (cell.getCellType() == CellType.FORMULA) {
            return "=" + cell.getCellFormula();
        }
        return cellAsText(cell);
    }

    private String cellAsText(Cell cell) {
        if (cell == null) {
            return null;
        }
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case FORMULA -> "=" + cell.getCellFormula();
            case NUMERIC -> Double.toString(cell.getNumericCellValue());
            case BOOLEAN -> Boolean.toString(cell.getBooleanCellValue());
            default -> null;
        };
    }

    private String detectColumnField(String expression) {
        if (expression == null || expression.isBlank()) {
            return null;
        }
        Matcher m = PLACEHOLDER_PATTERN.matcher(expression);
        if (!m.find()) {
            return null;
        }
        String placeholder = m.group();
        int underscore = placeholder.indexOf('_');
        return underscore >= 0 ? placeholder.substring(underscore + 1) : placeholder;
    }

    private GroupMarker parseGroupMarker(String marker) {
        Matcher m = GROUP_MARKER_PATTERN.matcher(marker);
        if (!m.matches()) {
            return null;
        }
        String groupType = m.group(1);
        String rhs = m.group(2).trim();
        int underscore = rhs.indexOf('_');
        String field = underscore >= 0 ? rhs.substring(underscore + 1) : rhs;
        if (field.isBlank()) {
            return null;
        }
        return new GroupMarker("groupH".equalsIgnoreCase(groupType), field);
    }

    private void applyGlobalPlaceholders(XSSFWorkbook wb, ReportConfig config) {
        for (Sheet sheet : wb) {
            config.getParams().forEach((k, v) -> PoiWorkbookOps.applyPlaceholder(sheet, "#" + k + "#", v));
            config.getEnvVars().forEach((k, v) -> PoiWorkbookOps.applyPlaceholder(sheet, "#" + k + "#", v));
            PoiWorkbookOps.applyPlaceholder(sheet, "#REPORT_TITLE#", config.getTitle());
            PoiWorkbookOps.applyPlaceholder(sheet, "#REPORT_SUBJECT#", config.getSubject());
            PoiWorkbookOps.applyPlaceholder(sheet, "#FULL_CODE#", config.getFullCode());
        }
        log.debug("Applied placeholders to {} sheet(s).", wb.getNumberOfSheets());
    }

    private record TemplateDef(
        Row detailsTemplateRow,
        Row totalsTemplateRow,
        Map<String, Row> groupHeaderByField,
        Map<String, Row> groupFooterByField,
        RowSnapshot detailsTemplateSnapshot,
        RowSnapshot totalsTemplateSnapshot,
        Map<String, RowSnapshot> groupHeaderSnapshotByField,
        Map<String, RowSnapshot> groupFooterSnapshotByField,
        List<String> columnFields,
        List<String> detailExpressions,
        List<String> totalDirectives,
        List<String> groupFields,
        Map<String, Integer> fieldToColumn,
        boolean leaveGroupData,
        int templateHeight
    ) {
    }

    private record RowSnapshot(
        short height,
        Map<Integer, CellSnapshot> cellsByColumn
    ) {
    }

    private record CellSnapshot(
        CellStyle style,
        CellType cellType,
        String formula,
        String stringValue,
        Double numericValue,
        Boolean boolValue
    ) {
    }

    private record BuildResult(
        List<OutputRow> rows,
        List<GroupSpan> groupSpans
    ) {
    }

    private record GroupMarker(
        boolean header,
        String fieldName
    ) {
    }

    private record OutputRow(
        RowKind kind,
        Map<String, Object> detailValues,
        List<Integer> detailIndexes,
        String groupField,
        Object groupValue,
        boolean rootTotal,
        int level,
        int scopeStart,
        int scopeEnd
    ) {
        static OutputRow detail(Map<String, Object> values) {
            return new OutputRow(RowKind.DETAIL, values, List.of(), null, null, false, -1, -1, -1);
        }

        static OutputRow header(String groupField, Object groupValue) {
            return new OutputRow(RowKind.GROUP_HEADER, Map.of(), List.of(), groupField, groupValue, false, -1, -1, -1);
        }

        static OutputRow total(
            List<Integer> detailIndexes,
            String groupField,
            Object groupValue,
            boolean rootTotal,
            int level,
            int scopeStart,
            int scopeEnd
        ) {
            return new OutputRow(RowKind.TOTAL, Map.of(), detailIndexes, groupField, groupValue, rootTotal, level, scopeStart, scopeEnd);
        }
    }

    private record GroupSpan(
        int startOutputIdx,
        int endOutputIdx,
        int level
    ) {
    }

    private enum RowKind {
        DETAIL,
        GROUP_HEADER,
        TOTAL
    }
}
