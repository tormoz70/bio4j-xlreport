package ru.mywayline.xlreport.js.api;

import java.util.HashMap;
import java.util.Map;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Workbook-scoped cache for JS-driven cell style changes.
 */
final class CellStyleCache {
    private final Map<String, CellStyle> formatCache = new HashMap<>();
    private final Map<Integer, CellStyle> cloneCache = new HashMap<>();
    private int created = 0;
    private int hits = 0;

    CellStyle styleWithFormat(Workbook workbook, CellStyle baseStyle, String formatString) {
        CellStyle base = baseStyle != null ? baseStyle : workbook.createCellStyle();
        String key = base.getIndex() + "\0" + formatString;
        CellStyle cached = formatCache.get(key);
        if (cached != null) {
            hits++;
            return cached;
        }
        CellStyle style = workbook.createCellStyle();
        created++;
        style.cloneStyleFrom(base);
        style.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat(formatString));
        formatCache.put(key, style);
        return style;
    }

    CellStyle cloneStyle(Workbook workbook, CellStyle sourceStyle) {
        if (sourceStyle == null) {
            CellStyle style = workbook.createCellStyle();
            created++;
            return style;
        }
        int identity = System.identityHashCode(sourceStyle);
        CellStyle cached = cloneCache.get(identity);
        if (cached != null) {
            hits++;
            return cached;
        }
        CellStyle style = workbook.createCellStyle();
        created++;
        style.cloneStyleFrom(sourceStyle);
        cloneCache.put(identity, style);
        return style;
    }

    int createdCount() {
        return created;
    }

    int hitCount() {
        return hits;
    }
}
