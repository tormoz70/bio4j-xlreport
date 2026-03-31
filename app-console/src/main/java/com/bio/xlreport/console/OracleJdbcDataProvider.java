package com.bio.xlreport.console;

import com.bio.xlreport.core.api.DataProvider;
import com.bio.xlreport.core.model.DataSourceConfig;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Types;
import java.time.Duration;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import lombok.RequiredArgsConstructor;

@RequiredArgsConstructor
final class OracleJdbcDataProvider implements DataProvider {
    private static final Pattern SQL_FILE_PATTERN = Pattern.compile("^\\s*\\{text-file:([^}]+)}\\s*$", Pattern.CASE_INSENSITIVE);
    private static final DateTimeFormatter DT_FMT = DateTimeFormatter.ofPattern("yyyy.MM.dd HH:mm:ss");

    private final String dbUrl;
    private final String dbUser;
    private final String dbPassword;
    private final String dbDriver;
    private final int fetchSize;
    private final Path reportXmlPath;
    private final Map<String, String> bindParams;
    private final List<QueryTrace> traces = Collections.synchronizedList(new ArrayList<>());

    void executeSqlHook(String hookName, String sqlText, int timeoutSeconds) throws Exception {
        if (sqlText == null || sqlText.isBlank()) {
            return;
        }
        String resolved = resolveSql(sqlText);
        NamedSql named = compileNamedSql(resolved);
        Class.forName(dbDriver);
        long startedNs = System.nanoTime();
        try (
            Connection cn = DriverManager.getConnection(dbUrl, dbUser, dbPassword);
            PreparedStatement ps = cn.prepareStatement(named.sql())
        ) {
            bindParams(ps, named.paramOrder(), bindParams);
            ps.setFetchSize(1);
            ps.setQueryTimeout(Math.max(1, timeoutSeconds));
            ps.execute();
            traces.add(new QueryTrace(hookName, resolved, named.paramOrder().size(), (System.nanoTime() - startedNs) / 1_000_000L));
            logMs("SQL hook '" + hookName + "' executed", startedNs);
        } catch (SQLException ex) {
            throw new SQLException(
                "SQL hook failed: " + hookName + ". SQL: " + resolved + ". " + ex.getMessage(),
                ex.getSQLState(),
                ex.getErrorCode(),
                ex
            );
        }
    }

    @Override
    public List<Map<String, Object>> fetch(DataSourceConfig config) throws Exception {
        long startedNs = System.nanoTime();
        String sql = resolveSql(config.getSql());
        NamedSql namedSql = compileNamedSql(sql);
        long maxRows = config.getMaxRowsLimit() > 0 ? config.getMaxRowsLimit() : Long.MAX_VALUE;
        log("DS[" + config.getRangeName() + "] SQL prepared, binds=" + namedSql.paramOrder().size());

        Class.forName(dbDriver);
        long connectStartedNs = System.nanoTime();
        try (
            Connection cn = DriverManager.getConnection(dbUrl, dbUser, dbPassword);
            PreparedStatement ps = cn.prepareStatement(namedSql.sql())
        ) {
            logMs("DS[" + config.getRangeName() + "] open connection+prepare", connectStartedNs);
            ps.setFetchSize(fetchSize);
            if (config.getTimeoutMinutes() > 0) {
                Duration timeout = Duration.ofMinutes(config.getTimeoutMinutes());
                ps.setQueryTimeout((int) Math.min(timeout.getSeconds(), Integer.MAX_VALUE));
            }
            long bindStartedNs = System.nanoTime();
            List<BoundParam> boundParams = bindParams(ps, namedSql.paramOrder(), this.bindParams);
            logMs("DS[" + config.getRangeName() + "] bind params", bindStartedNs);
            long executeStartedNs = System.nanoTime();
            try (ResultSet rs = ps.executeQuery()) {
                logMs("DS[" + config.getRangeName() + "] executeQuery", executeStartedNs);
                long readStartedNs = System.nanoTime();
                List<Map<String, Object>> rows = readRows(rs, maxRows);
                traces.add(
                    new QueryTrace(
                        "fetch:" + config.getRangeName(),
                        namedSql.sql(),
                        namedSql.paramOrder().size(),
                        (System.nanoTime() - executeStartedNs) / 1_000_000L
                    )
                );
                logMs("DS[" + config.getRangeName() + "] readRows count=" + rows.size(), readStartedNs);
                logMs("DS[" + config.getRangeName() + "] total fetch", startedNs);
                return rows;
            } catch (SQLException ex) {
                throw enrichSqlException(ex, config, namedSql.sql(), boundParams);
            }
        }
    }

    public Map<String, String> resolveReportParams(Map<String, String> xmlParams) throws Exception {
        Map<String, String> result = new LinkedHashMap<>();
        if (xmlParams == null || xmlParams.isEmpty()) {
            return result;
        }
        Class.forName(dbDriver);
        long allStartedNs = System.nanoTime();
        try (Connection cn = DriverManager.getConnection(dbUrl, dbUser, dbPassword)) {
            for (var e : xmlParams.entrySet()) {
                String name = e.getKey();
                String expr = e.getValue();
                long oneStartedNs = System.nanoTime();
                String value = resolveScalarStrict(cn, name, expr);
                result.put(name, value);
                logMs("XML param '" + name + "' resolved", oneStartedNs);
            }
        }
        logMs("resolveReportParams total, count=" + result.size(), allStartedNs);
        return result;
    }

    private String resolveSql(String rawSql) throws Exception {
        if (rawSql == null || rawSql.isBlank()) {
            throw new IllegalArgumentException("Data source SQL is empty.");
        }
        Matcher matcher = SQL_FILE_PATTERN.matcher(rawSql);
        if (!matcher.matches()) {
            return rawSql;
        }
        String rel = matcher.group(1).trim();
        Path base = reportXmlPath == null ? Path.of(".") : reportXmlPath.getParent();
        Path sqlPath = (base == null ? Path.of(rel) : base.resolve(rel)).normalize();
        if (!Files.exists(sqlPath)) {
            throw new IllegalArgumentException("SQL file not found: " + sqlPath);
        }
        return readTextWithFallback(sqlPath);
    }

    private String readTextWithFallback(Path path) throws Exception {
        Charset[] charsets = new Charset[] {
            StandardCharsets.UTF_8,
            Charset.forName("windows-1251"),
            Charset.forName("CP866")
        };
        Exception last = null;
        for (Charset cs : charsets) {
            try {
                return Files.readString(path, cs);
            } catch (Exception ex) {
                last = ex;
            }
        }
        throw new IllegalArgumentException("Cannot read SQL file in supported encodings: " + path, last);
    }

    private String resolveScalarStrict(Connection cn, String paramName, String expr) throws Exception {
        if (expr == null) {
            return "";
        }
        String sql = expr.trim();
        if (sql.isEmpty()) {
            return "";
        }
        try {
            NamedSql named = compileNamedSql(sql);
            try (PreparedStatement ps = cn.prepareStatement(named.sql())) {
                bindParams(ps, named.paramOrder(), bindParams);
                ps.setFetchSize(1);
                ps.setQueryTimeout(120);
                try (ResultSet rs = ps.executeQuery()) {
                    if (rs.next()) {
                        Object v = rs.getObject(1);
                        return v == null ? "" : String.valueOf(v);
                    }
                    return "";
                }
            }
        } catch (Exception ex) {
            throw new IllegalArgumentException(
                "Ошибка инициализации параметра отчета: " + paramName + ". SQL: " + sql + ". " + ex.getMessage(),
                ex
            );
        }
    }

    private List<BoundParam> bindParams(PreparedStatement ps, List<String> paramOrder, Map<String, String> params) throws Exception {
        Map<String, String> source = params == null ? Map.of() : params;
        List<BoundParam> bound = new ArrayList<>();
        for (int i = 0; i < paramOrder.size(); i++) {
            String key = paramOrder.get(i);
            String value = source.get(key);
            if (value == null) {
                value = source.get(key.toLowerCase(Locale.ROOT));
            }
            if (value == null) {
                value = source.get(key.toUpperCase(Locale.ROOT));
            }
            int jdbcIndex = i + 1;
            if (value == null) {
                ps.setNull(jdbcIndex, Types.VARCHAR);
                bound.add(new BoundParam(jdbcIndex, key, "NULL(VARCHAR)", null));
            } else {
                bound.add(bindTyped(ps, jdbcIndex, key, value));
            }
        }
        return bound;
    }

    private BoundParam bindTyped(PreparedStatement ps, int idx, String paramName, String rawValue) throws Exception {
        String value = rawValue.trim();
        if (value.isEmpty()) {
            ps.setNull(idx, Types.VARCHAR);
            return new BoundParam(idx, paramName, "NULL(VARCHAR)", null);
        }
        String key = paramName.toLowerCase(Locale.ROOT);

        LocalDate maybeDate = parseLocalDate(value);
        if (maybeDate != null && looksLikeDateParam(key)) {
            ps.setDate(idx, java.sql.Date.valueOf(maybeDate));
            return new BoundParam(idx, paramName, "DATE", maybeDate.toString());
        }

        if (looksLikeNumericParam(key) && isInteger(value)) {
            ps.setLong(idx, Long.parseLong(value));
            return new BoundParam(idx, paramName, "LONG", value);
        }

        ps.setString(idx, value);
        return new BoundParam(idx, paramName, "STRING", value);
    }

    private boolean looksLikeDateParam(String key) {
        return key.contains("date") || key.startsWith("dt");
    }

    private boolean looksLikeNumericParam(String key) {
        if ("sys_curuserroles".equals(key)) {
            return false;
        }
        if (key.endsWith("_number")) {
            return false;
        }
        return key.endsWith("_id")
            || key.endsWith("_uid")
            || key.endsWith("_num")
            || "org_id".equals(key)
            || "holding_id".equals(key)
            || "sys_curodepuid".equals(key);
    }

    private boolean isInteger(String value) {
        int start = value.startsWith("-") ? 1 : 0;
        if (start == value.length()) {
            return false;
        }
        for (int i = start; i < value.length(); i++) {
            if (!Character.isDigit(value.charAt(i))) {
                return false;
            }
        }
        return true;
    }

    private LocalDate parseLocalDate(String value) {
        try {
            return LocalDate.parse(value, DateTimeFormatter.ISO_LOCAL_DATE);
        } catch (DateTimeParseException ignored) {
        }
        try {
            return LocalDate.parse(value, DateTimeFormatter.BASIC_ISO_DATE);
        } catch (DateTimeParseException ignored) {
            return null;
        }
    }

    private SQLException enrichSqlException(
        SQLException ex,
        DataSourceConfig config,
        String sql,
        List<BoundParam> boundParams
    ) {
        StringBuilder msg = new StringBuilder();
        msg.append("SQL failed for range '").append(config.getRangeName()).append("'");
        msg.append(" with ").append(boundParams.size()).append(" bind(s): ");
        for (int i = 0; i < boundParams.size(); i++) {
            BoundParam p = boundParams.get(i);
            if (i > 0) {
                msg.append("; ");
            }
            msg.append("#").append(p.index()).append(" ").append(p.name()).append("=").append(p.type()).append("(");
            msg.append(p.value() == null ? "null" : p.value());
            msg.append(")");
        }
        msg.append(". SQL: ").append(sql);
        return new SQLException(msg.toString(), ex.getSQLState(), ex.getErrorCode(), ex);
    }

    private List<Map<String, Object>> readRows(ResultSet rs, long maxRows) throws Exception {
        ResultSetMetaData meta = rs.getMetaData();
        int cols = meta.getColumnCount();
        List<Map<String, Object>> rows = new ArrayList<>();
        while (rs.next() && rows.size() < maxRows) {
            Map<String, Object> row = new LinkedHashMap<>();
            for (int i = 1; i <= cols; i++) {
                String key = meta.getColumnLabel(i);
                if (key == null || key.isBlank()) {
                    key = meta.getColumnName(i);
                }
                row.put(key.toUpperCase(Locale.ROOT), rs.getObject(i));
            }
            rows.add(row);
        }
        return rows;
    }

    private NamedSql compileNamedSql(String sql) {
        StringBuilder out = new StringBuilder(sql.length());
        List<String> params = new ArrayList<>();
        boolean inString = false;
        boolean inLineComment = false;
        boolean inBlockComment = false;
        for (int i = 0; i < sql.length(); i++) {
            char ch = sql.charAt(i);
            char next = i + 1 < sql.length() ? sql.charAt(i + 1) : '\0';

            if (inLineComment) {
                out.append(ch);
                if (ch == '\n' || ch == '\r') {
                    inLineComment = false;
                }
                continue;
            }
            if (inBlockComment) {
                out.append(ch);
                if (ch == '*' && next == '/') {
                    out.append(next);
                    i++;
                    inBlockComment = false;
                }
                continue;
            }

            if (ch == '\'') {
                out.append(ch);
                if (inString && next == '\'') {
                    out.append(next);
                    i++;
                } else {
                    inString = !inString;
                }
                continue;
            }
            if (!inString && ch == '-' && next == '-') {
                out.append(ch).append(next);
                i++;
                inLineComment = true;
                continue;
            }
            if (!inString && ch == '/' && next == '*') {
                out.append(ch).append(next);
                i++;
                inBlockComment = true;
                continue;
            }
            if (!inString && ch == ':' && i + 1 < sql.length() && isIdentStart(sql.charAt(i + 1))) {
                int j = i + 2;
                while (j < sql.length() && isIdentPart(sql.charAt(j))) {
                    j++;
                }
                String name = sql.substring(i + 1, j);
                params.add(name);
                out.append('?');
                i = j - 1;
                continue;
            }
            out.append(ch);
        }
        return new NamedSql(out.toString(), params);
    }

    private boolean isIdentStart(char c) {
        return Character.isLetter(c) || c == '_';
    }

    private boolean isIdentPart(char c) {
        return Character.isLetterOrDigit(c) || c == '_';
    }

    private record NamedSql(String sql, List<String> paramOrder) {
    }

    private record BoundParam(int index, String name, String type, String value) {
    }

    List<QueryTrace> queryTracesSnapshot() {
        synchronized (traces) {
            return List.copyOf(traces);
        }
    }

    record QueryTrace(
        String stage,
        String sql,
        int bindCount,
        long elapsedMs
    ) {
    }

    private void log(String msg) {
        System.out.println(LocalDateTime.now().format(DT_FMT) + " - " + msg);
    }

    private void logMs(String msg, long startedNs) {
        long ms = (System.nanoTime() - startedNs) / 1_000_000L;
        log(msg + " took " + ms + " ms");
    }
}
