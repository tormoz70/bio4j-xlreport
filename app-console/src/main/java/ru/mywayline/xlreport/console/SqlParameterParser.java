package ru.mywayline.xlreport.console;

import java.util.ArrayList;
import java.util.List;

/**
 * Parses named SQL parameters ({@code :name}) into JDBC positional placeholders.
 */
final class SqlParameterParser {
    record NamedSql(String sql, List<String> paramOrder) {
    }

    NamedSql compile(String sql) {
        StringBuilder out = new StringBuilder(sql.length());
        List<String> params = new ArrayList<>();
        boolean inString = false;
        boolean inQuotedIdentifier = false;
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
            if (ch == '"' && !inString) {
                out.append(ch);
                if (inQuotedIdentifier && next == '"') {
                    out.append(next);
                    i++;
                } else {
                    inQuotedIdentifier = !inQuotedIdentifier;
                }
                continue;
            }
            if (!inString && !inQuotedIdentifier && ch == '-' && next == '-') {
                out.append(ch).append(next);
                i++;
                inLineComment = true;
                continue;
            }
            if (!inString && !inQuotedIdentifier && ch == '/' && next == '*') {
                out.append(ch).append(next);
                i++;
                inBlockComment = true;
                continue;
            }
            if (!inString && !inQuotedIdentifier && ch == ':' && i + 1 < sql.length() && isIdentStart(sql.charAt(i + 1))) {
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
}
