package ru.mywayline.xlreport.console;

import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.Test;

class SqlParameterParserTest {
    private final SqlParameterParser parser = new SqlParameterParser();

    @Test
    void ignoresColonInsideDoubleQuotedIdentifier() {
        var named = parser.compile("SELECT \"Время: 12:00\" AS t FROM demo WHERE id = :id");
        assertEquals("SELECT \"Время: 12:00\" AS t FROM demo WHERE id = ?", named.sql());
        assertEquals(List.of("id"), named.paramOrder());
    }

    @Test
    void ignoresColonInsideAliasWithDoubleQuotes() {
        var named = parser.compile("SELECT col AS \"Name: :value\" FROM t WHERE org_id = :org_id");
        assertEquals("SELECT col AS \"Name: :value\" FROM t WHERE org_id = ?", named.sql());
        assertEquals(List.of("org_id"), named.paramOrder());
    }

    @Test
    void ignoresColonInsideSqlComment() {
        var named = parser.compile("SELECT 1 FROM t -- :ignored\nWHERE id = :id");
        assertEquals("SELECT 1 FROM t -- :ignored\nWHERE id = ?", named.sql());
        assertEquals(List.of("id"), named.paramOrder());
    }

    @Test
    void ignoresEscapedDoubleQuoteInsideIdentifier() {
        var named = parser.compile("SELECT \"a\"\"b: :x\" FROM t WHERE id = :id");
        assertEquals("SELECT \"a\"\"b: :x\" FROM t WHERE id = ?", named.sql());
        assertEquals(List.of("id"), named.paramOrder());
    }

    @Test
    void ignoresColonInsideSingleQuotedString() {
        var named = parser.compile("SELECT 'text :notparam' FROM t WHERE uid = :uid");
        assertEquals("SELECT 'text :notparam' FROM t WHERE uid = ?", named.sql());
        assertEquals(List.of("uid"), named.paramOrder());
    }
}
