package ru.mywayline.xlreport.console;

import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.sql.PreparedStatement;
import java.sql.Types;
import org.junit.jupiter.api.Test;
import org.mockito.ArgumentCaptor;
import org.mockito.Mockito;

class OracleJdbcBindTest {

    @Test
    void bindsVeryLongIntegerAsBigDecimal() throws Exception {
        OracleJdbcDataProvider provider = new OracleJdbcDataProvider(
            "jdbc:oracle:thin:@//localhost:1521/XEPDB1",
            "user",
            "pass",
            "oracle.jdbc.OracleDriver",
            1000,
            null,
            Map.of()
        );
        PreparedStatement ps = Mockito.mock(PreparedStatement.class);
        String longId = "1234567890123456789012345";

        invokeBindTyped(provider, ps, 1, "org_id", longId);

        ArgumentCaptor<BigDecimal> captor = ArgumentCaptor.forClass(BigDecimal.class);
        Mockito.verify(ps).setBigDecimal(Mockito.eq(1), captor.capture());
        assertEquals(new BigDecimal(longId), captor.getValue());
    }

    @Test
    void bindsRegularLongWhenFits() throws Exception {
        OracleJdbcDataProvider provider = new OracleJdbcDataProvider(
            "jdbc:oracle:thin:@//localhost:1521/XEPDB1",
            "user",
            "pass",
            "oracle.jdbc.OracleDriver",
            1000,
            null,
            Map.of()
        );
        PreparedStatement ps = Mockito.mock(PreparedStatement.class);

        invokeBindTyped(provider, ps, 2, "org_id", "12345");

        Mockito.verify(ps).setLong(2, 12345L);
    }

    private static void invokeBindTyped(
        OracleJdbcDataProvider provider,
        PreparedStatement ps,
        int idx,
        String name,
        String value
    ) throws Exception {
        Method method = OracleJdbcDataProvider.class.getDeclaredMethod(
            "bindTyped", PreparedStatement.class, int.class, String.class, String.class
        );
        method.setAccessible(true);
        method.invoke(provider, ps, idx, name, value);
    }
}
