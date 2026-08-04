package ru.mywayline.xlreport.console;

import static org.junit.jupiter.api.Assertions.assertSame;
import static org.mockito.ArgumentMatchers.any;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.verify;
import static org.mockito.Mockito.verifyNoInteractions;
import static org.mockito.Mockito.when;

import ru.mywayline.xlreport.core.model.DataSourceConfig;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import org.junit.jupiter.api.Test;

class RoutingJdbcDataProviderTest {

    @Test
    void routesFetchToNamedConnectionProvider() throws Exception {
        OracleJdbcDataProvider defaultProvider = mock(OracleJdbcDataProvider.class);
        OracleJdbcDataProvider cub7Provider = mock(OracleJdbcDataProvider.class);
        when(cub7Provider.fetch(any())).thenReturn(List.of());

        Map<String, OracleJdbcDataProvider> byConn = new LinkedHashMap<>();
        byConn.put("cub7", cub7Provider);

        RoutingJdbcDataProvider routing = new RoutingJdbcDataProvider(defaultProvider, byConn);
        DataSourceConfig config = DataSourceConfig.builder()
            .rangeName("mRng")
            .connectionName("cub7")
            .build();

        routing.fetch(config);

        verify(cub7Provider).fetch(config);
        verifyNoInteractions(defaultProvider);
    }

    @Test
    void usesDefaultProviderWhenConnectionMissing() throws Exception {
        OracleJdbcDataProvider defaultProvider = mock(OracleJdbcDataProvider.class);
        when(defaultProvider.fetch(any())).thenReturn(List.of());

        RoutingJdbcDataProvider routing = new RoutingJdbcDataProvider(defaultProvider, Map.of());
        DataSourceConfig config = DataSourceConfig.builder()
            .rangeName("mRng")
            .connectionName("missing")
            .build();

        routing.fetch(config);

        verify(defaultProvider).fetch(config);
    }

    @Test
    void defaultProviderAccessorReturnsConfiguredDefault() {
        OracleJdbcDataProvider defaultProvider = mock(OracleJdbcDataProvider.class);
        RoutingJdbcDataProvider routing = new RoutingJdbcDataProvider(defaultProvider, Map.of());
        assertSame(defaultProvider, routing.defaultProvider());
    }
}
