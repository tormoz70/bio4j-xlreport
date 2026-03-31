package com.bio.xlreport.console;

import com.bio.xlreport.core.api.DataProvider;
import com.bio.xlreport.core.model.DataSourceConfig;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

final class RoutingJdbcDataProvider implements DataProvider {
    private final OracleJdbcDataProvider defaultProvider;
    private final Map<String, OracleJdbcDataProvider> providersByConnection;

    RoutingJdbcDataProvider(
        OracleJdbcDataProvider defaultProvider,
        Map<String, OracleJdbcDataProvider> providersByConnection
    ) {
        this.defaultProvider = defaultProvider;
        this.providersByConnection = new LinkedHashMap<>();
        if (providersByConnection != null) {
            providersByConnection.forEach((k, v) -> this.providersByConnection.put(normalize(k), v));
        }
    }

    @Override
    public List<Map<String, Object>> fetch(DataSourceConfig config) throws Exception {
        String connection = config == null ? null : config.getConnectionName();
        OracleJdbcDataProvider provider = providersByConnection.get(normalize(connection));
        if (provider == null) {
            provider = defaultProvider;
        }
        return provider.fetch(config);
    }

    OracleJdbcDataProvider defaultProvider() {
        return defaultProvider;
    }

    private String normalize(String value) {
        if (value == null || value.isBlank()) {
            return "default";
        }
        return value.trim().toLowerCase(Locale.ROOT);
    }
}
