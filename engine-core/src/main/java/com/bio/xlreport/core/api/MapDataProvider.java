package com.bio.xlreport.core.api;

import com.bio.xlreport.core.model.DataSourceConfig;
import java.util.Collections;
import java.util.List;
import java.util.Map;
import lombok.RequiredArgsConstructor;

@RequiredArgsConstructor
public class MapDataProvider implements DataProvider {
    private final Map<String, List<Map<String, Object>>> recordsByRangeName;

    @Override
    public List<Map<String, Object>> fetch(DataSourceConfig config) {
        return recordsByRangeName.getOrDefault(config.getRangeName(), Collections.emptyList());
    }
}
