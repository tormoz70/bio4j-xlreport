package com.bio.xlreport.core.api;

import com.bio.xlreport.core.model.DataSourceConfig;
import java.util.List;
import java.util.Map;

public interface DataProvider {
    List<Map<String, Object>> fetch(DataSourceConfig config) throws Exception;
}
