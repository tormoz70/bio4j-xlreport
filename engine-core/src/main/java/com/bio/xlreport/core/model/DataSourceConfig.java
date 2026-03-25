package com.bio.xlreport.core.model;

import java.util.List;
import lombok.Builder;
import lombok.Singular;
import lombok.Value;

@Value
@Builder
public class DataSourceConfig {
    String rangeName;
    String title;
    String connectionName;
    String sql;
    String commandType;
    boolean singleRow;
    @Builder.Default
    boolean leaveGroupData = false;
    int timeoutMinutes;
    long maxRowsLimit;

    @Singular
    List<FieldDef> fields;
    @Singular
    List<SortDef> sorts;
}
