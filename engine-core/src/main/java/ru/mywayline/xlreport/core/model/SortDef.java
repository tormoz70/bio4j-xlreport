package ru.mywayline.xlreport.core.model;

import lombok.Builder;
import lombok.Value;

@Value
@Builder
public class SortDef {
    String fieldName;
    SortDirection direction;
}
