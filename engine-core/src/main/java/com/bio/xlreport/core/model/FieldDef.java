package com.bio.xlreport.core.model;

import lombok.Builder;
import lombok.Value;

@Value
@Builder
public class FieldDef {
    String name;
    String type;
    String align;
    String header;
    String format;
    int width;
}
