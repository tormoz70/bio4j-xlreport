# bio4j-xlreport

Java 21 library for Excel report generation with high compatibility to legacy XML report definitions and Excel templates.

## Stack

- Gradle (Groovy DSL)
- Java 21
- Lombok
- Apache POI
- GraalJS (post-processing scripts)

## Modules

- `engine-core`: config model, XML parser, orchestration API
- `engine-poi`: Apache POI report builder
- `engine-js`: JavaScript post-processing over report object model
- `app-console`: console runner for report build

## Console Runner

- Interactive mode:
  - `gradlew :app-console:run --args="console"`
- Batch mode:
  - `gradlew :app-console:run --args="/rpt:C:/path/report.xml /template:C:/path/template.xlsx /data:C:/path/data.json /mode:strict /out:C:/path/out.xlsx"`

`data.json` example:

```json
{
  "mRng": [
    { "field1": "value", "field2": 10 }
  ]
}
```

Windows shortcut script: `run_as_console.cmd`

Legacy dataset batch example: `run_as_console_01_givc.cmd`

Additional notes: `examples/README-console-01-givc.md`

## JavaScript post-processing

Scripts run after the main report build and receive `report` object:

```javascript
const sheet = report.sheet("Sheet1");
sheet.cell("A1").setValue("Updated by JS");
sheet.groupRows(2, 10);
```

Use `postScripts/script` section in XML or inject script configs programmatically.
