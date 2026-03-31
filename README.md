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
  - Oracle mode:
    - `gradlew :app-console:run --args="/rpt:C:/path/report.xml /dbUrl:jdbc:oracle:thin:@host:1521:DB /dbUser:USER /dbPassword:PASS /mode:lenient /out:C:/path/out.xlsx"`
  - Oracle profiles mode (`connectionName` mapping):
    - `gradlew :app-console:run --args="/rpt:C:/path/report.xml /dbProfiles:C:/cfg/db-profiles.json /mode:lenient /out:C:/path/out.xlsx"`
  - Performance SLA + DBA pack:
    - `gradlew :app-console:run --args="/rpt:C:/path/report.xml /dbUrl:... /dbUser:... /dbPassword:... /perfSlaMs:60000 /dbaPackOut:C:/tmp/dba-pack.md"`

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

Legacy `macroBefore/macroAfter/autostart` are mapped to runtime JS hooks via `report.applyLegacyMacro(name)` bridge.

## Legacy baseline/regression

- Oracle baseline: `examples/legacy-baseline-01-givc-oracle.json`
- ClickHouse baseline: `examples/legacy-baseline-01-givc.json`
- Machine-readable regression outputs:
  - `app-console/build/reports/legacy-regression/01_givc_oracle-diff.json`
  - `app-console/build/reports/legacy-regression/01_givc_clickhouse-diff.json`
