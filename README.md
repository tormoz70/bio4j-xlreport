# bio4j-xlreport

[![Java](https://img.shields.io/badge/Java-21-orange?logo=openjdk)](https://openjdk.org/)
[![License](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](LICENSE)
[![Build](https://img.shields.io/badge/build-Gradle-02303A?logo=gradle)](https://gradle.org/)
[![Release](https://img.shields.io/github/v/release/tormoz70/bio4j-xlreport?include_prereleases)](https://github.com/tormoz70/bio4j-xlreport/releases)

**Excel report engine for Java 21** — turn legacy XML report definitions and `.xlsx` templates into filled workbooks.

Built for migrating and modernizing enterprise reporting stacks: Apache POI generation, GraalJS post-processing, JSON or JDBC/Oracle data sources, and a deliberate compatibility path for existing XML report formats.

## Why this project

- **Legacy-friendly**: parse existing XML report configs in `strict` or `lenient` mode instead of rewriting reports from scratch
- **Template-driven**: fill real Excel templates (named ranges, detail blocks, totals) via Apache POI
- **Scriptable**: post-process workbooks with GraalJS (`sheet`, `cell`, grouping, legacy macro bridge)
- **Data-flexible**: map/JSON providers or live Oracle/JDBC (including connection profiles and routing)
- **Operable**: console runner for batch builds, SLA timing, and DBA diagnostic packs

## Stack

| Layer | Tech |
| --- | --- |
| Build | Gradle (Groovy DSL) |
| Runtime | Java 21 |
| Workbooks | Apache POI |
| Scripts | GraalJS |
| Ergonomics | Lombok |

## Modules

| Module | Role |
| --- | --- |
| `engine-core` | Config model, XML parser, orchestration API |
| `engine-poi` | Apache POI report builder |
| `engine-js` | JavaScript post-processing over the report object model |
| `app-console` | Console runner for interactive and batch builds |

## Quick start

### Console (interactive)

```bash
gradlew :app-console:run --args="console"
```

### Console (batch, JSON data)

```bash
gradlew :app-console:run --args="/rpt:C:/path/report.xml /template:C:/path/template.xlsx /data:C:/path/data.json /mode:strict /out:C:/path/out.xlsx"
```

### Oracle / JDBC

```bash
gradlew :app-console:run --args="/rpt:C:/path/report.xml /dbUrl:jdbc:oracle:thin:@host:1521:DB /dbUser:USER /dbPassword:PASS /mode:lenient /out:C:/path/out.xlsx"
```

Connection profiles (`connectionName` mapping):

```bash
gradlew :app-console:run --args="/rpt:C:/path/report.xml /dbProfiles:C:/cfg/db-profiles.json /mode:lenient /out:C:/path/out.xlsx"
```

Performance SLA + DBA pack:

```bash
gradlew :app-console:run --args="/rpt:C:/path/report.xml /dbUrl:... /dbUser:... /dbPassword:... /perfSlaMs:60000 /dbaPackOut:C:/tmp/dba-pack.md"
```

`data.json` example:

```json
{
  "mRng": [
    { "field1": "value", "field2": 10 }
  ]
}
```

Windows helpers:

- `run_as_console.cmd` — generic shortcut
- `run_as_console_01_givc.cmd` — legacy dataset batch example
- `examples/README-console-01-givc.md` — walkthrough

## JavaScript post-processing

Scripts run after the main report build and receive a `report` object:

```javascript
const sheet = report.sheet("Sheet1");
sheet.cell("A1").setValue("Updated by JS");
sheet.groupRows(2, 10);
```

Configure via `postScripts/script` in XML or inject script configs programmatically.

Legacy `macroBefore` / `macroAfter` / `autostart` map to runtime JS hooks through `report.applyLegacyMacro(name)`.

## Legacy baseline / regression

- Oracle baseline: `examples/legacy-baseline-01-givc-oracle.json`
- ClickHouse baseline: `examples/legacy-baseline-01-givc.json`
- Machine-readable diffs:
  - `app-console/build/reports/legacy-regression/01_givc_oracle-diff.json`
  - `app-console/build/reports/legacy-regression/01_givc_clickhouse-diff.json`

## Status

Early public release (`v0.1.0`). APIs and XML compatibility surface may evolve; `lenient` mode is recommended while migrating real report packs.

## License

Licensed under the [Apache License 2.0](LICENSE).
