# Console Run Example: 01_givc

This example starts the Java console runner for a real legacy XML report definition:

- XML: `c:\data\tmp\rrequest-to-migrate\ekb-cabinet\ekb-rpt\rpts_v2\01_givc\form100(rpt).xml`
- Template: `c:\data\tmp\rrequest-to-migrate\ekb-cabinet\ekb-rpt\rpts_v2\01_givc\form100.xlsx` (set explicitly)
- Data sample: `examples/data-form100.json`
- Output: `out/form100-result.xlsx`

## Quick Start

Run:

`run_as_console_01_givc.cmd`

## Notes

- The sample JSON uses synthetic fields/values for initial pipeline testing.
- If the template is stored elsewhere, update `TEMPLATE_XLSX` in `run_as_console_01_givc.cmd`.
- Real report SQL in `<sql>{text-file:...}</sql>` is not executed automatically in this offline mode.
- Use `mode=lenient` first to maximize compatibility during migration.
- If your environment requires Oracle client binaries, set `ORACLE_HOME` before running.

## Customization

Edit `run_as_console_01_givc.cmd` variables:

- `RPT_XML` - path to target report XML.
- `DATA_JSON` - path to map-based input data.
- `OUT_XLSX` - output workbook path.
- `MODE` - `strict` or `lenient`.
- `RPT_PRMS` - semicolon-separated runtime params (`k=v;k2=v2`).
