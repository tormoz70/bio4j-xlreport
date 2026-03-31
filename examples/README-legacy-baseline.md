# Legacy baseline (01_givc)

This file set defines the acceptance baseline for legacy-parity work.

## Files

- `examples/legacy-baseline-01-givc-oracle.json`: Oracle-oriented subset.
- `examples/legacy-baseline-01-givc.json`: ClickHouse-oriented subset.

## Intended use

1. Run compatibility audit to discover risk areas.
2. Run regression for reports listed in the baseline file.
3. Validate output workbook structure against baseline checks:
   - build success
   - named range integrity
   - grouped-field suppression in details
   - totals formulas
   - SQL-derived header parameters
4. Verify runtime against configured SLA (milliseconds).

## Notes

- Each baseline has:
  - `enabledEnv` (toggle variable)
  - `executionProvider` (`oracle` or other provider id)
- Separate machine-readable reports are generated per baseline:
  - `app-console/build/reports/legacy-regression/<baselineId>-diff.json`
- The subset is intentionally focused on complex/high-risk reports first.
