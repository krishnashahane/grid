# Reverse-engineering investigation log

This log records what was checked with REA (Reverse Engineer Anything,
https://github.com/morluto/rea) for GRID. It separates verified findings,
hypotheses and unknowns, and it records nothing that was not observed.

## Tooling status — verified

| Check | Result |
| --- | --- |
| Global REA install | Present at `/opt/homebrew/bin/rea`, a symlink to the globally installed `rea-agents` package. CLI version reported: `6.10.1`. Left unchanged, with no reinstall and no configuration edits. |
| MCP connectivity | Verified with a live tool call: `guided_workflows` (`action: list`) returned the server's six built-in workflows (`investigate_feature`, `compare_application_versions`, `verify_reconstruction`, `trace_crash`, `audit_residual_unknowns`, `prepare_bounded_process_capture`). |

## Targets — verified

- No third-party desktop spreadsheet application bundle was installed on the
  development machine (the `/Applications` directory was checked). So **no
  authorised local binary artifact was available to inspect**, and no binary,
  bundle or runtime analysis was performed.
- No network-hosted spreadsheet product was inspected, and no proprietary code,
  assets or documentation were copied.

## Findings

- **Verified:** none about third-party products. There was nothing authorised to analyse.
- **Hypotheses:** none recorded.
- **Unknowns:** the internal architecture of commercial spreadsheet products
  remains unexamined by this project.

## How GRID's behaviour was established instead

GRID's formula semantics follow publicly documented, widely shared spreadsheet
conventions (operator precedence, the 1900 date-serial system, criteria syntax,
lookup modes, dynamic-array spilling and error values). Each convention GRID
relies on is pinned by an automated test in `tests/`, and every deliberate
difference is listed in [COMPATIBILITY.md](COMPATIBILITY.md).

## Re-running the investigation

If an authorised application artifact becomes available, start with:

1. `inspect_desktop_application` on the app bundle to identify its runtime.
2. `inspect_artifact` / `inspect_macho` for the inventory and linkage.
3. The `investigate_feature` guided workflow for a specific behaviour (for example
   "spill range blocked"), recording observations, inferences and residual
   unknowns separately.

Respect each product's licence and terms. Inspect only artifacts you are
authorised to analyse.
