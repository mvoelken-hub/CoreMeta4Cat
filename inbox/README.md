# Excel Inbox

This folder is the drop-zone for vocabulary workbook contributions via pull request.

## Contributor workflow

1. Download the current vocabulary reference workbook from
   [`docs/assets/coremeta4cat_vocabulary.xlsx`](../docs/assets/coremeta4cat_vocabulary.xlsx).

2. Edit the workbook in Excel. You may add or adjust rows in the data sheets
   (`Synthesis`, `Characterization`, `Reaction`, `Simulation`).
   Do **not** rename sheets or change the column headers.

3. Open a pull request and place your edited file here as:

   ```
   inbox/coremeta4cat_vocabulary.xlsx
   ```

4. The **Excel inbox** GitHub Actions workflow runs automatically and:
   - Validates the workbook structure (sheet names, column headers).
   - Runs a round-trip diff against the current schema and reports any
     differences as a comment on the PR.
   - If validation passes, the file is promoted to `docs/assets/` and the
     inbox copy is cleaned up automatically.
   - If validation fails, the PR is blocked until the issues are resolved.

## What the round-trip check does

The check compares every top-level slot listed in the workbook against the
LinkML schema and reports:

- Slots present in the workbook but missing from the schema
- Slots present in the schema but missing from the workbook
- Mandatory/Recommended/Optional (M/R/O) mismatches

The schema is the ground truth. If your edits require schema changes, please
open a schema issue or include the schema change in the same PR.

## Notes

- Only one file named `coremeta4cat_vocabulary.xlsx` is expected.
- The `inbox/` folder is otherwise empty and tracked by `.gitkeep`.
- Do not place other files in this folder.
