# Supervisor Goal Board — public snapshot

`goal/index.html` is an AES-GCM encrypted, view-only copy of the Goal Board artifact
(https://claude.ai/artifact/EaSmdzZLK52a2AcbA7FjpG). No data is stored unencrypted in this repo.

Refresh (every 3 days, scheduled):
1. Download the 4 Google Sheets as xlsx into a folder: staff.xlsx (1yE7oALXvKcKhP3U75C_vzq6BcGSi4XR4JBLApP5W8mM),
   sales.xlsx (1UjO_8Upfe2grXg0F2QcCrtM13md-60st629LD-zh7aE), audit.xlsx (1rqwdmccSqz60tEG039A77ivkc-ZW06fHw2dkO0klunY),
   demand.xlsx (1GI-TUQoQc3p5bHVmA_CsLvz_Ae23mSpdJWxJCO1cQIQ).
   A small sheet (demand) comes back inline instead of as a tool-results file: save it with
   `python3 -I goal/_build/extract_inline.py <fileId> <xlsxDir>/demand.xlsx`. Never skip a sheet.
2. Save the artifact db (collections config, monthly, stock, pos, team, sync, passport) with ArtifactData out_dir into a folder.
3. `npm i xlsx@0.18.5` then `node goal/_build/run.js <xlsxDir> <dbDir> <password>`.
4. run.js must not print "WARNING not refreshed". Write `_out/sync_*.json` back to the artifact db (collection `sync`),
   then commit `goal/index.html` only.
Never commit `_out/` or any downloaded data.

SC Passport (`passport/core|evals|items|photos|photos2`) is per-SC data incl. personal sales (approved for Sups 18 Oct 2026); no branch or company totals.

Focus Product sales (`pos/<cutDate>`, uploaded from Odoo in the board's Settings) are published as pieces sold only; snapshot.js turns baht into shares of the 17-branch total.

Focus Product stock locations (where each skin/colour sits, incl. Head Office and other shops) are stored in `stock/wh-<stockDate>` (field `wh`), separate from `stock/<stockDate>` to keep documents small. Save the whole `stock` collection so both come through.

## Odoo Stock / Sales Cut from the central Drive (since 21 Nov 2026)
Focus Product stock and sales no longer need the manual upload in Settings. Each refresh takes the raw Odoo exports from
Drive "BBY Data Update" (id 11jN94wU6n89fwrcxRmyHn_4byKQp1U7e): `Stock\` (1iQY1ojfu-4pSinuVDOul0E9Bf8cG89hQ) + `Form Uploads\Stock (File responses)`
(1NWvANO0ODRlcZvIpSx1MrLJcR-x4g79I1qeDlRqex3_O34CuqJdyiGzRekPc9XZxCBXxJ9M0), and `Sales Cut\` (1Ym3qa0PdRP2jcHSwo_Y_AwWp6tS9cpfd) + `Form Uploads\Sales Cut (File responses)`
(1awGDDrwdt-BZeIvdg7yk0HrWwIwWIJOnweOlLSY63IWyTbdBWPdAjmB3FG00PXMAT9w9_8w0).
- File rule (central rule 7): newest upload time (Drive modifiedTime) across the folders listed above wins, one file for Stock and one for Sales Cut. Same time to the second = stop and ask.
- Stock date: `YYYY-MM-DD` at the front of the file name, otherwise the upload time in Thai time (UTC+7). The raw export carries no date.
- `node drive_odoo.js <stock.xlsx|-> <pos.xlsx|-> <dbDir> <stockDate> [stockName] [posName]` reads the raw columns by name, checks 16,500-19,500 pieces
  and that pieces = unique lots, runs the same parseStock/parsePOS as the Settings upload, and writes `stock/<date>`, `stock/wh-<date>`, `pos/<cutDate>` into `<dbDir>` and `_out/db/`.
  Exit 2 / "ERROR" = check failed: keep the old stock/pos, publish the rest and report it. Write `_out/db/**` back with ArtifactData `set` BEFORE `node run.js` rebuilds the snapshot (or rerun run.js).
- Never read Stock / Sales Cut from the Data Update folder on the computer.
