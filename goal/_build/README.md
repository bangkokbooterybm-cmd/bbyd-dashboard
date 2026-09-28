# Supervisor Goal Board — public snapshot

`goal/index.html` is an AES-GCM encrypted, view-only copy of the Goal Board artifact
(https://claude.ai/artifact/EaSmdzZLK52a2AcbA7FjpG). No data is stored unencrypted in this repo.

Refresh (every 3 days, scheduled):
1. Download the 4 Google Sheets as xlsx into a folder: staff.xlsx (1yE7oALXvKcKhP3U75C_vzq6BcGSi4XR4JBLApP5W8mM),
   sales.xlsx (1UjO_8Upfe2grXg0F2QcCrtM13md-60st629LD-zh7aE), audit.xlsx (1rqwdmccSqz60tEG039A77ivkc-ZW06fHw2dkO0klunY),
   demand.xlsx (1GI-TUQoQc3p5bHVmA_CsLvz_Ae23mSpdJWxJCO1cQIQ).
2. Save the artifact db (collections config, monthly, stock, team, sync) with ArtifactData out_dir into a folder.
3. `npm i xlsx@0.18.5` then `node goal/_build/run.js <xlsxDir> <dbDir> <password>`.
4. Write `_out/sync_*.json` back to the artifact db (collection `sync`), then commit `goal/index.html` only.
Never commit `_out/` or any downloaded data.
