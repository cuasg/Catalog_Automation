# MSI Catalog Automation

## Operating Rule

After meaningful code edits, test runs, layout decisions, or workflow changes, update this README before ending the session. The goal is that a new Codex session can continue without relying on chat history.

## Current Status - 2026-07-14

This is a Google Apps Script / clasp project for generating MSI catalog price-list PDFs and customer net price files from the master workbook.

Current direction:
- Catalog PDFs are moving to a Google Slides template-based generator.
- The older generated-Google-Sheet PDF function is now retired from the user workflow. The active catalog path is the Slides generator.
- Customer net price files remain native Google Sheets. Automated XLSX conversion has intentionally been dropped.
- Latest local refactor: workflow UI, queued production runner, and catalog/price-list generation controls are all active in the guided workflow.
- Local syntax check status: `node --check 00_Main.js` passes.

Most recent local reference files:
- `MSI_Catalog_Automation_Workbook.xlsx`
- `Malleable-Iron-Fittings-PriceList.pdf`

Important local state:
- There is now a bound-sheet custom menu: `MSI Automation`.
- There is now a guided sidebar UI in `Sidebar.html`.
- The workflow now opens as a wider floating modeless dialog again. The docked Sheets sidebar was too narrow for the workflow surface, so the UI has been moved back to a larger floating container.
- The workflow defaults to the `Overview` tab when opened.
- The sidebar pricing tab now includes a `Price_Rules` editor for group-level, PLC-level, fitting-type-level, item-level, or all-groups increase updates.
- The `Current Rules` list on the pricing tab now supports local search, status filtering, bulk enable/disable on the currently filtered rules, inline `Enable/Disable`, `Edit`, and `Remove` actions for each existing `Price_Rules` row, and a recent rule-activity feed backed by `Generation_Log`. The rule editor itself now also exposes the `Active` state directly.
- The price-rule editor now supports deeper variant drill-down when needed. Depending on the selected product group/scope, rules can be narrowed by finish, grade, schedule, pressure class, or connection family through `Variant` and `Variant_2`.
- Stainless Steel Butt-Weld catalog sections now group by weight (`STD`, then `LW`) while showing `304` and `316` together in the same table. Carbon Steel Butt-Weld table headers now color `STD` as MSI blue and `XH` as silver/black.
- The `Catalogs` tab now supports multi-select product-group queueing. Operators can check multiple catalog groups and queue only that selected set in one production run, instead of running only one group at a time or all active groups.
- The `Price Lists` tab now mirrors the same multi-select queue workflow as `Catalogs`, with checkbox selection, selected-run queueing, and clear/reset controls.
- Catalog PDF generation now aggressively replaces stale slide decks for the same catalog in the `Slides` subfolder before creating the new deck, so old slide files do not accumulate in current output.
- Butt-weld catalog tables now use the shape-table path for better width control. The butt-weld fitting order has been tightened so elbows lead before tees/caps/reducing families, and weight-based headers have extra width to reduce 3-line wrapping.
- The `Price_Rules` workflow now supports optional `Variant` and `Variant_2` drill-down filters so rules can be narrowed by finish, grade, schedule, class, or connection family when that product group needs it.
- The sidebar now includes a separate `Maintenance` tab for queue/admin actions.
- The `Overview` tab is now read-only operational context: workbook metrics, group-level SKU/PLC/rule coverage, last successful price-file timestamps, and last successful catalog timestamps.
- The actionable queue controls for `Refresh Production Status` and `Cancel Active Production Run` now live on `Maintenance`, and the maintenance tab has been moved to the far right of the tab strip.
- The workflow does not auto-open on workbook open. The stable production pattern is: build the `MSI Automation` menu on open, show a short ready-toast, and let the operator launch the workflow manually from `MSI Automation > Open Workflow`.
- The workflow now includes a `Preflight` tab that audits workbook readiness before production and a `Last Production Run` summary section on `Overview`.
- Maintenance now includes `Rerun Failed Jobs`, which requeues only the failed jobs from the most recent production summary when rerunnable failure metadata exists.
- The workflow now also guides the operator more explicitly:
  - `Overview` includes a step-by-step recommended workflow with live status badges.
  - `Catalogs` and `Price Lists` show preflight banners.
  - generation buttons are disabled when preflight has blocking errors, so the UI steers the user back to `Preflight` before they queue output jobs.
- Generation actions from the sidebar now queue through the production runner so the UI can show job-level progress instead of only a blocking spinner.
- Pricing recalculation and generation-prep normalization now queue through the same resumable runner, with persisted cursors so long workbook maintenance can continue after Apps Script timeouts.
- Catalog PDF jobs now also run in resumable internal stages: deck initialization, page-chunk append, terms append, and PDF export.
- Production batching was tuned on 2026-07-02 so the runner continues advancing chunked jobs within the same execution window instead of forcing a one-minute wait between every partial stage. Catalog page append chunks are now 8 pages per pass.
- The sidebar progress bar now uses a live client-side progress model between server polls, active-stage pulse styling, and smoother interpolation. Active polling now refreshes more frequently during live runs, but the UI progress is still an estimate, not a true per-operation server metric.
- The workflow now shows a prominent production safety banner whenever a run is active, warning operators not to push code during that run.
- Operational rule: treat `clasp push` as blocked by default. Before any future push, explicitly confirm that no production run is active. Do not push while a catalog/price-list run is in progress.
- Catalog page-append stages now checkpoint after each appended page within the current 8-page chunk so long runs can show page-level progress instead of waiting for the whole chunk to finish.
- Targeted category runs from the sidebar intentionally do **not** depend on the `Catalog_Groups.Active` toggle.
- Full production runs still respect `Catalog_Groups.Active` plus `Generate_PDF` / `Generate_XLS`.
- Legacy `generateCatalogPDFs()` now throws a deprecation error so the old sheet-based catalog path is not used by mistake.
- Generation timestamps written by the automation now use explicit Central time via `America/Chicago` instead of relying on the script timezone.
- The production progress model was tightened again on 2026-07-14:
  - current-operation elapsed time now stays anchored to the true job start across trigger slices
  - operation ETA and full-run ETA are now derived separately
  - sidebar ETA math now blends live completion ratio with historical durations from `Generation_Log`
  - sidebar status now shows `Current Timing` and `Up Next` instead of echoing the current job as the next job
- The generated table-to-next-section title gap is intended to be `11pt`.
- Local `00_Main.js` has been synced to `11pt` so a future `clasp push` does not revert that manual tune.
- Local Git is now initialized and connected to `https://github.com/cuasg/Catalog_Automation`.
- Git tracking policy:
  - commit Apps Script source, README, `.clasp.json`, and `.gitignore`
  - do **not** commit generated PDFs, workbook snapshots, or local token files
- Latest local workbook inventory:
  - `Catalog_SKUs`: 8,172 SKU rows.
  - `Catalog_Groups`: 23 group rows.
  - `Price_Rules`: populated with generic 0% group rules.
  - Active SKU rows in the local XLSX sample: all 8,172 rows are active.
  - `Malleable Iron Fittings`, `Forged Stainless Steel Fittings`, and `Lead Free Bronze Fittings` naming is now aligned between `Catalog_SKUs` and `Catalog_Groups`.
- Latest workbook refresh inspected at 2026-06-26 4:06 PM local:
  - Active groups have matching SKU rows except the exact local text check for `Carbon Steel Butt-Weld Fittings  `, which has a trailing space in `Catalog_Groups`.
  - The script trims product group names during generation, so the trailing-space row should still generate, but the workbook label should be cleaned when convenient.
- Latest generated `Malleable-Iron-Fittings-PriceList.pdf` is accepted as the current layout baseline.
- Recent catalog logic changes not yet fully revalidated across every family:
  - single-variant catalog tables now use blue `Item Number` / `List Price` headers instead of variant-labeled black headers
  - valve second-variant headers now use MSI green with black text
  - forged steel now splits `A24...` and `A82...` into separate `3000LB` and `2000LB` section flows inferred from item number, instead of mixing both families into one table
  - forged steel uses stricter contiguous packing so the opening elbow sequence is preserved
  - forged stainless now keeps `Threaded` and `Socket-Weld` in separate packing lanes so page consolidation cannot interleave those two flows
  - merchant steel fittings and forged steel fittings have custom fitting-order overrides to push reduction/bushing families later
  - single-variant nipple families can now pack up to 3 sections per page
- Watch `Carbon Steel Butt-Weld Fittings `: the local workbook currently has a trailing space in the product group label. The catalog generator trims names before matching, so this is not blocking, but the workbook should be cleaned when convenient.

## Main Functions

Known working / active functions:
- `onOpen()`: builds the `MSI Automation` menu inside the workbook.
- `showMsiAutomationSidebar()`: opens the guided workflow sidebar.
- `showCatalogRunnerSidebar()`: opens the sidebar focused on catalog actions.
- `showPriceListRunnerSidebar()`: opens the sidebar focused on price-list actions.
- `getMsiAutomationSidebarData()`: returns grouped workflow inventory and production-run status for the sidebar.
- `runMsiAutomationAction(action, productGroup)`: sidebar dispatcher for pricing tools, previews, targeted runs, and full runs.
- `savePriceRuleFromSidebar(ruleInput, shouldRecalculate)`: saves or updates `Price_Rules` rows directly from the sidebar, with optional immediate repricing.
- `calculateCatalogPricing()`: now queues a resumable pricing recalculation job.
- `repairCatalogInnerMasterCartonCounts()`: retained as an internal/helper queue path, but no longer exposed in the operator workflow.
- `calculateCatalogPricingNow_()`: internal one-shot implementation retained for chunk processors.
- `repairCatalogInnerMasterCartonCountsNow_()`: internal one-shot implementation retained for queue prep and chunk processors.
- `setupCatalogTemplateWorkbook()`: appends missing workbook support columns without reordering existing columns.
- `calculateCatalogPricing()`: calculates pair keys, applied increases, calculated list prices, and update flags.
- `repairCatalogInnerMasterCartonCounts()`: fills missing `Inner_Carton` / `Master_Case` values inside confirmed paired variant groups only. It no longer flattens valid differing paired carton counts down to the lower value. This now runs automatically before catalog and price-list generation.
- `generateCatalogPriceFiles()`: creates native Google Sheets customer net price files.
- `previewCatalogSlidePlan()`: dry-runs active catalog PDF groups and logs compact page planning summaries.
- `generateCatalogSlidesPDFs(onlyProductGroup)`: active Slides-based catalog generator.
- `generateMerchantSteelCatalogSlidesPDF()`: convenience wrapper for Merchant Steel testing.
- `generateForgedStainlessCatalogSlidesPDF()`: convenience wrapper for forged stainless testing. Calls `Forged Stainless Steel Fittings`.
- `generateMalleableIronCatalogSlidesPDF()`: targeted combined 150LB/300LB Malleable catalog test.
- `generateMalleableIronPriceFiles()`: targeted Malleable price-file test, respecting configured PLC/file splits.
- `generateCarbonSteelNipplesCatalogSlidesPDF()`: convenience wrapper for carbon steel nipple testing.
- `generateStainlessSteelNipplesCatalogSlidesPDF()`: targeted Stainless Steel Nipples catalog test.
- `previewStainlessSteelNipplesCatalogSlidePlan()`: targeted Stainless Steel Nipples page-plan preview.
- `generateStainlessSteelNipplesPriceFile()`: targeted Stainless Steel Nipples net-price file test.
- `generateValvesCatalogSlidesPDF()`: targeted Valves catalog test.
- `previewValvesCatalogSlidePlan()`: targeted Valves page-plan preview.
- `generateValvesPriceFiles()`: targeted Valves net-price file test.
- Additional explicit category wrappers now exist for all current major catalog families, including Merchant Steel Fittings, Aluminium Fittings/Nipples, Brass Nipples, Bronze Fittings, Lead Free Bronze Fittings, Carbon Steel Flanges, Forged Steel Fittings, Forged Stainless Steel Flanges, Stainless Steel Cast Fittings, Stainless Steel Butt-Weld Fittings, and Carbon Steel Butt-Weld Fittings.

Legacy/reference function:
- `generateCatalogPDFs()`: intentionally retired. It now throws and points users to the Slides workflow.
- `buildCatalogPdfSheet_()`: legacy internal reference only.

Remote execution note:
- `clasp run ...` is not currently usable because the Apps Script project is not deployed/configured as an API executable.
- Run Apps Script functions manually from the Apps Script editor.
- `clasp push` is available and has been used throughout.
- Push safety rule: never push on assumption. Explicitly check workflow production status first and only push after confirming the run is idle.

## Pending UI Follow-Up

- Add an explicit queue confirmation message for full production runs that tells the operator the run has been handed off to Apps Script / Google infrastructure and is safe to leave running after closing the dialog, browser, or local machine.
- Improve progress-bar fidelity further if needed. It is materially better than the earlier version, but it is still partly a client-side model rather than a literal server-side progress stream.
- If stronger push protection is needed later, the real next step is an external `safe-clasp-push` wrapper or another out-of-band status check. Apps Script UI warnings alone cannot technically block a local shell push.

## Drive Hierarchy

Top-level folder:
- `MSI Catalog Automation`: `1sJ_D7rYiLPY76nDB_T8GvyCS1t3BDueq`

Workbook:
- Folder `01_Workbook`: `1N0WW06DMcA3CMwWiU0pPMETjzwIG1Baf`
- Master workbook `MSI_Catalog_Automation_Workbook`: `1dDwWTRVpm1CD1DRkRvbkdbq_vIYxA_X-MKqpzfOxhbc`

Templates:
- Folder `02_Templates`: `16uP3YnUcSF1tAGe-pZ6Xavn7GR3pQY8P`
- Cover template `01_Cover`: `1Aw8BWcvs1nqSlpv3QZ9Y68wrQdlDaMRA92Gmt-6dk58`
- Three-section product page: `1U4LRhGspQQ3z81N2qzXqNdrENRN20GkZB96vSx5OxoI`
- Two-section product page: `146ygE6clIYU_sSuUxgICY7tzDQ3mkSZKmsrxFEUytL4`
- One-section product page: `1Duc9LUa47N1_otgU7vJQX6KdFdTZ8_k5FS1_rWk8tQI`
- Continuation page: `1xesl381YmBcjCCF4HReJZZ2X4yjKwdFCQNmE5VOERBY`
- Terms page: `1CnvXdlmneFPedj0YeaEYFplHChNTnIrUmLVkR7zFvJA`

Generated files:
- Catalog PDFs Current: `1pMRtSCKZBOYkCOVeoY5f8qmAvRJLsrD8`
- Catalog PDFs Archive: `1WbXEUaLQSkjR9irWAVxcjW7OPOFxwiut`
- Net price files Current: `1_BUoqGiwK4cXrogAYLq-4DP9YcsRHSJn`
- Net price files Archive: `1xgHUt5qxI4BxaCI-ngs3X9KCo_XwvAsu`

Assets:
- Folder `06_Assets`: `18bQI51tq21bBU6ASrLKH7EaD2-WZlutQ`
- Logos: `1dmhLET9yTbPRIiqKFhzsrdStmj-DH7v0`
- Product images: `1y6gmTIF4uj0wa4JWrjDVaYpuilQ0i6vr`
- Product group images: `163iMGUWYmCUmSJYn7XTyzaDmU4-PBnY8`
- Icons: `1I74Xnc3ZQwjg2EcqUKRt1YFLtphg2Cq-`
- MSI logo file: `12syK5VFJfjqsyyWqhztQOHCqxJasIA7-`
- Product Image File ID List workbook: `1GBVqxtEZ3RympfLEg-udlK747cc_JSvDUpU2mc3qPWE`

## Workbook Schema

Required sheets:
- `Catalog_SKUs`
- `Catalog_Groups`
- `Price_Rules`
- `Price_History`
- `Generation_Log`
- `Sheet_Definitions`

Do not add extra workflow sheets unless there is a strong reason.

### Catalog_SKUs

Expected columns:
- `Active`
- `Product_Group`
- `Catalog_File_Name`
- `Price_File_Name`
- `PLC`
- `Fitting_Type`
- `Pair_Key_Auto`
- `Variant`
- `Variant_2`
- `Size_1`
- `Size_2`
- `Size_3`
- `Inner_Carton`
- `Master_Case`
- `Item_Number`
- `List_Price`
- `Item_Increase_Pct`
- `Applied_Increase_Pct`
- `Calculated_List_Price`
- `Prior_Published_Price`
- `Needs_Update`
- `Sort_Order`
- `product_image_filename`
- `Image_File_ID`
- `Notes`

Column behavior:
- `Variant` is the primary table-column variant, such as `Black`, `Galvanized`, `304`, or `316`.
- `Variant_2` is an optional secondary grouping variant, such as `150lb`, `300lb`, `Threaded`, or `Socket-Weld`.
- Forged steel currently infers pressure class from item number when the workbook does not carry that distinction in `Variant_2`:
  - `A24...` => `3000LB`
  - `A82...` => `2000LB`
- Whole-number decimal labels such as `304.0` and `316.0` are normalized to `304` and `316` by the script.
- Catalog sections group by `PLC + Fitting_Type + Variant_2`.
- If `Variant_2` is populated, section titles append it, for example `90° Elbows - Threaded`.
- If all sections on a product page share one nonblank `Variant_2`, the product page subtitle is derived from that value.
- For malleable iron, this allows `150lb Threaded Fittings` and `300lb Threaded Fittings` page subtitles.
- `Pair_Key_Auto`, `Applied_Increase_Pct`, `Calculated_List_Price`, and `Needs_Update` are script-generated.
- `Image_File_ID` is used for section product images. Missing images should not stop generation.

### Catalog_Groups

Expected columns:
- `Active`
- `Product_Group`
- `PLC`
- `Catalog_Title`
- `Catalog_Subtitle`
- `Catalog_Subheading`
- `Catalog_File_Name`
- `Price_File_Name`
- `Price_File_Title`
- `Price_File_Subtitle`
- `Catalog_Code`
- `Version_Code`
- `Effective_Date`
- `Product_Group_Image_ID`
- `Key_Bullet_1`
- `Key_Bullet_2`
- `Key_Bullet_3`
- `Key_Bullet_4`
- `Generate_PDF`
- `Generate_XLS`
- `Notes`

Behavior:
- Blank `PLC` means include all active SKUs for the product group.
- Populated `PLC` means include only matching SKU rows.
- Comma-separated PLC values are supported.
- Use one broad catalog row for combined catalog PDFs, and separate PLC-specific rows for split customer price files.

### Sheet_Definitions

Expected columns:
- `Cover_Template_ID`
- `02_Product_Page_Three_Sections`
- `03_Product_Page_Two_Sections`
- `04_Product_Page_One_Section`
- `05_Product_Page_Continuation`
- `Terms_Template_ID`
- `PDF_Output_Folder_ID`
- `PDF_Archive_Folder_ID`
- `Price_File_Output_Folder_ID`
- `Price_File_Archive_Folder_ID`
- `Logo_File_ID`
- `Product_Image_Folder_ID`
- `Product_Group_Image_Folder_ID`

`Sheet_Definitions` owns shared template, folder, logo, and asset IDs.

## Slides Placeholder Contract

Cover placeholders:
- `{{CATALOG_TITLE}}`
- `{{CATALOG_SUBTITLE}}`
- `{{CATALOG_SUBHEADING}}`
- `{{KEY_BULLET_1}}`
- `{{KEY_BULLET_2}}`
- `{{KEY_BULLET_3}}`
- `{{KEY_BULLET_4}}`
- `{{PRODUCT_GROUP_IMAGE}}`
- `{{CATALOG_CODE}}`
- `{{VERSION_CODE}}`
- `{{EFFECTIVE_DATE}}`

Product page placeholders:
- `{{CATALOG_TITLE}}`
- `{{CATALOG_SUBTITLE}}`
- `{{CATALOG_SUBHEADING}}`
- `{{SECTION_1_TITLE}}`
- `{{SECTION_1_PICTURE}}`
- `{{SECTION_1_TABLE}}`
- `{{SECTION_2_TITLE}}`
- `{{SECTION_2_PICTURE}}`
- `{{SECTION_2_TABLE}}`
- `{{SECTION_3_TITLE}}`
- `{{SECTION_3_PICTURE}}`
- `{{SECTION_3_TABLE}}`
- `{{PAGE_NUMBER}}`
- `{{CONTINUED_LABEL}}`

Terms placeholders:
- `{{PAGE_NUMBER}}`

Table placeholders are target areas. The script captures position/size and removes the placeholder. All catalog families use a generated grid of rectangle shapes so row heights and column widths remain exact. Do not switch to per-cell template placeholders.
- Visible blank cells in catalog PDF tables render as `-` unless the entire column is hidden for that section.

## UI Workflow

Primary operator surface:
- Open the workbook.
- Use the `MSI Automation` menu.
- Open the sidebar workflow.

Sidebar workflow tabs:
- `Overview`: workbook metrics, SKU/PLC/group coverage, recent run status, and at-a-glance production context.
- `Pricing`: recalculate pricing and edit `Price_Rules` from the sidebar.
- `Maintenance`: queue/admin actions such as refreshing production status, canceling the active run, and rerunning failed jobs.
- `Price Lists`: run all active price lists or run one product group directly.
- `Catalogs`: run all active catalog PDFs, preview page plans, run one product group directly, or run only the checked product groups.

Behavior:
- The sidebar reads `Catalog_Groups` dynamically and builds the product-group list from the workbook.
- The pricing workflow writes into the existing `Price_Rules` sheet rather than creating a separate rule source.
- Sidebar price-rule scopes support `All Product Groups`, `Product Group`, `PLC`, `Fitting Type`, and `Item Number`.
- For `Product Group`, `PLC`, and `Fitting Type` rule scopes, the sidebar now dynamically exposes optional variant checklists based on matching SKU rows:
  - finish families such as `Black`, `Galvanized`, `Plain (Steel)`, `Zinc Plated`
  - stainless grades such as `304` and `316`
  - schedules / classifications such as `Sch10`, `Sch40`, `Sch80`, `A106GrB`, `150LB`, `300LB`, `Threaded`, or `Socket-Weld`
- Carbon steel nipple rule filters now derive their schedule/type choices from PLC so the rule UI exposes `A106GrB`, `Sch40 Welded`, and `Sch80 Seamless` even when `Variant_2` is blank in the source workbook.
- Price-rule selector options are now generated from a centralized derived filter profile per SKU row, so the UI can expose the correct finish / grade / schedule / class / connection choices even when the workbook stores those distinctions indirectly.
- Flange catalog handling now has two special cases:
  - `Carbon Steel Flanges`: paired `150LB` / `300LB` column headers use MSI blue for `150LB` and a lighter blue for `300LB` instead of black/silver.
  - `Stainless Steel Flanges`: section/page flow is now grouped by pressure class (`150LB`, then `300LB`) while table columns pair `304` and `316` together; page subtitle now follows the class rather than the stainless grade.
- Leaving those optional variant checklists blank keeps the rule broad at the selected scope. Selecting one or more values writes narrower `Price_Rules` rows using `Variant` and/or `Variant_2`.
- PLC-level and item-number-level sidebar rule saves infer the matching product group from `Catalog_SKUs`, so those scopes do not require a visible product-group selector.
- The sidebar rule list is grouped by rule scope and can be filtered independently from the rule-entry form.
- Individual group actions use `Product_Group` targeting and do not require flipping `Active` on and off.
- Full-run actions still use the workbook’s `Active` and `Generate_*` flags.
- The shared queue now supports `catalog_pdf`, `price_file`, `pricing_calc`, and `repair_cartons` job types.
- `catalog_pdf` queue jobs now persist page-stage state, so a large catalog can resume between chunks instead of restarting the whole deck build after timeout.
- `catalog_pdf` queue jobs now also save progress after each appended product page inside a chunk, which gives the workflow more frequent status movement during long Slides runs.
- Pricing recalculation resumes by SKU row window.
- Carton repair resumes by paired-variant group window.
- Carton normalization is now automatic during generation prep. Operators do not need a separate repair step before catalogs or price lists.

## Current Slides Catalog Layout Rules

Page planning:
- Page packing uses a small lookahead window so a nearby fitting type can be pulled forward when that fills a page better.
- Non-nipple catalog packing uses a deeper lookahead window, up to 32 upcoming sections, so tiny later sections can be pulled forward onto earlier two-section or three-section pages when that reduces page count.
- After initial page construction, the planner runs a cleanup pass that tries to merge tiny one-section pages back into earlier compatible two-section pages, promoting them to three-section pages when allowed.
- Page packing still prefers fewer pages, but it is allowed to make local order exceptions to avoid stranded single-section pages.
- The opening fitting sequence is anchored, so early lead categories such as `90-degree elbows`, `45-degree elbows`, `90-degree street elbows`, and `45-degree street elbows` are not skipped by lookahead packing.
- Packing runs in separate `150LB` and `300LB` lanes, so those two families can still be optimized internally without interleaving pages.
- Forged stainless packing also runs in separate `Threaded` and `Socket-Weld` lanes, so optimization cannot mix those two flows in one page sequence.
- Three-section pages allow up to 36 combined rows, with no section over 14 rows.
- Two-section pages allow up to 42 combined rows, with no section over 26 rows.
- Single-section pages allow up to 56 rows.
- More than 56 rendered rows: first one-section page plus continuation pages.
- Oversized sections split in 56-rendered-row chunks.
- After the first full oversized page, continuation chunks stay contiguous and in order; only the final short continuation tail is eligible to share a two-section or three-section page while retaining the `(continued)` label.
- Sections with different page subtitles or secondary variants, such as 150LB/300LB or Threaded/Socket-Weld, are never packed together.
- Forged steel now also uses separate `3000LB` and `2000LB` section/page flows inferred from item number prefixes, so `A24...` and `A82...` families do not mix in the same catalog table.

Current spacing and table sizing:
- Gap from previous generated table bottom to next fitting type title: `11pt`.
- When a section image is taller than its table, the next section is positioned from the lower of the two bottoms so short tables do not overlap the following group.
- Conservative generated table layout estimate: `rowCount * 19`.
- Dense table body font: `6.1pt` when table has more than 10 rendered rows.
- Short table body font: `7.1pt`.
- Dense table row height request: `8.4pt`.
- Short table row height request: `9.4pt`.
- Shape-based rows have fixed dimensions and do not expand based on cell content.

Variant header rules:
- Black/galvanized and plain/zinc tables use one header row.
- Example paired header: `Size`, `Inner Carton`, `Master Carton`, `Black Item Number`, `Black List Price`, `Galvanized Item Number`, `Galvanized List Price`.
- Black/plain item and list price header cells are black with white text.
- Galvanized/zinc/316 item and list price header cells are silver with black text.
- 304/316 tables use one header row, not stacked headers.
- `304 Item Number` and `304 List Price` use MSI blue.
- `316 Item Number` and `316 List Price` use a lighter blue to visually split the 316 side.
- Forged stainless sections combine PLC `340` and `341` into the same fitting table so `304` and `316` appear side by side.
- Forged stainless page flow is `Threaded` sections first, then `Socket-Weld` sections.
- Forged stainless threaded titles suppress the `Threaded` suffix; socket-weld titles render as `- Socket`.
- Forged steel `3000LB` paired tables now collapse `Threaded` and `Socket-Weld` onto one size row when the source data is a true two-row pair for that size.
- Forged steel `A82...` rows are treated as a separate `2000LB` section family, not as extra threaded variants inside the `3000LB` tables.
- Forged steel multi-variant tables use the shape-table path.
- If an entire `Inner Carton` column is blank for a section, the generated Slides table hides that column.
- If an entire `Master Carton` column is blank for a section, the generated Slides table hides that column.
- Forged stainless 304/316 tables pivot by size only, so different 304/316 master carton quantities do not split the same size into duplicate visual rows.
- Compact forged stainless tables split oversized sections at 55 rows per page.
- Weld-o-let/outlet size display uses `Size_1 (Size_2-Size_3)` when all three size fields are populated.
- Compact forged stainless tables keep the `5.5pt` data font and stay inside the original template bounds.
- Shape tables use a fixed `20pt` header and fixed-height data rows, preventing content-driven row expansion.
- Shape-table column widths are calculated per section.
- Paired-variant tables use the full placeholder width; single-variant tables contract toward their natural width and remain left-aligned beside the product image.
- Master Carton and List Price columns remain constrained so Size and 304/316 Item Number columns receive the available space.
- Compact forged stainless size labels retain inch marks while keeping nonbreaking spacing, for example `1/4" x 1/8"`.
- Weld-o-let/outlet compact labels use `Size_1 (Size_2-Size_3)` with inch marks, for example `1/4" (3/8"-36")`.
- Compact forged size labels use non-breaking spaces and hyphens so Slides does not wrap size text and inflate row height.

Column width behavior:
- All catalog tables calculate exact widths without the Google Slides REST API or an Advanced Slides service.

Nipple catalog behavior:
- Any product group containing `Nipples` uses nipple-specific section planning.
- Nipple sections group by schedule/material, then nominal pipe size from `Size_1`.
- Nipple sections can pack up to two nominal pipe sizes per page within the same schedule/material group.
- Nipple packing runs in separate schedule/material lanes, so all `Schedule 40` pages finish before `Schedule 80` begins.
- Within each nipple schedule lane, page packing stays strictly contiguous from the smallest nominal size upward; the planner cannot skip ahead to a later size just to fill a page.
- Table rows use length from `Size_2` / `Size_3`, sorted shortest to longest with `Close` first.
- All nipple families use shape-based tables with fixed row heights and section-specific column widths.
- Nipple carton and price columns are constrained so Length and item-number columns receive more width.
- Carbon steel welded/seamless nipple pages render all Schedule 40 nominal sizes first, then all Schedule 80 nominal sizes.
- Schedule 40 pages use page subtitle `Schedule 40 Welded Nipples`.
- Schedule 80 pages use page subtitle `Schedule 80 Seamless Nipples`.
- Carbon steel A106GrB nipples render as a single `A106GrB` item/list pair and hide `Inner Carton` when all IC values are blank.
- Stainless steel nipples use the same schedule page flow as carbon steel nipples.
- Stainless steel nipple schedule comes from `Variant` (`Schedule 40`, then `Schedule 80`), while table columns come from `Variant_2` (`304`, `316`).
- Nipple tables use smaller text when the table is wide.
- Nipple item-number columns are weighted wider to avoid wrapping long item numbers.

Image behavior:
- Section images come from `Catalog_SKUs.Image_File_ID`.
- Cover image comes from `Catalog_Groups.Product_Group_Image_ID`.
- Missing section images render red `MISSING IMAGE` text.
- Do not fall back to a random image.

## Fitting Type Ordering

Catalog row sorting order:
- `Sort_Order` if populated.
- `PLC`.
- Configured fitting type order.
- Fitting type name fallback.
- `Variant_2`.
- Size fields.
- `Variant`.
- `Item_Number`.

General fitting order:
1. 90-degree elbows
2. 45-degree elbows
3. 90-degree street elbows
4. 45-degree street elbows
5. Side outlet elbows
6. Unions
7. 90-degree reducing elbows
8. Tees
9. Reducing tees
10. Street service tees
11. Side outlet tees
12. Crosses
13. Locknuts
14. Hex bushings
15. Square head plugs
16. Solid caps
17. Full flanges
18. Reducing street elbows
19. Reducing couplings
20. Standard couplings
21. Extension pieces
22. 45-degree Y laterals

Forged fitting order:
1. Full couplings
2. Reducing couplings
3. Half couplings
4. Socket weld inserts
5. Outlets
6. Tank flanges

Forged steel fitting order:
1. 90-degree elbows
2. 45-degree elbows
3. 90-degree street elbows
4. 45-degree street elbows
5. Unions
6. Tees
7. Solid caps
8. Full couplings
9. Half couplings
10. Square head plugs
11. Hex bushings
12. Crosses
13. Outlets
14. Socket weld inserts
15. Tank flanges
16. Reducing couplings

Forged stainless fitting order:
- Threaded sections follow the old stainless forged example flow: 90-degree elbows, 45-degree elbows, 90-degree street elbows, unions, tees, caps, full couplings, half couplings, plugs, hex bushings, reducing couplings, outlets, crosses, inserts.
- Socket-weld sections follow the old stainless forged example flow: tees, 90-degree elbows, 45-degree elbows, reducing couplings, caps, full couplings, unions, outlets, inserts, crosses.
- This forged stainless order is intentionally separate from the generic forged fitting order.

Flange order:
1. Threaded
2. Blind
3. Slip-on
4. Weld neck
5. Socket weld
6. Lap joint

Butt weld fitting order:
1. 90-degree long radius elbows
2. 90-degree short radius elbows
3. 45-degree elbows
4. Caps
5. Tees
6. Reducing tees
7. Concentric reducers
8. Eccentric reducers
9. Long radius return bends
10. Short radius return bends

Valve order:
1. 200 lb HORNET WOG Brass Gate Valves
2. 200 lb HORNET WOG Brass Globe Valves
3. 200 lb WOG Brass Hose Stop
4. 600 lb WOG Brass Full Port Ball Valves
5. 600 lb WOG Brass Full Port LEAD FREE Ball Valves
6. 600 lb WOG Brass Standard Port Ball Valves
7. FPT x FPT Swing Check Valves
8. FPT x FPT Gas Ball Valves
9. MPT x MALE FLARE Gas Ball Valves
10. MALE FLARE x MALE FLARE Gas Ball Valves
11. FEMALE FLARE x MALE FLARE Gas Ball Valves
12. MINI FPT x FPT Ball Valves
13. 200 lb WOG Brass Y-Strainer
14. MINI MPT x FPT Ball Valves
15. 1000 lb WOG 316ss Full Port Ball Valves
16. 2000 lb WOG 316ss Full Port Ball Valves
17. 2000 lb WOG 316ss 3 piece Full Port Ball Valves
18. 200 lb WOG 316ss Gate Valves
19. 200 lb WOG 316ss Swing Check Valves
20. 2000 lb Carbon Steel Full Port Ball Valves
21. 2000 lb Carbon Steel Standard Port Ball Valves
22. Iron FPT x Brass Sweat Dielectric Unions

The script normalizes common workbook label variations:
- `90° Elbows`
- `Street / Service Tee`
- `Couplings / Banded`
- `Couplings / Sockets`
- `Reducer Couplings`
- `Floor Flanges`
- `Weld - o - lets`

## Output Versioning

Before creating a new generated catalog PDF:
- If a matching PDF exists in Current and `PDF_Archive_Folder_ID` is configured, the old file is renamed with ` - archived yyyyMMdd-HHmmss` and moved to Archive.

Before creating a new generated customer price file:
- If a matching spreadsheet exists in Current and `Price_File_Archive_Folder_ID` is configured, the old file is renamed with ` - archived yyyyMMdd-HHmmss` and moved to Archive.

Slides working decks:
- Saved in a `Slides` subfolder under the catalog PDF Current folder.
- Old matching working decks are trashed, not archived.
- The generated PDF remains in the Current catalog PDF folder root.
- Template presentations are cached during generation, and large working decks are checkpointed every six product pages.

Transient Slides retry behavior:
- Slides operations retry up to three times in the current execution with short exponential delays.
- If a transient Slides failure remains, the exact failed catalog file is stored in Script Properties and retried by a one-time trigger in a fresh execution.
- Fresh-execution retries run up to three times with increasing delays.
- Existing Current PDFs are not archived until the replacement deck has completed successfully.

## Current Price File Behavior

`generateCatalogPriceFiles()` creates native Google Sheets files with:
- MSI logo/header.
- Static PLC multiplier input box.
- Item table with `MSI Item #`, `General Description`, `PLC`, `Inner Carton Qty`, `Master Case Qty`, `List Price`, `Net Ea`.
- VLOOKUP-based multiplier formula.
- All price files share the same static header block at the top, then use merged section-description rows between product groups.
- Nipple price files keep one frozen column-header row and insert merged group dividers by schedule, fitting type, nominal size, and material/finish.
- Stainless nipple dividers explicitly separate `304 Stainless` and `316 Stainless` groups.
- All other price files keep one frozen column-header row and use merged dividers by finish, secondary classification, and fitting type.
- Example Malleable divider: `Black | 150LB | 90° Elbows`, followed by sizes from smallest to largest.
- Forged steel price files now infer `3000LB` vs `2000LB` from item number (`A24...` vs `A82...`) and use that class in both section dividers and general descriptions.
- Valve price files split by fitting type and material family so brass, bronze, lead-free bronze, carbon steel, and stainless do not mix in one table.
- Stainless valve grades stay together under the stainless family, while brass, carbon steel, bronze, and lead-free bronze are split into separate sections.
- Valve catalog sections also split by item-number series when needed, so A12 economy/premium rows stay together while A14 rows move to their own section.
- Valve page 3 issue was corrected by splitting A14 item-number series away from the A12 economy/premium brass table.
- Valve catalog tables now pivot by size within a section, so the same nominal size does not repeat across multi-variant rows; if carton counts disagree, the lower count is retained on the merged line.
- Valve catalog sections now also split by connection type from `Variant_2`, so threaded and sweat/CxC rows do not get forced into the same table.
- Valve page packing now keys off the valve material subtitle for compatibility, so split stainless/brass sections can still consolidate onto the same page when they belong under the same material heading.
- Valve packing also allows a very small valve section to share a page across material subtitles when that avoids a stranded near-empty page and still stays visually light.
- If a carton column is blank for an entire price file, the script hides that column instead of printing an empty header and empty cells.
- Price file divider rows are intentionally compact so more valve sections can fit on the same printed page.
- PLC values in the price-file multiplier box and table are written as text, so codes such as `000` retain all three digits.
- The description column auto-resizes from the generated content, with a `330px` minimum width.
- General Description includes every populated size as `Size_1 x Size_2 x Size_3`.
- The carton repair helper only touches black/galvanized, plain/zinc plated, and 304/316 paired variant groups; it will not copy values across unrelated materials.
- Known Malleable 150LB reducer/bushing rows with blank secondary size columns use a guarded item-number suffix decoder. It only fills recognized fitting patterns when the decoded first size matches `Size_1`; explicit workbook sizes remain authoritative.
- Native Google Sheets output only. XLSX conversion is intentionally manual/out of scope for now.

## Scheduled Production Runs

User-facing production functions:
- `startFullCatalogProductionRun()`: queues active catalog PDFs first, then active price files.
- `startFullCatalogSlidesGeneration()`: queues only active catalog PDFs.
- `startFullPriceFileGeneration()`: queues only active price files.
- `getCatalogProductionRunStatus()`: returns/logs the active queue or latest completed summary.
- `cancelCatalogProductionRun()`: removes the active queue and its future trigger.

Production behavior:
- One configured output file runs per Apps Script execution.
- Queue progress is stored in Script Properties and the next job is started by a one-time trigger.
- A watchdog trigger is created before each job so a hard Apps Script timeout can retry the unfinished job.
- Each job receives up to three execution attempts; terminal failures are recorded and the queue continues.
- Calling `generateCatalogSlidesPDFs()` or `generateCatalogPriceFiles()` without a product-group filter starts the corresponding scheduled queue instead of looping through every group in one execution.
- Production safety rule: do not push Apps Script code while a queued production run is active. A push can cause later trigger executions in the same run to pick up different code than earlier stages.

## Known Risks / Current Watch Items

- All catalog tables now use shape-based cells; verify newly tested families for text fit and spacing as coverage expands.
- Multi-section page spacing is currently tuned by conservative estimates, not exact rendered measurement.
- The latest accepted table-to-next-title gap is `11pt`.
- The workflow progress bar is intentionally a blended model. It now combines client-side smoothing with more frequent persisted catalog page checkpoints, so long runs should show less stalling, but it is still not a literal server-side progress feed.
- Current forged stainless focus: verify final threaded-first / socket-weld-second page flow after the packing-lane separation fix.
- Forged stainless should have:
  - no stacked 304/316 overlay row,
  - no blank `Inner Carton` column when all IC values are blank,
  - 30-row oversized section chunks,
  - compact Size labels with inch marks,
  - weld-o-let/outlet labels in `Size_1 (Size_2-Size_3)` format using non-breaking spaces/hyphens.
- Confirm regenerated carbon steel nipple output. It should now show Schedule 40 nominal-size pages first, then Schedule 80 nominal-size pages, not S40/S80 in one table.
- Confirm regenerated stainless steel nipple output. It should show Schedule 40 nominal-size pages first, then Schedule 80, with 304/316 columns in each table.
- Regenerate and inspect `Steel-Butt-Weld-Fittings-PriceList.pdf` before doing any further butt-weld tuning. The latest local code changed the planner so strict butt-weld order should survive post-pack consolidation, but this has not been tested yet.
- If the regenerated butt-weld PDF is still out of order, instrument/log the exact section sequence after initial sort and again after page-packing so the reorder point is proven instead of guessed.
- Current workbook exact-text cleanup item: `Catalog_Groups` still has `Carbon Steel Butt-Weld Fittings  ` with a trailing space. The script trims names during generation, so it is not blocking, but clean the workbook label when convenient.
- `Price_History` is still not populated by automation.
- Archive behavior has been added, but replacement/archive behavior should be watched during the next generation runs.

## Recommended Next Session

1. Review the updated catalog families most affected by the latest header/order/grouping changes:
   - `Valves`
   - `Lead Free Bronze Fittings`
   - `Bronze Fittings`
   - `Forged Steel Fittings`
   - `Forged Stainless Steel Fittings`
   - `Merchant Steel Fittings`
   - `Aluminum Nipples`
2. Keep the butt-weld validation in that review pass:
   - verify the opening sequence starts with `90° Long Radius Elbows`, then `90° Short Radius Elbows`, then `45° Elbows`, then `Caps`, then `Tees`
   - verify the wider `STD` / `XH` list-price headers are holding to two lines where possible
3. Use `previewCatalogSlidePlan()` before a full production run if you want a cheap page-count sanity check across active groups.

## Verification From This Handoff

- Latest local workbook: `MSI_Catalog_Automation_Workbook.xlsx`.
- Latest local PDF references include the current valves output plus the current steel butt-weld example used for review.
- Latest `node --check 00_Main.js` status remains recorded as passing.
- GitHub remote repo: `https://github.com/cuasg/Catalog_Automation`
- Current local state assumption: confirm the workflow production status before any future push because queued Apps Script runs can continue after the UI closes.
- Local `00_Main.js` is intended to be the current source of truth for workflow UI, queued production, automatic carton normalization, slide-deck replacement, current catalog layout rules, page-level progress checkpoints, forged stainless lane separation, and the production push-safety warnings.
