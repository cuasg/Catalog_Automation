const MSI_WORKFLOW_GROUPS_ = [
  { key: 'merchantSteelCouplings', productGroup: 'Merchant Steel Couplings', label: 'Merchant Steel Couplings' },
  { key: 'merchantSteelFittings', productGroup: 'Merchant Steel Fittings', label: 'Merchant Steel Fittings' },
  { key: 'malleableIronFittings', productGroup: 'Malleable Iron Fittings', label: 'Malleable Iron Fittings' },
  { key: 'carbonSteelNipples', productGroup: 'Carbon Steel Nipples', label: 'Carbon Steel Nipples' },
  { key: 'stainlessSteelNipples', productGroup: 'Stainless Steel Nipples', label: 'Stainless Steel Nipples' },
  { key: 'aluminiumFittings', productGroup: 'Aluminium Fittings', label: 'Aluminium Fittings' },
  { key: 'aluminiumNipples', productGroup: 'Aluminium Nipples', label: 'Aluminium Nipples' },
  { key: 'brassNipples', productGroup: 'Brass Nipples', label: 'Brass Nipples' },
  { key: 'bronzeFittings', productGroup: 'Bronze Fittings', label: 'Bronze Fittings' },
  { key: 'leadFreeBronzeFittings', productGroup: 'Lead Free Bronze Fittings', label: 'Lead Free Bronze Fittings' },
  { key: 'carbonSteelFlanges', productGroup: 'Carbon Steel Flanges', label: 'Carbon Steel Flanges' },
  { key: 'forgedStainlessSteelFittings', productGroup: 'Forged Stainless Steel Fittings', label: 'Forged Stainless Steel Fittings' },
  { key: 'forgedStainlessSteelFlanges', productGroup: 'Forged Stainless Steel Flanges', label: 'Forged Stainless Steel Flanges' },
  { key: 'forgedSteelFittings', productGroup: 'Forged Steel Fittings', label: 'Forged Steel Fittings' },
  { key: 'stainlessSteelCastFittings', productGroup: 'Stainless Steel Cast Fittings', label: 'Stainless Steel Cast Fittings' },
  { key: 'stainlessSteelButtWeldFittings', productGroup: 'Stainless Steel Butt-Weld Fittings', label: 'Stainless Steel Butt-Weld Fittings' },
  { key: 'carbonSteelButtWeldFittings', productGroup: 'Carbon Steel Butt-Weld Fittings', label: 'Carbon Steel Butt-Weld Fittings' },
  { key: 'valves', productGroup: 'Valves', label: 'Valves' }
];

const MSI_TIME_ZONE_ = 'America/Chicago';

function getCatalogTimeZone_() {
  return MSI_TIME_ZONE_;
}

function getCatalogTimestamp_(date) {
  return Utilities.formatDate(date || new Date(), getCatalogTimeZone_(), "yyyy-MM-dd'T'HH:mm:ssZ");
}

function onOpen() {
  buildMsiAutomationMenu_();
  showMsiAutomationReadyToast_();
}

function onInstall() {
  onOpen();
}

function buildMsiAutomationMenu_() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('MSI Automation');
  const pricingMenu = ui.createMenu('Pricing')
    .addItem('Open Workflow', 'showMsiAutomationSidebar')
    .addSeparator()
    .addItem('Calculate Pricing', 'calculateCatalogPricing');

  const maintenanceMenu = ui.createMenu('Maintenance')
    .addItem('Open Workflow', 'showMsiAutomationSidebar');

  const priceListMenu = ui.createMenu('Price Lists')
    .addItem('Open Price List Runner', 'showPriceListRunnerSidebar')
    .addItem('Generate All Active Price Lists', 'startFullPriceFileGeneration');

  const catalogMenu = ui.createMenu('Catalogs')
    .addItem('Open Catalog Runner', 'showCatalogRunnerSidebar')
    .addItem('Generate All Active Catalog PDFs', 'startFullCatalogSlidesGeneration')
    .addItem('Preview All Active Catalog Plans', 'previewCatalogSlidePlan');

  const productionMenu = ui.createMenu('Production')
    .addItem('Open Workflow', 'showMsiAutomationSidebar')
    .addItem('Generate All Active Catalogs + Price Lists', 'startFullCatalogProductionRun')
    .addItem('Production Status', 'getCatalogProductionRunStatus')
    .addItem('Cancel Active Production Run', 'cancelCatalogProductionRun');

  menu
    .addItem('Open Workflow', 'showMsiAutomationSidebar')
    .addSubMenu(pricingMenu)
    .addSubMenu(maintenanceMenu)
    .addSubMenu(priceListMenu)
    .addSubMenu(catalogMenu)
    .addSubMenu(productionMenu)
    .addToUi();
}

function showMsiAutomationSidebar() {
  showMsiWorkflowSidebar_('overview');
}

function showCatalogRunnerSidebar() {
  showMsiWorkflowSidebar_('catalogs');
}

function showPriceListRunnerSidebar() {
  showMsiWorkflowSidebar_('priceLists');
}

function showMsiWorkflowSidebar_(initialTab) {
  showMsiWorkflowDialog_(initialTab, false);
}

function showMsiWorkflowDialog_(initialTab, modal) {
  const template = HtmlService.createTemplateFromFile('Sidebar');
  template.initialTab = String(initialTab || 'overview');
  const output = template
    .evaluate()
    .setTitle('MSI Automation')
    .setWidth(700)
    .setHeight(820);

  if (modal) {
    SpreadsheetApp.getUi().showModalDialog(output, 'MSI Automation');
    return;
  }

  SpreadsheetApp.getUi().showModelessDialog(output, 'MSI Automation');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function showMsiAutomationReadyToast_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;
    ss.toast('MSI Automation ready. Use MSI Automation > Open Workflow.', 'MSI Automation', 8);
  } catch (error) {
    // Toast failure should never block workbook open.
  }
}

function calculateCatalogPricing() {
  return queueExplicitCatalogProductionJobs_([
    { type: 'pricing_calc', fileName: 'Catalog Pricing Recalculation' }
  ], 'Pricing recalculation');
}

function calculateCatalogPricingNow_() {
  const ss = SpreadsheetApp.getActive();
  const skuSheet = ss.getSheetByName('Catalog_SKUs');
  const rulesSheet = ss.getSheetByName('Price_Rules');
  const logSheet = ss.getSheetByName('Generation_Log');
  const startedAt = getCatalogTimestamp_();

  if (!skuSheet || !rulesSheet) {
    throw new Error('Missing Catalog_SKUs or Price_Rules sheet.');
  }

  const skuValues = skuSheet.getDataRange().getValues();
  const ruleValues = rulesSheet.getDataRange().getValues();

  const skuHeaders = skuValues[0];
  const ruleHeaders = ruleValues[0];

  const skuCol = headerMap_(skuHeaders);
  const ruleCol = headerMap_(ruleHeaders);

  const rules = ruleValues.slice(1)
    .filter(row => isTruthy_(row[ruleCol.Active]))
    .filter(row => row[ruleCol.Product_Group]);

  const outputRows = [];

  for (let i = 1; i < skuValues.length; i++) {
    const row = skuValues[i];

    const active = isTruthy_(row[skuCol.Active]);
    if (!active) {
      outputRows.push([
        row[skuCol.Pair_Key_Auto],
        row[skuCol.Applied_Increase_Pct],
        row[skuCol.Calculated_List_Price],
        row[skuCol.Needs_Update]
      ]);
      continue;
    }

    const productGroup = row[skuCol.Product_Group];
    const plc = row[skuCol.PLC];
    const fittingType = row[skuCol.Fitting_Type];
    const itemNumber = row[skuCol.Item_Number];
    const listPrice = row[skuCol.List_Price];
    const itemIncrease = row[skuCol.Item_Increase_Pct];

    const pairKey = buildPairKey_(row, skuCol);
    const appliedIncrease = resolveIncrease_(
      itemIncrease,
      productGroup,
      plc,
      fittingType,
      itemNumber,
      getPricingRuleVariantValueForRow_(row, skuCol),
      getPricingRuleVariant2ValueForRow_(row, skuCol),
      rules,
      ruleCol
    );
    const calculatedPrice = calculateCatalogPriceValue_(listPrice, appliedIncrease);
    const priorPublishedPrice = row[skuCol.Prior_Published_Price];
    const needsUpdate = !catalogPriceValuesEqual_(calculatedPrice, priorPublishedPrice);

    outputRows.push([
      pairKey,
      appliedIncrease,
      calculatedPrice,
      needsUpdate
    ]);
  }

  const startRow = 2;
  skuSheet.getRange(startRow, skuCol.Pair_Key_Auto + 1, outputRows.length, 1)
    .setValues(outputRows.map(r => [r[0]]));

  skuSheet.getRange(startRow, skuCol.Applied_Increase_Pct + 1, outputRows.length, 1)
    .setValues(outputRows.map(r => [r[1]]));

  skuSheet.getRange(startRow, skuCol.Calculated_List_Price + 1, outputRows.length, 1)
    .setValues(outputRows.map(r => [r[2]]));

  skuSheet.getRange(startRow, skuCol.Needs_Update + 1, outputRows.length, 1)
    .setValues(outputRows.map(r => [r[3]]));

  skuSheet.getRange(startRow, skuCol.Applied_Increase_Pct + 1, outputRows.length, 1)
    .setNumberFormat('0.00%');

  skuSheet.getRange(startRow, skuCol.Calculated_List_Price + 1, outputRows.length, 1)
    .setNumberFormat('$#,##0.0000');

  appendCatalogLog_(
    logSheet,
    'Catalog_SKUs',
    'Pricing recalculation',
    'Success',
    Math.max(0, skuValues.length - 1),
    '',
    '',
    `Recalculated pricing across ${Math.max(0, skuValues.length - 1)} SKU rows.`,
    buildProductionLogDetails_(
      'pricing_calc',
      'Catalog_SKUs',
      'Catalog Pricing Recalculation',
      startedAt,
      getCatalogTimestamp_()
    )
  );
}

function repairCatalogInnerMasterCartonCounts() {
  return queueExplicitCatalogProductionJobs_([
    { type: 'repair_cartons', fileName: 'Repair Inner and Master Cartons' }
  ], 'Carton repair');
}

function repairCatalogInnerMasterCartonCountsNow_() {
  const ss = SpreadsheetApp.getActive();
  const skuSheet = ss.getSheetByName('Catalog_SKUs');
  const logSheet = ss.getSheetByName('Generation_Log');
  const startedAt = getCatalogTimestamp_();

  if (!skuSheet) throw new Error('Missing Catalog_SKUs sheet.');

  const values = skuSheet.getDataRange().getValues();
  if (values.length < 2) return { groupsProcessed: 0, fieldsFilled: 0, rowsUpdated: 0 };

  const headers = values[0];
  const col = headerMap_(headers);
  const groupMap = {};

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!isTruthy_(row[col.Active])) continue;
    if (!row[col.Item_Number]) continue;

    const pairKey = buildPairKey_(row, col);
    if (!pairKey) continue;

    if (!groupMap[pairKey]) {
      groupMap[pairKey] = [];
    }

    groupMap[pairKey].push({
      rowIndex: i + 1,
      row,
      variant: getDisplayVariantName_(row[col.Variant] || ''),
      innerCarton: row[col.Inner_Carton],
      masterCase: row[col.Master_Case]
    });
  }

  const updates = [];
  let fieldsFilled = 0;
  let rowsUpdated = 0;

  Object.keys(groupMap).forEach(pairKey => {
    const rows = groupMap[pairKey];
    const family = getCatalogPairedVariantFamily_(rows.map(entry => entry.variant));
    if (!family) return;
    if (!isCatalogAllowedPairedVariantFamily_(family)) return;
    if (rows.length < 2) return;

    const resolvedInner = resolveCatalogPairedCartonValue_(rows.map(entry => entry.innerCarton));
    const resolvedMaster = resolveCatalogPairedCartonValue_(rows.map(entry => entry.masterCase));
    if (!resolvedInner.changed && !resolvedMaster.changed) return;

    rows.forEach(entry => {
      let nextInner = entry.innerCarton;
      let nextMaster = entry.masterCase;
      let touched = false;

      if (resolvedInner.hasValue && String(entry.innerCarton || '').trim() !== String(resolvedInner.value || '').trim()) {
        nextInner = resolvedInner.value;
        touched = true;
        fieldsFilled++;
      }

      if (resolvedMaster.hasValue && String(entry.masterCase || '').trim() !== String(resolvedMaster.value || '').trim()) {
        nextMaster = resolvedMaster.value;
        touched = true;
        fieldsFilled++;
      }

      if (touched) {
        updates.push({
          rowIndex: entry.rowIndex,
          innerCarton: nextInner,
          masterCase: nextMaster
        });
        rowsUpdated++;
      }
    });
  });

  updates.forEach(update => {
    values[update.rowIndex - 1][col.Inner_Carton] = update.innerCarton;
    values[update.rowIndex - 1][col.Master_Case] = update.masterCase;
  });

  if (updates.length) {
    skuSheet.getRange(1, col.Inner_Carton + 1, values.length, 1)
      .setValues(values.map(row => [row[col.Inner_Carton]]));
    skuSheet.getRange(1, col.Master_Case + 1, values.length, 1)
      .setValues(values.map(row => [row[col.Master_Case]]));
  }

  if (logSheet) {
    appendCatalogLog_(
      logSheet,
      'Catalog_SKUs',
      'Repair carton counts',
      'Success',
      rowsUpdated,
      '',
      '',
      `Filled ${fieldsFilled} missing carton fields across ${rowsUpdated} rows in ${Object.keys(groupMap).length} pair groups.`,
      buildProductionLogDetails_(
        'repair_cartons',
        'Catalog_SKUs',
        'Repair Inner and Master Cartons',
        startedAt,
        getCatalogTimestamp_()
      )
    );
  }

  return {
    groupsProcessed: Object.keys(groupMap).length,
    fieldsFilled,
    rowsUpdated
  };
}

function runCatalogPricingCalculationStep_(job) {
  throwCatalogProductionCancelledIfNeeded_(job && job.runId);
  const ss = getCatalogWorkbook_();
  const skuSheet = ss.getSheetByName('Catalog_SKUs');
  const rulesSheet = ss.getSheetByName('Price_Rules');
  const logSheet = ss.getSheetByName('Generation_Log');

  if (!skuSheet || !rulesSheet) {
    throw new Error('Missing Catalog_SKUs or Price_Rules sheet.');
  }

  const skuValues = skuSheet.getDataRange().getValues();
  const ruleValues = rulesSheet.getDataRange().getValues();
  const skuCol = headerMap_(skuValues[0]);
  const ruleCol = headerMap_(ruleValues[0]);
  const rules = ruleValues.slice(1)
    .filter(row => isTruthy_(row[ruleCol.Active]))
    .filter(row => row[ruleCol.Product_Group]);

  const totalRows = Math.max(0, skuValues.length - 1);
  const startIndex = Math.max(1, Number(job.nextDataRowIndex) || 1);
  const chunkSize = 500;
  const endExclusive = Math.min(skuValues.length, startIndex + chunkSize);
  const outputRows = [];

  for (let i = startIndex; i < endExclusive; i++) {
    if ((i - startIndex) % 50 === 0) throwCatalogProductionCancelledIfNeeded_(job && job.runId);
    const row = skuValues[i];
    const active = isTruthy_(row[skuCol.Active]);
    if (!active) {
      outputRows.push([
        row[skuCol.Pair_Key_Auto],
        row[skuCol.Applied_Increase_Pct],
        row[skuCol.Calculated_List_Price],
        row[skuCol.Needs_Update]
      ]);
      continue;
    }

    const productGroup = row[skuCol.Product_Group];
    const plc = row[skuCol.PLC];
    const fittingType = row[skuCol.Fitting_Type];
    const itemNumber = row[skuCol.Item_Number];
    const listPrice = row[skuCol.List_Price];
    const itemIncrease = row[skuCol.Item_Increase_Pct];
    const pairKey = buildPairKey_(row, skuCol);
    const appliedIncrease = resolveIncrease_(
      itemIncrease,
      productGroup,
      plc,
      fittingType,
      itemNumber,
      getPricingRuleVariantValueForRow_(row, skuCol),
      getPricingRuleVariant2ValueForRow_(row, skuCol),
      rules,
      ruleCol
    );
    const calculatedPrice = calculateCatalogPriceValue_(listPrice, appliedIncrease);
    const priorPublishedPrice = row[skuCol.Prior_Published_Price];
    const needsUpdate = !catalogPriceValuesEqual_(calculatedPrice, priorPublishedPrice);

    outputRows.push([pairKey, appliedIncrease, calculatedPrice, needsUpdate]);
  }

  if (outputRows.length) {
    skuSheet.getRange(startIndex + 1, skuCol.Pair_Key_Auto + 1, outputRows.length, 1)
      .setValues(outputRows.map(r => [r[0]]));
    skuSheet.getRange(startIndex + 1, skuCol.Applied_Increase_Pct + 1, outputRows.length, 1)
      .setValues(outputRows.map(r => [r[1]]));
    skuSheet.getRange(startIndex + 1, skuCol.Calculated_List_Price + 1, outputRows.length, 1)
      .setValues(outputRows.map(r => [r[2]]));
    skuSheet.getRange(startIndex + 1, skuCol.Needs_Update + 1, outputRows.length, 1)
      .setValues(outputRows.map(r => [r[3]]));
  }

  const nextDataRowIndex = endExclusive;
  const completed = nextDataRowIndex >= skuValues.length;
  const updatedJob = Object.assign({}, job, {
    totalRows,
    nextDataRowIndex,
    processedRows: Math.min(totalRows, Math.max(0, nextDataRowIndex - 1)),
    progressAnchorAt: getCatalogTimestamp_()
  });

  if (completed) {
    const startRow = 2;
    if (totalRows > 0) {
      skuSheet.getRange(startRow, skuCol.Applied_Increase_Pct + 1, totalRows, 1).setNumberFormat('0.00%');
      skuSheet.getRange(startRow, skuCol.Calculated_List_Price + 1, totalRows, 1).setNumberFormat('$#,##0.0000');
    }

    appendCatalogLog_(
      logSheet,
      'Catalog_SKUs',
      'Pricing recalculation',
      'Success',
      totalRows,
      '',
      '',
      `Recalculated pricing across ${totalRows} SKU rows.`,
      buildProductionLogDetails_(
        'pricing_calc',
        'Catalog_SKUs',
        String(job.fileName || 'Catalog Pricing Recalculation'),
        job.startedAt || '',
        getCatalogTimestamp_(),
        { runId: job.runId || '' }
      )
    );
  }

  return {
    completed,
    job: updatedJob
  };
}

function runCatalogCartonRepairStep_(job) {
  throwCatalogProductionCancelledIfNeeded_(job && job.runId);
  const ss = getCatalogWorkbook_();
  const skuSheet = ss.getSheetByName('Catalog_SKUs');
  const logSheet = ss.getSheetByName('Generation_Log');
  if (!skuSheet) throw new Error('Missing Catalog_SKUs sheet.');

  const values = skuSheet.getDataRange().getValues();
  if (values.length < 2) {
    return { completed: true, job: Object.assign({}, job, { totalGroups: 0, nextGroupIndex: 0, rowsUpdated: 0, fieldsFilled: 0 }) };
  }

  const headers = values[0];
  const col = headerMap_(headers);
  const groupMap = {};

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!isTruthy_(row[col.Active])) continue;
    if (!row[col.Item_Number]) continue;

    const pairKey = buildPairKey_(row, col);
    if (!pairKey) continue;

    if (!groupMap[pairKey]) groupMap[pairKey] = [];
    groupMap[pairKey].push({
      rowIndex: i + 1,
      variant: getDisplayVariantName_(row[col.Variant] || ''),
      innerCarton: row[col.Inner_Carton],
      masterCase: row[col.Master_Case]
    });
  }

  const pairKeys = Object.keys(groupMap).sort((a, b) => a.localeCompare(b));
  const totalGroups = pairKeys.length;
  const startGroupIndex = Math.max(0, Number(job.nextGroupIndex) || 0);
  const chunkSize = 300;
  const endGroupIndex = Math.min(totalGroups, startGroupIndex + chunkSize);
  const updates = [];
  let fieldsFilled = Number(job.fieldsFilled) || 0;
  let rowsUpdated = Number(job.rowsUpdated) || 0;

  for (let index = startGroupIndex; index < endGroupIndex; index++) {
    if ((index - startGroupIndex) % 25 === 0) throwCatalogProductionCancelledIfNeeded_(job && job.runId);
    const rows = groupMap[pairKeys[index]];
    const family = getCatalogPairedVariantFamily_(rows.map(entry => entry.variant));
    if (!family) continue;
    if (!isCatalogAllowedPairedVariantFamily_(family)) continue;
    if (rows.length < 2) continue;

    const resolvedInner = resolveCatalogPairedCartonValue_(rows.map(entry => entry.innerCarton));
    const resolvedMaster = resolveCatalogPairedCartonValue_(rows.map(entry => entry.masterCase));
    if (!resolvedInner.changed && !resolvedMaster.changed) continue;

    rows.forEach(entry => {
      let nextInner = entry.innerCarton;
      let nextMaster = entry.masterCase;
      let touched = false;

      if (resolvedInner.hasValue && String(entry.innerCarton || '').trim() !== String(resolvedInner.value || '').trim()) {
        nextInner = resolvedInner.value;
        touched = true;
        fieldsFilled++;
      }

      if (resolvedMaster.hasValue && String(entry.masterCase || '').trim() !== String(resolvedMaster.value || '').trim()) {
        nextMaster = resolvedMaster.value;
        touched = true;
        fieldsFilled++;
      }

      if (touched) {
        updates.push({ rowIndex: entry.rowIndex, innerCarton: nextInner, masterCase: nextMaster });
        rowsUpdated++;
      }
    });
  }

  updates.forEach(update => {
    skuSheet.getRange(update.rowIndex, col.Inner_Carton + 1).setValue(update.innerCarton);
    skuSheet.getRange(update.rowIndex, col.Master_Case + 1).setValue(update.masterCase);
  });

  const nextGroupIndex = endGroupIndex;
  const completed = nextGroupIndex >= totalGroups;
  const updatedJob = Object.assign({}, job, {
    totalGroups,
    nextGroupIndex,
    rowsUpdated,
    fieldsFilled,
    progressAnchorAt: getCatalogTimestamp_()
  });

  if (completed) {
    appendCatalogLog_(
      logSheet,
      'Catalog_SKUs',
      'Repair carton counts',
      'Success',
      rowsUpdated,
      '',
      '',
      `Filled ${fieldsFilled} carton fields across ${rowsUpdated} rows in ${totalGroups} pair groups.`,
      buildProductionLogDetails_(
        'repair_cartons',
        'Catalog_SKUs',
        String(job.fileName || 'Normalize Inner and Master Cartons'),
        job.startedAt || '',
        getCatalogTimestamp_(),
        { runId: job.runId || '' }
      )
    );
  }

  return {
    completed,
    job: updatedJob
  };
}

function resolveCatalogPairedCartonValue_(values) {
  const entries = (values || []).map(value => ({
    raw: value,
    trimmed: String(value || '').trim()
  }));
  const populatedEntries = entries.filter(entry => entry.trimmed !== '');
  const uniquePopulatedValues = [...new Set(populatedEntries.map(entry => entry.trimmed))];
  const missingCount = entries.length - populatedEntries.length;
  const fillValue = populatedEntries.length ? populatedEntries[0].raw : '';

  return {
    hasValue: populatedEntries.length > 0,
    value: fillValue,
    changed: !!(populatedEntries.length && missingCount > 0),
    hasVariantConflict: uniquePopulatedValues.length > 1
  };
}

function resolveIncrease_(itemIncrease, productGroup, plc, fittingType, itemNumber, variant, variant2, rules, ruleCol) {
  const parsedItemIncrease = parsePercent_(itemIncrease);
  if (parsedItemIncrease !== null) return parsedItemIncrease;

  const itemRule = findRule_(rules, ruleCol, productGroup, '', '', itemNumber, variant, variant2);
  if (itemRule !== null) return itemRule;

  const fittingRule = findRule_(rules, ruleCol, productGroup, '', fittingType, '', variant, variant2);
  if (fittingRule !== null) return fittingRule;

  const plcRule = findRule_(rules, ruleCol, productGroup, plc, '', '', variant, variant2);
  if (plcRule !== null) return plcRule;

  const groupRule = findRule_(rules, ruleCol, productGroup, '', '', '', variant, variant2);
  if (groupRule !== null) return groupRule;

  return 0;
}

function findRule_(rules, ruleCol, productGroup, plc, fittingType, itemNumber, variant, variant2) {
  let bestMatch = null;
  let bestSpecificity = -1;

  for (const row of rules) {
    if (String(row[ruleCol.Product_Group]) !== String(productGroup)) continue;
    if (String(row[ruleCol.PLC] || '') !== String(plc || '')) continue;
    if (String(row[ruleCol.Fitting_Type] || '') !== String(fittingType || '')) continue;
    if (String(row[ruleCol.Item_Number] || '') !== String(itemNumber || '')) continue;
    if (!doesRuleVariantMatch_(row, ruleCol, 'Variant', variant)) continue;
    if (!doesRuleVariantMatch_(row, ruleCol, 'Variant_2', variant2)) continue;

    const specificity = getRuleVariantSpecificity_(row, ruleCol);
    if (specificity <= bestSpecificity) continue;

    bestMatch = row;
    bestSpecificity = specificity;
  }

  return bestMatch ? parsePercent_(bestMatch[ruleCol.Increase_Pct]) : null;
}

function doesRuleVariantMatch_(row, ruleCol, headerName, candidateValue) {
  if (ruleCol[headerName] === undefined) return true;

  const ruleValue = normalizeMatchValue_(row[ruleCol[headerName]]);
  if (!ruleValue) return true;

  return ruleValue === normalizeMatchValue_(candidateValue);
}

function getRuleVariantSpecificity_(row, ruleCol) {
  let specificity = 0;

  if (ruleCol.Variant !== undefined && normalizeMatchValue_(row[ruleCol.Variant])) {
    specificity += 1;
  }

  if (ruleCol.Variant_2 !== undefined && normalizeMatchValue_(row[ruleCol.Variant_2])) {
    specificity += 1;
  }

  return specificity;
}

function getCatalogPairedVariantFamily_(variants) {
  const normalized = (variants || [])
    .map(value => String(value || '').trim().toLowerCase())
    .filter(Boolean)
    .filter((value, index, array) => array.indexOf(value) === index);

  if (!normalized.length || normalized.length > 2) return '';

  const hasBlack = normalized.includes('black');
  const hasGalvanized = normalized.includes('galvanized');
  const hasPlain = normalized.includes('plain (steel)') || normalized.includes('steel') || normalized.includes('plain');
  const hasZinc = normalized.includes('zinc plated') || normalized.includes('zinc plated');
  const has304 = normalized.includes('304');
  const has316 = normalized.includes('316');

  if (hasBlack && hasGalvanized) return 'black/galvanized';
  if (hasPlain && hasZinc) return 'plain/zinc plated';
  if (has304 && has316) return '304/316';

  return '';
}

function isCatalogAllowedPairedVariantFamily_(family) {
  return ['black/galvanized', 'plain/zinc plated', '304/316'].includes(String(family || '').toLowerCase());
}

function buildPairKey_(row, col) {
  const sizeParts = getCatalogRowSizeParts_(row, col);
  return [
    row[col.Product_Group],
    row[col.Fitting_Type],
    getOptionalCellValue_(row, col, 'Variant_2'),
    sizeParts[0],
    sizeParts[1],
    sizeParts[2]
  ].filter(value => value !== '' && value !== null).join(' | ');
}

function parsePercent_(value) {
  if (value === '' || value === null || value === undefined) return null;

  if (typeof value === 'number') return value;

  const cleaned = String(value).replace('%', '').trim();
  if (cleaned === '') return null;

  const number = Number(cleaned);
  if (isNaN(number)) return null;

  return number > 1 ? number / 100 : number;
}

function headerMap_(headers) {
  return headers.reduce((map, header, index) => {
    map[String(header).trim()] = index;
    return map;
  }, {});
}

function setupCatalogTemplateWorkbook() {
  const ss = SpreadsheetApp.getActive();

  ensureSheetColumns_(ss, 'Catalog_Groups', [
    'Active',
    'Product_Group',
    'PLC',
    'Catalog_Title',
    'Catalog_Subtitle',
    'Catalog_Subheading',
    'Catalog_File_Name',
    'Price_File_Name',
    'Price_File_Title',
    'Price_File_Subtitle',
    'Catalog_Code',
    'Version_Code',
    'Effective_Date',
    'Product_Group_Image_ID',
    'Key_Bullet_1',
    'Key_Bullet_2',
    'Key_Bullet_3',
    'Key_Bullet_4',
    'Generate_PDF',
    'Generate_XLS',
    'Notes'
  ]);

  ensureSheetColumns_(ss, 'Sheet_Definitions', [
    'Cover_Template_ID',
    '02_Product_Page_Three_Sections',
    '03_Product_Page_Two_Sections',
    '04_Product_Page_One_Section',
    '05_Product_Page_Continuation',
    'Terms_Template_ID',
    'PDF_Output_Folder_ID',
    'PDF_Archive_Folder_ID',
    'Price_File_Output_Folder_ID',
    'Price_File_Archive_Folder_ID',
    'Logo_File_ID',
    'Product_Image_Folder_ID',
    'Product_Group_Image_Folder_ID'
  ]);

  ensureSheetColumns_(ss, 'Catalog_SKUs', [
    'Active',
    'Product_Group',
    'Catalog_File_Name',
    'Price_File_Name',
    'PLC',
    'Fitting_Type',
    'Pair_Key_Auto',
    'Variant',
    'Variant_2',
    'Size_1',
    'Size_2',
    'Size_3',
    'Inner_Carton',
    'Master_Case',
    'Item_Number',
    'List_Price',
    'Item_Increase_Pct',
    'Applied_Increase_Pct',
    'Calculated_List_Price',
    'Prior_Published_Price',
    'Needs_Update',
    'Sort_Order',
    'product_image_filename',
    'Image_File_ID',
    'Notes'
  ]);

  ensureSheetColumns_(ss, 'Price_Rules', [
    'Active',
    'Product_Group',
    'PLC',
    'Fitting_Type',
    'Item_Number',
    'Variant',
    'Variant_2',
    'Increase_Pct',
    'Notes'
  ]);
}

function ensureSheetColumns_(ss, sheetName, requiredHeaders) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Missing ${sheetName} sheet.`);

  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const existingHeaders = sheet.getRange(1, 1, 1, lastColumn).getValues()[0]
    .map(header => String(header || '').trim());
  const existingHeaderSet = existingHeaders.reduce((set, header) => {
    if (header) set[header] = true;
    return set;
  }, {});

  const missingHeaders = requiredHeaders.filter(header => !existingHeaderSet[header]);
  if (!missingHeaders.length) return;

  const currentLastColumn = sheet.getLastColumn();
  const startColumn = currentLastColumn + 1;

  if (currentLastColumn === 0) {
    sheet.insertColumns(1, missingHeaders.length);
  } else {
    sheet.insertColumnsAfter(currentLastColumn, missingHeaders.length);
  }

  sheet.getRange(1, startColumn, 1, missingHeaders.length)
    .setValues([missingHeaders])
    .setFontWeight('bold')
    .setBackground('#D9EAF7')
    .setWrap(true);

  sheet.autoResizeColumns(startColumn, missingHeaders.length);
}

function isTruthy_(value) {
  if (value === true) return true;
  if (typeof value === 'number') return value === 1;

  const normalized = String(value).trim().toUpperCase();
  return normalized === 'TRUE' || normalized === '1' || normalized === 'YES' || normalized === 'Y';
}

function normalizeMatchValue_(value) {
  if (value === '' || value === null || value === undefined) return '';

  const stringValue = String(value).trim();
  const numberValue = Number(stringValue);

  if (!isNaN(numberValue) && stringValue !== '') {
    return String(numberValue);
  }

  return stringValue.toUpperCase();
}

function getOptionalCellValue_(row, col, headerName) {
  return col[headerName] === undefined ? '' : row[col[headerName]];
}

function getSheetDefinitions_(ss) {
  const sheet = ss.getSheetByName('Sheet_Definitions') || ss.getSheetByName('Sheet_Definition');
  if (!sheet || sheet.getLastRow() < 2) return {};

  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const row = values[1];

  return headers.reduce((config, header, index) => {
    const key = String(header || '').trim();
    if (key) config[key] = row[index];
    return config;
  }, {});
}

function getConfigValue_(groupRow, groupCol, sheetDefinitions, headerName) {
  const groupValue = getOptionalCellValue_(groupRow, groupCol, headerName);
  if (groupValue !== '' && groupValue !== null && groupValue !== undefined) return groupValue;

  const definitionValue = sheetDefinitions[headerName];
  return definitionValue === undefined ? '' : definitionValue;
}

function getCatalogWorkbook_() {
  const properties = PropertiesService.getScriptProperties();
  const activeSpreadsheet = SpreadsheetApp.getActive();

  if (activeSpreadsheet) {
    properties.setProperty('CATALOG_WORKBOOK_ID', activeSpreadsheet.getId());
    return activeSpreadsheet;
  }

  const spreadsheetId = properties.getProperty('CATALOG_WORKBOOK_ID');
  if (!spreadsheetId) {
    throw new Error('Catalog workbook ID is not available for this scheduled execution.');
  }

  return SpreadsheetApp.openById(spreadsheetId);
}

function getCatalogRunTargetRows_(groupRows, groupCol, options) {
  const onlyProductGroup = String(options.onlyProductGroup || '').trim();
  const onlyFileName = String(options.onlyFileName || '').trim();
  const generateColumn = options.generateColumn;
  const fileColumn = options.fileColumn;
  const isTargetedRun = !!(onlyProductGroup || onlyFileName);

  return groupRows.filter(row => {
    const productGroup = String(row[groupCol.Product_Group] || '').trim();
    const fileName = String(row[groupCol[fileColumn]] || '').trim();

    if (!productGroup || !fileName) return false;
    if (!isTruthy_(row[groupCol[generateColumn]])) return false;
    if (!isTargetedRun && !isTruthy_(row[groupCol.Active])) return false;
    if (onlyProductGroup && productGroup !== onlyProductGroup) return false;
    if (onlyFileName && fileName !== onlyFileName) return false;

    return true;
  });
}

function getMsiAutomationSidebarData() {
  const ss = getCatalogWorkbook_();
  const groupSheet = ss.getSheetByName('Catalog_Groups');
  if (!groupSheet) throw new Error('Missing Catalog_Groups sheet.');

  const values = groupSheet.getDataRange().getValues();
  const col = headerMap_(values[0]);
  const groupsByName = {};

  values.slice(1).forEach(row => {
    const productGroup = String(row[col.Product_Group] || '').trim();
    if (!productGroup) return;

    if (!groupsByName[productGroup]) {
      groupsByName[productGroup] = {
        productGroup,
        isActive: false,
        catalogCount: 0,
        priceFileCount: 0,
        catalogFileNames: [],
        priceFileNames: []
      };
    }

    const entry = groupsByName[productGroup];
    entry.isActive = entry.isActive || isTruthy_(row[col.Active]);

    if (row[col.Catalog_File_Name] && isTruthy_(row[col.Generate_PDF])) {
      const catalogFileName = String(row[col.Catalog_File_Name]).trim();
      if (entry.catalogFileNames.indexOf(catalogFileName) === -1) {
        entry.catalogFileNames.push(catalogFileName);
        entry.catalogCount++;
      }
    }

    if (row[col.Price_File_Name] && isTruthy_(row[col.Generate_XLS])) {
      const priceFileName = String(row[col.Price_File_Name]).trim();
      if (entry.priceFileNames.indexOf(priceFileName) === -1) {
        entry.priceFileNames.push(priceFileName);
        entry.priceFileCount++;
      }
    }
  });

  const preferredOrder = MSI_WORKFLOW_GROUPS_.reduce((lookup, entry, index) => {
    lookup[entry.productGroup] = index;
    return lookup;
  }, {});

  const groups = Object.keys(groupsByName)
    .map(name => groupsByName[name])
    .sort((a, b) => {
      const preferredA = preferredOrder[a.productGroup];
      const preferredB = preferredOrder[b.productGroup];
      const hasPreferredA = preferredA !== undefined;
      const hasPreferredB = preferredB !== undefined;

      if (hasPreferredA && hasPreferredB) return preferredA - preferredB;
      if (hasPreferredA) return -1;
      if (hasPreferredB) return 1;

      return a.productGroup.localeCompare(b.productGroup);
    });

  return {
    groups,
    overview: getMsiAutomationOverviewData_(ss, groups),
    preflight: getMsiAutomationPreflightData_(ss),
    priceRules: getPriceRulesForSidebar_(ss),
    priceRuleOptions: getPriceRuleOptionsForSidebar_(ss),
    ruleHistory: getRecentPriceRuleHistoryForSidebar_(ss),
    productionStatus: getCatalogProductionRunState_() || null
  };
}

function getMsiAutomationPreflightData_(ss) {
  const workbook = ss || getCatalogWorkbook_();
  const requiredSheets = ['Catalog_Groups', 'Catalog_SKUs', 'Price_Rules', 'Sheet_Definitions', 'Generation_Log'];
  const issues = [];
  const seenIssueKeys = {};

  requiredSheets.forEach(sheetName => {
    if (!workbook.getSheetByName(sheetName)) {
      pushPreflightIssue_(issues, seenIssueKeys, 'error', 'Workbook', `Missing required sheet: ${sheetName}`, 'Create or restore the missing sheet before production.');
    }
  });

  const groupSheet = workbook.getSheetByName('Catalog_Groups');
  const skuSheet = workbook.getSheetByName('Catalog_SKUs');
  const ruleSheet = workbook.getSheetByName('Price_Rules');
  const sheetDefinitionSheet = workbook.getSheetByName('Sheet_Definitions') || workbook.getSheetByName('Sheet_Definition');

  if (!groupSheet || !skuSheet || !ruleSheet || !sheetDefinitionSheet) {
    return buildPreflightSummary_(issues);
  }

  const groupValues = groupSheet.getDataRange().getValues();
  const skuValues = skuSheet.getDataRange().getValues();
  const ruleValues = ruleSheet.getDataRange().getValues();
  const groupCol = headerMap_(groupValues[0] || []);
  const skuCol = headerMap_(skuValues[0] || []);
  const ruleCol = headerMap_(ruleValues[0] || []);
  const sheetDefinitions = getSheetDefinitions_(workbook);

  const activeSkuRows = skuValues.slice(1).filter(row => isTruthy_(row[skuCol.Active]));
  const activeGroups = groupValues.slice(1).filter(row => isTruthy_(row[groupCol.Active]));
  const groupLookup = {};
  const activeRuleCountByGroup = {};
  const skuCountByGroup = {};
  const imageMissingCountByGroup = {};
  const duplicateItemLookup = {};

  groupValues.slice(1).forEach(row => {
    const productGroup = String(row[groupCol.Product_Group] || '').trim();
    if (!productGroup) return;
    if (!groupLookup[productGroup]) groupLookup[productGroup] = [];
    groupLookup[productGroup].push(row);

    if (isTruthy_(row[groupCol.Active])) {
      if (isTruthy_(row[groupCol.Generate_PDF])) {
        validateCatalogGroupConfig_(issues, seenIssueKeys, row, groupCol, sheetDefinitions);
      }
      if (isTruthy_(row[groupCol.Generate_XLS])) {
        validatePriceFileGroupConfig_(issues, seenIssueKeys, row, groupCol, sheetDefinitions);
      }
    }
  });

  ruleValues.slice(1).forEach(row => {
    if (!isTruthy_(row[ruleCol.Active])) return;
    const productGroup = String(row[ruleCol.Product_Group] || '').trim();
    if (!productGroup) return;
    activeRuleCountByGroup[productGroup] = (activeRuleCountByGroup[productGroup] || 0) + 1;
  });

  activeSkuRows.forEach((row, index) => {
    const rowNumber = index + 2;
    const productGroup = String(row[skuCol.Product_Group] || '').trim();
    const itemNumber = String(row[skuCol.Item_Number] || '').trim();
    const listPrice = row[skuCol.List_Price];
    const imageFileId = colHasHeader_(skuCol, 'Image_File_ID') ? String(row[skuCol.Image_File_ID] || '').trim() : '';

    if (!productGroup) {
      pushPreflightIssue_(issues, seenIssueKeys, 'error', 'Catalog_SKUs', `Row ${rowNumber}: missing Product_Group`, 'Every active SKU needs a product group.');
      return;
    }

    skuCountByGroup[productGroup] = (skuCountByGroup[productGroup] || 0) + 1;
    if (!groupLookup[productGroup]) {
      pushPreflightIssue_(issues, seenIssueKeys, 'error', productGroup, 'Active SKU rows reference a product group not found in Catalog_Groups.', 'Add the missing Catalog_Groups row or correct the SKU product-group label.');
    }

    if (!itemNumber) {
      pushPreflightIssue_(issues, seenIssueKeys, 'error', productGroup, `Row ${rowNumber}: missing Item_Number`, 'Every active SKU needs an item number.');
    } else {
      if (!duplicateItemLookup[itemNumber]) duplicateItemLookup[itemNumber] = [];
      duplicateItemLookup[itemNumber].push({ rowNumber, productGroup });
    }

    const normalizedListPrice = normalizeCatalogPriceValue_(listPrice);
    if (normalizedListPrice === null) {
      pushPreflightIssue_(issues, seenIssueKeys, 'error', productGroup, `Row ${rowNumber}: invalid List_Price for ${itemNumber || 'unknown item'}`, 'Use a numeric price, leave it blank, or enter POA.');
    }

    if (!imageFileId) {
      imageMissingCountByGroup[productGroup] = (imageMissingCountByGroup[productGroup] || 0) + 1;
    }
  });

  Object.keys(duplicateItemLookup).forEach(itemNumber => {
    const entries = duplicateItemLookup[itemNumber];
    if (entries.length < 2) return;
    pushPreflightIssue_(
      issues,
      seenIssueKeys,
      'warning',
      entries[0].productGroup,
      `Duplicate active Item_Number found: ${itemNumber} (${entries.length} rows)`,
      'Confirm duplicates are intentional variants and not accidental duplicated SKUs.'
    );
  });

  activeGroups.forEach(row => {
    const productGroup = String(row[groupCol.Product_Group] || '').trim();
    if (!productGroup) return;

    const skuCount = skuCountByGroup[productGroup] || 0;
    if (!skuCount) {
      pushPreflightIssue_(issues, seenIssueKeys, 'error', productGroup, 'Group is active but has no active SKU rows.', 'Either deactivate the group or activate/fix its SKU rows.');
    }

    if ((activeRuleCountByGroup[productGroup] || 0) === 0) {
      pushPreflightIssue_(issues, seenIssueKeys, 'warning', productGroup, 'No active price rules found for this product group.', 'Confirm this is intentional before repricing.');
    }

    if (isTruthy_(row[groupCol.Generate_PDF]) && (imageMissingCountByGroup[productGroup] || 0) > 0) {
      pushPreflightIssue_(
        issues,
        seenIssueKeys,
        'warning',
        productGroup,
        `${imageMissingCountByGroup[productGroup]} active SKU rows are missing Image_File_ID.`,
        'Catalog generation will continue, but those sections may show MISSING IMAGE.'
      );
    }
  });

  return buildPreflightSummary_(issues);
}

function validateCatalogGroupConfig_(issues, seenIssueKeys, row, groupCol, sheetDefinitions) {
  const productGroup = String(row[groupCol.Product_Group] || '').trim() || 'Catalog_Groups';
  const catalogFileName = String(row[groupCol.Catalog_File_Name] || '').trim();
  if (!catalogFileName) {
    pushPreflightIssue_(issues, seenIssueKeys, 'error', productGroup, 'Generate_PDF is enabled but Catalog_File_Name is blank.', 'Populate Catalog_File_Name for this active catalog row.');
  }

  [
    ['Cover_Template_ID', 'Cover template ID'],
    ['Terms_Template_ID', 'Terms template ID'],
    ['PDF_Output_Folder_ID', 'PDF output folder ID'],
    ['PDF_Archive_Folder_ID', 'PDF archive folder ID']
  ].forEach(entry => {
    const value = String(getConfigValue_(row, groupCol, sheetDefinitions, entry[0]) || '').trim();
    if (!value) {
      pushPreflightIssue_(issues, seenIssueKeys, 'error', productGroup, `Missing ${entry[1]} for active catalog generation.`, `Populate ${entry[0]} in Catalog_Groups or Sheet_Definitions.`);
    }
  });
}

function validatePriceFileGroupConfig_(issues, seenIssueKeys, row, groupCol, sheetDefinitions) {
  const productGroup = String(row[groupCol.Product_Group] || '').trim() || 'Catalog_Groups';
  const priceFileName = String(row[groupCol.Price_File_Name] || '').trim();
  if (!priceFileName) {
    pushPreflightIssue_(issues, seenIssueKeys, 'error', productGroup, 'Generate_XLS is enabled but Price_File_Name is blank.', 'Populate Price_File_Name for this active price-file row.');
  }

  [
    ['Price_File_Output_Folder_ID', 'price-file output folder ID'],
    ['Price_File_Archive_Folder_ID', 'price-file archive folder ID']
  ].forEach(entry => {
    const value = String(getConfigValue_(row, groupCol, sheetDefinitions, entry[0]) || '').trim();
    if (!value) {
      pushPreflightIssue_(issues, seenIssueKeys, 'error', productGroup, `Missing ${entry[1]} for active price-file generation.`, `Populate ${entry[0]} in Catalog_Groups or Sheet_Definitions.`);
    }
  });
}

function pushPreflightIssue_(issues, seenIssueKeys, severity, scope, message, recommendation) {
  const key = [severity, scope, message].join('|');
  if (seenIssueKeys[key]) return;
  seenIssueKeys[key] = true;
  issues.push({
    severity,
    scope: String(scope || '').trim() || 'Workbook',
    message: String(message || '').trim(),
    recommendation: String(recommendation || '').trim()
  });
}

function buildPreflightSummary_(issues) {
  const normalizedIssues = (issues || []).slice().sort((a, b) =>
    getPreflightSeverityWeight_(a.severity) - getPreflightSeverityWeight_(b.severity) ||
    String(a.scope || '').localeCompare(String(b.scope || '')) ||
    String(a.message || '').localeCompare(String(b.message || ''))
  );
  const errorCount = normalizedIssues.filter(issue => issue.severity === 'error').length;
  const warningCount = normalizedIssues.filter(issue => issue.severity === 'warning').length;
  return {
    ready: errorCount === 0,
    errorCount,
    warningCount,
    issueCount: normalizedIssues.length,
    issues: normalizedIssues
  };
}

function getPreflightSeverityWeight_(severity) {
  return severity === 'error' ? 0 : 1;
}

function colHasHeader_(col, headerName) {
  return col[headerName] !== undefined;
}

function getMsiAutomationOverviewData_(ss, groups) {
  const workbook = ss || getCatalogWorkbook_();
  const skuSheet = workbook.getSheetByName('Catalog_SKUs');
  const ruleSheet = workbook.getSheetByName('Price_Rules');
  const logSheet = workbook.getSheetByName('Generation_Log');

  const skuValues = skuSheet ? skuSheet.getDataRange().getValues() : [];
  const ruleValues = ruleSheet ? ruleSheet.getDataRange().getValues() : [];
  const logValues = logSheet ? logSheet.getDataRange().getValues() : [];

  const skuCol = skuValues.length ? headerMap_(skuValues[0]) : {};
  const ruleCol = ruleValues.length ? headerMap_(ruleValues[0]) : {};

  const groupStats = {};
  const plcLookup = {};
  let totalSkuCount = 0;
  let activeSkuCount = 0;

  (groups || []).forEach(group => {
    groupStats[group.productGroup] = {
      productGroup: group.productGroup,
      activeSkuCount: 0,
      totalSkuCount: 0,
      plcLabels: [],
      plcLookup: {},
      activeRuleCount: 0,
      lastPriceFileAt: '',
      lastCatalogAt: ''
    };
  });

  skuValues.slice(1).forEach(row => {
    const productGroup = String(row[skuCol.Product_Group] || '').trim();
    if (!productGroup) return;

    if (!groupStats[productGroup]) {
      groupStats[productGroup] = {
        productGroup,
        activeSkuCount: 0,
        totalSkuCount: 0,
        plcLabels: [],
        plcLookup: {},
        activeRuleCount: 0,
        lastPriceFileAt: '',
        lastCatalogAt: ''
      };
    }

    const entry = groupStats[productGroup];
    const plc = String(row[skuCol.PLC] || '').trim();
    const isActive = isTruthy_(row[skuCol.Active]);

    totalSkuCount++;
    entry.totalSkuCount++;

    if (isActive) {
      activeSkuCount++;
      entry.activeSkuCount++;
      if (plc) plcLookup[plc] = true;
    }

    if (plc && !entry.plcLookup[plc]) {
      entry.plcLookup[plc] = true;
      entry.plcLabels.push(formatPriceFilePlc_(plc));
    }
  });

  ruleValues.slice(1).forEach(row => {
    if (!isTruthy_(row[ruleCol.Active])) return;
    const productGroup = String(row[ruleCol.Product_Group] || '').trim();
    if (!productGroup || !groupStats[productGroup]) return;
    groupStats[productGroup].activeRuleCount++;
  });

  logValues.slice(1).forEach(row => {
    const timestamp = row[0];
    const productGroup = String(row[1] || '').trim();
    const action = String(row[2] || '').trim();
    const status = String(row[3] || '').trim();
    if (!productGroup || !groupStats[productGroup] || status !== 'Success') return;

    const formatted = formatSidebarLogTimestamp_(timestamp);
    if (!formatted) return;

    if (action === 'Generate price file') {
      groupStats[productGroup].lastPriceFileAt = formatted;
    }

    if (action === 'Generate Slides catalog PDF') {
      groupStats[productGroup].lastCatalogAt = formatted;
    }
  });

  const latestPricingRun = getLatestSidebarLogTimestamp_(logValues, function(row) {
    return String(row[2] || '').trim() === 'Pricing recalculation' &&
      String(row[3] || '').trim() === 'Success';
  });

  const latestCartonRepair = getLatestSidebarLogTimestamp_(logValues, function(row) {
    return String(row[2] || '').trim() === 'Repair carton counts' &&
      String(row[3] || '').trim() === 'Success';
  });

  const orderedRows = Object.keys(groupStats)
    .map(name => groupStats[name])
    .sort((a, b) => (groups || []).findIndex(group => group.productGroup === a.productGroup) -
      (groups || []).findIndex(group => group.productGroup === b.productGroup));

  return {
    summary: {
      totalSkuCount,
      activeSkuCount,
      totalProductGroups: orderedRows.length,
      activeProductGroups: (groups || []).filter(group => group.isActive).length,
      activePlcCount: Object.keys(plcLookup).length,
      latestPricingRun,
      latestCartonRepair
    },
    groups: orderedRows.map(entry => ({
      productGroup: entry.productGroup,
      activeSkuCount: entry.activeSkuCount,
      totalSkuCount: entry.totalSkuCount,
      plcCount: Object.keys(entry.plcLookup).length,
      plcLabels: entry.plcLabels.sort((a, b) => a.localeCompare(b, undefined, { numeric: true })),
      activeRuleCount: entry.activeRuleCount,
      lastPriceFileAt: entry.lastPriceFileAt,
      lastCatalogAt: entry.lastCatalogAt
    }))
  };
}

function getLatestSidebarLogTimestamp_(logValues, predicate) {
  for (let i = logValues.length - 1; i >= 1; i--) {
    if (!predicate(logValues[i])) continue;
    return formatSidebarLogTimestamp_(logValues[i][0]);
  }
  return '';
}

function formatSidebarLogTimestamp_(value) {
  if (!value) return '';
  const date = Object.prototype.toString.call(value) === '[object Date]' ? value : new Date(value);
  if (!(date instanceof Date) || isNaN(date.getTime())) return '';
  return Utilities.formatDate(date, getCatalogTimeZone_(), 'M/d/yyyy h:mm a');
}

function getPriceRuleOptionsForSidebar_(ss) {
  const workbook = ss || getCatalogWorkbook_();
  const skuSheet = workbook.getSheetByName('Catalog_SKUs');
  if (!skuSheet) return { plcs: [], fittingTypesByGroup: {}, itemNumbers: [], skuIndex: [] };

  const values = skuSheet.getDataRange().getValues();
  if (values.length < 2) return { plcs: [], fittingTypesByGroup: {}, itemNumbers: [], skuIndex: [] };

  const col = headerMap_(values[0]);
  const plcLookup = {};
  const itemLookup = {};
  const fittingTypesByGroup = {};
  const skuIndex = [];

  values.slice(1).forEach(row => {
    if (!isTruthy_(row[col.Active])) return;

    const productGroup = String(row[col.Product_Group] || '').trim();
    const plc = String(row[col.PLC] || '').trim();
    const fittingType = String(row[col.Fitting_Type] || '').trim();
    const itemNumber = String(row[col.Item_Number] || '').trim();
    const pricingProfile = getPricingRuleFilterProfileForRow_(row, col);
    const variant = pricingProfile.variant;
    const variant2 = pricingProfile.variant2;

    if (plc) plcLookup[plc] = true;
    if (itemNumber) itemLookup[itemNumber] = true;

    if (productGroup && fittingType) {
      if (!fittingTypesByGroup[productGroup]) fittingTypesByGroup[productGroup] = {};
      fittingTypesByGroup[productGroup][fittingType] = true;
    }

    if (productGroup || plc || fittingType || itemNumber || variant || variant2) {
      skuIndex.push({
        productGroup,
        plc,
        plcLabel: formatPriceFilePlc_(plc),
        fittingType,
        itemNumber,
        variant,
        variantLabel: pricingProfile.variantLabel,
        variantAxis: pricingProfile.variantAxis,
        variant2,
        variant2Label: pricingProfile.variant2Label,
        variant2Axis: pricingProfile.variant2Axis
      });
    }
  });

  return {
    plcs: Object.keys(plcLookup).sort((a, b) => a.localeCompare(b, undefined, { numeric: true })),
    itemNumbers: Object.keys(itemLookup).sort((a, b) => a.localeCompare(b)),
    skuIndex,
    fittingTypesByGroup: Object.keys(fittingTypesByGroup).reduce((map, group) => {
      map[group] = Object.keys(fittingTypesByGroup[group]).sort((a, b) => a.localeCompare(b));
      return map;
    }, {})
  };
}

function getPriceRulesForSidebar_(ss) {
  const workbook = ss || getCatalogWorkbook_();
  const sheet = workbook.getSheetByName('Price_Rules');
  if (!sheet) return [];

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const col = headerMap_(values[0]);

  return values.slice(1)
    .map((row, index) => ({
      rowNumber: index + 2,
      active: isTruthy_(row[col.Active]),
      productGroup: String(row[col.Product_Group] || '').trim(),
      plc: String(row[col.PLC] || '').trim(),
      fittingType: String(row[col.Fitting_Type] || '').trim(),
      itemNumber: String(row[col.Item_Number] || '').trim(),
      variant: normalizeRuleSelectionValue_(col.Variant === undefined ? '' : row[col.Variant]),
      variant2: normalizeRuleSelectionValue_(getOptionalCellValue_(row, col, 'Variant_2')),
      increasePct: parsePercent_(row[col.Increase_Pct]),
      notes: String(row[col.Notes] || '').trim()
    }))
    .filter(rule =>
      rule.productGroup ||
      rule.plc ||
      rule.fittingType ||
      rule.itemNumber ||
      rule.variant ||
      rule.variant2 ||
      rule.increasePct !== null ||
      rule.notes
    )
    .sort((a, b) =>
      a.productGroup.localeCompare(b.productGroup) ||
      a.plc.localeCompare(b.plc, undefined, { numeric: true }) ||
      a.fittingType.localeCompare(b.fittingType) ||
      a.itemNumber.localeCompare(b.itemNumber) ||
      a.variant.localeCompare(b.variant) ||
      a.variant2.localeCompare(b.variant2)
    );
}

function getRecentPriceRuleHistoryForSidebar_(ss) {
  const workbook = ss || getCatalogWorkbook_();
  const logSheet = workbook.getSheetByName('Generation_Log');
  if (!logSheet) return [];

  const values = logSheet.getDataRange().getValues();
  if (values.length < 2) return [];

  return values.slice(1)
    .filter(row => /^Price Rule\b/i.test(String(row[2] || '').trim()))
    .filter(row => String(row[3] || '').trim() === 'Success')
    .slice(-12)
    .reverse()
    .map(row => ({
      timestamp: formatSidebarLogTimestamp_(row[0]),
      productGroup: String(row[1] || '').trim(),
      action: String(row[2] || '').trim(),
      message: String(row[7] || '').trim()
    }));
}

function savePriceRuleFromSidebar(ruleInput, shouldRecalculate) {
  const input = ruleInput || {};
  const ss = getCatalogWorkbook_();
  const sheet = ss.getSheetByName('Price_Rules');
  const logSheet = ss.getSheetByName('Generation_Log');
  if (!sheet) throw new Error('Missing Price_Rules sheet.');

  const values = sheet.getDataRange().getValues();
  if (!values.length) throw new Error('Price_Rules sheet is missing headers.');

  const col = headerMap_(values[0]);
  const increasePct = parsePercent_(input.increasePct);
  if (increasePct === null) throw new Error('Increase percent is required.');

  const editRowNumber = Number(input.rowNumber) || 0;
  if (editRowNumber) {
    if (editRowNumber < 2 || editRowNumber > values.length) {
      throw new Error(`Price rule row ${editRowNumber} was not found.`);
    }

    const target = buildSidebarSinglePriceRuleTarget_(input, increasePct);
    const nextRow = buildPriceRuleSheetRow_(values[0], col, target, values[editRowNumber - 1] || []);
    sheet.getRange(editRowNumber, 1, 1, nextRow.length).setValues([nextRow]);
    logPriceRuleChange_(logSheet, 'Price Rule Updated', buildPriceRuleLogGroupLabel_(target), `Row ${editRowNumber}: ${describePriceRuleTarget_(target)}`);

    if (col.Increase_Pct !== undefined && sheet.getLastRow() > 1) {
      sheet.getRange(2, col.Increase_Pct + 1, sheet.getLastRow() - 1, 1).setNumberFormat('0.00%');
    }

    if (shouldRecalculate) {
      const queueResult = queueExplicitCatalogProductionJobs_([
        { type: 'pricing_calc', fileName: 'Catalog Pricing Recalculation' }
      ], 'Pricing recalculation');
      return {
        ok: true,
        message: `Updated price rule on row ${editRowNumber} and queued pricing recalculation.`,
        status: queueResult.status
      };
    }

    return {
      ok: true,
      message: `Updated price rule on row ${editRowNumber}.`
    };
  }

  const targets = buildSidebarPriceRuleTargets_(ss, input, increasePct);
  if (!targets.length) throw new Error('No price-rule targets were generated.');

  const existingRows = values.slice(1);
  const updates = [];
  const appends = [];

  targets.forEach(target => {
    const existingIndex = existingRows.findIndex(row =>
      String(row[col.Product_Group] || '').trim() === target.productGroup &&
      String(row[col.PLC] || '').trim() === target.plc &&
      String(row[col.Fitting_Type] || '').trim() === target.fittingType &&
      String(row[col.Item_Number] || '').trim() === target.itemNumber &&
      normalizeVariantKey_(col.Variant === undefined ? '' : row[col.Variant]) === normalizeVariantKey_(target.variant) &&
      normalizeVariant2Key_(col.Variant_2 === undefined ? '' : row[col.Variant_2]) === normalizeVariant2Key_(target.variant2)
    );

    if (existingIndex === -1) {
      appends.push(target);
      return;
    }

    updates.push({
      rowNumber: existingIndex + 2,
      target
    });
  });

  updates.forEach(update => {
    const existingRow = values[update.rowNumber - 1] || [];
    const nextRow = buildPriceRuleSheetRow_(values[0], col, update.target, existingRow);
    sheet.getRange(update.rowNumber, 1, 1, nextRow.length).setValues([nextRow]);
  });

  if (appends.length) {
    const rowsToAppend = appends.map(target => buildPriceRuleSheetRow_(values[0], col, target, []));
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
  }

  logPriceRuleChange_(
    logSheet,
    'Price Rule Saved',
    buildPriceRuleLogGroupLabel_(targets[0]),
    `Saved ${targets.length} rule${targets.length === 1 ? '' : 's'}${targets.length === 1 ? ': ' + describePriceRuleTarget_(targets[0]) : '.'}`
  );

  if (col.Increase_Pct !== undefined && sheet.getLastRow() > 1) {
    sheet.getRange(2, col.Increase_Pct + 1, sheet.getLastRow() - 1, 1).setNumberFormat('0.00%');
  }

  if (shouldRecalculate) {
    const queueResult = queueExplicitCatalogProductionJobs_([
      { type: 'pricing_calc', fileName: 'Catalog Pricing Recalculation' }
    ], 'Pricing recalculation');
    return {
      ok: true,
      message:
        `Saved ${targets.length} price rule${targets.length === 1 ? '' : 's'} and queued pricing recalculation.`,
      status: queueResult.status
    };
  }

  return {
    ok: true,
    message:
      `Saved ${targets.length} price rule${targets.length === 1 ? '' : 's'}` +
      `${shouldRecalculate ? ' and recalculated pricing.' : '.'}`
  };
}

function deletePriceRuleFromSidebar(rowNumber) {
  const targetRowNumber = Number(rowNumber) || 0;
  if (!targetRowNumber) throw new Error('Rule row number is required.');

  const ss = getCatalogWorkbook_();
  const sheet = ss.getSheetByName('Price_Rules');
  const logSheet = ss.getSheetByName('Generation_Log');
  if (!sheet) throw new Error('Missing Price_Rules sheet.');
  if (targetRowNumber < 2 || targetRowNumber > sheet.getLastRow()) {
    throw new Error(`Price rule row ${targetRowNumber} was not found.`);
  }

  const values = sheet.getDataRange().getValues();
  const col = headerMap_(values[0]);
  const existingTarget = buildPriceRuleTargetFromSheetRow_(values[targetRowNumber - 1], col);

  sheet.deleteRow(targetRowNumber);
  logPriceRuleChange_(logSheet, 'Price Rule Removed', buildPriceRuleLogGroupLabel_(existingTarget), `Removed row ${targetRowNumber}: ${describePriceRuleTarget_(existingTarget)}`);
  return {
    ok: true,
    message: `Removed price rule from row ${targetRowNumber}.`
  };
}

function setPriceRuleActiveFromSidebar(rowNumber, isActive) {
  const targetRowNumber = Number(rowNumber) || 0;
  if (!targetRowNumber) throw new Error('Rule row number is required.');

  const ss = getCatalogWorkbook_();
  const sheet = ss.getSheetByName('Price_Rules');
  const logSheet = ss.getSheetByName('Generation_Log');
  if (!sheet) throw new Error('Missing Price_Rules sheet.');

  const values = sheet.getDataRange().getValues();
  if (!values.length) throw new Error('Price_Rules sheet is missing headers.');
  if (targetRowNumber < 2 || targetRowNumber > values.length) {
    throw new Error(`Price rule row ${targetRowNumber} was not found.`);
  }

  const col = headerMap_(values[0]);
  if (col.Active === undefined) throw new Error('Price_Rules sheet is missing the Active column.');

  sheet.getRange(targetRowNumber, col.Active + 1).setValue(!!isActive);
  const target = buildPriceRuleTargetFromSheetRow_(values[targetRowNumber - 1], col);
  target.active = !!isActive;
  logPriceRuleChange_(logSheet, 'Price Rule Status Updated', buildPriceRuleLogGroupLabel_(target), `${isActive ? 'Enabled' : 'Disabled'} row ${targetRowNumber}: ${describePriceRuleTarget_(target)}`);
  return {
    ok: true,
    message: `${isActive ? 'Enabled' : 'Disabled'} price rule on row ${targetRowNumber}.`
  };
}

function setPriceRulesActiveFromSidebar(rowNumbers, isActive) {
  const targetRowNumbers = (rowNumbers || [])
    .map(value => Number(value) || 0)
    .filter(value => value >= 2);
  if (!targetRowNumbers.length) throw new Error('At least one rule row number is required.');

  const ss = getCatalogWorkbook_();
  const sheet = ss.getSheetByName('Price_Rules');
  const logSheet = ss.getSheetByName('Generation_Log');
  if (!sheet) throw new Error('Missing Price_Rules sheet.');

  const values = sheet.getDataRange().getValues();
  if (!values.length) throw new Error('Price_Rules sheet is missing headers.');
  const col = headerMap_(values[0]);
  if (col.Active === undefined) throw new Error('Price_Rules sheet is missing the Active column.');

  const sortedRows = targetRowNumbers
    .filter((value, index, array) => array.indexOf(value) === index)
    .sort((a, b) => a - b);

  sortedRows.forEach(rowNumber => {
    if (rowNumber > values.length) {
      throw new Error(`Price rule row ${rowNumber} was not found.`);
    }
    sheet.getRange(rowNumber, col.Active + 1).setValue(!!isActive);
  });

  const firstTarget = buildPriceRuleTargetFromSheetRow_(values[sortedRows[0] - 1], col);
  logPriceRuleChange_(
    logSheet,
    'Price Rule Status Updated',
    buildPriceRuleLogGroupLabel_(firstTarget),
    `${isActive ? 'Enabled' : 'Disabled'} ${sortedRows.length} filtered rule${sortedRows.length === 1 ? '' : 's'} (rows ${sortedRows[0]}-${sortedRows[sortedRows.length - 1]}).`
  );

  return {
    ok: true,
    message: `${isActive ? 'Enabled' : 'Disabled'} ${sortedRows.length} filtered rule${sortedRows.length === 1 ? '' : 's'}.`
  };
}

function buildPriceRuleSheetRow_(headers, col, target, existingRow) {
  const row = headers.map((header, index) => existingRow[index] !== undefined ? existingRow[index] : '');
  row[col.Active] = target.active;
  row[col.Product_Group] = target.productGroup;
  row[col.PLC] = target.plc;
  row[col.Fitting_Type] = target.fittingType;
  row[col.Item_Number] = target.itemNumber;
  if (col.Variant !== undefined) row[col.Variant] = target.variant;
  if (col.Variant_2 !== undefined) row[col.Variant_2] = target.variant2;
  row[col.Increase_Pct] = target.increasePct;
  row[col.Notes] = target.notes;
  return row;
}

function buildPriceRuleTargetFromSheetRow_(row, col) {
  return {
    active: isTruthy_(row[col.Active]),
    productGroup: String(row[col.Product_Group] || '').trim(),
    plc: String(row[col.PLC] || '').trim(),
    fittingType: String(row[col.Fitting_Type] || '').trim(),
    itemNumber: String(row[col.Item_Number] || '').trim(),
    variant: normalizeRuleSelectionValue_(col.Variant === undefined ? '' : row[col.Variant]),
    variant2: normalizeRuleSelectionValue_(getOptionalCellValue_(row, col, 'Variant_2')),
    increasePct: parsePercent_(row[col.Increase_Pct]),
    notes: String(row[col.Notes] || '').trim()
  };
}

function describePriceRuleTarget_(target) {
  if (!target) return 'Price rule';

  const scopeDetail = target.itemNumber || target.fittingType || target.plc || target.productGroup || 'All Product Groups';
  const parts = [
    `${target.active === false ? 'Inactive' : 'Active'} ${scopeDetail}`,
    `Increase ${formatPercentForLog_(target.increasePct)}`
  ];

  if (target.variant) parts.push(`Variant ${target.variant}`);
  if (target.variant2) parts.push(`Variant 2 ${target.variant2}`);
  if (target.notes) parts.push(`Notes ${target.notes}`);
  return parts.join(' | ');
}

function buildPriceRuleLogGroupLabel_(target) {
  return String((target && target.productGroup) || 'Price Rules').trim() || 'Price Rules';
}

function formatPercentForLog_(value) {
  if (value === null || value === undefined || value === '') return '0%';
  return `${(Number(value) * 100).toFixed(2).replace(/\.00$/, '')}%`;
}

function logPriceRuleChange_(logSheet, action, productGroup, message) {
  appendCatalogLog_(logSheet, productGroup || 'Price Rules', action, 'Success', 1, '', '', message || '');
}

function buildSidebarSinglePriceRuleTarget_(input, increasePct) {
  const scope = String(input.scope || '').trim();
  const active = input.active === false ? false : true;
  const variants = normalizeRuleSelectionValues_(input.variants || input.variant);
  const variant2s = normalizeRuleSelectionValues_(input.variant2s || input.variant2);

  const target = {
    active,
    productGroup: String(input.productGroup || '').trim(),
    plc: '',
    fittingType: '',
    itemNumber: '',
    variant: variants[0] || '',
    variant2: variant2s[0] || '',
    increasePct,
    notes: String(input.notes || '').trim()
  };

  if (scope === 'plc') {
    target.plc = String(input.plc || '').trim();
    if (!target.plc) throw new Error('PLC is required for a PLC-level rule.');
  } else if (scope === 'fittingType') {
    target.fittingType = String(input.fittingType || '').trim();
    if (!target.productGroup) throw new Error('Product group is required for a fitting-type rule.');
    if (!target.fittingType) throw new Error('Fitting type is required for a fitting-type rule.');
  } else if (scope === 'itemNumber') {
    target.itemNumber = String(input.itemNumber || '').trim();
    if (!target.itemNumber) throw new Error('Item number is required for an item-level rule.');
  } else {
    if (!target.productGroup) throw new Error('Product group is required for a product-group rule.');
  }

  return target;
}

function buildSidebarPriceRuleTargets_(ss, input, increasePct) {
  const scope = String(input.scope || '').trim();
  const active = input.active === false ? false : true;
  const notes = String(input.notes || '').trim();
  const plc = String(input.plc || '').trim();
  const fittingType = String(input.fittingType || '').trim();
  const itemNumber = String(input.itemNumber || '').trim();
  const productGroup = String(input.productGroup || '').trim();
  const variants = normalizeRuleSelectionValues_(input.variants || input.variant);
  const variant2s = normalizeRuleSelectionValues_(input.variant2s || input.variant2);

  if (!scope) throw new Error('Rule scope is required.');

  if (scope === 'allGroups') {
    return getUniqueCatalogProductGroups_(ss).map(name => ({
      active,
      productGroup: name,
      plc: '',
      fittingType: '',
      itemNumber: '',
      variant: '',
      variant2: '',
      increasePct,
      notes
    }));
  }

  if (scope === 'group') {
    if (!productGroup) throw new Error('Product group is required for a product-group rule.');
    return buildRuleTargetsFromMatches_(
      ss,
      {
        active,
        productGroup,
        plc: '',
        fittingType: '',
        itemNumber: '',
        increasePct,
        notes
      },
      {
        productGroup,
        variants,
        variant2s
      }
    ) || [{
      active,
      productGroup,
      plc: '',
      fittingType: '',
      itemNumber: '',
      variant: '',
      variant2: '',
      increasePct,
      notes
    }];
  }

  if (scope === 'plc') {
    if (!plc) throw new Error('PLC is required for a PLC-level rule.');
    const matchingProductGroups = getCatalogProductGroupsForFieldValue_(ss, 'PLC', plc);
    if (!matchingProductGroups.length) {
      throw new Error(`No Catalog_SKUs rows found for PLC ${plc}.`);
    }

    return matchingProductGroups.reduce((targets, groupName) => {
      return targets.concat(buildRuleTargetsFromMatches_(
        ss,
        {
          active,
          productGroup: groupName,
          plc,
          fittingType: '',
          itemNumber: '',
          increasePct,
          notes
        },
        {
          productGroup: groupName,
          plc,
          variants,
          variant2s
        }
      ));
    }, []);
  }

  if (scope === 'fittingType') {
    if (!productGroup) throw new Error('Product group is required for a fitting-type rule.');
    if (!fittingType) throw new Error('Fitting type is required for a fitting-level rule.');
    return buildRuleTargetsFromMatches_(
      ss,
      {
        active,
        productGroup,
        plc: '',
        fittingType,
        itemNumber: '',
        increasePct,
        notes
      },
      {
        productGroup,
        fittingType,
        variants,
        variant2s
      }
    ) || [{
      active,
      productGroup,
      plc: '',
      fittingType,
      itemNumber: '',
      variant: '',
      variant2: '',
      increasePct,
      notes
    }];
  }

  if (scope === 'itemNumber') {
    if (!itemNumber) throw new Error('Item number is required for an item-level rule.');
    const matchingItemGroups = getCatalogProductGroupsForFieldValue_(ss, 'Item_Number', itemNumber);
    if (!matchingItemGroups.length) {
      throw new Error(`No Catalog_SKUs rows found for item number ${itemNumber}.`);
    }

    return matchingItemGroups.reduce((targets, groupName) => {
      return targets.concat(buildRuleTargetsFromMatches_(
        ss,
        {
          active,
          productGroup: groupName,
          plc: '',
          fittingType: '',
          itemNumber,
          increasePct,
          notes
        },
        {
          productGroup: groupName,
          itemNumber,
          variants,
          variant2s
        }
      ));
    }, []);
  }

  throw new Error(`Unsupported rule scope: ${scope}`);
}

function buildRuleTargetsFromMatches_(ss, baseTarget, filters) {
  const variants = normalizeRuleSelectionValues_(filters.variants);
  const variant2s = normalizeRuleSelectionValues_(filters.variant2s);

  if (!variants.length && !variant2s.length) {
    return [Object.assign({}, baseTarget, { variant: '', variant2: '' })];
  }

  const matchingRows = getCatalogSkuRowsForRuleFiltering_(ss).filter(row =>
    doesSidebarRuleSkuRowMatch_(row, filters)
  );

  if (!matchingRows.length) {
    throw new Error('No Catalog_SKUs rows matched the selected rule filters.');
  }

  const targetLookup = {};
  matchingRows.forEach(row => {
    const target = Object.assign({}, baseTarget, {
      variant: variants.length ? row.variant : '',
      variant2: variant2s.length ? row.variant2 : ''
    });
    const key = [
      target.productGroup,
      target.plc,
      target.fittingType,
      target.itemNumber,
      target.variant,
      target.variant2
    ].join('|');
    targetLookup[key] = target;
  });

  return Object.keys(targetLookup).sort().map(key => targetLookup[key]);
}

function getCatalogSkuRowsForRuleFiltering_(ss) {
  const workbook = ss || getCatalogWorkbook_();
  const skuSheet = workbook.getSheetByName('Catalog_SKUs');
  if (!skuSheet) throw new Error('Missing Catalog_SKUs sheet.');

  const values = skuSheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const col = headerMap_(values[0]);
  return values.slice(1)
    .filter(row => isTruthy_(row[col.Active]))
    .map(row => {
      const pricingProfile = getPricingRuleFilterProfileForRow_(row, col);
      return {
        productGroup: String(row[col.Product_Group] || '').trim(),
        plc: String(row[col.PLC] || '').trim(),
        fittingType: String(row[col.Fitting_Type] || '').trim(),
        itemNumber: String(row[col.Item_Number] || '').trim(),
        variant: pricingProfile.variant,
        variant2: pricingProfile.variant2,
        variantAxis: pricingProfile.variantAxis,
        variant2Axis: pricingProfile.variant2Axis
      };
    });
}

function doesSidebarRuleSkuRowMatch_(row, filters) {
  if (filters.productGroup && row.productGroup !== filters.productGroup) return false;
  if (filters.plc && row.plc !== filters.plc) return false;
  if (filters.fittingType && row.fittingType !== filters.fittingType) return false;
  if (filters.itemNumber && row.itemNumber !== filters.itemNumber) return false;

  const variants = normalizeRuleSelectionValues_(filters.variants);
  if (variants.length && variants.indexOf(row.variant) === -1) return false;

  const variant2s = normalizeRuleSelectionValues_(filters.variant2s);
  if (variant2s.length && variant2s.indexOf(row.variant2) === -1) return false;

  return true;
}

function normalizeRuleSelectionValues_(values) {
  const arrayValue = Array.isArray(values) ? values : (values === undefined || values === null || values === '' ? [] : [values]);
  return arrayValue
    .map(value => normalizeRuleSelectionValue_(value))
    .filter(Boolean)
    .filter((value, index, array) => array.indexOf(value) === index)
    .sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
}

function normalizeRuleSelectionValue_(value) {
  return normalizeCatalogWholeNumberLabel_(value);
}

function getPricingRuleFilterProfileForRow_(row, col) {
  const productGroup = String(row[col.Product_Group] || '').trim();
  const normalizedGroup = productGroup.toLowerCase();
  const rawVariant = normalizeRuleSelectionValue_(col.Variant === undefined ? '' : row[col.Variant]);
  const rawVariant2 = normalizeRuleSelectionValue_(getOptionalCellValue_(row, col, 'Variant_2'));

  let variant = rawVariant;
  let variant2 = rawVariant2;
  let variantAxis = '';
  let variant2Axis = '';

  if (isCarbonSteelNippleCatalogProductGroup_(productGroup)) {
    variantAxis = 'Finish';
    variant2Axis = 'Schedule / Type';
    variant2 = getCarbonSteelNippleRuleClassLabel_(row, col);
  } else if (normalizedGroup.indexOf('stainless steel nipple') !== -1) {
    variantAxis = 'Schedule / Type';
    variant2Axis = 'Stainless Grade';
  } else if (normalizedGroup.indexOf('malleable iron') !== -1) {
    variantAxis = 'Finish';
    variant2Axis = 'Class';
  } else if (normalizedGroup.indexOf('forged stainless steel fitting') !== -1 ||
      normalizedGroup.indexOf('stainless steel butt-weld') !== -1 ||
      normalizedGroup.indexOf('stainless steel butt weld') !== -1 ||
      normalizedGroup.indexOf('stainless steel cast') !== -1 ||
      normalizedGroup.indexOf('stainless steel flange') !== -1) {
    variantAxis = 'Stainless Grade';
    variant2Axis = rawVariant2 && /thread|socket|sweat|flare/i.test(rawVariant2) ? 'Connection' : 'Schedule / Type';
  } else if (normalizedGroup.indexOf('valve') !== -1) {
    variantAxis = 'Material';
    variant2Axis = rawVariant2 && /thread|socket|sweat|flare|cxc|fpt|mpt|male|female/i.test(rawVariant2) ? 'Connection' : '';
  } else {
    variantAxis = guessPricingRuleAxisLabel_([rawVariant], false);
    variant2Axis = guessPricingRuleAxisLabel_([rawVariant2], true);
  }

  return {
    variant,
    variantLabel: getPricingRuleDisplayValue_(variant),
    variantAxis,
    variant2,
    variant2Label: getPricingRuleDisplayValue_(variant2),
    variant2Axis
  };
}

function getPricingRuleVariantValueForRow_(row, col) {
  return getPricingRuleFilterProfileForRow_(row, col).variant;
}

function getPricingRuleVariant2ValueForRow_(row, col) {
  return getPricingRuleFilterProfileForRow_(row, col).variant2;
}

function isCarbonSteelNippleCatalogProductGroup_(productGroup) {
  const normalized = String(productGroup || '').toLowerCase();
  return normalized.indexOf('carbon steel nipple') !== -1;
}

function getCarbonSteelNippleRuleClassLabel_(row, col) {
  const plc = normalizeMatchValue_(row[col.PLC]);
  if (plc === '724') return 'A106GrB';
  if (plc === '727') return 'Sch40 Welded';
  if (plc === '729') return 'Sch80 Seamless';
  return normalizeRuleSelectionValue_(getOptionalCellValue_(row, col, 'Variant_2'));
}

function getPricingRuleDisplayValue_(value) {
  const normalized = String(value || '').trim().toLowerCase();
  if (normalized === 'plain' || normalized === 'steel') return 'Plain (Steel)';
  if (normalized === 'plated' || normalized === 'zinc plated') return 'Zinc Plated';
  if (normalized === 'black') return 'Black';
  if (normalized === 'galvanized') return 'Galvanized';
  if (normalized === '304') return '304';
  if (normalized === '316') return '316';
  return String(value || '').trim();
}

function guessPricingRuleAxisLabel_(values, isSecondary) {
  const normalized = (values || []).map(value => String(value || '').trim().toLowerCase()).filter(Boolean);
  if (!normalized.length) return isSecondary ? 'Secondary Variant' : 'Variant';
  if (normalized.some(value => value === '304' || value === '316')) return 'Stainless Grade';
  if (normalized.some(value => value === 'black' || value === 'galvanized' || value === 'plain' || value === 'steel' || value === 'zinc plated' || value === 'plated')) return 'Finish';
  if (normalized.some(value => value.indexOf('sch') !== -1 || value.indexOf('s40') !== -1 || value.indexOf('s80') !== -1 || value.indexOf('a106') !== -1 || value.indexOf('seamless') !== -1 || value.indexOf('welded') !== -1)) return 'Schedule / Type';
  if (normalized.some(value => value.indexOf('150') !== -1 || value.indexOf('300') !== -1)) return 'Class';
  if (normalized.some(value => value.indexOf('thread') !== -1 || value.indexOf('socket') !== -1 || value.indexOf('sweat') !== -1 || value.indexOf('flare') !== -1 || value.indexOf('cx') !== -1 || value.indexOf('fpt') !== -1 || value.indexOf('mpt') !== -1)) return 'Connection';
  return isSecondary ? 'Secondary Variant' : 'Variant';
}

function getUniqueCatalogProductGroups_(ss) {
  const workbook = ss || getCatalogWorkbook_();
  const groupSheet = workbook.getSheetByName('Catalog_Groups');
  if (!groupSheet) throw new Error('Missing Catalog_Groups sheet.');

  const values = groupSheet.getDataRange().getValues();
  const col = headerMap_(values[0]);
  const seen = {};

  return values.slice(1).reduce((groups, row) => {
    const productGroup = String(row[col.Product_Group] || '').trim();
    if (!productGroup || seen[productGroup]) return groups;
    seen[productGroup] = true;
    groups.push(productGroup);
    return groups;
  }, []).sort((a, b) => a.localeCompare(b));
}

function getCatalogProductGroupsForFieldValue_(ss, fieldName, fieldValue) {
  const workbook = ss || getCatalogWorkbook_();
  const skuSheet = workbook.getSheetByName('Catalog_SKUs');
  if (!skuSheet) throw new Error('Missing Catalog_SKUs sheet.');

  const values = skuSheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const col = headerMap_(values[0]);
  const normalizedTarget = String(fieldValue || '').trim();
  const lookup = {};

  values.slice(1).forEach(row => {
    const candidate = String(row[col[fieldName]] || '').trim();
    const productGroup = String(row[col.Product_Group] || '').trim();
    if (!candidate || !productGroup) return;
    if (candidate !== normalizedTarget) return;
    lookup[productGroup] = true;
  });

  return Object.keys(lookup).sort((a, b) => a.localeCompare(b));
}

function runMsiAutomationAction(action, productGroup) {
  const normalizedAction = String(action || '').trim();
  const normalizedGroup = String(productGroup || '').trim();

  switch (normalizedAction) {
    case 'calculatePricing':
      return calculateCatalogPricing();

    case 'generateCatalog':
      if (!normalizedGroup) throw new Error('Missing product group for catalog generation.');
      return queueCatalogProductionSelection_({
        includeCatalogPdfs: true,
        includePriceFiles: false,
        onlyProductGroup: normalizedGroup,
        requireActive: false,
        mode: `${normalizedGroup} catalog creation`
      });

    case 'generatePriceList':
      if (!normalizedGroup) throw new Error('Missing product group for price-list generation.');
      return queueCatalogProductionSelection_({
        includeCatalogPdfs: false,
        includePriceFiles: true,
        onlyProductGroup: normalizedGroup,
        requireActive: false,
        mode: `${normalizedGroup} price-list creation`
      });

    case 'previewCatalog':
      if (!normalizedGroup) throw new Error('Missing product group for catalog preview.');
      return {
        ok: true,
        message: `Catalog preview generated for ${normalizedGroup}. Check the Apps Script execution log for full planning detail.`,
        preview: buildCatalogSlidePlanLogSummary_(previewCatalogSlidePlan(normalizedGroup))
      };

    case 'generateAllCatalogs':
      return startFullCatalogSlidesGeneration();

    case 'generateAllPriceLists':
      return startFullPriceFileGeneration();

    case 'generateAll':
      return startFullCatalogProductionRun();

    case 'runPreflight':
      return {
        ok: true,
        message: 'Preflight refreshed.',
        preflight: getMsiAutomationPreflightData_(getCatalogWorkbook_())
      };

    case 'productionStatus':
      return {
        ok: true,
        message: 'Production status loaded.',
        status: getCatalogProductionRunState_() || null
      };

    case 'rerunFailedJobs':
      return rerunFailedCatalogProductionJobs_();

    case 'cancelProduction':
      cancelCatalogProductionRun();
      return {
        ok: true,
        message: 'Active production queue cancelled.',
        status: getCatalogProductionRunState_() || null
      };

    default:
      throw new Error(`Unsupported workflow action: ${normalizedAction}`);
  }
}

function queueSelectedCatalogGroupsFromSidebar(productGroups) {
  const selectedGroups = (productGroups || [])
    .map(value => String(value || '').trim())
    .filter(Boolean)
    .filter((value, index, array) => array.indexOf(value) === index);

  if (!selectedGroups.length) {
    throw new Error('Select at least one catalog product group.');
  }

  return queueCatalogProductionSelection_({
    includeCatalogPdfs: true,
    includePriceFiles: false,
    onlyProductGroups: selectedGroups,
    requireActive: false,
    mode: `${selectedGroups.length} selected catalog group${selectedGroups.length === 1 ? '' : 's'} creation`
  });
}

function queueSelectedPriceListGroupsFromSidebar(productGroups) {
  const selectedGroups = (productGroups || [])
    .map(value => String(value || '').trim())
    .filter(Boolean)
    .filter((value, index, array) => array.indexOf(value) === index);

  if (!selectedGroups.length) {
    throw new Error('Select at least one price-list product group.');
  }

  return queueCatalogProductionSelection_({
    includeCatalogPdfs: false,
    includePriceFiles: true,
    onlyProductGroups: selectedGroups,
    requireActive: false,
    mode: `${selectedGroups.length} selected price-list group${selectedGroups.length === 1 ? '' : 's'} creation`
  });
}

function matchesOptionalGroupPlc_(skuRow, skuCol, groupPlc) {
  const normalizedGroupPlc = normalizeMatchValue_(groupPlc);
  if (!normalizedGroupPlc) return true;

  const allowedPlcs = String(groupPlc)
    .split(',')
    .map(value => normalizeMatchValue_(value))
    .filter(Boolean);

  return allowedPlcs.includes(normalizeMatchValue_(skuRow[skuCol.PLC]));
}

function normalizeCatalogCartonsForGeneration_() {
  repairCatalogInnerMasterCartonCountsNow_();
}

function generateCatalogPriceFiles(onlyProductGroup, onlyPriceFileName, options) {
  if (!onlyProductGroup && !onlyPriceFileName) {
    return startFullPriceFileGeneration();
  }

  if (!(options && options.skipCartonNormalization)) {
    normalizeCatalogCartonsForGeneration_();
  }

  const ss = getCatalogWorkbook_();
  const skuSheet = ss.getSheetByName('Catalog_SKUs');
  const groupSheet = ss.getSheetByName('Catalog_Groups');
  const logSheet = ss.getSheetByName('Generation_Log');

  if (!skuSheet) throw new Error('Missing Catalog_SKUs sheet.');
  if (!groupSheet) throw new Error('Missing Catalog_Groups sheet.');

  const skuValues = skuSheet.getDataRange().getValues();
  const skuHeaders = skuValues[0];
  const skuCol = headerMap_(skuHeaders);

  const groupValues = groupSheet.getDataRange().getValues();
  const groupHeaders = groupValues[0];
  const groupCol = headerMap_(groupHeaders);
  const sheetDefinitions = getSheetDefinitions_(ss);

  const activeGroups = getCatalogRunTargetRows_(groupValues.slice(1), groupCol, {
    onlyProductGroup,
    onlyFileName: onlyPriceFileName,
    generateColumn: 'Generate_XLS',
    fileColumn: 'Price_File_Name'
  });

  activeGroups.forEach(group => {
    const productGroup = String(group[groupCol.Product_Group]).trim();
    const priceFileName = String(group[groupCol.Price_File_Name]).trim();
    const jobStartedAt = getCatalogTimestamp_();
    const title = String(getOptionalCellValue_(group, groupCol, 'Price_File_Title') || group[groupCol.Catalog_Title] || productGroup).trim();
    const subtitle = String(getOptionalCellValue_(group, groupCol, 'Price_File_Subtitle') || group[groupCol.Catalog_Subtitle] || '').trim();
    const versionCode = String(group[groupCol.Version_Code] || '').trim();
    const groupPlc = getOptionalCellValue_(group, groupCol, 'PLC');
    const outputFolderId = String(getConfigValue_(group, groupCol, sheetDefinitions, 'Price_File_Output_Folder_ID') || '').trim();
    const archiveFolderId = String(getConfigValue_(group, groupCol, sheetDefinitions, 'Price_File_Archive_Folder_ID') || '').trim();
    const outputFileName = priceFileName.replace(/\.xlsx$/i, '');
    const outputFolder = outputFolderId ? DriveApp.getFolderById(outputFolderId) : null;
    const archiveFolder = archiveFolderId ? DriveApp.getFolderById(archiveFolderId) : null;

    const rows = skuValues.slice(1).filter(row =>
      isTruthy_(row[skuCol.Active]) &&
      String(row[skuCol.Product_Group]).trim() === productGroup &&
      matchesOptionalGroupPlc_(row, skuCol, groupPlc) &&
      row[skuCol.Item_Number]
    ).sort((a, b) => comparePriceFileRows_(a, b, skuCol));

    if (!rows.length) {
      appendCatalogLog_(logSheet, productGroup, 'Generate price file', 'Skipped', 0, '', '', 'No active SKU rows found.');
      return;
    }

    const outputSs = SpreadsheetApp.create(outputFileName);
    const outputSheet = outputSs.getActiveSheet();
    outputSheet.setName('Price List');

    try {
      buildPriceFileSheet_(outputSheet, rows, skuCol, {
        title,
        subtitle,
        versionCode,
        logoFileId: String(getConfigValue_(group, groupCol, sheetDefinitions, 'Logo_File_ID') || '').trim()
      });
      SpreadsheetApp.flush();

      archiveExistingDriveFiles_(outputFolder, archiveFolder, [outputFileName, priceFileName]);

      if (outputFolder) {
        const file = DriveApp.getFileById(outputSs.getId());
        outputFolder.addFile(file);
        DriveApp.getRootFolder().removeFile(file);
      }

      appendCatalogLog_(
        logSheet,
        productGroup,
        'Generate price file',
        'Success',
        rows.length,
        '',
        outputSs.getId(),
        `Created ${priceFileName} with ${rows.length} rows.`,
        buildProductionLogDetails_('price_file', productGroup, priceFileName, jobStartedAt, getCatalogTimestamp_())
      );

    } catch (err) {
      DriveApp.getFileById(outputSs.getId()).setTrashed(true);
      appendCatalogLog_(
        logSheet,
        productGroup,
        'Generate price file',
        'Error',
        rows.length,
        '',
        '',
        err.message,
        buildProductionLogDetails_('price_file', productGroup, priceFileName, jobStartedAt, getCatalogTimestamp_())
      );
      throw err;
    }
  });
}

function generateCatalogPDFs() {
  throw new Error('Legacy sheet-based catalog PDF generation has been retired. Use generateCatalogSlidesPDFs() or the MSI Automation sidebar.');
}

function previewCatalogSlidePlan(onlyProductGroup) {
  const ss = SpreadsheetApp.getActive();
  const skuSheet = ss.getSheetByName('Catalog_SKUs');
  const groupSheet = ss.getSheetByName('Catalog_Groups');
  const logSheet = ss.getSheetByName('Generation_Log');

  if (!skuSheet) throw new Error('Missing Catalog_SKUs sheet.');
  if (!groupSheet) throw new Error('Missing Catalog_Groups sheet.');

  const skuValues = skuSheet.getDataRange().getValues();
  const skuHeaders = skuValues[0];
  const skuCol = headerMap_(skuHeaders);

  const groupValues = groupSheet.getDataRange().getValues();
  const groupHeaders = groupValues[0];
  const groupCol = headerMap_(groupHeaders);

  const activeGroups = getCatalogRunTargetRows_(groupValues.slice(1), groupCol, {
    onlyProductGroup,
    onlyFileName: '',
    generateColumn: 'Generate_PDF',
    fileColumn: 'Catalog_File_Name'
  });

  const plan = activeGroups.map(group => {
    const productGroup = String(group[groupCol.Product_Group]).trim();
    const catalogFileName = String(group[groupCol.Catalog_File_Name]).trim();
    const groupPlc = getOptionalCellValue_(group, groupCol, 'PLC');

    const rows = skuValues.slice(1)
      .filter(row =>
        isTruthy_(row[skuCol.Active]) &&
        String(row[skuCol.Product_Group]).trim() === productGroup &&
        matchesOptionalGroupPlc_(row, skuCol, groupPlc) &&
        row[skuCol.Item_Number]
      )
      .sort((a, b) => compareCatalogSkuRows_(a, b, skuCol));

    const sections = buildCatalogSectionModels_(rows, skuCol);
    const pages = buildCatalogSlidePages_(sections);

    const summary = {
      productGroup,
      plc: String(groupPlc || '').trim(),
      catalogFileName,
      skuRows: rows.length,
      sections: sections.length,
      pages: pages.length,
      pageMix: summarizeCatalogPageMix_(pages)
    };

    appendCatalogLog_(
      logSheet,
      productGroup,
      'Preview catalog slide plan',
      rows.length ? 'Success' : 'Skipped',
      rows.length,
      '',
      '',
      rows.length
        ? `${catalogFileName}: ${sections.length} sections across ${pages.length} planned product pages. ${formatCatalogPageMix_(summary.pageMix)}`
        : 'No active SKU rows found.'
    );

    return {
      summary,
      sections,
      pages
    };
  });

  Logger.log(JSON.stringify(buildCatalogSlidePlanLogSummary_(plan), null, 2));
  return plan;
}

function runCatalogSlidesForProductGroup_(productGroup) {
  return queueCatalogProductionSelection_({
    includeCatalogPdfs: true,
    includePriceFiles: false,
    onlyProductGroup: String(productGroup || '').trim(),
    requireActive: false,
    mode: `${productGroup} catalog creation`
  });
}

function runPriceFilesForProductGroup_(productGroup) {
  return generateCatalogPriceFiles(productGroup);
}

function generateMerchantSteelCouplingsCatalogSlidesPDF() {
  return runCatalogSlidesForProductGroup_('Merchant Steel Couplings');
}

function generateMerchantSteelCouplingsPriceFiles() {
  return runPriceFilesForProductGroup_('Merchant Steel Couplings');
}

function generateMerchantSteelCatalogSlidesPDF() {
  return generateMerchantSteelCouplingsCatalogSlidesPDF();
}

function generateMerchantSteelFittingsCatalogSlidesPDF() {
  return runCatalogSlidesForProductGroup_('Merchant Steel Fittings');
}

function generateMerchantSteelFittingsPriceFiles() {
  return runPriceFilesForProductGroup_('Merchant Steel Fittings');
}

function generateForgedStainlessSteelFittingsCatalogSlidesPDF() {
  return runCatalogSlidesForProductGroup_('Forged Stainless Steel Fittings');
}

function generateForgedStainlessSteelFittingsPriceFiles() {
  return runPriceFilesForProductGroup_('Forged Stainless Steel Fittings');
}

function generateForgedStainlessCatalogSlidesPDF() {
  return generateForgedStainlessSteelFittingsCatalogSlidesPDF();
}

function generateMalleableIronCatalogSlidesPDF() {
  return runCatalogSlidesForProductGroup_('Malleable Iron Fittings');
}

function generateMalleableIronPriceFiles() {
  return runPriceFilesForProductGroup_('Malleable Iron Fittings');
}

function generateCarbonSteelNipplesCatalogSlidesPDF() {
  return runCatalogSlidesForProductGroup_('Carbon Steel Nipples');
}

function generateCarbonSteelNipplesPriceFiles() {
  return runPriceFilesForProductGroup_('Carbon Steel Nipples');
}

function generateStainlessSteelNipplesCatalogSlidesPDF() {
  return runCatalogSlidesForProductGroup_('Stainless Steel Nipples');
}

function previewStainlessSteelNipplesCatalogSlidePlan() {
  return previewCatalogSlidePlan('Stainless Steel Nipples');
}

function generateStainlessSteelNipplesPriceFile() {
  return runPriceFilesForProductGroup_('Stainless Steel Nipples');
}

function generateStainlessSteelNipplesPriceFiles() {
  return runPriceFilesForProductGroup_('Stainless Steel Nipples');
}

function generateValvesCatalogSlidesPDF() {
  return runCatalogSlidesForProductGroup_('Valves');
}

function generateValvesPriceFiles() {
  return runPriceFilesForProductGroup_('Valves');
}

function previewValvesCatalogSlidePlan() {
  return previewCatalogSlidePlan('Valves');
}

function generateAluminiumFittingsCatalogSlidesPDF() {
  return runCatalogSlidesForProductGroup_('Aluminium Fittings');
}

function generateAluminiumFittingsPriceFiles() {
  return runPriceFilesForProductGroup_('Aluminium Fittings');
}

function generateAluminiumNipplesCatalogSlidesPDF() {
  return runCatalogSlidesForProductGroup_('Aluminium Nipples');
}

function generateAluminiumNipplesPriceFiles() {
  return runPriceFilesForProductGroup_('Aluminium Nipples');
}

function generateBrassNipplesCatalogSlidesPDF() {
  return runCatalogSlidesForProductGroup_('Brass Nipples');
}

function generateBrassNipplesPriceFiles() {
  return runPriceFilesForProductGroup_('Brass Nipples');
}

function generateBronzeFittingsCatalogSlidesPDF() {
  return runCatalogSlidesForProductGroup_('Bronze Fittings');
}

function generateBronzeFittingsPriceFiles() {
  return runPriceFilesForProductGroup_('Bronze Fittings');
}

function generateLeadFreeBronzeFittingsCatalogSlidesPDF() {
  return runCatalogSlidesForProductGroup_('Lead Free Bronze Fittings');
}

function generateLeadFreeBronzeFittingsPriceFiles() {
  return runPriceFilesForProductGroup_('Lead Free Bronze Fittings');
}

function generateCarbonSteelFlangesCatalogSlidesPDF() {
  return runCatalogSlidesForProductGroup_('Carbon Steel Flanges');
}

function generateCarbonSteelFlangesPriceFiles() {
  return runPriceFilesForProductGroup_('Carbon Steel Flanges');
}

function generateForgedSteelFittingsCatalogSlidesPDF() {
  return runCatalogSlidesForProductGroup_('Forged Steel Fittings');
}

function generateForgedSteelFittingsPriceFiles() {
  return runPriceFilesForProductGroup_('Forged Steel Fittings');
}

function generateForgedStainlessSteelFlangesCatalogSlidesPDF() {
  return runCatalogSlidesForProductGroup_('Forged Stainless Steel Flanges');
}

function generateForgedStainlessSteelFlangesPriceFiles() {
  return runPriceFilesForProductGroup_('Forged Stainless Steel Flanges');
}

function generateStainlessSteelCastFittingsCatalogSlidesPDF() {
  return runCatalogSlidesForProductGroup_('Stainless Steel Cast Fittings');
}

function generateStainlessSteelCastFittingsPriceFiles() {
  return runPriceFilesForProductGroup_('Stainless Steel Cast Fittings');
}

function generateStainlessSteelButtWeldFittingsCatalogSlidesPDF() {
  return runCatalogSlidesForProductGroup_('Stainless Steel Butt-Weld Fittings');
}

function generateStainlessSteelButtWeldFittingsPriceFiles() {
  return runPriceFilesForProductGroup_('Stainless Steel Butt-Weld Fittings');
}

function generateCarbonSteelButtWeldFittingsCatalogSlidesPDF() {
  return runCatalogSlidesForProductGroup_('Carbon Steel Butt-Weld Fittings');
}

function generateCarbonSteelButtWeldFittingsPriceFiles() {
  return runPriceFilesForProductGroup_('Carbon Steel Butt-Weld Fittings');
}

function generateCatalogSlidesPDFs(onlyProductGroup, onlyCatalogFileName, retryAttempt, options) {
  if (!onlyProductGroup && !onlyCatalogFileName) {
    return startFullCatalogSlidesGeneration();
  }

  if (!(options && options.skipCartonNormalization)) {
    normalizeCatalogCartonsForGeneration_();
  }

  const ss = getCatalogWorkbook_();
  const skuSheet = ss.getSheetByName('Catalog_SKUs');
  const groupSheet = ss.getSheetByName('Catalog_Groups');
  const logSheet = ss.getSheetByName('Generation_Log');

  if (!skuSheet) throw new Error('Missing Catalog_SKUs sheet.');
  if (!groupSheet) throw new Error('Missing Catalog_Groups sheet.');

  const skuValues = skuSheet.getDataRange().getValues();
  const skuHeaders = skuValues[0];
  const skuCol = headerMap_(skuHeaders);

  const groupValues = groupSheet.getDataRange().getValues();
  const groupHeaders = groupValues[0];
  const groupCol = headerMap_(groupHeaders);
  const sheetDefinitions = getSheetDefinitions_(ss);

  const activeGroups = getCatalogRunTargetRows_(groupValues.slice(1), groupCol, {
    onlyProductGroup,
    onlyFileName: onlyCatalogFileName,
    generateColumn: 'Generate_PDF',
    fileColumn: 'Catalog_File_Name'
  });

  for (let groupIndex = 0; groupIndex < activeGroups.length; groupIndex++) {
    const group = activeGroups[groupIndex];
    const productGroup = String(group[groupCol.Product_Group]).trim();
    const catalogFileName = String(group[groupCol.Catalog_File_Name]).trim();
    const jobStartedAt = getCatalogTimestamp_();
    const groupPlc = getOptionalCellValue_(group, groupCol, 'PLC');

    const rows = skuValues.slice(1)
      .filter(row =>
        isTruthy_(row[skuCol.Active]) &&
        String(row[skuCol.Product_Group]).trim() === productGroup &&
        matchesOptionalGroupPlc_(row, skuCol, groupPlc) &&
        row[skuCol.Item_Number]
      )
      .sort((a, b) => compareCatalogSkuRows_(a, b, skuCol));

    if (!rows.length) {
      appendCatalogLog_(logSheet, productGroup, 'Generate Slides catalog PDF', 'Skipped', 0, '', '', 'No active SKU rows found.');
      continue;
    }

    const sections = buildCatalogSectionModels_(rows, skuCol);
    const pages = buildCatalogSlidePages_(sections);
    const outputFolderId = String(getConfigValue_(group, groupCol, sheetDefinitions, 'PDF_Output_Folder_ID') || '').trim();
    const archiveFolderId = String(getConfigValue_(group, groupCol, sheetDefinitions, 'PDF_Archive_Folder_ID') || '').trim();
    const outputFolder = outputFolderId ? DriveApp.getFolderById(outputFolderId) : null;
    const archiveFolder = archiveFolderId ? DriveApp.getFolderById(archiveFolderId) : null;

    const meta = buildCatalogSlidesMeta_(group, groupCol, sheetDefinitions, productGroup, catalogFileName);

    try {
      const slidesFolder = outputFolder ? getOrCreateDriveSubfolder_(outputFolder, 'Slides') : null;
      trashExistingCatalogSlidesDecks_(outputFolder, slidesFolder, meta);

      const deckFile = buildCatalogSlidesDeck_(meta, pages, sheetDefinitions, slidesFolder);
      validateCatalogSlidesDeck_(
        deckFile.getId(),
        1 + pages.length + (meta.termsTemplateId ? 1 : 0),
        meta.deckName
      );
      const pdfBlob = deckFile.getBlob().setName(catalogFileName);
      archiveExistingDriveFiles_(outputFolder, archiveFolder, [catalogFileName]);
      const pdfFile = outputFolder ? outputFolder.createFile(pdfBlob) : DriveApp.createFile(pdfBlob);

      appendCatalogLog_(
        logSheet,
        productGroup,
        'Generate Slides catalog PDF',
        'Success',
        rows.length,
        pdfFile.getId(),
        '',
        `Created ${catalogFileName} from Slides deck ${deckFile.getId()}. PDF: ${pdfFile.getUrl()}`,
        buildProductionLogDetails_('catalog_pdf', productGroup, catalogFileName, jobStartedAt, getCatalogTimestamp_())
      );

      Logger.log(JSON.stringify({
        productGroup,
        catalogFileName,
        skuRows: rows.length,
        sections: sections.length,
        productPages: pages.length,
        deckId: deckFile.getId(),
        deckUrl: deckFile.getUrl(),
        pdfId: pdfFile.getId(),
        pdfUrl: pdfFile.getUrl()
      }, null, 2));
      clearCatalogSlidesRetryJob_(productGroup, catalogFileName);

    } catch (err) {
      const attempt = Number(retryAttempt) || 0;
      if (isTransientSlidesError_(err) && attempt < 3) {
        const nextAttempt = attempt + 1;

        try {
          scheduleCatalogSlidesRetry_(productGroup, catalogFileName, nextAttempt);
          appendCatalogLog_(
            logSheet,
            productGroup,
            'Generate Slides catalog PDF',
            'Retry scheduled',
            rows.length,
            '',
            '',
            `${catalogFileName}: transient Slides error. Retry ${nextAttempt} of 3 scheduled. ${err.message}`,
            buildProductionLogDetails_('catalog_pdf', productGroup, catalogFileName, jobStartedAt, getCatalogTimestamp_())
          );
          Logger.log(`${catalogFileName}: transient Slides error; retry ${nextAttempt} of 3 scheduled.`);
          return;
        } catch (scheduleErr) {
          err = new Error(`${err.message} Automatic retry scheduling also failed: ${scheduleErr.message}`);
        }
      }

      appendCatalogLog_(
        logSheet,
        productGroup,
        'Generate Slides catalog PDF',
        'Error',
        rows.length,
        '',
        '',
        err.message,
        buildProductionLogDetails_('catalog_pdf', productGroup, catalogFileName, jobStartedAt, getCatalogTimestamp_())
      );
      throw err;
    }
  }
}

function runCatalogSlidesProductionStep_(job, progressCallback) {
  const context = getCatalogSlidesProductionContext_(job);
  if (!context.rows.length) {
    appendCatalogLog_(context.logSheet, context.productGroup, 'Generate Slides catalog PDF', 'Skipped', 0, '', '', 'No active SKU rows found.');
    return { completed: true, job };
  }

  let nextJob = Object.assign({}, job, {
    totalPages: context.pages.length,
    totalSections: context.sections.length,
    totalSkuRows: context.rows.length
  });

  const stage = String(nextJob.stage || 'init');
  if (!nextJob.stageStartedAt) nextJob.stageStartedAt = getCatalogTimestamp_();
  if (!nextJob.progressAnchorAt) nextJob.progressAnchorAt = nextJob.stageStartedAt;

  if (stage === 'init') {
    const slidesFolder = context.outputFolder ? getOrCreateDriveSubfolder_(context.outputFolder, 'Slides') : null;
    trashExistingCatalogSlidesDecks_(context.outputFolder, slidesFolder, context.meta);
    const deckFile = createCatalogSlidesDeckFile_(context.meta, slidesFolder);
    initializeCatalogSlidesDeck_(deckFile.getId(), context.meta);

    nextJob.deckFileId = deckFile.getId();
    nextJob.deckName = context.meta.deckName;
    nextJob.nextPageIndex = 0;
    nextJob.progressPageIndex = 0;
    nextJob.stage = context.pages.length ? 'append_pages' : 'append_terms';
    nextJob.stageStartedAt = getCatalogTimestamp_();
    nextJob.progressAnchorAt = nextJob.stageStartedAt;
    return { completed: false, job: nextJob };
  }

  if (stage === 'append_pages') {
    const startIndex = Math.max(0, Number(nextJob.nextPageIndex) || 0);
    const chunkSize = 8;
    const endIndex = Math.min(context.pages.length, startIndex + chunkSize);
    appendCatalogSlidesPagesChunk_(
      nextJob.deckFileId,
      context.meta,
      context.pages,
      context.sheetDefinitions,
      startIndex,
      endIndex,
      nextJob.runId || '',
      function(partialNextPageIndex) {
        nextJob.nextPageIndex = partialNextPageIndex;
        nextJob.progressPageIndex = partialNextPageIndex;
        nextJob.progressAnchorAt = getCatalogTimestamp_();
        if (typeof progressCallback === 'function') {
          progressCallback(Object.assign({}, nextJob));
        }
      }
    );

    nextJob.nextPageIndex = endIndex;
    nextJob.progressPageIndex = endIndex;
    nextJob.progressAnchorAt = getCatalogTimestamp_();
    if (endIndex >= context.pages.length) {
      nextJob.stage = 'append_terms';
      nextJob.stageStartedAt = nextJob.progressAnchorAt;
    } else {
      nextJob.stage = 'append_pages';
    }
    return { completed: false, job: nextJob };
  }

  if (stage === 'append_terms') {
    if (nextJob.termsAppended) {
      nextJob.stage = 'export_pdf';
      nextJob.stageStartedAt = getCatalogTimestamp_();
      nextJob.progressAnchorAt = nextJob.stageStartedAt;
      return { completed: false, job: nextJob };
    }
    if (context.meta.termsTemplateId) {
      appendCatalogSlidesTermsPage_(nextJob.deckFileId, context.meta, context.pages.length + 2);
      nextJob.termsAppended = true;
    }
    nextJob.stage = 'export_pdf';
    nextJob.stageStartedAt = getCatalogTimestamp_();
    nextJob.progressAnchorAt = nextJob.stageStartedAt;
    return { completed: false, job: nextJob };
  }

  if (stage === 'export_pdf') {
    const deckFile = DriveApp.getFileById(nextJob.deckFileId);
    const expectedSlideCount = 1 + context.pages.length + (context.meta.termsTemplateId ? 1 : 0);
    validateCatalogSlidesDeck_(nextJob.deckFileId, expectedSlideCount, context.meta.deckName);
    const pdfBlob = deckFile.getBlob().setName(context.catalogFileName);
    archiveExistingDriveFiles_(context.outputFolder, context.archiveFolder, [context.catalogFileName]);
    const pdfFile = context.outputFolder ? context.outputFolder.createFile(pdfBlob) : DriveApp.createFile(pdfBlob);

    appendCatalogLog_(
      context.logSheet,
      context.productGroup,
      'Generate Slides catalog PDF',
      'Success',
      context.rows.length,
      pdfFile.getId(),
      '',
      `Created ${context.catalogFileName} from Slides deck ${deckFile.getId()}. PDF: ${pdfFile.getUrl()}`,
      buildProductionLogDetails_(
        'catalog_pdf',
        context.productGroup,
        context.catalogFileName,
        nextJob.startedAt || '',
        getCatalogTimestamp_(),
        { runId: nextJob.runId || '' }
      )
    );

    Logger.log(JSON.stringify({
      productGroup: context.productGroup,
      catalogFileName: context.catalogFileName,
      skuRows: context.rows.length,
      sections: context.sections.length,
      productPages: context.pages.length,
      deckId: deckFile.getId(),
      deckUrl: deckFile.getUrl(),
      pdfId: pdfFile.getId(),
      pdfUrl: pdfFile.getUrl()
    }, null, 2));

    return {
      completed: true,
      job: Object.assign({}, nextJob, {
        pdfFileId: pdfFile.getId(),
        stage: 'complete',
        nextPageIndex: context.pages.length
      })
    };
  }

  throw new Error(`Unsupported catalog production stage: ${stage}`);
}

function getCatalogSlidesProductionContext_(job) {
  const ss = getCatalogWorkbook_();
  const skuSheet = ss.getSheetByName('Catalog_SKUs');
  const groupSheet = ss.getSheetByName('Catalog_Groups');
  const logSheet = ss.getSheetByName('Generation_Log');
  if (!skuSheet) throw new Error('Missing Catalog_SKUs sheet.');
  if (!groupSheet) throw new Error('Missing Catalog_Groups sheet.');

  const skuValues = skuSheet.getDataRange().getValues();
  const skuCol = headerMap_(skuValues[0]);
  const groupValues = groupSheet.getDataRange().getValues();
  const groupCol = headerMap_(groupValues[0]);
  const sheetDefinitions = getSheetDefinitions_(ss);
  const productGroup = String(job.productGroup || '').trim();
  const catalogFileName = String(job.fileName || '').trim();
  const group = findCatalogGroupForJob_(groupValues.slice(1), groupCol, productGroup, catalogFileName);
  if (!group) throw new Error(`Catalog group row not found for ${productGroup} / ${catalogFileName}.`);

  const groupPlc = getOptionalCellValue_(group, groupCol, 'PLC');
  const rows = skuValues.slice(1)
    .filter(row =>
      isTruthy_(row[skuCol.Active]) &&
      String(row[skuCol.Product_Group]).trim() === productGroup &&
      matchesOptionalGroupPlc_(row, skuCol, groupPlc) &&
      row[skuCol.Item_Number]
    )
    .sort((a, b) => compareCatalogSkuRows_(a, b, skuCol));
  const sections = buildCatalogSectionModels_(rows, skuCol);
  const pages = buildCatalogSlidePages_(sections);
  const meta = buildCatalogSlidesMeta_(group, groupCol, sheetDefinitions, productGroup, catalogFileName);
  const outputFolderId = String(getConfigValue_(group, groupCol, sheetDefinitions, 'PDF_Output_Folder_ID') || '').trim();
  const archiveFolderId = String(getConfigValue_(group, groupCol, sheetDefinitions, 'PDF_Archive_Folder_ID') || '').trim();

  return {
    ss,
    logSheet,
    sheetDefinitions,
    productGroup,
    catalogFileName,
    group,
    rows,
    sections,
    pages,
    meta,
    outputFolder: outputFolderId ? DriveApp.getFolderById(outputFolderId) : null,
    archiveFolder: archiveFolderId ? DriveApp.getFolderById(archiveFolderId) : null
  };
}

function findCatalogGroupForJob_(groupRows, groupCol, productGroup, catalogFileName) {
  const matches = [];

  for (let i = 0; i < groupRows.length; i++) {
    const row = groupRows[i];
    if (String(row[groupCol.Product_Group] || '').trim() !== productGroup) continue;
    if (String(row[groupCol.Catalog_File_Name] || '').trim() !== catalogFileName) continue;
    matches.push(row);
  }

  if (!matches.length) return null;

  matches.sort((a, b) => {
    const aGeneratePdf = isTruthy_(a[groupCol.Generate_PDF]) ? 1 : 0;
    const bGeneratePdf = isTruthy_(b[groupCol.Generate_PDF]) ? 1 : 0;
    if (aGeneratePdf !== bGeneratePdf) return bGeneratePdf - aGeneratePdf;

    const aPlc = String(getOptionalCellValue_(a, groupCol, 'PLC') || '').trim();
    const bPlc = String(getOptionalCellValue_(b, groupCol, 'PLC') || '').trim();
    const aBlankPlc = aPlc ? 0 : 1;
    const bBlankPlc = bPlc ? 0 : 1;
    if (aBlankPlc !== bBlankPlc) return bBlankPlc - aBlankPlc;

    return 0;
  });

  return matches[0];
}

function createCatalogSlidesDeckFile_(meta, outputFolder) {
  if (!meta.coverTemplateId) throw new Error('Missing Cover_Template_ID.');
  const coverTemplateFile = DriveApp.getFileById(meta.coverTemplateId);
  return outputFolder
    ? coverTemplateFile.makeCopy(meta.deckName, outputFolder)
    : coverTemplateFile.makeCopy(meta.deckName);
}

function resetCopiedCatalogDeckToCoverSlide_(presentation) {
  const slides = presentation.getSlides();
  if (!slides.length) throw new Error('Cover template has no slides.');

  for (let index = slides.length - 1; index >= 1; index--) {
    slides[index].remove();
  }

  return presentation.getSlides()[0];
}

function trimCatalogDeckToSlideCount_(presentation, expectedSlideCount) {
  const slides = presentation.getSlides();
  if (!slides.length) throw new Error('Catalog deck has no slides.');

  const targetCount = Math.max(1, Number(expectedSlideCount) || 1);
  for (let index = slides.length - 1; index >= targetCount; index--) {
    slides[index].remove();
  }
}

function initializeCatalogSlidesDeck_(deckFileId, meta) {
  let presentation = runSlidesOperationWithRetry_(
    `Open ${meta.deckName}`,
    () => SlidesApp.openById(deckFileId)
  );
  const coverSlide = resetCopiedCatalogDeckToCoverSlide_(presentation);
  fillCatalogSlidePlaceholders_(coverSlide, buildCatalogTextPlaceholders_(meta, 1, ''));
  replaceCatalogImagePlaceholder_(coverSlide, '{{PRODUCT_GROUP_IMAGE}}', meta.productGroupImageId);
  runSlidesOperationWithRetry_(`Save ${meta.deckName}`, () => presentation.saveAndClose());
}

function appendCatalogSlidesPagesChunk_(deckFileId, meta, pages, sheetDefinitions, startIndex, endIndex, runId, progressCallback) {
  throwCatalogProductionCancelledIfNeeded_(runId);
  let presentation = runSlidesOperationWithRetry_(
    `Open ${meta.deckName}`,
    () => SlidesApp.openById(deckFileId)
  );
  const templateSlides = {};

  if (startIndex === 0) {
    resetCopiedCatalogDeckToCoverSlide_(presentation);
  } else {
    trimCatalogDeckToSlideCount_(presentation, startIndex + 1);
  }

  for (let index = startIndex; index < endIndex; index++) {
    throwCatalogProductionCancelledIfNeeded_(runId);
    const page = pages[index];
    const templateId = String(sheetDefinitions[page.template] || '').trim();
    if (!templateId) throw new Error(`Missing template ID for ${page.template}.`);

    if (!templateSlides[templateId]) {
      const templatePresentation = runSlidesOperationWithRetry_(
        `Open template ${page.template}`,
        () => SlidesApp.openById(templateId)
      );
      templateSlides[templateId] = templatePresentation.getSlides()[0];
    }

    const slide = presentation.appendSlide(templateSlides[templateId]);
    fillCatalogProductSlide_(slide, meta, page, index + 2, runId);
    throwCatalogProductionCancelledIfNeeded_(runId);
    runSlidesOperationWithRetry_(`Save ${meta.deckName} page ${index + 1}`, () => presentation.saveAndClose());
    if (typeof progressCallback === 'function') {
      progressCallback(index + 1);
    }
    if (index + 1 < endIndex) {
      throwCatalogProductionCancelledIfNeeded_(runId);
      presentation = runSlidesOperationWithRetry_(
        `Reopen ${meta.deckName}`,
        () => SlidesApp.openById(deckFileId)
      );
    }
  }
}

function appendCatalogSlidesTermsPage_(deckFileId, meta, pageNumber) {
  let presentation = runSlidesOperationWithRetry_(
    `Open ${meta.deckName}`,
    () => SlidesApp.openById(deckFileId)
  );
  trimCatalogDeckToSlideCount_(presentation, pageNumber - 1);
  const termsPresentation = runSlidesOperationWithRetry_(
    'Open terms template',
    () => SlidesApp.openById(meta.termsTemplateId)
  );
  const termsSlide = presentation.appendSlide(termsPresentation.getSlides()[0]);
  fillCatalogSlidePlaceholders_(termsSlide, buildCatalogTextPlaceholders_(meta, pageNumber, ''));
  runSlidesOperationWithRetry_(`Save ${meta.deckName}`, () => presentation.saveAndClose());
}

function validateCatalogSlidesDeck_(deckFileId, expectedSlideCount, deckName) {
  const presentation = runSlidesOperationWithRetry_(
    `Validate ${deckName || deckFileId}`,
    () => SlidesApp.openById(deckFileId)
  );
  const actualSlideCount = presentation.getSlides().length;
  runSlidesOperationWithRetry_(
    `Close ${deckName || deckFileId}`,
    () => presentation.saveAndClose()
  );

  if (actualSlideCount !== expectedSlideCount) {
    throw new Error(
      `Catalog deck slide-count mismatch for ${deckName || deckFileId}. ` +
      `Expected ${expectedSlideCount}, found ${actualSlideCount}.`
    );
  }
}

function buildCatalogSlidesMeta_(group, groupCol, sheetDefinitions, productGroup, catalogFileName) {
  const subtitle = String(group[groupCol.Catalog_Subtitle] || '').trim();
  let subheading = String(getOptionalCellValue_(group, groupCol, 'Catalog_Subheading') || '').trim();

  if (isForgedSteelCatalogProductGroup_(productGroup)) {
    subheading = 'ASTM A105 Carbon Steel';
  }

  return {
    productGroup,
    catalogFileName,
    deckName: catalogFileName.replace(/\.pdf$/i, '') + ' - Slides Catalog',
    title: String(group[groupCol.Catalog_Title] || productGroup).trim(),
    subtitle,
    subheading,
    catalogCode: String(group[groupCol.Catalog_Code] || '').trim(),
    versionCode: String(group[groupCol.Version_Code] || '').trim(),
    effectiveDate: formatCatalogEffectiveDate_(getOptionalCellValue_(group, groupCol, 'Effective_Date')),
    productGroupImageId: String(getOptionalCellValue_(group, groupCol, 'Product_Group_Image_ID') || '').trim(),
    keyBullet1: String(getOptionalCellValue_(group, groupCol, 'Key_Bullet_1') || '').trim(),
    keyBullet2: String(getOptionalCellValue_(group, groupCol, 'Key_Bullet_2') || '').trim(),
    keyBullet3: String(getOptionalCellValue_(group, groupCol, 'Key_Bullet_3') || '').trim(),
    keyBullet4: String(getOptionalCellValue_(group, groupCol, 'Key_Bullet_4') || '').trim(),
    coverTemplateId: String(getConfigValue_(group, groupCol, sheetDefinitions, 'Cover_Template_ID') || '').trim(),
    termsTemplateId: String(getConfigValue_(group, groupCol, sheetDefinitions, 'Terms_Template_ID') || '').trim()
  };
}

function buildCatalogSlidesDeck_(meta, pages, sheetDefinitions, outputFolder) {
  if (!meta.coverTemplateId) throw new Error('Missing Cover_Template_ID.');

  const coverTemplateFile = DriveApp.getFileById(meta.coverTemplateId);
  const deckFile = outputFolder
    ? coverTemplateFile.makeCopy(meta.deckName, outputFolder)
    : coverTemplateFile.makeCopy(meta.deckName);
  let presentation = runSlidesOperationWithRetry_(
    `Open ${meta.deckName}`,
    () => SlidesApp.openById(deckFile.getId())
  );
  const templateSlides = {};

  const coverSlide = resetCopiedCatalogDeckToCoverSlide_(presentation);
  fillCatalogSlidePlaceholders_(coverSlide, buildCatalogTextPlaceholders_(meta, 1, ''));
  replaceCatalogImagePlaceholder_(coverSlide, '{{PRODUCT_GROUP_IMAGE}}', meta.productGroupImageId);

  pages.forEach((page, index) => {
    const templateId = String(sheetDefinitions[page.template] || '').trim();
    if (!templateId) throw new Error(`Missing template ID for ${page.template}.`);

    if (!templateSlides[templateId]) {
      const templatePresentation = runSlidesOperationWithRetry_(
        `Open template ${page.template}`,
        () => SlidesApp.openById(templateId)
      );
      templateSlides[templateId] = templatePresentation.getSlides()[0];
    }

    const templateSlide = templateSlides[templateId];
    const slide = presentation.appendSlide(templateSlide);
    fillCatalogProductSlide_(slide, meta, page, index + 2);

    const shouldCheckpoint = (index + 1) % 6 === 0 && index < pages.length - 1;
    if (shouldCheckpoint) {
      runSlidesOperationWithRetry_(
        `Checkpoint ${meta.deckName} after product page ${index + 1}`,
        () => presentation.saveAndClose()
      );
      Utilities.sleep(500);
      presentation = runSlidesOperationWithRetry_(
        `Reopen ${meta.deckName}`,
        () => SlidesApp.openById(deckFile.getId())
      );
    }
  });

  if (meta.termsTemplateId) {
    const termsPresentation = runSlidesOperationWithRetry_(
      'Open terms template',
      () => SlidesApp.openById(meta.termsTemplateId)
    );
    const termsSlide = presentation.appendSlide(termsPresentation.getSlides()[0]);
    fillCatalogSlidePlaceholders_(termsSlide, buildCatalogTextPlaceholders_(meta, pages.length + 2, ''));
  }

  runSlidesOperationWithRetry_(
    `Save ${meta.deckName}`,
    () => presentation.saveAndClose()
  );
  return deckFile;
}

function runSlidesOperationWithRetry_(operationName, operation) {
  const maxAttempts = 3;
  let lastError;

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      return operation();
    } catch (err) {
      lastError = err;
      if (!isTransientSlidesError_(err) || attempt === maxAttempts) throw err;

      const delayMs = 1500 * Math.pow(2, attempt - 1);
      Logger.log(`${operationName} failed with a transient Slides error. Retrying in ${delayMs}ms (${attempt}/${maxAttempts}).`);
      Utilities.sleep(delayMs);
    }
  }

  throw lastError;
}

function isTransientSlidesError_(err) {
  const message = String(err && err.message ? err.message : err || '');
  return /service unavailable.*slides|slides.*service unavailable|internal error|timed out|try again|rate limit|too many calls|service invoked too many times/i.test(message);
}

function retryCatalogSlidesGeneration_() {
  const properties = PropertiesService.getScriptProperties();
  const propertyName = 'CATALOG_SLIDES_RETRY_JOB';
  const serializedJob = properties.getProperty(propertyName);

  properties.deleteProperty(propertyName);
  deleteCatalogSlidesRetryTriggers_();

  if (!serializedJob) {
    Logger.log('No pending catalog Slides retry job found.');
    return;
  }

  const job = JSON.parse(serializedJob);
  Logger.log(`Starting catalog Slides retry ${job.attempt} of 3 for ${job.catalogFileName}.`);
  generateCatalogSlidesPDFs(job.productGroup, job.catalogFileName, job.attempt);
}

function scheduleCatalogSlidesRetry_(productGroup, catalogFileName, attempt) {
  const properties = PropertiesService.getScriptProperties();
  const propertyName = 'CATALOG_SLIDES_RETRY_JOB';
  const job = {
    productGroup,
    catalogFileName,
    attempt,
    scheduledAt: getCatalogTimestamp_()
  };

  deleteCatalogSlidesRetryTriggers_();
  properties.setProperty(propertyName, JSON.stringify(job));

  try {
    ScriptApp.newTrigger('retryCatalogSlidesGeneration_')
      .timeBased()
      .after(Math.max(60000, Number(attempt) * 60000))
      .create();
  } catch (err) {
    properties.deleteProperty(propertyName);
    throw err;
  }
}

function clearCatalogSlidesRetryJob_(productGroup, catalogFileName) {
  const properties = PropertiesService.getScriptProperties();
  const propertyName = 'CATALOG_SLIDES_RETRY_JOB';
  const serializedJob = properties.getProperty(propertyName);
  if (!serializedJob) return;

  try {
    const job = JSON.parse(serializedJob);
    if (job.productGroup !== productGroup || job.catalogFileName !== catalogFileName) return;
  } catch (err) {
    Logger.log(`Discarding invalid catalog retry state: ${err.message}`);
  }

  properties.deleteProperty(propertyName);
  deleteCatalogSlidesRetryTriggers_();
}

function deleteCatalogSlidesRetryTriggers_() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'retryCatalogSlidesGeneration_') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function startFullCatalogProductionRun() {
  return startCatalogProductionRun_(true, true, 'Catalog PDFs and price files');
}

function startFullCatalogSlidesGeneration() {
  return startCatalogProductionRun_(true, false, 'Catalog PDFs');
}

function startFullPriceFileGeneration() {
  return startCatalogProductionRun_(false, true, 'Price files');
}

function startCatalogProductionRun_(includeCatalogPdfs, includePriceFiles, mode) {
  return queueCatalogProductionSelection_({
    includeCatalogPdfs,
    includePriceFiles,
    requireActive: true,
    mode
  });
}

function queueExplicitCatalogProductionJobs_(jobs, mode) {
  const explicitJobs = prependAutomaticCartonNormalizationJob_((jobs || []).slice());
  if (!explicitJobs.length) throw new Error(`No jobs found for ${mode}.`);

  return queueCatalogProductionJobs_(explicitJobs, mode, {
    catalogJobCount: explicitJobs.filter(job => job.type === 'catalog_pdf').length,
    priceJobCount: explicitJobs.filter(job => job.type === 'price_file').length
  });
}

function queueCatalogProductionSelection_(options) {
  const mode = options.mode;
  const properties = PropertiesService.getScriptProperties();
  const ss = getCatalogWorkbook_();
  const groupSheet = ss.getSheetByName('Catalog_Groups');
  if (!groupSheet) throw new Error('Missing Catalog_Groups sheet.');

  const values = groupSheet.getDataRange().getValues();
  const col = headerMap_(values[0]);
  const jobs = prependAutomaticCartonNormalizationJob_(buildCatalogProductionJobs_(values.slice(1), col, options));
  const catalogJobs = jobs.filter(job => job.type === 'catalog_pdf');
  const priceJobs = jobs.filter(job => job.type === 'price_file');

  return queueCatalogProductionJobs_(jobs, mode, {
    catalogJobCount: catalogJobs.length,
    priceJobCount: priceJobs.length
  });
}

function queueCatalogProductionJobs_(jobs, mode, counts) {
  const properties = PropertiesService.getScriptProperties();
  const propertyName = 'CATALOG_PRODUCTION_QUEUE';
  const existingQueue = properties.getProperty(propertyName);

  if (existingQueue) {
    const existingState = JSON.parse(existingQueue);
    throw new Error(
      `A catalog production run is already active (${existingState.completed}/${existingState.totalJobs} completed). ` +
      'Run getCatalogProductionRunStatus() for details or cancelCatalogProductionRun() before starting over.'
    );
  }

  if (!jobs.length) throw new Error(`No active jobs found for ${mode}.`);

  const logSheet = getCatalogWorkbook_().getSheetByName('Generation_Log');
  const state = {
    active: true,
    runId: Utilities.getUuid(),
    mode,
    startedAt: getCatalogTimestamp_(),
    updatedAt: getCatalogTimestamp_(),
    jobs,
    totalJobs: jobs.length,
    nextJobIndex: 0,
    currentAttempt: 0,
    completed: 0,
    failed: []
  };
  const serializedState = JSON.stringify(state);
  if (serializedState.length > 8500) {
    throw new Error(`Production queue is too large for Script Properties (${serializedState.length} characters).`);
  }

  properties.setProperty(propertyName, serializedState);
  appendCatalogLog_(
    logSheet,
    'Catalog Production',
    mode,
    'Started',
    jobs.length,
    '',
    '',
    `Queued ${(counts && counts.catalogJobCount) || 0} catalog PDFs, ${(counts && counts.priceJobCount) || 0} price files, ` +
    `${jobs.length - (((counts && counts.catalogJobCount) || 0) + ((counts && counts.priceJobCount) || 0))} preparation jobs. ` +
    `Run ID: ${state.runId}`
  );
  Logger.log(`Started ${mode} production run ${state.runId} with ${jobs.length} jobs.`);

  scheduleCatalogProductionTrigger_(1000);
  return {
    ok: true,
    message:
      `Queued ${jobs.length} job${jobs.length === 1 ? '' : 's'} for ${mode}. ` +
      'The run now continues on Google infrastructure, so it is safe to close the dialog, browser, or local machine after queue confirmation.',
    status: state
  };
}

function prependAutomaticCartonNormalizationJob_(jobs) {
  const queue = (jobs || []).slice();
  const needsGenerationNormalization = queue.some(job => job.type === 'catalog_pdf' || job.type === 'price_file');
  const alreadyIncluded = queue.some(job => job.type === 'repair_cartons');

  if (needsGenerationNormalization && !alreadyIncluded) {
    queue.unshift({ type: 'repair_cartons', fileName: 'Normalize Inner and Master Cartons' });
  }

  return queue;
}

function buildCatalogProductionJobs_(groupRows, col, options) {
  const includeCatalogPdfs = !!options.includeCatalogPdfs;
  const includePriceFiles = !!options.includePriceFiles;
  const requireActive = options.requireActive !== false;
  const onlyProductGroup = String(options.onlyProductGroup || '').trim();
  const onlyProductGroups = (options.onlyProductGroups || [])
    .map(value => String(value || '').trim())
    .filter(Boolean);
  const onlyProductGroupLookup = onlyProductGroups.reduce((lookup, value) => {
    lookup[value] = true;
    return lookup;
  }, {});
  const catalogJobs = [];
  const priceJobs = [];
  const seenJobs = {};

  groupRows.forEach(row => {
    const productGroup = String(row[col.Product_Group] || '').trim();
    if (!productGroup) return;
    if (requireActive && !isTruthy_(row[col.Active])) return;
    if (onlyProductGroup && productGroup !== onlyProductGroup) return;
    if (onlyProductGroups.length && !onlyProductGroupLookup[productGroup]) return;

    if (includeCatalogPdfs && isTruthy_(row[col.Generate_PDF]) && row[col.Catalog_File_Name]) {
      const fileName = String(row[col.Catalog_File_Name]).trim();
      const key = `catalog_pdf|${productGroup}|${fileName}`;
      if (!seenJobs[key]) {
        seenJobs[key] = true;
        catalogJobs.push({ type: 'catalog_pdf', productGroup, fileName });
      }
    }

    if (includePriceFiles && isTruthy_(row[col.Generate_XLS]) && row[col.Price_File_Name]) {
      const fileName = String(row[col.Price_File_Name]).trim();
      const key = `price_file|${productGroup}|${fileName}`;
      if (!seenJobs[key]) {
        seenJobs[key] = true;
        priceJobs.push({ type: 'price_file', productGroup, fileName });
      }
    }
  });

  return catalogJobs.concat(priceJobs);
}

function runNextCatalogProductionBatch_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    Logger.log('Skipped overlapping production trigger because another batch is already running.');
    return;
  }

  try {
  const properties = PropertiesService.getScriptProperties();
  const propertyName = 'CATALOG_PRODUCTION_QUEUE';
  const serializedState = properties.getProperty(propertyName);

  deleteCatalogProductionTriggers_();
  if (!serializedState) {
    Logger.log('No active catalog production queue found.');
    return;
  }

  const state = JSON.parse(serializedState);
  if (isCatalogProductionRunCancelled_(state.runId)) {
    finalizeCanceledProductionState_(state);
    return;
  }
  if (state.nextJobIndex >= state.jobs.length) {
    return completeCatalogProductionRun_(state);
  }

  const job = Object.assign({}, state.jobs[state.nextJobIndex]);
  if (!job.startedAt) {
    job.startedAt = getCatalogTimestamp_();
    job.runId = state.runId;
    state.jobs[state.nextJobIndex] = job;
  }
  if (state.currentAttempt >= 3) {
    const failedEndedAt = getCatalogTimestamp_();
    job.endedAt = failedEndedAt;
    job.durationMs = getDurationMsBetweenIso_(job.startedAt, failedEndedAt);
    state.jobs[state.nextJobIndex] = job;
    state.failed.push({
      type: job.type,
      productGroup: job.productGroup,
      fileName: job.fileName,
      message: 'Job did not finish within three execution attempts.'
    });
    state.nextJobIndex++;
    state.currentAttempt = 0;
    state.updatedAt = getCatalogTimestamp_();
    saveCatalogProductionState_(state);
    return scheduleOrCompleteCatalogProduction_(state);
  }

  state.currentAttempt++;
  state.currentJobStartedAt = job.startedAt || getCatalogTimestamp_();
  state.updatedAt = getCatalogTimestamp_();
  saveCatalogProductionState_(state);
  const batchStartedAt = Date.now();
  const batchDeadlineMs = batchStartedAt + (4.5 * 60 * 1000);

  // A consumed time trigger cannot recover from a hard execution timeout, so schedule a watchdog first.
  scheduleCatalogProductionTrigger_(8 * 60 * 1000);
  Logger.log(
    `Production job ${state.nextJobIndex + 1}/${state.totalJobs}, attempt ${state.currentAttempt}/3: ` +
    `${job.type} ${job.fileName}`
  );

  try {
    if (job.type === 'catalog_pdf') {
      const catalogStep = runCatalogSlidesProductionJobUntilPause_(job, batchDeadlineMs, function(partialJob) {
        state.jobs[state.nextJobIndex] = partialJob;
        state.updatedAt = getCatalogTimestamp_();
        if (!saveCatalogProductionState_(state)) {
          throw new Error('__CATALOG_PRODUCTION_CANCELLED__');
        }
      });
      state.jobs[state.nextJobIndex] = catalogStep.job;

      if (!catalogStep.completed) {
        deleteCatalogProductionTriggers_();
        state.currentAttempt = 0;
        state.updatedAt = getCatalogTimestamp_();
        if (!saveCatalogProductionState_(state)) {
          finalizeCanceledProductionState_(state);
          return;
        }
        scheduleCatalogProductionTrigger_(5000);
        return;
      }
    } else if (job.type === 'price_file') {
      generateCatalogPriceFiles(job.productGroup, job.fileName, { skipCartonNormalization: true });
    } else if (job.type === 'pricing_calc') {
      const pricingStep = runChunkedProductionJobUntilPause_(job, batchDeadlineMs, runCatalogPricingCalculationStep_);
      state.jobs[state.nextJobIndex] = pricingStep.job;

      if (!pricingStep.completed) {
        deleteCatalogProductionTriggers_();
        state.currentAttempt = 0;
        state.updatedAt = getCatalogTimestamp_();
        if (!saveCatalogProductionState_(state)) {
          finalizeCanceledProductionState_(state);
          return;
        }
        scheduleCatalogProductionTrigger_(5000);
        return;
      }
    } else if (job.type === 'repair_cartons') {
      const repairStep = runChunkedProductionJobUntilPause_(job, batchDeadlineMs, runCatalogCartonRepairStep_);
      state.jobs[state.nextJobIndex] = repairStep.job;

      if (!repairStep.completed) {
        deleteCatalogProductionTriggers_();
        state.currentAttempt = 0;
        state.updatedAt = getCatalogTimestamp_();
        if (!saveCatalogProductionState_(state)) {
          finalizeCanceledProductionState_(state);
          return;
        }
        scheduleCatalogProductionTrigger_(5000);
        return;
      }
    } else {
      throw new Error(`Unsupported production job type: ${job.type}`);
    }

    deleteCatalogProductionTriggers_();
    const completedEndedAt = getCatalogTimestamp_();
    state.jobs[state.nextJobIndex] = Object.assign({}, state.jobs[state.nextJobIndex], {
      endedAt: completedEndedAt,
      durationMs: getDurationMsBetweenIso_(state.jobs[state.nextJobIndex].startedAt, completedEndedAt)
    });
    state.completed++;
    state.nextJobIndex++;
    state.currentAttempt = 0;
    state.currentJobStartedAt = '';
    state.updatedAt = getCatalogTimestamp_();
    if (!saveCatalogProductionState_(state)) {
      finalizeCanceledProductionState_(state);
      return;
    }
    return scheduleOrCompleteCatalogProduction_(state);

  } catch (err) {
    deleteCatalogProductionTriggers_();
    if (String(err && err.message ? err.message : err || '') === '__CATALOG_PRODUCTION_CANCELLED__') {
      finalizeCanceledProductionState_(state);
      return;
    }
    const isTransient = isTransientProductionError_(err);

    if (isTransient && state.currentAttempt < 3) {
      state.updatedAt = getCatalogTimestamp_();
      state.lastError = String(err.message || err).slice(0, 500);
      if (!saveCatalogProductionState_(state)) {
        finalizeCanceledProductionState_(state);
        return;
      }
      scheduleCatalogProductionTrigger_(state.currentAttempt * 60000);
      Logger.log(`${job.fileName}: transient production error; attempt ${state.currentAttempt + 1} scheduled.`);
      return;
    }

    const failedEndedAt = getCatalogTimestamp_();
    state.jobs[state.nextJobIndex] = Object.assign({}, state.jobs[state.nextJobIndex], {
      endedAt: failedEndedAt,
      durationMs: getDurationMsBetweenIso_(state.jobs[state.nextJobIndex].startedAt, failedEndedAt)
    });
    state.failed.push({
      type: job.type,
      productGroup: job.productGroup,
      fileName: job.fileName,
      message: String(err.message || err).slice(0, 500)
    });
    state.nextJobIndex++;
    state.currentAttempt = 0;
    state.currentJobStartedAt = '';
    state.lastError = '';
    state.updatedAt = getCatalogTimestamp_();
    if (!saveCatalogProductionState_(state)) {
      finalizeCanceledProductionState_(state);
      return;
    }
    return scheduleOrCompleteCatalogProduction_(state);
  }
  } finally {
    lock.releaseLock();
  }
}

function scheduleOrCompleteCatalogProduction_(state) {
  if (state.nextJobIndex >= state.jobs.length) {
    return completeCatalogProductionRun_(state);
  }

  scheduleCatalogProductionTrigger_(5000);
  Logger.log(
    `Production progress: ${state.completed} completed, ${state.failed.length} failed, ` +
    `${state.totalJobs - state.nextJobIndex} remaining.`
  );
}

function scheduleCatalogProductionTrigger_(delayMs) {
  deleteCatalogProductionTriggers_();
  ScriptApp.newTrigger('runNextCatalogProductionBatch_')
    .timeBased()
    .after(Math.max(5000, Number(delayMs) || 5000))
    .create();
}

function runCatalogSlidesProductionJobUntilPause_(job, deadlineMs, progressCallback) {
  let nextJob = Object.assign({}, job);
  let step = { completed: false, job: nextJob };

  while (Date.now() < deadlineMs - 20000) {
    step = runCatalogSlidesProductionStep_(nextJob, progressCallback);
    nextJob = step.job;
    if (step.completed) return step;
  }

  return { completed: false, job: nextJob };
}

function runChunkedProductionJobUntilPause_(job, deadlineMs, stepRunner) {
  let nextJob = Object.assign({}, job);
  let step = { completed: false, job: nextJob };

  while (Date.now() < deadlineMs - 15000) {
    step = stepRunner(nextJob);
    nextJob = step.job;
    if (step.completed) return step;
  }

  return { completed: false, job: nextJob };
}

function saveCatalogProductionState_(state) {
  state.active = true;
  state.updatedAt = getCatalogTimestamp_();
  if (isCatalogProductionRunCancelled_(state.runId)) {
    const properties = PropertiesService.getScriptProperties();
    properties.deleteProperty('CATALOG_PRODUCTION_QUEUE');
    return false;
  }
  PropertiesService.getScriptProperties().setProperty('CATALOG_PRODUCTION_QUEUE', JSON.stringify(state));
  return true;
}

function completeCatalogProductionRun_(state) {
  const properties = PropertiesService.getScriptProperties();
  const logSheet = getCatalogWorkbook_().getSheetByName('Generation_Log');
  const completedAt = getCatalogTimestamp_();
  const summary = {
    active: false,
    status: state.failed && state.failed.length ? 'completed_with_errors' : 'completed',
    runId: state.runId,
    mode: state.mode,
    startedAt: state.startedAt,
    completedAt,
    totalJobs: state.totalJobs,
    completed: state.completed,
    succeeded: state.completed,
    failedCount: (state.failed || []).length,
    remaining: Math.max(0, state.totalJobs - state.completed - ((state.failed || []).length)),
    failed: state.failed || []
  };

  deleteCatalogProductionTriggers_();
  properties.deleteProperty('CATALOG_PRODUCTION_QUEUE');
  properties.setProperty('CATALOG_PRODUCTION_LAST_RUN', JSON.stringify(summary));
  appendCatalogLog_(
    logSheet,
    'Catalog Production',
    state.mode,
    summary.failed.length ? 'Completed with errors' : 'Completed',
    summary.completed,
    '',
    '',
    `${summary.completed}/${summary.totalJobs} jobs completed; ${summary.failed.length} failed. Run ID: ${summary.runId}`
  );
  Logger.log(JSON.stringify(summary, null, 2));
  return summary;
}

function rerunFailedCatalogProductionJobs_() {
  const state = getCatalogProductionRunState_();
  if (!state || !state.failed || !state.failed.length) {
    throw new Error('No failed jobs were found in the most recent production run.');
  }

  const jobs = state.failed
    .filter(entry => entry && entry.type && entry.productGroup && entry.fileName)
    .map(entry => ({
      type: entry.type,
      productGroup: entry.productGroup,
      fileName: entry.fileName
    }));

  if (!jobs.length) {
    throw new Error('The last run has failures, but no rerunnable failed jobs were captured.');
  }

  const dedupedJobs = [];
  const seen = {};
  jobs.forEach(job => {
    const key = [job.type, job.productGroup, job.fileName].join('|');
    if (seen[key]) return;
    seen[key] = true;
    dedupedJobs.push(job);
  });

  return queueCatalogProductionJobs_(dedupedJobs, 'Failed jobs rerun', {
    catalogJobCount: dedupedJobs.filter(job => job.type === 'catalog_pdf').length,
    priceJobCount: dedupedJobs.filter(job => job.type === 'price_file').length
  });
}

function getCatalogProductionRunStatus() {
  const state = getCatalogProductionRunState_();
  Logger.log(JSON.stringify(state, null, 2));
  return state;
}

function getCatalogProductionRunStatusForSidebar() {
  return getCatalogProductionRunState_();
}

function getCatalogProductionRunState_() {
  const workbook = getCatalogWorkbook_();
  const runtimeStats = getCatalogProductionRuntimeStats_(workbook);
  const properties = PropertiesService.getScriptProperties();
  const active = properties.getProperty('CATALOG_PRODUCTION_QUEUE');
  if (active) {
    const state = JSON.parse(active);
    state.active = true;
    state.runtimeStats = runtimeStats;
    return state;
  }

  const lastRun = JSON.parse(properties.getProperty('CATALOG_PRODUCTION_LAST_RUN') || 'null');
  if (lastRun) {
    lastRun.active = false;
    lastRun.runtimeStats = runtimeStats;
  }
  return lastRun;
}

function cancelCatalogProductionRun() {
  const properties = PropertiesService.getScriptProperties();
  const active = properties.getProperty('CATALOG_PRODUCTION_QUEUE');
  const canceledAt = getCatalogTimestamp_();

  deleteCatalogProductionTriggers_();
  if (active) {
    const state = JSON.parse(active);
    properties.setProperty('CATALOG_PRODUCTION_CANCELLED_RUN_ID', state.runId);
    properties.deleteProperty('CATALOG_PRODUCTION_QUEUE');
    finalizeCanceledProductionState_(state, canceledAt);
  }
  Logger.log(active ? 'Catalog production run cancelled.' : 'No active catalog production run found.');
}

function isCatalogProductionRunCancelled_(runId) {
  if (!runId) return false;
  const canceledRunId = PropertiesService.getScriptProperties().getProperty('CATALOG_PRODUCTION_CANCELLED_RUN_ID');
  return canceledRunId === runId;
}

function throwCatalogProductionCancelledIfNeeded_(runId) {
  if (isCatalogProductionRunCancelled_(runId)) {
    throw new Error('__CATALOG_PRODUCTION_CANCELLED__');
  }
}

function finalizeCanceledProductionState_(state, canceledAtOverride) {
  if (!state) return;
  const properties = PropertiesService.getScriptProperties();
  const canceledAt = canceledAtOverride || getCatalogTimestamp_();
  const summary = {
    active: false,
    status: 'cancelled',
    runId: state.runId,
    mode: state.mode,
    startedAt: state.startedAt,
    canceledAt,
    totalJobs: state.totalJobs,
    completed: state.completed || 0,
    succeeded: state.completed || 0,
    failedCount: (state.failed || []).length,
    remaining: Math.max(0, (state.totalJobs || 0) - (state.completed || 0) - ((state.failed || []).length)),
    failed: state.failed || []
  };
  deleteCatalogProductionTriggers_();
  properties.deleteProperty('CATALOG_PRODUCTION_QUEUE');
  properties.setProperty('CATALOG_PRODUCTION_LAST_RUN', JSON.stringify(summary));
  if (isCatalogProductionRunCancelled_(state.runId)) {
    properties.deleteProperty('CATALOG_PRODUCTION_CANCELLED_RUN_ID');
  }
}

function deleteCatalogProductionTriggers_() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'runNextCatalogProductionBatch_') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function isTransientProductionError_(err) {
  if (isTransientSlidesError_(err)) return true;
  const message = String(err && err.message ? err.message : err || '');
  return /service unavailable|internal error|timed out|try again|rate limit|too many calls|service invoked too many times/i.test(message);
}

function fillCatalogProductSlide_(slide, meta, page, pageNumber, runId) {
  throwCatalogProductionCancelledIfNeeded_(runId);
  const pageMeta = buildCatalogPageMeta_(meta, page);
  fillCatalogSlidePlaceholders_(slide, buildCatalogTextPlaceholders_(pageMeta, pageNumber, ''));

  let previousSectionBottom = null;

  page.sections.forEach((section, index) => {
    throwCatalogProductionCancelledIfNeeded_(runId);
    const sectionNumber = index + 1;
    const titlePlaceholder = `{{SECTION_${sectionNumber}_TITLE}}`;
    const picturePlaceholder = `{{SECTION_${sectionNumber}_PICTURE}}`;
    const tablePlaceholder = `{{SECTION_${sectionNumber}_TABLE}}`;
    const pictureBounds = getCatalogPlaceholderBounds_(slide, picturePlaceholder);

    if (previousSectionBottom !== null) {
      compactCatalogSectionLayout_(slide, sectionNumber, previousSectionBottom);
    }

    const titleElement = findCatalogPlaceholderElement_(slide, titlePlaceholder);
    const title = section.continued ? `${section.title} ${section.continuedLabel}` : section.title;
    replaceCatalogText_(slide, titlePlaceholder, title);
    const specBounds = insertCatalogSectionSpecInfo_(slide, sectionNumber, titleElement, section);
    replaceCatalogImagePlaceholder_(slide, picturePlaceholder, section.imageFileId);
    const tableBounds = replaceCatalogTablePlaceholder_(slide, tablePlaceholder, section);
    if (tableBounds) {
      const sectionBottomCandidates = [tableBounds.bottom];
      if (pictureBounds) sectionBottomCandidates.push(pictureBounds.bottom);
      if (specBounds) sectionBottomCandidates.push(specBounds.bottom);
      previousSectionBottom = Math.max.apply(null, sectionBottomCandidates.concat(previousSectionBottom || 0));
    }
  });

  for (let sectionNumber = page.sections.length + 1; sectionNumber <= 3; sectionNumber++) {
    removeCatalogPlaceholderElement_(slide, `{{SECTION_${sectionNumber}_TITLE}}`);
    removeCatalogPlaceholderElement_(slide, `{{SECTION_${sectionNumber}_PICTURE}}`);
    removeCatalogPlaceholderElement_(slide, `{{SECTION_${sectionNumber}_TABLE}}`);
  }

  replaceCatalogText_(slide, '{{CONTINUED_LABEL}}', page.sections[0] && page.sections[0].continued ? '(continued)' : '');
}

function insertCatalogSectionSpecInfo_(slide, sectionNumber, titleElement, section) {
  const specInfo = String(section && section.specInfo || '').trim();
  if (!specInfo || !titleElement) return null;

  const titleText = String((titleElement.getText && titleElement.getText().asString()) || '').trim();
  const titleTop = titleElement.getTop();
  const left = titleElement.getLeft();
  const width = titleElement.getWidth();
  const titleHeight = estimateCatalogTitleTextHeight_(titleText, width);
  const renderedTitleBottom = titleTop + Math.min(titleElement.getHeight(), titleHeight);
  const top = renderedTitleBottom + 1.8;
  const lineCount = estimateCatalogSpecInfoLineCount_(specInfo, width);
  const height = Math.max(5.6, (lineCount * 5.35));

  const shape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, left, top, width, height);
  shape.getFill().setTransparent();
  shape.getBorder().setTransparent();
  shape.getText().setText(specInfo);
  shape.getText().getTextStyle()
    .setForegroundColor('#000000')
    .setBold(false)
    .setFontFamily('Montserrat')
    .setFontSize(6.8);
  shape.getText().getParagraphStyle()
    .setParagraphAlignment(SlidesApp.ParagraphAlignment.START);

  const desiredContentTop = top + height + 0.8;
  [
    `{{SECTION_${sectionNumber}_PICTURE}}`,
    `{{SECTION_${sectionNumber}_TABLE}}`
  ].forEach(placeholder => {
    const element = findCatalogPlaceholderElement_(slide, placeholder);
    if (!element) return;
    const shiftDelta = Math.max(0, desiredContentTop - element.getTop());
    if (shiftDelta > 0.1) {
      element.setTop(element.getTop() + shiftDelta);
    }
  });

  return {
    left,
    top,
    width,
    height,
    bottom: top + height
  };
}

function estimateCatalogTitleTextHeight_(text, width) {
  const normalizedText = String(text || '').trim();
  if (!normalizedText) return 11.5;

  const lineCount = estimateCatalogWrappedLineCount_(normalizedText, width, 6.35, 1);
  return Math.max(11.5, lineCount * 9.25);
}

function estimateCatalogSpecInfoLineCount_(text, width) {
  const normalizedText = String(text || '').trim();
  if (!normalizedText) return 0;

  return normalizedText
    .split(/\s*\/\s*/)
    .map(part => estimateCatalogWrappedLineCount_(String(part || '').trim(), width, 4.55, 1))
    .reduce((sum, count) => sum + count, 0);
}

function estimateCatalogWrappedLineCount_(text, width, avgCharWidth, minLines) {
  const normalizedText = String(text || '').trim();
  if (!normalizedText) return Math.max(0, Number(minLines) || 0);

  const charsPerLine = Math.max(12, Math.floor((Number(width) || 180) / Math.max(1, Number(avgCharWidth) || 4.5)));
  return Math.max(Math.max(1, Number(minLines) || 1), Math.ceil(normalizedText.length / charsPerLine));
}

function compactCatalogSectionLayout_(slide, sectionNumber, previousSectionBottom) {
  const titleElement = findCatalogPlaceholderElement_(slide, `{{SECTION_${sectionNumber}_TITLE}}`);
  if (!titleElement) return;

  const desiredTitleTop = previousSectionBottom + 11;
  const currentTitleTop = titleElement.getTop();
  const delta = desiredTitleTop - currentTitleTop;

  if (Math.abs(delta) < 1) return;

  [
    `{{SECTION_${sectionNumber}_TITLE}}`,
    `{{SECTION_${sectionNumber}_PICTURE}}`,
    `{{SECTION_${sectionNumber}_TABLE}}`
  ].forEach(placeholder => {
    const element = findCatalogPlaceholderElement_(slide, placeholder);
    if (!element) return;
    element.setTop(element.getTop() + delta);
  });
}

function buildCatalogPageMeta_(meta, page) {
  const pageMeta = Object.assign({}, meta);
  const pageSubtitles = (page.sections || [])
    .map(section => String(section.pageSubtitle || '').trim())
    .filter(Boolean)
    .filter((value, index, values) => values.indexOf(value) === index);
  const pageVariant2Values = (page.sections || [])
    .map(section => normalizeVariant2Key_(section.variant2))
    .filter(Boolean)
    .filter((value, index, values) => values.indexOf(value) === index);

  if (pageSubtitles.length === 1) {
    pageMeta.subtitle = pageSubtitles[0];
  } else if (pageVariant2Values.length === 1) {
    pageMeta.subtitle = buildCatalogVariant2Subtitle_(meta.subtitle, pageVariant2Values[0]);
  }

  return pageMeta;
}

function buildCatalogVariant2Subtitle_(defaultSubtitle, variant2) {
  const normalizedVariant2 = String(variant2 || '').trim();
  if (!normalizedVariant2) return defaultSubtitle;

  if (/threaded fittings/i.test(String(defaultSubtitle || ''))) {
    return `${normalizedVariant2} Threaded Fittings`;
  }

  return normalizedVariant2;
}

function buildCatalogTextPlaceholders_(meta, pageNumber, continuedLabel) {
  return {
    '{{CATALOG_TITLE}}': meta.title,
    '{{CATALOG_SUBTITLE}}': meta.subtitle,
    '{{CATALOG_SUBHEADING}}': meta.subheading,
    '{{KEY_BULLET_1}}': meta.keyBullet1,
    '{{KEY_BULLET_2}}': meta.keyBullet2,
    '{{KEY_BULLET_3}}': meta.keyBullet3,
    '{{KEY_BULLET_4}}': meta.keyBullet4,
    '{{CATALOG_CODE}}': meta.catalogCode,
    '{{VERSION_CODE}}': meta.versionCode,
    '{{EFFECTIVE_DATE}}': meta.effectiveDate,
    '{{PAGE_NUMBER}}': String(pageNumber || ''),
    '{{CONTINUED_LABEL}}': continuedLabel || ''
  };
}

function fillCatalogSlidePlaceholders_(slide, replacements) {
  Object.keys(replacements).forEach(placeholder => {
    replaceCatalogText_(slide, placeholder, replacements[placeholder]);
  });
}

function replaceCatalogText_(slide, placeholder, replacement) {
  slide.replaceAllText(placeholder, replacement || '');
}

function replaceCatalogImagePlaceholder_(slide, placeholder, fileId) {
  const target = findCatalogPlaceholderElement_(slide, placeholder);
  if (!target) return;

  const left = target.getLeft();
  const top = target.getTop();
  const width = target.getWidth();
  const height = target.getHeight();
  target.remove();

  if (!fileId) {
    insertMissingSlidesImageText_(slide, left, top, width, height);
    return;
  }

  try {
    const blob = fetchDriveImageBlob_(fileId, Math.round(width * 3));
    slide.insertImage(blob, left, top, width, height);
  } catch (err) {
    Logger.log(`Image insert failed for ${fileId}: ${err.message}`);
    insertMissingSlidesImageText_(slide, left, top, width, height);
  }
}

function replaceCatalogTablePlaceholder_(slide, placeholder, section) {
  const target = findCatalogPlaceholderElement_(slide, placeholder);
  if (!target) return null;

  const left = target.getLeft();
  const top = target.getTop();
  const width = target.getWidth();
  const height = target.getHeight();
  target.remove();

  return insertCatalogSlidesShapeTable_(slide, section, left, top, width, height);
}

function getCatalogPlaceholderBounds_(slide, placeholder) {
  const target = findCatalogPlaceholderElement_(slide, placeholder);
  if (!target) return null;

  return {
    left: target.getLeft(),
    top: target.getTop(),
    width: target.getWidth(),
    height: target.getHeight(),
    bottom: target.getTop() + target.getHeight()
  };
}

function findCatalogPlaceholderElement_(slide, placeholder) {
  const elements = getCatalogSlidePageElements_(slide);

  for (let i = 0; i < elements.length; i++) {
    const element = elements[i];
    if (element.getPageElementType() !== SlidesApp.PageElementType.SHAPE) continue;

    let text = '';
    try {
      text = element.asShape().getText().asString();
    } catch (err) {
      text = '';
    }

    if (text.indexOf(placeholder) !== -1) return element;
  }

  return null;
}

function getCatalogSlidePageElements_(slide) {
  const elements = [];

  slide.getPageElements().forEach(element => {
    collectCatalogPageElement_(element, elements);
  });

  return elements;
}

function collectCatalogPageElement_(element, elements) {
  elements.push(element);

  if (element.getPageElementType() !== SlidesApp.PageElementType.GROUP) return;

  element.asGroup().getChildren().forEach(child => {
    collectCatalogPageElement_(child, elements);
  });
}

function removeCatalogPlaceholderElement_(slide, placeholder) {
  const target = findCatalogPlaceholderElement_(slide, placeholder);
  if (target) target.remove();
}

function insertMissingSlidesImageText_(slide, left, top, width, height) {
  const shape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, left, top, width, height);
  shape.getText().setText('MISSING IMAGE');
  shape.getText().getTextStyle()
    .setForegroundColor('#C00000')
    .setBold(true)
    .setFontSize(12);
  shape.getText().getParagraphStyle()
    .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
}

function fetchDriveImageBlob_(fileId, width) {
  const thumbnailUrl = `https://drive.google.com/thumbnail?id=${fileId}&sz=w${width || 1200}`;

  const response = UrlFetchApp.fetch(thumbnailUrl, {
    headers: {
      Authorization: `Bearer ${ScriptApp.getOAuthToken()}`
    },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error(`Thumbnail fetch failed with response ${response.getResponseCode()}`);
  }

  return response.getBlob();
}

function insertCatalogSlidesTable_(slide, section, left, top, width, height) {
  const blue = '#274A96';
  const veryLightBlue = '#EEF6FC';
  const white = '#FFFFFF';
  const variants = section.rawVariants || [];
  const forceForgedSteelVariantShapeTable = isForgedSteelCatalogProductGroup_(section.productGroup) && variants.length > 1;
  const tableRows = section.tableRows || [];
  const hideInnerCarton = !!section.hideInnerCarton;
  const hideMasterCase = !!section.hideMasterCase;
  let nextColumnIndex = 0;
  const sizeColumn = nextColumnIndex++;
  const innerCartonColumn = hideInnerCarton ? -1 : nextColumnIndex++;
  const masterCartonColumn = hideMasterCase ? -1 : nextColumnIndex++;
  const variantStartColumn = nextColumnIndex;
  const isSingleVariant = variants.length <= 1;
  const usesStackedVariantHeader = !isSingleVariant && shouldUseStackedCatalogVariantHeader_(variants);
  const columnCount = nextColumnIndex + (isSingleVariant ? 2 : variants.length * 2);
  const headerRowCount = 1;
  const rowCount = headerRowCount + tableRows.length;
  const isWideTable = columnCount > 8;
  const isCompactTable = !!section.compactTable;
  if (isCompactTable || section.isNippleTable || section.forceShapeTable || forceForgedSteelVariantShapeTable) {
    return insertCatalogSlidesShapeTable_(slide, section, left, top, width, height);
  }

  const rowHeight = isWideTable ? 7.2 : (isCompactTable ? 7.4 : (tableRows.length > 10 ? 8.0 : 9.0));
  const bodyFontSize = isWideTable ? 5.1 : (isCompactTable ? 5.5 : (tableRows.length > 10 ? 5.9 : 6.8));
  const sizeFontSize = bodyFontSize;
  const headerFontSize = isWideTable ? 4.8 : (isCompactTable ? 5.5 : 6.2);
  const variantHeaderFontSize = isWideTable ? 4.5 : (isCompactTable ? 5.3 : 5.8);
  const sizeHeader = section.sizeHeader || 'Size';
  const compactHeight = Math.min(height, Math.max(56, rowCount * rowHeight));
  const variantOverlayHeight = usesStackedVariantHeader ? 11 : 0;
  const tableBounds = getCatalogSlidesTableBounds_(left, top + variantOverlayHeight, width, compactHeight, {
    compactTable: isCompactTable
  });
  const table = slide.insertTable(rowCount, columnCount, tableBounds.left, tableBounds.top, tableBounds.width, tableBounds.height);
  setCatalogSlidesTableRowHeights_(table, rowCount, rowHeight);

  if (isSingleVariant) {
    const variantColors = getVariantHeaderColors_(variants[0] || '', {
      productGroup: section.productGroup,
      variantIndex: 0,
      variantCount: variants.length,
      isSingleVariant: true
    });
    setSlidesTableCell_(table, 0, sizeColumn, sizeHeader, blue, white, true, headerFontSize);
    if (!hideInnerCarton) {
      setSlidesTableCell_(table, 0, innerCartonColumn, 'Inner Carton', blue, white, true, headerFontSize);
    }
    if (!hideMasterCase) {
      setSlidesTableCell_(table, 0, masterCartonColumn, 'Master Carton', blue, white, true, headerFontSize);
    }
    setSlidesTableCell_(table, 0, variantStartColumn, 'Item Number', variantColors.background, variantColors.font, true, headerFontSize);
    setSlidesTableCell_(table, 0, variantStartColumn + 1, 'List Price', variantColors.background, variantColors.font, true, headerFontSize);
  } else if (usesStackedVariantHeader) {
    setSlidesTableCell_(table, 0, sizeColumn, sizeHeader, blue, white, true, headerFontSize);
    if (!hideInnerCarton) {
      setSlidesTableCell_(table, 0, innerCartonColumn, 'Inner Carton', blue, white, true, headerFontSize);
    }
    if (!hideMasterCase) {
      setSlidesTableCell_(table, 0, masterCartonColumn, 'Master Carton', blue, white, true, headerFontSize);
    }

    variants.forEach((variant, index) => {
      const startColumn = variantStartColumn + index * 2;
      const headerColors = getVariantHeaderColors_(variant, {
        productGroup: section.productGroup,
        variantIndex: index,
        variantCount: variants.length
      });
      insertCatalogVariantHeaderOverlay_(
        slide,
        getDisplayVariantName_(variant),
        left,
        top,
        width,
        variantOverlayHeight,
        1,
        columnCount,
        startColumn,
        headerColors.background,
        headerColors.font
      );
      setSlidesTableCell_(table, 0, startColumn, 'Item Number', blue, white, true, variantHeaderFontSize);
      setSlidesTableCell_(table, 0, startColumn + 1, 'List Price', blue, white, true, variantHeaderFontSize);
    });

  } else {
    setSlidesTableCell_(table, 0, sizeColumn, sizeHeader, blue, white, true, headerFontSize);
    if (!hideInnerCarton) {
      setSlidesTableCell_(table, 0, innerCartonColumn, 'Inner Carton', blue, white, true, headerFontSize);
    }
    if (!hideMasterCase) {
      setSlidesTableCell_(table, 0, masterCartonColumn, 'Master Carton', blue, white, true, headerFontSize);
    }

    variants.forEach((variant, index) => {
      const startColumn = variantStartColumn + index * 2;
      const headerColors = getVariantHeaderColors_(variant, {
        productGroup: section.productGroup,
        variantIndex: index,
        variantCount: variants.length
      });
      const variantName = getDisplayVariantName_(variant);
      setSlidesTableCell_(table, 0, startColumn, `${variantName} Item Number`, headerColors.background, headerColors.font, true, variantHeaderFontSize);
      setSlidesTableCell_(table, 0, startColumn + 1, `${variantName} List Price`, headerColors.background, headerColors.font, true, variantHeaderFontSize);
    });
  }

  tableRows.forEach((row, rowIndex) => {
    const tableRowIndex = rowIndex + headerRowCount;
    const background = rowIndex % 2 === 0 ? white : veryLightBlue;

    setSlidesTableCell_(table, tableRowIndex, sizeColumn, formatCatalogVisibleCellValue_(row.size), background, '#000000', false, sizeFontSize);
    if (!hideInnerCarton) {
      setSlidesTableCell_(table, tableRowIndex, innerCartonColumn, formatCatalogVisibleCellValue_(row.innerCarton), background, '#000000', false, bodyFontSize);
    }
    if (!hideMasterCase) {
      setSlidesTableCell_(table, tableRowIndex, masterCartonColumn, formatCatalogVisibleCellValue_(row.masterCase), background, '#000000', false, bodyFontSize);
    }

    if (isSingleVariant) {
      const item = row.variants[variants[0]] || {};
      setSlidesTableCell_(table, tableRowIndex, variantStartColumn, formatCatalogVisibleCellValue_(item.itemNumber), background, '#000000', false, bodyFontSize);
      setSlidesTableCell_(table, tableRowIndex, variantStartColumn + 1, formatCatalogCurrency_(item.price), background, '#000000', false, bodyFontSize);
    } else {
      variants.forEach((variant, variantIndex) => {
        const item = row.variants[variant] || {};
        const startColumn = variantStartColumn + variantIndex * 2;
        setSlidesTableCell_(table, tableRowIndex, startColumn, formatCatalogVisibleCellValue_(item.itemNumber), background, '#000000', false, bodyFontSize);
        setSlidesTableCell_(table, tableRowIndex, startColumn + 1, formatCatalogCurrency_(item.price), background, '#000000', false, bodyFontSize);
      });
    }
  });

  styleCatalogSlidesTable_(table, rowCount, columnCount);
  const layoutMultiplier = section.forceSingleSectionPage ? 12 : 19;
  const layoutHeight = Math.max(compactHeight, rowCount * layoutMultiplier);
  return {
    top,
    height: layoutHeight + variantOverlayHeight,
    bottom: top + layoutHeight + variantOverlayHeight
  };
}

function insertCatalogSlidesShapeTable_(slide, section, left, top, width, height) {
  const blue = '#274A96';
  const veryLightBlue = '#EEF6FC';
  const white = '#FFFFFF';
  const variants = section.rawVariants || [];
  const rows = section.tableRows || [];
  const normalizedProductGroup = String(section.productGroup || '').toLowerCase();
  const isValveTable = normalizedProductGroup.indexOf('valve') !== -1;
  const isDualVariantValveTable = isValveTable && variants.length > 1;
  const isForgedSteelFittings = normalizedProductGroup.indexOf('forged steel fittings') !== -1 &&
    normalizedProductGroup.indexOf('stainless') === -1;
  const isButtWeldCatalog = isButtWeldCatalogProductGroup_(section.productGroup);
  const variantSpecificInnerCarton = !!section.variantSpecificInnerCarton;
  const variantSpecificMasterCase = isForgedSteelFittings && variants.length > 1
    ? true
    : !!section.variantSpecificMasterCase;
  const isSingleVariant = variants.length <= 1;
  const isSingleVariantForgedSteel = isForgedSteelFittings && isSingleVariant;
  const hideInnerCarton = !!section.hideInnerCarton;
  const hideMasterCase = variantSpecificMasterCase ? false : !!section.hideMasterCase;
  const isNippleTable = !!section.isNippleTable;
  const hasWeightVariantHeaders = variants.some(variant => /\b(std|standard weight|lw|light weight|xh|extra heavy)\b/i.test(String(variant || '')));
  const isWeightPairedButtWeldTable = isButtWeldCatalog && hasWeightVariantHeaders && variants.length > 1;
  const isBlackGalvanizedPairedTable = variants.length === 2 &&
    variants.some(variant => String(variant || '').toLowerCase().indexOf('black') !== -1) &&
    variants.some(variant => String(variant || '').toLowerCase().indexOf('galvanized') !== -1);
  const sizeHeader = section.sizeHeader || 'Size';
  const cartonWidth = isNippleTable ? 44 : (hasWeightVariantHeaders ? 44 : (isSingleVariantForgedSteel ? 54 : (isForgedSteelFittings ? 30 : 50)));
  const priceWidth = isNippleTable ? 34 : (hasWeightVariantHeaders ? 54 : (isDualVariantValveTable ? 42 : (isBlackGalvanizedPairedTable ? 42 : (isForgedSteelFittings ? 24 : 50))));
  const sizeWidthBoost = hideMasterCase ? 14 : 0;
  const itemWidthBoost = (hideMasterCase ? 12 : 0) +
    (hasWeightVariantHeaders ? 8 : 0) +
    (isNippleTable ? 20 : 0) +
    (isDualVariantValveTable ? 24 : 0) +
    (isBlackGalvanizedPairedTable ? 18 : 0) +
    (isSingleVariantForgedSteel ? 68 : (isForgedSteelFittings ? 52 : 0));
  const priceWidthBoost = (hideMasterCase ? 8 : 0) + (hasWeightVariantHeaders ? 6 : 0) - (isBlackGalvanizedPairedTable ? 4 : 0);
  const columns = [{
    header: sizeHeader,
    headerBackground: blue,
    headerFont: white,
    value: row => formatCatalogVisibleCellValue_(row.size),
    targetWidth: Math.min(
      isForgedSteelFittings ? 138 : 126,
      Math.max(
        isForgedSteelFittings ? 98 : 90,
        58 + getLongestCatalogTextLength_([sizeHeader].concat(rows.map(row => row.size))) * (isForgedSteelFittings ? 3.4 : 3) + sizeWidthBoost
      )
    ),
    flexibleWeight: 2
  }];

  if (!hideInnerCarton && !variantSpecificInnerCarton) {
    columns.push({
      header: 'Inner Carton',
      headerBackground: blue,
      headerFont: white,
      value: row => formatCatalogVisibleCellValue_(row.innerCarton),
      targetWidth: cartonWidth,
      flexibleWeight: 0
    });
  }

  if (!hideMasterCase && !variantSpecificMasterCase) {
    columns.push({
      header: 'Master Carton',
      headerBackground: blue,
      headerFont: white,
      value: row => formatCatalogVisibleCellValue_(row.masterCase),
      targetWidth: cartonWidth,
      flexibleWeight: 0
    });
  }

  (variants.length ? variants : ['']).forEach((variant, variantIndex) => {
    const headerColors = getVariantHeaderColors_(variant, {
      productGroup: section.productGroup,
      variantIndex,
      variantCount: variants.length,
      isSingleVariant
    });
    const variantName = getDisplayVariantName_(variant);
    const itemHeader = isSingleVariant ? 'Item Number' : `${variantName} Item Number`;
    const priceHeader = isSingleVariant ? 'List Price' : `${variantName} List Price`;
    const itemLength = getLongestCatalogTextLength_(rows.map(row => {
      const item = row.variants && row.variants[variant];
      return item && item.itemNumber;
    }));
    const itemHeaderLength = getLongestCatalogTextLength_([itemHeader]);
    const priceHeaderLength = getLongestCatalogTextLength_([priceHeader]);

    columns.push({
      header: itemHeader,
      headerBackground: headerColors.background,
      headerFont: headerColors.font,
      value: row => {
        const item = row.variants && row.variants[variant];
        return formatCatalogVisibleCellValue_(item && item.itemNumber);
      },
      targetWidth: Math.min(
        isWeightPairedButtWeldTable ? 118 : (hasWeightVariantHeaders ? 132 : (isNippleTable ? 134 : (isSingleVariantForgedSteel ? 188 : (isDualVariantValveTable ? 136 : (isBlackGalvanizedPairedTable ? 126 : (isForgedSteelFittings ? 170 : 110)))))),
        Math.max(
          isWeightPairedButtWeldTable ? 84 : (hasWeightVariantHeaders ? 94 : (isNippleTable ? 98 : (isSingleVariantForgedSteel ? 128 : (isDualVariantValveTable ? 104 : (isBlackGalvanizedPairedTable ? 92 : (isForgedSteelFittings ? 118 : 72)))))),
          42 + Math.max(itemLength, itemHeaderLength) * (isForgedSteelFittings ? 3.0 : 2) + itemWidthBoost
        )
      ),
      flexibleWeight: 1
    });
    if (variantSpecificInnerCarton) {
      columns.push({
        header: isSingleVariant ? 'Inner Carton' : 'IC',
        headerBackground: headerColors.background,
        headerFont: headerColors.font,
        value: row => {
          const item = row.variants && row.variants[variant];
          return formatCatalogVisibleCellValue_(item && item.innerCarton);
        },
        targetWidth: cartonWidth,
        flexibleWeight: 0
      });
    }
    if (variantSpecificMasterCase) {
      columns.push({
        header: isSingleVariant ? 'Master Carton' : 'MC',
        headerBackground: headerColors.background,
        headerFont: headerColors.font,
        value: row => {
          const item = row.variants && row.variants[variant];
          return formatCatalogVisibleCellValue_(item && item.masterCase);
        },
        targetWidth: cartonWidth,
        flexibleWeight: 0
      });
    }
    columns.push({
      header: priceHeader,
      headerBackground: headerColors.background,
      headerFont: headerColors.font,
      value: row => {
        const item = row.variants && row.variants[variant];
        return formatCatalogCurrency_(item && item.price);
      },
      targetWidth: isDualVariantValveTable
        ? Math.min(68, Math.max(priceWidth + priceWidthBoost, 30 + priceHeaderLength * 1.65))
        : isWeightPairedButtWeldTable
        ? Math.min(76, Math.max(priceWidth + priceWidthBoost + 10, 36 + priceHeaderLength * 1.8))
        : isBlackGalvanizedPairedTable
        ? Math.min(48, Math.max(priceWidth + priceWidthBoost, 28 + priceHeaderLength * 1.5))
        : hasWeightVariantHeaders
        ? Math.min(54, Math.max(priceWidth + priceWidthBoost, 28 + priceHeaderLength * 1.35))
        : isForgedSteelFittings
        ? Math.min(58, Math.max(priceWidth + priceWidthBoost, 28 + priceHeaderLength * 1.5))
        : isNippleTable
        ? Math.max(priceWidth + priceWidthBoost, 26 + priceHeaderLength * 1.6)
        : Math.max(priceWidth + priceWidthBoost, 34 + priceHeaderLength * 2.4),
      flexibleWeight: 0
    });
  });

  const tableWidth = getCatalogShapeTableWidth_(columns, width, variants.length);
  const columnWidths = fitCatalogShapeColumnWidths_(columns, tableWidth);
  const headerHeight = 20;
  const availableBodyHeight = Math.max(0, height - headerHeight);
  const bodyRowHeight = rows.length ? Math.min(12.5, Math.max(10.5, availableBodyHeight / rows.length)) : 0;
  const tableHeight = headerHeight + bodyRowHeight * rows.length;

  let columnLeft = left;
  columns.forEach((column, columnIndex) => {
    insertCatalogSlidesShapeCell_(
      slide,
      columnLeft,
      top,
      columnWidths[columnIndex],
      headerHeight,
      column.header,
      column.headerBackground,
      column.headerFont,
      true,
      5.3
    );
    columnLeft += columnWidths[columnIndex];
  });

  rows.forEach((row, rowIndex) => {
    const rowTop = top + headerHeight + rowIndex * bodyRowHeight;
    const background = rowIndex % 2 === 0 ? white : veryLightBlue;
    let cellLeft = left;

    columns.forEach((column, columnIndex) => {
      insertCatalogSlidesShapeCell_(
        slide,
        cellLeft,
        rowTop,
        columnWidths[columnIndex],
        bodyRowHeight,
        column.value(row),
        background,
        '#000000',
        false,
        5.5
      );
      cellLeft += columnWidths[columnIndex];
    });
  });

  return {
    top,
    height: tableHeight,
    bottom: top + tableHeight
  };
}

function getCatalogShapeTableWidth_(columns, availableWidth, variantCount) {
  if (variantCount > 1) return availableWidth;

  const naturalWidth = columns.reduce((sum, column) => sum + column.targetWidth, 0);
  const singleVariantWidth = Math.min(availableWidth * 0.92, naturalWidth + 48);
  return Math.min(availableWidth, Math.max(naturalWidth, singleVariantWidth));
}

function fitCatalogShapeColumnWidths_(columns, totalWidth) {
  const widths = columns.map(column => column.targetWidth);
  const flexibleIndices = columns
    .map((column, index) => column.flexibleWeight ? index : -1)
    .filter(index => index >= 0);
  const currentTotal = widths.reduce((sum, columnWidth) => sum + columnWidth, 0);
  let difference = totalWidth - currentTotal;

  if (Math.abs(difference) < 0.1 || !flexibleIndices.length) return widths;

  if (difference > 0) {
    const weightTotal = flexibleIndices.reduce((sum, index) => sum + columns[index].flexibleWeight, 0);
    flexibleIndices.forEach(index => {
      widths[index] += difference * columns[index].flexibleWeight / weightTotal;
    });
    return widths;
  }

  const minimumWidths = columns.map((column, index) => index === 0 ? 76 : (column.flexibleWeight ? 58 : column.targetWidth));
  let remainingReduction = -difference;

  while (remainingReduction > 0.1) {
    const reducibleIndices = flexibleIndices.filter(index => widths[index] > minimumWidths[index] + 0.1);
    if (!reducibleIndices.length) break;

    const reductionPerColumn = remainingReduction / reducibleIndices.length;
    reducibleIndices.forEach(index => {
      const reduction = Math.min(reductionPerColumn, widths[index] - minimumWidths[index]);
      widths[index] -= reduction;
      remainingReduction -= reduction;
    });
  }

  if (remainingReduction > 0.1) {
    const scale = totalWidth / widths.reduce((sum, columnWidth) => sum + columnWidth, 0);
    return widths.map(columnWidth => columnWidth * scale);
  }

  return widths;
}

function insertCatalogSlidesShapeCell_(slide, left, top, width, height, value, background, fontColor, bold, fontSize) {
  const shape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, left, top, width, height);
  shape.getFill().setSolidFill(background);
  shape.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);

  const border = shape.getBorder();
  border.setWeight(0.5);
  border.getLineFill().setSolidFill('#8A8A8A');

  const textValue = String(value === null || value === undefined ? '' : value);
  if (textValue) {
    const textRange = shape.getText();
    textRange.setText(textValue);
    textRange.getTextStyle()
      .setForegroundColor(fontColor)
      .setBold(!!bold)
      .setFontFamily('Arial')
      .setFontSize(fontSize);
    textRange.getParagraphStyle()
      .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  }

  const autofit = shape.getAutofit();
  if (autofit) autofit.disableAutofit();
  return shape;
}

function setSlidesTableCell_(table, rowIndex, columnIndex, value, background, fontColor, bold, fontSize) {
  const cell = table.getCell(rowIndex, columnIndex);
  const textValue = String(value === null || value === undefined ? '' : value);

  cell.getFill().setSolidFill(background);

  if (!textValue) return;

  const textRange = cell.getText();
  textRange.setText(textValue);
  textRange.getTextStyle()
    .setForegroundColor(fontColor)
    .setBold(!!bold)
    .setFontSize(fontSize || 6);
  textRange.getParagraphStyle()
    .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
}

function getCatalogSlidesTableBounds_(left, top, width, height, options) {
  return { left, top, width, height };
}

function setCatalogSlidesTableRowHeights_(table, rowCount, rowHeight) {
  try {
    if (typeof table.getRow !== 'function') return;

    for (let row = 0; row < rowCount; row++) {
      const tableRow = table.getRow(row);
      if (!tableRow) continue;

      if (typeof tableRow.setMinimumHeight === 'function') {
        tableRow.setMinimumHeight(rowHeight);
      } else if (typeof tableRow.setHeight === 'function') {
        tableRow.setHeight(rowHeight);
      }
    }
  } catch (err) {
    Logger.log(`Table row height lock skipped: ${err.message}`);
  }
}

function getLongestCatalogTextLength_(values) {
  return (values || []).reduce((longest, value) => Math.max(longest, String(value || '').length), 0);
}

function shouldUseStackedCatalogVariantHeader_(variants) {
  return false;
}

function mergeCatalogSlidesTableCells_(table, rowIndex, startColumnIndex, columnCount) {
  if (columnCount < 2) return;

  try {
    const firstCell = table.getCell(rowIndex, startColumnIndex);
    const secondCell = table.getCell(rowIndex, startColumnIndex + 1);

    if (typeof firstCell.merge === 'function') {
      firstCell.merge();
      return;
    }

    if (typeof secondCell.merge === 'function') {
      secondCell.merge();
    }
  } catch (err) {
    Logger.log(`Variant header merge skipped: ${err.message}`);
  }
}

function insertCatalogVariantHeaderOverlay_(slide, text, tableLeft, tableTop, tableWidth, tableHeight, rowCount, columnCount, startColumn, background, fontColor) {
  const columnWidth = tableWidth / columnCount;
  const rowHeight = tableHeight / rowCount;
  const left = tableLeft + startColumn * columnWidth;
  const top = tableTop;
  const width = columnWidth * 2;
  const height = rowHeight;

  const shape = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, left, top, width, height);
  shape.getFill().setSolidFill(background);
  if (typeof shape.getLine === 'function') {
    shape.getLine().getFill().setSolidFill(background);
  }
  shape.getText().setText(text);
  shape.getText().getTextStyle()
    .setForegroundColor(fontColor)
    .setBold(true)
    .setFontSize(6.2);
  shape.getText().getParagraphStyle()
    .setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
}

function styleCatalogSlidesTable_(table, rowCount, columnCount) {
  for (let row = 0; row < rowCount; row++) {
    for (let column = 0; column < columnCount; column++) {
      const cell = table.getCell(row, column);
      styleCatalogSlidesCellBorder_(cell, 'getBorderTop');
      styleCatalogSlidesCellBorder_(cell, 'getBorderBottom');
      styleCatalogSlidesCellBorder_(cell, 'getBorderLeft');
      styleCatalogSlidesCellBorder_(cell, 'getBorderRight');
    }
  }
}

function styleCatalogSlidesCellBorder_(cell, borderGetterName) {
  if (typeof cell[borderGetterName] !== 'function') return;

  const border = cell[borderGetterName]();
  if (!border) return;

  if (typeof border.setWeight === 'function') border.setWeight(0.5);
  if (typeof border.getLineFill === 'function') {
    border.getLineFill().setSolidFill('#000000');
  }
}

function formatCatalogCurrency_(value) {
  const normalized = normalizeCatalogPriceValue_(value);
  if (normalized === '' || normalized === null) return '-';
  if (normalized === 'POA') return 'POA';
  const number = Number(normalized);
  if (isNaN(number)) return '-';
  if (!number) return value === 0 ? '$0.00' : '-';
  return `$${formatCatalogNumberWithCommas_(number, 2)}`;
}

function formatCatalogNumberWithCommas_(value, fixedDecimals) {
  const number = Number(value);
  if (isNaN(number)) return String(value || '').trim();

  const useFixedDecimals = fixedDecimals !== undefined && fixedDecimals !== null;
  const numberText = useFixedDecimals
    ? number.toFixed(Math.max(0, Number(fixedDecimals) || 0))
    : String(number);
  const parts = numberText.split('.');
  parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return parts.join('.');
}

function normalizeCatalogPriceValue_(value) {
  if (value === null || value === undefined) return '';
  if (typeof value === 'number') {
    return isNaN(value) ? null : value;
  }

  const trimmed = String(value).trim();
  if (!trimmed) return '';
  if (trimmed.toUpperCase() === 'POA') return 'POA';

  const number = Number(trimmed);
  if (!isNaN(number)) return number;
  return null;
}

function calculateCatalogPriceValue_(value, appliedIncrease) {
  const normalized = normalizeCatalogPriceValue_(value);
  if (normalized === '' || normalized === 'POA' || normalized === null) return normalized === null ? '' : normalized;
  return normalized * (1 + (Number(appliedIncrease) || 0));
}

function catalogPriceValuesEqual_(leftValue, rightValue) {
  const normalizedLeft = normalizeCatalogPriceValue_(leftValue);
  const normalizedRight = normalizeCatalogPriceValue_(rightValue);
  return normalizedLeft === normalizedRight;
}

function getCatalogRowPriceValue_(row, col) {
  const calculated = normalizeCatalogPriceValue_(row[col.Calculated_List_Price]);
  if (calculated !== '' && calculated !== null) return calculated;
  const listPrice = normalizeCatalogPriceValue_(row[col.List_Price]);
  if (listPrice !== null) return listPrice;
  return '';
}

function formatPriceFileCellValue_(value) {
  const normalized = normalizeCatalogPriceValue_(value);
  if (normalized === '' || normalized === null) return '-';
  return normalized;
}

function buildPriceFileNetFormula_(rowNumber, lookupRange) {
  return `=IF(UPPER(TO_TEXT(F${rowNumber}))="POA","POA",IF(OR(F${rowNumber}="",F${rowNumber}="-"),"-",IFERROR(F${rowNumber}*VLOOKUP(C${rowNumber},${lookupRange},2,FALSE),"-")))`;
}

function formatCatalogVisibleCellValue_(value) {
  const textValue = String(value === null || value === undefined ? '' : value).trim();
  if (!textValue) return '-';
  if (typeof value === 'number' && !isNaN(value)) return formatCatalogNumberWithCommas_(value);
  if (/^-?(?:0|[1-9]\d*)(?:\.\d+)?$/.test(textValue) && !/^0\d+/.test(textValue)) {
    return formatCatalogNumberWithCommas_(textValue);
  }
  return textValue;
}

function formatCatalogEffectiveDate_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(value, getCatalogTimeZone_(), 'MMMM yyyy');
  }
  return String(value).trim();
}

function buildPriceFileSheet_(sheet, rows, col, meta) {
  const uniquePlcs = [...new Set(
    rows.map(row => formatPriceFilePlc_(row[col.PLC])).filter(Boolean)
  )].sort();
  const hideInnerCartonColumn = rows.every(row => !String(row[col.Inner_Carton] || '').trim());
  const hideMasterCaseColumn = rows.every(row => !String(row[col.Master_Case] || '').trim());

  sheet.clear();
  sheet.setHiddenGridlines(true);

  // Logo
  if (meta.logoFileId) {
    insertResizedDriveImage_(sheet, meta.logoFileId, 1, 1, 210, 75);
  }

  sheet.setRowHeight(1, 48);
  sheet.setRowHeight(2, 28);
  sheet.setRowHeight(3, 32);
  sheet.setColumnWidth(1, 125);
  sheet.setColumnWidth(2, 145);

  // Header
  sheet.getRange('D1:G1').merge();
  sheet.getRange('D1').setValue(meta.title);

  sheet.getRange('D2:G2').merge();
  sheet.getRange('D2').setValue(
    meta.subtitle ? `${meta.subtitle} (${meta.versionCode})` : meta.versionCode
  );

  sheet.getRange('D3:G3').merge();
  sheet.getRange('D3').setValue(
    'Enter your multiplier next to each PLC below. Net Ea will calculate automatically from the List Price.'
  );

  // Static multiplier area
  sheet.getRange(6, 1, 1, 2).setValues([['PLC', 'Multiplier']]);

  if (uniquePlcs.length) {
    sheet.getRange(7, 1, uniquePlcs.length, 1).setNumberFormat('@');
    sheet.getRange(7, 1, uniquePlcs.length, 2)
      .setValues(uniquePlcs.map(plc => [plc, '']));
    sheet.getRange(7, 2, uniquePlcs.length, 1).setNumberFormat('0.0000');
  }

  const firstTableHeaderRow = 7 + uniquePlcs.length + 2;
  let currentRow = firstTableHeaderRow;

  const header = [[
    'MSI Item #',
    'General Description',
    'PLC',
    'Inner Carton Qty',
    'Master Case Qty',
    'List Price',
    'Net Ea'
  ]];

  const lookupRange = `$A$7:$B$${6 + uniquePlcs.length}`;
  const outputValues = [];
  const tableHeaderRows = [];
  const groupDividerRows = [];
  let previousGroupKey = '';

  tableHeaderRows.push(currentRow);
  outputValues.push(header[0]);
  currentRow++;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const group = buildPriceFileGroup_(row, col);
    if (group.key !== previousGroupKey) {
      groupDividerRows.push(currentRow);
      outputValues.push([group.label, '', '', '', '', '', '']);
      currentRow++;
      previousGroupKey = group.key;
    }

    const calculatedPrice = getCatalogRowPriceValue_(row, col);

    outputValues.push([
      row[col.Item_Number],
      buildCatalogDescription_(row, col),
      formatPriceFilePlc_(row[col.PLC]),
      row[col.Inner_Carton],
      row[col.Master_Case],
      formatPriceFileCellValue_(calculatedPrice),
      buildPriceFileNetFormula_(currentRow, lookupRange)
    ]);

    currentRow++;
  }

  if (outputValues.length) {
    sheet.getRange(firstTableHeaderRow, 3, outputValues.length, 1).setNumberFormat('@');
    sheet.getRange(firstTableHeaderRow, 1, outputValues.length, 7).setValues(outputValues);
  }
  if (hideInnerCartonColumn) sheet.hideColumns(4);
  if (hideMasterCaseColumn) sheet.hideColumns(5);
  tableHeaderRows.forEach(rowNumber => formatPriceFileTableHeader_(sheet, rowNumber));

  applyPriceFileFormatting_(
    sheet,
    firstTableHeaderRow,
    currentRow - 1,
    uniquePlcs.length,
    tableHeaderRows,
    groupDividerRows
  );
  groupDividerRows.forEach(rowNumber => formatPriceFileGroupDivider_(sheet, rowNumber));
}

function formatPriceFilePlc_(value) {
  const trimmed = String(value === null || value === undefined ? '' : value).trim();
  if (!trimmed) return '';
  if (/^\d+$/.test(trimmed)) return trimmed.padStart(3, '0');
  return trimmed;
}

function insertResizedDriveImage_(sheet, fileId, column, row, width, height, offsetX, offsetY, fallbackText) {
  try {
    const thumbnailUrl = `https://drive.google.com/thumbnail?id=${fileId}&sz=w${width * 3}`;

    const response = UrlFetchApp.fetch(thumbnailUrl, {
      headers: {
        Authorization: `Bearer ${ScriptApp.getOAuthToken()}`
      },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      throw new Error(`Thumbnail fetch failed with response ${response.getResponseCode()}`);
    }

    const blob = response.getBlob().setName('MSI_Logo.png');

    const logo = sheet.insertImage(blob, column, row, offsetX || 0, offsetY || 0);
    logo.setWidth(width);
    logo.setHeight(height);

  } catch (err) {
    Logger.log(`Logo insert failed: ${err.message}`);

    if (fallbackText === 'MISSING IMAGE') {
      setMissingImagePlaceholder_(sheet, row, column);
      return;
    }

    // Fallback: leave a clean text placeholder instead of failing generation.
    sheet.getRange(row, column).setValue(fallbackText || 'MSI');
    sheet.getRange(row, column)
      .setFontWeight('bold')
      .setFontSize(20)
      .setFontColor('#FFFFFF');
  }
}

function formatPriceFileTableHeader_(sheet, rowNumber) {
  const blue = '#1F4E78';
  const white = '#FFFFFF';

  sheet.getRange(rowNumber, 1, 1, 7)
    .setFontWeight('bold')
    .setFontColor(white)
    .setBackground(blue)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
}

function formatPriceFileGroupDivider_(sheet, rowNumber) {
  const range = sheet.getRange(rowNumber, 1, 1, 7);
  range.merge();
  range
    .setBackground('#D9EAF7')
    .setFontColor('#1F4E78')
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, '#1F4E78', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.setRowHeight(rowNumber, 18);
}

function buildPriceFileNippleGroup_(row, col) {
  const scheduleGroup = getCatalogNippleScheduleGroup_(row, col);
  const normalizedSchedule = String(scheduleGroup || '').toLowerCase();
  const scheduleLabel = normalizedSchedule === 's40'
    ? 'Schedule 40'
    : (normalizedSchedule === 's80' ? 'Schedule 80' : String(scheduleGroup || '').trim());
  const fittingType = String(row[col.Fitting_Type] || '').trim();
  const nominalSize = String(row[col.Size_1] || '').trim();
  const productGroup = String(row[col.Product_Group] || '').trim();
  const variant = buildCatalogNippleVariantKey_(row, col);
  let materialLabel = '';

  if (/stainless steel nipple/i.test(productGroup) && variant) {
    materialLabel = `${variant} Stainless`;
  } else if (variant && normalizeMatchValue_(variant) !== normalizeMatchValue_(scheduleGroup)) {
    materialLabel = variant;
  }

  return buildPriceFileSectionGroup_(
    [scheduleLabel, fittingType, nominalSize, materialLabel],
    'Nipples'
  );
}

function buildPriceFileValveGroup_(row, col) {
  const fittingType = String(row[col.Fitting_Type] || '').trim();
  const materialFamily = getValveMaterialFamily_(row, col);
  return buildPriceFileSectionGroup_([fittingType, materialFamily], 'Valves');
}

function buildPriceFileGroup_(row, col) {
  const productGroup = String(row[col.Product_Group] || '').trim();
  if (isNippleCatalogProductGroup_(productGroup)) {
    return buildPriceFileNippleGroup_(row, col);
  }
  if (isValveCatalogProductGroup_(productGroup)) {
    return buildPriceFileValveGroup_(row, col);
  }
  if (isForgedStainlessCatalogProductGroup_(productGroup)) {
    const connectionType = normalizeVariant2Key_(getOptionalCellValue_(row, col, 'Variant_2'));
    const fittingType = String(row[col.Fitting_Type] || '').trim();
    return buildPriceFileSectionGroup_([connectionType, fittingType], productGroup || 'Products');
  }

  if (isForgedSteelCatalogProductGroup_(productGroup)) {
    const pressureClass = getForgedSteelPressureClassFromRow_(row, col);
    const connectionType = getCatalogConnectionDisplayLabel_(row[col.Variant]);
    const fittingType = String(row[col.Fitting_Type] || '').trim();
    return buildPriceFileSectionGroup_([pressureClass, connectionType, fittingType], productGroup || 'Products');
  }

  const rawVariant = normalizeVariantKey_(row[col.Variant]);
  const finish = rawVariant ? getDisplayVariantName_(rawVariant) : '';
  const classification = normalizeVariant2Key_(getOptionalCellValue_(row, col, 'Variant_2'));
  const fittingType = String(row[col.Fitting_Type] || '').trim();
  return buildPriceFileSectionGroup_([finish, classification, fittingType], productGroup || 'Products');
}

function buildPriceFileSectionGroup_(parts, fallbackLabel) {
  const cleanParts = (parts || [])
    .map(value => String(value || '').trim())
    .filter(Boolean);

  if (!cleanParts.length) {
    cleanParts.push(String(fallbackLabel || 'Products').trim());
  }

  return {
    key: cleanParts.map(normalizeMatchValue_).join('|'),
    label: cleanParts.join(' | ')
  };
}

function buildCatalogDescription_(row, col) {
  const sizes = getCatalogRowSizeParts_(row, col).join(' x ');

  const productGroup = String(row[col.Product_Group] || '').trim();
  const fittingType = String(row[col.Fitting_Type] || '').trim();

  if (isNippleCatalogProductGroup_(productGroup)) {
    const materialGroup = productGroup.replace(/\s+nipples?\s*$/i, '').trim();
    return [
      sizes,
      row[col.Variant],
      getOptionalCellValue_(row, col, 'Variant_2'),
      materialGroup,
      fittingType
    ].filter(Boolean).join(' ');
  }

  if (isValveCatalogProductGroup_(productGroup)) {
    return [
      sizes,
      getValveMaterialDetail_(row, col),
      productGroup,
      fittingType
    ].filter(Boolean).join(' ');
  }
  if (isForgedStainlessCatalogProductGroup_(productGroup)) {
    return [
      sizes,
      getDisplayVariantName_(row[col.Variant]),
      normalizeVariant2Key_(getOptionalCellValue_(row, col, 'Variant_2')),
      productGroup,
      fittingType
    ].filter(Boolean).join(' ');
  }

  if (isForgedSteelCatalogProductGroup_(productGroup)) {
    return [
      sizes,
      getForgedSteelPressureClassFromRow_(row, col),
      getCatalogConnectionDisplayLabel_(row[col.Variant]),
      productGroup,
      fittingType
    ].filter(Boolean).join(' ');
  }

  return [
    sizes,
    row[col.Variant],
    getOptionalCellValue_(row, col, 'Variant_2'),
    productGroup,
    fittingType
  ].filter(Boolean).join(' ');
}

function getFirstCatalogImageFileId_(rows, col) {
  if (col.Image_File_ID === undefined) return '';

  const rowWithImage = rows.find(row => String(row[col.Image_File_ID] || '').trim());
  return rowWithImage ? String(rowWithImage[col.Image_File_ID]).trim() : '';
}

function isValveCatalogProductGroup_(productGroup) {
  return String(productGroup || '').toLowerCase().indexOf('valve') !== -1;
}

function isForgedSteelCatalogProductGroup_(productGroup) {
  const normalized = String(productGroup || '').toLowerCase();
  return normalized.indexOf('forged steel fittings') !== -1 &&
    normalized.indexOf('stainless') === -1;
}

function getForgedSteelPressureClassFromRow_(row, col) {
  const itemNumber = String(row[col.Item_Number] || '').trim().toUpperCase();
  if (/^A82/.test(itemNumber)) return '2000LB';
  if (/^A24/.test(itemNumber)) return '3000LB';
  return '';
}

function isForgedStainlessCatalogProductGroup_(productGroup) {
  return String(productGroup || '').toLowerCase().indexOf('forged stainless steel fittings') !== -1;
}

function getCatalogConnectionDisplayLabel_(value) {
  const normalized = normalizeVariantKey_(value).toLowerCase();
  if (normalized === 'threaded') return 'Threaded';
  if (normalized === 'socket-weld' || normalized === 'socket weld' || normalized === 'socket') return 'Socket-Weld';
  return getDisplayVariantName_(value || '');
}

function getValveMaterialFamily_(row, col) {
  const itemSeriesFamily = getValveMaterialFamilyFromSeries_(getValveItemSeries_(row, col));
  if (itemSeriesFamily) return itemSeriesFamily;

  const raw = [
    row[col.Variant],
    getOptionalCellValue_(row, col, 'Variant_2')
  ].map(value => String(value || '').trim()).filter(Boolean).join(' ');
  const normalized = raw.toLowerCase();

  if (normalized.indexOf('stainless') !== -1 || normalized.indexOf('304') !== -1 || normalized.indexOf('316') !== -1 || normalized.indexOf('ss') !== -1) {
    return 'Stainless';
  }
  if (normalized.indexOf('lead free bronze') !== -1 || normalized.indexOf('lf bronze') !== -1 || normalized.indexOf('lead free') !== -1) {
    return 'Lead Free Bronze';
  }
  if (normalized.indexOf('carbon steel') !== -1) return 'Carbon Steel';
  if (normalized.indexOf('bronze') !== -1) return 'Bronze';
  if (normalized.indexOf('brass') !== -1) return 'Brass';

  return getDisplayVariantName_(row[col.Variant] || '') || 'Material';
}

function getValveMaterialFamilyFromSeries_(itemSeries) {
  const normalized = String(itemSeries || '').trim().toUpperCase();
  if (normalized === 'A12') return 'Brass';
  if (normalized === 'A14') return 'Forged Brass';
  if (normalized === 'A16') return 'Lead Free Bronze';
  if (normalized === 'A24') return 'Carbon Steel';
  if (normalized === 'A35') return 'Stainless';
  return '';
}

function getValveMaterialDetail_(row, col) {
  const detail = getDisplayVariantName_(row[col.Variant] || '');
  const family = getValveMaterialFamily_(row, col);
  if (!detail) return family;
  if (detail.toLowerCase() === family.toLowerCase()) return family;
  return detail;
}

function getValveItemSeries_(row, col) {
  const itemNumber = String(row[col.Item_Number] || '').trim();
  const match = itemNumber.match(/^(A\d{2})/i);
  return match ? match[1].toUpperCase() : '';
}

function getValveConnectionType_(row, col) {
  return normalizeVariant2Key_(getOptionalCellValue_(row, col, 'Variant_2'));
}

function getValveMaterialSortWeight_(materialFamily) {
  const normalized = String(materialFamily || '').toLowerCase();
  if (normalized === 'brass') return 10;
  if (normalized === 'forged brass') return 15;
  if (normalized === 'bronze') return 20;
  if (normalized === 'lead free bronze') return 30;
  if (normalized === 'carbon steel') return 40;
  if (normalized === 'stainless') return 50;
  return 100;
}

function getValveItemSeriesSortWeight_(itemSeries) {
  const normalized = String(itemSeries || '').toUpperCase();
  if (normalized === 'A12') return 10;
  if (normalized === 'A14') return 20;
  return normalized ? 30 : 100;
}

function buildValveSectionQualifier_(group) {
  const fittingType = String(group && group.fittingType || '').trim().toLowerCase();
  const connectionType = String(group && group.connectionType || '').trim();
  const materialFamily = String(group && group.materialFamily || '').trim();

  if (fittingType === 'dielectric unions') {
    if (/^fpt\s*x\s*c$/i.test(connectionType)) {
      return 'Iron FPT x Brass Sweat';
    }
    return 'Iron FPT x Brass';
  }

  return [
    materialFamily,
    connectionType
  ].filter(Boolean).join(' - ');
}

function setMissingImagePlaceholder_(sheet, row, column) {
  sheet.getRange(row, column)
    .setValue('MISSING IMAGE')
    .setFontColor('#C00000')
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
}

function buildCatalogVariantTableRows_(rows, col, options) {
  const tableOptions = options || {};
  const rowMap = {};

  rows.forEach(row => {
    const sizeParts = getCatalogRowSizeParts_(row, col);
    const size = buildCatalogSize_(row, col);
    const displaySize = buildCatalogDisplaySize_(row, col);
    const normalizedSizeKey = normalizeCatalogSizeForComparison_(size) || String(size || '').trim();
    const key = tableOptions.pivotBySizeOnly ? normalizedSizeKey : [
      size,
      row[col.Inner_Carton],
      row[col.Master_Case]
    ].join('|');

    if (!rowMap[key]) {
      rowMap[key] = {
        size: displaySize,
        innerCarton: row[col.Inner_Carton],
        masterCase: row[col.Master_Case],
        variants: {},
        sortOrder: 0,
        size1Sort: parseCatalogSizeSortValue_(sizeParts[0]),
        size2Sort: parseCatalogSizeSortValue_(sizeParts[1]),
        size3Sort: parseCatalogSizeSortValue_(sizeParts[2])
      };
    } else if (tableOptions.pivotBySizeOnly) {
      rowMap[key].innerCarton = resolveMergedCatalogCartonValue_(rowMap[key].innerCarton, row[col.Inner_Carton]);
      rowMap[key].masterCase = resolveMergedCatalogCartonValue_(rowMap[key].masterCase, row[col.Master_Case]);
    }

    const variant = normalizeVariantKey_(row[col.Variant]);
    rowMap[key].variants[variant] = {
      itemNumber: row[col.Item_Number],
      price: getCatalogRowPriceValue_(row, col),
      innerCarton: row[col.Inner_Carton],
      masterCase: row[col.Master_Case]
    };
  });

  return Object.values(rowMap).sort((a, b) => {
    return [
      a.size1Sort - b.size1Sort,
      a.size2Sort - b.size2Sort,
      a.size3Sort - b.size3Sort,
      String(a.size).localeCompare(String(b.size), undefined, { numeric: true })
    ].find(result => result !== 0) || 0;
  });
}

function resolveMergedCatalogCartonValue_(currentValue, nextValue) {
  const currentTrimmed = String(currentValue || '').trim();
  const nextTrimmed = String(nextValue || '').trim();

  if (!currentTrimmed) return nextValue;
  if (!nextTrimmed) return currentValue;

  const currentNumber = Number(currentTrimmed);
  const nextNumber = Number(nextTrimmed);

  if (!isNaN(currentNumber) && !isNaN(nextNumber)) {
    return currentNumber <= nextNumber ? currentTrimmed : nextTrimmed;
  }

  return currentValue;
}

function buildCatalogCartonLayoutFlags_(pivotRows, variants, productGroup) {
  const forceVariantSpecificMasterCase = isForgedSteelCatalogProductGroup_(productGroup) &&
    (variants || []).length > 1 &&
    hasAnyCatalogVariantCartonValues_(pivotRows, variants, 'masterCase');
  const hasVariantSpecificInnerCarton = hasCatalogVariantCartonConflict_(pivotRows, variants, 'innerCarton');
  const hasVariantSpecificMasterCase = forceVariantSpecificMasterCase
    ? hasAnyCatalogVariantCartonValues_(pivotRows, variants, 'masterCase')
    : hasCatalogVariantCartonConflict_(pivotRows, variants, 'masterCase');
  const hideInnerCarton = !hasVariantSpecificInnerCarton &&
    pivotRows.every(row => !String(row.innerCarton || '').trim());
  const hideMasterCase = !hasVariantSpecificMasterCase &&
    pivotRows.every(row => !String(row.masterCase || '').trim());

  return {
    hideInnerCarton,
    hideMasterCase,
    variantSpecificInnerCarton: hasVariantSpecificInnerCarton,
    variantSpecificMasterCase: hasVariantSpecificMasterCase,
    compactTable: hideInnerCarton && hideMasterCase
  };
}

function hasCatalogVariantKey_(variants, targetValue) {
  const normalizedTarget = normalizeCatalogConnectionKey_(targetValue) || String(targetValue || '').toLowerCase();
  return (variants || []).some(variant => {
    const normalizedVariant = normalizeCatalogConnectionKey_(variant) || normalizeVariantKey_(variant).toLowerCase();
    return normalizedVariant === normalizedTarget;
  });
}

function hasAnyCatalogVariantCartonValues_(pivotRows, variants, fieldName) {
  return (pivotRows || []).some(row => {
    return (variants || []).some(variant => {
      const value = row && row.variants && row.variants[variant] ? row.variants[variant][fieldName] : '';
      return !!String(value || '').trim();
    });
  });
}

function hasCatalogVariantCartonConflict_(pivotRows, variants, fieldName) {
  if (!variants || variants.length <= 1) return false;

  return (pivotRows || []).some(row => {
    const distinctValues = variants
      .map(variant => row && row.variants && row.variants[variant] ? row.variants[variant][fieldName] : '')
      .map(value => String(value || '').trim())
      .filter(Boolean)
      .filter((value, index, values) => values.indexOf(value) === index);

    return distinctValues.length > 1;
  });
}

function shouldPivotCatalogRowsBySizeOnly_(rows, col, variants) {
  const normalizedVariants = (variants || []).map(variant => normalizeCatalogConnectionKey_(variant) || normalizeVariantKey_(variant).toLowerCase());
  const productGroup = rows[0] && String(rows[0][col.Product_Group] || '').toLowerCase();

  if (productGroup.indexOf('valve') !== -1 && normalizedVariants.length > 1) {
    return true;
  }

  if (productGroup.indexOf('forged steel fittings') !== -1 &&
      normalizedVariants.includes('threaded') &&
      normalizedVariants.includes('socketweld')) {
    return true;
  }

  return normalizedVariants.includes('304') &&
    normalizedVariants.includes('316') &&
    productGroup.indexOf('forged stainless steel fitting') !== -1;
}

function normalizeCatalogConnectionKey_(value) {
  const normalized = String(value || '')
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');

  if (!normalized) return '';
  if (normalized === 'threaded' || normalized === 'thread') return 'threaded';
  if (normalized === 'socketweld' || normalized === 'socket' || normalized === 'sw') return 'socketweld';
  return normalized;
}

function getCatalogVariantColumns_(rows, col) {
  const variants = [];

  rows.forEach(row => {
    const variant = normalizeVariantKey_(row[col.Variant]);
    if (variant && !variants.includes(variant)) variants.push(variant);
  });

  return variants.sort((a, b) => variantSortWeight_(a) - variantSortWeight_(b));
}

function buildCatalogSize_(row, col) {
  return getCatalogRowSizeParts_(row, col).join(' x ');
}

function buildCatalogDisplaySize_(row, col) {
  const fittingType = row[col.Fitting_Type];
  const normalizedFittingType = normalizeCatalogFittingTypeForSort_(fittingType);
  const productGroup = String(row[col.Product_Group] || '').toLowerCase();
  const sizeParts = getCatalogRowSizeParts_(row, col);
  const size1 = sizeParts[0];
  const size2 = sizeParts[1];
  const size3 = sizeParts[2];

  if (productGroup.indexOf('forged stainless steel fitting') !== -1) {
    if (normalizedFittingType === 'outlets' && size1 && size2 && size3) {
      return [
        formatCompactCatalogSizePart_(size1),
        `(${formatCompactCatalogSizePart_(size2)}\u2011${formatCompactCatalogSizePart_(size3)})`
      ].join('\u00A0');
    }

    return [size1, size2, size3]
      .filter(value => value !== '' && value !== null && value !== undefined)
      .map(formatCompactCatalogSizePart_)
      .join('\u00A0x\u00A0');
  }

  return buildCatalogSize_(row, col);
}

function getCatalogRowSizeParts_(row, col) {
  const explicitSizes = [row[col.Size_1], row[col.Size_2], row[col.Size_3]]
    .map(value => value === null || value === undefined ? '' : String(value).trim());
  const productGroup = String(row[col.Product_Group] || '').trim();
  const expectedSizeCount = getCatalogFittingSizeDimensionCount_(row[col.Fitting_Type]);

  if (productGroup !== 'Malleable Iron Fittings' || expectedSizeCount <= 1) {
    return explicitSizes.filter(Boolean);
  }
  if (explicitSizes.slice(0, expectedSizeCount).every(Boolean)) {
    return explicitSizes.filter(Boolean);
  }

  const inferredSizes = inferMalleableSizesFromItemNumber_(row[col.Item_Number], expectedSizeCount);
  if (!inferredSizes.length ||
      normalizeCatalogSizeForComparison_(explicitSizes[0]) !== normalizeCatalogSizeForComparison_(inferredSizes[0])) {
    return explicitSizes.filter(Boolean);
  }

  return inferredSizes.map((inferredSize, index) => explicitSizes[index] || inferredSize).filter(Boolean);
}

function getCatalogFittingSizeDimensionCount_(fittingType) {
  const normalized = normalizeCatalogFittingTypeForSort_(fittingType);
  if (normalized === 'reducing tees' || normalized === 'reducing crosses') return 3;
  if (normalized === '90 degree reducing elbows' ||
      normalized === '90 degree reducing street elbows' ||
      normalized === 'reducing street elbows' ||
      normalized === 'reducing couplings' ||
      normalized === 'hex bushings') return 2;
  return 1;
}

function inferMalleableSizesFromItemNumber_(itemNumber, sizeCount) {
  const suffixLength = sizeCount * 2;
  const match = String(itemNumber || '').trim().match(new RegExp(`(\\d{${suffixLength}})[A-Za-z]*$`));
  if (!match) return [];

  const sizeCodeMap = {
    '02': '1/8"',
    '04': '1/4"',
    '06': '3/8"',
    '08': '1/2"',
    '10': '5/8"',
    '12': '3/4"',
    '16': '1"',
    '20': '1-1/4"',
    '24': '1-1/2"',
    '28': '1-3/4"',
    '32': '2"',
    '40': '2-1/2"',
    '48': '3"',
    '56': '3-1/2"',
    '64': '4"',
    '72': '6"',
    '84': '8"'
  };
  const encodedSizes = match[1].match(/\d{2}/g) || [];
  const sizes = encodedSizes.map(code => sizeCodeMap[code] || '');
  return sizes.every(Boolean) ? sizes : [];
}

function getCatalogSectionSizeHeader_(fittingType) {
  const normalized = normalizeCatalogFittingTypeForSort_(fittingType);
  if (normalized === 'reducing tees') return 'Size (1 x 2 x 3)';
  return 'Size';
}

function normalizeCatalogSizeForComparison_(size) {
  return String(size || '')
    .replace(/&quot;/gi, '"')
    .replace(/\s+/g, '')
    .toLowerCase();
}

function formatCompactCatalogSizePart_(value) {
  const normalized = String(value || '')
    .replace(/&quot;/gi, '"')
    .replace(/-/g, '\u2011')
    .trim();

  if (!normalized) return '';
  return /"$/.test(normalized) ? normalized : `${normalized}"`;
}

function normalizeVariantKey_(variant) {
  return normalizeCatalogWholeNumberLabel_(variant);
}

function normalizeVariant2Key_(variant2) {
  return normalizeCatalogWholeNumberLabel_(variant2);
}

function normalizeCatalogWholeNumberLabel_(value) {
  const label = String(value || '').trim();
  const wholeNumberDecimal = label.match(/^(\d+)\.0+$/);
  return wholeNumberDecimal ? wholeNumberDecimal[1] : label;
}

function buildCatalogSectionDisplayTitle_(fittingType, variant2, productGroup) {
  const baseTitle = String(fittingType || 'Products').trim() || 'Products';
  const secondaryVariant = normalizeVariant2Key_(variant2);
  const normalizedProductGroup = String(productGroup || '').toLowerCase();

  if (normalizedProductGroup.indexOf('stainless steel flange') !== -1) {
    return baseTitle;
  }

  if (normalizedProductGroup.indexOf('forged steel fittings') !== -1 &&
      normalizedProductGroup.indexOf('stainless') === -1) {
    return baseTitle;
  }

  if (normalizedProductGroup.indexOf('forged stainless steel fitting') !== -1) {
    if (/^threaded$/i.test(secondaryVariant)) return `${baseTitle} - Threaded`;
    if (/^socket[-\s]?weld$/i.test(secondaryVariant)) return `${baseTitle} - Socket Weld`;
  }

  return secondaryVariant ? `${baseTitle} - ${secondaryVariant}` : baseTitle;
}

function compareCatalogSectionGroups_(a, b) {
  return [
    getCatalogVariant2SortWeight_(a.variant2) - getCatalogVariant2SortWeight_(b.variant2),
    getCatalogSectionFittingSortWeight_(a.productGroup, a.fittingType, a.variant2) -
      getCatalogSectionFittingSortWeight_(b.productGroup, b.fittingType, b.variant2),
    String(a.fittingType || '').localeCompare(String(b.fittingType || '')),
    String(a.variant2 || '').localeCompare(String(b.variant2 || ''))
  ].find(result => result !== 0) || 0;
}

function getCatalogSectionFittingSortWeight_(productGroup, fittingType, variant2) {
  const normalizedProductGroup = String(productGroup || '').toLowerCase();

  if (normalizedProductGroup.indexOf('forged stainless steel fitting') !== -1) {
    return getForgedStainlessFittingSortWeight_(fittingType, variant2);
  }

  return getCatalogFittingTypeSortWeight_(productGroup, fittingType);
}

function getForgedStainlessFittingSortWeight_(fittingType, variant2) {
  const normalizedFittingType = normalizeCatalogFittingTypeForSort_(fittingType);
  const normalizedVariant2 = normalizeVariant2Key_(variant2).toLowerCase();
  const socketOrder = [
    'tees',
    '90-degree elbows',
    '45-degree elbows',
    'reducing couplings',
    'solid caps',
    'full couplings',
    'unions',
    'outlets',
    'socket weld inserts',
    'crosses'
  ];
  const threadedOrder = [
    '90-degree elbows',
    '45-degree elbows',
    '90-degree street elbows',
    'unions',
    'tees',
    'solid caps',
    'full couplings',
    'half couplings',
    'square head plugs',
    'hex bushings',
    'reducing couplings',
    'outlets',
    'crosses',
    'socket weld inserts'
  ];
  const order = normalizedVariant2 === 'socket-weld' || normalizedVariant2 === 'socket weld'
    ? socketOrder
    : threadedOrder;
  const index = order.indexOf(normalizedFittingType);

  return index === -1 ? 1000 + getCatalogFittingTypeSortWeight_('Forged Stainless Steel Fittings', fittingType) : index;
}

function getCatalogVariant2SortWeight_(variant2) {
  const normalized = normalizeVariant2Key_(variant2).toLowerCase();

  if (!normalized) return 0;
  if (normalized === '150lb' || normalized === '150 lb') return 10;
  if (normalized === '300lb' || normalized === '300 lb') return 20;
  if (normalized === '3000lb' || normalized === '3000 lb') return 15;
  if (normalized === '2000lb' || normalized === '2000 lb') return 25;
  if (normalized === 'threaded') return 10;
  if (normalized === 'socket-weld' || normalized === 'socket weld' || normalized === 'socket') return 20;
  if (normalized.indexOf('std') !== -1 || normalized.indexOf('standard weight') !== -1) return 10;
  if (normalized.indexOf('lw') !== -1 || normalized.indexOf('light weight') !== -1) return 20;
  if (normalized.indexOf('xh') !== -1 || normalized.indexOf('extra heavy') !== -1) return 30;
  if (normalized === '304') return 10;
  if (normalized === '316') return 20;

  return 100;
}

function getCatalogFittingTypeSortWeight_(productGroup, fittingType) {
  const orderName = getCatalogFittingOrderName_(productGroup);
  const normalizedFittingType = normalizeCatalogFittingTypeForSort_(fittingType);
  const activeOrder = getCatalogFittingOrderMap_(orderName);
  const activeWeight = activeOrder[normalizedFittingType];

  if (activeWeight !== undefined) return activeWeight;

  const generalWeight = getCatalogFittingOrderMap_('general')[normalizedFittingType];
  if (generalWeight !== undefined) return 1000 + generalWeight;

  return 9999;
}

function getCatalogFittingOrderName_(productGroup) {
  const normalized = String(productGroup || '').toLowerCase();

  if (normalized.indexOf('merchant steel fittings') !== -1) return 'merchantSteelFittings';
  if (normalized.indexOf('forged steel fittings') !== -1) return 'forgedSteelFittings';
  if (normalized.indexOf('valve') !== -1) return 'valve';
  if (normalized.indexOf('butt') !== -1 || normalized.indexOf('weld fitting') !== -1) return 'buttWeld';
  if (normalized.indexOf('flange') !== -1) return 'flange';
  if (normalized.indexOf('forged') !== -1) return 'forged';

  return 'general';
}

function getCatalogFittingOrderMap_(orderName) {
  const orders = {
    general: [
      '90-degree elbows',
      '45-degree elbows',
      '90-degree street elbows',
      '45-degree street elbows',
      'side outlet elbows',
      'unions',
      '90-degree reducing elbows',
      'tees',
      'reducing tees',
      'street service tees',
      'side outlet tees',
      'crosses',
      'locknuts',
      'hex bushings',
      'square head plugs',
      'solid caps',
      'full flanges',
      'reducing street elbows',
      'reducing couplings',
      'standard couplings',
      'extension pieces',
      '45-degree y laterals'
    ],
    forged: [
      'full couplings',
      'half couplings',
      'socket weld inserts',
      'outlets',
      'tank flanges',
      'reducing couplings'
    ],
    forgedSteelFittings: [
      '90-degree elbows',
      '45-degree elbows',
      '90-degree street elbows',
      '45-degree street elbows',
      'unions',
      'tees',
      'solid caps',
      'full couplings',
      'half couplings',
      'square head plugs',
      'hex bushings',
      'crosses',
      'outlets',
      'socket weld inserts',
      'tank flanges',
      'reducing couplings'
    ],
    merchantSteelFittings: [
      'square head plugs',
      'flush hex socket countersunk plugs',
      'hex socket countersunk plugs',
      'square socket countersunk plugs',
      'face bushings',
      'hose barbs',
      'hex bushings'
    ],
    flange: [
      'threaded',
      'blind',
      'slip-on',
      'weld neck',
      'socket weld',
      'lap joint'
    ],
    buttWeld: [
      '90-degree long radius elbows',
      '90-degree short radius elbows',
      '45-degree elbows',
      'solid caps',
      'tees',
      'reducing tees',
      'concentric reducers',
      'eccentric reducers',
      'long radius return bends',
      'short radius return bends'
    ],
    valve: [
      '200 lb hornet wog brass gate valves',
      '200 lb hornet wog brass globe valves',
      '200 lb wog brass hose stop',
      '600 lb wog brass full port ball valves',
      '600 lb wog brass full port lead free ball valves',
      '600 lb wog brass standard port ball valves',
      'fpt x fpt swing check valves',
      'fpt x fpt gas ball valves',
      'mpt x male flare gas ball valves',
      'male flare x male flare gas ball valves',
      'female flare x male flare gas ball valves',
      'mini fpt x fpt ball valves',
      '200 lb wog brass y-strainer',
      'mini mpt x fpt ball valves',
      '1000 lb wog 316ss full port ball valves',
      '2000 lb wog 316ss full port ball valves',
      '2000 lb wog 316ss 3 piece full port ball valves',
      '200 lb wog 316ss gate valves',
      '200 lb wog 316ss swing check valves',
      '2000 lb carbon steel full port ball valves',
      '2000 lb carbon steel standard port ball valves',
      'iron fpt x brass sweat dielectric unions'
    ]
  };

  const order = orders[orderName] || orders.general;
  return order.reduce((map, fittingType, index) => {
    map[normalizeCatalogFittingTypeForSort_(fittingType)] = index;
    return map;
  }, {});
}

function normalizeCatalogFittingTypeForSort_(fittingType) {
  let normalized = String(fittingType || '').toLowerCase().trim();

  normalized = normalized
    .replace(/°/g, '-degree')
    .replace(/\s*-\s*degree/g, '-degree')
    .replace(/&/g, ' and ')
    .replace(/\//g, ' ')
    .replace(/[()]/g, ' ')
    .replace(/-/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  const aliases = {
    '90 degree elbows': '90-degree elbows',
    '45 degree elbows': '45-degree elbows',
    '90 degree street elbows': '90-degree street elbows',
    '45 degree street elbows': '45-degree street elbows',
    'reducing elbows': '90-degree reducing elbows',
    'street service tee': 'street service tees',
    'street service tees': 'street service tees',
    'square head plugs': 'square head plugs',
    'cored or hollow hex head plugs': 'square head plugs',
    'cored or hollow square head plugs': 'square head plugs',
    'solid square head plugs': 'square head plugs',
    'solid hex head plugs': 'square head plugs',
    'flush hex socket countersunk plugs': 'flush hex socket countersunk plugs',
    'flush hex socket countersunk plug': 'flush hex socket countersunk plugs',
    'hex socket countersunk plugs': 'hex socket countersunk plugs',
    'square socket countersunk plugs': 'square socket countersunk plugs',
    'face bushings': 'face bushings',
    'hose barb x mpt or combination nipple or king nipple': 'hose barbs',
    'hose barbs': 'hose barbs',
    'caps': 'solid caps',
    'floor flanges': 'full flanges',
    'full flanges': 'full flanges',
    'reducer couplings': 'reducing couplings',
    'reducing couplings': 'reducing couplings',
    'couplings banded': 'standard couplings',
    'standard couplings': 'standard couplings',
    'full couplings': 'full couplings',
    'couplings sockets': 'full couplings',
    'half couplings': 'half couplings',
    'inserts': 'socket weld inserts',
    'socket weld inserts': 'socket weld inserts',
    'weld o lets': 'outlets',
    'weldolets': 'outlets',
    'outlets': 'outlets',
    '45 degree y laterals': '45-degree y laterals',
    '90 degree long radius elbows': '90-degree long radius elbows',
    '90 degree short radius elbows': '90-degree short radius elbows',
    'return bends long radius': 'long radius return bends',
    'return bends short radius': 'short radius return bends'
  };

  return aliases[normalized] || normalized;
}

function variantSortWeight_(variant) {
  const normalized = String(variant).toLowerCase();
  if (normalized === 'threaded') return 10;
  if (normalized === 'socket-weld' || normalized === 'socket weld') return 20;
  if (normalized.indexOf('s40') !== -1 && normalized.indexOf('black') !== -1) return 10;
  if (normalized.indexOf('s40') !== -1 && (normalized.indexOf('galv') !== -1 || normalized.indexOf('zinc') !== -1)) return 20;
  if (normalized.indexOf('s80') !== -1 && normalized.indexOf('black') !== -1) return 30;
  if (normalized.indexOf('s80') !== -1 && (normalized.indexOf('galv') !== -1 || normalized.indexOf('zinc') !== -1)) return 40;
  if (normalized.indexOf('a106') !== -1) return 50;
  if (normalized === 'steel' || normalized === 'plain' || normalized === 'black' || normalized === '304') return 10;
  if (normalized === 'plated' || normalized === 'zinc plated' || normalized === 'galvanized' || normalized === '316') return 20;
  return 100;
}

function getDisplayVariantName_(variant) {
  const normalized = String(variant).toLowerCase();
  if (normalized === 'steel' || normalized === 'plain') return 'Plain (Steel)';
  if (normalized === 'plated' || normalized === 'zinc plated') return 'Zinc Plated';
  if (normalized === 'black') return 'Black';
  if (normalized === 'galvanized') return 'Galvanized';
  if (normalized === '304') return '304';
  if (normalized === '316') return '316';
  if (/^(s40|s80)\s+/i.test(String(variant || ''))) return String(variant || '').trim();
  if (/^a106/i.test(String(variant || ''))) return String(variant || '').trim();
  return String(variant || '').trim() || 'Variant';
}

function getVariantHeaderColors_(variant, options) {
  const context = options || {};
  const productGroup = String(context.productGroup || '').toLowerCase();
  const variantIndex = Number(context.variantIndex) || 0;
  const variantCount = Number(context.variantCount) || 0;
  const isSingleVariant = !!context.isSingleVariant || variantCount <= 1;
  const normalized = String(variant).toLowerCase();
  if (isSingleVariant) return { background: '#274A96', font: '#FFFFFF' };
  if (productGroup.indexOf('valve') !== -1) {
    if (variantIndex === 0) return { background: '#274A96', font: '#FFFFFF' };
    if (variantIndex === 1) return { background: '#4d9b61', font: '#000000' };
  }
  if (normalized === '150lb' || normalized === '150 lb') return { background: '#274A96', font: '#FFFFFF' };
  if (normalized === '300lb' || normalized === '300 lb') return { background: '#6F9ED3', font: '#FFFFFF' };
  if (normalized === 'threaded') return { background: '#274A96', font: '#FFFFFF' };
  if (normalized === 'socket-weld' || normalized === 'socket weld') return { background: '#6F9ED3', font: '#FFFFFF' };
  if (normalized.indexOf('std') !== -1 || normalized.indexOf('standard weight') !== -1) return { background: '#274A96', font: '#FFFFFF' };
  if (normalized.indexOf('xh') !== -1 || normalized.indexOf('extra heavy') !== -1) return { background: '#D9D9D9', font: '#000000' };
  if (normalized === '304') return { background: '#274A96', font: '#FFFFFF' };
  if (normalized === '316') return { background: '#6F9ED3', font: '#FFFFFF' };

  const isSecondary = normalized === 'plated' ||
    normalized === 'zinc plated' ||
    normalized === 'galvanized' ||
    normalized === '316' ||
    normalized.indexOf('galv') !== -1 ||
    normalized.indexOf('zinc') !== -1;

  return isSecondary
    ? { background: '#D9D9D9', font: '#000000' }
    : { background: '#000000', font: '#FFFFFF' };
}

function buildCatalogSectionTitle_(catalogTitle, fittingType) {
  const normalizedFittingType = String(fittingType || '').trim();
  return normalizedFittingType || String(catalogTitle || '').trim();
}

function parseCatalogSizeSortValue_(size) {
  const firstSize = String(size || '').split(' x ')[0].replace(/"/g, '').trim();
  if (!firstSize) return Number.MAX_SAFE_INTEGER;

  if (/^close\b/i.test(firstSize)) return 0;

  const mixedMatch = firstSize.match(/^(\d+)-(\d+)\/(\d+)$/);
  if (mixedMatch) {
    return Number(mixedMatch[1]) + Number(mixedMatch[2]) / Number(mixedMatch[3]);
  }

  const spacedMixedMatch = firstSize.match(/^(\d+)\s+(\d+)\/(\d+)$/);
  if (spacedMixedMatch) {
    return Number(spacedMixedMatch[1]) + Number(spacedMixedMatch[2]) / Number(spacedMixedMatch[3]);
  }

  const fractionMatch = firstSize.match(/^(\d+)\/(\d+)$/);
  if (fractionMatch) {
    return Number(fractionMatch[1]) / Number(fractionMatch[2]);
  }

  const number = Number(firstSize);
  return isNaN(number) ? 0 : number;
}

// Legacy reference implementation from the older sheet-based catalog PDF path.
function buildCatalogPdfSheet_(sheet, rows, col, meta) {
  const blue = '#1F4E78';
  const lightBlue = '#D9EAF7';
  const veryLightBlue = '#EEF6FC';
  const white = '#FFFFFF';

  sheet.clear();
  sheet.setHiddenGridlines(true);

  sheet.setColumnWidth(1, 104);
  sheet.setColumnWidth(2, 104);
  sheet.setColumnWidth(3, 60);
  sheet.setColumnWidth(4, 72);
  sheet.setColumnWidth(5, 72);
  sheet.setColumnWidth(6, 108);
  sheet.setColumnWidth(7, 82);
  sheet.setColumnWidth(8, 108);
  sheet.setColumnWidth(9, 82);

  sheet.getRange('A1:I4')
    .setBackground(blue)
    .setFontColor(white)
    .setVerticalAlignment('middle');

  if (meta.logoFileId) {
    insertResizedDriveImage_(sheet, meta.logoFileId, 1, 1, 132, 46, 12, 10);
  } else {
    sheet.getRange('A1').setValue('MSI').setFontSize(22).setFontWeight('bold');
  }

  sheet.getRange('C1:I1').merge();
  sheet.getRange('C1')
    .setValue(meta.title)
    .setFontSize(18)
    .setFontWeight('bold')
    .setHorizontalAlignment('right');

  sheet.getRange('C2:I2').merge();
  sheet.getRange('C2')
    .setValue(meta.subtitle)
    .setFontSize(11)
    .setFontWeight('bold')
    .setHorizontalAlignment('right');

  sheet.getRange('C3:I3').merge();
  sheet.getRange('C3')
    .setValue(buildCatalogMetaLine_(meta))
    .setFontSize(10)
    .setHorizontalAlignment('right');

  sheet.setRowHeight(1, 40);
  sheet.setRowHeight(2, 24);
  sheet.setRowHeight(3, 22);
  sheet.setRowHeight(4, 12);

  let currentRow = 6;
  const groupedRows = groupCatalogRowsByFittingType_(rows, col);
  const tableStartRow = currentRow;
  const tableStartCol = 3;
  let maxTableColumnCount = 7;

  groupedRows.forEach((group, groupIndex) => {
    if (groupIndex > 0) {
      sheet.setRowHeight(currentRow, 36);
      currentRow++;
    }

    const sectionStartRow = currentRow;
    const pivotRows = buildCatalogVariantTableRows_(group.rows, col);
    const variants = getCatalogVariantColumns_(group.rows, col);
    const tableColumnCount = 3 + variants.length * 2;
    maxTableColumnCount = Math.max(maxTableColumnCount, tableColumnCount);
    const imagePanelRows = pivotRows.length + 3;

    sheet.getRange(sectionStartRow, 1, imagePanelRows, 2)
      .merge()
      .setBackground(white)
      .setBorder(false, false, false, false, false, false)
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center');

    const imageFileId = getFirstCatalogImageFileId_(group.rows, col);
    if (imageFileId) {
      insertResizedDriveImage_(sheet, imageFileId, 1, sectionStartRow, 160, 142, 24, 12, 'MISSING IMAGE');
    } else {
      setMissingImagePlaceholder_(sheet, sectionStartRow, 1);
    }

    sheet.getRange(currentRow, tableStartCol, 1, tableColumnCount).merge();
    sheet.getRange(currentRow, tableStartCol)
      .setValue(buildCatalogSectionTitle_(meta.title, group.fittingType))
      .setBackground(blue)
      .setFontColor(white)
      .setFontWeight('bold')
      .setFontSize(13)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheet.setRowHeight(currentRow, 27);
    currentRow++;

    sheet.getRange(currentRow, tableStartCol, 1, 3)
      .setValues([['Size', 'Inner Carton', 'Master Carton']])
      .setBackground(lightBlue)
      .setFontColor('#000000')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setWrap(true);

    variants.forEach((variant, index) => {
      const startCol = tableStartCol + 3 + index * 2;
      const colors = getVariantHeaderColors_(variant);

      sheet.getRange(currentRow, startCol, 1, 2)
        .merge()
        .setValue(getDisplayVariantName_(variant))
        .setBackground(colors.background)
        .setFontColor(colors.font)
        .setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle')
        .setWrap(true);
    });

    sheet.getRange(currentRow, tableStartCol, 1, tableColumnCount)
      .setBorder(true, true, true, true, true, true);
    const subHeaderRow = currentRow;
    const dimensionHeaderRow = currentRow;
    currentRow++;

    sheet.getRange(currentRow, tableStartCol, 1, 3)
      .setValues([['', '', '']])
      .setBackground(lightBlue)
      .setFontColor('#000000')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setWrap(true);

    variants.forEach((variant, index) => {
      const startCol = tableStartCol + 3 + index * 2;

      sheet.getRange(currentRow, startCol, 1, 2)
        .setValues([['Item Number', 'List Price']])
        .setBackground(lightBlue)
        .setFontColor('#000000')
        .setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle')
        .setWrap(true)
        .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
    });

    for (let i = 0; i < 3; i++) {
      sheet.getRange(dimensionHeaderRow, tableStartCol + i, 2, 1)
        .merge()
        .setBackground(lightBlue)
        .setFontColor('#000000')
        .setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle')
        .setWrap(true);
    }

    sheet.getRange(currentRow, tableStartCol, 1, tableColumnCount)
      .setBorder(true, true, true, true, true, true);
    sheet.getRange(subHeaderRow, tableStartCol, 2, tableColumnCount)
      .setBorder(true, true, true, true, true, true);
    currentRow++;
    const firstDataRow = currentRow;

    pivotRows.forEach((pivotRow, index) => {
      const rowValues = [
        formatCatalogVisibleCellValue_(pivotRow.size),
        formatCatalogVisibleCellValue_(pivotRow.innerCarton),
        formatCatalogVisibleCellValue_(pivotRow.masterCase)
      ];

      variants.forEach(variant => {
        const item = pivotRow.variants[variant] || {};
        rowValues.push(formatCatalogVisibleCellValue_(item.itemNumber));
        rowValues.push(item.price === '' || item.price === null || item.price === undefined ? '-' : item.price);
      });

      sheet.getRange(currentRow, tableStartCol, 1, tableColumnCount).setValues([rowValues]);

      const rowColor = index % 2 === 0 ? white : veryLightBlue;
      sheet.getRange(currentRow, tableStartCol, 1, tableColumnCount)
        .setBackground(rowColor)
        .setBorder(false, false, false, false, false, false)
        .setVerticalAlignment('middle');

      sheet.getRange(currentRow, tableStartCol, 1, tableColumnCount).setHorizontalAlignment('center');
      variants.forEach((variant, variantIndex) => {
        const priceCol = tableStartCol + 4 + variantIndex * 2;
        sheet.getRange(currentRow, priceCol, 1, 1).setNumberFormat('$#,##0.00');
      });
      currentRow++;
    });

    if (variants.length) {
      const variantHeaderStartCol = tableStartCol + 3;
      const variantHeaderColumnCount = variants.length * 2;

      sheet.getRange(subHeaderRow + 1, variantHeaderStartCol, 1, variantHeaderColumnCount)
        .setBorder(null, null, true, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

      sheet.getRange(firstDataRow, variantHeaderStartCol, 1, variantHeaderColumnCount)
        .setBorder(true, null, null, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }

    currentRow = Math.max(currentRow, sectionStartRow + imagePanelRows);
  });

  const lastUsedRow = Math.max(currentRow, 1);
  sheet.getRange(1, 1, lastUsedRow, 9).setFontFamily('Arial');
  sheet.getRange(1, 1, lastUsedRow, 9).setFontSize(9);
  sheet.getRange('C1').setFontSize(18);
  sheet.getRange('C2').setFontSize(11);
  sheet.getRange('C3').setFontSize(10);
  sheet.setFrozenRows(5);
}

function groupCatalogRowsByFittingType_(rows, col) {
  const groupMap = {};

  rows.forEach(row => {
    const fittingType = String(row[col.Fitting_Type] || 'Products').trim() || 'Products';
    const variant2 = normalizeVariant2Key_(getOptionalCellValue_(row, col, 'Variant_2'));
    const key = `${fittingType}|${variant2}`;

    if (!groupMap[key]) {
      groupMap[key] = {
        plcValues: [],
        productGroup: row[col.Product_Group],
        fittingType,
        variant2,
        rows: []
      };
    }

    const plc = normalizeMatchValue_(row[col.PLC]);
    if (plc && !groupMap[key].plcValues.includes(plc)) {
      groupMap[key].plcValues.push(plc);
    }

    groupMap[key].rows.push(row);
  });

  return Object.values(groupMap).sort(compareCatalogSectionGroups_);
}

function buildCatalogSectionModels_(rows, col) {
  if (rows.length && isNippleCatalogProductGroup_(rows[0][col.Product_Group])) {
    return buildCatalogNippleSectionModels_(rows, col);
  }
  if (rows.length && isValveCatalogProductGroup_(rows[0][col.Product_Group])) {
    return buildCatalogValveSectionModels_(rows, col);
  }
  if (rows.length && isForgedSteelCatalogProductGroup_(rows[0][col.Product_Group])) {
    return buildCatalogForgedSteelSectionModels_(rows, col);
  }
  if (rows.length && isCarbonSteelButtWeldCatalogProductGroup_(rows[0][col.Product_Group])) {
    return buildCatalogCarbonSteelButtWeldSectionModels_(rows, col);
  }
  if (rows.length && isStainlessSteelButtWeldCatalogProductGroup_(rows[0][col.Product_Group])) {
    return buildCatalogStainlessSteelButtWeldSectionModels_(rows, col);
  }
  if (rows.length && isStainlessSteelFlangeCatalogProductGroup_(rows[0][col.Product_Group])) {
    return buildCatalogStainlessSteelFlangeSectionModels_(rows, col);
  }

  return groupCatalogRowsByFittingType_(rows, col).map(group => {
    const variants = getCatalogVariantColumns_(group.rows, col);
    const pivotRows = buildCatalogVariantTableRows_(group.rows, col, {
      pivotBySizeOnly: shouldPivotCatalogRowsBySizeOnly_(group.rows, col, variants)
    });
    const cartonLayout = buildCatalogCartonLayoutFlags_(pivotRows, variants, group.rows[0] && group.rows[0][col.Product_Group]);
    const rowCount = pivotRows.length;
    const recommendation = chooseCatalogSectionTemplate_(rowCount);

    return {
      productGroup: group.rows[0] && group.rows[0][col.Product_Group],
      plc: group.plcValues.join(','),
      fittingType: group.fittingType,
      variant2: group.variant2,
      title: buildCatalogSectionDisplayTitle_(group.fittingType, group.variant2, group.rows[0] && group.rows[0][col.Product_Group]),
      skuRows: group.rows.length,
      renderedRows: rowCount,
      tableRows: pivotRows,
      rawVariants: variants,
      variants: variants.map(variant => getDisplayVariantName_(variant)),
      sizeHeader: getCatalogSectionSizeHeader_(group.fittingType),
      imageFileId: getFirstCatalogImageFileId_(group.rows, col),
      recommendedTemplate: recommendation.template,
      reason: recommendation.reason,
      hideInnerCarton: cartonLayout.hideInnerCarton,
      hideMasterCase: cartonLayout.hideMasterCase,
      variantSpecificInnerCarton: cartonLayout.variantSpecificInnerCarton,
      variantSpecificMasterCase: cartonLayout.variantSpecificMasterCase,
      compactTable: cartonLayout.compactTable,
      forceShapeTable: isButtWeldCatalogProductGroup_(group.rows[0] && group.rows[0][col.Product_Group]) ||
        (isForgedSteelCatalogProductGroup_(group.rows[0] && group.rows[0][col.Product_Group]) && variants.length > 1) ||
        cartonLayout.variantSpecificInnerCarton ||
        cartonLayout.variantSpecificMasterCase
    };
  });
}

function isStainlessSteelButtWeldCatalogProductGroup_(productGroup) {
  return String(productGroup || '').toLowerCase().indexOf('stainless steel butt-weld fittings') !== -1 ||
    String(productGroup || '').toLowerCase().indexOf('stainless steel butt weld fittings') !== -1;
}

function isCarbonSteelButtWeldCatalogProductGroup_(productGroup) {
  const normalized = String(productGroup || '').toLowerCase();
  return normalized.indexOf('carbon steel butt-weld fittings') !== -1 ||
    normalized.indexOf('carbon steel butt weld fittings') !== -1;
}

function isButtWeldCatalogProductGroup_(productGroup) {
  const normalized = String(productGroup || '').toLowerCase();
  return normalized.indexOf('butt-weld fittings') !== -1 || normalized.indexOf('butt weld fittings') !== -1;
}

function buildCatalogForgedSteelSectionModels_(rows, col) {
  const groupMap = {};

  rows.forEach(row => {
    const fittingType = String(row[col.Fitting_Type] || 'Products').trim() || 'Products';
    const pressureClass = getForgedSteelPressureClassFromRow_(row, col);
    const key = `${pressureClass}|${fittingType}`;

    if (!groupMap[key]) {
      groupMap[key] = {
        plcValues: [],
        productGroup: row[col.Product_Group],
        fittingType,
        pressureClass,
        rows: []
      };
    }

    const plc = normalizeMatchValue_(row[col.PLC]);
    if (plc && !groupMap[key].plcValues.includes(plc)) {
      groupMap[key].plcValues.push(plc);
    }

    groupMap[key].rows.push(row);
  });

  return Object.values(groupMap).sort((a, b) => {
    return [
      getCatalogVariant2SortWeight_(a.pressureClass) - getCatalogVariant2SortWeight_(b.pressureClass),
      getCatalogFittingTypeSortWeight_(a.productGroup, a.fittingType) -
        getCatalogFittingTypeSortWeight_(b.productGroup, b.fittingType),
      String(a.fittingType || '').localeCompare(String(b.fittingType || '')),
      String(a.pressureClass || '').localeCompare(String(b.pressureClass || ''))
    ].find(result => result !== 0) || 0;
  }).map(group => {
    const variants = getCatalogVariantColumns_(group.rows, col);
    const pivotRows = buildCatalogVariantTableRows_(group.rows, col, {
      pivotBySizeOnly: shouldPivotCatalogRowsBySizeOnly_(group.rows, col, variants)
    });
    const cartonLayout = buildCatalogCartonLayoutFlags_(pivotRows, variants, group.rows[0] && group.rows[0][col.Product_Group]);
    const rowCount = pivotRows.length;
    const recommendation = chooseCatalogSectionTemplate_(rowCount);

    return {
      productGroup: group.rows[0] && group.rows[0][col.Product_Group],
      plc: group.plcValues.join(','),
      fittingType: group.fittingType,
      variant2: group.pressureClass,
      title: buildCatalogSectionDisplayTitle_(group.fittingType, group.pressureClass, group.rows[0] && group.rows[0][col.Product_Group]),
      skuRows: group.rows.length,
      renderedRows: rowCount,
      tableRows: pivotRows,
      rawVariants: variants,
      variants: variants.map(variant => getDisplayVariantName_(variant)),
      sizeHeader: getCatalogSectionSizeHeader_(group.fittingType),
      imageFileId: getFirstCatalogImageFileId_(group.rows, col),
      recommendedTemplate: recommendation.template,
      reason: recommendation.reason,
      hideInnerCarton: cartonLayout.hideInnerCarton,
      hideMasterCase: cartonLayout.hideMasterCase,
      variantSpecificInnerCarton: cartonLayout.variantSpecificInnerCarton,
      variantSpecificMasterCase: cartonLayout.variantSpecificMasterCase,
      compactTable: cartonLayout.compactTable,
      forceShapeTable: variants.length > 1 ||
        cartonLayout.variantSpecificInnerCarton ||
        cartonLayout.variantSpecificMasterCase,
      pageSubtitle: group.pressureClass
    };
  });
}

function buildCatalogCarbonSteelButtWeldSectionModels_(rows, col) {
  const groupMap = {};

  rows.forEach(row => {
    const fittingType = String(row[col.Fitting_Type] || 'Products').trim() || 'Products';

    if (!groupMap[fittingType]) {
      groupMap[fittingType] = {
        plcValues: [],
        productGroup: row[col.Product_Group],
        fittingType,
        rows: []
      };
    }

    const plc = normalizeMatchValue_(row[col.PLC]);
    if (plc && !groupMap[fittingType].plcValues.includes(plc)) {
      groupMap[fittingType].plcValues.push(plc);
    }

    groupMap[fittingType].rows.push(row);
  });

  return Object.values(groupMap).sort((a, b) => {
    return [
      getCatalogFittingTypeSortWeight_(a.productGroup, a.fittingType) -
        getCatalogFittingTypeSortWeight_(b.productGroup, b.fittingType),
      String(a.fittingType || '').localeCompare(String(b.fittingType || ''))
    ].find(result => result !== 0) || 0;
  }).map(group => {
    const variants = getCatalogVariantColumns_(group.rows, col);
    const pivotRows = buildCatalogVariantTableRows_(group.rows, col, {
      pivotBySizeOnly: shouldPivotCatalogRowsBySizeOnly_(group.rows, col, variants)
    });
    const cartonLayout = buildCatalogCartonLayoutFlags_(pivotRows, variants, group.rows[0] && group.rows[0][col.Product_Group]);
    const rowCount = pivotRows.length;
    const recommendation = chooseCatalogSectionTemplate_(rowCount);

    return {
      productGroup: group.rows[0] && group.rows[0][col.Product_Group],
      plc: group.plcValues.join(','),
      fittingType: group.fittingType,
      variant2: '',
      title: buildCatalogSectionDisplayTitle_(group.fittingType, '', group.rows[0] && group.rows[0][col.Product_Group]),
      skuRows: group.rows.length,
      renderedRows: rowCount,
      tableRows: pivotRows,
      rawVariants: variants,
      variants: variants.map(variant => getDisplayVariantName_(variant)),
      sizeHeader: getCatalogSectionSizeHeader_(group.fittingType),
      imageFileId: getFirstCatalogImageFileId_(group.rows, col),
      recommendedTemplate: recommendation.template,
      reason: recommendation.reason,
      hideInnerCarton: cartonLayout.hideInnerCarton,
      hideMasterCase: cartonLayout.hideMasterCase,
      variantSpecificInnerCarton: cartonLayout.variantSpecificInnerCarton,
      variantSpecificMasterCase: cartonLayout.variantSpecificMasterCase,
      compactTable: cartonLayout.compactTable,
      forceShapeTable: true
    };
  });
}

function buildCatalogStainlessSteelButtWeldSectionModels_(rows, col) {
  const groupMap = {};

  rows.forEach(row => {
    const fittingType = String(row[col.Fitting_Type] || 'Products').trim() || 'Products';
    const weight = normalizeVariantKey_(row[col.Variant]);
    const key = `${fittingType}|${weight}`;

    if (!groupMap[key]) {
      groupMap[key] = {
        plcValues: [],
        productGroup: row[col.Product_Group],
        fittingType,
        weight,
        rows: []
      };
    }

    const plc = normalizeMatchValue_(row[col.PLC]);
    if (plc && !groupMap[key].plcValues.includes(plc)) {
      groupMap[key].plcValues.push(plc);
    }

    groupMap[key].rows.push(row);
  });

  return Object.values(groupMap).sort((a, b) => {
    return [
      getCatalogButtWeldWeightSortWeight_(a.weight) - getCatalogButtWeldWeightSortWeight_(b.weight),
      getCatalogFittingTypeSortWeight_(a.productGroup, a.fittingType) - getCatalogFittingTypeSortWeight_(b.productGroup, b.fittingType),
      String(a.fittingType || '').localeCompare(String(b.fittingType || ''))
    ].find(result => result !== 0) || 0;
  }).map(group => {
    const variants = getCatalogStainlessSteelButtWeldVariantColumns_(group.rows, col);
    const pivotRows = buildCatalogStainlessSteelButtWeldTableRows_(group.rows, col);
    const cartonLayout = buildCatalogCartonLayoutFlags_(pivotRows, variants, group.rows[0] && group.rows[0][col.Product_Group]);
    const rowCount = pivotRows.length;
    const recommendation = chooseCatalogSectionTemplate_(rowCount);
    const weightLabel = getDisplayVariantName_(group.weight);

    return {
      productGroup: group.rows[0] && group.rows[0][col.Product_Group],
      plc: group.plcValues.join(','),
      fittingType: group.fittingType,
      variant2: weightLabel,
      title: buildCatalogSectionDisplayTitle_(group.fittingType, weightLabel, group.rows[0] && group.rows[0][col.Product_Group]),
      skuRows: group.rows.length,
      renderedRows: rowCount,
      tableRows: pivotRows,
      rawVariants: variants,
      variants: variants.map(variant => getDisplayVariantName_(variant)),
      sizeHeader: getCatalogSectionSizeHeader_(group.fittingType),
      imageFileId: getFirstCatalogImageFileId_(group.rows, col),
      recommendedTemplate: recommendation.template,
      reason: recommendation.reason,
      hideInnerCarton: cartonLayout.hideInnerCarton,
      hideMasterCase: cartonLayout.hideMasterCase,
      variantSpecificInnerCarton: cartonLayout.variantSpecificInnerCarton,
      variantSpecificMasterCase: cartonLayout.variantSpecificMasterCase,
      compactTable: cartonLayout.compactTable,
      pageSubtitle: weightLabel,
      forceShapeTable: true
    };
  });
}

function getCatalogStainlessSteelButtWeldVariantColumns_(rows, col) {
  const variants = [];

  rows.forEach(row => {
    const variant = normalizeVariant2Key_(getOptionalCellValue_(row, col, 'Variant_2'));
    if (variant && !variants.includes(variant)) variants.push(variant);
  });

  return variants.sort((a, b) => variantSortWeight_(a) - variantSortWeight_(b));
}

function buildCatalogStainlessSteelButtWeldTableRows_(rows, col) {
  const rowMap = {};

  rows.forEach(row => {
    const size = buildCatalogDisplaySize_(row, col);
    const sizeParts = getCatalogRowSizeParts_(row, col);
    const key = normalizeCatalogSizeForComparison_(size) || String(size || '').trim();

    if (!rowMap[key]) {
      rowMap[key] = {
        size,
        innerCarton: row[col.Inner_Carton],
        masterCase: row[col.Master_Case],
        variants: {},
        size1Sort: parseCatalogSizeSortValue_(sizeParts[0]),
        size2Sort: parseCatalogSizeSortValue_(sizeParts[1]),
        size3Sort: parseCatalogSizeSortValue_(sizeParts[2])
      };
    } else {
      rowMap[key].innerCarton = resolveMergedCatalogCartonValue_(rowMap[key].innerCarton, row[col.Inner_Carton]);
      rowMap[key].masterCase = resolveMergedCatalogCartonValue_(rowMap[key].masterCase, row[col.Master_Case]);
    }

    const variant = normalizeVariant2Key_(getOptionalCellValue_(row, col, 'Variant_2'));
    rowMap[key].variants[variant] = {
      itemNumber: row[col.Item_Number],
      price: getCatalogRowPriceValue_(row, col),
      innerCarton: row[col.Inner_Carton],
      masterCase: row[col.Master_Case]
    };
  });

  return Object.values(rowMap).sort((a, b) => {
    return [
      a.size1Sort - b.size1Sort,
      a.size2Sort - b.size2Sort,
      a.size3Sort - b.size3Sort,
      String(a.size).localeCompare(String(b.size), undefined, { numeric: true })
    ].find(result => result !== 0) || 0;
  });
}

function getCatalogButtWeldWeightSortWeight_(weight) {
  const normalized = normalizeVariantKey_(weight).toLowerCase();
  if (normalized.indexOf('std') !== -1 || normalized.indexOf('standard weight') !== -1) return 10;
  if (normalized.indexOf('lw') !== -1 || normalized.indexOf('light weight') !== -1) return 20;
  if (normalized.indexOf('xh') !== -1 || normalized.indexOf('extra heavy') !== -1) return 30;
  return 100;
}

function isStainlessSteelFlangeCatalogProductGroup_(productGroup) {
  return String(productGroup || '').toLowerCase().indexOf('stainless steel flange') !== -1;
}

function buildCatalogStainlessSteelFlangeSectionModels_(rows, col) {
  const groupMap = {};

  rows.forEach(row => {
    const fittingType = String(row[col.Fitting_Type] || 'Products').trim() || 'Products';
    const pressureClass = normalizeVariantKey_(row[col.Variant]);
    const key = `${fittingType}|${pressureClass}`;

    if (!groupMap[key]) {
      groupMap[key] = {
        plcValues: [],
        productGroup: row[col.Product_Group],
        fittingType,
        variant2: pressureClass,
        rows: []
      };
    }

    const plc = normalizeMatchValue_(row[col.PLC]);
    if (plc && !groupMap[key].plcValues.includes(plc)) {
      groupMap[key].plcValues.push(plc);
    }

    groupMap[key].rows.push(row);
  });

  return Object.values(groupMap).sort((a, b) => {
    return [
      getCatalogVariant2SortWeight_(a.variant2) - getCatalogVariant2SortWeight_(b.variant2),
      getCatalogFittingTypeSortWeight_(a.productGroup, a.fittingType) - getCatalogFittingTypeSortWeight_(b.productGroup, b.fittingType),
      String(a.fittingType || '').localeCompare(String(b.fittingType || '')),
      String(a.variant2 || '').localeCompare(String(b.variant2 || ''))
    ].find(result => result !== 0) || 0;
  }).map(group => {
    const variants = getCatalogStainlessSteelFlangeVariantColumns_(group.rows, col);
    const pivotRows = buildCatalogStainlessSteelFlangeTableRows_(group.rows, col, variants);
    const cartonLayout = buildCatalogCartonLayoutFlags_(pivotRows, variants, group.rows[0] && group.rows[0][col.Product_Group]);
    const rowCount = pivotRows.length;
    const recommendation = chooseCatalogSectionTemplate_(rowCount);

    return {
      productGroup: group.rows[0] && group.rows[0][col.Product_Group],
      plc: group.plcValues.join(','),
      fittingType: group.fittingType,
      variant2: group.variant2,
      title: buildCatalogSectionDisplayTitle_(group.fittingType, group.variant2, group.rows[0] && group.rows[0][col.Product_Group]),
      skuRows: group.rows.length,
      renderedRows: rowCount,
      tableRows: pivotRows,
      rawVariants: variants,
      variants: variants.map(variant => getDisplayVariantName_(variant)),
      sizeHeader: getCatalogSectionSizeHeader_(group.fittingType),
      imageFileId: getFirstCatalogImageFileId_(group.rows, col),
      recommendedTemplate: recommendation.template,
      reason: recommendation.reason,
      hideInnerCarton: cartonLayout.hideInnerCarton,
      hideMasterCase: cartonLayout.hideMasterCase,
      variantSpecificInnerCarton: cartonLayout.variantSpecificInnerCarton,
      variantSpecificMasterCase: cartonLayout.variantSpecificMasterCase,
      compactTable: cartonLayout.compactTable,
      pageSubtitle: group.variant2,
      forceShapeTable: cartonLayout.variantSpecificInnerCarton || cartonLayout.variantSpecificMasterCase
    };
  });
}

function getCatalogStainlessSteelFlangeVariantColumns_(rows, col) {
  const variants = [];

  rows.forEach(row => {
    const variant = normalizeVariant2Key_(getOptionalCellValue_(row, col, 'Variant_2'));
    if (variant && !variants.includes(variant)) variants.push(variant);
  });

  return variants.sort((a, b) => variantSortWeight_(a) - variantSortWeight_(b));
}

function buildCatalogStainlessSteelFlangeTableRows_(rows, col, variants) {
  const rowMap = {};

  rows.forEach(row => {
    const size = getCatalogRowSizeParts_(row, col).join(' x ');
    const key = normalizeCatalogSizeForComparison_(size) || String(size || '').trim();
    if (!rowMap[key]) {
      rowMap[key] = {
        size,
        innerCarton: row[col.Inner_Carton],
        masterCase: row[col.Master_Case],
        variants: {}
      };
    }

    const variant = normalizeVariant2Key_(getOptionalCellValue_(row, col, 'Variant_2'));
    rowMap[key].variants[variant] = {
      itemNumber: row[col.Item_Number],
      price: getCatalogRowPriceValue_(row, col),
      innerCarton: row[col.Inner_Carton],
      masterCase: row[col.Master_Case]
    };
  });

  return Object.values(rowMap).sort((a, b) => {
    const sizePartsA = String(a.size || '').split(/\s*x\s*/i);
    const sizePartsB = String(b.size || '').split(/\s*x\s*/i);
    return [
      parseCatalogSizeSortValue_(sizePartsA[0]) - parseCatalogSizeSortValue_(sizePartsB[0]),
      parseCatalogSizeSortValue_(sizePartsA[1]) - parseCatalogSizeSortValue_(sizePartsB[1]),
      String(a.size).localeCompare(String(b.size), undefined, { numeric: true })
    ].find(result => result !== 0) || 0;
  });
}

function buildCatalogValveSectionModels_(rows, col) {
  const groupMap = {};

  rows.forEach(row => {
    const fittingType = String(row[col.Fitting_Type] || 'Valves').trim() || 'Valves';
    const materialFamily = getValveMaterialFamily_(row, col);
    const materialDetail = getValveMaterialDetail_(row, col);
    const connectionType = getValveConnectionType_(row, col);
    const itemSeries = getValveItemSeries_(row, col);
    const key = `${fittingType}|${materialFamily}|${connectionType}|${itemSeries}`;

    if (!groupMap[key]) {
      groupMap[key] = {
        fittingType,
        materialFamily,
        materialDetail,
        connectionType,
        itemSeries,
        rows: []
      };
    }

    groupMap[key].rows.push(row);
  });

  return Object.values(groupMap).sort((a, b) => {
    return [
      getCatalogFittingTypeSortWeight_('Valves', a.fittingType) -
        getCatalogFittingTypeSortWeight_('Valves', b.fittingType),
      getValveMaterialSortWeight_(a.materialFamily) - getValveMaterialSortWeight_(b.materialFamily),
      String(a.materialDetail || '').localeCompare(String(b.materialDetail || '')),
      String(a.connectionType || '').localeCompare(String(b.connectionType || '')),
      String(a.itemSeries || '').localeCompare(String(b.itemSeries || ''))
    ].find(result => result !== 0) || 0;
  }).flatMap(group => {
    const variants = getCatalogVariantColumns_(group.rows, col);
    const variantChunks = variants.length > 2 ? chunkCatalogArray_(variants, 2) : [variants];
    const groupPlcs = [...new Set(group.rows.map(row => normalizeMatchValue_(row[col.PLC])))]
      .filter(Boolean)
      .join(',');

    return variantChunks.map(variantChunk => {
      const chunkRows = variantChunk.length
        ? group.rows.filter(row => variantChunk.includes(normalizeVariantKey_(row[col.Variant])))
        : group.rows.slice();
      const pivotRows = buildCatalogVariantTableRows_(chunkRows, col, {
        pivotBySizeOnly: shouldPivotCatalogRowsBySizeOnly_(chunkRows, col, variantChunk)
      });
      const cartonLayout = buildCatalogCartonLayoutFlags_(pivotRows, variantChunk, chunkRows[0] && chunkRows[0][col.Product_Group]);
      const rowCount = pivotRows.length;
      const recommendation = chooseCatalogSectionTemplate_(rowCount);
      const sectionQualifier = buildValveSectionQualifier_(group);
      const sectionTitle = buildCatalogSectionDisplayTitle_(
        group.fittingType,
        sectionQualifier,
        group.rows[0] && group.rows[0][col.Product_Group]
      );
      return {
        productGroup: chunkRows[0] && chunkRows[0][col.Product_Group],
        plc: groupPlcs,
        fittingType: group.fittingType,
        variant2: sectionQualifier || group.materialFamily,
        title: sectionTitle,
        specInfo: getValveSectionSpecInfo_(chunkRows, col),
        skuRows: chunkRows.length,
        renderedRows: rowCount,
        tableRows: pivotRows,
        rawVariants: variantChunk,
        variants: variantChunk.map(variant => getDisplayVariantName_(variant)),
        sizeHeader: getCatalogSectionSizeHeader_(group.fittingType),
        imageFileId: getFirstCatalogImageFileId_(chunkRows, col),
        recommendedTemplate: recommendation.template,
        reason: recommendation.reason,
        pageSubtitle: group.materialFamily,
        allowMixedPageSubtitlePacking: true,
        hideInnerCarton: cartonLayout.hideInnerCarton,
        hideMasterCase: cartonLayout.hideMasterCase,
        variantSpecificInnerCarton: cartonLayout.variantSpecificInnerCarton,
        variantSpecificMasterCase: cartonLayout.variantSpecificMasterCase,
        compactTable: cartonLayout.compactTable,
        forceShapeTable: variantChunk.length > 1 || cartonLayout.variantSpecificInnerCarton || cartonLayout.variantSpecificMasterCase
      };
    });
  });
}

function getValveSectionSpecInfo_(rows, col) {
  for (let i = 0; i < (rows || []).length; i++) {
    const value = String(getOptionalCellValue_(rows[i], col, 'Valve_Spec_Info') || '').trim();
    if (value) return value;
  }
  return '';
}

function chunkCatalogArray_(values, chunkSize) {
  const chunks = [];
  for (let i = 0; i < values.length; i += chunkSize) {
    chunks.push(values.slice(i, i + chunkSize));
  }
  return chunks;
}

function buildCatalogNippleSectionModels_(rows, col) {
  const groupMap = {};

  rows.forEach(row => {
    const fittingType = String(row[col.Fitting_Type] || 'Nipples').trim() || 'Nipples';
    const nominalSize = normalizeNippleNominalSize_(row[col.Size_1]);
    const scheduleGroup = getCatalogNippleScheduleGroup_(row, col);
    const key = `${scheduleGroup}|${fittingType}|${nominalSize}`;

    if (!groupMap[key]) {
      groupMap[key] = {
        scheduleGroup,
        fittingType,
        nominalSize,
        nominalSizeSort: parseCatalogSizeSortValue_(row[col.Size_1]),
        rows: []
      };
    }

    groupMap[key].rows.push(row);
  });

  return Object.values(groupMap).sort((a, b) => {
    return [
      getCatalogNippleScheduleSortWeight_(a.scheduleGroup) - getCatalogNippleScheduleSortWeight_(b.scheduleGroup),
      getCatalogFittingTypeSortWeight_(a.rows[0] && a.rows[0][col.Product_Group], a.fittingType) -
        getCatalogFittingTypeSortWeight_(b.rows[0] && b.rows[0][col.Product_Group], b.fittingType),
      a.nominalSizeSort - b.nominalSizeSort,
      String(a.nominalSize).localeCompare(String(b.nominalSize), undefined, { numeric: true })
    ].find(result => result !== 0) || 0;
  }).map(group => {
    const pivotRows = buildCatalogNippleTableRows_(group.rows, col);
    const variants = getCatalogNippleVariantColumns_(group.rows, col);
    const rowCount = pivotRows.length;
    const recommendation = chooseCatalogSectionTemplate_(rowCount);

    return {
      productGroup: group.rows[0] && group.rows[0][col.Product_Group],
      plc: '',
      fittingType: group.fittingType,
      variant2: '',
      title: `${group.nominalSize} ${group.fittingType}`,
      skuRows: group.rows.length,
      renderedRows: rowCount,
      tableRows: pivotRows,
      rawVariants: variants,
      variants: variants.map(variant => getDisplayVariantName_(variant)),
      imageFileId: getFirstCatalogImageFileId_(group.rows, col),
      recommendedTemplate: recommendation.template,
      reason: variants.length <= 1
        ? 'Single-variant nipple section; allow up to three nominal pipe sizes per page within the same schedule group.'
        : 'Multi-variant nipple section; allow up to two nominal pipe sizes per page within the same schedule group.',
      sizeHeader: 'Length',
      hideInnerCarton: group.rows.every(row => !String(row[col.Inner_Carton] || '').trim()),
      hideMasterCase: group.rows.every(row => !String(row[col.Master_Case] || '').trim()),
      maxSectionsPerPage: variants.length <= 1 ? 3 : 2,
      isNippleTable: true,
      pageSubtitle: getCatalogNipplePageSubtitle_(group.scheduleGroup)
    };
  });
}

function buildCatalogNippleTableRows_(rows, col) {
  const rowMap = {};

  rows.forEach(row => {
    const length = buildCatalogNippleLength_(row, col);
    const key = length;

    if (!rowMap[key]) {
      rowMap[key] = {
        size: length,
        innerCarton: row[col.Inner_Carton],
        masterCase: row[col.Master_Case],
        variants: {},
        size1Sort: parseCatalogSizeSortValue_(row[col.Size_2]),
        size2Sort: parseCatalogSizeSortValue_(row[col.Size_3]),
        size3Sort: 0
      };
    } else {
      if (!String(rowMap[key].innerCarton || '').trim() && String(row[col.Inner_Carton] || '').trim()) {
        rowMap[key].innerCarton = row[col.Inner_Carton];
      }
      if (!String(rowMap[key].masterCase || '').trim() && String(row[col.Master_Case] || '').trim()) {
        rowMap[key].masterCase = row[col.Master_Case];
      }
    }

    const variant = buildCatalogNippleVariantKey_(row, col);
    rowMap[key].variants[variant] = {
      itemNumber: row[col.Item_Number],
      price: getCatalogRowPriceValue_(row, col)
    };
  });

  return Object.values(rowMap).sort((a, b) => {
    return [
      a.size1Sort - b.size1Sort,
      a.size2Sort - b.size2Sort,
      String(a.size).localeCompare(String(b.size), undefined, { numeric: true })
    ].find(result => result !== 0) || 0;
  });
}

function getCatalogNippleVariantColumns_(rows, col) {
  const variants = [];

  rows.forEach(row => {
    const variant = buildCatalogNippleVariantKey_(row, col);
    if (variant && !variants.includes(variant)) variants.push(variant);
  });

  return variants.sort((a, b) => variantSortWeight_(a) - variantSortWeight_(b));
}

function buildCatalogNippleLength_(row, col) {
  return [row[col.Size_2], row[col.Size_3]]
    .filter(value => value !== '' && value !== null && value !== undefined)
    .join(' x ');
}

function buildCatalogNippleVariantKey_(row, col) {
  const productGroup = String(row[col.Product_Group] || '').toLowerCase();
  const plc = normalizeMatchValue_(row[col.PLC]);
  const variant = normalizeVariantKey_(row[col.Variant]);
  const variant2 = normalizeVariant2Key_(getOptionalCellValue_(row, col, 'Variant_2'));

  if (productGroup.indexOf('carbon steel nipple') !== -1) {
    const schedule = getCarbonSteelNippleScheduleLabel_(plc);
    if (schedule === 'A106GrB' && (!variant || variant.toLowerCase() === 'black')) return schedule;
    return getDisplayVariantName_(variant);
  }

  if (productGroup.indexOf('stainless steel nipple') !== -1) {
    return getDisplayVariantName_(variant2);
  }

  return [variant2, variant].filter(Boolean).join(' ') || variant || variant2 || plc || 'Item';
}

function getCarbonSteelNippleScheduleLabel_(plc) {
  if (plc === '727') return 'S40';
  if (plc === '729') return 'S80';
  if (plc === '724') return 'A106GrB';
  return plc ? `PLC ${plc}` : '';
}

function getCatalogNippleScheduleGroup_(row, col) {
  const productGroup = String(row[col.Product_Group] || '').toLowerCase();

  if (productGroup.indexOf('carbon steel nipple') !== -1) {
    return getCarbonSteelNippleScheduleLabel_(normalizeMatchValue_(row[col.PLC]));
  }

  if (productGroup.indexOf('stainless steel nipple') !== -1) {
    return getStainlessSteelNippleScheduleLabel_(normalizeVariantKey_(row[col.Variant]));
  }

  return normalizeVariant2Key_(getOptionalCellValue_(row, col, 'Variant_2')) || '';
}

function getStainlessSteelNippleScheduleLabel_(variant) {
  const normalized = String(variant || '').toLowerCase();
  if (normalized.indexOf('40') !== -1) return 'S40';
  if (normalized.indexOf('80') !== -1) return 'S80';
  return String(variant || '').trim();
}

function getCatalogNippleScheduleSortWeight_(scheduleGroup) {
  const normalized = String(scheduleGroup || '').toLowerCase();
  if (normalized === 's40') return 10;
  if (normalized === 's80') return 20;
  if (normalized === 'a106grb') return 30;
  if (normalized === '304') return 10;
  if (normalized === '316') return 20;
  return normalized ? 100 : 0;
}

function getCatalogNipplePageSubtitle_(scheduleGroup) {
  const normalized = String(scheduleGroup || '').toLowerCase();
  if (normalized === 's40') return 'Schedule 40 Welded Nipples';
  if (normalized === 's80') return 'Schedule 80 Seamless Nipples';
  return '';
}

function normalizeNippleNominalSize_(size) {
  return String(size || '').trim() || 'Nipples';
}

function isNippleCatalogProductGroup_(productGroup) {
  return String(productGroup || '').toLowerCase().indexOf('nipple') !== -1;
}

function chooseCatalogSectionTemplate_(renderedRows) {
  if (renderedRows <= 6) {
    return {
      template: '02_Product_Page_Three_Sections',
      reason: 'Small section; candidate for a three-section page.'
    };
  }

  if (renderedRows <= 13) {
    return {
      template: '03_Product_Page_Two_Sections',
      reason: 'Medium section; candidate for a two-section page.'
    };
  }

  if (renderedRows <= getCatalogSectionPageRowLimit_()) {
    return {
      template: '04_Product_Page_One_Section',
      reason: 'Large section; use a full product page.'
    };
  }

  return {
    template: '05_Product_Page_Continuation',
    reason: 'Oversized section; split across a one-section page plus continuation pages.'
  };
}

function buildCatalogSlidePages_(sections) {
  const pages = [];
  const sectionLanes = splitCatalogSectionsIntoPackingLanes_(sections);

  sectionLanes.forEach(laneSections => {
    const remainingSections = laneSections.slice();
    let buffer = [];

    while (remainingSections.length) {
      const section = remainingSections.shift();
      const sectionPageRowLimit = getCatalogSectionPageRowLimit_(section);

      if (section.forceSingleSectionPage) {
        flushCatalogSectionPageBuffer_(buffer, pages);
        buffer = [];

        if (section.renderedRows > sectionPageRowLimit) {
          appendOversizedCatalogSectionPages_(section, remainingSections, pages);
        } else {
          pages.push(buildCatalogSlidePage_('04_Product_Page_One_Section', [section]));
        }
        continue;
      }

      if (section.renderedRows > sectionPageRowLimit) {
        flushCatalogSectionPageBuffer_(buffer, pages);
        buffer = [];
        appendOversizedCatalogSectionPages_(section, remainingSections, pages);
        continue;
      }

      buffer.push(section);
    }

    flushCatalogSectionPageBuffer_(buffer, pages);
  });

  optimizeCatalogPages_(pages);

  return pages.map((page, index) => {
    page.pageNumber = index + 1;
    return page;
  });
}

function optimizeCatalogPages_(pages) {
  let changed = true;

  while (changed) {
    changed = false;

    for (let pageIndex = 1; pageIndex < pages.length; pageIndex++) {
      const page = pages[pageIndex];
      if (!page || page.sections.length !== 1) continue;
      if (page.sections.some(section => shouldPreserveStrictCatalogSectionOrder_(section))) continue;

      const section = page.sections[0];
      if (!section || section.continued) continue;
      if (Number(section.renderedRows || 0) > 4) continue;

      for (let targetIndex = 0; targetIndex < pageIndex; targetIndex++) {
        const targetPage = pages[targetIndex];
        if (!targetPage || targetPage.sections.length >= 3) continue;
        if (targetPage.sections.some(targetSection => shouldPreserveStrictCatalogSectionOrder_(targetSection))) continue;

        const combinedSections = targetPage.sections.concat([section]);
        const layout = getCatalogPackedPageLayout_(combinedSections);
        if (!layout) continue;

        targetPage.sections = combinedSections;
        targetPage.template = layout.template;
        targetPage.sectionCount = combinedSections.length;
        pages.splice(pageIndex, 1);
        changed = true;
        break;
      }

      if (changed) break;
    }
  }
}

function shouldPreserveStrictCatalogSectionOrder_(section) {
  if (!section) return false;
  return isButtWeldCatalogProductGroup_(section.productGroup) ||
    isForgedSteelCatalogProductGroup_(section.productGroup) ||
    isBronzeFittingsCatalogProductGroup_(section.productGroup) ||
    isMalleableIronCatalogProductGroup_(section.productGroup);
}

function splitCatalogSectionsIntoPackingLanes_(sections) {
  if (!sections || !sections.length) return [];

  const lanes = [];
  let currentLane = [];
  let currentLaneKey = null;

  sections.forEach(section => {
    const laneKey = getCatalogSectionPackingLaneKey_(section);
    if (!currentLane.length || laneKey === currentLaneKey) {
      currentLane.push(section);
      currentLaneKey = laneKey;
      return;
    }

    lanes.push(currentLane);
    currentLane = [section];
    currentLaneKey = laneKey;
  });

  if (currentLane.length) lanes.push(currentLane);
  return lanes;
}

function getCatalogSectionPackingLaneKey_(section) {
  if (section && section.isNippleTable) {
    return `nipple:${String(section.pageSubtitle || '').trim() || '__default__'}`;
  }

  const variant2 = normalizeVariant2Key_(section && section.variant2).toLowerCase();
  if (variant2 === '150lb' || variant2 === '150 lb') return '150lb';
  if (variant2 === '300lb' || variant2 === '300 lb') return '300lb';
  if (variant2 === '3000lb' || variant2 === '3000 lb') return '3000lb';
  if (variant2 === '2000lb' || variant2 === '2000 lb') return '2000lb';
  if (variant2 === 'threaded') return 'threaded';
  if (variant2 === 'socket-weld' || variant2 === 'socket weld' || variant2 === 'socket') return 'socket-weld';
  if (variant2.indexOf('std') !== -1 || variant2.indexOf('standard weight') !== -1) return 'std';
  if (variant2.indexOf('lw') !== -1 || variant2.indexOf('light weight') !== -1) return 'lw';
  if (variant2.indexOf('xh') !== -1 || variant2.indexOf('extra heavy') !== -1) return 'xh';
  return '__default__';
}

function isBronzeFittingsCatalogProductGroup_(productGroup) {
  const normalized = String(productGroup || '').toLowerCase();
  return normalized.indexOf('bronze fittings') !== -1;
}

function isMalleableIronCatalogProductGroup_(productGroup) {
  const normalized = String(productGroup || '').toLowerCase();
  return normalized.indexOf('malleable iron fittings') !== -1;
}

function flushCatalogSectionPageBuffer_(buffer, pages) {
  if (!buffer.length) return;

  packCatalogSectionsIntoPages_(buffer).forEach(page => pages.push(page));
  buffer.splice(0, buffer.length);
}

function packCatalogSectionsIntoPages_(sections) {
  const remaining = sections.slice();
  const pages = [];

  while (remaining.length) {
    const choice = chooseCatalogPageFromLookahead_(remaining);
    const layout = getCatalogPackedPageLayout_(choice.sections);

    if (!layout) {
      const fallbackSection = remaining[0];
      const fallbackLayout = getCatalogPackedPageLayout_([fallbackSection]) ||
        { template: '04_Product_Page_One_Section', capacity: getCatalogSectionPageRowLimit_(fallbackSection) };
      pages.push(buildCatalogSlidePage_(fallbackLayout.template, [fallbackSection]));
      remaining.splice(0, 1);
      continue;
    }

    pages.push(buildCatalogSlidePage_(layout.template, choice.sections));
    choice.indices
      .slice()
      .sort((a, b) => b - a)
      .forEach(index => remaining.splice(index, 1));
  }

  return pages;
}

function getCatalogPackedPageLayout_(sections) {
  if (!areCatalogSectionsPageCompatible_(sections)) return null;

  const sectionCount = sections.length;
  const rowTotal = getCatalogSectionsRenderedRowTotal_(sections);
  const maxSectionRows = Math.max.apply(null, sections.map(section => section.renderedRows));
  const maxAllowedSections = Math.min.apply(null, sections.map(section => Number(section.maxSectionsPerPage || 3)));
  const hasShortContinuedSection = sectionCount === 2 && sections.some(section => {
    const renderedRows = Number(section.renderedRowsOnPage || section.renderedRows || 0);
    return section.continued && renderedRows < (getCatalogSectionPageRowLimit_(section) / 2);
  });

  if (sectionCount === 3 && maxAllowedSections >= 3 && rowTotal <= 36 && maxSectionRows <= 14) {
    return { template: '02_Product_Page_Three_Sections', capacity: 36 };
  }
  if (sectionCount === 2 && maxAllowedSections >= 2 && rowTotal <= (hasShortContinuedSection ? 46 : 42) && maxSectionRows <= (hasShortContinuedSection ? 30 : 26)) {
    return { template: '03_Product_Page_Two_Sections', capacity: hasShortContinuedSection ? 46 : 42 };
  }
  if (sectionCount === 1 && rowTotal <= getCatalogSectionPageRowLimit_(sections[0])) {
    return { template: '04_Product_Page_One_Section', capacity: getCatalogSectionPageRowLimit_(sections[0]) };
  }

  return null;
}

function chooseCatalogPageFromLookahead_(sections) {
  const leadingSection = sections[0];
  const windowSize = Math.min(
    leadingSection && leadingSection.isNippleTable ? 8 : 32,
    sections.length
  );
  const candidates = [];
  const requireFirstSection = shouldRequireCatalogLeadingSection_(leadingSection);
  const requireContiguousFromStart = requireFirstSection &&
    (isButtWeldCatalogProductGroup_(leadingSection && leadingSection.productGroup) ||
      isForgedSteelCatalogProductGroup_(leadingSection && leadingSection.productGroup) ||
      isBronzeFittingsCatalogProductGroup_(leadingSection && leadingSection.productGroup) ||
      isMalleableIronCatalogProductGroup_(leadingSection && leadingSection.productGroup));

  for (let count = 3; count >= 1; count--) {
    const combos = getCatalogIndexCombinations_(windowSize, count);

    combos.forEach(indices => {
      if (requireFirstSection && indices.indexOf(0) === -1) return;
      if (requireContiguousFromStart && !areCatalogIndicesContiguousFromStart_(indices)) return;
      if (leadingSection && leadingSection.isNippleTable && !areCatalogIndicesContiguousFromStart_(indices)) return;
      if (leadingSection && leadingSection.continued && !leadingSection.isFinalContinuation) {
        if (indices.length !== 1 || indices[0] !== 0) return;
      }

      const pageSections = indices.map(index => sections[index]);
      const layout = getCatalogPackedPageLayout_(pageSections);
      if (!layout) return;

      const totalRows = getCatalogSectionsRenderedRowTotal_(pageSections);
      const slack = layout.capacity - totalRows;
      const includesFirstSection = indices.indexOf(0) !== -1;
      candidates.push({
        indices,
        sections: pageSections,
        layout,
        score: (pageSections.length * 2000) +
          (totalRows * 1000) -
          (slack * 24) -
          (indices.reduce((sum, index) => sum + index, 0) * 20) +
          (includesFirstSection ? 1500 : 0)
      });
    });
  }

  if (!candidates.length) {
    return { indices: [0], sections: [sections[0]] };
  }

  candidates.sort((a, b) => b.score - a.score);
  return candidates[0];
}

function areCatalogIndicesContiguousFromStart_(indices) {
  if (!indices || !indices.length) return false;
  for (let i = 0; i < indices.length; i++) {
    if (indices[i] !== i) return false;
  }
  return true;
}

function shouldRequireCatalogLeadingSection_(section) {
  if (!section) return false;
  if (section.continued) return true;
  if (section.isNippleTable) return true;
  const variant2 = normalizeVariant2Key_(section.variant2);
  const fittingWeight = getCatalogSectionFittingSortWeight_(section.productGroup, section.fittingType, section.variant2);

  if (String(section.productGroup || '').toLowerCase().indexOf('valve') !== -1) {
    return fittingWeight <= 2;
  }

  if (variant2 && /^(150 ?lb|300 ?lb|2000 ?lb|3000 ?lb|threaded|socket weld|socket-weld|304|316)$/i.test(variant2)) {
    return fittingWeight <= 3;
  }

  return fittingWeight <= 3;
}

function getCatalogIndexCombinations_(size, count) {
  const results = [];

  function build(startIndex, remaining, current) {
    if (!remaining) {
      results.push(current.slice());
      return;
    }

    for (let index = startIndex; index <= size - remaining; index++) {
      current.push(index);
      build(index + 1, remaining - 1, current);
      current.pop();
    }
  }

  build(0, count, []);
  return results;
}

function getCatalogSectionsRenderedRowTotal_(sections) {
  return sections.reduce((sum, section) => sum + Number(section.renderedRows || 0), 0);
}

function areCatalogSectionsPageCompatible_(sections) {
  const pageSubtitles = sections
    .map(section => String(section.pageSubtitle || '').trim())
    .filter(Boolean)
    .filter((value, index, values) => values.indexOf(value) === index);
  if (pageSubtitles.length) {
    if (pageSubtitles.length <= 1) return true;

    const canMixPageSubtitles = sections.every(section => !!section.allowMixedPageSubtitlePacking);
    const smallestSectionRows = Math.min.apply(null, sections.map(section => Number(section.renderedRows || 0)));
    const totalRows = getCatalogSectionsRenderedRowTotal_(sections);

    return canMixPageSubtitles &&
      sections.length <= 3 &&
      smallestSectionRows <= 4 &&
      totalRows <= 21;
  }

  const variant2Values = sections
    .map(section => normalizeVariant2Key_(section.variant2))
    .filter(Boolean)
    .filter((value, index, values) => values.indexOf(value) === index);

  return variant2Values.length <= 1;
}

function appendOversizedCatalogSectionPages_(section, remainingSections, pages) {
  const pageRowLimit = getCatalogSectionPageRowLimit_(section);
  let remainingRows = Number(section.renderedRows || 0);
  let part = 1;

  const firstPageSection = copyCatalogSectionForPage_(section, part, Math.min(pageRowLimit, remainingRows), false);
  pages.push(buildCatalogSlidePage_('04_Product_Page_One_Section', [firstPageSection]));
  remainingRows -= pageRowLimit;
  part++;

  const continuationSections = [];
  while (remainingRows > 0) {
    const renderedRowsOnPage = Math.min(pageRowLimit, remainingRows);
    continuationSections.push(copyCatalogSectionForPage_(section, part, renderedRowsOnPage, true, remainingRows <= pageRowLimit));
    remainingRows -= pageRowLimit;
    part++;
  }

  if (continuationSections.length) {
    remainingSections.unshift.apply(remainingSections, continuationSections);
  }
}

function takeCatalogContinuationCompanionSection_(continuedSection, remainingSections) {
  const windowSize = Math.min(8, remainingSections.length);
  let bestMatch = null;

  for (let index = 0; index < windowSize; index++) {
    const candidate = remainingSections[index];
    if (!candidate) continue;
    if (candidate.forceSingleSectionPage) continue;
    if (candidate.renderedRows > getCatalogSectionPageRowLimit_(candidate)) continue;

    const layout = getCatalogPackedPageLayout_([continuedSection, candidate]);
    if (!layout || layout.template !== '03_Product_Page_Two_Sections') continue;

    const totalRows = getCatalogSectionsRenderedRowTotal_([continuedSection, candidate]);
    const slack = layout.capacity - totalRows;
    const score = (totalRows * 1000) - (slack * 12) - (index * 10);

    if (!bestMatch || score > bestMatch.score) {
      bestMatch = {
        index,
        score
      };
    }
  }

  if (!bestMatch) return null;
  return remainingSections.splice(bestMatch.index, 1)[0];
}

function copyCatalogSectionForPage_(section, part, renderedRowsOnPage, continued, isFinalContinuation) {
  const copy = Object.assign({}, section);
  const pageRowLimit = getCatalogSectionPageRowLimit_(section);
  const startIndex = (part - 1) * pageRowLimit;
  copy.part = part;
  copy.originalRenderedRows = section.renderedRows;
  copy.originalTableRows = section.tableRows;
  copy.renderedRows = renderedRowsOnPage;
  copy.renderedRowsOnPage = renderedRowsOnPage;
  copy.tableRows = (section.tableRows || []).slice(startIndex, startIndex + renderedRowsOnPage);
  copy.continued = continued;
  copy.isFinalContinuation = !!isFinalContinuation;
  copy.continuedLabel = continued ? '(continued)' : '';
  return copy;
}

function getCatalogSectionPageRowLimit_(section) {
  return 56;
}

function buildCatalogSlidePage_(templateName, sections) {
  return {
    template: templateName,
    sectionCount: sections.length,
    sections: sections.map(section => ({
      productGroup: section.productGroup || '',
      plc: section.plc,
      fittingType: section.fittingType,
      variant2: section.variant2 || '',
      title: section.title,
      skuRows: section.skuRows,
      renderedRows: section.renderedRows,
      renderedRowsOnPage: section.renderedRowsOnPage || section.renderedRows,
      tableRows: section.tableRows || [],
      rawVariants: section.rawVariants || [],
      sizeHeader: section.sizeHeader || 'Size',
      specInfo: section.specInfo || '',
      hideInnerCarton: !!section.hideInnerCarton,
      hideMasterCase: !!section.hideMasterCase,
      compactTable: !!section.compactTable,
      isNippleTable: !!section.isNippleTable,
      pageSubtitle: section.pageSubtitle || '',
      forceSingleSectionPage: !!section.forceSingleSectionPage,
      part: section.part || 1,
      continued: !!section.continued,
      continuedLabel: section.continuedLabel || '',
      variants: section.variants,
      imageFileId: section.imageFileId,
      hasImage: !!section.imageFileId
    }))
  };
}

function summarizeCatalogPageMix_(pages) {
  return pages.reduce((summary, page) => {
    summary[page.template] = (summary[page.template] || 0) + 1;
    return summary;
  }, {});
}

function formatCatalogPageMix_(pageMix) {
  return Object.keys(pageMix)
    .sort()
    .map(templateName => `${templateName}: ${pageMix[templateName]}`)
    .join(', ');
}

function buildCatalogSlidePlanLogSummary_(plan) {
  return plan.map(groupPlan => {
    const oversizedSections = groupPlan.sections
      .filter(section => section.renderedRows > getCatalogSectionPageRowLimit_(section))
      .map(section => ({
        plc: section.plc,
        fittingType: section.fittingType,
        renderedRows: section.renderedRows,
        skuRows: section.skuRows,
        recommendedTemplate: section.recommendedTemplate
      }));

    const missingImages = groupPlan.sections
      .filter(section => !section.imageFileId)
      .map(section => ({
        plc: section.plc,
        fittingType: section.fittingType,
        renderedRows: section.renderedRows
      }));

    return {
      summary: groupPlan.summary,
      oversizedSections,
      missingImages
    };
  });
}

function comparePriceFileRows_(a, b, col) {
  const productGroupA = String(a[col.Product_Group] || '').trim();
  const productGroupB = String(b[col.Product_Group] || '').trim();

  if (productGroupA === productGroupB && isNippleCatalogProductGroup_(productGroupA)) {
    const sortA = Number(a[col.Sort_Order]) || 0;
    const sortB = Number(b[col.Sort_Order]) || 0;

    return [
      getCatalogNippleScheduleSortWeight_(getCatalogNippleScheduleGroup_(a, col)) -
        getCatalogNippleScheduleSortWeight_(getCatalogNippleScheduleGroup_(b, col)),
      getCatalogFittingTypeSortWeight_(productGroupA, a[col.Fitting_Type]) -
        getCatalogFittingTypeSortWeight_(productGroupB, b[col.Fitting_Type]),
      String(a[col.Fitting_Type] || '').localeCompare(String(b[col.Fitting_Type] || '')),
      parseCatalogSizeSortValue_(a[col.Size_1]) - parseCatalogSizeSortValue_(b[col.Size_1]),
      variantSortWeight_(buildCatalogNippleVariantKey_(a, col)) - variantSortWeight_(buildCatalogNippleVariantKey_(b, col)),
      parseCatalogSizeSortValue_(a[col.Size_2]) - parseCatalogSizeSortValue_(b[col.Size_2]),
      parseCatalogSizeSortValue_(a[col.Size_3]) - parseCatalogSizeSortValue_(b[col.Size_3]),
      sortA - sortB,
      String(a[col.Item_Number] || '').localeCompare(String(b[col.Item_Number] || ''))
    ].find(result => result !== 0) || 0;
  }

  if (productGroupA === productGroupB && isValveCatalogProductGroup_(productGroupA)) {
    const sortA = Number(a[col.Sort_Order]) || 0;
    const sortB = Number(b[col.Sort_Order]) || 0;
    const familyA = getValveMaterialFamily_(a, col);
    const familyB = getValveMaterialFamily_(b, col);
    const detailA = getValveMaterialDetail_(a, col);
    const detailB = getValveMaterialDetail_(b, col);

    return [
      normalizeMatchValue_(a[col.PLC]).localeCompare(normalizeMatchValue_(b[col.PLC]), undefined, { numeric: true }),
      getCatalogFittingTypeSortWeight_(productGroupA, a[col.Fitting_Type]) -
        getCatalogFittingTypeSortWeight_(productGroupB, b[col.Fitting_Type]),
      String(a[col.Fitting_Type] || '').localeCompare(String(b[col.Fitting_Type] || '')),
      getValveMaterialSortWeight_(familyA) - getValveMaterialSortWeight_(familyB),
      familyA.localeCompare(familyB),
      detailA.localeCompare(detailB),
      normalizeVariant2Key_(getOptionalCellValue_(a, col, 'Variant_2')).localeCompare(normalizeVariant2Key_(getOptionalCellValue_(b, col, 'Variant_2'))),
      parseCatalogSizeSortValue_(a[col.Size_1]) - parseCatalogSizeSortValue_(b[col.Size_1]),
      parseCatalogSizeSortValue_(a[col.Size_2]) - parseCatalogSizeSortValue_(b[col.Size_2]),
      parseCatalogSizeSortValue_(a[col.Size_3]) - parseCatalogSizeSortValue_(b[col.Size_3]),
      sortA - sortB,
      String(a[col.Item_Number] || '').localeCompare(String(b[col.Item_Number] || ''))
    ].find(result => result !== 0) || 0;
  }

  if (productGroupA === productGroupB) {
    const variantA = normalizeVariantKey_(a[col.Variant]);
    const variantB = normalizeVariantKey_(b[col.Variant]);
    const variant2A = normalizeVariant2Key_(getOptionalCellValue_(a, col, 'Variant_2'));
    const variant2B = normalizeVariant2Key_(getOptionalCellValue_(b, col, 'Variant_2'));
    const sortA = Number(a[col.Sort_Order]) || 0;
    const sortB = Number(b[col.Sort_Order]) || 0;
    const sizePartsA = getCatalogRowSizeParts_(a, col);
    const sizePartsB = getCatalogRowSizeParts_(b, col);

    return [
      normalizeMatchValue_(a[col.PLC]).localeCompare(normalizeMatchValue_(b[col.PLC]), undefined, { numeric: true }),
      getCatalogVariant2SortWeight_(variant2A) - getCatalogVariant2SortWeight_(variant2B),
      variant2A.localeCompare(variant2B, undefined, { numeric: true }),
      getCatalogFittingTypeSortWeight_(productGroupA, a[col.Fitting_Type]) -
        getCatalogFittingTypeSortWeight_(productGroupB, b[col.Fitting_Type]),
      String(a[col.Fitting_Type] || '').localeCompare(String(b[col.Fitting_Type] || '')),
      variantSortWeight_(variantA) - variantSortWeight_(variantB),
      variantA.localeCompare(variantB),
      parseCatalogSizeSortValue_(sizePartsA[0]) - parseCatalogSizeSortValue_(sizePartsB[0]),
      parseCatalogSizeSortValue_(sizePartsA[1]) - parseCatalogSizeSortValue_(sizePartsB[1]),
      parseCatalogSizeSortValue_(sizePartsA[2]) - parseCatalogSizeSortValue_(sizePartsB[2]),
      sortA - sortB,
      String(a[col.Item_Number] || '').localeCompare(String(b[col.Item_Number] || ''))
    ].find(result => result !== 0) || 0;
  }

  return productGroupA.localeCompare(productGroupB) || compareCatalogSkuRows_(a, b, col);
}

function compareCatalogSkuRows_(a, b, col) {
  const sortA = Number(a[col.Sort_Order]) || 0;
  const sortB = Number(b[col.Sort_Order]) || 0;
  if (sortA !== sortB) return sortA - sortB;

  const productGroupA = String(a[col.Product_Group] || '').trim();
  const productGroupB = String(b[col.Product_Group] || '').trim();
  if (productGroupA === productGroupB && isValveCatalogProductGroup_(productGroupA)) {
    const familyA = getValveMaterialFamily_(a, col);
    const familyB = getValveMaterialFamily_(b, col);
    const detailA = getValveMaterialDetail_(a, col);
    const detailB = getValveMaterialDetail_(b, col);
    const seriesA = getValveItemSeries_(a, col);
    const seriesB = getValveItemSeries_(b, col);

    return [
      getCatalogFittingTypeSortWeight_(productGroupA, a[col.Fitting_Type]) -
        getCatalogFittingTypeSortWeight_(productGroupB, b[col.Fitting_Type]),
      String(a[col.Fitting_Type] || '').localeCompare(String(b[col.Fitting_Type] || '')),
      getValveMaterialSortWeight_(familyA) - getValveMaterialSortWeight_(familyB),
      familyA.localeCompare(familyB),
      detailA.localeCompare(detailB),
      getValveItemSeriesSortWeight_(seriesA) - getValveItemSeriesSortWeight_(seriesB),
      seriesA.localeCompare(seriesB),
      normalizeVariant2Key_(getOptionalCellValue_(a, col, 'Variant_2')).localeCompare(normalizeVariant2Key_(getOptionalCellValue_(b, col, 'Variant_2'))),
      parseCatalogSizeSortValue_(getCatalogRowSizeParts_(a, col)[0]) - parseCatalogSizeSortValue_(getCatalogRowSizeParts_(b, col)[0]),
      parseCatalogSizeSortValue_(getCatalogRowSizeParts_(a, col)[1]) - parseCatalogSizeSortValue_(getCatalogRowSizeParts_(b, col)[1]),
      parseCatalogSizeSortValue_(getCatalogRowSizeParts_(a, col)[2]) - parseCatalogSizeSortValue_(getCatalogRowSizeParts_(b, col)[2]),
      String(a[col.Item_Number] || '').localeCompare(String(b[col.Item_Number] || ''))
    ].find(result => result !== 0) || 0;
  }
  if (productGroupA === productGroupB && isNippleCatalogProductGroup_(productGroupA)) {
    return [
      getCatalogNippleScheduleSortWeight_(getCatalogNippleScheduleGroup_(a, col)) -
        getCatalogNippleScheduleSortWeight_(getCatalogNippleScheduleGroup_(b, col)),
      getCatalogFittingTypeSortWeight_(productGroupA, a[col.Fitting_Type]) -
        getCatalogFittingTypeSortWeight_(productGroupB, b[col.Fitting_Type]),
      String(a[col.Fitting_Type] || '').localeCompare(String(b[col.Fitting_Type] || '')),
      parseCatalogSizeSortValue_(a[col.Size_1]) - parseCatalogSizeSortValue_(b[col.Size_1]),
      parseCatalogSizeSortValue_(a[col.Size_2]) - parseCatalogSizeSortValue_(b[col.Size_2]),
      parseCatalogSizeSortValue_(a[col.Size_3]) - parseCatalogSizeSortValue_(b[col.Size_3]),
      variantSortWeight_(buildCatalogNippleVariantKey_(a, col)) - variantSortWeight_(buildCatalogNippleVariantKey_(b, col)),
      String(a[col.Item_Number] || '').localeCompare(String(b[col.Item_Number] || ''))
    ].find(result => result !== 0) || 0;
  }

  const sizePartsA = getCatalogRowSizeParts_(a, col);
  const sizePartsB = getCatalogRowSizeParts_(b, col);

  return [
    normalizeMatchValue_(a[col.PLC]).localeCompare(normalizeMatchValue_(b[col.PLC]), undefined, { numeric: true }),
    getCatalogFittingTypeSortWeight_(a[col.Product_Group], a[col.Fitting_Type]) - getCatalogFittingTypeSortWeight_(b[col.Product_Group], b[col.Fitting_Type]),
    String(a[col.Fitting_Type] || '').localeCompare(String(b[col.Fitting_Type] || '')),
    normalizeVariant2Key_(getOptionalCellValue_(a, col, 'Variant_2')).localeCompare(normalizeVariant2Key_(getOptionalCellValue_(b, col, 'Variant_2'))),
    parseCatalogSizeSortValue_(sizePartsA[0]) - parseCatalogSizeSortValue_(sizePartsB[0]),
    parseCatalogSizeSortValue_(sizePartsA[1]) - parseCatalogSizeSortValue_(sizePartsB[1]),
    parseCatalogSizeSortValue_(sizePartsA[2]) - parseCatalogSizeSortValue_(sizePartsB[2]),
    String(a[col.Variant] || '').localeCompare(String(b[col.Variant] || '')),
    String(a[col.Item_Number] || '').localeCompare(String(b[col.Item_Number] || ''))
  ].find(result => result !== 0) || 0;
}

function buildCatalogMetaLine_(meta) {
  return [
    meta.catalogCode ? `Catalog ${meta.catalogCode}` : '',
    meta.versionCode ? `Version ${meta.versionCode}` : '',
    `Generated ${Utilities.formatDate(new Date(), getCatalogTimeZone_(), 'MM/dd/yyyy')}`
  ].filter(Boolean).join(' | ');
}

function exportSheetToPdfBlob_(spreadsheetId, sheetId, fileName, options) {
  const params = {
    format: 'pdf',
    gid: sheetId,
    size: 'letter',
    portrait: options.portrait ? 'true' : 'false',
    fitw: options.fitToWidth ? 'true' : 'false',
    sheetnames: options.sheetNames ? 'true' : 'false',
    printtitle: options.printTitle ? 'true' : 'false',
    pagenumbers: options.pageNumbers ? 'true' : 'false',
    gridlines: options.showGridlines ? 'true' : 'false',
    horizontal_alignment: 'CENTER',
    vertical_alignment: 'TOP',
    fzr: 'false',
    top_margin: '0.25',
    bottom_margin: '0.25',
    left_margin: '0.25',
    right_margin: '0.25'
  };

  const queryString = Object.keys(params)
    .map(key => `${encodeURIComponent(key)}=${encodeURIComponent(params[key])}`)
    .join('&');

  const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?${queryString}`;
  const response = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: `Bearer ${ScriptApp.getOAuthToken()}`
    },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error(`PDF export failed with response ${response.getResponseCode()}: ${response.getContentText()}`);
  }

  return response.getBlob().setName(fileName);
}

function applyPriceFileFormatting_(sheet, firstTableHeaderRow, lastUsedRow, multiplierRowCount, tableHeaderRows, groupDividerRows) {
  const blue = '#1F4E78';
  const lightBlue = '#D9EAF7';
  const veryLightBlue = '#EEF6FC';
  const white = '#FFFFFF';

  // Main header styling
  sheet.getRange('A1:G4')
    .setBackground(blue)
    .setFontColor(white)
    .setVerticalAlignment('middle');

  sheet.getRange('D1')
    .setFontSize(18)
    .setFontWeight('bold');

  sheet.getRange('D2')
    .setFontSize(11)
    .setFontWeight('bold');

  sheet.getRange('D3')
    .setFontSize(10)
    .setWrap(true);

  // Multiplier box
  sheet.getRange('A6:B6')
    .setFontWeight('bold')
    .setFontColor(white)
    .setBackground(blue)
    .setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true);

  if (multiplierRowCount > 0) {
    sheet.getRange(7, 1, multiplierRowCount, 2)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true);

    sheet.getRange(7, 2, multiplierRowCount, 1)
      .setBackground('#FFF2CC')
      .setFontWeight('bold')
      .setBorder(true, true, true, true, true, true, '#C65911', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }

  // Format all used data area
  if (lastUsedRow >= firstTableHeaderRow) {
    const fullRange = sheet.getRange(firstTableHeaderRow, 1, lastUsedRow - firstTableHeaderRow + 1, 7);
    fullRange.setBorder(true, true, true, true, true, true);
    fullRange.setVerticalAlignment('middle');

    // Center operational and pricing columns
    sheet.getRange(firstTableHeaderRow, 1, lastUsedRow - firstTableHeaderRow + 1, 1).setHorizontalAlignment('center'); // MSI Item #
    sheet.getRange(firstTableHeaderRow, 3, lastUsedRow - firstTableHeaderRow + 1, 5).setHorizontalAlignment('center'); // PLC through Net Ea

    // Description left aligned
    sheet.getRange(firstTableHeaderRow, 2, lastUsedRow - firstTableHeaderRow + 1, 1).setHorizontalAlignment('left');
    sheet.getRange(firstTableHeaderRow, 2, lastUsedRow - firstTableHeaderRow + 1, 1).setWrap(false);

    // PLC displayed as text so values like 000 are preserved
    sheet.getRange(firstTableHeaderRow, 3, lastUsedRow - firstTableHeaderRow + 1, 1).setNumberFormat('@');

    // Currency columns
    sheet.getRange(firstTableHeaderRow, 6, lastUsedRow - firstTableHeaderRow + 1, 2).setNumberFormat('$#,##0.00');

    const nonDataRowLookup = (tableHeaderRows || []).concat(groupDividerRows || []).reduce((lookup, rowNumber) => {
      lookup[rowNumber] = true;
      return lookup;
    }, {});
    const whiteRows = [];
    const blueRows = [];
    let dataRowIndex = 0;

    for (let rowNumber = firstTableHeaderRow; rowNumber <= lastUsedRow; rowNumber++) {
      if (nonDataRowLookup[rowNumber]) continue;
      const rangeName = `A${rowNumber}:G${rowNumber}`;
      (dataRowIndex % 2 === 0 ? whiteRows : blueRows).push(rangeName);
      dataRowIndex++;
    }

    if (whiteRows.length) sheet.getRangeList(whiteRows).setBackground(white);
    if (blueRows.length) sheet.getRangeList(blueRows).setBackground(veryLightBlue);
  }

  // Column widths tuned for 8.5 x 11 landscape style
  sheet.setColumnWidth(1, 115); // MSI Item #
  sheet.setColumnWidth(2, 330); // Description
  sheet.setColumnWidth(3, 60);  // PLC
  sheet.setColumnWidth(4, 112); // IC
  sheet.setColumnWidth(5, 120); // MC
  sheet.setColumnWidth(6, 98);  // List Price
  sheet.setColumnWidth(7, 98);  // Net Ea

  sheet.autoResizeColumn(2);
  if (sheet.getColumnWidth(2) < 330) {
    sheet.setColumnWidth(2, 330);
  }

  sheet.setRowHeights(1, 3, 24);
  sheet.setRowHeight(1, 36);
  sheet.setFrozenRows(firstTableHeaderRow);

  sheet.getRange('A:G').setFontFamily('Arial');
}

function appendCatalogLog_(logSheet, productGroup, action, status, rowsProcessed, pdfFileId, xlsFileId, message, details) {
  if (!logSheet) return;

  const metadata = details || {};
  logSheet.appendRow([
    getCatalogTimestamp_(),
    productGroup,
    action,
    status,
    rowsProcessed,
    pdfFileId,
    xlsFileId,
    message,
    metadata.runId || '',
    metadata.jobType || '',
    metadata.targetFile || '',
    metadata.startedAt || '',
    metadata.endedAt || '',
    metadata.durationMs || ''
  ]);

  try {
    CacheService.getScriptCache().remove('CATALOG_RUNTIME_STATS_V2');
  } catch (err) {
    Logger.log(`Runtime stats cache clear failed: ${err.message}`);
  }
}

function getDurationMsBetweenIso_(startedAt, endedAt) {
  const startMs = startedAt ? new Date(startedAt).getTime() : NaN;
  const endMs = endedAt ? new Date(endedAt).getTime() : NaN;
  if (isNaN(startMs) || isNaN(endMs)) return '';
  return Math.max(0, endMs - startMs);
}

function buildProductionLogDetails_(jobType, productGroup, targetFile, startedAt, endedAt, extra) {
  const details = Object.assign({}, extra || {});
  details.jobType = jobType || details.jobType || '';
  details.targetFile = targetFile || details.targetFile || '';
  details.startedAt = startedAt || details.startedAt || '';
  details.endedAt = endedAt || details.endedAt || '';
  details.durationMs = details.durationMs || getDurationMsBetweenIso_(details.startedAt, details.endedAt);
  if (!details.productGroup && productGroup) details.productGroup = productGroup;
  return details;
}

function mapCatalogLogActionToJobType_(action) {
  const normalized = String(action || '').trim().toLowerCase();
  if (normalized === 'generate slides catalog pdf') return 'catalog_pdf';
  if (normalized === 'generate price file') return 'price_file';
  if (normalized === 'pricing recalculation') return 'pricing_calc';
  if (normalized === 'repair carton counts') return 'repair_cartons';
  return '';
}

function getCatalogProductionRuntimeStats_(workbook) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'CATALOG_RUNTIME_STATS_V2';

  try {
    const cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);
  } catch (err) {
    Logger.log(`Runtime stats cache read failed: ${err.message}`);
  }

  const ss = workbook || getCatalogWorkbook_();
  const logSheet = ss.getSheetByName('Generation_Log');
  const stats = {
    generatedAt: getCatalogTimestamp_(),
    totalSamples: 0,
    typeAverages: {},
    groupTypeAverages: {},
    fileAverages: {}
  };

  if (!logSheet || logSheet.getLastRow() < 1) return stats;

  const values = logSheet.getDataRange().getValues();
  const samplesByType = {};
  const samplesByGroupType = {};
  const samplesByFile = {};

  values.forEach(row => {
    const status = String(row[3] || '').trim().toLowerCase();
    if (status !== 'success' && status !== 'completed') return;

    const productGroup = String(row[1] || '').trim();
    const action = String(row[2] || '').trim();
    const jobType = String(row[9] || '').trim() || mapCatalogLogActionToJobType_(action);
    const targetFile = String(row[10] || '').trim();
    const durationMs = Number(row[13]);
    if (!jobType || !durationMs || isNaN(durationMs) || durationMs <= 0) return;

    pushCatalogRuntimeSample_(samplesByType, jobType, durationMs);
    if (productGroup) pushCatalogRuntimeSample_(samplesByGroupType, `${jobType}|${productGroup}`, durationMs);
    if (productGroup && targetFile) pushCatalogRuntimeSample_(samplesByFile, `${jobType}|${productGroup}|${targetFile}`, durationMs);
    stats.totalSamples++;
  });

  stats.typeAverages = reduceCatalogRuntimeSamples_(samplesByType);
  stats.groupTypeAverages = reduceCatalogRuntimeSamples_(samplesByGroupType);
  stats.fileAverages = reduceCatalogRuntimeSamples_(samplesByFile);

  try {
    cache.put(cacheKey, JSON.stringify(stats), 600);
  } catch (err) {
    Logger.log(`Runtime stats cache write failed: ${err.message}`);
  }

  return stats;
}

function pushCatalogRuntimeSample_(lookup, key, durationMs) {
  if (!lookup[key]) lookup[key] = [];
  lookup[key].push(Number(durationMs) || 0);
}

function reduceCatalogRuntimeSamples_(lookup) {
  return Object.keys(lookup).reduce((result, key) => {
    const samples = (lookup[key] || []).filter(value => !isNaN(value) && value > 0);
    if (!samples.length) return result;

    const recentSamples = samples.slice(-20);
    const average = recentSamples.reduce((sum, value) => sum + value, 0) / recentSamples.length;
    result[key] = Math.round(average);
    return result;
  }, {});
}

function archiveExistingDriveFiles_(currentFolder, archiveFolder, fileNames) {
  if (!currentFolder || !archiveFolder) return [];

  const archivedFiles = [];
  const seenFileIds = {};

  fileNames.filter(Boolean).forEach(fileName => {
    const files = currentFolder.getFilesByName(fileName);

    while (files.hasNext()) {
      const file = files.next();
      const fileId = file.getId();
      if (seenFileIds[fileId]) continue;
      seenFileIds[fileId] = true;

      file.setName(buildArchivedDriveFileName_(file.getName()));
      archiveFolder.addFile(file);
      currentFolder.removeFile(file);
      archivedFiles.push(file);
    }
  });

  return archivedFiles;
}

function trashExistingDriveFiles_(folder, fileNames) {
  if (!folder) return [];

  const trashedFiles = [];
  const seenFileIds = {};

  fileNames.filter(Boolean).forEach(fileName => {
    const files = folder.getFilesByName(fileName);

    while (files.hasNext()) {
      const file = files.next();
      const fileId = file.getId();
      if (seenFileIds[fileId]) continue;
      seenFileIds[fileId] = true;

      file.setTrashed(true);
      trashedFiles.push(file);
    }
  });

  return trashedFiles;
}

function trashExistingCatalogSlidesDecks_(outputFolder, slidesFolder, meta) {
  const trashedFiles = [];
  const seenFileIds = {};
  const deckName = String(meta && meta.deckName || '').trim();
  const deckBaseName = String(meta && meta.catalogFileName || '')
    .replace(/\.pdf$/i, '')
    .trim();
  const deckPrefixes = [deckName, deckBaseName]
    .filter(Boolean)
    .filter((value, index, array) => array.indexOf(value) === index);

  function trashMatchingFiles(folder, matchMode) {
    if (!folder) return;
    const files = folder.getFiles();

    while (files.hasNext()) {
      const file = files.next();
      const fileId = file.getId();
      const fileName = String(file.getName() || '').trim();
      if (seenFileIds[fileId]) continue;

      const isMatch = deckPrefixes.some(prefix => {
        if (!prefix) return false;
        if (fileName === prefix) return true;
        if (matchMode === 'slidesPrefix') {
          return fileName.indexOf(prefix) === 0;
        }
        return false;
      });

      if (!isMatch) continue;

      seenFileIds[fileId] = true;
      file.setTrashed(true);
      trashedFiles.push(file);
    }
  }

  trashMatchingFiles(outputFolder, 'exact');
  trashMatchingFiles(slidesFolder, 'slidesPrefix');
  return trashedFiles;
}

function getOrCreateDriveSubfolder_(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();

  return parentFolder.createFolder(folderName);
}

function buildArchivedDriveFileName_(fileName) {
  const timestamp = Utilities.formatDate(new Date(), getCatalogTimeZone_(), 'yyyyMMdd-HHmmss');
  const name = String(fileName || '').trim();
  const extensionMatch = name.match(/(\.[^.]+)$/);

  if (!extensionMatch) return `${name} - archived ${timestamp}`;

  const extension = extensionMatch[1];
  const baseName = name.slice(0, -extension.length);
  return `${baseName} - archived ${timestamp}${extension}`;
}
