/**
 * Excel Pivot Builder - single-workbook Office Script.
 *
 * Run this script inside the raw-data workbook.
 * First run creates PivotBuilder_Setup. Later runs refresh fields and build pivots
 * directly in the same workbook. No Power Automate flow is required.
 */

const SETUP_SHEET = "PivotBuilder_Setup";
const SOURCE_SHEET = "__PivotSource";
const FIRST_SETUP_ROW = 9;
const LAST_SETUP_ROW = 200;

type CellValue = string | number | boolean;

// Embedded templates are installed once in each workbook. Edit the workbook rows
// freely afterward; normal runs never overwrite them. Change this block only when
// you want new workbooks to receive different organization-wide defaults.
const EMBEDDED_TEMPLATE_VERSION = "PB_EMBEDDED_1";
const SETUP_STYLE_VERSION = "PB_STYLE_6";
const EMBEDDED_TEMPLATE_ROWS: CellValue[][] = [
  [
    "Default", "Pivot 1", "Pivot_Output", "Save in this workbook",
    "", "", "", "", "", "", "",
    "", "No", "Choose fields from the dropdowns, then run Build pivots."
  ]
];

interface PivotSetup {
  template: string;
  pivotName: string;
  outputSheet: string;
  rows: string;
  rowNames: string;
  groupRules: string;
  values: string;
  filters: string;
  conditions: string;
  nextRight: boolean;
  sourceRow: number;
}

interface FieldCondition {
  field: string;
  values: string[];
  rawValue: string;
  measureField?: string;
  caption?: string;
}

interface Placement {
  row: number;
  col: number;
  bandBottom: number;
}

interface SourceModel {
  sheet: ExcelScript.Worksheet;
  sourceAddress: string;
  headers: string[];
  baseHeaders: string[];
  helperByKey: Map<string, string>;
  numericByHeader: Map<string, boolean>;
}

function main(workbook: ExcelScript.Workbook): string {
  let setup = workbook.getWorksheet(SETUP_SHEET);
  if (!setup) {
    setup = createSetupSheet(workbook);
    writeStatus(setup, "Setup created. Add or select a data sheet, complete row 9, then run this script again.");
    return "";
  }

  try {
    ensureEmbeddedTemplatesSeeded(setup);
    ensureCurrentSetupSchema(setup);
    if (text(setup.getRange("Z2").getValue()) !== SETUP_STYLE_VERSION) {
      applySetupStyle(setup);
    } else {
      resetSetupErrorStyle(setup);
    }
    refreshCloudSetup(workbook, setup);

    const action = text(setup.getRange("B5").getValue());
    if (equalsText(action, "Restore starter template") || equalsText(action, "Restore embedded templates")) {
      restoreEmbeddedTemplates(setup);
      refreshCloudSetup(workbook, setup);
      setup.getRange("B5").setValue("Build pivots");
      writeStatus(setup, "Embedded templates restored. Edit the setup rows normally; later builds will preserve your changes.");
      return "";
    }
    if (equalsText(action, "Refresh fields") || equalsText(action, "Refresh fields only")) {
      writeStatus(setup, "Fields, worksheets, and template choices refreshed.");
      return "";
    }

    const dataSheetName = text(setup.getRange("B3").getValue());
    const selectedTemplates = text(setup.getRange("B4").getValue()) || "All";
    const dataSheet = workbook.getWorksheet(dataSheetName);
    if (!dataSheet) throw new Error(`Data sheet not found: ${dataSheetName}`);
    if (dataSheet.getName() === SETUP_SHEET || dataSheet.getName() === SOURCE_SHEET) {
      throw new Error("Choose a normal data sheet, not the setup/helper sheet.");
    }

    const allSetups = readSetups(setup);
    const builds = allSetups.filter(row => templateSelected(row.template, selectedTemplates));
    if (builds.length === 0) throw new Error(`No PivotTable rows match template selection: ${selectedTemplates}`);

    writeStatus(setup, `Preparing source data for ${builds.length} PivotTable(s)...`);
    const used = dataSheet.getUsedRange(true);
    if (!used) throw new Error(`Data sheet '${dataSheet.getName()}' is empty.`);
    validateSetups(builds, readHeaderRow(used));
    const source = needsNormalizedSource(builds) || headersNeedNormalization(used)
      ? buildNormalizedSource(workbook, dataSheet, builds)
      : buildDirectSource(dataSheet, used);
    buildAllPivots(workbook, setup, source, builds);
    if (source.sheet.getName() === SOURCE_SHEET) {
      source.sheet.setVisibility(ExcelScript.SheetVisibility.veryHidden);
    }

    writeStatus(setup, `Finished: ${builds.length} PivotTable(s) built in this workbook.`);
    return headerPayloadJson(dataSheet);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    flagSetupError(setup, message);
    writeStatus(setup, `ERROR: ${message}`);
    throw error;
  }
}

function headerPayloadJson(dataSheet: ExcelScript.Worksheet): string {
  const used = dataSheet.getUsedRange(true);
  if (!used || used.getRowCount() < 1 || used.getColumnCount() < 1) {
    throw new Error(`Input data sheet has no header row: ${dataSheet.getName()}`);
  }
  const headers = readHeaderRow(used);
  return JSON.stringify({ dataSheet: dataSheet.getName(), headers });
}

function createSetupSheet(workbook: ExcelScript.Workbook): ExcelScript.Worksheet {
  const existing = workbook.getWorksheet(SETUP_SHEET);
  if (existing) return existing;
  const sheet = workbook.addWorksheet(SETUP_SHEET);
  sheet.setShowGridlines(false);

  sheet.getRange("A1:N1").merge(false);
  sheet.getRange("A1").setValue("Excel Pivot Builder");
  sheet.getRange("A3").setValue("Data sheet");
  sheet.getRange("A4").setValue("Selected template");
  sheet.getRange("A5").setValue("Action");
  sheet.getRange("A7").setValue("Status");
  sheet.getRange("B3").setValue(firstDataSheetName(workbook));
  sheet.getRange("B4").setValue("All");
  sheet.getRange("B5").setValue("Build pivots");
  sheet.getRange("B7:N7").merge(false);

  sheet.getRange("A8:N8").setValues([[
    "Template", "Pivot Name", "Output Sheet", "Save Behavior", "Rows", "Row Names",
    "Group Rules", "", "Values To Count/Sum", "Filters", "Conditions",
    "", "Next Pivot Right", "Notes"
  ]]);
  seedMissingEmbeddedTemplates(sheet);
  sheet.getRange("Z1").setValue(EMBEDDED_TEMPLATE_VERSION);

  sheet.getRange("P1:U1").merge(false);
  sheet.getRange("P1").setValue("Suggestions");
  sheet.getRange("P3").setValue("Available Fields");
  sheet.getRange("S3").setValue("Template Choices");
  sheet.getRange("T3").setValue("Sheet Choices");
  sheet.getRange("U3").setValue("Output Sheet Choices");

  applySetupStyle(sheet);
  refreshCloudSetup(workbook, sheet);
  return sheet;
}

function applySetupStyle(sheet: ExcelScript.Worksheet) {
  const burgundy = "#8D021F";
  const black = "#000000";
  const white = "#FFFFFF";
  const coolGray = "#F2F4F7";
  const border = "#C8CDD4";

  const used = sheet.getRange("A1:U500");
  used.getFormat().getFont().setName("Aptos");
  used.getFormat().getFont().setSize(10);
  used.getFormat().getFill().setColor(coolGray);

  const title = sheet.getRange("A1:N1");
  title.getFormat().getFill().setColor(burgundy);
  title.getFormat().getFont().setColor(white);
  title.getFormat().getFont().setBold(true);
  title.getFormat().getFont().setSize(16);
  title.getFormat().setRowHeight(28);

  sheet.getRange("A3:N7").getFormat().getFill().setColor(white);
  sheet.getRange("A3:A7").getFormat().getFont().setBold(true);
  sheet.getRange("A3:N7").getFormat().setWrapText(false);
  sheet.getRange("A3:N7").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
  sheet.getRange("A3:N7").getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.top);

  const headers = sheet.getRange("A8:N8");
  headers.getFormat().getFill().setColor(black);
  headers.getFormat().getFont().setColor(white);
  headers.getFormat().getFont().setBold(true);
  headers.getFormat().setWrapText(true);
  headers.getFormat().setRowHeight(30);

  const body = sheet.getRange(`A9:N${LAST_SETUP_ROW}`);
  body.getFormat().getFill().setColor(white);
  body.getFormat().setWrapText(false);
  body.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
  body.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.top);
  body.getFormat().setRowHeight(24);
  applySetupGridBorders(body, border);

  sheet.getRange("P1:U1").getFormat().getFill().setColor(burgundy);
  sheet.getRange("P1:U1").getFormat().getFont().setColor(white);
  sheet.getRange("P1:U1").getFormat().getFont().setBold(true);
  sheet.getRange("P3:U3").getFormat().getFill().setColor(black);
  sheet.getRange("P3:U3").getFormat().getFont().setColor(white);
  sheet.getRange("P3:U3").getFormat().getFont().setBold(true);
  sheet.getRange("P4:U500").getFormat().getFill().setColor(white);
  sheet.getRange("P4:U500").getFormat().setWrapText(false);
  sheet.getRange("P4:U500").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
  sheet.getRange("P4:U500").getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.top);

  const widths: { [key: string]: number } = {
    A: 110, B: 145, C: 120, D: 135, E: 160, F: 145, G: 250,
    H: 125, I: 140, J: 145, K: 185, L: 130, M: 110, N: 220,
    O: 25, P: 145, Q: 25, R: 25, S: 145, T: 145, U: 145
  };
  Object.keys(widths).forEach(column => sheet.getRange(`${column}:${column}`).getFormat().setColumnWidth(widths[column]));
  sheet.getRange("H:H").setColumnHidden(true);
  sheet.getRange("L:L").setColumnHidden(true);
  sheet.getFreezePanes().freezeRows(8);
  sheet.getRange("Z2").setValue(SETUP_STYLE_VERSION);
}

function resetSetupErrorStyle(sheet: ExcelScript.Worksheet) {
  const body = sheet.getRange(`A${FIRST_SETUP_ROW}:N${LAST_SETUP_ROW}`);
  body.getFormat().getFill().setColor("#FFFFFF");
  body.getFormat().getFont().setColor("#000000");
}

function ensureCurrentSetupSchema(sheet: ExcelScript.Worksheet) {
  if (equalsText(text(sheet.getRange("B5").getValue()), "Create output copy")) {
    sheet.getRange("B5").setValue("Build pivots");
  }
  sheet.getRange("Z3").clear(ExcelScript.ClearApplyTo.contents);
  sheet.getRange("B6:N6").unmerge();
  sheet.getRange("B6:N6").clear(ExcelScript.ClearApplyTo.contents);
  sheet.getRange("A6").clear(ExcelScript.ClearApplyTo.contents);
  sheet.getRange("A7").setValue("Status");
  sheet.getRange("B7:N7").unmerge();
  sheet.getRange("B7:N7").merge(false);
  sheet.getRange("H8").setValue("");
  sheet.getRange("I8").setValue("Values To Count/Sum");
  sheet.getRange("J8").setValue("Filters");
  sheet.getRange("K8").setValue("Conditions");
  sheet.getRange("L8").setValue("");
  sheet.getRange("P1:U1").unmerge();
  sheet.getRange("P1:U1").merge(false);
  sheet.getRange("U3").setValue("Output Sheet Choices");
  sheet.getRange("H:H").setColumnHidden(true);
  sheet.getRange("L:L").setColumnHidden(true);
}

function refreshCloudSetup(workbook: ExcelScript.Workbook, setup: ExcelScript.Worksheet) {
  populateSheetChoices(workbook, setup);
  populateTemplateChoices(setup);
  populateOutputSheetChoices(setup);
  applySetupValidations(setup);

  const dataSheetName = text(setup.getRange("B3").getValue());
  const dataSheet = workbook.getWorksheet(dataSheetName);
  if (dataSheet && dataSheet.getName() !== SETUP_SHEET && dataSheet.getName() !== SOURCE_SHEET) {
    const used = dataSheet.getUsedRange(true);
    if (used && used.getRowCount() > 0) {
      populateFieldSuggestions(setup, readHeaderRow(used));
    }
  }
}

function readHeaderRow(used: ExcelScript.Range): string[] {
  const firstRow = used.getWorksheet().getRangeByIndexes(
    used.getRowIndex(),
    used.getColumnIndex(),
    1,
    used.getColumnCount()
  );
  return normalizeHeaders(readValuesInChunks(firstRow)[0]);
}

function populateSheetChoices(workbook: ExcelScript.Workbook, setup: ExcelScript.Worksheet) {
  const configuredOutputs = setup.getRange(`C${FIRST_SETUP_ROW}:C${LAST_SETUP_ROW}`)
    .getValues()
    .map(row => text(row[0]))
    .filter(name => name !== "");
  const names = distinct(workbook.getWorksheets()
    .map(sheet => sheet.getName())
    .filter(name => name !== SETUP_SHEET &&
      name !== SOURCE_SHEET &&
      !name.startsWith("Pivot_Output") &&
      !configuredOutputs.some(output => equalsText(output, name))));
  setup.getRange("T4:T500").clear(ExcelScript.ClearApplyTo.contents);
  if (names.length > 0) setup.getRangeByIndexes(3, 19, names.length, 1).setValues(names.map(name => [name]));
}

function populateOutputSheetChoices(setup: ExcelScript.Worksheet) {
  const values = setup.getRange(`C${FIRST_SETUP_ROW}:C${LAST_SETUP_ROW}`).getValues();
  const names: string[] = ["Pivot_Output"];
  values.forEach(row => {
    const name = text(row[0]);
    if (name && !names.some(existing => equalsText(existing, name))) names.push(name);
  });
  setup.getRange("U4:U500").clear(ExcelScript.ClearApplyTo.contents);
  setup.getRangeByIndexes(3, 20, names.length, 1).setValues(names.map(name => [name]));
}

function populateTemplateChoices(setup: ExcelScript.Worksheet) {
  const values = setup.getRange(`A${FIRST_SETUP_ROW}:A${LAST_SETUP_ROW}`).getValues();
  const templates: string[] = ["All"];
  values.forEach(row => {
    const name = text(row[0]);
    if (name && !templates.some(existing => equalsText(existing, name))) templates.push(name);
  });
  if (templates.length === 1) templates.push("Default");
  setup.getRange("S4:S500").clear(ExcelScript.ClearApplyTo.contents);
  setup.getRangeByIndexes(3, 18, templates.length, 1).setValues(templates.map(name => [name]));
}

function populateFieldSuggestions(setup: ExcelScript.Worksheet, headers: string[]) {
  setup.getRange("P4:P500").clear(ExcelScript.ClearApplyTo.contents);
  const visibleHeaders = headers.slice(0, 497);
  if (visibleHeaders.length > 0) {
    setup.getRangeByIndexes(3, 15, visibleHeaders.length, 1).setValues(visibleHeaders.map(name => [name]));
  }
}

function applySetupValidations(setup: ExcelScript.Worksheet) {
  setListValidation(setup.getRange("B3"), "=$T$4:$T$500");
  setListValidation(setup.getRange("B4"), "=$S$4:$S$500");
  setListValidation(setup.getRange("B5"), "Build pivots,Refresh fields,Restore starter template");
  setListValidation(setup.getRange(`A${FIRST_SETUP_ROW}:A${LAST_SETUP_ROW}`), "=$S$5:$S$500", false);
  setListValidation(setup.getRange(`C${FIRST_SETUP_ROW}:C${LAST_SETUP_ROW}`), "=$U$4:$U$500", false);
  setListValidation(setup.getRange(`D${FIRST_SETUP_ROW}:D${LAST_SETUP_ROW}`), "Save in this workbook");
  setListValidation(setup.getRange(`M${FIRST_SETUP_ROW}:M${LAST_SETUP_ROW}`), "No,Yes");

  ["E", "I", "J", "K"].forEach(column => {
    setListValidation(setup.getRange(`${column}${FIRST_SETUP_ROW}:${column}${LAST_SETUP_ROW}`), "=$P$4:$P$500", false);
  });
}

function setListValidation(range: ExcelScript.Range, source: string, showError = true) {
  const validation = range.getDataValidation();
  validation.clear();
  validation.setRule({ list: { inCellDropDown: true, source } });
  validation.setIgnoreBlanks(true);
  validation.setErrorAlert({
    showAlert: showError,
    title: "Pivot Builder",
    message: "Choose a suggestion or type the required comma/semicolon expression."
  });
}

function applySetupGridBorders(range: ExcelScript.Range, color: string) {
  const format = range.getFormat();
  [ExcelScript.BorderIndex.diagonalDown, ExcelScript.BorderIndex.diagonalUp].forEach((index: ExcelScript.BorderIndex) => {
    format.getRangeBorder(index).setStyle(ExcelScript.BorderLineStyle.none);
  });
  [
    ExcelScript.BorderIndex.edgeTop,
    ExcelScript.BorderIndex.edgeBottom,
    ExcelScript.BorderIndex.edgeLeft,
    ExcelScript.BorderIndex.edgeRight,
    ExcelScript.BorderIndex.insideHorizontal,
    ExcelScript.BorderIndex.insideVertical
  ].forEach((index: ExcelScript.BorderIndex) => {
    const border = format.getRangeBorder(index);
    border.setColor(color);
    border.setStyle(ExcelScript.BorderLineStyle.continuous);
    border.setWeight(ExcelScript.BorderWeight.thin);
  });
}

function ensureEmbeddedTemplatesSeeded(setup: ExcelScript.Worksheet) {
  if (text(setup.getRange("Z1").getValue()) === EMBEDDED_TEMPLATE_VERSION) return;
  seedMissingEmbeddedTemplates(setup);
  setup.getRange("Z1").setValue(EMBEDDED_TEMPLATE_VERSION);
}

function seedMissingEmbeddedTemplates(setup: ExcelScript.Worksheet) {
  const target = setup.getRange(`A${FIRST_SETUP_ROW}:N${LAST_SETUP_ROW}`);
  const values = readValuesInChunks(target);
  const embeddedNames = distinct(EMBEDDED_TEMPLATE_ROWS.map(row => text(row[0])));

  embeddedNames.forEach(templateName => {
    const exists = values.some(row => equalsText(text(row[0]), templateName));
    if (exists) return;

    EMBEDDED_TEMPLATE_ROWS
      .filter(row => equalsText(text(row[0]), templateName))
      .forEach(templateRow => {
        const blankIndex = values.findIndex(row => text(row[0]) === "" && text(row[1]) === "");
        if (blankIndex < 0) throw new Error("No blank setup rows remain for embedded templates.");
        values[blankIndex] = templateRow.slice(0, 14);
      });
  });

  writeValuesInChunks(setup, values, 14, FIRST_SETUP_ROW - 1, 0);
}

function restoreEmbeddedTemplates(setup: ExcelScript.Worksheet) {
  const target = setup.getRange(`A${FIRST_SETUP_ROW}:N${LAST_SETUP_ROW}`);
  const values = readValuesInChunks(target);
  const embeddedNames = distinct(EMBEDDED_TEMPLATE_ROWS.map(row => text(row[0])));

  for (let index = 0; index < values.length; index++) {
    if (embeddedNames.some(name => equalsText(name, text(values[index][0])))) {
      values[index] = ["", "", "", "", "", "", "", "", "", "", "", "", "", ""];
    }
  }
  writeValuesInChunks(setup, values, 14, FIRST_SETUP_ROW - 1, 0);
  seedMissingEmbeddedTemplates(setup);
  setup.getRange("Z1").setValue(EMBEDDED_TEMPLATE_VERSION);
}

function readSetups(setup: ExcelScript.Worksheet): PivotSetup[] {
  const values = readValuesInChunks(setup.getRange(`A${FIRST_SETUP_ROW}:N${LAST_SETUP_ROW}`));
  const result: PivotSetup[] = [];
  values.forEach((row, index) => {
    const pivotName = text(row[1]);
    const template = text(row[0]);
    if (!pivotName || !template) return;
    result.push({
      template,
      pivotName,
      outputSheet: text(row[2]) || "Pivot_Output",
      rows: text(row[4]),
      rowNames: text(row[5]),
      groupRules: text(row[6]),
      values: text(row[8]),
      filters: text(row[9]),
      conditions: text(row[10]),
      nextRight: yes(text(row[12])),
      sourceRow: FIRST_SETUP_ROW + index
    });
  });
  return result;
}

function validateSetups(setups: PivotSetup[], headers: string[]) {
  const index = headerIndexMap(headers);
  setups.forEach(setup => {
    const rowFields = splitList(setup.rows);
    const rowNames = splitList(setup.rowNames);
    const ruleSets = splitRuleSets(setup.groupRules);
    const valueFields = splitList(setup.values);

    rowFields.forEach(field => requireSetupHeader(index, field, setup.sourceRow, "Rows"));
    valueFields.forEach(field => requireSetupHeader(index, field, setup.sourceRow, "Values To Count/Sum"));
    if (rowNames.length > rowFields.length && rowNames.some((name, position) => name && !rowFields[position] && !ruleSets[position])) {
      throw new Error(`Setup row ${setup.sourceRow}: Row Names contains a name without a matching Rows field or Group Rule.`);
    }
    ruleSets.forEach(rules => validateGroupRuleSet(rules, index, setup.sourceRow));

    let filters: FieldCondition[];
    let conditions: FieldCondition[];
    try {
      filters = parseFilterSpecs(setup.filters);
    } catch (error) {
      throw new Error(`Setup row ${setup.sourceRow}: ${errorDetail(error)}`);
    }
    try {
      conditions = parseConditions(setup.conditions);
    } catch (error) {
      throw new Error(`Setup row ${setup.sourceRow}: ${errorDetail(error)}`);
    }
    filters.forEach(filter => requireSetupHeader(index, filter.field, setup.sourceRow, "Filters"));
    conditions.forEach(condition => {
      requireSetupHeader(index, condition.field, setup.sourceRow, "Conditions");
      if (condition.measureField) {
        requireSetupHeader(index, condition.measureField, setup.sourceRow, "Conditions measure");
      }
    });
    if (conditions.some(condition => !condition.measureField) && valueFields.length === 0) {
      throw new Error(`Setup row ${setup.sourceRow}: Conditions need at least one Values To Count/Sum field.`);
    }
  });
}

function requireSetupHeader(
  index: Map<string, number>,
  field: string,
  setupRow: number,
  columnName: string
) {
  if (!index.has(lower(field))) {
    throw new Error(`Setup row ${setupRow}: ${columnName} field '${field}' was not found in the data headers.`);
  }
}

function validateGroupRuleSet(rulesText: string, index: Map<string, number>, setupRow: number) {
  const rules = rulesText.split("|").map(rule => rule.trim()).filter(rule => rule !== "");
  if (rules.length === 0) throw new Error(`Setup row ${setupRow}: Group Rules is blank or invalid.`);
  rules.forEach(rule => {
    const colon = rule.indexOf(":");
    if (colon < 1 || colon === rule.length - 1) {
      throw new Error(`Setup row ${setupRow}: Group Rule must use Name:Field=Value: '${rule}'.`);
    }
    const conditions = rule.slice(colon + 1).split(";").map(value => value.trim()).filter(value => value !== "");
    if (conditions.length === 0) throw new Error(`Setup row ${setupRow}: Group Rule has no conditions: '${rule}'.`);
    conditions.forEach(condition => {
      const match = condition.match(/^(.+?)(<=|>=|<>|!=|=|<|>)(.+)$/);
      if (!match) throw new Error(`Setup row ${setupRow}: Invalid Group Rule condition '${condition}'.`);
      requireSetupHeader(index, match[1].trim(), setupRow, "Group Rules");
    });
  });
}

function flagSetupError(setup: ExcelScript.Worksheet, message: string) {
  const match = message.match(/Setup row (\d+)/i);
  if (!match) return;
  const row = Number(match[1]);
  if (row < FIRST_SETUP_ROW || row > LAST_SETUP_ROW) return;
  const range = setup.getRange(`A${row}:N${row}`);
  range.getFormat().getFill().setColor("#FCE8EC");
  range.getFormat().getFont().setColor("#8D021F");
}

function needsNormalizedSource(setups: PivotSetup[]): boolean {
  return setups.some(setup => {
    const rows = splitList(setup.rows);
    const values = splitList(setup.values);
    const filterOverlap = parseFilterSpecs(setup.filters).some(filter =>
      containsText(rows, filter.field) || containsText(values, filter.field)
    );
    return setup.groupRules.trim() !== "" ||
      setup.conditions.trim() !== "" ||
      filterOverlap;
  });
}

function headersNeedNormalization(used: ExcelScript.Range): boolean {
  const firstRow = used.getWorksheet().getRangeByIndexes(
    used.getRowIndex(),
    used.getColumnIndex(),
    1,
    used.getColumnCount()
  );
  const raw = readValuesInChunks(firstRow)[0].map(value => text(value));
  const seen: string[] = [];
  return raw.some(header => {
    if (!header) return true;
    const duplicate = seen.some(existing => equalsText(existing, header));
    seen.push(header);
    return duplicate;
  });
}

function buildDirectSource(
  dataSheet: ExcelScript.Worksheet,
  used: ExcelScript.Range
): SourceModel {
  const headers = readHeaderRow(used);
  const sampleRowCount = Math.min(31, used.getRowCount());
  const sample = readValuesInChunks(dataSheet.getRangeByIndexes(
    used.getRowIndex(),
    used.getColumnIndex(),
    sampleRowCount,
    used.getColumnCount()
  ));
  const sampleRows = sample.slice(1);
  const numericByHeader = new Map<string, boolean>();
  headers.forEach((header, column) => {
    numericByHeader.set(lower(header), columnLooksNumeric(sampleRows, column));
  });
  return {
    sheet: dataSheet,
    sourceAddress: worksheetRangeAddress(
      dataSheet.getName(),
      used.getRowCount(),
      used.getColumnCount(),
      used.getRowIndex(),
      used.getColumnIndex()
    ),
    headers,
    baseHeaders: headers,
    helperByKey: new Map<string, string>(),
    numericByHeader
  };
}

function buildNormalizedSource(
  workbook: ExcelScript.Workbook,
  dataSheet: ExcelScript.Worksheet,
  setups: PivotSetup[]
): SourceModel {
  const used = dataSheet.getUsedRange(true);
  if (!used || used.getRowCount() < 2 || used.getColumnCount() < 1) {
    throw new Error(`Data sheet '${dataSheet.getName()}' needs one header row and at least one data row.`);
  }

  const sourceValues = readValuesInChunks(used);
  const baseHeaders = normalizeHeaders(sourceValues[0]);
  // Chunk reads already return rectangular rows. Reuse those rows instead of
  // keeping several full copies of a large source dataset in script memory.
  const dataRows = sourceValues.slice(1);
  const headers = [...baseHeaders];
  const outputRows = dataRows;
  const helperByKey = new Map<string, string>();
  const baseIndex = headerIndexMap(baseHeaders);

  const addHelper = (key: string, preferredName: string, values: CellValue[]) => {
    if (helperByKey.has(key)) return;
    const header = uniqueHeader(preferredName, headers);
    helperByKey.set(key, header);
    headers.push(header);
    outputRows.forEach((row, index) => row.push(values[index] ?? ""));
  };

  setups.forEach(setup => {
    const rowFields = splitList(setup.rows);
    const rowNames = splitList(setup.rowNames);
    const ruleSets = splitRuleSets(setup.groupRules);
    const valueFields = splitList(setup.values);

    const maxRows = Math.max(rowFields.length, rowNames.length, ruleSets.length);
    for (let position = 0; position < maxRows; position++) {
      const field = rowFields[position] || "";
      const rowName = rowNames[position] || "";
      const rules = ruleSets[position] || "";
      if (!field && !rules) continue;

      if (rules) {
        if (field) requireHeader(baseIndex, field, `Rows, setup row ${setup.sourceRow}`);
        const caption = rowName || (field ? `${field} Group` : "Condition Group");
        const key = groupKey(field, caption, rules);
        addHelper(key, `PB Group - ${caption}`, dataRows.map(row => applyGroupRules(row, baseIndex, rules)));
      }
    }

    parseFilterSpecs(setup.filters).forEach(spec => {
      if (!containsText(rowFields, spec.field) && !containsText(valueFields, spec.field)) return;
      const fieldIndex = requireHeader(baseIndex, spec.field, `Filters, setup row ${setup.sourceRow}`);
      addHelper(filterKey(spec.field), `PB Filter - ${spec.field}`, dataRows.map(row => row[fieldIndex]));
    });

    parseConditions(setup.conditions).forEach(condition => {
      const conditionIndex = requireHeader(baseIndex, condition.field, `Conditions, setup row ${setup.sourceRow}`);
      const conditionValueFields = condition.measureField ? [condition.measureField] : valueFields;
      conditionValueFields.forEach(valueField => {
        const valueIndex = requireHeader(baseIndex, valueField, `Values To Count/Sum, setup row ${setup.sourceRow}`);
        addHelper(
          measureKey(condition.field, condition.rawValue, valueField),
          `PB Measure - ${condition.field} - ${condition.rawValue} - ${valueField}`,
          dataRows.map(row => valueMatches(row[conditionIndex], condition.values) ? row[valueIndex] : "")
        );
      });
    });
  });

  const requiredBaseFields: string[] = [];
  setups.forEach(setup => {
    splitList(setup.rows).forEach(field => requiredBaseFields.push(field));
    splitList(setup.values).forEach(field => requiredBaseFields.push(field));
    parseFilterSpecs(setup.filters).forEach(filter => requiredBaseFields.push(filter.field));
  });
  const requiredKeys = distinct(requiredBaseFields).map(field => lower(field));
  const keptColumns: number[] = [];
  headers.forEach((header, index) => {
    if (index >= baseHeaders.length || requiredKeys.includes(lower(header))) keptColumns.push(index);
  });
  if (keptColumns.length === 0) keptColumns.push(0);
  const finalHeaders = keptColumns.map(index => headers[index]);
  const finalRows = outputRows.map(row => keptColumns.map(index => row[index]));

  let helper = workbook.getWorksheet(SOURCE_SHEET);
  if (!helper) helper = workbook.addWorksheet(SOURCE_SHEET);
  const matrix: CellValue[][] = [finalHeaders, ...finalRows];
  if (matrix.length > 1048576) {
    throw new Error(`Source has ${matrix.length} rows after adding headers; Excel allows at most 1,048,576.`);
  }
  if (finalHeaders.length > 16384) {
    throw new Error(`Source has ${finalHeaders.length} columns after helper fields; Excel allows at most 16,384.`);
  }
  writeValuesInChunks(helper, matrix, finalHeaders.length);
  const sourceAddress = worksheetRangeAddress(helper.getName(), matrix.length, finalHeaders.length);
  const numericByHeader = new Map<string, boolean>();
  finalHeaders.forEach((header, col) => numericByHeader.set(lower(header), columnLooksNumeric(finalRows, col)));
  return { sheet: helper, sourceAddress, headers: finalHeaders, baseHeaders, helperByKey, numericByHeader };
}

function worksheetRangeAddress(
  sheetName: string,
  rowCount: number,
  columnCount: number,
  startRow = 0,
  startColumn = 0
): string {
  const escapedSheet = sheetName.replace(/'/g, "''");
  const firstColumn = columnLetters(startColumn + 1);
  const lastColumn = columnLetters(startColumn + columnCount);
  return `'${escapedSheet}'!$${firstColumn}$${startRow + 1}:$${lastColumn}$${startRow + rowCount}`;
}

function columnLetters(columnCount: number): string {
  let value = columnCount;
  let result = "";
  while (value > 0) {
    value--;
    result = String.fromCharCode(65 + (value % 26)) + result;
    value = Math.floor(value / 26);
  }
  return result;
}

function readValuesInChunks(range: ExcelScript.Range): CellValue[][] {
  const sheet = range.getWorksheet();
  const startRow = range.getRowIndex();
  const startColumn = range.getColumnIndex();
  const rowCount = range.getRowCount();
  const columnCount = range.getColumnCount();
  const result: CellValue[][] = Array.from(
    { length: rowCount },
    () => Array(columnCount).fill("") as CellValue[]
  );

  for (let columnOffset = 0; columnOffset < columnCount; columnOffset += MAX_READ_COLUMNS_PER_CHUNK) {
    const chunkColumns = Math.min(MAX_READ_COLUMNS_PER_CHUNK, columnCount - columnOffset);
    const rowsPerChunk = chunkRowCount(chunkColumns);
    for (let rowOffset = 0; rowOffset < rowCount; rowOffset += rowsPerChunk) {
      const chunkRows = Math.min(rowsPerChunk, rowCount - rowOffset);
      const values = readChunkWithRetry(
        sheet,
        startRow + rowOffset,
        startColumn + columnOffset,
        chunkRows,
        chunkColumns
      );
      for (let row = 0; row < values.length; row++) {
        for (let column = 0; column < chunkColumns; column++) {
          result[rowOffset + row][columnOffset + column] = values[row][column];
        }
      }
    }
  }
  return result;
}

function readChunkWithRetry(
  sheet: ExcelScript.Worksheet,
  startRow: number,
  startColumn: number,
  rowCount: number,
  columnCount: number
): CellValue[][] {
  try {
    return sheet.getRangeByIndexes(
      startRow,
      startColumn,
      rowCount,
      columnCount
    ).getValues() as CellValue[][];
  } catch (error) {
    if (rowCount > 1) {
      const firstCount = Math.floor(rowCount / 2);
      const top = readChunkWithRetry(sheet, startRow, startColumn, firstCount, columnCount);
      const bottom = readChunkWithRetry(
        sheet,
        startRow + firstCount,
        startColumn,
        rowCount - firstCount,
        columnCount
      );
      return [...top, ...bottom];
    }
    if (columnCount > 1) {
      const firstCount = Math.floor(columnCount / 2);
      const left = readChunkWithRetry(sheet, startRow, startColumn, rowCount, firstCount);
      const right = readChunkWithRetry(
        sheet,
        startRow,
        startColumn + firstCount,
        rowCount,
        columnCount - firstCount
      );
      return left.map((row, index) => [...row, ...right[index]]);
    }
    const cell = sheet.getCell(startRow, startColumn);
    try {
      return [[cell.getValue() as CellValue]];
    } catch (valueError) {
      try {
        return [[cell.getText()]];
      } catch (textError) {
        throw new Error(
          `Could not read source cell at row ${startRow + 1}, column ${startColumn + 1}: ` +
          `${errorDetail(error)}; value fallback: ${errorDetail(valueError)}; ` +
          `text fallback: ${errorDetail(textError)}`
        );
      }
    }
  }
}

function writeValuesInChunks(
  sheet: ExcelScript.Worksheet,
  values: CellValue[][],
  columnCount: number,
  startRow = 0,
  startColumn = 0
) {
  for (let columnOffset = 0; columnOffset < columnCount; columnOffset += MAX_WRITE_COLUMNS_PER_CHUNK) {
    const chunkColumns = Math.min(MAX_WRITE_COLUMNS_PER_CHUNK, columnCount - columnOffset);
    const maximumRows = chunkRowCount(chunkColumns);
    let rowOffset = 0;
    while (rowOffset < values.length) {
      const chunk: CellValue[][] = [];
      let estimatedBytes = 2;
      while (rowOffset + chunk.length < values.length && chunk.length < maximumRows) {
        const nextRow = values[rowOffset + chunk.length].slice(
          columnOffset,
          columnOffset + chunkColumns
        );
        const nextBytes = estimateRowPayloadBytes(nextRow);
        if (chunk.length > 0 && estimatedBytes + nextBytes > MAX_WRITE_PAYLOAD_BYTES) break;
        chunk.push(nextRow);
        estimatedBytes += nextBytes;
      }
      writeChunkWithRetry(
        sheet,
        startRow + rowOffset,
        startColumn + columnOffset,
        chunk
      );
      rowOffset += chunk.length;
    }
  }
}

function writeChunkWithRetry(
  sheet: ExcelScript.Worksheet,
  startRow: number,
  startColumn: number,
  values: CellValue[][]
): void {
  const rowCount = values.length;
  const columnCount = values[0].length;
  try {
    sheet.getRangeByIndexes(startRow, startColumn, rowCount, columnCount).setValues(values);
    return;
  } catch (error) {
    if (rowCount > 1) {
      const firstCount = Math.floor(rowCount / 2);
      writeChunkWithRetry(sheet, startRow, startColumn, values.slice(0, firstCount));
      writeChunkWithRetry(sheet, startRow + firstCount, startColumn, values.slice(firstCount));
      return;
    }
    if (columnCount > 1) {
      const firstCount = Math.floor(columnCount / 2);
      const left = values.map(row => row.slice(0, firstCount));
      const right = values.map(row => row.slice(firstCount));
      writeChunkWithRetry(sheet, startRow, startColumn, left);
      writeChunkWithRetry(sheet, startRow, startColumn + firstCount, right);
      return;
    }
    try {
      sheet.getCell(startRow, startColumn).setValue(values[0][0]);
      return;
    } catch (cellError) {
      throw new Error(
        `Could not write source cell at row ${startRow + 1}, column ${startColumn + 1}: ` +
        `${errorDetail(error)}; cell fallback: ${errorDetail(cellError)}`
      );
    }
  }
}

function errorDetail(error: unknown): string {
  if (error instanceof Error) return error.message || error.name;
  if (typeof error === "string") return error;
  if (error && typeof error === "object") {
    const value = error as { message?: unknown; code?: unknown; name?: unknown };
    const pieces: string[] = [];
    if (value.code !== undefined) pieces.push(`code=${String(value.code)}`);
    if (value.name !== undefined) pieces.push(`name=${String(value.name)}`);
    if (value.message !== undefined) pieces.push(`message=${String(value.message)}`);
    if (pieces.length > 0) return pieces.join(", ");
    try {
      return JSON.stringify(error);
    } catch (_ignored) {
      return "Unknown Excel error";
    }
  }
  return String(error);
}

function estimateRowPayloadBytes(row: CellValue[]): number {
  let bytes = 4;
  row.forEach(value => {
    if (typeof value === "string") {
      // Four bytes per character safely covers UTF-8 plus JSON escaping for
      // ordinary worksheet text. Quotes, slashes, and separators get padding.
      bytes += value.length * 4 + 16;
    } else {
      bytes += 32;
    }
  });
  return bytes;
}

const MAX_READ_COLUMNS_PER_CHUNK = 100;
const MAX_WRITE_COLUMNS_PER_CHUNK = 20;
const MAX_WRITE_PAYLOAD_BYTES = 750000;

function chunkRowCount(columnCount: number): number {
  // Keep individual Excel API requests comfortably below the documented
  // payload ceiling, including workbooks containing longer text values.
  const targetCells = 25000;
  return Math.max(1, Math.min(10000, Math.floor(targetCells / Math.max(1, columnCount))));
}

function buildAllPivots(
  workbook: ExcelScript.Workbook,
  setupSheet: ExcelScript.Worksheet,
  source: SourceModel,
  setups: PivotSetup[]
) {
  const outputs = distinct(setups.map(setup => safeSheetName(setup.outputSheet || "Pivot_Output")));
  const outputSheets = new Map<string, ExcelScript.Worksheet>();
  outputs.forEach(name => {
    // Never delete or replace an existing worksheet. This also protects the
    // source data if a template accidentally requests its sheet name.
    const actualName = uniqueWorksheetName(workbook, name);
    const sheet = workbook.addWorksheet(actualName);
    sheet.setShowGridlines(false);
    outputSheets.set(lower(name), sheet);
  });

  const placements = new Map<string, Placement>();
  outputs.forEach(name => placements.set(lower(name), { row: 0, col: 0, bandBottom: 0 }));

  setups.forEach(setup => {
    const outputName = safeSheetName(setup.outputSheet || "Pivot_Output");
    const sheet = outputSheets.get(lower(outputName));
    const placement = placements.get(lower(outputName));
    if (!sheet || !placement) throw new Error(`Output sheet unavailable: ${outputName}`);

    buildOnePivot(workbook, sheet, source, setup, placement);
  });

  outputSheets.forEach(sheet => {
    const used = sheet.getUsedRange();
    if (used) {
      // Full-sheet autofit can time out when a PivotTable contains many rows.
      // A representative preview gives useful widths without scanning the
      // entire output. Pivot rows retain Excel's normal default height.
      const previewRows = Math.min(100, used.getRowCount());
      const previewColumns = Math.min(50, used.getColumnCount());
      sheet.getRangeByIndexes(
        used.getRowIndex(),
        used.getColumnIndex(),
        previewRows,
        previewColumns
      ).getFormat().autofitColumns();
    }
  });
}

function buildOnePivot(
  workbook: ExcelScript.Workbook,
  sheet: ExcelScript.Worksheet,
  source: SourceModel,
  setup: PivotSetup,
  placement: Placement
) {
  const rowFields = splitList(setup.rows);
  const rowNames = splitList(setup.rowNames);
  const ruleSets = splitRuleSets(setup.groupRules);
  const valueFields = splitList(setup.values);
  const filters = parseFilterSpecs(setup.filters);
  const conditions = parseConditions(setup.conditions);

  const titleRow = placement.row;
  const titleCol = placement.col;
  const reportFilterCount = filters.filter(spec => !containsText(valueFields, spec.field)).length;
  const pivotStartRow = titleRow + 4 + reportFilterCount;
  const title = sheet.getCell(titleRow, titleCol);
  title.setValue(setup.pivotName);
  title.getFormat().getFont().setBold(true);
  title.getFormat().getFont().setSize(14);
  title.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
  sheet.getCell(titleRow + 1, titleCol).setValue(summaryText(setup));
  sheet.getCell(titleRow + 1, titleCol).getFormat().getFont().setItalic(true);
  sheet.getCell(titleRow + 1, titleCol).getFormat().getFont().setColor("#5A6068");
  sheet.getCell(titleRow + 1, titleCol).getFormat().setWrapText(false);
  sheet.getCell(titleRow + 1, titleCol).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);

  const pivotName = uniquePivotName(workbook, setup.pivotName);
  // Address strings are more compatible across Excel web tenants than passing
  // Range objects from a hidden helper sheet directly.
  const pivot = sheet.addPivotTable(
    pivotName,
    source.sourceAddress,
    sheet.getCell(pivotStartRow, titleCol).getAddress()
  );

  const maxRows = Math.max(rowFields.length, rowNames.length, ruleSets.length);
  for (let position = 0; position < maxRows; position++) {
    const field = rowFields[position] || "";
    const rowName = rowNames[position] || "";
    const rules = ruleSets[position] || "";
    if (!field && !rules) continue;
    let actual = field;
    let caption = rowName;
    if (rules) {
      caption = rowName || (field ? `${field} Group` : "Condition Group");
      actual = requiredHelper(source, groupKey(field, caption, rules));
    }
    const hierarchy = pivot.getHierarchy(actual);
    if (!hierarchy) throw new Error(`Row field not found: ${actual}`);
    const added = pivot.addRowHierarchy(hierarchy);
    if (caption) added.setName(caption);
  }

  valueFields.forEach(field => {
    const actual = requirePivotHeader(source, field);
    addDataHierarchy(pivot, source, actual, field);
  });

  conditions.forEach(condition => {
    const conditionValueFields = condition.measureField ? [condition.measureField] : valueFields;
    conditionValueFields.forEach(valueField => {
      const actual = requiredHelper(source, measureKey(condition.field, condition.rawValue, valueField));
      const caption = condition.caption ||
        (conditionValueFields.length === 1 ? condition.rawValue : `${condition.rawValue} - ${valueField}`);
      addDataHierarchy(pivot, source, actual, caption);
    });
  });

  filters.forEach(spec => {
    const duplicateFilterField = source.helperByKey.get(filterKey(spec.field));
    const placeInColumns = containsText(valueFields, spec.field);
    const actual = duplicateFilterField || requirePivotHeader(source, spec.field);
    const hierarchy = pivot.getHierarchy(actual);
    if (!hierarchy) throw new Error(`Setup row ${setup.sourceRow}: Filter field not found: ${spec.field}`);
    let filterField: ExcelScript.PivotField;
    if (placeInColumns) {
      const added = pivot.addColumnHierarchy(hierarchy);
      added.setName(spec.field);
      filterField = added.getFields()[0];
    } else {
      const added = pivot.addFilterHierarchy(hierarchy);
      added.setName(spec.field);
      filterField = added.getFields()[0];
    }
    if (spec.values.length > 0) {
      try {
        filterField.applyFilter({ manualFilter: { selectedItems: spec.values } });
      } catch (_ignored) {
        // Keep the filter dropdown when a requested item is absent, but leave
        // it unselected instead of failing the entire PivotTable build.
      }
    }
  });

  pivot.getLayout().setLayoutType(ExcelScript.PivotLayoutType.tabular);
  const pivotRange = pivot.getLayout().getRange();
  const bottom = Math.max(titleRow + 1, pivotRange.getRowIndex() + pivotRange.getRowCount() - 1) + 5;
  const right = Math.max(titleCol, pivotRange.getColumnIndex() + pivotRange.getColumnCount() - 1) + 4;
  placement.bandBottom = Math.max(placement.bandBottom, bottom);
  if (setup.nextRight) {
    placement.col = right;
  } else {
    placement.row = placement.bandBottom + 1;
    placement.col = 0;
    placement.bandBottom = placement.row;
  }
}

function addDataHierarchy(
  pivot: ExcelScript.PivotTable,
  source: SourceModel,
  actualHeader: string,
  displayName: string
) {
  const hierarchy = pivot.getHierarchy(actualHeader);
  if (!hierarchy) throw new Error(`Value field not found: ${actualHeader}`);
  const added = pivot.addDataHierarchy(hierarchy);
  const numeric = source.numericByHeader.get(lower(actualHeader)) === true;
  added.setSummarizeBy(numeric ? ExcelScript.AggregationFunction.sum : ExcelScript.AggregationFunction.count);
  added.setName(`${numeric ? "Sum" : "Count"} of ${displayName}`);
}

function normalizeHeaders(values: CellValue[]): string[] {
  const result: string[] = [];
  values.forEach((value, index) => {
    let header = text(value) || `Column ${index + 1}`;
    header = uniqueHeader(header, result);
    result.push(header);
  });
  return result;
}

function headerIndexMap(headers: string[]): Map<string, number> {
  const result = new Map<string, number>();
  headers.forEach((header, index) => result.set(lower(header), index));
  return result;
}

function requireHeader(index: Map<string, number>, field: string, context: string): number {
  const found = index.get(lower(field));
  if (found === undefined) throw new Error(`Field not found in ${context}: ${field}`);
  return found;
}

function requirePivotHeader(source: SourceModel, field: string): string {
  const found = source.headers.find(header => equalsText(header, field));
  if (!found) throw new Error(`Field not found: ${field}`);
  return found;
}

function requiredHelper(source: SourceModel, key: string): string {
  const found = source.helperByKey.get(key);
  if (!found) throw new Error(`Internal helper was not created: ${key}`);
  return found;
}

function columnLooksNumeric(rows: CellValue[][], column: number): boolean {
  let checked = 0;
  for (const row of rows) {
    const value = row[column];
    if (value === "" || value === null || value === undefined) continue;
    checked++;
    if (typeof value !== "number") return false;
    if (checked >= 30) break;
  }
  return checked > 0;
}

function parseFilterSpecs(value: string): FieldCondition[] {
  const normalized = value.replace(/\r?\n/g, ";").trim();
  if (!normalized) return [];
  const pieces = normalized.includes("=") ? splitSemicolon(normalized) : splitList(normalized);
  return pieces.map(piece => {
    const equals = piece.indexOf("=");
    if (equals < 0) {
      const field = piece.trim();
      if (!field) throw new Error("Filter field is blank.");
      return { field, values: [], rawValue: "" };
    }
    if (equals < 1) throw new Error(`Filter must use Field or Field=Value: ${piece}`);
    const field = piece.slice(0, equals).trim();
    const rawValue = piece.slice(equals + 1).trim();
    if (!rawValue) throw new Error(`Filter value is blank: ${piece}`);
    return { field, values: splitList(rawValue), rawValue };
  }).filter(item => item.field !== "");
}

function parseConditions(value: string): FieldCondition[] {
  return splitSemicolon(value).map(piece => {
    const arrow = piece.indexOf("=>");
    let expression = arrow >= 0 ? piece.slice(0, arrow).trim() : piece.trim();
    const measureField = arrow >= 0 ? piece.slice(arrow + 2).trim() : "";
    if (arrow >= 0 && !measureField) throw new Error(`Condition measure field is blank: ${piece}`);
    const equals = expression.indexOf("=");
    if (equals < 1) throw new Error(`Condition must use Field=Value: ${piece}`);
    let caption = "";
    const colon = expression.indexOf(":");
    if (colon > 0 && colon < equals) {
      caption = expression.slice(0, colon).trim();
      expression = expression.slice(colon + 1).trim();
    }
    const adjustedEquals = expression.indexOf("=");
    const field = expression.slice(0, adjustedEquals).trim();
    const rawValue = expression.slice(adjustedEquals + 1).trim();
    if (!rawValue) throw new Error(`Condition value is blank: ${piece}`);
    return {
      field,
      values: splitList(rawValue),
      rawValue,
      measureField: measureField || undefined,
      caption: caption || undefined
    };
  });
}

function applyGroupRules(row: CellValue[], headerIndex: Map<string, number>, rulesText: string): string {
  const safeRulesText = typeof rulesText === "string" ? rulesText : "";
  const rules = safeRulesText
    .split("|")
    .map((item: string) => item.trim())
    .filter((item: string) => item.length > 0);
  for (const rule of rules) {
    const colon = rule.indexOf(":");
    if (colon < 1) continue;
    const label = rule.slice(0, colon).trim();
    const conditions = rule.slice(colon + 1).trim();
    if (label && rowMatches(row, headerIndex, conditions)) return label;
  }
  return "Other";
}

function rowMatches(row: CellValue[], headerIndex: Map<string, number>, conditionsText: string): boolean {
  const conditions = conditionsText
    .split(";")
    .map((item: string) => item.trim())
    .filter((item: string) => item.length > 0);
  if (conditions.length === 0) return false;
  return conditions.every(condition => {
    const match = condition.match(/^(.+?)(<=|>=|<>|!=|=|<|>)(.*)$/);
    if (!match) return false;
    const field = match[1].trim();
    const operator = match[2];
    const expected = match[3].trim();
    const index = headerIndex.get(lower(field));
    if (index === undefined) return false;
    return compareValues(row[index], expected, operator);
  });
}

function compareValues(actual: CellValue, expectedText: string, operator: string): boolean {
  const actualText = text(actual);
  const actualNumber = typeof actual === "number" ? actual : Number(actualText);
  const expectedNumber = Number(expectedText);
  const numeric = !Number.isNaN(actualNumber) && !Number.isNaN(expectedNumber) && expectedText !== "";
  if (numeric) return compareNumbers(actualNumber, expectedNumber, operator);
  return compareStrings(actualText.toLowerCase(), expectedText.toLowerCase(), operator);
}

function compareNumbers(left: number, right: number, operator: string): boolean {
  switch (operator) {
    case "=": return left === right;
    case "!=":
    case "<>": return left !== right;
    case ">": return left > right;
    case "<": return left < right;
    case ">=": return left >= right;
    case "<=": return left <= right;
    default: return false;
  }
}

function compareStrings(left: string, right: string, operator: string): boolean {
  switch (operator) {
    case "=": return left === right;
    case "!=":
    case "<>": return left !== right;
    case ">": return left > right;
    case "<": return left < right;
    case ">=": return left >= right;
    case "<=": return left <= right;
    default: return false;
  }
}

function valueMatches(actual: CellValue, expected: string[]): boolean {
  return expected.some(value => equalsText(text(actual), value));
}

function summaryText(setup: PivotSetup): string {
  const conditions = setup.conditions.trim() || "none";
  const filters = setup.filters.trim() || "none";
  return `Conditional values: ${conditions} | Filters: ${filters}`;
}

function templateSelected(template: string, selected: string): boolean {
  if (equalsText(selected, "All") || equalsText(selected, "All Templates")) return true;
  return splitList(selected).some(item => equalsText(item, template));
}

function firstDataSheetName(workbook: ExcelScript.Workbook): string {
  const found = workbook.getWorksheets().find(sheet => sheet.getName() !== SETUP_SHEET && sheet.getName() !== SOURCE_SHEET);
  return found ? found.getName() : "Data";
}

function uniquePivotName(workbook: ExcelScript.Workbook, preferred: string): string {
  const base = safeObjectName(preferred || "Pivot").slice(0, 180);
  const existing = workbook.getPivotTables().map(pivot => lower(pivot.getName()));
  if (!existing.includes(lower(base))) return base;
  for (let index = 2; index < 1000; index++) {
    const candidate = `${base}_${index}`;
    if (!existing.includes(lower(candidate))) return candidate;
  }
  return `${base}_${Date.now()}`;
}

function uniqueTableName(workbook: ExcelScript.Workbook, preferred: string): string {
  const existing: string[] = [];
  workbook.getWorksheets().forEach(sheet => sheet.getTables().forEach(table => existing.push(lower(table.getName()))));
  if (!existing.includes(lower(preferred))) return preferred;
  for (let index = 2; index < 1000; index++) {
    const candidate = `${preferred}_${index}`;
    if (!existing.includes(lower(candidate))) return candidate;
  }
  return `${preferred}_${Date.now()}`;
}

function uniqueHeader(preferred: string, headers: string[]): string {
  const base = (preferred.trim() || "Column").slice(0, 240);
  if (!headers.some(header => equalsText(header, base))) return base;
  for (let index = 2; index < 1000; index++) {
    const candidate = `${base.slice(0, 235)} ${index}`;
    if (!headers.some(header => equalsText(header, candidate))) return candidate;
  }
  return `${base.slice(0, 225)} ${Date.now()}`;
}

function safeSheetName(value: string): string {
  const cleaned = value.replace(/[\\/\?\*\[\]:]/g, "_").trim() || "Pivot_Output";
  return cleaned.slice(0, 31);
}

function uniqueWorksheetName(workbook: ExcelScript.Workbook, preferred: string): string {
  const base = safeSheetName(preferred);
  const existing = workbook.getWorksheets().map((sheet: ExcelScript.Worksheet) => lower(sheet.getName()));
  if (!existing.includes(lower(base))) return base;

  for (let index = 2; index < 10000; index++) {
    const suffix = `_${index}`;
    const candidate = `${base.slice(0, 31 - suffix.length)}${suffix}`;
    if (!existing.includes(lower(candidate))) return candidate;
  }
  return `Pivot_${Date.now()}`.slice(0, 31);
}

function safeObjectName(value: string): string {
  const cleaned = value.replace(/[^A-Za-z0-9_]/g, "_");
  return /^[A-Za-z_]/.test(cleaned) ? cleaned : `P_${cleaned}`;
}

function splitList(value: string): string[] {
  return value
    .replace(/\r?\n/g, ",")
    .split(",")
    .map((item: string) => item.trim())
    .filter((item: string) => item.length > 0);
}

function splitSemicolon(value: string): string[] {
  return value
    .replace(/\r?\n/g, ";")
    .split(";")
    .map((item: string) => item.trim())
    .filter((item: string) => item.length > 0);
}

function splitRuleSets(value: string): string[] {
  const normalized = value.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  return normalized.includes("\n")
    ? normalized.split("\n").map(item => item.trim())
    : (normalized.trim() ? [normalized.trim()] : []);
}

function containsText(values: string[], expected: string): boolean {
  return values.some(value => equalsText(value, expected));
}

function distinct(values: string[]): string[] {
  const result: string[] = [];
  values.forEach(value => {
    if (!result.some(existing => equalsText(existing, value))) result.push(value);
  });
  return result;
}

function text(value: unknown): string {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

function lower(value: string): string {
  return value.trim().toLowerCase();
}

function equalsText(left: string, right: string): boolean {
  return lower(left) === lower(right);
}

function yes(value: string): boolean {
  return ["yes", "y", "true", "1", "checked"].includes(lower(value));
}

function writeStatus(setup: ExcelScript.Worksheet, message: string) {
  setup.getRange("B7").setValue(message);
  setup.getRange("B7").getFormat().setWrapText(false);
  setup.getRange("B7").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
  setup.getRange("B7").getFormat().getFont().setColor(message.startsWith("ERROR") ? "#8D021F" : "#000000");
}

function aliasKey(field: string, alias: string): string {
  return `alias|${lower(field)}|${lower(alias)}`;
}

function groupKey(field: string, caption: string, rules: string): string {
  return `group|${lower(field)}|${lower(caption)}|${lower(rules)}`;
}

function valueKey(field: string): string {
  return `value|${lower(field)}`;
}

function filterKey(field: string): string {
  return `filter|${lower(field)}`;
}

function measureKey(conditionField: string, conditionValue: string, valueField: string): string {
  return `measure|${lower(conditionField)}|${lower(conditionValue)}|${lower(valueField)}`;
}
