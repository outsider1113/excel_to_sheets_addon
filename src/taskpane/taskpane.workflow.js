/* global Excel */
/*
  Workflow layer:
  - read from Excel for Create / Modify
  - create / modify workspace flows
  - workspace name validation
  - button enable / lock-out logic
*/

// ----------------- READ FROM EXCEL -----------------

async function readFieldsFromExcel_Create() {
  if (!selectedSheetId) {
    setStatus("Select a Google Sheets file first.");
    return;
  }
  if (createInFlight) return;

  setStatus("Reading fields from Excel (Create)...");
  try {
    await Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getActiveWorksheet();
      ws.load("name");

      // Header cells
      const headerRanges = {};
      HEADER_FIELDS.forEach(h => {
        const rng = ws.getRange(h.cell);
        rng.load("values");
        headerRanges[h.key] = rng;
      });

      // Line items block (A18:F...)
      const endRow = ITEM_START_ROW + MAX_LINE_ITEMS - 1;
      const lineRange = ws.getRange(`A${ITEM_START_ROW}:F${endRow}`);
      lineRange.load("values");

      await ctx.sync();

      // Fill header inputs
      HEADER_FIELDS.forEach(h => {
        const el = $(h.cId);
        const rng = headerRanges[h.key];
        const v = rng && rng.values && rng.values[0] ? rng.values[0][0] : "";
        if (el) el.value = v == null ? "" : String(v);
      });

      // Track where these Create values came from
      lastReadCreateSheetName = ws.name || "";

      // Line items
      createLineItems = [];
      const values = lineRange.values || [];
      for (let i = 0; i < values.length; i++) {
        const row = values[i] || [];
        const itemVal  = row[0];
        const comVal   = row[1];
        const grossVal = row[2];
        const tareVal  = row[3];
        const costVal  = row[5]; // F column

        const any =
          (itemVal && String(itemVal).trim().length) ||
          (comVal && String(comVal).trim().length) ||
          (grossVal && String(grossVal).trim().length) ||
          (tareVal && String(tareVal).trim().length) ||
          (costVal && String(costVal).trim().length);

        if (any) {
          createLineItems.push({
            item:  itemVal  == null ? "" : String(itemVal),
            com:   comVal   == null ? "" : String(comVal),
            gross: grossVal == null ? "" : String(grossVal),
            tare:  tareVal  == null ? "" : String(tareVal),
            cost:  costVal  == null ? "" : String(costVal)
          });
        }
      }

      renderLineItems("create");
    });

    setStatus("Fields populated from Excel (Create).");
  } catch (err) {
    console.error(err);
    setStatus("Failed to read Excel fields: " + (err.message || err));
  } finally {
    updateResyncWarning();
    refreshSendButtonsState();
  }
}

async function readFieldsFromExcel_Modify() {
  if (!selectedSheetId) {
    setStatus("Select a Google Sheets file first.");
    return;
  }
  if (modifyInFlight) return;

  setStatus("Reading fields from Excel (Modify)...");
  try {
    await Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getActiveWorksheet();
      ws.load("name");
      // Header cells
      const headerRanges = {};
      HEADER_FIELDS.forEach(h => {
        const rng = ws.getRange(h.cell);
        rng.load("values");
        headerRanges[h.key] = rng;
      });

      // Line items
      const endRow = ITEM_START_ROW + MAX_LINE_ITEMS - 1;
      const lineRange = ws.getRange(`A${ITEM_START_ROW}:F${endRow}`);
      lineRange.load("values");

      await ctx.sync();

      HEADER_FIELDS.forEach(h => {
        const el = $(h.mId);
        const rng = headerRanges[h.key];
        const v = rng && rng.values && rng.values[0] ? rng.values[0][0] : "";
        if (el) el.value = v == null ? "" : String(v);
      });

      // Track where Modify values came from
      lastReadModifySheetName = ws.name || "";

      modifyLineItems = [];
      const values = lineRange.values || [];
      for (let i = 0; i < values.length; i++) {
        const row = values[i] || [];
        const itemVal  = row[0];
        const comVal   = row[1];
        const grossVal = row[2];
        const tareVal  = row[3];
        const costVal  = row[5];

        const any =
          (itemVal && String(itemVal).trim().length) ||
          (comVal && String(comVal).trim().length) ||
          (grossVal && String(grossVal).trim().length) ||
          (tareVal && String(tareVal).trim().length) ||
          (costVal && String(costVal).trim().length);

        if (any) {
          modifyLineItems.push({
            item:  itemVal  == null ? "" : String(itemVal),
            com:   comVal   == null ? "" : String(comVal),
            gross: grossVal == null ? "" : String(grossVal),
            tare:  tareVal  == null ? "" : String(tareVal),
            cost:  costVal  == null ? "" : String(costVal)
          });
        }
      }

      renderLineItems("modify");
    });

    setStatus("Fields populated from Excel (Modify).");
  } catch (err) {
    console.error(err);
    setStatus("Failed to read Excel fields: " + (err.message || err));
  } finally {
    updateResyncWarning();
    refreshSendButtonsState();
  }
}

// ----------------- CREATE / MODIFY WORKSPACE -----------------

function isWorkspaceNameFormatValid(name) {
  if (!name) return false;
  return ID_RE.test(name.trim());
}

async function createWorkspaceAndSend() {
  if (!selectedSheetId) {
    setStatus("Select a Google Sheets file first.");
    return;
  }
  if (createInFlight) return;

  const wsNameInput = $("create_workspace_name");
  const name = (wsNameInput && wsNameInput.value || "").trim();

  if (!name) {
    setStatus("Workspace name is empty. Rename the Excel tab so it includes -MM-DD-XXX.");
    return;
  }
  if (!isWorkspaceNameFormatValid(name)) {
    setStatus("Workspace name must include identifier -MM-DD-XXX (day may be single digit).");
    return;
  }
  if (workspaceState.type === "existing") {
    setStatus("A workspace for this Excel tab already exists. Use Modify instead.");
    return;
  }
  if (workspaceState.type === "title-invalid") {
    setStatus("Excel tab title does not contain -MM-DD-XXX. Fix the title, then try Create again.");
    return;
  }
  if (!hasAnyCreateData()) {
    setStatus("Nothing to send. Fill at least one header or line item field.");
    return;
  }

  const currentName = currentActiveWorksheet;
  if (lastReadCreateSheetName && lastReadCreateSheetName !== currentName) {
    setStatus(
      `Create data was read from "${lastReadCreateSheetName}" but active sheet is ` +
      `"${currentName}". Re-read data from Excel before sending.`
    );
    return;
  }

  createInFlight = true;
  setButtonEnabled($("create_send"), false);
  setStatus("Creating workspace (copying template) and sending fields...");

  try {
    const templateName = "CMX METAL NEW TEMPLATE";

    const cr = await fetch(`${BACKEND}/api/createTab`, {
      method: "POST",
      credentials: "include",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ sheetId: selectedSheetId, templateName, newName: name })
    });
    if (!cr.ok) throw new Error("createTab failed: " + cr.status);

    const valuesMap = buildValuesMapFromUI("create");
    await writeFieldsToTab(selectedSheetId, name, valuesMap);

    setStatus("Workspace created and initialized.");
    // refresh tabs & workspace state
    await onSheetSelected();
  } catch (err) {
    console.error(err);
    setStatus("Create failed: " + (err.message || err));
  } finally {
    createInFlight = false;
    refreshSendButtonsState();
  }
}

async function modifySendToSheets() {
  if (!selectedSheetId) {
    setStatus("Select a Google Sheets file first.");
    return;
  }
  const tabName = $("modify_tab_select").value;
  if (!tabName) {
    setStatus("Select a workspace tab to modify.");
    return;
  }
  if (!hasAnyModifyData()) {
    setStatus("Nothing to send. Fill at least one header or line item field.");
    return;
  }
  const currentName = currentActiveWorksheet;
  if (lastReadModifySheetName && lastReadModifySheetName !== currentName) {
    setStatus(
      `Modify data was read from "${lastReadModifySheetName}" but active sheet is ` +
      `"${currentName}". Re-read data from Excel before sending.`
    );
    return;
  }
  if (modifyInFlight) return;

  modifyInFlight = true;
  setButtonEnabled($("modify_send"), false);
  setStatus("Writing values to Sheets...");

  try {
    const valuesMap = buildValuesMapFromUI("modify");
    await writeFieldsToTab(selectedSheetId, tabName, valuesMap);
    setStatus("Updated Sheets tab: " + tabName);
  } catch (err) {
    console.error(err);
    setStatus("Write failed: " + (err.message || err));
  } finally {
    modifyInFlight = false;
    refreshSendButtonsState();
  }
}

// ----------------- BUTTON STATE / LOCK-OUT LOGIC -----------------

function refreshSendButtonsState() {
  // default: everything off
  let createEnabled = false;
  let modifyEnabled = false;

  // Pre-conditions: auth + sheet selected
  if (!lastAuthOk) {
    setButtonEnabled($("create_send"), false);
    setButtonEnabled($("modify_send"), false);
    setSectionLocked($("createPanel"), true);
    setSectionLocked($("modifyPanel"), true);
    return;
  }

  setSectionLocked($("createPanel"), !selectedSheetId);
  setSectionLocked($("modifyPanel"), !selectedSheetId);

  if (!selectedSheetId) {
    setButtonEnabled($("create_send"), false);
    setButtonEnabled($("modify_send"), false);
    return;
  }

  // lock create completely if Excel tab already has a workspace in Sheets
  if (workspaceState.type === "existing") {
    setSectionLocked($("createPanel"), true);
    setButtonEnabled($("create_send"), false);
  } else {
    setSectionLocked($("createPanel"), false);
  }

  // CREATE: workspace name must be non-empty and valid format
  const wsName = ($("create_workspace_name") && $("create_workspace_name").value || "").trim();
  const nameValid = isWorkspaceNameFormatValid(wsName);

  // Resync requirement for Create
  const currentName = currentActiveWorksheet;
  const createNeedsResync =
    hasAnyCreateData() &&
    lastReadCreateSheetName &&
    currentName &&
    lastReadCreateSheetName !== currentName;

  // allow prefill regardless of format, but send only when valid & resynced
  setButtonEnabled($("create_prefill"), !!selectedSheetId && !createInFlight);

  if (!createInFlight &&
      selectedSheetId &&
      workspaceState.type !== "existing" &&
      workspaceState.type !== "title-invalid" &&
      nameValid &&
      hasAnyCreateData() &&
      !createNeedsResync) {
    createEnabled = true;
  }

  // MODIFY: need selected tab + some data + resynced
  const selectedTab = $("modify_tab_select").value;
  const modifyNeedsResync =
    hasAnyModifyData() &&
    lastReadModifySheetName &&
    currentName &&
    lastReadModifySheetName !== currentName;

  if (!modifyInFlight &&
      selectedSheetId &&
      selectedTab &&
      hasAnyModifyData() &&
      !modifyNeedsResync) {
    modifyEnabled = true;
  }

  // Respect in-flight flags (never re-enable mid-request)
  if (!createInFlight) {
    setButtonEnabled($("create_send"), createEnabled);
  }
  if (!modifyInFlight) {
    setButtonEnabled($("modify_send"), modifyEnabled);
  }
}
