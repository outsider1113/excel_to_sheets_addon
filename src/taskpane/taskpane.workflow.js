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

  setStatus("Reading Reciever From Excel...");
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

    setStatus("Fields Copied From Excel. You Can Edit Them Before Sending.");
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



// Build payload for /api/updateMasterFromReceiver from UI + in-memory items
function buildMasterPayloadFromUI(mode) {
  const isCreate = mode === "create";

  // Header inputs
  const pfEl       = $(isCreate ? "c_pf"       : "m_pf");
  const drEl       = $(isCreate ? "c_dr"       : "m_dr");
  const carrierEl  = $(isCreate ? "c_carrier"  : "m_carrier");
  const poEl       = $(isCreate ? "c_po"       : "m_po");
  const rcvEl      = $(isCreate ? "c_rcv"      : "m_rcv");
  const verEl      = $(isCreate ? "c_ver"      : "m_ver");

  const supplier   = pfEl      ? String(pfEl.value || "").trim()      : "";
  const date       = drEl      ? String(drEl.value || "").trim()      : "";
  const carrier    = carrierEl ? String(carrierEl.value || "").trim() : "";
  const poNumber   = poEl      ? String(poEl.value || "").trim()      : "";
  const rrNumber   = rcvEl     ? String(rcvEl.value || "").trim()     : "";
  const verifiedBy = verEl     ? String(verEl.value || "").trim()     : "";

  if (!rrNumber || !poNumber) {
    // Not enough to meaningfully update Master – let caller decide whether to skip
    return null;
  }

  const items = isCreate ? createLineItems : modifyLineItems;
  const lines = [];

  function parseNumber(val) {
    if (val == null) return "";
    const s = String(val).replace(/,/g, "").trim();
    if (!s) return "";
    const n = Number(s);
    return Number.isFinite(n) ? n : "";
  }

  for (const li of items) {
    const hasAny =
      (li.item && li.item.trim().length) ||
      (li.com && li.com.trim().length) ||
      (li.gross && li.gross.trim().length) ||
      (li.tare && li.tare.trim().length) ||
      (li.cost && li.cost.trim().length);

    if (!hasAny) continue;

    const gross = parseNumber(li.gross);
    const tare  = parseNumber(li.tare);
    const net   = (gross !== "" && tare !== "") ? (gross - tare) : "";

    const price = parseNumber(li.cost);

    lines.push({
      material: li.com || "",
      materialNotes: "",      // you can populate later if you add a column
      net: net === "" ? "" : net,
      price: price === "" ? "" : price,
      extension: "",          // let Sheets formulas calculate if you want
      poWeight: ""            // left blank; you adjust in Master
    });
  }

  if (!lines.length) {
    return null;
  }

  return {
    sheetId: selectedSheetId,
    rrNumber,
    date,
    supplier,
    term: "",                 // you can wire these later if you add them to Excel
    dueDate: "",
    daysTillDue: "",
    status: "",
    datePaid: "",
    poNumber,
    poStatus: "",
    poSageClosed: "",
    receiverSageEntry: verifiedBy || "",
    notes: "",                // base notes; carrier will be appended server-side
    carrier,
    lines
  };
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
    setStatus("Excel tab name must include -MM-DD-XXX (e.g. NOVA -1-17-025).");
    return;
  }
  if (workspaceState.type === "existing") {
    setStatus("This receiver already has a workspace in Sheets. Use Modify instead.");
    return;
  }
  if (workspaceState.type === "title-invalid") {
    setStatus("Rename the Excel tab to include -MM-DD-XXX before sending.");
    return;
  }
  if (!hasAnyCreateData()) {
    setStatus("Nothing to send yet. Fill some header or line items first.");
    return;
  }

  const currentName = currentActiveWorksheet;
  if (lastReadCreateSheetName && lastReadCreateSheetName !== currentName) {
    setStatus('Excel sheet changed. Click "Read Data From Excel" again before sending.');
    return;
  }

  createInFlight = true;
  setButtonEnabled($("create_send"), false);
  setGlobalLoading(true, "Creating new workspace and sending data to Google Sheets…");
  setStatus("Sending to Google Sheets…");

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

    // Update Master Receivers (create case)
    const masterPayload = buildMasterPayloadFromUI("create");
    if (masterPayload) {
      try {
        const mrRes = await fetch(`${BACKEND}/api/updateMasterFromReceiver`, {
          method: "POST",
          credentials: "include",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(masterPayload)
        });
        if (!mrRes.ok) {
          console.warn("updateMasterFromReceiver (create) failed:", mrRes.status);
          setStatus(
            "Workspace created, but Master Receivers update failed (" +
            mrRes.status +
            ")."
          );
        } else {
          setStatus("Done: created new workspace and updated Master Receivers.");
        }
      } catch (e) {
        console.warn("updateMasterFromReceiver (create) error:", e);
        setStatus(
          "Workspace created, but Master Receivers update errored: " +
          (e.message || e)
        );
      }
    } else {
      setStatus("Workspace created (Master not updated: missing RR#/PO#/lines).");
    }

    // refresh tabs & workspace state
    await onSheetSelected();
  } catch (err) {
    console.error(err);
    setStatus("Create failed: " + (err.message || err));
  } finally {
    createInFlight = false;
    setGlobalLoading(false);
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
    setStatus("Choose which workspace tab to update.");
    return;
  }
  if (!hasAnyModifyData()) {
    setStatus("Nothing to send yet. Fill some header or line items first.");
    return;
  }
  const currentName = currentActiveWorksheet;
  if (lastReadModifySheetName && lastReadModifySheetName !== currentName) {
    setStatus('Excel sheet changed. Click "Read Data From Excel" again before sending.');
    return;
  }
  if (modifyInFlight) return;

  modifyInFlight = true;
  setButtonEnabled($("modify_send"), false);
  setGlobalLoading(true, "Updating existing workspace in Google Sheets…");
  setStatus("Sending updates to Google Sheets…");

  try {
    const valuesMap = buildValuesMapFromUI("modify");
    await writeFieldsToTab(selectedSheetId, tabName, valuesMap);

    // Update Master Receivers (modify case)
    const masterPayload = buildMasterPayloadFromUI("modify");
    if (masterPayload) {
      try {
        const mrRes = await fetch(`${BACKEND}/api/updateMasterFromReceiver`, {
          method: "POST",
          credentials: "include",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(masterPayload)
        });
        if (!mrRes.ok) {
          console.warn("updateMasterFromReceiver (modify) failed:", mrRes.status);
          setStatus(
            "Workspace updated, but Master Receivers update failed (" +
            mrRes.status +
            ")."
          );
        } else {
          setStatus(
            "Done: updated workspace and Master Receivers for RR#: " +
            masterPayload.rrNumber +
            "."
          );
        }
      } catch (e) {
        console.warn("updateMasterFromReceiver (modify) error:", e);
        setStatus(
          "Workspace updated, but Master Receivers update errored: " +
          (e.message || e)
        );
      }
    } else {
      setStatus("Workspace updated (Master not updated: missing RR#/PO#/lines).");
    }
  } catch (err) {
    console.error(err);
    setStatus("Write failed: " + (err.message || err));
  } finally {
    modifyInFlight = false;
    setGlobalLoading(false);
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
