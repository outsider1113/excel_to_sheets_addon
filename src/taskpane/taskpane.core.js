/* global Office, Excel */
/*
  Core for Excel → Sheets add-in:
  - constants & config
  - DOM helpers
  - shared state
  - auth / Sheets helpers
  - active worksheet tracking
  - workspace state helpers
  - line-items helpers
  - build/write helpers
  - basic "has data" predicates
*/

const BACKEND = "https://excel-sheets-backend.onrender.com";

// ----------------- CONSTANTS / CONFIG -----------------

// Header fields shared by Create / Modify.
const HEADER_FIELDS = [
  { key: "pf",  cell: "B11", cId: "c_pf",  mId: "m_pf"  },
  { key: "dr",  cell: "B12", cId: "c_dr",  mId: "m_dr"  },
  { key: "carrier", cell: "B13", cId: "c_carrier", mId: "m_carrier" },
  { key: "po",  cell: "E11", cId: "c_po",  mId: "m_po"  },
  { key: "rcv", cell: "E12", cId: "c_rcv", mId: "m_rcv" },
  { key: "ver", cell: "E13", cId: "c_ver", mId: "m_ver" }
];

// Line item layout (multiple rows).
const ITEM_START_ROW = 18;   // first row for items
const MAX_LINE_ITEMS = 20;   // safety cap
const ITEM_COLUMNS = {
  item:  "A",    // ITEM NO
  com:   "B",    // COMMODITY
  gross: "C",    // GROSS
  tare:  "D",    // TARE
  cost:  "F"     // COST
};
const LINE_KEYS = ["item", "com", "gross", "tare", "cost"];

// Regex to extract id from Excel tab name, e.g. "...-11-02-345"
const ID_RE = /-(\d{1,2})-(\d{1,2})-(\d{3})$/;

// ----------------- DOM HELPERS -----------------

const $ = id => document.getElementById(id);
const show = el => el && el.classList.remove("hidden");
const hide = el => el && el.classList.add("hidden");

function setStatus(txt) {
  const box = $("statusBox");
  if (!box) return;
  if (txt && txt.length) {
    box.textContent = txt;
    show(box);
  } else {
    box.textContent = "";
    hide(box);
  }
}

// Robust button enable/disable
function setButtonEnabled(buttonEl, enabled) {
  if (!buttonEl) return;
  buttonEl.disabled = !enabled;
  if (enabled) {
    buttonEl.classList.remove("disabled");
    buttonEl.setAttribute("aria-disabled", "false");
  } else {
    buttonEl.classList.add("disabled");
    buttonEl.setAttribute("aria-disabled", "true");
  }
}

// Apply/remove locked visual state to a whole section
function setSectionLocked(sectionEl, locked) {
  if (!sectionEl) return;
  if (locked) sectionEl.classList.add("locked-section");
  else sectionEl.classList.remove("locked-section");
}

function setGlobalLoading(isLoading, message) {
  const overlay = $("globalLoadingOverlay");
  const msgEl = $("globalLoadingText");
  if (!overlay || !msgEl) return;

  if (isLoading) {
    if (message) msgEl.textContent = message;
    overlay.classList.remove("hidden");
  } else {
    overlay.classList.add("hidden");
  }
}

// ----------------- RUNTIME STATE -----------------

let pollHandle = null;

let lastAuthOk = false;
let sheetsList = [];
let selectedSheetId = null;
let currentTabs = [];   // tab titles for selected sheet

let currentActiveWorksheet = null;

// Which sheet did we read fields from last, per mode?
let lastReadCreateSheetName = null;
let lastReadModifySheetName = null;

let workspaceState = {
  type: "unknown",      // "existing" | "new-valid" | "title-invalid" | "unknown"
  identifier: null,     // -MM-DD-XXX extracted from Excel
  matchedTab: null      // tab title in Sheets (if any)
};

let createInFlight = false;
let modifyInFlight = false;

// dynamic line items (arrays of { item, com, gross, tare, cost })
let createLineItems = [];
let modifyLineItems = [];

// ----------------- AUTH / SHEETS -----------------

function openAuthDialog() {
  setStatus("Opening sign-in window...");
  Office.context.ui.displayDialogAsync(
    `${BACKEND}/auth`,
    { height: 60, width: 35, displayInIframe: true },
    (res) => {
      if (res.status === Office.AsyncResultStatus.Failed) {
        setStatus("Failed to open sign-in window.");
        return;
      }
      const dlg = res.value;
      dlg.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        if (arg && arg.message) {
          if (arg.message === "auth-success" || arg.message === "success") {
            setStatus("Authentication completed. Click Verify Sign-In.");
          } else {
            setStatus("Auth message: " + arg.message);
          }
        }
        try { dlg.close(); } catch (e) {}
      });
    }
  );
}

async function verifyAuth() {
  setStatus("Verifying sign-in...");
  try {
    const res = await fetch(`${BACKEND}/api/sheets`, { credentials: "include" });
    if (res.status === 200) {
      lastAuthOk = true;
      const json = await res.json();
      sheetsList = json.files || json || [];
      show($("mainContent"));
      populateSheetsSelect();
      $("authStatus").textContent = "Signed in.";
      setStatus("Signed in. Select a Google Sheets file to work inside.");

      await detectActiveWorksheet();
      refreshSendButtonsState();
    } else if (res.status === 401) {
      lastAuthOk = false;
      hide($("mainContent"));
      setStatus("Not signed in. Please complete sign-in.");
      $("authStatus").textContent = "Not signed in.";
    } else {
      lastAuthOk = false;
      hide($("mainContent"));
      setStatus("Auth verify returned: " + res.status);
      $("authStatus").textContent = "Auth error.";
    }
  } catch (err) {
    console.error(err);
    lastAuthOk = false;
    hide($("mainContent"));
    setStatus("Auth check failed: " + (err.message || err));
    $("authStatus").textContent = "Auth failed.";
  }
}

async function loadSheets() {
  if (!lastAuthOk) {
    setStatus("Please verify sign-in first.");
    return;
  }
  setStatus("Reloading Sheets list...");
  try {
    const res = await fetch(`${BACKEND}/api/sheets`, { credentials: "include" });
    if (!res.ok) throw new Error("Failed to load sheets: " + res.status);
    const json = await res.json();
    sheetsList = json.files || json || [];
    populateSheetsSelect();
    setStatus("Sheets list reloaded.");
  } catch (err) {
    console.error(err);
    setStatus("Reload failed: " + (err.message || err));
  }
}

function populateSheetsSelect() {
  const sel = $("modify_sheet_select");
  if (!sel) return;
  sel.innerHTML = "<option value=''>— choose sheet file —</option>";
  sheetsList.forEach(f => {
    const o = document.createElement("option");
    o.value = f.id;
    o.textContent = f.name;
    sel.appendChild(o);
  });
}

async function onSheetSelected() {
  selectedSheetId = $("modify_sheet_select").value || null;
  const tabSel = $("modify_tab_select");
  tabSel.innerHTML = "<option value=''>— choose tab —</option>";
  currentTabs = [];
  workspaceState = { type: "unknown", identifier: null, matchedTab: null };

  if (!selectedSheetId) {
    setStatus("Select a Google Sheets file to continue.");
    setSectionLocked($("createPanel"), true);
    setSectionLocked($("modifyPanel"), true);
    refreshSendButtonsState();
    return;
  }

  setStatus("Loading tabs for the selected file...");
  try {
    const res = await fetch(
      `${BACKEND}/api/tabs?sheetId=${encodeURIComponent(selectedSheetId)}`,
      { credentials: "include" }
    );
    if (!res.ok) {
      setStatus("Failed to load tabs: " + res.status);
      refreshSendButtonsState();
      return;
    }
    const data = await res.json();
    currentTabs = (data || []).map(t =>
      (t.properties && t.properties.title) ? t.properties.title : (t.title || t)
    );

    tabSel.innerHTML = "<option value=''>— choose tab —</option>";
    currentTabs.forEach(t => {
      const o = document.createElement("option");
      o.value = t;
      o.textContent = t;
      tabSel.appendChild(o);
    });

    // Evaluate workspace state based on current Excel tab + loaded tabs.
    // Evaluate workspace state based on current Excel tab + loaded tabs.
    await evaluateWorkspaceState();
    updateModeBanner();
    updateResyncWarning();


    // If this Excel tab already has a workspace in Sheets, auto-switch to Modify.
    if (workspaceState.type === "existing" && workspaceState.matchedTab) {
      $("modify_tab_select").value = workspaceState.matchedTab;
      switchMode("modify");
      setStatus(
        `Workspace for this Excel tab already exists as "${workspaceState.matchedTab}". ` +
        "Create is locked. Use Modify instead."
      );
    } else if (workspaceState.type === "title-invalid") {
      // evaluateWorkspaceState already set a detailed message about title format.
    } else {
      setStatus("Tabs loaded. Choose Create or Modify.");
    }

    // Sections are now unlocked (but buttons still controlled by refreshSendButtonsState)
    setSectionLocked($("createPanel"), false);
    setSectionLocked($("modifyPanel"), false);
  } catch (err) {
    console.error(err);
    setStatus("Error loading tabs: " + (err.message || err));
  } finally {
    refreshSendButtonsState();
  }
}

// ----------------- ACTIVE WORKSHEET TRACKING -----------------

async function detectActiveWorksheet() {
  try {
    await Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getActiveWorksheet();
      ws.load("name");
      await ctx.sync();

      const currentName = ws.name || "";
      currentActiveWorksheet = currentName;

      const nameSpan = $("activeSheetName");
      if (nameSpan) nameSpan.textContent = currentName;

      // Always mirror active Excel tab into the workspace-name field
      // Always mirror active Excel tab into the workspace-name field
      updateCreateWorkspaceNameFromActive(currentName);

      // Re-evaluate workspace existence and warnings on each poll
      await evaluateWorkspaceState();
      updateModeBanner();
      updateResyncWarning();
      refreshSendButtonsState();
    });
  } catch (err) {
    console.warn("detectActiveWorksheet error", err);
  }
}

// Update create workspace name from active Excel tab (always overwrite)
function updateCreateWorkspaceNameFromActive(activeName) {
  const input = $("create_workspace_name");
  if (input && activeName) {
    input.value = activeName;
  }
}

// ----------------- WORKSPACE STATE / WARNINGS -----------------

async function evaluateWorkspaceState() {
  const activeName = currentActiveWorksheet ||
    ($("activeSheetName") ? $("activeSheetName").textContent : "");
  const identifier = extractIdentifierFromName(activeName);
  workspaceState = { type: "unknown", identifier: null, matchedTab: null };

  if (!activeName) {
    workspaceState.type = "unknown";
    return;
  }

  if (!identifier) {
    // title doesn't contain -MM-DD-XXX at the end
    workspaceState.type = "title-invalid";
    workspaceState.identifier = null;
    if (selectedSheetId) {
      setStatus(
        "Active Excel tab name does not contain -MM-DD-XXX identifier. " +
        "You can read values, but Create is disabled until the tab title is fixed."
      );
    }
    return;
  }

  workspaceState.identifier = identifier;
  const matched = findTabByIdentifier(identifier);
  if (matched) {
    workspaceState.type = "existing";
    workspaceState.matchedTab = matched;
  } else {
    workspaceState.type = "new-valid";
  }
}

// extract identifier -MM-DD-XXX from name
function extractIdentifierFromName(name) {
  if (!name || typeof name !== "string") return null;
  const m = name.match(ID_RE);
  if (!m) return null;
  return `-${m[1]}-${m[2]}-${m[3]}`;
}

// Attempt to find a tab whose title contains the identifier (in a few padded forms)
function findTabByIdentifier(identifier) {
  if (!identifier || !currentTabs || !currentTabs.length) return null;
  const parts = identifier.split("-");
  // parts: ["", MM, DD, XXX]
  const mm = parts[1];
  const dd = parts[2];
  const xxx = parts[3];

  const mmP = mm.padStart(2, "0");
  const ddP = dd.padStart(2, "0");

  const candidates = new Set();
  candidates.add(identifier);                // -M-D-XXX
  candidates.add(`-${mmP}-${dd}-${xxx}`);
  candidates.add(`-${mm}-${ddP}-${xxx}`);
  candidates.add(`-${mmP}-${ddP}-${xxx}`);
  candidates.add(`${mm}-${dd}-${xxx}`);
  candidates.add(`${mmP}-${ddP}-${xxx}`);

  for (const t of currentTabs) {
    const lower = t.toLowerCase();
    for (const idf of candidates) {
      if (!idf) continue;
      if (lower.includes(idf.toLowerCase())) return t;
    }
  }
  return null;
}



function updateModeBanner() {
  const banner = $("modeBanner");
  if (!banner) return;

  const activeName =
    currentActiveWorksheet ||
    ($("activeSheetName") ? $("activeSheetName").textContent : "");

  // If we don't know the Excel tab yet or no sheet selected, hide
  if (!activeName || !selectedSheetId) {
    banner.className = "mode-banner hidden";
    banner.querySelector(".text").textContent = "Waiting for Excel tab...";
    return;
  }

  const labelEl = banner.querySelector(".label");
  const textEl = banner.querySelector(".text");

  if (!labelEl || !textEl) return;

  if (workspaceState.type === "title-invalid") {
    banner.className = "mode-banner invalid";
    labelEl.textContent = "Mode";
    textEl.textContent =
      "Cannot send yet. Rename the Excel tab to include -MM-DD-XXX.";
  } else if (workspaceState.type === "existing" && workspaceState.matchedTab) {
    banner.className = "mode-banner modify";
    labelEl.textContent = "Modifying";
    textEl.textContent = `Will update existing tab: "${workspaceState.matchedTab}" in Google Sheets.`;
  } else if (workspaceState.type === "new-valid") {
    banner.className = "mode-banner create";
    labelEl.textContent = "Creating";
    textEl.textContent = `Will create new tab from: "${activeName}" in Google Sheets.`;
  } else {
    banner.className = "mode-banner";
    labelEl.textContent = "Mode";
    textEl.textContent = "Waiting for workspace status…";
  }
}

// Warning about “data read from a different sheet”
function updateResyncWarning() {
  const warningEl = $("sheetChangeWarning");
  if (!warningEl) return;

  const currentName = currentActiveWorksheet ||
    ($("activeSheetName") ? $("activeSheetName").textContent : "");

  let showWarning = false;
  let msg = "";

  if (hasAnyCreateData() &&
      lastReadCreateSheetName &&
      lastReadCreateSheetName !== currentName) {
    showWarning = true;
    msg = `Excel sheet changed. Click "Read Data From Excel" again before sending.`;
  } else if (hasAnyModifyData() &&
             lastReadModifySheetName &&
             lastReadModifySheetName !== currentName) {
    showWarning = true;
    msg = `Excel sheet changed. Click "Read Data From Excel" again before sending.`;
  }


  if (showWarning) {
    warningEl.textContent = msg;
    show(warningEl);
  } else {
    warningEl.textContent = "";
    hide(warningEl);
  }
}

// ----------------- MULTI-ROW LINE ITEMS HELPERS -----------------

function addLineItem(mode) {
  const isCreate = mode === "create";
  const items = isCreate ? createLineItems : modifyLineItems;
  items.push({ item: "", com: "", gross: "", tare: "", cost: "" });
  renderLineItems(mode);
  refreshSendButtonsState();
}

function renderLineItems(mode) {
  const isCreate = mode === "create";
  const container = isCreate ? $("createLineItemsContainer") : $("modifyLineItemsContainer");
  const items = isCreate ? createLineItems : modifyLineItems;

  if (!container) return;
  container.innerHTML = "";

  if (!items.length) {
    const p = document.createElement("p");
    p.className = "status-msg";
    p.textContent = "No line items detected. Use \"Add Line\" to add one.";
    container.appendChild(p);
    return;
  }

  items.forEach((li, idx) => {
    const row = document.createElement("div");
    row.className = "line-item-row";
    row.dataset.index = String(idx);

    LINE_KEYS.forEach(key => {
      const input = document.createElement("input");
      input.value = li[key] || "";
      input.addEventListener("input", () => {
        li[key] = input.value;
        refreshSendButtonsState();
      });
      row.appendChild(input);
    });

    const removeBtn = document.createElement("button");
    removeBtn.type = "button";
    removeBtn.className = "secondary-btn small-btn";
    removeBtn.textContent = "×";
    removeBtn.addEventListener("click", () => {
      items.splice(idx, 1);
      renderLineItems(mode);
      refreshSendButtonsState();
    });
    row.appendChild(removeBtn);

    container.appendChild(row);
  });
}

// ----------------- BUILD VALUES MAP FROM UI -----------------

function buildValuesMapFromUI(mode) {
  const isCreate = mode === "create";

  const valuesMap = {};

  // headers
  HEADER_FIELDS.forEach(h => {
    const id = isCreate ? h.cId : h.mId;
    const el = $(id);
    valuesMap[h.cell] = el ? el.value || "" : "";
  });

  // line items
  const items = isCreate ? createLineItems : modifyLineItems;
  items.forEach((li, idx) => {
    const row = ITEM_START_ROW + idx;
    const any =
      (li.item && li.item.trim().length) ||
      (li.com && li.com.trim().length) ||
      (li.gross && li.gross.trim().length) ||
      (li.tare && li.tare.trim().length) ||
      (li.cost && li.cost.trim().length);

    if (!any) return;

    valuesMap[`${ITEM_COLUMNS.item}${row}`]  = li.item  || "";
    valuesMap[`${ITEM_COLUMNS.com}${row}`]   = li.com   || "";
    valuesMap[`${ITEM_COLUMNS.gross}${row}`] = li.gross || "";
    valuesMap[`${ITEM_COLUMNS.tare}${row}`]  = li.tare  || "";
    valuesMap[`${ITEM_COLUMNS.cost}${row}`]  = li.cost  || "";
  });

  return valuesMap;
}

// ----------------- WRITE TO SHEETS (BATCH + FALLBACK) -----------------

async function writeFieldsToTab(sheetId, tabName, valuesMap) {
  const batch = Object.entries(valuesMap).map(([cell, value]) => ({ cell, value }));

  // try batch first
  try {
    const res = await fetch(`${BACKEND}/api/writeBatch`, {
      method: "POST",
      credentials: "include",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ sheetId, tabName, values: batch })
    });
    if (res.ok) return;

    // fallback to individual calls
    for (const it of batch) {
      const r = await fetch(`${BACKEND}/api/writeCell`, {
        method: "POST",
        credentials: "include",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          sheetId,
          tabName,
          cell: it.cell,
          value: it.value
        })
      });
      if (!r.ok) throw new Error("writeCell failed: " + r.status);
    }
  } catch (err) {
    throw err;
  }
}

// ----------------- “HAS DATA” HELPERS -----------------

function hasAnyHeaderValues(prefix) {
  // prefix: "c_" or "m_"
  for (const h of HEADER_FIELDS) {
    const id = prefix === "c_" ? h.cId : h.mId;
    const el = $(id);
    if (el && String(el.value || "").trim().length) return true;
  }
  return false;
}

function hasAnyLineItems(mode) {
  const items = mode === "create" ? createLineItems : modifyLineItems;
  return items.some(li => {
    return (
      (li.item && li.item.trim().length) ||
      (li.com && li.com.trim().length) ||
      (li.gross && li.gross.trim().length) ||
      (li.tare && li.tare.trim().length) ||
      (li.cost && li.cost.trim().length)
    );
  });
}

function hasAnyCreateData() {
  return hasAnyHeaderValues("c_") || hasAnyLineItems("create");
}

function hasAnyModifyData() {
  return hasAnyHeaderValues("m_") || hasAnyLineItems("modify");
}
