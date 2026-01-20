/* global Office Excel */
/*
  Main entrypoint:
  - wires up DOM events in Office.onReady
  - starts active worksheet polling with blocking overlay on tab changes
  - routes send button to review modal handled in workflow.js
*/

function debounce_(fn, waitMs) {
  let t = null;
  return (...args) => {
    if (t) clearTimeout(t);
    t = setTimeout(() => fn(...args), waitMs);
  };
}

async function tryRegisterWorksheetActivatedEvent_() {
  // Prefer event-driven updates. If event registration fails (host quirks),
  // we fall back to a very slow poll as a safety net.
  try {
    await Excel.run(async (ctx) => {
      ctx.workbook.worksheets.onActivated.add(async () => {
        try {
          await detectActiveWorksheet(true);
          refreshSendButtonsState();
        } catch (e) {
          console.warn("onActivated handler error", e);
        }
      });
      await ctx.sync();
    });
    return true;
  } catch (e) {
    console.warn("Failed to register worksheet activation event; using fallback poll", e);
    return false;
  }
}

Office.onReady(async () => {
  // auth
  $("googleSignInBtn").addEventListener("click", openAuthDialog);
  $("verifyAuthBtn").addEventListener("click", verifyAuth);

  // sheet reload/select
  $("reloadSheetsBtn").addEventListener("click", loadSheets);
  $("modify_sheet_select").addEventListener("change", onSheetSelected);

  // create panel
  $("create_prefill").addEventListener("click", readFieldsFromExcel_Create);
  $("create_send").addEventListener("click", openSendReviewModal);
  $("create_add_line").addEventListener("click", () => addLineItem());
  $("create_workspace_name").addEventListener("input", refreshSendButtonsState);

  // Key fields: change-driven status refresh (no constant polling)
  const onKeyFieldsChanged = debounce_(async () => {
    try {
      invalidateReceiverRecordStatusCache_();
      await evaluateWorkspaceState();
      refreshSendButtonsState();
    } catch (e) {
      console.warn("key fields refresh error", e);
    }
  }, 350);

  ["c_rcv", "c_po", "c_dr"].forEach(id => {
    const el = $(id);
    if (el) el.addEventListener("input", onKeyFieldsChanged);
  });

  // modal actions
  $("actionModalCancel").addEventListener("click", closeActionModal);
  $("actionModalConfirm").addEventListener("click", confirmPendingSend);

  // Optional manual refresh button – still works as a quick resync
  const refreshBtn = $("refreshActiveSheet");
  if (refreshBtn) {
    refreshBtn.addEventListener("click", async () => {
      await detectActiveWorksheet(true);
      refreshSendButtonsState();
    });
  }

  hide($("mainContent"));
  setStatus("Please sign in to continue.");

  // initial detect
  await detectActiveWorksheet(true).catch(err => console.warn("initial detect error", err));

  // Event-driven updates preferred; fallback poll only if the event can't be registered.
  const ok = await tryRegisterWorksheetActivatedEvent_();
  if (!ok) {
    pollHandle = setInterval(
      () => detectActiveWorksheet().catch(err => console.warn("fallback poll err", err)),
      15000
    );
  }
});
