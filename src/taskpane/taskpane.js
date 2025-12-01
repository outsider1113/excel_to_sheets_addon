/* global Office */
/*
  Main entrypoint:
  - defines switchMode
  - wires up DOM events in Office.onReady
  - starts active worksheet polling
*/

// ----------------- MODE SWITCHING -----------------

function switchMode(mode) {
  if (mode === "create") {
    $("tabCreate").classList.add("active");
    $("tabModify").classList.remove("active");
    show($("createPanel"));
    hide($("modifyPanel"));
  } else {
    $("tabModify").classList.add("active");
    $("tabCreate").classList.remove("active");
    show($("modifyPanel"));
    hide($("createPanel"));
  }
  refreshSendButtonsState();
}

// ----------------- INIT / EVENT WIRING -----------------

Office.onReady(() => {
  // auth
  $("googleSignInBtn").addEventListener("click", openAuthDialog);
  $("verifyAuthBtn").addEventListener("click", verifyAuth);

  // sheet reload/select
  $("reloadSheetsBtn").addEventListener("click", loadSheets);
  $("modify_sheet_select").addEventListener("change", onSheetSelected);

  // mode tabs
  $("tabCreate").addEventListener("click", () => switchMode("create"));
  $("tabModify").addEventListener("click", () => switchMode("modify"));

  // create panel
  $("create_prefill").addEventListener("click", readFieldsFromExcel_Create);
  $("create_send").addEventListener("click", createWorkspaceAndSend);
  $("create_add_line").addEventListener("click", () => addLineItem("create"));
  $("create_workspace_name").addEventListener("input", () => {
    // Only affects format validation; existence is based on Excel tab id.
    refreshSendButtonsState();
  });

  // modify panel
  $("modify_prefill").addEventListener("click", readFieldsFromExcel_Modify);
  $("modify_send").addEventListener("click", modifySendToSheets);
  $("modify_add_line").addEventListener("click", () => addLineItem("modify"));
  $("modify_tab_select").addEventListener("change", refreshSendButtonsState);

  // Optional manual refresh button – still works as a quick resync
  const refreshBtn = $("refreshActiveSheet");
  if (refreshBtn) {
    refreshBtn.addEventListener("click", async () => {
      await detectActiveWorksheet();
      refreshSendButtonsState();
    });
  }

  hide($("mainContent"));
  setStatus("Please sign in to continue.");

  // initial detect and polling for changes
  detectActiveWorksheet().catch(err => console.warn("initial detect error", err));
  pollHandle = setInterval(
    () => detectActiveWorksheet().catch(err => console.warn("poll err", err)),
    2000
  );
});


//For scripting add in two files into the excel sheet and then aggregate them over time the