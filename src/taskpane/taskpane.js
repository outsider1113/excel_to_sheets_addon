/* global Office */
/*
  Main entrypoint:
  - wires up DOM events in Office.onReady
  - starts active worksheet polling with blocking overlay on tab changes
  - routes send button to review modal handled in workflow.js
*/

Office.onReady(() => {
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

  // initial detect and polling for changes
  detectActiveWorksheet(true).catch(err => console.warn("initial detect error", err));
  pollHandle = setInterval(
    () => detectActiveWorksheet().catch(err => console.warn("poll err", err)),
    2000
  );
});
