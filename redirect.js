// Fallback sign-in only (when the Outlook client doesn't support Nested App Auth).
// MSAL v5 popups land here; this hands the result back to the task pane.
import { broadcastResponseToMainFrame } from "@azure/msal-browser/redirect-bridge";

broadcastResponseToMainFrame().catch((e) => {
  document.body.textContent = "שגיאת התחברות: " + (e && e.message ? e.message : e);
});
