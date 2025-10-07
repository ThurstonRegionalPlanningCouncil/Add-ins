(function () {
  const QUARANTINE_URL = 'https://security.microsoft.com/quarantine';
  const LAUNCH_URL = 'https://cdn.jsdelivr.net/gh/ThurstonRegionalPlanningCouncil/Add-ins@main/quarantine/launch.html';

  Office.onReady(() => { /* Office.js ready */ });

  function openQuarantine(event) {
    try {
      if (Office?.context?.ui?.openBrowserWindow) {
        // Classic Outlook supports this
        Office.context.ui.openBrowserWindow(QUARANTINE_URL);
      } else if (Office?.context?.ui?.displayDialogAsync) {
        // New Outlook: open a small dialog hosting our launcher page
        Office.context.ui.displayDialogAsync(
          LAUNCH_URL,
          { height: 40, width: 30, promptBeforeOpen: false, displayInIframe: true },
          function () { /* no-op */ }
        );
      } else {
        // Fallback
        window.open(QUARANTINE_URL, '_blank');
      }
    } finally {
      event?.completed?.();
    }
  }

  // Export globally for the command
  window.openQuarantine = openQuarantine;
})();
