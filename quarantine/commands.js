(function () {
  const QUARANTINE_URL = 'https://security.microsoft.com/quarantine';

  Office.onReady(() => {
    // Office.js ready
  });

  function openQuarantine(event) {
    try {
      if (Office?.context?.ui?.openBrowserWindow) {
        Office.context.ui.openBrowserWindow(QUARANTINE_URL);
      } else {
        window.open(QUARANTINE_URL, '_blank');
      }
    } finally {
      if (event && typeof event.completed === 'function') {
        event.completed();
      }
    }
  }

  window.openQuarantine = openQuarantine;
})();

