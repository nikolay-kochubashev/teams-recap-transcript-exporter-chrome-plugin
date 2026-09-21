importScripts('batch-background.js');

async function configureSidePanel() {
  try {
    await chrome.sidePanel.setPanelBehavior({ openPanelOnActionClick: true });
  } catch (e) {
    console.error('Unable to configure side panel:', e);
  }
}

chrome.runtime.onInstalled.addListener(configureSidePanel);
chrome.runtime.onStartup.addListener(configureSidePanel);
configureSidePanel();
