importScripts('batch-background.js');

const PANEL_PATH = 'panel.html';
const OWNER_STORAGE_KEY = 'sidePanelOwnerByWindow';

async function configureSidePanel() {
  try {
    // The action click is handled explicitly so the panel can be bound to
    // exactly one tab instead of being inherited by the whole Chrome window.
    await chrome.sidePanel.setPanelBehavior({ openPanelOnActionClick: false });
    await chrome.sidePanel.setOptions({ enabled: false });
  } catch (e) {
    console.error('Unable to configure side panel:', e);
  }
}

async function getSidePanelOwners() {
  try {
    const data = await chrome.storage.session.get(OWNER_STORAGE_KEY);
    return data?.[OWNER_STORAGE_KEY] || {};
  } catch (_) {
    return {};
  }
}

async function saveSidePanelOwners(owners) {
  try {
    await chrome.storage.session.set({ [OWNER_STORAGE_KEY]: owners });
  } catch (e) {
    console.warn('Unable to persist side panel owner state:', e);
  }
}

async function bindSidePanelToTab(tab) {
  if (!tab?.id || !tab?.windowId) return;

  const owners = await getSidePanelOwners();
  const windowKey = String(tab.windowId);
  const previousTabId = Number(owners[windowKey] || 0);

  if (previousTabId && previousTabId !== tab.id) {
    try {
      await chrome.sidePanel.setOptions({
        tabId: previousTabId,
        enabled: false
      });
    } catch (_) {
      // The previous tab may already be closed.
    }
  }

  await chrome.sidePanel.setOptions({
    tabId: tab.id,
    path: PANEL_PATH,
    enabled: true
  });

  owners[windowKey] = tab.id;
  await saveSidePanelOwners(owners);

  await chrome.sidePanel.open({ tabId: tab.id });
}

async function forgetSidePanelOwnerForTab(tabId) {
  const owners = await getSidePanelOwners();
  let changed = false;

  for (const [windowKey, ownerTabId] of Object.entries(owners)) {
    if (Number(ownerTabId) === tabId) {
      delete owners[windowKey];
      changed = true;
    }
  }

  if (changed) await saveSidePanelOwners(owners);
}

async function forgetSidePanelOwnerForWindow(windowId) {
  const owners = await getSidePanelOwners();
  const windowKey = String(windowId);
  if (!(windowKey in owners)) return;

  delete owners[windowKey];
  await saveSidePanelOwners(owners);
}

chrome.action.onClicked.addListener(tab => {
  bindSidePanelToTab(tab).catch(e => {
    console.error('Unable to open tab-scoped side panel:', e);
  });
});

chrome.tabs.onRemoved.addListener(tabId => {
  forgetSidePanelOwnerForTab(tabId).catch(() => {});
});

chrome.windows.onRemoved.addListener(windowId => {
  forgetSidePanelOwnerForWindow(windowId).catch(() => {});
});

chrome.runtime.onInstalled.addListener(configureSidePanel);
chrome.runtime.onStartup.addListener(configureSidePanel);
configureSidePanel();
