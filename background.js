importScripts('batch-background.js');

const SIDE_PANEL_PATH = 'panel.html';
const SIDE_PANEL_TAB_KEY = 'sidePanelBoundTabId';

function isSupportedUrl(url) {
  try {
    const u = new URL(url || '');
    if (u.protocol !== 'https:') return false;
    const host = u.hostname.toLowerCase();
    return (
      host === 'teams.microsoft.com' ||
      host.endsWith('.teams.microsoft.com') ||
      host === 'teams.cloud.microsoft' ||
      host.endsWith('.teams.cloud.microsoft') ||
      host.endsWith('.sharepoint.com')
    );
  } catch (_) {
    return false;
  }
}

async function getBoundTabId() {
  try {
    const data = await chrome.storage.session.get(SIDE_PANEL_TAB_KEY);
    return Number.isInteger(data && data[SIDE_PANEL_TAB_KEY]) ? data[SIDE_PANEL_TAB_KEY] : null;
  } catch (_) {
    return null;
  }
}

async function setBoundTabId(tabId) {
  if (Number.isInteger(tabId)) {
    await chrome.storage.session.set({ [SIDE_PANEL_TAB_KEY]: tabId });
  } else {
    await chrome.storage.session.remove(SIDE_PANEL_TAB_KEY);
  }
}

async function disablePanelForTab(tabId) {
  if (!Number.isInteger(tabId)) return;
  try {
    await chrome.sidePanel.setOptions({ tabId, enabled: false });
  } catch (_) {}
}

async function bindSidePanelToTab(tab) {
  if (!tab || !tab.id) throw new Error('Не удалось определить вкладку.');
  if (!isSupportedUrl(tab.url)) {
    throw new Error('Панель доступна только для Microsoft Teams/SharePoint.');
  }

  const previousTabId = await getBoundTabId();
  if (previousTabId !== null && previousTabId !== tab.id) {
    await disablePanelForTab(previousTabId);
  }

  await chrome.sidePanel.setOptions({
    tabId: tab.id,
    path: SIDE_PANEL_PATH,
    enabled: true
  });

  await setBoundTabId(tab.id);
  await chrome.sidePanel.open({ tabId: tab.id });
}

async function configureSidePanel() {
  try {
    await chrome.sidePanel.setPanelBehavior({ openPanelOnActionClick: false });
    await chrome.sidePanel.setOptions({ enabled: false });

    const boundTabId = await getBoundTabId();
    if (boundTabId !== null) {
      try {
        const tab = await chrome.tabs.get(boundTabId);
        if (isSupportedUrl(tab.url)) {
          await chrome.sidePanel.setOptions({
            tabId: boundTabId,
            path: SIDE_PANEL_PATH,
            enabled: true
          });
        } else {
          await setBoundTabId(null);
        }
      } catch (_) {
        await setBoundTabId(null);
      }
    }
  } catch (e) {
    console.error('Unable to configure side panel:', e);
  }
}

chrome.action.onClicked.addListener(async tab => {
  try {
    await bindSidePanelToTab(tab);
  } catch (e) {
    console.error('Unable to open tab-specific side panel:', e);
  }
});

chrome.tabs.onRemoved.addListener(async tabId => {
  const boundTabId = await getBoundTabId();
  if (boundTabId === tabId) await setBoundTabId(null);
});

chrome.tabs.onUpdated.addListener(async (tabId, changeInfo, tab) => {
  if (!changeInfo.url) return;
  const boundTabId = await getBoundTabId();
  if (boundTabId !== tabId) return;
  if (!isSupportedUrl(tab.url)) {
    await disablePanelForTab(tabId);
    await setBoundTabId(null);
  }
});

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message && message.type === 'SIDE_PANEL_GET_BOUND_TAB') {
    (async () => {
      const tabId = await getBoundTabId();
      let tab = null;
      if (tabId !== null) {
        try { tab = await chrome.tabs.get(tabId); } catch (_) {}
      }
      sendResponse({
        ok: true,
        tabId,
        tab: tab ? { id: tab.id, windowId: tab.windowId, url: tab.url || '', title: tab.title || '' } : null
      });
    })();
    return true;
  }
});

chrome.runtime.onInstalled.addListener(configureSidePanel);
chrome.runtime.onStartup.addListener(configureSidePanel);
configureSidePanel();
