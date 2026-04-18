/**
 * Navegação explícita na aba — mais confiável que location.assign
 * dentro do content script no SPA do WhatsApp Web.
 * Foco da aba/janela — necessário para navigator.clipboard.* no content script
 * (ex.: após “Iniciar” no popup o documento do WA não fica focado).
 *
 * Lease de disparo: várias abas web.whatsapp.com recebem o mesmo storage.onChanged
 * e cada uma executava runOneContact → 3× mensagem. Só a aba que obtém o lease envia.
 */
const DISPATCH_LEASE_MS = 4 * 60 * 1000;
let dispatchLeaseTabId = null;
let dispatchLeaseUntil = 0;

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (!msg) return false;

  if (msg.type === "WPP_DISPATCH_LEASE_TRY") {
    const tabId = sender.tab?.id;
    const now = Date.now();
    if (!tabId) {
      sendResponse({ ok: false, reason: "no-tab" });
      return false;
    }
    if (dispatchLeaseUntil > now && dispatchLeaseTabId !== tabId) {
      sendResponse({ ok: false, reason: "leased" });
      return false;
    }
    dispatchLeaseTabId = tabId;
    dispatchLeaseUntil = now + DISPATCH_LEASE_MS;
    sendResponse({ ok: true });
    return false;
  }

  if (msg.type === "WPP_DISPATCH_LEASE_RELEASE") {
    const tabId = sender.tab?.id;
    if (tabId && dispatchLeaseTabId === tabId) {
      dispatchLeaseTabId = null;
      dispatchLeaseUntil = 0;
    }
    sendResponse({ ok: true });
    return false;
  }

  if (msg.type === "WPP_FOCUS_TAB") {
    const tabId = sender.tab?.id;
    const windowId = sender.tab?.windowId;
    if (!tabId) {
      sendResponse({ ok: false, reason: "no-tab" });
      return false;
    }
    chrome.tabs
      .update(tabId, { active: true })
      .then(() =>
        windowId != null
          ? chrome.windows.update(windowId, { focused: true })
          : Promise.resolve()
      )
      .then(() => sendResponse({ ok: true }))
      .catch((e) => sendResponse({ ok: false, reason: String(e) }));
    return true;
  }

  if (msg.type !== "WPP_NAVIGATE_TAB" || !msg.url) return false;

  const tabId = msg.tabId ?? sender.tab?.id;
  if (!tabId) {
    sendResponse({ ok: false, reason: "no-tab" });
    return false;
  }

  chrome.tabs
    .update(tabId, { url: msg.url })
    .then(() => sendResponse({ ok: true }))
    .catch((e) => sendResponse({ ok: false, reason: String(e) }));
  return true;
});
