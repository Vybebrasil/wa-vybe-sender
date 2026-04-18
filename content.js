/* WhatsApp Web — disparo humanizado, intervalos configuráveis, relatório CSV via storage */

const K = {
  queue: "wppQueue",
  template: "wppTemplate",
  index: "wppIndex",
  active: "wppActive",
  paused: "wppPaused",
  stats: "wppStats",
  logs: "wppLogs",
  results: "wppResults",
  minSec: "wppMinIntervalSec",
  maxSec: "wppMaxIntervalSec",
  navPending: "wppNavPending",
  antiBlocks: "wppAntiBlocks",
  antiBlockSize: "wppAntiBlockSize",
  antiBlockWait: "wppAntiBlockWait",
  antiRandom: "wppAntiRandom",
  antiIntervalOn: "wppAntiIntervalOn",
  batchCount: "wppBatchSentInBlock",
  pendingAttachment: "wppPendingAttachment",
  imageCaption: "wppImageCaption",
  enableCaption: "wppEnableCaption",
};

const NAV_PENDING_TTL_MS = 60000;

const delay = (ms) => new Promise((r) => setTimeout(r, ms));

/** Após atualizar a extensão, o script antigo no WA fica inválido — evita rejeições não tratadas. */
function isExtensionContextAlive() {
  try {
    return Boolean(chrome.runtime && chrome.runtime.id);
  } catch {
    return false;
  }
}

/** Erro síncrono/assíncrono quando o content script ficou órfão após reload da extensão. */
function isExtensionContextInvalidatedError(err) {
  if (!err) return false;
  const msg = typeof err.message === "string" ? err.message : String(err);
  return msg.includes("Extension context invalidated");
}

async function safeLocalGet(keys) {
  if (!isExtensionContextAlive()) return {};
  try {
    return await chrome.storage.local.get(keys);
  } catch {
    return {};
  }
}

async function safeLocalSet(obj) {
  if (!isExtensionContextAlive()) return;
  try {
    await chrome.storage.local.set(obj);
  } catch {
    /* contexto invalidado, quota, etc. */
  }
}

async function safeLocalRemove(keys) {
  if (!isExtensionContextAlive()) return;
  try {
    await chrome.storage.local.remove(keys);
  } catch {
    /* */
  }
}

/** Navegação via background (tabs.update) — fallback se o SW não responder */
function navigateToSendUrl(url) {
  return new Promise((resolve) => {
    const fallbackNav = () => {
      try {
        window.top.location.assign(url);
      } catch {
        window.location.assign(url);
      }
      resolve(false);
    };
    if (!isExtensionContextAlive()) {
      fallbackNav();
      return;
    }
    try {
      chrome.runtime.sendMessage({ type: "WPP_NAVIGATE_TAB", url }, (response) => {
        if (chrome.runtime.lastError) {
          fallbackNav();
          return;
        }
        if (response && response.ok) {
          resolve(true);
          return;
        }
        fallbackNav();
      });
    } catch {
      fallbackNav();
    }
  });
}

/** Após usar o popup, o documento do WA não está focado — Clipboard API falha sem isso. */
function focusWhatsappTabFromExtension() {
  return new Promise((resolve) => {
    if (!isExtensionContextAlive()) {
      resolve(false);
      return;
    }
    try {
      chrome.runtime.sendMessage({ type: "WPP_FOCUS_TAB" }, (response) => {
        if (chrome.runtime.lastError) {
          resolve(false);
          return;
        }
        resolve(!!(response && response.ok));
      });
    } catch {
      resolve(false);
    }
  });
}

async function ensureWhatsappTabFocused() {
  await focusWhatsappTabFromExtension();
  try {
    window.focus();
    if (window.top && window.top !== window) window.top.focus();
  } catch {
    /* */
  }
  const until = Date.now() + 4500;
  while (Date.now() < until && typeof document.hasFocus === "function" && !document.hasFocus()) {
    await delay(120);
  }
  await delay(260 + Math.random() * 240);
}

function gaussianRandom(mean, stdev) {
  const u = 1 - Math.random();
  const v = Math.random();
  const z = Math.sqrt(-2 * Math.log(u)) * Math.cos(2 * Math.PI * v);
  return z * stdev + mean;
}

/** Simula evento de clique completo para evitar bloqueios de UI */
function dispatchSyntheticClick(el) {
  if (!el) return;
  ["mousedown", "mouseup", "click"].forEach((name) => {
    el.dispatchEvent(new MouseEvent(name, { bubbles: true, cancelable: true, view: window }));
  });
}

/** Simula foco e scroll antes da interação */
async function simulateMouseToElement(el) {
  if (!el) return;
  el.scrollIntoView({ behavior: "smooth", block: "center" });
  await delay(200);
}

async function delayShortHuman() {
  const ms = Math.max(200, gaussianRandom(900, 350));
  await delay(ms);
}

async function delayBetweenMessages() {
  const s = await safeLocalGet([K.minSec, K.maxSec, K.antiIntervalOn]);
  let minS = Number(s[K.minSec]);
  let maxS = Number(s[K.maxSec]);
  if (Number.isNaN(minS)) minS = 5;
  if (Number.isNaN(maxS)) maxS = 15;
  minS = Math.max(0.3, minS);
  maxS = Math.max(minS, maxS);

  if (s[K.antiIntervalOn] === false) {
    const ms = Math.max(400, Math.round(minS * 1000));
    await delay(ms);
    return;
  }

  const seconds = minS + Math.random() * (maxS - minS);
  await delay(Math.round(seconds * 1000));
}

let logMutex = Promise.resolve();
function appendLog(level, text) {
  logMutex = logMutex
    .then(async () => {
      const data = await safeLocalGet([K.logs]);
      const logs = Array.isArray(data[K.logs]) ? [...data[K.logs]] : [];
      logs.push({ time: Date.now(), level, text });
      await safeLocalSet({ [K.logs]: logs.slice(-150) });
    })
    .catch(() => {});
  return logMutex;
}

let resultMutex = Promise.resolve();
function recordResult(phone, name, status) {
  resultMutex = resultMutex
    .then(async () => {
      const data = await safeLocalGet([K.results]);
      const results = Array.isArray(data[K.results]) ? [...data[K.results]] : [];
      results.push({
        phone,
        name: name || "",
        status,
        timestamp: new Date().toISOString(),
      });
      await safeLocalSet({ [K.results]: results.slice(-2000) });
    })
    .catch(() => {});
  return resultMutex;
}

function emptyDispatchState() {
  return {
    [K.active]: false,
    [K.paused]: false,
    [K.queue]: [],
    [K.index]: 0,
    [K.template]: "",
  };
}

async function getDispatchState() {
  const keys = [
    K.queue,
    K.template,
    K.index,
    K.active,
    K.paused,
    K.stats,
    K.minSec,
    K.maxSec,
  ];
  try {
    const data = await safeLocalGet(keys);
    if (!isExtensionContextAlive()) return emptyDispatchState();
    return data;
  } catch (e) {
    if (isExtensionContextInvalidatedError(e)) return emptyDispatchState();
    throw e;
  }
}

async function waitWhilePausedOrStopped() {
  for (;;) {
    try {
      if (!isExtensionContextAlive()) return false;
      const s = await getDispatchState();
      if (!s[K.active]) return false;
      if (!s[K.paused]) return true;
      await delay(400);
    } catch (e) {
      if (isExtensionContextInvalidatedError(e)) return false;
      throw e;
    }
  }
}

function parseSpinTax(text) {
  return text.replace(/\{([^{}]+)\}/g, (match, choices) => {
    // Evita processar {nome} como spintax
    if (match.toLowerCase() === "{nome}") return match;
    const parts = choices.split("|");
    return parts[Math.floor(Math.random() * parts.length)];
  });
}

function applyTemplate(template, row) {
  const nome = (row && row.name) || "";
  let text = template.replace(/\{nome\}/gi, nome);
  return parseSpinTax(text);
}

async function buildOutgoingMessage(template, row) {
  const s = await safeLocalGet([K.antiRandom]);
  let text = applyTemplate(template, row);
  if (s[K.antiRandom]) {
    text += "\u200B".repeat(1 + Math.floor(Math.random() * 2));
    const pad = ["", " ", "\n"][Math.floor(Math.random() * 3)];
    text += pad;
  }
  return text;
}

async function incrementBlockAndMaybePause() {
  const s = await safeLocalGet([
    K.antiBlocks,
    K.antiBlockSize,
    K.antiBlockWait,
    K.batchCount,
    K.active,
    K.paused,
  ]);
  if (!s[K.antiBlocks]) return;

  let c = Number(s[K.batchCount]) || 0;
  c += 1;
  const size = Math.max(5, Number(s[K.antiBlockSize]) || 50);
  const waitMin = Math.max(1, Number(s[K.antiBlockWait]) || 2);

  if (c >= size) {
    await appendLog(
      "warn",
      `Antibloqueio: pausa de ${waitMin} min após ${size} contatos.`
    );
    await safeLocalSet({ [K.batchCount]: 0 });
    const totalMs = waitMin * 60 * 1000;
    let waited = 0;
    while (waited < totalMs) {
      try {
        if (!isExtensionContextAlive()) return;
        const st = await safeLocalGet([K.active, K.paused]);
        if (!st[K.active]) return;
        if (st[K.paused]) {
          await delay(500);
          continue;
        }
        const step = Math.min(8000, totalMs - waited);
        await delay(step);
        waited += step;
      } catch (e) {
        if (isExtensionContextInvalidatedError(e)) return;
        throw e;
      }
    }
  } else {
    await safeLocalSet({ [K.batchCount]: c });
  }
}

function digitsOnly(v) {
  return String(v || "").replace(/\D/g, "");
}

/** Dígitos do telefone na URL atual (?phone= ou /send/...) */
function getPhoneDigitsFromLocation() {
  try {
    const u = new URL(location.href);
    const q = digitsOnly(u.searchParams.get("phone"));
    if (q) return q;
    const m = u.pathname.match(/\/send\/+(\d[\d]*)/i);
    if (m) return digitsOnly(m[1]);
    return "";
  } catch {
    return "";
  }
}

/**
 * Precisa abrir /send?phone=… — o WA costuma redirecionar para "/" mantendo o chat;
 * nesse caso a URL não tem mais /send e não podemos renavegar em loop.
 */
async function mustNavigateToOpenChat(rowPhone) {
  const want = digitsOnly(rowPhone);
  const fromUrl = getPhoneDigitsFromLocation();
  if (fromUrl && fromUrl === want) return false;

  const { [K.navPending]: pend } = await safeLocalGet(K.navPending);
  if (
    pend &&
    digitsOnly(pend.phone) === want &&
    Date.now() - (pend.at || 0) < NAV_PENDING_TTL_MS
  ) {
    return false;
  }

  try {
    const u = new URL(location.href);
    if (u.pathname.includes("/send")) {
      return fromUrl !== want;
    }
  } catch {
    /* ignora */
  }

  return true;
}

async function clearNavPending() {
  await safeLocalRemove(K.navPending);
}

function findInvalidPhoneDialog() {
  const dialogs = document.querySelectorAll('[role="dialog"]');
  for (const d of dialogs) {
    const t = d.innerText || "";
    if (/inv[aá]lid/i.test(t) && /(telefone|phone)/i.test(t)) return d;
  }
  return null;
}

function closeDialogBestEffort(dialog) {
  const candidates = dialog.querySelectorAll(
    '[role="button"], button, [data-testid="dialog-dismiss"]'
  );
  for (const b of candidates) {
    const tx = (b.textContent || "").trim().toUpperCase();
    if (
      tx === "OK" ||
      tx === "ENTENDI" ||
      tx === "FECHAR" ||
      tx === "GOT IT" ||
      tx === "CLOSE"
    ) {
      b.click();
      return true;
    }
  }
  const first = dialog.querySelector("button");
  if (first) {
    first.click();
    return true;
  }
  return false;
}

function isEditableVisible(el) {
  if (!el || !el.isContentEditable) return false;
  const r = el.getBoundingClientRect();
  return r.width > 2 && r.height > 2 && el.offsetParent !== null;
}

function findComposer() {
  const trySelectors = [
    'div[contenteditable="true"][data-tab="10"]',
    'footer div[contenteditable="true"]',
    'div[contenteditable="true"][role="textbox"]',
    '[aria-label][contenteditable="true"]',
  ];
  for (const sel of trySelectors) {
    const el = document.querySelector(sel);
    if (el && isEditableVisible(el)) return el;
  }
  const all = document.querySelectorAll('div[contenteditable="true"]');
  for (const el of all) {
    const lab = (el.getAttribute("aria-label") || "").toLowerCase();
    if (
      (lab.includes("mensagem") ||
        lab.includes("message") ||
        lab.includes("type")) &&
      isEditableVisible(el) &&
      !lab.includes("pesquis") &&
      !lab.includes("search")
    ) {
      return el;
    }
  }
  return null;
}

function gatherSendButtonElements() {
  const seen = new Set();
  const out = [];
  const add = (btn) => {
    if (!btn || seen.has(btn)) return;
    seen.add(btn);
    const r = btn.getBoundingClientRect();
    if (r.width < 6 || r.height < 6) return;
    if (r.bottom < -20 || r.top > window.innerHeight + 80) return;
    out.push(btn);
  };
  for (const icon of ["send", "wds-ic-send", "wds-ic-send-filled"]) {
    document.querySelectorAll(`span[data-icon="${icon}"]`).forEach((sp) => {
      add(sp.closest("button"));
      add(sp.closest('div[role="button"]'));
    });
  }
  document
    .querySelectorAll('button[aria-label*="Enviar"], button[aria-label*="Send"]')
    .forEach((b) => add(b));
  return out;
}

function isSendButtonInteractable(btn) {
  if (!btn?.isConnected) return false;
  if (btn.hasAttribute("disabled")) return false;
  if (btn.getAttribute("aria-disabled") === "true") return false;
  const cs = window.getComputedStyle(btn);
  if (cs.visibility === "hidden" || cs.display === "none") return false;
  if (cs.pointerEvents === "none") return false;
  return true;
}

/** O botão Enviar mais baixo na tela (compositor do chat). */
function rankSendButtonsForClick(list) {
  return [...list].sort((a, b) => {
    const ra = a.getBoundingClientRect();
    const rb = b.getBoundingClientRect();
    return rb.bottom + rb.height * 0.5 - (ra.bottom + ra.height * 0.5);
  });
}

/** Um único ciclo mouse — evita 3× envio (cadeia + span interno + click nativo). */
function dispatchSyntheticClick(el) {
  if (!el) return;
  const r = el.getBoundingClientRect();
  const x = r.left + Math.max(1, Math.min(r.width - 1, r.width * 0.5));
  const y = r.top + Math.max(1, Math.min(r.height - 1, r.height * 0.5));
  const base = {
    bubbles: true,
    cancelable: true,
    view: window,
    clientX: x,
    clientY: y,
    button: 0,
  };
  el.dispatchEvent(new MouseEvent("mousedown", { ...base, buttons: 1 }));
  el.dispatchEvent(new MouseEvent("mouseup", { ...base, buttons: 0 }));
  el.dispatchEvent(new MouseEvent("click", { ...base, buttons: 0 }));
}

/** Espera o botão Enviar e clica uma vez (mensagem de texto). */
async function clickSendWhenReady() {
  const waitBtnMs = 12000;
  const t0 = Date.now();
  let primary = null;

  while (Date.now() - t0 < waitBtnMs) {
    let candidates = gatherSendButtonElements().filter(isSendButtonInteractable);
    if (!candidates.length) candidates = gatherSendButtonElements();
    if (candidates.length) {
      const ranked = rankSendButtonsForClick(candidates);
      primary = ranked[0];
      if (primary) break;
    }
    await delay(200);
  }

  if (!primary) {
    const last = gatherSendButtonElements()[0];
    if (!last) return false;
    primary = last;
  }

  await simulateMouseToElement(primary);
  await delay(70 + Math.random() * 90);
  dispatchSyntheticClick(primary);
  await delay(520 + Math.random() * 280);
  return true;
}

async function simulateMouseToElement(el) {
  if (!el) return;
  const rect = el.getBoundingClientRect();
  const targetX = rect.left + rect.width * (0.25 + Math.random() * 0.5);
  const targetY = rect.top + rect.height * (0.25 + Math.random() * 0.5);
  let x = Math.random() * window.innerWidth;
  let y = Math.random() * window.innerHeight;
  const steps = 16 + Math.floor(Math.random() * 14);
  for (let i = 0; i <= steps; i++) {
    const t = i / steps;
    const nx = x + (targetX - x) * t + (Math.random() - 0.5) * 42;
    const ny = y + (targetY - y) * t + (Math.random() - 0.5) * 42;
    document.dispatchEvent(
      new MouseEvent("mousemove", {
        bubbles: true,
        cancelable: true,
        clientX: nx,
        clientY: ny,
        view: window,
      })
    );
    await delay(16 + Math.random() * 48);
  }
}

/** Converte um DataURL (base64) de volta para File */
function dataUrlToFile(dataUrl, fileName, mime) {
  try {
    const [, b64] = dataUrl.split(",");
    const byteString = atob(b64);
    const arr = new Uint8Array(byteString.length);
    for (let i = 0; i < byteString.length; i++) arr[i] = byteString.charCodeAt(i);
    return new File([arr], fileName || "imagem.jpg", { type: mime || "image/jpeg" });
  } catch (e) {
    return null;
  }
}

/**
 * Injeta um File num input[type=file] usando o native setter do prototype.
 */
async function injectFileNative(fileInput, file) {
  const dt = new DataTransfer();
  dt.items.add(file);
  const nativeSetter = Object.getOwnPropertyDescriptor(
    HTMLInputElement.prototype, "files"
  )?.set;
  if (nativeSetter) {
    nativeSetter.call(fileInput, dt.files);
  } else {
    try {
      Object.defineProperty(fileInput, "files", {
        value: dt.files, configurable: true,
      });
    } catch { /* erro silencioso */ }
  }
  fileInput.dispatchEvent(new Event("input",  { bubbles: true }));
  await delay(50);
  fileInput.dispatchEvent(new Event("change", { bubbles: true }));
}

/**
 * Verifica se o preview de mídia do WhatsApp está REALMENTE aberto.
 */
function isMediaPreviewVisible() {
  const allEditables = [...document.querySelectorAll('div[contenteditable="true"]')];
  const captionEl = allEditables.find((el) => {
    if (!el.offsetParent || el.closest("footer")) return false;
    const lab = (el.getAttribute("aria-label") || "").toLowerCase();
    return /legenda|caption|adicione|add a/i.test(lab);
  });
  if (captionEl) return true;

  const blobs = [...document.querySelectorAll('img[src^="blob:"]')].filter((img) => {
    if (!img.offsetParent) return false;
    const rect = img.getBoundingClientRect();
    return rect.width > 150 && rect.height > 150 && !img.closest("#main");
  });
  return blobs.length > 0;
}

/**
 * Aguarda o preview de mídia aparecer.
 */
async function waitForMediaPreview(timeoutMs = 15000) {
  const t0 = Date.now();
  while (Date.now() - t0 < timeoutMs) {
    if (!isExtensionContextAlive()) return false;
    if (isMediaPreviewVisible()) return true;
    await delay(500);
  }
  return false;
}

/**
 * Retorna o campo de legenda do preview.
 */
function findCaptionField() {
  // 1. Container oficial pelo data-testid
  const container = document.querySelector('[data-testid="media-caption-input-container"]');
  if (container) {
    const edit = container.querySelector('div[contenteditable="true"]');
    if (edit) return edit;
  }

  // 2. Tenta localizar pelo span lexical que o usuário forneceu
  const lexicalSpan = document.querySelector('span[data-lexical-text="true"]');
  if (lexicalSpan) {
    const parentEditable = lexicalSpan.closest('div[contenteditable="true"]');
    if (parentEditable && parentEditable.offsetParent) return parentEditable;
  }

  // 3. Fallback: Busca genérica por atributos e restrições de localização
  return [...document.querySelectorAll('div[contenteditable="true"]')].find((el) => {
    if (!el.offsetParent) return false;
    // O campo de legenda do preview nunca está no footer ou no chat principal (#main)
    if (el.closest("footer") || el.closest("#main")) return false;

    const lab = (el.getAttribute("aria-label") || "").toLowerCase();
    const ph = (el.getAttribute("placeholder") || "").toLowerCase();
    return /legenda|caption|adicione|add|type/i.test(lab) || /legenda|caption/i.test(ph);
  }) || null;
}

/**
 * Localiza o botão Enviar do preview de mídia.
 */
function findPreviewSendButton() {
  const previewSend = document.querySelector('[data-testid="media-caption-send-button"]');
  if (previewSend && previewSend.offsetParent) return previewSend;

  // Seletores alternativos de ícone ou tid
  const candidates = [...document.querySelectorAll('[data-testid="send"], [data-testid="compose-btn-send"], [data-icon="send"]')]
    .map(el => el.closest('button') || el.closest('[role="button"]') || el)
    .filter(btn => {
      if (!btn || !btn.offsetParent) return false;
      const rect = btn.getBoundingClientRect();
      // Em preview, o botão fica centralizado ou à direita, mas nunca no rodapé principal (#main)
      return !btn.closest("footer") && !btn.closest("#main") && rect.width > 20;
    });

  return candidates[0] || null;
}

/**
 * Localiza o botão (+) de anexo.
 */
function findPlusButton() {
  const footer = document.querySelector("footer");
  if (!footer) return null;
  for (const sel of [
    '[data-testid="clip"]',
    '[data-testid="attach-menu-plus"]',
    'span[data-icon="clip"]',
    'span[data-icon="plus-rounded"]',
    'span[data-icon="plus"]',
  ]) {
    const sp = footer.querySelector(sel);
    const btn = sp?.closest("button") || sp?.closest('div[role="button"]');
    if (btn && btn.offsetParent !== null) return btn;
  }
  return null;
}

async function sendImageWithCaption(attData, captionText) {
  await appendLog("info", "Injetando imagem virtualmente (simulação humana)...");
  const file = dataUrlToFile(attData.dataUrl, attData.name, attData.mime);
  if (!file) {
    await appendLog("err", "Erro ao processar arquivo da imagem.");
    return false;
  }

  try {
    // Busca o campo de mensagem atual
    const composer = document.querySelector('div[contenteditable="true"][data-tab="10"]') 
                  || document.activeElement 
                  || document.body;

    // Foca o campo para receber o comando colar
    if (composer && typeof composer.focus === "function") {
      composer.focus();
    }
    
    await delay(300);

    // Cria os dados de transferência forjando a imagem
    const dataTransfer = new DataTransfer();
    dataTransfer.items.add(file);

    // Despacha um evento "Paste" (Colar) perfeito que o WhatsApp reconhece
    const pasteEvent = new ClipboardEvent("paste", {
      bubbles: true,
      cancelable: true,
      clipboardData: dataTransfer
    });

    composer.dispatchEvent(pasteEvent);
    
    await appendLog("info", "Imagem colada. Aguardando a tela do preview...");

    // Aguarda o WhatsApp reagir e abrir o Modal
    const previewVisible = await waitForMediaPreview(10000);
    
    if (!previewVisible) {
      // Fallback: Tentativa via Drag and Drop nativo se o Paste for bloqueado
      await appendLog("info", "Preview atrasado, forçando Drag & Drop...");
      const dropZone = document.querySelector('#main') || document.body;
      const dtDrop = new DataTransfer();
      dtDrop.items.add(file);
      
      dropZone.dispatchEvent(new DragEvent('dragenter', { bubbles: true, cancelable: true, dataTransfer: dtDrop }));
      dropZone.dispatchEvent(new DragEvent('dragover', { bubbles: true, cancelable: true, dataTransfer: dtDrop }));
      dropZone.dispatchEvent(new DragEvent('drop', { bubbles: true, cancelable: true, dataTransfer: dtDrop }));
      
      const previewVisibleDrop = await waitForMediaPreview(10000);
      if (!previewVisibleDrop) {
         throw new Error("O WhatsApp bloqueou a injeção via Paste e Drop.");
      }
    }

    await delay(1200); // Tempo para o WhatsApp abrir o modal e focar o campo de legenda sozinho

    // Pega o que o WhatsApp focar automaticamente (geralmente a caixa de legenda)
    let focusTarget = document.activeElement;

    // Se ele não focou sozinho (raro), tentamos caçar manual
    if (!focusTarget || focusTarget.tagName === "BODY") {
      focusTarget = findCaptionField();
      if (focusTarget) focusTarget.focus();
    }

    // Se tem legenda, usa Ctrl+V simulado no alvo focado
    if (captionText && captionText.trim()) {
      if (focusTarget) {
        await pasteMessageInto(focusTarget, captionText);
        await appendLog("info", "Legenda colada no preview.");
      } else {
        await appendLog("warn", "Campo não ativado para legenda. Enviando sem legenda...");
      }
    }

    await delay(800);
    
    // Tenta enviar com Enter no campo focado
    if (focusTarget) {
      await appendLog("info", "Apertando Enter para disparar anexos...");
      const enterDown = new KeyboardEvent("keydown", { bubbles: true, cancelable: true, key: "Enter", code: "Enter", keyCode: 13, which: 13 });
      focusTarget.dispatchEvent(enterDown);
      
      await delay(100);
      
      const enterUp = new KeyboardEvent("keyup", { bubbles: true, cancelable: true, key: "Enter", code: "Enter", keyCode: 13, which: 13 });
      focusTarget.dispatchEvent(enterUp);
    } else {
      // Fallback extremo
      const sendBtn = findPreviewSendButton();
      if (!sendBtn) throw new Error("Botão de enviar ausente e campo também.");
      await simulateMouseToElement(sendBtn);
      dispatchSyntheticClick(sendBtn);
    }
    
    const tEnd = Date.now();
    while (isMediaPreviewVisible() && Date.now() - tEnd < 5000) await delay(500);

    await appendLog("info", "Imagem disparada 🚀.");
    return true;

  } catch (err) {
    await appendLog("err", `Falha no envio da imagem: ${err.message}`);
    return false;
  }
}



/**
 * Cola a mensagem inteira de uma vez (rápido).
 * Ordem: clipboard + colar → insertText completo → fallback DOM.
 */
/**
 * Simula status "Digitando..." focando o elemento e esperando um tempo
 * proporcional ao comprimento do texto.
 */
async function simulateTypingStatus(el, text) {
  if (!el || !text) return;
  el.focus();
  // Simula um delay de digitação: 40ms a 70ms por caractere, limitado a 8 segundos
  const typingMs = Math.min(8000, text.length * gaussianRandom(50, 15));
  
  // Pequena interação inicial para disparar o evento de início de digitação no WA
  el.dispatchEvent(new KeyboardEvent("keydown", { bubbles: true, cancelable: true, key: "Shift" }));
  
  await delay(typingMs);
}

async function pasteMessageInto(el, text) {
  el.focus();
  await delay(80 + Math.random() * 120);
  try {
    document.execCommand("selectAll", false, null);
    document.execCommand("delete", false, null);
  } catch {
    el.textContent = "";
    el.dispatchEvent(
      new InputEvent("input", { bubbles: true, data: null, inputType: "deleteContent" })
    );
  }
  await delay(40 + Math.random() * 80);

  let inserted = false;
  try {
    await navigator.clipboard.writeText(text);
    inserted = document.execCommand("paste");
  } catch {
    /* clipboard indisponível — segue para insertText */
  }
  if (!inserted) {
    inserted = document.execCommand("insertText", false, text);
  }
  if (!inserted) {
    el.textContent = text;
    el.dispatchEvent(
      new InputEvent("input", {
        bubbles: true,
        data: text,
        inputType: "insertFromPaste",
      })
    );
    inserted = true;
  }
  el.dispatchEvent(new Event("change", { bubbles: true }));
}

async function waitForComposerOrInvalid(timeoutMs) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    try {
      if (!isExtensionContextAlive()) return "stopped";
      const go = await waitWhilePausedOrStopped();
      if (!go) return "stopped";

      const inv = findInvalidPhoneDialog();
      if (inv) return { type: "invalid", el: inv };

      const comp = findComposer();
      if (comp) return { type: "composer", el: comp };

      await delay(500);
    } catch (e) {
      if (isExtensionContextInvalidatedError(e)) return "stopped";
      throw e;
    }
  }
  return "timeout";
}

async function bumpSent() {
  const data = await safeLocalGet([K.stats]);
  const cur = data[K.stats] || { sent: 0, failed: 0, total: 0 };
  await safeLocalSet({
    [K.stats]: { ...cur, sent: (cur.sent || 0) + 1 },
  });
}

async function bumpFailed() {
  const data = await safeLocalGet([K.stats]);
  const cur = data[K.stats] || { sent: 0, failed: 0, total: 0 };
  await safeLocalSet({
    [K.stats]: { ...cur, failed: (cur.failed || 0) + 1 },
  });
}

async function finishCampaign(msg) {
  await safeLocalSet({
    [K.active]: false,
    [K.paused]: false,
    [K.batchCount]: 0,
  });
  await safeLocalRemove("wppPendingAttachment");
  await clearNavPending();
  await appendLog("ok", msg);
}

let bootstrapLock = false;

/**
 * Contexto invalidado (extensão recarregada): não chamar runtime — evita “Uncaught (in promise)”.
 * SW indisponível: falha aberta (true) como antes.
 */
async function tryDispatchLease() {
  if (!isExtensionContextAlive()) return false;
  try {
    const r = await chrome.runtime.sendMessage({ type: "WPP_DISPATCH_LEASE_TRY" });
    return !!(r && r.ok);
  } catch (e) {
    if (isExtensionContextInvalidatedError(e)) return false;
    return true;
  }
}

async function releaseDispatchLease() {
  if (!isExtensionContextAlive()) return;
  try {
    await chrome.runtime.sendMessage({ type: "WPP_DISPATCH_LEASE_RELEASE" });
  } catch {
    /* SW inativo ou contexto invalidado */
  }
}

async function goToNextOrEnd(nextIndex, queue) {
  const live = await getDispatchState();
  if (!live[K.active]) return;

  if (nextIndex >= 1) {
    await incrementBlockAndMaybePause();
  }

  await safeLocalSet({ [K.index]: nextIndex });
  if (nextIndex >= queue.length) {
    await finishCampaign("Todos os contatos foram processados.");
    return;
  }
  const next = queue[nextIndex];
  const url = `https://web.whatsapp.com/send?phone=${encodeURIComponent(next.phone)}`;
  await safeLocalSet({
    [K.navPending]: { phone: digitsOnly(next.phone), at: Date.now() },
  });
  await navigateToSendUrl(url);
}

async function handleInvalidNumber(row) {
  const s = await getDispatchState();
  const queue = s[K.queue] || [];
  const idx = s[K.index] || 0;

  await appendLog("err", `Erro: número inválido — ${row.name || row.phone}.`);
  await recordResult(row.phone, row.name, "Falha");
  await bumpFailed();

  const next = idx + 1;
  await delayBetweenMessages();
  if (!(await getDispatchState())[K.active]) return;
  await goToNextOrEnd(next, queue);
}

async function handleSuccess(row) {
  const s = await getDispatchState();
  const queue = s[K.queue] || [];
  const idx = s[K.index] || 0;

  await appendLog("ok", `Mensagem entregue para ${row.name || row.phone}.`);
  await recordResult(row.phone, row.name, "Sucesso");
  await bumpSent();

  const next = idx + 1;
  await delayBetweenMessages();
  if (!(await getDispatchState())[K.active]) return;
  await goToNextOrEnd(next, queue);
}

async function runOneContact() {
  try {
    await runOneContactCore();
  } catch (e) {
    if (isExtensionContextInvalidatedError(e)) return;
    throw e;
  }
}

async function runOneContactCore() {
  const s = await getDispatchState();
  if (!s[K.active]) return;

  const queue = s[K.queue] || [];
  const idx = s[K.index] || 0;
  const template = s[K.template] || "";

  if (!queue.length) {
    await finishCampaign("Fila vazia.");
    return;
  }
  if (idx >= queue.length) {
    await finishCampaign("Todos os contatos foram processados.");
    return;
  }

  const row = queue[idx];
  const { [K.navPending]: stalePend } = await safeLocalGet(K.navPending);
  if (
    stalePend &&
    digitsOnly(stalePend.phone) !== digitsOnly(row.phone)
  ) {
    await clearNavPending();
  }

  const needNav = await mustNavigateToOpenChat(row.phone);
  if (needNav) {
    await appendLog(
      "info",
      `Iniciando envio para ${row.name ? row.name : row.phone}...`
    );
    if (!(await waitWhilePausedOrStopped())) return;
    await safeLocalSet({
      [K.navPending]: { phone: digitsOnly(row.phone), at: Date.now() },
    });
    await navigateToSendUrl(
      `https://web.whatsapp.com/send?phone=${encodeURIComponent(row.phone)}`
    );
    return;
  }

  await appendLog(
    "info",
    `Aguardando o campo de mensagem (${row.name ? row.name : row.phone})...`
  );

  const outcome = await waitForComposerOrInvalid(120000);
  if (outcome === "stopped") return;
  await clearNavPending();
  if (outcome === "timeout") {
    await appendLog("err", `Tempo esgotado aguardando chat — ${row.phone}`);
    await recordResult(row.phone, row.name, "Falha");
    await bumpFailed();
    const next = idx + 1;
    await delayBetweenMessages();
    if (!(await getDispatchState())[K.active]) return;
    await goToNextOrEnd(next, queue);
    return;
  }

  if (outcome.type === "invalid") {
    closeDialogBestEffort(outcome.el);
    await delay(600 + Math.random() * 400);
    await handleInvalidNumber(row);
    return;
  }

  const composer = outcome.el;
  await ensureWhatsappTabFocused();

  const personalized = await buildOutgoingMessage(template, row);

  // 1. Enviar primeiro o texto (se houver template)
  if (personalized && personalized.trim()) {
    try {
      await simulateTypingStatus(composer, personalized);
      await pasteMessageInto(composer, personalized);
      const sentOk = await clickSendWhenReady();
      if (!sentOk) {
        throw new Error("Botão de enviar texto não respondeu.");
      }
      await delay(1000 + Math.random() * 500); // Pausa entre texto e imagem
    } catch (e) {
      await appendLog("err", `Falha ao enviar texto: ${e.message}`);
      await recordResult(row.phone, row.name, "Falha");
      await bumpFailed();
      await delayBetweenMessages();
      if (!(await getDispatchState())[K.active]) return;
      await goToNextOrEnd(idx + 1, queue);
      return;
    }
  }

  // 2. Enviar a imagem (se houver anexo)
  const attData = await safeLocalGet([K.pendingAttachment, K.enableCaption, K.imageCaption]);
  const att = attData[K.pendingAttachment];
  const useCaption = !!attData[K.enableCaption];
  const caption = useCaption ? (attData[K.imageCaption] || "") : "";

  if (att && att.dataUrl) {
    const imgSent = await sendImageWithCaption(att, caption); 
    if (!imgSent) {
      await appendLog("err", `Falha ao enviar imagem para ${row.phone}.`);
      await recordResult(row.phone, row.name, "Falha");
      await bumpFailed();
      await delayBetweenMessages();
      if (!(await getDispatchState())[K.active]) return;
      await goToNextOrEnd(idx + 1, queue);
      return;
    }
  }

  await delay(500 + Math.random() * 400);

  const lateInvalid = findInvalidPhoneDialog();
  if (lateInvalid) {
    closeDialogBestEffort(lateInvalid);
    await delay(400);
    await handleInvalidNumber(row);
    return;
  }

  await handleSuccess(row);
}

async function safeBootstrap() {
  if (bootstrapLock) return;
  if (typeof document !== "undefined" && document.visibilityState === "hidden") return;
  const leased = await tryDispatchLease();
  if (!leased) return;
  bootstrapLock = true;
  try {
    const s = await getDispatchState();
    if (!s[K.active]) return;
    if (s[K.paused]) return;
    await runOneContact();
  } catch (e) {
    if (isExtensionContextInvalidatedError(e)) return;
    try {
      await appendLog("err", `Erro interno: ${e && e.message ? e.message : String(e)}`);
    } catch {
      /* contexto morto durante o log */
    }
  } finally {
    bootstrapLock = false;
    try {
      await releaseDispatchLease();
    } catch {
      /* idem */
    }
  }
}

/** Um único disparo: `active`+`paused` no mesmo set + WPP_KICK geravam 3× runOneContact. */
let bootstrapKickTimer = null;
function scheduleBootstrapKick() {
  if (bootstrapKickTimer) clearTimeout(bootstrapKickTimer);
  bootstrapKickTimer = setTimeout(() => {
    bootstrapKickTimer = null;
    if (!isExtensionContextAlive()) return;
    void safeLocalGet([K.active, K.paused])
      .then((d) => {
        try {
          if (!isExtensionContextAlive()) return;
          if (d[K.active] && !d[K.paused]) void safeBootstrap().catch(() => {});
        } catch {
          /* Extension context invalidated no tick do storage */
        }
      })
      .catch(() => {});
  }, 380);
}

chrome.runtime.onMessage.addListener((msg) => {
  try {
    if (!isExtensionContextAlive()) return;
    if (msg && msg.type === "WPP_KICK") {
      scheduleBootstrapKick();
    }
  } catch {
    /* Extension context invalidated ou outro erro no handler */
  }
});

chrome.storage.onChanged.addListener((changes, area) => {
  try {
    if (area !== "local" || !isExtensionContextAlive()) return;
    const unpaused = changes[K.paused] && changes[K.paused].newValue === false;
    const activated = changes[K.active] && changes[K.active].newValue === true;
    if (unpaused || activated) {
      scheduleBootstrapKick();
    }
  } catch {
    /* idem */
  }
});

(async () => {
  try {
    if (!isExtensionContextAlive()) return;
    const s = await getDispatchState();
    if (s[K.active] && !s[K.paused]) {
      await delay(800 + Math.random() * 700);
      scheduleBootstrapKick();
    }
  } catch {
    /* arranque com contexto morto */
  }
})();
