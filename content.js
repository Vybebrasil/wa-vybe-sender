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

function applyTemplate(template, row) {
  const nome = (row && row.name) || "";
  return template.replace(/\{nome\}/gi, nome);
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
  const [, b64] = dataUrl.split(",");
  const byteString = atob(b64);
  const arr = new Uint8Array(byteString.length);
  for (let i = 0; i < byteString.length; i++) arr[i] = byteString.charCodeAt(i);
  return new File([arr], fileName || "imagem.jpg", { type: mime || "image/jpeg" });
}

/**
 * Injeta um File num input[type=file] usando o native setter do prototype.
 * Necessário para acionar os handlers internos do React do WhatsApp Web.
 */
function injectFileNative(fileInput, file) {
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
    } catch { /* ignora */ }
  }
  fileInput.dispatchEvent(new Event("input",  { bubbles: true }));
  fileInput.dispatchEvent(new Event("change", { bubbles: true }));
}

/**
 * Verifica se o preview de mídia do WhatsApp está visível na tela.
 * Critério: imagem blob grande e visível, ou campo de legenda específico.
 */
function isMediaPreviewVisible() {
  // Campo de legenda com aria-label específico do preview
  const captionEl = [...document.querySelectorAll('div[contenteditable="true"]')].find((el) => {
    const lab = (el.getAttribute("aria-label") || "").toLowerCase();
    return /caption|legenda|adicione uma legenda|add a caption/i.test(lab) && el.offsetParent !== null;
  });
  if (captionEl) return true;

  // Imagem blob grande no centro da tela (preview de mídia)
  const blobs = [...document.querySelectorAll('img[src^="blob:"]')].filter((img) => {
    if (!img.offsetParent) return false;
    const r = img.getBoundingClientRect();
    return r.width > 80 && r.height > 80;
  });
  return blobs.length > 0;
}

/**
 * Aguarda o preview de mídia do WhatsApp aparecer.
 * Retorna true se abriu, false se timeout.
 */
async function waitForMediaPreview(timeoutMs = 18000) {
  const t0 = Date.now();
  while (Date.now() - t0 < timeoutMs) {
    if (!isExtensionContextAlive()) return false;
    if (isMediaPreviewVisible()) return true;
    await delay(300);
  }
  return false;
}

/**
 * Retorna o campo de legenda do preview de mídia.
 * Deve ser chamado DEPOIS de waitForMediaPreview retornar true.
 */
function findCaptionField() {
  // Buscar pelo aria-label do campo de legenda do WA (PT e EN)
  const byLabel = [...document.querySelectorAll('div[contenteditable="true"]')].find((el) => {
    const lab = (el.getAttribute("aria-label") || "").toLowerCase();
    return /caption|legenda|adicione uma legenda|add a caption/i.test(lab) && el.offsetParent !== null;
  });
  if (byLabel) return byLabel;

  // Buscar dentro da área do preview (próximo a imgs blob)
  const blobs = [...document.querySelectorAll('img[src^="blob:"]')]
    .filter((img) => img.offsetParent !== null && img.getBoundingClientRect().width > 80);
  for (const img of blobs) {
    // Subir até encontrar um container e buscar um contenteditable irmão/filho
    let el = img.parentElement;
    for (let d = 0; d < 15 && el; d++) {
      const found = el.querySelector('div[contenteditable="true"]');
      if (found && found.offsetParent !== null) return found;
      el = el.parentElement;
    }
  }
  return null;
}

/**
 * Localiza o botão Enviar do preview de mídia (FAB verde no canto inferior direito).
 */
function findPreviewSendButton() {
  for (const sel of [
    '[data-testid="media-caption-send-button"]',
    '[data-testid="send"]',
  ]) {
    const el = document.querySelector(sel);
    if (el && el.offsetParent !== null) {
      const btn = el.closest("button") || el.closest('div[role="button"]') || el;
      const r = btn.getBoundingClientRect();
      if (r.width > 4 && r.height > 4) return btn;
    }
  }
  // Buscar pelo ícone de enviar dentro de qualquer botão visível
  for (const iconName of ["send", "wds-ic-send-filled", "wds-ic-send"]) {
    const sp = document.querySelector(`span[data-icon="${iconName}"]`);
    if (!sp) continue;
    const btn = sp.closest("button") || sp.closest('div[role="button"]');
    if (btn && btn.offsetParent !== null) {
      const r = btn.getBoundingClientRect();
      if (r.width > 4 && r.height > 4) return btn;
    }
  }
  // Fallback posicional: botão mais à direita e abaixo (FAB do preview)
  const candidates = [...document.querySelectorAll('button, div[role="button"]')].filter((b) => {
    if (!b.offsetParent) return false;
    const r = b.getBoundingClientRect();
    return r.right > window.innerWidth * 0.55 && r.bottom > window.innerHeight * 0.55
      && r.width > 28 && r.height > 28;
  });
  if (!candidates.length) return null;
  return candidates.reduce((best, b) => {
    const rb = b.getBoundingClientRect(), rBest = best.getBoundingClientRect();
    return (rb.right + rb.bottom) > (rBest.right + rBest.bottom) ? b : best;
  });
}

/**
 * Localiza o botão (+) de anexo no rodapé.
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
  for (const b of footer.querySelectorAll('button, div[role="button"]')) {
    const a = (b.getAttribute("aria-label") || "").toLowerCase();
    if ((a.includes("anexar") || a.includes("attach") || a.includes("clip")) && b.offsetParent !== null) {
      return b;
    }
  }
  return null;
}

/**
 * Envia a imagem com legenda de texto para o contato atual.
 *
 * Técnica: intercepta temporariamente HTMLInputElement.prototype.click.
 * Quando o WA chama input.click() internamente ao clicar no menu "Fotos e vídeos",
 * o interceptor captura esse input e injeta o arquivo com o native setter,
 * eliminando a necessidade de abrir o seletor de arquivo nativo do OS.
 */
async function sendImageWithCaption(attData, captionText) {
  await appendLog("info", "Preparando envio de imagem com legenda...");
  const file = dataUrlToFile(attData.dataUrl, attData.name, attData.mime);

  // ── Passo 1: Interceptar input.click() ──
  let intercepted = false;
  const originalClick = HTMLInputElement.prototype.click;
  HTMLInputElement.prototype.click = function () {
    const accept = (this.getAttribute("accept") || "").toLowerCase();
    const isPhotoInput =
      this.type === "file" &&
      (accept.includes("image") || accept.includes("video") || accept.includes("mp4"));
    if (isPhotoInput && !intercepted) {
      intercepted = true;
      HTMLInputElement.prototype.click = originalClick; // restaurar imediatamente
      injectFileNative(this, file);
      return; // não abre o seletor de arquivo nativo
    }
    originalClick.call(this);
  };

  // ── Passo 2: Abrir menu (+) ──
  const plusBtn = findPlusButton();
  if (!plusBtn) {
    HTMLInputElement.prototype.click = originalClick;
    await appendLog("err", "Botão de anexo (+) não encontrado no rodapé do WhatsApp.");
    return false;
  }

  plusBtn.click();
  await delay(600 + Math.random() * 200);

  // ── Passo 3: Clicar no item de menu "Fotos e vídeos" ──
  // Isso acionará input.click() internamente, que será interceptado acima.
  let menuItemClicked = false;
  const menu = document.querySelector('[role="menu"]') || document.querySelector('[data-testid="attach-menu"]');
  if (menu) {
    const photoLabel = /foto|photo|image|imagem|v[ií]deo|video/i;
    const items = [
      ...menu.querySelectorAll('[role="menuitem"]'),
      ...menu.querySelectorAll("li"),
      ...menu.querySelectorAll("button"),
    ];
    for (const item of items) {
      if (!photoLabel.test(item.textContent || "")) continue;
      item.click();
      menuItemClicked = true;
      break;
    }
  }

  // Se não encontrou item no menu, aguardar um pouco e tentar clicar no menu entry diretamente
  if (!menuItemClicked) {
    await delay(400);
    // Tentar encontrar inputs de foto no DOM e clicar neles diretamente
    const footer = document.querySelector("footer");
    const inputs = [...(footer || document).querySelectorAll('input[type="file"]')];
    const photoInput = inputs.find((i) => {
      const a = (i.getAttribute("accept") || "").toLowerCase();
      return a.includes("image") && !a.endsWith(".webp");
    });
    if (photoInput) {
      photoInput.click(); // será interceptado
    }
  }

  // Aguardar a intercepção acontecer (max 3s)
  const tIntercept = Date.now();
  while (!intercepted && Date.now() - tIntercept < 3000) {
    await delay(100);
  }

  // Restaurar o prototype caso a intercepção não tenha ocorrido
  if (!intercepted) {
    HTMLInputElement.prototype.click = originalClick;
    await appendLog("warn", "Intercepção do input não ocorreu. Tentando injeção direta...");
    // Injeção direta como último recurso
    const footer = document.querySelector("footer");
    const inputs = [...(footer || document).querySelectorAll('input[type="file"]')];
    const photoInput = inputs.find((i) => {
      const a = (i.getAttribute("accept") || "").toLowerCase();
      return a.includes("image") && !a.includes("webp");
    }) || inputs[0];
    if (photoInput) {
      injectFileNative(photoInput, file);
    } else {
      await appendLog("err", "Não foi possível localizar o input de arquivo do WhatsApp.");
      return false;
    }
  }

  await appendLog("info", `Imagem "${file.name}" injetada. Aguardando preview...`);

  // ── Passo 4: Aguardar preview de mídia ──
  const previewOpened = await waitForMediaPreview(18000);
  if (!previewOpened) {
    await appendLog("err", "Preview de mídia não abriu após injeção do arquivo.");
    return false;
  }
  await appendLog("info", "Preview de mídia aberto.");

  // ── Passo 5: Inserir legenda ──
  if (captionText && captionText.trim()) {
    await delay(400 + Math.random() * 200);
    const captionField = findCaptionField();
    if (captionField) {
      try {
        await pasteMessageInto(captionField, captionText);
        await appendLog("info", "Legenda inserida no preview.");
      } catch (e) {
        await appendLog("warn", `Falha ao inserir legenda: ${e?.message || e}`);
      }
      await delay(400 + Math.random() * 200);
    } else {
      await appendLog("warn", "Campo de legenda não encontrado — imagem sem texto.");
    }
  }

  // ── Passo 6: Clicar em Enviar ──
  let sendBtn = findPreviewSendButton();
  if (!sendBtn) {
    const tSend = Date.now();
    while (Date.now() - tSend < 10000) {
      sendBtn = findPreviewSendButton();
      if (sendBtn) break;
      await delay(300);
    }
  }
  if (!sendBtn) {
    await appendLog("err", "Botão Enviar não encontrado no preview da imagem.");
    return false;
  }

  await simulateMouseToElement(sendBtn);
  await delay(80 + Math.random() * 80);
  dispatchSyntheticClick(sendBtn);
  await delay(2000 + Math.random() * 600);

  // Segunda tentativa caso ainda esteja aberto
  if (isMediaPreviewVisible()) {
    const btn2 = findPreviewSendButton();
    if (btn2) {
      dispatchSyntheticClick(btn2);
      await delay(1500 + Math.random() * 400);
    }
  }

  await appendLog("info", "Imagem enviada com sucesso.");
  return true;
}


/**
 * Injeta um arquivo em um input[type=file] usando o native setter do prototype,
 * que é a única forma de acionar os handlers internos do React no WhatsApp.
 */
function injectFileIntoInput(fileInput, file) {
  const dt = new DataTransfer();
  dt.items.add(file);
  // Usar o setter nativo do prototype — necessário para o React do WA reconhecer a mudança
  const nativeSetter = Object.getOwnPropertyDescriptor(
    window.HTMLInputElement.prototype, "files"
  )?.set;
  if (nativeSetter) {
    nativeSetter.call(fileInput, dt.files);
  } else {
    // Fallback: Object.defineProperty (funciona em alguns contextos)
    try {
      Object.defineProperty(fileInput, "files", {
        value: dt.files, writable: false, configurable: true,
      });
    } catch { /* ignora */ }
  }
  fileInput.dispatchEvent(new Event("input",  { bubbles: true }));
  fileInput.dispatchEvent(new Event("change", { bubbles: true }));
}

/**
 * Envia arquivo via drag-and-drop na área do chat.
 * Mais confiável pois ignora restrições do React no file input.
 */
async function dropFileOnChat(file) {
  // Alvos em prioridade: área do chat, footer, #main
  const dropTarget =
    document.querySelector('[data-testid="conversation-compose-box-input"]') ||
    document.querySelector('div[contenteditable="true"]') ||
    document.querySelector("#main") ||
    document.querySelector("footer");
  if (!dropTarget) return false;

  const dt = new DataTransfer();
  dt.items.add(file);

  const evOpts = { bubbles: true, cancelable: true, dataTransfer: dt };
  dropTarget.dispatchEvent(new DragEvent("dragenter", evOpts));
  await delay(120);
  dropTarget.dispatchEvent(new DragEvent("dragover",  evOpts));
  await delay(120);
  dropTarget.dispatchEvent(new DragEvent("drop",      evOpts));
  return true;
}

/** Localiza o input[type=file] "Fotos e vídeos" sem abrir o menu. */
function findPhotosVideosInputDirect() {
  const footer = document.querySelector("footer");
  const inputs = footer
    ? [...footer.querySelectorAll('input[type="file"]')]
    : [...document.querySelectorAll('input[type="file"]')];
  // Preferir input que aceite image E video (canal "Fotos e vídeos")
  const pvInput = inputs.find((i) => {
    const a = (i.getAttribute("accept") || "").toLowerCase();
    return a.includes("image") && a.includes("video") && !a.includes("webp");
  });
  if (pvInput) return pvInput;
  // Fallback: qualquer input que aceite imagem
  return inputs.find((i) => {
    const a = (i.getAttribute("accept") || "").toLowerCase();
    return a.includes("image") || a.includes("jpg") || a.includes("png");
  }) || null;
}

/** Abre menu (+), clica em "Fotos e vídeos" e injeta o arquivo via file input. */
async function openMenuAndInjectFile(file) {
  const footer = document.querySelector("footer");
  if (!footer) return false;

  // Localizar botão +
  let plusBtn = null;
  for (const sel of [
    '[data-testid="clip"]',
    '[data-testid="attach-menu-plus"]',
    'span[data-icon="clip"]',
    'span[data-icon="plus-rounded"]',
    'span[data-icon="plus"]',
  ]) {
    const sp = footer.querySelector(sel);
    const btn = sp?.closest("button") || sp?.closest('div[role="button"]');
    if (btn && btn.offsetParent !== null) { plusBtn = btn; break; }
  }
  if (!plusBtn) {
    for (const b of footer.querySelectorAll('button, div[role="button"]')) {
      const a = (b.getAttribute("aria-label") || "").toLowerCase();
      if ((a.includes("anexar") || a.includes("attach") || a.includes("clip")) && b.offsetParent !== null) {
        plusBtn = b; break;
      }
    }
  }
  if (!plusBtn) return false;

  plusBtn.click();
  await delay(600 + Math.random() * 200);

  // Tentar clicar no item "Fotos e vídeos" do menu para expor o input
  const menu = document.querySelector('[role="menu"]');
  if (menu) {
    const photoLabel = /foto|photo|image|imagem|vídeo|video/i;
    for (const item of [...menu.querySelectorAll('[role="menuitem"], li, button')]) {
      if (!photoLabel.test(item.textContent || "")) continue;
      // Achar e injetar via input filho
      let el = item;
      for (let d = 0; d < 10 && el; d++) {
        const inp = el.querySelector?.('input[type="file"]');
        if (inp) {
          injectFileIntoInput(inp, file);
          return true;
        }
        el = el.parentElement;
      }
    }
  }

  // Fallback: encontrar input direto e injetar
  await delay(300);
  const fileInput = findPhotosVideosInputDirect();
  if (fileInput) {
    injectFileIntoInput(fileInput, file);
    return true;
  }
  return false;
}

/**
 * Aguarda o campo de legenda do preview de mídia do WhatsApp aparecer.
 * Retorna o elemento ou null se timeout.
 */
async function waitForCaptionField(timeoutMs = 22000) {
  const t0 = Date.now();
  while (Date.now() - t0 < timeoutMs) {
    if (!isExtensionContextAlive()) return null;
    const allEditable = [...document.querySelectorAll('div[contenteditable="true"]')]
      .filter((el) => el.offsetParent !== null);
    for (const el of allEditable) {
      const lab = (el.getAttribute("aria-label") || "").toLowerCase();
      // Seletores conhecidos do campo de legenda do WA em PT e EN
      if (/caption|legenda|adicione uma legenda|add a caption|escreva uma mensagem|type a message|digite/i.test(lab)) {
        return el;
      }
    }
    // Detector alternativo: um contenteditable DENTRO de um elemento que contenha img[src^="blob:"]
    const previewArea = document.querySelector('div[class*="media-viewer"], div[data-testid*="media-viewer"]')
      || [...document.querySelectorAll("img[src^='blob:']")]
          .map((img) => img.closest('[class]'))
          .find(Boolean);
    if (previewArea) {
      const innerEdit = previewArea.querySelector('div[contenteditable="true"]');
      if (innerEdit && innerEdit.offsetParent !== null) return innerEdit;
    }
    await delay(350);
  }
  // Último recurso: retornar o contenteditable mais baixo na tela
  const visible = [...document.querySelectorAll('div[contenteditable="true"]')]
    .filter((el) => el.offsetParent !== null);
  if (!visible.length) return null;
  return visible.reduce((best, el) => {
    return el.getBoundingClientRect().bottom > best.getBoundingClientRect().bottom ? el : best;
  });
}

/**
 * Localiza o botão Enviar do modal de preview de imagem.
 * O WA usa um FAB verde no canto inferior direito durante o preview.
 */
function findPreviewSendButton() {
  // Tentar por data-testid específicos do preview
  for (const sel of [
    '[data-testid="media-caption-send-button"]',
    '[data-testid="send"]',
    'span[data-icon="send"]',
    'span[data-icon="wds-ic-send-filled"]',
  ]) {
    const sp = document.querySelector(sel);
    if (!sp) continue;
    const btn = sp.closest("button") || sp.closest('div[role="button"]') || sp;
    if (btn && btn.offsetParent !== null) {
      const r = btn.getBoundingClientRect();
      if (r.width > 4 && r.height > 4) return btn;
    }
  }
  // Fallback: botão Send mais à direita e mais abaixo (FAB do preview)
  const candidates = [...document.querySelectorAll('button, div[role="button"]')].filter((b) => {
    if (b.offsetParent === null) return false;
    const r = b.getBoundingClientRect();
    return r.right > window.innerWidth * 0.6 && r.bottom > window.innerHeight * 0.5 && r.width > 30 && r.height > 30;
  });
  if (!candidates.length) return null;
  return candidates.reduce((best, b) => {
    const rb = b.getBoundingClientRect(), rBest = best.getBoundingClientRect();
    return rb.right + rb.bottom > rBest.right + rBest.bottom ? b : best;
  });
}

/**
 * Envia a imagem com legenda de texto para o contato atual.
 * Estratégia 1 (primária): Drag & Drop — o WA aceita e não usa React para isso.
 * Estratégia 2 (fallback): Abrir menu (+) e injetar via file input com native setter.
 */
async function sendImageWithCaption(attData, captionText) {
  await appendLog("info", "Preparando envio de imagem com legenda...");
  const file = dataUrlToFile(attData.dataUrl, attData.name, attData.mime);

  // ── Estratégia 1: Drag & Drop ──
  let dropped = false;
  try {
    dropped = await dropFileOnChat(file);
  } catch (e) {
    await appendLog("warn", `Drag-drop falhou (${e?.message || e}), tentando menu (+)...`);
  }

  if (dropped) {
    await appendLog("info", `Imagem "${file.name}" enviada via drag-drop. Aguardando preview...`);
  } else {
    // ── Estratégia 2: Menu (+) + file input injection ──
    await appendLog("info", `Drag-drop não respondeu. Abrindo menu (+) para "${file.name}"...`);
    const injected = await openMenuAndInjectFile(file);
    if (!injected) {
      await appendLog("err", "Não foi possível localizar o input de arquivo do WhatsApp.");
      return false;
    }
    await appendLog("info", `Imagem "${file.name}" injetada via file input.`);
  }

  // ── Aguardar campo de legenda ──
  const captionField = await waitForCaptionField(22000);

  // ── Inserir legenda ──
  if (captionField && captionText && captionText.trim()) {
    await delay(400 + Math.random() * 200);
    try {
      await pasteMessageInto(captionField, captionText);
      await appendLog("info", "Legenda inserida no preview.");
    } catch (e) {
      await appendLog("warn", `Falha ao inserir legenda: ${e?.message || e}`);
    }
    await delay(400 + Math.random() * 200);
  } else if (!captionField) {
    await appendLog("warn", "Campo de legenda não encontrado — imagem será enviada sem texto.");
    await delay(800);
  }

  // ── Clicar em Enviar ──
  // Tentar primeiro o botão específico do preview
  let sendBtn = findPreviewSendButton();
  if (!sendBtn) {
    // Fallback: aguardar o botão padrão
    const waitMs = 12000;
    const t0 = Date.now();
    while (Date.now() - t0 < waitMs) {
      const list = gatherSendButtonElements().filter(isSendButtonInteractable);
      if (list.length) { sendBtn = rankSendButtonsForClick(list)[0]; break; }
      sendBtn = findPreviewSendButton();
      if (sendBtn) break;
      await delay(250);
    }
  }

  if (!sendBtn) {
    await appendLog("err", "Botão Enviar não encontrado no preview da imagem.");
    return false;
  }

  await simulateMouseToElement(sendBtn);
  await delay(80 + Math.random() * 80);
  dispatchSyntheticClick(sendBtn);
  await delay(1800 + Math.random() * 600);

  // Segunda tentativa de clique caso o preview ainda esteja aberto
  const secondBtn = findPreviewSendButton();
  if (secondBtn && secondBtn !== sendBtn) {
    dispatchSyntheticClick(secondBtn);
    await delay(1200 + Math.random() * 400);
  }

  await appendLog("info", "Imagem enviada com sucesso.");
  return true;
}

/**
 * Envia a imagem armazenada em wppPendingAttachment como "Fotos e vídeos"
 * com o texto personalizado como legenda.
 * Retorna true em sucesso, false em falha.
 */
async function sendImageWithCaption(attData, captionText) {
  await appendLog("info", "Iniciando envio de imagem com legenda...");

  // 1. Localizar input[file] sem abrir menu (inputs muitas vezes já estão renderizados)
  let fileInput = findPhotosVideosInputDirect();

  // 2. Se não encontrou, abrir o menu (+) e pegar
  if (!fileInput) {
    fileInput = await openAttachMenuAndGetPhotoInput();
  } else {
    // Abrir o menu de qualquer jeito para garantir que o WA está pronto
    const footer = document.querySelector("footer");
    let plusBtn = null;
    for (const sel of ['span[data-icon="clip"]', 'span[data-icon="plus-rounded"]', 'span[data-icon="plus"]', '[data-testid="clip"]']) {
      const sp = footer?.querySelector(sel);
      const btn = sp?.closest("button") || sp?.closest('div[role="button"]');
      if (btn && btn.offsetParent !== null) { plusBtn = btn; break; }
    }
    if (plusBtn) {
      plusBtn.click();
      await delay(400 + Math.random() * 150);
    }
    // Reconfirmar o input após menu abrir
    fileInput = findPhotosVideosInputDirect() || fileInput;
  }

  if (!fileInput) {
    await appendLog("err", "Não encontrou o campo de upload de imagem no WhatsApp.");
    return false;
  }

  // 3. Injetar o arquivo via DataTransfer
  const file = dataUrlToFile(attData.dataUrl, attData.name, attData.mime);
  const dt = new DataTransfer();
  dt.items.add(file);
  Object.defineProperty(fileInput, "files", { value: dt.files, writable: false, configurable: true });
  fileInput.dispatchEvent(new Event("change", { bubbles: true }));

  await appendLog("info", `Imagem "${file.name}" injetada. Aguardando preview...`);

  // 4. Aguardar o preview de mídia aparecer (modal de envio do WA)
  const previewTimeout = 20000;
  const t0 = Date.now();
  let captionField = null;

  while (Date.now() - t0 < previewTimeout) {
    if (!isExtensionContextAlive()) return false;
    // Procurar o campo de legenda
    const allEditable = [...document.querySelectorAll('div[contenteditable="true"]')];
    for (const el of allEditable) {
      const lab = (el.getAttribute("aria-label") || "").toLowerCase();
      if (
        /caption|legenda|adicione|add a|escreva|type a|digite/i.test(lab) &&
        el.offsetParent !== null
      ) {
        captionField = el;
        break;
      }
    }
    if (captionField) break;
    await delay(300);
  }

  if (!captionField) {
    // Tentar sem label — pegar o editable mais baixo na tela (preview abre novo campo)
    const allEditable = [...document.querySelectorAll('div[contenteditable="true"]')]
      .filter((el) => el.offsetParent !== null);
    if (allEditable.length) {
      captionField = allEditable.reduce((best, el) => {
        const rb = el.getBoundingClientRect();
        const rBest = best.getBoundingClientRect();
        return rb.bottom > rBest.bottom ? el : best;
      });
    }
  }

  // 5. Colar o texto de legenda (se houver)
  if (captionField && captionText && captionText.trim()) {
    await delay(300 + Math.random() * 200);
    try {
      await pasteMessageInto(captionField, captionText);
      await appendLog("info", "Legenda inserida.");
    } catch (e) {
      await appendLog("warn", `Falha ao inserir legenda: ${e?.message || e}`);
    }
  }

  await delay(600 + Math.random() * 400);

  // 6. Clicar no botão de Enviar do preview (FAB direito inferior)
  const sendOk = await clickSendWhenReady();
  if (!sendOk) {
    await appendLog("err", "Botão Enviar não encontrado no preview da imagem.");
    return false;
  }

  await delay(1000 + Math.random() * 500);
  return true;
}

/**
 * Cola a mensagem inteira de uma vez (rápido).
 * Ordem: clipboard + colar → insertText completo → fallback DOM.
 */
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

  // ── Verificar se h\u00e1 imagem pendente para enviar como legenda ──
  const attData = await safeLocalGet([K.pendingAttachment]);
  const att = attData[K.pendingAttachment];

  if (att && att.dataUrl) {
    // Fluxo: imagem + legenda de texto
    const imgSent = await sendImageWithCaption(att, personalized);
    if (!imgSent) {
      await appendLog("err", `Falha ao enviar imagem para ${row.phone}.`);
      await recordResult(row.phone, row.name, "Falha");
      await bumpFailed();
      await delayBetweenMessages();
      if (!(await getDispatchState())[K.active]) return;
      await goToNextOrEnd(idx + 1, queue);
      return;
    }
  } else {
    // Fluxo: apenas texto
    try {
      await pasteMessageInto(composer, personalized);
    } catch (e) {
      await appendLog(
        "err",
        `Falha ao digitar/enviar mensagem: ${row.phone}${e && e.message ? ` — ${e.message}` : ""}`
      );
      await recordResult(row.phone, row.name, "Falha");
      await bumpFailed();
      await delayBetweenMessages();
      if (!(await getDispatchState())[K.active]) return;
      await goToNextOrEnd(idx + 1, queue);
      return;
    }

    await delayShortHuman();
    const sentOk = await clickSendWhenReady();
    if (!sentOk) {
      await appendLog("err", `Envio n\u00e3o conclu\u00eddo (bot\u00e3o Enviar) — ${row.phone}`);
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
