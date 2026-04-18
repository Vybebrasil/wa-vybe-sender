const STORAGE_KEYS = {
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
  dialPrefix: "wppDialPrefix",
  antiPreset: "wppAntiPreset",
  antiBlocks: "wppAntiBlocks",
  antiBlockSize: "wppAntiBlockSize",
  antiBlockWait: "wppAntiBlockWait",
  antiDedupe: "wppAntiDedupe",
  antiRandom: "wppAntiRandom",
  antiIntervalOn: "wppAntiIntervalOn",
  batchCount: "wppBatchSentInBlock",
  messageTemplates: "wppMessageTemplates",
  pendingAttachment: "wppPendingAttachment",
};

const EMOJI_PALETTE =
  "😀 😃 😄 😁 😅 😂 🤣 😊 😍 🥰 😘 😋 😎 🤩 🥳 😇 🙂 😉 😌 😢 😭 😤 🤔 😬 🙏 👍 👎 👏 🙌 👋 ✨ 🔥 💯 ❤️ 🧡 💚 💙 🎉 ✅ ⚠️ 📌 💬 📱 💼 🚀".split(
    /\s+/
  );

const ANTIBLOCK_PRESETS = {
  conservative: {
    blocks: true,
    blockSize: 50,
    blockWait: 10,
    dedupe: true,
    random: true,
    intervalOn: true,
    minSec: 20,
    maxSec: 40,
  },
  balanced: {
    blocks: true,
    blockSize: 100,
    blockWait: 2,
    dedupe: true,
    random: true,
    intervalOn: true,
    minSec: 5,
    maxSec: 15,
  },
  fast: {
    blocks: false,
    blockSize: 50,
    blockWait: 5,
    dedupe: true,
    random: true,
    intervalOn: true,
    minSec: 0.5,
    maxSec: 2,
  },
};

const ANTIBLOCK_TITLES = {
  conservative: "Conservador",
  balanced: "Equilibrado",
  fast: "Rápido",
  custom: "Personalizado",
};

const $ = (id) => document.getElementById(id);

const COUNTRIES = [
  { dial: "55", label: "Brasil", flag: "🇧🇷" },
  { dial: "351", label: "Portugal", flag: "🇵🇹" },
  { dial: "1", label: "EUA / Canadá", flag: "🇺🇸" },
  { dial: "54", label: "Argentina", flag: "🇦🇷" },
  { dial: "598", label: "Uruguai", flag: "🇺🇾" },
  { dial: "244", label: "Angola", flag: "🇦🇴" },
  { dial: "258", label: "Moçambique", flag: "🇲🇿" },
  { dial: "49", label: "Alemanha", flag: "🇩🇪" },
  { dial: "34", label: "Espanha", flag: "🇪🇸" },
];

function escapeHtml(s) {
  const d = document.createElement("div");
  d.textContent = s;
  return d.innerHTML;
}

function getDialDigits() {
  const btn = $("countryBtn");
  const d = btn && btn.dataset.dial;
  return (d && String(d).replace(/\D/g, "")) || "55";
}

/** Se o número já começa com o DDI escolhido, mantém; senão prefixa (ex.: 11… → 5511…) */
function applyDialToPhone(digits, dial) {
  const d = String(dial || "55").replace(/\D/g, "");
  let p = String(digits || "").replace(/\D/g, "");
  if (!p) return "";
  if (p.startsWith(d)) return p;
  const local = p.replace(/^0+/, "");
  return d + local;
}

function parseQueue(raw) {
  const dial = getDialDigits();
  const dedupeOn = !$("antiDedupe") || $("antiDedupe").checked;
  const lines = raw.split("\n").map((l) => l.trim()).filter(Boolean);
  const queue = [];
  const seen = new Set();
  for (const line of lines) {
    const idx = line.indexOf(",");
    let phone = line.replace(/\D/g, "");
    let name = "";
    if (idx !== -1) {
      phone = line.slice(0, idx).replace(/\D/g, "");
      name = line.slice(idx + 1).trim();
    }
    phone = applyDialToPhone(phone, dial);
    if (phone.length < 10) continue;
    if (dedupeOn && seen.has(phone)) continue;
    seen.add(phone);
    queue.push({ phone, name });
  }
  return queue;
}

function updateCountryUi(dialDigits) {
  const d = String(dialDigits || "55").replace(/\D/g, "") || "55";
  const c = COUNTRIES.find((x) => x.dial === d) || {
    dial: d,
    label: "DDI",
    flag: "🌐",
  };
  const btn = $("countryBtn");
  const flagEl = $("countryFlag");
  const dialEl = $("dialLabel");
  const hintEl = $("dialHint");
  if (btn) btn.dataset.dial = c.dial;
  if (flagEl) flagEl.textContent = c.flag;
  if (dialEl) dialEl.textContent = `+${c.dial}`;
  if (hintEl) hintEl.textContent = `+${c.dial}`;
}

function closeCountryMenu() {
  const menu = $("countryMenu");
  const btn = $("countryBtn");
  if (menu) menu.classList.add("hidden");
  if (btn) btn.setAttribute("aria-expanded", "false");
}

function toggleCountryMenu(e) {
  if (e) e.stopPropagation();
  const menu = $("countryMenu");
  const btn = $("countryBtn");
  if (!menu || !btn) return;
  const opening = menu.classList.contains("hidden");
  menu.classList.toggle("hidden");
  btn.setAttribute("aria-expanded", String(opening));
}

function buildCountryMenu() {
  const ul = $("countryMenu");
  if (!ul) return;
  ul.innerHTML = "";
  COUNTRIES.forEach((c) => {
    const li = document.createElement("li");
    li.setAttribute("role", "option");
    li.dataset.dial = c.dial;
    li.innerHTML = `<span class="m-flag">${c.flag}</span><span class="m-dial">+${c.dial}</span><span>${escapeHtml(
      c.label
    )}</span>`;
    li.addEventListener("click", (ev) => {
      ev.stopPropagation();
      updateCountryUi(c.dial);
      chrome.storage.local.set({ [STORAGE_KEYS.dialPrefix]: c.dial });
      closeCountryMenu();
    });
    ul.appendChild(li);
  });
}

function renderLogs(logs) {
  const box = $("logBox");
  if (!Array.isArray(logs) || logs.length === 0) {
    box.innerHTML = '<div class="log-line info">Nenhum evento ainda.</div>';
    return;
  }
  box.innerHTML = logs
    .slice(-100)
    .map((e) => {
      const t = new Date(e.time).toLocaleTimeString("pt-BR", {
        hour: "2-digit",
        minute: "2-digit",
      });
      return `<div class="log-line ${e.level || "info"}">[${t}] ${escapeHtml(
        e.text
      )}</div>`;
    })
    .join("");
  box.scrollTop = box.scrollHeight;
}

function formatSecLabel(s) {
  const n = Number(s);
  if (Number.isNaN(n)) return "?";
  if (n < 1 && n >= 0.3) return `${n} seg.`;
  return `${Math.round(n)} seg.`;
}

function readAntiIntervals() {
  let minS = parseFloat($("antiMinSec")?.value);
  let maxS = parseFloat($("antiMaxSec")?.value);
  if (Number.isNaN(minS)) minS = 5;
  if (Number.isNaN(maxS)) maxS = 15;
  minS = Math.max(0.3, Math.min(120, minS));
  maxS = Math.max(0.3, Math.min(180, maxS));
  if (maxS < minS) maxS = minS;
  if ($("antiMinSec")) $("antiMinSec").value = String(minS);
  if ($("antiMaxSec")) $("antiMaxSec").value = String(maxS);
  return { minS, maxS };
}

function updateAntiInterSummary() {
  const { minS, maxS } = readAntiIntervals();
  const el = $("antiInterSummary");
  if (el) el.textContent = `de ${formatSecLabel(minS)} a ${formatSecLabel(maxS)}`;
}

function updateAntiBlockSummary() {
  const sz = parseInt($("antiBlockSize")?.value, 10) || 100;
  const w = parseInt($("antiBlockWait")?.value, 10) || 2;
  const el = $("antiBlockSummary");
  if (el) {
    el.textContent = `Blocos de ${sz} mensagens · ${w} min entre blocos`;
  }
}

function clearAntiPresetActive() {
  document.querySelectorAll(".anti-pill").forEach((p) => p.classList.remove("active"));
}

function applyAntiPreset(key) {
  const p = ANTIBLOCK_PRESETS[key];
  if (!p) return;
  if ($("antiBlocks")) $("antiBlocks").checked = p.blocks;
  if ($("antiDedupe")) $("antiDedupe").checked = p.dedupe;
  if ($("antiRandom")) $("antiRandom").checked = p.random;
  if ($("antiIntervalOn")) $("antiIntervalOn").checked = p.intervalOn;
  if ($("antiBlockSize")) $("antiBlockSize").value = String(p.blockSize);
  if ($("antiBlockWait")) $("antiBlockWait").value = String(p.blockWait);
  if ($("antiMinSec")) $("antiMinSec").value = String(p.minSec);
  if ($("antiMaxSec")) $("antiMaxSec").value = String(p.maxSec);
  clearAntiPresetActive();
  const btn = document.querySelector(`.anti-pill[data-preset="${key}"]`);
  if (btn) btn.classList.add("active");
  updateAntiInterSummary();
  updateAntiBlockSummary();
  updateAntiTitle(key);
}

function updateAntiTitle(presetKey) {
  const t = $("antiTitle");
  if (!t) return;
  const name = ANTIBLOCK_TITLES[presetKey] || ANTIBLOCK_TITLES.custom;
  t.textContent = `Antibloqueio: ${name}`;
}

function markAntiCustom() {
  clearAntiPresetActive();
  updateAntiTitle("custom");
}

function collectAntiForStorage() {
  const { minS, maxS } = readAntiIntervals();
  const presetBtn = document.querySelector(".anti-pill.active");
  const preset = presetBtn?.dataset?.preset || "custom";
  return {
    [STORAGE_KEYS.antiPreset]: preset,
    [STORAGE_KEYS.antiBlocks]: !!$("antiBlocks")?.checked,
    [STORAGE_KEYS.antiBlockSize]: Math.max(5, parseInt($("antiBlockSize")?.value, 10) || 100),
    [STORAGE_KEYS.antiBlockWait]: Math.max(1, parseInt($("antiBlockWait")?.value, 10) || 2),
    [STORAGE_KEYS.antiDedupe]: !!$("antiDedupe")?.checked,
    [STORAGE_KEYS.antiRandom]: !!$("antiRandom")?.checked,
    [STORAGE_KEYS.antiIntervalOn]: !!$("antiIntervalOn")?.checked,
    [STORAGE_KEYS.minSec]: minS,
    [STORAGE_KEYS.maxSec]: maxS,
  };
}

function persistAntiPartial() {
  const o = collectAntiForStorage();
  chrome.storage.local.set(o);
}

function updateStats(stats) {
  if (!stats) return;
  $("statSent").textContent = String(stats.sent ?? 0);
  $("statFailed").textContent = String(stats.failed ?? 0);
  $("statTotal").textContent = String(stats.total ?? 0);
  const total = stats.total || 0;
  const done = (stats.sent ?? 0) + (stats.failed ?? 0);
  const pct = total > 0 ? Math.min(100, Math.round((done / total) * 100)) : 0;
  $("progressFill").style.width = `${pct}%`;
  $("progressLabel").textContent = `${pct}%`;
}

function setStatusPill(active, paused) {
  const el = $("statusPill");
  el.classList.remove("running", "paused", "idle");
  if (active && paused) {
    el.classList.add("paused");
    el.textContent = "Pausado — use Retomar no botão Pausar ou exporte o CSV";
  } else if (active) {
    el.classList.add("running");
    el.textContent = "Disparo em execução";
  } else {
    el.classList.add("idle");
    el.textContent = "Pronto para iniciar";
  }
}

function setButtons(active, paused) {
  $("btnStart").disabled = active;
  $("btnPause").disabled = !active;
  const running = active && !paused;
  $("btnRestart").disabled = running;
  $("btnPause").textContent = active && paused ? "Retomar" : "Pausar";
}

function updateExportButton(active, paused, resultsLength) {
  const hasRows = resultsLength > 0;
  const endedOrPaused = !active || (active && paused);
  $("btnExport").disabled = !(hasRows && endedOrPaused);
}

function csvEscape(value) {
  const s = value == null ? "" : String(value);
  return `"${s.replace(/"/g, '""')}"`;
}

function buildCsv(rows) {
  const header = ["Número", "Nome", "Status", "Timestamp"];
  const lines = [header.map(csvEscape).join(",")];
  for (const r of rows) {
    lines.push(
      [
        csvEscape(r.phone),
        csvEscape(r.name || ""),
        csvEscape(r.status),
        csvEscape(r.timestamp),
      ].join(",")
    );
  }
  return "\uFEFF" + lines.join("\r\n");
}

async function loadUiFromStorage() {
  const data = await chrome.storage.local.get([
    STORAGE_KEYS.queue,
    STORAGE_KEYS.stats,
    STORAGE_KEYS.logs,
    STORAGE_KEYS.active,
    STORAGE_KEYS.paused,
    STORAGE_KEYS.template,
    STORAGE_KEYS.index,
    STORAGE_KEYS.results,
    STORAGE_KEYS.minSec,
    STORAGE_KEYS.maxSec,
    STORAGE_KEYS.dialPrefix,
    STORAGE_KEYS.antiPreset,
    STORAGE_KEYS.antiBlocks,
    STORAGE_KEYS.antiBlockSize,
    STORAGE_KEYS.antiBlockWait,
    STORAGE_KEYS.antiDedupe,
    STORAGE_KEYS.antiRandom,
    STORAGE_KEYS.antiIntervalOn,
  ]);

  if (Array.isArray(data[STORAGE_KEYS.logs])) renderLogs(data[STORAGE_KEYS.logs]);
  else renderLogs([]);

  updateStats(data[STORAGE_KEYS.stats] || { sent: 0, failed: 0, total: 0 });

  const active = !!data[STORAGE_KEYS.active];
  const paused = !!data[STORAGE_KEYS.paused];
  const results = Array.isArray(data[STORAGE_KEYS.results])
    ? data[STORAGE_KEYS.results]
    : [];

  setButtons(active, paused);
  setStatusPill(active, paused);
  updateExportButton(active, paused, results.length);

  if (data[STORAGE_KEYS.template]) $("message").value = data[STORAGE_KEYS.template];
  if (Array.isArray(data[STORAGE_KEYS.queue]) && data[STORAGE_KEYS.queue].length) {
    $("numbers").value = data[STORAGE_KEYS.queue]
      .map((r) => (r.name ? `${r.phone}, ${r.name}` : r.phone))
      .join("\n");
  }

  const preset = data[STORAGE_KEYS.antiPreset];
  const firstAntiRun =
    data[STORAGE_KEYS.antiPreset] == null && data[STORAGE_KEYS.antiBlocks] === undefined;
  if (preset && ANTIBLOCK_PRESETS[preset]) {
    applyAntiPreset(preset);
  } else if (firstAntiRun) {
    applyAntiPreset("balanced");
    chrome.storage.local.set(collectAntiForStorage());
  } else {
    if ($("antiBlocks")) $("antiBlocks").checked = !!data[STORAGE_KEYS.antiBlocks];
    if ($("antiDedupe")) $("antiDedupe").checked = data[STORAGE_KEYS.antiDedupe] !== false;
    if ($("antiRandom")) $("antiRandom").checked = data[STORAGE_KEYS.antiRandom] !== false;
    if ($("antiIntervalOn")) $("antiIntervalOn").checked = data[STORAGE_KEYS.antiIntervalOn] !== false;
    if ($("antiBlockSize") && data[STORAGE_KEYS.antiBlockSize] != null) {
      $("antiBlockSize").value = String(data[STORAGE_KEYS.antiBlockSize]);
    }
    if ($("antiBlockWait") && data[STORAGE_KEYS.antiBlockWait] != null) {
      $("antiBlockWait").value = String(data[STORAGE_KEYS.antiBlockWait]);
    }
    const minV = data[STORAGE_KEYS.minSec];
    const maxV = data[STORAGE_KEYS.maxSec];
    if ($("antiMinSec") && minV != null) $("antiMinSec").value = String(minV);
    if ($("antiMaxSec") && maxV != null) $("antiMaxSec").value = String(maxV);
    clearAntiPresetActive();
    updateAntiTitle(preset || "custom");
  }
  updateAntiInterSummary();
  updateAntiBlockSummary();

  const savedDial = data[STORAGE_KEYS.dialPrefix];
  updateCountryUi(
    savedDial != null ? String(savedDial).replace(/\D/g, "") : "55"
  );

  // Restaurar preview do anexo se houver imagem salva
  await refreshAttachPreview();
}

async function findWhatsAppTab() {
  const byPattern = await chrome.tabs.query({ url: "https://web.whatsapp.com/*" });
  if (byPattern[0]) return byPattern[0];
  const all = await chrome.tabs.query({});
  return (
    all.find(
      (t) =>
        t.url &&
        (t.url.startsWith("https://web.whatsapp.com/") ||
          t.url.startsWith("https://web.whatsapp.com?"))
    ) || null
  );
}

async function forceNavigateTab(tabId, url) {
  try {
    const r = await chrome.runtime.sendMessage({
      type: "WPP_NAVIGATE_TAB",
      tabId,
      url,
    });
    if (r && r.ok) return true;
  } catch {
    /* service worker inativo */
  }
  try {
    await chrome.tabs.update(tabId, { url });
    return true;
  } catch {
    return false;
  }
}

$("btnStart").addEventListener("click", async () => {
  const queue = parseQueue($("numbers").value);
  const template = $("message").value.trim();
  if (!queue.length || !template) {
    alert("Informe ao menos um número válido e a mensagem.");
    return;
  }
  const tab = await findWhatsAppTab();
  if (!tab?.id) {
    alert("Abra o WhatsApp Web (https://web.whatsapp.com) em uma aba antes de iniciar.");
    return;
  }

  const antiPayload = collectAntiForStorage();
  const stats = { sent: 0, failed: 0, total: queue.length };
  const logs = [
    {
      time: Date.now(),
      level: "info",
      text: `Iniciando envio — ${queue.length} contato(s) na fila.`,
    },
  ];

  await chrome.storage.local.set({
    [STORAGE_KEYS.queue]: queue,
    [STORAGE_KEYS.template]: template,
    [STORAGE_KEYS.index]: 0,
    [STORAGE_KEYS.active]: true,
    [STORAGE_KEYS.paused]: false,
    [STORAGE_KEYS.stats]: stats,
    [STORAGE_KEYS.logs]: logs,
    [STORAGE_KEYS.results]: [],
    [STORAGE_KEYS.navPending]: {
      phone: queue[0].phone.replace(/\D/g, ""),
      at: Date.now(),
    },
    [STORAGE_KEYS.dialPrefix]: getDialDigits(),
    [STORAGE_KEYS.batchCount]: 0,
    ...antiPayload,
  });

  setButtons(true, false);
  setStatusPill(true, false);
  updateStats(stats);
  updateExportButton(true, false, 0);
  renderLogs(logs);

  const firstUrl = `https://web.whatsapp.com/send?phone=${encodeURIComponent(
    queue[0].phone
  )}`;
  const navigated = await forceNavigateTab(tab.id, firstUrl);
  if (!navigated) {
    await chrome.storage.local.set({
      [STORAGE_KEYS.active]: false,
      [STORAGE_KEYS.paused]: false,
    });
    await chrome.storage.local.remove(STORAGE_KEYS.navPending);
    alert(
      "Não foi possível abrir a conversa na aba do WhatsApp. Recarregue a página do WhatsApp Web e tente de novo."
    );
    setButtons(false, false);
    setStatusPill(false, false);
    return;
  }

  logs.push({
    time: Date.now(),
    level: "info",
    text: "Abrindo conversa no WhatsApp Web…",
  });
  await chrome.storage.local.set({ [STORAGE_KEYS.logs]: logs });
  renderLogs(logs);

  /* O content script reage ao storage (debounced); WPP_KICK duplicava o disparo. */
});

$("btnPause").addEventListener("click", async () => {
  const cur = await chrome.storage.local.get([
    STORAGE_KEYS.active,
    STORAGE_KEYS.paused,
  ]);
  if (!cur[STORAGE_KEYS.active]) return;
  const nextPaused = !cur[STORAGE_KEYS.paused];
  await chrome.storage.local.set({ [STORAGE_KEYS.paused]: nextPaused });
  setButtons(true, nextPaused);
  setStatusPill(true, nextPaused);
  const res = await chrome.storage.local.get(STORAGE_KEYS.results);
  const results = Array.isArray(res[STORAGE_KEYS.results])
    ? res[STORAGE_KEYS.results]
    : [];
  updateExportButton(true, nextPaused, results.length);
});

$("btnRestart").addEventListener("click", async () => {
  const data = await chrome.storage.local.get([STORAGE_KEYS.queue]);
  let queue = Array.isArray(data[STORAGE_KEYS.queue]) ? data[STORAGE_KEYS.queue] : [];
  if (!queue.length) queue = parseQueue($("numbers").value);
  const total = queue.length;
  const antiPayload = collectAntiForStorage();

  const logs = [
    {
      time: Date.now(),
      level: "info",
      text: "Campanha reiniciada — índice e relatório da sessão zerados.",
    },
  ];

  await chrome.storage.local.set({
    [STORAGE_KEYS.active]: false,
    [STORAGE_KEYS.paused]: false,
    [STORAGE_KEYS.index]: 0,
    [STORAGE_KEYS.stats]: { sent: 0, failed: 0, total },
    [STORAGE_KEYS.results]: [],
    [STORAGE_KEYS.logs]: logs,
    [STORAGE_KEYS.queue]: queue.length ? queue : [],
    [STORAGE_KEYS.batchCount]: 0,
    ...antiPayload,
  });
  await chrome.storage.local.remove(STORAGE_KEYS.navPending);
  await chrome.storage.local.remove(STORAGE_KEYS.pendingAttachment);
  await refreshAttachPreview();

  setButtons(false, false);
  setStatusPill(false, false);
  updateExportButton(false, false, 0);
  updateStats({ sent: 0, failed: 0, total });
  renderLogs(logs);
});

$("btnExport").addEventListener("click", async () => {
  const data = await chrome.storage.local.get(STORAGE_KEYS.results);
  const rows = Array.isArray(data[STORAGE_KEYS.results])
    ? data[STORAGE_KEYS.results]
    : [];
  if (!rows.length) return;

  const csv = buildCsv(rows);
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  const stamp = new Date().toISOString().slice(0, 19).replace(/[:T]/g, "-");
  a.href = url;
  a.download = `relatorio-disparos-${stamp}.csv`;
  a.click();
  URL.revokeObjectURL(url);
});

function insertAroundSelection(prefix, suffix) {
  const ta = $("message");
  if (!ta) return;
  const start = ta.selectionStart ?? 0;
  const end = ta.selectionEnd ?? 0;
  const v = ta.value;
  const sel = v.slice(start, end);
  const mid =
    sel ||
    (prefix === "*" || prefix === "_" || prefix === "~" ? "texto" : "");
  ta.value = v.slice(0, start) + prefix + mid + suffix + v.slice(end);
  const caret = start + prefix.length + mid.length + suffix.length;
  ta.selectionStart = ta.selectionEnd = caret;
  ta.focus();
}

function insertAtCursor(str) {
  const ta = $("message");
  if (!ta) return;
  const start = ta.selectionStart ?? 0;
  const end = ta.selectionEnd ?? 0;
  const v = ta.value;
  ta.value = v.slice(0, start) + str + v.slice(end);
  const pos = start + str.length;
  ta.selectionStart = ta.selectionEnd = pos;
  ta.focus();
}

function renderPreviewHtml(text) {
  let h = escapeHtml(text);
  h = h.replace(/\*([^*]+)\*/g, "<strong>$1</strong>");
  h = h.replace(/_([^_]+)_/g, "<em>$1</em>");
  h = h.replace(/~([^~]+)~/g, "<del>$1</del>");
  return h.replace(/\n/g, "<br/>");
}

function buildEmojiGrid() {
  const g = $("emojiGrid");
  if (!g) return;
  g.innerHTML = "";
  EMOJI_PALETTE.forEach((em) => {
    if (!em) return;
    const b = document.createElement("button");
    b.type = "button";
    b.textContent = em;
    b.addEventListener("click", (e) => {
      e.stopPropagation();
      insertAtCursor(em);
      $("emojiPopover")?.classList.remove("open");
    });
    g.appendChild(b);
  });
}

function initMessageComposer() {
  buildEmojiGrid();
  $("fmtBold")?.addEventListener("click", () => insertAroundSelection("*", "*"));
  $("fmtItalic")?.addEventListener("click", () => insertAroundSelection("_", "_"));
  $("fmtStrike")?.addEventListener("click", () => insertAroundSelection("~", "~"));

  $("btnPreviewMsg")?.addEventListener("click", () => {
    const body = $("previewBody");
    const ta = $("message");
    if (body && ta) {
      body.innerHTML = renderPreviewHtml(ta.value);
      $("previewBackdrop")?.classList.add("open");
    }
  });

  $("previewBackdrop")?.addEventListener("click", (e) => {
    if (e.target === $("previewBackdrop")) $("previewBackdrop")?.classList.remove("open");
  });

  $("btnEmoji")?.addEventListener("click", (e) => {
    e.stopPropagation();
    $("emojiPopover")?.classList.toggle("open");
  });

  $("emojiPopover")?.addEventListener("click", (e) => e.stopPropagation());

  document.addEventListener("click", (e) => {
    if (!e.target.closest(".msg-composer-pos")) {
      $("emojiPopover")?.classList.remove("open");
    }
  });

  $("btnSaveTpl")?.addEventListener("click", async () => {
    const text = $("message")?.value?.trim();
    if (!text) {
      alert("Escreva um texto antes de salvar o modelo.");
      return;
    }
    const name = prompt("Nome do modelo?", "Meu modelo");
    if (!name?.trim()) return;
    const data = await chrome.storage.local.get(STORAGE_KEYS.messageTemplates);
    const list = Array.isArray(data[STORAGE_KEYS.messageTemplates])
      ? [...data[STORAGE_KEYS.messageTemplates]]
      : [];
    list.push({ id: Date.now(), name: name.trim(), text });
    await chrome.storage.local.set({
      [STORAGE_KEYS.messageTemplates]: list.slice(-30),
    });
    alert("Modelo salvo.");
  });

  $("btnLoadTpl")?.addEventListener("click", async () => {
    const data = await chrome.storage.local.get(STORAGE_KEYS.messageTemplates);
    const list = Array.isArray(data[STORAGE_KEYS.messageTemplates])
      ? data[STORAGE_KEYS.messageTemplates]
      : [];
    const box = $("tplList");
    if (!box) return;
    box.innerHTML = "";
    if (!list.length) {
      box.innerHTML =
        '<p class="hint" style="margin:0">Nenhum modelo salvo. Use o botão 📄+ para criar.</p>';
    } else {
      [...list].reverse().forEach((t) => {
        const b = document.createElement("button");
        b.type = "button";
        b.className = "msg-tpl-item";
        b.textContent = t.name || "Sem nome";
        b.addEventListener("click", () => {
          if ($("message")) $("message").value = t.text || "";
          $("tplBackdrop")?.classList.remove("open");
        });
        box.appendChild(b);
      });
    }
    $("tplBackdrop")?.classList.add("open");
  });

  $("tplBackdrop")?.addEventListener("click", (e) => {
    if (e.target === $("tplBackdrop")) $("tplBackdrop")?.classList.remove("open");
  });

  document.querySelectorAll(".msg-modal-close").forEach((btn) => {
    btn.addEventListener("click", () => {
      btn.closest(".msg-modal-backdrop")?.classList.remove("open");
    });
  });

}

function formatFileSize(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}

async function refreshAttachPreview() {
  const data = await chrome.storage.local.get(STORAGE_KEYS.pendingAttachment);
  const att = data[STORAGE_KEYS.pendingAttachment];
  const wrap = $("attachPreviewWrap");
  const thumb = $("attachThumb");
  const nameEl = $("attachName");
  const sizeEl = $("attachSize");
  if (!wrap) return;
  if (att && att.dataUrl) {
    thumb.src = att.dataUrl;
    nameEl.textContent = att.name || "imagem";
    sizeEl.textContent = att.size ? formatFileSize(att.size) : "";
    wrap.classList.add("has-image");
  } else {
    thumb.src = "";
    nameEl.textContent = "";
    sizeEl.textContent = "";
    wrap.classList.remove("has-image");
  }
}

async function clearAttachment() {
  await chrome.storage.local.remove(STORAGE_KEYS.pendingAttachment);
  const inp = $("attachInput");
  if (inp) inp.value = "";
  await refreshAttachPreview();
}

function initAttachmentButton() {
  const MAX_BYTES = 1.5 * 1024 * 1024; // 1.5 MB

  $("btnAttach")?.addEventListener("click", () => {
    $("attachInput")?.click();
  });

  $("attachInput")?.addEventListener("change", async () => {
    const inp = $("attachInput");
    const file = inp?.files?.[0];
    if (!file) return;
    if (file.size > MAX_BYTES) {
      alert(`Imagem muito grande (${formatFileSize(file.size)}). Máximo: 1,5 MB.`);
      inp.value = "";
      return;
    }
    const reader = new FileReader();
    reader.onload = async (e) => {
      const dataUrl = e.target.result;
      await chrome.storage.local.set({
        [STORAGE_KEYS.pendingAttachment]: {
          name: file.name,
          mime: file.type || "image/jpeg",
          size: file.size,
          dataUrl,
        },
      });
      await refreshAttachPreview();
    };
    reader.readAsDataURL(file);
  });

  $("btnAttachRemove")?.addEventListener("click", async () => {
    await clearAttachment();
  });
}

function initAntiblockPanel() {
  $("antiToggleDetail")?.addEventListener("click", () => {
    const det = $("antiDetails");
    const btn = $("antiToggleDetail");
    if (!det || !btn) return;
    det.classList.toggle("hidden");
    const open = !det.classList.contains("hidden");
    btn.textContent = open ? "Ocultar detalhe ▲" : "Mostrar detalhe ▼";
  });
  $("antiPresets")?.addEventListener("click", (e) => {
    const pill = e.target.closest(".anti-pill");
    if (!pill?.dataset?.preset) return;
    applyAntiPreset(pill.dataset.preset);
    persistAntiPartial();
  });
  [
    "antiBlocks",
    "antiDedupe",
    "antiRandom",
    "antiIntervalOn",
  ].forEach((id) => {
    $(id)?.addEventListener("change", () => {
      markAntiCustom();
      persistAntiPartial();
    });
  });
  ["antiBlockSize", "antiBlockWait"].forEach((id) => {
    $(id)?.addEventListener("input", () => {
      updateAntiBlockSummary();
      markAntiCustom();
      persistAntiPartial();
    });
  });
  ["antiMinSec", "antiMaxSec"].forEach((id) => {
    $(id)?.addEventListener("input", () => {
      readAntiIntervals();
      updateAntiInterSummary();
      markAntiCustom();
      persistAntiPartial();
    });
  });
  $("antiDedupeEdit")?.addEventListener("click", (e) => {
    e.preventDefault();
    $("numbers")?.focus();
  });
  $("antiHelp")?.addEventListener("click", () => {
    alert(
      "Perfis ajustam tempos e opções de uma vez.\n\n• Conservador: intervalos maiores e blocos menores.\n• Equilibrado: ritmo médio.\n• Rápido: menos espera (maior risco).\n\nRespeite as regras do WhatsApp — uso responsável é sua responsabilidade."
    );
  });
}

chrome.storage.onChanged.addListener((changes, area) => {
  if (area !== "local") return;
  if (changes[STORAGE_KEYS.logs]) {
    renderLogs(changes[STORAGE_KEYS.logs].newValue || []);
  }
  if (changes[STORAGE_KEYS.stats]) {
    updateStats(changes[STORAGE_KEYS.stats].newValue);
  }
  if (changes[STORAGE_KEYS.active] || changes[STORAGE_KEYS.paused]) {
    chrome.storage.local
      .get([
        STORAGE_KEYS.active,
        STORAGE_KEYS.paused,
        STORAGE_KEYS.results,
      ])
      .then((d) => {
        const active = !!d[STORAGE_KEYS.active];
        const paused = !!d[STORAGE_KEYS.paused];
        const results = Array.isArray(d[STORAGE_KEYS.results])
          ? d[STORAGE_KEYS.results]
          : [];
        setButtons(active, paused);
        setStatusPill(active, paused);
        updateExportButton(active, paused, results.length);
      });
  }
  if (changes[STORAGE_KEYS.results]) {
    chrome.storage.local
      .get([
        STORAGE_KEYS.active,
        STORAGE_KEYS.paused,
        STORAGE_KEYS.results,
      ])
      .then((d) => {
        updateExportButton(
          !!d[STORAGE_KEYS.active],
          !!d[STORAGE_KEYS.paused],
          (Array.isArray(d[STORAGE_KEYS.results]) && d[STORAGE_KEYS.results].length) || 0
        );
      });
  }
});

function initPhoneCountryUi() {
  buildCountryMenu();
  const btn = $("countryBtn");
  if (btn) {
    btn.addEventListener("click", (e) => toggleCountryMenu(e));
  }
  document.addEventListener("click", (e) => {
    const wrap = document.querySelector(".phone-input-wrap");
    if (wrap && !wrap.contains(e.target)) closeCountryMenu();
  });
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", () => {
    initPhoneCountryUi();
    initAntiblockPanel();
    initMessageComposer();
    initAttachmentButton();
    loadUiFromStorage();
  });
} else {
  initPhoneCountryUi();
  initAntiblockPanel();
  initMessageComposer();
  initAttachmentButton();
  loadUiFromStorage();
}
