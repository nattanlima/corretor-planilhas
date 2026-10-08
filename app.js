/* Tratador de Planilhas para WhatsApp
 * Roda 100% no navegador. Usa SheetJS (xlsx) carregado via CDN.
 * Fluxo: soltar a planilha -> mapeamento aplicado (lembrado do último uso)
 * -> resultado recalculado a cada ajuste -> baixar o Modelo de Disparo.
 */
(function () {
  "use strict";

  // ---------- Modelo de Disparo ----------

  // Cabeçalho idêntico ao "Modelo de Disparo.xlsx" (colunas A–P, aba Página1).
  // Índice 10 (coluna K) é intencionalmente vazio, como no arquivo original.
  const MODEL_HEADERS = [
    "telefone", "nome", "email", "cpfcnpj", "genero", "estado",
    "cidade", "referencia", "aniversario", "endereco", null,
    "atualizar", "carteira", "whatsapp", "tag", "status",
  ];
  // Larguras das colunas no modelo original.
  const MODEL_WIDTHS = { 0: 24.8, 1: 34.1, 2: 26.4, 8: 13.9, 10: 5.9 };

  // Campos mapeáveis. `match` reconhece o cabeçalho da planilha de origem;
  // `fixed` é o valor fixo padrão quando não há escolha salva.
  const FIELDS = [
    { key: "telefone", out: 0, phone: true, match: /telefone|celular|\bfone\b|\btel\b|whats|contato/i },
    { key: "nome", out: 1, match: /\bnome\b|\bname\b|cliente|raz[aã]o/i },
    { key: "email", out: 2, match: /e-?mail/i },
    { key: "cpfcnpj", out: 3, match: /cpf|cnpj|documento/i },
    { key: "genero", out: 4, match: /g[eê]nero|sexo/i },
    { key: "estado", out: 5, match: /^uf$|estado/i },
    { key: "cidade", out: 6, match: /cidade|munic[ií]pio/i },
    { key: "referencia", out: 7, match: /refer[eê]ncia/i },
    { key: "aniversario", out: 8, match: /anivers|nascimento/i },
    { key: "endereco", out: 9, match: /endere[cç]o|logradouro/i },
    { key: "atualizar", out: 11, fixed: "1" },
    { key: "carteira", out: 12, match: /carteira/i },
    { key: "whatsapp", out: 13, fixed: "56" },
    { key: "tag", out: 14, match: /^tags?$/i },
    { key: "status", out: 15, match: /status|situa[cç][aã]o/i },
  ];

  const PREVIEW_LIMIT = 200;
  const PREFS_KEY = "corretor-planilhas:prefs:v2";

  const $ = (id) => document.getElementById(id);
  const els = {
    app: $("app"),
    dropzone: $("dropzone"),
    fileInput: $("file-input"),
    dropOverlay: $("drop-overlay"),
    modelColsPreview: $("model-cols-preview"),

    fileName: $("file-name"),
    fileExtra: $("file-extra"),
    btnSwap: $("btn-swap"),
    sheetField: $("sheet-field"),
    sheetSelect: $("sheet-select"),
    hasHeader: $("has-header"),

    mapList: $("map-list"),
    mapCount: $("map-count"),

    rules: $("rules"),
    rulesSummary: $("rules-summary"),
    countryCode: $("country-code"),
    minDigits: $("min-digits"),
    smartPrefix: $("smart-prefix"),
    removeDuplicates: $("remove-duplicates"),
    filterInvalid: $("filter-invalid"),

    statTotal: $("stat-total"),
    statClean: $("stat-clean"),
    statDup: $("stat-dup"),
    statInvalid: $("stat-invalid"),

    btnExport: $("btn-export"),
    btnExportDisparo: $("btn-export-disparo"),

    tabOk: $("tab-ok"),
    tabOut: $("tab-out"),
    tabOkCount: $("tab-ok-count"),
    tabOutCount: $("tab-out-count"),
    showEmpty: $("show-empty"),
    emptyColsToggle: $("empty-cols-toggle"),
    previewWrap: $("preview-wrap"),
    previewTable: $("preview-table"),
    tableEmpty: $("table-empty"),
    tableFoot: $("table-foot"),

    toast: $("toast"),
  };

  const state = {
    workbook: null,
    fileName: "",
    fileSize: 0,
    sheetData: [], // array de arrays da aba atual
    headers: [],   // rótulos das colunas de origem
    map: {},       // key -> { type: "col" | "fixed" | "empty", index?, value? }
    processed: null,
    tab: "ok",
  };

  // ---------- Preferências (lembradas entre usos) ----------

  function loadPrefs() {
    try {
      return JSON.parse(localStorage.getItem(PREFS_KEY)) || {};
    } catch (e) {
      return {};
    }
  }

  function savePrefs() {
    const map = {};
    FIELDS.forEach((f) => {
      const m = state.map[f.key];
      if (!m) return;
      if (m.type === "col") map[f.key] = { type: "col", index: m.index, header: state.headers[m.index] };
      else if (m.type === "fixed") map[f.key] = { type: "fixed", value: m.value };
      else map[f.key] = { type: "empty" };
    });
    const prefs = {
      hasHeader: els.hasHeader.checked,
      countryCode: els.countryCode.value,
      minDigits: els.minDigits.value,
      smartPrefix: els.smartPrefix.checked,
      removeDuplicates: els.removeDuplicates.checked,
      filterInvalid: els.filterInvalid.checked,
      showEmpty: els.showEmpty.checked,
      map,
    };
    try {
      localStorage.setItem(PREFS_KEY, JSON.stringify(prefs));
    } catch (e) {
      /* navegação privada ou armazenamento bloqueado: segue sem lembrar */
    }
  }

  function applyPrefsToControls(prefs) {
    if (typeof prefs.hasHeader === "boolean") els.hasHeader.checked = prefs.hasHeader;
    if (prefs.countryCode) els.countryCode.value = prefs.countryCode;
    if (prefs.minDigits) els.minDigits.value = prefs.minDigits;
    ["smartPrefix", "removeDuplicates", "filterInvalid", "showEmpty"].forEach((k) => {
      if (typeof prefs[k] === "boolean") els[k].checked = prefs[k];
    });
  }

  // ---------- Utilidades ----------

  function showToast(message, type) {
    els.toast.textContent = message;
    els.toast.className = "toast " + (type || "");
    clearTimeout(showToast._t);
    showToast._t = setTimeout(() => els.toast.classList.add("hidden"), 4000);
  }

  /** Cede a vez ao navegador para pintar estados de carregamento antes de
   * trabalho bloqueante (leitura/escrita do XLSX). O timeout cobre abas em
   * segundo plano, onde o requestAnimationFrame fica pausado. */
  function uiYield() {
    return new Promise((resolve) => {
      const done = () => { clearTimeout(t); setTimeout(resolve, 0); };
      const t = setTimeout(done, 60);
      requestAnimationFrame(done);
    });
  }

  function setButtonLoading(btn, loading, loadingText) {
    const label = btn.querySelector(".btn-label");
    if (loading) {
      btn.dataset.originalLabel = label.textContent;
      label.textContent = loadingText || "Gerando...";
      btn.classList.add("is-loading");
      btn.disabled = true;
    } else {
      if (btn.dataset.originalLabel) label.textContent = btn.dataset.originalLabel;
      btn.classList.remove("is-loading");
      btn.disabled = !state.processed || state.processed.rows.length === 0;
    }
  }

  const fmt = (n) => n.toLocaleString("pt-BR");

  function formatBytes(n) {
    if (n < 1024) return n + " B";
    if (n < 1024 * 1024) return (n / 1024).toFixed(1).replace(".", ",") + " KB";
    return (n / 1024 / 1024).toFixed(2).replace(".", ",") + " MB";
  }

  function colLetter(index) {
    // 0 -> A, 25 -> Z, 26 -> AA
    let s = "";
    let n = index;
    while (n >= 0) {
      s = String.fromCharCode(65 + (n % 26)) + s;
      n = Math.floor(n / 26) - 1;
    }
    return s;
  }

  const isBlank = (v) => v === "" || v === null || v === undefined;
  const norm = (s) => String(s || "").trim().toLowerCase();

  function displayValue(v) {
    if (v instanceof Date) return v.toLocaleDateString("pt-BR");
    if (isBlank(v)) return "";
    return String(v);
  }

  /** Limpa o telefone: só dígitos e aplica a regra do DDI. */
  function cleanPhone(raw, opts) {
    if (isBlank(raw)) return "";
    // Remove tudo que não é dígito (espaços, parênteses, hífens, pontos, "+", letras...)
    let digits = String(raw).replace(/\D+/g, "");
    if (!digits) return "";

    const cc = opts.countryCode;
    if (opts.smartPrefix) {
      // Só adiciona o DDI se o número ainda não estiver "completo" com ele.
      // Para o Brasil (55): prefixado tem 12 (fixo) ou 13 (celular) dígitos.
      const alreadyPrefixed =
        digits.startsWith(cc) &&
        (digits.length === cc.length + 10 || digits.length === cc.length + 11);
      if (!alreadyPrefixed) digits = cc + digits;
    } else {
      digits = cc + digits;
    }
    return digits;
  }

  function getOptions() {
    const cc = (els.countryCode.value || "").replace(/\D+/g, "") || "55";
    return {
      hasHeader: els.hasHeader.checked,
      countryCode: cc,
      smartPrefix: els.smartPrefix.checked,
      removeDuplicates: els.removeDuplicates.checked,
      filterInvalid: els.filterInvalid.checked,
      minDigits: Math.max(6, parseInt(els.minDigits.value, 10) || 10),
    };
  }

  // ---------- Arquivo ----------

  function setupFileInputs() {
    els.dropzone.addEventListener("click", () => els.fileInput.click());
    els.btnSwap.addEventListener("click", () => els.fileInput.click());
    els.fileInput.addEventListener("change", (e) => {
      const file = e.target.files && e.target.files[0];
      if (file) handleFile(file);
      els.fileInput.value = "";
    });

    // Soltar arquivo em qualquer lugar da página (também para trocar a planilha).
    let depth = 0;
    const hasFiles = (e) => e.dataTransfer && Array.from(e.dataTransfer.types || []).includes("Files");
    window.addEventListener("dragenter", (e) => {
      if (!hasFiles(e)) return;
      e.preventDefault();
      depth++;
      els.app.classList.add("is-dragging");
    });
    window.addEventListener("dragover", (e) => {
      if (hasFiles(e)) e.preventDefault();
    });
    window.addEventListener("dragleave", (e) => {
      if (!hasFiles(e)) return;
      depth = Math.max(0, depth - 1);
      if (depth === 0) els.app.classList.remove("is-dragging");
    });
    window.addEventListener("drop", (e) => {
      if (!hasFiles(e)) return;
      e.preventDefault();
      depth = 0;
      els.app.classList.remove("is-dragging");
      const file = e.dataTransfer.files && e.dataTransfer.files[0];
      if (file) handleFile(file);
    });
  }

  function handleFile(file) {
    if (!/\.(xlsx|xls|csv)$/i.test(file.name)) {
      showToast("Arquivo não suportado. Use .xlsx, .xls ou .csv.", "error");
      return;
    }
    if (typeof XLSX === "undefined") {
      showToast("A biblioteca de planilhas não carregou. Verifique a conexão e recarregue a página.", "error");
      return;
    }
    els.app.classList.add("is-reading");
    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        await uiYield();
        const wb = XLSX.read(new Uint8Array(e.target.result), { type: "array", cellDates: true });
        state.workbook = wb;
        state.fileName = file.name;
        state.fileSize = file.size;
        els.fileName.textContent = file.name;
        populateSheetSelector();
        els.app.dataset.state = "ready";
      } catch (err) {
        console.error(err);
        showToast("Não foi possível ler a planilha: " + err.message, "error");
      } finally {
        els.app.classList.remove("is-reading");
      }
    };
    reader.onerror = () => {
      els.app.classList.remove("is-reading");
      showToast("Erro ao ler o arquivo.", "error");
    };
    reader.readAsArrayBuffer(file);
  }

  // ---------- Aba e colunas de origem ----------

  function populateSheetSelector() {
    els.sheetSelect.innerHTML = "";
    state.workbook.SheetNames.forEach((name) => {
      els.sheetSelect.appendChild(new Option(name, name));
    });
    els.sheetField.classList.toggle("hidden", state.workbook.SheetNames.length < 2);
    loadSelectedSheet();
  }

  function loadSelectedSheet() {
    const ws = state.workbook.Sheets[els.sheetSelect.value];
    state.sheetData = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: true });
    rebuildSource();
  }

  /** Recalcula cabeçalhos de origem, aplica o mapeamento lembrado e processa. */
  function rebuildSource() {
    const hasHeader = els.hasHeader.checked;
    const colCount = state.sheetData.reduce((m, r) => Math.max(m, r.length), 0);
    const first = state.sheetData[0] || [];

    state.headers = [];
    for (let i = 0; i < colCount; i++) {
      const cell = first[i];
      state.headers.push(hasHeader && !isBlank(cell) ? String(cell).trim() : `Coluna ${colLetter(i)}`);
    }

    const rows = Math.max(0, state.sheetData.length - (hasHeader ? 1 : 0));
    els.fileExtra.textContent = `${fmt(rows)} linha${rows === 1 ? "" : "s"} · ${formatBytes(state.fileSize)}`;

    resolveMapping();
    renderMapList();
    processNow();
  }

  function sampleOf(index) {
    const start = els.hasHeader.checked ? 1 : 0;
    for (let r = start; r < Math.min(state.sheetData.length, start + 20); r++) {
      const v = state.sheetData[r][index];
      if (!isBlank(v)) return displayValue(v);
    }
    return "";
  }

  /** Coluna que mais parece telefone (10 a 13 dígitos nas primeiras linhas). */
  function guessPhoneColumn() {
    const start = els.hasHeader.checked ? 1 : 0;
    const sample = state.sheetData.slice(start, start + 30);
    let best = -1;
    let bestScore = 0;
    for (let c = 0; c < state.headers.length; c++) {
      let score = 0;
      for (const row of sample) {
        const v = row[c];
        if (isBlank(v) || v instanceof Date) continue;
        const d = String(v).replace(/\D+/g, "").length;
        if (d >= 10 && d <= 13) score += 2;
        else if (d >= 8) score += 1;
      }
      if (score > bestScore) {
        bestScore = score;
        best = c;
      }
    }
    return best;
  }

  function findHeader(re) {
    if (!els.hasHeader.checked) return -1;
    return state.headers.findIndex((h) => re.test(h));
  }

  function resolveMapping() {
    const saved = loadPrefs().map || {};
    const colCount = state.headers.length;
    const hasHeader = els.hasHeader.checked;
    const next = {};

    FIELDS.forEach((f) => {
      const s = saved[f.key];
      let m = null;

      if (s && s.type === "fixed" && !f.phone) m = { type: "fixed", value: s.value };
      else if (s && s.type === "empty" && !f.phone) m = { type: "empty" };
      else if (s && s.type === "col") {
        // Mesma coluna do último uso: pelo nome do cabeçalho; sem cabeçalho, pela posição.
        let idx = -1;
        if (hasHeader && s.header && !/^Coluna [A-Z]+$/.test(s.header)) {
          idx = state.headers.findIndex((h) => norm(h) === norm(s.header));
        } else if (typeof s.index === "number" && s.index < colCount) {
          idx = s.index;
        }
        if (idx >= 0) m = { type: "col", index: idx };
      }

      if (!m) {
        let idx = f.match ? findHeader(f.match) : -1;
        if (f.phone && idx < 0) idx = guessPhoneColumn();
        if (f.phone && idx < 0 && colCount) idx = 0;
        if (idx >= 0) m = { type: "col", index: idx };
        else if (f.fixed !== undefined) m = { type: "fixed", value: f.fixed };
        else m = { type: "empty" };
      }
      next[f.key] = m;
    });
    state.map = next;
  }

  // ---------- Mapeamento (UI) ----------

  function renderModelStrip() {
    MODEL_HEADERS.forEach((h, i) => {
      if (h === null) return;
      const li = document.createElement("li");
      li.innerHTML = `<span class="col-letter">${colLetter(i)}</span>`;
      li.appendChild(document.createTextNode(h));
      els.modelColsPreview.appendChild(li);
    });
  }

  function renderMapList() {
    els.mapList.innerHTML = "";
    FIELDS.forEach((f) => {
      const m = state.map[f.key];
      const id = "map-" + f.key;

      const row = document.createElement("div");
      row.className = "map-row";
      row.dataset.type = m.type;

      const letter = document.createElement("span");
      letter.className = "col-letter";
      letter.textContent = colLetter(f.out);
      letter.setAttribute("aria-hidden", "true");

      const label = document.createElement("label");
      label.htmlFor = id;
      label.className = "map-label";
      label.textContent = f.key;

      const select = document.createElement("select");
      select.id = id;
      if (!f.phone) {
        select.appendChild(new Option("vazio", "empty"));
        select.appendChild(new Option("valor fixo…", "fixed"));
      }
      const group = document.createElement("optgroup");
      group.label = "Colunas da planilha";
      state.headers.forEach((h, i) => {
        const sample = sampleOf(i);
        const text = `${colLetter(i)} · ${h}${sample ? "  (" + sample.slice(0, 28) + ")" : ""}`;
        group.appendChild(new Option(text, "col:" + i));
      });
      select.appendChild(group);
      select.value = m.type === "col" ? "col:" + m.index : m.type;

      const fixed = document.createElement("input");
      fixed.type = "text";
      fixed.className = "map-fixed";
      fixed.placeholder = "valor";
      fixed.setAttribute("aria-label", `Valor fixo de ${f.key}`);
      fixed.value = m.type === "fixed" ? m.value : "";

      select.addEventListener("change", () => {
        const v = select.value;
        if (v.startsWith("col:")) state.map[f.key] = { type: "col", index: parseInt(v.slice(4), 10) };
        else if (v === "fixed") state.map[f.key] = { type: "fixed", value: f.fixed || "" };
        else state.map[f.key] = { type: "empty" };
        row.dataset.type = state.map[f.key].type;
        fixed.value = state.map[f.key].value || "";
        if (v === "fixed") fixed.focus();
        onMappingChange();
      });
      fixed.addEventListener("input", () => {
        state.map[f.key] = { type: "fixed", value: fixed.value };
        onMappingChange();
      });

      row.append(letter, label, select, fixed);
      if (f.phone) {
        const tag = document.createElement("span");
        tag.className = "map-tag";
        tag.textContent = "tratado";
        row.appendChild(tag);
      }
      els.mapList.appendChild(row);
    });
    updateMapCount();
  }

  function updateMapCount() {
    const filled = FIELDS.filter((f) => state.map[f.key].type !== "empty").length;
    els.mapCount.textContent = `${filled} de ${FIELDS.length} preenchidas`;
  }

  function onMappingChange() {
    updateMapCount();
    savePrefs();
    scheduleProcess();
  }

  function updateRulesSummary() {
    const o = getOptions();
    const parts = [`DDI ${o.countryCode}`];
    if (o.filterInvalid) parts.push(`mín. ${o.minDigits} dígitos`);
    if (o.removeDuplicates) parts.push("sem repetidos");
    els.rulesSummary.textContent = parts.join(" · ");
  }

  // ---------- Processamento ----------

  let processTimer = 0;
  function scheduleProcess() {
    clearTimeout(processTimer);
    els.app.classList.add("is-busy");
    processTimer = setTimeout(processNow, 140);
  }

  function processNow() {
    clearTimeout(processTimer);
    if (!state.workbook) return;
    const t0 = performance.now();
    const opts = getOptions();
    const phone = state.map.telefone;
    const phoneIdx = phone && phone.type === "col" ? phone.index : -1;
    const start = opts.hasHeader ? 1 : 0;
    const data = state.sheetData;
    const minTotalLen = opts.minDigits + opts.countryCode.length;

    const seen = new Set();
    const rows = [];      // { phone, src }
    const discarded = []; // { line, raw, reason, src }
    let total = 0;
    let duplicates = 0;
    let invalids = 0;

    for (let r = start; r < data.length; r++) {
      const src = data[r];
      // Ignora linhas totalmente vazias (laço manual é mais rápido que .every()).
      let nonEmpty = false;
      for (let c = 0; c < src.length; c++) {
        if (!isBlank(src[c])) { nonEmpty = true; break; }
      }
      if (!nonEmpty) continue;
      total++;

      const raw = phoneIdx >= 0 ? src[phoneIdx] : "";
      const cleaned = cleanPhone(raw, opts);
      if (opts.filterInvalid ? cleaned.length < minTotalLen : cleaned.length === 0) {
        invalids++;
        discarded.push({ line: r + 1, raw, reason: isBlank(raw) ? "sem número" : "curto demais", src });
        continue;
      }
      if (opts.removeDuplicates) {
        if (seen.has(cleaned)) {
          duplicates++;
          discarded.push({ line: r + 1, raw, reason: "repetido", src });
          continue;
        }
        seen.add(cleaned);
      }
      rows.push({ phone: cleaned, src });
    }

    state.processed = { rows, discarded, stats: { total, clean: rows.length, duplicates, invalids } };
    console.info(`[bench] process: ${(performance.now() - t0).toFixed(1)} ms para ${total} linhas`);
    els.app.classList.remove("is-busy");
    updateRulesSummary();
    renderResult();
  }

  /** Valor final de um campo do modelo para uma linha tratada. */
  function fieldValue(f, item) {
    if (f.phone) return item.phone;
    const m = state.map[f.key];
    if (m.type === "col") {
      const v = item.src[m.index];
      return isBlank(v) ? null : v;
    }
    if (m.type === "fixed") {
      const raw = String(m.value || "").trim();
      if (raw === "") return null;
      const n = Number(raw);
      return Number.isNaN(n) ? raw : n;
    }
    return null;
  }

  // ---------- Resultado ----------

  function renderResult() {
    const { stats, rows, discarded } = state.processed;
    els.statTotal.textContent = fmt(stats.total);
    els.statClean.textContent = fmt(stats.clean);
    els.statDup.textContent = fmt(stats.duplicates);
    els.statInvalid.textContent = fmt(stats.invalids);
    els.tabOkCount.textContent = fmt(rows.length);
    els.tabOutCount.textContent = fmt(discarded.length);

    const none = rows.length === 0;
    els.btnExportDisparo.disabled = none;
    els.btnExport.disabled = none;

    if (state.tab === "ok") renderFinalTable();
    else renderDiscardedTable();
  }

  function buildTable(columns, items, cellFn) {
    const tbl = els.previewTable;
    tbl.innerHTML = "";
    const thead = tbl.createTHead();
    const trh = thead.insertRow();
    columns.forEach((c) => {
      const th = document.createElement("th");
      th.scope = "col";
      if (c.letter) {
        const l = document.createElement("span");
        l.className = "col-letter";
        l.textContent = c.letter;
        th.appendChild(l);
      }
      th.appendChild(document.createTextNode(c.title));
      if (c.cls) th.className = c.cls;
      trh.appendChild(th);
    });
    const tbody = tbl.createTBody();
    const frag = document.createDocumentFragment();
    items.slice(0, PREVIEW_LIMIT).forEach((item) => {
      const tr = document.createElement("tr");
      columns.forEach((c, i) => {
        const td = document.createElement("td");
        const out = cellFn(item, c, i);
        td.textContent = out.text;
        if (out.cls) td.className = out.cls;
        tr.appendChild(td);
      });
      frag.appendChild(tr);
    });
    tbody.appendChild(frag);
  }

  function renderFinalTable() {
    const { rows } = state.processed;
    const showEmpty = els.showEmpty.checked;
    const fields = FIELDS.filter((f) => showEmpty || state.map[f.key].type !== "empty");
    const hidden = FIELDS.length - fields.length;
    const columns = fields.map((f) => ({
      f,
      letter: colLetter(f.out),
      title: f.key,
      cls: state.map[f.key].type === "fixed" ? "is-fixed" : "",
    }));

    buildTable(columns, rows, (item, c) => {
      const v = fieldValue(c.f, item);
      let cls = "";
      if (c.f.phone) cls = "cell-phone";
      else if (state.map[c.f.key].type === "fixed") cls = "cell-fixed";
      return { text: displayValue(v), cls };
    });

    els.emptyColsToggle.classList.remove("hidden");
    setEmptyMessage(rows.length ? "" : "Nenhum contato passou na limpeza. Confira a coluna do telefone e as regras de limpeza.");
    els.tableFoot.textContent = rows.length
      ? footText(rows.length) + (hidden && !showEmpty ? ` · ${hidden} coluna${hidden === 1 ? "" : "s"} vazia${hidden === 1 ? "" : "s"} oculta${hidden === 1 ? "" : "s"}` : "")
      : "";
  }

  function renderDiscardedTable() {
    const { discarded } = state.processed;
    const nome = state.map.nome;
    const columns = [
      { title: "linha", cls: "num" },
      { title: "telefone na planilha" },
      { title: "motivo" },
    ];
    if (nome && nome.type === "col") columns.push({ title: "nome" });

    buildTable(columns, discarded, (item, c, i) => {
      if (i === 0) return { text: String(item.line), cls: "num" };
      if (i === 1) return { text: displayValue(item.raw) || "(vazio)", cls: "cell-raw" };
      if (i === 2) return { text: item.reason, cls: item.reason === "repetido" ? "cell-warn" : "cell-err" };
      return { text: displayValue(item.src[nome.index]) };
    });

    els.emptyColsToggle.classList.add("hidden");
    setEmptyMessage(discarded.length ? "" : "Nenhuma linha foi descartada.");
    els.tableFoot.textContent = discarded.length ? footText(discarded.length) : "";
  }

  function footText(n) {
    return n > PREVIEW_LIMIT ? `Mostrando ${fmt(PREVIEW_LIMIT)} de ${fmt(n)} linhas` : `${fmt(n)} linha${n === 1 ? "" : "s"}`;
  }

  function setEmptyMessage(msg) {
    els.tableEmpty.textContent = msg;
    els.tableEmpty.classList.toggle("hidden", !msg);
    els.previewTable.classList.toggle("hidden", !!msg);
  }

  function selectTab(tab) {
    state.tab = tab;
    els.tabOk.setAttribute("aria-selected", String(tab === "ok"));
    els.tabOut.setAttribute("aria-selected", String(tab === "out"));
    if (state.processed) renderResult();
  }

  // ---------- Exportação ----------

  function baseName() {
    return state.fileName.replace(/\.(xlsx|xls|csv)$/i, "");
  }

  async function exportDisparo() {
    if (!state.processed || !state.processed.rows.length) return;
    setButtonLoading(els.btnExportDisparo, true, "Gerando arquivo...");
    await uiYield();
    try {
      const t0 = performance.now();
      const { rows } = state.processed;
      const aoa = [MODEL_HEADERS.slice()];
      for (const item of rows) {
        const out = new Array(MODEL_HEADERS.length).fill(null);
        FIELDS.forEach((f) => { out[f.out] = fieldValue(f, item); });
        aoa.push(out);
      }

      const ws = XLSX.utils.aoa_to_sheet(aoa);
      // "telefone" gravado como TEXTO, para o Excel não cortar o 55
      // nem converter para notação científica.
      for (let r = 1; r < aoa.length; r++) {
        const ref = XLSX.utils.encode_cell({ r, c: 0 });
        if (ws[ref]) {
          ws[ref].t = "s";
          ws[ref].v = String(ws[ref].v);
        }
      }
      ws["!cols"] = MODEL_HEADERS.map((h, i) => ({ wch: MODEL_WIDTHS[i] || 12.7 }));

      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Página1");
      const outName = `${baseName()}_disparo.xlsx`;
      XLSX.writeFile(wb, outName);

      console.info(`[bench] exportDisparo: ${(performance.now() - t0).toFixed(1)} ms para ${rows.length} linhas`);
      showToast(`Arquivo salvo: ${outName}`, "success");
    } catch (err) {
      console.error(err);
      showToast("Erro ao gerar o Modelo de Disparo: " + err.message, "error");
    } finally {
      setButtonLoading(els.btnExportDisparo, false);
    }
  }

  async function exportOriginal() {
    if (!state.processed || !state.processed.rows.length) return;
    setButtonLoading(els.btnExport, true, "Gerando...");
    await uiYield();
    try {
      const { rows } = state.processed;
      const headers = ["whatsapp", ...state.headers];
      const aoa = [headers];
      for (const item of rows) aoa.push([item.phone, ...item.src]);

      const ws = XLSX.utils.aoa_to_sheet(aoa);
      for (let r = 1; r < aoa.length; r++) {
        const ref = XLSX.utils.encode_cell({ r, c: 0 });
        if (ws[ref]) {
          ws[ref].t = "s";
          ws[ref].v = String(ws[ref].v);
        }
      }
      ws["!cols"] = headers.map((h, i) => ({ wch: i === 0 ? 16 : Math.min(40, Math.max(12, String(h).length + 4)) }));

      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "whatsapp");
      const outName = `${baseName()}_whatsapp.xlsx`;
      XLSX.writeFile(wb, outName);
      showToast(`Arquivo salvo: ${outName}`, "success");
    } catch (err) {
      console.error(err);
      showToast("Erro ao exportar: " + err.message, "error");
    } finally {
      setButtonLoading(els.btnExport, false);
    }
  }

  // ---------- Inicialização ----------

  function init() {
    if (typeof XLSX === "undefined") {
      showToast("A biblioteca de planilhas não carregou do CDN. Verifique a conexão.", "error");
    }
    applyPrefsToControls(loadPrefs());
    renderModelStrip();
    updateRulesSummary();
    setupFileInputs();

    els.sheetSelect.addEventListener("change", loadSelectedSheet);
    els.hasHeader.addEventListener("change", () => {
      savePrefs();
      rebuildSource();
    });

    [els.countryCode, els.minDigits].forEach((el) =>
      el.addEventListener("input", () => {
        updateRulesSummary();
        savePrefs();
        scheduleProcess();
      })
    );
    [els.smartPrefix, els.removeDuplicates, els.filterInvalid].forEach((el) =>
      el.addEventListener("change", () => {
        updateRulesSummary();
        savePrefs();
        scheduleProcess();
      })
    );
    els.showEmpty.addEventListener("change", () => {
      savePrefs();
      if (state.processed) renderResult();
    });

    els.tabOk.addEventListener("click", () => selectTab("ok"));
    els.tabOut.addEventListener("click", () => selectTab("out"));
    els.btnExportDisparo.addEventListener("click", exportDisparo);
    els.btnExport.addEventListener("click", exportOriginal);
  }

  init();
})();
