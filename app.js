/* LINK10 · Tratador de Planilhas para WhatsApp
 * Roda 100% no navegador. Usa SheetJS (xlsx) carregado via CDN.
 */
(function () {
  "use strict";

  const els = {
    dropzone: document.getElementById("dropzone"),
    fileInput: document.getElementById("file-input"),
    fileInfo: document.getElementById("file-info"),
    fileName: document.getElementById("file-name"),
    fileExtra: document.getElementById("file-extra"),
    btnReset: document.getElementById("btn-reset"),

    stepConfig: document.getElementById("step-config"),
    stepResult: document.getElementById("step-result"),

    sheetSelect: document.getElementById("sheet-select"),
    hasHeader: document.getElementById("has-header"),
    phoneColumn: document.getElementById("phone-column"),
    phoneSample: document.getElementById("phone-sample"),
    countryCode: document.getElementById("country-code"),
    smartPrefix: document.getElementById("smart-prefix"),
    removeDuplicates: document.getElementById("remove-duplicates"),
    filterInvalid: document.getElementById("filter-invalid"),
    minDigits: document.getElementById("min-digits"),

    btnProcess: document.getElementById("btn-process"),
    btnBack: document.getElementById("btn-back"),
    btnExport: document.getElementById("btn-export"),

    statTotal: document.getElementById("stat-total"),
    statClean: document.getElementById("stat-clean"),
    statDup: document.getElementById("stat-dup"),
    statInvalid: document.getElementById("stat-invalid"),

    previewTable: document.getElementById("preview-table"),
    toast: document.getElementById("toast"),
  };

  const state = {
    workbook: null,
    fileName: "",
    sheetData: [], // array of arrays for current sheet
    headers: [],   // computed column headers
    processed: null, // {rows, stats, headers}
  };

  // ---------- Utilities ----------

  function showToast(message, type) {
    els.toast.textContent = message;
    els.toast.className = "toast " + (type || "");
    setTimeout(() => els.toast.classList.add("hidden"), 50);
    requestAnimationFrame(() => {
      els.toast.classList.remove("hidden");
    });
    clearTimeout(showToast._t);
    showToast._t = setTimeout(() => els.toast.classList.add("hidden"), 4000);
  }

  /** Yields to the browser so that pending UI updates (e.g. spinners) can paint
   * before we run blocking work like XLSX parsing. */
  function uiYield() {
    return new Promise((resolve) => {
      if (typeof requestAnimationFrame === "function") {
        requestAnimationFrame(() => setTimeout(resolve, 0));
      } else {
        setTimeout(resolve, 0);
      }
    });
  }

  function setButtonLoading(btn, loading, loadingText) {
    if (!btn) return;
    if (loading) {
      btn.dataset.originalLabel = btn.querySelector(".btn-label").textContent;
      btn.querySelector(".btn-label").textContent = loadingText || "Processando...";
      btn.classList.add("is-loading");
      btn.disabled = true;
    } else {
      if (btn.dataset.originalLabel) {
        btn.querySelector(".btn-label").textContent = btn.dataset.originalLabel;
      }
      btn.classList.remove("is-loading");
      btn.disabled = false;
    }
  }

  function formatBytes(n) {
    if (n < 1024) return n + " B";
    if (n < 1024 * 1024) return (n / 1024).toFixed(1) + " KB";
    return (n / 1024 / 1024).toFixed(2) + " MB";
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

  /** Cleans a phone value: keeps only digits, then applies country-code logic. */
  function cleanPhone(raw, opts) {
    if (raw === null || raw === undefined) return "";
    // Strip any non-digit (handles spaces, parens, dashes, dots, "+", letters, etc.)
    let digits = String(raw).replace(/\D+/g, "");
    if (!digits) return "";

    const cc = opts.countryCode;
    if (opts.smartPrefix) {
      // Only add CC if not already present at start of a "fully prefixed" number.
      // For Brazil (cc=55): valid prefixed lengths are 12 (landline) or 13 (mobile).
      const alreadyPrefixed =
        digits.startsWith(cc) &&
        (digits.length === cc.length + 10 || digits.length === cc.length + 11);
      if (!alreadyPrefixed) {
        digits = cc + digits;
      }
    } else {
      digits = cc + digits;
    }
    return digits;
  }

  function isValidPhone(digits, opts) {
    if (!opts.filterInvalid) return digits.length > 0;
    return digits.length >= opts.minDigits + opts.countryCode.length;
  }

  // ---------- File handling ----------

  function setupDropzone() {
    els.dropzone.addEventListener("click", () => els.fileInput.click());
    els.dropzone.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        els.fileInput.click();
      }
    });
    ["dragenter", "dragover"].forEach((evt) =>
      els.dropzone.addEventListener(evt, (e) => {
        e.preventDefault();
        e.stopPropagation();
        els.dropzone.classList.add("dragover");
      })
    );
    ["dragleave", "drop"].forEach((evt) =>
      els.dropzone.addEventListener(evt, (e) => {
        e.preventDefault();
        e.stopPropagation();
        els.dropzone.classList.remove("dragover");
      })
    );
    els.dropzone.addEventListener("drop", (e) => {
      const file = e.dataTransfer.files && e.dataTransfer.files[0];
      if (file) handleFile(file);
    });
    els.fileInput.addEventListener("change", (e) => {
      const file = e.target.files && e.target.files[0];
      if (file) handleFile(file);
    });
    els.btnReset.addEventListener("click", reset);
  }

  function reset() {
    state.workbook = null;
    state.fileName = "";
    state.sheetData = [];
    state.headers = [];
    state.processed = null;
    els.fileInput.value = "";
    els.fileInfo.classList.add("hidden");
    els.stepConfig.classList.add("hidden");
    els.stepResult.classList.add("hidden");
  }

  function handleFile(file) {
    if (!/\.(xlsx|xls|csv)$/i.test(file.name)) {
      showToast("Arquivo inválido. Use .xlsx, .xls ou .csv.", "error");
      return;
    }
    if (typeof XLSX === "undefined") {
      showToast("Biblioteca XLSX não carregou (sem conexão?). Recarregue a página.", "error");
      return;
    }
    state.fileName = file.name;
    els.fileName.textContent = file.name;
    els.fileExtra.textContent = formatBytes(file.size) + " · lendo...";
    els.fileInfo.classList.remove("hidden");

    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        await uiYield();
        const t0 = performance.now();
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: "array", cellDates: true });
        const readMs = performance.now() - t0;
        state.workbook = wb;
        populateSheetSelector();
        const totalRows = state.sheetData.length;
        els.fileExtra.textContent =
          formatBytes(file.size) +
          " · " +
          totalRows.toLocaleString("pt-BR") +
          " linha" + (totalRows === 1 ? "" : "s") +
          " · lido em " + readMs.toFixed(0) + " ms";
        els.stepConfig.classList.remove("hidden");
        els.stepResult.classList.add("hidden");
        showToast("Planilha carregada: " + file.name, "success");
      } catch (err) {
        console.error(err);
        els.fileExtra.textContent = formatBytes(file.size);
        showToast("Falha ao ler a planilha: " + err.message, "error");
      }
    };
    reader.onerror = () => {
      els.fileExtra.textContent = formatBytes(file.size);
      showToast("Erro ao ler o arquivo.", "error");
    };
    reader.readAsArrayBuffer(file);
  }

  // ---------- Sheet/column selection ----------

  function populateSheetSelector() {
    els.sheetSelect.innerHTML = "";
    state.workbook.SheetNames.forEach((name) => {
      const opt = document.createElement("option");
      opt.value = name;
      opt.textContent = name;
      els.sheetSelect.appendChild(opt);
    });
    els.sheetSelect.onchange = loadSelectedSheet;
    loadSelectedSheet();
  }

  function loadSelectedSheet() {
    const sheetName = els.sheetSelect.value;
    const ws = state.workbook.Sheets[sheetName];
    // header: 1 -> array of arrays, defval keeps cells aligned
    state.sheetData = XLSX.utils.sheet_to_json(ws, {
      header: 1,
      defval: "",
      raw: true,
    });
    refreshColumnSelector();
  }

  function refreshColumnSelector() {
    if (!state.sheetData.length) {
      els.phoneColumn.innerHTML = "<option>Planilha vazia</option>";
      return;
    }
    const hasHeader = els.hasHeader.checked;
    const firstRow = state.sheetData[0];
    const colCount = firstRow.length;

    state.headers = [];
    for (let i = 0; i < colCount; i++) {
      let label;
      if (hasHeader) {
        const cell = firstRow[i];
        label = cell !== "" && cell !== null && cell !== undefined
          ? String(cell)
          : `Coluna ${colLetter(i)}`;
      } else {
        label = `Coluna ${colLetter(i)}`;
      }
      state.headers.push(label);
    }

    els.phoneColumn.innerHTML = "";
    state.headers.forEach((h, i) => {
      const opt = document.createElement("option");
      opt.value = String(i);
      // Build a sample preview from up to 3 rows
      const sampleRowStart = hasHeader ? 1 : 0;
      const samples = state.sheetData
        .slice(sampleRowStart, sampleRowStart + 3)
        .map((r) => r[i])
        .filter((v) => v !== "" && v !== null && v !== undefined)
        .map((v) => String(v));
      const sampleText = samples.length ? " — ex: " + samples[0] : "";
      opt.textContent = `${h} (${colLetter(i)})${sampleText}`;
      els.phoneColumn.appendChild(opt);
    });

    // Try to auto-detect phone column (longest numeric ratio in first ~20 rows)
    autoDetectPhoneColumn(hasHeader);
    updatePhoneSample();
  }

  function autoDetectPhoneColumn(hasHeader) {
    const start = hasHeader ? 1 : 0;
    const sampleRows = state.sheetData.slice(start, start + 20);
    if (!sampleRows.length) return;

    let bestIdx = 0;
    let bestScore = -1;
    const colCount = state.headers.length;
    for (let c = 0; c < colCount; c++) {
      let score = 0;
      for (const row of sampleRows) {
        const v = row[c];
        if (v === undefined || v === null || v === "") continue;
        const digits = String(v).replace(/\D+/g, "");
        // Score by likely-phone digit length (10-13 digits typical)
        if (digits.length >= 10 && digits.length <= 13) score += 2;
        else if (digits.length >= 8) score += 1;
      }
      if (score > bestScore) {
        bestScore = score;
        bestIdx = c;
      }
    }
    if (bestScore > 0) els.phoneColumn.value = String(bestIdx);
  }

  function updatePhoneSample() {
    const idx = parseInt(els.phoneColumn.value, 10);
    if (isNaN(idx)) {
      els.phoneSample.textContent = "";
      return;
    }
    const hasHeader = els.hasHeader.checked;
    const start = hasHeader ? 1 : 0;
    const samples = state.sheetData
      .slice(start, start + 3)
      .map((r) => r[idx])
      .filter((v) => v !== "" && v !== null && v !== undefined);
    els.phoneSample.textContent = samples.length
      ? "Amostra: " + samples.map((s) => String(s)).join(", ")
      : "";
  }

  // ---------- Processing ----------

  function getOptions() {
    const cc = (els.countryCode.value || "").replace(/\D+/g, "") || "55";
    return {
      sheetName: els.sheetSelect.value,
      hasHeader: els.hasHeader.checked,
      phoneIdx: parseInt(els.phoneColumn.value, 10),
      countryCode: cc,
      smartPrefix: els.smartPrefix.checked,
      removeDuplicates: els.removeDuplicates.checked,
      filterInvalid: els.filterInvalid.checked,
      minDigits: Math.max(6, parseInt(els.minDigits.value, 10) || 10),
    };
  }

  async function processRows() {
    const opts = getOptions();
    if (isNaN(opts.phoneIdx)) {
      showToast("Selecione a coluna do telefone.", "error");
      return;
    }
    setButtonLoading(els.btnProcess, true, "Processando...");
    // Give the browser a chance to paint the spinner before we block the main thread.
    await uiYield();

    try {
      const t0 = performance.now();
      const start = opts.hasHeader ? 1 : 0;
      const dataRows = state.sheetData;
      const totalDataRows = dataRows.length - start;

      const seen = new Set();
      let duplicates = 0;
      let invalids = 0;
      const outRows = [];
      const phoneIdx = opts.phoneIdx;
      const cc = opts.countryCode;
      const smart = opts.smartPrefix;
      const filterInvalid = opts.filterInvalid;
      const minTotalLen = opts.minDigits + cc.length;
      const removeDuplicates = opts.removeDuplicates;

      for (let r = start; r < dataRows.length; r++) {
        const row = dataRows[r];

        // Manual empty-row check (faster than `.every()` for large sheets).
        let nonEmpty = false;
        for (let c = 0; c < row.length; c++) {
          const v = row[c];
          if (v !== "" && v !== null && v !== undefined) { nonEmpty = true; break; }
        }
        if (!nonEmpty) continue;

        const cleaned = cleanPhone(row[phoneIdx], { countryCode: cc, smartPrefix: smart });

        if (filterInvalid ? cleaned.length < minTotalLen : cleaned.length === 0) {
          invalids++;
          continue;
        }

        if (removeDuplicates) {
          if (seen.has(cleaned)) { duplicates++; continue; }
          seen.add(cleaned);
        }

        const outRow = new Array(row.length + 1);
        outRow[0] = cleaned;
        for (let c = 0; c < row.length; c++) outRow[c + 1] = row[c];
        outRows.push(outRow);
      }

      const outHeaders = ["whatsapp", ...state.headers];

      state.processed = {
        headers: outHeaders,
        rows: outRows,
        stats: {
          total: totalDataRows,
          clean: outRows.length,
          duplicates,
          invalids,
        },
      };

      const elapsed = performance.now() - t0;
      console.info(`[bench] processRows: ${elapsed.toFixed(1)} ms for ${totalDataRows} rows`);
      renderResult();
    } catch (err) {
      console.error(err);
      showToast("Erro ao processar: " + err.message, "error");
    } finally {
      setButtonLoading(els.btnProcess, false);
    }
  }

  function renderResult() {
    const { headers, rows, stats } = state.processed;
    els.statTotal.textContent = stats.total;
    els.statClean.textContent = stats.clean;
    els.statDup.textContent = stats.duplicates;
    els.statInvalid.textContent = stats.invalids;

    const opts = getOptions();
    const phoneIdxInOut = opts.phoneIdx + 1; // because we prepended whatsapp col

    // Preview first 10 rows
    const tbl = els.previewTable;
    tbl.innerHTML = "";
    const thead = document.createElement("thead");
    const trh = document.createElement("tr");
    headers.forEach((h) => {
      const th = document.createElement("th");
      th.textContent = h;
      trh.appendChild(th);
    });
    thead.appendChild(trh);
    tbl.appendChild(thead);

    const tbody = document.createElement("tbody");
    rows.slice(0, 10).forEach((row) => {
      const tr = document.createElement("tr");
      row.forEach((val, i) => {
        const td = document.createElement("td");
        if (i === 0) td.className = "whatsapp-cell";
        else if (i === phoneIdxInOut) td.className = "original-cell";
        let display = val;
        if (val instanceof Date) {
          display = val.toLocaleDateString("pt-BR");
        } else if (val === null || val === undefined) {
          display = "";
        }
        td.textContent = String(display);
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
    tbl.appendChild(tbody);

    els.stepResult.classList.remove("hidden");
    els.stepResult.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  // ---------- Export ----------

  async function exportXlsx() {
    if (!state.processed) return;
    setButtonLoading(els.btnExport, true, "Gerando arquivo...");
    await uiYield();

    try {
      const t0 = performance.now();
      const { headers, rows } = state.processed;

      const aoa = [headers, ...rows];
      const ws = XLSX.utils.aoa_to_sheet(aoa);

      // Force the whatsapp column (column A) to be stored as TEXT so the leading
      // "55" is never converted to scientific notation by Excel.
      for (let r = 1; r < aoa.length; r++) {
        const cellRef = XLSX.utils.encode_cell({ r, c: 0 });
        if (ws[cellRef]) {
          ws[cellRef].t = "s";
          ws[cellRef].v = String(ws[cellRef].v);
        }
      }

      ws["!cols"] = headers.map((h, i) => {
        if (i === 0) return { wch: 16 };
        return { wch: Math.min(40, Math.max(12, String(h).length + 4)) };
      });

      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "whatsapp");

      const baseName = state.fileName.replace(/\.(xlsx|xls|csv)$/i, "");
      const outName = `${baseName}_whatsapp.xlsx`;
      XLSX.writeFile(wb, outName);

      const elapsed = performance.now() - t0;
      console.info(`[bench] exportXlsx: ${elapsed.toFixed(1)} ms for ${rows.length} rows`);
      showToast(`Arquivo salvo: ${outName}`, "success");
    } catch (err) {
      console.error(err);
      showToast("Erro ao exportar: " + err.message, "error");
    } finally {
      setButtonLoading(els.btnExport, false);
    }
  }

  // ---------- Wire-up ----------

  function init() {
    if (typeof XLSX === "undefined") {
      showToast(
        "Biblioteca XLSX não pôde ser carregada do CDN. Verifique sua conexão.",
        "error"
      );
    }
    setupDropzone();

    els.hasHeader.addEventListener("change", refreshColumnSelector);
    els.phoneColumn.addEventListener("change", updatePhoneSample);

    els.btnProcess.addEventListener("click", processRows);
    els.btnExport.addEventListener("click", exportXlsx);
    els.btnBack.addEventListener("click", () => {
      els.stepResult.classList.add("hidden");
      els.stepConfig.scrollIntoView({ behavior: "smooth", block: "start" });
    });
  }

  init();
})();
