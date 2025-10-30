(function () {
  "use strict";

  // Datos de ejemplo. Puedes reemplazar/expandir según tu catálogo real.
  /**
   * Estructura por producto
   * {
   *   id: string,
   *   nombre: string,
   *   procesos: Array<{
   *     codigo: string,
   *     analisis: string,
   *     metodo: string,
   *     cMtra_g: number,
   *     cantidad: number,
   *     vrUnit1: number,
   *     vrUnit2: number,
   *     vrUnit3: number,
   *     vrUnitUSD: number
   *   }>
   * }
   */
  let productos = [
    {
      id: "A-ACEITE-EUCALIPTO-USP",
      nombre: "ACEITE DE EUCALIPTO USP",
      procesos: [
        { codigo: "AEU-001", analisis: "Identificación", metodo: "USP <197>", cMtra_g: 5, cantidad: 1, vrUnit1: 190000, vrUnit2: 180000, vrUnit3: 165000, vrUnitUSD: 190000 },
        { codigo: "AEU-002", analisis: "Índice de refracción", metodo: "USP <831>", cMtra_g: 3, cantidad: 1, vrUnit1: 240000, vrUnit2: 225000, vrUnit3: 210000, vrUnitUSD: 240000 },
        { codigo: "AEU-003", analisis: "Cromatografía GC", metodo: "USP <621>", cMtra_g: 2, cantidad: 1, vrUnit1: 880000, vrUnit2: 850000, vrUnit3: 820000, vrUnitUSD: 880000 }
      ]
    },
    {
      id: "A-ALCOHOL-ISOPROPILICO",
      nombre: "ALCOHOL ISOPROPÍLICO",
      procesos: [
        { codigo: "AIS-010", analisis: "Pureza", metodo: "GC", cMtra_g: 10, cantidad: 1, vrUnit1: 420000, vrUnit2: 400000, vrUnit3: 380000, vrUnitUSD: 420000 },
        { codigo: "AIS-011", analisis: "Identificación", metodo: "IR", cMtra_g: 4, cantidad: 1, vrUnit1: 170000, vrUnit2: 160000, vrUnit3: 150000, vrUnitUSD: 170000 }
      ]
    },
    {
      id: "B-BENZOATO-SODIO",
      nombre: "BENZOATO DE SODIO",
      procesos: [
        { codigo: "BZS-020", analisis: "Ensayo", metodo: "HPLC", cMtra_g: 2, cantidad: 1, vrUnit1: 720000, vrUnit2: 690000, vrUnit3: 650000, vrUnitUSD: 720000 },
        { codigo: "BZS-021", analisis: "Impurezas", metodo: "HPLC", cMtra_g: 2, cantidad: 1, vrUnit1: 1150000, vrUnit2: 1100000, vrUnit3: 1050000, vrUnitUSD: 1150000 }
      ]
    },
    {
      id: "C-CAFEINA",
      nombre: "CAFEÍNA",
      procesos: [
        { codigo: "CAF-030", analisis: "Ensayo", metodo: "UV-Vis", cMtra_g: 1, cantidad: 1, vrUnit1: 260000, vrUnit2: 245000, vrUnit3: 230000, vrUnitUSD: 260000 },
        { codigo: "CAF-031", analisis: "Identificación", metodo: "IR", cMtra_g: 1, cantidad: 1, vrUnit1: 180000, vrUnit2: 170000, vrUnit3: 160000, vrUnitUSD: 180000 }
      ]
    },
    {
      id: "E-ETANOL",
      nombre: "ETANOL",
      procesos: [
        { codigo: "ETA-050", analisis: "Pureza", metodo: "GC", cMtra_g: 8, cantidad: 1, vrUnit1: 400000, vrUnit2: 390000, vrUnit3: 370000, vrUnitUSD: 400000 }
      ]
    }
  ];

  const $alphabetNav = document.getElementById("alphabetNav");
  const $productsContainer = document.getElementById("productsContainer");
  const $noResults = document.getElementById("noResults");
  const $selectedCount = document.getElementById("selectedCount");
  const $cartCount = document.getElementById("cartCount");
  const $openCartBtn = document.getElementById("openCartBtn");
  const $btnGeneratePDF = document.getElementById("btnGeneratePDF");
  const $clientName = document.getElementById("clientName");
  const $clientEmail = document.getElementById("clientEmail");
  const $quoteDate = document.getElementById("quoteDate");
  const $historyTableBody = document.querySelector("#historyTable tbody");
  const $noHistory = document.getElementById("noHistory");

  const $filterClient = document.getElementById("filterClient");
  const $filterProduct = document.getElementById("filterProduct");
  const $filterFrom = document.getElementById("filterFrom");
  const $filterTo = document.getElementById("filterTo");
  const $btnApplyFilters = document.getElementById("btnApplyFilters");
  const $btnClearFilters = document.getElementById("btnClearFilters");
  const $btnExportXlsx = document.getElementById("btnExportXlsx");
  const $btnLoadExcel = document.getElementById("btnLoadExcel");
  const $excelInput = document.getElementById("excelInput");
  const $productSearch = document.getElementById("productSearch");
  const $scrollTopBtn = document.getElementById("scrollTopBtn");
  const $statsCanvas = document.getElementById("statsTopProducts");
  const $noStatsEl = document.getElementById("noStats");
  let logoAsset = null; // { el?: HTMLImageElement, dataUrl?: string }
  let topProductsChart = null;

  let selectedProductIds = new Set();
  let currentLetter = "A";
  let currentSearchQuery = "";

  function init() {
    // Fecha por defecto
    if (!$quoteDate.value) {
      $quoteDate.valueAsDate = new Date();
    }
    // Cargar catálogo previo si existe
    const persistedCatalog = loadCatalog();
    if (persistedCatalog && Array.isArray(persistedCatalog) && persistedCatalog.length > 0) {
      productos = persistedCatalog;
    }
    buildAlphabetNav();
    renderProductsByLetter(currentLetter);
    renderHistory();

    $btnGeneratePDF.addEventListener("click", onGeneratePDF);
    $btnApplyFilters.addEventListener("click", () => renderHistory(getHistoryFilters()));
    if ($btnClearFilters) $btnClearFilters.addEventListener("click", clearFilters);
    if ($btnExportXlsx) $btnExportXlsx.addEventListener("click", () => exportHistoryToXlsx(getHistoryFilters()));

    setupMainNav();

    if ($openCartBtn) {
      $openCartBtn.addEventListener("click", () => {
        const modalEl = document.getElementById("quoteModal");
        if (modalEl && typeof bootstrap !== "undefined") {
          const modal = bootstrap.Modal.getOrCreateInstance(modalEl);
          modal.show();
        }
      });
    }

    if ($btnLoadExcel && $excelInput) {
      $btnLoadExcel.addEventListener("click", () => $excelInput.click());
      $excelInput.addEventListener("change", onExcelSelected);
    }

    if ($productSearch) {
      $productSearch.addEventListener("input", () => {
        currentSearchQuery = ($productSearch.value || "").trim().toLowerCase();
        renderProductsByLetter(currentLetter);
      });
    }

    // Scroll-to-top
    if ($scrollTopBtn) {
      $scrollTopBtn.addEventListener("click", () => window.scrollTo({ top: 0, behavior: "smooth" }));
      window.addEventListener("scroll", onScrollToggleScrollTop);
      onScrollToggleScrollTop();
    }
  }

  function onExcelSelected(e) {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const data = new Uint8Array(ev.target.result);
        const wb = XLSX.read(data, { type: "array" });
        productos = buildProductsFromWorkbook(wb);
        saveCatalog(productos);
        selectedProductIds.clear();
        updateSelectionStateUI();
        renderProductsByLetter(currentLetter);
      } catch (err) {
        alert("No se pudo leer el Excel. Verifica el formato.");
        console.error(err);
      } finally {
        e.target.value = "";
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function buildProductsFromWorkbook(wb) {
    const products = [];
    (wb.SheetNames || []).forEach((sheetName) => {
      const sheet = wb.Sheets[sheetName];
      if (!sheet) return;
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: "" });
      if (!rows || rows.length === 0) return;

      let i = 0;
      let lastProductName = null;
      let tableIndex = 0;
      while (i < rows.length) {
        const row = ensureArray(rows[i]);
        const maybeName = extractProductNameFromRow(row);
        if (maybeName) {
          lastProductName = maybeName;
          i += 1;
          continue;
        }
        if (!isHeaderRow(row)) { i += 1; continue; }

        const colIndex = mapHeaderIndices(row);
        if (!colIndex) { i += 1; continue; }
        i += 1; // avanzar a la primera fila de datos

        const procesos = [];
        for (; i < rows.length; i++) {
          const r = ensureArray(rows[i]);
          if (extractProductNameFromRow(r) || isHeaderRow(r)) break; // nueva tabla o nuevo producto
          const textRow = r.map((x) => String(x || "").trim());
          const isEmpty = textRow.every((c) => c === "");
          const hasTotal = textRow.some((c) => normalizeText(c) === "total");
          if (isEmpty || hasTotal) { i += 1; break; }

          const get = (idx) => (idx != null && idx >= 0 ? r[idx] : "");
          const proceso = {
            codigo: String(get(colIndex.codigo) || "").trim(),
            analisis: String(get(colIndex.analisis) || "").trim(),
            metodo: String(get(colIndex.metodo) || "").trim(),
            cMtra: String(get(colIndex.cmtra_g) || "").trim(),
            cantidad: String(get(colIndex.cantidad) || "").trim(),
            vrUnit1: String(get(colIndex.vrunit1) || "").trim(),
            vrUnit2: String(get(colIndex.vrunit2) || "").trim(),
            vrUnit3: String(get(colIndex.vrunit3) || "").trim(),
            vrUnitUSD: String(get(colIndex.vrunit) || "").trim(),
            vrTotal: String(get(colIndex.vrtotal) || "").trim()
          };
          const meaningful = proceso.analisis || proceso.metodo || proceso.vrUnitUSD > 0 || proceso.cMtra_g > 0;
          if (meaningful) procesos.push(proceso);
        }

        const nameForThisTable = lastProductName || `${sheetName} ${++tableIndex}`;
        lastProductName = null; // reset para la próxima tabla

        if (procesos.length > 0) {
          products.push({
            id: `${nameForThisTable.charAt(0).toUpperCase()}-${sanitizeFilename(nameForThisTable).toUpperCase()}`,
            nombre: String(nameForThisTable).toUpperCase(),
            procesos
          });
        }
      }
    });
    return products;
  }

  function productoMatchRegex() { return /^producto\s*:/i; }

  function ensureArray(r) { return Array.isArray(r) ? r : []; }

  function isHeaderRow(r) {
    const cells = ensureArray(r).map((c) => slug(String(c || "")));
    const hasCodigo = cells.includes("codigo");
    const hasAnalisis = cells.includes("analisis");
    const hasMetodo = cells.includes("metodo");
    return (hasCodigo && hasAnalisis) || (hasAnalisis && hasMetodo);
  }

  function mapHeaderIndices(r) {
    const cells = ensureArray(r).map((c) => slug(String(c || "")));
    const findIdx = (...keys) => {
      const set = new Set(keys.map((k) => slug(k)));
      let idx = -1;
      cells.forEach((c, i) => { if (set.has(c)) idx = i; });
      return idx >= 0 ? idx : null;
    };
    const map = {
      codigo: findIdx("codigo", "code"),
      analisis: findIdx("analisis", "analysis"),
      metodo: findIdx("metodo", "method"),
      cmtra_g: findIdx("cmtrag", "cmtra", "cmtra g", "cmtra[g]"),
      cantidad: findIdx("cant", "cantidad"),
      vrunit1: findIdx("vrunitario1", "vrunit1"),
      vrunit2: findIdx("vrunitario2", "vrunit2"),
      vrunit3: findIdx("vrunitario3", "vrunit3"),
      vrunit: findIdx("vrunitariousd", "vrunitario", "vrunitariocop", "vrunitariousd"),
      vrtotal: findIdx("vrtotal")
    };
    if (map.analisis == null && map.metodo == null) return null;
    return map;
  }

  function extractProductNameFromRow(r) {
    const cells = ensureArray(r).map((c) => String(c || "").trim());
    // Caso 1: una sola celda tipo "Producto: Nombre"
    for (let c of cells) {
      const m = /^\s*producto\s*:\s*(.+)$/i.exec(c);
      if (m && m[1]) return m[1].trim();
    }
    // Caso 2: celda "Producto" y nombre en la siguiente celda no vacía
    const idx = cells.findIndex((c) => normalizeText(c) === "producto");
    if (idx >= 0) {
      for (let j = idx + 1; j < cells.length; j++) {
        if (String(cells[j]).trim() !== "") return String(cells[j]).trim();
      }
    }
    return null;
  }

  function normalizeText(s) {
    return String(s)
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/\s+/g, " ")
      .trim();
  }

  function slug(s) {
    return String(s)
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[^a-z0-9]+/g, "");
  }

  function toNumber(v, fallback = 0) {
    if (v == null) return fallback;
    if (typeof v === "number" && isFinite(v)) return v;
    const s = String(v).replace(/[^0-9,.-]/g, "").replace(/\.(?=.*\.)/g, "").replace(/,(?=\d{3}\b)/g, "");
    const val = Number(s.replace(",", "."));
    return isNaN(val) ? fallback : val;
  }

  function setupMainNav() {
    const nav = document.getElementById("mainNav");
    const viewProducts = document.getElementById("viewProducts");
    const viewHistory = document.getElementById("viewHistory");
    const viewStats = document.getElementById("viewStats");
    if (!nav || !viewProducts || !viewHistory || !viewStats) return;
    const links = nav.querySelectorAll(".nav-link");
    links.forEach((link) => {
      link.addEventListener("click", () => {
        links.forEach((l) => l.classList.remove("active"));
        link.classList.add("active");
        const target = link.getAttribute("data-target");
        const showHistory = target === "#viewHistory";
        const showStats = target === "#viewStats";
        if (showHistory) {
          viewProducts.classList.add("d-none");
          viewStats.classList.add("d-none");
          viewHistory.classList.remove("d-none");
        } else if (showStats) {
          viewProducts.classList.add("d-none");
          viewHistory.classList.add("d-none");
          viewStats.classList.remove("d-none");
          renderStats();
        } else {
          viewHistory.classList.add("d-none");
          viewStats.classList.add("d-none");
          viewProducts.classList.remove("d-none");
        }
      });
    });
  }

  function clearFilters() {
    $filterClient.value = "";
    $filterProduct.value = "";
    $filterFrom.value = "";
    $filterTo.value = "";
    renderHistory();
  }

  function buildAlphabetNav() {
    const fragment = document.createDocumentFragment();
    const letters = Array.from({ length: 26 }, (_, i) => String.fromCharCode(65 + i));

    const allBtn = document.createElement("button");
    allBtn.type = "button";
    allBtn.className = `btn btn-sm btn-outline-secondary alphabet-btn`;
    allBtn.textContent = "Todos";
    allBtn.addEventListener("click", () => {
      currentLetter = "*";
      updateAlphabetActive();
      renderProductsByLetter(currentLetter);
    });
    fragment.appendChild(allBtn);

    letters.forEach((letter) => {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.dataset.letter = letter;
      btn.className = `btn btn-sm btn-outline-primary alphabet-btn`;
      btn.textContent = letter;
      btn.addEventListener("click", () => {
        currentLetter = letter;
        updateAlphabetActive();
        renderProductsByLetter(letter);
      });
      fragment.appendChild(btn);
    });
    $alphabetNav.appendChild(fragment);
    updateAlphabetActive();
    // Asegurar inicio al principio en móviles
    try { $alphabetNav.scrollLeft = 0; } catch {}
  }

  function updateAlphabetActive() {
    const buttons = $alphabetNav.querySelectorAll("button");
    buttons.forEach((b) => b.classList.remove("active"));
    const active = Array.from(buttons).find((b) => (currentLetter === "*" ? b.textContent === "Todos" : b.textContent === currentLetter));
    if (active) active.classList.add("active");
  }

  function renderProductsByLetter(letter) {
    $productsContainer.innerHTML = "";
    const list = productos
      .slice()
      .sort((a, b) => a.nombre.localeCompare(b.nombre, "es", { sensitivity: "base" }))
      .filter((p) => {
        const matchesSearch = currentSearchQuery ? p.nombre.toLowerCase().includes(currentSearchQuery) : true;
        const matchesLetter = currentSearchQuery ? true : (letter === "*" ? true : p.nombre.trim().toUpperCase().startsWith(letter));
        return matchesSearch && matchesLetter;
      });

    if (list.length === 0) {
      $noResults.classList.remove("d-none");
      return;
    }
    $noResults.classList.add("d-none");

    const fragment = document.createDocumentFragment();
    list.forEach((prod) => {
      fragment.appendChild(renderProductCard(prod));
    });
    $productsContainer.appendChild(fragment);
    updateSelectionStateUI();
  }

  function renderProductCard(product) {
    const col = document.createElement("div");
    col.className = "col-12"; // ancho completo por fila

    const card = document.createElement("div");
    card.className = "card product-card h-100";

    const header = document.createElement("div");
    header.className = "card-header d-flex justify-content-between align-items-start gap-2";
    header.innerHTML = `
      <div class="product-title">
        <input class="form-check-input me-2 product-select" type="checkbox" data-product-id="${product.id}">
        <span class="fw-semibold">${product.nombre}</span>
      </div>
      <span class="badge bg-light text-dark badge-unit">COP</span>
    `;

    const body = document.createElement("div");
    body.className = "card-body p-0";
    body.appendChild(renderProductTable(product));

    card.appendChild(header);
    card.appendChild(body);
    col.appendChild(card);

    const checkbox = header.querySelector(".product-select");
    checkbox.checked = selectedProductIds.has(product.id);
    checkbox.addEventListener("change", (e) => {
      if (e.target.checked) {
        selectedProductIds.add(product.id);
      } else {
        selectedProductIds.delete(product.id);
      }
      updateSelectionStateUI();
    });

    return col;
  }

  function renderProductTable(product) {
    const tableWrapper = document.createElement("div");
    tableWrapper.className = "table-responsive";
    const table = document.createElement("table");
    table.className = "table table-sm table-striped table-hover mb-0";

    const thead = document.createElement("thead");
    thead.innerHTML = `
      <tr>
        <th>Código</th>
        <th>Análisis</th>
        <th>Método</th>
        <th>C. Mtra. [g]</th>
        <th>Cant.</th>
        <th>Vr. Unitario 1</th>
        <th>Vr. Unitario 2</th>
        <th>Vr. Unitario 3</th>
        <th>Vr. Unitario USD</th>
        <th>Vr. Total</th>
      </tr>
    `;

    const tbody = document.createElement("tbody");
    let sumCMtra = 0;
    let sumTotal = 0;
    product.procesos.forEach((row) => {
      const getText = (v) => (v == null ? "" : String(v));
      const valueOrFormat = (text, num, isMoney = false) => {
        if (text !== undefined && text !== null && String(text) !== "") return String(text);
        if (num === undefined || num === null || String(num) === "") return "";
        return isMoney ? formatMoney(num) : formatNumber(num);
      };

      // Sumas solo para el pie (no cambian valores mostrados)
      const cMtraForSum = parseNumStrict(row.cMtra ?? row.cMtra_g);
      const vrTotalForSum = parseNumStrict(row.vrTotal);
      sumCMtra += cMtraForSum;
      sumTotal += vrTotalForSum;

      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${getText(row.codigo)}</td>
        <td>${getText(row.analisis)}</td>
        <td>${getText(row.metodo)}</td>
        <td class=\"text-center\">${valueOrFormat(row.cMtra, row.cMtra_g)}</td>
        <td class=\"text-center\">${valueOrFormat(row.cantidad, row.cantidad)}</td>
        <td class=\"text-center\">${valueOrFormat(row.vrUnit1)}</td>
        <td class=\"text-center\">${valueOrFormat(row.vrUnit2)}</td>
        <td class=\"text-center\">${valueOrFormat(row.vrUnit3)}</td>
        <td class=\"text-center\">${valueOrFormat(row.vrUnitUSD)}</td>
        <td class=\"text-center\">${valueOrFormat(row.vrTotal)}</td>
      `;
      tbody.appendChild(tr);
    });

    const tfoot = document.createElement("tfoot");
    tfoot.innerHTML = `
      <tr>
        <td colspan="3" class="text-end">Totales:</td>
        <td class="text-center">${formatNumber(sumCMtra)}</td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td class="text-center fw-semibold">Subtotal</td>
        <td class="text-center fw-semibold">${formatMoney(sumTotal)}</td>
      </tr>
    `;

    table.appendChild(thead);
    table.appendChild(tbody);
    table.appendChild(tfoot);
    tableWrapper.appendChild(table);

    // Guardar totales calculados en el objeto para uso en PDF
    product._sumCMtra = sumCMtra;
    product._sumTotal = sumTotal;
    return tableWrapper;
  }

  function updateSelectionStateUI() {
    const count = selectedProductIds.size;
    $selectedCount.textContent = `${count} seleccionados`;
    $btnGeneratePDF.disabled = count === 0;
    if ($cartCount) $cartCount.textContent = `${count}`;
    // Sincronizar checkboxes visibles
    document.querySelectorAll(".product-select").forEach((cb) => {
      const id = cb.getAttribute("data-product-id");
      cb.checked = selectedProductIds.has(id);
    });
  }

  async function onGeneratePDF() {
    const clientName = $clientName.value.trim();
    const clientEmail = $clientEmail.value.trim();
    const dateStr = $quoteDate.value ? $quoteDate.value : new Date().toISOString().slice(0, 10);

    if (!clientName) {
      alert("Por favor ingresa el nombre del cliente.");
      return;
    }

    const selectedProducts = productos
      .filter((p) => selectedProductIds.has(p.id))
      .map(recomputeTotalsForProduct);

    if (selectedProducts.length === 0) {
      alert("Selecciona al menos un producto.");
      return;
    }

    // Totales
    const totalGeneral = selectedProducts.reduce((acc, p) => acc + (p._sumTotal || 0), 0);

    // Cargar logo (cache) y generar PDF
    if (!logoAsset) {
      try { logoAsset = await loadLogo(); } catch {}
    }
    generatePDF({ clientName, clientEmail, dateStr, products: selectedProducts, totalGeneral, logo: logoAsset });

    // Guardar historial
    const quote = {
      id: `Q-${Date.now()}`,
      date: dateStr,
      clientName,
      clientEmail,
      products: selectedProducts.map((p) => ({ id: p.id, nombre: p.nombre, subtotal: p._sumTotal })),
      totalCOP: totalGeneral
    };
    saveQuote(quote);
    renderHistory();
    renderStats();

    // Cerrar modal y resetear formulario
    closeQuoteModal();
    resetClientForm();
    clearCartSelection();
  }

  function recomputeTotalsForProduct(product) {
    let sumCMtra = 0;
    let sumTotal = 0;
    product.procesos.forEach((row) => {
      const cMtraNum = parseNumStrict(row.cMtra_g ?? row.cMtra);
      const qty = parseNumStrict(row.cantidad);
      const unit = parseNumStrict(row.vrUnitUSD ?? row.vrUnit1 ?? row.vrUnit2 ?? row.vrUnit3);
      const rowTotal = parseNumStrict(row.vrTotal);
      if (isFinite(cMtraNum)) sumCMtra += cMtraNum;
      if (rowTotal > 0) {
        sumTotal += rowTotal;
      } else if (unit > 0 && qty > 0) {
        sumTotal += unit * qty;
      }
    });
    return { ...product, _sumCMtra: sumCMtra, _sumTotal: sumTotal };
  }

  function closeQuoteModal() {
    const modalEl = document.getElementById("quoteModal");
    if (modalEl && typeof bootstrap !== "undefined") {
      const modal = bootstrap.Modal.getOrCreateInstance(modalEl);
      modal.hide();
    }
  }

  function resetClientForm() {
    const form = document.getElementById("clientForm");
    if (form) form.reset();
    if ($quoteDate) {
      $quoteDate.value = "";
      $quoteDate.valueAsDate = new Date();
    }
  }

  function clearCartSelection() {
    selectedProductIds.clear();
    updateSelectionStateUI();
  }

  function generatePDF({ clientName, clientEmail, dateStr, products, totalGeneral, logo }) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ unit: "pt", format: "a4" });
    // Colores de marca (alineados con la web)
    const brandPrimary = [68, 194, 196]; // #44c2c4
    const brandSecondary = [243, 192, 44]; // #f3c02c
    const dark = [33, 37, 41];

    const pageWidth = doc.internal.pageSize.getWidth();
    const left = 40;
    const right = 555; // ancho útil para alinear a la derecha

    // Logo centrado arriba con tamaño fijo solicitado
    const logoW = 200; // ancho fijo en px
    const logoH = 65;  // alto fijo en px
    const xLogo = (pageWidth - logoW) / 2;
    let yTop = 24 + logoH + 36; // espacio garantizado bajo el logo
    try {
      if (logo && logo.el) {
        try { doc.addImage(logo.el, "PNG", xLogo, 24, logoW, logoH); }
        catch {
          const data = getLogoDataUrlSync(logo);
          if (data) doc.addImage(data, "PNG", xLogo, 24, logoW, logoH);
        }
      } else if (logo && logo.dataUrl) {
        doc.addImage(logo.dataUrl, "PNG", xLogo, 24, logoW, logoH);
      }
    } catch {}

    // Encabezado empresa (bloque izquierdo)
    doc.setFont("helvetica", "bold");
    doc.setFontSize(14);
    doc.setTextColor(...brandPrimary);
    doc.text("Olgroup - Laboratorio de Soluciones Analíticas", left, yTop);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(10);
    doc.setTextColor(0, 0, 0);
    doc.text("Servicios analíticos y control de calidad", left, yTop + 16);
    doc.text("Email: contacto@olgroup.example | Tel: +57 000 000 0000", left, yTop + 28);

    // Bloque derecho (cotización + cliente + fecha) en la misma línea del encabezado
    doc.setFont("helvetica", "bold");
    doc.setFontSize(14);
    doc.setTextColor(...brandPrimary);
    doc.text("Cotización", right, yTop, { align: "right" });
    doc.setFont("helvetica", "normal");
    doc.setFontSize(10);
    doc.setTextColor(0, 0, 0);
    doc.text(`Cliente: ${clientName}`, right, yTop + 16, { align: "right" });
    doc.text(`Fecha: ${formatDateHuman(dateStr)}`, right, yTop + 28, { align: "right" });

    // Separador sutil con color de marca
    doc.setDrawColor(...brandPrimary);
    doc.line(left, yTop + 40, 555, yTop + 40);
    let y = yTop + 58;

    // Productos
    products.forEach((p, idx) => {
      if (idx > 0) {
        y += 8;
      }
      doc.setFont("helvetica", "bold");
      doc.setFontSize(11);
      doc.setTextColor(...brandPrimary);
      doc.text(p.nombre, left, y);
      doc.setFont("helvetica", "normal");
      doc.setFontSize(9);
      doc.setTextColor(0, 0, 0);
      y += 12;

      const tableData = p.procesos.map((row) => [
        String(row.codigo ?? ""),
        String(row.analisis ?? ""),
        String(row.metodo ?? ""),
        String((row.cMtra ?? row.cMtra_g) ?? ""),
        String(row.cantidad ?? ""),
        String(row.vrUnit1 ?? ""),
        String(row.vrUnit2 ?? ""),
        String(row.vrUnit3 ?? ""),
        String(row.vrUnitUSD ?? ""),
        String(row.vrTotal ?? "")
      ]);

      doc.autoTable({
        head: [["Código", "Análisis", "Método", "C. Mtra. [g]", "Cant.", "Vr. Unitario 1", "Vr. Unitario 2", "Vr. Unitario 3", "Vr. Unitario COP", "Vr. Total"]],
        body: tableData,
        startY: y,
        styles: { fontSize: 8 },
        headStyles: { fillColor: brandPrimary, textColor: [255, 255, 255] },
        columnStyles: {
          3: { halign: "right" },
          4: { halign: "right" },
          5: { halign: "right" },
          6: { halign: "right" },
          7: { halign: "right" },
          8: { halign: "right" },
          9: { halign: "right" }
        },
        didDrawPage: (data) => {},
        willDrawCell: (data) => {}
      });
      y = doc.lastAutoTable.finalY + 16;

      // Subtotal por producto (alineado a la derecha bajo la tabla)
      doc.setFont("helvetica", "bold");
      doc.setTextColor(...brandSecondary);
      doc.text(`Subtotal: ${formatMoney(p._sumTotal)}`, right, y, { align: "right" });
      doc.setTextColor(0, 0, 0);
      doc.setFont("helvetica", "normal");
      y += 12;
    });

    // Total general
    y += 4;
    doc.setDrawColor(...brandPrimary);
    doc.line(left, y, 555, y);
    y += 16;
    doc.setFont("helvetica", "bold");
    doc.setFontSize(12);
    doc.setTextColor(...brandSecondary);
    doc.text(`Total general (COP): ${formatMoney(totalGeneral)}`, left, y);
    doc.setTextColor(0, 0, 0);

    // Pie de página
    doc.setFont("helvetica", "normal");
    doc.setFontSize(8);
    doc.text("Esta cotización es válida por 15 días. Precios en COP.", left, 812);

    doc.save(`Cotizacion_${sanitizeFilename(clientName)}_${dateStr}.pdf`);
  }

  function saveQuote(quote) {
    const key = "olgroup_quotes";
    const list = JSON.parse(localStorage.getItem(key) || "[]");
    list.unshift(quote);
    localStorage.setItem(key, JSON.stringify(list));
  }

  function getQuotes() {
    try {
      return JSON.parse(localStorage.getItem("olgroup_quotes") || "[]");
    } catch {
      return [];
    }
  }

  function deleteQuote(id) {
    const list = getQuotes().filter((q) => q.id !== id);
    localStorage.setItem("olgroup_quotes", JSON.stringify(list));
  }

  function getHistoryFilters() {
    return {
      client: $filterClient.value.trim().toLowerCase(),
      product: $filterProduct.value.trim().toLowerCase(),
      from: $filterFrom.value,
      to: $filterTo.value
    };
  }

  function renderHistory(filters) {
    const all = getQuotes();
    const list = (filters ? applyFilters(all, filters) : all).slice();
    $historyTableBody.innerHTML = "";

    if (list.length === 0) {
      $noHistory.classList.remove("d-none");
      return;
    }
    $noHistory.classList.add("d-none");

    list.forEach((q, idx) => {
      const tr = document.createElement("tr");
      const productsText = q.products.map((p) => p.nombre).join(", ");
      tr.innerHTML = `
        <td>${idx + 1}</td>
        <td>${formatDateHuman(q.date)}</td>
        <td>${escapeHtml(q.clientName)}</td>
        <td>${escapeHtml(productsText)}</td>
        <td class="text-center">${formatMoney(q.totalCOP != null ? q.totalCOP : q.totalUSD)}</td>
        <td class="text-nowrap">
          <button class="btn btn-sm btn-outline-primary me-1" data-action="view" data-id="${q.id}">PDF</button>
          <button class="btn btn-sm btn-outline-danger" data-action="delete" data-id="${q.id}">Borrar</button>
        </td>
      `;
      $historyTableBody.appendChild(tr);
    });

    $historyTableBody.querySelectorAll("button").forEach((btn) => {
      btn.addEventListener("click", () => onHistoryAction(btn.dataset.action, btn.dataset.id));
    });

    // Actualizar estadísticas al renderizar historial completo
    try { renderStats(); } catch {}
  }

  function exportHistoryToXlsx(filters) {
    if (typeof XLSX === "undefined") {
      alert("No se pudo cargar la librería de Excel. Verifica tu conexión.");
      return;
    }
    const all = getQuotes();
    const list = (filters ? applyFilters(all, filters) : all).slice();
    if (list.length === 0) {
      alert("No hay registros para exportar.");
      return;
    }
    const rows = list.map((q, idx) => ({
      "#": idx + 1,
      Fecha: formatDateHuman(q.date),
      Cliente: q.clientName,
      Productos: (q.products || []).map((p) => p.nombre).join(", "),
      "Total COP": Number((q.totalCOP != null ? q.totalCOP : (q.totalUSD != null ? q.totalUSD : 0)))
    }));
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(rows, { header: ["#", "Fecha", "Cliente", "Productos", "Total COP"] });
    XLSX.utils.book_append_sheet(wb, ws, "Historial");
    const filename = `Historial_Cotizaciones_${new Date().toISOString().slice(0, 10)}.xlsx`;
    XLSX.writeFile(wb, filename);
  }

  // Estadísticas: productos más cotizados
  function computeTopProducts(limit = 10) {
    const quotes = getQuotes();
    const countByName = new Map();
    quotes.forEach((q) => {
      (q.products || []).forEach((p) => {
        const name = String(p.nombre || "").trim();
        if (!name) return;
        countByName.set(name, (countByName.get(name) || 0) + 1);
      });
    });
    const sorted = Array.from(countByName.entries()).sort((a, b) => b[1] - a[1]).slice(0, limit);
    return { labels: sorted.map((x) => x[0]), data: sorted.map((x) => x[1]) };
  }

  function renderStats() {
    if (!$statsCanvas) return;
    const { labels, data } = computeTopProducts(10);
    const hasData = labels.length > 0;
    if ($noStatsEl) $noStatsEl.classList.toggle("d-none", hasData);
    $statsCanvas.classList.toggle("d-none", !hasData);
    if (!hasData) {
      if (topProductsChart) { try { topProductsChart.destroy(); } catch {} topProductsChart = null; }
      return;
    }
    const ctx = $statsCanvas.getContext("2d");
    if (topProductsChart) { try { topProductsChart.destroy(); } catch {} }
    topProductsChart = new Chart(ctx, {
      type: "bar",
      data: {
        labels,
        datasets: [{
          label: "Veces cotizado",
          data,
          backgroundColor: "rgba(68, 194, 196, 0.6)",
          borderColor: "rgb(68, 194, 196)",
          borderWidth: 1
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        scales: {
          y: { beginAtZero: true, ticks: { precision: 0, stepSize: 1 } }
        },
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: { label: (ctx) => ` ${ctx.parsed.y} cotizaciones` } }
        }
      }
    });
  }

  async function loadLogo() {
    // 1) Intentar usar la imagen existente en el DOM para evitar CORS
    try {
      const el = document.querySelector(".brand-logo");
      if (el && el.complete && el.naturalWidth > 0) {
        try { return { el, dataUrl: elementToDataUrl(el) }; } catch { return { el }; }
      }
    } catch {}
    // 2) Probar una lista de rutas comunes para el logo
    const candidates = [
      "logo.png",
      "logo2.png",
      "assets/logo.png",
      "assets/logo2.png"
    ];
    for (const src of candidates) {
      try {
        const img = await loadImageElement(src);
        try { return { el: img, dataUrl: elementToDataUrl(img) }; } catch { return { el: img }; }
      } catch {}
    }
    return null;
  }

  function getLogoDataUrlSync(logo) {
    if (!logo) return null;
    if (logo.dataUrl) return logo.dataUrl;
    try {
      if (logo.el) return elementToDataUrl(logo.el);
    } catch {}
    return null;
  }

  function elementToDataUrl(imgEl) {
    const w = Math.max(1, imgEl.naturalWidth || 64);
    const h = Math.max(1, imgEl.naturalHeight || 64);
    const canvas = document.createElement("canvas");
    canvas.width = w;
    canvas.height = h;
    const ctx = canvas.getContext("2d");
    ctx.drawImage(imgEl, 0, 0, w, h);
    return canvas.toDataURL("image/png");
  }


  function loadImageElement(src) {
    return new Promise((resolve, reject) => {
      const img = new Image();
      try { img.crossOrigin = "anonymous"; } catch {}
      img.onload = () => resolve(img);
      img.onerror = reject;
      img.src = src;
    });
  }

  // Calcula tamaño de render conservando proporción dentro de un límite
  function getLogoRenderSize(logo, maxW, maxH) {
    let naturalW = 64;
    let naturalH = 64;
    try {
      if (logo && logo.el) {
        naturalW = Math.max(1, Number(logo.el.naturalWidth || logo.el.width || 64));
        naturalH = Math.max(1, Number(logo.el.naturalHeight || logo.el.height || 64));
      }
    } catch {}
    const scale = Math.max(0.0001, Math.min(maxW / naturalW, maxH / naturalH));
    const w = Math.round(naturalW * scale);
    const h = Math.round(naturalH * scale);
    return { w, h };
  }

  async function onHistoryAction(action, id) {
    const list = getQuotes();
    const q = list.find((x) => x.id === id);
    if (!q) return;
    if (action === "view") {
      // Reconstruir productos desde nombres y subtotales. Usamos catálogo actual para filas.
      const selected = (q.products || [])
        .map((qp) => productos.find((p) => p.id === qp.id) || null)
        .filter(Boolean)
        .map(recomputeTotalsForProduct);
      // Asegurar que el logo esté disponible también al abrir desde historial
      if (!logoAsset) {
        try { logoAsset = await loadLogo(); } catch {}
      }
      generatePDF({ clientName: q.clientName, clientEmail: q.clientEmail, dateStr: q.date, products: selected, totalGeneral: (q.totalCOP != null ? q.totalCOP : q.totalUSD), logo: logoAsset });
    } else if (action === "delete") {
      showConfirm("¿Borrar esta cotización del historial?", "Borrar").then((ok) => {
        if (ok) {
          deleteQuote(id);
          renderHistory(getHistoryFilters());
          renderStats();
        }
      });
    }
  }

  function showConfirm(message, acceptLabel) {
    return new Promise((resolve) => {
      const modalEl = document.getElementById("confirmModal");
      const msgEl = document.getElementById("confirmMessage");
      const acceptBtn = document.getElementById("confirmAcceptBtn");
      if (!modalEl || !msgEl || !acceptBtn || typeof bootstrap === "undefined") {
        // Fallback seguro si no carga Bootstrap
        resolve(window.confirm(message || "¿Confirmar?"));
        return;
      }
      msgEl.textContent = message || "¿Confirmar?";
      if (acceptLabel) acceptBtn.textContent = acceptLabel;
      const modal = bootstrap.Modal.getOrCreateInstance(modalEl);
      let resolved = false;
      const onAccept = () => {
        resolved = true;
        cleanup();
        modal.hide();
        resolve(true);
      };
      const onHidden = () => {
        if (!resolved) {
          cleanup();
          resolve(false);
        }
      };
      function cleanup() {
        acceptBtn.removeEventListener("click", onAccept);
        modalEl.removeEventListener("hidden.bs.modal", onHidden);
      }
      acceptBtn.addEventListener("click", onAccept, { once: true });
      modalEl.addEventListener("hidden.bs.modal", onHidden, { once: true });
      modal.show();
    });
  }

  function applyFilters(list, { client, product, from, to }) {
    return list.filter((q) => {
      const matchesClient = client ? (q.clientName || "").toLowerCase().includes(client) : true;
      const matchesProduct = product ? (q.products || []).some((p) => (p.nombre || "").toLowerCase().includes(product)) : true;
      const date = q.date || "";
      const matchesFrom = from ? date >= from : true;
      const matchesTo = to ? date <= to : true;
      return matchesClient && matchesProduct && matchesFrom && matchesTo;
    });
  }

  function onScrollToggleScrollTop() {
    if (!$scrollTopBtn) return;
    if (window.scrollY > 200) {
      $scrollTopBtn.classList.add("show");
    } else {
      $scrollTopBtn.classList.remove("show");
    }
  }

  // Utilidades
  function formatMoney(n) {
    const num = Number(n) || 0;
    return num.toLocaleString("es-CO", { style: "currency", currency: "COP", minimumFractionDigits: 0, maximumFractionDigits: 0 });
  }
  function formatNumber(n) {
    const num = Number(n) || 0;
    return num.toLocaleString("es-CO", { maximumFractionDigits: 4 });
  }
  function parseNumStrict(v) {
    if (v == null) return 0;
    if (typeof v === "number" && isFinite(v)) return v;
    const s = String(v).replace(/[^0-9,.-]/g, "").replace(/\.(?=.*\.)/g, "").replace(/,(?=\d{3}\b)/g, "");
    const val = Number(s.replace(",", "."));
    return isNaN(val) ? 0 : val;
  }

  function saveCatalog(list) {
    try { localStorage.setItem("olgroup_catalog", JSON.stringify(list || [])); } catch {}
  }
  function loadCatalog() {
    try { return JSON.parse(localStorage.getItem("olgroup_catalog") || "null"); } catch { return null; }
  }
  function formatDateHuman(d) {
    const [y, m, da] = d.split("-").map((x) => Number(x));
    if (!y || !m || !da) return d;
    return `${da.toString().padStart(2, "0")}/${m.toString().padStart(2, "0")}/${y}`;
  }
  function sanitizeFilename(s) {
    return (s || "").replace(/[^a-z0-9\-_]+/gi, "_");
  }
  function escapeHtml(str) {
    return String(str)
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  }

  // Inicio robusto: si el DOM ya está listo, inicializa de inmediato
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();


