const STORAGE_KEYS = {
  inventory: 'newscan.inventory.v1',
  request: 'newscan.request.v1',
};

const state = {
  inventory: {},
  request: {},
};

const elements = {
  excelFile: document.getElementById('excelFile'),
  inventoryStatus: document.getElementById('inventoryStatus'),
  scanInput: document.getElementById('scanInput'),
  scanStatus: document.getElementById('scanStatus'),
  requestBody: document.getElementById('requestBody'),
  clearInventory: document.getElementById('clearInventory'),
  clearRequest: document.getElementById('clearRequest'),
  exportCsv: document.getElementById('exportCsv'),
  addManual: document.getElementById('addManual'),
  focusScan: document.getElementById('focusScan'),
};

function normalizeHeader(header) {
  return String(header || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .trim();
}

function findColumnIndex(headers, possibilities) {
  return headers.findIndex((header) => possibilities.some((alias) => header.includes(alias)));
}

function readCell(row, index) {
  if (index < 0) return '';
  return String(row[index] || '').trim();
}

function saveState() {
  localStorage.setItem(STORAGE_KEYS.inventory, JSON.stringify(state.inventory));
  localStorage.setItem(STORAGE_KEYS.request, JSON.stringify(state.request));
}

function loadState() {
  state.inventory = JSON.parse(localStorage.getItem(STORAGE_KEYS.inventory) || '{}');
  state.request = JSON.parse(localStorage.getItem(STORAGE_KEYS.request) || '{}');
}

function setStatus(target, text, isError = false) {
  target.textContent = text;
  target.style.color = isError ? '#b42222' : '#0d4f87';
}

function parseWorkbook(file) {
  const reader = new FileReader();
  reader.onload = (event) => {
    const data = new Uint8Array(event.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });

    if (!rows.length) {
      setStatus(elements.inventoryStatus, 'Le fichier est vide.', true);
      return;
    }

    const headers = rows[0].map(normalizeHeader);
    const barcodeIdx = findColumnIndex(headers, ['barcode_1', 'code barre', 'codebarre', 'barcode', 'ean', 'upc']);
    const productNumberIdx = findColumnIndex(headers, [
      'textbox6',
      'numero de produit',
      'numéro de produit',
      'produit',
      'product',
      'article',
      'designation',
      'nom',
      'title',
      'titre',
    ]);
    const shortDescIdx = findColumnIndex(headers, ['shortdesc', 'description courte']);
    const docNumberIdx = findColumnIndex(headers, ['docnumber', 'document']);
    const longDescIdx = findColumnIndex(headers, ['description']);
    const locationIdx = findColumnIndex(headers, ['locationid', 'location', 'emplacement', 'bin', 'zone']);
    const qtyIdx = findColumnIndex(headers, ['qty', 'quantite', 'quantité', 'quantity', 'stock']);

    if (barcodeIdx === -1 || locationIdx === -1) {
      setStatus(
        elements.inventoryStatus,
        'Colonnes introuvables. Vérifiez au minimum: barcode_1 (code barre) et locationId.',
        true
      );
      return;
    }

    const inventory = {};

    rows.slice(1).forEach((row) => {
      const barcode = readCell(row, barcodeIdx);
      if (!barcode) return;

      const productNumber = readCell(row, productNumberIdx);
      const shortDescription = readCell(row, shortDescIdx);
      const longDescription = readCell(row, longDescIdx);
      const documentNumber = readCell(row, docNumberIdx);

      const productLabel =
        productNumber || shortDescription || longDescription || documentNumber || readCell(row, productNumberIdx);

      inventory[barcode] = {
        barcode,
        product: productLabel,
        location: readCell(row, locationIdx),
        qty: Number(readCell(row, qtyIdx) || 0),
        shortDescription,
        longDescription,
        documentNumber,
      };
    });

    state.inventory = inventory;
    saveState();
    setStatus(elements.inventoryStatus, `${Object.keys(inventory).length} article(s) importé(s).`);
  };

  reader.readAsArrayBuffer(file);
}

function addScan(barcode) {
  const normalized = String(barcode || '').trim();
  if (!normalized) return;

  const item = state.inventory[normalized];
  if (!item) {
    setStatus(elements.scanStatus, `Code ${normalized} introuvable dans le stock importé.`, true);
    return;
  }

  if (!state.request[normalized]) {
    state.request[normalized] = { ...item, requestedQty: 0 };
  }

  state.request[normalized].requestedQty += 1;
  saveState();
  renderRequestTable();
  setStatus(elements.scanStatus, `${item.product || normalized} ajouté à la demande.`);
}

function removeLine(barcode) {
  delete state.request[barcode];
  saveState();
  renderRequestTable();
}

function renderRequestTable() {
  const lines = Object.values(state.request);
  if (!lines.length) {
    elements.requestBody.innerHTML = '<tr><td colspan="5">Aucun article scanné.</td></tr>';
    return;
  }

  elements.requestBody.innerHTML = lines
    .map(
      (line) => `
      <tr>
        <td>${line.barcode}</td>
        <td>${line.product || '-'}</td>
        <td>${line.location || '-'}</td>
        <td class="qty-cell">${line.requestedQty}</td>
        <td>
          <button class="inline-btn" data-action="minus" data-barcode="${line.barcode}">-1</button>
          <button class="inline-btn danger" data-action="remove" data-barcode="${line.barcode}">Suppr.</button>
        </td>
      </tr>
    `
    )
    .join('');
}

function exportRequestCsv() {
  const lines = Object.values(state.request);
  if (!lines.length) {
    setStatus(elements.scanStatus, 'La demande est vide, rien à exporter.', true);
    return;
  }

  const headers = ['barcode', 'product', 'location', 'requested_qty'];
  const csv = [
    headers.join(','),
    ...lines.map((line) =>
      [line.barcode, line.product, line.location, line.requestedQty]
        .map((value) => `"${String(value || '').replaceAll('"', '""')}"`)
        .join(',')
    ),
  ].join('\n');

  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = `demande-commande-${new Date().toISOString().slice(0, 10)}.csv`;
  link.click();
  URL.revokeObjectURL(url);
}

function bindEvents() {
  elements.excelFile.addEventListener('change', (event) => {
    const file = event.target.files?.[0];
    if (file) parseWorkbook(file);
  });

  elements.scanInput.addEventListener('keydown', (event) => {
    if (event.key === 'Enter') {
      event.preventDefault();
      addScan(elements.scanInput.value);
      elements.scanInput.value = '';
    }
  });

  elements.addManual.addEventListener('click', () => {
    addScan(elements.scanInput.value);
    elements.scanInput.value = '';
    elements.scanInput.focus();
  });

  elements.focusScan.addEventListener('click', () => {
    elements.scanInput.focus();
  });

  elements.clearInventory.addEventListener('click', () => {
    state.inventory = {};
    saveState();
    setStatus(elements.inventoryStatus, 'Stock importé effacé.');
  });

  elements.clearRequest.addEventListener('click', () => {
    state.request = {};
    saveState();
    renderRequestTable();
    setStatus(elements.scanStatus, 'Demande vidée.');
  });

  elements.exportCsv.addEventListener('click', exportRequestCsv);

  elements.requestBody.addEventListener('click', (event) => {
    const button = event.target.closest('button[data-action]');
    if (!button) return;

    const barcode = button.dataset.barcode;
    const action = button.dataset.action;

    if (action === 'remove') {
      removeLine(barcode);
      return;
    }

    if (action === 'minus' && state.request[barcode]) {
      state.request[barcode].requestedQty -= 1;
      if (state.request[barcode].requestedQty <= 0) {
        delete state.request[barcode];
      }
      saveState();
      renderRequestTable();
    }
  });
}

function initPwa() {
  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('sw.js').catch(() => {
      setStatus(elements.inventoryStatus, 'Mode offline non disponible.', true);
    });
  }
}

(function init() {
  loadState();
  bindEvents();
  renderRequestTable();
  setStatus(elements.inventoryStatus, `${Object.keys(state.inventory).length} article(s) en local.`);
  initPwa();
  elements.scanInput.focus();
})();
