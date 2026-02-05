const excelFile = document.getElementById('excelFile');
const sheetSelect = document.getElementById('sheetSelect');
const searchInput = document.getElementById('searchInput');
const statusText = document.getElementById('status');
const dataTable = document.getElementById('dataTable');

let workbook = null;
let currentRows = [];

excelFile.addEventListener('change', handleFileUpload);
sheetSelect.addEventListener('change', handleSheetChange);
searchInput.addEventListener('input', renderFilteredTable);

function handleFileUpload(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      workbook = XLSX.read(data, { type: 'array' });
      fillSheetOptions(workbook.SheetNames);
      statusText.textContent = `Planilha carregada: ${file.name}`;
      searchInput.disabled = false;
      sheetSelect.disabled = false;
    } catch {
      statusText.textContent = 'Erro ao ler a planilha. Verifique o arquivo enviado.';
      clearTable();
    }
  };
  reader.readAsArrayBuffer(file);
}

function fillSheetOptions(sheetNames) {
  sheetSelect.innerHTML = '';

  sheetNames.forEach((name, index) => {
    const option = document.createElement('option');
    option.value = name;
    option.textContent = name;
    if (index === 0) option.selected = true;
    sheetSelect.appendChild(option);
  });

  if (sheetNames.length > 0) {
    loadSheet(sheetNames[0]);
  }
}

function handleSheetChange() {
  if (!workbook) return;
  loadSheet(sheetSelect.value);
}

function loadSheet(sheetName) {
  const sheet = workbook.Sheets[sheetName];
  currentRows = XLSX.utils.sheet_to_json(sheet, { defval: '' });

  if (currentRows.length === 0) {
    statusText.textContent = `A aba "${sheetName}" não possui dados.`;
    clearTable();
    return;
  }

  statusText.textContent = `Exibindo aba "${sheetName}" com ${currentRows.length} linha(s).`;
  renderTable(currentRows);
}

function renderFilteredTable() {
  if (!currentRows.length) return;

  const query = searchInput.value.trim().toLowerCase();
  if (!query) {
    renderTable(currentRows);
    return;
  }

  const filteredRows = currentRows.filter((row) =>
    Object.values(row).some((value) => String(value).toLowerCase().includes(query))
  );

  statusText.textContent = `Filtro: ${filteredRows.length} resultado(s) encontrado(s).`;
  renderTable(filteredRows);
}

function renderTable(rows) {
  clearTable();
  if (!rows.length) return;

  const headers = Object.keys(rows[0]);

  const thead = document.createElement('thead');
  const headRow = document.createElement('tr');
  headers.forEach((header) => {
    const th = document.createElement('th');
    th.textContent = header;
    headRow.appendChild(th);
  });
  thead.appendChild(headRow);

  const tbody = document.createElement('tbody');
  rows.forEach((row) => {
    const tr = document.createElement('tr');
    headers.forEach((header) => {
      const td = document.createElement('td');
      td.textContent = row[header];
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  dataTable.append(thead, tbody);
}

function clearTable() {
  dataTable.innerHTML = '';
}
