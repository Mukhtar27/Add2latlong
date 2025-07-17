const dropArea = document.getElementById('drop-area');
const fileElem = document.getElementById('fileElem');
const tablePreview = document.getElementById('tablePreview');
const togglePassword = document.getElementById('togglePassword');
const apiKeyInput = document.getElementById('apiKey');

// Password toggle icon logic
togglePassword.addEventListener('click', () => {
  const type = apiKeyInput.getAttribute('type') === 'password' ? 'text' : 'password';
  apiKeyInput.setAttribute('type', type);
  togglePassword.textContent = type === 'password' ? '👁️' : '🙈';
});

// Drag-and-drop handlers
['dragenter', 'dragover'].forEach(eventName => {
  dropArea.addEventListener(eventName, e => {
    e.preventDefault();
    dropArea.classList.add('highlight');
  });
});

['dragleave', 'drop'].forEach(eventName => {
  dropArea.addEventListener(eventName, e => {
    e.preventDefault();
    dropArea.classList.remove('highlight');
  });
});

dropArea.addEventListener('click', () => fileElem.click());

dropArea.addEventListener('drop', e => {
  const file = e.dataTransfer.files[0];
  handleFile(file);
});

fileElem.addEventListener('change', e => {
  const file = e.target.files[0];
  handleFile(file);
});

function handleFile(file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const json = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
    renderTable(json);
  };
  reader.readAsArrayBuffer(file);
}

function renderTable(data) {
  if (data.length === 0) {
    tablePreview.innerHTML = "<p>No data to preview.</p>";
    return;
  }

  const table = document.createElement('table');
  const thead = document.createElement('thead');
  const headerRow = document.createElement('tr');

  Object.keys(data[0]).forEach(key => {
    const th = document.createElement('th');
    th.textContent = key;
    headerRow.appendChild(th);
  });

  thead.appendChild(headerRow);
  table.appendChild(thead);

  const tbody = document.createElement('tbody');
  data.forEach(row => {
    const tr = document.createElement('tr');
    Object.values(row).forEach(cell => {
      const td = document.createElement('td');
      td.textContent = cell;
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  table.appendChild(tbody);
  tablePreview.innerHTML = '';
  tablePreview.appendChild(table);
}
