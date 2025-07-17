const fileElem = document.getElementById('fileElem');
const dropArea = document.getElementById('drop-area');
const tablePreview = document.getElementById('tablePreview');
const apiKeyInput = document.getElementById('apiKey');
const togglePassword = document.getElementById('togglePassword');

// Password eye toggle
togglePassword.addEventListener('click', () => {
  const type = apiKeyInput.type === 'password' ? 'text' : 'password';
  apiKeyInput.type = type;
  togglePassword.textContent = type === 'password' ? '👁️' : '🙈';
});

// File drag/drop logic
dropArea.addEventListener('click', () => fileElem.click());

['dragenter', 'dragover'].forEach(evt =>
  dropArea.addEventListener(evt, e => {
    e.preventDefault();
    dropArea.classList.add('highlight');
  })
);

['dragleave', 'drop'].forEach(evt =>
  dropArea.addEventListener(evt, e => {
    e.preventDefault();
    dropArea.classList.remove('highlight');
  })
);

dropArea.addEventListener('drop', e => {
  const file = e.dataTransfer.files[0];
  if (file) handleFile(file);
});

fileElem.addEventListener('change', e => {
  const file = e.target.files[0];
  if (file) handleFile(file);
});

function handleFile(file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });
    renderTable(json);
  };
  reader.readAsArrayBuffer(file);
}

function renderTable(data) {
  if (!data.length) {
    tablePreview.innerHTML = "<p>No data found.</p>";
    return;
  }

  const table = document.createElement('table');
  const thead = document.createElement('thead');
  const tbody = document.createElement('tbody');

  const headers = Object.keys(data[0]);
  const headRow = document.createElement('tr');
  headers.forEach(header => {
    const th = document.createElement('th');
    th.textContent = header;
    headRow.appendChild(th);
  });
  thead.appendChild(headRow);

  data.forEach(row => {
    const tr = document.createElement('tr');
    headers.forEach(header => {
      const td = document.createElement('td');
      td.textContent = row[header];
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  table.appendChild(thead);
  table.appendChild(tbody);
  tablePreview.innerHTML = "";
  tablePreview.appendChild(table);
}
