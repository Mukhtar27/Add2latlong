let selectedColumn = null;
let parsedData = [];

document.getElementById('browseBtn').addEventListener('click', () => {
  document.getElementById('fileInput').click();
});

document.getElementById('fileInput').addEventListener('change', handleFile);

const dropZone = document.getElementById('dropZone');
dropZone.addEventListener('dragover', e => {
  e.preventDefault();
  dropZone.classList.add('dragover');
});
dropZone.addEventListener('dragleave', () => {
  dropZone.classList.remove('dragover');
});
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('dragover');
  const file = e.dataTransfer.files[0];
  if (file) handleExcel(file);
});

function handleFile(event) {
  const file = event.target.files[0];
  if (file) handleExcel(file);
}

function handleExcel(file) {
  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    parsedData = XLSX.utils.sheet_to_json(sheet);
    if (parsedData.length > 0) {
      previewExcel(parsedData);
      populateColumnSelect(Object.keys(parsedData[0]));
    }
  };
  reader.readAsArrayBuffer(file);
}

function previewExcel(data) {
  const previewContainer = document.getElementById('previewContainer');
  previewContainer.innerHTML = '';

  const heading = document.createElement('h3');
  heading.textContent = '📄 Preview of uploaded file:';
  previewContainer.appendChild(heading);

  const table = document.createElement('table');
  table.classList.add('preview-table');

  const thead = document.createElement('thead');
  const tbody = document.createElement('tbody');

  const headers = Object.keys(data[0]);
  const headerRow = document.createElement('tr');

  headers.forEach(header => {
    const th = document.createElement('th');
    th.textContent = header;
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);

  data.slice(0, 5).forEach(row => {
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
  previewContainer.appendChild(table);
}

function populateColumnSelect(columns) {
  const select = document.getElementById('columnSelect');
  select.innerHTML = '';
  columns.forEach(col => {
    const option = document.createElement('option');
    option.value = col;
    option.textContent = col;
    select.appendChild(option);
  });
  document.getElementById('columnSelectContainer').style.display = 'block';
}

document.getElementById('processBtn').addEventListener('click', async () => {
  const apiKey = document.getElementById('apiKey').value.trim();
  const column = document.getElementById('columnSelect').value;
  if (!apiKey || !parsedData.length || !column) {
    showMessage('Please ensure API key, file, and column are selected.', true);
    return;
  }

  showMessage('Fetching coordinates...', false);
  const updatedData = [];
  for (const row of parsedData) {
    const address = row[column];
    try {
      const coords = await fetchCoordinates(address, apiKey);
      updatedData.push({ ...row, Latitude: coords.lat, Longitude: coords.lng });
    } catch (err) {
      updatedData.push({ ...row, Latitude: '', Longitude: '' });
    }
  }
  downloadExcel(updatedData);
  showMessage('✅ Coordinates fetched and file downloaded!');
});

function fetchCoordinates(address, apiKey) {
  const url = `https://maps.googleapis.com/maps/api/geocode/json?address=${encodeURIComponent(address)}&key=${apiKey}`;
  return fetch(url)
    .then(res => res.json())
    .then(data => {
      if (data.status === 'OK') return data.results[0].geometry.location;
      throw new Error('Not found');
    });
}

function downloadExcel(data) {
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Geocoded');
  const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });

  const blob = new Blob([wbout], { type: 'application/octet-stream' });
  const url = URL.createObjectURL(blob);

  const link = document.getElementById('downloadLink');
  link.href = url;
  link.download = 'geocoded_addresses.xlsx';
  link.style.display = 'inline-block';
  link.click();
}

function showMessage(message, isError = false) {
  const msgDiv = document.getElementById('message');
  msgDiv.textContent = message;
  msgDiv.style.color = isError ? 'red' : 'green';
}

document.getElementById('reset-btn').addEventListener('click', () => {
  location.reload();
});

function toggleAPIVisibility() {
  const input = document.getElementById('apiKey');
  const toggle = document.getElementById('toggleApi');
  if (input.type === 'password') {
    input.type = 'text';
    toggle.textContent = '🙈';
  } else {
    input.type = 'password';
    toggle.textContent = '👁️';
  }
}
