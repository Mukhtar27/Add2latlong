let selectedFile;
const fileInput = document.getElementById('fileInput');
const dropZone = document.getElementById('dropZone');
const browseBtn = document.getElementById('browseBtn');
const columnSelect = document.getElementById('columnSelect');
const columnSelectContainer = document.getElementById('columnSelectContainer');
const previewContainer = document.getElementById('previewContainer');
const message = document.getElementById('message');
const downloadLink = document.getElementById('downloadLink');

// Handle file browse
browseBtn.addEventListener('click', () => fileInput.click());

// Handle drag and drop
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
  handleFile(e.dataTransfer.files[0]);
});
fileInput.addEventListener('change', () => {
  if (fileInput.files.length > 0) {
    handleFile(fileInput.files[0]);
  }
});

function handleFile(file) {
  if (!file.name.endsWith('.xlsx')) {
    message.textContent = 'Please upload a valid .xlsx file.';
    return;
  }

  selectedFile = file;
  document.getElementById('fileInfo').textContent = `📁 Selected: ${file.name}`;

  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet);

    if (json.length === 0) return;

    // Show preview
    previewContainer.innerHTML = '<h4>📄 Preview:</h4><pre>' + JSON.stringify(json.slice(0, 5), null, 2) + '</pre>';

    // Populate column selector
    const keys = Object.keys(json[0]);
    columnSelect.innerHTML = '';
    keys.forEach(key => {
      const option = document.createElement('option');
      option.value = key;
      option.textContent = key;
      columnSelect.appendChild(option);
    });
    columnSelectContainer.style.display = 'block';
  };
  reader.readAsArrayBuffer(file);
}

document.getElementById('processBtn').addEventListener('click', async () => {
  if (!selectedFile) return alert('Please upload an Excel file.');

  const apiKey = document.getElementById('apiKey').value.trim();
  if (!apiKey) return alert('Please enter your API key.');

  const reader = new FileReader();
  reader.onload = async (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet);

    const column = columnSelect.value;
    message.textContent = '⏳ Fetching coordinates...';

    for (let row of json) {
      const address = row[column];
      const coords = await getCoordinates(address, apiKey);
      row.Latitude = coords.lat;
      row.Longitude = coords.lng;
    }

    message.textContent = '✅ Done!';

    const newSheet = XLSX.utils.json_to_sheet(json);
    const newWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWorkbook, newSheet, 'Geocoded');
    const wbout = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'array' });

    const blob = new Blob([wbout], { type: 'application/octet-stream' });
    downloadLink.href = URL.createObjectURL(blob);
    downloadLink.download = 'geocoded_addresses.xlsx';
    downloadLink.style.display = 'inline-block';
    downloadLink.textContent = '⬇ Download Geocoded Excel';
  };
  reader.readAsArrayBuffer(selectedFile);
});

async function getCoordinates(address, apiKey) {
  const url = `https://maps.googleapis.com/maps/api/geocode/json?address=${encodeURIComponent(address)}&key=${apiKey}`;
  try {
    const response = await fetch(url);
    const data = await response.json();
    if (data.status === 'OK') {
      const loc = data.results[0].geometry.location;
      return { lat: loc.lat, lng: loc.lng };
    }
  } catch (err) {
    console.error('Error fetching geocode:', err);
  }
  return { lat: '', lng: '' };
}

document.getElementById('reset-btn').addEventListener('click', () => {
  location.reload();
});

function toggleAPIVisibility() {
  const input = document.getElementById('apiKey');
  input.type = input.type === 'password' ? 'text' : 'password';
}
