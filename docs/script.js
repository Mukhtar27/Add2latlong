// ===== Password Toggle (Custom SVG) =====
const apiKeyInput = document.getElementById('apiKey');
const togglePassword = document.getElementById('togglePassword');
const eyeIcon = document.getElementById('eyeIcon');

togglePassword.addEventListener('click', () => {
  const type = apiKeyInput.type === 'password' ? 'text' : 'password';
  apiKeyInput.type = type;

  eyeIcon.innerHTML = type === 'password'
    ? `<path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8S1 12 1 12z"/><circle cx="12" cy="12" r="3"/>`
    : `<path d="M17.94 17.94A10.94 10.94 0 0112 20c-7 0-11-8-11-8a21.91 21.91 0 014.14-5.94"/><path d="M1 1l22 22"/>`;
});

// ===== Excel Upload + Preview =====
const dropArea = document.getElementById('drop-area');
const fileElem = document.getElementById('fileElem');
const tablePreview = document.getElementById('tablePreview');
let workbookData = null;

dropArea.addEventListener('click', () => fileElem.click());

dropArea.addEventListener('dragover', e => {
  e.preventDefault();
  dropArea.classList.add('highlight');
});

dropArea.addEventListener('dragleave', () => {
  dropArea.classList.remove('highlight');
});

dropArea.addEventListener('drop', e => {
  e.preventDefault();
  dropArea.classList.remove('highlight');
  const file = e.dataTransfer.files[0];
  handleFile(file);
});

fileElem.addEventListener('change', () => {
  const file = fileElem.files[0];
  handleFile(file);
});

function handleFile(file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    const sheet = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: '' });
    workbookData = sheet;
    showTable(sheet);
    createAddressSelector(sheet);
  };
  reader.readAsArrayBuffer(file);
}

// ===== Show Table Preview =====
function showTable(data) {
  if (!data.length) return;
  const headers = Object.keys(data[0]);
  let html = '<table><thead><tr>';
  headers.forEach(h => html += `<th>${h}</th>`);
  html += '</tr></thead><tbody>';
  data.forEach(row => {
    html += '<tr>';
    headers.forEach(h => html += `<td>${row[h]}</td>`);
    html += '</tr>';
  });
  html += '</tbody></table>';
  tablePreview.innerHTML = html;
}

// ===== Address Column Selector =====
function createAddressSelector(data) {
  const headers = Object.keys(data[0]);
  let selectorHTML = `<div class="input-group"><label>Select Address Column</label><select id="addressColumn">`;
  headers.forEach(col => {
    selectorHTML += `<option value="${col}">${col}</option>`;
  });
  selectorHTML += `</select></div>`;

  if (!document.getElementById('addressColumn')) {
    dropArea.insertAdjacentHTML('afterend', selectorHTML);
    const goButton = `<button id="geocodeBtn" class="action-btn">Geocode</button>
                      <button id="downloadBtn" class="action-btn">Download .xlsx</button>
                      <button id="resetBtn" class="action-btn reset">Reset</button>`;
    tablePreview.insertAdjacentHTML('afterend', goButton);

    document.getElementById('geocodeBtn').addEventListener('click', geocodeAddresses);
    document.getElementById('downloadBtn').addEventListener('click', downloadExcel);
    document.getElementById('resetBtn').addEventListener('click', resetApp);
  }
}

// ===== Geocode Addresses =====
async function geocodeAddresses() {
  const apiKey = apiKeyInput.value.trim();
  if (!apiKey) return alert("Please enter your API key.");

  const col = document.getElementById('addressColumn').value;
  for (let i = 0; i < workbookData.length; i++) {
    const address = workbookData[i][col];
    try {
      const response = await fetch(`https://maps.googleapis.com/maps/api/geocode/json?address=${encodeURIComponent(address)}&key=${apiKey}`);
      const result = await response.json();
      if (result.status === "OK") {
        const location = result.results[0].geometry.location;
        workbookData[i]["Latitude"] = location.lat;
        workbookData[i]["Longitude"] = location.lng;
      } else {
        workbookData[i]["Latitude"] = "";
        workbookData[i]["Longitude"] = "";
      }
    } catch (err) {
      workbookData[i]["Latitude"] = "";
      workbookData[i]["Longitude"] = "";
    }
  }
  showTable(workbookData);
}

// ===== Download Updated Excel =====
function downloadExcel() {
  const ws = XLSX.utils.json_to_sheet(workbookData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Geocoded Data");
  XLSX.writeFile(wb, "geocoded_output.xlsx");
}

// ===== Reset Everything =====
function resetApp() {
  apiKeyInput.value = "";
  fileElem.value = "";
  workbookData = null;
  tablePreview.innerHTML = "";
  const addressSelect = document.getElementById('addressColumn');
  if (addressSelect) addressSelect.parentElement.remove();
  const buttons = document.querySelectorAll('.action-btn');
  buttons.forEach(btn => btn.remove());
}
