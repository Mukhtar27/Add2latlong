let workbook, sheetName, data;

document.getElementById('fileInput')?.addEventListener('change', handleFile);
document.getElementById('processBtn').addEventListener('click', processAddresses);
document.getElementById('reset-btn').addEventListener('click', () => location.reload());

function toggleAPIVisibility() {
  const apiInput = document.getElementById('apiKey');
  apiInput.type = apiInput.type === 'password' ? 'text' : 'password';
}

// Drag-and-drop support
const dropZone = document.querySelector(".drop-zone");
if (dropZone) {
  dropZone.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropZone.style.borderColor = "#007bff";
  });

  dropZone.addEventListener("dragleave", () => {
    dropZone.style.borderColor = "#ccc";
  });

  dropZone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropZone.style.borderColor = "#ccc";
    const file = e.dataTransfer.files[0];
    if (file && file.name.endsWith(".xlsx")) {
      document.getElementById("fileInput").files = e.dataTransfer.files;
      handleFile({ target: { files: [file] } });
    } else {
      alert("Please upload a valid .xlsx file.");
    }
  });

  // Trigger file browse on button click
  document.getElementById("browseBtn")?.addEventListener("click", () => {
    document.getElementById("fileInput").click();
  });
}

function handleFile(event) {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (e) => {
    const dataBinary = new Uint8Array(e.target.result);
    workbook = XLSX.read(dataBinary, { type: "array" });
    sheetName = workbook.SheetNames[0];
    data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

    if (data.length === 0) {
      document.getElementById('message').textContent = "Excel file is empty.";
      return;
    }

    const columns = Object.keys(data[0]);
    const columnSelect = document.getElementById('columnSelect');
    columnSelect.innerHTML = "";
    columns.forEach(col => {
      const option = document.createElement("option");
      option.value = col;
      option.textContent = col;
      columnSelect.appendChild(option);
    });

    document.getElementById('columnSelectContainer').style.display = 'block';
    document.getElementById('message').textContent = `Loaded ${data.length} rows.`;
  };

  reader.readAsArrayBuffer(file);
}

async function processAddresses() {
  const apiKey = document.getElementById('apiKey').value;
  const column = document.getElementById('columnSelect').value;

  if (!apiKey || !data || !column) {
    document.getElementById('message').textContent = "Missing required input.";
    return;
  }

  document.getElementById('message').textContent = "Fetching coordinates...";
  for (let i = 0; i < data.length; i++) {
    const address = data[i][column];
    if (!address) continue;

    try {
      const response = await fetch(`https://maps.googleapis.com/maps/api/geocode/json?address=${encodeURIComponent(address)}&key=${apiKey}`);
      const json = await response.json();

      if (json.status === "OK" && json.results[0]) {
        const loc = json.results[0].geometry.location;
        data[i]["Latitude"] = loc.lat;
        data[i]["Longitude"] = loc.lng;
      } else {
        data[i]["Latitude"] = "";
        data[i]["Longitude"] = "";
      }
    } catch (error) {
      console.error("Error fetching geocode:", error);
      data[i]["Latitude"] = "";
      data[i]["Longitude"] = "";
    }
  }

  const worksheet = XLSX.utils.json_to_sheet(data);
  const newWorkbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(newWorkbook, worksheet, "Geocoded");

  const blob = XLSX.write(newWorkbook, { bookType: "xlsx", type: "blob" });
  const url = URL.createObjectURL(blob);
  const downloadLink = document.getElementById("downloadLink");
  downloadLink.href = url;
  downloadLink.download = "Geocoded_Addresses.xlsx";
  downloadLink.style.display = "inline-block";
  downloadLink.textContent = "⬇ Download Excel";
  document.getElementById('message').textContent = "Geocoding complete.";
}
