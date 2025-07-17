let workbookData = null;

document.getElementById("fileUpload").addEventListener("change", handleFile, false);

function toggleApiKey() {
  const input = document.getElementById("apiKey");
  input.type = input.type === "password" ? "text" : "password";
}

function handleFile(e) {
  const file = e.target.files[0];
  const reader = new FileReader();

  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    workbookData = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    const columns = Object.keys(workbookData[0]);
    const select = document.getElementById("columnSelect");
    select.innerHTML = columns.map(col => `<option value="${col}">${col}</option>`).join('');

    previewTable(workbookData.slice(0, 10));
  };

  reader.readAsArrayBuffer(file);
}

function previewTable(data) {
  const table = document.getElementById("previewTable");
  if (!data || data.length === 0) {
    table.innerHTML = "";
    return;
  }

  const headers = Object.keys(data[0]);
  const thead = `<thead><tr>${headers.map(h => `<th>${h}</th>`).join('')}</tr></thead>`;
  const tbody = `<tbody>${data.map(row =>
    `<tr>${headers.map(h => `<td>${row[h]}</td>`).join('')}</tr>`
  ).join('')}</tbody>`;
  table.innerHTML = thead + tbody;
}

async function geocodeAddress(address, apiKey) {
  const url = `https://maps.googleapis.com/maps/api/geocode/json?address=${encodeURIComponent(address)}&key=${apiKey}`;
  try {
    const res = await fetch(url);
    const data = await res.json();
    if (data.status === "OK") {
      const loc = data.results[0].geometry.location;
      return [loc.lat, loc.lng];
    }
  } catch (e) {
    console.error("Geocode error:", e);
  }
  return [null, null];
}

async function processFile() {
  const apiKey = document.getElementById("apiKey").value.trim();
  const addressColumn = document.getElementById("columnSelect").value;
  const output = document.getElementById("output");
  const downloadLink = document.getElementById("downloadLink");

  if (!apiKey || !workbookData || !addressColumn) {
    output.textContent = "Please fill all fields!";
    output.className = "alert alert-danger";
    output.classList.remove("d-none");
    return;
  }

  output.className = "alert alert-warning";
  output.classList.remove("d-none");
  output.textContent = "⏳ Fetching coordinates...";

  for (let row of workbookData) {
    const address = row[addressColumn];
    const [lat, lon] = await geocodeAddress(address, apiKey);
    row["Latitude"] = lat;
    row["Longitude"] = lon;
  }

  output.className = "alert success";
  output.classList.add("success");
  output.textContent = "✅ Coordinates added successfully!";

  previewTable(workbookData.slice(0, 10));

  const ws = XLSX.utils.json_to_sheet(workbookData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Geocoded");

  const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([wbout], { type: "application/octet-stream" });

  downloadLink.href = URL.createObjectURL(blob);
  downloadLink.classList.remove("d-none");
}
