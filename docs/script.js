document.getElementById("browseBtn").addEventListener("click", () => {
  document.getElementById("fileInput").click();
});

document.getElementById("dropZone").addEventListener("dragover", (e) => {
  e.preventDefault();
  e.currentTarget.classList.add("hover");
});

document.getElementById("dropZone").addEventListener("dragleave", (e) => {
  e.currentTarget.classList.remove("hover");
});

document.getElementById("dropZone").addEventListener("drop", (e) => {
  e.preventDefault();
  e.currentTarget.classList.remove("hover");
  const file = e.dataTransfer.files[0];
  handleFile(file);
});

document.getElementById("fileInput").addEventListener("change", (e) => {
  const file = e.target.files[0];
  handleFile(file);
});

function toggleAPIVisibility() {
  const apiKeyInput = document.getElementById("apiKey");
  if (apiKeyInput.type === "password") {
    apiKeyInput.type = "text";
  } else {
    apiKeyInput.type = "password";
  }
}

function handleFile(file) {
  if (!file || !file.name.endsWith(".xlsx")) return;
  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    if (json.length === 0) return;

    displayPreviewTable(json);
    populateColumnSelect(Object.keys(json[0]));
  };
  reader.readAsArrayBuffer(file);
}

function displayPreviewTable(data) {
  const container = document.getElementById("previewContainer");
  container.innerHTML = `<label>Preview of uploaded file:</label>`;
  const table = document.createElement("table");

  const thead = document.createElement("thead");
  const headerRow = document.createElement("tr");
  const headers = Object.keys(data[0]);

  headerRow.innerHTML = `<th>#</th>` + headers.map(h => `<th>${h}</th>`).join("");
  thead.appendChild(headerRow);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  data.slice(0, 5).forEach((row, idx) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>${idx}</td>` + headers.map(h => `<td>${row[h]}</td>`).join("");
    tbody.appendChild(tr);
  });

  table.appendChild(tbody);
  container.appendChild(table);
  document.getElementById("columnSelectContainer").style.display = "block";
}

function populateColumnSelect(columns) {
  const select = document.getElementById("columnSelect");
  select.innerHTML = columns.map(col => `<option value="${col}">${col}</option>`).join("");
}

document.getElementById("reset-btn").addEventListener("click", () => {
  location.reload();
});
