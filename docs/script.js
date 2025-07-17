let workbook, sheetData, header, selectedColumn;

document.getElementById("excelFile").addEventListener("change", handleFile, false);
document.getElementById("processBtn").addEventListener("click", processAddresses);

function handleFile(e) {
  const file = e.target.files[0];
  const reader = new FileReader();

  reader.onload = function (event) {
    const data = new Uint8Array(event.target.result);
    workbook = XLSX.read(data, { type: "array" });

    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    sheetData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

    if (sheetData.length > 0) {
      header = sheetData[0];
      const select = document.getElementById("columnSelect");
      select.innerHTML = "";

      header.forEach((col, i) => {
        const option = document.createElement("option");
        option.value = i;
        option.textContent = col;
        select.appendChild(option);
      });

      document.getElementById("columnSelectWrapper").style.display = "block";
      document.getElementById("processBtn").disabled = false;
    }
  };

  reader.readAsArrayBuffer(file);
}

async function processAddresses() {
  const apiKey = document.getElementById("apiKey").value.trim();
  const colIndex = parseInt(document.getElementById("columnSelect").value);
  const statusDiv = document.getElementById("status");

  if (!apiKey) {
    alert("Please enter your API key.");
    return;
  }

  statusDiv.innerHTML = "⏳ Processing... Please wait.";

  const results = [header.concat(["Latitude", "Longitude"])];

  for (let i = 1; i < sheetData.length; i++) {
    const row = sheetData[i];
    const address = row[colIndex];
    let lat = "", lon = "";

    if (address) {
      try {
        const response = await fetch(
          `https://maps.googleapis.com/maps/api/geocode/json?address=${encodeURIComponent(address)}&key=${apiKey}`
        );
        const data = await response.json();

        if (data.status === "OK") {
          lat = data.results[0].geometry.location.lat;
          lon = data.results[0].geometry.location.lng;
        }
      } catch (error) {
        console.error("Error fetching coordinates:", error);
      }
    }

    results.push(row.concat([lat, lon]));
  }

  const newSheet = XLSX.utils.aoa_to_sheet(results);
  const newWB = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(newWB, newSheet, "Geocoded");

  const wbout = XLSX.write(newWB, { bookType: "xlsx", type: "array" });
  const blob = new Blob([wbout], { type: "application/octet-stream" });
  const url = URL.createObjectURL(blob);

  const link = document.getElementById("downloadLink");
  link.href = url;
  link.download = "geocoded_addresses.xlsx";
  link.style.display = "block";
  link.textContent = "⬇ Download Geocoded Excel";

  statusDiv.innerHTML = "✅ Done! You can now download the results.";
}
