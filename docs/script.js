document.getElementById("geoForm").addEventListener("submit", async (e) => {
  e.preventDefault();
  const apiKey = document.getElementById("apiKey").value;
  const address = document.getElementById("address").value;
  const resultDiv = document.getElementById("result");

  const url = `https://maps.googleapis.com/maps/api/geocode/json?address=${encodeURIComponent(address)}&key=${apiKey}`;

  resultDiv.innerHTML = "🔍 Searching...";

  try {
    const response = await fetch(url);
    const data = await response.json();

    if (data.status === "OK") {
      const location = data.results[0].geometry.location;
      resultDiv.innerHTML = `
        ✅ Coordinates: <strong>${location.lat}, ${location.lng}</strong><br/>
        📫 Verified Address: ${data.results[0].formatted_address}
      `;
    } else {
      resultDiv.innerHTML = `❌ Error: ${data.status}`;
    }
  } catch (error) {
    resultDiv.innerHTML = `❌ Fetch failed: ${error}`;
  }
});
