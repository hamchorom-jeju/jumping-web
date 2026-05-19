async function test() {
  const GAS_URL = "https://script.google.com/macros/s/AKfycbxaJ8W0pDK2Ig78BwE1necJm4FnA_kneBTukBtRyA7DtkLKBTFfq2lAWeMNpDY_oWZtYg/exec";
  try {
    console.log("Calling GAS Web App via fetch...");
    const response = await fetch(GAS_URL + "?action=getUserDashboardData", {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ phone: '01012345678' })
    });
    console.log("Status:", response.status);
    const text = await response.text();
    console.log("Response text (first 1000 chars):");
    console.log(text.substring(0, 1000));
  } catch (error) {
    console.error("Error occurred:", error.message);
  }
}

test();
