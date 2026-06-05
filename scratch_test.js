async function test() {
  const GAS_URL = "https://script.google.com/macros/s/AKfycbxaJ8W0pDK2Ig78BwE1necJm4FnA_kneBTukBtRyA7DtkLKBTFfq2lAWeMNpDY_oWZtYg/exec";
  try {
    console.log("Calling searchMemberByPin for '0689'...");
    const response = await fetch(GAS_URL + "?action=searchMemberByPin", {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ args: ["0689"] })
    });
    console.log("Status:", response.status);
    const result = await response.json();
    console.log("Result:", JSON.stringify(result, null, 2));
  } catch (error) {
    console.error("Error occurred:", error.message);
  }
}

test();


