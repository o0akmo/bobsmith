/* global Office, Excel, document */

const GEMINI_API_KEY = "AIzaSyDd93frsqmmwybBpjBvOIC4_5A9pOIDLhg";

// 🔥 Cache storage
let lastInput = "";
let lastResult = null;

async function askAI(prompt) {
  // ✅ CACHE HIT (no API call)
  if (prompt === lastInput && lastResult) {
    console.log("Cache hit - skipping API call");
    return lastResult;
  }

  try {
    const res = await fetch(
      "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" +
        GEMINI_API_KEY,
      {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          contents: [{ parts: [{ text: prompt }] }]
        })
      }
    );

    const data = await res.json();

    if (!res.ok) return "HTTP Error: " + res.status;
    if (data.error) return "API Error: " + data.error.message;

    const result =
      data.candidates?.[0]?.content?.parts?.[0]?.text || "No response";

    // 💾 STORE CACHE
    lastInput = prompt;
    lastResult = result;

    return result;
  } catch (e) {
    return "Fetch Error: " + e.message;
  }
}

Office.onReady(() => {
  const runBtn = document.getElementById("run");
  const appBody = document.getElementById("app-body");
  const loading = document.getElementById("sideload-msg");

  if (loading) loading.style.display = "none";
  if (appBody) appBody.style.display = "block";

  if (runBtn) runBtn.onclick = runAI;
});

async function runAI() {
  const output = document.getElementById("result");

  try {
    output.innerText = "Thinking...";

    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("values");

      await context.sync();

      const input = range.values
        .map(r => r.join(","))
        .join("\n");

      const result = await askAI(
        "Analyze this Excel data:\n\n" + input
      );

      output.innerText = result;
    });
  } catch (err) {
    output.innerText = "Error: " + err.message;
  }
}
