/* global console, document, Excel, Office */

const GEMINI_API_KEY = "AIzaSyDd93frsqmmwybBpjBvOIC4_5A9pOIDLhg";

let isRunning = false;

// -------------------- AI CALL --------------------
async function askAI(prompt) {
  try {
    const res = await fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${GEMINI_API_KEY}`,
      {
        method: "POST",
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          contents: [
            {
              parts: [{ text: prompt }]
            }
          ]
        })
      }
    );

    const data = await res.json();

    // Handle errors cleanly
    if (!res.ok) {
      return `HTTP Error: ${res.status} - ${data?.error?.message || "Request failed"}`;
    }

    return (
      data.candidates?.[0]?.content?.parts?.[0]?.text ||
      "No response from AI"
    );

  } catch (err) {
    return "Fetch Error: " + err.message;
  }
}

// -------------------- INIT --------------------
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    const runBtn = document.getElementById("run");

    if (runBtn) {
      runBtn.onclick = run;
    }

    const resultBox = document.getElementById("result");
    if (!resultBox) {
      console.warn("Missing <div id='result'> in HTML");
    }
  }
});

// -------------------- MAIN RUN --------------------
export async function run() {
  if (isRunning) return; // prevent spam clicks
  isRunning = true;

  const resultBox = document.getElementById("result");

  try {
    if (resultBox) resultBox.innerText = "Thinking...";

    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("values");

      await context.sync();

      const values = range.values;

      // ---------------- VALIDATION ----------------
      if (!values || values.length === 0) {
        if (resultBox) resultBox.innerText = "Select some Excel cells first.";
        return;
      }

      const input = values
        .map(row => row.join(","))
        .join("\n")
        .trim();

      if (!input) {
        if (resultBox) resultBox.innerText = "Selected cells are empty.";
        return;
      }

      // ---------------- RATE LIMIT SAFETY ----------------
      await new Promise(r => setTimeout(r, 1200));

      // ---------------- AI CALL ----------------
      const result = await askAI(
        "You are an Excel AI assistant. Analyze this data and give insights, formulas, or patterns:\n\n" +
        input
      );

      if (resultBox) resultBox.innerText = result;
    });

  } catch (error) {
    console.error(error);
    if (resultBox) {
      resultBox.innerText = "Error: " + error.message;
    }
  }

  isRunning = false;
}
