/* global Office, Excel, document */
// 🔥 Cache storage
let lastInput = "";
let lastResult = null;

const GROQ_API_KEY = "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";

async function askAI(prompt) {
  try {
    const res = await fetch("https://api.groq.com/openai/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": "Bearer " + GROQ_API_KEY
      },
      body: JSON.stringify({
        model: "llama-3.1-8b-instant",
        messages: [
          {
            role: "user",
            content: prompt
          }
        ],
        temperature: 0.7
      })
    });

    const data = await res.json();

    console.log("Groq response:", data);

    if (!res.ok) {
      return "HTTP Error: " + res.status + " - " + (data.error?.message || "");
    }

    return data.choices?.[0]?.message?.content || "No response from AI";
  } catch (err) {
    return "Fetch Error: " + err.message;
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
