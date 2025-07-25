
const API_KEY = "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";
const chatBox = document.getElementById("chatBox");
const chatForm = document.getElementById("chatForm");
const userInput = document.getElementById("userInput");

chatForm.addEventListener("submit", async (e) => {
  e.preventDefault();
  const input = userInput.value.trim();
  if (!input) return;

  appendMessage("user", input);
  userInput.value = "";

  const reply = await getBotReply(input);
  appendMessage("bot", reply);
});

function appendMessage(sender, text) {
  const message = document.createElement("div");
  message.classList.add(sender === "bot" ? "bot-message" : "user-message");
  message.textContent = text;
  chatBox.appendChild(message);
  chatBox.scrollTop = chatBox.scrollHeight;
}

async function getBotReply(message) {
  try {
    const res = await fetch("https://openrouter.ai/api/v1/chat/completions", {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${API_KEY}`,
        "Content-Type": "application/json",
        "HTTP-Referer": "https://chat.openai.com/",
        "X-Title": "Codificador Quirúrgico"
      },
      body: JSON.stringify({
        model: "openrouter/gpt-3.5-turbo",
        messages: [
          {
            role: "system",
            content: "Sos un asistente médico especializado en codificación quirúrgica. Tu tarea es identificar el código correcto de un nomenclador quirúrgico a partir de una descripción natural del procedimiento. Respondé solo con el código, la descripción y el nivel de complejidad si lo conocés."
          },
          {
            role: "user",
            content: message
          }
        ]
      })
    });

    const data = await res.json();
    return data.choices[0].message.content;
  } catch (err) {
    console.error(err);
    return "⚠️ Error al conectar con el modelo. Verificá tu conexión o clave.";
  }
}
