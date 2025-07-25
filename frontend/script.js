
async function sendMessage() {
  const input = document.getElementById("input");
  const chat = document.getElementById("chat");
  const userMsg = input.value;
  if (!userMsg) return;

  const userDiv = document.createElement("div");
  userDiv.className = "msg user";
  userDiv.textContent = userMsg;
  chat.appendChild(userDiv);
  input.value = "";

  const botDiv = document.createElement("div");
  botDiv.className = "msg bot";
  botDiv.textContent = "Pensando...";
  chat.appendChild(botDiv);

  try {
    const response = await fetch("https://your-api-url.onrender.com/api/codificar", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ mensaje: userMsg })
    });
    const data = await response.json();
    botDiv.textContent = data.respuesta;
  } catch (error) {
    botDiv.textContent = "Ocurrió un error al contactar al servidor.";
  }
}
