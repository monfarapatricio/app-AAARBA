async function getBotReply(message) {
  const response = await fetch("https://openrouter.ai/api/v1/chat/completions", {
    method: "POST",
    headers: {
      "Authorization": "Bearer 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",
      "Content-Type": "application/json",
      "HTTP-Referer": "https://app-AAAARBA.onrender.com",  // O el dominio real de tu app
      "X-Title": "Codificador Quirúrgico"
    },
    body: JSON.stringify({
      model: "openai/gpt-3.5-turbo",
      messages: [
        { role: "system", content: "Actuá como un asistente que ayuda a codificar procedimientos quirúrgicos según el nomenclador médico argentino." },
        { role: "user", content: message }
      ]
    })
  });

  if (!response.ok) {
    throw new Error("⚠️ Error al conectar con el modelo. Verificá tu clave.");
  }

  const data = await response.json();

  const reply = data.choices[0].message.content;
  return reply;
}
