
const express = require('express');
const cors = require('cors');
const { OpenAI } = require('openai');

require('dotenv').config();

const app = express();
app.use(cors());
app.use(express.json());

const openai = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });

app.post('/api/codificar', async (req, res) => {
  const { mensaje } = req.body;

  try {
    const completion = await openai.chat.completions.create({
      model: "gpt-3.5-turbo",
      messages: [{ role: "user", content: `Codifica quirúrgicamente este procedimiento: ${mensaje}` }],
    });

    const respuesta = completion.choices[0].message.content;
    res.json({ respuesta });
  } catch (error) {
    console.error(error);
    res.status(500).json({ error: "Error al generar respuesta de GPT" });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Servidor corriendo en puerto ${PORT}`));
