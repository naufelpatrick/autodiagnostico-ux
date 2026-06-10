/*export default async function handler(req, res) {

  if (req.method !== 'POST') {
    return res.status(405).json({
      error: 'Método não permitido'
    });
  }

  try {

    const { payload } = req.body;

    const apiKey = process.env.GEMINI_API_KEY;
    const modelName = "gemini-2.0-flash-lite";

    console.log("MODEL USADO:", modelName);
    console.log("CHAVE CONFIGURADA:", !!apiKey);

    const response = await fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${apiKey}`,
      {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(payload)
      }
    );

    const data = await response.json();
    console.log(JSON.stringify(data, null, 2));

    if (!response.ok) {
      return res.status(response.status).json(data);
    }

    return res.status(200).json(data);

  } catch (error) {

    console.error("Erro Gemini:", error);

    return res.status(500).json({
      error: {
        message: "Erro interno ao processar requisição Gemini."
      }
    });
  }
}*/


export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({
      error: {
        message: "Método não permitido."
      }
    });
  }

  try {
    const apiKey = process.env.OPENROUTER_API_KEY;

    if (!apiKey) {
      return res.status(500).json({
        error: {
          message: "Chave OpenRouter não configurada."
        }
      });
    }

    const { payload } = req.body;

    const userPrompt = payload?.contents?.[0]?.parts?.[0]?.text || "";
    const systemPrompt = payload?.systemInstruction?.parts?.[0]?.text || "";

    const response = await fetch("https://openrouter.ai/api/v1/chat/completions", {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${apiKey}`,
        "Content-Type": "application/json",
        "HTTP-Referer": "https://autodiagnostico-ux.vercel.app",
        "X-Title": "Autodiagnostico UX"
      },
      body: JSON.stringify({
        model: "deepseek/deepseek-chat-v3-0324",
        messages: [
          {
            role: "system",
            content: systemPrompt
          },
          {
            role: "user",
            content: userPrompt
          }
        ],
        temperature: 0.4
      })
    });

    const data = await response.json();

    if (!response.ok) {
      return res.status(response.status).json(data);
    }

    const generatedText = data?.choices?.[0]?.message?.content;

    if (!generatedText) {
      return res.status(500).json({
        error: {
          message: "OpenRouter não retornou conteúdo."
        }
      });
    }

    return res.status(200).json({
      candidates: [
        {
          content: {
            parts: [
              {
                text: generatedText
              }
            ]
          }
        }
      ]
    });

  } catch (error) {
    console.error("Erro OpenRouter:", error);

    return res.status(500).json({
      error: {
        message: "Erro interno ao processar requisição via OpenRouter."
      }
    });
  }
}