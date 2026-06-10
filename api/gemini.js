export default async function handler(req, res) {

  if (req.method !== 'POST') {
    return res.status(405).json({
      error: 'Método não permitido'
    });
  }

  try {

    const { payload } = req.body;

    const apiKey = process.env.GEMINI_API_KEY;
    const modelName = "gemini-2.0-flash";

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
}