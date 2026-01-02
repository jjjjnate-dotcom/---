exports.handler = async (event) => {
  if (event.httpMethod !== "POST") {
    return { statusCode: 405, body: "Method Not Allowed" };
  }

  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) {
    return { statusCode: 500, body: "Missing OPENAI_API_KEY" };
  }

  let payload = {};
  try {
    payload = JSON.parse(event.body || "{}");
  } catch {
    return { statusCode: 400, body: "Invalid JSON" };
  }

  const { mode, base64, prompt } = payload;

  if (!mode) {
    return { statusCode: 400, body: "mode is required" };
  }

  try {
    let messages = [];
    if (mode === "ocr") {
      if (!base64) return { statusCode: 400, body: "base64 is required for ocr" };
      messages = [
        { role: "system", content: "???대?吏瑜?JSON?쇰줈 ?뺥솗?섍쾶 異붿텧?섎뒗 ?꾩슦誘몄엯?덈떎." },
        {
          role: "user",
          content: [
            {
              type: "text",
              text:
                prompt ||
                "???대?吏瑜?蹂닿퀬 1?쒖쐞(泥??? ?낆껜 ?뺣낫瑜?JSON?쇰줈 異쒕젰?섏꽭?? ?꾨뱶: company, bizno, ceo, contact, amount, reason, period, address.",
            },
            { type: "image_url", image_url: { url: base64 } },
          ],
        },
      ];
    } else if (mode === "genBody") {
      if (!prompt) return { statusCode: 400, body: "prompt is required for genBody" };
      messages = [
        { role: "system", content: "?꾪뙆??怨듭?臾?蹂몃Ц???묒꽦?섎뒗 ?꾩슦誘몄엯?덈떎." },
        { role: "user", content: prompt },
      ];
    } else {
      return { statusCode: 400, body: "unsupported mode" };
    }

    const requestBody = {
      model: "gpt-4o",
      messages,
      max_tokens: 800,
      temperature: 0.2,
    };
    if (mode === "ocr") {
      requestBody.response_format = { type: "json_object" };
    }
    const res = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${apiKey}`,
      },
      body: JSON.stringify(requestBody),
    });

    if (!res.ok) {
      const errText = await res.text();
      return { statusCode: 500, body: `OpenAI error: ${res.status} ${errText}` };
    }

    const data = await res.json();
    const content = data.choices?.[0]?.message?.content;
    if (!content) {
      return { statusCode: 500, body: "No content from OpenAI" };
    }

    return {
      statusCode: 200,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ content }),
    };
  } catch (err) {
    return { statusCode: 500, body: `Server error: ${err.message}` };
  }
};

