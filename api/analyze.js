const schema = {
  type: "object",
  additionalProperties: false,
  required: ["score", "findings"],
  properties: {
    score: { type: "integer", minimum: 0, maximum: 100 },
    findings: {
      type: "array",
      maxItems: 30,
      items: {
        type: "object",
        additionalProperties: false,
        required: [
          "step",
          "severity",
          "title",
          "area",
          "detail",
          "recommendation",
          "evidence",
        ],
        properties: {
          step: { type: "integer", minimum: 1 },
          severity: { type: "string", enum: ["Critical", "Major", "Minor"] },
          title: { type: "string" },
          area: { type: "string" },
          detail: { type: "string" },
          recommendation: { type: "string" },
          evidence: { type: "string" },
        },
      },
    },
  },
};

const systemPrompt = `Ты — старший CX/UX-аудитор. Анализируй только то, что действительно видно на переданных экранах. Экраны идут в порядке клиентского пути.
Проверяй понятность следующего действия, навигацию, информационную архитектуру, визуальную иерархию, читаемость, доступность, предотвращение и обработку ошибок, обратную связь системы, консистентность, доверие и непрерывность переходов.
Не выдумывай скрытые состояния, бизнес-правила или отсутствующие переходы. Каждое замечание обязано содержать конкретное визуальное доказательство в evidence. Если проблему нельзя уверенно подтвердить изображением, не включай её.
Critical — сценарий невозможно завершить, есть риск потери денег/данных, необратимое действие без защиты или серьёзная доступностная блокировка. Блокирует релиз.
Major — задача выполнима, но велик риск ошибки, отказа или существенного затруднения. Блокирует релиз.
Minor — локальная проблема качества, не мешающая завершить задачу. Релиз допустим.
Пиши по-русски, конкретно и без общих фраз. Не создавай замечания ради количества. Если недостатков нет, верни пустой массив. Score: 90–100 отлично, 75–89 хорошо, 60–74 проблемно, ниже 60 — высокий риск.`;

export default async function handler(req, res) {
  if (req.method !== "POST")
    return res.status(405).json({ error: "Метод не поддерживается" });
  if (!process.env.OPENAI_API_KEY)
    return res
      .status(500)
      .json({ error: "На сервере не настроен OPENAI_API_KEY" });
  const images = req.body?.images;
  if (!Array.isArray(images) || !images.length || images.length > 8)
    return res.status(400).json({ error: "Передайте от 1 до 8 изображений" });
  if (
    images.some(
      (image) => typeof image !== "string" || !image.startsWith("data:image/"),
    )
  )
    return res.status(400).json({ error: "Некорректный формат изображения" });
  try {
    const content = [
      {
        type: "input_text",
        text: `Проанализируй путь из ${images.length} экранов. Номер изображения — номер шага.`,
      },
    ];
    images.forEach((image, index) =>
      content.push(
        { type: "input_text", text: `Шаг ${index + 1}` },
        { type: "input_image", image_url: image, detail: "high" },
      ),
    );
    const openaiResponse = await fetch("https://api.openai.com/v1/responses", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${process.env.OPENAI_API_KEY}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        model: process.env.OPENAI_MODEL || "gpt-5.6-sol",
        input: [
          {
            role: "system",
            content: [{ type: "input_text", text: systemPrompt }],
          },
          { role: "user", content },
        ],
        text: {
          format: {
            type: "json_schema",
            name: "cx_review",
            strict: true,
            schema,
          },
        },
        max_output_tokens: 12000,
      }),
    });
    const response = await openaiResponse.json();
    if (!openaiResponse.ok)
      return res
        .status(openaiResponse.status)
        .json({ error: response?.error?.message || "Ошибка OpenAI API" });
    const outputText = response.output
      ?.flatMap((item) => item.content || [])
      .find((item) => item.type === "output_text")?.text;
    if (!outputText)
      return res
        .status(502)
        .json({ error: "Модель не вернула результат анализа" });
    const result = JSON.parse(outputText);
    result.findings = result.findings.filter(
      (finding) => finding.step <= images.length,
    );
    return res.status(200).json(result);
  } catch (error) {
    console.error("Analysis failed", error);
    return res
      .status(500)
      .json({ error: "Не удалось выполнить анализ. Попробуйте ещё раз." });
  }
}
