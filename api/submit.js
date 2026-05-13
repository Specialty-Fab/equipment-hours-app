const POWER_AUTOMATE_URL =
  "https://defaulta1dce605051e42ce9ba7342cabd36c.67.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/883f9b78868849b2ab5ebb5a7727a4ac/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=SfrBLRZrx8AiL4yrjPuFJHtd72FaLdKvALwM9Oga4ac";

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Only POST requests are allowed." });
  }

  try {
    const response = await fetch(POWER_AUTOMATE_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(req.body),
    });

    if (!response.ok) {
      const errorText = await response.text();
      return res.status(response.status).json({
        error: `Power Automate failed: ${response.status} ${errorText}`,
      });
    }

    return res.status(200).json({ ok: true });
  } catch (error) {
    return res.status(500).json({
      error: error.message || "Unknown server error.",
    });
  }
}
