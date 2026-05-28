import express from "express";
import { fileURLToPath } from "url";
import { dirname, join } from "path";

const __dirname = dirname(fileURLToPath(import.meta.url));
const app = express();
const PORT = process.env.PORT || 8080;

const ANTHROPIC_BASE_URL = process.env.ANTHROPIC_BASE_URL || "https://hyperion-lms-api.prod.cvshealth.com";
const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY;
const CLIENT_NAME = process.env.X_CLIENT_NAME || "aetna-roadmap-app";

app.use(express.json({ limit: "25mb" }));

// Health check (required by CAP)
app.get("/health", (req, res) => res.status(200).json({ status: "ok" }));

// Serve static frontend
app.use(express.static(join(__dirname, "../dist")));

// Proxy endpoint for Claude API calls
app.post("/api/messages", async (req, res) => {
  if (!ANTHROPIC_API_KEY) {
    return res.status(500).json({ error: { message: "ANTHROPIC_API_KEY not configured on server" } });
  }

  try {
    const response = await fetch(`${ANTHROPIC_BASE_URL}/v1/messages`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": ANTHROPIC_API_KEY,
        "anthropic-version": "2023-06-01",
        "X-Client-Name": CLIENT_NAME,
      },
      body: JSON.stringify(req.body),
    });

    const data = await response.json();
    res.status(response.status).json(data);
  } catch (err) {
    console.error("Proxy error:", err.message);
    res.status(502).json({ error: { message: `Proxy error: ${err.message}` } });
  }
});

// SPA fallback
app.get("*", (req, res) => {
  res.sendFile(join(__dirname, "../dist/index.html"));
});

app.listen(PORT, () => {
  console.log(`Roadmap app running on port ${PORT}`);
});
