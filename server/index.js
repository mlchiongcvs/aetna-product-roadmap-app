import express from "express";
import { fileURLToPath } from "url";
import { dirname, join } from "path";
import { readFileSync, writeFileSync } from "fs";
import { execSync } from "child_process";

const __dirname = dirname(fileURLToPath(import.meta.url));
const app = express();
const PORT = process.env.PORT || 8080;

const ANTHROPIC_BASE_URL = process.env.ANTHROPIC_BASE_URL || "https://hyperion-lms-api.prod.cvshealth.com";
const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY;
const CLIENT_NAME = process.env.X_CLIENT_NAME || "aetna-roadmap-app";

// Rally config
const RALLY_TOKEN = process.env.RALLY_TOKEN || "_DmfFV93JSCqP39zIM6h9SeOwpVnEy4odUH0U1D7k4";
const RALLY_BASE = "https://rally1.rallydev.com/slm/webservice/v2.0";
const RALLY_PROJECTS = {
  "MCAID: Train 9 - Member and Enrollment (MME)": "Train 9 - MME",
  "MCAID - MME: Enrollment Optimization": "Enrollment Optimization",
  "MBU - Enrollment/Member (E/M)": "MBU - E/M",
};
const PRODUCT_PREFIX_MAP = [
  ["AD9", "Assign a Doc"],
  ["EO9", "Enrollment Ops"],
  ["E09", "Enrollment Ops"],
  ["EX9", "Extracts & ID Cards"],
  ["DU9", "Duals/D-SNP"],
  ["EA9", "Enrollment Ancillary"],
];
const PRODUCT_PILLAR_MAP = {
  "Assign a Doc": "Growth",
  "Enrollment Ops": "Efficiency",
  "Extracts & ID Cards": "Efficiency",
  "Duals/D-SNP": "Growth",
  "Enrollment Ancillary": "Cost Reduction",
  "Other": "Enterprise Value",
};

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

// Rally features - serve cached data
app.get("/api/rally/features", (req, res) => {
  try {
    const data = readFileSync(join(__dirname, "features.json"), "utf-8");
    res.json(JSON.parse(data));
  } catch (err) {
    res.status(500).json({ error: "No Rally data available" });
  }
});

// Rally features - refresh from Rally API
app.post("/api/rally/refresh", async (req, res) => {
  try {
    const allFeatures = [];
    const allReleases = {};

    for (const [rallyName, displayName] of Object.entries(RALLY_PROJECTS)) {
      // Fetch releases
      const relQuery = encodeURIComponent(`(Project.Name = "${rallyName}")`);
      const relUrl = `${RALLY_BASE}/release?query=${relQuery}&fetch=Name,ReleaseDate,ReleaseStartDate&pagesize=200&order=ReleaseDate`;
      const relResult = execSync(`curl -s -H "ZSESSIONID: ${RALLY_TOKEN}" "${relUrl}"`, { encoding: "utf-8" });
      const relData = JSON.parse(relResult).QueryResult;
      for (const r of relData.Results) {
        if (r.ReleaseDate) allReleases[r.Name] = new Date(r.ReleaseDate);
      }

      // Fetch features
      let start = 1;
      while (true) {
        const query = encodeURIComponent(`(Project.Name = "${rallyName}")`);
        const fetchFields = "FormattedID,Name,State,Project,Release,PlannedEndDate,PlannedStartDate,ActualEndDate,ObjectID,c_ValueSector";
        const url = `${RALLY_BASE}/portfolioitem/feature?query=${query}&fetch=${fetchFields}&pagesize=200&start=${start}`;
        const result = execSync(`curl -s -H "ZSESSIONID: ${RALLY_TOKEN}" "${url}"`, { encoding: "utf-8" });
        const data = JSON.parse(result).QueryResult;

        for (const f of data.Results) {
          const state = f.State ? (typeof f.State === "object" ? f.State._refObjectName || "" : String(f.State)) : "";
          const release = f.Release ? (typeof f.Release === "object" ? f.Release._refObjectName || "" : String(f.Release)) : "";
          const actualEnd = f.ActualEndDate ? new Date(f.ActualEndDate) : null;
          const plannedEnd = f.PlannedEndDate ? new Date(f.PlannedEndDate) : null;

          const thirtyDaysAgo = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000);
          const isDone = state === "Done" || actualEnd !== null;
          if (isDone && (!actualEnd || actualEnd < thirtyDaysAgo)) continue;

          const name = f.Name || "";
          const upper = name.trim().toUpperCase();
          let product = "Other";
          for (const [prefix, prod] of PRODUCT_PREFIX_MAP) {
            if (upper.startsWith(prefix)) { product = prod; break; }
          }
          const pillar = PRODUCT_PILLAR_MAP[product] || "Enterprise Value";

          // Lane assignment (simplified - uses release name patterns)
          let lane;
          if (isDone) {
            lane = "done";
          } else if (state === "In Progress") {
            lane = "now";
          } else {
            // Use release end dates to determine current/next PI
            const releaseEnd = allReleases[release];
            const now = new Date();
            const sixMonths = new Date(now.getTime() + 182 * 24 * 60 * 60 * 1000);
            if (releaseEnd && releaseEnd >= now && releaseEnd <= new Date(now.getTime() + 90 * 24 * 60 * 60 * 1000)) {
              lane = "now";
            } else if (releaseEnd && releaseEnd > new Date(now.getTime() + 90 * 24 * 60 * 60 * 1000) && releaseEnd <= sixMonths) {
              lane = "next";
            } else {
              lane = "future";
            }
          }

          allFeatures.push({
            id: f.FormattedID || "",
            name,
            state: isDone ? "Done" : state,
            project: displayName,
            product,
            pillar,
            release,
            date: plannedEnd ? plannedEnd.toISOString().slice(0, 10) : "",
            lane,
            valueSector: f.c_ValueSector || "",
            rallyUrl: `https://rally1.rallydev.com/#/search?keywords=${f.FormattedID}`,
            oid: f.ObjectID || 0,
          });
        }

        if (start + 200 > data.TotalResultCount) break;
        start += 200;
      }
    }

    // Sort and save
    const laneOrder = { now: 0, next: 1, future: 2, done: 3 };
    allFeatures.sort((a, b) => (laneOrder[a.lane] ?? 9) - (laneOrder[b.lane] ?? 9) || a.id.localeCompare(b.id));

    writeFileSync(join(__dirname, "features.json"), JSON.stringify(allFeatures, null, 2));
    res.json({ success: true, count: allFeatures.length });
  } catch (err) {
    console.error("Rally refresh error:", err.message);
    res.status(500).json({ error: `Rally refresh failed: ${err.message}` });
  }
});

// SPA fallback
app.get("*", (req, res) => {
  res.sendFile(join(__dirname, "../dist/index.html"));
});

app.listen(PORT, () => {
  console.log(`Roadmap app running on port ${PORT}`);
});
