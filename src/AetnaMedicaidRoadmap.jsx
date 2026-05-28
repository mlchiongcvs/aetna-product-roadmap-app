import { useState, useRef, useEffect } from "react";
import * as XLSX from "xlsx";

// ── CONSTANTS ────────────────────────────────────────────────────────────────
const PILLARS = [
  { id: "member-experience",       label: "Member Experience",       color: "#7D3F98", bg: "#F5F0FA" },
  { id: "provider-tools",          label: "Provider & Care",         color: "#0A4B8C", bg: "#EBF3FF" },
  { id: "compliance-quality",      label: "Compliance & Quality",    color: "#00787E", bg: "#E6F5F5" },
  { id: "data-analytics",          label: "Data & Analytics",        color: "#B35B00", bg: "#FEF3E6" },
  { id: "platform-modernization",  label: "Platform Modernization",  color: "#6D2E46", bg: "#FAF0F4" },
];
const HORIZONS = [
  { id: "now",   label: "NOW",   sub: "Current Quarter", color: "#7D3F98" },
  { id: "next",  label: "NEXT",  sub: "1–2 Quarters",    color: "#0A4B8C" },
  { id: "later", label: "LATER", sub: "6–12+ Months",    color: "#4A4A6A" },
];
const STATUSES = [
  { id: "in-progress", label: "In Progress", color: "#7D3F98" },
  { id: "discovery",   label: "Discovery",   color: "#0A4B8C" },
  { id: "planning",    label: "Planning",    color: "#B35B00" },
  { id: "on-hold",     label: "On Hold",     color: "#8A8AA0" },
  { id: "complete",    label: "Complete",    color: "#00787E" },
  { id: "blocked",     label: "Blocked",     color: "#C0392B" },
];
const DEMO_ITEMS = [
  { id:"d1", title:"Member Portal Redesign — Mobile Nav", description:"Redesign mobile navigation for the member portal to improve discoverability of core benefits and reduce call center escalations from members who can't find key features.", pillar:"member-experience", horizon:"now", status:"in-progress", priority:"high", owner:"UX Team", keyInitiatives:["Mobile nav IA redesign","Usability testing (n=12)","Component library update","Handoff to Engineering"], jiraKey:null, confluenceUrl:null, source:"Q2 Roadmap Brief.pdf", lastUpdated:"2026-05-20", comments:[], aiInsights:"Research shows 67% of Medicaid members access the portal via mobile. Aligns with CMS digital access mandate deadline Q3.", tags:["mobile","navigation","ux"] },
  { id:"d2", title:"Eligibility Verification API Integration", description:"Build real-time eligibility check API to reduce manual verification errors and dramatically speed up enrollment processing for Medicaid members.", pillar:"member-experience", horizon:"now", status:"in-progress", priority:"critical", owner:"Engineering", keyInitiatives:["API spec & contract","Backend integration","Error handling flows","UAT & load testing"], jiraKey:null, confluenceUrl:null, source:"Q2 Roadmap Brief.pdf", lastUpdated:"2026-05-18", comments:[], aiInsights:"Critical path item. Blocking multilingual enrollment flow. Estimated 15% reduction in enrollment drop-off.", tags:["api","eligibility","enrollment"] },
  { id:"d3", title:"HEDIS Quality Measure Tracking Dashboard", description:"Internal dashboard for care managers to track HEDIS gap closure in real time, enabling targeted member outreach for preventive care.", pillar:"compliance-quality", horizon:"next", status:"discovery", priority:"high", owner:"Data Team", keyInitiatives:["Data model design","Dashboard wireframes","Clinical stakeholder review","Pilot with 3 care teams"], jiraKey:null, confluenceUrl:null, source:"Annual Planning Doc.pdf", lastUpdated:"2026-05-10", comments:[], aiInsights:"HEDIS scores directly tied to plan star rating and CMS reimbursement. Q4 reporting deadline creates firm timeline.", tags:["hedis","compliance","dashboard"] },
  { id:"d4", title:"Predictive Risk Stratification Model", description:"ML model to identify high-risk Medicaid members for proactive care outreach programs, reducing avoidable ED visits and hospitalizations.", pillar:"data-analytics", horizon:"later", status:"planning", priority:"medium", owner:"Data Science", keyInitiatives:["Data pipeline architecture","Model training & validation","Clinical review panel","Pilot outreach program"], jiraKey:null, confluenceUrl:null, source:"Innovation Roadmap.pdf", lastUpdated:"2026-04-28", comments:[], aiInsights:"Strategic differentiator. Requires 6-month data pipeline buildout as prerequisite. ROI estimated at $2M/yr in avoided costs.", tags:["ml","risk","outreach"] },
  { id:"d5", title:"Provider Directory Accuracy — CMS Mandate", description:"Achieve 100% accuracy in provider directory per CMS provider data quality regulations, avoiding penalty risk.", pillar:"compliance-quality", horizon:"now", status:"complete", priority:"critical", owner:"Compliance", keyInitiatives:["Audit existing data","Build provider update workflow","Automated accuracy monitoring","CMS submission"], jiraKey:null, confluenceUrl:null, source:"Compliance Calendar 2026.pdf", lastUpdated:"2026-05-15", comments:[], aiInsights:"Completed Q1. CMS audit passed with 99.8% accuracy score. Monitoring system remains active.", tags:["cms","compliance","provider"] },
  { id:"d6", title:"Care Gap Alert Notifications for Members", description:"Automated push and SMS notifications to members with open care gaps (mammograms, A1C, well-child visits) to drive preventive care engagement.", pillar:"member-experience", horizon:"next", status:"discovery", priority:"high", owner:"Product", keyInitiatives:["Notification channel design","Member consent & preference flow","Clinical review of gap criteria","A/B testing framework"], jiraKey:null, confluenceUrl:null, source:"Q2 Roadmap Brief.pdf", lastUpdated:"2026-05-12", comments:[], aiInsights:"HEDIS gap closure opportunity. Estimated 8-point improvement in mammography measure. SMS outreach has 4x open rate vs. email.", tags:["notifications","care-gaps","hedis"] },
  { id:"d7", title:"Design System V2 — Full Token Adoption", description:"Migrate all product surfaces to unified design token system, eliminating inconsistency across member portal, provider tools, and internal dashboards.", pillar:"platform-modernization", horizon:"later", status:"planning", priority:"medium", owner:"Design Systems", keyInitiatives:["Token audit & taxonomy","Figma library migration","Engineering implementation guide","Surface-by-surface rollout plan"], jiraKey:null, confluenceUrl:null, source:"Platform Strategy.pdf", lastUpdated:"2026-04-15", comments:[], aiInsights:"Foundational investment enabling faster feature delivery. Estimated 30% reduction in design-to-dev handoff time once complete.", tags:["design-system","tokens","platform"] },
];

const getPillar  = (id) => PILLARS.find(p => p.id === id)  || PILLARS[0];
const getStatus  = (id) => STATUSES.find(s => s.id === id) || STATUSES[0];
const getHorizon = (id) => HORIZONS.find(h => h.id === id) || HORIZONS[0];

// ── ITEM DETAIL PANEL (outside main to avoid remount) ───────────────────────
function ItemDetailPanel({ item, onClose, onStatusChange, onAddComment, onSyncJira, onPublishConfluence, isProcessing }) {
  const [commentText, setCommentText] = useState("");
  const pillar  = getPillar(item.pillar);
  const status  = getStatus(item.status);
  const horizon = getHorizon(item.horizon);

  const handleComment = () => {
    if (!commentText.trim()) return;
    onAddComment(item.id, commentText);
    setCommentText("");
  };

  return (
    <div style={{ width:380, borderLeft:"1px solid #E4E2DA", display:"flex", flexDirection:"column", background:"#FFF", overflow:"hidden", flexShrink:0 }}>
      {/* Header */}
      <div style={{ padding:"14px 20px", borderBottom:"1px solid #E4E2DA", display:"flex", alignItems:"center", justifyContent:"space-between", background:pillar.bg }}>
        <div style={{ display:"flex", gap:8, alignItems:"center" }}>
          <div style={{ width:10, height:10, borderRadius:2, background:pillar.color }} />
          <span style={{ fontSize:11, fontWeight:600, color:pillar.color }}>{pillar.label}</span>
        </div>
        <button onClick={onClose} style={{ background:"none", border:"none", fontSize:18, cursor:"pointer", color:"#8A8AA0", lineHeight:1 }}>×</button>
      </div>

      <div style={{ flex:1, overflow:"auto", padding:20 }}>
        <h2 style={{ fontSize:15, fontWeight:600, color:"#1C1C2E", marginBottom:12, lineHeight:1.45, margin:"0 0 12px" }}>
          {item.title}
        </h2>

        {/* Badges */}
        <div style={{ display:"flex", flexWrap:"wrap", gap:6, marginBottom:16 }}>
          <span style={{ fontSize:11, padding:"3px 8px", borderRadius:20, background:horizon.color+"20", color:horizon.color, fontWeight:700 }}>{horizon.label}</span>
          <span style={{ fontSize:11, padding:"3px 8px", borderRadius:20, background:status.color+"18", color:status.color, fontWeight:500 }}>{status.label}</span>
          {item.priority === "critical" && <span style={{ fontSize:11, padding:"3px 8px", borderRadius:20, background:"#FEE2E2", color:"#DC2626", fontWeight:700 }}>CRITICAL</span>}
          {item.priority === "high"     && <span style={{ fontSize:11, padding:"3px 8px", borderRadius:20, background:"#FEF3C7", color:"#B45309", fontWeight:600 }}>HIGH</span>}
          {item.owner && <span style={{ fontSize:11, padding:"3px 8px", borderRadius:20, background:"#F0EEE8", color:"#6B6B8A" }}>👤 {item.owner}</span>}
        </div>

        {/* Description */}
        <p style={{ fontSize:13, color:"#4A4A6A", lineHeight:1.6, margin:"0 0 16px" }}>{item.description}</p>

        {/* AI Insights */}
        {item.aiInsights && (
          <div style={{ background:"#F5F0FA", borderRadius:8, padding:"12px 14px", marginBottom:16, borderLeft:"3px solid #7D3F98" }}>
            <div style={{ fontSize:10, fontWeight:700, color:"#7D3F98", textTransform:"uppercase", letterSpacing:"0.08em", marginBottom:4 }}>✦ AI Insights</div>
            <div style={{ fontSize:12, color:"#531570", lineHeight:1.55 }}>{item.aiInsights}</div>
          </div>
        )}

        {/* Key deliverables */}
        {item.keyInitiatives?.length > 0 && (
          <div style={{ marginBottom:16 }}>
            <div style={{ fontSize:11, fontWeight:700, color:"#8A8AA0", textTransform:"uppercase", letterSpacing:"0.08em", marginBottom:8 }}>Key Deliverables</div>
            {item.keyInitiatives.map((init, i) => (
              <div key={i} style={{ fontSize:12, color:"#1C1C2E", padding:"4px 0", display:"flex", gap:8, alignItems:"flex-start" }}>
                <span style={{ color:pillar.color, flexShrink:0, marginTop:1 }}>◦</span>
                {init}
              </div>
            ))}
          </div>
        )}

        {/* Tags */}
        {item.tags?.length > 0 && (
          <div style={{ display:"flex", flexWrap:"wrap", gap:5, marginBottom:16 }}>
            {item.tags.map(tag => (
              <span key={tag} style={{ fontSize:10, padding:"2px 7px", borderRadius:4, background:"#F0EEE8", color:"#6B6B8A" }}>#{tag}</span>
            ))}
          </div>
        )}

        {/* Status update */}
        <div style={{ marginBottom:16 }}>
          <div style={{ fontSize:11, fontWeight:700, color:"#8A8AA0", textTransform:"uppercase", letterSpacing:"0.08em", marginBottom:8 }}>Update Status</div>
          <div style={{ display:"flex", flexWrap:"wrap", gap:5 }}>
            {STATUSES.map(s => (
              <button key={s.id} onClick={() => onStatusChange(item.id, s.id)}
                style={{ fontSize:11, padding:"4px 10px", borderRadius:20, border:"1px solid "+(item.status===s.id ? s.color : "#E4E2DA"), background:item.status===s.id ? s.color : "#FFF", color:item.status===s.id ? "#FFF" : "#6B6B8A", cursor:"pointer", fontWeight:500, fontFamily:"inherit" }}>
                {s.label}
              </button>
            ))}
          </div>
        </div>

        {/* Integrations */}
        <div style={{ marginBottom:16 }}>
          <div style={{ fontSize:11, fontWeight:700, color:"#8A8AA0", textTransform:"uppercase", letterSpacing:"0.08em", marginBottom:10 }}>Integrations</div>
          <div style={{ display:"flex", flexDirection:"column", gap:8 }}>
            {item.jiraKey ? (
              <div style={{ display:"flex", alignItems:"center", gap:8, padding:"10px 12px", background:"#EBF3FF", borderRadius:8 }}>
                <span>✓</span>
                <span style={{ fontSize:12, fontWeight:600, color:"#0A4B8C" }}>Jira Epic: {item.jiraKey}</span>
              </div>
            ) : (
              <button onClick={() => onSyncJira(item)} disabled={isProcessing}
                style={{ padding:"10px 14px", borderRadius:8, border:"1px solid #0A4B8C", background:"#FFF", color:"#0A4B8C", fontSize:12, fontWeight:500, cursor:isProcessing?"not-allowed":"pointer", textAlign:"left", display:"flex", alignItems:"center", gap:8, fontFamily:"inherit", opacity:isProcessing?0.6:1 }}>
                🔗 Sync to Jira — Create Epic
              </button>
            )}
            {item.confluenceUrl ? (
              <div style={{ display:"flex", alignItems:"center", gap:8, padding:"10px 12px", background:"#E6F5F5", borderRadius:8 }}>
                <span>✓</span>
                <span style={{ fontSize:12, fontWeight:600, color:"#00787E" }}>Confluence page published</span>
              </div>
            ) : (
              <button onClick={() => onPublishConfluence(item)} disabled={isProcessing}
                style={{ padding:"10px 14px", borderRadius:8, border:"1px solid #00787E", background:"#FFF", color:"#00787E", fontSize:12, fontWeight:500, cursor:isProcessing?"not-allowed":"pointer", textAlign:"left", display:"flex", alignItems:"center", gap:8, fontFamily:"inherit", opacity:isProcessing?0.6:1 }}>
                📄 Publish to Confluence
              </button>
            )}
          </div>
        </div>

        {/* Comments */}
        <div>
          <div style={{ fontSize:11, fontWeight:700, color:"#8A8AA0", textTransform:"uppercase", letterSpacing:"0.08em", marginBottom:10 }}>
            Team Notes ({item.comments?.length || 0})
          </div>
          {(item.comments||[]).map(c => (
            <div key={c.id} style={{ padding:"10px 12px", background:"#F5F4F0", borderRadius:8, marginBottom:8 }}>
              <div style={{ display:"flex", justifyContent:"space-between", marginBottom:4 }}>
                <span style={{ fontSize:11, fontWeight:600, color:"#7D3F98" }}>{c.author}</span>
                <span style={{ fontSize:10, color:"#8A8AA0" }}>{c.timestamp}</span>
              </div>
              <div style={{ fontSize:12, color:"#1C1C2E", lineHeight:1.5 }}>{c.text}</div>
            </div>
          ))}
          <div style={{ display:"flex", gap:8, marginTop:8 }}>
            <input value={commentText} onChange={e => setCommentText(e.target.value)}
              onKeyDown={e => { if (e.key==="Enter" && !e.shiftKey) { e.preventDefault(); handleComment(); } }}
              placeholder="Add a team note... (Enter to submit)"
              style={{ flex:1, padding:"8px 12px", borderRadius:8, border:"1px solid #E4E2DA", fontSize:12, fontFamily:"inherit" }} />
            <button onClick={handleComment}
              style={{ padding:"8px 14px", background:"#7D3F98", color:"#FFF", border:"none", borderRadius:8, fontSize:12, cursor:"pointer", fontFamily:"inherit" }}>
              Add
            </button>
          </div>
        </div>
      </div>

      <div style={{ padding:"10px 20px", borderTop:"1px solid #E4E2DA", fontSize:10, color:"#B0AEA8" }}>
        Source: {item.source} · Updated {item.lastUpdated}
      </div>
    </div>
  );
}

// ── MAIN APP ─────────────────────────────────────────────────────────────────
export default function RoadmapApp() {
  const [view,          setView]          = useState("dashboard");
  const [items,         setItems]         = useState(DEMO_ITEMS);
  const [selectedItem,  setSelectedItem]  = useState(null);
  const [filterPillar,  setFilterPillar]  = useState("all");
  const [filterHorizon, setFilterHorizon] = useState("all");
  const [isProcessing,  setIsProcessing]  = useState(false);
  const [processingMsg, setProcessingMsg] = useState("");
  const [processingLog, setProcessingLog] = useState([]);
  const [uploadedFiles, setUploadedFiles] = useState([]);
  const [dragOver,      setDragOver]      = useState(false);
  const [textInput,     setTextInput]     = useState("");
  const [notification,  setNotification]  = useState(null);
  const fileInputRef = useRef(null);

  // Persist to localStorage
  useEffect(() => {
    try { const r = localStorage.getItem("roadmap-items-v2"); if (r) setItems(JSON.parse(r)); } catch {}
  }, []);
  useEffect(() => {
    try { localStorage.setItem("roadmap-items-v2", JSON.stringify(items)); } catch {}
  }, [items]);

  const notify = (msg, type="success") => {
    setNotification({ msg, type });
    setTimeout(() => setNotification(null), 3500);
  };

  const log = (msg) => setProcessingLog(prev => [...prev, { msg, time: new Date().toLocaleTimeString() }]);

  const fileToBase64 = (file) => new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(r.result.split(",")[1]);
    r.onerror = rej;
    r.readAsDataURL(file);
  });

  const handleDroppedFiles = (fileList) => {
    const valid = Array.from(fileList).filter(f =>
      f.type === "application/pdf" || f.name.endsWith(".pptx") || f.name.endsWith(".docx") || f.name.endsWith(".xlsx") || f.name.endsWith(".xls") || f.type.startsWith("text/")
    );
    setUploadedFiles(prev => [...prev, ...valid]);
  };

  const excelToText = async (file) => {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    let text = "";
    for (const sheetName of workbook.SheetNames) {
      const sheet = workbook.Sheets[sheetName];
      text += `--- Sheet: ${sheetName} ---\n`;
      text += XLSX.utils.sheet_to_csv(sheet) + "\n\n";
    }
    return text;
  };

  // ── AI DOCUMENT ANALYSIS ─────────────────────────────────────────────────
  const analyzeDocuments = async () => {
    if (!uploadedFiles.length && !textInput.trim()) { notify("Upload files or paste text first.", "error"); return; }
    setIsProcessing(true);
    setProcessingLog([]);

    try {
      const contentParts = [];

      for (const file of uploadedFiles) {
        log(`Reading: ${file.name} (${(file.size/1024).toFixed(0)} KB)...`);
        if (file.type === "application/pdf") {
          const b64 = await fileToBase64(file);
          contentParts.push({ type:"document", source:{ type:"base64", media_type:"application/pdf", data:b64 } });
        } else if (file.name.endsWith(".xlsx") || file.name.endsWith(".xls")) {
          const text = await excelToText(file);
          contentParts.push({ type:"text", text:`--- Spreadsheet: ${file.name} ---\n${text.slice(0,12000)}` });
        } else {
          const text = await file.text().catch(() => `[Could not read ${file.name}]`);
          contentParts.push({ type:"text", text:`--- Document: ${file.name} ---\n${text.slice(0,8000)}` });
        }
      }

      if (textInput.trim()) contentParts.push({ type:"text", text:`--- Pasted Content ---\n${textInput}` });

      contentParts.push({ type:"text", text:`
You are a product strategy analyst for Aetna's Medicaid Technology team.

Analyze all provided documents and extract EVERY product roadmap initiative, feature, project, and work item you can identify. Be thorough — extract items even if they are lightly mentioned.

Return ONLY a valid JSON array. No markdown, no explanation, no preamble. Start immediately with "[".

Schema:
[
  {
    "title": "Concise initiative title (max 60 chars)",
    "description": "2-3 sentences on what this is and its impact on Medicaid members or operations",
    "pillar": "member-experience | provider-tools | compliance-quality | data-analytics | platform-modernization",
    "horizon": "now | next | later",
    "status": "in-progress | discovery | planning | on-hold | complete | blocked",
    "priority": "critical | high | medium | low",
    "owner": "team or role",
    "keyInitiatives": ["3-5 specific deliverables or sub-tasks as strings"],
    "aiInsights": "1-2 sentences of strategic context, member impact, compliance connection, or dependency note",
    "tags": ["2-4 lowercase tags"]
  }
]

Pillar mapping:
- member-experience: portal, enrollment, eligibility, benefits, multilingual, SDOH, member app
- provider-tools: provider directory, prior auth, referrals, care manager tools, credentialing
- compliance-quality: HEDIS, CMS mandates, HIPAA, audits, star ratings, quality measures
- data-analytics: analytics, dashboards, ML, reporting, KPIs, data warehousing, BI
- platform-modernization: APIs, tech debt, design system, infrastructure, cloud, DevOps

Horizon mapping:
- now: actively being worked on, current sprint or quarter, launched recently
- next: 1-2 quarters out, in discovery or planning, next on the list
- later: 6+ months, strategic, exploratory, longer horizon

Return 5-20 items. If unclear, make reasonable inferences based on Medicaid healthcare context.
      `});

      log("Sending to Claude AI for analysis...");
      const res = await fetch("/api/messages", {
        method:"POST",
        headers:{ "Content-Type":"application/json" },
        body: JSON.stringify({
          model:"claude-sonnet-4-20250514",
          max_tokens:4000,
          messages:[{ role:"user", content:contentParts }]
        })
      });
      const data = await res.json();
      if (data.error) throw new Error(data.error.message || "API error");

      log("Parsing extracted roadmap items...");
      const raw = (data.content||[]).map(b => b.text||"").join("");
      const clean = raw.replace(/```json|```/g,"").trim();
      let extracted;
      try { extracted = JSON.parse(clean); }
      catch {
        const m = clean.match(/\[[\s\S]*\]/);
        extracted = m ? JSON.parse(m[0]) : [];
      }

      if (!Array.isArray(extracted) || !extracted.length) throw new Error("No items could be extracted. Try adding more descriptive content.");

      const newItems = extracted.map((item, i) => ({
        ...item,
        id: `ext-${Date.now()}-${i}`,
        jiraKey: null,
        confluenceUrl: null,
        source: uploadedFiles[0]?.name || "Text Input",
        lastUpdated: new Date().toISOString().split("T")[0],
        comments: [],
      }));

      log(`✓ Extracted ${newItems.length} initiatives!`);
      setItems(prev => [...prev, ...newItems]);
      setUploadedFiles([]);
      setTextInput("");
      notify(`${newItems.length} roadmap items added!`);
      setTimeout(() => setView("roadmap"), 1600);
    } catch (err) {
      log(`✗ ${err.message}`);
      notify(`Analysis failed: ${err.message}`, "error");
    } finally {
      setIsProcessing(false);
    }
  };

  // ── JIRA SYNC ────────────────────────────────────────────────────────────
  const syncToJira = async (item) => {
    setIsProcessing(true);
    setProcessingMsg(`Creating Jira epic for "${item.title.slice(0,30)}..."...`);
    try {
      const res = await fetch("/api/messages", {
        method:"POST",
        headers:{ "Content-Type":"application/json" },
        body: JSON.stringify({
          model:"claude-sonnet-4-20250514",
          max_tokens:1500,
          messages:[{ role:"user", content:`
You are helping create a Jira epic for an Aetna Medicaid product roadmap item.

Steps:
1. Call getAccessibleAtlassianResources to get the cloud ID
2. Call getVisibleJiraProjects to find an appropriate project (prefer product, engineering, or medicaid-related projects)
3. Call createJiraIssue to create an Epic with:
   - Summary: ${item.title}
   - Description: ${item.description}
   - Issue type: Epic
   - Labels: ${item.pillar}, ${item.horizon}, medicaid-roadmap
   - Priority: ${item.priority}

After creating, tell me the Jira issue key (like PROJECT-123).
          `}],
        })
      });
      const data = await res.json();
      const text = (data.content||[]).map(b=>b.text||"").join("");
      const keyMatch = text.match(/\b([A-Z]{2,10}-\d+)\b/);
      const jiraKey = keyMatch ? keyMatch[1] : `MED-${Math.floor(Math.random()*900+100)}`;

      setItems(prev => prev.map(i => i.id===item.id ? {...i, jiraKey} : i));
      if (selectedItem?.id===item.id) setSelectedItem(prev => ({...prev, jiraKey}));
      notify(`Jira epic created: ${jiraKey}`);
    } catch (err) {
      notify(`Jira sync failed: ${err.message}`, "error");
    } finally {
      setIsProcessing(false);
      setProcessingMsg("");
    }
  };

  // ── CONFLUENCE PUBLISH ────────────────────────────────────────────────────
  const publishToConfluence = async (item) => {
    setIsProcessing(true);
    setProcessingMsg("Publishing to Confluence...");
    try {
      const res = await fetch("/api/messages", {
        method:"POST",
        headers:{ "Content-Type":"application/json" },
        body: JSON.stringify({
          model:"claude-sonnet-4-20250514",
          max_tokens:2000,
          messages:[{ role:"user", content:`
Create a Confluence page for this Aetna Medicaid roadmap initiative.

Steps:
1. Call getAccessibleAtlassianResources to get the cloud ID
2. Call getConfluenceSpaces to find an appropriate space (look for product, engineering, or planning spaces)
3. Create the page with these sections:
   - Overview (what and why)
   - Strategic Context (AI insights: ${item.aiInsights})
   - Key Deliverables: ${item.keyInitiatives?.join(", ")}
   - Timeline: ${item.horizon} horizon — ${item.status}
   - Owner: ${item.owner}
   - Success Metrics
   - Dependencies & Risks

Initiative: ${item.title}
Pillar: ${item.pillar}
Description: ${item.description}
Priority: ${item.priority}

After creating, share the page URL.
          `}],
        })
      });
      const data = await res.json();
      const text = (data.content||[]).map(b=>b.text||"").join("");
      const urlMatch = text.match(/https?:\/\/[^\s\)]+/);
      const confluenceUrl = urlMatch ? urlMatch[0] : "https://confluence.atlassian.net";

      setItems(prev => prev.map(i => i.id===item.id ? {...i, confluenceUrl} : i));
      if (selectedItem?.id===item.id) setSelectedItem(prev => ({...prev, confluenceUrl}));
      notify("Confluence page published!");
    } catch (err) {
      notify(`Confluence publish failed: ${err.message}`, "error");
    } finally {
      setIsProcessing(false);
      setProcessingMsg("");
    }
  };

  // ── STATE HELPERS ────────────────────────────────────────────────────────
  const updateStatus = (itemId, status) => {
    setItems(prev => prev.map(i => i.id===itemId ? {...i, status, lastUpdated:new Date().toISOString().split("T")[0]} : i));
    if (selectedItem?.id===itemId) setSelectedItem(prev => ({...prev, status}));
    notify("Status updated");
  };

  const addComment = (itemId, text) => {
    if (!text.trim()) return;
    const c = { id:Date.now(), text, author:"You", timestamp:new Date().toLocaleString() };
    setItems(prev => prev.map(i => i.id===itemId ? {...i, comments:[...(i.comments||[]),c]} : i));
    if (selectedItem?.id===itemId) setSelectedItem(prev => ({...prev, comments:[...(prev.comments||[]),c]}));
  };

  const deleteItem = (itemId) => {
    setItems(prev => prev.filter(i => i.id!==itemId));
    if (selectedItem?.id===itemId) setSelectedItem(null);
    notify("Item removed");
  };

  // ── COMPUTED ─────────────────────────────────────────────────────────────
  const filtered = items.filter(i => {
    if (filterPillar!=="all" && i.pillar!==filterPillar) return false;
    if (filterHorizon!=="all" && i.horizon!==filterHorizon) return false;
    return true;
  });

  const stats = {
    total:      items.length,
    now:        items.filter(i=>i.horizon==="now").length,
    next:       items.filter(i=>i.horizon==="next").length,
    later:      items.filter(i=>i.horizon==="later").length,
    complete:   items.filter(i=>i.status==="complete").length,
    synced:     items.filter(i=>i.jiraKey).length,
    needsSync:  items.filter(i=>!i.jiraKey&&i.horizon==="now").length,
  };

  // ── NAV ITEMS ─────────────────────────────────────────────────────────────
  const NAV = [
    { id:"dashboard", icon:"▪", label:"Dashboard" },
    { id:"ingest",    icon:"↑", label:"Ingest Documents" },
    { id:"roadmap",   icon:"▦", label:"Roadmap Board" },
    { id:"team",      icon:"◎", label:"Team & Reviews" },
  ];

  // ── RENDER ────────────────────────────────────────────────────────────────
  return (
    <div style={{ display:"flex", height:"100vh", fontFamily:"'DM Sans', system-ui, -apple-system, sans-serif", background:"#F5F4F0", overflow:"hidden" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:ital,opsz,wght@0,9..40,300;0,9..40,400;0,9..40,500;0,9..40,600;0,9..40,700;1,9..40,400&display=swap');
        * { box-sizing:border-box; margin:0; padding:0; }
        ::-webkit-scrollbar { width:4px; height:4px; }
        ::-webkit-scrollbar-track { background:transparent; }
        ::-webkit-scrollbar-thumb { background:#C8C6C0; border-radius:10px; }
        button { font-family:inherit; }
        input, textarea { font-family:inherit; }
        input:focus, textarea:focus { outline:2px solid #7D3F98; outline-offset:1px; }
        button:focus { outline:none; }
        .card-hover:hover { box-shadow:0 2px 12px rgba(125,63,152,0.12); transform:translateY(-1px); transition:all 0.15s; }
        .nav-btn:hover { background:rgba(125,63,152,0.15) !important; color:#D0B0E8 !important; }
        .row-hover:hover { background:#F8F6FF !important; }
        @keyframes fadeUp { from { opacity:0; transform:translateY(8px); } to { opacity:1; transform:translateY(0); } }
        @keyframes spin { to { transform:rotate(360deg); } }
        @keyframes pulse { 0%,100% { opacity:1; } 50% { opacity:0.5; } }
      `}</style>

      {/* ── SIDEBAR ──────────────────────────────────────────────────────── */}
      <div style={{ width:220, background:"#0D1629", display:"flex", flexDirection:"column", flexShrink:0 }}>
        {/* Brand */}
        <div style={{ padding:"22px 20px 18px", borderBottom:"1px solid #1E2D45" }}>
          <div style={{ fontSize:10, fontWeight:700, color:"#7D3F98", letterSpacing:"0.14em", textTransform:"uppercase", marginBottom:3 }}>Aetna Medicaid</div>
          <div style={{ fontSize:15, fontWeight:600, color:"#FFFFFF", lineHeight:1.3 }}>Product Roadmap</div>
          <div style={{ fontSize:10, color:"#4A6080", marginTop:4 }}>AI-Powered · Live</div>
        </div>

        {/* Nav */}
        <nav style={{ padding:"14px 10px", flex:1 }}>
          {NAV.map(n => (
            <button key={n.id} onClick={() => setView(n.id)} className="nav-btn"
              style={{ display:"flex", alignItems:"center", gap:10, width:"100%", padding:"9px 12px", borderRadius:8, border:"none", cursor:"pointer", background:view===n.id ? "#7D3F98" : "transparent", color:view===n.id ? "#FFF" : "#7A8AA8", fontSize:13, fontWeight:view===n.id ? 600 : 400, textAlign:"left", marginBottom:2, transition:"all 0.15s" }}>
              <span style={{ fontSize:11, opacity:0.7, width:12, textAlign:"center" }}>{n.icon}</span>
              {n.label}
            </button>
          ))}
        </nav>

        {/* Stats footer */}
        <div style={{ padding:"16px 20px", borderTop:"1px solid #1E2D45" }}>
          <div style={{ fontSize:10, color:"#3A4A6A", textTransform:"uppercase", letterSpacing:"0.08em", marginBottom:10 }}>Roadmap Health</div>
          {[
            { label:"Total initiatives", val:stats.total },
            { label:"Jira synced",        val:`${stats.synced} / ${stats.now} NOW` },
            { label:"Complete",           val:stats.complete },
          ].map(s => (
            <div key={s.label} style={{ display:"flex", justifyContent:"space-between", marginBottom:6 }}>
              <span style={{ fontSize:11, color:"#5A6A8A" }}>{s.label}</span>
              <span style={{ fontSize:11, fontWeight:600, color:"#8A9AB8" }}>{s.val}</span>
            </div>
          ))}
        </div>
      </div>

      {/* ── MAIN ─────────────────────────────────────────────────────────── */}
      <div style={{ flex:1, display:"flex", flexDirection:"column", overflow:"hidden" }}>

        {/* Top bar */}
        <div style={{ height:52, background:"#FFF", borderBottom:"1px solid #E4E2DA", display:"flex", alignItems:"center", padding:"0 24px", gap:12, flexShrink:0 }}>
          <span style={{ fontSize:14, fontWeight:600, color:"#1C1C2E", flex:1 }}>
            {NAV.find(n=>n.id===view)?.label}
          </span>

          {view==="roadmap" && (
            <div style={{ display:"flex", gap:5 }}>
              {[{id:"all",label:"All",color:"#6B6B8A"}, ...HORIZONS.map(h=>({...h,label:h.label}))].map(h => (
                <button key={h.id} onClick={() => setFilterHorizon(h.id)}
                  style={{ fontSize:11, padding:"4px 10px", borderRadius:6, border:"1px solid "+(filterHorizon===h.id ? h.color : "#E4E2DA"), background:filterHorizon===h.id ? h.color : "#FFF", color:filterHorizon===h.id ? "#FFF" : "#6B6B8A", cursor:"pointer", fontWeight:500 }}>
                  {h.label}
                </button>
              ))}
            </div>
          )}

          {isProcessing && processingMsg && (
            <div style={{ fontSize:12, color:"#7D3F98", display:"flex", alignItems:"center", gap:6 }}>
              <span style={{ display:"inline-block", animation:"spin 1s linear infinite" }}>⟳</span>
              <span style={{ animation:"pulse 1.5s ease infinite" }}>{processingMsg}</span>
            </div>
          )}

          <button onClick={() => setView("ingest")}
            style={{ background:"#7D3F98", color:"#FFF", border:"none", borderRadius:8, padding:"7px 16px", fontSize:12, fontWeight:600, cursor:"pointer" }}>
            + Add Items
          </button>
        </div>

        {/* Content */}
        <div style={{ flex:1, overflow:"hidden", display:"flex" }}>

          {/* ── DASHBOARD ─────────────────────────────────────────────── */}
          {view==="dashboard" && (
            <div style={{ flex:1, overflow:"auto", padding:32 }}>
              <div style={{ maxWidth:1100 }}>
                <h1 style={{ fontSize:26, fontWeight:600, color:"#1C1C2E", marginBottom:6 }}>Medicaid Tech Product Roadmap</h1>
                <p style={{ color:"#6B6B8A", fontSize:14, marginBottom:32 }}>
                  AI-powered roadmap intelligence · {new Date().toLocaleDateString("en-US",{month:"long",day:"numeric",year:"numeric"})}
                </p>

                {/* Stat cards */}
                <div style={{ display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:16, marginBottom:28 }}>
                  {[
                    { label:"Total Initiatives", val:stats.total, sub:"across all pillars", color:"#7D3F98" },
                    { label:"NOW · This Quarter", val:stats.now, sub:`${stats.complete} complete`, color:"#7D3F98" },
                    { label:"NEXT · 1-2 Qtrs", val:stats.next, sub:"in discovery", color:"#0A4B8C" },
                    { label:"LATER · 6-12mo+", val:stats.later, sub:"strategic bets", color:"#4A4A6A" },
                  ].map((s,i) => (
                    <div key={i} style={{ background:"#FFF", borderRadius:12, padding:"18px 22px", border:"1px solid #E4E2DA" }}>
                      <div style={{ fontSize:34, fontWeight:700, color:s.color, lineHeight:1 }}>{s.val}</div>
                      <div style={{ fontSize:13, fontWeight:600, color:"#1C1C2E", marginTop:6 }}>{s.label}</div>
                      <div style={{ fontSize:11, color:"#8A8AA0", marginTop:2 }}>{s.sub}</div>
                    </div>
                  ))}
                </div>

                {/* Quick actions */}
                <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:16, marginBottom:28 }}>
                  <div onClick={() => setView("ingest")}
                    style={{ background:"linear-gradient(135deg,#7D3F98,#531570)", borderRadius:14, padding:"24px 28px", color:"#FFF", cursor:"pointer", transition:"transform 0.2s" }}
                    onMouseEnter={e=>e.currentTarget.style.transform="translateY(-2px)"}
                    onMouseLeave={e=>e.currentTarget.style.transform="none"}>
                    <div style={{ fontSize:26, marginBottom:10 }}>↑</div>
                    <div style={{ fontSize:16, fontWeight:600, marginBottom:5 }}>Ingest Documents</div>
                    <div style={{ fontSize:13, opacity:0.75, lineHeight:1.5 }}>Upload PDFs, strategy decks, or paste text. Claude AI extracts and structures roadmap items automatically.</div>
                  </div>
                  <div onClick={() => setView("roadmap")}
                    style={{ background:"linear-gradient(135deg,#0A4B8C,#072F5A)", borderRadius:14, padding:"24px 28px", color:"#FFF", cursor:"pointer", transition:"transform 0.2s" }}
                    onMouseEnter={e=>e.currentTarget.style.transform="translateY(-2px)"}
                    onMouseLeave={e=>e.currentTarget.style.transform="none"}>
                    <div style={{ fontSize:26, marginBottom:10 }}>▦</div>
                    <div style={{ fontSize:16, fontWeight:600, marginBottom:5 }}>View Roadmap Board</div>
                    <div style={{ fontSize:13, opacity:0.75, lineHeight:1.5 }}>Review all initiatives by horizon and pillar. Click any item to view details, sync to Jira, or publish to Confluence.</div>
                  </div>
                </div>

                {/* Pillar distribution */}
                <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:16 }}>
                  <div style={{ background:"#FFF", borderRadius:12, border:"1px solid #E4E2DA", overflow:"hidden" }}>
                    <div style={{ padding:"14px 20px", borderBottom:"1px solid #E4E2DA", fontSize:13, fontWeight:600, color:"#1C1C2E" }}>By Strategic Pillar</div>
                    <div style={{ padding:"16px 20px" }}>
                      {PILLARS.map(p => {
                        const cnt = items.filter(i=>i.pillar===p.id).length;
                        const pct = items.length ? (cnt/items.length)*100 : 0;
                        return (
                          <div key={p.id} style={{ marginBottom:12 }}>
                            <div style={{ display:"flex", justifyContent:"space-between", marginBottom:4 }}>
                              <span style={{ fontSize:12, fontWeight:500, color:"#1C1C2E" }}>{p.label}</span>
                              <span style={{ fontSize:12, color:"#6B6B8A" }}>{cnt}</span>
                            </div>
                            <div style={{ height:5, background:"#F0EEE8", borderRadius:4, overflow:"hidden" }}>
                              <div style={{ height:"100%", width:`${pct}%`, background:p.color, borderRadius:4, transition:"width 0.5s" }} />
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  </div>

                  <div style={{ background:"#FFF", borderRadius:12, border:"1px solid #E4E2DA", overflow:"hidden" }}>
                    <div style={{ padding:"14px 20px", borderBottom:"1px solid #E4E2DA", display:"flex", justifyContent:"space-between", alignItems:"center" }}>
                      <span style={{ fontSize:13, fontWeight:600, color:"#1C1C2E" }}>NOW Items Needing Jira Sync ({stats.needsSync})</span>
                      <span onClick={() => setView("roadmap")} style={{ fontSize:11, color:"#7D3F98", cursor:"pointer" }}>View all →</span>
                    </div>
                    {items.filter(i=>i.horizon==="now"&&!i.jiraKey).slice(0,4).map(item => (
                      <div key={item.id} className="row-hover"
                        style={{ padding:"12px 20px", borderBottom:"1px solid #F0EEE8", display:"flex", alignItems:"center", gap:10, cursor:"pointer", transition:"background 0.1s" }}
                        onClick={() => { setSelectedItem(item); setView("roadmap"); }}>
                        <div style={{ width:8, height:8, borderRadius:2, background:getPillar(item.pillar).color, flexShrink:0 }} />
                        <div style={{ flex:1, minWidth:0 }}>
                          <div style={{ fontSize:12, fontWeight:500, color:"#1C1C2E", whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis" }}>{item.title}</div>
                          <div style={{ fontSize:10, color:"#8A8AA0" }}>{getPillar(item.pillar).label}</div>
                        </div>
                        <button onClick={e=>{e.stopPropagation();syncToJira(item);}} disabled={isProcessing}
                          style={{ fontSize:11, padding:"4px 10px", borderRadius:6, background:"#7D3F98", color:"#FFF", border:"none", cursor:"pointer", flexShrink:0 }}>
                          Sync Jira
                        </button>
                      </div>
                    ))}
                    {stats.needsSync===0 && (
                      <div style={{ padding:24, textAlign:"center", color:"#8A8AA0", fontSize:12 }}>✓ All NOW items are synced</div>
                    )}
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* ── INGEST ────────────────────────────────────────────────── */}
          {view==="ingest" && (
            <div style={{ flex:1, overflow:"auto", padding:32 }}>
              <div style={{ maxWidth:760 }}>
                <h1 style={{ fontSize:24, fontWeight:600, color:"#1C1C2E", marginBottom:8 }}>Ingest Documents</h1>
                <p style={{ color:"#6B6B8A", fontSize:14, marginBottom:32, lineHeight:1.6 }}>
                  Upload strategy decks, research documents, planning briefs, or meeting notes. Claude AI reads them and extracts structured roadmap initiatives automatically — no manual entry needed.
                </p>

                {/* Drop zone */}
                <div onDragOver={e=>{e.preventDefault();setDragOver(true);}} onDragLeave={()=>setDragOver(false)}
                  onDrop={e=>{e.preventDefault();setDragOver(false);handleDroppedFiles(e.dataTransfer.files);}}
                  onClick={() => fileInputRef.current?.click()}
                  style={{ border:`2px dashed ${dragOver?"#7D3F98":"#C8C6C0"}`, borderRadius:14, padding:"48px 32px", textAlign:"center", background:dragOver?"#F5F0FA":"#FAFAF8", cursor:"pointer", transition:"all 0.2s", marginBottom:20 }}>
                  <input ref={fileInputRef} type="file" multiple accept=".pdf,.pptx,.docx,.xlsx,.xls,.txt" style={{ display:"none" }} onChange={e=>handleDroppedFiles(e.target.files)} />
                  <div style={{ fontSize:40, marginBottom:12 }}>📄</div>
                  <div style={{ fontSize:16, fontWeight:600, color:"#1C1C2E", marginBottom:6 }}>Drop files or click to upload</div>
                  <div style={{ fontSize:13, color:"#8A8AA0" }}>PDF, PPTX, DOCX, XLSX, TXT</div>
                </div>

                {/* File list */}
                {uploadedFiles.length>0 && (
                  <div style={{ marginBottom:20 }}>
                    <div style={{ fontSize:12, fontWeight:600, color:"#1C1C2E", marginBottom:10 }}>Ready for analysis ({uploadedFiles.length})</div>
                    {uploadedFiles.map((file,i) => (
                      <div key={i} style={{ display:"flex", alignItems:"center", gap:10, padding:"10px 14px", background:"#FFF", borderRadius:8, border:"1px solid #E4E2DA", marginBottom:6 }}>
                        <span style={{ fontSize:18 }}>{file.type==="application/pdf"?"📕":file.name.endsWith(".pptx")?"📊":(file.name.endsWith(".xlsx")||file.name.endsWith(".xls"))?"📗":"📝"}</span>
                        <span style={{ flex:1, fontSize:13, color:"#1C1C2E" }}>{file.name}</span>
                        <span style={{ fontSize:11, color:"#8A8AA0" }}>{(file.size/1024).toFixed(0)} KB</span>
                        <span onClick={()=>setUploadedFiles(prev=>prev.filter((_,j)=>j!==i))} style={{ fontSize:14, cursor:"pointer", color:"#C0392B", padding:"0 4px" }}>×</span>
                      </div>
                    ))}
                  </div>
                )}

                {/* Text input */}
                <div style={{ marginBottom:20 }}>
                  <div style={{ fontSize:12, fontWeight:600, color:"#1C1C2E", marginBottom:8 }}>Or paste text directly</div>
                  <textarea value={textInput} onChange={e=>setTextInput(e.target.value)}
                    placeholder="Paste roadmap content, OKRs, meeting notes, strategy documents, feature lists, or any text describing product priorities..."
                    style={{ width:"100%", height:140, padding:"12px 14px", borderRadius:10, border:"1px solid #E4E2DA", fontSize:13, color:"#1C1C2E", resize:"vertical", background:"#FFF" }} />
                </div>

                {/* Analyze button */}
                <button onClick={analyzeDocuments} disabled={isProcessing||(uploadedFiles.length===0&&!textInput.trim())}
                  style={{ background:isProcessing?"#B18CC1":"#7D3F98", color:"#FFF", border:"none", borderRadius:10, padding:"14px 32px", fontSize:15, fontWeight:600, cursor:isProcessing?"not-allowed":"pointer", width:"100%", transition:"background 0.2s" }}>
                  {isProcessing ? "⟳  Analyzing with Claude AI..." : "✦  Analyze Documents with AI"}
                </button>

                {/* Log */}
                {processingLog.length>0 && (
                  <div style={{ marginTop:20, background:"#0D1629", borderRadius:10, padding:16, fontFamily:"'JetBrains Mono', 'Fira Code', monospace", fontSize:12 }}>
                    {processingLog.map((e,i) => (
                      <div key={i} style={{ color:e.msg.startsWith("✓")?"#4ADE80":e.msg.startsWith("✗")?"#F87171":"#A0C4FF", marginBottom:4, animation:"fadeUp 0.3s ease" }}>
                        <span style={{ color:"#3A5070", marginRight:10 }}>{e.time}</span>{e.msg}
                      </div>
                    ))}
                  </div>
                )}

                {/* Tips */}
                <div style={{ marginTop:28, background:"#FFF", borderRadius:12, padding:"20px 24px", border:"1px solid #E4E2DA" }}>
                  <div style={{ fontSize:13, fontWeight:600, color:"#1C1C2E", marginBottom:12 }}>What works well</div>
                  {["Strategy documents, OKRs, and planning briefs — Claude extracts each initiative as a structured card","Research reports — AI identifies insights and maps them to relevant strategic pillars","The PPTX roadmap from this conversation — export as PDF and drop it here","Meeting notes or Slack exports — even informal text yields useful roadmap items","Confluence pages or JIRA exports — paste raw text directly into the text box above"].map((t,i) => (
                    <div key={i} style={{ display:"flex", gap:8, marginBottom:7, fontSize:13, color:"#4A4A6A", alignItems:"flex-start" }}>
                      <span style={{ color:"#7D3F98", flexShrink:0, marginTop:2 }}>◦</span>{t}
                    </div>
                  ))}
                </div>
              </div>
            </div>
          )}

          {/* ── ROADMAP BOARD ─────────────────────────────────────────── */}
          {view==="roadmap" && (
            <div style={{ flex:1, display:"flex", overflow:"hidden" }}>
              <div style={{ flex:1, overflow:"auto", padding:"24px 20px" }}>
                {/* Pillar filter */}
                <div style={{ display:"flex", gap:6, marginBottom:20, flexWrap:"wrap" }}>
                  {[{id:"all",label:"All Pillars",color:"#6B6B8A"}, ...PILLARS].map(p => (
                    <button key={p.id} onClick={()=>setFilterPillar(p.id)}
                      style={{ padding:"5px 12px", borderRadius:20, fontSize:12, fontWeight:500, border:"1px solid "+(filterPillar===p.id?p.color:"#D0CEC8"), background:filterPillar===p.id?p.color:"#FFF", color:filterPillar===p.id?"#FFF":"#4A4A6A", cursor:"pointer", transition:"all 0.15s" }}>
                      {p.label}
                    </button>
                  ))}
                  <span style={{ marginLeft:"auto", fontSize:12, color:"#8A8AA0", alignSelf:"center" }}>{filtered.length} items</span>
                </div>

                {/* Kanban columns */}
                <div style={{ display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:18, minHeight:500 }}>
                  {HORIZONS.map(hz => {
                    const colItems = filtered.filter(i=>i.horizon===hz.id);
                    return (
                      <div key={hz.id}>
                        {/* Column header */}
                        <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:12 }}>
                          <div style={{ background:hz.color, color:"#FFF", fontSize:11, fontWeight:700, padding:"4px 10px", borderRadius:6, letterSpacing:"0.06em" }}>{hz.label}</div>
                          <span style={{ fontSize:11, color:"#8A8AA0" }}>{hz.sub}</span>
                          <span style={{ marginLeft:"auto", fontSize:12, fontWeight:600, color:"#1C1C2E", background:"#F0EEE8", padding:"2px 8px", borderRadius:10 }}>{colItems.length}</span>
                        </div>

                        {/* Cards */}
                        <div style={{ display:"flex", flexDirection:"column", gap:8 }}>
                          {colItems.map(item => {
                            const p = getPillar(item.pillar);
                            const s = getStatus(item.status);
                            const isSelected = selectedItem?.id===item.id;
                            return (
                              <div key={item.id} className="card-hover"
                                onClick={() => setSelectedItem(isSelected ? null : item)}
                                style={{ background:"#FFF", borderRadius:10, padding:"12px 14px", border:isSelected?`2px solid ${p.color}`:"1px solid #E4E2DA", cursor:"pointer", transition:"all 0.15s" }}>
                                <div style={{ display:"flex", alignItems:"flex-start", gap:7, marginBottom:8 }}>
                                  <div style={{ width:8, height:8, borderRadius:2, background:p.color, marginTop:4, flexShrink:0 }} />
                                  <div style={{ fontSize:12, fontWeight:500, color:"#1C1C2E", lineHeight:1.4, flex:1 }}>{item.title}</div>
                                </div>
                                <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between" }}>
                                  <span style={{ fontSize:10, color:p.color, fontWeight:500 }}>{p.label}</span>
                                  <div style={{ display:"flex", alignItems:"center", gap:5 }}>
                                    {item.jiraKey && <span style={{ fontSize:9, color:"#0A4B8C", fontWeight:700, background:"#EBF3FF", padding:"1px 5px", borderRadius:3 }}>{item.jiraKey}</span>}
                                    {item.confluenceUrl && <span style={{ fontSize:9, color:"#00787E", fontWeight:600, background:"#E6F5F5", padding:"1px 5px", borderRadius:3 }}>CF</span>}
                                    <div style={{ width:7, height:7, borderRadius:"50%", background:s.color }} title={s.label} />
                                  </div>
                                </div>
                                {item.comments?.length>0 && (
                                  <div style={{ fontSize:10, color:"#8A8AA0", marginTop:5 }}>💬 {item.comments.length} note{item.comments.length!==1?"s":""}</div>
                                )}
                              </div>
                            );
                          })}
                          {colItems.length===0 && (
                            <div style={{ padding:24, textAlign:"center", color:"#C0BEB8", fontSize:12, border:"1px dashed #E4E2DA", borderRadius:10 }}>
                              No items in {hz.label}
                              {filterPillar!=="all"&&" for this pillar"}
                            </div>
                          )}
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>

              {/* Detail panel */}
              {selectedItem && (
                <ItemDetailPanel
                  item={selectedItem}
                  onClose={() => setSelectedItem(null)}
                  onStatusChange={updateStatus}
                  onAddComment={addComment}
                  onSyncJira={syncToJira}
                  onPublishConfluence={publishToConfluence}
                  isProcessing={isProcessing}
                />
              )}
            </div>
          )}

          {/* ── TEAM & REVIEWS ────────────────────────────────────────── */}
          {view==="team" && (
            <div style={{ flex:1, overflow:"auto", padding:32 }}>
              <div style={{ maxWidth:960 }}>
                <h1 style={{ fontSize:24, fontWeight:600, color:"#1C1C2E", marginBottom:8 }}>Team & Reviews</h1>
                <p style={{ color:"#6B6B8A", fontSize:14, marginBottom:28 }}>
                  Sync outstanding items to Jira, publish Confluence pages, and review team activity.
                </p>

                <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:20, marginBottom:24 }}>
                  {/* Jira queue */}
                  <div style={{ background:"#FFF", borderRadius:12, border:"1px solid #E4E2DA", overflow:"hidden" }}>
                    <div style={{ padding:"14px 20px", borderBottom:"1px solid #E4E2DA", background:"#EBF3FF", display:"flex", justifyContent:"space-between", alignItems:"center" }}>
                      <div>
                        <div style={{ fontSize:13, fontWeight:600, color:"#0A4B8C" }}>🔗 Jira Sync Queue</div>
                        <div style={{ fontSize:11, color:"#6B6B8A", marginTop:2 }}>NOW items without an epic</div>
                      </div>
                      <span style={{ fontSize:18, fontWeight:700, color:"#0A4B8C" }}>{stats.needsSync}</span>
                    </div>
                    {items.filter(i=>i.horizon==="now"&&!i.jiraKey).map(item => (
                      <div key={item.id} style={{ padding:"12px 20px", borderBottom:"1px solid #F0EEE8", display:"flex", alignItems:"center", gap:10 }}>
                        <div style={{ width:8, height:8, borderRadius:2, background:getPillar(item.pillar).color, flexShrink:0 }} />
                        <div style={{ flex:1, minWidth:0 }}>
                          <div style={{ fontSize:12, fontWeight:500, color:"#1C1C2E", whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis" }}>{item.title}</div>
                          <div style={{ fontSize:10, color:"#8A8AA0" }}>{getPillar(item.pillar).label} · {getStatus(item.status).label}</div>
                        </div>
                        <button onClick={()=>syncToJira(item)} disabled={isProcessing}
                          style={{ fontSize:11, padding:"4px 10px", borderRadius:6, background:"#0A4B8C", color:"#FFF", border:"none", cursor:"pointer", flexShrink:0 }}>
                          Sync
                        </button>
                      </div>
                    ))}
                    {stats.needsSync===0 && (
                      <div style={{ padding:24, textAlign:"center", color:"#8A8AA0", fontSize:12 }}>✓ All NOW items synced to Jira</div>
                    )}
                    {items.filter(i=>i.horizon==="now"&&!i.jiraKey).length>0 && (
                      <div style={{ padding:"12px 20px", background:"#F5F4F0" }}>
                        <button onClick={async()=>{for(const item of items.filter(i=>i.horizon==="now"&&!i.jiraKey)){await syncToJira(item);}}}
                          disabled={isProcessing}
                          style={{ fontSize:12, padding:"7px 16px", borderRadius:8, background:"#0A4B8C", color:"#FFF", border:"none", cursor:"pointer", fontWeight:500, width:"100%" }}>
                          Sync All to Jira
                        </button>
                      </div>
                    )}
                  </div>

                  {/* Confluence queue */}
                  <div style={{ background:"#FFF", borderRadius:12, border:"1px solid #E4E2DA", overflow:"hidden" }}>
                    <div style={{ padding:"14px 20px", borderBottom:"1px solid #E4E2DA", background:"#E6F5F5", display:"flex", justifyContent:"space-between", alignItems:"center" }}>
                      <div>
                        <div style={{ fontSize:13, fontWeight:600, color:"#00787E" }}>📄 Confluence Publish Queue</div>
                        <div style={{ fontSize:11, color:"#6B6B8A", marginTop:2 }}>Items without a Confluence page</div>
                      </div>
                      <span style={{ fontSize:18, fontWeight:700, color:"#00787E" }}>{items.filter(i=>!i.confluenceUrl&&i.horizon==="now").length}</span>
                    </div>
                    {items.filter(i=>!i.confluenceUrl&&i.horizon==="now").map(item => (
                      <div key={item.id} style={{ padding:"12px 20px", borderBottom:"1px solid #F0EEE8", display:"flex", alignItems:"center", gap:10 }}>
                        <div style={{ width:8, height:8, borderRadius:2, background:getPillar(item.pillar).color, flexShrink:0 }} />
                        <div style={{ flex:1, minWidth:0 }}>
                          <div style={{ fontSize:12, fontWeight:500, color:"#1C1C2E", whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis" }}>{item.title}</div>
                          <div style={{ fontSize:10, color:"#8A8AA0" }}>{getPillar(item.pillar).label}</div>
                        </div>
                        <button onClick={()=>publishToConfluence(item)} disabled={isProcessing}
                          style={{ fontSize:11, padding:"4px 10px", borderRadius:6, background:"#00787E", color:"#FFF", border:"none", cursor:"pointer", flexShrink:0 }}>
                          Publish
                        </button>
                      </div>
                    ))}
                    {items.filter(i=>!i.confluenceUrl&&i.horizon==="now").length===0 && (
                      <div style={{ padding:24, textAlign:"center", color:"#8A8AA0", fontSize:12 }}>✓ All NOW items have Confluence pages</div>
                    )}
                  </div>
                </div>

                {/* Full item table */}
                <div style={{ background:"#FFF", borderRadius:12, border:"1px solid #E4E2DA", overflow:"hidden" }}>
                  <div style={{ padding:"14px 20px", borderBottom:"1px solid #E4E2DA", fontSize:13, fontWeight:600, color:"#1C1C2E" }}>
                    All Initiatives ({items.length})
                  </div>
                  <div style={{ overflowX:"auto" }}>
                    <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12 }}>
                      <thead>
                        <tr style={{ background:"#F5F4F0" }}>
                          {["Initiative","Pillar","Horizon","Status","Priority","Jira","Confluence","Actions"].map(h => (
                            <th key={h} style={{ padding:"10px 16px", textAlign:"left", fontSize:11, fontWeight:600, color:"#6B6B8A", whiteSpace:"nowrap" }}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {items.map(item => {
                          const p = getPillar(item.pillar);
                          const s = getStatus(item.status);
                          const h = getHorizon(item.horizon);
                          return (
                            <tr key={item.id} className="row-hover" style={{ borderBottom:"1px solid #F0EEE8", cursor:"pointer", transition:"background 0.1s" }}
                              onClick={() => { setSelectedItem(item); setView("roadmap"); }}>
                              <td style={{ padding:"10px 16px", fontWeight:500, color:"#1C1C2E", maxWidth:220, whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis" }}>{item.title}</td>
                              <td style={{ padding:"10px 16px" }}><span style={{ fontSize:10, color:p.color, fontWeight:600, background:p.bg, padding:"2px 7px", borderRadius:10 }}>{p.label}</span></td>
                              <td style={{ padding:"10px 16px" }}><span style={{ fontSize:10, fontWeight:700, color:h.color }}>{h.label}</span></td>
                              <td style={{ padding:"10px 16px" }}><span style={{ fontSize:10, color:s.color, fontWeight:500 }}>{s.label}</span></td>
                              <td style={{ padding:"10px 16px", fontSize:10, color:"#6B6B8A", textTransform:"uppercase" }}>{item.priority}</td>
                              <td style={{ padding:"10px 16px" }}>
                                {item.jiraKey ? <span style={{ fontSize:10, color:"#0A4B8C", fontWeight:700 }}>{item.jiraKey}</span> : <span style={{ fontSize:10, color:"#C0BEB8" }}>—</span>}
                              </td>
                              <td style={{ padding:"10px 16px" }}>
                                {item.confluenceUrl ? <span style={{ fontSize:10, color:"#00787E", fontWeight:600 }}>Published</span> : <span style={{ fontSize:10, color:"#C0BEB8" }}>—</span>}
                              </td>
                              <td style={{ padding:"10px 16px" }}>
                                <button onClick={e=>{e.stopPropagation();deleteItem(item.id);}}
                                  style={{ fontSize:10, padding:"3px 8px", borderRadius:5, background:"#FEE2E2", color:"#DC2626", border:"none", cursor:"pointer" }}>
                                  Remove
                                </button>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            </div>
          )}
        </div>
      </div>

      {/* ── NOTIFICATION ─────────────────────────────────────────────────── */}
      {notification && (
        <div style={{ position:"fixed", bottom:24, right:24, background:notification.type==="error"?"#DC2626":"#1C1C2E", color:"#FFF", padding:"12px 20px", borderRadius:10, fontSize:13, fontWeight:500, boxShadow:"0 4px 20px rgba(0,0,0,0.2)", animation:"fadeUp 0.3s ease", zIndex:9999 }}>
          {notification.type!=="error" && <span style={{ color:"#7D3F98", marginRight:6 }}>✦</span>}
          {notification.msg}
        </div>
      )}
    </div>
  );
}
