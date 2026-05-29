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

// ── TIMELINE DATA ───────────────────────────────────────────────────────────
const TIMELINE_PILLARS = [
  { id:"enterprise-value", label:"Enterprise Value", sub:"Enterprise-wide platforms", color:"#1565C0", dot:"#1565C0" },
  { id:"efficiency",       label:"Efficiency",       sub:"Standardized processes",    color:"#2E7D32", dot:"#2E7D32" },
  { id:"growth",           label:"Growth",           sub:"Innovative technology",     color:"#6A1B9A", dot:"#6A1B9A" },
  { id:"cost-reduction",   label:"Cost Reduction",   sub:"Shared operational platforms", color:"#B71C1C", dot:"#B71C1C" },
];
const TIMELINE_QUARTERS = ["Q3 '26","Q4 '26","Q1 '27","Q2 '27","Q3 '27","Q4 '27","2028","2029+"];
const TIMELINE_PHASES = [
  { label:"Now", cols:2 },
  { label:"Next", cols:3 },
  { label:"Future", cols:3 },
];
const TIMELINE_ITEMS = [
  // Product capabilities (solid bars)
  { id:"t1",  name:"Dynamo",     domain:"Claims",            pillar:"efficiency",     type:"product", startCol:0, spanCols:8, users:[0,1,1,0,0] },
  { id:"t2",  name:"D&A",        domain:"Digital",           pillar:"efficiency",     type:"product", startCol:0, spanCols:8, users:[0,1,1,0,0] },
  { id:"t3",  name:"UM",         domain:"Digital",           pillar:"efficiency",     type:"product", startCol:0, spanCols:8, users:[0,1,1,0,0] },
  { id:"t4",  name:"Platform",   domain:"Member Enrollment", pillar:"efficiency",     type:"product", startCol:0, spanCols:3, users:[2,0,0,0,0] },
  { id:"t5",  name:"FCC",        domain:"Digital",           pillar:"efficiency",     type:"product", startCol:5, spanCols:3, users:[2,0,0,0] },
  { id:"t6",  name:"GPS",        domain:"Member Enrollment", pillar:"efficiency",     type:"product", startCol:5, spanCols:3, users:[2,0,0,0] },
  { id:"t7",  name:"Infra",      domain:"Claims",            pillar:"efficiency",     type:"product", startCol:5, spanCols:3, users:[0,2,1,0,0] },
  { id:"t8",  name:"Enablers",   domain:"Digital",           pillar:"efficiency",     type:"product", startCol:5, spanCols:3, users:[2,0,0,0] },
  { id:"t9",  name:"PopHealth",  domain:"Digital",           pillar:"efficiency",     type:"product", startCol:5, spanCols:3, users:[2,0,0,0] },
  { id:"t10", name:"QNXT",       domain:"Digital",           pillar:"efficiency",     type:"product", startCol:5, spanCols:3, users:[0,2,2,0] },
  // Technology investments (dashed bars) - Enterprise Value
  { id:"t11", name:"Real-Time Event Streaming",               domain:"All", pillar:"enterprise-value", type:"tech", startCol:0, spanCols:5, users:[0,2,2,2,0,0] },
  { id:"t12", name:"Modern Data Platform Migration",          domain:"All", pillar:"enterprise-value", type:"tech", startCol:0, spanCols:5, users:[0,2,2,2,0,0] },
  { id:"t13", name:"Agile Delivery Standardization",          domain:"All", pillar:"enterprise-value", type:"tech", startCol:0, spanCols:5, users:[0,2,2,0,2,0] },
  { id:"t14", name:"Medicaid Data & Analytics Modernization", domain:"All", pillar:"enterprise-value", type:"tech", startCol:0, spanCols:5, users:[0,2,2,2,0,0] },
  { id:"t15", name:"Portal consolidation",                    domain:"Provider", pillar:"enterprise-value", type:"tech", startCol:2, spanCols:3, users:[0,2,2,0] },
  { id:"t16", name:"Central DB",                              domain:"Provider", pillar:"enterprise-value", type:"tech", startCol:2, spanCols:3, users:[0,2,2,0] },
  { id:"t17", name:"Code cleanup",                            domain:"Enterprise", pillar:"enterprise-value", type:"tech", startCol:2, spanCols:3, users:[0,2,2,0] },
  { id:"t18", name:"Standardize the application stack (.Net to Java/Py)", domain:"Integration", pillar:"enterprise-value", type:"tech", startCol:3, spanCols:2, users:[2,0] },
  // Growth
  { id:"t19", name:"Member Self-Service Portal",     domain:"Digital",  pillar:"growth", type:"product", startCol:0, spanCols:4, users:[2,1,0,0,0] },
  { id:"t20", name:"Provider Network Expansion",    domain:"Network",  pillar:"growth", type:"product", startCol:1, spanCols:5, users:[0,2,1,0,0] },
  // Cost Reduction
  { id:"t21", name:"Automated Claims Processing",    domain:"Claims",     pillar:"cost-reduction", type:"tech", startCol:0, spanCols:5, users:[0,2,2,2,0,0] },
  { id:"t22", name:"Legacy System Decommission",     domain:"Enterprise", pillar:"cost-reduction", type:"tech", startCol:2, spanCols:6, users:[0,1,2,2,0] },
  { id:"t23", name:"Shared Service Bus",             domain:"Integration",pillar:"cost-reduction", type:"tech", startCol:1, spanCols:4, users:[2,1,1,0,0] },
  { id:"t24", name:"RPA for Manual Workflows",       domain:"Ops",        pillar:"cost-reduction", type:"product", startCol:0, spanCols:3, users:[2,0,0,0,0] },
  { id:"t25", name:"Cloud Cost Optimization",        domain:"Infra",      pillar:"cost-reduction", type:"tech", startCol:3, spanCols:5, users:[0,2,2,1,0] },
  { id:"t26", name:"Vendor Consolidation Program",   domain:"Enterprise", pillar:"cost-reduction", type:"product", startCol:4, spanCols:4, users:[0,0,2,1,0] },
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
  const [rallyFeatures, setRallyFeatures] = useState([]);
  const [rallyLoading,  setRallyLoading]  = useState(false);
  const [rallyFilter,   setRallyFilter]   = useState({ product:"all", pillar:"all", vs:"all", release:"all", project:"all", search:"" });
  const [timelinePanel, setTimelinePanel] = useState(false);
  const [timelineFilter, setTimelineFilter] = useState({ pillar:"all", type:"all", domain:"all", phase:"all", search:"" });
  const [rallyPanel, setRallyPanel] = useState(false);
  const fileInputRef = useRef(null);

  // Persist to localStorage
  useEffect(() => {
    try { const r = localStorage.getItem("roadmap-items-v2"); if (r) setItems(JSON.parse(r)); } catch {}
  }, []);
  useEffect(() => {
    try { localStorage.setItem("roadmap-items-v2", JSON.stringify(items)); } catch {}
  }, [items]);

  // Load Rally features on mount
  useEffect(() => {
    fetch("/api/rally/features").then(r => r.ok ? r.json() : []).then(setRallyFeatures).catch(() => {});
  }, []);

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
    { id:"rally",     icon:"◆", label:"Rally Roadmap" },
    { id:"timeline",  icon:"━", label:"3-Year Timeline" },
    { id:"strategy",  icon:"◇", label:"Strategy & Org" },
    { id:"investments", icon:"$", label:"Investments" },
    { id:"team",      icon:"◎", label:"Team & Reviews" },
    { id:"stratsite", icon:"◈", label:"Strategy Site" },
    { id:"storyboard", icon:"▤", label:"Storyboard" },
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

          {/* ── 3-YEAR TIMELINE ───────────────────────────────────────────── */}
          {view==="timeline" && (() => {
            const tf = timelineFilter;
            const allDomains = [...new Set(TIMELINE_ITEMS.map(i=>i.domain))].sort();
            const filteredItems = TIMELINE_ITEMS.filter(item => {
              if (tf.pillar !== "all" && item.pillar !== tf.pillar) return false;
              if (tf.type !== "all" && item.type !== tf.type) return false;
              if (tf.domain !== "all" && item.domain !== tf.domain) return false;
              if (tf.phase !== "all") {
                const phaseStart = tf.phase === "now" ? 0 : tf.phase === "next" ? 2 : 5;
                const phaseEnd = tf.phase === "now" ? 2 : tf.phase === "next" ? 5 : 8;
                if (item.startCol >= phaseEnd || item.startCol + item.spanCols <= phaseStart) return false;
              }
              if (tf.search) {
                const q = tf.search.toLowerCase();
                if (!`${item.name} ${item.domain} ${item.pillar}`.toLowerCase().includes(q)) return false;
              }
              return true;
            });

            const pillarStats = TIMELINE_PILLARS.map(p => {
              const items = filteredItems.filter(i => i.pillar === p.id);
              return { ...p, feat: items.filter(i=>i.type==="product").length, tech: items.filter(i=>i.type==="tech").length };
            });
            const totalCols = TIMELINE_QUARTERS.length;
            const activeFilterCount = [tf.pillar, tf.type, tf.domain, tf.phase].filter(v=>v!=="all").length + (tf.search ? 1 : 0);

            const UserDots = ({ users }) => (
              <div style={{ display:"flex", gap:2, alignItems:"center" }}>
                {(users||[]).map((u, i) => (
                  <span key={i} style={{ width:8, height:8, borderRadius:"50%", display:"inline-block",
                    background: u===2 ? ["#1565C0","#2E7D32","#6A1B9A","#B71C1C","#E65100"][i%5] : u===1 ? ["#1565C0","#2E7D32","#6A1B9A","#B71C1C","#E65100"][i%5] : "transparent",
                    border: u===0 ? "1.5px solid #B0B8C4" : "none",
                    opacity: u===1 ? 0.5 : 1,
                  }} />
                ))}
              </div>
            );

            return (
            <div style={{ flex:1, display:"flex", overflow:"hidden" }}>
              {/* Main timeline content */}
              <div style={{ flex:1, overflow:"auto", padding:28, background:"#F5F4F0" }}>
                {/* Header */}
                <div style={{ marginBottom:20 }}>
                  <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:6 }}>
                    <h1 style={{ fontSize:22, fontWeight:700, color:"#1C1C2E", margin:0 }}>Aetna Medicaid Technology 3 Year Roadmap</h1>
                    <div style={{ display:"flex", gap:8 }}>
                      <button onClick={()=>setView("dashboard")} style={{ fontSize:12, padding:"6px 14px", borderRadius:8, border:"1px solid #E4E2DA", background:"#FFF", color:"#1C1C2E", cursor:"pointer", fontWeight:500 }}>Home</button>
                      <button onClick={()=>setView("rally")} style={{ fontSize:12, padding:"6px 14px", borderRadius:8, border:"1px solid #E4E2DA", background:"#FFF", color:"#1C1C2E", cursor:"pointer", fontWeight:500 }}>Feature Details →</button>
                      <button onClick={()=>setTimelinePanel(!timelinePanel)} style={{ fontSize:12, padding:"6px 14px", borderRadius:8, border: timelinePanel ? "1.5px solid #7D3F98" : "1px solid #E4E2DA", background: timelinePanel ? "#F5F0FA" : "#FFF", color: timelinePanel ? "#7D3F98" : "#1C1C2E", cursor:"pointer", fontWeight:600, display:"flex", alignItems:"center", gap:5 }}>
                        <span style={{ fontSize:14 }}>⚙</span> Filters{activeFilterCount > 0 && <span style={{ background:"#7D3F98", color:"#FFF", fontSize:10, fontWeight:700, borderRadius:10, padding:"1px 6px", marginLeft:2 }}>{activeFilterCount}</span>}
                      </button>
                    </div>
                  </div>
                  <p style={{ fontSize:13, color:"#5A6170", margin:0 }}>Product themes and technology investments aligned to strategic pillars — click product bars to drill into features — Last updated: May 29, 2026</p>
                </div>

                {/* Pillar stats */}
                <div style={{ display:"flex", gap:12, marginBottom:20 }}>
                  {pillarStats.map(p => (
                    <div key={p.id} style={{ flex:1, background:"#0D1629", borderRadius:10, padding:"12px 16px", display:"flex", alignItems:"center", gap:10, cursor:"pointer", border: tf.pillar===p.id ? `2px solid ${p.dot}` : "2px solid transparent", transition:"border 0.15s" }}
                      onClick={() => setTimelineFilter({...tf, pillar: tf.pillar===p.id ? "all" : p.id})}>
                      <div style={{ width:8, height:8, borderRadius:"50%", background:p.dot }} />
                      <span style={{ fontSize:11, color:"rgba(255,255,255,0.8)", fontWeight:500 }}>{p.label}</span>
                      <span style={{ fontSize:18, fontWeight:700, color:"#FFF", marginLeft:6 }}>{p.feat}</span>
                      <span style={{ fontSize:10, color:"rgba(255,255,255,0.5)" }}>feat</span>
                      <span style={{ fontSize:10, color:"rgba(255,255,255,0.3)", margin:"0 2px" }}>·</span>
                      <span style={{ fontSize:18, fontWeight:700, color:"#FFF" }}>{p.tech}</span>
                      <span style={{ fontSize:10, color:"rgba(255,255,255,0.5)" }}>tech</span>
                    </div>
                  ))}
                </div>

                {/* Legend */}
                <div style={{ display:"flex", alignItems:"center", gap:20, marginBottom:16, padding:"10px 16px", background:"#FFF", borderRadius:8, border:"1px solid #E4E2DA" }}>
                  <span style={{ fontSize:11, fontWeight:600, color:"#6B6B8A", textTransform:"uppercase", letterSpacing:"0.06em" }}>Legend</span>
                  <div style={{ display:"flex", alignItems:"center", gap:6 }}>
                    <div style={{ width:28, height:14, borderRadius:4, background:"#E3ECFA", border:"1.5px solid #1565C0" }} />
                    <span style={{ fontSize:11, color:"#4A4A6A" }}>Product capability (Rally)</span>
                  </div>
                  <div style={{ display:"flex", alignItems:"center", gap:6 }}>
                    <div style={{ width:28, height:14, borderRadius:4, background:"#F8F8FA", border:"1.5px dashed #8A8AA0" }} />
                    <span style={{ fontSize:11, color:"#4A4A6A" }}>Technology investment (Architecture)</span>
                  </div>
                  <div style={{ display:"flex", alignItems:"center", gap:6 }}>
                    <div style={{ width:8, height:8, borderRadius:"50%", background:"#1565C0" }} />
                    <span style={{ fontSize:11, color:"#4A4A6A" }}>Explicit user</span>
                  </div>
                  <div style={{ display:"flex", alignItems:"center", gap:6 }}>
                    <div style={{ width:8, height:8, borderRadius:"50%", border:"1.5px solid #8A8AA0" }} />
                    <span style={{ fontSize:11, color:"#4A4A6A" }}>Implied user</span>
                  </div>
                  <span style={{ marginLeft:"auto", fontSize:11, color:"#8A8AA0" }}>Showing {filteredItems.length} of {TIMELINE_ITEMS.length}</span>
                </div>

                {/* Timeline chart */}
                <div style={{ background:"#FFF", borderRadius:12, border:"1px solid #E4E2DA", overflow:"hidden" }}>
                  {/* Phase header (Now / Next / Future) */}
                  <div style={{ display:"grid", gridTemplateColumns:`repeat(${totalCols}, 1fr)`, borderBottom:"2px solid #1C1C2E" }}>
                    {TIMELINE_PHASES.map((phase, pi) => (
                      <div key={pi} style={{ gridColumn:`span ${phase.cols}`, textAlign:"center", padding:"10px 0", fontWeight:700, fontSize:13, color:"#1C1C2E", borderRight: pi < TIMELINE_PHASES.length-1 ? "1px solid #E4E2DA" : "none", background:"#F8F7F4" }}>
                        {phase.label}
                      </div>
                    ))}
                  </div>

                  {/* Quarter labels */}
                  <div style={{ display:"grid", gridTemplateColumns:`repeat(${totalCols}, 1fr)`, borderBottom:"1px solid #E4E2DA", position:"relative" }}>
                    {TIMELINE_QUARTERS.map((q, i) => (
                      <div key={i} style={{ textAlign:"center", padding:"8px 0", fontSize:11, color:"#6B6B8A", fontWeight:500, borderRight: i < totalCols-1 ? "1px solid #F0EEE8" : "none" }}>
                        {q}
                      </div>
                    ))}
                  </div>

                  {/* Pillar sections with items */}
                  {TIMELINE_PILLARS.map(pillar => {
                    const pillarItems = filteredItems.filter(i => i.pillar === pillar.id);
                    if (pillarItems.length === 0) return null;
                    return (
                      <div key={pillar.id} style={{ borderBottom:"1px solid #E4E2DA" }}>
                        {/* Pillar label on left */}
                        <div style={{ display:"flex" }}>
                          <div style={{ width:140, flexShrink:0, padding:"16px 14px", borderRight:"1px solid #E4E2DA", display:"flex", flexDirection:"column", justifyContent:"center" }}>
                            <div style={{ fontSize:12, fontWeight:700, color:pillar.color }}>{pillar.label}</div>
                            <div style={{ fontSize:10, color:"#8A8AA0", marginTop:2 }}>{pillar.sub}</div>
                          </div>
                          {/* Items grid area */}
                          <div style={{ flex:1, position:"relative", minHeight: pillarItems.length * 40 + 16 }}>
                            {/* Vertical grid lines */}
                            <div style={{ position:"absolute", inset:0, display:"grid", gridTemplateColumns:`repeat(${totalCols}, 1fr)`, pointerEvents:"none" }}>
                              {TIMELINE_QUARTERS.map((_, i) => (
                                <div key={i} style={{ borderRight: i < totalCols-1 ? "1px solid #F5F4F0" : "none" }} />
                              ))}
                            </div>
                            {/* Bars */}
                            <div style={{ position:"relative", padding:"8px 0" }}>
                              {pillarItems.map((item, idx) => {
                                const leftPct = (item.startCol / totalCols) * 100;
                                const widthPct = (item.spanCols / totalCols) * 100;
                                const isProduct = item.type === "product";
                                return (
                                  <div key={item.id} style={{ position:"relative", height:32, marginBottom:4, paddingLeft:8, paddingRight:8 }}>
                                    <div className="card-hover" style={{
                                      position:"absolute", top:2, left:`calc(${leftPct}% + 4px)`, width:`calc(${widthPct}% - 8px)`, height:28,
                                      borderRadius:6,
                                      background: isProduct ? pillar.color+"14" : "#F8F8FA",
                                      border: isProduct ? `1.5px solid ${pillar.color}50` : "1.5px dashed #B0B8C4",
                                      display:"flex", alignItems:"center", padding:"0 10px", gap:8, cursor:"pointer",
                                    }}>
                                      <span style={{ fontSize:12, fontWeight:600, color:"#1C1C2E", whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis" }}>{item.name}</span>
                                      <span style={{ fontSize:10, padding:"1px 6px", borderRadius:4, background:"#0D1629", color:"#FFF", fontWeight:500, whiteSpace:"nowrap", flexShrink:0 }}>{item.domain}</span>
                                      <div style={{ marginLeft:"auto", flexShrink:0 }}>
                                        <UserDots users={item.users} />
                                      </div>
                                    </div>
                                  </div>
                                );
                              })}
                            </div>
                          </div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>

              {/* ── FILTER PANEL (slides in from right) ──────────────────── */}
              {timelinePanel && (
                <div style={{ width:320, borderLeft:"1px solid #E4E2DA", background:"#FFF", display:"flex", flexDirection:"column", overflow:"hidden", flexShrink:0, animation:"fadeUp 0.2s ease" }}>
                  {/* Panel header */}
                  <div style={{ padding:"16px 20px", borderBottom:"1px solid #E4E2DA", display:"flex", alignItems:"center", justifyContent:"space-between", background:"#F8F7F4" }}>
                    <div>
                      <div style={{ fontSize:14, fontWeight:700, color:"#1C1C2E" }}>Filter Timeline</div>
                      <div style={{ fontSize:11, color:"#8A8AA0", marginTop:2 }}>{filteredItems.length} of {TIMELINE_ITEMS.length} items visible</div>
                    </div>
                    <button onClick={()=>setTimelinePanel(false)} style={{ background:"none", border:"none", fontSize:18, cursor:"pointer", color:"#8A8AA0", lineHeight:1 }}>×</button>
                  </div>

                  <div style={{ flex:1, overflow:"auto", padding:20 }}>
                    {/* Search */}
                    <div style={{ marginBottom:20 }}>
                      <label style={{ fontSize:11, fontWeight:600, color:"#6B6B8A", textTransform:"uppercase", letterSpacing:"0.06em", display:"block", marginBottom:6 }}>Search</label>
                      <input type="search" placeholder="Search items..."
                        value={tf.search} onChange={e=>setTimelineFilter({...tf, search:e.target.value})}
                        style={{ width:"100%", padding:"8px 12px", borderRadius:8, border:"1px solid #E4E2DA", fontSize:12 }} />
                    </div>

                    {/* Strategic Pillar */}
                    <div style={{ marginBottom:20 }}>
                      <label style={{ fontSize:11, fontWeight:600, color:"#6B6B8A", textTransform:"uppercase", letterSpacing:"0.06em", display:"block", marginBottom:8 }}>Strategic Pillar</label>
                      <div style={{ display:"flex", flexDirection:"column", gap:4 }}>
                        <button onClick={()=>setTimelineFilter({...tf, pillar:"all"})}
                          style={{ display:"flex", alignItems:"center", gap:8, padding:"8px 12px", borderRadius:8, border: tf.pillar==="all" ? "1.5px solid #7D3F98" : "1px solid #E4E2DA", background: tf.pillar==="all" ? "#F5F0FA" : "#FFF", cursor:"pointer", fontSize:12, fontWeight: tf.pillar==="all" ? 600 : 400, color: tf.pillar==="all" ? "#7D3F98" : "#4A4A6A", textAlign:"left" }}>
                          All Pillars <span style={{ marginLeft:"auto", fontSize:11, color:"#8A8AA0" }}>{TIMELINE_ITEMS.length}</span>
                        </button>
                        {TIMELINE_PILLARS.map(p => {
                          const count = TIMELINE_ITEMS.filter(i=>i.pillar===p.id).length;
                          return (
                            <button key={p.id} onClick={()=>setTimelineFilter({...tf, pillar: tf.pillar===p.id ? "all" : p.id})}
                              style={{ display:"flex", alignItems:"center", gap:8, padding:"8px 12px", borderRadius:8, border: tf.pillar===p.id ? `1.5px solid ${p.color}` : "1px solid #E4E2DA", background: tf.pillar===p.id ? p.color+"10" : "#FFF", cursor:"pointer", fontSize:12, fontWeight: tf.pillar===p.id ? 600 : 400, color: tf.pillar===p.id ? p.color : "#4A4A6A", textAlign:"left" }}>
                              <div style={{ width:8, height:8, borderRadius:"50%", background:p.color, flexShrink:0 }} />
                              {p.label} <span style={{ marginLeft:"auto", fontSize:11, color:"#8A8AA0" }}>{count}</span>
                            </button>
                          );
                        })}
                      </div>
                    </div>

                    {/* Item Type */}
                    <div style={{ marginBottom:20 }}>
                      <label style={{ fontSize:11, fontWeight:600, color:"#6B6B8A", textTransform:"uppercase", letterSpacing:"0.06em", display:"block", marginBottom:8 }}>Item Type</label>
                      <div style={{ display:"flex", flexDirection:"column", gap:4 }}>
                        {[{id:"all",label:"All Types",icon:null},{id:"product",label:"Product Capability",icon:"▪"},{id:"tech",label:"Technology Investment",icon:"┄"}].map(t => (
                          <button key={t.id} onClick={()=>setTimelineFilter({...tf, type: tf.type===t.id ? "all" : t.id})}
                            style={{ display:"flex", alignItems:"center", gap:8, padding:"8px 12px", borderRadius:8, border: tf.type===t.id ? "1.5px solid #7D3F98" : "1px solid #E4E2DA", background: tf.type===t.id ? "#F5F0FA" : "#FFF", cursor:"pointer", fontSize:12, fontWeight: tf.type===t.id ? 600 : 400, color: tf.type===t.id ? "#7D3F98" : "#4A4A6A", textAlign:"left" }}>
                            {t.icon && <span style={{ fontSize:10, opacity:0.6 }}>{t.icon}</span>}
                            {t.label}
                            <span style={{ marginLeft:"auto", fontSize:11, color:"#8A8AA0" }}>{t.id==="all" ? TIMELINE_ITEMS.length : TIMELINE_ITEMS.filter(i=>i.type===t.id).length}</span>
                          </button>
                        ))}
                      </div>
                    </div>

                    {/* Domain */}
                    <div style={{ marginBottom:20 }}>
                      <label style={{ fontSize:11, fontWeight:600, color:"#6B6B8A", textTransform:"uppercase", letterSpacing:"0.06em", display:"block", marginBottom:8 }}>Domain</label>
                      <div style={{ display:"flex", flexDirection:"column", gap:4 }}>
                        <button onClick={()=>setTimelineFilter({...tf, domain:"all"})}
                          style={{ display:"flex", alignItems:"center", gap:8, padding:"8px 12px", borderRadius:8, border: tf.domain==="all" ? "1.5px solid #7D3F98" : "1px solid #E4E2DA", background: tf.domain==="all" ? "#F5F0FA" : "#FFF", cursor:"pointer", fontSize:12, fontWeight: tf.domain==="all" ? 600 : 400, color: tf.domain==="all" ? "#7D3F98" : "#4A4A6A", textAlign:"left" }}>
                          All Domains <span style={{ marginLeft:"auto", fontSize:11, color:"#8A8AA0" }}>{TIMELINE_ITEMS.length}</span>
                        </button>
                        {allDomains.map(d => {
                          const count = TIMELINE_ITEMS.filter(i=>i.domain===d).length;
                          return (
                            <button key={d} onClick={()=>setTimelineFilter({...tf, domain: tf.domain===d ? "all" : d})}
                              style={{ display:"flex", alignItems:"center", gap:8, padding:"8px 12px", borderRadius:8, border: tf.domain===d ? "1.5px solid #7D3F98" : "1px solid #E4E2DA", background: tf.domain===d ? "#F5F0FA" : "#FFF", cursor:"pointer", fontSize:12, fontWeight: tf.domain===d ? 600 : 400, color: tf.domain===d ? "#7D3F98" : "#4A4A6A", textAlign:"left" }}>
                              <span style={{ fontSize:10, padding:"1px 6px", borderRadius:4, background:"#0D1629", color:"#FFF", fontWeight:500 }}>{d}</span>
                              <span style={{ marginLeft:"auto", fontSize:11, color:"#8A8AA0" }}>{count}</span>
                            </button>
                          );
                        })}
                      </div>
                    </div>

                    {/* Time Phase */}
                    <div style={{ marginBottom:20 }}>
                      <label style={{ fontSize:11, fontWeight:600, color:"#6B6B8A", textTransform:"uppercase", letterSpacing:"0.06em", display:"block", marginBottom:8 }}>Time Phase</label>
                      <div style={{ display:"flex", flexDirection:"column", gap:4 }}>
                        {[{id:"all",label:"All Phases"},{id:"now",label:"Now (Q3-Q4 '26)"},{id:"next",label:"Next (Q1-Q3 '27)"},{id:"future",label:"Future (Q4 '27+)"}].map(p => (
                          <button key={p.id} onClick={()=>setTimelineFilter({...tf, phase: tf.phase===p.id ? "all" : p.id})}
                            style={{ display:"flex", alignItems:"center", gap:8, padding:"8px 12px", borderRadius:8, border: tf.phase===p.id ? "1.5px solid #7D3F98" : "1px solid #E4E2DA", background: tf.phase===p.id ? "#F5F0FA" : "#FFF", cursor:"pointer", fontSize:12, fontWeight: tf.phase===p.id ? 600 : 400, color: tf.phase===p.id ? "#7D3F98" : "#4A4A6A", textAlign:"left" }}>
                            {p.label}
                          </button>
                        ))}
                      </div>
                    </div>

                    {/* Reset button */}
                    {activeFilterCount > 0 && (
                      <button onClick={()=>setTimelineFilter({ pillar:"all", type:"all", domain:"all", phase:"all", search:"" })}
                        style={{ width:"100%", padding:"10px", borderRadius:8, border:"1px solid #DC2626", background:"#FEF2F2", color:"#DC2626", cursor:"pointer", fontSize:12, fontWeight:600 }}>
                        Clear All Filters ({activeFilterCount})
                      </button>
                    )}
                  </div>
                </div>
              )}
            </div>
            );
          })()}

          {/* ── STRATEGY & ORG ──────────────────────────────────────────── */}
          {view==="strategy" && (
            <div style={{ flex:1, overflow:"auto", padding:28, background:"#F5F4F0" }}>
              {/* Hero */}
              <div style={{ background:"linear-gradient(135deg, #0f1629, #1a2744, #1e3a5f)", borderRadius:14, padding:"28px 32px", marginBottom:20, color:"#FFF" }}>
                <div style={{ fontSize:10, fontWeight:700, color:"#7D3F98", letterSpacing:"0.14em", textTransform:"uppercase", marginBottom:4 }}>Aetna Medicaid DDAT</div>
                <h1 style={{ fontSize:22, fontWeight:700, margin:"0 0 8px" }}>Strategic Context & Organization</h1>
                <p style={{ fontSize:13, color:"rgba(255,255,255,0.7)", margin:0, lineHeight:1.5 }}>From projects to products — transforming Medicaid technology into a unified, product-driven platform</p>
              </div>

              {/* Story Arc */}
              <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr 1fr", gap:12, marginBottom:20 }}>
                {[
                  { phase:"Phase 1", title:"Growth & Fragmentation", desc:"Rapid expansion via market wins; organic tech growth to ~350 apps across ~9 product lines", color:"#B71C1C" },
                  { phase:"Phase 2", title:"Scale Without Integration", desc:"More capabilities layered on without consolidation; technical debt accelerated", color:"#E65100" },
                  { phase:"Phase 3", title:"Complexity Peaks", desc:"Frequent leadership changes; rising costs; limited ROI visibility across portfolio", color:"#F9A825" },
                  { phase:"Phase 4", title:"Turning Point (Now)", desc:"Platform consolidation, product operating model, cloud-native strategy underway", color:"#2E7D32" },
                ].map((p, i) => (
                  <div key={i} style={{ background:"#FFF", borderRadius:10, border:"1px solid #E4E2DA", borderTop:`3px solid ${p.color}`, padding:16 }}>
                    <div style={{ fontSize:10, fontWeight:700, color:p.color, textTransform:"uppercase", letterSpacing:"0.06em", marginBottom:4 }}>{p.phase}</div>
                    <div style={{ fontSize:13, fontWeight:700, color:"#1C1C2E", marginBottom:6 }}>{p.title}</div>
                    <div style={{ fontSize:12, color:"#5A6170", lineHeight:1.5 }}>{p.desc}</div>
                  </div>
                ))}
              </div>

              {/* Vision */}
              <div style={{ background:"#FFF", borderRadius:12, border:"1px solid #E4E2DA", padding:20, marginBottom:20 }}>
                <div style={{ fontSize:11, fontWeight:700, color:"#7D3F98", textTransform:"uppercase", letterSpacing:"0.08em", marginBottom:12 }}>3-Year Vision</div>
                <div style={{ display:"flex", alignItems:"center", gap:16, marginBottom:16 }}>
                  <div style={{ flex:1, background:"#FEF2F2", borderRadius:8, padding:14, textAlign:"center" }}>
                    <div style={{ fontSize:11, fontWeight:600, color:"#B71C1C", marginBottom:4 }}>FROM</div>
                    <div style={{ fontSize:13, fontWeight:700, color:"#1C1C2E" }}>Projects → Fragmentation → Cost Center</div>
                  </div>
                  <div style={{ fontSize:20, color:"#7D3F98", fontWeight:700 }}>→</div>
                  <div style={{ flex:1, background:"#F0FAF0", borderRadius:8, padding:14, textAlign:"center" }}>
                    <div style={{ fontSize:11, fontWeight:600, color:"#2E7D32", marginBottom:4 }}>TO</div>
                    <div style={{ fontSize:13, fontWeight:700, color:"#1C1C2E" }}>Products → Platforms → Value Engine</div>
                  </div>
                </div>
                <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10 }}>
                  {["Shift to shared enterprise platforms (Aetna Health, OneCM, QNXT modernization)",
                    "Cloud-native, API-first, event-driven architecture (Azure-based)",
                    "Product operating model with OKRs and outcome-based funding",
                    "Unified data + integration strategy (Snowflake, GCP, ATC)"].map((s, i) => (
                    <div key={i} style={{ display:"flex", alignItems:"flex-start", gap:8, padding:"8px 12px", background:"#F8F7F4", borderRadius:6, fontSize:12, color:"#4A4A6A", lineHeight:1.45 }}>
                      <span style={{ color:"#7D3F98", fontWeight:700, flexShrink:0 }}>✦</span>{s}
                    </div>
                  ))}
                </div>
              </div>

              <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:16 }}>
                {/* Leadership */}
                <div style={{ background:"#FFF", borderRadius:12, border:"1px solid #E4E2DA", padding:20 }}>
                  <div style={{ fontSize:11, fontWeight:700, color:"#7D3F98", textTransform:"uppercase", letterSpacing:"0.08em", marginBottom:14 }}>Leadership Team</div>
                  {[
                    { name:"Greg Krause", role:"AVP, Software Development Engineering", domain:"Overall Medicaid Technology Lead" },
                    { name:"Jeffrey Krasner", role:"Lead Director, Product Management", domain:"Product operating model, OKRs, roadmaps" },
                    { name:"Manjusha Nair", role:"Exec Director, Engineering", domain:"Claims & Encounters, Tech Operations" },
                    { name:"Rajesh Johnson", role:"Exec Director, Engineering", domain:"Member Enrollment & Strategic Programs" },
                    { name:"Paul Cavanaugh", role:"Exec Director, Engineering", domain:"Digital, Clinical Innovation, Data Integration" },
                    { name:"Swati Nanda", role:"Lead Director, Architecture", domain:"Solution Architecture, blueprints" },
                    { name:"Vikas Sharma", role:"Lead Director, Engineering", domain:"Finance Reporting, Provider Payments, Data" },
                    { name:"Lan Thanh Hoang", role:"Exec Director, Engineering", domain:"Provider & Growth Initiatives" },
                  ].map((l, i) => (
                    <div key={i} style={{ display:"flex", alignItems:"center", gap:10, padding:"8px 0", borderBottom: i < 7 ? "1px solid #F0EEE8" : "none" }}>
                      <div style={{ width:32, height:32, borderRadius:8, background:"#F5F0FA", display:"flex", alignItems:"center", justifyContent:"center", fontSize:12, fontWeight:700, color:"#7D3F98", flexShrink:0 }}>{l.name.split(" ").map(w=>w[0]).join("")}</div>
                      <div style={{ flex:1 }}>
                        <div style={{ fontSize:12, fontWeight:600, color:"#1C1C2E" }}>{l.name}</div>
                        <div style={{ fontSize:10, color:"#8A8AA0" }}>{l.role}</div>
                      </div>
                      <div style={{ fontSize:10, color:"#5A6170", maxWidth:160, textAlign:"right" }}>{l.domain}</div>
                    </div>
                  ))}
                </div>

                {/* 8 Product Lines */}
                <div style={{ background:"#FFF", borderRadius:12, border:"1px solid #E4E2DA", padding:20 }}>
                  <div style={{ fontSize:11, fontWeight:700, color:"#7D3F98", textTransform:"uppercase", letterSpacing:"0.08em", marginBottom:14 }}>8 Proposed Product Lines</div>
                  {[
                    { num:1, name:"Member Access & Self-Service", north:"Every member can enroll, understand benefits, and act digitally", owner:"Scott Waller" },
                    { num:2, name:"Enrollment & Eligibility Platform", north:"Real-time enrollment activation, zero fallout, full compliance", owner:"Rajesh Johnson" },
                    { num:3, name:"Claims & Encounters", north:"Single standardized claims platform, real-time adjudication", owner:"Manjusha Nair" },
                    { num:4, name:"Provider & Network", north:"Providers onboard and interact with minimal friction", owner:"Lan Thanh Hoang" },
                    { num:5, name:"Clinical & Care Management", north:"Proactive, data-driven care management", owner:"Mark Buckley" },
                    { num:6, name:"Finance & Billing", north:"Automated financial ops, near real-time reconciliation", owner:"Vikas Sharma" },
                    { num:7, name:"Data & Analytics Platform", north:"One trusted data platform, real-time insights", owner:"Swati Nanda" },
                    { num:8, name:"Contract & Compliance", north:"Automated contract/regulatory execution", owner:"Lan Thanh Hoang" },
                  ].map((pl, i) => (
                    <div key={i} style={{ padding:"10px 0", borderBottom: i < 7 ? "1px solid #F0EEE8" : "none" }}>
                      <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:3 }}>
                        <span style={{ width:20, height:20, borderRadius:4, background:"#7D3F98", color:"#FFF", fontSize:10, fontWeight:700, display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0 }}>{pl.num}</span>
                        <span style={{ fontSize:12, fontWeight:600, color:"#1C1C2E" }}>{pl.name}</span>
                        <span style={{ fontSize:10, color:"#8A8AA0", marginLeft:"auto" }}>{pl.owner}</span>
                      </div>
                      <div style={{ fontSize:11, color:"#5A6170", paddingLeft:28, lineHeight:1.4 }}>{pl.north}</div>
                    </div>
                  ))}
                </div>
              </div>

              {/* Users & Technology Layers */}
              <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:16, marginTop:16 }}>
                <div style={{ background:"#FFF", borderRadius:12, border:"1px solid #E4E2DA", padding:20 }}>
                  <div style={{ fontSize:11, fontWeight:700, color:"#7D3F98", textTransform:"uppercase", letterSpacing:"0.08em", marginBottom:14 }}>User Types</div>
                  {[
                    { type:"Member", desc:"The Medicaid enrollee" },
                    { type:"Caregiver", desc:"Family/support for the member" },
                    { type:"Provider", desc:"Physicians, facilities" },
                    { type:"Case Manager", desc:"Coordinates care" },
                    { type:"Clinician", desc:"Clinical decision-making" },
                  ].map((u, i) => (
                    <div key={i} style={{ display:"flex", alignItems:"center", gap:10, padding:"8px 0", borderBottom: i < 4 ? "1px solid #F0EEE8" : "none" }}>
                      <div style={{ width:28, height:28, borderRadius:"50%", background:"#F5F0FA", display:"flex", alignItems:"center", justifyContent:"center", fontSize:11, fontWeight:700, color:"#7D3F98" }}>{i+1}</div>
                      <div>
                        <div style={{ fontSize:12, fontWeight:600, color:"#1C1C2E" }}>{u.type}</div>
                        <div style={{ fontSize:10, color:"#8A8AA0" }}>{u.desc}</div>
                      </div>
                    </div>
                  ))}
                </div>
                <div style={{ background:"#FFF", borderRadius:12, border:"1px solid #E4E2DA", padding:20 }}>
                  <div style={{ fontSize:11, fontWeight:700, color:"#7D3F98", textTransform:"uppercase", letterSpacing:"0.08em", marginBottom:14 }}>Technology Layers</div>
                  {[
                    { layer:"Integration Platform", desc:"The connective tissue", color:"#1565C0" },
                    { layer:"Transactional", desc:"Claims, Constituent Service, Clinical, Network & Provider", color:"#2E7D32" },
                    { layer:"Core Admin", desc:"Enrollment, Benefits, Plan, Health Plan Admin", color:"#6A1B9A" },
                    { layer:"Data & Insights", desc:"Data Platforms and Data Sources", color:"#B71C1C" },
                  ].map((t, i) => (
                    <div key={i} style={{ display:"flex", alignItems:"center", gap:10, padding:"10px 0", borderBottom: i < 3 ? "1px solid #F0EEE8" : "none" }}>
                      <div style={{ width:4, height:32, borderRadius:2, background:t.color }} />
                      <div>
                        <div style={{ fontSize:12, fontWeight:600, color:"#1C1C2E" }}>{t.layer}</div>
                        <div style={{ fontSize:10, color:"#8A8AA0" }}>{t.desc}</div>
                      </div>
                    </div>
                  ))}
                </div>
              </div>

              {/* Terminology */}
              <div style={{ background:"#FFF", borderRadius:12, border:"1px solid #E4E2DA", padding:20, marginTop:16 }}>
                <div style={{ fontSize:11, fontWeight:700, color:"#7D3F98", textTransform:"uppercase", letterSpacing:"0.08em", marginBottom:14 }}>Terminology Reference</div>
                <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:8 }}>
                  {[
                    ["QNXT","Core claims/enrollment admin platform"],
                    ["GPS","Guided Personal Services (Salesforce)"],
                    ["Aetna Health","Enterprise member portal"],
                    ["OneCM","Enterprise case management platform"],
                    ["HEDIS","Quality measurement standard"],
                    ["STARS","CMS quality rating system"],
                    ["ABX","Cost structure & efficiency initiative"],
                    ["ATC","Real-time data streaming layer"],
                    ["MDH","Medicaid Data Hub API layer"],
                    ["FHIR","Healthcare interoperability standard"],
                    ["FCC","Family Care Central"],
                    ["834","EDI enrollment/eligibility standard"],
                    ["OKR","Objectives & Key Results"],
                    ["DDAT","Digital, Data & Analytics Technology"],
                    ["Duals","Members in both Medicare & Medicaid"],
                  ].map(([term, def], i) => (
                    <div key={i} style={{ padding:"6px 10px", background:"#F8F7F4", borderRadius:6 }}>
                      <span style={{ fontSize:11, fontWeight:700, color:"#7D3F98", marginRight:6 }}>{term}</span>
                      <span style={{ fontSize:10, color:"#5A6170" }}>{def}</span>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          )}

          {/* ── INVESTMENTS ──────────────────────────────────────────────── */}
          {view==="investments" && (
            <div style={{ flex:1, overflow:"auto", padding:28, background:"#F5F4F0" }}>
              {/* Hero */}
              <div style={{ background:"linear-gradient(135deg, #0f1629, #1a2744, #1e3a5f)", borderRadius:14, padding:"28px 32px", marginBottom:20, color:"#FFF" }}>
                <div style={{ fontSize:10, fontWeight:700, color:"#7D3F98", letterSpacing:"0.14em", textTransform:"uppercase", marginBottom:4 }}>Capital Portfolio</div>
                <h1 style={{ fontSize:22, fontWeight:700, margin:"0 0 8px" }}>FY2026–2027 Investment Overview</h1>
                <p style={{ fontSize:13, color:"rgba(255,255,255,0.7)", margin:0 }}>Total projected EBIT benefit (2026–2031): $676.6M</p>
              </div>

              {/* Top-line stats */}
              <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr 1fr", gap:12, marginBottom:20 }}>
                {[
                  { label:"2026 CapEx", value:"$45.3M", color:"#1565C0" },
                  { label:"2027 CapEx", value:"$67.9M", color:"#2E7D32" },
                  { label:"2026 Budget", value:"$105.2M", color:"#6A1B9A" },
                  { label:"FTE Count", value:"~125", color:"#B71C1C" },
                ].map((s, i) => (
                  <div key={i} style={{ background:"#FFF", borderRadius:10, border:"1px solid #E4E2DA", borderLeft:`3px solid ${s.color}`, padding:"16px 20px" }}>
                    <div style={{ fontSize:10, fontWeight:600, color:"#8A8AA0", textTransform:"uppercase", letterSpacing:"0.06em", marginBottom:4 }}>{s.label}</div>
                    <div style={{ fontSize:24, fontWeight:700, color:"#1C1C2E" }}>{s.value}</div>
                  </div>
                ))}
              </div>

              <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:16, marginBottom:16 }}>
                {/* Carryover Investments */}
                <div style={{ background:"#FFF", borderRadius:12, border:"1px solid #E4E2DA", padding:20 }}>
                  <div style={{ fontSize:11, fontWeight:700, color:"#7D3F98", textTransform:"uppercase", letterSpacing:"0.08em", marginBottom:4 }}>Carryover Investments</div>
                  <div style={{ fontSize:10, color:"#8A8AA0", marginBottom:14 }}>~$43.52M current 2027 CapEx</div>
                  {[
                    { name:"Aetna Health & Innovation Platform", amount:"$9.13M", desc:"Member experience modernization" },
                    { name:"Medicaid Contractual Projects", amount:"$8.44M", desc:"Compliance (Final Rule/Duals, OB3)" },
                    { name:"Claims Carry-Over", amount:"$5.23M", desc:"Universal Claim Tracker, Xten Upgrade, Fee Schedule" },
                    { name:"Enrollment Carry-Over", amount:"$4.69M", desc:"End-to-end automation, eligibility processing" },
                    { name:"Claims GPS Carry-Over", amount:"$4.04M", desc:"Moving Medicaid to GPS (Salesforce)" },
                    { name:"Claims Workflow Management", amount:"$3.92M", desc:"QNXT Upgrade, remedy replacement" },
                    { name:"Medicaid Clinical Carry-Over", amount:"$3.47M", desc:"Prior Auth, care coordination" },
                  ].map((inv, i) => (
                    <div key={i} style={{ padding:"10px 0", borderBottom: i < 6 ? "1px solid #F0EEE8" : "none" }}>
                      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:2 }}>
                        <span style={{ fontSize:12, fontWeight:600, color:"#1C1C2E" }}>{inv.name}</span>
                        <span style={{ fontSize:12, fontWeight:700, color:"#7D3F98" }}>{inv.amount}</span>
                      </div>
                      <div style={{ fontSize:11, color:"#5A6170" }}>{inv.desc}</div>
                    </div>
                  ))}
                </div>

                {/* New 2027 Investments */}
                <div style={{ background:"#FFF", borderRadius:12, border:"1px solid #E4E2DA", padding:20 }}>
                  <div style={{ fontSize:11, fontWeight:700, color:"#2E7D32", textTransform:"uppercase", letterSpacing:"0.08em", marginBottom:4 }}>New 2027 Investments</div>
                  <div style={{ fontSize:10, color:"#8A8AA0", marginBottom:14 }}>~$20.12M new CapEx</div>
                  {[
                    { name:"Aetna Medicaid Quality Optimization", amount:"$9.70M", desc:"HEDIS gaps, STARS, AI-driven quality", highlight:true },
                    { name:"Medicaid Cloud Optimization", amount:"$5.44M", desc:"Legacy to Azure cloud-native migration" },
                    { name:"Premium Billing Modernization", amount:"$0.83M", desc:"End-to-end billing capability" },
                    { name:"Member Domain AI/Automation", amount:"$0.73M", desc:"AI/ML for enrollment, 834, pharmacy reconciliation" },
                    { name:"AI Data Abstractions for HEDIS", amount:"$0.50M", desc:"Extract clinical insights from unstructured data" },
                  ].map((inv, i) => (
                    <div key={i} style={{ padding:"10px 0", borderBottom: i < 4 ? "1px solid #F0EEE8" : "none", background: inv.highlight ? "#F0FAF0" : "transparent", borderRadius: inv.highlight ? 6 : 0, paddingLeft: inv.highlight ? 10 : 0, paddingRight: inv.highlight ? 10 : 0 }}>
                      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:2 }}>
                        <span style={{ fontSize:12, fontWeight:600, color:"#1C1C2E" }}>{inv.name}</span>
                        <span style={{ fontSize:12, fontWeight:700, color:"#2E7D32" }}>{inv.amount}</span>
                      </div>
                      <div style={{ fontSize:11, color:"#5A6170" }}>{inv.desc}</div>
                    </div>
                  ))}

                  <div style={{ marginTop:20, borderTop:"1px solid #E4E2DA", paddingTop:16 }}>
                    <div style={{ fontSize:11, fontWeight:700, color:"#B71C1C", textTransform:"uppercase", letterSpacing:"0.08em", marginBottom:4 }}>ABX Initiatives</div>
                    <div style={{ fontSize:10, color:"#8A8AA0", marginBottom:10 }}>Cost structure focus — ~$12.8M 2027</div>
                    {["Medicaid Cloud Optimization","Triangulation","QNXT Consolidation","Pursuant HRQ In-House Migration","Enrollment Carryover"].map((a, i) => (
                      <div key={i} style={{ fontSize:11, color:"#4A4A6A", padding:"4px 0", display:"flex", alignItems:"center", gap:6 }}>
                        <span style={{ width:4, height:4, borderRadius:2, background:"#B71C1C", flexShrink:0 }} />{a}
                      </div>
                    ))}
                  </div>
                </div>
              </div>

              {/* Historical Spend */}
              <div style={{ background:"#FFF", borderRadius:12, border:"1px solid #E4E2DA", padding:20, marginBottom:16 }}>
                <div style={{ fontSize:11, fontWeight:700, color:"#7D3F98", textTransform:"uppercase", letterSpacing:"0.08em", marginBottom:14 }}>Historical Spend Trend</div>
                <div style={{ display:"flex", alignItems:"flex-end", gap:16, height:120 }}>
                  {[
                    { year:"FY2023", val:74.9, max:105.2 },
                    { year:"FY2024", val:64.1, max:105.2 },
                    { year:"FY2025", val:82.7, max:105.2 },
                    { year:"2026 1RF", val:103.5, max:105.2 },
                    { year:"2026 Budget", val:105.2, max:105.2 },
                  ].map((d, i) => (
                    <div key={i} style={{ flex:1, display:"flex", flexDirection:"column", alignItems:"center", gap:4 }}>
                      <span style={{ fontSize:11, fontWeight:700, color:"#1C1C2E" }}>${d.val}M</span>
                      <div style={{ width:"100%", height:`${(d.val/d.max)*80}px`, background: i >= 3 ? "linear-gradient(180deg, #7D3F98, #5c2d82)" : "#E4E2DA", borderRadius:"4px 4px 0 0", transition:"height 0.3s" }} />
                      <span style={{ fontSize:9, color:"#8A8AA0", fontWeight:500 }}>{d.year}</span>
                    </div>
                  ))}
                </div>
              </div>

              {/* Key Initiatives In-Flight */}
              <div style={{ background:"#FFF", borderRadius:12, border:"1px solid #E4E2DA", padding:20, marginBottom:16 }}>
                <div style={{ fontSize:11, fontWeight:700, color:"#7D3F98", textTransform:"uppercase", letterSpacing:"0.08em", marginBottom:14 }}>Key Strategic Initiatives In-Flight</div>
                <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:8 }}>
                  {[
                    { code:"A", name:"Aetna Health (Medicaid)", desc:"Member portal migration" },
                    { code:"B", name:"Family Care Central", desc:"Complex care portal" },
                    { code:"C", name:"Guided Personal Services (GPS)", desc:"Salesforce-based service" },
                    { code:"D", name:"Your Aetna Virtual Assistant", desc:"AI-powered member support" },
                    { code:"E", name:"OneCM Platform (SFHC)", desc:"Case Management modernization" },
                    { code:"F", name:"Network & Provider Modernization", desc:"Provider platform upgrade" },
                    { code:"G", name:"Data Platform (GCP/Snowflake)", desc:"Unified data strategy" },
                  ].map((init, i) => (
                    <div key={i} style={{ display:"flex", alignItems:"center", gap:10, padding:"10px 12px", background:"#F8F7F4", borderRadius:8 }}>
                      <span style={{ width:24, height:24, borderRadius:6, background:"#7D3F98", color:"#FFF", fontSize:11, fontWeight:700, display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0 }}>{init.code}</span>
                      <div>
                        <div style={{ fontSize:12, fontWeight:600, color:"#1C1C2E" }}>{init.name}</div>
                        <div style={{ fontSize:10, color:"#8A8AA0" }}>{init.desc}</div>
                      </div>
                      <span style={{ marginLeft:"auto", fontSize:10, padding:"2px 8px", borderRadius:10, background:"#E6F4EA", color:"#1E7E34", fontWeight:600 }}>In-Flight</span>
                    </div>
                  ))}
                </div>
              </div>

              {/* Critical Deadlines */}
              <div style={{ background:"#FFF", borderRadius:12, border:"1px solid #E4E2DA", padding:20 }}>
                <div style={{ fontSize:11, fontWeight:700, color:"#B71C1C", textTransform:"uppercase", letterSpacing:"0.08em", marginBottom:14 }}>Critical Deadlines</div>
                {[
                  { item:"Claims Remedy Replacement", deadline:"Jan 1, 2027", note:"Day-1 scope lock + stakeholder sign-off", severity:"critical" },
                  { item:"QNXT 2025.R1 Upgrade", deadline:"Active", note:"Both Claims and Enrollment teams", severity:"high" },
                  { item:"TX STAR CHP Implementation", deadline:"Completed 4/1", note:"Post-KY 1115 Waiver", severity:"done" },
                  { item:"Open Enrollment / Welcome Season", deadline:"Ongoing", note:"Operational readiness priority", severity:"high" },
                ].map((d, i) => (
                  <div key={i} style={{ display:"flex", alignItems:"center", gap:12, padding:"10px 0", borderBottom: i < 3 ? "1px solid #F0EEE8" : "none" }}>
                    <span style={{ fontSize:11, padding:"3px 8px", borderRadius:6, fontWeight:700, background: d.severity==="critical"?"#FEE2E2":d.severity==="done"?"#E6F4EA":"#FEF3C7", color: d.severity==="critical"?"#DC2626":d.severity==="done"?"#1E7E34":"#B45309" }}>{d.deadline}</span>
                    <div style={{ flex:1 }}>
                      <div style={{ fontSize:12, fontWeight:600, color:"#1C1C2E" }}>{d.item}</div>
                      <div style={{ fontSize:10, color:"#8A8AA0" }}>{d.note}</div>
                    </div>
                  </div>
                ))}
              </div>

              {/* Vendor Ecosystem */}
              <div style={{ background:"#FFF", borderRadius:12, border:"1px solid #E4E2DA", padding:20, marginTop:16 }}>
                <div style={{ fontSize:11, fontWeight:700, color:"#7D3F98", textTransform:"uppercase", letterSpacing:"0.08em", marginBottom:14 }}>Vendor Ecosystem</div>
                <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr 1fr", gap:8 }}>
                  {[
                    { vendor:"Microsoft Azure", role:"Primary cloud platform" },
                    { vendor:"Salesforce Health Cloud", role:"GPS platform" },
                    { vendor:"Infosys", role:"Product Support Services" },
                    { vendor:"Cognizant", role:"QNXT platform support" },
                    { vendor:"edifecs (Cotiviti)", role:"Interoperability" },
                    { vendor:"Availity", role:"Provider portal / HIN" },
                    { vendor:"Echo Payments", role:"Payment platform" },
                  ].map((v, i) => (
                    <div key={i} style={{ padding:"10px 12px", background:"#F8F7F4", borderRadius:8, textAlign:"center" }}>
                      <div style={{ fontSize:11, fontWeight:600, color:"#1C1C2E", marginBottom:2 }}>{v.vendor}</div>
                      <div style={{ fontSize:9, color:"#8A8AA0" }}>{v.role}</div>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          )}

          {/* ── RALLY ROADMAP ──────────────────────────────────────────── */}
          {view==="rally" && (() => {
            const rf = rallyFilter;
            const visible = rallyFeatures.filter(f => {
              if (rf.product !== "all" && f.product !== rf.product) return false;
              if (rf.pillar !== "all" && f.pillar !== rf.pillar) return false;
              if (rf.vs !== "all") { const fvs = f.valueSector||""; if (rf.vs === "(none)" ? fvs !== "" : fvs !== rf.vs) return false; }
              if (rf.release !== "all") { const fr = f.release||""; if (rf.release === "(none)" ? fr !== "" : fr !== rf.release) return false; }
              if (rf.project !== "all" && f.project !== rf.project) return false;
              if (rf.search) { const q = rf.search.toLowerCase(); if (!`${f.name} ${f.id} ${f.project} ${f.product} ${f.pillar} ${f.release} ${f.valueSector}`.toLowerCase().includes(q)) return false; }
              return true;
            });
            const laneCounts = { done:0, now:0, next:0, future:0 };
            visible.forEach(f => laneCounts[f.lane]++);

            const products = [...new Set(rallyFeatures.map(f=>f.product))].sort();
            const pillars = [...new Set(rallyFeatures.map(f=>f.pillar))].sort();
            const valueSectors = [...new Set(rallyFeatures.map(f=>f.valueSector||""))].filter(Boolean).sort();
            const releases = [...new Set(rallyFeatures.map(f=>f.release||""))].filter(Boolean).sort((a,b)=>a.localeCompare(b,undefined,{numeric:true}));
            const projects = [...new Set(rallyFeatures.map(f=>f.project))].sort();
            const activeFilterCount = [rf.product, rf.pillar, rf.vs, rf.release, rf.project].filter(v=>v!=="all").length + (rf.search ? 1 : 0);

            const LANE_META = {
              done:   { label:"Done",   sub:"Completed last 30 days", dot:"#43a047" },
              now:    { label:"Now",    sub:"In flight this PI",      dot:"#1a73e8" },
              next:   { label:"Next",   sub:"On deck next PI",        dot:"#6b4fd6" },
              future: { label:"Future", sub:"Exploring & defining",   dot:"#717784" },
            };
            const STATUS_COLORS = { "Done":"#1565c0", "In Progress":"#1e7e34", "Ready":"#1565c0", "Prepare":"#b84400", "Backlog":"#545b64", "Ask":"#545b64" };
            const PILLAR_COLORS = { "Enterprise Value":"#1565C0", "Efficiency":"#2E7D32", "Growth":"#6A1B9A", "Cost Reduction":"#B71C1C" };
            const PRODUCT_COLORS = { "Assign a Doc":"#1b5e20", "Enrollment Ops":"#0d47a1", "Extracts & ID Cards":"#e65100", "Duals/D-SNP":"#880e4f", "Enrollment Ancillary":"#4a148c", "Other":"#37474f" };

            return (
            <div style={{ flex:1, display:"flex", overflow:"hidden" }}>
              {/* Main content */}
              <div style={{ flex:1, overflow:"auto", padding:24 }}>
                {/* Stats header */}
                <div style={{ background:"linear-gradient(135deg, #0f1629, #1a2744, #1e3a5f)", borderRadius:14, padding:"20px 28px", marginBottom:16, color:"#FFF" }}>
                  <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:12 }}>
                    <div style={{ fontSize:16, fontWeight:700 }}>Medicaid Member Enrollment — Rally Roadmap</div>
                    <button onClick={()=>setRallyPanel(!rallyPanel)} style={{ fontSize:11, padding:"5px 14px", borderRadius:6, background: rallyPanel ? "rgba(125,63,152,0.3)" : "rgba(255,255,255,0.12)", color:"#FFF", border: rallyPanel ? "1px solid rgba(125,63,152,0.6)" : "1px solid rgba(255,255,255,0.2)", cursor:"pointer", fontWeight:600, display:"flex", alignItems:"center", gap:5 }}>
                      <span>⚙</span> Filters{activeFilterCount > 0 && <span style={{ background:"#7D3F98", fontSize:10, fontWeight:700, borderRadius:10, padding:"1px 6px" }}>{activeFilterCount}</span>}
                    </button>
                  </div>
                  <div style={{ display:"flex", gap:24, alignItems:"center" }}>
                    {Object.entries(LANE_META).map(([key, meta]) => (
                      <div key={key} style={{ textAlign:"center" }}>
                        <div style={{ fontSize:22, fontWeight:700 }}>{laneCounts[key]}</div>
                        <div style={{ fontSize:10, color:"rgba(255,255,255,0.7)", textTransform:"uppercase", letterSpacing:"0.08em" }}>{meta.label}</div>
                      </div>
                    ))}
                    <div style={{ marginLeft:"auto", fontSize:11, color:"rgba(255,255,255,0.6)" }}>
                      {visible.length} of {rallyFeatures.length} features
                      <button onClick={async () => { setRallyLoading(true); try { const r = await fetch("/api/rally/refresh",{method:"POST"}); const d = await r.json(); if (d.success) { const f2 = await fetch("/api/rally/features"); setRallyFeatures(await f2.json()); notify(`Refreshed ${d.count} features from Rally`); } else { notify(d.error||"Refresh failed","error"); } } catch(e){ notify(e.message,"error"); } setRallyLoading(false); }}
                        style={{ marginLeft:12, fontSize:11, padding:"4px 12px", borderRadius:6, background:"rgba(255,255,255,0.15)", color:"#FFF", border:"1px solid rgba(255,255,255,0.2)", cursor:"pointer" }}>
                        {rallyLoading ? "⟳ Refreshing..." : "↻ Refresh from Rally"}
                      </button>
                    </div>
                  </div>
                </div>

                {/* Search */}
                <div style={{ marginBottom:16 }}>
                  <input type="search" placeholder="Search features by name, ID, project, product..."
                    value={rf.search} onChange={e => setRallyFilter({...rf, search:e.target.value})}
                    style={{ width:"100%", padding:"10px 16px", borderRadius:10, border:"1px solid #E4E2DA", fontSize:13, background:"#FFF" }} />
                </div>

                {/* Lanes grid */}
                <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr 1fr", gap:16, alignItems:"start" }}>
                  {["done","now","next","future"].map(laneKey => {
                    const meta = LANE_META[laneKey];
                    const STATE_ORDER = { "":0, "Backlog":1, "Ask":2, "Prepare":3, "Ready":4, "In Progress":5, "Done":6 };
                    const laneFeatures = visible.filter(f => f.lane===laneKey).sort((a,b) => (STATE_ORDER[b.state]||0) - (STATE_ORDER[a.state]||0));
                    return (
                      <div key={laneKey} style={{ background:"#FFF", borderRadius:14, border:"1px solid #E4E2DA", overflow:"hidden" }}>
                        <div style={{ padding:"14px 16px 10px", borderBottom:"1px solid #F0EEE8" }}>
                          <div style={{ display:"flex", alignItems:"center", gap:8 }}>
                            <div style={{ width:10, height:10, borderRadius:"50%", background:meta.dot }} />
                            <span style={{ fontSize:14, fontWeight:700, color:"#1C1C2E" }}>{meta.label}</span>
                            <span style={{ fontSize:12, color:"#8A8AA0", marginLeft:"auto" }}>{laneFeatures.length}</span>
                          </div>
                          <div style={{ fontSize:11, color:"#8A8AA0", marginTop:4 }}>{meta.sub}</div>
                        </div>
                        <div style={{ padding:12, maxHeight:600, overflow:"auto" }}>
                          {laneFeatures.length === 0 && <div style={{ textAlign:"center", padding:20, color:"#8A8AA0", fontSize:12 }}>No features</div>}
                          {laneFeatures.map(f => (
                            <div key={f.id} className="card-hover" style={{ background:"#FAFBFD", border:"1px solid #E9ECF3", borderLeft:`3px solid ${meta.dot}`, borderRadius:8, padding:"10px 12px", marginBottom:8, cursor:"default" }}>
                              <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:4 }}>
                                <span style={{ fontSize:10, padding:"2px 7px", borderRadius:10, background:(STATUS_COLORS[f.state]||"#545b64")+"18", color:STATUS_COLORS[f.state]||"#545b64", fontWeight:600, textTransform:"uppercase" }}>{f.state||"Defining"}</span>
                                <a href={f.rallyUrl} target="_blank" rel="noopener noreferrer" style={{ fontSize:11, color:"#1558b0", textDecoration:"none", fontFamily:"monospace", fontWeight:500 }}>{f.id}</a>
                              </div>
                              <div style={{ fontSize:12, fontWeight:600, color:"#1C1C2E", lineHeight:1.35, marginBottom:6, display:"-webkit-box", WebkitLineClamp:2, WebkitBoxOrient:"vertical", overflow:"hidden" }}>{f.name}</div>
                              <div style={{ display:"flex", flexWrap:"wrap", gap:4 }}>
                                <span style={{ fontSize:9, padding:"2px 6px", borderRadius:3, background:(PRODUCT_COLORS[f.product]||"#37474f")+"14", color:PRODUCT_COLORS[f.product]||"#37474f", fontWeight:600 }}>{f.product}</span>
                                <span style={{ fontSize:9, padding:"2px 6px", borderRadius:10, background:(PILLAR_COLORS[f.pillar]||"#1565C0")+"14", color:PILLAR_COLORS[f.pillar]||"#1565C0", fontWeight:700 }}>{f.pillar}</span>
                                {f.release && <span style={{ fontSize:9, padding:"2px 5px", borderRadius:3, background:"#f0f2f5", color:"#6B7280", fontWeight:600 }}>{f.release}</span>}
                                {f.date && <span style={{ fontSize:9, color:"#8A8AA0", marginLeft:"auto" }}>{f.date}</span>}
                              </div>
                            </div>
                          ))}
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>

              {/* ── FILTER PANEL (right side) ──────────────────────────────── */}
              {rallyPanel && (
                <div style={{ width:320, borderLeft:"1px solid #E4E2DA", background:"#FFF", display:"flex", flexDirection:"column", overflow:"hidden", flexShrink:0, animation:"fadeUp 0.2s ease" }}>
                  {/* Panel header */}
                  <div style={{ padding:"16px 20px", borderBottom:"1px solid #E4E2DA", display:"flex", alignItems:"center", justifyContent:"space-between", background:"#F8F7F4" }}>
                    <div>
                      <div style={{ fontSize:14, fontWeight:700, color:"#1C1C2E" }}>Filter Features</div>
                      <div style={{ fontSize:11, color:"#8A8AA0", marginTop:2 }}>{visible.length} of {rallyFeatures.length} features visible</div>
                    </div>
                    <button onClick={()=>setRallyPanel(false)} style={{ background:"none", border:"none", fontSize:18, cursor:"pointer", color:"#8A8AA0", lineHeight:1 }}>×</button>
                  </div>

                  <div style={{ flex:1, overflow:"auto", padding:20 }}>
                    {/* Search */}
                    <div style={{ marginBottom:20 }}>
                      <label style={{ fontSize:11, fontWeight:600, color:"#6B6B8A", textTransform:"uppercase", letterSpacing:"0.06em", display:"block", marginBottom:6 }}>Search</label>
                      <input type="search" placeholder="Search features..."
                        value={rf.search} onChange={e=>setRallyFilter({...rf, search:e.target.value})}
                        style={{ width:"100%", padding:"8px 12px", borderRadius:8, border:"1px solid #E4E2DA", fontSize:12 }} />
                    </div>

                    {/* Product */}
                    <div style={{ marginBottom:20 }}>
                      <label style={{ fontSize:11, fontWeight:600, color:"#6B6B8A", textTransform:"uppercase", letterSpacing:"0.06em", display:"block", marginBottom:8 }}>Product</label>
                      <div style={{ display:"flex", flexDirection:"column", gap:4 }}>
                        <button onClick={()=>setRallyFilter({...rf, product:"all"})}
                          style={{ display:"flex", alignItems:"center", gap:8, padding:"8px 12px", borderRadius:8, border: rf.product==="all" ? "1.5px solid #7D3F98" : "1px solid #E4E2DA", background: rf.product==="all" ? "#F5F0FA" : "#FFF", cursor:"pointer", fontSize:12, fontWeight: rf.product==="all" ? 600 : 400, color: rf.product==="all" ? "#7D3F98" : "#4A4A6A", textAlign:"left" }}>
                          All Products <span style={{ marginLeft:"auto", fontSize:11, color:"#8A8AA0" }}>{rallyFeatures.length}</span>
                        </button>
                        {products.map(p => {
                          const count = rallyFeatures.filter(f=>f.product===p).length;
                          const color = PRODUCT_COLORS[p]||"#37474f";
                          return (
                            <button key={p} onClick={()=>setRallyFilter({...rf, product: rf.product===p ? "all" : p})}
                              style={{ display:"flex", alignItems:"center", gap:8, padding:"8px 12px", borderRadius:8, border: rf.product===p ? `1.5px solid ${color}` : "1px solid #E4E2DA", background: rf.product===p ? color+"10" : "#FFF", cursor:"pointer", fontSize:12, fontWeight: rf.product===p ? 600 : 400, color: rf.product===p ? color : "#4A4A6A", textAlign:"left" }}>
                              <div style={{ width:8, height:8, borderRadius:2, background:color, flexShrink:0 }} />
                              {p} <span style={{ marginLeft:"auto", fontSize:11, color:"#8A8AA0" }}>{count}</span>
                            </button>
                          );
                        })}
                      </div>
                    </div>

                    {/* Strategic Pillar */}
                    <div style={{ marginBottom:20 }}>
                      <label style={{ fontSize:11, fontWeight:600, color:"#6B6B8A", textTransform:"uppercase", letterSpacing:"0.06em", display:"block", marginBottom:8 }}>Strategic Pillar</label>
                      <div style={{ display:"flex", flexDirection:"column", gap:4 }}>
                        <button onClick={()=>setRallyFilter({...rf, pillar:"all"})}
                          style={{ display:"flex", alignItems:"center", gap:8, padding:"8px 12px", borderRadius:8, border: rf.pillar==="all" ? "1.5px solid #7D3F98" : "1px solid #E4E2DA", background: rf.pillar==="all" ? "#F5F0FA" : "#FFF", cursor:"pointer", fontSize:12, fontWeight: rf.pillar==="all" ? 600 : 400, color: rf.pillar==="all" ? "#7D3F98" : "#4A4A6A", textAlign:"left" }}>
                          All Pillars <span style={{ marginLeft:"auto", fontSize:11, color:"#8A8AA0" }}>{rallyFeatures.length}</span>
                        </button>
                        {pillars.map(p => {
                          const count = rallyFeatures.filter(f=>f.pillar===p).length;
                          const color = PILLAR_COLORS[p]||"#1565C0";
                          return (
                            <button key={p} onClick={()=>setRallyFilter({...rf, pillar: rf.pillar===p ? "all" : p})}
                              style={{ display:"flex", alignItems:"center", gap:8, padding:"8px 12px", borderRadius:8, border: rf.pillar===p ? `1.5px solid ${color}` : "1px solid #E4E2DA", background: rf.pillar===p ? color+"10" : "#FFF", cursor:"pointer", fontSize:12, fontWeight: rf.pillar===p ? 600 : 400, color: rf.pillar===p ? color : "#4A4A6A", textAlign:"left" }}>
                              <div style={{ width:8, height:8, borderRadius:"50%", background:color, flexShrink:0 }} />
                              {p} <span style={{ marginLeft:"auto", fontSize:11, color:"#8A8AA0" }}>{count}</span>
                            </button>
                          );
                        })}
                      </div>
                    </div>

                    {/* PI / Release */}
                    <div style={{ marginBottom:20 }}>
                      <label style={{ fontSize:11, fontWeight:600, color:"#6B6B8A", textTransform:"uppercase", letterSpacing:"0.06em", display:"block", marginBottom:8 }}>PI / Release</label>
                      <div style={{ display:"flex", flexDirection:"column", gap:4 }}>
                        <button onClick={()=>setRallyFilter({...rf, release:"all"})}
                          style={{ display:"flex", alignItems:"center", gap:8, padding:"8px 12px", borderRadius:8, border: rf.release==="all" ? "1.5px solid #7D3F98" : "1px solid #E4E2DA", background: rf.release==="all" ? "#F5F0FA" : "#FFF", cursor:"pointer", fontSize:12, fontWeight: rf.release==="all" ? 600 : 400, color: rf.release==="all" ? "#7D3F98" : "#4A4A6A", textAlign:"left" }}>
                          All Releases <span style={{ marginLeft:"auto", fontSize:11, color:"#8A8AA0" }}>{rallyFeatures.length}</span>
                        </button>
                        {releases.map(r => {
                          const count = rallyFeatures.filter(f=>f.release===r).length;
                          return (
                            <button key={r} onClick={()=>setRallyFilter({...rf, release: rf.release===r ? "all" : r})}
                              style={{ display:"flex", alignItems:"center", gap:8, padding:"8px 12px", borderRadius:8, border: rf.release===r ? "1.5px solid #7D3F98" : "1px solid #E4E2DA", background: rf.release===r ? "#F5F0FA" : "#FFF", cursor:"pointer", fontSize:12, fontWeight: rf.release===r ? 600 : 400, color: rf.release===r ? "#7D3F98" : "#4A4A6A", textAlign:"left" }}>
                              {r} <span style={{ marginLeft:"auto", fontSize:11, color:"#8A8AA0" }}>{count}</span>
                            </button>
                          );
                        })}
                      </div>
                    </div>

                    {/* Value Sector */}
                    <div style={{ marginBottom:20 }}>
                      <label style={{ fontSize:11, fontWeight:600, color:"#6B6B8A", textTransform:"uppercase", letterSpacing:"0.06em", display:"block", marginBottom:8 }}>Value Sector</label>
                      <div style={{ display:"flex", flexDirection:"column", gap:4 }}>
                        <button onClick={()=>setRallyFilter({...rf, vs:"all"})}
                          style={{ display:"flex", alignItems:"center", gap:8, padding:"8px 12px", borderRadius:8, border: rf.vs==="all" ? "1.5px solid #7D3F98" : "1px solid #E4E2DA", background: rf.vs==="all" ? "#F5F0FA" : "#FFF", cursor:"pointer", fontSize:12, fontWeight: rf.vs==="all" ? 600 : 400, color: rf.vs==="all" ? "#7D3F98" : "#4A4A6A", textAlign:"left" }}>
                          All Sectors <span style={{ marginLeft:"auto", fontSize:11, color:"#8A8AA0" }}>{rallyFeatures.length}</span>
                        </button>
                        {valueSectors.map(vs => {
                          const count = rallyFeatures.filter(f=>(f.valueSector||"")===vs).length;
                          return (
                            <button key={vs} onClick={()=>setRallyFilter({...rf, vs: rf.vs===vs ? "all" : vs})}
                              style={{ display:"flex", alignItems:"center", gap:8, padding:"8px 12px", borderRadius:8, border: rf.vs===vs ? "1.5px solid #7D3F98" : "1px solid #E4E2DA", background: rf.vs===vs ? "#F5F0FA" : "#FFF", cursor:"pointer", fontSize:12, fontWeight: rf.vs===vs ? 600 : 400, color: rf.vs===vs ? "#7D3F98" : "#4A4A6A", textAlign:"left" }}>
                              {vs} <span style={{ marginLeft:"auto", fontSize:11, color:"#8A8AA0" }}>{count}</span>
                            </button>
                          );
                        })}
                      </div>
                    </div>

                    {/* Rally Project */}
                    <div style={{ marginBottom:20 }}>
                      <label style={{ fontSize:11, fontWeight:600, color:"#6B6B8A", textTransform:"uppercase", letterSpacing:"0.06em", display:"block", marginBottom:8 }}>Rally Project</label>
                      <div style={{ display:"flex", flexDirection:"column", gap:4 }}>
                        <button onClick={()=>setRallyFilter({...rf, project:"all"})}
                          style={{ display:"flex", alignItems:"center", gap:8, padding:"8px 12px", borderRadius:8, border: rf.project==="all" ? "1.5px solid #7D3F98" : "1px solid #E4E2DA", background: rf.project==="all" ? "#F5F0FA" : "#FFF", cursor:"pointer", fontSize:12, fontWeight: rf.project==="all" ? 600 : 400, color: rf.project==="all" ? "#7D3F98" : "#4A4A6A", textAlign:"left" }}>
                          All Projects <span style={{ marginLeft:"auto", fontSize:11, color:"#8A8AA0" }}>{rallyFeatures.length}</span>
                        </button>
                        {projects.map(p => {
                          const count = rallyFeatures.filter(f=>f.project===p).length;
                          return (
                            <button key={p} onClick={()=>setRallyFilter({...rf, project: rf.project===p ? "all" : p})}
                              style={{ display:"flex", alignItems:"center", gap:8, padding:"8px 12px", borderRadius:8, border: rf.project===p ? "1.5px solid #7D3F98" : "1px solid #E4E2DA", background: rf.project===p ? "#F5F0FA" : "#FFF", cursor:"pointer", fontSize:12, fontWeight: rf.project===p ? 600 : 400, color: rf.project===p ? "#7D3F98" : "#4A4A6A", textAlign:"left" }}>
                              {p} <span style={{ marginLeft:"auto", fontSize:11, color:"#8A8AA0" }}>{count}</span>
                            </button>
                          );
                        })}
                      </div>
                    </div>

                    {/* Reset */}
                    {activeFilterCount > 0 && (
                      <button onClick={()=>setRallyFilter({ product:"all", pillar:"all", vs:"all", release:"all", project:"all", search:"" })}
                        style={{ width:"100%", padding:"10px", borderRadius:8, border:"1px solid #DC2626", background:"#FEF2F2", color:"#DC2626", cursor:"pointer", fontSize:12, fontWeight:600 }}>
                        Clear All Filters ({activeFilterCount})
                      </button>
                    )}
                  </div>
                </div>
              )}
            </div>
            );
          })()}

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

          {view==="stratsite" && (
            <iframe src="/strategy-overview.html" style={{ flex:1, width:"100%", height:"calc(100vh - 56px)", border:"none" }} title="Strategy Site" />
          )}

          {view==="storyboard" && (
            <iframe src="/storyboard.html" style={{ flex:1, width:"100%", height:"calc(100vh - 56px)", border:"none" }} title="Storyboard" />
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
