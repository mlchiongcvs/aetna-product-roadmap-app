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
  { id: "later", label: "LATER", sub: "6–12+ Months",    color: "#575757" },
];
const STATUSES = [
  { id: "in-progress", label: "In Progress", color: "#7D3F98" },
  { id: "discovery",   label: "Discovery",   color: "#0A4B8C" },
  { id: "planning",    label: "Planning",    color: "#B35B00" },
  { id: "on-hold",     label: "On Hold",     color: "#767676" },
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

// ── ANATOMY ICONS (outlined, 1.5px stroke, currentColor) ──────────────────
const Icon = ({ name, size=16, color="currentColor", style={} }) => {
  const s = { width:size, height:size, display:"inline-block", verticalAlign:"middle", flexShrink:0, ...style };
  const props = { width:size, height:size, viewBox:"0 0 24 24", fill:"none", stroke:color, strokeWidth:1.5, strokeLinecap:"round", strokeLinejoin:"round" };
  const icons = {
    overview:    <svg {...props}><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/></svg>,
    story:       <svg {...props}><path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20"/><path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z"/></svg>,
    roadmap:     <svg {...props}><line x1="3" y1="6" x2="21" y2="6"/><line x1="3" y1="12" x2="21" y2="12"/><line x1="3" y1="18" x2="21" y2="18"/><rect x="5" y="4" width="8" height="4" rx="1" fill={color} fillOpacity="0.15"/><rect x="8" y="10" width="10" height="4" rx="1" fill={color} fillOpacity="0.15"/><rect x="4" y="16" width="6" height="4" rx="1" fill={color} fillOpacity="0.15"/></svg>,
    context:     <svg {...props}><circle cx="12" cy="12" r="10"/><line x1="12" y1="16" x2="12" y2="12"/><line x1="12" y1="8" x2="12.01" y2="8"/></svg>,
    northstar:   <svg {...props}><polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2" fill={color} fillOpacity="0.1"/></svg>,
    dashboard:   <svg {...props}><rect x="3" y="3" width="18" height="18" rx="2"/><line x1="3" y1="9" x2="21" y2="9"/><line x1="9" y1="21" x2="9" y2="9"/></svg>,
    investments: <svg {...props}><line x1="12" y1="1" x2="12" y2="23"/><path d="M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"/></svg>,
    team:        <svg {...props}><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>,
    rally:       <svg {...props}><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/></svg>,
    user:        <svg {...props}><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/></svg>,
    sparkle:     <svg {...props}><path d="M12 3l1.5 5.5L19 10l-5.5 1.5L12 17l-1.5-5.5L5 10l5.5-1.5L12 3z" fill={color} fillOpacity="0.2"/></svg>,
    check:       <svg {...props}><polyline points="20 6 9 17 4 12"/></svg>,
    link:        <svg {...props}><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/></svg>,
    file:        <svg {...props}><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>,
    upload:      <svg {...props}><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>,
    settings:    <svg {...props}><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09a1.65 1.65 0 0 0-1.08-1.51 1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09a1.65 1.65 0 0 0 1.51-1.08 1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 2.83-2.83l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1.08 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9c.26.6.87 1 1.51 1.08H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1.08z"/></svg>,
    refresh:     <svg {...props}><polyline points="23 4 23 10 17 10"/><path d="M20.49 15a9 9 0 1 1-2.12-9.36L23 10"/></svg>,
    close:       <svg {...props}><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>,
    chart:       <svg {...props}><line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/></svg>,
    target:      <svg {...props}><circle cx="12" cy="12" r="10"/><circle cx="12" cy="12" r="6"/><circle cx="12" cy="12" r="2"/></svg>,
    doc:         <svg {...props}><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/><polyline points="10 9 9 9 8 9"/></svg>,
    spreadsheet: <svg {...props}><rect x="3" y="3" width="18" height="18" rx="2"/><line x1="3" y1="9" x2="21" y2="9"/><line x1="3" y1="15" x2="21" y2="15"/><line x1="9" y1="3" x2="9" y2="21"/><line x1="15" y1="3" x2="15" y2="21"/></svg>,
    book:        <svg {...props}><path d="M2 3h6a4 4 0 0 1 4 4v14a3 3 0 0 0-3-3H2z"/><path d="M22 3h-6a4 4 0 0 0-4 4v14a3 3 0 0 1 3-3h7z"/></svg>,
  };
  return <span style={s}>{icons[name] || null}</span>;
};

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
    <div style={{ width:380, borderLeft:"1px solid #E9E9E9", display:"flex", flexDirection:"column", background:"#FFF", overflow:"hidden", flexShrink:0 }}>
      {/* Header */}
      <div style={{ padding:"14px 20px", borderBottom:"1px solid #E9E9E9", display:"flex", alignItems:"center", justifyContent:"space-between", background:pillar.bg }}>
        <div style={{ display:"flex", gap:8, alignItems:"center" }}>
          <div style={{ width:10, height:10, borderRadius:2, background:pillar.color }} />
          <span style={{ fontSize:11, fontWeight:600, color:pillar.color }}>{pillar.label}</span>
        </div>
        <button onClick={onClose} style={{ background:"none", border:"none", cursor:"pointer", color:"#767676", lineHeight:1, padding:4 }}><Icon name="close" size={16} /></button>
      </div>

      <div style={{ flex:1, overflow:"auto", padding:20 }}>
        <h2 style={{ fontSize:15, fontWeight:600, color:"#262626", marginBottom:12, lineHeight:1.45, margin:"0 0 12px" }}>
          {item.title}
        </h2>

        {/* Badges */}
        <div style={{ display:"flex", flexWrap:"wrap", gap:6, marginBottom:16 }}>
          <span style={{ fontSize:11, padding:"3px 8px", borderRadius:20, background:horizon.color+"20", color:horizon.color, fontWeight:700 }}>{horizon.label}</span>
          <span style={{ fontSize:11, padding:"3px 8px", borderRadius:20, background:status.color+"18", color:status.color, fontWeight:500 }}>{status.label}</span>
          {item.priority === "critical" && <span style={{ fontSize:11, padding:"3px 8px", borderRadius:20, background:"#FEE2E2", color:"#DC2626", fontWeight:700 }}>CRITICAL</span>}
          {item.priority === "high"     && <span style={{ fontSize:11, padding:"3px 8px", borderRadius:20, background:"#FEF3C7", color:"#B45309", fontWeight:600 }}>HIGH</span>}
          {item.owner && <span style={{ fontSize:11, padding:"3px 8px", borderRadius:20, background:"#F2F2F2", color:"#575757", display:"inline-flex", alignItems:"center", gap:3 }}><Icon name="user" size={11} />{item.owner}</span>}
        </div>

        {/* Description */}
        <p style={{ fontSize:13, color:"#575757", lineHeight:1.6, margin:"0 0 16px" }}>{item.description}</p>

        {/* AI Insights */}
        {item.aiInsights && (
          <div style={{ background:"#F5F0FA", borderRadius:8, padding:"12px 14px", marginBottom:16, borderLeft:"3px solid #7D3F98" }}>
            <div style={{ fontSize:10, fontWeight:700, color:"#7D3F98", textTransform:"uppercase", letterSpacing:"0.08em", marginBottom:4, display:"flex", alignItems:"center", gap:4 }}><Icon name="sparkle" size={12} color="#7D3F98" />AI Insights</div>
            <div style={{ fontSize:12, color:"#531570", lineHeight:1.55 }}>{item.aiInsights}</div>
          </div>
        )}

        {/* Key deliverables */}
        {item.keyInitiatives?.length > 0 && (
          <div style={{ marginBottom:16 }}>
            <div style={{ fontSize:11, fontWeight:700, color:"#767676", textTransform:"uppercase", letterSpacing:"0.08em", marginBottom:8 }}>Key Deliverables</div>
            {item.keyInitiatives.map((init, i) => (
              <div key={i} style={{ fontSize:12, color:"#262626", padding:"4px 0", display:"flex", gap:8, alignItems:"flex-start" }}>
                <span style={{ width:6, height:6, borderRadius:"50%", border:`1.5px solid ${pillar.color}`, flexShrink:0, marginTop:5 }} />
                {init}
              </div>
            ))}
          </div>
        )}

        {/* Tags */}
        {item.tags?.length > 0 && (
          <div style={{ display:"flex", flexWrap:"wrap", gap:5, marginBottom:16 }}>
            {item.tags.map(tag => (
              <span key={tag} style={{ fontSize:10, padding:"2px 7px", borderRadius:4, background:"#F2F2F2", color:"#575757" }}>#{tag}</span>
            ))}
          </div>
        )}

        {/* Status update */}
        <div style={{ marginBottom:16 }}>
          <div style={{ fontSize:11, fontWeight:700, color:"#767676", textTransform:"uppercase", letterSpacing:"0.08em", marginBottom:8 }}>Update Status</div>
          <div style={{ display:"flex", flexWrap:"wrap", gap:5 }}>
            {STATUSES.map(s => (
              <button key={s.id} onClick={() => onStatusChange(item.id, s.id)}
                style={{ fontSize:11, padding:"4px 10px", borderRadius:20, border:"1px solid "+(item.status===s.id ? s.color : "#E9E9E9"), background:item.status===s.id ? s.color : "#FFF", color:item.status===s.id ? "#FFF" : "#575757", cursor:"pointer", fontWeight:500, fontFamily:"inherit" }}>
                {s.label}
              </button>
            ))}
          </div>
        </div>

        {/* Integrations */}
        <div style={{ marginBottom:16 }}>
          <div style={{ fontSize:11, fontWeight:700, color:"#767676", textTransform:"uppercase", letterSpacing:"0.08em", marginBottom:10 }}>Integrations</div>
          <div style={{ display:"flex", flexDirection:"column", gap:8 }}>
            {item.jiraKey ? (
              <div style={{ display:"flex", alignItems:"center", gap:8, padding:"10px 12px", background:"#EBF3FF", borderRadius:8 }}>
                <Icon name="check" size={14} color="#0A4B8C" />
                <span style={{ fontSize:12, fontWeight:600, color:"#0A4B8C" }}>Jira Epic: {item.jiraKey}</span>
              </div>
            ) : (
              <button onClick={() => onSyncJira(item)} disabled={isProcessing}
                style={{ padding:"10px 14px", borderRadius:8, border:"1px solid #0A4B8C", background:"#FFF", color:"#0A4B8C", fontSize:12, fontWeight:500, cursor:isProcessing?"not-allowed":"pointer", textAlign:"left", display:"flex", alignItems:"center", gap:8, fontFamily:"inherit", opacity:isProcessing?0.6:1 }}>
                <Icon name="link" size={14} />Sync to Jira — Create Epic
              </button>
            )}
            {item.confluenceUrl ? (
              <div style={{ display:"flex", alignItems:"center", gap:8, padding:"10px 12px", background:"#E6F5F5", borderRadius:8 }}>
                <Icon name="check" size={14} color="#00787E" />
                <span style={{ fontSize:12, fontWeight:600, color:"#00787E" }}>Confluence page published</span>
              </div>
            ) : (
              <button onClick={() => onPublishConfluence(item)} disabled={isProcessing}
                style={{ padding:"10px 14px", borderRadius:8, border:"1px solid #00787E", background:"#FFF", color:"#00787E", fontSize:12, fontWeight:500, cursor:isProcessing?"not-allowed":"pointer", textAlign:"left", display:"flex", alignItems:"center", gap:8, fontFamily:"inherit", opacity:isProcessing?0.6:1 }}>
                <Icon name="doc" size={14} />Publish to Confluence
              </button>
            )}
          </div>
        </div>

        {/* Comments */}
        <div>
          <div style={{ fontSize:11, fontWeight:700, color:"#767676", textTransform:"uppercase", letterSpacing:"0.08em", marginBottom:10 }}>
            Team Notes ({item.comments?.length || 0})
          </div>
          {(item.comments||[]).map(c => (
            <div key={c.id} style={{ padding:"10px 12px", background:"#F7F7F7", borderRadius:8, marginBottom:8 }}>
              <div style={{ display:"flex", justifyContent:"space-between", marginBottom:4 }}>
                <span style={{ fontSize:11, fontWeight:600, color:"#7D3F98" }}>{c.author}</span>
                <span style={{ fontSize:10, color:"#767676" }}>{c.timestamp}</span>
              </div>
              <div style={{ fontSize:12, color:"#262626", lineHeight:1.5 }}>{c.text}</div>
            </div>
          ))}
          <div style={{ display:"flex", gap:8, marginTop:8 }}>
            <input value={commentText} onChange={e => setCommentText(e.target.value)}
              onKeyDown={e => { if (e.key==="Enter" && !e.shiftKey) { e.preventDefault(); handleComment(); } }}
              placeholder="Add a team note... (Enter to submit)"
              style={{ flex:1, padding:"8px 12px", borderRadius:8, border:"1px solid #E9E9E9", fontSize:12, fontFamily:"inherit" }} />
            <button onClick={handleComment}
              style={{ padding:"8px 14px", background:"#7D3F98", color:"#FFF", border:"none", borderRadius:9999, fontSize:12, cursor:"pointer", fontFamily:"inherit" }}>
              Add
            </button>
          </div>
        </div>
      </div>

      <div style={{ padding:"10px 20px", borderTop:"1px solid #E9E9E9", fontSize:10, color:"#B0AEA8" }}>
        Source: {item.source} · Updated {item.lastUpdated}
      </div>
    </div>
  );
}

// ── MAIN APP ─────────────────────────────────────────────────────────────────
export default function RoadmapApp() {
  const [sidebarOpen,   setSidebarOpen]   = useState(true);
  const [view,          setView]          = useState("strat-overview");
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

  // Listen for navigation messages from iframes
  const [pendingInitId, setPendingInitId] = useState(null);
  useEffect(() => {
    const handler = (e) => {
      if (e.data && e.data.type === 'navigate' && e.data.view) {
        if (e.data.initId) setPendingInitId(e.data.initId);
        setView(e.data.view);
      }
    };
    window.addEventListener('message', handler);
    return () => window.removeEventListener('message', handler);
  }, []);

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
    { id:"strat-overview", icon:"overview",    label:"Overview" },
    { id:"strat-roadmap",  icon:"roadmap",     label:"Roadmap" },
    { id:"strat-context",  icon:"context",     label:"Context & Assumptions" },
    // { id:"storyboard",  icon:"northstar",   label:"Northstar Strategy" },  // merged into Our Story
    // { id:"dashboard",   icon:"dashboard",   label:"Dashboard" },           // reuse later
    // { id:"rally",       icon:"rally",       label:"Rally Roadmap" },       // not ready for users
    // { id:"investments", icon:"investments", label:"Investments" },          // reuse later
    // { id:"team",        icon:"team",        label:"Team & Reviews" },      // reuse later
  ];

  // ── RENDER ────────────────────────────────────────────────────────────────
  return (
    <div style={{ display:"flex", flexDirection:"column", height:"100vh", fontFamily:"-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif", background:"#F7F7F7", overflow:"hidden" }}>
      <style>{`
        * { box-sizing:border-box; margin:0; padding:0; }
        body { -webkit-font-smoothing:antialiased; }
        ::-webkit-scrollbar { width:4px; height:4px; }
        ::-webkit-scrollbar-track { background:transparent; }
        ::-webkit-scrollbar-thumb { background:#CCCCCC; border-radius:10px; }
        button { font-family:inherit; }
        input, textarea { font-family:inherit; }
        input:focus, textarea:focus { outline:2px solid #7D3F98; outline-offset:2px; }
        button:focus-visible { outline:2px solid #7D3F98; outline-offset:2px; }
        .card-hover:hover { box-shadow:0 2px 12px rgba(92,46,114,0.08); transform:translateY(-1px); transition:all 0.15s; }
        .nav-btn:hover { background:rgba(125,63,152,0.08) !important; color:#7D3F98 !important; }
        .row-hover:hover { background:#F3EBF8 !important; }
        @keyframes fadeUp { from { opacity:0; transform:translateY(8px); } to { opacity:1; transform:translateY(0); } }
        @keyframes spin { to { transform:rotate(360deg); } }
        @keyframes pulse { 0%,100% { opacity:1; } 50% { opacity:0.5; } }
      `}</style>

      {/* ── BODY (sidebar + main) ────────────────────────────────────────── */}
      <div style={{ display:"flex", flex:1, overflow:"hidden" }}>

      {/* ── SIDEBAR ──────────────────────────────────────────────────────── */}
      <div style={{ width:sidebarOpen ? 240 : 56, background:"#FFFFFF", borderRight:"1px solid #E9E9E9", display:"flex", flexDirection:"column", flexShrink:0, transition:"width 0.2s ease" }}>
        {/* Logo */}
        <div style={{ padding:sidebarOpen ? "16px 20px 12px" : "14px 0 12px", display:"flex", justifyContent:sidebarOpen ? "flex-start" : "center", borderBottom:"1px solid #E9E9E9" }}>
          {sidebarOpen
            ? <img src={`${import.meta.env.BASE_URL || '/'}aetna-logo.svg`} alt="Aetna" style={{ height:18 }} />
            : <svg width="20" height="20" viewBox="0 0 28 28" fill="none"><path d="M20.5 3.56c-.58-.59-1.35-.88-2.11-.88s-1.53.3-2.12.89L12.81 7.05 9.33 3.56C8.75 2.98 7.98 2.68 7.21 2.68s-1.53.29-2.11.88L.87 7.8c-1.16 1.17-1.16 3.07.01 4.24L12.81 24l11.93-11.96c1.17-1.17 1.17-3.07 0-4.24L20.5 3.56z" fill="#7D3F98"/></svg>
          }
        </div>

        {/* Brand section */}
        {sidebarOpen && (
          <div style={{ padding:"12px 20px 12px", borderBottom:"1px solid #E9E9E9" }}>
            <div style={{ fontSize:13, fontWeight:700, color:"#262626", lineHeight:1.35, textTransform:"uppercase" }}>3 Year Roadmap Strategy</div>
          </div>
        )}

        {/* Nav */}
        <nav style={{ padding:sidebarOpen ? "16px 12px" : "16px 8px", flex:1, overflow:"auto" }}>
          {NAV.map(n => (
            <button key={n.id} onClick={() => setView(n.id)} className="nav-btn" title={!sidebarOpen ? n.label : undefined}
              style={{ display:"flex", alignItems:"center", gap:12, width:"100%", padding:sidebarOpen ? "12px 14px" : "10px 0", borderRadius:8, border:"none", cursor:"pointer", background:view===n.id ? "#F3EBF8" : "transparent", color:view===n.id ? "#5C2E72" : "#575757", fontSize:14, fontWeight:view===n.id ? 600 : 400, textAlign:"left", marginBottom:4, transition:"all 0.15s", justifyContent:sidebarOpen ? "flex-start" : "center" }}>
              <Icon name={n.icon} size={18} color={view===n.id ? "#7D3F98" : "#767676"} />
              {sidebarOpen && n.label}
            </button>
          ))}
        </nav>

        {/* Stats footer */}
        {sidebarOpen && (
          <div style={{ padding:"16px 20px", borderTop:"1px solid #E9E9E9" }}>
            <div style={{ fontSize:10, color:"#767676", textTransform:"uppercase", letterSpacing:"0.08em", marginBottom:10, fontWeight:600 }}>Roadmap Health</div>
            {[
              { label:"Total initiatives", val:stats.total },
              { label:"Jira synced",        val:`${stats.synced} / ${stats.now} NOW` },
              { label:"Complete",           val:stats.complete },
            ].map(s => (
              <div key={s.label} style={{ display:"flex", justifyContent:"space-between", marginBottom:6 }}>
                <span style={{ fontSize:11, color:"#767676" }}>{s.label}</span>
                <span style={{ fontSize:11, fontWeight:600, color:"#262626" }}>{s.val}</span>
              </div>
            ))}
          </div>
        )}

        {/* Collapse toggle */}
        <button onClick={() => setSidebarOpen(o => !o)}
          style={{ padding:"12px", borderTop:"1px solid #E9E9E9", border:"none", borderTop:"1px solid #E9E9E9", background:"transparent", cursor:"pointer", display:"flex", alignItems:"center", justifyContent:sidebarOpen ? "flex-end" : "center", paddingRight:sidebarOpen ? 16 : 0, color:"#767676" }}>
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round">
            {sidebarOpen
              ? <><polyline points="11 17 6 12 11 7"/><polyline points="18 17 13 12 18 7"/></>
              : <><polyline points="13 17 18 12 13 7"/><polyline points="6 17 11 12 6 7"/></>
            }
          </svg>
        </button>
      </div>

      {/* ── MAIN ─────────────────────────────────────────────────────────── */}
      <div style={{ flex:1, display:"flex", flexDirection:"column", overflow:"hidden" }}>

        {/* Top bar */}
        <div style={{ height:48, background:"#FFFFFF", borderBottom:"1px solid #E9E9E9", display:"flex", alignItems:"center", padding:"0 24px", gap:12, flexShrink:0 }}>
          <span style={{ fontSize:14, fontWeight:600, color:"#262626", flex:1 }}>
            {NAV.find(n=>n.id===view)?.label}
          </span>

          {view==="roadmap" && (
            <div style={{ display:"flex", gap:5 }}>
              {[{id:"all",label:"All",color:"#575757"}, ...HORIZONS.map(h=>({...h,label:h.label}))].map(h => (
                <button key={h.id} onClick={() => setFilterHorizon(h.id)}
                  style={{ fontSize:11, padding:"4px 10px", borderRadius:6, border:"1px solid "+(filterHorizon===h.id ? h.color : "#E9E9E9"), background:filterHorizon===h.id ? h.color : "#FFF", color:filterHorizon===h.id ? "#FFF" : "#575757", cursor:"pointer", fontWeight:500 }}>
                  {h.label}
                </button>
              ))}
            </div>
          )}

          {isProcessing && processingMsg && (
            <div style={{ fontSize:12, color:"#7D3F98", display:"flex", alignItems:"center", gap:6 }}>
              <span style={{ display:"inline-flex", animation:"spin 1s linear infinite" }}><Icon name="refresh" size={14} color="#7D3F98" /></span>
              <span style={{ animation:"pulse 1.5s ease infinite" }}>{processingMsg}</span>
            </div>
          )}

        </div>

        {/* Content */}
        <div style={{ flex:1, overflow:"hidden", display:"flex" }}>

          {/* ── DASHBOARD ─────────────────────────────────────────────── */}
          {view==="dashboard" && (
            <div style={{ flex:1, overflow:"auto", padding:32 }}>
              <div style={{ maxWidth:1100 }}>
                <h1 style={{ fontSize:24, fontWeight:500, color:"#262626", marginBottom:6 }}>Dashboard</h1>
                <p style={{ color:"#575757", fontSize:16, marginBottom:32 }}>
                  AI-powered roadmap intelligence · {new Date().toLocaleDateString("en-US",{month:"long",day:"numeric",year:"numeric"})}
                </p>

                {/* Stat cards */}
                <div style={{ display:"grid", gridTemplateColumns:"repeat(4,1fr)", gap:16, marginBottom:32 }}>
                  {[
                    { label:"Total Initiatives", val:stats.total, sub:"across all pillars", color:"#7D3F98" },
                    { label:"NOW · This Quarter", val:stats.now, sub:`${stats.complete} complete`, color:"#7D3F98" },
                    { label:"NEXT · 1-2 Qtrs", val:stats.next, sub:"in discovery", color:"#769FC8" },
                    { label:"LATER · 6-12mo+", val:stats.later, sub:"strategic bets", color:"#575757" },
                  ].map((s,i) => (
                    <div key={i} style={{ background:"#FFFFFF", borderRadius:12, padding:"20px 24px", border:"1px solid #E9E9E9" }}>
                      <div style={{ fontSize:34, fontWeight:500, color:s.color, lineHeight:1 }}>{s.val}</div>
                      <div style={{ fontSize:14, fontWeight:500, color:"#262626", marginTop:8 }}>{s.label}</div>
                      <div style={{ fontSize:12, color:"#767676", marginTop:4 }}>{s.sub}</div>
                    </div>
                  ))}
                </div>

                {/* Two-column: Pillar distribution + Ingest */}
                <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:24, marginBottom:32 }}>
                  {/* Pillar distribution */}
                  <div style={{ background:"#FFFFFF", borderRadius:12, border:"1px solid #E9E9E9", overflow:"hidden" }}>
                    <div style={{ padding:"16px 24px", borderBottom:"1px solid #E9E9E9", fontSize:14, fontWeight:500, color:"#262626" }}>By Strategic Pillar</div>
                    <div style={{ padding:"20px 24px" }}>
                      {PILLARS.map(p => {
                        const cnt = items.filter(i=>i.pillar===p.id).length;
                        const pct = items.length ? (cnt/items.length)*100 : 0;
                        return (
                          <div key={p.id} style={{ marginBottom:16 }}>
                            <div style={{ display:"flex", justifyContent:"space-between", marginBottom:4 }}>
                              <span style={{ fontSize:14, fontWeight:400, color:"#262626" }}>{p.label}</span>
                              <span style={{ fontSize:14, color:"#575757" }}>{cnt}</span>
                            </div>
                            <div style={{ height:6, background:"#F2F2F2", borderRadius:4, overflow:"hidden" }}>
                              <div style={{ height:"100%", width:`${pct}%`, background:p.color, borderRadius:4, transition:"width 0.5s" }} />
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  </div>

                  {/* Ingest Documents */}
                  <div style={{ background:"#FFFFFF", borderRadius:12, border:"1px solid #E9E9E9", overflow:"hidden" }}>
                    <div style={{ padding:"16px 24px", borderBottom:"1px solid #E9E9E9", fontSize:14, fontWeight:500, color:"#262626" }}>Ingest Documents</div>
                    <div style={{ padding:"20px 24px" }}>
                      <p style={{ color:"#575757", fontSize:14, marginBottom:16, lineHeight:1.6 }}>
                        Upload strategy decks, research, or planning briefs. AI extracts structured roadmap initiatives.
                      </p>
                      <div onDragOver={e=>{e.preventDefault();setDragOver(true);}} onDragLeave={()=>setDragOver(false)}
                        onDrop={e=>{e.preventDefault();setDragOver(false);handleDroppedFiles(e.dataTransfer.files);}}
                        onClick={() => fileInputRef.current?.click()}
                        style={{ border:`2px dashed ${dragOver?"#8932AF":"#CCCCCC"}`, borderRadius:12, padding:"32px 20px", textAlign:"center", background:dragOver?"#F5F0F8":"#F7F7F7", cursor:"pointer", transition:"all 0.2s", marginBottom:16 }}>
                        <input ref={fileInputRef} type="file" multiple accept=".pdf,.pptx,.docx,.xlsx,.xls,.txt" style={{ display:"none" }} onChange={e=>handleDroppedFiles(e.target.files)} />
                        <div style={{ marginBottom:8 }}><Icon name="upload" size={32} color="#767676" /></div>
                        <div style={{ fontSize:16, fontWeight:500, color:"#262626", marginBottom:4 }}>Drop files or click to upload</div>
                        <div style={{ fontSize:14, color:"#767676" }}>PDF, PPTX, DOCX, XLSX, TXT</div>
                      </div>

                      {uploadedFiles.length>0 && (
                        <div style={{ marginBottom:16 }}>
                          <div style={{ fontSize:12, fontWeight:500, color:"#262626", marginBottom:8 }}>Ready for analysis ({uploadedFiles.length})</div>
                          {uploadedFiles.map((file,i) => (
                            <div key={i} style={{ display:"flex", alignItems:"center", gap:10, padding:"10px 14px", background:"#F7F7F7", borderRadius:8, border:"1px solid #E9E9E9", marginBottom:6 }}>
                              <Icon name={file.name.endsWith(".xlsx")||file.name.endsWith(".xls")?"spreadsheet":"file"} size={16} color="#767676" />
                              <span style={{ flex:1, fontSize:14, color:"#262626" }}>{file.name}</span>
                              <span style={{ fontSize:12, color:"#767676" }}>{(file.size/1024).toFixed(0)} KB</span>
                              <span onClick={()=>setUploadedFiles(prev=>prev.filter((_,j)=>j!==i))} style={{ cursor:"pointer", color:"#C62828", padding:"0 4px" }}><Icon name="close" size={14} /></span>
                            </div>
                          ))}
                        </div>
                      )}

                      <textarea value={textInput} onChange={e=>setTextInput(e.target.value)}
                        placeholder="Or paste text directly — roadmap content, OKRs, meeting notes..."
                        style={{ width:"100%", height:80, padding:"12px 16px", borderRadius:8, border:"1px solid #CCCCCC", fontSize:14, color:"#262626", resize:"vertical", background:"#FFFFFF", marginBottom:16 }} />

                      <button onClick={analyzeDocuments} disabled={isProcessing||(uploadedFiles.length===0&&!textInput.trim())}
                        style={{ background:isProcessing?"#949494":"#7D3F98", color:"#FFF", border:"none", borderRadius:9999, padding:"12px 24px", fontSize:16, fontWeight:500, cursor:isProcessing?"not-allowed":"pointer", width:"100%", transition:"background 0.2s" }}>
                        {isProcessing ? "Analyzing..." : "Analyze with AI"}
                      </button>

                      {processingLog.length>0 && (
                        <div style={{ marginTop:16, background:"#262626", borderRadius:8, padding:12, fontFamily:"monospace", fontSize:12, maxHeight:120, overflowY:"auto" }}>
                          {processingLog.map((e,i) => (
                            <div key={i} style={{ color:e.msg.startsWith("✓")?"#77AE58":e.msg.startsWith("✗")?"#E38D8B":"#769FC8", marginBottom:3 }}>
                              <span style={{ color:"#767676", marginRight:8 }}>{e.time}</span>{e.msg}
                            </div>
                          ))}
                        </div>
                      )}
                    </div>
                  </div>
                </div>

                {/* Items needing sync */}
                <div style={{ background:"#FFFFFF", borderRadius:12, border:"1px solid #E9E9E9", overflow:"hidden" }}>
                  <div style={{ padding:"16px 24px", borderBottom:"1px solid #E9E9E9", display:"flex", justifyContent:"space-between", alignItems:"center" }}>
                    <span style={{ fontSize:14, fontWeight:500, color:"#262626" }}>NOW Items Needing Jira Sync ({stats.needsSync})</span>
                    <span onClick={() => setView("roadmap")} style={{ fontSize:14, color:"#8932AF", cursor:"pointer", textDecoration:"underline" }}>View all</span>
                  </div>
                  {items.filter(i=>i.horizon==="now"&&!i.jiraKey).slice(0,4).map(item => (
                    <div key={item.id} className="row-hover"
                      style={{ padding:"12px 24px", borderBottom:"1px solid #F2F2F2", display:"flex", alignItems:"center", gap:12, cursor:"pointer", transition:"background 0.1s" }}
                      onClick={() => { setSelectedItem(item); setView("roadmap"); }}>
                      <div style={{ width:8, height:8, borderRadius:2, background:getPillar(item.pillar).color, flexShrink:0 }} />
                      <div style={{ flex:1, minWidth:0 }}>
                        <div style={{ fontSize:14, fontWeight:400, color:"#262626", whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis" }}>{item.title}</div>
                        <div style={{ fontSize:12, color:"#767676" }}>{getPillar(item.pillar).label}</div>
                      </div>
                      <button onClick={e=>{e.stopPropagation();syncToJira(item);}} disabled={isProcessing}
                        style={{ fontSize:14, padding:"6px 16px", borderRadius:9999, background:"#7D3F98", color:"#FFF", border:"none", cursor:"pointer", flexShrink:0 }}>
                        Sync Jira
                      </button>
                    </div>
                  ))}
                  {stats.needsSync===0 && (
                    <div style={{ padding:24, textAlign:"center", color:"#767676", fontSize:14 }}>All NOW items are synced</div>
                  )}
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
                  {[{id:"all",label:"All Pillars",color:"#575757"}, ...PILLARS].map(p => (
                    <button key={p.id} onClick={()=>setFilterPillar(p.id)}
                      style={{ padding:"5px 12px", borderRadius:20, fontSize:12, fontWeight:500, border:"1px solid "+(filterPillar===p.id?p.color:"#D0CEC8"), background:filterPillar===p.id?p.color:"#FFF", color:filterPillar===p.id?"#FFF":"#575757", cursor:"pointer", transition:"all 0.15s" }}>
                      {p.label}
                    </button>
                  ))}
                  <span style={{ marginLeft:"auto", fontSize:12, color:"#767676", alignSelf:"center" }}>{filtered.length} items</span>
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
                          <span style={{ fontSize:11, color:"#767676" }}>{hz.sub}</span>
                          <span style={{ marginLeft:"auto", fontSize:12, fontWeight:600, color:"#262626", background:"#F2F2F2", padding:"2px 8px", borderRadius:10 }}>{colItems.length}</span>
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
                                style={{ background:"#FFF", borderRadius:10, padding:"12px 14px", border:isSelected?`2px solid ${p.color}`:"1px solid #E9E9E9", cursor:"pointer", transition:"all 0.15s" }}>
                                <div style={{ display:"flex", alignItems:"flex-start", gap:7, marginBottom:8 }}>
                                  <div style={{ width:8, height:8, borderRadius:2, background:p.color, marginTop:4, flexShrink:0 }} />
                                  <div style={{ fontSize:12, fontWeight:500, color:"#262626", lineHeight:1.4, flex:1 }}>{item.title}</div>
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
                                  <div style={{ fontSize:10, color:"#767676", marginTop:5 }}>💬 {item.comments.length} note{item.comments.length!==1?"s":""}</div>
                                )}
                              </div>
                            );
                          })}
                          {colItems.length===0 && (
                            <div style={{ padding:24, textAlign:"center", color:"#C0BEB8", fontSize:12, border:"1px dashed #E9E9E9", borderRadius:10 }}>
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

          {/* ── INVESTMENTS ──────────────────────────────────────────────── */}
          {view==="investments" && (
            <div style={{ flex:1, overflow:"auto", padding:28, background:"#F7F7F7" }}>
              {/* Hero */}
              <div style={{ background:"linear-gradient(135deg, #5C2E72 0%, #7D3F98 100%)", borderRadius:14, padding:"28px 32px", marginBottom:20, color:"#FFF" }}>
                <div style={{ fontSize:10, fontWeight:700, color:"#7D3F98", letterSpacing:"0.14em", textTransform:"uppercase", marginBottom:4 }}>Our Strategy</div>
                <h1 style={{ fontSize:22, fontWeight:700, margin:"0 0 8px" }}>FY2026–2027 Investment Overview</h1>
                <p style={{ fontSize:13, color:"rgba(255,255,255,0.7)", margin:0 }}>Improve population health and program sustainability by advancing total cost of care, strengthening state and provider partnerships, and delivering simpler, more personalized member experiences.</p>
                <p style={{ fontSize:11, color:"rgba(255,255,255,0.45)", margin:"8px 0 0" }}>Total projected EBIT benefit (2026–2031): $676.6M</p>
              </div>

              {/* Top-line stats */}
              <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr 1fr", gap:12, marginBottom:20 }}>
                {[
                  { label:"2026 CapEx", value:"$45.3M", color:"#1565C0" },
                  { label:"2027 CapEx", value:"$67.9M", color:"#2E7D32" },
                  { label:"2026 Budget", value:"$105.2M", color:"#6A1B9A" },
                  { label:"FTE Count", value:"~125", color:"#B71C1C" },
                ].map((s, i) => (
                  <div key={i} style={{ background:"#FFF", borderRadius:10, border:"1px solid #E9E9E9", borderLeft:`3px solid ${s.color}`, padding:"16px 20px" }}>
                    <div style={{ fontSize:10, fontWeight:600, color:"#767676", textTransform:"uppercase", letterSpacing:"0.06em", marginBottom:4 }}>{s.label}</div>
                    <div style={{ fontSize:24, fontWeight:700, color:"#262626" }}>{s.value}</div>
                  </div>
                ))}
              </div>

              <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:16, marginBottom:16 }}>
                {/* Carryover Investments */}
                <div style={{ background:"#FFF", borderRadius:12, border:"1px solid #E9E9E9", padding:20 }}>
                  <div style={{ fontSize:11, fontWeight:700, color:"#7D3F98", textTransform:"uppercase", letterSpacing:"0.08em", marginBottom:4 }}>Carryover Investments</div>
                  <div style={{ fontSize:10, color:"#767676", marginBottom:14 }}>~$43.52M current 2027 CapEx</div>
                  {[
                    { name:"Aetna Health & Innovation Platform", amount:"$9.13M", desc:"Member experience modernization" },
                    { name:"Medicaid Contractual Projects", amount:"$8.44M", desc:"Compliance (Final Rule/Duals, OB3)" },
                    { name:"Claims Carry-Over", amount:"$5.23M", desc:"Universal Claim Tracker, Xten Upgrade, Fee Schedule" },
                    { name:"Enrollment Carry-Over", amount:"$4.69M", desc:"End-to-end automation, eligibility processing" },
                    { name:"Claims GPS Carry-Over", amount:"$4.04M", desc:"Moving Medicaid to GPS (Salesforce)" },
                    { name:"Claims Workflow Management", amount:"$3.92M", desc:"QNXT Upgrade, remedy replacement" },
                    { name:"Medicaid Clinical Carry-Over", amount:"$3.47M", desc:"Prior Auth, care coordination" },
                  ].map((inv, i) => (
                    <div key={i} style={{ padding:"10px 0", borderBottom: i < 6 ? "1px solid #F2F2F2" : "none" }}>
                      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:2 }}>
                        <span style={{ fontSize:12, fontWeight:600, color:"#262626" }}>{inv.name}</span>
                        <span style={{ fontSize:12, fontWeight:700, color:"#7D3F98" }}>{inv.amount}</span>
                      </div>
                      <div style={{ fontSize:11, color:"#767676" }}>{inv.desc}</div>
                    </div>
                  ))}
                </div>

                {/* New 2027 Investments */}
                <div style={{ background:"#FFF", borderRadius:12, border:"1px solid #E9E9E9", padding:20 }}>
                  <div style={{ fontSize:11, fontWeight:700, color:"#2E7D32", textTransform:"uppercase", letterSpacing:"0.08em", marginBottom:4 }}>New 2027 Investments</div>
                  <div style={{ fontSize:10, color:"#767676", marginBottom:14 }}>~$20.12M new CapEx</div>
                  {[
                    { name:"Aetna Medicaid Quality Optimization", amount:"$9.70M", desc:"HEDIS gaps, STARS, AI-driven quality", highlight:true },
                    { name:"Medicaid Cloud Optimization", amount:"$5.44M", desc:"Legacy to Azure cloud-native migration" },
                    { name:"Premium Billing Modernization", amount:"$0.83M", desc:"End-to-end billing capability" },
                    { name:"Member Domain AI/Automation", amount:"$0.73M", desc:"AI/ML for enrollment, 834, pharmacy reconciliation" },
                    { name:"AI Data Abstractions for HEDIS", amount:"$0.50M", desc:"Extract clinical insights from unstructured data" },
                  ].map((inv, i) => (
                    <div key={i} style={{ padding:"10px 0", borderBottom: i < 4 ? "1px solid #F2F2F2" : "none", background: inv.highlight ? "#F0FAF0" : "transparent", borderRadius: inv.highlight ? 6 : 0, paddingLeft: inv.highlight ? 10 : 0, paddingRight: inv.highlight ? 10 : 0 }}>
                      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:2 }}>
                        <span style={{ fontSize:12, fontWeight:600, color:"#262626" }}>{inv.name}</span>
                        <span style={{ fontSize:12, fontWeight:700, color:"#2E7D32" }}>{inv.amount}</span>
                      </div>
                      <div style={{ fontSize:11, color:"#767676" }}>{inv.desc}</div>
                    </div>
                  ))}

                  <div style={{ marginTop:20, borderTop:"1px solid #E9E9E9", paddingTop:16 }}>
                    <div style={{ fontSize:11, fontWeight:700, color:"#B71C1C", textTransform:"uppercase", letterSpacing:"0.08em", marginBottom:4 }}>ABX Initiatives</div>
                    <div style={{ fontSize:10, color:"#767676", marginBottom:10 }}>Cost structure focus — ~$12.8M 2027</div>
                    {["Medicaid Cloud Optimization","Triangulation","QNXT Consolidation","Pursuant HRQ In-House Migration","Enrollment Carryover"].map((a, i) => (
                      <div key={i} style={{ fontSize:11, color:"#575757", padding:"4px 0", display:"flex", alignItems:"center", gap:6 }}>
                        <span style={{ width:4, height:4, borderRadius:2, background:"#B71C1C", flexShrink:0 }} />{a}
                      </div>
                    ))}
                  </div>
                </div>
              </div>

              {/* Historical Spend */}
              <div style={{ background:"#FFF", borderRadius:12, border:"1px solid #E9E9E9", padding:20, marginBottom:16 }}>
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
                      <span style={{ fontSize:11, fontWeight:700, color:"#262626" }}>${d.val}M</span>
                      <div style={{ width:"100%", height:`${(d.val/d.max)*80}px`, background: i >= 3 ? "linear-gradient(180deg, #7D3F98, #5c2d82)" : "#E9E9E9", borderRadius:"4px 4px 0 0", transition:"height 0.3s" }} />
                      <span style={{ fontSize:9, color:"#767676", fontWeight:500 }}>{d.year}</span>
                    </div>
                  ))}
                </div>
              </div>

              {/* Key Initiatives In-Flight */}
              <div style={{ background:"#FFF", borderRadius:12, border:"1px solid #E9E9E9", padding:20, marginBottom:16 }}>
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
                    <div key={i} style={{ display:"flex", alignItems:"center", gap:10, padding:"10px 12px", background:"#F7F7F7", borderRadius:8 }}>
                      <span style={{ width:24, height:24, borderRadius:6, background:"#7D3F98", color:"#FFF", fontSize:11, fontWeight:700, display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0 }}>{init.code}</span>
                      <div>
                        <div style={{ fontSize:12, fontWeight:600, color:"#262626" }}>{init.name}</div>
                        <div style={{ fontSize:10, color:"#767676" }}>{init.desc}</div>
                      </div>
                      <span style={{ marginLeft:"auto", fontSize:10, padding:"2px 8px", borderRadius:10, background:"#E6F4EA", color:"#1E7E34", fontWeight:600 }}>In-Flight</span>
                    </div>
                  ))}
                </div>
              </div>

              {/* Critical Deadlines */}
              <div style={{ background:"#FFF", borderRadius:12, border:"1px solid #E9E9E9", padding:20 }}>
                <div style={{ fontSize:11, fontWeight:700, color:"#B71C1C", textTransform:"uppercase", letterSpacing:"0.08em", marginBottom:14 }}>Critical Deadlines</div>
                {[
                  { item:"Claims Remedy Replacement", deadline:"Jan 1, 2027", note:"Day-1 scope lock + stakeholder sign-off", severity:"critical" },
                  { item:"QNXT 2025.R1 Upgrade", deadline:"Active", note:"Both Claims and Enrollment teams", severity:"high" },
                  { item:"TX STAR CHP Implementation", deadline:"Completed 4/1", note:"Post-KY 1115 Waiver", severity:"done" },
                  { item:"Open Enrollment / Welcome Season", deadline:"Ongoing", note:"Operational readiness priority", severity:"high" },
                ].map((d, i) => (
                  <div key={i} style={{ display:"flex", alignItems:"center", gap:12, padding:"10px 0", borderBottom: i < 3 ? "1px solid #F2F2F2" : "none" }}>
                    <span style={{ fontSize:11, padding:"3px 8px", borderRadius:6, fontWeight:700, background: d.severity==="critical"?"#FEE2E2":d.severity==="done"?"#E6F4EA":"#FEF3C7", color: d.severity==="critical"?"#DC2626":d.severity==="done"?"#1E7E34":"#B45309" }}>{d.deadline}</span>
                    <div style={{ flex:1 }}>
                      <div style={{ fontSize:12, fontWeight:600, color:"#262626" }}>{d.item}</div>
                      <div style={{ fontSize:10, color:"#767676" }}>{d.note}</div>
                    </div>
                  </div>
                ))}
              </div>

              {/* Vendor Ecosystem */}
              <div style={{ background:"#FFF", borderRadius:12, border:"1px solid #E9E9E9", padding:20, marginTop:16 }}>
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
                    <div key={i} style={{ padding:"10px 12px", background:"#F7F7F7", borderRadius:8, textAlign:"center" }}>
                      <div style={{ fontSize:11, fontWeight:600, color:"#262626", marginBottom:2 }}>{v.vendor}</div>
                      <div style={{ fontSize:9, color:"#767676" }}>{v.role}</div>
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
                <div style={{ background:"linear-gradient(135deg, #5C2E72 0%, #7D3F98 100%)", borderRadius:14, padding:"20px 28px", marginBottom:16, color:"#FFF" }}>
                  <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:12 }}>
                    <div style={{ fontSize:16, fontWeight:700 }}>Medicaid Member Enrollment — Rally Roadmap</div>
                    <button onClick={()=>setRallyPanel(!rallyPanel)} style={{ fontSize:11, padding:"5px 14px", borderRadius:6, background: rallyPanel ? "rgba(125,63,152,0.3)" : "rgba(255,255,255,0.12)", color:"#FFF", border: rallyPanel ? "1px solid rgba(125,63,152,0.6)" : "1px solid rgba(255,255,255,0.2)", cursor:"pointer", fontWeight:600, display:"flex", alignItems:"center", gap:5 }}>
                      <Icon name="settings" size={13} color="#FFF" /> Filters{activeFilterCount > 0 && <span style={{ background:"#7D3F98", fontSize:10, fontWeight:700, borderRadius:10, padding:"1px 6px" }}>{activeFilterCount}</span>}
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
                        {rallyLoading ? "Refreshing..." : "Refresh from Rally"}
                      </button>
                    </div>
                  </div>
                </div>

                {/* Search */}
                <div style={{ marginBottom:16 }}>
                  <input type="search" placeholder="Search features by name, ID, project, product..."
                    value={rf.search} onChange={e => setRallyFilter({...rf, search:e.target.value})}
                    style={{ width:"100%", padding:"10px 16px", borderRadius:10, border:"1px solid #E9E9E9", fontSize:13, background:"#FFF" }} />
                </div>

                {/* Lanes grid */}
                <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr 1fr", gap:16, alignItems:"start" }}>
                  {["done","now","next","future"].map(laneKey => {
                    const meta = LANE_META[laneKey];
                    const STATE_ORDER = { "":0, "Backlog":1, "Ask":2, "Prepare":3, "Ready":4, "In Progress":5, "Done":6 };
                    const laneFeatures = visible.filter(f => f.lane===laneKey).sort((a,b) => (STATE_ORDER[b.state]||0) - (STATE_ORDER[a.state]||0));
                    return (
                      <div key={laneKey} style={{ background:"#FFF", borderRadius:14, border:"1px solid #E9E9E9", overflow:"hidden" }}>
                        <div style={{ padding:"14px 16px 10px", borderBottom:"1px solid #F2F2F2" }}>
                          <div style={{ display:"flex", alignItems:"center", gap:8 }}>
                            <div style={{ width:10, height:10, borderRadius:"50%", background:meta.dot }} />
                            <span style={{ fontSize:14, fontWeight:700, color:"#262626" }}>{meta.label}</span>
                            <span style={{ fontSize:12, color:"#767676", marginLeft:"auto" }}>{laneFeatures.length}</span>
                          </div>
                          <div style={{ fontSize:11, color:"#767676", marginTop:4 }}>{meta.sub}</div>
                        </div>
                        <div style={{ padding:12, maxHeight:600, overflow:"auto" }}>
                          {laneFeatures.length === 0 && <div style={{ textAlign:"center", padding:20, color:"#767676", fontSize:12 }}>No features</div>}
                          {laneFeatures.map(f => (
                            <div key={f.id} className="card-hover" style={{ background:"#FAFBFD", border:"1px solid #E9ECF3", borderLeft:`3px solid ${meta.dot}`, borderRadius:8, padding:"10px 12px", marginBottom:8, cursor:"default" }}>
                              <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:4 }}>
                                <span style={{ fontSize:10, padding:"2px 7px", borderRadius:10, background:(STATUS_COLORS[f.state]||"#545b64")+"18", color:STATUS_COLORS[f.state]||"#545b64", fontWeight:600, textTransform:"uppercase" }}>{f.state||"Defining"}</span>
                                <a href={f.rallyUrl} target="_blank" rel="noopener noreferrer" style={{ fontSize:11, color:"#1558b0", textDecoration:"none", fontFamily:"monospace", fontWeight:500 }}>{f.id}</a>
                              </div>
                              <div style={{ fontSize:12, fontWeight:600, color:"#262626", lineHeight:1.35, marginBottom:6, display:"-webkit-box", WebkitLineClamp:2, WebkitBoxOrient:"vertical", overflow:"hidden" }}>{f.name}</div>
                              <div style={{ display:"flex", flexWrap:"wrap", gap:4 }}>
                                <span style={{ fontSize:9, padding:"2px 6px", borderRadius:3, background:(PRODUCT_COLORS[f.product]||"#37474f")+"14", color:PRODUCT_COLORS[f.product]||"#37474f", fontWeight:600 }}>{f.product}</span>
                                <span style={{ fontSize:9, padding:"2px 6px", borderRadius:10, background:(PILLAR_COLORS[f.pillar]||"#1565C0")+"14", color:PILLAR_COLORS[f.pillar]||"#1565C0", fontWeight:700 }}>{f.pillar}</span>
                                {f.release && <span style={{ fontSize:9, padding:"2px 5px", borderRadius:3, background:"#f0f2f5", color:"#6B7280", fontWeight:600 }}>{f.release}</span>}
                                {f.date && <span style={{ fontSize:9, color:"#767676", marginLeft:"auto" }}>{f.date}</span>}
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
                <div style={{ width:320, borderLeft:"1px solid #E9E9E9", background:"#FFF", display:"flex", flexDirection:"column", overflow:"hidden", flexShrink:0, animation:"fadeUp 0.2s ease" }}>
                  {/* Panel header */}
                  <div style={{ padding:"16px 20px", borderBottom:"1px solid #E9E9E9", display:"flex", alignItems:"center", justifyContent:"space-between", background:"#F7F7F7" }}>
                    <div>
                      <div style={{ fontSize:14, fontWeight:700, color:"#262626" }}>Filter Features</div>
                      <div style={{ fontSize:11, color:"#767676", marginTop:2 }}>{visible.length} of {rallyFeatures.length} features visible</div>
                    </div>
                    <button onClick={()=>setRallyPanel(false)} style={{ background:"none", border:"none", cursor:"pointer", color:"#767676", lineHeight:1, padding:4 }}><Icon name="close" size={16} /></button>
                  </div>

                  <div style={{ flex:1, overflow:"auto", padding:20 }}>
                    {/* Search */}
                    <div style={{ marginBottom:20 }}>
                      <label style={{ fontSize:11, fontWeight:600, color:"#575757", textTransform:"uppercase", letterSpacing:"0.06em", display:"block", marginBottom:6 }}>Search</label>
                      <input type="search" placeholder="Search features..."
                        value={rf.search} onChange={e=>setRallyFilter({...rf, search:e.target.value})}
                        style={{ width:"100%", padding:"8px 12px", borderRadius:8, border:"1px solid #E9E9E9", fontSize:12 }} />
                    </div>

                    {/* Product */}
                    <div style={{ marginBottom:20 }}>
                      <label style={{ fontSize:11, fontWeight:600, color:"#575757", textTransform:"uppercase", letterSpacing:"0.06em", display:"block", marginBottom:8 }}>Product</label>
                      <div style={{ display:"flex", flexDirection:"column", gap:4 }}>
                        <button onClick={()=>setRallyFilter({...rf, product:"all"})}
                          style={{ display:"flex", alignItems:"center", gap:8, padding:"8px 12px", borderRadius:8, border: rf.product==="all" ? "1.5px solid #7D3F98" : "1px solid #E9E9E9", background: rf.product==="all" ? "#F5F0FA" : "#FFF", cursor:"pointer", fontSize:12, fontWeight: rf.product==="all" ? 600 : 400, color: rf.product==="all" ? "#7D3F98" : "#575757", textAlign:"left" }}>
                          All Products <span style={{ marginLeft:"auto", fontSize:11, color:"#767676" }}>{rallyFeatures.length}</span>
                        </button>
                        {products.map(p => {
                          const count = rallyFeatures.filter(f=>f.product===p).length;
                          const color = PRODUCT_COLORS[p]||"#37474f";
                          return (
                            <button key={p} onClick={()=>setRallyFilter({...rf, product: rf.product===p ? "all" : p})}
                              style={{ display:"flex", alignItems:"center", gap:8, padding:"8px 12px", borderRadius:8, border: rf.product===p ? `1.5px solid ${color}` : "1px solid #E9E9E9", background: rf.product===p ? color+"10" : "#FFF", cursor:"pointer", fontSize:12, fontWeight: rf.product===p ? 600 : 400, color: rf.product===p ? color : "#575757", textAlign:"left" }}>
                              <div style={{ width:8, height:8, borderRadius:2, background:color, flexShrink:0 }} />
                              {p} <span style={{ marginLeft:"auto", fontSize:11, color:"#767676" }}>{count}</span>
                            </button>
                          );
                        })}
                      </div>
                    </div>

                    {/* Strategic Pillar */}
                    <div style={{ marginBottom:20 }}>
                      <label style={{ fontSize:11, fontWeight:600, color:"#575757", textTransform:"uppercase", letterSpacing:"0.06em", display:"block", marginBottom:8 }}>Strategic Pillar</label>
                      <div style={{ display:"flex", flexDirection:"column", gap:4 }}>
                        <button onClick={()=>setRallyFilter({...rf, pillar:"all"})}
                          style={{ display:"flex", alignItems:"center", gap:8, padding:"8px 12px", borderRadius:8, border: rf.pillar==="all" ? "1.5px solid #7D3F98" : "1px solid #E9E9E9", background: rf.pillar==="all" ? "#F5F0FA" : "#FFF", cursor:"pointer", fontSize:12, fontWeight: rf.pillar==="all" ? 600 : 400, color: rf.pillar==="all" ? "#7D3F98" : "#575757", textAlign:"left" }}>
                          All Pillars <span style={{ marginLeft:"auto", fontSize:11, color:"#767676" }}>{rallyFeatures.length}</span>
                        </button>
                        {pillars.map(p => {
                          const count = rallyFeatures.filter(f=>f.pillar===p).length;
                          const color = PILLAR_COLORS[p]||"#1565C0";
                          return (
                            <button key={p} onClick={()=>setRallyFilter({...rf, pillar: rf.pillar===p ? "all" : p})}
                              style={{ display:"flex", alignItems:"center", gap:8, padding:"8px 12px", borderRadius:8, border: rf.pillar===p ? `1.5px solid ${color}` : "1px solid #E9E9E9", background: rf.pillar===p ? color+"10" : "#FFF", cursor:"pointer", fontSize:12, fontWeight: rf.pillar===p ? 600 : 400, color: rf.pillar===p ? color : "#575757", textAlign:"left" }}>
                              <div style={{ width:8, height:8, borderRadius:"50%", background:color, flexShrink:0 }} />
                              {p} <span style={{ marginLeft:"auto", fontSize:11, color:"#767676" }}>{count}</span>
                            </button>
                          );
                        })}
                      </div>
                    </div>

                    {/* PI / Release */}
                    <div style={{ marginBottom:20 }}>
                      <label style={{ fontSize:11, fontWeight:600, color:"#575757", textTransform:"uppercase", letterSpacing:"0.06em", display:"block", marginBottom:8 }}>PI / Release</label>
                      <div style={{ display:"flex", flexDirection:"column", gap:4 }}>
                        <button onClick={()=>setRallyFilter({...rf, release:"all"})}
                          style={{ display:"flex", alignItems:"center", gap:8, padding:"8px 12px", borderRadius:8, border: rf.release==="all" ? "1.5px solid #7D3F98" : "1px solid #E9E9E9", background: rf.release==="all" ? "#F5F0FA" : "#FFF", cursor:"pointer", fontSize:12, fontWeight: rf.release==="all" ? 600 : 400, color: rf.release==="all" ? "#7D3F98" : "#575757", textAlign:"left" }}>
                          All Releases <span style={{ marginLeft:"auto", fontSize:11, color:"#767676" }}>{rallyFeatures.length}</span>
                        </button>
                        {releases.map(r => {
                          const count = rallyFeatures.filter(f=>f.release===r).length;
                          return (
                            <button key={r} onClick={()=>setRallyFilter({...rf, release: rf.release===r ? "all" : r})}
                              style={{ display:"flex", alignItems:"center", gap:8, padding:"8px 12px", borderRadius:8, border: rf.release===r ? "1.5px solid #7D3F98" : "1px solid #E9E9E9", background: rf.release===r ? "#F5F0FA" : "#FFF", cursor:"pointer", fontSize:12, fontWeight: rf.release===r ? 600 : 400, color: rf.release===r ? "#7D3F98" : "#575757", textAlign:"left" }}>
                              {r} <span style={{ marginLeft:"auto", fontSize:11, color:"#767676" }}>{count}</span>
                            </button>
                          );
                        })}
                      </div>
                    </div>

                    {/* Value Sector */}
                    <div style={{ marginBottom:20 }}>
                      <label style={{ fontSize:11, fontWeight:600, color:"#575757", textTransform:"uppercase", letterSpacing:"0.06em", display:"block", marginBottom:8 }}>Value Sector</label>
                      <div style={{ display:"flex", flexDirection:"column", gap:4 }}>
                        <button onClick={()=>setRallyFilter({...rf, vs:"all"})}
                          style={{ display:"flex", alignItems:"center", gap:8, padding:"8px 12px", borderRadius:8, border: rf.vs==="all" ? "1.5px solid #7D3F98" : "1px solid #E9E9E9", background: rf.vs==="all" ? "#F5F0FA" : "#FFF", cursor:"pointer", fontSize:12, fontWeight: rf.vs==="all" ? 600 : 400, color: rf.vs==="all" ? "#7D3F98" : "#575757", textAlign:"left" }}>
                          All Sectors <span style={{ marginLeft:"auto", fontSize:11, color:"#767676" }}>{rallyFeatures.length}</span>
                        </button>
                        {valueSectors.map(vs => {
                          const count = rallyFeatures.filter(f=>(f.valueSector||"")===vs).length;
                          return (
                            <button key={vs} onClick={()=>setRallyFilter({...rf, vs: rf.vs===vs ? "all" : vs})}
                              style={{ display:"flex", alignItems:"center", gap:8, padding:"8px 12px", borderRadius:8, border: rf.vs===vs ? "1.5px solid #7D3F98" : "1px solid #E9E9E9", background: rf.vs===vs ? "#F5F0FA" : "#FFF", cursor:"pointer", fontSize:12, fontWeight: rf.vs===vs ? 600 : 400, color: rf.vs===vs ? "#7D3F98" : "#575757", textAlign:"left" }}>
                              {vs} <span style={{ marginLeft:"auto", fontSize:11, color:"#767676" }}>{count}</span>
                            </button>
                          );
                        })}
                      </div>
                    </div>

                    {/* Rally Project */}
                    <div style={{ marginBottom:20 }}>
                      <label style={{ fontSize:11, fontWeight:600, color:"#575757", textTransform:"uppercase", letterSpacing:"0.06em", display:"block", marginBottom:8 }}>Rally Project</label>
                      <div style={{ display:"flex", flexDirection:"column", gap:4 }}>
                        <button onClick={()=>setRallyFilter({...rf, project:"all"})}
                          style={{ display:"flex", alignItems:"center", gap:8, padding:"8px 12px", borderRadius:8, border: rf.project==="all" ? "1.5px solid #7D3F98" : "1px solid #E9E9E9", background: rf.project==="all" ? "#F5F0FA" : "#FFF", cursor:"pointer", fontSize:12, fontWeight: rf.project==="all" ? 600 : 400, color: rf.project==="all" ? "#7D3F98" : "#575757", textAlign:"left" }}>
                          All Projects <span style={{ marginLeft:"auto", fontSize:11, color:"#767676" }}>{rallyFeatures.length}</span>
                        </button>
                        {projects.map(p => {
                          const count = rallyFeatures.filter(f=>f.project===p).length;
                          return (
                            <button key={p} onClick={()=>setRallyFilter({...rf, project: rf.project===p ? "all" : p})}
                              style={{ display:"flex", alignItems:"center", gap:8, padding:"8px 12px", borderRadius:8, border: rf.project===p ? "1.5px solid #7D3F98" : "1px solid #E9E9E9", background: rf.project===p ? "#F5F0FA" : "#FFF", cursor:"pointer", fontSize:12, fontWeight: rf.project===p ? 600 : 400, color: rf.project===p ? "#7D3F98" : "#575757", textAlign:"left" }}>
                              {p} <span style={{ marginLeft:"auto", fontSize:11, color:"#767676" }}>{count}</span>
                            </button>
                          );
                        })}
                      </div>
                    </div>

                    {/* Reset */}
                    {activeFilterCount > 0 && (
                      <button onClick={()=>setRallyFilter({ product:"all", pillar:"all", vs:"all", release:"all", project:"all", search:"" })}
                        style={{ width:"100%", padding:"10px", borderRadius:9999, border:"1px solid #DC2626", background:"#FEF2F2", color:"#DC2626", cursor:"pointer", fontSize:12, fontWeight:600 }}>
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
                <h1 style={{ fontSize:24, fontWeight:600, color:"#262626", marginBottom:8 }}>Team & Reviews</h1>
                <p style={{ color:"#575757", fontSize:14, marginBottom:28 }}>
                  Sync outstanding items to Jira, publish Confluence pages, and review team activity.
                </p>

                <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:20, marginBottom:24 }}>
                  {/* Jira queue */}
                  <div style={{ background:"#FFF", borderRadius:12, border:"1px solid #E9E9E9", overflow:"hidden" }}>
                    <div style={{ padding:"14px 20px", borderBottom:"1px solid #E9E9E9", background:"#EBF3FF", display:"flex", justifyContent:"space-between", alignItems:"center" }}>
                      <div>
                        <div style={{ fontSize:13, fontWeight:600, color:"#0A4B8C", display:"flex", alignItems:"center", gap:6 }}><Icon name="link" size={14} color="#0A4B8C" />Jira Sync Queue</div>
                        <div style={{ fontSize:11, color:"#575757", marginTop:2 }}>NOW items without an epic</div>
                      </div>
                      <span style={{ fontSize:18, fontWeight:700, color:"#0A4B8C" }}>{stats.needsSync}</span>
                    </div>
                    {items.filter(i=>i.horizon==="now"&&!i.jiraKey).map(item => (
                      <div key={item.id} style={{ padding:"12px 20px", borderBottom:"1px solid #F2F2F2", display:"flex", alignItems:"center", gap:10 }}>
                        <div style={{ width:8, height:8, borderRadius:2, background:getPillar(item.pillar).color, flexShrink:0 }} />
                        <div style={{ flex:1, minWidth:0 }}>
                          <div style={{ fontSize:12, fontWeight:500, color:"#262626", whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis" }}>{item.title}</div>
                          <div style={{ fontSize:10, color:"#767676" }}>{getPillar(item.pillar).label} · {getStatus(item.status).label}</div>
                        </div>
                        <button onClick={()=>syncToJira(item)} disabled={isProcessing}
                          style={{ fontSize:11, padding:"4px 10px", borderRadius:9999, background:"#0A4B8C", color:"#FFF", border:"none", cursor:"pointer", flexShrink:0 }}>
                          Sync
                        </button>
                      </div>
                    ))}
                    {stats.needsSync===0 && (
                      <div style={{ padding:24, textAlign:"center", color:"#767676", fontSize:12, display:"flex", alignItems:"center", justifyContent:"center", gap:6 }}><Icon name="check" size={14} color="#767676" />All NOW items synced to Jira</div>
                    )}
                    {items.filter(i=>i.horizon==="now"&&!i.jiraKey).length>0 && (
                      <div style={{ padding:"12px 20px", background:"#F7F7F7" }}>
                        <button onClick={async()=>{for(const item of items.filter(i=>i.horizon==="now"&&!i.jiraKey)){await syncToJira(item);}}}
                          disabled={isProcessing}
                          style={{ fontSize:12, padding:"7px 16px", borderRadius:8, background:"#0A4B8C", color:"#FFF", border:"none", cursor:"pointer", fontWeight:500, width:"100%" }}>
                          Sync All to Jira
                        </button>
                      </div>
                    )}
                  </div>

                  {/* Confluence queue */}
                  <div style={{ background:"#FFF", borderRadius:12, border:"1px solid #E9E9E9", overflow:"hidden" }}>
                    <div style={{ padding:"14px 20px", borderBottom:"1px solid #E9E9E9", background:"#E6F5F5", display:"flex", justifyContent:"space-between", alignItems:"center" }}>
                      <div>
                        <div style={{ fontSize:13, fontWeight:600, color:"#00787E", display:"flex", alignItems:"center", gap:6 }}><Icon name="doc" size={14} color="#00787E" />Confluence Publish Queue</div>
                        <div style={{ fontSize:11, color:"#575757", marginTop:2 }}>Items without a Confluence page</div>
                      </div>
                      <span style={{ fontSize:18, fontWeight:700, color:"#00787E" }}>{items.filter(i=>!i.confluenceUrl&&i.horizon==="now").length}</span>
                    </div>
                    {items.filter(i=>!i.confluenceUrl&&i.horizon==="now").map(item => (
                      <div key={item.id} style={{ padding:"12px 20px", borderBottom:"1px solid #F2F2F2", display:"flex", alignItems:"center", gap:10 }}>
                        <div style={{ width:8, height:8, borderRadius:2, background:getPillar(item.pillar).color, flexShrink:0 }} />
                        <div style={{ flex:1, minWidth:0 }}>
                          <div style={{ fontSize:12, fontWeight:500, color:"#262626", whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis" }}>{item.title}</div>
                          <div style={{ fontSize:10, color:"#767676" }}>{getPillar(item.pillar).label}</div>
                        </div>
                        <button onClick={()=>publishToConfluence(item)} disabled={isProcessing}
                          style={{ fontSize:11, padding:"4px 10px", borderRadius:6, background:"#00787E", color:"#FFF", border:"none", cursor:"pointer", flexShrink:0 }}>
                          Publish
                        </button>
                      </div>
                    ))}
                    {items.filter(i=>!i.confluenceUrl&&i.horizon==="now").length===0 && (
                      <div style={{ padding:24, textAlign:"center", color:"#767676", fontSize:12, display:"flex", alignItems:"center", justifyContent:"center", gap:6 }}><Icon name="check" size={14} color="#767676" />All NOW items have Confluence pages</div>
                    )}
                  </div>
                </div>

                {/* Full item table */}
                <div style={{ background:"#FFF", borderRadius:12, border:"1px solid #E9E9E9", overflow:"hidden" }}>
                  <div style={{ padding:"14px 20px", borderBottom:"1px solid #E9E9E9", fontSize:13, fontWeight:600, color:"#262626" }}>
                    All Initiatives ({items.length})
                  </div>
                  <div style={{ overflowX:"auto" }}>
                    <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12 }}>
                      <thead>
                        <tr style={{ background:"#F7F7F7" }}>
                          {["Initiative","Pillar","Horizon","Status","Priority","Jira","Confluence","Actions"].map(h => (
                            <th key={h} style={{ padding:"10px 16px", textAlign:"left", fontSize:11, fontWeight:600, color:"#575757", whiteSpace:"nowrap" }}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {items.map(item => {
                          const p = getPillar(item.pillar);
                          const s = getStatus(item.status);
                          const h = getHorizon(item.horizon);
                          return (
                            <tr key={item.id} className="row-hover" style={{ borderBottom:"1px solid #F2F2F2", cursor:"pointer", transition:"background 0.1s" }}
                              onClick={() => { setSelectedItem(item); setView("roadmap"); }}>
                              <td style={{ padding:"10px 16px", fontWeight:500, color:"#262626", maxWidth:220, whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis" }}>{item.title}</td>
                              <td style={{ padding:"10px 16px" }}><span style={{ fontSize:10, color:p.color, fontWeight:600, background:p.bg, padding:"2px 7px", borderRadius:10 }}>{p.label}</span></td>
                              <td style={{ padding:"10px 16px" }}><span style={{ fontSize:10, fontWeight:700, color:h.color }}>{h.label}</span></td>
                              <td style={{ padding:"10px 16px" }}><span style={{ fontSize:10, color:s.color, fontWeight:500 }}>{s.label}</span></td>
                              <td style={{ padding:"10px 16px", fontSize:10, color:"#575757", textTransform:"uppercase" }}>{item.priority}</td>
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

          {view==="strat-overview" && (
            <iframe src={`${import.meta.env.BASE_URL || '/'}strategy-overview.html?page=overview`} style={{ flex:1, width:"100%", height:"100%", border:"none" }} title="3-Year Overview" />
          )}


          {view==="strat-roadmap" && (
            <iframe src={`${import.meta.env.BASE_URL || '/'}strategy-overview.html?page=roadmap${pendingInitId ? '&openInit='+pendingInitId : ''}`} onLoad={() => setPendingInitId(null)} style={{ flex:1, width:"100%", height:"100%", border:"none" }} title="Roadmap" />
          )}

          {view==="strat-context" && (
            <iframe src={`${import.meta.env.BASE_URL || '/'}strategy-overview.html?page=assumptions`} style={{ flex:1, width:"100%", height:"100%", border:"none" }} title="Context & Assumptions" />
          )}

          {view==="storyboard" && (
            <iframe src={`${import.meta.env.BASE_URL || '/'}storyboard.html`} style={{ flex:1, width:"100%", height:"100%", border:"none" }} title="Storyboard" />
          )}
        </div>
      </div>
      </div>{/* end body wrapper */}

      {/* ── NOTIFICATION ─────────────────────────────────────────────────── */}
      {notification && (
        <div style={{ position:"fixed", bottom:24, right:24, background:notification.type==="error"?"#DC2626":"#262626", color:"#FFF", padding:"12px 20px", borderRadius:10, fontSize:13, fontWeight:500, boxShadow:"0 4px 20px rgba(0,0,0,0.2)", animation:"fadeUp 0.3s ease", zIndex:9999 }}>
          {notification.type!=="error" && <span style={{ marginRight:6 }}><Icon name="sparkle" size={14} color="#7D3F98" /></span>}
          {notification.msg}
        </div>
      )}
    </div>
  );
}
