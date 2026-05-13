import * as XLSX from "xlsx";
import { saveAs } from "file-saver";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import emailjs from '@emailjs/browser';
import { motion, AnimatePresence } from 'framer-motion';
import { initializeApp } from 'firebase/app';
import { getFirestore, collection, doc, setDoc, getDocs, onSnapshot, query, where, orderBy } from 'firebase/firestore';
const EMAIL_SERVICE_ID = import.meta.env.VITE_EMAILJS_SERVICE_ID || 'service_ctyqqbc';
const EMAIL_CREATE_TEMPLATE_ID = "template_a30g4md";
const EMAIL_PUBLIC_KEY = import.meta.env.VITE_EMAILJS_PUBLIC_KEY || 'N9OlDxPyO0uf_IlxJ';
const EMAILJS_SERVICE_ID = EMAIL_SERVICE_ID;
const EMAILJS_TEMPLATE_ID = EMAIL_CREATE_TEMPLATE_ID;
const EMAILJS_PUBLIC_KEY = EMAIL_PUBLIC_KEY;
const FIREBASE_PROJECT_ID = import.meta.env.VITE_FIREBASE_PROJECT_ID;
const FIREBASE_API_KEY = import.meta.env.VITE_FIREBASE_API_KEY;
const FIRESTORE_DATABASE_ID = import.meta.env.VITE_FIRESTORE_DATABASE_ID || "(default)";

console.log("Firebase Project:", FIREBASE_PROJECT_ID);
console.log("Firestore DB:", FIRESTORE_DATABASE_ID);

const ONLINE_TICKETS_ENABLED = Boolean(FIREBASE_PROJECT_ID && FIREBASE_API_KEY);
const FIRESTORE_DATABASE_PATH = FIRESTORE_DATABASE_ID || "(default)";
const FIRESTORE_BASE_URL = ONLINE_TICKETS_ENABLED
  ? `https://firestore.googleapis.com/v1/projects/${FIREBASE_PROJECT_ID}/databases/${FIRESTORE_DATABASE_PATH}/documents`
  : "";

// Initialize Firebase
let firebaseApp = null;
let firestoreDb = null;
if (ONLINE_TICKETS_ENABLED) {
  firebaseApp = initializeApp({
    apiKey: FIREBASE_API_KEY,
    projectId: FIREBASE_PROJECT_ID,
    databaseURL: FIRESTORE_DATABASE_ID !== "(default)" ? `https://${FIREBASE_PROJECT_ID}-default-rtdb.firebaseio.com` : undefined,
  });
  firestoreDb = getFirestore(firebaseApp);
}

function toFirestoreValue(value) {
  if (value === undefined || value === null) return { nullValue: null };
  if (typeof value === "string") return { stringValue: value };
  if (typeof value === "boolean") return { booleanValue: value };
  if (typeof value === "number") return Number.isInteger(value) ? { integerValue: String(value) } : { doubleValue: value };
  if (Array.isArray(value)) return value.length ? { arrayValue: { values: value.map(toFirestoreValue) } } : { arrayValue: {} };
  if (typeof value === "object") return { mapValue: { fields: toFirestoreFields(value) } };
  return { stringValue: String(value) };
}

function toFirestoreFields(obj = {}) {
  return Object.fromEntries(Object.entries(obj).filter(([, value]) => value !== undefined).map(([key, value]) => [key, toFirestoreValue(value)]));
}

function fromFirestoreValue(value) {
  if (!value) return null;
  if ("stringValue" in value) return value.stringValue;
  if ("integerValue" in value) return Number(value.integerValue);
  if ("doubleValue" in value) return Number(value.doubleValue);
  if ("booleanValue" in value) return value.booleanValue;
  if ("nullValue" in value) return null;
  if ("arrayValue" in value) return (value.arrayValue.values || []).map(fromFirestoreValue);
  if ("mapValue" in value) return fromFirestoreFields(value.mapValue.fields || {});
  return null;
}

function fromFirestoreFields(fields = {}) {
  return Object.fromEntries(Object.entries(fields).map(([key, value]) => [key, fromFirestoreValue(value)]));
}

function normalizeTicket(ticket = {}) {
  return {
    ...ticket,
    status: ticket.status || "Assigned",
    priority: ticket.priority || "Medium",
    assigneeId: Number(ticket.assigneeId || STAFF_BASE[0]?.id || 1),
    comments: Array.isArray(ticket.comments) ? ticket.comments : [],
    timeline: Array.isArray(ticket.timeline) ? ticket.timeline : [],
    feedbackSubmitted: Boolean(ticket.feedbackSubmitted),
    closedAt: ticket.closedAt ?? null,
    closingRemarks: ticket.closingRemarks || "",
  };
}

function getFirestoreErrorMessage(body) {
  try {
    const parsed = JSON.parse(body);
    return parsed?.error?.message || body;
  } catch {
    return body;
  }
}

async function fetchTickets() {
  if (!ONLINE_TICKETS_ENABLED) {
    console.warn("Firestore is not configured. Check Vercel env vars and redeploy.");
    return [];
  }
  const res = await fetch(`${FIRESTORE_BASE_URL}/tickets?key=${encodeURIComponent(FIREBASE_API_KEY)}`);
  if (!res.ok) throw new Error(`Firestore ${res.status}: ${getFirestoreErrorMessage(await res.text())}`);
  const data = await res.json();
  return (data.documents || [])
    .map(doc => normalizeTicket({ id: doc.name?.split("/").pop(), ...fromFirestoreFields(doc.fields || {}) }))
    .sort((a, b) => (b.createdAt || 0) - (a.createdAt || 0));
}

async function saveTicket(ticket) {
  if (!ONLINE_TICKETS_ENABLED) {
    throw new Error("Firestore is not configured. Check Vercel env vars and redeploy.");
  }
  const cleanTicket = normalizeTicket(ticket);
  const url = `${FIRESTORE_BASE_URL}/tickets/${encodeURIComponent(cleanTicket.id)}?key=${encodeURIComponent(FIREBASE_API_KEY)}`;
  console.log("Firestore save URL:", url);
  const res = await fetch(url, {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ fields: toFirestoreFields(cleanTicket) }),
  });
  if (!res.ok) {
    const body = await res.text();
    console.error("Firestore save response:", res.status, body);
    throw new Error(`Firestore ${res.status}: ${getFirestoreErrorMessage(body)}`);
  }
}

async function updateTicket(ticket) {
  return saveTicket(ticket);
}

async function saveTickets(tickets) {
  if (!ONLINE_TICKETS_ENABLED) return;
  await Promise.all((tickets || []).map(ticket => saveTicket(ticket)));
}

async function deleteTicket(ticketId) {
  if (!ONLINE_TICKETS_ENABLED || !ticketId) return;
  const res = await fetch(`${FIRESTORE_BASE_URL}/tickets/${encodeURIComponent(ticketId)}?key=${encodeURIComponent(FIREBASE_API_KEY)}`, { method: "DELETE" });
  if (!res.ok && res.status !== 404) throw new Error(`Firestore ${res.status}: ${getFirestoreErrorMessage(await res.text())}`);
}

function normalizeFeedback(entry = {}) {
  return {
    ...entry,
    id: entry.id || genFeedbackId(),
    ticketId: entry.ticketId || "",
    name: entry.name || "",
    email: entry.email || "",
    dept: entry.dept || "",
    category: entry.category || "",
    rating: Number(entry.rating || 0),
    satisfaction: entry.satisfaction || "",
    recommend: entry.recommend || "Yes",
    message: entry.message || "",
    suggestions: entry.suggestions || "",
    createdAt: entry.createdAt || Date.now(),
    reviewed: Boolean(entry.reviewed),
  };
}

async function fetchFeedback() {
  if (!ONLINE_TICKETS_ENABLED) {
    console.warn("Firestore feedback storage is not configured. Check Vercel env vars and redeploy.");
    return [];
  }
  const res = await fetch(`${FIRESTORE_BASE_URL}/feedback?key=${encodeURIComponent(FIREBASE_API_KEY)}`);
  if (!res.ok) throw new Error(`Firestore feedback ${res.status}: ${getFirestoreErrorMessage(await res.text())}`);
  const data = await res.json();
  return (data.documents || [])
    .map(doc => normalizeFeedback({ id: doc.name?.split("/").pop(), ...fromFirestoreFields(doc.fields || {}) }))
    .sort((a, b) => (b.createdAt || 0) - (a.createdAt || 0));
}

async function saveFeedback(entry) {
  if (!ONLINE_TICKETS_ENABLED) {
    throw new Error("Firestore is not configured. Check Vercel env vars and redeploy.");
  }
  const cleanEntry = normalizeFeedback(entry);
  const res = await fetch(`${FIRESTORE_BASE_URL}/feedback/${encodeURIComponent(cleanEntry.id)}?key=${encodeURIComponent(FIREBASE_API_KEY)}`, {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ fields: toFirestoreFields(cleanEntry) }),
  });
  if (!res.ok) {
    const body = await res.text();
    console.error("Firestore feedback save response:", res.status, body);
    throw new Error(`Firestore feedback ${res.status}: ${getFirestoreErrorMessage(body)}`);
  }
  return cleanEntry;
}

async function updateFeedback(entry) {
  return saveFeedback(entry);
}

// ── MESSAGES (CHAT) ───────────────────────────────────────────────────────
function normalizeMessage(msg = {}) {
  return {
    ...msg,
    id: msg.id || genToken(),
    thread: msg.thread || "",
    from: Number(msg.from) || 0,
    to: Number(msg.to) || 0,
    text: msg.text || "",
    at: msg.at || Date.now(),
    read: Boolean(msg.read),
  };
}

async function fetchMessages(thread = "") {
  if (!ONLINE_TICKETS_ENABLED || !thread) return [];
  const res = await fetch(`${FIRESTORE_BASE_URL}/messages?key=${encodeURIComponent(FIREBASE_API_KEY)}`);
  if (!res.ok) throw new Error(`Firestore messages ${res.status}: ${getFirestoreErrorMessage(await res.text())}`);
  const data = await res.json();
  return (data.documents || [])
    .map(doc => normalizeMessage({ id: doc.name?.split("/").pop(), ...fromFirestoreFields(doc.fields || {}) }))
    .filter(msg => msg.thread === thread)
    .sort((a, b) => (a.at || 0) - (b.at || 0));
}

async function saveStaffProfile(staffId, profile) {
  if (!ONLINE_TICKETS_ENABLED || !firestoreDb) return;
  await setDoc(doc(firestoreDb, 'staff_profiles', staffId.toString()), profile, { merge: true });
}

async function fetchStaffProfiles() {
  if (!ONLINE_TICKETS_ENABLED || !firestoreDb) return {};
  const snapshot = await getDocs(collection(firestoreDb, 'staff_profiles'));
  const profiles = {};
  snapshot.forEach(doc => {
    profiles[doc.id] = doc.data();
  });
  return profiles;
}

function getTicketFeedbackLink(ticketId) {
  return typeof window !== "undefined"
    ? `${window.location.origin}/?feedbackTicket=${encodeURIComponent(ticketId)}`
    : `Feedback Ticket: ${ticketId}`;
}
const sendTicketEmail = async (ticket, user) => {
  try {
    await emailjs.send(
      EMAILJS_SERVICE_ID,
      EMAILJS_TEMPLATE_ID,
      {
          user_name: user?.name || 'User',
          to_email: user?.email,
          ticket_id: ticket.id,
          issue: ticket.issue || ticket.description || ticket.category || "IT Helpdesk Ticket",
          status: ticket.status || "Assigned",
      },
      EMAILJS_PUBLIC_KEY
    );

    console.log("Email sent");
  } catch (error) {
    console.log("Email error:", error);
  }
};
import { useState, useEffect, useRef, useCallback } from "react";

// ── CRYPTO HELPERS (client-side hashing simulation) ───────────────────────
async function hashPassword(pwd) {
  const buf = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(pwd + "jaipuria_salt_2024"));
  return Array.from(new Uint8Array(buf)).map(b => b.toString(16).padStart(2, "0")).join("");
}
async function verifyPassword(pwd, hash) { return (await hashPassword(pwd)) === hash; }

// ── CONSTANTS ─────────────────────────────────────────────────────────────
const STAFF_BASE = [
  { id: 1, name: "Raj Parkash Singh", role: "Senior IT Support Executive", email: "raj.parkash@jaipuria.ac.in", avatar: "RPS", color: "#6366f1", permissions: ["view_all","assign","close","export","manage_users"] },
  { id: 2, name: "Rohit Jangid", role: "IT Support Executive", email: "rohit.jangid@jaipuria.ac.in", avatar: "RJ", color: "#0ea5e9", permissions: ["view_assigned","close","comment"] },
  { id: 3, name: "Vishal Swami", role: "IT Executive", email: "vishal.swami@jaipuria.ac.in", avatar: "VS", color: "#10b981", permissions: ["view_assigned","assign","close","comment"] },
];

const CATEGORIES = [
  { id:"laptop", label:"Laptop Issue", icon:"💻", color:"#6366f1", bg:"rgba(99,102,241,0.15)" },
  { id:"desktop", label:"Desktop Issue", icon:"🖥️", color:"#0ea5e9", bg:"rgba(14,165,233,0.15)" },
  { id:"wifi", label:"WiFi Problem", icon:"📶", color:"#10b981", bg:"rgba(16,185,129,0.15)" },
  { id:"internet", label:"Internet Not Working", icon:"🌐", color:"#f59e0b", bg:"rgba(245,158,11,0.15)" },
  { id:"printer", label:"Printer Issue", icon:"🖨️", color:"#ef4444", bg:"rgba(239,68,68,0.15)" },
  { id:"software", label:"Software Installation", icon:"⚙️", color:"#8b5cf6", bg:"rgba(139,92,246,0.15)" },
  { id:"email", label:"Email Login Issue", icon:"📧", color:"#ec4899", bg:"rgba(236,72,153,0.15)" },
  { id:"projector", label:"Projector/Smart Board", icon:"📽️", color:"#14b8a6", bg:"rgba(20,184,166,0.15)" },
  { id:"biometric", label:"Biometric Issue", icon:"🔐", color:"#f97316", bg:"rgba(249,115,22,0.15)" },
  { id:"cctv", label:"CCTV Issue", icon:"📹", color:"#06b6d4", bg:"rgba(6,182,212,0.15)" },
  { id:"erp", label:"ERP/LMS Issue", icon:"📊", color:"#84cc16", bg:"rgba(132,204,22,0.15)" },
  { id:"network", label:"Network Problem", icon:"🔌", color:"#a855f7", bg:"rgba(168,85,247,0.15)" },
  { id:"hardware", label:"Hardware Damage", icon:"🔧", color:"#dc2626", bg:"rgba(220,38,38,0.15)" },
  { id:"asset", label:"New IT Asset Request", icon:"📦", color:"#0891b2", bg:"rgba(8,145,178,0.15)" },
  { id:"password", label:"Password Reset", icon:"🔑", color:"#d97706", bg:"rgba(217,119,6,0.15)" },
  { id:"other", label:"Other Complaint", icon:"💬", color:"#6b7280", bg:"rgba(107,114,128,0.15)" },
];

const DEPTS = ["Director's Office","Faculty","Student","Admin. Office","HR","Accounts","PMC","Student Affairs","MRC Office","Examination","FPM","IT","Library","Admissions & Marketing","Training","Placements & Corporate Relations","MDP","Training & Consultancy","IRC & E-Cell","Support Staff"];
const PRIORITIES = ["Low","Medium","High","Critical"];
const STATUSES = ["Open","Assigned","In Progress","Resolved","Closed"];
const SLA_HOURS = { Low:72, Medium:48, High:24, Critical:4 };
const FEEDBACK_CATEGORIES = ["Ticket Resolution","Staff Support","Internet/WiFi","Laptop/Desktop Support","Printer Support","Software Support","Overall IT Service"];
const SATISFACTION_LEVELS = [
  { id:"Excellent", icon:"😊", color:"#10b981", bg:"rgba(16,185,129,0.16)" },
  { id:"Good", icon:"🙂", color:"#3b82f6", bg:"rgba(59,130,246,0.16)" },
  { id:"Average", icon:"😐", color:"#f59e0b", bg:"rgba(245,158,11,0.16)" },
  { id:"Poor", icon:"🙁", color:"#ef4444", bg:"rgba(239,68,68,0.16)" },
];

// ── LOCAL STORAGE DB ──────────────────────────────────────────────────────
const hasStorage = () => typeof window !== "undefined" && typeof window.localStorage !== "undefined";
const safeJsonParse = (value, fallback = null) => {
  try { return value ? JSON.parse(value) : fallback; } catch { return fallback; }
};
const DB = {
  get: (k, def) => { try { if (!hasStorage()) return def; const v = localStorage.getItem("helpdesk_"+k); return safeJsonParse(v, def); } catch { return def; } },
  set: (k, v) => { try { if (hasStorage()) localStorage.setItem("helpdesk_"+k, JSON.stringify(v)); } catch {} },
};

function getActiveStaffForAssignment() {
  const statuses = DB.get("staff_statuses", {});
  return STAFF_BASE.find(staff => (statuses[staff.id] || "Online") !== "Offline") || STAFF_BASE[0];
}
// ── UTILS ─────────────────────────────────────────────────────────────────
function genId() { return "TKT-"+Date.now().toString(36).toUpperCase()+Math.random().toString(36).slice(2,5).toUpperCase(); }
function genOTP() { return Math.floor(100000+Math.random()*900000).toString(); }
function genToken() { return Math.random().toString(36).slice(2)+Date.now().toString(36); }
function formatDuration(ms) {
  if (!ms || ms<=0) return "—";
  const h=Math.floor(ms/3600000), m=Math.floor((ms%3600000)/60000);
  if(h>24) return `${Math.floor(h/24)}d ${h%24}h ${m}m`;
  return `${h}h ${m}m`;
}
function formatTime(ms) {
  if(ms<=0) return "Overdue";
  const h=Math.floor(ms/3600000), m=Math.floor((ms%3600000)/60000);
  if(h>24) return `${Math.floor(h/24)}d ${h%24}h`;
  return `${h}h ${m}m`;
}
function timeAgo(ts) {
  const d=Date.now()-ts, m=Math.floor(d/60000);
  if(m<1) return "just now"; if(m<60) return `${m}m ago`;
  const h=Math.floor(m/60); if(h<24) return `${h}h ago`;
  return `${Math.floor(h/24)}d ago`;
}
function fmtDate(ts) { return ts ? new Date(ts).toLocaleString("en-IN",{dateStyle:"medium",timeStyle:"short"}) : "—"; }
function isInstitutionEmail(value) { return /^[^\s@]+@jaipuria\.ac\.in$/i.test((value || "").trim()); }
function statusColor(status) { return {Open:"#6366f1",Assigned:"#0ea5e9","In Progress":"#f59e0b",Resolved:"#10b981",Closed:"#6b7280"}[status] || "#64748b"; }
function priorityColor(priority) { return {Low:"#64748b",Medium:"#f59e0b",High:"#f97316",Critical:"#ef4444"}[priority] || "#64748b"; }
function categoryLabel(id) { return CATEGORIES.find(c=>c.id===id)?.label || id || "—"; }
function staffName(id) { return STAFF_BASE.find(s=>s.id===Number(id))?.name || "Unassigned"; }
function genFeedbackId() { return "FDB-"+Date.now().toString(36).toUpperCase()+Math.random().toString(36).slice(2,5).toUpperCase(); }
function satisfactionColor(level) { return SATISFACTION_LEVELS.find(s=>s.id===level)?.color || "#64748b"; }
function cleanFeedbackRow(f) {
  return {
    "Feedback ID": f.id || "—",
    "Ticket ID": f.ticketId || "—",
    "User Name": f.name || "—",
    Email: f.email || "—",
    Department: f.dept || "—",
    "Service Category": f.category || "—",
    Rating: f.rating ? `${f.rating}/5` : "—",
    Satisfaction: f.satisfaction || "—",
    Recommendation: f.recommend || "—",
    "Feedback Message": f.message || "—",
    Suggestions: f.suggestions || "—",
    "Submitted At": fmtDate(f.createdAt),
    "Reviewed Status": f.reviewed ? "Reviewed" : "New",
  };
}
function cleanTicketRow(t) {
  return {
    "Ticket ID": t.id || "—",
    Name: t.name || "—",
    Email: t.email || "—",
    Department: t.dept || "—",
    Category: categoryLabel(t.category),
    Priority: t.priority || "—",
    Status: t.status || "—",
    "Assigned To": staffName(t.assigneeId),
    "Created At": fmtDate(t.createdAt),
    "Closed At": t.closedAt ? fmtDate(t.closedAt) : "—",
    "Resolution Time": t.closedAt && t.createdAt ? formatDuration(t.closedAt - t.createdAt) : "Active",
    Description: t.description || "—",
    "Closing Remarks": t.closingRemarks || "—",
  };
}
const STAFF_STATUS = {
  Online:{label:"Online",color:"#10b981"},
  Away:{label:"Away",color:"#f59e0b"},
  Busy:{label:"Busy",color:"#ef4444"},
  Offline:{label:"Offline",color:"#6b7280"},
};
function getStaffStatus(staffId,statuses={}) { return statuses[staffId] || "Online"; }
function getStaffPhoto(staffId,profiles={}) { return profiles[staffId]?.photo || ""; }
function getStaffInitials(staff) { return staff?.avatar || (staff?.name||"IT").split(" ").map(p=>p[0]).join("").slice(0,2).toUpperCase(); }
function StaffAvatar({staff,profiles={},statuses={},size=44,showStatus=false}) {
  const photo=getStaffPhoto(staff?.id,profiles);
  const status=getStaffStatus(staff?.id,statuses);
  const statusMeta=STAFF_STATUS[status] || STAFF_STATUS.Online;
  return (
    <div style={{position:"relative",width:size,height:size,flexShrink:0}}>
      <div style={{width:size,height:size,borderRadius:"50%",overflow:"hidden",background:staff?.color+"33",border:`2px solid ${staff?.color || "#6366f1"}`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:Math.max(11,size/3.4),fontWeight:800,color:staff?.color || "#818cf8"}}>
        {photo ? <img src={photo} alt={staff?.name || "Staff"} style={{width:"100%",height:"100%",objectFit:"cover"}}/> : getStaffInitials(staff)}
      </div>
      {showStatus&&<span title={statusMeta.label} style={{position:"absolute",right:0,bottom:0,width:Math.max(10,size/4),height:Math.max(10,size/4),borderRadius:"50%",background:statusMeta.color,border:"2px solid #0a0a0f"}}/>}
    </div>
  );
}
function StatusDot({status}) {
  const meta=STAFF_STATUS[status] || STAFF_STATUS.Online;
  return <span style={{display:"inline-flex",alignItems:"center",gap:6,fontSize:12,color:meta.color,fontWeight:700}}><span style={{width:8,height:8,borderRadius:"50%",background:meta.color,display:"inline-block"}}/>{meta.label}</span>;
}
// Password strength
function pwdStrength(pwd) {
  let s=0;
  if(pwd.length>=8) s++;
  if(/[A-Z]/.test(pwd)) s++;
  if(/[0-9]/.test(pwd)) s++;
  if(/[^A-Za-z0-9]/.test(pwd)) s++;
  if(pwd.length>=12) s++;
  return s;
}
function pwdLabel(s) { return ["","Weak","Fair","Good","Strong","Very Strong"][s]||""; }
function pwdColor(s) { return ["","#ef4444","#f97316","#f59e0b","#10b981","#6366f1"][s]||"#ef4444"; }

// ── EMAIL SIMULATION ──────────────────────────────────────────────────────
function simulateEmail(to, subject, body) {
  const log = DB.get("emaillog", []);
  log.unshift({ id: genToken().slice(0,8), to, subject, body, sentAt: Date.now() });
  DB.set("emaillog", log.slice(0,50));
}

function emailTicketCreated(ticket, assignee) {
  simulateEmail(ticket.email, `[${ticket.id}] IT Support Ticket Created`, `
Dear ${ticket.name},

Your IT support ticket has been raised successfully.

━━━━━━━━━━━━━━━━━━━━━━━━━
TICKET DETAILS
━━━━━━━━━━━━━━━━━━━━━━━━━
Ticket ID     : ${ticket.id}
Category      : ${CATEGORIES.find(c=>c.id===ticket.category)?.label}
Priority      : ${ticket.priority}
Status        : Open → Assigned
Assigned To   : ${assignee?.name} (${assignee?.role})
Created At    : ${fmtDate(ticket.createdAt)}
━━━━━━━━━━━━━━━━━━━━━━━━━

Our IT team will resolve your issue within the SLA timeline.

Regards,
Jaipuria Institute of Management IT Support Team
  `.trim());
}

function emailTicketAssigned(ticket, newAssignee, previousAssignee, assignedBy="Admin", remark="") {
  const category = categoryLabel(ticket.category);
  const base = `
Ticket ID    : ${ticket.id}
Category     : ${category}
Priority     : ${ticket.priority}
Status       : ${ticket.status}
Assigned To  : ${newAssignee?.name || "Unassigned"}
Assigned By  : ${assignedBy}
Updated At   : ${fmtDate(Date.now())}
${remark ? `Remark       : ${remark}` : ""}
  `.trim();

  simulateEmail(ticket.email, `[${ticket.id}] Ticket Assignment Updated`, `Dear ${ticket.name},\n\nYour IT support ticket assignment has been updated.\n\n${base}\n\nRegards,\nJaipuria Institute of Management IT Support Team`);

  if (newAssignee?.email) {
    simulateEmail(newAssignee.email, `[${ticket.id}] Ticket Assigned To You`, `A ticket has been assigned to you.\n\n${base}`);
  }

  if (previousAssignee?.email && previousAssignee.email !== newAssignee?.email) {
    simulateEmail(previousAssignee.email, `[${ticket.id}] Ticket Reassigned`, `This ticket has been reassigned from you.\n\n${base}`);
  }

  simulateEmail("admin@jaipuria.ac.in", `[${ticket.id}] Assignment Updated`, `Admin notification:\n\n${base}`);
}

// ── CSV / EXPORT HELPERS ──────────────────────────────────────────────────
function downloadExcel(data, filename) {
  const cleanData = data.map(row => {
    const cleaned = {};
    Object.keys(row).forEach(key => {
      cleaned[key] = typeof row[key] === "object" && row[key] !== null ? JSON.stringify(row[key]) : row[key];
    });
    return cleaned;
  });

  const worksheet = XLSX.utils.json_to_sheet(cleanData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Tickets");
  XLSX.writeFile(workbook, filename);
}

function countBy(items, values, getter) {
  return values.map(value => ({ label: value, value: items.filter(item => getter(item) === value).length }));
}

function drawSummaryCards(doc, cards, y) {
  const pageWidth = doc.internal.pageSize.getWidth();
  const margin = 12;
  const gap = 5;
  const cardW = (pageWidth - margin * 2 - gap * (cards.length - 1)) / cards.length;
  cards.forEach((card, i) => {
    const x = margin + i * (cardW + gap);
    doc.setFillColor(...card.rgb);
    doc.roundedRect(x, y, cardW, 18, 3, 3, "F");
    doc.setTextColor(255, 255, 255);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(12);
    doc.text(String(card.value), x + 4, y + 7);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(7.5);
    doc.text(card.label, x + 4, y + 14);
  });
}

function drawBarSummary(doc, title, rows, x, y, w, colors) {
  doc.setTextColor(30, 41, 59);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(10);
  doc.text(title, x, y);
  const total = Math.max(1, rows.reduce((sum, row) => sum + row.value, 0));
  let cy = y + 7;
  rows.forEach((row, idx) => {
    const color = colors[row.label] || colors[idx] || "#64748b";
    const rgb = hexToRgb(color);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(7.5);
    doc.setTextColor(71, 85, 105);
    doc.text(`${row.label} (${row.value})`, x, cy + 3);
    doc.setFillColor(226, 232, 240);
    doc.roundedRect(x + 35, cy, w - 43, 4, 1.5, 1.5, "F");
    doc.setFillColor(rgb.r, rgb.g, rgb.b);
    doc.roundedRect(x + 35, cy, Math.max(1, (w - 43) * (row.value / total)), 4, 1.5, 1.5, "F");
    cy += 8;
  });
  return cy;
}

function hexToRgb(hex) {
  const clean = (hex || "#64748b").replace("#", "");
  const value = parseInt(clean.length === 3 ? clean.split("").map(c => c + c).join("") : clean, 16);
  return { r: (value >> 16) & 255, g: (value >> 8) & 255, b: value & 255 };
}

function addPdfFooter(doc) {
  const pageCount = doc.internal.getNumberOfPages();
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  for (let i = 1; i <= pageCount; i++) {
    doc.setPage(i);
    doc.setDrawColor(226, 232, 240);
    doc.line(12, pageHeight - 12, pageWidth - 12, pageHeight - 12);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(8);
    doc.setTextColor(100, 116, 139);
    doc.text("Jaipuria Institute of Management - IT Helpdesk", 12, pageHeight - 7);
    doc.text(`Page ${i} of ${pageCount}`, pageWidth - 32, pageHeight - 7);
  }
}

function downloadPDF(data, filename, options = {}) {
  const tickets = options.sourceTickets || data || [];
  const rows = tickets.map(cleanTicketRow);
  const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4", compress: true });
  const pageWidth = doc.internal.pageSize.getWidth();
  const generatedAt = new Date().toLocaleString("en-IN", { dateStyle: "medium", timeStyle: "short" });
  const dateRange = options.dateFrom || options.dateTo
    ? `${options.dateFrom || "Beginning"} to ${options.dateTo || "Today"}`
    : "All dates";
  const closed = tickets.filter(t => t.status === "Closed" || t.status === "Resolved");
  const avgMs = closed.length ? closed.reduce((sum, t) => sum + ((t.closedAt || Date.now()) - t.createdAt), 0) / closed.length : 0;

  doc.setFillColor(15, 23, 42);
  doc.rect(0, 0, pageWidth, 34, "F");
  doc.setFillColor(99, 102, 241);
  doc.circle(18, 17, 8, "F");
  doc.setTextColor(255, 255, 255);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(17);
  doc.text("Jaipuria Institute of Management", 31, 14);
  doc.setFontSize(10);
  doc.setFont("helvetica", "normal");
  doc.text("IT Helpdesk Report", 31, 22);
  doc.setFontSize(8);
  doc.text(`Date range: ${dateRange}`, pageWidth - 82, 13);
  doc.text(`Generated: ${generatedAt}`, pageWidth - 82, 20);

  drawSummaryCards(doc, [
    { label: "Total Tickets", value: tickets.length, rgb: [99, 102, 241] },
    { label: "Open", value: tickets.filter(t => t.status === "Open").length, rgb: [14, 165, 233] },
    { label: "In Progress", value: tickets.filter(t => t.status === "In Progress").length, rgb: [245, 158, 11] },
    { label: "Closed/Resolved", value: closed.length, rgb: [16, 185, 129] },
    { label: "Critical", value: tickets.filter(t => t.priority === "Critical").length, rgb: [239, 68, 68] },
    { label: "Avg Resolution", value: avgMs ? formatDuration(avgMs) : "—", rgb: [139, 92, 246] },
  ], 42);

  const statusRows = countBy(tickets, STATUSES, t => t.status);
  const priorityRows = countBy(tickets, PRIORITIES, t => t.priority);
  const staffRows = STAFF_BASE.map(s => ({ label: s.name.split(" ").slice(0,2).join(" "), value: tickets.filter(t => t.assigneeId === s.id).length }));
  const categoryRows = CATEGORIES.map(c => ({ label: c.label, value: tickets.filter(t => t.category === c.id).length })).filter(r => r.value).slice(0, 8);

  drawBarSummary(doc, "Status Distribution", statusRows, 12, 72, 65, Object.fromEntries(STATUSES.map(s => [s, statusColor(s)])));
  drawBarSummary(doc, "Priority Distribution", priorityRows, 84, 72, 58, Object.fromEntries(PRIORITIES.map(p => [p, priorityColor(p)])));
  drawBarSummary(doc, "Staff Ticket Count", staffRows, 149, 72, 58, { 0: "#6366f1", 1: "#0ea5e9", 2: "#10b981" });
  drawBarSummary(doc, "Top Categories", categoryRows.length ? categoryRows : [{ label: "No category data", value: 0 }], 214, 72, 70, { 0: "#8b5cf6", 1: "#06b6d4", 2: "#f97316" });

  autoTable(doc, {
    startY: 122,
    head: [["Ticket ID", "Name", "Email", "Dept", "Category", "Priority", "Status", "Assigned To", "Created", "Closed", "Resolution", "Description", "Closing Remarks"]],
    body: rows.map(row => [
      row["Ticket ID"], row.Name, row.Email, row.Department, row.Category, row.Priority, row.Status,
      row["Assigned To"], row["Created At"], row["Closed At"], row["Resolution Time"], row.Description, row["Closing Remarks"]
    ]),
    theme: "grid",
    tableWidth: "auto",
    margin: { left: 10, right: 10, bottom: 16 },
    styles: {
      font: "helvetica",
      fontSize: 6.8,
      cellPadding: 1.7,
      overflow: "linebreak",
      valign: "top",
      lineColor: [226, 232, 240],
      lineWidth: 0.1,
      textColor: [30, 41, 59],
    },
    headStyles: { fillColor: [79, 70, 229], textColor: 255, fontStyle: "bold", fontSize: 7.2, halign: "center" },
    alternateRowStyles: { fillColor: [248, 250, 252] },
    columnStyles: {
      0: { cellWidth: 21, fontStyle: "bold" },
      1: { cellWidth: 22 },
      2: { cellWidth: 34 },
      3: { cellWidth: 24 },
      4: { cellWidth: 27 },
      5: { cellWidth: 17, halign: "center" },
      6: { cellWidth: 20, halign: "center" },
      7: { cellWidth: 25 },
      8: { cellWidth: 23 },
      9: { cellWidth: 23 },
      10: { cellWidth: 18, halign: "center" },
      11: { cellWidth: 38 },
      12: { cellWidth: 34 },
    },
    didParseCell: data => {
      if (data.section !== "body") return;
      if (data.column.index === 5) {
        const rgb = hexToRgb(priorityColor(data.cell.raw));
        data.cell.styles.textColor = [rgb.r, rgb.g, rgb.b];
        data.cell.styles.fontStyle = "bold";
      }
      if (data.column.index === 6) {
        const rgb = hexToRgb(statusColor(data.cell.raw));
        data.cell.styles.textColor = [rgb.r, rgb.g, rgb.b];
        data.cell.styles.fontStyle = "bold";
      }
    },
  });

  if (!rows.length) {
    doc.setFont("helvetica", "bold");
    doc.setFontSize(13);
    doc.setTextColor(100, 116, 139);
    doc.text("No tickets found for the selected filters.", 12, 130);
  }

  addPdfFooter(doc);
  doc.save(filename);
}

function downloadGenericReportPDF(title, rows, filename, options = {}) {
  const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4", compress: true });
  const pageWidth = doc.internal.pageSize.getWidth();
  const generatedAt = new Date().toLocaleString("en-IN", { dateStyle: "medium", timeStyle: "short" });
  const dateRange = `${options.dateFrom || "—"} to ${options.dateTo || "—"}`;
  const headers = rows.length ? Object.keys(rows[0]) : ["Report"];

  doc.setFillColor(15, 23, 42);
  doc.rect(0, 0, pageWidth, 32, "F");
  doc.setFillColor(99, 102, 241);
  doc.circle(18, 16, 7, "F");
  doc.setTextColor(255,255,255);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(16);
  doc.text("Jaipuria Institute of Management", 30, 13);
  doc.setFontSize(10);
  doc.setFont("helvetica", "normal");
  doc.text(title, 30, 21);
  doc.setFontSize(8);
  doc.text(`Date range: ${dateRange}`, pageWidth - 82, 12);
  doc.text(`Generated: ${generatedAt}`, pageWidth - 82, 19);

  autoTable(doc, {
    startY: 42,
    head: [headers],
    body: rows.map(row => headers.map(h => row[h] ?? "—")),
    theme: "grid",
    margin: { left: 12, right: 12, bottom: 16 },
    styles: { font: "helvetica", fontSize: 8, cellPadding: 2, overflow: "linebreak", valign: "top", lineColor: [226,232,240], lineWidth: 0.1, textColor: [30,41,59] },
    headStyles: { fillColor: [79,70,229], textColor: 255, fontStyle: "bold", halign: "center" },
    alternateRowStyles: { fillColor: [248,250,252] },
  });
  addPdfFooter(doc);
  doc.save(filename);
}

function downloadStaffPerformancePDF(staffRows, filename, options = {}) {
  const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4", compress: true });
  const pageWidth = doc.internal.pageSize.getWidth();
  const generatedAt = new Date().toLocaleString("en-IN", { dateStyle: "medium", timeStyle: "short" });
  const dateRange = `${options.dateFrom || "—"} to ${options.dateTo || "—"}`;
  const totalAssigned = staffRows.reduce((sum, s) => sum + s.assigned, 0);
  const totalResolved = staffRows.reduce((sum, s) => sum + s.resolved, 0);
  const top = [...staffRows].sort((a,b) => b.resolutionRate - a.resolutionRate || b.resolved - a.resolved)[0];

  doc.setFillColor(15, 23, 42);
  doc.rect(0, 0, pageWidth, 36, "F");
  doc.setFillColor(139, 92, 246);
  doc.circle(18, 18, 8, "F");
  doc.setTextColor(255,255,255);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(17);
  doc.text("Jaipuria Institute of Management", 31, 14);
  doc.setFontSize(11);
  doc.text("IT Staff Performance Report", 31, 23);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(8);
  doc.text(`Date range: ${dateRange}`, pageWidth - 82, 13);
  doc.text(`Generated: ${generatedAt}`, pageWidth - 82, 21);

  drawSummaryCards(doc, [
    { label: "Total Staff", value: staffRows.length, rgb: [99,102,241] },
    { label: "Assigned Tickets", value: totalAssigned, rgb: [14,165,233] },
    { label: "Resolved Tickets", value: totalResolved, rgb: [16,185,129] },
    { label: "Pending Tickets", value: Math.max(0,totalAssigned-totalResolved), rgb: [245,158,11] },
    { label: "Top Performer", value: top?.name?.split(" ")[0] || "—", rgb: [139,92,246] },
  ], 44);

  if (top) {
    doc.setFillColor(240, 253, 244);
    doc.setDrawColor(34, 197, 94);
    doc.roundedRect(12, 70, pageWidth - 24, 22, 3, 3, "FD");
    doc.setTextColor(22, 101, 52);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(11);
    doc.text(`Top Performer: ${top.name}`, 18, 79);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(8.5);
    doc.text(`${top.resolved} resolved of ${top.assigned} assigned tickets · ${top.resolutionRate}% resolution rate`, 18, 86);
  }

  autoTable(doc, {
    startY: 102,
    head: [["Staff Name", "Role", "Email", "Assigned Tickets", "Resolved Tickets", "Pending Tickets", "Resolution Rate", "Performance"]],
    body: staffRows.map(s => [s.name, s.role, s.email, s.assigned, s.resolved, s.pending, `${s.resolutionRate}%`, ""]),
    theme: "grid",
    margin: { left: 12, right: 12, bottom: 16 },
    styles: { font: "helvetica", fontSize: 8, cellPadding: 2, overflow: "linebreak", valign: "middle", lineColor: [226,232,240], lineWidth: 0.1, textColor: [30,41,59] },
    headStyles: { fillColor: [79,70,229], textColor: 255, fontStyle: "bold", halign: "center" },
    alternateRowStyles: { fillColor: [248,250,252] },
    columnStyles: { 0:{cellWidth:34,fontStyle:"bold"}, 1:{cellWidth:27}, 2:{cellWidth:48}, 3:{cellWidth:24,halign:"center"}, 4:{cellWidth:24,halign:"center"}, 5:{cellWidth:24,halign:"center"}, 6:{cellWidth:24,halign:"center",fontStyle:"bold"}, 7:{cellWidth:70} },
    didParseCell: data => {
      if (data.section === "body" && data.column.index === 6) {
        const rate = parseFloat(String(data.cell.raw));
        const color = rate >= 75 ? "#16a34a" : rate >= 40 ? "#ca8a04" : "#dc2626";
        const rgb = hexToRgb(color);
        data.cell.styles.textColor = [rgb.r, rgb.g, rgb.b];
      }
    },
    didDrawCell: data => {
      if (data.section === "body" && data.column.index === 7) {
        const s = staffRows[data.row.index];
        const color = s.resolutionRate >= 75 ? "#16a34a" : s.resolutionRate >= 40 ? "#ca8a04" : "#dc2626";
        const rgb = hexToRgb(color);
        const x = data.cell.x + 3;
        const y = data.cell.y + data.cell.height / 2 - 2;
        const w = data.cell.width - 8;
        doc.setFillColor(226,232,240);
        doc.roundedRect(x, y, w, 4, 1.5, 1.5, "F");
        doc.setFillColor(rgb.r, rgb.g, rgb.b);
        doc.roundedRect(x, y, Math.max(1, w * (s.resolutionRate / 100)), 4, 1.5, 1.5, "F");
      }
    },
  });

  addPdfFooter(doc);
  doc.save(filename);
}

function downloadFeedbackPDF(feedbackRows, filename, options = {}) {
  const items = feedbackRows || [];
  const rows = items.map(cleanFeedbackRow);
  const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4", compress: true });
  const pageWidth = doc.internal.pageSize.getWidth();
  const generatedAt = new Date().toLocaleString("en-IN", { dateStyle:"medium", timeStyle:"short" });
  const dateRange = `${options.dateFrom || "—"} to ${options.dateTo || "—"}`;
  const avgRating = items.length ? (items.reduce((sum,f)=>sum+Number(f.rating||0),0)/items.length).toFixed(1) : "0.0";
  const recommendYes = items.filter(f=>f.recommend === "Yes").length;
  const unread = items.filter(f=>!f.reviewed).length;

  doc.setFillColor(15, 23, 42);
  doc.rect(0, 0, pageWidth, 36, "F");
  doc.setFillColor(99, 102, 241);
  doc.circle(18, 18, 8, "F");
  doc.setTextColor(255,255,255);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(17);
  doc.text("Jaipuria Institute of Management", 31, 14);
  doc.setFontSize(11);
  doc.text("IT Services Feedback Report", 31, 23);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(8);
  doc.text(`Date range: ${dateRange}`, pageWidth - 82, 13);
  doc.text(`Generated: ${generatedAt}`, pageWidth - 82, 21);

  drawSummaryCards(doc, [
    { label:"Total Feedback", value:items.length, rgb:[99,102,241] },
    { label:"Average Rating", value:`${avgRating}/5`, rgb:[245,158,11] },
    { label:"Recommended", value:recommendYes, rgb:[16,185,129] },
    { label:"Needs Review", value:unread, rgb:[239,68,68] },
    { label:"Excellent", value:items.filter(f=>f.satisfaction==="Excellent").length, rgb:[14,165,233] },
  ], 44);

  const ratingRows = [5,4,3,2,1].map(n => ({ label:`${n} Star`, value:items.filter(f=>Number(f.rating)===n).length }));
  const satisfactionRows = SATISFACTION_LEVELS.map(s => ({ label:s.id, value:items.filter(f=>f.satisfaction===s.id).length }));
  const deptRows = DEPTS.map(d => ({ label:d, value:items.filter(f=>f.dept===d).length })).filter(r=>r.value).slice(0,7);
  const categoryRows = FEEDBACK_CATEGORIES.map(c => ({ label:c, value:items.filter(f=>f.category===c).length })).filter(r=>r.value).slice(0,7);

  drawBarSummary(doc, "Rating Distribution", ratingRows, 12, 74, 62, { "5 Star":"#10b981", "4 Star":"#22c55e", "3 Star":"#f59e0b", "2 Star":"#f97316", "1 Star":"#ef4444" });
  drawBarSummary(doc, "Satisfaction", satisfactionRows, 82, 74, 58, Object.fromEntries(SATISFACTION_LEVELS.map(s=>[s.id,s.color])));
  drawBarSummary(doc, "Department Summary", deptRows.length ? deptRows : [{label:"No data",value:0}], 147, 74, 64, {0:"#6366f1",1:"#0ea5e9",2:"#10b981",3:"#f59e0b"});
  drawBarSummary(doc, "Service Category", categoryRows.length ? categoryRows : [{label:"No data",value:0}], 220, 74, 64, {0:"#8b5cf6",1:"#06b6d4",2:"#f97316",3:"#10b981"});

  autoTable(doc, {
    startY: 126,
    head: [["Feedback ID", "Ticket ID", "User", "Email", "Dept", "Rating", "Satisfaction", "Recommend", "Feedback", "Suggestions", "Submitted", "Status"]],
    body: rows.map(row => [row["Feedback ID"], row["Ticket ID"], row["User Name"], row.Email, row.Department, row.Rating, row.Satisfaction, row.Recommendation, row["Feedback Message"], row.Suggestions, row["Submitted At"], row["Reviewed Status"]]),
    theme: "grid",
    margin: { left:10, right:10, bottom:16 },
    styles: { font:"helvetica", fontSize:7, cellPadding:1.7, overflow:"linebreak", valign:"top", lineColor:[226,232,240], lineWidth:0.1, textColor:[30,41,59] },
    headStyles: { fillColor:[79,70,229], textColor:255, fontStyle:"bold", halign:"center", fontSize:7.3 },
    alternateRowStyles: { fillColor:[248,250,252] },
    columnStyles: { 0:{cellWidth:21,fontStyle:"bold"}, 1:{cellWidth:22,fontStyle:"bold"}, 2:{cellWidth:24}, 3:{cellWidth:36}, 4:{cellWidth:24}, 5:{cellWidth:15,halign:"center"}, 6:{cellWidth:20,halign:"center"}, 7:{cellWidth:18,halign:"center"}, 8:{cellWidth:46}, 9:{cellWidth:40}, 10:{cellWidth:24}, 11:{cellWidth:16,halign:"center"} },
    didParseCell: data => {
      if (data.section !== "body") return;
      if (data.column.index === 6) {
        const rgb = hexToRgb(satisfactionColor(data.cell.raw));
        data.cell.styles.textColor = [rgb.r, rgb.g, rgb.b];
        data.cell.styles.fontStyle = "bold";
      }
      if (data.column.index === 11 && data.cell.raw === "New") {
        data.cell.styles.textColor = [239,68,68];
        data.cell.styles.fontStyle = "bold";
      }
    },
  });

  if (!rows.length) {
    doc.setFont("helvetica", "bold");
    doc.setFontSize(13);
    doc.setTextColor(100,116,139);
    doc.text("No feedback found for the selected filters.", 12, 132);
  }

  addPdfFooter(doc);
  doc.save(filename);
}
// ── GLOBAL CSS ─────────────────────────────────────────────────────────────
const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Sans:wght@300;400;500;600&display=swap');
*{box-sizing:border-box;margin:0;padding:0}
:root{--bg:#070913;--panel:rgba(13,18,35,.74);--panel2:rgba(20,28,52,.72);--stroke:rgba(148,163,184,.16);--stroke2:rgba(125,211,252,.22);--text:#e6edf7;--muted:rgba(226,232,240,.58);--purple:#8b5cf6;--blue:#3b82f6;--cyan:#06b6d4;--green:#10b981;--amber:#f59e0b;--red:#ef4444}
html{background:var(--bg)}
body{font-family:'DM Sans',sans-serif;color:var(--text);min-height:100vh;background:radial-gradient(circle at 16% 8%,rgba(139,92,246,.18),transparent 31%),radial-gradient(circle at 86% 14%,rgba(6,182,212,.14),transparent 30%),radial-gradient(circle at 70% 92%,rgba(16,185,129,.11),transparent 36%),linear-gradient(145deg,#050712 0%,#090d1b 44%,#07111d 100%);overflow-x:hidden}
body::before{content:"";position:fixed;inset:-25%;pointer-events:none;z-index:0;background:radial-gradient(circle at 22% 28%,rgba(99,102,241,.16),transparent 18%),radial-gradient(circle at 75% 22%,rgba(14,165,233,.12),transparent 19%),radial-gradient(circle at 58% 74%,rgba(16,185,129,.1),transparent 18%);filter:blur(34px);opacity:.9;animation:bgDrift 18s ease-in-out infinite alternate}
body::after{content:"";position:fixed;inset:0;pointer-events:none;z-index:0;background:linear-gradient(rgba(255,255,255,.025) 1px,transparent 1px),linear-gradient(90deg,rgba(255,255,255,.018) 1px,transparent 1px);background-size:44px 44px;mask-image:linear-gradient(to bottom,rgba(0,0,0,.42),transparent 72%)}
#root{position:relative;z-index:1;min-height:100vh}
::-webkit-scrollbar{width:8px;height:8px}
::-webkit-scrollbar-track{background:rgba(255,255,255,.035)}
::-webkit-scrollbar-thumb{background:linear-gradient(180deg,rgba(139,92,246,.65),rgba(6,182,212,.55));border-radius:999px;border:2px solid rgba(7,9,19,.72)}
input,select,textarea{background:linear-gradient(180deg,rgba(255,255,255,.075),rgba(255,255,255,.045));border:1px solid rgba(148,163,184,.18);color:var(--text);border-radius:12px;padding:11px 14px;font-family:'DM Sans',sans-serif;font-size:14px;width:100%;outline:none;transition:border-color .2s,background .2s,box-shadow .2s,transform .2s;box-shadow:inset 0 1px 0 rgba(255,255,255,.05),0 8px 24px rgba(0,0,0,.12)}
input:hover,select:hover,textarea:hover{border-color:rgba(125,211,252,.28);background:rgba(255,255,255,.075)}
input:focus,select:focus,textarea:focus{border-color:rgba(34,211,238,.72);background:rgba(8,145,178,.12);box-shadow:0 0 0 3px rgba(6,182,212,.14),0 14px 36px rgba(6,182,212,.08),inset 0 1px 0 rgba(255,255,255,.07)}
input::placeholder,textarea::placeholder{color:rgba(226,232,240,.34)}
select option{background:#0f172a;color:#e2e8f0}
button{cursor:pointer;font-family:'DM Sans',sans-serif;transition:transform .2s,box-shadow .2s,border-color .2s,background .2s,opacity .2s}
button:hover{filter:saturate(1.08)}
.glass{position:relative;background:linear-gradient(145deg,rgba(15,23,42,.76),rgba(17,24,39,.52));backdrop-filter:blur(22px) saturate(1.22);border:1px solid var(--stroke);border-radius:16px;box-shadow:0 18px 55px rgba(0,0,0,.28),inset 0 1px 0 rgba(255,255,255,.06);overflow:hidden}
.glass::before{content:"";position:absolute;inset:0;border-radius:inherit;pointer-events:none;background:linear-gradient(135deg,rgba(139,92,246,.16),transparent 30%,rgba(6,182,212,.1) 58%,rgba(16,185,129,.08));opacity:.72}
.glass>*{position:relative;z-index:1}
.glass2{position:relative;background:linear-gradient(145deg,rgba(30,41,59,.68),rgba(15,23,42,.55));backdrop-filter:blur(18px) saturate(1.15);border:1px solid rgba(148,163,184,.17);border-radius:12px;box-shadow:0 14px 38px rgba(0,0,0,.22),inset 0 1px 0 rgba(255,255,255,.055);overflow:hidden}
.glass2:hover,.glass:hover{border-color:rgba(125,211,252,.26);box-shadow:0 22px 64px rgba(0,0,0,.34),0 0 0 1px rgba(99,102,241,.08),inset 0 1px 0 rgba(255,255,255,.075)}
.glow-btn{background:linear-gradient(135deg,#6366f1 0%,#8b5cf6 42%,#06b6d4 100%);border:none;color:#fff;padding:12px 28px;border-radius:13px;font-size:15px;font-weight:700;letter-spacing:.25px;box-shadow:0 13px 32px rgba(99,102,241,.32),inset 0 1px 0 rgba(255,255,255,.18);position:relative;overflow:hidden}
.glow-btn::before{content:"";position:absolute;inset:-1px;background:linear-gradient(90deg,transparent,rgba(255,255,255,.28),transparent);transform:translateX(-120%);transition:transform .55s}
.glow-btn:hover{transform:translateY(-2px);box-shadow:0 18px 42px rgba(99,102,241,.38),0 8px 26px rgba(6,182,212,.16)}
.glow-btn:hover::before{transform:translateX(120%)}
.glow-btn:active{transform:translateY(0) scale(.99)}
.glow-btn:disabled{opacity:.5;transform:none;box-shadow:none;cursor:not-allowed}
.danger-btn{background:linear-gradient(135deg,#ef4444,#f97316);border:none;color:#fff;padding:10px 20px;border-radius:11px;font-size:14px;font-weight:700;box-shadow:0 12px 28px rgba(239,68,68,.24)}
.danger-btn:hover{transform:translateY(-1px);box-shadow:0 16px 34px rgba(239,68,68,.34)}
.success-btn{background:linear-gradient(135deg,#10b981,#06b6d4);border:none;color:#fff;padding:10px 20px;border-radius:11px;font-size:14px;font-weight:700;box-shadow:0 12px 28px rgba(16,185,129,.24)}
.success-btn:hover{transform:translateY(-1px);box-shadow:0 16px 34px rgba(16,185,129,.34)}
.tag{display:inline-flex;align-items:center;gap:5px;padding:4px 10px;border-radius:999px;font-size:12px;font-weight:650;border:1px solid rgba(255,255,255,.08);box-shadow:inset 0 1px 0 rgba(255,255,255,.06)}
.app-sidebar{background:linear-gradient(180deg,rgba(9,12,28,.98),rgba(12,18,36,.94))!important;border-right:1px solid rgba(125,211,252,.14)!important;box-shadow:12px 0 40px rgba(0,0,0,.28),inset -1px 0 0 rgba(255,255,255,.035)}
.app-sidebar nav button{position:relative;overflow:hidden}
.app-sidebar nav button:hover{background:rgba(99,102,241,.13)!important;color:#dbeafe!important;transform:translateX(2px)}
.app-sidebar nav button[style*="rgba(99,102,241,0.2)"]{background:linear-gradient(90deg,rgba(99,102,241,.28),rgba(6,182,212,.12))!important;color:#bfdbfe!important;box-shadow:inset 0 0 0 1px rgba(125,211,252,.16),0 8px 22px rgba(37,99,235,.12)}
.app-header{background:linear-gradient(90deg,rgba(7,10,24,.9),rgba(15,23,42,.78),rgba(8,47,73,.56))!important;border-bottom:1px solid rgba(125,211,252,.14)!important;box-shadow:0 10px 32px rgba(0,0,0,.24);backdrop-filter:blur(22px) saturate(1.2)!important}
.app-content{background:radial-gradient(circle at 20% 10%,rgba(99,102,241,.075),transparent 22%),radial-gradient(circle at 84% 22%,rgba(6,182,212,.055),transparent 24%)}
.modal-overlay{background:rgba(1,5,15,.78)!important;backdrop-filter:blur(13px) saturate(1.1)!important}
.modal-panel{border:1px solid rgba(125,211,252,.18)!important;box-shadow:0 28px 90px rgba(0,0,0,.55),0 0 0 1px rgba(139,92,246,.1)!important}
.modal-header{background:linear-gradient(90deg,rgba(99,102,241,.14),rgba(6,182,212,.08),transparent);border-bottom:1px solid rgba(125,211,252,.14)!important}
.pwd-input-wrap{position:relative}
.pwd-input-wrap input{padding-right:44px}
.pwd-toggle{position:absolute;right:12px;top:50%;transform:translateY(-50%);background:none;border:none;color:rgba(226,232,240,.46);font-size:18px;padding:0;line-height:1;cursor:pointer}
.pwd-toggle:hover{color:rgba(226,232,240,.86)}
@keyframes fadeUp{from{opacity:0;transform:translateY(20px)}to{opacity:1;transform:translateY(0)}}
@keyframes fadeIn{from{opacity:0}to{opacity:1}}
@keyframes pulse{0%,100%{opacity:1}50%{opacity:.5}}
@keyframes spin{to{transform:rotate(360deg)}}
@keyframes bounce{0%,100%{transform:scale(1)}50%{transform:scale(1.15)}}
@keyframes slideDown{from{opacity:0;transform:translateY(-10px)}to{opacity:1;transform:translateY(0)}}
@keyframes confetti{0%{transform:translateY(0) rotate(0);opacity:1}100%{transform:translateY(100px) rotate(720deg);opacity:0}}
@keyframes bgDrift{0%{transform:translate3d(-1%,0,0) scale(1)}100%{transform:translate3d(1.5%,1%,0) scale(1.04)}}
.fade-up{animation:fadeUp .4s ease forwards}
.fade-in{animation:fadeIn .3s ease forwards}
.pulse{animation:pulse 2s infinite}
.spin{animation:spin 1s linear infinite}
.slide-down{animation:slideDown .3s ease forwards}
/* Premium colorful dark theme overrides */
body{background:radial-gradient(circle at 9% 10%,rgba(168,85,247,.46),transparent 34%),radial-gradient(circle at 88% 8%,rgba(14,165,233,.38),transparent 34%),radial-gradient(circle at 78% 78%,rgba(16,185,129,.3),transparent 36%),radial-gradient(circle at 18% 96%,rgba(59,130,246,.28),transparent 32%),linear-gradient(135deg,#081126 0%,#0b1740 34%,#171553 63%,#08243a 100%)!important;color:#edf5ff!important}
body::before{opacity:1!important;filter:blur(22px)!important;background:radial-gradient(circle at 12% 20%,rgba(168,85,247,.48),transparent 18%),radial-gradient(circle at 82% 18%,rgba(34,211,238,.4),transparent 19%),radial-gradient(circle at 64% 72%,rgba(34,197,94,.28),transparent 18%),radial-gradient(circle at 34% 84%,rgba(37,99,235,.36),transparent 19%)!important;animation:bgDrift 13s ease-in-out infinite alternate}
body::after{opacity:.75!important;background:linear-gradient(rgba(125,211,252,.045) 1px,transparent 1px),linear-gradient(90deg,rgba(168,85,247,.035) 1px,transparent 1px)!important;background-size:42px 42px!important}
.app-content{background:radial-gradient(circle at 13% 2%,rgba(168,85,247,.2),transparent 28%),radial-gradient(circle at 88% 14%,rgba(34,211,238,.17),transparent 30%),radial-gradient(circle at 70% 100%,rgba(16,185,129,.14),transparent 34%)!important}
.glass{background:linear-gradient(145deg,rgba(42,52,126,.86),rgba(24,34,82,.76) 42%,rgba(9,79,104,.66))!important;border:1px solid rgba(125,211,252,.43)!important;box-shadow:0 26px 76px rgba(0,0,0,.36),0 0 0 1px rgba(168,85,247,.26),0 0 52px rgba(6,182,212,.17),inset 0 1px 0 rgba(255,255,255,.16)!important;transition:transform .24s ease,border-color .24s ease,box-shadow .24s ease,background .24s ease!important}
.glass::before{background:linear-gradient(135deg,rgba(168,85,247,.36),rgba(37,99,235,.1) 34%,rgba(6,182,212,.27) 64%,rgba(16,185,129,.2))!important;opacity:1!important}
.glass::after{content:"";position:absolute;left:0;right:0;top:0;height:1px;background:linear-gradient(90deg,transparent,rgba(216,180,254,.92),rgba(34,211,238,1),rgba(52,211,153,.9),transparent);pointer-events:none;z-index:2}
.glass2{background:linear-gradient(145deg,rgba(50,58,135,.8),rgba(22,33,79,.74) 45%,rgba(10,91,116,.62))!important;border:1px solid rgba(148,163,184,.34)!important;box-shadow:0 20px 54px rgba(0,0,0,.3),0 0 42px rgba(99,102,241,.15),inset 0 1px 0 rgba(255,255,255,.14)!important;transition:transform .22s ease,border-color .22s ease,box-shadow .22s ease!important}
.glass2::before{content:"";position:absolute;inset:0;pointer-events:none;background:linear-gradient(140deg,rgba(139,92,246,.22),transparent 40%,rgba(34,211,238,.18),rgba(16,185,129,.08));opacity:1}
.glass2>*{position:relative;z-index:1}
.glass:hover,.glass2:hover{transform:translateY(-3px);border-color:rgba(34,211,238,.72)!important;box-shadow:0 34px 92px rgba(0,0,0,.42),0 0 0 1px rgba(168,85,247,.34),0 0 68px rgba(34,211,238,.26),0 0 36px rgba(16,185,129,.1),inset 0 1px 0 rgba(255,255,255,.2)!important}
.app-sidebar{background:radial-gradient(circle at 18% 8%,rgba(168,85,247,.4),transparent 33%),radial-gradient(circle at 100% 34%,rgba(6,182,212,.28),transparent 34%),radial-gradient(circle at 45% 96%,rgba(16,185,129,.14),transparent 34%),linear-gradient(180deg,#0b1230 0%,#151758 48%,#08283c 100%)!important;border-right:1px solid rgba(34,211,238,.48)!important;box-shadow:20px 0 62px rgba(0,0,0,.46),inset -1px 0 0 rgba(255,255,255,.1)!important}
.app-sidebar nav button{border:1px solid transparent!important;margin-bottom:3px!important;color:rgba(226,239,255,.74)!important}
.app-sidebar nav button:hover{background:linear-gradient(90deg,rgba(168,85,247,.34),rgba(6,182,212,.22))!important;border-color:rgba(125,211,252,.36)!important;color:#fff!important;box-shadow:0 12px 28px rgba(6,182,212,.16)!important;transform:translateX(3px)!important}
.app-sidebar nav button[style*="rgba(99,102,241,0.2)"],.app-sidebar nav button[style*="#6366f1"]{background:linear-gradient(100deg,#8b5cf6 0%,#2563eb 44%,#06b6d4 82%,#10b981 120%)!important;color:#fff!important;border-color:rgba(255,255,255,.28)!important;box-shadow:0 16px 38px rgba(37,99,235,.38),0 0 32px rgba(6,182,212,.28),inset 0 1px 0 rgba(255,255,255,.2)!important}
.app-header{background:linear-gradient(90deg,rgba(13,23,56,.94),rgba(34,34,105,.86) 40%,rgba(9,87,117,.78))!important;border-bottom:1px solid rgba(34,211,238,.5)!important;box-shadow:0 18px 48px rgba(0,0,0,.34),0 0 42px rgba(99,102,241,.18)!important}
.modal-panel{background:linear-gradient(145deg,rgba(30,35,93,.96),rgba(9,72,97,.86))!important;border-color:rgba(34,211,238,.48)!important;box-shadow:0 38px 118px rgba(0,0,0,.62),0 0 82px rgba(99,102,241,.24),0 0 38px rgba(6,182,212,.15)!important}
.modal-header{background:linear-gradient(90deg,rgba(139,92,246,.3),rgba(37,99,235,.22),rgba(6,182,212,.18))!important;border-bottom-color:rgba(34,211,238,.38)!important}
.glow-btn{background:linear-gradient(135deg,#9333ea 0%,#4f46e5 30%,#2563eb 55%,#06b6d4 82%,#10b981 115%)!important;color:#fff!important;box-shadow:0 18px 48px rgba(37,99,235,.42),0 0 34px rgba(6,182,212,.3),inset 0 1px 0 rgba(255,255,255,.28)!important;text-shadow:0 1px 12px rgba(0,0,0,.28)}
.glow-btn:hover{transform:translateY(-3px)!important;box-shadow:0 26px 64px rgba(147,51,234,.38),0 14px 44px rgba(6,182,212,.36),0 0 24px rgba(16,185,129,.16)!important}
.success-btn{background:linear-gradient(135deg,#059669,#10b981,#06b6d4)!important;box-shadow:0 18px 44px rgba(16,185,129,.38),0 0 24px rgba(6,182,212,.16)!important}
.danger-btn{background:linear-gradient(135deg,#ef4444,#f97316,#f59e0b)!important;box-shadow:0 18px 44px rgba(239,68,68,.38)!important}
input,select,textarea{background:linear-gradient(180deg,rgba(18,31,74,.86),rgba(11,55,82,.68))!important;border-color:rgba(125,211,252,.34)!important;color:#f3f8ff!important;box-shadow:inset 0 1px 0 rgba(255,255,255,.11),0 12px 34px rgba(0,0,0,.18)!important}
input:hover,select:hover,textarea:hover{border-color:rgba(34,211,238,.56)!important;background:linear-gradient(180deg,rgba(25,43,96,.88),rgba(13,71,101,.72))!important}
input:focus,select:focus,textarea:focus{border-color:rgba(34,211,238,1)!important;background:linear-gradient(180deg,rgba(10,83,112,.78),rgba(43,48,126,.7))!important;box-shadow:0 0 0 3px rgba(34,211,238,.24),0 0 0 7px rgba(168,85,247,.12),0 20px 52px rgba(6,182,212,.22)!important}
.tag{background:linear-gradient(135deg,rgba(139,92,246,.24),rgba(6,182,212,.18))!important;border-color:rgba(125,211,252,.3)!important;color:#eef6ff!important}
@media (max-width:768px){body{background:radial-gradient(circle at 20% 8%,rgba(168,85,247,.38),transparent 38%),radial-gradient(circle at 82% 18%,rgba(34,211,238,.3),transparent 36%),linear-gradient(145deg,#081126,#10194a 54%,#08283c)!important}.glass,.glass2{box-shadow:0 18px 50px rgba(0,0,0,.34),0 0 34px rgba(6,182,212,.14),inset 0 1px 0 rgba(255,255,255,.14)!important}.app-sidebar{background:radial-gradient(circle at 25% 6%,rgba(168,85,247,.44),transparent 34%),linear-gradient(180deg,#0b1230,#17175b 55%,#08283c)!important}.glow-btn{box-shadow:0 16px 38px rgba(37,99,235,.38),0 0 26px rgba(6,182,212,.26)!important}}
@media (max-width:768px){
  html,body,#root{width:100%;max-width:100%;overflow-x:hidden}
  body::before{inset:-45%;filter:blur(28px);opacity:.68}
  .app-shell{display:block!important;min-height:100vh!important;width:100%!important;overflow-x:hidden!important}
  .app-main{width:100%!important;min-width:0!important;margin:0!important}
  .app-sidebar{position:fixed!important;top:0!important;left:0!important;width:min(82vw,300px)!important;max-width:300px!important;height:100dvh!important;z-index:80!important;transform:translateX(-105%)!important;transition:transform .25s ease!important;box-shadow:18px 0 44px rgba(0,0,0,.5)!important;border-right:1px solid rgba(125,211,252,.18)!important}
  .app-sidebar.mobile-open{transform:translateX(0)!important}
  .sidebar-overlay{display:block!important;position:fixed!important;inset:0!important;background:rgba(0,0,0,.6)!important;backdrop-filter:blur(5px)!important;z-index:70!important}
  .app-header{padding:10px 12px!important;gap:10px!important;position:sticky!important;top:0!important;z-index:40!important;width:100%!important;max-width:100vw!important}
  .header-identity{min-width:0!important;flex:1!important;overflow:hidden!important}
  .header-identity span{min-width:0!important;max-width:54vw!important;overflow:hidden!important;text-overflow:ellipsis!important;white-space:nowrap!important;display:block!important}
  .header-actions{gap:7px!important;flex-shrink:0!important}
  .header-actions span{font-size:11px!important}
  .header-actions button{padding:6px 10px!important;max-width:96px!important;overflow:hidden!important;text-overflow:ellipsis!important;white-space:nowrap!important}
  .app-content{padding:16px 12px!important;width:100%!important;max-width:100vw!important;overflow-x:hidden!important}
  .app-content h1,.app-content h2{font-size:20px!important;line-height:1.25!important}
  .app-content h3{font-size:16px!important}
  .app-content [style*="grid-template-columns"]{grid-template-columns:1fr!important}
  .app-content [style*="minmax(280px"],.app-content [style*="minmax(290px"],.app-content [style*="minmax(300px"],.app-content [style*="minmax(320px"]{grid-template-columns:1fr!important}
  .app-content [style*="display: flex"]{max-width:100%!important}
  .app-content button,.app-content .glow-btn,.app-content .success-btn,.app-content .danger-btn{max-width:100%;white-space:normal!important}
  .app-content input,.app-content select,.app-content textarea{min-width:0!important;max-width:100%!important}
  .glass,.glass2{max-width:100%!important;border-radius:14px!important}
  .modal-overlay{padding:8px!important;align-items:stretch!important;justify-content:center!important}
  .modal-panel{max-width:100%!important;width:100%!important;max-height:calc(100dvh - 16px)!important;border-radius:14px!important}
  .modal-header{padding:14px 14px 12px!important;gap:10px!important}
  .modal-header h2{font-size:16px!important;line-height:1.25!important;overflow:hidden!important;text-overflow:ellipsis!important;white-space:nowrap!important}
  .modal-body{padding:14px!important;overflow-y:auto!important}
  .staff-profile-menu{position:fixed!important;left:12px!important;right:12px!important;top:58px!important;width:auto!important;min-width:0!important;max-width:none!important;z-index:90!important}
}
@media (max-width:480px){
  .app-sidebar{width:86vw!important}
  .app-content{padding:12px 10px!important}
  .app-header{padding:9px 10px!important}
  .header-identity span{max-width:42vw!important;font-size:12px!important}
  .header-actions .pulse,.header-actions span{display:none!important}
  .header-actions button{font-size:12px!important;padding:6px 9px!important}
  .app-content [style*="padding: 24px"],.app-content [style*="padding:24px"]{padding:16px!important}
  .app-content [style*="gap: 24px"],.app-content [style*="gap:24px"]{gap:16px!important}
  .app-content .tag{font-size:11px!important;padding:3px 8px!important}
  .modal-overlay{padding:0!important}
  .modal-panel{min-height:100dvh!important;max-height:100dvh!important;border-radius:0!important}
/* ===== FORCE PREMIUM OVERRIDE ===== */

body{
  background:
    radial-gradient(circle at top left, rgba(99,102,241,.35), transparent 35%),
    radial-gradient(circle at top right, rgba(6,182,212,.28), transparent 32%),
    radial-gradient(circle at bottom left, rgba(16,185,129,.20), transparent 38%),
    #050816 !important;
}

.glass,
.glass2,
.ticket-card,
.stat-card,
.dashboard-card{
  background:
    linear-gradient(135deg,
      rgba(15,23,42,.92),
      rgba(30,41,59,.72)
    ) !important;

  border:1px solid rgba(34,211,238,.35)!important;

  box-shadow:
    0 0 0 1px rgba(99,102,241,.15),
    0 20px 60px rgba(0,0,0,.45),
    0 0 40px rgba(6,182,212,.12)!important;

  backdrop-filter:blur(20px)!important;
}

.glass:hover,
.glass2:hover,
.ticket-card:hover,
.stat-card:hover,
.dashboard-card:hover{
  transform:translateY(-4px);
  border-color:rgba(34,211,238,.65)!important;

  box-shadow:
    0 0 0 1px rgba(99,102,241,.22),
    0 24px 80px rgba(0,0,0,.55),
    0 0 55px rgba(6,182,212,.22)!important;
}

.glow-btn,
button{
  background:
    linear-gradient(
      135deg,
      #8b5cf6 0%,
      #3b82f6 50%,
      #06b6d4 100%
    ) !important;

  color:white!important;
}

.app-sidebar{
  background:
    linear-gradient(
      180deg,
      rgba(23, 2, 5, 0.98),
      rgba(15,23,42,.96),
      rgba(30,41,59,.92)
    ) !important;

  border-right:1px solid rgba(34,211,238,.28)!important;
}

.app-header{
  background:
    linear-gradient(
      135deg,
      rgba(160, 182, 233, 0.92),
      rgba(30,41,59,.70)
    ) !important;

  border-bottom:1px solid rgba(34,211,238,.25)!important;

  backdrop-filter:blur(18px)!important;
}

input,
select,
textarea{
  background:rgba(2,6,23,.72)!important;

  border:1px solid rgba(148,163,184,.28)!important;
}

input:focus,
select:focus,
textarea:focus{
  border-color:#06b6d4!important;

  box-shadow:
    0 0 0 3px rgba(6,182,212,.18),
    0 0 30px rgba(99,102,241,.22)!important;
}
  }`;

// ── TOAST ─────────────────────────────────────────────────────────────────
function Toast({ toasts, remove }) {
  return (
    
    <div style={{position:"fixed",top:20,right:20,zIndex:9999,display:"flex",flexDirection:"column",gap:10,maxWidth:340}}>
      {toasts.map(t=>(
        <div key={t.id} className="fade-up" style={{
          background:t.type==="success"?"rgba(16,185,129,0.18)":t.type==="error"?"rgba(239,68,68,0.18)":t.type==="email"?"rgba(99,102,241,0.18)":"rgba(245,158,11,0.18)",
          border:`1px solid ${t.type==="success"?"rgba(16,185,129,0.4)":t.type==="error"?"rgba(239,68,68,0.4)":t.type==="email"?"rgba(99,102,241,0.4)":"rgba(245,158,11,0.4)"}`,
          color:"#fff",padding:"12px 16px",borderRadius:12,fontSize:14,fontWeight:500,
          display:"flex",alignItems:"flex-start",gap:10,backdropFilter:"blur(20px)"
        }}>
          <span style={{fontSize:18,flexShrink:0}}>{t.type==="success"?"✅":t.type==="error"?"❌":t.type==="email"?"📧":"⚠️"}</span>
          <span style={{flex:1,lineHeight:1.4}}>{t.msg}</span>
          <button onClick={()=>remove(t.id)} style={{background:"none",border:"none",color:"rgba(255,255,255,0.5)",fontSize:18,flexShrink:0,lineHeight:1}}>×</button>
        </div>
      ))}
    </div>
  );
}

function useToast() {
  const [toasts,setToasts]=useState([]);
  const toast=(msg,type="info")=>{
    const id=Date.now()+Math.random();
    setToasts(ts=>[...ts,{id,msg,type}]);
    setTimeout(()=>setToasts(ts=>ts.filter(t=>t.id!==id)),5000);
  };
  const remove=(id)=>setToasts(ts=>ts.filter(t=>t.id!==id));
  return {toasts,toast,remove};
}

// ── PASSWORD INPUT ────────────────────────────────────────────────────────
function PwdInput({
  value,
  onChange,
  placeholder="Password",
  showStrength=false,
  id,
  autoComplete,
  name,
  onKeyDown
}) {

  const [show,setShow]=useState(false);

  const s = showStrength ? pwdStrength(value) : 0;

  return (
    <div>

      <div className="pwd-input-wrap">

        <input
  id={id}
  type={show ? "text" : "password"}
  value={value}
  name={name || "not-password"}
  autoComplete={autoComplete || "off"}
  data-lpignore="true"
  onChange={e=>onChange(e.target.value)}
  onKeyDown={onKeyDown}
  placeholder={placeholder}
/>

        <button
          type="button"
          className="pwd-toggle"
          onClick={()=>setShow(v=>!v)}
          title={show ? "Hide password" : "Show password"}
        >
          {show ? "🙈" : "👁"}
        </button>

      </div>

      {showStrength && value.length > 0 && (
        <div style={{marginTop:8}}>

          <div
            style={{
              height:4,
              background:"rgba(255,255,255,0.08)",
              borderRadius:2,
              overflow:"hidden"
            }}
          >
            <div
              style={{
                height:"100%",
                width:`${(s/5)*100}%`,
                background:pwdColor(s),
                borderRadius:2,
                transition:"all .3s"
              }}
            />
          </div>

          <div
            style={{
              fontSize:12,
              marginTop:4,
              color:pwdColor(s),
              fontWeight:500
            }}
          >
            {pwdLabel(s)}
            {s < 3
              ? " — Use 8+ chars, uppercase, numbers & symbols"
              : ""}
          </div>

        </div>
      )}

    </div>
  );
}

// ── MODAL ─────────────────────────────────────────────────────────────────
function Modal({ title, children, onClose, wide=false }) {
  return (
    <div className="modal-overlay" style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.75)",backdropFilter:"blur(8px)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}
      onClick={e=>e.target===e.currentTarget&&onClose&&onClose()}>
      <div className="glass fade-up modal-panel" style={{width:"100%",maxWidth:wide?860:720,maxHeight:"92vh",overflow:"hidden",display:"flex",flexDirection:"column"}}>
        <div className="modal-header" style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"20px 24px 16px",borderBottom:"1px solid rgba(255,255,255,0.08)",flexShrink:0}}>
          <h2 style={{fontFamily:"Syne",fontSize:18,fontWeight:700,color:"#e2e8f0"}}>{title}</h2>
          {onClose&&<button onClick={onClose} style={{background:"rgba(255,255,255,0.08)",border:"1px solid rgba(255,255,255,0.1)",color:"#e2e8f0",width:32,height:32,borderRadius:8,fontSize:20,display:"flex",alignItems:"center",justifyContent:"center",lineHeight:1}}>×</button>}
        </div>
        <div className="modal-body" style={{padding:"20px 24px",overflowY:"auto",flex:1}}>{children}</div>
      </div>
    </div>
  );
}

// ── STATUS / PRIORITY BADGES ─────────────────────────────────────────────
function StatusBadge({status}) {
  const c={Open:{bg:"rgba(99,102,241,0.2)",col:"#818cf8"},Assigned:{bg:"rgba(14,165,233,0.2)",col:"#38bdf8"},"In Progress":{bg:"rgba(245,158,11,0.2)",col:"#fbbf24"},Resolved:{bg:"rgba(16,185,129,0.2)",col:"#34d399"},Closed:{bg:"rgba(107,114,128,0.2)",col:"#9ca3af"}}[status]||{bg:"rgba(99,102,241,0.2)",col:"#818cf8"};
  return <span className="tag" style={{background:c.bg,color:c.col}}>● {status}</span>;
}
function PriorityBadge({p}) {
  const c={Low:"#94a3b8",Medium:"#fbbf24",High:"#f97316",Critical:"#ef4444"}[p]||"#94a3b8";
  return <span className="tag" style={{background:`${c}20`,color:c}}>▲ {p}</span>;
}

// ── SLA TIMER ─────────────────────────────────────────────────────────────
function TimerBadge({ticket}) {
  const [now,setNow]=useState(Date.now());
  useEffect(()=>{const iv=setInterval(()=>setNow(Date.now()),60000);return()=>clearInterval(iv);},[]);
  if(ticket.status==="Closed"||ticket.status==="Resolved") return <span className="tag" style={{background:"rgba(16,185,129,0.15)",color:"#10b981"}}>✅ {ticket.closedAt?formatDuration(ticket.closedAt-ticket.createdAt):"Resolved"}</span>;
  const slaMs=SLA_HOURS[ticket.priority]*3600000;
  const elapsed=now-ticket.createdAt;
  const remaining=slaMs-elapsed;
  const pct=Math.min(100,(elapsed/slaMs)*100);
  const col=pct<50?"#10b981":pct<80?"#f59e0b":"#ef4444";
  return (
    <div style={{display:"flex",flexDirection:"column",gap:4}}>
      <div style={{display:"flex",justifyContent:"space-between",fontSize:11}}>
        <span style={{color:"rgba(226,232,240,0.5)"}}>SLA Remaining</span>
        <span style={{color:col,fontWeight:600}}>{formatTime(remaining)}</span>
      </div>
      <div style={{height:4,background:"rgba(255,255,255,0.08)",borderRadius:2}}>
        <div style={{height:"100%",width:`${pct}%`,background:col,borderRadius:2,transition:"width .3s"}}/>
      </div>
    </div>
  );
}

// ── STAT CARD ─────────────────────────────────────────────────────────────
function StatCard({label,value,icon,color,sub}) {
  return (
    <div className="glass" style={{padding:"20px 22px",display:"flex",flexDirection:"column",gap:8}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
        <span style={{fontSize:13,color:"rgba(226,232,240,0.5)",fontWeight:500}}>{label}</span>
        <div style={{fontSize:24}}>{icon}</div>
      </div>
      <div style={{fontSize:32,fontWeight:700,fontFamily:"Syne",color,letterSpacing:"-1px"}}>{value}</div>
      {sub&&<div style={{fontSize:12,color:"rgba(226,232,240,0.4)"}}>{sub}</div>}
    </div>
  );
}

// ── AUDIT TIMELINE ────────────────────────────────────────────────────────
function AuditTimeline({timeline}) {
  const icons={Created:"🆕",Assigned:"👤",Reassigned:"🔄","Status changed":"📋",Commented:"💬",Closed:"✅",default:"📌"};
  return (
    <div style={{display:"flex",flexDirection:"column",gap:0}}>
      {(timeline||[]).map((ev,i)=>{
        const icon=Object.keys(icons).find(k=>ev.action.startsWith(k))||"default";
        return (
          <div key={i} style={{display:"flex",gap:14,alignItems:"flex-start",position:"relative",paddingBottom:i<timeline.length-1?16:0}}>
            {i<timeline.length-1&&<div style={{position:"absolute",left:15,top:32,width:2,height:"calc(100% - 10px)",background:"rgba(255,255,255,0.07)"}}/>}
            <div style={{width:30,height:30,borderRadius:"50%",background:"rgba(99,102,241,0.15)",border:"1px solid rgba(99,102,241,0.3)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:14,flexShrink:0}}>
              {icons[icon]}
            </div>
            <div style={{paddingTop:4}}>
              <div style={{fontSize:13,color:"#e2e8f0",fontWeight:500}}>{ev.action}</div>
              {ev.remark&&<div style={{fontSize:12,color:"rgba(226,232,240,0.5)",marginTop:2,fontStyle:"italic"}}>"{ev.remark}"</div>}
              <div style={{fontSize:11,color:"rgba(226,232,240,0.35)",marginTop:3}}>{fmtDate(ev.at)} · {ev.by}</div>
            </div>
          </div>
        );
      })}
    </div>
  );
}

// ── CLOSE TICKET DIALOG ───────────────────────────────────────────────────
function CloseTicketDialog({ticket,onClose,onConfirm}) {
  const [remarks,setRemarks]=useState("");
  const [loading,setLoading]=useState(false);
  const handleClose=async()=>{
    if(!remarks.trim()){return;}
    setLoading(true);
    await new Promise(r=>setTimeout(r,800));
    try {
      await Promise.resolve(onConfirm(remarks));
    } finally {
      setLoading(false);
    }
  };
  const duration=formatDuration(Date.now()-ticket.createdAt);
  return (
    <div style={{display:"flex",flexDirection:"column",gap:20}}>
      {/* Summary */}
      <div className="glass" style={{padding:"16px 18px"}}>
        <div style={{fontSize:12,color:"rgba(226,232,240,0.5)",marginBottom:12,letterSpacing:".5px"}}>TICKET SUMMARY</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
          {[["Ticket ID",ticket.id],["Category",CATEGORIES.find(c=>c.id===ticket.category)?.label],["Opened",fmtDate(ticket.createdAt)],["Active for",duration]].map(([l,v])=>(
            <div key={l}><div style={{fontSize:11,color:"rgba(226,232,240,0.4)"}}>{l}</div><div style={{fontSize:13,fontWeight:600,color:"#e2e8f0",marginTop:2}}>{v}</div></div>
          ))}
        </div>
      </div>
      {/* Closing remarks */}
      <div>
        <label style={{fontSize:13,color:"rgba(226,232,240,0.7)",marginBottom:8,display:"block",fontWeight:500}}>Closing Remarks <span style={{color:"#ef4444"}}>*</span></label>
        <textarea rows={4} placeholder="Describe how the issue was resolved..." value={remarks} onChange={e=>setRemarks(e.target.value)} style={{resize:"vertical"}} />
        {!remarks.trim()&&<div style={{fontSize:12,color:"rgba(239,68,68,0.8)",marginTop:4}}>Remarks required to close ticket</div>}
      </div>
      {/* Timeline preview */}
      <div className="glass" style={{padding:"14px 16px"}}>
        <div style={{display:"flex",flexWrap:"wrap",gap:16,justifyContent:"space-around",textAlign:"center"}}>
          {[["🕐","Open Time",fmtDate(ticket.createdAt)],["🕑","Close Time",fmtDate(Date.now())],["⏱","Duration",duration]].map(([ic,l,v])=>(
            <div key={l}><div style={{fontSize:20}}>{ic}</div><div style={{fontSize:11,color:"rgba(226,232,240,0.4)",marginTop:4}}>{l}</div><div style={{fontSize:12,fontWeight:600,color:"#e2e8f0",marginTop:2}}>{v}</div></div>
          ))}
        </div>
      </div>
      <div style={{display:"flex",gap:10,justifyContent:"flex-end"}}>
        <button onClick={onClose} style={{background:"rgba(255,255,255,0.06)",border:"1px solid rgba(255,255,255,0.1)",color:"#e2e8f0",padding:"10px 20px",borderRadius:10,fontSize:14}}>Cancel</button>
        <button className="success-btn" onClick={handleClose} disabled={!remarks.trim()||loading}>
          {loading?"⏳ Closing...":"✅ Close Ticket"}
        </button>
      </div>
    </div>
  );
}

// ── TICKET DETAIL ─────────────────────────────────────────────────────────
function TicketDetail({ticketId,tickets,setTickets,onClose,isAdmin,isStaff,staffId,staffName,toast,staffProfiles={},staffStatuses={}}) {
  const ticket=tickets.find(t=>t.id===ticketId)||{};
  const [comment,setComment]=useState("");
  const [editStatus,setEditStatus]=useState(ticket.status);
  const [editAssignee,setEditAssignee]=useState(ticket.assigneeId);
  const [showClose,setShowClose]=useState(false);
  const assignee=STAFF_BASE.find(s=>s.id===ticket.assigneeId);
  const cat=CATEGORIES.find(c=>c.id===ticket.category);

  const updateTicket = async (changes, auditAction, remark = "") => {
    let closedByStatusChange = false;

    setTickets(ts=>ts.map(t=>{
      if(t.id!==ticketId) return t;
      const actor=isAdmin?"Admin":staffName||"User";
      const tl=[...(t.timeline||[]),{action:auditAction,remark,at:Date.now(),by:actor}];
      const oldStatus=t.status;
      const closingNow=changes.status==="Closed"&&oldStatus!=="Closed";
      const updated={...t,...changes,updatedAt:Date.now(),timeline:tl};
      if(closingNow){
        updated.closedAt=Date.now();
        updated.closingRemarks=remark||"Closed from ticket controls.";
        updated.resolutionTime=updated.closedAt-t.createdAt;
        updated.feedbackSubmitted=false;
        closedByStatusChange=true;
        console.log("Ticket closed without email:", updated);
      }
      if(changes.assigneeId&&Number(changes.assigneeId)!==Number(t.assigneeId)){
        emailTicketAssigned(updated,STAFF_BASE.find(s=>s.id===Number(changes.assigneeId)),STAFF_BASE.find(s=>s.id===t.assigneeId),actor,remark);
      }
      return updated;
    }));

    toast(closedByStatusChange ? "Ticket closed" : auditAction,"success");
  };
  const addComment=()=>{
    if(!comment.trim()) return;
    setTickets(ts=>ts.map(t=>{
      if(t.id!==ticketId) return t;
      const comments=[...(t.comments||[]),{text:comment,at:Date.now(),by:isAdmin?"Admin":staffName||"User"}];
      const tl=[...(t.timeline||[]),{action:"Commented",remark:comment.slice(0,60),at:Date.now(),by:isAdmin?"Admin":staffName||"User"}];
      return {...t,comments,timeline:tl,updatedAt:Date.now()};
    }));
    setComment("");
    toast("Comment added","success");
  };

  const handleCloseTicket=async(remarks)=>{
    const closedAt=Date.now();
    const closedBy = staffName || (isAdmin ? "Admin" : "User");
    let closedTicket = null;
    setTickets(ts=>ts.map(t=>{
      if(t.id!==ticketId) return t;
      if(t.status==="Closed") return t;
      const updated={...t,status:"Closed",closedAt,closingRemarks:remarks,resolutionTime:closedAt-t.createdAt,updatedAt:closedAt,feedbackSubmitted:false,
        timeline:[...(t.timeline||[]),{action:"Closed",remark:remarks,at:closedAt,by:closedBy}]};
      closedTicket = updated;
      console.log("Ticket closed without email:", updated);
      return updated;
    }));

    setShowClose(false);
    if(closedTicket){
      toast("Ticket closed", "success");
    }
  };
  const canClose=(isAdmin||(isStaff&&ticket.assigneeId===staffId))&&ticket.status!=="Closed";
  const currentTicket=tickets.find(t=>t.id===ticketId)||ticket;

  if(showClose) return (
    <div>
      <button onClick={()=>setShowClose(false)} style={{background:"none",border:"none",color:"rgba(226,232,240,0.5)",fontSize:13,marginBottom:16,display:"flex",alignItems:"center",gap:6}}>← Back to ticket</button>
      <CloseTicketDialog ticket={currentTicket} onClose={()=>setShowClose(false)} onConfirm={handleCloseTicket}/>
    </div>
  );

  return (
    <div style={{display:"flex",flexDirection:"column",gap:20}}>
      {/* Header */}
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",flexWrap:"wrap",gap:12}}>
        <div style={{display:"flex",gap:12,alignItems:"center"}}>
          <div style={{width:44,height:44,borderRadius:12,background:cat?.bg,display:"flex",alignItems:"center",justifyContent:"center",fontSize:22}}>{cat?.icon}</div>
          <div>
            <div style={{fontSize:18,fontWeight:700,fontFamily:"Syne",color:"#e2e8f0"}}>{currentTicket.id}</div>
            <div style={{fontSize:13,color:"rgba(226,232,240,0.5)"}}>{cat?.label} · {fmtDate(currentTicket.createdAt)}</div>
          </div>
        </div>
        <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
          <StatusBadge status={currentTicket.status}/>
          <PriorityBadge p={currentTicket.priority}/>
          {canClose&&<button className="success-btn" onClick={()=>setShowClose(true)} style={{padding:"7px 16px",fontSize:13}}>✅ Close Ticket</button>}
        </div>
      </div>

      {/* SLA */}
      <div className="glass" style={{padding:"16px 18px"}}>
        <div style={{fontSize:12,color:"rgba(226,232,240,0.5)",marginBottom:8,letterSpacing:".5px"}}>SLA TRACKING</div>
        <TimerBadge ticket={currentTicket}/>
        <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:12,marginTop:14}}>
          {[["Opened",fmtDate(currentTicket.createdAt)],["SLA Limit",`${SLA_HOURS[currentTicket.priority]}h`],
            currentTicket.closedAt?["Closed",fmtDate(currentTicket.closedAt)]:["Last Update",timeAgo(currentTicket.updatedAt)]].map(([l,v])=>(
            <div key={l} style={{textAlign:"center"}}><div style={{fontSize:11,color:"rgba(226,232,240,0.4)"}}>{l}</div><div style={{fontSize:12,fontWeight:600,color:"#e2e8f0",marginTop:2}}>{v}</div></div>
          ))}
        </div>
        {currentTicket.closedAt&&(
          <div style={{marginTop:12,textAlign:"center",padding:"10px",background:"rgba(16,185,129,0.1)",borderRadius:8,border:"1px solid rgba(16,185,129,0.2)"}}>
            <div style={{fontSize:13,fontWeight:600,color:"#34d399"}}>✅ Resolved in {formatDuration(currentTicket.closedAt-currentTicket.createdAt)}</div>
            {currentTicket.closingRemarks&&<div style={{fontSize:12,color:"rgba(226,232,240,0.5)",marginTop:4}}>Remarks: {currentTicket.closingRemarks}</div>}
          </div>
        )}
      </div>

      {/* Info */}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
        {[["Submitted by",currentTicket.name],["Email",currentTicket.email],["Department",currentTicket.dept],["Mobile",currentTicket.mobile||"—"],["Location",currentTicket.location||"—"],["Priority",currentTicket.priority]].map(([l,v])=>(
          <div key={l} className="glass" style={{padding:"12px 14px"}}><div style={{fontSize:11,color:"rgba(226,232,240,0.4)"}}>{l}</div><div style={{fontSize:13,fontWeight:500,color:"#e2e8f0",marginTop:3}}>{v}</div></div>
        ))}
      </div>

      {/* Description */}
      <div className="glass" style={{padding:"16px 18px"}}>
        <div style={{fontSize:12,color:"rgba(226,232,240,0.5)",marginBottom:8}}>DESCRIPTION</div>
        <p style={{fontSize:14,lineHeight:1.7,color:"rgba(226,232,240,0.8)"}}>{currentTicket.description}</p>
      </div>

      {/* Assigned To */}
      {assignee&&<div className="glass" style={{padding:"16px 18px",display:"flex",gap:14,alignItems:"center"}}>
        <StaffAvatar staff={assignee} profiles={staffProfiles} statuses={staffStatuses} size={44} showStatus />
        <div><div style={{fontSize:11,color:"rgba(226,232,240,0.4)"}}>ASSIGNED TO</div><div style={{fontSize:15,fontWeight:600,color:"#e2e8f0"}}>{assignee.name}</div><div style={{fontSize:12,color:"rgba(226,232,240,0.5)"}}>{assignee.role} · {assignee.email}</div></div>
      </div>}

      {/* Admin Controls */}
{(isAdmin||(isStaff&&STAFF_BASE.find(s=>s.id===staffId)?.permissions.includes("assign")))&&(

  <div
    className="glass"
    style={{padding:"16px 18px"}}
  >

    <div
      style={{
        fontSize:12,
        color:"rgba(226,232,240,0.5)",
        marginBottom:16
      }}
    >
      CONTROLS
    </div>

    <div
      style={{
        display:"grid",
        gridTemplateColumns:"1fr 1fr",
        gap:12
      }}
    >

      <div>

        <label
          style={{
            fontSize:12,
            color:"rgba(226,232,240,0.5)",
            marginBottom:6,
            display:"block"
          }}
        >
          Status
        </label>

        <select
          value={editStatus}
          onChange={e=>setEditStatus(e.target.value)}
        >
          {STATUSES.map(s=>
            <option key={s}>
              {s}
            </option>
          )}
        </select>

      </div>

      <div>

        <label
          style={{
            fontSize:12,
            color:"rgba(226,232,240,0.5)",
            marginBottom:6,
            display:"block"
          }}
        >
          Assign To
        </label>

        <select
          value={editAssignee}
          onChange={e=>setEditAssignee(Number(e.target.value))}
        >
          {STAFF_BASE.map(s=>
            <option
              key={s.id}
              value={s.id}
            >
              {s.name} ({s.role}) - {getStaffStatus(s.id,staffStatuses)}
            </option>
          )}
        </select>

      </div>

    </div>

    <button
      className="glow-btn"

      style={{
        marginTop:12,
        width:"100%",
        fontSize:14
      }}

      onClick={() => {

        const oldStatus =
          currentTicket.status;

        const oldAssignee =
          currentTicket.assigneeId;

        const changes = {
          status:editStatus,
          assigneeId:Number(editAssignee)
        };

        const actions = [];

        if(editStatus !== oldStatus){

          actions.push(
            `Status changed: ${oldStatus} → ${editStatus}`
          );

        }

        if(Number(editAssignee)!==oldAssignee){

          actions.push(
            `Reassigned to ${
              STAFF_BASE.find(
                s=>s.id===Number(editAssignee)
              )?.name
            }`
          );

        }

        updateTicket(
          changes,
          actions.join("; ") || "Updated"
        );

      }}
    >
      💾 Save Changes
    </button>

  </div>

)}
      {/* Audit Timeline */}
      <div className="glass" style={{padding:"16px 18px"}}>
        <div style={{fontSize:12,color:"rgba(226,232,240,0.5)",marginBottom:14,letterSpacing:".5px"}}>AUDIT TRAIL</div>
        <AuditTimeline timeline={currentTicket.timeline}/>
      </div>

      {/* Comments */}
      <div className="glass" style={{padding:"16px 18px"}}>
        <div style={{fontSize:12,color:"rgba(226,232,240,0.5)",marginBottom:12}}>COMMENTS ({(currentTicket.comments||[]).length})</div>
        {(currentTicket.comments||[]).map((c,i)=>(
          <div key={i} style={{background:"rgba(255,255,255,0.04)",borderRadius:10,padding:"10px 14px",marginBottom:8}}>
            <div style={{fontSize:13,color:"#e2e8f0"}}>{c.text}</div>
            <div style={{fontSize:11,color:"rgba(226,232,240,0.4)",marginTop:4}}>{c.by} · {timeAgo(c.at)}</div>
          </div>
        ))}
        <div style={{display:"flex",gap:10,marginTop:12}}>
          <input value={comment} onChange={e=>setComment(e.target.value)} placeholder="Add a comment..." onKeyDown={e=>e.key==="Enter"&&!e.shiftKey&&addComment()}/>
          <button className="glow-btn" style={{padding:"10px 18px",fontSize:13,whiteSpace:"nowrap"}} onClick={addComment}>Send</button>
        </div>
      </div>
    </div>
  );
}

// ── TICKET FORM ───────────────────────────────────────────────────────────
function TicketForm({userEmail,initialCategory,onSubmit,onCancel,toast}) {
  const [form,setForm]=useState({name:"",email:userEmail||"",dept:"",mobile:"",category:initialCategory||"",priority:"Medium",description:"",location:""});
  const [loading,setLoading]=useState(false);
  const set=(k,v)=>setForm(f=>({...f,[k]:v}));

  const submit=async()=>{
    if(!form.name||!form.email||!form.dept||!form.category||!form.description){toast("Fill all required fields","error");return;}
    if(!/\S+@\S+\.\S+/.test(form.email)){toast("Invalid email format","error");return;}
    setLoading(true);
    await new Promise(r=>setTimeout(r,700));
    try {
      await Promise.resolve(onSubmit({...form}));
    } finally {
      setLoading(false);
    }
  };

  return (
    <div style={{display:"flex",flexDirection:"column",gap:16}}>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
        {[["name","Full Name *","text"],["email","Email *","email"],["mobile","Mobile Number","tel"],["location","Location/Room No","text"]].map(([k,label,type])=>(
          <div key={k}><label style={{fontSize:12,color:"rgba(226,232,240,0.5)",marginBottom:6,display:"block"}}>{label}</label>
          <input type={type} placeholder={label} value={form[k]} onChange={e=>set(k,e.target.value)}/></div>
        ))}
        <div><label style={{fontSize:12,color:"rgba(226,232,240,0.5)",marginBottom:6,display:"block"}}>Department *</label>
          <select value={form.dept} onChange={e=>set("dept",e.target.value)}><option value="">Select Department</option>{DEPTS.map(d=><option key={d}>{d}</option>)}</select></div>
        <div><label style={{fontSize:12,color:"rgba(226,232,240,0.5)",marginBottom:6,display:"block"}}>Category *</label>
          <select value={form.category} onChange={e=>set("category",e.target.value)}><option value="">Select Category</option>{CATEGORIES.map(c=><option key={c.id} value={c.id}>{c.icon} {c.label}</option>)}</select></div>
        <div><label style={{fontSize:12,color:"rgba(226,232,240,0.5)",marginBottom:6,display:"block"}}>Priority</label>
          <select value={form.priority} onChange={e=>set("priority",e.target.value)}>{PRIORITIES.map(p=><option key={p}>{p}</option>)}</select></div>
      </div>
      <div><label style={{fontSize:12,color:"rgba(226,232,240,0.5)",marginBottom:6,display:"block"}}>Description *</label>
        <textarea rows={4} placeholder="Describe your issue in detail..." value={form.description} onChange={e=>set("description",e.target.value)} style={{resize:"vertical"}}/></div>
      <div style={{display:"flex",gap:10,justifyContent:"flex-end",flexWrap:"wrap"}}>
        <button onClick={onCancel} style={{background:"rgba(255,255,255,0.06)",border:"1px solid rgba(255,255,255,0.1)",color:"#e2e8f0",padding:"10px 22px",borderRadius:10,fontSize:14}}>Cancel</button>
        <button className="glow-btn" onClick={submit} disabled={loading}>{loading?"⏳ Submitting...":"🚀 Submit Ticket"}</button>
      </div>
    </div>
  );
}

// ── TICKET CARD ───────────────────────────────────────────────────────────
function TicketCard({ticket,onView}) {
  const cat=CATEGORIES.find(c=>c.id===ticket.category);
  return (
    <div className="glass2" style={{padding:"16px 18px",cursor:"pointer",transition:"all .2s"}}
      onClick={()=>onView(ticket.id)}
      onMouseEnter={e=>{e.currentTarget.style.borderColor="rgba(99,102,241,0.4)";e.currentTarget.style.background="rgba(255,255,255,0.08)";}}
      onMouseLeave={e=>{e.currentTarget.style.borderColor="rgba(255,255,255,0.1)";e.currentTarget.style.background="rgba(255,255,255,0.06)";}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:12,marginBottom:10}}>
        <div style={{display:"flex",gap:10,alignItems:"center"}}>
          <div style={{width:36,height:36,borderRadius:10,background:cat?.bg,display:"flex",alignItems:"center",justifyContent:"center",fontSize:18,flexShrink:0}}>{cat?.icon}</div>
          <div><div style={{fontSize:13,fontWeight:600,color:"#e2e8f0"}}>{ticket.id}</div><div style={{fontSize:12,color:"rgba(226,232,240,0.5)"}}>{ticket.name}</div></div>
        </div>
        <StatusBadge status={ticket.status}/>
      </div>
      <div style={{fontSize:13,color:"rgba(226,232,240,0.7)",marginBottom:8,display:"-webkit-box",WebkitLineClamp:2,WebkitBoxOrient:"vertical",overflow:"hidden"}}>{ticket.description}</div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:8,marginBottom:10}}>
        <PriorityBadge p={ticket.priority}/>
        <span style={{fontSize:11,color:"rgba(226,232,240,0.35)"}}>{timeAgo(ticket.createdAt)}</span>
      </div>
      <TimerBadge ticket={ticket}/>
    </div>
  );
}

// ── CATEGORY GRID ─────────────────────────────────────────────────────────
function CategoryGrid({onSelect}) {
  const [hover,setHover]=useState(null);
  return (
    <div>
      <h2 style={{fontFamily:"Syne",fontSize:22,fontWeight:700,color:"#e2e8f0",marginBottom:6}}>Raise IT Support Ticket</h2>
      <p style={{fontSize:14,color:"rgba(226,232,240,0.5)",marginBottom:24}}>Select the issue category to get started</p>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(145px,1fr))",gap:14}}>
        {CATEGORIES.map(cat=>(
          <button key={cat.id} onClick={()=>onSelect(cat.id)} onMouseEnter={()=>setHover(cat.id)} onMouseLeave={()=>setHover(null)} style={{
            background:hover===cat.id?cat.bg:"rgba(255,255,255,0.04)",border:`1px solid ${hover===cat.id?cat.color+"60":"rgba(255,255,255,0.08)"}`,
            borderRadius:14,padding:"18px 14px",textAlign:"center",cursor:"pointer",transition:"all .2s",
            transform:hover===cat.id?"translateY(-3px)":"none",boxShadow:hover===cat.id?`0 8px 24px ${cat.color}30`:"none",
          }}>
            <div style={{fontSize:28,marginBottom:8}}>{cat.icon}</div>
            <div style={{fontSize:12,fontWeight:600,color:hover===cat.id?cat.color:"rgba(226,232,240,0.7)",lineHeight:1.3}}>{cat.label}</div>
          </button>
        ))}
      </div>
    </div>
  );
}

// ── EXPORT PANEL ──────────────────────────────────────────────────────────
function ExportPanel({tickets,toast}) {
  const [loading,setLoading]=useState("");
  const [dateFrom,setDateFrom]=useState("");
  const [dateTo,setDateTo]=useState("");
  const [format,setFormat]=useState("excel");

  const filtered=tickets.filter(t=>{
    if(dateFrom&&t.createdAt<new Date(dateFrom).getTime()) return false;
    if(dateTo&&t.createdAt>new Date(dateTo).getTime()+86399999) return false;
    return true;
  });

  const requireDateRange=()=>{
    if(!dateFrom||!dateTo){toast("Please select date range","error");return false;}
    return true;
  };

  const doExport=async(type)=>{
    if(!requireDateRange()) return;
    setLoading(type);
    await new Promise(r=>setTimeout(r,900));
    const now=new Date().toISOString().slice(0,10);
    let data=filtered;
    if(type==="open") data=filtered.filter(t=>t.status==="Open"||t.status==="Assigned"||t.status==="In Progress");
    if(type==="closed") data=filtered.filter(t=>t.status==="Closed"||t.status==="Resolved");
    const rows = data.map(cleanTicketRow);

    if(format==="excel") {
      downloadExcel(rows, `tickets_${type}_${now}.xlsx`);
    }
    if(format==="pdf") {
      downloadPDF(rows, `tickets_${type}_${now}.pdf`, { sourceTickets:data, dateFrom, dateTo, type });
    }

    toast(`${type} report exported (${data.length} tickets) ✅`,"success");
    setLoading("");
  };

  const exportSlaReport=()=>{
    if(!requireDateRange()) return;
    const now=new Date().toISOString().slice(0,10);
    if(format==="excel") downloadExcel(slaData, `sla_report_${now}.xlsx`);
    if(format==="pdf") downloadGenericReportPDF("IT Helpdesk SLA Report", slaData, `sla_report_${now}.pdf`, {dateFrom,dateTo});
    toast("SLA report exported ✅","success");
  };

  const exportStaffReport=()=>{
    if(!requireDateRange()) return;
    const now=new Date().toISOString().slice(0,10);
    if(format==="excel") downloadExcel(staffData, `staff_performance_${now}.xlsx`);
    if(format==="pdf") downloadStaffPerformancePDF(staffData, `staff_performance_${now}.pdf`, {dateFrom,dateTo});
    toast("Staff performance exported ✅","success");
  };

  const slaData=filtered.map(t=>({
    id:t.id,priority:t.priority,slaHours:SLA_HOURS[t.priority],
    status:t.status,elapsed:t.closedAt?formatDuration(t.closedAt-t.createdAt):"Active",met:t.closedAt?(t.closedAt-t.createdAt<SLA_HOURS[t.priority]*3600000?"✅ Met":"❌ Breached"):"—"
  }));

  const staffData=STAFF_BASE.map(s=>{
    const assigned=filtered.filter(t=>t.assigneeId===s.id).length;
    const resolved=filtered.filter(t=>t.assigneeId===s.id&&(t.status==="Resolved"||t.status==="Closed")).length;
    const pending=Math.max(0,assigned-resolved);
    const resolutionRate=assigned?Math.round((resolved/assigned)*100):0;
    return {name:s.name,role:s.role,email:s.email,assigned,resolved,pending,resolutionRate};
  });

  const exports=[
    {id:"all",label:"All Tickets",icon:"🎫",color:"#6366f1",desc:`${filtered.length} tickets`},
    {id:"open",label:"Open Tickets",icon:"🔵",color:"#0ea5e9",desc:`${filtered.filter(t=>!["Resolved","Closed"].includes(t.status)).length} active`},
    {id:"closed",label:"Closed Tickets",icon:"✅",color:"#10b981",desc:`${filtered.filter(t=>t.status==="Closed").length} resolved`},
    {id:"sla",label:"SLA Report",icon:"⏱",color:"#f59e0b",desc:"SLA performance"},
    {id:"staff",label:"Staff Performance",icon:"👥",color:"#8b5cf6",desc:"Per-agent stats"},
  ];

  return (
    <div style={{display:"flex",flexDirection:"column",gap:20}}>
      <h2 style={{fontFamily:"Syne",fontSize:22,fontWeight:700,color:"#e2e8f0"}}>Export Reports</h2>
      {/* Filters */}
      <div className="glass" style={{padding:"18px 20px"}}>
        <div style={{fontSize:13,fontWeight:600,color:"rgba(226,232,240,0.6)",marginBottom:14,letterSpacing:".5px"}}>FILTERS & FORMAT</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:14,flexWrap:"wrap"}}>
          <div><label style={{fontSize:12,color:"rgba(226,232,240,0.5)",marginBottom:6,display:"block"}}>From Date</label><input type="date" value={dateFrom} onChange={e=>setDateFrom(e.target.value)}/></div>
          <div><label style={{fontSize:12,color:"rgba(226,232,240,0.5)",marginBottom:6,display:"block"}}>To Date</label><input type="date" value={dateTo} onChange={e=>setDateTo(e.target.value)}/></div>
          <div><label style={{fontSize:12,color:"rgba(226,232,240,0.5)",marginBottom:6,display:"block"}}>Format</label>
            <select value={format} onChange={e=>setFormat(e.target.value)}><option value="excel">Excel (.xlsx)</option><option value="pdf">PDF (.pdf)</option></select>
          </div>
        </div>
        <div style={{marginTop:10,fontSize:12,color:"rgba(226,232,240,0.4)"}}>Showing {filtered.length} of {tickets.length} tickets</div>
      </div>

      {/* Export cards */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(200px,1fr))",gap:14}}>
        {exports.map(ex=>(
          <div key={ex.id} className="glass2" style={{padding:"18px 16px",textAlign:"center"}}>
            <div style={{fontSize:32,marginBottom:8}}>{ex.icon}</div>
            <div style={{fontSize:14,fontWeight:600,color:"#e2e8f0",marginBottom:4}}>{ex.label}</div>
            <div style={{fontSize:12,color:"rgba(226,232,240,0.4)",marginBottom:14}}>{ex.desc}</div>
            <button onClick={()=>{
              if(ex.id==="sla") { exportSlaReport(); return; }
              if(ex.id==="staff") { exportStaffReport(); return; }
              doExport(ex.id);
            }} disabled={loading===ex.id} style={{
              background:`${ex.color}20`,border:`1px solid ${ex.color}40`,color:ex.color,
              padding:"9px 18px",borderRadius:8,fontSize:13,fontWeight:600,width:"100%",display:"flex",alignItems:"center",justifyContent:"center",gap:8,
            }}>
              {loading===ex.id?<><span className="spin" style={{display:"inline-block"}}>⟳</span> Exporting...</>:<>⬇ Download</>}
            </button>
          </div>
        ))}
      </div>
    </div>
  );
}

// ── EMAIL LOG VIEW ─────────────────────────────────────────────────────────
function EmailLog() {
  const log=DB.get("emaillog",[]);
  return (
    <div style={{display:"flex",flexDirection:"column",gap:16}}>
      <h2 style={{fontFamily:"Syne",fontSize:22,fontWeight:700,color:"#e2e8f0"}}>Email Notifications Log</h2>
      <p style={{fontSize:13,color:"rgba(226,232,240,0.5)"}}>Simulated email notifications (no real SMTP — shows what would be sent)</p>
      {log.length===0&&<div style={{textAlign:"center",padding:"60px 0",color:"rgba(226,232,240,0.3)"}}>📧 No emails sent yet</div>}
      {log.map(e=>(
        <div key={e.id} className="glass2" style={{padding:"16px 18px"}}>
          <div style={{display:"flex",justifyContent:"space-between",flexWrap:"wrap",gap:8,marginBottom:8}}>
            <div style={{fontSize:14,fontWeight:600,color:"#818cf8"}}>📧 {e.subject}</div>
            <span style={{fontSize:12,color:"rgba(226,232,240,0.4)"}}>{fmtDate(e.sentAt)}</span>
          </div>
          <div style={{fontSize:12,color:"rgba(226,232,240,0.5)",marginBottom:10}}>To: {e.to}</div>
          <pre style={{fontSize:12,color:"rgba(226,232,240,0.7)",background:"rgba(0,0,0,0.2)",padding:"12px 14px",borderRadius:8,whiteSpace:"pre-wrap",fontFamily:"monospace",lineHeight:1.6}}>{e.body}</pre>
        </div>
      ))}
    </div>
  );
}

// ── ANALYTICS ─────────────────────────────────────────────────────────────
function Analytics({tickets}) {
  const total=tickets.length;
  const byStatus=STATUSES.reduce((a,s)=>{a[s]=tickets.filter(t=>t.status===s).length;return a;},{});
  const byPriority=PRIORITIES.reduce((a,p)=>{a[p]=tickets.filter(t=>t.priority===p).length;return a;},{});
  const byCat=CATEGORIES.map(c=>({...c,count:tickets.filter(t=>t.category===c.id).length})).filter(c=>c.count>0).sort((a,b)=>b.count-a.count).slice(0,8);
  const resolved=tickets.filter(t=>t.closedAt&&t.createdAt);
  const avgResMs=resolved.length?resolved.reduce((a,t)=>a+(t.closedAt-t.createdAt),0)/resolved.length:0;
  const slaBreaches=tickets.filter(t=>{const ms=SLA_HOURS[t.priority]*3600000;return(t.closedAt||Date.now())-t.createdAt>ms;}).length;

  return (
    <div style={{display:"flex",flexDirection:"column",gap:24}}>
      <h2 style={{fontFamily:"Syne",fontSize:22,fontWeight:700,color:"#e2e8f0"}}>Analytics Dashboard</h2>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(160px,1fr))",gap:14}}>
        <StatCard label="Total Tickets" value={total} icon="🎫" color="#818cf8"/>
        <StatCard label="Open" value={byStatus.Open} icon="🔵" color="#60a5fa"/>
        <StatCard label="In Progress" value={byStatus["In Progress"]} icon="🟡" color="#fbbf24"/>
        <StatCard label="Resolved" value={(byStatus.Resolved||0)+(byStatus.Closed||0)} icon="🟢" color="#34d399"/>
        <StatCard label="SLA Breaches" value={slaBreaches} icon="⚠️" color="#f87171"/>
        <StatCard label="Avg Resolution" value={avgResMs?formatDuration(avgResMs):"—"} icon="⚡" color="#f97316"/>
      </div>
      <div className="glass" style={{padding:"20px 22px"}}>
        <div style={{fontSize:13,fontWeight:600,color:"rgba(226,232,240,0.6)",marginBottom:16,letterSpacing:".5px"}}>STATUS BREAKDOWN</div>
        {STATUSES.map(s=>{const cnt=byStatus[s];const pct=total?Math.round((cnt/total)*100):0;
          const col={Open:"#818cf8",Assigned:"#38bdf8","In Progress":"#fbbf24",Resolved:"#34d399",Closed:"#6b7280"}[s];
          return(<div key={s} style={{display:"flex",alignItems:"center",gap:12,marginBottom:10}}>
            <div style={{width:90,fontSize:13,color:"rgba(226,232,240,0.7)"}}>{s}</div>
            <div style={{flex:1,height:8,background:"rgba(255,255,255,0.06)",borderRadius:4}}><div style={{width:`${pct}%`,height:"100%",background:col,borderRadius:4,transition:"width .6s"}}/></div>
            <div style={{width:36,fontSize:13,fontWeight:600,color:col,textAlign:"right"}}>{cnt}</div>
          </div>);
        })}
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
        <div className="glass" style={{padding:"20px 22px"}}>
          <div style={{fontSize:13,fontWeight:600,color:"rgba(226,232,240,0.6)",marginBottom:16,letterSpacing:".5px"}}>PRIORITY SPLIT</div>
          {PRIORITIES.map(p=>{const c={Low:"#94a3b8",Medium:"#fbbf24",High:"#f97316",Critical:"#ef4444"}[p];return(
            <div key={p} style={{textAlign:"center",background:`${c}15`,border:`1px solid ${c}30`,borderRadius:10,padding:"12px 8px",marginBottom:10}}>
              <div style={{fontSize:22,fontFamily:"Syne",fontWeight:800,color:c}}>{byPriority[p]}</div>
              <div style={{fontSize:12,color:"rgba(226,232,240,0.5)",marginTop:2}}>{p}</div>
            </div>
          );})}
        </div>
        <div className="glass" style={{padding:"20px 22px"}}>
          <div style={{fontSize:13,fontWeight:600,color:"rgba(226,232,240,0.6)",marginBottom:16,letterSpacing:".5px"}}>TOP CATEGORIES</div>
          {byCat.slice(0,6).map(c=>(
            <div key={c.id} style={{display:"flex",alignItems:"center",gap:10,marginBottom:10}}>
              <span style={{fontSize:16}}>{c.icon}</span>
              <div style={{flex:1,fontSize:12,color:"rgba(226,232,240,0.7)"}}>{c.label}</div>
              <div style={{width:60,height:5,background:"rgba(255,255,255,0.06)",borderRadius:3}}><div style={{width:`${(c.count/(byCat[0]?.count||1))*100}%`,height:"100%",background:c.color,borderRadius:3}}/></div>
              <div style={{width:20,fontSize:12,fontWeight:600,color:c.color,textAlign:"right"}}>{c.count}</div>
            </div>
          ))}
          {byCat.length===0&&<div style={{color:"rgba(226,232,240,0.3)",fontSize:13}}>No data yet</div>}
        </div>
      </div>
      <div className="glass" style={{padding:"20px 22px"}}>
        <div style={{fontSize:13,fontWeight:600,color:"rgba(226,232,240,0.6)",marginBottom:16,letterSpacing:".5px"}}>STAFF PERFORMANCE</div>
        {STAFF_BASE.map(s=>{
          const asgn=tickets.filter(t=>t.assigneeId===s.id).length;
          const res=tickets.filter(t=>t.assigneeId===s.id&&(t.status==="Resolved"||t.status==="Closed")).length;
          return(
            <div key={s.id} style={{display:"flex",alignItems:"center",gap:14,marginBottom:14}}>
              <div style={{width:38,height:38,borderRadius:"50%",background:s.color+"33",border:`2px solid ${s.color}`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:12,fontWeight:700,color:s.color,flexShrink:0}}>{s.avatar}</div>
              <div style={{flex:1}}>
                <div style={{fontSize:14,fontWeight:600,color:"#e2e8f0"}}>{s.name}</div>
                <div style={{fontSize:12,color:"rgba(226,232,240,0.5)"}}>{s.role} · {asgn} assigned, {res} resolved</div>
                <div style={{height:4,background:"rgba(255,255,255,0.06)",borderRadius:2,marginTop:6}}><div style={{width:asgn?`${(res/asgn)*100}%`:"0%",height:"100%",background:s.color,borderRadius:2,transition:"width .6s"}}/></div>
              </div>
              <div style={{textAlign:"right"}}><div style={{fontSize:14,fontWeight:600,color:"#34d399"}}>{asgn?Math.round((res/asgn)*100):0}%</div><div style={{fontSize:11,color:"rgba(226,232,240,0.4)"}}>rate</div></div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

// ── IT FEEDBACK ───────────────────────────────────────────────────────────
function MiniBar({label,value,total,color="#6366f1"}) {
  const pct = total ? Math.round((value / total) * 100) : 0;
  return (
    <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:10}}>
      <div style={{width:140,fontSize:12,color:"rgba(226,232,240,0.65)",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{label}</div>
      <div style={{flex:1,height:8,background:"rgba(255,255,255,0.07)",borderRadius:8,overflow:"hidden"}}>
        <div style={{width:`${pct}%`,height:"100%",background:color,borderRadius:8,transition:"width .4s"}} />
      </div>
      <div style={{width:42,textAlign:"right",fontSize:12,fontWeight:700,color}}>{value}</div>
    </div>
  );
}

function FeedbackForm({userEmail,onSubmit,toast,ticket=null}) {
  const empty = {ticketId:ticket?.id||"",name:ticket?.name||"",email:ticket?.email||userEmail||"",dept:ticket?.dept||"",category:ticket?"Ticket Resolution":"",rating:0,satisfaction:"",message:"",suggestions:"",recommend:"Yes"};
  const [form,setForm]=useState(empty);
  const [hoverRating,setHoverRating]=useState(0);
  const [reactionBursts,setReactionBursts]=useState([]);
  const reactionTimers=useRef([]);
  const set=(k,v)=>setForm(f=>({...f,[k]:v}));

  useEffect(()=>()=>{
    reactionTimers.current.forEach(clearTimeout);
  },[]);

  const triggerSatisfactionReaction=(s,e)=>{
    if(typeof window === "undefined") return;
    const rect=e.currentTarget.getBoundingClientRect();
    const originX=rect.left+rect.width/2;
    const originY=rect.top+rect.height/2;
    const count=12+Math.floor(Math.random()*9);
    const particles=Array.from({length:count},(_,i)=>({
      id:`${Date.now()}-${s.id}-${i}-${Math.random()}`,
      emoji:s.icon,
      x:originX+(Math.random()*28-14),
      y:originY+(Math.random()*16-8),
      dx:Math.random()*220-110,
      dy:150+Math.random()*170,
      rotate:Math.random()*110-55,
      scale:.78+Math.random()*.7,
      duration:1.15+Math.random()*.65,
      delay:Math.random()*.14,
      sparkle:Math.random()>.45,
      color:s.color,
    }));
    setReactionBursts(prev=>[...prev,...particles]);
    const timer=setTimeout(()=>{
      setReactionBursts(prev=>prev.filter(p=>!particles.some(n=>n.id===p.id)));
    },2100);
    reactionTimers.current.push(timer);
  };

  useEffect(()=>setForm(f=>({...f,ticketId:ticket?.id||"",name:ticket?.name||f.name,email:ticket?.email||userEmail||f.email,dept:ticket?.dept||f.dept,category:ticket?"Ticket Resolution":f.category})),[userEmail,ticket]);

  const submit=async()=>{
    if(!form.name.trim()||!form.email.trim()||!form.dept||!form.category||!form.rating||!form.satisfaction||!form.message.trim()){
      toast("Please complete all required feedback fields","error");
      return;
    }
    if(!isInstitutionEmail(form.email)){
      toast("Only @jaipuria.ac.in email ID is allowed","error");
      return;
    }
    const entry={...form,ticketId:form.ticketId||"",name:form.name.trim(),email:form.email.trim(),message:form.message.trim(),suggestions:form.suggestions.trim(),id:genFeedbackId(),createdAt:Date.now(),reviewed:false};
    try {
      await Promise.resolve(onSubmit(entry));
      setForm({...empty,ticketId:ticket?.id||"",name:ticket?.name||"",email:ticket?.email||userEmail||"",dept:ticket?.dept||"",category:ticket?"Ticket Resolution":""});
      setHoverRating(0);
    } catch (error) {
      console.error("Feedback submit failed:", error);
    }
  };

  return (
    <div style={{display:"flex",flexDirection:"column",gap:22,maxWidth:1040}}>
      <div className="glass" style={{padding:"24px",background:"linear-gradient(135deg,rgba(99,102,241,0.16),rgba(20,184,166,0.08)),rgba(255,255,255,0.04)"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",gap:16,flexWrap:"wrap"}}>
          <div>
            <div style={{fontSize:13,color:"#38bdf8",fontWeight:700,letterSpacing:".5px",textTransform:"uppercase"}}>IT Feedback</div>
            <h2 style={{fontFamily:"Syne",fontSize:26,fontWeight:800,color:"#e2e8f0",marginTop:4}}>Help us improve IT support</h2>
            <p style={{fontSize:14,color:"rgba(226,232,240,0.58)",marginTop:8,maxWidth:620,lineHeight:1.6}}>Share how your IT service experience felt. Your feedback helps the team improve ticket resolution, support quality, and campus technology services.</p>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:8,minWidth:220}}>
            {["Fast Help","Clear Updates","Better Service"].map((x,i)=><div key={x} style={{padding:"12px",borderRadius:12,background:["rgba(99,102,241,0.16)","rgba(16,185,129,0.14)","rgba(245,158,11,0.14)"][i],border:"1px solid rgba(255,255,255,0.08)",textAlign:"center",fontSize:12,fontWeight:700,color:"#e2e8f0"}}>{x}</div>)}
          </div>
        </div>
      </div>

      <div className="glass" style={{padding:"24px",display:"flex",flexDirection:"column",gap:18}}>
        {ticket&&<div className="glass2" style={{padding:"14px 16px",borderColor:"rgba(99,102,241,0.35)",background:"rgba(99,102,241,0.1)"}}><div style={{fontSize:12,color:"#a5b4fc",fontWeight:800,letterSpacing:".5px",marginBottom:4}}>CLOSED TICKET FEEDBACK</div><div style={{fontSize:14,color:"#e2e8f0"}}>Your ticket <strong>{ticket.id}</strong> has been closed. Please share your experience with the IT support team.</div></div>}
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(220px,1fr))",gap:14}}>
          {ticket&&<div><label style={{fontSize:12,color:"rgba(226,232,240,0.55)",marginBottom:6,display:"block"}}>Ticket ID</label><input value={form.ticketId} readOnly style={{opacity:.75}} /></div>}
          <div><label style={{fontSize:12,color:"rgba(226,232,240,0.55)",marginBottom:6,display:"block"}}>User Name *</label><input value={form.name} onChange={e=>set("name",e.target.value)} placeholder="Your full name" readOnly={!!ticket} /></div>
          <div><label style={{fontSize:12,color:"rgba(226,232,240,0.55)",marginBottom:6,display:"block"}}>User Email *</label><input type="email" value={form.email} onChange={e=>set("email",e.target.value)} placeholder="name@jaipuria.ac.in" readOnly={!!ticket} /></div>
          <div><label style={{fontSize:12,color:"rgba(226,232,240,0.55)",marginBottom:6,display:"block"}}>Department *</label><select value={form.dept} onChange={e=>set("dept",e.target.value)} disabled={!!ticket}><option value="">Select Department</option>{DEPTS.map(d=><option key={d}>{d}</option>)}</select></div>
          <div><label style={{fontSize:12,color:"rgba(226,232,240,0.55)",marginBottom:6,display:"block"}}>Service Category *</label><select value={form.category} onChange={e=>set("category",e.target.value)}><option value="">Select Service</option>{FEEDBACK_CATEGORIES.map(c=><option key={c}>{c}</option>)}</select></div>
        </div>

        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(260px,1fr))",gap:16}}>
          <div className="glass2" style={{padding:"18px"}}>
            <div style={{fontSize:13,fontWeight:700,color:"rgba(226,232,240,0.75)",marginBottom:12}}>Rate IT Service *</div>
            <div style={{display:"flex",gap:6,alignItems:"center"}} onMouseLeave={()=>setHoverRating(0)}>
              {[1,2,3,4,5].map(n=>(
                <button key={n} type="button" onMouseEnter={()=>setHoverRating(n)} onClick={()=>set("rating",n)} style={{background:"transparent",border:"none",fontSize:34,lineHeight:1,color:n <= (hoverRating||form.rating) ? "#fbbf24" : "rgba(255,255,255,0.18)",filter:n <= (hoverRating||form.rating) ? "drop-shadow(0 0 10px rgba(251,191,36,0.35))" : "none",transform:n <= hoverRating ? "translateY(-2px) scale(1.08)" : "none"}}>★</button>
              ))}
              <span style={{marginLeft:8,fontSize:13,color:"rgba(226,232,240,0.5)"}}>{form.rating ? `${form.rating}/5` : "Select rating"}</span>
            </div>
          </div>

          <div className="glass2" style={{padding:"18px",position:"relative",overflow:"hidden"}}>
            <AnimatePresence>
              {reactionBursts.map(p=>(
                <motion.span
                  key={p.id}
                  aria-hidden="true"
                  initial={{x:p.x,y:p.y,opacity:0,scale:.35,rotate:0}}
                  animate={{
                    x:p.x+p.dx,
                    y:p.y-p.dy,
                    opacity:[0,1,.92,0],
                    scale:[.35,p.scale,p.scale*.92],
                    rotate:p.rotate,
                  }}
                  exit={{opacity:0,scale:.2}}
                  transition={{duration:p.duration,delay:p.delay,ease:[.16,1,.3,1]}}
                  style={{
                    position:"fixed",
                    left:0,
                    top:0,
                    zIndex:9998,
                    pointerEvents:"none",
                    fontSize:26,
                    lineHeight:1,
                    filter:"drop-shadow(0 10px 18px rgba(0,0,0,.35))",
                    willChange:"transform, opacity",
                  }}
                >
                  {p.emoji}
                  {p.sparkle&&<span style={{position:"absolute",left:"58%",top:"-28%",fontSize:11,color:p.color,textShadow:`0 0 12px ${p.color}`}}>✦</span>}
                </motion.span>
              ))}
            </AnimatePresence>
            <div style={{fontSize:13,fontWeight:700,color:"rgba(226,232,240,0.75)",marginBottom:12}}>Satisfaction Level *</div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:10}}>
              {SATISFACTION_LEVELS.map(s=>{
                const selected=form.satisfaction===s.id;
                return (
                  <motion.button
                    key={s.id}
                    type="button"
                    aria-pressed={selected}
                    aria-label={`Select ${s.id} satisfaction`}
                    onClick={(e)=>{set("satisfaction",s.id);triggerSatisfactionReaction(s,e);}}
                    whileTap={{scale:.93}}
                    animate={{
                      scale:selected?1.03:1,
                      boxShadow:selected?`0 0 0 1px ${s.color}, 0 0 28px ${s.color}55, inset 0 0 22px ${s.color}22`:"0 0 0 rgba(0,0,0,0)",
                    }}
                    transition={{type:"spring",stiffness:380,damping:24}}
                    style={{
                      padding:"12px",
                      borderRadius:12,
                      border:`1px solid ${selected?s.color:"rgba(255,255,255,0.08)"}`,
                      background:selected?`linear-gradient(135deg,${s.bg},rgba(255,255,255,0.055))`:"rgba(255,255,255,0.04)",
                      color:selected?s.color:"rgba(226,232,240,0.72)",
                      fontWeight:700,
                      display:"flex",
                      alignItems:"center",
                      justifyContent:"center",
                      gap:8,
                      position:"relative",
                      overflow:"hidden",
                      transformOrigin:"center",
                    }}
                  >
                    {selected&&(
                      <motion.span
                        aria-hidden="true"
                        initial={{opacity:.55,scale:.5}}
                        animate={{opacity:0,scale:2.7}}
                        transition={{duration:.75,ease:"easeOut"}}
                        style={{position:"absolute",inset:0,margin:"auto",width:52,height:52,borderRadius:"50%",background:s.color,filter:"blur(18px)"}}
                      />
                    )}
                    <motion.span animate={{scale:selected?[1,1.18,1]:1,rotate:selected?[0,-5,5,0]:0}} transition={{duration:.45}} style={{fontSize:20,position:"relative"}}>{s.icon}</motion.span>
                    <span style={{position:"relative"}}>{s.id}</span>
                  </motion.button>
                );
              })}
            </div>
          </div>
        </div>

        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(280px,1fr))",gap:14}}>
          <div><label style={{fontSize:12,color:"rgba(226,232,240,0.55)",marginBottom:6,display:"block"}}>Feedback Message *</label><textarea rows={5} value={form.message} onChange={e=>set("message",e.target.value)} placeholder="Tell us what worked well or what did not..." style={{resize:"vertical"}} /></div>
          <div><label style={{fontSize:12,color:"rgba(226,232,240,0.55)",marginBottom:6,display:"block"}}>Suggestions for Improvement</label><textarea rows={5} value={form.suggestions} onChange={e=>set("suggestions",e.target.value)} placeholder="Any ideas to improve IT support?" style={{resize:"vertical"}} /></div>
        </div>

        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",gap:14,flexWrap:"wrap"}}>
          <div>
            <div style={{fontSize:13,fontWeight:700,color:"rgba(226,232,240,0.75)",marginBottom:8}}>Would you recommend IT support?</div>
            <div style={{display:"flex",gap:10}}>
              {["Yes","No"].map(v=><button key={v} type="button" onClick={()=>set("recommend",v)} style={{padding:"10px 18px",borderRadius:10,border:`1px solid ${form.recommend===v?(v==="Yes"?"#10b981":"#ef4444"):"rgba(255,255,255,0.1)"}`,background:form.recommend===v?(v==="Yes"?"rgba(16,185,129,0.16)":"rgba(239,68,68,0.16)"):"rgba(255,255,255,0.04)",color:form.recommend===v?(v==="Yes"?"#34d399":"#f87171"):"rgba(226,232,240,0.65)",fontWeight:700}}>{v}</button>)}
            </div>
          </div>
          <button className="glow-btn" onClick={submit} style={{minWidth:190}}>Submit Feedback</button>
        </div>
      </div>
    </div>
  );
}

function AdminFeedbackPage({feedback,setFeedback,toast}) {
  const [dateFrom,setDateFrom]=useState("");
  const [dateTo,setDateTo]=useState("");
  const [format,setFormat]=useState("excel");
  const [ticketFilter,setTicketFilter]=useState("");
  const [deptFilter,setDeptFilter]=useState("All");
  const [ratingFilter,setRatingFilter]=useState("All");
  const [reviewFilter,setReviewFilter]=useState("All");
  const filtered=feedback.filter(f=>{
    if(dateFrom && f.createdAt < new Date(dateFrom).getTime()) return false;
    if(dateTo && f.createdAt > new Date(dateTo).getTime()+86399999) return false;
    if(ticketFilter.trim() && !(f.ticketId||"").toLowerCase().includes(ticketFilter.trim().toLowerCase())) return false;
    if(deptFilter!=="All" && f.dept!==deptFilter) return false;
    if(ratingFilter!=="All" && Number(f.rating)!==Number(ratingFilter)) return false;
    if(reviewFilter==="Reviewed" && !f.reviewed) return false;
    if(reviewFilter==="Unreviewed" && f.reviewed) return false;
    return true;
  });
  const total=feedback.length;
  const avg=filtered.length ? (filtered.reduce((s,f)=>s+Number(f.rating||0),0)/filtered.length).toFixed(1) : "0.0";
  const satCounts=Object.fromEntries(SATISFACTION_LEVELS.map(s=>[s.id,filtered.filter(f=>f.satisfaction===s.id).length]));
  const yes=filtered.filter(f=>f.recommend==="Yes").length;
  const no=filtered.filter(f=>f.recommend==="No").length;
  const unread=feedback.filter(f=>!f.reviewed).length;
  const deptRows=DEPTS.map(d=>({label:d,value:filtered.filter(f=>f.dept===d).length})).filter(r=>r.value);
  const categoryRows=FEEDBACK_CATEGORIES.map(c=>({label:c,value:filtered.filter(f=>f.category===c).length})).filter(r=>r.value);

  const markReviewed=async(id)=>{
    const current = feedback.find(f=>f.id===id);
    if(!current) return;
    const updated = {...current,reviewed:true,reviewedAt:Date.now()};
    try {
      await updateFeedback(updated);
      setFeedback(fs=>fs.map(f=>f.id===id?updated:f));
      toast("Feedback marked as reviewed","success");
    } catch (error) {
      console.error("Feedback review update failed:", error);
      toast("Feedback review update failed","error");
    }
  };

  const markAllReviewed=async()=>{
    const changed = feedback.filter(f=>!f.reviewed).map(f=>({...f,reviewed:true,reviewedAt:f.reviewedAt||Date.now()}));
    if(!changed.length) return;
    try {
      await Promise.all(changed.map(updateFeedback));
      setFeedback(fs=>fs.map(f=>changed.find(c=>c.id===f.id)||f));
      toast("All feedback marked reviewed","success");
    } catch (error) {
      console.error("Feedback bulk review update failed:", error);
      toast("Feedback review update failed","error");
    }
  };

  const doExport=()=>{
    if(!dateFrom || !dateTo){ toast("Please select date range","error"); return; }
    const now=new Date().toISOString().slice(0,10);
    if(format==="excel") downloadExcel(filtered.map(f=>{ const r=cleanFeedbackRow(f); return {"Feedback ID":r["Feedback ID"],"Ticket ID":r["Ticket ID"],"User Name":r["User Name"],Email:r.Email,Department:r.Department,Rating:r.Rating,Satisfaction:r.Satisfaction,Recommendation:r.Recommendation,"Feedback Message":r["Feedback Message"],Suggestions:r.Suggestions,"Submitted At":r["Submitted At"],"Reviewed Status":r["Reviewed Status"]}; }), `it_feedback_report_${now}.xlsx`);
    else downloadFeedbackPDF(filtered, `it_feedback_report_${now}.pdf`, {dateFrom,dateTo});
    toast(`Feedback report exported (${filtered.length} records)`,"success");
  };

  return (
    <div style={{display:"flex",flexDirection:"column",gap:22}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:14,flexWrap:"wrap"}}>
        <div><h2 style={{fontFamily:"Syne",fontSize:24,fontWeight:800,color:"#e2e8f0"}}>IT Feedback Dashboard</h2><p style={{fontSize:14,color:"rgba(226,232,240,0.5)",marginTop:4}}>Review user feedback, service quality, ratings, and improvement suggestions.</p></div>
        {unread>0&&<span className="tag" style={{background:"rgba(239,68,68,0.16)",color:"#f87171",border:"1px solid rgba(239,68,68,0.32)"}}>{unread} new</span>}
      </div>

      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(155px,1fr))",gap:14}}>
        <StatCard label="Total Feedback" value={total} icon="★" color="#818cf8" />
        <StatCard label="Average Rating" value={`${avg}/5`} icon="★" color="#fbbf24" />
        <StatCard label="Excellent" value={satCounts.Excellent||0} icon="✓" color="#34d399" />
        <StatCard label="Good" value={satCounts.Good||0} icon="+" color="#60a5fa" />
        <StatCard label="Recommend Yes" value={yes} icon="↑" color="#10b981" />
        <StatCard label="Recommend No" value={no} icon="↓" color="#f87171" />
      </div>

      <div className="glass" style={{padding:"18px 20px"}}>
        <div style={{fontSize:13,fontWeight:700,color:"rgba(226,232,240,0.62)",marginBottom:14}}>FEEDBACK REPORT EXPORT</div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(180px,1fr))",gap:14,alignItems:"end"}}>
          <div><label style={{fontSize:12,color:"rgba(226,232,240,0.5)",marginBottom:6,display:"block"}}>From Date</label><input type="date" value={dateFrom} onChange={e=>setDateFrom(e.target.value)} /></div>
          <div><label style={{fontSize:12,color:"rgba(226,232,240,0.5)",marginBottom:6,display:"block"}}>To Date</label><input type="date" value={dateTo} onChange={e=>setDateTo(e.target.value)} /></div>
          <div><label style={{fontSize:12,color:"rgba(226,232,240,0.5)",marginBottom:6,display:"block"}}>Format</label><select value={format} onChange={e=>setFormat(e.target.value)}><option value="excel">Excel (.xlsx)</option><option value="pdf">PDF (.pdf)</option></select></div>
          <button className="glow-btn" onClick={doExport}>Export Feedback</button>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(170px,1fr))",gap:14,alignItems:"end",marginTop:14}}>
          <div><label style={{fontSize:12,color:"rgba(226,232,240,0.5)",marginBottom:6,display:"block"}}>Ticket ID Filter</label><input value={ticketFilter} onChange={e=>setTicketFilter(e.target.value)} placeholder="TKT-..." /></div>
          <div><label style={{fontSize:12,color:"rgba(226,232,240,0.5)",marginBottom:6,display:"block"}}>Department</label><select value={deptFilter} onChange={e=>setDeptFilter(e.target.value)}><option value="All">All Departments</option>{DEPTS.map(d=><option key={d}>{d}</option>)}</select></div>
          <div><label style={{fontSize:12,color:"rgba(226,232,240,0.5)",marginBottom:6,display:"block"}}>Rating</label><select value={ratingFilter} onChange={e=>setRatingFilter(e.target.value)}><option value="All">All Ratings</option>{[5,4,3,2,1].map(r=><option key={r} value={r}>{r} Star</option>)}</select></div>
          <div><label style={{fontSize:12,color:"rgba(226,232,240,0.5)",marginBottom:6,display:"block"}}>Review Status</label><select value={reviewFilter} onChange={e=>setReviewFilter(e.target.value)}><option value="All">All</option><option value="Reviewed">Reviewed</option><option value="Unreviewed">Unreviewed</option></select></div>
        </div>
        <div style={{fontSize:12,color:"rgba(226,232,240,0.42)",marginTop:10}}>Showing {filtered.length} of {feedback.length} feedback submissions</div>
      </div>

      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(320px,1fr))",gap:14}}>
        <div className="glass" style={{padding:"18px 20px"}}>
          <div style={{fontSize:13,fontWeight:700,color:"rgba(226,232,240,0.62)",marginBottom:14}}>SATISFACTION DISTRIBUTION</div>
          {SATISFACTION_LEVELS.map(s=><MiniBar key={s.id} label={s.id} value={satCounts[s.id]||0} total={filtered.length} color={s.color} />)}
        </div>
        <div className="glass" style={{padding:"18px 20px"}}>
          <div style={{fontSize:13,fontWeight:700,color:"rgba(226,232,240,0.62)",marginBottom:14}}>DEPARTMENT SUMMARY</div>
          {(deptRows.length?deptRows:[{label:"No data",value:0}]).map((r,i)=><MiniBar key={r.label} label={r.label} value={r.value} total={filtered.length} color={["#6366f1","#0ea5e9","#10b981","#f59e0b","#ef4444"][i%5]} />)}
        </div>
        <div className="glass" style={{padding:"18px 20px"}}>
          <div style={{fontSize:13,fontWeight:700,color:"rgba(226,232,240,0.62)",marginBottom:14}}>SERVICE CATEGORY SUMMARY</div>
          {(categoryRows.length?categoryRows:[{label:"No data",value:0}]).map((r,i)=><MiniBar key={r.label} label={r.label} value={r.value} total={filtered.length} color={["#8b5cf6","#06b6d4","#f97316","#10b981","#ec4899"][i%5]} />)}
        </div>
      </div>

      <div className="glass" style={{padding:"18px 20px"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14,gap:10,flexWrap:"wrap"}}><div style={{fontSize:13,fontWeight:700,color:"rgba(226,232,240,0.62)"}}>RECENT FEEDBACK</div>{unread>0&&<button onClick={markAllReviewed} style={{background:"rgba(16,185,129,0.14)",border:"1px solid rgba(16,185,129,0.28)",color:"#34d399",padding:"8px 12px",borderRadius:8,fontSize:12,fontWeight:700}}>Mark all reviewed</button>}</div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(300px,1fr))",gap:12}}>
          {filtered.map(f=>{
            const sat=SATISFACTION_LEVELS.find(s=>s.id===f.satisfaction);
            return <div key={f.id} className="glass2" style={{padding:"16px",borderColor:f.reviewed?"rgba(255,255,255,0.1)":"rgba(245,158,11,0.45)",background:f.reviewed?"rgba(255,255,255,0.06)":"rgba(245,158,11,0.08)"}}>
              <div style={{display:"flex",justifyContent:"space-between",gap:10,alignItems:"flex-start",marginBottom:10}}><div><div style={{fontSize:13,fontWeight:800,color:"#e2e8f0"}}>{f.id}</div>{f.ticketId&&<div style={{fontSize:12,color:"#a5b4fc",marginTop:2}}>Ticket: {f.ticketId}</div>}<div style={{fontSize:12,color:"rgba(226,232,240,0.45)",marginTop:2}}>{fmtDate(f.createdAt)}</div></div>{!f.reviewed&&<span className="tag" style={{background:"rgba(245,158,11,0.16)",color:"#fbbf24"}}>New</span>}</div>
              <div style={{fontSize:14,fontWeight:700,color:"#e2e8f0"}}>{f.name}</div>
              <div style={{fontSize:12,color:"rgba(226,232,240,0.48)",marginTop:2}}>{f.email} · {f.dept}</div>
              <div style={{display:"flex",gap:8,flexWrap:"wrap",marginTop:10}}><span className="tag" style={{background:"rgba(99,102,241,0.14)",color:"#a5b4fc"}}>{f.category}</span><span className="tag" style={{background:"rgba(251,191,36,0.14)",color:"#fbbf24"}}>★ {f.rating}/5</span><span className="tag" style={{background:sat?.bg||"rgba(255,255,255,0.08)",color:sat?.color||"#e2e8f0"}}>{sat?.icon} {f.satisfaction}</span></div>
              <p style={{fontSize:13,lineHeight:1.6,color:"rgba(226,232,240,0.78)",marginTop:12}}>{f.message}</p>
              {f.suggestions&&<p style={{fontSize:12,lineHeight:1.5,color:"rgba(226,232,240,0.52)",marginTop:8}}>Suggestion: {f.suggestions}</p>}
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",gap:10,marginTop:12}}><span style={{fontSize:12,color:f.recommend==="Yes"?"#34d399":"#f87171",fontWeight:700}}>Recommend: {f.recommend}</span>{!f.reviewed&&<button onClick={()=>markReviewed(f.id)} style={{background:"rgba(99,102,241,0.16)",border:"1px solid rgba(99,102,241,0.32)",color:"#a5b4fc",padding:"7px 10px",borderRadius:8,fontSize:12,fontWeight:700}}>Mark reviewed</button>}</div>
            </div>;
          })}
          {filtered.length===0&&<div style={{gridColumn:"1/-1",textAlign:"center",padding:"50px 0",color:"rgba(226,232,240,0.34)"}}>No feedback found</div>}
        </div>
      </div>
    </div>
  );
}
// ── SIDEBAR ───────────────────────────────────────────────────────────────
function Sidebar({current,onChange,isAdmin,isStaff,tickets,feedback=[],mobileOpen,setMobileOpen}) {
  const adminNav=[{id:"dashboard",icon:"🏠",label:"Dashboard"},{id:"tickets",icon:"🎫",label:"All Tickets"},{id:"staff",icon:"👥",label:"IT Staff"},{id:"analytics",icon:"📊",label:"Analytics"},{id:"feedback",icon:"★",label:"IT Feedback"},{id:"export",icon:"⬇",label:"Export Reports"},{id:"staff-management",icon:"👥",label:"Staff Management"},{id:"emaillog",icon:"📧",label:"Email Log"}];
  const userNav=[{id:"home",icon:"🏠",label:"Home"},{id:"my-tickets",icon:"🎫",label:"My Tickets"},{id:"know-staff",icon:"👥",label:"Know Your IT Staff"},{id:"feedback",icon:"★",label:"IT Feedback"},{id:"new-ticket",icon:"➕",label:"New Ticket"},{id:"track",icon:"🔍",label:"Track Ticket"}];
  const staffNav=[{id:"staff-dash",icon:"🏠",label:"My Dashboard"},{id:"assigned",icon:"📋",label:"Assigned Tickets"}];
  const nav=isAdmin?adminNav:isStaff?staffNav:userNav;
  const open=tickets.filter(t=>t.status==="Open").length;
  const unreadFeedback=feedback.filter(f=>!f.reviewed).length;

  return (
    <>
      {mobileOpen&&<div className="sidebar-overlay" style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:49}} onClick={()=>setMobileOpen(false)}/>}
      <aside className={`app-sidebar ${mobileOpen ? "mobile-open" : ""}`} style={{
        width:240,flexShrink:0,background:"rgba(10,10,20,0.95)",backdropFilter:"blur(20px)",
        borderRight:"1px solid rgba(255,255,255,0.07)",display:"flex",flexDirection:"column",
        height:"100vh",position:"sticky",top:0,overflowY:"auto",
      }}>
        <div style={{padding:"18px 16px 14px",borderBottom:"1px solid rgba(255,255,255,0.06)"}}>
          <div style={{background:"rgba(255,255,255,0.95)",borderRadius:14,padding:"8px 10px",boxShadow:"0 12px 28px rgba(0,0,0,0.24)",display:"flex",alignItems:"center",justifyContent:"center"}}>
            <img src="/jaipuria-logo.png" alt="Jaipuria Institute of Management" style={{width:"100%",maxWidth:178,height:"auto",display:"block",objectFit:"contain"}} />
          </div>
        </div>
        <nav style={{flex:1,padding:"12px",display:"flex",flexDirection:"column",gap:3}}>
          {nav.map(item=>(
            <button key={item.id} onClick={()=>{onChange(item.id);setMobileOpen(false);}} style={{
              display:"flex",alignItems:"center",gap:12,padding:"10px 14px",borderRadius:10,border:"none",textAlign:"left",width:"100%",fontSize:14,
              background:current===item.id?"rgba(99,102,241,0.2)":"transparent",
              color:current===item.id?"#818cf8":"rgba(226,232,240,0.6)",
              fontWeight:current===item.id?600:400,
              borderLeft:current===item.id?"3px solid #6366f1":"3px solid transparent",
            }}>
              <span style={{fontSize:18}}>{item.icon}</span><span style={{flex:1}}>{item.label}</span>
              {item.id==="tickets"&&open>0&&<span style={{background:"rgba(239,68,68,0.2)",color:"#f87171",fontSize:11,padding:"2px 8px",borderRadius:20,fontWeight:600}}>{open}</span>}
              {item.id==="feedback"&&isAdmin&&unreadFeedback>0&&<span style={{background:"rgba(245,158,11,0.2)",color:"#fbbf24",fontSize:11,padding:"2px 8px",borderRadius:20,fontWeight:600}}>{unreadFeedback}</span>}
            </button>
          ))}
        </nav>
        <div style={{padding:"16px",borderTop:"1px solid rgba(255,255,255,0.06)",fontSize:11,color:"rgba(226,232,240,0.25)",textAlign:"center"}}>
          Jaipuria Institute of Management IT Helpdesk v2.0<br/>Enterprise Edition
        </div>
      </aside>
    </>
  );
}

// ── FIRST LOGIN / SET PASSWORD ─────────────────────────────────────────────
function SetPasswordScreen({staff,onComplete,toast}) {
  const [pwd,setPwd]=useState("");
  const [confirm,setConfirm]=useState("");
  const [loading,setLoading]=useState(false);
  const s=pwdStrength(pwd);

  const submit=async()=>{
    if(s<2){toast("Password too weak — add uppercase, numbers & symbols","error");return;}
    if(pwd!==confirm){toast("Passwords do not match","error");return;}
    setLoading(true);
    const hash=await hashPassword(pwd);
    await new Promise(r=>setTimeout(r,600));
    onComplete(hash);
    setLoading(false);
    toast("Password set successfully! 🎉","success");
  };

  return (
    <div style={{minHeight:"100vh",background:"radial-gradient(ellipse at 30% 40%,rgba(99,102,241,0.18) 0%,transparent 60%),#0a0a0f",display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{width:"100%",maxWidth:420}} className="fade-up">
        <div style={{textAlign:"center",marginBottom:32}}>
          <div style={{width:64,height:64,borderRadius:18,background:`${staff.color}33`,border:`2px solid ${staff.color}`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:24,fontWeight:800,color:staff.color,margin:"0 auto 16px"}}>{staff.avatar}</div>
          <h1 style={{fontFamily:"Syne",fontSize:24,fontWeight:800,color:"#e2e8f0",marginBottom:6}}>Welcome, {staff.name.split(" ")[0]}!</h1>
          <p style={{fontSize:14,color:"rgba(226,232,240,0.5)"}}>First login detected. Please create your secure password.</p>
        </div>
        <div className="glass" style={{padding:"28px 24px",display:"flex",flexDirection:"column",gap:18}}>
          <div style={{background:"rgba(245,158,11,0.1)",border:"1px solid rgba(245,158,11,0.3)",borderRadius:10,padding:"12px 14px",fontSize:13,color:"#fbbf24"}}>
            🔐 Your password is encrypted and stored securely. It cannot be recovered — keep it safe.
          </div>
          <div>
            <label style={{fontSize:12,color:"rgba(226,232,240,0.5)",marginBottom:6,display:"block"}}>Create Password</label>
            <PwdInput value={pwd} onChange={setPwd} placeholder="Minimum 8 characters" showStrength />
          </div>
          <div>
            <label style={{fontSize:12,color:"rgba(226,232,240,0.5)",marginBottom:6,display:"block"}}>Confirm Password</label>
            <PwdInput value={confirm} onChange={setConfirm} placeholder="Repeat your password"/>
            {confirm&&pwd!==confirm&&<div style={{fontSize:12,color:"#f87171",marginTop:4}}>Passwords do not match</div>}
          </div>
          <button className="glow-btn" style={{width:"100%"}} onClick={submit} disabled={loading||s<2||pwd!==confirm}>
            {loading?"⏳ Securing...":"🔐 Set Password & Continue"}
          </button>
        </div>
      </div>
    </div>
  );
}

// ── FORGOT PASSWORD ────────────────────────────────────────────────────────
function ForgotPassword({onBack,toast}) {
  const [step,setStep]=useState(1); // 1=email, 2=otp, 3=newpwd
  const [email,setEmail]=useState("");
  const [otp,setOtp]=useState("");
  const [inputOtp,setInputOtp]=useState("");
  const [newPwd,setNewPwd]=useState("");
  const [confirmPwd,setConfirmPwd]=useState("");
  const [loading,setLoading]=useState(false);
  const [token,setToken]=useState("");

  const sendOTP=async()=>{
    const cleanEmail=email.trim().toLowerCase();
    const isAdmin=cleanEmail==="admin@jaipuria.ac.in"||cleanEmail==="admin";
    if(cleanEmail.includes("@") && !isInstitutionEmail(cleanEmail)){toast("Only @jaipuria.ac.in email ID is allowed","error");return;}
    const staff=STAFF_BASE.find(s=>s.email.toLowerCase()===cleanEmail);
    if(!staff&&!isAdmin){toast("Email not found in our system","error");return;}
    setLoading(true);
    await new Promise(r=>setTimeout(r,800));
    const code=genOTP(); const tok=genToken();
    setOtp(code); setToken(tok);
    simulateEmail(isAdmin ? "admin@jaipuria.ac.in" : cleanEmail,"IT Helpdesk Password Reset OTP",`Your OTP for password reset is: ${code}\n\nThis OTP expires in 10 minutes.\n\nIf you did not request this, please ignore this email.`);
    toast(`OTP sent to ${isAdmin ? "admin@jaipuria.ac.in" : cleanEmail} (check Email Log for demo) 📧`,"email");
    setLoading(false); setStep(2);
  };

  const verifyOTP=()=>{
    if(inputOtp.trim()===otp){setStep(3);}
    else{toast("Invalid OTP. Check the Email Log.","error");}
  };

  const resetPassword=async()=>{
    if(pwdStrength(newPwd)<2){toast("Password too weak","error");return;}
    if(newPwd!==confirmPwd){toast("Passwords do not match","error");return;}
    setLoading(true);
    const hash=await hashPassword(newPwd);
    const staff=STAFF_BASE.find(s=>s.email.toLowerCase()===email.trim().toLowerCase());
    if(staff){
      const staffPasswords = DB.get("staff_passwords", {});
      staffPasswords[staff.id] = hash;
      DB.set("staff_passwords", staffPasswords);
    } else {
      DB.set("admin_password",hash);
    }
    await new Promise(r=>setTimeout(r,600));
    setLoading(false);
    toast("Password reset successfully! 🎉","success");
    setTimeout(onBack,1500);
  };

  return (
    <div style={{width:"100%",maxWidth:420}} className="fade-up">
      <div style={{textAlign:"center",marginBottom:28}}>
        <div style={{fontSize:48,marginBottom:12}}>🔑</div>
        <h2 style={{fontFamily:"Syne",fontSize:22,fontWeight:700,color:"#e2e8f0",marginBottom:6}}>Reset Password</h2>
        <p style={{fontSize:13,color:"rgba(226,232,240,0.5)"}}>
          {step===1?"Enter your registered email to receive OTP":step===2?"Enter the OTP sent to your email":"Create your new secure password"}
        </p>
      </div>
      {/* Step indicator */}
      <div style={{display:"flex",gap:8,marginBottom:24,justifyContent:"center"}}>
        {[1,2,3].map(n=><div key={n} style={{width:32,height:4,borderRadius:2,background:step>=n?"#6366f1":"rgba(255,255,255,0.1)",transition:"all .3s"}}/>)}
      </div>
      <div className="glass" style={{padding:"24px",display:"flex",flexDirection:"column",gap:16}}>
        {step===1&&<>
          <div><label style={{fontSize:12,color:"rgba(226,232,240,0.5)",marginBottom:6,display:"block"}}>Registered Email</label>
            <input type="email" placeholder="your@email.com" value={email} onChange={e=>setEmail(e.target.value)} onKeyDown={e=>e.key==="Enter"&&sendOTP()}/></div>
          <div style={{fontSize:12,color:"rgba(226,232,240,0.4)"}}>Use: admin, raj.singh@jaipuria.ac.in, rohit.jangid@jaipuria.ac.in, or vishal.swami@jaipuria.ac.in</div>
          <button className="glow-btn" style={{width:"100%"}} onClick={sendOTP} disabled={loading}>{loading?"⏳ Sending OTP...":"📧 Send OTP"}</button>
        </>}
        {step===2&&<>
          <div style={{background:"rgba(99,102,241,0.1)",border:"1px solid rgba(99,102,241,0.3)",borderRadius:8,padding:"12px 14px",fontSize:13,color:"#818cf8"}}>
            📧 OTP sent! Check the <strong>Email Log</strong> in Admin portal to see the demo OTP.
          </div>
          <div><label style={{fontSize:12,color:"rgba(226,232,240,0.5)",marginBottom:6,display:"block"}}>Enter 6-digit OTP</label>
            <input placeholder="123456" value={inputOtp} onChange={e=>setInputOtp(e.target.value)} maxLength={6} onKeyDown={e=>e.key==="Enter"&&verifyOTP()}/></div>
          <button className="glow-btn" style={{width:"100%"}} onClick={verifyOTP}>Verify OTP →</button>
        </>}
        {step===3&&<>
          <div><label style={{fontSize:12,color:"rgba(226,232,240,0.5)",marginBottom:6,display:"block"}}>New Password</label>
            <PwdInput value={newPwd} onChange={setNewPwd} placeholder="New password" showStrength onKeyDown={e=>e.key==="Enter"&&resetPassword()}/></div>
          <div><label style={{fontSize:12,color:"rgba(226,232,240,0.5)",marginBottom:6,display:"block"}}>Confirm Password</label>
            <PwdInput value={confirmPwd} onChange={setConfirmPwd} placeholder="Repeat new password" onKeyDown={e=>e.key==="Enter"&&resetPassword()}/>
            {confirmPwd&&newPwd!==confirmPwd&&<div style={{fontSize:12,color:"#f87171",marginTop:4}}>Passwords do not match</div>}
          </div>
          <button className="glow-btn" style={{width:"100%"}} onClick={resetPassword} disabled={loading||pwdStrength(newPwd)<2||newPwd!==confirmPwd}>
            {loading?"⏳ Updating...":"✅ Reset Password"}
          </button>
        </>}
        <button onClick={onBack} style={{background:"none",border:"none",color:"rgba(226,232,240,0.4)",fontSize:13,textAlign:"center"}}>← Back to Login</button>
      </div>
    </div>
  );
}

// ── LANDING / LOGIN ────────────────────────────────────────────────────────
function Landing({onLogin,tickets=[]}) {
  const [mode,setMode]=useState("user");
  const [email,setEmail]=useState("");
  const [pwd,setPwd]=useState("");
  const [loading,setLoading]=useState(false);
  const [showForgot,setShowForgot]=useState(false);
  const {toasts,toast,remove}=useToast();
  const [adminUser,setAdminUser]=useState("");

  const resetMode = (nextMode) => {
    setMode(nextMode);
    setPwd("");
    setEmail("");
    setAdminUser("");
  };

  const validateInstituteEmail = (value) => {
    if(!isInstitutionEmail(value)){
      toast("Only @jaipuria.ac.in email ID is allowed","error");
      return false;
    }
    return true;
  };

  const handleLogin=async()=>{
    const cleanEmail=email.trim().toLowerCase();
    const cleanAdmin=adminUser.trim().toLowerCase();

    if(mode==="user"){
      if(!cleanEmail){toast("Enter your email address","error");return;}
      if(!validateInstituteEmail(cleanEmail)) return;
      onLogin({type:"user",email:cleanEmail});
      return;
    }

    if(mode==="admin"){
      if(!cleanAdmin||!pwd.trim()){toast("Enter username and password","error");return;}
      if(cleanAdmin.includes("@") && !validateInstituteEmail(cleanAdmin)) return;
      if(cleanAdmin!=="admin" && cleanAdmin!=="admin@jaipuria.ac.in"){toast("Invalid username","error");return;}
      setLoading(true);
      await new Promise(r=>setTimeout(r,500));
      const storedHash=DB.get("admin_password",null);
      let valid=false;
      if(storedHash){valid=await verifyPassword(pwd,storedHash);}
      else{valid=pwd==="Admin@123";}
      setLoading(false);
      if(valid){onLogin({type:"admin",email:"admin@jaipuria.ac.in"});}
      else{toast("Invalid credentials","error");}
      return;
    }

    if(mode==="staff"){
      if(!cleanEmail){toast("Enter staff email","error");return;}
      if(!validateInstituteEmail(cleanEmail)) return;
      const staff=STAFF_BASE.find(s=>s.email.toLowerCase()===cleanEmail);
      if(!staff){toast("Staff email not found","error");return;}
      setLoading(true);
      await new Promise(r=>setTimeout(r,500));
      const staffPasswords=DB.get("staff_passwords",{});
      const storedHash=staffPasswords[staff.id];
      if(!storedHash){
        setLoading(false);
        onLogin({type:"staff_firstlogin",staffId:staff.id,staff});
        return;
      }
      if(!pwd){toast("Enter your password","error");setLoading(false);return;}
      const valid=await verifyPassword(pwd,storedHash);
      setLoading(false);
      if(valid){onLogin({type:"staff",staffId:staff.id,email:staff.email,name:staff.name,role:staff.role,permissions:staff.permissions});}
      else{toast("Incorrect password","error");}
    }
  };

  if(showForgot) return (
    <div style={{minHeight:"100vh",background:"linear-gradient(135deg,#111436 0%,#1d1450 48%,#071827 100%)",display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <ForgotPassword onBack={()=>setShowForgot(false)} toast={toast}/>
      <Toast toasts={toasts} remove={remove}/>
    </div>
  );

  const modeCopy = {
    user:{title:"User Login",emailLabel:"Email",placeholder:"name@jaipuria.ac.in",needsPassword:false},
    staff:{title:"Staff Login",emailLabel:"Email",placeholder:"staff@jaipuria.ac.in",needsPassword:true},
    admin:{title:"Admin Login",emailLabel:"Username / Email",placeholder:"Admin or admin@jaipuria.ac.in",needsPassword:true},
  }[mode];

  const liveTickets = [...tickets]
    .sort((a,b)=>(b.updatedAt||b.createdAt||0)-(a.updatedAt||a.createdAt||0))
    .slice(0,3);
  const activeTickets = tickets.filter(t=>!["Closed","Resolved"].includes(t.status));
  const criticalTickets = tickets.filter(t=>t.priority==="Critical");

  return (
    <div style={{minHeight:"100vh",background:"linear-gradient(135deg,#111436 0%,#211552 52%,#071827 100%)",display:"flex",alignItems:"center",justifyContent:"center",padding:"24px 16px",position:"relative",overflow:"hidden"}}>
      <div style={{position:"absolute",inset:0,background:"radial-gradient(circle at 15% 20%,rgba(99,102,241,0.28),transparent 32%),radial-gradient(circle at 85% 18%,rgba(14,165,233,0.20),transparent 30%),radial-gradient(circle at 72% 86%,rgba(16,185,129,0.16),transparent 28%)",pointerEvents:"none"}} />
      <div className="fade-up" style={{width:"100%",maxWidth:1040,display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(min(100%,360px),1fr))",background:"rgba(8,12,34,0.82)",border:"1px solid rgba(255,255,255,0.10)",borderRadius:26,boxShadow:"0 34px 110px rgba(0,0,0,0.42)",overflow:"hidden",position:"relative",backdropFilter:"blur(22px)"}}>
        <form onSubmit={e=>{e.preventDefault();handleLogin();}} style={{padding:"clamp(26px,5vw,54px)",display:"flex",flexDirection:"column",justifyContent:"center",minHeight:560}}>
          <div style={{display:"flex",alignItems:"center",gap:14,marginBottom:30}}>
            <div style={{background:"rgba(255,255,255,0.96)",border:"1px solid rgba(255,255,255,0.35)",borderRadius:16,padding:"8px 12px",boxShadow:"0 14px 34px rgba(0,0,0,0.24)",display:"flex",alignItems:"center",justifyContent:"center"}}>
              <img src="/jaipuria-logo.png" alt="Jaipuria Institute of Management" style={{width:"clamp(210px,25vw,260px)",height:"auto",display:"block",objectFit:"contain"}} />
            </div>
          </div>

          <div style={{marginBottom:24}}>
            <div style={{fontFamily:"Syne",fontWeight:800,fontSize:"clamp(20px,2.5vw,28px)",lineHeight:1.12,color:"#f8fafc",letterSpacing:0,whiteSpace:"nowrap"}}>Login With Jaipuria Email ID</div>
            <div style={{fontSize:14,color:"rgba(226,232,240,0.62)",marginTop:10}}>Secure access for users, IT staff, and administrators.</div>
          </div>

          <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",background:"rgba(255,255,255,0.06)",border:"1px solid rgba(255,255,255,0.09)",borderRadius:14,padding:5,marginBottom:22,gap:4}}>
            {[["user","User"],["staff","Staff"],["admin","Admin"]].map(([m,l])=>(
              <button type="button" key={m} onClick={()=>resetMode(m)} style={{padding:"10px 8px",borderRadius:10,border:"none",fontSize:13,fontWeight:700,background:mode===m?"linear-gradient(135deg,#6366f1,#14b8a6)":"transparent",color:mode===m?"#fff":"rgba(226,232,240,0.58)",boxShadow:mode===m?"0 10px 28px rgba(99,102,241,0.25)":"none"}}>{l}</button>
            ))}
          </div>

          <div style={{display:"flex",flexDirection:"column",gap:14}}>
            <div>
              <label style={{fontSize:12,color:"rgba(226,232,240,0.68)",marginBottom:7,display:"block",fontWeight:700}}>{modeCopy.emailLabel}</label>
              {mode==="admin" ? (
                <input placeholder={modeCopy.placeholder} value={adminUser} autoComplete="username" onChange={e=>setAdminUser(e.target.value)} />
              ) : (
                <input type="email" name="login-email" autoComplete="username" placeholder={modeCopy.placeholder} value={email} onChange={e=>setEmail(e.target.value)} />
              )}
              {mode!=="admin"&&<div style={{fontSize:11,color:"rgba(226,232,240,0.36)",marginTop:7}}>Only @jaipuria.ac.in email IDs are allowed.</div>}
            </div>

            {modeCopy.needsPassword&&(
              <div>
                <label style={{fontSize:12,color:"rgba(226,232,240,0.68)",marginBottom:7,display:"block",fontWeight:700}}>Password</label>
                <PwdInput value={pwd} onChange={setPwd} autoComplete="current-password" placeholder="Enter password" />
              </div>
            )}
          </div>

          <button type="submit" className="glow-btn" style={{width:"100%",padding:"14px",marginTop:22,boxShadow:"0 16px 38px rgba(99,102,241,0.32)"}} disabled={loading}>
            {loading?"Authenticating...":mode==="admin"?"Login as Admin":"Login"}
          </button>

          <div style={{display:"flex",justifyContent:mode==="admin"?"space-between":"flex-end",alignItems:"center",gap:12,marginTop:14,flexWrap:"wrap"}}>
            {mode!=="user"&&<button type="button" onClick={()=>setShowForgot(true)} style={{background:"none",border:"none",color:"rgba(129,140,248,0.92)",fontSize:13,textDecoration:"underline"}}>Forgot Password?</button>}
          </div>
        </form>

        <div style={{minHeight:560,padding:"clamp(26px,5vw,54px)",background:"linear-gradient(160deg,rgba(99,102,241,0.18),rgba(20,184,166,0.08))",borderLeft:"1px solid rgba(255,255,255,0.08)",display:"flex",alignItems:"center",justifyContent:"center"}}>
          <div style={{width:"100%",maxWidth:390,position:"relative"}}>
            <div className="glass" style={{padding:22,borderRadius:20,boxShadow:"0 22px 70px rgba(0,0,0,0.30)"}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:22}}>
                <div>
                  <div style={{fontSize:12,color:"rgba(226,232,240,0.48)",fontWeight:700}}>LIVE QUEUE</div>
                  <div style={{fontFamily:"Syne",fontSize:26,fontWeight:800,color:"#fff",marginTop:4}}>Helpdesk Tickets</div>
                </div>
                <div style={{width:50,height:50,borderRadius:16,background:"linear-gradient(135deg,#10b981,#0ea5e9)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:26}}>✓</div>
              </div>

              {liveTickets.length > 0 ? liveTickets.map(ticket=>{
                const color=statusColor(ticket.status);
                const assignee=staffName(ticket.assigneeId);
                return (
                  <div key={ticket.id} style={{background:"rgba(255,255,255,0.06)",border:"1px solid rgba(255,255,255,0.08)",borderRadius:14,padding:"14px 15px",marginBottom:12,display:"flex",alignItems:"center",gap:12}}>
                    <div style={{width:10,height:44,borderRadius:999,background:color,boxShadow:`0 0 18px ${color}88`}} />
                    <div style={{flex:1,minWidth:0}}>
                      <div style={{fontSize:14,fontWeight:800,color:"#e2e8f0",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{categoryLabel(ticket.category)}</div>
                      <div style={{fontSize:12,color:"rgba(226,232,240,0.45)",marginTop:3,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{ticket.id} · Assigned to {assignee}</div>
                    </div>
                    <span style={{fontSize:11,fontWeight:800,color,background:`${color}20`,border:`1px solid ${color}55`,borderRadius:999,padding:"5px 8px",whiteSpace:"nowrap"}}>{ticket.status}</span>
                  </div>
                );
              }) : (
                <div style={{background:"rgba(255,255,255,0.06)",border:"1px solid rgba(255,255,255,0.08)",borderRadius:14,padding:"26px 16px",textAlign:"center",color:"rgba(226,232,240,0.58)"}}>
                  No live tickets yet
                </div>
              )}

              <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10,marginTop:18}}>
                {[[tickets.length,"Total"],[activeTickets.length,"Active"],[criticalTickets.length,"Critical"]].map(([v,l])=>(
                  <div key={l} style={{textAlign:"center",background:"rgba(15,23,42,0.5)",border:"1px solid rgba(255,255,255,0.08)",borderRadius:14,padding:"14px 8px"}}>
                    <div style={{fontFamily:"Syne",fontSize:20,fontWeight:800,color:l==="Critical"?"#f87171":l==="Active"?"#38bdf8":"#818cf8"}}>{v}</div>
                    <div style={{fontSize:11,color:"rgba(226,232,240,0.42)",marginTop:2}}>{l}</div>
                  </div>
                ))}
              </div>
            </div>
          </div>
        </div>
      </div>
      <Toast toasts={toasts} remove={remove}/>
    </div>
  );
}
// ── TICKETS TABLE ─────────────────────────────────────────────────────────
function TicketsTable({tickets,onView,isAdmin,onDelete}) {
  const [search,setSearch]=useState("");
  const [filterStatus,setFilterStatus]=useState("All");
  const [filterPriority,setFilterPriority]=useState("All");
  const [sort,setSort]=useState("newest");
  let filtered=tickets.filter(t=>{const q=search.toLowerCase();return !q||t.id.toLowerCase().includes(q)||t.name.toLowerCase().includes(q)||t.description.toLowerCase().includes(q);});
  if(filterStatus!=="All") filtered=filtered.filter(t=>t.status===filterStatus);
  if(filterPriority!=="All") filtered=filtered.filter(t=>t.priority===filterPriority);
  filtered=[...filtered].sort((a,b)=>sort==="newest"?b.createdAt-a.createdAt:PRIORITIES.indexOf(b.priority)-PRIORITIES.indexOf(a.priority));
  return (
    <div style={{display:"flex",flexDirection:"column",gap:16}}>
      <h2 style={{fontFamily:"Syne",fontSize:22,fontWeight:700,color:"#e2e8f0"}}>All Tickets</h2>
      <div style={{display:"flex",gap:10,flexWrap:"wrap",alignItems:"center"}}>
        <input placeholder="🔍 Search by ID, name, description..." value={search} onChange={e=>setSearch(e.target.value)} style={{flex:1,minWidth:180}}/>
        <select value={filterStatus} onChange={e=>setFilterStatus(e.target.value)} style={{width:"auto"}}><option value="All">All Status</option>{STATUSES.map(s=><option key={s}>{s}</option>)}</select>
        <select value={filterPriority} onChange={e=>setFilterPriority(e.target.value)} style={{width:"auto"}}><option value="All">All Priority</option>{PRIORITIES.map(p=><option key={p}>{p}</option>)}</select>
        <select value={sort} onChange={e=>setSort(e.target.value)} style={{width:"auto"}}><option value="newest">Newest</option><option value="priority">Priority</option></select>
      </div>
      <div style={{fontSize:13,color:"rgba(226,232,240,0.4)"}}>{filtered.length} ticket{filtered.length!==1?"s":""}</div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(290px,1fr))",gap:12}}>
        {filtered.map(t=>(
          <div key={t.id} style={{position:"relative"}}>
            <TicketCard ticket={t} onView={onView}/>
            {isAdmin&&<button onClick={e=>{e.stopPropagation();onDelete(t.id);}} style={{position:"absolute",top:10,right:10,background:"rgba(239,68,68,0.15)",border:"1px solid rgba(239,68,68,0.3)",color:"#f87171",width:26,height:26,borderRadius:6,fontSize:13,display:"flex",alignItems:"center",justifyContent:"center"}}>🗑</button>}
          </div>
        ))}
        {filtered.length===0&&<div style={{gridColumn:"1/-1",textAlign:"center",padding:"60px 0",color:"rgba(226,232,240,0.3)"}}><div style={{fontSize:48,marginBottom:12}}>🎫</div><div>No tickets found</div></div>}
      </div>
    </div>
  );
}

// ── TRACK TICKET ──────────────────────────────────────────────────────────
function TrackTicket({tickets,onView}) {
  const [id,setId]=useState(""); const [result,setResult]=useState(null); const [searched,setSearched]=useState(false);
  const search=()=>{setSearched(true);setResult(tickets.find(t=>t.id.toLowerCase()===id.trim().toLowerCase())||null);};
  return (
    <div style={{maxWidth:520}}>
      <h2 style={{fontFamily:"Syne",fontSize:22,fontWeight:700,color:"#e2e8f0",marginBottom:20}}>Track Ticket</h2>
      <div style={{display:"flex",gap:10,marginBottom:20}}>
        <input placeholder="Enter Ticket ID (e.g. TKT-ABC123)" value={id} onChange={e=>setId(e.target.value)} onKeyDown={e=>e.key==="Enter"&&search()}/>
        <button className="glow-btn" onClick={search} style={{whiteSpace:"nowrap"}}>Search</button>
      </div>
      {searched&&!result&&<div className="glass" style={{padding:"24px",textAlign:"center",color:"rgba(226,232,240,0.4)"}}><div style={{fontSize:36,marginBottom:8}}>🔍</div><div>No ticket found: {id}</div></div>}
      {result&&<div className="glass2" style={{padding:"20px",cursor:"pointer"}} onClick={()=>onView(result.id)}>
        <div style={{display:"flex",justifyContent:"space-between",marginBottom:10}}><div style={{fontFamily:"Syne",fontSize:16,fontWeight:700,color:"#e2e8f0"}}>{result.id}</div><StatusBadge status={result.status}/></div>
        <div style={{fontSize:14,color:"rgba(226,232,240,0.7)",marginBottom:8}}>{result.description.slice(0,100)}...</div>
        <PriorityBadge p={result.priority}/><div style={{marginTop:10}}><TimerBadge ticket={result}/></div>
        <div style={{marginTop:10,fontSize:12,color:"#818cf8"}}>Click to view full details →</div>
      </div>}
    </div>
  );
}

// ── STAFF PANEL ───────────────────────────────────────────────────────────
function StaffPanel({staffId,tickets,setTickets,toast,onViewTicket,permissions,staffProfiles={},staffStatuses={}}) {
  const staff=STAFF_BASE.find(s=>s.id===staffId);
  const myTickets=tickets.filter(t=>t.assigneeId===staffId);
  const active=myTickets.filter(t=>!["Resolved","Closed"].includes(t.status)).length;
  const resolved=myTickets.filter(t=>t.status==="Resolved"||t.status==="Closed").length;
  return (
    <div style={{display:"flex",flexDirection:"column",gap:20}}>
      <div className="glass" style={{padding:"24px",display:"flex",gap:16,alignItems:"center",flexWrap:"wrap"}}>
        <StaffAvatar staff={staff} profiles={staffProfiles} statuses={staffStatuses} size={56} showStatus />
        <div style={{flex:1}}>
          <div style={{fontFamily:"Syne",fontSize:20,fontWeight:700,color:"#e2e8f0"}}>{staff.name}</div>
          <div style={{fontSize:13,color:"rgba(226,232,240,0.5)"}}>{staff.role} · {staff.email}</div>
          <div style={{display:"flex",gap:8,marginTop:6,flexWrap:"wrap"}}>
            {(permissions||[]).map(p=><span key={p} className="tag" style={{background:"rgba(99,102,241,0.1)",color:"#818cf8",fontSize:11}}>{p}</span>)}
          </div>
        </div>
        <div style={{display:"flex",gap:14,textAlign:"center",flexWrap:"wrap"}}>
          {[["Total",myTickets.length,"#818cf8"],["Active",active,"#fbbf24"],["Resolved",resolved,"#34d399"]].map(([l,v,c])=>(
            <div key={l}><div style={{fontSize:24,fontFamily:"Syne",fontWeight:800,color:c}}>{v}</div><div style={{fontSize:11,color:"rgba(226,232,240,0.4)"}}>{l}</div></div>
          ))}
        </div>
      </div>
      <h3 style={{fontFamily:"Syne",fontSize:16,fontWeight:700,color:"#e2e8f0"}}>Assigned Tickets</h3>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(280px,1fr))",gap:12}}>
        {[...myTickets].sort((a,b)=>b.createdAt-a.createdAt).map(t=><TicketCard key={t.id} ticket={t} onView={onViewTicket}/>)}
        {myTickets.length===0&&<div style={{gridColumn:"1/-1",textAlign:"center",padding:"60px 0",color:"rgba(226,232,240,0.3)"}}><div style={{fontSize:48,marginBottom:12}}>✅</div><div>No tickets assigned yet</div></div>}
      </div>
    </div>
  );
}

// ── STAFF PROFILE / PERFORMANCE / CHAT ───────────────────────────────────
function StaffProfileMenu({staff,profiles,statuses,onStatusChange,onOpen,onLogout,onClose}) {
  const status=getStaffStatus(staff.id,statuses);
  return (
    <div className="glass staff-profile-menu" style={{position:"absolute",left:0,top:"calc(100% + 12px)",width:"min(320px, calc(100vw - 32px))",padding:12,zIndex:2000,boxShadow:"0 28px 90px rgba(0,0,0,0.48)",background:"rgba(10,10,20,0.96)",border:"1px solid rgba(255,255,255,0.14)",backdropFilter:"blur(24px)"}}>
      <div style={{display:"flex",gap:12,alignItems:"center",padding:"10px 10px 14px",borderBottom:"1px solid rgba(255,255,255,0.08)",marginBottom:10}}>
        <StaffAvatar staff={staff} profiles={profiles} statuses={statuses} size={52} showStatus />
        <div style={{minWidth:0}}>
          <div style={{fontSize:15,fontWeight:800,color:"#fff",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{staff.name}</div>
          <div style={{fontSize:12,color:"rgba(226,232,240,0.52)",marginTop:2}}>{staff.role}</div>
          <div style={{marginTop:4}}><StatusDot status={status}/></div>
        </div>
      </div>
      <div style={{padding:"6px 8px 10px"}}>
        <label style={{fontSize:11,color:"rgba(226,232,240,0.45)",display:"block",marginBottom:6,fontWeight:700}}>Live Status</label>
        <select value={status} onChange={e=>onStatusChange(e.target.value)} style={{fontSize:12,padding:"8px 10px"}}>
          {Object.keys(STAFF_STATUS).map(s=><option key={s} value={s}>{s}</option>)}
        </select>
      </div>
      {[['profile','My Profile'],['password','Change Password'],['performance','My Performance'],['chat','Staff Chat']].map(([id,label])=>(
        <button key={id} onClick={()=>{onClose&&onClose();onOpen(id);}} style={{width:"100%",textAlign:"left",background:"transparent",border:"none",color:"rgba(226,232,240,0.82)",padding:"11px 12px",borderRadius:10,fontSize:13,fontWeight:700}}>{label}</button>
      ))}
      <button onClick={onLogout} style={{width:"100%",textAlign:"left",background:"rgba(239,68,68,0.13)",border:"1px solid rgba(239,68,68,0.26)",color:"#f87171",padding:"11px 12px",borderRadius:10,fontSize:13,fontWeight:800,marginTop:8}}>Logout</button>
    </div>
  );
}

function StaffProfileModal({staff,profiles,statuses,onSave,toast}) {
  const [photo,setPhoto]=useState(getStaffPhoto(staff.id,profiles));
  const handleFile=e=>{
    const file=e.target.files?.[0];
    if(!file) return;
    const reader=new FileReader();
    reader.onload=()=>setPhoto(reader.result);
    reader.readAsDataURL(file);
  };
  return (
    <div style={{display:"flex",flexDirection:"column",gap:18}}>
      <div style={{display:"flex",gap:16,alignItems:"center"}}>
        <StaffAvatar staff={{...staff}} profiles={{[staff.id]:{photo}}} statuses={statuses} size={82} showStatus />
        <div><h3 style={{fontFamily:"Syne",fontSize:20,color:"#fff"}}>{staff.name}</h3><div style={{fontSize:13,color:"rgba(226,232,240,0.5)"}}>{staff.role}</div><StatusDot status={getStaffStatus(staff.id,statuses)}/></div>
      </div>
      <div className="glass" style={{padding:16,display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
        {[['Name',staff.name],['Email',staff.email],['Role',staff.role],['Permissions',staff.permissions.join(', ')]].map(([l,v])=><div key={l}><div style={{fontSize:11,color:"rgba(226,232,240,0.4)"}}>{l}</div><div style={{fontSize:13,color:"#e2e8f0",fontWeight:600,marginTop:3}}>{v}</div></div>)}
      </div>
      <div><label style={{fontSize:12,color:"rgba(226,232,240,0.65)",marginBottom:7,display:"block",fontWeight:700}}>Profile Photo</label><input type="file" accept="image/*" onChange={handleFile}/></div>
      <button className="glow-btn" onClick={()=>{onSave(staff.id,{photo});toast('Profile updated','success');}}>Save Profile</button>
    </div>
  );
}

function StaffChangePasswordModal({staff,toast}) {
  const [oldPwd,setOldPwd]=useState('');
  const [newPwd,setNewPwd]=useState('');
  const [confirm,setConfirm]=useState('');
  const save=async()=>{
    const passwords=DB.get('staff_passwords',{});
    const current=passwords[staff.id];
    if(!current){toast('No password set yet. Please complete first login.','error');return;}
    if(!(await verifyPassword(oldPwd,current))){toast('Old password is incorrect','error');return;}
    if(pwdStrength(newPwd)<3){toast('New password is too weak','error');return;}
    if(newPwd!==confirm){toast('Passwords do not match','error');return;}
    passwords[staff.id]=await hashPassword(newPwd);
    DB.set('staff_passwords',passwords);
    setOldPwd('');setNewPwd('');setConfirm('');
    toast('Password changed successfully','success');
  };
  return <div style={{display:'flex',flexDirection:'column',gap:16}}>
    <div><label style={{fontSize:12,color:'rgba(226,232,240,0.65)',marginBottom:6,display:'block'}}>Old Password</label><PwdInput value={oldPwd} onChange={setOldPwd} placeholder="Old password"/></div>
    <div><label style={{fontSize:12,color:'rgba(226,232,240,0.65)',marginBottom:6,display:'block'}}>New Password</label><PwdInput value={newPwd} onChange={setNewPwd} placeholder="New password" showStrength/></div>
    <div><label style={{fontSize:12,color:'rgba(226,232,240,0.65)',marginBottom:6,display:'block'}}>Confirm New Password</label><PwdInput value={confirm} onChange={setConfirm} placeholder="Confirm password"/></div>
    <button className="glow-btn" onClick={save}>Change Password</button>
  </div>;
}

function StaffPerformanceModal({staff,tickets}) {
  const mine=tickets.filter(t=>t.assigneeId===staff.id);
  const closed=mine.filter(t=>t.status==='Closed'||t.status==='Resolved');
  const avg=closed.length?closed.reduce((s,t)=>s+((t.closedAt||Date.now())-t.createdAt),0)/closed.length:0;
  const breached=mine.filter(t=>(t.closedAt||Date.now())-t.createdAt>SLA_HOURS[t.priority]*3600000).length;
  const rate=mine.length?Math.round((closed.length/mine.length)*100):0;
  const months=[...Array(6)].map((_,i)=>{const d=new Date();d.setMonth(d.getMonth()-5+i);return {key:`${d.getFullYear()}-${d.getMonth()}`,label:d.toLocaleString('en-IN',{month:'short'}),count:closed.filter(t=>{const c=new Date(t.closedAt||0);return c.getFullYear()===d.getFullYear()&&c.getMonth()===d.getMonth();}).length};});
  const cats=CATEGORIES.map(c=>({label:c.label,color:c.color,count:mine.filter(t=>t.category===c.id).length})).filter(c=>c.count);
  return <div style={{display:'flex',flexDirection:'column',gap:20}}>
    <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fill,minmax(140px,1fr))',gap:12}}>
      {[["Assigned",mine.length,'#818cf8'],["Open",mine.filter(t=>t.status==='Open'||t.status==='Assigned').length,'#38bdf8'],["In Progress",mine.filter(t=>t.status==='In Progress').length,'#fbbf24'],["Closed",closed.length,'#34d399'],["Avg Time",avg?formatDuration(avg):'—','#f97316'],["SLA Breach",breached,'#f87171'],["Rate",`${rate}%`,'#10b981']].map(([l,v,c])=><div key={l} className="glass" style={{padding:16}}><div style={{fontSize:12,color:'rgba(226,232,240,0.45)'}}>{l}</div><div style={{fontFamily:'Syne',fontSize:24,fontWeight:800,color:c,marginTop:6}}>{v}</div></div>)}
    </div>
    <div className="glass" style={{padding:18}}><h3 style={{fontFamily:'Syne',fontSize:16,color:'#fff',marginBottom:14}}>Monthly Closed Tickets</h3>{months.map(m=><div key={m.key} style={{display:'flex',alignItems:'center',gap:12,marginBottom:10}}><div style={{width:42,fontSize:12,color:'rgba(226,232,240,0.55)'}}>{m.label}</div><div style={{flex:1,height:8,background:'rgba(255,255,255,0.07)',borderRadius:6}}><div style={{width:`${Math.min(100,m.count*20)}%`,height:'100%',background:'#6366f1',borderRadius:6}}/></div><div style={{width:24,color:'#e2e8f0',fontSize:12,fontWeight:700}}>{m.count}</div></div>)}</div>
    <div className="glass" style={{padding:18}}><h3 style={{fontFamily:'Syne',fontSize:16,color:'#fff',marginBottom:14}}>Category-wise Tickets</h3>{cats.length?cats.map(c=><div key={c.label} style={{display:'flex',alignItems:'center',gap:12,marginBottom:10}}><div style={{width:150,fontSize:12,color:'rgba(226,232,240,0.68)'}}>{c.label}</div><div style={{flex:1,height:8,background:'rgba(255,255,255,0.07)',borderRadius:6}}><div style={{width:`${Math.min(100,c.count*18)}%`,height:'100%',background:c.color,borderRadius:6}}/></div><div style={{width:24,color:c.color,fontSize:12,fontWeight:800}}>{c.count}</div></div>):<div style={{color:'rgba(226,232,240,0.4)'}}>No category data yet</div>}</div>
  </div>;
}

function KnowYourITStaff({staffProfiles}) {
  const staffData = [
    {
      id: 1,
      name: "Raj Parkash Singh",
      post: "Senior IT Support Executive",
      email: "raj.parkash@jaipuria.ac.in",
      contact: "9887283825",
      whatsapp: "9887283825"
    },
    {
      id: 2,
      name: "Rohit Jangid",
      post: "IT Support Executive",
      email: "rohit.jangid@jaipuria.ac.in",
      contact: "8005978632",
      whatsapp: "8005978632"
    },
    {
      id: 3,
      name: "Vishal Swami",
      post: "IT Executive",
      email: "vishal.swami@jaipuria.ac.in",
      contact: "8233771101",
      whatsapp: "8233771101"
    }
  ];

  return (
    <div>
      <h2 style={{fontFamily:"Syne",fontSize:22,fontWeight:700,color:"#e2e8f0",marginBottom:20}}>Know Your IT Staff</h2>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(320px,1fr))",gap:20}}>
        {staffData.map(staff => {
          const profile = staffProfiles[staff.id] || {};
          return (
            <div key={staff.id} className="glass" style={{
              background: "linear-gradient(135deg, rgba(99,102,241,0.15), rgba(6,182,212,0.08), rgba(16,185,129,0.06))",
              border: "1px solid rgba(255,255,255,0.1)",
              borderRadius: 16,
              padding: 24,
              position: "relative",
              overflow: "hidden"
            }}>
              <div style={{
                position: "absolute",
                top: 0,
                left: 0,
                right: 0,
                bottom: 0,
                background: "radial-gradient(circle at 20% 20%, rgba(168,85,247,0.1), transparent 50%), radial-gradient(circle at 80% 80%, rgba(6,182,212,0.08), transparent 50%)",
                pointerEvents: "none"
              }} />
              <div style={{display:"flex",flexDirection:"column",alignItems:"center",textAlign:"center",position:"relative",zIndex:1}}>
                <StaffAvatar staff={{...staff, avatar: staff.name.split(' ').map(n=>n[0]).join(''), color: ["#6366f1", "#0ea5e9", "#10b981"][staff.id-1]}} profiles={{[staff.id]: profile}} statuses={{}} size={80} showStatus={false} />
                <h3 style={{fontFamily:"Syne",fontSize:20,fontWeight:800,color:"#e2e8f0",margin:"16px 0 4px"}}>{staff.name}</h3>
                <div style={{fontSize:14,color:"rgba(6,182,212,0.9)",fontWeight:600,marginBottom:20}}>{staff.post}</div>
                <div style={{display:"flex",flexDirection:"column",gap:12,width:"100%"}}>
                  <div style={{display:"flex",gap:10}}>
                    <button
                      onClick={() => window.open(`mailto:${staff.email}`, '_blank', 'noopener,noreferrer')}
                      style={{
                        flex:1,
                        background:"rgba(99,102,241,0.2)",
                        border:"1px solid rgba(99,102,241,0.4)",
                        color:"#dbeafe",
                        padding:"10px 14px",
                        borderRadius:10,
                        fontSize:13,
                        fontWeight:600,
                        cursor:"pointer",
                        transition:"all 0.2s",
                        display:"flex",
                        alignItems:"center",
                        justifyContent:"center",
                        gap:8
                      }}
                      onMouseOver={e => e.target.style.background = "rgba(99,102,241,0.3)"}
                      onMouseOut={e => e.target.style.background = "rgba(99,102,241,0.2)"}
                    >
                      📧 Email
                    </button>
                    <button
                      onClick={() => window.open(`tel:${staff.contact}`, '_blank', 'noopener,noreferrer')}
                      style={{
                        flex:1,
                        background:"rgba(16,185,129,0.2)",
                        border:"1px solid rgba(16,185,129,0.4)",
                        color:"#d1fae5",
                        padding:"10px 14px",
                        borderRadius:10,
                        fontSize:13,
                        fontWeight:600,
                        cursor:"pointer",
                        transition:"all 0.2s",
                        display:"flex",
                        alignItems:"center",
                        justifyContent:"center",
                        gap:8
                      }}
                      onMouseOver={e => e.target.style.background = "rgba(16,185,129,0.3)"}
                      onMouseOut={e => e.target.style.background = "rgba(16,185,129,0.2)"}
                    >
                      📞 Call
                    </button>
                  </div>
                  <button
                    onClick={() => window.open(`https://wa.me/91${staff.whatsapp}`, '_blank', 'noopener,noreferrer')}
                    style={{
                      width:"100%",
                      background:"linear-gradient(135deg, rgba(34,197,94,0.2), rgba(22,163,74,0.3))",
                      border:"1px solid rgba(34,197,94,0.5)",
                      color:"#d1fae5",
                      padding:"12px 16px",
                      borderRadius:12,
                      fontSize:14,
                      fontWeight:700,
                      cursor:"pointer",
                      transition:"all 0.3s",
                      display:"flex",
                      alignItems:"center",
                      justifyContent:"center",
                      gap:10,
                      boxShadow:"0 8px 24px rgba(34,197,94,0.2)"
                    }}
                    onMouseOver={e => {
                      e.target.style.transform = "translateY(-2px)";
                      e.target.style.boxShadow = "0 12px 32px rgba(34,197,94,0.3)";
                    }}
                    onMouseOut={e => {
                      e.target.style.transform = "translateY(0)";
                      e.target.style.boxShadow = "0 8px 24px rgba(34,197,94,0.2)";
                    }}
                  >
                    💬 Chat on WhatsApp
                  </button>
                </div>
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

function StaffChatModal({staff,profiles,statuses}) {
  const [selected,setSelected]=useState(STAFF_BASE.find(s=>s.id!==staff.id)?.id || STAFF_BASE[0].id);
  const [text,setText]=useState('');
  const [messages,setMessages]=useState([]);
  const [unreadCounts,setUnreadCounts]=useState({});
  const peer=STAFF_BASE.find(s=>s.id===selected);
  const thread=[staff.id,selected].sort((a,b)=>a-b).join('-');

  useEffect(() => {
    if (!firestoreDb || !thread) return;
    const q = query(collection(firestoreDb, 'messages'), where('thread', '==', thread), orderBy('at', 'asc'));
    const unsubscribe = onSnapshot(q, (snapshot) => {
      const msgs = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
      setMessages(msgs);
      // Mark messages as read when viewing
      msgs.filter(m => m.to === staff.id && !m.read).forEach(async (m) => {
        await setDoc(doc(firestoreDb, 'messages', m.id), { ...m, read: true }, { merge: true });
      });
    });
    return unsubscribe;
  }, [thread, staff.id]);

  useEffect(() => {
    if (!firestoreDb) return;
    const allThreads = STAFF_BASE.filter(s => s.id !== staff.id).map(s => [staff.id, s.id].sort((a,b)=>a-b).join('-'));
    const unsubscribes = allThreads.map(thread => {
      const q = query(collection(firestoreDb, 'messages'), where('thread', '==', thread), where('to', '==', staff.id), where('read', '==', false));
      return onSnapshot(q, (snapshot) => {
        setUnreadCounts(prev => ({ ...prev, [thread]: snapshot.size }));
      });
    });
    return () => unsubscribes.forEach(unsub => unsub());
  }, [staff.id]);

  const send=async()=>{
    if(!text.trim() || !firestoreDb) return;
    const msg = {thread,from:staff.id,to:selected,text:text.trim(),at:Date.now(),read:false};
    await setDoc(doc(collection(firestoreDb, 'messages')), msg);
    setText('');
  };

  return <div style={{display:'grid',gridTemplateColumns:'220px 1fr',gap:14,minHeight:430}}>
    <div className="glass" style={{padding:10,overflowY:'auto'}}>
      {STAFF_BASE.filter(s=>s.id!==staff.id).map(s=>{
        const threadKey = [staff.id,s.id].sort((a,b)=>a-b).join('-');
        const unread = unreadCounts[threadKey] || 0;
        return <button key={s.id} onClick={()=>setSelected(s.id)} style={{width:'100%',display:'flex',alignItems:'center',gap:10,background:selected===s.id?'rgba(99,102,241,0.18)':'transparent',border:'none',borderRadius:10,padding:10,color:'#e2e8f0',textAlign:'left',marginBottom:6}}>
          <StaffAvatar staff={s} profiles={profiles} statuses={statuses} size={34} showStatus/>
          <div style={{flex:1}}><div style={{fontSize:13,fontWeight:700}}>{s.name}</div><StatusDot status={getStaffStatus(s.id,statuses)}/></div>
          {unread>0&&<span style={{background:'#ef4444',color:'#fff',borderRadius:999,padding:'2px 7px',fontSize:11}}>{unread}</span>}
        </button>
      })}
    </div>
    <div className="glass" style={{display:'flex',flexDirection:'column',overflow:'hidden'}}>
      <div style={{padding:14,borderBottom:'1px solid rgba(255,255,255,0.08)',display:'flex',gap:10,alignItems:'center'}}>
        <StaffAvatar staff={peer} profiles={profiles} statuses={statuses} size={38} showStatus/>
        <div><div style={{fontSize:14,fontWeight:800,color:'#fff'}}>{peer?.name}</div><StatusDot status={getStaffStatus(peer?.id,statuses)}/></div>
      </div>
      <div style={{flex:1,padding:14,overflowY:'auto'}}>
        {messages.map(m=>{
          const mine=m.from===staff.id;
          return <div key={m.id} style={{display:'flex',justifyContent:mine?'flex-end':'flex-start',marginBottom:10}}>
            <div style={{maxWidth:'72%',background:mine?'rgba(99,102,241,0.28)':'rgba(255,255,255,0.07)',border:'1px solid rgba(255,255,255,0.08)',borderRadius:14,padding:'9px 12px'}}>
              <div style={{fontSize:12,color:'rgba(226,232,240,0.45)',marginBottom:3}}>{STAFF_BASE.find(s=>s.id===m.from)?.name} · {timeAgo(m.at)}</div>
              <div style={{fontSize:13,color:'#e2e8f0',lineHeight:1.4}}>{m.text}</div>
            </div>
          </div>
        })}
        {messages.length===0&&<div style={{textAlign:'center',color:'rgba(226,232,240,0.35)',paddingTop:80}}>No messages yet</div>}
      </div>
      <div style={{padding:12,borderTop:'1px solid rgba(255,255,255,0.08)',display:'flex',gap:10}}>
        <input value={text} onChange={e=>setText(e.target.value)} onKeyDown={e=>e.key==='Enter'&&send()} placeholder="Type a message..."/>
        <button className="glow-btn" style={{padding:'10px 18px'}} onClick={send}>Send</button>
      </div>
    </div>
  </div>;
}
// ── QUICK ASSIGN DIALOG ───────────────────────────────────────────────────
function QuickAssignDialog({ticket,onClose,onSave,statuses={}}) {
  const [assigneeId,setAssigneeId]=useState(ticket?.assigneeId || STAFF_BASE[0]?.id || "");
  const [remark,setRemark]=useState("");
  const currentAssignee=STAFF_BASE.find(s=>s.id===ticket?.assigneeId);

  if(!ticket) return null;

  return (
    <div style={{display:"flex",flexDirection:"column",gap:16}}>
      <div className="glass" style={{padding:"14px 16px"}}>
        <div style={{fontSize:12,color:"rgba(226,232,240,0.45)",marginBottom:8}}>TICKET</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
          <div><div style={{fontSize:11,color:"rgba(226,232,240,0.4)"}}>Ticket ID</div><div style={{fontSize:13,fontWeight:700,color:"#e2e8f0",marginTop:2}}>{ticket.id}</div></div>
          <div><div style={{fontSize:11,color:"rgba(226,232,240,0.4)"}}>Current Assignee</div><div style={{fontSize:13,fontWeight:700,color:"#e2e8f0",marginTop:2}}>{currentAssignee?.name || "Unassigned"}</div></div>
          <div><div style={{fontSize:11,color:"rgba(226,232,240,0.4)"}}>Category</div><div style={{fontSize:13,fontWeight:600,color:"#e2e8f0",marginTop:2}}>{categoryLabel(ticket.category)}</div></div>
          <div><div style={{fontSize:11,color:"rgba(226,232,240,0.4)"}}>Priority</div><div style={{fontSize:13,fontWeight:600,color:priorityColor(ticket.priority),marginTop:2}}>{ticket.priority}</div></div>
        </div>
      </div>

      <div>
        <label style={{fontSize:12,color:"rgba(226,232,240,0.65)",marginBottom:6,display:"block",fontWeight:600}}>Assign To</label>
        <select value={assigneeId} onChange={e=>setAssigneeId(Number(e.target.value))}>
          {STAFF_BASE.map(staff=>(
            <option key={staff.id} value={staff.id}>{staff.name} ({staff.role}) - {getStaffStatus(staff.id,statuses)}</option>
          ))}
        </select>
      </div>

      <div>
        <label style={{fontSize:12,color:"rgba(226,232,240,0.65)",marginBottom:6,display:"block",fontWeight:600}}>Optional Remark</label>
        <textarea rows={4} value={remark} onChange={e=>setRemark(e.target.value)} placeholder="Add assignment context for audit trail and notifications..." style={{resize:"vertical"}} />
      </div>

      <div style={{display:"flex",gap:10,justifyContent:"flex-end",flexWrap:"wrap"}}>
        <button onClick={onClose} style={{background:"rgba(255,255,255,0.06)",border:"1px solid rgba(255,255,255,0.1)",color:"#e2e8f0",padding:"10px 18px",borderRadius:10,fontSize:14}}>Cancel</button>
        <button className="glow-btn" style={{padding:"10px 20px",fontSize:14}} onClick={()=>onSave(ticket.id,Number(assigneeId),remark)}>Save Assignment</button>
      </div>
    </div>
  );
}
// ── MAIN APP ──────────────────────────────────────────────────────────────
const getSavedSession = () => {
  if (!hasStorage()) return null;
  return safeJsonParse(localStorage.getItem("helpdesk_session"), null);
};

const getInitialPage = (sess) => {
  if (!sess) return "home";
  if (sess.type === "admin") return "dashboard";
  if (sess.type === "staff") return "staff-dash";
  return "home";
};

export default function App() {
  const { toasts, toast, remove } = useToast();
  const [session, setSession] = useState(getSavedSession);
  const [page, setPage] = useState(() => getInitialPage(getSavedSession()));
  const [tickets, setTickets] = useState([]);
  const [ticketsLoaded, setTicketsLoaded] = useState(false);
  const [feedback, setFeedback] = useState([]);
  const [feedbackLoaded, setFeedbackLoaded] = useState(false);
  const [feedbackTicketId, setFeedbackTicketId] = useState("");
  const [dismissedFeedbackTickets, setDismissedFeedbackTickets] = useState([]);
  const [viewTicketId, setViewTicketId] = useState(null);
  const [quickAssignTicketId, setQuickAssignTicketId] = useState(null);
  const [formCat, setFormCat] = useState(null);
  const [mobileOpen, setMobileOpen] = useState(false);
  const [staffProfiles, setStaffProfiles] = useState(() => DB.get("staff_profiles", {}));
  const [staffStatuses, setStaffStatuses] = useState(() => DB.get("staff_statuses", {}));
  const [staffPanel, setStaffPanel] = useState(null);
  const [staffMenuOpen, setStaffMenuOpen] = useState(false);
  useEffect(() => {
    let alive = true;
    const loadTickets = async () => {
      try {
        const onlineTickets = await fetchTickets();
        if (alive) {
          setTickets(onlineTickets);
          setTicketsLoaded(true);
        }
      } catch (error) {
        console.error("Online ticket load failed:", error);
        if (alive) setTicketsLoaded(true);
      }
    };
    loadTickets();
    const interval = setInterval(loadTickets, 15000);
    return () => { alive = false; clearInterval(interval); };
  }, []);
  useEffect(() => {
    let alive = true;
    const loadFeedback = async () => {
      try {
        const onlineFeedback = await fetchFeedback();
        if (alive) {
          setFeedback(onlineFeedback);
          setFeedbackLoaded(true);
        }
      } catch (error) {
        console.error("Online feedback load failed:", error);
        if (alive) setFeedbackLoaded(true);
      }
    };
    loadFeedback();
    const interval = setInterval(loadFeedback, 15000);
    return () => { alive = false; clearInterval(interval); };
  }, []);

  useEffect(() => {
    let alive = true;
    const loadStaffProfiles = async () => {
      try {
        const onlineProfiles = await fetchStaffProfiles();
        if (alive) {
          setStaffProfiles(onlineProfiles);
        }
      } catch (error) {
        console.error("Online staff profiles load failed:", error);
      }
    };
    if (ONLINE_TICKETS_ENABLED) {
      loadStaffProfiles();
    }
  }, []);

  useEffect(() => {
    if (!ticketsLoaded || !ONLINE_TICKETS_ENABLED) return;
    saveTickets(tickets).catch(error => console.error("Online ticket sync failed:", error));
  }, [tickets, ticketsLoaded]);
  useEffect(() => DB.set("staff_profiles", staffProfiles), [staffProfiles]);
  useEffect(() => DB.set("staff_statuses", staffStatuses), [staffStatuses]);
  useEffect(() => {
    if (typeof window === "undefined") return;
    const ticketId = new URLSearchParams(window.location.search).get("feedbackTicket");
    if (ticketId) {
      setFeedbackTicketId(ticketId);
      setPage("feedback");
    }
  }, []);

  const handleLogin = (sess) => {
    if (sess.type === "staff_firstlogin") {
      setSession(sess);
      return;
    }

    setSession(sess);
    if (hasStorage()) localStorage.setItem("helpdesk_session", JSON.stringify(sess));
    setPage(sess.type === "admin" ? "dashboard" : sess.type === "staff" ? "staff-dash" : "home");
    toast(`Welcome${sess.name ? `, ${sess.name}` : ""}!`, "success");
  };

  const logoutUser = () => {
    if (hasStorage()) localStorage.removeItem("helpdesk_session");
    setSession(null);
    setPage("home");
    setViewTicketId(null);
  };

  const reloadTickets = useCallback(async () => {
    try {
      const onlineTickets = await fetchTickets();
      setTickets(onlineTickets);
      setTicketsLoaded(true);
      return onlineTickets;
    } catch (error) {
      console.error("Online ticket refresh failed:", error);
      return tickets;
    }
  }, [tickets]);

  const reloadFeedback = useCallback(async () => {
    try {
      const onlineFeedback = await fetchFeedback();
      setFeedback(onlineFeedback);
      setFeedbackLoaded(true);
      return onlineFeedback;
    } catch (error) {
      console.error("Online feedback refresh failed:", error);
      return feedback;
    }
  }, [feedback]);
  const handleDeleteTicket = async (id) => {
    try {
      await deleteTicket(id);
      setTickets(ts => ts.filter(t => t.id !== id));
      await reloadTickets();
      toast("Ticket deleted", "info");
    } catch (error) {
      console.error("Online ticket delete failed:", error);
      toast("Ticket delete failed", "error");
    }
  };

const handleNewTicket = async (form) => {

  const assignee =
    getActiveStaffForAssignment() ||
    STAFF_BASE[0];

  const now = Date.now();

  const newTicket = {
      id: genId(),
      name: form.name,
      email: form.email,
      dept: form.dept,
      mobile: form.mobile || "",
      location: form.location || "",
      category: form.category,
      description: form.description,
      priority: form.priority || "Medium",
      status: "Assigned",
      assigneeId: assignee.id,
      assigneeName: assignee.name,
      createdAt: now,
      updatedAt: now,
      closedAt: null,
      feedbackSubmitted: false,
      closingRemarks: "",
      comments: [],
      timeline: [
        {
          action: "Created",
          by: form.name,
          at: now
        },
        {
          action: `Assigned to ${assignee.name}`,
          by: "System",
          at: now
        }
      ]
    };

    try {
      console.log("Saving ticket to Firestore:", newTicket);
      await saveTicket(newTicket);
      console.log("Firestore save success:", newTicket.id);

      setTickets(prev => [newTicket, ...prev.filter(t => t.id !== newTicket.id)]);
      setFormCat(null);
      if (session?.type === "user") setPage("my-tickets");
      reloadTickets().catch(error => console.error("Post-create ticket refresh failed:", error));

      await sendTicketEmail(newTicket, { name: newTicket.name, email: newTicket.email });
      emailTicketCreated(newTicket, assignee);

      if (assignee?.email) {
        simulateEmail(
          assignee.email,
          `[${newTicket.id}] New Ticket Assigned`,
          `${newTicket.name} submitted a ${categoryLabel(newTicket.category)} ticket.\n\n${newTicket.description}`
        );
      }

      toast("Ticket created successfully", "success");
    } catch (error) {
      console.error("Ticket create/save failed:", error);
      toast(`Ticket save failed: ${error?.message || "Firestore error"}`, "error");
    }
  };
  const handleFeedbackSubmit = async (entry) => {
    const feedbackEntry = normalizeFeedback({
      id: entry.id || genFeedbackId(),
      ticketId: entry.ticketId || "",
      name: entry.name || "",
      email: entry.email || "",
      dept: entry.dept || "",
      category: entry.category || "",
      rating: Number(entry.rating || 0),
      satisfaction: entry.satisfaction || "",
      recommend: entry.recommend || "Yes",
      message: entry.message || "",
      suggestions: entry.suggestions || "",
      createdAt: entry.createdAt || Date.now(),
      reviewed: false,
    });

    try {
      console.log("Saving feedback:", feedbackEntry);
      await saveFeedback(feedbackEntry);
      console.log("Feedback save success:", feedbackEntry.id);

      setFeedback(fs => [feedbackEntry, ...fs.filter(f => f.id !== feedbackEntry.id)]);

      if (feedbackEntry.ticketId) {
        const currentTicket = tickets.find(t => t.id === feedbackEntry.ticketId);
        const updatedTicket = currentTicket
          ? {...currentTicket, feedbackSubmitted:true, feedbackId:feedbackEntry.id, updatedAt:Date.now()}
          : null;
        if (updatedTicket) {
          await updateTicket(updatedTicket);
          setTickets(ts => ts.map(t => t.id === feedbackEntry.ticketId ? updatedTicket : t));
        }
        setDismissedFeedbackTickets(ids => Array.from(new Set([...ids, feedbackEntry.ticketId])));
      }

      simulateEmail(
        "admin@jaipuria.ac.in",
        `New IT feedback submitted by ${feedbackEntry.name}`,
        `New IT feedback submitted by ${feedbackEntry.name}\n\nFeedback ID: ${feedbackEntry.id}\nEmail: ${feedbackEntry.email}\nDepartment: ${feedbackEntry.dept}\nService: ${feedbackEntry.category}\nRating: ${feedbackEntry.rating}/5\nSatisfaction: ${feedbackEntry.satisfaction}\nRecommendation: ${feedbackEntry.recommend}\nSubmitted: ${fmtDate(feedbackEntry.createdAt)}\n\nFeedback:\n${feedbackEntry.message}\n\nSuggestions:\n${feedbackEntry.suggestions || "—"}`
      );

      reloadFeedback().catch(error => console.error("Post-submit feedback refresh failed:", error));
      toast("Thank you for your feedback", "success");
    } catch (error) {
      console.error("Feedback save failed:", error);
      toast(`Feedback save failed: ${error?.message || "Firestore error"}`, "error");
      throw error;
    }
  };
  const handleQuickAssign = (ticketId, assigneeId, remark="") => {
    let assignedTicket = null;
    let newAssignee = null;
    let previousAssignee = null;

    setTickets(ts => ts.map(t => {
      if (t.id !== ticketId) return t;
      newAssignee = STAFF_BASE.find(s => s.id === assigneeId);
      previousAssignee = STAFF_BASE.find(s => s.id === t.assigneeId);
      const cleanRemark = remark.trim();
      const timelineEntry = {
        action: `Quick assigned to ${newAssignee?.name || "Unassigned"}`,
        remark: cleanRemark,
        at: Date.now(),
        by: "Admin",
      };
      const comments = cleanRemark
        ? [...(t.comments || []), { text: cleanRemark, at: Date.now(), by: "Admin" }]
        : (t.comments || []);
      assignedTicket = {
        ...t,
        assigneeId,
        status: t.status === "Open" ? "Assigned" : t.status,
        updatedAt: Date.now(),
        comments,
        timeline: [...(t.timeline || []), timelineEntry],
      };
      return assignedTicket;
    }));

    if (assignedTicket) {
      emailTicketAssigned(assignedTicket, newAssignee, previousAssignee, "Admin", remark.trim());
      toast(`Ticket assigned to ${newAssignee?.name || "staff"}`,"success");
    }
    setQuickAssignTicketId(null);
  };

  const updateStaffProfile = async (staffId, changes) => {
    setStaffProfiles(p => ({...p, [staffId]: {...(p[staffId] || {}), ...changes}}));
    if (ONLINE_TICKETS_ENABLED) {
      await saveStaffProfile(staffId, {...(staffProfiles[staffId] || {}), ...changes});
    }
  };

  const updateOwnStatus = (status) => {
    if (!session?.staffId) return;
    setStaffStatuses(s => ({...s, [session.staffId]: status}));
  };

  const handleFirstLoginComplete = (hash) => {
    const staffPasswords = DB.get("staff_passwords", {});
    staffPasswords[session.staffId] = hash;
    DB.set("staff_passwords", staffPasswords);

    const staff = STAFF_BASE.find(s => s.id === session.staffId);
    const staffSession = {
      type: "staff",
      staffId: staff.id,
      email: staff.email,
      name: staff.name,
      role: staff.role,
      permissions: staff.permissions,
    };

    setSession(staffSession);
    if (hasStorage()) localStorage.setItem("helpdesk_session", JSON.stringify(staffSession));
    setPage("staff-dash");
  };

  if (!session) {
  return (
    <>
      <style>{CSS}</style>
      <Landing onLogin={handleLogin} tickets={tickets} />
      <Toast toasts={toasts} remove={remove} />
    </>
  );
}

  if (session.type === "staff_firstlogin") {
    return (
      <>
        <style>{CSS}</style>
        <SetPasswordScreen staff={session.staff} onComplete={handleFirstLoginComplete} toast={toast} />
        <Toast toasts={toasts} remove={remove} />
      </>
    );
  }

  const isAdmin = session.type === "admin";
  const isStaff = session.type === "staff";
  const myTickets = isAdmin || isStaff ? tickets : tickets.filter(t => t.email === session.email);
  const quickAssignTicket = tickets.find(t => t.id === quickAssignTicketId);
  const linkedFeedbackTicket = tickets.find(t => t.id === feedbackTicketId);
  const pendingFeedbackTicket = !isAdmin && !isStaff ? tickets.find(t => t.email === session.email && t.status === "Closed" && !t.feedbackSubmitted && !feedback.some(f => f.ticketId === t.id && f.email === session.email) && !dismissedFeedbackTickets.includes(t.id)) : null;

  const renderStaffManagement = () => (
    <div style={{display:"flex",flexDirection:"column",gap:18}}>
      <h2 style={{fontFamily:"Syne",fontSize:22,fontWeight:700,color:"#e2e8f0"}}>Staff Management</h2>
      {STAFF_BASE.map(staff => (
        <div key={staff.id} className="glass" style={{padding:"18px 20px"}}>
          <div style={{display:"flex",gap:12,alignItems:"center",marginBottom:12}}>
            <StaffAvatar staff={staff} profiles={staffProfiles} statuses={staffStatuses} size={46} showStatus />
            <div>
              <div style={{fontSize:18,fontWeight:700,color:"#fff"}}>{staff.name}</div>
              <div style={{fontSize:13,color:"rgba(226,232,240,0.5)",marginTop:4}}>{staff.email}</div>
              <StatusDot status={getStaffStatus(staff.id,staffStatuses)} />
            </div>
          </div>
          <button
            onClick={async () => {
              const newPwd = prompt("Enter new password");
              if (!newPwd) return;

              const passwords = DB.get("staff_passwords", {});
              passwords[staff.id] = await hashPassword(newPwd);
              DB.set("staff_passwords", passwords);
              toast("Password updated", "success");
            }}
            style={{marginTop:14,padding:"10px 16px",border:"none",borderRadius:10,background:"#2563eb",color:"#fff",cursor:"pointer"}}
          >
            Change Password
          </button>
        </div>
      ))}
    </div>
  );

  const renderPage = () => {
    if (isAdmin) {
      if (page === "dashboard") return (
        <div style={{display:"flex",flexDirection:"column",gap:24}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",flexWrap:"wrap",gap:12}}>
            <div>
              <h2 style={{fontFamily:"Syne",fontSize:22,fontWeight:700,color:"#e2e8f0"}}>Admin Dashboard</h2>
              <p style={{fontSize:14,color:"rgba(226,232,240,0.5)"}}>{new Date().toLocaleDateString("en-IN",{weekday:"long",year:"numeric",month:"long",day:"numeric"})}</p>
            </div>
            <button className="glow-btn" onClick={() => setFormCat("")}>+ New Ticket</button>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(155px,1fr))",gap:14}}>
            {[["Total",tickets.length,"🎫","#818cf8"],["Open",tickets.filter(t=>t.status==="Open").length,"🔵","#60a5fa"],["In Progress",tickets.filter(t=>t.status==="In Progress").length,"🟡","#fbbf24"],["Resolved",tickets.filter(t=>t.status==="Resolved"||t.status==="Closed").length,"🟢","#34d399"],["Critical",tickets.filter(t=>t.priority==="Critical").length,"🔴","#f87171"],["Closed",tickets.filter(t=>t.status==="Closed").length,"⚫","#6b7280"]].map(([l,v,i,c]) => <StatCard key={l} label={l} value={v} icon={i} color={c} />)}
          </div>
          <h3 style={{fontFamily:"Syne",fontSize:16,fontWeight:700,color:"#e2e8f0"}}>Recent Tickets</h3>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(280px,1fr))",gap:12}}>
            {tickets.slice(0,6).map(t => (
              <div key={t.id} style={{position:"relative"}}>
                <TicketCard ticket={t} onView={setViewTicketId} />
                <button
                  onClick={e=>{e.stopPropagation();setQuickAssignTicketId(t.id);}}
                  style={{position:"absolute",right:12,bottom:12,background:"rgba(99,102,241,0.92)",border:"1px solid rgba(255,255,255,0.18)",color:"#fff",padding:"7px 12px",borderRadius:8,fontSize:12,fontWeight:700,boxShadow:"0 10px 24px rgba(99,102,241,0.25)"}}
                >
                  Assign
                </button>
              </div>
            ))}
            {tickets.length === 0 && <div style={{gridColumn:"1/-1",textAlign:"center",padding:"60px 0",color:"rgba(226,232,240,0.3)"}}><div style={{fontSize:48,marginBottom:12}}>🎫</div><div>No tickets yet</div></div>}
          </div>
        </div>
      );
      if (page === "tickets") return <TicketsTable tickets={tickets} onView={setViewTicketId} isAdmin onDelete={handleDeleteTicket} />;
      if (page === "analytics") return <Analytics tickets={tickets} />;
      if (page === "feedback") return <AdminFeedbackPage feedback={feedback} setFeedback={setFeedback} toast={toast} />;
      if (page === "staff-management") return renderStaffManagement();
      if (page === "export") return <ExportPanel tickets={tickets} toast={toast} />;
      if (page === "emaillog") return <EmailLog />;
      if (page === "staff") return (
        <div>
          <h2 style={{fontFamily:"Syne",fontSize:22,fontWeight:700,color:"#e2e8f0",marginBottom:20}}>IT Staff Management</h2>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(280px,1fr))",gap:16}}>
            {STAFF_BASE.map(s => {
              const asgn = tickets.filter(t => t.assigneeId === s.id).length;
              const res = tickets.filter(t => t.assigneeId === s.id && (t.status === "Resolved" || t.status === "Closed")).length;
              const active = asgn - res;
              return (
                <div key={s.id} className="glass" style={{padding:"22px"}}>
                  <div style={{display:"flex",gap:14,alignItems:"center",marginBottom:16}}>
                    <StaffAvatar staff={s} profiles={staffProfiles} statuses={staffStatuses} size={50} showStatus />
                    <div><div style={{fontSize:15,fontWeight:700,color:"#e2e8f0"}}>{s.name}</div><div style={{fontSize:12,color:"rgba(226,232,240,0.5)"}}>{s.role}</div><div style={{fontSize:11,color:"rgba(226,232,240,0.35)"}}>{s.email}</div><StatusDot status={getStaffStatus(s.id,staffStatuses)} /></div>
                  </div>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10,marginBottom:12}}>
                    {[["Assigned",asgn,"#818cf8"],["Active",active,"#fbbf24"],["Resolved",res,"#34d399"]].map(([l,v,c]) => (
                      <div key={l} style={{textAlign:"center",background:`${c}15`,borderRadius:8,padding:"10px 4px"}}><div style={{fontSize:20,fontWeight:800,fontFamily:"Syne",color:c}}>{v}</div><div style={{fontSize:11,color:"rgba(226,232,240,0.5)",marginTop:2}}>{l}</div></div>
                    ))}
                  </div>
                  <div style={{fontSize:12,color:"rgba(226,232,240,0.4)",marginBottom:6}}>Permissions</div>
                  <div style={{display:"flex",flexWrap:"wrap",gap:4}}>{s.permissions.map(p => <span key={p} className="tag" style={{background:"rgba(99,102,241,0.1)",color:"#818cf8",fontSize:10}}>{p}</span>)}</div>
                </div>
              );
            })}
          </div>
        </div>
      );
    }

    if (isStaff && (page === "staff-dash" || page === "assigned")) {
      return <StaffPanel staffId={session.staffId} tickets={tickets} setTickets={setTickets} toast={toast} onViewTicket={setViewTicketId} permissions={session.permissions} staffProfiles={staffProfiles} staffStatuses={staffStatuses} />;
    }

    if (page === "home") return <CategoryGrid onSelect={cat => setFormCat(cat)} />;
    if (page === "my-tickets") return (
      <div>
        <h2 style={{fontFamily:"Syne",fontSize:22,fontWeight:700,color:"#e2e8f0",marginBottom:20}}>My Tickets</h2>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(280px,1fr))",gap:12}}>
          {myTickets.map(t => <TicketCard key={t.id} ticket={t} onView={setViewTicketId} />)}
          {myTickets.length === 0 && <div style={{gridColumn:"1/-1",textAlign:"center",padding:"60px 0",color:"rgba(226,232,240,0.3)"}}><div style={{fontSize:48,marginBottom:12}}>🎫</div><div>No tickets yet</div><button className="glow-btn" style={{marginTop:16}} onClick={() => setPage("home")}>Raise Ticket</button></div>}
        </div>
      </div>
    );
    if (page === "know-staff") return <KnowYourITStaff staffProfiles={staffProfiles} />;
    if (page === "feedback") return <FeedbackForm userEmail={session.email} onSubmit={handleFeedbackSubmit} toast={toast} ticket={linkedFeedbackTicket || null} />;
    if (page === "track") return <TrackTicket tickets={tickets} onView={setViewTicketId} />;
    if (page === "new-ticket") return (
      <div>
        <h2 style={{fontFamily:"Syne",fontSize:22,fontWeight:700,color:"#e2e8f0",marginBottom:20}}>New Ticket</h2>
        <div className="glass" style={{padding:"24px"}}>
          <TicketForm userEmail={session.email} initialCategory="" onSubmit={async t => { await handleNewTicket(t); setPage("my-tickets"); }} onCancel={() => setPage("home")} toast={toast} />
        </div>
      </div>
    );

    return null;
  };

  return (
    <>
      <style>{CSS}</style>
      <div className="app-shell" style={{display:"flex",minHeight:"100vh"}}>
        <Sidebar current={page} onChange={setPage} isAdmin={isAdmin} isStaff={isStaff} tickets={tickets} feedback={feedback} mobileOpen={mobileOpen} setMobileOpen={setMobileOpen} />
        <div className="app-main" style={{flex:1,display:"flex",flexDirection:"column",minWidth:0}}>
          <div className="app-header" style={{padding:"14px 24px",borderBottom:"1px solid rgba(255,255,255,0.07)",display:"flex",justifyContent:"space-between",alignItems:"center",background:"rgba(10,10,20,0.9)",backdropFilter:"blur(20px)",position:"sticky",top:0,zIndex:10}}>
            <div className="header-identity" style={{display:"flex",alignItems:"center",gap:12,position:"relative"}}>
              <button onClick={() => { const isMobile = typeof window !== "undefined" && window.innerWidth <= 768; if (isMobile) setMobileOpen(o => !o); else if (isStaff) setStaffMenuOpen(o=>!o); else setMobileOpen(o => !o); }} style={{background:"rgba(255,255,255,0.06)",border:"1px solid rgba(255,255,255,0.1)",color:"#e2e8f0",width:38,height:38,borderRadius:10,fontSize:18,display:"flex",alignItems:"center",justifyContent:"center"}}>☰</button>
              {isStaff&&staffMenuOpen&&(
                <StaffProfileMenu
                  staff={STAFF_BASE.find(s=>s.id===session.staffId)}
                  profiles={staffProfiles}
                  statuses={staffStatuses}
                  onStatusChange={updateOwnStatus}
                  onOpen={setStaffPanel}
                  onLogout={logoutUser}
                  onClose={()=>setStaffMenuOpen(false)}
                />
              )}
              {!isStaff&&<span style={{fontSize:13,color:"rgba(226,232,240,0.4)"}}>{isAdmin ? "Admin Portal" : session.email}</span>}
            </div>
            <div className="header-actions" style={{display:"flex",gap:10,alignItems:"center"}}>
              <div className="pulse" style={{width:8,height:8,borderRadius:"50%",background:"#10b981"}} />
              <span style={{fontSize:12,color:"rgba(226,232,240,0.4)"}}>Live</span>
              {!isStaff&&<button onClick={logoutUser} style={{background:"rgba(239,68,68,0.1)",border:"1px solid rgba(239,68,68,0.3)",color:"#f87171",padding:"6px 14px",borderRadius:8,fontSize:13}}>Logout</button>}
            </div>
          </div>
          <div className="app-content" style={{padding:"24px 28px",flex:1,overflowY:"auto"}}>{renderPage()}</div>
        </div>
      </div>

      {formCat !== null && (
        <Modal title="Raise IT Support Ticket" onClose={() => setFormCat(null)}>
          <TicketForm userEmail={session?.email} initialCategory={formCat} onSubmit={handleNewTicket} onCancel={() => setFormCat(null)} toast={toast} />
        </Modal>
      )}

      {quickAssignTicket && (
        <Modal title={`Assign Ticket - ${quickAssignTicket.id}`} onClose={() => setQuickAssignTicketId(null)}>
          <QuickAssignDialog
            ticket={quickAssignTicket}
            onClose={() => setQuickAssignTicketId(null)}
            onSave={handleQuickAssign}
            statuses={staffStatuses}
          />
        </Modal>
      )}


      {pendingFeedbackTicket && page !== "feedback" && (
        <Modal title="Share IT Support Feedback" onClose={() => setDismissedFeedbackTickets(ids => Array.from(new Set([...ids, pendingFeedbackTicket.id])))}>
          <div style={{display:"flex",flexDirection:"column",gap:18}}>
            <div className="glass" style={{padding:"18px 20px",background:"rgba(99,102,241,0.1)",borderColor:"rgba(99,102,241,0.3)"}}>
              <div style={{fontSize:18,fontWeight:800,fontFamily:"Syne",color:"#e2e8f0",marginBottom:8}}>Your ticket {pendingFeedbackTicket.id} has been closed.</div>
              <div style={{fontSize:14,lineHeight:1.6,color:"rgba(226,232,240,0.68)"}}>Please share your feedback so the IT team can improve support quality.</div>
            </div>
            <div style={{display:"flex",justifyContent:"flex-end",gap:10,flexWrap:"wrap"}}>
              <button onClick={() => setDismissedFeedbackTickets(ids => Array.from(new Set([...ids, pendingFeedbackTicket.id])))} style={{background:"rgba(255,255,255,0.06)",border:"1px solid rgba(255,255,255,0.1)",color:"#e2e8f0",padding:"10px 18px",borderRadius:10,fontSize:14}}>Later</button>
              <button className="glow-btn" onClick={() => { setFeedbackTicketId(pendingFeedbackTicket.id); setPage("feedback"); }}>Give Feedback</button>
            </div>
          </div>
        </Modal>
      )}
      {viewTicketId && (
        <Modal title={`Ticket - ${viewTicketId}`} onClose={() => setViewTicketId(null)}>
          <TicketDetail
            ticketId={viewTicketId}
            tickets={tickets}
            setTickets={setTickets}
            isAdmin={isAdmin}
            isStaff={isStaff}
            staffId={session.staffId}
            staffName={session.name}
            toast={toast}
            staffProfiles={staffProfiles}
            staffStatuses={staffStatuses}
          />
        </Modal>
      )}

      {isStaff&&staffPanel==="profile"&&(
        <Modal title="My Profile" onClose={()=>setStaffPanel(null)}>
          <StaffProfileModal staff={STAFF_BASE.find(s=>s.id===session.staffId)} profiles={staffProfiles} statuses={staffStatuses} onSave={updateStaffProfile} toast={toast}/>
        </Modal>
      )}
      {isStaff&&staffPanel==="password"&&(
        <Modal title="Change Password" onClose={()=>setStaffPanel(null)}>
          <StaffChangePasswordModal staff={STAFF_BASE.find(s=>s.id===session.staffId)} toast={toast}/>
        </Modal>
      )}
      {isStaff&&staffPanel==="performance"&&(
        <Modal title="My Performance" onClose={()=>setStaffPanel(null)} wide>
          <StaffPerformanceModal staff={STAFF_BASE.find(s=>s.id===session.staffId)} tickets={tickets}/>
        </Modal>
      )}
      {isStaff&&staffPanel==="chat"&&(
        <Modal title="Staff Chat" onClose={()=>setStaffPanel(null)} wide>
          <StaffChatModal staff={STAFF_BASE.find(s=>s.id===session.staffId)} profiles={staffProfiles} statuses={staffStatuses} />
        </Modal>
      )}

      <Toast toasts={toasts} remove={remove} />
    </>
  );
}














































































