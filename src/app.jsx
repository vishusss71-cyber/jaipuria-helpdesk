import * as XLSX from "xlsx";
import { saveAs } from "file-saver";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import emailjs from '@emailjs/browser';
import { motion, AnimatePresence } from 'framer-motion';
import { initializeApp } from 'firebase/app';
import { getFirestore, collection, doc, setDoc, getDocs, onSnapshot, query, where, orderBy, addDoc } from 'firebase/firestore';
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
    watchers: Array.isArray(ticket.watchers) ? ticket.watchers : [],
    notifiedStaff: Array.isArray(ticket.notifiedStaff) ? ticket.notifiedStaff : [],
    feedbackSubmitted: Boolean(ticket.feedbackSubmitted),
    userConfirmedResolved: Boolean(ticket.userConfirmedResolved),
    userFeedbackStatus: ticket.userFeedbackStatus || "",
    userReviewedAt: ticket.userReviewedAt || null,
    reopenReason: ticket.reopenReason || "",
    reopenedAt: ticket.reopenedAt || null,
    reopenedBy: ticket.reopenedBy || "",
    closedAt: ticket.closedAt ?? null,
    closingRemarks: ticket.closingRemarks || "",
    userFeedbackRating: Number(ticket.userFeedbackRating || 0),
    userFeedbackComment: ticket.userFeedbackComment || "",
    userFeedbackAt: ticket.userFeedbackAt || null,
    feedbackStatus: ticket.feedbackStatus || "",
    feedbackReadByAdmin: Boolean(ticket.feedbackReadByAdmin),
    feedbackReadByStaff: Boolean(ticket.feedbackReadByStaff),
    remoteSupportRequested: Boolean(ticket.remoteSupportRequested),
    remoteSupportTool: ticket.remoteSupportTool || "",
    remoteSupportId: ticket.remoteSupportId || "",
    remoteSupportNote: ticket.remoteSupportNote || "",
    remoteSupportRequestedAt: ticket.remoteSupportRequestedAt || null,
    aiSummary: ticket.aiSummary || "",
    suggestedAction: ticket.suggestedAction || "",
    escalated: Boolean(ticket.escalated),
    escalationLevel: Number(ticket.escalationLevel || 0),
    escalatedAt: ticket.escalatedAt || null,
    escalationHistory: Array.isArray(ticket.escalationHistory) ? ticket.escalationHistory : [],
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

function genIncidentId() { return "INC-"+Date.now().toString(36).toUpperCase()+Math.random().toString(36).slice(2,5).toUpperCase(); }
function normalizeIncident(entry = {}) {
  return {
    ...entry,
    id: entry.id || genIncidentId(),
    title: entry.title || "",
    message: entry.message || "",
    severity: ["Info","Warning","Critical"].includes(entry.severity) ? entry.severity : "Info",
    active: entry.active !== false,
    expiryAt: entry.expiryAt || null,
    createdAt: entry.createdAt || Date.now(),
    updatedAt: entry.updatedAt || Date.now(),
  };
}
async function fetchIncidents() {
  if (!ONLINE_TICKETS_ENABLED) return [];
  const res = await fetch(`${FIRESTORE_BASE_URL}/incidents?key=${encodeURIComponent(FIREBASE_API_KEY)}`);
  if (!res.ok) throw new Error(`Firestore incidents ${res.status}: ${getFirestoreErrorMessage(await res.text())}`);
  const data = await res.json();
  return (data.documents || [])
    .map(doc => normalizeIncident({ id: doc.name?.split("/").pop(), ...fromFirestoreFields(doc.fields || {}) }))
    .sort((a, b) => (b.updatedAt || 0) - (a.updatedAt || 0));
}
async function saveIncident(entry) {
  if (!ONLINE_TICKETS_ENABLED) throw new Error("Firestore is not configured. Check Vercel env vars and redeploy.");
  const incident = normalizeIncident(entry);
  const res = await fetch(`${FIRESTORE_BASE_URL}/incidents/${encodeURIComponent(incident.id)}?key=${encodeURIComponent(FIREBASE_API_KEY)}`, {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ fields: toFirestoreFields(incident) }),
  });
  if (!res.ok) throw new Error(`Firestore incident save ${res.status}: ${getFirestoreErrorMessage(await res.text())}`);
  return incident;
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

function genPortalFeedbackId() {
  return "PFB-" + Date.now().toString(36).toUpperCase() + Math.random().toString(36).slice(2,5).toUpperCase();
}

function normalizePortalFeedback(entry = {}) {
  return {
    ...entry,
    id: entry.id || genPortalFeedbackId(),
    name: entry.name || "",
    email: entry.email || "",
    role: entry.role || "User",
    rating: Number(entry.rating || 0),
    feedbackType: entry.feedbackType || "General Feedback",
    message: entry.message || "",
    createdAt: entry.createdAt || Date.now(),
    status: entry.status || "New",
    reviewed: Boolean(entry.reviewed),
  };
}

async function fetchPortalFeedback() {
  if (!ONLINE_TICKETS_ENABLED) {
    console.warn("Firestore portal feedback storage is not configured.");
    return [];
  }
  const res = await fetch(`${FIRESTORE_BASE_URL}/portalFeedback?key=${encodeURIComponent(FIREBASE_API_KEY)}`);
  if (!res.ok) throw new Error(`Firestore portalFeedback ${res.status}: ${getFirestoreErrorMessage(await res.text())}`);
  const data = await res.json();
  return (data.documents || [])
    .map(doc => normalizePortalFeedback({ id: doc.name?.split("/").pop(), ...fromFirestoreFields(doc.fields || {}) }))
    .sort((a, b) => (b.createdAt || 0) - (a.createdAt || 0));
}

async function savePortalFeedback(entry) {
  if (!ONLINE_TICKETS_ENABLED) {
    throw new Error("Firestore is not configured. Check Vercel env vars and redeploy.");
  }
  const cleanEntry = normalizePortalFeedback(entry);
  const res = await fetch(`${FIRESTORE_BASE_URL}/portalFeedback/${encodeURIComponent(cleanEntry.id)}?key=${encodeURIComponent(FIREBASE_API_KEY)}`, {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ fields: toFirestoreFields(cleanEntry) }),
  });
  if (!res.ok) {
    const body = await res.text();
    console.error("Firestore portal feedback save response:", res.status, body);
    throw new Error(`Firestore portalFeedback ${res.status}: ${getFirestoreErrorMessage(body)}`);
  }
  return cleanEntry;
}

async function updatePortalFeedback(entry) {
  return savePortalFeedback(entry);
}
function normalizeTempIssue(issue = {}) {
  return {
    requestId:
      issue.requestId ||
      `TI-${Date.now().toString(36).toUpperCase()}-${Math.random()
        .toString(36)
        .slice(2, 5)
        .toUpperCase()}`,

    // USER INFO
    userId: issue.userId || "",
    userName: issue.userName || "",
    userEmail: issue.userEmail || "",
    mobile: issue.mobile || "",

    // ITEM INFO
    item: issue.item || "",
    customItem: issue.customItem || "",

    // APPROVAL INFO
    permissionApprovedBy: issue.permissionApprovedBy || issue.permissionBy || "",
    permissionBy: issue.permissionBy || issue.permissionApprovedBy || "",
    requestedToStaff: issue.requestedToStaff || issue.requestToStaff || "",
    requestToStaff: issue.requestToStaff || issue.requestedToStaff || "",
    requestedToStaffEmail: issue.requestedToStaffEmail || issue.requestToStaffEmail || "",
    requestToStaffEmail: issue.requestToStaffEmail || issue.requestedToStaffEmail || "",

    // DATES
    issueDate: issue.issueDate || "",
    returnDate: issue.returnDate || "",

    // STATUS FLOW
    status: issue.status || "Pending Approval",

    // REMARKS
    remarks: issue.remarks || "",
    purpose: issue.purpose || "",

    // TIMESTAMPS
    createdAt: issue.createdAt || Date.now(),
    updatedAt: issue.updatedAt || Date.now(),

    // APPROVAL ACTIONS
    approvedBy: issue.approvedBy || "",
    approvedAt: issue.approvedAt || null,

    rejectedBy: issue.rejectedBy || "",
    rejectedAt: issue.rejectedAt || null,

    // ISSUE ACTIONS
    issuedBy: issue.issuedBy || "",
    issuedAt: issue.issuedAt || null,

    notIssuedBy: issue.notIssuedBy || "",
    notIssuedAt: issue.notIssuedAt || null,

    // RETURN FLOW
    returnRequestedAt: issue.returnRequestedAt || null,

    returnAcceptedBy: issue.returnAcceptedBy || "",
    returnAcceptedAt: issue.returnAcceptedAt || null,

    returnRejectedBy: issue.returnRejectedBy || "",
    returnRejectedAt: issue.returnRejectedAt || null,

    // FINAL RETURN
    returnedBy: issue.returnedBy || "",
    returnedAt: issue.returnedAt || null,

    // ADMIN CONTROL
    forceClosedBy: issue.forceClosedBy || "",
    forceClosedAt: issue.forceClosedAt || null,

    // TRACKING
    requestHistory: Array.isArray(issue.requestHistory)
      ? issue.requestHistory
      : [],

    notifications: Array.isArray(issue.notifications)
      ? issue.notifications
      : [],

    // LEGACY SUPPORT
    requestedBy: issue.requestedBy || "",
  };
}

async function fetchTempIssues() {
  if (!ONLINE_TICKETS_ENABLED) {
    console.warn("Firestore temp issue storage is not configured.");
    return [];
  }
  const res = await fetch(`${FIRESTORE_BASE_URL}/tempIssues?key=${encodeURIComponent(FIREBASE_API_KEY)}`);
  if (!res.ok) throw new Error(`Firestore tempIssues ${res.status}: ${getFirestoreErrorMessage(await res.text())}`);
  const data = await res.json();
  return (data.documents || [])
    .map(doc => normalizeTempIssue({ requestId: doc.name?.split("/").pop(), ...fromFirestoreFields(doc.fields || {}) }))
    .sort((a, b) => (b.createdAt || 0) - (a.createdAt || 0));
}

async function saveTempIssue(issue) {
  if (!ONLINE_TICKETS_ENABLED) {
    throw new Error("Firestore is not configured. Check Vercel env vars and redeploy.");
  }
  const cleanIssue = normalizeTempIssue(issue);
  const res = await fetch(`${FIRESTORE_BASE_URL}/tempIssues/${encodeURIComponent(cleanIssue.requestId)}?key=${encodeURIComponent(FIREBASE_API_KEY)}`, {
    method: "PATCH",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ fields: toFirestoreFields(cleanIssue) }),
  });
  if (!res.ok) {
    const body = await res.text();
    console.error("Firestore temp issue save response:", res.status, body);
    throw new Error(`Firestore tempIssues ${res.status}: ${getFirestoreErrorMessage(body)}`);
  }
  return cleanIssue;
}

async function updateTempIssue(issue) {
  return saveTempIssue(issue);
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

async function fetchStaffProfile(staffId) {
  if (!staffId) return null;
  if (!ONLINE_TICKETS_ENABLED || !firestoreDb) return null;
  const profiles = await fetchStaffProfiles();
  return profiles[String(staffId)] || null;
}

function staffPasswordExists(profile) {
  if (!profile) return false;
  if (profile.requiresPasswordSetup === true) return false;
  if (profile.passwordSet === true) return true;
  if (profile.passwordHash || profile.password) return true;
  return false;
}

function clearStaffPasswordSetupStorage() {
  if (!hasStorage()) return;
  ["helpdesk_firstLogin", "helpdesk_requiresPasswordSetup", "helpdesk_passwordSetup", "firstLogin", "requiresPasswordSetup", "passwordSetup"].forEach(key => {
    try { localStorage.removeItem(key); sessionStorage.removeItem(key); } catch {}
  });
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
  { id: 1, name: "Raj Prakash Singh", role: "Manager", email: "raj.singh@jaipuria.ac.in", avatar: "RPS", color: "#6366f1", permissions: ["view_all","assign","close","export","manage_users"] },
  { id: 2, name: "Rohit Jangid", role: "Executive", email: "rohit.jangid@jaipuria.ac.in", avatar: "RJ", color: "#0ea5e9", permissions: ["view_assigned","close","comment"] },
  { id: 3, name: "Vishal Swami", role: "Senior Executive", email: "vishal.swami@jaipuria.ac.in", avatar: "VS", color: "#10b981", permissions: ["view_assigned","assign","close","comment"] },
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
  { id:"erp", label:"LMS/Moodle Issue", icon:"📊", color:"#84cc16", bg:"rgba(132,204,22,0.15)" },
  { id:"echo360", label:"Echo360 Lecture Capture System", icon:"🎥", color:"#22d3ee", bg:"rgba(34,211,238,0.15)" },
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
function isClosedTicket(ticket) { return ["Closed","Resolved"].includes(ticket?.status); }
function hasTicketFeedback(ticket) { return Boolean(ticket?.feedbackSubmitted || ticket?.feedbackStatus === "Submitted" || ticket?.userFeedbackAt); }
function isTicketFeedbackPending(ticket) { return isClosedTicket(ticket) && !hasTicketFeedback(ticket); }
function isTicketFeedbackUnread(ticket, isAdmin=false, isStaff=false) {
  if(ticket?.feedbackStatus !== "Submitted") return false;
  if(isAdmin) return !ticket.feedbackReadByAdmin;
  if(isStaff) return !ticket.feedbackReadByStaff;
  return false;
}
function getEscalationInfo(ticket, now = Date.now()) {
  if(!ticket || isClosedTicket(ticket)) return { overdue:false, level:0, label:"" };
  const thresholds = { High:4, Critical:4, Medium:24, Low:48 };
  const hours = thresholds[ticket.priority] || 24;
  const createdAt = Number(ticket.createdAt || now);
  const elapsed = Math.max(0, now - createdAt);
  const thresholdMs = hours * 3600000;
  if(elapsed <= thresholdMs) return { overdue:false, level:0, label:"" };
  const level = elapsed > thresholdMs * 3 ? 3 : elapsed > thresholdMs * 2 ? 2 : 1;
  const label = level === 3 ? "Overdue L3 - Admin attention" : level === 2 ? "Overdue L2 - All IT staff" : "Overdue L1 - Staff reminder";
  return { overdue:true, level, label, overdueBy: elapsed - thresholdMs, thresholdHours: hours };
}
function getAiTicketFallback(form = {}) {
  const source = String(form.issueSummary || form.description || form.subCategory || "IT support request").trim();
  return {
    aiSummary: source.length > 180 ? `${source.slice(0, 177)}...` : source,
    suggestedAction: form.recommendedAction || "Review the issue details and proceed with standard troubleshooting."
  };
}
async function generateTicketAiSummary(form = {}) {
  const fallback = getAiTicketFallback(form);
  try {
    const prompt = `Create a concise IT helpdesk staff summary. Return JSON only with keys aiSummary and suggestedAction.\nCategory: ${categoryLabel(form.category)}\nSub-category: ${form.subCategory || "Not provided"}\nPriority: ${form.priority || "Medium"}\nIssue: ${form.description || form.issueSummary || ""}`;
    const response = await fetch("/api/chat", {
      method:"POST",
      headers:{"Content-Type":"application/json"},
      cache:"no-store",
      body:JSON.stringify({message:prompt,user:{name:form.name || "Portal User",email:form.email || ""}})
    });
    if(!response.ok) throw new Error(`AI summary endpoint failed: ${response.status}`);
    const data = await response.json().catch(()=>({reply:""}));
    const reply = String(data?.reply || "");
    const jsonText = reply.match(/\{[\s\S]*\}/)?.[0] || "";
    const parsed = jsonText ? JSON.parse(jsonText) : {};
    return {
      aiSummary: String(parsed.aiSummary || parsed.summary || fallback.aiSummary).slice(0, 240),
      suggestedAction: String(parsed.suggestedAction || parsed.recommendedAction || fallback.suggestedAction).slice(0, 320)
    };
  } catch(error) {
    console.error("AI ticket summary failed:", error);
    return fallback;
  }
}
function getDisplayName(session) {
  return session?.name || session?.staff?.name || session?.email?.split("@")[0]?.replace(/[._-]+/g," ") || "there";
}
function getGreetingLabel() {
  const hour=new Date().getHours();
  if(hour<12) return "Good Morning";
  if(hour<17) return "Good Afternoon";
  return "Good Evening";
}
function showBrowserNotification(title, body) {
  try {
    if(typeof window==="undefined" || !("Notification" in window)) return false;
    if(Notification.permission!=="granted") return false;
    new Notification(title, { body, icon:"/pwa-icon-192.png", badge:"/pwa-icon-192.png" });
    return true;
  } catch(error) {
    console.error("Browser notification failed:", error);
    return false;
  }
}
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

function emailAllITStaff(subject, body) {
  STAFF_BASE.forEach(staff => {
    if(staff.email) simulateEmail(staff.email, subject, body);
  });
  simulateEmail("admin@jaipuria.ac.in", subject, body);
}

function notifyTicketCreated(ticket) {
  const body=`New IT Ticket Created
Ticket ID: ${ticket.id}
Category: ${categoryLabel(ticket.category)}
User: ${ticket.name}
Mobile: ${ticket.mobile || "Not provided"}
  Issue: ${ticket.description || ticket.issueSummary || "Not provided"}`;
  emailAllITStaff(`[${ticket.id}] New IT Ticket Created`, body);
}

function notifyTicketClosed(ticket, closedBy="IT Support", resolution="") {
  const body=`Ticket Closed
Ticket ID: ${ticket.id}
Closed By: ${closedBy}
  Resolution: ${resolution || ticket.closingRemarks || "Resolved by IT Support"}`;
  if(ticket.email) simulateEmail(ticket.email, `[${ticket.id}] Ticket Closed`, `Dear ${ticket.name},\n\nYour ticket has been closed. Please review the resolution.\n\n${body}\n\nRegards,\nJaipuria IT Support`);
  emailAllITStaff(`[${ticket.id}] Ticket Closed`, body);
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

function downloadTicketCsv(tickets = [], filename = "it-helpdesk-tickets.csv") {
  const headers = ["ticketId","userName","email","mobile","category","subCategory","status","priority","assignedTo","createdAt","closedAt","source"];
  const escape = value => {
    const text = value === undefined || value === null ? "" : String(value);
    return /[",\n\r]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
  };
  const rows = (tickets || []).map(ticket => ({
    ticketId: ticket.id || "",
    userName: ticket.name || "",
    email: ticket.email || "",
    mobile: ticket.mobile || "",
    category: categoryLabel(ticket.category),
    subCategory: ticket.subCategory || "",
    status: ticket.status || "",
    priority: ticket.priority || "",
    assignedTo: ticket.assignedTo || ticket.assigneeName || staffName(ticket.assigneeId),
    createdAt: fmtDate(ticket.createdAt),
    closedAt: fmtDate(ticket.closedAt),
    source: ticket.source || "Portal"
  }));
  const csv = [headers.join(","), ...rows.map(row => headers.map(header => escape(row[header])).join(","))].join("\r\n");
  saveAs(new Blob([csv], { type:"text/csv;charset=utf-8" }), filename);
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
  .app-sidebar{position:fixed!important;top:0!important;left:0!important;width:min(82vw,300px)!important;max-width:300px!important;height:100dvh!important;z-index:120!important;transform:translateX(-105%)!important;transition:transform .25s ease!important;box-shadow:18px 0 44px rgba(0,0,0,.5)!important;border-right:1px solid rgba(125,211,252,.18)!important;overflow-y:auto!important;-webkit-overflow-scrolling:touch!important}
  .app-sidebar.mobile-open{transform:translateX(0)!important}
  .sidebar-overlay{display:block!important;position:fixed!important;inset:0!important;background:rgba(0,0,0,.6)!important;backdrop-filter:blur(5px)!important;z-index:110!important}
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
  .staff-profile-menu{position:fixed!important;left:12px!important;right:12px!important;top:58px!important;width:auto!important;min-width:0!important;max-width:none!important;z-index:95!important}
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

/* ═══ FIRST LOGIN PASSWORD SETUP - MOBILE RESPONSIVE ═══ */
@media (max-width:768px){
  div[style*="minHeight"*="100dvh"][style*="background"*="radial-gradient"]{
    min-height:100dvh!important;
    height:auto!important;
    padding-top:max(16px,env(safe-area-inset-top,0px))!important;
    padding-bottom:max(16px,env(safe-area-inset-bottom,0px))!important;
  }
}

@media (max-width:480px){
  /* Ensure full viewport coverage on small screens */
  div[style*="minHeight"*="100dvh"][style*="background"*="radial-gradient"]{
    min-height:100dvh!important;
    padding:max(12px,env(safe-area-inset-top,0px)) max(12px,env(safe-area-inset-right,0px)) max(12px,env(safe-area-inset-bottom,0px)) max(12px,env(safe-area-inset-left,0px))!important;
    gap:clamp(12px,3vw,16px)!important;
  }

  div[style*="minHeight"*="100dvh"][style*="background"*="radial-gradient"] > div[style*="width"*="100%"]{
    width:100%!important;
    max-width:100%!important;
    padding:0!important;
  }

  div[style*="minHeight"*="100dvh"][style*="background"*="radial-gradient"] h1{
    font-size:clamp(18px,5vw,24px)!important;
    line-height:1.2!important;
    margin-bottom:8px!important;
  }

  div[style*="minHeight"*="100dvh"][style*="background"*="radial-gradient"] p{
    font-size:clamp(11px,3vw,13px)!important;
    line-height:1.4!important;
  }

  .fade-up{
    animation:fadeUp .4s ease forwards;
  }

  div[style*="minHeight"*="100dvh"][style*="background"*="radial-gradient"] .glass{
    padding:clamp(14px,4vw,20px)!important;
    gap:clamp(10px,3vw,14px)!important;
    max-width:100%!important;
    width:100%!important;
    border-radius:12px!important;
  }

  div[style*="minHeight"*="100dvh"][style*="background"*="radial-gradient"] .glass > div:first-child{
    font-size:clamp(10px,2.5vw,12px)!important;
    padding:10px 12px!important;
  }

  div[style*="minHeight"*="100dvh"][style*="background"*="radial-gradient"] label{
    font-size:clamp(10px,2.5vw,11px)!important;
    margin-bottom:4px!important;
  }

  div[style*="minHeight"*="100dvh"][style*="background"*="radial-gradient"] input,
  div[style*="minHeight"*="100dvh"][style*="background"*="radial-gradient"] .pwd-input-wrap input{
    padding:10px 12px!important;
    font-size:14px!important;
    min-height:42px!important;
  }

  div[style*="minHeight"*="100dvh"][style*="background"*="radial-gradient"] .glow-btn{
    width:100%!important;
    min-height:44px!important;
    font-size:clamp(13px,3vw,15px)!important;
    padding:12px 20px!important;
    margin-top:4px!important;
  }

  div[style*="minHeight"*="100dvh"][style*="background"*="radial-gradient"] .pwd-toggle{
    right:10px!important;
    top:50%!important;
  }
}
  }
@media (max-width:768px){
  .staff-chat-shell{
    display:flex!important;
    width:100%!important;
    max-width:100%!important;
    min-height:calc(100dvh - 112px)!important;
    height:calc(100dvh - 112px)!important;
    overflow:hidden!important;
    grid-template-columns:1fr!important;
    gap:0!important;
  }
  .staff-chat-list{
    width:100%!important;
    max-width:100%!important;
    height:100%!important;
    overflow-y:auto!important;
    overflow-x:hidden!important;
    display:block!important;
  }
  .staff-chat-window{
    width:100%!important;
    max-width:100%!important;
    height:100%!important;
    min-height:0!important;
    display:none!important;
    overflow:hidden!important;
  }
  .staff-chat-shell.mobile-chat-open .staff-chat-list{display:none!important}
  .staff-chat-shell.mobile-chat-open .staff-chat-window{display:flex!important}
  .staff-chat-header{flex-shrink:0!important;padding:10px 12px!important}
  .staff-chat-back{display:inline-flex!important;align-items:center;justify-content:center;flex-shrink:0}
  .staff-chat-messages{
    flex:1!important;
    min-height:0!important;
    overflow-y:auto!important;
    overflow-x:hidden!important;
    padding:12px!important;
    padding-bottom:16px!important;
    -webkit-overflow-scrolling:touch!important;
  }
  .staff-chat-bubble{max-width:86%!important}
  .staff-chat-input{
    flex-shrink:0!important;
    position:sticky!important;
    bottom:0!important;
    background:rgba(10,10,20,.96)!important;
    backdrop-filter:blur(18px)!important;
    padding:10px!important;
    padding-bottom:max(12px,env(safe-area-inset-bottom,0px))!important;
    gap:8px!important;
    z-index:3!important;
  }
  .staff-chat-input input{min-width:0!important}
  .staff-chat-input .glow-btn{padding:10px 12px!important;min-width:64px}
  .portal-feedback-tab{bottom:74px!important;right:10px!important}
}
@media (max-width:480px){
  .staff-chat-shell{
    min-height:calc(100dvh - 94px)!important;
    height:calc(100dvh - 94px)!important;
  }
  .staff-chat-input .glow-btn{font-size:12px!important}
}
.ai-helpdesk-wrap{position:fixed;right:12px;bottom:62px;z-index:920;display:flex;flex-direction:column;align-items:flex-end;gap:12px;pointer-events:none}
.ai-helpdesk-wrap>*{pointer-events:auto}
.ai-helpdesk-button{border:1px solid rgba(255,255,255,.2);border-radius:999px;padding:11px 16px;color:#fff;font-weight:900;font-size:13px;background:linear-gradient(135deg,#10b981,#06b6d4,#2563eb,#8b5cf6);box-shadow:0 18px 44px rgba(37,99,235,.38),0 0 30px rgba(6,182,212,.26);backdrop-filter:blur(18px);display:flex;align-items:center;gap:8px;transition:transform .2s,box-shadow .2s}
.ai-helpdesk-button:hover{transform:translateY(-2px);box-shadow:0 24px 58px rgba(37,99,235,.48),0 0 40px rgba(16,185,129,.28)}
.ai-helpdesk-button span{width:22px;height:22px;border-radius:50%;background:rgba(255,255,255,.16);display:inline-flex;align-items:center;justify-content:center;color:#d1fae5}
.ai-helpdesk-panel{width:min(380px,calc(100vw - 24px));height:min(620px,calc(100dvh - 148px));display:flex;flex-direction:column;overflow:hidden;border-radius:20px!important;background:radial-gradient(circle at 18% 0,rgba(16,185,129,.2),transparent 34%),radial-gradient(circle at 82% 8%,rgba(139,92,246,.25),transparent 38%),rgba(8,13,28,.92)!important;border:1px solid rgba(125,211,252,.28)!important;box-shadow:0 28px 80px rgba(0,0,0,.48),0 0 42px rgba(6,182,212,.22)!important}
.ai-helpdesk-head{padding:14px 15px;display:flex;align-items:center;gap:10px;border-bottom:1px solid rgba(255,255,255,.09);background:linear-gradient(135deg,rgba(37,99,235,.24),rgba(139,92,246,.18),rgba(6,182,212,.12))}
.ai-helpdesk-avatar{width:40px;height:40px;border-radius:14px;background:linear-gradient(135deg,#10b981,#06b6d4,#8b5cf6);display:flex;align-items:center;justify-content:center;color:#fff;font-weight:900;box-shadow:0 0 24px rgba(6,182,212,.3);flex-shrink:0}
.ai-helpdesk-close{width:32px;height:32px;border-radius:10px;border:1px solid rgba(255,255,255,.12);background:rgba(255,255,255,.08);color:#e2e8f0;font-size:20px;line-height:1;display:flex;align-items:center;justify-content:center}
.ai-helpdesk-messages{flex:1;min-height:0;overflow-y:auto;overflow-x:hidden;padding:14px;background:linear-gradient(180deg,rgba(255,255,255,.02),rgba(255,255,255,.04));-webkit-overflow-scrolling:touch}
.ai-helpdesk-row{display:flex;margin-bottom:10px}.ai-helpdesk-row.user{justify-content:flex-end}.ai-helpdesk-row.bot{justify-content:flex-start}
.ai-helpdesk-bubble{max-width:82%;border-radius:16px;padding:10px 12px;font-size:13px;line-height:1.45;border:1px solid rgba(255,255,255,.1);box-shadow:0 10px 28px rgba(0,0,0,.2)}
.ai-helpdesk-bubble.user{background:linear-gradient(135deg,rgba(99,102,241,.48),rgba(6,182,212,.25));color:#fff;border-top-right-radius:5px}
.ai-helpdesk-bubble.bot{background:rgba(255,255,255,.075);color:#e2e8f0;border-top-left-radius:5px}
.ai-helpdesk-menu-card,.ai-helpdesk-steps-card{min-width:min(292px,100%);border-radius:14px;background:linear-gradient(180deg,rgba(15,23,42,.7),rgba(15,23,42,.46));border:1px solid rgba(125,211,252,.18);overflow:hidden}
.ai-helpdesk-card-title{padding:10px 11px;font-weight:900;color:#f8fafc;background:rgba(14,165,233,.11);border-bottom:1px solid rgba(125,211,252,.14)}
.ai-helpdesk-menu-list{display:grid;gap:6px;padding:9px}
.ai-helpdesk-menu-item{display:flex;align-items:center;gap:9px;border-radius:11px;background:rgba(255,255,255,.065);border:1px solid rgba(255,255,255,.08);padding:8px 9px;color:#e2e8f0}
.ai-helpdesk-menu-item span{width:24px;height:24px;border-radius:8px;background:rgba(16,185,129,.18);color:#a7f3d0;display:inline-flex;align-items:center;justify-content:center;font-size:12px;font-weight:900;flex-shrink:0}
.ai-helpdesk-menu-item strong{font-size:12px;line-height:1.25}
.ai-helpdesk-menu-hint{margin:0 9px 9px;border-radius:10px;background:rgba(255,255,255,.055);border:1px solid rgba(255,255,255,.08);padding:7px 9px;color:#94a3b8;font-size:11px;font-weight:800}
.ai-helpdesk-steps-card ol{margin:0;padding:10px 12px 4px 30px}
.ai-helpdesk-steps-card li{margin:0 0 7px;padding-left:2px;color:#e2e8f0}
.ai-helpdesk-causes{margin:6px 10px 8px;border-radius:12px;background:rgba(15,23,42,.58);border:1px solid rgba(255,255,255,.08);padding:9px 10px;color:#cbd5e1;font-size:12px}
.ai-helpdesk-causes div{font-weight:900;color:#f8fafc;margin-bottom:6px}
.ai-helpdesk-causes ul{margin:0;padding-left:18px}.ai-helpdesk-causes li{margin:0 0 5px;color:#cbd5e1}
.ai-helpdesk-url{display:block;margin:6px 10px 8px;border-radius:10px;background:rgba(14,165,233,.12);border:1px solid rgba(125,211,252,.2);padding:8px 9px;color:#7dd3fc;font-size:12px;font-weight:900;text-decoration:none;overflow-wrap:anywhere}
.ai-helpdesk-url:hover{background:rgba(14,165,233,.2);color:#bae6fd}
.ai-helpdesk-card-footer{margin:8px 10px 10px;border-radius:12px;background:rgba(16,185,129,.1);border:1px solid rgba(16,185,129,.18);padding:9px 10px;white-space:pre-wrap;color:#bbf7d0;font-weight:800;font-size:12px}
.ai-helpdesk-input{display:flex;gap:8px;padding:12px;border-top:1px solid rgba(255,255,255,.08);background:rgba(10,10,20,.94);backdrop-filter:blur(18px);flex-shrink:0}
.ai-helpdesk-input input{min-width:0}.ai-helpdesk-input .glow-btn{padding:10px 13px;font-size:12px}
.mic-btn{width:42px;min-width:42px;height:42px;border-radius:13px;border:1px solid rgba(125,211,252,.2);background:rgba(14,165,233,.12)!important;color:#bae6fd!important;display:inline-flex;align-items:center;justify-content:center;gap:7px;font-size:18px;font-weight:900;box-shadow:0 10px 24px rgba(6,182,212,.12);white-space:nowrap;overflow:hidden}
.mic-btn:hover{transform:translateY(-2px);border-color:rgba(34,211,238,.52)!important;box-shadow:0 16px 36px rgba(6,182,212,.2)}
.mic-btn.listening{width:auto;min-width:118px;padding:0 12px;background:linear-gradient(135deg,#ef4444,#8b5cf6,#06b6d4)!important;color:#fff!important;animation:micGlow 1.1s ease-in-out infinite;box-shadow:0 0 0 6px rgba(239,68,68,.12),0 0 30px rgba(34,211,238,.34)}
.mic-btn-text{font-size:11px;font-weight:1000;letter-spacing:.1px}
@keyframes micGlow{0%,100%{transform:scale(1);filter:saturate(1)}50%{transform:scale(1.08);filter:saturate(1.3)}}
.ai-typing{display:flex;align-items:center;gap:6px;color:rgba(226,232,240,.74)!important;font-style:italic}.ai-typing span{width:5px;height:5px;border-radius:50%;background:#67e8f9;display:inline-block;animation:pulse 1s infinite}.ai-typing span:nth-child(2){animation-delay:.15s}.ai-typing span:nth-child(3){animation-delay:.3s}
.theme-glow{position:fixed;inset:0;pointer-events:none;z-index:0;overflow:hidden;background:radial-gradient(circle at 14% 18%,rgba(37,99,235,.08),transparent 30%),radial-gradient(circle at 82% 8%,rgba(139,92,246,.07),transparent 32%),radial-gradient(circle at 70% 86%,rgba(6,182,212,.06),transparent 34%);animation:themeGlowShift 14s ease-in-out infinite alternate}
@keyframes themeGlowShift{from{filter:hue-rotate(0deg);opacity:.72;transform:scale(1)}to{filter:hue-rotate(12deg);opacity:.95;transform:scale(1.03)}}
.smart-welcome-once{animation:smartWelcomeFade 5s ease forwards}
@keyframes smartWelcomeFade{0%,82%{opacity:1;transform:translateY(0)}100%{opacity:0;transform:translateY(-8px)}}
.incident-banner{margin:14px 24px 0;border-radius:14px;border:1px solid rgba(125,211,252,.22);overflow:hidden;box-shadow:0 18px 44px rgba(0,0,0,.24),0 0 24px rgba(6,182,212,.12)}
.incident-banner.info{background:linear-gradient(135deg,rgba(14,165,233,.18),rgba(37,99,235,.12))}
.incident-banner.warning{background:linear-gradient(135deg,rgba(245,158,11,.2),rgba(249,115,22,.12))}
.incident-banner.critical{background:linear-gradient(135deg,rgba(239,68,68,.24),rgba(147,51,234,.12))}
.incident-track{display:inline-block;white-space:nowrap;padding:10px 16px;color:#f8fafc;font-size:13px;font-weight:900;animation:incidentMarquee 22s linear infinite;text-shadow:0 0 18px rgba(255,255,255,.18)}
@keyframes incidentMarquee{from{transform:translateX(18%)}to{transform:translateX(-100%)}}
.app-sidebar,.app-main{position:relative;z-index:1}
@media (max-width:768px){.ai-helpdesk-wrap{right:10px;bottom:132px}.ai-helpdesk-panel{width:calc(100vw - 20px);height:calc(100dvh - 176px);border-radius:18px!important}.ai-helpdesk-bubble{max-width:90%}.ai-helpdesk-button{padding:10px 13px;font-size:12px}.ai-helpdesk-input{padding-bottom:max(12px,env(safe-area-inset-bottom,0px))}.incident-banner{margin:10px 12px 0}.incident-track{font-size:12px;padding:9px 12px}.mic-btn{width:38px;min-width:38px;height:38px}.mic-btn.listening{min-width:102px}}
@media (max-width:480px){.ai-helpdesk-wrap{right:8px;bottom:126px}.ai-helpdesk-panel{width:calc(100vw - 16px);height:calc(100dvh - 164px)}.ai-helpdesk-head{padding:12px}.ai-helpdesk-messages{padding:12px}.ai-helpdesk-input .glow-btn{min-width:58px}}`;

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
  const styles={
    Low:{bg:"rgba(16,185,129,.16)",col:"#86efac",border:"rgba(16,185,129,.28)"},
    Medium:{bg:"rgba(245,158,11,.18)",col:"#fbbf24",border:"rgba(245,158,11,.32)"},
    High:{bg:"rgba(239,68,68,.18)",col:"#f87171",border:"rgba(239,68,68,.34)"},
    Critical:{bg:"rgba(220,38,38,.22)",col:"#fecaca",border:"rgba(248,113,113,.42)"}
  }[p]||{bg:"rgba(148,163,184,.16)",col:"#cbd5e1",border:"rgba(148,163,184,.24)"};
  return <span className="tag" style={{background:styles.bg,color:styles.col,border:`1px solid ${styles.border}`}}>▲ {p}</span>;
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
function StatCard({label,value,icon,color,sub,onClick}) {
  return (
    <div className="glass" onClick={onClick} style={{padding:"20px 22px",display:"flex",flexDirection:"column",gap:8,cursor:onClick?"pointer":"default"}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
        <span style={{fontSize:13,color:"rgba(226,232,240,0.5)",fontWeight:500}}>{label}</span>
        <div style={{fontSize:24}}>{icon}</div>
      </div>
      <div style={{fontSize:32,fontWeight:700,fontFamily:"Syne",color,letterSpacing:"-1px"}}>{value}</div>
      {sub&&<div style={{fontSize:12,color:"rgba(226,232,240,0.4)"}}>{sub}</div>}
    </div>
  );
}

function SmartWelcome({session,visible=true}) {
  const name=getDisplayName(session);
  if(!visible) return null;
  return (
    <div className="glass smart-welcome-once" style={{padding:"18px 20px",marginBottom:18,background:"linear-gradient(135deg,rgba(14,165,233,.13),rgba(139,92,246,.1),rgba(255,255,255,.04))",border:"1px solid rgba(125,211,252,.16)"}}>
      <h2 style={{fontFamily:"Syne",fontSize:"clamp(20px,4vw,26px)",fontWeight:900,color:"#f8fafc",marginBottom:5,lineHeight:1.18}}>{getGreetingLabel()}, {name} 👋</h2>
      <p style={{fontSize:14,color:"rgba(226,232,240,.62)",margin:0}}>How can IT Helpdesk assist you today?</p>
    </div>
  );
}

function CampusIncidentBanner({incidents=[]}) {
  const active = (incidents || []).filter(incident => incident.active && (!incident.expiryAt || Number(incident.expiryAt) > Date.now()));
  if(!active.length) return null;
  const text = active.map(incident => `${incident.severity}: ${incident.title} - ${incident.message}`).join("     |     ");
  const severity = active.some(i=>i.severity==="Critical") ? "critical" : active.some(i=>i.severity==="Warning") ? "warning" : "info";
  return (
    <div className={`incident-banner ${severity}`} title="Campus IT incident update">
      <span className="incident-track">⚡ Campus Incident: {text}</span>
    </div>
  );
}

function IncidentManager({incidents,onSave,toast}) {
  const empty={title:"",message:"",severity:"Info",active:true,expiryAt:""};
  const [form,setForm]=useState(empty);
  const [saving,setSaving]=useState(false);
  const set=(k,v)=>setForm(f=>({...f,[k]:v}));
  const edit=incident=>setForm({
    ...incident,
    expiryAt: incident.expiryAt ? new Date(Number(incident.expiryAt)).toISOString().slice(0,16) : ""
  });
  const submit=async()=>{
    if(!form.title.trim() || !form.message.trim()) {
      toast("Incident title and message are required.","error");
      return;
    }
    setSaving(true);
    try {
      await onSave({
        ...form,
        title:form.title.trim(),
        message:form.message.trim(),
        active:Boolean(form.active),
        expiryAt:form.expiryAt ? new Date(form.expiryAt).getTime() : null,
        updatedAt:Date.now(),
        createdAt:form.createdAt || Date.now()
      });
      setForm(empty);
    } catch(error) {
      console.error("Incident save failed:", error);
      toast("Incident could not be saved.","error");
    } finally {
      setSaving(false);
    }
  };
  const disable=incident=>onSave({...incident,active:false,updatedAt:Date.now()}).catch(error=>{
    console.error("Incident disable failed:", error);
    toast("Incident could not be disabled.","error");
  });
  return (
    <div className="glass" style={{padding:"18px",display:"grid",gap:14}}>
      <div>
        <h3 style={{fontFamily:"Syne",fontSize:16,fontWeight:900,color:"#e2e8f0"}}>Campus Incident Banner</h3>
        <p style={{fontSize:12,color:"rgba(226,232,240,.52)",marginTop:4}}>Create a small portal-wide alert for maintenance or campus IT incidents.</p>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
        <input value={form.title} onChange={e=>set("title",e.target.value)} placeholder="Title, e.g. Moodle maintenance" />
        <select value={form.severity} onChange={e=>set("severity",e.target.value)}>
          {["Info","Warning","Critical"].map(item=><option key={item}>{item}</option>)}
        </select>
        <input value={form.message} onChange={e=>set("message",e.target.value)} placeholder="Message shown to users" />
        <input type="datetime-local" value={form.expiryAt || ""} onChange={e=>set("expiryAt",e.target.value)} />
      </div>
      <label style={{display:"inline-flex",gap:8,alignItems:"center",fontSize:13,color:"#e2e8f0",fontWeight:800}}>
        <input type="checkbox" checked={Boolean(form.active)} onChange={e=>set("active",e.target.checked)} style={{width:"auto"}} /> Active
      </label>
      <div style={{display:"flex",justifyContent:"flex-end",gap:10}}>
        <button type="button" onClick={()=>setForm(empty)} style={{background:"rgba(255,255,255,.06)",border:"1px solid rgba(255,255,255,.1)",color:"#e2e8f0",padding:"9px 14px",borderRadius:10,fontWeight:800}}>Clear</button>
        <button className="glow-btn" type="button" onClick={submit} disabled={saving}>{saving ? "Saving..." : form.id ? "Update Incident" : "Create Incident"}</button>
      </div>
      <div style={{display:"grid",gap:8}}>
        {(incidents || []).slice(0,5).map(incident=>(
          <div key={incident.id} className="glass2" style={{padding:"12px",display:"flex",justifyContent:"space-between",gap:10,alignItems:"center",flexWrap:"wrap"}}>
            <div>
              <div style={{fontSize:13,fontWeight:900,color:"#fff"}}>{incident.title} <span className="tag" style={{fontSize:10}}>{incident.severity}</span></div>
              <div style={{fontSize:12,color:"rgba(226,232,240,.58)",marginTop:3}}>{incident.message}</div>
              <div style={{fontSize:11,color:incident.active?"#86efac":"#cbd5e1",marginTop:3}}>{incident.active ? "Active" : "Disabled"}{incident.expiryAt ? ` · Expires ${fmtDate(incident.expiryAt)}` : ""}</div>
            </div>
            <div style={{display:"flex",gap:8}}>
              <button type="button" onClick={()=>edit(incident)} style={{background:"rgba(14,165,233,.12)",border:"1px solid rgba(125,211,252,.22)",color:"#bae6fd",padding:"7px 10px",borderRadius:9,fontWeight:800}}>Edit</button>
              {incident.active&&<button type="button" onClick={()=>disable(incident)} style={{background:"rgba(239,68,68,.12)",border:"1px solid rgba(239,68,68,.25)",color:"#fecaca",padding:"7px 10px",borderRadius:9,fontWeight:800}}>Disable</button>}
            </div>
          </div>
        ))}
        {!incidents.length&&<EmptyState message="No campus incidents configured." icon="⚡" />}
      </div>
    </div>
  );
}

function EmptyState({message,icon="ℹ️"}) {
  return <div className="glass2" style={{gridColumn:"1/-1",padding:"28px 18px",textAlign:"center",color:"rgba(226,232,240,.64)",borderStyle:"dashed"}}><div style={{fontSize:34,marginBottom:8}}>{icon}</div><div style={{fontSize:14,fontWeight:800}}>{message}</div></div>;
}

// ── AUDIT TIMELINE ────────────────────────────────────────────────────────
function AuditTimeline({timeline}) {
  const icons={Created:"🆕",Assigned:"👤",Reassigned:"🔄","Status changed":"📋",Commented:"💬",Closed:"✅",default:"📌"};
  return (
    <div style={{display:"flex",flexDirection:"column",gap:0}}>
      {(timeline||[]).map((ev,i)=>{
        const action=ev?.action || "Updated";
        const icon=Object.keys(icons).find(k=>action.startsWith(k))||"default";
        return (
          <div key={i} style={{display:"flex",gap:14,alignItems:"flex-start",position:"relative",paddingBottom:i<timeline.length-1?16:0}}>
            {i<timeline.length-1&&<div style={{position:"absolute",left:15,top:32,width:2,height:"calc(100% - 10px)",background:"rgba(255,255,255,0.07)"}}/>}
            <div style={{width:30,height:30,borderRadius:"50%",background:"rgba(99,102,241,0.15)",border:"1px solid rgba(99,102,241,0.3)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:14,flexShrink:0}}>
              {icons[icon]}
            </div>
            <div style={{paddingTop:4}}>
              <div style={{fontSize:13,color:"#e2e8f0",fontWeight:500}}>{action}</div>
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

function TicketFeedbackSection({ticket,onSubmit,toast}) {
  const [rating,setRating]=useState(Number(ticket?.userFeedbackRating || 0));
  const [comment,setComment]=useState(ticket?.userFeedbackComment || "");
  const [saving,setSaving]=useState(false);
  if(!ticket || !isTicketFeedbackPending(ticket)) return null;
  const submit=async()=>{
    if(!rating){
      toast("Please select a rating from 1 to 5.","error");
      return;
    }
    setSaving(true);
    try {
      await Promise.resolve(onSubmit({rating,comment:comment.trim()}));
    } catch(error) {
      console.error("Ticket feedback submit failed:", error);
    } finally {
      setSaving(false);
    }
  };
  return (
    <div className="glass2" style={{padding:"16px",borderColor:"rgba(16,185,129,.28)",background:"rgba(16,185,129,.08)",display:"flex",flexDirection:"column",gap:12}}>
      <div style={{fontSize:16,fontWeight:900,color:"#fff"}}>Your ticket has been closed. Please share your feedback.</div>
      <div>
        <div style={{fontSize:12,color:"rgba(226,232,240,.58)",marginBottom:8,fontWeight:800}}>Rating</div>
        <div style={{display:"flex",gap:10,flexWrap:"wrap",alignItems:"center"}}>
          {[1,2,3,4,5].map(value=>{
            const selected=rating===value;
            return (
              <button
                key={value}
                type="button"
                onClick={()=>setRating(value)}
                onMouseEnter={e=>{if(!selected){e.currentTarget.style.transform="translateY(-2px) scale(1.04)";e.currentTarget.style.borderColor="rgba(125,211,252,.35)";}}}
                onMouseLeave={e=>{if(!selected){e.currentTarget.style.transform="translateY(0) scale(1)";e.currentTarget.style.borderColor="rgba(255,255,255,.12)";}}}
                onMouseDown={e=>{e.currentTarget.style.transform=selected?"scale(1.1)":"scale(.97)";}}
                onMouseUp={e=>{e.currentTarget.style.transform=selected?"scale(1.16)":"scale(1.04)";}}
                aria-pressed={selected}
                style={{
                  minWidth:selected?62:42,
                  height:selected?50:42,
                  borderRadius:14,
                  border:`1px solid ${selected?"rgba(125,211,252,.72)":"rgba(255,255,255,.12)"}`,
                  background:selected
                    ?"linear-gradient(135deg,#2563eb 0%,#7c3aed 48%,#06b6d4 100%)"
                    :"rgba(255,255,255,.06)",
                  color:selected?"#fff":"#e2e8f0",
                  fontSize:selected?20:15,
                  fontWeight:selected?1000:900,
                  cursor:"pointer",
                  transform:selected?"scale(1.16)":"scale(1)",
                  transition:"transform .2s ease, box-shadow .2s ease, background .2s ease, border-color .2s ease, color .2s ease, min-width .2s ease, height .2s ease",
                  boxShadow:selected
                    ?"0 0 0 4px rgba(14,165,233,.16), 0 14px 34px rgba(79,70,229,.42), 0 0 24px rgba(34,211,238,.28)"
                    :"0 6px 16px rgba(0,0,0,.16)",
                  display:"inline-flex",
                  flexDirection:"column",
                  alignItems:"center",
                  justifyContent:"center",
                  gap:1,
                  position:"relative",
                  WebkitTapHighlightColor:"transparent"
                }}
              >
                <span style={{lineHeight:1}}>{value}</span>
                {selected&&<span style={{fontSize:9,fontWeight:900,letterSpacing:.2,lineHeight:1,color:"rgba(255,255,255,.9)"}}>✓ Selected</span>}
              </button>
            );
          })}
        </div>
      </div>
      <textarea rows={3} value={comment} onChange={e=>setComment(e.target.value)} placeholder="Optional comment" style={{resize:"vertical"}} />
      <div style={{display:"flex",gap:10,justifyContent:"flex-end",flexWrap:"wrap"}}>
        <button type="button" onClick={()=>toast("No problem. Feedback reminder will stay in My Tickets.","info")} style={{background:"rgba(255,255,255,.06)",border:"1px solid rgba(255,255,255,.1)",color:"#e2e8f0",padding:"9px 14px",borderRadius:10,fontWeight:900}}>Later</button>
        <button className="glow-btn" type="button" onClick={submit} disabled={saving}>{saving?"Submitting...":"Submit Feedback"}</button>
      </div>
    </div>
  );
}

function RemoteSupportDialog({ticket,onClose,onSubmit,toast}) {
  const [tool,setTool]=useState(ticket?.remoteSupportTool || "AnyDesk");
  const [remoteId,setRemoteId]=useState(ticket?.remoteSupportId || "");
  const [note,setNote]=useState(ticket?.remoteSupportNote || "");
  const [saving,setSaving]=useState(false);
  const submit=async()=>{
    if(!ticket?.id) {
      toast("Ticket not found. Please reopen the ticket.","error");
      console.error("Remote support request failed: missing ticket", ticket);
      return;
    }
    if(!remoteId.trim()) {
      toast("Please enter your remote support ID.","error");
      return;
    }
    setSaving(true);
    try {
      await Promise.resolve(onSubmit({tool,remoteId:remoteId.trim(),note:note.trim()}));
      onClose();
    } catch(error) {
      console.error("Remote support request failed:", error);
      toast("Remote support request could not be saved.","error");
    } finally {
      setSaving(false);
    }
  };
  return (
    <div style={{display:"flex",flexDirection:"column",gap:16}}>
      <div className="glass2" style={{padding:"14px",borderColor:"rgba(245,158,11,.28)",background:"rgba(245,158,11,.08)",color:"#fde68a",fontSize:13,fontWeight:800}}>
        Only share remote access ID with official IT staff.
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
        <div><label style={{fontSize:12,color:"rgba(226,232,240,.58)",display:"block",marginBottom:6}}>Remote Support Tool</label><select value={tool} onChange={e=>setTool(e.target.value)}><option>AnyDesk</option><option>TeamViewer</option></select></div>
        <div><label style={{fontSize:12,color:"rgba(226,232,240,.58)",display:"block",marginBottom:6}}>Remote ID</label><input value={remoteId} onChange={e=>setRemoteId(e.target.value)} placeholder="Enter AnyDesk/TeamViewer ID" /></div>
      </div>
      <div><label style={{fontSize:12,color:"rgba(226,232,240,.58)",display:"block",marginBottom:6}}>Optional Note</label><textarea rows={3} value={note} onChange={e=>setNote(e.target.value)} placeholder="Best time to connect or short note" style={{resize:"vertical"}} /></div>
      <div style={{display:"flex",justifyContent:"flex-end",gap:10,flexWrap:"wrap"}}>
        <button type="button" onClick={onClose} style={{background:"rgba(255,255,255,.06)",border:"1px solid rgba(255,255,255,.1)",color:"#e2e8f0",padding:"9px 14px",borderRadius:10,fontWeight:800}}>Cancel</button>
        <button className="glow-btn" type="button" onClick={submit} disabled={saving}>{saving ? "Saving..." : "Submit Remote Request"}</button>
      </div>
    </div>
  );
}

// ── TICKET DETAIL ─────────────────────────────────────────────────────────
function TicketDetail({ticketId,tickets,setTickets,onClose,isAdmin,isStaff,staffId,staffName,toast,staffProfiles={},staffStatuses={}}) {
  const ticket=tickets.find(t=>t.id===ticketId)||null;
  const [comment,setComment]=useState("");
  const [editStatus,setEditStatus]=useState(ticket?.status || "Open");
  const [editAssignee,setEditAssignee]=useState(ticket?.assigneeId || STAFF_BASE[0]?.id || 1);
  const [showClose,setShowClose]=useState(false);
  const [showRemoteSupport,setShowRemoteSupport]=useState(false);
  const currentTicket=ticket;

  useEffect(()=>{
    if(!currentTicket) return;
    setEditStatus(currentTicket.status || "Open");
    setEditAssignee(currentTicket.assigneeId || STAFF_BASE[0]?.id || 1);
  },[currentTicket?.id,currentTicket?.status,currentTicket?.assigneeId]);

  const applyTicketChanges = async (changes, auditAction, remark = "") => {
    if(!currentTicket){
      toast("Ticket not found. Please reopen the ticket.", "error");
      return null;
    }
    if(!currentTicket.id){
      toast("Ticket update failed: missing ticket ID.", "error");
      console.error("Ticket update failed: missing ticket id", currentTicket);
      return null;
    }
    let closedByStatusChange = false;
    let closedTicket = null;
    let updatedTicket = null;

    setTickets(ts=>ts.map(t=>{
      if(t.id!==currentTicket.id) return t;
      const actor=isAdmin?"Admin":staffName||"User";
      const tl=[...(t.timeline||[]),{action:auditAction || "Updated",remark,at:Date.now(),by:actor}];
      const oldStatus=t.status;
      const closingNow=["Closed","Resolved"].includes(changes.status)&&!["Closed","Resolved"].includes(oldStatus);
      const nextAssignee=changes.assigneeId ? STAFF_BASE.find(s=>s.id===Number(changes.assigneeId)) : null;
      const updated={...t,...changes,updatedAt:Date.now(),timeline:tl};
      if(nextAssignee){
        updated.assigneeName=nextAssignee.name;
        updated.assignedTo=nextAssignee.name;
      } else if(Number(changes.assigneeId)===0) {
        updated.assigneeName="IT Support Team";
        updated.assignedTo="IT Support Team";
      }
      if(closingNow){
        updated.closedAt=Date.now();
        updated.closingRemarks=remark||"Closed from ticket controls.";
        updated.resolutionTime=updated.closedAt-t.createdAt;
        updated.feedbackSubmitted=false;
        updated.userConfirmedResolved=false;
        updated.userFeedbackStatus="";
        updated.userReviewedAt=null;
        updated.userFeedbackRating=0;
        updated.userFeedbackComment="";
        updated.userFeedbackAt=null;
        updated.feedbackStatus="";
        updated.feedbackReadByAdmin=false;
        updated.feedbackReadByStaff=false;
        closedByStatusChange=true;
        closedTicket=updated;
      }
      if(changes.assigneeId&&Number(changes.assigneeId)!==Number(t.assigneeId)){
        emailTicketAssigned(updated,STAFF_BASE.find(s=>s.id===Number(changes.assigneeId)),STAFF_BASE.find(s=>s.id===t.assigneeId),actor,remark);
      }
      updatedTicket = updated;
      return updated;
    }));

    toast(closedByStatusChange ? "Ticket closed" : auditAction,"success");
    if(closedTicket) {
      notifyTicketClosed(closedTicket, isAdmin?"Admin":staffName||"IT Support", remark);
      showBrowserNotification("Ticket closed", `${closedTicket.id} has been closed/resolved.`);
    }
    if(updatedTicket && isClosedTicket(currentTicket) && changes.status && !isClosedTicket(updatedTicket)) {
      showBrowserNotification("Ticket reopened", `${updatedTicket.id} has been reopened.`);
    }
    return updatedTicket;
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

  useEffect(()=>{
    if(!currentTicket?.id || !isTicketFeedbackUnread(currentTicket,isAdmin,isStaff)) return;
    const now=Date.now();
    const updated={
      ...currentTicket,
      feedbackReadByAdmin:isAdmin ? true : currentTicket.feedbackReadByAdmin,
      feedbackReadByStaff:isStaff ? true : currentTicket.feedbackReadByStaff,
      feedbackReadAt:now,
      feedbackReadBy:isAdmin ? "Admin" : staffName || "IT Staff",
      updatedAt:now
    };
    setTickets(ts=>ts.map(t=>t.id===currentTicket.id?updated:t));
    if(ONLINE_TICKETS_ENABLED) saveTicket(updated).catch(error=>console.error("Feedback read update failed:", error));
  },[currentTicket?.id,currentTicket?.feedbackStatus,currentTicket?.feedbackReadByAdmin,currentTicket?.feedbackReadByStaff,isAdmin,isStaff,staffName,setTickets]);

  const submitTicketFeedback=async({rating,comment})=>{
    try {
      if(!currentTicket){
        toast("Ticket not found. Please reopen the ticket.","error");
        return;
      }
      if(!currentTicket.id){
        toast("Feedback could not be saved: missing ticket ID.","error");
        console.error("Ticket feedback submit failed: missing ticket id", currentTicket);
        return;
      }
      if(!isClosedTicket(currentTicket)){
        toast("Feedback is available after the ticket is closed.","error");
        return;
      }
      const now=Date.now();
      const updated={
        ...currentTicket,
        userFeedbackRating:Number(rating),
        userFeedbackComment:comment || "",
        userFeedbackAt:now,
        feedbackStatus:"Submitted",
        feedbackSubmitted:true,
        feedbackReadByAdmin:false,
        feedbackReadByStaff:false,
        updatedAt:now,
        timeline:[...(currentTicket.timeline||[]),{action:"Feedback submitted",remark:`Rating: ${rating}/5`,at:now,by:currentTicket.email || "User"}]
      };
      if(ONLINE_TICKETS_ENABLED) await saveTicket(updated);
      setTickets(ts=>ts.map(t=>t.id===currentTicket.id?updated:t));
      toast("Feedback submitted. Thank you.","success");
    } catch(error) {
      console.error("Ticket feedback submit failed:", error);
      toast("Feedback could not be submitted right now. Please try again.","error");
      throw error;
    }
  };

  const handleCloseTicket=async(remarks)=>{
    const closedAt=Date.now();
    const closedBy = staffName || (isAdmin ? "Admin" : "User");
    let closedTicket = null;
    setTickets(ts=>ts.map(t=>{
      if(t.id!==ticketId) return t;
      if(t.status==="Closed") return t;
      const updated={...t,status:"Closed",closedAt,closingRemarks:remarks,resolutionTime:closedAt-t.createdAt,updatedAt:closedAt,feedbackSubmitted:false,userConfirmedResolved:false,userFeedbackStatus:"",userReviewedAt:null,userFeedbackRating:0,userFeedbackComment:"",userFeedbackAt:null,feedbackStatus:"",feedbackReadByAdmin:false,feedbackReadByStaff:false,
        timeline:[...(t.timeline||[]),{action:"Closed",remark:remarks,at:closedAt,by:closedBy}]};
      closedTicket = updated;
      return updated;
    }));

    setShowClose(false);
    if(closedTicket){
      toast("Ticket closed", "success");
      notifyTicketClosed(closedTicket, closedBy, remarks);
      showBrowserNotification("Ticket closed", `${closedTicket.id} has been closed/resolved.`);
    }
  };

  const submitRemoteSupport=async({tool,remoteId,note})=>{
    if(!currentTicket?.id) {
      toast("Ticket update failed: missing ticket ID.","error");
      console.error("Remote support request failed: missing ticket id", currentTicket);
      return;
    }
    const now=Date.now();
    const updated={
      ...currentTicket,
      remoteSupportRequested:true,
      remoteSupportTool:tool,
      remoteSupportId:remoteId,
      remoteSupportNote:note,
      remoteSupportRequestedAt:now,
      updatedAt:now,
      timeline:[...(currentTicket.timeline||[]),{action:"Remote support requested",remark:`${tool}: ${remoteId}`,at:now,by:currentTicket.email || currentTicket.name || "User"}]
    };
    try {
      if(ONLINE_TICKETS_ENABLED) await saveTicket(updated);
      setTickets(ts=>ts.map(t=>t.id===currentTicket.id?updated:t));
      toast("Remote support request saved.","success");
      showBrowserNotification("Remote support requested", `${updated.id} needs remote support.`);
    } catch(error) {
      console.error("Remote support request save failed:", error);
      toast("Remote support request could not be saved.","error");
      throw error;
    }
  };

  if(!currentTicket) {
    return (
      <div className="glass2" style={{padding:"16px",borderColor:"rgba(239,68,68,.28)",background:"rgba(239,68,68,.08)",color:"#fecaca"}}>
        Ticket details are not available. Please close this window and open the ticket again.
      </div>
    );
  }

  const assignee=STAFF_BASE.find(s=>s.id===currentTicket.assigneeId);
  const cat=CATEGORIES.find(c=>c.id===currentTicket.category);
  const canClose=(isAdmin||(isStaff&&currentTicket.assigneeId===staffId))&&currentTicket.status!=="Closed";

  if(showClose) return (
    <div>
      <button onClick={()=>setShowClose(false)} style={{background:"none",border:"none",color:"rgba(226,232,240,0.5)",fontSize:13,marginBottom:16,display:"flex",alignItems:"center",gap:6}}>← Back to ticket</button>
      <CloseTicketDialog ticket={currentTicket} onClose={()=>setShowClose(false)} onConfirm={handleCloseTicket}/>
    </div>
  );

  return (
    <>
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
          {currentTicket.remoteSupportRequested&&<span className="tag" style={{background:"rgba(14,165,233,.14)",color:"#bae6fd"}}>Remote Support Requested</span>}
          {!isAdmin&&!isStaff&&!isClosedTicket(currentTicket)&&<button className="glow-btn" type="button" onClick={()=>setShowRemoteSupport(true)} style={{padding:"7px 14px",fontSize:13}}>Request Remote Support</button>}
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

      {!isAdmin&&!isStaff&&isTicketFeedbackPending(currentTicket)&&(
        <TicketFeedbackSection ticket={currentTicket} onSubmit={submitTicketFeedback} toast={toast} />
      )}

      {(isAdmin||isStaff)&&currentTicket.feedbackStatus==="Submitted"&&(
        <div className="glass2" style={{padding:"16px",borderColor:"rgba(16,185,129,.24)",background:"rgba(16,185,129,.07)"}}>
          <div style={{display:"flex",justifyContent:"space-between",gap:10,alignItems:"center",marginBottom:10,flexWrap:"wrap"}}>
            <div style={{fontSize:13,fontWeight:900,color:"#bbf7d0"}}>User Feedback</div>
            {isTicketFeedbackUnread(currentTicket,isAdmin,isStaff)&&<span className="tag" style={{background:"rgba(16,185,129,.16)",color:"#86efac"}}>New feedback</span>}
          </div>
          <div style={{fontSize:13,color:"rgba(226,232,240,.78)",lineHeight:1.6}}>
            <div>Rating: <span style={{color:"#fff",fontWeight:900}}>{currentTicket.userFeedbackRating || "—"}/5</span></div>
            {currentTicket.userFeedbackComment&&<div>Comment: <span style={{color:"#fff"}}>{currentTicket.userFeedbackComment}</span></div>}
            <div style={{color:"rgba(226,232,240,.45)",fontSize:12,marginTop:4}}>Submitted: {fmtDate(currentTicket.userFeedbackAt)}</div>
          </div>
        </div>
      )}

      {/* Info */}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
        {[["Submitted by",currentTicket.name],["Email",currentTicket.email],["Department",currentTicket.dept],["Mobile",currentTicket.mobile||"—"],["Location",currentTicket.location||"—"],["Priority",currentTicket.priority]].map(([l,v])=>(
          <div key={l} className="glass" style={{padding:"12px 14px"}}><div style={{fontSize:11,color:"rgba(226,232,240,0.4)"}}>{l}</div><div style={{fontSize:13,fontWeight:500,color:"#e2e8f0",marginTop:3}}>{v}</div></div>
        ))}
      </div>

      <div className="glass" style={{padding:"16px 18px"}}>
        <div style={{fontSize:12,color:"rgba(226,232,240,0.5)",marginBottom:10}}>TICKET METADATA</div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(2,minmax(0,1fr))",gap:10,fontSize:13,color:"rgba(226,232,240,.76)"}}>
          <div><b>Source:</b> {currentTicket.source || "Portal"}</div>
          <div><b>AI Status:</b> {getTicketStatusExplanation(currentTicket.status)}</div>
          <div><b>Assigned To:</b> {currentTicket.assigneeName || currentTicket.assignedTo || staffName(currentTicket.assigneeId)}</div>
          <div><b>Escalation:</b> {getEscalationInfo(currentTicket).overdue ? getEscalationInfo(currentTicket).label : "Not escalated"}</div>
          <div style={{gridColumn:"1/-1"}}><b>Watchers:</b> {(currentTicket.watchers||currentTicket.notifiedStaff||[]).map(w=>w.name).filter(Boolean).join(", ") || "All IT Staff"}</div>
        </div>
      </div>

      {/* Description */}
      <div className="glass" style={{padding:"16px 18px"}}>
        <div style={{fontSize:12,color:"rgba(226,232,240,0.5)",marginBottom:8}}>DESCRIPTION</div>
        <p style={{fontSize:14,lineHeight:1.7,color:"rgba(226,232,240,0.8)"}}>{currentTicket.description}</p>
      </div>

      {(currentTicket.aiSummary || currentTicket.suggestedAction)&&(
        <div className="glass2" style={{padding:"16px",borderColor:"rgba(125,211,252,.22)",background:"rgba(14,165,233,.07)"}}>
          <div style={{fontSize:12,color:"#7dd3fc",fontWeight:900,marginBottom:8}}>AI TICKET SUMMARY FOR STAFF</div>
          {currentTicket.aiSummary&&<div style={{fontSize:14,color:"#fff",fontWeight:800,lineHeight:1.5}}>{currentTicket.aiSummary}</div>}
          {currentTicket.suggestedAction&&<div style={{fontSize:13,color:"rgba(226,232,240,.72)",lineHeight:1.55,marginTop:7}}><b>Suggested action:</b> {currentTicket.suggestedAction}</div>}
        </div>
      )}

      {currentTicket.remoteSupportRequested&&(
        <div className="glass2" style={{padding:"16px",borderColor:"rgba(14,165,233,.24)",background:"rgba(14,165,233,.08)"}}>
          <div style={{display:"flex",justifyContent:"space-between",gap:8,flexWrap:"wrap",marginBottom:8}}>
            <div style={{fontSize:13,fontWeight:900,color:"#bae6fd"}}>Remote Support Requested</div>
            <span className="tag">{fmtDate(currentTicket.remoteSupportRequestedAt)}</span>
          </div>
          <div style={{fontSize:13,color:"rgba(226,232,240,.76)",lineHeight:1.6}}>
            <div>Tool: <b>{currentTicket.remoteSupportTool || "—"}</b></div>
            <div>Remote ID: <b>{currentTicket.remoteSupportId || "—"}</b></div>
            {currentTicket.remoteSupportNote&&<div>Note: {currentTicket.remoteSupportNote}</div>}
          </div>
        </div>
      )}

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
          <option value={0}>IT Support Team</option>
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

        applyTicketChanges(
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
    {showRemoteSupport&&(
      <Modal title="Request Remote Support" onClose={()=>setShowRemoteSupport(false)}>
        <RemoteSupportDialog ticket={currentTicket} onClose={()=>setShowRemoteSupport(false)} onSubmit={submitRemoteSupport} toast={toast} />
      </Modal>
    )}
    </>
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
function TicketCard({ticket,onView,showFeedbackPending=false,showFeedbackUnread=false}) {
  const cat=CATEGORIES.find(c=>c.id===ticket.category);
  const feedbackPending=showFeedbackPending && isTicketFeedbackPending(ticket);
  const unreadFeedback=showFeedbackUnread && ticket.feedbackStatus==="Submitted";
  const escalation=getEscalationInfo(ticket);
  if(!ticket?.id) return null;
  const assigneeLabel=ticket.assigneeName || ticket.assignedTo || staffName(ticket.assigneeId);
  return (
    <div className="glass2" style={{padding:"12px 13px",cursor:"pointer",transition:"all .2s",borderRadius:11,minHeight:128,borderColor:escalation.overdue?"rgba(239,68,68,.46)":undefined,boxShadow:escalation.overdue?"0 14px 34px rgba(239,68,68,.14),0 0 24px rgba(249,115,22,.1)":undefined}}
      onClick={()=>onView(ticket.id)}
      onMouseEnter={e=>{e.currentTarget.style.borderColor="rgba(99,102,241,0.4)";e.currentTarget.style.background="rgba(255,255,255,0.08)";}}
      onMouseLeave={e=>{e.currentTarget.style.borderColor="rgba(255,255,255,0.1)";e.currentTarget.style.background="rgba(255,255,255,0.06)";}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:10,marginBottom:8}}>
        <div style={{display:"flex",gap:9,alignItems:"center",minWidth:0}}>
          <div style={{width:31,height:31,borderRadius:9,background:cat?.bg,display:"flex",alignItems:"center",justifyContent:"center",fontSize:15,flexShrink:0}}>{cat?.icon}</div>
          <div style={{minWidth:0}}>
            <div style={{fontSize:13,fontWeight:900,color:"#f8fafc",lineHeight:1.1,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{ticket.id}</div>
            <div style={{fontSize:11,color:"rgba(226,232,240,0.58)",marginTop:3,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{ticket.name || "User"}</div>
          </div>
        </div>
        <StatusBadge status={ticket.status}/>
      </div>
      <div style={{display:"grid",gap:5,marginBottom:8,fontSize:11,color:"rgba(226,232,240,.52)"}}>
        <div style={{display:"flex",justifyContent:"space-between",gap:8}}>
          <span>Assigned To</span>
          <span style={{color:"#e2e8f0",fontWeight:800,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis",maxWidth:150}}>{assigneeLabel}</span>
        </div>
        <div style={{display:"flex",justifyContent:"space-between",gap:8}}>
          <span>TAT Time</span>
          <span style={{color:"#bae6fd",fontWeight:900}}>{isClosedTicket(ticket) && ticket.closedAt ? formatDuration(ticket.closedAt-ticket.createdAt) : timeAgo(ticket.createdAt)}</span>
        </div>
      </div>
      <div style={{display:"flex",gap:6,flexWrap:"wrap",alignItems:"center",marginBottom:8}}>
        {escalation.overdue&&<span className="tag" style={{background:"rgba(239,68,68,.17)",color:"#fecaca",fontSize:10,padding:"3px 7px"}}>Escalated L{escalation.level}</span>}
        {ticket.status==="Reopened"&&<span className="tag" style={{background:"rgba(245,158,11,.16)",color:"#fde68a",fontSize:10,padding:"3px 7px"}}>Reopened</span>}
        {(feedbackPending||unreadFeedback)&&(
          <span style={{display:"inline-flex",alignItems:"center",gap:6,border:"1px solid rgba(16,185,129,.28)",background:"rgba(16,185,129,.12)",color:"#bbf7d0",borderRadius:999,padding:"3px 7px",fontSize:10,fontWeight:900}}>
            <span className="pulse" style={{width:8,height:8,borderRadius:"50%",background:"#22c55e",display:"inline-block"}}/>
            {feedbackPending ? "Feedback Pending" : "New Feedback"}
          </span>
        )}
      </div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:8}}>
        <PriorityBadge p={ticket.priority}/>
        <span style={{fontSize:11,color:"rgba(226,232,240,0.38)",fontWeight:800}}>View details</span>
      </div>
    </div>
  );
}

// ── CATEGORY GRID ─────────────────────────────────────────────────────────
function AIHelpCards({onSmartTicket}) {
  return (
    <div style={{display:"grid",gridTemplateColumns:"minmax(0,360px)",gap:12,marginBottom:22}}>
      <button type="button" onClick={onSmartTicket} className="glass2" style={{textAlign:"left",padding:"16px 15px",border:"1px solid rgba(125,211,252,.16)",background:"linear-gradient(135deg,rgba(14,165,233,.11),rgba(139,92,246,.08),rgba(255,255,255,.04))",cursor:"pointer"}}>
        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:8}}>
          <span className="pulse" style={{width:34,height:34,borderRadius:12,display:"inline-flex",alignItems:"center",justifyContent:"center",background:"rgba(14,165,233,.15)",fontSize:18}}>✨</span>
          <span style={{fontSize:13,fontWeight:900,color:"#f8fafc"}}>Raise Smart Ticket</span>
        </div>
        <div style={{fontSize:12,lineHeight:1.45,color:"rgba(226,232,240,.58)"}}>I can create a ticket for you.</div>
      </button>
    </div>
  );
}

function SmartTicketModal({session,onSubmit,onClose,toast}) {
  const [issue,setIssue]=useState("");
  const [summary,setSummary]=useState(null);
  const [resolution,setResolution]=useState(null);
  const [ready,setReady]=useState(false);
  const [loading,setLoading]=useState(false);
  const [listening,setListening]=useState(false);
  const recognitionRef=useRef(null);

  const stopVoiceInput=useCallback(()=>{
    try {
      const recognition=recognitionRef.current;
      if(recognition) {
        recognition.onstart=null;
        recognition.onresult=null;
        recognition.onerror=null;
        recognition.onend=null;
        if(typeof recognition.abort==="function") recognition.abort();
        else if(typeof recognition.stop==="function") recognition.stop();
      }
    } catch(error) {
      console.error("Smart ticket voice stop failed:", error);
    } finally {
      recognitionRef.current=null;
      setListening(false);
    }
  },[]);

  useEffect(()=>{
    const onVisibility=()=>{ if(document.hidden) stopVoiceInput(); };
    const onBeforeUnload=()=>stopVoiceInput();
    document.addEventListener("visibilitychange",onVisibility);
    window.addEventListener("beforeunload",onBeforeUnload);
    return()=>{
      document.removeEventListener("visibilitychange",onVisibility);
      window.removeEventListener("beforeunload",onBeforeUnload);
      stopVoiceInput();
    };
  },[stopVoiceInput]);

  const startVoiceInput=()=>{
    try {
      stopVoiceInput();
      const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
      if(!SpeechRecognition) {
        toast("Voice input is not supported on this browser.","info");
        return;
      }
      const recognition = new SpeechRecognition();
      recognition.lang = "en-IN";
      recognition.interimResults = false;
      recognition.maxAlternatives = 1;
      recognitionRef.current=recognition;
      recognition.onstart = () => setListening(true);
      recognition.onerror = error => {
        console.error("Smart ticket voice input failed:", error);
        const message=error?.error==="not-allowed" || error?.error==="permission-denied"
          ? "Microphone permission was denied. Please allow mic access to use voice input."
          : "Voice input could not capture audio. Please try again.";
        toast(message,"error");
        stopVoiceInput();
      };
      recognition.onresult = event => {
        const transcript = event.results?.[0]?.[0]?.transcript || "";
        if(transcript) setIssue(prev=>`${prev ? `${prev} ` : ""}${transcript}`.trim());
      };
      recognition.onend = () => {
        recognitionRef.current=null;
        setListening(false);
      };
      recognition.start();
    } catch(error) {
      console.error("Smart ticket voice input failed:", error);
      setListening(false);
      toast("Voice input is not supported on this browser.","info");
    }
  };

  const analyze=()=>{
    if(!issue.trim()){toast("Describe your issue in simple words.","error");return;}
    const next=detectSmartTicketIssue(issue);
    setSummary(next);
    setResolution(getSmartResolution(next));
    setReady(false);
  };

  const create=async()=>{
    if(!summary) return;
    setLoading(true);
    try {
      await Promise.resolve(onSubmit({
        name:session?.name || session?.email?.split("@")[0] || "Portal User",
        email:session?.email || "",
        dept:"Not provided",
        mobile:"",
        location:"Not provided",
        category:summary.category,
        subCategory:summary.subCategory,
        priority:summary.priority,
        description:`${summary.issueSummary}\n\nAI Summary: ${summary.subCategory}\nRecommended Action: ${summary.recommendedAction}`,
        source:"AI Smart Ticket",
        assignmentGroup:getAssignmentGroup(summary.category),
        issueSummary:summary.issueSummary,
        recommendedAction:summary.recommendedAction,
        notes:`AI guided ticket created from: ${issue}`
      }));
      onClose();
    } finally {
      setLoading(false);
    }
  };

  return (
    <div style={{display:"flex",flexDirection:"column",gap:16}}>
      <div className="glass2" style={{padding:"16px 18px",background:"linear-gradient(135deg,rgba(14,165,233,.12),rgba(139,92,246,.1))"}}>
        <div style={{display:"inline-flex",gap:8,alignItems:"center",fontSize:12,fontWeight:900,color:"#7dd3fc",marginBottom:8}}><span className="pulse">✦</span> AI Guided Ticket Creation</div>
        <div style={{fontSize:18,fontWeight:900,color:"#fff",marginBottom:6}}>Describe your issue in simple words.</div>
        <div style={{fontSize:13,color:"rgba(226,232,240,.62)"}}>I can create a ticket for you after a quick fix check.</div>
      </div>
      <div style={{display:"flex",gap:8,alignItems:"stretch"}}>
        <textarea rows={3} value={issue} onChange={e=>setIssue(e.target.value)} placeholder={listening ? "Listening..." : "Example: OneJaipuria WiFi is not connecting on my laptop"} style={{resize:"vertical",flex:1}}/>
        <button type="button" className={`mic-btn ${listening ? "listening" : ""}`} onClick={startVoiceInput} title={listening ? "Listening..." : "Use voice input"} aria-label="Use voice input"><span>🎙️</span>{listening&&<span className="mic-btn-text">Listening...</span>}</button>
      </div>
      <button className="glow-btn" type="button" onClick={analyze}>Analyze with AI</button>
      {summary&&(
        <div className="glass2" style={{padding:"16px",display:"grid",gap:10}}>
          <div style={{fontWeight:900,color:"#fff"}}>Your ticket is ready. Please confirm.</div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))",gap:8,fontSize:12}}>
            <div><b>Category</b><br/>{summary.categoryLabel}</div>
            <div><b>Sub-category</b><br/>{summary.subCategory}</div>
            <div><b>Priority</b><br/>{summary.priority}</div>
            <div><b>Assign group</b><br/>{getAssignmentGroup(summary.category)}</div>
          </div>
          <div style={{fontSize:13,color:"rgba(226,232,240,.72)"}}><b>Summary:</b> {summary.issueSummary}</div>
          <div style={{fontSize:13,color:"#bfdbfe"}}><b>Recommended action:</b> {summary.recommendedAction}</div>
        </div>
      )}
      {summary&&resolution&&!ready&&(
        <div className="glass2" style={{padding:"16px",borderColor:"rgba(16,185,129,.24)"}}>
          <div style={{fontWeight:900,color:"#bbf7d0",marginBottom:8}}>Try this quick fix first.</div>
          <ol style={{margin:0,paddingLeft:20,color:"rgba(226,232,240,.76)",fontSize:13,lineHeight:1.55}}>
            {resolution.steps.slice(0,4).map((step,index)=><li key={index}>{step}</li>)}
          </ol>
          <div style={{display:"flex",gap:10,marginTop:14,flexWrap:"wrap"}}>
            <button type="button" onClick={onClose} style={{background:"rgba(16,185,129,.16)",border:"1px solid rgba(16,185,129,.28)",color:"#bbf7d0",padding:"9px 14px",borderRadius:10,fontWeight:900}}>YES, solved</button>
            <button className="glow-btn" type="button" onClick={()=>setReady(true)}>NO, create ticket</button>
          </div>
        </div>
      )}
      {ready&&<button className="glow-btn" type="button" onClick={create} disabled={loading}>{loading?"Creating...":"Confirm & Submit Ticket"}</button>}
    </div>
  );
}

function CategoryGrid({onSelect,onSmartTicket,session,showWelcome}) {
  const [hover,setHover]=useState(null);
  return (
    <div>
      <SmartWelcome session={session} visible={showWelcome} />
      <div className="glass" style={{padding:"22px",marginBottom:18,background:"radial-gradient(circle at 0 0,rgba(14,165,233,.22),transparent 36%),linear-gradient(135deg,rgba(15,23,42,.9),rgba(30,41,59,.78))"}}>
        <div style={{display:"inline-flex",alignItems:"center",gap:8,border:"1px solid rgba(125,211,252,.22)",borderRadius:999,padding:"6px 10px",color:"#7dd3fc",fontSize:12,fontWeight:900,marginBottom:12}}><span className="pulse">✦</span> AI Powered Support</div>
        <h2 style={{fontFamily:"Syne",fontSize:24,fontWeight:900,color:"#e2e8f0",marginBottom:6}}>IT Helpdesk Portal</h2>
        <p style={{fontSize:14,color:"rgba(226,232,240,0.62)",margin:0}}>Minimum effort support: describe your issue, try a smart fix, and let AI prepare the ticket.</p>
      </div>
      <AIHelpCards onSmartTicket={onSmartTicket} />
      <h2 style={{fontFamily:"Syne",fontSize:22,fontWeight:700,color:"#e2e8f0",marginBottom:6}}>Raise IT Support Ticket</h2>
      <p style={{fontSize:14,color:"rgba(226,232,240,0.5)",marginBottom:24}}>Select the issue category or let AI create a smart ticket</p>
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

function PortalFeedbackForm({session,onSubmit,toast,onClose}) {
  const roleLabel = session?.type === "admin" ? "Admin" : session?.type === "staff" ? "Staff" : session?.type === "user" ? "User" : "Guest";
  const [form,setForm]=useState({
    name: session?.name || "",
    email: session?.email || "",
    role: roleLabel,
    rating: 0,
    feedbackType: "",
    message: "",
  });
  const [loading,setLoading]=useState(false);
  const types=["Bug Report","Feature Request","UI/Design Feedback","Performance Issue","General Feedback"];
  const set=(k,v)=>setForm(f=>({...f,[k]:v}));
  const submit=async()=>{
    if(!form.name.trim() || !form.email.trim() || !form.role || !form.rating || !form.feedbackType || !form.message.trim()){
      toast("Please complete all portal feedback fields","error");
      return;
    }
    setLoading(true);
    try {
      await onSubmit({
        ...form,
        name: form.name.trim(),
        email: form.email.trim(),
        message: form.message.trim(),
      });
      onClose?.();
    } finally {
      setLoading(false);
    }
  };
  return (
    <div style={{display:"flex",flexDirection:"column",gap:18}}>
      <div className="glass2" style={{padding:"16px 18px",background:"linear-gradient(135deg,rgba(99,102,241,0.14),rgba(6,182,212,0.08))",borderColor:"rgba(125,211,252,0.24)"}}>
        <div style={{fontFamily:"Syne",fontSize:18,fontWeight:800,color:"#fff",marginBottom:4}}>Help us improve the portal</div>
        <div style={{fontSize:13,color:"rgba(226,232,240,0.58)",lineHeight:1.5}}>Share bugs, ideas, design feedback, or performance issues. This goes directly to the admin portal.</div>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(220px,1fr))",gap:14}}>
        <div><label style={{fontSize:12,color:"rgba(226,232,240,0.6)",marginBottom:6,display:"block"}}>Name</label><input value={form.name} onChange={e=>set("name",e.target.value)} placeholder="Your name" /></div>
        <div><label style={{fontSize:12,color:"rgba(226,232,240,0.6)",marginBottom:6,display:"block"}}>Email</label><input type="email" value={form.email} onChange={e=>set("email",e.target.value)} placeholder="name@jaipuria.ac.in" /></div>
        <div><label style={{fontSize:12,color:"rgba(226,232,240,0.6)",marginBottom:6,display:"block"}}>Role / User Type</label><select value={form.role} onChange={e=>set("role",e.target.value)}><option>User</option><option>Staff</option><option>Admin</option><option>Guest</option></select></div>
        <div><label style={{fontSize:12,color:"rgba(226,232,240,0.6)",marginBottom:6,display:"block"}}>Feedback Type</label><select value={form.feedbackType} onChange={e=>set("feedbackType",e.target.value)}><option value="">Select Feedback Type</option>{types.map(t=><option key={t}>{t}</option>)}</select></div>
      </div>
      <div className="glass2" style={{padding:"16px 18px"}}>
        <div style={{fontSize:13,fontWeight:700,color:"rgba(226,232,240,0.74)",marginBottom:10}}>Rating</div>
        <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
          {[1,2,3,4,5].map(n=><button key={n} type="button" onClick={()=>set("rating",n)} style={{background:"transparent",border:"none",fontSize:32,lineHeight:1,color:n<=form.rating?"#fbbf24":"rgba(255,255,255,0.18)",filter:n<=form.rating?"drop-shadow(0 0 10px rgba(251,191,36,0.35))":"none",transform:n===form.rating?"translateY(-2px) scale(1.08)":"none"}}>★</button>)}
          <span style={{fontSize:13,color:"rgba(226,232,240,0.48)"}}>{form.rating ? `${form.rating}/5` : "Select rating"}</span>
        </div>
      </div>
      <div><label style={{fontSize:12,color:"rgba(226,232,240,0.6)",marginBottom:6,display:"block"}}>Message / Feedback</label><textarea rows={5} value={form.message} onChange={e=>set("message",e.target.value)} placeholder="Tell us what should be fixed or improved..." style={{resize:"vertical"}} /></div>
      <div style={{display:"flex",justifyContent:"flex-end",gap:10,flexWrap:"wrap"}}>
        <button onClick={onClose} style={{background:"rgba(255,255,255,0.06)",border:"1px solid rgba(255,255,255,0.1)",color:"#e2e8f0",padding:"10px 18px",borderRadius:10}}>Cancel</button>
        <button className="glow-btn" onClick={submit} disabled={loading}>{loading?"Submitting...":"Submit Feedback"}</button>
      </div>
    </div>
  );
}

function AdminPortalFeedbackPage({portalFeedback,setPortalFeedback,toast}) {
  const [statusFilter,setStatusFilter]=useState("All");
  const [typeFilter,setTypeFilter]=useState("All");
  const [ratingFilter,setRatingFilter]=useState("All");
  const types=["Bug Report","Feature Request","UI/Design Feedback","Performance Issue","General Feedback"];
  const filtered=portalFeedback.filter(f=>{
    if(statusFilter!=="All" && f.status!==statusFilter) return false;
    if(typeFilter!=="All" && f.feedbackType!==typeFilter) return false;
    if(ratingFilter!=="All" && Number(f.rating)!==Number(ratingFilter)) return false;
    return true;
  });
  const avg=filtered.length?(filtered.reduce((a,f)=>a+Number(f.rating||0),0)/filtered.length).toFixed(1):"0.0";
  const markReviewed=async(id)=>{
    const current=portalFeedback.find(f=>f.id===id);
    if(!current) return;
    const updated={...current,status:"Reviewed",reviewed:true,reviewedAt:Date.now()};
    try{
      await updatePortalFeedback(updated);
      setPortalFeedback(fs=>fs.map(f=>f.id===id?updated:f));
      toast("Portal feedback marked reviewed","success");
    }catch(error){
      console.error("Portal feedback review failed:",error);
      toast("Portal feedback update failed","error");
    }
  };
  const exportRows=filtered.map(f=>({"Feedback ID":f.id,Name:f.name,Email:f.email,Role:f.role,Rating:f.rating,"Feedback Type":f.feedbackType,Message:f.message,"Created At":fmtDate(f.createdAt),Status:f.status,Reviewed:f.reviewed?"Yes":"No"}));
  return (
    <div style={{display:"flex",flexDirection:"column",gap:22}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:14,flexWrap:"wrap"}}>
        <div><h2 style={{fontFamily:"Syne",fontSize:24,fontWeight:800,color:"#e2e8f0"}}>Portal Feedback</h2><p style={{fontSize:14,color:"rgba(226,232,240,0.5)",marginTop:4}}>Review portal bugs, feature ideas, UI feedback, and performance reports.</p></div>
        <button className="glow-btn" onClick={()=>downloadExcel(exportRows,`portal_feedback_${new Date().toISOString().slice(0,10)}.xlsx`)}>Export Excel</button>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(155px,1fr))",gap:14}}>
        <StatCard label="Total" value={portalFeedback.length} icon="★" color="#818cf8" />
        <StatCard label="New" value={portalFeedback.filter(f=>f.status==="New"&&!f.reviewed).length} icon="●" color="#fbbf24" />
        <StatCard label="Reviewed" value={portalFeedback.filter(f=>f.reviewed).length} icon="✓" color="#34d399" />
        <StatCard label="Avg Rating" value={`${avg}/5`} icon="★" color="#f59e0b" />
      </div>
      <div className="glass" style={{padding:"18px 20px",display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(170px,1fr))",gap:14}}>
        <select value={statusFilter} onChange={e=>setStatusFilter(e.target.value)}><option value="All">All Status</option><option>New</option><option>Reviewed</option></select>
        <select value={typeFilter} onChange={e=>setTypeFilter(e.target.value)}><option value="All">All Types</option>{types.map(t=><option key={t}>{t}</option>)}</select>
        <select value={ratingFilter} onChange={e=>setRatingFilter(e.target.value)}><option value="All">All Ratings</option>{[5,4,3,2,1].map(r=><option key={r} value={r}>{r} Star</option>)}</select>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(300px,1fr))",gap:14}}>
        {filtered.map(f=><div key={f.id} className="glass2" style={{padding:"18px 16px",borderColor:f.reviewed?"rgba(16,185,129,0.24)":"rgba(245,158,11,0.28)"}}>
          <div style={{display:"flex",justifyContent:"space-between",gap:10,alignItems:"flex-start",marginBottom:10}}><div><div style={{fontSize:14,fontWeight:800,color:"#fff"}}>{f.feedbackType}</div><div style={{fontSize:12,color:"rgba(226,232,240,0.44)",marginTop:3}}>{f.name} · {f.role}</div></div><span className="tag" style={{background:f.reviewed?"rgba(16,185,129,0.14)":"rgba(245,158,11,0.14)",color:f.reviewed?"#6ee7b7":"#fbbf24"}}>{f.status}</span></div>
          <div style={{fontSize:13,color:"rgba(226,232,240,0.7)",lineHeight:1.55,marginBottom:10}}>{f.message}</div>
          <div style={{fontSize:12,color:"rgba(226,232,240,0.45)",display:"flex",justifyContent:"space-between",gap:8,flexWrap:"wrap"}}><span>{f.email}</span><span>{f.rating}/5 · {fmtDate(f.createdAt)}</span></div>
          {!f.reviewed&&<button className="glow-btn" style={{width:"100%",marginTop:14,padding:"9px 12px",fontSize:13}} onClick={()=>markReviewed(f.id)}>Mark Reviewed</button>}
        </div>)}
        {filtered.length===0&&<div className="glass" style={{gridColumn:"1/-1",padding:40,textAlign:"center",color:"rgba(226,232,240,0.42)"}}>No portal feedback found.</div>}
      </div>
    </div>
  );
}

const HELPDESK_MENU_ITEMS=[
  {id:"wifi",label:"WiFi",aliases:["wifi","wi-fi","wifi issue","wi fi","onejaipuria","one jaipuria"]},
  {id:"printer",label:"Printer",aliases:["printer","printer issue","printing","print"]},
  {id:"moodle",label:"Moodle / LMS",aliases:["moodle","lms","moodle issue","lms issue","learning management system"]},
  {id:"internet",label:"Internet",aliases:["internet","internet issue","network","browsing"]},
  {id:"laptop",label:"Laptop",aliases:["laptop","laptop issue","notebook"]},
  {id:"desktop",label:"Desktop",aliases:["desktop","desktop issue","pc","computer"]},
  {id:"ms-office",label:"MS Office",aliases:["ms office","office","microsoft office","word","excel","powerpoint"]},
  {id:"email",label:"Email",aliases:["email","mail","outlook","email issue"]},
  {id:"login",label:"Login Issue",aliases:["login","login issue","password","signin","sign in","account"]},
  {id:"echo360",label:"Echo360 Lecture Capture System",aliases:["echo360","echo 360","lecture capture","recording issue","classroom recording","lecture not uploaded","capture device","lecture playback"]},
  {id:"resources",label:"Other IT Resources",aliases:["other it resources","it resources","resources","other"]},
  {id:"ai",label:"Talk to AI",aliases:["ai","ask ai","talk to ai","assistant"]}
];

function getHelpdeskMenu() {
  return HELPDESK_MENU_ITEMS.map((item,index)=>({number:index+1,id:item.id,label:item.label}));
}

const HELPDESK_SUBCATEGORIES={
  wifi:["WiFi Network not showing","Not able to connect OneJaipuria","OneJaipuria connected but internet not working","Guest WiFi password for guest user","WiFi certificate/configuration issue","WiFi slow","Mobile connected but laptop not connecting","Other WiFi issue"],
  printer:["Need to install printer driver","Printer showing offline","Print command given but print not coming","Printer paper jam","Printer toner/ink issue","Scan not working","Printer not visible in system","Wrong printer selected","Other printer issue"],
  moodle:["Moodle login issue","Course not visible","Assignment upload issue","Quiz/test not opening","Password reset issue","File download issue","Attendance/course material not visible","Other Moodle issue"],
  internet:["Internet not working","Internet slow","LAN cable connected but no internet","Website not opening","Only some websites not opening","IP conflict / network issue","Other internet issue"],
  laptop:["Laptop login password forgot","Laptop slow","Laptop not turning on","Battery/charging issue","Keyboard/mouse not working","WiFi not working on laptop","Software installation required","Windows update issue","Other laptop issue"],
  desktop:["Desktop login password forgot","Desktop not turning on","Monitor no display","Keyboard/mouse not working","Desktop slow","LAN/internet issue","Software installation required","Other desktop issue"],
  "ms-office":["Office license expired","Word/Excel/PowerPoint not opening","Outlook login issue","Teams login issue","OneDrive sync issue","Office activation issue","Other MS Office issue"],
  email:["Email login issue","Password reset issue","Email not sending","Email not receiving","Outlook configuration issue","Mailbox full","Attachment issue","Other email issue"],
  login:["Portal login issue","Laptop/Desktop login password forgot","Moodle login issue","Email login issue","MS Teams login issue","Other login issue"],
  echo360:["Lecture recording not starting","Lecture recording missing","Audio not recording","Video not recording","Echo360 login issue","Lecture not uploaded","Classroom capture device offline","Recording quality issue","Echo360 playback issue","Faculty access issue","Student unable to view recording","Other Echo360 issue"],
  resources:["Projector issue","Smart classroom issue","Biometric issue","CCTV request","Software access request","New system/peripheral request","Other IT request"]
};

const wifiKnowledgeBase={
  category:"WiFi",
  ticketCategory:"wifi",
  priority:"Medium",
  title:"OneJaipuria WiFi not connecting",
  keywords:[
    "wifi not working",
    "onejaipuria issue",
    "onejaipuria wifi",
    "internet not connecting",
    "laptop wifi issue",
    "mobile wifi issue",
    "wifi setup",
    "wifi configuration",
    "wifi certificate",
    "campus wifi"
  ],
  steps:[
    "First connect your laptop/mobile with your mobile hotspot.",
    "Open your default browser such as Chrome, Edge, or Firefox.",
    "Open: https://tinyurl.com/jimjwifi",
    "Click on \"Join Now\" and download the configuration file.",
    "Run the downloaded file.",
    "Click \"Next\".",
    "Select your Jaipuria email account.",
    "Complete setup.",
    "Connect to OneJaipuria WiFi.",
    "Restart WiFi if required."
  ],
  commonCauses:[
    "Old or missing WiFi certificate",
    "Wrong Jaipuria email selected",
    "Browser blocked download",
    "WiFi adapter disabled",
    "Old WiFi profile conflict"
  ],
  url:"https://tinyurl.com/jimjwifi"
};

function detectWifiIssue(userText) {
  const text=String(userText || "").toLowerCase().replace(/\s+/g," ").trim();
  return wifiKnowledgeBase.keywords.some(keyword=>text.includes(keyword));
}

function getWifiTroubleshooting(subCategory="Not able to connect OneJaipuria") {
  return {
    type:"steps",
    categoryId:"wifi",
    categoryLabel:wifiKnowledgeBase.category,
    subCategory,
    title:subCategory,
    steps:wifiKnowledgeBase.steps,
    commonCauses:wifiKnowledgeBase.commonCauses,
    url:wifiKnowledgeBase.url,
    footer:"Is your issue resolved? Type YES or NO."
  };
}

function handleWifiTicketFlow(subCategory="Not able to connect OneJaipuria") {
  return {
    categoryId:"wifi",
    categoryLabel:wifiKnowledgeBase.category,
    subCategory,
    notes:`AI Chatbot troubleshooting shown:\n${subCategory}\n${wifiKnowledgeBase.steps.map((step,index)=>`${index+1}. ${step}`).join("\n")}\n\nCommon Causes:\n${wifiKnowledgeBase.commonCauses.map(cause=>`- ${cause}`).join("\n")}`,
    description:`${subCategory}\n\nTroubleshooting shown:\n${wifiKnowledgeBase.steps.map((step,index)=>`${index+1}. ${step}`).join("\n")}`
  };
}

function getWifiSteps() {
  return getWifiTroubleshooting();
}

function getCategorySteps(category) {
  const normalized=String(category || "").toLowerCase();
  const categoryMap={
    printer:{title:"Printer troubleshooting",steps:["Confirm the printer is powered on and showing ready.","Check that your laptop/desktop is connected to the campus network.","Restart the printer queue or try printing a test page.","If paper, toner, or access errors appear, note the exact message."]},
    moodle:{title:"Moodle / LMS troubleshooting",steps:["Check your internet connection and reopen Moodle/LMS in a fresh browser tab.","Clear browser cache or try Chrome, Edge, or Firefox.","Verify that you are using your Jaipuria email/login credentials.","Capture the error message or course name if the issue continues."]},
    internet:{title:"Internet troubleshooting",steps:["Check whether WiFi or LAN is connected.","Restart WiFi or unplug/replug the LAN cable.","Try opening another website in a fresh browser window.","Restart the device if the network shows connected but pages do not load."]},
    laptop:{title:"Laptop troubleshooting",steps:["Restart the laptop and check whether the issue repeats.","Confirm charger, battery, keyboard, touchpad, and display behavior.","Run pending OS updates only if you have enough time and battery.","Note any error message, noise, heating, or application that triggers the issue."]},
    desktop:{title:"Desktop troubleshooting",steps:["Check power cable, monitor cable, keyboard, mouse, and LAN cable.","Restart the system once if it is responsive.","Confirm whether the issue is with display, login, internet, or a specific application.","Note the system location and any error shown on screen."]},
    "ms-office":{title:"MS Office troubleshooting",steps:["Close and reopen Word, Excel, PowerPoint, or Outlook.","Check whether your Jaipuria account is signed in and licensed.","Try opening the file from local storage if it fails from email/cloud.","Restart the device if Office keeps freezing or asking for activation."]},
    email:{title:"Email troubleshooting",steps:["Check internet connectivity and open email in a browser.","Confirm you are using your Jaipuria email address.","Reset browser cache or try another browser/device.","Note any Outlook, password, MFA, or mailbox error message."]},
    login:{title:"Login issue troubleshooting",steps:["Confirm the username/email is typed correctly.","Check Caps Lock and try resetting the password if available.","Try signing in from a different browser or private window.","Capture the exact login error for IT Support."]},
    echo360:{title:"Echo360 Lecture Capture System troubleshooting",steps:["Confirm the classroom capture device is powered on and connected.","Check the lecture schedule, room mapping, and faculty access.","Restart the Echo360 capture device or playback browser once if available.","Note the lecture name, classroom, date, time, and exact recording/playback issue."]},
    resources:{title:"Other IT Resources",steps:["Identify the exact resource, portal, software, or device you need help with.","Check whether the issue happens on one device or multiple devices.","Restart the app/browser and try again once.","Keep your system/location and error details ready for support."]}
  };
  const fallback={title:"IT troubleshooting",steps:["Restart the device or application once.","Check internet connectivity and login status.","Note the exact error message and when it occurs.","Try again from another browser or device if available."]};
  const flow=categoryMap[normalized] || fallback;
  return {
    ...flow,
    footer:"Is your issue resolved? Type YES or NO."
  };
}

function getTicketCategoryFromHelpdesk(categoryId) {
  const categoryMap={
    wifi:"wifi",
    printer:"printer",
    moodle:"erp",
    internet:"internet",
    laptop:"laptop",
    desktop:"desktop",
    "ms-office":"software",
    email:"email",
    login:"password",
    echo360:"echo360",
    resources:"other"
  };
  return categoryMap[categoryId] || "other";
}

function getHelpdeskCategoryLabel(categoryId) {
  return HELPDESK_MENU_ITEMS.find(item=>item.id===categoryId)?.label || "Other IT Resources";
}

function normalizeHelpdeskText(value) {
  return String(value || "").toLowerCase().replace(/[?.!]+$/g,"").replace(/\s+/g," ").trim();
}

function getSubCategoryMenu(categoryId) {
  const categoryLabel=getHelpdeskCategoryLabel(categoryId);
  return {
    type:"menu",
    menuType:"sub",
    categoryId,
    text:`${categoryLabel} issue type:`,
    hint:"Type back to return to categories.",
    menu:(HELPDESK_SUBCATEGORIES[categoryId] || []).map((label,index)=>({number:index+1,id:`${categoryId}-${index+1}`,label}))
  };
}

function findMainCategorySelection(userText) {
  const text=normalizeHelpdeskText(userText);
  const number=Number(text);
  if(Number.isInteger(number)) return HELPDESK_MENU_ITEMS[number-1] || null;
  return HELPDESK_MENU_ITEMS.find(item=>item.aliases.includes(text) || item.label.toLowerCase()===text) || null;
}

function findSubCategorySelection(userText, categoryId=null) {
  const text=normalizeHelpdeskText(userText);
  const categories=categoryId ? [categoryId] : Object.keys(HELPDESK_SUBCATEGORIES);
  for(const currentCategory of categories) {
    const options=HELPDESK_SUBCATEGORIES[currentCategory] || [];
    const number=Number(text);
    const selectedByNumber=Number.isInteger(number) && categoryId ? options[number-1] : null;
    const selected=selectedByNumber || options.find(label=>{
      const normalized=label.toLowerCase();
      return normalized===text || normalized.includes(text) || text.includes(normalized);
    });
    if(selected) return {categoryId:currentCategory,categoryLabel:getHelpdeskCategoryLabel(currentCategory),subCategory:selected};
  }
  return null;
}

function getBasicTroubleshootingSteps(categoryLabel, subCategory) {
  const echo360Steps={
    "Lecture recording not starting":["Check classroom capture device power.","Verify internet connection in the classroom.","Restart the Echo360 capture device.","Check lecture schedule mapping.","Try manual recording start."],
    "Lecture recording missing":["Confirm the correct course, section, and lecture date in Echo360.","Check whether the recording is still processing or unpublished.","Verify the classroom and schedule mapping.","Ask faculty to refresh the Echo360 course page.","Note the lecture date, time, classroom, and course code for IT Support."],
    "Audio not recording":["Check classroom microphone power and mute status.","Verify the correct audio input is selected on the capture device.","Restart the capture device if audio is not detected.","Run a short test recording if available.","Report the classroom, microphone type, and lecture time if issue continues."],
    "Video not recording":["Check camera power and cable connection.","Confirm the camera lens is not blocked and the source is selected.","Restart the Echo360 capture device.","Verify the classroom capture profile includes video.","Capture the room, date, and schedule details for IT Support."],
    "Echo360 login issue":["Confirm you are using your Jaipuria email account.","Try signing in from Chrome, Edge, or Firefox.","Clear browser cache or use a private window.","Check whether Moodle/LMS access opens the correct Echo360 link.","Share the exact login error if it continues."],
    "Lecture not uploaded":["Check whether the recording is still processing.","Verify internet connectivity from the classroom capture device.","Confirm the schedule and course mapping are correct.","Refresh the Echo360 course section after a few minutes.","Report lecture date, time, room, and course details if upload is delayed."],
    "Classroom capture device offline":["Check device power and network cable.","Restart the classroom capture device if accessible.","Confirm classroom internet is working.","Check whether only one room or multiple rooms are affected.","Escalate with classroom number and device status."],
    "Recording quality issue":["Check microphone distance, camera view, and classroom lighting.","Confirm there is no background noise near the microphone.","Restart the capture device before the next recording.","Try playback in another browser to rule out streaming issues.","Share a sample lecture link and timestamp with IT Support."],
    "Echo360 playback issue":["Refresh the page and try another browser.","Check internet speed and disable VPN/proxy if used.","Clear browser cache or try private window.","Confirm the recording is published and visible to your role.","Share the recording link and error message if playback still fails."],
    "Faculty access issue":["Confirm the faculty member is mapped to the correct course.","Check access through Moodle/LMS and Echo360 directly.","Sign out and sign back in with Jaipuria email.","Ask the course admin to verify role/access mapping.","Share faculty email, course, and section details if unresolved."],
    "Student unable to view recording":["Confirm the recording is published for students.","Check whether the student is enrolled in the correct course/section.","Ask the student to try another browser or private window.","Verify the Echo360 link opens from Moodle/LMS.","Share student email, course, and recording link if unresolved."],
    "Other Echo360 issue":["Identify the exact Echo360 feature or classroom device affected.","Try refreshing the browser or restarting the capture/playback device once.","Check whether the issue affects one user, one classroom, or multiple users.","Keep course, room, lecture date, and screenshot details ready.","Escalate to IT Support if the issue continues."]
  };
  if(categoryLabel==="Echo360 Lecture Capture System" || echo360Steps[subCategory]) {
    return echo360Steps[subCategory] || echo360Steps["Other Echo360 issue"];
  }
  return [
    `Confirm the exact ${categoryLabel} issue: ${subCategory}.`,
    "Restart the related app, device, or connection once.",
    "Check login, cable, network, power, or account status as applicable.",
    "Try again from another browser/device if available.",
    "Keep a screenshot or exact error message ready for IT Support."
  ];
}

function getSubCategoryTroubleshooting(categoryId, subCategory) {
  if(categoryId==="wifi" && subCategory==="Not able to connect OneJaipuria") return getWifiTroubleshooting(subCategory);
  const categoryLabel=getHelpdeskCategoryLabel(categoryId);
  return {
    type:"steps",
    categoryId,
    categoryLabel,
    subCategory,
    title:subCategory,
    steps:getBasicTroubleshootingSteps(categoryLabel, subCategory),
    footer:"Is your issue resolved? Type YES or NO."
  };
}

function getTicketContextFromTroubleshooting(reply) {
  if(reply.categoryId==="wifi" && reply.subCategory==="Not able to connect OneJaipuria") return handleWifiTicketFlow(reply.subCategory);
  return {
    categoryId:reply.categoryId,
    categoryLabel:reply.categoryLabel,
    subCategory:reply.subCategory,
    notes:`AI Chatbot troubleshooting shown:\n${reply.title}\n${reply.steps.map((step,index)=>`${index+1}. ${step}`).join("\n")}`,
    description:`${reply.subCategory}\n\nTroubleshooting shown:\n${reply.steps.map((step,index)=>`${index+1}. ${step}`).join("\n")}`
  };
}

function detectSmartTicketIssue(text) {
  const q=normalizeHelpdeskText(text);
  const rules=[
    {category:"wifi",subCategory:"Not able to connect OneJaipuria",priority:"Medium",words:["wifi","wi-fi","onejaipuria","certificate","campus wifi"]},
    {category:"printer",subCategory:"Printer showing offline",priority:"Medium",words:["printer","print","scan","toner","paper jam"]},
    {category:"erp",label:"Moodle / LMS",subCategory:"Moodle login issue",priority:"Medium",words:["moodle","lms","course","assignment","quiz"]},
    {category:"echo360",subCategory:"Lecture not uploaded",priority:"Medium",words:["lecture not uploaded","recording not uploaded","upload missing"]},
    {category:"echo360",subCategory:"Classroom capture device offline",priority:"Medium",words:["capture device","device offline","classroom capture device"]},
    {category:"echo360",subCategory:"Echo360 playback issue",priority:"Medium",words:["lecture playback","playback issue","recording playback"]},
    {category:"echo360",subCategory:"Lecture recording not starting",priority:"Medium",words:["echo360","echo 360","lecture capture","recording issue","classroom recording"]},
    {category:"internet",subCategory:"Internet not working",priority:"High",words:["internet","lan","website","network","ip conflict"]},
    {category:"laptop",subCategory:"Other laptop issue",priority:"Medium",words:["laptop","battery","charging","keyboard","windows"]},
    {category:"desktop",subCategory:"Other desktop issue",priority:"Medium",words:["desktop","monitor","cpu","mouse"]},
    {category:"software",label:"MS Office",subCategory:"Other MS Office issue",priority:"Medium",words:["office","word","excel","powerpoint","teams","onedrive","outlook"]},
    {category:"email",subCategory:"Email login issue",priority:"Medium",words:["email","mail","mailbox","attachment"]},
    {category:"password",label:"Login Issue",subCategory:"Portal login issue",priority:"High",words:["login","password","signin","sign in","forgot"]},
    {category:"resources",subCategory:"Other IT request",priority:"Medium",words:["projector","biometric","cctv","software access","new system","peripheral"]}
  ];
  const matched=rules.find(rule=>rule.words.some(word=>q.includes(word))) || rules[rules.length-1];
  const friendlyCategoryLabel=matched.label || categoryLabel(matched.category);
  const issueSummary=String(text || "").trim() || matched.subCategory;
  const recommendedAction=matched.category==="wifi"
    ? "Try the OneJaipuria configuration setup and restart WiFi once."
    : `Try a quick restart/check for ${friendlyCategoryLabel}; if it continues, submit this AI-prepared ticket.`;
  return {
    category:matched.category,
    categoryLabel:friendlyCategoryLabel,
    subCategory:matched.subCategory,
    priority:matched.priority,
    issueSummary,
    recommendedAction
  };
}

function getSmartResolution(ticketSummary) {
  if(ticketSummary.category==="wifi") return getWifiTroubleshooting(ticketSummary.subCategory);
  return {
    type:"steps",
    categoryId:ticketSummary.category,
    categoryLabel:ticketSummary.categoryLabel,
    subCategory:ticketSummary.subCategory,
    title:ticketSummary.subCategory,
    steps:getBasicTroubleshootingSteps(ticketSummary.categoryLabel, ticketSummary.subCategory),
    footer:"Did this solve your issue? YES / NO"
  };
}

function getAssignmentGroup(category) {
  if(["wifi","internet"].includes(category)) return "Network Support";
  if(["printer","desktop","laptop"].includes(category)) return "Hardware Support";
  if(["moodle","erp","login","password","email","ms-office","software"].includes(category)) return "Application Support";
  return "IT Support Team";
}

function getTicketStatusExplanation(status) {
  const text={
    Open:"Your request is received and waiting for IT review.",
    Assigned:"Your ticket has been assigned to the right IT support team.",
    "In Progress":"IT support is actively working on your issue.",
    Resolved:"IT has marked this issue resolved. Please verify once.",
    Closed:"This ticket is closed after resolution."
  };
  return text[status] || "Your ticket is being tracked by IT Support.";
}

function getTicketFlowPrompt(step, draft={}) {
  const prompts={
    name:"Please enter your name.",
    email:"Please enter your email.",
    mobile:"Please enter your mobile number."
  };
  return prompts[step] || "";
}

function handleMenuSelection(userText, activeCategoryId=null) {
  const raw=String(userText || "").trim();
  const text=normalizeHelpdeskText(raw);
  if(!text) return null;
  if(["hi","hello","hey","help","menu","start"].includes(text)) {
    return {type:"menu",text:"Please choose an IT helpdesk option:",menu:getHelpdeskMenu()};
  }
  if(["escalate","esc","talk to it","talk to support","it support","support"].includes(text)) {
    return {type:"escalate",text:"Your issue has been marked for IT Support escalation. Please share your name, system/location, and issue details."};
  }
  const subCategory=activeCategoryId ? findSubCategorySelection(text, activeCategoryId) : findSubCategorySelection(text);
  if(subCategory) return getSubCategoryTroubleshooting(subCategory.categoryId, subCategory.subCategory);
  const selected=findMainCategorySelection(text);
  if(!selected) return null;
  if(selected.id==="ai") return {type:"ai",aiPrompt:"Talk to AI"};
  return getSubCategoryMenu(selected.id);
}

function isMainMenuResponse(response) {
  return response?.type==="menu" && !response.menuType;
}

function isSubMenuResponse(response) {
  return response?.type==="menu" && response.menuType==="sub";
}

function findCategoryFromSubMenu(response) {
  if(isSubMenuResponse(response)) return response.categoryId;
  return null;
}

function AIHelpdeskChat({session,onCreateTicket}) {
  const [open,setOpen]=useState(false);
  const [input,setInput]=useState("");
  const [loading,setLoading]=useState(false);
  const [ticketFlow,setTicketFlow]=useState(null);
  const [activeCategoryId,setActiveCategoryId]=useState(null);
  const [lastHelpdeskContext,setLastHelpdeskContext]=useState(null);
  const [listening,setListening]=useState(false);
  const [messages,setMessages]=useState([
    {id:"welcome",role:"assistant",text:"Hello! How may I help you today?",at:Date.now()}
  ]);
  const endRef=useRef(null);
  const recognitionRef=useRef(null);

  useEffect(()=>{
    if(open) endRef.current?.scrollIntoView({behavior:"smooth",block:"end"});
  },[messages.length,loading,open]);

  const buildPrompt=(question)=>`You are Jaipuria Helpdesk AI, a friendly IT support assistant for Jaipuria Institute of Management. Answer only campus IT/helpdesk questions such as login issues, WiFi, Moodle/LMS, Echo360 lecture capture, printers, MS Office, email, laptop/software troubleshooting, and general IT support. Give concise, practical steps. If the question cannot be answered confidently, reply exactly: I have forwarded this issue to IT Support Team.\n\nUser: ${question}`;

  const stopVoiceInput=useCallback(()=>{
    try {
      const recognition=recognitionRef.current;
      if(recognition) {
        recognition.onstart=null;
        recognition.onresult=null;
        recognition.onerror=null;
        recognition.onend=null;
        if(typeof recognition.abort==="function") recognition.abort();
        else if(typeof recognition.stop==="function") recognition.stop();
      }
    } catch(error) {
      console.error("Chatbot voice stop failed:", error);
    } finally {
      recognitionRef.current=null;
      setListening(false);
    }
  },[]);

  useEffect(()=>{
    if(!open) stopVoiceInput();
  },[open,stopVoiceInput]);

  useEffect(()=>{
    const onVisibility=()=>{ if(document.hidden) stopVoiceInput(); };
    const onBeforeUnload=()=>stopVoiceInput();
    window.addEventListener("pagehide",onBeforeUnload);
    window.addEventListener("beforeunload",onBeforeUnload);
    document.addEventListener("visibilitychange",onVisibility);
    return()=>{
      window.removeEventListener("pagehide",onBeforeUnload);
      window.removeEventListener("beforeunload",onBeforeUnload);
      document.removeEventListener("visibilitychange",onVisibility);
      stopVoiceInput();
    };
  },[stopVoiceInput]);

  const startVoiceInput=()=>{
    try {
      stopVoiceInput();
      const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
      if(!SpeechRecognition) {
        setMessages(prev=>[...prev,{id:genToken(),role:"assistant",text:"Voice input is not supported on this browser.",at:Date.now(),error:true}]);
        return;
      }
      const recognition = new SpeechRecognition();
      recognition.lang = "en-IN";
      recognition.interimResults = false;
      recognition.maxAlternatives = 1;
      recognitionRef.current=recognition;
      recognition.onstart = () => setListening(true);
      recognition.onerror = error => {
        console.error("Chatbot voice input failed:", error);
        const message=error?.error==="not-allowed" || error?.error==="permission-denied"
          ? "Microphone permission was denied. Please allow mic access to use voice input."
          : "Voice input could not capture audio. Please try again.";
        setMessages(prev=>[...prev,{id:genToken(),role:"assistant",text:message,at:Date.now(),error:true}]);
        stopVoiceInput();
      };
      recognition.onresult = event => {
        const transcript = event.results?.[0]?.[0]?.transcript || "";
        if(transcript) setInput(prev=>`${prev ? `${prev} ` : ""}${transcript}`.trim());
      };
      recognition.onend = () => {
        recognitionRef.current=null;
        setListening(false);
      };
      recognition.start();
    } catch(error) {
      console.error("Chatbot voice input failed:", error);
      setListening(false);
      setMessages(prev=>[...prev,{id:genToken(),role:"assistant",text:"Voice input is not supported on this browser.",at:Date.now(),error:true}]);
    }
  };

  const startTicketFlow=(context=lastHelpdeskContext)=>{
    const categoryId=context?.categoryId || "resources";
    const categoryLabel=context?.categoryLabel || getHelpdeskCategoryLabel(categoryId);
    const ticketCategory=getTicketCategoryFromHelpdesk(categoryId);
    const draft={
      name:"",
      email:"",
      mobile:"",
      dept:"Not provided",
      location:"Not provided",
      category:ticketCategory,
      categoryLabel,
      subCategory:context?.subCategory || categoryLabel,
      description:context?.description || `${categoryLabel} issue reported from AI Chatbot.\n\n${context?.notes || "Troubleshooting context was not available."}`,
      priority:"Medium",
      source:"AI Chatbot",
      notes:context?.notes || ""
    };
    const firstStep="name";
    setTicketFlow({step:firstStep,draft});
    setMessages(prev=>[...prev,{id:genToken(),role:"assistant",text:getTicketFlowPrompt(firstStep,draft),at:Date.now(),type:"ticket-flow"}]);
  };

  const handleTicketFlowInput=async(clean)=>{
    if(!ticketFlow) return false;
    const step=ticketFlow.step;
    const value=clean.trim();
    const draft={...ticketFlow.draft};
    if(step==="email" && !/\S+@\S+\.\S+/.test(value)) {
      setMessages(prev=>[...prev,{id:genToken(),role:"assistant",text:"Please enter a valid email address.",at:Date.now(),error:true}]);
      return true;
    }
    draft[step]=value;
    const steps=["name","email","mobile"];
    const nextStep=steps[steps.indexOf(step)+1];
    if(nextStep) {
      setTicketFlow({step:nextStep,draft});
      setMessages(prev=>[...prev,{id:genToken(),role:"assistant",text:getTicketFlowPrompt(nextStep,draft),at:Date.now(),type:"ticket-flow"}]);
      return true;
    }
    setTicketFlow(null);
    setLoading(true);
    try {
      const ticket=await Promise.resolve(onCreateTicket?.({
        name:draft.name,
        email:draft.email,
        dept:"Not provided",
        mobile:draft.mobile,
        location:"Not provided",
        category:draft.category,
        subCategory:draft.subCategory,
        description:draft.description,
        priority:"Medium",
        source:"AI Chatbot",
        notes:draft.notes
      }));
      if(!ticket?.id) throw new Error("Ticket could not be created.");
      setMessages(prev=>[...prev,{id:genToken(),role:"assistant",text:`Your ticket has been generated successfully. Ticket ID: ${ticket.id}. Our IT Support Team will contact you soon.`,at:Date.now(),type:"ticket-success"}]);
    } catch (error) {
      console.error("Chatbot ticket creation failed:",error);
      setMessages(prev=>[...prev,{id:genToken(),role:"assistant",text:"Ticket could not be created right now. Please try again or contact IT Support.",at:Date.now(),error:true}]);
    } finally {
      setLoading(false);
    }
    return true;
  };

  const sendMessage=async(value)=>{
    const clean=(value ?? input).trim();
    if(!clean || loading) return;
    const userMessage={id:genToken(),role:"user",text:clean,at:Date.now()};
    setMessages(prev=>[...prev,userMessage]);
    setInput("");
    if(ticketFlow && await handleTicketFlowInput(clean)) return;
    const normalized=clean.toLowerCase().trim();
    if(normalized==="yes") {
      setMessages(prev=>[...prev,{id:genToken(),role:"assistant",text:"Great! Happy to help.",at:Date.now()}]);
      return;
    }
    if(normalized==="menu") {
      setActiveCategoryId(null);
      setMessages(prev=>[...prev,{id:genToken(),role:"assistant",at:Date.now(),type:"menu",text:"Please choose an IT helpdesk option:",menu:getHelpdeskMenu()}]);
      return;
    }
    if(normalized==="back") {
      if(activeCategoryId) {
        setActiveCategoryId(null);
        setMessages(prev=>[...prev,{id:genToken(),role:"assistant",at:Date.now(),type:"menu",text:"Please choose an IT helpdesk option:",menu:getHelpdeskMenu()}]);
      } else {
        setMessages(prev=>[...prev,{id:genToken(),role:"assistant",text:"You are already at the main menu. Type menu to view categories.",at:Date.now()}]);
      }
      return;
    }
    if(normalized==="no" || ["escalate","esc","talk to it","talk to support","it support","support"].includes(normalized)) {
      startTicketFlow();
      return;
    }
    if(detectWifiIssue(clean)) {
      const wifiReply=getWifiTroubleshooting();
      setLastHelpdeskContext(handleWifiTicketFlow());
      setActiveCategoryId(null);
      setMessages(prev=>[...prev,{id:genToken(),role:"assistant",at:Date.now(),...wifiReply}]);
      return;
    }
    const localReply=handleMenuSelection(clean, activeCategoryId);
    const wantsAi=["ai","ask ai","talk to ai"].includes(clean.toLowerCase().trim());
    if(localReply && localReply.type!=="ai" && !wantsAi) {
      if(isMainMenuResponse(localReply)) setActiveCategoryId(null);
      if(isSubMenuResponse(localReply)) setActiveCategoryId(findCategoryFromSubMenu(localReply));
      if(localReply.type==="steps") {
        setActiveCategoryId(null);
        setLastHelpdeskContext(getTicketContextFromTroubleshooting(localReply));
      }
      setMessages(prev=>[...prev,{id:genToken(),role:"assistant",at:Date.now(),...localReply}]);
      return;
    }
    const aiMessage=localReply?.aiPrompt || clean;
    setLoading(true);
    try {
      console.log("Calling AI chat API", { message: aiMessage });
      const response=await fetch("/api/chat",{
        method:"POST",
        headers:{"Content-Type":"application/json"},
        cache:"no-store",
        body:JSON.stringify({message:aiMessage,user:{name:session?.name || session?.email || "Portal User",email:session?.email || ""}})
      });
      const data=await response.json().catch(()=>({reply:"",error:"INVALID_JSON_RESPONSE"}));
      console.log("AI chat API response", { status:response.status, ok:response.ok, data });
      if(!response.ok) throw new Error(data?.detail || data?.error || `AI endpoint failed: ${response.status}`);
      const reply=(data?.reply || "").trim() || "I have forwarded this issue to IT Support Team.";
      setMessages(prev=>[...prev,{id:genToken(),role:"assistant",text:reply,at:Date.now()}]);
    } catch (error) {
      console.error("AI chatbot error:",error);
      const errorText = import.meta.env.DEV ? `AI API error: ${error.message}` : "Sorry, I am unable to respond right now. Please contact IT Support.";
      setMessages(prev=>[...prev,{id:genToken(),role:"assistant",text:errorText,at:Date.now(),error:true}]);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="ai-helpdesk-wrap">
      {open&&(
        <div className="ai-helpdesk-panel glass">
          <div className="ai-helpdesk-head">
            <div className="ai-helpdesk-avatar">AI</div>
            <div style={{minWidth:0,flex:1}}>
              <div style={{fontSize:15,fontWeight:900,color:'#fff',whiteSpace:'nowrap',overflow:'hidden',textOverflow:'ellipsis'}}>Jaipuria Helpdesk AI</div>
              <div style={{fontSize:11,color:'rgba(226,232,240,.58)',display:'flex',alignItems:'center',gap:6}}><span style={{width:7,height:7,borderRadius:'50%',background:'#10b981',display:'inline-block'}}/>Online support assistant</div>
            </div>
            <button type="button" onClick={()=>setOpen(false)} className="ai-helpdesk-close" aria-label="Close AI helpdesk chat">×</button>
          </div>

          <div className="ai-helpdesk-messages">
            {messages.map(m=>(
              <div key={m.id} className={`ai-helpdesk-row ${m.role==='user'?'user':'bot'}`}>
                <div className={`ai-helpdesk-bubble ${m.role==='user'?'user':'bot'}`}>
                  {m.type==="menu"&&(
                    <div className="ai-helpdesk-menu-card">
                      <div className="ai-helpdesk-card-title">{m.text}</div>
                      <div className="ai-helpdesk-menu-list">
                        {m.menu.map(item=>(
                          <div key={item.id} className="ai-helpdesk-menu-item">
                            <span>{item.number}</span>
                            <strong>{item.label}</strong>
                          </div>
                        ))}
                      </div>
                      {m.hint&&<div className="ai-helpdesk-menu-hint">{m.hint}</div>}
                    </div>
                  )}
                  {m.type==="steps"&&(
                    <div className="ai-helpdesk-steps-card">
                      <div className="ai-helpdesk-card-title">{m.title}</div>
                      <ol>
                        {m.steps.map((step,index)=><li key={`${m.id}-${index}`}>{step}</li>)}
                      </ol>
                      {m.commonCauses&&(
                        <div className="ai-helpdesk-causes">
                          <div>Common Causes:</div>
                          <ul>
                            {m.commonCauses.map((cause,index)=><li key={`${m.id}-cause-${index}`}>{cause}</li>)}
                          </ul>
                        </div>
                      )}
                      {m.url&&<a className="ai-helpdesk-url" href={m.url} target="_blank" rel="noreferrer">{m.url}</a>}
                      <div className="ai-helpdesk-card-footer">{m.footer}</div>
                    </div>
                  )}
                  {m.type!=="menu"&&m.type!=="steps"&&<div style={{whiteSpace:'pre-wrap',overflowWrap:'anywhere'}}>{m.text}</div>}
                  <div style={{fontSize:10,color:'rgba(226,232,240,.38)',marginTop:5,textAlign:m.role==='user'?'right':'left'}}>{fmtDate(m.at)}</div>
                </div>
              </div>
            ))}
            {loading&&(
              <div className="ai-helpdesk-row bot">
                <div className="ai-helpdesk-bubble bot ai-typing">AI is analyzing your issue... <span></span><span></span><span></span></div>
              </div>
            )}
            <div ref={endRef}/>
          </div>

          <div className="ai-helpdesk-input">
            <input value={input} onChange={e=>setInput(e.target.value)} onKeyDown={e=>{if(e.key==='Enter'&&!e.shiftKey){e.preventDefault();sendMessage();}}} placeholder={listening ? "Listening..." : "Type hi, menu, WiFi, AI, or ESCALATE..."} />
            <button type="button" className={`mic-btn ${listening ? "listening" : ""}`} onClick={startVoiceInput} title={listening ? "Listening..." : "Use voice input"} aria-label="Use voice input"><span>🎙️</span>{listening&&<span className="mic-btn-text">Listening...</span>}</button>
            <button className="glow-btn" type="button" onClick={()=>sendMessage()} disabled={loading||!input.trim()}>Send</button>
          </div>
        </div>
      )}
      <button type="button" className="ai-helpdesk-button" onClick={()=>setOpen(o=>!o)} aria-label="Open AI helpdesk chat">
        <span>✦</span> {open?"Minimize":"AI Help"}
      </button>
    </div>
  );
}function PortalFeedbackChrome({onOpen}) {
  return (
    <>
      <button
        type="button"
        onClick={onOpen}
        className="portal-feedback-tab" style={{position:"fixed",right:12,bottom:12,zIndex:900,border:"1px solid rgba(255,255,255,0.18)",borderRadius:999,padding:"9px 14px",fontSize:12,fontWeight:800,color:"#fff",background:"linear-gradient(135deg,#7c3aed,#2563eb,#06b6d4,#10b981)",boxShadow:"0 14px 34px rgba(37,99,235,0.36),0 0 22px rgba(6,182,212,0.24)",backdropFilter:"blur(16px)",whiteSpace:"nowrap"}}
        onMouseEnter={e=>{e.currentTarget.style.transform="translateY(-2px)";e.currentTarget.style.boxShadow="0 18px 44px rgba(37,99,235,0.46),0 0 30px rgba(6,182,212,0.34)";}}
        onMouseLeave={e=>{e.currentTarget.style.transform="translateY(0)";e.currentTarget.style.boxShadow="0 14px 34px rgba(37,99,235,0.36),0 0 22px rgba(6,182,212,0.24)";}}
      >
        Portal Feedback
      </button>
    </>
  );
}

function PWAInstallPrompt() {
  const [promptEvent,setPromptEvent]=useState(null);
  const [dismissed,setDismissed]=useState(()=>DB.get("pwa_install_dismissed", false));
  const [installed,setInstalled]=useState(false);

  useEffect(()=>{
    const isStandalone = window.matchMedia?.("(display-mode: standalone)")?.matches || window.navigator.standalone;
    setInstalled(Boolean(isStandalone));
    const onPrompt=(event)=>{
      event.preventDefault();
      setPromptEvent(event);
    };
    const onInstalled=()=>{
      setInstalled(true);
      setPromptEvent(null);
      DB.set("pwa_install_dismissed", true);
    };
    window.addEventListener("beforeinstallprompt",onPrompt);
    window.addEventListener("appinstalled",onInstalled);
    return ()=>{
      window.removeEventListener("beforeinstallprompt",onPrompt);
      window.removeEventListener("appinstalled",onInstalled);
    };
  },[]);

  const isIOS = /iphone|ipad|ipod/i.test(window.navigator.userAgent || "");
  const canShow = !installed && !dismissed && (promptEvent || isIOS);
  if(!canShow) return null;

  const install=async()=>{
    if(!promptEvent) return;
    promptEvent.prompt();
    await promptEvent.userChoice.catch(()=>null);
    setPromptEvent(null);
    setDismissed(true);
    DB.set("pwa_install_dismissed", true);
  };
  const close=()=>{
    setDismissed(true);
    DB.set("pwa_install_dismissed", true);
  };

  return (
    <div style={{position:"fixed",left:12,bottom:12,zIndex:910,maxWidth:360,border:"1px solid rgba(125,211,252,.22)",borderRadius:16,background:"linear-gradient(135deg,rgba(7,17,31,.96),rgba(15,23,42,.94))",boxShadow:"0 18px 44px rgba(0,0,0,.38),0 0 24px rgba(14,165,233,.18)",padding:"12px 13px",color:"#e2e8f0",backdropFilter:"blur(18px)"}}>
      <div style={{display:"flex",gap:10,alignItems:"flex-start"}}>
        <div className="pulse" style={{width:34,height:34,borderRadius:12,background:"rgba(14,165,233,.16)",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>✦</div>
        <div style={{minWidth:0,flex:1}}>
          <div style={{fontSize:13,fontWeight:900,color:"#fff"}}>Install Jaipuria Helpdesk App</div>
          <div style={{fontSize:12,lineHeight:1.45,color:"rgba(226,232,240,.62)",marginTop:3}}>
            {isIOS ? "Tap Share button and select Add to Home Screen." : "Open the portal faster from your home screen."}
          </div>
          <div style={{display:"flex",gap:8,marginTop:10,flexWrap:"wrap"}}>
            {promptEvent&&<button className="glow-btn" type="button" onClick={install} style={{padding:"8px 12px",fontSize:12}}>Install</button>}
            <button type="button" onClick={close} style={{background:"rgba(255,255,255,.06)",border:"1px solid rgba(255,255,255,.1)",color:"#cbd5e1",padding:"8px 12px",borderRadius:10,fontSize:12,fontWeight:800}}>Later</button>
          </div>
        </div>
      </div>
    </div>
  );
}

function NotificationButton({toast,enabled,setEnabled}) {
  const enable=async()=>{
    try {
      if(typeof window==="undefined" || !("Notification" in window)) {
        toast("Browser notifications are not supported on this device.","info");
        return;
      }
      const permission=await Notification.requestPermission();
      if(permission==="granted") {
        setEnabled(true);
        toast("Notifications enabled","success");
        showBrowserNotification("Jaipuria IT Helpdesk", "Notifications are enabled.");
      } else {
        setEnabled(false);
        toast("Notifications permission was not enabled.","info");
      }
    } catch(error) {
      console.error("Notification permission failed:", error);
      toast("Notifications could not be enabled on this browser.","error");
    }
  };
  if(enabled) return <span style={{fontSize:12,color:"#86efac",fontWeight:800,whiteSpace:"nowrap"}}>Notifications On</span>;
  return <button type="button" onClick={enable} style={{background:"rgba(14,165,233,.1)",border:"1px solid rgba(125,211,252,.24)",color:"#bae6fd",padding:"6px 10px",borderRadius:999,fontSize:12,fontWeight:900,whiteSpace:"nowrap"}}>Enable Notifications</button>;
}
// ── SIDEBAR ───────────────────────────────────────────────────────────────
function Sidebar({current,onChange,isAdmin,isStaff,tickets,feedback=[],mobileOpen,setMobileOpen,onStaffAction,feedbackPendingCount=0}) {
  const adminNav=[{id:"dashboard",icon:"🏠",label:"Dashboard"},{id:"tickets",icon:"🎫",label:"All Tickets"},{id:"staff",icon:"👥",label:"IT Staff"},{id:"analytics",icon:"📊",label:"Analytics"},{id:"feedback",icon:"★",label:"IT Feedback"},{id:"export",icon:"⬇",label:"Export Reports"},{id:"staff-management",icon:"👥",label:"Staff Management"},{id:"emaillog",icon:"📧",label:"Email Log"},{id:"portal-feedback",icon:"★",label:"Portal Feedback"},{id:"temp-issue",icon:"📦",label:"Temp Issue"}];
  const userNav=[{id:"home",icon:"🏠",label:"Home"},{id:"my-tickets",icon:"🎫",label:"My Tickets"},{id:"know-staff",icon:"👥",label:"Connect with IT Staff"},{id:"feedback",icon:"★",label:"IT Feedback"},{id:"new-ticket",icon:"➕",label:"New Ticket"},{id:"track",icon:"🔍",label:"Track Ticket"},{id:"temp-issue",icon:"🧾",label:"Temp Issue"}];
  const staffNav=[{id:"staff-dash",icon:"🏠",label:"My Dashboard"},{id:"assigned",icon:"📋",label:"Assigned Tickets"},{id:"chat",icon:"💬",label:"Staff Chat",staffAction:true},{id:"know-staff",icon:"👥",label:"Connect with IT Staff"},{id:"temp-issue",icon:"🧾",label:"Temp Issue"},{id:"profile",icon:"👤",label:"My Profile",staffAction:true},{id:"password",icon:"🔐",label:"Change Password",staffAction:true},{id:"logout",icon:"↩",label:"Logout",staffAction:true}];
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
            <button key={item.id} onClick={()=>{if(isStaff&&item.staffAction&&onStaffAction){onStaffAction(item.id);}else{onChange(item.id);}setMobileOpen(false);}} style={{
              display:"flex",alignItems:"center",gap:12,padding:"10px 14px",borderRadius:10,border:"none",textAlign:"left",width:"100%",fontSize:14,
              background:current===item.id?"rgba(99,102,241,0.2)":"transparent",
              color:current===item.id?"#818cf8":"rgba(226,232,240,0.6)",
              fontWeight:current===item.id?600:400,
              borderLeft:current===item.id?"3px solid #6366f1":"3px solid transparent",
            }}>
              <span style={{fontSize:18}}>{item.icon}</span><span style={{flex:1}}>{item.label}</span>
              {item.id==="my-tickets"&&feedbackPendingCount>0&&<span title="Closed tickets pending feedback" className="pulse" style={{width:9,height:9,borderRadius:"50%",background:"#22c55e",boxShadow:"0 0 0 4px rgba(34,197,94,.12)",display:"inline-block"}}/>}
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
    try {
      await Promise.resolve(onComplete(hash));
      toast("Password set successfully! 🎉","success");
    } catch (error) {
      console.error("Staff password setup failed:", error);
      toast(`Password setup failed: ${error?.message || "Please try again"}`, "error");
    } finally {
      setLoading(false);
    }  };

  return (
    <div style={{minHeight:"100dvh",background:"radial-gradient(ellipse at 30% 40%,rgba(99,102,241,0.18) 0%,transparent 60%),#0a0a0f",display:"flex",alignItems:"center",justifyContent:"center",padding:"max(20px, clamp(12px, 4vw, 32px))",overflowY:"auto",overflowX:"hidden"}}>
      <div style={{width:"100%",maxWidth:420}} className="fade-up">
        <div style={{textAlign:"center",marginBottom:"clamp(20px, 6vw, 32px)"}}>
          <div style={{width:64,height:64,borderRadius:18,background:`${staff.color}33`,border:`2px solid ${staff.color}`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:24,fontWeight:800,color:staff.color,margin:"0 auto 16px"}}>{staff.avatar}</div>
          <h1 style={{fontFamily:"Syne",fontSize:"clamp(20px, 5.5vw, 28px)",fontWeight:800,color:"#e2e8f0",marginBottom:6,lineHeight:1.2}}>Welcome, {staff.name.split(" ")[0]}!</h1>
          <p style={{fontSize:"clamp(12px, 3.5vw, 14px)",color:"rgba(226,232,240,0.5)",lineHeight:1.4}}>First login detected. Please create your secure password.</p>
        </div>
        <div className="glass" style={{padding:"clamp(16px, 5vw, 28px)",display:"flex",flexDirection:"column",gap:"clamp(12px, 3vw, 18px)",flexShrink:0}}>
          <div style={{background:"rgba(245,158,11,0.1)",border:"1px solid rgba(245,158,11,0.3)",borderRadius:10,padding:"12px 14px",fontSize:"clamp(11px, 2.5vw, 13px)",color:"#fbbf24",lineHeight:1.4}}>
            🔐 Your password is encrypted and stored securely. It cannot be recovered — keep it safe.
          </div>
          <div>
            <label style={{fontSize:"clamp(10px, 2.5vw, 12px)",color:"rgba(226,232,240,0.5)",marginBottom:6,display:"block"}}>Create Password</label>
            <PwdInput value={pwd} onChange={setPwd} placeholder="Minimum 8 characters" showStrength />
          </div>
          <div>
            <label style={{fontSize:"clamp(10px, 2.5vw, 12px)",color:"rgba(226,232,240,0.5)",marginBottom:6,display:"block"}}>Confirm Password</label>
            <PwdInput value={confirm} onChange={setConfirm} placeholder="Repeat your password"/>
            {confirm&&pwd!==confirm&&<div style={{fontSize:"clamp(10px, 2.5vw, 12px)",color:"#f87171",marginTop:4}}>Passwords do not match</div>}
          </div>
          <button className="glow-btn" style={{width:"100%",padding:"clamp(10px, 2.5vw, 12px) 28px"}} onClick={submit} disabled={loading||s<2||pwd!==confirm}>
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
      const existingProfile = await fetchStaffProfile(staff.id).catch(error => {
        console.error("Staff reset profile lookup failed:", error);
        return null;
      });
      await saveStaffProfile(staff.id, {
        ...(existingProfile || {}),
        id: String(staff.id),
        staffId: staff.id,
        email: staff.email,
        name: staff.name,
        role: staff.role,
        permissions: staff.permissions,
        passwordHash: hash,
        passwordSet: true,
        passwordUpdatedAt: Date.now(),
      });
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
      const staffPasswords = DB.get("staff_passwords", {});
      const profile = await fetchStaffProfile(staff.id).catch(error => {
        console.error("Staff profile lookup failed:", error);
        return null;
      });

      const firestoreHash = profile?.passwordHash || profile?.password || "";
      const storedHash = firestoreHash || staffPasswords[staff.id] || "";
      const hasPassword = Boolean(profile?.passwordSet === true || firestoreHash || storedHash);

      if (!hasPassword) {
        setLoading(false);
        onLogin({type:"staff_firstlogin",staffId:staff.id,staff,requiresPasswordSetup:true});
        return;
      }

      if(!pwd){toast("Enter your password","error");setLoading(false);return;}
      const valid = storedHash ? await verifyPassword(pwd,storedHash) : profile?.passwordSet === true;
      setLoading(false);
      if(valid){
        clearStaffPasswordSetupStorage();
        if(firestoreHash && staffPasswords[staff.id]!==firestoreHash){
          staffPasswords[staff.id]=firestoreHash;
          DB.set("staff_passwords",staffPasswords);
        }
        onLogin({type:"staff",staffId:staff.id,email:staff.email,name:staff.name,role:staff.role,permissions:staff.permissions,passwordSet:true,requiresPasswordSetup:false});
      }
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
        <form onSubmit={e=>{e.preventDefault();handleLogin();}} style={{padding:"clamp(26px,5vw,54px)",display:"flex",flexDirection:"column",justifyContent:"center",minHeight:560,overflow:"hidden"}}>
          <div style={{display:"flex",alignItems:"center",gap:14,marginBottom:30}}>
            <div style={{background:"rgba(255,255,255,0.96)",border:"1px solid rgba(255,255,255,0.35)",borderRadius:16,padding:"8px 12px",boxShadow:"0 14px 34px rgba(0,0,0,0.24)",display:"flex",alignItems:"center",justifyContent:"center"}}>
              <img src="/jaipuria-logo.png" alt="Jaipuria Institute of Management" style={{width:"clamp(210px,25vw,260px)",height:"auto",display:"block",objectFit:"contain"}} />
            </div>
          </div>

          <div style={{marginBottom:24}}>
            <div style={{fontFamily:"Syne",fontWeight:800,fontSize:"clamp(22px,3vw,34px)",lineHeight:1.15,color:"#f8fafc",letterSpacing:0,whiteSpace:"normal",textWrap:"balance",maxWidth:"100%",overflowWrap:"break-word",wordBreak:"normal"}}>Login With Jaipuria Email ID</div>
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
function TicketsTable({tickets,onView,isAdmin,onDelete,emptyKind=""}) {
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
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(240px,1fr))",gap:10}}>
        {filtered.map(t=>(
          <div key={t.id} style={{position:"relative"}}>
            <TicketCard ticket={t} onView={onView} showFeedbackUnread={isAdmin&&isTicketFeedbackUnread(t,true,false)}/>
            {isAdmin&&<button onClick={e=>{e.stopPropagation();onDelete(t.id);}} style={{position:"absolute",top:10,right:10,background:"rgba(239,68,68,0.15)",border:"1px solid rgba(239,68,68,0.3)",color:"#f87171",width:26,height:26,borderRadius:6,fontSize:13,display:"flex",alignItems:"center",justifyContent:"center"}}>🗑</button>}
          </div>
        ))}
        {filtered.length===0&&(
          <EmptyState
            icon={filterStatus==="Open" ? "🎉" : "🎫"}
            message={filterStatus==="Open" || emptyKind==="Open" ? "No open tickets. Everything looks good." : filterStatus==="Closed" || emptyKind==="Closed" ? "No closed tickets yet." : "No tickets found."}
          />
        )}
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
        <div style={{marginBottom:10,border:"1px solid rgba(125,211,252,.22)",background:"rgba(14,165,233,.1)",borderRadius:12,padding:"10px 12px",fontSize:13,color:"#bfdbfe"}}>AI says: {getTicketStatusExplanation(result.status)}</div>
        <div style={{fontSize:14,color:"rgba(226,232,240,0.7)",marginBottom:8}}>{result.description.slice(0,100)}...</div>
        <PriorityBadge p={result.priority}/><div style={{marginTop:10}}><TimerBadge ticket={result}/></div>
        <div style={{marginTop:10,fontSize:12,color:"#818cf8"}}>Click to view full details →</div>
      </div>}
    </div>
  );
}

const ACTION_TAB_LABELS = {
  view_all:"View All",
  assign:"Assign",
  close:"Closed",
  export:"Export",
  manage_users:"Manage Users"
};

function ActionTabs({permissions=[],active,onSelect}) {
  const available = ["view_all","assign","close","export","manage_users"].filter(action => permissions.includes(action));
  if(!available.length) return null;
  return (
    <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
      {available.map(action=>(
        <button
          key={action}
          type="button"
          onClick={()=>onSelect(action)}
          style={{
            border:"1px solid rgba(125,211,252,.22)",
            borderRadius:999,
            padding:"7px 11px",
            fontSize:12,
            fontWeight:900,
            color:active===action ? "#fff" : "#bae6fd",
            background:active===action ? "linear-gradient(135deg,#2563eb,#8b5cf6,#06b6d4)" : "rgba(14,165,233,.1)",
            boxShadow:active===action ? "0 12px 30px rgba(37,99,235,.32),0 0 20px rgba(6,182,212,.18)" : "none"
          }}
        >
          {ACTION_TAB_LABELS[action] || action}
        </button>
      ))}
    </div>
  );
}

function ManageUsersPanel() {
  return (
    <div className="glass" style={{padding:"18px",display:"grid",gap:14}}>
      <div>
        <h3 style={{fontFamily:"Syne",fontSize:17,fontWeight:900,color:"#fff"}}>Manage Users</h3>
        <p style={{fontSize:12,color:"rgba(226,232,240,.55)",marginTop:4}}>User management module is ready for integration.</p>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(220px,1fr))",gap:10}}>
        {[{name:"Admin",email:"admin@jaipuria.ac.in",role:"Admin"},...STAFF_BASE.map(staff=>({name:staff.name,email:staff.email,role:staff.role}))].map(user=>(
          <div key={user.email} className="glass2" style={{padding:"12px"}}>
            <div style={{fontSize:13,fontWeight:900,color:"#f8fafc"}}>{user.name}</div>
            <div style={{fontSize:12,color:"rgba(226,232,240,.58)",marginTop:3}}>{user.email}</div>
            <span className="tag" style={{marginTop:8,fontSize:11}}>{user.role}</span>
          </div>
        ))}
      </div>
    </div>
  );
}

function filterActionTickets(action,tickets,baseTickets) {
  if(action==="view_all") return tickets;
  if(action==="assign") return tickets.filter(ticket => ["Open","Assigned"].includes(ticket.status) || !ticket.assignedTo || ticket.assignedTo==="IT Support Team");
  if(action==="close") return tickets.filter(ticket => ["Closed","Resolved"].includes(ticket.status));
  return baseTickets;
}

// ── STAFF PANEL ───────────────────────────────────────────────────────────
function StaffPanel({staffId,tickets,setTickets,toast,onViewTicket,onQuickAssign,permissions,staffProfiles={},staffStatuses={},showWelcome=false}) {
  const [activeAction,setActiveAction]=useState("assigned");
  const staff=STAFF_BASE.find(s=>s.id===staffId);
  const visibleTickets=(permissions||[]).includes("view_all")
    ? tickets
    : tickets.filter(t=>t.assigneeId===staffId || (t.watchers||[]).some(w=>Number(w.id)===Number(staffId)) || (t.notifiedStaff||[]).some(w=>Number(w.id)===Number(staffId)));
  const myTickets=filterActionTickets(activeAction, tickets, visibleTickets);
  const active=myTickets.filter(t=>!["Resolved","Closed"].includes(t.status)).length;
  const resolved=myTickets.filter(t=>t.status==="Resolved"||t.status==="Closed").length;
  const selectAction=action=>{
    if(action==="export") {
      downloadTicketCsv(myTickets.length ? myTickets : visibleTickets);
      toast("Tickets exported to CSV","success");
      setActiveAction("export");
      return;
    }
    setActiveAction(action);
  };
  return (
    <div style={{display:"flex",flexDirection:"column",gap:20}}>
      <SmartWelcome session={{name:staff?.name}} visible={showWelcome} />
      <div className="glass" style={{padding:"24px",display:"flex",gap:16,alignItems:"center",flexWrap:"wrap"}}>
        <StaffAvatar staff={staff} profiles={staffProfiles} statuses={staffStatuses} size={56} showStatus />
        <div style={{flex:1}}>
          <div style={{fontFamily:"Syne",fontSize:20,fontWeight:700,color:"#e2e8f0"}}>{staff.name}</div>
          <div style={{fontSize:13,color:"rgba(226,232,240,0.5)"}}>{staff.role} · {staff.email}</div>
          <div style={{marginTop:8}}><ActionTabs permissions={permissions} active={activeAction} onSelect={selectAction} /></div>
        </div>
        <div style={{display:"flex",gap:14,textAlign:"center",flexWrap:"wrap"}}>
          {[["Total",myTickets.length,"#818cf8"],["Active",active,"#fbbf24"],["Resolved",resolved,"#34d399"]].map(([l,v,c])=>(
            <div key={l}><div style={{fontSize:24,fontFamily:"Syne",fontWeight:800,color:c}}>{v}</div><div style={{fontSize:11,color:"rgba(226,232,240,0.4)"}}>{l}</div></div>
          ))}
        </div>
      </div>
      {activeAction==="manage_users" ? <ManageUsersPanel /> : (
        <>
          <h3 style={{fontFamily:"Syne",fontSize:16,fontWeight:700,color:"#e2e8f0"}}>{ACTION_TAB_LABELS[activeAction] || "Assigned Tickets"}</h3>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(240px,1fr))",gap:10}}>
            {[...myTickets].sort((a,b)=>b.createdAt-a.createdAt).map(t=>(
              <div key={t.id} style={{position:"relative"}}>
                <TicketCard ticket={t} onView={onViewTicket} showFeedbackUnread={isTicketFeedbackUnread(t,false,true)}/>
                {activeAction==="assign"&&(permissions||[]).includes("assign")&&<button onClick={e=>{e.stopPropagation();onQuickAssign?.(t.id);}} style={{position:"absolute",right:10,bottom:10,background:"rgba(99,102,241,0.92)",border:"1px solid rgba(255,255,255,0.18)",color:"#fff",padding:"6px 10px",borderRadius:8,fontSize:11,fontWeight:900}}>Assign</button>}
              </div>
            ))}
            {myTickets.length===0&&<EmptyState icon={activeAction==="close"?"🎫":"✅"} message={activeAction==="close"?"No closed tickets yet.":activeAction==="assign"?"No tickets need assignment right now.":"No tickets found."} />}
          </div>
        </>
      )}
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
    const profile=await fetchStaffProfile(staff.id).catch(error=>{console.error('Staff password profile lookup failed:',error);return null;});
    const current=profile?.passwordHash || passwords[staff.id];
    if(!current){toast('No password set yet. Please complete first login.','error');return;}
    if(!(await verifyPassword(oldPwd,current))){toast('Old password is incorrect','error');return;}
    if(pwdStrength(newPwd)<3){toast('New password is too weak','error');return;}
    if(newPwd!==confirm){toast('Passwords do not match','error');return;}
    const nextHash=await hashPassword(newPwd);
    await saveStaffProfile(staff.id, {...(profile||{}),id:String(staff.id),staffId:staff.id,email:staff.email,name:staff.name,role:staff.role,passwordHash:nextHash,passwordSet:true,passwordUpdatedAt:Date.now()});
    passwords[staff.id]=nextHash;
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
      name: "Raj Prakash Singh",
      post: "Manager",
      email: "raj.singh@jaipuria.ac.in",
      contact: "9887283825",
      whatsapp: "9887283825"
    },
    {
      id: 2,
      name: "Rohit Jangid",
      post: "Executive",
      email: "rohit.jangid@jaipuria.ac.in",
      contact: "8005978632",
      whatsapp: "8005978632"
    },
    {
      id: 3,
      name: "Vishal Swami",
      post: "Senior Executive",
      email: "vishal.swami@jaipuria.ac.in",
      contact: "8233771101",
      whatsapp: "8233771101"
    }
  ];

  return (
    <div>
      <h2 style={{fontFamily:"Syne",fontSize:22,fontWeight:700,color:"#e2e8f0",marginBottom:20}}>Connect with IT Staff</h2>
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

function TempIssuePanel({session, tempIssues, tempIssuesLoaded, filters, setFilters, onSubmit, onAction, toast}) {
  const isNormalUser = session?.type === "user";
  const currentStaffName = session?.type === "staff" ? STAFF_BASE.find(s => s.id === session.staffId)?.name || "" : "";
  const canAdmin = session?.type === "admin";
  const today = new Date().toISOString().slice(0, 10);

  const emptyForm = {
    userName: session?.name || "",
    userEmail: session?.email || "",
    mobile: "",
    item: "",
    customItem: "",
    permissionApprovedBy: "",
    requestedToStaff: "",
    issueDate: today,
    purpose: "",
  };

  const [form, setForm] = useState(emptyForm);
  const [loading, setLoading] = useState(false);
  const [returnDrafts, setReturnDrafts] = useState({});

  useEffect(() => {
    setForm(f => ({
      ...f,
      userName: session?.name || f.userName,
      userEmail: session?.email || f.userEmail,
    }));
  }, [session?.name, session?.email]);

  const itemOptions = ["Laptop","Projector","HDMI Cable","VGA Cable","LAN Cable","Mouse","Keyboard","Charger","Speaker","Microphone","Webcam","Extension Board","Pen Drive","External Hard Disk","Tablet","Other"];
  const permissionOptions = ["Director's Office","Faculty","Admin. Office","HR","Accounts","PMC","Student Affairs","MRC Office","Examination","FPM","IT","Library","Admissions & Marketing","Training","Placements & Corporate Relations","MDP","Training & Consultancy","IRC & E-Cell","Support Staff", ...STAFF_BASE.map(s => s.name)];
  const staffOptions = STAFF_BASE.map(s => s.name);
  const returnStaffOptions = [
    { name: "Vishal Swami", email: STAFF_BASE.find(s => s.name === "Vishal Swami")?.email || "vishal.swami@jaipuria.ac.in" },
    { name: "Raj Prakash Singh", email: STAFF_BASE.find(s => s.id === 1)?.email || "raj.singh@jaipuria.ac.in" },
    { name: "Rohit Jangid", email: STAFF_BASE.find(s => s.name === "Rohit Jangid")?.email || "rohit.jangid@jaipuria.ac.in" },
  ];
  const statusOptions = ["All","Pending Approval","Approved","Issued","Return Requested","Returned","Rejected","Not Issued","Return Rejected","Force Closed"];

  const getStaffForIssue = issue => STAFF_BASE.find(s => s.name === (issue.requestedToStaff || issue.requestToStaff));
  const issueStaffName = issue => issue.requestedToStaff || issue.requestToStaff || "";
  const issuePermission = issue => issue.permissionApprovedBy || issue.permissionBy || "";
  const issueItem = issue => issue.item === "Other" ? issue.customItem || issue.item : issue.item || issue.customItem || "";
  const canManageIssue = issue => canAdmin || (session?.type === "staff" && (issueStaffName(issue) === currentStaffName || issue.returnToStaff === currentStaffName));
  const isOverdue = issue => issue.status === "Issued" && issue.issueDate && issue.issueDate < today;

  const visibleIssues = canAdmin
    ? tempIssues
    : session?.type === "staff"
      ? tempIssues.filter(issue => issueStaffName(issue) === currentStaffName || issue.returnToStaff === currentStaffName)
      : tempIssues.filter(issue => (issue.userEmail || "").toLowerCase() === (session?.email || "").toLowerCase());

  const filteredIssues = visibleIssues.filter(issue => {
    const statusFilter = filters.status || "All";
    const staffFilter = filters.staff || "All";
    const itemFilter = filters.item || "All";
    if (statusFilter !== "All" && issue.status !== statusFilter) return false;
    if (staffFilter !== "All" && issueStaffName(issue) !== staffFilter) return false;
    if (itemFilter !== "All" && issueItem(issue) !== itemFilter) return false;
    if (filters.from && new Date(issue.issueDate || issue.createdAt) < new Date(filters.from)) return false;
    if (filters.to && new Date(issue.issueDate || issue.createdAt) > new Date(filters.to)) return false;
    const q = (filters.search || "").trim().toLowerCase();
    if (q && ![issue.requestId, issue.userName, issue.userEmail, issue.mobile, issueItem(issue), issuePermission(issue), issueStaffName(issue), issue.status].some(value => (value || "").toLowerCase().includes(q))) return false;
    return true;
  });

  const adminCards = [
    ["Total Requests", tempIssues.length, "All", "#818cf8"],
    ["Pending Approval", tempIssues.filter(i => i.status === "Pending Approval").length, "Pending Approval", "#fbbf24"],
    ["Approved", tempIssues.filter(i => i.status === "Approved").length, "Approved", "#38bdf8"],
    ["Issued", tempIssues.filter(i => i.status === "Issued").length, "Issued", "#22c55e"],
    ["Return Requested", tempIssues.filter(i => i.status === "Return Requested").length, "Return Requested", "#a78bfa"],
    ["Returned", tempIssues.filter(i => i.status === "Returned").length, "Returned", "#10b981"],
    ["Rejected", tempIssues.filter(i => ["Rejected","Not Issued","Return Rejected"].includes(i.status)).length, "Rejected", "#f87171"],
    ["Overdue Items", tempIssues.filter(isOverdue).length, "Issued", "#fb7185"],
  ];

  const resetForm = () => setForm({ ...emptyForm, userEmail: session?.email || "", userName: session?.name || "" });

  const handleSubmit = async () => {
    if (!isNormalUser) {
      toast("Only users can create Temp Issue requests", "error");
      return;
    }
    if (!form.userName.trim() || !form.userEmail.trim() || !form.mobile.trim() || !form.item || !form.permissionApprovedBy || !form.requestedToStaff) {
      toast("Please fill all required request fields.", "error");
      return;
    }
    if (form.item === "Other" && !form.customItem.trim()) {
      toast("Please enter the item name for 'Other'.", "error");
      return;
    }
    setLoading(true);
    try {
      await onSubmit({
        userId: session?.email || form.userEmail.trim(),
        userName: form.userName.trim(),
        userEmail: form.userEmail.trim(),
        mobile: form.mobile.trim(),
        item: form.item === "Other" ? form.customItem.trim() : form.item,
        customItem: form.item === "Other" ? form.customItem.trim() : "",
        permissionApprovedBy: form.permissionApprovedBy,
        permissionBy: form.permissionApprovedBy,
        requestedToStaff: form.requestedToStaff,
        requestToStaff: form.requestedToStaff,
        issueDate: form.issueDate || today,
        returnDate: "",
        purpose: form.purpose.trim(),
      });
      resetForm();
    } catch (error) {
      console.error(error);
    } finally {
      setLoading(false);
    }
  };

  const statusChip = status => {
    const map = {
      "Pending Approval": ["rgba(245,158,11,0.16)", "#fbbf24"],
      Approved: ["rgba(14,165,233,0.16)", "#38bdf8"],
      Issued: ["rgba(34,197,94,0.16)", "#86efac"],
      "Return Requested": ["rgba(167,139,250,0.16)", "#c4b5fd"],
      Returned: ["rgba(16,185,129,0.16)", "#6ee7b7"],
      Rejected: ["rgba(239,68,68,0.16)", "#fca5a5"],
      "Not Issued": ["rgba(249,115,22,0.16)", "#fdba74"],
      "Return Rejected": ["rgba(244,63,94,0.16)", "#fda4af"],
      "Force Closed": ["rgba(100,116,139,0.2)", "#cbd5e1"],
    };
    const [bg, color] = map[status] || ["rgba(99,102,241,0.12)", "#dbeafe"];
    return <span className="tag" style={{background:bg,color,border:`1px solid ${color}40`}}>{status}</span>;
  };

  const act = (issue, action, label, style = {}) => (
    <button
      onClick={() => {
        const remark = prompt(`${label} remarks (optional):`, "");
        onAction(issue.requestId, action, session?.name || currentStaffName || "Admin", remark || "");
      }}
      style={{padding:"7px 10px",fontSize:12,borderRadius:10,border:"1px solid rgba(255,255,255,0.14)",background:"rgba(255,255,255,0.07)",color:"#e2e8f0",...style}}
    >
      {label}
    </button>
  );

  const requestReturnButton = issue => {
    const selectedName = returnDrafts[issue.requestId] || "";
    const selectedStaff = returnStaffOptions.find(s => s.name === selectedName);
    return (
      <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
        <select
          value={selectedName}
          onChange={e=>setReturnDrafts(d=>({...d,[issue.requestId]:e.target.value}))}
          style={{minWidth:170,padding:"7px 10px",fontSize:12}}
        >
          <option value="">Select Return Staff</option>
          {returnStaffOptions.map(s => <option key={s.name} value={s.name}>{s.name}</option>)}
        </select>
        <button
          onClick={() => {
            if (!selectedStaff) {
              toast("Please select return IT staff", "error");
              return;
            }
            onAction(
              issue.requestId,
              "requestReturn",
              session?.name || issue.userName || session?.email || "User",
              "Return requested by user",
              { returnToStaff: selectedStaff.name, returnToStaffEmail: selectedStaff.email }
            );
          }}
          className="glow-btn"
          style={{padding:"7px 10px",fontSize:12}}
        >
          Request Return
        </button>
      </div>
    );
  };
  const exportRows = filteredIssues.map(issue => ({
    "Request ID": issue.requestId,
    "User Name": issue.userName,
    Email: issue.userEmail,
    Mobile: issue.mobile,
    Item: issueItem(issue),
    "Permission Approved By": issuePermission(issue),
    "Requested To Staff": issueStaffName(issue),
    "Issue Date": issue.issueDate || "",
    Status: issue.status,
    Remarks: issue.remarks || "",
    "Created At": fmtDate(issue.createdAt),
    "Approved By": issue.approvedBy || "",
    "Approved At": issue.approvedAt ? fmtDate(issue.approvedAt) : "",
    "Issued By": issue.issuedBy || "",
    "Issued At": issue.issuedAt ? fmtDate(issue.issuedAt) : "",
    "Return To Staff": issue.returnToStaff || "",
    "Return Requested At": issue.returnRequestedAt ? fmtDate(issue.returnRequestedAt) : "",
    "Returned By": issue.returnedBy || "",
    "Returned At": issue.returnedAt ? fmtDate(issue.returnedAt) : "",
    "Return Remarks": issue.returnRemarks || "",
  }));

  const downloadExcelReport = () => downloadExcel(exportRows, `temp_issues_${new Date().toISOString().slice(0,10)}.xlsx`);
  const downloadPdfReport = () => {
    const doc = new jsPDF({orientation:"landscape"});
    doc.setFillColor(15,23,42); doc.rect(0,0,297,28,"F");
    doc.setTextColor(255); doc.setFontSize(15); doc.text("Jaipuria Institute of Management",14,12);
    doc.setFontSize(12); doc.text("Temp Issue Report",14,21);
    doc.setFontSize(8); doc.text(`Generated: ${fmtDate(Date.now())}`,235,12);
    autoTable(doc,{startY:34,head:[Object.keys(exportRows[0]||{"Request ID":"","User Name":"","Email":"","Item":"","Status":""})],body:exportRows.map(r=>Object.values(r)),styles:{fontSize:7,cellPadding:2,overflow:"linebreak"},headStyles:{fillColor:[79,70,229],textColor:255},alternateRowStyles:{fillColor:[248,250,252]},margin:{left:8,right:8}});
    doc.save(`temp_issues_${new Date().toISOString().slice(0,10)}.pdf`);
  };

  const requestTitle = canAdmin ? "All Temp Issue Requests" : session?.type === "staff" ? "Assigned Temp Issue Requests" : "My Temp Issue Requests";

  return (
    <div style={{display:"flex",flexDirection:"column",gap:22}}>
      <div style={{display:"flex",justifyContent:"space-between",flexWrap:"wrap",gap:14,alignItems:"flex-end"}}>
        <div>
          <h2 style={{fontFamily:"Syne",fontSize:22,fontWeight:700,color:"#e2e8f0",marginBottom:4}}>Temp Issue Request</h2>
          <div style={{fontSize:14,color:"rgba(226,232,240,0.5)"}}>{isNormalUser ? "Request temporary IT items and track approval, issue, and return status." : "Manage approval, issue, return, and reporting for temporary IT items."}</div>
        </div>
        {(canAdmin || session?.type === "staff") && <div style={{display:"flex",gap:10,flexWrap:"wrap"}}><button className="glow-btn" onClick={downloadExcelReport}>Export Excel</button><button className="glow-btn" onClick={downloadPdfReport}>Export PDF</button></div>}
      </div>

      {canAdmin && (
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(155px,1fr))",gap:14}}>
          {adminCards.map(([label,value,status,color]) => <StatCard key={label} label={label} value={value} icon="📦" color={color} onClick={() => setFilters({...filters,status})} />)}
        </div>
      )}

      {isNormalUser && (
        <div className="glass" style={{padding:24,display:"grid",gridTemplateColumns:"1fr 1fr",gap:18}}>
          <div style={{display:"grid",gap:14}}>
            <label style={{fontSize:13,color:"rgba(226,232,240,0.6)"}}>Name *</label>
            <input value={form.userName} onChange={e=>setForm({...form,userName:e.target.value})} placeholder="Enter your name" />
            <label style={{fontSize:13,color:"rgba(226,232,240,0.6)"}}>Email ID *</label>
            <input value={form.userEmail} onChange={e=>setForm({...form,userEmail:e.target.value})} placeholder="user@jaipuria.ac.in" />
            <label style={{fontSize:13,color:"rgba(226,232,240,0.6)"}}>Mobile Number *</label>
            <input value={form.mobile} onChange={e=>setForm({...form,mobile:e.target.value})} placeholder="Mobile number" />
            <label style={{fontSize:13,color:"rgba(226,232,240,0.6)"}}>Item *</label>
            <select value={form.item} onChange={e=>setForm({...form,item:e.target.value,customItem:e.target.value==="Other"?form.customItem:""})}>
              <option value="">Select Option</option>
              {itemOptions.map(item => <option key={item} value={item}>{item}</option>)}
            </select>
            {form.item === "Other" && <input value={form.customItem} onChange={e=>setForm({...form,customItem:e.target.value})} placeholder="Specify item" />}
          </div>
          <div style={{display:"grid",gap:14}}>
            <label style={{fontSize:13,color:"rgba(226,232,240,0.6)"}}>Permission Approved By *</label>
            <select value={form.permissionApprovedBy} onChange={e=>setForm({...form,permissionApprovedBy:e.target.value})}>
              <option value="">Select Option</option>
              {permissionOptions.map(name => <option key={name} value={name}>{name}</option>)}
            </select>
            <label style={{fontSize:13,color:"rgba(226,232,240,0.6)"}}>Issue Date</label>
            <input type="date" value={form.issueDate} onChange={e=>setForm({...form,issueDate:e.target.value})} />
            <label style={{fontSize:13,color:"rgba(226,232,240,0.6)"}}>Request To IT Staff *</label>
            <select value={form.requestedToStaff} onChange={e=>setForm({...form,requestedToStaff:e.target.value})}>
              <option value="">Select Option</option>
              {staffOptions.map(name => <option key={name} value={name}>{name}</option>)}
            </select>
            <label style={{fontSize:13,color:"rgba(226,232,240,0.6)"}}>Reason / Purpose</label>
            <textarea rows={4} value={form.purpose} onChange={e=>setForm({...form,purpose:e.target.value})} placeholder="Optional context for IT staff" style={{resize:"vertical"}} />
            <button className="glow-btn" onClick={handleSubmit} disabled={loading}>{loading ? "Submitting..." : "Submit Request"}</button>
          </div>
        </div>
      )}

      <div className="glass" style={{padding:24}}>
        <div style={{display:"flex",flexWrap:"wrap",gap:12,marginBottom:18,alignItems:"center",justifyContent:"space-between"}}>
          <h3 style={{fontFamily:"Syne",fontSize:18,fontWeight:700,color:"#e2e8f0",margin:0}}>{requestTitle}</h3>
          <div style={{display:"flex",flexWrap:"wrap",gap:10,alignItems:"center"}}>
            <select value={filters.status || "All"} onChange={e=>setFilters({...filters,status:e.target.value})} style={{minWidth:160}}>{statusOptions.map(status => <option key={status} value={status}>{status}</option>)}</select>
            <select value={filters.staff || "All"} onChange={e=>setFilters({...filters,staff:e.target.value})} style={{minWidth:160}}><option value="All">All Staff</option>{staffOptions.map(name => <option key={name} value={name}>{name}</option>)}</select>
            <select value={filters.item || "All"} onChange={e=>setFilters({...filters,item:e.target.value})} style={{minWidth:150}}><option value="All">All Items</option>{itemOptions.filter(i=>i!=="Other").map(item => <option key={item} value={item}>{item}</option>)}</select>
            <input type="date" value={filters.from || ""} onChange={e=>setFilters({...filters,from:e.target.value})} />
            <input type="date" value={filters.to || ""} onChange={e=>setFilters({...filters,to:e.target.value})} />
            <input type="search" value={filters.search || ""} onChange={e=>setFilters({...filters,search:e.target.value})} placeholder="Search user, item, ticket" style={{minWidth:220}} />
            <button onClick={() => setFilters({ status:"All", staff:"All", item:"All", search:"", from:"", to:"" })} style={{background:"rgba(255,255,255,0.06)",border:"1px solid rgba(255,255,255,0.12)",color:"#e2e8f0",padding:"10px 16px",borderRadius:10}}>Clear</button>
          </div>
        </div>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse",minWidth:1280}}>
            <thead><tr style={{textAlign:"left",borderBottom:"1px solid rgba(255,255,255,0.1)"}}>{["Request ID","User","Email","Mobile","Item","Approved By","Issue Date","Requested To","Status","Timeline / Remarks","Actions"].map(label => <th key={label} style={{padding:"12px 10px",fontSize:12,color:"rgba(226,232,240,0.65)",fontWeight:700}}>{label}</th>)}</tr></thead>
            <tbody>
              {tempIssuesLoaded && filteredIssues.length === 0 && <tr><td colSpan={11} style={{padding:24,textAlign:"center",color:"rgba(226,232,240,0.4)"}}>No temp issue requests found.</td></tr>}
              {filteredIssues.map(issue => {
                const manageable = canManageIssue(issue);
                const ownedByUser = isNormalUser && (issue.userEmail || "").toLowerCase() === (session?.email || "").toLowerCase();
                const history = issue.requestHistory || [];
                return (
                  <tr key={issue.requestId} style={{borderBottom:"1px solid rgba(255,255,255,0.06)",verticalAlign:"middle"}}>
                    <td style={{padding:"12px 10px",fontWeight:700,color:"#c4b5fd",verticalAlign:"middle"}}>{issue.requestId}</td>
                    <td style={{padding:"12px 10px",verticalAlign:"middle"}}>{issue.userName}</td>
                    <td style={{padding:"12px 10px",verticalAlign:"middle"}}>{issue.userEmail}</td>
                    <td style={{padding:"12px 10px",verticalAlign:"middle"}}>{issue.mobile}</td>
                    <td style={{padding:"12px 10px",verticalAlign:"middle"}}>{issueItem(issue)}</td>
                    <td style={{padding:"12px 10px",verticalAlign:"middle"}}>{issuePermission(issue)}</td>
                    <td style={{padding:"12px 10px",verticalAlign:"middle"}}>{issue.issueDate}</td>
                    <td style={{padding:"12px 10px",verticalAlign:"middle"}}>{issueStaffName(issue)}<div style={{fontSize:11,color:"rgba(226,232,240,0.4)"}}>{getStaffForIssue(issue)?.email || issue.requestedToStaffEmail}</div></td>
                    <td style={{padding:"12px 10px",verticalAlign:"middle",textAlign:"center"}}>{statusChip(issue.status)}{isOverdue(issue)&&<div style={{marginTop:6,fontSize:11,color:"#fb7185",fontWeight:800}}>Overdue / not returned</div>}</td>
                    <td style={{padding:"12px 10px",minWidth:240,verticalAlign:"middle"}}>
                      <div style={{fontSize:12,color:"rgba(226,232,240,0.75)",lineHeight:1.5}}>{issue.remarks || "No remarks"}</div>
                      <div style={{fontSize:11,color:"rgba(226,232,240,0.38)",marginTop:5}}>Created: {fmtDate(issue.createdAt)}{issue.approvedAt ? ` · Approved: ${fmtDate(issue.approvedAt)}` : ""}{issue.issuedAt ? ` · Issued: ${fmtDate(issue.issuedAt)}` : ""}{issue.returnRequestedAt ? ` · Return requested: ${fmtDate(issue.returnRequestedAt)}` : ""}{issue.returnedAt ? ` · Returned: ${fmtDate(issue.returnedAt)}` : ""}</div>
                      {history.length>0&&<div style={{marginTop:6,fontSize:11,color:"rgba(226,232,240,0.42)"}}>Latest: {history[history.length-1]?.action} by {history[history.length-1]?.by}</div>}
                    </td>
                    <td style={{padding:"12px 10px",verticalAlign:"middle"}}><div style={{display:"flex",flexWrap:"wrap",gap:6,minWidth:230,alignItems:"center"}}>
                      {manageable && issue.status === "Pending Approval" && act(issue,"approve","Approve",{background:"rgba(14,165,233,0.16)",color:"#dbeafe"})}
                      {manageable && issue.status === "Pending Approval" && act(issue,"reject","Reject",{background:"rgba(239,68,68,0.16)",color:"#fee2e2"})}
                      {manageable && issue.status === "Approved" && act(issue,"issue","Mark Issued",{background:"rgba(16,185,129,0.16)",color:"#d1fae5"})}
                      {manageable && issue.status === "Approved" && act(issue,"notIssued","Mark Not Issued",{background:"rgba(249,115,22,0.16)",color:"#ffedd5"})}
                      {(ownedByUser || canAdmin) && issue.status === "Issued" && requestReturnButton(issue)}
                      {manageable && issue.status === "Return Requested" && act(issue,"acceptReturn","Receive Return",{background:"rgba(16,185,129,0.16)",color:"#d1fae5"})}
                      {manageable && issue.status === "Return Requested" && act(issue,"rejectReturn","Reject Return",{background:"rgba(244,63,94,0.16)",color:"#ffe4e6"})}
                      {canAdmin && !["Returned","Rejected","Force Closed"].includes(issue.status) && act(issue,"forceClose","Force Close",{background:"rgba(100,116,139,0.18)",color:"#e2e8f0"})}
                      {canAdmin && act(issue,"editStatus","Edit Status",{background:"rgba(99,102,241,0.16)",color:"#ddd6fe"})}
                      {issue.status === "Returned" && <span style={{fontSize:12,color:"#86efac",fontWeight:800}}>Returned to {issue.returnToStaff || issue.returnedBy || "IT Staff"}</span>}
                    </div></td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}
function StaffChatModal({staff,profiles,statuses}) {
  const firstPeerId = STAFF_BASE.find(s => s.id !== staff.id)?.id || STAFF_BASE[0].id;
  const [selected, setSelected] = useState(firstPeerId);
  const [mobileChatOpen, setMobileChatOpen] = useState(false);
  const [text, setText] = useState('');
  const [messages, setMessages] = useState([]);
  const [unreadCounts, setUnreadCounts] = useState({});
  const messagesEndRef = useRef(null);

  const peer = STAFF_BASE.find(s => s.id === selected);
  const thread = [staff.id, selected].sort((a, b) => a - b).join('-');

  useEffect(() => {
    setMessages([]);

    if (!firestoreDb || !thread || !staff?.id) {
      return undefined;
    }

    const q = query(
      collection(firestoreDb, 'messages'),
      where('thread', '==', thread)
    );

    const unsubscribe = onSnapshot(q, (snapshot) => {
      const msgs = snapshot.docs
        .map((snap) => normalizeMessage({ id: snap.id, ...snap.data() }))
        .sort((a, b) => (a.at || 0) - (b.at || 0));

      setMessages(msgs);
    });

    return () => unsubscribe();
  }, [thread, staff?.id]);

  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({behavior:'smooth',block:'end'});
  }, [messages.length, thread]);

  useEffect(() => {
    if (!firestoreDb || !staff?.id) return undefined;
    const unsubscribes = STAFF_BASE
      .filter(s => s.id !== staff.id)
      .map(peerStaff => {
        const threadKey = [staff.id,peerStaff.id].sort((a,b)=>a-b).join('-');
        const q = query(collection(firestoreDb, 'messages'), where('thread', '==', threadKey), where('to', '==', staff.id), where('read', '==', false));
        return onSnapshot(
          q,
          (snapshot) => setUnreadCounts(prev => ({ ...prev, [threadKey]: snapshot.size })),
          (error) => console.error('Unread chat listener failed:', error)
        );
      });
    return () => unsubscribes.forEach(unsub => unsub());
  }, [staff?.id]);

  const send=async()=>{
    const cleanText=text.trim();
    if(!cleanText || !firestoreDb || !peer) return;
    const createdAt=Date.now();
    const msg = {
      thread,
      from:staff.id,
      fromName:staff.name,
      to:selected,
      toName:peer.name,
      text:cleanText,
      at:createdAt,
      createdAt,
      read:false,
    };
    setText('');
   try {
await addDoc(collection(firestoreDb, 'messages'), msg);

// AI RESPONSE
const aiResponse = await fetch("/api/chat", {
method: "POST",
headers: {
"Content-Type": "application/json",
},
body: JSON.stringify({
message: cleanText,
}),
});

const aiData = await aiResponse.json();

const aiMsg = {
text: aiData.reply,
sender: "JIMJ AI Helpdesk",
createdAt: new Date(),
};

await addDoc(collection(firestoreDb, 'messages'), aiMsg);

} catch (error) {
console.error("Chat message send failed:", error);
setText(cleanText);
}
};



  return <div className={`staff-chat-shell ${mobileChatOpen ? 'mobile-chat-open' : ''}`} style={{display:'grid',gridTemplateColumns:'220px 1fr',gap:14,minHeight:430}}>
    <div className="glass staff-chat-list" style={{padding:10,overflowY:'auto'}}>
      {STAFF_BASE.filter(s=>s.id!==staff.id).map(s=>{
        const threadKey = [staff.id,s.id].sort((a,b)=>a-b).join('-');
        const unread = unreadCounts[threadKey] || 0;
        return <button key={s.id} onClick={()=>{setSelected(s.id);setMobileChatOpen(true);}} style={{width:'100%',display:'flex',alignItems:'center',gap:10,background:selected===s.id?'rgba(99,102,241,0.18)':'transparent',border:'none',borderRadius:10,padding:10,color:'#e2e8f0',textAlign:'left',marginBottom:6}}>
          <StaffAvatar staff={s} profiles={profiles} statuses={statuses} size={34} showStatus/>
          <div style={{flex:1,minWidth:0}}><div style={{fontSize:13,fontWeight:700,overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{s.name}</div><StatusDot status={getStaffStatus(s.id,statuses)}/></div>
          {unread>0&&<span style={{background:'#ef4444',color:'#fff',borderRadius:999,padding:'2px 7px',fontSize:11,flexShrink:0}}>{unread}</span>}
        </button>
      })}
    </div>
    <div className="glass staff-chat-window" style={{display:'flex',flexDirection:'column',overflow:'hidden'}}>
      <div className="staff-chat-header" style={{padding:14,borderBottom:'1px solid rgba(255,255,255,0.08)',display:'flex',gap:10,alignItems:'center'}}>
        <button type="button" className="staff-chat-back" onClick={()=>setMobileChatOpen(false)} style={{display:'none',background:'rgba(255,255,255,0.08)',border:'1px solid rgba(255,255,255,0.12)',color:'#e2e8f0',borderRadius:10,padding:'8px 10px',fontSize:12,fontWeight:700}}>← Back</button>
        <StaffAvatar staff={peer} profiles={profiles} statuses={statuses} size={38} showStatus/>
        <div style={{minWidth:0}}><div style={{fontSize:14,fontWeight:800,color:'#fff',overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>{peer?.name}</div><StatusDot status={getStaffStatus(peer?.id,statuses)}/></div>
      </div>
      <div className="staff-chat-messages" style={{flex:1,padding:14,overflowY:'auto',overflowX:'hidden'}}>
        {messages.map(m=>{
          const mine=m.from===staff.id;
          return <div key={m.id} style={{display:'flex',justifyContent:mine?'flex-end':'flex-start',marginBottom:10}}>
            <div className="staff-chat-bubble" style={{maxWidth:'72%',background:mine?'rgba(99,102,241,0.28)':'rgba(255,255,255,0.07)',border:'1px solid rgba(255,255,255,0.08)',borderRadius:14,padding:'9px 12px'}}>
              <div style={{fontSize:12,color:'rgba(226,232,240,0.45)',marginBottom:3}}>{STAFF_BASE.find(s=>s.id===m.from)?.name || m.fromName || 'Staff'} · {timeAgo(m.at)}</div>
              <div style={{fontSize:13,color:'#e2e8f0',lineHeight:1.4,whiteSpace:'pre-wrap',overflowWrap:'anywhere'}}>{m.text}</div>
            </div>
          </div>
        })}
        {messages.length===0&&<div style={{textAlign:'center',color:'rgba(226,232,240,0.35)',paddingTop:80}}>No messages yet</div>}
        <div ref={messagesEndRef} />
      </div>
      <div className="staff-chat-input" style={{padding:12,borderTop:'1px solid rgba(255,255,255,0.08)',display:'flex',gap:10}}>
        <input value={text} onChange={e=>setText(e.target.value)} onKeyDown={e=>{if(e.key==='Enter'&&!e.shiftKey){e.preventDefault();send();}}} placeholder="Type a message..."/>
        <button className="glow-btn" style={{padding:'10px 18px'}} onClick={send}>Send</button>
      </div>
    </div>
  </div>;
}// ── QUICK ASSIGN DIALOG ───────────────────────────────────────────────────
function QuickAssignDialog({ticket,onClose,onSave,statuses={}}) {
  const [assigneeId,setAssigneeId]=useState(ticket?.assignedTo==="IT Support Team" ? "team" : String(ticket?.assigneeId || STAFF_BASE[0]?.id || ""));
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
          <select value={assigneeId} onChange={e=>setAssigneeId(e.target.value)}>
            <option value="team">IT Support Team</option>
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
        <button className="glow-btn" style={{padding:"10px 20px",fontSize:14}} onClick={()=>onSave(ticket.id,assigneeId,remark)}>Save Assignment</button>
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
  const [portalFeedback, setPortalFeedback] = useState([]);
  const [portalFeedbackLoaded, setPortalFeedbackLoaded] = useState(false);
  const [showPortalFeedback, setShowPortalFeedback] = useState(false);
  const [incidents, setIncidents] = useState([]);
  const [incidentsLoaded, setIncidentsLoaded] = useState(false);
  const [tempIssues, setTempIssues] = useState([]);
  const [tempIssuesLoaded, setTempIssuesLoaded] = useState(false);
  const [tempIssueFilters, setTempIssueFilters] = useState({ status:"All", staff:"All", item:"All", search:"", from:"", to:"" });
  const [dashboardFilter, setDashboardFilter] = useState({ type:"Total", label:"Total" });
  const [adminActionTab, setAdminActionTab] = useState("view_all");
  const [feedbackTicketId, setFeedbackTicketId] = useState("");
  const [dismissedFeedbackTickets, setDismissedFeedbackTickets] = useState([]);
  const [viewTicketId, setViewTicketId] = useState(null);
  const [quickAssignTicketId, setQuickAssignTicketId] = useState(null);
  const [formCat, setFormCat] = useState(null);
  const [smartTicketOpen, setSmartTicketOpen] = useState(false);
  const [mobileOpen, setMobileOpen] = useState(false);
  const [staffProfiles, setStaffProfiles] = useState(() => DB.get("staff_profiles", {}));
  const [staffStatuses, setStaffStatuses] = useState(() => DB.get("staff_statuses", {}));
  const [staffPanel, setStaffPanel] = useState(null);
  const [staffMenuOpen, setStaffMenuOpen] = useState(false);
  const [notificationsEnabled, setNotificationsEnabled] = useState(() => typeof window !== "undefined" && "Notification" in window && Notification.permission === "granted");
  const [showSmartWelcome, setShowSmartWelcome] = useState(() => Boolean(getSavedSession()));
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
    const loadPortalFeedback = async () => {
      try {
        const onlinePortalFeedback = await fetchPortalFeedback();
        if (alive) {
          setPortalFeedback(onlinePortalFeedback);
          setPortalFeedbackLoaded(true);
        }
      } catch (error) {
        console.error("Online portal feedback load failed:", error);
        if (alive) setPortalFeedbackLoaded(true);
      }
    };
    loadPortalFeedback();
    const interval = setInterval(loadPortalFeedback, 15000);
    return () => { alive = false; clearInterval(interval); };
  }, []);
  useEffect(() => {
    let alive = true;
    const loadIncidents = async () => {
      try {
        const onlineIncidents = await fetchIncidents();
        if (alive) {
          setIncidents(onlineIncidents);
          setIncidentsLoaded(true);
        }
      } catch (error) {
        console.error("Online incidents load failed:", error);
        if (alive) setIncidentsLoaded(true);
      }
    };
    loadIncidents();
    const interval = setInterval(loadIncidents, 30000);
    return () => { alive = false; clearInterval(interval); };
  }, []);
  useEffect(() => {
    let alive = true;
    const loadTempIssues = async () => {
      try {
        const onlineTempIssues = await fetchTempIssues();
        if (alive) {
          setTempIssues(onlineTempIssues);
          setTempIssuesLoaded(true);
        }
      } catch (error) {
        console.error("Online temp issues load failed:", error);
        if (alive) setTempIssuesLoaded(true);
      }
    };
    loadTempIssues();
    const interval = setInterval(loadTempIssues, 15000);
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
  useEffect(() => {
    if (!ticketsLoaded) return;
    const now=Date.now();
    let changed=false;
    const updated=tickets.map(ticket=>{
      const info=getEscalationInfo(ticket, now);
      if(!info.overdue) return ticket;
      if(ticket.escalated && Number(ticket.escalationLevel)===info.level) return ticket;
      changed=true;
      return {
        ...ticket,
        escalated:true,
        escalationLevel:info.level,
        escalatedAt:now,
        escalationHistory:[...(ticket.escalationHistory||[]),{level:info.level,at:now,remark:info.label}],
        updatedAt:ticket.updatedAt || now
      };
    });
    if(changed) setTickets(updated);
  }, [ticketsLoaded, tickets.map(t=>`${t.id}:${t.status}:${t.priority}:${t.escalationLevel}`).join("|")]);
  useEffect(() => DB.set("staff_profiles", staffProfiles), [staffProfiles]);
  useEffect(() => DB.set("staff_statuses", staffStatuses), [staffStatuses]);
  useEffect(() => {
    if (!session || !showSmartWelcome) return;
    const timer=setTimeout(()=>setShowSmartWelcome(false),5000);
    return()=>clearTimeout(timer);
  }, [session, showSmartWelcome]);
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
    setShowSmartWelcome(true);
    if (hasStorage()) localStorage.setItem("helpdesk_session", JSON.stringify(sess));
    setPage(sess.type === "admin" ? "dashboard" : sess.type === "staff" ? "staff-dash" : "home");
    toast(`Welcome${sess.name ? `, ${sess.name}` : ""}!`, "success");
  };

  const logoutUser = () => {
    if (hasStorage()) localStorage.removeItem("helpdesk_session");
    setSession(null);
    setShowSmartWelcome(false);
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
  const reloadPortalFeedback = useCallback(async () => {
    try {
      const onlinePortalFeedback = await fetchPortalFeedback();
      setPortalFeedback(onlinePortalFeedback);
      setPortalFeedbackLoaded(true);
      return onlinePortalFeedback;
    } catch (error) {
      console.error("Online portal feedback refresh failed:", error);
      return portalFeedback;
    }
  }, [portalFeedback]);

  const handlePortalFeedbackSubmit = async (entry) => {
    const portalEntry = normalizePortalFeedback({
      id: entry.id || genPortalFeedbackId(),
      name: entry.name || "",
      email: entry.email || "",
      role: entry.role || session?.type || "User",
      rating: Number(entry.rating || 0),
      feedbackType: entry.feedbackType || "General Feedback",
      message: entry.message || "",
      createdAt: Date.now(),
      status: "New",
      reviewed: false,
    });
    try {
      await savePortalFeedback(portalEntry);
      setPortalFeedback(fs => [portalEntry, ...fs.filter(f => f.id !== portalEntry.id)]);
      reloadPortalFeedback().catch(error => console.error("Post-submit portal feedback refresh failed:", error));
      toast("Portal feedback submitted. Thank you!", "success");
      return portalEntry;
    } catch (error) {
      console.error("Portal feedback save failed:", error);
      toast(`Portal feedback save failed: ${error?.message || "Firestore error"}`, "error");
      throw error;
    }
  };
  const handleSaveIncident = async (entry) => {
    try {
      const incident = await saveIncident(entry);
      setIncidents(prev => [incident, ...prev.filter(item => item.id !== incident.id)].sort((a,b)=>(b.updatedAt||0)-(a.updatedAt||0)));
      toast("Campus incident updated","success");
      return incident;
    } catch (error) {
      console.error("Incident save failed:", error);
      toast(`Incident save failed: ${error?.message || "Firestore error"}`, "error");
      throw error;
    }
  };
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
    STAFF_BASE[0] ||
    { id: 0, name: "IT Support Team", role: "IT Support", email: "" };

  const now = Date.now();
  const allStaffWatchers = STAFF_BASE.map(staff => ({ id: staff.id, name: staff.name, email: staff.email, role: staff.role }));
  const aiTicketSummary = await generateTicketAiSummary(form);

  const newTicket = {
      id: genId(),
      name: form.name,
      email: form.email,
      dept: form.dept,
      mobile: form.mobile || "",
      location: form.location || "",
      category: form.category,
      subCategory: form.subCategory || "",
      description: form.description,
      priority: form.priority || "Medium",
      source: form.source || "Portal",
      assignmentGroup: form.assignmentGroup || getAssignmentGroup(form.category),
      issueSummary: form.issueSummary || "",
      recommendedAction: form.recommendedAction || "",
      aiSummary: aiTicketSummary.aiSummary,
      suggestedAction: aiTicketSummary.suggestedAction,
      notes: form.notes || "",
      status: "Assigned",
      assigneeId: assignee.id,
      assigneeName: assignee.name,
      assignedTo: assignee.name,
      watchers: allStaffWatchers,
      notifiedStaff: allStaffWatchers,
      createdAt: now,
      updatedAt: now,
      closedAt: null,
      feedbackSubmitted: false,
      closingRemarks: "",
      remoteSupportRequested: false,
      remoteSupportTool: "",
      remoteSupportId: "",
      remoteSupportNote: "",
      remoteSupportRequestedAt: null,
      escalated: false,
      escalationLevel: 0,
      escalatedAt: null,
      escalationHistory: [],
      comments: [],
      timeline: [
        {
          action: "Created",
          by: form.name,
          at: now,
          remark: form.source ? `Source: ${form.source}${form.assignmentGroup ? ` · ${form.assignmentGroup}` : ""}` : ""
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
      notifyTicketCreated(newTicket);
      showBrowserNotification("Ticket created", `${newTicket.id} has been created successfully.`);

      toast("Ticket created successfully", "success");
      return newTicket;
    } catch (error) {
      console.error("Ticket create/save failed:", error);
      toast(`Ticket save failed: ${error?.message || "Firestore error"}`, "error");
      return null;
    }
  };

  const reloadTempIssues = useCallback(async () => {
    try {
      const onlineTempIssues = await fetchTempIssues();
      setTempIssues(onlineTempIssues);
      setTempIssuesLoaded(true);
      return onlineTempIssues;
    } catch (error) {
      console.error("Online temp issue refresh failed:", error);
      return tempIssues;
    }
  }, [tempIssues]);

  const handleSaveTempIssue = async (entry) => {
    if (session?.type !== "user") {
      toast("Only users can create Temp Issue requests", "error");
      throw new Error("Only users can create Temp Issue requests");
    }
    const targetStaff = STAFF_BASE.find(s => s.name === (entry.requestedToStaff || entry.requestToStaff));
    if (!targetStaff) {
      toast("Please select IT staff", "error");
      throw new Error("Missing requested IT staff");
    }
    const now = Date.now();
    const request = normalizeTempIssue({
      ...entry,
      requestId: entry.requestId || `TI-${now.toString(36).toUpperCase()}-${Math.random().toString(36).slice(2,5).toUpperCase()}`,
      userId: session?.email || entry.userEmail || "",
      userName: entry.userName || session?.name || "",
      userEmail: entry.userEmail || session?.email || "",
      requestedToStaff: targetStaff.name,
      requestToStaff: targetStaff.name,
      requestedToStaffEmail: targetStaff.email,
      requestToStaffEmail: targetStaff.email,
      permissionApprovedBy: entry.permissionApprovedBy || entry.permissionBy || "",
      permissionBy: entry.permissionApprovedBy || entry.permissionBy || "",
      status: "Pending Approval",
      returnDate: "",
      createdAt: now,
      updatedAt: now,
      requestedBy: session?.email || entry.userEmail || "",
      requestHistory: [
        { action: "Created", by: entry.userName || session?.email || "User", at: now, remark: "Temp issue request submitted" },
        { action: "New Temp Issue Request", by: "System", at: now, remark: `Assigned to ${targetStaff.name}` },
      ],
      notifications: [
        { type: "New request", to: targetStaff.email, message: "New Temp Issue Request", at: now, read: false },
        { type: "New request", to: "admin@jaipuria.ac.in", message: "New Temp Issue Request", at: now, read: false },
      ],
    });
    try {
      await saveTempIssue(request);
      setTempIssues(prev => [request, ...prev.filter(it => it.requestId !== request.requestId)]);
      simulateEmail(targetStaff.email, "New Temp Issue Request", `${request.userName} requested ${request.item}. Permission approved by: ${request.permissionApprovedBy}.`);
      simulateEmail("admin@jaipuria.ac.in", "New Temp Issue Request", `${request.userName} requested ${request.item} from ${targetStaff.name}.`);
      await reloadTempIssues();
      toast("Temp Issue request submitted", "success");
      return request;
    } catch (error) {
      console.error("Temp issue save failed:", error);
      toast(`Temp Issue save failed: ${error?.message || "Firestore error"}`, "error");
      throw error;
    }
  };

  const handleTempIssueAction = async (requestId, action, actorName, remarks = "", meta = {}) => {
    const existing = tempIssues.find(t => t.requestId === requestId);
    if (!existing) return;
    const now = Date.now();
    const actor = actorName || (session?.type === "admin" ? "Admin" : session?.name || session?.email || "User");
    const updated = {
      ...existing,
      updatedAt: now,
      remarks: remarks || existing.remarks || "",
      requestHistory: Array.isArray(existing.requestHistory) ? [...existing.requestHistory] : [],
      notifications: Array.isArray(existing.notifications) ? [...existing.notifications] : [],
    };
    const pushHistory = (label, message = "") => {
      updated.requestHistory.push({ action: label, by: actor, at: now, remark: remarks || message || "" });
    };
    const pushNotify = (type, message, extraTo = []) => {
      const staffEmail = updated.returnToStaffEmail || updated.requestedToStaffEmail || updated.requestToStaffEmail || STAFF_BASE.find(s => s.name === (updated.returnToStaff || updated.requestedToStaff || updated.requestToStaff))?.email || "";
      const targets = [updated.userEmail, staffEmail, "admin@jaipuria.ac.in", ...extraTo].filter(Boolean);
      targets.forEach(to => updated.notifications.push({ type, to, message, at: now, read: false }));
      if (staffEmail) simulateEmail(staffEmail, `Temp Issue ${type}`, message);
      simulateEmail("admin@jaipuria.ac.in", `Temp Issue ${type}`, message);
    };

    if (action === "approve") {
      if (existing.status !== "Pending Approval") return;
      updated.status = "Approved";
      updated.approvedBy = actor;
      updated.approvedAt = now;
      pushHistory("Approved");
      pushNotify("Approved", `${updated.requestId} approved by ${actor}`);
    } else if (action === "reject") {
      if (existing.status !== "Pending Approval") return;
      updated.status = "Rejected";
      updated.rejectedBy = actor;
      updated.rejectedAt = now;
      pushHistory("Rejected");
      pushNotify("Rejected", `${updated.requestId} rejected by ${actor}`);
    } else if (action === "issue") {
      if (existing.status !== "Approved") return;
      updated.status = "Issued";
      updated.issuedBy = actor;
      updated.issuedAt = now;
      pushHistory("Issued");
      pushNotify("Issued", `${updated.requestId} item issued by ${actor}`);
    } else if (action === "notIssued") {
      if (existing.status !== "Approved") return;
      updated.status = "Not Issued";
      updated.notIssuedBy = actor;
      updated.notIssuedAt = now;
      pushHistory("Not Issued");
      pushNotify("Not Issued", `${updated.requestId} marked not issued by ${actor}`);
    } else if (action === "requestReturn") {
      if (existing.status !== "Issued") return;
      if (!meta.returnToStaff || !meta.returnToStaffEmail) {
        toast("Please select return IT staff", "error");
        return;
      }
      updated.status = "Return Requested";
      updated.returnToStaff = meta.returnToStaff;
      updated.returnToStaffEmail = meta.returnToStaffEmail;
      updated.returnRequestedAt = now;
      updated.returnRemarks = remarks || updated.returnRemarks || "";
      pushHistory("Return Requested", `User requested return to ${meta.returnToStaff}`);
      pushNotify("Return request received", "Temp Issue return request received", [meta.returnToStaffEmail]);
    } else if (action === "acceptReturn") {
      if (existing.status !== "Return Requested") return;
      updated.status = "Returned";
      updated.returnAcceptedBy = actor;
      updated.returnAcceptedAt = now;
      updated.returnedBy = actor;
      updated.returnedAt = now;
      updated.returnRemarks = remarks || updated.returnRemarks || "";
      pushHistory("Return Received", remarks || `Item handover to ${actor}`);
      pushNotify("Return accepted", `${updated.requestId} returned to ${actor}`);
    } else if (action === "rejectReturn") {
      if (existing.status !== "Return Requested") return;
      updated.status = "Return Rejected";
      updated.returnRejectedBy = actor;
      updated.returnRejectedAt = now;
      updated.returnRemarks = remarks || updated.returnRemarks || "";
      pushHistory("Return Rejected");
      pushNotify("Return rejected", `${updated.requestId} return rejected by ${actor}`);
    } else if (action === "forceClose") {
      updated.status = "Force Closed";
      updated.forceClosedBy = actor;
      updated.forceClosedAt = now;
      pushHistory("Force Closed");
      pushNotify("Force closed", `${updated.requestId} force closed by ${actor}`);
    } else if (action === "editStatus") {
      if (session?.type !== "admin") return;
      const nextStatus = prompt("Enter new status", existing.status);
      const allowed = ["Pending Approval","Approved","Issued","Return Requested","Returned","Rejected","Not Issued","Return Rejected","Force Closed"];
      if (!nextStatus || !allowed.includes(nextStatus)) {
        toast("Invalid status", "error");
        return;
      }
      updated.status = nextStatus;
      pushHistory("Status Edited", `Status changed to ${nextStatus}`);
      pushNotify("Status updated", `${updated.requestId} status changed to ${nextStatus} by ${actor}`);
    }

    try {
      await saveTempIssue(updated);
      setTempIssues(prev => prev.map(t => t.requestId === requestId ? updated : t));
      await reloadTempIssues();
      toast(`Temp Issue ${updated.status.toLowerCase()}`, "success");
    } catch (error) {
      console.error("Temp issue status update failed:", error);
      toast(`Temp Issue update failed: ${error?.message || "Firestore error"}`, "error");
      throw error;
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
    const assignToTeam = assigneeId === "team" || Number(assigneeId) === 0;
    const numericAssigneeId = assignToTeam ? 0 : Number(assigneeId);

    setTickets(ts => ts.map(t => {
      if (t.id !== ticketId) return t;
      newAssignee = assignToTeam ? { id:0, name:"IT Support Team", role:"IT Support", email:"" } : STAFF_BASE.find(s => s.id === numericAssigneeId);
      previousAssignee = STAFF_BASE.find(s => s.id === t.assigneeId);
      const cleanRemark = remark.trim();
      const actor = session?.type === "staff" ? session.name || "IT Staff" : "Admin";
      const timelineEntry = {
        action: `Quick assigned to ${newAssignee?.name || "Unassigned"}`,
        remark: cleanRemark,
        at: Date.now(),
        by: actor,
      };
      const comments = cleanRemark
        ? [...(t.comments || []), { text: cleanRemark, at: Date.now(), by: actor }]
        : (t.comments || []);
      assignedTicket = {
        ...t,
        assigneeId:numericAssigneeId,
        assigneeName:newAssignee?.name || t.assigneeName,
        assignedTo:newAssignee?.name || t.assignedTo,
        status: t.status === "Open" ? "Assigned" : t.status,
        updatedAt: Date.now(),
        comments,
        timeline: [...(t.timeline || []), timelineEntry],
      };
      return assignedTicket;
    }));

    if (assignedTicket) {
      emailTicketAssigned(assignedTicket, newAssignee, previousAssignee, session?.type === "staff" ? session.name || "IT Staff" : "Admin", remark.trim());
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

  const handleFirstLoginComplete = async (hash) => {
    const staff = STAFF_BASE.find(s => s.id === session.staffId);
    if (!staff) return;

    const existingProfile = await fetchStaffProfile(staff.id).catch(error => {
      console.error("Staff first-login profile lookup failed:", error);
      return null;
    });

    const existingHash = existingProfile?.passwordHash || existingProfile?.password || "";
    const finalHash = existingHash || hash;

    if (!existingHash) {
      await saveStaffProfile(staff.id, {
        ...(existingProfile || {}),
        id: String(staff.id),
        staffId: staff.id,
        email: staff.email,
        name: staff.name,
        role: staff.role,
        permissions: staff.permissions,
        passwordHash: hash,
        passwordSet: true,
        passwordUpdatedAt: Date.now(),
      });
      setStaffProfiles(p => ({...p, [staff.id]: {...(p[staff.id] || {}), passwordHash: hash, passwordSet: true}}));
    }

    const staffPasswords = DB.get("staff_passwords", {});
    staffPasswords[staff.id] = finalHash;
    DB.set("staff_passwords", staffPasswords);

    const staffSession = {
      type: "staff",
      staffId: staff.id,
      email: staff.email,
      name: staff.name,
      role: staff.role,
      permissions: staff.permissions,
    };

    const completedSession = {...staffSession, passwordSet: true, requiresPasswordSetup: false};
    clearStaffPasswordSetupStorage();
    setSession(completedSession);
    if (hasStorage()) localStorage.setItem("helpdesk_session", JSON.stringify(completedSession));
    setPage("staff-dash");
  };

  if (!session) {
  return (
    <>
      <style>{CSS}</style>
      <Landing onLogin={handleLogin} tickets={tickets} />
      <PortalFeedbackChrome onOpen={() => setShowPortalFeedback(true)} />
      {showPortalFeedback && (
        <Modal title="Portal Feedback" onClose={() => setShowPortalFeedback(false)}>
          <PortalFeedbackForm session={session} onSubmit={handlePortalFeedbackSubmit} toast={toast} onClose={() => setShowPortalFeedback(false)} />
        </Modal>
      )}
      <Toast toasts={toasts} remove={remove} />
    </>
  );
}

  if (session.type === "staff_firstlogin") {
    return (
      <>
        <style>{CSS}</style>
        <SetPasswordScreen staff={session.staff} onComplete={handleFirstLoginComplete} toast={toast} />
        <PortalFeedbackChrome onOpen={() => setShowPortalFeedback(true)} />
        {showPortalFeedback && (
          <Modal title="Portal Feedback" onClose={() => setShowPortalFeedback(false)}>
            <PortalFeedbackForm session={session} onSubmit={handlePortalFeedbackSubmit} toast={toast} onClose={() => setShowPortalFeedback(false)} />
          </Modal>
        )}
        <Toast toasts={toasts} remove={remove} />
      </>
    );
  }

  const isAdmin = session.type === "admin";
  const isStaff = session.type === "staff";
  const myTickets = isAdmin || isStaff ? tickets : tickets.filter(t => t.email === session.email);
  const userFeedbackPendingCount = !isAdmin && !isStaff ? myTickets.filter(isTicketFeedbackPending).length : 0;
  const quickAssignTicket = tickets.find(t => t.id === quickAssignTicketId);
  const linkedFeedbackTicket = tickets.find(t => t.id === feedbackTicketId);
  const selectedViewTicket = tickets.find(t => t.id === viewTicketId) || null;
  const pendingFeedbackTicket = null;

  const handleStaffMenuAction = (id) => {
    setStaffMenuOpen(false);
    setMobileOpen(false);
    if (id === "logout") {
      logoutUser();
      return;
    }
    if (id === "know-staff") {
      setPage("know-staff");
      return;
    }
    setStaffPanel(id);
  };

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

              const nextHash = await hashPassword(newPwd);
              const existingProfile = await fetchStaffProfile(staff.id).catch(error => { console.error("Admin staff password profile lookup failed:", error); return null; });
              await saveStaffProfile(staff.id, {...(existingProfile||{}),id:String(staff.id),staffId:staff.id,email:staff.email,name:staff.name,role:staff.role,passwordHash:nextHash,passwordSet:true,passwordUpdatedAt:Date.now()});
              const passwords = DB.get("staff_passwords", {});
              passwords[staff.id] = nextHash;
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

  const handleAdminActionTab = (action) => {
    if(action==="export") {
      downloadTicketCsv(tickets);
      toast("Tickets exported to CSV","success");
      setAdminActionTab("export");
      return;
    }
    setAdminActionTab(action);
    if(action==="view_all") setDashboardFilter({ type:"Total", label:"Total" });
    if(action==="assign") setDashboardFilter({ type:"Assign", label:"Assignment Queue" });
    if(action==="close") setDashboardFilter({ type:"Closed", label:"Closed / Resolved" });
  };

  const renderPage = () => {
    if (isAdmin) {
      if (page === "dashboard") {
        const filteredDashboardTickets = dashboardFilter.type === "Total"
          ? tickets
          : dashboardFilter.type === "Assign"
            ? filterActionTickets("assign", tickets, tickets)
          : dashboardFilter.type === "Open"
            ? tickets.filter(t => t.status === "Open")
            : dashboardFilter.type === "In Progress"
              ? tickets.filter(t => t.status === "In Progress")
              : dashboardFilter.type === "Resolved"
                ? tickets.filter(t => t.status === "Resolved" || t.status === "Closed")
                : dashboardFilter.type === "Critical"
                  ? tickets.filter(t => t.priority === "Critical")
                  : dashboardFilter.type === "Closed"
                    ? tickets.filter(t => t.status === "Closed" || t.status === "Resolved")
                    : tickets;
        const escalatedTickets = tickets.filter(t => getEscalationInfo(t).overdue).sort((a,b)=>getEscalationInfo(b).level-getEscalationInfo(a).level || (b.createdAt||0)-(a.createdAt||0));

        return (
          <div style={{display:"flex",flexDirection:"column",gap:24}}>
            <SmartWelcome session={session} visible={showSmartWelcome} />
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",flexWrap:"wrap",gap:12}}>
              <div>
                <h2 style={{fontFamily:"Syne",fontSize:22,fontWeight:700,color:"#e2e8f0"}}>Admin Dashboard</h2>
                <p style={{fontSize:14,color:"rgba(226,232,240,0.5)"}}>{new Date().toLocaleDateString("en-IN",{weekday:"long",year:"numeric",month:"long",day:"numeric"})}</p>
              </div>
              <button className="glow-btn" onClick={() => setFormCat("")}>+ New Ticket</button>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(155px,1fr))",gap:14}}>
              {[["Total",tickets.length,"🎫","#818cf8"],["Open",tickets.filter(t=>t.status==="Open").length,"🔵","#60a5fa"],["In Progress",tickets.filter(t=>t.status==="In Progress").length,"🟡","#fbbf24"],["Resolved",tickets.filter(t=>t.status==="Resolved"||t.status==="Closed").length,"🟢","#34d399"],["Critical",tickets.filter(t=>t.priority==="Critical").length,"🔴","#f87171"],["Closed",tickets.filter(t=>t.status==="Closed").length,"⚫","#6b7280"]].map(([l,v,i,c]) => (
                <StatCard
                  key={l}
                  label={l}
                  value={v}
                  icon={i}
                  color={c}
                  onClick={() => setDashboardFilter({ type: l, label: l })}
                  style={{cursor: 'pointer'}}
                />
              ))}
            </div>
            <div className="glass" style={{padding:"14px 16px",display:"flex",justifyContent:"space-between",alignItems:"center",gap:12,flexWrap:"wrap"}}>
              <div>
                <div style={{fontSize:13,fontWeight:900,color:"#fff"}}>Admin Actions</div>
                <div style={{fontSize:12,color:"rgba(226,232,240,.5)",marginTop:3}}>Filter, assign, export, or manage users.</div>
              </div>
              <ActionTabs permissions={["view_all","assign","close","export","manage_users"]} active={adminActionTab} onSelect={handleAdminActionTab} />
            </div>
            {adminActionTab==="manage_users"&&<ManageUsersPanel />}
            <IncidentManager incidents={incidents} onSave={handleSaveIncident} toast={toast} />
            <div className="glass" style={{padding:"18px",display:"grid",gap:14}}>
              <div style={{display:"flex",justifyContent:"space-between",gap:12,alignItems:"center",flexWrap:"wrap"}}>
                <div>
                  <h3 style={{fontFamily:"Syne",fontSize:16,fontWeight:900,color:"#e2e8f0"}}>Escalated Tickets</h3>
                  <p style={{fontSize:12,color:"rgba(226,232,240,.52)",marginTop:4}}>Overdue tickets are highlighted for staff/admin follow-up.</p>
                </div>
                <span className="tag" style={{background:"rgba(239,68,68,.16)",color:"#fecaca"}}>{escalatedTickets.length} overdue</span>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(240px,1fr))",gap:10}}>
                {escalatedTickets.slice(0,6).map(ticket=><TicketCard key={ticket.id} ticket={ticket} onView={setViewTicketId} showFeedbackUnread={isTicketFeedbackUnread(ticket,true,false)} />)}
                {escalatedTickets.length===0&&<EmptyState message="No escalated tickets right now." icon="✅" />}
              </div>
            </div>
            {dashboardFilter.type !== "Total" && (
              <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',gap:14,flexWrap:'wrap'}}>
                <h3 style={{fontFamily:"Syne",fontSize:16,fontWeight:700,color:"#e2e8f0",margin:0}}>{`Showing ${dashboardFilter.label} Tickets (${filteredDashboardTickets.length})`}</h3>
                <button className="glow-btn" onClick={() => setDashboardFilter({ type: "Total", label: "Total" })}>Show All</button>
              </div>
            )}
            {dashboardFilter.type === "Total" && adminActionTab!=="manage_users" ? (
              <>
                <h3 style={{fontFamily:"Syne",fontSize:16,fontWeight:700,color:"#e2e8f0"}}>Recent Tickets</h3>
                <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(240px,1fr))",gap:10}}>
                  {tickets.slice(0,6).map(t => (
                    <div key={t.id} style={{position:"relative"}}>
                      <TicketCard ticket={t} onView={setViewTicketId} showFeedbackUnread={isTicketFeedbackUnread(t,true,false)} />
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
              </>
            ) : (
              <TicketsTable tickets={filteredDashboardTickets} onView={setViewTicketId} isAdmin onDelete={handleDeleteTicket} emptyKind={dashboardFilter.type} />
            )}
          </div>
        );
      }
      if (page === "tickets") return <TicketsTable tickets={tickets} onView={setViewTicketId} isAdmin onDelete={handleDeleteTicket} />;
      if (page === "analytics") return <Analytics tickets={tickets} />;
      if (page === "feedback") return <AdminFeedbackPage feedback={feedback} setFeedback={setFeedback} toast={toast} />;
      if (page === "portal-feedback") return <AdminPortalFeedbackPage portalFeedback={portalFeedback} setPortalFeedback={setPortalFeedback} toast={toast} />;
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
      return <StaffPanel staffId={session.staffId} tickets={tickets} setTickets={setTickets} toast={toast} onViewTicket={setViewTicketId} onQuickAssign={setQuickAssignTicketId} permissions={session.permissions} staffProfiles={staffProfiles} staffStatuses={staffStatuses} showWelcome={showSmartWelcome} />;
    }

    if (page === "home") return <CategoryGrid onSelect={cat => setFormCat(cat)} onSmartTicket={()=>setSmartTicketOpen(true)} session={session} showWelcome={showSmartWelcome} />;
    if (page === "my-tickets") return (
      <div>
        <h2 style={{fontFamily:"Syne",fontSize:22,fontWeight:700,color:"#e2e8f0",marginBottom:20}}>My Tickets</h2>
        {userFeedbackPendingCount===0&&<div className="glass2" style={{padding:"12px 14px",marginBottom:14,color:"#bbf7d0",borderColor:"rgba(16,185,129,.22)",background:"rgba(16,185,129,.07)",fontSize:13,fontWeight:800}}>No pending feedback. You're all caught up.</div>}
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(240px,1fr))",gap:10}}>
          {myTickets.map(t => <TicketCard key={t.id} ticket={t} onView={setViewTicketId} showFeedbackPending={isTicketFeedbackPending(t)} />)}
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
    if (page === "temp-issue") return <TempIssuePanel session={session} tempIssues={tempIssues} tempIssuesLoaded={tempIssuesLoaded} filters={tempIssueFilters} setFilters={setTempIssueFilters} onSubmit={handleSaveTempIssue} onAction={handleTempIssueAction} toast={toast} />;

    return null;
  };

  return (
    <>
      <style>{CSS}</style>
      <div className="theme-glow" />
      <div className="app-shell" style={{display:"flex",minHeight:"100vh"}}>
        <Sidebar current={page} onChange={setPage} isAdmin={isAdmin} isStaff={isStaff} tickets={tickets} feedback={feedback} mobileOpen={mobileOpen} setMobileOpen={setMobileOpen} onStaffAction={handleStaffMenuAction} feedbackPendingCount={userFeedbackPendingCount} />
        <div className="app-main" style={{flex:1,display:"flex",flexDirection:"column",minWidth:0}}>
          <div className="app-header" style={{padding:"14px 24px",borderBottom:"1px solid rgba(255,255,255,0.07)",display:"flex",justifyContent:"space-between",alignItems:"center",background:"rgba(10,10,20,0.9)",backdropFilter:"blur(20px)",position:"sticky",top:0,zIndex:10}}>
            <div className="header-identity" style={{display:"flex",alignItems:"center",gap:12,position:"relative"}}>
              <button onClick={() => { const isMobile = typeof window !== "undefined" && window.innerWidth <= 768; if (isMobile) { setStaffMenuOpen(false); setMobileOpen(o => !o); } else if (isStaff) { setMobileOpen(false); setStaffMenuOpen(o=>!o); } else { setStaffMenuOpen(false); setMobileOpen(o => !o); } }} style={{background:"rgba(255,255,255,0.06)",border:"1px solid rgba(255,255,255,0.1)",color:"#e2e8f0",width:38,height:38,borderRadius:10,fontSize:18,display:"flex",alignItems:"center",justifyContent:"center"}}>☰</button>
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
              <NotificationButton toast={toast} enabled={notificationsEnabled} setEnabled={setNotificationsEnabled} />
              {!isStaff&&<button onClick={logoutUser} style={{background:"rgba(239,68,68,0.1)",border:"1px solid rgba(239,68,68,0.3)",color:"#f87171",padding:"6px 14px",borderRadius:8,fontSize:13}}>Logout</button>}
            </div>
          </div>
          <CampusIncidentBanner incidents={incidents} />
          <div className="app-content" style={{padding:"24px 28px",flex:1,overflowY:"auto"}}>{renderPage()}</div>
        </div>
      </div>

      {formCat !== null && (
        <Modal title="Raise IT Support Ticket" onClose={() => setFormCat(null)}>
          <TicketForm userEmail={session?.email} initialCategory={formCat} onSubmit={handleNewTicket} onCancel={() => setFormCat(null)} toast={toast} />
        </Modal>
      )}

      {smartTicketOpen && (
        <Modal title="AI Smart Ticket" onClose={() => setSmartTicketOpen(false)} wide>
          <SmartTicketModal session={session} onSubmit={handleNewTicket} onClose={() => setSmartTicketOpen(false)} toast={toast} />
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
      {viewTicketId && selectedViewTicket && (
        <Modal title={`Ticket - ${selectedViewTicket.id}`} onClose={() => setViewTicketId(null)}>
          <TicketDetail
            ticketId={selectedViewTicket.id}
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

      <PortalFeedbackChrome onOpen={() => setShowPortalFeedback(true)} />
      <PWAInstallPrompt />
      {!isAdmin && !isStaff && <AIHelpdeskChat session={session} onCreateTicket={handleNewTicket} />}
      {showPortalFeedback && (
        <Modal title="Portal Feedback" onClose={() => setShowPortalFeedback(false)}>
          <PortalFeedbackForm session={session} onSubmit={handlePortalFeedbackSubmit} toast={toast} onClose={() => setShowPortalFeedback(false)} />
        </Modal>
      )}
      <Toast toasts={toasts} remove={remove} />
    </>
  );
}













































































































