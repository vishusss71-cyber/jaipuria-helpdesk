import * as XLSX from "xlsx";
import { saveAs } from "file-saver";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import emailjs from '@emailjs/browser';
const sendTicketEmail = async (ticket, user) => {
  try {
    await emailjs.send(
      'service_ctyqqbc',
      'template_vuv4jtd',
      {
          user_name: user?.name || 'User',
          to_email: user?.email,
          ticket_id: ticket.id,
          issue: ticket.issue || ticket.description || ticket.category || "IT Helpdesk Ticket",
      },
      'N9OlDxPyO0uf_IlxJ'
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
  { id: 1, name: "Raj Parkash Singh", role: "Manager", email: "raj.singh@jaipuria.ac.in", avatar: "RPS", color: "#6366f1", permissions: ["view_all","assign","close","export","manage_users"] },
  { id: 2, name: "Rohit Jangid", role: "Executive", email: "rohit.jangid@jaipuria.ac.in", avatar: "RJ", color: "#0ea5e9", permissions: ["view_assigned","close","comment"] },
  { id: 3, name: "Vishal Swami", role: "Asst. Manager", email: "vishal.swami@jaipuria.ac.in", avatar: "VS", color: "#10b981", permissions: ["view_assigned","assign","close","comment"] },
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

const DEPTS = ["IT Department","Computer Science","Electronics","Administration","Library","Finance","HR","Management","Maintenance","Faculty"];
const PRIORITIES = ["Low","Medium","High","Critical"];
const STATUSES = ["Open","Assigned","In Progress","Resolved","Closed"];
const SLA_HOURS = { Low:72, Medium:48, High:24, Critical:4 };

// ── LOCAL STORAGE DB ──────────────────────────────────────────────────────
const hasStorage = () => typeof window !== "undefined" && typeof window.localStorage !== "undefined";
const safeJsonParse = (value, fallback = null) => {
  try { return value ? JSON.parse(value) : fallback; } catch { return fallback; }
};
const DB = {
  get: (k, def) => { try { if (!hasStorage()) return def; const v = localStorage.getItem("helpdesk_"+k); return safeJsonParse(v, def); } catch { return def; } },
  set: (k, v) => { try { if (hasStorage()) localStorage.setItem("helpdesk_"+k, JSON.stringify(v)); } catch {} },
};

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

function emailTicketClosed(ticket, assignee, closedBy="Admin") {
  const duration = formatDuration(ticket.closedAt - ticket.createdAt);
  const category = categoryLabel(ticket.category);
  const body = `
Dear ${ticket.name},

Your IT support ticket has been closed.

━━━━━━━━━━━━━━━━━━━━━━━━━
TICKET CLOSURE DETAILS
━━━━━━━━━━━━━━━━━━━━━━━━━
Ticket ID           : ${ticket.id}
Category            : ${category}
Priority            : ${ticket.priority}
Final Status        : Closed
Closed By           : ${closedBy}
Closed Date/Time    : ${fmtDate(ticket.closedAt)}
Resolution Duration : ${duration}
Closing Remarks     : ${ticket.closingRemarks || "Issue resolved successfully."}
━━━━━━━━━━━━━━━━━━━━━━━━━

Thank you for using Jaipuria Institute of Management IT Support Portal.

Regards,
Jaipuria Institute of Management IT Support Team
  `.trim();

  simulateEmail(ticket.email, `[${ticket.id}] Ticket Closed`, body);

  if (assignee?.email) {
    simulateEmail(assignee.email, `[${ticket.id}] Assigned Ticket Closed`, `${body}\n\nStaff copy: this ticket was assigned to you.`);
  }

  simulateEmail("admin@jaipuria.ac.in", `[${ticket.id}] Ticket Closed Notification`, `${body}\n\nAdmin copy: ticket closure has been recorded.`);
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
}// ── GLOBAL CSS ─────────────────────────────────────────────────────────────
const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Sans:wght@300;400;500;600&display=swap');
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'DM Sans',sans-serif;background:#0a0a0f;color:#e2e8f0;min-height:100vh}
::-webkit-scrollbar{width:6px;height:6px}
::-webkit-scrollbar-track{background:rgba(255,255,255,0.03)}
::-webkit-scrollbar-thumb{background:rgba(255,255,255,0.12);border-radius:3px}
input,select,textarea{background:rgba(255,255,255,0.06);border:1px solid rgba(255,255,255,0.1);color:#e2e8f0;border-radius:10px;padding:10px 14px;font-family:'DM Sans',sans-serif;font-size:14px;width:100%;outline:none;transition:all .2s}
input:focus,select:focus,textarea:focus{border-color:rgba(99,102,241,0.6);background:rgba(99,102,241,0.08);box-shadow:0 0 0 3px rgba(99,102,241,0.12)}
input::placeholder,textarea::placeholder{color:rgba(226,232,240,0.3)}
select option{background:#1a1a2e;color:#e2e8f0}
button{cursor:pointer;font-family:'DM Sans',sans-serif;transition:all .2s}
.glass{background:rgba(255,255,255,0.04);backdrop-filter:blur(20px);border:1px solid rgba(255,255,255,0.08);border-radius:16px}
.glass2{background:rgba(255,255,255,0.06);backdrop-filter:blur(20px);border:1px solid rgba(255,255,255,0.1);border-radius:12px}
.glow-btn{background:linear-gradient(135deg,#6366f1,#8b5cf6);border:none;color:#fff;padding:12px 28px;border-radius:12px;font-size:15px;font-weight:600;letter-spacing:.3px}
.glow-btn:hover{transform:translateY(-1px);box-shadow:0 8px 30px rgba(99,102,241,0.4)}
.glow-btn:active{transform:translateY(0)}
.glow-btn:disabled{opacity:.5;transform:none;box-shadow:none;cursor:not-allowed}
.danger-btn{background:linear-gradient(135deg,#ef4444,#dc2626);border:none;color:#fff;padding:10px 20px;border-radius:10px;font-size:14px;font-weight:600}
.danger-btn:hover{transform:translateY(-1px);box-shadow:0 8px 24px rgba(239,68,68,0.4)}
.success-btn{background:linear-gradient(135deg,#10b981,#059669);border:none;color:#fff;padding:10px 20px;border-radius:10px;font-size:14px;font-weight:600}
.success-btn:hover{transform:translateY(-1px);box-shadow:0 8px 24px rgba(16,185,129,0.4)}
.tag{display:inline-flex;align-items:center;gap:5px;padding:4px 10px;border-radius:20px;font-size:12px;font-weight:500}
.pwd-input-wrap{position:relative}
.pwd-input-wrap input{padding-right:44px}
.pwd-toggle{position:absolute;right:12px;top:50%;transform:translateY(-50%);background:none;border:none;color:rgba(226,232,240,0.4);font-size:18px;padding:0;line-height:1;cursor:pointer}
.pwd-toggle:hover{color:rgba(226,232,240,0.8)}
@keyframes fadeUp{from{opacity:0;transform:translateY(20px)}to{opacity:1;transform:translateY(0)}}
@keyframes fadeIn{from{opacity:0}to{opacity:1}}
@keyframes pulse{0%,100%{opacity:1}50%{opacity:.5}}
@keyframes spin{to{transform:rotate(360deg)}}
@keyframes bounce{0%,100%{transform:scale(1)}50%{transform:scale(1.15)}}
@keyframes slideDown{from{opacity:0;transform:translateY(-10px)}to{opacity:1;transform:translateY(0)}}
@keyframes confetti{0%{transform:translateY(0) rotate(0);opacity:1}100%{transform:translateY(100px) rotate(720deg);opacity:0}}
.fade-up{animation:fadeUp .4s ease forwards}
.fade-in{animation:fadeIn .3s ease forwards}
.pulse{animation:pulse 2s infinite}
.spin{animation:spin 1s linear infinite}
.slide-down{animation:slideDown .3s ease forwards}
`;

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
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.75)",backdropFilter:"blur(8px)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}
      onClick={e=>e.target===e.currentTarget&&onClose&&onClose()}>
      <div className="glass fade-up" style={{width:"100%",maxWidth:wide?860:720,maxHeight:"92vh",overflow:"hidden",display:"flex",flexDirection:"column"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"20px 24px 16px",borderBottom:"1px solid rgba(255,255,255,0.08)",flexShrink:0}}>
          <h2 style={{fontFamily:"Syne",fontSize:18,fontWeight:700,color:"#e2e8f0"}}>{title}</h2>
          {onClose&&<button onClick={onClose} style={{background:"rgba(255,255,255,0.08)",border:"1px solid rgba(255,255,255,0.1)",color:"#e2e8f0",width:32,height:32,borderRadius:8,fontSize:20,display:"flex",alignItems:"center",justifyContent:"center",lineHeight:1}}>×</button>}
        </div>
        <div style={{padding:"20px 24px",overflowY:"auto",flex:1}}>{children}</div>
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
    onConfirm(remarks);
    setLoading(false);
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
function TicketDetail({ticketId,tickets,setTickets,onClose,isAdmin,isStaff,staffId,staffName,toast}) {
  const ticket=tickets.find(t=>t.id===ticketId)||{};
  const [comment,setComment]=useState("");
  const [editStatus,setEditStatus]=useState(ticket.status);
  const [editAssignee,setEditAssignee]=useState(ticket.assigneeId);
  const [showClose,setShowClose]=useState(false);
  const assignee=STAFF_BASE.find(s=>s.id===ticket.assigneeId);
  const cat=CATEGORIES.find(c=>c.id===ticket.category);

  const 
  updateTicket=(changes,auditAction,remark="")=>{
    setTickets(ts=>ts.map(t=>{
      if(t.id!==ticketId) return t;
      const actor=isAdmin?"Admin":staffName||"User";
      const tl=[...(t.timeline||[]),{action:auditAction,remark,at:Date.now(),by:actor}];
      const closingNow=changes.status==="Closed"&&t.status!=="Closed";
      const updated={...t,...changes,updatedAt:Date.now(),timeline:tl};
      if(closingNow){
        updated.closedAt=Date.now();
        updated.closingRemarks=remark||"Closed from ticket controls.";
        updated.resolutionTime=updated.closedAt-t.createdAt;
        emailTicketClosed(updated,STAFF_BASE.find(s=>s.id===updated.assigneeId),actor);
      }
      if(changes.assigneeId&&Number(changes.assigneeId)!==Number(t.assigneeId)){
        emailTicketAssigned(updated,STAFF_BASE.find(s=>s.id===Number(changes.assigneeId)),STAFF_BASE.find(s=>s.id===t.assigneeId),actor,remark);
      }
      return updated;
    }));
    toast(auditAction,"success");
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

  const handleCloseTicket=(remarks)=>{
    const closedAt=Date.now();
    setTickets(ts=>ts.map(t=>{
      if(t.id!==ticketId) return t;
      const closedBy = staffName || (isAdmin ? "Admin" : "User");
      const updated={...t,status:"Closed",closedAt,closingRemarks:remarks,resolutionTime:closedAt-t.createdAt,updatedAt:closedAt,
        timeline:[...(t.timeline||[]),{action:"Closed",remark:remarks,at:closedAt,by:closedBy}]};
      emailTicketClosed(updated,STAFF_BASE.find(s=>s.id===t.assigneeId),closedBy);
      return updated;
    }));
    setShowClose(false);
    toast("Ticket closed! Email notification sent 📧","success");
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
        <div style={{width:44,height:44,borderRadius:"50%",background:assignee.color+"33",border:`2px solid ${assignee.color}`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:13,fontWeight:700,color:assignee.color,flexShrink:0}}>{assignee.avatar}</div>
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
              {s.name} ({s.role})
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
    const assignee=STAFF_BASE[Math.floor(Math.random()*STAFF_BASE.length)];
    const ticket={...form,id:genId(),status:"Assigned",assigneeId:assignee.id,createdAt:Date.now(),updatedAt:Date.now(),comments:[],
      timeline:[{action:"Created",at:Date.now(),by:"User"},{action:`Assigned to ${assignee.name}`,at:Date.now(),by:"System (Auto)"}]};
    emailTicketCreated(ticket,assignee);
    onSubmit(ticket);
    setLoading(false);
    toast(`Ticket created! Confirmation sent to ${ticket.email} 📧`,"email");
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

  const doExport=async(type)=>{
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

  const slaData=filtered.map(t=>({
    id:t.id,priority:t.priority,slaHours:SLA_HOURS[t.priority],
    status:t.status,elapsed:t.closedAt?formatDuration(t.closedAt-t.createdAt):"Active",met:t.closedAt?(t.closedAt-t.createdAt<SLA_HOURS[t.priority]*3600000?"✅ Met":"❌ Breached"):"—"
  }));

  const staffData=STAFF_BASE.map(s=>({
    name:s.name,role:s.role,email:s.email,
    assigned:filtered.filter(t=>t.assigneeId===s.id).length,
    resolved:filtered.filter(t=>t.assigneeId===s.id&&(t.status==="Resolved"||t.status==="Closed")).length,
  }));

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
              if(ex.id==="sla") {
  downloadExcel(
    slaData,
    `sla_report_${new Date().toISOString().slice(0,10)}.xlsx`
  );

  toast("SLA report exported ✅","success");
  return;
}

if(ex.id==="staff") {
  downloadExcel(
    staffData,
    `staff_performance_${new Date().toISOString().slice(0,10)}.xlsx`
  );

  toast("Staff performance exported ✅","success");
  return;
}

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

// ── SIDEBAR ───────────────────────────────────────────────────────────────
function Sidebar({current,onChange,isAdmin,isStaff,tickets,mobileOpen,setMobileOpen}) {
  const adminNav=[{id:"dashboard",icon:"🏠",label:"Dashboard"},{id:"tickets",icon:"🎫",label:"All Tickets"},{id:"staff",icon:"👥",label:"IT Staff"},{id:"analytics",icon:"📊",label:"Analytics"},{id:"export",icon:"⬇",label:"Export Reports"},{id:"staff-management",icon:"👥",label:"Staff Management"},{id:"emaillog",icon:"📧",label:"Email Log"}];
  const userNav=[{id:"home",icon:"🏠",label:"Home"},{id:"my-tickets",icon:"🎫",label:"My Tickets"},{id:"new-ticket",icon:"➕",label:"New Ticket"},{id:"track",icon:"🔍",label:"Track Ticket"}];
  const staffNav=[{id:"staff-dash",icon:"🏠",label:"My Dashboard"},{id:"assigned",icon:"📋",label:"Assigned Tickets"}];
  const nav=isAdmin?adminNav:isStaff?staffNav:userNav;
  const open=tickets.filter(t=>t.status==="Open").length;

  return (
    <>
      {mobileOpen&&<div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.5)",zIndex:49}} onClick={()=>setMobileOpen(false)}/>}
      <aside style={{
        width:240,flexShrink:0,background:"rgba(10,10,20,0.95)",backdropFilter:"blur(20px)",
        borderRight:"1px solid rgba(255,255,255,0.07)",display:"flex",flexDirection:"column",
        height:"100vh",position:"sticky",top:0,overflowY:"auto",
      }}>
        <div style={{padding:"24px 20px 16px",borderBottom:"1px solid rgba(255,255,255,0.06)"}}>
          <div style={{display:"flex",alignItems:"center",gap:10}}>
            <div style={{width:36,height:36,borderRadius:10,background:"linear-gradient(135deg,#6366f1,#8b5cf6)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:18}}>🛡️</div>
            <div><div style={{fontFamily:"Syne",fontWeight:800,fontSize:14,color:"#e2e8f0",letterSpacing:".5px"}}>JAIPURIA</div><div style={{fontSize:10,color:"rgba(226,232,240,0.4)",letterSpacing:"1px"}}>IT HELPDESK v2</div></div>
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
function Landing({onLogin}) {
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
    <div style={{minHeight:"100vh",background:"radial-gradient(circle at 18% 18%,rgba(99,102,241,0.22),transparent 32%),radial-gradient(circle at 86% 30%,rgba(20,184,166,0.14),transparent 30%),#0a0a0f",display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <ForgotPassword onBack={()=>setShowForgot(false)} toast={toast}/>
      <Toast toasts={toasts} remove={remove}/>
    </div>
  );

  const modeCopy = {
    user:{title:"User Access",subtitle:"Raise and track IT support tickets with your institutional email.",placeholder:"name@jaipuria.ac.in"},
    staff:{title:"IT Staff Access",subtitle:"Manage assigned tickets, updates, SLA progress, and closures.",placeholder:"staff@jaipuria.ac.in"},
    admin:{title:"Admin Portal",subtitle:"Full dashboard access for tickets, analytics, exports, and staff management.",placeholder:"Admin or admin@jaipuria.ac.in"},
  }[mode];

  return (
    <div style={{minHeight:"100vh",background:"radial-gradient(circle at 16% 20%,rgba(99,102,241,0.24),transparent 34%),radial-gradient(circle at 84% 28%,rgba(16,185,129,0.16),transparent 30%),radial-gradient(circle at 50% 100%,rgba(14,165,233,0.10),transparent 38%),#0a0a0f",display:"flex",alignItems:"center",justifyContent:"center",padding:"24px 16px"}}>
      <div style={{width:"100%",maxWidth:980,display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(min(100%,360px),1fr))",gap:28,alignItems:"center"}} className="fade-up">
        <div style={{padding:"10px 4px"}}>
          <div style={{display:"inline-flex",alignItems:"center",gap:12,background:"rgba(255,255,255,0.06)",border:"1px solid rgba(255,255,255,0.1)",borderRadius:16,padding:"10px 14px",marginBottom:22}}>
            <div style={{width:42,height:42,borderRadius:12,background:"linear-gradient(135deg,#6366f1,#14b8a6)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:24}}>🛡️</div>
            <div>
              <div style={{fontFamily:"Syne",fontWeight:800,fontSize:13,color:"#fff",letterSpacing:0}}>JAIPURIA</div>
              <div style={{fontSize:11,color:"rgba(226,232,240,0.52)"}}>Institute IT Helpdesk</div>
            </div>
          </div>
          <h1 style={{fontFamily:"Syne",fontSize:"clamp(30px,5vw,54px)",lineHeight:1.02,fontWeight:800,color:"#f8fafc",letterSpacing:0,marginBottom:16}}>IT Support Portal</h1>
          <p style={{fontSize:16,lineHeight:1.7,color:"rgba(226,232,240,0.66)",maxWidth:560,marginBottom:24}}>Fast ticket logging, transparent assignment, SLA tracking, analytics, and export-ready reports for Jaipuria Institute of Management.</p>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))",gap:12,maxWidth:560}}>
            {[["24/7","Support Desk"],["SLA","Live Tracking"],["PDF/XLSX","Reports"],["Secure","Role Access"]].map(([value,label])=>(
              <div key={label} className="glass" style={{padding:"16px 14px"}}>
                <div style={{fontFamily:"Syne",fontWeight:800,fontSize:22,color:"#818cf8"}}>{value}</div>
                <div style={{fontSize:12,color:"rgba(226,232,240,0.5)",marginTop:3}}>{label}</div>
              </div>
            ))}
          </div>
        </div>

        <form onSubmit={e=>{e.preventDefault();handleLogin();}} className="glass" style={{padding:"28px 24px",boxShadow:"0 24px 80px rgba(0,0,0,0.35)",borderRadius:20}}>
          <div style={{marginBottom:20}}>
            <div style={{fontSize:12,color:"#818cf8",fontWeight:700,textTransform:"uppercase",letterSpacing:1,marginBottom:6}}>Welcome back</div>
            <h2 style={{fontFamily:"Syne",fontSize:24,fontWeight:800,color:"#f8fafc",letterSpacing:0}}>{modeCopy.title}</h2>
            <p style={{fontSize:13,lineHeight:1.6,color:"rgba(226,232,240,0.52)",marginTop:6}}>{modeCopy.subtitle}</p>
          </div>

          <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",background:"rgba(255,255,255,0.05)",border:"1px solid rgba(255,255,255,0.08)",borderRadius:14,padding:5,marginBottom:20,gap:4}}>
            {[["user","User"],["staff","IT Staff"],["admin","Admin"]].map(([m,l])=>(
              <button type="button" key={m} onClick={()=>resetMode(m)} style={{padding:"10px 8px",borderRadius:10,border:"none",fontSize:13,fontWeight:700,background:mode===m?"linear-gradient(135deg,rgba(99,102,241,0.95),rgba(20,184,166,0.82))":"transparent",color:mode===m?"#fff":"rgba(226,232,240,0.58)",boxShadow:mode===m?"0 10px 28px rgba(99,102,241,0.25)":"none"}}>{l}</button>
            ))}
          </div>

          {mode==="admin"&&(
            <div style={{display:"flex",flexDirection:"column",gap:14,marginBottom:18}}>
              <div><label style={{fontSize:12,color:"rgba(226,232,240,0.62)",marginBottom:6,display:"block",fontWeight:600}}>Username or Admin Email</label>
                <input placeholder={modeCopy.placeholder} value={adminUser} autoComplete="username" onChange={e=>setAdminUser(e.target.value)}/></div>
              <div><label style={{fontSize:12,color:"rgba(226,232,240,0.62)",marginBottom:6,display:"block",fontWeight:600}}>Password</label>
                <PwdInput value={pwd} onChange={setPwd} autoComplete="current-password" placeholder="Enter admin password"/></div>
            </div>
          )}

          {mode==="staff"&&(
            <div style={{display:"flex",flexDirection:"column",gap:14,marginBottom:18}}>
              <div><label style={{fontSize:12,color:"rgba(226,232,240,0.62)",marginBottom:6,display:"block",fontWeight:600}}>Staff Email</label>
                <input type="email" autoComplete="username" placeholder={modeCopy.placeholder} value={email} onChange={e=>setEmail(e.target.value)}/></div>
              <div><label style={{fontSize:12,color:"rgba(226,232,240,0.62)",marginBottom:6,display:"block",fontWeight:600}}>Password</label>
                <PwdInput value={pwd} onChange={setPwd} autoComplete="current-password" placeholder="Enter your password"/></div>
            </div>
          )}

          {mode==="user"&&(
            <div style={{marginBottom:18}}>
              <label style={{fontSize:12,color:"rgba(226,232,240,0.62)",marginBottom:6,display:"block",fontWeight:600}}>Institutional Email Address</label>
              <input type="email" name="login-email" autoComplete="username" placeholder={modeCopy.placeholder} value={email} onChange={e=>setEmail(e.target.value)}/>
              <div style={{fontSize:11,color:"rgba(226,232,240,0.36)",marginTop:7}}>Only @jaipuria.ac.in email IDs are allowed.</div>
            </div>
          )}

          <button type="submit" className="glow-btn" style={{width:"100%",padding:"13px",boxShadow:"0 14px 36px rgba(99,102,241,0.28)"}} disabled={loading}>
            {loading?"Authenticating...":mode==="admin"?"Access Admin Portal":"Continue"}
          </button>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",gap:12,marginTop:14,flexWrap:"wrap"}}>
            <button type="button" onClick={()=>setShowForgot(true)} style={{background:"none",border:"none",color:"rgba(129,140,248,0.9)",fontSize:13,textDecoration:"underline"}}>Forgot Password?</button>
            {mode==="admin"&&<div style={{fontSize:12,color:"rgba(226,232,240,0.34)"}}>Default: Admin / Admin@123</div>}
          </div>
        </form>
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
function StaffPanel({staffId,tickets,setTickets,toast,onViewTicket,permissions}) {
  const staff=STAFF_BASE.find(s=>s.id===staffId);
  const myTickets=tickets.filter(t=>t.assigneeId===staffId);
  const active=myTickets.filter(t=>!["Resolved","Closed"].includes(t.status)).length;
  const resolved=myTickets.filter(t=>t.status==="Resolved"||t.status==="Closed").length;
  return (
    <div style={{display:"flex",flexDirection:"column",gap:20}}>
      <div className="glass" style={{padding:"24px",display:"flex",gap:16,alignItems:"center",flexWrap:"wrap"}}>
        <div style={{width:56,height:56,borderRadius:"50%",background:staff.color+"33",border:`3px solid ${staff.color}`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:18,fontWeight:800,color:staff.color}}>{staff.avatar}</div>
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

// ── QUICK ASSIGN DIALOG ───────────────────────────────────────────────────
function QuickAssignDialog({ticket,onClose,onSave}) {
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
            <option key={staff.id} value={staff.id}>{staff.name} ({staff.role})</option>
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
  const [tickets, setTickets] = useState(() => DB.get("tickets", []));
  const [viewTicketId, setViewTicketId] = useState(null);
  const [quickAssignTicketId, setQuickAssignTicketId] = useState(null);
  const [formCat, setFormCat] = useState(null);
  const [mobileOpen, setMobileOpen] = useState(false);

  useEffect(() => {
    DB.set("tickets", tickets);
  }, [tickets]);

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

  const handleDeleteTicket = (id) => {
    setTickets(ts => ts.filter(t => t.id !== id));
    toast("Ticket deleted", "info");
  };

  const handleNewTicket = async (ticket) => {
    setTickets(ts => [ticket, ...ts]);
    setFormCat(null);

    const assignee = STAFF_BASE.find(s => s.id === ticket.assigneeId);
    await sendTicketEmail(ticket, {
      name: ticket.name,
      email: ticket.email,
    });

    if (assignee) {
      simulateEmail(
        assignee.email,
        `[${ticket.id}] New Ticket Assigned`,
        `${ticket.name} submitted a ${CATEGORIES.find(c => c.id === ticket.category)?.label || ticket.category} ticket.\n\n${ticket.description}`
      );
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
        <Landing onLogin={handleLogin} />
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

  const renderStaffManagement = () => (
    <div style={{display:"flex",flexDirection:"column",gap:18}}>
      <h2 style={{fontFamily:"Syne",fontSize:22,fontWeight:700,color:"#e2e8f0"}}>Staff Management</h2>
      {STAFF_BASE.map(staff => (
        <div key={staff.id} className="glass" style={{padding:"18px 20px"}}>
          <div style={{fontSize:18,fontWeight:700,color:"#fff"}}>{staff.name}</div>
          <div style={{fontSize:13,color:"rgba(226,232,240,0.5)",marginTop:4}}>{staff.email}</div>
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
                    <div style={{width:50,height:50,borderRadius:"50%",background:s.color+"33",border:`3px solid ${s.color}`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:15,fontWeight:800,color:s.color}}>{s.avatar}</div>
                    <div><div style={{fontSize:15,fontWeight:700,color:"#e2e8f0"}}>{s.name}</div><div style={{fontSize:12,color:"rgba(226,232,240,0.5)"}}>{s.role}</div><div style={{fontSize:11,color:"rgba(226,232,240,0.35)"}}>{s.email}</div></div>
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
      return <StaffPanel staffId={session.staffId} tickets={tickets} setTickets={setTickets} toast={toast} onViewTicket={setViewTicketId} permissions={session.permissions} />;
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
    if (page === "track") return <TrackTicket tickets={tickets} onView={setViewTicketId} />;
    if (page === "new-ticket") return (
      <div>
        <h2 style={{fontFamily:"Syne",fontSize:22,fontWeight:700,color:"#e2e8f0",marginBottom:20}}>New Ticket</h2>
        <div className="glass" style={{padding:"24px"}}>
          <TicketForm userEmail={session.email} initialCategory="" onSubmit={t => { handleNewTicket(t); setPage("my-tickets"); }} onCancel={() => setPage("home")} toast={toast} />
        </div>
      </div>
    );

    return null;
  };

  return (
    <>
      <style>{CSS}</style>
      <div style={{display:"flex",minHeight:"100vh"}}>
        <Sidebar current={page} onChange={setPage} isAdmin={isAdmin} isStaff={isStaff} tickets={tickets} mobileOpen={mobileOpen} setMobileOpen={setMobileOpen} />
        <div style={{flex:1,display:"flex",flexDirection:"column",minWidth:0}}>
          <div style={{padding:"14px 24px",borderBottom:"1px solid rgba(255,255,255,0.07)",display:"flex",justifyContent:"space-between",alignItems:"center",background:"rgba(10,10,20,0.9)",backdropFilter:"blur(20px)",position:"sticky",top:0,zIndex:10}}>
            <div style={{display:"flex",alignItems:"center",gap:12}}>
              <button onClick={() => setMobileOpen(o => !o)} style={{background:"rgba(255,255,255,0.06)",border:"1px solid rgba(255,255,255,0.1)",color:"#e2e8f0",width:36,height:36,borderRadius:8,fontSize:18}}>☰</button>
              <span style={{fontSize:13,color:"rgba(226,232,240,0.4)"}}>{isAdmin ? "Admin Portal" : isStaff ? `${session.name} (${session.role})` : session.email}</span>
            </div>
            <div style={{display:"flex",gap:10,alignItems:"center"}}>
              <div className="pulse" style={{width:8,height:8,borderRadius:"50%",background:"#10b981"}} />
              <span style={{fontSize:12,color:"rgba(226,232,240,0.4)"}}>Live</span>
              <button onClick={logoutUser} style={{background:"rgba(239,68,68,0.1)",border:"1px solid rgba(239,68,68,0.3)",color:"#f87171",padding:"6px 14px",borderRadius:8,fontSize:13}}>Logout</button>
            </div>
          </div>
          <div style={{padding:"24px 28px",flex:1,overflowY:"auto"}}>{renderPage()}</div>
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
          />
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
          />
        </Modal>
      )}

      <Toast toasts={toasts} remove={remove} />
    </>
  );
}
















