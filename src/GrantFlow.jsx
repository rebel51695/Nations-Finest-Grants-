import { useState, useEffect, useMemo, useRef, Fragment } from "react";
import { doc, getDoc } from "firebase/firestore";
import { db } from "./firebaseConfig";
import * as XLSX from "xlsx";
import {
  LayoutDashboard, FileText, Wallet, BarChart3, Plus, X, Pencil, Trash2,
  ExternalLink, Download, Search, ArrowRight, AlertCircle, CheckCircle2,
  ClipboardList, Circle, CheckCircle, Users, PieChart, TrendingUp, History, CheckSquare, Upload, Printer, RefreshCw, Receipt, Menu, Shield, FlaskConical,
} from "lucide-react";
import AdminPanel from "./AdminPanel.jsx";
import {
  BarChart, Bar, LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip,
  ResponsiveContainer, Legend,
} from "recharts";

// ---------- constants ----------

const MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

// Given a budget's Period Start date, returns 12 entries describing which real
// calendar month/year each of the budget's 12 monthly slots actually falls on.
// This lets a budget's columns be labeled correctly (e.g. "Oct 2025" instead of
// always assuming "Jan") for grants that don't run on the calendar year.
function monthColumnsForBudget(periodStart) {
  let startYear, startMonth; // startMonth is 0-indexed
  if (periodStart) {
    const d = new Date(periodStart + "T00:00:00");
    if (!isNaN(d)) { startYear = d.getFullYear(); startMonth = d.getMonth(); }
  }
  if (startYear === undefined) { startYear = new Date().getFullYear(); startMonth = 0; }
  return Array.from({ length: 12 }, (_, i) => {
    const totalMonth = startMonth + i;
    const year = startYear + Math.floor(totalMonth / 12);
    const monthIndex = ((totalMonth % 12) + 12) % 12;
    return { year, monthIndex, label: `${MONTHS[monthIndex]} ${String(year).slice(2)}` };
  });
}

const CATEGORIES = [
  { name: "Grants and Contracts", type: "revenue", subs: ["4100 - Grants and Contracts"] },
  { name: "Grants and Contracts Indirect Billing", type: "revenue", subs: ["4101 - Grants and Contracts Indirect Billing"] },
  { name: "Donations Without In Kind", type: "revenue", subs: ["4000 - Donation"] },
  { name: "Other Non-SSVF Revenue", type: "revenue", subs: [
    "4004 - Event Revenue", "4020 - Interest/Dividends", "4025 - F/B transfer in", "4060 - WEG Misc. Income",
    "4120 - Program Fees Revenue", "4601 - Gain on Sale", "4625 - Forgiveness of Debt", "4990 - MISC", "4999 - Insurance Claim",
  ] },
  { name: "Wages and Benefits", type: "expense", subs: ["5000 - Salary and Wages", "5900 - Payroll taxes and benefits"] },
  { name: "Operations", type: "expense", subs: [
    "6000 - Outreach/Recruitment", "6001 - Marketing", "6005 - Employee Expenses", "6015 - Dues and Membership",
    "6016 - Donations", "6018 - Grant Expense", "6020 - Equipment Expenses", "6025 - Vehicle Expenses",
    "6027 - F/B tsfr out", "6040 - Facility Expenses", "6045 - Fees/Licenses", "6047 - Software & Licensing",
    "6050 - Office Supplies", "6055 - Postage/Freight", "6060 - Printing/Duplication",
    "6071 - Homeless Management Information System (HMIS)", "6085 - Non Federal Grant Client Expenses",
    "6120 - Consulting Expenses", "6127 - CARF Certification", "6130 - G & S - Expenses", "6135 - Insurance",
    "6140 - Interest", "6145 - Ret. Forfeiture", "6150 - Spec. Projects", "6155 - Taxes", "6175 - Miscellaneous",
    "6185 - Insurance Claim Expense", "6190 - Refunds/NSF Checks", "6300 - Bad Debt", "6700 - Depreciation",
    "6750 - Intercompany Billing", "6989 - Other Total Budget", "6990 - Indirect", "6998 - REC F.A./EQ. Grantor",
    "6999 - Asses Sale Gain/Loss",
  ] },
  { name: "TFA", type: "expense", subs: ["7100 - Category 1 Temporary Financial Assistance", "7200 - Category 2 Temporary Financial Assistance"] },
];
const CUSTOM_CATEGORY = "__custom__";

const STAGES = ["Prospecting", "Writing", "Applied", "Awarded", "Rejected", "Active", "Closing", "Closed"];
const SITE_OPTIONS = [
  "Carson City", "Chico", "Flagstaff", "Mather", "Menlo Park", "Monterey", "Prescott",
  "Redding", "Reno", "Sacramento", "Santa Cruz", "Vacaville", "Bullhead City", "Eureka",
  "Santa Rosa", "All CA", "All AZ", "All NV", "Residential", "MSU", "Corp",
];
const RISKS = ["Low", "Medium", "High"];
const CADENCES = ["Weekly", "Monthly", "Quarterly", "Semi-annual", "Annually", "End of grant"];
const BUDGET_STATUSES = ["Draft", "Pending Approval", "Active", "Rejected", "Closed"];
const STAFF_STATUSES = ["Active", "Inactive", "Leave of Absence"];
const AUTO_BACKUP_RETENTION_DAYS = 30;
const INVOICE_STATUSES = ["Draft", "Submitted", "Paid", "Rejected"];
const PAYMENT_METHODS = ["Billable Service", "Interval Lump Sum", "Lump Sum", "Per Diem Rate", "Reimbursement"];

const REPORT_STATUSES = ["Not started", "In progress", "Completed"];
const REPORT_PRIORITIES = [
  { label: "Urgent", color: "#B5443A" },
  { label: "Important", color: "#C08A2E" },
  { label: "Medium", color: "#2F6F53" },
  { label: "Low", color: "#5B7FA6" },
];
const REPORT_REPEATS = ["None", "Weekly", "Monthly", "Quarterly", "Annually"];
const DEFAULT_BUCKETS = ["Backlog", "Upcoming", "Up next", "In progress", "Complete", "Submitted"];
const TASK_STATUSES = ["Not started", "In progress", "Done"];
const TASK_CATEGORIES = ["Application/Submission", "Site Visit", "Renewal Prep", "Document Collection", "Board Approval", "Compliance", "Personnel Reallocation", "Report Submission", "Other"];

const APP_VERSION = "1.1.2";
const uid = () => Math.random().toString(36).slice(2, 10);
const stripNonce = (v) => (v ? v.split("::")[0] : "");
const fmt = (n) => (Number(n) || 0).toLocaleString("en-US", { style: "currency", currency: "USD", maximumFractionDigits: 0 });
const fmtDate = (d) => (d ? new Date(d + "T00:00:00").toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" }) : "—");

function newLine() {
  const c = CATEGORIES[0];
  return {
    id: uid(), category: c.name, type: c.type, categoryCustom: false,
    subcategory: "", subcategoryCustom: false, description: "",
    amounts: Array(12).fill(0), actuals: Array(12).fill(0),
  };
}

function lineTotal(line) {
  return line.amounts.reduce((a, b) => a + (Number(b) || 0), 0);
}

function lineActualTotal(line) {
  return (line.actuals || []).reduce((a, b) => a + (Number(b) || 0), 0);
}

function distributeEvenly(total) {
  const cents = Math.round((Number(total) || 0) * 100);
  const base = Math.floor(cents / 12);
  const remainder = cents - base * 12;
  return Array.from({ length: 12 }, (_, i) => (base + (i < remainder ? 1 : 0)) / 100);
}

function budgetTotals(budget) {
  let revenue = 0, expense = 0;
  const monthly = Array(12).fill(0);
  for (const line of budget.lines) {
    const t = lineTotal(line);
    if (line.type === "revenue") revenue += t; else expense += t;
    line.amounts.forEach((a, i) => { monthly[i] += (line.type === "revenue" ? 1 : -1) * (Number(a) || 0); });
  }
  return { revenue, expense, net: revenue - expense, monthly };
}

function budgetActualTotals(budget) {
  let revenue = 0, expense = 0;
  const monthly = Array(12).fill(0);
  for (const line of budget.lines) {
    const t = lineActualTotal(line);
    if (line.type === "revenue") revenue += t; else expense += t;
    (line.actuals || []).forEach((a, i) => { monthly[i] += (line.type === "revenue" ? 1 : -1) * (Number(a) || 0); });
  }
  return { revenue, expense, net: revenue - expense, monthly };
}

function grantBudgetTotals(grantId, budgets) {
  const mine = budgets.filter((b) => b.grantId === grantId);
  return mine.reduce((acc, b) => {
    const t = budgetTotals(b);
    acc.revenue += t.revenue; acc.expense += t.expense;
    return acc;
  }, { revenue: 0, expense: 0 });
}

function expenseMonthlyArray(budget) {
  const arr = Array(12).fill(0);
  budget.lines.forEach((l) => { if (l.type === "expense") l.amounts.forEach((a, i) => { arr[i] += Number(a) || 0; }); });
  return arr;
}

function budgetElapsedMonths(budget) {
  if (!budget.periodStart) return null;
  const start = new Date(budget.periodStart + "T00:00:00");
  const today = new Date();
  const months = (today.getFullYear() - start.getFullYear()) * 12 + (today.getMonth() - start.getMonth()) + 1;
  return Math.max(0, Math.min(months, 12));
}

function expenseActualMonthlyArray(budget) {
  const arr = Array(12).fill(0);
  budget.lines.forEach((l) => { if (l.type === "expense") (l.actuals || []).forEach((a, i) => { arr[i] += Number(a) || 0; }); });
  return arr;
}

function budgetBurnInfo(budget) {
  const expenseMonthly = expenseMonthlyArray(budget);
  const totalExpense = expenseMonthly.reduce((a, b) => a + b, 0);
  const actualExpenseMonthly = expenseActualMonthlyArray(budget);
  const actualToDate = actualExpenseMonthly.reduce((a, b) => a + b, 0);
  const elapsed = budgetElapsedMonths(budget);
  if (elapsed === null) return { totalExpense, toDate: null, elapsed: null, actualToDate };
  const toDate = expenseMonthly.slice(0, elapsed).reduce((a, b) => a + b, 0);
  return { totalExpense, toDate, elapsed, actualToDate };
}

function grantBurn(grant, budgets) {
  const mine = budgets.filter((b) => b.grantId === grant.id);
  let totalExpense = 0, toDate = 0, actualToDate = 0, maxElapsed = 0, elapsedKnown = false;
  mine.forEach((b) => {
    const info = budgetBurnInfo(b);
    totalExpense += info.totalExpense;
    actualToDate += info.actualToDate;
    if (info.elapsed !== null) {
      elapsedKnown = true;
      toDate += info.toDate;
      maxElapsed = Math.max(maxElapsed, info.elapsed);
    }
  });
  const award = Number(grant.awardAmount) || 0;
  const hasActuals = actualToDate > 0;
  const spendToDate = hasActuals ? actualToDate : toDate;
  const pctTimeElapsed = elapsedKnown ? maxElapsed / 12 : null;
  const pctBudgetUsed = totalExpense > 0 ? spendToDate / totalExpense : 0;
  const monthlyAvg = elapsedKnown && maxElapsed > 0 ? spendToDate / maxElapsed : totalExpense / 12;
  const projectedFullYear = monthlyAvg * 12;
  const variance = actualToDate - toDate;
  let status = "No budget period set";
  if (elapsedKnown) {
    if (maxElapsed === 0) status = "Not started";
    else if (pctBudgetUsed > pctTimeElapsed + 0.07) status = "Ahead of pace";
    else if (pctBudgetUsed < pctTimeElapsed - 0.07) status = "Behind pace";
    else status = "On pace";
  }
  return {
    totalExpense, toDate, actualToDate, hasActuals, variance, pctTimeElapsed, pctBudgetUsed, monthlyAvg, projectedFullYear,
    status, projectedOverAward: award > 0 && projectedFullYear > award, award, elapsedMonths: maxElapsed, elapsedKnown,
  };
}

const riskColor = { Low: "#2F6F53", Medium: "#C08A2E", High: "#B5443A" };
const stageColor = {
  Prospecting: "#8A8F87", Writing: "#5B7FA6", Applied: "#5B7FA6", Awarded: "#A8791F", Rejected: "#B5443A",
  Active: "#2F6F53", Closing: "#C08A2E", Closed: "#8A8F87",
};
const ANNUAL_HOURS = 1768;

function newAllocation() {
  return { id: uid(), type: "grant", grantId: "", costCenterId: "", percent: 0 };
}

function staffAnnualCost(staff) {
  const fte = Number(staff.fte) || 0;
  if (staff.payType === "Hourly") {
    return (Number(staff.hourlyRate) || 0) * (Number(staff.annualHours) || ANNUAL_HOURS) * fte;
  }
  return (Number(staff.annualSalary) || 0) * fte;
}

function staffAllocatedTotal(staff) {
  return (staff.allocations || []).reduce((a, al) => a + (Number(al.percent) || 0), 0);
}

function personnelCostByGrant(staffList) {
  const map = {};
  staffList.forEach((s) => {
    const cost = staffAnnualCost(s);
    (s.allocations || []).forEach((al) => {
      if (!al.grantId) return;
      map[al.grantId] = (map[al.grantId] || 0) + cost * ((Number(al.percent) || 0) / 100);
    });
  });
  return map;
}

function personnelCostByCostCenter(staffList) {
  const map = {};
  staffList.forEach((s) => {
    const cost = staffAnnualCost(s);
    (s.allocations || []).forEach((al) => {
      if (!al.costCenterId) return;
      map[al.costCenterId] = (map[al.costCenterId] || 0) + cost * ((Number(al.percent) || 0) / 100);
    });
  });
  return map;
}

function pushTrash(setTrash, entityType, data, deletedBy, extra) {
  setTrash((prev) => [
    { id: uid(), entityType, data, extra: extra || null, deletedAt: new Date().toISOString(), deletedBy: deletedBy || "Unknown" },
    ...prev,
  ].slice(0, 500));
}

function newScenario(basedOn) {
  return {
    id: uid(), title: "", notes: "", createdBy: "", createdAt: new Date().toISOString(),
    basedOn: basedOn || { type: "blank" },
    fy: "", periodStart: "", periodEnd: "",
    lines: [newLine()],
  };
}

// Computes the CURRENT real numbers a scenario should be compared against,
// based on what it was snapshotted from. Always re-derived live — never a
// frozen copy — so the comparison reflects reality as of right now.
function liveComparisonForScenario(scenario, grants, budgets, costCenters) {
  const basedOn = scenario.basedOn || { type: "blank" };
  if (basedOn.type === "grant" || basedOn.type === "costCenter") {
    const b = budgets.find((x) => x.id === basedOn.budgetId);
    if (!b) return { available: false, reason: "The original budget this was based on no longer exists." };
    const map = {};
    b.lines.forEach((l) => {
      if (!map[l.category]) map[l.category] = Array(12).fill(0);
      (l.amounts || Array(12).fill(0)).forEach((a, i) => { map[l.category][i] += Number(a) || 0; });
    });
    return { available: true, byCategory: map, periodStart: b.periodStart };
  }
  if (basedOn.type === "org") {
    const scope = basedOn.scope || "all";
    const calYear = basedOn.calYear ?? "All";
    const scopedGrantIds = scope === "all" ? null : new Set(grants.filter((g) => g.budgetGroupId === scope).map((g) => g.id));
    const scopedCcIds = scope === "all" ? null : new Set((costCenters || []).filter((c) => c.budgetGroupId === scope).map((c) => c.id));
    const scopedBudgets = (scope === "all" ? budgets : budgets.filter((b) => (b.grantId && scopedGrantIds.has(b.grantId)) || (b.costCenterId && scopedCcIds.has(b.costCenterId)))).filter((b) => b.status === "Active");
    const map = {};
    scopedBudgets.forEach((b) => {
      const cols = monthColumnsForBudget(b.periodStart);
      b.lines.forEach((l) => {
        if (!map[l.category]) map[l.category] = Array(12).fill(0);
        (l.amounts || Array(12).fill(0)).forEach((a, i) => {
          const col = cols[i];
          if (calYear !== "All" && col.year !== calYear) return;
          map[l.category][col.monthIndex] += Number(a) || 0;
        });
      });
    });
    return { available: true, byCategory: map, periodStart: "" };
  }
  return { available: false, reason: "This scenario started blank, with nothing to compare against." };
}

function daysOutstanding(inv) {
  if (!inv.submittedDate || inv.paidDate) return null;
  const diff = Date.now() - new Date(inv.submittedDate + "T00:00:00").getTime();
  return Math.max(0, Math.round(diff / 86400000));
}

function isInvoiceOverdue(inv) {
  return inv.status === "Submitted" && inv.dueDate && !inv.paidDate && new Date(inv.dueDate) < new Date(new Date().toDateString());
}

const priorityColor = (label) => (REPORT_PRIORITIES.find((p) => p.label === label) || REPORT_PRIORITIES[2]).color;

const checklistProgress = (report) => {
  const items = report.checklist || [];
  return { done: items.filter((i) => i.done).length, total: items.length };
};
const isOverdue = (report) => report.dueDate && report.status !== "Completed" && new Date(report.dueDate) < new Date(new Date().toDateString());

// ---------- shared bits ----------

function Badge({ children, color }) {
  return (
    <span
      className="inline-flex items-center px-2 py-0.5 rounded-full text-xs font-medium"
      style={{ background: `${color}1a`, color }}
    >
      {children}
    </span>
  );
}

function StatCard({ label, value, sub }) {
  return (
    <div className="bg-white rounded-lg border p-4" style={{ borderColor: "#E1E5DE" }}>
      <div className="text-xs uppercase tracking-wide" style={{ color: "#5B6B66" }}>{label}</div>
      <div className="text-2xl mt-1 font-display" style={{ color: "#1C2624", fontVariantNumeric: "tabular-nums" }}>{value}</div>
      {sub && <div className="text-xs mt-1" style={{ color: "#8A8F87" }}>{sub}</div>}
    </div>
  );
}

function Field({ label, children }) {
  return (
    <label className="block">
      <span className="block text-xs font-medium mb-1" style={{ color: "#5B6B66" }}>{label}</span>
      {children}
    </label>
  );
}

const inputCls = "w-full rounded-md border px-3 py-2 text-sm outline-none focus:ring-2";
const inputStyle = { borderColor: "#E1E5DE", color: "#1C2624" };

function grantLabel(g) {
  return g.programCode ? `${g.programCode} - ${g.title}` : g.title;
}

async function loadData(baseKey) {
  // Plain single-key format (the normal, simple case)
  try {
    const plain = await window.storage.get(baseKey, true);
    if (plain?.value) return JSON.parse(plain.value);
  } catch (e) { /* not stored this way, try the chunked fallback below */ }
  // Fallback: in case data was ever written in the old chunked format
  try {
    const countRes = await window.storage.get(`${baseKey}:count`, true);
    if (countRes?.value) {
      const count = parseInt(countRes.value, 10) || 0;
      let all = [];
      for (let i = 0; i < count; i++) {
        const part = await window.storage.get(`${baseKey}:${i}`, true);
        if (part?.value) all = all.concat(JSON.parse(part.value));
      }
      return all;
    }
  } catch (e) { /* nothing stored yet */ }
  return null;
}

async function saveData(baseKey, value) {
  await window.storage.set(baseKey, JSON.stringify(value), true);
}

function downloadFile(filename, content, mime) {
  try {
    const blob = new Blob([content], { type: mime });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 2000);
    return true;
  } catch (err) {
    return false;
  }
}

function printSection(elementId, title) {
  const el = document.getElementById(elementId);
  if (!el) return;
  const win = window.open("", "_blank");
  if (!win) {
    alert("Your browser blocked the print window. Please allow pop-ups for this site and try again.");
    return;
  }
  win.document.write(`
    <!DOCTYPE html>
    <html>
      <head>
        <title>${title}</title>
        <style>
          body { font-family: Inter, system-ui, sans-serif; color: #1C2624; padding: 24px; }
          h1, h2 { font-family: Oswald, sans-serif; text-transform: uppercase; letter-spacing: 0.02em; }
          table { border-collapse: collapse; width: 100%; }
          th, td { padding: 6px 10px; font-size: 12px; text-align: left; }
          th { background: #F6F7F3; }
          tr { border-top: 1px solid #E1E5DE; }
          .no-print, button, select, input { display: none !important; }
          @media print { body { padding: 0; } }
        </style>
      </head>
      <body>
        <h1>${title}</h1>
        ${el.innerHTML}
      </body>
    </html>
  `);
  win.document.close();
  win.focus();
  setTimeout(() => { win.print(); }, 400);
}

function GrantPicker({ grants, value, onChange, placeholder = "Select a grant", noneLabel, noneValue = "", wrapStyle }) {
  const [query, setQuery] = useState("");
  const [open, setOpen] = useState(false);
  const [highlighted, setHighlighted] = useState(0);
  const selected = grants.find((g) => g.id === value);
  const q = query.trim().toLowerCase();
  const filtered = q ? grants.filter((g) => grantLabel(g).toLowerCase().includes(q)) : grants;
  const rows = noneLabel ? [{ id: "__none__", label: noneLabel, isNone: true }, ...filtered] : filtered;

  const selectRow = (row) => {
    onChange(row.isNone ? noneValue : row.id);
    setQuery(""); setOpen(false); setHighlighted(0);
  };
  const onKeyDown = (e) => {
    if (!open || rows.length === 0) return;
    if (e.key === "ArrowDown") { e.preventDefault(); setHighlighted((h) => (h + 1) % rows.length); }
    else if (e.key === "ArrowUp") { e.preventDefault(); setHighlighted((h) => (h - 1 + rows.length) % rows.length); }
    else if (e.key === "Enter") { e.preventDefault(); selectRow(rows[highlighted]); }
    else if (e.key === "Escape") { setOpen(false); }
  };

  return (
    <div className="relative" style={wrapStyle}>
      <input
        value={open ? query : (selected ? grantLabel(selected) : "")}
        onChange={(e) => { setQuery(e.target.value); setOpen(true); setHighlighted(0); }}
        onFocus={() => { setQuery(""); setOpen(true); setHighlighted(0); }}
        onBlur={() => setTimeout(() => setOpen(false), 150)}
        onKeyDown={onKeyDown}
        placeholder={open ? "Type to search…" : (noneLabel && !selected ? noneLabel : placeholder)}
        className={inputCls}
        style={inputStyle}
      />
      {open && (
        <div className="absolute mt-1 w-full bg-white rounded-md border shadow-lg z-50 max-h-60 overflow-y-auto" style={{ borderColor: "#E1E5DE" }}>
          {rows.length === 0 ? (
            <div className="px-3 py-2 text-sm" style={{ color: "#8A8F87" }}>No matching grants</div>
          ) : rows.map((row, i) => (
            <button
              key={row.id}
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => selectRow(row)}
              onMouseEnter={() => setHighlighted(i)}
              className="w-full text-left px-3 py-2 text-sm border-b last:border-b-0"
              style={{
                borderColor: "#E1E5DE",
                background: highlighted === i ? "#F6F7F3" : "transparent",
                color: row.isNone ? "#5B6B66" : (value === row.id ? "#2F6F53" : "#1C2624"),
                fontWeight: !row.isNone && value === row.id ? 600 : 400,
              }}
            >
              {row.isNone ? row.label : grantLabel(row)}
            </button>
          ))}
        </div>
      )}
    </div>
  );
}

function Modal({ title, onClose, children, wide, size }) {
  const widthClass = size === "xl" ? "max-w-[1400px]" : wide ? "max-w-4xl" : "max-w-lg";
  return (
    <div className="fixed inset-0 z-50 flex items-start justify-center overflow-y-auto py-8 px-4" style={{ background: "rgba(28,38,36,0.45)" }}>
      <div className={`bg-white rounded-xl shadow-xl w-full ${widthClass} my-auto`}>
        <div className="flex items-center justify-between px-6 py-4 border-b" style={{ borderColor: "#E1E5DE" }}>
          <h2 className="font-display text-lg" style={{ color: "#1C2624" }}>{title}</h2>
          <button onClick={onClose} className="p-1 rounded hover:bg-stone-100">
            <X size={18} style={{ color: "#5B6B66" }} />
          </button>
        </div>
        <div className="p-6">{children}</div>
      </div>
    </div>
  );
}

function ConfirmModal({ message, onConfirm, onCancel }) {
  return (
    <Modal title="Confirm delete" onClose={onCancel}>
      <p className="text-sm mb-6" style={{ color: "#5B6B66" }}>{message}</p>
      <div className="flex justify-end gap-2">
        <button onClick={onCancel} className="px-4 py-2 rounded-md text-sm border" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>Cancel</button>
        <button onClick={onConfirm} className="px-4 py-2 rounded-md text-sm text-white" style={{ background: "#B5443A" }}>Delete</button>
      </div>
    </Modal>
  );
}

// ---------- grant form ----------

function BudgetGroupModal({ budgetGroup, onSave, onClose, onDelete }) {
  const [form, setForm] = useState(budgetGroup || { id: uid(), name: "", description: "" });
  const set = (k) => (e) => setForm({ ...form, [k]: e.target.value });

  return (
    <Modal title={budgetGroup ? "Edit budget group" : "New budget group"} onClose={onClose}>
      <div className="space-y-4">
        <Field label="Name">
          <input className={inputCls} style={inputStyle} value={form.name} onChange={set("name")} placeholder="e.g. Housing Programs, Veteran Support Services" autoFocus />
        </Field>
        <Field label="Description (optional)">
          <textarea className={inputCls} style={inputStyle} rows={2} value={form.description} onChange={set("description")} />
        </Field>
      </div>
      <div className="flex justify-between gap-2 mt-6">
        {onDelete ? (
          <button onClick={onDelete} className="px-4 py-2 rounded-md text-sm border" style={{ borderColor: "#E1E5DE", color: "#B5443A" }}>Delete</button>
        ) : <span />}
        <div className="flex gap-2">
          <button onClick={onClose} className="px-4 py-2 rounded-md text-sm border" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>Cancel</button>
          <button
            onClick={() => { if (!form.name.trim()) return; onSave(form); }}
            className="px-4 py-2 rounded-md text-sm text-white"
            style={{ background: "#1F5C6B" }}
          >
            Save
          </button>
        </div>
      </div>
    </Modal>
  );
}

function CostCenterModal({ costCenter, budgetGroups, setBudgetGroups, logActivity, onSave, onClose, onDelete }) {
  const [form, setForm] = useState(costCenter || { id: uid(), name: "", description: "", budgetGroupId: "" });
  const [bgModal, setBgModal] = useState(null);
  const set = (k) => (e) => setForm({ ...form, [k]: e.target.value });

  const saveBudgetGroup = (bg) => {
    setBudgetGroups((prev) => {
      const exists = prev.some((x) => x.id === bg.id);
      logActivity?.("Budget Group", exists ? "Updated" : "Created", bg.name || "Untitled budget group");
      return exists ? prev.map((x) => (x.id === bg.id ? bg : x)) : [...prev, bg];
    });
    setForm((f) => ({ ...f, budgetGroupId: bg.id }));
    setBgModal(null);
  };

  return (
    <Modal title={costCenter ? "Edit cost center" : "New cost center"} onClose={onClose}>
      <div className="space-y-4">
        <Field label="Name">
          <input className={inputCls} style={inputStyle} value={form.name} onChange={set("name")} placeholder="e.g. Administration, Fundraising, Facilities" autoFocus />
        </Field>
        <Field label="Budget group (optional)">
          <div className="flex items-center gap-2">
            <select value={form.budgetGroupId || ""} onChange={set("budgetGroupId")} className={inputCls} style={inputStyle}>
              <option value="">No group</option>
              {(budgetGroups || []).map((bg) => <option key={bg.id} value={bg.id}>{bg.name}</option>)}
            </select>
            <button onClick={() => setBgModal("new")} className="shrink-0 text-xs px-3 py-2 rounded-md border" style={{ borderColor: "#E1E5DE", color: "#1F5C6B" }}>+ New</button>
          </div>
        </Field>
        <Field label="Description (optional)">
          <textarea className={inputCls} style={inputStyle} rows={2} value={form.description} onChange={set("description")} />
        </Field>
      </div>
      <div className="flex justify-between gap-2 mt-6">
        {onDelete ? (
          <button onClick={onDelete} className="px-4 py-2 rounded-md text-sm border" style={{ borderColor: "#E1E5DE", color: "#B5443A" }}>Delete</button>
        ) : <span />}
        <div className="flex gap-2">
          <button onClick={onClose} className="px-4 py-2 rounded-md text-sm border" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>Cancel</button>
          <button
            onClick={() => { if (!form.name.trim()) return; onSave(form); }}
            className="px-4 py-2 rounded-md text-sm text-white"
            style={{ background: "#1F5C6B" }}
          >
            Save
          </button>
        </div>
      </div>
      {bgModal && <BudgetGroupModal budgetGroup={bgModal === "new" ? null : bgModal} onSave={saveBudgetGroup} onClose={() => setBgModal(null)} />}
    </Modal>
  );
}

function GrantModal({ grant, budgetGroups, setBudgetGroups, logActivity, canEdit = true, onSave, onClose }) {
  const [form, setForm] = useState(grant || {
    id: uid(), title: "", programCode: "", funding: "", sites: [], stage: "Prospecting",
    awardAmount: 0, start: "", end: "", riskStatus: "Low", cadence: [],
    complianceOwner: "", financeOwner: "", internalOwner: "", operationsOwner: "", renewal: false,
    doclibUrl: "", contractUrl: "", notes: "",
    budgetPeriodStart: "", budgetPeriodEnd: "", obligatedFunds: 0, obligatedFundsRemaining: 0, paymentMethod: PAYMENT_METHODS[0],
    beds: "", bedRate: 0, grantPoc: "", awardAmountRemaining: 0, budgetGroupId: "",
  });
  const [bgModal, setBgModal] = useState(null);
  const saveBudgetGroup = (bg) => {
    setBudgetGroups((prev) => {
      const exists = prev.some((x) => x.id === bg.id);
      logActivity?.("Budget Group", exists ? "Updated" : "Created", bg.name || "Untitled budget group");
      return exists ? prev.map((x) => (x.id === bg.id ? bg : x)) : [...prev, bg];
    });
    setForm((f) => ({ ...f, budgetGroupId: bg.id }));
    setBgModal(null);
  };
  const set = (k) => (e) => setForm({ ...form, [k]: e.target.type === "checkbox" ? e.target.checked : e.target.value });
  const toggleSite = (site) => {
    setForm((f) => ({
      ...f,
      sites: f.sites.includes(site) ? f.sites.filter((s) => s !== site) : [...f.sites, site],
    }));
  };
  const toggleCadence = (c) => {
    setForm((f) => ({
      ...f,
      cadence: f.cadence.includes(c) ? f.cadence.filter((x) => x !== c) : [...f.cadence, c],
    }));
  };

  return (
    <Modal title={grant ? (canEdit ? "Edit grant" : "View grant") : "New grant"} onClose={onClose} wide>
      <fieldset disabled={!canEdit} style={{ border: "none", margin: 0, padding: 0 }}>
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
        <Field label="Grant title">
          <input className={inputCls} style={inputStyle} value={form.title} onChange={set("title")} placeholder="e.g. SSVF Supportive Services" />
        </Field>
        <Field label="Program code">
          <input className={inputCls} style={inputStyle} value={form.programCode} onChange={set("programCode")} placeholder="e.g. SSVF-26" />
        </Field>
        <Field label="Funding source">
          <input className={inputCls} style={inputStyle} value={form.funding} onChange={set("funding")} placeholder="e.g. VA, HUD, private foundation" />
        </Field>
        <div className="col-span-2">
          <Field label={`Site / location${form.sites.length ? ` (${form.sites.length} selected)` : ""}`}>
            <div className="grid grid-cols-2 sm:grid-cols-3 gap-x-4 gap-y-1 rounded-md border p-3 max-h-48 overflow-y-auto" style={inputStyle}>
              {SITE_OPTIONS.map((site) => (
                <label key={site} className="flex items-center gap-2 text-sm" style={{ color: "#1C2624" }}>
                  <input type="checkbox" checked={form.sites.includes(site)} onChange={() => toggleSite(site)} />
                  {site}
                </label>
              ))}
            </div>
          </Field>
        </div>
        <Field label="Stage">
          <select className={inputCls} style={inputStyle} value={form.stage} onChange={set("stage")}>
            {STAGES.map((s) => <option key={s}>{s}</option>)}
          </select>
        </Field>
        <Field label="Risk status">
          <select className={inputCls} style={inputStyle} value={form.riskStatus} onChange={set("riskStatus")}>
            {RISKS.map((s) => <option key={s}>{s}</option>)}
          </select>
        </Field>
        <Field label="Award amount">
          <input type="number" className={inputCls} style={inputStyle} value={form.awardAmount} onChange={set("awardAmount")} />
        </Field>
        <Field label="Award amount remaining">
          <input type="number" className={inputCls} style={inputStyle} value={form.awardAmountRemaining} onChange={set("awardAmountRemaining")} />
        </Field>
        <div className="col-span-2">
          <Field label={`Reporting cadence${form.cadence.length ? ` (${form.cadence.length} selected)` : ""}`}>
            <div className="grid grid-cols-2 sm:grid-cols-3 gap-x-4 gap-y-1 rounded-md border p-3" style={inputStyle}>
              {CADENCES.map((c) => (
                <label key={c} className="flex items-center gap-2 text-sm" style={{ color: "#1C2624" }}>
                  <input type="checkbox" checked={form.cadence.includes(c)} onChange={() => toggleCadence(c)} />
                  {c}
                </label>
              ))}
            </div>
          </Field>
        </div>
        <Field label="Start date">
          <input type="date" className={inputCls} style={inputStyle} value={form.start} onChange={set("start")} />
        </Field>
        <Field label="End date">
          <input type="date" className={inputCls} style={inputStyle} value={form.end} onChange={set("end")} />
        </Field>
        <Field label="Compliance owner">
          <input className={inputCls} style={inputStyle} value={form.complianceOwner} onChange={set("complianceOwner")} />
        </Field>
        <Field label="Finance owner">
          <input className={inputCls} style={inputStyle} value={form.financeOwner} onChange={set("financeOwner")} />
        </Field>
        <Field label="Internal owner">
          <input className={inputCls} style={inputStyle} value={form.internalOwner} onChange={set("internalOwner")} />
        </Field>
        <Field label="Operations owner">
          <input className={inputCls} style={inputStyle} value={form.operationsOwner} onChange={set("operationsOwner")} />
        </Field>
        <Field label="Budget group (optional)">
          <div className="flex items-center gap-2">
            <select value={form.budgetGroupId || ""} onChange={set("budgetGroupId")} className={inputCls} style={inputStyle}>
              <option value="">No group</option>
              {(budgetGroups || []).map((bg) => <option key={bg.id} value={bg.id}>{bg.name}</option>)}
            </select>
            <button onClick={() => setBgModal("new")} className="shrink-0 text-xs px-3 py-2 rounded-md border" style={{ borderColor: "#E1E5DE", color: "#1F5C6B" }}>+ New</button>
          </div>
        </Field>
        <Field label="Document library URL">
          <input className={inputCls} style={inputStyle} value={form.doclibUrl} onChange={set("doclibUrl")} placeholder="https://…" />
        </Field>
        <Field label="Contract URL (SharePoint)">
          <input className={inputCls} style={inputStyle} value={form.contractUrl} onChange={set("contractUrl")} placeholder="https://…sharepoint.com/…" />
        </Field>
        <Field label="Budget period start">
          <input type="date" className={inputCls} style={inputStyle} value={form.budgetPeriodStart} onChange={set("budgetPeriodStart")} />
        </Field>
        <Field label="Budget period end">
          <input type="date" className={inputCls} style={inputStyle} value={form.budgetPeriodEnd} onChange={set("budgetPeriodEnd")} />
        </Field>
        <Field label="Obligated funds">
          <input type="number" className={inputCls} style={inputStyle} value={form.obligatedFunds} onChange={set("obligatedFunds")} />
        </Field>
        <Field label="Obligated funds remaining">
          <input type="number" className={inputCls} style={inputStyle} value={form.obligatedFundsRemaining} onChange={set("obligatedFundsRemaining")} />
        </Field>
        <Field label="Payment method">
          <select className={inputCls} style={inputStyle} value={form.paymentMethod} onChange={set("paymentMethod")}>
            {PAYMENT_METHODS.map((s) => <option key={s}>{s}</option>)}
          </select>
        </Field>
        <Field label="Grant POC">
          <input className={inputCls} style={inputStyle} value={form.grantPoc} onChange={set("grantPoc")} placeholder="Point of contact name" />
        </Field>
        <Field label="Beds">
          <input type="number" className={inputCls} style={inputStyle} value={form.beds} onChange={set("beds")} />
        </Field>
        <Field label="Bed rate">
          <input type="number" className={inputCls} style={inputStyle} value={form.bedRate} onChange={set("bedRate")} placeholder="$ per bed" />
        </Field>
        <div className="col-span-2">
          <Field label="Notes">
            <textarea className={inputCls} style={inputStyle} rows={2} value={form.notes} onChange={set("notes")} />
          </Field>
        </div>
        <label className="col-span-2 flex items-center gap-2 text-sm" style={{ color: "#1C2624" }}>
          <input type="checkbox" checked={!!form.renewal} onChange={set("renewal")} />
          Up for renewal
        </label>
      </div>
      </fieldset>
      <div className="flex justify-end gap-2 mt-6">
        <button onClick={onClose} className="px-4 py-2 rounded-md text-sm border" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>Cancel</button>
        {canEdit && (
          <button
            onClick={() => { if (!form.title.trim()) return; onSave(form); }}
            className="px-4 py-2 rounded-md text-sm text-white"
            style={{ background: "#1F5C6B" }}
          >
            Save grant
          </button>
        )}
      </div>
      {bgModal && canEdit && <BudgetGroupModal budgetGroup={bgModal === "new" ? null : bgModal} onSave={saveBudgetGroup} onClose={() => setBgModal(null)} />}
    </Modal>
  );
}

// ---------- budget form ----------

function BudgetModal({ budget, grantId, costCenterId, canEdit = true, onSave, onClose }) {
  const [form, setForm] = useState(budget || {
    id: uid(), grantId, costCenterId, title: "", fy: "", periodStart: "", periodEnd: "",
    status: "Draft", notes: "", lines: [newLine()],
    approvedBy: "", approvedAt: "", rejectionReason: "",
  });
  const [approverName, setApproverName] = useState("");
  const [rejectReason, setRejectReason] = useState("");
  const [showRejectBox, setShowRejectBox] = useState(false);
  const set = (k) => (e) => setForm({ ...form, [k]: e.target.value });

  const updateLine = (id, patch) => {
    setForm({ ...form, lines: form.lines.map((l) => (l.id === id ? { ...l, ...patch } : l)) });
  };
  const updateAmount = (id, idx, val, field = "amounts") => {
    setForm({
      ...form,
      lines: form.lines.map((l) => {
        if (l.id !== id) return l;
        const arr = [...(l[field] || Array(12).fill(0))];
        arr[idx] = val === "" ? 0 : Number(val);
        return { ...l, [field]: arr };
      }),
    });
  };
  const addLine = () => setForm({ ...form, lines: [...form.lines, newLine()] });
  const deleteLine = (id) => setForm({ ...form, lines: form.lines.filter((l) => l.id !== id) });
  const [lineSort, setLineSort] = useState("none");
  const [yearlyDraft, setYearlyDraft] = useState({});
  const [mode, setMode] = useState("plan");
  const field = mode === "plan" ? "amounts" : "actuals";

  const applyYearlyTotal = (id) => {
    const val = yearlyDraft[id];
    if (val === undefined || val === "") return;
    updateLine(id, { [field]: distributeEvenly(val) });
  };

  const submitForApproval = () => setForm({ ...form, status: "Pending Approval", rejectionReason: "" });
  const approveBudget = () => {
    if (!approverName.trim()) return;
    setForm({ ...form, status: "Active", approvedBy: approverName.trim(), approvedAt: new Date().toISOString().slice(0, 10) });
  };
  const rejectBudget = () => {
    if (!rejectReason.trim()) return;
    setForm({ ...form, status: "Rejected", rejectionReason: rejectReason.trim() });
    setShowRejectBox(false);
  };
  const revise = () => setForm({ ...form, status: "Draft" });

  const totals = budgetTotals(form);
  const actualTotals = budgetActualTotals(form);
  const sortedLines = [...form.lines].sort((a, b) => {
    if (lineSort === "category") return a.category.localeCompare(b.category);
    if (lineSort === "subcategory") return (a.subcategory || "").localeCompare(b.subcategory || "");
    if (lineSort === "type") return a.type.localeCompare(b.type);
    if (lineSort === "total") return lineTotal(b) - lineTotal(a);
    return 0;
  });

  return (
    <Modal title={budget ? "Edit budget" : "New budget"} onClose={onClose} size="xl">
      {budget?.status === "Closed" && (
        <div className="rounded-md px-3 py-2 mb-4 flex items-start gap-2" style={{ background: "#FBEAE8", border: "1px solid #B5443A" }}>
          <AlertCircle size={15} style={{ color: "#B5443A", marginTop: 1 }} className="shrink-0" />
          <div className="text-sm" style={{ color: "#B5443A" }}>
            This budget is marked <strong>Closed</strong>. It's usually treated as finalized — double-check before making changes.
          </div>
        </div>
      )}

      <div className="rounded-md px-3 py-3 mb-4" style={{ background: "#F6F7F3", border: "1px solid #E1E5DE" }}>
        {form.status === "Draft" && (
          <div className="flex items-center justify-between">
            <span className="text-sm" style={{ color: "#5B6B66" }}>This budget is a draft. Submit it for approval before it can go Active.</span>
            {canEdit && (
              <button onClick={submitForApproval} className="text-xs px-3 py-1.5 rounded-md text-white shrink-0" style={{ background: "#1F5C6B" }}>Submit for approval</button>
            )}
          </div>
        )}
        {form.status === "Pending Approval" && !showRejectBox && (
          <div>
            <div className="text-sm mb-2" style={{ color: "#5B6B66" }}>Awaiting approval before this budget can go Active.</div>
            {canEdit && (
              <div className="flex items-center gap-2">
                <input
                  value={approverName} onChange={(e) => setApproverName(e.target.value)}
                  placeholder="Approver name" className="rounded-md border px-2 py-1.5 text-sm flex-1" style={inputStyle}
                />
                <button onClick={approveBudget} className="text-xs px-3 py-1.5 rounded-md text-white shrink-0" style={{ background: "#1F5C6B" }}>Approve</button>
                <button onClick={() => setShowRejectBox(true)} className="text-xs px-3 py-1.5 rounded-md border shrink-0" style={{ borderColor: "#B5443A", color: "#B5443A" }}>Reject</button>
              </div>
            )}
          </div>
        )}
        {form.status === "Pending Approval" && showRejectBox && canEdit && (
          <div>
            <div className="text-sm mb-2" style={{ color: "#5B6B66" }}>Reason for rejection:</div>
            <div className="flex items-center gap-2">
              <input
                value={rejectReason} onChange={(e) => setRejectReason(e.target.value)}
                placeholder="e.g. Numbers don't match award amount" className="rounded-md border px-2 py-1.5 text-sm flex-1" style={inputStyle}
              />
              <button onClick={rejectBudget} className="text-xs px-3 py-1.5 rounded-md text-white shrink-0" style={{ background: "#B5443A" }}>Confirm reject</button>
              <button onClick={() => setShowRejectBox(false)} className="text-xs px-3 py-1.5 rounded-md border shrink-0" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>Cancel</button>
            </div>
          </div>
        )}
        {form.status === "Rejected" && (
          <div className="flex items-center justify-between">
            <div className="text-sm" style={{ color: "#B5443A" }}>
              <strong>Rejected: </strong>{form.rejectionReason || "No reason given."}
            </div>
            {canEdit && (
              <button onClick={revise} className="text-xs px-3 py-1.5 rounded-md border shrink-0" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>Revise & resubmit</button>
            )}
          </div>
        )}
        {form.status === "Active" && (
          <div className="text-sm" style={{ color: "#2F6F53" }}>
            <CheckCircle size={13} className="inline mr-1" style={{ marginBottom: 2 }} />
            Approved{form.approvedBy ? ` by ${form.approvedBy}` : ""}{form.approvedAt ? ` on ${fmtDate(form.approvedAt)}` : ""}
          </div>
        )}
        {form.status === "Closed" && (
          <div className="text-sm" style={{ color: "#8A8F87" }}>This budget is closed.</div>
        )}
      </div>

      <fieldset disabled={!canEdit} style={{ border: "none", margin: 0, padding: 0 }}>
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-4 mb-4">
        <Field label="Budget title">
          <input className={inputCls} style={inputStyle} value={form.title} onChange={set("title")} placeholder="e.g. FY26 Operating Budget" />
        </Field>
        <Field label="Fiscal year">
          <input className={inputCls} style={inputStyle} value={form.fy} onChange={set("fy")} placeholder="e.g. FY26" />
        </Field>
        <Field label="Status">
          <select className={inputCls} style={inputStyle} value={form.status} onChange={set("status")}>
            {BUDGET_STATUSES.map((s) => <option key={s}>{s}</option>)}
          </select>
        </Field>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
          <Field label="Period start">
            <input type="date" className={inputCls} style={inputStyle} value={form.periodStart} onChange={set("periodStart")} />
          </Field>
          <Field label="Period end">
            <input type="date" className={inputCls} style={inputStyle} value={form.periodEnd} onChange={set("periodEnd")} />
          </Field>
        </div>
      </div>

      <div className="flex items-center justify-between mb-2">
        <h3 className="text-sm font-medium" style={{ color: "#1C2624" }}>Budget line items</h3>
        <div className="flex items-center gap-2">
          <select value={lineSort} onChange={(e) => setLineSort(e.target.value)} className="rounded-md border px-2 py-1.5 text-xs" style={inputStyle}>
            <option value="none">Sort: Row order</option>
            <option value="category">Sort: Category</option>
            <option value="subcategory">Sort: Subcategory</option>
            <option value="type">Sort: Revenue / Expense</option>
            <option value="total">Sort: Total (high–low)</option>
          </select>
          <button onClick={addLine} className="inline-flex items-center gap-1 text-sm px-3 py-1.5 rounded-md border" style={{ borderColor: "#E1E5DE", color: "#1F5C6B" }}>
            <Plus size={14} /> Add row
          </button>
        </div>
      </div>

      <div className="flex items-center justify-between mb-2">
        <p className="text-xs" style={{ color: "#8A8F87" }}>
          Enter a number in <strong>Annual total</strong> to split it evenly across all 12 months, then fine-tune any month directly.
        </p>
        <div className="inline-flex rounded-md border overflow-hidden shrink-0" style={{ borderColor: "#E1E5DE" }}>
          <button
            onClick={() => setMode("plan")}
            className="px-3 py-1.5 text-xs font-medium"
            style={{ background: mode === "plan" ? "#2F6F53" : "#FFFFFF", color: mode === "plan" ? "#FFFFFF" : "#5B6B66" }}
          >
            Plan
          </button>
          <button
            onClick={() => setMode("actual")}
            className="px-3 py-1.5 text-xs font-medium"
            style={{ background: mode === "actual" ? "#2F6F53" : "#FFFFFF", color: mode === "actual" ? "#FFFFFF" : "#5B6B66" }}
          >
            Actual
          </button>
        </div>
      </div>

      <div className="overflow-x-auto border rounded-lg" style={{ borderColor: "#E1E5DE" }}>
        <table className="text-xs w-full" style={{ fontFamily: "var(--mono-font)" }}>
          <thead>
            <tr style={{ background: "#F6F7F3" }}>
              <th className="text-left px-2 py-2 sticky left-0" style={{ background: "#F6F7F3", minWidth: 230 }}>Category</th>
              <th className="text-left px-2 py-2" style={{ minWidth: 230 }}>Subcategory</th>
              <th className="text-left px-2 py-2" style={{ minWidth: 180 }}>Description</th>
              <th className="text-right px-2 py-2" style={{ minWidth: 110 }}>Annual total</th>
              {monthColumnsForBudget(form.periodStart).map((col, i) => <th key={i} className="text-right px-2 py-2" style={{ minWidth: 78 }}>{col.label}</th>)}
              <th className="text-right px-2 py-2" style={{ minWidth: 90 }}>Total</th>
              <th className="px-2 py-2"></th>
            </tr>
          </thead>
          <tbody>
            {form.lines.length === 0 && (
              <tr><td colSpan={16} className="text-center py-6" style={{ color: "#8A8F87" }}>No expense lines yet.</td></tr>
            )}
            {sortedLines.map((line) => {
              const cat = CATEGORIES.find((c) => c.name === line.category);
              const values = line[field] || Array(12).fill(0);
              return (
                <tr key={line.id} className="border-t" style={{ borderColor: "#E1E5DE" }}>
                  <td className="px-2 py-1.5 sticky left-0 bg-white">
                    {line.categoryCustom ? (
                      <div className="flex gap-1">
                        <input
                          value={line.category}
                          onChange={(e) => updateLine(line.id, { category: e.target.value })}
                          placeholder="Custom category"
                          className="w-full rounded border px-1.5 py-1 text-xs"
                          style={inputStyle}
                          autoFocus
                        />
                        <select
                          value={line.type}
                          onChange={(e) => updateLine(line.id, { type: e.target.value })}
                          className="shrink-0 rounded border px-1 py-1 text-xs"
                          style={inputStyle}
                        >
                          <option value="expense">Exp</option>
                          <option value="revenue">Rev</option>
                        </select>
                        <button
                          onClick={() => updateLine(line.id, { categoryCustom: false, category: CATEGORIES[0].name, type: CATEGORIES[0].type, subcategory: "", subcategoryCustom: false })}
                          className="shrink-0 px-1 rounded hover:bg-red-50"
                          title="Back to category list"
                        >
                          <X size={12} style={{ color: "#B5443A" }} />
                        </button>
                      </div>
                    ) : (
                      <select
                        value={line.category}
                        onChange={(e) => {
                          if (e.target.value === CUSTOM_CATEGORY) {
                            updateLine(line.id, { categoryCustom: true, category: "", subcategory: "", subcategoryCustom: false });
                            return;
                          }
                          const nc = CATEGORIES.find((c) => c.name === e.target.value);
                          updateLine(line.id, { category: nc.name, type: nc.type, subcategory: "", subcategoryCustom: false });
                        }}
                        className="w-full rounded border px-1.5 py-1 text-xs"
                        style={inputStyle}
                      >
                        {CATEGORIES.map((c) => <option key={c.name} value={c.name}>{c.name}</option>)}
                        <option value={CUSTOM_CATEGORY}>Other (write in)…</option>
                      </select>
                    )}
                  </td>
                  <td className="px-2 py-1.5">
                    {line.categoryCustom || line.subcategoryCustom ? (
                      <div className="flex gap-1">
                        <input
                          value={line.subcategory}
                          onChange={(e) => updateLine(line.id, { subcategory: e.target.value })}
                          placeholder="Custom subcategory"
                          className="w-full rounded border px-1.5 py-1 text-xs"
                          style={inputStyle}
                          autoFocus={line.subcategoryCustom && !line.categoryCustom}
                        />
                        {!line.categoryCustom && (
                          <button
                            onClick={() => updateLine(line.id, { subcategoryCustom: false, subcategory: "" })}
                            className="shrink-0 px-1 rounded hover:bg-red-50"
                            title="Back to subcategory list"
                          >
                            <X size={12} style={{ color: "#B5443A" }} />
                          </button>
                        )}
                      </div>
                    ) : (
                      <select
                        value={line.subcategory}
                        onChange={(e) => {
                          if (e.target.value === CUSTOM_CATEGORY) {
                            updateLine(line.id, { subcategoryCustom: true, subcategory: "" });
                            return;
                          }
                          updateLine(line.id, { subcategory: e.target.value });
                        }}
                        className="w-full rounded border px-1.5 py-1 text-xs"
                        style={inputStyle}
                      >
                        <option value="">Select subcategory</option>
                        {cat?.subs.map((s) => <option key={s} value={s}>{s}</option>)}
                        <option value={CUSTOM_CATEGORY}>Other (write in)…</option>
                      </select>
                    )}
                  </td>
                  <td className="px-2 py-1.5">
                    <input
                      value={line.description || ""}
                      onChange={(e) => updateLine(line.id, { description: e.target.value })}
                      placeholder="Optional note"
                      className="w-full rounded border px-1.5 py-1 text-xs"
                      style={inputStyle}
                    />
                  </td>
                  <td className="px-2 py-1.5">
                    <input
                      type="number"
                      value={yearlyDraft[line.id] ?? ""}
                      placeholder={fmt(mode === "plan" ? lineTotal(line) : lineActualTotal(line)).replace("$", "")}
                      onChange={(e) => setYearlyDraft({ ...yearlyDraft, [line.id]: e.target.value })}
                      onBlur={() => applyYearlyTotal(line.id)}
                      onKeyDown={(e) => { if (e.key === "Enter") { e.preventDefault(); applyYearlyTotal(line.id); } }}
                      className="w-full rounded border px-1.5 py-1 text-xs text-right"
                      style={{ ...inputStyle, fontVariantNumeric: "tabular-nums" }}
                    />
                  </td>
                  {values.map((amt, idx) => (
                    <td key={idx} className="px-1 py-1.5">
                      <input
                        type="number"
                        value={amt === 0 ? "" : amt}
                        placeholder="0"
                        onChange={(e) => updateAmount(line.id, idx, e.target.value, field)}
                        className="w-full rounded border px-1.5 py-1 text-xs text-right"
                        style={{ ...inputStyle, fontVariantNumeric: "tabular-nums" }}
                      />
                    </td>
                  ))}
                  <td className="px-2 py-1.5 text-right font-medium" style={{ fontVariantNumeric: "tabular-nums", color: line.type === "revenue" ? "#2F6F53" : "#1C2624" }}>
                    {fmt(mode === "plan" ? lineTotal(line) : lineActualTotal(line))}
                  </td>
                  <td className="px-2 py-1.5 text-center">
                    <button onClick={() => deleteLine(line.id)} className="p-1 rounded hover:bg-red-50">
                      <Trash2 size={14} style={{ color: "#B5443A" }} />
                    </button>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      <div className="grid grid-cols-1 sm:grid-cols-3 gap-6 mt-4 text-sm">
        <div>
          <div className="text-xs font-medium mb-1" style={{ color: "#8A8F87" }}>Plan</div>
          <div><span style={{ color: "#5B6B66" }}>Revenue: </span><span className="font-medium" style={{ color: "#2F6F53" }}>{fmt(totals.revenue)}</span></div>
          <div><span style={{ color: "#5B6B66" }}>Expense: </span><span className="font-medium" style={{ color: "#1C2624" }}>{fmt(totals.expense)}</span></div>
          <div><span style={{ color: "#5B6B66" }}>Net: </span><span className="font-medium" style={{ color: totals.net >= 0 ? "#2F6F53" : "#B5443A" }}>{fmt(totals.net)}</span></div>
        </div>
        <div>
          <div className="text-xs font-medium mb-1" style={{ color: "#8A8F87" }}>Actual</div>
          <div><span style={{ color: "#5B6B66" }}>Revenue: </span><span className="font-medium" style={{ color: "#2F6F53" }}>{fmt(actualTotals.revenue)}</span></div>
          <div><span style={{ color: "#5B6B66" }}>Expense: </span><span className="font-medium" style={{ color: "#1C2624" }}>{fmt(actualTotals.expense)}</span></div>
          <div><span style={{ color: "#5B6B66" }}>Net: </span><span className="font-medium" style={{ color: actualTotals.net >= 0 ? "#2F6F53" : "#B5443A" }}>{fmt(actualTotals.net)}</span></div>
        </div>
        <div>
          <div className="text-xs font-medium mb-1" style={{ color: "#8A8F87" }}>Variance (Actual − Plan)</div>
          <div><span style={{ color: "#5B6B66" }}>Revenue: </span><span className="font-medium" style={{ color: "#1C2624" }}>{fmt(actualTotals.revenue - totals.revenue)}</span></div>
          <div><span style={{ color: "#5B6B66" }}>Expense: </span><span className="font-medium" style={{ color: (actualTotals.expense - totals.expense) > 0 ? "#B5443A" : "#2F6F53" }}>{fmt(actualTotals.expense - totals.expense)}</span></div>
          <div><span style={{ color: "#5B6B66" }}>Net: </span><span className="font-medium" style={{ color: "#1C2624" }}>{fmt(actualTotals.net - totals.net)}</span></div>
        </div>
      </div>

      <Field label="Notes">
        <textarea className={inputCls} style={inputStyle} rows={2} value={form.notes} onChange={set("notes")} />
      </Field>
      </fieldset>

      <div className="flex justify-end gap-2 mt-4">
        <button onClick={onClose} className="px-4 py-2 rounded-md text-sm border" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>Cancel</button>
        {canEdit && (
          <button
            onClick={() => { if (!form.title.trim()) return; onSave(form); }}
            className="px-4 py-2 rounded-md text-sm text-white"
            style={{ background: "#1F5C6B" }}
          >
            Save budget
          </button>
        )}
      </div>
    </Modal>
  );
}

// ---------- grant report form ----------

function ReportModal({ report, grants, canEdit = true, onSave, onClose, onDelete, onCreateTask }) {
  const [taskCreated, setTaskCreated] = useState(!!report?.linkedTaskCreated);
  const [form, setForm] = useState(report || {
    id: uid(), title: "", grantId: "", assignedTo: "", status: "Not started",
    priority: "Medium", startDate: "", dueDate: "", repeat: "None", repeatDetail: "",
    bucket: DEFAULT_BUCKETS[0], checklist: [], notes: "", portalUrl: "", linkedTaskCreated: false,
    createdAt: new Date().toISOString().slice(0, 10),
  });
  const [newStep, setNewStep] = useState("");
  const set = (k) => (e) => setForm({ ...form, [k]: e.target.value });

  const addStep = () => {
    if (!newStep.trim()) return;
    setForm({ ...form, checklist: [...form.checklist, { id: uid(), text: newStep.trim(), done: false }] });
    setNewStep("");
  };
  const toggleStep = (id) => setForm({ ...form, checklist: form.checklist.map((s) => (s.id === id ? { ...s, done: !s.done } : s)) });
  const deleteStep = (id) => setForm({ ...form, checklist: form.checklist.filter((s) => s.id !== id) });

  const grant = grants.find((g) => g.id === form.grantId);

  return (
    <Modal title={report ? (canEdit ? "Edit report" : "View report") : "New report"} onClose={onClose} wide>
      <fieldset disabled={!canEdit} style={{ border: "none", margin: 0, padding: 0 }}>
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
        <div className="col-span-2">
          <Field label="Report title">
            <input className={inputCls} style={inputStyle} value={form.title} onChange={set("title")} placeholder="e.g. Monthly Invoice" />
          </Field>
        </div>

        <Field label="Grant">
          <GrantPicker grants={grants} value={form.grantId} onChange={(v) => setForm({ ...form, grantId: v })} noneLabel="No grant linked" />
        </Field>
        <Field label="Assigned to">
          <input className={inputCls} style={inputStyle} value={form.assignedTo} onChange={set("assignedTo")} placeholder="Name" />
        </Field>

        <Field label="Status">
          <select className={inputCls} style={inputStyle} value={form.status} onChange={set("status")}>
            {REPORT_STATUSES.map((s) => <option key={s}>{s}</option>)}
          </select>
        </Field>
        <Field label="Priority">
          <select className={inputCls} style={inputStyle} value={form.priority} onChange={set("priority")}>
            {REPORT_PRIORITIES.map((p) => <option key={p.label}>{p.label}</option>)}
          </select>
        </Field>

        <Field label="Start date">
          <input type="date" className={inputCls} style={inputStyle} value={form.startDate} onChange={set("startDate")} />
        </Field>
        <Field label="Due date">
          <input type="date" className={inputCls} style={inputStyle} value={form.dueDate} onChange={set("dueDate")} />
        </Field>

        <Field label="Repeat">
          <select className={inputCls} style={inputStyle} value={form.repeat} onChange={set("repeat")}>
            {REPORT_REPEATS.map((r) => <option key={r}>{r}</option>)}
          </select>
        </Field>
        <Field label="Bucket">
          <select className={inputCls} style={inputStyle} value={form.bucket} onChange={set("bucket")}>
            {[...new Set([...DEFAULT_BUCKETS, form.bucket])].map((b) => <option key={b} value={b}>{b}</option>)}
          </select>
        </Field>
        <Field label="Submission portal URL">
          <input className={inputCls} style={inputStyle} value={form.portalUrl} onChange={set("portalUrl")} placeholder="https://…" />
        </Field>

        {form.repeat !== "None" && (
          <div className="col-span-2">
            <Field label="Repeat details">
              <input className={inputCls} style={inputStyle} value={form.repeatDetail} onChange={set("repeatDetail")} placeholder="e.g. on the second Tuesday" />
            </Field>
          </div>
        )}
      </div>

      <div className="mt-5">
        <h3 className="text-sm font-medium mb-2" style={{ color: "#1C2624" }}>Checklist</h3>
        <div className="space-y-1.5">
          {form.checklist.map((s) => (
            <div key={s.id} className="flex items-center gap-2 text-sm">
              <button onClick={() => toggleStep(s.id)} className="shrink-0">
                {s.done ? <CheckCircle size={16} style={{ color: "#2F6F53" }} /> : <Circle size={16} style={{ color: "#8A8F87" }} />}
              </button>
              <span className={s.done ? "line-through" : ""} style={{ color: s.done ? "#8A8F87" : "#1C2624" }}>{s.text}</span>
              <button onClick={() => deleteStep(s.id)} className="ml-auto p-1 rounded hover:bg-red-50">
                <X size={13} style={{ color: "#B5443A" }} />
              </button>
            </div>
          ))}
        </div>
        <div className="flex gap-2 mt-2">
          <input
            className={inputCls} style={inputStyle} value={newStep}
            onChange={(e) => setNewStep(e.target.value)}
            onKeyDown={(e) => { if (e.key === "Enter") { e.preventDefault(); addStep(); } }}
            placeholder="Add steps to complete this report…"
          />
          <button onClick={addStep} className="px-3 py-2 rounded-md text-sm border shrink-0" style={{ borderColor: "#E1E5DE", color: "#1F5C6B" }}>Add</button>
        </div>
      </div>

      <div className="mt-4 rounded-md px-3 py-3" style={{ background: "#F6F7F3", border: "1px solid #E1E5DE" }}>
        {taskCreated ? (
          <div className="flex items-center gap-2 text-sm" style={{ color: "#2F6F53" }}>
            <CheckCircle size={14} /> A task has been created for this report in Tasks.
          </div>
        ) : (
          <div className="flex items-center justify-between gap-3">
            <span className="text-sm" style={{ color: "#5B6B66" }}>
              Want this tracked as a to-do too? Create a linked task for {form.assignedTo || "whoever you assign it to"} in the Tasks module.
            </span>
            <button
              onClick={() => {
                if (!form.title.trim()) return;
                onCreateTask?.(form);
                setTaskCreated(true);
                setForm({ ...form, linkedTaskCreated: true });
              }}
              className="text-xs px-3 py-1.5 rounded-md text-white shrink-0"
              style={{ background: "#1F5C6B" }}
            >
              Create linked task
            </button>
          </div>
        )}
      </div>

      <div className="mt-4">
        <Field label="Notes">
          <textarea className={inputCls} style={inputStyle} rows={4} value={form.notes} onChange={set("notes")} placeholder="Reporting requirements, citations, submission instructions…" />
        </Field>
      </div>
      </fieldset>

      <div className="flex justify-between gap-2 mt-6">
        {onDelete ? (
          <button onClick={onDelete} className="px-4 py-2 rounded-md text-sm border" style={{ borderColor: "#E1E5DE", color: "#B5443A" }}>Delete report</button>
        ) : <span />}
        <div className="flex gap-2">
          <button onClick={onClose} className="px-4 py-2 rounded-md text-sm border" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>Cancel</button>
          {canEdit && (
            <button
              onClick={() => { if (!form.title.trim()) return; onSave(form); }}
              className="px-4 py-2 rounded-md text-sm text-white"
              style={{ background: "#1F5C6B" }}
            >
              Save report
            </button>
          )}
        </div>
      </div>
    </Modal>
  );
}



// ---------- main views ----------

function Dashboard({ grants, budgets, reports, tasks, staff, invoices, goTo }) {
  const activeGrants = grants.filter((g) => g.stage === "Active");
  const totalAward = activeGrants.reduce((a, g) => a + (Number(g.awardAmount) || 0), 0);
  const totalRemaining = activeGrants.reduce((a, g) => {
    const t = grantBudgetTotals(g.id, budgets);
    return a + ((Number(g.awardAmount) || 0) - t.expense);
  }, 0);
  const activeGrantIds = new Set(activeGrants.map((g) => g.id));
  const activeBudgets = budgets.filter((b) => b.status === "Active" && activeGrantIds.has(b.grantId)).length;

  const fyYear = new Date().getFullYear();
  const fyStart = new Date(fyYear, 0, 1);
  const fyEnd = new Date(fyYear, 11, 31);
  const fyGrants = activeGrants.filter((g) => {
    const start = g.start ? new Date(g.start) : null;
    const end = g.end ? new Date(g.end) : null;
    return (!start || start <= fyEnd) && (!end || end >= fyStart);
  });
  const fyTotalAward = fyGrants.reduce((a, g) => a + (Number(g.obligatedFunds) || 0), 0);
  const fyTotalRemaining = fyGrants.reduce((a, g) => a + (Number(g.obligatedFundsRemaining) || 0), 0);
  const upcoming = [...grants]
    .filter((g) => g.end)
    .sort((a, b) => new Date(a.end) - new Date(b.end))
    .filter((g) => new Date(g.end) >= new Date(new Date().toDateString()))
    .slice(0, 5);
  const upcomingReports = [...reports]
    .filter((r) => r.dueDate && r.status !== "Completed")
    .sort((a, b) => new Date(a.dueDate) - new Date(b.dueDate))
    .slice(0, 5);
  const upcomingTasks = [...tasks]
    .filter((t) => t.dueDate && t.status !== "Done")
    .sort((a, b) => new Date(a.dueDate) - new Date(b.dueDate))
    .slice(0, 5);
  const closedGrantsWithStaff = grants
    .filter((g) => g.stage === "Closed")
    .map((g) => ({
      grant: g,
      staffAllocs: (staff || []).filter((s) => (s.allocations || []).some((a) => a.grantId === g.id && Number(a.percent) > 0)),
    }))
    .filter((x) => x.staffAllocs.length > 0);
  const overdueInvoices = (invoices || []).filter(isInvoiceOverdue);

  return (
    <div className="space-y-6">
      <div className="flex items-center justify-between">
        <div>
          <h1 className="font-display text-2xl" style={{ color: "#1C2624" }}>Dashboard</h1>
          <p className="text-sm mt-1" style={{ color: "#5B6B66" }}>Active grants only</p>
        </div>
        <button onClick={() => printSection("dashboard-print-area", "GrantFlow Dashboard")} className="no-print inline-flex items-center gap-1.5 px-3 py-2 rounded-md text-sm border" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>
          <Printer size={15} /> Print / Save PDF
        </button>
      </div>

      <div id="dashboard-print-area">

      {closedGrantsWithStaff.length > 0 && (
        <div className="rounded-lg px-4 py-3 flex items-start gap-3" style={{ background: "#FBEAE8", border: "1px solid #B5443A" }}>
          <AlertCircle size={18} style={{ color: "#B5443A", marginTop: 1 }} className="shrink-0" />
          <div className="flex-1">
            <div className="font-medium text-sm" style={{ color: "#B5443A" }}>
              {closedGrantsWithStaff.length} closed grant{closedGrantsWithStaff.length > 1 ? "s" : ""} still {closedGrantsWithStaff.length > 1 ? "have" : "has"} staff allocated
            </div>
            <ul className="text-xs mt-1 space-y-0.5" style={{ color: "#8A4A44" }}>
              {closedGrantsWithStaff.map(({ grant, staffAllocs }) => (
                <li key={grant.id}>
                  <strong>{grant.title}</strong>: {staffAllocs.map((s) => `${s.name} (${s.allocations.find((a) => a.grantId === grant.id)?.percent}%)`).join(", ")}
                </li>
              ))}
            </ul>
          </div>
          <button onClick={() => goTo("personnel")} className="shrink-0 text-xs px-3 py-1.5 rounded-md border" style={{ borderColor: "#B5443A", color: "#B5443A" }}>
            Go to Personnel
          </button>
        </div>
      )}

      {overdueInvoices.length > 0 && (
        <div className="rounded-lg px-4 py-3 flex items-start gap-3" style={{ background: "#FBEAE8", border: "1px solid #B5443A" }}>
          <AlertCircle size={18} style={{ color: "#B5443A", marginTop: 1 }} className="shrink-0" />
          <div className="flex-1">
            <div className="font-medium text-sm" style={{ color: "#B5443A" }}>
              {overdueInvoices.length} invoice{overdueInvoices.length > 1 ? "s" : ""} overdue for payment
            </div>
            <ul className="text-xs mt-1 space-y-0.5" style={{ color: "#8A4A44" }}>
              {overdueInvoices.slice(0, 5).map((inv) => {
                const g = grants.find((x) => x.id === inv.grantId);
                return (
                  <li key={inv.id}>
                    <strong>{inv.invoiceNumber || "Untitled invoice"}</strong> ({g?.title || "Unknown grant"}) — {fmt(inv.amount)}, expected {fmtDate(inv.dueDate)}
                  </li>
                );
              })}
            </ul>
          </div>
          <button onClick={() => goTo("invoicing")} className="shrink-0 text-xs px-3 py-1.5 rounded-md border" style={{ borderColor: "#B5443A", color: "#B5443A" }}>
            Go to Invoicing
          </button>
        </div>
      )}

      <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
        <StatCard label="Total active grants" value={activeGrants.length} />
        <StatCard label="Active budgets" value={activeBudgets} />
        <StatCard label="Total award" value={fmt(totalAward)} />
        <StatCard label="Award remaining" value={fmt(totalRemaining)} sub="Per budgeted plan, not actuals" />
      </div>

      <div>
        <h2 className="font-display text-base mb-2" style={{ color: "#1C2624" }}>Fiscal year {fyYear} (Jan–Dec)</h2>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          <StatCard label={`FY${fyYear} total award`} value={fmt(fyTotalAward)} sub={`${fyGrants.length} grants active in FY${fyYear}`} />
          <StatCard label={`FY${fyYear} award remaining`} value={fmt(fyTotalRemaining)} sub="Sum of Obligated funds remaining" />
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        <div className="col-span-2 space-y-6">
          <div className="bg-white rounded-lg border p-5" style={{ borderColor: "#E1E5DE" }}>
            <h2 className="font-display text-base mb-3" style={{ color: "#1C2624" }}>Upcoming grant end dates</h2>
            {upcoming.length === 0 ? (
              <p className="text-sm" style={{ color: "#8A8F87" }}>No grants ending soon.</p>
            ) : (
              <ul className="divide-y" style={{ borderColor: "#E1E5DE" }}>
                {upcoming.map((g) => (
                  <li key={g.id} className="py-2.5 flex items-center justify-between text-sm">
                    <div>
                      <div style={{ color: "#1C2624" }}>{g.title}</div>
                      <div className="text-xs" style={{ color: "#8A8F87" }}>{g.programCode}</div>
                    </div>
                    <div className="text-right">
                      <div style={{ color: "#1C2624", fontVariantNumeric: "tabular-nums" }}>{fmtDate(g.end)}</div>
                      <Badge color={riskColor[g.riskStatus]}>{g.riskStatus} risk</Badge>
                    </div>
                  </li>
                ))}
              </ul>
            )}
          </div>

          <div className="bg-white rounded-lg border p-5" style={{ borderColor: "#E1E5DE" }}>
            <h2 className="font-display text-base mb-3" style={{ color: "#1C2624" }}>Upcoming grant reports due</h2>
            {upcomingReports.length === 0 ? (
              <p className="text-sm" style={{ color: "#8A8F87" }}>No grant reports due soon.</p>
            ) : (
              <ul className="divide-y" style={{ borderColor: "#E1E5DE" }}>
                {upcomingReports.map((r) => {
                  const g = grants.find((x) => x.id === r.grantId);
                  const overdue = isOverdue(r);
                  return (
                    <li key={r.id} className="py-2.5 flex items-center justify-between text-sm">
                      <div>
                        <div style={{ color: "#1C2624" }}>{r.title}</div>
                        <div className="text-xs" style={{ color: "#8A8F87" }}>{g ? (g.programCode ? `${g.programCode} - ${g.title}` : g.title) : "No grant linked"}</div>
                      </div>
                      <div className="text-right">
                        <div style={{ color: overdue ? "#B5443A" : "#1C2624", fontVariantNumeric: "tabular-nums" }}>{overdue ? "Overdue: " : ""}{fmtDate(r.dueDate)}</div>
                        <span className="w-2 h-2 rounded-full inline-block" style={{ background: priorityColor(r.priority) }} title={r.priority} />
                      </div>
                    </li>
                  );
                })}
              </ul>
            )}
          </div>

          <div className="bg-white rounded-lg border p-5" style={{ borderColor: "#E1E5DE" }}>
            <h2 className="font-display text-base mb-3" style={{ color: "#1C2624" }}>Upcoming tasks</h2>
            {upcomingTasks.length === 0 ? (
              <p className="text-sm" style={{ color: "#8A8F87" }}>No tasks due soon.</p>
            ) : (
              <ul className="divide-y" style={{ borderColor: "#E1E5DE" }}>
                {upcomingTasks.map((t) => {
                  const g = grants.find((x) => x.id === t.grantId);
                  const overdue = t.dueDate && new Date(t.dueDate) < new Date(new Date().toDateString());
                  return (
                    <li key={t.id} className="py-2.5 flex items-center justify-between text-sm">
                      <div>
                        <div style={{ color: "#1C2624" }}>{t.title}</div>
                        <div className="text-xs" style={{ color: "#8A8F87" }}>{t.category}{g ? ` · ${g.title}` : ""}</div>
                      </div>
                      <div className="text-right">
                        <div style={{ color: overdue ? "#B5443A" : "#1C2624", fontVariantNumeric: "tabular-nums" }}>{overdue ? "Overdue: " : ""}{fmtDate(t.dueDate)}</div>
                        <span className="w-2 h-2 rounded-full inline-block" style={{ background: priorityColor(t.priority) }} title={t.priority} />
                      </div>
                    </li>
                  );
                })}
              </ul>
            )}
          </div>
        </div>

        <div className="bg-white rounded-lg border p-5 space-y-2 h-fit" style={{ borderColor: "#E1E5DE" }}>
          <h2 className="font-display text-base mb-1" style={{ color: "#1C2624" }}>Quick actions</h2>
          <button onClick={() => goTo("grants", "new")} className="w-full flex items-center justify-between px-3 py-2 rounded-md border text-sm" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>
            New grant <ArrowRight size={14} />
          </button>
          <button onClick={() => goTo("budgets", "new")} className="w-full flex items-center justify-between px-3 py-2 rounded-md border text-sm" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>
            Manage budgets <ArrowRight size={14} />
          </button>
          <button onClick={() => goTo("tasks", "new")} className="w-full flex items-center justify-between px-3 py-2 rounded-md border text-sm" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>
            New task <ArrowRight size={14} />
          </button>
          <button onClick={() => goTo("grant-reports")} className="w-full flex items-center justify-between px-3 py-2 rounded-md border text-sm" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>
            View grant reports <ArrowRight size={14} />
          </button>
          <button onClick={() => goTo("reporting")} className="w-full flex items-center justify-between px-3 py-2 rounded-md border text-sm" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>
            View reporting <ArrowRight size={14} />
          </button>
        </div>
      </div>
      </div>
    </div>
  );
}

function GrantsView({ grants, budgets, reports, tasks, invoices, staff, budgetGroups, setBudgetGroups, setGrants, setBudgets, setReports, setTasks, setStaff, setInvoices, setTrash, currentUserEmail, canEdit, autoOpenNew, initialExpandId, goTo, logActivity }) {
  const [search, setSearch] = useState("");
  const [stageFilter, setStageFilter] = useState("All");
  const [riskFilter, setRiskFilter] = useState("All");
  const [sortBy, setSortBy] = useState("title");
  const [modal, setModal] = useState(autoOpenNew ? "new" : null);
  const [confirm, setConfirm] = useState(null);
  const [expanded, setExpanded] = useState(stripNonce(initialExpandId) || null);

  const filtered = grants
    .filter((g) => {
      const matchesSearch = (g.title + " " + g.programCode + " " + (g.sites || []).join(" ")).toLowerCase().includes(search.toLowerCase());
      const matchesStage = stageFilter === "All" || g.stage === stageFilter;
      const matchesRisk = riskFilter === "All" || g.riskStatus === riskFilter;
      return matchesSearch && matchesStage && matchesRisk;
    })
    .sort((a, b) => {
      if (sortBy === "location") return (a.sites?.[0] || "").localeCompare(b.sites?.[0] || "");
      if (sortBy === "end") return new Date(a.end || 0) - new Date(b.end || 0);
      if (sortBy === "award") return (Number(b.awardAmount) || 0) - (Number(a.awardAmount) || 0);
      return (a.title || "").localeCompare(b.title || "");
    });

  const saveGrant = (g) => {
    const prevGrant = grants.find((x) => x.id === g.id);
    const justClosed = prevGrant && prevGrant.stage !== "Closed" && g.stage === "Closed";
    setGrants((prev) => {
      const exists = prev.some((x) => x.id === g.id);
      logActivity?.("Grant", exists ? "Updated" : "Created", g.title || "Untitled grant");
      return exists ? prev.map((x) => (x.id === g.id ? g : x)) : [...prev, g];
    });
    if (justClosed) {
      const allocatedStaff = (staff || []).filter((s) => (s.allocations || []).some((a) => a.grantId === g.id && Number(a.percent) > 0));
      if (allocatedStaff.length > 0) {
        const names = allocatedStaff.map((s) => {
          const pct = s.allocations.find((a) => a.grantId === g.id)?.percent;
          return `${s.name} (${pct}%)`;
        }).join(", ");
        setTasks?.((prev) => [...prev, {
          id: uid(), title: `Reassign staff off closed grant: ${g.title}`, category: "Personnel Reallocation",
          grantId: g.id, dueDate: new Date(Date.now() + 7 * 86400000).toISOString().slice(0, 10),
          priority: "Urgent", status: "Not started", assignedTo: "",
          notes: `Grant "${g.title}" was just marked Closed. Staff still allocated: ${names}. Reassign their time or update allocations in Personnel.`,
        }]);
        logActivity?.("Task", "Created", `Reassign staff off closed grant: ${g.title}`);
      }
    }
    setModal(null);
  };

  const deleteGrant = (id) => {
    const g = grants.find((x) => x.id === id);
    const cascadedBudgets = budgets.filter((b) => b.grantId === id);
    const cascadedReports = (reports || []).filter((r) => r.grantId === id);
    const cascadedTasks = (tasks || []).filter((t) => t.grantId === id);
    const cascadedInvoices = (invoices || []).filter((i) => i.grantId === id);
    const cascadedAllocations = (staff || [])
      .map((s) => ({ staffId: s.id, allocations: (s.allocations || []).filter((a) => a.grantId === id) }))
      .filter((x) => x.allocations.length > 0);
    pushTrash(setTrash, "grant", g, currentUserEmail, {
      budgets: cascadedBudgets, reports: cascadedReports, tasks: cascadedTasks, invoices: cascadedInvoices, staffAllocations: cascadedAllocations,
    });
    setGrants((prev) => prev.filter((g) => g.id !== id));
    setBudgets((prev) => prev.filter((b) => b.grantId !== id));
    setReports?.((prev) => prev.filter((r) => r.grantId !== id));
    setTasks?.((prev) => prev.filter((t) => t.grantId !== id));
    setInvoices?.((prev) => prev.filter((i) => i.grantId !== id));
    setStaff?.((prev) => prev.map((s) => ({ ...s, allocations: (s.allocations || []).filter((a) => a.grantId !== id) })));
    logActivity?.("Grant", "Deleted", g?.title || "Untitled grant");
    setConfirm(null);
    setExpanded(null);
  };

  return (
    <div className="space-y-4">
      <div className="flex items-center justify-between">
        <h1 className="font-display text-2xl" style={{ color: "#1C2624" }}>Grants</h1>
        {canEdit && (
          <button onClick={() => setModal("new")} className="inline-flex items-center gap-1.5 px-3 py-2 rounded-md text-sm text-white" style={{ background: "#1F5C6B" }}>
            <Plus size={16} /> New grant
          </button>
        )}
      </div>

      <div className="flex gap-3">
        <div className="relative flex-1">
          <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2" style={{ color: "#8A8F87" }} />
          <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Search grants…" className={inputCls + " pl-9"} style={inputStyle} />
        </div>
        <select value={stageFilter} onChange={(e) => setStageFilter(e.target.value)} className={inputCls} style={{ ...inputStyle, width: 160 }}>
          <option>All</option>{STAGES.map((s) => <option key={s}>{s}</option>)}
        </select>
        <select value={riskFilter} onChange={(e) => setRiskFilter(e.target.value)} className={inputCls} style={{ ...inputStyle, width: 160 }}>
          <option>All</option>{RISKS.map((s) => <option key={s}>{s}</option>)}
        </select>
        <select value={sortBy} onChange={(e) => setSortBy(e.target.value)} className={inputCls} style={{ ...inputStyle, width: 180 }}>
          <option value="title">Sort: Title (A–Z)</option>
          <option value="location">Sort: Location (A–Z)</option>
          <option value="end">Sort: End date</option>
          <option value="award">Sort: Award amount</option>
        </select>
      </div>

      {filtered.length === 0 ? (
        <div className="bg-white rounded-lg border p-10 text-center" style={{ borderColor: "#E1E5DE", color: "#8A8F87" }}>
          No grants match your filters.
        </div>
      ) : (
        <div className="space-y-3">
          {filtered.map((g) => {
            const totals = grantBudgetTotals(g.id, budgets);
            const remaining = (Number(g.awardAmount) || 0) - totals.expense;
            const isOpen = expanded === g.id;
            const closedWithStaff = g.stage === "Closed" ? (staff || []).filter((s) => (s.allocations || []).some((a) => a.grantId === g.id && Number(a.percent) > 0)) : [];
            return (
              <div key={g.id} className="bg-white rounded-lg border" style={{ borderColor: "#E1E5DE" }}>
                <button onClick={() => setExpanded(isOpen ? null : g.id)} className="w-full flex items-center justify-between px-5 py-4 text-left">
                  <div>
                    <div className="flex items-center gap-2">
                      <span className="font-medium" style={{ color: "#1C2624" }}>{g.title}</span>
                      <Badge color={stageColor[g.stage]}>{g.stage}</Badge>
                      <Badge color={riskColor[g.riskStatus]}>{g.riskStatus} risk</Badge>
                      {g.renewal && <Badge color="#A8791F">Up for renewal</Badge>}
                    </div>
                    <div className="text-xs mt-1" style={{ color: "#8A8F87" }}>{g.programCode} · {g.sites?.length ? g.sites.join(", ") : "no site set"} · ends {fmtDate(g.end)}</div>
                  </div>
                  <div className="text-right text-sm" style={{ fontVariantNumeric: "tabular-nums" }}>
                    <div style={{ color: "#1C2624" }}>{fmt(g.awardAmount)}</div>
                    <div style={{ color: remaining >= 0 ? "#2F6F53" : "#B5443A" }}>{fmt(remaining)} remaining (budgeted)</div>
                  </div>
                </button>
                {isOpen && (
                  <div className="border-t px-5 py-4 text-sm grid grid-cols-1 sm:grid-cols-2 gap-x-8 gap-y-2" style={{ borderColor: "#E1E5DE" }}>
                    {closedWithStaff.length > 0 && (
                      <div className="col-span-2 rounded-md px-3 py-2 flex items-start gap-2" style={{ background: "#FBEAE8", border: "1px solid #B5443A" }}>
                        <AlertCircle size={15} style={{ color: "#B5443A", marginTop: 1 }} className="shrink-0" />
                        <div>
                          <div className="font-medium" style={{ color: "#B5443A" }}>This grant is closed but staff are still allocated to it</div>
                          <div className="text-xs mt-0.5" style={{ color: "#8A4A44" }}>
                            {closedWithStaff.map((s) => `${s.name} (${s.allocations.find((a) => a.grantId === g.id)?.percent}%)`).join(", ")}
                            {" — "}reassign them in Personnel.
                          </div>
                        </div>
                      </div>
                    )}
                    <div><span style={{ color: "#8A8F87" }}>Funding source: </span><span style={{ color: "#1C2624" }}>{g.funding || "—"}</span></div>
                    <div className="col-span-2">
                      <span style={{ color: "#8A8F87" }}>Sites: </span>
                      {g.sites?.length ? (
                        <span className="inline-flex flex-wrap gap-1 align-middle">
                          {g.sites.map((s) => <Badge key={s} color="#5B7FA6">{s}</Badge>)}
                        </span>
                      ) : <span style={{ color: "#1C2624" }}>—</span>}
                    </div>
                    <div className="col-span-2">
                      <span style={{ color: "#8A8F87" }}>Reporting cadence: </span>
                      {g.cadence?.length ? (
                        <span className="inline-flex flex-wrap gap-1 align-middle">
                          {g.cadence.map((c) => <Badge key={c} color="#A8791F">{c}</Badge>)}
                        </span>
                      ) : <span style={{ color: "#1C2624" }}>—</span>}
                    </div>
                    <div><span style={{ color: "#8A8F87" }}>Start: </span><span style={{ color: "#1C2624" }}>{fmtDate(g.start)}</span></div>
                    <div><span style={{ color: "#8A8F87" }}>Budget period: </span><span style={{ color: "#1C2624" }}>{g.budgetPeriodStart || g.budgetPeriodEnd ? `${fmtDate(g.budgetPeriodStart)} – ${fmtDate(g.budgetPeriodEnd)}` : "—"}</span></div>
                    <div><span style={{ color: "#8A8F87" }}>Obligated funds: </span><span style={{ color: "#1C2624", fontVariantNumeric: "tabular-nums" }}>{fmt(g.obligatedFunds)}</span></div>
                    <div><span style={{ color: "#8A8F87" }}>Obligated funds remaining: </span><span style={{ color: "#1C2624", fontVariantNumeric: "tabular-nums" }}>{fmt(g.obligatedFundsRemaining)}</span></div>
                    <div><span style={{ color: "#8A8F87" }}>Award amount remaining (manual): </span><span style={{ color: "#1C2624", fontVariantNumeric: "tabular-nums" }}>{fmt(g.awardAmountRemaining)}</span></div>
                    <div><span style={{ color: "#8A8F87" }}>Payment method: </span><span style={{ color: "#1C2624" }}>{g.paymentMethod || "—"}</span></div>
                    <div><span style={{ color: "#8A8F87" }}>Grant POC: </span><span style={{ color: "#1C2624" }}>{g.grantPoc || "—"}</span></div>
                    <div><span style={{ color: "#8A8F87" }}>Beds: </span><span style={{ color: "#1C2624", fontVariantNumeric: "tabular-nums" }}>{g.beds || "—"}</span></div>
                    <div><span style={{ color: "#8A8F87" }}>Bed rate: </span><span style={{ color: "#1C2624", fontVariantNumeric: "tabular-nums" }}>{g.bedRate ? fmt(g.bedRate) : "—"}</span></div>
                    <div><span style={{ color: "#8A8F87" }}>Compliance owner: </span><span style={{ color: "#1C2624" }}>{g.complianceOwner || "—"}</span></div>
                    <div><span style={{ color: "#8A8F87" }}>Finance owner: </span><span style={{ color: "#1C2624" }}>{g.financeOwner || "—"}</span></div>
                    <div><span style={{ color: "#8A8F87" }}>Internal owner: </span><span style={{ color: "#1C2624" }}>{g.internalOwner || "—"}</span></div>
                    <div><span style={{ color: "#8A8F87" }}>Operations owner: </span><span style={{ color: "#1C2624" }}>{g.operationsOwner || "—"}</span></div>
                    {g.notes && <div className="col-span-2"><span style={{ color: "#8A8F87" }}>Notes: </span><span style={{ color: "#1C2624" }}>{g.notes}</span></div>}
                    <div className="col-span-2 flex items-center gap-2 pt-2 flex-wrap">
                      <button onClick={() => goTo("budgets", null, g.id)} className="inline-flex items-center gap-1 text-xs px-3 py-1.5 rounded-md border" style={{ borderColor: "#E1E5DE", color: "#1F5C6B" }}>
                        <Wallet size={13} /> Budget
                      </button>
                      <button onClick={() => goTo("grant-reports", null, g.id)} className="inline-flex items-center gap-1 text-xs px-3 py-1.5 rounded-md border" style={{ borderColor: "#E1E5DE", color: "#1F5C6B" }}>
                        <ClipboardList size={13} /> Reports
                      </button>
                      {g.doclibUrl && (
                        <a href={g.doclibUrl} target="_blank" rel="noreferrer" className="inline-flex items-center gap-1 text-xs px-3 py-1.5 rounded-md border" style={{ borderColor: "#E1E5DE", color: "#1F5C6B" }}>
                          <ExternalLink size={13} /> Document library
                        </a>
                      )}
                      {g.contractUrl && (
                        <a href={g.contractUrl} target="_blank" rel="noreferrer" className="inline-flex items-center gap-1 text-xs px-3 py-1.5 rounded-md border" style={{ borderColor: "#E1E5DE", color: "#1F5C6B" }}>
                          <ExternalLink size={13} /> Current contract
                        </a>
                      )}
                      {canEdit && (
                        <>
                          <button onClick={() => setModal(g)} className="inline-flex items-center gap-1 text-xs px-3 py-1.5 rounded-md border" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>
                            <Pencil size={13} /> Edit grant
                          </button>
                          <button onClick={() => setConfirm(g.id)} className="inline-flex items-center gap-1 text-xs px-3 py-1.5 rounded-md border" style={{ borderColor: "#E1E5DE", color: "#B5443A" }}>
                            <Trash2 size={13} /> Delete grant
                          </button>
                        </>
                      )}
                    </div>
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}

      {modal && <GrantModal grant={modal === "new" ? null : modal} budgetGroups={budgetGroups} setBudgetGroups={setBudgetGroups} logActivity={logActivity} canEdit={canEdit} onSave={saveGrant} onClose={() => setModal(null)} />}
      {confirm && (
        <ConfirmModal
          message="This moves the grant, all of its budgets, invoices, linked grant reports and tasks to Trash, and removes it from any staff allocations. It can be restored from Trash later if needed."
          onConfirm={() => deleteGrant(confirm)}
          onCancel={() => setConfirm(null)}
        />
      )}
    </div>
  );
}

function BudgetsView({ grants, budgets, setBudgets, selectedGrantId, setSelectedGrantId, costCenters, setCostCenters, selectedCostCenterId, setSelectedCostCenterId, budgetGroups, setBudgetGroups, setTrash, currentUserEmail, canEdit, initialOpenBudgetId, logActivity }) {
  const [modal, setModal] = useState(() => (initialOpenBudgetId ? budgets.find((b) => b.id === stripNonce(initialOpenBudgetId)) || null : null));
  const [confirm, setConfirm] = useState(null);
  const [ccModal, setCcModal] = useState(null); // null | "new" | costCenter object
  const [budgetMode, setBudgetMode] = useState("grant"); // grant | costCenter

  const grant = grants.find((g) => g.id === selectedGrantId);
  const costCenter = costCenters.find((c) => c.id === selectedCostCenterId);
  const activeSelection = budgetMode === "grant" ? grant : costCenter;
  const myBudgets = budgets.filter((b) =>
    budgetMode === "grant" ? b.grantId === selectedGrantId && !b.costCenterId : b.costCenterId === selectedCostCenterId && !b.grantId
  );

  const saveCostCenter = (cc) => {
    setCostCenters((prev) => {
      const exists = prev.some((x) => x.id === cc.id);
      logActivity?.("Cost Center", exists ? "Updated" : "Created", cc.name || "Untitled cost center");
      return exists ? prev.map((x) => (x.id === cc.id ? cc : x)) : [...prev, cc];
    });
    setSelectedCostCenterId(cc.id);
    setCcModal(null);
  };
  const deleteCostCenter = (id) => {
    const cc = costCenters.find((x) => x.id === id);
    const cascadedBudgets = budgets.filter((b) => b.costCenterId === id);
    pushTrash(setTrash, "costCenter", cc, currentUserEmail, { budgets: cascadedBudgets });
    setCostCenters((prev) => prev.filter((x) => x.id !== id));
    setBudgets((prev) => prev.filter((b) => b.costCenterId !== id));
    logActivity?.("Cost Center", "Deleted", cc?.name || "Untitled cost center");
    if (selectedCostCenterId === id) setSelectedCostCenterId("");
    setCcModal(null);
  };

  const saveBudget = (b) => {
    setBudgets((prev) => {
      const exists = prev.some((x) => x.id === b.id);
      logActivity?.("Budget", exists ? "Updated" : "Created", `${b.title || "Untitled budget"}${activeSelection ? ` (${activeSelection.title || activeSelection.name})` : ""}`);
      return exists ? prev.map((x) => (x.id === b.id ? b : x)) : [...prev, b];
    });
    setModal(null);
  };
  const deleteBudget = (id) => {
    const b = budgets.find((x) => x.id === id);
    pushTrash(setTrash, "budget", b, currentUserEmail);
    setBudgets((prev) => prev.filter((b) => b.id !== id));
    logActivity?.("Budget", "Deleted", `${b?.title || "Untitled budget"}${activeSelection ? ` (${activeSelection.title || activeSelection.name})` : ""}`);
    setConfirm(null);
  };
  const nextFyLabel = (fy) => {
    if (!fy) return "";
    const match = fy.match(/(\d+)$/);
    if (match) return fy.slice(0, match.index) + (Number(match[1]) + 1);
    return `${fy} (next)`;
  };
  const shiftYear = (dateStr) => {
    if (!dateStr) return "";
    const d = new Date(dateStr + "T00:00:00");
    d.setFullYear(d.getFullYear() + 1);
    return d.toISOString().slice(0, 10);
  };
  const duplicateBudget = (budget) => {
    const newBudget = {
      ...budget,
      id: uid(),
      fy: nextFyLabel(budget.fy),
      periodStart: shiftYear(budget.periodStart),
      periodEnd: shiftYear(budget.periodEnd),
      status: "Draft",
      approvedBy: "", approvedAt: "", rejectionReason: "",
      lines: budget.lines.map((l) => ({
        ...l, id: uid(), amounts: Array(12).fill(0), actuals: Array(12).fill(0),
      })),
    };
    setBudgets((prev) => [...prev, newBudget]);
    logActivity?.("Budget", "Created", `${newBudget.title || "Untitled budget"}${activeSelection ? ` (${activeSelection.title || activeSelection.name})` : ""} — rolled over from ${budget.fy || "prior year"}`);
    setModal(newBudget);
  };
  const exportCsv = (budget) => {
    const labels = monthColumnsForBudget(budget.periodStart).map((c) => c.label);
    const rows = [["Category", "Subcategory", "Description", "Type", ...labels, "Total"]];
    budget.lines.forEach((l) => {
      rows.push([l.category, l.subcategory, l.description || "", l.type, ...l.amounts, lineTotal(l)]);
    });
    const csv = rows.map((r) => r.map((c) => `"${String(c).replace(/"/g, '""')}"`).join(",")).join("\n");
    downloadFile("nations-finest-budget-lines.csv", csv, "text/csv");
  };
  const exportXlsx = (budget) => {
    const g = grants.find((x) => x.id === budget.grantId);
    const cc = costCenters.find((x) => x.id === budget.costCenterId);
    const label = g ? (g.programCode ? `${g.programCode} - ${g.title}` : g.title) : cc ? cc.name : "Budget";
    const t = budgetTotals(budget);
    const labels = monthColumnsForBudget(budget.periodStart).map((c) => c.label);
    const rows = [
      [label],
      [`${budget.title}${budget.fy ? ` (${budget.fy})` : ""}`],
      [`Period: ${fmtDate(budget.periodStart)} – ${fmtDate(budget.periodEnd)}`, `Status: ${budget.status}`],
      [],
      ["Category", "Subcategory", "Description", "Type", ...labels, "Total"],
      ...budget.lines.map((l) => [l.category, l.subcategory || "", l.description || "", l.type, ...l.amounts, lineTotal(l)]),
      [],
      ["Total Revenue", "", "", "", ...Array(12).fill(""), t.revenue],
      ["Total Expense", "", "", "", ...Array(12).fill(""), t.expense],
      ["Net", "", "", "", ...Array(12).fill(""), t.net],
    ];
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"] = [{ wch: 30 }, { wch: 18 }, { wch: 24 }, { wch: 10 }, ...MONTHS.map(() => ({ wch: 12 })), { wch: 14 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Budget");
    const safe = (s) => (s || "budget").replace(/[^a-z0-9]+/gi, "_").slice(0, 40);
    const arrayBuffer = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    downloadFile(`${safe(g?.title || cc?.name)}-${safe(budget.title)}.xlsx`, arrayBuffer, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  };

  return (
    <div className="space-y-4">
      <div className="flex items-center justify-between">
        <h1 className="font-display text-2xl" style={{ color: "#1C2624" }}>Budgets</h1>
        {activeSelection && canEdit && (
          <button onClick={() => setModal("new")} className="inline-flex items-center gap-1.5 px-3 py-2 rounded-md text-sm text-white" style={{ background: "#1F5C6B" }}>
            <Plus size={16} /> New budget
          </button>
        )}
      </div>

      <div className="inline-flex rounded-md border overflow-hidden" style={{ borderColor: "#E1E5DE" }}>
        <button
          onClick={() => setBudgetMode("grant")}
          className="px-3 py-2 text-sm font-medium"
          style={{ background: budgetMode === "grant" ? "#1F5C6B" : "#FFFFFF", color: budgetMode === "grant" ? "#FFFFFF" : "#5B6B66" }}
        >
          Grants
        </button>
        <button
          onClick={() => setBudgetMode("costCenter")}
          className="px-3 py-2 text-sm font-medium"
          style={{ background: budgetMode === "costCenter" ? "#1F5C6B" : "#FFFFFF", color: budgetMode === "costCenter" ? "#FFFFFF" : "#5B6B66" }}
        >
          Cost Centers
        </button>
      </div>

      {budgetMode === "grant" ? (
        <Field label="Select a grant to manage budgets">
          <GrantPicker grants={grants} value={selectedGrantId} onChange={setSelectedGrantId} noneLabel="Select a grant" wrapStyle={{ maxWidth: 400 }} />
        </Field>
      ) : (
        <Field label="Select a cost center to manage budgets">
          <div className="flex items-center gap-2" style={{ maxWidth: 500 }}>
            <select
              value={selectedCostCenterId}
              onChange={(e) => setSelectedCostCenterId(e.target.value)}
              className={inputCls}
              style={{ ...inputStyle, maxWidth: 320 }}
            >
              <option value="">Select a cost center</option>
              {costCenters.map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
            </select>
            {canEdit && (
              <button onClick={() => setCcModal("new")} className="inline-flex items-center gap-1 text-xs px-3 py-2 rounded-md border shrink-0" style={{ borderColor: "#E1E5DE", color: "#1F5C6B" }}>
                <Plus size={13} /> New
              </button>
            )}
            {costCenter && canEdit && (
              <button onClick={() => setCcModal(costCenter)} className="inline-flex items-center gap-1 text-xs px-3 py-2 rounded-md border shrink-0" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>
                <Pencil size={13} /> Edit
              </button>
            )}
          </div>
        </Field>
      )}

      {!activeSelection ? (
        <div className="bg-white rounded-lg border p-10 text-center" style={{ borderColor: "#E1E5DE", color: "#8A8F87" }}>
          {budgetMode === "grant" ? "Select a grant to view its details." : "Select or create a cost center to view its budgets."}
        </div>
      ) : myBudgets.length === 0 ? (
        <div className="bg-white rounded-lg border p-10 text-center space-y-3" style={{ borderColor: "#E1E5DE", color: "#8A8F87" }}>
          <p>No budgets for this {budgetMode === "grant" ? "grant" : "cost center"} yet.</p>
          <button onClick={() => setModal("new")} className="inline-flex items-center gap-1.5 px-3 py-2 rounded-md text-sm text-white" style={{ background: "#1F5C6B" }}>
            <Plus size={16} /> Create the first budget
          </button>
        </div>
      ) : (
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          {myBudgets.map((b) => {
            const t = budgetTotals(b);
            return (
              <div key={b.id} className="bg-white rounded-lg border p-4" style={{ borderColor: "#E1E5DE" }}>
                <div className="flex items-center justify-between">
                  <div>
                    <div className="font-medium" style={{ color: "#1C2624" }}>{b.title}</div>
                    <div className="text-xs mt-0.5" style={{ color: "#8A8F87" }}>{b.fy} · {fmtDate(b.periodStart)} – {fmtDate(b.periodEnd)}</div>
                  </div>
                  <Badge color={b.status === "Active" ? "#2F6F53" : b.status === "Closed" ? "#8A8F87" : b.status === "Rejected" ? "#B5443A" : b.status === "Pending Approval" ? "#C08A2E" : "#5B7FA6"}>{b.status}</Badge>
                </div>
                <div className="flex gap-5 mt-3 text-sm" style={{ fontVariantNumeric: "tabular-nums" }}>
                  <div><span style={{ color: "#8A8F87" }}>Revenue </span><span style={{ color: "#2F6F53" }}>{fmt(t.revenue)}</span></div>
                  <div><span style={{ color: "#8A8F87" }}>Expense </span><span style={{ color: "#1C2624" }}>{fmt(t.expense)}</span></div>
                  <div><span style={{ color: "#8A8F87" }}>Net </span><span style={{ color: t.net >= 0 ? "#2F6F53" : "#B5443A" }}>{fmt(t.net)}</span></div>
                </div>
                <div className="flex gap-2 mt-4 flex-wrap">
                  <button onClick={() => setModal(b)} className="inline-flex items-center gap-1 text-xs px-3 py-1.5 rounded-md border" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>
                    <Pencil size={13} /> {canEdit ? "Edit budget" : "View budget"}
                  </button>
                  <button onClick={() => exportCsv(b)} className="inline-flex items-center gap-1 text-xs px-3 py-1.5 rounded-md border" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>
                    <Download size={13} /> Export CSV
                  </button>
                  <button onClick={() => exportXlsx(b)} className="inline-flex items-center gap-1 text-xs px-3 py-1.5 rounded-md border" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>
                    <Download size={13} /> Export Excel
                  </button>
                  {canEdit && (
                    <>
                      <button onClick={() => duplicateBudget(b)} className="inline-flex items-center gap-1 text-xs px-3 py-1.5 rounded-md border" style={{ borderColor: "#E1E5DE", color: "#1F5C6B" }}>
                        <Plus size={13} /> Duplicate to next FY
                      </button>
                      <button onClick={() => setConfirm(b.id)} className="inline-flex items-center gap-1 text-xs px-3 py-1.5 rounded-md border" style={{ borderColor: "#E1E5DE", color: "#B5443A" }}>
                        <Trash2 size={13} /> Delete
                      </button>
                    </>
                  )}
                </div>
              </div>
            );
          })}
        </div>
      )}

      {modal && (
        <BudgetModal
          budget={modal === "new" ? null : modal}
          grantId={budgetMode === "grant" ? selectedGrantId : ""}
          costCenterId={budgetMode === "costCenter" ? selectedCostCenterId : ""}
          canEdit={canEdit}
          onSave={saveBudget}
          onClose={() => setModal(null)}
        />
      )}
      {confirm && (
        <ConfirmModal message="This will permanently delete this budget." onConfirm={() => deleteBudget(confirm)} onCancel={() => setConfirm(null)} />
      )}
      {ccModal && (
        <CostCenterModal
          costCenter={ccModal === "new" ? null : ccModal}
          budgetGroups={budgetGroups}
          setBudgetGroups={setBudgetGroups}
          logActivity={logActivity}
          onSave={saveCostCenter}
          onClose={() => setCcModal(null)}
          onDelete={ccModal === "new" ? undefined : () => deleteCostCenter(ccModal.id)}
        />
      )}
    </div>
  );
}

function ReportCard({ report, grant, onToggleDone, onBucketChange, onEdit }) {
  const progress = checklistProgress(report);
  const overdue = isOverdue(report);
  return (
    <div className="bg-white rounded-lg border p-3.5" style={{ borderColor: overdue ? "#B5443A" : "#E1E5DE" }}>
      <div className="flex items-start gap-2">
        <button onClick={onToggleDone} disabled={!onToggleDone} className="mt-0.5 shrink-0">
          {report.status === "Completed" ? <CheckCircle size={17} style={{ color: "#2F6F53" }} /> : <Circle size={17} style={{ color: "#8A8F87" }} />}
        </button>
        <button onClick={onEdit} className="text-left flex-1">
          <div className="text-sm font-medium" style={{ color: report.status === "Completed" ? "#8A8F87" : "#1C2624", textDecoration: report.status === "Completed" ? "line-through" : "none" }}>
            {report.title}
          </div>
        </button>
        <span className="w-2 h-2 rounded-full mt-1.5 shrink-0" style={{ background: priorityColor(report.priority) }} title={report.priority} />
      </div>

      {grant && (
        <div className="mt-2">
          <Badge color="#B5443A">{grant.programCode ? `${grant.programCode} - ${grant.title}` : grant.title}</Badge>
        </div>
      )}

      <div className="flex items-center justify-between mt-2.5 text-xs" style={{ color: "#8A8F87" }}>
        <span>{report.assignedTo || "Unassigned"}</span>
        <span style={{ color: overdue ? "#B5443A" : "#8A8F87", fontVariantNumeric: "tabular-nums" }}>
          {report.dueDate ? `Due ${fmtDate(report.dueDate)}` : "No due date"}
        </span>
      </div>

      {progress.total > 0 && (
        <div className="text-xs mt-1" style={{ color: "#8A8F87" }}>{progress.done}/{progress.total} checklist steps</div>
      )}

      {report.portalUrl && (
        <a
          href={report.portalUrl} target="_blank" rel="noreferrer"
          onClick={(e) => e.stopPropagation()}
          className="inline-flex items-center gap-1 text-xs px-2 py-1 rounded-md border mt-2.5"
          style={{ borderColor: "#E1E5DE", color: "#1F5C6B" }}
        >
          <ExternalLink size={12} /> Submission portal
        </a>
      )}

      <select
        value={report.bucket}
        onChange={(e) => onBucketChange?.(e.target.value)}
        onClick={(e) => e.stopPropagation()}
        disabled={!onBucketChange}
        className="w-full mt-2.5 rounded border px-2 py-1 text-xs"
        style={inputStyle}
      >
        {[...new Set([...DEFAULT_BUCKETS, report.bucket])].map((b) => <option key={b} value={b}>{b}</option>)}
      </select>
    </div>
  );
}

function ReportsView({ grants, reports, setReports, setTasks, grantFilter, setGrantFilter, setTrash, currentUserEmail, canEdit, initialOpenReportId, logActivity }) {
  const [modal, setModal] = useState(() => (initialOpenReportId ? reports.find((r) => r.id === stripNonce(initialOpenReportId)) || null : null));
  const [confirm, setConfirm] = useState(null);

  const buckets = [...new Set([...DEFAULT_BUCKETS, ...reports.map((r) => r.bucket)])];
  const visible = grantFilter === "All" ? reports : reports.filter((r) => r.grantId === grantFilter);

  const saveReport = (r) => {
    setReports((prev) => {
      const exists = prev.some((x) => x.id === r.id);
      logActivity?.("Report", exists ? "Updated" : "Created", r.title || "Untitled report");
      return exists ? prev.map((x) => (x.id === r.id ? r : x)) : [...prev, r];
    });
    setModal(null);
  };
  const createTaskFromReport = (r) => {
    const newTask = {
      id: uid(), title: `Complete report: ${r.title}`, category: "Report Submission",
      grantId: r.grantId, dueDate: r.dueDate || "", priority: r.priority || "Medium",
      status: "Not started", assignedTo: r.assignedTo || "",
      notes: `Linked to the grant report "${r.title}"${r.portalUrl ? ` — submission portal: ${r.portalUrl}` : ""}.`,
    };
    setTasks?.((prev) => [...prev, newTask]);
    logActivity?.("Task", "Created", newTask.title);
  };
  const deleteReport = (id) => {
    const r = reports.find((x) => x.id === id);
    pushTrash(setTrash, "report", r, currentUserEmail);
    setReports((prev) => prev.filter((r) => r.id !== id));
    logActivity?.("Report", "Deleted", r?.title || "Untitled report");
    setConfirm(null);
  };
  const toggleDone = (r) => {
    setReports((prev) => prev.map((x) => (x.id === r.id ? { ...x, status: x.status === "Completed" ? "Not started" : "Completed" } : x)));
  };
  const changeBucket = (r, bucket) => {
    setReports((prev) => prev.map((x) => (x.id === r.id ? { ...x, bucket } : x)));
  };

  return (
    <div className="space-y-4">
      <div className="flex items-center justify-between">
        <div>
          <h1 className="font-display text-2xl" style={{ color: "#1C2624" }}>Grant Reports</h1>
          <p className="text-sm mt-1" style={{ color: "#5B6B66" }}>Track reports and deliverables due to funders</p>
        </div>
        {canEdit && (
          <button onClick={() => setModal("new")} className="inline-flex items-center gap-1.5 px-3 py-2 rounded-md text-sm text-white" style={{ background: "#1F5C6B" }}>
            <Plus size={16} /> New report
          </button>
        )}
      </div>

      <Field label="Filter by grant">
        <GrantPicker grants={grants} value={grantFilter === "All" ? "" : grantFilter} onChange={(v) => setGrantFilter(v || "All")} noneLabel="All grants" noneValue="All" wrapStyle={{ maxWidth: 320 }} />
      </Field>

      {visible.length === 0 ? (
        <div className="bg-white rounded-lg border p-10 text-center" style={{ borderColor: "#E1E5DE", color: "#8A8F87" }}>
          No grant reports yet.
        </div>
      ) : (
        <div className="overflow-x-auto pb-2">
          <div className="grid gap-4" style={{ gridTemplateColumns: `repeat(${buckets.length}, minmax(240px, 1fr))`, minWidth: buckets.length * 250 }}>
            {buckets.map((bucket) => (
              <div key={bucket}>
                <h3 className="text-xs font-semibold uppercase tracking-wide mb-2" style={{ color: "#5B6B66" }}>
                  {bucket} <span style={{ color: "#8A8F87", fontWeight: 400 }}>({visible.filter((r) => r.bucket === bucket).length})</span>
                </h3>
                <div className="space-y-3">
                  {visible.filter((r) => r.bucket === bucket).map((r) => (
                    <ReportCard
                      key={r.id}
                      report={r}
                      grant={grants.find((g) => g.id === r.grantId)}
                      onToggleDone={canEdit ? () => toggleDone(r) : undefined}
                      onBucketChange={canEdit ? (b) => changeBucket(r, b) : undefined}
                      onEdit={() => setModal(r)}
                    />
                  ))}
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {modal && (
        <ReportModal
          report={modal === "new" ? null : modal}
          grants={grants}
          canEdit={canEdit}
          onSave={saveReport}
          onClose={() => setModal(null)}
          onDelete={modal === "new" || !canEdit ? undefined : () => { setConfirm(modal.id); setModal(null); }}
          onCreateTask={createTaskFromReport}
        />
      )}
      {confirm && (
        <ConfirmModal message="This will permanently delete this report." onConfirm={() => deleteReport(confirm)} onCancel={() => setConfirm(null)} />
      )}
    </div>
  );
}

// ---------- tasks / reminders ----------

function TaskModal({ task, grants, canEdit = true, onSave, onClose, onDelete }) {
  const [form, setForm] = useState(task || {
    id: uid(), title: "", category: TASK_CATEGORIES[0], grantId: "", dueDate: "",
    priority: "Medium", status: "Not started", assignedTo: "", notes: "",
  });
  const set = (k) => (e) => setForm({ ...form, [k]: e.target.value });

  return (
    <Modal title={task ? (canEdit ? "Edit task" : "View task") : "New task"} onClose={onClose}>
      <fieldset disabled={!canEdit} style={{ border: "none", margin: 0, padding: 0 }}>
      <div className="space-y-4">
        <Field label="Task title">
          <input className={inputCls} style={inputStyle} value={form.title} onChange={set("title")} placeholder="e.g. Submit LOI to funder" />
        </Field>
        <Field label="Category">
          <select className={inputCls} style={inputStyle} value={form.category} onChange={set("category")}>
            {TASK_CATEGORIES.map((c) => <option key={c}>{c}</option>)}
          </select>
        </Field>
        <Field label="Grant">
          <GrantPicker grants={grants} value={form.grantId} onChange={(v) => setForm({ ...form, grantId: v })} noneLabel="No grant linked" />
        </Field>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          <Field label="Due date">
            <input type="date" className={inputCls} style={inputStyle} value={form.dueDate} onChange={set("dueDate")} />
          </Field>
          <Field label="Priority">
            <select className={inputCls} style={inputStyle} value={form.priority} onChange={set("priority")}>
              {REPORT_PRIORITIES.map((p) => <option key={p.label}>{p.label}</option>)}
            </select>
          </Field>
        </div>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          <Field label="Status">
            <select className={inputCls} style={inputStyle} value={form.status} onChange={set("status")}>
              {TASK_STATUSES.map((s) => <option key={s}>{s}</option>)}
            </select>
          </Field>
          <Field label="Assigned to">
            <input className={inputCls} style={inputStyle} value={form.assignedTo} onChange={set("assignedTo")} placeholder="Name" />
          </Field>
        </div>
        <Field label="Notes">
          <textarea className={inputCls} style={inputStyle} rows={3} value={form.notes} onChange={set("notes")} />
        </Field>
      </div>
      </fieldset>

      <div className="flex justify-between gap-2 mt-6">
        {onDelete ? (
          <button onClick={onDelete} className="px-4 py-2 rounded-md text-sm border" style={{ borderColor: "#E1E5DE", color: "#B5443A" }}>Delete task</button>
        ) : <span />}
        <div className="flex gap-2">
          <button onClick={onClose} className="px-4 py-2 rounded-md text-sm border" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>Cancel</button>
          {canEdit && (
            <button
              onClick={() => { if (!form.title.trim()) return; onSave(form); }}
              className="px-4 py-2 rounded-md text-sm text-white"
              style={{ background: "#1F5C6B" }}
            >
              Save task
            </button>
          )}
        </div>
      </div>
    </Modal>
  );
}

function TasksView({ grants, tasks, setTasks, setTrash, currentUserEmail, canEdit, autoOpenNew, initialOpenTaskId, logActivity }) {
  const [modal, setModal] = useState(() => (autoOpenNew ? "new" : initialOpenTaskId ? tasks.find((t) => t.id === stripNonce(initialOpenTaskId)) || null : null));
  const [confirm, setConfirm] = useState(null);
  const [categoryFilter, setCategoryFilter] = useState("All");
  const [statusFilter, setStatusFilter] = useState("All");
  const [grantFilter, setGrantFilter] = useState("All");

  const visible = tasks
    .filter((t) => categoryFilter === "All" || t.category === categoryFilter)
    .filter((t) => statusFilter === "All" || t.status === statusFilter)
    .filter((t) => grantFilter === "All" || t.grantId === grantFilter)
    .sort((a, b) => {
      if (!a.dueDate && !b.dueDate) return 0;
      if (!a.dueDate) return 1;
      if (!b.dueDate) return -1;
      return new Date(a.dueDate) - new Date(b.dueDate);
    });

  const saveTask = (t) => {
    setTasks((prev) => {
      const exists = prev.some((x) => x.id === t.id);
      logActivity?.("Task", exists ? "Updated" : "Created", t.title || "Untitled task");
      return exists ? prev.map((x) => (x.id === t.id ? t : x)) : [...prev, t];
    });
    setModal(null);
  };
  const deleteTask = (id) => {
    const t = tasks.find((x) => x.id === id);
    pushTrash(setTrash, "task", t, currentUserEmail);
    setTasks((prev) => prev.filter((t) => t.id !== id));
    logActivity?.("Task", "Deleted", t?.title || "Untitled task");
    setConfirm(null);
  };
  const toggleDone = (t) => {
    setTasks((prev) => prev.map((x) => (x.id === t.id ? { ...x, status: x.status === "Done" ? "Not started" : "Done" } : x)));
  };
  const taskOverdue = (t) => t.dueDate && t.status !== "Done" && new Date(t.dueDate) < new Date(new Date().toDateString());

  return (
    <div className="space-y-4">
      <div className="flex items-center justify-between">
        <div>
          <h1 className="font-display text-2xl" style={{ color: "#1C2624" }}>Tasks</h1>
          <p className="text-sm mt-1" style={{ color: "#5B6B66" }}>Deadlines and to-dos beyond funder reports — site visits, renewals, approvals, and more</p>
        </div>
        {canEdit && (
          <button onClick={() => setModal("new")} className="inline-flex items-center gap-1.5 px-3 py-2 rounded-md text-sm text-white" style={{ background: "#1F5C6B" }}>
            <Plus size={16} /> New task
          </button>
        )}
      </div>

      <div className="flex gap-3">
        <select value={categoryFilter} onChange={(e) => setCategoryFilter(e.target.value)} className={inputCls} style={{ ...inputStyle, width: 200 }}>
          <option>All</option>
          {TASK_CATEGORIES.map((c) => <option key={c}>{c}</option>)}
        </select>
        <select value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)} className={inputCls} style={{ ...inputStyle, width: 160 }}>
          <option>All</option>
          {TASK_STATUSES.map((s) => <option key={s}>{s}</option>)}
        </select>
        <GrantPicker grants={grants} value={grantFilter === "All" ? "" : grantFilter} onChange={(v) => setGrantFilter(v || "All")} noneLabel="All grants" noneValue="All" wrapStyle={{ width: 220 }} />
      </div>

      {visible.length === 0 ? (
        <div className="bg-white rounded-lg border p-10 text-center" style={{ borderColor: "#E1E5DE", color: "#8A8F87" }}>
          No tasks match your filters.
        </div>
      ) : (
        <div className="bg-white rounded-lg border divide-y" style={{ borderColor: "#E1E5DE" }}>
          {visible.map((t) => {
            const g = grants.find((x) => x.id === t.grantId);
            const overdue = taskOverdue(t);
            return (
              <div key={t.id} className="px-4 py-3 flex items-start gap-3">
                <button onClick={() => toggleDone(t)} disabled={!canEdit} className="mt-0.5 shrink-0">
                  {t.status === "Done" ? <CheckCircle size={17} style={{ color: "#2F6F53" }} /> : <Circle size={17} style={{ color: "#8A8F87" }} />}
                </button>
                <button onClick={() => setModal(t)} className="flex-1 text-left">
                  <div className="text-sm font-medium" style={{ color: t.status === "Done" ? "#8A8F87" : "#1C2624", textDecoration: t.status === "Done" ? "line-through" : "none" }}>
                    {t.title}
                  </div>
                  <div className="flex items-center gap-2 mt-1 flex-wrap">
                    <Badge color="#5B6B66">{t.category}</Badge>
                    {g && <Badge color="#5B7FA6">{g.programCode ? `${g.programCode} - ${g.title}` : g.title}</Badge>}
                    {t.assignedTo && <span className="text-xs" style={{ color: "#8A8F87" }}>{t.assignedTo}</span>}
                  </div>
                </button>
                <div className="text-right shrink-0">
                  <div className="text-sm" style={{ color: overdue ? "#B5443A" : "#1C2624", fontVariantNumeric: "tabular-nums" }}>
                    {t.dueDate ? `${overdue ? "Overdue: " : ""}${fmtDate(t.dueDate)}` : "No due date"}
                  </div>
                  <span className="w-2 h-2 rounded-full inline-block mt-1" style={{ background: priorityColor(t.priority) }} title={t.priority} />
                </div>
                <button onClick={() => setConfirm(t.id)} className="p-1 rounded hover:bg-red-50 shrink-0">
                  <Trash2 size={14} style={{ color: "#B5443A" }} />
                </button>
              </div>
            );
          })}
        </div>
      )}

      {modal && (
        <TaskModal
          task={modal === "new" ? null : modal}
          grants={grants}
          canEdit={canEdit}
          onSave={saveTask}
          onClose={() => setModal(null)}
          onDelete={modal === "new" || !canEdit ? undefined : () => { setConfirm(modal.id); setModal(null); }}
        />
      )}
      {confirm && (
        <ConfirmModal message="This will permanently delete this task." onConfirm={() => deleteTask(confirm)} onCancel={() => setConfirm(null)} />
      )}
    </div>
  );
}

function ReportingView({ grants, budgets }) {
  const [scope, setScope] = useState("all");
  const relevantBudgets = scope === "all" ? budgets : budgets.filter((b) => b.grantId === scope);

  const byCategory = useMemo(() => {
    const map = {};
    relevantBudgets.forEach((b) => b.lines.forEach((l) => {
      if (l.type !== "expense") return;
      map[l.category] = (map[l.category] || 0) + lineTotal(l);
    }));
    return Object.entries(map).map(([category, total]) => ({ category, total })).sort((a, b) => b.total - a.total);
  }, [relevantBudgets]);

  const byMonth = useMemo(() => {
    const revenue = Array(12).fill(0), expense = Array(12).fill(0);
    relevantBudgets.forEach((b) => b.lines.forEach((l) => l.amounts.forEach((a, i) => {
      if (l.type === "revenue") revenue[i] += Number(a) || 0; else expense[i] += Number(a) || 0;
    })));
    return MONTHS.map((m, i) => ({ month: m, revenue: revenue[i], expense: expense[i] }));
  }, [relevantBudgets]);

  const byGrant = useMemo(() => {
    return grants.map((g) => {
      const t = grantBudgetTotals(g.id, budgets);
      return { grant: g.title, revenue: t.revenue, expense: t.expense };
    }).filter((r) => r.revenue || r.expense);
  }, [grants, budgets]);

  const byProgram = useMemo(() => {
    const map = {};
    grants.forEach((g) => {
      const t = grantBudgetTotals(g.id, budgets);
      const key = g.programCode || "Unassigned";
      if (!map[key]) map[key] = { program: key, revenue: 0, expense: 0 };
      map[key].revenue += t.revenue; map[key].expense += t.expense;
    });
    return Object.values(map);
  }, [grants, budgets]);

  const totals = relevantBudgets.reduce((acc, b) => {
    const t = budgetTotals(b);
    acc.revenue += t.revenue; acc.expense += t.expense;
    return acc;
  }, { revenue: 0, expense: 0 });

  const exportAllCsv = () => {
    const monthCols = MONTHS.map((_, i) => `Month ${i + 1}`);
    const rows = [["Grant", "Budget", "Period Start", "Category", "Subcategory", "Description", "Type", ...monthCols, "Total"]];
    budgets.forEach((b) => {
      const g = grants.find((x) => x.id === b.grantId);
      b.lines.forEach((l) => rows.push([g?.title || "", b.title, b.periodStart || "", l.category, l.subcategory, l.description || "", l.type, ...l.amounts, lineTotal(l)]));
    });
    const csv = rows.map((r) => r.map((c) => `"${String(c).replace(/"/g, '""')}"`).join(",")).join("\n");
    downloadFile("nations-finest-budget-lines.csv", csv, "text/csv");
  };

  return (
    <div className="space-y-6">
      <div className="flex items-center justify-between">
        <h1 className="font-display text-2xl" style={{ color: "#1C2624" }}>Reporting</h1>
        <button onClick={exportAllCsv} className="inline-flex items-center gap-1.5 px-3 py-2 rounded-md text-sm border" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>
          <Download size={15} /> Export all as CSV
        </button>
      </div>

      <Field label="Scope">
        <GrantPicker grants={grants} value={scope === "all" ? "" : scope} onChange={(v) => setScope(v || "all")} noneLabel="All grants" noneValue="all" wrapStyle={{ maxWidth: 320 }} />
      </Field>

      <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
        <StatCard label="Total revenue" value={fmt(totals.revenue)} />
        <StatCard label="Total expense" value={fmt(totals.expense)} />
        <StatCard label="Total net" value={fmt(totals.revenue - totals.expense)} />
      </div>

      <div className="grid grid-cols-1 sm:grid-cols-2 gap-6">
        <div className="bg-white rounded-lg border p-5" style={{ borderColor: "#E1E5DE" }}>
          <h2 className="font-display text-base mb-3" style={{ color: "#1C2624" }}>Expense by category</h2>
          {byCategory.length === 0 ? <p className="text-sm" style={{ color: "#8A8F87" }}>No monthly data yet.</p> : (
            <ResponsiveContainer width="100%" height={260}>
              <BarChart data={byCategory} layout="vertical" margin={{ left: 10 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#E1E5DE" />
                <XAxis type="number" tickFormatter={(v) => `$${v / 1000}k`} tick={{ fontSize: 11 }} />
                <YAxis type="category" dataKey="category" width={150} tick={{ fontSize: 10 }} />
                <Tooltip formatter={(v) => fmt(v)} />
                <Bar dataKey="total" fill="#B5443A" radius={[0, 3, 3, 0]} />
              </BarChart>
            </ResponsiveContainer>
          )}
        </div>

        <div className="bg-white rounded-lg border p-5" style={{ borderColor: "#E1E5DE" }}>
          <h2 className="font-display text-base mb-3" style={{ color: "#1C2624" }}>Budget by month</h2>
          <ResponsiveContainer width="100%" height={260}>
            <LineChart data={byMonth}>
              <CartesianGrid strokeDasharray="3 3" stroke="#E1E5DE" />
              <XAxis dataKey="month" tick={{ fontSize: 11 }} />
              <YAxis tickFormatter={(v) => `$${v / 1000}k`} tick={{ fontSize: 11 }} />
              <Tooltip formatter={(v) => fmt(v)} />
              <Legend />
              <Line type="monotone" dataKey="revenue" stroke="#2F6F53" strokeWidth={2} dot={false} />
              <Line type="monotone" dataKey="expense" stroke="#B5443A" strokeWidth={2} dot={false} />
            </LineChart>
          </ResponsiveContainer>
        </div>

        <div className="bg-white rounded-lg border p-5" style={{ borderColor: "#E1E5DE" }}>
          <h2 className="font-display text-base mb-3" style={{ color: "#1C2624" }}>Budget by program</h2>
          {byProgram.length === 0 ? <p className="text-sm" style={{ color: "#8A8F87" }}>No grant budget data yet.</p> : (
            <ResponsiveContainer width="100%" height={240}>
              <BarChart data={byProgram}>
                <CartesianGrid strokeDasharray="3 3" stroke="#E1E5DE" />
                <XAxis dataKey="program" tick={{ fontSize: 10 }} />
                <YAxis tickFormatter={(v) => `$${v / 1000}k`} tick={{ fontSize: 11 }} />
                <Tooltip formatter={(v) => fmt(v)} />
                <Legend />
                <Bar dataKey="revenue" fill="#2F6F53" radius={[3, 3, 0, 0]} />
                <Bar dataKey="expense" fill="#B5443A" radius={[3, 3, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          )}
        </div>

        <div className="bg-white rounded-lg border p-5" style={{ borderColor: "#E1E5DE" }}>
          <h2 className="font-display text-base mb-3" style={{ color: "#1C2624" }}>Budget by grant</h2>
          {byGrant.length === 0 ? <p className="text-sm" style={{ color: "#8A8F87" }}>No grant budget data yet.</p> : (
            <ResponsiveContainer width="100%" height={240}>
              <BarChart data={byGrant}>
                <CartesianGrid strokeDasharray="3 3" stroke="#E1E5DE" />
                <XAxis dataKey="grant" tick={{ fontSize: 9 }} interval={0} angle={-20} textAnchor="end" height={60} />
                <YAxis tickFormatter={(v) => `$${v / 1000}k`} tick={{ fontSize: 11 }} />
                <Tooltip formatter={(v) => fmt(v)} />
                <Legend />
                <Bar dataKey="revenue" fill="#2F6F53" radius={[3, 3, 0, 0]} />
                <Bar dataKey="expense" fill="#B5443A" radius={[3, 3, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          )}
        </div>
      </div>
    </div>
  );
}

function OrgBudgetRow({ label, values, bold, indent, color, isHeader }) {
  const total = values.reduce((a, b) => a + b, 0);
  return (
    <tr className={isHeader ? "" : "border-t"} style={{ borderColor: "#E1E5DE", background: isHeader ? "#F6F7F3" : "transparent" }}>
      <td className={`px-3 py-1.5 text-xs sticky left-0 ${isHeader ? "" : "bg-white"}`} style={{ background: isHeader ? "#F6F7F3" : undefined, paddingLeft: indent ? 28 : 12, fontWeight: bold ? 600 : 400, color: color || "#1C2624" }}>
        {label}
      </td>
      {values.map((v, i) => (
        <td key={i} className="px-2 py-1.5 text-right text-xs" style={{ fontVariantNumeric: "tabular-nums", fontWeight: bold ? 600 : 400, color: color || "#1C2624" }}>
          {v ? fmt(v) : "—"}
        </td>
      ))}
      <td className="px-3 py-1.5 text-right text-xs" style={{ fontVariantNumeric: "tabular-nums", fontWeight: bold ? 700 : 500, color: color || "#1C2624" }}>
        {fmt(total)}
      </td>
    </tr>
  );
}

function NewScenarioModal({ grants, costCenters, budgets, budgetGroups, onCreate, onClose }) {
  const [title, setTitle] = useState("");
  const [startMode, setStartMode] = useState("blank"); // blank | existing | org
  const [pickMode, setPickMode] = useState("grant"); // grant | costCenter (for "existing")
  const [pickedGrantId, setPickedGrantId] = useState("");
  const [pickedCcId, setPickedCcId] = useState("");
  const [pickedBudgetId, setPickedBudgetId] = useState("");
  const [orgScope, setOrgScope] = useState("all");
  const [orgYear, setOrgYear] = useState("All");

  const candidateBudgets = pickMode === "grant"
    ? budgets.filter((b) => b.grantId === pickedGrantId)
    : budgets.filter((b) => b.costCenterId === pickedCcId);

  const orgCalendarYears = useMemo(() => {
    const scopedGrantIds = orgScope === "all" ? null : new Set(grants.filter((g) => g.budgetGroupId === orgScope).map((g) => g.id));
    const scopedCcIds = orgScope === "all" ? null : new Set((costCenters || []).filter((c) => c.budgetGroupId === orgScope).map((c) => c.id));
    const scoped = (orgScope === "all" ? budgets : budgets.filter((b) => (b.grantId && scopedGrantIds.has(b.grantId)) || (b.costCenterId && scopedCcIds.has(b.costCenterId)))).filter((b) => b.status === "Active");
    const years = new Set();
    scoped.forEach((b) => monthColumnsForBudget(b.periodStart).forEach((col) => years.add(col.year)));
    return [...years].sort();
  }, [orgScope, grants, costCenters, budgets]);

  const canCreate = title.trim() && (
    startMode === "blank" ||
    (startMode === "existing" && pickedBudgetId) ||
    startMode === "org"
  );

  const handleCreate = () => {
    let scen;
    if (startMode === "blank") {
      scen = newScenario({ type: "blank" });
    } else if (startMode === "existing") {
      const b = budgets.find((x) => x.id === pickedBudgetId);
      const g = pickMode === "grant" ? grants.find((x) => x.id === pickedGrantId) : null;
      const cc = pickMode === "costCenter" ? costCenters.find((x) => x.id === pickedCcId) : null;
      scen = newScenario({
        type: pickMode, grantId: pickedGrantId || "", costCenterId: pickedCcId || "", budgetId: pickedBudgetId,
        label: g ? (g.programCode ? `${g.programCode} - ${g.title}` : g.title) : cc ? cc.name : "",
      });
      scen.fy = b.fy;
      scen.periodStart = b.periodStart;
      scen.periodEnd = b.periodEnd;
      scen.lines = b.lines.map((l) => ({ ...l, id: uid(), amounts: [...l.amounts] }));
    } else {
      const bg = budgetGroups.find((x) => x.id === orgScope);
      scen = newScenario({ type: "org", scope: orgScope, calYear: orgYear, label: orgScope === "all" ? "Whole Organization" : (bg?.name || "Budget Group") });
      scen.periodStart = orgYear !== "All" ? `${orgYear}-01-01` : "";
      const scopedGrantIds = orgScope === "all" ? null : new Set(grants.filter((g) => g.budgetGroupId === orgScope).map((g) => g.id));
      const scopedCcIds = orgScope === "all" ? null : new Set((costCenters || []).filter((c) => c.budgetGroupId === orgScope).map((c) => c.id));
      const scoped = (orgScope === "all" ? budgets : budgets.filter((b) => (b.grantId && scopedGrantIds.has(b.grantId)) || (b.costCenterId && scopedCcIds.has(b.costCenterId)))).filter((b) => b.status === "Active");
      const map = {};
      scoped.forEach((b) => {
        const cols = monthColumnsForBudget(b.periodStart);
        b.lines.forEach((l) => {
          const key = l.category;
          if (!map[key]) {
            const catDef = CATEGORIES.find((c) => c.name === l.category);
            map[key] = { category: l.category, type: catDef ? catDef.type : l.type, subcategory: "", amounts: Array(12).fill(0) };
          }
          (l.amounts || Array(12).fill(0)).forEach((a, i) => {
            const col = cols[i];
            if (orgYear !== "All" && col.year !== orgYear) return;
            map[key].amounts[col.monthIndex] += Number(a) || 0;
          });
        });
      });
      scen.lines = Object.values(map).map((l) => ({ id: uid(), category: l.category, type: l.type, categoryCustom: false, subcategory: "", subcategoryCustom: false, amounts: l.amounts }));
      if (scen.lines.length === 0) scen.lines = [newLine()];
    }
    scen.title = title.trim();
    onCreate(scen);
  };

  return (
    <Modal title="New scenario" onClose={onClose} wide>
      <div className="space-y-4">
        <Field label="Scenario name">
          <input className={inputCls} style={inputStyle} value={title} onChange={(e) => setTitle(e.target.value)} placeholder="e.g. FY27 Conservative Case" autoFocus />
        </Field>

        <Field label="Starting point">
          <div className="grid grid-cols-3 gap-2">
            {[
              { key: "blank", label: "Start blank" },
              { key: "existing", label: "Snapshot a grant/cost center budget" },
              { key: "org", label: "Snapshot the Org Budget rollup" },
            ].map((opt) => (
              <button
                key={opt.key}
                onClick={() => setStartMode(opt.key)}
                className="text-sm px-3 py-2.5 rounded-md border text-left"
                style={{
                  borderColor: startMode === opt.key ? "#1F5C6B" : "#E1E5DE",
                  background: startMode === opt.key ? "rgba(31,92,107,0.06)" : "#FFFFFF",
                  color: "#1C2624",
                }}
              >
                {opt.label}
              </button>
            ))}
          </div>
        </Field>

        {startMode === "existing" && (
          <div className="space-y-3 rounded-md border p-4" style={{ borderColor: "#E1E5DE" }}>
            <div className="inline-flex rounded-md border overflow-hidden" style={{ borderColor: "#E1E5DE" }}>
              <button onClick={() => { setPickMode("grant"); setPickedBudgetId(""); }} className="px-3 py-1.5 text-sm font-medium" style={{ background: pickMode === "grant" ? "#1F5C6B" : "#FFFFFF", color: pickMode === "grant" ? "#FFFFFF" : "#5B6B66" }}>Grant</button>
              <button onClick={() => { setPickMode("costCenter"); setPickedBudgetId(""); }} className="px-3 py-1.5 text-sm font-medium" style={{ background: pickMode === "costCenter" ? "#1F5C6B" : "#FFFFFF", color: pickMode === "costCenter" ? "#FFFFFF" : "#5B6B66" }}>Cost Center</button>
            </div>
            {pickMode === "grant" ? (
              <GrantPicker grants={grants} value={pickedGrantId} onChange={(v) => { setPickedGrantId(v); setPickedBudgetId(""); }} noneLabel="Select a grant" />
            ) : (
              <select value={pickedCcId} onChange={(e) => { setPickedCcId(e.target.value); setPickedBudgetId(""); }} className={inputCls} style={inputStyle}>
                <option value="">Select a cost center</option>
                {(costCenters || []).map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
              </select>
            )}
            {(pickedGrantId || pickedCcId) && (
              <select value={pickedBudgetId} onChange={(e) => setPickedBudgetId(e.target.value)} className={inputCls} style={inputStyle}>
                <option value="">Select a budget to snapshot</option>
                {candidateBudgets.map((b) => <option key={b.id} value={b.id}>{b.title}{b.fy ? ` (${b.fy})` : ""}</option>)}
              </select>
            )}
          </div>
        )}

        {startMode === "org" && (
          <div className="grid grid-cols-2 gap-3 rounded-md border p-4" style={{ borderColor: "#E1E5DE" }}>
            <Field label="Scope">
              <select value={orgScope} onChange={(e) => { setOrgScope(e.target.value); setOrgYear("All"); }} className={inputCls} style={inputStyle}>
                <option value="all">Whole Organization</option>
                {(budgetGroups || []).map((bg) => <option key={bg.id} value={bg.id}>{bg.name}</option>)}
              </select>
            </Field>
            <Field label="Calendar year">
              <select value={orgYear} onChange={(e) => setOrgYear(e.target.value === "All" ? "All" : Number(e.target.value))} className={inputCls} style={inputStyle}>
                <option value="All">All years combined</option>
                {orgCalendarYears.map((y) => <option key={y} value={y}>{y}</option>)}
              </select>
            </Field>
          </div>
        )}
      </div>
      <div className="flex justify-end gap-2 mt-6">
        <button onClick={onClose} className="px-4 py-2 rounded-md text-sm border" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>Cancel</button>
        <button
          disabled={!canCreate}
          onClick={handleCreate}
          className="px-4 py-2 rounded-md text-sm text-white"
          style={{ background: canCreate ? "#1F5C6B" : "#8A8F87" }}
        >
          Create scenario
        </button>
      </div>
    </Modal>
  );
}

function ScenarioEditor({ scenario, grants, costCenters, budgets, canEdit = true, onSave, onDelete, onBack }) {
  const [form, setForm] = useState(scenario);
  const [showCompare, setShowCompare] = useState(true);
  const cols = monthColumnsForBudget(form.periodStart);

  const updateLine = (id, patch) => setForm((f) => ({ ...f, lines: f.lines.map((l) => (l.id === id ? { ...l, ...patch } : l)) }));
  const addLine = () => setForm((f) => ({ ...f, lines: [...f.lines, newLine()] }));
  const removeLine = (id) => setForm((f) => ({ ...f, lines: f.lines.filter((l) => l.id !== id) }));
  const setAnnual = (id, val) => {
    const per = Math.round((Number(val) || 0) / 12 * 100) / 100;
    updateLine(id, { amounts: Array(12).fill(per) });
  };

  const totals = useMemo(() => {
    const revenue = form.lines.filter((l) => l.type === "revenue").reduce((a, l) => a + lineTotal(l), 0);
    const expense = form.lines.filter((l) => l.type === "expense").reduce((a, l) => a + lineTotal(l), 0);
    return { revenue, expense, net: revenue - expense };
  }, [form.lines]);

  const comparison = useMemo(() => liveComparisonForScenario(form, grants, budgets, costCenters), [form, grants, budgets, costCenters]);

  const exportXlsx = () => {
    const labels = cols.map((c) => c.label);
    const rows = [
      [form.title],
      [`Scenario${form.basedOn?.label ? ` — based on ${form.basedOn.label}` : " — started blank"}`],
      [],
      ["Category", "Subcategory", "Description", "Type", ...labels, "Total"],
      ...form.lines.map((l) => [l.category, l.subcategory || "", l.description || "", l.type, ...l.amounts, lineTotal(l)]),
      [],
      ["Total Revenue", "", "", "", ...Array(12).fill(""), totals.revenue],
      ["Total Expense", "", "", "", ...Array(12).fill(""), totals.expense],
      ["Net", "", "", "", ...Array(12).fill(""), totals.net],
    ];
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"] = [{ wch: 30 }, { wch: 18 }, { wch: 24 }, { wch: 10 }, ...MONTHS.map(() => ({ wch: 12 })), { wch: 14 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Scenario");
    const safe = (s) => (s || "scenario").replace(/[^a-z0-9]+/gi, "_").slice(0, 40);
    const arrayBuffer = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    downloadFile(`${safe(form.title)}.xlsx`, arrayBuffer, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  };

  return (
    <div className="space-y-4">
      <div className="flex items-center justify-between">
        <button onClick={onBack} className="text-sm inline-flex items-center gap-1" style={{ color: "#1F5C6B" }}>
          <ArrowRight size={14} style={{ transform: "rotate(180deg)" }} /> Back to scenarios
        </button>
        <div className="flex items-center gap-2">
          <button onClick={exportXlsx} className="inline-flex items-center gap-1.5 text-xs px-3 py-2 rounded-md border" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>
            <Download size={14} /> Export Excel
          </button>
          {canEdit && (
            <>
              <button onClick={() => onSave(form)} className="px-4 py-2 rounded-md text-sm text-white" style={{ background: "#1F5C6B" }}>Save scenario</button>
              <button onClick={onDelete} className="px-3 py-2 rounded-md text-sm border" style={{ borderColor: "#B5443A", color: "#B5443A" }}>Delete</button>
            </>
          )}
        </div>
      </div>

      <div className="bg-white rounded-lg border p-5 space-y-4" style={{ borderColor: "#E1E5DE" }}>
        <fieldset disabled={!canEdit} style={{ border: "none", margin: 0, padding: 0 }}>
        <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
          <Field label="Scenario name">
            <input className={inputCls} style={inputStyle} value={form.title} onChange={(e) => setForm({ ...form, title: e.target.value })} />
          </Field>
          <Field label="Fiscal year label (optional)">
            <input className={inputCls} style={inputStyle} value={form.fy} onChange={(e) => setForm({ ...form, fy: e.target.value })} placeholder="e.g. FY27" />
          </Field>
          <Field label="Period start (for month labels)">
            <input type="date" className={inputCls} style={inputStyle} value={form.periodStart} onChange={(e) => setForm({ ...form, periodStart: e.target.value })} />
          </Field>
          <Field label="Period end">
            <input type="date" className={inputCls} style={inputStyle} value={form.periodEnd} onChange={(e) => setForm({ ...form, periodEnd: e.target.value })} />
          </Field>
        </div>
        <p className="text-xs" style={{ color: "#8A8F87" }}>
          {form.basedOn?.type === "blank" ? "Started blank — not tied to any real grant or budget." : `Based on: ${form.basedOn?.label || "unknown"}`}
        </p>

        <div className="overflow-x-auto border rounded-lg" style={{ borderColor: "#E1E5DE" }}>
          <table className="text-xs w-full" style={{ fontFamily: "var(--mono-font)" }}>
            <thead>
              <tr style={{ background: "#F6F7F3" }}>
                <th className="text-left px-2 py-2 sticky left-0" style={{ background: "#F6F7F3", minWidth: 230 }}>Category</th>
                <th className="text-left px-2 py-2" style={{ minWidth: 230 }}>Subcategory</th>
                <th className="text-left px-2 py-2" style={{ minWidth: 180 }}>Description</th>
                <th className="text-right px-2 py-2" style={{ minWidth: 110 }}>Annual total</th>
                {cols.map((c, i) => <th key={i} className="text-right px-2 py-2" style={{ minWidth: 78 }}>{c.label}</th>)}
                <th className="text-right px-2 py-2" style={{ minWidth: 90 }}>Total</th>
                <th className="px-2 py-2"></th>
              </tr>
            </thead>
            <tbody>
              {form.lines.map((line) => {
                const cat = CATEGORIES.find((c) => c.name === line.category);
                return (
                  <tr key={line.id} className="border-t" style={{ borderColor: "#E1E5DE" }}>
                    <td className="px-2 py-1.5 sticky left-0 bg-white">
                      {line.categoryCustom ? (
                        <div className="flex gap-1">
                          <input value={line.category} onChange={(e) => updateLine(line.id, { category: e.target.value })} placeholder="Custom category" className="w-full rounded border px-1.5 py-1 text-xs" style={inputStyle} />
                          <select value={line.type} onChange={(e) => updateLine(line.id, { type: e.target.value })} className="shrink-0 rounded border px-1 py-1 text-xs" style={inputStyle}>
                            <option value="expense">Exp</option>
                            <option value="revenue">Rev</option>
                          </select>
                          <button onClick={() => updateLine(line.id, { categoryCustom: false, category: CATEGORIES[0].name, type: CATEGORIES[0].type, subcategory: "", subcategoryCustom: false })} className="shrink-0 px-1 rounded hover:bg-red-50">
                            <X size={12} style={{ color: "#B5443A" }} />
                          </button>
                        </div>
                      ) : (
                        <select
                          value={line.category}
                          onChange={(e) => {
                            if (e.target.value === CUSTOM_CATEGORY) { updateLine(line.id, { categoryCustom: true, category: "", subcategory: "", subcategoryCustom: false }); return; }
                            const nc = CATEGORIES.find((c) => c.name === e.target.value);
                            updateLine(line.id, { category: nc.name, type: nc.type, subcategory: "", subcategoryCustom: false });
                          }}
                          className="w-full rounded border px-1.5 py-1 text-xs"
                          style={inputStyle}
                        >
                          {CATEGORIES.map((c) => <option key={c.name} value={c.name}>{c.name}</option>)}
                          <option value={CUSTOM_CATEGORY}>Other (write in)…</option>
                        </select>
                      )}
                    </td>
                    <td className="px-2 py-1.5">
                      {line.categoryCustom || line.subcategoryCustom ? (
                        <div className="flex gap-1">
                          <input value={line.subcategory} onChange={(e) => updateLine(line.id, { subcategory: e.target.value })} placeholder="Custom subcategory" className="w-full rounded border px-1.5 py-1 text-xs" style={inputStyle} />
                          {!line.categoryCustom && (
                            <button onClick={() => updateLine(line.id, { subcategoryCustom: false, subcategory: "" })} className="shrink-0 px-1 rounded hover:bg-red-50">
                              <X size={12} style={{ color: "#B5443A" }} />
                            </button>
                          )}
                        </div>
                      ) : (
                        <select
                          value={line.subcategory}
                          onChange={(e) => {
                            if (e.target.value === CUSTOM_CATEGORY) { updateLine(line.id, { subcategoryCustom: true, subcategory: "" }); return; }
                            updateLine(line.id, { subcategory: e.target.value });
                          }}
                          className="w-full rounded border px-1.5 py-1 text-xs"
                          style={inputStyle}
                        >
                          <option value="">Select subcategory</option>
                          {cat?.subs.map((s) => <option key={s} value={s}>{s}</option>)}
                          <option value={CUSTOM_CATEGORY}>Other (write in)…</option>
                        </select>
                      )}
                    </td>
                    <td className="px-2 py-1.5">
                      <input
                        value={line.description || ""}
                        onChange={(e) => updateLine(line.id, { description: e.target.value })}
                        placeholder="Optional note"
                        className="w-full rounded border px-1.5 py-1 text-xs"
                        style={inputStyle}
                      />
                    </td>
                    <td className="px-2 py-1.5">
                      <input
                        type="number"
                        className="w-full rounded border px-1.5 py-1 text-xs text-right"
                        style={inputStyle}
                        placeholder="0"
                        onChange={(e) => setAnnual(line.id, e.target.value)}
                      />
                    </td>
                    {line.amounts.map((a, i) => (
                      <td key={i} className="px-1 py-1.5">
                        <input
                          type="number"
                          value={a}
                          onChange={(e) => {
                            const vals = [...line.amounts];
                            vals[i] = Number(e.target.value) || 0;
                            updateLine(line.id, { amounts: vals });
                          }}
                          className="w-full rounded border px-1.5 py-1 text-xs text-right"
                          style={inputStyle}
                        />
                      </td>
                    ))}
                    <td className="px-2 py-1.5 text-right font-medium" style={{ color: "#1C2624" }}>{fmt(lineTotal(line))}</td>
                    <td className="px-2 py-1.5">
                      <button onClick={() => removeLine(line.id)} className="p-1 rounded hover:bg-red-50">
                        <Trash2 size={13} style={{ color: "#B5443A" }} />
                      </button>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
        <button onClick={addLine} className="inline-flex items-center gap-1 text-xs px-3 py-1.5 rounded-md border" style={{ borderColor: "#E1E5DE", color: "#1F5C6B" }}>
          <Plus size={13} /> Add row
        </button>

        <div className="grid grid-cols-3 gap-4 text-sm pt-2 border-t" style={{ borderColor: "#E1E5DE" }}>
          <div>Revenue: <span style={{ color: "#2F6F53", fontWeight: 600 }}>{fmt(totals.revenue)}</span></div>
          <div>Expense: <span style={{ color: "#B5443A", fontWeight: 600 }}>{fmt(totals.expense)}</span></div>
          <div>Net: <span style={{ fontWeight: 600 }}>{fmt(totals.net)}</span></div>
        </div>
        </fieldset>
      </div>

      <div className="bg-white rounded-lg border p-5" style={{ borderColor: "#E1E5DE" }}>
        <div className="flex items-center justify-between mb-3">
          <h2 className="font-display text-base" style={{ color: "#1C2624" }}>Compare to real numbers</h2>
          <button onClick={() => setShowCompare((v) => !v)} className="text-xs" style={{ color: "#1F5C6B" }}>{showCompare ? "Hide" : "Show"}</button>
        </div>
        {showCompare && (
          comparison.available ? (
            <div className="overflow-x-auto">
              <table className="text-sm w-full">
                <thead>
                  <tr style={{ color: "#8A8F87" }}>
                    <th className="text-left py-1.5 font-medium">Category</th>
                    <th className="text-right py-1.5 font-medium">Scenario</th>
                    <th className="text-right py-1.5 font-medium">Real (current)</th>
                    <th className="text-right py-1.5 font-medium">Variance</th>
                  </tr>
                </thead>
                <tbody>
                  {form.lines.reduce((acc, l) => { if (!acc.includes(l.category)) acc.push(l.category); return acc; }, []).map((catName) => {
                    const scenarioTotal = form.lines.filter((l) => l.category === catName).reduce((a, l) => a + lineTotal(l), 0);
                    const realVals = comparison.byCategory[catName] || Array(12).fill(0);
                    const realTotal = realVals.reduce((a, b) => a + b, 0);
                    const variance = scenarioTotal - realTotal;
                    return (
                      <tr key={catName} className="border-t" style={{ borderColor: "#E1E5DE" }}>
                        <td className="py-1.5" style={{ color: "#1C2624" }}>{catName}</td>
                        <td className="py-1.5 text-right" style={{ fontVariantNumeric: "tabular-nums" }}>{fmt(scenarioTotal)}</td>
                        <td className="py-1.5 text-right" style={{ fontVariantNumeric: "tabular-nums", color: "#8A8F87" }}>{fmt(realTotal)}</td>
                        <td className="py-1.5 text-right font-medium" style={{ fontVariantNumeric: "tabular-nums", color: variance >= 0 ? "#2F6F53" : "#B5443A" }}>
                          {variance >= 0 ? "+" : ""}{fmt(variance)}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          ) : (
            <p className="text-sm" style={{ color: "#8A8F87" }}>{comparison.reason}</p>
          )
        )}
      </div>
    </div>
  );
}

function TrashView({ trash, setTrash, setGrants, setBudgets, setReports, setTasks, setInvoices, setStaff, setCostCenters, setScenarios, isAdmin, canEdit, logActivity }) {
  const [confirm, setConfirm] = useState(null);

  const labelFor = (t) => {
    switch (t.entityType) {
      case "grant": return t.data?.title || "Untitled grant";
      case "budget": return t.data?.title || "Untitled budget";
      case "costCenter": return t.data?.name || "Untitled cost center";
      case "report": return t.data?.title || "Untitled report";
      case "task": return t.data?.title || "Untitled task";
      case "invoice": return t.data?.invoiceNumber || "Untitled invoice";
      case "staff": return t.data?.name || "Untitled staff member";
      case "scenario": return t.data?.title || "Untitled scenario";
      default: return "Item";
    }
  };
  const typeLabel = {
    grant: "Grant", budget: "Budget", costCenter: "Cost Center", report: "Grant Report",
    task: "Task", invoice: "Invoice", staff: "Staff", scenario: "Scenario",
  };

  const restore = (t) => {
    if (!t.data) return;
    switch (t.entityType) {
      case "grant":
        setGrants((prev) => [...prev, t.data]);
        if (t.extra?.budgets?.length) setBudgets((prev) => [...prev, ...t.extra.budgets]);
        if (t.extra?.reports?.length) setReports?.((prev) => [...prev, ...t.extra.reports]);
        if (t.extra?.tasks?.length) setTasks?.((prev) => [...prev, ...t.extra.tasks]);
        if (t.extra?.invoices?.length) setInvoices?.((prev) => [...prev, ...t.extra.invoices]);
        if (t.extra?.staffAllocations?.length) {
          setStaff?.((prev) => prev.map((s) => {
            const match = t.extra.staffAllocations.find((x) => x.staffId === s.id);
            if (!match) return s;
            return { ...s, allocations: [...(s.allocations || []), ...match.allocations] };
          }));
        }
        break;
      case "costCenter":
        setCostCenters((prev) => [...prev, t.data]);
        if (t.extra?.budgets?.length) setBudgets((prev) => [...prev, ...t.extra.budgets]);
        break;
      case "budget": setBudgets((prev) => [...prev, t.data]); break;
      case "report": setReports?.((prev) => [...prev, t.data]); break;
      case "task": setTasks?.((prev) => [...prev, t.data]); break;
      case "invoice": setInvoices?.((prev) => [...prev, t.data]); break;
      case "staff": setStaff?.((prev) => [...prev, t.data]); break;
      case "scenario": setScenarios?.((prev) => [...prev, t.data]); break;
      default: break;
    }
    setTrash((prev) => prev.filter((x) => x.id !== t.id));
    logActivity?.("Data", "Restored", `Restored ${typeLabel[t.entityType] || "item"} "${labelFor(t)}" from Trash`);
  };

  const permanentlyDelete = (id) => {
    setTrash((prev) => prev.filter((x) => x.id !== id));
    setConfirm(null);
  };

  const sorted = [...trash].sort((a, b) => new Date(b.deletedAt) - new Date(a.deletedAt));

  return (
    <div className="space-y-4">
      <div>
        <h1 className="font-display text-2xl" style={{ color: "#1C2624" }}>Trash</h1>
        <p className="text-sm mt-1" style={{ color: "#5B6B66" }}>
          Deleted items land here and can be restored. Items are automatically removed for good after 90 days.
        </p>
      </div>

      {sorted.length === 0 ? (
        <div className="bg-white rounded-lg border p-10 text-center" style={{ borderColor: "#E1E5DE", color: "#8A8F87" }}>
          Trash is empty.
        </div>
      ) : (
        <div className="bg-white rounded-lg border divide-y" style={{ borderColor: "#E1E5DE" }}>
          {sorted.map((t) => (
            <div key={t.id} className="px-4 py-3 flex items-center justify-between text-sm">
              <div>
                <div className="flex items-center gap-2">
                  <span className="text-xs px-1.5 py-0.5 rounded-full" style={{ background: "#F6F7F3", color: "#5B6B66" }}>{typeLabel[t.entityType] || t.entityType}</span>
                  <span style={{ color: "#1C2624" }}>{labelFor(t)}</span>
                </div>
                <div className="text-xs mt-0.5" style={{ color: "#8A8F87" }}>
                  Deleted by {t.deletedBy || "Unknown"} · {new Date(t.deletedAt).toLocaleString("en-US", { month: "short", day: "numeric", year: "numeric", hour: "numeric", minute: "2-digit" })}
                </div>
              </div>
              <div className="flex items-center gap-2 shrink-0">
                {canEdit && (
                  <button onClick={() => restore(t)} className="text-xs px-3 py-1.5 rounded-md text-white" style={{ background: "#1F5C6B" }}>Restore</button>
                )}
                {isAdmin && (
                  <button onClick={() => setConfirm(t.id)} className="text-xs px-3 py-1.5 rounded-md border" style={{ borderColor: "#B5443A", color: "#B5443A" }}>Delete forever</button>
                )}
              </div>
            </div>
          ))}
        </div>
      )}

      {confirm && (
        <ConfirmModal
          message="This permanently and irreversibly deletes this item. There is no way to get it back after this."
          onConfirm={() => permanentlyDelete(confirm)}
          onCancel={() => setConfirm(null)}
        />
      )}
    </div>
  );
}

function ScenariosView({ scenarios, setScenarios, grants, budgets, costCenters, budgetGroups, whoami, setTrash, canEdit, logActivity }) {
  const [openId, setOpenId] = useState(null);
  const [wizardOpen, setWizardOpen] = useState(false);
  const [confirm, setConfirm] = useState(null);

  const open = scenarios.find((s) => s.id === openId);

  const createScenario = (scen) => {
    scen.createdBy = whoami || "Unknown";
    setScenarios((prev) => [...prev, scen]);
    logActivity?.("Scenario", "Created", scen.title);
    setWizardOpen(false);
    setOpenId(scen.id);
  };
  const saveScenario = (scen) => {
    setScenarios((prev) => prev.map((s) => (s.id === scen.id ? scen : s)));
    logActivity?.("Scenario", "Updated", scen.title);
  };
  const deleteScenario = (id) => {
    const s = scenarios.find((x) => x.id === id);
    pushTrash(setTrash, "scenario", s, whoami);
    setScenarios((prev) => prev.filter((x) => x.id !== id));
    logActivity?.("Scenario", "Deleted", s?.title || "Untitled scenario");
    setOpenId(null);
    setConfirm(null);
  };

  if (open) {
    return (
      <ScenarioEditor
        scenario={open}
        grants={grants}
        costCenters={costCenters}
        budgets={budgets}
        canEdit={canEdit}
        onSave={saveScenario}
        onDelete={canEdit ? () => setConfirm(open.id) : undefined}
        onBack={() => setOpenId(null)}
      />
    );
  }

  return (
    <div className="space-y-4">
      <div className="flex items-center justify-between">
        <div>
          <h1 className="font-display text-2xl" style={{ color: "#1C2624" }}>Scenarios</h1>
          <p className="text-sm mt-1" style={{ color: "#5B6B66" }}>A sandbox to play with what-if numbers — never touches real budgets or actuals</p>
        </div>
        {canEdit && (
          <button onClick={() => setWizardOpen(true)} className="inline-flex items-center gap-1.5 px-3 py-2 rounded-md text-sm text-white" style={{ background: "#1F5C6B" }}>
            <Plus size={16} /> New scenario
          </button>
        )}
      </div>

      {scenarios.length === 0 ? (
        <div className="bg-white rounded-lg border p-10 text-center" style={{ borderColor: "#E1E5DE", color: "#8A8F87" }}>
          No scenarios yet — create one to start playing with what-if numbers.
        </div>
      ) : (
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
          {scenarios.map((s) => {
            const revenue = s.lines.filter((l) => l.type === "revenue").reduce((a, l) => a + lineTotal(l), 0);
            const expense = s.lines.filter((l) => l.type === "expense").reduce((a, l) => a + lineTotal(l), 0);
            return (
              <button
                key={s.id}
                onClick={() => setOpenId(s.id)}
                className="text-left bg-white rounded-lg border p-4 hover:shadow-sm transition-shadow"
                style={{ borderColor: "#E1E5DE" }}
              >
                <div className="font-medium" style={{ color: "#1C2624" }}>{s.title || "Untitled scenario"}</div>
                <div className="text-xs mt-1" style={{ color: "#8A8F87" }}>{s.basedOn?.label ? `Based on ${s.basedOn.label}` : "Started blank"}</div>
                <div className="flex items-center gap-4 mt-3 text-xs">
                  <span style={{ color: "#2F6F53" }}>Rev {fmt(revenue)}</span>
                  <span style={{ color: "#B5443A" }}>Exp {fmt(expense)}</span>
                </div>
                <div className="text-xs mt-2" style={{ color: "#8A8F87" }}>By {s.createdBy || "Unknown"}</div>
              </button>
            );
          })}
        </div>
      )}

      {wizardOpen && (
        <NewScenarioModal
          grants={grants}
          costCenters={costCenters}
          budgets={budgets}
          budgetGroups={budgetGroups}
          onCreate={createScenario}
          onClose={() => setWizardOpen(false)}
        />
      )}
      {confirm && (
        <ConfirmModal
          message="This will permanently delete this scenario. It has no effect on any real budget or actual data."
          onConfirm={() => deleteScenario(confirm)}
          onCancel={() => setConfirm(null)}
        />
      )}
    </div>
  );
}

function OrgBudgetView({ grants, budgets, costCenters, budgetGroups }) {
  const [calYear, setCalYear] = useState("All");
  const [scope, setScope] = useState("all"); // all | a budget group id
  const [viewMode, setViewMode] = useState("monthly");
  const [dataMode, setDataMode] = useState("plan");

  const scopedGrantIds = useMemo(() => {
    if (scope === "all") return null;
    return new Set(grants.filter((g) => g.budgetGroupId === scope).map((g) => g.id));
  }, [scope, grants]);
  const scopedCostCenterIds = useMemo(() => {
    if (scope === "all") return null;
    return new Set((costCenters || []).filter((c) => c.budgetGroupId === scope).map((c) => c.id));
  }, [scope, costCenters]);

  const scopedBudgets = (scope === "all"
    ? budgets
    : budgets.filter((b) => (b.grantId && scopedGrantIds.has(b.grantId)) || (b.costCenterId && scopedCostCenterIds.has(b.costCenterId)))
  ).filter((b) => b.status === "Active");

  // Real calendar years actually touched by any in-scope budget's period, so the
  // year picker reflects reality even when grants run off the calendar year.
  const calendarYears = useMemo(() => {
    const years = new Set();
    scopedBudgets.forEach((b) => monthColumnsForBudget(b.periodStart).forEach((col) => years.add(col.year)));
    return [...years].sort();
  }, [scopedBudgets]);

  const amountsField = dataMode === "plan" ? "amounts" : "actuals";
  const lineValue = (l) => (dataMode === "plan" ? lineTotal(l) : lineActualTotal(l));

  const revenueCats = CATEGORIES.filter((c) => c.type === "revenue");
  const expenseCats = CATEGORIES.filter((c) => c.type === "expense");

  const grouped = useMemo(() => {
    const map = {};
    CATEGORIES.forEach((c) => { map[c.name] = { type: c.type, monthly: Array(12).fill(0), subs: {} }; });
    scopedBudgets.forEach((b) => {
      const cols = monthColumnsForBudget(b.periodStart);
      b.lines.forEach((l) => {
        if (!map[l.category]) map[l.category] = { type: l.type, monthly: Array(12).fill(0), subs: {} };
        const bucket = map[l.category];
        const vals = l[amountsField] || Array(12).fill(0);
        vals.forEach((a, i) => {
          const col = cols[i];
          if (calYear !== "All" && col.year !== calYear) return;
          const slot = calYear === "All" ? col.monthIndex : col.monthIndex;
          bucket.monthly[slot] += Number(a) || 0;
          if (l.subcategory) {
            if (!bucket.subs[l.subcategory]) bucket.subs[l.subcategory] = Array(12).fill(0);
            bucket.subs[l.subcategory][slot] += Number(a) || 0;
          }
        });
      });
    });
    return map;
  }, [scopedBudgets, amountsField, calYear]);

  const yearCompare = useMemo(() => {
    const years = [...new Set(scopedBudgets.map((b) => b.fy || "Unspecified"))].sort();
    const byYear = {};
    years.forEach((fy) => {
      const budgetsForFy = scopedBudgets.filter((b) => (b.fy || "Unspecified") === fy);
      const catTotals = {};
      CATEGORIES.forEach((c) => { catTotals[c.name] = 0; });
      budgetsForFy.forEach((b) => b.lines.forEach((l) => {
        if (catTotals[l.category] === undefined) catTotals[l.category] = 0;
        catTotals[l.category] += lineValue(l);
      }));
      const revenue = revenueCats.reduce((a, c) => a + (catTotals[c.name] || 0), 0);
      const expense = expenseCats.reduce((a, c) => a + (catTotals[c.name] || 0), 0);
      byYear[fy] = { catTotals, revenue, expense, net: revenue - expense };
    });
    return { years, byYear };
  }, [scopedBudgets, dataMode]);

  const sumRows = (rows) => rows.reduce((acc, r) => acc.map((v, i) => v + r[i]), Array(12).fill(0));
  const totalRevenue = sumRows(revenueCats.map((c) => grouped[c.name].monthly));
  const totalExpense = sumRows(expenseCats.map((c) => grouped[c.name].monthly));
  const net = totalRevenue.map((v, i) => v - totalExpense[i]);

  const exportCsv = () => {
    const monthLabels = MONTHS.map((m) => (calYear === "All" ? m : `${m} ${calYear}`));
    const rows = [["Category", "Subcategory", ...monthLabels, "Total"]];
    const pushSection = (cats) => cats.forEach((c) => {
      const bucket = grouped[c.name];
      rows.push([c.name, "", ...bucket.monthly, bucket.monthly.reduce((a, b) => a + b, 0)]);
      Object.entries(bucket.subs).forEach(([sub, vals]) => {
        rows.push([c.name, sub, ...vals, vals.reduce((a, b) => a + b, 0)]);
      });
    });
    pushSection(revenueCats);
    rows.push(["Total Revenue", "", ...totalRevenue, totalRevenue.reduce((a, b) => a + b, 0)]);
    pushSection(expenseCats);
    rows.push(["Total Expense", "", ...totalExpense, totalExpense.reduce((a, b) => a + b, 0)]);
    rows.push(["Net", "", ...net, net.reduce((a, b) => a + b, 0)]);
    const csv = rows.map((r) => r.map((c) => `"${String(c).replace(/"/g, '""')}"`).join(",")).join("\n");
    downloadFile("nations-finest-organizational-budget.csv", csv, "text/csv");
  };

  return (
    <div className="space-y-4">
      <div className="flex items-center justify-between">
        <div>
          <h1 className="font-display text-2xl" style={{ color: "#1C2624" }}>Organizational Budget</h1>
          <p className="text-sm mt-1" style={{ color: "#5B6B66" }}>Rolled up across every grant's budget, by category and month — showing {dataMode === "plan" ? "planned" : "actual"} figures</p>
        </div>
        <div className="flex items-center gap-2">
          <div className="inline-flex rounded-md border overflow-hidden" style={{ borderColor: "#E1E5DE" }}>
            <button
              onClick={() => setDataMode("plan")}
              className="px-3 py-2 text-sm font-medium"
              style={{ background: dataMode === "plan" ? "#2F6F53" : "#FFFFFF", color: dataMode === "plan" ? "#FFFFFF" : "#5B6B66" }}
            >
              Plan
            </button>
            <button
              onClick={() => setDataMode("actual")}
              className="px-3 py-2 text-sm font-medium"
              style={{ background: dataMode === "actual" ? "#2F6F53" : "#FFFFFF", color: dataMode === "actual" ? "#FFFFFF" : "#5B6B66" }}
            >
              Actual
            </button>
          </div>
          <div className="inline-flex rounded-md border overflow-hidden" style={{ borderColor: "#E1E5DE" }}>
            <button
              onClick={() => setViewMode("monthly")}
              className="px-3 py-2 text-sm font-medium"
              style={{ background: viewMode === "monthly" ? "#2F6F53" : "#FFFFFF", color: viewMode === "monthly" ? "#FFFFFF" : "#5B6B66" }}
            >
              Monthly detail
            </button>
            <button
              onClick={() => setViewMode("compare")}
              className="px-3 py-2 text-sm font-medium"
              style={{ background: viewMode === "compare" ? "#2F6F53" : "#FFFFFF", color: viewMode === "compare" ? "#FFFFFF" : "#5B6B66" }}
            >
              Year comparison
            </button>
          </div>
          <button onClick={exportCsv} className="inline-flex items-center gap-1.5 px-3 py-2 rounded-md text-sm border" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>
            <Download size={15} /> Export CSV
          </button>
          <button onClick={() => printSection("org-budget-print-area", "GrantFlow Organizational Budget")} className="no-print inline-flex items-center gap-1.5 px-3 py-2 rounded-md text-sm border" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>
            <Printer size={15} /> Print / Save PDF
          </button>
        </div>
      </div>

      <div className="flex flex-wrap gap-4">
        <Field label="Scope">
          <select value={scope} onChange={(e) => setScope(e.target.value)} className={inputCls} style={{ ...inputStyle, maxWidth: 260 }}>
            <option value="all">Whole Organization</option>
            {(budgetGroups || []).map((bg) => <option key={bg.id} value={bg.id}>{bg.name}</option>)}
          </select>
        </Field>
        {viewMode === "monthly" && (
          <Field label="Calendar year">
            <select value={calYear} onChange={(e) => setCalYear(e.target.value === "All" ? "All" : Number(e.target.value))} className={inputCls} style={{ ...inputStyle, maxWidth: 240 }}>
              <option value="All">All years combined</option>
              {calendarYears.map((y) => <option key={y} value={y}>{y}</option>)}
            </select>
          </Field>
        )}
      </div>

      <div id="org-budget-print-area">
      {viewMode === "compare" ? (
        yearCompare.years.length === 0 ? (
          <div className="bg-white rounded-lg border p-10 text-center" style={{ borderColor: "#E1E5DE", color: "#8A8F87" }}>
            No budget data yet — add grant budgets to compare across years.
          </div>
        ) : (
          <div className="overflow-x-auto border rounded-lg bg-white" style={{ borderColor: "#E1E5DE" }}>
            <table className="w-full" style={{ fontFamily: "var(--mono-font)" }}>
              <thead>
                <tr style={{ background: "#F6F7F3" }}>
                  <th className="text-left px-3 py-2 text-xs sticky left-0" style={{ background: "#F6F7F3", minWidth: 220 }}>Account</th>
                  {yearCompare.years.map((fy) => <th key={fy} className="text-right px-3 py-2 text-xs" style={{ minWidth: 120 }}>{fy}</th>)}
                </tr>
              </thead>
              <tbody>
                <tr style={{ background: "#F6F7F3" }}>
                  <td className="px-3 py-1.5 text-xs font-semibold sticky left-0" style={{ background: "#F6F7F3" }}>Revenue</td>
                  {yearCompare.years.map((fy) => <td key={fy} />)}
                </tr>
                {revenueCats.map((c) => (
                  <tr key={c.name} className="border-t" style={{ borderColor: "#E1E5DE" }}>
                    <td className="px-3 py-1.5 text-xs sticky left-0 bg-white" style={{ paddingLeft: 28, color: "#2F6F53" }}>{c.name}</td>
                    {yearCompare.years.map((fy) => (
                      <td key={fy} className="px-3 py-1.5 text-xs text-right" style={{ fontVariantNumeric: "tabular-nums", color: "#2F6F53" }}>
                        {yearCompare.byYear[fy].catTotals[c.name] ? fmt(yearCompare.byYear[fy].catTotals[c.name]) : "—"}
                      </td>
                    ))}
                  </tr>
                ))}
                <tr className="border-t" style={{ borderColor: "#E1E5DE" }}>
                  <td className="px-3 py-1.5 text-xs font-semibold sticky left-0 bg-white" style={{ color: "#2F6F53" }}>Total Revenue</td>
                  {yearCompare.years.map((fy) => (
                    <td key={fy} className="px-3 py-1.5 text-xs text-right font-semibold" style={{ fontVariantNumeric: "tabular-nums", color: "#2F6F53" }}>{fmt(yearCompare.byYear[fy].revenue)}</td>
                  ))}
                </tr>

                <tr style={{ background: "#F6F7F3" }}>
                  <td className="px-3 py-1.5 text-xs font-semibold sticky left-0" style={{ background: "#F6F7F3" }}>Expense</td>
                  {yearCompare.years.map((fy) => <td key={fy} />)}
                </tr>
                {expenseCats.map((c) => (
                  <tr key={c.name} className="border-t" style={{ borderColor: "#E1E5DE" }}>
                    <td className="px-3 py-1.5 text-xs sticky left-0 bg-white" style={{ paddingLeft: 28, color: "#1C2624" }}>{c.name}</td>
                    {yearCompare.years.map((fy) => (
                      <td key={fy} className="px-3 py-1.5 text-xs text-right" style={{ fontVariantNumeric: "tabular-nums", color: "#1C2624" }}>
                        {yearCompare.byYear[fy].catTotals[c.name] ? fmt(yearCompare.byYear[fy].catTotals[c.name]) : "—"}
                      </td>
                    ))}
                  </tr>
                ))}
                <tr className="border-t" style={{ borderColor: "#E1E5DE" }}>
                  <td className="px-3 py-1.5 text-xs font-semibold sticky left-0 bg-white">Total Expense</td>
                  {yearCompare.years.map((fy) => (
                    <td key={fy} className="px-3 py-1.5 text-xs text-right font-semibold" style={{ fontVariantNumeric: "tabular-nums" }}>{fmt(yearCompare.byYear[fy].expense)}</td>
                  ))}
                </tr>
                <tr className="border-t" style={{ borderColor: "#E1E5DE" }}>
                  <td className="px-3 py-1.5 text-xs font-semibold sticky left-0 bg-white">Net Total</td>
                  {yearCompare.years.map((fy) => (
                    <td key={fy} className="px-3 py-1.5 text-xs text-right font-semibold" style={{ fontVariantNumeric: "tabular-nums", color: yearCompare.byYear[fy].net >= 0 ? "#2F6F53" : "#B5443A" }}>
                      {fmt(yearCompare.byYear[fy].net)}
                    </td>
                  ))}
                </tr>
              </tbody>
            </table>
          </div>
        )
      ) : scopedBudgets.length === 0 ? (
        <div className="bg-white rounded-lg border p-10 text-center" style={{ borderColor: "#E1E5DE", color: "#8A8F87" }}>
          No budget data yet — add grant budgets to see the organizational rollup.
        </div>
      ) : (
        <div className="overflow-x-auto border rounded-lg bg-white" style={{ borderColor: "#E1E5DE" }}>
          <table className="w-full" style={{ fontFamily: "var(--mono-font)" }}>
            <thead>
              <tr style={{ background: "#F6F7F3" }}>
                <th className="text-left px-3 py-2 text-xs sticky left-0" style={{ background: "#F6F7F3", minWidth: 220 }}>Account</th>
                {MONTHS.map((m) => <th key={m} className="text-right px-2 py-2 text-xs" style={{ minWidth: 90 }}>{calYear === "All" ? m : `${m} ${String(calYear).slice(2)}`}</th>)}
                <th className="text-right px-3 py-2 text-xs" style={{ minWidth: 100 }}>Total</th>
              </tr>
            </thead>
            <tbody>
              <OrgBudgetRow label="Revenue" values={Array(12).fill(0)} isHeader bold />
              {revenueCats.map((c) => (
                <Fragment key={c.name}>
                  <OrgBudgetRow label={c.name} values={grouped[c.name].monthly} indent color="#2F6F53" />
                  {Object.entries(grouped[c.name].subs).map(([sub, vals]) => (
                    <OrgBudgetRow key={c.name + sub} label={sub} values={vals} indent color="#5B6B66" />
                  ))}
                </Fragment>
              ))}
              <OrgBudgetRow label="Total Revenue" values={totalRevenue} bold color="#2F6F53" />

              <OrgBudgetRow label="Expense" values={Array(12).fill(0)} isHeader bold />
              {expenseCats.map((c) => (
                <Fragment key={c.name}>
                  <OrgBudgetRow label={c.name} values={grouped[c.name].monthly} indent />
                  {Object.entries(grouped[c.name].subs).map(([sub, vals]) => (
                    <OrgBudgetRow key={c.name + sub} label={sub} values={vals} indent color="#5B6B66" />
                  ))}
                </Fragment>
              ))}
              <OrgBudgetRow label="Total Expense" values={totalExpense} bold />
              <OrgBudgetRow label="Net Total" values={net} bold color={net.reduce((a, b) => a + b, 0) >= 0 ? "#2F6F53" : "#B5443A"} />
            </tbody>
          </table>
        </div>
      )}
      </div>
    </div>
  );
}

// ---------- personnel / payroll ----------

function StaffModal({ staff, grants, costCenters, canEdit = true, onSave, onClose, onDelete }) {
  const [form, setForm] = useState(staff ? { ...staff, status: staff.status || "Active" } : {
    id: uid(), name: "", position: "", department: "", exempt: "Non-exempt",
    payType: "Salary", annualSalary: 0, hourlyRate: 0, annualHours: ANNUAL_HOURS,
    fte: 1, allocations: [], site: "", status: "Active",
  });
  const set = (k) => (e) => setForm({ ...form, [k]: e.target.value });

  const addAlloc = () => setForm({ ...form, allocations: [...form.allocations, newAllocation()] });
  const updateAlloc = (id, patch) => setForm({ ...form, allocations: form.allocations.map((a) => (a.id === id ? { ...a, ...patch } : a)) });
  const removeAlloc = (id) => setForm({ ...form, allocations: form.allocations.filter((a) => a.id !== id) });

  const cost = staffAnnualCost(form);
  const allocatedPct = staffAllocatedTotal(form);

  return (
    <Modal title={staff ? (canEdit ? "Edit staff member" : "View staff member") : "New staff member"} onClose={onClose} wide>
      <fieldset disabled={!canEdit} style={{ border: "none", margin: 0, padding: 0 }}>
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
        <Field label="Name">
          <input className={inputCls} style={inputStyle} value={form.name} onChange={set("name")} placeholder="Last, First" />
        </Field>
        <Field label="Position">
          <input className={inputCls} style={inputStyle} value={form.position} onChange={set("position")} />
        </Field>
        <Field label="Department">
          <input className={inputCls} style={inputStyle} value={form.department} onChange={set("department")} placeholder="e.g. Residential, GPD, HCHV" />
        </Field>
        <Field label="Site">
          <select className={inputCls} style={inputStyle} value={form.site} onChange={set("site")}>
            <option value="">No site set</option>
            {SITE_OPTIONS.map((s) => <option key={s} value={s}>{s}</option>)}
          </select>
        </Field>
        <Field label="Status">
          <select className={inputCls} style={inputStyle} value={form.status || "Active"} onChange={set("status")}>
            {STAFF_STATUSES.map((s) => <option key={s} value={s}>{s}</option>)}
          </select>
        </Field>
        <Field label="Exempt status">
          <select className={inputCls} style={inputStyle} value={form.exempt} onChange={set("exempt")}>
            <option>Exempt</option>
            <option>Non-exempt</option>
          </select>
        </Field>
        <Field label="Pay type">
          <select className={inputCls} style={inputStyle} value={form.payType} onChange={set("payType")}>
            <option>Salary</option>
            <option>Hourly</option>
          </select>
        </Field>
        <Field label="FTE">
          <input type="number" step="0.05" min="0" max="1" className={inputCls} style={inputStyle} value={form.fte} onChange={set("fte")} />
        </Field>
        {form.payType === "Salary" ? (
          <Field label="Annual salary">
            <input type="number" className={inputCls} style={inputStyle} value={form.annualSalary} onChange={set("annualSalary")} />
          </Field>
        ) : (
          <>
            <Field label="Hourly rate">
              <input type="number" className={inputCls} style={inputStyle} value={form.hourlyRate} onChange={set("hourlyRate")} />
            </Field>
            <Field label="Annual hours">
              <input type="number" className={inputCls} style={inputStyle} value={form.annualHours} onChange={set("annualHours")} />
            </Field>
          </>
        )}
      </div>

      <div className="mt-3 text-sm">
        <span style={{ color: "#8A8F87" }}>Computed annual cost: </span>
        <span className="font-medium" style={{ color: "#1C2624" }}>{fmt(cost)}</span>
      </div>

      <div className="mt-5">
        <div className="flex items-center justify-between mb-2">
          <h3 className="text-sm font-medium" style={{ color: "#1C2624" }}>Allocations</h3>
          <span className="text-xs" style={{ color: allocatedPct > 100 ? "#B5443A" : "#8A8F87" }}>{allocatedPct}% allocated</span>
        </div>
        <div className="space-y-2">
          {form.allocations.map((a) => {
            const isCostCenter = a.type === "costCenter";
            return (
              <div key={a.id} className="flex items-center gap-2">
                <div className="inline-flex rounded-md border overflow-hidden shrink-0" style={{ borderColor: "#E1E5DE" }}>
                  <button
                    onClick={() => updateAlloc(a.id, { type: "grant", costCenterId: "" })}
                    className="px-2 py-1.5 text-xs font-medium"
                    style={{ background: !isCostCenter ? "#1F5C6B" : "#FFFFFF", color: !isCostCenter ? "#FFFFFF" : "#5B6B66" }}
                  >
                    Grant
                  </button>
                  <button
                    onClick={() => updateAlloc(a.id, { type: "costCenter", grantId: "" })}
                    className="px-2 py-1.5 text-xs font-medium"
                    style={{ background: isCostCenter ? "#1F5C6B" : "#FFFFFF", color: isCostCenter ? "#FFFFFF" : "#5B6B66" }}
                  >
                    Cost Center
                  </button>
                </div>
                {isCostCenter ? (
                  <select
                    value={a.costCenterId}
                    onChange={(e) => updateAlloc(a.id, { costCenterId: e.target.value })}
                    className={inputCls}
                    style={{ ...inputStyle, flex: 1 }}
                  >
                    <option value="">Select a cost center</option>
                    {(costCenters || []).map((c) => <option key={c.id} value={c.id}>{c.name}</option>)}
                  </select>
                ) : (
                  <GrantPicker
                    grants={grants}
                    value={a.grantId}
                    onChange={(v) => updateAlloc(a.id, { grantId: v })}
                    noneLabel="Select a grant"
                    wrapStyle={{ flex: 1 }}
                  />
                )}
                <div className="relative w-28">
                  <input
                    type="number" min="0" max="100"
                    value={a.percent}
                    onChange={(e) => updateAlloc(a.id, { percent: e.target.value })}
                    className="w-full rounded-md border px-2 py-1.5 text-sm text-right"
                    style={inputStyle}
                  />
                </div>
                <span className="text-xs" style={{ color: "#8A8F87" }}>%</span>
                <button onClick={() => removeAlloc(a.id)} className="p-1 rounded hover:bg-red-50">
                  <Trash2 size={14} style={{ color: "#B5443A" }} />
                </button>
              </div>
            );
          })}
        </div>
        <button onClick={addAlloc} className="inline-flex items-center gap-1 text-xs px-3 py-1.5 rounded-md border mt-2" style={{ borderColor: "#E1E5DE", color: "#1F5C6B" }}>
          <Plus size={13} /> Add allocation
        </button>
      </div>
      </fieldset>

      <div className="flex justify-between gap-2 mt-6">
        {onDelete ? (
          <button onClick={onDelete} className="px-4 py-2 rounded-md text-sm border" style={{ borderColor: "#E1E5DE", color: "#B5443A" }}>Delete staff member</button>
        ) : <span />}
        <div className="flex gap-2">
          <button onClick={onClose} className="px-4 py-2 rounded-md text-sm border" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>Cancel</button>
          {canEdit && (
            <button
              onClick={() => { if (!form.name.trim()) return; onSave(form); }}
              className="px-4 py-2 rounded-md text-sm text-white"
              style={{ background: "#1F5C6B" }}
            >
              Save staff member
            </button>
          )}
        </div>
      </div>
    </Modal>
  );
}

function PersonnelView({ grants, staff, setStaff, costCenters, setTrash, currentUserEmail, canEdit, initialOpenStaffId, logActivity }) {
  const [modal, setModal] = useState(() => (initialOpenStaffId ? staff.find((s) => s.id === stripNonce(initialOpenStaffId)) || null : null));
  const [confirm, setConfirm] = useState(null);
  const [deptFilter, setDeptFilter] = useState("All");
  const [siteFilter, setSiteFilter] = useState("All");
  const [statusFilter, setStatusFilter] = useState("All");
  const [sortBy, setSortBy] = useState("name");

  const departments = ["All", ...new Set(staff.map((s) => s.department).filter(Boolean))];
  const visible = staff
    .filter((s) => deptFilter === "All" || s.department === deptFilter)
    .filter((s) => siteFilter === "All" || s.site === siteFilter)
    .filter((s) => statusFilter === "All" || (s.status || "Active") === statusFilter)
    .slice()
    .sort((a, b) => {
      if (sortBy === "status") return (a.status || "Active").localeCompare(b.status || "Active") || (a.name || "").localeCompare(b.name || "");
      if (sortBy === "department") return (a.department || "").localeCompare(b.department || "") || (a.name || "").localeCompare(b.name || "");
      return (a.name || "").localeCompare(b.name || "");
    });
  const activeStaff = staff.filter((s) => (s.status || "Active") !== "Inactive");
  const costByGrant = personnelCostByGrant(activeStaff);
  const costByCostCenter = personnelCostByCostCenter(activeStaff);
  const totalPersonnelCost = activeStaff.reduce((a, s) => a + staffAnnualCost(s), 0);

  const saveStaff = (s) => {
    setStaff((prev) => {
      const exists = prev.some((x) => x.id === s.id);
      logActivity?.("Staff", exists ? "Updated" : "Created", s.name || "Untitled staff member");
      return exists ? prev.map((x) => (x.id === s.id ? s : x)) : [...prev, s];
    });
    setModal(null);
  };
  const deleteStaff = (id) => {
    const s = staff.find((x) => x.id === id);
    pushTrash(setTrash, "staff", s, currentUserEmail);
    setStaff((prev) => prev.filter((s) => s.id !== id));
    logActivity?.("Staff", "Deleted", s?.name || "Untitled staff member");
    setConfirm(null);
  };

  return (
    <div className="space-y-4">
      <div className="flex items-center justify-between">
        <div>
          <h1 className="font-display text-2xl" style={{ color: "#1C2624" }}>Personnel & Payroll</h1>
          <p className="text-sm mt-1" style={{ color: "#5B6B66" }}>Staff cost and grant allocation — separate from budget line items</p>
        </div>
        {canEdit && (
          <button onClick={() => setModal("new")} className="inline-flex items-center gap-1.5 px-3 py-2 rounded-md text-sm text-white" style={{ background: "#1F5C6B" }}>
            <Plus size={16} /> New staff member
          </button>
        )}
      </div>

      <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
        <StatCard label="Total staff" value={activeStaff.length} />
        <StatCard label="Total annual personnel cost" value={fmt(totalPersonnelCost)} />
        <StatCard label="Grants with allocated staff" value={Object.keys(costByGrant).length} />
      </div>

      {Object.keys(costByGrant).length > 0 && (
        <div className="bg-white rounded-lg border p-4" style={{ borderColor: "#E1E5DE" }}>
          <h2 className="font-display text-base mb-3" style={{ color: "#1C2624" }}>Personnel cost by grant</h2>
          <div className="space-y-1.5 text-sm">
            {Object.entries(costByGrant).map(([grantId, cost]) => {
              const g = grants.find((x) => x.id === grantId);
              return (
                <div key={grantId} className="flex items-center justify-between">
                  <span style={{ color: "#1C2624" }}>{g ? (g.programCode ? `${g.programCode} - ${g.title}` : g.title) : "Unknown grant"}</span>
                  <span style={{ color: "#1C2624", fontVariantNumeric: "tabular-nums" }}>{fmt(cost)}</span>
                </div>
              );
            })}
          </div>
        </div>
      )}

      {Object.keys(costByCostCenter).length > 0 && (
        <div className="bg-white rounded-lg border p-5" style={{ borderColor: "#E1E5DE" }}>
          <h2 className="font-display text-base mb-3" style={{ color: "#1C2624" }}>Personnel cost by cost center</h2>
          <div className="space-y-1.5 text-sm">
            {Object.entries(costByCostCenter).map(([ccId, cost]) => {
              const cc = (costCenters || []).find((x) => x.id === ccId);
              return (
                <div key={ccId} className="flex items-center justify-between">
                  <span style={{ color: "#1C2624" }}>{cc ? cc.name : "Unknown cost center"}</span>
                  <span style={{ color: "#1C2624", fontVariantNumeric: "tabular-nums" }}>{fmt(cost)}</span>
                </div>
              );
            })}
          </div>
        </div>
      )}

      <div className="flex gap-3 flex-wrap">
        <Field label="Filter by department">
          <select value={deptFilter} onChange={(e) => setDeptFilter(e.target.value)} className={inputCls} style={{ ...inputStyle, maxWidth: 260 }}>
            {departments.map((d) => <option key={d}>{d}</option>)}
          </select>
        </Field>
        <Field label="Filter by site">
          <select value={siteFilter} onChange={(e) => setSiteFilter(e.target.value)} className={inputCls} style={{ ...inputStyle, maxWidth: 260 }}>
            <option>All</option>
            {SITE_OPTIONS.map((s) => <option key={s}>{s}</option>)}
          </select>
        </Field>
        <Field label="Filter by status">
          <select value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)} className={inputCls} style={{ ...inputStyle, maxWidth: 200 }}>
            <option>All</option>
            {STAFF_STATUSES.map((s) => <option key={s}>{s}</option>)}
          </select>
        </Field>
        <Field label="Sort by">
          <select value={sortBy} onChange={(e) => setSortBy(e.target.value)} className={inputCls} style={{ ...inputStyle, maxWidth: 200 }}>
            <option value="name">Name</option>
            <option value="status">Status</option>
            <option value="department">Department</option>
          </select>
        </Field>
      </div>

      {visible.length === 0 ? (
        <div className="bg-white rounded-lg border p-10 text-center" style={{ borderColor: "#E1E5DE", color: "#8A8F87" }}>
          No staff members yet.
        </div>
      ) : (
        <div className="space-y-3">
          {visible.map((s) => {
            const cost = staffAnnualCost(s);
            const pct = staffAllocatedTotal(s);
            const status = s.status || "Active";
            const statusColor = status === "Active" ? "#2F6F53" : status === "Leave of Absence" ? "#C08A2E" : "#8A8F87";
            return (
              <div key={s.id} className="bg-white rounded-lg border p-4" style={{ borderColor: "#E1E5DE" }}>
                <div className="flex items-center justify-between">
                  <div>
                    <div className="flex items-center gap-2">
                      <span className="font-medium" style={{ color: "#1C2624" }}>{s.name}</span>
                      <span className="text-xs px-1.5 py-0.5 rounded-full" style={{ background: `${statusColor}1A`, color: statusColor }}>{status}</span>
                    </div>
                    <div className="text-xs mt-0.5" style={{ color: "#8A8F87" }}>{s.position}{s.department ? ` · ${s.department}` : ""}{s.site ? ` · ${s.site}` : ""} · {s.exempt}</div>
                  </div>
                  <div className="text-right text-sm" style={{ fontVariantNumeric: "tabular-nums" }}>
                    <div style={{ color: "#1C2624" }}>{fmt(cost)}/yr</div>
                    <div style={{ color: pct > 100 ? "#B5443A" : "#8A8F87" }}>{pct}% allocated</div>
                  </div>
                </div>
                {s.allocations?.length > 0 && (
                  <div className="flex flex-wrap gap-1 mt-2">
                    {s.allocations.filter((a) => a.grantId).map((a) => {
                      const g = grants.find((x) => x.id === a.grantId);
                      return <Badge key={a.id} color="#5B7FA6">{g ? g.title : "Unknown"} · {a.percent}%</Badge>;
                    })}
                    {s.allocations.filter((a) => a.costCenterId).map((a) => {
                      const cc = (costCenters || []).find((x) => x.id === a.costCenterId);
                      return <Badge key={a.id} color="#8A8F87">{cc ? cc.name : "Unknown"} · {a.percent}%</Badge>;
                    })}
                  </div>
                )}
                <div className="flex gap-2 mt-3">
                  <button onClick={() => setModal(s)} className="inline-flex items-center gap-1 text-xs px-3 py-1.5 rounded-md border" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>
                    <Pencil size={13} /> {canEdit ? "Edit" : "View"}
                  </button>
                  {canEdit && (
                    <button onClick={() => setConfirm(s.id)} className="inline-flex items-center gap-1 text-xs px-3 py-1.5 rounded-md border" style={{ borderColor: "#E1E5DE", color: "#B5443A" }}>
                      <Trash2 size={13} /> Delete
                    </button>
                  )}
                </div>
              </div>
            );
          })}
        </div>
      )}

      {modal && (
        <StaffModal
          staff={modal === "new" ? null : modal}
          grants={grants}
          costCenters={costCenters}
          canEdit={canEdit}
          onSave={saveStaff}
          onClose={() => setModal(null)}
          onDelete={modal === "new" || !canEdit ? undefined : () => { setConfirm(modal.id); setModal(null); }}
        />
      )}
      {confirm && (
        <ConfirmModal message="This will permanently delete this staff member and their allocations." onConfirm={() => deleteStaff(confirm)} onCancel={() => setConfirm(null)} />
      )}
    </div>
  );
}

// ---------- invoicing ----------

function InvoiceModal({ invoice, grants, canEdit = true, onSave, onClose, onDelete }) {
  const grantDefault = grants[0]?.id || "";
  const [form, setForm] = useState(invoice || {
    id: uid(), grantId: grantDefault, invoiceNumber: "", amount: 0,
    submittedDate: "", dueDate: "", paidDate: "", status: "Draft", notes: "",
  });
  const set = (k) => (e) => setForm({ ...form, [k]: e.target.value });

  return (
    <Modal title={invoice ? (canEdit ? "Edit invoice" : "View invoice") : "New invoice"} onClose={onClose}>
      <fieldset disabled={!canEdit} style={{ border: "none", margin: 0, padding: 0 }}>
      <div className="space-y-4">
        <Field label="Grant">
          <GrantPicker grants={grants} value={form.grantId} onChange={(v) => setForm({ ...form, grantId: v })} placeholder="Select a grant" />
        </Field>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          <Field label="Invoice number">
            <input className={inputCls} style={inputStyle} value={form.invoiceNumber} onChange={set("invoiceNumber")} placeholder="e.g. INV-0142" />
          </Field>
          <Field label="Amount">
            <input type="number" className={inputCls} style={inputStyle} value={form.amount} onChange={set("amount")} />
          </Field>
        </div>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          <Field label="Status">
            <select className={inputCls} style={inputStyle} value={form.status} onChange={set("status")}>
              {INVOICE_STATUSES.map((s) => <option key={s}>{s}</option>)}
            </select>
          </Field>
          <Field label="Submitted date">
            <input type="date" className={inputCls} style={inputStyle} value={form.submittedDate} onChange={set("submittedDate")} />
          </Field>
        </div>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          <Field label="Expected payment date">
            <input type="date" className={inputCls} style={inputStyle} value={form.dueDate} onChange={set("dueDate")} />
          </Field>
          <Field label="Date paid">
            <input type="date" className={inputCls} style={inputStyle} value={form.paidDate} onChange={set("paidDate")} />
          </Field>
        </div>
        <Field label="Notes">
          <textarea className={inputCls} style={inputStyle} rows={3} value={form.notes} onChange={set("notes")} placeholder="Submission method, contact, follow-up notes…" />
        </Field>
      </div>
      </fieldset>

      <div className="flex justify-between gap-2 mt-6">
        {onDelete ? (
          <button onClick={onDelete} className="px-4 py-2 rounded-md text-sm border" style={{ borderColor: "#E1E5DE", color: "#B5443A" }}>Delete invoice</button>
        ) : <span />}
        <div className="flex gap-2">
          <button onClick={onClose} className="px-4 py-2 rounded-md text-sm border" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>Cancel</button>
          {canEdit && (
            <button
              onClick={() => { if (!form.grantId) return; onSave(form); }}
              className="px-4 py-2 rounded-md text-sm text-white"
              style={{ background: "#1F5C6B" }}
            >
              Save invoice
            </button>
          )}
        </div>
      </div>
    </Modal>
  );
}

function InvoicingView({ grants, invoices, setInvoices, setTrash, currentUserEmail, canEdit, initialOpenInvoiceId, logActivity }) {
  const [modal, setModal] = useState(() => (initialOpenInvoiceId ? invoices.find((i) => i.id === stripNonce(initialOpenInvoiceId)) || null : null));
  const [confirm, setConfirm] = useState(null);
  const [grantFilter, setGrantFilter] = useState("All");
  const [statusFilter, setStatusFilter] = useState("All");

  const visible = invoices
    .filter((i) => grantFilter === "All" || i.grantId === grantFilter)
    .filter((i) => statusFilter === "All" || i.status === statusFilter)
    .sort((a, b) => new Date(b.submittedDate || 0) - new Date(a.submittedDate || 0));

  const totalInvoiced = invoices.reduce((a, i) => a + (Number(i.amount) || 0), 0);
  const totalPaid = invoices.filter((i) => i.status === "Paid").reduce((a, i) => a + (Number(i.amount) || 0), 0);
  const totalOutstanding = invoices.filter((i) => i.status === "Submitted").reduce((a, i) => a + (Number(i.amount) || 0), 0);
  const overdueCount = invoices.filter(isInvoiceOverdue).length;

  const saveInvoice = (inv) => {
    setInvoices((prev) => {
      const exists = prev.some((x) => x.id === inv.id);
      logActivity?.("Invoice", exists ? "Updated" : "Created", inv.invoiceNumber || "Untitled invoice");
      return exists ? prev.map((x) => (x.id === inv.id ? inv : x)) : [...prev, inv];
    });
    setModal(null);
  };
  const deleteInvoice = (id) => {
    const inv = invoices.find((x) => x.id === id);
    pushTrash(setTrash, "invoice", inv, currentUserEmail);
    setInvoices((prev) => prev.filter((i) => i.id !== id));
    logActivity?.("Invoice", "Deleted", inv?.invoiceNumber || "Untitled invoice");
    setConfirm(null);
  };
  const exportCsv = () => {
    const rows = [["Grant", "Invoice #", "Amount", "Status", "Submitted", "Due", "Paid", "Days outstanding"]];
    invoices.forEach((i) => {
      const g = grants.find((x) => x.id === i.grantId);
      rows.push([g?.title || "", i.invoiceNumber, i.amount, i.status, i.submittedDate, i.dueDate, i.paidDate, daysOutstanding(i) ?? ""]);
    });
    const csv = rows.map((r) => r.map((c) => `"${String(c).replace(/"/g, '""')}"`).join(",")).join("\n");
    downloadFile("nations-finest-invoices.csv", csv, "text/csv");
  };

  const statusColor = { Draft: "#8A8F87", Submitted: "#5B7FA6", Paid: "#2F6F53", Rejected: "#B5443A" };

  return (
    <div className="space-y-4">
      <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3">
        <div>
          <h1 className="font-display text-2xl" style={{ color: "#1C2624" }}>Invoicing</h1>
          <p className="text-sm mt-1" style={{ color: "#5B6B66" }}>Track invoices submitted to funders and what's still outstanding</p>
        </div>
        <div className="flex items-center gap-2 flex-wrap">
          <a href="https://pmsapp.psc.gov/pms/app/login" target="_blank" rel="noreferrer" className="inline-flex items-center gap-1.5 px-3 py-2 rounded-md text-sm border" style={{ borderColor: "#E1E5DE", color: "#1F5C6B" }}>
            <ExternalLink size={15} /> PMS System
          </a>
          <a href="https://authentication.tungsten-network.com/login" target="_blank" rel="noreferrer" className="inline-flex items-center gap-1.5 px-3 py-2 rounded-md text-sm border" style={{ borderColor: "#E1E5DE", color: "#1F5C6B" }}>
            <ExternalLink size={15} /> Tungsten System
          </a>
          <button onClick={exportCsv} className="inline-flex items-center gap-1.5 px-3 py-2 rounded-md text-sm border" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>
            <Download size={15} /> Export CSV
          </button>
          {canEdit && (
            <button onClick={() => setModal("new")} className="inline-flex items-center gap-1.5 px-3 py-2 rounded-md text-sm text-white" style={{ background: "#1F5C6B" }}>
              <Plus size={16} /> New invoice
            </button>
          )}
        </div>
      </div>

      <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
        <StatCard label="Total invoiced" value={fmt(totalInvoiced)} />
        <StatCard label="Total paid" value={fmt(totalPaid)} />
        <StatCard label="Outstanding" value={fmt(totalOutstanding)} sub={`${invoices.filter((i) => i.status === "Submitted").length} submitted, unpaid`} />
        <StatCard label="Overdue" value={overdueCount} sub={overdueCount > 0 ? "Past expected payment date" : "None"} />
      </div>

      <div className="flex gap-3">
        <GrantPicker grants={grants} value={grantFilter === "All" ? "" : grantFilter} onChange={(v) => setGrantFilter(v || "All")} noneLabel="All grants" noneValue="All" wrapStyle={{ width: 260 }} />
        <select value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)} className={inputCls} style={{ ...inputStyle, width: 160 }}>
          <option>All</option>
          {INVOICE_STATUSES.map((s) => <option key={s}>{s}</option>)}
        </select>
      </div>

      {visible.length === 0 ? (
        <div className="bg-white rounded-lg border p-10 text-center" style={{ borderColor: "#E1E5DE", color: "#8A8F87" }}>
          No invoices match your filters.
        </div>
      ) : (
        <div className="overflow-x-auto border rounded-lg bg-white" style={{ borderColor: "#E1E5DE" }}>
          <table className="w-full text-sm">
            <thead>
              <tr style={{ background: "#F6F7F3" }}>
                <th className="text-left px-3 py-2 text-xs">Grant</th>
                <th className="text-left px-3 py-2 text-xs">Invoice #</th>
                <th className="text-right px-3 py-2 text-xs">Amount</th>
                <th className="text-left px-3 py-2 text-xs">Status</th>
                <th className="text-right px-3 py-2 text-xs">Submitted</th>
                <th className="text-right px-3 py-2 text-xs">Expected payment</th>
                <th className="text-right px-3 py-2 text-xs">Paid</th>
                <th className="text-right px-3 py-2 text-xs">Days outstanding</th>
                <th className="px-3 py-2"></th>
              </tr>
            </thead>
            <tbody>
              {visible.map((i) => {
                const g = grants.find((x) => x.id === i.grantId);
                const overdue = isInvoiceOverdue(i);
                const outstanding = daysOutstanding(i);
                return (
                  <tr key={i.id} className="border-t cursor-pointer hover:bg-stone-50" style={{ borderColor: overdue ? "#B5443A" : "#E1E5DE" }} onClick={() => setModal(i)}>
                    <td className="px-3 py-2" style={{ color: "#1C2624" }}>{g?.title || "Unknown grant"}</td>
                    <td className="px-3 py-2" style={{ color: "#1C2624" }}>{i.invoiceNumber || "—"}</td>
                    <td className="px-3 py-2 text-right" style={{ fontVariantNumeric: "tabular-nums", color: "#1C2624" }}>{fmt(i.amount)}</td>
                    <td className="px-3 py-2"><Badge color={statusColor[i.status]}>{i.status}</Badge></td>
                    <td className="px-3 py-2 text-right" style={{ fontVariantNumeric: "tabular-nums", color: "#5B6B66" }}>{fmtDate(i.submittedDate)}</td>
                    <td className="px-3 py-2 text-right" style={{ fontVariantNumeric: "tabular-nums", color: overdue ? "#B5443A" : "#5B6B66" }}>{fmtDate(i.dueDate)}</td>
                    <td className="px-3 py-2 text-right" style={{ fontVariantNumeric: "tabular-nums", color: "#5B6B66" }}>{i.paidDate ? fmtDate(i.paidDate) : "—"}</td>
                    <td className="px-3 py-2 text-right" style={{ fontVariantNumeric: "tabular-nums", color: overdue ? "#B5443A" : "#5B6B66" }}>{outstanding !== null ? `${outstanding}d` : "—"}</td>
                    <td className="px-3 py-2 text-center" onClick={(e) => e.stopPropagation()}>
                      {canEdit && (
                        <button onClick={() => setConfirm(i.id)} className="p-1 rounded hover:bg-red-50">
                          <Trash2 size={14} style={{ color: "#B5443A" }} />
                        </button>
                      )}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

      {modal && (
        <InvoiceModal
          invoice={modal === "new" ? null : modal}
          grants={grants}
          canEdit={canEdit}
          onSave={saveInvoice}
          onClose={() => setModal(null)}
          onDelete={modal === "new" || !canEdit ? undefined : () => { setConfirm(modal.id); setModal(null); }}
        />
      )}
      {confirm && (
        <ConfirmModal message="This will permanently delete this invoice record." onConfirm={() => deleteInvoice(confirm)} onCancel={() => setConfirm(null)} />
      )}
    </div>
  );
}

// ---------- burn rate ----------

const paceColor = {
  "On pace": "#2F6F53", "Ahead of pace": "#B5443A", "Behind pace": "#C08A2E",
  "Not started": "#8A8F87", "No budget period set": "#8A8F87",
};

function BurnRateView({ grants, budgets }) {
  const withBudgets = grants.filter((g) => budgets.some((b) => b.grantId === g.id));

  return (
    <div className="space-y-4">
      <div>
        <h1 className="font-display text-2xl" style={{ color: "#1C2624" }}>Burn Rate</h1>
        <p className="text-sm mt-1" style={{ color: "#5B6B66" }}>
          Uses recorded actuals where you've entered them (Budgets → Actual). Falls back to the planned schedule for any grant without actuals yet.
        </p>
      </div>

      {withBudgets.length === 0 ? (
        <div className="bg-white rounded-lg border p-10 text-center" style={{ borderColor: "#E1E5DE", color: "#8A8F87" }}>
          No grants with budgets yet — add a budget to see burn rate.
        </div>
      ) : (
        <div className="overflow-x-auto border rounded-lg bg-white" style={{ borderColor: "#E1E5DE" }}>
          <table className="w-full text-sm">
            <thead>
              <tr style={{ background: "#F6F7F3" }}>
                <th className="text-left px-3 py-2 text-xs" style={{ minWidth: 200 }}>Grant</th>
                <th className="text-right px-3 py-2 text-xs">Award</th>
                <th className="text-right px-3 py-2 text-xs">Budgeted (full term)</th>
                <th className="text-right px-3 py-2 text-xs">Planned to date</th>
                <th className="text-right px-3 py-2 text-xs">Actual to date</th>
                <th className="text-right px-3 py-2 text-xs">Variance</th>
                <th className="text-right px-3 py-2 text-xs">% time elapsed</th>
                <th className="text-right px-3 py-2 text-xs">% budget used</th>
                <th className="text-left px-3 py-2 text-xs">Pace</th>
                <th className="text-right px-3 py-2 text-xs">Projected full-term spend</th>
              </tr>
            </thead>
            <tbody>
              {withBudgets.map((g) => {
                const b = grantBurn(g, budgets);
                return (
                  <tr key={g.id} className="border-t" style={{ borderColor: "#E1E5DE" }}>
                    <td className="px-3 py-2" style={{ color: "#1C2624" }}>
                      {g.title}
                      <div className="text-xs" style={{ color: "#8A8F87" }}>{g.programCode}</div>
                    </td>
                    <td className="px-3 py-2 text-right" style={{ fontVariantNumeric: "tabular-nums", color: "#1C2624" }}>{fmt(b.award)}</td>
                    <td className="px-3 py-2 text-right" style={{ fontVariantNumeric: "tabular-nums", color: "#1C2624" }}>{fmt(b.totalExpense)}</td>
                    <td className="px-3 py-2 text-right" style={{ fontVariantNumeric: "tabular-nums", color: "#5B6B66" }}>{b.elapsedKnown ? fmt(b.toDate) : "—"}</td>
                    <td className="px-3 py-2 text-right" style={{ fontVariantNumeric: "tabular-nums", color: b.hasActuals ? "#1C2624" : "#8A8F87" }}>{b.hasActuals ? fmt(b.actualToDate) : "Not entered"}</td>
                    <td className="px-3 py-2 text-right" style={{ fontVariantNumeric: "tabular-nums", color: b.hasActuals ? (b.variance > 0 ? "#B5443A" : "#2F6F53") : "#8A8F87" }}>{b.hasActuals ? fmt(b.variance) : "—"}</td>
                    <td className="px-3 py-2 text-right" style={{ fontVariantNumeric: "tabular-nums", color: "#5B6B66" }}>{b.elapsedKnown ? `${Math.round(b.pctTimeElapsed * 100)}%` : "—"}</td>
                    <td className="px-3 py-2 text-right" style={{ fontVariantNumeric: "tabular-nums", color: "#5B6B66" }}>{b.elapsedKnown ? `${Math.round(b.pctBudgetUsed * 100)}%` : "—"}</td>
                    <td className="px-3 py-2"><Badge color={paceColor[b.status]}>{b.status}</Badge></td>
                    <td className="px-3 py-2 text-right" style={{ fontVariantNumeric: "tabular-nums", color: b.projectedOverAward ? "#B5443A" : "#1C2624" }}>
                      {b.elapsedKnown ? fmt(b.projectedFullYear) : "—"}
                      {b.projectedOverAward && <div className="text-xs" style={{ color: "#B5443A" }}>Over award</div>}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

// ---------- activity log ----------

function WhoamiModal({ current, onSave, onSkip }) {
  const [name, setName] = useState(current || "");
  return (
    <Modal title="Who are you?" onClose={onSkip}>
      <p className="text-sm mb-4" style={{ color: "#5B6B66" }}>
        This tags the changes you make so your team can see who did what in the Activity Log. There's no password — it's just a label, stored in your own browser.
      </p>
      <Field label="Your name">
        <input className={inputCls} style={inputStyle} value={name} onChange={(e) => setName(e.target.value)} placeholder="e.g. Sarah Johnson" autoFocus />
      </Field>
      <div className="flex justify-end gap-2 mt-6">
        <button onClick={onSkip} className="px-4 py-2 rounded-md text-sm border" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>Skip for now</button>
        <button
          onClick={() => { if (name.trim()) onSave(name.trim()); }}
          className="px-4 py-2 rounded-md text-sm text-white"
          style={{ background: "#1F5C6B" }}
        >
          Save
        </button>
      </div>
    </Modal>
  );
}

function ActivityLogView({ activity }) {
  const [entityFilter, setEntityFilter] = useState("All");
  const [personFilter, setPersonFilter] = useState("All");
  const entities = ["All", "Grant", "Budget", "Cost Center", "Budget Group", "Scenario", "Report", "Staff", "Task", "Invoice", "Data"];
  const people = ["All", ...new Set(activity.map((a) => a.by).filter(Boolean))];
  const visible = activity
    .filter((a) => entityFilter === "All" || a.entity === entityFilter)
    .filter((a) => personFilter === "All" || a.by === personFilter);
  const actionColor = { Created: "#2F6F53", Updated: "#5B7FA6", Deleted: "#B5443A" };
  const fmtWhen = (iso) => new Date(iso).toLocaleString("en-US", { month: "short", day: "numeric", year: "numeric", hour: "numeric", minute: "2-digit" });

  return (
    <div className="space-y-4">
      <div>
        <h1 className="font-display text-2xl" style={{ color: "#1C2624" }}>Activity Log</h1>
        <p className="text-sm mt-1" style={{ color: "#5B6B66" }}>A record of changes made across the team</p>
      </div>

      <div className="flex gap-3">
        <Field label="Filter by type">
          <select value={entityFilter} onChange={(e) => setEntityFilter(e.target.value)} className={inputCls} style={{ ...inputStyle, maxWidth: 220 }}>
            {entities.map((e) => <option key={e}>{e}</option>)}
          </select>
        </Field>
        <Field label="Filter by person">
          <select value={personFilter} onChange={(e) => setPersonFilter(e.target.value)} className={inputCls} style={{ ...inputStyle, maxWidth: 220 }}>
            {people.map((p) => <option key={p}>{p}</option>)}
          </select>
        </Field>
      </div>

      {visible.length === 0 ? (
        <div className="bg-white rounded-lg border p-10 text-center" style={{ borderColor: "#E1E5DE", color: "#8A8F87" }}>
          No activity recorded yet.
        </div>
      ) : (
        <div className="bg-white rounded-lg border divide-y" style={{ borderColor: "#E1E5DE" }}>
          {visible.map((a) => (
            <div key={a.id} className="px-4 py-3 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-1 text-sm">
              <div className="flex items-center gap-2 flex-wrap">
                <Badge color={actionColor[a.action] || "#8A8F87"}>{a.action}</Badge>
                <Badge color="#5B6B66">{a.entity}</Badge>
                <span style={{ color: "#1C2624" }}>{a.label}</span>
              </div>
              <div className="text-xs shrink-0" style={{ color: "#8A8F87" }}>
                {a.by ? <span style={{ color: "#5B6B66" }}>{a.by} · </span> : null}
                {fmtWhen(a.timestamp)}
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

// ---------- global search ----------

function GlobalSearch({ grants, budgets, reports, staff, tasks, invoices, costCenters, goTo }) {
  const [query, setQuery] = useState("");
  const [open, setOpen] = useState(false);
  const [highlighted, setHighlighted] = useState(0);

  const q = query.trim().toLowerCase();
  const matches = (text) => (text || "").toLowerCase().includes(q);
  const results = q.length < 2 ? [] : [
    ...grants.filter((g) => matches(g.title) || matches(g.programCode) || matches(g.notes))
      .slice(0, 5).map((g) => ({ type: "Grant", label: g.title, sub: g.programCode, action: () => goTo("grants", null, null, g.id) })),
    ...budgets.filter((b) => matches(b.title) || matches(b.notes))
      .slice(0, 5).map((b) => {
        const g = grants.find((x) => x.id === b.grantId);
        const cc = costCenters?.find((x) => x.id === b.costCenterId);
        return { type: "Budget", label: b.title, sub: g?.title || cc?.name || "", action: () => goTo("budgets", null, b.grantId, b.id) };
      }),
    ...reports.filter((r) => matches(r.title) || matches(r.notes))
      .slice(0, 5).map((r) => {
        const g = grants.find((x) => x.id === r.grantId);
        return { type: "Grant Report", label: r.title, sub: g?.title || "", action: () => goTo("grant-reports", null, r.grantId, r.id) };
      }),
    ...staff.filter((s) => matches(s.name) || matches(s.position))
      .slice(0, 5).map((s) => ({ type: "Staff", label: s.name, sub: s.position, action: () => goTo("personnel", null, null, s.id) })),
    ...tasks.filter((t) => matches(t.title) || matches(t.notes))
      .slice(0, 5).map((t) => {
        const g = grants.find((x) => x.id === t.grantId);
        return { type: "Task", label: t.title, sub: g?.title || t.category, action: () => goTo("tasks", null, null, t.id) };
      }),
    ...invoices.filter((i) => matches(i.invoiceNumber) || matches(i.notes))
      .slice(0, 5).map((i) => {
        const g = grants.find((x) => x.id === i.grantId);
        return { type: "Invoice", label: i.invoiceNumber || "Untitled invoice", sub: g?.title || "", action: () => goTo("invoicing", null, null, i.id) };
      }),
  ].slice(0, 12);

  const typeColor = { Grant: "#2F6F53", Budget: "#5B7FA6", "Grant Report": "#A8791F", Staff: "#B5443A", Task: "#8A8F87", Invoice: "#A8791F" };

  const selectResult = (r) => { r.action(); setQuery(""); setOpen(false); setHighlighted(0); };
  const onKeyDown = (e) => {
    if (!open || results.length === 0) return;
    if (e.key === "ArrowDown") { e.preventDefault(); setHighlighted((h) => (h + 1) % results.length); }
    else if (e.key === "ArrowUp") { e.preventDefault(); setHighlighted((h) => (h - 1 + results.length) % results.length); }
    else if (e.key === "Enter") { e.preventDefault(); selectResult(results[highlighted]); }
    else if (e.key === "Escape") { setOpen(false); }
  };

  return (
    <div className="relative w-full md:w-80">
      <div className="relative">
        <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2" style={{ color: "#8A8F87" }} />
        <input
          value={query}
          onChange={(e) => { setQuery(e.target.value); setOpen(true); setHighlighted(0); }}
          onFocus={() => setOpen(true)}
          onBlur={() => setTimeout(() => setOpen(false), 150)}
          onKeyDown={onKeyDown}
          placeholder="Search grants, budgets, reports, staff, tasks…"
          className="w-full rounded-md border pl-8 pr-3 py-1.5 text-sm outline-none"
          style={{ borderColor: "#E1E5DE", color: "#1C2624" }}
        />
      </div>
      {open && q.length >= 2 && (
        <div className="absolute mt-1 w-full bg-white rounded-md border shadow-lg z-50 max-h-80 overflow-y-auto" style={{ borderColor: "#E1E5DE" }}>
          {results.length === 0 ? (
            <div className="px-3 py-3 text-sm" style={{ color: "#8A8F87" }}>No matches</div>
          ) : (
            results.map((r, i) => (
              <button
                key={i}
                onClick={() => selectResult(r)}
                onMouseEnter={() => setHighlighted(i)}
                className="w-full flex items-center justify-between px-3 py-2 text-left border-t first:border-t-0"
                style={{ borderColor: "#E1E5DE", background: highlighted === i ? "#F6F7F3" : "transparent" }}
              >
                <div>
                  <div className="text-sm" style={{ color: "#1C2624" }}>{r.label}</div>
                  {r.sub && <div className="text-xs" style={{ color: "#8A8F87" }}>{r.sub}</div>}
                </div>
                <Badge color={typeColor[r.type]}>{r.type}</Badge>
              </button>
            ))
          )}
        </div>
      )}
    </div>
  );
}

// ---------- data / backup ----------

function trimRowKeys(row) {
  const out = {};
  Object.keys(row).forEach((k) => { out[k.trim()] = row[k]; });
  return out;
}

function getField(row, ...names) {
  for (const n of names) {
    if (row[n] !== undefined && row[n] !== null && String(row[n]).trim() !== "") return row[n];
  }
  return "";
}

function stripSpTokens(val) {
  if (!val) return "";
  return String(val).split(";#").map((t) => t.trim()).filter((t) => t && !/^\d+$/.test(t)).join(", ");
}

function excelDateToStr(val) {
  if (!val) return "";
  if (val instanceof Date && !isNaN(val)) return val.toISOString().slice(0, 10);
  if (typeof val === "number") {
    const d = new Date(Math.round((val - 25569) * 86400 * 1000));
    if (!isNaN(d)) return d.toISOString().slice(0, 10);
  }
  const d = new Date(val);
  return isNaN(d) ? "" : d.toISOString().slice(0, 10);
}

function parseBudgetPeriod(val) {
  if (!val) return { start: "", end: "" };
  const parts = String(val).split(/\s*-\s*/);
  return { start: excelDateToStr(parts[0]), end: excelDateToStr(parts[1]) };
}

function pickStage(val) {
  if (!val) return "Prospecting";
  const tokens = String(val).split(";#").map((t) => t.trim().toLowerCase());
  const priority = ["active", "applied", "awarded", "closing", "prospecting"];
  for (const p of priority) if (tokens.includes(p)) return STAGES.find((s) => s.toLowerCase() === p);
  if (tokens.includes("inactive")) return "Closed";
  return STAGES.find((s) => s.toLowerCase() === tokens[0]) || "Prospecting";
}

function pickCadences(val) {
  if (!val) return { matched: [], leftover: "" };
  const tokens = String(val).split(",").map((t) => t.trim()).filter(Boolean);
  const matched = [];
  const leftover = [];
  tokens.forEach((t) => {
    const low = t.toLowerCase();
    if (low.includes("week")) matched.push("Weekly");
    else if (low.includes("month")) matched.push("Monthly");
    else if (low.includes("quarter")) matched.push("Quarterly");
    else if (low.includes("semi")) matched.push("Semi-annual");
    else if (low.includes("year") || low.includes("annual")) matched.push("Annually");
    else if (low.includes("end of")) matched.push("End of grant");
    else leftover.push(t);
  });
  return { matched: [...new Set(matched)], leftover: leftover.join(", ") };
}

function DataView({ grants, budgets, reports, staff, tasks, activity, invoices, costCenters, budgetGroups, scenarios, trash, setGrants, setBudgets, setReports, setStaff, setTasks, setActivity, setInvoices, setCostCenters, setBudgetGroups, setScenarios, setTrash, canEdit, logActivity }) {
  const [restoreError, setRestoreError] = useState("");
  const [restoreSummary, setRestoreSummary] = useState("");
  const [importError, setImportError] = useState("");
  const [importSummary, setImportSummary] = useState("");
  const [budgetImportError, setBudgetImportError] = useState("");
  const [budgetImportSummary, setBudgetImportSummary] = useState("");
  const [reportImportError, setReportImportError] = useState("");
  const [reportImportSummary, setReportImportSummary] = useState("");

  const downloadBackup = () => {
    const payload = { exportedAt: new Date().toISOString(), grants, budgets, reports, staff, tasks, invoices, costCenters, budgetGroups, scenarios, trash, activity };
    downloadFile(`nations-finest-grantflow-backup-${new Date().toISOString().slice(0, 10)}.json`, JSON.stringify(payload, null, 2), "application/json");
  };

  const [showBackupText, setShowBackupText] = useState(false);
  const [autoBackupList, setAutoBackupList] = useState(null);
  const [autoBackupLoading, setAutoBackupLoading] = useState(false);
  const [autoBackupError, setAutoBackupError] = useState("");

  const loadAutoBackups = async () => {
    setAutoBackupLoading(true);
    setAutoBackupError("");
    try {
      const list = await window.storage.list("grantflow:autobackup:", true);
      const dates = (list?.keys || [])
        .filter((k) => k !== "grantflow:autobackup:meta")
        .map((k) => k.replace("grantflow:autobackup:", ""))
        .sort()
        .reverse();
      setAutoBackupList(dates);
    } catch (err) {
      setAutoBackupError("Couldn't load the list of automatic backups.");
    }
    setAutoBackupLoading(false);
  };

  const downloadAutoBackup = async (date) => {
    try {
      const res = await window.storage.get(`grantflow:autobackup:${date}`, true);
      if (!res?.value) { setAutoBackupError(`Couldn't find the backup for ${date}.`); return; }
      downloadFile(`nations-finest-grantflow-autobackup-${date}.json`, res.value, "application/json");
    } catch (err) {
      setAutoBackupError(`Couldn't download the backup for ${date}.`);
    }
  };
  const [copyStatus, setCopyStatus] = useState("");
  const backupText = useMemo(() => {
    const payload = { exportedAt: new Date().toISOString(), grants, budgets, reports, staff, tasks, invoices, costCenters, budgetGroups, scenarios, trash, activity };
    return JSON.stringify(payload, null, 2);
  }, [grants, budgets, reports, staff, tasks, invoices, costCenters, budgetGroups, scenarios, trash, activity]);

  const copyBackupText = async () => {
    try {
      await navigator.clipboard.writeText(backupText);
      setCopyStatus("Copied to clipboard.");
    } catch (err) {
      setCopyStatus("Couldn't auto-copy — select the text below manually and copy it (Ctrl/Cmd+C).");
    }
  };

  const handleRestore = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setRestoreError(""); setRestoreSummary("");
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const data = JSON.parse(reader.result);
        if (!data || typeof data !== "object") throw new Error("Not a valid backup file");
        if (Array.isArray(data.grants)) setGrants(data.grants);
        if (Array.isArray(data.budgets)) setBudgets(data.budgets);
        if (Array.isArray(data.reports)) setReports(data.reports);
        if (Array.isArray(data.staff)) setStaff(data.staff);
        if (Array.isArray(data.tasks)) setTasks(data.tasks);
        if (Array.isArray(data.invoices)) setInvoices(data.invoices);
        if (Array.isArray(data.costCenters)) setCostCenters(data.costCenters);
        if (Array.isArray(data.budgetGroups)) setBudgetGroups(data.budgetGroups);
        if (Array.isArray(data.scenarios)) setScenarios(data.scenarios);
        if (Array.isArray(data.trash)) setTrash(data.trash);
        if (Array.isArray(data.activity)) setActivity(data.activity);
        logActivity?.("Data", "Restored", `Restored from backup file "${file.name}"`);
        setRestoreSummary(`Restored ${data.grants?.length || 0} grants, ${data.budgets?.length || 0} budgets, ${data.reports?.length || 0} reports, ${data.staff?.length || 0} staff, ${data.tasks?.length || 0} tasks, ${data.invoices?.length || 0} invoices, ${data.costCenters?.length || 0} cost centers, ${data.budgetGroups?.length || 0} budget groups, ${data.scenarios?.length || 0} scenarios.`);
      } catch (err) {
        setRestoreError("Couldn't read that file — make sure it's a GrantFlow backup JSON exported from this app.");
      }
    };
    reader.readAsText(file);
    e.target.value = "";
  };

  const handleImportGrants = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setImportError(""); setImportSummary("");
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const wb = XLSX.read(reader.result, { type: "binary", cellDates: true });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" }).map(trimRowKeys);
        if (rows.length === 0) throw new Error("No rows found");
        const imported = rows.map((row) => {
          const title = String(getField(row, "Title", "title", "Grant title")).trim();
          if (!title) return null;

          const period = parseBudgetPeriod(getField(row, "Budget period"));
          const cadenceInfo = pickCadences(getField(row, "Reporting Interval"));
          const extraNotes = [
            getField(row, "Contract #") ? `Contract #: ${getField(row, "Contract #")}` : "",
            getField(row, "Reporting Requirements") ? `Reporting requirements: ${getField(row, "Reporting Requirements")}` : "",
            getField(row, "Performance Metrics/Deliverables") ? `Performance metrics/deliverables: ${getField(row, "Performance Metrics/Deliverables")}` : "",
            getField(row, "Reporting Center") ? `Reporting center: ${stripSpTokens(getField(row, "Reporting Center"))}` : "",
            getField(row, "Payor Systems") ? `Payor system: ${getField(row, "Payor Systems")}` : "",
            getField(row, "Payment Interval") ? `Payment interval: ${getField(row, "Payment Interval")}` : "",
            cadenceInfo.leftover ? `Reporting interval (unmapped): ${cadenceInfo.leftover}` : "",
          ].filter(Boolean).join("\n");

          const paymentMethodRaw = String(getField(row, "Payment Method to NF")).trim();
          const paymentMethod = PAYMENT_METHODS.find((p) => p.toLowerCase() === paymentMethodRaw.toLowerCase()) || PAYMENT_METHODS[0];

          const bedRateRaw = String(getField(row, "Rate")).replace(/[^0-9.]/g, "");

          return {
            id: uid(),
            title,
            programCode: String(getField(row, "Program Code", "programCode", "Program code")).trim(),
            funding: String(getField(row, "Funding source", "funding")).trim(),
            sites: getField(row, "Site(s)", "Sites", "sites")
              ? String(getField(row, "Site(s)", "Sites", "sites")).split(";#").map((s) => s.trim()).filter(Boolean)
              : [],
            stage: pickStage(getField(row, "Status", "Stage", "stage")),
            awardAmount: Number(getField(row, "Award amount", "awardAmount")) || 0,
            awardAmountRemaining: Number(getField(row, "Award amount remaining")) || 0,
            start: excelDateToStr(getField(row, "Start date", "Start", "start")),
            end: excelDateToStr(getField(row, "End date", "End", "end")),
            riskStatus: RISKS.includes(getField(row, "Risk status", "riskStatus")) ? getField(row, "Risk status", "riskStatus") : "Low",
            cadence: cadenceInfo.matched,
            complianceOwner: String(getField(row, "Compliance owner")).trim(),
            financeOwner: String(getField(row, "Finance owner")).trim(),
            internalOwner: stripSpTokens(getField(row, "NF POC", "Internal owner")),
            operationsOwner: String(getField(row, "Operations owner")).trim(),
            renewal: false, doclibUrl: "", contractUrl: "",
            notes: extraNotes,
            budgetPeriodStart: period.start, budgetPeriodEnd: period.end,
            obligatedFunds: Number(getField(row, "Obligated funds")) || 0,
            obligatedFundsRemaining: Number(getField(row, "Obligated funds remaining")) || 0,
            paymentMethod,
            beds: getField(row, "Beds") || "",
            bedRate: Number(bedRateRaw) || 0,
            grantPoc: String(getField(row, "Grant POC", "grantPoc")).trim(),
          };
        }).filter(Boolean);
        if (imported.length === 0) throw new Error("No valid rows with a Title column found");
        setGrants((prev) => [...prev, ...imported]);
        logActivity?.("Data", "Imported", `Imported ${imported.length} grants from "${file.name}"`);
        setImportSummary(`Imported ${imported.length} grant${imported.length > 1 ? "s" : ""}. Fields with no clear match on your sheet default to blank — open each grant to check the details.`);
      } catch (err) {
        setImportError("Couldn't read that file. Use a .csv or .xlsx with a header row — at minimum a \"Title\" column.");
      }
    };
    reader.readAsBinaryString(file);
    e.target.value = "";
  };

  const handleImportBudgets = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setBudgetImportError(""); setBudgetImportSummary("");
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const wb = XLSX.read(reader.result, { type: "binary", cellDates: true });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" }).map(trimRowKeys);
        if (rows.length === 0) throw new Error("No rows found");

        const findGrant = (name) => {
          const n = String(name || "").trim().toLowerCase();
          if (!n) return null;
          return grants.find((g) => g.title.toLowerCase() === n || (g.programCode && g.programCode.toLowerCase() === n))
            || grants.find((g) => (g.programCode ? `${g.programCode} - ${g.title}` : g.title).toLowerCase() === n);
        };

        const groups = {};
        let skipped = 0;
        rows.forEach((row) => {
          const grantName = getField(row, "Grant", "grant");
          const grant = findGrant(grantName);
          const budgetTitle = String(getField(row, "Budget", "Budget Title", "budget") || "Imported budget").trim();
          const category = String(getField(row, "Category", "category")).trim();
          if (!grant || !category) { skipped++; return; }
          const catDef = CATEGORIES.find((c) => c.name.toLowerCase() === category.toLowerCase());
          const key = `${grant.id}::${budgetTitle}`;
          if (!groups[key]) {
            groups[key] = {
              id: uid(), grantId: grant.id, title: budgetTitle,
              fy: String(getField(row, "Fiscal Year", "fy")).trim(),
              periodStart: excelDateToStr(getField(row, "Period Start", "periodStart")),
              periodEnd: excelDateToStr(getField(row, "Period End", "periodEnd")),
              status: BUDGET_STATUSES.includes(getField(row, "Status", "status")) ? getField(row, "Status", "status") : "Draft",
              notes: "", lines: [], approvedBy: "", approvedAt: "", rejectionReason: "",
            };
          }
          const amounts = MONTHS.map((m) => Number(row[m] || 0) || 0);
          groups[key].lines.push({
            id: uid(),
            category: catDef ? catDef.name : category,
            type: catDef ? catDef.type : (String(getField(row, "Type", "type") || "expense").toLowerCase() === "revenue" ? "revenue" : "expense"),
            subcategory: String(getField(row, "Subcategory", "subcategory")).trim(),
            amounts, actuals: Array(12).fill(0),
          });
        });

        const newBudgets = Object.values(groups);
        if (newBudgets.length === 0) throw new Error("No rows matched an existing grant and category");
        setBudgets((prev) => [...prev, ...newBudgets]);
        logActivity?.("Data", "Imported", `Imported ${newBudgets.length} budget(s) from "${file.name}"`);
        setBudgetImportSummary(
          `Imported ${newBudgets.length} budget${newBudgets.length > 1 ? "s" : ""} (${rows.length - skipped} line item${rows.length - skipped === 1 ? "" : "s"}).` +
          (skipped > 0 ? ` Skipped ${skipped} row${skipped > 1 ? "s" : ""} — no matching grant or missing category.` : "")
        );
      } catch (err) {
        setBudgetImportError("Couldn't read that file, or no rows matched an existing grant. Use the same column layout as \"Export all as CSV\" on the Reporting tab: Grant, Budget, Category, Subcategory, Type, Jan…Dec.");
      }
    };
    reader.readAsBinaryString(file);
    e.target.value = "";
  };

  const handleImportReports = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setReportImportError(""); setReportImportSummary("");
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const wb = XLSX.read(reader.result, { type: "binary", cellDates: true });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" }).map(trimRowKeys);
        if (rows.length === 0) throw new Error("No rows found");

        const findGrant = (name) => {
          const n = String(name || "").trim().toLowerCase();
          if (!n) return null;
          return grants.find((g) => g.title.toLowerCase() === n || (g.programCode && g.programCode.toLowerCase() === n))
            || grants.find((g) => (g.programCode ? `${g.programCode} - ${g.title}` : g.title).toLowerCase() === n);
        };

        let skipped = 0;
        const imported = rows.map((row) => {
          const title = String(getField(row, "Title", "title")).trim();
          if (!title) { skipped++; return null; }
          const grant = findGrant(getField(row, "Grant", "grant"));
          const priorityRaw = String(getField(row, "Priority", "priority") || "Medium").trim();
          const priority = REPORT_PRIORITIES.some((p) => p.label.toLowerCase() === priorityRaw.toLowerCase())
            ? REPORT_PRIORITIES.find((p) => p.label.toLowerCase() === priorityRaw.toLowerCase()).label
            : "Medium";
          const statusRaw = String(getField(row, "Status", "status") || "Not started").trim();
          const status = REPORT_STATUSES.some((s) => s.toLowerCase() === statusRaw.toLowerCase())
            ? REPORT_STATUSES.find((s) => s.toLowerCase() === statusRaw.toLowerCase())
            : "Not started";
          const repeatRaw = String(getField(row, "Repeat", "repeat") || "None").trim();
          const repeat = REPORT_REPEATS.some((r) => r.toLowerCase() === repeatRaw.toLowerCase())
            ? REPORT_REPEATS.find((r) => r.toLowerCase() === repeatRaw.toLowerCase())
            : "None";
          return {
            id: uid(), title, grantId: grant ? grant.id : "",
            assignedTo: String(getField(row, "Assigned To", "assignedTo")).trim(),
            status, priority,
            startDate: excelDateToStr(getField(row, "Start Date", "startDate")),
            dueDate: excelDateToStr(getField(row, "Due Date", "dueDate")),
            repeat, repeatDetail: String(getField(row, "Repeat Detail", "repeatDetail")).trim(),
            bucket: String(getField(row, "Bucket", "bucket") || DEFAULT_BUCKETS[0]).trim(),
            checklist: [], notes: String(getField(row, "Notes", "notes")).trim(),
            portalUrl: String(getField(row, "Submission Portal URL", "portalUrl")).trim(),
            linkedTaskCreated: false,
            createdAt: new Date().toISOString().slice(0, 10),
          };
        }).filter(Boolean);

        if (imported.length === 0) throw new Error("No valid rows with a Title column found");
        setReports((prev) => [...prev, ...imported]);
        logActivity?.("Data", "Imported", `Imported ${imported.length} grant report(s) from "${file.name}"`);
        setReportImportSummary(
          `Imported ${imported.length} report${imported.length > 1 ? "s" : ""}.` +
          (skipped > 0 ? ` Skipped ${skipped} row${skipped > 1 ? "s" : ""} with no title.` : "")
        );
      } catch (err) {
        setReportImportError("Couldn't read that file. Use a .csv or .xlsx with a header row — at minimum a \"Title\" column.");
      }
    };
    reader.readAsBinaryString(file);
    e.target.value = "";
  };

  return (
    <div className="space-y-6 max-w-2xl">
      <div>
        <h1 className="font-display text-2xl" style={{ color: "#1C2624" }}>Data & Backup</h1>
        <p className="text-sm mt-1" style={{ color: "#5B6B66" }}>Export a full backup, restore from one, or bulk-import grants from a spreadsheet</p>
      </div>

      <div className="bg-white rounded-lg border p-5" style={{ borderColor: "#E1E5DE" }}>
        <h2 className="font-display text-base mb-1" style={{ color: "#1C2624" }}>Download backup</h2>
        <p className="text-sm mb-3" style={{ color: "#5B6B66" }}>Saves everything — grants, budgets, reports, staff, tasks, and the activity log — to one JSON file you can keep as a safety copy.</p>
        <div className="flex flex-wrap items-center gap-2">
          <button onClick={downloadBackup} className="inline-flex items-center gap-1.5 px-3 py-2 rounded-md text-sm text-white" style={{ background: "#1F5C6B" }}>
            <Download size={15} /> Download backup (.json)
          </button>
          <button
            onClick={() => { setShowBackupText((v) => !v); setCopyStatus(""); }}
            className="inline-flex items-center gap-1.5 px-3 py-2 rounded-md text-sm border"
            style={{ borderColor: "#E1E5DE", color: "#1C2624" }}
          >
            {showBackupText ? "Hide" : "View & copy backup data"}
          </button>
        </div>
        <p className="text-xs mt-2" style={{ color: "#8A8F87" }}>
          If the download button doesn't work in your browser, use "View & copy backup data" instead — it always works, since it's just text on screen you copy and paste into a .json file yourself.
        </p>
        {showBackupText && (
          <div className="mt-3">
            <div className="flex items-center gap-2 mb-2">
              <button onClick={copyBackupText} className="text-xs px-3 py-1.5 rounded-md border" style={{ borderColor: "#E1E5DE", color: "#1F5C6B" }}>
                Copy to clipboard
              </button>
              {copyStatus && <span className="text-xs" style={{ color: "#5B6B66" }}>{copyStatus}</span>}
            </div>
            <textarea
              readOnly value={backupText}
              onFocus={(e) => e.target.select()}
              rows={12}
              className="w-full rounded-md border px-3 py-2 text-xs"
              style={{ ...inputStyle, fontFamily: "var(--mono-font)" }}
            />
            <p className="text-xs mt-1" style={{ color: "#8A8F87" }}>
              Click inside the box to select all, then copy (Ctrl/Cmd+C) and paste into a plain text file saved with a .json extension.
            </p>
          </div>
        )}
      </div>

      <div className="bg-white rounded-lg border p-5" style={{ borderColor: "#E1E5DE" }}>
        <h2 className="font-display text-base mb-1" style={{ color: "#1C2624" }}>Automatic backups</h2>
        <p className="text-sm mb-3" style={{ color: "#5B6B66" }}>
          A backup is taken automatically once a day, the first time anyone opens the app that day — no action needed.
          Kept for {AUTO_BACKUP_RETENTION_DAYS} days, then cleaned up automatically.
        </p>
        <button onClick={loadAutoBackups} disabled={autoBackupLoading} className="text-xs px-3 py-1.5 rounded-md border" style={{ borderColor: "#E1E5DE", color: "#1F5C6B" }}>
          {autoBackupLoading ? "Loading…" : "Show available automatic backups"}
        </button>
        {autoBackupError && <p className="text-sm mt-2" style={{ color: "#B5443A" }}>{autoBackupError}</p>}
        {autoBackupList !== null && (
          autoBackupList.length === 0 ? (
            <p className="text-sm mt-2" style={{ color: "#8A8F87" }}>No automatic backups yet — one will appear after the app has been open on any day.</p>
          ) : (
            <div className="mt-3 divide-y" style={{ borderColor: "#E1E5DE" }}>
              {autoBackupList.map((date) => (
                <div key={date} className="py-2 flex items-center justify-between text-sm">
                  <span style={{ color: "#1C2624" }}>{date}</span>
                  <button onClick={() => downloadAutoBackup(date)} className="text-xs px-3 py-1.5 rounded-md border" style={{ borderColor: "#E1E5DE", color: "#1F5C6B" }}>
                    Download
                  </button>
                </div>
              ))}
            </div>
          )
        )}
      </div>

      <div className="bg-white rounded-lg border p-5" style={{ borderColor: "#E1E5DE" }}>
        <h2 className="font-display text-base mb-1" style={{ color: "#1C2624" }}>Restore from backup</h2>
        <p className="text-sm mb-3" style={{ color: "#5B6B66" }}>
          <strong style={{ color: "#B5443A" }}>This replaces all current data</strong> — grants, budgets, reports, staff, tasks, and activity log — with what's in the file. Since data here is shared, this affects everyone.
        </p>
        {canEdit ? (
          <label className="inline-flex items-center gap-1.5 px-3 py-2 rounded-md text-sm border cursor-pointer" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>
            <Upload size={15} /> Choose backup file
            <input type="file" accept=".json" className="hidden" onChange={handleRestore} />
          </label>
        ) : (
          <p className="text-sm" style={{ color: "#8A8F87" }}>View-only access — restoring is disabled.</p>
        )}
        {restoreSummary && <p className="text-sm mt-2" style={{ color: "#2F6F53" }}>{restoreSummary}</p>}
        {restoreError && <p className="text-sm mt-2" style={{ color: "#B5443A" }}>{restoreError}</p>}
      </div>

      <div className="bg-white rounded-lg border p-5" style={{ borderColor: "#E1E5DE" }}>
        <h2 className="font-display text-base mb-1" style={{ color: "#1C2624" }}>Bulk import grants</h2>
        <p className="text-sm mb-3" style={{ color: "#5B6B66" }}>
          Upload a .csv or .xlsx with a header row. Recognized columns: Title (required), Program Code, Funding source, Sites, Stage, Award amount, Start, End, Risk status, Compliance owner, Finance owner, Internal owner, Operations owner, Grant POC, Notes. Anything else is ignored — new grants are added, existing ones aren't touched.
        </p>
        {canEdit ? (
          <label className="inline-flex items-center gap-1.5 px-3 py-2 rounded-md text-sm border cursor-pointer" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>
            <Upload size={15} /> Choose spreadsheet
            <input type="file" accept=".csv,.xlsx,.xls" className="hidden" onChange={handleImportGrants} />
          </label>
        ) : (
          <p className="text-sm" style={{ color: "#8A8F87" }}>View-only access — importing is disabled.</p>
        )}
        {importSummary && <p className="text-sm mt-2" style={{ color: "#2F6F53" }}>{importSummary}</p>}
        {importError && <p className="text-sm mt-2" style={{ color: "#B5443A" }}>{importError}</p>}
      </div>

      <div className="bg-white rounded-lg border p-5" style={{ borderColor: "#E1E5DE" }}>
        <h2 className="font-display text-base mb-1" style={{ color: "#1C2624" }}>Bulk import budgets</h2>
        <p className="text-sm mb-3" style={{ color: "#5B6B66" }}>
          Upload a .csv or .xlsx with one row per budget line item — the same layout as "Export all as CSV" on the Reporting tab. Columns: <strong>Grant</strong> (must match an existing grant's title or program code exactly), <strong>Budget</strong>, Fiscal Year, Period Start, Period End, Status, <strong>Category</strong> (must match a GrantFlow category), Subcategory, Type, and Jan–Dec. Rows with the same Grant + Budget are grouped into one budget with multiple line items.
        </p>
        {canEdit ? (
          <label className="inline-flex items-center gap-1.5 px-3 py-2 rounded-md text-sm border cursor-pointer" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>
            <Upload size={15} /> Choose spreadsheet
            <input type="file" accept=".csv,.xlsx,.xls" className="hidden" onChange={handleImportBudgets} />
          </label>
        ) : (
          <p className="text-sm" style={{ color: "#8A8F87" }}>View-only access — importing is disabled.</p>
        )}
        {budgetImportSummary && <p className="text-sm mt-2" style={{ color: "#2F6F53" }}>{budgetImportSummary}</p>}
        {budgetImportError && <p className="text-sm mt-2" style={{ color: "#B5443A" }}>{budgetImportError}</p>}
      </div>

      <div className="bg-white rounded-lg border p-5" style={{ borderColor: "#E1E5DE" }}>
        <h2 className="font-display text-base mb-1" style={{ color: "#1C2624" }}>Bulk import grant reports</h2>
        <p className="text-sm mb-3" style={{ color: "#5B6B66" }}>
          Upload a .csv or .xlsx with one row per report. Columns: <strong>Title</strong> (required), Grant (matches an existing grant's title or program code — left unlinked if it doesn't match), Assigned To, Status, Priority, Start Date, Due Date, Repeat, Repeat Detail, Bucket, Submission Portal URL, Notes. New reports are added; existing ones aren't touched.
        </p>
        {canEdit ? (
          <label className="inline-flex items-center gap-1.5 px-3 py-2 rounded-md text-sm border cursor-pointer" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>
            <Upload size={15} /> Choose spreadsheet
            <input type="file" accept=".csv,.xlsx,.xls" className="hidden" onChange={handleImportReports} />
          </label>
        ) : (
          <p className="text-sm" style={{ color: "#8A8F87" }}>View-only access — importing is disabled.</p>
        )}
        {reportImportSummary && <p className="text-sm mt-2" style={{ color: "#2F6F53" }}>{reportImportSummary}</p>}
        {reportImportError && <p className="text-sm mt-2" style={{ color: "#B5443A" }}>{reportImportError}</p>}
      </div>
    </div>
  );
}

// ---------- app shell ----------

const NAV = [
  { key: "dashboard", label: "Dashboard", icon: LayoutDashboard },
  { key: "grants", label: "Grants", icon: FileText },
  { key: "budgets", label: "Budgets", icon: Wallet },
  { key: "invoicing", label: "Invoicing", icon: Receipt },
  { key: "tasks", label: "Tasks", icon: CheckSquare },
  { key: "grant-reports", label: "Grant Reports", icon: ClipboardList },
  { key: "reporting", label: "Reporting", icon: BarChart3 },
  { key: "org-budget", label: "Org Budget", icon: PieChart },
  { key: "scenarios", label: "Scenarios", icon: FlaskConical },
  { key: "burn-rate", label: "Burn Rate", icon: TrendingUp },
  { key: "personnel", label: "Personnel", icon: Users },
  { key: "activity-log", label: "Activity Log", icon: History },
  { key: "trash", label: "Trash", icon: Trash2 },
  { key: "data", label: "Data & Backup", icon: Upload },
];

export default function GrantFlow({ currentUserEmail, isAdmin, userRole, disabledModules, onSignOut } = {}) {
  const canEdit = isAdmin || userRole !== "viewer";
  const hiddenModules = isAdmin ? [] : (disabledModules || []);
  const [tab, setTab] = useState("dashboard");

  useEffect(() => {
    if (hiddenModules.includes(tab)) setTab("dashboard");
  }, [tab, hiddenModules]);
  const [grants, setGrants] = useState([]);
  const [budgets, setBudgets] = useState([]);
  const [reports, setReports] = useState([]);
  const [staff, setStaff] = useState([]);
  const [activity, setActivity] = useState([]);
  const [tasks, setTasks] = useState([]);
  const [invoices, setInvoices] = useState([]);
  const [selectedGrantId, setSelectedGrantId] = useState("");
  const [costCenters, setCostCenters] = useState([]);
  const [budgetGroups, setBudgetGroups] = useState([]);
  const [scenarios, setScenarios] = useState([]);
  const [trash, setTrash] = useState([]);
  const [announcement, setAnnouncement] = useState(null);
  const [announcementDismissed, setAnnouncementDismissed] = useState(false);
  const [selectedCostCenterId, setSelectedCostCenterId] = useState("");
  const [reportsGrantFilter, setReportsGrantFilter] = useState("All");
  const [loaded, setLoaded] = useState(false);
  const [pendingNewGrant, setPendingNewGrant] = useState(false);
  const [pendingNewTask, setPendingNewTask] = useState(false);
  const [pendingExpandGrantId, setPendingExpandGrantId] = useState("");
  const [pendingOpenBudgetId, setPendingOpenBudgetId] = useState("");
  const [pendingOpenReportId, setPendingOpenReportId] = useState("");
  const [pendingOpenTaskId, setPendingOpenTaskId] = useState("");
  const [pendingOpenStaffId, setPendingOpenStaffId] = useState("");
  const [pendingOpenInvoiceId, setPendingOpenInvoiceId] = useState("");
  const [saveErrors, setSaveErrors] = useState({});
  const saveError = Object.keys(saveErrors).length > 0;

  const saveKey = (key, value, label) => {
    let attempt = 0;
    const tryOnce = () => {
      attempt += 1;
      saveData(key, value)
        .then(() => {
          setSaveErrors((prev) => {
            if (!(key in prev)) return prev;
            const next = { ...prev };
            delete next[key];
            return next;
          });
        })
        .catch(() => {
          if (attempt < 2) {
            setTimeout(tryOnce, 6000);
          } else {
            setSaveErrors((prev) => ({ ...prev, [key]: label }));
          }
        });
    };
    tryOnce();
  };
  const [lastSyncedAt, setLastSyncedAt] = useState(null);
  const [mobileNavOpen, setMobileNavOpen] = useState(false);
  const isSyncingRef = useRef(false);
  const [whoami, setWhoami] = useState(null);
  const [whoamiLoaded, setWhoamiLoaded] = useState(false);
  const [editingWhoami, setEditingWhoami] = useState(false);
  const [skippedWhoami, setSkippedWhoami] = useState(false);

  const withTimeout = (promise, ms = 8000) =>
    Promise.race([promise, new Promise((resolve) => setTimeout(() => resolve(undefined), ms))]);

  const refreshAll = async () => {
    isSyncingRef.current = true;
    try {
      const g = await withTimeout(loadData("grantflow:grants"));
      if (g) setGrants(g);
    } catch (e) { /* no data yet */ }
    try {
      const b = await withTimeout(loadData("grantflow:budgets"));
      if (b) setBudgets(b);
    } catch (e) { /* no data yet */ }
    try {
      const r = await withTimeout(loadData("grantflow:reports"));
      if (r) setReports(r);
    } catch (e) { /* no data yet */ }
    try {
      const s = await withTimeout(loadData("grantflow:staff"));
      if (s) setStaff(s);
    } catch (e) { /* no data yet */ }
    try {
      const act = await withTimeout(loadData("grantflow:activity"));
      if (act) setActivity(act.slice(0, 150));
    } catch (e) { /* no data yet */ }
    try {
      const tk = await withTimeout(loadData("grantflow:tasks"));
      if (tk) setTasks(tk);
    } catch (e) { /* no data yet */ }
    try {
      const iv = await withTimeout(loadData("grantflow:invoices"));
      if (iv) setInvoices(iv);
    } catch (e) { /* no data yet */ }
    try {
      const cc = await withTimeout(loadData("grantflow:costcenters"));
      if (cc) setCostCenters(cc);
    } catch (e) { /* no data yet */ }
    try {
      const bg = await withTimeout(loadData("grantflow:budgetgroups"));
      if (bg) setBudgetGroups(bg);
    } catch (e) { /* no data yet */ }
    try {
      const sc = await withTimeout(loadData("grantflow:scenarios"));
      if (sc) setScenarios(sc);
    } catch (e) { /* no data yet */ }
    try {
      const tr = await withTimeout(loadData("grantflow:trash"));
      if (tr) {
        const cutoff = Date.now() - 90 * 24 * 60 * 60 * 1000;
        setTrash(tr.filter((t) => new Date(t.deletedAt).getTime() > cutoff));
      }
    } catch (e) { /* no data yet */ }
    try {
      const annSnap = await withTimeout(getDoc(doc(db, "app_config", "announcement")));
      if (annSnap?.exists?.() && annSnap.data()?.message) {
        setAnnouncement(annSnap.data());
      } else {
        setAnnouncement(null);
      }
    } catch (e) { /* no announcement set */ }
    setLastSyncedAt(Date.now());
    setTimeout(() => { isSyncingRef.current = false; }, 500);
  };

  useEffect(() => {
    (async () => {
      await refreshAll();
      setLoaded(true);
    })();
    (async () => {
      if (currentUserEmail) {
        setWhoami(currentUserEmail);
        setWhoamiLoaded(true);
        return;
      }
      try {
        const w = await window.storage.get("grantflow:whoami", false);
        if (w?.value) setWhoami(w.value);
      } catch (e) { /* not set yet */ }
      setWhoamiLoaded(true);
    })();
  }, []);

  useEffect(() => {
    if (!whoamiLoaded || currentUserEmail) return;
    if (whoami) window.storage.set("grantflow:whoami", whoami, false).catch(() => {});
  }, [whoami, whoamiLoaded]);

  useEffect(() => {
    if (!loaded) return;
    const interval = setInterval(() => { refreshAll(); }, 60000);
    return () => clearInterval(interval);
  }, [loaded]);

  useEffect(() => {
    if (!loaded || isSyncingRef.current) return;
    saveKey("grantflow:grants", grants, "Grants");
  }, [grants, loaded]);

  useEffect(() => {
    if (!loaded || isSyncingRef.current) return;
    saveKey("grantflow:budgets", budgets, "Budgets");
  }, [budgets, loaded]);

  useEffect(() => {
    if (!loaded || isSyncingRef.current) return;
    saveKey("grantflow:reports", reports, "Grant reports");
  }, [reports, loaded]);

  useEffect(() => {
    if (!loaded || isSyncingRef.current) return;
    saveKey("grantflow:staff", staff, "Personnel");
  }, [staff, loaded]);

  useEffect(() => {
    if (!loaded || isSyncingRef.current) return;
    saveKey("grantflow:activity", activity, "Activity log");
  }, [activity, loaded]);

  useEffect(() => {
    if (!loaded || isSyncingRef.current) return;
    saveKey("grantflow:tasks", tasks, "Tasks");
  }, [tasks, loaded]);

  useEffect(() => {
    if (!loaded || isSyncingRef.current) return;
    saveKey("grantflow:invoices", invoices, "Invoices");
  }, [invoices, loaded]);

  useEffect(() => {
    if (!loaded || isSyncingRef.current) return;
    saveKey("grantflow:costcenters", costCenters, "Cost Centers");
  }, [costCenters, loaded]);

  useEffect(() => {
    if (!loaded || isSyncingRef.current) return;
    saveKey("grantflow:budgetgroups", budgetGroups, "Budget Groups");
  }, [budgetGroups, loaded]);

  useEffect(() => {
    if (!loaded || isSyncingRef.current) return;
    saveKey("grantflow:scenarios", scenarios, "Scenarios");
  }, [scenarios, loaded]);

  useEffect(() => {
    if (!loaded || isSyncingRef.current) return;
    saveKey("grantflow:trash", trash, "Trash");
  }, [trash, loaded]);

  const logActivity = (entity, action, label) => {
    setActivity((prev) => [{ id: uid(), timestamp: new Date().toISOString(), entity, action, label, by: whoami || "Unknown" }, ...prev].slice(0, 150));
  };

  const staffMigratedRef = useRef(false);
  useEffect(() => {
    if (!loaded || staffMigratedRef.current) return;
    staffMigratedRef.current = true;
    const needsMigration = staff.some((s) => !s.status);
    if (needsMigration) {
      setStaff((prev) => prev.map((s) => (s.status ? s : { ...s, status: "Active" })));
      logActivity?.("Staff", "Updated", "Set status to Active for all existing staff members (one-time update)");
    }
  }, [loaded]);

  const autoBackupRef = useRef(false);
  useEffect(() => {
    if (!loaded || autoBackupRef.current) return;
    autoBackupRef.current = true;
    (async () => {
      const today = new Date().toISOString().slice(0, 10);
      const metaKey = "grantflow:autobackup:meta";
      let meta = null;
      try {
        const m = await window.storage.get(metaKey, true);
        if (m?.value) meta = JSON.parse(m.value);
      } catch (e) { /* no meta yet */ }
      if (meta?.lastBackupDate === today) return; // already backed up today
      try {
        const payload = { exportedAt: new Date().toISOString(), grants, budgets, reports, staff, tasks, invoices, costCenters, budgetGroups, scenarios, trash, activity };
        await window.storage.set(`grantflow:autobackup:${today}`, JSON.stringify(payload), true);
        await window.storage.set(metaKey, JSON.stringify({ lastBackupDate: today }), true);
        const list = await window.storage.list("grantflow:autobackup:", true);
        const cutoff = Date.now() - AUTO_BACKUP_RETENTION_DAYS * 24 * 60 * 60 * 1000;
        for (const key of (list?.keys || [])) {
          if (key === metaKey) continue;
          const dateStr = key.replace("grantflow:autobackup:", "");
          const d = new Date(dateStr);
          if (!isNaN(d) && d.getTime() < cutoff) {
            window.storage.delete(key, true).catch(() => {});
          }
        }
      } catch (e) { /* fail silently — this should never interrupt normal use */ }
    })();
  }, [loaded]);

  const navNonceRef = useRef(0);
  const goTo = (nextTab, action, grantId, recordId) => {
    setTab(nextTab);
    navNonceRef.current += 1;
    const nonce = navNonceRef.current;
    setPendingNewGrant(nextTab === "grants" && action === "new");
    setPendingNewTask(nextTab === "tasks" && action === "new");
    setPendingExpandGrantId(nextTab === "grants" && recordId ? `${recordId}::${nonce}` : "");
    if (nextTab === "budgets" && grantId) setSelectedGrantId(grantId);
    setPendingOpenBudgetId(nextTab === "budgets" && recordId ? `${recordId}::${nonce}` : "");
    if (nextTab === "grant-reports" && grantId) setReportsGrantFilter(grantId);
    setPendingOpenReportId(nextTab === "grant-reports" && recordId ? `${recordId}::${nonce}` : "");
    setPendingOpenTaskId(nextTab === "tasks" && recordId ? `${recordId}::${nonce}` : "");
    setPendingOpenStaffId(nextTab === "personnel" && recordId ? `${recordId}::${nonce}` : "");
    setPendingOpenInvoiceId(nextTab === "invoicing" && recordId ? `${recordId}::${nonce}` : "");
  };

  return (
    <div className="min-h-screen flex" style={{ background: "#F6F7F3" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Oswald:wght@500;600;700&family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@400;500&display=swap');
        .font-display { font-family: 'Oswald', sans-serif; text-transform: uppercase; letter-spacing: 0.02em; }
        :root { --mono-font: 'JetBrains Mono', monospace; }
        body, input, select, textarea, button { font-family: 'Inter', system-ui, sans-serif; }
        @media print {
          .no-print { display: none !important; }
          main { max-width: 100% !important; padding: 0 !important; }
          body { background: #fff !important; }
        }
      `}</style>

      {mobileNavOpen && (
        <div className="fixed inset-0 z-40 md:hidden no-print" style={{ background: "rgba(28,38,36,0.45)" }} onClick={() => setMobileNavOpen(false)} />
      )}

      <aside
        className={`${mobileNavOpen ? "fixed inset-y-0 left-0 z-50 flex" : "hidden"} md:flex md:static md:z-auto w-64 md:w-56 shrink-0 flex-col no-print`}
        style={{ background: "#17313A" }}
      >
        <div className="px-5 py-5 flex items-center gap-2.5 justify-between">
          <div className="flex items-center gap-2.5">
            <svg width="34" height="34" viewBox="0 0 40 40" className="shrink-0">
              <path d="M6 4 H34 V24 L20 36 L6 24 Z" fill="none" stroke="#F0B21E" strokeWidth="2" strokeLinejoin="round" />
              <path d="M11 11 L12.3 13.8 L15.3 14.2 L13.1 16.3 L13.7 19.3 L11 17.8 L8.3 19.3 L8.9 16.3 L6.7 14.2 L9.7 13.8 Z" fill="#F0B21E" />
              <rect x="17" y="10" width="14" height="2.2" fill="#F0B21E" />
              <rect x="17" y="15" width="14" height="2.2" fill="#F0B21E" />
              <rect x="8" y="20" width="23" height="2.2" fill="#F0B21E" />
            </svg>
            <div className="leading-tight">
              <div className="font-display text-sm tracking-wide" style={{ color: "#FFFFFF" }}>NATION'S FINEST</div>
              <div className="text-xs tracking-wide" style={{ color: "#F0B21E" }}>GRANT PORTAL</div>
            </div>
          </div>
          <button onClick={() => setMobileNavOpen(false)} className="md:hidden p-1 rounded hover:bg-white/10">
            <X size={18} style={{ color: "#B9CBCF" }} />
          </button>
        </div>
        <nav className="px-3 space-y-1 overflow-y-auto">
          {[...NAV.filter((n) => !hiddenModules.includes(n.key)), ...(isAdmin ? [{ key: "user-access", label: "User Access", icon: Shield }] : [])].map(({ key, label, icon: Icon }) => (
            <button
              key={key}
              onClick={() => { setTab(key); setMobileNavOpen(false); }}
              className="w-full flex items-center gap-2.5 px-3 py-2 rounded-md text-sm text-left border-l-2"
              style={{
                background: tab === key ? "rgba(240,178,30,0.12)" : "transparent",
                borderColor: tab === key ? "#F0B21E" : "transparent",
                color: tab === key ? "#F0B21E" : "#B9CBCF",
                fontWeight: tab === key ? 600 : 400,
              }}
            >
              <Icon size={16} /> {label}
            </button>
          ))}
        </nav>
        <div className="mt-auto px-5 py-4 text-xs space-y-1.5" style={{ color: "#7C9298" }}>
          {saveError ? (
            <span className="inline-flex items-start gap-1" style={{ color: "#E08A82" }}>
              <AlertCircle size={12} className="mt-0.5 shrink-0" />
              {Object.values(saveErrors).join(", ")} failed to save after retrying — try "Refresh now" below, or check support.claude.com if this continues.
            </span>
          ) : (
            <span className="inline-flex items-center gap-1"><CheckCircle2 size={12} /> Synced (shared)</span>
          )}
          <button onClick={refreshAll} className="w-full flex items-center gap-1.5 text-xs hover:underline" style={{ color: "#B9CBCF" }}>
            <RefreshCw size={11} /> Refresh now{lastSyncedAt ? ` · ${Math.max(0, Math.round((Date.now() - lastSyncedAt) / 1000))}s ago` : ""}
          </button>
          {currentUserEmail ? (
            <div className="flex items-center justify-between gap-1.5 text-xs" style={{ color: "#B9CBCF" }}>
              <span className="inline-flex items-center gap-1.5 truncate"><Users size={11} className="shrink-0" /> {currentUserEmail}</span>
              <button onClick={onSignOut} className="shrink-0 hover:underline" style={{ color: "#F0B21E" }}>Sign out</button>
            </div>
          ) : (
            <button onClick={() => setEditingWhoami(true)} className="w-full flex items-center gap-1.5 text-xs hover:underline" style={{ color: "#B9CBCF" }}>
              <Users size={11} /> You: {whoami || "set your name"}
            </button>
          )}
        </div>
      </aside>

      <div className="flex-1 flex flex-col min-w-0">
        <div className="flex items-center gap-3 px-4 md:px-8 py-3 border-b no-print" style={{ borderColor: "#E1E5DE", background: "#FFFFFF" }}>
          <button onClick={() => setMobileNavOpen(true)} className="md:hidden p-1.5 rounded border shrink-0" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>
            <Menu size={18} />
          </button>
          <div className="flex-1 flex justify-end">
            <GlobalSearch grants={grants} budgets={budgets} reports={reports} staff={staff} tasks={tasks} invoices={invoices} costCenters={costCenters} goTo={goTo} />
          </div>
        </div>
        {announcement && !announcementDismissed && (
          <div className="no-print px-4 md:px-8 py-2.5 flex items-start gap-2" style={{ background: "#FFF7E6", borderBottom: "1px solid #F0B21E" }}>
            <AlertCircle size={15} style={{ color: "#8A6D1F", marginTop: 2 }} className="shrink-0" />
            <div className="flex-1 text-sm" style={{ color: "#5B4A0F" }}>
              {announcement.message}
              {announcement.setBy && <span style={{ color: "#8A8F87" }}> — {announcement.setBy}</span>}
            </div>
            <button onClick={() => setAnnouncementDismissed(true)} className="shrink-0 p-0.5 rounded hover:bg-black/5">
              <X size={15} style={{ color: "#8A6D1F" }} />
            </button>
          </div>
        )}
        <main className="flex-1 px-4 md:px-8 py-4 md:py-8" style={{ maxWidth: (tab === "grant-reports" || tab === "org-budget" || tab === "burn-rate") ? "100%" : "72rem" }}>
        {!loaded ? (
          <div className="text-sm" style={{ color: "#8A8F87" }}>Loading…</div>
        ) : tab === "dashboard" ? (
          <Dashboard grants={grants} budgets={budgets} reports={reports} tasks={tasks} staff={staff} invoices={invoices} goTo={goTo} />
        ) : tab === "grants" ? (
          <GrantsView
            key={pendingNewGrant ? "grants-new" : pendingExpandGrantId ? `grants-expand-${pendingExpandGrantId}` : "grants"}
            grants={grants} budgets={budgets} reports={reports} tasks={tasks} invoices={invoices} setGrants={setGrants} setBudgets={setBudgets}
            setReports={setReports} setTasks={setTasks} setStaff={setStaff} setInvoices={setInvoices} staff={staff}
            budgetGroups={budgetGroups} setBudgetGroups={setBudgetGroups}
            setTrash={setTrash} currentUserEmail={currentUserEmail || whoami} canEdit={canEdit}
            autoOpenNew={pendingNewGrant} initialExpandId={pendingExpandGrantId} goTo={goTo} logActivity={logActivity}
          />
        ) : tab === "budgets" ? (
          <BudgetsView
            key={pendingOpenBudgetId ? `budgets-open-${pendingOpenBudgetId}` : "budgets"}
            grants={grants} budgets={budgets} setBudgets={setBudgets}
            selectedGrantId={selectedGrantId} setSelectedGrantId={setSelectedGrantId}
            costCenters={costCenters} setCostCenters={setCostCenters}
            selectedCostCenterId={selectedCostCenterId} setSelectedCostCenterId={setSelectedCostCenterId}
            budgetGroups={budgetGroups} setBudgetGroups={setBudgetGroups}
            setTrash={setTrash} currentUserEmail={currentUserEmail || whoami} canEdit={canEdit}
            initialOpenBudgetId={pendingOpenBudgetId} logActivity={logActivity}
          />
        ) : tab === "invoicing" ? (
          <InvoicingView
            key={pendingOpenInvoiceId ? `invoices-open-${pendingOpenInvoiceId}` : "invoicing"}
            grants={grants} invoices={invoices} setInvoices={setInvoices}
            setTrash={setTrash} currentUserEmail={currentUserEmail || whoami} canEdit={canEdit}
            initialOpenInvoiceId={pendingOpenInvoiceId} logActivity={logActivity}
          />
        ) : tab === "tasks" ? (
          <TasksView
            key={pendingNewTask ? "tasks-new" : pendingOpenTaskId ? `tasks-open-${pendingOpenTaskId}` : "tasks"}
            grants={grants} tasks={tasks} setTasks={setTasks}
            setTrash={setTrash} currentUserEmail={currentUserEmail || whoami} canEdit={canEdit}
            autoOpenNew={pendingNewTask} initialOpenTaskId={pendingOpenTaskId} logActivity={logActivity}
          />
        ) : tab === "grant-reports" ? (
          <ReportsView
            key={pendingOpenReportId ? `reports-open-${pendingOpenReportId}` : "grant-reports"}
            grants={grants} reports={reports} setReports={setReports} setTasks={setTasks}
            grantFilter={reportsGrantFilter} setGrantFilter={setReportsGrantFilter}
            setTrash={setTrash} currentUserEmail={currentUserEmail || whoami} canEdit={canEdit}
            initialOpenReportId={pendingOpenReportId} logActivity={logActivity}
          />
        ) : tab === "org-budget" ? (
          <OrgBudgetView grants={grants} budgets={budgets} costCenters={costCenters} budgetGroups={budgetGroups} />
        ) : tab === "scenarios" ? (
          <ScenariosView
            scenarios={scenarios} setScenarios={setScenarios}
            grants={grants} budgets={budgets} costCenters={costCenters} budgetGroups={budgetGroups}
            whoami={currentUserEmail || whoami} setTrash={setTrash} canEdit={canEdit} logActivity={logActivity}
          />
        ) : tab === "burn-rate" ? (
          <BurnRateView grants={grants} budgets={budgets} />
        ) : tab === "personnel" ? (
          <PersonnelView
            key={pendingOpenStaffId ? `staff-open-${pendingOpenStaffId}` : "personnel"}
            grants={grants} staff={staff} setStaff={setStaff} costCenters={costCenters}
            setTrash={setTrash} currentUserEmail={currentUserEmail || whoami} canEdit={canEdit}
            initialOpenStaffId={pendingOpenStaffId} logActivity={logActivity}
          />
        ) : tab === "activity-log" ? (
          <ActivityLogView activity={activity} />
        ) : tab === "trash" ? (
          <TrashView
            trash={trash} setTrash={setTrash}
            setGrants={setGrants} setBudgets={setBudgets} setReports={setReports} setTasks={setTasks}
            setInvoices={setInvoices} setStaff={setStaff} setCostCenters={setCostCenters} setScenarios={setScenarios}
            isAdmin={isAdmin} canEdit={canEdit} logActivity={logActivity}
          />
        ) : tab === "data" ? (
          <DataView
            grants={grants} budgets={budgets} reports={reports} staff={staff} tasks={tasks} invoices={invoices} costCenters={costCenters} budgetGroups={budgetGroups} scenarios={scenarios} trash={trash} activity={activity}
            setGrants={setGrants} setBudgets={setBudgets} setReports={setReports} setStaff={setStaff} setTasks={setTasks} setInvoices={setInvoices} setCostCenters={setCostCenters} setBudgetGroups={setBudgetGroups} setScenarios={setScenarios} setTrash={setTrash} setActivity={setActivity}
            canEdit={canEdit}
            logActivity={logActivity}
          />
        ) : tab === "user-access" && isAdmin ? (
          <AdminPanel currentUserEmail={currentUserEmail || whoami} />
        ) : (
          <ReportingView grants={grants} budgets={budgets} />
        )}
        </main>
      </div>

      {!currentUserEmail && editingWhoami && (
        <WhoamiModal current={whoami} onSave={(n) => { setWhoami(n); setEditingWhoami(false); }} onSkip={() => setEditingWhoami(false)} />
      )}
      {!currentUserEmail && !editingWhoami && whoamiLoaded && !whoami && !skippedWhoami && (
        <WhoamiModal current={whoami} onSave={(n) => setWhoami(n)} onSkip={() => setSkippedWhoami(true)} />
      )}
      <div className="no-print fixed bottom-2 right-3 text-xs z-40 pointer-events-none" style={{ color: "#8A8F87" }}>
        v{APP_VERSION}
      </div>
    </div>
  );
}

