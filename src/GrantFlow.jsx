import { useState, useEffect, useMemo, useRef, Fragment, Component } from "react";
import { doc, getDoc } from "firebase/firestore";
import { db } from "./firebaseConfig";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";
import {
  LayoutDashboard, FileText, Wallet, BarChart3, Plus, X, Pencil, Trash2,
  ExternalLink, Download, Search, ArrowRight, AlertCircle, CheckCircle2,
  ClipboardList, Circle, CheckCircle, Users, PieChart, TrendingUp, History, CheckSquare, Upload, Printer, RefreshCw, Receipt, Menu, Shield, FlaskConical, Undo2,
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
// Parses a budget's "actuals complete through" marker (stored as "YYYY-MM")
// into a comparable { year, monthIndex } shape. Returns null if unset.
function parseActualsThrough(str) {
  if (!str) return null;
  const [y, m] = str.split("-").map(Number);
  if (!y || !m) return null;
  return { year: y, monthIndex: m - 1 };
}
// True if a month column falls on/before the marked cutoff (i.e. it's real,
// entered data) — false if it's past the cutoff and should be projected.
function colIsWithinCutoff(col, cutoff) {
  if (!cutoff) return true;
  if (col.year !== cutoff.year) return col.year < cutoff.year;
  return col.monthIndex <= cutoff.monthIndex;
}

// Copies actuals from a Template budget's lines onto every linked
// Operational budget, matching by category+subcategory and by actual
// calendar month — not array position — since periods can start on
// different dates. A calendar-year Template can span more than one
// Operational budget (e.g. two Oct–Sep fiscal years); each linked budget
// only receives whichever months actually fall within its own period,
// determined independently per target. If a target doesn't already have a
// matching line, one is created (Plan left at $0) so no data is silently
// dropped. One-way only: Operational edits never flow back to the Template.
// Replicates the Org Budget page's "Whole Organization" scope (Template
// budgets only, Active/Awarded, Deferred Revenue group excluded) so any
// other view — like the Dashboard's FY summary — always matches what Org
// Budget itself shows, rather than maintaining a second, drifting version
// of the same math. Plan revenue is a straight sum; Actual expense reuses
// the same owner-chained cross-year bridging as the Org Budget page, so a
// marked "actuals complete through" cutoff projects forward here too.
function computeOrgFYTotals(budgets, grants, costCenters, budgetGroups, calYear) {
  const deferredRevenueGroupId = (budgetGroups || []).find((bg) => bg.name?.trim().toLowerCase() === "deferred revenue")?.id || null;
  const excludedGrantIds = deferredRevenueGroupId ? new Set(grants.filter((g) => g.budgetGroupId === deferredRevenueGroupId).map((g) => g.id)) : null;
  const excludedCcIds = deferredRevenueGroupId ? new Set((costCenters || []).filter((c) => c.budgetGroupId === deferredRevenueGroupId).map((c) => c.id)) : null;

  const scopedBudgets = budgets
    .filter((b) => !(b.grantId && excludedGrantIds?.has(b.grantId)) && !(b.costCenterId && excludedCcIds?.has(b.costCenterId)))
    .filter((b) => isActiveBudget(b.status) && b.budgetType === "Template");

  let planRevenue = 0;
  scopedBudgets.forEach((b) => {
    const cols = monthColumnsForBudget(b.periodStart, b.periodEnd);
    b.lines.forEach((l) => {
      if (l.type !== "revenue") return;
      const vals = l.amounts || [];
      cols.forEach((col, i) => {
        if (col && col.year === calYear) planRevenue += Number(vals[i]) || 0;
      });
    });
  });

  const ownerKey = (b) => b.grantId || b.costCenterId || `__standalone_${b.id}`;
  const byOwner = {};
  scopedBudgets.forEach((b) => { (byOwner[ownerKey(b)] = byOwner[ownerKey(b)] || []).push(b); });

  let actualExpense = 0;
  Object.values(byOwner).forEach((ownerBudgets) => {
    ownerBudgets.sort((a, b) => new Date(a.periodStart) - new Date(b.periodStart));
    const lineMeta = {};
    ownerBudgets.forEach((b) => b.lines.forEach((l) => {
      const key = `${l.category}|||${l.subcategory || ""}`;
      if (!lineMeta[key]) lineMeta[key] = { type: l.type };
    }));

    const budgetLineAvg = {};
    const budgetLineActuals = {};
    ownerBudgets.forEach((b) => {
      const cols = monthColumnsForBudget(b.periodStart, b.periodEnd);
      Object.keys(lineMeta).forEach((key) => {
        const matching = b.lines.filter((l) => `${l.category}|||${l.subcategory || ""}` === key);
        if (matching.length === 0) return;
        const combined = Array(cols.length).fill(0);
        matching.forEach((l) => (l.actuals || []).forEach((v, i) => { if (i < combined.length) combined[i] += Number(v) || 0; }));
        budgetLineActuals[`${b.id}|||${key}`] = combined;
        const cutoff = parseActualsThrough(b.actualsThrough);
        if (!cutoff) return;
        const vals = cols.map((col, i) => (colIsWithinCutoff(col, cutoff) ? combined[i] : null)).filter((v) => v !== null);
        budgetLineAvg[`${b.id}|||${key}`] = vals.length ? vals.reduce((a, x) => a + x, 0) / vals.length : 0;
      });
    });

    Object.keys(lineMeta).forEach((key) => {
      if (lineMeta[key].type !== "expense") return;
      const timeline = [];
      ownerBudgets.forEach((b) => {
        const cols = monthColumnsForBudget(b.periodStart, b.periodEnd);
        const combined = budgetLineActuals[`${b.id}|||${key}`];
        const cutoff = parseActualsThrough(b.actualsThrough);
        cols.forEach((col, i) => {
          timeline.push({
            year: col.year,
            monthIndex: col.monthIndex,
            isRealEntry: !!(combined && cutoff && colIsWithinCutoff(col, cutoff)),
            rawValue: combined ? (Number(combined[i]) || 0) : 0,
            budgetId: b.id,
            atOrAfterCutoff: cutoff ? !colIsWithinCutoff(col, cutoff) : false,
          });
        });
      });
      timeline.sort((a, b) => (a.year - b.year) || (a.monthIndex - b.monthIndex));
      let basis = null;
      timeline.forEach((t) => {
        const avgKey = `${t.budgetId}|||${key}`;
        if (t.atOrAfterCutoff && budgetLineAvg[avgKey] !== undefined) basis = budgetLineAvg[avgKey];
        const value = t.isRealEntry ? t.rawValue : (basis !== null ? basis : t.rawValue);
        if (t.year === calYear) actualExpense += value;
      });
    });
  });

  return { planRevenue, actualExpense };
}

function syncActualsToLinkedBudgets(templateBudget, allBudgets) {
  const linkedIds = templateBudget.linkedBudgetIds && templateBudget.linkedBudgetIds.length
    ? templateBudget.linkedBudgetIds
    : (templateBudget.linkedBudgetId ? [templateBudget.linkedBudgetId] : []); // legacy single-link fallback
  if (templateBudget.budgetType !== "Template" || linkedIds.length === 0 || !templateBudget.periodStart || !templateBudget.periodEnd) return [];

  const templateCols = monthColumnsForBudget(templateBudget.periodStart, templateBudget.periodEnd);
  const results = [];

  linkedIds.forEach((targetId) => {
    const target = allBudgets.find((b) => b.id === targetId);
    if (!target || !target.periodStart || !target.periodEnd) return;

    const targetCols = monthColumnsForBudget(target.periodStart, target.periodEnd);
    const targetColIndexByYM = {};
    targetCols.forEach((c, j) => { targetColIndexByYM[`${c.year}-${c.monthIndex}`] = j; });

    let changed = false;
    const newLines = target.lines.map((l) => ({ ...l, actuals: [...(l.actuals || Array(targetCols.length).fill(0))] }));

    templateBudget.lines.forEach((tl) => {
      const key = `${tl.category}|||${tl.subcategory || ""}`;
      let targetLine = newLines.find((l) => `${l.category}|||${l.subcategory || ""}` === key);
      // Only months from this template line that actually fall within this
      // specific target's period will match below — a line with no
      // in-range months for this target simply won't create one.
      const hasInRangeMonth = (tl.actuals || []).some((v, i) => {
        const col = templateCols[i];
        return col && targetColIndexByYM[`${col.year}-${col.monthIndex}`] !== undefined;
      });
      if (!targetLine && !hasInRangeMonth) return;
      if (!targetLine) {
        targetLine = { ...newLine(), category: tl.category, subcategory: tl.subcategory, type: tl.type, amounts: Array(targetCols.length).fill(0), actuals: Array(targetCols.length).fill(0) };
        newLines.push(targetLine);
        changed = true;
      }
      (tl.actuals || []).forEach((v, i) => {
        const col = templateCols[i];
        if (!col) return;
        const j = targetColIndexByYM[`${col.year}-${col.monthIndex}`];
        if (j === undefined) return;
        const numVal = Number(v) || 0;
        if ((Number(targetLine.actuals[j]) || 0) !== numVal) {
          targetLine.actuals[j] = numVal;
          changed = true;
        }
      });
    });

    if (changed) results.push({ ...target, lines: newLines });
  });

  return results;
}

function monthColumnsForBudget(periodStart, periodEnd) {
  let startYear, startMonth; // startMonth is 0-indexed
  if (periodStart) {
    const d = new Date(periodStart + "T00:00:00");
    if (!isNaN(d)) { startYear = d.getFullYear(); startMonth = d.getMonth(); }
  }
  if (startYear === undefined) { startYear = new Date().getFullYear(); startMonth = 0; }

  let count = 12; // default when there's no period end to measure against yet
  if (periodEnd) {
    const endD = new Date(periodEnd + "T00:00:00");
    if (!isNaN(endD)) {
      const diff = (endD.getFullYear() - startYear) * 12 + (endD.getMonth() - startMonth) + 1;
      if (diff > 0) count = diff;
    }
  }

  return Array.from({ length: count }, (_, i) => {
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
    "6999 - Asses Sale Gain/Loss", "8000 - Unallocated Expenses",
  ] },
  { name: "In-Kind Expenses", type: "expense", subs: ["6066 - InKind Rent Expense", "6067 - InKind Event Expense"] },
  { name: "TFA", type: "expense", subs: ["7100 - Category 1 Temporary Financial Assistance", "7200 - Category 2 Temporary Financial Assistance"] },
  // Balance sheet accounts — tracked with the same Plan/Actual monthly
  // structure as any other line, but deliberately excluded from
  // revenue/expense/net everywhere (see budgetTotals, budgetActualTotals,
  // and OrgBudgetView) since they're not income statement activity.
  { name: "Deferred Revenue", type: "balance", subs: ["1290 - Deferred Revenue - SSVF", "2600 - Deferred Revenue"] },
];
const CUSTOM_CATEGORY = "__custom__";

// Maps a bare account code (e.g. "4100") to its canonical category,
// subcategory string, and type — built once from CATEGORIES so an imported
// Excel template lands under the exact same names every other budget in the
// app already uses, regardless of minor wording/punctuation differences in
// the source file (e.g. "Grants & Contracts" vs "Grants and Contracts").
const CODE_TO_CATEGORY = {};
CATEGORIES.forEach((cat) => {
  cat.subs.forEach((sub) => {
    const m = sub.match(/^(\d{3,4}[A-Za-z]*)\s*-/);
    if (m) CODE_TO_CATEGORY[m[1]] = { category: cat.name, subcategory: sub, type: cat.type };
  });
});

const EXCEL_ERROR_STRINGS = ["#REF!", "#DIV/0!", "#VALUE!", "#N/A", "#NAME?", "#NULL!", "#NUM!"];

// Parses one of Nation's Finest's "Rolling 12 Months P&L" grant budget
// exports (a fixed, known layout per grant, confirmed consistent across
// templates) into the shape needed to create or update a Template budget.
// Rows are matched purely by leading account code against CODE_TO_CATEGORY
// — this sidesteps needing to parse the file's own header/indentation
// hierarchy at all, and guarantees whatever lands in GrantFlow uses exactly
// the category names already used everywhere else.
function parseExcelTemplateWorkbook(aoa, sheet) {
  const warnings = [];
  const projectLabel = String(aoa[6]?.[1] || "");
  const projectCodeMatch = projectLabel.match(/^(\d+)/);
  const projectCode = projectCodeMatch ? projectCodeMatch[1] : null;

  const actualsNote = String(aoa[4]?.[5] || "");
  const actualsNoteMatch = actualsNote.match(/Updated up to (\d{2})\/(\d{2})\/(\d{4})/);
  const actualsThroughDate = actualsNoteMatch ? { year: Number(actualsNoteMatch[3]), monthIndex: Number(actualsNoteMatch[1]) - 1 } : null;

  const monthCols = [];
  for (let c = 1; c <= 12; c++) {
    const raw = aoa[9]?.[c];
    if (raw === null || raw === undefined) { monthCols.push(null); continue; }
    const d = new Date(raw);
    monthCols.push(isNaN(d.getTime()) ? null : { year: d.getUTCFullYear(), monthIndex: d.getUTCMonth() });
  }
  const validMonths = monthCols.filter(Boolean);
  if (validMonths.length === 0) {
    return { error: "Couldn't read the month headers in row 10 — this doesn't look like the expected template layout." };
  }
  const fy = aoa[9]?.[13];
  const firstMonth = validMonths[0];
  const lastMonth = validMonths[validMonths.length - 1];
  const periodStart = `${firstMonth.year}-${String(firstMonth.monthIndex + 1).padStart(2, "0")}-01`;
  const lastDay = new Date(lastMonth.year, lastMonth.monthIndex + 1, 0).getDate();
  const periodEnd = `${lastMonth.year}-${String(lastMonth.monthIndex + 1).padStart(2, "0")}-${String(lastDay).padStart(2, "0")}`;

  const isMonthActual = (col) => {
    if (!col) return false;
    if (actualsThroughDate) {
      return col.year < actualsThroughDate.year || (col.year === actualsThroughDate.year && col.monthIndex <= actualsThroughDate.monthIndex);
    }
    return false;
  };

  const linesByKey = {};
  const unrecognizedCodes = new Set();
  let inKindBlock = false;

  for (let i = 11; i < aoa.length; i++) {
    const row = aoa[i];
    if (!row) continue;
    const rawLabel = row[0];
    if (rawLabel === null || rawLabel === undefined) continue;
    const label = String(rawLabel).trim();
    if (!label) continue;

    if (/^in\s*kind\s*activity$/i.test(label)) { inKindBlock = true; continue; }
    if (/^total\s+in\s*kind\s*activity/i.test(label)) { inKindBlock = false; continue; }
    if (inKindBlock) continue;

    const codeMatch = label.match(/^(\d{3,4}[A-Za-z]*)\s*-/);
    if (!codeMatch) continue; // header, "Total ..." rollup, ratio row, etc. — not a real line

    const target = CODE_TO_CATEGORY[codeMatch[1]];
    if (!target) { unrecognizedCodes.add(label); continue; }

    const key = `${target.category}|||${target.subcategory}`;
    if (!linesByKey[key]) linesByKey[key] = { ...target, amounts: Array(12).fill(0), actuals: Array(12).fill(0) };

    for (let m = 0; m < 12; m++) {
      const col = monthCols[m];
      if (!col) continue;
      const raw = row[m + 1];
      const rawCell = sheet ? sheet[XLSX.utils.encode_cell({ r: i, c: m + 1 })] : null;
      let val = 0;
      if (rawCell && rawCell.t === "e") {
        warnings.push(`"${target.category} / ${target.subcategory}", ${MONTHS[col.monthIndex]} ${col.year}: source cell contains a formula error (${rawCell.w || "#ERROR"}) — imported as $0.`);
      } else if (typeof raw === "number") {
        val = raw;
      } else if (typeof raw === "string" && EXCEL_ERROR_STRINGS.includes(raw.trim())) {
        warnings.push(`"${target.category} / ${target.subcategory}", ${MONTHS[col.monthIndex]} ${col.year}: source cell contains ${raw.trim()} — imported as $0.`);
      } else if (raw !== null && raw !== undefined && raw !== "") {
        const n = Number(raw);
        if (!isNaN(n)) val = n;
      }
      if (isMonthActual(col)) linesByKey[key].actuals[m] += val;
      else linesByKey[key].amounts[m] += val;
    }
  }

  return {
    projectCode, fy, periodStart, periodEnd, actualsThroughDate, monthCols,
    lines: Object.values(linesByKey),
    warnings,
    unrecognizedCodes: [...unrecognizedCodes],
  };
}

const STAGES = ["Prospecting", "Writing", "Applied", "Awarded", "Rejected", "Active", "Closing", "Closed"];
const SITE_OPTIONS = [
  "Carson City", "Chico", "Flagstaff", "Mather", "Menlo Park", "Monterey", "Prescott",
  "Redding", "Reno", "Sacramento", "Santa Cruz", "Vacaville", "Bullhead City", "Eureka",
  "Santa Rosa", "All CA", "All AZ", "All NV", "Residential", "MSU", "Corp",
];
const RISKS = ["Low", "Medium", "High"];
const CADENCES = ["Weekly", "Monthly", "Quarterly", "Semi-annual", "Annually", "End of grant"];
const BUDGET_STATUSES = ["Draft", "Pending Approval", "Active", "Awarded", "Rejected", "Closed"];
// "Operational" is what actually goes to the grantor. "Template" is Nation's
// Finest's own internal version of the same budget, used to roll up into the
// org-wide Org Budget view — the two can differ (internal cost allocation,
// admin categorization, etc.) without one distorting the other.
const BUDGET_TYPES = ["Operational", "Template"];
// "Awarded" is functionally identical to "Active" everywhere budgets are scoped,
// totaled, or rolled up (org budget, scenarios, burn rate, dashboard counts, etc).
// It exists as a separate status purely so it can be labeled/badged differently.
const isActiveBudget = (status) => status === "Active" || status === "Awarded";
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
const DEFAULT_BUCKETS = ["Upcoming", "Up next", "Overdue", "In progress", "Complete", "Submitted"];
const TASK_STATUSES = ["Not started", "In progress", "Done"];
const TASK_CATEGORIES = ["Application/Submission", "Site Visit", "Renewal Prep", "Document Collection", "Board Approval", "Compliance", "Personnel Reallocation", "Report Submission", "Other"];

const APP_VERSION = "1.2.0";
const uid = () => Math.random().toString(36).slice(2, 10);
const stripNonce = (v) => (v ? v.split("::")[0] : "");
const fmt = (n) => {
  let v = Number(n) || 0;
  // Sums of many decimal values can leave tiny float residue (e.g.
  // -0.0000000000003) that's technically negative but should read as a
  // clean $0, not "-$0".
  if (Math.round(v * 100) === 0) v = 0;
  return v.toLocaleString("en-US", { style: "currency", currency: "USD", maximumFractionDigits: 0 });
};
// Sums of many decimal (cents) values can leave tiny floating-point residue
// (e.g. -0.0000000000003) that's technically negative but displays as $0
// after fmt()'s rounding. Round to the nearest cent before deciding color so
// a value that visibly reads as $0 never gets colored as if it were negative.
const isNetNegative = (n) => Math.round((Number(n) || 0) * 100) < 0;
const fmtDate = (d) => (d ? new Date(d + "T00:00:00").toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" }) : "—");

function newLine(monthCount = 12) {
  const c = CATEGORIES[0];
  return {
    id: uid(), category: c.name, type: c.type, categoryCustom: false,
    subcategory: "", subcategoryCustom: false, description: "",
    amounts: Array(monthCount).fill(0), actuals: Array(monthCount).fill(0),
  };
}

function lineTotal(line) {
  return line.amounts.reduce((a, b) => a + (Number(b) || 0), 0);
}

function lineActualTotal(line) {
  return (line.actuals || []).reduce((a, b) => a + (Number(b) || 0), 0);
}

function varianceInfo(plan, actual) {
  const p = Number(plan) || 0;
  const a = Number(actual) || 0;
  if (p === 0 && a === 0) return { label: "—", color: "#8A8F87" };
  if (p === 0) return { label: "New", color: "#B5443A" };
  const pct = ((a - p) / Math.abs(p)) * 100;
  const sign = pct >= 0 ? "+" : "";
  return { label: `${sign}${Math.round(pct)}%`, color: pct >= 0 ? "#2F6F53" : "#B5443A" };
}

function resizeMonthlyArray(arr, newLength) {
  const a = arr || [];
  if (a.length === newLength) return a;
  if (a.length > newLength) return a.slice(0, newLength);
  return [...a, ...Array(newLength - a.length).fill(0)];
}

function distributeEvenly(total, count = 12) {
  const cents = Math.round((Number(total) || 0) * 100);
  const base = Math.floor(cents / count);
  const remainder = cents - base * count;
  return Array.from({ length: count }, (_, i) => (base + (i < remainder ? 1 : 0)) / 100);
}

function budgetTotals(budget) {
  let revenue = 0, expense = 0;
  const monthly = Array(12).fill(0);
  for (const line of budget.lines) {
    const t = lineTotal(line);
    if (line.type === "revenue") { revenue += t; line.amounts.forEach((a, i) => { monthly[i] += Number(a) || 0; }); }
    else if (line.type === "expense") { expense += t; line.amounts.forEach((a, i) => { monthly[i] -= Number(a) || 0; }); }
    // Balance-sheet-type lines (e.g. Deferred Revenue) are tracked but
    // deliberately excluded from revenue/expense/net — they're not income
    // statement activity.
  }
  return { revenue, expense, net: revenue - expense, monthly };
}

function budgetActualTotals(budget) {
  let revenue = 0, expense = 0;
  const monthly = Array(12).fill(0);
  for (const line of budget.lines) {
    const t = lineActualTotal(line);
    if (line.type === "revenue") { revenue += t; (line.actuals || []).forEach((a, i) => { monthly[i] += Number(a) || 0; }); }
    else if (line.type === "expense") { expense += t; (line.actuals || []).forEach((a, i) => { monthly[i] -= Number(a) || 0; }); }
  }
  return { revenue, expense, net: revenue - expense, monthly };
}

// A drop-in replacement for useState that keeps a short undo history.
// Rapid changes (typing, dragging a slider) get grouped into one undo step
// after a brief pause, so "Undo" steps back through meaningful edits rather
// than one keystroke at a time. This is purely in-memory and local to the
// open card — it has no effect on saved data until the user clicks Save.
function useUndoableState(initialValue) {
  const [history, setHistory] = useState([initialValue]);
  const [index, setIndex] = useState(0);
  const groupingRef = useRef(false);
  const timerRef = useRef(null);

  const current = history[index];

  const setValue = (updater) => {
    const base = history[index];
    const next = typeof updater === "function" ? updater(base) : updater;

    if (groupingRef.current) {
      // Still within the grouping window — replace this step instead of
      // creating a new one, so continuous typing counts as one undo step.
      setHistory((h) => {
        const copy = h.slice(0, index + 1);
        copy[index] = next;
        return copy;
      });
    } else {
      groupingRef.current = true;
      setHistory((h) => [...h.slice(0, index + 1), next]);
      setIndex((i) => i + 1);
    }

    clearTimeout(timerRef.current);
    timerRef.current = setTimeout(() => { groupingRef.current = false; }, 800);
  };

  const undo = () => setIndex((i) => Math.max(0, i - 1));
  const canUndo = index > 0;

  return [current, setValue, undo, canUndo];
}

function grantBudgetTotals(grantId, budgets) {
  const mine = budgets.filter((b) => b.grantId === grantId);
  return mine.reduce((acc, b) => {
    const t = budgetTotals(b);
    acc.revenue += t.revenue; acc.expense += t.expense;
    return acc;
  }, { revenue: 0, expense: 0 });
}

// Same idea as grantBudgetTotals, but only counts budgets whose period is
// currently active (today falls within Period Start–Period End) — avoids
// stacking multiple budget years together for grants with several budgets.
function grantCurrentPeriodBudgetTotals(grantId, budgets) {
  const today = new Date();
  const mine = budgets.filter((b) => {
    if (b.grantId !== grantId) return false;
    if (!b.periodStart || !b.periodEnd) return false;
    const start = new Date(b.periodStart);
    const end = new Date(b.periodEnd);
    return today >= start && today <= end;
  });
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
  const today = new Date();
  const mine = budgets.filter((b) => {
    if (b.grantId !== grant.id) return false;
    if (b.budgetType !== "Operational") return false; // pace against what's actually contracted with the grantor, not the internal Template version
    if (!b.periodStart || !b.periodEnd) return false;
    return today >= new Date(b.periodStart) && today <= new Date(b.periodEnd);
  });
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

// Column classification for Paylocity's "Labor Distribution Percentages"
// report (grouped by Worked Program). Verified against two real exports
// (single pay period AND full year-to-date) with zero mismatches across all
// 398 employee-records: WAGE = every "E - ... Amt" column EXCEPT the
// specific employer benefit-cost columns below. An exclusion list is used
// rather than an inclusion whitelist so a new earning code Paylocity adds
// later (bonus, holiday, jury duty, etc.) is correctly treated as wages by
// default instead of silently dropped.
const PAYLOCITY_BENEFIT_COLS = ["E - 401ER Amt", "E - 401PS Amt", "E - ERCBU Amt", "E - ERCIG Amt", "E - ERDEN Amt", "E - ERKBU Amt", "E - ERKSR Amt", "E - ERLIF Amt", "E - ERWSH Amt", "E - RPHRS Amt", "E - REGM Amt"];
const PAYLOCITY_EMPLOYER_TAX_COLS = ["R - MED-R Amt", "R - SS-R Amt", "R - AZSUI Amt", "R - CAETT Amt", "R - CASUI Amt", "R - FLSUI Amt", "R - LASUI Amt", "R - NVCLA Amt", "R - NVSUI Amt", "R - OHSUI Amt", "R - ORSUI Amt", "R - ORWC Amt", "R - SCAST Amt", "R - SCSUI Amt", "R - TNSUI Amt", "R - TXAST Amt", "R - TXETT Amt", "R - TXSUI Amt", "R - VASUI Amt", "R - WACLA Amt", "R - WASUI Amt", "R - OR-LAN1 Amt"];
const PAYLOCITY_FFCRA_CREDIT_COLS = ["R - FFCRAMC Amt", "R - FFCRAMPC Amt", "R - FFCRASC Amt", "R - FFCRAWC Amt"];

const sumCols = (row, cols) => cols.reduce((a, c) => a + (Number(row[c]) || 0), 0);
// Sums every "E - ... Amt" column present on a row except the known
// benefit-cost columns — see note above on why this is exclusion-based.
const sumWageCols = (row) => Object.keys(row).reduce((a, k) => (
  k.startsWith("E - ") && k.endsWith(" Amt") && !PAYLOCITY_BENEFIT_COLS.includes(k) ? a + (Number(row[k]) || 0) : a
), 0);
// Paylocity exports names as "LAST, FIRST M." — normalize and compare just
// last+first (ignoring middle initial/punctuation/case) to suggest a likely
// match against an existing staff record's "Last, First" name.
const paylocityNameKey = (n) => {
  const norm = (n || "").toLowerCase().replace(/[.,]/g, "").replace(/\s+/g, " ").trim();
  return norm.split(" ").slice(0, 2).join(" ");
};
const daysBetweenInclusive = (start, end) => Math.round((new Date(end + "T00:00:00") - new Date(start + "T00:00:00")) / 86400000) + 1;
// Splits a lump sum earned over [periodStart, periodEnd] across the calendar
// months it spans, proportional to how many days of the period fall in each
// month — used for writing real dollars into the right Actuals columns
// rather than annualizing (annualizing is only for the Personnel card).
function splitAmountAcrossMonths(amount, periodStart, periodEnd) {
  const totalDays = daysBetweenInclusive(periodStart, periodEnd);
  if (totalDays <= 0) return [];
  const start = new Date(periodStart + "T00:00:00");
  const end = new Date(periodEnd + "T00:00:00");
  const byMonth = {};
  for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
    const key = `${d.getFullYear()}-${d.getMonth()}`;
    byMonth[key] = (byMonth[key] || 0) + 1;
  }
  return Object.entries(byMonth).map(([key, days]) => {
    const [year, monthIndex] = key.split("-").map(Number);
    return { year, monthIndex, amount: amount * (days / totalDays) };
  });
}

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

// Base salary/wages plus bonus, benefits, and employer payroll tax — the
// "fully loaded" cost of an employee, used for generating budget lines from
// Personnel. Deliberately separate from staffAnnualCost() above: that
// function is base pay only and is what every existing Personnel-cost
// figure has already been validated against, so it's left untouched.
function staffFullyLoadedCost(staff) {
  const fte = Number(staff.fte) || 0;
  const base = staffAnnualCost(staff);
  const bonus = (Number(staff.bonus) || 0) * fte;
  const benefits = (Number(staff.benefits) || 0) * fte;
  const taxableWages = base + bonus;
  const payrollTax = taxableWages * ((Number(staff.payrollTaxRate) || 0) / 100);
  return {
    base, bonus, benefits, payrollTax,
    wagesTotal: base + bonus, // maps to the "5000 - Salary and Wages" budget subcategory
    taxAndBenefitsTotal: benefits + payrollTax, // maps to "5900 - Payroll taxes and benefits"
    total: base + bonus + benefits + payrollTax,
  };
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
    const scopedBudgets = (scope === "all" ? budgets : budgets.filter((b) => (b.grantId && scopedGrantIds.has(b.grantId)) || (b.costCenterId && scopedCcIds.has(b.costCenterId)))).filter((b) => isActiveBudget(b.status));
    const map = {};
    scopedBudgets.forEach((b) => {
      const cols = monthColumnsForBudget(b.periodStart, b.periodEnd);
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

function daysToPay(inv) {
  if (!inv.submittedDate || !inv.paidDate) return null;
  const diff = new Date(inv.paidDate + "T00:00:00").getTime() - new Date(inv.submittedDate + "T00:00:00").getTime();
  return Math.max(0, Math.round(diff / 86400000));
}

function inPeriod(dateStr, period) {
  if (!dateStr) return false;
  if (period === "all") return true;
  const d = new Date(dateStr + "T00:00:00");
  const now = new Date();
  if (period === "month") return d.getFullYear() === now.getFullYear() && d.getMonth() === now.getMonth();
  if (period === "year") return d.getFullYear() === now.getFullYear();
  return true;
}

function isInvoiceOverdue(inv) {
  return inv.status === "Submitted" && inv.dueDate && !inv.paidDate && new Date(inv.dueDate) < new Date(new Date().toDateString());
}

const priorityColor = (label) => (REPORT_PRIORITIES.find((p) => p.label === label) || REPORT_PRIORITIES[2]).color;

function advanceDateByRepeat(dateStr, repeat) {
  if (!dateStr) return "";
  const d = new Date(dateStr + "T00:00:00");
  if (repeat === "Weekly") d.setDate(d.getDate() + 7);
  else if (repeat === "Monthly") d.setMonth(d.getMonth() + 1);
  else if (repeat === "Quarterly") d.setMonth(d.getMonth() + 3);
  else if (repeat === "Annually") d.setFullYear(d.getFullYear() + 1);
  else return "";
  return d.toISOString().slice(0, 10);
}

function spawnNextReportOccurrence(report, grant) {
  if (!report.repeat || report.repeat === "None") return null;
  const nextDue = advanceDateByRepeat(report.dueDate, report.repeat);
  if (!nextDue) return null;
  if (grant) {
    if (grant.stage === "Closed" || grant.stage === "Rejected") return null;
    if (grant.end && new Date(nextDue + "T00:00:00") > new Date(grant.end + "T00:00:00")) return null;
  }
  return {
    ...report,
    id: uid(),
    dueDate: nextDue,
    startDate: report.startDate ? advanceDateByRepeat(report.startDate, report.repeat) : "",
    status: "Not started",
    bucket: "Upcoming",
    linkedTaskCreated: false,
    checklist: (report.checklist || []).map((c) => ({ ...c, done: false })),
    createdAt: new Date().toISOString().slice(0, 10),
  };
}

const checklistProgress = (report) => {
  const items = report.checklist || [];
  return { done: items.filter((i) => i.done).length, total: items.length };
};
const isOverdue = (report) => report.dueDate && report.status !== "Completed" && new Date(report.dueDate) < new Date(new Date().toDateString());
const isAtRisk = (report) => {
  if (!report.dueDate || report.status === "Completed") return false;
  if (isOverdue(report)) return false;
  if (report.bucket !== "Upcoming") return false;
  const daysUntilDue = (new Date(report.dueDate) - new Date(new Date().toDateString())) / 86400000;
  return daysUntilDue >= 0 && daysUntilDue <= 14;
};

// Backlog, Upcoming, Up next, and Overdue are date-driven — recomputed live every time
// the board is viewed, based on today's date vs. the due date. Anything moved into
// In progress, Complete, or Submitted reflects real human progress and is never
// overridden by this — only these four "not yet actively worked" buckets are affected.
const PASSIVE_REPORT_BUCKETS = ["Upcoming", "Up next", "Overdue"];
function effectiveReportBucket(report) {
  if (!PASSIVE_REPORT_BUCKETS.includes(report.bucket)) return report.bucket;
  if (!report.dueDate) return report.bucket;
  if (report.status === "Completed") return report.bucket;
  const today = new Date(new Date().toDateString());
  const due = new Date(report.dueDate + "T00:00:00");
  if (due < today) return "Overdue";
  const daysUntil = (due - today) / 86400000;
  return daysUntil > 90 ? "Upcoming" : "Up next";
}

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

// Closed-by-default multi-select with a search box — used wherever a picklist
// could grow unbounded over time (e.g. linking budgets) and an always-open
// checkbox list would become unwieldy.
function MultiSelectDropdown({ options, selectedIds, onChange, placeholder = "Select…", renderLabel }) {
  const [open, setOpen] = useState(false);
  const [search, setSearch] = useState("");
  const ref = useRef(null);

  useEffect(() => {
    if (!open) return;
    const handler = (e) => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); };
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, [open]);

  const filtered = options.filter((o) => renderLabel(o).toLowerCase().includes(search.toLowerCase()));
  const toggle = (id) => onChange(selectedIds.includes(id) ? selectedIds.filter((x) => x !== id) : [...selectedIds, id]);
  const selectedOptions = options.filter((o) => selectedIds.includes(o.id));

  const summary = selectedIds.length === 0
    ? placeholder
    : selectedIds.length <= 2
      ? selectedOptions.map(renderLabel).join(", ")
      : `${selectedIds.length} selected`;

  return (
    <div className="relative" ref={ref}>
      <button
        type="button"
        onClick={() => setOpen((v) => !v)}
        className={inputCls}
        style={{ ...inputStyle, textAlign: "left", display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8 }}
      >
        <span className="truncate" style={{ color: selectedIds.length === 0 ? "#8A8F87" : "#1C2624" }}>{summary}</span>
        <span style={{ color: "#8A8F87", flexShrink: 0 }}>▾</span>
      </button>
      {open && (
        <div className="absolute z-30 mt-1 w-full rounded-md border bg-white shadow-lg" style={{ borderColor: "#E1E5DE" }}>
          <div className="p-2 border-b" style={{ borderColor: "#E1E5DE" }}>
            <input
              autoFocus
              className={inputCls}
              style={inputStyle}
              placeholder="Search…"
              value={search}
              onChange={(e) => setSearch(e.target.value)}
            />
          </div>
          <div className="max-h-52 overflow-y-auto p-1">
            {filtered.length === 0 ? (
              <p className="text-xs px-2 py-2" style={{ color: "#8A8F87" }}>No matches.</p>
            ) : filtered.map((o) => (
              <label key={o.id} className="flex items-center gap-2 text-sm px-2 py-1.5 rounded hover:bg-stone-50 cursor-pointer" style={{ color: "#1C2624" }}>
                <input type="checkbox" checked={selectedIds.includes(o.id)} onChange={() => toggle(o.id)} />
                {renderLabel(o)}
              </label>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

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

// Captures a rendered chart's <svg> element as a PNG, so an Excel export can
// embed an actual picture of the chart rather than just its underlying data
// — ExcelJS can embed images but has no support for creating live, editable
// Excel chart objects.
function svgElementToPngBase64(svgEl, widthPx, heightPx) {
  return new Promise((resolve, reject) => {
    if (!svgEl) { resolve(null); return; }
    try {
      const clone = svgEl.cloneNode(true);
      clone.setAttribute("width", String(widthPx));
      clone.setAttribute("height", String(heightPx));
      const svgString = new XMLSerializer().serializeToString(clone);
      const svgBlob = new Blob([svgString], { type: "image/svg+xml;charset=utf-8" });
      const url = URL.createObjectURL(svgBlob);
      const img = new Image();
      img.onload = () => {
        const scale = 2; // render at 2x for a crisper embedded image
        const canvas = document.createElement("canvas");
        canvas.width = widthPx * scale;
        canvas.height = heightPx * scale;
        const ctx = canvas.getContext("2d");
        ctx.fillStyle = "#ffffff";
        ctx.fillRect(0, 0, canvas.width, canvas.height);
        ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
        URL.revokeObjectURL(url);
        resolve(canvas.toDataURL("image/png").split(",")[1]);
      };
      img.onerror = () => { URL.revokeObjectURL(url); resolve(null); };
      img.src = url;
    } catch (e) {
      resolve(null);
    }
  });
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
  const widthClass = size === "xl" ? "w-[97vw] max-w-[2000px]" : wide ? "max-w-4xl" : "max-w-lg";
  return (
    <div className="fixed inset-0 z-50 flex items-start justify-center overflow-y-auto overflow-x-hidden py-8 px-4" style={{ background: "rgba(28,38,36,0.45)" }}>
      <div className={`bg-white rounded-xl shadow-xl w-full min-w-0 ${widthClass} my-auto`}>
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
  const [form, setForm, undoForm, canUndoForm] = useUndoableState(grant || {
    id: uid(), title: "", programCode: "", funding: "", sites: [], stage: "Prospecting",
    awardAmount: 0, start: "", end: "", riskStatus: "Low", cadence: [],
    complianceOwner: "", financeOwner: "", internalOwner: "", operationsOwner: "", renewal: false,
    doclibUrl: "", contractUrl: "", coverPageUrl: "", notes: "",
    budgetPeriodStart: "", budgetPeriodEnd: "", obligatedFunds: 0, obligatedFundsRemaining: 0, paymentMethod: PAYMENT_METHODS[0],
    beds: "", bedRate: 0, grantPoc: "", awardAmountRemaining: 0, budgetGroupId: "", indirectRate: 0,
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
      <fieldset disabled={!canEdit} style={{ border: "none", margin: 0, padding: 0, minWidth: 0 }}>
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
        <Field label="Indirect cost rate (%)">
          <input type="number" step="0.01" className={inputCls} style={inputStyle} value={form.indirectRate} onChange={set("indirectRate")} placeholder="e.g. 15" />
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
        <Field label="Grant cover page URL (SharePoint)">
          <input className={inputCls} style={inputStyle} value={form.coverPageUrl} onChange={set("coverPageUrl")} placeholder="https://…sharepoint.com/…" />
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
        {canEdit && canUndoForm && (
          <button onClick={undoForm} className="px-3 py-2 rounded-md text-sm border inline-flex items-center gap-1.5 mr-auto" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>
            <Undo2 size={14} /> Undo
          </button>
        )}
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

function BudgetModal({ budget, grantId, costCenterId, canEdit = true, onSave, onClose, currentUserEmail, grants = [], budgets = [] }) {
  const initial = budget
    ? { ...budget, linkedBudgetIds: budget.linkedBudgetIds || (budget.linkedBudgetId ? [budget.linkedBudgetId] : []) }
    : {
      id: uid(), grantId, costCenterId, title: "", fy: "", periodStart: "", periodEnd: "",
      status: "Draft", notes: "", lines: [newLine()], budgetType: "", linkedBudgetIds: [],
      approvedBy: "", approvedAt: "", rejectionReason: "",
    };
  const [form, setForm, undoForm, canUndoForm] = useUndoableState(initial);
  const [approverName, setApproverName] = useState("");
  const [rejectReason, setRejectReason] = useState("");
  const [showRejectBox, setShowRejectBox] = useState(false);
  const set = (k) => (e) => setForm({ ...form, [k]: e.target.value });
  const cols = monthColumnsForBudget(form.periodStart, form.periodEnd);
  // Plain 4-digit years only (e.g. "2026", never "FY26") — keeps every budget's
  // fy field in a consistent, comparable format for year-based filtering.
  const fyYearOptions = Array.from({ length: 8 }, (_, i) => new Date().getFullYear() - 2 + i);

  useEffect(() => {
    const needsResize = form.lines.some((l) => (l.amounts?.length || 0) !== cols.length || (l.actuals?.length || 0) !== cols.length);
    if (!needsResize) return;
    setForm((f) => ({
      ...f,
      lines: f.lines.map((l) => ({
        ...l,
        amounts: resizeMonthlyArray(l.amounts, cols.length),
        actuals: resizeMonthlyArray(l.actuals, cols.length),
      })),
    }));
  }, [cols.length]);

  const updateLine = (id, patch) => {
    setForm({ ...form, lines: form.lines.map((l) => (l.id === id ? { ...l, ...patch } : l)) });
  };
  const updateAmount = (id, idx, val, field = "amounts") => {
    setForm({
      ...form,
      lines: form.lines.map((l) => {
        if (l.id !== id) return l;
        const arr = [...(l[field] || Array(cols.length).fill(0))];
        arr[idx] = val === "" ? 0 : Number(val);
        return { ...l, [field]: arr };
      }),
    });
  };
  const addLine = () => setForm({ ...form, lines: [...form.lines, newLine(cols.length)] });
  const deleteLine = (id) => setForm({ ...form, lines: form.lines.filter((l) => l.id !== id) });
  const [lineSort, setLineSort] = useState("none");
  const [yearlyDraft, setYearlyDraft] = useState({});
  const [mode, setMode] = useState("plan");
  const field = mode === "plan" ? "amounts" : "actuals";
  // Fixed pixel widths for every column — must stay in sync with the <th>/<td>
  // widths below. table-layout: fixed locks columns to these exact sizes so
  // the frozen (sticky) Category/Subcategory/Description columns never drift
  // out of alignment with their hardcoded `left` offsets, even when a long
  // category name would otherwise make the browser widen that column.
  const COL_W = { category: 160, subcategory: 200, description: 120, annual: 96, month: 92, variance: 60, total: 80, totalVariance: 64, trash: 36 };
  const tableTotalWidth = COL_W.category + COL_W.subcategory + COL_W.description + COL_W.annual
    + cols.length * (COL_W.month + (mode === "actual" ? COL_W.variance : 0))
    + COL_W.total + (mode === "actual" ? COL_W.totalVariance : 0) + COL_W.trash;

  const applyYearlyTotal = (id) => {
    const val = yearlyDraft[id];
    if (val === undefined || val === "") return;
    updateLine(id, { [field]: distributeEvenly(val, cols.length) });
  };

  // If this grant has a negotiated indirect cost rate on file, a line whose
  // category is some flavor of "Indirect Billing" can be auto-calculated as
  // that rate applied to its matching direct-revenue sibling line, instead of
  // hand-computing the percentage every renewal year.
  const grant = grants.find((g) => g.id === form.grantId);
  const calcIndirect = (lineId) => {
    const line = form.lines.find((l) => l.id === lineId);
    if (!line || !grant?.indirectRate) return;
    const baseCategory = line.category.replace(/\s*Indirect Billing\s*$/i, "").trim();
    const directLine = form.lines.find((l) => l.id !== lineId && l.category.trim() === baseCategory);
    if (!directLine) return;
    const directTotal = mode === "plan" ? lineTotal(directLine) : lineActualTotal(directLine);
    const indirectTotal = Math.round(directTotal * (Number(grant.indirectRate) / 100) * 100) / 100;
    updateLine(lineId, { [field]: distributeEvenly(indirectTotal, cols.length) });
    setYearlyDraft((d) => ({ ...d, [lineId]: "" }));
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
        {(form.status === "Active" || form.status === "Awarded") && (
          <div className="text-sm" style={{ color: "#2F6F53" }}>
            <CheckCircle size={13} className="inline mr-1" style={{ marginBottom: 2 }} />
            {form.status === "Awarded" ? "Awarded" : "Approved"}{form.approvedBy ? ` by ${form.approvedBy}` : ""}{form.approvedAt ? ` on ${fmtDate(form.approvedAt)}` : ""}
          </div>
        )}
        {form.status === "Closed" && (
          <div className="text-sm" style={{ color: "#8A8F87" }}>This budget is closed.</div>
        )}
      </div>

      <fieldset disabled={!canEdit} style={{ border: "none", margin: 0, padding: 0, minWidth: 0 }}>
      <div className="grid grid-cols-2 sm:grid-cols-5 gap-4 mb-4">
        <Field label="Budget title">
          <input className={inputCls} style={inputStyle} value={form.title} onChange={set("title")} placeholder="e.g. FY26 Operating Budget" />
        </Field>
        <Field label="Fiscal year">
          <select className={inputCls} style={inputStyle} value={form.fy} onChange={set("fy")}>
            <option value="">Select year</option>
            {fyYearOptions.map((y) => <option key={y} value={String(y)}>{y}</option>)}
            {form.fy && !fyYearOptions.includes(Number(form.fy)) && (
              <option value={form.fy}>{form.fy} (non-standard — re-select to fix)</option>
            )}
          </select>
        </Field>
        <Field label="Budget type">
          <select className={inputCls} style={inputStyle} value={form.budgetType || ""} onChange={set("budgetType")}>
            <option value="">Select type</option>
            {BUDGET_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}
          </select>
        </Field>
        <Field label="Status">
          <select
            className={inputCls}
            style={inputStyle}
            value={form.status}
            onChange={(e) => {
              const newStatus = e.target.value;
              // Most budgets get their status set directly here rather than
              // through the formal Submit/Approve flow above. Capture who
              // approved it and when the moment it goes Active/Awarded,
              // unless it's already been formally approved — never overwrite
              // an existing approval record.
              if ((newStatus === "Active" || newStatus === "Awarded") && !form.approvedBy) {
                setForm({ ...form, status: newStatus, approvedBy: currentUserEmail || "Unknown", approvedAt: new Date().toISOString().slice(0, 10) });
              } else {
                setForm({ ...form, status: newStatus });
              }
            }}
          >
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

      {form.budgetType === "Template" && (
        <div className="mb-4">
          <Field label="Linked Operational budgets">
            <MultiSelectDropdown
              options={budgets.filter((b) => {
                if (b.id === form.id || (!form.grantId && !form.costCenterId)) return false;
                if (!(b.grantId === form.grantId || b.costCenterId === form.costCenterId)) return false;
                const isLinkable = b.status === "Draft" || b.status === "Active" || b.status === "Closed";
                const alreadyLinked = (form.linkedBudgetIds || []).includes(b.id);
                return isLinkable || alreadyLinked; // keep a since-changed link visible so it can still be unlinked
              })}
              selectedIds={form.linkedBudgetIds || []}
              onChange={(ids) => setForm({ ...form, linkedBudgetIds: ids })}
              placeholder="Select budgets to link…"
              renderLabel={(b) => `${b.title || "Untitled budget"} (${b.fy}${b.budgetType ? `, ${b.budgetType}` : ""})`}
            />
          </Field>
          <p className="text-xs mt-1" style={{ color: "#8A8F87" }}>
            A calendar-year Template can span more than one Operational fiscal year — link every budget this one should feed, including a Closed one if the grant's period overlaps into it. Each linked budget only receives whichever months actually fall within its own period.
          </p>
          <p className="text-xs mt-1" style={{ color: "#8A8F87" }}>
            Whenever this Template budget is saved, its actuals are copied onto every linked budget for the matching calendar months — matched by category/subcategory, not row order. One-way only; editing a linked budget directly won't flow back here.
          </p>
        </div>
      )}

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
        <div className="flex items-center gap-2 shrink-0">
          {mode === "actual" && (
            <label className="flex items-center gap-1.5 text-xs" style={{ color: "#8A8F87" }}>
              Actuals complete through
              <select
                className={inputCls}
                style={{ ...inputStyle, width: "auto", padding: "4px 8px", fontSize: 12 }}
                value={form.actualsThrough || ""}
                onChange={(e) => setForm({ ...form, actualsThrough: e.target.value || "" })}
              >
                <option value="">Not marked</option>
                {cols
                  .filter((c) => {
                    const now = new Date();
                    return c.year < now.getFullYear() || (c.year === now.getFullYear() && c.monthIndex <= now.getMonth());
                  })
                  .map((c, i) => (
                    <option key={i} value={`${c.year}-${String(c.monthIndex + 1).padStart(2, "0")}`}>{c.label}</option>
                  ))}
              </select>
            </label>
          )}
          <div className="inline-flex rounded-md border overflow-hidden" style={{ borderColor: "#E1E5DE" }}>
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
      </div>

      <div className="overflow-x-auto border rounded-lg" style={{ borderColor: "#E1E5DE" }}>
        <table className="text-xs" style={{ fontFamily: "var(--mono-font)", width: tableTotalWidth, tableLayout: "fixed" }}>
          <thead>
            <tr style={{ background: "#F6F7F3" }}>
              <th className="text-left px-2 py-2 sticky left-0 z-20" style={{ background: "#F6F7F3", width: COL_W.category }}>Category</th>
              <th className="text-left px-2 py-2 sticky z-10" style={{ left: COL_W.category, background: "#F6F7F3", width: COL_W.subcategory }}>Subcategory</th>
              <th className="text-left px-2 py-2 sticky z-10" style={{ left: COL_W.category + COL_W.subcategory, background: "#F6F7F3", width: COL_W.description }}>Description</th>
              <th className="text-right px-2 py-2" style={{ width: COL_W.annual }}>Annual total</th>
              {cols.map((col, i) => (
                <Fragment key={i}>
                  <th className="text-right px-2 py-2" style={{ width: COL_W.month }}>{col.label}</th>
                  {mode === "actual" && <th className="text-right px-2 py-2" style={{ width: COL_W.variance, color: "#8A8F87", fontWeight: 400 }}>% var</th>}
                </Fragment>
              ))}
              <th className="text-right px-2 py-2" style={{ width: COL_W.total }}>Total</th>
              {mode === "actual" && <th className="text-right px-2 py-2" style={{ width: COL_W.totalVariance, color: "#8A8F87", fontWeight: 400 }}>% var</th>}
              <th className="px-2 py-2" style={{ width: COL_W.trash }}></th>
            </tr>
          </thead>
          <tbody>
            {form.lines.length === 0 && (
              <tr><td colSpan={cols.length + 6 + (mode === "actual" ? cols.length + 1 : 0)} className="text-center py-6" style={{ color: "#8A8F87" }}>No expense lines yet.</td></tr>
            )}
            {sortedLines.map((line) => {
              const cat = CATEGORIES.find((c) => c.name === line.category);
              const values = line[field] || Array(12).fill(0);
              return (
                <tr key={line.id} className="border-t" style={{ borderColor: "#E1E5DE" }}>
                  <td className="px-2 py-1.5 sticky left-0 z-20 bg-white">
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
                  <td className="px-2 py-1.5 sticky z-10 bg-white" style={{ left: COL_W.category }}>
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
                  <td className="px-2 py-1.5 sticky z-10 bg-white" style={{ left: COL_W.category + COL_W.subcategory }}>
                    <input
                      value={line.description || ""}
                      onChange={(e) => updateLine(line.id, { description: e.target.value })}
                      placeholder="Optional note"
                      className="w-full rounded border px-1.5 py-1 text-xs"
                      style={inputStyle}
                    />
                  </td>
                  <td className="px-2 py-1.5">
                    <div className="flex items-center gap-1">
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
                      {grant?.indirectRate > 0 && /indirect/i.test(line.category) && (
                        <button
                          type="button"
                          onClick={() => calcIndirect(line.id)}
                          title={`Calculate as ${grant.indirectRate}% of the matching direct revenue line`}
                          className="shrink-0"
                          style={{ color: "#1F5C6B" }}
                        >
                          <RefreshCw size={11} />
                        </button>
                      )}
                    </div>
                  </td>
                  {values.map((amt, idx) => {
                    const v = mode === "actual" ? varianceInfo(line.amounts?.[idx], line.actuals?.[idx]) : null;
                    return (
                      <Fragment key={idx}>
                        <td className="px-1 py-1.5">
                          <input
                            type="number"
                            value={amt === 0 ? "" : amt}
                            placeholder="0"
                            onChange={(e) => updateAmount(line.id, idx, e.target.value, field)}
                            className="w-full rounded border px-1.5 py-1 text-xs text-right"
                            style={{ ...inputStyle, fontVariantNumeric: "tabular-nums" }}
                          />
                        </td>
                        {mode === "actual" && (
                          <td className="px-2 py-1.5 text-right text-xs" style={{ fontVariantNumeric: "tabular-nums", color: v.color }}>
                            {v.label}
                          </td>
                        )}
                      </Fragment>
                    );
                  })}
                  <td className="px-2 py-1.5 text-right font-medium" style={{ fontVariantNumeric: "tabular-nums", color: line.type === "revenue" ? "#2F6F53" : "#1C2624" }}>
                    {fmt(mode === "plan" ? lineTotal(line) : lineActualTotal(line))}
                  </td>
                  {mode === "actual" && (() => {
                    const overall = varianceInfo(lineTotal(line), lineActualTotal(line));
                    return (
                      <td className="px-2 py-1.5 text-right text-xs font-medium" style={{ fontVariantNumeric: "tabular-nums", color: overall.color }}>
                        {overall.label}
                      </td>
                    );
                  })()}
                  <td className="px-2 py-1.5 text-center">
                    <button onClick={() => deleteLine(line.id)} className="p-1 rounded hover:bg-red-50">
                      <Trash2 size={14} style={{ color: "#B5443A" }} />
                    </button>
                  </td>
                </tr>
              );
            })}
            {form.lines.length > 0 && (
              <tr className="border-t-2" style={{ borderColor: "#1C2624" }}>
                <td className="px-2 py-1.5 sticky left-0 z-20 font-medium text-xs" style={{ background: "#F6F7F3", color: "#1C2624" }}>Monthly net</td>
                <td className="px-2 py-1.5 sticky z-10" style={{ left: COL_W.category, background: "#F6F7F3" }}></td>
                <td className="px-2 py-1.5 sticky z-10" style={{ left: COL_W.category + COL_W.subcategory, background: "#F6F7F3" }}></td>
                <td className="px-2 py-1.5" style={{ background: "#F6F7F3" }}></td>
                {cols.map((_, i) => {
                  const net = form.lines.reduce((sum, l) => {
                    const v = Number((l[field] || [])[i]) || 0;
                    if (l.type === "revenue") return sum + v;
                    if (l.type === "expense") return sum - v;
                    return sum; // balance-sheet-type lines aren't income statement activity
                  }, 0);
                  return (
                    <Fragment key={i}>
                      <td className="px-2 py-1.5 text-right text-xs font-medium" style={{ fontVariantNumeric: "tabular-nums", color: net >= 0 ? "#2F6F53" : "#B5443A", background: "#F6F7F3" }}>
                        {fmt(net)}
                      </td>
                      {mode === "actual" && <td style={{ background: "#F6F7F3" }}></td>}
                    </Fragment>
                  );
                })}
                <td className="px-2 py-1.5" style={{ background: "#F6F7F3" }}></td>
                {mode === "actual" && <td style={{ background: "#F6F7F3" }}></td>}
                <td className="px-2 py-1.5" style={{ background: "#F6F7F3" }}></td>
              </tr>
            )}
          </tbody>
        </table>
      </div>

      <div className="grid grid-cols-1 sm:grid-cols-3 gap-6 mt-4 text-sm">
        <div>
          <div className="text-xs font-medium mb-1" style={{ color: "#8A8F87" }}>Plan</div>
          <div><span style={{ color: "#5B6B66" }}>Revenue: </span><span className="font-medium" style={{ color: "#2F6F53" }}>{fmt(totals.revenue)}</span></div>
          <div><span style={{ color: "#5B6B66" }}>Expense: </span><span className="font-medium" style={{ color: "#1C2624" }}>{fmt(totals.expense)}</span></div>
          <div><span style={{ color: "#5B6B66" }}>Net: </span><span className="font-medium" style={{ color: !isNetNegative(totals.net) ? "#2F6F53" : "#B5443A" }}>{fmt(totals.net)}</span></div>
        </div>
        <div>
          <div className="text-xs font-medium mb-1" style={{ color: "#8A8F87" }}>Actual</div>
          <div><span style={{ color: "#5B6B66" }}>Revenue: </span><span className="font-medium" style={{ color: "#2F6F53" }}>{fmt(actualTotals.revenue)}</span></div>
          <div><span style={{ color: "#5B6B66" }}>Expense: </span><span className="font-medium" style={{ color: "#1C2624" }}>{fmt(actualTotals.expense)}</span></div>
          <div><span style={{ color: "#5B6B66" }}>Net: </span><span className="font-medium" style={{ color: !isNetNegative(actualTotals.net) ? "#2F6F53" : "#B5443A" }}>{fmt(actualTotals.net)}</span></div>
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
        {canEdit && canUndoForm && (
          <button onClick={undoForm} className="px-3 py-2 rounded-md text-sm border inline-flex items-center gap-1.5 mr-auto" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>
            <Undo2 size={14} /> Undo
          </button>
        )}
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
  const [form, setForm, undoForm, canUndoForm] = useUndoableState(report || {
    id: uid(), title: "", grantId: "", assignedTo: "", status: "Not started",
    priority: "Medium", startDate: "", dueDate: "", repeat: "None", repeatDetail: "",
    bucket: DEFAULT_BUCKETS[0], checklist: [], notes: "", portalUrl: "", supportingDocsUrl: "", linkedTaskCreated: false,
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
      <fieldset disabled={!canEdit} style={{ border: "none", margin: 0, padding: 0, minWidth: 0 }}>
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
        <Field label="Supporting documents (SharePoint folder URL)">
          <input className={inputCls} style={inputStyle} value={form.supportingDocsUrl || ""} onChange={set("supportingDocsUrl")} placeholder="https://…" />
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
          {canEdit && canUndoForm && (
            <button onClick={undoForm} className="px-3 py-2 rounded-md text-sm border inline-flex items-center gap-1.5" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>
              <Undo2 size={14} /> Undo
            </button>
          )}
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

function Dashboard({ grants, budgets, reports, tasks, staff, invoices, goTo, costCenters, budgetGroups }) {
  const activeGrants = grants.filter((g) => g.stage === "Active");
  const totalAward = activeGrants.reduce((a, g) => a + (Number(g.awardAmount) || 0), 0);
  const totalRemaining = activeGrants.reduce((a, g) => {
    const t = grantCurrentPeriodBudgetTotals(g.id, budgets);
    return a + ((Number(g.awardAmount) || 0) - t.expense);
  }, 0);
  const activeGrantIds = new Set(activeGrants.map((g) => g.id));
  const activeBudgets = budgets.filter((b) => isActiveBudget(b.status) && activeGrantIds.has(b.grantId)).length;

  const fyYear = new Date().getFullYear();
  const fyStart = new Date(fyYear, 0, 1);
  const fyEnd = new Date(fyYear, 11, 31);
  const fyGrants = activeGrants.filter((g) => {
    const start = g.start ? new Date(g.start) : null;
    const end = g.end ? new Date(g.end) : null;
    return (!start || start <= fyEnd) && (!end || end >= fyStart);
  });
  // Sourced from the Org Budget page's own math (Whole Organization scope,
  // Template budgets) rather than grant-level award fields, so this always
  // matches what Org Budget itself shows for the same fiscal year.
  const { planRevenue: fyTotalAward, actualExpense: fyActualExpense } = computeOrgFYTotals(budgets, grants, costCenters, budgetGroups, fyYear);
  const fyTotalRemaining = fyTotalAward - fyActualExpense;
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
          <StatCard label={`FY${fyYear} total award`} value={fmt(fyTotalAward)} sub={`Planned revenue, Org Budget (Template) — ${fyGrants.length} grants active in FY${fyYear}`} />
          <StatCard label={`FY${fyYear} award remaining`} value={fmt(fyTotalRemaining)} sub="Planned revenue minus actual expense, Org Budget" />
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

      <div className="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-8 gap-3">
        {STAGES.map((s) => {
          const count = grants.filter((g) => g.stage === s).length;
          const active = stageFilter === s;
          return (
            <button
              key={s}
              onClick={() => setStageFilter(active ? "All" : s)}
              className="text-left bg-white rounded-lg border p-3"
              style={{ borderColor: active ? stageColor[s] : "#E1E5DE", borderWidth: active ? 2 : 1 }}
            >
              <div className="text-xs" style={{ color: stageColor[s] }}>{s}</div>
              <div className="text-xl font-medium mt-0.5" style={{ color: "#1C2624", fontVariantNumeric: "tabular-nums" }}>{count}</div>
            </button>
          );
        })}
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
            const totals = grantCurrentPeriodBudgetTotals(g.id, budgets);
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
                      {g.coverPageUrl && (
                        <a href={g.coverPageUrl} target="_blank" rel="noreferrer" className="inline-flex items-center gap-1 text-xs px-3 py-1.5 rounded-md border" style={{ borderColor: "#E1E5DE", color: "#1F5C6B" }}>
                          <ExternalLink size={13} /> Grant cover page
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
  const [duplicatePrompt, setDuplicatePrompt] = useState(null); // the budget being duplicated, or null
  const [overviewSearch, setOverviewSearch] = useState("");
  const [overviewSort, setOverviewSort] = useState({ key: "title", dir: "asc" });
  const [showExcelImport, setShowExcelImport] = useState(false);

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
      let next = exists ? prev.map((x) => (x.id === b.id ? b : x)) : [...prev, b];
      const syncedTargets = syncActualsToLinkedBudgets(b, next);
      syncedTargets.forEach((synced) => {
        logActivity?.("Budget", "Updated", `${synced.title || "Untitled budget"} (actuals synced from "${b.title || "Untitled budget"}")`);
        next = next.map((x) => (x.id === synced.id ? synced : x));
      });
      return next;
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
  const duplicateBudget = (budget, carryOverAmounts) => {
    const newBudget = {
      ...budget,
      id: uid(),
      fy: nextFyLabel(budget.fy),
      periodStart: shiftYear(budget.periodStart),
      periodEnd: shiftYear(budget.periodEnd),
      status: "Draft",
      approvedBy: "", approvedAt: "", rejectionReason: "",
      lines: budget.lines.map((l) => ({
        ...l, id: uid(),
        amounts: carryOverAmounts ? [...l.amounts] : Array(l.amounts.length).fill(0),
        actuals: Array(l.amounts.length).fill(0),
      })),
    };
    setBudgets((prev) => [...prev, newBudget]);
    logActivity?.("Budget", "Created", `${newBudget.title || "Untitled budget"}${activeSelection ? ` (${activeSelection.title || activeSelection.name})` : ""} — rolled over from ${budget.fy || "prior year"}${carryOverAmounts ? " (carried over amounts)" : ""}`);
    setModal(newBudget);
    setDuplicatePrompt(null);
  };
  const exportCsv = (budget) => {
    const labels = monthColumnsForBudget(budget.periodStart, budget.periodEnd).map((c) => c.label);
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
    const cols = monthColumnsForBudget(budget.periodStart, budget.periodEnd);
    const labels = cols.map((c) => c.label);
    const rows = [
      [label],
      [`${budget.title}${budget.fy ? ` (${budget.fy})` : ""}`],
      [`Period: ${fmtDate(budget.periodStart)} – ${fmtDate(budget.periodEnd)}`, `Status: ${budget.status}`],
      [],
      ["Category", "Subcategory", "Description", "Type", ...labels, "Total"],
      ...budget.lines.map((l) => [l.category, l.subcategory || "", l.description || "", l.type, ...l.amounts, lineTotal(l)]),
      [],
      ["Total Revenue", "", "", "", ...Array(cols.length).fill(""), t.revenue],
      ["Total Expense", "", "", "", ...Array(cols.length).fill(""), t.expense],
      ["Net", "", "", "", ...Array(cols.length).fill(""), t.net],
    ];
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"] = [{ wch: 30 }, { wch: 18 }, { wch: 24 }, { wch: 10 }, ...cols.map(() => ({ wch: 12 })), { wch: 14 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Budget");
    const safe = (s) => (s || "budget").replace(/[^a-z0-9]+/gi, "_").slice(0, 40);
    const arrayBuffer = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    downloadFile(`${safe(g?.title || cc?.name)}-${safe(budget.title)}.xlsx`, arrayBuffer, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  };

  const allBudgetsEnriched = useMemo(() => budgets.map((b) => {
    const g = b.grantId ? grants.find((x) => x.id === b.grantId) : null;
    const cc = b.costCenterId ? costCenters.find((x) => x.id === b.costCenterId) : null;
    const t = budgetTotals(b);
    return {
      ...b,
      ownerName: g ? (g.programCode ? `${g.programCode} - ${g.title}` : g.title) : cc ? cc.name : "Unknown",
      ownerType: g ? "grant" : "costCenter",
      netTotal: t.revenue - t.expense,
    };
  }), [budgets, grants, costCenters]);

  const statusCounts = useMemo(() => {
    const counts = {};
    BUDGET_STATUSES.forEach((s) => { counts[s] = 0; });
    allBudgetsEnriched.forEach((b) => { counts[b.status] = (counts[b.status] || 0) + 1; });
    return counts;
  }, [allBudgetsEnriched]);

  const needsAttention = allBudgetsEnriched.filter((b) => b.status === "Pending Approval");

  const overviewRows = useMemo(() => {
    const q = overviewSearch.trim().toLowerCase();
    let rows = allBudgetsEnriched.filter((b) =>
      !q || b.title.toLowerCase().includes(q) || b.ownerName.toLowerCase().includes(q) || (b.fy || "").toLowerCase().includes(q)
    );
    rows = [...rows].sort((a, b) => {
      const dir = overviewSort.dir === "asc" ? 1 : -1;
      const av = a[overviewSort.key], bv = b[overviewSort.key];
      if (typeof av === "string") return av.localeCompare(bv) * dir;
      return ((av ?? 0) - (bv ?? 0)) * dir;
    });
    return rows;
  }, [allBudgetsEnriched, overviewSearch, overviewSort]);

  const toggleOverviewSort = (key) => setOverviewSort((s) => (s.key === key ? { key, dir: s.dir === "asc" ? "desc" : "asc" } : { key, dir: "asc" }));

  const openBudget = (b) => {
    setBudgetMode(b.ownerType === "grant" ? "grant" : "costCenter");
    if (b.ownerType === "grant") setSelectedGrantId(b.grantId); else setSelectedCostCenterId(b.costCenterId);
    setModal(b);
  };

  return (
    <div className="space-y-4">
      <div className="flex items-center justify-between">
        <h1 className="font-display text-2xl" style={{ color: "#1C2624" }}>Budgets</h1>
        <div className="flex items-center gap-2">
          {canEdit && (
            <button onClick={() => setShowExcelImport(true)} className="inline-flex items-center gap-1.5 px-3 py-2 rounded-md text-sm border" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>
              <Upload size={16} /> Import Excel Template
            </button>
          )}
          {activeSelection && canEdit && (
            <button onClick={() => setModal("new")} className="inline-flex items-center gap-1.5 px-3 py-2 rounded-md text-sm text-white" style={{ background: "#1F5C6B" }}>
              <Plus size={16} /> New budget
            </button>
          )}
        </div>
      </div>

      <div className="grid grid-cols-3 sm:grid-cols-6 gap-3">
        <StatCard label="Total budgets" value={allBudgetsEnriched.length} />
        <StatCard label="Draft" value={statusCounts["Draft"] || 0} />
        <StatCard label="Pending approval" value={statusCounts["Pending Approval"] || 0} />
        <StatCard label="Active" value={statusCounts["Active"] || 0} />
        <StatCard label="Awarded" value={statusCounts["Awarded"] || 0} />
        <StatCard label="Closed" value={statusCounts["Closed"] || 0} />
      </div>

      {needsAttention.length > 0 && (
        <div className="bg-white rounded-lg border p-4" style={{ borderColor: "#C08A2E" }}>
          <h2 className="font-display text-sm mb-2" style={{ color: "#C08A2E" }}>Needs attention — {needsAttention.length} budget{needsAttention.length === 1 ? "" : "s"} awaiting approval</h2>
          <div className="divide-y" style={{ borderColor: "#E1E5DE" }}>
            {needsAttention.map((b) => (
              <button
                key={b.id}
                onClick={() => openBudget(b)}
                className="w-full text-left py-2 flex items-center justify-between text-sm hover:bg-stone-50"
              >
                <span style={{ color: "#1C2624" }}>{b.title || "Untitled budget"} <span style={{ color: "#8A8F87" }}>— {b.ownerName}</span></span>
                <span style={{ color: "#8A8F87" }}>{b.fy}</span>
              </button>
            ))}
          </div>
        </div>
      )}

      <div className="bg-white rounded-lg border p-4" style={{ borderColor: "#E1E5DE" }}>
        <div className="flex items-center justify-between mb-3">
          <h2 className="font-display text-sm" style={{ color: "#1C2624" }}>All budgets</h2>
          <input
            value={overviewSearch}
            onChange={(e) => setOverviewSearch(e.target.value)}
            placeholder="Search by title, grant, or FY…"
            className={inputCls}
            style={{ ...inputStyle, maxWidth: 280 }}
          />
        </div>
        {overviewRows.length === 0 ? (
          <p className="text-sm py-4 text-center" style={{ color: "#8A8F87" }}>No budgets match your search.</p>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead>
                <tr style={{ color: "#8A8F87" }}>
                  <th className="text-left py-1.5 font-medium cursor-pointer" onClick={() => toggleOverviewSort("ownerName")}>Grant / Cost Center</th>
                  <th className="text-left py-1.5 font-medium cursor-pointer" onClick={() => toggleOverviewSort("title")}>Budget</th>
                  <th className="text-left py-1.5 font-medium cursor-pointer" onClick={() => toggleOverviewSort("fy")}>FY</th>
                  <th className="text-left py-1.5 font-medium cursor-pointer" onClick={() => toggleOverviewSort("status")}>Status</th>
                  <th className="text-left py-1.5 font-medium cursor-pointer" onClick={() => toggleOverviewSort("budgetType")}>Type</th>
                  <th className="text-right py-1.5 font-medium cursor-pointer" onClick={() => toggleOverviewSort("netTotal")}>Net</th>
                </tr>
              </thead>
              <tbody>
                {overviewRows.map((b) => (
                  <tr
                    key={b.id}
                    onClick={() => openBudget(b)}
                    className="border-t cursor-pointer hover:bg-stone-50"
                    style={{ borderColor: "#E1E5DE" }}
                  >
                    <td className="py-1.5" style={{ color: "#1C2624" }}>{b.ownerName}</td>
                    <td className="py-1.5" style={{ color: "#1C2624" }}>{b.title || "Untitled budget"}</td>
                    <td className="py-1.5" style={{ color: "#8A8F87" }}>{b.fy}</td>
                    <td className="py-1.5">
                      <Badge color={isActiveBudget(b.status) ? "#2F6F53" : b.status === "Pending Approval" ? "#C08A2E" : b.status === "Rejected" ? "#B5443A" : "#8A8F87"}>{b.status}</Badge>
                    </td>
                    <td className="py-1.5">
                      {b.budgetType ? (
                        <Badge color={b.budgetType === "Template" ? "#5B7FA6" : "#8A8F87"}>{b.budgetType}</Badge>
                      ) : (
                        <span className="text-xs" style={{ color: "#C08A2E" }}>Not set</span>
                      )}
                    </td>
                    <td className="py-1.5 text-right" style={{ fontVariantNumeric: "tabular-nums", color: !isNetNegative(b.netTotal) ? "#2F6F53" : "#B5443A" }}>{fmt(b.netTotal)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
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
                  <Badge color={isActiveBudget(b.status) ? "#2F6F53" : b.status === "Closed" ? "#8A8F87" : b.status === "Rejected" ? "#B5443A" : b.status === "Pending Approval" ? "#C08A2E" : "#5B7FA6"}>{b.status}</Badge>
                </div>
                <div className="flex gap-5 mt-3 text-sm" style={{ fontVariantNumeric: "tabular-nums" }}>
                  <div><span style={{ color: "#8A8F87" }}>Revenue </span><span style={{ color: "#2F6F53" }}>{fmt(t.revenue)}</span></div>
                  <div><span style={{ color: "#8A8F87" }}>Expense </span><span style={{ color: "#1C2624" }}>{fmt(t.expense)}</span></div>
                  <div><span style={{ color: "#8A8F87" }}>Net </span><span style={{ color: !isNetNegative(t.net) ? "#2F6F53" : "#B5443A" }}>{fmt(t.net)}</span></div>
                </div>
                <div className="flex gap-2 mt-4 flex-wrap">
                  <button onClick={() => setModal(b)} className="inline-flex items-center gap-1 text-xs px-3 py-1.5 rounded-md border" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>
                    <Pencil size={13} /> {canEdit ? "View/edit budget" : "View budget"}
                  </button>
                  <button onClick={() => exportCsv(b)} className="inline-flex items-center gap-1 text-xs px-3 py-1.5 rounded-md border" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>
                    <Download size={13} /> Export CSV
                  </button>
                  <button onClick={() => exportXlsx(b)} className="inline-flex items-center gap-1 text-xs px-3 py-1.5 rounded-md border" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>
                    <Download size={13} /> Export Excel
                  </button>
                  {canEdit && (
                    <>
                      <button onClick={() => setDuplicatePrompt(b)} className="inline-flex items-center gap-1 text-xs px-3 py-1.5 rounded-md border" style={{ borderColor: "#E1E5DE", color: "#1F5C6B" }}>
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
          currentUserEmail={currentUserEmail}
          grants={grants}
          budgets={budgets}
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
      {duplicatePrompt && (
        <Modal title="Duplicate to next FY" onClose={() => setDuplicatePrompt(null)}>
          <p className="text-sm mb-4" style={{ color: "#5B6B66" }}>
            Should the new budget start with last year's numbers already filled in, or start blank so you can build it fresh?
          </p>
          <div className="flex flex-col gap-2">
            <button
              onClick={() => duplicateBudget(duplicatePrompt, true)}
              className="px-4 py-3 rounded-md text-sm text-white text-left"
              style={{ background: "#1F5C6B" }}
            >
              Carry over this year's numbers as a starting point
            </button>
            <button
              onClick={() => duplicateBudget(duplicatePrompt, false)}
              className="px-4 py-3 rounded-md text-sm border text-left"
              style={{ borderColor: "#E1E5DE", color: "#1C2624" }}
            >
              Start blank
            </button>
          </div>
        </Modal>
      )}
      {showExcelImport && (
        <ExcelTemplateImportModal
          grants={grants}
          budgets={budgets}
          setBudgets={setBudgets}
          logActivity={logActivity}
          onClose={() => setShowExcelImport(false)}
        />
      )}
    </div>
  );
}

function ExcelTemplateImportModal({ grants, budgets, setBudgets, logActivity, onClose }) {
  const [step, setStep] = useState("upload"); // upload -> review -> done
  const [error, setError] = useState("");
  const [parsed, setParsed] = useState(null); // { projectCode, projectLabel, periodStart, periodEnd, fy, lines, warnings, actualsThrough }
  const [fyOverride, setFyOverride] = useState("");
  const [manualGrantId, setManualGrantId] = useState("");

  const parseFile = async (file) => {
    setError("");
    setParsed(null);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array", cellDates: true });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });

      const result = parseExcelTemplateWorkbook(aoa, sheet);
      if (result.error) { setError(result.error); return; }
      if (result.lines.length === 0) {
        setError("No recognized account lines were found — double-check this is the expected template export.");
        return;
      }

      const matchedGrant = grants.find((g) => g.programCode && String(g.programCode).trim() === result.projectCode);
      const fy = String(result.fy || new Date(result.periodStart).getUTCFullYear());
      const actualsThrough = result.actualsThroughDate
        ? `${result.actualsThroughDate.year}-${String(result.actualsThroughDate.monthIndex + 1).padStart(2, "0")}`
        : "";

      const warnings = [...result.warnings];
      result.unrecognizedCodes.forEach((label) => {
        warnings.push(`"${label}" has an account code not found in GrantFlow's category list — this line was skipped. Add it to CATEGORIES if it should be tracked.`);
      });

      setParsed({
        projectCode: result.projectCode, periodStart: result.periodStart, periodEnd: result.periodEnd,
        fy, lines: result.lines, warnings, actualsThrough, matchedGrantId: matchedGrant?.id || "",
      });
      setFyOverride(fy);
      setManualGrantId(matchedGrant?.id || "");
      setStep("review");
    } catch (e) {
      setError("Couldn't read that file — make sure it's the expected .xlsx template export.");
    }
  };

  const targetGrantId = manualGrantId || parsed?.matchedGrantId || "";
  const existingBudget = parsed
    ? budgets.find((b) => b.grantId === targetGrantId && b.budgetType === "Template" && String(b.fy) === String(fyOverride))
    : null;

  const applyImport = () => {
    if (!targetGrantId) { setError("Select a grant before continuing."); return; }
    const grant = grants.find((g) => g.id === targetGrantId);

    setBudgets((prev) => {
      if (existingBudget) {
        const cols = monthColumnsForBudget(existingBudget.periodStart, existingBudget.periodEnd);
        const colIndexByYM = {};
        cols.forEach((c, i) => { colIndexByYM[`${c.year}-${c.monthIndex}`] = i; });
        const importCols = monthColumnsForBudget(parsed.periodStart, parsed.periodEnd);

        let newLines = existingBudget.lines.map((l) => ({ ...l, amounts: [...(l.amounts || [])], actuals: [...(l.actuals || [])] }));
        parsed.lines.forEach((pl) => {
          let target = newLines.find((l) => l.category === pl.category && l.subcategory === pl.subcategory);
          if (!target) {
            target = { ...newLine(), category: pl.category, subcategory: pl.subcategory, type: pl.type, amounts: Array(cols.length).fill(0), actuals: Array(cols.length).fill(0) };
            newLines.push(target);
          }
          importCols.forEach((col, i) => {
            const j = colIndexByYM[`${col.year}-${col.monthIndex}`];
            if (j === undefined) return;
            if (pl.amounts[i]) target.amounts[j] = pl.amounts[i];
            if (pl.actuals[i]) target.actuals[j] = pl.actuals[i];
          });
        });
        const newCutoff = parsed.actualsThrough;
        const cutoff = newCutoff && (!existingBudget.actualsThrough || newCutoff > existingBudget.actualsThrough) ? newCutoff : existingBudget.actualsThrough;
        const updated = { ...existingBudget, lines: newLines, actualsThrough: cutoff };
        logActivity?.("Budget", "Updated", `${updated.title} (Excel template import)`);
        return prev.map((b) => (b.id === existingBudget.id ? updated : b));
      }

      const newBudget = {
        id: uid(), grantId: targetGrantId, costCenterId: "",
        title: `${grant?.title || "Untitled"} — Template ${fyOverride}`,
        fy: fyOverride, periodStart: parsed.periodStart, periodEnd: parsed.periodEnd,
        status: "Draft", budgetType: "Template", linkedBudgetIds: [],
        notes: "", approvedBy: "", approvedAt: "", rejectionReason: "",
        actualsThrough: parsed.actualsThrough,
        lines: parsed.lines.map((pl) => ({ ...newLine(), category: pl.category, subcategory: pl.subcategory, type: pl.type, amounts: pl.amounts, actuals: pl.actuals })),
      };
      logActivity?.("Budget", "Created", `${newBudget.title} (Excel template import)`);
      return [...prev, newBudget];
    });

    setStep("done");
  };

  return (
    <Modal title="Import Excel Template" onClose={onClose} wide>
      {step === "upload" && (
        <div className="space-y-4">
          <p className="text-sm" style={{ color: "#5B6B66" }}>
            Upload the Rolling 12-Month P&amp;L export for a grant. Account lines with a numeric code (e.g. "4100 - Grants and Contracts") are imported — category headers, "Total" rollups, and individual named sub-lines with no code are skipped automatically.
          </p>
          <Field label="Template export (.xlsx)">
            <input type="file" accept=".xlsx" className={inputCls} style={inputStyle} onChange={(e) => e.target.files[0] && parseFile(e.target.files[0])} />
          </Field>
          {error && <p className="text-xs" style={{ color: "#B5443A" }}>{error}</p>}
          <div className="flex justify-end gap-2 pt-2">
            <button onClick={onClose} className="text-sm px-4 py-2 rounded-md border" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>Cancel</button>
          </div>
        </div>
      )}

      {step === "review" && parsed && (
        <div className="space-y-4">
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
            <Field label="Grant">
              <GrantPicker grants={grants} value={targetGrantId} onChange={setManualGrantId} placeholder="Select a grant" />
            </Field>
            <Field label="Fiscal year">
              <input className={inputCls} style={inputStyle} value={fyOverride} onChange={(e) => setFyOverride(e.target.value)} />
            </Field>
          </div>
          {!parsed.matchedGrantId && (
            <p className="text-xs" style={{ color: "#C08A2E" }}>
              No grant found with program code "{parsed.projectCode}" — select the correct grant above before continuing.
            </p>
          )}
          <p className="text-xs px-3 py-2 rounded-md" style={{ background: existingBudget ? "#EAF1F7" : "#F0F5F2", color: existingBudget ? "#1F5C6B" : "#2F6F53" }}>
            {existingBudget
              ? `Will update the existing Template budget "${existingBudget.title}" for FY${fyOverride} — new lines are added, matching lines are overwritten for the months this file covers.`
              : `Will create a new Template budget for FY${fyOverride}, period ${fmtDate(parsed.periodStart)}–${fmtDate(parsed.periodEnd)}.`}
          </p>
          <p className="text-xs" style={{ color: "#8A8F87" }}>
            {parsed.lines.length} account lines parsed. "Actuals complete through" will be set to {parsed.actualsThrough ? fmtDate(`${parsed.actualsThrough}-01`) : "—"}.
          </p>

          {parsed.warnings.length > 0 && (
            <div>
              <p className="text-xs font-medium mb-1" style={{ color: "#C08A2E" }}>{parsed.warnings.length} formula error(s) in the source file — imported as $0</p>
              <div className="max-h-32 overflow-y-auto text-xs rounded-md border p-2" style={{ borderColor: "#E1E5DE", color: "#8A8F87" }}>
                {parsed.warnings.slice(0, 20).map((w, i) => <div key={i}>{w}</div>)}
              </div>
            </div>
          )}

          <div className="max-h-64 overflow-y-auto rounded-md border" style={{ borderColor: "#E1E5DE" }}>
            <table className="w-full text-xs">
              <thead>
                <tr style={{ background: "#F6F7F3", color: "#5B6B66" }}>
                  <th className="text-left px-2 py-1.5">Category</th>
                  <th className="text-left px-2 py-1.5">Subcategory</th>
                  <th className="text-left px-2 py-1.5">Type</th>
                </tr>
              </thead>
              <tbody>
                {parsed.lines.map((l, i) => (
                  <tr key={i} className="border-t" style={{ borderColor: "#E1E5DE" }}>
                    <td className="px-2 py-1" style={{ color: "#1C2624" }}>{l.category}</td>
                    <td className="px-2 py-1" style={{ color: "#1C2624" }}>{l.subcategory}</td>
                    <td className="px-2 py-1" style={{ color: "#8A8F87" }}>{l.type}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {error && <p className="text-xs" style={{ color: "#B5443A" }}>{error}</p>}
          <div className="flex justify-end gap-2 pt-2">
            <button onClick={() => setStep("upload")} className="text-sm px-4 py-2 rounded-md border" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>Back</button>
            <button onClick={applyImport} className="text-sm px-4 py-2 rounded-md text-white" style={{ background: "#1F5C6B" }}>
              {existingBudget ? "Update budget" : "Create budget"}
            </button>
          </div>
        </div>
      )}

      {step === "done" && (
        <div className="space-y-4 text-center py-6">
          <CheckCircle size={28} style={{ color: "#2F6F53", margin: "0 auto" }} />
          <p className="text-sm" style={{ color: "#1C2624" }}>Template budget {existingBudget ? "updated" : "created"} successfully.</p>
          <button onClick={onClose} className="text-sm px-4 py-2 rounded-md text-white" style={{ background: "#1F5C6B" }}>Done</button>
        </div>
      )}
    </Modal>
  );
}

function ReportCard({ report, grant, onToggleDone, onBucketChange, onEdit }) {
  const progress = checklistProgress(report);
  const overdue = isOverdue(report);
  const atRisk = isAtRisk(report);
  return (
    <div className="bg-white rounded-lg border p-3.5" style={{ borderColor: overdue ? "#B5443A" : atRisk ? "#C08A2E" : "#E1E5DE" }}>
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

      {atRisk && (
        <div className="mt-2 text-xs px-2 py-1 rounded-md inline-flex items-center gap-1" style={{ background: "rgba(192,138,46,0.12)", color: "#C08A2E" }}>
          <AlertCircle size={12} /> Due soon, not started yet
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

      <div className="flex flex-wrap gap-1.5 mt-2.5">
        {report.portalUrl && (
          <a
            href={report.portalUrl} target="_blank" rel="noreferrer"
            onClick={(e) => e.stopPropagation()}
            className="inline-flex items-center gap-1 text-xs px-2 py-1 rounded-md border"
            style={{ borderColor: "#E1E5DE", color: "#1F5C6B" }}
          >
            <ExternalLink size={12} /> Submission portal
          </a>
        )}
        {report.supportingDocsUrl && (
          <a
            href={report.supportingDocsUrl} target="_blank" rel="noreferrer"
            onClick={(e) => e.stopPropagation()}
            className="inline-flex items-center gap-1 text-xs px-2 py-1 rounded-md border"
            style={{ borderColor: "#E1E5DE", color: "#1F5C6B" }}
          >
            <ExternalLink size={12} /> Supporting Documents
          </a>
        )}
      </div>

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
  const [assignedFilter, setAssignedFilter] = useState("All");

  const buckets = [...new Set([...DEFAULT_BUCKETS, ...reports.map((r) => r.bucket)])];
  const assignees = [...new Set(reports.map((r) => (r.assignedTo || "").trim()).filter(Boolean))].sort();
  const visible = reports
    .filter((r) => grantFilter === "All" || r.grantId === grantFilter)
    .filter((r) => assignedFilter === "All" || (r.assignedTo || "").trim() === assignedFilter);

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
    const willComplete = r.status !== "Completed";
    setReports((prev) => {
      let next = prev.map((x) => (x.id === r.id ? { ...x, status: willComplete ? "Completed" : "Not started" } : x));
      if (willComplete) {
        const spawned = spawnNextReportOccurrence(r, grants.find((g) => g.id === r.grantId));
        if (spawned) {
          next = [...next, spawned];
          logActivity?.("Report", "Created", `${spawned.title || "Untitled report"} — next occurrence auto-created`);
        }
      }
      return next;
    });
  };
  const changeBucket = (r, bucket) => {
    const willComplete = bucket === "Complete" && r.bucket !== "Complete";
    setReports((prev) => {
      let next = prev.map((x) => (x.id === r.id ? { ...x, bucket } : x));
      if (willComplete) {
        const spawned = spawnNextReportOccurrence(r, grants.find((g) => g.id === r.grantId));
        if (spawned) {
          next = [...next, spawned];
          logActivity?.("Report", "Created", `${spawned.title || "Untitled report"} — next occurrence auto-created`);
        }
      }
      return next;
    });
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

      <div className="flex gap-3 flex-wrap">
        <Field label="Filter by grant">
          <GrantPicker grants={grants} value={grantFilter === "All" ? "" : grantFilter} onChange={(v) => setGrantFilter(v || "All")} noneLabel="All grants" noneValue="All" wrapStyle={{ maxWidth: 320 }} />
        </Field>
        <Field label="Assigned to">
          <select value={assignedFilter} onChange={(e) => setAssignedFilter(e.target.value)} className={inputCls} style={{ ...inputStyle, maxWidth: 220 }}>
            <option>All</option>
            {assignees.map((a) => <option key={a}>{a}</option>)}
          </select>
        </Field>
      </div>

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
                  {bucket} <span style={{ color: "#8A8F87", fontWeight: 400 }}>({visible.filter((r) => effectiveReportBucket(r) === bucket).length})</span>
                </h3>
                <div className="space-y-3">
                  {visible.filter((r) => effectiveReportBucket(r) === bucket).sort((a, b) => new Date(a.dueDate || "9999-12-31") - new Date(b.dueDate || "9999-12-31")).map((r) => (
                    <ReportCard
                      key={r.id}
                      report={{ ...r, bucket: effectiveReportBucket(r) }}
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
  const [form, setForm, undoForm, canUndoForm] = useUndoableState(task || {
    id: uid(), title: "", category: TASK_CATEGORIES[0], grantId: "", dueDate: "",
    priority: "Medium", status: "Not started", assignedTo: "", notes: "",
  });
  const set = (k) => (e) => setForm({ ...form, [k]: e.target.value });

  return (
    <Modal title={task ? (canEdit ? "Edit task" : "View task") : "New task"} onClose={onClose}>
      <fieldset disabled={!canEdit} style={{ border: "none", margin: 0, padding: 0, minWidth: 0 }}>
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
          {canEdit && canUndoForm && (
            <button onClick={undoForm} className="px-3 py-2 rounded-md text-sm border inline-flex items-center gap-1.5" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>
              <Undo2 size={14} /> Undo
            </button>
          )}
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

function ReportingView({ grants, budgets, costCenters, budgetGroups, invoices }) {
  const [scope, setScope] = useState("all"); // all | a budget group id
  const [calYear, setCalYear] = useState("All");
  const [grantSort, setGrantSort] = useState({ key: "planExpense", dir: "desc" });
  const [exporting, setExporting] = useState(false);
  const barChartWrapRef = useRef(null);
  const lineChartWrapRef = useRef(null);

  const scopedGrantIds = useMemo(() => {
    if (scope === "all") return null;
    return new Set(grants.filter((g) => g.budgetGroupId === scope).map((g) => g.id));
  }, [scope, grants]);
  const scopedCostCenterIds = useMemo(() => {
    if (scope === "all") return null;
    return new Set((costCenters || []).filter((c) => c.budgetGroupId === scope).map((c) => c.id));
  }, [scope, costCenters]);

  const scopedBudgets = useMemo(() => (scope === "all"
    ? budgets
    : budgets.filter((b) => (b.grantId && scopedGrantIds.has(b.grantId)) || (b.costCenterId && scopedCostCenterIds.has(b.costCenterId)))
  ).filter((b) => isActiveBudget(b.status)), [scope, budgets, scopedGrantIds, scopedCostCenterIds]);

  const calendarYears = useMemo(() => {
    const years = new Set();
    scopedBudgets.forEach((b) => monthColumnsForBudget(b.periodStart, b.periodEnd).forEach((col) => years.add(col.year)));
    return [...years].sort();
  }, [scopedBudgets]);

  // One pass across every in-scope budget line, bucketing Plan and Actual figures
  // both by real calendar month (for the trend chart) and by grant/category (for
  // the tables below) — respecting the selected calendar year throughout.
  const agg = useMemo(() => {
    const monthly = Array.from({ length: 12 }, () => ({ planRevenue: 0, planExpense: 0, actualRevenue: 0, actualExpense: 0 }));
    const byCategory = {};
    const byGrant = {};
    const totals = { planRevenue: 0, planExpense: 0, actualRevenue: 0, actualExpense: 0 };

    scopedBudgets.forEach((b) => {
      const cols = monthColumnsForBudget(b.periodStart, b.periodEnd);
      const g = b.grantId ? grants.find((x) => x.id === b.grantId) : null;
      const cc = b.costCenterId ? costCenters.find((x) => x.id === b.costCenterId) : null;
      const ownerKey = b.grantId || b.costCenterId || "unknown";
      const ownerName = g ? (g.programCode ? `${g.programCode} - ${g.title}` : g.title) : cc ? cc.name : "Unknown";
      if (!byGrant[ownerKey]) byGrant[ownerKey] = { name: ownerName, planRevenue: 0, planExpense: 0, actualRevenue: 0, actualExpense: 0 };

      b.lines.forEach((l) => {
        (l.amounts || []).forEach((amt, i) => {
          const col = cols[i];
          if (!col || (calYear !== "All" && col.year !== calYear)) return;
          const v = Number(amt) || 0;
          if (l.type === "revenue") { monthly[col.monthIndex].planRevenue += v; totals.planRevenue += v; byGrant[ownerKey].planRevenue += v; }
          else { monthly[col.monthIndex].planExpense += v; totals.planExpense += v; byGrant[ownerKey].planExpense += v; }
          if (l.type === "expense") {
            if (!byCategory[l.category]) byCategory[l.category] = { category: l.category, plan: 0, actual: 0 };
            byCategory[l.category].plan += v;
          }
        });
        (l.actuals || []).forEach((amt, i) => {
          const col = cols[i];
          if (!col || (calYear !== "All" && col.year !== calYear)) return;
          const v = Number(amt) || 0;
          if (l.type === "revenue") { monthly[col.monthIndex].actualRevenue += v; totals.actualRevenue += v; byGrant[ownerKey].actualRevenue += v; }
          else { monthly[col.monthIndex].actualExpense += v; totals.actualExpense += v; byGrant[ownerKey].actualExpense += v; }
          if (l.type === "expense") {
            if (!byCategory[l.category]) byCategory[l.category] = { category: l.category, plan: 0, actual: 0 };
            byCategory[l.category].actual += v;
          }
        });
      });
    });

    return {
      monthly: MONTHS.map((m, i) => ({ month: m, ...monthly[i] })),
      byCategory: Object.values(byCategory).sort((a, b) => b.plan - a.plan),
      byGrant: Object.entries(byGrant).map(([id, v]) => ({ id, ...v })).filter((r) => r.planRevenue || r.planExpense || r.actualRevenue || r.actualExpense),
      totals,
    };
  }, [scopedBudgets, grants, costCenters, calYear]);

  const sortedByGrant = useMemo(() => {
    const dir = grantSort.dir === "asc" ? 1 : -1;
    return [...agg.byGrant].sort((a, b) => {
      const av = a[grantSort.key], bv = b[grantSort.key];
      if (typeof av === "string") return av.localeCompare(bv) * dir;
      return ((av ?? 0) - (bv ?? 0)) * dir;
    });
  }, [agg.byGrant, grantSort]);
  const toggleGrantSort = (key) => setGrantSort((s) => (s.key === key ? { key, dir: s.dir === "asc" ? "desc" : "asc" } : { key, dir: "desc" }));

  // Needs attention: grants behind pace per Burn Rate, and grants with overdue invoices.
  const behindPaceGrants = useMemo(() => grants
    .filter((g) => scope === "all" || scopedGrantIds?.has(g.id))
    .map((g) => ({ grant: g, burn: grantBurn(g, budgets) }))
    .filter((r) => r.burn.status === "Behind pace"), [grants, budgets, scope, scopedGrantIds]);

  const overdueByGrant = useMemo(() => {
    const map = {};
    (invoices || []).filter(isInvoiceOverdue).forEach((inv) => {
      if (scope !== "all" && !scopedGrantIds?.has(inv.grantId)) return;
      const g = grants.find((x) => x.id === inv.grantId);
      const key = inv.grantId || "unlinked";
      if (!map[key]) map[key] = { name: g ? (g.programCode ? `${g.programCode} - ${g.title}` : g.title) : "No grant linked", amount: 0, count: 0 };
      map[key].amount += Number(inv.amount) || 0;
      map[key].count += 1;
    });
    return Object.values(map);
  }, [invoices, grants, scope, scopedGrantIds]);

  const exportAllCsv = async () => {
    setExporting(true);
    try {
      const HEADER_FILL = "FFF6F7F3";
      const GREEN = "FF2F6F53";
      const RED = "FFB5443A";
      const round2 = (n) => Math.round((Number(n) || 0) * 100) / 100;

      // Capture the two live charts as images before building the workbook —
      // ExcelJS can embed a picture of a chart, but has no support for
      // creating a real, editable Excel chart object.
      const barSvg = barChartWrapRef.current?.querySelector("svg") || null;
      const lineSvg = lineChartWrapRef.current?.querySelector("svg") || null;
      const [barPng, linePng] = await Promise.all([
        svgElementToPngBase64(barSvg, 520, 280),
        svgElementToPngBase64(lineSvg, 520, 280),
      ]);

      const wb = new ExcelJS.Workbook();

      // ---- Sheet 1: Summary (charts + category table) ----
      const ws1 = wb.addWorksheet("Summary");
      ws1.mergeCells(1, 1, 1, 6);
      ws1.getCell(1, 1).value = "Nation's Finest — Reporting Summary";
      ws1.getCell(1, 1).font = { bold: true, size: 13 };
      ws1.mergeCells(2, 1, 2, 6);
      ws1.getCell(2, 1).value = `Scope: ${scope === "all" ? "Whole Organization" : (budgetGroups.find((g) => g.id === scope)?.name || "Scoped view")} — ${calYear === "All" ? "All years" : `Calendar year ${calYear}`} — Generated ${fmtDate(new Date().toISOString().slice(0, 10))}`;
      ws1.getCell(2, 1).font = { italic: true, size: 9, color: { argb: "FF8A8F87" } };

      let imgRow = 4;
      if (barPng) {
        const imgId = wb.addImage({ base64: barPng, extension: "png" });
        ws1.addImage(imgId, { tl: { col: 0, row: imgRow - 1 }, ext: { width: 520, height: 280 } });
      }
      if (linePng) {
        const imgId = wb.addImage({ base64: linePng, extension: "png" });
        ws1.addImage(imgId, { tl: { col: 7, row: imgRow - 1 }, ext: { width: 520, height: 280 } });
      }
      let r = imgRow + 16; // clear the embedded images before starting the table

      ws1.getRow(r).values = ["Category", "Plan", "Actual", "Variance"];
      ws1.getRow(r).eachCell((cell) => {
        cell.font = { bold: true };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: HEADER_FILL } };
      });
      r++;
      agg.byCategory.forEach((c) => {
        const variance = round2(c.actual - c.plan);
        ws1.getRow(r).values = [c.category, round2(c.plan), round2(c.actual), variance];
        ws1.getCell(r, 2).numFmt = "$#,##0";
        ws1.getCell(r, 3).numFmt = "$#,##0";
        ws1.getCell(r, 4).numFmt = "$#,##0";
        ws1.getCell(r, 4).font = { color: { argb: isNetNegative(variance) ? RED : GREEN } };
        r++;
      });
      ws1.columns = [{ width: 30 }, { width: 16 }, { width: 16 }, { width: 16 }, { width: 4 }, { width: 4 }, { width: 4 }, { width: 30 }];

      // ---- Sheet 2: By Grant / Cost Center ----
      const ws2 = wb.addWorksheet("By Grant");
      const header2 = ["Grant / Cost Center", "Plan revenue", "Plan expense", "Plan net", "Actual revenue", "Actual expense", "Actual net"];
      ws2.getRow(1).values = header2;
      ws2.getRow(1).eachCell((cell) => {
        cell.font = { bold: true };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: HEADER_FILL } };
      });
      sortedByGrant.forEach((g, i) => {
        const rowIdx = i + 2;
        const planNet = round2(g.planRevenue - g.planExpense);
        const actualNet = round2(g.actualRevenue - g.actualExpense);
        ws2.getRow(rowIdx).values = [g.name, round2(g.planRevenue), round2(g.planExpense), planNet, round2(g.actualRevenue), round2(g.actualExpense), actualNet];
        [2, 3, 4, 5, 6, 7].forEach((col) => { ws2.getCell(rowIdx, col).numFmt = "$#,##0"; });
        ws2.getCell(rowIdx, 4).font = { color: { argb: isNetNegative(planNet) ? RED : GREEN } };
        ws2.getCell(rowIdx, 7).font = { color: { argb: isNetNegative(actualNet) ? RED : GREEN } };
      });
      ws2.views = [{ state: "frozen", ySplit: 1 }];
      ws2.columns = [{ width: 34 }, { width: 15 }, { width: 15 }, { width: 15 }, { width: 15 }, { width: 15 }, { width: 15 }];

      // ---- Sheet 3: Budget Lines (full detail, same data as before) ----
      const ws3 = wb.addWorksheet("Budget Lines");
      const maxMonths = budgets.reduce((max, b) => Math.max(max, ...b.lines.map((l) => (l.amounts || []).length), 12), 12);
      const monthCols = Array.from({ length: maxMonths }, (_, i) => `Month ${i + 1}`);
      const header3 = ["Grant", "Budget", "Period Start", "Category", "Subcategory", "Description", "Type", ...monthCols, "Total"];
      ws3.getRow(1).values = header3;
      ws3.getRow(1).eachCell((cell) => {
        cell.font = { bold: true };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: HEADER_FILL } };
      });
      let r3 = 2;
      budgets.forEach((b) => {
        const g = grants.find((x) => x.id === b.grantId);
        b.lines.forEach((l) => {
          const padded = resizeMonthlyArray(l.amounts, maxMonths);
          ws3.getRow(r3).values = [g?.title || "", b.title, b.periodStart || "", l.category, l.subcategory, l.description || "", l.type, ...padded.map(round2), round2(lineTotal(l))];
          for (let c = 8; c <= 8 + maxMonths; c++) ws3.getCell(r3, c).numFmt = "$#,##0";
          r3++;
        });
      });
      ws3.views = [{ state: "frozen", ySplit: 1 }];
      ws3.columns = [{ width: 26 }, { width: 22 }, { width: 12 }, { width: 20 }, { width: 24 }, { width: 20 }, { width: 10 }, ...monthCols.map(() => ({ width: 11 })), { width: 13 }];

      const buffer = await wb.xlsx.writeBuffer();
      downloadFile("nations-finest-reporting.xlsx", buffer, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    } finally {
      setExporting(false);
    }
  };

  return (
    <div className="space-y-6">
      <div className="flex items-center justify-between">
        <h1 className="font-display text-2xl" style={{ color: "#1C2624" }}>Reporting</h1>
        <button onClick={exportAllCsv} disabled={exporting} className="inline-flex items-center gap-1.5 px-3 py-2 rounded-md text-sm border" style={{ borderColor: "#E1E5DE", color: "#1C2624", opacity: exporting ? 0.6 : 1 }}>
          <Download size={15} /> {exporting ? "Building export…" : "Export Excel"}
        </button>
      </div>

      <div className="flex flex-wrap gap-4">
        <Field label="Scope">
          <select value={scope} onChange={(e) => setScope(e.target.value)} className={inputCls} style={{ ...inputStyle, maxWidth: 260 }}>
            <option value="all">Whole Organization</option>
            {(budgetGroups || []).map((bg) => <option key={bg.id} value={bg.id}>{bg.name}</option>)}
          </select>
        </Field>
        <Field label="Calendar year">
          <select value={calYear} onChange={(e) => setCalYear(e.target.value === "All" ? "All" : Number(e.target.value))} className={inputCls} style={{ ...inputStyle, maxWidth: 220 }}>
            <option value="All">All years combined</option>
            {calendarYears.map((y) => <option key={y} value={y}>{y}</option>)}
          </select>
        </Field>
      </div>
      <p className="text-xs" style={{ color: "#8A8F87" }}>Only Active budgets are counted, matching Org Budget and Burn Rate.</p>

      <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
        <StatCard label="Total revenue" value={fmt(agg.totals.planRevenue)} sub={`Actual: ${fmt(agg.totals.actualRevenue)}`} />
        <StatCard label="Total expense" value={fmt(agg.totals.planExpense)} sub={`Actual: ${fmt(agg.totals.actualExpense)}`} />
        <StatCard label="Total net" value={fmt(agg.totals.planRevenue - agg.totals.planExpense)} sub={`Actual: ${fmt(agg.totals.actualRevenue - agg.totals.actualExpense)}`} />
      </div>

      {(behindPaceGrants.length > 0 || overdueByGrant.length > 0) && (
        <div className="bg-white rounded-lg border p-5" style={{ borderColor: "#C08A2E" }}>
          <h2 className="font-display text-base mb-3" style={{ color: "#C08A2E" }}>Needs attention</h2>
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-6">
            {behindPaceGrants.length > 0 && (
              <div>
                <div className="text-xs font-medium mb-2" style={{ color: "#8A8F87" }}>Behind pace (per Burn Rate)</div>
                <div className="space-y-1.5">
                  {behindPaceGrants.map(({ grant, burn }) => (
                    <div key={grant.id} className="flex items-center justify-between text-sm">
                      <span style={{ color: "#1C2624" }}>{grant.programCode ? `${grant.programCode} - ${grant.title}` : grant.title}</span>
                      <span style={{ color: "#B5443A" }}>{Math.round(burn.pctBudgetUsed * 100)}% used, {Math.round((burn.pctTimeElapsed || 0) * 100)}% time elapsed</span>
                    </div>
                  ))}
                </div>
              </div>
            )}
            {overdueByGrant.length > 0 && (
              <div>
                <div className="text-xs font-medium mb-2" style={{ color: "#8A8F87" }}>Overdue invoices</div>
                <div className="space-y-1.5">
                  {overdueByGrant.map((r) => (
                    <div key={r.name} className="flex items-center justify-between text-sm">
                      <span style={{ color: "#1C2624" }}>{r.name}</span>
                      <span style={{ color: "#B5443A" }}>{fmt(r.amount)} ({r.count} invoice{r.count === 1 ? "" : "s"})</span>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
        </div>
      )}

      <div className="grid grid-cols-1 sm:grid-cols-2 gap-6">
        <div className="bg-white rounded-lg border p-5" style={{ borderColor: "#E1E5DE" }} ref={barChartWrapRef}>
          <h2 className="font-display text-base mb-3" style={{ color: "#1C2624" }}>Expense by category — Plan vs. Actual</h2>
          {agg.byCategory.length === 0 ? <p className="text-sm" style={{ color: "#8A8F87" }}>No monthly data yet.</p> : (
            <ResponsiveContainer width="100%" height={280}>
              <BarChart data={agg.byCategory} layout="vertical" margin={{ left: 10 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#E1E5DE" />
                <XAxis type="number" tickFormatter={(v) => `$${v / 1000}k`} tick={{ fontSize: 11 }} />
                <YAxis type="category" dataKey="category" width={150} tick={{ fontSize: 10 }} />
                <Tooltip formatter={(v) => fmt(v)} />
                <Legend />
                <Bar dataKey="plan" name="Plan" fill="#5B7FA6" radius={[0, 3, 3, 0]} />
                <Bar dataKey="actual" name="Actual" fill="#B5443A" radius={[0, 3, 3, 0]} />
              </BarChart>
            </ResponsiveContainer>
          )}
        </div>

        <div className="bg-white rounded-lg border p-5" style={{ borderColor: "#E1E5DE" }} ref={lineChartWrapRef}>
          <h2 className="font-display text-base mb-3" style={{ color: "#1C2624" }}>Monthly trend — Plan vs. Actual</h2>
          <ResponsiveContainer width="100%" height={280}>
            <LineChart data={agg.monthly}>
              <CartesianGrid strokeDasharray="3 3" stroke="#E1E5DE" />
              <XAxis dataKey="month" tick={{ fontSize: 11 }} />
              <YAxis tickFormatter={(v) => `$${v / 1000}k`} tick={{ fontSize: 11 }} />
              <Tooltip formatter={(v) => fmt(v)} />
              <Legend />
              <Line type="monotone" dataKey="planRevenue" name="Plan revenue" stroke="#2F6F53" strokeWidth={2} dot={false} />
              <Line type="monotone" dataKey="planExpense" name="Plan expense" stroke="#B5443A" strokeWidth={2} dot={false} />
              <Line type="monotone" dataKey="actualRevenue" name="Actual revenue" stroke="#2F6F53" strokeWidth={2} strokeDasharray="4 3" dot={false} />
              <Line type="monotone" dataKey="actualExpense" name="Actual expense" stroke="#B5443A" strokeWidth={2} strokeDasharray="4 3" dot={false} />
            </LineChart>
          </ResponsiveContainer>
        </div>
      </div>

      <div className="bg-white rounded-lg border p-5" style={{ borderColor: "#E1E5DE" }}>
        <h2 className="font-display text-base mb-3" style={{ color: "#1C2624" }}>By grant / cost center</h2>
        {sortedByGrant.length === 0 ? (
          <p className="text-sm" style={{ color: "#8A8F87" }}>No budget data yet.</p>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead>
                <tr style={{ color: "#8A8F87" }}>
                  <th className="text-left py-1.5 font-medium cursor-pointer" onClick={() => toggleGrantSort("name")}>Grant / Cost Center</th>
                  <th className="text-right py-1.5 font-medium cursor-pointer" onClick={() => toggleGrantSort("planRevenue")}>Plan revenue</th>
                  <th className="text-right py-1.5 font-medium cursor-pointer" onClick={() => toggleGrantSort("actualRevenue")}>Actual revenue</th>
                  <th className="text-right py-1.5 font-medium cursor-pointer" onClick={() => toggleGrantSort("planExpense")}>Plan expense</th>
                  <th className="text-right py-1.5 font-medium cursor-pointer" onClick={() => toggleGrantSort("actualExpense")}>Actual expense</th>
                </tr>
              </thead>
              <tbody>
                {sortedByGrant.map((r) => (
                  <tr key={r.id} className="border-t" style={{ borderColor: "#E1E5DE" }}>
                    <td className="py-1.5" style={{ color: "#1C2624" }}>{r.name}</td>
                    <td className="py-1.5 text-right" style={{ fontVariantNumeric: "tabular-nums", color: "#2F6F53" }}>{fmt(r.planRevenue)}</td>
                    <td className="py-1.5 text-right" style={{ fontVariantNumeric: "tabular-nums", color: "#2F6F53" }}>{fmt(r.actualRevenue)}</td>
                    <td className="py-1.5 text-right" style={{ fontVariantNumeric: "tabular-nums" }}>{fmt(r.planExpense)}</td>
                    <td className="py-1.5 text-right" style={{ fontVariantNumeric: "tabular-nums" }}>{fmt(r.actualExpense)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}

function OrgBudgetRow({ label, values, projected, bold, indent, color, isHeader }) {
  const total = values.reduce((a, b) => a + b, 0);
  return (
    <tr className={isHeader ? "" : "border-t"} style={{ borderColor: "#E1E5DE", background: isHeader ? "#F6F7F3" : "transparent" }}>
      <td className={`px-3 py-1.5 text-xs sticky left-0 ${isHeader ? "" : "bg-white"}`} style={{ background: isHeader ? "#F6F7F3" : undefined, paddingLeft: indent ? 28 : 12, fontWeight: bold ? 600 : 400, color: color || "#1C2624" }}>
        {label}
      </td>
      {values.map((v, i) => {
        const isProj = !!(projected && projected[i]);
        return (
          <td
            key={i}
            className="px-2 py-1.5 text-right text-xs"
            style={{ fontVariantNumeric: "tabular-nums", fontWeight: bold ? 600 : 400, color: isProj ? "#8A8F87" : (color || "#1C2624"), fontStyle: isProj ? "italic" : "normal" }}
            title={isProj ? "Projected from this line's average actuals so far — not entered data" : undefined}
          >
            {v ? fmt(v) : "—"}
          </td>
        );
      })}
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
    const scoped = (orgScope === "all" ? budgets : budgets.filter((b) => (b.grantId && scopedGrantIds.has(b.grantId)) || (b.costCenterId && scopedCcIds.has(b.costCenterId)))).filter((b) => isActiveBudget(b.status));
    const years = new Set();
    scoped.forEach((b) => monthColumnsForBudget(b.periodStart, b.periodEnd).forEach((col) => years.add(col.year)));
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
      const scoped = (orgScope === "all" ? budgets : budgets.filter((b) => (b.grantId && scopedGrantIds.has(b.grantId)) || (b.costCenterId && scopedCcIds.has(b.costCenterId)))).filter((b) => isActiveBudget(b.status));
      const map = {};
      scoped.forEach((b) => {
        const cols = monthColumnsForBudget(b.periodStart, b.periodEnd);
        b.lines.forEach((l) => {
          const key = l.category;
          if (!map[key]) {
            const catDef = CATEGORIES.find((c) => c.name === l.category);
            map[key] = { category: l.category, type: catDef ? catDef.type : l.type, subcategory: "", amounts: Array(12).fill(0) };
          }
          (l.amounts || Array(12).fill(0)).forEach((a, i) => {
            const col = cols[i];
            if (!col) return;
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
  const [form, setForm, undoForm, canUndoForm] = useUndoableState(scenario);
  const [showCompare, setShowCompare] = useState(true);
  const cols = monthColumnsForBudget(form.periodStart, form.periodEnd);

  useEffect(() => {
    const needsResize = form.lines.some((l) => (l.amounts?.length || 0) !== cols.length);
    if (!needsResize) return;
    setForm((f) => ({
      ...f,
      lines: f.lines.map((l) => ({
        ...l,
        amounts: resizeMonthlyArray(l.amounts, cols.length),
      })),
    }));
  }, [cols.length]);

  const updateLine = (id, patch) => setForm((f) => ({ ...f, lines: f.lines.map((l) => (l.id === id ? { ...l, ...patch } : l)) }));
  const addLine = () => setForm((f) => ({ ...f, lines: [...f.lines, newLine(cols.length)] }));
  const removeLine = (id) => setForm((f) => ({ ...f, lines: f.lines.filter((l) => l.id !== id) }));
  const setAnnual = (id, val) => {
    const per = Math.round((Number(val) || 0) / cols.length * 100) / 100;
    updateLine(id, { amounts: Array(cols.length).fill(per) });
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
      ["Total Revenue", "", "", "", ...Array(cols.length).fill(""), totals.revenue],
      ["Total Expense", "", "", "", ...Array(cols.length).fill(""), totals.expense],
      ["Net", "", "", "", ...Array(cols.length).fill(""), totals.net],
    ];
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"] = [{ wch: 30 }, { wch: 18 }, { wch: 24 }, { wch: 10 }, ...cols.map(() => ({ wch: 12 })), { wch: 14 }];
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
          {canEdit && canUndoForm && (
            <button onClick={undoForm} className="inline-flex items-center gap-1.5 text-xs px-3 py-2 rounded-md border" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>
              <Undo2 size={14} /> Undo
            </button>
          )}
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
        <fieldset disabled={!canEdit} style={{ border: "none", margin: 0, padding: 0, minWidth: 0 }}>
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
          <table className="text-xs" style={{ fontFamily: "var(--mono-font)", width: "max-content" }}>
            <thead>
              <tr style={{ background: "#F6F7F3" }}>
                <th className="text-left px-2 py-2 sticky left-0" style={{ background: "#F6F7F3", minWidth: 190 }}>Category</th>
                <th className="text-left px-2 py-2" style={{ minWidth: 190 }}>Subcategory</th>
                <th className="text-left px-2 py-2" style={{ minWidth: 140 }}>Description</th>
                <th className="text-right px-2 py-2" style={{ minWidth: 95 }}>Annual total</th>
                {cols.map((c, i) => <th key={i} className="text-right px-2 py-2" style={{ minWidth: 88 }}>{c.label}</th>)}
                <th className="text-right px-2 py-2" style={{ minWidth: 80 }}>Total</th>
                <th className="px-2 py-2" style={{ minWidth: 36 }}></th>
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

  const deferredRevenueGroupId = useMemo(
    () => (budgetGroups || []).find((bg) => bg.name?.trim().toLowerCase() === "deferred revenue")?.id || null,
    [budgetGroups]
  );

  const scopedGrantIds = useMemo(() => {
    if (scope === "all") return null;
    return new Set(grants.filter((g) => g.budgetGroupId === scope).map((g) => g.id));
  }, [scope, grants]);
  const scopedCostCenterIds = useMemo(() => {
    if (scope === "all") return null;
    return new Set((costCenters || []).filter((c) => c.budgetGroupId === scope).map((c) => c.id));
  }, [scope, costCenters]);

  // When viewing "Whole Organization," anything tagged into the Deferred
  // Revenue budget group is deliberately left out of the overall totals —
  // it's still visible by selecting that group as its own scope.
  const excludedGrantIds = useMemo(() => {
    if (scope !== "all" || !deferredRevenueGroupId) return null;
    return new Set(grants.filter((g) => g.budgetGroupId === deferredRevenueGroupId).map((g) => g.id));
  }, [scope, deferredRevenueGroupId, grants]);
  const excludedCcIds = useMemo(() => {
    if (scope !== "all" || !deferredRevenueGroupId) return null;
    return new Set((costCenters || []).filter((c) => c.budgetGroupId === deferredRevenueGroupId).map((c) => c.id));
  }, [scope, deferredRevenueGroupId, costCenters]);

  const scopedBudgets = (scope === "all"
    ? budgets.filter((b) => !(b.grantId && excludedGrantIds?.has(b.grantId)) && !(b.costCenterId && excludedCcIds?.has(b.costCenterId)))
    : budgets.filter((b) => (b.grantId && scopedGrantIds.has(b.grantId)) || (b.costCenterId && scopedCostCenterIds.has(b.costCenterId)))
  ).filter((b) => isActiveBudget(b.status) && b.budgetType === "Template");

  // Real calendar years actually touched by any in-scope budget's period, so the
  // year picker reflects reality even when grants run off the calendar year.
  const calendarYears = useMemo(() => {
    const years = new Set();
    scopedBudgets.forEach((b) => monthColumnsForBudget(b.periodStart, b.periodEnd).forEach((col) => years.add(col.year)));
    return [...years].sort();
  }, [scopedBudgets]);

  const amountsField = dataMode === "plan" ? "amounts" : "actuals";
  const lineValue = (l) => (dataMode === "plan" ? lineTotal(l) : lineActualTotal(l));

  const revenueCats = CATEGORIES.filter((c) => c.type === "revenue");
  const expenseCats = CATEGORIES.filter((c) => c.type === "expense");
  const balanceCats = CATEGORIES.filter((c) => c.type === "balance");

  const grouped = useMemo(() => {
    const map = {};
    CATEGORIES.forEach((c) => { map[c.name] = { type: c.type, monthly: Array(12).fill(0), monthlyProjected: Array(12).fill(false), subs: {} }; });
    const addToBucket = (category, subcategory, type, monthIndex, value, isProjected) => {
      if (!map[category]) map[category] = { type, monthly: Array(12).fill(0), monthlyProjected: Array(12).fill(false), subs: {} };
      const bucket = map[category];
      bucket.monthly[monthIndex] += value;
      if (isProjected) bucket.monthlyProjected[monthIndex] = true;
      if (subcategory) {
        if (!bucket.subs[subcategory]) bucket.subs[subcategory] = { values: Array(12).fill(0), projected: Array(12).fill(false) };
        bucket.subs[subcategory].values[monthIndex] += value;
        if (isProjected) bucket.subs[subcategory].projected[monthIndex] = true;
      }
    };

    if (dataMode !== "actual") {
      // Plan mode: no projection concept applies — figures are exactly what
      // was entered, per budget, same as before this feature existed.
      scopedBudgets.forEach((b) => {
        const cols = monthColumnsForBudget(b.periodStart, b.periodEnd);
        b.lines.forEach((l) => {
          const vals = l.amounts || Array(12).fill(0);
          cols.forEach((col, i) => {
            if (!col) return;
            if (calYear !== "All" && col.year !== calYear) return;
            addToBucket(l.category, l.subcategory, l.type, col.monthIndex, Number(vals[i]) || 0, false);
          });
        });
      });
      return map;
    }

    // Actual mode: chain each grant/cost center's budgets together so a
    // projection can carry forward across a fiscal-year boundary into the
    // next budget, instead of stopping dead at the marked budget's own
    // periodEnd. A newer budget's own marked cutoff always takes priority
    // over an inherited average the moment it's reached.
    const ownerKey = (b) => b.grantId || b.costCenterId || `__standalone_${b.id}`;
    const byOwner = {};
    scopedBudgets.forEach((b) => { (byOwner[ownerKey(b)] = byOwner[ownerKey(b)] || []).push(b); });

    Object.values(byOwner).forEach((ownerBudgets) => {
      ownerBudgets.sort((a, b) => new Date(a.periodStart) - new Date(b.periodStart));

      // Every distinct category+subcategory line seen anywhere for this owner.
      const lineMeta = {};
      ownerBudgets.forEach((b) => b.lines.forEach((l) => {
        const key = `${l.category}|||${l.subcategory || ""}`;
        if (!lineMeta[key]) lineMeta[key] = { category: l.category, subcategory: l.subcategory || "", type: l.type };
      }));

      // Each budget's own run-rate average per line — summed across EVERY
      // line that shares that category+subcategory. A single budget can have
      // many individually-entered lines under the same category (e.g. one
      // payroll line per staff position, all filed under "Wages and
      // Benefits"), so matching only the first line found would silently
      // drop the rest.
      const budgetLineAvg = {}; // `${budgetId}|||${lineKey}` -> average
      const budgetLineActuals = {}; // `${budgetId}|||${lineKey}` -> combined actuals array
      ownerBudgets.forEach((b) => {
        const cols = monthColumnsForBudget(b.periodStart, b.periodEnd);
        Object.keys(lineMeta).forEach((key) => {
          const matching = b.lines.filter((l) => `${l.category}|||${l.subcategory || ""}` === key);
          if (matching.length === 0) return;
          const combined = Array(cols.length).fill(0);
          matching.forEach((l) => {
            (l.actuals || []).forEach((v, i) => { if (i < combined.length) combined[i] += Number(v) || 0; });
          });
          budgetLineActuals[`${b.id}|||${key}`] = combined;
          const cutoff = parseActualsThrough(b.actualsThrough);
          if (!cutoff) return;
          const vals = cols
            .map((col, i) => (colIsWithinCutoff(col, cutoff) ? combined[i] : null))
            .filter((v) => v !== null);
          budgetLineAvg[`${b.id}|||${key}`] = vals.length ? vals.reduce((a, x) => a + x, 0) / vals.length : 0;
        });
      });

      Object.keys(lineMeta).forEach((key) => {
        const meta = lineMeta[key];
        // Flatten this line's data across every owner budget into one
        // chronological timeline of calendar months.
        const timeline = [];
        ownerBudgets.forEach((b) => {
          const cols = monthColumnsForBudget(b.periodStart, b.periodEnd);
          const combined = budgetLineActuals[`${b.id}|||${key}`]; // undefined if this budget has no matching lines
          const cutoff = parseActualsThrough(b.actualsThrough);
          cols.forEach((col, i) => {
            timeline.push({
              year: col.year,
              monthIndex: col.monthIndex,
              isRealEntry: !!(combined && cutoff && colIsWithinCutoff(col, cutoff)),
              rawValue: combined ? (Number(combined[i]) || 0) : 0,
              budgetId: b.id,
              cutoff,
              atOrAfterCutoff: cutoff ? !colIsWithinCutoff(col, cutoff) : false,
            });
          });
        });
        timeline.sort((a, b) => (a.year - b.year) || (a.monthIndex - b.monthIndex));

        let basis = null; // most recently established run-rate average for this line
        timeline.forEach((t) => {
          const avgKey = `${t.budgetId}|||${key}`;
          if (t.atOrAfterCutoff && budgetLineAvg[avgKey] !== undefined) basis = budgetLineAvg[avgKey];

          let value, isProjected;
          if (t.isRealEntry) {
            value = t.rawValue;
            isProjected = false;
          } else if (basis !== null) {
            value = basis;
            isProjected = true;
          } else {
            value = t.rawValue; // no basis established anywhere yet — behaves as before this feature existed
            isProjected = false;
          }

          if (calYear !== "All" && t.year !== calYear) return;
          addToBucket(meta.category, meta.subcategory, meta.type, t.monthIndex, value, isProjected);
        });
      });
    });

    return map;
  }, [scopedBudgets, amountsField, calYear, dataMode]);

  // For Total Revenue / Total Expense / Net rows, a month is "projected" if
  // any contributing category had projected data in it.
  const combineProjected = (cats) => {
    const arr = Array(12).fill(false);
    cats.forEach((c) => grouped[c.name].monthlyProjected.forEach((p, i) => { if (p) arr[i] = true; }));
    return arr;
  };

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
  const totalRevenueProjected = combineProjected(revenueCats);
  const totalExpenseProjected = combineProjected(expenseCats);
  const netProjected = totalRevenueProjected.map((p, i) => p || totalExpenseProjected[i]);
  const anyProjected = dataMode === "actual" && scopedBudgets.some((b) => parseActualsThrough(b.actualsThrough));

  const exportCsv = async () => {
    const monthLabels = MONTHS.map((m) => (calYear === "All" ? m : `${m} ${calYear}`));
    const HEADER_FILL = "FFF6F7F3";
    const GREEN = "FF2F6F53";
    const RED = "FFB5443A";
    const round2 = (n) => Math.round((Number(n) || 0) * 100) / 100;

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Org Budget");

    const titleParts = [
      "Nation's Finest — Organizational Budget",
      scope === "all" ? "Whole Organization" : (budgetGroups.find((g) => g.id === scope)?.name || "Scoped view"),
      dataMode === "plan" ? "Plan" : "Actual",
      calYear === "All" ? "All years" : `Calendar year ${calYear}`,
    ];
    ws.mergeCells(1, 1, 1, monthLabels.length + 3);
    const titleCell = ws.getCell(1, 1);
    titleCell.value = titleParts.join(" — ");
    titleCell.font = { bold: true, size: 13 };
    ws.getRow(1).height = 22;

    ws.mergeCells(2, 1, 2, monthLabels.length + 3);
    const subCell = ws.getCell(2, 1);
    subCell.value = `Generated ${fmtDate(new Date().toISOString().slice(0, 10))}${anyProjected ? " — italicized figures in the app are run-rate projections; this export shows the same blended totals as numbers" : ""}`;
    subCell.font = { italic: true, size: 9, color: { argb: "FF8A8F87" } };

    const headerRowIdx = 4;
    const header = ["Category", "Subcategory", ...monthLabels, "Total"];
    ws.getRow(headerRowIdx).values = header;
    ws.getRow(headerRowIdx).eachCell((cell) => {
      cell.font = { bold: true };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: HEADER_FILL } };
    });

    let r = headerRowIdx + 1;
    const writeRow = (category, subcategory, values, opts = {}) => {
      const total = values.reduce((a, v) => a + v, 0);
      const rowVals = [category, subcategory, ...values.map(round2), round2(total)];
      ws.getRow(r).values = rowVals;
      ws.getRow(r).eachCell((cell, colNumber) => {
        if (colNumber >= 3) cell.numFmt = "$#,##0";
        if (opts.bold) cell.font = { ...(cell.font || {}), bold: true };
        if (opts.fill) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: opts.fill } };
        if (opts.color) cell.font = { ...(cell.font || {}), bold: true, color: { argb: opts.color } };
      });
      r++;
    };

    const pushSection = (cats) => cats.forEach((c) => {
      const bucket = grouped[c.name];
      writeRow(c.name, "", bucket.monthly, { bold: true });
      Object.entries(bucket.subs).forEach(([sub, subData]) => {
        writeRow("", sub, subData.values);
      });
    });

    pushSection(revenueCats);
    writeRow("Total Revenue", "", totalRevenue, { bold: true, fill: HEADER_FILL, color: GREEN });
    pushSection(expenseCats);
    writeRow("Total Expense", "", totalExpense, { bold: true, fill: HEADER_FILL });
    const netTotal = round2(net.reduce((a, b) => a + b, 0));
    writeRow("Net", "", net, { bold: true, fill: HEADER_FILL, color: isNetNegative(netTotal) ? RED : GREEN });

    if (balanceCats.length > 0) {
      r++; // blank separator row
      writeRow("Balance Sheet (not included in Net)", "", Array(monthLabels.length).fill(0), { bold: true, fill: HEADER_FILL });
      pushSection(balanceCats);
    }

    ws.columns = [{ width: 26 }, { width: 30 }, ...monthLabels.map(() => ({ width: 13 })), { width: 15 }];
    ws.views = [{ state: "frozen", xSplit: 2, ySplit: headerRowIdx }];

    const buffer = await wb.xlsx.writeBuffer();
    downloadFile("nations-finest-organizational-budget.xlsx", buffer, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
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
            <Download size={15} /> Export Excel
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
      <p className="text-xs" style={{ color: "#8A8F87" }}>
        Only budgets marked as budget type "Template" are included here — the "Operational" version sent to a grantor doesn't affect this rollup.
      </p>
      {scope === "all" && deferredRevenueGroupId && (
        <p className="text-xs" style={{ color: "#8A8F87" }}>
          Grants and cost centers in the "Deferred Revenue" budget group are excluded from Whole Organization totals — select that group above to view them.
        </p>
      )}
      {viewMode === "monthly" && anyProjected && (
        <p className="text-xs flex items-center gap-1.5" style={{ color: "#8A8F87" }}>
          <span style={{ fontStyle: "italic" }}>Italic, muted figures</span> are projected from a grant's average actuals so far — carried forward across fiscal-year boundaries until real or newly-marked data takes over. Set "Actuals complete through" on a budget to turn this on.
        </p>
      )}

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
                    <td key={fy} className="px-3 py-1.5 text-xs text-right font-semibold" style={{ fontVariantNumeric: "tabular-nums", color: !isNetNegative(yearCompare.byYear[fy].net) ? "#2F6F53" : "#B5443A" }}>
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
                  <OrgBudgetRow label={c.name} values={grouped[c.name].monthly} projected={grouped[c.name].monthlyProjected} indent color="#2F6F53" />
                  {Object.entries(grouped[c.name].subs).map(([sub, subData]) => (
                    <OrgBudgetRow key={c.name + sub} label={sub} values={subData.values} projected={subData.projected} indent color="#5B6B66" />
                  ))}
                </Fragment>
              ))}
              <OrgBudgetRow label="Total Revenue" values={totalRevenue} projected={totalRevenueProjected} bold color="#2F6F53" />

              <OrgBudgetRow label="Expense" values={Array(12).fill(0)} isHeader bold />
              {expenseCats.map((c) => (
                <Fragment key={c.name}>
                  <OrgBudgetRow label={c.name} values={grouped[c.name].monthly} projected={grouped[c.name].monthlyProjected} indent />
                  {Object.entries(grouped[c.name].subs).map(([sub, subData]) => (
                    <OrgBudgetRow key={c.name + sub} label={sub} values={subData.values} projected={subData.projected} indent color="#5B6B66" />
                  ))}
                </Fragment>
              ))}
              <OrgBudgetRow label="Total Expense" values={totalExpense} projected={totalExpenseProjected} bold />
              <OrgBudgetRow label="Net Total" values={net} projected={netProjected} bold color={!isNetNegative(net.reduce((a, b) => a + b, 0)) ? "#2F6F53" : "#B5443A"} />
              {balanceCats.length > 0 && (
                <>
                  <tr><td colSpan={MONTHS.length + 2} className="py-2"></td></tr>
                  <OrgBudgetRow label="Balance Sheet (not included in Net)" values={Array(12).fill(0)} isHeader bold />
                  {balanceCats.map((c) => (
                    <Fragment key={c.name}>
                      <OrgBudgetRow label={c.name} values={grouped[c.name].monthly} projected={grouped[c.name].monthlyProjected} indent color="#5B7FA6" />
                      {Object.entries(grouped[c.name].subs).map(([sub, subData]) => (
                        <OrgBudgetRow key={c.name + sub} label={sub} values={subData.values} projected={subData.projected} indent color="#5B6B66" />
                      ))}
                    </Fragment>
                  ))}
                </>
              )}
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
  const defaultStaff = {
    id: uid(), name: "", position: "", department: "", exempt: "Non-exempt",
    payType: "Salary", annualSalary: 0, hourlyRate: 0, annualHours: ANNUAL_HOURS,
    fte: 1, allocations: [], site: "", status: "Active",
    bonus: 0, benefits: 0, payrollTaxRate: 0, raiseDate: "", raisePercent: 0, paylocityId: "",
  };
  const [form, setForm, undoForm, canUndoForm] = useUndoableState(
    staff ? { ...defaultStaff, ...staff, status: staff.status || "Active" } : defaultStaff
  );
  const set = (k) => (e) => setForm({ ...form, [k]: e.target.value });

  const addAlloc = () => setForm({ ...form, allocations: [...form.allocations, newAllocation()] });
  const updateAlloc = (id, patch) => setForm({ ...form, allocations: form.allocations.map((a) => (a.id === id ? { ...a, ...patch } : a)) });
  const removeAlloc = (id) => setForm({ ...form, allocations: form.allocations.filter((a) => a.id !== id) });

  const cost = staffAnnualCost(form);
  const loaded = staffFullyLoadedCost(form);
  const allocatedPct = staffAllocatedTotal(form);

  return (
    <Modal title={staff ? (canEdit ? "Edit staff member" : "View staff member") : "New staff member"} onClose={onClose} wide>
      <fieldset disabled={!canEdit} style={{ border: "none", margin: 0, padding: 0, minWidth: 0 }}>
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
        <Field label="Name">
          <input className={inputCls} style={inputStyle} value={form.name} onChange={set("name")} placeholder="Last, First" />
        </Field>
        <Field label="Paylocity ID">
          <input className={inputCls} style={inputStyle} value={form.paylocityId} onChange={set("paylocityId")} placeholder="e.g. BARTK" />
        </Field>
      </div>
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 mt-4">
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

      <div className="mt-5">
        <div className="flex items-center justify-between mb-2">
          <h3 className="font-display text-sm" style={{ color: "#1C2624" }}>Compensation details</h3>
          {form.lastPaylocitySync && (
            <span className="text-xs" style={{ color: "#2F6F53" }}>
              Synced from Paylocity — {fmtDate(form.lastPaylocitySync.periodStart)}–{fmtDate(form.lastPaylocitySync.periodEnd)}
            </span>
          )}
        </div>
        <p className="text-xs mb-2" style={{ color: "#8A8F87" }}>
          Used to generate budget lines directly from Personnel — separate from base pay above.
        </p>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          <Field label="Annual bonus">
            <input type="number" className={inputCls} style={inputStyle} value={form.bonus} onChange={set("bonus")} />
          </Field>
          <Field label="Annual benefits ($)">
            <input type="number" className={inputCls} style={inputStyle} value={form.benefits} onChange={set("benefits")} placeholder="Health, dental, retirement, etc." />
          </Field>
          <Field label="Payroll tax rate (%)">
            <input type="number" step="0.01" className={inputCls} style={inputStyle} value={form.payrollTaxRate} onChange={set("payrollTaxRate")} placeholder="e.g. 9.5" />
          </Field>
          <Field label="Next raise/anniversary date">
            <input type="date" className={inputCls} style={inputStyle} value={form.raiseDate || ""} onChange={set("raiseDate")} />
          </Field>
          <Field label="Expected raise (%)">
            <input type="number" step="0.01" className={inputCls} style={inputStyle} value={form.raisePercent} onChange={set("raisePercent")} placeholder="e.g. 3" />
          </Field>
        </div>
        <p className="text-xs mt-2" style={{ color: "#8A8F87" }}>
          Raise date/% are captured for future use but don't yet affect generated budget lines — those still generate a flat annual figure.
        </p>
      </div>

      <div className="mt-3 text-sm space-y-0.5">
        <div><span style={{ color: "#8A8F87" }}>Base annual cost: </span><span className="font-medium" style={{ color: "#1C2624" }}>{fmt(cost)}</span></div>
        <div><span style={{ color: "#8A8F87" }}>Fully loaded annual cost: </span><span className="font-medium" style={{ color: "#1C2624" }}>{fmt(loaded.total)}</span></div>
        <div className="text-xs" style={{ color: "#8A8F87" }}>
          Wages: {fmt(loaded.wagesTotal)} · Taxes &amp; benefits: {fmt(loaded.taxAndBenefitsTotal)}
        </div>
      </div>

      <div className="mt-5">
        <div className="flex items-center justify-between mb-2">
          <h3 className="text-sm font-medium" style={{ color: "#1C2624" }}>Allocations</h3>
          <span className="text-xs" style={{ color: allocatedPct > 100 ? "#B5443A" : "#8A8F87" }}>
            {allocatedPct}% allocated{form.lastPaylocitySync ? <> · <span style={{ color: "#2F6F53" }}>synced {fmtDate(form.lastPaylocitySync.periodStart)}–{fmtDate(form.lastPaylocitySync.periodEnd)}</span></> : ""}
          </span>
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
          {canEdit && canUndoForm && (
            <button onClick={undoForm} className="px-3 py-2 rounded-md text-sm border inline-flex items-center gap-1.5" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>
              <Undo2 size={14} /> Undo
            </button>
          )}
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

function PersonnelView({ grants, staff, setStaff, costCenters, setTrash, currentUserEmail, canEdit, initialOpenStaffId, logActivity, budgets = [], setBudgets, paylocityProgramMap = [], setPaylocityProgramMap, paylocityLastImport, setPaylocityLastImport }) {
  const [modal, setModal] = useState(() => (initialOpenStaffId ? staff.find((s) => s.id === stripNonce(initialOpenStaffId)) || null : null));
  const [confirm, setConfirm] = useState(null);
  const [deptFilter, setDeptFilter] = useState("All");
  const [siteFilter, setSiteFilter] = useState("All");
  const [statusFilter, setStatusFilter] = useState("All");
  const [sortBy, setSortBy] = useState("name");
  const [showImport, setShowImport] = useState(false);

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
  const totalWages = activeStaff.reduce((a, s) => a + staffFullyLoadedCost(s).wagesTotal, 0);
  const totalTaxesAndBenefits = activeStaff.reduce((a, s) => a + staffFullyLoadedCost(s).taxAndBenefitsTotal, 0);
  const totalPersonnelCost = totalWages + totalTaxesAndBenefits;

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
          {paylocityLastImport && (
            <p className="text-xs mt-1" style={{ color: "#8A8F87" }}>
              Last synced from Paylocity: covering {fmtDate(paylocityLastImport.periodStart)}–{fmtDate(paylocityLastImport.periodEnd)}, imported {fmtDate(paylocityLastImport.importedAt)}
            </p>
          )}
        </div>
        {canEdit && (
          <div className="flex items-center gap-2">
            <button onClick={() => setShowImport(true)} className="inline-flex items-center gap-1.5 px-3 py-2 rounded-md text-sm border" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>
              <Upload size={16} /> Import from Paylocity
            </button>
            <button onClick={() => setModal("new")} className="inline-flex items-center gap-1.5 px-3 py-2 rounded-md text-sm text-white" style={{ background: "#1F5C6B" }}>
              <Plus size={16} /> New staff member
            </button>
          </div>
        )}
      </div>

      <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
        <StatCard label="Total annual personnel cost" value={fmt(totalPersonnelCost)} />
        <StatCard label="Total annual wages" value={fmt(totalWages)} />
        <StatCard label="Total annual taxes & benefits" value={fmt(totalTaxesAndBenefits)} />
        <StatCard label="Total staff" value={activeStaff.length} />
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
      {showImport && (
        <PaylocityImportModal
          staff={staff}
          setStaff={setStaff}
          grants={grants}
          costCenters={costCenters}
          budgets={budgets}
          setBudgets={setBudgets}
          paylocityProgramMap={paylocityProgramMap}
          setPaylocityProgramMap={setPaylocityProgramMap}
          paylocityLastImport={paylocityLastImport}
          setPaylocityLastImport={setPaylocityLastImport}
          logActivity={logActivity}
          onClose={() => setShowImport(false)}
          onOpenStaff={(s) => { setShowImport(false); setModal(s); }}
        />
      )}
    </div>
  );
}

function PaylocityImportModal({ staff, setStaff, grants, costCenters, budgets, setBudgets, paylocityProgramMap, setPaylocityProgramMap, paylocityLastImport, setPaylocityLastImport, logActivity, onClose, onOpenStaff }) {
  const [step, setStep] = useState("upload"); // upload -> link -> crosswalk -> review -> done
  const [periodStart, setPeriodStart] = useState(paylocityLastImport ? "" : "2026-01-01");
  const [periodEnd, setPeriodEnd] = useState("");
  const [error, setError] = useState("");
  const [rows, setRows] = useState(null);
  const [linkDraft, setLinkDraft] = useState({}); // paylocityId -> staffId | "__new__"
  const [crosswalkDraft, setCrosswalkDraft] = useState({}); // code -> { grantId, costCenterId, ignore }
  const [updateComp, setUpdateComp] = useState(true);
  const [updateAllocations, setUpdateAllocations] = useState(true);
  const [result, setResult] = useState(null);

  const parseFile = async (file) => {
    setError("");
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
      const headerRow = aoa[1];
      if (!headerRow || !headerRow.includes("ID") || !headerRow.includes("Worked Program")) {
        setError("This doesn't look like a Paylocity Labor Distribution Percentages export grouped by Worked Program. Double-check the file.");
        return;
      }
      const parsed = aoa.slice(2)
        .filter((r) => r && r[headerRow.indexOf("ID")])
        .map((r) => {
          const obj = {};
          headerRow.forEach((h, i) => { if (h) obj[h] = r[i]; });
          return obj;
        });
      setRows(parsed);
    } catch (e) {
      setError("Couldn't read that file — make sure it's the .xlsx Paylocity export.");
    }
  };

  const unmatchedEmployees = useMemo(() => {
    if (!rows) return [];
    const knownIds = new Set(staff.map((s) => s.paylocityId).filter(Boolean));
    const byId = {};
    rows.forEach((r) => {
      const id = r["ID"];
      if (id && !knownIds.has(id) && !byId[id]) byId[id] = r["Employee Name"] || id;
    });
    return Object.entries(byId).map(([paylocityId, name]) => ({ paylocityId, name }));
  }, [rows, staff]);

  const suggestMatch = (paylocityName) => {
    const key = paylocityNameKey(paylocityName);
    return staff.find((s) => !s.paylocityId && paylocityNameKey(s.name) === key) || null;
  };

  const distinctCodes = useMemo(() => {
    if (!rows) return [];
    const seen = new Map();
    rows.forEach((r) => {
      const label = r["Worked Program"];
      if (!label) return;
      const code = String(label).split(":")[0].trim();
      if (!seen.has(code)) seen.set(code, label);
    });
    return [...seen.entries()].map(([code, label]) => ({ code, label }));
  }, [rows]);

  const unmappedCodes = useMemo(() => {
    const mappedCodes = new Set(paylocityProgramMap.map((m) => m.code));
    return distinctCodes.filter((c) => !mappedCodes.has(c.code));
  }, [distinctCodes, paylocityProgramMap]);

  const [wideRangeAck, setWideRangeAck] = useState(false);

  const goToNextStep = () => {
    if (!periodStart || !periodEnd) { setError("Enter the date range this report covers."); return; }
    if (!rows || rows.length === 0) { setError("Upload a file first."); return; }
    if (!updateComp && !updateAllocations) { setError("Check at least one of Compensation or Allocations to import."); return; }
    const days = daysBetweenInclusive(periodStart, periodEnd);
    if (updateAllocations && days > 45 && !wideRangeAck) {
      setError(`This is a ${days}-day range for updating allocations — that's a lot wider than a typical single pay period, and will blend everyone's percentages toward this period's average instead of reflecting who's currently working where. If that's intentional, click Continue again to proceed.`);
      setWideRangeAck(true);
      return;
    }
    setError("");
    setWideRangeAck(false);
    if (unmatchedEmployees.length > 0) {
      setStep("link");
    } else {
      goToCrosswalkOrReview();
    }
  };

  const confirmLinks = () => {
    // Anything left un-chosen defaults to whatever's showing in the
    // dropdown (the suggested match, or "new employee") — finalize those
    // defaults into state now so downstream matching can rely on them.
    const finalized = { ...linkDraft };
    unmatchedEmployees.forEach((e) => {
      if (finalized[e.paylocityId] === undefined) {
        const suggested = suggestMatch(e.name);
        finalized[e.paylocityId] = suggested ? suggested.id : "__new__";
      }
    });
    setLinkDraft(finalized);
    goToCrosswalkOrReview();
  };

  // Existing staff records, with any confirmed name-based links applied
  // in-memory — this is what matching uses from here on, so a person linked
  // this way gets fully treated as matched for the rest of this import.
  const getWorkingStaff = () => staff.map((s) => {
    const linkedPaylocityId = Object.entries(linkDraft).find(([, sid]) => sid === s.id)?.[0];
    return linkedPaylocityId && !s.paylocityId ? { ...s, paylocityId: linkedPaylocityId } : s;
  });

  const goToCrosswalkOrReview = () => {
    if (unmappedCodes.length > 0) {
      const draft = {};
      unmappedCodes.forEach((c) => { draft[c.code] = { grantId: "", costCenterId: "", ignore: false }; });
      setCrosswalkDraft(draft);
      setStep("crosswalk");
    } else {
      buildReview();
    }
  };

  const confirmCrosswalk = () => {
    const unresolved = unmappedCodes.filter((c) => {
      const d = crosswalkDraft[c.code];
      return !d || (!d.ignore && !d.grantId && !d.costCenterId);
    });
    if (unresolved.length > 0) { setError("Map or ignore every listed program code before continuing."); return; }
    setError("");
    const additions = unmappedCodes.map((c) => ({ code: c.code, label: c.label, ...crosswalkDraft[c.code] }));
    setPaylocityProgramMap((prev) => [...prev, ...additions]);
    buildReview(additions);
  };

  const buildReview = (freshMappings = []) => {
    const fullMap = [...paylocityProgramMap, ...freshMappings];
    const codeToTarget = {};
    fullMap.forEach((m) => { codeToTarget[m.code] = m; });

    const workingStaff = getWorkingStaff();

    const byEmployee = {};
    rows.forEach((r) => {
      const id = r["ID"];
      if (!id) return;
      (byEmployee[id] = byEmployee[id] || []).push(r);
    });

    const staffByPaylocityId = {};
    workingStaff.forEach((s) => { if (s.paylocityId) staffByPaylocityId[s.paylocityId] = s; });

    const totalDays = daysBetweenInclusive(periodStart, periodEnd);
    const factor = totalDays > 0 ? 365 / totalDays : 0;

    const matchedUpdates = []; // { staffId, patch }
    const newEmployees = []; // { paylocityId, name }
    const grantMonthlyTotals = {}; // grantId -> { wage: {year-month: $}, benefitAndTax: {year-month: $} }

    Object.entries(byEmployee).forEach(([paylocityId, employeeRows]) => {
      const existingStaff = staffByPaylocityId[paylocityId];
      const totalWage = employeeRows.reduce((a, r) => a + sumWageCols(r), 0);
      const totalBenefit = employeeRows.reduce((a, r) => a + sumCols(r, PAYLOCITY_BENEFIT_COLS), 0);
      const totalEmployerTax = employeeRows.reduce((a, r) => a + sumCols(r, PAYLOCITY_EMPLOYER_TAX_COLS) - sumCols(r, PAYLOCITY_FFCRA_CREDIT_COLS), 0);

      // Every program row for this employee must resolve to a grant/cost
      // center (or an explicit "ignore") before we touch their record —
      // a partial picture is worse than flagging it for next time.
      const targets = employeeRows.map((r) => {
        const label = r["Worked Program"];
        const code = label ? String(label).split(":")[0].trim() : "";
        return { row: r, target: codeToTarget[code] };
      });

      if (!existingStaff) {
        newEmployees.push({ paylocityId, name: employeeRows[0]["Employee Name"] || paylocityId });
        return;
      }

      const fte = Number(existingStaff.fte) || 1;
      const annualWage = totalWage * factor;
      const annualBenefit = totalBenefit * factor;
      const annualEmployerTax = totalEmployerTax * factor;
      const payrollTaxRate = annualWage > 0 ? (annualEmployerTax / annualWage) * 100 : 0;

      const patch = { paylocityId };
      if (updateComp) {
        patch.benefits = fte > 0 ? annualBenefit / fte : annualBenefit;
        patch.payrollTaxRate = payrollTaxRate;
        patch.lastPaylocitySync = { periodStart, periodEnd, importedAt: new Date().toISOString().slice(0, 10) };
        if (existingStaff.payType === "Hourly") {
          const hours = Number(existingStaff.annualHours) || ANNUAL_HOURS;
          patch.hourlyRate = fte > 0 && hours > 0 ? annualWage / (hours * fte) : 0;
        } else {
          patch.annualSalary = fte > 0 ? annualWage / fte : annualWage;
        }
      }

      // Allocations replace entirely, from this import's Labor % breakdown —
      // rows with an "ignore" target simply don't contribute an allocation.
      if (updateAllocations) {
        const newAllocations = [];
        targets.forEach(({ row, target }) => {
          if (!target || target.ignore) return;
          const pct = (Number(row["Labor %"]) || 0) * 100;
          if (target.grantId) newAllocations.push({ id: uid(), type: "grant", grantId: target.grantId, costCenterId: "", percent: Math.round(pct * 100) / 100 });
          else if (target.costCenterId) newAllocations.push({ id: uid(), type: "costCenter", grantId: "", costCenterId: target.costCenterId, percent: Math.round(pct * 100) / 100 });
        });
        patch.allocations = newAllocations;
      }

      matchedUpdates.push({ staffId: existingStaff.id, patch });

      if (!updateComp) return; // allocations-only run — never touch budget actuals

      // Real dollars for this specific period, split by calendar month, into
      // whichever grant each program row maps to — this is what feeds the
      // Template budget's Actuals.
      targets.forEach(({ row, target }) => {
        if (!target || target.ignore || !target.grantId) return; // budget actuals only make sense for grant-linked budgets
        const rowWage = sumWageCols(row);
        const rowBenefitAndTax = sumCols(row, PAYLOCITY_BENEFIT_COLS) + sumCols(row, PAYLOCITY_EMPLOYER_TAX_COLS) - sumCols(row, PAYLOCITY_FFCRA_CREDIT_COLS);
        if (!grantMonthlyTotals[target.grantId]) grantMonthlyTotals[target.grantId] = { wage: {}, benefitAndTax: {} };
        splitAmountAcrossMonths(rowWage, periodStart, periodEnd).forEach(({ year, monthIndex, amount }) => {
          const key = `${year}-${monthIndex}`;
          grantMonthlyTotals[target.grantId].wage[key] = (grantMonthlyTotals[target.grantId].wage[key] || 0) + amount;
        });
        splitAmountAcrossMonths(rowBenefitAndTax, periodStart, periodEnd).forEach(({ year, monthIndex, amount }) => {
          const key = `${year}-${monthIndex}`;
          grantMonthlyTotals[target.grantId].benefitAndTax[key] = (grantMonthlyTotals[target.grantId].benefitAndTax[key] || 0) + amount;
        });
      });
    });

    const matchedIds = new Set(Object.keys(byEmployee).map((pid) => staffByPaylocityId[pid]?.id).filter(Boolean));
    const missingStaff = workingStaff.filter((s) => s.paylocityId && !matchedIds.has(s.id) && (s.status || "Active") !== "Inactive");

    // Find, for each grant with $ activity, its Template budget(s) and which
    // months couldn't be placed anywhere.
    const budgetTargets = []; // { budgetId, lineUpdates: [{col, wage, benefitAndTax}] }
    const unplacedMonths = []; // { grantTitle, year, monthIndex }
    Object.entries(grantMonthlyTotals).forEach(([grantId, totals]) => {
      const g = grants.find((x) => x.id === grantId);
      const templateBudgets = budgets.filter((b) => b.grantId === grantId && b.budgetType === "Template" && (b.status === "Active" || b.status === "Awarded"));
      const allKeys = new Set([...Object.keys(totals.wage), ...Object.keys(totals.benefitAndTax)]);
      allKeys.forEach((key) => {
        const [year, monthIndex] = key.split("-").map(Number);
        const owner = templateBudgets.find((b) => {
          const cols = monthColumnsForBudget(b.periodStart, b.periodEnd);
          return cols.some((c) => c.year === year && c.monthIndex === monthIndex);
        });
        if (!owner) {
          unplacedMonths.push({ grantTitle: g ? (g.programCode ? `${g.programCode} - ${g.title}` : g.title) : grantId, year, monthIndex });
          return;
        }
        let entry = budgetTargets.find((t) => t.budgetId === owner.id);
        if (!entry) { entry = { budgetId: owner.id, months: {} }; budgetTargets.push(entry); }
        entry.months[key] = { wage: totals.wage[key] || 0, benefitAndTax: totals.benefitAndTax[key] || 0 };
      });
    });

    setResult({ matchedUpdates, newEmployees, missingStaff, budgetTargets, unplacedMonths, matchedCount: matchedUpdates.length, factor, totalDays });
    setStep("review");
  };

  const applyImport = () => {
    const { matchedUpdates, budgetTargets } = result;
    const patchById = {};
    matchedUpdates.forEach((u) => { patchById[u.staffId] = u.patch; });
    setStaff((prev) => prev.map((s) => (patchById[s.id] ? { ...s, ...patchById[s.id] } : s)));

    setBudgets((prev) => {
      let next = prev.map((b) => {
        const target = budgetTargets.find((t) => t.budgetId === b.id);
        if (!target) return b;
        const cols = monthColumnsForBudget(b.periodStart, b.periodEnd);
        const colIndexByYM = {};
        cols.forEach((c, i) => { colIndexByYM[`${c.year}-${c.monthIndex}`] = i; });

        let lines = b.lines.map((l) => ({ ...l, actuals: [...(l.actuals || Array(cols.length).fill(0))] }));
        const findOrCreate = (category, subcategory, type) => {
          let line = lines.find((l) => l.category === category && l.subcategory === subcategory);
          if (!line) {
            line = { ...newLine(), category, subcategory, type, amounts: Array(cols.length).fill(0), actuals: Array(cols.length).fill(0) };
            lines.push(line);
          }
          return line;
        };
        const wageLine = findOrCreate("Wages and Benefits", "5000 - Salary and Wages", "expense");
        const taxLine = findOrCreate("Wages and Benefits", "5900 - Payroll taxes and benefits", "expense");

        Object.entries(target.months).forEach(([key, amounts]) => {
          const i = colIndexByYM[key];
          if (i === undefined) return;
          wageLine.actuals[i] = amounts.wage;
          taxLine.actuals[i] = amounts.benefitAndTax;
        });

        // Advance the cutoff to the end of this import's period, but never
        // backward — a re-upload of an earlier range shouldn't undo a
        // marker that's already more current.
        const newCutoff = periodEnd.slice(0, 7); // "YYYY-MM"
        const cutoff = !b.actualsThrough || newCutoff > b.actualsThrough ? newCutoff : b.actualsThrough;

        return { ...b, lines, actualsThrough: cutoff };
      });

      // Re-run the existing Template -> Operational sync for every budget we
      // just touched, so linked budgets and Org Budget pick this up
      // immediately without a separate save step.
      budgetTargets.forEach((t) => {
        const b = next.find((x) => x.id === t.budgetId);
        if (!b) return;
        const synced = syncActualsToLinkedBudgets(b, next);
        synced.forEach((s) => { next = next.map((x) => (x.id === s.id ? s : x)); });
      });

      return next;
    });

    // The "last import" marker (which future full imports use to suggest
    // their start date) should only move when this run actually touched
    // Actuals — an allocations-only run is a quick snapshot, not a ledger
    // event.
    if (updateComp) {
      setPaylocityLastImport({ periodStart, periodEnd, importedAt: new Date().toISOString().slice(0, 10) });
    }
    const modeTag = updateComp && updateAllocations ? "" : updateComp ? " [Compensation only]" : " [Allocations only]";
    logActivity?.("Personnel", "Updated", `Paylocity import applied — ${matchedUpdates.length} staff updated (${fmtDate(periodStart)}–${fmtDate(periodEnd)})${modeTag}`);
    setStep("done");
  };

  return (
    <Modal title="Import from Paylocity" onClose={onClose} wide>
      {step === "upload" && (
        <div className="space-y-4">
          <p className="text-sm" style={{ color: "#5B6B66" }}>
            Upload the "Labor Distribution Percentages" export grouped by Worked Program (.xlsx). Wages, benefits, payroll tax, and grant allocations are all read from this one file.
          </p>
          {paylocityLastImport && (
            <p className="text-xs" style={{ color: "#8A8F87" }}>
              Last import covered {fmtDate(paylocityLastImport.periodStart)}–{fmtDate(paylocityLastImport.periodEnd)}. Start this report's date range right after that.
            </p>
          )}
          <div className="grid grid-cols-2 gap-4">
            <Field label="Report period start">
              <input type="date" className={inputCls} style={inputStyle} value={periodStart} onChange={(e) => { setPeriodStart(e.target.value); setWideRangeAck(false); }} />
            </Field>
            <Field label="Report period end">
              <input type="date" className={inputCls} style={inputStyle} value={periodEnd} onChange={(e) => { setPeriodEnd(e.target.value); setWideRangeAck(false); }} />
            </Field>
          </div>
          <Field label="Paylocity export (.xlsx)">
            <input type="file" accept=".xlsx" className={inputCls} style={inputStyle} onChange={(e) => e.target.files[0] && parseFile(e.target.files[0])} />
          </Field>
          {rows && <p className="text-xs" style={{ color: "#2F6F53" }}>{rows.length} rows read, covering {distinctCodes.length} program codes.</p>}
          <p className="text-xs font-medium" style={{ color: "#5B6B66" }}>What should this import update?</p>
          <label className="flex items-start gap-2 text-sm p-3 rounded-md border" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>
            <input type="checkbox" className="mt-0.5" checked={updateComp} onChange={(e) => { setUpdateComp(e.target.checked); setWideRangeAck(false); }} />
            <span>
              Compensation — wages, benefits, payroll tax rate
              <span className="block text-xs mt-0.5" style={{ color: "#8A8F87" }}>
                Annualizes this period's rate. Use a full year-to-date range to match a blended actual-plus-forecast total. Writes to the linked Template budget's Actuals, and advances "actuals complete through" to the end of this period.
              </span>
            </span>
          </label>
          <label className="flex items-start gap-2 text-sm p-3 rounded-md border" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>
            <input type="checkbox" className="mt-0.5" checked={updateAllocations} onChange={(e) => { setUpdateAllocations(e.target.checked); setWideRangeAck(false); }} />
            <span>
              Allocations — current grant percentages
              <span className="block text-xs mt-0.5" style={{ color: "#8A8F87" }}>
                Use a single recent pay period for the freshest snapshot of who's working where right now. Never touches any budget.
              </span>
            </span>
          </label>
          {error && <p className="text-xs" style={{ color: "#B5443A" }}>{error}</p>}
          <div className="flex justify-end gap-2 pt-2">
            <button onClick={onClose} className="text-sm px-4 py-2 rounded-md border" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>Cancel</button>
            <button onClick={goToNextStep} className="text-sm px-4 py-2 rounded-md text-white" style={{ background: "#1F5C6B" }}>Continue</button>
          </div>
        </div>
      )}

      {step === "link" && (
        <div className="space-y-4">
          <p className="text-sm" style={{ color: "#5B6B66" }}>
            These Paylocity employees don't have a matching Paylocity ID on an existing staff record yet. Link each to the right person — this only needs to happen once, and future imports will match them automatically. Anyone left as "New employee" won't be updated; they'll show up on the review screen to add as a fresh record instead.
          </p>
          <div className="space-y-2 max-h-96 overflow-y-auto">
            {unmatchedEmployees.map((e) => {
              const suggested = suggestMatch(e.name);
              const currentValue = linkDraft[e.paylocityId] ?? (suggested ? suggested.id : "__new__");
              return (
                <div key={e.paylocityId} className="flex items-center gap-2 p-2 rounded-md border" style={{ borderColor: "#E1E5DE" }}>
                  <div className="text-sm flex-1" style={{ color: "#1C2624" }}>
                    {e.name} <span style={{ color: "#8A8F87", fontFamily: "var(--mono-font)" }}>({e.paylocityId})</span>
                  </div>
                  <select
                    className={inputCls}
                    style={{ ...inputStyle, flex: 2 }}
                    value={currentValue}
                    onChange={(ev) => setLinkDraft((d) => ({ ...d, [e.paylocityId]: ev.target.value }))}
                  >
                    <option value="__new__">New employee — not in the portal yet</option>
                    {staff.filter((s) => !s.paylocityId).sort((a, b) => (a.name || "").localeCompare(b.name || "")).map((s) => (
                      <option key={s.id} value={s.id}>{s.name}{suggested?.id === s.id ? " (suggested match)" : ""}</option>
                    ))}
                  </select>
                </div>
              );
            })}
          </div>
          {error && <p className="text-xs" style={{ color: "#B5443A" }}>{error}</p>}
          <div className="flex justify-end gap-2 pt-2">
            <button onClick={() => setStep("upload")} className="text-sm px-4 py-2 rounded-md border" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>Back</button>
            <button onClick={confirmLinks} className="text-sm px-4 py-2 rounded-md text-white" style={{ background: "#1F5C6B" }}>Continue</button>
          </div>
        </div>
      )}

      {step === "crosswalk" && (
        <div className="space-y-4">
          <p className="text-sm" style={{ color: "#5B6B66" }}>
            These program codes haven't been mapped to a grant or cost center yet. This only needs to happen once per code — future imports will remember it.
          </p>
          <div className="space-y-2 max-h-96 overflow-y-auto">
            {unmappedCodes.map((c) => (
              <div key={c.code} className="flex items-center gap-2 p-2 rounded-md border" style={{ borderColor: "#E1E5DE" }}>
                <div className="text-sm flex-1" style={{ color: "#1C2624" }}>{c.label}</div>
                <select
                  className={inputCls} style={{ ...inputStyle, flex: 2 }}
                  value={crosswalkDraft[c.code]?.ignore ? "__ignore__" : (crosswalkDraft[c.code]?.grantId ? `g:${crosswalkDraft[c.code].grantId}` : crosswalkDraft[c.code]?.costCenterId ? `c:${crosswalkDraft[c.code].costCenterId}` : "")}
                  onChange={(e) => {
                    const v = e.target.value;
                    let next = { grantId: "", costCenterId: "", ignore: false };
                    if (v === "__ignore__") next.ignore = true;
                    else if (v.startsWith("g:")) next.grantId = v.slice(2);
                    else if (v.startsWith("c:")) next.costCenterId = v.slice(2);
                    setCrosswalkDraft((d) => ({ ...d, [c.code]: next }));
                  }}
                >
                  <option value="">Select a grant or cost center…</option>
                  {grants.map((g) => <option key={g.id} value={`g:${g.id}`}>{g.programCode ? `${g.programCode} - ${g.title}` : g.title}</option>)}
                  {costCenters.map((cc) => <option key={cc.id} value={`c:${cc.id}`}>{cc.name}</option>)}
                  <option value="__ignore__">Ignore this program code</option>
                </select>
              </div>
            ))}
          </div>
          {error && <p className="text-xs" style={{ color: "#B5443A" }}>{error}</p>}
          <div className="flex justify-end gap-2 pt-2">
            <button onClick={() => setStep(unmatchedEmployees.length > 0 ? "link" : "upload")} className="text-sm px-4 py-2 rounded-md border" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>Back</button>
            <button onClick={confirmCrosswalk} className="text-sm px-4 py-2 rounded-md text-white" style={{ background: "#1F5C6B" }}>Continue</button>
          </div>
        </div>
      )}

      {step === "review" && result && (
        <div className="space-y-4">
          <p className="text-xs px-3 py-2 rounded-md" style={{ background: "#F6F7F3", color: "#5B6B66" }}>
            Annualizing at ×{result.factor.toFixed(2)} ({result.totalDays}-day period → 365 days). Double-check this looks right before applying — a much larger or smaller multiplier than expected usually means the date range doesn't match the file.
          </p>
          {!updateComp && (
            <p className="text-xs px-3 py-2 rounded-md" style={{ background: "#EAF1F7", color: "#1F5C6B" }}>
              Allocations only — no budget will be touched and no "actuals complete through" marker will move.
            </p>
          )}
          {updateComp && !updateAllocations && (
            <p className="text-xs px-3 py-2 rounded-md" style={{ background: "#EAF1F7", color: "#1F5C6B" }}>
              Compensation only — allocation percentages below are left exactly as they are today.
            </p>
          )}
          <div className="flex items-center gap-3 text-xs">
            <span className="px-2 py-1 rounded-md" style={{ background: "#F0F5F2", color: "#2F6F53" }}>{result.matchedCount} matched and will be updated</span>
            {result.newEmployees.length > 0 && <span className="px-2 py-1 rounded-md" style={{ background: "#FBF3E4", color: "#8A5A0B" }}>{result.newEmployees.length} new, unmatched</span>}
            {result.missingStaff.length > 0 && <span className="px-2 py-1 rounded-md" style={{ background: "#FBEAE8", color: "#B5443A" }}>{result.missingStaff.length} not found in this import</span>}
          </div>

          {result.newEmployees.length > 0 && (
            <div>
              <p className="text-xs font-medium mb-1.5" style={{ color: "#5B6B66" }}>New in file — no matching staff record</p>
              <div className="rounded-md border divide-y" style={{ borderColor: "#E1E5DE" }}>
                {result.newEmployees.map((e) => (
                  <div key={e.paylocityId} className="flex items-center justify-between px-3 py-2 text-sm">
                    <span style={{ color: "#1C2624" }}>{e.name} — Paylocity ID <span style={{ fontFamily: "var(--mono-font)", color: "#8A8F87" }}>{e.paylocityId}</span></span>
                    <button
                      onClick={() => onOpenStaff({ id: uid(), name: e.name, paylocityId: e.paylocityId, allocations: [], status: "Active" })}
                      className="text-xs px-3 py-1.5 rounded-md border"
                      style={{ borderColor: "#E1E5DE", color: "#1F5C6B" }}
                    >
                      Add as staff
                    </button>
                  </div>
                ))}
              </div>
            </div>
          )}

          {result.missingStaff.length > 0 && (
            <div>
              <p className="text-xs font-medium mb-1.5" style={{ color: "#5B6B66" }}>In the portal but not found in this import</p>
              <div className="rounded-md border divide-y" style={{ borderColor: "#E1E5DE" }}>
                {result.missingStaff.map((s) => (
                  <div key={s.id} className="flex items-center justify-between px-3 py-2 text-sm">
                    <span style={{ color: "#1C2624" }}>{s.name}</span>
                    <button onClick={() => onOpenStaff(s)} className="text-xs px-3 py-1.5 rounded-md border" style={{ borderColor: "#E1E5DE", color: "#8A8F87" }}>
                      Review staff record
                    </button>
                  </div>
                ))}
              </div>
              <p className="text-xs mt-1" style={{ color: "#8A8F87" }}>Nothing changes automatically — check each before updating their status.</p>
            </div>
          )}

          {result.unplacedMonths.length > 0 && (
            <div>
              <p className="text-xs font-medium mb-1.5" style={{ color: "#B5443A" }}>No Template budget found to receive these months</p>
              <div className="rounded-md border divide-y text-xs" style={{ borderColor: "#E1E5DE" }}>
                {result.unplacedMonths.map((m, i) => (
                  <div key={i} className="px-3 py-1.5" style={{ color: "#1C2624" }}>{m.grantTitle} — {MONTHS[m.monthIndex]} {m.year}</div>
                ))}
              </div>
            </div>
          )}

          <div className="flex justify-end gap-2 pt-2">
            <button onClick={onClose} className="text-sm px-4 py-2 rounded-md border" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>Cancel</button>
            <button onClick={applyImport} className="text-sm px-4 py-2 rounded-md text-white" style={{ background: "#1F5C6B" }}>Apply import</button>
          </div>
        </div>
      )}

      {step === "done" && (
        <div className="space-y-4 text-center py-6">
          <CheckCircle size={28} style={{ color: "#2F6F53", margin: "0 auto" }} />
          <p className="text-sm" style={{ color: "#1C2624" }}>Import applied — Personnel and Template budget actuals are up to date.</p>
          <button onClick={onClose} className="text-sm px-4 py-2 rounded-md text-white" style={{ background: "#1F5C6B" }}>Done</button>
        </div>
      )}
    </Modal>
  );
}

// ---------- invoicing ----------

function InvoiceModal({ invoice, grants, costCenters = [], budgets = [], currentUserEmail, canEdit = true, onSave, onClose, onDelete }) {
  const grantDefault = grants[0]?.id || "";
  const [form, setForm, undoForm, canUndoForm] = useUndoableState(invoice || {
    id: uid(), grantId: grantDefault, costCenterId: "", invoiceNumber: "", amount: 0,
    submittedDate: "", dueDate: "", paidDate: "", status: "Draft", notes: "", supportingDocsUrl: "",
    periodStart: "", periodEnd: "",
    verification: {
      expectedChecks: {}, vendors: [], glTotal: "", plTotal: "", docsUrl: "",
      discrepancies: [], signedOff: false, signedOffBy: "", signedOffAt: "",
    },
  });
  // Whether this invoice is billed against a grant or a cost center — an
  // invoice can only be one or the other, mirroring how budgets themselves
  // are owned (grantId XOR costCenterId).
  const [ownerType, setOwnerType] = useState(invoice?.costCenterId ? "costCenter" : "grant");
  const set = (k) => (e) => setForm({ ...form, [k]: e.target.value });
  const verification = form.verification || { expectedChecks: {}, vendors: [], glTotal: "", plTotal: "", docsUrl: "", discrepancies: [], signedOff: false };
  const setV = (patch) => setForm({ ...form, verification: { ...verification, ...patch } });

  // Every category/subcategory with planned activity in this invoice's period,
  // pulled from the grant's own budget — so Larry isn't retyping what he
  // expects to see each time, and the checklist can't drift out of sync with
  // what's actually budgeted.
  const ymNum = (y, m) => y * 12 + m;
  const toYM = (dateStr) => { const [y, m] = dateStr.split("-").map(Number); return ymNum(y, m - 1); };
  const expectedCategories = useMemo(() => {
    if ((!form.grantId && !form.costCenterId) || !form.periodStart || !form.periodEnd) return [];
    const startYM = toYM(form.periodStart);
    const endYM = toYM(form.periodEnd);
    const relevant = budgets.filter((b) => (
      (form.grantId && b.grantId === form.grantId) || (form.costCenterId && b.costCenterId === form.costCenterId)
    ) && (b.status === "Active" || b.status === "Awarded"));
    const seen = new Set();
    const result = [];
    relevant.forEach((b) => {
      if (!b.periodStart || !b.periodEnd) return;
      const cols = monthColumnsForBudget(b.periodStart, b.periodEnd);
      (b.lines || []).forEach((l) => {
        const key = `${l.category}|||${l.subcategory || ""}`;
        if (seen.has(key)) return;
        const amounts = l.amounts || [];
        const hasActivity = cols.some((col, i) => {
          const colYM = ymNum(col.year, col.monthIndex);
          return colYM >= startYM && colYM <= endYM && (Number(amounts[i]) || 0) > 0;
        });
        if (hasActivity) { seen.add(key); result.push({ key, category: l.category, subcategory: l.subcategory }); }
      });
    });
    return result;
  }, [form.grantId, form.costCenterId, form.periodStart, form.periodEnd, budgets]);

  const addVendor = () => setV({ vendors: [...verification.vendors, { id: uid(), name: "", expectedAmount: "", found: "pending" }] });
  const updateVendor = (id, patch) => setV({ vendors: verification.vendors.map((v) => (v.id === id ? { ...v, ...patch } : v)) });
  const removeVendor = (id) => setV({ vendors: verification.vendors.filter((v) => v.id !== id) });

  const [newDiscrepancy, setNewDiscrepancy] = useState("");
  const addDiscrepancy = () => {
    if (!newDiscrepancy.trim()) return;
    setV({ discrepancies: [...verification.discrepancies, { id: uid(), text: newDiscrepancy.trim(), status: "open", createdAt: new Date().toISOString().slice(0, 10) }] });
    setNewDiscrepancy("");
  };
  const toggleDiscrepancy = (id) => setV({
    discrepancies: verification.discrepancies.map((d) => (d.id === id
      ? { ...d, status: d.status === "resolved" ? "open" : "resolved", resolvedAt: d.status === "resolved" ? "" : new Date().toISOString().slice(0, 10) }
      : d)),
  });

  const glTotal = Number(verification.glTotal) || 0;
  const plTotal = Number(verification.plTotal) || 0;
  const variance = glTotal - plTotal;
  const hasOpenDiscrepancy = verification.discrepancies.some((d) => d.status !== "resolved");

  return (
    <Modal title={invoice ? (canEdit ? "Edit invoice" : "View invoice") : "New invoice"} onClose={onClose}>
      <fieldset disabled={!canEdit} style={{ border: "none", margin: 0, padding: 0, minWidth: 0 }}>
      <div className="space-y-4">
        <Field label="Grant or cost center">
          <div className="flex items-center gap-2">
            <div className="inline-flex rounded-md border overflow-hidden shrink-0" style={{ borderColor: "#E1E5DE" }}>
              <button
                type="button"
                onClick={() => { setOwnerType("grant"); setForm({ ...form, costCenterId: "" }); }}
                className="px-3 py-2 text-xs font-medium"
                style={{ background: ownerType === "grant" ? "#1F5C6B" : "#FFFFFF", color: ownerType === "grant" ? "#FFFFFF" : "#5B6B66" }}
              >
                Grant
              </button>
              <button
                type="button"
                onClick={() => { setOwnerType("costCenter"); setForm({ ...form, grantId: "" }); }}
                className="px-3 py-2 text-xs font-medium"
                style={{ background: ownerType === "costCenter" ? "#1F5C6B" : "#FFFFFF", color: ownerType === "costCenter" ? "#FFFFFF" : "#5B6B66" }}
              >
                Cost Center
              </button>
            </div>
            {ownerType === "grant" ? (
              <div className="flex-1">
                <GrantPicker grants={grants} value={form.grantId} onChange={(v) => setForm({ ...form, grantId: v })} placeholder="Select a grant" />
              </div>
            ) : (
              <select className={inputCls} style={inputStyle} value={form.costCenterId} onChange={(e) => setForm({ ...form, costCenterId: e.target.value })}>
                <option value="">Select a cost center</option>
                {costCenters.map((cc) => <option key={cc.id} value={cc.id}>{cc.name}</option>)}
              </select>
            )}
          </div>
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
          <Field label="Period start">
            <input type="date" className={inputCls} style={inputStyle} value={form.periodStart} onChange={set("periodStart")} />
          </Field>
          <Field label="Period end">
            <input type="date" className={inputCls} style={inputStyle} value={form.periodEnd} onChange={set("periodEnd")} />
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
        <Field label="Supporting documents (SharePoint folder URL)">
          <div className="flex items-center gap-2">
            <input className={inputCls} style={inputStyle} value={form.supportingDocsUrl || ""} onChange={set("supportingDocsUrl")} placeholder="https://…" />
            {form.supportingDocsUrl && (
              <a
                href={form.supportingDocsUrl} target="_blank" rel="noreferrer"
                className="inline-flex items-center gap-1 text-xs px-3 py-2 rounded-md border shrink-0"
                style={{ borderColor: "#E1E5DE", color: "#1F5C6B" }}
              >
                <ExternalLink size={13} /> Supporting Documents
              </a>
            )}
          </div>
        </Field>

        <div className="border-t pt-4 mt-2" style={{ borderColor: "#E1E5DE" }}>
          <div className="flex items-center justify-between mb-3">
            <h3 className="font-display text-sm" style={{ color: "#1C2624" }}>Pre-invoice verification</h3>
            {hasOpenDiscrepancy && (
              <span className="inline-flex items-center gap-1 text-xs px-2 py-0.5 rounded-md" style={{ background: "#FBEAE8", color: "#B5443A" }}>
                <AlertCircle size={12} /> Unresolved discrepancy
              </span>
            )}
          </div>

          <div className="mb-4">
            <p className="text-xs font-medium mb-1.5" style={{ color: "#5B6B66" }}>Expected activity this period</p>
            {!form.grantId || !form.periodStart || !form.periodEnd ? (
              <p className="text-xs" style={{ color: "#8A8F87" }}>Set the grant and invoice period above to pull expected budget categories.</p>
            ) : expectedCategories.length === 0 ? (
              <p className="text-xs" style={{ color: "#8A8F87" }}>No budgeted activity found for this grant in this period.</p>
            ) : (
              <div className="space-y-1">
                {expectedCategories.map((c) => (
                  <label key={c.key} className="flex items-center gap-2 text-sm" style={{ color: "#1C2624" }}>
                    <input
                      type="checkbox"
                      checked={!!verification.expectedChecks[c.key]}
                      onChange={(e) => setV({ expectedChecks: { ...verification.expectedChecks, [c.key]: e.target.checked } })}
                    />
                    {c.category}{c.subcategory ? ` / ${c.subcategory}` : ""}
                  </label>
                ))}
              </div>
            )}
          </div>

          <div className="mb-4">
            <p className="text-xs font-medium mb-1.5" style={{ color: "#5B6B66" }}>Vendor / subcontractor match to GL</p>
            <div className="space-y-2">
              {verification.vendors.map((v) => (
                <div key={v.id} className="flex items-center gap-2">
                  <input
                    className={inputCls} style={{ ...inputStyle, flex: 2 }}
                    value={v.name} onChange={(e) => updateVendor(v.id, { name: e.target.value })}
                    placeholder="Vendor name"
                  />
                  <input
                    type="number" className={inputCls} style={{ ...inputStyle, flex: 1 }}
                    value={v.expectedAmount} onChange={(e) => updateVendor(v.id, { expectedAmount: e.target.value })}
                    placeholder="Amount"
                  />
                  <select
                    className={inputCls} style={{ ...inputStyle, flex: 1 }}
                    value={v.found} onChange={(e) => updateVendor(v.id, { found: e.target.value })}
                  >
                    <option value="pending">Pending</option>
                    <option value="matched">Matched</option>
                    <option value="notfound">Not found</option>
                  </select>
                  <button onClick={() => removeVendor(v.id)} className="p-1.5 rounded hover:bg-red-50 shrink-0">
                    <X size={14} style={{ color: "#B5443A" }} />
                  </button>
                </div>
              ))}
              <button onClick={addVendor} className="text-xs inline-flex items-center gap-1" style={{ color: "#1F5C6B" }}>
                <Plus size={13} /> Add vendor / subcontractor
              </button>
            </div>
          </div>

          <div className="mb-4">
            <p className="text-xs font-medium mb-1.5" style={{ color: "#5B6B66" }}>GL vs P/L comparison</p>
            <div className="grid grid-cols-3 gap-3 mb-3">
              <Field label="GL detail total">
                <input type="number" className={inputCls} style={inputStyle} value={verification.glTotal} onChange={(e) => setV({ glTotal: e.target.value })} />
              </Field>
              <Field label="P/L total">
                <input type="number" className={inputCls} style={inputStyle} value={verification.plTotal} onChange={(e) => setV({ plTotal: e.target.value })} />
              </Field>
              <Field label="Variance">
                <div className="px-3 py-2 rounded-md border text-sm" style={{ borderColor: "#E1E5DE", color: variance !== 0 ? "#B5443A" : "#2F6F53", fontVariantNumeric: "tabular-nums" }}>
                  {fmt(variance)}
                </div>
              </Field>
            </div>
            <Field label="Verification documents (Sage GL/P&L export folder)">
              <div className="flex items-center gap-2">
                <input className={inputCls} style={inputStyle} value={verification.docsUrl} onChange={(e) => setV({ docsUrl: e.target.value })} placeholder="https://…" />
                {verification.docsUrl && (
                  <a
                    href={verification.docsUrl} target="_blank" rel="noreferrer"
                    className="inline-flex items-center gap-1 text-xs px-3 py-2 rounded-md border shrink-0"
                    style={{ borderColor: "#E1E5DE", color: "#1F5C6B" }}
                  >
                    <ExternalLink size={13} /> Open
                  </a>
                )}
              </div>
            </Field>
          </div>

          <div className="mb-4">
            <p className="text-xs font-medium mb-1.5" style={{ color: "#5B6B66" }}>Discrepancy log</p>
            <div className="space-y-2 mb-2">
              {verification.discrepancies.map((d) => (
                <div
                  key={d.id}
                  className="flex items-start gap-2 text-xs px-3 py-2 rounded-md"
                  style={{ background: d.status === "resolved" ? "#F0F5F2" : "#FBEAE8", color: d.status === "resolved" ? "#2F6F53" : "#8A392F" }}
                >
                  <button onClick={() => toggleDiscrepancy(d.id)} className="shrink-0 mt-0.5">
                    {d.status === "resolved" ? <CheckCircle size={13} /> : <Circle size={13} />}
                  </button>
                  <div className="flex-1">
                    <div>{d.text}</div>
                    <div style={{ opacity: 0.75 }}>
                      Flagged {fmtDate(d.createdAt)}{d.status === "resolved" && d.resolvedAt ? ` · Resolved ${fmtDate(d.resolvedAt)}` : ""}
                    </div>
                  </div>
                </div>
              ))}
            </div>
            <div className="flex items-center gap-2">
              <input
                className={inputCls} style={{ ...inputStyle, flex: 1 }}
                value={newDiscrepancy} onChange={(e) => setNewDiscrepancy(e.target.value)}
                placeholder="Describe what's missing or doesn't match…"
                onKeyDown={(e) => { if (e.key === "Enter") { e.preventDefault(); addDiscrepancy(); } }}
              />
              <button onClick={addDiscrepancy} className="text-xs px-3 py-2 rounded-md border shrink-0" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>
                Add
              </button>
            </div>
          </div>

          <label className="flex items-center gap-2 text-sm" style={{ color: "#1C2624" }}>
            <input
              type="checkbox"
              checked={!!verification.signedOff}
              onChange={(e) => setV(e.target.checked
                ? { signedOff: true, signedOffBy: currentUserEmail || "Unknown", signedOffAt: new Date().toISOString().slice(0, 10) }
                : { signedOff: false, signedOffBy: "", signedOffAt: "" })}
            />
            All flagged items resolved or documented — ready to invoice
          </label>
          {verification.signedOff && verification.signedOffBy && (
            <p className="text-xs mt-1" style={{ color: "#8A8F87" }}>Signed off by {verification.signedOffBy} on {fmtDate(verification.signedOffAt)}</p>
          )}
        </div>
      </div>
      </fieldset>

      <div className="flex justify-between gap-2 mt-6">
        {onDelete ? (
          <button onClick={onDelete} className="px-4 py-2 rounded-md text-sm border" style={{ borderColor: "#E1E5DE", color: "#B5443A" }}>Delete invoice</button>
        ) : <span />}
        <div className="flex gap-2">
          {canEdit && canUndoForm && (
            <button onClick={undoForm} className="px-3 py-2 rounded-md text-sm border inline-flex items-center gap-1.5" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>
              <Undo2 size={14} /> Undo
            </button>
          )}
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

function InvoicingView({ grants, invoices, setInvoices, setTrash, currentUserEmail, canEdit, initialOpenInvoiceId, logActivity, budgets = [], costCenters = [] }) {
  const [modal, setModal] = useState(() => (initialOpenInvoiceId ? invoices.find((i) => i.id === stripNonce(initialOpenInvoiceId)) || null : null));
  const [confirm, setConfirm] = useState(null);
  const [grantFilter, setGrantFilter] = useState("All");
  const [statusFilter, setStatusFilter] = useState("All");
  const [period, setPeriod] = useState("month"); // month | year | all

  const visible = invoices
    .filter((i) => grantFilter === "All" || i.grantId === grantFilter)
    .filter((i) => statusFilter === "All" || i.status === statusFilter)
    .sort((a, b) => new Date(b.submittedDate || 0) - new Date(a.submittedDate || 0));

  // "Flow" figures (Invoiced, Paid, Avg Days to Pay) are scoped to the selected period.
  // "Stock" figures (Outstanding, Overdue, aging) always reflect right now, regardless
  // of the period picker — money owed today doesn't stop being owed because it was
  // invoiced outside the selected window.
  const invoicedInPeriod = invoices.filter((i) => inPeriod(i.submittedDate, period));
  const paidInPeriod = invoices.filter((i) => i.status === "Paid" && inPeriod(i.paidDate, period));
  const totalInvoiced = invoicedInPeriod.reduce((a, i) => a + (Number(i.amount) || 0), 0);
  const totalPaid = paidInPeriod.reduce((a, i) => a + (Number(i.amount) || 0), 0);
  const totalOutstanding = invoices.filter((i) => i.status === "Submitted").reduce((a, i) => a + (Number(i.amount) || 0), 0);
  const overdueInvoices = invoices.filter(isInvoiceOverdue);
  const overdueCount = overdueInvoices.length;

  const payDelays = paidInPeriod.map(daysToPay).filter((d) => d !== null);
  const avgDaysToPay = payDelays.length > 0 ? Math.round(payDelays.reduce((a, d) => a + d, 0) / payDelays.length) : null;

  const agingBuckets = useMemo(() => {
    const buckets = { "0-30": 0, "31-60": 0, "61-90": 0, "90+": 0 };
    invoices.filter((i) => i.status === "Submitted").forEach((i) => {
      const d = daysOutstanding(i) ?? 0;
      const amt = Number(i.amount) || 0;
      if (d <= 30) buckets["0-30"] += amt;
      else if (d <= 60) buckets["31-60"] += amt;
      else if (d <= 90) buckets["61-90"] += amt;
      else buckets["90+"] += amt;
    });
    return buckets;
  }, [invoices]);

  const trendData = useMemo(() => {
    const months = [];
    const now = new Date();
    for (let i = 11; i >= 0; i--) {
      const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
      months.push({ year: d.getFullYear(), month: d.getMonth(), label: `${MONTHS[d.getMonth()]} ${String(d.getFullYear()).slice(2)}`, invoiced: 0, paid: 0 });
    }
    invoices.forEach((inv) => {
      if (inv.submittedDate) {
        const d = new Date(inv.submittedDate + "T00:00:00");
        const bucket = months.find((m) => m.year === d.getFullYear() && m.month === d.getMonth());
        if (bucket) bucket.invoiced += Number(inv.amount) || 0;
      }
      if (inv.paidDate) {
        const d = new Date(inv.paidDate + "T00:00:00");
        const bucket = months.find((m) => m.year === d.getFullYear() && m.month === d.getMonth());
        if (bucket) bucket.paid += Number(inv.amount) || 0;
      }
    });
    return months;
  }, [invoices]);

  const [grantSort, setGrantSort] = useState({ key: "outstanding", dir: "desc" });
  const byGrantData = useMemo(() => {
    const map = {};
    invoices.forEach((inv) => {
      const key = inv.grantId || "unlinked";
      if (!map[key]) {
        const g = grants.find((x) => x.id === inv.grantId);
        map[key] = { grantId: inv.grantId, title: g ? (g.programCode ? `${g.programCode} - ${g.title}` : g.title) : "No grant linked", invoiced: 0, outstanding: 0, delays: [] };
      }
      map[key].invoiced += Number(inv.amount) || 0;
      if (inv.status === "Submitted") map[key].outstanding += Number(inv.amount) || 0;
      const d = daysToPay(inv);
      if (d !== null) map[key].delays.push(d);
    });
    const rows = Object.values(map).map((r) => ({
      ...r,
      avgDays: r.delays.length > 0 ? Math.round(r.delays.reduce((a, d) => a + d, 0) / r.delays.length) : null,
    }));
    rows.sort((a, b) => {
      const dir = grantSort.dir === "asc" ? 1 : -1;
      const av = a[grantSort.key] ?? -1, bv = b[grantSort.key] ?? -1;
      return av < bv ? -1 * dir : av > bv ? 1 * dir : 0;
    });
    return rows;
  }, [invoices, grants, grantSort]);
  const toggleGrantSort = (key) => setGrantSort((s) => (s.key === key ? { key, dir: s.dir === "asc" ? "desc" : "asc" } : { key, dir: "desc" }));

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
  const exportCsv = async () => {
    const HEADER_FILL = "FFF6F7F3";
    const RED = "FFB5443A";
    const GREEN = "FF2F6F53";
    const AMBER = "FFC08A2E";

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Invoices");

    const header = ["Grant / Cost Center", "Invoice #", "Amount", "Status", "Submitted", "Due", "Paid", "Days outstanding"];
    ws.mergeCells(1, 1, 1, header.length);
    const titleCell = ws.getCell(1, 1);
    titleCell.value = "Nation's Finest — Invoices";
    titleCell.font = { bold: true, size: 13 };
    ws.getRow(1).height = 22;

    ws.mergeCells(2, 1, 2, header.length);
    ws.getCell(2, 1).value = `Generated ${fmtDate(new Date().toISOString().slice(0, 10))}`;
    ws.getCell(2, 1).font = { italic: true, size: 9, color: { argb: "FF8A8F87" } };

    const headerRowIdx = 4;
    ws.getRow(headerRowIdx).values = header;
    ws.getRow(headerRowIdx).eachCell((cell) => {
      cell.font = { bold: true };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: HEADER_FILL } };
    });

    const statusColor = { Paid: GREEN, Submitted: AMBER, Rejected: RED, Draft: "FF8A8F87" };
    let r = headerRowIdx + 1;
    invoices.forEach((i) => {
      const g = grants.find((x) => x.id === i.grantId);
      const cc = costCenters.find((x) => x.id === i.costCenterId);
      const overdue = isInvoiceOverdue(i);
      const days = daysOutstanding(i);
      ws.getRow(r).values = [
        g?.title || cc?.name || "",
        i.invoiceNumber || "",
        Number(i.amount) || 0,
        i.status || "",
        i.submittedDate ? fmtDate(i.submittedDate) : "",
        i.dueDate ? fmtDate(i.dueDate) : "",
        i.paidDate ? fmtDate(i.paidDate) : "",
        days ?? "",
      ];
      ws.getCell(r, 3).numFmt = "$#,##0";
      ws.getCell(r, 4).font = { bold: true, color: { argb: statusColor[i.status] || "FF1C2624" } };
      if (overdue) {
        ws.getCell(r, 6).font = { bold: true, color: { argb: RED } };
      }
      r++;
    });

    ws.columns = [{ width: 34 }, { width: 16 }, { width: 14 }, { width: 12 }, { width: 13 }, { width: 13 }, { width: 13 }, { width: 16 }];
    ws.views = [{ state: "frozen", ySplit: headerRowIdx }];

    const buffer = await wb.xlsx.writeBuffer();
    downloadFile("nations-finest-invoices.xlsx", buffer, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
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
            <Download size={15} /> Export Excel
          </button>
          {canEdit && (
            <button onClick={() => setModal("new")} className="inline-flex items-center gap-1.5 px-3 py-2 rounded-md text-sm text-white" style={{ background: "#1F5C6B" }}>
              <Plus size={16} /> New invoice
            </button>
          )}
        </div>
      </div>

      <div className="flex items-center gap-2">
        <span className="text-xs" style={{ color: "#8A8F87" }}>Show Invoiced/Paid/Avg Days to Pay for:</span>
        <div className="inline-flex rounded-md border overflow-hidden" style={{ borderColor: "#E1E5DE" }}>
          {[{ key: "month", label: "This Month" }, { key: "year", label: "This Year" }, { key: "all", label: "All Time" }].map((p) => (
            <button
              key={p.key}
              onClick={() => setPeriod(p.key)}
              className="px-3 py-1.5 text-xs font-medium"
              style={{ background: period === p.key ? "#1F5C6B" : "#FFFFFF", color: period === p.key ? "#FFFFFF" : "#5B6B66" }}
            >
              {p.label}
            </button>
          ))}
        </div>
      </div>

      <div className="grid grid-cols-2 sm:grid-cols-5 gap-4">
        <StatCard label="Total invoiced" value={fmt(totalInvoiced)} />
        <StatCard label="Total paid" value={fmt(totalPaid)} />
        <StatCard label="Avg days to pay" value={avgDaysToPay === null ? "—" : `${avgDaysToPay}d`} sub={payDelays.length > 0 ? `${payDelays.length} paid invoice${payDelays.length === 1 ? "" : "s"}` : "No paid invoices yet"} />
        <StatCard label="Outstanding" value={fmt(totalOutstanding)} sub={`${invoices.filter((i) => i.status === "Submitted").length} submitted, unpaid — as of today`} />
        <StatCard label="Overdue" value={overdueCount} sub={overdueCount > 0 ? fmt(overdueInvoices.reduce((a, i) => a + (Number(i.amount) || 0), 0)) + " past due" : "None"} />
      </div>

      <div className="bg-white rounded-lg border p-5" style={{ borderColor: "#E1E5DE" }}>
        <h2 className="font-display text-base mb-3" style={{ color: "#1C2624" }}>Outstanding by age</h2>
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
          {[["0-30", "0–30 days"], ["31-60", "31–60 days"], ["61-90", "61–90 days"], ["90+", "90+ days"]].map(([key, label]) => (
            <div key={key}>
              <div className="text-xs" style={{ color: "#8A8F87" }}>{label}</div>
              <div className="text-lg font-medium" style={{ color: key === "90+" && agingBuckets[key] > 0 ? "#B5443A" : "#1C2624", fontVariantNumeric: "tabular-nums" }}>
                {fmt(agingBuckets[key])}
              </div>
            </div>
          ))}
        </div>
      </div>

      <div className="bg-white rounded-lg border p-5" style={{ borderColor: "#E1E5DE" }}>
        <h2 className="font-display text-base mb-3" style={{ color: "#1C2624" }}>Invoiced vs. paid, last 12 months</h2>
        <ResponsiveContainer width="100%" height={260}>
          <LineChart data={trendData}>
            <CartesianGrid strokeDasharray="3 3" stroke="#E1E5DE" />
            <XAxis dataKey="label" tick={{ fontSize: 11 }} />
            <YAxis tickFormatter={(v) => `$${v / 1000}k`} tick={{ fontSize: 11 }} />
            <Tooltip formatter={(v) => fmt(v)} />
            <Legend />
            <Line type="monotone" dataKey="invoiced" name="Invoiced" stroke="#5B7FA6" strokeWidth={2} dot={false} />
            <Line type="monotone" dataKey="paid" name="Paid" stroke="#2F6F53" strokeWidth={2} dot={false} />
          </LineChart>
        </ResponsiveContainer>
      </div>

      <div className="bg-white rounded-lg border p-5" style={{ borderColor: "#E1E5DE" }}>
        <h2 className="font-display text-base mb-3" style={{ color: "#1C2624" }}>By grant — all time</h2>
        {byGrantData.length === 0 ? (
          <p className="text-sm" style={{ color: "#8A8F87" }}>No invoices yet.</p>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead>
                <tr style={{ color: "#8A8F87" }}>
                  <th className="text-left py-1.5 font-medium cursor-pointer" onClick={() => toggleGrantSort("title")}>Grant</th>
                  <th className="text-right py-1.5 font-medium cursor-pointer" onClick={() => toggleGrantSort("invoiced")}>Total invoiced</th>
                  <th className="text-right py-1.5 font-medium cursor-pointer" onClick={() => toggleGrantSort("outstanding")}>Outstanding</th>
                  <th className="text-right py-1.5 font-medium cursor-pointer" onClick={() => toggleGrantSort("avgDays")}>Avg days to pay</th>
                </tr>
              </thead>
              <tbody>
                {byGrantData.map((r) => (
                  <tr key={r.grantId || "unlinked"} className="border-t" style={{ borderColor: "#E1E5DE" }}>
                    <td className="py-1.5" style={{ color: "#1C2624" }}>{r.title}</td>
                    <td className="py-1.5 text-right" style={{ fontVariantNumeric: "tabular-nums" }}>{fmt(r.invoiced)}</td>
                    <td className="py-1.5 text-right" style={{ fontVariantNumeric: "tabular-nums", color: r.outstanding > 0 ? "#B5443A" : "#8A8F87" }}>{fmt(r.outstanding)}</td>
                    <td className="py-1.5 text-right" style={{ fontVariantNumeric: "tabular-nums" }}>{r.avgDays === null ? "—" : `${r.avgDays}d`}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
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
                const cc = costCenters.find((x) => x.id === i.costCenterId);
                const overdue = isInvoiceOverdue(i);
                const outstanding = daysOutstanding(i);
                const flagged = i.verification?.discrepancies?.some((d) => d.status !== "resolved");
                return (
                  <tr key={i.id} className="border-t cursor-pointer hover:bg-stone-50" style={{ borderColor: overdue ? "#B5443A" : "#E1E5DE" }} onClick={() => setModal(i)}>
                    <td className="px-3 py-2" style={{ color: "#1C2624" }}>{g?.title || cc?.name || "Unknown grant"}</td>
                    <td className="px-3 py-2" style={{ color: "#1C2624" }}>
                      <div className="flex items-center gap-1.5">
                        {i.invoiceNumber || "—"}
                        {flagged && (
                          <span title="Unresolved verification discrepancy">
                            <AlertCircle size={13} style={{ color: "#B5443A" }} />
                          </span>
                        )}
                      </div>
                    </td>
                    <td className="px-3 py-2 text-right" style={{ fontVariantNumeric: "tabular-nums", color: "#1C2624" }}>{fmt(i.amount)}</td>
                    <td className="px-3 py-2"><Badge color={statusColor[i.status]}>{i.status}</Badge></td>
                    <td className="px-3 py-2 text-right" style={{ fontVariantNumeric: "tabular-nums", color: "#5B6B66" }}>{fmtDate(i.submittedDate)}</td>
                    <td className="px-3 py-2 text-right" style={{ fontVariantNumeric: "tabular-nums", color: overdue ? "#B5443A" : "#5B6B66" }}>{fmtDate(i.dueDate)}</td>
                    <td className="px-3 py-2 text-right" style={{ fontVariantNumeric: "tabular-nums", color: "#5B6B66" }}>{i.paidDate ? fmtDate(i.paidDate) : "—"}</td>
                    <td className="px-3 py-2 text-right" style={{ fontVariantNumeric: "tabular-nums", color: overdue ? "#B5443A" : "#5B6B66" }}>{outstanding !== null ? `${outstanding}d` : "—"}</td>
                    <td className="px-3 py-2 text-center" onClick={(e) => e.stopPropagation()}>
                      <div className="flex items-center justify-center gap-1">
                        {i.supportingDocsUrl && (
                          <a href={i.supportingDocsUrl} target="_blank" rel="noreferrer" className="p-1 rounded hover:bg-stone-100" title="Supporting Documents">
                            <ExternalLink size={14} style={{ color: "#1F5C6B" }} />
                          </a>
                        )}
                        {canEdit && (
                          <button onClick={() => setConfirm(i.id)} className="p-1 rounded hover:bg-red-50">
                            <Trash2 size={14} style={{ color: "#B5443A" }} />
                          </button>
                        )}
                      </div>
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
          costCenters={costCenters}
          budgets={budgets}
          currentUserEmail={currentUserEmail}
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
  const withBudgets = grants.filter((g) => budgets.some((b) => b.grantId === g.id && b.budgetType === "Operational"));

  return (
    <div className="space-y-4">
      <div>
        <h1 className="font-display text-2xl" style={{ color: "#1C2624" }}>Burn Rate</h1>
        <p className="text-sm mt-1" style={{ color: "#5B6B66" }}>
          Paced against each grant's Operational budget — the version actually contracted with the grantor — not the internal Template. Uses recorded actuals where you've entered them (Budgets → Actual). Falls back to the planned schedule for any grant without actuals yet.
        </p>
      </div>

      {withBudgets.length === 0 ? (
        <div className="bg-white rounded-lg border p-10 text-center" style={{ borderColor: "#E1E5DE", color: "#8A8F87" }}>
          No grants with an Operational budget yet — add one to see burn rate.
        </div>
      ) : (
        <div className="overflow-x-auto border rounded-lg bg-white" style={{ borderColor: "#E1E5DE" }}>
          <table className="w-full text-sm">
            <thead>
              <tr style={{ background: "#F6F7F3" }}>
                <th className="text-left px-3 py-2 text-xs" style={{ minWidth: 200 }}>Grant</th>
                <th className="text-right px-3 py-2 text-xs">Award</th>
                <th className="text-right px-3 py-2 text-xs">Budgeted (current period)</th>
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
        const cc = costCenters.find((x) => x.id === i.costCenterId);
        return { type: "Invoice", label: i.invoiceNumber || "Untitled invoice", sub: g?.title || cc?.name || "", action: () => goTo("invoicing", null, null, i.id) };
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
  const [showHealthCheck, setShowHealthCheck] = useState(false);

  // Runs the same categories of checks we've done manually against backup
  // exports in the past — orphaned records, format drift, missing links,
  // overdue compliance items — directly against live data, on demand.
  const healthIssues = useMemo(() => {
    const issues = [];
    const grantById = {};
    grants.forEach((g) => { grantById[g.id] = g; });
    const ccIds = new Set((costCenters || []).map((c) => c.id));
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    budgets.forEach((b) => {
      if (b.grantId && !grantById[b.grantId]) issues.push({ level: "error", area: "Budgets", text: `"${b.title}" references a grant that no longer exists.` });
      if (b.costCenterId && !ccIds.has(b.costCenterId)) issues.push({ level: "error", area: "Budgets", text: `"${b.title}" references a cost center that no longer exists.` });
      if (!b.grantId && !b.costCenterId) issues.push({ level: "error", area: "Budgets", text: `"${b.title}" isn't linked to a grant or cost center.` });
      if (b.fy && !/^\d{4}$/.test(String(b.fy))) issues.push({ level: "warn", area: "Budgets", text: `"${b.title}" has a non-standard fiscal year value ("${b.fy}") — expected a plain 4-digit year.` });

      if (!b.periodStart || !b.periodEnd) {
        issues.push({ level: "error", area: "Budgets", text: `"${b.title}" is missing a period start or end date.` });
        return;
      }
      const cols = monthColumnsForBudget(b.periodStart, b.periodEnd);
      (b.lines || []).forEach((l) => {
        const amtLen = (l.amounts || []).length;
        if (amtLen > 0 && amtLen < cols.length) {
          issues.push({ level: "warn", area: "Budgets", text: `"${b.title}" — line "${l.category}${l.subcategory ? "/" + l.subcategory : ""}" has fewer months of data (${amtLen}) than its period (${cols.length}).` });
        }
        (l.actuals || []).forEach((v, i) => {
          if (Number(v) < 0) issues.push({ level: "info", area: "Budgets", text: `"${b.title}" — "${l.category}${l.subcategory ? "/" + l.subcategory : ""}" has a negative actual in ${cols[i]?.label || `month ${i + 1}`}.` });
        });
      });
      const cutoff = parseActualsThrough(b.actualsThrough);
      if (cutoff) {
        if (cutoff.year > today.getFullYear() || (cutoff.year === today.getFullYear() && cutoff.monthIndex > today.getMonth())) {
          issues.push({ level: "error", area: "Budgets", text: `"${b.title}" has "Actuals complete through" set to a future month.` });
        }
        if (!cols.some((c) => c.year === cutoff.year && c.monthIndex === cutoff.monthIndex)) {
          issues.push({ level: "warn", area: "Budgets", text: `"${b.title}"'s "Actuals complete through" date falls outside its own budget period.` });
        }
      }
      if ((b.status === "Active" || b.status === "Awarded") && (!b.approvedBy || !b.approvedAt)) {
        issues.push({ level: "info", area: "Budgets", text: `"${b.title}" is ${b.status} but has no recorded approver or approval date.` });
      }
    });

    const grantIdsWithBudgets = new Set(budgets.map((b) => b.grantId).filter(Boolean));
    grants.forEach((g) => {
      if ((g.stage === "Active" || g.stage === "Awarded") && !grantIdsWithBudgets.has(g.id)) {
        issues.push({ level: "warn", area: "Grants", text: `"${g.programCode ? g.programCode + " - " : ""}${g.title}" is ${g.stage} but has no budget yet.` });
      }
    });

    (reports || []).forEach((r) => {
      if (!r.grantId) issues.push({ level: "info", area: "Reports", text: `"${r.title}" isn't linked to a grant.` });
      if (r.dueDate && !["Submitted", "Complete", "Completed", "Done"].includes(r.status)) {
        const d = new Date(r.dueDate + "T00:00:00");
        if (d < today) issues.push({ level: "warn", area: "Reports", text: `"${r.title}" was due ${r.dueDate} and is still "${r.status}".` });
      }
    });

    (tasks || []).forEach((t) => {
      if (t.dueDate && !["Done", "Completed", "Closed"].includes(t.status)) {
        const d = new Date(t.dueDate + "T00:00:00");
        if (d < today) issues.push({ level: "warn", area: "Tasks", text: `"${t.title}" was due ${t.dueDate} and is still "${t.status}".` });
      }
    });

    return issues;
  }, [grants, budgets, reports, tasks, costCenters]);

  const issueCounts = { error: 0, warn: 0, info: 0 };
  healthIssues.forEach((i) => { issueCounts[i.level] = (issueCounts[i.level] || 0) + 1; });
  const levelColor = { error: "#B5443A", warn: "#C08A2E", info: "#5B7FA6" };
  const levelLabel = { error: "Needs fixing", warn: "Worth a look", info: "FYI" };

  const [restoreError, setRestoreError] = useState("");
  const [restoreSummary, setRestoreSummary] = useState("");
  const [importError, setImportError] = useState("");
  const [importSummary, setImportSummary] = useState("");
  const [budgetImportError, setBudgetImportError] = useState("");
  const [budgetImportSummary, setBudgetImportSummary] = useState("");
  const [actualsImportError, setActualsImportError] = useState("");
  const [actualsImportSummary, setActualsImportSummary] = useState("");
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

  const handleImportActuals = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setActualsImportError(""); setActualsImportSummary("");
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

        // Work on a deep-enough copy so multiple rows can update the same budget.
        const updated = {}; // budgetId -> working copy of that budget
        let rowsApplied = 0, linesUpdated = 0, linesCreated = 0, skippedNoGrant = 0, skippedNoBudget = 0, skippedNoCategory = 0;

        rows.forEach((row) => {
          const grant = findGrant(getField(row, "Grant", "grant"));
          if (!grant) { skippedNoGrant++; return; }
          const budgetTitle = String(getField(row, "Budget", "Budget Title", "budget") || "").trim().toLowerCase();
          const category = String(getField(row, "Category", "category")).trim();
          if (!category) { skippedNoCategory++; return; }
          const subcategory = String(getField(row, "Subcategory", "subcategory")).trim();

          const targetBudget = budgets.find((b) => b.grantId === grant.id && b.title.trim().toLowerCase() === budgetTitle);
          if (!targetBudget) { skippedNoBudget++; return; }

          if (!updated[targetBudget.id]) {
            updated[targetBudget.id] = { ...targetBudget, lines: targetBudget.lines.map((l) => ({ ...l, actuals: [...(l.actuals || Array(12).fill(0))] })) };
          }
          const workingBudget = updated[targetBudget.id];
          const catDef = CATEGORIES.find((c) => c.name.toLowerCase() === category.toLowerCase());
          const actuals = MONTHS.map((m) => Number(row[m] || 0) || 0);

          const existingLine = workingBudget.lines.find((l) =>
            l.category.trim().toLowerCase() === category.toLowerCase() && (l.subcategory || "").trim().toLowerCase() === subcategory.toLowerCase()
          );
          if (existingLine) {
            existingLine.actuals = actuals;
            linesUpdated++;
          } else {
            workingBudget.lines.push({
              id: uid(),
              category: catDef ? catDef.name : category,
              type: catDef ? catDef.type : (String(getField(row, "Type", "type") || "expense").toLowerCase() === "revenue" ? "revenue" : "expense"),
              subcategory, categoryCustom: !catDef, subcategoryCustom: !!subcategory,
              amounts: Array(12).fill(0), actuals,
            });
            linesCreated++;
          }
          rowsApplied++;
        });

        const changedBudgets = Object.values(updated);
        if (changedBudgets.length === 0) throw new Error("No rows matched an existing grant and budget");
        setBudgets((prev) => prev.map((b) => updated[b.id] || b));
        logActivity?.("Data", "Imported", `Imported actuals into ${changedBudgets.length} budget(s) from "${file.name}"`);
        const skippedTotal = skippedNoGrant + skippedNoBudget + skippedNoCategory;
        setActualsImportSummary(
          `Updated ${changedBudgets.length} budget${changedBudgets.length > 1 ? "s" : ""} — ${linesUpdated} line${linesUpdated === 1 ? "" : "s"} updated, ${linesCreated} new line${linesCreated === 1 ? "" : "s"} added.` +
          (skippedTotal > 0 ? ` Skipped ${skippedTotal} row${skippedTotal > 1 ? "s" : ""} (${skippedNoGrant} no matching grant, ${skippedNoBudget} no matching budget title, ${skippedNoCategory} missing category).` : "")
        );
      } catch (err) {
        setActualsImportError("Couldn't read that file, or no rows matched an existing grant + budget. The Grant and Budget columns must exactly match an existing grant and an existing budget already created for it.");
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
        <div className="flex items-center justify-between mb-1">
          <h2 className="font-display text-base" style={{ color: "#1C2624" }}>Data health check</h2>
          <button
            onClick={() => setShowHealthCheck((v) => !v)}
            className="text-xs px-3 py-1.5 rounded-md border shrink-0"
            style={{ borderColor: "#E1E5DE", color: "#1C2624" }}
          >
            {showHealthCheck ? "Hide" : "Run health check"}
          </button>
        </div>
        <p className="text-sm mb-3" style={{ color: "#5B6B66" }}>
          Scans budgets, grants, reports, and tasks for the kinds of issues that are easy to miss — orphaned links, format drift, missing approvals, overdue compliance items.
        </p>
        {showHealthCheck && (
          healthIssues.length === 0 ? (
            <div className="flex items-center gap-2 text-sm rounded-md px-3 py-2" style={{ background: "#F0F5F2", color: "#2F6F53" }}>
              <CheckCircle size={15} /> All clear — nothing found.
            </div>
          ) : (
            <div className="space-y-3">
              <div className="flex items-center gap-3 text-xs" style={{ color: "#8A8F87" }}>
                {issueCounts.error > 0 && <span style={{ color: levelColor.error }}>{issueCounts.error} needs fixing</span>}
                {issueCounts.warn > 0 && <span style={{ color: levelColor.warn }}>{issueCounts.warn} worth a look</span>}
                {issueCounts.info > 0 && <span style={{ color: levelColor.info }}>{issueCounts.info} FYI</span>}
              </div>
              <div className="rounded-md border divide-y max-h-96 overflow-y-auto" style={{ borderColor: "#E1E5DE" }}>
                {["error", "warn", "info"].flatMap((level) => healthIssues.filter((i) => i.level === level)).map((issue, i) => (
                  <div key={i} className="flex items-start gap-2 px-3 py-2 text-sm">
                    <span className="shrink-0 mt-0.5 text-[10px] font-semibold uppercase px-1.5 py-0.5 rounded" style={{ background: `${levelColor[issue.level]}1A`, color: levelColor[issue.level] }}>
                      {issue.area}
                    </span>
                    <span style={{ color: "#1C2624" }}>{issue.text}</span>
                  </div>
                ))}
              </div>
            </div>
          )
        )}
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
        <h2 className="font-display text-base mb-1" style={{ color: "#1C2624" }}>Bulk import budget actuals</h2>
        <p className="text-sm mb-3" style={{ color: "#5B6B66" }}>
          Updates the <strong>Actual</strong> figures on budgets that already exist — it never creates a new budget. Upload a .csv or .xlsx with one row per account: <strong>Grant</strong> (must match an existing grant's title or program code), <strong>Budget</strong> (must exactly match an existing budget's title already created for that grant), <strong>Category</strong>, Subcategory, Type, and Jan–Dec (actual dollar amounts). If a matching line item already exists in that budget, its Actuals are updated; if not, a new line is added with $0 Plan and the given Actuals.
        </p>
        {canEdit ? (
          <label className="inline-flex items-center gap-1.5 px-3 py-2 rounded-md text-sm border cursor-pointer" style={{ borderColor: "#E1E5DE", color: "#1C2624" }}>
            <Upload size={15} /> Choose spreadsheet
            <input type="file" accept=".csv,.xlsx,.xls" className="hidden" onChange={handleImportActuals} />
          </label>
        ) : (
          <p className="text-sm" style={{ color: "#8A8F87" }}>View-only access — importing is disabled.</p>
        )}
        {actualsImportSummary && <p className="text-sm mt-2" style={{ color: "#2F6F53" }}>{actualsImportSummary}</p>}
        {actualsImportError && <p className="text-sm mt-2" style={{ color: "#B5443A" }}>{actualsImportError}</p>}
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

function GrantFlowApp({ currentUserEmail, isAdmin, userRole, disabledModules, onSignOut } = {}) {
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
  const [paylocityProgramMap, setPaylocityProgramMap] = useState([]);
  const [paylocityLastImport, setPaylocityLastImport] = useState(null);
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
      const pm = await withTimeout(loadData("grantflow:paylocityprogrammap"));
      if (pm) setPaylocityProgramMap(pm);
    } catch (e) { /* no data yet */ }
    try {
      const pli = await withTimeout(loadData("grantflow:paylocitylastimport"));
      if (pli) setPaylocityLastImport(pli);
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
    saveKey("grantflow:paylocityprogrammap", paylocityProgramMap, "Paylocity program mapping");
  }, [paylocityProgramMap, loaded]);

  useEffect(() => {
    if (!loaded || isSyncingRef.current) return;
    saveKey("grantflow:paylocitylastimport", paylocityLastImport, "Paylocity last import");
  }, [paylocityLastImport, loaded]);

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
          <Dashboard grants={grants} budgets={budgets} reports={reports} tasks={tasks} staff={staff} invoices={invoices} goTo={goTo} costCenters={costCenters} budgetGroups={budgetGroups} />
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
            grants={grants} invoices={invoices} setInvoices={setInvoices} budgets={budgets} costCenters={costCenters}
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
            budgets={budgets} setBudgets={setBudgets}
            paylocityProgramMap={paylocityProgramMap} setPaylocityProgramMap={setPaylocityProgramMap}
            paylocityLastImport={paylocityLastImport} setPaylocityLastImport={setPaylocityLastImport}
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
          <ReportingView grants={grants} budgets={budgets} costCenters={costCenters} budgetGroups={budgetGroups} invoices={invoices} />
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

// Catches uncaught render errors anywhere in the app. Without this, a single
// thrown error (e.g. from unexpected/malformed data) unmounts the entire
// React tree and leaves a blank screen — the only recovery being a hard
// refresh. This shows a plain-language message and the actual error instead,
// so a real bug is visible and reportable rather than silently blanking out.
class ErrorBoundary extends Component {
  constructor(props) {
    super(props);
    this.state = { error: null };
  }
  static getDerivedStateFromError(error) {
    return { error };
  }
  componentDidCatch(error, info) {
    console.error("GrantFlow crashed:", error, info?.componentStack);
  }
  render() {
    if (this.state.error) {
      return (
        <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: "#0B0F0D", padding: 24 }}>
          <div style={{ maxWidth: 520, textAlign: "center", color: "#E7E9E5" }}>
            <AlertCircle size={28} style={{ color: "#B5443A", margin: "0 auto 12px" }} />
            <h1 style={{ fontSize: 16, fontWeight: 600, marginBottom: 8 }}>Something went wrong</h1>
            <p style={{ fontSize: 13, color: "#8A8F87", marginBottom: 16 }}>
              GrantFlow hit an unexpected error and couldn't continue. Your data is unaffected — everything already saved is still there. Reloading the page should fix it.
            </p>
            <button
              onClick={() => window.location.reload()}
              className="text-sm px-4 py-2 rounded-md text-white"
              style={{ background: "#1F5C6B" }}
            >
              Reload page
            </button>
            <details style={{ marginTop: 20, textAlign: "left", fontSize: 11, color: "#8A8F87" }}>
              <summary style={{ cursor: "pointer" }}>Technical details</summary>
              <pre style={{ whiteSpace: "pre-wrap", marginTop: 8 }}>{String(this.state.error?.stack || this.state.error)}</pre>
            </details>
          </div>
        </div>
      );
    }
    return this.props.children;
  }
}

export default function GrantFlow(props) {
  return (
    <ErrorBoundary>
      <GrantFlowApp {...props} />
    </ErrorBoundary>
  );
}

