import React, { useEffect, useMemo, useRef, useState } from "react";
import { isSupabaseConfigured, supabase } from "./lib/supabase";

const PRESS_ORDER = ["5.1", "6.1", "1.1", "2.1", "8", "9", "Rewind"];
const STORAGE_KEY = "labeltraxx-scheduler-v4";
const SESSION_STORAGE_KEY = "labeltraxx-scheduler-session-v1";
const SHARED_STATE_ROW_ID = "labeltraxx-shared-state";
const LOGIN_SESSION_DURATION_MS = 8 * 60 * 60 * 1000;
const BASE_TABS = ["Scheduler", "Notes", "New Request", "Open Requests", "Request History", "Pull Paper Request", "Supplies Request", "Daily Shipment", "Shipment Emails"];
const ACCESS_MODE_OPTIONS = ["edit", "view"];
const ATTACHMENT_ACCEPT =
  ".pdf,.doc,.docx,.xls,.xlsx,.csv,.txt,.rtf,.png,.jpg,.jpeg,.zip,.msg,.eml";
const PULL_PAPER_TARGETS = ["Press 5.1", "Press 6.1", "Press 2.1", "Press 1.1", "Digital"];
const ROLE_OPTIONS = ["Management", "Warehouse/Shipper", "Operator"];
const DEFAULT_SHIPMENT_METHODS = ["Skid", "FedEx", "UPS", "LTL", "Customer Pickup"];

const EMPTY_REQUEST_FORM = {
  jobNumber: "",
  customer: "",
  requestorName: "",
  description: "",
};

const EMPTY_PULL_PAPER_FORM = {
  details: "",
  target: PULL_PAPER_TARGETS[0],
};

const EMPTY_SUPPLIES_FORM = {
  details: "",
};

const EMPTY_NOTE_FORM = {
  text: "",
};

const EMPTY_MANUAL_SCHEDULE_FORM = {
  title: "",
  press: PRESS_ORDER[0],
  dayKey: "",
};

const EMPTY_LOGIN_FORM = {
  username: "",
  password: "",
};

const EMPTY_REGISTER_FORM = {
  username: "",
  password: "",
};

const EMPTY_USER_FORM = {
  username: "",
  password: "",
  accessMode: "edit",
  canManageUsers: false,
  tabs: BASE_TABS.filter((tab) => !["New Request", "Open Requests"].includes(tab)),
};

const EMPTY_SHIPMENT_FORM = {
  label: "",
  method: DEFAULT_SHIPMENT_METHODS[0],
  packageCount: "",
  packageType: "Skids",
  totalCost: "",
  billAmount: "",
  notes: "",
};

const EMPTY_EMAIL_FORM = {
  recipients: "",
};

function safeText(value) {
  return value == null ? "" : String(value).trim();
}

function parseNumber(value) {
  const cleaned = safeText(value).replace(/,/g, "");
  if (!cleaned) return 0;
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
}

function parseCurrency(value) {
  const cleaned = safeText(value).replace(/[$,]/g, "");
  if (!cleaned) return 0;
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
}

function parseDate(value) {
  const raw = safeText(value);
  if (!raw || raw === "00/00/00") return null;
  const parts = raw.split(/[\/\-]/).map((part) => part.trim());
  if (parts.length < 3) return null;
  let [monthValue, dayValue, yearValue] = parts;
  const month = Number(monthValue);
  const day = Number(dayValue);
  let year = Number(yearValue);
  if (!Number.isFinite(month) || !Number.isFinite(day) || !Number.isFinite(year)) return null;
  if (year < 100) year += 2000;
  const dt = new Date(year, month - 1, day);
  return Number.isNaN(dt.getTime()) ? null : dt;
}

function formatDate(date) {
  if (!date) return "-";
  return new Intl.DateTimeFormat("en-US", {
    month: "numeric",
    day: "numeric",
    year: "2-digit",
  }).format(date);
}

function formatShortDate(date) {
  if (!date) return "";
  return new Intl.DateTimeFormat("en-US", {
    weekday: "short",
    month: "numeric",
    day: "numeric",
  }).format(date);
}

function formatDateTime(value) {
  if (!value) return "-";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "-";
  return new Intl.DateTimeFormat("en-US", {
    month: "numeric",
    day: "numeric",
    year: "2-digit",
    hour: "numeric",
    minute: "2-digit",
  }).format(date);
}

function formatCurrency(value) {
  return new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
  }).format(value || 0);
}

function formatFileSize(bytes) {
  if (!Number.isFinite(bytes) || bytes <= 0) return "0 KB";
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}

function startOfWeek(input) {
  const date = new Date(input);
  const day = date.getDay();
  const diff = day === 0 ? -6 : 1 - day;
  date.setDate(date.getDate() + diff);
  date.setHours(0, 0, 0, 0);
  return date;
}

function addDays(date, days) {
  const copy = new Date(date);
  copy.setDate(copy.getDate() + days);
  return copy;
}

function isoDate(date) {
  if (!date) return "";
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function todayKey() {
  return isoDate(new Date());
}

function sameDay(dateValue, dayKey) {
  if (!dateValue || !dayKey) return false;
  return isoDate(new Date(dateValue)) === dayKey;
}

function comparableUsername(value) {
  return safeText(value).toLowerCase();
}

function getDefaultTabsForRole(role, isAdmin = false) {
  const normalizedRole = normalizeRole(role, isAdmin);
  if (normalizedRole === "Management") return [...BASE_TABS, "User Admin"];
  if (normalizedRole === "Warehouse/Shipper") {
    return BASE_TABS.filter((tab) => !["New Request", "Open Requests"].includes(tab));
  }
  if (normalizedRole === "Operator") return ["Scheduler"];
  return ["Scheduler"];
}

function normalizeAccessMode(mode, role, isAdmin = false) {
  const normalized = safeText(mode).toLowerCase();
  if (ACCESS_MODE_OPTIONS.includes(normalized)) return normalized;
  return normalizeRole(role, isAdmin) === "Operator" ? "view" : "edit";
}

function normalizeUserTabs(tabs, role, isAdmin = false, canManageUsers = false) {
  const next = Array.isArray(tabs)
    ? tabs
        .map((tab) => safeText(tab))
        .filter((tab) => BASE_TABS.includes(tab) || (canManageUsers && tab === "User Admin"))
    : [];
  const fallback = next.length ? Array.from(new Set(next)) : getDefaultTabsForRole(role, isAdmin);
  const filtered = canManageUsers ? fallback : fallback.filter((tab) => tab !== "User Admin");
  if (canManageUsers && !filtered.includes("User Admin")) filtered.push("User Admin");
  return filtered;
}

function normalizeRole(role, isAdmin) {
  const normalized = safeText(role);
  if (ROLE_OPTIONS.includes(normalized)) return normalized;
  return isAdmin ? "Management" : "Warehouse/Shipper";
}

function getUserRole(user) {
  if (user?.canManageUsers) return "Management";
  return normalizeRole(user?.role, user?.isAdmin);
}

function hasManagementAccess(user) {
  return !!user?.canManageUsers || getUserRole(user) === "Management";
}

function getUserAccessMode(user) {
  return normalizeAccessMode(user?.accessMode, user?.role, user?.canManageUsers || user?.isAdmin);
}

function canEdit(user) {
  return getUserAccessMode(user) === "edit";
}

function getVisibleTabs(user) {
  if (!user) return ["Scheduler"];
  return normalizeUserTabs(user.tabs, user.role, user.isAdmin, user.canManageUsers);
}

function canMoveJobs(user) {
  return canEdit(user) && canAccessTab(user, "Scheduler");
}

function canAccessTab(user, tab) {
  return getVisibleTabs(user).includes(tab);
}

function buildDefaultAdmin() {
  return {
    id: "user-admin",
    username: "Admin",
    password: "1234",
    role: "Management",
    accessMode: "edit",
    tabs: [...BASE_TABS, "User Admin"],
    canManageUsers: true,
    isAdmin: true,
    createdAt: new Date().toISOString(),
    createdBy: "system",
  };
}

function normalizeUsers(users) {
  const normalized = Array.isArray(users)
    ? users
        .map((user, index) => ({
          id: user.id || `user-${index + 1}`,
          username: safeText(user.username),
          password: safeText(user.password),
          role: normalizeRole(user.role, user.isAdmin || user.canManageUsers),
          accessMode: normalizeAccessMode(user.accessMode, user.role, user.isAdmin || user.canManageUsers),
          canManageUsers: !!user.canManageUsers || normalizeRole(user.role, user.isAdmin) === "Management",
          tabs: normalizeUserTabs(
            user.tabs,
            user.role,
            user.isAdmin || user.canManageUsers,
            !!user.canManageUsers || normalizeRole(user.role, user.isAdmin) === "Management"
          ),
          isAdmin: !!user.canManageUsers || normalizeRole(user.role, user.isAdmin) === "Management",
          createdAt: user.createdAt || new Date().toISOString(),
          createdBy: user.createdBy || "system",
        }))
        .filter((user) => user.username)
    : [];

  const hasAdmin = normalized.some((user) => comparableUsername(user.username) === "admin");
  if (!hasAdmin) normalized.unshift(buildDefaultAdmin());
  return normalized;
}

function normalizeJobs(jobs) {
  return Array.isArray(jobs)
    ? jobs.map((job) => ({
        ...job,
        press: normalizePressValue(job.press),
        shipByDate: job.shipByDate ? new Date(job.shipByDate) : null,
        entryDate: job.entryDate ? new Date(job.entryDate) : null,
        dueOnSiteDate: job.dueOnSiteDate ? new Date(job.dueOnSiteDate) : null,
      }))
    : [];
}

function normalizeRequests(requests) {
  return Array.isArray(requests)
    ? requests.map((request) => ({
        ...request,
        attachments: Array.isArray(request.attachments) ? request.attachments : [],
        createdByAccount: request.createdByAccount || "",
        completedByAccount: request.completedByAccount || "",
      }))
    : [];
}

function normalizePullPaperRequests(requests) {
  return Array.isArray(requests) ? requests : [];
}

function normalizeNotes(notes) {
  return Array.isArray(notes)
    ? notes
        .map((note, index) => ({
          id: note.id || `note-${index + 1}`,
          ownerUsername: safeText(note.ownerUsername || note.username),
          text: safeText(note.text),
          completed: !!note.completed,
          createdAt: note.createdAt || new Date().toISOString(),
          completedAt: note.completedAt || "",
        }))
        .filter((note) => note.ownerUsername && note.text)
    : [];
}

function normalizeRegistrationRequests(requests) {
  return Array.isArray(requests)
    ? requests
        .map((request, index) => ({
          id: request.id || `registration-${index + 1}`,
          username: safeText(request.username),
          password: safeText(request.password),
          status: safeText(request.status) || "pending",
          createdAt: request.createdAt || new Date().toISOString(),
          createdBy: request.createdBy || safeText(request.username),
          approvedAt: request.approvedAt || "",
          approvedBy: request.approvedBy || "",
          deniedAt: request.deniedAt || "",
          deniedBy: request.deniedBy || "",
        }))
        .filter((request) => request.username && request.password)
    : [];
}

function normalizeSuppliesRequests(requests) {
  return Array.isArray(requests)
    ? requests.map((request) => ({
        ...request,
        attachments: Array.isArray(request.attachments) ? request.attachments : [],
      }))
    : [];
}

function normalizeAssignments(assignments) {
  return Array.isArray(assignments)
    ? assignments
        .filter((assignment) => assignment.kind !== "rewind")
        .map((assignment, index) => ({
          id: assignment.id || `assignment-${index + 1}`,
          ...assignment,
          manualTitle: safeText(assignment.manualTitle),
          kind: safeText(assignment.kind) || "press",
          status: safeText(assignment.status) || "scheduled",
        }))
    : [];
}

function normalizeShipmentGroups(groups) {
  return Array.isArray(groups)
    ? groups.map((group) => ({
        ...group,
        attachments: Array.isArray(group.attachments) ? group.attachments : [],
        billAmount: parseCurrency(group.billAmount),
      }))
    : [];
}

function normalizeShipmentMethods(methods) {
  const next = Array.isArray(methods)
    ? methods.map((method) => safeText(method)).filter(Boolean)
    : [];
  return next.length ? Array.from(new Set(next)) : [...DEFAULT_SHIPMENT_METHODS];
}

function normalizeShipmentEmailLogs(logs) {
  return Array.isArray(logs)
    ? logs
        .map((log, index) => ({
          id: log.id || `email-${index + 1}`,
          shipDate: safeText(log.shipDate),
          recipients: safeText(log.recipients),
          subject: safeText(log.subject),
          body: safeText(log.body),
          jobCount: parseNumber(log.jobCount),
          totalCost: parseCurrency(log.totalCost),
          totalBill: parseCurrency(log.totalBill),
          methods: Array.isArray(log.methods) ? log.methods.map((method) => safeText(method)).filter(Boolean) : [],
          createdAt: log.createdAt || new Date().toISOString(),
          createdBy: safeText(log.createdBy),
        }))
        .filter((log) => log.shipDate)
    : [];
}

function defaultSharedSnapshot() {
  return {
    jobs: [],
    assignments: [],
    requests: [],
    pullPaperRequests: [],
    notes: [],
    registrationRequests: [],
    suppliesRequests: [],
    shipmentGroups: [],
    shipmentEmailLogs: [],
    shipmentMethods: [...DEFAULT_SHIPMENT_METHODS],
    users: [buildDefaultAdmin()],
    weekStart: startOfWeek(new Date()).toISOString(),
  };
}

function normalizeSharedSnapshot(snapshot) {
  const source = snapshot && typeof snapshot === "object" ? snapshot : {};
  return {
    jobs: normalizeJobs(source.jobs),
    assignments: normalizeAssignments(source.assignments),
    requests: normalizeRequests(source.requests),
    pullPaperRequests: normalizePullPaperRequests(source.pullPaperRequests),
    notes: normalizeNotes(source.notes),
    registrationRequests: normalizeRegistrationRequests(source.registrationRequests),
    suppliesRequests: normalizeSuppliesRequests(source.suppliesRequests),
    shipmentGroups: normalizeShipmentGroups(source.shipmentGroups),
    shipmentEmailLogs: normalizeShipmentEmailLogs(source.shipmentEmailLogs),
    shipmentMethods: normalizeShipmentMethods(source.shipmentMethods),
    users: normalizeUsers(source.users),
    weekStart: source.weekStart ? new Date(source.weekStart) : startOfWeek(new Date()),
  };
}

function buildSharedSnapshot(state) {
  return {
    jobs: state.jobs,
    assignments: state.assignments,
    requests: state.requests,
    pullPaperRequests: state.pullPaperRequests,
    notes: state.notes,
    registrationRequests: state.registrationRequests,
    suppliesRequests: state.suppliesRequests,
    shipmentGroups: state.shipmentGroups,
    shipmentEmailLogs: state.shipmentEmailLogs,
    shipmentMethods: state.shipmentMethods,
    users: state.users,
    weekStart: state.weekStart.toISOString(),
  };
}

function readFileAsAttachment(file, uploadedBy) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () =>
      resolve({
        id: makeId("att"),
        name: file.name,
        size: file.size,
        type: file.type || "application/octet-stream",
        uploadedAt: new Date().toISOString(),
        uploadedBy,
        dataUrl: String(reader.result || ""),
      });
    reader.onerror = () => reject(reader.error || new Error("Failed to read file."));
    reader.readAsDataURL(file);
  });
}

async function buildAttachments(fileList, uploadedBy) {
  const files = Array.from(fileList || []);
  if (!files.length) return [];
  return Promise.all(files.map((file) => readFileAsAttachment(file, uploadedBy)));
}

function looksLikeJobRow(parts) {
  if (parts.length < 2) return false;
  const press = safeText(parts[0]);
  const number = safeText(parts[1]);
  return /^\d+(?:\.\d+)?$/.test(press) && /^\d+$/.test(number);
}

function normalizeJobStatus(status) {
  const value = safeText(status).toLowerCase();
  if (!value) return "open";
  if (value === "done" || value === "closed" || value === "finished" || value === "complete") return "closed";
  return "open";
}

function normalizePressValue(value) {
  const normalized = safeText(value);
  if (!normalized) return "";
  if (normalized === "1") return "1.1";
  if (normalized === "2") return "2.1";
  return normalized;
}

function parseLabelTraxxText(text) {
  const normalized = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  const lines = normalized.split("\n");
  const headerLine = lines.find(
    (line) => line.includes("Press\t") && line.includes("Number\t") && line.includes("CustomerName")
  );
  if (!headerLine) return [];

  const startIndex = lines.indexOf(headerLine);
  const headers = headerLine.split("\t");
  const headerCount = headers.length;
  const rows = [];
  let current = null;

  for (let index = startIndex + 1; index < lines.length; index += 1) {
    const line = lines[index];
    if (!line.trim()) {
      if (current) current[headerCount - 1] = `${current[headerCount - 1] || ""}\n`;
      continue;
    }

    const parts = line.split("\t");
    if (looksLikeJobRow(parts)) {
      if (current) rows.push(current);
      current = Array.from({ length: headerCount }, (_, itemIndex) => parts[itemIndex] ?? "");
      if (parts.length > headerCount) {
        current[headerCount - 1] = parts.slice(headerCount - 1).join("\t");
      }
      continue;
    }

    if (current) {
      current[headerCount - 1] = `${current[headerCount - 1] || ""}${current[headerCount - 1] ? "\n" : ""}${line}`;
    }
  }

  if (current) rows.push(current);

  return rows
    .map((row) => {
      const record = Object.fromEntries(headers.map((header, index) => [header, safeText(row[index])]));
      const stockParts = [safeText(record.StockNum2), safeText(record.StockNum1)].filter(Boolean);
      const importedStatus = safeText(record.Status) || safeText(record.TicketStatus) || "Open";
      return {
        id: safeText(record.Number),
        press: normalizePressValue(record.Press),
        number: safeText(record.Number),
        customerName: safeText(record.CustomerName),
        generalDescr: safeText(record.GeneralDescr),
        custPoNum: safeText(record.CustPONum),
        priority: safeText(record.Priority),
        shipByDate: parseDate(record.Ship_by_Date),
        entryDate: parseDate(record.EntryDate),
        dueOnSiteDate: parseDate(record.Due_on_Site_Date),
        stockNum2: safeText(record.StockNum2),
        stockNum1: safeText(record.StockNum1),
        stockDisplay: stockParts.join(" / "),
        ticketStatus: importedStatus,
        normalizedStatus: normalizeJobStatus(importedStatus),
        mainTool: safeText(record.MainTool),
        toolNo2: safeText(record.ToolNo2),
        ticQuantity: parseNumber(record.TicQuantity),
        estFootage: parseNumber(record.EstFootage),
        estPressTime: parseNumber(record.EstPressTime),
        notes: safeText(record.Notes),
        raw: record,
      };
    })
    .filter((job) => job.number);
}

function priorityTone(priority) {
  const value = safeText(priority).toLowerCase();
  if (value.includes("high")) return "bg-amber-100 text-amber-900 border-amber-300";
  if (value.includes("release")) return "bg-stone-200 text-stone-800 border-stone-300";
  if (value.includes("inventory")) return "bg-lime-100 text-lime-900 border-lime-300";
  if (value.includes("digital")) return "bg-slate-200 text-slate-800 border-slate-300";
  if (value.includes("test")) return "bg-neutral-200 text-neutral-800 border-neutral-300";
  return "bg-emerald-100 text-emerald-900 border-emerald-300";
}

function statusTone(status) {
  const value = safeText(status).toLowerCase();
  if (value === "finished") return "bg-stone-900 text-stone-50";
  if (value === "ship") return "bg-emerald-800 text-white";
  if (value === "scheduled") return "bg-stone-700 text-stone-50";
  if (value === "done") return "bg-emerald-100 text-emerald-900";
  if (value === "note") return "bg-amber-100 text-amber-900";
  if (value === "open") return "bg-stone-200 text-stone-800";
  return "bg-stone-100 text-stone-700";
}

function syncTone(status) {
  const value = safeText(status).toLowerCase();
  if (value.includes("live")) return "bg-emerald-100 text-emerald-900";
  if (value.includes("saving") || value.includes("connecting")) return "bg-amber-100 text-amber-900";
  if (value.includes("error")) return "bg-rose-100 text-rose-900";
  return "bg-stone-100 text-stone-700";
}

function buildWeekColumns(weekStart) {
  return Array.from({ length: 5 }, (_, index) => {
    const date = addDays(weekStart, index);
    return {
      key: isoDate(date),
      date,
      label: new Intl.DateTimeFormat("en-US", { weekday: "long" }).format(date),
    };
  });
}

function csvEscape(value) {
  const text = safeText(value).replace(/"/g, '""');
  return `"${text}"`;
}

function makeId(prefix) {
  return `${prefix}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

function makeDragPayload(payload) {
  return JSON.stringify(payload);
}

function parseDragPayload(raw) {
  if (!raw) return null;
  try {
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === "object" ? parsed : null;
  } catch {
    return null;
  }
}

function dateSortValue(value) {
  if (!value) return 0;
  const time = new Date(value).getTime();
  return Number.isFinite(time) ? time : 0;
}

function downloadFile(name, content, type = "text/plain;charset=utf-8") {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = name;
  link.click();
  URL.revokeObjectURL(url);
}

function snapshotJobForShipment(job, finishMeta) {
  return {
    id: job.id,
    number: job.number,
    customerName: job.customerName,
    generalDescr: job.generalDescr,
    press: job.press,
    ticQuantity: job.ticQuantity,
    shipByDate: job.shipByDate ? job.shipByDate.toISOString() : null,
    finishedAt: finishMeta?.finishedAt || null,
    finishedBy: finishMeta?.finishedBy || "",
  };
}

function effectiveFinishedShipDate(finishMeta) {
  if (!finishMeta) return "";
  if (finishMeta.shipDate) return finishMeta.shipDate;
  if (finishMeta.dayKey) return finishMeta.dayKey;
  return finishMeta.finishedAt ? isoDate(new Date(finishMeta.finishedAt)) : "";
}

function getShipmentItems(group, jobMap, finishedMetaByJobId) {
  if (Array.isArray(group.jobItems) && group.jobItems.length) return group.jobItems;
  return (group.jobIds || [])
    .map((jobId) => {
      const job = jobMap.get(jobId);
      if (!job) return null;
      return snapshotJobForShipment(job, finishedMetaByJobId.get(jobId));
    })
    .filter(Boolean);
}

function deriveVisibleJobState(jobId, activePressJobIds, finishedJobIds) {
  if (finishedJobIds.has(jobId)) return "finished";
  if (activePressJobIds.has(jobId)) return "scheduled";
  return "open";
}

export default function App() {
  const [isReady, setIsReady] = useState(false);
  const [jobs, setJobs] = useState([]);
  const [assignments, setAssignments] = useState([]);
  const [requests, setRequests] = useState([]);
  const [pullPaperRequests, setPullPaperRequests] = useState([]);
  const [notes, setNotes] = useState([]);
  const [registrationRequests, setRegistrationRequests] = useState([]);
  const [suppliesRequests, setSuppliesRequests] = useState([]);
  const [shipmentGroups, setShipmentGroups] = useState([]);
  const [shipmentEmailLogs, setShipmentEmailLogs] = useState([]);
  const [shipmentMethods, setShipmentMethods] = useState([...DEFAULT_SHIPMENT_METHODS]);
  const [users, setUsers] = useState([buildDefaultAdmin()]);
  const [currentUsername, setCurrentUsername] = useState("");
  const [weekStart, setWeekStart] = useState(startOfWeek(new Date()));
  const [pasteText, setPasteText] = useState("");
  const [search, setSearch] = useState("");
  const [unscheduledSearch, setUnscheduledSearch] = useState("");
  const [queueStatusFilter, setQueueStatusFilter] = useState("All");
  const [queuePressFilter, setQueuePressFilter] = useState("All");
  const [selectedJobId, setSelectedJobId] = useState(null);
  const [activeTab, setActiveTab] = useState("Scheduler");
  const [requestForm, setRequestForm] = useState(EMPTY_REQUEST_FORM);
  const [pullPaperForm, setPullPaperForm] = useState(EMPTY_PULL_PAPER_FORM);
  const [suppliesForm, setSuppliesForm] = useState(EMPTY_SUPPLIES_FORM);
  const [noteForm, setNoteForm] = useState(EMPTY_NOTE_FORM);
  const [requestDraftAttachments, setRequestDraftAttachments] = useState([]);
  const [suppliesDraftAttachments, setSuppliesDraftAttachments] = useState([]);
  const [selectedShipDate, setSelectedShipDate] = useState(todayKey());
  const [shipDateDraft, setShipDateDraft] = useState(todayKey());
  const [selectedShipmentJobs, setSelectedShipmentJobs] = useState([]);
  const [selectedShipQueueJobs, setSelectedShipQueueJobs] = useState([]);
  const [shipmentForm, setShipmentForm] = useState({ ...EMPTY_SHIPMENT_FORM, shipDate: todayKey() });
  const [shipmentDraftAttachments, setShipmentDraftAttachments] = useState([]);
  const [newShipmentMethod, setNewShipmentMethod] = useState("");
  const [shipmentEmailForm, setShipmentEmailForm] = useState(EMPTY_EMAIL_FORM);
  const [loginForm, setLoginForm] = useState(EMPTY_LOGIN_FORM);
  const [loginError, setLoginError] = useState("");
  const [authView, setAuthView] = useState("login");
  const [registerForm, setRegisterForm] = useState(EMPTY_REGISTER_FORM);
  const [registerError, setRegisterError] = useState("");
  const [registerSuccess, setRegisterSuccess] = useState("");
  const [sessionExpiresAt, setSessionExpiresAt] = useState("");
  const [userForm, setUserForm] = useState(EMPTY_USER_FORM);
  const [userPasswordDrafts, setUserPasswordDrafts] = useState({});
  const [requestHistoryFilterDate, setRequestHistoryFilterDate] = useState("");
  const [shipmentHistoryFilterDate, setShipmentHistoryFilterDate] = useState("");
  const [shipmentEmailHistoryFilterDate, setShipmentEmailHistoryFilterDate] = useState("");
  const [queueCategoryFilter, setQueueCategoryFilter] = useState("All");
  const [schedulePressFilter, setSchedulePressFilter] = useState("All");
  const [condensedSchedule, setCondensedSchedule] = useState(true);
  const [manualScheduleForm, setManualScheduleForm] = useState(EMPTY_MANUAL_SCHEDULE_FORM);
  const [locationSearch, setLocationSearch] = useState("");
  const [syncStatus, setSyncStatus] = useState(isSupabaseConfigured ? "Connecting..." : "Local only");
  const [lastSyncAt, setLastSyncAt] = useState("");
  const jobDetailsRef = useRef(null);
  const lastSharedSnapshotRef = useRef("");
  const saveTimerRef = useRef(null);

  useEffect(() => {
    let isCancelled = false;

    function applySharedState(snapshot) {
      const normalized = normalizeSharedSnapshot(snapshot);
      setJobs(normalized.jobs);
      setAssignments(normalized.assignments);
      setRequests(normalized.requests);
      setPullPaperRequests(normalized.pullPaperRequests);
      setNotes(normalized.notes);
      setRegistrationRequests(normalized.registrationRequests);
      setSuppliesRequests(normalized.suppliesRequests);
      setShipmentGroups(normalized.shipmentGroups);
      setShipmentEmailLogs(normalized.shipmentEmailLogs);
      setShipmentMethods(normalized.shipmentMethods);
      setUsers(normalized.users);
      setWeekStart(normalized.weekStart);
      return normalized;
    }

    async function hydrate() {
      let saved = {};
      try {
        saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || "{}");
        const session = JSON.parse(localStorage.getItem(SESSION_STORAGE_KEY) || "{}");
        let sharedSnapshot = saved;

        if (isSupabaseConfigured && supabase) {
          setSyncStatus("Connecting...");
          const { data, error } = await supabase
            .from("app_state")
            .select("payload, updated_at")
            .eq("id", SHARED_STATE_ROW_ID)
            .maybeSingle();

          if (error) throw error;

          if (data?.payload) {
            sharedSnapshot = data.payload;
            if (!isCancelled) {
              setSyncStatus("Live sync");
              setLastSyncAt(data.updated_at || new Date().toISOString());
            }
          } else {
            const seedSnapshot =
              Object.keys(saved || {}).length > 0 ? buildSharedSnapshot(normalizeSharedSnapshot(saved)) : defaultSharedSnapshot();
            const { error: seedError } = await supabase.from("app_state").upsert({
              id: SHARED_STATE_ROW_ID,
              payload: seedSnapshot,
              updated_by: safeText(session.currentUsername) || "system",
            });
            if (seedError) throw seedError;
            sharedSnapshot = seedSnapshot;
            if (!isCancelled) {
              setSyncStatus("Live sync");
              setLastSyncAt(new Date().toISOString());
            }
          }
        }

        if (isCancelled) return;

        const normalized = applySharedState(sharedSnapshot);
        const digest = JSON.stringify(buildSharedSnapshot(normalized));
        lastSharedSnapshotRef.current = digest;
        const sessionUsername = safeText(session.currentUsername);
        const storedExpiresAt = safeText(session.expiresAt);
        const expiresAtValue = storedExpiresAt ? new Date(storedExpiresAt).getTime() : 0;
        const isSessionActive = sessionUsername && Number.isFinite(expiresAtValue) && expiresAtValue > Date.now();
        if (isSessionActive) {
          const match = normalized.users.find(
            (user) => comparableUsername(user.username) === comparableUsername(sessionUsername)
          );
          if (match) {
            setCurrentUsername(match.username);
            setSessionExpiresAt(storedExpiresAt);
          }
        } else if (sessionUsername || storedExpiresAt) {
          localStorage.setItem(
            SESSION_STORAGE_KEY,
            JSON.stringify({
              currentUsername: "",
              expiresAt: "",
            })
          );
        }
      } catch (error) {
        console.error("Failed to load shared scheduler state.", error);
        const fallbackSource = Object.keys(saved || {}).length ? saved : defaultSharedSnapshot();
        const fallback = normalizeSharedSnapshot(fallbackSource);
        applySharedState(fallback);
        lastSharedSnapshotRef.current = JSON.stringify(buildSharedSnapshot(fallback));
        setSyncStatus(isSupabaseConfigured ? "Sync error" : "Local only");
      } finally {
        if (!isCancelled) setIsReady(true);
      }
    }

    hydrate();

    return () => {
      isCancelled = true;
    };
  }, []);

  useEffect(() => {
    if (!isReady) return;
    const sharedSnapshot = buildSharedSnapshot({
      jobs,
      assignments,
      requests,
      pullPaperRequests,
      notes,
      registrationRequests,
      suppliesRequests,
      shipmentGroups,
      shipmentEmailLogs,
      shipmentMethods,
      users,
      weekStart,
    });
    const digest = JSON.stringify(sharedSnapshot);

    localStorage.setItem(STORAGE_KEY, digest);

    if (digest === lastSharedSnapshotRef.current) return;
    lastSharedSnapshotRef.current = digest;

    if (!isSupabaseConfigured || !supabase) {
      setSyncStatus("Local only");
      return;
    }

    setSyncStatus("Saving...");
    window.clearTimeout(saveTimerRef.current);
    saveTimerRef.current = window.setTimeout(async () => {
      const { error } = await supabase.from("app_state").upsert({
        id: SHARED_STATE_ROW_ID,
        payload: sharedSnapshot,
        updated_by: currentUsername || "system",
      });

      if (error) {
        console.error("Failed to save shared scheduler state.", error);
        setSyncStatus("Sync error");
        return;
      }

      setSyncStatus("Live sync");
      setLastSyncAt(new Date().toISOString());
    }, 300);

    return () => {
      window.clearTimeout(saveTimerRef.current);
    };
  }, [assignments, currentUsername, isReady, jobs, notes, pullPaperRequests, registrationRequests, requests, shipmentEmailLogs, shipmentGroups, shipmentMethods, suppliesRequests, users, weekStart]);

  useEffect(() => {
    localStorage.setItem(
      SESSION_STORAGE_KEY,
      JSON.stringify({
        currentUsername,
        expiresAt: currentUsername ? sessionExpiresAt : "",
      })
    );
  }, [currentUsername, sessionExpiresAt]);

  const currentUser = useMemo(
    () =>
      users.find((user) => comparableUsername(user.username) === comparableUsername(currentUsername)) || null,
    [currentUsername, users]
  );

  const currentUserRole = useMemo(() => getUserRole(currentUser), [currentUser]);
  const currentUserAccessMode = useMemo(() => getUserAccessMode(currentUser), [currentUser]);
  const userCanEdit = useMemo(() => canEdit(currentUser), [currentUser]);
  const userCanManageUsers = useMemo(() => hasManagementAccess(currentUser), [currentUser]);
  const userCanMoveJobs = useMemo(() => canMoveJobs(currentUser), [currentUser]);

  useEffect(() => {
    if (currentUser) return;
    if (!currentUsername) return;
    setCurrentUsername("");
    setSessionExpiresAt("");
  }, [currentUser, currentUsername]);

  useEffect(() => {
    if (!currentUsername || !sessionExpiresAt) return;
    const expiresAtValue = new Date(sessionExpiresAt).getTime();
    if (!Number.isFinite(expiresAtValue)) {
      setCurrentUsername("");
      setSessionExpiresAt("");
      return;
    }
    const timeoutMs = expiresAtValue - Date.now();
    if (timeoutMs <= 0) {
      setCurrentUsername("");
      setSessionExpiresAt("");
      return;
    }
    const timer = window.setTimeout(() => {
      setCurrentUsername("");
      setSessionExpiresAt("");
    }, timeoutMs);
    return () => window.clearTimeout(timer);
  }, [currentUsername, sessionExpiresAt]);

  useEffect(() => {
    if (!currentUser) return;
    if (canAccessTab(currentUser, activeTab)) return;
    setActiveTab("Scheduler");
  }, [activeTab, currentUser]);

  useEffect(() => {
    if (!isReady || !isSupabaseConfigured || !supabase) return undefined;

    const channel = supabase
      .channel("labeltraxx-shared-state")
      .on(
        "postgres_changes",
        {
          event: "*",
          schema: "public",
          table: "app_state",
          filter: `id=eq.${SHARED_STATE_ROW_ID}`,
        },
        (payload) => {
          const nextPayload = payload.new?.payload;
          if (!nextPayload) return;
          const digest = JSON.stringify(nextPayload);
          if (digest === lastSharedSnapshotRef.current) {
            setSyncStatus("Live sync");
            setLastSyncAt(payload.new?.updated_at || new Date().toISOString());
            return;
          }

          const normalized = normalizeSharedSnapshot(nextPayload);
          lastSharedSnapshotRef.current = JSON.stringify(buildSharedSnapshot(normalized));
          setJobs(normalized.jobs);
          setAssignments(normalized.assignments);
          setRequests(normalized.requests);
          setPullPaperRequests(normalized.pullPaperRequests);
          setNotes(normalized.notes);
          setRegistrationRequests(normalized.registrationRequests);
          setSuppliesRequests(normalized.suppliesRequests);
          setShipmentGroups(normalized.shipmentGroups);
          setShipmentEmailLogs(normalized.shipmentEmailLogs);
          setShipmentMethods(normalized.shipmentMethods);
          setUsers(normalized.users);
          setWeekStart(normalized.weekStart);
          setSyncStatus("Live sync");
          setLastSyncAt(payload.new?.updated_at || new Date().toISOString());
        }
      )
      .subscribe((status) => {
        if (status === "SUBSCRIBED") {
          setSyncStatus("Live sync");
          return;
        }
        if (status === "CHANNEL_ERROR" || status === "TIMED_OUT") {
          setSyncStatus("Sync error");
        }
      });

    return () => {
      supabase.removeChannel(channel);
    };
  }, [isReady]);

  useEffect(() => {
    if (!jobs.some((job) => job.id === selectedJobId)) setSelectedJobId(null);
  }, [jobs, selectedJobId]);

  useEffect(() => {
    setShipmentForm((current) => ({ ...current, shipDate: selectedShipDate }));
    setShipDateDraft(selectedShipDate);
    setSelectedShipmentJobs([]);
  }, [selectedShipDate]);

  useEffect(() => {
    if (!shipmentMethods.length) return;
    setShipmentForm((current) =>
      shipmentMethods.includes(current.method) ? current : { ...current, method: shipmentMethods[0] }
    );
  }, [shipmentMethods]);

  useEffect(() => {
    if (!weekColumns.length) return;
    setManualScheduleForm((current) =>
      current.dayKey && weekColumns.some((day) => day.key === current.dayKey)
        ? current
        : { ...current, dayKey: weekColumns[0].key, press: PRESS_ORDER.includes(current.press) ? current.press : PRESS_ORDER[0] }
    );
  }, [weekColumns]);

  const tabs = useMemo(
    () => [...BASE_TABS, "User Admin"].filter((tab) => canAccessTab(currentUser, tab)),
    [currentUser]
  );

  const weekColumns = useMemo(() => buildWeekColumns(weekStart), [weekStart]);
  const weekKeys = useMemo(() => new Set(weekColumns.map((column) => column.key)), [weekColumns]);
  const jobMap = useMemo(() => new Map(jobs.map((job) => [job.id, job])), [jobs]);

  const filteredJobs = useMemo(() => {
    return jobs.filter((job) => {
      const haystack = `${job.number} ${job.customerName} ${job.generalDescr} ${job.notes}`.toLowerCase();
      const matchesSearch = !search.trim() || haystack.includes(search.toLowerCase());
      return matchesSearch;
    });
  }, [jobs, search]);

  const filteredJobIds = useMemo(() => new Set(filteredJobs.map((job) => job.id)), [filteredJobs]);

  const importedSummary = useMemo(() => {
    const openCount = jobs.filter((job) => job.normalizedStatus === "open").length;
    const closedCount = jobs.filter((job) => job.normalizedStatus === "closed").length;
    return { openCount, closedCount };
  }, [jobs]);

  const userFinishedAssignments = useMemo(
    () => assignments.filter((assignment) => assignment.kind === "press" && assignment.status === "finished"),
    [assignments]
  );

  const finishedMetaByJobId = useMemo(() => {
    const map = new Map();
    userFinishedAssignments.forEach((assignment) => {
      const existing = map.get(assignment.jobId);
      if (!existing || dateSortValue(assignment.finishedAt) > dateSortValue(existing.finishedAt)) {
        map.set(assignment.jobId, {
          finishedAt: assignment.finishedAt || null,
          finishedBy: assignment.finishedBy || "",
          dayKey: assignment.finishedAt ? isoDate(new Date(assignment.finishedAt)) : assignment.dayKey,
          shipDate: assignment.shipDate || (assignment.finishedAt ? isoDate(new Date(assignment.finishedAt)) : assignment.dayKey),
        });
      }
    });
    return map;
  }, [userFinishedAssignments]);

  const userFinishedJobIds = useMemo(
    () => new Set(Array.from(finishedMetaByJobId.keys())),
    [finishedMetaByJobId]
  );

  const activePressAssignments = useMemo(
    () =>
      assignments.filter(
        (assignment) =>
          assignment.kind === "press" &&
          assignment.status !== "finished" &&
          filteredJobIds.has(assignment.jobId)
      ),
    [assignments, filteredJobIds]
  );

  const activePressJobIds = useMemo(
    () => new Set(activePressAssignments.map((assignment) => assignment.jobId)),
    [activePressAssignments]
  );

  const unscheduledJobs = useMemo(() => {
    return filteredJobs
      .filter((job) => !userFinishedJobIds.has(job.id))
      .filter((job) => {
        if (queueStatusFilter === "All") return true;
        if (queueStatusFilter === "Open") return job.normalizedStatus === "open";
        return job.normalizedStatus === "closed";
      })
      .filter((job) => {
        if (queuePressFilter === "All") return true;
        return safeText(job.press) === queuePressFilter;
      })
      .filter((job) => {
        if (queueCategoryFilter === "All") return true;
        return safeText(job.priority) === queueCategoryFilter;
      })
      .filter((job) => {
        const haystack = `${job.number} ${job.customerName} ${job.generalDescr}`.toLowerCase();
        return !unscheduledSearch.trim() || haystack.includes(unscheduledSearch.toLowerCase());
      })
      .sort((left, right) => {
        const leftDate = left.shipByDate ? left.shipByDate.getTime() : Number.MAX_SAFE_INTEGER;
        const rightDate = right.shipByDate ? right.shipByDate.getTime() : Number.MAX_SAFE_INTEGER;
        if (leftDate !== rightDate) return leftDate - rightDate;
        return right.estPressTime - left.estPressTime;
      });
  }, [filteredJobs, queueCategoryFilter, queuePressFilter, queueStatusFilter, unscheduledSearch, userFinishedJobIds]);

  const allUserFinishedJobs = useMemo(() => {
    return jobs
      .filter((job) => userFinishedJobIds.has(job.id))
      .map((job) => ({
        ...job,
        finishMeta: finishedMetaByJobId.get(job.id) || null,
      }))
      .sort((left, right) => dateSortValue(right.finishMeta?.finishedAt) - dateSortValue(left.finishMeta?.finishedAt));
  }, [finishedMetaByJobId, jobs, userFinishedJobIds]);

  const doneJobs = useMemo(
    () => allUserFinishedJobs.filter((job) => filteredJobIds.has(job.id)),
    [allUserFinishedJobs, filteredJobIds]
  );

  const board = useMemo(() => {
    const map = {};
    weekColumns.forEach((day) => {
      map[day.key] = {};
      PRESS_ORDER.forEach((press) => {
        map[day.key][press] = [];
      });
    });

    assignments.forEach((assignment) => {
      if (!map[assignment.dayKey]) return;
      if (assignment.kind === "manual") {
        map[assignment.dayKey][assignment.press].push({ assignment, job: null });
        return;
      }
      if (!filteredJobIds.has(assignment.jobId)) return;
      const job = jobMap.get(assignment.jobId);
      if (!job) return;
      map[assignment.dayKey][assignment.press].push({ assignment, job });
    });

    Object.values(map).forEach((pressMap) => {
      Object.values(pressMap).forEach((laneJobs) => {
        laneJobs.sort((left, right) => {
          if (!left.job && !right.job) return left.assignment.manualTitle.localeCompare(right.assignment.manualTitle);
          if (!left.job) return -1;
          if (!right.job) return 1;
          return right.job.estPressTime - left.job.estPressTime;
        });
      });
    });

    return map;
  }, [assignments, filteredJobIds, jobMap, weekColumns]);

  const openRequests = useMemo(
    () =>
      requests
        .filter((request) => request.status === "open")
        .sort((left, right) => dateSortValue(right.createdAt) - dateSortValue(left.createdAt)),
    [requests]
  );

  const requestHistory = useMemo(
    () =>
      requests
        .filter((request) => request.status === "done")
        .filter((request) => !requestHistoryFilterDate || sameDay(request.completedAt, requestHistoryFilterDate))
        .sort((left, right) => dateSortValue(right.completedAt) - dateSortValue(left.completedAt)),
    [requestHistoryFilterDate, requests]
  );

  const pendingRegistrationRequests = useMemo(
    () =>
      registrationRequests
        .filter((request) => request.status === "pending")
        .sort((left, right) => dateSortValue(right.createdAt) - dateSortValue(left.createdAt)),
    [registrationRequests]
  );

  const openPullPaperRequests = useMemo(
    () =>
      pullPaperRequests
        .filter((request) => request.status === "open")
        .sort((left, right) => dateSortValue(right.createdAt) - dateSortValue(left.createdAt)),
    [pullPaperRequests]
  );

  const completedPullPaperRequests = useMemo(
    () =>
      pullPaperRequests
        .filter((request) => request.status === "done")
        .sort((left, right) => dateSortValue(right.completedAt) - dateSortValue(left.completedAt)),
    [pullPaperRequests]
  );

  const openSuppliesRequests = useMemo(
    () =>
      suppliesRequests
        .filter((request) => request.status === "open")
        .sort((left, right) => dateSortValue(right.createdAt) - dateSortValue(left.createdAt)),
    [suppliesRequests]
  );

  const completedSuppliesRequests = useMemo(
    () =>
      suppliesRequests
        .filter((request) => request.status === "done")
        .sort((left, right) => dateSortValue(right.completedAt) - dateSortValue(left.completedAt)),
    [suppliesRequests]
  );

  const userNotes = useMemo(
    () =>
      notes
        .filter((note) => comparableUsername(note.ownerUsername) === comparableUsername(currentUser?.username))
        .sort((left, right) => {
          if (left.completed !== right.completed) return left.completed ? 1 : -1;
          return dateSortValue(right.createdAt) - dateSortValue(left.createdAt);
        }),
    [currentUser?.username, notes]
  );

  const assignedShipmentJobIds = useMemo(() => {
    const ids = new Set();
    shipmentGroups.forEach((group) => {
      getShipmentItems(group, jobMap, finishedMetaByJobId).forEach((item) => ids.add(item.id));
    });
    return ids;
  }, [finishedMetaByJobId, jobMap, shipmentGroups]);

  const unassignedFinishedJobs = useMemo(() => {
    return allUserFinishedJobs
      .filter((job) => !assignedShipmentJobIds.has(job.id))
      .sort((left, right) => dateSortValue(right.finishMeta?.finishedAt) - dateSortValue(left.finishMeta?.finishedAt));
  }, [allUserFinishedJobs, assignedShipmentJobIds]);

  const readyToShipJobs = useMemo(() => {
    return unassignedFinishedJobs
      .filter((job) => effectiveFinishedShipDate(job.finishMeta) === selectedShipDate)
      .sort((left, right) => dateSortValue(right.finishMeta?.finishedAt) - dateSortValue(left.finishMeta?.finishedAt));
  }, [selectedShipDate, unassignedFinishedJobs]);

  const shipmentGroupsForDay = useMemo(
    () =>
      shipmentGroups
        .filter((group) => group.shipDate === selectedShipDate)
        .sort((left, right) => dateSortValue(right.createdAt) - dateSortValue(left.createdAt)),
    [selectedShipDate, shipmentGroups]
  );

  const shipmentHistoryDays = useMemo(() => {
    const grouped = new Map();
    shipmentGroups.forEach((group) => {
      if (shipmentHistoryFilterDate && group.shipDate !== shipmentHistoryFilterDate) return;
      const existing = grouped.get(group.shipDate) || {
        shipDate: group.shipDate,
        groupCount: 0,
        jobCount: 0,
        totalCost: 0,
      };
      const items = getShipmentItems(group, jobMap, finishedMetaByJobId);
      grouped.set(group.shipDate, {
        shipDate: group.shipDate,
        groupCount: existing.groupCount + 1,
        jobCount: existing.jobCount + items.length,
        totalCost: existing.totalCost + parseCurrency(group.totalCost),
      });
    });
    return Array.from(grouped.values()).sort((left, right) => right.shipDate.localeCompare(left.shipDate));
  }, [finishedMetaByJobId, jobMap, shipmentGroups, shipmentHistoryFilterDate]);

  const selectedJob = selectedJobId ? jobs.find((job) => job.id === selectedJobId) : null;
  const selectedJobFinishMeta = selectedJob ? finishedMetaByJobId.get(selectedJob.id) : null;

  const jobLocationResults = useMemo(() => {
    const query = locationSearch.trim().toLowerCase();
    if (!query) return [];
    return jobs
      .filter((job) => {
        const haystack = `${job.number} ${job.customerName} ${job.generalDescr} ${job.notes}`.toLowerCase();
        return haystack.includes(query);
      })
      .slice(0, 20)
      .map((job) => {
        const locations = assignments
          .filter((assignment) => assignment.kind === "press" && assignment.jobId === job.id)
          .sort((left, right) => {
            if (left.dayKey !== right.dayKey) return left.dayKey.localeCompare(right.dayKey);
            return left.press.localeCompare(right.press);
          });
        return { job, locations };
      });
  }, [assignments, jobs, locationSearch]);

  const queuePressOptions = useMemo(
    () => ["All", ...Array.from(new Set(jobs.map((job) => safeText(job.press)).filter(Boolean))).sort()],
    [jobs]
  );

  const queueCategoryOptions = useMemo(
    () => ["All", ...Array.from(new Set(jobs.map((job) => safeText(job.priority)).filter(Boolean))).sort()],
    [jobs]
  );

  const visiblePresses = useMemo(
    () => (schedulePressFilter === "All" ? PRESS_ORDER : PRESS_ORDER.filter((press) => press === schedulePressFilter)),
    [schedulePressFilter]
  );

  const summary = useMemo(() => {
    return {
      txtOpen: importedSummary.openCount,
      txtClosed: importedSummary.closedCount,
      markedFinished: allUserFinishedJobs.length,
      openRequests: openRequests.length,
      scheduledJobs: activePressAssignments.length,
      shipGroupsOnDate: shipmentGroupsForDay.length,
    };
  }, [activePressAssignments.length, allUserFinishedJobs.length, importedSummary.closedCount, importedSummary.openCount, openRequests.length, shipmentGroupsForDay.length]);

  const shipmentEmailsForSelectedDate = useMemo(
    () =>
      shipmentEmailLogs
        .filter((log) => log.shipDate === selectedShipDate)
        .sort((left, right) => dateSortValue(right.createdAt) - dateSortValue(left.createdAt)),
    [selectedShipDate, shipmentEmailLogs]
  );

  const shipmentEmailHistory = useMemo(
    () =>
      shipmentEmailLogs
        .filter((log) => !shipmentEmailHistoryFilterDate || log.shipDate === shipmentEmailHistoryFilterDate)
        .sort((left, right) => dateSortValue(right.createdAt) - dateSortValue(left.createdAt)),
    [shipmentEmailHistoryFilterDate, shipmentEmailLogs]
  );

  useEffect(() => {
    const validIds = new Set(readyToShipJobs.map((job) => job.id));
    setSelectedShipmentJobs((current) => current.filter((jobId) => validIds.has(jobId)));
  }, [readyToShipJobs]);

  useEffect(() => {
    const validIds = new Set(unassignedFinishedJobs.map((job) => job.id));
    setSelectedShipQueueJobs((current) => current.filter((jobId) => validIds.has(jobId)));
  }, [unassignedFinishedJobs]);

  function selectJob(jobId, shouldScroll = false) {
    setSelectedJobId(jobId);
    if (!shouldScroll) return;
    requestAnimationFrame(() => {
      jobDetailsRef.current?.scrollIntoView({ behavior: "smooth", block: "start" });
    });
  }

  function importText(text) {
    const parsed = parseLabelTraxxText(text);
    if (!parsed.length) return;
    setJobs(parsed);
    setAssignments((current) =>
      current.filter(
        (assignment) => assignment.kind === "manual" || parsed.some((job) => job.id === assignment.jobId)
      )
    );
    setShipmentGroups((current) =>
      current.map((group) => {
        const nextJobMap = new Map(parsed.map((job) => [job.id, job]));
        const jobItems = getShipmentItems(group, nextJobMap, finishedMetaByJobId).filter((item) =>
          parsed.some((job) => job.id === item.id)
        );
        return { ...group, jobItems };
      })
    );
    const firstDate = parsed.find((job) => job.shipByDate)?.shipByDate;
    if (firstDate) setWeekStart(startOfWeek(firstDate));
  }

  function handleUpload(event) {
    const file = event.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => importText(String(reader.result || ""));
    reader.readAsText(file);
  }

  function handleScheduleDrop(event, dayKey, press) {
    event.preventDefault();
    if (!userCanMoveJobs) return;
    const payload = parseDragPayload(event.dataTransfer.getData("application/json"));
    if (payload?.type === "queue" && payload.jobId) {
      addAssignment(payload.jobId, dayKey, press);
      return;
    }
    if (payload?.type === "scheduled" && payload.assignmentId) {
      moveAssignment(payload.assignmentId, dayKey, press);
      return;
    }
    const fallbackJobId = event.dataTransfer.getData("text/plain");
    if (fallbackJobId) addAssignment(fallbackJobId, dayKey, press);
  }

  function addAssignment(jobId, dayKey, press) {
    if (!userCanMoveJobs) return;
    setAssignments((current) => {
      const exists = current.some(
        (assignment) =>
          assignment.jobId === jobId &&
          assignment.dayKey === dayKey &&
          assignment.press === press &&
          assignment.kind === "press" &&
          assignment.status !== "finished"
      );
      let next = current;
      if (!exists) {
        next = [
          ...next,
          {
            id: makeId("asg"),
            jobId,
            dayKey,
            press,
            kind: "press",
            status: "scheduled",
            createdAt: new Date().toISOString(),
            finishedAt: null,
            finishedBy: "",
          },
        ];
      }
      return next;
    });
  }

  function addManualScheduleEntry(event) {
    event.preventDefault();
    if (!userCanMoveJobs) return;
    const title = safeText(manualScheduleForm.title);
    const dayKey = safeText(manualScheduleForm.dayKey);
    const press = safeText(manualScheduleForm.press);
    if (!title || !dayKey || !PRESS_ORDER.includes(press)) return;

    setAssignments((current) => [
      {
        id: makeId("manual"),
        jobId: "",
        manualTitle: title,
        dayKey,
        press,
        kind: "manual",
        status: "scheduled",
        createdAt: new Date().toISOString(),
        finishedAt: null,
        finishedBy: "",
      },
      ...current,
    ]);
    setManualScheduleForm((current) => ({ ...current, title: "" }));
  }

  function moveAssignment(assignmentId, dayKey, press) {
    if (!userCanMoveJobs) return;
    setAssignments((current) => {
      const assignmentToMove = current.find((assignment) => assignment.id === assignmentId);
      if (!assignmentToMove) return current;
      if (assignmentToMove.kind === "press" && assignmentToMove.status === "finished") return current;

      const duplicate = current.some(
        (assignment) =>
          assignment.id !== assignmentId &&
          assignment.dayKey === dayKey &&
          assignment.press === press &&
          (
            (assignmentToMove.kind === "manual" &&
              assignment.kind === "manual" &&
              comparableUsername(assignment.manualTitle) === comparableUsername(assignmentToMove.manualTitle)) ||
            (assignmentToMove.kind === "press" &&
              assignment.jobId === assignmentToMove.jobId &&
              assignment.kind === "press" &&
              assignment.status !== "finished")
          )
      );
      if (duplicate) return current;

      return current.map((assignment) => {
        if (assignment.id === assignmentId) {
          return {
            ...assignment,
            dayKey,
            press,
          };
        }
        return assignment;
      });
    });
  }

  function duplicateAssignmentToNextDay(assignmentId) {
    if (!userCanMoveJobs) return;
    setAssignments((current) => {
      const assignmentToCopy = current.find((assignment) => assignment.id === assignmentId);
      if (!assignmentToCopy) return current;
      if (assignmentToCopy.kind === "press" && assignmentToCopy.status === "finished") return current;

      const currentIndex = weekColumns.findIndex((day) => day.key === assignmentToCopy.dayKey);
      if (currentIndex < 0 || currentIndex >= weekColumns.length - 1) return current;

      const nextDayKey = weekColumns[currentIndex + 1].key;
      const exists = current.some(
        (assignment) =>
          assignment.id !== assignmentId &&
          assignment.dayKey === nextDayKey &&
          assignment.press === assignmentToCopy.press &&
          (
            (assignmentToCopy.kind === "manual" &&
              assignment.kind === "manual" &&
              comparableUsername(assignment.manualTitle) === comparableUsername(assignmentToCopy.manualTitle)) ||
            (assignmentToCopy.kind === "press" &&
              assignment.jobId === assignmentToCopy.jobId &&
              assignment.kind === "press" &&
              assignment.status !== "finished")
          )
      );
      if (exists) return current;

      return [
        ...current,
        {
          ...assignmentToCopy,
          id: makeId("asg"),
          dayKey: nextDayKey,
          createdAt: new Date().toISOString(),
          finishedAt: null,
          finishedBy: "",
          status: "scheduled",
        },
      ];
    });
  }

  function removeAssignment(assignmentId) {
    if (!userCanMoveJobs) return;
    setAssignments((current) => current.filter((assignment) => assignment.id !== assignmentId));
  }

  function finishJob(jobId) {
    if (!currentUser || !userCanEdit) return;
    const finishedAt = new Date().toISOString();
    const defaultShipDate = isoDate(new Date(finishedAt));
    const finishedBy = currentUser.username;
    const job = jobMap.get(jobId);
    const fallbackDayKey = job?.shipByDate ? isoDate(job.shipByDate) : weekColumns[0]?.key || todayKey();
    const fallbackPress = job?.press && PRESS_ORDER.includes(job.press) ? job.press : "Rewind";

    setAssignments((current) => {
      let foundPress = false;
      let next = current.map((assignment) => {
        if (assignment.jobId !== jobId) return assignment;
        foundPress = true;
        return {
          ...assignment,
          status: "finished",
          finishedAt,
          finishedBy,
          shipDate: assignment.shipDate || defaultShipDate,
        };
      });

      if (!foundPress) {
        next = [
          ...next,
          {
            id: makeId("asg"),
            jobId,
            dayKey: fallbackDayKey,
            press: fallbackPress,
            kind: "press",
            status: "finished",
            createdAt: finishedAt,
            finishedAt,
            finishedBy,
            shipDate: defaultShipDate,
          },
        ];
      }

      return next;
    });
  }

  function autoPlace() {
    if (!userCanMoveJobs) return;
    setAssignments((current) => {
      let next = [...current];
      jobs.forEach((job) => {
        if (!job.shipByDate) return;
        const dayKey = isoDate(job.shipByDate);
        if (!weekKeys.has(dayKey)) return;
        const press = PRESS_ORDER.includes(job.press) ? job.press : "Rewind";
        const exists = next.some(
          (assignment) =>
            assignment.jobId === job.id &&
            assignment.dayKey === dayKey &&
            assignment.press === press &&
            assignment.kind === "press" &&
            assignment.status !== "finished"
        );
        if (!exists) {
          next.push({
            id: makeId("asg"),
            jobId: job.id,
            dayKey,
            press,
            kind: "press",
            status: "scheduled",
            createdAt: new Date().toISOString(),
            finishedAt: null,
            finishedBy: "",
          });
        }
      });
      return next;
    });
  }

  function clearBoard() {
    if (!userCanMoveJobs) return;
    setAssignments([]);
  }

  function exportSchedule() {
    const rows = [
      [
        "Day",
        "Press",
        "Number",
        "CustomerName",
        "Description",
        "Priority",
        "EST Time",
        "Quantity",
        "Footage",
        "Stock",
        "Ship Date",
        "Imported Status",
        "Marked Finished By",
        "Marked Finished At",
      ],
    ];

    weekColumns.forEach((day) => {
      PRESS_ORDER.forEach((press) => {
        (board[day.key]?.[press] || []).forEach(({ job }) => {
          const finishMeta = finishedMetaByJobId.get(job.id);
          rows.push([
            day.label,
            press,
            job.number,
            job.customerName,
            job.generalDescr,
            job.priority,
            job.estPressTime,
            job.ticQuantity,
            job.estFootage,
            job.stockDisplay,
            formatDate(job.shipByDate),
            job.ticketStatus,
            finishMeta?.finishedBy || "",
            finishMeta?.finishedAt || "",
          ]);
        });
      });
    });

    doneJobs.forEach((job) => {
      rows.push([
        "Ship Ready",
        "History",
        job.number,
        job.customerName,
        job.generalDescr,
        job.priority,
        job.estPressTime,
        job.ticQuantity,
        job.estFootage,
        job.stockDisplay,
        formatDate(job.shipByDate),
        job.ticketStatus,
        job.finishMeta?.finishedBy || "",
        job.finishMeta?.finishedAt || "",
      ]);
    });

    const csv = rows.map((row) => row.map(csvEscape).join(",")).join("\n");
    downloadFile(`schedule-${isoDate(weekStart)}.csv`, csv, "text/csv;charset=utf-8");
  }

  async function handleDraftAttachmentChange(event) {
    if (!currentUser || !userCanEdit) return;
    const nextAttachments = await buildAttachments(event.target.files, currentUser.username);
    if (nextAttachments.length) {
      setRequestDraftAttachments((current) => [...current, ...nextAttachments]);
    }
    event.target.value = "";
  }

  function removeDraftAttachment(attachmentId) {
    setRequestDraftAttachments((current) => current.filter((attachment) => attachment.id !== attachmentId));
  }

  async function handleSuppliesDraftAttachmentChange(event) {
    if (!currentUser || !userCanEdit) return;
    const nextAttachments = await buildAttachments(event.target.files, currentUser.username);
    if (nextAttachments.length) {
      setSuppliesDraftAttachments((current) => [...current, ...nextAttachments]);
    }
    event.target.value = "";
  }

  function removeSuppliesDraftAttachment(attachmentId) {
    setSuppliesDraftAttachments((current) => current.filter((attachment) => attachment.id !== attachmentId));
  }

  async function handleShipmentDraftAttachmentChange(event) {
    if (!currentUser || !userCanEdit) return;
    const nextAttachments = await buildAttachments(event.target.files, currentUser.username);
    if (nextAttachments.length) {
      setShipmentDraftAttachments((current) => [...current, ...nextAttachments]);
    }
    event.target.value = "";
  }

  function removeShipmentDraftAttachment(attachmentId) {
    setShipmentDraftAttachments((current) => current.filter((attachment) => attachment.id !== attachmentId));
  }

  function submitRequest(event) {
    event.preventDefault();
    if (!currentUser || !userCanEdit) return;
    if (!requestForm.jobNumber || !requestForm.customer || !requestForm.requestorName || !requestForm.description) {
      return;
    }

    setRequests((current) => [
      {
        id: makeId("req"),
        ...requestForm,
        attachments: requestDraftAttachments,
        createdAt: new Date().toISOString(),
        createdByAccount: currentUser.username,
        completedAt: null,
        completedByAccount: "",
        status: "open",
      },
      ...current,
    ]);
    setRequestForm(EMPTY_REQUEST_FORM);
    setRequestDraftAttachments([]);
    setActiveTab("Open Requests");
  }

  function submitPullPaperRequest(event) {
    event.preventDefault();
    if (!currentUser || !userCanEdit) return;
    if (!pullPaperForm.details.trim()) return;

    setPullPaperRequests((current) => [
      {
        id: makeId("paper"),
        details: pullPaperForm.details.trim(),
        target: pullPaperForm.target,
        status: "open",
        createdAt: new Date().toISOString(),
        createdBy: currentUser.username,
        completedAt: null,
        completedBy: "",
      },
      ...current,
    ]);
    setPullPaperForm(EMPTY_PULL_PAPER_FORM);
  }

  function submitSuppliesRequest(event) {
    event.preventDefault();
    if (!currentUser || !userCanEdit) return;
    if (!suppliesForm.details.trim()) return;

    setSuppliesRequests((current) => [
      {
        id: makeId("supply"),
        details: suppliesForm.details.trim(),
        status: "open",
        attachments: suppliesDraftAttachments,
        createdAt: new Date().toISOString(),
        createdBy: currentUser.username,
        completedAt: null,
        completedBy: "",
      },
      ...current,
    ]);
    setSuppliesForm(EMPTY_SUPPLIES_FORM);
    setSuppliesDraftAttachments([]);
  }

  function markPullPaperRequestDone(requestId) {
    if (!currentUser || !userCanEdit) return;
    const completedAt = new Date().toISOString();
    setPullPaperRequests((current) =>
      current.map((request) =>
        request.id === requestId
          ? {
              ...request,
              status: "done",
              completedAt,
              completedBy: currentUser.username,
            }
          : request
      )
    );
  }

  function deletePullPaperRequest(requestId) {
    if (!userCanEdit) return;
    setPullPaperRequests((current) => current.filter((request) => request.id !== requestId));
  }

  function markSuppliesRequestDone(requestId) {
    if (!currentUser || !userCanEdit) return;
    const completedAt = new Date().toISOString();
    setSuppliesRequests((current) =>
      current.map((request) =>
        request.id === requestId
          ? {
              ...request,
              status: "done",
              completedAt,
              completedBy: currentUser.username,
            }
          : request
      )
    );
  }

  function deleteSuppliesRequest(requestId) {
    if (!userCanEdit) return;
    setSuppliesRequests((current) => current.filter((request) => request.id !== requestId));
  }

  async function addSuppliesRequestAttachments(requestId, fileList) {
    if (!currentUser || !userCanEdit) return;
    const attachments = await buildAttachments(fileList, currentUser.username);
    if (!attachments.length) return;
    setSuppliesRequests((current) =>
      current.map((request) =>
        request.id === requestId
          ? {
              ...request,
              attachments: [...(request.attachments || []), ...attachments],
            }
          : request
      )
    );
  }

  function removeSuppliesRequestAttachment(requestId, attachmentId) {
    if (!userCanEdit) return;
    setSuppliesRequests((current) =>
      current.map((request) =>
        request.id === requestId
          ? {
              ...request,
              attachments: (request.attachments || []).filter((attachment) => attachment.id !== attachmentId),
            }
          : request
      )
    );
  }

  function markRequestDone(requestId) {
    if (!currentUser || !userCanEdit) return;
    const completedAt = new Date().toISOString();
    setRequests((current) =>
      current.map((request) =>
        request.id === requestId
          ? {
              ...request,
              status: "done",
              completedAt,
              completedByAccount: currentUser.username,
            }
          : request
      )
    );
  }

  function deleteRequest(requestId) {
    if (!userCanEdit) return;
    setRequests((current) => current.filter((request) => request.id !== requestId));
  }

  async function addRequestAttachments(requestId, fileList) {
    if (!currentUser || !userCanEdit) return;
    const attachments = await buildAttachments(fileList, currentUser.username);
    if (!attachments.length) return;
    setRequests((current) =>
      current.map((request) =>
        request.id === requestId
          ? {
              ...request,
              attachments: [...(request.attachments || []), ...attachments],
            }
          : request
      )
    );
  }

  function removeRequestAttachment(requestId, attachmentId) {
    if (!userCanEdit) return;
    setRequests((current) =>
      current.map((request) =>
        request.id === requestId
          ? {
              ...request,
              attachments: (request.attachments || []).filter((attachment) => attachment.id !== attachmentId),
            }
          : request
      )
    );
  }

  function toggleShipmentJob(jobId) {
    setSelectedShipmentJobs((current) =>
      current.includes(jobId) ? current.filter((value) => value !== jobId) : [...current, jobId]
    );
  }

  function toggleShipQueueJob(jobId) {
    setSelectedShipQueueJobs((current) =>
      current.includes(jobId) ? current.filter((value) => value !== jobId) : [...current, jobId]
    );
  }

  function assignShipDateToFinishedJobs() {
    if (!userCanEdit) return;
    if (!selectedShipQueueJobs.length || !shipDateDraft) return;
    const targetIds = new Set(selectedShipQueueJobs);
    setAssignments((current) =>
      current.map((assignment) => {
        if (!targetIds.has(assignment.jobId) || assignment.kind !== "press" || assignment.status !== "finished") {
          return assignment;
        }
        return {
          ...assignment,
          shipDate: shipDateDraft,
        };
      })
    );
    setSelectedShipDate(shipDateDraft);
    setSelectedShipQueueJobs([]);
  }

  function addShipmentMethod() {
    if (!userCanEdit) return;
    const method = safeText(newShipmentMethod);
    if (!method) return;
    setShipmentMethods((current) => {
      if (current.some((item) => comparableUsername(item) === comparableUsername(method))) return current;
      return [...current, method];
    });
    setShipmentForm((current) => ({ ...current, method }));
    setNewShipmentMethod("");
  }

  function removeShipmentMethod(methodToRemove) {
    if (!userCanEdit) return;
    setShipmentMethods((current) => {
      const next = current.filter((method) => method !== methodToRemove);
      return next.length ? next : current;
    });
    setShipmentForm((current) => {
      if (current.method !== methodToRemove) return current;
      const fallback = shipmentMethods.find((method) => method !== methodToRemove) || current.method;
      return { ...current, method: fallback };
    });
  }

  function createShipmentGroup(event) {
    event.preventDefault();
    if (!userCanEdit) return;
    if (!selectedShipmentJobs.length) return;
    const jobItems = selectedShipmentJobs
      .map((jobId) => {
        const job = jobMap.get(jobId);
        if (!job) return null;
        return snapshotJobForShipment(job, finishedMetaByJobId.get(jobId));
      })
      .filter(Boolean);

    setShipmentGroups((current) => [
      {
        id: makeId("ship"),
        label: shipmentForm.label || `${shipmentForm.method} shipment`,
        method: shipmentForm.method,
        packageCount: parseNumber(shipmentForm.packageCount),
        packageType: shipmentForm.packageType,
        totalCost: parseCurrency(shipmentForm.totalCost),
        billAmount: parseCurrency(shipmentForm.billAmount),
        notes: shipmentForm.notes,
        shipDate: shipmentForm.shipDate,
        createdAt: new Date().toISOString(),
        createdBy: currentUser?.username || "",
        attachments: shipmentDraftAttachments,
        jobItems,
      },
      ...current,
    ]);

    setSelectedShipmentJobs([]);
    setShipmentForm({ ...EMPTY_SHIPMENT_FORM, shipDate: selectedShipDate });
    setShipmentDraftAttachments([]);
  }

  function deleteShipmentGroup(groupId) {
    if (!userCanEdit) return;
    setShipmentGroups((current) => current.filter((group) => group.id !== groupId));
  }

  async function addShipmentGroupAttachments(groupId, fileList) {
    if (!currentUser || !userCanEdit) return;
    const attachments = await buildAttachments(fileList, currentUser.username);
    if (!attachments.length) return;
    setShipmentGroups((current) =>
      current.map((group) =>
        group.id === groupId
          ? {
              ...group,
              attachments: [...(group.attachments || []), ...attachments],
            }
          : group
      )
    );
  }

  function removeShipmentGroupAttachment(groupId, attachmentId) {
    if (!userCanEdit) return;
    setShipmentGroups((current) =>
      current.map((group) =>
        group.id === groupId
          ? {
              ...group,
              attachments: (group.attachments || []).filter((attachment) => attachment.id !== attachmentId),
            }
          : group
      )
    );
  }

  function updateJobRecommendedPress(jobId, press) {
    if (!userCanEdit) return;
    const nextPress = normalizePressValue(press);
    if (!PRESS_ORDER.includes(nextPress)) return;
    setJobs((current) =>
      current.map((job) =>
        job.id === jobId
          ? {
              ...job,
              press: nextPress,
            }
          : job
      )
    );
  }

  function addNote(event) {
    event.preventDefault();
    if (!currentUser || !userCanEdit) return;
    const text = safeText(noteForm.text);
    if (!text) return;
    setNotes((current) => [
      {
        id: makeId("note"),
        ownerUsername: currentUser.username,
        text,
        completed: false,
        createdAt: new Date().toISOString(),
        completedAt: "",
      },
      ...current,
    ]);
    setNoteForm(EMPTY_NOTE_FORM);
  }

  function toggleNote(noteId) {
    if (!currentUser || !userCanEdit) return;
    setNotes((current) =>
      current.map((note) =>
        note.id === noteId
          ? {
              ...note,
              completed: !note.completed,
              completedAt: note.completed ? "" : new Date().toISOString(),
            }
          : note
      )
    );
  }

  function deleteNote(noteId) {
    if (!currentUser || !userCanEdit) return;
    setNotes((current) => current.filter((note) => note.id !== noteId));
  }

  function toggleUserTab(userId, tab) {
    if (!userCanManageUsers) return;
    setUsers((current) =>
      current.map((user) => {
        if (user.id !== userId) return user;
        const nextTabs = user.tabs.includes(tab)
          ? user.tabs.filter((value) => value !== tab)
          : [...user.tabs, tab];
        return { ...user, tabs: normalizeUserTabs(nextTabs, user.role, user.isAdmin, user.canManageUsers) };
      })
    );
  }

  function updateUserAccess(userId, updates) {
    if (!userCanManageUsers) return;
    setUsers((current) => {
      const target = current.find((user) => user.id === userId);
      if (!target) return current;

      if (Object.prototype.hasOwnProperty.call(updates, "canManageUsers")) {
        const managers = current.filter((user) => user.canManageUsers);
        if (target.canManageUsers && !updates.canManageUsers && managers.length <= 1) {
          window.alert("Keep at least one user with admin access.");
          return current;
        }
      }

      return current.map((user) => {
        if (user.id !== userId) return user;
        const nextCanManageUsers = Object.prototype.hasOwnProperty.call(updates, "canManageUsers")
          ? !!updates.canManageUsers
          : user.canManageUsers;
        const nextRole = nextCanManageUsers ? "Management" : user.role === "Management" ? "Warehouse/Shipper" : user.role;
        return {
          ...user,
          ...updates,
          role: nextRole,
          isAdmin: nextCanManageUsers,
          canManageUsers: nextCanManageUsers,
          accessMode: Object.prototype.hasOwnProperty.call(updates, "accessMode")
            ? normalizeAccessMode(updates.accessMode, nextRole, nextCanManageUsers)
            : user.accessMode,
          tabs: normalizeUserTabs(
            Object.prototype.hasOwnProperty.call(updates, "tabs") ? updates.tabs : user.tabs,
            nextRole,
            nextCanManageUsers,
            nextCanManageUsers
          ),
        };
      });
    });
  }

  function buildShipmentEmailDraft() {
    const groups = shipmentGroupsForDay;
    const totalCost = groups.reduce((sum, group) => sum + parseCurrency(group.totalCost), 0);
    const totalBill = groups.reduce((sum, group) => sum + parseCurrency(group.billAmount), 0);
    const methods = Array.from(new Set(groups.map((group) => group.method).filter(Boolean)));
    const lines = [`Daily shipment summary for ${selectedShipDate}`, ""];

    groups.forEach((group) => {
      const items = getShipmentItems(group, jobMap, finishedMetaByJobId);
      lines.push(`${group.label} | ${group.method}`);
      lines.push(`Skids / cartons: ${group.packageCount || 0} ${group.packageType || "Skids"}`);
      lines.push(`Our cost: ${formatCurrency(group.totalCost)}`);
      lines.push(`Bill: ${formatCurrency(group.billAmount)}`);
      items.forEach((item) => {
        lines.push(`- ${item.customerName} ${item.number}: ${item.generalDescr}`);
      });
      if (group.notes) lines.push(`Notes: ${group.notes}`);
      lines.push("");
    });

    if (!groups.length) {
      lines.push("No shipment groups have been created for this date yet.");
      lines.push("");
    }

    lines.push(`Total cost: ${formatCurrency(totalCost)}`);
    lines.push(`Total bill: ${formatCurrency(totalBill)}`);

    return {
      subject: `Daily shipments for ${selectedShipDate}`,
      body: lines.join("\n"),
      groupCount: groups.length,
      jobCount: groups.reduce((sum, group) => sum + getShipmentItems(group, jobMap, finishedMetaByJobId).length, 0),
      totalCost,
      totalBill,
      methods,
    };
  }

  function openShipmentEmailDraft() {
    if (!userCanEdit) return;
    const recipients = safeText(shipmentEmailForm.recipients);
    const draft = buildShipmentEmailDraft();
    const params = new URLSearchParams({
      subject: draft.subject,
      body: draft.body,
    });
    window.location.href = `mailto:${encodeURIComponent(recipients)}?${params.toString()}`;
  }

  function logShipmentEmail() {
    if (!currentUser || !selectedShipDate || !userCanEdit) return;
    const draft = buildShipmentEmailDraft();
    setShipmentEmailLogs((current) => [
      {
        id: makeId("email"),
        shipDate: selectedShipDate,
        recipients: safeText(shipmentEmailForm.recipients),
        subject: draft.subject,
        body: draft.body,
        jobCount: draft.jobCount,
        totalCost: draft.totalCost,
        totalBill: draft.totalBill,
        methods: draft.methods,
        createdAt: new Date().toISOString(),
        createdBy: currentUser.username,
      },
      ...current,
    ]);
  }

  function loadDemo() {
    const sample = `Press\tNumber\tCustomerName\tGeneralDescr\tCustPONum\tPriority\tShip_by_Date\tEntryDate\tDue_on_Site_Date\tStockNum2\tStockNum1\tStatus\tMainTool\tToolNo2\tTicQuantity\tEstFootage\tEstPressTime\tNotes
5.1\t10159\tData Graphics\t1.625" Cap One Circle 70072\t223000DG\tHigh\t04/28/26\t04/18/26\t04/29/26\t266\t\tOpen\t946\t\t3,612,279\t113,847\t9.76\tExample long notes
8\t11180\tPremio Foods\t3.125"x4.1875" Premio\t4500081640\tDigital\t04/29/26\t04/16/26\t04/30/26\t266\t590\tDone\tD-904\t\t96,000\t13,328\t1.73\tImported done should not ship until you mark it finished.
6.1\t11194\tPremio Foods\t3.25"x5" Premio Contract Release\t4500081729\tRelease\t04/28/26\t04/21/26\t04/29/26\t266\t590\tOpen\t668\t\t48,000\t13,020\t4.64\tContract PO 4600004905
9\t11022\tData Graphics\t1.625" Cap One Circle 69797\t\tHigh\t04/29/26\t03/24/26\t05/04/26\t266\t\tOpen\t946\t\t2,686,950\t86,266\t7.91\tArt Due 4/23`;
    importText(sample);
  }

  function switchAuthView(nextView) {
    setAuthView(nextView);
    setLoginError("");
    setRegisterError("");
    setRegisterSuccess("");
  }

  function handleLogin(event) {
    event.preventDefault();
    const match = users.find(
      (user) =>
        comparableUsername(user.username) === comparableUsername(loginForm.username) &&
        user.password === loginForm.password
    );
    if (!match) {
      setLoginError("Invalid username or password.");
      return;
    }
    setCurrentUsername(match.username);
    setSessionExpiresAt(new Date(Date.now() + LOGIN_SESSION_DURATION_MS).toISOString());
    setLoginForm(EMPTY_LOGIN_FORM);
    setLoginError("");
    setRegisterSuccess("");
    setAuthView("login");
    setActiveTab("Scheduler");
  }

  function submitRegistrationRequest(event) {
    event.preventDefault();
    const username = safeText(registerForm.username);
    const password = safeText(registerForm.password);

    if (!username || !password) {
      setRegisterError("Enter both a name and password.");
      setRegisterSuccess("");
      return;
    }

    const hasUser = users.some((user) => comparableUsername(user.username) === comparableUsername(username));
    if (hasUser) {
      setRegisterError("That username already exists.");
      setRegisterSuccess("");
      return;
    }

    const hasPendingRequest = registrationRequests.some(
      (request) =>
        comparableUsername(request.username) === comparableUsername(username) && request.status === "pending"
    );
    if (hasPendingRequest) {
      setRegisterError("That account request is already waiting for approval.");
      setRegisterSuccess("");
      return;
    }

    setRegistrationRequests((current) => [
      ...current,
      {
        id: makeId("registration"),
        username,
        password,
        status: "pending",
        createdAt: new Date().toISOString(),
        createdBy: username,
        approvedAt: "",
        approvedBy: "",
        deniedAt: "",
        deniedBy: "",
      },
    ]);
    setRegisterForm(EMPTY_REGISTER_FORM);
    setRegisterError("");
    setRegisterSuccess("Registration request sent. A manager will need to approve it.");
    setAuthView("login");
  }

  function handleLogout() {
    setCurrentUsername("");
    setSessionExpiresAt("");
    setAuthView("login");
    setActiveTab("Scheduler");
  }

  function createUser(event) {
    event.preventDefault();
    if (!userCanManageUsers) return;
    const username = safeText(userForm.username);
    const password = safeText(userForm.password);
    if (!username || !password) return;
    const exists = users.some((user) => comparableUsername(user.username) === comparableUsername(username));
    if (exists) {
      window.alert("That username already exists.");
      return;
    }
    setUsers((current) => [
      ...current,
      {
        id: makeId("user"),
        username,
        password,
        role: userForm.canManageUsers ? "Management" : "Warehouse/Shipper",
        accessMode: normalizeAccessMode(userForm.accessMode, userForm.canManageUsers ? "Management" : "Warehouse/Shipper", userForm.canManageUsers),
        tabs: normalizeUserTabs(userForm.tabs, userForm.canManageUsers ? "Management" : "Warehouse/Shipper", userForm.canManageUsers, userForm.canManageUsers),
        canManageUsers: !!userForm.canManageUsers,
        isAdmin: !!userForm.canManageUsers,
        createdAt: new Date().toISOString(),
        createdBy: currentUser.username,
      },
    ]);
    setUserForm({ ...EMPTY_USER_FORM, tabs: [...EMPTY_USER_FORM.tabs] });
  }

  function updateUserPassword(userId) {
    if (!userCanManageUsers) return;
    const nextPassword = safeText(userPasswordDrafts[userId]);
    if (!nextPassword) return;
    setUsers((current) =>
      current.map((user) => (user.id === userId ? { ...user, password: nextPassword } : user))
    );
    setUserPasswordDrafts((current) => ({ ...current, [userId]: "" }));
  }

  function approveRegistrationRequest(requestId) {
    if (!userCanManageUsers) return;
    const request = registrationRequests.find((item) => item.id === requestId);
    if (!request) return;

    const exists = users.some(
      (user) => comparableUsername(user.username) === comparableUsername(request.username)
    );
    if (exists) {
      window.alert("That username already exists. Delete or rename the pending request first.");
      return;
    }

    setUsers((current) => [
      ...current,
      {
        id: makeId("user"),
        username: request.username,
        password: request.password,
        role: "Warehouse/Shipper",
        accessMode: "edit",
        tabs: normalizeUserTabs(EMPTY_USER_FORM.tabs, "Warehouse/Shipper", false, false),
        canManageUsers: false,
        isAdmin: false,
        createdAt: new Date().toISOString(),
        createdBy: currentUser.username,
      },
    ]);
    setRegistrationRequests((current) => current.filter((item) => item.id !== requestId));
  }

  function denyRegistrationRequest(requestId) {
    if (!userCanManageUsers) return;
    setRegistrationRequests((current) => current.filter((item) => item.id !== requestId));
  }

  if (!isReady) {
    return (
      <div className="min-h-screen bg-stone-100 p-6 text-stone-900">
        <div className="mx-auto max-w-xl rounded-3xl border border-stone-300 bg-gradient-to-br from-stone-50 via-white to-stone-100 p-8 text-center shadow-sm shadow-stone-300/40">
          <div className="text-lg font-semibold">Loading scheduler...</div>
        </div>
      </div>
    );
  }

  if (!currentUser) {
    return (
      <LoginScreen
        authView={authView}
        loginForm={loginForm}
        loginError={loginError}
        registerForm={registerForm}
        registerError={registerError}
        registerSuccess={registerSuccess}
        users={users}
        onChangeLogin={setLoginForm}
        onChangeRegister={setRegisterForm}
        onChangeView={switchAuthView}
        onSubmitLogin={handleLogin}
        onSubmitRegister={submitRegistrationRequest}
      />
    );
  }

  const schedulerCards = [
    ["TXT open", summary.txtOpen],
    ["TXT closed", summary.txtClosed],
    ["Marked finished", summary.markedFinished],
    ["Open requests", summary.openRequests],
    ["Scheduled jobs", summary.scheduledJobs],
    ["Ship groups", summary.shipGroupsOnDate],
  ];

  const tabBadges = {
    Notes: userNotes.filter((note) => !note.completed).length,
    "Open Requests": openRequests.length,
    "Pull Paper Request": openPullPaperRequests.length,
    "Supplies Request": openSuppliesRequests.length,
  };

  return (
    <div className="min-h-screen bg-stone-100 text-stone-900">
      <div className="mx-auto max-w-[1900px] p-4 md:p-6">
        <div className="mb-6 rounded-[2rem] border border-stone-300 bg-gradient-to-br from-stone-50 via-white to-stone-100 p-5 shadow-sm shadow-stone-300/40">
          <div className="flex flex-col gap-4 xl:flex-row xl:items-end xl:justify-between">
            <div>
              <p className="text-xs font-semibold uppercase tracking-[0.18em] text-emerald-900">Production board</p>
              <h1 className="mt-1 text-2xl font-semibold tracking-tight">Label Traxx Scheduler</h1>
              <p className="mt-2 max-w-3xl text-sm text-stone-700">
                Logged in as {currentUser.username}. Request history, attachments, and completed work are now tied to user accounts.
              </p>
            </div>
            <div className="flex flex-col gap-3 xl:items-end">
              <div className="flex flex-wrap gap-2">
                {tabs.map((tab) => (
                  <button
                    key={tab}
                    onClick={() => setActiveTab(tab)}
                    className={`rounded-2xl px-4 py-2 text-sm font-medium transition ${
                      activeTab === tab
                        ? "bg-emerald-900 text-stone-50 shadow-sm"
                        : "border border-stone-300 bg-stone-50 text-stone-700 hover:bg-stone-100"
                    }`}
                  >
                    <span>{tab}</span>
                    {tabBadges[tab] > 0 && (
                      <span
                        className={`ml-2 rounded-full px-2 py-0.5 text-[11px] font-semibold ${
                          activeTab === tab ? "bg-white/15 text-stone-50" : "bg-stone-200 text-stone-800"
                        }`}
                      >
                        {tabBadges[tab]}
                      </span>
                    )}
                  </button>
                ))}
              </div>
              <div className="flex items-center gap-2 text-sm">
                <span className={`rounded-full px-3 py-2 ${syncTone(syncStatus)}`}>
                  {syncStatus}
                  {lastSyncAt ? ` • ${formatDateTime(lastSyncAt)}` : ""}
                </span>
                <span className="rounded-full bg-stone-200 px-3 py-2 text-stone-800">
                  {currentUserRole} / {currentUserAccessMode === "edit" ? "Edit" : "View only"}: {currentUser.username}
                </span>
                <button onClick={handleLogout} className="rounded-2xl border border-stone-300 bg-stone-50 px-3 py-2 text-stone-800">
                  Log out
                </button>
              </div>
            </div>
          </div>
        </div>

        <div className="mb-6 grid gap-3 md:grid-cols-3 xl:grid-cols-6">
          {schedulerCards.map(([label, value]) => (
            <div key={label} className="rounded-3xl border border-stone-300 bg-stone-50 p-4 shadow-sm shadow-stone-300/30">
              <div className="text-xs uppercase tracking-[0.16em] text-stone-600">{label}</div>
              <div className="mt-2 text-2xl font-semibold tracking-tight">{value}</div>
            </div>
          ))}
        </div>

        {activeTab === "Scheduler" && (
          <div className="space-y-4">
            <div className="grid gap-4 xl:grid-cols-[340px_minmax(0,1fr)]">
              <div className="rounded-3xl border border-stone-300 bg-stone-50 p-5 shadow-sm shadow-stone-300/30">
                <div className="mb-4">
                  <div className="text-sm font-semibold">Import and controls</div>
                  <div className="text-xs text-stone-600">Imported done tickets stay informational only until a logged-in user marks them finished.</div>
                </div>

                <div className="grid gap-3">
                  <label className="rounded-2xl border border-dashed border-stone-300 bg-white p-4 text-sm text-stone-700 hover:border-emerald-800">
                    <div className="font-medium text-stone-900">Upload Label Traxx TXT</div>
                    <input type="file" accept=".txt,.tsv,text/plain" onChange={handleUpload} className="mt-3 block w-full text-xs" />
                  </label>

                  <div className="rounded-2xl border border-stone-300 bg-white p-3">
                    <div className="mb-2 flex items-center justify-between">
                      <div className="text-sm font-medium">Paste export text</div>
                      <button onClick={() => importText(pasteText)} className="rounded-xl bg-emerald-900 px-3 py-2 text-xs font-medium text-white">
                        Import
                      </button>
                    </div>
                    <textarea
                      value={pasteText}
                      onChange={(event) => setPasteText(event.target.value)}
                      placeholder="Paste the full TXT export here..."
                      className="h-32 w-full rounded-2xl border border-stone-300 bg-stone-50 p-3 text-xs outline-none placeholder:text-stone-400 focus:border-emerald-800"
                    />
                  </div>

                  <div className="rounded-2xl bg-stone-200/70 px-4 py-3 text-sm text-stone-800">
                    Finishing jobs and completing requests will be recorded under <span className="font-semibold">{currentUser.username}</span>.
                  </div>

                  <div className="grid grid-cols-2 gap-2">
                    <button onClick={loadDemo} className="rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm text-stone-800">
                      Load demo
                    </button>
                    <button onClick={autoPlace} className="rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm text-stone-800">
                      Auto-place
                    </button>
                    <button onClick={exportSchedule} className="rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm text-stone-800">
                      Export CSV
                    </button>
                    <button onClick={clearBoard} className="rounded-2xl border border-amber-300 bg-amber-50 px-3 py-2 text-sm text-amber-900">
                      Clear
                    </button>
                  </div>

                  {canAccessTab(currentUser, "New Request") && (
                    <button
                      onClick={() => setActiveTab("New Request")}
                      className="rounded-2xl bg-emerald-900 px-3 py-2 text-sm font-medium text-white"
                    >
                      Create a request
                    </button>
                  )}

                  <div className="rounded-2xl border border-stone-300 bg-white p-4">
                    <div className="mb-3">
                      <div className="text-sm font-medium text-stone-900">Manual schedule block</div>
                      <div className="text-xs text-stone-600">Reserve a press/day for something like Maintenance, Changeover, or Wash-up.</div>
                    </div>
                    <form onSubmit={addManualScheduleEntry} className="grid gap-3">
                      <Field
                        label="Title"
                        value={manualScheduleForm.title}
                        onChange={(value) => setManualScheduleForm((current) => ({ ...current, title: value }))}
                        placeholder="Maintenance"
                      />
                      <div>
                        <div className="mb-2 text-sm font-medium text-stone-800">Press</div>
                        <select
                          value={manualScheduleForm.press}
                          onChange={(event) => setManualScheduleForm((current) => ({ ...current, press: event.target.value }))}
                          className="w-full rounded-2xl border border-stone-300 bg-stone-50 px-4 py-3 text-sm outline-none focus:border-emerald-800"
                        >
                          {PRESS_ORDER.map((press) => (
                            <option key={press} value={press}>
                              Press {press}
                            </option>
                          ))}
                        </select>
                      </div>
                      <div>
                        <div className="mb-2 text-sm font-medium text-stone-800">Day</div>
                        <select
                          value={manualScheduleForm.dayKey}
                          onChange={(event) => setManualScheduleForm((current) => ({ ...current, dayKey: event.target.value }))}
                          className="w-full rounded-2xl border border-stone-300 bg-stone-50 px-4 py-3 text-sm outline-none focus:border-emerald-800"
                        >
                          {weekColumns.map((day) => (
                            <option key={day.key} value={day.key}>
                              {day.label} - {formatShortDate(day.date)}
                            </option>
                          ))}
                        </select>
                      </div>
                      <button
                        type="submit"
                        disabled={!userCanMoveJobs}
                        className="rounded-2xl border border-stone-300 bg-stone-50 px-4 py-3 text-sm text-stone-800 disabled:cursor-not-allowed disabled:opacity-50"
                      >
                        Add block
                      </button>
                    </form>
                  </div>
                </div>
              </div>

              <div className="rounded-3xl border border-stone-300 bg-stone-50 p-4 shadow-sm shadow-stone-300/30">
                <div className="mb-3 flex items-center justify-between">
                  <div>
                    <div className="text-sm font-semibold">Press queue</div>
                    <div className="text-xs text-stone-600">Search by ticket, then filter the queue by status, press, or category. Jobs stay here even after they are scheduled so you can place them on multiple days.</div>
                  </div>
                  <div className="rounded-xl bg-stone-200 px-2 py-1 text-xs text-stone-700">{unscheduledJobs.length}</div>
                </div>

                <input
                  type="text"
                  value={unscheduledSearch}
                  onChange={(event) => setUnscheduledSearch(event.target.value)}
                  placeholder="Search ticket, customer, or description"
                  className="mb-3 w-full rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm outline-none focus:border-emerald-800"
                />

                <div className="mb-3 grid gap-2 md:grid-cols-3">
                  <select
                    value={queueStatusFilter}
                    onChange={(event) => setQueueStatusFilter(event.target.value)}
                    className="rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm outline-none focus:border-emerald-800"
                  >
                    <option>All</option>
                    <option>Open</option>
                    <option>Done</option>
                  </select>
                  <select
                    value={queuePressFilter}
                    onChange={(event) => setQueuePressFilter(event.target.value)}
                    className="rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm outline-none focus:border-emerald-800"
                  >
                    {queuePressOptions.map((press) => (
                      <option key={press}>{press}</option>
                    ))}
                  </select>
                  <select
                    value={queueCategoryFilter}
                    onChange={(event) => setQueueCategoryFilter(event.target.value)}
                    className="rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm outline-none focus:border-emerald-800"
                  >
                    {queueCategoryOptions.map((category) => (
                      <option key={category}>{category}</option>
                    ))}
                  </select>
                </div>

                <div className="max-h-[480px] space-y-3 overflow-y-auto pr-1">
                  {unscheduledJobs.map((job) => (
                    <JobCard
                      key={job.id}
                      job={job}
                      state={deriveVisibleJobState(job.id, activePressJobIds, userFinishedJobIds)}
                      onClick={() => selectJob(job.id)}
                      onDoubleClick={() => selectJob(job.id, true)}
                      onFinish={userCanEdit ? () => finishJob(job.id) : undefined}
                      weekColumns={weekColumns}
                      canMove={userCanMoveJobs}
                      pressOptions={PRESS_ORDER}
                      onUpdatePress={userCanMoveJobs ? (press) => updateJobRecommendedPress(job.id, press) : undefined}
                      onQuickAssign={(dayKey) => addAssignment(job.id, dayKey, PRESS_ORDER.includes(job.press) ? job.press : "Rewind")}
                    />
                  ))}
                  {!unscheduledJobs.length && (
                    <div className="rounded-2xl border border-dashed border-stone-300 bg-white/70 p-4 text-sm text-stone-600">
                      No open queue jobs match your search.
                    </div>
                  )}
                </div>
              </div>
            </div>

            <div className="rounded-3xl border border-stone-300 bg-stone-50 p-3 shadow-sm shadow-stone-300/30">
              <div className="mb-3 flex flex-col gap-3 rounded-2xl border border-stone-300 bg-gradient-to-r from-stone-100 via-stone-50 to-stone-100 px-4 py-3 xl:flex-row xl:items-center xl:justify-between">
                <div>
                  <div className="text-[11px] font-semibold uppercase tracking-[0.18em] text-stone-600">Week view</div>
                  <div className="mt-1 text-base font-semibold">
                    {formatShortDate(weekColumns[0]?.date)} - {formatShortDate(weekColumns[4]?.date)}
                  </div>
                </div>
                <div className="flex flex-wrap items-center gap-2">
                  <button onClick={() => setWeekStart(addDays(weekStart, -7))} className="rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm text-stone-800">
                    Previous
                  </button>
                  <button onClick={() => setWeekStart(startOfWeek(new Date()))} className="rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm text-stone-800">
                    This week
                  </button>
                  <button onClick={() => setWeekStart(addDays(weekStart, 7))} className="rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm text-stone-800">
                    Next
                  </button>
                  <select
                    value={schedulePressFilter}
                    onChange={(event) => setSchedulePressFilter(event.target.value)}
                    className="rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm outline-none focus:border-emerald-800"
                  >
                    <option value="All">All presses</option>
                    {PRESS_ORDER.map((press) => (
                      <option key={press} value={press}>
                        Press {press}
                      </option>
                    ))}
                  </select>
                  <button
                    type="button"
                    onClick={() => setCondensedSchedule((current) => !current)}
                    className="rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm text-stone-800"
                  >
                    {condensedSchedule ? "Expanded cards" : "Condensed cards"}
                  </button>
                  <input
                    type="text"
                    value={search}
                    onChange={(event) => setSearch(event.target.value)}
                    placeholder="Search all jobs"
                    className="rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm outline-none focus:border-emerald-800"
                  />
                </div>
              </div>
              <div className="grid grid-cols-5 gap-3">
                {weekColumns.map((day) => (
                  <div key={day.key} className="rounded-3xl bg-stone-200/60 p-3">
                    <div className="mb-3 rounded-2xl border border-stone-300 bg-white px-3 py-3">
                      <div className="text-xs uppercase tracking-[0.16em] text-stone-600">{day.label}</div>
                      <div className="mt-1 text-lg font-semibold">{formatShortDate(day.date)}</div>
                    </div>
                    <div className="space-y-3">
                      {visiblePresses.map((press) => {
                        const laneJobs = board[day.key]?.[press] || [];
                        const totalHours = laneJobs.reduce((sum, item) => sum + (item.job?.estPressTime || 0), 0);
                        return (
                          <div
                            key={`${day.key}-${press}`}
                            onDragOver={(event) => event.preventDefault()}
                            onDrop={(event) => handleScheduleDrop(event, day.key, press)}
                            className="rounded-2xl border border-stone-300 bg-white p-2"
                          >
                            <div className="mb-2 flex items-start justify-between gap-2">
                              <div>
                                <div className="text-sm font-semibold">Press {press}</div>
                                <div className="text-[11px] text-stone-600">
                                  {laneJobs.length} jobs - {totalHours.toFixed(2)} hrs
                                </div>
                              </div>
                              <span className="rounded-full bg-stone-200 px-2 py-1 text-[10px] font-medium text-stone-700">
                                drop
                              </span>
                            </div>
                            <div className="space-y-2">
                              {laneJobs.map(({ assignment, job }) => (
                                <CompactScheduleCard
                                  key={assignment.id}
                                  job={job}
                                  assignment={assignment}
                                  finishMeta={job ? finishedMetaByJobId.get(job.id) : null}
                                  compact={condensedSchedule}
                                  onSelect={job ? () => selectJob(job.id) : undefined}
                                  onUnschedule={userCanMoveJobs ? () => removeAssignment(assignment.id) : undefined}
                                  onFinish={job && userCanEdit ? () => finishJob(job.id) : undefined}
                                  onDuplicate={userCanMoveJobs ? () => duplicateAssignmentToNextDay(assignment.id) : undefined}
                                  draggable={userCanMoveJobs && assignment.status !== "finished"}
                                />
                              ))}
                              {!laneJobs.length && (
                                <div className="rounded-2xl border border-dashed border-stone-300 bg-stone-50 p-3 text-center text-[11px] text-stone-500">
                                  Drop here
                                </div>
                              )}
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                ))}
              </div>
            </div>

            <div className="grid gap-4 xl:grid-cols-[minmax(0,1fr)_340px]">
                <div ref={jobDetailsRef} className="rounded-3xl border border-stone-300 bg-stone-50 p-4 shadow-sm shadow-stone-300/30">
                  <div className="mb-3 flex items-center justify-between">
                    <div>
                      <div className="text-sm font-semibold">Job details</div>
                      <div className="text-xs text-stone-600">Select a job from the queue or board.</div>
                    </div>
                    {selectedJob && (
                      <span
                        className={`rounded-full px-2 py-1 text-xs font-medium ${statusTone(
                          deriveVisibleJobState(selectedJob.id, activePressJobIds, userFinishedJobIds)
                        )}`}
                      >
                        {deriveVisibleJobState(selectedJob.id, activePressJobIds, userFinishedJobIds)}
                      </span>
                    )}
                  </div>

                  {selectedJob ? (
                    <div className="space-y-3 text-sm">
                      <div>
                        <div className="text-lg font-semibold">
                          {selectedJob.customerName} {selectedJob.number}
                        </div>
                        <div className="text-stone-700">{selectedJob.generalDescr}</div>
                      </div>
                      <div className="grid grid-cols-2 gap-3 text-xs text-stone-700">
                        <Detail label="Default press" value={selectedJob.press || "-"} />
                        <Detail label="Priority" value={selectedJob.priority || "-"} />
                        <Detail label="Ship by" value={formatDate(selectedJob.shipByDate)} />
                        <Detail label="Imported status" value={selectedJob.ticketStatus || "-"} />
                        <Detail label="Quantity" value={selectedJob.ticQuantity.toLocaleString()} />
                        <Detail label="EST time" value={`${selectedJob.estPressTime.toFixed(2)} hrs`} />
                        <Detail label="PO number" value={selectedJob.custPoNum || "-"} />
                        <Detail label="Main tool" value={selectedJob.mainTool || "-"} />
                        <Detail label="Footage" value={selectedJob.estFootage.toLocaleString()} />
                        <Detail label="Stock" value={selectedJob.stockDisplay || "-"} />
                      </div>
                      <div>
                        <div className="mb-1 text-xs font-semibold uppercase tracking-[0.16em] text-stone-600">Notes</div>
                        <div className="max-h-56 overflow-auto whitespace-pre-wrap rounded-2xl bg-stone-100 p-3 text-sm text-stone-800">
                          {selectedJob.notes || "No notes on this job."}
                        </div>
                      </div>
                    </div>
                  ) : (
                    <div className="rounded-2xl border border-dashed border-stone-300 bg-white/60 p-4 text-sm text-stone-600">
                      No job selected yet.
                    </div>
                  )}
                </div>

                <div className="space-y-4">
                  <div className="rounded-3xl border border-stone-300 bg-stone-50 p-4 shadow-sm shadow-stone-300/30">
                    <div className="mb-3">
                      <div className="text-sm font-semibold">Job location search</div>
                      <div className="text-xs text-stone-600">Search any ticket, customer, or description to see every scheduled or finished day and press.</div>
                    </div>
                    <input
                      type="text"
                      value={locationSearch}
                      onChange={(event) => setLocationSearch(event.target.value)}
                      placeholder="Search job number, customer, or description"
                      className="w-full rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm outline-none focus:border-emerald-800"
                    />
                    <div className="mt-3 max-h-[32vh] space-y-3 overflow-y-auto">
                      {jobLocationResults.map(({ job, locations }) => (
                        <div key={job.id} className="rounded-2xl border border-stone-300 bg-white p-3">
                          <div className="flex items-start justify-between gap-3">
                            <div>
                              <div className="text-sm font-semibold">{job.customerName} {job.number}</div>
                              <div className="mt-1 text-xs text-stone-700">{job.generalDescr}</div>
                            </div>
                            <button
                              onClick={() => selectJob(job.id, true)}
                              className="rounded-xl border border-stone-300 bg-stone-50 px-2 py-1 text-[11px] text-stone-800"
                            >
                              Open
                            </button>
                          </div>
                          <div className="mt-3 space-y-2">
                            {locations.length ? locations.map((location) => (
                              <button
                                key={location.id}
                                onClick={() => {
                                  setWeekStart(startOfWeek(new Date(location.dayKey)));
                                  selectJob(job.id, true);
                                }}
                                className="flex w-full items-center justify-between rounded-2xl bg-stone-100 px-3 py-2 text-left text-xs text-stone-800"
                              >
                                <span>{location.dayKey}</span>
                                <span>
                                  Press {location.press} · {location.status === "finished" ? "done" : location.status}
                                </span>
                              </button>
                            )) : (
                              <div className="rounded-2xl border border-dashed border-stone-300 bg-white/60 p-3 text-xs text-stone-600">
                                This job is not on the schedule yet.
                              </div>
                            )}
                          </div>
                        </div>
                      ))}
                      {!jobLocationResults.length && (
                        <div className="rounded-2xl border border-dashed border-stone-300 bg-white/60 p-4 text-sm text-stone-600">
                          {locationSearch.trim() ? "No jobs matched that search." : "Type a search to see where jobs are scheduled or finished."}
                        </div>
                      )}
                    </div>
                  </div>

                  <div className="rounded-3xl border border-stone-300 bg-stone-50 p-4 shadow-sm shadow-stone-300/30">
                    <div className="mb-3">
                      <div className="text-sm font-semibold">Ship ready today</div>
                      <div className="text-xs text-stone-600">These are only the jobs marked finished today by a logged-in user.</div>
                    </div>
                    <div className="max-h-[56vh] space-y-3 overflow-y-auto">
                      {doneJobs
                        .filter((job) => sameDay(job.finishMeta?.finishedAt, todayKey()))
                        .map((job) => (
                          <JobCard
                            key={job.id}
                            job={job}
                            state="ship"
                            onClick={() => selectJob(job.id)}
                            finishedAt={job.finishMeta?.finishedAt}
                            finishedBy={job.finishMeta?.finishedBy}
                          />
                        ))}
                      {!doneJobs.filter((job) => sameDay(job.finishMeta?.finishedAt, todayKey())).length && (
                        <div className="rounded-2xl border border-dashed border-stone-300 bg-white/60 p-4 text-sm text-stone-600">
                          No jobs marked finished today yet.
                        </div>
                      )}
                    </div>
                  </div>
                </div>
            </div>
          </div>
        )}

        {activeTab === "Notes" && (
          <div className="grid gap-4 xl:grid-cols-[420px_minmax(0,1fr)]">
            <div className="rounded-3xl border border-stone-300 bg-stone-50 p-6 shadow-sm shadow-stone-300/30">
              <div className="mb-5">
                <div className="text-sm font-semibold">My internal checklist</div>
                <div className="text-xs text-stone-600">These notes are tied only to the signed-in user.</div>
              </div>
              <form onSubmit={addNote} className="grid gap-4">
                <div>
                  <div className="mb-2 text-sm font-medium text-stone-800">Checklist item</div>
                  <textarea
                    value={noteForm.text}
                    onChange={(event) => setNoteForm({ text: event.target.value })}
                    placeholder="Write a reminder, checklist item, or personal note"
                    className="h-36 w-full rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm outline-none focus:border-emerald-800"
                  />
                </div>
                <button
                  type="submit"
                  disabled={!userCanEdit}
                  className="rounded-2xl bg-emerald-900 px-4 py-3 text-sm font-medium text-white disabled:cursor-not-allowed disabled:opacity-50"
                >
                  Add note
                </button>
              </form>
            </div>

            <div className="rounded-3xl border border-stone-300 bg-stone-50 p-6 shadow-sm shadow-stone-300/30">
              <div className="mb-4 flex items-center justify-between">
                <div>
                  <div className="text-sm font-semibold">My notes</div>
                  <div className="text-xs text-stone-600">Only {currentUser.username} can see these checklist items.</div>
                </div>
                <div className="rounded-xl bg-stone-200 px-2 py-1 text-xs text-stone-700">
                  {userNotes.filter((note) => !note.completed).length} open
                </div>
              </div>
              <div className="space-y-3">
                {userNotes.map((note) => (
                  <div key={note.id} className="rounded-2xl border border-stone-300 bg-white p-4">
                    <div className="flex gap-3">
                      <input
                        type="checkbox"
                        checked={note.completed}
                        disabled={!userCanEdit}
                        onChange={() => toggleNote(note.id)}
                        className="mt-1 h-4 w-4"
                      />
                      <div className="min-w-0 flex-1">
                        <div className={`whitespace-pre-wrap text-sm ${note.completed ? "text-stone-500 line-through" : "text-stone-800"}`}>
                          {note.text}
                        </div>
                        <div className="mt-2 text-xs text-stone-600">
                          Added {formatDateTime(note.createdAt)}
                          {note.completedAt ? ` · completed ${formatDateTime(note.completedAt)}` : ""}
                        </div>
                      </div>
                      <button
                        onClick={() => deleteNote(note.id)}
                        disabled={!userCanEdit}
                        className="rounded-xl border border-rose-200 px-3 py-2 text-xs text-rose-700 disabled:cursor-not-allowed disabled:opacity-50"
                      >
                        Delete
                      </button>
                    </div>
                  </div>
                ))}
                {!userNotes.length && (
                  <div className="rounded-2xl border border-dashed border-stone-300 bg-white/60 p-5 text-sm text-stone-600">
                    No personal notes yet.
                  </div>
                )}
              </div>
            </div>
          </div>
        )}

        {activeTab === "New Request" && (
          <div className="grid gap-4 xl:grid-cols-[minmax(0,760px)_minmax(0,1fr)]">
            <div className="rounded-3xl border border-stone-300 bg-stone-50 p-6 shadow-sm shadow-stone-300/30">
              <div className="mb-5">
                <div className="text-sm font-semibold">Request form</div>
                <div className="text-xs text-stone-600">Requests are stamped with the logged-in account, and you can attach supporting documents before saving.</div>
              </div>
              <form onSubmit={submitRequest} className="grid gap-4">
                <Field
                  label="Job number"
                  value={requestForm.jobNumber}
                  onChange={(value) => setRequestForm((current) => ({ ...current, jobNumber: value }))}
                  placeholder="Internal request number"
                />
                <Field
                  label="Customer"
                  value={requestForm.customer}
                  onChange={(value) => setRequestForm((current) => ({ ...current, customer: value }))}
                  placeholder="Customer name"
                />
                <Field
                  label="Requestor name"
                  value={requestForm.requestorName}
                  onChange={(value) => setRequestForm((current) => ({ ...current, requestorName: value }))}
                  placeholder="Who is asking for it"
                />
                <div>
                  <div className="mb-2 text-sm font-medium text-stone-800">Description</div>
                  <textarea
                    value={requestForm.description}
                    onChange={(event) =>
                      setRequestForm((current) => ({ ...current, description: event.target.value }))
                    }
                    placeholder="Include QTY and any extra details here"
                    className="h-36 w-full rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm outline-none focus:border-emerald-800"
                  />
                </div>
                <div>
                  <div className="mb-2 text-sm font-medium text-stone-800">Documents</div>
                  <label className="block rounded-2xl border border-dashed border-stone-300 bg-white p-4 text-sm text-stone-700 hover:border-emerald-800">
                    <div className="font-medium text-stone-900">Upload PDF, Word, Excel, or other request files</div>
                    <input
                      type="file"
                      accept={ATTACHMENT_ACCEPT}
                      multiple
                      onChange={handleDraftAttachmentChange}
                      className="mt-3 block w-full text-xs"
                    />
                  </label>
                </div>
                <AttachmentList attachments={requestDraftAttachments} onRemove={removeDraftAttachment} />
                <div className="flex gap-2">
                  <button type="submit" className="rounded-2xl bg-emerald-900 px-4 py-3 text-sm font-medium text-white">
                    Save request
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                      setRequestForm(EMPTY_REQUEST_FORM);
                      setRequestDraftAttachments([]);
                    }}
                    className="rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm text-stone-800"
                  >
                    Reset
                  </button>
                </div>
              </form>
            </div>

            <div className="rounded-3xl border border-stone-300 bg-stone-50 p-6 shadow-sm shadow-stone-300/30">
              <div className="mb-4">
                <div className="text-sm font-semibold">Request workflow</div>
                <div className="text-xs text-stone-600">Open requests live in their own tab. Mark done records which logged-in user completed the request and when.</div>
              </div>
              <div className="grid gap-3">
                <StatRow label="Open requests" value={openRequests.length} />
                <StatRow label="Completed requests" value={requestHistory.length} />
                <StatRow label="Signed-in account" value={currentUser.username} />
                {canAccessTab(currentUser, "Open Requests") && (
                  <button
                    onClick={() => setActiveTab("Open Requests")}
                    className="rounded-2xl border border-stone-300 bg-white px-4 py-3 text-left text-sm text-stone-800"
                  >
                    Open the request queue
                  </button>
                )}
              </div>
            </div>
          </div>
        )}

        {activeTab === "Open Requests" && (
          <div className="rounded-3xl border border-stone-300 bg-stone-50 p-6 shadow-sm shadow-stone-300/30">
            <div className="mb-4 flex flex-col gap-2 md:flex-row md:items-end md:justify-between">
              <div>
                <div className="text-sm font-semibold">Open requests</div>
                <div className="text-xs text-stone-600">Mark done to move a request into history under your login. Uploads added here stay attached to the request.</div>
              </div>
              <button
                onClick={() => setActiveTab("New Request")}
                className="rounded-2xl bg-emerald-900 px-4 py-2 text-sm font-medium text-white"
              >
                New request
              </button>
            </div>

            <div className="space-y-3">
                  {openRequests.map((request) => (
                    <RequestCard
                      key={request.id}
                      request={request}
                      onDone={() => markRequestDone(request.id)}
                      onDelete={() => deleteRequest(request.id)}
                      onAddAttachments={(files) => addRequestAttachments(request.id, files)}
                      onRemoveAttachment={(attachmentId) => removeRequestAttachment(request.id, attachmentId)}
                      readOnly={!userCanEdit}
                    />
                  ))}
              {!openRequests.length && (
                <div className="rounded-2xl border border-dashed border-stone-300 bg-white/60 p-5 text-sm text-stone-600">
                  No open requests right now.
                </div>
              )}
            </div>
          </div>
        )}

        {activeTab === "Request History" && (
          <div className="rounded-3xl border border-stone-300 bg-stone-50 p-6 shadow-sm shadow-stone-300/30">
            <div className="mb-4 flex flex-col gap-3 md:flex-row md:items-end md:justify-between">
              <div>
                <div className="text-sm font-semibold">Request history</div>
                <div className="text-xs text-stone-600">Completed requests are sorted by completion time and show which login completed them.</div>
              </div>
              <div className="flex items-end gap-2">
                <div>
                  <div className="mb-1 text-xs font-semibold uppercase tracking-[0.16em] text-stone-600">Filter date</div>
                  <input
                    type="date"
                    value={requestHistoryFilterDate}
                    onChange={(event) => setRequestHistoryFilterDate(event.target.value)}
                    className="rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm outline-none focus:border-emerald-800"
                  />
                </div>
                <button
                  onClick={() => setRequestHistoryFilterDate("")}
                  className="rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm text-stone-800"
                >
                  Clear
                </button>
              </div>
            </div>
            <div className="space-y-3">
              {requestHistory.map((request) => (
                <RequestCard key={request.id} request={request} readOnly />
              ))}
              {!requestHistory.length && (
                <div className="rounded-2xl border border-dashed border-stone-300 bg-white/60 p-5 text-sm text-stone-600">
                  No completed requests yet.
                </div>
              )}
            </div>
          </div>
        )}

        {activeTab === "Pull Paper Request" && (
          <div className="grid gap-4 xl:grid-cols-[420px_minmax(0,1fr)]">
            <div className="rounded-3xl border border-stone-300 bg-stone-50 p-6 shadow-sm shadow-stone-300/30">
              <div className="mb-5">
                <div className="text-sm font-semibold">Pull paper request</div>
                <div className="text-xs text-stone-600">Send a paper pull note to a press or digital and keep the open list here.</div>
              </div>
              <form onSubmit={submitPullPaperRequest} className="grid gap-4">
                <div>
                  <div className="mb-2 text-sm font-medium text-stone-800">Send to</div>
                  <select
                    value={pullPaperForm.target}
                    onChange={(event) => setPullPaperForm((current) => ({ ...current, target: event.target.value }))}
                    className="w-full rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm outline-none focus:border-emerald-800"
                  >
                    {PULL_PAPER_TARGETS.map((target) => (
                      <option key={target} value={target}>
                        {target}
                      </option>
                    ))}
                  </select>
                </div>
                <div>
                  <div className="mb-2 text-sm font-medium text-stone-800">Paper note</div>
                  <textarea
                    value={pullPaperForm.details}
                    onChange={(event) => setPullPaperForm((current) => ({ ...current, details: event.target.value }))}
                    placeholder='Stock 266 Width 9.5" 1 roll'
                    className="h-40 w-full rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm outline-none focus:border-emerald-800"
                  />
                </div>
                <div className="rounded-2xl bg-stone-200/70 px-4 py-3 text-sm text-stone-800">
                  Requesting as <span className="font-semibold">{currentUser.username}</span>
                </div>
                <div className="flex gap-2">
                  <button type="submit" className="rounded-2xl bg-emerald-900 px-4 py-3 text-sm font-medium text-white">
                    Save pull request
                  </button>
                  <button
                    type="button"
                    onClick={() => setPullPaperForm(EMPTY_PULL_PAPER_FORM)}
                    className="rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm text-stone-800"
                  >
                    Reset
                  </button>
                </div>
              </form>
            </div>

            <div className="space-y-4">
              <div className="rounded-3xl border border-stone-300 bg-stone-50 p-6 shadow-sm shadow-stone-300/30">
                <div className="mb-4 flex items-center justify-between">
                  <div>
                    <div className="text-sm font-semibold">Open pull paper requests</div>
                    <div className="text-xs text-stone-600">Mark done when the paper is pulled, or delete it if it is no longer needed.</div>
                  </div>
                  <div className="rounded-xl bg-stone-200 px-2 py-1 text-xs text-stone-700">{openPullPaperRequests.length}</div>
                </div>
                <div className="space-y-3">
                  {openPullPaperRequests.map((request) => (
                    <PaperPullCard
                      key={request.id}
                      request={request}
                      onDone={() => markPullPaperRequestDone(request.id)}
                      onDelete={() => deletePullPaperRequest(request.id)}
                      readOnly={!userCanEdit}
                    />
                  ))}
                  {!openPullPaperRequests.length && (
                    <div className="rounded-2xl border border-dashed border-stone-300 bg-white/60 p-5 text-sm text-stone-600">
                      No open pull paper requests right now.
                    </div>
                  )}
                </div>
              </div>

              <div className="rounded-3xl border border-stone-300 bg-stone-50 p-6 shadow-sm shadow-stone-300/30">
                <div className="mb-4 flex items-center justify-between">
                  <div>
                    <div className="text-sm font-semibold">Completed pull paper requests</div>
                    <div className="text-xs text-stone-600">Recent completed paper pulls stay here so you can see who closed them out.</div>
                  </div>
                  <div className="rounded-xl bg-stone-200 px-2 py-1 text-xs text-stone-700">{completedPullPaperRequests.length}</div>
                </div>
                <div className="space-y-3">
                  {completedPullPaperRequests.slice(0, 12).map((request) => (
                    <PaperPullCard key={request.id} request={request} readOnly />
                  ))}
                  {!completedPullPaperRequests.length && (
                    <div className="rounded-2xl border border-dashed border-stone-300 bg-white/60 p-5 text-sm text-stone-600">
                      No completed paper pull requests yet.
                    </div>
                  )}
                </div>
              </div>
            </div>
          </div>
        )}

        {activeTab === "Supplies Request" && (
          <div className="grid gap-4 xl:grid-cols-[420px_minmax(0,1fr)]">
            <div className="rounded-3xl border border-stone-300 bg-stone-50 p-6 shadow-sm shadow-stone-300/30">
              <div className="mb-5">
                <div className="text-sm font-semibold">Supplies request</div>
                <div className="text-xs text-stone-600">Create a supplies request, attach supporting files if needed, and keep the open count visible in the tab.</div>
              </div>
              <form onSubmit={submitSuppliesRequest} className="grid gap-4">
                <div>
                  <div className="mb-2 text-sm font-medium text-stone-800">Supplies note</div>
                  <textarea
                    value={suppliesForm.details}
                    onChange={(event) => setSuppliesForm((current) => ({ ...current, details: event.target.value }))}
                    placeholder="Boxes, labels, cores, bags, tape, or any other supply request details"
                    className="h-40 w-full rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm outline-none focus:border-emerald-800"
                  />
                </div>
                <div>
                  <div className="mb-2 text-sm font-medium text-stone-800">Documents</div>
                  <label className="block rounded-2xl border border-dashed border-stone-300 bg-white p-4 text-sm text-stone-700 hover:border-emerald-800">
                    <div className="font-medium text-stone-900">Upload PDF, PNG, JPG, JPEG, or other supply request files</div>
                    <input
                      type="file"
                      accept={ATTACHMENT_ACCEPT}
                      multiple
                      onChange={handleSuppliesDraftAttachmentChange}
                      className="mt-3 block w-full text-xs"
                    />
                  </label>
                </div>
                <AttachmentList attachments={suppliesDraftAttachments} onRemove={removeSuppliesDraftAttachment} />
                <div className="rounded-2xl bg-stone-200/70 px-4 py-3 text-sm text-stone-800">
                  Requesting as <span className="font-semibold">{currentUser.username}</span>
                </div>
                <div className="flex gap-2">
                  <button type="submit" className="rounded-2xl bg-emerald-900 px-4 py-3 text-sm font-medium text-white">
                    Save supplies request
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                      setSuppliesForm(EMPTY_SUPPLIES_FORM);
                      setSuppliesDraftAttachments([]);
                    }}
                    className="rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm text-stone-800"
                  >
                    Reset
                  </button>
                </div>
              </form>
            </div>

            <div className="space-y-4">
              <div className="rounded-3xl border border-stone-300 bg-stone-50 p-6 shadow-sm shadow-stone-300/30">
                <div className="mb-4 flex items-center justify-between">
                  <div>
                    <div className="text-sm font-semibold">Open supplies requests</div>
                    <div className="text-xs text-stone-600">Mark done when the supplies request is complete, or delete it if it is no longer needed.</div>
                  </div>
                  <div className="rounded-xl bg-stone-200 px-2 py-1 text-xs text-stone-700">{openSuppliesRequests.length}</div>
                </div>
                <div className="space-y-3">
                  {openSuppliesRequests.map((request) => (
                    <SuppliesRequestCard
                      key={request.id}
                      request={request}
                      onDone={() => markSuppliesRequestDone(request.id)}
                      onDelete={() => deleteSuppliesRequest(request.id)}
                      onAddAttachments={(files) => addSuppliesRequestAttachments(request.id, files)}
                      onRemoveAttachment={(attachmentId) => removeSuppliesRequestAttachment(request.id, attachmentId)}
                      readOnly={!userCanEdit}
                    />
                  ))}
                  {!openSuppliesRequests.length && (
                    <div className="rounded-2xl border border-dashed border-stone-300 bg-white/60 p-5 text-sm text-stone-600">
                      No open supplies requests right now.
                    </div>
                  )}
                </div>
              </div>

              <div className="rounded-3xl border border-stone-300 bg-stone-50 p-6 shadow-sm shadow-stone-300/30">
                <div className="mb-4 flex items-center justify-between">
                  <div>
                    <div className="text-sm font-semibold">Completed supplies requests</div>
                    <div className="text-xs text-stone-600">Recent completed supplies requests stay here so you can see who closed them out.</div>
                  </div>
                  <div className="rounded-xl bg-stone-200 px-2 py-1 text-xs text-stone-700">{completedSuppliesRequests.length}</div>
                </div>
                <div className="space-y-3">
                  {completedSuppliesRequests.slice(0, 12).map((request) => (
                    <SuppliesRequestCard key={request.id} request={request} readOnly />
                  ))}
                  {!completedSuppliesRequests.length && (
                    <div className="rounded-2xl border border-dashed border-stone-300 bg-white/60 p-5 text-sm text-stone-600">
                      No completed supplies requests yet.
                    </div>
                  )}
                </div>
              </div>
            </div>
          </div>
        )}

        {activeTab === "Daily Shipment" && (
          <div className="space-y-4">
            <div className="rounded-3xl border border-stone-300 bg-stone-50 p-5 shadow-sm shadow-stone-300/30">
              <div className="flex flex-col gap-3 md:flex-row md:items-end md:justify-between">
                <div>
                  <div className="text-sm font-semibold">Daily shipment</div>
                  <div className="text-xs text-stone-600">
                    Finished jobs can be assigned to a ship date first, then grouped under that shipping day.
                  </div>
                  <div className="mt-2 text-xs text-stone-700">
                    {shipmentEmailsForSelectedDate.length
                      ? `${shipmentEmailsForSelectedDate.length} shipment email${shipmentEmailsForSelectedDate.length === 1 ? "" : "s"} already logged for this date.`
                      : "No shipment email has been logged for this date yet."}
                  </div>
                </div>
                <div>
                  <div className="mb-1 text-xs font-semibold uppercase tracking-[0.16em] text-stone-600">Ship date</div>
                  <input
                    type="date"
                    value={selectedShipDate}
                    onChange={(event) => setSelectedShipDate(event.target.value)}
                    className="rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm outline-none focus:border-emerald-800"
                  />
                </div>
              </div>
            </div>

            <div className="rounded-3xl border border-stone-300 bg-stone-50 p-5 shadow-sm shadow-stone-300/30">
              <div className="mb-4 flex flex-col gap-3 md:flex-row md:items-end md:justify-between">
                <div>
                  <div className="text-sm font-semibold">Assign ship date</div>
                  <div className="text-xs text-stone-600">
                    Pick the finished jobs that have not been grouped yet, assign them to a ship date, and they will move into that day's shipment queue.
                  </div>
                </div>
                <div className="flex flex-wrap items-end gap-2">
                  <div>
                    <div className="mb-1 text-xs font-semibold uppercase tracking-[0.16em] text-stone-600">Ship date</div>
                    <input
                      type="date"
                      value={shipDateDraft}
                      onChange={(event) => setShipDateDraft(event.target.value)}
                      className="rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm outline-none focus:border-emerald-800"
                    />
                  </div>
                  <button
                    type="button"
                    onClick={() => setSelectedShipQueueJobs(unassignedFinishedJobs.map((job) => job.id))}
                    className="rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm text-stone-800"
                  >
                    Select all
                  </button>
                  <button
                    type="button"
                    onClick={() => setSelectedShipQueueJobs([])}
                    className="rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm text-stone-800"
                  >
                    Clear
                  </button>
                  <button
                    type="button"
                    onClick={assignShipDateToFinishedJobs}
                    className="rounded-2xl bg-emerald-900 px-4 py-2 text-sm font-medium text-white"
                  >
                    Assign date
                  </button>
                </div>
              </div>

              <div className="mb-3 flex items-center justify-between">
                <div className="text-xs text-stone-600">
                  {selectedShipQueueJobs.length} selected
                </div>
                <div className="rounded-xl bg-stone-200 px-2 py-1 text-xs text-stone-700">{unassignedFinishedJobs.length}</div>
              </div>

              <div className="max-h-[320px] space-y-3 overflow-y-auto pr-1">
                {unassignedFinishedJobs.map((job) => (
                  <label key={job.id} className="flex gap-3 rounded-2xl border border-stone-300 bg-white p-3">
                    <input
                      type="checkbox"
                      checked={selectedShipQueueJobs.includes(job.id)}
                      onChange={() => toggleShipQueueJob(job.id)}
                      className="mt-1 h-4 w-4"
                    />
                    <div className="min-w-0 flex-1">
                      <div className="flex flex-wrap items-start justify-between gap-2">
                        <div>
                          <div className="text-sm font-semibold">
                            {job.customerName} {job.number}
                          </div>
                          <div className="mt-1 text-xs text-stone-700">{job.generalDescr}</div>
                        </div>
                        <span className={`rounded-full px-2 py-1 text-[11px] font-medium ${statusTone("done")}`}>done</span>
                      </div>
                      <div className="mt-2 grid grid-cols-2 gap-2 text-xs text-stone-700 md:grid-cols-4">
                        <InfoPill label="Press" value={job.press || "-"} />
                        <InfoPill label="Qty" value={job.ticQuantity.toLocaleString()} />
                        <InfoPill label="Finished" value={formatDateTime(job.finishMeta?.finishedAt)} />
                        <InfoPill label="Ship date" value={effectiveFinishedShipDate(job.finishMeta) || "-"} />
                      </div>
                    </div>
                  </label>
                ))}
                {!unassignedFinishedJobs.length && (
                  <div className="rounded-2xl border border-dashed border-stone-300 bg-white/60 p-4 text-sm text-stone-600">
                    All finished jobs are already grouped into shipments.
                  </div>
                )}
              </div>
            </div>

            <div className="grid gap-4 xl:grid-cols-[minmax(0,1.1fr)_420px]">
              <div className="rounded-3xl border border-stone-300 bg-stone-50 p-5 shadow-sm shadow-stone-300/30">
                <div className="mb-4 flex items-center justify-between">
                  <div>
                    <div className="text-sm font-semibold">Ready to ship on {selectedShipDate}</div>
                    <div className="text-xs text-stone-600">Select one or more jobs, then create a shipment group.</div>
                  </div>
                  <div className="rounded-xl bg-stone-200 px-2 py-1 text-xs text-stone-700">{readyToShipJobs.length}</div>
                </div>

                <div className="space-y-3">
                  {readyToShipJobs.map((job) => (
                    <label key={job.id} className="flex gap-3 rounded-2xl border border-stone-300 bg-white p-3">
                      <input
                        type="checkbox"
                        checked={selectedShipmentJobs.includes(job.id)}
                        onChange={() => toggleShipmentJob(job.id)}
                        className="mt-1 h-4 w-4"
                      />
                      <div className="min-w-0 flex-1">
                        <div className="flex flex-wrap items-start justify-between gap-2">
                          <div>
                            <div className="text-sm font-semibold">
                              {job.customerName} {job.number}
                            </div>
                            <div className="mt-1 text-xs text-stone-700">{job.generalDescr}</div>
                          </div>
                          <span className={`rounded-full px-2 py-1 text-[11px] font-medium ${statusTone("ship")}`}>ship</span>
                        </div>
                        <div className="mt-2 grid grid-cols-2 gap-2 text-xs text-stone-700 md:grid-cols-4">
                          <InfoPill label="Press" value={job.press || "-"} />
                          <InfoPill label="Qty" value={job.ticQuantity.toLocaleString()} />
                          <InfoPill label="Finished" value={formatDateTime(job.finishMeta?.finishedAt)} />
                          <InfoPill label="By" value={job.finishMeta?.finishedBy || "-"} />
                        </div>
                      </div>
                    </label>
                  ))}
                  {!readyToShipJobs.length && (
                    <div className="rounded-2xl border border-dashed border-stone-300 bg-white/60 p-4 text-sm text-stone-600">
                      No user-finished jobs are waiting to ship on this date.
                    </div>
                  )}
                </div>
              </div>

              <div className="space-y-4">
                <div className="rounded-3xl border border-stone-300 bg-stone-50 p-5 shadow-sm shadow-stone-300/30">
                  <div className="mb-4">
                    <div className="text-sm font-semibold">Create shipment group</div>
                    <div className="text-xs text-stone-600">Example: one skid with 3 jobs for $141.17, separate FedEx transactions, or a custom method like JP Express.</div>
                  </div>
                  <div className="mb-4 rounded-2xl border border-stone-300 bg-white p-4">
                    <div className="mb-3 text-sm font-semibold">Shipping methods</div>
                    <div className="mb-3 flex gap-2">
                      <input
                        type="text"
                        value={newShipmentMethod}
                        onChange={(event) => setNewShipmentMethod(event.target.value)}
                        placeholder="Add a method like JP Express"
                        className="flex-1 rounded-2xl border border-stone-300 bg-stone-50 px-4 py-3 text-sm outline-none focus:border-emerald-800"
                      />
                      <button
                        type="button"
                        onClick={addShipmentMethod}
                        className="rounded-2xl bg-emerald-900 px-4 py-3 text-sm font-medium text-white"
                      >
                        Add
                      </button>
                    </div>
                    <div className="flex flex-wrap gap-2">
                      {shipmentMethods.map((method) => (
                        <div key={method} className="flex items-center gap-2 rounded-full border border-stone-300 bg-stone-100 px-3 py-2 text-xs text-stone-800">
                          <span>{method}</span>
                          <button
                            type="button"
                            onClick={() => removeShipmentMethod(method)}
                            className="text-rose-700"
                          >
                            remove
                          </button>
                        </div>
                      ))}
                    </div>
                  </div>
                  <form onSubmit={createShipmentGroup} className="grid gap-3">
                    <Field
                      label="Label"
                      value={shipmentForm.label}
                      onChange={(value) => setShipmentForm((current) => ({ ...current, label: value }))}
                      placeholder="Skid A or FedEx 1"
                    />
                    <div>
                      <div className="mb-2 text-sm font-medium text-stone-800">Method</div>
                      <select
                        value={shipmentForm.method}
                        onChange={(event) => setShipmentForm((current) => ({ ...current, method: event.target.value }))}
                        className="w-full rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm outline-none focus:border-emerald-800"
                      >
                        {shipmentMethods.map((method) => (
                          <option key={method} value={method}>
                            {method}
                          </option>
                        ))}
                      </select>
                    </div>
                    <div className="grid grid-cols-[minmax(0,1fr)_160px] gap-3">
                      <Field
                        label="Skids / cartons out"
                        value={shipmentForm.packageCount}
                        onChange={(value) => setShipmentForm((current) => ({ ...current, packageCount: value }))}
                        placeholder="3"
                      />
                      <div>
                        <div className="mb-2 text-sm font-medium text-stone-800">Type</div>
                        <select
                          value={shipmentForm.packageType}
                          onChange={(event) =>
                            setShipmentForm((current) => ({ ...current, packageType: event.target.value }))
                          }
                          className="w-full rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm outline-none focus:border-emerald-800"
                        >
                          <option>Skids</option>
                          <option>Cartons</option>
                        </select>
                      </div>
                    </div>
                    <Field
                      label="Total price"
                      value={shipmentForm.totalCost}
                      onChange={(value) => setShipmentForm((current) => ({ ...current, totalCost: value }))}
                      placeholder="141.17"
                    />
                    <Field
                      label="Bill customer"
                      value={shipmentForm.billAmount}
                      onChange={(value) => setShipmentForm((current) => ({ ...current, billAmount: value }))}
                      placeholder="185.00"
                    />
                    <div>
                      <div className="mb-2 text-sm font-medium text-stone-800">Notes</div>
                      <textarea
                        value={shipmentForm.notes}
                        onChange={(event) => setShipmentForm((current) => ({ ...current, notes: event.target.value }))}
                        placeholder="Optional shipment notes"
                        className="h-24 w-full rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm outline-none focus:border-emerald-800"
                      />
                    </div>
                    <div>
                      <div className="mb-2 text-sm font-medium text-stone-800">Documents</div>
                      <label className="block rounded-2xl border border-dashed border-stone-300 bg-white p-4 text-sm text-stone-700 hover:border-emerald-800">
                        <div className="font-medium text-stone-900">Upload PDF, PNG, JPG, JPEG, or other shipment files</div>
                        <input
                          type="file"
                          accept={ATTACHMENT_ACCEPT}
                          multiple
                          onChange={handleShipmentDraftAttachmentChange}
                          className="mt-3 block w-full text-xs"
                        />
                      </label>
                    </div>
                    <AttachmentList attachments={shipmentDraftAttachments} onRemove={removeShipmentDraftAttachment} />
                    <button type="submit" className="rounded-2xl bg-emerald-900 px-4 py-3 text-sm font-medium text-white">
                      Create shipment group
                    </button>
                  </form>
                </div>

                <div className="rounded-3xl border border-stone-300 bg-stone-50 p-5 shadow-sm shadow-stone-300/30">
                  <div className="mb-4">
                    <div className="text-sm font-semibold">Daily shipment email</div>
                    <div className="text-xs text-stone-600">Draft the summary email for this ship date, then log it so everyone can see it was already sent.</div>
                  </div>
                  <div className="grid gap-3">
                    <Field
                      label="Recipients"
                      value={shipmentEmailForm.recipients}
                      onChange={(value) => setShipmentEmailForm({ recipients: value })}
                      placeholder="shipping@company.com; billing@company.com"
                    />
                    <div className="rounded-2xl bg-stone-100 p-4 text-sm text-stone-800">
                      <div className="font-semibold">{buildShipmentEmailDraft().subject}</div>
                      <pre className="mt-2 max-h-52 overflow-auto whitespace-pre-wrap text-xs text-stone-700">{buildShipmentEmailDraft().body}</pre>
                    </div>
                    <div className="flex flex-wrap gap-2">
                      <button
                        type="button"
                        disabled={!userCanEdit}
                        onClick={openShipmentEmailDraft}
                        className="rounded-2xl bg-emerald-900 px-4 py-3 text-sm font-medium text-white disabled:cursor-not-allowed disabled:opacity-50"
                      >
                        Open email draft
                      </button>
                      <button
                        type="button"
                        disabled={!userCanEdit}
                        onClick={logShipmentEmail}
                        className="rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm text-stone-800 disabled:cursor-not-allowed disabled:opacity-50"
                      >
                        Mark emailed
                      </button>
                    </div>
                  </div>
                </div>

                <div className="rounded-3xl border border-stone-300 bg-stone-50 p-5 shadow-sm shadow-stone-300/30">
                  <div className="mb-4">
                    <div className="text-sm font-semibold">Shipment history</div>
                    <div className="text-xs text-stone-600">Pick a past date to see the exact jobs that shipped that day.</div>
                  </div>
                  <div className="mb-4 flex items-end gap-2">
                    <div className="flex-1">
                      <div className="mb-1 text-xs font-semibold uppercase tracking-[0.16em] text-stone-600">Filter date</div>
                      <input
                        type="date"
                        value={shipmentHistoryFilterDate}
                        onChange={(event) => setShipmentHistoryFilterDate(event.target.value)}
                        className="w-full rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm outline-none focus:border-emerald-800"
                      />
                    </div>
                    <button
                      onClick={() => setShipmentHistoryFilterDate("")}
                      className="rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm text-stone-800"
                    >
                      Clear
                    </button>
                  </div>
                  <div className="space-y-3">
                    {shipmentHistoryDays.map((item) => (
                      <button
                        key={item.shipDate}
                        onClick={() => setSelectedShipDate(item.shipDate)}
                        className={`w-full rounded-2xl border px-4 py-3 text-left ${
                          selectedShipDate === item.shipDate
                            ? "border-emerald-900 bg-emerald-900 text-white"
                            : "border-stone-300 bg-white text-stone-800"
                        }`}
                      >
                        <div className="text-sm font-semibold">{item.shipDate}</div>
                        <div className={`mt-1 text-xs ${selectedShipDate === item.shipDate ? "text-stone-200" : "text-stone-600"}`}>
                          {item.groupCount} groups - {item.jobCount} jobs - {formatCurrency(item.totalCost)}
                        </div>
                      </button>
                    ))}
                    {!shipmentHistoryDays.length && (
                      <div className="rounded-2xl border border-dashed border-stone-300 bg-white/60 p-4 text-sm text-stone-600">
                        No shipment history yet.
                      </div>
                    )}
                  </div>
                </div>
              </div>
            </div>

            <div className="rounded-3xl border border-stone-300 bg-stone-50 p-5 shadow-sm shadow-stone-300/30">
              <div className="mb-4">
                <div className="text-sm font-semibold">Shipped on {selectedShipDate}</div>
                <div className="text-xs text-stone-600">Each shipment group keeps the exact jobs and who marked them finished.</div>
              </div>
              <div className="space-y-3">
                {shipmentGroupsForDay.map((group) => {
                  const items = getShipmentItems(group, jobMap, finishedMetaByJobId);
                  return (
                    <div key={group.id} className="rounded-2xl border border-stone-300 bg-white p-4">
                      <div className="flex flex-col gap-3 md:flex-row md:items-start md:justify-between">
                        <div>
                          <div className="flex flex-wrap items-center gap-2">
                            <div className="text-sm font-semibold">{group.label}</div>
                            <span className="rounded-full bg-emerald-900 px-2 py-1 text-[11px] font-medium text-white">
                              {group.method}
                            </span>
                          </div>
                          <div className="mt-1 text-xs text-stone-600">
                            {items.length} jobs - {formatCurrency(group.totalCost)} - created {formatDateTime(group.createdAt)}
                            {group.createdBy ? ` by ${group.createdBy}` : ""}
                          </div>
                          <div className="mt-1 text-xs text-stone-600">
                            Bill {formatCurrency(group.billAmount)}
                          </div>
                          {!!group.packageCount && (
                            <div className="mt-1 text-xs text-stone-600">
                              {group.packageCount} {safeText(group.packageType || "Skids").toLowerCase()}
                            </div>
                          )}
                          {group.notes && <div className="mt-2 text-sm text-stone-800">{group.notes}</div>}
                          <div className="mt-3">
                            <div className="mb-2 text-xs font-semibold uppercase tracking-[0.16em] text-stone-600">Documents</div>
                            <AttachmentList
                              attachments={group.attachments || []}
                              onRemove={(attachmentId) => removeShipmentGroupAttachment(group.id, attachmentId)}
                            />
                          </div>
                          <label className="mt-3 block rounded-2xl border border-dashed border-stone-300 bg-stone-50 p-4 text-sm text-stone-700 hover:border-emerald-800">
                            <div className="font-medium text-stone-900">Add more shipment files</div>
                            <input
                              type="file"
                              accept={ATTACHMENT_ACCEPT}
                              multiple
                              onChange={(event) => {
                                const files = event.target.files;
                                if (files?.length) addShipmentGroupAttachments(group.id, files);
                                event.target.value = "";
                              }}
                              className="mt-3 block w-full text-xs"
                            />
                          </label>
                        </div>
                        <button
                          onClick={() => deleteShipmentGroup(group.id)}
                          className="rounded-2xl border border-rose-200 px-3 py-2 text-sm text-rose-700"
                        >
                          Delete group
                        </button>
                      </div>
                      <div className="mt-3 grid gap-2 md:grid-cols-2">
                        {items.map((item) => (
                            <div key={item.id} className="rounded-2xl bg-stone-100 p-3">
                            <div className="text-sm font-semibold">
                              {item.customerName} {item.number}
                            </div>
                              <div className="mt-1 text-xs text-stone-700">{item.generalDescr}</div>
                              <div className="mt-2 text-[11px] text-stone-600">
                              Finished {formatDateTime(item.finishedAt)} by {item.finishedBy || "-"}
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>
                  );
                })}
                {!shipmentGroupsForDay.length && (
                  <div className="rounded-2xl border border-dashed border-stone-300 bg-white/60 p-4 text-sm text-stone-600">
                    No shipment groups created for this date yet.
                  </div>
                )}
              </div>
            </div>
          </div>
        )}

        {activeTab === "Shipment Emails" && (
          <div className="rounded-3xl border border-stone-300 bg-stone-50 p-6 shadow-sm shadow-stone-300/30">
            <div className="mb-4 flex flex-col gap-3 md:flex-row md:items-end md:justify-between">
              <div>
                <div className="text-sm font-semibold">Shipment email log</div>
                <div className="text-xs text-stone-600">Use this to confirm which shipment dates were already emailed so you do not double-send them.</div>
              </div>
              <div className="flex items-end gap-2">
                <div>
                  <div className="mb-1 text-xs font-semibold uppercase tracking-[0.16em] text-stone-600">Filter date</div>
                  <input
                    type="date"
                    value={shipmentEmailHistoryFilterDate}
                    onChange={(event) => setShipmentEmailHistoryFilterDate(event.target.value)}
                    className="rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm outline-none focus:border-emerald-800"
                  />
                </div>
                <button
                  onClick={() => setShipmentEmailHistoryFilterDate("")}
                  className="rounded-2xl border border-stone-300 bg-white px-3 py-2 text-sm text-stone-800"
                >
                  Clear
                </button>
              </div>
            </div>
            <div className="space-y-3">
              {shipmentEmailHistory.map((log) => (
                <div key={log.id} className="rounded-2xl border border-stone-300 bg-white p-4">
                  <div className="flex flex-col gap-3 md:flex-row md:items-start md:justify-between">
                    <div className="min-w-0 flex-1">
                      <div className="flex flex-wrap items-center gap-2">
                        <div className="text-sm font-semibold">{log.shipDate}</div>
                        <span className="rounded-full bg-stone-200 px-2 py-1 text-[11px] font-medium text-stone-800">
                          {log.jobCount} jobs
                        </span>
                      </div>
                      <div className="mt-1 text-xs text-stone-600">
                        Sent by {log.createdBy || "-"} on {formatDateTime(log.createdAt)}
                      </div>
                      <div className="mt-1 text-xs text-stone-600">Recipients: {log.recipients || "-"}</div>
                      <div className="mt-1 text-xs text-stone-600">
                        Methods: {log.methods.length ? log.methods.join(", ") : "-"} · Cost {formatCurrency(log.totalCost)} · Bill {formatCurrency(log.totalBill)}
                      </div>
                      <div className="mt-3 rounded-2xl bg-stone-100 p-3">
                        <div className="text-xs font-semibold uppercase tracking-[0.16em] text-stone-600">Subject</div>
                        <div className="mt-1 text-sm text-stone-800">{log.subject}</div>
                        <div className="mt-3 text-xs font-semibold uppercase tracking-[0.16em] text-stone-600">Body</div>
                        <pre className="mt-1 whitespace-pre-wrap text-xs text-stone-700">{log.body}</pre>
                      </div>
                    </div>
                    <button
                      onClick={() => {
                        setSelectedShipDate(log.shipDate);
                        setActiveTab("Daily Shipment");
                      }}
                      className="rounded-2xl border border-stone-300 bg-stone-50 px-3 py-2 text-sm text-stone-800"
                    >
                      Open date
                    </button>
                  </div>
                </div>
              ))}
              {!shipmentEmailHistory.length && (
                <div className="rounded-2xl border border-dashed border-stone-300 bg-white/60 p-5 text-sm text-stone-600">
                  No shipment emails have been logged yet.
                </div>
              )}
            </div>
          </div>
        )}

        {activeTab === "User Admin" && userCanManageUsers && (
          <div className="grid gap-4 xl:grid-cols-[420px_minmax(0,1fr)]">
            <div className="rounded-3xl border border-stone-300 bg-stone-50 p-6 shadow-sm shadow-stone-300/30">
              <div className="mb-5">
                <div className="text-sm font-semibold">Add user</div>
                <div className="text-xs text-stone-600">Management can create logins, choose edit or view-only access, and decide which tabs each user can see.</div>
              </div>
              <form onSubmit={createUser} className="grid gap-4">
                <Field
                  label="Username"
                  value={userForm.username}
                  onChange={(value) => setUserForm((current) => ({ ...current, username: value }))}
                  placeholder="New username"
                />
                <Field
                  label="Password"
                  value={userForm.password}
                  onChange={(value) => setUserForm((current) => ({ ...current, password: value }))}
                  placeholder="Set a password"
                />
                <div>
                  <div className="mb-2 text-sm font-medium text-stone-800">Access</div>
                  <select
                    value={userForm.accessMode}
                    onChange={(event) => setUserForm((current) => ({ ...current, accessMode: event.target.value }))}
                    className="w-full rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm outline-none focus:border-emerald-800"
                  >
                    {ACCESS_MODE_OPTIONS.map((mode) => (
                      <option key={mode} value={mode}>
                        {mode === "edit" ? "Edit" : "View only"}
                      </option>
                    ))}
                  </select>
                </div>
                <label className="flex items-center gap-3 rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm text-stone-800">
                  <input
                    type="checkbox"
                    checked={userForm.canManageUsers}
                    onChange={(event) => setUserForm((current) => ({ ...current, canManageUsers: event.target.checked }))}
                    className="h-4 w-4"
                  />
                  <span>Allow user admin access</span>
                </label>
                <div>
                  <div className="mb-2 text-sm font-medium text-stone-800">Visible tabs</div>
                  <div className="grid gap-2 md:grid-cols-2">
                    {[...BASE_TABS, "User Admin"].map((tab) => (
                      <label key={tab} className="flex items-center gap-3 rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm text-stone-800">
                        <input
                          type="checkbox"
                          checked={userForm.canManageUsers && tab === "User Admin" ? true : userForm.tabs.includes(tab)}
                          disabled={tab === "User Admin"}
                          onChange={() =>
                            setUserForm((current) => {
                              const nextTabs = current.tabs.includes(tab)
                                ? current.tabs.filter((value) => value !== tab)
                                : [...current.tabs, tab];
                              return { ...current, tabs: normalizeUserTabs(nextTabs, "Warehouse/Shipper", false, current.canManageUsers) };
                            })
                          }
                          className="h-4 w-4"
                        />
                        <span>{tab}</span>
                      </label>
                    ))}
                  </div>
                </div>
                <button type="submit" className="rounded-2xl bg-emerald-900 px-4 py-3 text-sm font-medium text-white">
                  Create user
                </button>
              </form>
            </div>

            <div className="space-y-4">
              <div className="rounded-3xl border border-stone-300 bg-stone-50 p-6 shadow-sm shadow-stone-300/30">
                <div className="mb-5 flex items-center justify-between gap-3">
                  <div>
                    <div className="text-sm font-semibold">Registration approvals</div>
                    <div className="text-xs text-stone-600">Approve new account requests before they can log in.</div>
                  </div>
                  <div className="rounded-full bg-stone-200 px-3 py-1 text-xs font-semibold text-stone-800">
                    {pendingRegistrationRequests.length} pending
                  </div>
                </div>
                <div className="space-y-3">
                  {pendingRegistrationRequests.map((request) => (
                    <div key={request.id} className="rounded-2xl border border-stone-300 bg-white p-4">
                      <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
                        <div>
                          <div className="text-sm font-semibold">{request.username}</div>
                          <div className="mt-1 text-xs text-stone-600">
                            Requested {formatDateTime(request.createdAt)}
                          </div>
                        </div>
                        <div className="flex flex-wrap gap-2">
                          <button
                            onClick={() => approveRegistrationRequest(request.id)}
                            className="rounded-2xl bg-emerald-900 px-4 py-2 text-sm font-medium text-white"
                          >
                            Approve
                          </button>
                          <button
                            onClick={() => denyRegistrationRequest(request.id)}
                            className="rounded-2xl border border-stone-300 bg-white px-4 py-2 text-sm text-stone-800"
                          >
                            Deny
                          </button>
                        </div>
                      </div>
                    </div>
                  ))}
                  {!pendingRegistrationRequests.length && (
                    <div className="rounded-2xl border border-dashed border-stone-300 bg-white/60 p-4 text-sm text-stone-600">
                      No pending registration requests right now.
                    </div>
                  )}
                </div>
              </div>

              <div className="rounded-3xl border border-stone-300 bg-stone-50 p-6 shadow-sm shadow-stone-300/30">
                <div className="mb-5">
                  <div className="text-sm font-semibold">Manage users</div>
                  <div className="text-xs text-stone-600">Reset passwords, choose edit or view-only, and control which tabs each account can see.</div>
                </div>
                <div className="space-y-3">
                  {users.map((user) => (
                    <div key={user.id} className="rounded-2xl border border-stone-300 bg-white p-4">
                      <div className="space-y-4">
                        <div className="flex flex-col gap-3 md:flex-row md:items-start md:justify-between">
                          <div>
                            <div className="flex flex-wrap items-center gap-2">
                              <div className="text-sm font-semibold">{user.username}</div>
                              <span className={`rounded-full px-2 py-1 text-[11px] font-medium ${hasManagementAccess(user) ? "bg-emerald-900 text-white" : "bg-stone-200 text-stone-800"}`}>
                                {hasManagementAccess(user) ? "User Admin" : user.accessMode === "edit" ? "Edit" : "View only"}
                              </span>
                            </div>
                            <div className="mt-1 text-xs text-stone-600">
                              Created {formatDateTime(user.createdAt)} by {user.createdBy || "-"}
                            </div>
                          </div>
                          <div className="flex flex-col gap-2 md:w-[320px]">
                            <input
                              type="text"
                              value={userPasswordDrafts[user.id] || ""}
                              onChange={(event) =>
                                setUserPasswordDrafts((current) => ({ ...current, [user.id]: event.target.value }))
                              }
                              placeholder={`Set new password for ${user.username}`}
                              className="w-full rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm outline-none focus:border-emerald-800"
                            />
                            <button
                              onClick={() => updateUserPassword(user.id)}
                              className="rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm text-stone-800"
                            >
                              Save password
                            </button>
                          </div>
                        </div>

                        <div className="grid gap-3 md:grid-cols-[180px_1fr]">
                          <div>
                            <div className="mb-2 text-xs font-semibold uppercase tracking-[0.16em] text-stone-600">Access</div>
                            <select
                              value={user.accessMode}
                              onChange={(event) => updateUserAccess(user.id, { accessMode: event.target.value })}
                              className="w-full rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm outline-none focus:border-emerald-800"
                            >
                              {ACCESS_MODE_OPTIONS.map((mode) => (
                                <option key={mode} value={mode}>
                                  {mode === "edit" ? "Edit" : "View only"}
                                </option>
                              ))}
                            </select>
                          </div>
                          <label className="flex items-center gap-3 rounded-2xl border border-stone-300 bg-stone-50 px-4 py-3 text-sm text-stone-800">
                            <input
                              type="checkbox"
                              checked={!!user.canManageUsers}
                              onChange={(event) => updateUserAccess(user.id, { canManageUsers: event.target.checked })}
                              className="h-4 w-4"
                            />
                            <span>Can manage users and registration approvals</span>
                          </label>
                        </div>

                        <div>
                          <div className="mb-2 text-xs font-semibold uppercase tracking-[0.16em] text-stone-600">Visible tabs</div>
                          <div className="grid gap-2 md:grid-cols-2">
                            {[...BASE_TABS, "User Admin"].map((tab) => (
                              <label key={tab} className="flex items-center gap-3 rounded-2xl border border-stone-300 bg-stone-50 px-4 py-3 text-sm text-stone-800">
                                <input
                                  type="checkbox"
                                  checked={user.tabs.includes(tab)}
                                  disabled={tab === "User Admin"}
                                  onChange={() => toggleUserTab(user.id, tab)}
                                  className="h-4 w-4"
                                />
                                <span>{tab}</span>
                              </label>
                            ))}
                          </div>
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

function LoginScreen({
  authView,
  loginForm,
  loginError,
  registerForm,
  registerError,
  registerSuccess,
  users,
  onChangeLogin,
  onChangeRegister,
  onChangeView,
  onSubmitLogin,
  onSubmitRegister,
}) {
  return (
    <div className="min-h-screen bg-stone-100 p-6 text-stone-900">
      <div className="mx-auto max-w-xl rounded-[2rem] border border-stone-300 bg-gradient-to-br from-stone-50 via-white to-stone-100 p-8 shadow-sm shadow-stone-300/40">
        <div className="mb-6">
          <p className="text-xs font-semibold uppercase tracking-[0.18em] text-emerald-900">Secure Access</p>
          <h1 className="mt-1 text-3xl font-semibold tracking-tight">Label Traxx Scheduler login</h1>
          <p className="mt-2 text-sm text-stone-700">
            Sign in to open the scheduler, save your work, and stamp finished jobs and request history with your account.
          </p>
        </div>
        <div className="mb-5 grid grid-cols-2 gap-2 rounded-2xl bg-stone-200/70 p-1">
          <button
            type="button"
            onClick={() => onChangeView("login")}
            className={`rounded-xl px-4 py-2 text-sm font-medium transition ${
              authView === "login" ? "bg-white text-stone-900 shadow-sm" : "text-stone-700"
            }`}
          >
            Log in
          </button>
          <button
            type="button"
            onClick={() => onChangeView("register")}
            className={`rounded-xl px-4 py-2 text-sm font-medium transition ${
              authView === "register" ? "bg-white text-stone-900 shadow-sm" : "text-stone-700"
            }`}
          >
            Register
          </button>
        </div>
        {authView === "login" ? (
          <form onSubmit={onSubmitLogin} className="grid gap-4">
            <div>
              <div className="mb-2 text-sm font-medium text-stone-800">Username</div>
              <select
                value={loginForm.username}
                onChange={(event) => onChangeLogin((current) => ({ ...current, username: event.target.value }))}
                className="w-full rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm outline-none focus:border-emerald-800"
              >
                <option value="">Select a username</option>
                {users.map((user) => (
                  <option key={user.id} value={user.username}>
                    {user.username}
                  </option>
                ))}
              </select>
            </div>
            <div>
              <div className="mb-2 text-sm font-medium text-stone-800">Password</div>
              <input
                type="password"
                value={loginForm.password}
                onChange={(event) => onChangeLogin((current) => ({ ...current, password: event.target.value }))}
                placeholder="Enter password"
                className="w-full rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm outline-none focus:border-emerald-800"
              />
            </div>
            {loginError && <div className="rounded-2xl bg-rose-50 px-4 py-3 text-sm text-rose-700">{loginError}</div>}
            {registerSuccess && (
              <div className="rounded-2xl bg-emerald-50 px-4 py-3 text-sm text-emerald-800">{registerSuccess}</div>
            )}
            <button type="submit" className="rounded-2xl bg-emerald-900 px-4 py-3 text-sm font-medium text-white">
              Log in
            </button>
          </form>
        ) : (
          <form onSubmit={onSubmitRegister} className="grid gap-4">
            <div className="rounded-2xl bg-stone-200/70 px-4 py-3 text-sm text-stone-700">
              Create a name and password. A manager will need to approve the account before you can sign in.
            </div>
            <div>
              <div className="mb-2 text-sm font-medium text-stone-800">Name</div>
              <input
                type="text"
                value={registerForm.username}
                onChange={(event) =>
                  onChangeRegister((current) => ({ ...current, username: event.target.value }))
                }
                placeholder="Enter your name"
                className="w-full rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm outline-none focus:border-emerald-800"
              />
            </div>
            <div>
              <div className="mb-2 text-sm font-medium text-stone-800">Password</div>
              <input
                type="password"
                value={registerForm.password}
                onChange={(event) =>
                  onChangeRegister((current) => ({ ...current, password: event.target.value }))
                }
                placeholder="Create a password"
                className="w-full rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm outline-none focus:border-emerald-800"
              />
            </div>
            {registerError && (
              <div className="rounded-2xl bg-rose-50 px-4 py-3 text-sm text-rose-700">{registerError}</div>
            )}
            <button type="submit" className="rounded-2xl bg-emerald-900 px-4 py-3 text-sm font-medium text-white">
              Send approval request
            </button>
          </form>
        )}
        <div className="mt-4 rounded-2xl bg-stone-200/70 px-4 py-3 text-sm text-stone-700">
          Need to make an account? Email <span className="font-semibold">sinthavong@data-mail.com</span>
        </div>
      </div>
    </div>
  );
}

function Detail({ label, value }) {
  return (
    <div className="rounded-2xl bg-stone-100 p-3">
      <div className="text-[11px] font-semibold uppercase tracking-[0.16em] text-stone-600">{label}</div>
      <div className="mt-1 text-sm font-medium text-stone-800">{value}</div>
    </div>
  );
}

function Field({ label, value, onChange, placeholder }) {
  return (
    <div>
      <div className="mb-2 text-sm font-medium text-stone-800">{label}</div>
      <input
        type="text"
        value={value}
        onChange={(event) => onChange(event.target.value)}
        placeholder={placeholder}
        className="w-full rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm outline-none focus:border-emerald-800"
      />
    </div>
  );
}

function StatRow({ label, value }) {
  return (
    <div className="rounded-2xl bg-stone-100 px-4 py-3">
      <div className="text-[11px] font-semibold uppercase tracking-[0.16em] text-stone-600">{label}</div>
      <div className="mt-1 text-lg font-semibold text-stone-900">{value}</div>
    </div>
  );
}

function AttachmentList({ attachments, onRemove = null }) {
  if (!attachments.length) {
    return (
      <div className="rounded-2xl border border-dashed border-stone-300 bg-white/60 p-4 text-sm text-stone-600">
        No documents attached yet.
      </div>
    );
  }

  return (
    <div className="space-y-2">
      {attachments.map((attachment) => (
        <div key={attachment.id} className="flex flex-col gap-3 rounded-2xl border border-stone-300 bg-white p-3 md:flex-row md:items-center md:justify-between">
          <div className="min-w-0">
            <a
              href={attachment.dataUrl}
              download={attachment.name}
              className="block truncate text-sm font-semibold text-stone-900 underline-offset-2 hover:underline"
            >
              {attachment.name}
            </a>
            <div className="mt-1 text-xs text-stone-600">
              {formatFileSize(attachment.size)} • uploaded {formatDateTime(attachment.uploadedAt)}
              {attachment.uploadedBy ? ` by ${attachment.uploadedBy}` : ""}
            </div>
          </div>
          {onRemove && (
            <button
              onClick={() => onRemove(attachment.id)}
              className="rounded-2xl border border-rose-200 px-3 py-2 text-sm text-rose-700"
            >
              Remove
            </button>
          )}
        </div>
      ))}
    </div>
  );
}

function RequestCard({
  request,
  onDone,
  onDelete,
  onAddAttachments,
  onRemoveAttachment,
  readOnly = false,
}) {
  return (
    <div className="rounded-2xl border border-stone-300 bg-white p-4">
      <div className="flex flex-col gap-3">
        <div className="flex flex-col gap-3 md:flex-row md:items-start md:justify-between">
          <div className="min-w-0 flex-1">
            <div className="flex flex-wrap items-center gap-2">
              <div className="text-sm font-semibold">
                {request.customer} - {request.jobNumber}
              </div>
              <span className={`rounded-full px-2 py-1 text-[11px] font-medium ${statusTone(request.status)}`}>
                {request.status}
              </span>
            </div>
            <div className="mt-1 text-xs text-stone-600">
              Requested by {request.requestorName} on {formatDateTime(request.createdAt)}
              {request.createdByAccount ? ` using ${request.createdByAccount}` : ""}
            </div>
            {request.completedAt && (
              <div className="mt-1 text-xs text-stone-600">
                Completed on {formatDateTime(request.completedAt)}
                {request.completedByAccount ? ` by ${request.completedByAccount}` : ""}
              </div>
            )}
            <div className="mt-3 whitespace-pre-wrap text-sm text-stone-800">{request.description}</div>
          </div>
          {!readOnly && (
            <div className="flex gap-2">
              <button onClick={onDone} className="rounded-2xl bg-emerald-900 px-3 py-2 text-sm text-white">
                Mark done
              </button>
              <button onClick={onDelete} className="rounded-2xl border border-rose-200 px-3 py-2 text-sm text-rose-700">
                Delete
              </button>
            </div>
          )}
        </div>

        <div>
          <div className="mb-2 text-xs font-semibold uppercase tracking-[0.16em] text-stone-600">Documents</div>
          <AttachmentList
            attachments={request.attachments || []}
            onRemove={readOnly ? null : onRemoveAttachment}
          />
        </div>

        {!readOnly && onAddAttachments && (
          <label className="block rounded-2xl border border-dashed border-stone-300 bg-stone-50 p-4 text-sm text-stone-700 hover:border-emerald-800">
            <div className="font-medium text-stone-900">Add more files to this request</div>
            <input
              type="file"
              accept={ATTACHMENT_ACCEPT}
              multiple
              onChange={(event) => {
                const files = event.target.files;
                if (files?.length) onAddAttachments(files);
                event.target.value = "";
              }}
              className="mt-3 block w-full text-xs"
            />
          </label>
        )}
      </div>
    </div>
  );
}

function JobCard({
  job,
  state,
  onClick,
  onDoubleClick,
  onQuickAssign,
  onFinish,
  onUpdatePress,
  pressOptions = [],
  weekColumns,
  finishedAt,
  finishedBy,
  canMove = true,
}) {
  const isDraggable = state !== "finished" && canMove;

  return (
    <div
      draggable={isDraggable}
      onDragStart={(event) => {
        if (!isDraggable) return;
        event.dataTransfer.effectAllowed = "move";
        event.dataTransfer.setData(
          "application/json",
          makeDragPayload({ type: "queue", jobId: job.id })
        );
        event.dataTransfer.setData("text/plain", job.id);
      }}
      onClick={() => onClick?.()}
      onDoubleClick={() => onDoubleClick?.()}
      className={`rounded-2xl border border-stone-300 bg-white p-3 shadow-sm shadow-stone-300/20 ${isDraggable ? "cursor-grab" : ""}`}
    >
      <div className="mb-2 flex items-start justify-between gap-2">
        <div
          onDoubleClick={onDoubleClick}
          role="button"
          tabIndex={0}
          onKeyDown={(event) => {
            if (event.key === "Enter" || event.key === " ") onClick?.();
          }}
          className="min-w-0 flex-1 cursor-pointer text-left"
        >
          <div className="text-sm font-semibold leading-tight text-stone-900">
            {job.customerName} {job.number}
          </div>
          <div className="mt-1 line-clamp-2 text-xs leading-5 text-stone-700">{job.generalDescr}</div>
        </div>
        <div className="flex flex-col items-end gap-2">
          <span className={`rounded-full border px-2 py-1 text-[11px] font-medium ${priorityTone(job.priority)}`}>
            {job.priority || "-"}
          </span>
          {isDraggable && (
            <span className="rounded-full bg-stone-200 px-2 py-1 text-[10px] font-medium text-stone-700">
              drag
            </span>
          )}
        </div>
      </div>

      <div className="grid grid-cols-2 gap-2 text-xs text-stone-700">
        <InfoPill label="Press" value={job.press || "-"} />
        <InfoPill label="EST" value={`${job.estPressTime.toFixed(2)}h`} />
        <InfoPill label="Qty" value={job.ticQuantity.toLocaleString()} />
        <InfoPill label="Ship" value={formatDate(job.shipByDate)} />
      </div>

      {onUpdatePress && pressOptions.length > 0 && (
        <div className="mt-2">
          <div className="mb-1 text-[10px] font-semibold uppercase tracking-[0.16em] text-stone-600">Recommended press</div>
          <select
            value={PRESS_ORDER.includes(job.press) ? job.press : ""}
            onClick={(event) => event.stopPropagation()}
            onChange={(event) => {
              event.stopPropagation();
              onUpdatePress(event.target.value);
            }}
            className="w-full rounded-xl border border-stone-300 bg-stone-50 px-3 py-2 text-xs outline-none focus:border-emerald-800"
          >
            {pressOptions.map((press) => (
              <option key={press} value={press}>
                Press {press}
              </option>
            ))}
          </select>
        </div>
      )}

      <div className="mt-2 flex flex-wrap items-center gap-2">
        <span className={`rounded-full px-2 py-1 text-[11px] font-medium ${statusTone(state)}`}>{state}</span>
        {job.ticketStatus && (
          <span className="rounded-full bg-stone-200 px-2 py-1 text-[11px] font-medium text-stone-700">
            import: {job.ticketStatus}
          </span>
        )}
      </div>

      {finishedAt && <div className="mt-2 text-xs text-stone-600">Finished {formatDateTime(finishedAt)}</div>}
      {finishedBy && <div className="mt-1 text-xs text-stone-600">Finished by {finishedBy}</div>}

      {state !== "finished" && ((weekColumns && canMove) || onFinish) && (
        <div className="mt-3 flex flex-wrap gap-2">
          {weekColumns && canMove &&
            weekColumns.map((day) => (
              <button
                key={day.key}
                onClick={(event) => {
                  event.stopPropagation();
                  onQuickAssign?.(day.key);
                }}
                className="rounded-xl border border-stone-300 bg-stone-50 px-2 py-1 text-[11px] text-stone-800 hover:bg-stone-100"
              >
                {day.label.slice(0, 3)}
              </button>
            ))}
          {onFinish && (
            <button
              onClick={(event) => {
                event.stopPropagation();
                onFinish();
              }}
              className="rounded-xl bg-emerald-900 px-2 py-1 text-[11px] text-white"
            >
              Finish
            </button>
          )}
        </div>
      )}
    </div>
  );
}

function PaperPullCard({ request, onDone, onDelete, readOnly = false }) {
  return (
    <div className="rounded-2xl border border-stone-300 bg-white p-4">
      <div className="flex flex-col gap-3 md:flex-row md:items-start md:justify-between">
        <div className="space-y-2">
          <div className="flex flex-wrap items-center gap-2">
            <span className="rounded-full bg-stone-200 px-2 py-1 text-[11px] font-medium text-stone-800">{request.target}</span>
            <span className={`rounded-full px-2 py-1 text-[11px] font-medium ${request.status === "done" ? statusTone("done") : statusTone("open")}`}>
              {request.status}
            </span>
          </div>
          <div className="whitespace-pre-wrap text-sm text-stone-800">{request.details}</div>
          <div className="text-xs text-stone-600">
            Requested by {request.createdBy || "-"} on {formatDateTime(request.createdAt)}
          </div>
          {request.completedAt && (
            <div className="text-xs text-stone-600">
              Completed by {request.completedBy || "-"} on {formatDateTime(request.completedAt)}
            </div>
          )}
        </div>
        {!readOnly && (
          <div className="flex gap-2">
            <button onClick={onDone} className="rounded-2xl bg-emerald-900 px-3 py-2 text-sm font-medium text-white">
              Done
            </button>
            <button onClick={onDelete} className="rounded-2xl border border-rose-200 px-3 py-2 text-sm text-rose-700">
              Delete
            </button>
          </div>
        )}
      </div>
    </div>
  );
}

function SuppliesRequestCard({
  request,
  onDone,
  onDelete,
  onAddAttachments,
  onRemoveAttachment,
  readOnly = false,
}) {
  return (
    <div className="rounded-2xl border border-stone-300 bg-white p-4">
      <div className="flex flex-col gap-3">
        <div className="flex flex-col gap-3 md:flex-row md:items-start md:justify-between">
          <div className="min-w-0 flex-1">
            <div className="flex flex-wrap items-center gap-2">
              <div className="text-sm font-semibold">Supplies request</div>
              <span className={`rounded-full px-2 py-1 text-[11px] font-medium ${statusTone(request.status)}`}>
                {request.status}
              </span>
            </div>
            <div className="mt-1 text-xs text-stone-600">
              Requested by {request.createdBy || "-"} on {formatDateTime(request.createdAt)}
            </div>
            {request.completedAt && (
              <div className="mt-1 text-xs text-stone-600">
                Completed by {request.completedBy || "-"} on {formatDateTime(request.completedAt)}
              </div>
            )}
            <div className="mt-3 whitespace-pre-wrap text-sm text-stone-800">{request.details}</div>
          </div>
          {!readOnly && (
            <div className="flex gap-2">
              <button onClick={onDone} className="rounded-2xl bg-emerald-900 px-3 py-2 text-sm font-medium text-white">
                Done
              </button>
              <button onClick={onDelete} className="rounded-2xl border border-rose-200 px-3 py-2 text-sm text-rose-700">
                Delete
              </button>
            </div>
          )}
        </div>

        <div>
          <div className="mb-2 text-xs font-semibold uppercase tracking-[0.16em] text-stone-600">Documents</div>
          <AttachmentList attachments={request.attachments || []} onRemove={readOnly ? null : onRemoveAttachment} />
        </div>

        {!readOnly && onAddAttachments && (
          <label className="block rounded-2xl border border-dashed border-stone-300 bg-stone-50 p-4 text-sm text-stone-700 hover:border-emerald-800">
            <div className="font-medium text-stone-900">Add more files to this supplies request</div>
            <input
              type="file"
              accept={ATTACHMENT_ACCEPT}
              multiple
              onChange={(event) => {
                const files = event.target.files;
                if (files?.length) onAddAttachments(files);
                event.target.value = "";
              }}
              className="mt-3 block w-full text-xs"
            />
          </label>
        )}
      </div>
    </div>
  );
}

function CompactScheduleCard({
  job,
  assignment,
  finishMeta,
  compact = false,
  onSelect,
  onUnschedule,
  onFinish,
  onDuplicate,
  draggable = false,
}) {
  const isManual = assignment.kind === "manual";
  const state = isManual ? "note" : assignment.status === "finished" ? "done" : assignment.status;
  const title = isManual ? assignment.manualTitle || "Manual block" : `${job.customerName} ${job.number}`;
  const subtitle = isManual ? "Manual schedule block" : job.generalDescr;
  const isCardDraggable = draggable && !!job;

  return (
    <div
      draggable={isCardDraggable}
      onDragStart={(event) => {
        if (!isCardDraggable) return;
        event.dataTransfer.effectAllowed = "move";
        event.dataTransfer.setData(
          "application/json",
          makeDragPayload({ type: "scheduled", assignmentId: assignment.id, jobId: job.id })
        );
      }}
      className={`rounded-2xl border border-stone-300 bg-stone-100 p-2 ${isCardDraggable ? "cursor-grab" : ""}`}
    >
      <div className="flex items-start justify-between gap-2">
        <div
          onClick={onSelect}
          role="button"
          tabIndex={0}
          onKeyDown={(event) => {
            if (event.key === "Enter" || event.key === " ") onSelect?.();
          }}
          className={`min-w-0 flex-1 ${onSelect ? "cursor-pointer" : ""}`}
        >
          <div className="truncate text-xs font-semibold text-stone-900">
            {title}
          </div>
          {!compact && <div className="mt-1 line-clamp-2 text-[11px] text-stone-700">{subtitle}</div>}
        </div>
        <span className={`rounded-full px-2 py-1 text-[10px] font-medium ${statusTone(state)}`}>{state}</span>
      </div>
      {isCardDraggable && !compact && <div className="mt-1 text-[10px] text-stone-600">Drag to move</div>}
      {!compact && !isManual && (
        <div className="mt-2 grid grid-cols-2 gap-2 text-[11px] text-stone-700">
          <InfoPill label="Est" value={`${job.estPressTime.toFixed(2)}h`} />
          <InfoPill label="Qty" value={job.ticQuantity.toLocaleString()} />
        </div>
      )}
      {finishMeta?.finishedAt && !compact && !isManual && (
        <div className="mt-2 text-[11px] text-stone-600">
          {formatDateTime(finishMeta.finishedAt)} {finishMeta.finishedBy ? `- ${finishMeta.finishedBy}` : ""}
        </div>
      )}
      {(onUnschedule || onFinish || onDuplicate) && (
        <div className={`${compact ? "mt-1" : "mt-2"} flex flex-wrap gap-2`}>
          {onUnschedule && (
            <button onClick={onUnschedule} className="rounded-xl border border-stone-300 bg-white px-2 py-1 text-[11px] text-stone-800">
              Remove
            </button>
          )}
          {onDuplicate && (
            <button onClick={onDuplicate} className="rounded-xl border border-stone-300 bg-white px-2 py-1 text-[11px] text-stone-800">
              Duplicate
            </button>
          )}
          {onFinish && (
            <button onClick={onFinish} className="rounded-xl bg-emerald-900 px-2 py-1 text-[11px] text-white">
              Finish
            </button>
          )}
        </div>
      )}
    </div>
  );
}

function InfoPill({ label, value }) {
  return (
    <div className="rounded-xl border border-stone-200 bg-white px-2 py-2">
      <div className="text-[10px] font-semibold uppercase tracking-[0.16em] text-stone-600">{label}</div>
      <div className="mt-1 font-medium text-stone-800">{value}</div>
    </div>
  );
}
