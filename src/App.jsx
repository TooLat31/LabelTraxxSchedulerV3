import React, { useEffect, useMemo, useRef, useState } from "react";
import { isSupabaseConfigured, supabase } from "./lib/supabase";

const PRESS_ORDER = ["5.1", "6.1", "1.1", "2.1", "8", "9", "Rewind"];
const STORAGE_KEY = "labeltraxx-scheduler-v4";
const SESSION_STORAGE_KEY = "labeltraxx-scheduler-session-v1";
const SHARED_STATE_ROW_ID = "labeltraxx-shared-state";
const BASE_TABS = ["Scheduler", "New Request", "Open Requests", "Request History", "Pull Paper Request", "Daily Shipment"];
const ATTACHMENT_ACCEPT =
  ".pdf,.doc,.docx,.xls,.xlsx,.csv,.txt,.rtf,.png,.jpg,.jpeg,.zip,.msg,.eml";
const PULL_PAPER_TARGETS = ["Press 5.1", "Press 6.1", "Press 2.1", "Press 1.1", "Digital"];
const ROLE_OPTIONS = ["Management", "Warehouse/Shipper", "Operator"];

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

const EMPTY_LOGIN_FORM = {
  username: "",
  password: "",
};

const EMPTY_USER_FORM = {
  username: "",
  password: "",
  role: ROLE_OPTIONS[0],
};

const EMPTY_SHIPMENT_FORM = {
  label: "",
  method: "Skid",
  packageCount: "",
  packageType: "Skids",
  totalCost: "",
  notes: "",
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

function normalizeRole(role, isAdmin) {
  const normalized = safeText(role);
  if (ROLE_OPTIONS.includes(normalized)) return normalized;
  return isAdmin ? "Management" : "Warehouse/Shipper";
}

function getUserRole(user) {
  return normalizeRole(user?.role, user?.isAdmin);
}

function hasManagementAccess(user) {
  return getUserRole(user) === "Management";
}

function canMoveJobs(user) {
  return getUserRole(user) !== "Operator";
}

function canAccessTab(user, tab) {
  const role = getUserRole(user);
  if (role === "Management") return true;
  if (role === "Warehouse/Shipper") {
    return !["New Request", "Open Requests", "User Admin"].includes(tab);
  }
  if (role === "Operator") {
    return tab === "Scheduler";
  }
  return tab === "Scheduler";
}

function buildDefaultAdmin() {
  return {
    id: "user-admin",
    username: "Admin",
    password: "1234",
    role: "Management",
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
          role: normalizeRole(user.role, user.isAdmin),
          isAdmin: hasManagementAccess(user),
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

function normalizeAssignments(assignments) {
  return Array.isArray(assignments)
    ? assignments.filter((assignment) => assignment.kind !== "rewind")
    : [];
}

function normalizeShipmentGroups(groups) {
  return Array.isArray(groups) ? groups : [];
}

function defaultSharedSnapshot() {
  return {
    jobs: [],
    assignments: [],
    requests: [],
    pullPaperRequests: [],
    shipmentGroups: [],
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
    shipmentGroups: normalizeShipmentGroups(source.shipmentGroups),
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
    shipmentGroups: state.shipmentGroups,
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
  const [shipmentGroups, setShipmentGroups] = useState([]);
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
  const [requestDraftAttachments, setRequestDraftAttachments] = useState([]);
  const [selectedShipDate, setSelectedShipDate] = useState(todayKey());
  const [shipDateDraft, setShipDateDraft] = useState(todayKey());
  const [selectedShipmentJobs, setSelectedShipmentJobs] = useState([]);
  const [selectedShipQueueJobs, setSelectedShipQueueJobs] = useState([]);
  const [shipmentForm, setShipmentForm] = useState({ ...EMPTY_SHIPMENT_FORM, shipDate: todayKey() });
  const [loginForm, setLoginForm] = useState(EMPTY_LOGIN_FORM);
  const [loginError, setLoginError] = useState("");
  const [userForm, setUserForm] = useState(EMPTY_USER_FORM);
  const [userPasswordDrafts, setUserPasswordDrafts] = useState({});
  const [requestHistoryFilterDate, setRequestHistoryFilterDate] = useState("");
  const [shipmentHistoryFilterDate, setShipmentHistoryFilterDate] = useState("");
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
      setShipmentGroups(normalized.shipmentGroups);
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
        if (sessionUsername) {
          const match = normalized.users.find(
            (user) => comparableUsername(user.username) === comparableUsername(sessionUsername)
          );
          if (match) setCurrentUsername(match.username);
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
      shipmentGroups,
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
  }, [assignments, currentUsername, isReady, jobs, pullPaperRequests, requests, shipmentGroups, users, weekStart]);

  useEffect(() => {
    localStorage.setItem(
      SESSION_STORAGE_KEY,
      JSON.stringify({
        currentUsername,
      })
    );
  }, [currentUsername]);

  const currentUser = useMemo(
    () =>
      users.find((user) => comparableUsername(user.username) === comparableUsername(currentUsername)) || null,
    [currentUsername, users]
  );

  const currentUserRole = useMemo(() => getUserRole(currentUser), [currentUser]);
  const userCanManageUsers = useMemo(() => hasManagementAccess(currentUser), [currentUser]);
  const userCanMoveJobs = useMemo(() => canMoveJobs(currentUser), [currentUser]);

  useEffect(() => {
    if (currentUser) return;
    if (!currentUsername) return;
    setCurrentUsername("");
  }, [currentUser, currentUsername]);

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
          setShipmentGroups(normalized.shipmentGroups);
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
      .filter((job) => !activePressJobIds.has(job.id))
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
        const haystack = `${job.number} ${job.customerName} ${job.generalDescr}`.toLowerCase();
        return !unscheduledSearch.trim() || haystack.includes(unscheduledSearch.toLowerCase());
      })
      .sort((left, right) => {
        const leftDate = left.shipByDate ? left.shipByDate.getTime() : Number.MAX_SAFE_INTEGER;
        const rightDate = right.shipByDate ? right.shipByDate.getTime() : Number.MAX_SAFE_INTEGER;
        if (leftDate !== rightDate) return leftDate - rightDate;
        return right.estPressTime - left.estPressTime;
      });
  }, [activePressJobIds, filteredJobs, queuePressFilter, queueStatusFilter, unscheduledSearch, userFinishedJobIds]);

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
      if (!filteredJobIds.has(assignment.jobId)) return;
      if (!map[assignment.dayKey]) return;
      const job = jobMap.get(assignment.jobId);
      if (!job) return;
      map[assignment.dayKey][assignment.press].push({ assignment, job });
    });

    Object.values(map).forEach((pressMap) => {
      Object.values(pressMap).forEach((laneJobs) => {
        laneJobs.sort((left, right) => right.job.estPressTime - left.job.estPressTime);
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

  const queuePressOptions = useMemo(
    () => ["All", ...Array.from(new Set(jobs.map((job) => safeText(job.press)).filter(Boolean))).sort()],
    [jobs]
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
    setAssignments((current) => current.filter((assignment) => parsed.some((job) => job.id === assignment.jobId)));
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

  function moveAssignment(assignmentId, dayKey, press) {
    if (!userCanMoveJobs) return;
    setAssignments((current) => {
      const assignmentToMove = current.find((assignment) => assignment.id === assignmentId);
      if (!assignmentToMove) return current;
      if (assignmentToMove.kind !== "press") return current;
      if (assignmentToMove.status === "finished") return current;

      const duplicate = current.some(
        (assignment) =>
          assignment.id !== assignmentId &&
          assignment.jobId === assignmentToMove.jobId &&
          assignment.kind === "press" &&
          assignment.status !== "finished" &&
          assignment.dayKey === dayKey &&
          assignment.press === press
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
      if (assignmentToCopy.kind !== "press") return current;
      if (assignmentToCopy.status === "finished") return current;

      const currentIndex = weekColumns.findIndex((day) => day.key === assignmentToCopy.dayKey);
      if (currentIndex < 0 || currentIndex >= weekColumns.length - 1) return current;

      const nextDayKey = weekColumns[currentIndex + 1].key;
      const exists = current.some(
        (assignment) =>
          assignment.id !== assignmentId &&
          assignment.jobId === assignmentToCopy.jobId &&
          assignment.kind === "press" &&
          assignment.status !== "finished" &&
          assignment.dayKey === nextDayKey &&
          assignment.press === assignmentToCopy.press
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
    if (!currentUser) return;
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
    if (!currentUser) return;
    const nextAttachments = await buildAttachments(event.target.files, currentUser.username);
    if (nextAttachments.length) {
      setRequestDraftAttachments((current) => [...current, ...nextAttachments]);
    }
    event.target.value = "";
  }

  function removeDraftAttachment(attachmentId) {
    setRequestDraftAttachments((current) => current.filter((attachment) => attachment.id !== attachmentId));
  }

  function submitRequest(event) {
    event.preventDefault();
    if (!currentUser) return;
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
    if (!currentUser) return;
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

  function markPullPaperRequestDone(requestId) {
    if (!currentUser) return;
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
    setPullPaperRequests((current) => current.filter((request) => request.id !== requestId));
  }

  function markRequestDone(requestId) {
    if (!currentUser) return;
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
    setRequests((current) => current.filter((request) => request.id !== requestId));
  }

  async function addRequestAttachments(requestId, fileList) {
    if (!currentUser) return;
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

  function createShipmentGroup(event) {
    event.preventDefault();
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
        notes: shipmentForm.notes,
        shipDate: shipmentForm.shipDate,
        createdAt: new Date().toISOString(),
        createdBy: currentUser?.username || "",
        jobItems,
      },
      ...current,
    ]);

    setSelectedShipmentJobs([]);
    setShipmentForm({ ...EMPTY_SHIPMENT_FORM, shipDate: selectedShipDate });
  }

  function deleteShipmentGroup(groupId) {
    setShipmentGroups((current) => current.filter((group) => group.id !== groupId));
  }

  function loadDemo() {
    const sample = `Press\tNumber\tCustomerName\tGeneralDescr\tCustPONum\tPriority\tShip_by_Date\tEntryDate\tDue_on_Site_Date\tStockNum2\tStockNum1\tStatus\tMainTool\tToolNo2\tTicQuantity\tEstFootage\tEstPressTime\tNotes
5.1\t10159\tData Graphics\t1.625" Cap One Circle 70072\t223000DG\tHigh\t04/28/26\t04/18/26\t04/29/26\t266\t\tOpen\t946\t\t3,612,279\t113,847\t9.76\tExample long notes
8\t11180\tPremio Foods\t3.125"x4.1875" Premio\t4500081640\tDigital\t04/29/26\t04/16/26\t04/30/26\t266\t590\tDone\tD-904\t\t96,000\t13,328\t1.73\tImported done should not ship until you mark it finished.
6.1\t11194\tPremio Foods\t3.25"x5" Premio Contract Release\t4500081729\tRelease\t04/28/26\t04/21/26\t04/29/26\t266\t590\tOpen\t668\t\t48,000\t13,020\t4.64\tContract PO 4600004905
9\t11022\tData Graphics\t1.625" Cap One Circle 69797\t\tHigh\t04/29/26\t03/24/26\t05/04/26\t266\t\tOpen\t946\t\t2,686,950\t86,266\t7.91\tArt Due 4/23`;
    importText(sample);
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
    setLoginForm(EMPTY_LOGIN_FORM);
    setLoginError("");
    setActiveTab("Scheduler");
  }

  function handleLogout() {
    setCurrentUsername("");
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
        role: normalizeRole(userForm.role),
        isAdmin: normalizeRole(userForm.role) === "Management",
        createdAt: new Date().toISOString(),
        createdBy: currentUser.username,
      },
    ]);
    setUserForm(EMPTY_USER_FORM);
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
        loginForm={loginForm}
        loginError={loginError}
        users={users}
        onChange={setLoginForm}
        onSubmit={handleLogin}
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
    "Open Requests": openRequests.length,
    "Pull Paper Request": openPullPaperRequests.length,
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
                  {currentUserRole}: {currentUser.username}
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
                </div>
              </div>

              <div className="rounded-3xl border border-stone-300 bg-stone-50 p-4 shadow-sm shadow-stone-300/30">
                <div className="mb-3 flex items-center justify-between">
                  <div>
                    <div className="text-sm font-semibold">Press queue</div>
                    <div className="text-xs text-stone-600">Search by ticket, then filter the queue by imported status or press number.</div>
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

                <div className="mb-3 grid grid-cols-2 gap-2">
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
                </div>

                <div className="max-h-[480px] space-y-3 overflow-y-auto pr-1">
                  {unscheduledJobs.map((job) => (
                    <JobCard
                      key={job.id}
                      job={job}
                      state={deriveVisibleJobState(job.id, activePressJobIds, userFinishedJobIds)}
                      onClick={() => selectJob(job.id)}
                      onDoubleClick={() => selectJob(job.id, true)}
                      onFinish={() => finishJob(job.id)}
                      weekColumns={weekColumns}
                      canMove={userCanMoveJobs}
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
                      {PRESS_ORDER.map((press) => {
                        const laneJobs = board[day.key]?.[press] || [];
                        const totalHours = laneJobs.reduce((sum, item) => sum + item.job.estPressTime, 0);
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
                                  finishMeta={finishedMetaByJobId.get(job.id)}
                                  onSelect={() => selectJob(job.id)}
                                  onUnschedule={userCanMoveJobs ? () => removeAssignment(assignment.id) : undefined}
                                  onFinish={() => finishJob(job.id)}
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

        {activeTab === "Daily Shipment" && (
          <div className="space-y-4">
            <div className="rounded-3xl border border-stone-300 bg-stone-50 p-5 shadow-sm shadow-stone-300/30">
              <div className="flex flex-col gap-3 md:flex-row md:items-end md:justify-between">
                <div>
                  <div className="text-sm font-semibold">Daily shipment</div>
                  <div className="text-xs text-stone-600">
                    Finished jobs can be assigned to a ship date first, then grouped under that shipping day.
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
                    <div className="text-xs text-stone-600">Example: one skid with 3 jobs for $141.17, or separate FedEx transactions.</div>
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
                        <option>Skid</option>
                        <option>FedEx</option>
                        <option>UPS</option>
                        <option>LTL</option>
                        <option>Customer Pickup</option>
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
                    <div>
                      <div className="mb-2 text-sm font-medium text-stone-800">Notes</div>
                      <textarea
                        value={shipmentForm.notes}
                        onChange={(event) => setShipmentForm((current) => ({ ...current, notes: event.target.value }))}
                        placeholder="Optional shipment notes"
                        className="h-24 w-full rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm outline-none focus:border-emerald-800"
                      />
                    </div>
                    <button type="submit" className="rounded-2xl bg-emerald-900 px-4 py-3 text-sm font-medium text-white">
                      Create shipment group
                    </button>
                  </form>
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
                          {!!group.packageCount && (
                            <div className="mt-1 text-xs text-stone-600">
                              {group.packageCount} {safeText(group.packageType || "Skids").toLowerCase()}
                            </div>
                          )}
                          {group.notes && <div className="mt-2 text-sm text-stone-800">{group.notes}</div>}
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

        {activeTab === "User Admin" && userCanManageUsers && (
          <div className="grid gap-4 xl:grid-cols-[420px_minmax(0,1fr)]">
            <div className="rounded-3xl border border-stone-300 bg-stone-50 p-6 shadow-sm shadow-stone-300/30">
              <div className="mb-5">
                <div className="text-sm font-semibold">Add user</div>
                <div className="text-xs text-stone-600">Management can create logins, assign roles, and set passwords here.</div>
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
                  <div className="mb-2 text-sm font-medium text-stone-800">Role</div>
                  <select
                    value={userForm.role}
                    onChange={(event) => setUserForm((current) => ({ ...current, role: event.target.value }))}
                    className="w-full rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm outline-none focus:border-emerald-800"
                  >
                    {ROLE_OPTIONS.map((role) => (
                      <option key={role} value={role}>
                        {role}
                      </option>
                    ))}
                  </select>
                </div>
                <button type="submit" className="rounded-2xl bg-emerald-900 px-4 py-3 text-sm font-medium text-white">
                  Create user
                </button>
              </form>
            </div>

            <div className="rounded-3xl border border-stone-300 bg-stone-50 p-6 shadow-sm shadow-stone-300/30">
              <div className="mb-5">
                <div className="text-sm font-semibold">Manage users</div>
                <div className="text-xs text-stone-600">Reset passwords for any user and manage access from here.</div>
              </div>
              <div className="space-y-3">
                {users.map((user) => (
                  <div key={user.id} className="rounded-2xl border border-stone-300 bg-white p-4">
                    <div className="flex flex-col gap-3 md:flex-row md:items-center md:justify-between">
                      <div>
                        <div className="flex items-center gap-2">
                          <div className="text-sm font-semibold">{user.username}</div>
                          <span className={`rounded-full px-2 py-1 text-[11px] font-medium ${hasManagementAccess(user) ? "bg-emerald-900 text-white" : "bg-stone-200 text-stone-800"}`}>
                            {getUserRole(user)}
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
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

function LoginScreen({ loginForm, loginError, users, onChange, onSubmit }) {
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
        <form onSubmit={onSubmit} className="grid gap-4">
          <div>
            <div className="mb-2 text-sm font-medium text-stone-800">Username</div>
            <select
              value={loginForm.username}
              onChange={(event) => onChange((current) => ({ ...current, username: event.target.value }))}
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
              onChange={(event) => onChange((current) => ({ ...current, password: event.target.value }))}
              placeholder="Enter password"
              className="w-full rounded-2xl border border-stone-300 bg-white px-4 py-3 text-sm outline-none focus:border-emerald-800"
            />
          </div>
          {loginError && <div className="rounded-2xl bg-rose-50 px-4 py-3 text-sm text-rose-700">{loginError}</div>}
          <button type="submit" className="rounded-2xl bg-emerald-900 px-4 py-3 text-sm font-medium text-white">
            Log in
          </button>
        </form>
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
  weekColumns,
  finishedAt,
  finishedBy,
  canMove = true,
}) {
  const isDraggable = state === "open" && canMove;

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

      {state === "open" && ((weekColumns && canMove) || onFinish) && (
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

function CompactScheduleCard({
  job,
  assignment,
  finishMeta,
  onSelect,
  onUnschedule,
  onFinish,
  onDuplicate,
  draggable = false,
}) {
  const state = assignment.status === "finished" ? "done" : assignment.status;

  return (
    <div
      draggable={draggable}
      onDragStart={(event) => {
        if (!draggable) return;
        event.dataTransfer.effectAllowed = "move";
        event.dataTransfer.setData(
          "application/json",
          makeDragPayload({ type: "scheduled", assignmentId: assignment.id, jobId: job.id })
        );
      }}
      className={`rounded-2xl border border-stone-300 bg-stone-100 p-2 ${draggable ? "cursor-grab" : ""}`}
    >
      <div className="flex items-start justify-between gap-2">
        <div
          onClick={onSelect}
          role="button"
          tabIndex={0}
          onKeyDown={(event) => {
            if (event.key === "Enter" || event.key === " ") onSelect?.();
          }}
          className="min-w-0 flex-1 cursor-pointer"
        >
          <div className="truncate text-xs font-semibold text-stone-900">
            {job.customerName} {job.number}
          </div>
          <div className="mt-1 line-clamp-2 text-[11px] text-stone-700">{job.generalDescr}</div>
        </div>
        <span className={`rounded-full px-2 py-1 text-[10px] font-medium ${statusTone(state)}`}>{state}</span>
      </div>
      {draggable && <div className="mt-1 text-[10px] text-stone-600">Drag to move</div>}
      <div className="mt-2 grid grid-cols-2 gap-2 text-[11px] text-stone-700">
        <InfoPill label="Est" value={`${job.estPressTime.toFixed(2)}h`} />
        <InfoPill label="Qty" value={job.ticQuantity.toLocaleString()} />
      </div>
      {finishMeta?.finishedAt && (
        <div className="mt-2 text-[11px] text-stone-600">
          {formatDateTime(finishMeta.finishedAt)} {finishMeta.finishedBy ? `- ${finishMeta.finishedBy}` : ""}
        </div>
      )}
      {(onUnschedule || onFinish || onDuplicate) && (
        <div className="mt-2 flex gap-2">
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
