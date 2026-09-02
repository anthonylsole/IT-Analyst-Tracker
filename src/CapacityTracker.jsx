import { useState, useEffect, useRef } from "react";
import { Plus, X, Check, AlertTriangle, Gauge, Bell, Download, Upload, History } from "lucide-react";

const PALETTE = [
  "#1F6F63",
  "#E8A33D",
  "#C1443C",
  "#4F5FA6",
  "#6B8F3E",
  "#8E5B9E",
  "#5C7080",
  "#B08B4F",
];

const DATA_KEY = "capacity-board-data";
const USER_KEY = "capacity-current-user";
const BACKUP_PREFIX = "capacity-board-backup:";
const MAX_BACKUPS = 10;

function uid() {
  return Math.random().toString(36).slice(2) + Date.now().toString(36);
}

function clamp(n, lo, hi) {
  return Math.max(lo, Math.min(hi, n));
}

// Seeded from SharePoint: ITKnowledgeBase / Analyst Allocation page
const SEED_MEMBER_NAMES = ["Dusko", "Paul", "Sandy", "Diego", "Shelly", "Kim", "Bhaskar", "Carmen", "Brent"];
// "Last updated" dates from the SharePoint page header (e.g. DUSKO 7.1 = July 1, 2026)
const SEED_MEMBER_DATES = ["2026-07-01", "2026-07-29", "2026-06-25", "2026-06-25", "2026-08-04", "2026-07-02", "2026-05-13", "2026-06-25", "2026-05-13"];

const SEED_PROJECT_NAMES = [
  "Tifton WMS",
  "BN \u2013 Barrett Integration",
  "BN \u2013 Barrett PROD Support",
  "BN \u2013 General Support",
  "SPL \u2013 Anaplan",
  "SAP",
  "7FAM \u2013 POS - xStore",
  "Fontana \u2013 Passport Integration",
  "Fontana \u2013 WMS Support",
  "General \u2013 Ecom Projects",
  "General \u2013 Drop Ship / EDI Projects",
  "General \u2013 Other gen support",
];

// One row per project above, values = % allocated by that analyst
const SEED_ALLOCATION_ROWS = [
  { Dusko: 60, Paul: 70, Sandy: 10, Brent: 40 },
  { Diego: 30, Brent: 40 },
  { Paul: 5, Diego: 10, Shelly: 60 },
  { Paul: 25, Sandy: 10, Diego: 25 },
  { Dusko: 10, Brent: 10 },
  { Sandy: 25, Diego: 5, Kim: 75, Bhaskar: 60, Carmen: 80 },
  {},
  {},
  { Shelly: 10, Bhaskar: 40 },
  { Diego: 15 },
  { Sandy: 30, Kim: 5, Carmen: 10 },
  { Dusko: 30, Sandy: 25, Diego: 15, Shelly: 30, Kim: 20, Carmen: 10, Brent: 10 },
];

function buildSeedData() {
  const members = SEED_MEMBER_NAMES.map((name, i) => ({ id: uid(), name, lastUpdated: SEED_MEMBER_DATES[i] || "", lastRequested: "" }));
  const idByName = {};
  members.forEach((m) => (idByName[m.name] = m.id));

  const projects = SEED_PROJECT_NAMES.map((name, i) => ({
    id: uid(),
    name,
    color: PALETTE[i % PALETTE.length],
  }));

  const allocations = {};
  members.forEach((m) => (allocations[m.id] = {}));
  projects.forEach((p, i) => {
    const row = SEED_ALLOCATION_ROWS[i] || {};
    Object.entries(row).forEach(([name, pct]) => {
      const mid = idByName[name];
      if (mid) allocations[mid][p.id] = pct;
    });
  });

  return { projects, members, allocations };
}

export default function CapacityTracker() {
  const [mode, setMode] = useState("dashboard"); // "dashboard" | "selfserve"
  const [data, setData] = useState({ projects: [], members: [], allocations: {} });
  const [loaded, setLoaded] = useState(false);
  const [loadError, setLoadError] = useState(false);
  const [saveState, setSaveState] = useState("idle");
  const [currentUserId, setCurrentUserId] = useState("");
  const [selectedProjectId, setSelectedProjectId] = useState("");
  const [addingProject, setAddingProject] = useState(false);
  const [addingMember, setAddingMember] = useState(false);
  const [newProjectName, setNewProjectName] = useState("");
  const [newMemberName, setNewMemberName] = useState("");
  const [selfServeUserId, setSelfServeUserId] = useState("");
  const [selfServeSubmitted, setSelfServeSubmitted] = useState(false);
  const saveTimeout = useRef(null);
  const pendingSaveData = useRef(null);
  const saveChain = useRef(Promise.resolve());
  const fileInputRef = useRef(null);
  const [showBackups, setShowBackups] = useState(false);
  const [backupsList, setBackupsList] = useState([]);
  const [backupsLoading, setBackupsLoading] = useState(false);
  const [restoreNotice, setRestoreNotice] = useState("");

  async function loadData() {
    setLoadError(false);
    try {
      const res = await window.storage.get(DATA_KEY, true);
      if (res && res.value) {
        setData(JSON.parse(res.value));
      } else {
        // Key genuinely doesn't exist yet — this is a true first-time load.
        const seed = buildSeedData();
        setData(seed);
        try {
          await window.storage.set(DATA_KEY, JSON.stringify(seed), true);
        } catch (e2) {
          // seed will still show locally even if save fails
        }
      }
    } catch (e) {
      if (e && e.code === "NOT_FOUND") {
        // Normal first-time load (e.g. a brand-new database) — not an error.
        const seed = buildSeedData();
        setData(seed);
        try {
          await window.storage.set(DATA_KEY, JSON.stringify(seed), true);
        } catch (e2) {
          // seed will still show locally even if save fails
        }
      } else {
        // A genuine fetch/connectivity error — NOT proof the data doesn't
        // exist. Never overwrite shared storage here, or a brief blip could
        // silently wipe out everyone's real data.
        setLoadError(true);
        const seed = buildSeedData();
        setData(seed);
      }
    }
    try {
      const res2 = await window.storage.get(USER_KEY, false);
      if (res2 && res2.value) setCurrentUserId(res2.value);
    } catch (e) {
      // no local user set yet
    }
    setLoaded(true);
  }

  useEffect(() => {
    loadData();
  }, []);

  useEffect(() => {
    function handleVisibilityOrUnload() {
      if (document.hidden) {
        flushPendingSave();
      }
    }
    document.addEventListener("visibilitychange", handleVisibilityOrUnload);
    window.addEventListener("pagehide", handleVisibilityOrUnload);
    window.addEventListener("beforeunload", handleVisibilityOrUnload);
    return () => {
      document.removeEventListener("visibilitychange", handleVisibilityOrUnload);
      window.removeEventListener("pagehide", handleVisibilityOrUnload);
      window.removeEventListener("beforeunload", handleVisibilityOrUnload);
    };
  }, []);

  function persist(nextData) {
    setData(nextData);
    setSaveState("saving");
    pendingSaveData.current = nextData;
    if (saveTimeout.current) clearTimeout(saveTimeout.current);
    saveTimeout.current = setTimeout(() => {
      flushPendingSave();
    }, 400);
  }

  async function flushPendingSave() {
    if (saveTimeout.current) {
      clearTimeout(saveTimeout.current);
      saveTimeout.current = null;
    }
    const toSave = pendingSaveData.current;
    if (toSave === null) return;
    pendingSaveData.current = null;
    // Chain this write after any write already in flight completes. This
    // guarantees writes land in the order they were made — a faster network
    // response for a later edit can never overwrite an earlier one anymore.
    saveChain.current = saveChain.current.then(async () => {
      try {
        const ok = await window.storage.set(DATA_KEY, JSON.stringify(toSave), true);
        setSaveState(ok ? "saved" : "error");
        if (ok) {
          // Best-effort recovery point. Never blocks or fails the main save.
          writeBackupSnapshot(toSave).catch(() => {});
        }
      } catch (e) {
        setSaveState("error");
      }
    });
    await saveChain.current;
  }

  async function writeBackupSnapshot(snapshotData) {
    const key = `${BACKUP_PREFIX}${Date.now()}`;
    await window.storage.set(key, JSON.stringify(snapshotData), true);
    try {
      const list = await window.storage.list(BACKUP_PREFIX, true);
      const keys = (list && list.keys) || [];
      if (keys.length > MAX_BACKUPS) {
        const sorted = [...keys].sort(); // ISO-ish timestamp suffix sorts chronologically
        const toDelete = sorted.slice(0, sorted.length - MAX_BACKUPS);
        await Promise.all(toDelete.map((k) => window.storage.delete(k, true).catch(() => {})));
      }
    } catch (e) {
      // pruning is best-effort only
    }
  }

  async function createBackupNow() {
    setRestoreNotice("Creating backup…");
    try {
      await writeBackupSnapshot(data);
      setRestoreNotice("Backup created just now.");
      if (showBackups) loadBackupsList();
    } catch (e) {
      setRestoreNotice("Couldn't create a backup — please try again.");
    }
  }

  async function loadBackupsList() {
    setBackupsLoading(true);
    try {
      const list = await window.storage.list(BACKUP_PREFIX, true);
      const keys = (list && list.keys) || [];
      const withTimes = keys
        .map((k) => {
          const ts = Number(k.slice(BACKUP_PREFIX.length));
          return { key: k, ts };
        })
        .filter((entry) => !Number.isNaN(entry.ts))
        .sort((a, b) => b.ts - a.ts);
      setBackupsList(withTimes);
    } catch (e) {
      setBackupsList([]);
    }
    setBackupsLoading(false);
  }

  async function restoreBackup(key, label) {
    const confirmed = window.confirm(
      `Restore the board to the snapshot from ${label}? This will overwrite what's currently on screen (your current state is not automatically saved first).`
    );
    if (!confirmed) return;
    try {
      const res = await window.storage.get(key, true);
      if (res && res.value) {
        const parsed = JSON.parse(res.value);
        persist(parsed);
        setRestoreNotice(`Restored to snapshot from ${label}.`);
        setShowBackups(false);
      }
    } catch (e) {
      setRestoreNotice("Could not restore that snapshot — please try again.");
    }
  }

  function exportData() {
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `capacity-board-backup-${todayISO()}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  function triggerImport() {
    if (fileInputRef.current) fileInputRef.current.click();
  }

  function handleImportFile(e) {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const parsed = JSON.parse(reader.result);
        if (parsed && parsed.projects && parsed.members && parsed.allocations) {
          const confirmed = window.confirm(
            "Import this file? This will overwrite the current board with the data in this file."
          );
          if (confirmed) {
            persist(parsed);
            setRestoreNotice("Imported data from file.");
          }
        } else {
          window.alert("That file doesn't look like a valid Capacity board backup.");
        }
      } catch (err) {
        window.alert("Couldn't read that file as JSON.");
      }
    };
    reader.readAsText(file);
    e.target.value = "";
  }

  function selectUser(id) {
    setCurrentUserId(id);
    window.storage.set(USER_KEY, id, false).catch(() => {});
  }

  function addProject(name) {
    const trimmed = name.trim();
    if (!trimmed) return;
    const color = PALETTE[data.projects.length % PALETTE.length];
    const project = { id: uid(), name: trimmed, color };
    persist({ ...data, projects: [...data.projects, project] });
    setNewProjectName("");
    setAddingProject(false);
  }

  function removeProject(id) {
    const allocations = {};
    Object.keys(data.allocations).forEach((mid) => {
      const rest = { ...data.allocations[mid] };
      delete rest[id];
      allocations[mid] = rest;
    });
    persist({ ...data, projects: data.projects.filter((p) => p.id !== id), allocations });
  }

  function addMember(name) {
    const trimmed = name.trim();
    if (!trimmed) return;
    const member = { id: uid(), name: trimmed, lastUpdated: "", lastRequested: "" };
    persist({
      ...data,
      members: [...data.members, member],
      allocations: { ...data.allocations, [member.id]: {} },
    });
    setNewMemberName("");
    setAddingMember(false);
  }

  function restoreMissingMembers() {
    const existingNames = new Set(data.members.map((m) => m.name));
    const missingNames = SEED_MEMBER_NAMES.filter((name) => !existingNames.has(name));
    if (missingNames.length === 0) return;

    // Ensure all seeded projects exist too, so restored allocations have somewhere to land
    const existingProjectNames = new Map(data.projects.map((p) => [p.name, p.id]));
    const newProjects = [];
    SEED_PROJECT_NAMES.forEach((name, i) => {
      if (!existingProjectNames.has(name)) {
        const color = PALETTE[(data.projects.length + newProjects.length) % PALETTE.length];
        const project = { id: uid(), name, color };
        newProjects.push(project);
        existingProjectNames.set(name, project.id);
      }
    });

    const newMembers = missingNames.map((name) => {
      const idx = SEED_MEMBER_NAMES.indexOf(name);
      return { id: uid(), name, lastUpdated: SEED_MEMBER_DATES[idx] || "", lastRequested: "" };
    });

    const allocations = { ...data.allocations };
    newMembers.forEach((m) => {
      allocations[m.id] = {};
    });
    newMembers.forEach((m) => {
      const idx = SEED_MEMBER_NAMES.indexOf(m.name);
      SEED_ALLOCATION_ROWS.forEach((row, projIdx) => {
        const pct = row[m.name];
        if (pct) {
          const projName = SEED_PROJECT_NAMES[projIdx];
          const projId = existingProjectNames.get(projName);
          if (projId) allocations[m.id][projId] = pct;
        }
      });
    });

    persist({
      ...data,
      projects: [...data.projects, ...newProjects],
      members: [...data.members, ...newMembers],
      allocations,
    });
  }

  function setLastUpdated(memberId, dateStr) {
    const members = data.members.map((m) => (m.id === memberId ? { ...m, lastUpdated: dateStr } : m));
    persist({ ...data, members });
  }

  function requestUpdate(memberId) {
    const members = data.members.map((m) => (m.id === memberId ? { ...m, lastRequested: todayISO() } : m));
    persist({ ...data, members });
  }

  function todayISO() {
    return new Date().toISOString().slice(0, 10);
  }

  function setAllocation(memberId, projectId, value) {
    const v = clamp(Math.round(value), 0, 100);
    const memberAlloc = { ...(data.allocations[memberId] || {}), [projectId]: v };
    const members = data.members.map((m) => (m.id === memberId ? { ...m, lastUpdated: todayISO() } : m));
    persist({ ...data, allocations: { ...data.allocations, [memberId]: memberAlloc }, members });
  }

  function totalFor(memberId) {
    const alloc = data.allocations[memberId] || {};
    return Object.values(alloc).reduce((a, b) => a + b, 0);
  }

  function projectTotal(projectId) {
    return data.members.reduce(
      (sum, m) => sum + ((data.allocations[m.id] || {})[projectId] || 0),
      0
    );
  }

  function statusColor(total) {
    if (total > 100) return "var(--c-danger)";
    if (total >= 90) return "var(--c-amber)";
    return "var(--c-teal)";
  }

  function formatDate(iso) {
    if (!iso) return "no date";
    const [y, m, d] = iso.split("-");
    return `${m}/${d}/${y.slice(2)}`;
  }

  function daysSince(iso) {
    if (!iso) return Infinity;
    const then = new Date(iso + "T00:00:00");
    const now = new Date();
    return Math.floor((now - then) / (1000 * 60 * 60 * 24));
  }

  function isStale(iso) {
    return daysSince(iso) > 30;
  }

  function isRequestPending(member) {
    if (!member.lastRequested) return false;
    if (!member.lastUpdated) return true;
    // Fulfilled if the person updated on or after the day the request was made
    return member.lastRequested > member.lastUpdated;
  }

  const currentUser = data.members.find((m) => m.id === currentUserId);
  const selectedProject = data.projects.find((p) => p.id === selectedProjectId);
  const sortedMembers = [...data.members].sort((a, b) => totalFor(b.id) - totalFor(a.id));
  const sortedProjects = [...data.projects].sort((a, b) => projectTotal(b.id) - projectTotal(a.id));
  const staleMembers = data.members.filter((m) => isStale(m.lastUpdated));
  const missingSeedMembers = SEED_MEMBER_NAMES.filter((name) => !data.members.some((m) => m.name === name));

  if (!loaded) {
    return (
      <div className="ct-app">
        <Style />
        <div className="ct-loading">Loading capacity board…</div>
      </div>
    );
  }

  return (
    <div className="ct-app">
      <Style />

      {loadError && (
        <div className="ct-load-error">
          <AlertTriangle size={15} />
          <span>
            Couldn't connect to the team's saved data — you're viewing sample starting data, not your team's real data. Don't make changes until this reconnects, or you may overwrite real data.
          </span>
          <button type="button" className="ct-btn ct-btn-icon" onClick={loadData} title="Try reconnecting">
            Retry
          </button>
        </div>
      )}

      <div className="ct-topbar">
        <div className="ct-brand">
          <div className="ct-brand-mark">
            <Gauge size={20} strokeWidth={2} />
          </div>
          <div>
            <h1>DELTA IT Resource and Project Allocation</h1>
            <p>Team allocation across projects</p>
          </div>
        </div>
        <div className="ct-topbar-actions">
          {mode === "dashboard" && (
            <span className={`ct-save ct-save-${saveState}`}>
              {saveState === "saving" ? "Saving…" : saveState === "saved" ? "Saved" : saveState === "error" ? "Save failed" : ""}
            </span>
          )}
          {mode === "selfserve" && (
            <button
              className="ct-btn ct-mode-toggle"
              onClick={() => {
                setMode("dashboard");
                setSelfServeSubmitted(false);
              }}
            >
              Back to dashboard
            </button>
          )}
        </div>
      </div>

      {mode === "selfserve" ? (
        <SelfServeForm
          data={data}
          selfServeUserId={selfServeUserId}
          setSelfServeUserId={setSelfServeUserId}
          setAllocation={setAllocation}
          setLastUpdated={setLastUpdated}
          totalFor={totalFor}
          statusColor={statusColor}
          saveState={saveState}
          submitted={selfServeSubmitted}
          setSubmitted={setSelfServeSubmitted}
          formatDate={formatDate}
          isRequestPending={isRequestPending}
        />
      ) : (
      <>
      <div className="ct-addrow">
        {addingProject ? (
          <div className="ct-inline-form">
            <input
              autoFocus
              placeholder="Project name"
              value={newProjectName}
              onChange={(e) => setNewProjectName(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === "Enter") {
                  e.preventDefault();
                  addProject(newProjectName);
                } else if (e.key === "Escape") {
                  setAddingProject(false);
                  setNewProjectName("");
                }
              }}
            />
            <button type="button" className="ct-btn ct-btn-icon" onClick={() => addProject(newProjectName)}><Check size={15} /></button>
            <button type="button" className="ct-btn ct-btn-icon" onClick={() => { setAddingProject(false); setNewProjectName(""); }}><X size={15} /></button>
          </div>
        ) : (
          <button className="ct-btn" onClick={() => setAddingProject(true)}>
            <Plus size={14} /> Project
          </button>
        )}

        {addingMember ? (
          <div className="ct-inline-form">
            <input
              autoFocus
              placeholder="Teammate name"
              value={newMemberName}
              onChange={(e) => setNewMemberName(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === "Enter") {
                  e.preventDefault();
                  addMember(newMemberName);
                } else if (e.key === "Escape") {
                  setAddingMember(false);
                  setNewMemberName("");
                }
              }}
            />
            <button type="button" className="ct-btn ct-btn-icon" onClick={() => addMember(newMemberName)}><Check size={15} /></button>
            <button type="button" className="ct-btn ct-btn-icon" onClick={() => { setAddingMember(false); setNewMemberName(""); }}><X size={15} /></button>
          </div>
        ) : (
          <button className="ct-btn" onClick={() => setAddingMember(true)}>
            <Plus size={14} /> Teammate
          </button>
        )}

        {missingSeedMembers.length > 0 && (
          <button
            className="ct-btn ct-restore-btn"
            onClick={restoreMissingMembers}
            title={`Missing: ${missingSeedMembers.join(", ")}`}
          >
            <Plus size={14} /> Restore team ({missingSeedMembers.length})
          </button>
        )}

        <button
          className="ct-btn ct-mode-toggle"
          onClick={() => {
            setMode("selfserve");
            setSelfServeSubmitted(false);
          }}
        >
          Update Allocations
        </button>

        <button className="ct-btn" onClick={exportData} title="Download the current board as a JSON file you keep">
          <Download size={14} /> Export
        </button>

        <button className="ct-btn" onClick={triggerImport} title="Restore the board from a previously exported JSON file">
          <Upload size={14} /> Import
        </button>
        <input
          type="file"
          accept="application/json"
          ref={fileInputRef}
          onChange={handleImportFile}
          style={{ display: "none" }}
        />

        <button
          className="ct-btn"
          onClick={() => {
            setShowBackups((v) => !v);
            if (!showBackups) loadBackupsList();
          }}
          title="See automatic recovery snapshots"
        >
          <History size={14} /> Backups
        </button>

        <button className="ct-btn" onClick={createBackupNow} title="Save a recovery snapshot right now">
          <Check size={14} /> Create backup
        </button>
      </div>

      {restoreNotice && (
        <div className="ct-restore-notice">
          <Check size={13} /> {restoreNotice}
          <button type="button" className="ct-restore-notice-close" onClick={() => setRestoreNotice("")}>
            <X size={12} />
          </button>
        </div>
      )}

      {showBackups && (
        <div className="ct-panel ct-backups-panel">
          <h2>Recovery snapshots</h2>
          <p className="ct-hint" style={{ marginBottom: 10 }}>
            The board automatically saves a snapshot after every change. If something ever looks wrong, you can jump back to any of the last {MAX_BACKUPS} recovery points below.
          </p>
          {backupsLoading ? (
            <p className="ct-hint">Loading snapshots…</p>
          ) : backupsList.length === 0 ? (
            <p className="ct-hint">No snapshots yet — they'll appear here after your first change.</p>
          ) : (
            <div className="ct-backups-list">
              {backupsList.map((b) => {
                const label = new Date(b.ts).toLocaleString();
                return (
                  <div key={b.key} className="ct-backups-row">
                    <span>{label}</span>
                    <button type="button" className="ct-btn ct-btn-icon" onClick={() => restoreBackup(b.key, label)}>
                      Restore
                    </button>
                  </div>
                );
              })}
            </div>
          )}
        </div>
      )}

      {data.members.length === 0 ? (
        <div className="ct-empty">Add your first teammate to start tracking allocation.</div>
      ) : (
        <>

          <section className="ct-panel">
            <div className="ct-identity-row">
              <div className="ct-identity">
                <label htmlFor="ct-who">View resource</label>
                <select id="ct-who" value={currentUserId} onChange={(e) => selectUser(e.target.value)}>
                  <option value="">Team overview</option>
                  {data.members.map((m) => (
                    <option key={m.id} value={m.id}>{m.name}</option>
                  ))}
                </select>
              </div>
              <div className="ct-identity">
                <label htmlFor="ct-proj-select">View project</label>
                <select
                  id="ct-proj-select"
                  value={selectedProjectId}
                  onChange={(e) => setSelectedProjectId(e.target.value)}
                >
                  <option value="">Choose a project…</option>
                  {data.projects.map((p) => (
                    <option key={p.id} value={p.id}>{p.name}</option>
                  ))}
                </select>
              </div>
            </div>

            {currentUser && (
              <div className="ct-editor">
                <div className="ct-editor-head">
                  <h3>{currentUser.name}'s allocation</h3>
                  <span style={{ color: statusColor(totalFor(currentUser.id)) }} className="ct-total">
                    {totalFor(currentUser.id)}%
                  </span>
                </div>

                <div className="ct-updated-row">
                  <label htmlFor="ct-updated">Last updated</label>
                  <input
                    id="ct-updated"
                    type="date"
                    value={currentUser.lastUpdated || ""}
                    onChange={(e) => setLastUpdated(currentUser.id, e.target.value)}
                  />
                  <span className="ct-req-sep">Last requested</span>
                  <span className="ct-req-date">
                    {currentUser.lastRequested ? formatDate(currentUser.lastRequested) : "never"}
                    {currentUser.lastRequested && !isRequestPending(currentUser) && (
                      <span className="ct-req-fulfilled"> · fulfilled</span>
                    )}
                  </span>
                  <button
                    type="button"
                    className="ct-btn ct-btn-icon ct-req-btn"
                    onClick={() => requestUpdate(currentUser.id)}
                    title="Log that an update was requested today"
                  >
                    <Bell size={13} />
                  </button>
                </div>

                {data.projects.length === 0 ? (
                  <p className="ct-hint">Add a project below to allocate time against it.</p>
                ) : (
                  <div className="ct-sliders">
                    {data.projects.map((p) => {
                      const val = (data.allocations[currentUser.id] || {})[p.id] || 0;
                      return (
                        <div key={p.id} className="ct-slider-stack">
                          <div className="ct-slider-stack-top">
                            <span className="ct-dot" style={{ background: p.color }} />
                            <span className="ct-proj-name-full" title={p.name}>{p.name}</span>
                          </div>
                          <div className="ct-slider-stack-bottom">
                            <input
                              type="range"
                              min={0}
                              max={100}
                              step={5}
                              value={val}
                              onChange={(e) => setAllocation(currentUser.id, p.id, Number(e.target.value))}
                            />
                            <input
                              type="number"
                              min={0}
                              max={100}
                              className="ct-num"
                              value={val}
                              onChange={(e) => setAllocation(currentUser.id, p.id, Number(e.target.value))}
                            />
                            <span className="ct-pct">%</span>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                )}

                <AllocMessage total={totalFor(currentUser.id)} />
              </div>
            )}
            {selectedProject && (
              <div className="ct-editor">
                <div className="ct-editor-head">
                  <h3>
                    <span className="ct-dot" style={{ background: selectedProject.color, marginRight: 8, display: "inline-block" }} />
                    {selectedProject.name}
                  </h3>
                  <span className="ct-total">
                    {(projectTotal(selectedProject.id) / 100).toFixed(1)} FTE
                  </span>
                </div>

                {data.members.length === 0 ? (
                  <p className="ct-hint">Add a teammate to allocate time on this project.</p>
                ) : (
                  <div className="ct-sliders">
                    {sortedMembers.map((m) => {
                      const val = (data.allocations[m.id] || {})[selectedProject.id] || 0;
                      return (
                        <div key={m.id} className="ct-slider-row">
                          <span className="ct-dot" style={{ background: "transparent" }} />
                          <span className="ct-proj-name">{m.name}</span>
                          <input
                            type="range"
                            min={0}
                            max={100}
                            step={5}
                            value={val}
                            onChange={(e) => setAllocation(m.id, selectedProject.id, Number(e.target.value))}
                          />
                          <input
                            type="number"
                            min={0}
                            max={100}
                            className="ct-num"
                            value={val}
                            onChange={(e) => setAllocation(m.id, selectedProject.id, Number(e.target.value))}
                          />
                          <span className="ct-pct">%</span>
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
            )}
          </section>

          <section className="ct-panel">
            <h2>Team</h2>
            {data.projects.length === 0 && (
              <p className="ct-hint">Add a project so allocation can be tracked against it.</p>
            )}
            <div className="ct-table-wrap">
              <table className="ct-table">
                <thead>
                  <tr>
                    <th className="ct-th-name">Project</th>
                    {sortedMembers.map((m) => (
                      <th key={m.id} className={m.id === currentUserId ? "ct-th-proj ct-col-self" : "ct-th-proj"}>
                        <div className="ct-th-proj-inner">
                          <div className="ct-th-name-rotated">
                            <span title={m.name}>
                              {m.name}
                              {m.id === currentUserId && <span className="ct-you-inline"> · you</span>}
                            </span>
                          </div>
                          <div className="ct-th-meta-row">
                            <div className={isStale(m.lastUpdated) ? "ct-th-date ct-th-date-stale" : "ct-th-date"} title={isStale(m.lastUpdated) ? "Stale — over 30 days since last update" : "Last updated"}>
                              {isStale(m.lastUpdated) && <AlertTriangle size={9} />} {formatDate(m.lastUpdated)}
                            </div>
                            <button
                              type="button"
                              className={isRequestPending(m) ? "ct-req-bell ct-req-bell-active" : "ct-req-bell"}
                              onClick={() => requestUpdate(m.id)}
                              title={
                                isRequestPending(m)
                                  ? `Update requested ${formatDate(m.lastRequested)} — still waiting on them — click to log a new request`
                                  : m.lastRequested
                                  ? `Last requested ${formatDate(m.lastRequested)} — fulfilled — click to log a new request`
                                  : "Never requested — click to log a request"
                              }
                            >
                              <Bell size={9} />
                            </button>
                          </div>
                        </div>
                      </th>
                    ))}
                    <th className="ct-th-demand" title="Full-time equivalent">FTE</th>
                  </tr>
                </thead>
                <tbody>
                  {data.projects.map((p) => (
                    <tr key={p.id}>
                      <td className="ct-td-name" title={p.name}>
                        <span className="ct-dot" style={{ background: p.color }} />
                        <span className="ct-td-name-text">{p.name}</span>
                        <button className="ct-btn ct-btn-icon ct-remove" onClick={() => removeProject(p.id)} aria-label={`Remove ${p.name}`}>
                          <X size={12} />
                        </button>
                      </td>
                      {sortedMembers.map((m) => {
                        const val = (data.allocations[m.id] || {})[p.id] || 0;
                        return (
                          <td key={m.id} className={m.id === currentUserId ? "ct-td-cell ct-col-self" : "ct-td-cell"}>
                            {val ? `${val}%` : <span className="ct-td-empty">–</span>}
                          </td>
                        );
                      })}
                      <td className="ct-td-demand">{(projectTotal(p.id) / 100).toFixed(1)}</td>
                    </tr>
                  ))}
                </tbody>
                <tfoot>
                  <tr>
                    <td className="ct-td-name ct-foot-label">Total</td>
                    {sortedMembers.map((m) => {
                      const total = totalFor(m.id);
                      return (
                        <td key={m.id} className={m.id === currentUserId ? "ct-td-total ct-col-self" : "ct-td-total"} style={{ color: statusColor(total) }}>
                          {total}%
                        </td>
                      );
                    })}
                    <td className="ct-td-demand"></td>
                  </tr>
                </tfoot>
              </table>
            </div>
            <p className="ct-caption">FTE = full-time equivalent, the sum of allocated % on a project divided by 100.</p>
          </section>

          {data.projects.length > 0 && (
            <section className="ct-panel">
              <h2>Project demand</h2>
              <div className="ct-projects">
                {sortedProjects.map((p) => (
                  <div key={p.id} className="ct-project-row">
                    <span className="ct-dot" style={{ background: p.color }} />
                    <span className="ct-proj-name" title={p.name}>{p.name}</span>
                    <div className="ct-demand-track">
                      <div
                        className="ct-demand-fill"
                        style={{
                          width: `${Math.min(100, (projectTotal(p.id) / (Math.max(data.members.length, 1) * 100)) * 100)}%`,
                          background: p.color,
                        }}
                      />
                    </div>
                    <span className="ct-fte">{(projectTotal(p.id) / 100).toFixed(1)} FTE</span>
                    <button className="ct-btn ct-btn-icon ct-remove" onClick={() => removeProject(p.id)} aria-label={`Remove ${p.name}`}>
                      <X size={13} />
                    </button>
                  </div>
                ))}
              </div>
            </section>
          )}

          {staleMembers.length > 0 && (
            <section className="ct-panel ct-alert-panel">
              <div className="ct-alert">
                <AlertTriangle size={15} />
                <span>
                  {staleMembers.length === 1 ? (
                    <><strong>{staleMembers[0].name}</strong> hasn't updated their allocation in over 30 days ({formatDate(staleMembers[0].lastUpdated)}).</>
                  ) : (
                    <>
                      {staleMembers.length} people haven't updated their allocation in over 30 days:{" "}
                      {staleMembers.map((m, i) => (
                        <span key={m.id}>
                          <strong>{m.name}</strong> ({formatDate(m.lastUpdated)}){i < staleMembers.length - 1 ? ", " : ""}
                        </span>
                      ))}
                    </>
                  )}
                </span>
              </div>
            </section>
          )}
        </>
      )}
      </>
      )}
    </div>
  );
}

function AllocMessage({ total }) {
  if (total > 100) {
    return (
      <div className="ct-msg ct-msg-danger">
        <AlertTriangle size={14} /> {total - 100}% over capacity — trim an allocation.
      </div>
    );
  }
  if (total === 100) {
    return <div className="ct-msg ct-msg-ok">Fully allocated.</div>;
  }
  return <div className="ct-msg ct-msg-muted">{100 - total}% unallocated.</div>;
}

function SelfServeForm({
  data,
  selfServeUserId,
  setSelfServeUserId,
  setAllocation,
  setLastUpdated,
  totalFor,
  statusColor,
  saveState,
  submitted,
  setSubmitted,
  formatDate,
  isRequestPending,
}) {
  const user = data.members.find((m) => m.id === selfServeUserId);

  function handleDateChange(dateStr) {
    setLastUpdated(user.id, dateStr);
  }

  return (
    <div className="ct-selfserve">
      <section className="ct-panel ct-selfserve-panel">
        <h2>Update my allocation</h2>
        <p className="ct-hint" style={{ marginBottom: 16 }}>
          Pick your name, adjust your time on each project, then hit Submit. Your changes save straight to the team board.
        </p>

        <div className="ct-identity">
          <label htmlFor="ct-selfserve-who">I am</label>
          <select
            id="ct-selfserve-who"
            value={selfServeUserId}
            onChange={(e) => {
              setSelfServeUserId(e.target.value);
              setSubmitted(false);
            }}
          >
            <option value="">Select your name…</option>
            {data.members.map((m) => (
              <option key={m.id} value={m.id}>{m.name}</option>
            ))}
          </select>
        </div>

        {!user && (
          <p className="ct-hint" style={{ marginTop: 16 }}>
            Don't see your name? Ask whoever manages the board to add you first.
          </p>
        )}

        {user && (
          <div className="ct-editor">
            <div className="ct-editor-head">
              <h3>{user.name}'s allocation</h3>
              <span style={{ color: statusColor(totalFor(user.id)) }} className="ct-total">
                {totalFor(user.id)}%
              </span>
            </div>

            {isRequestPending(user) && (
              <p className="ct-hint" style={{ marginBottom: 10 }}>
                An update was requested from you on {formatDate(user.lastRequested)}.
              </p>
            )}

            {data.projects.length === 0 ? (
              <p className="ct-hint">No projects have been added yet.</p>
            ) : (
              <div className="ct-sliders">
                {data.projects.map((p) => {
                  const val = (data.allocations[user.id] || {})[p.id] || 0;
                  return (
                    <div key={p.id} className="ct-slider-stack">
                      <div className="ct-slider-stack-top">
                        <span className="ct-dot" style={{ background: p.color }} />
                        <span className="ct-proj-name-full" title={p.name}>{p.name}</span>
                      </div>
                      <div className="ct-slider-stack-bottom">
                        <input
                          type="range"
                          min={0}
                          max={100}
                          step={5}
                          value={val}
                          onChange={(e) => {
                            setAllocation(user.id, p.id, Number(e.target.value));
                            setSubmitted(false);
                          }}
                        />
                        <input
                          type="number"
                          min={0}
                          max={100}
                          className="ct-num"
                          value={val}
                          onChange={(e) => {
                            setAllocation(user.id, p.id, Number(e.target.value));
                            setSubmitted(false);
                          }}
                        />
                        <span className="ct-pct">%</span>
                      </div>
                    </div>
                  );
                })}
              </div>
            )}

            <AllocMessage total={totalFor(user.id)} />

            <div className="ct-selfserve-submitrow">
              <button
                className="ct-btn ct-btn-primary"
                onClick={() => setSubmitted(true)}
              >
                <Check size={15} /> Submit
              </button>
              {submitted && saveState !== "saving" && (
                <span className="ct-selfserve-confirm">
                  <Check size={13} /> Saved — thanks, {user.name}!
                </span>
              )}
              {submitted && saveState === "saving" && (
                <span className="ct-selfserve-confirm">Saving…</span>
              )}
            </div>
          </div>
        )}
      </section>
    </div>
  );
}

function Style() {
  return (
    <style>{`
      .ct-app {
        --bg: #ECEEE7;
        --surface: #FFFFFF;
        --ink: #1B2320;
        --ink-soft: #58635C;
        --line: #D9DCD1;
        --c-teal: #1F6F63;
        --c-amber: #B9791F;
        --c-danger: #C1443C;
        font-family: "IBM Plex Sans", -apple-system, sans-serif;
        color: var(--ink);
        background: linear-gradient(180deg, #DCEBF7 0%, #FFFFFF 55%);
        padding: 18px;
        border-radius: 14px;
        max-width: 1320px;
        margin: 0 auto;
      }
      .ct-app * { box-sizing: border-box; }
      .ct-loading { padding: 40px; text-align: center; color: var(--ink-soft); }
      .ct-load-error {
        display: flex;
        align-items: center;
        gap: 10px;
        background: #FBEBE9;
        border: 1px solid #F3C9C5;
        color: var(--c-danger);
        border-radius: 10px;
        padding: 10px 14px;
        font-size: 12.5px;
        line-height: 1.4;
        margin-bottom: 14px;
      }
      .ct-load-error svg { flex-shrink: 0; }
      .ct-load-error span { flex: 1; }
      .ct-load-error button { border: 1px solid var(--c-danger); color: var(--c-danger); background: #fff; flex-shrink: 0; }
      .ct-load-error button:hover { background: var(--c-danger); color: #fff; }
      .ct-restore-notice {
        display: flex;
        align-items: center;
        gap: 8px;
        background: #E6F1EE;
        color: var(--c-teal);
        border: 1px solid #BFE0D6;
        border-radius: 10px;
        padding: 8px 12px;
        font-size: 12.5px;
        margin: 10px 0;
      }
      .ct-restore-notice-close { margin-left: auto; background: none; border: none; color: var(--c-teal); cursor: pointer; padding: 2px; }
      .ct-backups-panel { margin-top: 12px; }
      .ct-backups-list { display: flex; flex-direction: column; gap: 6px; max-height: 260px; overflow-y: auto; }
      .ct-backups-row { display: flex; align-items: center; justify-content: space-between; padding: 8px 10px; border: 1px solid var(--line); border-radius: 8px; font-size: 12.5px; }
      .ct-topbar { display: flex; justify-content: space-between; align-items: flex-start; gap: 12px; flex-wrap: wrap; }
      .ct-brand { display: flex; align-items: center; gap: 10px; }
      .ct-brand-mark { width: 36px; height: 36px; border-radius: 8px; background: var(--c-teal); color: #fff; display: flex; align-items: center; justify-content: center; }
      .ct-brand h1 { font-family: "Space Grotesk", sans-serif; font-size: 19px; line-height: 1.25; margin: 0; letter-spacing: -0.01em; max-width: 420px; }
      .ct-brand p { margin: 2px 0 0; font-size: 13px; color: var(--ink-soft); }
      .ct-topbar-actions { display: flex; align-items: center; gap: 12px; }
      .ct-mode-toggle { background: var(--c-teal); color: #fff; border-color: var(--c-teal); font-weight: 500; }
      .ct-restore-btn { background: #FDF3E3; border-color: var(--c-amber); color: #8A5A12; font-weight: 500; }
      .ct-restore-btn:hover { background: #FBEAC8; }
      .ct-mode-toggle:hover { background: #185349; border-color: #185349; }
      .ct-selfserve { max-width: 560px; margin: 22px auto 0; }
      .ct-selfserve-panel h2 { margin-bottom: 4px; }
      .ct-selfserve-submitrow { display: flex; align-items: center; gap: 12px; margin-top: 16px; }
      .ct-btn-primary { background: var(--c-teal); color: #fff; border-color: var(--c-teal); font-weight: 500; padding: 9px 16px; }
      .ct-btn-primary:hover { background: #185349; border-color: #185349; }
      .ct-selfserve-confirm { display: inline-flex; align-items: center; gap: 5px; font-size: 12.5px; color: var(--c-teal); font-weight: 500; }
      .ct-save { font-size: 12px; color: var(--ink-soft); font-family: "IBM Plex Mono", monospace; }
      .ct-save-error { color: var(--c-danger); }
      .ct-addrow { display: flex; gap: 8px; margin: 12px 0 4px; flex-wrap: wrap; }
      .ct-btn { display: inline-flex; align-items: center; gap: 6px; background: var(--surface); border: 1px solid var(--line); border-radius: 8px; padding: 7px 12px; font-size: 13px; color: var(--ink); cursor: pointer; font-family: inherit; }
      .ct-btn:hover { border-color: var(--c-teal); }
      .ct-btn-icon { padding: 7px 9px; }
      .ct-remove { border: none; background: transparent; color: var(--ink-soft); padding: 4px; }
      .ct-remove:hover { color: var(--c-danger); }
      .ct-inline-form { display: flex; gap: 6px; align-items: center; }
      .ct-inline-form input { border: 1px solid var(--line); border-radius: 8px; padding: 7px 10px; font-size: 13px; font-family: inherit; min-width: 160px; }
      .ct-empty { margin-top: 14px; padding: 20px; text-align: center; border: 1px dashed var(--line); border-radius: 12px; color: var(--ink-soft); font-size: 14px; }
      .ct-panel { margin-top: 14px; background: var(--surface); border: 1px solid var(--line); border-radius: 12px; padding: 12px 16px; }
      .ct-panel h2 { font-family: "Space Grotesk", sans-serif; font-size: 15px; margin: 0 0 8px; }
      .ct-identity-row { display: flex; align-items: center; gap: 24px; flex-wrap: wrap; }
      .ct-identity { display: flex; align-items: center; gap: 10px; font-size: 13px; }
      .ct-identity label { color: var(--ink-soft); }
      .ct-identity select { border: 1px solid var(--line); border-radius: 8px; padding: 7px 10px; font-family: inherit; font-size: 13px; background: var(--surface); }
      .ct-editor { margin-top: 10px; border-top: 1px solid var(--line); padding-top: 10px; }
      .ct-editor-head { display: flex; justify-content: space-between; align-items: baseline; margin-bottom: 6px; }
      .ct-editor-head h3 { font-size: 14px; margin: 0; font-weight: 600; }
      .ct-total { font-family: "IBM Plex Mono", monospace; font-size: 16px; font-weight: 600; }
      .ct-hint { font-size: 13px; color: var(--ink-soft); }
      .ct-sliders { display: flex; flex-direction: column; gap: 4px; }
      .ct-slider-row { display: grid; grid-template-columns: 10px 110px 1fr 48px 12px; align-items: center; gap: 10px; padding: 2px 0; }
      .ct-slider-stack { display: flex; flex-direction: column; gap: 3px; padding: 5px 0; border-bottom: 1px solid var(--line); }
      .ct-slider-stack:last-child { border-bottom: none; }
      .ct-slider-stack-top { display: flex; align-items: center; gap: 8px; }
      .ct-proj-name-full { font-size: 13px; font-weight: 500; word-break: break-word; overflow-wrap: anywhere; }
      .ct-slider-stack-bottom { display: grid; grid-template-columns: 1fr 48px 12px; align-items: center; gap: 10px; }
      .ct-slider-stack-bottom input[type=range] { width: 100%; accent-color: var(--c-teal); }
      .ct-dot { width: 10px; height: 10px; border-radius: 50%; flex-shrink: 0; }
      .ct-proj-name { font-size: 13px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
      .ct-slider-row input[type=range] { width: 100%; accent-color: var(--c-teal); }
      .ct-num { width: 48px; border: 1px solid var(--line); border-radius: 6px; padding: 4px 6px; font-family: "IBM Plex Mono", monospace; font-size: 12px; text-align: right; }
      .ct-pct { font-size: 12px; color: var(--ink-soft); }
      .ct-msg { margin-top: 8px; font-size: 12px; display: flex; align-items: center; gap: 6px; padding: 6px 10px; border-radius: 8px; }
      .ct-msg-danger { background: #FBEBE9; color: var(--c-danger); }
      .ct-msg-ok { background: #E6F1EE; color: var(--c-teal); }
      .ct-msg-muted { background: #F1F2ED; color: var(--ink-soft); }
      .ct-caption { margin: 4px 2px 0; font-size: 11.5px; color: var(--ink-soft); }
      .ct-updated-row { display: flex; align-items: center; gap: 8px; margin-bottom: 8px; font-size: 12px; color: var(--ink-soft); flex-wrap: wrap; }
      .ct-updated-row input[type=date] { border: 1px solid var(--line); border-radius: 6px; padding: 5px 8px; font-family: inherit; font-size: 12px; background: var(--surface); }
      .ct-req-sep { margin-left: 6px; padding-left: 12px; border-left: 1px solid var(--line); }
      .ct-req-date { font-family: "IBM Plex Mono", monospace; font-size: 12px; color: var(--ink); }
      .ct-req-fulfilled { font-family: inherit; color: var(--c-teal); font-weight: 500; }
      .ct-req-btn { color: var(--ink-soft); }
      .ct-req-btn:hover { color: var(--c-teal); border-color: var(--c-teal); }
      .ct-req-bell { border: none; background: transparent; color: var(--line); padding: 2px; display: inline-flex; align-items: center; justify-content: center; cursor: pointer; margin-top: 2px; }
      .ct-req-bell:hover { color: var(--c-teal); }
      .ct-req-bell-active { color: var(--c-amber); }
      .ct-req-bell-active:hover { color: var(--c-teal); }
      .ct-th-date { font-weight: 400; font-size: 9.5px; color: var(--ink-soft); text-transform: none; white-space: nowrap; }
      .ct-alert-panel { padding: 0; border: none; background: none; }
      .ct-alert { display: flex; align-items: flex-start; gap: 8px; padding: 12px 16px; background: #FBEBE9; color: var(--c-danger); border-radius: 12px; font-size: 12.5px; line-height: 1.5; border: 1px solid #F3C9C5; }
      .ct-alert svg { flex-shrink: 0; margin-top: 2px; }
      .ct-th-date-stale { color: var(--c-danger); display: inline-flex; align-items: center; gap: 3px; }
      .ct-you-inline { color: var(--c-teal); font-weight: 500; }
      .ct-col-self { background: #F5F8F1; }

      .ct-table-wrap {
        overflow-x: auto;
        border: 1px solid var(--line);
        border-radius: 10px;
        scrollbar-width: thin;
        scrollbar-color: var(--line) transparent;
      }
      .ct-table-wrap::-webkit-scrollbar { height: 6px; }
      .ct-table-wrap::-webkit-scrollbar-track { background: transparent; }
      .ct-table-wrap::-webkit-scrollbar-thumb { background: var(--line); border-radius: 6px; }
      .ct-table { border-collapse: collapse; width: 100%; table-layout: fixed; font-size: 12.5px; }
      .ct-table th, .ct-table td { padding: 5px 4px; border-bottom: 1px solid var(--line); text-align: center; }
      .ct-table thead th { background: #F2F3ED; color: var(--ink-soft); font-weight: 500; font-size: 11.5px; position: sticky; top: 0; }
      .ct-th-name {
        width: 148px;
        text-align: left !important;
        position: sticky;
        left: 0;
        background: #F2F3ED;
        z-index: 3;
        padding-right: 10px;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
      }
      .ct-th-proj { width: 40px; position: relative; white-space: nowrap; padding: 0 1px 6px; vertical-align: bottom; }
      .ct-th-proj-inner { display: flex; flex-direction: column; align-items: center; }
      .ct-th-name-rotated { height: 78px; display: flex; align-items: flex-end; justify-content: center; overflow: visible; }
      .ct-th-name-rotated span {
        display: inline-block;
        transform: rotate(-40deg);
        transform-origin: bottom left;
        white-space: nowrap;
        font-size: 12.5px;
        font-weight: 600;
        color: var(--ink);
        letter-spacing: -0.01em;
      }
      .ct-th-meta-row { display: flex; align-items: center; gap: 2px; margin-top: 3px; }
      .ct-th-demand, .ct-td-demand {
        width: 46px;
        position: sticky;
        right: 0;
        z-index: 2;
        font-family: "IBM Plex Mono", monospace;
        font-weight: 600;
        color: var(--ink-soft);
        box-shadow: -6px 0 6px -6px rgba(0,0,0,0.12);
      }
      .ct-th-demand { background: #F2F3ED; z-index: 3; }
      .ct-td-demand { background: var(--surface); }
      .ct-table tbody tr:hover .ct-td-demand { background: #FAFBF7; }
      .ct-td-name {
        text-align: left !important;
        font-weight: 500;
        position: sticky;
        left: 0;
        background: var(--surface);
        z-index: 1;
        display: flex;
        align-items: center;
        gap: 6px;
        padding-right: 10px;
        overflow: hidden;
      }
      .ct-td-name-text { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; min-width: 0; }
      .ct-foot-label { font-weight: 600; }
      .ct-table tfoot td { border-bottom: none; border-top: 1px solid var(--line); }
      .ct-table tfoot .ct-td-demand { background: var(--surface); }
      .ct-td-cell { font-family: "IBM Plex Mono", monospace; color: var(--ink); }
      .ct-td-empty { color: var(--line); }
      .ct-td-total { font-family: "IBM Plex Mono", monospace; font-weight: 600; color: var(--ink-soft); }
      .ct-td-name .ct-remove { opacity: 0; transition: opacity 0.1s; margin-left: auto; flex-shrink: 0; }
      .ct-table tbody tr:hover .ct-remove { opacity: 1; }

      .ct-projects { display: flex; flex-direction: column; gap: 6px; }
      .ct-project-row { display: grid; grid-template-columns: 10px 130px 1fr 56px 20px; align-items: center; gap: 10px; }
      .ct-demand-track { height: 8px; background: #EEF0E8; border-radius: 6px; overflow: hidden; }
      .ct-demand-fill { height: 100%; }
      .ct-fte { font-family: "IBM Plex Mono", monospace; font-size: 12px; color: var(--ink-soft); text-align: right; }
      @media (max-width: 480px) {
        .ct-slider-row { grid-template-columns: 8px 80px 1fr 40px 10px; }
        .ct-member-row { grid-template-columns: 90px 1fr 44px 18px; }
        .ct-project-row { grid-template-columns: 8px 80px 1fr 44px 18px; }
      }
    `}</style>
  );
}
