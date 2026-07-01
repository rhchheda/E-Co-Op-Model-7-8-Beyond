// Shared helpers used across all pages.

function escapeHtml(str) {
  return String(str ?? "").replace(/[&<>"']/g, (c) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  }[c]));
}

function setDataStatusEl(sourceKind) {
  const el = document.getElementById("data-status");
  if (!el) return;
  if (sourceKind === "live") {
    el.textContent = "● Live data from Google Sheet";
    el.className = "live";
  } else {
    el.textContent = "○ Showing bundled snapshot (connect a sheet in js/config.js)";
    el.className = "fallback";
  }
}

// Maps a free-text status value to a badge CSS class.
function statusBadgeClass(status) {
  const s = (status || "").toLowerCase();
  if (["complete", "completed", "on track"].some((k) => s.includes(k))) return "good";
  if (["progress", "at risk"].some((k) => s.includes(k))) return "warn";
  if (["blocked", "behind"].some((k) => s.includes(k))) return "bad";
  return "neutral";
}

function badge(status) {
  const label = status && status.trim() ? status : "Not set";
  return `<span class="badge ${statusBadgeClass(status)}">${escapeHtml(label)}</span>`;
}

function excitementBadgeClass(text) {
  const s = (text || "").toLowerCase();
  if (s.includes("high") || s.includes("excited")) return "good";
  if (s.includes("not sure") || s.includes("no clarity")) return "bad";
  return "neutral";
}

function hideLoadingScreen() {
  const el = document.getElementById("loading-screen");
  if (!el) return;
  el.classList.add("hidden");
  setTimeout(() => { el.style.display = "none"; }, 500);
}

// ---------------- students / duplicate SRN detection ----------------

function studentsDisplayText(t) {
  return (t.students || []).map((s) => `${s.name} (${s.srn}${s.phone ? ", " + s.phone : ""})`).join("\n");
}

function duplicateSrnSet(teams) {
  const counts = {};
  teams.forEach((t) => (t.students || []).forEach((s) => {
    if (!s.srn) return;
    const key = String(s.srn).trim().toLowerCase();
    counts[key] = (counts[key] || 0) + 1;
  }));
  const dupes = new Set();
  Object.keys(counts).forEach((k) => { if (counts[k] > 1) dupes.add(k); });
  return dupes;
}

function teamHasDuplicate(t, dupSet) {
  return (t.students || []).some((s) => s.srn && dupSet.has(String(s.srn).trim().toLowerCase()));
}

function parseDMY(s) {
  const m = String(s || "").trim().match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
  if (!m) return null;
  return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
}

// ---------------- funnel + insights (shared by index.html) ----------------

function computeFunnel(teams) {
  const total = teams.length || 1;
  const stages = [
    { stage: "Registered", count: teams.length, color: "var(--brand)" },
    { stage: "Mentor Assigned", count: teams.filter((t) => t.facultyMentor && String(t.facultyMentor).trim()).length, color: "var(--brand-mid)" },
    { stage: "Status Reported", count: teams.filter((t) => t.status && String(t.status).trim()).length, color: "var(--gold)" },
    { stage: "On Track / Completed", count: teams.filter((t) => ["On Track", "Completed"].includes(t.status)).length, color: "var(--gold-light)" },
    { stage: "Completed", count: teams.filter((t) => t.status === "Completed").length, color: "var(--good)" },
  ];
  return stages.map((s) => Object.assign({}, s, { pct: Math.round((s.count / total) * 100) }));
}

function renderFunnel(data) {
  const wrap = document.getElementById("funnel-wrap");
  if (!wrap) return;
  const funnel = computeFunnel(data.teams);
  const max = Math.max(1, ...funnel.map((f) => f.count));
  const sub = document.getElementById("funnel-sub");
  if (sub) sub.textContent = `${data.teams.length} teams across ${funnel.length} pipeline stages, from registration to MVP completion.`;
  wrap.innerHTML = funnel.map((f) => `
    <div class="funnel-row">
      <div class="funnel-stage" style="background:${f.color}">${escapeHtml(f.stage)}</div>
      <div class="funnel-bar-wrap"><div class="funnel-bar" style="width:${Math.round((f.count / max) * 100)}%"></div></div>
      <div class="funnel-count">${f.count}</div>
      <div class="funnel-pct">${f.pct}%</div>
    </div>`).join("");
}

function computeInsights(data) {
  const teams = data.teams, tasks = data.tasks;
  const insights = [];

  const noStatus = teams.filter((t) => !t.status || !String(t.status).trim()).length;
  if (noStatus > 0) insights.push({ icon: "\u{1F4CB}", title: `${noStatus} Team${noStatus === 1 ? "" : "s"} Need a Status Update`, text: `${noStatus} of ${teams.length} teams don't have a Team Status set yet.`, dark: false });

  const noMentor = teams.filter((t) => !t.facultyMentor || !String(t.facultyMentor).trim()).length;
  if (noMentor > 0) insights.push({ icon: "\u{1F468}‍\u{1F3EB}", title: `${noMentor} Team${noMentor === 1 ? "" : "s"} Missing a Faculty Mentor`, text: `${noMentor} team${noMentor === 1 ? " has" : "s have"} no faculty mentor assigned yet — without one, they can't progress through milestone reviews.`, dark: false });

  const dupSet = duplicateSrnSet(teams);
  if (dupSet.size > 0) insights.push({ icon: "⚠️", title: `${dupSet.size} Duplicate SRN${dupSet.size === 1 ? "" : "s"} Detected`, text: `A student SRN appears on more than one team, or twice within the same team.`, dark: true });

  const overdue = tasks.filter((t) => { const d = parseDMY(t.due); return d && d < new Date() && !(t.status || "").toLowerCase().includes("complete"); }).length;
  if (overdue > 0) insights.push({ icon: "⏰", title: `${overdue} Overdue Task${overdue === 1 ? "" : "s"}`, text: `${overdue} program task${overdue === 1 ? " is" : "s are"} past due and not marked complete.`, dark: true });

  const uncertain = teams.filter((t) => /not sure|no clarity/i.test(t.excitement || "")).length;
  if (uncertain > 0) insights.push({ icon: "\u{1F914}", title: `${uncertain} Team${uncertain === 1 ? "" : "s"} Uncertain About Their Product`, text: `${uncertain} team${uncertain === 1 ? "" : "s"} flagged uncertainty about their product direction — good candidates for extra mentoring time.`, dark: false });

  const byDept = {};
  teams.forEach((t) => { const d = t.department || "Unspecified"; byDept[d] = (byDept[d] || 0) + 1; });
  const deptEntries = Object.entries(byDept).sort((a, b) => b[1] - a[1]);
  if (deptEntries.length) {
    const [topDept, topCount] = deptEntries[0];
    insights.push({ icon: "\u{1F3E2}", title: `${topDept} Leads with ${topCount} Team${topCount === 1 ? "" : "s"}`, text: `Out of ${deptEntries.length} department${deptEntries.length === 1 ? "" : "s"} represented, ${topDept} has the most teams in this cohort.`, dark: false });
  }

  return insights.slice(0, 6);
}

function renderInsights(data) {
  const el = document.getElementById("insights-grid");
  if (!el) return;
  const insights = computeInsights(data);
  el.innerHTML = insights.length
    ? insights.map((ins) => `
      <div class="insight-card ${ins.dark ? "dark" : ""}">
        <div class="insight-icon">${ins.icon}</div>
        <div class="insight-title">${escapeHtml(ins.title)}</div>
        <div class="insight-text">${escapeHtml(ins.text)}</div>
      </div>`).join("")
    : `<div class="empty-state">No notable flags right now &mdash; looking good.</div>`;
}
