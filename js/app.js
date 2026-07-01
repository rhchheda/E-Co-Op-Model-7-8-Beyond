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
