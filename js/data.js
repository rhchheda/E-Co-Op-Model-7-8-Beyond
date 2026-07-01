// Data layer: tries the live Google Sheet first, falls back to the bundled
// snapshot in data/seed.json if no sheet is configured or the fetch fails.

const Data = (() => {
  let source = null; // "live" | "fallback"

  async function fetchTab(tabName) {
    const url = `https://docs.google.com/spreadsheets/d/${CONFIG.SHEET_ID}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(tabName)}`;
    const res = await fetch(url, { cache: "no-store" });
    if (!res.ok) throw new Error(`Sheet fetch failed for ${tabName}: ${res.status}`);
    const text = await res.text();
    const match = text.match(/google\.visualization\.Query\.setResponse\(([\s\S]*)\);?\s*$/);
    if (!match) throw new Error(`Unexpected response for ${tabName}`);
    const json = JSON.parse(match[1]);
    const cols = json.table.cols.map((c) => c.label || c.id);
    return json.table.rows.map((r) => {
      const obj = {};
      cols.forEach((c, i) => {
        const cell = r.c[i];
        obj[c] = cell ? (cell.f !== undefined ? cell.f : cell.v) : "";
      });
      return obj;
    });
  }

  function fallbackTeams(seed) {
    return seed.teams.map((t) => {
      const mentors = (t.facultyMentor || "").split("\n").filter(Boolean);
      const names = [];
      const phones = [];
      mentors.forEach((m) => {
        if (m.includes("/")) {
          const [n, p] = m.split("/");
          names.push(n.trim());
          phones.push(p.trim());
        } else {
          names.push(m.trim());
        }
      });
      return {
        id: t.id,
        department: t.department,
        project: t.project,
        students: t.students.slice(0, 6),
        facultyMentor: names.join(", "),
        facultyPhone: phones.join(", "),
        endUsers: t.endUsers,
        description: t.description,
        excitement: t.excitement,
        status: "",
      };
    });
  }

  function fallbackTasks(seed) {
    return seed.tasks.map((t, i) => ({
      sl: i + 1,
      activity: t.activity,
      details: t.details,
      due: t.due,
      owner: t.owner,
      status: t.status,
      outcomes: t.outcomes,
      remarks: t.remarks,
    }));
  }

  function fallbackFaculty(seed) {
    return seed.facultyMentors.map((f) => ({
      name: f.name,
      phone: f.phone,
      teams: f.teams.join(", "),
      count: f.teams.length,
    }));
  }

  function fallbackIndustry(seed) {
    return seed.industryMentors.map((m) => ({
      name: m.name,
      industry: m.industry,
      background: m.background,
      phone: m.phone,
      email: m.email,
      willingness: m.willingness || "Not yet contacted",
    }));
  }

  const MAX_STUDENTS = 6;

  function mapSheetTeams(rows) {
    return rows
      .filter((r) => r["ID"] !== "" && r["ID"] !== undefined)
      .map((r) => {
        const students = [];
        for (let i = 1; i <= MAX_STUDENTS; i++) {
          const name = r[`Student ${i} Name`];
          if (name && String(name).trim()) {
            students.push({ name: name, srn: r[`Student ${i} SRN`] || "", phone: r[`Student ${i} Phone`] || "" });
          }
        }
        return {
          id: r["ID"],
          department: r["Department"],
          project: r["Project Name"],
          students: students,
          facultyMentor: r["Faculty Mentor"],
          facultyPhone: r["Faculty Phone"],
          endUsers: r["End Users"],
          description: r["Description"],
          excitement: r["Excitement / Interest"],
          status: r["Team Status"],
        };
      });
  }

  function mapSheetTasks(rows) {
    return rows
      .filter((r) => r["Task / Activity"])
      .map((r, i) => ({
        sl: r["Sl. No"] || i + 1,
        activity: r["Task / Activity"],
        details: r["Details / Link"],
        due: r["Due Date"],
        owner: r["Owner"],
        status: r["Status"] || "Pending",
        outcomes: r["Outcomes"],
        remarks: r["Remarks"],
      }));
  }

  function mapSheetFaculty(rows) {
    return rows
      .filter((r) => r["Name"])
      .map((r) => ({
        name: r["Name"],
        phone: r["Phone"],
        teams: r["Teams Mentored"],
        count: r["Number of Teams"],
      }));
  }

  function mapSheetIndustry(rows) {
    return rows
      .filter((r) => r["Name"])
      .map((r) => ({
        name: r["Name"],
        industry: r["Industry"],
        background: r["Background"],
        phone: r["Phone"],
        email: r["Email"],
        willingness: r["Willingness to Participate"] || "Not yet contacted",
      }));
  }

  async function load() {
    if (CONFIG.SHEET_ID) {
      try {
        const [teamsRaw, tasksRaw, facultyRaw, industryRaw] = await Promise.all([
          fetchTab("Teams"),
          fetchTab("Tasks"),
          fetchTab("Faculty Mentors"),
          fetchTab("Industry Mentors"),
        ]);
        source = "live";
        return {
          teams: mapSheetTeams(teamsRaw),
          tasks: mapSheetTasks(tasksRaw),
          facultyMentors: mapSheetFaculty(facultyRaw),
          industryMentors: mapSheetIndustry(industryRaw),
        };
      } catch (e) {
        console.warn("Falling back to local data:", e);
      }
    }
    const res = await fetch("data/seed.json", { cache: "no-store" });
    const seed = await res.json();
    source = "fallback";
    return {
      teams: fallbackTeams(seed),
      tasks: fallbackTasks(seed),
      facultyMentors: fallbackFaculty(seed),
      industryMentors: fallbackIndustry(seed),
    };
  }

  function getSource() {
    return source;
  }

  return { load, getSource };
})();
