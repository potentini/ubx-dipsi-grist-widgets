/* global grist, gantt */

const state = {
  options: null,
  tables: [],
  columnsByTable: new Map(),
  suppressWrites: false,
};

const DEFAULT_OPTIONS = {
  levelsCount: 3,
  levels: [
    // Exemple (à configurer)
    // {
    //   name: "Niveau 1",
    //   table: "PROGRAMMES",
    //   idKeyCol: "Code",       // clé métier stable (recommandé)
    //   labelCol: "Libelle",
    //   startCol: "DateDebut",
    //   endCol: "DateFin",      // optionnel
    //   ownerCol: "Responsable",
    //   statusCol: "Statut",
    //   progressCol: "Avancement",
    //   parentKeyCol: null      // pas de parent pour niveau 1
    // }
  ],
  scale: "week"
};

function $(id){ return document.getElementById(id); }

function setStatus(msg){ $("status").textContent = msg; }

async function listTablesAndColumns(){
  // Grist docApi
  const tables = await grist.docApi.listTables();
  state.tables = tables.map(t => t.id);

  state.columnsByTable.clear();
  for (const t of tables) {
    // t.columns contains schema (id, fields). Keep only column ids.
    const cols = (t.columns || []).map(c => c.id);
    state.columnsByTable.set(t.id, cols);
  }
}

function openDrawer(){
  $("settingsDrawer").classList.add("open");
  $("settingsDrawer").setAttribute("aria-hidden", "false");
  $("backdrop").hidden = false;
}
function closeDrawer(){
  $("settingsDrawer").classList.remove("open");
  $("settingsDrawer").setAttribute("aria-hidden", "true");
  $("backdrop").hidden = true;
}

function ensureOptionsShape(opt){
  const merged = structuredClone(DEFAULT_OPTIONS);
  if (opt && typeof opt === "object") {
    if (Number.isFinite(opt.levelsCount)) merged.levelsCount = opt.levelsCount;
    if (Array.isArray(opt.levels)) merged.levels = opt.levels;
    if (opt.scale) merged.scale = opt.scale;
  }
  // normalize levels array length
  if (!Array.isArray(merged.levels) || merged.levels.length !== merged.levelsCount) {
    merged.levels = Array.from({length: merged.levelsCount}, (_, i) => ({
      name: `Niveau ${i+1}`,
      table: "",
      idKeyCol: "",
      parentKeyCol: i === 0 ? null : "",
      labelCol: "",
      startCol: "",
      endCol: "",
      ownerCol: "",
      statusCol: "",
      progressCol: ""
    }));
  }
  return merged;
}

function buildLevelsUI(){
  const levels = state.options.levels;
  const container = $("levelsContainer");
  container.innerHTML = "";

  levels.forEach((lvl, idx) => {
    const card = document.createElement("div");
    card.className = "levelCard";

    const title = document.createElement("h4");
    title.textContent = `${lvl.name || `Niveau ${idx+1}`}`;
    card.appendChild(title);

    const grid = document.createElement("div");
    grid.className = "grid";

    const makeSelect = (label, key, tableForColsFn, full=false) => {
      const wrap = document.createElement("div");
      if (full) wrap.classList.add("full");

      const lab = document.createElement("div");
      lab.className = "small";
      lab.textContent = label;

      const sel = document.createElement("select");
      sel.className = "select";
      sel.dataset.level = String(idx);
      sel.dataset.key = key;

      const options = tableForColsFn();
      sel.appendChild(new Option("—", ""));
      for (const v of options) sel.appendChild(new Option(v, v));
      sel.value = lvl[key] ?? "";

      sel.addEventListener("change", () => {
        const i = Number(sel.dataset.level);
        const k = sel.dataset.key;
        state.options.levels[i][k] = sel.value;
        // If table changed, rebuild this level card to refresh columns list
        if (k === "table") buildLevelsUI();
      });

      wrap.appendChild(lab);
      wrap.appendChild(sel);
      return wrap;
    };

    const makeInput = (label, key, full=false) => {
      const wrap = document.createElement("div");
      if (full) wrap.classList.add("full");

      const lab = document.createElement("div");
      lab.className = "small";
      lab.textContent = label;

      const inp = document.createElement("input");
      inp.className = "input";
      inp.dataset.level = String(idx);
      inp.dataset.key = key;
      inp.value = lvl[key] ?? "";

      inp.addEventListener("input", () => {
        const i = Number(inp.dataset.level);
        const k = inp.dataset.key;
        state.options.levels[i][k] = inp.value;
      });

      wrap.appendChild(lab);
      wrap.appendChild(inp);
      return wrap;
    };

    const tableSelect = makeSelect(
      "Table",
      "table",
      () => state.tables,
      true
    );
    grid.appendChild(tableSelect);

    const cols = state.columnsByTable.get(lvl.table) || [];

    grid.appendChild(makeInput("Nom du niveau (optionnel)", "name", true));
    grid.appendChild(makeSelect("Clé (ID métier) — recommandé", "idKeyCol", () => cols));
    if (idx > 0) grid.appendChild(makeSelect("Colonne Parent (référence clé du parent)", "parentKeyCol", () => cols));
    grid.appendChild(makeSelect("Libellé", "labelCol", () => cols));
    grid.appendChild(makeSelect("Date début", "startCol", () => cols));
    grid.appendChild(makeSelect("Date fin (optionnel)", "endCol", () => cols));
    grid.appendChild(makeSelect("Responsable (optionnel)", "ownerCol", () => cols));
    grid.appendChild(makeSelect("Statut (optionnel)", "statusCol", () => cols));
    grid.appendChild(makeSelect("Avancement (optionnel)", "progressCol", () => cols));

    card.appendChild(grid);
    container.appendChild(card);
  });
}

function parseDate(value){
  if (!value) return null;
  // Grist dates are often ISO or Date objects serialized; Date() handles ISO.
  const d = new Date(value);
  return isNaN(d.getTime()) ? null : d;
}

function addDays(date, days){
  const d = new Date(date);
  d.setDate(d.getDate() + days);
  return d;
}

function normalizeProgress(v){
  if (v == null || v === "") return 0;
  const n = Number(v);
  if (!isFinite(n)) return 0;
  // accept 0..1 or 0..100
  if (n > 1) return Math.max(0, Math.min(1, n / 100));
  return Math.max(0, Math.min(1, n));
}

async function fetchTableRecords(tableId){
  // returns {columns: [...], records: [{id, fields}]}
  return await grist.docApi.fetchTable(tableId);
}

function makeTaskId(levelIndex, key){
  return `${levelIndex}::${String(key)}`;
}

async function buildTasksFromLevels(){
  const levels = state.options.levels;

  // Validate minimal mapping
  for (let i = 0; i < levels.length; i++){
    const l = levels[i];
    if (!l.table || !l.labelCol || !l.startCol) {
      throw new Error(`Paramètres incomplets au niveau ${i+1} (table, libellé, date début requis).`);
    }
    if (!l.idKeyCol) {
      // on autorise sans idKeyCol, mais c’est moins robuste : on tombe sur record id
      // (dans fetchTable, record.id existe toujours)
    }
    if (i > 0 && !l.parentKeyCol) {
      throw new Error(`Paramètres incomplets au niveau ${i+1} (colonne Parent requise).`);
    }
  }

  // Fetch all tables once
  const levelData = [];
  for (const l of levels) {
    const data = await fetchTableRecords(l.table);
    levelData.push(data);
  }

  // Build index of parent keys per level for robust linking
  // For each level: keyValue -> recordId + fields
  const indexes = levelData.map((data, idx) => {
    const l = levels[idx];
    const map = new Map();
    for (const r of data.records) {
      const key = l.idKeyCol ? r.fields[l.idKeyCol] : r.id; // fallback record id
      if (key != null && key !== "") map.set(String(key), r);
    }
    return map;
  });

  const tasks = [];
  const links = [];

  // Root synthetic task (optional) to group everything
  const ROOT_ID = "root";
  tasks.push({
    id: ROOT_ID,
    text: "Portefeuille",
    start_date: gantt.date.date_to_str("%d-%m-%Y")(new Date()),
    duration: 1,
    open: true
  });

  for (let levelIndex = 0; levelIndex < levels.length; levelIndex++){
    const l = levels[levelIndex];
    const data = levelData[levelIndex];

    for (const r of data.records){
      const key = l.idKeyCol ? r.fields[l.idKeyCol] : r.id;
      if (key == null || key === "") continue;

      const label = r.fields[l.labelCol] ?? `(sans libellé)`;
      const start = parseDate(r.fields[l.startCol]);
      if (!start) continue;

      let end = null;
      if (l.endCol) end = parseDate(r.fields[l.endCol]);
      if (!end) end = addDays(start, 1); // durée 1 jour si pas de fin

      const parentTaskId = (() => {
        if (levelIndex === 0) return ROOT_ID;
        const parentKey = r.fields[l.parentKeyCol];
        if (parentKey == null || parentKey === "") return ROOT_ID;
        return makeTaskId(levelIndex - 1, parentKey);
      })();

      const taskId = makeTaskId(levelIndex, key);
      const durationDays = Math.max(1, Math.ceil((end - start) / (24*3600*1000)));

      tasks.push({
        id: taskId,
        text: String(label),
        start_date: gantt.date.date_to_str("%d-%m-%Y")(start),
        duration: durationDays,
        parent: parentTaskId,
        open: true,
        // Custom payload for writing back
        _levelIndex: levelIndex,
        _table: l.table,
        _recordId: r.id,
        _mapping: l,
        _raw: r.fields,
        owner: l.ownerCol ? r.fields[l.ownerCol] : "",
        status: l.statusCol ? r.fields[l.statusCol] : "",
        progress: l.progressCol ? normalizeProgress(r.fields[l.progressCol]) : 0
      });
    }
  }

  return { tasks, links };
}

function initGantt(){
  // Basic config
  gantt.config.autosize = "y";
  gantt.config.fit_tasks = true;
  gantt.config.open_tree_initially = true;

  gantt.config.drag_move = true;
  gantt.config.drag_resize = true;
  gantt.config.drag_progress = false;

  gantt.config.readonly = false;
  gantt.config.show_progress = true;

  // Grid columns (best practice: show only a few, keep readable)
  gantt.config.columns = [
    { name: "text", label: "Élément", tree: true, width: 260, resize: true },
    { name: "owner", label: "Responsable", align: "left", width: 140, resize: true },
    { name: "status", label: "Statut", align: "left", width: 110, resize: true }
  ];

  // Scale presets
  applyScale(state.options?.scale || "week");

  // Tooltip (survol)
  gantt.plugins({ tooltip: true });

  gantt.templates.tooltip_text = function(start, end, task){
    const s = gantt.templates.tooltip_date_format(start);
    const e = gantt.templates.tooltip_date_format(end);
    const owner = task.owner ? `<br><b>Resp.</b> ${escapeHtml(String(task.owner))}` : "";
    const status = task.status ? `<br><b>Statut</b> ${escapeHtml(String(task.status))}` : "";
    return `<b>${escapeHtml(task.text)}</b><br>${s} → ${e}${owner}${status}`;
  };

  // Progress rendering
  gantt.templates.progress_text = function(start, end, task){
    const p = Math.round((task.progress || 0) * 100);
    return `${p}%`;
  };

  // Write-back on drag
  gantt.attachEvent("onAfterTaskDrag", async function(id, mode, e){
    const task = gantt.getTask(id);
    await writeBackDates(task).catch(err => {
      console.error(err);
      setStatus(`Erreur MAJ: ${err.message || err}`);
    });
    return true;
  });

  gantt.init("gantt_here");
}

function escapeHtml(str){
  return str.replace(/[&<>"']/g, (c) => ({
    "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"
  }[c]));
}

function applyScale(scale){
  const today = new Date();
  if (scale === "day") {
    gantt.config.scale_unit = "day";
    gantt.config.date_scale = "%d %M";
    gantt.config.subscales = [{ unit: "hour", step: 6, date: "%Hh" }];
  } else if (scale === "month") {
    gantt.config.scale_unit = "month";
    gantt.config.date_scale = "%F %Y";
    gantt.config.subscales = [{ unit: "week", step: 1, date: "S%W" }];
  } else {
    gantt.config.scale_unit = "week";
    gantt.config.date_scale = "S%W";
    gantt.config.subscales = [{ unit: "day", step: 1, date: "%D %d" }];
  }
  // Keep a sensible start position
  gantt.showDate(today);
}

async function writeBackDates(task){
  // Ignore synthetic root
  if (!task._table || !task._recordId || !task._mapping) return;

  const m = task._mapping;
  const start = task.start_date;
  const end = task.end_date;

  // Convert: DHTMLX uses end_date exclusive sometimes; we’ll keep consistent with duration.
  // We’ll write:
  // - startCol: start date
  // - endCol: end date (if mapped) -> start + duration days
  const payload = {};
  payload[m.startCol] = start;

  if (m.endCol) {
    // write an inclusive-ish end: start + duration days
    // dhtmlx end_date is computed; use start + duration
    payload[m.endCol] = addDays(start, task.duration);
  }

  state.suppressWrites = true;
  await grist.docApi.updateRecord(m.table, task._recordId, payload);
  state.suppressWrites = false;

  setStatus(`Mis à jour: ${m.table} #${task._recordId}`);
}

async function reloadData(){
  setStatus("Chargement…");

  const { tasks, links } = await buildTasksFromLevels();

  // Parse & render
  gantt.clearAll();
  gantt.parse({ data: tasks, links });

  setStatus(`OK — ${tasks.length - 1} éléments`);
}

async function saveOptions(){
  await grist.setOptions(state.options);
  setStatus("Paramètres enregistrés");
}

function wireUI(){
  $("btnSettings").addEventListener("click", openDrawer);
  $("btnCloseSettings").addEventListener("click", closeDrawer);
  $("backdrop").addEventListener("click", closeDrawer);

  $("scale").addEventListener("change", async () => {
    state.options.scale = $("scale").value;
    applyScale(state.options.scale);
    await saveOptions();
    gantt.render();
  });

  $("btnPrev").addEventListener("click", () => gantt.showDate(gantt.date.add(gantt.getState().min_date, -1, gantt.config.scale_unit)));
  $("btnNext").addEventListener("click", () => gantt.showDate(gantt.date.add(gantt.getState().min_date, 1, gantt.config.scale_unit)));
  $("btnToday").addEventListener("click", () => gantt.showDate(new Date()));

  $("btnBuildLevels").addEventListener("click", () => {
    const n = Math.max(1, Math.min(6, Number($("levelsCount").value || 1)));
    state.options.levelsCount = n;
    state.options = ensureOptionsShape(state.options);
    buildLevelsUI();
  });

  $("btnSave").addEventListener("click", async () => {
    await saveOptions();
    closeDrawer();
  });

  $("btnReload").addEventListener("click", async () => {
    await saveOptions();
    await reloadData();
  });
}

async function main(){
  setStatus("Init…");

  // Tell Grist we want full access if needed
  grist.ready({ requiredAccess: 'full' });

  // Load schema for tables/columns listing
  await listTablesAndColumns();

  // Receive options from Grist
  grist.onOptions(async (opt) => {
    state.options = ensureOptionsShape(opt);
    $("levelsCount").value = String(state.options.levelsCount);
    $("scale").value = state.options.scale || "week";

    buildLevelsUI();
    applyScale(state.options.scale || "week");

    // First init if not already
    if (!main._ganttInit) {
      initGantt();
      wireUI();
      main._ganttInit = true;
    }

    // Load data (if mappings exist)
    try {
      // don’t try if the first level isn't configured at all
      const l0 = state.options.levels?.[0];
      if (l0?.table && l0?.labelCol && l0?.startCol) {
        await reloadData();
      } else {
        setStatus("Configurer Paramètres → puis Recharger");
      }
    } catch (e) {
      console.error(e);
      setStatus(`Paramètres requis. (${e.message || e})`);
    }
  });

  // Optional: when records change in the linked table, you can reload.
  // Here we keep it simple (button "Recharger").
}

main().catch(err => {
  console.error(err);
  setStatus(`Erreur: ${err.message || err}`);
});
