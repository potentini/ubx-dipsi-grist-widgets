(function () {
  const ZOOMS = {
    day: { key: "day", spanDays: 14 },
    week: { key: "week", spanDays: 56 },
    month: { key: "month", spanDays: 365 },
    year: { key: "year", spanDays: 365 * 3 },
    all: { key: "all", spanDays: null }
  };

  const PALETTE = [
    "#0072B2", "#E69F00", "#009E73", "#D55E00", "#CC79A7",
    "#56B4E9", "#F0E442", "#000000", "#4C9A2A", "#8B5CF6",
    "#0EA5E9", "#DC2626"
  ];

  let zoomMode = "day";
  let allRecords = [];
  let parentGroups = [];
  let leafTasks = [];
  let expandedParents = {};
  let globalMinDate = null;
  let globalMaxDate = null;
  let visibleStart = null;
  let visibleEnd = null;

  let colorField = "priority";
  const availableColorFields = [
    "parent", "child", "start", "end",
    "priority", "status", "respPol", "respOp", "respChild", "selector", "order"
  ];

  let labelsVisible = true;
  let childrenOnOneRow = false;

  let currentTableId = null;
  let currentMappingsOk = false;
  let latestMappings = null;

  const mappingInfoEl = document.getElementById("mappingInfo");
  const taskListEl = document.getElementById("taskList");
  const timelineGridEl = document.getElementById("timelineGrid");
  const yearsRowEl = document.getElementById("yearsRow");
  const monthsRowEl = document.getElementById("monthsRow");
  const weeksRowEl = document.getElementById("weeksRow");
  const daysRowEl = document.getElementById("daysRow");
  const timelineHeaderEl = document.getElementById("timelineHeader");
  const timelineBodyEl = document.getElementById("timelineBody");
  const currentPeriodEl = document.getElementById("currentPeriod");
  const colorFieldSelect = document.getElementById("colorFieldSelect");
  const toastContainer = document.getElementById("toastContainer");
  const tooltipEl = document.getElementById("tooltip");
  const ttStartEl = document.getElementById("ttStart");
  const ttEndEl = document.getElementById("ttEnd");
  const ttExtraEl = document.getElementById("ttExtra");
  const dragBubbleEl = document.getElementById("dragBubble");
  const taskCountEl = document.getElementById("taskCount");
  const expandAllBtn = document.getElementById("expandAllBtn");
  const collapseAllBtn = document.getElementById("collapseAllBtn");
  const prevBtn = document.getElementById("prevBtn");
  const nextBtn = document.getElementById("nextBtn");
  const todayBtn = document.getElementById("todayBtn");
  const toggleSidebarBtn = document.getElementById("toggleSidebarBtn");
  const toggleLabelsBtn = document.getElementById("toggleLabelsBtn");
  const groupChildrenBtn = document.getElementById("groupChildrenBtn");
  const ganttContainer = document.getElementById("ganttContainer");

  const dragState = {
    active: false,
    type: null,
    bar: null,
    milestone: null,
    taskId: null,
    parentKey: null,
    originalStart: null,
    originalEnd: null,
    originalMilestoneDate: null,
    originalChildren: null,
    startX: 0,
    pxPerDay: 0
  };

  let sideDragInfo = null;
  const STORAGE_KEY = "grist_gantt_state_v11";

  function loadState() {
    try {
      const raw = window.localStorage.getItem(STORAGE_KEY);
      if (!raw) return;
      const s = JSON.parse(raw);
      if (s.zoomMode) zoomMode = s.zoomMode;
      if (s.colorField) colorField = s.colorField;
      if (typeof s.labelsVisible === "boolean") labelsVisible = s.labelsVisible;
      if (typeof s.childrenOnOneRow === "boolean") childrenOnOneRow = s.childrenOnOneRow;
      if (s.expandedParents && typeof s.expandedParents === "object") {
        expandedParents = s.expandedParents;
      }
    } catch (e) {
      console.warn("Impossible de charger l’état persistant :", e);
    }
  }

  function saveState() {
    try {
      const state = {
        zoomMode,
        colorField,
        labelsVisible,
        childrenOnOneRow,
        expandedParents
      };
      window.localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    } catch (e) {
      console.warn("Impossible de sauvegarder l’état persistant :", e);
    }
  }

  loadState();

  function normalizeDate(value) {
    if (!value) return null;
    if (value instanceof Date) {
      return new Date(value.getFullYear(), value.getMonth(), value.getDate());
    }
    const d = new Date(value);
    if (isNaN(d.getTime())) return null;
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }

  function addDays(date, n) {
    const d = new Date(date.getTime());
    d.setDate(d.getDate() + n);
    return d;
  }

  function diffInDays(a, b) {
    const msPerDay = 24 * 60 * 60 * 1000;
    const da = normalizeDate(a);
    const db = normalizeDate(b);
    return Math.round((db - da) / msPerDay);
  }

  function startOfYear(date) {
    return new Date(date.getFullYear(), 0, 1);
  }

  function endOfYear(date) {
    return new Date(date.getFullYear(), 11, 31);
  }

  function isWeekend(d) {
    const day = d.getDay();
    return day === 0 || day === 6;
  }

  function isSameDay(a, b) {
    return (
      a.getFullYear() === b.getFullYear() &&
      a.getMonth() === b.getMonth() &&
      a.getDate() === b.getDate()
    );
  }

  function formatDate(d) {
    if (!d) return "–";
    return d.toLocaleDateString("fr-FR", {
      day: "2-digit",
      month: "short",
      year: "numeric"
    });
  }

  function formatDateShort(d) {
    if (!d) return "–";
    const day = String(d.getDate());
    const month = String(d.getMonth() + 1);
    const year = String(d.getFullYear()).slice(-2);
    return `${day}/${month}/${year}`;
  }

  function toGristDateString(d) {
    if (!d) return null;
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    return `${y}-${m}-${day}`;
  }

  function isoWeekNumber(date) {
    const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
    const dayNum = d.getUTCDay() || 7;
    d.setUTCDate(d.getUTCDate() + 4 - dayNum);
    const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    return Math.ceil(((d - yearStart) / 86400000 + 1) / 7);
  }

  function showToast(message, type = "info") {
    const el = document.createElement("div");
    el.className =
      "toast " +
      (type === "success" ? "success" : type === "error" ? "error" : "");
    el.textContent = message;
    toastContainer.appendChild(el);
    setTimeout(() => el.remove(), 2600);
  }

  function showTooltip(x, y, title, start, end, extras) {
    tooltipEl.querySelector(".tooltip-title").textContent = title;
    ttStartEl.textContent = formatDate(start);
    ttEndEl.textContent = formatDate(end);

    const lines = [];
    if (extras.parentLabel) lines.push(`<div><span>Parent</span><span>${extras.parentLabel}</span></div>`);
    if (extras.childLabel) lines.push(`<div><span>Enfant</span><span>${extras.childLabel}</span></div>`);
    if (extras.status) lines.push(`<div><span>Statut</span><span>${extras.status}</span></div>`);
    if (extras.priority) lines.push(`<div><span>Priorité</span><span>${extras.priority}</span></div>`);
    if (extras.respPol) lines.push(`<div><span>Réf. pol.</span><span>${extras.respPol}</span></div>`);
    if (extras.respOp) lines.push(`<div><span>Réf. op.</span><span>${extras.respOp}</span></div>`);
    if (extras.respChild) lines.push(`<div><span>Responsable</span><span>${extras.respChild}</span></div>`);
    ttExtraEl.innerHTML = lines.join("");

    tooltipEl.classList.add("visible");
    tooltipEl.style.left = x + 12 + "px";
    tooltipEl.style.top = y + 10 + "px";

    const rect = tooltipEl.getBoundingClientRect();
    const margin = 12;
    let left = x + 12;
    let top = y + 10;

    if (rect.right > window.innerWidth - margin) left = x - rect.width - 12;
    if (left < margin) left = margin;
    if (rect.bottom > window.innerHeight - margin) {
      top = window.innerHeight - rect.height - margin;
    }
    if (top < margin) top = margin;

    tooltipEl.style.left = left + "px";
    tooltipEl.style.top = top + "px";
  }

  function hideTooltip() {
    tooltipEl.classList.remove("visible");
  }

  function showDragBubble(html, x, y) {
    dragBubbleEl.innerHTML = html;
    dragBubbleEl.style.left = x + "px";
    dragBubbleEl.style.top = y + "px";
    dragBubbleEl.classList.add("visible");
  }

  function hideDragBubble() {
    dragBubbleEl.classList.remove("visible");
  }

  function hashStringToInt(str) {
    let h = 2166136261;
    for (let i = 0; i < str.length; i++) {
      h ^= str.charCodeAt(i);
      h = Math.imul(h, 16777619);
    }
    return Math.abs(h);
  }

  function getColorForTask(task) {
    const f = colorField;
    let v = null;

    if (f === "priority") v = task.priority;
    else if (f === "status") v = task.status;
    else if (f === "respPol") v = task.respPol;
    else if (f === "respOp") v = task.respOp;
    else if (f === "respChild") v = task.respChild;
    else if (f === "parent") v = task.parentLabel;
    else if (f === "child") v = task.childLabel;
    else if (f === "start") v = task.startDate ? task.startDate.toISOString().slice(0, 10) : "";
    else if (f === "end") v = task.endDate ? task.endDate.toISOString().slice(0, 10) : "";
    else if (f === "selector") v = task.selector;
    else if (f === "order") v = task.order != null ? String(task.order) : "";

    if (f === "priority") {
      const p = String(v || "");
      if (p.startsWith("1")) return "#ef4444";
      if (p.startsWith("2")) return "#f59e0b";
      if (p.startsWith("3")) return "#3b82f6";
      if (!v) return "#64748b";
    }

    if (f === "status") {
      const s = String(v || "").trim().toLowerCase();
      if (["terminé", "termine", "done", "clos", "clôturé", "cloture"].includes(s)) return "#10b981";
      if (["en cours", "ongoing", "started"].includes(s)) return "#3b82f6";
      if (["bloqué", "bloque", "blocked"].includes(s)) return "#ef4444";
      if (["à faire", "a faire", "todo", "non démarré", "non demarre"].includes(s)) return "#64748b";
    }

    const key = v == null ? "" : String(v);
    const idx = hashStringToInt(key) % PALETTE.length;
    return PALETTE[idx];
  }

  function initColorFieldSelect() {
    colorFieldSelect.innerHTML = "";
    const labels = {
      parent: "Parent",
      child: "Enfant",
      start: "Date début",
      end: "Date fin",
      priority: "Priorité",
      status: "Statut",
      respPol: "Référent politique",
      respOp: "Référent opérationnel",
      respChild: "Responsable enfant",
      selector: "Sélecteur O/N",
      order: "Ordre d’affichage"
    };

    for (const f of availableColorFields) {
      const opt = document.createElement("option");
      opt.value = f;
      opt.textContent = labels[f] || f;
      colorFieldSelect.appendChild(opt);
    }
    if (!availableColorFields.includes(colorField)) colorField = availableColorFields[0];
    colorFieldSelect.value = colorField;
  }

  colorFieldSelect.addEventListener("change", (e) => {
    colorField = e.target.value;
    saveState();
    render();
  });

  function cleanRecordForUpdate(obj) {
    const out = {};
    for (const [key, value] of Object.entries(obj || {})) {
      if (value !== undefined) out[key] = value;
    }
    return out;
  }

  function buildLogicalRecords(records) {
    const result = [];

    for (const raw of records) {
      if (!raw) continue;
      const mapped = grist.mapColumnNames(raw, { mappings: latestMappings });
      if (!mapped) continue;

      const selectorRaw = (mapped.selector || "").toString().trim().toUpperCase();
      if (selectorRaw && selectorRaw !== "O" && selectorRaw !== "1" && selectorRaw !== "TRUE") {
        continue;
      }

      const parentVal = (mapped.parent || "").toString().trim();
      const childVal = (mapped.child || "").toString().trim();
      const startDate = normalizeDate(mapped.start);
      const endDate = normalizeDate(mapped.end);
      const order = mapped.order != null && !isNaN(mapped.order) ? Number(mapped.order) : null;

      const hasParent = !!parentVal;
      const hasChild = !!childVal;
      const kind = hasParent && hasChild ? "hierarchical" : "leaf";

      let isMilestone = false;
      let milestoneDate = null;
      if (!startDate && endDate) {
        isMilestone = true;
        milestoneDate = endDate;
      }

      const parentLabel = hasParent ? parentVal : "";
      const childLabel = hasChild ? childVal : hasParent ? parentVal : "(Sans nom)";

      result.push({
        rowId: raw.id || raw.Id || raw.ID,
        kind,
        parentKey: parentVal || "",
        parentLabel,
        childLabel,
        startDate,
        endDate,
        isMilestone,
        milestoneDate,
        priority: mapped.priority || "",
        status: mapped.status || "",
        respPol: mapped.respPol || "",
        respOp: mapped.respOp || "",
        respChild: mapped.respChild || "",
        selector: mapped.selector || "",
        order
      });
    }

    return result;
  }

  function groupData(records) {
    parentGroups = [];
    leafTasks = [];
    const groupsMap = new Map();

    for (const r of records) {
      if (r.kind === "hierarchical") {
        const key = r.parentKey || "(Sans parent)";
        if (!groupsMap.has(key)) {
          groupsMap.set(key, {
            parentKey: key,
            parentLabel: r.parentLabel || key || "(Sans parent)",
            children: [],
            aggStart: null,
            aggEnd: null,
            hasOnlyMilestones: false,
            onlySingleMilestone: false,
            order: null
          });
        }
        groupsMap.get(key).children.push(r);
      } else {
        leafTasks.push(r);
      }
    }

    for (const g of groupsMap.values()) {
      let minDate = null;
      let maxDate = null;
      let allMilestones = true;
      let minOrder = null;

      for (const c of g.children) {
        if (!c.isMilestone) allMilestones = false;

        if (c.order != null && !isNaN(c.order)) {
          if (minOrder == null || c.order < minOrder) minOrder = c.order;
        }

        const ds = c.startDate || c.milestoneDate || c.endDate;
        const de = c.endDate || c.milestoneDate || c.startDate;

        if (ds && (!minDate || ds < minDate)) minDate = ds;
        if (de && (!maxDate || de > maxDate)) maxDate = de;
      }

      g.aggStart = minDate;
      g.aggEnd = maxDate;
      g.hasOnlyMilestones = allMilestones && g.children.length > 0;
      g.onlySingleMilestone =
        g.children.length === 1 &&
        g.children[0].isMilestone &&
        g.hasOnlyMilestones;
      g.order = minOrder;

      g.children.sort((a, b) => {
        const ao = a.order != null ? a.order : Infinity;
        const bo = b.order != null ? b.order : Infinity;
        if (ao !== bo) return ao - bo;
        return a.childLabel.localeCompare(b.childLabel, "fr");
      });

      parentGroups.push(g);
    }

    parentGroups.sort((a, b) => {
      const ao = a.order != null ? a.order : Infinity;
      const bo = b.order != null ? b.order : Infinity;
      if (ao !== bo) return ao - bo;
      return a.parentLabel.localeCompare(b.parentLabel, "fr");
    });

    leafTasks.sort((a, b) => {
      const ao = a.order != null ? a.order : Infinity;
      const bo = b.order != null ? b.order : Infinity;
      if (ao !== bo) return ao - bo;
      return a.childLabel.localeCompare(b.childLabel, "fr");
    });
  }

  function computeGlobalRange(records) {
    let min = null;
    let max = null;
    for (const r of records) {
      const dates = [];
      if (r.startDate) dates.push(r.startDate);
      if (r.endDate) dates.push(r.endDate);
      if (r.milestoneDate) dates.push(r.milestoneDate);
      for (const d of dates) {
        if (!min || d < min) min = d;
        if (!max || d > max) max = d;
      }
    }
    return { min, max };
  }

  function getNavigationBounds() {
    if (!globalMinDate || !globalMaxDate) {
      return { minAllowed: null, maxAllowed: null };
    }

    if (zoomMode === "all") {
      return {
        minAllowed: new Date(globalMinDate.getFullYear() - 2, 0, 1),
        maxAllowed: new Date(globalMaxDate.getFullYear() + 2, 11, 31)
      };
    }

    const fullSpan = diffInDays(globalMinDate, globalMaxDate) + 1;
    const requested = ZOOMS[zoomMode]?.spanDays || fullSpan;
    const span = Math.max(1, Math.min(requested, Math.max(fullSpan, requested)));
    const marginDays = Math.max(15, span);

    return {
      minAllowed: addDays(globalMinDate, -marginDays),
      maxAllowed: addDays(globalMaxDate, marginDays)
    };
  }

  function getShiftDaysForZoom() {
    if (zoomMode === "day") return 3;
    if (zoomMode === "week") return 14;
    if (zoomMode === "month") return 90;
    if (zoomMode === "year") return 365;
    if (zoomMode === "all") return 365;
    return 30;
  }

  function setVisibleRangeForZoom(centerOnToday = false) {
    if (!globalMinDate || !globalMaxDate) {
      visibleStart = null;
      visibleEnd = null;
      return;
    }

    const { minAllowed, maxAllowed } = getNavigationBounds();

    if (zoomMode === "all") {
      visibleStart = new Date(globalMinDate.getFullYear() - 1, 0, 1);
      visibleEnd = new Date(globalMaxDate.getFullYear() + 1, 11, 31);
      return;
    }

    const fullSpan = diffInDays(globalMinDate, globalMaxDate) + 1;
    const requested = ZOOMS[zoomMode]?.spanDays || fullSpan;
    const span = Math.max(1, requested);

    const center = centerOnToday
      ? normalizeDate(new Date())
      : addDays(globalMinDate, Math.floor(fullSpan / 2));

    const offset = centerOnToday ? Math.floor(span * 0.25) : Math.floor(span / 2);
    let start = addDays(center, -offset);
    let end = addDays(start, span - 1);

    if (start < minAllowed) {
      start = new Date(minAllowed.getTime());
      end = addDays(start, span - 1);
    }
    if (end > maxAllowed) {
      end = new Date(maxAllowed.getTime());
      start = addDays(end, -span + 1);
    }

    visibleStart = start;
    visibleEnd = end;
  }

  function shiftVisibleRange(direction) {
    if (!visibleStart || !visibleEnd) return;

    const { minAllowed, maxAllowed } = getNavigationBounds();
    if (!minAllowed || !maxAllowed) return;

    const step = getShiftDaysForZoom();
    const delta = direction === "left" ? -step : step;
    const span = diffInDays(visibleStart, visibleEnd) + 1;

    let start = addDays(visibleStart, delta);
    let end = addDays(visibleEnd, delta);

    if (zoomMode !== "all") {
      if (diffInDays(start, end) + 1 !== span) {
        end = addDays(start, span - 1);
      }
    }

    if (start < minAllowed) {
      start = new Date(minAllowed.getTime());
      end = addDays(start, span - 1);
    }
    if (end > maxAllowed) {
      end = new Date(maxAllowed.getTime());
      start = addDays(end, -span + 1);
    }

    visibleStart = start;
    visibleEnd = end;
    render();
  }

  function updateZoomButtons() {
    document.querySelectorAll(".zoom-controls .btn").forEach((btn) => {
      btn.classList.toggle("active", btn.dataset.zoom === zoomMode);
    });
  }

  document.querySelectorAll(".zoom-controls .btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      zoomMode = btn.dataset.zoom;
      updateZoomButtons();
      setVisibleRangeForZoom(false);
      saveState();
      render();
    });
  });

  prevBtn.addEventListener("click", () => shiftVisibleRange("left"));
  nextBtn.addEventListener("click", () => shiftVisibleRange("right"));
  todayBtn.addEventListener("click", () => {
    setVisibleRangeForZoom(true);
    render();
  });

  toggleSidebarBtn.addEventListener("click", () => {
    const collapsed = ganttContainer.classList.toggle("sidebar-collapsed");
    toggleSidebarBtn.textContent = collapsed ? "Afficher liste" : "Masquer liste";
  });

  toggleLabelsBtn.addEventListener("click", () => {
    labelsVisible = !labelsVisible;
    toggleLabelsBtn.textContent = labelsVisible ? "Masquer labels" : "Afficher labels";
    saveState();
    render();
  });

  groupChildrenBtn.addEventListener("click", () => {
    childrenOnOneRow = !childrenOnOneRow;
    groupChildrenBtn.textContent = childrenOnOneRow
      ? "Enfants : 1 ligne"
      : "Enfants : multi-lignes";
    saveState();
    render();
  });

  function isGroupExpanded(g) {
    if (g.onlySingleMilestone && !(g.parentKey in expandedParents)) {
      return false;
    }
    if (g.parentKey in expandedParents) {
      return !!expandedParents[g.parentKey];
    }
    return true;
  }

  expandAllBtn.addEventListener("click", () => {
    parentGroups.forEach((g) => {
      if (!g.onlySingleMilestone) expandedParents[g.parentKey] = true;
    });
    saveState();
    render();
  });

  collapseAllBtn.addEventListener("click", () => {
    parentGroups.forEach((g) => {
      expandedParents[g.parentKey] = false;
    });
    saveState();
    render();
  });

  function recomputeCellWidth(totalDays) {
    const containerWidth =
      timelineBodyEl.clientWidth ||
      timelineHeaderEl.clientWidth ||
      window.innerWidth ||
      600;
    const cellWidth = containerWidth / Math.max(1, totalDays);
    document.documentElement.style.setProperty("--cell-width", cellWidth + "px");
    return { containerWidth, cellWidth };
  }

  function buildHeaders() {
    yearsRowEl.innerHTML = "";
    monthsRowEl.innerHTML = "";
    weeksRowEl.innerHTML = "";
    daysRowEl.innerHTML = "";

    yearsRowEl.style.display = "none";
    monthsRowEl.style.display = "none";
    weeksRowEl.style.display = "none";
    daysRowEl.style.display = "none";

    yearsRowEl.style.position = "";
    yearsRowEl.style.width = "";
    yearsRowEl.style.height = "";
    yearsRowEl.style.gridTemplateColumns = "";
    monthsRowEl.style.gridTemplateColumns = "";
    weeksRowEl.style.gridTemplateColumns = "";
    daysRowEl.style.gridTemplateColumns = "";

    if (!visibleStart || !visibleEnd) return;

    const totalDays = diffInDays(visibleStart, visibleEnd) + 1;
    if (totalDays <= 0) return;

    const { containerWidth } = recomputeCellWidth(totalDays);

    const dates = [];
    for (let i = 0; i < totalDays; i++) dates.push(addDays(visibleStart, i));
    const today = normalizeDate(new Date());

    if (zoomMode === "all") {
      yearsRowEl.style.display = "block";
      yearsRowEl.style.position = "relative";
      yearsRowEl.style.width = containerWidth + "px";
      yearsRowEl.style.height = "24px";

      const firstYear = visibleStart.getFullYear();
      const lastYear = visibleEnd.getFullYear();

      for (let year = firstYear; year <= lastYear; year++) {
        const segStart = year === firstYear ? visibleStart : startOfYear(new Date(year, 0, 1));
        const segEnd = year === lastYear ? visibleEnd : endOfYear(new Date(year, 0, 1));

        const leftPx =
          (diffInDays(visibleStart, segStart) / totalDays) * containerWidth;

        const widthPx =
          ((diffInDays(segStart, segEnd) + 1) / totalDays) * containerWidth;

        const cell = document.createElement("div");
        cell.className = "time-cell";
        cell.textContent = String(year);
        cell.style.position = "absolute";
        cell.style.left = leftPx + "px";
        cell.style.width = widthPx + "px";
        cell.style.top = "0";
        cell.style.height = "24px";
        cell.style.display = "flex";
        cell.style.alignItems = "center";
        cell.style.justifyContent = "center";

        yearsRowEl.appendChild(cell);
      }
    } else if (zoomMode === "year") {
      monthsRowEl.style.display = "grid";
      monthsRowEl.style.gridTemplateColumns = `repeat(${totalDays}, var(--cell-width))`;

      let monthStartIndex = 0;
      for (let i = 0; i < dates.length; i++) {
        const isLast = i === dates.length - 1;
        const m = dates[i].getMonth();
        const nextM = !isLast ? dates[i + 1].getMonth() : null;
        if (isLast || nextM !== m) {
          const cell = document.createElement("div");
          cell.className = "time-cell";
          cell.textContent = (m + 1).toString().padStart(2, "0");
          cell.style.gridColumn = `${monthStartIndex + 1} / ${i + 2}`;
          monthsRowEl.appendChild(cell);
          monthStartIndex = i + 1;
        }
      }
    } else if (zoomMode === "month") {
      monthsRowEl.style.display = "grid";
      weeksRowEl.style.display = "grid";

      monthsRowEl.style.gridTemplateColumns = `repeat(${totalDays}, var(--cell-width))`;
      weeksRowEl.style.gridTemplateColumns = `repeat(${totalDays}, var(--cell-width))`;

      let monthStartIndex = 0;
      for (let i = 0; i < dates.length; i++) {
        const isLast = i === dates.length - 1;
        const m = dates[i].getMonth();
        const nextM = !isLast ? dates[i + 1].getMonth() : null;
        if (isLast || nextM !== m) {
          const cell = document.createElement("div");
          cell.className = "time-cell";
          cell.textContent = dates[i].toLocaleDateString("fr-FR", {
            month: "short",
            year: "numeric"
          });
          cell.style.gridColumn = `${monthStartIndex + 1} / ${i + 2}`;
          monthsRowEl.appendChild(cell);
          monthStartIndex = i + 1;
        }
      }

      let weekStartIndex = 0;
      let currentWeek = isoWeekNumber(dates[0]);
      let currentYear = dates[0].getFullYear();

      for (let i = 0; i < dates.length; i++) {
        const isLast = i === dates.length - 1;
        const w = isoWeekNumber(dates[i]);
        const y = dates[i].getFullYear();
        const changes = w !== currentWeek || y !== currentYear;

        if (changes || isLast) {
          const endIndex = changes ? i : i + 1;
          const cell = document.createElement("div");
          cell.className = "time-cell";
          cell.textContent = "S" + currentWeek.toString().padStart(2, "0");
          cell.style.gridColumn = `${weekStartIndex + 1} / ${endIndex + 1}`;
          weeksRowEl.appendChild(cell);

          currentWeek = w;
          currentYear = y;
          weekStartIndex = i;
        }
      }
    } else {
      monthsRowEl.style.display = "grid";
      daysRowEl.style.display = "grid";

      monthsRowEl.style.gridTemplateColumns = `repeat(${totalDays}, var(--cell-width))`;
      daysRowEl.style.gridTemplateColumns = `repeat(${totalDays}, var(--cell-width))`;

      let monthStartIndex = 0;
      for (let i = 0; i < dates.length; i++) {
        const isLast = i === dates.length - 1;
        const m = dates[i].getMonth();
        const nextM = !isLast ? dates[i + 1].getMonth() : null;
        if (isLast || nextM !== m) {
          const cell = document.createElement("div");
          cell.className = "time-cell";
          cell.textContent = dates[i].toLocaleDateString("fr-FR", {
            month: "short",
            year: "numeric"
          });
          cell.style.gridColumn = `${monthStartIndex + 1} / ${i + 2}`;
          monthsRowEl.appendChild(cell);
          monthStartIndex = i + 1;
        }
      }

      for (const d of dates) {
        const cell = document.createElement("div");
        const weekend = isWeekend(d);
        const isTodayFlag = isSameDay(d, today);
        cell.className =
          "time-cell " +
          (weekend ? " weekend" : "") +
          (isTodayFlag ? " today" : "");
        cell.textContent = d.getDate().toString().padStart(2, "0");
        daysRowEl.appendChild(cell);
      }
    }

    currentPeriodEl.textContent = `${formatDate(visibleStart)} – ${formatDate(visibleEnd)}`;
  }

  function buildTracks() {
    const tracks = [];

    if (!childrenOnOneRow) {
      parentGroups.forEach((g) => {
        tracks.push({ kind: "parent", group: g });
        const expanded = isGroupExpanded(g);
        if (expanded) {
          g.children.forEach((c) => {
            tracks.push({ kind: "child", group: g, task: c });
          });
        }
      });
    } else {
      parentGroups.forEach((g) => {
        const expanded = isGroupExpanded(g);
        if (expanded && g.children.length) {
          tracks.push({ kind: "groupChildren", group: g });
        }
      });
    }

    leafTasks.forEach((r) => {
      tracks.push({ kind: "leaf", task: r });
    });

    return tracks;
  }

  function buildSidebarMeta(task) {
    const parts = [];
    if (task.respOp) parts.push(`Resp. op: ${task.respOp}`);
    if (task.status) parts.push(`Stat.: ${task.status}`);
    if (task.startDate || task.endDate || task.milestoneDate) {
      parts.push(
        `${formatDateShort(task.startDate || task.milestoneDate)} – ${formatDateShort(
          task.endDate || task.milestoneDate
        )}`
      );
    }
    return parts.join(" · ") || "–";
  }

  function renderTaskList() {
    taskListEl.innerHTML = "";

    const tracks = buildTracks();
    if (!tracks.length) {
      const div = document.createElement("div");
      div.className = "empty";
      div.textContent = "Aucune ligne à afficher.";
      taskListEl.appendChild(div);
      taskCountEl.textContent = "";
      return;
    }

    let rowCount = 0;

    for (const t of tracks) {
      if (t.kind === "parent") {
        const g = t.group;
        const parentRow = document.createElement("div");
        parentRow.className = "task-row parent-row";
        parentRow.draggable = true;
        parentRow.dataset.kind = "parent";
        parentRow.dataset.parentKey = g.parentKey;

        const toggle = document.createElement("button");
        toggle.type = "button";
        toggle.className = "parent-toggle";
        const expanded = isGroupExpanded(g);
        toggle.textContent = expanded ? "▾" : "▸";
        toggle.addEventListener("click", (ev) => {
          ev.stopPropagation();
          expandedParents[g.parentKey] = !expanded;
          saveState();
          render();
        });

        const info = document.createElement("div");
        info.className = "task-info";

        const main = document.createElement("div");
        main.className = "task-name";
        main.textContent = g.parentLabel;

        const meta = document.createElement("div");
        meta.className = "task-meta";
        meta.textContent =
          g.aggStart && g.aggEnd
            ? `${formatDateShort(g.aggStart)} – ${formatDateShort(g.aggEnd)}`
            : "Aucune date";

        info.appendChild(main);
        info.appendChild(meta);
        parentRow.appendChild(toggle);
        parentRow.appendChild(info);

        attachSideDragHandlers(parentRow);
        taskListEl.appendChild(parentRow);
        rowCount++;
      } else if (t.kind === "child") {
        const c = t.task;
        const row = document.createElement("div");
        row.className = "task-row child-row";
        row.draggable = true;
        row.dataset.kind = "child";
        row.dataset.parentKey = t.group.parentKey;
        row.dataset.rowId = String(c.rowId);

        const spacer = document.createElement("span");
        spacer.style.display = "inline-block";
        spacer.style.width = "16px";

        const info = document.createElement("div");
        info.className = "task-info";

        const main = document.createElement("div");
        main.className = "task-name";
        main.textContent = c.childLabel;

        const meta = document.createElement("div");
        meta.className = "task-meta";
        meta.textContent = buildSidebarMeta(c);

        info.appendChild(main);
        info.appendChild(meta);
        row.appendChild(spacer);
        row.appendChild(info);

        attachSideDragHandlers(row);
        taskListEl.appendChild(row);
        rowCount++;
      } else if (t.kind === "groupChildren") {
        const g = t.group;
        const row = document.createElement("div");
        row.className = "task-row parent-row";
        row.draggable = true;
        row.dataset.kind = "parent";
        row.dataset.parentKey = g.parentKey;

        const info = document.createElement("div");
        info.className = "task-info";

        const main = document.createElement("div");
        main.className = "task-name";
        main.textContent = g.parentLabel;

        const meta = document.createElement("div");
        meta.className = "task-meta";
        meta.textContent = `${g.children.length} élément(s) enfant`;

        info.appendChild(main);
        info.appendChild(meta);
        row.appendChild(info);

        attachSideDragHandlers(row);
        taskListEl.appendChild(row);
        rowCount++;
      } else if (t.kind === "leaf") {
        const r = t.task;
        const row = document.createElement("div");
        row.className = "task-row child-row";
        row.draggable = true;
        row.dataset.kind = "leaf";
        row.dataset.rowId = String(r.rowId);

        const spacer = document.createElement("span");
        spacer.style.display = "inline-block";
        spacer.style.width = "16px";

        const info = document.createElement("div");
        info.className = "task-info";

        const main = document.createElement("div");
        main.className = "task-name";
        main.textContent = r.childLabel;

        const meta = document.createElement("div");
        meta.className = "task-meta";
        meta.textContent = buildSidebarMeta(r);

        info.appendChild(main);
        info.appendChild(meta);
        row.appendChild(spacer);
        row.appendChild(info);

        attachSideDragHandlers(row);
        taskListEl.appendChild(row);
        rowCount++;
      }
    }

    taskCountEl.textContent = `${rowCount} lignes`;
  }

  function attachSideDragHandlers(rowEl) {
    rowEl.addEventListener("dragstart", (e) => {
      const kind = rowEl.dataset.kind;
      if (kind !== "child" && kind !== "leaf" && kind !== "parent") return;

      sideDragInfo = {
        kind,
        parentKey: rowEl.dataset.parentKey || null,
        rowId: rowEl.dataset.rowId ? Number(rowEl.dataset.rowId) : null
      };

      rowEl.classList.add("dragging");
      e.dataTransfer.effectAllowed = "move";
    });

    rowEl.addEventListener("dragend", () => {
      sideDragInfo = null;
      document.querySelectorAll(".task-row").forEach((r) => {
        r.classList.remove("dragging", "drag-over");
      });
    });

    rowEl.addEventListener("dragover", (e) => {
      if (!sideDragInfo) return;

      const targetKind = rowEl.dataset.kind;
      if (targetKind !== sideDragInfo.kind) return;

      if (targetKind === "child") {
        if ((rowEl.dataset.parentKey || "") !== (sideDragInfo.parentKey || "")) return;
      }

      e.preventDefault();
      document.querySelectorAll(".task-row").forEach((r) => {
        r.classList.remove("drag-over");
      });
      rowEl.classList.add("drag-over");
    });

    rowEl.addEventListener("drop", async (e) => {
      e.preventDefault();
      if (!sideDragInfo) return;

      const targetKind = rowEl.dataset.kind;
      if (targetKind !== sideDragInfo.kind) return;

      try {
        if (targetKind === "child") {
          const targetRowId = Number(rowEl.dataset.rowId);
          if ((rowEl.dataset.parentKey || "") !== (sideDragInfo.parentKey || "")) return;
          if (targetRowId === sideDragInfo.rowId) return;
          await reorderChildrenWithinParent(
            sideDragInfo.parentKey,
            sideDragInfo.rowId,
            targetRowId
          );
          showToast("Ordre mis à jour", "success");
        } else if (targetKind === "leaf") {
          const targetRowId = Number(rowEl.dataset.rowId);
          if (targetRowId === sideDragInfo.rowId) return;
          await reorderLeafTasks(sideDragInfo.rowId, targetRowId);
          showToast("Ordre mis à jour", "success");
        } else if (targetKind === "parent") {
          const targetParentKey = rowEl.dataset.parentKey || "";
          if (targetParentKey === sideDragInfo.parentKey) return;
          await reorderParentGroups(sideDragInfo.parentKey, targetParentKey);
          showToast("Ordre des parents mis à jour", "success");
        }
      } catch (err) {
        console.error(err);
        showToast("Erreur lors de la mise à jour de l’ordre", "error");
      } finally {
        sideDragInfo = null;
        document.querySelectorAll(".task-row").forEach((r) => {
          r.classList.remove("dragging", "drag-over");
        });
      }
    });
  }

  async function refreshTableInfo() {
    try {
      currentTableId = await grist.selectedTable.getTableId();
    } catch (e) {
      currentTableId = null;
    }

    const mappedCols =
      latestMappings && latestMappings.columns
        ? Object.keys(latestMappings.columns).length
        : latestMappings
        ? Object.keys(latestMappings).length
        : 0;

    mappingInfoEl.textContent =
      "Mapping actif : " +
      (currentMappingsOk ? "oui" : "non") +
      ", table = " +
      (currentTableId || "inconnue") +
      ", mappings reçus = " +
      mappedCols;
  }

  function buildBackPayload(aliasValues) {
    if (!latestMappings) {
      throw new Error("Aucun mapping courant disponible pour remapper les colonnes.");
    }

    const mapped = grist.mapColumnNamesBack(aliasValues, {
      mappings: latestMappings
    });

    if (!mapped || typeof mapped !== "object") {
      throw new Error("Le remappage inverse des colonnes a échoué.");
    }

    const cleaned = cleanRecordForUpdate(mapped);
    const { id, ...fields } = cleaned;

    if (id == null) {
      throw new Error("Payload sans id.");
    }

    return { id, fields };
  }

  async function updateRows(records) {
    if (!records) return;

    const arr = Array.isArray(records) ? records : [records];
    const cleaned = arr.filter(
      (r) => r && r.id != null && r.fields && Object.keys(r.fields).length > 0
    );

    if (!cleaned.length) {
      console.warn("[GANTT DEBUG] Aucun champ modifiable à envoyer à Grist.");
      return;
    }

    await grist.selectedTable.update(cleaned);
  }

  async function updateChildDates(rowId, startDate, endDate) {
    if (!rowId) return;
    const aliasValues = { id: rowId };
    if (startDate) aliasValues.start = toGristDateString(startDate);
    if (endDate) aliasValues.end = toGristDateString(endDate);
    const payload = buildBackPayload(aliasValues);
    await updateRows(payload);
  }

  async function updateMilestoneDate(rowId, newDate) {
    if (!rowId) return;
    const aliasValues = { id: rowId };
    if (newDate) aliasValues.end = toGristDateString(newDate);
    const payload = buildBackPayload(aliasValues);
    await updateRows(payload);
  }

  async function moveParentGroup(parentKey, originalChildren, deltaDays) {
    if (!parentKey || !originalChildren || !deltaDays) return;

    const updates = [];
    for (const c of originalChildren) {
      const aliasValues = { id: c.rowId };
      if (c.startDate) aliasValues.start = toGristDateString(addDays(c.startDate, deltaDays));
      if (c.endDate) aliasValues.end = toGristDateString(addDays(c.endDate, deltaDays));
      updates.push(buildBackPayload(aliasValues));
    }
    if (updates.length) await updateRows(updates);
  }

  async function reorderChildrenWithinParent(parentKey, sourceRowId, targetRowId) {
    const g = parentGroups.find((pg) => pg.parentKey === parentKey);
    if (!g) return;

    const arr = g.children.slice();
    const fromIndex = arr.findIndex((c) => c.rowId === sourceRowId);
    const toIndex = arr.findIndex((c) => c.rowId === targetRowId);
    if (fromIndex < 0 || toIndex < 0) return;

    const [moved] = arr.splice(fromIndex, 1);
    arr.splice(toIndex, 0, moved);

    const updates = arr.map((c, idx) =>
      buildBackPayload({
        id: c.rowId,
        order: idx + 1
      })
    );

    await updateRows(updates);
  }

  async function reorderLeafTasks(sourceRowId, targetRowId) {
    const arr = leafTasks.slice();
    const fromIndex = arr.findIndex((c) => c.rowId === sourceRowId);
    const toIndex = arr.findIndex((c) => c.rowId === targetRowId);
    if (fromIndex < 0 || toIndex < 0) return;

    const [moved] = arr.splice(fromIndex, 1);
    arr.splice(toIndex, 0, moved);

    const updates = arr.map((c, idx) =>
      buildBackPayload({
        id: c.rowId,
        order: idx + 1
      })
    );

    await updateRows(updates);
  }

  async function reorderParentGroups(sourceParentKey, targetParentKey) {
    const groups = parentGroups.slice();
    const fromIndex = groups.findIndex((g) => g.parentKey === sourceParentKey);
    const toIndex = groups.findIndex((g) => g.parentKey === targetParentKey);
    if (fromIndex < 0 || toIndex < 0) return;

    const [moved] = groups.splice(fromIndex, 1);
    groups.splice(toIndex, 0, moved);

    const updates = [];
    let baseOrder = 1;

    for (const g of groups) {
      for (const child of g.children) {
        updates.push(
          buildBackPayload({
            id: child.rowId,
            order: baseOrder++
          })
        );
      }
    }

    for (const leaf of leafTasks) {
      updates.push(
        buildBackPayload({
          id: leaf.rowId,
          order: baseOrder++
        })
      );
    }

    await updateRows(updates);
  }

  function renderTimeline() {
    timelineGridEl.innerHTML = "";
    if (!visibleStart || !visibleEnd) return;

    const tracks = buildTracks();
    if (!tracks.length) return;

    const totalDays = diffInDays(visibleStart, visibleEnd) + 1;
    if (totalDays <= 0) return;

    const { containerWidth } = recomputeCellWidth(totalDays);
    timelineGridEl.style.width = containerWidth + "px";

    const rowHeight = 34;
    const totalHeight = tracks.length * rowHeight;
    timelineGridEl.style.height = totalHeight + "px";
    timelineGridEl.style.minHeight = totalHeight + "px";
    timelineBodyEl.style.height = totalHeight + "px";
    timelineBodyEl.style.minHeight = totalHeight + "px";

    const today = normalizeDate(new Date());
    const milestoneLabelByTrack = new Map();

    function dateToFrac(d) {
      if (!d) return null;
      const clamped =
        d < visibleStart ? visibleStart : d > visibleEnd ? visibleEnd : d;
      const daysFromStart = diffInDays(visibleStart, clamped);
      return daysFromStart / totalDays;
    }

    function extrasFromTask(task, isParent) {
      return {
        parentLabel: task.parentLabel,
        childLabel: isParent ? "" : task.childLabel,
        status: task.status || "",
        priority: task.priority || "",
        respPol: task.respPol || "",
        respOp: task.respOp || "",
        respChild: isParent ? "" : (task.respChild || "")
      };
    }

    function addGroupLabelAtRightmostVisibleItem(trackIndex, group) {
      if (!labelsVisible || !group || !group.children || !group.children.length) return;

      const centerY = trackIndex * rowHeight + rowHeight / 2;
      let rightmostX = null;

      for (const c of group.children) {
        if (c.isMilestone && c.milestoneDate && !c.startDate) {
          if (c.milestoneDate < visibleStart || c.milestoneDate > visibleEnd) continue;
          const frac = dateToFrac(c.milestoneDate);
          if (frac == null) continue;
          const x = frac * containerWidth;
          if (rightmostX == null || x > rightmostX) rightmostX = x;
          continue;
        }

        const start = c.startDate || c.milestoneDate || c.endDate;
        const end = c.endDate || start;
        if (!start || !end) continue;

        const s = normalizeDate(start);
        const e = normalizeDate(end);
        if (e < visibleStart || s > visibleEnd) continue;

        const rightFrac = dateToFrac(e);
        if (rightFrac == null) continue;
        const x = (rightFrac + 1 / totalDays) * containerWidth;
        if (rightmostX == null || x > rightmostX) rightmostX = x;
      }

      if (rightmostX == null) return;

      const label = document.createElement("span");
      label.className = "group-row-label";
      label.textContent = group.parentLabel;
      label.style.left = (rightmostX + 20) + "px";
      label.style.top = centerY + "px";

      timelineGridEl.appendChild(label);
    }

    function getMilestoneLabelYOffset(trackIndex, x) {
      const placements = milestoneLabelByTrack.get(trackIndex) || [];
      const threshold = 42;
      let level = 0;

      for (const p of placements) {
        if (Math.abs(p.x - x) < threshold && p.level === level) {
          level++;
        }
      }

      placements.push({ x, level });
      milestoneLabelByTrack.set(trackIndex, placements);

      const offsets = [0, -10, 10, -18, 18];
      return offsets[Math.min(level, offsets.length - 1)];
    }

    function addMilestone(trackIndex, task, isParent) {
      const frac = dateToFrac(task.milestoneDate);
      if (frac == null) return;

      const x = frac * containerWidth;
      const centerY = trackIndex * rowHeight + rowHeight / 2;

      const m = document.createElement("div");
      m.className = "gantt-milestone";
      m.style.left = x.toFixed(1) + "px";
      m.style.top = centerY.toFixed(1) + "px";
      m.style.background = getColorForTask(task);
      m.style.border = "1.5px solid #000";

      m.dataset.role = isParent ? "parent-milestone" : "child-milestone";
      if (!isParent) m.dataset.rowId = task.rowId;

      const label = document.createElement("span");
      label.className = "milestone-label";
      label.textContent = isParent ? task.parentLabel : task.childLabel;

      const hideMilestoneLabel =
        !labelsVisible || (childrenOnOneRow && !isParent);

      if (hideMilestoneLabel) {
        label.style.display = "none";
      }

      const yOffset = getMilestoneLabelYOffset(trackIndex, x);
      label.style.left = (x + 12) + "px";
      label.style.top = (centerY + yOffset) + "px";

      const extras = extrasFromTask(task, isParent);

      m.addEventListener("mousemove", (ev) => {
        showTooltip(
          ev.clientX,
          ev.clientY,
          isParent ? task.parentLabel : task.childLabel,
          task.milestoneDate,
          task.milestoneDate,
          extras
        );
      });
      m.addEventListener("mouseenter", (ev) => {
        showTooltip(
          ev.clientX,
          ev.clientY,
          isParent ? task.parentLabel : task.childLabel,
          task.milestoneDate,
          task.milestoneDate,
          extras
        );
      });
      m.addEventListener("mouseleave", hideTooltip);

      attachMilestoneDrag(m, task, isParent);
      timelineGridEl.appendChild(m);
      timelineGridEl.appendChild(label);
    }

    function addBar(trackIndex, task, opts) {
      const { isParent, parentKey, forceLabelText, hideLabel } = opts || {};

      if (task.isMilestone && task.milestoneDate && !task.startDate) {
        addMilestone(trackIndex, task, isParent);
        return;
      }

      const start = task.startDate || task.milestoneDate || task.endDate;
      const end = task.endDate || start;
      if (!start || !end) return;

      const s = normalizeDate(start);
      const e = normalizeDate(end);
      if (e < visibleStart || s > visibleEnd) return;

      const leftFrac = dateToFrac(s);
      const rightFrac = dateToFrac(e);
      if (leftFrac == null || rightFrac == null) return;

      const widthFrac = Math.max(0.01, (rightFrac - leftFrac) + 1 / totalDays);
      const leftPx = leftFrac * containerWidth;
      const widthPx = widthFrac * containerWidth;

      const bar = document.createElement("div");
      bar.className = "gantt-bar" + (isParent ? " parent" : "");
      bar.style.left = leftPx.toFixed(1) + "px";
      bar.style.width = widthPx.toFixed(1) + "px";
      bar.style.top = trackIndex * rowHeight + 8 + "px";

      bar.dataset.role = isParent ? "parent" : "child";
      if (isParent) {
        bar.dataset.parentKey = parentKey;
        bar.dataset.title = task.parentLabel;
      } else {
        bar.dataset.rowId = task.rowId;
        bar.dataset.title = task.childLabel;
      }
      bar.dataset.start = s.toISOString();
      bar.dataset.end = e.toISOString();
      bar.style.background = getColorForTask(task);

      const labelText = forceLabelText || (isParent ? task.parentLabel : task.childLabel);

      if (!hideLabel) {
        const labelSpan = document.createElement("span");
        labelSpan.textContent = labelText;
        labelSpan.className = widthPx >= 110 ? "bar-label inside" : "bar-label outside";
        if (!labelsVisible) labelSpan.style.display = "none";
        bar.appendChild(labelSpan);
      }

      const extras = extrasFromTask(task, isParent);

      bar.addEventListener("mousemove", (ev) => {
        setBarCursor(bar, ev);
        showTooltip(ev.clientX, ev.clientY, bar.dataset.title, s, e, extras);
      });
      bar.addEventListener("mouseenter", (ev) => {
        setBarCursor(bar, ev);
        showTooltip(ev.clientX, ev.clientY, bar.dataset.title, s, e, extras);
      });
      bar.addEventListener("mouseleave", () => {
        bar.style.cursor = "default";
        hideTooltip();
      });

      attachBarDrag(bar);
      timelineGridEl.appendChild(bar);
    }

    for (let t = 0; t < tracks.length; t++) {
      const row = document.createElement("div");
      row.className = "grid-row";
      for (let i = 0; i < totalDays; i++) {
        const d = addDays(visibleStart, i);
        const cell = document.createElement("div");
        cell.className = "grid-cell" + (isWeekend(d) ? " weekend" : "");
        row.appendChild(cell);
      }
      timelineGridEl.appendChild(row);
    }

    const todayDiff = diffInDays(visibleStart, today);
    if (todayDiff >= 0 && todayDiff < totalDays) {
      const line = document.createElement("div");
      line.className = "today-line";
      const x = todayDiff * (containerWidth / Math.max(1, totalDays));
      line.style.left = x + "px";
      timelineGridEl.appendChild(line);
    }

    let trackIndex = 0;
    for (const t of tracks) {
      if (t.kind === "parent") {
        const g = t.group;

        const parentRespPol =
          g.children.find((c) => c.respPol && String(c.respPol).trim())?.respPol || "";
        const parentRespOp =
          g.children.find((c) => c.respOp && String(c.respOp).trim())?.respOp || "";
        const parentStatus =
          g.children.find((c) => c.status && String(c.status).trim())?.status || "";

        const parentTask = {
          parentLabel: g.parentLabel,
          childLabel: g.parentLabel,
          startDate: g.aggStart,
          endDate: g.aggEnd,
          isMilestone: g.onlySingleMilestone,
          milestoneDate: g.onlySingleMilestone ? g.aggEnd : null,
          priority: null,
          status: parentStatus,
          respPol: parentRespPol,
          respOp: parentRespOp,
          respChild: "",
          selector: "",
          order: g.order
        };

        addBar(trackIndex, parentTask, { isParent: true, parentKey: g.parentKey });
      } else if (t.kind === "child") {
        addBar(trackIndex, t.task, { isParent: false });
      } else if (t.kind === "groupChildren") {
        t.group.children.forEach((c) => {
          addBar(trackIndex, c, { isParent: false, hideLabel: true });
        });
        addGroupLabelAtRightmostVisibleItem(trackIndex, t.group);
      } else if (t.kind === "leaf") {
        addBar(trackIndex, t.task, { isParent: false });
      }
      trackIndex++;
    }
  }

  function setBarCursor(bar, e) {
    const role = bar.dataset.role;
    if (role === "parent") {
      bar.style.cursor = "grab";
      return;
    }
    const rect = bar.getBoundingClientRect();
    const offsetX = e.clientX - rect.left;
    if (offsetX < 8 || rect.right - e.clientX < 8) {
      bar.style.cursor = "ew-resize";
    } else {
      bar.style.cursor = "grab";
    }
  }

  function attachBarDrag(bar) {
    bar.addEventListener("mousedown", (e) => {
      if (e.button !== 0) return;
      e.preventDefault();
      hideTooltip();

      const role = bar.dataset.role;
      const rect = bar.getBoundingClientRect();

      const totalDays = diffInDays(visibleStart, visibleEnd) + 1;
      const containerWidth =
        timelineBodyEl.clientWidth ||
        timelineHeaderEl.clientWidth ||
        rect.width;

      dragState.pxPerDay = containerWidth / Math.max(1, totalDays);
      dragState.active = true;
      dragState.bar = bar;
      dragState.milestone = null;
      dragState.startX = e.clientX;
      dragState.originalStart = normalizeDate(bar.dataset.start);
      dragState.originalEnd = normalizeDate(bar.dataset.end);

      if (role === "child") {
        dragState.taskId = parseInt(bar.dataset.rowId, 10);
        const offsetX = e.clientX - rect.left;
        if (offsetX < 8) dragState.type = "resize-left-child";
        else if (rect.right - e.clientX < 8) dragState.type = "resize-right-child";
        else dragState.type = "move-child";
      } else {
        dragState.type = "move-parent";
        dragState.parentKey = bar.dataset.parentKey;
        const grp = parentGroups.find((g) => g.parentKey === dragState.parentKey);
        dragState.originalChildren = grp
          ? grp.children.map((c) => ({
              rowId: c.rowId,
              startDate: c.startDate ? new Date(c.startDate.getTime()) : null,
              endDate: c.endDate ? new Date(c.endDate.getTime()) : null
            }))
          : [];
      }

      const midY = rect.top + rect.height / 2;
      const txtStart = formatDate(dragState.originalStart);
      const txtEnd = formatDate(dragState.originalEnd);
      const typeLabel =
        dragState.type === "move-child" || dragState.type === "move-parent"
          ? "déplacement"
          : dragState.type === "resize-left-child"
          ? "début"
          : "fin";

      showDragBubble(
        `${txtStart} → ${txtEnd}<span class="muted">${typeLabel}</span>`,
        e.clientX,
        midY
      );

      document.addEventListener("mousemove", onDragMove);
      document.addEventListener("mouseup", onDragEnd);
    });
  }

  function attachMilestoneDrag(m, task, isParent) {
    if (isParent) return;

    m.addEventListener("mousedown", (e) => {
      if (e.button !== 0) return;
      e.preventDefault();
      hideTooltip();

      const rect = m.getBoundingClientRect();
      const totalDays = diffInDays(visibleStart, visibleEnd) + 1;
      const containerWidth =
        timelineBodyEl.clientWidth ||
        timelineHeaderEl.clientWidth ||
        rect.width;

      dragState.pxPerDay = containerWidth / Math.max(1, totalDays);
      dragState.active = true;
      dragState.bar = null;
      dragState.milestone = m;
      dragState.type = "move-milestone";
      dragState.taskId = task.rowId;
      dragState.originalMilestoneDate = new Date(task.milestoneDate.getTime());
      dragState.startX = e.clientX;

      const midY = rect.top + rect.height / 2;
      showDragBubble(
        `${formatDate(task.milestoneDate)}<span class="muted">jalon</span>`,
        e.clientX,
        midY
      );

      document.addEventListener("mousemove", onDragMove);
      document.addEventListener("mouseup", onDragEnd);
    });
  }

  function onDragMove(e) {
    if (!dragState.active) return;
    e.preventDefault();

    const deltaX = e.clientX - dragState.startX;
    const deltaDays = Math.round(deltaX / dragState.pxPerDay);

    const totalDays = diffInDays(visibleStart, visibleEnd) + 1;
    const containerWidth = timelineBodyEl.clientWidth || timelineHeaderEl.clientWidth;
    if (!containerWidth || !totalDays) return;

    if (
      dragState.type === "move-child" ||
      dragState.type === "resize-left-child" ||
      dragState.type === "resize-right-child"
    ) {
      const origStart = dragState.originalStart;
      const origEnd = dragState.originalEnd;
      let newStart = new Date(origStart.getTime());
      let newEnd = new Date(origEnd.getTime());

      if (dragState.type === "move-child") {
        newStart = addDays(newStart, deltaDays);
        newEnd = addDays(newEnd, deltaDays);
      } else if (dragState.type === "resize-left-child") {
        newStart = addDays(newStart, deltaDays);
        if (newStart > newEnd) newStart = new Date(newEnd.getTime());
      } else if (dragState.type === "resize-right-child") {
        newEnd = addDays(newEnd, deltaDays);
        if (newEnd < newStart) newEnd = new Date(newStart.getTime());
      }

      const leftFrac = diffInDays(visibleStart, newStart) / totalDays;
      const rightFrac = diffInDays(visibleStart, newEnd) / totalDays;
      const widthFrac = Math.max(0.01, (rightFrac - leftFrac) + (1 / totalDays));

      dragState.bar.style.left = (leftFrac * containerWidth).toFixed(1) + "px";
      dragState.bar.style.width = (widthFrac * containerWidth).toFixed(1) + "px";
      dragState.bar.dataset.start = newStart.toISOString();
      dragState.bar.dataset.end = newEnd.toISOString();

      const rect = dragState.bar.getBoundingClientRect();
      const midY = rect.top + rect.height / 2;
      const label =
        dragState.type === "move-child"
          ? "déplacement"
          : dragState.type === "resize-left-child"
          ? "début"
          : "fin";

      showDragBubble(
        `${formatDate(newStart)} → ${formatDate(newEnd)}<span class="muted">${label}</span>`,
        e.clientX,
        midY
      );
    }

    if (dragState.type === "move-parent") {
      const grp = parentGroups.find((g) => g.parentKey === dragState.parentKey);
      if (!grp || !grp.aggStart || !grp.aggEnd) return;

      const newStart = addDays(grp.aggStart, deltaDays);
      const newEnd = addDays(grp.aggEnd, deltaDays);

      const leftFrac = diffInDays(visibleStart, newStart) / totalDays;
      const rightFrac = diffInDays(visibleStart, newEnd) / totalDays;
      const widthFrac = Math.max(0.01, (rightFrac - leftFrac) + (1 / totalDays));

      dragState.bar.style.left = (leftFrac * containerWidth).toFixed(1) + "px";
      dragState.bar.style.width = (widthFrac * containerWidth).toFixed(1) + "px";

      const rect = dragState.bar.getBoundingClientRect();
      const midY = rect.top + rect.height / 2;
      showDragBubble(
        `${formatDate(newStart)} → ${formatDate(newEnd)}<span class="muted">groupe</span>`,
        e.clientX,
        midY
      );
    }

    if (dragState.type === "move-milestone") {
      const orig = dragState.originalMilestoneDate;
      const newDate = addDays(orig, deltaDays);
      const frac = diffInDays(visibleStart, newDate) / totalDays;
      const x = frac * containerWidth;
      dragState.milestone.style.left = x.toFixed(1) + "px";

      const rect = dragState.milestone.getBoundingClientRect();
      const midY = rect.top + rect.height / 2;
      showDragBubble(
        `${formatDate(newDate)}<span class="muted">jalon</span>`,
        e.clientX,
        midY
      );
    }
  }

  async function onDragEnd(e) {
    if (!dragState.active) return;
    e.preventDefault();

    document.removeEventListener("mousemove", onDragMove);
    document.removeEventListener("mouseup", onDragEnd);
    hideDragBubble();

    const deltaX = e.clientX - dragState.startX;
    const deltaDays = Math.round(deltaX / dragState.pxPerDay);

    try {
      if (
        dragState.type === "move-child" ||
        dragState.type === "resize-left-child" ||
        dragState.type === "resize-right-child"
      ) {
        const newStart = normalizeDate(dragState.bar.dataset.start);
        const newEnd = normalizeDate(dragState.bar.dataset.end);
        await updateChildDates(dragState.taskId, newStart, newEnd);
        showToast("Dates mises à jour", "success");
      } else if (dragState.type === "move-parent" && dragState.originalChildren) {
        await moveParentGroup(dragState.parentKey, dragState.originalChildren, deltaDays);
        showToast("Groupe déplacé", "success");
      } else if (dragState.type === "move-milestone") {
        const newDate = addDays(dragState.originalMilestoneDate, deltaDays);
        await updateMilestoneDate(dragState.taskId, newDate);
        showToast("Jalon déplacé", "success");
      }
    } catch (err) {
      console.error(err);
      showToast("Erreur lors de la mise à jour", "error");
    } finally {
      dragState.active = false;
      dragState.bar = null;
      dragState.milestone = null;
      dragState.type = null;
      dragState.taskId = null;
      dragState.parentKey = null;
      dragState.originalChildren = null;
    }
  }

  function render() {
    if (!allRecords.length) {
      taskListEl.innerHTML = '<div class="empty">En attente de données…</div>';
      timelineGridEl.innerHTML = "";
      yearsRowEl.innerHTML = "";
      monthsRowEl.innerHTML = "";
      weeksRowEl.innerHTML = "";
      daysRowEl.innerHTML = "";
      currentPeriodEl.textContent = "–";
      taskCountEl.textContent = "";
      return;
    }

    initColorFieldSelect();
    buildHeaders();
    renderTaskList();
    renderTimeline();
  }

  window.addEventListener("resize", () => {
    if (!visibleStart || !visibleEnd || !allRecords.length) return;
    render();
  });

  grist.ready({
    requiredAccess: "full",
    columns: [
      { name: "parent", title: "Élément parent", optional: true },
      { name: "child", title: "Élément enfant", optional: true },
      { name: "start", title: "Date de début", optional: true, type: "Date,DateTime" },
      { name: "end", title: "Date de fin", optional: true, type: "Date,DateTime" },
      { name: "priority", title: "Priorité", optional: true },
      { name: "status", title: "Statut", optional: true },
      { name: "respPol", title: "Référent politique", optional: true },
      { name: "respOp", title: "Référent opérationnel", optional: true },
      { name: "respChild", title: "Responsable enfant", optional: true },
      { name: "selector", title: "Sélecteur O/N", optional: true },
      { name: "order", title: "Ordre d’affichage", optional: true }
    ]
  });

  grist.onRecords(async function (records, mappings) {
    latestMappings = mappings || null;

    if (!records || !records.length) {
      allRecords = [];
      parentGroups = [];
      leafTasks = [];
      globalMinDate = null;
      globalMaxDate = null;
      visibleStart = null;
      visibleEnd = null;
      currentMappingsOk = false;
      await refreshTableInfo();
      render();
      return;
    }

    try {
      const probe = grist.mapColumnNames(records[0], { mappings: latestMappings });
      currentMappingsOk = !!probe;
    } catch (e) {
      currentMappingsOk = false;
    }

    await refreshTableInfo();

    allRecords = buildLogicalRecords(records);
    groupData(allRecords);

    const range = computeGlobalRange(allRecords);
    globalMinDate = range.min;
    globalMaxDate = range.max;
    setVisibleRangeForZoom(false);

    updateZoomButtons();
    toggleLabelsBtn.textContent = labelsVisible ? "Masquer labels" : "Afficher labels";
    groupChildrenBtn.textContent = childrenOnOneRow
      ? "Enfants : 1 ligne"
      : "Enfants : multi-lignes";

    render();
  });
})();
