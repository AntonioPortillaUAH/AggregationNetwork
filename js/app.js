/* Interfaz de la aplicación: carga del Excel, ejecución y visualización. */
(function () {
  "use strict";

  const A = window.Aggregation;
  const $ = function (id) { return document.getElementById(id); };

  const els = {
    fileInput: $("fileInput"),
    exampleBtn: $("exampleBtn"),
    fileInfo: $("fileInfo"),
    numControllers: $("numControllers"),
    capacity: $("capacity"),
    numRounds: $("numRounds"),
    seed: $("seed"),
    capacityInfo: $("capacityInfo"),
    runBtn: $("runBtn"),
    stepCard: $("stepCard"),
    stepHint: $("stepHint"),
    stepBack: $("stepBack"),
    stepFwd: $("stepFwd"),
    stepAll: $("stepAll"),
    playBtn: $("playBtn"),
    stepSlider: $("stepSlider"),
    stepInfo: $("stepInfo"),
    tabsWrap: $("tabsWrap"),
    tabButtons: Array.prototype.slice.call(document.querySelectorAll("[data-tab]")),
    methodInfo: $("methodInfo"),
    mapCard: $("mapCard"),
    canvas: $("map"),
    tooltip: $("tooltip"),
    legend: $("legend"),
    results: $("results"),
    stats: $("stats"),
    controllerTable: $("controllerTable"),
    assignmentTable: $("assignmentTable"),
    exportBtn: $("exportBtn"),
    exportBtn2: $("exportBtn2"),
    compare: $("compare"),
    compareTable: $("compareTable"),
    compareNote: $("compareNote"),
    compareCtrlTable: $("compareCtrlTable"),
  };

  // Métodos de asignación. Los dos usan los mismos controladores para que la
  // comparación sea justa.
  const METHODS = {
    greedy: { label: "Greedy clásico", short: "Greedy" },
    rounds: { label: "Greedy por rondas", short: "Rondas" },
  };

  const state = {
    bts: [],
    fileName: "",
    run: null,      // { controllers, capacity, rounds, seed, colors, results: {greedy, rounds} }
    tab: "greedy",  // "greedy" | "rounds" | "compare"
    shown: { greedy: 0, rounds: 0 }, // pasos visibles en el mapa de cada método
    timer: null,
    view: null,     // transformación mundo -> pantalla del último dibujo
  };

  // ---------------------------------------------------------------- utilidades

  const BASE_COLORS = [
    "#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd", "#8c564b",
    "#e377c2", "#17becf", "#bcbd22", "#393b79", "#ad494a", "#637939",
  ];

  function controllerColors(n) {
    const colors = [];
    for (let i = 0; i < n; i++) {
      colors.push(i < BASE_COLORS.length
        ? BASE_COLORS[i]
        : "hsl(" + Math.round((i * 137.508) % 360) + ", 65%, 48%)");
    }
    return colors;
  }

  function fmtMeters(m) {
    if (m === null || m === undefined) return "—";
    return Math.abs(m) >= 1000
      ? (m / 1000).toLocaleString("es-ES", { maximumFractionDigits: 2 }) + " km"
      : m.toLocaleString("es-ES", { maximumFractionDigits: 1 }) + " m";
  }

  function fmtNum(v, digits) {
    return v.toLocaleString("es-ES", { maximumFractionDigits: digits === undefined ? 1 : digits });
  }

  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  }

  function setInfo(el, kind, html) {
    el.className = "info" + (kind ? " " + kind : "");
    el.innerHTML = html;
  }

  function readInt(input) {
    const v = Number(input.value);
    return Number.isInteger(v) ? v : NaN;
  }

  function cssVar(name) {
    return getComputedStyle(document.documentElement).getPropertyValue(name).trim();
  }

  function swatch(color) {
    return "<span class=\"swatch\" style=\"background:" + color + "\"></span>";
  }

  // Resultado del método de la pestaña activa (null en la pestaña de comparación).
  function current() {
    return state.run && METHODS[state.tab] ? state.run.results[state.tab] : null;
  }

  // ------------------------------------------------------------ carga de datos

  function loadBts(bts, sourceName, errors) {
    stopPlay();
    state.bts = bts;
    state.fileName = sourceName;
    state.run = null;
    state.shown = { greedy: 0, rounds: 0 };
    els.tabsWrap.hidden = true;
    els.results.hidden = true;
    els.compare.hidden = true;
    els.mapCard.hidden = false;
    els.stepCard.hidden = true;
    els.legend.innerHTML = "";

    if (bts.length === 0) {
      setInfo(els.fileInfo, "error", "No se ha encontrado ninguna BTS válida en <b>" + escapeHtml(sourceName) + "</b>.");
    } else {
      let html = "<b>" + bts.length + "</b> BTS cargadas de <b>" + escapeHtml(sourceName) + "</b>.";
      if (errors && errors.length) {
        const shown = errors.slice(0, 5).map(function (e) { return "<li>" + escapeHtml(e) + "</li>"; }).join("");
        html += " Se han ignorado " + errors.length + " filas:<ul>" + shown +
          (errors.length > 5 ? "<li>…</li>" : "") + "</ul>";
      }
      setInfo(els.fileInfo, errors && errors.length ? "warn" : "ok", html);
    }
    updateCapacityInfo();
    draw();
  }

  function onFile(file) {
    const reader = new FileReader();
    reader.onload = function (e) {
      try {
        const wb = XLSX.read(new Uint8Array(e.target.result), { type: "array" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, blankrows: false });
        const parsed = A.parseBtsRows(rows);
        loadBts(parsed.bts, file.name, parsed.errors);
      } catch (err) {
        loadBts([], file.name, []);
        setInfo(els.fileInfo, "error", "No se ha podido leer el fichero: " + escapeHtml(err.message));
      }
    };
    reader.readAsArrayBuffer(file);
  }

  // Datos de ejemplo: BTS agrupadas alrededor de varios núcleos urbanos en un área de 20 x 15 km.
  function exampleBts() {
    const rng = A.createRng("ejemplo");
    const towns = [
      { x: 4000, y: 4000, r: 1800, n: 22 },
      { x: 14000, y: 5000, r: 2200, n: 26 },
      { x: 9000, y: 11000, r: 1500, n: 16 },
      { x: 17000, y: 12500, r: 1200, n: 10 },
    ];
    const bts = [];
    towns.forEach(function (t) {
      for (let i = 0; i < t.n; i++) {
        const ang = rng() * 2 * Math.PI;
        const rad = t.r * Math.sqrt(rng());
        bts.push({ x: t.x + rad * Math.cos(ang), y: t.y + rad * Math.sin(ang) });
      }
    });
    for (let i = 0; i < 12; i++) bts.push({ x: rng() * 20000, y: rng() * 15000 });
    return bts.map(function (b, i) {
      return { id: "BTS" + String(i + 1).padStart(3, "0"), x: Math.round(b.x), y: Math.round(b.y) };
    });
  }

  function updateCapacityInfo() {
    const n = readInt(els.numControllers);
    const cap = readInt(els.capacity);
    const rounds = readInt(els.numRounds);
    const total = state.bts.length;
    let problem = "";

    if (!(n >= 1)) problem = "El número de controladores debe ser un entero ≥ 1.";
    else if (!(cap >= 1)) problem = "El máximo de BTS por controlador debe ser un entero ≥ 1.";
    else if (!(rounds >= 1)) problem = "El número de rondas debe ser un entero ≥ 1.";
    else if (rounds > cap) problem = "El número de rondas (" + rounds + ") no puede superar el máximo de BTS por controlador (" + cap + ").";
    else if (total && n > total) problem = "Hay más controladores (" + n + ") que BTS (" + total + ").";

    els.runBtn.disabled = !!problem || total === 0;

    if (problem) {
      setInfo(els.capacityInfo, "error", problem);
      return;
    }
    if (total === 0) {
      setInfo(els.capacityInfo, "", "");
      return;
    }

    let html;
    let kind = "ok";
    if (n * cap < total) {
      kind = "warn";
      html = "Capacidad total " + n + " × " + cap + " = <b>" + n * cap + "</b> &lt; " + total +
        " BTS: quedarán <b>" + (total - n * cap) + "</b> BTS sin asignar.";
    } else {
      html = "Capacidad total " + n + " × " + cap + " = <b>" + n * cap + "</b> ≥ " + total +
        " BTS (ocupación media " + fmtNum((100 * total) / (n * cap), 0) + " %).";
    }
    const limits = A.roundLimits(cap, rounds);
    html += "<br>Rondas: hasta " + fmtNum(cap / rounds, 1) + " BTS por controlador y ronda " +
      "(límites acumulados: " + (limits.length <= 10 ? limits.join(", ") : limits.slice(0, 9).join(", ") + ", …, " + cap) + ").";
    setInfo(els.capacityInfo, kind, html);
  }

  // ------------------------------------------------------------------ ejecución

  function run() {
    stopPlay();
    const n = readInt(els.numControllers);
    const cap = readInt(els.capacity);
    const rounds = readInt(els.numRounds);
    try {
      const rng = A.createRng(els.seed.value);
      const controllers = A.selectControllers(state.bts.length, n, rng);
      state.run = {
        controllers: controllers,
        capacity: cap,
        rounds: rounds,
        seed: els.seed.value,
        colors: controllerColors(n),
        results: {
          greedy: A.greedyAssign(state.bts, controllers, cap),
          rounds: A.greedyRoundsAssign(state.bts, controllers, cap, rounds),
        },
      };
      state.shown = {
        greedy: state.run.results.greedy.steps.length,
        rounds: state.run.results.rounds.steps.length,
      };
    } catch (err) {
      setInfo(els.capacityInfo, "error", escapeHtml(err.message));
      return;
    }
    els.tabsWrap.hidden = false;
    renderCompare();
    showTab(state.tab);
  }

  // ----------------------------------------------------------------- pestañas

  function showTab(tab) {
    stopPlay();
    state.tab = tab;
    els.tabButtons.forEach(function (b) {
      b.setAttribute("aria-selected", b.dataset.tab === tab ? "true" : "false");
    });
    const run = state.run;
    const isMethod = !!METHODS[tab];
    els.mapCard.hidden = !isMethod;
    els.results.hidden = !isMethod;
    els.stepCard.hidden = !isMethod;
    els.compare.hidden = isMethod;

    if (tab === "greedy") {
      els.methodInfo.textContent = "Recorre todas las parejas (BTS, controlador) de menor a mayor distancia y asigna " +
        "la BTS si aún está libre y el controlador no está lleno (máx. " + run.capacity + " BTS).";
      els.stepHint.textContent = "Recorre las asignaciones en el orden en que las realiza el algoritmo (de menor a mayor distancia).";
    } else if (tab === "rounds") {
      els.methodInfo.textContent = "Igual que el Greedy clásico, pero en " + run.rounds + " rondas: en cada ronda cada " +
        "controlador solo puede recibir unas " + fmtNum(run.capacity / run.rounds, 1) + " BTS más (límites acumulados " +
        run.results.rounds.limits.join(", ") + ").";
      els.stepHint.textContent = "Recorre las asignaciones ronda a ronda; dentro de cada ronda, de menor a mayor distancia.";
    } else {
      els.methodInfo.textContent = "Ambos métodos usan los mismos " + run.controllers.length +
        " controladores elegidos al azar, así que las diferencias se deben solo al algoritmo.";
    }

    if (isMethod) {
      const res = current();
      els.stepSlider.max = res.steps.length;
      renderLegend(res);
      renderStats(res);
      renderTables(res, tab);
      updateStep();
    }
  }

  // ------------------------------------------------------------- paso a paso

  function setShown(k) {
    const res = current();
    if (!res) return;
    state.shown[state.tab] = Math.max(0, Math.min(res.steps.length, k));
    updateStep();
  }

  function updateStep() {
    const res = current();
    if (!res) return;
    const steps = res.steps;
    const shown = state.shown[state.tab];
    els.stepSlider.value = shown;
    if (shown === 0) {
      setInfo(els.stepInfo, "", "Paso 0 de " + steps.length + ": controladores situados, ninguna BTS asignada.");
    } else {
      const s = steps[shown - 1];
      setInfo(els.stepInfo, "",
        (state.tab === "rounds" ? "Ronda <b>" + s.round + "</b> de " + state.run.rounds + " · " : "") +
        "Paso <b>" + shown + "</b> de " + steps.length + ": " +
        escapeHtml(state.bts[s.bts].id) + " → <b>C" + (s.controller + 1) + "</b> (" + fmtMeters(s.distance) + ")");
    }
    draw();
  }

  function stopPlay() {
    if (state.timer) clearInterval(state.timer);
    state.timer = null;
    els.playBtn.textContent = "▶ Reproducir";
  }

  function togglePlay() {
    if (state.timer) { stopPlay(); return; }
    const res = current();
    if (!res) return;
    const total = res.steps.length;
    if (state.shown[state.tab] >= total) setShown(0);
    const delay = Math.max(30, Math.min(400, 10000 / Math.max(1, total)));
    els.playBtn.textContent = "❚❚ Pausa";
    state.timer = setInterval(function () {
      if (state.shown[state.tab] >= total) { stopPlay(); return; }
      setShown(state.shown[state.tab] + 1);
    }, delay);
  }

  // ------------------------------------------------------------------- dibujo

  function niceStep(range, target) {
    const raw = range / target;
    const mag = Math.pow(10, Math.floor(Math.log10(raw)));
    const norm = raw / mag;
    return (norm < 1.5 ? 1 : norm < 3 ? 2 : norm < 7 ? 5 : 10) * mag;
  }

  function draw() {
    const canvas = els.canvas;
    if (els.mapCard.hidden) return;
    const dpr = window.devicePixelRatio || 1;
    const w = canvas.clientWidth;
    const h = canvas.clientHeight;
    canvas.width = Math.round(w * dpr);
    canvas.height = Math.round(h * dpr);
    const ctx = canvas.getContext("2d");
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.clearRect(0, 0, w, h);

    const text = cssVar("--text");
    const muted = cssVar("--muted");

    if (state.bts.length === 0) {
      state.view = null;
      ctx.fillStyle = muted;
      ctx.font = "15px system-ui, sans-serif";
      ctx.textAlign = "center";
      ctx.fillText("Carga un fichero Excel con las estaciones base para empezar", w / 2, h / 2);
      return;
    }

    // Transformación con la misma escala en X e Y (coordenadas métricas).
    let minX = Infinity, maxX = -Infinity, minY = Infinity, maxY = -Infinity;
    state.bts.forEach(function (b) {
      minX = Math.min(minX, b.x); maxX = Math.max(maxX, b.x);
      minY = Math.min(minY, b.y); maxY = Math.max(maxY, b.y);
    });
    const spanX = Math.max(maxX - minX, 1);
    const spanY = Math.max(maxY - minY, 1);
    const margin = { l: 64, r: 20, t: 20, b: 40 };
    const plotW = w - margin.l - margin.r;
    const plotH = h - margin.t - margin.b;
    const scale = Math.min(plotW / (spanX * 1.06), plotH / (spanY * 1.06));
    const cx = (minX + maxX) / 2;
    const cy = (minY + maxY) / 2;
    const ox = margin.l + plotW / 2;
    const oy = margin.t + plotH / 2;
    const sx = function (x) { return ox + (x - cx) * scale; };
    const sy = function (y) { return oy - (y - cy) * scale; };
    state.view = { sx: sx, sy: sy };

    // Rejilla y ejes.
    const useKm = Math.max(spanX, spanY) >= 5000;
    const x0 = cx - plotW / 2 / scale, x1 = cx + plotW / 2 / scale;
    const y0 = cy - plotH / 2 / scale, y1 = cy + plotH / 2 / scale;
    const step = niceStep(Math.max(x1 - x0, y1 - y0), 8);
    const label = function (v) {
      if (Math.abs(v) < step * 1e-6) v = 0; // evita "-0" por redondeo
      return useKm ? fmtNum(v / 1000, 2) : fmtNum(v, 0);
    };

    ctx.lineWidth = 1;
    ctx.font = "11px system-ui, sans-serif";
    ctx.fillStyle = muted;
    ctx.strokeStyle = cssVar("--grid");
    ctx.textAlign = "center";
    ctx.textBaseline = "top";
    for (let x = Math.ceil(x0 / step) * step; x <= x1; x += step) {
      ctx.beginPath(); ctx.moveTo(sx(x), margin.t); ctx.lineTo(sx(x), h - margin.b); ctx.stroke();
      ctx.fillText(label(x), sx(x), h - margin.b + 6);
    }
    ctx.textAlign = "right";
    ctx.textBaseline = "middle";
    for (let y = Math.ceil(y0 / step) * step; y <= y1; y += step) {
      ctx.beginPath(); ctx.moveTo(margin.l, sy(y)); ctx.lineTo(w - margin.r, sy(y)); ctx.stroke();
      ctx.fillText(label(y), margin.l - 6, sy(y));
    }
    ctx.strokeStyle = cssVar("--axis");
    ctx.strokeRect(margin.l, margin.t, plotW, plotH);
    ctx.textAlign = "center";
    ctx.textBaseline = "bottom";
    ctx.fillText("X (" + (useKm ? "km" : "m") + ")", margin.l + plotW / 2, h - 4);
    ctx.save();
    ctx.translate(14, margin.t + plotH / 2);
    ctx.rotate(-Math.PI / 2);
    ctx.textBaseline = "middle";
    ctx.fillText("Y (" + (useKm ? "km" : "m") + ")", 0, 0);
    ctx.restore();

    ctx.save();
    ctx.beginPath();
    ctx.rect(margin.l, margin.t, plotW, plotH);
    ctx.clip();

    const run = state.run;
    const res = current();
    const shown = res ? state.shown[state.tab] : 0;
    const owner = new Array(state.bts.length).fill(-1);
    if (res) {
      const steps = res.steps;
      // Enlaces BTS -> controlador de los pasos visibles.
      ctx.globalAlpha = 0.55;
      ctx.lineWidth = 1.2;
      for (let k = 0; k < shown; k++) {
        const s = steps[k];
        owner[s.bts] = s.controller;
        const b = state.bts[s.bts];
        const c = state.bts[run.controllers[s.controller]];
        ctx.strokeStyle = run.colors[s.controller];
        ctx.beginPath(); ctx.moveTo(sx(b.x), sy(b.y)); ctx.lineTo(sx(c.x), sy(c.y)); ctx.stroke();
      }
      ctx.globalAlpha = 1;
      // Última asignación resaltada.
      if (shown > 0 && shown < steps.length) {
        const s = steps[shown - 1];
        const b = state.bts[s.bts];
        const c = state.bts[run.controllers[s.controller]];
        ctx.strokeStyle = run.colors[s.controller];
        ctx.lineWidth = 3.5;
        ctx.beginPath(); ctx.moveTo(sx(b.x), sy(b.y)); ctx.lineTo(sx(c.x), sy(c.y)); ctx.stroke();
        ctx.lineWidth = 2;
        ctx.strokeStyle = text;
        ctx.beginPath(); ctx.arc(sx(b.x), sy(b.y), 8, 0, 2 * Math.PI); ctx.stroke();
      }
    }

    // BTS.
    const finished = res && shown === res.steps.length;
    state.bts.forEach(function (b, i) {
      const x = sx(b.x), y = sy(b.y);
      if (finished && owner[i] === -1) {
        ctx.strokeStyle = cssVar("--error-text");
        ctx.lineWidth = 2;
        ctx.beginPath();
        ctx.moveTo(x - 4, y - 4); ctx.lineTo(x + 4, y + 4);
        ctx.moveTo(x + 4, y - 4); ctx.lineTo(x - 4, y + 4);
        ctx.stroke();
        return;
      }
      ctx.fillStyle = owner[i] === -1 ? muted : run.colors[owner[i]];
      ctx.beginPath(); ctx.arc(x, y, 3.5, 0, 2 * Math.PI); ctx.fill();
    });

    // Controladores.
    if (run) {
      ctx.font = "bold 12px system-ui, sans-serif";
      ctx.textAlign = "left";
      ctx.textBaseline = "bottom";
      run.controllers.forEach(function (idx, c) {
        const b = state.bts[idx];
        const x = sx(b.x), y = sy(b.y);
        ctx.fillStyle = run.colors[c];
        ctx.strokeStyle = cssVar("--surface");
        ctx.lineWidth = 2;
        ctx.beginPath();
        ctx.rect(x - 7, y - 7, 14, 14);
        ctx.fill(); ctx.stroke();
        ctx.fillStyle = text;
        ctx.fillText("C" + (c + 1), x + 9, y - 5);
      });
    }
    ctx.restore();
  }

  function onHover(ev) {
    if (!state.view) return;
    const rect = els.canvas.getBoundingClientRect();
    const mx = ev.clientX - rect.left;
    const my = ev.clientY - rect.top;
    let best = -1, bestD = 10;
    state.bts.forEach(function (b, i) {
      const d = Math.hypot(state.view.sx(b.x) - mx, state.view.sy(b.y) - my);
      if (d < bestD) { bestD = d; best = i; }
    });
    if (best === -1) { els.tooltip.hidden = true; return; }

    const b = state.bts[best];
    let html = "<b>" + escapeHtml(b.id) + "</b><br>X = " + fmtNum(b.x) + " m, Y = " + fmtNum(b.y) + " m";
    const res = current();
    if (res) {
      const c = state.run.controllers.indexOf(best);
      if (c !== -1) html += "<br>Emplazamiento del controlador <b>C" + (c + 1) + "</b>";
      const a = res.assignment[best];
      html += a === -1
        ? "<br>Sin asignar"
        : "<br>Asignada a <b>C" + (a + 1) + "</b> (" + fmtMeters(res.distances[best]) + ")";
    }
    els.tooltip.innerHTML = html;
    els.tooltip.hidden = false;
    const tw = els.tooltip.offsetWidth;
    const left = mx + 14 + tw > rect.width ? mx - 14 - tw : mx + 14;
    els.tooltip.style.left = left + 8 + "px";
    els.tooltip.style.top = my + 8 + "px";
  }

  // --------------------------------------------------------------- resultados

  function controllerSummary(res) {
    const run = state.run;
    const rows = run.controllers.map(function (idx, c) {
      return { c: c, site: state.bts[idx], load: res.loads[c], total: 0, max: 0 };
    });
    res.assignment.forEach(function (c, i) {
      if (c === -1) return;
      const d = res.distances[i];
      rows[c].total += d;
      rows[c].max = Math.max(rows[c].max, d);
    });
    return rows;
  }

  // Indicadores globales de un resultado.
  function metrics(res) {
    const assigned = state.bts.length - res.unassigned.length;
    let maxD = 0, sq = 0;
    const mean = assigned ? res.totalDistance / assigned : 0;
    res.distances.forEach(function (d) {
      if (d === null) return;
      maxD = Math.max(maxD, d);
      sq += (d - mean) * (d - mean);
    });
    return {
      assigned: assigned,
      unassigned: res.unassigned.length,
      total: res.totalDistance,
      mean: mean,
      max: maxD,
      std: assigned ? Math.sqrt(sq / assigned) : 0,
      minLoad: Math.min.apply(null, res.loads),
      maxLoad: Math.max.apply(null, res.loads),
    };
  }

  function renderLegend(res) {
    const run = state.run;
    let html = run.colors.map(function (col, c) {
      return "<span><i style=\"background:" + col + "\"></i>C" + (c + 1) + "</span>";
    }).join("");
    html += "<span>■ controlador · ● BTS" + (res.unassigned.length ? " · ✕ sin asignar" : "") + "</span>";
    els.legend.innerHTML = html;
  }

  function renderStats(res) {
    const m = metrics(res);
    const items = [
      ["BTS asignadas", m.assigned + " / " + state.bts.length],
      ["Sin asignar", String(m.unassigned)],
      ["Distancia total", fmtMeters(m.total)],
      ["Distancia media", fmtMeters(m.mean)],
      ["Distancia máxima", fmtMeters(m.max)],
    ];
    els.stats.innerHTML = items.map(function (it) {
      return "<div class=\"stat\"><div class=\"label\">" + it[0] + "</div><div class=\"value\">" + it[1] + "</div></div>";
    }).join("");
  }

  function renderTables(res, method) {
    const run = state.run;
    const cap = run.capacity;

    let html = "<thead><tr><th>Controlador</th><th>Emplazamiento</th><th class=\"num\">X (m)</th><th class=\"num\">Y (m)</th>" +
      "<th class=\"num\">BTS</th><th>Ocupación</th><th class=\"num\">Dist. total</th><th class=\"num\">Dist. media</th><th class=\"num\">Dist. máx.</th></tr></thead><tbody>";
    controllerSummary(res).forEach(function (r) {
      const col = run.colors[r.c];
      html += "<tr><td>" + swatch(col) + "C" + (r.c + 1) + "</td>" +
        "<td>" + escapeHtml(r.site.id) + "</td>" +
        "<td class=\"num\">" + fmtNum(r.site.x) + "</td><td class=\"num\">" + fmtNum(r.site.y) + "</td>" +
        "<td class=\"num\">" + r.load + " / " + cap + "</td>" +
        "<td><div class=\"bar\"><div style=\"width:" + (100 * r.load / cap) + "%;background:" + col + "\"></div></div></td>" +
        "<td class=\"num\">" + fmtMeters(r.total) + "</td>" +
        "<td class=\"num\">" + fmtMeters(r.load ? r.total / r.load : 0) + "</td>" +
        "<td class=\"num\">" + fmtMeters(r.max) + "</td></tr>";
    });
    els.controllerTable.innerHTML = html + "</tbody>";

    const withRound = method === "rounds";
    const order = new Array(state.bts.length).fill(null);
    const roundOf = new Array(state.bts.length).fill(null);
    res.steps.forEach(function (s, k) { order[s.bts] = k + 1; roundOf[s.bts] = s.round; });
    html = "<thead><tr><th>Id BTS</th><th class=\"num\">X (m)</th><th class=\"num\">Y (m)</th><th>Controlador</th>" +
      "<th class=\"num\">Distancia</th>" + (withRound ? "<th class=\"num\">Ronda</th>" : "") +
      "<th class=\"num\">Orden</th></tr></thead><tbody>";
    state.bts.forEach(function (b, i) {
      const c = res.assignment[i];
      html += "<tr" + (c === -1 ? " class=\"unassigned\"" : "") + "><td>" + escapeHtml(b.id) + "</td>" +
        "<td class=\"num\">" + fmtNum(b.x) + "</td><td class=\"num\">" + fmtNum(b.y) + "</td>" +
        "<td>" + (c === -1 ? "Sin asignar" : swatch(run.colors[c]) + "C" + (c + 1)) + "</td>" +
        "<td class=\"num\">" + fmtMeters(res.distances[i]) + "</td>" +
        (withRound ? "<td class=\"num\">" + (roundOf[i] === null ? "—" : roundOf[i]) + "</td>" : "") +
        "<td class=\"num\">" + (order[i] === null ? "—" : order[i]) + "</td></tr>";
    });
    els.assignmentTable.innerHTML = html + "</tbody>";
  }

  function renderCompare() {
    const run = state.run;
    const g = run.results.greedy;
    const r = run.results.rounds;
    const mg = metrics(g);
    const mr = metrics(r);

    // [etiqueta, valor greedy, valor rondas, formato, true si "menos es mejor" / false si "más es mejor" / null]
    const rows = [
      ["BTS asignadas", mg.assigned, mr.assigned, String, false],
      ["BTS sin asignar", mg.unassigned, mr.unassigned, String, true],
      ["Distancia total", mg.total, mr.total, fmtMeters, true],
      ["Distancia media", mg.mean, mr.mean, fmtMeters, true],
      ["Distancia máxima", mg.max, mr.max, fmtMeters, true],
      ["Desviación típica de la distancia", mg.std, mr.std, fmtMeters, true],
      ["BTS por controlador (mín. – máx.)", null, null, null, null],
    ];

    let html = "<thead><tr><th>Indicador</th><th class=\"num\">" + METHODS.greedy.label + "</th><th class=\"num\">" +
      METHODS.rounds.label + " (" + run.rounds + " rondas)</th><th class=\"num\">Diferencia (rondas − clásico)</th></tr></thead><tbody>";
    rows.forEach(function (row) {
      if (row[3] === null) {
        html += "<tr><td>" + row[0] + "</td><td class=\"num\">" + mg.minLoad + " – " + mg.maxLoad + "</td>" +
          "<td class=\"num\">" + mr.minLoad + " – " + mr.maxLoad + "</td><td class=\"num\">—</td></tr>";
        return;
      }
      const a = row[1], b = row[2], fmt = row[3], lessIsBetter = row[4];
      const diff = b - a;
      const eps = 1e-9;
      let clsA = "", clsB = "";
      if (Math.abs(diff) > eps) {
        const roundsBetter = lessIsBetter ? diff < 0 : diff > 0;
        clsA = roundsBetter ? "" : " better";
        clsB = roundsBetter ? " better" : "";
      }
      const diffText = Math.abs(diff) <= eps ? "=" : (diff > 0 ? "+" : "−") + fmt(Math.abs(diff));
      html += "<tr><td>" + row[0] + "</td><td class=\"num" + clsA + "\">" + fmt(a) + "</td>" +
        "<td class=\"num" + clsB + "\">" + fmt(b) + "</td><td class=\"num\">" + diffText + "</td></tr>";
    });
    els.compareTable.innerHTML = html + "</tbody>";

    let changed = 0;
    g.assignment.forEach(function (c, i) { if (c !== r.assignment[i]) changed++; });
    els.compareNote.textContent = changed === 0
      ? "Los dos métodos producen exactamente la misma asignación."
      : changed + " de " + state.bts.length + " BTS (" + fmtNum((100 * changed) / state.bts.length, 0) +
        " %) tienen un controlador distinto en cada método. En verde, el método con mejor valor en cada indicador.";

    const sg = controllerSummary(g);
    const sr = controllerSummary(r);
    html = "<thead><tr><th>Controlador</th><th>Emplazamiento</th>" +
      "<th class=\"num\">BTS clásico</th><th class=\"num\">BTS rondas</th>" +
      "<th class=\"num\">Dist. media clásico</th><th class=\"num\">Dist. media rondas</th>" +
      "<th class=\"num\">Dist. máx. clásico</th><th class=\"num\">Dist. máx. rondas</th></tr></thead><tbody>";
    sg.forEach(function (a, c) {
      const b = sr[c];
      html += "<tr><td>" + swatch(run.colors[c]) + "C" + (c + 1) + "</td><td>" + escapeHtml(a.site.id) + "</td>" +
        "<td class=\"num\">" + a.load + "</td><td class=\"num\">" + b.load + "</td>" +
        "<td class=\"num\">" + fmtMeters(a.load ? a.total / a.load : 0) + "</td>" +
        "<td class=\"num\">" + fmtMeters(b.load ? b.total / b.load : 0) + "</td>" +
        "<td class=\"num\">" + fmtMeters(a.max) + "</td><td class=\"num\">" + fmtMeters(b.max) + "</td></tr>";
    });
    els.compareCtrlTable.innerHTML = html + "</tbody>";
  }

  function exportResults() {
    const run = state.run;
    const round2 = function (v) { return v === null ? null : Math.round(v * 100) / 100; };
    const g = run.results.greedy;
    const r = run.results.rounds;
    const info = {};
    ["greedy", "rounds"].forEach(function (m) {
      const order = new Array(state.bts.length).fill(null);
      const rnd = new Array(state.bts.length).fill(null);
      run.results[m].steps.forEach(function (s, k) { order[s.bts] = k + 1; rnd[s.bts] = s.round; });
      info[m] = { order: order, round: rnd };
    });
    const ctrlName = function (c) { return c === -1 ? "Sin asignar" : "C" + (c + 1); };

    const assignRows = state.bts.map(function (b, i) {
      return {
        "Id BTS": b.id, "X (m)": b.x, "Y (m)": b.y,
        "Controlador (clásico)": ctrlName(g.assignment[i]),
        "Distancia clásico (m)": round2(g.distances[i]),
        "Orden clásico": info.greedy.order[i],
        "Controlador (rondas)": ctrlName(r.assignment[i]),
        "Distancia rondas (m)": round2(r.distances[i]),
        "Ronda": info.rounds.round[i],
        "Orden rondas": info.rounds.order[i],
      };
    });
    const sg = controllerSummary(g);
    const sr = controllerSummary(r);
    const ctrlRows = sg.map(function (a, c) {
      const b = sr[c];
      return {
        "Controlador": "C" + (c + 1), "Id BTS emplazamiento": a.site.id,
        "X (m)": a.site.x, "Y (m)": a.site.y, "Capacidad": run.capacity,
        "BTS clásico": a.load, "Dist. total clásico (m)": round2(a.total),
        "Dist. media clásico (m)": round2(a.load ? a.total / a.load : 0), "Dist. máx. clásico (m)": round2(a.max),
        "BTS rondas": b.load, "Dist. total rondas (m)": round2(b.total),
        "Dist. media rondas (m)": round2(b.load ? b.total / b.load : 0), "Dist. máx. rondas (m)": round2(b.max),
      };
    });
    const mg = metrics(g), mr = metrics(r);
    const cmpRows = [
      ["BTS asignadas", mg.assigned, mr.assigned],
      ["BTS sin asignar", mg.unassigned, mr.unassigned],
      ["Distancia total (m)", round2(mg.total), round2(mr.total)],
      ["Distancia media (m)", round2(mg.mean), round2(mr.mean)],
      ["Distancia máxima (m)", round2(mg.max), round2(mr.max)],
      ["Desviación típica de la distancia (m)", round2(mg.std), round2(mr.std)],
      ["Mín. BTS por controlador", mg.minLoad, mr.minLoad],
      ["Máx. BTS por controlador", mg.maxLoad, mr.maxLoad],
    ].map(function (x) { return { "Indicador": x[0], "Greedy clásico": x[1], "Greedy por rondas": x[2] }; });
    const params = [
      { "Parámetro": "Fichero", "Valor": state.fileName },
      { "Parámetro": "Número de BTS", "Valor": state.bts.length },
      { "Parámetro": "Número de controladores", "Valor": run.controllers.length },
      { "Parámetro": "Máx. BTS por controlador", "Valor": run.capacity },
      { "Parámetro": "Número de rondas", "Valor": run.rounds },
      { "Parámetro": "Límites acumulados por ronda", "Valor": r.limits.join(", ") },
      { "Parámetro": "Semilla", "Valor": run.seed || "(aleatoria)" },
    ];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(assignRows), "Asignacion");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(ctrlRows), "Controladores");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(cmpRows), "Comparacion");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(params), "Parametros");
    XLSX.writeFile(wb, "resultado_red_agregacion.xlsx");
  }

  // ------------------------------------------------------------------ eventos

  els.fileInput.addEventListener("change", function () {
    if (els.fileInput.files.length) onFile(els.fileInput.files[0]);
  });
  els.exampleBtn.addEventListener("click", function () {
    els.fileInput.value = "";
    loadBts(exampleBts(), "datos de ejemplo", []);
  });
  els.numControllers.addEventListener("input", updateCapacityInfo);
  els.capacity.addEventListener("input", updateCapacityInfo);
  els.numRounds.addEventListener("input", updateCapacityInfo);
  els.runBtn.addEventListener("click", run);
  els.tabButtons.forEach(function (b) {
    b.addEventListener("click", function () { if (state.run) showTab(b.dataset.tab); });
  });
  els.stepBack.addEventListener("click", function () { stopPlay(); setShown(state.shown[state.tab] - 1); });
  els.stepFwd.addEventListener("click", function () { stopPlay(); setShown(state.shown[state.tab] + 1); });
  els.stepAll.addEventListener("click", function () { stopPlay(); setShown(Infinity); });
  els.stepSlider.addEventListener("input", function () { stopPlay(); setShown(Number(els.stepSlider.value)); });
  els.playBtn.addEventListener("click", togglePlay);
  els.exportBtn.addEventListener("click", exportResults);
  els.exportBtn2.addEventListener("click", exportResults);
  els.canvas.addEventListener("mousemove", onHover);
  els.canvas.addEventListener("mouseleave", function () { els.tooltip.hidden = true; });
  window.addEventListener("resize", draw);
  if (window.matchMedia) {
    window.matchMedia("(prefers-color-scheme: dark)").addEventListener("change", draw);
  }
  new MutationObserver(draw).observe(document.documentElement, { attributes: true, attributeFilter: ["data-theme"] });

  updateCapacityInfo();
  draw();
})();
