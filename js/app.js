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
    minDist: $("minDist"),
    numRounds: $("numRounds"),
    seed: $("seed"),
    capacityInfo: $("capacityInfo"),
    runBtn: $("runBtn"),
    newSelBtn: $("newSelBtn"),
    selectionInfo: $("selectionInfo"),
    stepCard: $("stepCard"),
    stepHint: $("stepHint"),
    stepBack: $("stepBack"),
    stepFwd: $("stepFwd"),
    stepAll: $("stepAll"),
    playBtn: $("playBtn"),
    stepSlider: $("stepSlider"),
    stepInfo: $("stepInfo"),
    roundNav: $("roundNav"),
    roundCard: $("roundCard"),
    roundTable: $("roundTable"),
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

  // Métodos de asignación. Todos usan los mismos controladores para que la
  // comparación sea justa. `rounds` indica si el método trabaja por rondas.
  const METHODS = {
    greedy: {
      label: "Greedy ordenado", short: "Ordenado", rounds: false,
      assign: function (bts, ctrls, cap) { return A.greedyAssign(bts, ctrls, cap); },
    },
    rounds: {
      label: "Greedy ordenado por rondas", short: "Ordenado por rondas", rounds: true,
      assign: function (bts, ctrls, cap, r) { return A.greedyRoundsAssign(bts, ctrls, cap, r); },
    },
    seq: {
      label: "Greedy secuencial por controlador", short: "Secuencial", rounds: false,
      assign: function (bts, ctrls, cap) { return A.sequentialAssign(bts, ctrls, cap); },
    },
    seqRounds: {
      label: "Greedy secuencial por controlador por rondas", short: "Secuencial por rondas", rounds: true,
      assign: function (bts, ctrls, cap, r) { return A.sequentialRoundsAssign(bts, ctrls, cap, r); },
    },
  };
  const METHOD_KEYS = Object.keys(METHODS);

  function emptyShown() {
    const shown = {};
    METHOD_KEYS.forEach(function (m) { shown[m] = 0; });
    return shown;
  }

  const state = {
    bts: [],
    fileName: "",
    // Controladores elegidos. Se mantienen entre ejecuciones de RUN para poder
    // comparar resultados; solo cambian con "Nueva selección de controladores"
    // o si cambian el número de controladores, la distancia mínima o la semilla.
    selection: null, // { controllers, minSeparation, count, minDist, seed, number }
    selectionCount: 0, // selecciones hechas con la semilla actual
    run: null,      // { controllers, capacity, rounds, minDist, minSeparation, seed, selection, colors, results: {<método>: resultado} }
    tab: "greedy",  // clave de METHODS o "compare"
    shown: null,    // pasos visibles en el mapa de cada método
    timer: null,
    pickedRound: 0, // última ronda elegida con los botones de ronda
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

  // Distancia mínima entre controladores en metros (el campo está en km). NaN si no es válida.
  function readMinDist() {
    const text = String(els.minDist.value).trim().replace(",", ".");
    if (text === "") return 0;
    const km = Number(text);
    return Number.isFinite(km) && km >= 0 ? km * 1000 : NaN;
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
    state.selection = null;
    state.selectionCount = 0;
    state.run = null;
    state.shown = emptyShown();
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
    else if (Number.isNaN(readMinDist())) problem = "La distancia mínima entre controladores debe ser un número ≥ 0 (en km).";
    else if (!(rounds >= 1)) problem = "El número de rondas debe ser un entero ≥ 1.";
    else if (rounds > cap) problem = "El número de rondas (" + rounds + ") no puede superar el máximo de BTS por controlador (" + cap + ").";
    else if (total && n > total) problem = "Hay más controladores (" + n + ") que BTS (" + total + ").";

    els.runBtn.disabled = !!problem || total === 0;
    els.newSelBtn.disabled = els.runBtn.disabled;
    updateSelectionInfo();

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
    const minDist = readMinDist();
    if (minDist > 0) {
      html += "<br>Los controladores se elegirán a " + fmtMeters(minDist) + " o más unos de otros.";
    }
    setInfo(els.capacityInfo, kind, html);
  }

  // ------------------------------------------------------------------ ejecución

  // Quita los resultados anteriores y deja solo las BTS en el mapa.
  function clearRun() {
    state.run = null;
    state.shown = emptyShown();
    els.tabsWrap.hidden = true;
    els.results.hidden = true;
    els.compare.hidden = true;
    els.stepCard.hidden = true;
    els.mapCard.hidden = false;
    els.legend.innerHTML = "";
    draw();
  }

  // Motivo por el que la selección actual no sirve para los parámetros actuales ("" si sirve).
  function selectionStaleReason() {
    const sel = state.selection;
    if (!sel) return "";
    if (sel.count !== readInt(els.numControllers)) return "ha cambiado el número de controladores";
    if (sel.minDist !== readMinDist()) return "ha cambiado la distancia mínima entre controladores";
    if (sel.seed !== els.seed.value.trim()) return "ha cambiado la semilla";
    return "";
  }

  function updateSelectionInfo() {
    const sel = state.selection;
    if (!sel) {
      setInfo(els.selectionInfo, "", "");
      return;
    }
    const stale = selectionStaleReason();
    if (stale) {
      setInfo(els.selectionInfo, "warn",
        "Al pulsar RUN se elegirán nuevos controladores porque " + stale + ".");
      return;
    }
    const ids = sel.controllers.map(function (idx, c) { return "C" + (c + 1) + " = " + escapeHtml(state.bts[idx].id); });
    setInfo(els.selectionInfo, "",
      "<b>Controladores fijados</b> (selección nº " + sel.number + (sel.seed ? ", semilla " + escapeHtml(sel.seed) : "") + "): " +
      (ids.length <= 12 ? ids.join(", ") : ids.slice(0, 12).join(", ") + ", …") +
      ". RUN los mantiene; para cambiarlos pulsa «Nueva selección de controladores».");
  }

  // Elige nuevos controladores. Devuelve false (y muestra el motivo) si no hay solución.
  function newSelection() {
    const n = readInt(els.numControllers);
    const minDist = readMinDist();
    const seed = els.seed.value.trim();
    if (!state.selection || state.selection.seed !== seed) state.selectionCount = 0;
    const number = state.selectionCount + 1;
    // Con semilla, la selección nº k es reproducible: se deriva de la semilla y de k.
    const rng = seed === "" ? Math.random : A.createRng(number === 1 ? seed : seed + "#" + number);
    const sel = A.selectControllersWithMinDistance(state.bts, n, minDist, rng);
    if (!sel.controllers) {
      state.selection = null;
      clearRun();
      setInfo(els.capacityInfo, "error",
        "<b>No se han ejecutado los algoritmos.</b> " +
        (sel.exhaustive
          ? "No existe ninguna forma de elegir " + n + " controladores entre estas BTS con una distancia mínima de " +
            fmtMeters(minDist) + " entre ellos" +
            (sel.best === n - 1 ? ": como máximo caben " + sel.best + "." : " (el mejor reparto encontrado coloca " + sel.best + ").")
          : "No se ha encontrado ninguna forma de elegir " + n + " controladores con una distancia mínima de " +
            fmtMeters(minDist) + " (la búsqueda se ha detenido por ser demasiado larga; el mejor intento coloca " +
            sel.best + "). Es probable que no exista.") +
        " Reduce la distancia mínima o el número de controladores.");
      updateSelectionInfo();
      return false;
    }
    state.selectionCount = number;
    state.selection = {
      controllers: sel.controllers,
      minSeparation: sel.minSeparation,
      count: n,
      minDist: minDist,
      seed: seed,
      number: number,
    };
    return true;
  }

  // Ejecuta los cuatro Greedy. Reutiliza los controladores ya elegidos salvo que se
  // pida una selección nueva o la actual no valga para los parámetros actuales.
  function run(forceNewSelection) {
    stopPlay();
    const cap = readInt(els.capacity);
    const rounds = readInt(els.numRounds);
    try {
      if (forceNewSelection === true || !state.selection || selectionStaleReason()) {
        if (!newSelection()) return;
      }
      const sel = state.selection;
      const controllers = sel.controllers;
      state.run = {
        controllers: controllers,
        capacity: cap,
        rounds: rounds,
        minDist: sel.minDist,
        minSeparation: sel.minSeparation,
        seed: sel.seed,
        selection: sel.number,
        colors: controllerColors(controllers.length),
        results: {},
      };
      state.shown = emptyShown();
      METHOD_KEYS.forEach(function (m) {
        state.run.results[m] = METHODS[m].assign(state.bts, controllers, cap, rounds);
        state.shown[m] = state.run.results[m].steps.length;
      });
    } catch (err) {
      clearRun();
      setInfo(els.capacityInfo, "error", escapeHtml(err.message));
      return;
    }
    updateCapacityInfo();
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

    const perRound = run.capacity / run.rounds;
    const limitsText = "límites acumulados " + A.roundLimits(run.capacity, run.rounds).join(", ");
    if (tab === "greedy") {
      els.methodInfo.textContent = "Recorre todas las parejas (BTS, controlador) de menor a mayor distancia y asigna " +
        "la BTS si aún está libre y el controlador no está lleno (máx. " + run.capacity + " BTS).";
      els.stepHint.textContent = "Recorre las asignaciones en el orden en que las realiza el algoritmo (de menor a mayor distancia).";
    } else if (tab === "rounds") {
      els.methodInfo.textContent = "Igual que el Greedy ordenado, pero en " + run.rounds + " rondas: en cada ronda cada " +
        "controlador solo puede recibir unas " + fmtNum(perRound, 1) + " BTS más (" + limitsText + ").";
      els.stepHint.textContent = "Recorre las asignaciones ronda a ronda; dentro de cada ronda, de menor a mayor distancia.";
    } else if (tab === "seq") {
      els.methodInfo.textContent = "Los controladores se atienden por orden (C1, C2, …): cada uno toma sus BTS libres más " +
        "cercanas hasta llenarse (máx. " + run.capacity + " BTS) y después se pasa al siguiente. " +
        "El emplazamiento de cada controlador queda reservado para él.";
      els.stepHint.textContent = "Recorre las asignaciones controlador a controlador; cada uno, de su BTS más cercana a la más lejana.";
    } else if (tab === "seqRounds") {
      els.methodInfo.textContent = "Igual que el secuencial por controlador, pero en " + run.rounds + " rondas: en cada ronda " +
        "C1, C2, … toman por turno sus BTS libres más cercanas hasta unas " + fmtNum(perRound, 1) + " más (" + limitsText + ").";
      els.stepHint.textContent = "Recorre las asignaciones ronda a ronda; dentro de cada ronda, controlador a controlador.";
    } else {
      els.methodInfo.textContent = "Los cuatro métodos usan los mismos " + run.controllers.length +
        " controladores elegidos al azar, así que las diferencias se deben solo al algoritmo." + separationText();
    }

    if (isMethod) els.methodInfo.textContent += separationText();

    const withRounds = isRoundsTab();
    state.pickedRound = 0;
    els.roundNav.hidden = !withRounds;
    els.roundCard.hidden = !withRounds;
    if (withRounds) renderRounds();

    if (isMethod) {
      const res = current();
      els.stepSlider.max = res.steps.length;
      renderLegend(res);
      renderStats(res);
      renderTables(res, tab);
      updateStep();
    }
  }

  function isRoundsTab() {
    return !!(METHODS[state.tab] && METHODS[state.tab].rounds);
  }

  function separationText() {
    const run = state.run;
    if (run.minSeparation === null) return "";
    return " Distancia mínima entre los controladores elegidos: " + fmtMeters(run.minSeparation) +
      (run.minDist > 0 ? " (exigida: " + fmtMeters(run.minDist) + ")." : ".");
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
    } else if (isRoundsTab() && roundAtEnd(shown)) {
      const r = roundAtEnd(shown);
      const ends = roundEnds(res);
      const added = shown - (r === 1 ? 0 : ends[r - 2]);
      const pending = state.bts.length - shown;
      setInfo(els.stepInfo, "",
        "<b>Fin de la ronda " + r + "</b> de " + state.run.rounds + " (límite " + res.limits[r - 1] +
        " BTS por controlador): " + (added ? added + " BTS asignadas en esta ronda, " : "en esta ronda no se asigna ninguna BTS; ") +
        shown + " en total" + (pending ? " y " + pending + " aún sin asignar (apagadas en el mapa)." : "."));
    } else {
      const s = steps[shown - 1];
      setInfo(els.stepInfo, "",
        (isRoundsTab() ? "Ronda <b>" + s.round + "</b> de " + state.run.rounds + " · " : "") +
        "Paso <b>" + shown + "</b> de " + steps.length + ": " +
        escapeHtml(state.bts[s.bts].id) + " → <b>C" + (s.controller + 1) + "</b> (" + fmtMeters(s.distance) + ")");
    }
    if (isRoundsTab()) highlightRound();
    draw();
  }

  // ------------------------------------------------------------------ rondas

  // Número de pasos acumulados al final de cada ronda: ends[r - 1].
  function roundEnds(res) {
    const ends = [];
    for (let r = 1; r <= state.run.rounds; r++) {
      let k = r === 1 ? 0 : ends[r - 2];
      while (k < res.steps.length && res.steps[k].round <= r) k++;
      ends.push(k);
    }
    return ends;
  }

  // Ronda que termina justo en el paso `shown` (0 si no coincide con el final de una ronda).
  // Si varias rondas acaban en el mismo paso (rondas en las que no se asigna nada),
  // se respeta la que se haya elegido con los botones; si no, la primera.
  function roundAtEnd(shown) {
    const ends = roundEnds(current());
    const picked = state.pickedRound;
    if (picked && ends[picked - 1] === shown) return picked;
    const r = ends.indexOf(shown);
    return r === -1 ? 0 : r + 1;
  }

  // Ronda a la que pertenece lo que se ve en el mapa (0 antes de empezar).
  function currentRound() {
    const shown = state.shown[state.tab];
    return shown === 0 ? 0 : current().steps[shown - 1].round;
  }

  function showRound(r) {
    stopPlay();
    state.pickedRound = r;
    setShown(r === 0 ? 0 : roundEnds(current())[r - 1]);
  }

  function renderRounds() {
    const run = state.run;
    const res = current();
    const ends = roundEnds(res);

    let nav = "<button type=\"button\" data-round=\"0\" aria-pressed=\"false\">Inicio</button>";
    for (let r = 1; r <= run.rounds; r++) {
      nav += "<button type=\"button\" data-round=\"" + r + "\" aria-pressed=\"false\">Ronda " + r + "</button>";
    }
    els.roundNav.innerHTML = nav;

    let html = "<thead><tr><th class=\"num\">Ronda</th><th class=\"num\">Límite por controlador</th>" +
      "<th class=\"num\">Asignadas en la ronda</th><th class=\"num\">Asignadas acumuladas</th>" +
      "<th class=\"num\">Sin asignar</th><th class=\"num\">Dist. media de la ronda</th>" +
      "<th class=\"num\">Dist. máx. de la ronda</th></tr></thead><tbody>";
    for (let r = 1; r <= run.rounds; r++) {
      const from = r === 1 ? 0 : ends[r - 2];
      const to = ends[r - 1];
      let sum = 0, max = 0;
      for (let k = from; k < to; k++) { sum += res.steps[k].distance; max = Math.max(max, res.steps[k].distance); }
      const count = to - from;
      html += "<tr data-round=\"" + r + "\" tabindex=\"0\" style=\"cursor:pointer\"><td class=\"num\">" + r + "</td>" +
        "<td class=\"num\">" + res.limits[r - 1] + "</td>" +
        "<td class=\"num\">" + count + "</td><td class=\"num\">" + to + "</td>" +
        "<td class=\"num\">" + (state.bts.length - to) + "</td>" +
        "<td class=\"num\">" + (count ? fmtMeters(sum / count) : "—") + "</td>" +
        "<td class=\"num\">" + (count ? fmtMeters(max) : "—") + "</td></tr>";
    }
    els.roundTable.innerHTML = html + "</tbody>";
  }

  // Marca el botón y la fila de la ronda que se ve en el mapa.
  function highlightRound() {
    const shown = state.shown[state.tab];
    const atEnd = shown === 0 ? 0 : roundAtEnd(shown);
    const cur = currentRound();
    Array.prototype.forEach.call(els.roundNav.querySelectorAll("button"), function (b) {
      const r = Number(b.dataset.round);
      b.setAttribute("aria-pressed", (shown === 0 ? r === 0 : atEnd === r) ? "true" : "false");
    });
    Array.prototype.forEach.call(els.roundTable.querySelectorAll("tbody tr"), function (tr) {
      tr.classList.toggle("current", Number(tr.dataset.round) === cur);
    });
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
      // Enlaces BTS -> controlador de los pasos visibles. En el método por rondas,
      // los de rondas anteriores se dibujan más tenues que los de la ronda actual.
      const lastRound = isRoundsTab() && shown > 0 ? steps[shown - 1].round : 0;
      ctx.lineWidth = 1.2;
      for (let k = 0; k < shown; k++) {
        const s = steps[k];
        owner[s.bts] = s.controller;
        ctx.globalAlpha = lastRound && s.round < lastRound ? 0.22 : 0.65;
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
      if (owner[i] === -1) {
        // Sin asignar todavía: punto apagado.
        ctx.fillStyle = res ? cssVar("--pending") : muted;
        ctx.beginPath(); ctx.arc(x, y, 3, 0, 2 * Math.PI); ctx.fill();
        return;
      }
      ctx.fillStyle = run.colors[owner[i]];
      ctx.beginPath(); ctx.arc(x, y, 4, 0, 2 * Math.PI); ctx.fill();
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
    html += "<span>■ controlador · ● BTS asignada</span>" +
      "<span><i style=\"background:var(--pending);border-radius:50%\"></i>BTS aún sin asignar</span>" +
      (res.unassigned.length ? "<span>✕ sin asignar al final</span>" : "");
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

    const withRound = METHODS[method].rounds;
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
    const res = METHOD_KEYS.map(function (m) { return run.results[m]; });
    const ms = res.map(metrics);
    const base = run.results.greedy;
    const changed = res.map(function (r) {
      let n = 0;
      r.assignment.forEach(function (c, i) { if (c !== base.assignment[i]) n++; });
      return n;
    });

    // [etiqueta, valores, formato, true si "menos es mejor" / false si "más es mejor" / null si no se compara]
    const rows = [
      ["BTS asignadas", ms.map(function (m) { return m.assigned; }), String, false],
      ["BTS sin asignar", ms.map(function (m) { return m.unassigned; }), String, true],
      ["Distancia total", ms.map(function (m) { return m.total; }), fmtMeters, true],
      ["Distancia media", ms.map(function (m) { return m.mean; }), fmtMeters, true],
      ["Distancia máxima", ms.map(function (m) { return m.max; }), fmtMeters, true],
      ["Desviación típica de la distancia", ms.map(function (m) { return m.std; }), fmtMeters, true],
      ["BTS por controlador (mín. – máx.)", ms.map(function (m) { return m.minLoad + " – " + m.maxLoad; }), String, null],
      ["BTS con distinto controlador que el Greedy ordenado", changed, String, null],
    ];

    let html = "<thead><tr><th>Indicador</th>" + METHOD_KEYS.map(function (m) {
      return "<th class=\"num\" title=\"" + METHODS[m].label + "\">" + METHODS[m].short +
        (METHODS[m].rounds ? "<br>(" + run.rounds + " rondas)" : "") + "</th>";
    }).join("") + "</tr></thead><tbody>";
    rows.forEach(function (row) {
      const vals = row[1], fmt = row[2], lessIsBetter = row[3];
      let best = null;
      if (lessIsBetter !== null) {
        best = lessIsBetter ? Math.min.apply(null, vals) : Math.max.apply(null, vals);
        // Si todos empatan no se marca ninguno.
        if (vals.every(function (v) { return Math.abs(v - best) <= 1e-9; })) best = null;
      }
      html += "<tr><td>" + row[0] + "</td>" + vals.map(function (v, k) {
        const isBest = best !== null && Math.abs(v - best) <= 1e-9;
        const text = row[0].indexOf("distinto controlador") !== -1 && k === 0 ? "—" : fmt(v);
        return "<td class=\"num" + (isBest ? " better" : "") + "\">" + text + "</td>";
      }).join("") + "</tr>";
    });
    els.compareTable.innerHTML = html + "</tbody>";
    els.compareNote.textContent = "En verde, el método o métodos con mejor valor en cada indicador.";

    const sums = res.map(controllerSummary);
    html = "<thead><tr><th>Controlador</th><th>Emplazamiento</th>" +
      METHOD_KEYS.map(function (m) { return "<th class=\"num\">BTS<br>" + METHODS[m].short + "</th>"; }).join("") +
      METHOD_KEYS.map(function (m) { return "<th class=\"num\">Dist. media<br>" + METHODS[m].short + "</th>"; }).join("") +
      "</tr></thead><tbody>";
    sums[0].forEach(function (a, c) {
      html += "<tr><td>" + swatch(run.colors[c]) + "C" + (c + 1) + "</td><td>" + escapeHtml(a.site.id) + "</td>" +
        sums.map(function (sm) { return "<td class=\"num\">" + sm[c].load + "</td>"; }).join("") +
        sums.map(function (sm) {
          return "<td class=\"num\">" + fmtMeters(sm[c].load ? sm[c].total / sm[c].load : 0) + "</td>";
        }).join("") + "</tr>";
    });
    els.compareCtrlTable.innerHTML = html + "</tbody>";
  }

  function exportResults() {
    const run = state.run;
    const round2 = function (v) { return v === null ? null : Math.round(v * 100) / 100; };
    const ctrlName = function (c) { return c === -1 ? "Sin asignar" : "C" + (c + 1); };
    const info = {};
    METHOD_KEYS.forEach(function (m) {
      const order = new Array(state.bts.length).fill(null);
      const rnd = new Array(state.bts.length).fill(null);
      run.results[m].steps.forEach(function (s, k) { order[s.bts] = k + 1; rnd[s.bts] = s.round; });
      info[m] = { order: order, round: rnd, summary: controllerSummary(run.results[m]), metrics: metrics(run.results[m]) };
    });

    const assignRows = state.bts.map(function (b, i) {
      const row = { "Id BTS": b.id, "X (m)": b.x, "Y (m)": b.y };
      METHOD_KEYS.forEach(function (m) {
        const res = run.results[m], name = METHODS[m].short;
        row["Controlador (" + name + ")"] = ctrlName(res.assignment[i]);
        row["Distancia " + name + " (m)"] = round2(res.distances[i]);
        if (METHODS[m].rounds) row["Ronda " + name] = info[m].round[i];
        row["Orden " + name] = info[m].order[i];
      });
      return row;
    });
    const ctrlRows = run.controllers.map(function (idx, c) {
      const site = state.bts[idx];
      const row = {
        "Controlador": "C" + (c + 1), "Id BTS emplazamiento": site.id,
        "X (m)": site.x, "Y (m)": site.y, "Capacidad": run.capacity,
      };
      METHOD_KEYS.forEach(function (m) {
        const sm = info[m].summary[c], name = METHODS[m].short;
        row["BTS " + name] = sm.load;
        row["Dist. total " + name + " (m)"] = round2(sm.total);
        row["Dist. media " + name + " (m)"] = round2(sm.load ? sm.total / sm.load : 0);
        row["Dist. máx. " + name + " (m)"] = round2(sm.max);
      });
      return row;
    });
    const cmpDefs = [
      ["BTS asignadas", function (x) { return x.assigned; }],
      ["BTS sin asignar", function (x) { return x.unassigned; }],
      ["Distancia total (m)", function (x) { return round2(x.total); }],
      ["Distancia media (m)", function (x) { return round2(x.mean); }],
      ["Distancia máxima (m)", function (x) { return round2(x.max); }],
      ["Desviación típica de la distancia (m)", function (x) { return round2(x.std); }],
      ["Mín. BTS por controlador", function (x) { return x.minLoad; }],
      ["Máx. BTS por controlador", function (x) { return x.maxLoad; }],
    ];
    const cmpRows = cmpDefs.map(function (d) {
      const row = { "Indicador": d[0] };
      METHOD_KEYS.forEach(function (m) { row[METHODS[m].label] = d[1](info[m].metrics); });
      return row;
    });
    const params = [
      { "Parámetro": "Fichero", "Valor": state.fileName },
      { "Parámetro": "Número de BTS", "Valor": state.bts.length },
      { "Parámetro": "Número de controladores", "Valor": run.controllers.length },
      { "Parámetro": "Máx. BTS por controlador", "Valor": run.capacity },
      { "Parámetro": "Distancia mínima exigida entre controladores (m)", "Valor": run.minDist },
      { "Parámetro": "Distancia mínima entre los controladores elegidos (m)", "Valor": round2(run.minSeparation) },
      { "Parámetro": "Número de rondas", "Valor": run.rounds },
      { "Parámetro": "Límites acumulados por ronda", "Valor": A.roundLimits(run.capacity, run.rounds).join(", ") },
      { "Parámetro": "Semilla", "Valor": run.seed || "(aleatoria)" },
      { "Parámetro": "Selección de controladores nº", "Valor": run.selection },
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
  els.minDist.addEventListener("input", updateCapacityInfo);
  els.runBtn.addEventListener("click", function () { run(false); });
  els.newSelBtn.addEventListener("click", function () { run(true); });
  els.seed.addEventListener("input", updateSelectionInfo);
  els.tabButtons.forEach(function (b) {
    b.addEventListener("click", function () { if (state.run) showTab(b.dataset.tab); });
  });
  els.stepBack.addEventListener("click", function () { stopPlay(); setShown(state.shown[state.tab] - 1); });
  els.stepFwd.addEventListener("click", function () { stopPlay(); setShown(state.shown[state.tab] + 1); });
  els.stepAll.addEventListener("click", function () { stopPlay(); setShown(Infinity); });
  els.stepSlider.addEventListener("input", function () { stopPlay(); setShown(Number(els.stepSlider.value)); });
  els.playBtn.addEventListener("click", togglePlay);
  els.roundNav.addEventListener("click", function (ev) {
    const b = ev.target.closest("button[data-round]");
    if (b) showRound(Number(b.dataset.round));
  });
  els.roundTable.addEventListener("click", function (ev) {
    const tr = ev.target.closest("tr[data-round]");
    if (tr) {
      showRound(Number(tr.dataset.round));
      els.mapCard.scrollIntoView({ behavior: "smooth", block: "start" });
    }
  });
  els.roundTable.addEventListener("keydown", function (ev) {
    const tr = ev.target.closest("tr[data-round]");
    if (tr && (ev.key === "Enter" || ev.key === " ")) { ev.preventDefault(); showRound(Number(tr.dataset.round)); }
  });
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
