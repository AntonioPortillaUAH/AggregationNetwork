/*
 * Lógica de la red de agregación (independiente de la interfaz).
 *
 *  - Lectura y validación de las filas del Excel (Id, X, Y).
 *  - Generador pseudoaleatorio con semilla (para repetir experimentos).
 *  - Selección aleatoria de los emplazamientos de los controladores.
 *  - Algoritmo Greedy de asignación BTS -> controlador por distancia
 *    respetando la capacidad máxima de cada controlador, en su versión
 *    clásica y en su versión por rondas.
 *
 * Funciona tanto en el navegador (window.Aggregation) como en Node (require).
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) {
    module.exports = factory();
  } else {
    root.Aggregation = factory();
  }
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  function toNumber(value) {
    if (typeof value === "number") return value;
    if (typeof value !== "string") return NaN;
    const text = value.trim().replace(",", ".");
    return text === "" ? NaN : Number(text);
  }

  /**
   * Convierte las filas de la hoja (array de arrays) en una lista de BTS.
   * Se usan las tres primeras columnas: Id, X, Y. Si la primera fila no
   * contiene coordenadas numéricas se considera cabecera y se ignora.
   * Las filas vacías se saltan; el resto de filas no válidas se reportan.
   */
  function parseBtsRows(rows) {
    const bts = [];
    const errors = [];
    const seenIds = new Set();

    rows.forEach(function (row, i) {
      const rowNumber = i + 1;
      if (!row || row.every(function (c) { return c === null || c === undefined || String(c).trim() === ""; })) {
        return;
      }
      const id = row[0] === undefined || row[0] === null ? "" : String(row[0]).trim();
      const x = toNumber(row[1]);
      const y = toNumber(row[2]);

      if (!Number.isFinite(x) || !Number.isFinite(y)) {
        if (bts.length === 0 && errors.length === 0) return; // cabecera
        errors.push("Fila " + rowNumber + ": coordenadas X/Y no numéricas.");
        return;
      }
      if (id === "") {
        errors.push("Fila " + rowNumber + ": falta el Id de la BTS.");
        return;
      }
      if (seenIds.has(id)) {
        errors.push("Fila " + rowNumber + ": Id de BTS duplicado (" + id + ").");
        return;
      }
      seenIds.add(id);
      bts.push({ id: id, x: x, y: y });
    });

    return { bts: bts, errors: errors };
  }

  /**
   * Generador pseudoaleatorio. Con semilla (número o texto) es reproducible
   * (mulberry32); sin semilla usa Math.random.
   */
  function createRng(seed) {
    if (seed === undefined || seed === null || String(seed).trim() === "") {
      return Math.random;
    }
    let h = 1779033703 ^ String(seed).length;
    const s = String(seed);
    for (let i = 0; i < s.length; i++) {
      h = Math.imul(h ^ s.charCodeAt(i), 3432918353);
      h = (h << 13) | (h >>> 19);
    }
    let a = h >>> 0;
    return function () {
      a = (a + 0x6d2b79f5) | 0;
      let t = Math.imul(a ^ (a >>> 15), 1 | a);
      t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
  }

  /**
   * Elige aleatoriamente `count` emplazamientos de BTS distintos donde se
   * instalarán los controladores (Fisher-Yates parcial).
   * Devuelve los índices de las BTS elegidas.
   */
  function selectControllers(btsCount, count, rng) {
    if (!Number.isInteger(count) || count < 1) {
      throw new Error("El número de controladores debe ser un entero mayor que 0.");
    }
    if (count > btsCount) {
      throw new Error("No puede haber más controladores (" + count + ") que BTS (" + btsCount + ").");
    }
    const indices = [];
    for (let i = 0; i < btsCount; i++) indices.push(i);
    for (let i = 0; i < count; i++) {
      const j = i + Math.floor((rng || Math.random)() * (btsCount - i));
      const tmp = indices[i];
      indices[i] = indices[j];
      indices[j] = tmp;
    }
    return indices.slice(0, count);
  }

  function distance(a, b) {
    return Math.hypot(a.x - b.x, a.y - b.y);
  }

  function sortedPairs(bts, controllers) {
    const pairs = [];
    for (let i = 0; i < bts.length; i++) {
      for (let c = 0; c < controllers.length; c++) {
        pairs.push({ bts: i, controller: c, distance: distance(bts[i], bts[controllers[c]]) });
      }
    }
    // Orden estable: a igual distancia, primero la BTS y el controlador de menor índice.
    pairs.sort(function (p, q) {
      return p.distance - q.distance || p.bts - q.bts || p.controller - q.controller;
    });
    return pairs;
  }

  /**
   * Límite acumulado de BTS por controlador al final de cada ronda.
   * En cada ronda cada controlador puede recibir capacidad / rondas BTS más
   * (redondeado), de modo que en la última ronda el límite es la capacidad.
   * Ej.: capacidad 20 y 3 rondas -> [7, 13, 20].
   */
  function roundLimits(capacity, rounds) {
    const limits = [];
    for (let r = 1; r <= rounds; r++) limits.push(Math.round((capacity * r) / rounds));
    return limits;
  }

  /**
   * Algoritmo Greedy de asignación por distancia con capacidad, en rondas.
   *
   * 1. Se calculan las distancias de todas las parejas (BTS, controlador).
   * 2. Se ordenan de menor a mayor distancia.
   * 3. En cada ronda r se recorren las parejas en ese orden: si la BTS aún no
   *    está asignada y el controlador no ha alcanzado el límite de la ronda
   *    (limits[r]), se asigna.
   *
   * Con una sola ronda el límite es directamente la capacidad y se obtiene el
   * Greedy clásico.
   *
   * Parámetros:
   *   bts          lista de {id, x, y}
   *   controllers  índices (en `bts`) de los emplazamientos de controlador
   *   capacity     número máximo de BTS por controlador
   *   rounds       número de rondas (1 = Greedy clásico)
   *
   * Devuelve:
   *   assignment[i]   índice del controlador (0..C-1) de la BTS i, o -1
   *   distances[i]    distancia de la BTS i a su controlador (o null)
   *   steps           asignaciones en el orden en que las hace el algoritmo,
   *                   {bts, controller, distance, round}
   *   loads[c]        número de BTS asignadas al controlador c
   *   unassigned      índices de BTS sin controlador (capacidad insuficiente)
   *   totalDistance   suma de las distancias de todas las asignaciones
   *   limits          límite acumulado por controlador en cada ronda
   */
  function greedyRoundsAssign(bts, controllers, capacity, rounds) {
    if (!Number.isInteger(capacity) || capacity < 1) {
      throw new Error("La capacidad por controlador debe ser un entero mayor que 0.");
    }
    if (!Number.isInteger(rounds) || rounds < 1) {
      throw new Error("El número de rondas debe ser un entero mayor que 0.");
    }
    if (rounds > capacity) {
      throw new Error("El número de rondas (" + rounds + ") no puede ser mayor que el máximo de BTS por controlador (" + capacity + ").");
    }

    const pairs = sortedPairs(bts, controllers);
    const limits = roundLimits(capacity, rounds);
    const assignment = new Array(bts.length).fill(-1);
    const distances = new Array(bts.length).fill(null);
    const loads = new Array(controllers.length).fill(0);
    const steps = [];
    let remaining = Math.min(bts.length, controllers.length * capacity);

    for (let r = 0; r < rounds && remaining > 0; r++) {
      const limit = limits[r];
      for (let k = 0; k < pairs.length && remaining > 0; k++) {
        const p = pairs[k];
        if (assignment[p.bts] !== -1 || loads[p.controller] >= limit) continue;
        assignment[p.bts] = p.controller;
        distances[p.bts] = p.distance;
        loads[p.controller]++;
        remaining--;
        steps.push({ bts: p.bts, controller: p.controller, distance: p.distance, round: r + 1 });
      }
    }

    const unassigned = [];
    let totalDistance = 0;
    for (let i = 0; i < bts.length; i++) {
      if (assignment[i] === -1) unassigned.push(i);
      else totalDistance += distances[i];
    }

    return {
      assignment: assignment,
      distances: distances,
      steps: steps,
      loads: loads,
      unassigned: unassigned,
      totalDistance: totalDistance,
      limits: limits,
    };
  }

  /**
   * Algoritmo Greedy clásico de asignación por distancia con capacidad:
   * se recorren las parejas (BTS, controlador) de menor a mayor distancia y se
   * asigna la BTS si aún está libre y el controlador no está lleno.
   * Equivale a greedyRoundsAssign con una única ronda.
   */
  function greedyAssign(bts, controllers, capacity) {
    return greedyRoundsAssign(bts, controllers, capacity, 1);
  }

  return {
    parseBtsRows: parseBtsRows,
    createRng: createRng,
    selectControllers: selectControllers,
    greedyAssign: greedyAssign,
    greedyRoundsAssign: greedyRoundsAssign,
    roundLimits: roundLimits,
    distance: distance,
  };
});
