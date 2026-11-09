/*
 * Lógica de la red de agregación (independiente de la interfaz).
 *
 *  - Lectura y validación de las filas del Excel (Id, X, Y).
 *  - Generador pseudoaleatorio con semilla (para repetir experimentos).
 *  - Selección aleatoria de los emplazamientos de los controladores, con una
 *    distancia mínima opcional entre ellos.
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

  /**
   * Elige aleatoriamente `count` emplazamientos de BTS para los controladores
   * de forma que la distancia entre cualquier pareja de controladores sea
   * al menos `minDistance` (en metros).
   *
   * 1. Intentos aleatorios: se recorren las BTS en un orden aleatorio y se toma
   *    cada una que quede a `minDistance` o más de las ya elegidas. Se repite
   *    hasta reunir `count` controladores o agotar los intentos.
   * 2. Si ningún intento lo consigue, búsqueda exhaustiva con vuelta atrás para
   *    encontrar una selección o demostrar que no existe. Para no bloquear la
   *    aplicación se limita el número de nodos explorados.
   *
   * Con `minDistance` <= 0 equivale a selectControllers.
   *
   * Devuelve:
   *   controllers   índices de las BTS elegidas, o null si no se ha encontrado solución
   *   best          mayor número de controladores que se ha conseguido colocar
   *                 respetando la distancia mínima
   *   exhaustive    false si la búsqueda se cortó por el límite de nodos; con
   *                 controllers === null y exhaustive === true está demostrado
   *                 que no hay solución
   *   minSeparation menor distancia entre dos controladores elegidos (null si hay < 2)
   */
  function selectControllersWithMinDistance(bts, count, minDistance, rng, options) {
    const n = bts.length;
    const random = rng || Math.random;

    if (!(minDistance > 0)) {
      const chosen = selectControllers(n, count, random);
      return { controllers: chosen, best: count, exhaustive: true, minSeparation: minSeparation(bts, chosen) };
    }
    selectControllers(n, count, function () { return 0; }); // misma validación de `count`

    const opts = options || {};
    const attempts = opts.attempts || 300;
    const maxNodes = opts.maxNodes || 2000000;
    const min2 = minDistance * minDistance;
    const far = function (a, b) {
      const dx = bts[a].x - bts[b].x, dy = bts[a].y - bts[b].y;
      return dx * dx + dy * dy >= min2;
    };
    const shuffled = function () {
      const order = [];
      for (let i = 0; i < n; i++) order.push(i);
      for (let i = n - 1; i > 0; i--) {
        const j = Math.floor(random() * (i + 1));
        const tmp = order[i]; order[i] = order[j]; order[j] = tmp;
      }
      return order;
    };
    const result = function (chosen, best, exhaustive) {
      return {
        controllers: chosen,
        best: best,
        exhaustive: exhaustive,
        minSeparation: chosen ? minSeparation(bts, chosen) : null,
      };
    };

    // 1. Intentos aleatorios.
    let best = 0;
    for (let a = 0; a < attempts; a++) {
      const chosen = [];
      const order = shuffled();
      for (let k = 0; k < n && chosen.length < count; k++) {
        const c = order[k];
        if (chosen.every(function (o) { return far(c, o); })) chosen.push(c);
      }
      if (chosen.length === count) return result(chosen, count, true);
      if (chosen.length > best) best = chosen.length;
    }

    // 2. Búsqueda exhaustiva con vuelta atrás.
    const picked = [];
    let nodes = 0;
    let aborted = false;
    function search(candidates) {
      if (picked.length === count) return true;
      if (++nodes > maxNodes) { aborted = true; return false; }
      for (let k = 0; k < candidates.length; k++) {
        if (picked.length + candidates.length - k < count) return false; // no quedan suficientes
        const c = candidates[k];
        const rest = [];
        for (let m = k + 1; m < candidates.length; m++) {
          if (far(c, candidates[m])) rest.push(candidates[m]);
        }
        picked.push(c);
        if (picked.length > best) best = picked.length;
        if (search(rest)) return true;
        picked.pop();
        if (aborted) return false;
      }
      return false;
    }
    if (search(shuffled())) return result(picked.slice(), count, true);
    return result(null, best, !aborted);
  }

  // Menor distancia entre dos de los emplazamientos indicados (null si hay menos de 2).
  function minSeparation(bts, indices) {
    let min = null;
    for (let a = 0; a < indices.length; a++) {
      for (let b = a + 1; b < indices.length; b++) {
        const d = distance(bts[indices[a]], bts[indices[b]]);
        if (min === null || d < min) min = d;
      }
    }
    return min;
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
    selectControllersWithMinDistance: selectControllersWithMinDistance,
    minSeparation: minSeparation,
    greedyAssign: greedyAssign,
    greedyRoundsAssign: greedyRoundsAssign,
    roundLimits: roundLimits,
    distance: distance,
  };
});
