const test = require("node:test");
const assert = require("node:assert/strict");
const A = require("../js/algorithm.js");

test("parseBtsRows ignora la cabecera y las filas vacías", () => {
  const { bts, errors } = A.parseBtsRows([
    ["Id", "X", "Y"],
    ["BTS1", 0, 0],
    [],
    ["BTS2", "10,5", "3"],
  ]);
  assert.deepEqual(bts, [{ id: "BTS1", x: 0, y: 0 }, { id: "BTS2", x: 10.5, y: 3 }]);
  assert.deepEqual(errors, []);
});

test("parseBtsRows informa de filas erróneas y duplicados", () => {
  const { bts, errors } = A.parseBtsRows([
    ["A", 1, 1],
    ["B", "x", 1],
    ["A", 2, 2],
    [null, 3, 3],
  ]);
  assert.equal(bts.length, 1);
  assert.equal(errors.length, 3);
});

test("createRng con semilla es reproducible", () => {
  const a = A.createRng("42"), b = A.createRng("42");
  for (let i = 0; i < 5; i++) assert.equal(a(), b());
});

test("selectControllers devuelve índices distintos y valida el número", () => {
  const idx = A.selectControllers(10, 4, A.createRng(1));
  assert.equal(new Set(idx).size, 4);
  idx.forEach((i) => assert.ok(i >= 0 && i < 10));
  assert.throws(() => A.selectControllers(3, 4, Math.random));
  assert.throws(() => A.selectControllers(3, 0, Math.random));
});

test("greedyAssign asigna por distancia respetando la capacidad", () => {
  // Dos controladores en x=0 y x=100; tres BTS cerca de x=0 y capacidad 2.
  const bts = [
    { id: "C0", x: 0, y: 0 },
    { id: "C1", x: 100, y: 0 },
    { id: "a", x: 10, y: 0 },
    { id: "b", x: 20, y: 0 },
  ];
  const r = A.greedyAssign(bts, [0, 1], 2);
  // C0 recibe su propio emplazamiento y la BTS "a"; "b" se desborda a C1.
  assert.deepEqual(r.assignment, [0, 1, 0, 1]);
  assert.deepEqual(r.loads, [2, 2]);
  assert.deepEqual(r.unassigned, []);
  assert.equal(r.distances[3], 80);
  assert.equal(r.totalDistance, 0 + 0 + 10 + 80);
  // Los pasos van en orden no decreciente de distancia.
  for (let k = 1; k < r.steps.length; k++) assert.ok(r.steps[k].distance >= r.steps[k - 1].distance);
});

test("greedyAssign deja BTS sin asignar si no hay capacidad suficiente", () => {
  const bts = [0, 1, 2, 3, 4].map((i) => ({ id: "B" + i, x: i * 10, y: 0 }));
  const r = A.greedyAssign(bts, [0], 3);
  assert.deepEqual(r.loads, [3]);
  assert.deepEqual(r.unassigned, [3, 4]);
  assert.equal(r.steps.length, 3);
});

test("roundLimits reparte la capacidad entre las rondas", () => {
  assert.deepEqual(A.roundLimits(20, 3), [7, 13, 20]);
  assert.deepEqual(A.roundLimits(10, 5), [2, 4, 6, 8, 10]);
  assert.deepEqual(A.roundLimits(7, 1), [7]);
});

test("greedyRoundsAssign con una ronda coincide con el Greedy clásico", () => {
  const rng = A.createRng("rondas");
  const bts = Array.from({ length: 60 }, (_, i) => ({ id: "B" + i, x: rng() * 1000, y: rng() * 1000 }));
  const ctrls = A.selectControllers(bts.length, 4, rng);
  const a = A.greedyAssign(bts, ctrls, 12);
  const b = A.greedyRoundsAssign(bts, ctrls, 12, 1);
  assert.deepEqual(b.assignment, a.assignment);
  assert.equal(b.totalDistance, a.totalDistance);
});

test("greedyRoundsAssign limita las asignaciones de cada ronda", () => {
  // Controladores en x=0 (A) y x=100 (B); dos BTS cerca de A; capacidad 3.
  const bts = [
    { id: "A", x: 0, y: 0 },
    { id: "B", x: 100, y: 0 },
    { id: "a1", x: 1, y: 0 },
    { id: "a2", x: 2, y: 0 },
  ];
  // Greedy clásico: A se queda con a1 y a2.
  assert.deepEqual(A.greedyAssign(bts, [0, 1], 3).assignment, [0, 1, 0, 0]);
  // Rondas (límites 1, 2, 3): en la ronda 2 A solo admite una BTS más (a1),
  // así que a2 pasa a B en esa misma ronda.
  const r = A.greedyRoundsAssign(bts, [0, 1], 3, 3);
  assert.deepEqual(r.limits, [1, 2, 3]);
  assert.deepEqual(r.assignment, [0, 1, 0, 1]);
  assert.deepEqual(r.steps.map((s) => s.round), [1, 1, 2, 2]);
  assert.deepEqual(r.loads, [2, 2]);
});

test("greedyRoundsAssign asigna tantas BTS como permite la capacidad total", () => {
  const rng = A.createRng("total");
  const bts = Array.from({ length: 50 }, (_, i) => ({ id: "B" + i, x: rng() * 1000, y: rng() * 1000 }));
  const ctrls = A.selectControllers(bts.length, 3, rng);
  const r = A.greedyRoundsAssign(bts, ctrls, 10, 4);
  assert.equal(r.steps.length, 30);
  assert.deepEqual(r.loads, [10, 10, 10]);
  assert.equal(r.unassigned.length, 20);
});

test("greedyRoundsAssign valida el número de rondas", () => {
  const bts = [{ id: "A", x: 0, y: 0 }];
  assert.throws(() => A.greedyRoundsAssign(bts, [0], 3, 0));
  assert.throws(() => A.greedyRoundsAssign(bts, [0], 3, 4));
});

test("selectControllersWithMinDistance respeta la distancia mínima", () => {
  const rng = A.createRng("separacion");
  const bts = Array.from({ length: 80 }, (_, i) => ({ id: "B" + i, x: rng() * 200000, y: rng() * 200000 }));
  const r = A.selectControllersWithMinDistance(bts, 6, 50000, A.createRng(3));
  assert.equal(r.controllers.length, 6);
  assert.equal(new Set(r.controllers).size, 6);
  assert.ok(r.minSeparation >= 50000);
  // Con la misma semilla se obtiene la misma selección.
  assert.deepEqual(A.selectControllersWithMinDistance(bts, 6, 50000, A.createRng(3)).controllers, r.controllers);
});

test("selectControllersWithMinDistance detecta que no hay solución", () => {
  // Cuatro BTS en las esquinas de un cuadrado de 100 m: caben 4 a 100 m, pero a 141 m solo caben las dos de una diagonal (141,4 m).
  const bts = [[0, 0], [100, 0], [0, 100], [100, 100]].map((p, i) => ({ id: "B" + i, x: p[0], y: p[1] }));
  assert.equal(A.selectControllersWithMinDistance(bts, 4, 100, A.createRng(1)).controllers.length, 4);
  const r = A.selectControllersWithMinDistance(bts, 3, 141, A.createRng(1));
  assert.equal(r.controllers, null);
  assert.equal(r.exhaustive, true);
  assert.equal(r.best, 2);
});

test("selectControllersWithMinDistance sin distancia mínima equivale a selectControllers", () => {
  const bts = Array.from({ length: 20 }, (_, i) => ({ id: "B" + i, x: i, y: 0 }));
  const r = A.selectControllersWithMinDistance(bts, 5, 0, A.createRng("x"));
  assert.deepEqual(r.controllers, A.selectControllers(20, 5, A.createRng("x")));
  assert.throws(() => A.selectControllersWithMinDistance(bts, 21, 10, A.createRng("x")));
});

test("sequentialAssign llena cada controlador por orden antes de pasar al siguiente", () => {
  // A en x=0, B en x=10; p en x=6 está más cerca de B (4) que de A (6).
  const bts = [
    { id: "A", x: 0, y: 0 },
    { id: "B", x: 10, y: 0 },
    { id: "p", x: 6, y: 0 },
  ];
  // Greedy ordenado: p va a B, su controlador más cercano.
  assert.deepEqual(A.greedyAssign(bts, [0, 1], 2).assignment, [0, 1, 1]);
  // Secuencial: A elige primero y se queda con p.
  const r = A.sequentialAssign(bts, [0, 1], 2);
  assert.deepEqual(r.assignment, [0, 1, 0]);
  assert.deepEqual(r.steps.map((s) => s.controller), [0, 0, 1]);
  // Cambiando el orden de los controladores, p va a B.
  assert.deepEqual(A.sequentialAssign(bts, [1, 0], 2).assignment, [1, 0, 0]);
});

test("sequentialAssign reserva a cada controlador su propio emplazamiento", () => {
  const bts = [
    { id: "A", x: 0, y: 0 },
    { id: "B", x: 1, y: 0 },
    { id: "c", x: 2, y: 0 },
  ];
  const r = A.sequentialAssign(bts, [0, 1], 5);
  assert.equal(r.assignment[0], 0);
  assert.equal(r.assignment[1], 1);
  assert.equal(r.assignment[2], 0); // A no puede tomar B, pero sí c
});

test("sequentialRoundsAssign reparte por rondas y controladores", () => {
  const bts = [
    { id: "A", x: 0, y: 0 },
    { id: "B", x: 100, y: 0 },
    { id: "a1", x: 1, y: 0 },
    { id: "a2", x: 2, y: 0 },
  ];
  // Límites 1, 2, 3. Ronda 1: A toma A, B toma B. Ronda 2: A toma a1, B toma a2.
  const r = A.sequentialRoundsAssign(bts, [0, 1], 3, 3);
  assert.deepEqual(r.assignment, [0, 1, 0, 1]);
  assert.deepEqual(r.steps.map((s) => [s.bts, s.round]), [[0, 1], [1, 1], [2, 2], [3, 2]]);
  // Con una ronda coincide con el secuencial normal.
  assert.deepEqual(A.sequentialRoundsAssign(bts, [0, 1], 3, 1).assignment, A.sequentialAssign(bts, [0, 1], 3).assignment);
});

test("sequentialRoundsAssign asigna tantas BTS como permite la capacidad total", () => {
  const rng = A.createRng("secuencial");
  const bts = Array.from({ length: 50 }, (_, i) => ({ id: "B" + i, x: rng() * 1000, y: rng() * 1000 }));
  const ctrls = A.selectControllers(bts.length, 4, rng);
  const r = A.sequentialRoundsAssign(bts, ctrls, 10, 3);
  assert.equal(r.steps.length, 40);
  assert.deepEqual(r.loads, [10, 10, 10, 10]);
  ctrls.forEach((site, c) => assert.equal(r.assignment[site], c));
  assert.throws(() => A.sequentialRoundsAssign(bts, ctrls, 3, 4));
});
