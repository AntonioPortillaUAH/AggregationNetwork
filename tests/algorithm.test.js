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
