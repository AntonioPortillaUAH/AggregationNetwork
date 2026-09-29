# Red de Agregación

Aplicación docente para demostrar el diseño de la **red de agregación** de una red de
comunicaciones móviles: cómo se conectan las estaciones base (BTS) a un conjunto de
controladores.

## Uso

La aplicación es una página web estática: no necesita instalación ni servidor.

1. Abrir `index.html` en un navegador (Chrome, Firefox, Edge…).
2. Cargar un fichero Excel con las estaciones base (o pulsar *Cargar datos de ejemplo*).
3. Indicar el **número de controladores** y el **máximo de BTS por controlador**.
4. Opcionalmente, indicar una **semilla** para que la selección aleatoria sea repetible
   (útil para que todos los alumnos obtengan el mismo resultado).
5. Pulsar **RUN**.

### Formato del fichero de entrada

Primera hoja del libro, tres columnas (la fila de cabecera es opcional):

| Id     | X (m) | Y (m) |
|--------|-------|-------|
| BTS001 | 3540  | 2291  |
| BTS002 | 4128  | 3680  |

Las coordenadas son métricas (metros). Se admiten `.xlsx`, `.xls`, `.ods` y `.csv`.
Hay un fichero de ejemplo en [`ejemplos/bts_ejemplo.xlsx`](ejemplos/bts_ejemplo.xlsx).

## Algoritmo

1. **Selección de controladores**: se eligen al azar tantos emplazamientos de BTS
   (sin repetición) como controladores se hayan indicado. El controlador se instala
   en ese emplazamiento.
2. **Asignación Greedy por distancia con capacidad**:
   - se calculan las distancias euclídeas de todas las parejas (BTS, controlador);
   - se ordenan de menor a mayor;
   - se recorren en ese orden y se asigna la BTS al controlador si la BTS aún no está
     asignada y el controlador no ha alcanzado su capacidad máxima.

Si `nº controladores × capacidad < nº BTS`, algunas BTS quedan sin asignar y se
marcan con ✕ en el mapa.

## Resultados

- Mapa con las BTS, los controladores (■) y los enlaces coloreados por controlador.
- Modo **paso a paso**: reproduce las asignaciones en el orden en que las hace el
  algoritmo, lo que permite ver cómo un controlador se llena y las BTS siguientes
  "desbordan" a controladores más lejanos.
- Indicadores: BTS asignadas, distancia total, media y máxima.
- Tabla por controlador (ocupación y distancias) y tabla de asignación por BTS.
- Exportación de los resultados a Excel.

## Estructura

```
index.html              interfaz
css/style.css           estilos
js/algorithm.js         lectura de datos, selección aleatoria y algoritmo Greedy
js/app.js               lógica de la interfaz y dibujo del mapa
lib/xlsx.full.min.js    SheetJS 0.18.5 (lectura/escritura de Excel, licencia Apache 2.0)
ejemplos/               fichero Excel de ejemplo
tools/                  generador del fichero de ejemplo
tests/                  pruebas del algoritmo
```

## Desarrollo

Requiere Node.js ≥ 18 solo para las pruebas:

```
npm test            # pruebas del algoritmo
npm run ejemplo     # regenera ejemplos/bts_ejemplo.xlsx
```
