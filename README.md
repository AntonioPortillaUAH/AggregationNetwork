# Red de Agregación

Aplicación docente para demostrar el diseño de la **red de agregación** de una red de
comunicaciones móviles: cómo se conectan las estaciones base (BTS) a un conjunto de
controladores.

## Uso

La aplicación es una página web estática: no necesita instalación ni servidor.

1. Abrir `index.html` en un navegador (Chrome, Firefox, Edge…).
2. Cargar un fichero Excel con las estaciones base (o pulsar *Cargar datos de ejemplo*).
3. Indicar el **número de controladores**, el **máximo de BTS por controlador** y el
   **número de rondas** del Greedy por rondas.
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

## Algoritmos

1. **Selección de controladores**: se eligen al azar tantos emplazamientos de BTS
   (sin repetición) como controladores se hayan indicado. El controlador se instala
   en ese emplazamiento. Los dos métodos usan los mismos controladores, de modo que
   las diferencias entre ellos se deben solo al algoritmo.
2. **Greedy clásico (por distancia con capacidad)**:
   - se calculan las distancias euclídeas de todas las parejas (BTS, controlador);
   - se ordenan de menor a mayor;
   - se recorren en ese orden y se asigna la BTS al controlador si la BTS aún no está
     asignada y el controlador no ha alcanzado su capacidad máxima.
3. **Greedy por rondas**: igual que el anterior, pero el recorrido se repite en
   `R` rondas (parámetro *Número de rondas*). En cada ronda cada controlador solo
   puede recibir `capacidad / R` BTS más; el límite acumulado al final de la ronda
   `r` es `round(capacidad · r / R)`, de modo que en la última ronda es la capacidad.
   Por ejemplo, con capacidad 20 y 5 rondas los límites son 4, 8, 12, 16 y 20.
   Con una sola ronda se obtiene el Greedy clásico. El número de rondas no puede
   superar la capacidad.

Si `nº controladores × capacidad < nº BTS`, algunas BTS quedan sin asignar y se
marcan con ✕ en el mapa.

## Resultados

Tras pulsar **RUN** se ejecutan los dos métodos y los resultados se muestran en
tres pestañas:

- **Greedy clásico** y **Greedy por rondas**: mapa con las BTS, los controladores (■)
  y los enlaces coloreados por controlador; modo **paso a paso** (en el método por
  rondas indica también la ronda de cada asignación); indicadores (BTS asignadas,
  distancia total, media y máxima); tabla por controlador y tabla de asignación por BTS.
- **Comparación**: los indicadores de los dos métodos lado a lado (incluida la
  desviación típica de la distancia y el reparto de BTS entre controladores), el
  número de BTS que cambian de controlador y una tabla por controlador.

El botón **Descargar resultados (.xlsx)** exporta la asignación de ambos métodos,
el resumen por controlador, la comparación y los parámetros.

## Estructura

```
index.html              interfaz
css/style.css           estilos
js/algorithm.js         lectura de datos, selección aleatoria y algoritmos Greedy
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
