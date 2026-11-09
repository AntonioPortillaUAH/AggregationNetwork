# Red de Agregación

Aplicación docente para demostrar el diseño de la **red de agregación** de una red de
comunicaciones móviles: cómo se conectan las estaciones base (BTS) a un conjunto de
controladores.

## Uso

La aplicación es una página web estática: no necesita instalación ni servidor.

1. Abrir `index.html` en un navegador (Chrome, Firefox, Edge…).
2. Cargar un fichero Excel con las estaciones base (o pulsar *Cargar datos de ejemplo*).
3. Indicar el **número de controladores**, el **máximo de BTS por controlador**, la
   **distancia mínima entre controladores** (en km; 0 = sin restricción) y el
   **número de rondas** de los métodos por rondas.
4. Opcionalmente, indicar una **semilla** para que la selección aleatoria sea repetible
   (útil para que todos los alumnos obtengan el mismo resultado).
5. Pulsar **RUN**.

Los controladores se eligen al azar en el primer **RUN** y **se mantienen fijos** en
las siguientes ejecuciones, de modo que se pueden cambiar la capacidad o el número de
rondas y comparar resultados con los mismos controladores. Para sortear otros, se
pulsa **Nueva selección de controladores**. También se eligen nuevos automáticamente
si cambian el número de controladores, la distancia mínima o la semilla. Con semilla,
cada selección está numerada y es reproducible: la selección nº *k* con la semilla *S*
es siempre la misma.

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
   en ese emplazamiento. Si se indica una **distancia mínima entre controladores**,
   la selección debe cumplir que cualquier pareja de controladores esté al menos a esa
   distancia: primero se prueban órdenes aleatorios de las BTS y, si ninguno lo
   consigue, se hace una búsqueda exhaustiva con vuelta atrás. Si no existe ninguna
   selección posible, la aplicación lo indica y **no se ejecutan los algoritmos de
   asignación**. Todos los métodos usan los mismos controladores, de modo que las
   diferencias entre ellos se deben solo al algoritmo.
Hay cuatro métodos de asignación, todos con la capacidad máxima de cada controlador:

2. **Greedy ordenado**:
   - se calculan las distancias euclídeas de todas las parejas (BTS, controlador);
   - se ordenan de menor a mayor, sin agruparlas por controlador;
   - se recorren en ese orden y se asigna la BTS al controlador si la BTS aún no está
     asignada y el controlador no ha alcanzado su capacidad máxima.
3. **Greedy ordenado por rondas**: igual que el anterior, pero el recorrido se repite en
   `R` rondas (parámetro *Número de rondas*). En cada ronda cada controlador solo
   puede recibir `capacidad / R` BTS más; el límite acumulado al final de la ronda
   `r` es `round(capacidad · r / R)`, de modo que en la última ronda es la capacidad.
   Por ejemplo, con capacidad 20 y 5 rondas los límites son 4, 8, 12, 16 y 20.
   Con una sola ronda se obtiene el Greedy ordenado. El número de rondas no puede
   superar la capacidad.
4. **Greedy secuencial por controlador**: los controladores se atienden por orden
   (C1, C2, …). Cada uno toma sus BTS libres más cercanas hasta llenarse y después se
   pasa al siguiente. El resultado depende del orden de los controladores: los
   primeros pueden quedarse con BTS que están más cerca de otro controlador. La BTS
   donde está instalado cada controlador queda reservada para él.
5. **Greedy secuencial por controlador por rondas**: igual que el anterior, pero con
   los mismos límites por ronda que el Greedy ordenado por rondas: en cada ronda C1,
   C2, … toman por turno sus BTS libres más cercanas hasta el límite de la ronda.

Si `nº controladores × capacidad < nº BTS`, algunas BTS quedan sin asignar y se
marcan con ✕ en el mapa.

## Resultados

Tras pulsar **RUN** se ejecutan los cuatro métodos y los resultados se muestran en
cinco pestañas:

- **Greedy ordenado**, **Ordenado por rondas**, **Secuencial por controlador** y
  **Secuencial por rondas**: mapa con las BTS, los controladores (■)
  y los enlaces coloreados por controlador; modo **paso a paso** (en el método por
  rondas indica también la ronda de cada asignación); indicadores (BTS asignadas,
  distancia total, media y máxima); tabla por controlador y tabla de asignación por BTS.
- En las pestañas de los métodos **por rondas** se puede ver el **resultado de cada ronda**: los
  botones *Ronda 1 … R* del panel *Paso a paso* (o las filas de la tabla *Resultado
  de cada ronda*) muestran el mapa al terminar esa ronda. Las BTS que aún no tienen
  controlador se ven apagadas y los enlaces de rondas anteriores, más tenues. La tabla
  resume, para cada ronda, el límite por controlador, las BTS asignadas en la ronda y
  acumuladas, las pendientes y la distancia media y máxima de la ronda.
- **Comparación**: los indicadores de los cuatro métodos lado a lado (incluida la
  desviación típica de la distancia y el reparto de BTS entre controladores), el
  número de BTS con distinto controlador que en el Greedy ordenado y una tabla por
  controlador.

El botón **Descargar resultados (.xlsx)** exporta la asignación de los cuatro métodos,
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
