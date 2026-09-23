# Planificación en Chile

Visor web del estado de tres instrumentos de planificación en las 345 comunas de Chile:

- **PACCC**: Plan de Acción Comunal de Cambio Climático (Ministerio del Medio Ambiente).
- **Plan comunal RRD**: Plan Comunal para la Reducción del Riesgo de Desastres (SENAPRED).
- **PRC**: Plan Regulador Comunal (MINVU, instrumentos de planificación territorial).

Es un sitio estático: HTML, CSS y JavaScript, con Leaflet y mapa base gris de Esri (sin API key).
No necesita servidor ni base de datos. El navegador lee los archivos originales de las fuentes desde `data/` y los procesa al cargar la página.

## Estructura

```
index.html            página del visor
assets/app.js         lectura de fuentes, clasificación y mapa
assets/style.css      estilos
data/fuentes.json     qué archivo usar para cada fuente, y su fecha
data/comunas.geojson  límites comunales (desde COMUNAS_v1.shp, simplificado)
data/*.csv, *.xlsx    descargas originales de cada fuente
```

## Publicar en GitHub Pages

1. Crea un repositorio nuevo en GitHub (por ejemplo `planificacion-en-chile`).
2. Sube todo el contenido de esta carpeta, incluido el archivo oculto `.nojekyll`.
   Desde la web: *Add file → Upload files* y arrastra los archivos y carpetas.
3. En *Settings → Pages*, elige *Deploy from a branch*, rama `main`, carpeta `/ (root)`.
4. En uno o dos minutos el visor queda en `https://<usuario>.github.io/planificacion-en-chile/`.

## Actualizar los datos

1. Descarga el archivo nuevo desde la fuente, sin modificarlo.
2. Súbelo a la carpeta `data/` del repositorio (*Add file → Upload files*).
3. Edita `data/fuentes.json` con el lápiz de GitHub: cambia `archivo` por el nombre del archivo nuevo y `fecha` por la fecha de corte (`AAAA-MM-DD`).
   Si dejas `fecha` vacía, el visor la toma del nombre del CSV de PACCC o del título del Excel RRD.
4. Guarda el cambio (*Commit changes*). GitHub Pages se actualiza solo.
5. Opcional: borra el archivo antiguo de `data/` para no acumular versiones.

Antes de subir, puedes revisar las descargas nuevas con el botón **Probar archivos nuevos** del visor.
Te dice cuántas filas leyó, si alguna comuna no se pudo enlazar y te deja ver el resultado en el mapa sin publicar nada.

### Formatos que reconoce

| Fuente | Formato | Columnas que usa |
|---|---|---|
| PACCC | CSV (`;` o `,`) | `cut_com`, `comuna`, `estado`, `plan_url`, `decreto_url` |
| Plan comunal RRD | Excel | `Región`, `Comuna`, `Estado`, `Año`, `N° Decreto Alcaldicio`, `Estado de Revisión por Informe Técnico`, `FINANCIAMIENTO` |
| PRC | CSV (`;`) | `Comunas`, `Tipo de planificación`, `Denominación`, `Estado`, `Fecha de inicio de vigencia` |

Si una fuente cambia el nombre de una comuna (por ejemplo "Coyhaique" y "Coihaique"), agrega la equivalencia en el objeto `ALIAS` de `assets/app.js`.

## Criterios de clasificación

**PACCC**: aprobado, en proceso o sin información, según la columna `estado`.

**Plan comunal RRD**: con decreto alcaldicio. Si no tiene decreto, se clasifica según la revisión técnica: recomendado y en formalización, en revisión, no recomendado o sin plan presentado.

**PRC**: se considera el instrumento de origen vigente más reciente de la comuna.
Está al día si tiene 10 años o menos (constante `PRC_MAX_AGE` en `assets/app.js`).
Si es más antiguo, se distingue si hay otro PRC "En Desarrollo".
El archivo de instrumentos no incluye modificaciones ni enmiendas, así que la antigüedad corresponde al instrumento de origen.

**Los tres**: cuenta cuántos instrumentos tiene al día cada comuna (PACCC aprobado, plan RRD con decreto, PRC vigente). Una casilla permite contar cualquier PRC vigente, sin importar su antigüedad.

## Probar en tu computador

El visor lee archivos con `fetch`, así que no funciona abriendo `index.html` con doble clic. En la carpeta del proyecto:

```
python -m http.server
```

y abre `http://localhost:8000`.

## Regenerar la geometría

`data/comunas.geojson` se generó desde `COMUNAS_v1.shp` (SIRGAS-Chile) con [mapshaper](https://mapshaper.org):

```
mapshaper -i COMUNAS_v1.shp encoding=utf8 \
  -simplify 0.7% keep-shapes -filter-islands min-area=1km2 remove-empty \
  -each 'cut=+CUT_COM' -filter-fields cut,COMUNA -rename-fields comuna=COMUNA \
  -proj wgs84 -o format=geojson precision=0.0005 comunas.geojson
```

El visor solo necesita las propiedades `cut` (código comunal numérico) y `comuna`.

## Créditos

Diseño: Pablo Sarricolea, Departamento de Geografía, Universidad de Chile, y Centro de Ciencia del Clima y la Resiliencia (CR2).

Mapa base © Esri, HERE, Garmin, © OpenStreetMap contributors. Datos: Ministerio del Medio Ambiente, SENAPRED y MINVU.
