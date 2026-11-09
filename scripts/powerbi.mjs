// Lectura de un informe público de Power BI ("Publicar en la web").
// Usa la misma API que el visor de Power BI en el navegador: no necesita cuenta ni clave.
//
//   const pbi = await abrirInforme(url);              // url = https://app.powerbi.com/view?r=...
//   const filas = await consultar(pbi, "Tabla53", ["Comuna", "Estado"]);
//   const links = enlaces(pbi, /realfilename/);       // URLs escritas en cuadros de texto del informe

const UA = { "User-Agent": "planificacion-en-chile (GitHub Actions)" };
const CLUSTERS = [
  "https://wabi-paas-1-scus-api.analysis.windows.net",
  "https://wabi-south-central-us-api.analysis.windows.net",
  "https://wabi-us-north-central-api.analysis.windows.net"
];

/* Decodifica el formato comprimido DSR de Power BI a un arreglo de objetos {columna: valor}. */
export function decodificarDSR(data, nombres) {
  const ds = data.dsr.DS[0];
  const dicts = ds.ValueDicts || {};
  const filas = ds.PH.flatMap(ph => ph.DM0 || []);
  let schema = null, prev = [];
  const out = [];
  for (const f of filas) {
    if (f.S) schema = f.S;
    const vals = [];
    let ci = 0;
    const C = f.C || [], R = f.R || 0, N = f["Ø"] || 0;
    for (let i = 0; i < schema.length; i++) {
      if ((R >> i) & 1) vals[i] = prev[i];
      else if ((N >> i) & 1) vals[i] = null;
      else {
        // Con varias columnas los valores vienen en C; con una sola, como clave G0, G1...
        let v = f.C ? C[ci++] : f[schema[i].N];
        if (schema[i].DN && typeof v === "number") v = dicts[schema[i].DN][v];
        vals[i] = v;
      }
    }
    prev = vals;
    const o = {};
    nombres.forEach((n, i) => { o[n] = vals[i]; });
    out.push(o);
  }
  return { filas: out, continuacion: ds.RT || null, completo: ds.IC !== false };
}

export async function abrirInforme(url) {
  const r = new URL(url).searchParams.get("r");
  const { k: key } = JSON.parse(Buffer.from(r, "base64").toString("utf8"));
  let ultimoError;
  for (const base of CLUSTERS) {
    try {
      const res = await fetch(`${base}/public/reports/${key}/modelsAndExploration?preferReadOnlySession=true`,
        { headers: { ...UA, "X-PowerBI-ResourceKey": key } });
      if (!res.ok) throw new Error(`modelsAndExploration respondió ${res.status}`);
      const texto = await res.text();
      const json = JSON.parse(texto);
      const model = json.models[0];
      return { base, key, modelId: model.id, datasetId: model.dbName, reportId: json.exploration?.report?.objectId, texto };
    } catch (e) { ultimoError = e; }
  }
  throw new Error(`No pude abrir el informe de Power BI: ${ultimoError?.message}`);
}

export async function consultar(pbi, entidad, columnas, maxFilas = 30000) {
  const q = {
    Version: 2,
    From: [{ Name: "t", Entity: entidad, Type: 0 }],
    Select: columnas.map(c => ({ Column: { Expression: { SourceRef: { Source: "t" } }, Property: c }, Name: `${entidad}.${c}` }))
  };
  const body = {
    version: "1.0.0",
    queries: [{
      Query: { Commands: [{ SemanticQueryDataShapeCommand: {
        Query: q,
        Binding: { Primary: { Groupings: [{ Projections: columnas.map((_, i) => i) }] }, DataReduction: { DataVolume: 3, Primary: { Window: { Count: maxFilas } } }, Version: 1 },
        ExecutionMetricsKind: 1
      } }] },
      QueryId: "",
      ApplicationContext: { DatasetId: pbi.datasetId, Sources: [{ ReportId: pbi.reportId }] }
    }],
    cancelQueries: [],
    modelId: pbi.modelId
  };
  const res = await fetch(`${pbi.base}/public/reports/querydata?synchronous=true`, {
    method: "POST",
    headers: { ...UA, "X-PowerBI-ResourceKey": pbi.key, "Content-Type": "application/json" },
    body: JSON.stringify(body)
  });
  if (!res.ok) throw new Error(`querydata respondió ${res.status}`);
  const json = await res.json();
  const data = json.results?.[0]?.result?.data;
  if (!data?.dsr) throw new Error(`Power BI no devolvió datos para ${entidad}: ${JSON.stringify(json).slice(0, 300)}`);
  if (data.dsr.DataShapes?.[0]?.["odata.error"]) throw new Error(`Power BI: ${JSON.stringify(data.dsr.DataShapes[0]["odata.error"]).slice(0, 300)}`);
  return decodificarDSR(data, columnas).filas;
}

/* URLs (sin el envoltorio de Outlook SafeLinks) que aparecen en el informe y cumplen el patrón. */
export function enlaces(pbi, patron) {
  const out = new Set();
  for (const m of pbi.texto.matchAll(/https?:\/\/[^"\\\s]+/g)) {
    let u = m[0];
    try {
      const x = new URL(u);
      if (/safelinks\.protection\.outlook\.com$/.test(x.hostname) && x.searchParams.get("url")) u = x.searchParams.get("url");
    } catch { continue; }
    if (patron.test(u)) out.add(u);
  }
  return [...out];
}

/* Tablas y columnas del modelo: [{entidad, columnas:[...]}] */
export async function esquema(pbi) {
  const res = await fetch(`${pbi.base}/public/reports/conceptualschema`, {
    method: "POST",
    headers: { ...UA, "X-PowerBI-ResourceKey": pbi.key, "Content-Type": "application/json" },
    body: JSON.stringify({ modelIds: [pbi.modelId], userPreferredLocale: "es-ES" })
  });
  if (!res.ok) throw new Error(`conceptualschema respondió ${res.status}`);
  const json = await res.json();
  return json.schemas[0].schema.Entities.map(e => ({ entidad: e.Name, columnas: (e.Properties || []).map(p => p.Name) }));
}
