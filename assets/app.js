/* Visor de planificación comunal en Chile
   Lee las fuentes originales desde data/ (según data/fuentes.json) y las procesa en el navegador. */
(function () {
"use strict";

const COL = {good:"#1B6A8A", goodL:"#7DB5C4", mid:"#E3A82B", badL:"#E08D78", bad:"#A63D55", none:"#BCC3C5"};
const REGIONS = [[15,"Arica y Parinacota"],[1,"Tarapacá"],[2,"Antofagasta"],[3,"Atacama"],[4,"Coquimbo"],[5,"Valparaíso"],[13,"Metropolitana"],[6,"O'Higgins"],[7,"Maule"],[16,"Ñuble"],[8,"Biobío"],[9,"La Araucanía"],[14,"Los Ríos"],[10,"Los Lagos"],[11,"Aysén"],[12,"Magallanes"]];
const REG_NAME = Object.fromEntries(REGIONS);
const ISLANDS = new Set([5104, 5201]);            // Juan Fernández, Isla de Pascua
const PRC_MAX_AGE = 10;                            // años para considerar un PRC "al día"

const LAYERS = {
  paccc: {title:"Plan de Acción Comunal de Cambio Climático", good:"aprobado", metric:"Porcentaje de comunas con PACCC aprobado.",
    cats:[{k:"aprobado",l:"Aprobado",c:COL.good},{k:"proceso",l:"En proceso",c:COL.mid},{k:"sininfo",l:"Sin información",c:COL.none}]},
  rrd: {title:"Plan Comunal para la Reducción del Riesgo de Desastres", good:"decreto", metric:"Porcentaje de comunas con plan RRD aprobado por decreto alcaldicio.",
    cats:[{k:"decreto",l:"Con decreto alcaldicio",c:COL.good},{k:"formalizacion",l:"Recomendado, en formalización",c:COL.goodL},{k:"revision",l:"En revisión técnica",c:COL.mid},{k:"norecomendado",l:"No recomendado técnicamente",c:COL.badL},{k:"sinplan",l:"Sin plan presentado",c:COL.bad},{k:"sininfo",l:"Sin información",c:COL.none}]},
  prc: {title:"Plan Regulador Comunal", good:"actual", metric:`Porcentaje de comunas con un PRC vigente de ${PRC_MAX_AGE} años o menos.`,
    cats:[{k:"actual",l:`Vigente de ${PRC_MAX_AGE} años o menos`,c:COL.good},{k:"actualizando",l:`Más de ${PRC_MAX_AGE} años, en actualización`,c:COL.goodL},{k:"antiguo",l:`Más de ${PRC_MAX_AGE} años, sin actualización`,c:COL.mid},{k:"elaborando",l:"Sin PRC vigente, en elaboración",c:COL.badL},{k:"sinprc",l:"Sin PRC vigente",c:COL.bad}]},
  all: {title:"Instrumentos al día", good:"3", metric:"Porcentaje de comunas con los tres instrumentos al día.",
    cats:[{k:"3",l:"Los tres",c:COL.good},{k:"2",l:"Dos de tres",c:COL.goodL},{k:"1",l:"Uno de tres",c:COL.mid},{k:"0",l:"Ninguno",c:COL.bad}]}
};

/* Nombres que difieren entre fuentes. Agrega aquí si una actualización reporta comunas sin enlazar. */
const ALIAS = {CABODEHORNOSYANTARTICACHILENA:"CABODEHORNOS", COYHAIQUE:"COIHAIQUE", ELOLIVAR:"OLIVAR", LACALERA:"CALERA", MARCHIGUE:"MARCHIHUE", PAIHUANO:"PAIGUANO", SANVICENTETT:"SANVICENTE", SANVICENTEDETAGUATAGUA:"SANVICENTE", AISEN:"AYSEN", PUERTOAYSEN:"AYSEN", LLAILLAY:"LLAYLLAY", TREGUACO:"TREHUACO"};

const norm = s => String(s == null ? "" : s).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase().replace(/[^A-Z]/g, "");
const esc = s => String(s == null ? "" : s).replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const pdate = s => { const m = String(s || "").match(/(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/); return m ? new Date(+m[3], +m[2]-1, +m[1]) : null; };
const iso = s => { const m = String(s || "").match(/(\d{4})-(\d{2})-(\d{2})/); return m ? new Date(+m[1], +m[2]-1, +m[3]) : null; };
const toIso = d => `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
const fdate = d => d ? d.toLocaleDateString("es-CL", {day:"numeric", month:"short", year:"numeric"}) : "sin fecha";
const pct = (a, b) => b ? Math.round(100 * a / b) : 0;
const cssVar = n => getComputedStyle(document.documentElement).getPropertyValue(n).trim();
const $ = id => document.getElementById(id);

/* ---------- Lectura de archivos fuente ---------- */
function decodeText(buf) {
  let t = new TextDecoder("utf-8").decode(buf);
  if (t.includes("\ufffd")) t = new TextDecoder("windows-1252").decode(buf);
  return t.replace(/^\ufeff/, "");
}
function pick(row, re) { for (const k in row) if (re.test(norm(k))) return row[k]; return ""; }
function need(fields, list, label) {
  for (const [n, re] of list) if (!fields.some(f => re.test(norm(f)))) throw new Error(`Al archivo de ${label} le falta la columna “${n}”.`);
}
/* Reconoce el archivo por su contenido y devuelve {kind, rows, fecha} */
function parseSource(name, buf, fallbackDate) {
  if (/\.xlsx?$/i.test(name)) {
    const wb = XLSX.read(buf, {type:"array"});
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {header:1, defval:"", raw:false});
    const hi = rows.findIndex(r => r.some(v => norm(v) === "COMUNA"));
    if (hi < 0) throw new Error("No encontré una fila de encabezado con la columna “Comuna”.");
    const H = rows[hi].map(norm), col = f => H.findIndex(f);
    const ix = {region:col(h=>h.startsWith("REGION")), comuna:col(h=>h==="COMUNA"), estado:col(h=>h==="ESTADO"), anio:col(h=>h.startsWith("ANO")),
                decreto:col(h=>h.includes("DECRETO")), revision:col(h=>h.includes("REVISION")), financiamiento:col(h=>h.includes("FINANC"))};
    if (ix.estado < 0 || ix.revision < 0) throw new Error("El Excel no tiene las columnas “Estado” y “Estado de Revisión” del formato RRD.");
    let fecha = null;
    for (const r of rows.slice(0, hi)) { for (const v of r) { const d = pdate(v); if (d) { fecha = d; break; } } if (fecha) break; }
    const out = [];
    for (const r of rows.slice(hi + 1)) {
      const o = {}; for (const k in ix) o[k] = ix[k] >= 0 ? String(r[ix[k]] == null ? "" : r[ix[k]]).trim() : "";
      if (o.comuna) out.push(o);
    }
    return {kind:"rrd", rows:out, fecha:fecha || fallbackDate};
  }
  const p = Papa.parse(decodeText(buf), {header:true, skipEmptyLines:true, delimiter:"", transformHeader:h => h.replace(/^\ufeff/, "").trim()});
  const f = p.meta.fields || [], fn = f.map(norm);
  if (fn.includes("CUTCOM")) {
    need(f, [["cut_com",/^CUTCOM$/],["comuna",/^COMUNA$/],["estado",/^ESTADO$/]], "PACCC");
    const rows = p.data.map(r => ({cut_com:String(pick(r,/^CUTCOM$/)).trim(), comuna:pick(r,/^COMUNA$/), region:pick(r,/^REGION$/), estado:pick(r,/^ESTADO$/),
                                   plan_url:pick(r,/^PLANURL$/), decreto_url:pick(r,/^DECRETOURL$/)})).filter(r => r.cut_com);
    const m = name.match(/(\d{4}-\d{2}-\d{2})/);
    return {kind:"paccc", rows, fecha:m ? iso(m[1]) : fallbackDate};
  }
  if (fn.some(h => h.startsWith("TIPODEPLANIFICACION")) || fn.includes("DENOMINACION")) {
    need(f, [["Comunas",/^COMUNAS$/],["Denominación",/^DENOMINACION$/],["Estado",/^ESTADO$/],["Fecha de inicio de vigencia",/^FECHADEINICIODEVIGENCIA$/]], "instrumentos");
    const tipo = f.find(h => norm(h).startsWith("TIPODEPLANIFICACION"));
    const rows = p.data.filter(r => !tipo || norm(r[tipo]) === "PRC").map(r => ({comunas:pick(r,/^COMUNAS$/), nombre:pick(r,/^DENOMINACION$/), estado:pick(r,/^ESTADO$/),
                 vigencia:pick(r,/^FECHADEINICIODEVIGENCIA$/), hito:pick(r,/^FECHADEULTIMOHITOCUMPLIDO$/)}));
    return {kind:"prc", rows, fecha:fallbackDate};
  }
  throw new Error("No reconocí el formato: se esperaba el CSV de PACCC (columna cut_com), el de instrumentos (columna Denominación) o el Excel RRD.");
}

/* ---------- De filas fuente a estado por comuna ---------- */
let GEO = null, RAW = null, D = null;
function process(raw) {
  const C = new Map(), idx = new Map();
  for (const f of GEO.features) {
    const cod = +f.properties.cut;
    C.set(cod, {cod, name:f.properties.comuna, reg:Math.floor(cod / 1000), prcRows:[]});
    idx.set(norm(f.properties.comuna), cod);
  }
  for (const r of raw.paccc) { const c = C.get(+r.cut_com); if (c && r.comuna) idx.set(norm(r.comuna), c.cod); }
  const find = n => { const k = norm(n); return idx.has(k) ? idx.get(k) : idx.get(ALIAS[k]); };
  const miss = {paccc:new Set(), rrd:new Set(), prc:new Set()};
  for (const r of raw.paccc) { const c = C.get(+r.cut_com); if (!c) { miss.paccc.add(r.comuna || r.cut_com); continue; } c.pacccRow = r; }
  for (const r of raw.rrd) { const cod = find(r.comuna); if (cod == null) { miss.rrd.add(r.comuna); continue; } C.get(cod).rrdRow = r; }
  for (const r of raw.prc) for (const n of String(r.comunas || "").split(",")) {
    const t = n.trim(); if (!t) continue;
    const cod = find(t); if (cod == null) { miss.prc.add(t); continue; }
    C.get(cod).prcRows.push(r);
  }
  const now = new Date();
  for (const c of C.values()) {
    const e = norm(c.pacccRow && c.pacccRow.estado);
    c.paccc = e === "APROBADO" ? "aprobado" : (e.startsWith("ENPROCESO") || e.startsWith("ENELABORACION")) ? "proceso" : "sininfo";

    const r = c.rrdRow;
    if (!r) c.rrd = "sininfo";
    else {
      const est = norm(r.estado), rev = norm(r.revision);
      c.rrd = est.includes("DECRETO") ? "decreto"
            : rev.includes("NORECOMENDADO") ? "norecomendado"
            : rev.includes("RECOMENDADO") ? "formalizacion"
            : rev.includes("ENREVISION") ? "revision" : "sinplan";
    }

    const vig = c.prcRows.filter(x => norm(x.estado) === "VIGENTE"), dev = c.prcRows.filter(x => norm(x.estado) === "ENDESARROLLO");
    let latest = null; for (const v of vig) { const d = pdate(v.vigencia); if (d && (!latest || d > latest)) latest = d; }
    c.prcLatest = latest; c.prcAge = latest ? (now - latest) / 31557600000 : null; c.prcVig = vig.length; c.prcDev = dev.length;
    if (vig.length) c.prc = (c.prcAge != null && c.prcAge <= PRC_MAX_AGE) ? "actual" : dev.length ? "actualizando" : "antiguo";
    else c.prc = dev.length ? "elaborando" : "sinprc";
  }
  return {C, miss};
}

const state = {layer:"paccc", focus:null, sel:null, reg:null, strict:true};
const score = c => (c.paccc === "aprobado") + (c.rrd === "decreto") + (state.strict ? c.prc === "actual" : c.prcVig > 0);
const catOf = c => state.layer === "all" ? String(score(c)) : c[state.layer];
const catInfo = (layer, k) => LAYERS[layer].cats.find(x => x.k === k) || {l:k, c:COL.none};

/* ---------- Mapa ---------- */
const CONT = [[-55.9,-75.9],[-17.45,-66.3]];
const map = L.map("map", {preferCanvas:true, zoomSnap:0.25, minZoom:3, maxZoom:13});
map.attributionControl.setPrefix(false);
map.createPane("labels"); map.getPane("labels").style.zIndex = 650; map.getPane("labels").style.pointerEvents = "none";
const dark = window.matchMedia("(prefers-color-scheme: dark)");
const ATTR = 'Mapa base &copy; Esri, HERE, Garmin, &copy; colaboradores de OpenStreetMap';
let base, labels;
function setBase() {
  if (base) map.removeLayer(base); if (labels) map.removeLayer(labels);
  const s = dark.matches ? "Dark" : "Light";
  const esri = n => `https://server.arcgisonline.com/ArcGIS/rest/services/Canvas/World_${s}_Gray_${n}/MapServer/tile/{z}/{y}/{x}`;
  base = L.tileLayer(esri("Base"), {maxZoom:16, attribution:ATTR}).addTo(map);
  labels = L.tileLayer(esri("Reference"), {maxZoom:16, pane:"labels"}).addTo(map);
}
setBase();
map.fitBounds(CONT);

const layerOf = new Map();
let geoLayer = null;
function styleFor(cod) {
  const c = D.C.get(cod), cat = catOf(c), info = catInfo(state.layer, cat);
  const dim = state.focus != null && cat !== state.focus, sel = state.sel === cod;
  return {fillColor:info.c, fillOpacity:dim ? 0.08 : 0.85, color:sel ? cssVar("--ink") : cssVar("--edge"), weight:sel ? 2.6 : 0.5, opacity:dim && !sel ? 0.4 : 1};
}
function drawGeo() {
  geoLayer = L.geoJSON(GEO, {
    style: f => styleFor(+f.properties.cut),
    onEachFeature: (f, l) => {
      const cod = +f.properties.cut; layerOf.set(cod, l);
      l.bindTooltip(() => { const c = D.C.get(cod); return `<b>${esc(c.name)}</b><span>${esc(catInfo(state.layer, catOf(c)).l)}</span>`; },
                    {sticky:true, className:"tip", direction:"top", offset:[0,-8]});
      l.on("mouseover", () => l.setStyle({weight:1.8, color:cssVar("--ink")}));
      l.on("mouseout", () => l.setStyle(styleFor(cod)));
      l.on("click", () => select(cod, false));
    }
  }).addTo(map);
}
function restyle() {
  if (!geoLayer) return;
  geoLayer.eachLayer(l => l.setStyle(styleFor(+l.feature.properties.cut)));
  if (state.sel && layerOf.get(state.sel)) layerOf.get(state.sel).bringToFront();
}
dark.addEventListener("change", () => { setBase(); restyle(); });
document.querySelectorAll("[data-jump]").forEach(b => b.addEventListener("click", () => {
  const j = b.dataset.jump;
  if (j === "cont") map.fitBounds(CONT);
  else if (layerOf.size) map.fitBounds(layerOf.get(j === "rn" ? 5201 : 5104).getBounds(), {padding:[30,30]});
}));

/* ---------- Leyenda ---------- */
function renderLegend() {
  const Ly = LAYERS[state.layer];
  const list = [...D.C.values()].filter(c => !state.reg || c.reg === state.reg);
  const counts = {}; for (const c of list) { const k = catOf(c); counts[k] = (counts[k] || 0) + 1; }
  let h = `<h3>${esc(Ly.title)} <span class="n" style="font-weight:400">en ${esc(state.reg ? REG_NAME[state.reg] : "Chile")}</span></h3><div class="items">`;
  for (const k of Ly.cats) {
    const on = state.focus === k.k, dim = state.focus != null && !on;
    h += `<button data-cat="${k.k}" aria-pressed="${on}" class="${dim ? "dim" : ""}"><span class="sw" style="background:${k.c}"></span><span>${esc(k.l)}</span><span class="n">${counts[k.k] || 0}</span></button>`;
  }
  h += "</div>";
  if (state.layer === "all") h += `<p class="hint">Al día: PACCC aprobado, plan RRD con decreto y PRC vigente.</p><label><input type="checkbox" id="strict" ${state.strict ? "checked" : ""}> Exigir que el PRC tenga ${PRC_MAX_AGE} años o menos</label>`;
  else if (state.layer === "prc") h += `<p class="hint">La antigüedad se mide desde el último instrumento de origen vigente.</p>`;
  h += `<p class="hint">${state.focus ? "Toca la categoría otra vez para ver todas." : "Toca una categoría para destacarla."}</p>`;
  const el = $("legend"); el.innerHTML = h;
  el.querySelectorAll("[data-cat]").forEach(b => b.addEventListener("click", () => { state.focus = state.focus === b.dataset.cat ? null : b.dataset.cat; restyle(); renderLegend(); }));
  const s = el.querySelector("#strict"); if (s) s.addEventListener("change", () => { state.strict = s.checked; renderAll(); });
}

/* ---------- Panel regional ---------- */
function rowHtml(code, name, list, nation) {
  const Ly = LAYERS[state.layer], n = list.length, counts = {};
  for (const c of list) { const k = catOf(c); counts[k] = (counts[k] || 0) + 1; }
  const bar = Ly.cats.map(k => counts[k.k] ? `<i style="width:${100 * counts[k.k] / n}%;background:${k.c}" title="${esc(k.l)}: ${counts[k.k]}"></i>` : "").join("");
  return `<button class="reg${nation ? " nation" : ""}" data-reg="${code}" aria-pressed="${nation ? state.reg == null : state.reg === code}">
    <span class="nm">${esc(name)}<small>${n} comunas</small></span><span class="pct">${pct(counts[Ly.good] || 0, n)}%</span><span class="bar">${bar}</span></button>`;
}
function renderRegions() {
  const all = [...D.C.values()];
  let h = rowHtml(0, "Chile", all, true);
  for (const [code, name] of REGIONS) h += rowHtml(code, name, all.filter(c => c.reg === code), false);
  const el = $("regs"); el.innerHTML = h;
  $("metric").textContent = LAYERS[state.layer].metric + " La barra muestra todas las categorías.";
  el.querySelectorAll("[data-reg]").forEach(b => b.addEventListener("click", () => {
    const code = +b.dataset.reg;
    if (!code || state.reg === code) { state.reg = null; map.fitBounds(CONT); }
    else {
      state.reg = code; let bb = null;
      for (const c of D.C.values()) {
        if (c.reg !== code || ISLANDS.has(c.cod)) continue;
        const lb = layerOf.get(c.cod).getBounds();
        bb = bb ? bb.extend(lb) : L.latLngBounds(lb.getSouthWest(), lb.getNorthEast());
      }
      if (bb) map.fitBounds(bb, {padding:[20,20]});
    }
    renderRegions(); renderLegend();
  }));
}

/* ---------- Ficha comunal ---------- */
const chip = (layer, k) => { const i = catInfo(layer, k); return `<span class="chip"><span class="sw" style="background:${i.c}"></span>${esc(i.l)}</span>`; };
const link = (u, t) => /^https?:\/\//.test(u || "") ? `<a href="${esc(u)}" target="_blank" rel="noopener">${t}</a>` : "";
function renderDetail() {
  const el = $("detail");
  if (!state.sel) { el.hidden = true; el.innerHTML = ""; return; }
  const c = D.C.get(state.sel), s = score(c), p = c.pacccRow || {}, r = c.rrdRow;
  const rows = c.prcRows.slice().sort((a, b) => (pdate(b.vigencia) || 0) - (pdate(a.vigencia) || 0));
  const links = [link(p.plan_url, "Ver plan"), link(p.decreto_url, "Ver decreto")].filter(Boolean).join("");
  const val = v => v && v !== "-";
  el.innerHTML = `
  <div class="det-head"><div><h2>${esc(c.name)}</h2><p>Región de ${esc(REG_NAME[c.reg] || "")}, código CUT ${c.cod}</p></div><button class="x" id="closeDet" aria-label="Cerrar ficha">×</button></div>
  <div class="score" aria-label="${s} de 3 instrumentos al día">${[0,1,2].map(i => `<i class="${i < s ? "on" : ""}"></i>`).join("")}<span>${s} de 3 instrumentos al día</span></div>
  <div class="inst"><h3>Plan de Acción Comunal de Cambio Climático</h3>${chip("paccc", c.paccc)}${links ? `<div class="links">${links}</div>` : ""}</div>
  <div class="inst"><h3>Plan Comunal para la Reducción del Riesgo de Desastres</h3>${chip("rrd", c.rrd)}
    ${r ? `<dl>${val(r.anio) ? `<dt>Año</dt><dd>${esc(r.anio)}</dd>` : ""}${val(r.decreto) ? `<dt>Decreto</dt><dd>N° ${esc(r.decreto)}</dd>` : ""}<dt>Revisión técnica</dt><dd>${esc(r.revision || "Sin dato")}</dd>${r.financiamiento ? `<dt>Financiamiento</dt><dd>${esc(r.financiamiento)}</dd>` : ""}</dl>` : ""}</div>
  <div class="inst"><h3>Plan Regulador Comunal</h3>${chip("prc", c.prc)}
    ${c.prcLatest ? `<dl><dt>Último vigente</dt><dd>${fdate(c.prcLatest)} (${Math.floor(c.prcAge)} años)</dd></dl>` : ""}
    ${rows.length ? `<ul class="plist">${rows.map(x => `<li><span>${esc(x.nombre)}</span><span class="st ${norm(x.estado) === "VIGENTE" ? "v" : ""}">${esc(x.estado)}${x.vigencia ? ", " + esc(String(x.vigencia).slice(-4)) : ""}</span></li>`).join("")}</ul>`
                  : `<p class="lead" style="margin:6px 0 0">No hay instrumentos registrados.</p>`}</div>`;
  el.hidden = false;
  el.querySelector("#closeDet").addEventListener("click", () => { state.sel = null; restyle(); renderDetail(); });
}
function select(cod, zoom) {
  state.sel = cod; restyle(); renderDetail();
  if (zoom) map.fitBounds(layerOf.get(cod).getBounds(), {padding:[60,60], maxZoom:11});
  if (window.matchMedia("(max-width:860px)").matches) $("detail").scrollIntoView({behavior:"smooth", block:"start"});
  else document.querySelector(".side").scrollTop = 0;
}

/* ---------- Búsqueda ---------- */
function renderSearch() {
  const list = [...D.C.values()].sort((a, b) => a.name.localeCompare(b.name, "es"));
  $("comunas").innerHTML = list.map(c => `<option value="${esc(c.name)} (${esc(REG_NAME[c.reg] || "")})">`).join("");
}
const q = $("q");
function doSearch() {
  const k = norm(q.value.replace(/\s*\(.*\)\s*$/, "")); if (!k) return;
  const all = [...D.C.values()];
  const hit = all.find(c => norm(c.name) === k) || all.find(c => norm(c.name).startsWith(k));
  if (hit) { select(hit.cod, true); q.blur(); }
}
q.addEventListener("change", doSearch);
q.addEventListener("keydown", e => { if (e.key === "Enter") doSearch(); });

/* ---------- Fuentes ---------- */
function renderSources() {
  $("sources").innerHTML = ["paccc", "rrd", "prc"].map(k => {
    const m = RAW.meta[k];
    const inst = m.url ? `<a href="${esc(m.url)}" target="_blank" rel="noopener">${esc(m.institucion)}</a>` : esc(m.institucion || "");
    return `<li><b>${esc(m.nombre)}</b><br>${inst}. Datos al ${fdate(iso(m.fecha))}, archivo ${esc(m.archivo)}${m.local ? " (vista de prueba, no publicado)" : ""}.</li>`;
  }).join("");
}

/* ---------- Capa activa ---------- */
document.querySelectorAll("[data-layer]").forEach(b => b.addEventListener("click", () => {
  state.layer = b.dataset.layer; state.focus = null;
  document.querySelectorAll("[data-layer]").forEach(x => x.setAttribute("aria-pressed", x === b));
  renderAll();
}));
function renderAll() { restyle(); renderLegend(); renderRegions(); renderDetail(); }

/* ---------- Exportar CSV ---------- */
$("csv").addEventListener("click", () => {
  const head = ["cut_com","comuna","region","paccc","rrd","rrd_anio","rrd_decreto","rrd_revision","rrd_financiamiento","prc","prc_ultimo_vigente","prc_antiguedad_anios","prc_vigentes","prc_en_desarrollo","instrumentos_al_dia"];
  const qv = v => { v = String(v == null ? "" : v); return /[;"\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; };
  const ord = c => REGIONS.findIndex(r => r[0] === c.reg);
  const lines = [head.join(";")];
  for (const c of [...D.C.values()].sort((a, b) => ord(a) - ord(b) || a.cod - b.cod)) {
    const r = c.rrdRow || {};
    lines.push([c.cod, c.name, REG_NAME[c.reg], catInfo("paccc", c.paccc).l, catInfo("rrd", c.rrd).l, r.anio, r.decreto, r.revision, r.financiamiento,
      catInfo("prc", c.prc).l, c.prcLatest ? toIso(c.prcLatest) : "", c.prcAge != null ? c.prcAge.toFixed(1) : "", c.prcVig, c.prcDev, score(c)].map(qv).join(";"));
  }
  const a = document.createElement("a");
  a.href = URL.createObjectURL(new Blob(["\ufeff" + lines.join("\n")], {type:"text/csv"}));
  a.download = "estado-planificacion-comunal.csv"; document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
});

/* ---------- Probar archivos nuevos (solo local) ---------- */
const dlg = $("upd"), resEl = $("res"), btnLocal = $("applyLocal");
const LABEL = {paccc:"PACCC", rrd:"Planes comunales RRD", prc:"Planes reguladores"};
let pending = null;
async function handleFiles(files) {
  resEl.innerHTML = `<li>Leyendo ${files.length} archivo(s)…</li>`; btnLocal.disabled = true;
  const next = JSON.parse(JSON.stringify(RAW)); const out = []; let ok = 0;
  for (const file of files) {
    try {
      const r = parseSource(file.name, await file.arrayBuffer(), new Date(file.lastModified));
      next[r.kind] = r.rows;
      next.meta[r.kind] = Object.assign({}, next.meta[r.kind], {fecha:toIso(r.fecha), archivo:file.name, local:true});
      out.push({ok:true, file:file.name, kind:r.kind, n:r.rows.length}); ok++;
    } catch (e) { out.push({ok:false, file:file.name, msg:e.message || String(e)}); }
  }
  const test = process(next);
  resEl.innerHTML = out.map(o => {
    if (!o.ok) return `<li class="err"><b>${esc(o.file)}</b><small>${esc(o.msg)}</small></li>`;
    const miss = [...test.miss[o.kind]];
    return `<li class="ok"><b>${esc(LABEL[o.kind])}</b>: ${o.n} filas desde ${esc(o.file)}<small>${miss.length ? `No se enlazaron ${miss.length} nombre(s): ${esc(miss.slice(0,12).join(", "))}${miss.length > 12 ? "…" : ""}. Agrégalos a ALIAS en assets/app.js.` : "Todas las comunas se enlazaron."}</small></li>`;
  }).join("");
  if (ok) { pending = next; btnLocal.disabled = false; }
}
$("openUpd").addEventListener("click", () => { resEl.innerHTML = ""; pending = null; btnLocal.disabled = true; dlg.showModal(); });
$("cancelUpd").addEventListener("click", () => dlg.close());
$("files").addEventListener("change", e => { if (e.target.files.length) handleFiles([...e.target.files]); e.target.value = ""; });
const drop = $("drop");
["dragenter","dragover"].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.add("over"); }));
["dragleave","drop"].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.remove("over"); }));
drop.addEventListener("drop", e => { const f = [...((e.dataTransfer && e.dataTransfer.files) || [])]; if (f.length) handleFiles(f); });
btnLocal.addEventListener("click", () => { if (!pending) return; RAW = pending; D = process(RAW); renderSearch(); renderSources(); renderAll(); dlg.close(); });

/* ---------- Carga inicial ---------- */
async function fetchBuf(path) {
  const r = await fetch(path, {cache:"no-cache"});
  if (!r.ok) throw new Error(`No se encontró ${path} (HTTP ${r.status}).`);
  return {buf:await r.arrayBuffer(), date:r.headers.get("last-modified") ? new Date(r.headers.get("last-modified")) : new Date()};
}
async function init() {
  const cfg = await (await fetch("data/fuentes.json", {cache:"no-cache"})).json();
  const [g, ...src] = await Promise.all([fetchBuf("data/" + cfg.geometria.archivo), ...["paccc","rrd","prc"].map(k => fetchBuf("data/" + cfg[k].archivo))]);
  GEO = JSON.parse(new TextDecoder().decode(g.buf));
  RAW = {meta:{}};
  ["paccc","rrd","prc"].forEach((k, i) => {
    const r = parseSource(cfg[k].archivo, src[i].buf, src[i].date);
    if (r.kind !== k) throw new Error(`El archivo ${cfg[k].archivo} no tiene el formato esperado para “${k}” en fuentes.json.`);
    RAW[k] = r.rows;
    RAW.meta[k] = Object.assign({}, cfg[k], {fecha:cfg[k].fecha || toIso(r.fecha)});
  });
  D = process(RAW);
  const miss = Object.entries(D.miss).filter(([, s]) => s.size).map(([k, s]) => `${k}: ${[...s].join(", ")}`);
  if (miss.length) console.warn("Comunas sin enlazar →", miss.join(" | "));
  drawGeo(); renderSearch(); renderSources(); renderAll();
  $("loading").remove();
}
init().catch(e => {
  const el = $("loading"); el.classList.add("err");
  el.innerHTML = location.protocol === "file:"
    ? `<div>El visor necesita un servidor para leer <code>data/</code>.<br>En la carpeta del proyecto ejecuta <code>python -m http.server</code> y abre <code>http://localhost:8000</code>.</div>`
    : `<div>No se pudieron cargar los datos.<br>${esc(e.message || e)}</div>`;
  console.error(e);
});
})();
