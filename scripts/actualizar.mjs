// Actualización automática de las fuentes del visor y registro de seguimiento.
// La ejecuta .github/workflows/actualizar.yml cada semana; también puede correrse a mano:
//   npm install xlsx papaparse && node scripts/actualizar.mjs
//
// 1. PACCC: lee las 345 comunas del visualizador del MMA y, si algo cambió, guarda un CSV nuevo.
// 2. RRD: lee el Power BI público de SENAPRED (data/fuentes.json → rrd.powerbi) y, si cambió,
//    guarda RRD_DD_MM_AAAA.xlsx con el mismo formato del Excel oficial.
// 3. PRC: descarga el CSV de instrumentos de origen desde la API del Portal IPT (prc.api).
// 4. Calcula los indicadores con la misma lógica de assets/app.js y agrega una fila por región
//    a data/seguimiento.csv cada vez que cambia algún dato.
import fs from "node:fs";
import path from "node:path";
import crypto from "node:crypto";
import os from "node:os";
import { execFileSync } from "node:child_process";
import { createRequire } from "node:module";
import { abrirInforme, consultar, enlaces, esquema } from "./powerbi.mjs";

const require = createRequire(import.meta.url);
const XLSX = require("xlsx");
const Papa = require("papaparse");
const ROOT = path.resolve(path.dirname(new URL(import.meta.url).pathname), "..");
const DATA = path.join(ROOT, "data");
const FUENTES = path.join(DATA, "fuentes.json");
const SEGUIMIENTO = path.join(DATA, "seguimiento.csv");
const MMA_URL = "https://cambioclimatico.mma.gob.cl/planes-climaticos/comunales";

const hoy = process.env.FECHA_REGISTRO || new Intl.DateTimeFormat("en-CA", { timeZone: "America/Santiago" }).format(new Date()); // AAAA-MM-DD
const cfg = JSON.parse(fs.readFileSync(FUENTES, "utf8"));
const sha = b => crypto.createHash("sha256").update(b).digest("hex");
const log = (...a) => console.log("•", ...a);
let cambios = [];

async function get(url) {
  const r = await fetch(url, { headers: { "User-Agent": "planificacion-en-chile (GitHub Actions)" } });
  if (!r.ok) throw new Error(`${url} respondió ${r.status}`);
  return Buffer.from(await r.arrayBuffer());
}

/* ---------- 1. PACCC desde el visualizador del MMA ---------- */
async function actualizarPaccc() {
  const html = (await get(MMA_URL)).toString("utf8");
  const m = html.match(/const COMUNAS = (\[.*?\]);/s);
  if (!m) throw new Error("No encontré la lista COMUNAS en el visualizador del MMA; puede haber cambiado la página.");
  const lista = JSON.parse(m[1]);
  if (lista.length < 340) throw new Error(`El MMA entregó solo ${lista.length} comunas; no actualizo.`);
  const EST = { aprobado: "Aprobado", en_proceso: "En proceso", sin_informacion: "Sin información" };
  const q = v => (/[\s;"]/.test(v) ? `"${String(v).replace(/"/g, '""')}"` : String(v ?? ""));
  const filas = lista.map(c => [c.codigo, c.nombre, c.provincia, c.region, EST[c.estado] || c.estado, c.documento_url || "", c.decreto_url || ""].map(q).join(";"));
  const csv = "﻿" + ["cut_com;comuna;provincia;region;estado;plan_url;decreto_url", ...filas].join("\n") + "\n";

  // Compara contenido (no el nombre del archivo) con el CSV vigente
  const actual = Papa.parse(fs.readFileSync(path.join(DATA, cfg.paccc.archivo), "utf8").replace(/^﻿/, ""), { header: true, delimiter: ";", skipEmptyLines: true }).data;
  const firma = rows => rows.map(r => [String(r.cut_com).padStart(5, "0"), r.estado, r.plan_url || "", r.decreto_url || ""].join("|")).sort().join("\n");
  const nueva = Papa.parse(csv.replace(/^﻿/, ""), { header: true, delimiter: ";", skipEmptyLines: true }).data;
  if (firma(actual) === firma(nueva)) { log("PACCC: sin cambios"); return; }

  const archivo = `dataset-paccc-comunas-${hoy}.csv`;
  fs.writeFileSync(path.join(DATA, archivo), csv);
  if (cfg.paccc.archivo !== archivo) fs.rmSync(path.join(DATA, cfg.paccc.archivo), { force: true });
  Object.assign(cfg.paccc, { archivo, fecha: hoy });
  cambios.push("PACCC");
  log("PACCC: actualizado ->", archivo);
}

/* ---------- 2. Plan comunal RRD desde el Power BI de SENAPRED ---------- */
// Primero intenta la "Base de datos Planes RRD" que el informe enlaza (Excel oficial, a veces en .rar).
// Si no se puede, arma el mismo Excel con la tabla del informe (todo menos la columna FINANCIAMIENTO).
const normH = s => String(s ?? "").normalize("NFD").replace(/[̀-ͯ]/g, "").toUpperCase().replace(/[^A-Z]/g, "");

function leerRrd(buf) {
  const wb = XLSX.read(buf, { type: "buffer" });
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: "", raw: false });
  const hi = rows.findIndex(r => r.some(v => normH(v) === "COMUNA"));
  if (hi < 0) throw new Error("el Excel no tiene la columna Comuna");
  const H = rows[hi].map(normH), col = f => H.findIndex(f);
  const ix = { comuna: col(h => h === "COMUNA"), estado: col(h => h === "ESTADO"), anio: col(h => h.startsWith("ANO")),
               decreto: col(h => h.includes("DECRETO")), revision: col(h => h.includes("REVISION")) };
  if (ix.estado < 0 || ix.revision < 0) throw new Error("el Excel no tiene las columnas Estado y Estado de Revisión");
  const filas = rows.slice(hi + 1).filter(r => String(r[ix.comuna]).trim());
  const val = v => String(v ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase().replace(/\s+/g, " ").trim();
  const firma = filas.map(r => Object.values(ix).map(i => (i < 0 ? "" : val(r[i]))).join("|")).sort().join("\n");
  return { n: filas.length, firma };
}

function buscarXlsx(dir) {
  for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
    const p = path.join(dir, e.name);
    if (e.isDirectory()) { const r = buscarXlsx(p); if (r) return r; }
    else if (/\.xlsx$/i.test(e.name) && !e.name.startsWith("~$")) return p;
  }
  return null;
}

async function rrdDesdeArchivo(pbi) {
  const cands = enlaces(pbi, /realfilename=RRD/i).map(u => {
    const nombre = new URL(u).searchParams.get("realfilename") || "";
    const m = nombre.match(/(\d{2})[._-](\d{2})[._-](\d{4})/);
    return { u, nombre, fecha: m ? `${m[3]}-${m[2]}-${m[1]}` : "" };
  }).filter(c => c.fecha).sort((a, b) => b.fecha.localeCompare(a.fecha));
  if (!cands.length) throw new Error("el informe no enlaza una base de datos RRD con fecha");
  const c = cands[0];
  const buf = await get(c.u);
  if (/\.xlsx$/i.test(c.nombre) && buf[0] === 0x50 && buf[1] === 0x4b) return { buf, fecha: c.fecha, origen: c.nombre };
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "rrd-"));
  try {
    const arch = path.join(tmp, "descarga" + (path.extname(c.nombre) || ".rar"));
    fs.writeFileSync(arch, buf);
    const dest = path.join(tmp, "x");
    try { execFileSync("7z", ["x", "-y", `-o${dest}`, arch], { stdio: "ignore" }); }
    catch { fs.mkdirSync(dest, { recursive: true }); execFileSync("unrar", ["x", "-y", arch, dest + "/"], { stdio: "ignore" }); }
    const x = buscarXlsx(dest);
    if (!x) throw new Error(`${c.nombre} no trae un Excel`);
    return { buf: fs.readFileSync(x), fecha: c.fecha, origen: c.nombre };
  } finally { fs.rmSync(tmp, { recursive: true, force: true }); }
}

async function rrdDesdeTabla(pbi) {
  const ents = await esquema(pbi);
  const tiene = (e, re) => e.columnas.find(c => re.test(normH(c)));
  const e = ents.find(e => e.entidad === cfg.rrd.tabla) ||
            ents.find(e => !/EMERG/.test(normH(e.entidad)) && tiene(e, /^COMUNA$/) && tiene(e, /REVISION/) && tiene(e, /DECRETO/));
  if (!e) throw new Error("no encontré la tabla de planes comunales RRD en el modelo");
  const want = [["Región", /^REGION$/], ["CUT", /^CUT$/], ["Comuna", /^COMUNA$/], ["Estado", /^ESTADO(PLAN)?$/], ["Año", /^ANO/],
                ["N° Decreto Alcaldicio", /DECRETO/], ["Estado de Revisión por Informe Técnico", /REVISION/], ["Enlace", /^ENLACE$/]]
    .map(([titulo, re]) => [titulo, tiene(e, re)]).filter(([, c]) => c);
  const filas = await consultar(pbi, e.entidad, want.map(([, c]) => c));
  let fecha = hoy;
  try {
    const f = ents.find(x => x.entidad === "Fecha");
    const [r] = await consultar(pbi, "Fecha", [f.columnas[0]]);
    const v = Object.values(r)[0];
    if (typeof v === "number") fecha = new Date(v).toISOString().slice(0, 10);
  } catch (err) { log(`RRD: no pude leer la fecha del informe (${err.message}); uso la de hoy`); }
  const [a, m, d] = fecha.split("-");
  const aoa = [[`Fecha de Actualización de Planes Comunales RRD: ${d}-${m}-${a}`], [], want.map(([t]) => t),
               ...filas.map(r => want.map(([, c]) => r[c] ?? ""))];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), "RRD");
  return { buf: XLSX.write(wb, { type: "buffer", bookType: "xlsx" }), fecha, origen: `tabla ${e.entidad} del Power BI` };
}

async function actualizarRrd() {
  if (!cfg.rrd.powerbi) { log("RRD: sin URL de Power BI, se mantiene", cfg.rrd.archivo); return; }
  const pbi = await abrirInforme(cfg.rrd.powerbi);
  let nuevo;
  for (const [nombre, fn] of [["archivo", rrdDesdeArchivo], ["tabla", rrdDesdeTabla]]) {
    try {
      const r = await fn(pbi);
      const l = leerRrd(r.buf);
      if (l.n < 340) throw new Error(`solo ${l.n} comunas`);
      nuevo = { ...r, ...l };
      break;
    } catch (e) { log(`RRD: no pude usar el ${nombre} (${e.message})`); }
  }
  if (!nuevo) throw new Error("no pude leer el Power BI de SENAPRED");
  const actual = leerRrd(fs.readFileSync(path.join(DATA, cfg.rrd.archivo)));
  if (actual.firma === nuevo.firma) { log(`RRD: sin cambios (${nuevo.origen}, ${nuevo.fecha})`); return; }
  const [a, m, d] = nuevo.fecha.split("-");
  const archivo = `RRD_${d}_${m}_${a}.xlsx`;
  fs.writeFileSync(path.join(DATA, archivo), nuevo.buf);
  if (cfg.rrd.archivo !== archivo) fs.rmSync(path.join(DATA, cfg.rrd.archivo), { force: true });
  Object.assign(cfg.rrd, { archivo, fecha: nuevo.fecha });
  cambios.push("RRD");
  log(`RRD: actualizado -> ${archivo} (${nuevo.origen})`);
}

/* ---------- 3. Planes reguladores desde la API del Portal IPT (MINVU) ---------- */
// Es el mismo CSV del botón "Descargar CSV" del buscador, con el filtro "Instrumento de origen".
async function actualizarPrc() {
  if (!cfg.prc.api) { log("PRC: sin URL de API, se mantiene", cfg.prc.archivo); return; }
  const filtro = "clasificacion=Instrumento+de+origen&format=csv";
  const intentos = [`paginated=false&${filtro}`, `paginated=true&page=1&perPage=10000&${filtro}`, `paginated=true&page=1&perPage=10&${filtro}`];
  let mejor = null; const vistos = [];
  for (const q of intentos) {
    try {
      const buf = await get(`${cfg.prc.api}?${q}`);
      let t = buf.toString("utf8"); if (t.includes("�")) t = buf.toString("latin1");
      t = t.replace(/^﻿/, "");
      const p = Papa.parse(t, { header: true, skipEmptyLines: true, delimiter: "" });
      const f = (p.meta.fields || []).map(normH);
      if (!f.includes("DENOMINACION") || !f.includes("COMUNAS")) { vistos.push(`${q}: sin columnas esperadas`); continue; }
      vistos.push(`${q}: ${p.data.length} filas`);
      if (!mejor || p.data.length > mejor.n) mejor = { n: p.data.length, texto: t };
      if (p.data.length >= 700) break;
    } catch (e) { vistos.push(`${q}: ${e.message}`); }
  }
  if (!mejor || mejor.n < 700) throw new Error(`el Portal IPT no entregó el listado completo (${vistos.join("; ")})`);
  const destino = path.join(DATA, cfg.prc.archivo);
  const limpia = s => s.replace(/^﻿/, "").replace(/\r/g, "").trim();
  if (fs.existsSync(destino) && limpia(fs.readFileSync(destino, "utf8")) === limpia(mejor.texto)) { log("PRC: sin cambios"); return; }
  fs.writeFileSync(destino, "﻿" + mejor.texto);
  cfg.prc.fecha = hoy;
  cambios.push("PRC");
  log(`PRC: actualizado (${mejor.n} instrumentos de origen)`);
}

/* ---------- 4. Indicadores con la lógica del visor ---------- */
function indicadores() {
  const src = fs.readFileSync(path.join(ROOT, "assets", "app.js"), "utf8");
  const body = src.slice(src.indexOf("const COL"), src.indexOf("const state ="));
  const GEO = JSON.parse(fs.readFileSync(path.join(DATA, cfg.geometria.archivo), "utf8"));
  const A = new Function("XLSX", "Papa", "GEO_IN", body.replace("let GEO = null", "let GEO = GEO_IN") + "\nreturn {parseSource, process, REGIONS};")(XLSX, Papa, GEO);
  const raw = {}, fechas = {};
  for (const k of ["paccc", "rrd", "prc"]) {
    const b = fs.readFileSync(path.join(DATA, cfg[k].archivo));
    const r = A.parseSource(cfg[k].archivo, b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength), null);
    if (r.kind !== k) throw new Error(`${cfg[k].archivo} no tiene el formato de ${k}`);
    raw[k] = r.rows;
    fechas[k] = cfg[k].fecha || (r.fecha ? r.fecha.toISOString().slice(0, 10) : "");
    if (k === "rrd" && !cfg.rrd.fecha && fechas.rrd) cfg.rrd.fecha = fechas.rrd;
  }
  const { C, miss } = A.process(raw);
  for (const k in miss) if (miss[k].size) console.warn(`⚠ ${k}: comunas sin enlazar: ${[...miss[k]].join(", ")} (agregar a ALIAS en assets/app.js)`);
  const filas = [];
  const agrega = (nombre, cs) => filas.push({
    region: nombre, comunas: cs.length,
    paccc_aprobado: cs.filter(c => c.paccc === "aprobado").length,
    rrd_decreto: cs.filter(c => c.rrd === "decreto").length,
    prc_al_dia: cs.filter(c => c.prc === "actual").length,
    tres_al_dia: cs.filter(c => c.paccc === "aprobado" && c.rrd === "decreto" && c.prc === "actual").length,
    ninguno: cs.filter(c => c.paccc !== "aprobado" && c.rrd !== "decreto" && c.prc !== "actual").length
  });
  const todas = [...C.values()];
  agrega("Chile", todas);
  for (const [r, n] of A.REGIONS) agrega(n, todas.filter(c => c.reg === r));
  return { filas, fechas };
}

function registrar({ filas, fechas }) {
  const head = "fecha_registro;fecha_paccc;fecha_rrd;fecha_prc;region;comunas;paccc_aprobado;rrd_decreto;prc_al_dia;tres_al_dia;ninguno";
  const prev = fs.existsSync(SEGUIMIENTO) ? fs.readFileSync(SEGUIMIENTO, "utf8").trim().split("\n") : [head];
  const ultimo = prev.slice(1).filter(l => l.split(";")[4] === "Chile").pop();
  const nuevoChile = filas[0];
  const clave = l => l.split(";").slice(1, 4).join(";");
  const lineas = filas.map(f => [hoy, fechas.paccc, fechas.rrd, fechas.prc, f.region, f.comunas, f.paccc_aprobado, f.rrd_decreto, f.prc_al_dia, f.tres_al_dia, f.ninguno].join(";"));
  if (ultimo && clave(ultimo) === clave(lineas[0])) { log("Seguimiento: mismas fechas de corte, no agrego filas"); return; }
  fs.writeFileSync(SEGUIMIENTO, [...prev, ...lineas].join("\n") + "\n");
  log(`Seguimiento: Chile ${nuevoChile.paccc_aprobado} PACCC, ${nuevoChile.rrd_decreto} RRD, ${nuevoChile.prc_al_dia} PRC, ${nuevoChile.tres_al_dia} con los tres`);
}

const errores = [];
if (!process.env.SIN_DESCARGAS) for (const [nombre, fn] of [["PACCC", actualizarPaccc], ["RRD", actualizarRrd], ["PRC", actualizarPrc]]) {
  try { await fn(); } catch (e) { errores.push(`${nombre}: ${e.message}`); console.error("✗", nombre, e.message); }
}
registrar(indicadores());
fs.writeFileSync(FUENTES, JSON.stringify(cfg, null, 2) + "\n");
fs.writeFileSync(path.join(ROOT, ".resumen-actualizacion"), cambios.length ? `Datos al ${hoy}: ${cambios.join(", ")}` : "");
if (errores.length) { console.error(errores.join("\n")); process.exitCode = 1; }
