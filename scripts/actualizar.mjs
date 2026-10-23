// Actualización automática de las fuentes del visor y registro de seguimiento.
// La ejecuta .github/workflows/actualizar.yml cada semana; también puede correrse a mano:
//   npm install xlsx papaparse && node scripts/actualizar.mjs
//
// 1. PACCC: lee las 345 comunas del visualizador del MMA y, si algo cambió, guarda un CSV nuevo.
// 2. RRD y PRC: si en data/fuentes.json la fuente tiene "descarga" (URL directa al archivo),
//    lo descarga y lo reemplaza cuando cambia. Sin "descarga", se actualizan a mano como siempre.
// 3. Calcula los indicadores con la misma lógica de assets/app.js y agrega una fila por región
//    a data/seguimiento.csv cada vez que cambia algún dato.
import fs from "node:fs";
import path from "node:path";
import crypto from "node:crypto";
import { createRequire } from "node:module";

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

/* ---------- 2. Fuentes con URL de descarga directa ---------- */
async function actualizarDescarga(k) {
  const url = cfg[k].descarga;
  if (!url) { log(`${k.toUpperCase()}: sin URL de descarga, se mantiene ${cfg[k].archivo}`); return; }
  const nuevo = await get(url);
  const destino = path.join(DATA, cfg[k].archivo);
  if (fs.existsSync(destino) && sha(fs.readFileSync(destino)) === sha(nuevo)) { log(`${k.toUpperCase()}: sin cambios`); return; }
  fs.writeFileSync(destino, nuevo);
  cfg[k].fecha = k === "rrd" ? "" : hoy; // RRD: el visor toma la fecha del título del Excel
  cambios.push(k.toUpperCase());
  log(`${k.toUpperCase()}: actualizado desde ${url}`);
}

/* ---------- 3. Indicadores con la lógica del visor ---------- */
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
if (!process.env.SIN_DESCARGAS) for (const [nombre, fn] of [["PACCC", actualizarPaccc], ["RRD", () => actualizarDescarga("rrd")], ["PRC", () => actualizarDescarga("prc")]]) {
  try { await fn(); } catch (e) { errores.push(`${nombre}: ${e.message}`); console.error("✗", nombre, e.message); }
}
registrar(indicadores());
fs.writeFileSync(FUENTES, JSON.stringify(cfg, null, 2) + "\n");
fs.writeFileSync(path.join(ROOT, ".resumen-actualizacion"), cambios.length ? `Datos al ${hoy}: ${cambios.join(", ")}` : "");
if (errores.length) { console.error(errores.join("\n")); process.exitCode = 1; }
