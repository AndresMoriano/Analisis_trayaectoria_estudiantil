import React, { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";

/* ============================================================
   ANÁLISIS DE TRAYECTORIAS ESTUDIANTILES
   UDR VALLE DEL GUAMUEZ (UNAD · ZONA SUR)

   Cruza hasta tres bases de datos por documento para identificar
   estudiantes que continúan, ausentes en el período, nunca
   matriculados, y trayectorias completas. Genera tablas dinámicas
   y descargables en CSV y Excel. Marca alertas de permanencia
   prolongada (9 o más matrículas).
   ============================================================ */

const C = {
  tinta: "#0F2440",
  azul: "#1F4E9C",
  azulSuave: "#E8EFFA",
  ambar: "#F5B300",
  verde: "#1B8A5A",
  verdeSuave: "#E7F4EC",
  rojo: "#C0402E",
  rojoSuave: "#FBEAE6",
  gris: "#5B6B80",
  grisClaro: "#F4F6FA",
  borde: "#DDE4EE",
};

const ESCUELAS = ["ECACEN", "ECISA", "ECBTI", "ECJP", "ECAPMA", "ECSAH", "ECEDU", "INVIL"];
const UMBRAL_MATRICULAS = 9; // 9 o más matrículas = alerta de permanencia

/* ---------------------- utilidades ---------------------- */
const sinTildes = (s) =>
  String(s ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase().trim();

const fmt = (v) =>
  v === null || v === undefined || v === "" ? "—" : Number(v).toLocaleString("es-CO");

/* Normaliza un documento para usarlo como llave: quita espacios, puntos de
   miles, sufijo ".0" que agregan las hojas de cálculo, y ceros a la izquierda
   solo si el resto es numérico (para que "0123" y "123" crucen). */
function normalizarDoc(v) {
  if (v === null || v === undefined) return null;
  let s = String(v).trim();
  if (!s || s === "-") return null;
  s = s.replace(/\s+/g, "");
  if (/^\d+\.0+$/.test(s)) s = s.split(".")[0]; // 12345.0 -> 12345
  s = s.replace(/\.(?=\d{3}\b)/g, ""); // 1.234.567 -> 1234567
  if (/^0+\d+$/.test(s)) s = s.replace(/^0+/, ""); // quita ceros a la izquierda
  return s || null;
}

function normalizarEscuela(valor) {
  if (valor === null || valor === undefined) return null;
  const v = sinTildes(valor);
  if (!v) return null;
  if (ESCUELAS.includes(v)) return v;
  if (v.includes("ADMINISTRATIVAS") || v.includes("CONTABLES")) return "ECACEN";
  if (v.includes("SALUD")) return "ECISA";
  if (v.includes("BASICAS") || v.includes("TECNOLOGIA") || v.includes("INGENIERIA")) return "ECBTI";
  if (v.includes("JURIDICAS") || v.includes("POLITICAS")) return "ECJP";
  if (v.includes("AGRICOLAS") || v.includes("PECUARIAS") || v.includes("AMBIENTE")) return "ECAPMA";
  if (v.includes("SOCIALES") || v.includes("ARTES") || v.includes("HUMANIDADES")) return "ECSAH";
  if (v.includes("EDUCACION")) return "ECEDU";
  if (v.includes("INVIL") || v.includes("LENGUAS") || v.includes("ENGLISH")) return "INVIL";
  return v; // deja el valor tal cual si no reconoce (no lo pierde)
}

/* Busca la fila de encabezados: la primera fila (en las primeras 20) que
   contenga una celda tipo "documento". Devuelve su índice y los encabezados. */
/* Algunos reportes traen un encabezado de dos niveles: una fila superior con
   el período repetido ("2026 I PERIODO 16-01" en cada celda) y debajo los
   nombres reales de columna. Si detecta esa fila superior repetitiva sin datos
   de documento, la descarta para quedarse con los encabezados verdaderos. */
function aplanarEncabezadoDoble(filas) {
  if (filas.length < 2) return filas;
  const fila0 = filas[0] || [];
  const noVacias = fila0.filter((c) => c !== null && c !== undefined && c !== "");
  if (noVacias.length < 2) return filas;
  // ¿todas las celdas no vacías de la fila 0 son el mismo texto? (período repetido)
  const primeras = noVacias.map((c) => sinTildes(c));
  const todasIguales = primeras.every((c) => c === primeras[0]);
  // ¿la fila 1 parece encabezado real? (contiene "codigo", "documento" o "nombres")
  const fila1 = (filas[1] || []).map((c) => sinTildes(c));
  const fila1EsEncabezado = fila1.some((c) => c.includes("CODIGO") || c.includes("DOCUMENTO") || c.includes("NOMBRES"));
  if (todasIguales && fila1EsEncabezado) {
    return filas.slice(1); // descarta la fila superior repetida
  }
  return filas;
}

function detectarEncabezado(filas) {
  for (let i = 0; i < Math.min(filas.length, 20); i++) {
    const fila = (filas[i] || []).map((c) => sinTildes(c));
    if (fila.some((c) => c.includes("DOCUMENTO") || c === "DOC" || c === "IDENTIFICACION" || c === "CEDULA" || c === "CODIGO")) {
      return i;
    }
  }
  // si no encontró "documento", asume la primera fila no vacía
  for (let i = 0; i < filas.length; i++) {
    if ((filas[i] || []).some((c) => c !== null && c !== undefined && c !== "")) return i;
  }
  return 0;
}

function indiceCol(enc, pruebas) {
  const cols = enc.map((c) => sinTildes(c));
  for (const prueba of pruebas) {
    const i = cols.findIndex((c) => c && prueba(c));
    if (i >= 0) return i;
  }
  return -1;
}

/* Pruebas para reconocer la columna de documento. Incluye "Código", que es
   como el SII llama al documento en sus reportes. */
const PRUEBAS_DOC = [
  (c) => c.includes("DOCUMENTO"),
  (c) => c === "DOC" || c === "IDENTIFICACION" || c === "CEDULA" || c === "ID",
  (c) => c === "CODIGO" || c === "COD",
];

/* Detecta si el contenido del archivo es en realidad HTML (algunos sistemas,
   como el SII, exportan una tabla web con extensión .xls). */
function pareceHTML(buf) {
  const muestra = new TextDecoder("utf-8").decode(new Uint8Array(buf).slice(0, 2000)).toLowerCase();
  return muestra.includes("<table") || muestra.includes("<html") || muestra.includes("<tr");
}

/* Lee un archivo y devuelve un objeto base: encabezados + filas de registros,
   cada uno con su documento normalizado y un mapa de sus columnas. */
async function leerBase(file) {
  const buf = await file.arrayBuffer();
  const wb = pareceHTML(buf)
    ? XLSX.read(new TextDecoder("utf-8").decode(new Uint8Array(buf)), { type: "string", cellDates: true })
    : XLSX.read(buf, { type: "array", cellDates: true });
  const resultados = [];
  for (const nombreHoja of wb.SheetNames) {
    const hoja = wb.Sheets[nombreHoja];
    let filas = XLSX.utils.sheet_to_json(hoja, { header: 1, defval: null, raw: true });
    if (!filas.length) continue;
    filas = aplanarEncabezadoDoble(filas);
    const hIdx = detectarEncabezado(filas);
    const enc = (filas[hIdx] || []).map((c) => (c === null || c === undefined ? "" : String(c).trim()));
    const colDoc = indiceCol(enc, PRUEBAS_DOC);
    if (colDoc < 0) continue;
    const registros = [];
    for (let i = hIdx + 1; i < filas.length; i++) {
      const fila = filas[i];
      if (!fila || fila.every((c) => c === null || c === undefined || c === "")) continue;
      const doc = normalizarDoc(fila[colDoc]);
      if (!doc) continue;
      const obj = {};
      for (let c = 0; c < enc.length; c++) {
        const clave = enc[c] || `col${c}`;
        obj[clave] = fila[c];
      }
      obj.__doc = doc;
      registros.push(obj);
    }
    resultados.push({ nombreHoja, enc, registros });
  }
  if (!resultados.length) {
    throw new Error("No se encontró una columna de documento en ninguna hoja del archivo.");
  }
  // combina todas las hojas en una sola base (concatena registros y une encabezados)
  const encTotal = [];
  const vistos = new Set();
  for (const r of resultados) for (const h of r.enc) {
    const k = sinTildes(h);
    if (!vistos.has(k) && h) { vistos.add(k); encTotal.push(h); }
  }
  const registros = resultados.flatMap((r) => r.registros);
  return { enc: encTotal, registros, hojas: resultados.map((r) => r.nombreHoja) };
}

/* Índice documento -> primer registro, para cruces rápidos. */
function indexarPorDoc(base) {
  const m = new Map();
  for (const reg of base.registros) {
    if (!m.has(reg.__doc)) m.set(reg.__doc, reg);
  }
  return m;
}

/* Descargas ------------------------------------------------ */
function descargar(nombre, blob) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = nombre;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function exportarCSV(nombre, columnas, filas) {
  const esc = (v) => {
    if (v === null || v === undefined) return "";
    const s = String(v);
    return /[",\n;]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  const lineas = [columnas.map(esc).join(";")];
  for (const f of filas) lineas.push(columnas.map((c) => esc(f[c])).join(";"));
  // BOM para que Excel abra los acentos bien
  descargar(nombre + ".csv", new Blob(["\uFEFF" + lineas.join("\n")], { type: "text/csv;charset=utf-8" }));
}

function exportarXLSX(nombre, columnas, filas) {
  const aoa = [columnas, ...filas.map((f) => columnas.map((c) => f[c] ?? ""))];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Datos");
  const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  descargar(nombre + ".xlsx", new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
}

/* ============================ APP ============================ */
export default function App() {
  const [bases, setBases] = useState([]); // {id, etiqueta, base, archivo}
  const [vista, setVista] = useState("cargar");

  useEffect(() => {
    const link = document.createElement("link");
    link.rel = "stylesheet";
    link.href = "https://fonts.googleapis.com/css2?family=Archivo:wght@400;500;600;700;800&display=swap";
    document.head.appendChild(link);
  }, []);

  const hayBases = bases.length > 0;

  return (
    <div style={{ minHeight: "100vh", background: C.grisClaro, fontFamily: "Archivo, system-ui, sans-serif", color: C.tinta }}>
      <style>{`
        .num { font-variant-numeric: tabular-nums; }
        table.dx { border-collapse: collapse; width: 100%; }
        table.dx th, table.dx td { padding: 6px 10px; border-bottom: 1px solid ${C.borde}; font-size: 13px; }
        table.dx th { text-align: left; color: ${C.gris}; font-size: 11px; letter-spacing: 0.04em; text-transform: uppercase; }
        button.tab { border: none; background: transparent; padding: 10px 14px; font: inherit; font-weight: 600; color: #B9C6DC; cursor: pointer; border-bottom: 3px solid transparent; }
        button.tab.activa { color: #fff; border-color: ${C.ambar}; }
        button.tab:disabled { opacity: 0.4; cursor: not-allowed; }
        .btn { border: none; border-radius: 8px; padding: 9px 16px; font: inherit; font-weight: 600; cursor: pointer; }
        .btn:focus-visible, button.tab:focus-visible { outline: 2px solid ${C.ambar}; outline-offset: 2px; }
        select, input[type=text] { font: inherit; padding: 7px 10px; border: 1px solid ${C.borde}; border-radius: 8px; }
      `}</style>

      <header style={{ background: C.tinta, color: "#fff" }}>
        <div style={{ maxWidth: 1180, margin: "0 auto", padding: "20px 20px 0" }}>
          <div style={{ fontSize: 11, letterSpacing: "0.18em", fontWeight: 600, color: C.ambar }}>UNAD · ZONA SUR · UDR VALLE DEL GUAMUEZ</div>
          <h1 style={{ margin: "4px 0 2px", fontSize: 26, fontWeight: 800 }}>Análisis de trayectorias estudiantiles</h1>
          <div style={{ fontSize: 13, color: "#B9C6DC" }}>Cruce de bases · retención, permanencia y continuidad · descargables en CSV y Excel</div>
          <nav style={{ marginTop: 14, display: "flex", gap: 4, flexWrap: "wrap" }}>
            {[
              ["cargar", "1 · Cargar bases", false],
              ["cruce", "2 · Cruce de trayectorias", !hayBases],
              ["dinamica", "3 · Tabla dinámica", !hayBases],
              ["alertas", "4 · Alertas de permanencia", !hayBases],
            ].map(([id, tx, dis]) => (
              <button key={id} className={`tab ${vista === id ? "activa" : ""}`} disabled={dis} onClick={() => setVista(id)}>{tx}</button>
            ))}
          </nav>
        </div>
      </header>

      <main style={{ maxWidth: 1180, margin: "0 auto", padding: "20px 20px 60px" }}>
        {vista === "cargar" && <VistaCargar bases={bases} setBases={setBases} setVista={setVista} />}
        {vista === "cruce" && <VistaCruce bases={bases} />}
        {vista === "dinamica" && <VistaDinamica bases={bases} />}
        {vista === "alertas" && <VistaAlertas bases={bases} />}
      </main>
    </div>
  );
}

/* ---------------- Tarjeta ---------------- */
function Tarjeta({ children, titulo, extra }) {
  return (
    <section style={{ background: "#fff", border: `1px solid ${C.borde}`, borderRadius: 12, padding: 18, marginBottom: 18 }}>
      {(titulo || extra) && (
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10, marginBottom: 12 }}>
          {titulo && <h2 style={{ margin: 0, fontSize: 16, fontWeight: 700 }}>{titulo}</h2>}
          {extra}
        </div>
      )}
      {children}
    </section>
  );
}

function BotonesDescarga({ nombre, columnas, filas }) {
  if (!filas.length) return null;
  return (
    <div style={{ display: "flex", gap: 8 }}>
      <button className="btn" style={{ background: C.verde, color: "#fff", fontSize: 13 }} onClick={() => exportarCSV(nombre, columnas, filas)}>Descargar CSV</button>
      <button className="btn" style={{ background: C.azul, color: "#fff", fontSize: 13 }} onClick={() => exportarXLSX(nombre, columnas, filas)}>Descargar Excel</button>
    </div>
  );
}

/* ================= VISTA 1: CARGAR ================= */
function VistaCargar({ bases, setBases, setVista }) {
  const [error, setError] = useState(null);
  const [cargando, setCargando] = useState(false);
  const inputRef = useRef(null);

  const alSeleccionar = async (e) => {
    const files = Array.from(e.target.files || []);
    if (!files.length) return;
    setCargando(true);
    setError(null);
    const nuevas = [];
    for (const file of files) {
      try {
        const base = await leerBase(file);
        const nombreLimpio = file.name.replace(/\.(xlsx|xls|csv)$/i, "");
        nuevas.push({ id: `${Date.now()}-${nuevas.length}`, etiqueta: nombreLimpio.slice(0, 40), base, archivo: file.name });
      } catch (err) {
        setError(`No se pudo leer «${file.name}»: ${err.message}`);
      }
    }
    setBases((prev) => [...prev, ...nuevas].slice(0, 3));
    setCargando(false);
    if (inputRef.current) inputRef.current.value = "";
  };

  const cambiarEtiqueta = (id, etiqueta) => setBases((prev) => prev.map((b) => (b.id === id ? { ...b, etiqueta } : b)));
  const quitar = (id) => setBases((prev) => prev.filter((b) => b.id !== id));

  return (
    <>
      <Tarjeta titulo="Cargar bases de datos (hasta 3)">
        <p style={{ fontSize: 14, color: C.gris, marginTop: 0 }}>
          Suba de una a tres bases en Excel o CSV, tal como las descarga del sistema (incluye los reportes «.xls» del SII, que
          en realidad son tablas web). La aplicación detecta sola la fila de encabezados y la columna de documento (o
          «Código»), que es la llave para cruzar. Póngale a cada base una etiqueta clara (por
          ejemplo «Matriculados 16-01», «Matriculados 16-04»); esas etiquetas se usan en los cruces. Todo el
          procesamiento ocurre en su navegador: los datos no se envían a ningún servidor.
        </p>
        {bases.length < 3 && (
          <label className="btn" style={{ background: C.ambar, color: C.tinta, display: "inline-block" }}>
            {cargando ? "Procesando…" : "Seleccionar archivo(s)"}
            <input ref={inputRef} type="file" accept=".xlsx,.xls,.csv" multiple onChange={alSeleccionar} style={{ display: "none" }} />
          </label>
        )}
        {error && (
          <div style={{ marginTop: 12, background: C.rojoSuave, border: `1px solid ${C.rojo}`, color: C.rojo, borderRadius: 8, padding: "10px 14px", fontSize: 14, fontWeight: 600 }}>{error}</div>
        )}
      </Tarjeta>

      {bases.map((b, i) => {
        const docsUnicos = new Set(b.base.registros.map((r) => r.__doc)).size;
        return (
          <Tarjeta key={b.id}>
            <div style={{ display: "flex", gap: 14, alignItems: "center", flexWrap: "wrap" }}>
              <div style={{ background: C.azul, color: "#fff", width: 30, height: 30, borderRadius: 8, display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 800, flexShrink: 0 }}>{String.fromCharCode(65 + i)}</div>
              <div style={{ flex: 1, minWidth: 220 }}>
                <label style={{ fontSize: 12, color: C.gris, fontWeight: 600 }}>Etiqueta de la base</label><br />
                <input type="text" value={b.etiqueta} onChange={(e) => cambiarEtiqueta(b.id, e.target.value)} style={{ width: "100%", maxWidth: 340, marginTop: 3 }} />
              </div>
              <div style={{ fontSize: 13, color: C.gris }}>
                <div><strong>{fmt(b.base.registros.length)}</strong> registros · <strong>{fmt(docsUnicos)}</strong> documentos únicos</div>
                <div style={{ fontSize: 12 }}>{b.archivo}{b.base.hojas.length > 1 ? ` · ${b.base.hojas.length} hojas combinadas` : ""}</div>
              </div>
              <button className="btn" style={{ background: "#fff", border: `1px solid ${C.rojo}`, color: C.rojo, fontSize: 13 }} onClick={() => quitar(b.id)}>Quitar</button>
            </div>
            <details style={{ marginTop: 10 }}>
              <summary style={{ fontSize: 12.5, color: C.azul, cursor: "pointer" }}>Ver columnas detectadas ({b.base.enc.length})</summary>
              <div style={{ fontSize: 12, color: C.gris, marginTop: 6, lineHeight: 1.7 }}>{b.base.enc.join(" · ")}</div>
            </details>
          </Tarjeta>
        );
      })}

      {bases.length > 0 && (
        <div style={{ display: "flex", gap: 10 }}>
          <button className="btn" style={{ background: C.tinta, color: "#fff" }} onClick={() => setVista("cruce")}>Continuar al cruce de trayectorias →</button>
        </div>
      )}
    </>
  );
}

/* ================= VISTA 2: CRUCE ================= */
const CATEGORIAS = [
  { id: "AyB", nom: "En A y en B", desc: "Están en ambas bases", req: 2 },
  { id: "AnoB", nom: "En A pero NO en B", desc: "Ausentes: estaban en A y no aparecen en B", req: 2 },
  { id: "BnoA", nom: "En B pero NO en A", desc: "Nuevos en B respecto de A", req: 2 },
  { id: "AyByC", nom: "En A, B y C (trayectoria completa)", desc: "Presentes en las tres bases", req: 3 },
  { id: "AyBnoC", nom: "En A y B pero NO en C", desc: "Avanzaron hasta B pero no continúan en C", req: 3 },
  { id: "AnoByC", nom: "En A, ausentes en B y C", desc: "Solo en la primera base", req: 3 },
];

function VistaCruce({ bases }) {
  const [catId, setCatId] = useState("AyB");
  const [modo, setModo] = useState("contacto"); // resumen | contacto | completo
  const cat = CATEGORIAS.find((c) => c.id === catId);
  const nBases = bases.length;

  const idxs = useMemo(() => bases.map((b) => indexarPorDoc(b.base)), [bases]);

  const resultado = useMemo(() => {
    if (nBases < cat.req) return null;
    const [A, B, Cc] = idxs;
    const setA = A ? new Set(A.keys()) : new Set();
    const setB = B ? new Set(B.keys()) : new Set();
    const setC = Cc ? new Set(Cc.keys()) : new Set();
    let docs = [];
    if (catId === "AyB") docs = [...setA].filter((d) => setB.has(d));
    else if (catId === "AnoB") docs = [...setA].filter((d) => !setB.has(d));
    else if (catId === "BnoA") docs = [...setB].filter((d) => !setA.has(d));
    else if (catId === "AyByC") docs = [...setA].filter((d) => setB.has(d) && setC.has(d));
    else if (catId === "AyBnoC") docs = [...setA].filter((d) => setB.has(d) && !setC.has(d));
    else if (catId === "AnoByC") docs = [...setA].filter((d) => !setB.has(d) && !setC.has(d));

    const indices = [A, B, Cc];

    // Toma el primer valor no vacío de un campo, buscándolo (sin tildes ni
    // mayúsculas) en los registros del estudiante en las tres bases.
    const primerValor = (regs, nombres) => {
      const claves = nombres.map(sinTildes);
      for (const r of regs) {
        if (!r) continue;
        for (const k of Object.keys(r)) {
          if (claves.includes(sinTildes(k))) {
            const v = r[k];
            if (v !== null && v !== undefined && String(v).trim() !== "") return v;
          }
        }
      }
      return "";
    };

    const filas = docs.map((d) => {
      const regs = indices.map((idx) => (idx ? idx.get(d) : null));
      // Campos clave de identidad y contacto, tomados de donde estén.
      const fila = {
        Documento: d,
        Nombres: primerValor(regs, ["NOMBRES", "NOMBRE"]),
        Apellidos: primerValor(regs, ["APELLIDOS", "APELLIDO"]),
        Telefono: primerValor(regs, ["TELEFONO", "CELULAR", "TELEFONO CELULAR"]),
        "Telefono alterno": primerValor(regs, ["TELEFONO ALTERNATIVO", "TELEFONO ALTERNO", "TELEFONO 2"]),
        Correo: primerValor(regs, ["CORREO", "CORREO CONTACTO", "EMAIL", "CORREO ELECTRONICO"]),
        "Correo alterno": primerValor(regs, ["CORREO ALTERNATIVO", "CORREO ALTERNO"]),
        "Correo institucional": primerValor(regs, ["CORREO INSTITUCIONAL", "EMAIL INSTITUCIONAL"]),
        Escuela: primerValor(regs, ["ESCUELA"]),
        Programa: primerValor(regs, ["PROGRAMA"]),
        Centro: primerValor(regs, ["CENTRO"]),
        "Ciudad residencia": primerValor(regs, ["CIUDAD DE RESIDENCIA", "CIUDAD"]),
        "Departamento residencia": primerValor(regs, ["DEPARTAMENTO DE RESIDENCIA", "DEPARTAMENTO"]),
        Condicion: primerValor(regs, ["CONDICION", "ESTADO", "TIPO ESTUDIANTE", "TIPO"]),
      };

      // Marca de presencia en cada base.
      const presencia = {};
      bases.forEach((b, i) => { presencia[`En ${b.etiqueta}`] = regs[i] ? "Sí" : "No"; });

      // Modo completo: además vuelca TODAS las columnas de cada base con el
      // prefijo de su letra, para no perder ningún dato.
      const completo = {};
      if (modo === "completo") {
        regs.forEach((r, i) => {
          if (!r) return;
          const letra = String.fromCharCode(65 + i);
          for (const k of Object.keys(r)) {
            if (k === "__doc") continue;
            completo[`[${letra}] ${k}`] = r[k];
          }
        });
      }

      return { fila, presencia, completo };
    });

    filas.sort((a, b) => String(a.fila.Nombres).localeCompare(String(b.fila.Nombres)));

    // Ensambla las columnas según el modo elegido.
    const CLAVE = ["Documento", "Nombres", "Apellidos", "Escuela", "Programa", "Centro"];
    const CONTACTO = ["Telefono", "Telefono alterno", "Correo", "Correo alterno", "Correo institucional", "Ciudad residencia", "Departamento residencia", "Condicion"];
    return filas.map(({ fila, presencia, completo }) => {
      if (modo === "resumen") {
        const o = {};
        for (const c of CLAVE) o[c] = fila[c];
        return { ...o, ...presencia };
      }
      if (modo === "contacto") {
        const o = {};
        for (const c of [...CLAVE.slice(0, 3), ...CONTACTO, ...CLAVE.slice(3)]) o[c] = fila[c];
        return { ...o, ...presencia };
      }
      // completo
      return { ...fila, ...presencia, ...completo };
    });
  }, [idxs, catId, cat, nBases, bases, modo]);

  // Columnas = unión de las claves de todas las filas (en modo completo cada
  // estudiante puede traer columnas distintas según en qué bases aparezca).
  const columnas = useMemo(() => {
    if (!resultado || !resultado.length) return [];
    const set = [];
    const vistos = new Set();
    for (const f of resultado) for (const k of Object.keys(f)) {
      if (!vistos.has(k)) { vistos.add(k); set.push(k); }
    }
    return set;
  }, [resultado]);
  const nombreArchivo = `cruce_${catId}_${modo}_${new Date().toISOString().slice(0, 10)}`;

  const etiquetaCat = (txt) => txt
    .replace(/\bA\b/g, bases[0] ? bases[0].etiqueta : "A")
    .replace(/\bB\b/g, bases[1] ? bases[1].etiqueta : "B")
    .replace(/\bC\b/g, bases[2] ? bases[2].etiqueta : "C");

  return (
    <>
      <Tarjeta titulo="Cruce de trayectorias">
        <p style={{ fontSize: 13.5, color: C.gris, marginTop: 0 }}>
          Elija qué grupo quiere identificar. Las letras corresponden a sus bases en orden de carga:
          {bases.map((b, i) => <span key={b.id}> <strong>{String.fromCharCode(65 + i)}</strong> = {b.etiqueta}{i < bases.length - 1 ? "," : "."}</span>)}
        </p>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(240px, 1fr))", gap: 10 }}>
          {CATEGORIAS.filter((c) => c.req <= nBases).map((c) => (
            <button key={c.id} onClick={() => setCatId(c.id)} className="btn"
              style={{ textAlign: "left", padding: "12px 14px", background: catId === c.id ? C.azulSuave : "#fff", border: `1px solid ${catId === c.id ? C.azul : C.borde}` }}>
              <div style={{ fontWeight: 700, fontSize: 13.5 }}>{etiquetaCat(c.nom)}</div>
              <div style={{ fontSize: 12, color: C.gris, marginTop: 2 }}>{etiquetaCat(c.desc)}</div>
            </button>
          ))}
        </div>
      </Tarjeta>

      {resultado && (
        <Tarjeta
          titulo={`${fmt(resultado.length)} estudiantes — ${etiquetaCat(cat.nom)}`}
          extra={<BotonesDescarga nombre={nombreArchivo} columnas={columnas} filas={resultado} />}
        >
          <div style={{ display: "flex", gap: 12, alignItems: "center", flexWrap: "wrap", marginBottom: 12 }}>
            <span style={{ fontSize: 12.5, fontWeight: 600, color: C.gris }}>Nivel de detalle:</span>
            <div style={{ display: "inline-flex", border: `1px solid ${C.borde}`, borderRadius: 8, overflow: "hidden" }}>
              {[["resumen", "Resumen"], ["contacto", "Con datos de contacto"], ["completo", "Todos los campos"]].map(([id, tx]) => (
                <button key={id} className="btn" onClick={() => setModo(id)}
                  style={{ borderRadius: 0, background: modo === id ? C.azul : "#fff", color: modo === id ? "#fff" : C.tinta, fontWeight: 600, fontSize: 12.5 }}>
                  {tx}
                </button>
              ))}
            </div>
          </div>
          <p style={{ fontSize: 12.5, color: C.gris, marginTop: 0, marginBottom: 12 }}>
            {modo === "resumen" && "Solo identificación básica. Útil para un vistazo rápido."}
            {modo === "contacto" && "Incluye teléfonos, correos y ciudad para poder contactar al estudiante. Los datos se toman de la base que los tenga (por ejemplo, teléfono y correo suelen venir de la base de inscritos)."}
            {modo === "completo" && "Vuelca TODAS las columnas de las tres bases, cada una con el prefijo de su letra ([A], [B], [C]). La tabla se ve ancha; la descarga trae absolutamente todo."}
          </p>
          {resultado.length === 0 ? (
            <p style={{ color: C.gris, fontSize: 14 }}>Ningún estudiante cumple esta condición con las bases cargadas.</p>
          ) : (
            <TablaResultado columnas={columnas} filas={resultado} />
          )}
        </Tarjeta>
      )}
    </>
  );
}

function TablaResultado({ columnas, filas, max = 200 }) {
  const mostradas = filas.slice(0, max);
  return (
    <>
      <div style={{ overflowX: "auto" }}>
        <table className="dx">
          <thead><tr>{columnas.map((c) => <th key={c}>{c}</th>)}</tr></thead>
          <tbody>
            {mostradas.map((f, i) => (
              <tr key={i}>{columnas.map((c) => <td key={c} className={typeof f[c] === "number" ? "num" : ""}>{f[c] === null || f[c] === undefined || f[c] === "" ? "—" : String(f[c])}</td>)}</tr>
            ))}
          </tbody>
        </table>
      </div>
      {filas.length > max && (
        <p style={{ fontSize: 12.5, color: C.gris, marginTop: 8 }}>Mostrando los primeros {max} de {fmt(filas.length)}. La descarga incluye todos.</p>
      )}
    </>
  );
}

/* ================= VISTA 3: TABLA DINÁMICA ================= */
function VistaDinamica({ bases }) {
  const [baseId, setBaseId] = useState(bases[0] ? bases[0].id : "");
  const b = bases.find((x) => x.id === baseId) || bases[0];
  const [campoFila, setCampoFila] = useState("");
  const [campoCol, setCampoCol] = useState("");

  useEffect(() => {
    if (b && !campoFila) {
      const pref = b.base.enc.find((c) => /ESCUELA/i.test(c)) || b.base.enc.find((c) => /PROGRAMA|CENTRO|ESTADO|CONDICI/i.test(c)) || b.base.enc[0];
      setCampoFila(pref || "");
    }
  }, [b]); // eslint-disable-line

  const pivote = useMemo(() => {
    if (!b || !campoFila) return null;
    const valFila = (r) => {
      const v = r[campoFila];
      if (/ESCUELA/i.test(campoFila)) return normalizarEscuela(v) || "(sin dato)";
      return v === null || v === undefined || v === "" ? "(sin dato)" : String(v).trim();
    };
    const valCol = (r) => {
      const v = r[campoCol];
      return v === null || v === undefined || v === "" ? "(sin dato)" : String(v).trim();
    };
    const filasSet = new Map();
    const colsSet = new Set();
    // cuenta documentos únicos por celda
    const vistos = new Set();
    for (const r of b.base.registros) {
      const cf = valFila(r);
      const cc = campoCol ? valCol(r) : "Total";
      colsSet.add(cc);
      if (!filasSet.has(cf)) filasSet.set(cf, {});
      const fila = filasSet.get(cf);
      const claveUnica = r.__doc + "|" + cf + "|" + cc;
      if (!vistos.has(claveUnica)) {
        vistos.add(claveUnica);
        fila[cc] = (fila[cc] || 0) + 1;
      }
    }
    const cols = [...colsSet].sort();
    const filas = [...filasSet.keys()].sort().map((cf) => {
      const o = { [campoFila]: cf };
      let tot = 0;
      for (const c of cols) { o[c] = filasSet.get(cf)[c] || 0; tot += o[c]; }
      o["Total"] = tot;
      return o;
    });
    // fila de totales
    const totalRow = { [campoFila]: "TOTAL" };
    let gran = 0;
    for (const c of cols) { totalRow[c] = filas.reduce((a, f) => a + f[c], 0); gran += totalRow[c]; }
    totalRow["Total"] = gran;
    return { cols: [campoFila, ...cols, "Total"], filas, totalRow };
  }, [b, campoFila, campoCol]);

  if (!b) return null;
  const columnasExport = pivote ? pivote.cols : [];
  const filasExport = pivote ? [...pivote.filas, pivote.totalRow] : [];

  return (
    <>
      <Tarjeta titulo="Tabla dinámica">
        <p style={{ fontSize: 13.5, color: C.gris, marginTop: 0 }}>
          Cuenta estudiantes (documentos únicos) agrupados por los campos que elija. Deje la columna en «(ninguno)» para un
          conteo simple por filas.
        </p>
        <div style={{ display: "flex", gap: 14, flexWrap: "wrap" }}>
          <label style={{ fontSize: 12.5, fontWeight: 600 }}>Base<br />
            <select value={baseId} onChange={(e) => { setBaseId(e.target.value); setCampoFila(""); setCampoCol(""); }} style={{ marginTop: 3 }}>
              {bases.map((x, i) => <option key={x.id} value={x.id}>{String.fromCharCode(65 + i)} · {x.etiqueta}</option>)}
            </select>
          </label>
          <label style={{ fontSize: 12.5, fontWeight: 600 }}>Agrupar por (filas)<br />
            <select value={campoFila} onChange={(e) => setCampoFila(e.target.value)} style={{ marginTop: 3 }}>
              {b.base.enc.map((c) => <option key={c} value={c}>{c}</option>)}
            </select>
          </label>
          <label style={{ fontSize: 12.5, fontWeight: 600 }}>Cruzar con (columnas)<br />
            <select value={campoCol} onChange={(e) => setCampoCol(e.target.value)} style={{ marginTop: 3 }}>
              <option value="">(ninguno)</option>
              {b.base.enc.map((c) => <option key={c} value={c}>{c}</option>)}
            </select>
          </label>
        </div>
      </Tarjeta>

      {pivote && (
        <Tarjeta
          titulo={`Conteo por ${campoFila}${campoCol ? ` × ${campoCol}` : ""}`}
          extra={<BotonesDescarga nombre={`dinamica_${new Date().toISOString().slice(0, 10)}`} columnas={columnasExport} filas={filasExport} />}
        >
          <div style={{ overflowX: "auto" }}>
            <table className="dx">
              <thead><tr>{pivote.cols.map((c) => <th key={c} style={{ textAlign: c === campoFila ? "left" : "right" }}>{c}</th>)}</tr></thead>
              <tbody>
                {pivote.filas.map((f, i) => (
                  <tr key={i}>{pivote.cols.map((c) => <td key={c} className="num" style={{ textAlign: c === campoFila ? "left" : "right", fontWeight: c === campoFila ? 600 : 400 }}>{c === campoFila ? f[c] : fmt(f[c])}</td>)}</tr>
                ))}
                <tr style={{ background: C.azulSuave, fontWeight: 800 }}>
                  {pivote.cols.map((c) => <td key={c} className="num" style={{ textAlign: c === campoFila ? "left" : "right" }}>{c === campoFila ? "TOTAL" : fmt(pivote.totalRow[c])}</td>)}
                </tr>
              </tbody>
            </table>
          </div>
        </Tarjeta>
      )}
    </>
  );
}

/* ================= VISTA 4: ALERTAS DE PERMANENCIA ================= */
function VistaAlertas({ bases }) {
  const [baseId, setBaseId] = useState("");
  const [umbral, setUmbral] = useState(UMBRAL_MATRICULAS);

  // detecta automáticamente la base que tenga columna de número de matrícula
  const basesConMatriculas = useMemo(() => bases.filter((b) => b.base.enc.some((c) => /NUM.*MATRICUL|MATRICULA/i.test(sinTildes(c)) && !/PERIODO/i.test(c))), [bases]);
  useEffect(() => {
    if (!baseId && basesConMatriculas[0]) setBaseId(basesConMatriculas[0].id);
  }, [basesConMatriculas, baseId]);

  const b = bases.find((x) => x.id === baseId);
  const colMat = b ? b.base.enc.find((c) => /NUM.*MATRICUL/i.test(sinTildes(c))) || b.base.enc.find((c) => /MATRICULA/i.test(sinTildes(c)) && !/PERIODO/i.test(c)) : null;

  const resultado = useMemo(() => {
    if (!b || !colMat) return null;
    const filas = [];
    const vistos = new Set();
    for (const r of b.base.registros) {
      const n = Number(r[colMat]);
      if (!isFinite(n) || n < umbral) continue;
      if (vistos.has(r.__doc)) continue;
      vistos.add(r.__doc);
      filas.push({
        Documento: r.__doc,
        Nombres: r["NOMBRES"] || r["Nombres"] || "",
        Escuela: normalizarEscuela(r["ESCUELA"] || r["Escuela"]) || "",
        Programa: r["PROGRAMA"] || r["Programa"] || "",
        "N° matrículas": n,
        Promedio: r["Promedio"] ?? r["PROMEDIO"] ?? "",
        Resultado: r["Resultado"] ?? r["RESULTADO"] ?? "",
      });
    }
    filas.sort((a, b2) => b2["N° matrículas"] - a["N° matrículas"]);
    return filas;
  }, [b, colMat, umbral]);

  const porEscuela = useMemo(() => {
    if (!resultado) return null;
    const m = {};
    for (const f of resultado) { const e = f.Escuela || "(sin dato)"; m[e] = (m[e] || 0) + 1; }
    return m;
  }, [resultado]);

  const columnas = resultado && resultado.length ? Object.keys(resultado[0]) : [];

  return (
    <>
      <Tarjeta titulo="Alertas de permanencia prolongada">
        <p style={{ fontSize: 13.5, color: C.gris, marginTop: 0 }}>
          Identifica estudiantes con muchas matrículas acumuladas, señal de permanencia prolongada que amerita seguimiento
          dentro de la política de retención y permanencia. Usa la columna de número de matrícula de las actas.
        </p>
        {basesConMatriculas.length === 0 ? (
          <div style={{ background: C.rojoSuave, border: `1px solid ${C.rojo}`, color: C.rojo, borderRadius: 8, padding: "10px 14px", fontSize: 14 }}>
            Ninguna base cargada tiene una columna de «Número de matrícula». Suba la base de actas de matrícula, que sí la incluye.
          </div>
        ) : (
          <div style={{ display: "flex", gap: 16, flexWrap: "wrap", alignItems: "end" }}>
            <label style={{ fontSize: 12.5, fontWeight: 600 }}>Base<br />
              <select value={baseId} onChange={(e) => setBaseId(e.target.value)} style={{ marginTop: 3 }}>
                {basesConMatriculas.map((x, i) => <option key={x.id} value={x.id}>{x.etiqueta}</option>)}
              </select>
            </label>
            <label style={{ fontSize: 12.5, fontWeight: 600 }}>Umbral (matrículas o más)<br />
              <input type="number" min={1} value={umbral} onChange={(e) => setUmbral(Number(e.target.value) || 1)} style={{ width: 90, marginTop: 3, padding: "7px 10px", border: `1px solid ${C.borde}`, borderRadius: 8 }} />
            </label>
          </div>
        )}
      </Tarjeta>

      {resultado && (
        <>
          <Tarjeta titulo={`${fmt(resultado.length)} estudiantes con ${umbral} o más matrículas`}
            extra={<BotonesDescarga nombre={`alertas_permanencia_${umbral}mas_${new Date().toISOString().slice(0, 10)}`} columnas={columnas} filas={resultado} />}>
            {porEscuela && Object.keys(porEscuela).length > 0 && (
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 14 }}>
                {Object.entries(porEscuela).sort((a, b2) => b2[1] - a[1]).map(([e, n]) => (
                  <div key={e} style={{ background: C.grisClaro, border: `1px solid ${C.borde}`, borderRadius: 8, padding: "6px 12px", fontSize: 13 }}>
                    <strong>{e}</strong>: {n}
                  </div>
                ))}
              </div>
            )}
            {resultado.length === 0 ? (
              <p style={{ color: C.gris, fontSize: 14 }}>Ningún estudiante alcanza ese umbral en la base seleccionada.</p>
            ) : (
              <TablaResultado columnas={columnas} filas={resultado} />
            )}
          </Tarjeta>
        </>
      )}
    </>
  );
}
