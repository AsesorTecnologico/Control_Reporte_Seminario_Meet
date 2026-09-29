let padron = [];
let nombresMd = [];
let estadosActuales = [];
let resultadoActual = [];
let noUbicadosActual = [];

const $ = id => document.getElementById(id);

const CAMPOS = [
  "SEDE",
  "NIVEL",
  "TURNO",
  "TIPO GRADO",
  "GRADO",
  "SECCION",
  "NOMBRE AULA",
  "DOCUMENTO",
  "NOMBRES"
];

function normalizar(s=""){
  return String(s)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g,"")
    .toUpperCase()
    .replace(/Ñ/g,"N")
    .replace(/[^A-Z0-9 ]/g," ")
    .replace(/\s+/g," ")
    .trim();
}

function escapeHtml(v){
  return String(v ?? "")
    .replaceAll("&","&amp;")
    .replaceAll("<","&lt;")
    .replaceAll(">","&gt;")
    .replaceAll('"',"&quot;")
    .replaceAll("'","&#039;");
}

async function leerExcel(file){
  const buffer = await file.arrayBuffer();
  const wb = XLSX.read(buffer,{type:"array"});
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws,{defval:""});
}

function mapearHeaders(row){
  const mapa = {};
  Object.keys(row).forEach(k => {
    mapa[normalizar(k)] = k;
  });
  return mapa;
}

function validarEstructura(rows){
  if(!rows.length){
    throw new Error("El archivo no contiene registros.");
  }

  const mapa = mapearHeaders(rows[0]);
  const faltantes = CAMPOS.filter(c => !mapa[normalizar(c)]);

  if(faltantes.length){
    throw new Error("Faltan columnas: " + faltantes.join(", "));
  }

  return mapa;
}

function esTimestamp(linea){
  return /^\s*\d{1,2}:\d{2}:\d{2}(?:\.\d+)?\s*,\s*\d{1,2}:\d{2}:\d{2}/.test(linea);
}

function extraerNombreLinea(linea){
  let s = String(linea || "").trim();

  if(!s || esTimestamp(s) || !s.includes(":")){
    return "";
  }

  const izquierda = s.split(":")[0].trim();
  const n = normalizar(izquierda);

  if(["SEDE","TUTORA","TUTOR","PROFESOR","PROFE"].includes(n)){
    return "";
  }

  const tokens = izquierda.split(/\s+/).filter(Boolean);

  if(
    tokens.length >= 2 &&
    /^[A-Za-zÁÉÍÓÚÜÑáéíóúüñ .'-]+$/.test(izquierda)
  ){
    return izquierda;
  }

  return "";
}

async function leerMd(file){
  const txt = await file.text();

  const nombres = txt
    .split(/\r?\n/)
    .map(extraerNombreLinea)
    .filter(Boolean);

  return [
    ...new Map(
      nombres.map(n => [normalizar(n), n])
    ).values()
  ];
}

function similitudNombre(a,b){
  a = normalizar(a);
  b = normalizar(b);

  if(!a || !b) return 0;
  if(a === b) return 1;

  const sa = new Set(a.split(" ").filter(Boolean));
  const sb = new Set(b.split(" ").filter(Boolean));

  const comunes = [...sa].filter(x => sb.has(x)).length;

  if(comunes < 2) return 0;

  const containment =
    comunes / Math.max(1,Math.min(sa.size,sb.size));

  const union =
    new Set([...sa,...sb]).size;

  const jaccard =
    comunes / Math.max(1,union);

  return Math.max(containment,jaccard);
}

function gradoNumero(valor){
  const s = normalizar(valor);

  // Regla especial solicitada:
  // "1° Y 2° GRADO" debe contabilizarse como 2.º Secundaria.
  if(
    s.includes("1 Y 2 GRADO") ||
    s.includes("1 Y 2") ||
    s.includes("1RO Y 2DO") ||
    s.includes("1ERO Y 2DO")
  ){
    return 2;
  }

  // Primer Año = 1.º Secundaria
  if(
    s === "1" ||
    s.includes("1RO") ||
    s.includes("1ERO") ||
    s.includes("PRIMERO") ||
    s.includes("PRIMER ANO")
  ){
    return 1;
  }

  // 2.º Secundaria
  if(
    s === "2" ||
    s.includes("2DO") ||
    s.includes("SEGUNDO") ||
    s.includes("2 SECUNDARIA")
  ){
    return 2;
  }

  return 0;
}

function activarProcesar(){
  $("procesarBtn").disabled = !(padron.length && nombresMd.length);
}

$("padronFile").addEventListener("change", async e => {
  try{
    const file = e.target.files[0];
    if(!file) return;

    const rows = await leerExcel(file);
    const mapa = validarEstructura(rows);

    padron = rows
      .map(r => ({
        SEDE: r[mapa["SEDE"]],
        NIVEL: r[mapa["NIVEL"]],
        TURNO: r[mapa["TURNO"]],
        TIPO_GRADO: r[mapa["TIPO GRADO"]],
        GRADO: r[mapa["GRADO"]],
        SECCION: r[mapa["SECCION"]],
        NOMBRE_AULA: r[mapa["NOMBRE AULA"]],
        DOCUMENTO: r[mapa["DOCUMENTO"]],
        NOMBRES: r[mapa["NOMBRES"]]
      }))
      .filter(r => String(r.NOMBRES || "").trim());

    $("padronStatus").textContent =
      `Padrón válido: ${padron.length} alumnos cargados.`;

    $("padronStatus").className = "status ok";
    activarProcesar();

  }catch(err){
    padron = [];
    $("padronStatus").textContent = "Error: " + err.message;
    $("padronStatus").className = "status error";
    activarProcesar();
  }
});

$("mdFile").addEventListener("change", async e => {
  try{
    const file = e.target.files[0];
    if(!file) return;

    nombresMd = await leerMd(file);

    if(!nombresMd.length){
      throw new Error("No se detectaron nombres.");
    }

    $("mdStatus").textContent =
      `Archivo leído: ${nombresMd.length} nombres únicos detectados.`;

    $("mdStatus").className = "status ok";
    activarProcesar();

  }catch(err){
    nombresMd = [];
    $("mdStatus").textContent = "Error: " + err.message;
    $("mdStatus").className = "status error";
    activarProcesar();
  }
});

function mejorCoincidencia(nombre,candidatos,umbral){
  const objetivo = normalizar(nombre);

  const exacta = candidatos.find(
    r => normalizar(r.NOMBRES) === objetivo
  );

  if(exacta){
    return {row:exacta,score:1};
  }

  let mejor = null;
  let score = 0;

  for(const r of candidatos){
    const s = similitudNombre(nombre,r.NOMBRES);

    if(s > score){
      score = s;
      mejor = r;
    }
  }

  return mejor && score >= umbral
    ? {row:mejor,score}
    : null;
}

function generarResultadoPorSede(estados){
  const mapa = new Map();

  estados
    .filter(x =>
      x.estado === "Ubicado" &&
      (x.gradoNumero === 1 || x.gradoNumero === 2)
    )
    .forEach(x => {
      if(!mapa.has(x.sede)){
        mapa.set(x.sede,{
          sede:x.sede,
          primero:0,
          segundo:0,
          total:0
        });
      }

      const fila = mapa.get(x.sede);

      if(x.gradoNumero === 1){
        fila.primero++;
      }

      if(x.gradoNumero === 2){
        fila.segundo++;
      }

      fila.total++;
    });

  return [...mapa.values()]
    .sort((a,b) => a.sede.localeCompare(b.sede,"es"));
}

$("procesarBtn").addEventListener("click", () => {
  const umbral = Number($("umbral").value);
  const soloSec = $("soloSecundaria").checked;

  let candidatos = [...padron];

  if(soloSec){
    candidatos = candidatos.filter(r => {
      const nivel = normalizar(r.NIVEL);
      return nivel.includes("SECUNDARIA") || nivel === "SEC";
    });
  }

  const estados = [];
  const noUbicados = [];

  for(const nombre of nombresMd){
    const match = mejorCoincidencia(nombre,candidatos,umbral);

    if(!match){
      estados.push({
        nombre,
        estado:"No ubicado",
        sede:"",
        grado:"",
        gradoNumero:0
      });

      noUbicados.push(nombre);
      continue;
    }

    const sede =
      String(match.row.SEDE || "").trim() || "SIN SEDE";

    const gradoNum =
      gradoNumero(match.row.GRADO);

    const gradoTexto =
      gradoNum === 1
        ? "1.º SECUNDARIA"
        : gradoNum === 2
          ? "2.º SECUNDARIA"
          : String(match.row.GRADO || "");

    estados.push({
      nombre,
      estado:"Ubicado",
      sede,
      grado:gradoTexto,
      gradoNumero:gradoNum,
      nombrePadron:match.row.NOMBRES,
      score:match.score
    });
  }

  estadosActuales = estados;
  noUbicadosActual = noUbicados;

  resultadoActual =
    generarResultadoPorSede(estadosActuales);

  renderNombresUnicos(estadosActuales);
  renderResultado(resultadoActual);
  renderNoUbicados(noUbicadosActual);

  $("kpiUnicos").textContent = nombresMd.length;
  $("kpiUbicados").textContent =
    estadosActuales.filter(x => x.estado === "Ubicado").length;
  $("kpiNoUbicados").textContent = noUbicadosActual.length;
  $("kpiSedes").textContent = resultadoActual.length;

  [
    "resumenSection",
    "nombresUnicosSection",
    "resultadoSection"
  ].forEach(id =>
    $(id).classList.remove("hidden")
  );

  $("noUbicadosSection").classList.toggle(
    "hidden",
    noUbicadosActual.length === 0
  );

  $("exportarBtn").disabled = false;
});

function renderNombresUnicos(datos){
  const tbody =
    $("nombresUnicosTabla").querySelector("tbody");

  tbody.innerHTML = "";

  datos.forEach((r,i) => {
    const ubicado = r.estado === "Ubicado";

    const tr = document.createElement("tr");

    tr.innerHTML = `
      <td>${i+1}</td>
      <td>${escapeHtml(r.nombre)}</td>
      <td>
        <span class="estado ${ubicado ? "ubicado" : "no-ubicado"}">
          ${escapeHtml(r.estado)}
        </span>
      </td>
      <td>${escapeHtml(r.sede || "—")}</td>
      <td>${escapeHtml(r.grado || "—")}</td>
    `;

    tbody.appendChild(tr);
  });
}

function renderResultado(datos){
  const tbody =
    $("resultadoTabla").querySelector("tbody");

  tbody.innerHTML = "";

  let total1 = 0;
  let total2 = 0;
  let totalGeneral = 0;

  datos.forEach((r,i) => {
    total1 += r.primero;
    total2 += r.segundo;
    totalGeneral += r.total;

    const tr = document.createElement("tr");

    tr.innerHTML = `
      <td>${i+1}</td>
      <td><strong>${escapeHtml(r.sede)}</strong></td>
      <td><strong>${r.primero}</strong></td>
      <td><strong>${r.segundo}</strong></td>
      <td><strong>${r.total}</strong></td>
    `;

    tbody.appendChild(tr);
  });

  $("total1").textContent = total1;
  $("total2").textContent = total2;
  $("totalGeneral").textContent = totalGeneral;
}

function renderNoUbicados(datos){
  const cont = $("noUbicados");
  cont.innerHTML = "";

  datos.forEach(n => {
    const chip = document.createElement("span");
    chip.className = "chip";
    chip.textContent = n;
    cont.appendChild(chip);
  });
}

$("buscarNombre").addEventListener("input", e => {
  const q = normalizar(e.target.value);

  const filtrados = !q
    ? estadosActuales
    : estadosActuales.filter(r =>
        normalizar(r.nombre).includes(q) ||
        normalizar(r.sede).includes(q) ||
        normalizar(r.grado).includes(q) ||
        normalizar(r.estado).includes(q)
      );

  renderNombresUnicos(filtrados);
});

$("buscarSede").addEventListener("input", e => {
  const q = normalizar(e.target.value);

  const filtrados = !q
    ? resultadoActual
    : resultadoActual.filter(r =>
        normalizar(r.sede).includes(q)
      );

  renderResultado(filtrados);
});

$("limpiarBtn").addEventListener("click", () => {
  padron = [];
  nombresMd = [];
  estadosActuales = [];
  resultadoActual = [];
  noUbicadosActual = [];

  $("padronFile").value = "";
  $("mdFile").value = "";

  $("padronStatus").textContent = "Sin archivo cargado.";
  $("mdStatus").textContent = "Sin archivo cargado.";

  $("padronStatus").className = "status";
  $("mdStatus").className = "status";

  [
    "resumenSection",
    "nombresUnicosSection",
    "resultadoSection",
    "noUbicadosSection"
  ].forEach(id =>
    $(id).classList.add("hidden")
  );

  $("procesarBtn").disabled = true;
  $("exportarBtn").disabled = true;
});

$("exportarBtn").addEventListener("click", () => {
  const resumen =
    resultadoActual.map((r,i) => ({
      "N.°":i+1,
      "SEDE":r.sede,
      "1.º SECUNDARIA":r.primero,
      "2.º SECUNDARIA":r.segundo,
      "CANTIDAD TOTAL":r.total
    }));

  const total1 =
    resultadoActual.reduce((s,r) => s+r.primero,0);

  const total2 =
    resultadoActual.reduce((s,r) => s+r.segundo,0);

  const totalGeneral =
    resultadoActual.reduce((s,r) => s+r.total,0);

  resumen.push({
    "N.°":"",
    "SEDE":"TOTAL GENERAL",
    "1.º SECUNDARIA":total1,
    "2.º SECUNDARIA":total2,
    "CANTIDAD TOTAL":totalGeneral
  });

  const nombres =
    estadosActuales.map((r,i) => ({
      "N.°":i+1,
      "NOMBRE DETECTADO":r.nombre,
      "ESTADO":r.estado,
      "SEDE":r.sede,
      "GRADO":r.grado
    }));

  const noUbicados =
    noUbicadosActual.map(n => ({
      "NO UBICADO":n
    }));

  const wb =
    XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.json_to_sheet(resumen),
    "Resultado por sede"
  );

  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.json_to_sheet(nombres),
    "Nombres únicos"
  );

  if(noUbicados.length){
    XLSX.utils.book_append_sheet(
      wb,
      XLSX.utils.json_to_sheet(noUbicados),
      "No ubicados"
    );
  }

  XLSX.writeFile(
    wb,
    "sectorizacion_alumnos.xlsx"
  );
});
