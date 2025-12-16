function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('AMPU');

  // Submenú: Herramientas de búsqueda
  const menuBusqueda = ui.createMenu('Herramientas de búsqueda');
  menuBusqueda
    .addItem('Nomenclador AAARBA', 'showBuscadorAAARBA');
  menu.addSubMenu(menuBusqueda);

  // Submenú: Procesamiento de datos
  const menuProcesamiento = ui.createMenu('Procesamiento de datos');
  menuProcesamiento
    .addItem('Normalizar base de datos', 'generarResumenPorCirugia')
    .addItem('Dashboard', 'generarDashboardDatos')
    .addItem('Colorear', 'showSidebar');
  menu.addSubMenu(menuProcesamiento);

  // Submenú: Cálculos
  const menuCalculos = ui.createMenu('Cálculos');
  menuCalculos
    .addItem('Calculadora AAARBA', 'Calculadora');
  menu.addSubMenu(menuCalculos);

  // Agregar todo el menú a la UI
  menu.addToUi();
}


// Abre la nueva barra lateral de búsqueda
function showBuscadorAAARBA() {
  var html = HtmlService
    .createHtmlOutputFromFile('modalDialog')
    .setTitle('BuscadorAAARBA');
  SpreadsheetApp.getUi().showSidebar(html);
}

// Calculadora AAARBA
function Calculadora() {
  var html = HtmlService.createHtmlOutputFromFile('calculadoraAAARBA')
    .setTitle('Calculadora Honorarios');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
Calcula totales y OP para dos combinaciones
@param {string[]} codesInit
@param {boolean[]} facInit
@param {boolean} opInit
@param {string[]} codesProp
@param {boolean[]} facProp
@param {boolean} opProp
@return {{inicial:{sum:number,op:number},propuesto:{sum:number,op:number},diferencia:number}}
*/
function calcular(codesInit, facInit, opInit, codesProp, facProp, opProp) {
  var ss = SpreadsheetApp.openById('1JpCxL0fbFg5wB5G92HhUaWtw1SoSEw264Lh4w6IFs7s');
  var sheet = ss.getSheetByName('Niveles');
  var data = sheet.getRange('A2:D56').getValues();
  var tarifaMap = {};
  data.forEach(function(r) {
    tarifaMap[r[0]] = r[3];
  });
  function compute(codes, facs, includeOp) {
    var subtotal = 0;
    codes.forEach(function(code, i) {
      if (code && tarifaMap[code] != null) {
        var factor = facs[i] ? 0.5 : 1;
        subtotal += tarifaMap[code] * factor;
      }
    });
    var opValue = 0;
    if (includeOp) {
      opValue = subtotal * 0.3;
      subtotal += opValue;
    }
    return { sum: subtotal, op: opValue };
  }
  var init = compute(codesInit, facInit, opInit);
  var prop = compute(codesProp, facProp, opProp);
  return { inicial: init, propuesto: prop, diferencia: init.sum - prop.sum };
}

// Búsqueda en la hoja 'datos'
function searchDatabase(searchTerm) {
  var ss = SpreadsheetApp.openById('1R9D3XJ9FiCaA3YEZtxB-FtS6JaQRhFaq2CyUGQq0yZ8');
  var sheet = ss.getSheetByName('datos');
  var data = sheet.getDataRange().getValues();
  searchTerm = removeAccents(searchTerm.toLowerCase());
  var results = [];
  results.push(data[0]); // encabezados
  for (var i = 1; i < data.length; i++) {
    for (var j = 0; j < data[i].length; j++) {
      var cellValue = removeAccents(data[i][j].toString().toLowerCase());
      if (cellValue.includes(searchTerm)) {
        results.push(data[i]);
        break;
      }
    }
  }
  return results;
}

// Elimina acentos
function removeAccents(str) {
  return str.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

// Barra lateral para 'Colorear'
function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Page')
    .setTitle('Find and Highlight');
  SpreadsheetApp.getUi().showSidebar(html);
}

// Resalta texto en color
function findAndHighlight(values, color) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var searchValues = (values || '').split(/\s+/).map(function(item) { return item.trim(); }).filter(Boolean);
  if (!searchValues.length) return "Sin valores para buscar";
  var searchSet = new Set(searchValues);
  for (var i = 0; i < data.length; i++) {
    for (var j = 0; j < data[i].length; j++) {
      if (searchSet.has(data[i][j].toString())) {
        sheet.getRange(i+1,1,1,data[i].length).setFontColor(color);
        break;
      }
    }
  }
  return "Resaltado con éxito";
}

// Resalta filas completas
function findAndHighlightRows(values, color) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var searchValues = (values || '').split(/\s+/).map(function(item) { return item.trim(); }).filter(Boolean);
  if (!searchValues.length) return "Sin valores para buscar";
  var searchSet = new Set(searchValues);
  for (var i = 0; i < data.length; i++) {
    for (var j = 0; j < data[i].length; j++) {
      if (searchSet.has(data[i][j].toString())) {
        sheet.getRange(i+1,1,1,data[i].length).setBackground(color);
        break;
      }
    }
  }
  return "Resaltado con éxito";
}

/** ==================== UTILIDADES ==================== **/
function limpiarValor(valor) {
  if (typeof valor === "number") return valor;
  if (typeof valor === "string") {
    return Number(
      valor.replace(/\$/g, "").replace(/\./g, "").replace(",", ".").trim()
    ) || 0;
  }
  return 0;
}

// Normaliza texto: mayúsculas, quita acentos, signos y espacios extra
function normalizarTexto(s) {
  return (s || "")
    .toString()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "") // quita tildes
    .replace(/[^A-Z0-9\s]/gi, " ") // reemplaza signos por espacio
    .toUpperCase()
    .replace(/\s+/g, " ") // colapsa espacios
    .trim();
}

/** ==================== MAPA DESCRIPCION -> ESPECIALIDAD ==================== **/
function cargarMapaEspecialidadesPorDescripcion() {
  const hojaNom = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Nomenclador");
  const datosNom = hojaNom.getDataRange().getValues();
  const mapa = new Map();
  const lista = []; // para búsqueda flexible

  // Nomenclador: Col A = Especialidad, Col C = Descripción
  for (let i = 1; i < datosNom.length; i++) {
    const especialidad = (datosNom[i][0] || "").toString().trim();
    const descripcion = (datosNom[i][2] || "").toString().trim();
    if (!descripcion) continue;

    const descNorm = normalizarTexto(descripcion);
    if (!mapa.has(descNorm)) mapa.set(descNorm, especialidad);
    lista.push([descNorm, especialidad]);
  }

  return { mapa, lista };
}

/** ==================== RESUMEN POR CIRUGIA ==================== **/
function generarResumenPorCirugia() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaOriginal = ss.getSheetByName("BD");
  const datos = hojaOriginal.getDataRange().getValues();
  const encabezadosOriginales = datos.shift();

  const encabezados = encabezadosOriginales.map(h => h.toString().trim().toUpperCase());

  const idx = {
    boleta: encabezados.indexOf("BOLETA"),
    fecha: encabezados.indexOf("FECHA"),
    afiliado: encabezados.indexOf("AFILIADO"),
    nombre: encabezados.indexOf("NOMBRE AFILIADO"),
    sanatorio: encabezados.indexOf("SANATORIO"),
    anestesista: encabezados.indexOf("ANESTESISTA"),
    nivel: encabezados.indexOf("NIVEL"),
    descripcion: encabezados.indexOf("DESCRIPCION"),
    valor: encabezados.indexOf("VALOR"),
    pq: encabezados.indexOf("PQ"),
    practica: encabezados.indexOf("PRACTICA")
  };

  const requeridos = ["boleta","fecha","afiliado","nombre","sanatorio","anestesista","nivel","descripcion","valor","pq"];
  for (const k of requeridos) {
    if (idx[k] === -1) throw new Error("No se encontró la columna requerida: " + k);
  }

  const prioridadNiveles = [
    "EV","EP","EO","EN","EM","EL","EK","EJ","EI","EH","EG","EF","EE","ED","EC","EB","EA",
    "MI","MF","ME","MD","MB","OP","UF","UO","US","UD"
  ];
  const nivelesValidosParaEspecialidad = [
    "EP","EO","EN","EM","EL","EK","EJ","EI","EH","EG","EF","EE","ED","EC","EB","EA",
    "MI","MF","ME","MD","MB"
  ];
  const getPrioridad = (nivel) => {
    const cod = (nivel || "").substring(0,2).toUpperCase();
    const i = prioridadNiveles.indexOf(cod);
    return i === -1 ? 999 : i;
  };

  const { mapa: mapaDescExacto, lista: listaDesc } = cargarMapaEspecialidadesPorDescripcion();
  const agrupados = [];

  for (let fila of datos) {
    const fechaObj = new Date(fila[idx.fecha]);
    const fechaTexto = Utilities.formatDate(fechaObj, Session.getScriptTimeZone(), "dd/MM/yyyy");

    const clave = [
      fila[idx.boleta],
      fechaTexto,
      fila[idx.afiliado],
      fila[idx.anestesista]
    ].join("|");

    let entrada = agrupados.find(e => e.clave === clave);
    if (!entrada) {
      entrada = {
        clave,
        boleta: fila[idx.boleta],
        fecha: fechaTexto,
        afiliado: fila[idx.afiliado],
        nombre: fila[idx.nombre],
        sanatorio: fila[idx.sanatorio],
        anestesista: fila[idx.anestesista],
        pq: fila[idx.pq],
        niveles: [],
        descripciones: [],
        total: 0,
        boletas: new Set()
      };
      agrupados.push(entrada);
    }

    entrada.total += limpiarValor(fila[idx.valor]);
    entrada.niveles.push(fila[idx.nivel]);
    entrada.descripciones.push(fila[idx.descripcion]);
    entrada.boletas.add(fila[idx.boleta]);
  }

  agrupados.sort((a, b) => b.niveles.length - a.niveles.length);

  for (let entrada of agrupados) {
    const combinados = entrada.niveles.map((nivel, i) => ({
      nivel,
      descripcion: entrada.descripciones[i],
      prioridad: getPrioridad(nivel)
    })).sort((a, b) => a.prioridad - b.prioridad);

    entrada.niveles = combinados.map(e => e.nivel);
    entrada.descripciones = combinados.map(e => e.descripcion);
  }

  const hojaResumen = ss.getSheetByName("Resumen por Cirugía") || ss.insertSheet("Resumen por Cirugía");
  hojaResumen.clearContents();

  const salida = [];
  let maxItems = 0;

  for (let entrada of agrupados) {
    const niveles = entrada.niveles;
    const descripciones = entrada.descripciones;
    const cantidadItems = niveles.length;
    if (cantidadItems > maxItems) maxItems = cantidadItems;

    const nivelEspecial = niveles.find(niv =>
      ["OP","UF","UO","US","UN","UD"].includes((niv || "").substring(0,2).toUpperCase())
    ) || "";

    const candidatos = niveles
      .map((nivel, i) => ({
        nivel,
        descripcion: descripciones[i],
        prioridad: getPrioridad(nivel)
      }))
      .filter(e => nivelesValidosParaEspecialidad.includes((e.nivel || "").substring(0,2).toUpperCase()))
      .sort((a, b) => a.prioridad - b.prioridad);

    let especialidad = "";
    if (candidatos.length > 0) {
      const descPrincipalNorm = normalizarTexto(candidatos[0].descripcion);
      if (mapaDescExacto.has(descPrincipalNorm)) {
        especialidad = mapaDescExacto.get(descPrincipalNorm);
      } else {
        const MIN_LEN = 10;
        if (descPrincipalNorm.length >= MIN_LEN) {
          for (let i = 0; i < listaDesc.length; i++) {
            const [descNomNorm, especNom] = listaDesc[i];
            if (descNomNorm.length < MIN_LEN) continue;

            if (descPrincipalNorm.includes(descNomNorm) || descNomNorm.includes(descPrincipalNorm)) {
              especialidad = especNom;
              break;
            }
          }
        }
        if (!especialidad) {
          especialidad = "NO ENCONTRADO";
        }
      }
    }

    // 👉 Agregamos la columna "AMPU" al inicio de cada fila (por ahora vacío)
    const filaBase = [
      "", // columna AMPU
      entrada.boleta,
      entrada.fecha,
      entrada.nombre,
      entrada.sanatorio,
      entrada.anestesista,
      entrada.pq,
      cantidadItems,
      entrada.total,
      nivelEspecial,
      especialidad
    ];

    const nivelesCompletos = niveles.concat(Array(maxItems - niveles.length).fill(""));
    const fila = filaBase.concat(nivelesCompletos);
    fila.push(...descripciones);

    salida.push(fila);
  }

  // 👉 Encabezados: agregamos "AMPU" al inicio
  const encabezadoFinal = [
    "AMPU",
    "BOLETA","FECHA","NOMBRE AFILIADO","SANATORIO","ANESTESISTA",
    "PQ","CANTIDAD_ITEMS","TOTAL PROCEDIMIENTO",
    "NIVEL ESPECIAL","ESPECIALIDAD"
  ];
  for (let i = 1; i <= maxItems; i++) encabezadoFinal.push(`NIVEL_${i}`);
  for (let i = 1; i <= maxItems; i++) encabezadoFinal.push(`DESCRIPCION_${i}`);

  const columnasEsperadas = encabezadoFinal.length;
  const salidaNormalizada = salida.map(fila => {
    const faltan = columnasEsperadas - fila.length;
    return fila.concat(Array(faltan).fill(""));
  });

  hojaResumen.getRange(1, 1, 1, columnasEsperadas).setValues([encabezadoFinal]);
  hojaResumen.getRange(2, 1, salidaNormalizada.length, columnasEsperadas).setValues(salidaNormalizada);

  const colTotal = encabezadoFinal.indexOf("TOTAL PROCEDIMIENTO") + 1;
  hojaResumen.getRange(2, colTotal, salida.length).setNumberFormat("$#,##0.00");
}

function generarDashboardDatos() { 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaResumen = ss.getSheetByName("Resumen por Cirugía");
  const hojaDashboard = ss.getSheetByName("Dashboard") || ss.insertSheet("Dashboard");
  const hojaBD = ss.getSheetByName("BD");
  hojaDashboard.clear();

  const datos = hojaResumen.getDataRange().getValues();
  const encabezados = datos[0];
  const filas = datos.slice(1);

  const idx = {
    descripcion: encabezados.findIndex(h => h.toString().toUpperCase().startsWith("DESCRIPCION_")),
    nivel: encabezados.findIndex(h => h.toString().toUpperCase().startsWith("NIVEL_")),
    anestesista: encabezados.indexOf("ANESTESISTA"),
    procedimientoTotal: encabezados.indexOf("TOTAL PROCEDIMIENTO"),
    sanatorio: encabezados.indexOf("SANATORIO"),
    boleta: encabezados.indexOf("BOLETA")
  };

  const parseMonto = (valor) => limpiarValor(valor);

  const contarFrecuencias = (valores) => {
    const mapa = {};
    valores.forEach((v) => {
      if (!v || v === "") return;
      if (!mapa[v]) mapa[v] = 0;
      mapa[v]++;
    });
    return Object.entries(mapa).sort((a, b) => b[1] - a[1]);
  };

  const flattenCols = (rows, prefix, excluir = []) => {
    const cols = encabezados.map((h, i) => h.toString().startsWith(prefix) ? i : -1).filter(i => i !== -1);
    const valores = [];
    rows.forEach(r => {
      cols.forEach(i => {
        const val = r[i];
        if (val && !excluir.some(ex => (val+"").toLowerCase().includes(ex.toLowerCase()))) {
          valores.push(val);
        }
      });
    });
    return valores;
  };

  const agruparConMonto = (rows, campo) => {
    const mapa = {};
    rows.forEach(r => {
      const clave = (r[campo] || "").toString().trim();
      const monto = parseMonto(r[idx.procedimientoTotal]);
      if (!clave) return;
      if (!mapa[clave]) mapa[clave] = { cantidad: 0, monto: 0 };
      mapa[clave].cantidad++;
      mapa[clave].monto += monto;
    });
    return Object.entries(mapa).sort((a, b) => b[1].cantidad - a[1].cantidad);
  };

  const escribirTablaConPorcentaje = (titulo, data, filaInicio, totalBase, chartType = Charts.ChartType.BAR, isDoughnut = false) => {
    hojaDashboard.getRange(filaInicio, 1, 1, 3).setValues([[titulo, "Cantidad", "% sobre Total"]]);
    const tabla = data.map(([k, cantidad]) => [k, cantidad, cantidad / totalBase]);
    hojaDashboard.getRange(filaInicio + 1, 1, tabla.length, 3).setValues(tabla);
    hojaDashboard.getRange(filaInicio + 1, 3, tabla.length, 1).setNumberFormat("0.00%");
    aplicarFormatoTabla(hojaDashboard, filaInicio, 3, tabla.length);

    const chartRange = hojaDashboard.getRange(filaInicio + 1, 1, tabla.length, 2);
    let chartBuilder = hojaDashboard.newChart()
      .setChartType(chartType)
      .addRange(chartRange)
      .setPosition(filaInicio, 5, 0, 0)
      .setOption("title", titulo)
      .setOption("legend", { position: "right" })
      .setOption("is3D", isDoughnut)
      .setOption("pieSliceText", "label");

    hojaDashboard.insertChart(chartBuilder.build());
  };

  const totalBoletas = new Set(filas.map(r => r[idx.boleta])).size;

  const topProcedimientos = contarFrecuencias(
    flattenCols(filas, "DESCRIPCION_", ["Evaluación anestésica", "urgencia", "programada", "domingo", "sábado", "feriado", "diurna", "nocturna"])
  ).slice(0, 10);

  const topNiveles = contarFrecuencias(
    flattenCols(filas, "NIVEL_", ["EV100", "UD100", "UF100", "UN100", "UO100", "US100", "OP30"])
  ).slice(0, 10);

  const rankingAnestesistas = agruparConMonto(filas, idx.anestesista).slice(0, 10);

  const urgenciasYProgramadas = contarFrecuencias(
    flattenCols(filas, "DESCRIPCION_", ["Evaluación anestésica"]).filter(v => /urgencia|programada|domingo|sábado|feriado|diurna|nocturna/i.test(v))
  );

  escribirTablaConPorcentaje("Top 10 Procedimientos Más Frecuentes", topProcedimientos, 1, totalBoletas);
  escribirTablaConPorcentaje("Top 10 Niveles Facturados", topNiveles, 20, totalBoletas);

  hojaDashboard.getRange(39, 1, 1, 4).setValues([["Ranking de Anestesistas (Cantidad)", "Cantidad", "Monto", "% sobre Total"]]);
  const totalFacturado = filas.reduce((acum, r) => acum + parseMonto(r[idx.procedimientoTotal]), 0);
  const tablaAnestesistas = rankingAnestesistas.map(([k, v]) => [k, v.cantidad, v.monto, v.monto / totalFacturado]);
  hojaDashboard.getRange(40, 1, tablaAnestesistas.length, 4).setValues(tablaAnestesistas);
  hojaDashboard.getRange(40, 3, tablaAnestesistas.length, 1).setNumberFormat("$#,##0.00");
  hojaDashboard.getRange(40, 4, tablaAnestesistas.length, 1).setNumberFormat("0.00%");
  aplicarFormatoTabla(hojaDashboard, 39, 4, tablaAnestesistas.length);

  const chartAnestesistas = hojaDashboard.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(hojaDashboard.getRange(40, 1, tablaAnestesistas.length, 2))
    .setPosition(39, 5, 0, 0)
    .setOption("title", "Ranking de Anestesistas")
    .setOption("legend", { position: "right" })
    .build();
  hojaDashboard.insertChart(chartAnestesistas);

  escribirTablaConPorcentaje("Conteo de Operaciones Programadas y Urgencias", urgenciasYProgramadas, 60, totalBoletas);

    /*** === AUDITORÍA POR COLORES: ahora BOLETA está en B y los colores en C === ***/
  const lastRowRes = hojaResumen.getLastRow();
  if (lastRowRes >= 2) {
    const valores = hojaResumen.getRange(1, 2, lastRowRes, 2).getValues(); // B (boleta) y C (coloreada)
    const fontColC = hojaResumen.getRange(1, 3, lastRowRes, 1).getFontColors();
    const backColC = hojaResumen.getRange(1, 3, lastRowRes, 1).getBackgrounds();

    const mapeo = new Map(); // boleta -> estado
    for (let i = 2; i <= lastRowRes; i++) {
      const boleta = valores[i - 1][0];                     // Col B
      const texto = (fontColC[i - 1][0] || "").toLowerCase(); // fuente Col C
      const fondo = (backColC[i - 1][0] || "").toLowerCase(); // fondo Col C
      if (!boleta || mapeo.has(boleta)) continue;

      if (texto === "#0000ff" && fondo === "#ffffff") {
        mapeo.set(boleta, "Sin Observaciones");
      } else if (texto === "#000000" && fondo === "#ffff00") {
        mapeo.set(boleta, "Auditado sin Observaciones");
      } else if (texto === "#ff0000" && fondo === "#ffff00") {
        mapeo.set(boleta, "Auditado con Observaciones");
      }
    }

    // Contadores
    const resumenColores = {
      "Auditado sin Observaciones": 0,
      "Sin Observaciones": 0,
      "Auditado con Observaciones": 0
    };
    for (let tipo of mapeo.values()) {
      if (resumenColores.hasOwnProperty(tipo)) resumenColores[tipo]++;
    }

    const totalAuditado = mapeo.size;
    const pendienteAuditoria = totalBoletas - totalAuditado;

    // Orden fijo para colores estables
    const ordenCategorias = [
      "Auditado sin Observaciones",
      "Sin Observaciones",
      "Auditado con Observaciones",
      "Pendiente de Auditoría"
    ];

    const tablaAuditoria = ordenCategorias.map(cat => {
      const valor = (cat === "Pendiente de Auditoría") ? pendienteAuditoria : (resumenColores[cat] || 0);
      return [cat, valor, totalBoletas ? valor / totalBoletas : 0];
    });
    // Fila informativa
    tablaAuditoria.push(["Total Procedimientos", totalBoletas, 1]);

    // Escribir tabla y formato
    hojaDashboard.getRange(81, 1, 1, 3).setValues([["Auditoría por Colores", "Cantidad", "% sobre Total"]]);
    hojaDashboard.getRange(82, 1, tablaAuditoria.length, 3).setValues(tablaAuditoria);
    hojaDashboard.getRange(82, 3, tablaAuditoria.length, 1).setNumberFormat("0.00%");
    aplicarFormatoTabla(hojaDashboard, 81, 3, tablaAuditoria.length);

    // Gráfico donut con colores pastel en el MISMO orden que la tabla (primeras 4 filas)
    const coloresPastel = ["#F9E79F", "#A9CCE3", "#F5B7B1", "#D7DBDD"];
    const chartAuditoria = hojaDashboard.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(hojaDashboard.getRange(82, 1, 4, 2)) // 4 categorías, 2 columnas (Etiqueta, Cantidad)
      .setPosition(81, 5, 0, 0)
      .setOption("title", "Distribución de Auditoría")
      .setOption("pieHole", 0.4) // donut
      .setOption("legend", { position: "right" })
      .setOption("colors", coloresPastel)
      .build();
    hojaDashboard.insertChart(chartAuditoria);
  }


  insertarBotonActualizarAuditoria();
}

// 🎨 APLICA FORMATO DE TABLA
function aplicarFormatoTabla(hoja, filaInicio, columnas, cantidadFilas) {
  const rangoEncabezado = hoja.getRange(filaInicio, 1, 1, columnas);
  rangoEncabezado.setFontWeight("bold");
  rangoEncabezado.setBackground("#D6DBDF");
  rangoEncabezado.setHorizontalAlignment("center");
  rangoEncabezado.setVerticalAlignment("middle");
  rangoEncabezado.setBorder(true, true, true, true, true, true);

  const rangoDatos = hoja.getRange(filaInicio + 1, 1, cantidadFilas, columnas);
  rangoDatos.setBorder(true, true, true, true, true, true);
  rangoDatos.setHorizontalAlignment("center");
  rangoDatos.setVerticalAlignment("middle");
}

// 🔘 INSERCIÓN DE BOTÓN
function insertarBotonActualizarAuditoria() {
  const hojaDashboard = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dashboard");
  const fileId = "1hW4pOgBDohB2MKw4Fnb5p6X0LDTBfGcA"; // ID de la imagen en Drive
  const blob = DriveApp.getFileById(fileId).getBlob();
  hojaDashboard.insertImage(blob, 2, 88).assignScript("actualizarTablaAuditoria").setAltTextTitle("Actualizar Auditoría");
}

function actualizarTablaAuditoria() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaDashboard = ss.getSheetByName("Dashboard");
  const hojaResumen = ss.getSheetByName("Resumen por Cirugía");
  if (!hojaDashboard || !hojaResumen) return;

  const datosResumen = hojaResumen.getDataRange().getValues();
  const encabezados = datosResumen[0];
  const filasResumen = datosResumen.slice(1);
  const idxBoletaResumen = encabezados.indexOf("BOLETA");
  const totalBoletas = new Set(filasResumen.map(r => r[idxBoletaResumen])).size;

  // BOLETA en B (2), colores en C (3)
  const lastRowRes = hojaResumen.getLastRow();
  if (lastRowRes < 2) return;

  const valoresBC = hojaResumen.getRange(1, 2, lastRowRes, 2).getValues(); 
  const fontColC  = hojaResumen.getRange(1, 3, lastRowRes, 1).getFontColors();
  const backColC  = hojaResumen.getRange(1, 3, lastRowRes, 1).getBackgrounds();

  const mapeo = new Map(); // boleta -> estado
  for (let i = 2; i <= lastRowRes; i++) {
    const boleta = valoresBC[i - 1][0]; 
    const texto = (fontColC[i - 1][0] || "").toLowerCase();
    const fondo = (backColC[i - 1][0] || "").toLowerCase();
    if (!boleta || mapeo.has(boleta)) continue;

    if (texto === "#0000ff" && fondo === "#ffffff") {
      mapeo.set(boleta, "Sin Observaciones");
    } else if (texto === "#000000" && fondo === "#ffff00") {
      mapeo.set(boleta, "Auditado sin Observaciones");
    } else if (texto === "#ff0000" && fondo === "#ffff00") {
      mapeo.set(boleta, "Auditado con Observaciones");
    }
  }

  // Contadores
  const resumenColores = {
    "Auditado sin Observaciones": 0,
    "Sin Observaciones": 0,
    "Auditado con Observaciones": 0
  };
  for (let tipo of mapeo.values()) {
    if (resumenColores.hasOwnProperty(tipo)) resumenColores[tipo]++;
  }

  const totalAuditado = mapeo.size;
  const pendienteAuditoria = totalBoletas - totalAuditado;

  // Orden fijo
  const ordenCategorias = [
    "Auditado sin Observaciones",
    "Sin Observaciones",
    "Auditado con Observaciones",
    "Pendiente de Auditoría"
  ];

  const tablaAuditoria = ordenCategorias.map(cat => {
    const valor = (cat === "Pendiente de Auditoría") ? pendienteAuditoria : (resumenColores[cat] || 0);
    return [cat, valor, totalBoletas ? valor / totalBoletas : 0];
  });
  tablaAuditoria.push(["Total Procedimientos", totalBoletas, 1]);

  // 👉 Solo actualizamos tabla
  hojaDashboard.getRange(82, 1, tablaAuditoria.length, 3).setValues(tablaAuditoria);
  hojaDashboard.getRange(82, 3, tablaAuditoria.length, 1).setNumberFormat("0.00%");
}

function SinDebito1() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getActiveRange();
  var lastColumn = sheet.getLastColumn();
  var rows = [];

  // Recorremos todas las celdas seleccionadas
  for (var i = 1; i <= range.getNumRows(); i++) {
    for (var j = 1; j <= range.getNumColumns(); j++) {
      var row = range.getCell(i, j).getRow();
      if (!rows.includes(row)) {
        rows.push(row); // Guardamos filas únicas
      }
    }
  }

  // Pintamos cada fila completa
  rows.forEach(function(row) {
    var rowRange = sheet.getRange(row, 1, 1, lastColumn);
    rowRange.setBackground('#ffff00');
    rowRange.setFontColor('#000000');
  });
}

function ConDebito1() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getActiveRange();
  var lastColumn = sheet.getLastColumn();
  var rows = [];

  // Recorremos cada celda seleccionada
  for (var i = 0; i < range.getNumRows(); i++) {
    var row = range.getCell(i + 1, 1).getRow();
    if (!rows.includes(row)) {
      rows.push(row); // Agregamos filas únicas
    }
  }

  // Pintamos cada fila completa
  rows.forEach(function(row) {
    var rowRange = sheet.getRange(row, 1, 1, lastColumn);
    rowRange.setBackground('#ffff00');
    rowRange.setFontColor('#ff0000');
  });
}

function macroAbrirBuscador() {
  showBuscadorAAARBA();
}
