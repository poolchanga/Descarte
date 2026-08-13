function doGet() {
  // Despliega la interfaz del Login inicialmente
  return HtmlService.createTemplateFromFile('Login')
      .evaluate()
      .setTitle('Sistema de Gestión de Descarte')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Función auxiliar indispensable para incluir archivos HTML aislados (CSS o JS)
function incluir(nombreArchivo) {
  return HtmlService.createHtmlOutputFromFile(nombreArchivo).getContent();
}

// Función para validar el Login desde las pestañas del Sheet
function validarCredenciales(usuario, contrasenia, empresa) {
  try {
    var sheetId = "110fiUgfJizj_JR7RG3B-xnwfC3wx0NAlbAD4kie-oZ4";
    var ss = SpreadsheetApp.openById(sheetId);
    
    // Determinar a qué pestaña ir según la selección del usuario
    var nombrePestana = "ACCESO-" + empresa;
    var sheet = ss.getSheetByName(nombrePestana);
    
    if (!sheet) {
      return { exito: false, mensaje: "Error: No se encontró la pestaña de acceso para " + empresa };
    }
    
    var datos = sheet.getDataRange().getValues();
    
    // Buscar usuario y contraseña (omitimos la cabecera fila 1)
    for (var i = 1; i < datos.length; i++) {
      var userSheet = datos[i][0].toString().trim().toLowerCase();
      var passSheet = datos[i][1].toString().trim();
      var estado = datos[i][2] ? datos[i][2].toString().trim().toUpperCase() : "ACTIVO";
      
      if (userSheet === usuario.trim().toLowerCase() && passSheet === contrasenia.trim()) {
        if (estado === "ACTIVO") {
          return { exito: true, usuario: usuario, empresa: empresa };
        } else {
          return { exito: false, mensaje: "El usuario se encuentra INACTIVO." };
        }
      }
    }
    return { exito: false, mensaje: "Usuario o contraseña incorrectos." };
    
  } catch (error) {
    return { exito: false, mensaje: "Error de conexión: " + error.toString() };
  }
}

// Función que cargará la interfaz principal una vez logueado
function cargarInterfazPrincipal() {
  return HtmlService.createTemplateFromFile('Interfaz').evaluate().getContent();
}

function obtenerDatosReportePorFechas(fechaInicioStr, fechaFinStr) {
  try {
    var sheetId = "110fiUgfJizj_JR7RG3B-xnwfC3wx0NAlbAD4kie-oZ4";
    var ss = SpreadsheetApp.openById(sheetId);
    var sheetLinks = ss.getSheetByName("LINK-FECHA");
    
    if (!sheetLinks) {
      return { exito: false, mensaje: "No se encontró la pestaña 'LINK-FECHA' en el documento maestro." };
    }
    
    var datosLinks = sheetLinks.getDataRange().getValues();
    
    // Estandarizar inputs YYYY-MM-DD a objeto Date
    var parsearInput = function(str) {
      var p = str.split("-");
      return new Date(parseInt(p[0],10), parseInt(p[1],10) - 1, parseInt(p[2],10));
    };
    var dInicio = parsearInput(fechaInicioStr);
    var dFin = parsearInput(fechaFinStr);
    
    // Convertir celdas de Sheets a objeto Date de comparación limpia
    var normalizarFechaCelda = function(celda) {
      if (celda instanceof Date) {
        return new Date(celda.getFullYear(), celda.getMonth(), celda.getDate());
      }
      if (celda) {
        var str = celda.toString().trim();
        var p = str.split("/");
        if (p.length === 3) {
          return new Date(parseInt(p[2],10), parseInt(p[1],10) - 1, parseInt(p[0],10));
        }
      }
      return null;
    };

    var enlacesAProcesar = [];
    
    // 1. Filtrar los enlaces por rango de fecha en la pestaña maestra
    for (var i = 1; i < datosLinks.length; i++) {
      var empresaMatriz = datosLinks[i][0] ? datosLinks[i][0].toString().trim().toUpperCase() : "";
      var fechaObj = normalizarFechaCelda(datosLinks[i][1]);
      var urlLink = datosLinks[i][2] ? datosLinks[i][2].toString().trim() : "";
      
      if (fechaObj && urlLink) {
        if (fechaObj >= dInicio && fechaObj <= dFin) {
          enlacesAProcesar.push({ empresa: empresaMatriz, url: urlLink });
        }
      }
    }
    
    if (enlacesAProcesar.length === 0) {
      return { exito: true, datos: [], mensaje: "No se encontraron enlaces en el rango seleccionado." };
    }
    
    // Objeto temporal para acumular y agrupar las sumas
    var mapaAgrupado = {};
    
    // 2. Procesar dinámicamente cada archivo linkeado (Lectura corrida por Columnas)
    enlacesAProcesar.forEach(function(item) {
      try {
        var match = item.url.match(/\/d\/([a-zA-Z0-9-_]+)/);
        if (!match || !match[1]) return;
        
        var subSs = SpreadsheetApp.openById(match[1]);
        var subSheet = subSs.getSheets()[0]; 
        var subDatos = subSheet.getDataRange().getValues();
        
        // Recorremos TODAS las filas de la hoja en un solo ciclo dinámico
        for (var j = 0; j < subDatos.length; j++) {
          
          // Evaluamos si la fila actual pertenece al formato de estructura LARAMA (Packing arriba)
          var tipoDescarteLarama = subDatos[j][3] ? subDatos[j][3].toString().trim().toUpperCase() : ""; // Columna D
          
          if (tipoDescarteLarama.includes("DESCARTE")) {
            var fRaw = subDatos[j][0]; // Columna A
            var fechaTxt = (fRaw instanceof Date) ? Utilities.formatDate(fRaw, Session.getScriptTimeZone(), "dd/MM/yyyy") : fRaw.toString().trim();
            
            var plantaOrigen = subDatos[j][13] ? subDatos[j][13].toString().trim().toUpperCase() : ""; // Columna N
            var variedad = subDatos[j][2] ? subDatos[j][2].toString().trim().toUpperCase() : ""; // Columna C
            var pesoNeto = parseFloat(subDatos[j][9]) || 0; // Columna J
            
            if (pesoNeto > 0 && fechaTxt.includes("/")) {
              var llavePacking = item.empresa + "|" + fechaTxt + "|" + plantaOrigen + "|" + tipoDescarteLarama + "|" + variedad;
              
              if (!mapaAgrupado[llavePacking]) {
                mapaAgrupado[llavePacking] = {
                  empresa: item.empresa,
                  fecha: fechaTxt,
                  planta: subDatos[j][13] ? subDatos[j][13].toString().trim() : "", 
                  presentacion: subDatos[j][3] ? subDatos[j][3].toString().trim() : "", 
                  variedad: subDatos[j][2] ? subDatos[j][2].toString().trim() : "", 
                  peso: 0
                };
              }
              mapaAgrupado[llavePacking].peso += pesoNeto;
            }
          }
          
          // Evaluamos en la misma fila si cumple con el formato AGA / ARENUVA o bloque inferior de Campo
          var tipoDescarteCampo = subDatos[j][5] ? subDatos[j][5].toString().trim().toUpperCase() : ""; // Columna F
          
          if (tipoDescarteCampo.includes("DESCARTE") || tipoDescarteCampo.includes("GRANEL")) {
            var fRawC = subDatos[j][2]; // Columna C
            var fechaTxtC = (fRawC instanceof Date) ? Utilities.formatDate(fRawC, Session.getScriptTimeZone(), "dd/MM/yyyy") : fRawC.toString().trim();
            
            var plantaOrigenCampo = subDatos[j][21] ? subDatos[j][21].toString().trim().toUpperCase() : ""; // Columna V
            var variedadCampo = subDatos[j][9] ? subDatos[j][9].toString().trim().toUpperCase() : ""; // Columna J
            var pesoNetoCampo = parseFloat(subDatos[j][12]) || 0; // Columna M
            
            if (pesoNetoCampo > 0 && fechaTxtC.includes("/")) {
              var llaveCampo = item.empresa + "|" + fechaTxtC + "|" + plantaOrigenCampo + "|" + tipoDescarteCampo + "|" + variedadCampo;
              
              if (!mapaAgrupado[llaveCampo]) {
                mapaAgrupado[llaveCampo] = {
                  empresa: item.empresa,
                  fecha: fechaTxtC,
                  planta: subDatos[j][21] ? subDatos[j][21].toString().trim() : "", 
                  presentacion: subDatos[j][5] ? subDatos[j][5].toString().trim() : "", 
                  variedad: subDatos[j][9] ? subDatos[j][9].toString().trim() : "", 
                  peso: 0
                };
              }
              mapaAgrupado[llaveCampo].peso += pesoNetoCampo;
            }
          }
        }
        
      } catch (e) {
        console.error("Error en link de " + item.empresa + ": " + e.toString());
      }
    });
    
    // 3. Transformar el mapa agrupado de vuelta a una lista limpia para la web
    var consolidadoGlobal = [];
    for (var clave in mapaAgrupado) {
      consolidadoGlobal.push(mapaAgrupado[clave]);
    }
    
    return { exito: true, datos: consolidadoGlobal };
    
  } catch (err) {
    return { exito: false, mensaje: "Error general en servidor: " + err.toString() };
  }
}

// ==========================================
// CRUCE DE STOCKS (LECTURA VERTICAL DINÁMICA)
// ==========================================

function obtenerCruceStocks(empresaSeleccionada, fechaInicioStr, fechaFinStr) {
  try {
    var sheetId = "110fiUgfJizj_JR7RG3B-xnwfC3wx0NAlbAD4kie-oZ4";
    var ss = SpreadsheetApp.openById(sheetId);
    
    // --- 0. CARGAR LISTA DE PRODUCTOS DESDE BD-PRODUCTOS ---
    var listaProductosBD = [];
    var sheetProductos = ss.getSheetByName("BD-PRODUCTOS");
    if (sheetProductos) {
      var datosProd = sheetProductos.getDataRange().getValues();
      for (var p = 1; p < datosProd.length; p++) {
        var codCelda = datosProd[p][0] ? datosProd[p][0].toString().trim().toUpperCase() : "";
        var descCelda = datosProd[p][1] ? datosProd[p][1].toString().trim().toUpperCase() : "";
        if (descCelda !== "" && codCelda !== "") {
          listaProductosBD.push({ codigo: codCelda, descripcion: descCelda });
        }
      }
    }

    // --- 1. OBTENER ENLACES DE SISPACKING (HOJA LINK-FECHA) ---
    var sheetLinks = ss.getSheetByName("LINK-FECHA");
    if (!sheetLinks) return { exito: false, mensaje: "No se encontró la pestaña 'LINK-FECHA'." };
    
    var datosLinks = sheetLinks.getDataRange().getValues();
    
    var parsearInput = function(str) {
      var p = str.split("-");
      return new Date(parseInt(p[0],10), parseInt(p[1],10) - 1, parseInt(p[2],10));
    };
    var dInicio = parsearInput(fechaInicioStr);
    var dFin = parsearInput(fechaFinStr);
    
    var normalizarFechaCelda = function(celda) {
      if (!celda) return null;
      if (celda instanceof Date) {
        return new Date(celda.getFullYear(), celda.getMonth(), celda.getDate());
      }
      var texto = celda.toString().trim();
      if (texto === "") return null;
      
      if (texto.includes("/")) {
        var p = texto.split("/");
        if (p.length === 3) {
          return new Date(parseInt(p[2],10), parseInt(p[1],10) - 1, parseInt(p[0],10));
        }
      }
      if (texto.includes("-")) {
        var p = texto.split("-");
        if (p.length === 3) {
          return new Date(parseInt(p[0],10), parseInt(p[1],10) - 1, parseInt(p[2],10));
        }
      }
      return null;
    };

    var enlacesAProcesar = [];
    for (var i = 1; i < datosLinks.length; i++) {
      var empMatriz = datosLinks[i][0] ? datosLinks[i][0].toString().trim().toUpperCase() : "";
      var fechaObj = normalizarFechaCelda(datosLinks[i][1]);
      var urlLink = datosLinks[i][2] ? datosLinks[i][2].toString().trim() : "";
      
      if (fechaObj && urlLink && empMatriz === empresaSeleccionada.toUpperCase()) {
        if (fechaObj >= dInicio && fechaObj <= dFin) {
          enlacesAProcesar.push(urlLink);
        }
      }
    }

    // --- 2. ACUMULAR KILOS DE SISPACKING ---
    var mapaSispacking = {};
    enlacesAProcesar.forEach(function(url) {
      try {
        var match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
        if (!match || !match[1]) return;
        var subSs = SpreadsheetApp.openById(match[1]);
        var subSheet = subSs.getSheets()[0];
        var subDatos = subSheet.getDataRange().getValues();
        
        for (var j = 0; j < subDatos.length; j++) {
          var tipoC = subDatos[j][5] ? subDatos[j][5].toString().trim().toUpperCase() : "";
          if (tipoC.includes("DESCARTE") || tipoC.includes("GRANEL")) {
            var fRawC = subDatos[j][2];
            var fObjC = normalizarFechaCelda(fRawC);
            if (fObjC) {
              var fTxtC = Utilities.formatDate(fObjC, Session.getScriptTimeZone(), "dd/MM/yyyy");
              var varC = subDatos[j][9] ? subDatos[j][9].toString().trim().toUpperCase() : "";
              var pesoC = parseFloat(subDatos[j][12]) || 0;
              if (pesoC > 0) {
                var llaveC = fTxtC + "|" + varC + "|CAMPO";
                mapaSispacking[llaveC] = (mapaSispacking[llaveC] || 0) + pesoC;
              }
            }
          }
          
          var tipoP = subDatos[j][3] ? subDatos[j][3].toString().trim().toUpperCase() : "";
          if (tipoP.includes("DESCARTE")) {
            var fRawP = subDatos[j][0];
            var fObjP = normalizarFechaCelda(fRawP);
            if (fObjP) {
              var fTxtP = Utilities.formatDate(fObjP, Session.getScriptTimeZone(), "dd/MM/yyyy");
              var varP = subDatos[j][2] ? subDatos[j][2].toString().trim().toUpperCase() : "";
              var pesoP = parseFloat(subDatos[j][9]) || 0;
              if (pesoP > 0) {
                var llaveP = fTxtP + "|" + varP + "|PACKING";
                mapaSispacking[llaveP] = (mapaSispacking[llaveP] || 0) + pesoP;
              }
            }
          }
        }
      } catch (e) { console.error("Error subarchivo: " + e.toString()); }
    });

    // --- 3. LECTURA VERTICAL DE SAP (COLUMNAS A, B, D, E, G, J) ---
    var nombrePestanaSap = "SAP " + empresaSeleccionada.toUpperCase();
    var sheetSap = ss.getSheetByName(nombrePestanaSap);
    var mapaSap = {};
    var mapaStockFinal = {}; 
    
    if (sheetSap) {
      var datosSap = sheetSap.getDataRange().getValues();
      
      var variedadActual = "";
      var tipoGrupoActual = "";
      var ultimoValorColJ = null;
      var llaveStockActual = "";

      for (var r = 1; r < datosSap.length; r++) {
        var colA = datosSap[r][0] ? datosSap[r][0].toString().trim() : "";
        var colB = datosSap[r][1] ? datosSap[r][1].toString().trim().toUpperCase() : "";
        
        // A) DETECTAR NUEVO BLOQUE VERTICAL / VARIEDAD
        if (colA !== "" || colB !== "") {
          if (colB !== "" && colB !== "DESCRIPCIÓN" && colB !== "DESCRIPCION") {
            
            // GUARDAR EL ÚLTIMO STOCK DEL BLOQUE ANTERIOR ANTES DE INICIAR EL NUEVO
            if (llaveStockActual !== "" && ultimoValorColJ !== null) {
              mapaStockFinal[llaveStockActual] = ultimoValorColJ;
            }

            tipoGrupoActual = colB.includes("PACKING") ? "PACKING" : "CAMPO";
            
            if (colB.includes("ROSITA")) {
              variedadActual = "ROSITA";
            } else if (colB.includes("RAYMI")) {
              variedadActual = "RAYMI BLU";
            } else if (colB.includes("DINA")) {
              variedadActual = "OZBLU® DINA";
            } else if (colB.includes("CAROLINA")) {
              variedadActual = "OZBLU® CAROLINA";
            } else if (colB.includes("JULIETA")) {
              variedadActual = "OZBLU® JULIETA";
            } else if (colB.includes("ANDREA")) {
              variedadActual = "OZBLU® ANDREA";
            } else if (colB.includes("MÁGICA") || colB.includes("MAGICA")) {
              variedadActual = "OZBLU® MÁGICA";
            } else if (colB.includes("OLIVIA")) {
              variedadActual = "OZBLU® OLIVIA";
            } else {
              var limpiar = colB
                .replace("ARÁNDANO FRESCO", "")
                .replace("- DESCARTE CAMPO", "")
                .replace("- DESCARTE PACKING", "")
                .replace("CAMPO", "")
                .replace("PACKING", "")
                .trim();
              if (limpiar) variedadActual = limpiar;
            }

            llaveStockActual = variedadActual + "|" + tipoGrupoActual;
            ultimoValorColJ = null; // Reiniciar para el nuevo bloque
          }
        }
        
        // B) RASTREO DE COLUMNA J (ÍNDICE 9) - CAPTURA INCLUYENDO 0
        var valJ = parseFloat(datosSap[r][9]);
        if (!isNaN(valJ)) {
          ultimoValorColJ = valJ;
        }

        // C) ACUMULAR TABLAS SUPERIORES POR FECHA
        var fRawS = datosSap[r][3]; 
        var docCodigo = datosSap[r][4] ? datosSap[r][4].toString().toUpperCase().replace(/\s+/g, '') : ""; 
        var cantidadS = parseFloat(datosSap[r][6]) || 0; 
        
        if (docCodigo.indexOf("IM") === 0 || docCodigo.indexOf("EM") === 0) {
          var fObjS = normalizarFechaCelda(fRawS);
          if (fObjS && fObjS >= dInicio && fObjS <= dFin) {
            if (cantidadS !== 0 && variedadActual !== "") {
              var dia = ("0" + fObjS.getDate()).slice(-2);
              var mes = ("0" + (fObjS.getMonth() + 1)).slice(-2);
              var anio = fObjS.getFullYear();
              var fTxtS = dia + "/" + mes + "/" + anio;
              
              var llaveS = fTxtS + "|" + variedadActual + "|" + tipoGrupoActual;
              mapaSap[llaveS] = (mapaSap[llaveS] || 0) + cantidadS;
            }
          }
        }
      }

      // GUARDAR EL ÚLTIMO BLOQUE AL LLEGAR AL FINAL DE LA HOJA
      if (llaveStockActual !== "" && ultimoValorColJ !== null) {
        mapaStockFinal[llaveStockActual] = ultimoValorColJ;
      }
    }

    // --- 4. CONSOLIDAR SISPACKING Y SAP ---
    var todasLasLlaves = Object.keys(mapaSispacking).concat(Object.keys(mapaSap).filter(function(item) {
      return Object.keys(mapaSispacking).indexOf(item) < 0;
    }));

    var listaCruce = [];
    todasLasLlaves.forEach(function(llave) {
      var partes = llave.split("|");
      var fecha = partes[0];
      var varNombre = partes[1];
      var tipo = partes[2];
      
      var totalSispacking = mapaSispacking[llave] || 0;
      var totalSap = mapaSap[llave] || 0;
      var diferencia = totalSap - totalSispacking;
      
      if (Math.abs(diferencia) <= 0.01) {
        diferencia = 0;
      }

      var codigoProducto = "-";
      var nombreBusqueda = varNombre.trim().toUpperCase();
      
      for (var b = 0; b < listaProductosBD.length; b++) {
        var descBD = listaProductosBD[b].descripcion;
        if (descBD.indexOf(nombreBusqueda) !== -1 || nombreBusqueda.indexOf(descBD) !== -1) {
          codigoProducto = listaProductosBD[b].codigo;
          break;
        }
      }
      
      listaCruce.push({
        empresa: empresaSeleccionada.toUpperCase(),
        codigo: codigoProducto,
        fecha: fecha,
        variedad: varNombre,
        tipoDescarte: tipo === "CAMPO" ? "GRANEL DESCARTE CAMPO" : "DESCARTE PACKING",
        sispacking: totalSispacking,
        sap: totalSap,
        diferencia: diferencia,
        grupo: tipo
      });
    });

    listaCruce.sort(function(a, b) {
      var pA = a.fecha.split("/"), pB = b.fecha.split("/");
      return new Date(pA[2], pA[1]-1, pA[0]) - new Date(pB[2], pB[1]-1, pB[0]);
    });

    // --- 5. CONSTRUIR LISTA DE STOCK TOTAL ACUMULADO ---
    var listaStockTotal = [];
    Object.keys(mapaStockFinal).forEach(function(key) {
      var p = key.split("|");
      var vNombre = p[0];
      var tGrupo = p[1];
      var codP = "-";
      var nBusq = vNombre.trim().toUpperCase();

      for (var b2 = 0; b2 < listaProductosBD.length; b2++) {
        var descBD2 = listaProductosBD[b2].descripcion;
        if (descBD2.indexOf(nBusq) !== -1 || nBusq.indexOf(descBD2) !== -1) {
          codP = listaProductosBD[b2].codigo;
          break;
        }
      }

      listaStockTotal.push({
        empresa: empresaSeleccionada.toUpperCase(),
        codigo: codP,
        variedad: vNombre,
        tipoDescarte: tGrupo === "CAMPO" ? "GRANEL DESCARTE CAMPO" : "DESCARTE PACKING",
        stockTotal: (mapaStockFinal[key] !== undefined && mapaStockFinal[key] !== null) ? mapaStockFinal[key] : 0
      });
    });

    return { exito: true, datos: listaCruce, stockTotal: listaStockTotal };
  } catch (err) {
    return { exito: false, mensaje: "Error crítico en proceso de cruce SAP: " + err.toString() };
  }
}

// ==========================================
// SERVICIOS PARA MÓDULO DE GUÍAS PROVISIONALES
// ==========================================

var ID_SS_MAESTRO_GUIAS = "110fiUgfJizj_JR7RG3B-xnwfC3wx0NAlbAD4kie-oZ4";

/**
 * Busca un producto en BD-PRODUCTOS mediante su código.
 * Columna A = Código / Columna B = Descripción.
 */
function buscarProductoPorCodigo(codigo) {
  try {
    var ss = SpreadsheetApp.openById(ID_SS_MAESTRO_GUIAS);
    var sheet = ss.getSheetByName("BD-PRODUCTOS");

    if (!sheet) {
      return { exito: false, descripcion: "Error: No existe BD-PRODUCTOS" };
    }

    var datos = sheet.getDataRange().getValues();
    var codigoBuscado = String(codigo || "").trim().toLowerCase();

    for (var i = 1; i < datos.length; i++) {
      var codCelda = String(datos[i][0] || "").trim().toLowerCase();
      if (codCelda === codigoBuscado) {
        return {
          exito: true,
          descripcion: String(datos[i][1] || "").trim().toUpperCase()
        };
      }
    }

    return { exito: false, descripcion: "PRODUCTO NO ENCONTRADO" };
  } catch (e) {
    return { exito: false, descripcion: "Error: " + e.toString() };
  }
}

/**
 * Directorio oficial de empresas emisoras.
 */
function obtenerDatosEmpresaEmisora(codigoEmpresa) {
  var directorio = {
    "AGA": {
      razonSocial: "AGRÍCOLA ANDREA S.A.C.",
      ruc: "20505688903"
    },
    "LARAMA": {
      razonSocial: "LARAMA BERRIES S.A.C.",
      ruc: "20602842739"
    },
    "ARENUVA": {
      razonSocial: "ARENUVA S.A.C.",
      ruc: "20607519391"
    }
  };

  var codigo = String(codigoEmpresa || "").trim().toUpperCase();

  return directorio[codigo] || {
    razonSocial: "",
    ruc: ""
  };
}

/**
 * Guarda la cabecera y detalle de la guía en HISTORIAL_GUIAS.
 * El correlativo se genera por empresa y se protege con LockService.
 */
function guardarGuiaProvisional(datosCabecera, listaProductos) {
  var lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    var ss = SpreadsheetApp.openById(ID_SS_MAESTRO_GUIAS);
    var sheet = ss.getSheetByName("HISTORIAL_GUIAS");

    if (!sheet) {
      return { exito: false, mensaje: "Error: No se encontró la pestaña HISTORIAL_GUIAS" };
    }

    if (!datosCabecera || !listaProductos || !listaProductos.length) {
      return { exito: false, mensaje: "No se recibieron los datos completos de la guía." };
    }

    var empresa = String(datosCabecera.empresa || "").trim().toUpperCase();
    var empresaData = obtenerDatosEmpresaEmisora(empresa);

    if (!empresaData.razonSocial) {
      return { exito: false, mensaje: "La empresa emisora seleccionada no es válida." };
    }

    var siguiente = obtenerSiguienteCorrelativoInterno_(sheet, empresa);
    var nuevoCorrelativo = siguiente.correlativo;

    var filasParaInsertar = [];
    for (var i = 0; i < listaProductos.length; i++) {
      var prod = listaProductos[i] || {};
      filasParaInsertar.push([
        nuevoCorrelativo,
        empresa,
        empresaData.ruc,
        String(datosCabecera.cliente || "").trim(),
        String(datosCabecera.documento || "").trim(),
        String(datosCabecera.direccion || "").trim(),
        String(datosCabecera.fechaEmision || "").trim(),
        String(datosCabecera.fechaVencimiento || "").trim(),
        String(datosCabecera.condicion || "CONTADO").trim(),
        String(datosCabecera.autorizado || "").trim(),
        String(prod.codigo || "").trim(),
        String(prod.descripcion || "").trim(),
        parseFloat(prod.cantidad) || 0
      ]);
    }

    var primeraFila = sheet.getLastRow() + 1;
    sheet.getRange(primeraFila, 1, filasParaInsertar.length, filasParaInsertar[0].length).setValues(filasParaInsertar);

    var urlLogo = obtenerUrlLogo(empresa);

    return {
      exito: true,
      correlativo: nuevoCorrelativo,
      empresa: empresa,
      razonSocial: empresaData.razonSocial,
      rucEmisor: empresaData.ruc,
      urlLogo: urlLogo
    };

  } catch (e) {
    return { exito: false, mensaje: "Error crítico al guardar la guía: " + e.toString() };
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

/**
 * Calcula el siguiente correlativo dentro del mismo lock de guardado.
 * Columna A de HISTORIAL_GUIAS = correlativo.
 */
function obtenerSiguienteCorrelativoInterno_(sheet, codigoEmpresa) {
  var prefijos = {
    "AGA": "T001-",
    "ARENUVA": "T002-",
    "LARAMA": "T003-"
  };

  var empresa = String(codigoEmpresa || "").trim().toUpperCase();
  var prefijo = prefijos[empresa] || "T001-";
  var maxNumero = 0;
  var ultimaFila = sheet.getLastRow();

  if (ultimaFila > 1) {
    var correlativos = sheet.getRange(2, 1, ultimaFila - 1, 1).getDisplayValues();

    for (var i = 0; i < correlativos.length; i++) {
      var corr = String(correlativos[i][0] || "").trim().toUpperCase();
      if (corr.indexOf(prefijo) === 0) {
        var num = parseInt(corr.substring(prefijo.length), 10);
        if (!isNaN(num) && num > maxNumero) {
          maxNumero = num;
        }
      }
    }
  }

  var siguiente = maxNumero + 1;
  return {
    prefijo: prefijo,
    numero: siguiente,
    correlativo: prefijo + String(siguiente).padStart(8, '0')
  };
}

/**
 * Obtiene el siguiente correlativo para mostrarlo en pantalla.
 * No reserva el número; el número definitivo se genera al guardar.
 */
function obtenerSiguienteCorrelativo(codigoEmpresa) {
  try {
    var ss = SpreadsheetApp.openById(ID_SS_MAESTRO_GUIAS);
    var sheet = ss.getSheetByName("HISTORIAL_GUIAS");
    var prefijos = {
      "AGA": "T001-",
      "ARENUVA": "T002-",
      "LARAMA": "T003-"
    };
    var empresa = String(codigoEmpresa || "").trim().toUpperCase();
    var prefijo = prefijos[empresa] || "T001-";

    if (!sheet) return prefijo + "00000001";

    var datos = obtenerSiguienteCorrelativoInterno_(sheet, empresa);
    return datos.correlativo;
  } catch (e) {
    Logger.log("Error al obtener correlativo: " + e.toString());
    return "T001-00000001";
  }
}

/**
 * Obtiene la URL del logo desde CONFIGURACION.
 * CONFIGURACION: A = empresa, C = URL del logo.
 */
function obtenerUrlLogo(codigoEmpresa) {
  var empresa = String(codigoEmpresa || '').trim().toUpperCase();

  var logos = {
    'AGA': LOGO_AGA,
    'LARAMA': LOGO_LARAMA,
    'ARENUVA': LOGO_ARENUVA
  };

  return logos[empresa] || '';
}

/**
 * Devuelve el HTML de impresión de la guía.
 */
function generarHtmlImpresionGuia() {
  return HtmlService.createHtmlOutputFromFile('FormatoImpresion').getContent();
}

/**
 * Lee el link de origen utilizado en Guías Provisionales.
 * Estructura: B=EMPRESA, C=SEDE, D=CÓDIGO, E=DESCRIPCIÓN, F=auxiliar, G=CANT/KG, H=UM.
 */
function obtenerProductosDesdeLinkGuia(empresaSeleccionada, urlOrigen) {
  try {
    var empresa = String(empresaSeleccionada || '').trim().toUpperCase();
    var url = String(urlOrigen || '').trim();

    if (!empresa) return { exito: false, mensaje: 'Debe seleccionar una empresa.' };
    if (!url) return { exito: false, mensaje: 'Debe ingresar un link de origen.' };

    var match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (!match || !match[1]) {
      return { exito: false, mensaje: 'El link no corresponde a un Google Sheets válido.' };
    }

    var ssOrigen = SpreadsheetApp.openById(match[1]);
    var gidMatch = url.match(/[?#&]gid=(\d+)/);
    var sheet = gidMatch ? ssOrigen.getSheetById(Number(gidMatch[1])) : null;
    if (!sheet) sheet = ssOrigen.getSheetByName('DESCARTE');
    if (!sheet) sheet = ssOrigen.getSheets()[0];
    if (!sheet) return { exito: false, mensaje: 'No se pudo identificar la pestaña del archivo.' };

    var lastRow = sheet.getLastRow();
    if (lastRow < 1) return { exito: false, mensaje: 'La pestaña indicada no contiene datos.' };

    var datos = sheet.getRange(1, 2, lastRow, 7).getDisplayValues();
    var normalizar = function(v) {
      return String(v == null ? '' : v).trim().toUpperCase()
        .normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    };
    var aliases = {
      AGA: ['AGA', 'AGRICOLA ANDREA', 'AGRICOLA ANDREA SAC'],
      LARAMA: ['LARAMA', 'AGROPECUARIA LARAMA', 'AGROPECUARIA LARAMA SAC'],
      ARENUVA: ['ARENUVA', 'AGRICOLA ARENUVA', 'AGRICOLA ARENUVA SAC']
    }[empresa] || [empresa];

    var coincideEmpresa = function(valor) {
      var valorNormalizado = normalizar(valor);
      return aliases.some(function(alias) {
        var aliasNormalizado = normalizar(alias);
        return valorNormalizado === aliasNormalizado || valorNormalizado.indexOf(aliasNormalizado) >= 0;
      });
    };

    var toNumber = function(valor) {
      var s = String(valor == null ? '' : valor).trim().replace(/\s/g, '');
      if (!s) return 0;
      if (s.indexOf(',') >= 0 && s.indexOf('.') >= 0) {
        s = s.replace(/,/g, '');
      } else if (s.indexOf(',') >= 0) {
        s = s.replace(/,/g, '.');
      }
      var n = parseFloat(s);
      return isNaN(n) ? 0 : n;
    };

    var productos = [];

    for (var i = 0; i < datos.length; i++) {
      var fila = datos[i];
      var empresaFila = fila[0];
      var codigo = String(fila[2] || '').trim();
      var descripcion = String(fila[3] || '').trim();
      var cantidad = toNumber(fila[5]);
      var um = String(fila[6] || '').trim();

      if (!coincideEmpresa(empresaFila) || !codigo || !descripcion) continue;
      if (!/^\d+$/.test(codigo.replace(/\s/g, ''))) continue;
      if (cantidad <= 0) continue;

      productos.push({
        codigo: codigo,
        descripcion: descripcion.toUpperCase(),
        cantidad: cantidad,
        um: um || 'KG'
      });
    }

    if (!productos.length) {
      return {
        exito: false,
        mensaje: 'No se encontraron productos válidos para ' + empresa + ' en la pestaña ' + sheet.getName() + '.'
      };
    }

    return {
      exito: true,
      empresa: empresa,
      productos: productos,
      cantidad: productos.length,
      pestaña: sheet.getName(),
      gid: String(sheet.getSheetId())
    };

  } catch (error) {
    return { exito: false, mensaje: 'No se pudo leer el link de origen: ' + error.toString() };
  }
}

/**
 * Busca un cliente por RUC/DNI en el historial de Ventas.
 */
function buscarClientePorRuc(ruc) {
  try {
    var ss = SpreadsheetApp.openById(ID_SS_MAESTRO_GUIAS);
    var sheet = ss.getSheetByName('VENTAS');
    if (!sheet) return { exito: false, cliente: '' };

    var buscado = String(ruc || '').trim().toUpperCase().replace(/[\s\-.]/g, '');
    if (!buscado) return { exito: false, cliente: '' };

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { exito: false, cliente: '' };

    var datos = sheet.getRange(2, 3, lastRow - 1, 2).getDisplayValues();
    for (var i = datos.length - 1; i >= 0; i--) {
      var rucFila = String(datos[i][0] || '').trim().toUpperCase().replace(/[\s\-.]/g, '');
      var cliente = String(datos[i][1] || '').trim();
      if (rucFila === buscado && cliente) {
        return { exito: true, cliente: cliente };
      }
    }

    return { exito: false, cliente: '' };
  } catch (error) {
    return { exito: false, cliente: '', mensaje: error.toString() };
  }
}

/**
 * Convierte enlaces compartidos de Google Drive en URLs directas de imagen.
 */
function transformarUrlDriveDirecta(url) {
  if (!url) return "";
  var match = String(url).match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (match && match[1]) {
    return "https://lh3.googleusercontent.com/d/" + match[1];
  }
  return url;
}

// ============================================================
// SERVICIOS PARA MÓDULO DE VENTAS
// Se agregan únicamente las funciones que necesita Ventas.
// No modifica las funciones existentes de Guías, Reportes o Cruce.
// ============================================================

function extraerIdSpreadsheetDesdeUrl_(url) {
  var texto = String(url || '').trim();
  var match = texto.match(/\/d\/([a-zA-Z0-9-_]+)/);
  return match && match[1] ? match[1] : '';
}

/**
 * Lee el link de origen del módulo Ventas.
 * Estructura fija del reporte:
 * B = EMPRESA
 * C = SEDE
 * D = COD
 * E = DESCRIPCION
 * F = auxiliar
 * G = CANT / KG
 * H = UM
 *
 * La empresa seleccionada filtra las filas y se devuelve también
 * el valor de G para que Ventas lo coloque automáticamente.
 */
/**
 * Compatibilidad con versiones anteriores de Guías que solicitaban
 * directamente el HTML del formato de impresión.
 */
function obtenerFormatoImpresion() {
  return HtmlService.createHtmlOutputFromFile('FormatoImpresion').getContent();
}

function obtenerProductosDesdeLinkVenta(empresaSeleccionada, urlOrigen) {
  try {
    var empresa = String(empresaSeleccionada || '').trim().toUpperCase();
    var url = String(urlOrigen || '').trim();

    if (!empresa) {
      return { exito: false, mensaje: 'Debe seleccionar una empresa.' };
    }

    if (!url) {
      return { exito: false, mensaje: 'Debe ingresar un link de origen.' };
    }

    var spreadsheetId = extraerIdSpreadsheetDesdeUrl_(url);
    if (!spreadsheetId) {
      return { exito: false, mensaje: 'El link no corresponde a un Google Sheets válido.' };
    }

    var ssOrigen = SpreadsheetApp.openById(spreadsheetId);
    var gidMatch = url.match(/[?#&]gid=(\d+)/);
    var sheet = gidMatch ? ssOrigen.getSheetById(Number(gidMatch[1])) : null;

    if (!sheet) sheet = ssOrigen.getSheetByName('DESCARTE');
    if (!sheet) sheet = ssOrigen.getSheets()[0];
    if (!sheet) {
      return { exito: false, mensaje: 'No se pudo identificar la pestaña del archivo.' };
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < 1) {
      return { exito: false, mensaje: 'La pestaña indicada no contiene datos.' };
    }

    // B:H = 7 columnas. Así respetamos exactamente la estructura
    // que ya utiliza el reporte de origen.
    var datos = sheet.getRange(1, 2, lastRow, 7).getDisplayValues();

    var normalizar = function(valor) {
      return String(valor == null ? '' : valor).trim().toUpperCase()
        .normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    };

    var aliases = {
      AGA: ['AGA', 'AGRICOLA ANDREA', 'AGRICOLA ANDREA SAC'],
      LARAMA: ['LARAMA', 'AGROPECUARIA LARAMA', 'AGROPECUARIA LARAMA SAC'],
      ARENUVA: ['ARENUVA', 'AGRICOLA ARENUVA', 'AGRICOLA ARENUVA SAC']
    }[empresa] || [empresa];

    var coincideEmpresa = function(valor) {
      var valorNormalizado = normalizar(valor);
      return aliases.some(function(alias) {
        var aliasNormalizado = normalizar(alias);
        return valorNormalizado === aliasNormalizado ||
               valorNormalizado.indexOf(aliasNormalizado) >= 0;
      });
    };

    var toNumber = function(valor) {
      var s = String(valor == null ? '' : valor).trim().replace(/\s/g, '');
      if (!s) return 0;

      // El reporte mostrado utiliza 4,049.47, por lo que una coma
      // junto a punto representa separador de miles.
      if (s.indexOf(',') >= 0 && s.indexOf('.') >= 0) {
        s = s.replace(/,/g, '');
      } else if (s.indexOf(',') >= 0) {
        s = s.replace(/,/g, '.');
      }

      var n = parseFloat(s);
      return isNaN(n) ? 0 : n;
    };

   var productos = [];

    for (var i = 0; i < datos.length; i++) {
      var fila = datos[i];

      var empresaFila = fila[0]; // B
      var codigo = String(fila[2] || '').trim(); // D
      var descripcion = String(fila[3] || '').trim(); // E
      var cantidad = toNumber(fila[5]); // G
      var um = String(fila[6] || '').trim(); // H

      if (!coincideEmpresa(empresaFila)) continue;
      if (!codigo || !descripcion) continue;
      if (!/^\d+$/.test(codigo.replace(/\s/g, ''))) continue;
      if (cantidad <= 0) continue;

      productos.push({
        codigo: codigo,
        descripcion: descripcion.toUpperCase(),
        cantidad: cantidad,
        um: um || 'KG'
      });
    }

    if (!productos.length) {
      return {
        exito: false,
        mensaje: 'No se encontraron códigos para la empresa ' + empresa +
                 ' en la pestaña indicada por el link.'
      };
    }

    return {
      exito: true,
      empresa: empresa,
      productos: productos,
      cantidad: productos.length,
      pestaña: sheet.getName(),
      gid: String(sheet.getSheetId())
    };

  } catch (error) {
    return {
      exito: false,
      mensaje: 'No se pudo leer el link de origen: ' + error.toString()
    };
  }
}

/**
 * Guarda la venta en la pestaña VENTAS.
 * A = CORRELATIVO
 * B = FECHA DE DESPACHO
 * C = RUC
 * D = CLIENTE
 * E = CODIGO
 * F = VARIEDAD
 * G = KG TOTAL
 * H = EMPRESA
 * I = LINK ORIGEN
 */
function guardarVenta(datosVenta) {
  try {
    var ss = SpreadsheetApp.openById(ID_SS_MAESTRO_GUIAS);
    var sheet = ss.getSheetByName('VENTAS');

    if (!sheet) {
      return { exito: false, mensaje: 'No se encontró la pestaña VENTAS.' };
    }

    if (!datosVenta) {
      return { exito: false, mensaje: 'No se recibieron los datos de la venta.' };
    }

    var correlativo = String(datosVenta.correlativo || '').trim();
    var fechaDespacho = String(datosVenta.fechaDespacho || '').trim();
    var empresa = String(datosVenta.empresa || '').trim().toUpperCase();
    var linkOrigen = String(datosVenta.linkOrigen || '').trim();
    var ruc = String(datosVenta.ruc || '').trim();
    var cliente = String(datosVenta.cliente || '').trim();
    var detalle = datosVenta.detalle || [];

    if (!correlativo) return { exito: false, mensaje: 'El correlativo es obligatorio.' };
    if (!fechaDespacho) return { exito: false, mensaje: 'La fecha de despacho es obligatoria.' };
    if (!empresa) return { exito: false, mensaje: 'La empresa es obligatoria.' };
    if (!linkOrigen) return { exito: false, mensaje: 'El link de origen es obligatorio.' };
    if (!ruc) return { exito: false, mensaje: 'El RUC es obligatorio.' };
    if (!cliente) return { exito: false, mensaje: 'El cliente es obligatorio.' };
    if (!detalle.length) return { exito: false, mensaje: 'La venta debe tener al menos un producto.' };

    var ultimaFila = sheet.getLastRow();
    if (ultimaFila > 1) {
      var correlativosExistentes = sheet.getRange(2, 1, ultimaFila - 1, 1).getValues();
      var buscado = correlativo.toUpperCase();

      for (var i = 0; i < correlativosExistentes.length; i++) {
        var existente = String(correlativosExistentes[i][0] || '').trim().toUpperCase();
        if (existente === buscado) {
          return {
            exito: false,
            mensaje: 'El correlativo ' + correlativo + ' ya existe en VENTAS.'
          };
        }
      }
    }

    var partesFecha = fechaDespacho.split('-');
    var fechaParaGuardar;
    if (partesFecha.length === 3) {
      fechaParaGuardar = new Date(
        parseInt(partesFecha[0], 10),
        parseInt(partesFecha[1], 10) - 1,
        parseInt(partesFecha[2], 10)
      );
    } else {
      fechaParaGuardar = fechaDespacho;
    }

    var filasParaGuardar = [];

    for (var j = 0; j < detalle.length; j++) {
      var producto = detalle[j] || {};
      var codigo = String(producto.codigo || '').trim();
      var variedad = String(producto.variedad || '').trim();
      var kg = parseFloat(producto.kgTotal) || 0;

      if (!codigo) return { exito: false, mensaje: 'Existe una línea sin código de producto.' };
      if (!variedad) return { exito: false, mensaje: 'Existe una línea sin variedad.' };
      if (kg <= 0) return { exito: false, mensaje: 'Todos los productos deben tener KG mayor a cero.' };

      filasParaGuardar.push([
        correlativo,
        fechaParaGuardar,
        ruc,
        cliente,
        codigo,
        variedad,
        kg,
        empresa,
        linkOrigen
      ]);
    }

    if (!sheet.getRange(1, 8).getValue()) sheet.getRange(1, 8).setValue('EMPRESA');
    if (!sheet.getRange(1, 9).getValue()) sheet.getRange(1, 9).setValue('LINK ORIGEN');

    var filaInicial = sheet.getLastRow() + 1;
    sheet.getRange(filaInicial, 1, filasParaGuardar.length, 9).setValues(filasParaGuardar);

    return {
      exito: true,
      mensaje: 'Venta guardada correctamente.',
      correlativo: correlativo
    };

  } catch (error) {
    return {
      exito: false,
      mensaje: 'Error al guardar la venta: ' + error.toString()
    };
  }
}

/**
 * Busca una venta completa mediante su correlativo.
 */
function buscarVentaPorCorrelativo(correlativo) {
  try {
    var ss = SpreadsheetApp.openById(ID_SS_MAESTRO_GUIAS);
    var sheet = ss.getSheetByName('VENTAS');

    if (!sheet) return { exito: false, mensaje: 'No se encontró la pestaña VENTAS.' };

    var buscado = String(correlativo || '').trim().toUpperCase();
    if (!buscado) return { exito: false, mensaje: 'Ingrese un correlativo para buscar.' };

    var ultimaFila = sheet.getLastRow();
    if (ultimaFila <= 1) return { exito: false, mensaje: 'La pestaña VENTAS no contiene registros.' };

    var datos = sheet.getRange(2, 1, ultimaFila - 1, 9).getValues();
    var resultados = [];

    for (var i = 0; i < datos.length; i++) {
      var correlativoFila = String(datos[i][0] || '').trim().toUpperCase();
      if (correlativoFila !== buscado) continue;

      var fecha = datos[i][1];
      var fechaFormateada = '';
      if (fecha instanceof Date) {
        fechaFormateada = Utilities.formatDate(
          fecha,
          Session.getScriptTimeZone(),
          'dd/MM/yyyy'
        );
      } else {
        fechaFormateada = String(fecha || '').trim();
      }

      resultados.push({
        correlativo: String(datos[i][0] || '').trim(),
        fechaDespacho: fechaFormateada,
        ruc: String(datos[i][2] || '').trim(),
        cliente: String(datos[i][3] || '').trim(),
        codigo: String(datos[i][4] || '').trim(),
        variedad: String(datos[i][5] || '').trim(),
        kgTotal: parseFloat(datos[i][6]) || 0,
        empresa: String(datos[i][7] || '').trim(),
        linkOrigen: String(datos[i][8] || '').trim()
      });
    }

    if (!resultados.length) {
      return {
        exito: false,
        mensaje: 'No se encontró ninguna venta con el correlativo: ' + correlativo
      };
    }

    return { exito: true, datos: resultados };

  } catch (error) {
    return {
      exito: false,
      mensaje: 'Error al buscar la venta: ' + error.toString()
    };
  }
}
/**
 * ============================================================
 * BUSCAR GUÍA PROVISIONAL POR CORRELATIVO
 * ============================================================
 *
 * Busca todas las filas de HISTORIAL_GUIAS que pertenezcan
 * al mismo correlativo y reconstruye la guía completa.
 *
 * NO modifica ni genera ningún correlativo.
 * SOLO LEE información existente.
 */
function buscarGuiaPorCorrelativo(correlativoBuscado) {

  try {

    var correlativo = String(correlativoBuscado || '')
      .trim()
      .toUpperCase();

    if (!correlativo) {
      return {
        exito: false,
        mensaje: 'Ingrese un correlativo para buscar.'
      };
    }

    var ss = SpreadsheetApp.openById(ID_SS_MAESTRO_GUIAS);

    var sheet = ss.getSheetByName('HISTORIAL_GUIAS');

    if (!sheet) {
      return {
        exito: false,
        mensaje: 'No se encontró la pestaña HISTORIAL_GUIAS.'
      };
    }

    var ultimaFila = sheet.getLastRow();

    if (ultimaFila < 2) {
      return {
        exito: false,
        mensaje: 'No existen guías registradas.'
      };
    }

    /*
     * Estructura actual:
     *
     * A = CORRELATIVO
     * B = EMPRESA
     * C = RUC EMISOR
     * D = CLIENTE
     * E = DOCUMENTO
     * F = DIRECCIÓN
     * G = FECHA EMISIÓN
     * H = FECHA VENCIMIENTO
     * I = CONDICIÓN
     * J = AUTORIZADO
     * K = CÓDIGO PRODUCTO
     * L = DESCRIPCIÓN
     * M = CANTIDAD
     */

    var datos = sheet
      .getRange(2, 1, ultimaFila - 1, 13)
      .getDisplayValues();

    var filasGuia = [];

    for (var i = 0; i < datos.length; i++) {

      var fila = datos[i];

      var corrFila = String(fila[0] || '')
        .trim()
        .toUpperCase();

      if (corrFila === correlativo) {
        filasGuia.push(fila);
      }
    }

    if (filasGuia.length === 0) {

      return {
        exito: false,
        mensaje: 'No se encontró ninguna guía con el correlativo ' + correlativo + '.'
      };

    }

    var primera = filasGuia[0];

    var empresa = String(primera[1] || '')
      .trim()
      .toUpperCase();

    var empresaData = obtenerDatosEmpresaEmisora(empresa);

    var productos = [];

    for (var j = 0; j < filasGuia.length; j++) {

      var f = filasGuia[j];

      var codigo = String(f[10] || '').trim();

      if (!codigo) {
        continue;
      }

      productos.push({
        codigo: codigo,
        descripcion: String(f[11] || '').trim(),
        cantidad: parseFloat(
          String(f[12] || '0').replace(',', '.')
        ) || 0,
        um: 'KG'
      });
    }

    return {

      exito: true,

      correlativo: correlativo,

      cabecera: {

        empresa: empresaData.razonSocial || empresa,

        rucEmisor:
          empresaData.ruc ||
          String(primera[2] || '').trim(),

        cliente:
          String(primera[3] || '').trim(),

        documento:
          String(primera[4] || '').trim(),

        direccion:
          String(primera[5] || '').trim(),

        fechaEmision:
          String(primera[6] || '').trim(),

        fechaVencimiento:
          String(primera[7] || '').trim(),

        condicion:
          String(primera[8] || '').trim(),

        autorizado:
          String(primera[9] || '').trim(),

        /*
         * La versión actual del historial no guarda
         * observaciones ni creador.
         */
        observaciones: '',

        urlLogo:
          obtenerUrlLogo(empresa)

      },

      creador: 'SISTEMA',

      productos: productos

    };

  } catch (e) {

    Logger.log(
      'Error buscando guía por correlativo: ' +
      e.toString()
    );

    return {
      exito: false,
      mensaje:
        'Error al buscar la guía: ' +
        e.toString()
    };

  }

}
