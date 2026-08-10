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

/**
 * Busca un producto en la pestaña 'BD-PRODUCTOS' mediante su código.
 * Asume Columna A = Código, Columna B = Descripción/Variedad.
 */
function buscarProductoPorCodigo(codigo) {
  try {
    var sheetId = "110fiUgfJizj_JR7RG3B-xnwfC3wx0NAlbAD4kie-oZ4";
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName("BD-PRODUCTOS");
    
    if (!sheet) {
      return { exito: false, descripcion: "Error: No existe BD-PRODUCTOS" };
    }
    
    var datos = sheet.getDataRange().getValues();
    var codigoBuscado = codigo.toString().trim().toLowerCase();
    
    for (var i = 1; i < datos.length; i++) {
      var codCelda = datos[i][0].toString().trim().toLowerCase();
      if (codCelda === codigoBuscado) {
        return { exito: true, descripcion: datos[i][1].toString().trim().toUpperCase() };
      }
    }
    
    return { exito: false, descripcion: "PRODUCTO NO ENCONTRADO" };
  } catch (e) {
    return { exito: false, descripcion: "Error: " + e.toString() };
  }
}

/**
 * Mapeo oficial centralizado de Empresas emisoras, Razones Sociales y RUCs
 */
function obtenerDatosEmpresaEmisora(codigoEmpresa) {
  var directorio = {
    "AGA": { razonSocial: "AGRÍCOLA ANDREA S.A.C.", ruc: "20522108118" },
    "LARAMA": { razonSocial: "AGROPECUARIA LARAMA S.A.C.", ruc: "20601344444" }, 
    "ARENUVA": { razonSocial: "AGRÍCOLA ARENUVA S.A.C.", ruc: "20602511111" }   
  };
  
  return directorio[codigoEmpresa.toUpperCase()] || { razonSocial: "", ruc: "" };
}

/**
 * Guarda la cabecera y el detalle de la guía en 'HISTORIAL_GUIAS'
 * Genera de forma segura el siguiente correlativo incremental.
 */
function guardarGuiaProvisional(datosCabecera, listaProductos) {
  try {
    var sheetId = "110fiUgfJizj_JR7RG3B-xnwfC3wx0NAlbAD4kie-oZ4";
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheetByName("HISTORIAL_GUIAS");
    
    if (!sheet) {
      return { exito: false, mensaje: "Error: No se encontró la pestaña HISTORIAL_GUIAS" };
    }
    
    // Generar Correlativo Automático
    var ultimaFila = sheet.getLastRow();
    var sgteNumero = 1;

    if (ultimaFila > 1) {
      var ultimoCorrelativo = sheet.getRange(ultimaFila, 1).getValue().toString();
      var numeroExtraido = parseInt(ultimoCorrelativo.replace("DESC-CAL-", ""), 10);
      if (!isNaN(numeroExtraido)) {
        sgteNumero = numeroExtraido + 1;
      }
    }

    var nuevoCorrelativo = "DESC-CAL-" + String(sgteNumero).padStart(3, '0');
    
    // Preparar el lote de filas para setValues()
    var filasParaInsertar = [];
    
    for (var i = 0; i < listaProductos.length; i++) {
      var prod = listaProductos[i];
      
      filasParaInsertar.push([
        nuevoCorrelativo,
        datosCabecera.empresa,
        datosCabecera.rucEmisor,
        datosCabecera.cliente,
        datosCabecera.documento,
        datosCabecera.direccion,
        datosCabecera.fechaEmision,
        datosCabecera.fechaVencimiento,
        datosCabecera.condicion,
        datosCabecera.autorizado,
        prod.codigo,
        prod.descripcion,
        parseFloat(prod.cantidad) || 0
      ]);
    }
    
    if (filasParaInsertar.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, filasParaInsertar.length, filasParaInsertar[0].length).setValues(filasParaInsertar);
    }
    
    return { exito: true, correlativo: nuevoCorrelativo };
    
  } catch (e) {
    return { exito: false, mensaje: "Error crítico: " + e.toString() };
  }
}

function obtenerFormatoImpresion() {
  return HtmlService.createHtmlOutputFromFile('FormatoImpresion').getContent();
}

/**
 * Obtiene el siguiente correlativo para la empresa elegida.
 * Revisa el correlativo máximo guardado en "HISTORIAL_GUIAS".
 */
function obtenerSiguienteCorrelativo(codigoEmpresa) {
  try {
    var ss = SpreadsheetApp.openById(ID_SS_DESTINO);
    var sheet = ss.getSheetByName("HISTORIAL_GUIAS");
    
    var prefijo = "T001-";
    if (codigoEmpresa === "AGA") prefijo = "T001-";
    else if (codigoEmpresa === "ARENUVA") prefijo = "T002-";
    else if (codigoEmpresa === "LARAMA") prefijo = "T003-";
    
    if (!sheet) return prefijo + "00000001";
    
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return prefijo + "00000001";
    
    var maxNumero = 0;
    for (var i = 1; i < data.length; i++) {
      var corr = String(data[i][1] || ""); // Columna B: Correlativo
      if (corr.startsWith(prefijo)) {
        var numStr = corr.replace(prefijo, "");
        var num = parseInt(numStr, 10);
        if (!isNaN(num) && num > maxNumero) {
          maxNumero = num;
        }
      }
    }
    
    var siguiente = maxNumero + 1;
    var numFormateado = ("00000000" + siguiente).slice(-8);
    return prefijo + numFormateado;
    
  } catch (e) {
    Logger.log("Error al obtener correlativo: " + e.toString());
    return "T001-00000001";
  }
} // <-- Llave de cierre corregida

/**
 * Obtiene la URL del logo de la empresa desde la pestaña CONFIGURACION.
 */
function obtenerUrlLogo(codigoEmpresa) {
  try {
    var ss = SpreadsheetApp.openById(ID_SS_DESTINO);
    var sheet = ss.getSheetByName("CONFIGURACION");
    if (!sheet) return "";
    
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toUpperCase() === String(codigoEmpresa).trim().toUpperCase()) {
        return transformarUrlDriveDirecta(String(data[i][2])); 
      }
    }
    return "";
  } catch (e) {
    return "";
  }
}

/**
 * Convierte enlaces compartidos de Google Drive en URLs directas de imagen.
 */
function transformarUrlDriveDirecta(url) {
  if (!url) return "";
  var match = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (match && match[1]) {
    return "https://lh3.googleusercontent.com/d/" + match[1];
  }
  return url;
}