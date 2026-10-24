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
function _crearSesionDespachos_(usuario, tipoAcceso, cliente, empresa) {
  var token = Utilities.getUuid();
  CacheService.getScriptCache().put(
    "DESPACHOS_SESION_" + token,
    JSON.stringify({
      usuario: String(usuario || ""),
      tipoAcceso: String(tipoAcceso || ""),
      cliente: String(cliente || ""),
      empresa: String(empresa || "")
    }),
    21600
  );
  return token;
}

function _obtenerSesionDespachos_(token) {
  if (!token) return null;
  var raw = CacheService.getScriptCache().get("DESPACHOS_SESION_" + token);
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
}

function _exigirAdminDespachos_(token) {
  var sesion = _obtenerSesionDespachos_(token);
  if (!sesion || String(sesion.tipoAcceso || "").toUpperCase() === "CLIENTE" || String(sesion.empresa || "").toUpperCase() !== "ADMIN") {
    throw new Error("No tiene permisos para administrar despachos.");
  }
  return sesion;
}

function _exigirClienteDespachos_(token) {
  var sesion = _obtenerSesionDespachos_(token);
  if (!sesion || String(sesion.tipoAcceso || "").toUpperCase() !== "CLIENTE" || !sesion.cliente) {
    throw new Error("Sesión de cliente no válida o expirada.");
  }
  return sesion;
}

// Función para validar el Login desde las pestañas del Sheet
function validarCredenciales(usuario, contrasenia, empresa) {
  try {
    var sheetId = "110fiUgfJizj_JR7RG3B-xnwfC3wx0NAlbAD4kie-oZ4";
    var ss = SpreadsheetApp.openById(sheetId);
    var usuarioLimpio = String(usuario || "").trim();
    var contraseniaLimpia = String(contrasenia || "").trim();
    var accesoSeleccionado = String(empresa || "").trim().toUpperCase();

    // Acceso exclusivo de clientes externos.
    if (accesoSeleccionado === "CLIENTES") {
      var sheetClientesUsuarios = ss.getSheetByName("USUARIOS-CLIENTES");
      if (!sheetClientesUsuarios) {
        return { exito: false, mensaje: "No se encontró la pestaña 'USUARIOS-CLIENTES'." };
      }

      var datosClientesUsuarios = sheetClientesUsuarios.getDataRange().getValues();
      for (var c = 1; c < datosClientesUsuarios.length; c++) {
        var userCliente = String(datosClientesUsuarios[c][0] || "").trim().toLowerCase();
        var passCliente = String(datosClientesUsuarios[c][1] || "").trim();
        var nombreCliente = String(datosClientesUsuarios[c][2] || "").trim();
        var estadoCliente = String(datosClientesUsuarios[c][3] || "ACTIVO").trim().toUpperCase();

        if (userCliente === usuarioLimpio.toLowerCase() && passCliente === contraseniaLimpia) {
          if (estadoCliente !== "ACTIVO") {
            return { exito: false, mensaje: "El usuario cliente se encuentra INACTIVO." };
          }
          if (!nombreCliente) {
            return { exito: false, mensaje: "El usuario cliente no tiene un cliente asignado." };
          }

          var clienteActivo = _clienteActivoDespacho_(nombreCliente);
          if (!clienteActivo) {
            return { exito: false, mensaje: "El cliente asignado al usuario no está ACTIVO en la pestaña CLIENTES." };
          }

          var tokenCliente = _crearSesionDespachos_(usuarioLimpio, "CLIENTE", clienteActivo, "CLIENTES");
          return {
            exito: true,
            usuario: usuarioLimpio,
            empresa: "CLIENTES",
            cliente: clienteActivo,
            tipoAcceso: "CLIENTE",
            token: tokenCliente
          };
        }
      }

      return { exito: false, mensaje: "Usuario o contraseña incorrectos." };
    }

    // Acceso interno existente: se conserva la validación por ACCESO-{empresa}.
    var nombrePestana = "ACCESO-" + accesoSeleccionado;
    var sheet = ss.getSheetByName(nombrePestana);

    if (!sheet) {
      return { exito: false, mensaje: "Error: No se encontró la pestaña de acceso para " + accesoSeleccionado };
    }

    var datos = sheet.getDataRange().getValues();

    for (var i = 1; i < datos.length; i++) {
      var userSheet = String(datos[i][0] || "").trim().toLowerCase();
      var passSheet = String(datos[i][1] || "").trim();
      var estado = datos[i][2] ? String(datos[i][2]).trim().toUpperCase() : "ACTIVO";

      if (userSheet === usuarioLimpio.toLowerCase() && passSheet === contraseniaLimpia) {
        if (estado === "ACTIVO") {
          var tipoInterno = accesoSeleccionado === "ADMIN" ? "ADMIN" : "INTERNO";
          var tokenInterno = _crearSesionDespachos_(usuarioLimpio, tipoInterno, "", accesoSeleccionado);
          return {
            exito: true,
            usuario: usuarioLimpio,
            empresa: accesoSeleccionado,
            tipoAcceso: tipoInterno,
            cliente: "",
            token: tokenInterno
          };
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

// Función que cargará la interfaz principal una vez logueado.
// Los clientes reciben una interfaz aislada que contiene únicamente Despachos.
function cargarInterfazPrincipal(tipoAcceso, cliente, usuario) {
  if (String(tipoAcceso || "").toUpperCase() === "CLIENTE") {
    return HtmlService.createTemplateFromFile('InterfazCliente').evaluate().getContent();
  }
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
    
    // Cargar códigos desde BD-PRODUCTOS.
    // La relación variedad/código se construye dinámicamente para que
    // nuevas variedades agregadas a BD-PRODUCTOS no requieran cambios en código.
    var listaProductosBD = [];
    var mapaCodigosProducto = {};
    var sheetProductos = ss.getSheetByName("BD-PRODUCTOS");

    var normalizarProducto = function(texto) {
      return String(texto || "")
        .toUpperCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .replace(/OZBLU®/g, "")
        .replace(/OZBLU/g, "")
        .replace(/ARANDANO ORGANICO FRESCO/g, "ORGANICO")
        .replace(/ARANDANO FRESCO/g, "")
        .replace(/GRANEL/g, "")
        .replace(/DESCARTE CAMPO/g, "")
        .replace(/DESCARTE PACKING/g, "")
        .replace(/DESCARTE/g, "")
        .replace(/CAMPO/g, "")
        .replace(/PACKING/g, "")
        .replace(/[^A-Z0-9]+/g, " ")
        .replace(/\s+/g, " ")
        .trim();
    };

    var obtenerTipoDescarte = function(texto) {
      var valor = String(texto || "").toUpperCase();
      return valor.includes("PACKING") ? "DESCARTE PACKING" : "DESCARTE CAMPO";
    };

    if (sheetProductos) {
      var datosProductos = sheetProductos.getDataRange().getValues();

      for (var p = 1; p < datosProductos.length; p++) {
        var codigo = datosProductos[p][0] != null
          ? String(datosProductos[p][0]).trim()
          : "";

        var descripcion = datosProductos[p][1] != null
          ? String(datosProductos[p][1]).trim().toUpperCase()
          : "";

        if (!codigo || !descripcion) continue;

        var producto = {
          codigo: codigo,
          descripcion: descripcion,
          tipo: obtenerTipoDescarte(descripcion),
          variedadClave: normalizarProducto(descripcion)
        };

        listaProductosBD.push(producto);

        // Guardar por tipo + variedad normalizada.
        // Se usa la descripción completa de BD-PRODUCTOS como fuente única.
        var claveProducto = producto.tipo + "|" + producto.variedadClave;
        if (!mapaCodigosProducto[claveProducto]) {
          mapaCodigosProducto[claveProducto] = codigo;
        }
      }
    }

    var buscarCodigoProducto = function(variedadProducto, presentacionProducto) {
      var variedadNormalizada = normalizarProducto(variedadProducto);
      if (!variedadNormalizada) return "-";

      var tipoDescarte = obtenerTipoDescarte(presentacionProducto);
      var claveDirecta = tipoDescarte + "|" + variedadNormalizada;

      // Coincidencia directa: permite agregar nuevas variedades sin cambiar el código.
      if (mapaCodigosProducto[claveDirecta]) {
        return mapaCodigosProducto[claveDirecta];
      }

      // Respaldo flexible para pequeñas diferencias de texto entre SISPACKING y BD.
      for (var p = 0; p < listaProductosBD.length; p++) {
        var productoBD = listaProductosBD[p];
        if (productoBD.tipo !== tipoDescarte) continue;

        if (productoBD.variedadClave === variedadNormalizada) {
          return productoBD.codigo;
        }
      }

      return "-";
    };

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
                  codigo: buscarCodigoProducto(
                    subDatos[j][2] ? subDatos[j][2].toString().trim() : "",
                    subDatos[j][3] ? subDatos[j][3].toString().trim() : ""
                  ),
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
                  codigo: buscarCodigoProducto(
                    subDatos[j][9] ? subDatos[j][9].toString().trim() : "",
                    subDatos[j][5] ? subDatos[j][5].toString().trim() : ""
                  ),
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
      var datosProd = sheetProductos.getDataRange().getDisplayValues();
      for (var p = 1; p < datosProd.length; p++) {
        var codCelda = datosProd[p][0] ? datosProd[p][0].toString().trim().toUpperCase() : "";
        var descCelda = datosProd[p][1] ? datosProd[p][1].toString().trim().toUpperCase() : "";
        if (descCelda !== "" && codCelda !== "") {
          listaProductosBD.push({ codigo: codCelda, descripcion: descCelda });
        }
      }
    }

    // Normaliza únicamente el texto para comparar productos sin perder
    // diferencias como ROSITA y ORGÁNICO ROSITA.
    var normalizarProducto = function(texto) {
      return String(texto || "")
        .toUpperCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .replace(/ARANDANO ORGANICO FRESCO/g, "ORGANICO")
        .replace(/ARANDANO FRESCO/g, "")
        .replace(/OZBLU®?/g, "")
        .replace(/GRANEL/g, "")
        .replace(/DESCARTE\s+(CAMPO|PACKING)/g, "")
        .replace(/DESCARTE/g, "")
        .replace(/CAMPO/g, "")
        .replace(/PACKING/g, "")
        .replace(/[-]/g, " ")
        .replace(/\s+/g, " ")
        .trim();
    };

    var obtenerVariedadDesdeBD = function(texto) {
      var clave = normalizarProducto(texto);
      if (!clave) return "";

      for (var b = 0; b < listaProductosBD.length; b++) {
        if (normalizarProducto(listaProductosBD[b].descripcion) === clave) {
          return clave;
        }
      }
      return "";
    };

    var buscarCodigoProductoCruce = function(variedadProducto, tipoDescarte) {
      var variedadNormalizada = normalizarProducto(variedadProducto);
      var tipoNormalizado = String(tipoDescarte || "").toUpperCase().includes("PACKING")
        ? "DESCARTE PACKING"
        : "DESCARTE CAMPO";

      if (!variedadNormalizada) return "-";

      for (var b = 0; b < listaProductosBD.length; b++) {
        var productoBD = listaProductosBD[b];
        var descripcionBD = productoBD.descripcion;
        var tipoBD = descripcionBD.includes("DESCARTE PACKING")
          ? "DESCARTE PACKING"
          : "DESCARTE CAMPO";

        if (tipoBD === tipoNormalizado && normalizarProducto(descripcionBD) === variedadNormalizada) {
          return productoBD.codigo;
        }
      }

      return "-";
    };

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
    var mapaNombreVariedadCruce = {};
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
                var claveVariedadC = normalizarProducto(varC);
                if (!claveVariedadC) continue;
                var llaveC = fTxtC + "|" + claveVariedadC + "|CAMPO";
                if (!mapaNombreVariedadCruce[claveVariedadC]) mapaNombreVariedadCruce[claveVariedadC] = varC;
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
                var claveVariedadP = normalizarProducto(varP);
                if (!claveVariedadP) continue;
                var llaveP = fTxtP + "|" + claveVariedadP + "|PACKING";
                if (!mapaNombreVariedadCruce[claveVariedadP]) mapaNombreVariedadCruce[claveVariedadP] = varP;
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

            // La variedad se obtiene primero desde BD-PRODUCTOS para conservar
            // diferencias reales como ROSITA y ORGÁNICO ROSITA.
            var variedadBD = obtenerVariedadDesdeBD(colB);
            if (variedadBD) {
              variedadActual = variedadBD;
            } else if (colB.includes("ROSITA")) {
              variedadActual = colB.includes("ORGÁNICO") || colB.includes("ORGANICO")
                ? "ORGÁNICO ROSITA"
                : "ROSITA";
            } else if (colB.includes("RAYMI")) {
              variedadActual = colB.includes("ORGÁNICO") || colB.includes("ORGANICO")
                ? "ORGÁNICO RAYMI BLU"
                : "RAYMI BLU";
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
                .replace("ARÁNDANO ORGÁNICO FRESCO", "")
                .replace("ARANDANO FRESCO", "")
                .replace("ARANDANO ORGANICO FRESCO", "")
                .replace("- DESCARTE CAMPO", "")
                .replace("- DESCARTE PACKING", "")
                .replace("CAMPO", "")
                .replace("PACKING", "")
                .trim();
              if (limpiar) variedadActual = limpiar;
            }

            var claveVariedadStock = normalizarProducto(variedadActual);
            llaveStockActual = claveVariedadStock + "|" + tipoGrupoActual;
            if (claveVariedadStock && !mapaNombreVariedadCruce[claveVariedadStock]) mapaNombreVariedadCruce[claveVariedadStock] = variedadActual;
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
              
              var claveVariedadS = normalizarProducto(variedadActual);
              if (claveVariedadS) {
                var llaveS = fTxtS + "|" + claveVariedadS + "|" + tipoGrupoActual;
                if (!mapaNombreVariedadCruce[claveVariedadS]) mapaNombreVariedadCruce[claveVariedadS] = variedadActual;
                mapaSap[llaveS] = (mapaSap[llaveS] || 0) + cantidadS;
              }
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
      var claveVariedad = partes[1];
      var varNombre = mapaNombreVariedadCruce[claveVariedad] || claveVariedad;
      var tipo = partes[2];
      
      var totalSispacking = mapaSispacking[llave] || 0;
      var totalSap = mapaSap[llave] || 0;
      var diferencia = totalSap - totalSispacking;
      
      if (Math.abs(diferencia) <= 0.01) {
        diferencia = 0;
      }

      var codigoProducto = buscarCodigoProductoCruce(varNombre, tipo);
      
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
      var codP = buscarCodigoProductoCruce(vNombre, tGrupo);

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


// ==========================================
// MÓDULO DE DESPACHOS / CLIENTES
// ==========================================

function _obtenerHojaDespachos_() {
  var ss = SpreadsheetApp.openById("110fiUgfJizj_JR7RG3B-xnwfC3wx0NAlbAD4kie-oZ4");
  var sheet = ss.getSheetByName("DESPACHOS");

  if (!sheet) {
    sheet = ss.insertSheet("DESPACHOS");
    sheet.getRange(1, 1, 1, 10).setValues([[
      "ID", "FECHA", "EMPRESA", "CLIENTE", "CÓDIGO", "PRODUCTO", "KG", "USUARIO REGISTRO", "FECHA REGISTRO", "CORRELATIVO"
    ]]);
    sheet.setFrozenRows(1);
  } else {
    if (sheet.getRange(1, 10).getValue() !== "CORRELATIVO") {
      sheet.getRange(1, 10).setValue("CORRELATIVO");
    }
  }

  return sheet;
}

function _catalogoClientesDespachos_() {
  var ss = SpreadsheetApp.openById("110fiUgfJizj_JR7RG3B-xnwfC3wx0NAlbAD4kie-oZ4");
  var sheet = ss.getSheetByName("CLIENTES");
  if (!sheet) return [];

  var datos = sheet.getDataRange().getValues();
  var lista = [];
  for (var i = 1; i < datos.length; i++) {
    var id = String(datos[i][0] || "").trim();
    var cliente = String(datos[i][1] || "").trim();
    var estado = String(datos[i][4] || "ACTIVO").trim().toUpperCase();
    if (cliente && estado === "ACTIVO") lista.push({ id: id, cliente: cliente, estado: estado });
  }
  return lista;
}

function _catalogoProductosDespachos_() {
  var ss = SpreadsheetApp.openById("110fiUgfJizj_JR7RG3B-xnwfC3wx0NAlbAD4kie-oZ4");
  var sheet = ss.getSheetByName("BD-PRODUCTOS");
  if (!sheet) return [];

  var datos = sheet.getDataRange().getValues();
  var lista = [];
  for (var i = 1; i < datos.length; i++) {
    var codigo = String(datos[i][0] || "").trim();
    var descripcion = String(datos[i][1] || "").trim();
    if (codigo && descripcion) lista.push({ codigo: codigo, descripcion: descripcion });
  }
  return lista;
}

function obtenerCatalogosDespachos(token) {
  _exigirAdminDespachos_(token);
  return { exito: true, clientes: _catalogoClientesDespachos_(), productos: _catalogoProductosDespachos_() };
}

function _buscarProductoDespacho_(codigo) {
  var productos = _catalogoProductosDespachos_();
  var buscado = String(codigo || "").trim();
  for (var i = 0; i < productos.length; i++) if (productos[i].codigo === buscado) return productos[i];
  return null;
}

function _clienteActivoDespacho_(cliente) {
  var lista = _catalogoClientesDespachos_();
  var buscado = String(cliente || "").trim().toUpperCase();
  if (!buscado) return "";

  function normalizarNombreCliente_(valor) {
    return String(valor || "")
      .trim()
      .toUpperCase()
      .replace(/[.,]/g, " ")
      .replace(/\b(SAC|S\.A\.C|SA|S\.A|EIRL|E\.I\.R\.L)\b/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  var buscadoNormalizado = normalizarNombreCliente_(buscado);

  // Primero se prioriza coincidencia exacta.
  for (var i = 0; i < lista.length; i++) {
    if (lista[i].cliente.toUpperCase() === buscado) return lista[i].cliente;
  }

  // Luego se acepta el mismo nombre con razón social/sufijo empresarial.
  for (var j = 0; j < lista.length; j++) {
    var clienteNormalizado = normalizarNombreCliente_(lista[j].cliente);
    if (clienteNormalizado === buscadoNormalizado) return lista[j].cliente;
  }

  return "";
}

function guardarDespacho(token, datosDespacho) {
  var sesion = _exigirAdminDespachos_(token);
  datosDespacho = datosDespacho || {};

  var fechaTexto = String(datosDespacho.fecha || "").trim();
  var empresa = String(datosDespacho.empresa || "").trim().toUpperCase();
  var cliente = _clienteActivoDespacho_(datosDespacho.cliente);
  var codigo = String(datosDespacho.codigo || "").trim();
  var kg = Number(datosDespacho.kg || 0);

  if (!fechaTexto) return { exito: false, mensaje: "La fecha es obligatoria." };
  if (["AGA", "LARAMA", "ARENUVA"].indexOf(empresa) === -1) return { exito: false, mensaje: "Seleccione una empresa válida." };
  if (!cliente) return { exito: false, mensaje: "Seleccione un cliente activo." };
  if (!codigo) return { exito: false, mensaje: "Seleccione un producto." };
  if (!isFinite(kg) || kg <= 0) return { exito: false, mensaje: "Ingrese una cantidad de KG mayor a 0." };

  var producto = _buscarProductoDespacho_(codigo);
  if (!producto) return { exito: false, mensaje: "El código seleccionado no existe en BD-PRODUCTOS." };

  var partes = fechaTexto.split("-");
  if (partes.length !== 3) return { exito: false, mensaje: "La fecha del despacho no es válida." };
  var fecha = new Date(Number(partes[0]), Number(partes[1]) - 1, Number(partes[2]));
  if (isNaN(fecha.getTime())) return { exito: false, mensaje: "La fecha del despacho no es válida." };

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = _obtenerHojaDespachos_();
    var id = "D" + Utilities.formatString("%06d", Math.max(sheet.getLastRow(), 1));
    sheet.appendRow([id, fecha, empresa, cliente, producto.codigo, producto.descripcion, kg, sesion.usuario, new Date()]);
    return { exito: true, id: id, mensaje: "Despacho registrado correctamente." };
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function _normalizarFechaDespacho_(valor) {
  if (valor instanceof Date) return new Date(valor.getFullYear(), valor.getMonth(), valor.getDate());
  var texto = String(valor || "").trim();
  if (!texto) return null;
  var p = texto.indexOf("/") !== -1 ? texto.split("/") : texto.split("-");
  if (p.length !== 3) return null;
  return texto.indexOf("/") !== -1 ? new Date(Number(p[2]), Number(p[1]) - 1, Number(p[0])) : new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
}

function _fechaFiltroDespacho_(texto) {
  if (!texto) return null;
  var p = String(texto).split("-");
  if (p.length !== 3) return null;
  return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
}

function _grupoDespachoDesdeId_(id) {
  var texto = String(id || '').trim();
  if (!texto) return '';
  var m = texto.match(/^(.*)-\d+$/);
  return m ? m[1] : texto;
}

function _leerDespachosFiltrados_(fechaInicioStr, fechaFinStr, empresaFiltro, clienteFiltro, soloVigentesCliente) {
  var sheet = _obtenerHojaDespachos_();
  var data = sheet.getDataRange().getValues(); var bdSheet = SpreadsheetApp.openById('110fiUgfJizj_JR7RG3B-xnwfC3wx0NAlbAD4kie-oZ4').getSheetByName('BD-PRODUCTOS'); var bdCodigoMap = {}; if (bdSheet) { var bdData = bdSheet.getDataRange().getValues(); for (var b = 1; b < bdData.length; b++) { var bdCodigo = String(bdData[b][0] || '').trim(); var bdDesc = String(bdData[b][1] || '').trim(); if (bdCodigo && bdDesc) { var bdTipo = bdDesc.toUpperCase().indexOf('PACKING') !== -1 ? 'DESCARTE PACKING' : 'DESCARTE CAMPO'; bdCodigoMap[bdTipo + '|' + _normalizarVariedadDespacho_(bdDesc)] = bdCodigo; } } }
  var dInicio = _fechaFiltroDespacho_(fechaInicioStr);
  var dFin = _fechaFiltroDespacho_(fechaFinStr);
  var empresa = String(empresaFiltro || "TODAS").trim().toUpperCase();
  var cliente = String(clienteFiltro || "TODOS").trim().toUpperCase();
  var resultado = [];
  var ahora = new Date();
  var DURACION_24_HORAS = 24 * 60 * 60 * 1000;

  for (var i = 1; i < data.length; i++) {
    var fecha = _normalizarFechaDespacho_(data[i][1]);
    if (!fecha) continue;
    if (dInicio && fecha < dInicio) continue;
    if (dFin && fecha > dFin) continue;
    var empresaFila = String(data[i][2] || "").trim().toUpperCase();
    var clienteFila = String(data[i][3] || "").trim();
    if (empresa !== "TODAS" && empresaFila !== empresa) continue;
    if (cliente !== "TODOS" && clienteFila.toUpperCase() !== cliente) continue;

    var fechaRegistroObj = data[i][8] instanceof Date ? new Date(data[i][8].getTime()) : null;
    if (soloVigentesCliente) {
      if (!fechaRegistroObj) continue;
      var antiguedad = ahora.getTime() - fechaRegistroObj.getTime();
      if (antiguedad < 0 || antiguedad >= DURACION_24_HORAS) continue;
    }

    var productoTexto = String(data[i][5] || "").trim();
    var productoUpper = productoTexto.toUpperCase();
    var tipoDescarte = productoUpper.indexOf("PACKING") !== -1 ? "DESCARTE PACKING" : "DESCARTE CAMPO";
    var variedad = productoTexto
      .replace(/^ARANDANO FRESCO\s*/i, "")
      .replace(/\s*-\s*DESCARTE\s+(CAMPO|PACKING)\s*$/i, "")
      .trim();

    resultado.push({
      id: String(data[i][0] || ""),
      grupoDespacho: _grupoDespachoDesdeId_(data[i][0]),
      correlativo: String(data[i][9] || "").trim(),
      fecha: Utilities.formatDate(fecha, Session.getScriptTimeZone(), "dd/MM/yyyy"),
      fechaRegistro: fechaRegistroObj ? fechaRegistroObj.toISOString() : "",
      empresa: empresaFila,
      cliente: clienteFila,
      codigo: bdCodigoMap[tipoDescarte + '|' + _normalizarVariedadDespacho_(productoTexto)] || String(data[i][4] || ""),
      producto: productoTexto,
      variedad: variedad,
      tipo: tipoDescarte,
      kg: Number(data[i][6] || 0),
      usuario: String(data[i][7] || "")
    });
  }

  resultado.sort(function(a, b) {
    return a.fecha.split("/").reverse().join("").localeCompare(b.fecha.split("/").reverse().join(""));
  });
  return resultado;
}

function obtenerDespachosAdmin(token, fechaInicioStr, fechaFinStr, empresaFiltro, clienteFiltro) {
  _exigirAdminDespachos_(token);
  return { exito: true, datos: _leerDespachosFiltrados_(fechaInicioStr, fechaFinStr, empresaFiltro, clienteFiltro) };
}

function obtenerDespachosCliente(token, fechaInicioStr, fechaFinStr) {
  var sesion = _exigirClienteDespachos_(token);
  return { exito: true, cliente: sesion.cliente, datos: _leerDespachosFiltrados_(fechaInicioStr, fechaFinStr, "TODAS", sesion.cliente, true) };
}

// ==========================================
// DESPACHOS DESDE REPORTE DIARIO
// ==========================================

function _normalizarVariedadDespacho_(valor) {
  return String(valor || '')
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/ARANDANO ORGANICO FRESCO/g, 'ORGANICO').replace(/ARANDANO FRESCO/g, '')
    .replace(/OZBLU®?/g, '')
    .replace(/DESCARTE\s+(CAMPO|PACKING)/g, '')
    .replace(/DESCARTE/g, '')
    .replace(/[-]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function _agruparClaveDespacho_(item) {
  return [
    String(item.empresa || '').trim().toUpperCase(),
    String(item.tipo || '').replace(/^DESCARTE\s+/i, '').trim().toUpperCase(),
    String(item.codigo || '').trim(),
    _normalizarVariedadDespacho_(item.variedad)
  ].join('|');
}

function _obtenerBaseDespachoInterna_(fecha, empresaFiltro, clienteFiltro) {
  var reporte = obtenerDatosReportePorFechas(fecha, fecha);
  if (!reporte || !reporte.exito) {
    throw new Error((reporte && reporte.mensaje) || 'No se pudo obtener el reporte del día.');
  }

  var mapa = {};
  (reporte.datos || []).forEach(function(r) {
    var empresa = String(r.empresa || '').trim().toUpperCase();
    if (empresaFiltro && empresaFiltro !== 'TODAS' && empresa !== empresaFiltro) return;

    var tipo = String(r.presentacion || '').toUpperCase().indexOf('PACKING') !== -1 ? 'PACKING' : 'CAMPO';
    var codigo = String(r.codigo || '').trim();
    var variedad = String(r.variedad || '').trim();
    var peso = Number(r.peso || 0);
    if (!empresa || !codigo || !variedad || peso <= 0) return;

    var item = {
      empresa: empresa,
      tipo: tipo,
      codigo: codigo,
      variedad: variedad,
      disponible: peso
    };
    var clave = _agruparClaveDespacho_(item);

    if (!mapa[clave]) {
      mapa[clave] = {
        empresa: empresa,
        tipo: tipo,
        codigo: codigo,
        variedad: variedad,
        disponible: 0,
        despachado: 0,
        saldo: 0
      };
    }
    mapa[clave].disponible += peso;
  });

  var hoja = _obtenerHojaDespachos_();
  var data = hoja.getDataRange().getValues(); var bdSheet = SpreadsheetApp.openById('110fiUgfJizj_JR7RG3B-xnwfC3wx0NAlbAD4kie-oZ4').getSheetByName('BD-PRODUCTOS'); var bdCodigoMap = {}; if (bdSheet) { var bdData = bdSheet.getDataRange().getValues(); for (var b = 1; b < bdData.length; b++) { var bdCodigo = String(bdData[b][0] || '').trim(); var bdDesc = String(bdData[b][1] || '').trim(); if (bdCodigo && bdDesc) { var bdTipo = bdDesc.toUpperCase().indexOf('PACKING') !== -1 ? 'PACKING' : 'CAMPO'; bdCodigoMap[bdTipo + '|' + _normalizarVariedadDespacho_(bdDesc)] = bdCodigo; } } }
  var fechaClave = String(fecha || '').trim();

  for (var i = 1; i < data.length; i++) {
    var fechaFila = _normalizarFechaDespacho_(data[i][1]);
    if (!fechaFila) continue;
    var fechaFilaTexto = Utilities.formatDate(fechaFila, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (fechaFilaTexto !== fechaClave) continue;

    var empresaFila = String(data[i][2] || '').trim().toUpperCase();
    if (empresaFiltro && empresaFiltro !== 'TODAS' && empresaFila !== empresaFiltro) continue;

    var codigoFila = String(data[i][4] || '').trim();
    var productoFila = String(data[i][5] || '').trim();
    var kgFila = Number(data[i][6] || 0);
    if (!empresaFila || !codigoFila || kgFila <= 0) continue;

    var tipoFila = productoFila.toUpperCase().indexOf('PACKING') !== -1 ? 'PACKING' : 'CAMPO'; codigoFila = bdCodigoMap[tipoFila + '|' + _normalizarVariedadDespacho_(productoFila)] || codigoFila;
    var variedadFila = productoFila
      .replace(/^ARANDANO FRESCO\s*/i, '')
      .replace(/\s*-\s*DESCARTE\s+(CAMPO|PACKING)\s*$/i, '')
      .replace(/^OZBLU®?\s*/i, '')
      .trim();

    var itemFila = {
      empresa: empresaFila,
      tipo: tipoFila,
      codigo: codigoFila,
      variedad: variedadFila
    };
    var claveFila = _agruparClaveDespacho_(itemFila);

    if (!mapa[claveFila]) continue;
    mapa[claveFila].despachado += kgFila;
  }

  var resultado = [];
  Object.keys(mapa).forEach(function(clave) {
    var item = mapa[clave];
    item.saldo = Math.max(0, Number(item.disponible || 0) - Number(item.despachado || 0));
    resultado.push(item);
  });

  resultado.sort(function(a, b) {
    return (a.empresa + a.tipo + a.variedad + a.codigo).localeCompare(b.empresa + b.tipo + b.variedad + b.codigo);
  });

  return resultado;
}

function obtenerDatosClienteDespacho(cliente) {
  var buscado = String(cliente || '').trim();
  if (!buscado) return { exito: false, mensaje: 'Cliente vacío.' };

  var ss = SpreadsheetApp.openById('110fiUgfJizj_JR7RG3B-xnwfC3wx0NAlbAD4kie-oZ4');
  var sheet = ss.getSheetByName('CLIENTES');
  if (!sheet) return { exito: false, mensaje: 'No se encontró la pestaña CLIENTES.' };

  var datos = sheet.getDataRange().getValues();
  var buscadoUpper = buscado.toUpperCase();

  function normalizarCliente_(valor) {
    return String(valor || '')
      .trim()
      .toUpperCase()
      .replace(/[.,]/g, ' ')
      .replace(/\b(SAC|S\.A\.C|SA|S\.A|EIRL|E\.I\.R\.L)\b/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  // Primero coincidencia exacta por razón social.
  for (var i = 1; i < datos.length; i++) {
    var clienteFila = String(datos[i][1] || '').trim();
    var estado = String(datos[i][4] || 'ACTIVO').trim().toUpperCase();
    if (estado !== 'ACTIVO') continue;
    if (clienteFila.toUpperCase() === buscadoUpper) {
      return {
        exito: true,
        cliente: clienteFila,
        documento: String(datos[i][2] || '').trim(),
        direccion: String(datos[i][3] || '').trim(),
        condicionPago: String(datos[i][5] || 'CONTADO').trim() || 'CONTADO'
      };
    }
  }

  // Luego acepta nombres cortos como "GRUPO CLUNA" vs "GRUPO CLUNA E.I.R.L.".
  var buscadoNormalizado = normalizarCliente_(buscado);
  for (var j = 1; j < datos.length; j++) {
    var clienteFila2 = String(datos[j][1] || '').trim();
    var estado2 = String(datos[j][4] || 'ACTIVO').trim().toUpperCase();
    if (estado2 !== 'ACTIVO') continue;
    if (normalizarCliente_(clienteFila2) === buscadoNormalizado) {
      return {
        exito: true,
        cliente: clienteFila2,
        documento: String(datos[j][2] || '').trim(),
        direccion: String(datos[j][3] || '').trim(),
        condicionPago: String(datos[j][5] || 'CONTADO').trim() || 'CONTADO'
      };
    }
  }

  return { exito: false, mensaje: 'No se encontró el cliente en CLIENTES.' };
}

/**
 * Obtiene la base disponible para despachar acumulando todo el rango.
 * El saldo de cada producto es global y no depende del cliente seleccionado;
 * la empresa solo limita los productos cuando el usuario selecciona una empresa concreta.
 */
function obtenerBaseDespachoPorFechas(token, fechaInicio, fechaFin, empresaFiltro, cliente) {
  _exigirAdminDespachos_(token);
  var inicio = String(fechaInicio || '').trim();
  var fin = String(fechaFin || '').trim();
  var clienteActivo = _clienteActivoDespacho_(cliente);
  if (!clienteActivo) throw new Error('Seleccione un cliente activo.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(inicio) || !/^\d{4}-\d{2}-\d{2}$/.test(fin)) {
    throw new Error('El rango de fechas seleccionado no es válido.');
  }
  if (inicio > fin) throw new Error('La fecha de inicio no puede ser mayor que la fecha de fin.');

  return {
    exito: true,
    cliente: clienteActivo,
    fechaInicio: inicio,
    fechaFin: fin,
    datos: _obtenerBaseDespachoInternaPorRango_(inicio, fin, String(empresaFiltro || 'TODAS').trim().toUpperCase(), clienteActivo)
  };
}

/**
 * Construye la base disponible consolidada de un rango de fechas.
 * Disponible = kilos del reporte del rango.
 * Despachado = kilos ya registrados en DESPACHOS dentro del mismo rango.
 */
function _obtenerBaseDespachoInternaPorRango_(fechaInicio, fechaFin, empresaFiltro, clienteFiltro) {
  var reporte = obtenerDatosReportePorFechas(fechaInicio, fechaFin);
  if (!reporte || !reporte.exito) {
    throw new Error((reporte && reporte.mensaje) || 'No se pudo obtener el reporte del rango.');
  }

  var mapa = {};
  (reporte.datos || []).forEach(function(r) {
    var empresa = String(r.empresa || '').trim().toUpperCase();
    if (empresaFiltro && empresaFiltro !== 'TODAS' && empresa !== empresaFiltro) return;

    var tipo = String(r.presentacion || '').toUpperCase().indexOf('PACKING') !== -1 ? 'PACKING' : 'CAMPO';
    var codigo = String(r.codigo || '').trim();
    var variedad = String(r.variedad || '').trim();
    var peso = Number(r.peso || 0);
    if (!empresa || !codigo || !variedad || peso <= 0) return;

    var item = { empresa: empresa, tipo: tipo, codigo: codigo, variedad: variedad, disponible: peso };
    var clave = _agruparClaveDespacho_(item);

    if (!mapa[clave]) {
      mapa[clave] = {
        empresa: empresa,
        tipo: tipo,
        codigo: codigo,
        variedad: variedad,
        disponible: 0,
        despachado: 0,
        saldo: 0
      };
    }
    mapa[clave].disponible += peso;
  });

  var hoja = _obtenerHojaDespachos_();
  var data = hoja.getDataRange().getValues(); var bdSheet = SpreadsheetApp.openById('110fiUgfJizj_JR7RG3B-xnwfC3wx0NAlbAD4kie-oZ4').getSheetByName('BD-PRODUCTOS'); var bdCodigoMap = {}; if (bdSheet) { var bdData = bdSheet.getDataRange().getValues(); for (var b = 1; b < bdData.length; b++) { var bdCodigo = String(bdData[b][0] || '').trim(); var bdDesc = String(bdData[b][1] || '').trim(); if (bdCodigo && bdDesc) { var bdTipo = bdDesc.toUpperCase().indexOf('PACKING') !== -1 ? 'PACKING' : 'CAMPO'; bdCodigoMap[bdTipo + '|' + _normalizarVariedadDespacho_(bdDesc)] = bdCodigo; } } }
  var dInicio = _fechaFiltroDespacho_(fechaInicio);
  var dFin = _fechaFiltroDespacho_(fechaFin);
  var clienteUpper = String(clienteFiltro || 'TODOS').trim().toUpperCase();

  for (var i = 1; i < data.length; i++) {
    var fechaFila = _normalizarFechaDespacho_(data[i][1]);
    if (!fechaFila) continue;
    if (dInicio && fechaFila < dInicio) continue;
    if (dFin && fechaFila > dFin) continue;

    var empresaFila = String(data[i][2] || '').trim().toUpperCase();
    if (empresaFiltro && empresaFiltro !== 'TODAS' && empresaFila !== empresaFiltro) continue;

    var codigoFila = String(data[i][4] || '').trim();
    var productoFila = String(data[i][5] || '').trim();
    var kgFila = Number(data[i][6] || 0);
    if (!empresaFila || !codigoFila || kgFila <= 0) continue;

    var tipoFila = productoFila.toUpperCase().indexOf('PACKING') !== -1 ? 'PACKING' : 'CAMPO'; codigoFila = bdCodigoMap[tipoFila + '|' + _normalizarVariedadDespacho_(productoFila)] || codigoFila;
    var variedadFila = productoFila
      .replace(/^ARANDANO FRESCO\s*/i, '')
      .replace(/\s*-\s*DESCARTE\s+(CAMPO|PACKING)\s*$/i, '')
      .replace(/^OZBLU®?\s*/i, '')
      .trim();

    var itemFila = { empresa: empresaFila, tipo: tipoFila, codigo: codigoFila, variedad: variedadFila };
    var claveFila = _agruparClaveDespacho_(itemFila);
    if (!mapa[claveFila]) continue;
    mapa[claveFila].despachado += kgFila;
  }

  var resultado = [];
  Object.keys(mapa).forEach(function(clave) {
    var item = mapa[clave];
    item.saldo = Math.max(0, Number(item.disponible || 0) - Number(item.despachado || 0));
    resultado.push(item);
  });

  resultado.sort(function(a, b) {
    return (a.empresa + a.tipo + a.variedad + a.codigo).localeCompare(b.empresa + b.tipo + b.variedad + b.codigo);
  });

  return resultado;
}

// Compatibilidad con cualquier llamada anterior de una sola fecha.
function obtenerBaseDespachoPorFecha(token, fecha, empresaFiltro, cliente) {
  return obtenerBaseDespachoPorFechas(token, fecha, fecha, empresaFiltro, cliente);
}

function guardarDespachosMasivosPorRango(token, fechaInicio, fechaFin, cliente, registros) {
  var sesion = _exigirAdminDespachos_(token);
  var inicio = String(fechaInicio || '').trim();
  var fin = String(fechaFin || '').trim();
  var clienteActivo = _clienteActivoDespacho_(cliente);
  var lista = Array.isArray(registros) ? registros : [];

  if (!clienteActivo) return { exito: false, mensaje: 'Seleccione un cliente activo.' };
  if (!/^\d{4}-\d{2}-\d{2}$/.test(inicio) || !/^\d{4}-\d{2}-\d{2}$/.test(fin)) return { exito: false, mensaje: 'El rango de fechas seleccionado no es válido.' };
  if (inicio > fin) return { exito: false, mensaje: 'La fecha de inicio no puede ser mayor que la fecha de fin.' };
  if (!lista.length) return { exito: false, mensaje: 'No hay cantidades para registrar.' };

  var lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    var base = _obtenerBaseDespachoInternaPorRango_(inicio, fin, 'TODAS', clienteActivo);
    var mapaBase = {};
    base.forEach(function(item) {
      mapaBase[_agruparClaveDespacho_(item)] = item;
    });

    var acumuladoSolicitud = {};
    var registrosValidos = [];

    lista.forEach(function(r) {
      var empresa = String(r.empresa || '').trim().toUpperCase();
      var tipo = String(r.tipo || '').trim().toUpperCase();
      var codigo = String(r.codigo || '').trim();
      var variedad = String(r.variedad || '').trim();
      var kg = Number(r.kg || 0);
      if (!empresa || !tipo || !codigo || !variedad || !(kg > 0)) return;

      var clave = _agruparClaveDespacho_({ empresa: empresa, tipo: tipo, codigo: codigo, variedad: variedad });
      if (!mapaBase[clave]) {
        throw new Error('El producto ' + codigo + ' - ' + variedad + ' ya no está disponible en el rango seleccionado.');
      }

      acumuladoSolicitud[clave] = Number(acumuladoSolicitud[clave] || 0) + kg;
      if (acumuladoSolicitud[clave] > Number(mapaBase[clave].saldo || 0) + 0.000001) {
        throw new Error('La cantidad a despachar de ' + codigo + ' - ' + variedad + ' supera el saldo disponible de ' + Number(mapaBase[clave].saldo || 0).toFixed(3) + ' KG.');
      }

      registrosValidos.push({
        empresa: empresa,
        tipo: tipo,
        codigo: codigo,
        variedad: variedad,
        kg: kg
      });
    });

    if (!registrosValidos.length) return { exito: false, mensaje: 'No hay cantidades válidas para registrar.' };

    var hoja = _obtenerHojaDespachos_();
    var ahora = new Date();
    var filas = [];
    // El despacho queda registrado con la fecha final del rango seleccionado.
    var prefijo = 'D-' + fin.replace(/-/g, '');
    var baseId = ahora.getTime();

    registrosValidos.forEach(function(r, index) {
      filas.push([
        prefijo + '-' + baseId + '-' + (index + 1),
        fin,
        r.empresa,
        clienteActivo,
        r.codigo,
        'ARANDANO FRESCO ' + r.variedad + ' - DESCARTE ' + r.tipo,
        r.kg,
        sesion.usuario || 'ADMIN',
        ahora,
        ''
      ]);
    });

    hoja.getRange(hoja.getLastRow() + 1, 1, filas.length, filas[0].length).setValues(filas);

    return {
      exito: true,
      mensaje: 'Despachos registrados correctamente para ' + clienteActivo + ' con fecha ' + fin.split('-').reverse().join('/') + '.',
      registros: filas.length,
      cliente: clienteActivo,
      fechaInicio: inicio,
      fechaFin: fin
    };
  } finally {
    lock.releaseLock();
  }
}

// Compatibilidad con el guardado anterior de una sola fecha.
function guardarDespachosMasivos(token, fecha, cliente, registros) {
  var sesion = _exigirAdminDespachos_(token);
  var fechaTexto = String(fecha || '').trim();
  var clienteActivo = _clienteActivoDespacho_(cliente);
  var lista = Array.isArray(registros) ? registros : [];

  if (!clienteActivo) return { exito: false, mensaje: 'Seleccione un cliente activo.' };
  if (!/^\d{4}-\d{2}-\d{2}$/.test(fechaTexto)) return { exito: false, mensaje: 'La fecha seleccionada no es válida.' };
  if (!lista.length) return { exito: false, mensaje: 'No hay cantidades para registrar.' };

  var lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    var base = _obtenerBaseDespachoInterna_(fechaTexto, 'TODAS', clienteActivo);
    var mapaBase = {};
    base.forEach(function(item) {
      mapaBase[_agruparClaveDespacho_(item)] = item;
    });

    var acumuladoSolicitud = {};
    var registrosValidos = [];

    lista.forEach(function(r) {
      var empresa = String(r.empresa || '').trim().toUpperCase();
      var tipo = String(r.tipo || '').trim().toUpperCase();
      var codigo = String(r.codigo || '').trim();
      var variedad = String(r.variedad || '').trim();
      var kg = Number(r.kg || 0);
      if (!empresa || !tipo || !codigo || !variedad || !(kg > 0)) return;

      var clave = _agruparClaveDespacho_({ empresa: empresa, tipo: tipo, codigo: codigo, variedad: variedad });
      if (!mapaBase[clave]) {
        throw new Error('El producto ' + codigo + ' - ' + variedad + ' ya no está disponible en el reporte seleccionado.');
      }

      acumuladoSolicitud[clave] = Number(acumuladoSolicitud[clave] || 0) + kg;
      if (acumuladoSolicitud[clave] > Number(mapaBase[clave].saldo || 0) + 0.000001) {
        throw new Error('La cantidad a despachar de ' + codigo + ' - ' + variedad + ' supera el saldo disponible de ' + Number(mapaBase[clave].saldo || 0).toFixed(3) + ' KG.');
      }

      registrosValidos.push({
        empresa: empresa,
        tipo: tipo,
        codigo: codigo,
        variedad: variedad,
        kg: kg
      });
    });

    if (!registrosValidos.length) return { exito: false, mensaje: 'No hay cantidades válidas para registrar.' };

    var hoja = _obtenerHojaDespachos_();
    var ahora = new Date();
    var filas = [];
    var prefijo = 'D-' + fechaTexto.replace(/-/g, '');
    var baseId = ahora.getTime();

    registrosValidos.forEach(function(r, index) {
      filas.push([
        prefijo + '-' + baseId + '-' + (index + 1),
        fechaTexto,
        r.empresa,
        clienteActivo,
        r.codigo,
        'ARANDANO FRESCO ' + r.variedad + ' - DESCARTE ' + r.tipo,
        r.kg,
        sesion.usuario || 'ADMIN',
        ahora,
        ''
      ]);
    });

    hoja.getRange(hoja.getLastRow() + 1, 1, filas.length, filas[0].length).setValues(filas);

    return {
      exito: true,
      mensaje: 'Despachos registrados correctamente para ' + clienteActivo + '.',
      registros: filas.length,
      cliente: clienteActivo
    };
  } finally {
    lock.releaseLock();
  }
}

function asignarCorrelativoDespacho(token, grupoDespacho, empresaOcorrelativo, correlativoOpcional) {
  _exigirAdminDespachos_(token);
  var grupo = String(grupoDespacho || '').trim();
  var empresa = '';
  var nuevo = '';

  // Nueva firma: token, grupo, empresa, correlativo.
  if (correlativoOpcional !== undefined) {
    empresa = String(empresaOcorrelativo || '').trim().toUpperCase();
    nuevo = String(correlativoOpcional || '').trim();
  } else {
    // Compatibilidad con la firma anterior.
    nuevo = String(empresaOcorrelativo || '').trim();
  }

  if (!grupo) return { exito: false, mensaje: 'No se identificó el despacho seleccionado.' };
  if (correlativoOpcional !== undefined && ['AGA', 'LARAMA', 'ARENUVA'].indexOf(empresa) === -1) {
    return { exito: false, mensaje: 'No se identificó una empresa válida para el despacho.' };
  }
  if (!nuevo) return { exito: false, mensaje: 'Ingrese un correlativo.' };

  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sheet = _obtenerHojaDespachos_();
    var data = sheet.getDataRange().getValues();
    var filasGrupo = [];

    for (var i = 1; i < data.length; i++) {
      var id = String(data[i][0] || '');
      var grupoFila = _grupoDespachoDesdeId_(id);
      var empresaFila = String(data[i][2] || '').trim().toUpperCase();
      if (grupoFila === grupo && (!empresa || empresaFila === empresa)) filasGrupo.push(i + 1);
    }

    if (!filasGrupo.length) return { exito: false, mensaje: 'No se encontró el despacho seleccionado' + (empresa ? ' para ' + empresa : '') + '.' };

    var buscado = nuevo.toUpperCase();
    for (var j = 1; j < data.length; j++) {
      var corrExistente = String(data[j][9] || '').trim().toUpperCase();
      if (!corrExistente || corrExistente !== buscado) continue;

      var grupoExistente = _grupoDespachoDesdeId_(data[j][0]);
      var empresaExistente = String(data[j][2] || '').trim().toUpperCase();
      var mismoDespacho = grupoExistente === grupo && (!empresa || empresaExistente === empresa);
      if (!mismoDespacho) {
        if (empresa && empresaExistente !== empresa) continue;
        return { exito: false, mensaje: 'El correlativo ' + nuevo + ' ya está asignado a otro despacho' + (empresa ? ' de ' + empresa : '') + '.' };
      }
    }

    filasGrupo.forEach(function(fila) {
      sheet.getRange(fila, 10).setValue(nuevo);
    });

    return { exito: true, correlativo: nuevo, filas: filasGrupo.length, empresa: empresa || 'TODAS' };
  } finally {
    lock.releaseLock();
  }
}
