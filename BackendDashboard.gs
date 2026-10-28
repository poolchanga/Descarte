/**
 * ============================================================
 * DASHBOARD - BACKEND
 * ============================================================
 * Regla de oro:
 * - Este archivo SOLO alimenta el Dashboard.
 * - No modifica Ventas, Guías Provisionales, Reporte ni Cruce.
 * - Lee HISTORIAL_GUIAS y VENTAS del maestro existente.
 * ============================================================
 */

function getDashboardData(empresa, fechaInicio, fechaFin, variedad, producto) {
  try {
    var ss = SpreadsheetApp.openById(_dashboardSpreadsheetId_());
    var sheetVentas = ss.getSheetByName('VENTAS');

    var perfilProducto = String(producto || 'ARANDANOS').toUpperCase() === 'UVA' ? 'UVA' : 'ARANDANOS';
    if (perfilProducto === 'UVA' && String(empresa || 'TODAS').toUpperCase() === 'LARAMA') return { exito: false, mensaje: 'El perfil UVA solo permite AGA y ARENUVA.' };
    var filtroEmpresa = _dashboardNormalizar_(empresa || 'TODAS');
    var filtroVariedad = _dashboardNormalizar_(variedad || 'TODAS');
    var inicio = _dashboardFecha_(fechaInicio, false);
    var fin = _dashboardFecha_(fechaFin, true);

    var guias = [];
    var ventas = _dashboardLeerVentas_(sheetVentas, filtroEmpresa, inicio, fin, perfilProducto);

    if (filtroVariedad && filtroVariedad !== 'TODAS') {
      ventas = ventas.filter(function(v) {
        var vVariedad = _dashboardNormalizar_(v.variedad);
        return vVariedad === filtroVariedad ||
               vVariedad.indexOf(filtroVariedad) >= 0 ||
               filtroVariedad.indexOf(vVariedad) >= 0;
      });
    }

    var reporte = _dashboardLeerReporte_(empresa || 'TODAS', fechaInicio, fechaFin, filtroVariedad, perfilProducto);
    var stock = _dashboardLeerStockSAP_(empresa || 'TODAS', fechaInicio, fechaFin, filtroVariedad, reporte.cruceMapa, perfilProducto);

    var kpis = _dashboardKpis_(guias, ventas);

    kpis.kilosDescarte = reporte.totalKg;
    kpis.kilosCampo = reporte.campo;
    kpis.kilosPacking = reporte.packing;
    kpis.stockActual = stock.total;
    kpis.stockPendienteVenta = stock.pendienteVenta;
    kpis.diferenciasStock = stock.diferenciasTotal;

    var ventasPorVariedad = {};
    ventas.forEach(function(v) {
      var key = _dashboardNormalizar_(v.variedad || 'SIN VARIEDAD');
      if (!ventasPorVariedad[key]) ventasPorVariedad[key] = 0;
      ventasPorVariedad[key] += v.kg;
    });

    stock.controlStock = stock.porVariedad.map(function(item) {
      var key = _dashboardNormalizar_(item.variedad);
      var vendido = ventasPorVariedad[key] || 0;
      var pendiente = item.stock - vendido;
      return {
        variedad: item.variedad,
        stock: item.stock,
        vendido: vendido,
        pendiente: pendiente,
        estado: pendiente > 0.01 ? 'PENDIENTE DE VENTA' : (pendiente < -0.01 ? 'SALIDA SUPERIOR AL STOCK' : 'OK')
      };
    });

    kpis.stockPendienteVenta = stock.controlStock.reduce(function(suma, item) {
      return suma + Math.max(item.pendiente, 0);
    }, 0);

    return {
      exito: true,

      filtros: {
        empresa: filtroEmpresa || 'TODAS',
        variedad: filtroVariedad || 'TODAS',
        fechaInicio: fechaInicio || '',
        fechaFin: fechaFin || ''
      },

      kpis: kpis,

      porEmpresa: [],
      porFecha: [],
      porProducto: [],
      porCliente: _dashboardTopClientes_(ventas, 5),

      porFechaReporte: reporte.porFecha,
      porEmpresaReporte: reporte.porEmpresa,
      porVariedadReporte: reporte.porVariedad,

      stockPorEmpresa: stock.porEmpresa,
      stockPorProducto: stock.porProducto,
      stockPorVariedad: stock.porVariedad,
      controlStock: stock.controlStock,
      diferencias: stock.diferencias,
      variedadesDisponibles: stock.variedadesDisponibles,

      recientesGuias: _dashboardRecientesGuias_(guias, 6),
      recientesVentas: _dashboardRecientesVentas_(ventas, 6)
    };

  } catch (error) {
    return {
      exito: false,
      mensaje: 'Error al cargar Dashboard: ' + error.toString()
    };
  }
}
function _dashboardSpreadsheetId_() {
  // Reutiliza el ID global del proyecto si ya existe.
  // Si no existe, usa el maestro actual del sistema.
  try {
    if (typeof ID_SS_MAESTRO_GUIAS !== 'undefined' && ID_SS_MAESTRO_GUIAS) {
      return ID_SS_MAESTRO_GUIAS;
    }
  } catch (e) {}

  return '110fiUgfJizj_JR7RG3B-xnwfC3wx0NAlbAD4kie-oZ4';
}

function _dashboardNormalizar_(valor) {
  return String(valor == null ? '' : valor)
    .trim()
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

function _dashboardFechaDesdeCelda_(celda) {
  if (!celda) return null;

  if (Object.prototype.toString.call(celda) === '[object Date]') {
    var d = new Date(celda.getTime());
    d.setHours(0, 0, 0, 0);
    return isNaN(d.getTime()) ? null : d;
  }

  var texto = String(celda).trim();
  if (!texto) return null;

  var p = texto.indexOf('/') >= 0 ? texto.split('/') : texto.split('-');
  if (p.length !== 3) return null;

  var d2;
  if (texto.indexOf('/') >= 0) {
    d2 = new Date(Number(p[2]), Number(p[1]) - 1, Number(p[0]));
  } else {
    d2 = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
  }

  d2.setHours(0, 0, 0, 0);
  return isNaN(d2.getTime()) ? null : d2;
}

function _dashboardFecha_(valor, finDelDia) {
  if (!valor) return null;

  if (Object.prototype.toString.call(valor) === '[object Date]') {
    var d0 = new Date(valor.getTime());
    if (finDelDia) d0.setHours(23, 59, 59, 999);
    else d0.setHours(0, 0, 0, 0);
    return isNaN(d0.getTime()) ? null : d0;
  }

  var texto = String(valor).trim();
  var partes = texto.split('-');

  if (partes.length === 3) {
    var d = new Date(
      Number(partes[0]),
      Number(partes[1]) - 1,
      Number(partes[2])
    );

    if (finDelDia) d.setHours(23, 59, 59, 999);
    else d.setHours(0, 0, 0, 0);

    return isNaN(d.getTime()) ? null : d;
  }

  return null;
}

function _dashboardEnRango_(fecha, inicio, fin) {
  if (!inicio && !fin) return true;
  if (!fecha) return false;

  var d = Object.prototype.toString.call(fecha) === '[object Date]'
    ? new Date(fecha.getTime())
    : new Date(fecha);

  if (isNaN(d.getTime())) return false;

  if (inicio && d < inicio) return false;
  if (fin && d > fin) return false;

  return true;
}

function _dashboardEmpresaCoincide_(filaEmpresa, filtroEmpresa) {
  if (!filtroEmpresa || filtroEmpresa === 'TODAS') return true;

  var fila = _dashboardNormalizar_(filaEmpresa);

  var aliases = {
    AGA: ['AGA', 'AGRICOLA ANDREA', 'AGRICOLA ANDREA SAC'],
    LARAMA: ['LARAMA', 'LARAMA BERRIES', 'LARAMA BERRIES SAC'],
    ARENUVA: ['ARENUVA', 'ARENUVA SAC']
  };

  var permitidos = aliases[filtroEmpresa] || [filtroEmpresa];

  return permitidos.some(function(alias) {
    var a = _dashboardNormalizar_(alias);
    return fila === a || fila.indexOf(a) >= 0 || a.indexOf(fila) >= 0;
  });
}

function _dashboardNumero_(valor) {
  if (typeof valor === 'number') return isNaN(valor) ? 0 : valor;

  var s = String(valor == null ? '' : valor).trim().replace(/\s/g, '');
  if (!s) return 0;

  if (s.indexOf(',') >= 0 && s.indexOf('.') >= 0) {
    s = s.replace(/,/g, '');
  } else if (s.indexOf(',') >= 0) {
    s = s.replace(/,/g, '.');
  }

  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function _dashboardFechaTexto_(fecha) {
  if (!fecha) return '';

  var d = Object.prototype.toString.call(fecha) === '[object Date]'
    ? fecha
    : new Date(fecha);

  if (isNaN(d.getTime())) return '';

  return Utilities.formatDate(
    d,
    Session.getScriptTimeZone(),
    'dd/MM/yyyy'
  );
}

function _dashboardFechaKey_(fecha) {
  if (!fecha) return '';

  var d = Object.prototype.toString.call(fecha) === '[object Date]'
    ? fecha
    : new Date(fecha);

  if (isNaN(d.getTime())) return '';

  return Utilities.formatDate(
    d,
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );
}

function _dashboardLeerGuias_(sheet, filtroEmpresa, inicio, fin) {
  var resultado = [];
  if (!sheet || sheet.getLastRow() <= 1) return resultado;

  // HISTORIAL_GUIAS:
  // A correlativo | B empresa | C RUC | D cliente | E documento
  // F dirección | G fecha emisión | H vencimiento | I condición
  // J autorizado | K código | L descripción | M cantidad
  var datos = sheet.getRange(
    2, 1, sheet.getLastRow() - 1, 13
  ).getValues();

  var vistos = {};

  for (var i = 0; i < datos.length; i++) {
    var fila = datos[i];

    var correlativo = String(fila[0] || '').trim();
    var empresa = String(fila[1] || '').trim();
    var cliente = String(fila[3] || '').trim();
    var fecha = fila[6];
    var codigo = String(fila[10] || '').trim();
    var descripcion = String(fila[11] || '').trim();
    var kg = _dashboardNumero_(fila[12]);

    if (!correlativo || !empresa || !codigo || kg <= 0) continue;
    if (familiaPorCodigo[codigo] !== perfilVentas) continue;
    if (!_dashboardEmpresaCoincide_(empresa, filtroEmpresa)) continue;
    if (!_dashboardEnRango_(fecha, inicio, fin)) continue;

    var key = correlativo + '|' + codigo + '|' + i;

    resultado.push({
      correlativo: correlativo,
      empresa: empresa,
      cliente: cliente,
      fecha: _dashboardFechaTexto_(fecha),
      fechaKey: _dashboardFechaKey_(fecha),
      codigo: codigo,
      descripcion: descripcion,
      kg: kg,
      tipo: _dashboardTipoDescarte_(descripcion),
      _key: key
    });

    vistos[correlativo] = true;
  }

  return resultado;
}

function _dashboardLeerVentas_(sheet, filtroEmpresa, inicio, fin, producto) {
  var resultado = [];
  if (!sheet || sheet.getLastRow() <= 1) return resultado;

  // VENTAS:
  // A correlativo | B fecha | C RUC | D cliente | E código
  // F variedad | G KG | H empresa | I link origen
  var datos = sheet.getRange(
    2, 1, sheet.getLastRow() - 1, 9
  ).getValues();
  // Leer la columna de códigos como texto visible para evitar pérdida de precisión
  // en códigos largos que Sheets pueda interpretar como números.
  var codigosVisibles = sheet.getRange(2, 5, sheet.getLastRow() - 1, 1).getDisplayValues();

  // Aislar las ventas por familia usando el código de BD-PRODUCTOS.
  var familiaPorCodigo = {};
  try {
    var ssProductos = SpreadsheetApp.openById(_dashboardSpreadsheetId_());
    var shProductos = ssProductos.getSheetByName('BD-PRODUCTOS');
    if (shProductos && shProductos.getLastRow() > 1) {
      var bd = shProductos.getDataRange().getDisplayValues();
      for (var b = 1; b < bd.length; b++) {
        var codigoBD = String(bd[b][0] || '').trim();
        var descBD = String(bd[b][1] || '').toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
        if (!codigoBD || !descBD) continue;
        if (/\bUVA(S)?\s+FRESCA(S)?\b/.test(descBD)) familiaPorCodigo[codigoBD] = 'UVA';
        else if (/\bARANDANO(S)?\b/.test(descBD)) familiaPorCodigo[codigoBD] = 'ARANDANOS';
      }
    }
  } catch (e) {}
  var perfilVentas = String(producto || 'ARANDANOS').toUpperCase() === 'UVA' ? 'UVA' : 'ARANDANOS';

  for (var i = 0; i < datos.length; i++) {
    var fila = datos[i];

    var correlativo = String(fila[0] || '').trim();
    var fecha = fila[1];
    var ruc = String(fila[2] || '').trim();
    var cliente = String(fila[3] || '').trim();
    var codigo = String((codigosVisibles[i] && codigosVisibles[i][0]) || fila[4] || '').trim();
    var variedad = String(fila[5] || '').trim();
    var kg = _dashboardNumero_(fila[6]);
    var empresa = String(fila[7] || '').trim();

    if (!correlativo || !empresa || !codigo || kg <= 0) continue;
    if (!_dashboardEmpresaCoincide_(empresa, filtroEmpresa)) continue;
    if (!_dashboardEnRango_(fecha, inicio, fin)) continue;

    resultado.push({
      correlativo: correlativo,
      empresa: empresa,
      cliente: cliente,
      ruc: ruc,
      fecha: _dashboardFechaTexto_(fecha),
      fechaKey: _dashboardFechaKey_(fecha),
      codigo: codigo,
      variedad: variedad,
      kg: kg
    });
  }

  return resultado;
}

function _dashboardTipoDescarte_(descripcion) {
  var texto = _dashboardNormalizar_(descripcion);

  if (texto.indexOf('PACKING') >= 0) return 'PACKING';
  if (texto.indexOf('CAMPO') >= 0 || texto.indexOf('GRANEL') >= 0) return 'CAMPO';

  return 'OTRO';
}

function _dashboardKpis_(guias, ventas) {
  var kilosProcesados = 0;
  var kilosPacking = 0;
  var kilosCampo = 0;

  var guiasSet = {};
  var ventasSet = {};
  var productosSet = {};
  var empresasSet = {};

  guias.forEach(function(g) {
    kilosProcesados += g.kg;
    guiasSet[g.correlativo] = true;
    productosSet[g.codigo] = true;
    empresasSet[_dashboardNormalizar_(g.empresa)] = true;

    if (g.tipo === 'PACKING') kilosPacking += g.kg;
    else if (g.tipo === 'CAMPO') kilosCampo += g.kg;
  });

  var kilosVendidos = 0;
  var clientesSet = {};

  ventas.forEach(function(v) {
    kilosVendidos += v.kg;
    ventasSet[v.correlativo] = true;
    productosSet[v.codigo] = true;
    empresasSet[_dashboardNormalizar_(v.empresa)] = true;

    if (v.cliente) clientesSet[_dashboardNormalizar_(v.cliente)] = true;
  });

  return {
    kilosProcesados: kilosProcesados,
    kilosPacking: kilosPacking,
    kilosCampo: kilosCampo,
    kilosVendidos: kilosVendidos,
    guiasEmitidas: Object.keys(guiasSet).length,
    ventasRegistradas: Object.keys(ventasSet).length,
    productos: Object.keys(productosSet).length,
    empresas: Object.keys(empresasSet).length,
    clientes: Object.keys(clientesSet).length
  };
}

function _dashboardAgruparEmpresa_(guias) {
  var mapa = {};

  guias.forEach(function(g) {
    var key = _dashboardNormalizar_(g.empresa);
    if (!mapa[key]) {
      mapa[key] = {
        empresa: g.empresa,
        kg: 0
      };
    }
    mapa[key].kg += g.kg;
  });

  return Object.keys(mapa)
    .map(function(k) { return mapa[k]; })
    .sort(function(a, b) { return b.kg - a.kg; });
}

function _dashboardAgruparFecha_(guias) {
  var mapa = {};

  guias.forEach(function(g) {
    if (!g.fechaKey) return;

    if (!mapa[g.fechaKey]) {
      mapa[g.fechaKey] = {
        fecha: g.fecha,
        fechaKey: g.fechaKey,
        kg: 0,
        packing: 0,
        campo: 0
      };
    }

    mapa[g.fechaKey].kg += g.kg;

    if (g.tipo === 'PACKING') mapa[g.fechaKey].packing += g.kg;
    if (g.tipo === 'CAMPO') mapa[g.fechaKey].campo += g.kg;
  });

  return Object.keys(mapa)
    .map(function(k) { return mapa[k]; })
    .sort(function(a, b) {
      return a.fechaKey.localeCompare(b.fechaKey);
    });
}

function _dashboardTopProductos_(guias, limite) {
  var mapa = {};

  guias.forEach(function(g) {
    var key = g.codigo;

    if (!mapa[key]) {
      mapa[key] = {
        codigo: g.codigo,
        descripcion: g.descripcion || g.codigo,
        kg: 0
      };
    }

    mapa[key].kg += g.kg;
  });

  return Object.keys(mapa)
    .map(function(k) { return mapa[k]; })
    .sort(function(a, b) { return b.kg - a.kg; })
    .slice(0, limite || 10);
}

function _dashboardTopClientes_(ventas, limite) {
  var mapa = {};

  ventas.forEach(function(v) {
    var key = _dashboardNormalizar_(v.cliente);
    if (!key) return;

    if (!mapa[key]) {
      mapa[key] = {
        cliente: v.cliente,
        kg: 0
      };
    }

    mapa[key].kg += v.kg;
  });

  return Object.keys(mapa)
    .map(function(k) { return mapa[k]; })
    .sort(function(a, b) { return b.kg - a.kg; })
    .slice(0, limite || 5);
}

function _dashboardRecientesGuias_(guias, limite) {
  var mapa = {};

  guias.forEach(function(g) {
    if (!mapa[g.correlativo]) {
      mapa[g.correlativo] = {
        correlativo: g.correlativo,
        empresa: g.empresa,
        cliente: g.cliente,
        fecha: g.fecha,
        kg: 0
      };
    }
    mapa[g.correlativo].kg += g.kg;
  });

  return Object.keys(mapa)
    .map(function(k) { return mapa[k]; })
    .sort(function(a, b) {
      return b.correlativo.localeCompare(a.correlativo);
    })
    .slice(0, limite || 6);
}

function _dashboardRecientesVentas_(ventas, limite) {
  var mapa = {};

  ventas.forEach(function(v) {
    if (!mapa[v.correlativo]) {
      mapa[v.correlativo] = {
        correlativo: v.correlativo,
        empresa: v.empresa,
        cliente: v.cliente,
        fecha: v.fecha,
        kg: 0
      };
    }
    mapa[v.correlativo].kg += v.kg;
  });

  return Object.keys(mapa)
    .map(function(k) { return mapa[k]; })
    .sort(function(a, b) {
      return b.fecha.localeCompare(a.fecha);
    })
    .slice(0, limite || 6);
}
function _dashboardLeerReporte_(empresa, fechaInicio, fechaFin, filtroVariedad, producto) {
  var resultado = {
    totalKg: 0,
    campo: 0,
    packing: 0,
    porFecha: [],
    porEmpresa: [],
    porVariedad: []
  };

  try {
    if (!fechaInicio || !fechaFin) return resultado;

    var respuesta = obtenerDatosReportePorFechas(fechaInicio, fechaFin, producto);

    if (!respuesta || !respuesta.exito) return resultado;

    var mapaFecha = {};
    var mapaEmpresa = {};
    var mapaVariedad = {};
    var cruceMapa = {};

    var filtro = _dashboardNormalizar_(empresa || 'TODAS');

    (respuesta.datos || []).forEach(function(item) {

      var empresaFila = String(item.empresa || '').trim();

      if (!_dashboardEmpresaCoincide_(empresaFila, filtro)) return;

      var kg = _dashboardNumero_(item.peso);

      if (kg <= 0) return;

      var fecha = String(item.fecha || '').trim();
      var variedad = String(item.variedad || '').trim() || 'SIN VARIEDAD';
      var presentacion = _dashboardNormalizar_(item.presentacion);

      var tipo = presentacion.indexOf('PACKING') >= 0
        ? 'PACKING'
        : 'CAMPO';

      var fechaKeyCruce = fecha;
      var partesFechaCruce = fecha.split('/');
      if (partesFechaCruce.length === 3) {
        fechaKeyCruce = partesFechaCruce[0] + '/' + partesFechaCruce[1] + '/' + partesFechaCruce[2];
      }

      var cruceKey = _dashboardNormalizar_(empresaFila) + '|' +
        fechaKeyCruce + '|' +
        _dashboardNormalizar_(variedad) + '|' +
        tipo;

      cruceMapa[cruceKey] = (cruceMapa[cruceKey] || 0) + kg;

      resultado.totalKg += kg;

      if (tipo === 'PACKING') {
        resultado.packing += kg;
      } else {
        resultado.campo += kg;
      }

      if (fecha) {
        if (!mapaFecha[fecha]) {
          mapaFecha[fecha] = {
            fecha: fecha,
            kg: 0,
            campo: 0,
            packing: 0
          };
        }

        mapaFecha[fecha].kg += kg;
        mapaFecha[fecha][tipo.toLowerCase()] += kg;
      }

      var empKey = _dashboardNormalizar_(empresaFila);

      if (!mapaEmpresa[empKey]) {
        mapaEmpresa[empKey] = {
          empresa: empresaFila,
          kg: 0
        };
      }

      mapaEmpresa[empKey].kg += kg;

      var varKey = _dashboardNormalizar_(variedad);

      if (!mapaVariedad[varKey]) {
        mapaVariedad[varKey] = {
          variedad: variedad,
          kg: 0
        };
      }

      mapaVariedad[varKey].kg += kg;
    });

    resultado.porFecha = Object.keys(mapaFecha)
      .map(function(k) {
        return mapaFecha[k];
      });

    resultado.porFecha.sort(function(a, b) {
      return a.fecha.localeCompare(b.fecha);
    });

    resultado.porEmpresa = Object.keys(mapaEmpresa)
      .map(function(k) {
        return mapaEmpresa[k];
      });

    resultado.porEmpresa.sort(function(a, b) {
      return b.kg - a.kg;
    });

    resultado.porVariedad = Object.keys(mapaVariedad)
      .map(function(k) {
        return mapaVariedad[k];
      });

    resultado.porVariedad.sort(function(a, b) {
      return b.kg - a.kg;
    });

    resultado.cruceMapa = cruceMapa;
    return resultado;

  } catch (e) {
    return resultado;
  }
}


function _dashboardLeerStockSAP_(empresa, fechaInicio, fechaFin, filtroVariedad, mapaSispacking, producto) {
  var resultado = {
    total: 0,
    pendienteVenta: 0,
    diferenciasTotal: 0,
    diferencias: [],
    porEmpresa: [],
    porProducto: [],
    porVariedad: [],
    controlStock: [],
    variedadesDisponibles: []
  };

  try {
    var filtro = _dashboardNormalizar_(empresa || 'TODAS');
    var esUva = String(producto || 'ARANDANOS').toUpperCase() === 'UVA';
    var empresas = filtro === 'TODAS'
      ? (esUva ? ['AGA', 'ARENUVA'] : ['AGA', 'LARAMA', 'ARENUVA'])
      : [filtro];

    mapaSispacking = mapaSispacking || {};

    var inicio = _dashboardFecha_(fechaInicio, false);
    var fin = _dashboardFecha_(fechaFin, true);

    var ss = SpreadsheetApp.openById(_dashboardSpreadsheetId_());
    var listaProductosBD = [];
    var sheetProductos = ss.getSheetByName('BD-PRODUCTOS');

    if (sheetProductos && sheetProductos.getLastRow() > 1) {
      var datosProd = sheetProductos.getDataRange().getValues();
      for (var p = 1; p < datosProd.length; p++) {
        var codCelda = datosProd[p][0] ? datosProd[p][0].toString().trim().toUpperCase() : '';
        var descCelda = datosProd[p][1] ? datosProd[p][1].toString().trim().toUpperCase() : '';
        if (codCelda && descCelda) {
          listaProductosBD.push({ codigo: codCelda, descripcion: descCelda });
        }
      }
    }

    var mapaEmpresa = {};
    var mapaProducto = {};
    var mapaVariedad = {};
    var mapaDiferencias = {};
    var mapaSap = {};
    var mapaStockFinal = {};

    // SAP se lee una sola vez por empresa. No se vuelven a abrir los archivos SISPACKING.
    empresas.forEach(function(nombreEmpresa) {
      var nombrePestanaSap = 'SAP ' + nombreEmpresa;
      var sheetSap = ss.getSheetByName(nombrePestanaSap);
      if (!sheetSap || sheetSap.getLastRow() <= 1) return;

      var datosSap = sheetSap.getDataRange().getValues();

      var variedadActual = '';
      var tipoGrupoActual = '';
      var ultimoValorColJ = null;
      var llaveStockActual = '';

      for (var r = 1; r < datosSap.length; r++) {
        var colA = datosSap[r][0] ? datosSap[r][0].toString().trim() : '';
        var colB = datosSap[r][1] ? datosSap[r][1].toString().trim().toUpperCase() : '';

        if (colA !== '' || colB !== '') {
          if (colB !== '' && colB !== 'DESCRIPCIÓN' && colB !== 'DESCRIPCION') {

            if (llaveStockActual !== '' && ultimoValorColJ !== null) {
              mapaStockFinal[nombreEmpresa + '|' + llaveStockActual] = ultimoValorColJ;
            }

            tipoGrupoActual = colB.includes('PACKING') ? 'PACKING' : 'CAMPO';

            if (esUva) {
              var limpiarUva = colB
                .replace(/DESCARTE\s+(CAMPO|PACKING)/g, '')
                .replace(/ARANDANO(S)?\s+FRESCO(S)?/g, '')
                .replace(/UVA(S)?\s+FRESCA(S)?/g, '')
                .replace(/CAMPO/g, '')
                .replace(/PACKING/g, '')
                .replace(/[-]/g, ' ')
                .replace(/\s+/g, ' ')
                .trim();
              if (limpiarUva) variedadActual = limpiarUva;
            } else if (colB.includes('ROSITA')) {
              variedadActual = 'ROSITA';
            } else if (colB.includes('RAYMI')) {
              variedadActual = 'RAYMI BLU';
            } else if (colB.includes('DINA')) {
              variedadActual = 'OZBLU® DINA';
            } else if (colB.includes('CAROLINA')) {
              variedadActual = 'OZBLU® CAROLINA';
            } else if (colB.includes('JULIETA')) {
              variedadActual = 'OZBLU® JULIETA';
            } else if (colB.includes('ANDREA')) {
              variedadActual = 'OZBLU® ANDREA';
            } else if (colB.includes('MÁGICA') || colB.includes('MAGICA')) {
              variedadActual = 'OZBLU® MÁGICA';
            } else if (colB.includes('OLIVIA')) {
              variedadActual = 'OZBLU® OLIVIA';
            } else {
              var limpiar = colB
                .replace('ARÁNDANO FRESCO', '')
                .replace('- DESCARTE CAMPO', '')
                .replace('- DESCARTE PACKING', '')
                .replace('CAMPO', '')
                .replace('PACKING', '')
                .trim();
              if (limpiar) variedadActual = limpiar;
            }

            // No permitir que un bloque SAP de la otra familia alimente este perfil.
            var variedadPerfilNorm = typeof _normalizarClaveVariedadProducto_ === 'function'
              ? _normalizarClaveVariedadProducto_(variedadActual)
              : _dashboardNormalizar_(variedadActual).replace(/[^A-Z0-9]/g, '');
            var familiaBloqueSAP = '';
            for (var fb = 0; fb < listaProductosBD.length; fb++) {
              var descPerfil = listaProductosBD[fb].descripcion;
              var descPerfilNorm = typeof _normalizarClaveVariedadProducto_ === 'function'
                ? _normalizarClaveVariedadProducto_(descPerfil)
                : _dashboardNormalizar_(descPerfil).replace(/[^A-Z0-9]/g, '');
              if (descPerfilNorm === variedadPerfilNorm) {
                familiaBloqueSAP = typeof _familiaProductoDesdeDescripcion_ === 'function'
                  ? _familiaProductoDesdeDescripcion_(descPerfil)
                  : (/\bUVA(S)?\s+FRESCA(S)?\b/.test(descPerfil) ? 'UVA' : (/\bARANDANO(S)?\b/.test(descPerfil) ? 'ARANDANOS' : ''));
                if (familiaBloqueSAP) break;
              }
            }
            if (familiaBloqueSAP !== (esUva ? 'UVA' : 'ARANDANOS')) {
              variedadActual = '';
              llaveStockActual = '';
              tipoGrupoActual = '';
            } else {
              llaveStockActual = variedadActual + '|' + tipoGrupoActual;
            }
            ultimoValorColJ = null;
          }
        }

        var valJ = parseFloat(datosSap[r][9]);
        if (!isNaN(valJ)) ultimoValorColJ = valJ;

        // Movimientos SAP: solo los que caen dentro del rango seleccionado.
        var fRawS = datosSap[r][3];
        var docCodigo = datosSap[r][4]
          ? datosSap[r][4].toString().toUpperCase().replace(/\s+/g, '')
          : '';
        var cantidadS = parseFloat(datosSap[r][6]) || 0;

        if ((docCodigo.indexOf('IM') === 0 || docCodigo.indexOf('EM') === 0) &&
            cantidadS !== 0 && variedadActual !== '') {

          var fObjS = _dashboardFechaDesdeCelda_(fRawS);
          if (fObjS && _dashboardEnRango_(fObjS, inicio, fin)) {
            var fTxtS = Utilities.formatDate(
              fObjS,
              Session.getScriptTimeZone(),
              'dd/MM/yyyy'
            );

            var llaveS = nombreEmpresa + '|' + fTxtS + '|' +
              _dashboardNormalizar_(variedadActual) + '|' + tipoGrupoActual;

            mapaSap[llaveS] = (mapaSap[llaveS] || 0) + cantidadS;
          }
        }
      }

      if (llaveStockActual !== '' && ultimoValorColJ !== null) {
        mapaStockFinal[nombreEmpresa + '|' + llaveStockActual] = ultimoValorColJ;
      }
    });

    // STOCK ACTUAL: último valor disponible del bloque SAP.
    Object.keys(mapaStockFinal).forEach(function(key) {
      var partes = key.split('|');
      var nombreEmpresa = partes.shift();
      var variedad = partes[0] || 'SIN VARIEDAD';
      var tipo = partes[1] || 'OTRO';

      var variedadNormalizada = _dashboardNormalizar_(variedad);
      if (
        filtroVariedad &&
        filtroVariedad !== 'TODAS' &&
        variedadNormalizada !== filtroVariedad &&
        variedadNormalizada.indexOf(filtroVariedad) < 0 &&
        filtroVariedad.indexOf(variedadNormalizada) < 0
      ) return;

      var kg = _dashboardNumero_(mapaStockFinal[key]);
      var empKey = nombreEmpresa;
      var prodKey = nombreEmpresa + '|' + variedadNormalizada + '|' + tipo;

      if (!mapaEmpresa[empKey]) {
        mapaEmpresa[empKey] = { empresa: nombreEmpresa, kg: 0 };
      }
      mapaEmpresa[empKey].kg += kg;

      if (!mapaProducto[prodKey]) {
        mapaProducto[prodKey] = {
          empresa: nombreEmpresa,
          variedad: variedad,
          tipo: tipo === 'CAMPO' ? 'GRANEL DESCARTE CAMPO' : 'DESCARTE PACKING',
          kg: 0
        };
      }
      mapaProducto[prodKey].kg += kg;

      if (!mapaVariedad[variedadNormalizada]) {
        mapaVariedad[variedadNormalizada] = {
          variedad: variedad,
          stock: 0,
          kg: 0
        };
      }
      mapaVariedad[variedadNormalizada].stock += kg;
      mapaVariedad[variedadNormalizada].kg += kg;
    });

    // DIFERENCIAS: SAP del rango seleccionado vs SISPACKING del mismo rango.
    var clavesCruce = {};
    Object.keys(mapaSap).forEach(function(key) { clavesCruce[key] = true; });
    Object.keys(mapaSispacking).forEach(function(key) { clavesCruce[key] = true; });

    Object.keys(clavesCruce).forEach(function(key) {
      var totalSap = mapaSap[key] || 0;
      var totalSispacking = mapaSispacking[key] || 0;
      var diferencia = totalSap - totalSispacking;

      if (Math.abs(diferencia) <= 0.01) return;

      var partes = key.split('|');
      var nombreEmpresa = partes[0];
      var fecha = partes[1];
      var variedadNormalizada = partes[2];
      var tipo = partes[3];

      var variedad = variedadNormalizada;
      var prodDescripcion = variedadNormalizada;

      for (var b = 0; b < listaProductosBD.length; b++) {
        var desc = listaProductosBD[b].descripcion;
        if (
          desc.indexOf(variedadNormalizada) >= 0 ||
          variedadNormalizada.indexOf(desc) >= 0
        ) {
          prodDescripcion = desc;
          break;
        }
      }

      var diffKey = nombreEmpresa + '|' + variedadNormalizada + '|' + tipo;

      if (!mapaDiferencias[diffKey]) {
        mapaDiferencias[diffKey] = {
          empresa: nombreEmpresa,
          variedad: variedad || prodDescripcion,
          diferencia: 0,
          fecha: fecha
        };
      }

      mapaDiferencias[diffKey].diferencia += diferencia;
    });

    resultado.porEmpresa = Object.keys(mapaEmpresa)
      .map(function(k) { return mapaEmpresa[k]; })
      .sort(function(a, b) { return b.kg - a.kg; });

    resultado.porProducto = Object.keys(mapaProducto)
      .map(function(k) { return mapaProducto[k]; })
      .sort(function(a, b) { return b.kg - a.kg; });

    resultado.porVariedad = Object.keys(mapaVariedad)
      .map(function(k) { return mapaVariedad[k]; })
      .sort(function(a, b) { return b.stock - a.stock; });

    resultado.total = resultado.porEmpresa.reduce(function(suma, item) {
      return suma + item.kg;
    }, 0);

    resultado.diferencias = Object.keys(mapaDiferencias)
      .map(function(k) { return mapaDiferencias[k]; })
      .filter(function(item) { return Math.abs(item.diferencia) > 0.01; })
      .sort(function(a, b) {
        return Math.abs(b.diferencia) - Math.abs(a.diferencia);
      });

    resultado.diferenciasTotal = resultado.diferencias.reduce(function(suma, item) {
      return suma + Math.abs(item.diferencia);
    }, 0);

    resultado.controlStock = resultado.porVariedad.map(function(item) {
      return {
        variedad: item.variedad,
        stock: item.stock,
        vendido: 0,
        pendiente: item.stock,
        estado: item.stock > 0 ? 'PENDIENTE DE VENTA' : 'SIN STOCK'
      };
    });

    resultado.variedadesDisponibles = resultado.porVariedad
      .map(function(item) { return item.variedad; })
      .sort(function(a, b) { return a.localeCompare(b); });

    return resultado;

  } catch (e) {
    return resultado;
  }
}
