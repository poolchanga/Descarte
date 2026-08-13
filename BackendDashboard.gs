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

function getDashboardData(empresa, fechaInicio, fechaFin) {
  try {
    var ss = SpreadsheetApp.openById(_dashboardSpreadsheetId_());
    var sheetGuias = ss.getSheetByName('HISTORIAL_GUIAS');
    var sheetVentas = ss.getSheetByName('VENTAS');

    var filtroEmpresa = _dashboardNormalizar_(empresa || 'TODAS');
    var inicio = _dashboardFecha_(fechaInicio, false);
    var fin = _dashboardFecha_(fechaFin, true);

    var guias = _dashboardLeerGuias_(sheetGuias, filtroEmpresa, inicio, fin);
    var ventas = _dashboardLeerVentas_(sheetVentas, filtroEmpresa, inicio, fin);

    var reporte = _dashboardLeerReporte_(empresa || 'TODAS', fechaInicio, fechaFin);
    var stock = _dashboardLeerStockSAP_(empresa || 'TODAS');

    var kpis = _dashboardKpis_(guias, ventas);

    kpis.kilosDescarte = reporte.totalKg;
    kpis.kilosCampo = reporte.campo;
    kpis.kilosPacking = reporte.packing;
    kpis.stockActual = stock.total;

    return {
      exito: true,

      filtros: {
        empresa: filtroEmpresa || 'TODAS',
        fechaInicio: fechaInicio || '',
        fechaFin: fechaFin || ''
      },

      kpis: kpis,

      porEmpresa: _dashboardAgruparEmpresa_(guias),
      porFecha: _dashboardAgruparFecha_(guias),
      porProducto: _dashboardTopProductos_(guias, 10),
      porCliente: _dashboardTopClientes_(ventas, 5),

      porFechaReporte: reporte.porFecha,
      porEmpresaReporte: reporte.porEmpresa,
      porVariedadReporte: reporte.porVariedad,

      stockPorEmpresa: stock.porEmpresa,
      stockPorProducto: stock.porProducto,

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

function _dashboardLeerVentas_(sheet, filtroEmpresa, inicio, fin) {
  var resultado = [];
  if (!sheet || sheet.getLastRow() <= 1) return resultado;

  // VENTAS:
  // A correlativo | B fecha | C RUC | D cliente | E código
  // F variedad | G KG | H empresa | I link origen
  var datos = sheet.getRange(
    2, 1, sheet.getLastRow() - 1, 9
  ).getValues();

  for (var i = 0; i < datos.length; i++) {
    var fila = datos[i];

    var correlativo = String(fila[0] || '').trim();
    var fecha = fila[1];
    var ruc = String(fila[2] || '').trim();
    var cliente = String(fila[3] || '').trim();
    var codigo = String(fila[4] || '').trim();
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
function _dashboardLeerReporte_(empresa, fechaInicio, fechaFin) {
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

    var respuesta = obtenerDatosReportePorFechas(fechaInicio, fechaFin);

    if (!respuesta || !respuesta.exito) return resultado;

    var mapaFecha = {};
    var mapaEmpresa = {};
    var mapaVariedad = {};

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

    return resultado;

  } catch (e) {
    return resultado;
  }
}


function _dashboardLeerStockSAP_(empresa) {

  var resultado = {
    total: 0,
    porEmpresa: [],
    porProducto: []
  };

  try {

    var ss = SpreadsheetApp.openById(
      _dashboardSpreadsheetId_()
    );

    var filtro = _dashboardNormalizar_(
      empresa || 'TODAS'
    );

    var empresas = filtro === 'TODAS'
      ? ['AGA', 'LARAMA', 'ARENUVA']
      : [filtro];

    var mapaEmpresa = {};
    var mapaProducto = {};

    empresas.forEach(function(nombreEmpresa) {

      var sheet = ss.getSheetByName(
        'SAP ' + nombreEmpresa
      );

      if (!sheet) return;

      var datos = sheet.getDataRange().getValues();

      var variedadActual = '';
      var tipoActual = '';
      var ultimoStock = null;
      var llaveActual = '';

      for (var r = 1; r < datos.length; r++) {

        var colA = datos[r][0] != null
          ? String(datos[r][0]).trim()
          : '';

        var colB = datos[r][1] != null
          ? String(datos[r][1]).trim().toUpperCase()
          : '';

        if (colA !== '' || colB !== '') {

          if (
            colB !== '' &&
            colB !== 'DESCRIPCIÓN' &&
            colB !== 'DESCRIPCION'
          ) {

            if (
              llaveActual !== '' &&
              ultimoStock !== null
            ) {

              _dashboardAcumularStock_(
                mapaEmpresa,
                mapaProducto,
                nombreEmpresa,
                variedadActual,
                tipoActual,
                ultimoStock
              );
            }

            tipoActual =
              colB.indexOf('PACKING') >= 0
                ? 'PACKING'
                : 'CAMPO';

            if (colB.indexOf('ROSITA') >= 0) {
              variedadActual = 'ROSITA';

            } else if (colB.indexOf('RAYMI') >= 0) {
              variedadActual = 'RAYMI BLU';

            } else if (colB.indexOf('DINA') >= 0) {
              variedadActual = 'OZBLU® DINA';

            } else if (colB.indexOf('CAROLINA') >= 0) {
              variedadActual = 'OZBLU® CAROLINA';

            } else if (colB.indexOf('JULIETA') >= 0) {
              variedadActual = 'OZBLU® JULIETA';

            } else if (colB.indexOf('ANDREA') >= 0) {
              variedadActual = 'OZBLU® ANDREA';

            } else if (
              colB.indexOf('MÁGICA') >= 0 ||
              colB.indexOf('MAGICA') >= 0
            ) {
              variedadActual = 'OZBLU® MÁGICA';

            } else if (colB.indexOf('OLIVIA') >= 0) {
              variedadActual = 'OZBLU® OLIVIA';

            } else {

              variedadActual = colB
                .replace('ARÁNDANO FRESCO', '')
                .replace('- DESCARTE CAMPO', '')
                .replace('- DESCARTE PACKING', '')
                .replace('CAMPO', '')
                .replace('PACKING', '')
                .trim();
            }

            llaveActual =
              variedadActual +
              '|' +
              tipoActual;

            ultimoStock = null;
          }
        }

        var valorJ = parseFloat(
          datos[r][9]
        );

        if (!isNaN(valorJ)) {
          ultimoStock = valorJ;
        }
      }

      if (
        llaveActual !== '' &&
        ultimoStock !== null
      ) {

        _dashboardAcumularStock_(
          mapaEmpresa,
          mapaProducto,
          nombreEmpresa,
          variedadActual,
          tipoActual,
          ultimoStock
        );
      }
    });

    resultado.porEmpresa =
      Object.keys(mapaEmpresa)
        .map(function(k) {
          return mapaEmpresa[k];
        });

    resultado.porEmpresa.sort(function(a, b) {
      return b.kg - a.kg;
    });

    resultado.porProducto =
      Object.keys(mapaProducto)
        .map(function(k) {
          return mapaProducto[k];
        });

    resultado.porProducto.sort(function(a, b) {
      return b.kg - a.kg;
    });

    resultado.total =
      resultado.porEmpresa.reduce(
        function(suma, item) {
          return suma + item.kg;
        },
        0
      );

    return resultado;

  } catch (e) {

    return resultado;
  }
}


function _dashboardAcumularStock_(
  mapaEmpresa,
  mapaProducto,
  empresa,
  variedad,
  tipo,
  stock
) {

  var kg = _dashboardNumero_(stock);

  var empKey = empresa;

  var prodKey =
    empresa +
    '|' +
    _dashboardNormalizar_(variedad) +
    '|' +
    tipo;

  if (!mapaEmpresa[empKey]) {

    mapaEmpresa[empKey] = {
      empresa: empresa,
      kg: 0
    };
  }

  mapaEmpresa[empKey].kg += kg;

  if (!mapaProducto[prodKey]) {

    mapaProducto[prodKey] = {
      empresa: empresa,
      variedad: variedad,
      tipo: tipo,
      kg: 0
    };
  }

  mapaProducto[prodKey].kg += kg;
}
