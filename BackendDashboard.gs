function getDashboardData(empresa, fechaInicio, fechaFin) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Configuración de hojas
    var sheetCampo = ss.getSheetByName("DESCARTE CAMPO");
    var sheetPacking = ss.getSheetByName("DESCARTE PACKING");
    
    var dataCampo = sheetCampo ? sheetCampo.getDataRange().getValues() : [];
    var dataPacking = sheetPacking ? sheetPacking.getDataRange().getValues() : [];
    
    var totalCampo = 0;
    var totalPacking = 0;
    
    var variedadesMap = {}; // { "VARIEDAD": { campo: 0, packing: 0 } }
    var fechasMap = {};     // { "YYYY-MM-DD": { campo: 0, packing: 0 } }
    
    // Procesar Descarte Campo
    if (dataCampo.length > 1) {
      var headers = dataCampo[0].map(function(h) { return h.toString().toUpperCase().trim(); });
      var idxFecha = headers.indexOf("FECHA");
      var idxEmpresa = headers.indexOf("EMPRESA");
      var idxVariedad = headers.indexOf("VARIEDAD");
      var idxKilos = headers.indexOf("KILOS");
      if (idxKilos === -1) idxKilos = headers.indexOf("KG");
      
      for (var i = 1; i < dataCampo.length; i++) {
        var row = dataCampo[i];
        var rowEmpresa = idxEmpresa !== -1 ? row[idxEmpresa] : "";
        var rowFecha = idxFecha !== -1 ? row[idxFecha] : null;
        var rowVariedad = idxVariedad !== -1 ? row[idxVariedad] : "OTRAS";
        var rowKilos = idxKilos !== -1 ? parseFloat(row[idxKilos]) || 0 : 0;
        
        if (cumpleFiltro(rowEmpresa, rowFecha, empresa, fechaInicio, fechaFin)) {
          totalCampo += rowKilos;
          
          // Variedad
          if (!variedadesMap[rowVariedad]) variedadesMap[rowVariedad] = { campo: 0, packing: 0 };
          variedadesMap[rowVariedad].campo += rowKilos;
          
          // Fecha
          var strFecha = formatearFechaKey(rowFecha);
          if (strFecha) {
            if (!fechasMap[strFecha]) fechasMap[strFecha] = { campo: 0, packing: 0 };
            fechasMap[strFecha].campo += rowKilos;
          }
        }
      }
    }
    
    // Procesar Descarte Packing
    if (dataPacking.length > 1) {
      var headersP = dataPacking[0].map(function(h) { return h.toString().toUpperCase().trim(); });
      var idxFechaP = headersP.indexOf("FECHA");
      var idxEmpresaP = headersP.indexOf("EMPRESA");
      var idxVariedadP = headersP.indexOf("VARIEDAD");
      var idxKilosP = headersP.indexOf("KILOS");
      if (idxKilosP === -1) idxKilosP = headersP.indexOf("KG");
      
      for (var j = 1; j < dataPacking.length; j++) {
        var rowP = dataPacking[j];
        var rowEmpresaP = idxEmpresaP !== -1 ? rowP[idxEmpresaP] : "";
        var rowFechaP = idxFechaP !== -1 ? rowP[idxFechaP] : null;
        var rowVariedadP = idxVariedadP !== -1 ? rowP[idxVariedadP] : "OTRAS";
        var rowKilosP = idxKilosP !== -1 ? parseFloat(rowP[idxKilosP]) || 0 : 0;
        
        if (cumpleFiltro(rowEmpresaP, rowFechaP, empresa, fechaInicio, fechaFin)) {
          totalPacking += rowKilosP;
          
          // Variedad
          if (!variedadesMap[rowVariedadP]) variedadesMap[rowVariedadP] = { campo: 0, packing: 0 };
          variedadesMap[rowVariedadP].packing += rowKilosP;
          
          // Fecha
          var strFechaP = formatearFechaKey(rowFechaP);
          if (strFechaP) {
            if (!fechasMap[strFechaP]) fechasMap[strFechaP] = { campo: 0, packing: 0 };
            fechasMap[strFechaP].packing += rowKilosP;
          }
        }
      }
    }
    
    return {
      success: true,
      kpis: {
        descarteCampo: totalCampo,
        descartePacking: totalPacking,
        totalConsolidado: totalCampo + totalPacking
      },
      variedades: variedadesMap,
      fechas: fechasMap
    };
    
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function cumpleFiltro(rowEmpresa, rowFecha, filtroEmpresa, fechaInicio, fechaFin) {
  if (filtroEmpresa && filtroEmpresa !== "TODAS") {
    if (rowEmpresa.toString().toUpperCase().trim() !== filtroEmpresa.toUpperCase().trim()) {
      return false;
    }
  }
  
  if (fechaInicio || fechaFin) {
    if (!rowFecha) return false;
    var dFecha = new Date(rowFecha);
    if (isNaN(dFecha.getTime())) return false;
    
    if (fechaInicio) {
      var dInicio = new Date(fechaInicio);
      if (dFecha < dInicio) return false;
    }
    if (fechaFin) {
      var dFin = new Date(fechaFin);
      dFin.setHours(23, 59, 59, 999);
      if (dFecha > dFin) return false;
    }
  }
  return true;
}

function formatearFechaKey(fecha) {
  if (!fecha) return "";
  var d = new Date(fecha);
  if (isNaN(d.getTime())) return "";
  var year = d.getFullYear();
  var month = ("0" + (d.getMonth() + 1)).slice(-2);
  var day = ("0" + d.getDate()).slice(-2);
  return year + "-" + month + "-" + day;
}