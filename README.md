# SCRIPT-1
SCRIT VALIDO PARA MASIVO Y ACTUALIZAR SEGUN PIVOTE



/**
 * SUITE ROCKWELL AUTOMATION - GOOGLE SHEETS
 * Consulta el estado de Ciclo de Vida de los PNs.
 */

// 1. CREA EL MENÚ EN GOOGLE SHEETS AL ABRIR EL ARCHIVO
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🤖 Suite Rockwell')
      .addItem('🔍 Consultar Ciclo de Vida', 'consultarCicloDeVida')
      .addToUi();
}

function consultarCicloDeVida() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  if (data.length < 2) {
    SpreadsheetApp.getUi().alert("No hay suficientes datos en la hoja.");
    return;
  }

  var header = data[0];
  var pnColIndex = -1;
  
  // 🚀 MEJORA 1: BUSCAR LA COLUMNA PN DINÁMICAMENTE
  for (var i = 0; i < header.length; i++) {
    var val = String(header[i]).toUpperCase().trim();
    if (val === "PN" || val === "P/N" || val === "PART NUMBER") {
      pnColIndex = i;
      break;
    }
  }

  if (pnColIndex === -1) {
    SpreadsheetApp.getUi().alert("⚠️ No se encontró la columna 'PN', 'P/N' o 'PART NUMBER' en la primera fila.");
    return;
  }

  var statusColIndex = pnColIndex + 1;
  var statusHeader = header.length > statusColIndex ? String(header[statusColIndex]).toUpperCase().trim() : "";

  // 🚀 MEJORA 2: CREAR COLUMNA DE ESTADO SI NO EXISTE
  if (statusHeader !== "ESTADO ROCKWELL") {
    sheet.insertColumnAfter(pnColIndex + 1);
    statusColIndex = pnColIndex + 1;
    var headerCell = sheet.getRange(1, statusColIndex + 1);
    headerCell.setValue("ESTADO ROCKWELL").setFontWeight("bold").setBackground("#33475b").setFontColor("white");
  }

  var numRows = sheet.getLastRow();
  var procesados = 0;

  // Empezar desde la fila 2 (omitiendo encabezados)
  for (var r = 2; r <= numRows; r++) {
    var pn = sheet.getRange(r, pnColIndex + 1).getValue().toString().trim();
    var cellEstado = sheet.getRange(r, statusColIndex + 1);
    var estadoActual = cellEstado.getValue().toString().trim();

    // 🚀 MEJORA 4: SMART SKIP (Salta vacíos, GEN- y los que ya tienen estado)
    if (pn === "" || pn.toUpperCase().indexOf("GEN-") === 0 || estadoActual !== "") {
      continue; 
    }

    // Indicar que está procesando
    cellEstado.setValue("Consultando... ⏳").setFontColor("#516f90").setBackground("#f5f8fa");
    SpreadsheetApp.flush(); // Fuerza a Google Sheets a mostrar el cambio en vivo

    // Llamar a la web de Rockwell
    var estado = fetchRockwellStatus(pn);
    cellEstado.setValue(estado);
    
    // Colores de estado
    var eLow = estado.toLowerCase();
    if (eLow.indexOf("active mature") > -1) {
      cellEstado.setBackground("#fff9e6").setFontColor("#e59a10"); // Amarillo
    } else if (eLow === "active") {
      cellEstado.setBackground("#e5fdf4").setFontColor("#00bda5"); // Verde
    } else if (eLow.indexOf("discontinued") > -1 || eLow.indexOf("end of life") > -1) {
      cellEstado.setBackground("#ffe5e5").setFontColor("#ff3333"); // Rojo
    } else {
      cellEstado.setBackground("white").setFontColor("#33475b"); // Default
    }

    SpreadsheetApp.flush(); // Actualiza la pantalla de nuevo
    procesados++;
    
    // 🚀 MEJORA 3: PAUSA ALEATORIA ANTI-BLOQUEO (Entre 2 y 4 segundos)
    var pausa = Math.floor(Math.random() * 2000) + 2000;
    Utilities.sleep(pausa);
  }
  
  if(procesados > 0) {
    SpreadsheetApp.getUi().alert("✅ Consulta finalizada. Se procesaron " + procesados + " equipos.");
  } else {
    SpreadsheetApp.getUi().alert("👍 Todo al día. No había PNs nuevos por consultar.");
  }
}

function fetchRockwellStatus(pn) {
  // Usamos el endpoint directo de detalles de producto (el más confiable)
  var url = "https://www.rockwellautomation.com/en-us/products/details." + encodeURIComponent(pn) + ".html";
  
  var options = {
    "method": "get",
    "muteHttpExceptions": true, // Evita que el script colapse si RA nos bloquea
    "headers": {
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
      "Accept-Language": "en-US,en;q=0.9"
    }
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var rc = response.getResponseCode();
    
    if (rc === 200) {
      var html = response.getContentText();
      // Buscamos el HTML específico usando una expresión regular
      var regex = /<span[^>]*class=["'][^"']*ra-product-new__product-detail-options-status-value[^"']*["'][^>]*>([^<]+)<\/span>/i;
      var match = html.match(regex);
      
      if (match && match[1]) {
        return match[1].trim();
      } else {
        return "No Encontrado";
      }
    } else if (rc === 403 || rc === 401) {
      return "Bloqueado por RA (Bot)";
    } else {
      return "Error Servidor (" + rc + ")";
    }
  } catch (e) {
    return "Error de Conexión";
  }
}
