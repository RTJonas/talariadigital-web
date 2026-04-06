// =========================================================
// TALARIA DIGITAL — Apps Script
// Archivo: talaria-ledger.gs
// Asociado a: Talaria Ledger 2026 (Google Sheets)
// Última actualización: 2026-04-06
// =========================================================

// =========================================================
// CARGAR GASTO
// Registra un gasto manual asociado a una creadora.
// El gasto entra en negativo en Bruto. Honorarios = 20% del ABS.
// Genera saldo deudor en la siguiente liquidación si el balance queda negativo.
// =========================================================
function cargarGasto() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Transacciones");

  if (!sheet) {
    ui.alert("❌ Error: No se encontró la pestaña 'Transacciones'.");
    return;
  }

  var idResponse = ui.prompt("Cargar Gasto 💸", "1. Ingresá el ID de la Creadora (ej: 21):", ui.ButtonSet.OK_CANCEL);
  if (idResponse.getSelectedButton() !== ui.Button.OK || idResponse.getResponseText() === "") return;
  var idCreadora = idResponse.getResponseText().trim();

  var montoResponse = ui.prompt("Cargar Gasto 💸", "2. Ingresá el monto del gasto total en EUROS (ej: 120):", ui.ButtonSet.OK_CANCEL);
  if (montoResponse.getSelectedButton() !== ui.Button.OK || montoResponse.getResponseText() === "") return;

  var montoText = montoResponse.getResponseText().replace(",", ".").replace("€", "").trim();
  var monto = parseFloat(montoText);
  if (isNaN(monto) || monto <= 0) {
    ui.alert("❌ Error: Ingresá un número válido mayor a 0.");
    return;
  }

  var conceptoResponse = ui.prompt("Cargar Gasto 💸", "3. ¿Qué tipo de gasto es? (ej: Publicidad, Diseño):", ui.ButtonSet.OK_CANCEL);
  if (conceptoResponse.getSelectedButton() !== ui.Button.OK) return;
  var concepto = conceptoResponse.getResponseText().trim().toUpperCase();
  if (concepto === "") concepto = "VARIOS";

  var refResponse = ui.prompt("Cargar Gasto 💸", "4. Ingresá el ID de la transacción, recibo o factura:", ui.ButtonSet.OK_CANCEL);
  if (refResponse.getSelectedButton() !== ui.Button.OK) return;
  var referencia = refResponse.getResponseText().trim();

  var fechaHoy = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
  var bruto = -monto;
  var comision = 0;

  var nuevaFila = [
    fechaHoy,              // Col A: Fecha y Hora
    "GASTO - " + concepto, // Col B: Origen / Tipo
    referencia,            // Col C: ID Transaccion
    "",                    // Col D: Email Cliente Final
    idCreadora,            // Col E: ID Proveedor
    bruto,                 // Col F: Bruto
    comision,              // Col G: Comision
    "",                    // Col H: Neto (fórmula)
    "",                    // Col I: Honorarios (fórmula)
    "",                    // Col J: A Liquidar (fórmula)
    "",                    // Col K: Fecha de Liquidacion
    "Pendiente"            // Col L: Estado Liquidacion
  ];

  sheet.appendRow(nuevaFila);
  var row = sheet.getLastRow();

  sheet.getRange(row, 8).setFormula("=F" + row + "-G" + row);
  sheet.getRange(row, 9).setFormula("=ABS(H" + row + ")*0.2");
  sheet.getRange(row, 10).setFormula("=H" + row + "-I" + row);

  if (row > 2) {
    sheet.getRange(row - 1, 11).copyTo(sheet.getRange(row, 11));
  }

  var misHonorarios = monto * 0.2;
  var aLiquidar = bruto - misHonorarios;

  ui.alert("✅ Gasto registrado con éxito:\n\n" +
           "👤 ID: " + idCreadora + "\n" +
           "📝 Concepto: " + concepto + "\n" +
           "💶 Costo Gasto: €" + monto.toFixed(2) + "\n" +
           "📈 Tus Honorarios (20%): €" + misHonorarios.toFixed(2) + "\n" +
           "💥 Total a descontar: €" + Math.abs(aLiquidar).toFixed(2));
}

// =========================================================
// CARGAR TIPO DE CAMBIO
// Pide el precio EUR/USDT de Kraken y calcula los USDT
// exactos dividiendo el monto en EUR por la tasa.
// Guarda el historial de tasas en cols K-L de "Semanal - Liquidacion".
// =========================================================
function cargarTipoCambio() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var hojaTransacciones = ss.getSheetByName("Transacciones");
  var hojaLiquidacion = ss.getSheetByName("Semanal - Liquidacion");

  if (!hojaTransacciones || !hojaLiquidacion) {
    ui.alert("❌ Error: Revisá que las pestañas se llamen exactamente 'Transacciones' y 'Semanal - Liquidacion'.");
    return;
  }

  var rateResponse = ui.prompt("Cotización Kraken 💱", "Ingresá el Precio en EUR por cada USDT de hoy\n(El número que te da Kraken, ej: 0.8626):", ui.ButtonSet.OK_CANCEL);
  if (rateResponse.getSelectedButton() !== ui.Button.OK || rateResponse.getResponseText() === "") return;

  var rateText = rateResponse.getResponseText().replace(",", ".").trim();
  var tipoDeCambio = parseFloat(rateText);

  if (isNaN(tipoDeCambio) || tipoDeCambio <= 0) {
    ui.alert("❌ Error: Ingresá un número válido para la tasa de cambio.");
    return;
  }

  var colK = hojaLiquidacion.getRange("K:K").getValues();
  var ultimaFilaHistorial = 1;
  for (var j = 0; j < colK.length; j++) {
    if (colK[j][0] !== "") {
      ultimaFilaHistorial = j + 1;
    }
  }

  var fechaHoy = new Date();
  hojaLiquidacion.getRange(ultimaFilaHistorial + 1, 11).setValue(fechaHoy);
  hojaLiquidacion.getRange(ultimaFilaHistorial + 1, 12).setValue(tipoDeCambio);

  var datos = hojaTransacciones.getDataRange().getValues();
  var encabezados = datos[0];

  var colEstado = encabezados.indexOf("Estado Liquidacion");
  var colFechaLiq = encabezados.indexOf("Fecha de Liquidacion");
  var colALiquidar = encabezados.indexOf("A Liquidar");
  var colUsdt = encabezados.indexOf("USDT");

  if (colEstado === -1 || colFechaLiq === -1 || colALiquidar === -1 || colUsdt === -1) {
    ui.alert("❌ Error: Faltan columnas en 'Transacciones'. Asegurate de tener los títulos 'Estado Liquidacion', 'Fecha de Liquidacion', 'A Liquidar' y 'USDT' en la fila 1.");
    return;
  }

  var tope = new Date();
  tope.setHours(23, 59, 59, 999);
  var filasActualizadas = 0;

  for (var i = 1; i < datos.length; i++) {
    var estado = datos[i][colEstado];
    var fecha = new Date(datos[i][colFechaLiq]);

    if (estado === "Pendiente" && fecha <= tope) {
      var montoEur = parseFloat(String(datos[i][colALiquidar]).replace("€", "").replace(",", "."));
      if (!isNaN(montoEur)) {
        var montoUsdt = montoEur / tipoDeCambio;
        hojaTransacciones.getRange(i + 1, colUsdt + 1).setValue(montoUsdt);
        filasActualizadas++;
      }
    }
  }

  ui.alert("✅ Tasa de " + tipoDeCambio + " EUR/USDT guardada.\n\nSe calcularon los USDT exactos (dividiendo por la tasa) para " + filasActualizadas + " pagos pendientes.");
}

// =========================================================
// ACTUALIZAR LIQUIDACIONES
// Cierra la semana:
// - Marca transacciones pendientes como "Procesado"
// - Protege las filas cerradas (warning only)
// - Escribe resumen en "Historial de Pagos"
// - Genera "SALDO DEUDOR" si el balance de la creadora es negativo
// - Registra el fondeo a Kraken en "Fondeo Kraken"
// =========================================================
function actualizarLiquidaciones() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaTransacciones = ss.getSheetByName("Transacciones");
  var hojaHistorial = ss.getSheetByName("Historial de Pagos");
  var hojaLiquidacion = ss.getSheetByName("Semanal - Liquidacion");
  var ui = SpreadsheetApp.getUi();

  if (!hojaTransacciones || !hojaHistorial) {
    ui.alert("❌ Error: Faltan las pestañas 'Transacciones' o 'Historial de Pagos'.");
    return;
  }

  var eurosKrakenResponse = ui.prompt("Fondeo Kraken 🏦", "¿Cuántos Euros enviaste a Kraken para esta liquidación?\n\n👉 Dejá en blanco o poné 0 si no enviaste nada.", ui.ButtonSet.OK_CANCEL);
  if (eurosKrakenResponse.getSelectedButton() === ui.Button.CANCEL) return;

  var eurosKrakenText = eurosKrakenResponse.getResponseText().replace(",", ".").trim();
  var eurosKraken = parseFloat(eurosKrakenText) || 0;

  var fechaKraken = "";
  var fechaHoy = new Date();

  if (eurosKraken > 0) {
    var fechaKrakenResponse = ui.prompt("Fondeo Kraken 🏦", "¿En qué fecha enviaste los euros a Kraken? (DD/MM/AAAA)\n\n👉 Dejá en blanco para usar la fecha de hoy.", ui.ButtonSet.OK_CANCEL);
    if (fechaKrakenResponse.getSelectedButton() === ui.Button.CANCEL) return;

    fechaKraken = fechaKrakenResponse.getResponseText().trim();
    if (fechaKraken === "") {
      var d = ("0" + fechaHoy.getDate()).slice(-2);
      var m = ("0" + (fechaHoy.getMonth() + 1)).slice(-2);
      var a = fechaHoy.getFullYear();
      fechaKraken = d + "/" + m + "/" + a;
    }
  }

  var rangoDatos = hojaTransacciones.getDataRange();
  var valores = rangoDatos.getValues();
  var encabezados = valores[0];

  var colEstado = encabezados.indexOf("Estado Liquidacion");
  var colFechaLiq = encabezados.indexOf("Fecha de Liquidacion");
  var colID = encabezados.indexOf("ID Proveedor");
  var colBruto = encabezados.indexOf("Bruto");
  var colALiquidar = encabezados.indexOf("A Liquidar");
  var colUsdt = encabezados.indexOf("USDT");
  var colHonorarios = encabezados.indexOf("Honorarios");

  var colComision = -1;
  for (var c = 0; c < encabezados.length; c++) {
    var titulo = encabezados[c].toString().toLowerCase();
    if (titulo.indexOf("comision") !== -1 || titulo.indexOf("fee") !== -1 || titulo.indexOf("pasarela") !== -1) {
      colComision = c; break;
    }
  }

  if (colEstado === -1 || colFechaLiq === -1 || colID === -1) {
    ui.alert("❌ Error: Faltan columnas clave en 'Transacciones'.");
    return;
  }

  fechaHoy.setHours(23, 59, 59, 999);
  var filasProcesadas = 0;
  var deudasGeneradas = 0;
  var resumenPorID = {};
  var fechaLiquidacionResumen = "";

  function limpiarNumero(val) {
    if (!val) return 0;
    return parseFloat(String(val).replace(/[€$]/g, "").replace(",", ".").trim()) || 0;
  }

  for (var i = 1; i < valores.length; i++) {
    var estadoActual = valores[i][colEstado];
    var fechaCelda = new Date(valores[i][colFechaLiq]);

    if (estadoActual === "Pendiente" && fechaCelda <= fechaHoy) {
      var idActual = String(valores[i][colID]).trim();

      if (fechaLiquidacionResumen === "") {
        var dia = ("0" + fechaHoy.getDate()).slice(-2);
        var mes = ("0" + (fechaHoy.getMonth() + 1)).slice(-2);
        var anio = fechaHoy.getFullYear();
        fechaLiquidacionResumen = dia + "/" + mes + "/" + anio;
      }

      if (!resumenPorID[idActual]) {
        resumenPorID[idActual] = { bruto: 0, comision: 0, honorarios: 0, aLiquidar: 0, usdt: 0 };
      }

      resumenPorID[idActual].bruto += colBruto !== -1 ? limpiarNumero(valores[i][colBruto]) : 0;
      resumenPorID[idActual].comision += colComision !== -1 ? limpiarNumero(valores[i][colComision]) : 0;
      resumenPorID[idActual].honorarios += colHonorarios !== -1 ? limpiarNumero(valores[i][colHonorarios]) : 0;
      resumenPorID[idActual].aLiquidar += colALiquidar !== -1 ? limpiarNumero(valores[i][colALiquidar]) : 0;
      resumenPorID[idActual].usdt += colUsdt !== -1 ? limpiarNumero(valores[i][colUsdt]) : 0;

      hojaTransacciones.getRange(i + 1, colEstado + 1).setValue("Procesado");
      var rangoFila = hojaTransacciones.getRange(i + 1, 1, 1, encabezados.length);
      var proteccion = rangoFila.protect().setDescription("Liquidacion Cerrada (Fila " + (i + 1) + ")");
      proteccion.setWarningOnly(true);

      filasProcesadas++;
    }
  }

  var tipoDeCambio = "No especificado";
  if (hojaLiquidacion) {
    var colK = hojaLiquidacion.getRange("K:K").getValues();
    var colL = hojaLiquidacion.getRange("L:L").getValues();
    var indiceUltimaFila = -1;
    for (var j = 0; j < colK.length; j++) {
      if (String(colK[j][0]).trim() !== "") {
        indiceUltimaFila = j;
      }
    }
    if (indiceUltimaFila !== -1) {
      tipoDeCambio = colL[indiceUltimaFila][0];
    }
  }

  var idsLiquidados = Object.keys(resumenPorID);

  if (idsLiquidados.length > 0) {
    var ultimaFilaHist = hojaHistorial.getLastRow();

    for (var k = 0; k < idsLiquidados.length; k++) {
      var idL = idsLiquidados[k];
      var datosID = resumenPorID[idL];
      var reservaHacienda = datosID.honorarios * 0.20;

      var nuevaFilaHistorial = [
        fechaLiquidacionResumen,
        idL,
        datosID.bruto,
        datosID.comision,
        datosID.honorarios,
        datosID.aLiquidar,
        Math.trunc(datosID.usdt),
        tipoDeCambio,
        reservaHacienda
      ];

      hojaHistorial.getRange(ultimaFilaHist + 1 + k, 1, 1, 9).setValues([nuevaFilaHistorial]);

      if (datosID.aLiquidar < 0) {
        var proxViernes = new Date(fechaHoy);
        var diasFaltantes = (5 - proxViernes.getDay() + 7) % 7;
        if (diasFaltantes === 0) diasFaltantes = 7;
        proxViernes.setDate(proxViernes.getDate() + diasFaltantes);

        var dV = ("0" + proxViernes.getDate()).slice(-2);
        var mV = ("0" + (proxViernes.getMonth() + 1)).slice(-2);
        var aV = proxViernes.getFullYear();
        var fechaProxViernes = dV + "/" + mV + "/" + aV;

        var fechaHoraActual = ("0" + fechaHoy.getDate()).slice(-2) + "/" + ("0" + (fechaHoy.getMonth() + 1)).slice(-2) + "/" + fechaHoy.getFullYear() + " " + fechaHoy.getHours() + ":" + fechaHoy.getMinutes() + ":" + fechaHoy.getSeconds();

        var filaDeuda = [
          fechaHoraActual, "SALDO DEUDOR SEMANA ANTERIOR", "ARRASTRE", "", idL,
          datosID.aLiquidar, 0, "", "", "", fechaProxViernes, "Pendiente"
        ];

        hojaTransacciones.appendRow(filaDeuda);
        var row = hojaTransacciones.getLastRow();

        hojaTransacciones.getRange(row, 8).setFormula("=F" + row + "-G" + row);
        hojaTransacciones.getRange(row, 9).setFormula('=IF(B' + row + '="SALDO DEUDOR SEMANA ANTERIOR", 0, ABS(H' + row + ')*0.2)');
        hojaTransacciones.getRange(row, 10).setFormula("=H" + row + "-I" + row);

        deudasGeneradas++;
      }
    }

    var rangoAFiltrar = hojaHistorial.getRange(2, 1, hojaHistorial.getLastRow() - 1, 9);
    rangoAFiltrar.sort([{column: 2, ascending: true}, {column: 1, ascending: false}]);

    if (eurosKraken > 0) {
      var hojaFondeo = ss.getSheetByName("Fondeo Kraken");
      if (hojaFondeo) {
        var usdtCompradosReales = 0;
        var tasaLimpia = 0;
        if (tipoDeCambio !== "No especificado") {
          tasaLimpia = parseFloat(String(tipoDeCambio).replace(",", ".").trim());
        }
        if (tasaLimpia > 0) {
          usdtCompradosReales = eurosKraken / tasaLimpia;
        }
        hojaFondeo.appendRow([
          fechaKraken, eurosKraken,
          usdtCompradosReales > 0 ? parseFloat(usdtCompradosReales.toFixed(2)) : "Falta Tipo de Cambio",
          tipoDeCambio, "Liq. " + fechaLiquidacionResumen
        ]);
      } else {
        ui.alert("⚠️ Aviso: No se encontró la pestaña 'Fondeo Kraken'. El registro del envío no se guardó.");
      }
    }
  }

  if (filasProcesadas > 0) {
    var msjExtra = deudasGeneradas > 0 ? "\nSe generaron " + deudasGeneradas + " saldos deudores para la próxima semana." : "";
    ui.alert("🔒 ¡Liquidacion Cerrada!\n\nSe procesaron " + filasProcesadas + " transacciones." + msjExtra);
  } else {
    ui.alert("ℹ️ No habia nada pendiente con fecha de hoy o anterior para procesar.");
  }
}

// =========================================================
// GENERAR EXTRACTO PDF
// Genera PDFs semanales por creadora con el desglose de
// transacciones, fees, honorarios y USDT enviados.
// Los guarda en carpeta "Extractos Semanales" de Drive.
// Acepta IDs separados por coma o vacío para TODOS.
// =========================================================
function generarExtractoPDF() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaTransacciones = ss.getSheetByName("Transacciones");
  var hojaLiquidacion = ss.getSheetByName("Semanal - Liquidacion");

  if (!hojaTransacciones) {
    ui.alert("❌ Error: No se encontró la pestaña 'Transacciones'.");
    return;
  }

  var idResponse = ui.prompt(
    "Generar Extracto(s) 📄",
    "Ingresá los IDs separados por coma (ej: 42, 21, 63)\n\n👉 Dejá el texto VACÍO y dale a OK para generar TODOS automáticamente:",
    ui.ButtonSet.OK_CANCEL
  );
  if (idResponse.getSelectedButton() !== ui.Button.OK) return;
  var inputIDs = idResponse.getResponseText().trim().toUpperCase();
  if (inputIDs === "") inputIDs = "TODOS";

  var hoy = new Date();
  var diaHoy = ("0" + hoy.getDate()).slice(-2);
  var mesHoy = ("0" + (hoy.getMonth() + 1)).slice(-2);
  var anioHoy = hoy.getFullYear();
  var fechaSugerida = diaHoy + "/" + mesHoy + "/" + anioHoy;

  var fechaResponse = ui.prompt(
    "Generar Extracto(s) 📄",
    "Fecha de Liquidación a buscar.\n\n👉 Dejá el texto VACÍO y dale a OK para usar la fecha de hoy: " + fechaSugerida + "\n(O escribí manualmente otra fecha):",
    ui.ButtonSet.OK_CANCEL
  );
  if (fechaResponse.getSelectedButton() !== ui.Button.OK) return;
  var fechaBuscada = fechaResponse.getResponseText().trim();
  if (fechaBuscada === "") fechaBuscada = fechaSugerida;

  var tipoDeCambio = "No especificado";
  if (hojaLiquidacion) {
    var datosHistorial = hojaLiquidacion.getRange("K:L").getValues();
    for (var h = 0; h < datosHistorial.length; h++) {
      var fechaHist = datosHistorial[h][0];
      if (fechaHist !== "") {
        var fechaHistStr = "";
        if (fechaHist instanceof Date) {
          var diaH = ("0" + fechaHist.getDate()).slice(-2);
          var mesH = ("0" + (fechaHist.getMonth() + 1)).slice(-2);
          var anioH = fechaHist.getFullYear();
          fechaHistStr = diaH + "/" + mesH + "/" + anioH;
        } else {
          fechaHistStr = String(fechaHist).trim();
        }
        if (fechaHistStr === fechaBuscada) {
          tipoDeCambio = datosHistorial[h][1];
          break;
        }
      }
    }
  }

  var datos = hojaTransacciones.getDataRange().getValues();
  var encabezados = datos[0];

  var colID = encabezados.indexOf("ID Proveedor");
  var colFechaVenta = encabezados.indexOf("Fecha y Hora");
  var colOrigen = encabezados.indexOf("Origen / TIpo");
  if (colOrigen === -1) colOrigen = encabezados.indexOf("Origen");
  var colBruto = encabezados.indexOf("Bruto");
  var colHonorarios = encabezados.indexOf("Honorarios");
  var colALiquidar = encabezados.indexOf("A Liquidar");
  var colUsdt = encabezados.indexOf("USDT");
  var colEstado = encabezados.indexOf("Estado Liquidacion");
  var colFechaLiq = encabezados.indexOf("Fecha de Liquidacion");

  var colComision = -1;
  for (var c = 0; c < encabezados.length; c++) {
    var titulo = encabezados[c].toString().toLowerCase();
    if (titulo.indexOf("comision") !== -1 || titulo.indexOf("fee") !== -1 || titulo.indexOf("pasarela") !== -1) {
      colComision = c; break;
    }
  }

  if (colID === -1 || colALiquidar === -1 || colUsdt === -1 || colEstado === -1 || colFechaLiq === -1) {
    ui.alert("❌ Error: Faltan columnas clave en la pestaña 'Transacciones'.");
    return;
  }

  var transaccionesFecha = [];
  for (var i = 1; i < datos.length; i++) {
    var fila = datos[i];
    if (fila[colEstado] === "Procesado") {
      var fechaCelda = fila[colFechaLiq];
      var fechaFormateada = "";
      if (fechaCelda instanceof Date) {
        var dia = ("0" + fechaCelda.getDate()).slice(-2);
        var mes = ("0" + (fechaCelda.getMonth() + 1)).slice(-2);
        var anio = fechaCelda.getFullYear();
        fechaFormateada = dia + "/" + mes + "/" + anio;
      } else {
        fechaFormateada = String(fechaCelda).trim();
      }
      if (fechaFormateada === fechaBuscada) {
        transaccionesFecha.push(fila);
      }
    }
  }

  if (transaccionesFecha.length === 0) {
    ui.alert("ℹ️ No encontré ninguna transacción 'Procesada' en la fecha " + fechaBuscada + ".");
    return;
  }

  var idsAProcesar = [];
  if (inputIDs === "TODOS") {
    var idsUnicos = {};
    for (var f = 0; f < transaccionesFecha.length; f++) {
      var idCreadora = String(transaccionesFecha[f][colID]).trim();
      if (idCreadora !== "") idsUnicos[idCreadora] = true;
    }
    idsAProcesar = Object.keys(idsUnicos);
  } else {
    var partes = inputIDs.replace(/;/g, ",").split(",");
    for (var p = 0; p < partes.length; p++) {
      var idLimpiado = partes[p].trim();
      if (idLimpiado !== "") idsAProcesar.push(idLimpiado);
    }
  }

  if (idsAProcesar.length === 0) return;

  function limpiarNumero(val) {
    if (!val) return 0;
    return parseFloat(String(val).replace(/[€$]/g, "").replace(",", ".").trim()) || 0;
  }

  var carpetas = DriveApp.getFoldersByName("Extractos Semanales");
  var carpetaDestino = carpetas.hasNext() ? carpetas.next() : DriveApp.createFolder("Extractos Semanales");

  var generadosCount = 0;
  var listaNombres = [];

  for (var k = 0; k < idsAProcesar.length; k++) {
    var idActual = idsAProcesar[k];
    var transaccionesDelID = [];
    var totalUsdt = 0;
    var totalEuros = 0;

    for (var t = 0; t < transaccionesFecha.length; t++) {
      if (String(transaccionesFecha[t][colID]).trim() === idActual) {
        transaccionesDelID.push(transaccionesFecha[t]);
        totalUsdt += limpiarNumero(transaccionesFecha[t][colUsdt]);
      }
    }

    if (transaccionesDelID.length === 0) continue;

    var html = "<div style='font-family: Arial, sans-serif; color: #333;'>";
    html += "<h2 style='color: #2c3e50; border-bottom: 2px solid #2c3e50; padding-bottom: 5px;'>Talaria Digital - Liquidación de Honorarios</h2>";
    html += "<p><strong>ID de Creadora:</strong> " + idActual + "<br>";
    html += "<strong>Fecha de Liquidación:</strong> " + fechaBuscada + "</p>";
    html += "<table border='1' cellpadding='8' cellspacing='0' style='width: 100%; border-collapse: collapse; font-size: 14px; text-align: right;'>";
    html += "<tr style='background-color: #f8f9fa; color: #333; text-align: center;'>";
    html += "<th>Fecha</th><th>Origen</th><th>Bruto</th><th>Pasarela (Fees)</th><th>Honorarios</th><th>A Liquidar</th></tr>";

    for (var f = 0; f < transaccionesDelID.length; f++) {
      var filaT = transaccionesDelID[f];
      var valFecha = "-";
      if (colFechaVenta !== -1 && filaT[colFechaVenta] instanceof Date) {
        valFecha = ("0" + filaT[colFechaVenta].getDate()).slice(-2) + "/" + ("0" + (filaT[colFechaVenta].getMonth()+1)).slice(-2);
      } else if (colFechaVenta !== -1) {
        valFecha = String(filaT[colFechaVenta]).split(" ")[0];
      }
      var valOrigen = colOrigen !== -1 ? filaT[colOrigen] : "-";
      var valBruto = colBruto !== -1 ? limpiarNumero(filaT[colBruto]) : 0;
      var valComision = colComision !== -1 ? limpiarNumero(filaT[colComision]) : 0;
      var valHonorarios = colHonorarios !== -1 ? limpiarNumero(filaT[colHonorarios]) : 0;
      var valALiquidar = colALiquidar !== -1 ? limpiarNumero(filaT[colALiquidar]) : 0;

      totalEuros += valALiquidar;

      html += "<tr>";
      html += "<td style='text-align: center;'>" + valFecha + "</td>";
      html += "<td style='text-align: left;'>" + valOrigen + "</td>";
      html += "<td>€ " + valBruto.toFixed(2) + "</td>";
      html += "<td style='color: #c0392b;'>-€ " + valComision.toFixed(2) + "</td>";
      html += "<td style='color: #c0392b;'>-€ " + valHonorarios.toFixed(2) + "</td>";
      html += "<td style='font-weight: bold; color: #27ae60;'>€ " + valALiquidar.toFixed(2) + "</td>";
      html += "</tr>";
    }
    html += "</table>";
    html += "<h3 style='text-align: right; margin-top: 20px; font-size: 18px; color: #27ae60;'>TOTAL A LIQUIDAR: € " + totalEuros.toFixed(2) + "</h3>";
    html += "<h3 style='text-align: right; margin-top: 5px; font-size: 20px; color: #2980b9;'>TOTAL ENVIADO: " + Math.trunc(totalUsdt) + " USDT</h3>";
    html += "<p style='font-size: 11px; color: #7f8c8d; margin-top: 30px;'>* Las comisiones de infraestructura (Stripe/Bizum) y los honorarios de gestión (20%) ya han sido descontados del Bruto para obtener el monto A Liquidar en Euros.<br>";
    if (tipoDeCambio !== "No especificado") {
      html += "* Tipo de cambio aplicado (Kraken): <strong>1 USDT = " + tipoDeCambio + " EUR</strong><br>";
    }
    html += "* El Total Enviado refleja la conversión a Criptoactivos (USDT) al tipo de cambio de mercado ejecutado en la fecha de liquidación.</p>";
    html += "</div>";

    var nombreArchivo = "Extracto_ID" + idActual + "_" + fechaBuscada.replace(/\//g, "-") + ".pdf";
    var blob = Utilities.newBlob(html, MimeType.HTML, nombreArchivo).getAs(MimeType.PDF);

    var archivosExistentes = carpetaDestino.getFilesByName(nombreArchivo);
    while (archivosExistentes.hasNext()) {
      archivosExistentes.next().setTrashed(true);
    }

    carpetaDestino.createFile(blob);
    generadosCount++;
    listaNombres.push(nombreArchivo);
  }

  if (generadosCount > 0) {
    ui.alert("✅ ¡Generación Masiva Exitosa!\n\nSe crearon " + generadosCount + " extractos en tu Google Drive.\n\nArchivos guardados:\n" + listaNombres.join("\n"));
  } else {
    ui.alert("ℹ️ No se generó ningún extracto. Revisá que los IDs ingresados realmente tengan pagos 'Procesados' en la fecha elegida.");
  }
}

// =========================================================
// ENVIAR EXTRACTOS POR EMAIL
// Envía los PDFs de la carpeta "Extractos Semanales" a cada
// creadora por email. Anota la fecha de envío en Col R de
// "Semanal - Liquidacion".
// Emails mapeados desde Col Q (índice 16) de "Semanal - Liquidacion".
// =========================================================
function enviarExtractosPorEmail() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaLiquidacion = ss.getSheetByName("Semanal - Liquidacion");

  if (!hojaLiquidacion) {
    ui.alert("❌ Error: No se encontró la pestaña 'Semanal - Liquidacion' para buscar los emails.");
    return;
  }

  var idResponse = ui.prompt(
    "Enviar Extracto(s) ✉️",
    "Ingresá los IDs separados por coma (ej: 42, 21, 63)\n\n👉 Dejá el texto VACÍO y dale a OK para enviar TODOS automáticamente:",
    ui.ButtonSet.OK_CANCEL
  );
  if (idResponse.getSelectedButton() !== ui.Button.OK) return;
  var inputIDs = idResponse.getResponseText().trim().toUpperCase();
  if (inputIDs === "") inputIDs = "TODOS";

  var hoy = new Date();
  var diaHoy = ("0" + hoy.getDate()).slice(-2);
  var mesHoy = ("0" + (hoy.getMonth() + 1)).slice(-2);
  var anioHoy = hoy.getFullYear();
  var fechaSugerida = diaHoy + "/" + mesHoy + "/" + anioHoy;

  var fechaResponse = ui.prompt(
    "Enviar Extracto(s) ✉️",
    "Fecha de los extractos a enviar.\n\n👉 Dejá VACÍO para usar hoy: " + fechaSugerida + "\n(O escribí la fecha exacta DD/MM/AAAA):",
    ui.ButtonSet.OK_CANCEL
  );
  if (fechaResponse.getSelectedButton() !== ui.Button.OK) return;
  var fechaBuscada = fechaResponse.getResponseText().trim();
  if (fechaBuscada === "") fechaBuscada = fechaSugerida;

  // Mapear emails: Col O (índice 14) = ID, Col Q (índice 16) = Email
  var datosLiquidacion = hojaLiquidacion.getDataRange().getValues();
  var idDataMap = {};

  for (var i = 0; i < datosLiquidacion.length; i++) {
    var idFila = String(datosLiquidacion[i][14]).trim();
    var emailFila = String(datosLiquidacion[i][16]).trim();

    if (idFila !== "" && idFila !== "ID" && emailFila.indexOf("@") !== -1) {
      idDataMap[idFila] = { email: emailFila, filaExcel: i + 1 };
    }
  }

  if (Object.keys(idDataMap).length === 0) {
    ui.alert("❌ Error: No se encontraron correos electrónicos válidos en las columnas O y Q.");
    return;
  }

  var idsAProcesar = [];
  if (inputIDs === "TODOS") {
    idsAProcesar = Object.keys(idDataMap);
  } else {
    var partes = inputIDs.replace(/;/g, ",").split(",");
    for (var p = 0; p < partes.length; p++) {
      var idLimpiado = partes[p].trim();
      if (idLimpiado !== "") idsAProcesar.push(idLimpiado);
    }
  }

  if (idsAProcesar.length === 0) return;

  var carpetas = DriveApp.getFoldersByName("Extractos Semanales");
  if (!carpetas.hasNext()) {
    ui.alert("❌ Error: No se encontró la carpeta 'Extractos Semanales' en Google Drive.");
    return;
  }
  var carpetaExtractos = carpetas.next();

  var enviadosCount = 0;
  var erroresCount = 0;
  var resumenErrores = "";

  for (var k = 0; k < idsAProcesar.length; k++) {
    var idActual = idsAProcesar[k];

    if (!idDataMap[idActual]) {
      erroresCount++;
      resumenErrores += "- ID " + idActual + ": No tiene email registrado.\n";
      continue;
    }

    var dataDestino = idDataMap[idActual];
    var nombreArchivo = "Extracto_ID" + idActual + "_" + fechaBuscada.replace(/\//g, "-") + ".pdf";
    var archivosEncontrados = carpetaExtractos.getFilesByName(nombreArchivo);

    if (archivosEncontrados.hasNext()) {
      var archivoPdf = archivosEncontrados.next();
      var asunto = "Liquidación Talaria Digital - ID " + idActual + " (" + fechaBuscada + ")";
      var cuerpo = "Hola,\n\nAdjunto encontrarás el extracto detallado de la liquidación correspondiente al " + fechaBuscada + ".\n\n";
      cuerpo += "El envío de los USDT ya se encuentra procesado hacia tu wallet registrada.\n\n";
      cuerpo += "Cualquier consulta o duda sobre las operaciones detalladas, por favor respondenos a este correo.\n\n";
      cuerpo += "Saludos,\nEquipo de Talaria Digital";

      try {
        MailApp.sendEmail({
          to: dataDestino.email,
          subject: asunto,
          body: cuerpo,
          attachments: [archivoPdf.getAs(MimeType.PDF)]
        });
        enviadosCount++;
        // Col R (índice 18) = fecha último extracto enviado
        hojaLiquidacion.getRange(dataDestino.filaExcel, 18).setValue(fechaBuscada);
      } catch (e) {
        erroresCount++;
        resumenErrores += "- ID " + idActual + ": Error al intentar enviar el correo.\n";
      }
    } else {
      erroresCount++;
      resumenErrores += "- ID " + idActual + ": No se encontró el PDF '" + nombreArchivo + "' en Drive.\n";
    }
  }

  if (erroresCount === 0 && enviadosCount > 0) {
    ui.alert("✅ ¡Éxito Total!\n\nSe enviaron correctamente " + enviadosCount + " extractos por email y se actualizó la fecha en la planilla.");
  } else if (enviadosCount > 0 && erroresCount > 0) {
    ui.alert("⚠️ Envío Parcial\n\nSe enviaron " + enviadosCount + " correos, pero hubo " + erroresCount + " errores:\n\n" + resumenErrores);
  } else {
    ui.alert("❌ Error: No se pudo enviar ningún correo.\n\n" + resumenErrores);
  }
}

// =========================================================
// GENERAR FACTURAS MENSUALES
// Suma honorarios del mes por creadora desde "Transacciones".
// Genera factura PDF legal con datos fiscales (Art. 69 / sin IRPF).
// Guarda en carpeta "Facturas" de Drive y registra en "Registro Facturas".
// Datos de creadoras: Col O=ID, S=Nombre (índice 18), T=CUIT (índice 19)
// =========================================================
function generarFacturasMensuales() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaTransacciones = ss.getSheetByName("Transacciones");
  var hojaLiquidacion = ss.getSheetByName("Semanal - Liquidacion");
  var hojaRegistro = ss.getSheetByName("Registro Facturas");

  if (!hojaTransacciones || !hojaRegistro) {
    ui.alert("❌ Error: Faltan pestañas necesarias (Transacciones o Registro Facturas).");
    return;
  }

  var idResponse = ui.prompt("Generar Facturas 🧾", "Ingresá IDs separados por coma (ej: 42, 21)\n👉 Dejá VACÍO para generar TODOS:", ui.ButtonSet.OK_CANCEL);
  if (idResponse.getSelectedButton() !== ui.Button.OK) return;
  var inputIDs = idResponse.getResponseText().trim().toUpperCase() || "TODOS";

  var mesResponse = ui.prompt("Mes a Facturar 📅", "Ingresá el número del mes (ej: 3 para Marzo):", ui.ButtonSet.OK_CANCEL);
  if (mesResponse.getSelectedButton() !== ui.Button.OK) return;
  var mesFacturar = parseInt(mesResponse.getResponseText().trim());

  var anioActual = 2026;

  // Mapear datos de creadoras: Col O=ID (índice 14), S=Nombre (índice 18), T=CUIT (índice 19)
  var datosLiq = hojaLiquidacion.getDataRange().getValues();
  var datosCreadoras = {};
  for (var i = 1; i < datosLiq.length; i++) {
    var idFila = String(datosLiq[i][14]).trim();
    if (idFila !== "") {
      datosCreadoras[idFila] = {
        nombre: String(datosLiq[i][18] || "Sin Nombre"),
        cuit: String(datosLiq[i][19] || "Sin CUIT")
      };
    }
  }

  var transDatos = hojaTransacciones.getDataRange().getValues();
  var encT = transDatos[0];
  var colEstado = encT.indexOf("Estado Liquidacion");
  var colFechaLiq = encT.indexOf("Fecha de Liquidacion");
  var colID = encT.indexOf("ID Proveedor");
  var colHonorarios = encT.indexOf("Honorarios");

  var facturacionPorID = {};

  for (var t = 1; t < transDatos.length; t++) {
    if (transDatos[t][colEstado] === "Procesado") {
      var fecha = new Date(transDatos[t][colFechaLiq]);
      if (fecha.getMonth() + 1 === mesFacturar && fecha.getFullYear() === anioActual) {
        var id = String(transDatos[t][colID]).trim();
        var honorarios = parseFloat(String(transDatos[t][colHonorarios]).replace(/[€$]/g, "").replace(",", ".")) || 0;
        if (!facturacionPorID[id]) facturacionPorID[id] = 0;
        facturacionPorID[id] += honorarios;
      }
    }
  }

  var idsAProcesar = inputIDs === "TODOS" ? Object.keys(facturacionPorID) : inputIDs.replace(/;/g, ",").split(",").map(function(item) { return item.trim(); });

  var carpetas = DriveApp.getFoldersByName("Facturas");
  var carpetaDestino = carpetas.hasNext() ? carpetas.next() : DriveApp.createFolder("Facturas");
  var generadas = 0;

  for (var k = 0; k < idsAProcesar.length; k++) {
    var idActual = idsAProcesar[k];
    if (!facturacionPorID[idActual] || facturacionPorID[idActual] <= 0) continue;

    var totalHonorarios = facturacionPorID[idActual];
    var dataCliente = datosCreadoras[idActual] || {nombre: "Consumidor Final", cuit: "N/A"};

    var ultimaFilaReg = hojaRegistro.getLastRow();
    var numCorrelativo = ultimaFilaReg === 0 ? 1 : ultimaFilaReg;
    var numeroFactura = anioActual + "-" + ("000" + numCorrelativo).slice(-3);
    var fechaEmision = new Date().toLocaleDateString('es-ES');

    var html = "<div style='font-family: Arial, sans-serif; color: #333; max-width: 800px; margin: auto;'>";
    html += "<table width='100%'><tr>";
    html += "<td width='50%'><img src='https://i.imgur.com/jTwtbwK.png' width='150'></td>";
    html += "<td width='50%' style='text-align: right;'><h2>FACTURA</h2><b>Nº:</b> " + numeroFactura + "<br><b>Fecha:</b> " + fechaEmision + "</td>";
    html += "</tr></table><hr style='border-top: 2px solid #2c3e50;'>";

    html += "<table width='100%' style='margin-top:20px; margin-bottom: 30px; font-size: 14px;'><tr>";
    html += "<td width='50%'><b>Emisor:</b><br>Jonas Gabriel Ramirez Torres<br>NIE/NIF: Z3318599A<br>Avenida De La Estacion 24<br>50298 Pinseque, Zaragoza<br>España</td>";
    html += "<td width='50%' style='text-align: right;'><b>Cliente (ID " + idActual + "):</b><br>" + dataCliente.nombre + "<br>CUIT: " + dataCliente.cuit + "<br>Argentina</td>";
    html += "</tr></table>";

    html += "<table border='1' cellpadding='10' cellspacing='0' style='width: 100%; border-collapse: collapse; text-align: left; font-size: 14px;'>";
    html += "<tr style='background-color: #f8f9fa; color: #333;'><th>Concepto</th><th style='text-align: right;'>Importe</th></tr>";
    html += "<tr><td>Servicios de gestión de agencia digital y marketing - Mes " + ("0"+mesFacturar).slice(-2) + "/" + anioActual + "</td><td style='text-align: right;'>€ " + totalHonorarios.toFixed(2) + "</td></tr>";
    html += "</table>";

    html += "<h3 style='text-align: right; color: #2c3e50; margin-top: 20px;'>TOTAL: € " + totalHonorarios.toFixed(2) + "</h3>";
    html += "<p style='font-size: 10px; color: #7f8c8d; margin-top: 50px; text-align: justify;'><i>Operación no sujeta a IVA en aplicación de las reglas de localización de los servicios (Art. 69 de la Ley 37/1992). Operación no sujeta a retención de IRPF por destinatario no residente.</i></p>";
    html += "</div>";

    var nombreArchivo = "Factura_" + numeroFactura + "_ID" + idActual + "_" + mesFacturar + "-" + anioActual + ".pdf";
    var blob = Utilities.newBlob(html, MimeType.HTML, nombreArchivo).getAs(MimeType.PDF);
    carpetaDestino.createFile(blob);

    hojaRegistro.appendRow([numeroFactura, fechaEmision, idActual, dataCliente.nombre, totalHonorarios.toFixed(2)]);
    generadas++;
  }

  ui.alert(generadas > 0 ? "✅ Se generaron " + generadas + " facturas en la carpeta 'Facturas'." : "ℹ️ No se encontraron honorarios para facturar en ese mes/IDs.");
}

// =========================================================
// ENVIAR FACTURAS POR EMAIL
// Envía las facturas generadas a los emails de las creadoras.
// Busca PDFs por patrón: _ID{id}_{mes}-2026.pdf
// Emails desde Col Q (índice 16) de "Semanal - Liquidacion".
// =========================================================
function enviarFacturasPorEmail() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaLiquidacion = ss.getSheetByName("Semanal - Liquidacion");

  var idResponse = ui.prompt("Enviar Facturas ✉️", "Ingresá IDs separados por coma\n👉 Dejá VACÍO para enviar a TODOS:", ui.ButtonSet.OK_CANCEL);
  if (idResponse.getSelectedButton() !== ui.Button.OK) return;
  var inputIDs = idResponse.getResponseText().trim().toUpperCase() || "TODOS";

  var mesResponse = ui.prompt("Mes a Enviar 📅", "Número del mes de la factura (ej: 3 para Marzo):", ui.ButtonSet.OK_CANCEL);
  if (mesResponse.getSelectedButton() !== ui.Button.OK) return;
  var mesEnviar = parseInt(mesResponse.getResponseText().trim());

  var datosLiq = hojaLiquidacion.getDataRange().getValues();
  var correos = {};
  for (var i = 1; i < datosLiq.length; i++) {
    var idFila = String(datosLiq[i][14]).trim();
    var email = String(datosLiq[i][16]).trim();
    if (idFila !== "" && email.indexOf("@") !== -1) correos[idFila] = email;
  }

  var carpetas = DriveApp.getFoldersByName("Facturas");
  if (!carpetas.hasNext()) { ui.alert("❌ No existe la carpeta 'Facturas'."); return; }
  var carpetaFacturas = carpetas.next();

  var idsAProcesar = inputIDs === "TODOS" ? Object.keys(correos) : inputIDs.replace(/;/g, ",").split(",").map(function(i){return i.trim()});
  var enviados = 0;

  for (var k = 0; k < idsAProcesar.length; k++) {
    var id = idsAProcesar[k];
    if (!correos[id]) continue;

    var searchStr = "_ID" + id + "_" + mesEnviar + "-2026.pdf";
    var archivosIter = carpetaFacturas.searchFiles("title contains '" + searchStr + "'");

    if (archivosIter.hasNext()) {
      var pdf = archivosIter.next();
      MailApp.sendEmail({
        to: correos[id],
        subject: "Talaria Digital - Factura Mensual " + ("0"+mesEnviar).slice(-2) + "/2026",
        body: "Hola,\n\nAdjunto encontrarás la factura correspondiente a los honorarios de gestión del mes.\n\nSaludos,\nEquipo de Talaria Digital",
        attachments: [pdf.getAs(MimeType.PDF)]
      });
      enviados++;
    }
  }
  ui.alert("✅ Se enviaron " + enviados + " facturas por correo.");
}
