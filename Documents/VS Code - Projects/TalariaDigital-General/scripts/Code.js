function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏢 Talaria Digital')
    .addItem('💸 Cargar Gasto', 'cargarGasto')
    .addSeparator()
    .addItem('💱 Cargar Tipo de Cambio', 'cargarTipoCambio')
    .addItem('🔒 Actualizar Liquidaciones', 'actualizarLiquidaciones')
    .addSeparator()
    .addItem('📄 Generar Extractos PDF', 'generarExtractoPDF')
    .addItem('✉️ Enviar Extractos por Email', 'enviarExtractosPorEmail')
    .addSeparator()
    .addItem('🧾 Generar Facturas Mensuales', 'generarFacturasMensuales')
    .addItem('📬 Enviar Facturas por Email', 'enviarFacturasPorEmail')
    .addSeparator()
    .addItem('📋 Ayuda / Guía de Uso', 'mostrarAyuda')
    .addToUi();
}

// Parsea fechas que pueden venir como Date object o como string DD/MM/YYYY (con o sin hora)
function parseFechaTalaria(valor) {
  if (!valor) return null;
  if (valor instanceof Date) return valor;
  var str = String(valor).trim();
  var m = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1]));
  return null;
}

function cargarGasto() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Transacciones");

  if (!sheet) {
    ui.alert("❌ Error: No se encontró la pestaña 'Transacciones'.");
    return;
  }

  // 1. Preguntar ID
  var idResponse = ui.prompt("Cargar Gasto 💸", "1. Ingresá el ID de la Creadora (ej: 21):", ui.ButtonSet.OK_CANCEL);
  if (idResponse.getSelectedButton() !== ui.Button.OK || idResponse.getResponseText() === "") return;
  var idCreadora = idResponse.getResponseText().trim();

  // 2. Preguntar Monto
  var montoResponse = ui.prompt("Cargar Gasto 💸", "2. Ingresá el monto del gasto total en EUROS (ej: 120):", ui.ButtonSet.OK_CANCEL);
  if (montoResponse.getSelectedButton() !== ui.Button.OK || montoResponse.getResponseText() === "") return;
  
  var montoText = montoResponse.getResponseText().replace(",", ".").replace("€", "").trim();
  var monto = parseFloat(montoText);
  if (isNaN(monto) || monto <= 0) {
    ui.alert("❌ Error: Ingresá un número válido mayor a 0.");
    return;
  }

  // 3. Preguntar Concepto
  var conceptoResponse = ui.prompt("Cargar Gasto 💸", "3. ¿Qué tipo de gasto es? (ej: Publicidad, Diseño):", ui.ButtonSet.OK_CANCEL);
  if (conceptoResponse.getSelectedButton() !== ui.Button.OK) return;
  var concepto = conceptoResponse.getResponseText().trim().toUpperCase();
  if (concepto === "") concepto = "VARIOS";

  // 4. Preguntar Factura / Referencia
  var refResponse = ui.prompt("Cargar Gasto 💸", "4. Ingresá el ID de la transacción, recibo o factura:", ui.ButtonSet.OK_CANCEL);
  if (refResponse.getSelectedButton() !== ui.Button.OK) return;
  var referencia = refResponse.getResponseText().trim();

  // Damos formato a la fecha de hoy
  var fechaHoy = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");

  // El gasto entra en negativo en la columna Bruto. Comisión es 0.
  var bruto = -monto;
  var comision = 0;

  // Armamos la fila parcial (dejamos las columnas contables vacías para que el script meta las fórmulas)
  var nuevaFila = [
    fechaHoy,              // Col A (1): Fecha y Hora
    "GASTO - " + concepto, // Col B (2): Origen / Tipo
    referencia,            // Col C (3): ID Transaccion
    "",                    // Col D (4): Email Cliente Final
    idCreadora,            // Col E (5): ID Proveedor
    bruto,                 // Col F (6): Bruto
    comision,              // Col G (7): Comision
    "",                    // Col H (8): Neto (Se llena con fórmula)
    "",                    // Col I (9): Honorarios (Se llena con fórmula)
    "",                    // Col J (10): A Liquidar (Se llena con fórmula)
    "",                    // Col K (11): Fecha de Liquidacion (Se arrastra)
    "Pendiente"            // Col L (12): Estado Liquidacion
  ];

  // 1. Agregamos la fila a la tabla
  sheet.appendRow(nuevaFila);

  // 2. INYECTAMOS LAS FÓRMULAS en la fila recién creada para no romper la secuencia
  var row = sheet.getLastRow();
  
  // Col H: Neto = Bruto - Comision
  sheet.getRange(row, 8).setFormula("=F" + row + "-G" + row);
  
  // Col I: Honorarios — 0 si es saldo deudor, negativo si es reintegro ads, 20% normal en el resto
  sheet.getRange(row, 9).setFormula('=IF(OR(UPPER(B' + row + ')="SALDO DEUDOR SEMANA ANTERIOR",UPPER(B' + row + ')="RESERVA ADS",UPPER(B' + row + ')="REINTEGRO ADS"),0,ABS(H' + row + ')*0.2)');
  
  // Col J: A Liquidar = Neto - Honorarios (Ej: -100 - 20 = -120 de descuento)
  sheet.getRange(row, 10).setFormula("=H" + row + "-I" + row);

  // Col K: Copiamos la fórmula de fecha de la fila de arriba para que se calcule sola
  if (row > 2) {
    sheet.getRange(row - 1, 11).copyTo(sheet.getRange(row, 11));
  }

  // Calculamos para el cartelito de aviso final
  var misHonorarios = monto * 0.2;
  var aLiquidar = bruto - misHonorarios; // -100 - 20 = -120

  // Aviso final
  ui.alert("✅ Gasto registrado con éxito:\n\n" +
           "👤 ID: " + idCreadora + "\n" +
           "📝 Concepto: " + concepto + "\n" +
           "💶 Costo Gasto: €" + monto.toFixed(2) + "\n" +
           "📈 Tus Honorarios (20%): €" + misHonorarios.toFixed(2) + "\n" +
           "💥 Total a descontar: €" + Math.abs(aLiquidar).toFixed(2));
}

function cargarTipoCambio() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Pestañas
  var hojaTransacciones = ss.getSheetByName("Transacciones");
  var hojaLiquidacion = ss.getSheetByName("Semanal - Liquidacion");

  if (!hojaTransacciones || !hojaLiquidacion) {
    ui.alert("❌ Error: Revisá que las pestañas se llamen exactamente 'Transacciones' y 'Semanal - Liquidacion'.");
    return;
  }

  // 1. Pedir la tasa de Kraken (Modificado para el formato de Kraken)
  var rateResponse = ui.prompt("Cotización Kraken 💱", "Ingresá el Precio en EUR por cada USDT de hoy\n(El número que te da Kraken, ej: 0.8626):", ui.ButtonSet.OK_CANCEL);
  if (rateResponse.getSelectedButton() !== ui.Button.OK || rateResponse.getResponseText() === "") return;

  var rateText = rateResponse.getResponseText().replace(",", ".").trim();
  var tipoDeCambio = parseFloat(rateText);

  if (isNaN(tipoDeCambio) || tipoDeCambio <= 0) {
    ui.alert("❌ Error: Ingresá un número válido para la tasa de cambio.");
    return;
  }

  // 2. Guardar en el Historial (Columnas K y L de 'Semanal - Liquidacion')
  var colK = hojaLiquidacion.getRange("K:K").getValues();
  var ultimaFilaHistorial = 1;
  
  // Buscamos la última fila que tenga texto en la columna K
  for (var j = 0; j < colK.length; j++) {
    if (colK[j][0] !== "") {
      ultimaFilaHistorial = j + 1;
    }
  }
  
  var fechaHoy = new Date();
  hojaLiquidacion.getRange(ultimaFilaHistorial + 1, 11).setValue(fechaHoy); // Columna K
  hojaLiquidacion.getRange(ultimaFilaHistorial + 1, 12).setValue(tipoDeCambio); // Columna L

  // 3. Aplicar a las Transacciones (SIN TRUNCAR)
  var datos = hojaTransacciones.getDataRange().getValues();
  var encabezados = datos[0];
  
  var colEstado = encabezados.indexOf("Estado Liquidacion");
  var colFechaLiq = encabezados.indexOf("Fecha de Liquidacion");
  var colALiquidar = encabezados.indexOf("A Liquidar");
  var colUsdt = encabezados.indexOf("USDT"); // Ahora busca la columna "USDT"

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
      // Extraemos el valor de Euros y lo limpiamos
      var montoEur = parseFloat(String(datos[i][colALiquidar]).replace("€", "").replace(",", "."));
      
      if (!isNaN(montoEur)) {
        // AHORA DIVIDIMOS POR LA TASA DE KRAKEN
        var montoUsdt = montoEur / tipoDeCambio;
        
        // Escribimos en la columna USDT
        hojaTransacciones.getRange(i + 1, colUsdt + 1).setValue(montoUsdt);
        filasActualizadas++;
      }
    }
  }

  ui.alert("✅ Tasa de " + tipoDeCambio + " EUR/USDT guardada.\n\nSe calcularon los USDT exactos (dividiendo por la tasa) para " + filasActualizadas + " pagos pendientes.");
}

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

  // --- PREGUNTAS FONDEO KRAKEN ---
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

  // 1. Recorremos y procesamos
  for (var i = 1; i < valores.length; i++) {
    var estadoActual = valores[i][colEstado];
    var fechaCelda = new Date(valores[i][colFechaLiq]);

    if (estadoActual === "Pendiente" && fechaCelda <= fechaHoy) {
      
      var idActual = String(valores[i][colID]).trim();
      
      if (fechaLiquidacionResumen === "") {
        // --- CORRECCIÓN 1: Ahora usa la fecha del momento de hacer click (fechaHoy) ---
        var dia = ("0" + fechaHoy.getDate()).slice(-2);
        var mes = ("0" + (fechaHoy.getMonth() + 1)).slice(-2);
        var anio = fechaHoy.getFullYear();
        fechaLiquidacionResumen = dia + "/" + mes + "/" + anio;
      }

      if (!resumenPorID[idActual]) {
        resumenPorID[idActual] = {
          bruto: 0,
          comision: 0,
          honorarios: 0,
          aLiquidar: 0,
          usdt: 0
        };
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

  // --- CORRECCIÓN 2: Búsqueda segura del Exchange Rate vinculada a la fecha ---
  var tipoDeCambio = "No especificado";
  if (hojaLiquidacion) {
    var colK = hojaLiquidacion.getRange("K:K").getValues();
    var colL = hojaLiquidacion.getRange("L:L").getValues();
    var indiceUltimaFila = -1;
    
    // Escaneamos de arriba a abajo para encontrar la última carga real
    for (var j = 0; j < colK.length; j++) {
      if (String(colK[j][0]).trim() !== "") {
        indiceUltimaFila = j;
      }
    }
    
    if (indiceUltimaFila !== -1) {
      tipoDeCambio = colL[indiceUltimaFila][0];
    }
  }

  // 3. Escribimos en el Historial de Pagos, Ordenamos y CREAMOS SALDOS DEUDORES
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
          fechaHoraActual,                 
          "SALDO DEUDOR SEMANA ANTERIOR",  
          "ARRASTRE",                      
          "",                              
          idL,                             
          datosID.aLiquidar,               
          0,                               
          "",                              
          "",                              
          "",                              
          fechaProxViernes,                
          "Pendiente"                      
        ];
        
        hojaTransacciones.appendRow(filaDeuda);
        var row = hojaTransacciones.getLastRow();
        
        hojaTransacciones.getRange(row, 8).setFormula("=F" + row + "-G" + row);
        hojaTransacciones.getRange(row, 9).setFormula('=IF(OR(UPPER(B' + row + ')="SALDO DEUDOR SEMANA ANTERIOR",UPPER(B' + row + ')="RESERVA ADS",UPPER(B' + row + ')="REINTEGRO ADS"),0,ABS(H' + row + ')*0.2)');
        hojaTransacciones.getRange(row, 10).setFormula("=H" + row + "-I" + row);

        deudasGeneradas++;
      }
    }

    var rangoAFiltrar = hojaHistorial.getRange(2, 1, hojaHistorial.getLastRow() - 1, 9);
    rangoAFiltrar.sort([
      {column: 2, ascending: true}, 
      {column: 1, ascending: false} 
    ]);
    
    // --- GUARDADO EN FONDEO KRAKEN ---
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
          fechaKraken,
          eurosKraken,
          usdtCompradosReales > 0 ? parseFloat(usdtCompradosReales.toFixed(2)) : "Falta Tipo de Cambio",
          tipoDeCambio,
          "Liq. " + fechaLiquidacionResumen
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

function generarExtractoPDF() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaTransacciones = ss.getSheetByName("Transacciones");
  var hojaLiquidacion = ss.getSheetByName("Semanal - Liquidacion");

  if (!hojaTransacciones) {
    ui.alert("❌ Error: No se encontró la pestaña 'Transacciones'.");
    return;
  }

  // 1. Pedimos los IDs (Múltiples o TODOS - Modificado para aceptar en blanco)
  var idResponse = ui.prompt(
    "Generar Extracto(s) 📄", 
    "Ingresá los IDs separados por coma (ej: 42, 21, 63)\n\n👉 Dejá el texto VACÍO y dale a OK para generar TODOS automáticamente:", 
    ui.ButtonSet.OK_CANCEL
  );
  
  if (idResponse.getSelectedButton() !== ui.Button.OK) return;
  var inputIDs = idResponse.getResponseText().trim().toUpperCase();
  if (inputIDs === "") inputIDs = "TODOS"; // <--- MAGIA AQUI

  // 2. Pedimos la Fecha de Liquidación 
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

  // 3. Buscar el Tipo de Cambio en el Historial
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
  var carpetaDestino;
  if (carpetas.hasNext()) {
    carpetaDestino = carpetas.next();
  } else {
    carpetaDestino = DriveApp.createFolder("Extractos Semanales");
  }

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
      var archivoViejo = archivosExistentes.next(); 
      archivoViejo.setTrashed(true);
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
function enviarExtractosPorEmail() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaLiquidacion = ss.getSheetByName("Semanal - Liquidacion");

  if (!hojaLiquidacion) {
    ui.alert("❌ Error: No se encontró la pestaña 'Semanal - Liquidacion' para buscar los emails.");
    return;
  }

  // 1. Pedimos los IDs (Múltiples o TODOS - Modificado para aceptar en blanco)
  var idResponse = ui.prompt(
    "Enviar Extracto(s) ✉️", 
    "Ingresá los IDs separados por coma (ej: 42, 21, 63)\n\n👉 Dejá el texto VACÍO y dale a OK para enviar TODOS automáticamente:", 
    ui.ButtonSet.OK_CANCEL
  );
  if (idResponse.getSelectedButton() !== ui.Button.OK) return;
  var inputIDs = idResponse.getResponseText().trim().toUpperCase();
  if (inputIDs === "") inputIDs = "TODOS"; // <--- MAGIA AQUI

  // 2. Pedimos la Fecha 
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

  // 3. Mapear Emails y Filas (Columnas O y Q)
  var datosLiquidacion = hojaLiquidacion.getDataRange().getValues();
  var idDataMap = {}; 

  for (var i = 0; i < datosLiquidacion.length; i++) {
    var idFila = String(datosLiquidacion[i][14]).trim(); // Columna O (Índice 14)
    var emailFila = String(datosLiquidacion[i][16]).trim(); // Columna Q (Índice 16) - ACTUALIZADO
    
    if (idFila !== "" && idFila !== "ID" && emailFila.indexOf("@") !== -1) {
      idDataMap[idFila] = {
        email: emailFila,
        filaExcel: i + 1 
      };
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

        // --- NUEVO: Anotar la fecha en la Columna R (Último Extracto) ---
        // Se cambió el índice 17 (Columna Q) al 18 (Columna R) para no pisar el email
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
// ==========================================
// 1. GENERAR FACTURAS MENSUALES
// ==========================================
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

  // Inputs del usuario
  var idResponse = ui.prompt("Generar Facturas 🧾", "Ingresá IDs separados por coma (ej: 42, 21)\n👉 Dejá VACÍO para generar TODOS:", ui.ButtonSet.OK_CANCEL);
  if (idResponse.getSelectedButton() !== ui.Button.OK) return;
  var inputIDs = idResponse.getResponseText().trim().toUpperCase() || "TODOS";

  var mesResponse = ui.prompt("Mes a Facturar 📅", "Ingresá el número del mes (ej: 3 para Marzo):", ui.ButtonSet.OK_CANCEL);
  if (mesResponse.getSelectedButton() !== ui.Button.OK) return;
  var mesFacturar = parseInt(mesResponse.getResponseText().trim());

  var anioActual = 2026; 

  // --- LOGO: base64 generado en runtime desde Drive ---
  var logoSrc = "";
  try {
    var logoFile = DriveApp.getFileById("1w1iBy9qf87o15uWlVP2z7tf53nLEkuaW");
    var logoBase64 = Utilities.base64Encode(logoFile.getBlob().getBytes());
    logoSrc = "data:image/png;base64," + logoBase64;
  } catch(e) {
    // Si falla, el logo no se muestra pero la factura se genera igual
  }

  // Mapear Datos de Creadoras
  var datosLiq = hojaLiquidacion.getDataRange().getValues();
  var datosCreadoras = {};
  for (var i = 1; i < datosLiq.length; i++) {
    var idFila = String(datosLiq[i][14]).trim();
    if (idFila !== "" && idFila !== "ID") {
      datosCreadoras[idFila] = {
        nombre: String(datosLiq[i][18] || "Sin Nombre"), 
        cuit: String(datosLiq[i][19] || "Sin CUIT")      
      };
    }
  }

  // Sumar Honorarios
  var transDatos = hojaTransacciones.getDataRange().getValues();
  var encT = transDatos[0];
  var colEstado = encT.indexOf("Estado Liquidacion");
  var colFechaTx = encT.indexOf("Fecha y Hora");
  var colID = encT.indexOf("ID Proveedor");
  var colHonorarios = encT.indexOf("Honorarios");

  var facturacionPorID = {};

  for (var t = 1; t < transDatos.length; t++) {
    if (transDatos[t][colEstado] === "Procesado" || transDatos[t][colEstado] === "Pendiente") {
      var fecha = parseFechaTalaria(transDatos[t][colFechaTx]);
      if (fecha && fecha.getMonth() + 1 === mesFacturar && fecha.getFullYear() === anioActual) {
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
  var sobreescritas = 0;

  for (var k = 0; k < idsAProcesar.length; k++) {
    var idActual = idsAProcesar[k];
    
    // Ignorar si no hay honorarios o si el ID no existe en la pestaña de liquidación
    if (!facturacionPorID[idActual] || facturacionPorID[idActual] <= 0) continue;
    if (!datosCreadoras[idActual]) continue; 

    var totalHonorarios = facturacionPorID[idActual];
    var dataCliente = datosCreadoras[idActual];

    // --- LÓGICA DE SOBREESCRITURA INTELIGENTE ---
    var searchStr = "_ID" + idActual + "_" + mesFacturar + "-" + anioActual + ".pdf";
    var archivosViejos = carpetaDestino.searchFiles("title contains '" + searchStr + "'");
    
    var esSobreescritura = false;
    var numeroFactura = "";

    if (archivosViejos.hasNext()) {
      var archivoViejo = archivosViejos.next();
      var nombreViejo = archivoViejo.getName();
      var partes = nombreViejo.split("_");
      if (partes.length > 1) {
        numeroFactura = partes[1]; // Recupera el número original
        esSobreescritura = true;
      }
      archivoViejo.setTrashed(true); // Envía el viejo a la papelera
      
      // Limpia duplicados si existen
      while (archivosViejos.hasNext()) {
        archivosViejos.next().setTrashed(true);
      }
    }

    if (!esSobreescritura) {
      var ultimaFilaReg = hojaRegistro.getLastRow();
      var numCorrelativo = ultimaFilaReg === 0 ? 1 : ultimaFilaReg; 
      numeroFactura = anioActual + "-" + String(numCorrelativo).padStart(3, '0');
    }

    var fechaEmision = new Date().toLocaleDateString('es-ES');

    // HTML de la Factura
    var html = "<div style='font-family: Arial, sans-serif; color: #333; max-width: 800px; margin: auto;'>";
    html += "<table width='100%'><tr>";
    
    // Logo desde Drive (base64 en runtime)
    if (logoSrc !== "") {
      html += "<td width='50%'><img src='" + logoSrc + "' width='150'></td>";
    } else {
      html += "<td width='50%'><h2>Talaria Digital</h2></td>";
    }
    
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
    var archivoFactura = carpetaDestino.createFile(blob);
    var urlFactura = archivoFactura.getUrl();

    // Agregar header de Link si la hoja está vacía o no tiene col F
    if (hojaRegistro.getLastRow() === 0) {
      hojaRegistro.appendRow(["N° Factura", "Fecha", "ID", "Cliente", "Total", "Link Factura"]);
    } else if (hojaRegistro.getRange(1, 6).getValue() === "") {
      hojaRegistro.getRange(1, 6).setValue("Link Factura");
    }

    // Actualizar el Registro de Facturas
    if (esSobreescritura) {
      var datosReg = hojaRegistro.getRange("A:A").getValues();
      var filaEncontrada = -1;
      for (var r = 0; r < datosReg.length; r++) {
        if (datosReg[r][0] === numeroFactura) {
          filaEncontrada = r + 1;
          break;
        }
      }
      if (filaEncontrada !== -1) {
        hojaRegistro.getRange(filaEncontrada, 2, 1, 5).setValues([[fechaEmision, idActual, dataCliente.nombre, totalHonorarios.toFixed(2), urlFactura]]);
      } else {
        hojaRegistro.appendRow([numeroFactura, fechaEmision, idActual, dataCliente.nombre, totalHonorarios.toFixed(2), urlFactura]);
      }
      sobreescritas++;
    } else {
      hojaRegistro.appendRow([numeroFactura, fechaEmision, idActual, dataCliente.nombre, totalHonorarios.toFixed(2), urlFactura]);
      generadas++;
    }

    // --- LIBRO FACTURAS EXPEDIDAS (Hacienda) ---
    var hojaLibro = ss.getSheetByName("Libro Facturas Expedidas");
    if (!hojaLibro) {
      hojaLibro = ss.insertSheet("Libro Facturas Expedidas");
      hojaLibro.appendRow(["Nº Factura", "Fecha Emisión", "Destinatario", "NIF/CUIT", "Concepto", "Base Imponible", "% IVA", "Cuota IVA", "Total", "Observaciones"]);
    }

    var filaLibro = [
      numeroFactura,
      fechaEmision,
      dataCliente.nombre,
      dataCliente.cuit,
      "Servicios de gestión de agencia digital y marketing - Mes " + ("0" + mesFacturar).slice(-2) + "/" + anioActual,
      parseFloat(totalHonorarios.toFixed(2)),
      "0%",
      0,
      parseFloat(totalHonorarios.toFixed(2)),
      "No sujeta a IVA — Art. 69 Ley 37/1992"
    ];

    var colA_Libro = hojaLibro.getRange("A:A").getValues();
    var filaLibroExistente = -1;
    for (var rl = 0; rl < colA_Libro.length; rl++) {
      if (String(colA_Libro[rl][0]) === String(numeroFactura)) { filaLibroExistente = rl + 1; break; }
    }
    if (filaLibroExistente !== -1) {
      hojaLibro.getRange(filaLibroExistente, 1, 1, filaLibro.length).setValues([filaLibro]);
    } else {
      hojaLibro.appendRow(filaLibro);
    }
  }

  var msjResultado = "";
  if (generadas > 0) msjResultado += "✅ Se generaron " + generadas + " facturas nuevas.\n";
  if (sobreescritas > 0) msjResultado += "🔄 Se actualizaron y sobreescribieron " + sobreescritas + " facturas existentes.\n";
  if (generadas === 0 && sobreescritas === 0) msjResultado = "ℹ️ No se encontraron honorarios para facturar en ese mes/IDs.";

  ui.alert("Resumen de Facturación:\n\n" + msjResultado);
}

// ==========================================
// 2. ENVIAR FACTURAS
// ==========================================
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
    var email = String(datosLiq[i][16]).trim(); // Columna Q
    if (idFila !== "" && email.indexOf("@") !== -1) correos[idFila] = email;
  }

  var carpetas = DriveApp.getFoldersByName("Facturas");
  if (!carpetas.hasNext()) { ui.alert("❌ No existe la carpeta 'Facturas'."); return; }
  var carpetaFacturas = carpetas.next();
  var archivos = carpetaFacturas.getFiles();
  
  var idsAProcesar = inputIDs === "TODOS" ? Object.keys(correos) : inputIDs.replace(/;/g, ",").split(",").map(function(i){return i.trim()});
  var enviados = 0;

  for (var k = 0; k < idsAProcesar.length; k++) {
    var id = idsAProcesar[k];
    if (!correos[id]) continue;

    // Buscar archivo que coincida con ID y Mes
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

// ==========================================
// AYUDA / GUÍA DE USO
// ==========================================
function mostrarAyuda() {
  var ui = SpreadsheetApp.getUi();

  var guia =
    "⚡ FLUJO SEMANAL (ejecutar en orden cada viernes)\n" +
    "─────────────────────────────────────────────\n\n" +
    "  1. 💱 CARGAR TIPO DE CAMBIO\n" +
    "     Abrí Kraken → buscá EUR/USDT → ingresá el precio en EUR\n" +
    "     de 1 USDT (ej: 0.8626, NO al revés).\n" +
    "     Hacelo ANTES de cerrar la semana.\n\n" +
    "  2. 🔒 ACTUALIZAR LIQUIDACIONES\n" +
    "     Cierra la semana: pasa a 'Procesado' todo lo 'Pendiente'\n" +
    "     con fecha ≤ hoy. Te pregunta cuántos euros enviaste a Kraken\n" +
    "     (poné 0 si no enviaste nada).\n" +
    "     ⚠️  Siempre DESPUÉS del Tipo de Cambio.\n\n" +
    "  3. 📄 GENERAR EXTRACTOS PDF\n" +
    "     Dejá el campo de IDs VACÍO + OK → genera TODOS.\n" +
    "     Dejá la fecha VACÍA → usa la de hoy automáticamente.\n" +
    "     Los PDFs quedan en Drive > carpeta 'Extractos Semanales'.\n\n" +
    "  4. ✉️  ENVIAR EXTRACTOS POR EMAIL\n" +
    "     Igual que el anterior: VACÍO + OK = envía a todas.\n" +
    "     Los emails se toman de la columna Q de 'Semanal - Liquidacion'.\n\n\n" +
    "💸 CARGAR GASTO (cuando ocurra durante la semana)\n" +
    "─────────────────────────────────────────────\n" +
    "  Ingresá: ID de creadora → monto en EUROS (positivo) →\n" +
    "  concepto → referencia o número de factura.\n" +
    "  El sistema lo registra en negativo y calcula el 20% de honorarios.\n\n\n" +
    "🧾 FACTURAS MENSUALES (a fin de cada mes)\n" +
    "─────────────────────────────────────────────\n" +
    "  1. Generar Facturas → ingresá el número de mes (ej: 4 = Abril).\n" +
    "  2. Enviar Facturas  → mismo número de mes.\n" +
    "  Si regenerás una factura ya existente, la sobreescribe\n" +
    "  manteniendo el mismo número de factura.\n\n\n" +
    "📌 DATOS CLAVE PARA NO EQUIVOCARSE\n" +
    "─────────────────────────────────────────────\n" +
    "  • IDs de creadoras : números simples (ej: 21, 42, 63)\n" +
    "  • Fechas           : formato DD/MM/AAAA\n" +
    "  • Tasa Kraken      : EUR por 1 USDT (ej: 0.86), no al revés\n" +
    "  • TODOS            : dejar el campo VACÍO y darle OK\n" +
    "  • Saldo deudor     : si una creadora termina en negativo,\n" +
    "                       se arrastra automáticamente a la semana siguiente";

  ui.alert("📋 Guía de Uso — Talaria Digital", guia, ui.ButtonSet.OK);
}