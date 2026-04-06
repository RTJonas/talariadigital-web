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
  
  // Col I: Honorarios = ABS(Neto) * 0.2 (Siempre te da positivo el 20%)
  sheet.getRange(row, 9).setFormula("=ABS(H" + row + ")*0.2");
  
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
        hojaTransacciones.getRange(row, 9).setFormula('=IF(B' + row + '="SALDO DEUDOR SEMANA ANTERIOR", 0, ABS(H' + row + ')*0.2)');
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

  // --- SOLUCIÓN LOGO: Base64 Hardcodeado ---
  // Reemplazá el texto de abajo con el código de tu imagen. No borres las comillas.
  var codigoBase64Puro = "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";
  
  var base64Logo = "";
  if (codigoBase64Puro !== "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" && codigoBase64Puro !== "") {
    base64Logo = "data:image/png;base64," + codigoBase64Puro;
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
    
    // Inyecta el logo si el Base64 está configurado
    if (base64Logo !== "") {
      html += "<td width='50%'><img src='" + base64Logo + "' width='150'></td>";
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
    carpetaDestino.createFile(blob);

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
        hojaRegistro.getRange(filaEncontrada, 2, 1, 4).setValues([[fechaEmision, idActual, dataCliente.nombre, totalHonorarios.toFixed(2)]]);
      } else {
        hojaRegistro.appendRow([numeroFactura, fechaEmision, idActual, dataCliente.nombre, totalHonorarios.toFixed(2)]);
      }
      sobreescritas++;
    } else {
      hojaRegistro.appendRow([numeroFactura, fechaEmision, idActual, dataCliente.nombre, totalHonorarios.toFixed(2)]);
      generadas++;
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