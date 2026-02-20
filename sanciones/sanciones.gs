function onOpen() {
  createDropdown()//ACTUALIZA LA LISTA DESPLEGABLE DE Nº EXPEDIENTE
  colorearCeldasPliegos();
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Sanciones')
    //.addItem('Aplicar sanciones', 'probarAccesoFormulario')
    .addItem('Historial del denunciante', 'buscarHistorialDenunciante')
    .addItem('Actualizar Pliegos', 'colorearCeldasPliegos')
    .addItem("Formulario de sanciones", "abrirFormulario")
    .addItem('Formulario de pliegos', 'abrirFormularioPliegoConCopia')
    .addToUi();
    var ui = SpreadsheetApp.getUi();
  ui.createMenu('Resolución')
    .addItem('Sanción Nula', 'checkAndSetNula')
    .addItem('Introducir seguimiento', 'promptAndAppendToA8AndDirectiva')
    .addItem('Introducir Beredicto', 'promptAndAppendBeredicto')
    .addItem('Historial del denunciante', 'mostrarResolucion')
    
    //.addItem('Falta muy Grave', 'FaltaMuyGrave')
    .addItem('Exportar pago', 'exportarDatosPago')
    .addItem('ejecutar expediente', 'updateResolutionFields')
    .addItem('Importar Estado del pago de sanciones', 'actualizarEstadoDePagoDeSanciones')
    .addItem('🖨️Imprimir copia sanción en firme', 'imprimirCopiaSancionEnFirme')
    //.addItem('Impromir informe Sancionado', 'imprimirInformeSancionado')
    //.addItem('Impromir sancion en firme', 'imprimirInformeSancionado')
    .addToUi(); 
  ui.createMenu('Mayusculas')
    //.addItem('Aplicar sanciones', 'probarAccesoFormulario')
    .addItem('🔠Convertir a Mayusculas las celdas seleccionadas', 'convertirCeldaSeleccionadaAMayusculas')
    .addItem('🔤Convertir a Nombre propio en todas las celdas seleccionadas', 'convertirCeldaSeleccionadaANombrePropio')
    .addItem('🔡Convertir a minusculas en todas las celdas seleccionadas', 'convertirCeldaSeleccionadaAMinusculas')
    .addItem('🔁Convertir espacios en blanco en todas las celdas seleccionadas', 'limpiarEspaciosEnBlanco')
    .addItem('📞Eliminar espacios en blanco en todos los teléfonos', 'eliminarEspaciosEnBlanco')
    .addToUi();
crearBarraLateral();
}


function onFormSubmit(e) {

  sanciones();
  limpiarLmSanciones();
  
}


// ++++++++++VARIABLES GLOBALES +++++++++++++++++++++++//
var sheetReg = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registro entrada sanciones");
var sheetPli = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registro pliegos");
var sheetTes = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Testigos");
var sheetRes = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Resolución");
var sheetSocios = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Socios");
// ++++++++++ FIN - VARIABLES GLOBALES +++++++++++++++++++++++//


// ++++++++++ APLICAR BORDES A LA HOJA Reg +++++++++++++++//
function sanciones() {
  // Aplicar bordes a todas las celdas con datos en la hoja "Registro entrada sanciones"
  applyBordersToCells(sheetReg);

  // Ordenar la columna A de mayor a menor
  sortColumnADescending(sheetReg);

  // Actualizar columna M según la condición especificada
  updateColumnM(sheetReg);
  separador();
}
// ++++++++++ FIN - APLICAR BORDES A LA HOJA Reg +++++++++++++++//


// ++++++++++ COLOREAR DE AZUL LOS BORDES DE LA HOJA Reg +++++++//
function applyBordersToCells(sheet) {
  var range = sheet.getDataRange();
  range.setBorder(true, true, true, true, true, true, "blue", SpreadsheetApp.BorderStyle.SOLID);
}
// ++++++++++ COLOREAR DE AZUL LOS BORDES DE LA HOJA Reg +++++++//


// ++++++++++ ORDENAR LA COLUMNA A DE LA HOJA Reg ++++++++++++++//
function sortColumnADescending(sheet) {
  var range = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  range.sort({column: 1, ascending: false});
}
// ++++++++++ FIN - ORDENAR LA COLUMNA A DE LA HOJA Reg ++++++++//


//++++++++++ NUMERAR EXPEDIENTES +++++++++++++++++++++++++++++++//
function updateColumnM(sheet) {
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();

  for (var i = 1; i < data.length; i++) { // Empezamos desde 1 para saltar la fila del encabezado
    if (!data[i][12]) { // Columna M es el índice 12
      if (i < data.length - 1) {
        data[i][12] = data[i + 1][12] + 1;
      } else {
        data[i][12] = 1; // Si es la última fila y está vacía, asignar 1
      }
    }
  }
  dataRange.setValues(data);
  actualizarContenidoCelda();
}

function actualizarContenidoCelda() { // Esta función numera el número de expediente sumando 1 a la celda inferior
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var celda = hoja.getRange('M2');
  var valor = celda.getValue();
  
  // Encuentra la posición del símbolo "/"
  var indice = valor.indexOf('/');
  
  if (indice !== -1) {
    // Obtiene el número antes del símbolo "/"
    var numero = valor.substring(0, indice);
    
    // Intenta convertir el número a un valor entero
    var numeroActual = parseInt(numero, 10);
    
    if (!isNaN(numeroActual)) {
      // Suma 1 al número
      var nuevoNumero = numeroActual + 1;
      
      // Obtiene el año actual
      var fechaActual = new Date();
      var anoActual = fechaActual.getFullYear();
      
      // Actualiza el contenido de la celda
      celda.setValue(nuevoNumero + '/' + anoActual);
    } else {
      // Manejo en caso de que el valor antes del símbolo no sea un número
      Logger.log('El valor antes del símbolo "/" no es un número.');
    }
  } else {
    // Manejo en caso de que el símbolo "/" no se encuentre
    Logger.log('El símbolo "/" no se encontró en la celda.');
  }
  procesarCeldaH2();
}
function procesarCeldaH2() { // ESTA FUNCIÓN DIVIDE SANCIONES MÚLTIPLES
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var celdaH2 = hoja.getRange('H2').getValue();
  var celdaM2 = hoja.getRange('M2').getValue();
  
  // Asegúrate de que celdaH2 sea una cadena
  celdaH2 = String(celdaH2);
  
  while (celdaH2.includes(',')) {
    // Contar el número de comas
    var numComas = (celdaH2.match(/,/g) || []).length;
    
    // Insertar una fila debajo de la fila 2
    hoja.insertRowAfter(2);
    
    // Copiar el contenido de la fila 2 a la nueva fila 3
    hoja.getRange('2:2').copyTo(hoja.getRange('3:3'));
    
    // Obtener el último item a la derecha de la celda H2 después de la última coma
    var partes = celdaH2.split(',');
    var ultimoItem = partes.pop();
    
    // Establecer el valor en la celda H3
    hoja.getRange('H3').setValue(ultimoItem);
    
    // Establecer el valor en la celda M3
    hoja.getRange('M3').setValue(celdaM2 + '-' + numComas);
    
    // Eliminar el último item y la última coma de la celda H2
    celdaH2 = partes.join(',');
    
    // Actualizar la celda H2 con el nuevo contenido
    hoja.getRange('H2').setValue(celdaH2);
  }
  
  // Mensaje de finalización
  // SpreadsheetApp.getUi().alert('Proceso completado.');
  //exportarSanciones(); //modificado 21/05-------------------------------------
}





//++++++++++ FIN - NUMERAR EXPEDIENTES +++++++++++++++++++++++++++++++//


//*********************************************************************************************************************//
//++++++++++++++++++ACTUALIZA LA LISTA DESPLEGABLE DE Nº EXPEDIENTE++++++++++++++++++++++++++++++++++++++++++++++++++++//
//*********************************************************************************************************************//

/*function onEdit(e) {    //---Registro entrada sanciones//          ****** Cambio realizado por Kapi ****
  // Definir la hoja y el rango donde se realizará la edición
  var sheet = e.source.getActiveSheet();
  var range = e.range;

  // Verificar si el cambio se realizó en la columna N de la hoja "Registro entrada sanciones"
  if (sheet.getName() === "Registro entrada sanciones" && range.getColumn() === 14) {
    // Llamar a la función que crea la lista desplegable
    createDropdown();
  }
}

function createDropdown() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var registroSheet = ss.getSheetByName("Registro entrada sanciones");
  var resolucionSheet = ss.getSheetByName("Resolución");

  // Obtener datos de la columna M donde la celda offset(0,1) es ""
  var data = registroSheet.getRange("M2:M" + registroSheet.getLastRow()).getValues();
  var validData = [];

  for (var i = 0; i < data.length; i++) {
    if (registroSheet.getRange(i + 2, 14).getValue() === "") { // columna N está vacía
      validData.push(data[i][0]);
    }
  }

  // Crear la validación de datos con los valores obtenidos
  var cell = resolucionSheet.getRange("F3");
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(validData).build();
  cell.setDataValidation(rule);
}

//*********************************************************************************************************************//
//++++++++++++++++++RELLENAR INFORME ENCABEZADO Y CONTENIDO++++++++++++++++++++++++++++++++++++++++++++++++++++//
//*********************************************************************************************************************//

function onEdit(e) {
  //------Evento al cambiar número expediente//
  var sheet = e.source.getActiveSheet();
  var range = e.range;

  // Verificar si el cambio se realizó en la celda F3 de la hoja "Resolución"
  if (sheet.getName() === "Resolución" && range.getA1Notation() === "F3") {
    updateResolutionFields();
  }

  //-----Faltas, leves, graves y muy graves//
  var sheetResolucion = e.source.getSheetByName("Resolución");

  // Verificar si la celda editada es C136
  if (sheetResolucion.getName() === "Resolución" && range.getA1Notation() === "C136") {
    // Cambiar el color de la celda C136 a blanco
    range.setBackground("white");

    // Limpiar el contenido de las celdas D136 y E136
    sheetResolucion.getRange("D136").setValue("");
    sheetResolucion.getRange("E136").setValue("");
    // Asegúrate de definir la celda "e111" correctamente
    // por ejemplo: sheetResolucion.setActiveRange(sheetResolucion.getRange("E111"));
  }

  //--- Buscar y actualizar en la hoja "Socios" cuando cambie B117 ---
  // Verificar si el cambio ocurrió en la celda B117 de la hoja Resolución
  if (sheetResolucion.getName() === "Resolución" && range.getA1Notation() === "B117") {
    // Obtener el valor de la celda B117
    var nuevoValorB117 = range.getValue();
    
    // Obtener el valor de la celda A11 de la hoja Resolución
    var valorA11 = sheetResolucion.getRange("A11").getValue();
    
    // Obtener la hoja Socios
    var sheetSocios = e.source.getSheetByName("Socios");
    
    // Obtener el rango A2:A de la hoja Socios
    var rangoSocios = sheetSocios.getRange("A2:A" + sheetSocios.getLastRow());
    var valoresSocios = rangoSocios.getValues();
    
    // Iterar sobre cada fila de la columna A en la hoja Socios
    for (var i = 0; i < valoresSocios.length; i++) {
      // Si encontramos un valor coincidente
      if (valoresSocios[i][0] === valorA11) {
        // Establecer el valor de la columna I en la misma fila
        sheetSocios.getRange(i + 2, 9).setValue(nuevoValorB117);
        break; // Detener la búsqueda una vez encontrado
      }
    }
  }
  
  //--- Actualizar B147 y A139 cuando cambie E111 ---
  if (sheetResolucion.getName() === "Resolución" && range.getA1Notation() === "E111") {
    range.setBackground("white");
    
    // Obtener la fecha actual
    var fechaActual = new Date();

    // Formatear la fecha actual
    var opcionesFecha = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
    var fechaFormateada = "En Los Cristianos a " + fechaActual.toLocaleDateString('es-ES', opcionesFecha);

    // Actualizar la celda B147 con la fecha actual formateada
    sheetResolucion.getRange("B147").setValue(fechaFormateada);

    // Obtener el valor de la celda E111 (fecha objetivo)
    var fechaObjetivo = new Date(sheetResolucion.getRange("E111").getValue());

    // Obtener el número de días de la celda I1
    var diasAdicionales = parseInt(sheetResolucion.getRange("I1").getValue(), 10);

    // Calcular la nueva fecha objetivo sumando los días adicionales
    fechaObjetivo.setDate(fechaObjetivo.getDate() + diasAdicionales);

    // Formatear la nueva fecha objetivo
    var fechaObjetivoFormateada = fechaObjetivo.toLocaleDateString('es-ES', opcionesFecha);

    // Actualizar la celda A139 con el texto y la fecha calculada
    var textoA139 = "Informándole que Usted tiene de plazo hasta el día " + fechaObjetivoFormateada + " a las 12:00 horas para hacer efectiva la sanción en la cuenta de esta asociación";
    sheetResolucion.getRange("A139").setValue(textoA139);
    
    //--- Buscar y actualizar en la hoja "Pago Sanciones" cuando cambie E111 ---
    // Obtener la hoja Pago Sanciones
    var sheetPagoSanciones = e.source.getSheetByName("Pago Sanciones");
    
    // Obtener el valor de F109 de la hoja Resolución
    var valorF109 = sheetResolucion.getRange("F109").getValue();
    
    // Obtener los valores de la hoja Pago Sanciones columna A
    var rangoPagoSanciones = sheetPagoSanciones.getRange("A2:A" + sheetPagoSanciones.getLastRow());
    var valoresPagoSanciones = rangoPagoSanciones.getValues();

    // Bandera para verificar si el valor fue encontrado
    var encontrado = false;

    // Buscar en la columna A de Pago Sanciones
    for (var j = 0; j < valoresPagoSanciones.length; j++) {
      if (valoresPagoSanciones[j][0] === valorF109) {
        // Si se encuentra el valor, actualizar las columnas correspondientes
        sheetPagoSanciones.getRange(j + 2, 2).setValue(sheetResolucion.getRange("A117").getValue()); // Columna B
        sheetPagoSanciones.getRange(j + 2, 3).setValue(sheetResolucion.getRange("A120").getValue()); // Columna C
        sheetPagoSanciones.getRange(j + 2, 4).setValue(sheetResolucion.getRange("A123").getValue()); // Columna D
        sheetPagoSanciones.getRange(j + 2, 5).setValue(sheetResolucion.getRange("C136").getValue()); // Columna E
        // Calcular la fecha en Columna F
        var fechaE111 = new Date(sheetResolucion.getRange("E111").getValue());
        fechaE111.setDate(fechaE111.getDate() + diasAdicionales);
        sheetPagoSanciones.getRange(j + 2, 6).setValue(fechaE111.toLocaleDateString('es-ES', opcionesFecha)); // Columna F
        
        encontrado = true;
        break;
      }
    }
    
    // Si no se encuentra el valor, insertar una nueva fila y agregar los valores
    if (!encontrado) {
      // Insertar una nueva fila en la posición 2
      sheetPagoSanciones.insertRowAfter(1);
      
      // Asignar los valores a la nueva fila
      sheetPagoSanciones.getRange(2, 1).setValue(valorF109); // Columna A
      sheetPagoSanciones.getRange(2, 2).setValue(sheetResolucion.getRange("A117").getValue()); // Columna B
      sheetPagoSanciones.getRange(2, 3).setValue(sheetResolucion.getRange("A120").getValue()); // Columna C
      sheetPagoSanciones.getRange(2, 4).setValue(sheetResolucion.getRange("A123").getValue()); // Columna D
      sheetPagoSanciones.getRange(2, 5).setValue(sheetResolucion.getRange("C136").getValue()); // Columna E
      // Calcular la fecha en Columna F
      var nuevaFecha = new Date(sheetResolucion.getRange("E111").getValue());
      nuevaFecha.setDate(nuevaFecha.getDate() + diasAdicionales);
      sheetPagoSanciones.getRange(2, 6).setValue(nuevaFecha.toLocaleDateString('es-ES', opcionesFecha)); // Columna F
    }
    
    //--- Sumar las celdas E2:E y actualizar G1 ---
    var rangoE = sheetPagoSanciones.getRange("E2:E" + sheetPagoSanciones.getLastRow());
    var valoresE = rangoE.getValues();
    var sumaE = 0;

    for (var k = 0; k < valoresE.length; k++) {
      sumaE += parseFloat(valoresE[k][0]) || 0; // Sumar solo números
    }

    // Actualizar la celda G1 con la suma de la columna E
    sheetPagoSanciones.getRange("G1").setValue(sumaE);

    //--- Cambiar color de fechas en F2:F ---
    var rangoF = sheetPagoSanciones.getRange("F2:F" + sheetPagoSanciones.getLastRow());
    var valoresF = rangoF.getValues();

    for (var m = 0; m < valoresF.length; m++) {
      var fechaColumnaF = new Date(valoresF[m][0]);
      if (fechaColumnaF > fechaActual) {
        // Si la fecha es mayor que la actual, color rojo
        sheetPagoSanciones.getRange(m + 2, 6).setFontColor("green");
      } else {
        // Si la fecha es menor o igual que la actual, color verde
        sheetPagoSanciones.getRange(m + 2, 6).setFontColor("red");
      }
    }
  }
  buscarHistorialDenunciante();
}




//+++++++++++++++++++FIN ON EDIT(E)+++++++++++++++++++++++++++++//


function updateResolutionFields() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var resolucionSheet = ss.getSheetByName("Resolución");
  var registroSheet = ss.getSheetByName("Registro entrada sanciones");
  var sociosSheet= ss.getSheetByName("Socios");

  var f3Value = resolucionSheet.getRange("F3").getValue();

  // Obtener datos de la hoja "Registro entrada sanciones"
  var data = registroSheet.getRange("A2:R" + registroSheet.getLastRow()).getValues();


  for (var i = 0; i < data.length; i++) {
    if (data[i][12] == f3Value) { // Columna M corresponde al índice 12 en el array data
        var unidadSancionadora = data[i][5]; // Columna F
        var unidadSancionada = data[i][7]; // Columna H
        var testigos = data[i][11]; // Columna L
        var denunciante = data[i][2].toUpperCase(); // Columna C
        var dni = data[i][3]; // Columna D
        var telefono = data[i][4]; // Columna E
        
        // Convertir FechaDeLosHechos a un objeto Date
        var FechaDeLosHechos = new Date(data[i][8]); // Columna I
        
        // Asumiendo que HoraDeLosHechos es un string o un objeto Date
        var HoraDeLosHechos = new Date(data[i][9]); // Columna J
        
        // Formatear la fecha a DD/MMMM/AAAA
        var opcionesFecha = { day: '2-digit', month: 'long', year: 'numeric' };
        var fechaFormateada = FechaDeLosHechos.toLocaleDateString('es-ES', opcionesFecha);
        
        // Formatear la hora a HH:mm
        var opcionesHora = { hour: '2-digit', minute: '2-digit', hour12: false };
        var horaFormateada = HoraDeLosHechos.toLocaleTimeString('es-ES', opcionesHora);

        var contenido = data[i][10]; // Columna K
        var FirmDte = data[i][16]; 
        var FirmDdo = data[i][17]; 

        resolucionSheet.getRange("E4").setValue("La unidad " + unidadSancionadora + " sanciona a la unidad " + unidadSancionada + ".");
        resolucionSheet.getRange("E5").setValue("Testigos: " + testigos);
        
        // Mostrar la fecha y hora formateada
        resolucionSheet.getRange("C6").setValue("Fecha y hora de los hechos: " + fechaFormateada + " a las " + horaFormateada);
        resolucionSheet.getRange("C60").setValue("Fecha y hora de los hechos: " + fechaFormateada + " a las " + horaFormateada);
        
        resolucionSheet.getRange("A8").setValue("La persona denunciante " + denunciante + ", Con D.N.I nº : " + dni + ", teléfono Nº : " + telefono + " expone que:\n" + contenido);
        resolucionSheet.getRange("F56").setValue(f3Value);
        resolucionSheet.getRange("A61").setValue("La persona denunciante expone que:\n\n" + contenido);
        resolucionSheet.getRange("C2").setValue("");
        resolucionSheet.getRange("C54").setValue("");
        //resolucionSheet.getRange("A51").setValue(FirmDte); // Firma denunciante
        //resolucionSheet.getRange("A52").setValue(FirmDdo); // Firma denunciado
        resolucionSheet.getRange("F109").setValue(f3Value);
        resolucionSheet.getRange("A117").setValue(unidadSancionada);
        
        break;
    }
}

  ocultarFilasReg();
  updateA8BasedOnF3();
  

}

//++++++++++++++++++++++ SANCIÓN NULA ++++++++++++++++++++++++++++++++++++++++++++++++++//

/*function checkAndSetNula() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var resolucionSheet = ss.getSheetByName("Resolución");
  var registroSheet = ss.getSheetByName("Registro entrada sanciones");

  // Obtener el valor de la celda F3 de la hoja Resolución
  var f3Value = resolucionSheet.getRange("F3").getValue();

  // Obtener todos los valores de la columna M de la hoja Registro entrada sanciones
  var mValues = registroSheet.getRange("M2:M" + registroSheet.getLastRow()).getValues();

  for (var i = 0; i < mValues.length; i++) {
    // Verificar si el valor en la columna M es igual al valor de F3
    if (mValues[i][0] == f3Value) {
      // Establecer el valor "NULA" en la celda adyacente en la columna N
      registroSheet.getRange(i + 2, 14).setValue("NULA");
    }
  }

createDropdown();

}*/
function checkAndSetNula() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var resolucionSheet = ss.getSheetByName("Resolución");
  var registroSheet = ss.getSheetByName("Registro entrada sanciones");

  // Obtener el valor de la celda F3 de la hoja Resolución
  var f3Value = resolucionSheet.getRange("F3").getValue();

  // Obtener todos los valores de la columna M de la hoja Registro entrada sanciones
  var mValues = registroSheet.getRange("M2:M" + registroSheet.getLastRow()).getValues();

  for (var i = 0; i < mValues.length; i++) {
    // Verificar si el valor en la columna M es igual al valor de F3
    if (mValues[i][0] == f3Value) {
      // Establecer el valor "NULA" en la celda adyacente en la columna N
      registroSheet.getRange(i + 2, 14).setValue("NULA");
    }
  }

  // Establecer el valor "SANCIÓN NULA" en la celda C2 de la hoja Resolución
  resolucionSheet.getRange("C2").setValue("SANCIÓN NULA");
  resolucionSheet.getRange("C54").setValue("SANCIÓN NULA");
  // Llamar a la función para crear el desplegable (si es necesario)
  createDropdown();
}


function checkAndSetLeve() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var resolucionSheet = ss.getSheetByName("Resolución");
  var registroSheet = ss.getSheetByName("Registro entrada sanciones");

  // Obtener el valor de la celda F3 de la hoja Resolución
  var f3Value = resolucionSheet.getRange("F3").getValue();

  // Obtener todos los valores de la columna M de la hoja Registro entrada sanciones
  var mValues = registroSheet.getRange("M2:M" + registroSheet.getLastRow()).getValues();

  for (var i = 0; i < mValues.length; i++) {
    // Verificar si el valor en la columna M es igual al valor de F3
    if (mValues[i][0] == f3Value) {
      // Establecer el valor "leve" en la celda adyacente en la columna N
      registroSheet.getRange(i + 2, 14).setValue("LEVE");
    }
  }

createDropdown();

}
function checkAndSetGrave() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var resolucionSheet = ss.getSheetByName("Resolución");
  var registroSheet = ss.getSheetByName("Registro entrada sanciones");

  // Obtener el valor de la celda F3 de la hoja Resolución
  var f3Value = resolucionSheet.getRange("F3").getValue();

  // Obtener todos los valores de la columna M de la hoja Registro entrada sanciones
  var mValues = registroSheet.getRange("M2:M" + registroSheet.getLastRow()).getValues();

  for (var i = 0; i < mValues.length; i++) {
    // Verificar si el valor en la columna M es igual al valor de F3
    if (mValues[i][0] == f3Value) {
      // Establecer el valor "leve" en la celda adyacente en la columna N
      registroSheet.getRange(i + 2, 14).setValue("GRAVE");
    }
  }

createDropdown();

}

function checkAndSetMuyGrave() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var resolucionSheet = ss.getSheetByName("Resolución");
  var registroSheet = ss.getSheetByName("Registro entrada sanciones");

  // Obtener el valor de la celda F3 de la hoja Resolución
  var f3Value = resolucionSheet.getRange("F3").getValue();

  // Obtener todos los valores de la columna M de la hoja Registro entrada sanciones
  var mValues = registroSheet.getRange("M2:M" + registroSheet.getLastRow()).getValues();

  for (var i = 0; i < mValues.length; i++) {
    // Verificar si el valor en la columna M es igual al valor de F3
    if (mValues[i][0] == f3Value) {
      // Establecer el valor "leve" en la celda adyacente en la columna N
      registroSheet.getRange(i + 2, 14).setValue("MUY GRAVE");
    }
  }

createDropdown();

}
 //++++++++++++++++++++++ AÑADIR PLIEGO AL INFORME +++++++++++++++++++++++++++//
function updateA8BasedOnF3() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var resolucionSheet = ss.getSheetByName("Resolución");
  var registroPliegosSheet = ss.getSheetByName("Registro pliegos");

  // Obtener el valor de la celda F3 de la hoja Resolución
  var f3Value = resolucionSheet.getRange("F3").getValue();

  // Obtener todos los valores de la columna C de la hoja Registro pliegos
  var cValues = registroPliegosSheet.getRange("C2:C" + registroPliegosSheet.getLastRow()).getValues();

  // Recorrer los valores de la columna C
  for (var i = 0; i < cValues.length; i++) {
    // Verificar si el valor en la columna C es igual al valor de F3
    if (cValues[i][0] == f3Value) {
      // Obtener el valor de la celda desplazada 6 columnas a la derecha (offset(0,6)) desde la columna C
      var ContenidoPliego = registroPliegosSheet.getRange(i + 2, 9).getValue(); // Columna I es la 9 (C + 6)
      var NombrePliego = registroPliegosSheet.getRange(i + 2, 10).getValue().toUpperCase(); 
      var DniPliego = registroPliegosSheet.getRange(i + 2, 4).getValue(); 
      var TelefonoPliego = registroPliegosSheet.getRange(i + 2, 4).getValue();      

      // Obtener el valor actual de la celda A8 de la hoja Resolución
      var currentA8Value = resolucionSheet.getRange("A8").getValue();
      var currentA61Value = resolucionSheet.getRange("A61").getValue();

      
      // Concatenar el valor actual de A8 con el valor desplazado
      var newA8Value = currentA8Value + "\n\nEl denunciado, " + NombrePliego + ", con D.N.I. nº: " + DniPliego + ", teléfono: "+ TelefonoPliego + ", expone que:\n" + ContenidoPliego;
      var newA61Value = currentA61Value + "\n\nEl denunciado expone que:\n" + ContenidoPliego;

      // Establecer el nuevo valor en la celda A8
      resolucionSheet.getRange("A8").setValue(newA8Value);
      resolucionSheet.getRange("A61").setValue(newA61Value);
      resolucionSheet.getRange("A123").setValue(NombrePliego);
      // Salir del bucle ya que solo necesitamos la primera coincidencia
      break;
    }
  }
  updateA8WithTestigos();
}
 //++++++++++++++++++++++++ AÑADIR TESTIGOS AL INFORME ++++++++++++++++++++++++++++++//

 function updateA8WithTestigos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var resolucionSheet = ss.getSheetByName("Resolución");
  var testigosSheet = ss.getSheetByName("Testigos");

  // Obtener el valor de la celda F3 de la hoja Resolución
  var f3Value = resolucionSheet.getRange("F3").getValue();

  // Obtener todos los valores de la columna C de la hoja Testigos
  var cValues = testigosSheet.getRange("C2:C" + testigosSheet.getLastRow()).getValues();

  // Inicializar una variable para acumular el texto de A8
  var newA8Value = resolucionSheet.getRange("A8").getValue();

  // Recorrer los valores de la columna C
  for (var i = 0; i < cValues.length; i++) {
    // Verificar si el valor en la columna C es igual al valor de F3
    if (cValues[i][0] == f3Value) {
      // Obtener el valor de la celda desplazada 5 columnas a la derecha (offset(0,5)) desde la columna C
      var LmTestigo = testigosSheet.getRange(i + 2, 8).getValue(); // Columna H es la 8 (C + 5)
      var DNITestigo = testigosSheet.getRange(i + 2, 5).getValue(); 
      var NombreTestigo = testigosSheet.getRange(i + 2, 4).getValue().toUpperCase(); 
      var TelefonoTestigo = testigosSheet.getRange(i + 2, 7).getValue(); 
      var ContenidoTestigo = testigosSheet.getRange(i + 2, 10).getValue(); 
    

      // Concatenar el nuevo texto con un nuevo párrafo
      newA8Value += "\n\nEl testigo, la unidad: " + LmTestigo + ", " + NombreTestigo + ", con D.N.I. nº: " + DNITestigo + ", teléfono: " + TelefonoTestigo + ", expone que: \n" + ContenidoTestigo;
    }
  }

  // Establecer el nuevo valor en la celda A8
  resolucionSheet.getRange("A8").setValue(newA8Value);
  RegistroSeguimientos();
}

//++++++++++++++++++++++++ AÑADIR REGISTRO SEGUIMIENTOS ++++++++++++++++++++++++++++++//

 function RegistroSeguimientos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var resolucionSheet = ss.getSheetByName("Resolución");
  var directivaSheet = ss.getSheetByName("Directiva");

  // Obtener el valor de la celda F3 de la hoja Resolución
  var f3Value = resolucionSheet.getRange("F3").getValue();

  // Obtener todos los valores de la columna A de la hoja Directiva
  var cValues = directivaSheet.getRange("A2:A" + directivaSheet.getLastRow()).getValues();

  // Inicializar una variable para acumular el texto de A8
  var newA8Value = resolucionSheet.getRange("A8").getValue();
  // Inicializar una variable para acumular el texto de A61
  var newA61Value = resolucionSheet.getRange("A61").getValue();

  // Recorrer los valores de la columna A
  for (var i = 0; i < cValues.length; i++) {
    // Verificar si el valor en la columna A es igual al valor de F3
    if (cValues[i][0] == f3Value) {
      
      var seguimiento = directivaSheet.getRange(i + 2, 2).getValue(); 

      // Concatenar el nuevo texto con un nuevo párrafo
      newA8Value += "\n\n" + seguimiento;
      newA61Value += "\n\n" + seguimiento;
    }
  }

  // Establecer el nuevo valor en la celda A8
  resolucionSheet.getRange("A8").setValue(newA8Value);
  // Establecer el nuevo valor en la celda A61
  resolucionSheet.getRange("A61").setValue(newA61Value);
}


//+++++++++++++++ INPUT BOX POSICIONAMIENTOS +++++++++++++++++++++++++++++//

function promptAndAppendToA8AndDirectiva() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Ingresar posicionamientos', 'Por favor, ingresa los posicionamientos:', ui.ButtonSet.OK_CANCEL);

  // Procesar la respuesta del usuario
  if (response.getSelectedButton() == ui.Button.OK) {
    var userInput = response.getResponseText();

    // Obtener la hoja Resolución
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var resolucionSheet = ss.getSheetByName("Resolución");

    // Obtener el valor actual de la celda F3, A8 y A61
    var f3Value = resolucionSheet.getRange("F3").getValue();
    var currentA8Value = resolucionSheet.getRange("A8").getValue();
    var currentA61Value = resolucionSheet.getRange("A61").getValue();
    
    // Concatenar el nuevo texto en la celda A8
    var newA8Value = currentA8Value + "\n\nLos posicionamientos muestran que:\n" + userInput;
    resolucionSheet.getRange("A8").setValue(newA8Value);

    // Concatenar el nuevo texto en la celda A61
    var newA61Value = currentA61Value + "\n\nLos posicionamientos muestran que:\n" + userInput;
    resolucionSheet.getRange("A61").setValue(newA61Value);

    // Obtener la hoja Directiva
    var directivaSheet = ss.getSheetByName("Directiva");

    // Encontrar la última fila con datos en la hoja Directiva
    var lastRow = directivaSheet.getLastRow() + 1;

    // Insertar el valor de F3 en la columna A y el valor del input box en la columna B
    directivaSheet.getRange(lastRow, 1).setValue(f3Value);
    directivaSheet.getRange(lastRow, 2).setValue(userInput);
  } else {
    ui.alert('Operación cancelada. No se realizaron cambios.');
  }
}
//+++++++++++++++ FIN INPUT BOX POSICIONAMIENTOS +++++++++++++++++++++++++++++//

//+++++++++++++++ INPUT BOX BEREDICTO +++++++++++++++++++++++++++++//

function promptAndAppendBeredicto() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Beredicto', 'Por favor, escribe el beredicto:', ui.ButtonSet.OK_CANCEL);

  // Procesar la respuesta del usuario
  if (response.getSelectedButton() == ui.Button.OK) {
    var userInput = response.getResponseText();

    // Obtener la hoja Resolución
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var resolucionSheet = ss.getSheetByName("Resolución");

    // Obtener el valor actual de la celda F3, A8 y A61
    var f3Value = resolucionSheet.getRange("F3").getValue();
    var currentA8Value = resolucionSheet.getRange("A8").getValue();
    var currentA61Value = resolucionSheet.getRange("A61").getValue();
    
    // Concatenar el nuevo texto en la celda A8
    var newA8Value = currentA8Value + "\n\nLa junta directiva determina que:\n" + userInput;
    resolucionSheet.getRange("A8").setValue(newA8Value);

    // Concatenar el nuevo texto en la celda A61
    var newA61Value = currentA61Value + "\n\nLa junta directiva determina que:\n" + userInput;
    resolucionSheet.getRange("A61").setValue(newA61Value);

    // Obtener la hoja Directiva
    var directivaSheet = ss.getSheetByName("Directiva");

    // Obtener todos los valores de la columna A de la hoja Directiva
    var columnAValues = directivaSheet.getRange("A2:A" + directivaSheet.getLastRow()).getValues();

    // Verificar si f3Value ya existe en la columna A
    var rowIndex = -1;
    for (var i = 0; i < columnAValues.length; i++) {
      if (columnAValues[i][0] === f3Value) {
        rowIndex = i + 2; // Ajustar índice para que sea el índice real de la hoja
        break;
      }
    }

    if (rowIndex > -1) {
      // Si el valor ya existe, insertar el valor del prompt en la columna D
      directivaSheet.getRange(rowIndex, 4).setValue(userInput);
    } else {
      // Si el valor no existe, insertar una nueva fila debajo de A1
      directivaSheet.insertRowAfter(1);
      directivaSheet.getRange("A2").setValue(f3Value);
      directivaSheet.getRange("D2").setValue(userInput);
    }
  } else {
    ui.alert('Operación cancelada. No se realizaron cambios.');
  }
}

//+++++++++++++++ FIN INPUT BOX BEREDICTO +++++++++++++++++++++++++++++//

//+++++++++++++++++++++++++++ FALTA GRAVE +++++++++++++++++++++++++++++++++++++++//
function FaltaLeve() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Resolución");
  //sheet.getRange("A123").setValue(""); // borra nombre del infractor antes de ejecutar todo el código
  sheet.getRange("A114").setValue(""); //muy señor nuestro=""
  var hRange = sheet.getRange("H3:H15");
  var selectedCell = sheet.getActiveCell();
  
  
  // Verificar si la celda seleccionada está en el rango H3:H15
  if (selectedCell.getRow() >= 3 && selectedCell.getRow() <= 15 && selectedCell.getColumn() == 8) {
    var cellValue = selectedCell.getValue();
    
    if (cellValue !== "") {
      // Asignar el valor de la celda seleccionada a la celda E133
      sheet.getRange("E133").setValue(cellValue);
      
      // Asignar el texto a la celda C132
      sheet.getRange("A132").setValue("La infracción se ha tipificado según nuestros estatutos como: LEVE.");
      
      // Dejar la celda C136 vacía y rellenarla de color amarillo
      var c136 = sheet.getRange("C136");
      c136.setValue("");
      c136.setBackground("yellow");
      
      // Colorear de amarillo la celda E111
      sheet.getRange("E111").setBackground("yellow");

      // Establecer los valores de las celdas D136 y E136
      var j2Value = sheet.getRange("J2").getValue();
      var k2Value = sheet.getRange("K2").getValue();
      sheet.getRange("D136").setValue("Min: " + j2Value + "€");
      sheet.getRange("E136").setValue("Max: " + k2Value + "€");

      // Seleccionar la celda C136
      sheet.setActiveRange(c136);
      checkAndSetLeve();
    }
  } else {
    // Mostrar mensaje de alerta si la celda seleccionada está fuera del rango H3:H15
    SpreadsheetApp.getUi().alert("Debe seleccionar una celda dentro de la columna H que indique una falta leve");
  }

  SancionenFirme();
}


//+++++++++++++++++++++++++++ FALTA GRAVE +++++++++++++++++++++++++++++++++++++++//
function FaltaGrave() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Resolución");
  //sheet.getRange("A123").setValue(""); // borra nombre del infractor antes de ejecutar todo el código
  sheet.getRange("A114").setValue(""); //muy señor nuestro=""
  var hRange = sheet.getRange("H17:H32");
  var selectedCell = sheet.getActiveCell();
  
  // Verificar si la celda seleccionada está en el rango H3:H15
  if (selectedCell.getRow() >= 17 && selectedCell.getRow() <= 32 && selectedCell.getColumn() == 8) {
    var cellValue = selectedCell.getValue();
    
    if (cellValue !== "") {
      // Asignar el valor de la celda seleccionada a la celda E133
      sheet.getRange("E133").setValue(cellValue);
      
      // Asignar el texto a la celda C132
      sheet.getRange("A132").setValue("La infracción se ha tipificado según nuestros estatutos como: GRAVE.");
      
      // Dejar la celda C136 vacía y rellenarla de color amarillo
      var c136 = sheet.getRange("C136");
      c136.setValue("");
      c136.setBackground("yellow");
      
      // Colorear de amarillo la celda E111
      sheet.getRange("E111").setBackground("yellow");

      // Establecer los valores de las celdas D136 y E136
      var j16Value = sheet.getRange("J16").getValue();
      var k16Value = sheet.getRange("K16").getValue();
      sheet.getRange("D136").setValue("Min: " + j16Value + "€");
      sheet.getRange("E136").setValue("Max: " + k16Value + "€");
      sheet.getRange("C2").setValue("FALTA GRAVE");
      sheet.getRange("C54").setValue("FALTA GRAVE");
      // Seleccionar la celda C136
      sheet.setActiveRange(c136);
    }
  } else {
    // Mostrar mensaje de alerta si la celda seleccionada está fuera del rango H3:H15
    SpreadsheetApp.getUi().alert("Debe seleccionar una celda dentro de la columna H que indique una falta grave");
  }
  SancionenFirme();
  checkAndSetGrave();
}


//+++++++++++++++++++++++++++ FALTA MUY GRAVE +++++++++++++++++++++++++++++++++++++++//
function FaltaMuyGrave() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Resolución");
  //sheet.getRange("A123").setValue(""); // borra nombre del infractor antes de ejecutar todo el código
  sheet.getRange("A114").setValue(""); //muy señor nuestro=""
  var hRange = sheet.getRange("H34:H51");
  var selectedCell = sheet.getActiveCell();
  
  // Verificar si la celda seleccionada está en el rango H3:H15
  if (selectedCell.getRow() >= 34 && selectedCell.getRow() <= 51 && selectedCell.getColumn() == 8) {
    var cellValue = selectedCell.getValue();
    
    if (cellValue !== "") {
      // Asignar el valor de la celda seleccionada a la celda E133
      sheet.getRange("E133").setValue(cellValue);
      
      // Asignar el texto a la celda C132
      sheet.getRange("A132").setValue("La infracción se ha tipificado según nuestros estatutos como: MUY GRAVE.");
      
      // Dejar la celda C136 vacía y rellenarla de color amarillo
      var c136 = sheet.getRange("C136");
      c136.setValue("");
      c136.setBackground("yellow");
      
      // Colorear de amarillo la celda E111
      sheet.getRange("E111").setBackground("yellow");

      // Establecer los valores de las celdas D136 y E136
      var j33Value = sheet.getRange("J33").getValue();
      var k33Value = sheet.getRange("K33").getValue();
      sheet.getRange("D136").setValue("Min: " + j33Value + "€");
      sheet.getRange("E136").setValue("Max: " + k33Value + "€");


      // Seleccionar la celda C136
      sheet.setActiveRange(c136);
    }
  } else {
    // Mostrar mensaje de alerta si la celda seleccionada está fuera del rango H3:H15
    SpreadsheetApp.getUi().alert("Debe seleccionar una celda dentro de la columna H que indique una falta muy grave");
  }
  SancionenFirme();
  checkAndSetMuyGrave();
}


//++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++//
//----------- SANCION EN FIRME--------------------------//

function SancionenFirme() {
  mostrarFilasReg();
  // Obtener las hojas de trabajo
  const hojaSocios = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Socios");
  const hojaResolucion = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Resolución");

  // Obtener el valor buscado en la celda A117 de la hoja Resolución
  const valorBuscado = hojaResolucion.getRange("A117").getValue();

  // Obtener todos los valores de la columna A de la hoja Socios
  const valoresSocios = hojaSocios.getRange("A:A").getValues();

  // Variable para verificar si se encontró el valor
  let valorEncontrado = false;

  // Recorrer los valores de la columna A para encontrar el valor buscado
  for (let i = 0; i < valoresSocios.length; i++) {
    if (valoresSocios[i][0] === valorBuscado) {
      // Si se encuentra el valor, obtener el valor de la columna I en la misma fila
      const valorMatricula = hojaSocios.getRange(i + 1, 9).getValue();
      const valorNombre = hojaSocios.getRange(i + 1, 2).getValue();
      
      // Asignar el valor correspondiente a la celda B117 de la hoja Resolución
      hojaResolucion.getRange("B117").setValue(valorMatricula);
      hojaResolucion.getRange("A120").setValue(valorNombre);
      // Indicar que el valor fue encontrado
      valorEncontrado = true;
      break; // Detener el bucle al encontrar el valor
    }
  }
  matriculasPpliegos();
}

//+++++++++++++   OCULTAR SANCIÓN EN FIRME   ++++++++++++//
function ocultarFilasReg() {
  // Obtener la hoja específica por nombre
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Resolución");
  
  if (sheet) {
    // Ocultar las filas 104 a 157
    sheet.hideRows(104, 157 - 104 + 1);
  } else {
    Logger.log("La hoja 'Registro entrada sanciones' no se encontró.");
  }
}
//+++++++++++++   FIN OCULTAR SANCIÓN EN FIRME   ++++++++++++//

//+++++++++++++   MOSTRAR SANCIÓN EN FIRME   ++++++++++++//
function mostrarFilasReg() {
  // Obtener la hoja específica por nombre
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Resolución");
  
  if (sheet) {
    // Mostrar las filas 104 a 157
    sheet.showRows(104, 157 - 104 + 1);
    Logger.log("Filas 104-157 mostradas."); // Registro para confirmar la acción
  } else {
    Logger.log("La hoja 'Registro entrada sanciones' no se encontró.");
  }
}
//+++++++++++++   FIN MOSTRAR SANCIÓN EN FIRME   ++++++++++++//

//+++++++++++++ EXPORTAR PAGO DE SANCIONES +++++++++++++++++//
/*function exportarDatosPago() {
  // ID del libro de origen
  const idLibroOrigen = '1DgjB0vTBB2WZhr_Nnr7AhYWo5Djxut5ZstjLZvs8614'; // Reemplaza con el ID del libro "Nuevo registro de sanciones"
  
  // ID del libro de destino
  const idLibroDestino = '1PTmboi4op93bqBuuYLPboTg9KbRob8edUNpLm5OPQOM'; // Reemplaza con el ID del libro "Sanciones para administrativo"

  // Nombre de la hoja de origen
  const nombreHojaOrigen = 'Pago Sanciones';
  
  // Nombre de la hoja de destino
  const nombreHojaDestino = 'Informe Sanciones';

  // Obtén el libro de origen
  const libroOrigen = SpreadsheetApp.openById(idLibroOrigen);
  const hojaOrigen = libroOrigen.getSheetByName(nombreHojaOrigen);

  // Obtén el libro de destino
  const libroDestino = SpreadsheetApp.openById(idLibroDestino);
  const hojaDestino = libroDestino.getSheetByName(nombreHojaDestino);

  // Obtén todos los datos de la hoja de origen
  const rangoOrigen = hojaOrigen.getDataRange();
  const datosOrigen = rangoOrigen.getValues();

  // Limpia la hoja de destino antes de copiar los datos
  hojaDestino.clear();

  // Copia los datos en la hoja de destino
  hojaDestino.getRange(1, 1, datosOrigen.length, datosOrigen[0].length).setValues(datosOrigen);
}*/
function exportarDatosPago() {
  // ID del libro de origen
  const idLibroOrigen = '1DgjB0vTBB2WZhr_Nnr7AhYWo5Djxut5ZstjLZvs8614'; // Reemplaza con el ID del libro "Nuevo registro de sanciones"
  
  // ID del libro de destino
  const idLibroDestino = '1PTmboi4op93bqBuuYLPboTg9KbRob8edUNpLm5OPQOM'; // Reemplaza con el ID del libro "Sanciones para administrativo"

  // Nombre de la hoja de origen
  const nombreHojaOrigen = 'Pago Sanciones';
  
  // Nombre de la hoja de destino
  const nombreHojaDestino = 'Informe Sanciones';

  // Obtén el libro de origen
  const libroOrigen = SpreadsheetApp.openById(idLibroOrigen);
  const hojaOrigen = libroOrigen.getSheetByName(nombreHojaOrigen);

  // Obtén el libro de destino
  const libroDestino = SpreadsheetApp.openById(idLibroDestino);
  const hojaDestino = libroDestino.getSheetByName(nombreHojaDestino);

  // Define el rango específico de la hoja de origen que deseas copiar
  const rangoOrigen = hojaOrigen.getRange('A2:G'); // Selecciona el rango A2:G de la hoja de origen
  const datosOrigen = rangoOrigen.getValues();

  // Limpia los datos de la hoja de destino desde la fila 3 hacia abajo antes de copiar los nuevos datos
  hojaDestino.getRange(3, 1, hojaDestino.getLastRow() - 2, hojaDestino.getLastColumn()).clear();

  // Copia los datos en la hoja de destino comenzando desde la fila 3
  hojaDestino.getRange(3, 1, datosOrigen.length, datosOrigen[0].length).setValues(datosOrigen);
}


//+++++++++++++ FIN EXPORTAR PAGO DE SANCIONES +++++++++++++++++//

//++++++++++ EXPORTAR DATOS A LA HOJA IMPRIMIR SANCIONES +++++++//

  /*function exportarSanciones() {
  // Abre los libros de trabajo
  var libroOrigen = SpreadsheetApp.openById("1DgjB0vTBB2WZhr_Nnr7AhYWo5Djxut5ZstjLZvs8614"); // Reemplaza con el ID de "Nuevo registro de sanciones"
  var libroDestino = SpreadsheetApp.openById("1DmKQQtkx4z1xCkGxL_Hr7sJdLgoyDT71K1dFZ7gtrpQ"); // Reemplaza con el ID de "Imprimir sanciones"
  
  // Obtén las hojas relevantes
  var hojaOrigen = libroOrigen.getSheetByName("Registro entrada sanciones");
  var hojaDestino = libroDestino.getSheetByName("datos");
  
  // Borra todos los datos de la hoja destino
  hojaDestino.getRange(2, 1, hojaDestino.getLastRow() - 1, hojaDestino.getLastColumn()).clearContent();
  
  // Obtén los datos de la hoja origen, excluyendo la primera fila
  var datos = hojaOrigen.getDataRange().getValues();
  
  // Inicializa un array para almacenar las filas que cumplen la condición
  var filasParaExportar = [];
  
  // Recorre los datos comenzando desde la segunda fila (índice 1) para omitir los encabezados
  for (var i = 1; i < datos.length; i++) {
    if (datos[i][13] === "") { // Columna N (índice 13) está vacía
      filasParaExportar.push(datos[i]);
    }
  }
  
  // Exporta las filas filtradas a la hoja destino si hay filas que exportar
  if (filasParaExportar.length > 0) {
    hojaDestino.getRange(2, 1, filasParaExportar.length, filasParaExportar[0].length).setValues(filasParaExportar);
  }
}

*/
function exportarSanciones() {
  // Abre los libros de trabajo
  var libroOrigen = SpreadsheetApp.openById("1DgjB0vTBB2WZhr_Nnr7AhYWo5Djxut5ZstjLZvs8614"); // Reemplaza con el ID de "Nuevo registro de sanciones"
  var libroDestino = SpreadsheetApp.openById("1DmKQQtkx4z1xCkGxL_Hr7sJdLgoyDT71K1dFZ7gtrpQ"); // Reemplaza con el ID de "Imprimir sanciones"
  
  // Obtén las hojas relevantes
  var hojaOrigen = libroOrigen.getSheetByName("Registro entrada sanciones");
  var hojaDestino = libroDestino.getSheetByName("datos");
  
  // Borra todos los datos de la hoja destino
  hojaDestino.getRange(2, 1, hojaDestino.getLastRow() - 1, hojaDestino.getLastColumn()).clearContent();
  
  // Obtén los datos de la hoja origen, excluyendo la primera fila
  var datos = hojaOrigen.getDataRange().getValues();
  
  // Inicializa un array para almacenar las filas que cumplen la condición
  var filasParaExportar = [];
  
  // Recorre los datos comenzando desde la segunda fila (índice 1) para omitir los encabezados
  for (var i = 1; i < datos.length; i++) {
    if (datos[i][13] === "") { // Columna N (índice 13) está vacía
      
      // Extrae la hora si la columna J contiene un valor de fecha/hora
      if (datos[i][9] instanceof Date) {
        // Convierte el valor de la columna J a una cadena de solo hora
        var horas = datos[i][9].getHours().toString().padStart(2, '0');
        var minutos = datos[i][9].getMinutes().toString().padStart(2, '0');
        var segundos = datos[i][9].getSeconds().toString().padStart(2, '0');
        datos[i][9] = horas + ':' + minutos + ':' + segundos;
      }
      
      filasParaExportar.push(datos[i]); // Agrega la fila completa
    }
  }
  
  // Pega los datos en la hoja destino si hay filas que exportar
  if (filasParaExportar.length > 0) {
    hojaDestino.getRange(2, 1, filasParaExportar.length, filasParaExportar[0].length).setValues(filasParaExportar);
  }
}



 
//++++++ FIN EXPORTAR DATOS A LA HOJA IMPRIMIR SANCIONES +++++++//


//++++++++++++ IMPORTAR ESTADO DEL PAGO DE SANCIONES DE ADMINISTRATIVO +++++++++++//
function actualizarEstadoDePagoDeSanciones() {
  // IDs de los documentos
  var libroNuevoRegistroID = "1DgjB0vTBB2WZhr_Nnr7AhYWo5Djxut5ZstjLZvs8614";  // Reemplaza con el ID real del libro "Nuevo registro de sanciones"
  var libroSancionesID = "1PTmboi4op93bqBuuYLPboTg9KbRob8edUNpLm5OPQOM";  // Reemplaza con el ID real del libro "Sanciones para administrativo"

  // Abrir los documentos
  var libroNuevoRegistro = SpreadsheetApp.openById(libroNuevoRegistroID);
  var libroSanciones = SpreadsheetApp.openById(libroSancionesID);

  // Obtener las hojas
  var hojaPagoSanciones = libroNuevoRegistro.getSheetByName("Pago Sanciones");
  var hojaInformeSanciones = libroSanciones.getSheetByName("Informe Sanciones");

  // Obtener los datos de las columnas A2:A
  var datosPagoSanciones = hojaPagoSanciones.getRange("A2:A" + hojaPagoSanciones.getLastRow()).getValues();
  var datosInformeSanciones = hojaInformeSanciones.getRange("A2:A" + hojaInformeSanciones.getLastRow()).getValues();

  // Crear un mapa de búsqueda para los datos de la hoja Informe Sanciones
  var mapaInforme = {};
  for (var i = 0; i < datosInformeSanciones.length; i++) {
    var clave = datosInformeSanciones[i][0];
    if (clave) {
      mapaInforme[clave] = {
        G: hojaInformeSanciones.getRange("G" + (i + 2)).getValue(),
        H: hojaInformeSanciones.getRange("H" + (i + 2)).getValue()
      };
    }
  }

  // Actualizar las celdas G y H en Pago Sanciones según coincidencias
  for (var j = 0; j < datosPagoSanciones.length; j++) {
    var valorPago = datosPagoSanciones[j][0];
    if (valorPago && mapaInforme[valorPago]) {
      hojaPagoSanciones.getRange("G" + (j + 2)).setValue(mapaInforme[valorPago].G);
      hojaPagoSanciones.getRange("H" + (j + 2)).setValue(mapaInforme[valorPago].H);
    }
  }

  // Obtener las fechas de la columna F2:F en Informe Sanciones
  var fechasInformeSanciones = hojaInformeSanciones.getRange("F2:F" + hojaInformeSanciones.getLastRow()).getValues();
  var fechaActual = new Date();

  // Actualizar el color de la fuente en la columna F basado en la fecha
  for (var k = 0; k < fechasInformeSanciones.length; k++) {
    var fecha = fechasInformeSanciones[k][0];
    var celda = hojaInformeSanciones.getRange("F" + (k + 2));
    if (fecha instanceof Date) {
      if (fecha > fechaActual) {
        celda.setFontColor("red");
      } else {
        celda.setFontColor("green");
      }
    }
  }
}
//++++++++++++ FIN IMPORTAR ESTADO DEL PAGO DE SANCIONES DE ADMINISTRATIVO +++++++++++//



/*function colorearCeldasPliegos() {
  var hojaSanciones = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registro entrada sanciones");
  var hojaPliegos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registro pliegos");
  
  // Obtener los datos de las columnas
  var rangoM = hojaSanciones.getRange("M2:M" + hojaSanciones.getLastRow());
  var valoresM = rangoM.getValues().flat();
  
  var rangoC = hojaPliegos.getRange("C2:C" + hojaPliegos.getLastRow());
  var valoresC = rangoC.getValues().flat();
  
  // Recorrer cada celda en la columna M
  for (var i = 0; i < valoresM.length; i++) {
    var valorM = valoresM[i];
    
    // Verificar si el valor en la columna M está en la columna C
    if (valoresC.includes(valorM)) {
      // Si se encuentra el valor, colorear de amarillo claro
      rangoM.getCell(i + 1, 1).setBackground("#ffff99");
    } else {
      // Si no se encuentra el valor, colorear de azul claro
      rangoM.getCell(i + 1, 1).setBackground("#cfe2f3");
    }
  }
}
*/
function colorearCeldasPliegos() {
  var hojaSanciones = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registro entrada sanciones");
  var hojaPliegos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registro pliegos");
  
  // Obtener los datos de las columnas
  var rangoM = hojaSanciones.getRange("M2:M" + hojaSanciones.getLastRow());
  var valoresM = rangoM.getValues().flat();
  
  var rangoC = hojaPliegos.getRange("C2:C" + hojaPliegos.getLastRow());
  var valoresC = rangoC.getValues().flat();
  
  // Obtener los valores de las columnas R y S
  var rangoR = hojaSanciones.getRange("R2:R" + hojaSanciones.getLastRow());
  var valoresR = rangoR.getValues().flat();
  
  var rangoS = hojaSanciones.getRange("S2:S" + hojaSanciones.getLastRow());
  var valoresS = rangoS.getValues().flat();
  
  // Recorrer cada celda en la columna M
  for (var i = 0; i < valoresM.length; i++) {
    var valorM = valoresM[i];
    var valorR = valoresR[i];
    var valorS = valoresS[i];
    
    // Verificar si las columnas R o S están vacías
    if (!valorR || !valorS) {
      // Si cualquiera de las columnas R o S está vacía, colorear la celda de la columna M de beige
      rangoM.getCell(i + 1, 1).setBackground("#f5f5dc"); // Beige
    } else if (valoresC.includes(valorM)) {
      // Si el valor de la columna M está en la columna C, colorear de amarillo claro
      rangoM.getCell(i + 1, 1).setBackground("#ffff99");
    } else {
      // Si no se encuentra el valor en la columna C, colorear de azul claro
      rangoM.getCell(i + 1, 1).setBackground("#cfe2f3");
    }
  }
}
function buscarHistorialDenunciante() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Obtener hojas
  var hojaResolucion = ss.getSheetByName("Resolución");
  var hojaRegistro = ss.getSheetByName("Registro entrada sanciones");
  var hojaPliegos = ss.getSheetByName("Registro pliegos");


// Lista de filas que deseas actualizar
var filas = [3, 5, 7, 9, 11, 15, 18, 20, 22, 24, 26, 28, 30];

// Asignar "---" a las celdas en la columna L
for (var i = 0; i < filas.length; i++) {
  hojaResolucion.getRange("L" + filas[i]).setValue("---");
}

// Asignar "---" a las celdas en la columna M
for (var i = 0; i < filas.length; i++) {
  hojaResolucion.getRange("M" + filas[i]).setValue("---");
}


  // Obtener valores de la celda F3 de la hoja Resolución
  var valorF3 = hojaResolucion.getRange("F3").getValue();
  
  // Obtener el rango de la columna M de la hoja Registro entrada sanciones
  var rangoColumnaM = hojaRegistro.getRange("M:M").getValues();

  // Obtener el rango de la columna C de la hoja Pliegos
  var rangoColumnaC = hojaPliegos.getRange("C:C").getValues();
  
  // Recorrer la columna M para buscar el valor
  for (var i = 0; i < rangoColumnaM.length; i++) {
    if (rangoColumnaM[i][0] == valorF3) {  // Si coincide el valor
      // Obtener valores de las columnas C y D en la misma fila
      var valorC = hojaRegistro.getRange(i + 1, 3).getValue(); // Columna C
      var valorD = hojaRegistro.getRange(i + 1, 4).getValue(); // Columna D
      var valorE = hojaRegistro.getRange(i + 1, 5).getValue(); // Columna E
      var valorO = hojaRegistro.getRange(i + 1, 15).getValue(); 
      var valorQ = hojaRegistro.getRange(i + 1, 17).getValue(); 
      var valorF = hojaRegistro.getRange(i + 1, 6).getValue();
      var valorH = hojaRegistro.getRange(i + 1, 8).getValue();

       
      // Asignar valores 
    
      hojaResolucion.getRange("L7").setValue(valorC);
      hojaResolucion.getRange("L9").setValue(valorD);
      hojaResolucion.getRange("L11").setValue(valorE);
      hojaResolucion.getRange("L13").setValue(valorO+valorQ);
      hojaResolucion.getRange("L3").setValue(valorF);
      hojaResolucion.getRange("L18").setValue(valorH);
      break; // Terminar el bucle una vez encontrado
    }
  }

   for (var i = 0; i < rangoColumnaC.length; i++) {
    if (rangoColumnaC[i][0] == valorF3) {  // Si coincide el valor
      // Obtener valores de las columnas C y D en la misma fila

      var PliegoJ = hojaPliegos.getRange(i + 1, 10).getValue(); // Columna J
      var PliegoC = hojaPliegos.getRange(i + 1, 4).getValue(); 
      var PliegoD = hojaPliegos.getRange(i + 1, 6).getValue(); 
      var PliegoE = hojaPliegos.getRange(i + 1, 5).getValue(); 

      hojaResolucion.getRange("L22").setValue(PliegoJ);
      hojaResolucion.getRange("L24").setValue(PliegoC);
      hojaResolucion.getRange("L26").setValue(PliegoD);
      hojaResolucion.getRange("L28").setValue(PliegoE);

      break; // Terminar el bucle una vez encontrado
    }
  }
  // Verificar si la celda L28 está vacía
if (hojaResolucion.getRange("L28").getValue() === "") {
  hojaResolucion.getRange("L28").setValue("---");
}
  NumeroDePliegos();
}
function NumeroDePliegos() {
  // Acceder a las hojas de cálculo
  const hojaResolucion = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Resolución");
  const hojaRegistro = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registro pliegos");
  
  // Obtener los valores de las celdas L22, L24, L26 y L28 de la hoja Resolución
  const valoresBuscados = [
    hojaResolucion.getRange("L22").getValue(),
    hojaResolucion.getRange("L24").getValue(),
    hojaResolucion.getRange("L26").getValue(),
    hojaResolucion.getRange("L28").getValue()
  ];
  
  // Inicializar contador
  let contador = 0;
  
  // Obtener todos los datos de la hoja Registro pliegos
  const rangoDatos = hojaRegistro.getDataRange();
  const datos = rangoDatos.getValues();
  
  // Buscar coincidencias en cada fila de la hoja Registro pliegos
  for (let i = 0; i < datos.length; i++) {
    for (let j = 0; j < valoresBuscados.length; j++) {
      // Comprobar si el valor buscado está en la fila actual
      if (datos[i].includes(valoresBuscados[j])) {
        contador++; // Incrementar contador si se encuentra
        break; // Salir del bucle interno si se encuentra una coincidencia
      }
    }
  }
  
  // Actualizar la celda M18 de la hoja Resolución
  hojaResolucion.getRange("M18").setValue(contador === 0 ? 0 : contador);
}
function NumeroDePliegosDte() {
    // Verificar si la celda L28 está vacía
if (hojaResolucion.getRange("L13").getValue() === "") {
  hojaResolucion.getRange("L13").setValue("---");
}
  NumeroDePliegos();
  // Acceder a las hojas de cálculo
  const hojaResolucion = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Resolución");
  const hojaRegistro = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registro entrada sanciones");
  
  // Obtener los valores de las celdas L7, L9, L11 y L13 de la hoja Resolución
  const valoresBuscados = [
    hojaResolucion.getRange("L7").getValue(),
    hojaResolucion.getRange("L9").getValue(),
    hojaResolucion.getRange("L11").getValue(),
    hojaResolucion.getRange("L13").getValue()
  ];
  
  // Inicializar contador
  let contador = 0;
  
  // Obtener todos los datos de la hoja Registro entrada sanciones
  const rangoDatos = hojaRegistro.getDataRange();
  const datos = rangoDatos.getValues();
  
  // Buscar coincidencias en cada fila de la hoja Registro entrada sanciones
  for (let i = 0; i < datos.length; i++) {
    for (let j = 0; j < valoresBuscados.length; j++) {
      // Comprobar si el valor buscado está en la fila actual
      if (datos[i].includes(valoresBuscados[j])) {
        contador++; // Incrementar contador si se encuentra
        break; // Salir del bucle interno si se encuentra una coincidencia
      }
    }
  }
  
  // Actualizar la celda M18 de la hoja Resolución
  hojaResolucion.getRange("M3").setValue(contador === 0 ? 0 : contador);
}

function matriculasPpliegos() {
  // Obtén la hoja de cálculo activa
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Obtén las hojas 'Registro pliegos' y 'Resolución'
  var hojaRegistroPliegos = ss.getSheetByName('Registro pliegos');
  var hojaResolucion = ss.getSheetByName('Resolución');
  
  // Obtén el valor de la celda F109 de la hoja 'Resolución'
  var valorBusqueda = hojaResolucion.getRange('F109').getValue();
  
  // Obtén los valores de la columna C de la hoja 'Registro pliegos'
  var valoresColumnaC = hojaRegistroPliegos.getRange('C2:C' + hojaRegistroPliegos.getLastRow()).getValues();
  
  // Recorre la columna C para encontrar una coincidencia
  for (var i = 0; i < valoresColumnaC.length; i++) {
    if (valoresColumnaC[i][0] === valorBusqueda) {
      // Si encuentra una coincidencia, obtén el valor de la columna H correspondiente
      var valorColumnaH = hojaRegistroPliegos.getRange('H' + (i + 2)).getValue();
      
      // Convierte el valor a mayúsculas
      var valorEnMayusculas = valorColumnaH.toString().toUpperCase();
      
      // Actualiza la celda I116 en la hoja 'Resolución' con el valor en mayúsculas
      hojaResolucion.getRange('I116').setValue(valorEnMayusculas);
      
      // Sal del bucle una vez que encuentres la coincidencia
      break;
    }
  }
  
  // Obtén los valores actuales de las celdas I116 y B117
  var valorI116 = hojaResolucion.getRange('I116').getValue();
  var valorB117 = hojaResolucion.getRange('B117').getValue();
  
  // Compara los valores y establece el color de fondo
  if (valorI116 !== valorB117) {
    // Si son diferentes, cambia el fondo a amarillo
    hojaResolucion.getRange('I116').setBackground('yellow');
    hojaResolucion.getRange('B117').setBackground('yellow');
  } else {
    // Si son iguales, cambia el fondo a blanco
    hojaResolucion.getRange('I116').setBackground('white');
    hojaResolucion.getRange('B117').setBackground('white');
  }
}
function convertirCeldaSeleccionadaAMayusculas() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rango = hoja.getActiveRange(); // Rango seleccionado
  var valores = rango.getValues();   // Obtiene todos los valores del rango

  for (var i = 0; i < valores.length; i++) {
    for (var j = 0; j < valores[i].length; j++) {
      var valor = valores[i][j];

      // Si es texto, lo convierte a mayúsculas
      if (typeof valor === 'string') {
        valores[i][j] = valor.toUpperCase();
      }
    }
  }

  rango.setValues(valores); // Aplica los cambios

  SpreadsheetApp.getUi().alert("Las celdas seleccionadas han sido convertidas a MAYÚSCULAS.");
}

function convertirCeldaSeleccionadaANombrePropio() {

  var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rango = hoja.getActiveRange(); // Rango seleccionado
  var valores = rango.getValues();   // Obtiene todos los valores del rango

  for (var i = 0; i < valores.length; i++) {
    for (var j = 0; j < valores[i].length; j++) {
      var valor = valores[i][j];

      // Si es texto, lo transforma a nombre propio
      if (typeof valor === 'string') {
        valores[i][j] = valor
          .toLowerCase()
          .split(' ')
          .map(palabra => palabra.charAt(0).toUpperCase() + palabra.slice(1))
          .join(' ');
      }
    }
  }

  rango.setValues(valores); // Establece los valores modificados en el mismo rango

  SpreadsheetApp.getUi().alert("Las celdas seleccionadas han sido convertidas a formato de nombre propio.");
}

function convertirCeldaSeleccionadaAMinusculas() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rango = hoja.getActiveRange(); // Rango seleccionado
  var valores = rango.getValues();   // Todos los valores del rango

  // Recorre cada celda del rango
  for (var i = 0; i < valores.length; i++) {
    for (var j = 0; j < valores[i].length; j++) {
      var celda = valores[i][j];
      
      // Si es texto, lo convierte a minúsculas
      if (typeof celda === 'string') {
        valores[i][j] = celda.toLowerCase();
      }
    }
  }

  // Establece los valores modificados en el mismo rango
  rango.setValues(valores);

  SpreadsheetApp.getUi().alert("Las celdas seleccionadas han sido convertidas a minúsculas.");
}

//_______IMPRIMIR____________________//

function imprimirInformeDirectiva() {
 var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Resolución");
  if (hoja) {
    var rango = hoja.getRange("A1:G52");
    hoja.setActiveRange(rango);
  } else {
    Logger.log("No se encontró la hoja 'Resolución'.");
  }
}

/*function imprimirInformeSancionado() {
 var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Resolución");
  if (hoja) {
    var rango = hoja.getRange("A1:G104");
    hoja.setActiveRange(rango);
  } else {
    Logger.log("No se encontró la hoja 'Resolución'.");
  }
}*/
function imprimirInformeSancionado() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Resolución");
  if (hoja) {
    // Establecer el valor "AMONESTACIÓN" en las celdas C2 y C54
    hoja.getRange("C2").setValue("AMONES- TACIÓN");
    hoja.getRange("C54").setValue("AMONESTACIÓN");

    // Establecer el rango A1:G104 como el rango activo
    var rango = hoja.getRange("A1:G104");
    hoja.setActiveRange(rango);
  } else {
    Logger.log("No se encontró la hoja 'Resolución'.");
  }
}

function imprimirSancionEnFirme() {
 var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Resolución");
  if (hoja) {
    var rango = hoja.getRange("A1:G156");
    hoja.setActiveRange(rango);
  } else {
    Logger.log("No se encontró la hoja 'Resolución'.");
  }
}

function imprimirCopiaSancionEnFirme() {
 var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Resolución");
  if (hoja) {
    var rango = hoja.getRange("A107:G156");
    hoja.setActiveRange(rango);
  } else {
    Logger.log("No se encontró la hoja 'Resolución'.");
  }
}
//---------------------CAMBIAR MATRICULA----------------------//

function CambiarMatricula() {

  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Resolución"); 
  var celdaDestino = hoja.getRange("B117");

  // Copiar valor de I116 a B117
  var valor = hoja.getRange("I116").getValue();
  celdaDestino.setValue(valor);

  // Cambiar el fondo de la celda a blanco
  celdaDestino.setBackground("#FFFFFF");

  // Aplicar bordes blancos
  celdaDestino.setBorder(true, true, true, true, true, true, "#FFFFFF", SpreadsheetApp.BorderStyle.SOLID);
}

function IntroducirCantidadAPagar() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Resolución"); 
  var celdaDestino = hoja.getRange("C136");

  // Copiar valor de I116 a B117
  var valor = hoja.getRange("I132").getValue();
  celdaDestino.setValue(valor);

  // Cambiar el fondo de la celda a blanco
  celdaDestino.setBackground("#FFFFFF");

  // Aplicar bordes blancos
  celdaDestino.setBorder(true, true, true, true, true, true, "#FFFFFF", SpreadsheetApp.BorderStyle.SOLID);
  hoja.getRange("D136").setValue("");
  hoja.getRange("E136").setValue("");
  hoja.getRange("I132").setValue("");
}
function limpiarLmPliegos() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registro pliegos");
  const rango = hoja.getRange("G2:G" + hoja.getLastRow());
  const valores = rango.getValues();

  const valoresLimpios = valores.map(fila => {
    const celda = fila[0];
    if (typeof celda === 'string' || typeof celda === 'number') {
      const numerosSolo = String(celda).replace(/[^\d]/g, ''); // elimina todo excepto dígitos
      const conCeros = numerosSolo.padStart(3, '0'); // asegura que tenga al menos 3 cifras
      return [conCeros];
    } else {
      return [''];
    }
  });

  rango.setValues(valoresLimpios);
}
function limpiarLmSanciones() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registro entrada sanciones");
  const rango = hoja.getRange("H2:H" + hoja.getLastRow());
  const valores = rango.getValues();

  const valoresLimpios = valores.map(fila => {
    const celda = fila[0];
    if (typeof celda === 'string' || typeof celda === 'number') {
      const numerosSolo = String(celda).replace(/[^\d]/g, ''); // elimina todo excepto dígitos
      const conCeros = numerosSolo.padStart(3, '0'); // asegura que tenga al menos 3 cifras
      return [conCeros];
    } else {
      return [''];
    }
  });

  rango.setValues(valoresLimpios);
  limpiarLmSancionador();
}
function limpiarLmSancionador() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registro entrada sanciones");
  const rango = hoja.getRange("F2:F" + hoja.getLastRow());
  const valores = rango.getValues();

  const valoresLimpios = valores.map(fila => {
    const celda = fila[0];
    if (typeof celda === 'string' || typeof celda === 'number') {
      const numerosSolo = String(celda).replace(/[^\d]/g, ''); // elimina todo excepto dígitos
      const conCeros = numerosSolo.padStart(3, '0'); // asegura que tenga al menos 3 cifras
      return [conCeros];
    } else {
      return [''];
    }
  });

  rango.setValues(valoresLimpios);
}
function showSelectedCellContent() { //abrir ventana con el contenido de una celda
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var activeCell = spreadsheet.getActiveCell();
  var cellValue = activeCell.getValue();
  
  // Crea una ventana emergente HTML con el contenido de la celda
  var htmlOutput = HtmlService.createHtmlOutput('<h1>Contenido de la Celda Seleccionada</h1><H2>' + cellValue + '</H2>')
    .setWidth(500)
    .setHeight(550);
  
  // Muestra la ventana emergente
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Contenido de la Celda');
}
function AumentarPliego() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var celda = hoja.getActiveCell();
  
  // Verificar si la celda seleccionada está en la hoja "Registro entrada sanciones"
  if (hoja.getName() === "Registro entrada sanciones") {
    // Verificar si la celda seleccionada está en el rango M2:M
    if (celda.getColumn() === 13 && celda.getRow() >= 2) {
      var valor = celda.getValue();
      var hojaPliegos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registro pliegos");
      var rangoPliegos = hojaPliegos.getRange(2, 3, hojaPliegos.getLastRow() - 1);
      var valoresPliegos = rangoPliegos.getValues();
      
      // Buscar el valor en el rango C2:C de la hoja "Registro pliegos"
      for (var i = 0; i < valoresPliegos.length; i++) {
        if (valoresPliegos[i][0] === valor) {
          var celdaJ = hojaPliegos.getRange(i + 2, 10).getValue();
          var celdaI = hojaPliegos.getRange(i + 2, 9).getValue();
          var mensaje = celdaJ + "\n\n" + celdaI;
          
          // Crear un diálogo modal con contenido HTML
          var html = HtmlService.createHtmlOutput("<div style='font-size: 18px;'>" + celdaJ + '<br>' + '<br>' + celdaI + "</div>");
          
          // Mostrar el diálogo modal
          SpreadsheetApp.getUi().showModalDialog(html, "Mensaje");
          break;
        }
      }
    }
  }
}

function crearBarraLateral() {
  var html = HtmlService.createHtmlOutputFromFile('barraLateral');
  html.setWidth(200);
  SpreadsheetApp.getUi().showSidebar(html);
}

function sancionar() {
  // Aquí debes agregar el código para la función sancionar()
  // Por ejemplo:
  SpreadsheetApp.getUi().alert('La función sancionar() se ha ejecutado');
}
function activarHoja1() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = spreadsheet.getSheetByName("Registro entrada sanciones");
  if (hoja) {
    spreadsheet.setActiveSheet(hoja);
  }
}
function activarHoja2() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = spreadsheet.getSheetByName("Registro pliegos");
  if (hoja) {
    spreadsheet.setActiveSheet(hoja);
  }
}
function activarHoja3() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = spreadsheet.getSheetByName("Testigos");
  if (hoja) {
    spreadsheet.setActiveSheet(hoja);
  }
}
function activarHoja4() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = spreadsheet.getSheetByName("Resolución");
  if (hoja) {
    spreadsheet.setActiveSheet(hoja);
  }
}
function activarHoja5() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = spreadsheet.getSheetByName("Socios");
  if (hoja) {
    spreadsheet.setActiveSheet(hoja);
  }
}
function activarHoja6() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = spreadsheet.getSheetByName("Directiva");
  if (hoja) {
    spreadsheet.setActiveSheet(hoja);
  }
}
function activarHoja7() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = spreadsheet.getSheetByName("Pago Sanciones");
  if (hoja) {
    spreadsheet.setActiveSheet(hoja);
  }
}
function buscarCoincidencias(texto) { //Busca coincidencias al escribir un nombre
  const hojas = ["Registro entrada sanciones", "Registro pliegos"];
  let coincidencias = [];

  hojas.forEach(nombreHoja => {
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHoja);
    const valores = hoja.getRange("C2:C" + hoja.getLastRow()).getValues().flat();
    const filtrados = valores.filter(nombre =>
      nombre && nombre.toString().toLowerCase().includes(texto.toLowerCase())
    );
    coincidencias = coincidencias.concat(filtrados);
  });

  // Elimina duplicados
  return [...new Set(coincidencias)];
}

function FiltrarNombre(nombre) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaEntrada = ss.getSheetByName("Registro entrada sanciones");
  const hojaPliegos = ss.getSheetByName("Registro pliegos");

  const datosEntrada = hojaEntrada.getRange("C2:N" + hojaEntrada.getLastRow()).getValues();
  const datosPliegos = hojaPliegos.getRange("C2:K" + hojaPliegos.getLastRow()).getValues();

  let html = `<div style="font-family: Arial, sans-serif; font-size: 14px;">`;

  // Encabezado con nombre buscado
  html += `<p><b><i><u>${nombre}</u></i></b></p><br><br>`;

  // Buscar coincidencias en Registro entrada sanciones
  datosEntrada.forEach((fila, i) => {
    const celdaC = fila[0]?.toString().toLowerCase();
    if (celdaC && celdaC.includes(nombre.toLowerCase())) {
      const expediente = fila[10] || "";
      const dni = fila[1] || "";
      const telefono = fila[2] || "";
      const lm = fila[3] || "";
      const unidad = fila[5] || "";
      const testigos = fila[9] || "";
      const resolucion = fila[11] || "";
      const exponeEntrada = fila[8] || "";

      html += `
        <p>
          Expediente nº: ${expediente}, Denunciante: ${fila[0]}; DNI: ${dni}, Teléfono: ${telefono}, Lm: ${lm}, denuncia a la unidad: ${unidad} con testigos: ${testigos}, cuya resolución es: <u>${resolucion}</u><br>
          Expone que: ${exponeEntrada}<br>
      `;

      // Buscar en pliegos por expediente
      const matchPliego = datosPliegos.find(p => (p[0] || "").toString().trim() === expediente);
      if (matchPliego) {
        const denunciado = matchPliego[7] || "";
        const lmPliego = matchPliego[4] || "";
        const exponePliego = matchPliego[6] || "";

        html += `<span style="color:blue">Denunciado: ${denunciado}, Lm: ${lmPliego}, Expone que: ${exponePliego}</span>`;
      }

      html += `</p><br><br>`;
    }
  });

  html += `<hr><br><br>`;

  // Segunda parte: Pliegos
  html += `<p><b><i><u style="color:black;">Pliegos:</u></i></b></p><br><br>`;

  datosPliegos.forEach((fila, i) => {
    const denunciadoNombre = fila[7]?.toString().toLowerCase();
    if (denunciadoNombre && denunciadoNombre.includes(nombre.toLowerCase())) {
      const expedientePliego = fila[0] || "";

      // Buscar expediente en entrada sanciones
      const matchEntrada = datosEntrada.find(e => (e[10] || "").toString().trim() === expedientePliego);
      if (matchEntrada) {
        const dni = matchEntrada[1] || "";
        const telefono = matchEntrada[2] || "";
        const lm = matchEntrada[3] || "";
        const unidad = matchEntrada[5] || "";
        const testigos = matchEntrada[9] || "";
        const resolucion = matchEntrada[11] || "";
        const exponeEntrada = matchEntrada[8] || "";

        html += `
          <p>
            Expediente nº: ${expedientePliego}, Denunciante: ${matchEntrada[0]}; DNI: ${dni}, Teléfono: ${telefono}, Lm: ${lm}, denuncia a la unidad: ${unidad} con testigos: ${testigos}, cuya resolución es: <u>${resolucion}</u><br>
            Expone que: ${exponeEntrada}<br>
        `;

        // Buscar en pliegos por expediente
        const denunciado = fila[7] || "";
        const lmPliego = fila[4] || "";
        const exponePliego = fila[6] || "";

        html += `<span style="color:blue">Denunciado: ${denunciado}, Lm: ${lmPliego}, Expone que: ${exponePliego}</span>`;
        html += `</p><br><br>`;
      }
    }
  });

  html += `</div>`;

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(700).setHeight(600),
    `Resultados de búsqueda`
  );
}

function limpiarEspaciosEnBlanco() { //Limpiar espacios en blanco

  // Obtener la hoja activa
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Obtener el rango de la columna seleccionada
  var columnaSeleccionada = hoja.getActiveRange().getColumn();
  
  // Obtener todas las celdas de esa columna (de la fila 2 en adelante)
  var rangoColumna = hoja.getRange(2, columnaSeleccionada, hoja.getLastRow() - 1, 1);
  
  // Obtener los valores de la columna seleccionada
  var valores = rangoColumna.getValues();
  
  // Iterar sobre cada celda de la columna
  for (var i = 0; i < valores.length; i++) {
    var valor = valores[i][0];
    
    // Verificar si la celda no está vacía y es texto
    if (typeof valor === 'string' && valor.trim() !== "") {
      // Reemplazar los espacios consecutivos por un solo espacio
      valores[i][0] = valor.replace(/\s+/g, ' ').trim();
    }
  }
  
  // Establecer los valores corregidos en el rango
  rangoColumna.setValues(valores);
  
  // Mostrar un mensaje de éxito
  SpreadsheetApp.getUi().alert('Los espacios duplicados han sido corregidos en la columna seleccionada.');
}
/*


// ver sancion completa en Registro entrada sanciones //
function mostrarDatosCeldaSeleccionadaSancion() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("Registro entrada sanciones");
  const celda = hoja.getActiveCell();
  const fila = celda.getRow();
  const columna = celda.getColumn();

  if (hoja.getName() !== "Registro entrada sanciones" || columna !== 13 || fila < 2) {
    SpreadsheetApp.getUi().alert("Selecciona una celda en la columna M (M2:M) de la hoja 'Registro entrada sanciones'.");
    return;
  }

  const valorM = hoja.getRange(fila, 13).getValue();
  const valorIraw = hoja.getRange(fila, 9).getValue();  // Fecha
  const valorJraw = hoja.getRange(fila, 10).getValue(); // Hora
  const valorC = hoja.getRange(fila, 3).getValue();
  const valorF = hoja.getRange(fila, 6).getValue();
  const valorK = hoja.getRange(fila, 11).getValue();
  const valorR = hoja.getRange(fila, 18).getValue();
  const valorS = hoja.getRange(fila, 19).getValue();

  const zona = ss.getSpreadsheetTimeZone();
  const valorI = Utilities.formatDate(new Date(valorIraw), zona, "dd/MM/yyyy");
  const valorJ = Utilities.formatDate(new Date(valorJraw), zona, "HH:mm");

  let mensaje = "";

  // Encabezado en "negrita, cursiva, subrayado"
  mensaje += `Expediente: ${valorM}, Fecha: ${valorI}, Hora: ${valorJ}.\n\n`;
  mensaje += `Denunciante: ${valorC}, lm: ${valorF}.\n`;

  mensaje += `${valorK}\n`;

  // Buscar en Registro pliegos
  const hojaPliegos = ss.getSheetByName("Registro pliegos");
  const datosPliegos = hojaPliegos.getRange("C2:C" + hojaPliegos.getLastRow()).getValues();
  const indexPliego = datosPliegos.findIndex(row => row[0] == valorM);

  if (indexPliego !== -1) {
    const filaPliego = indexPliego + 2;
    const valJ = hojaPliegos.getRange(filaPliego, 10).getValue();
    const valG = hojaPliegos.getRange(filaPliego, 7).getValue();
    const valI = hojaPliegos.getRange(filaPliego, 9).getValue();

    mensaje += `\n Denunciado:  ${valJ}, lm: ${valG}\n${valI}\n`;
  }

  // Buscar en Testigos
  const hojaTestigos = ss.getSheetByName("Testigos");
  const datosTestigos = hojaTestigos.getRange("C2:C" + hojaTestigos.getLastRow()).getValues();
  const indexTestigo = datosTestigos.findIndex(row => row[0] == valorM);

  if (indexTestigo !== -1) {
    const filaTestigo = indexTestigo + 2;
    const valJ = hojaTestigos.getRange(filaTestigo, 10).getValue();
    const valG = hojaTestigos.getRange(filaTestigo, 7).getValue();
    const valI = hojaTestigos.getRange(filaTestigo, 9).getValue();

    mensaje += `Testigo: ${valJ}, lm: ${valG}\n${valI}.\n`;
  }

  if (!valorR) {
    mensaje += `\n***__Falta firmar__***\n`;
  } else if (valorR && !valorS) {
    mensaje += `\n***__falta pliego__***\n`;
  }

  SpreadsheetApp.getUi().alert(mensaje);
}
*/

//VENTANA SANCION//
function generarContenidoMensaje() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("Registro entrada sanciones");
  const fila = parseInt(PropertiesService.getScriptProperties().getProperty("filaActiva"), 10);

  const valorM = hoja.getRange(fila, 13).getValue();
  const valorIraw = hoja.getRange(fila, 9).getValue();
  const valorJraw = hoja.getRange(fila, 10).getValue();
  const valorC = hoja.getRange(fila, 3).getValue();
  const valorF = hoja.getRange(fila, 6).getValue();
  const valorK = hoja.getRange(fila, 11).getValue();
  const valorR = hoja.getRange(fila, 18).getValue();
  const valorS = hoja.getRange(fila, 19).getValue();

  const zona = ss.getSpreadsheetTimeZone();
  const valorI = Utilities.formatDate(new Date(valorIraw), zona, "dd/MM/yyyy");
  const valorJ = Utilities.formatDate(new Date(valorJraw), zona, "HH:mm");

  let mensaje = "";
  mensaje += `<b>Expediente:</b> ${valorM}, <b>Fecha:</b> ${valorI}, <b>Hora:</b> ${valorJ}<br><br>`;
  mensaje += `<b>Denunciante:</b> ${valorC}, lm: ${valorF}<br>${valorK}<br>`;

  const hojaPliegos = ss.getSheetByName("Registro pliegos");
  const datosPliegos = hojaPliegos.getRange("C2:C" + hojaPliegos.getLastRow()).getValues();
  const indexPliego = datosPliegos.findIndex(row => row[0] == valorM);

  if (indexPliego !== -1) {
    const filaPliego = indexPliego + 2;
    const valJ = hojaPliegos.getRange(filaPliego, 10).getValue();
    const valG = hojaPliegos.getRange(filaPliego, 7).getValue();
    const valI = hojaPliegos.getRange(filaPliego, 9).getValue();
    mensaje += `<br><b>Denunciado:</b> ${valJ}, lm: ${valG}<br>${valI}<br>`;
  }

  const hojaTestigos = ss.getSheetByName("Testigos");
  const datosTestigos = hojaTestigos.getRange("C2:C" + hojaTestigos.getLastRow()).getValues();
  const indexTestigo = datosTestigos.findIndex(row => row[0] == valorM);

  if (indexTestigo !== -1) {
    const filaTestigo = indexTestigo + 2;
    const valJ = hojaTestigos.getRange(filaTestigo, 10).getValue();
    const valG = hojaTestigos.getRange(filaTestigo, 7).getValue();
    const valI = hojaTestigos.getRange(filaTestigo, 9).getValue();
    mensaje += `<br><b>Testigo:</b> ${valJ}, lm: ${valG}<br>${valI}.<br>`;
  }

  if (!valorR) {
    mensaje += `<br><b><i>Falta firmar</i></b><br>`;
  } else if (valorR && !valorS) {
    mensaje += `<br><b><i>Falta pliego</i></b><br>`;
  }

  const hoy = new Date();
  const onceDiasEnMs = 11 * 24 * 60 * 60 * 1000;
  const fechaLimite = new Date(hoy.getTime() - onceDiasEnMs);

  if (!valorS && valorR instanceof Date && valorR < fechaLimite) {
    const diasFaltantes = Math.ceil((valorR.getTime() + onceDiasEnMs - hoy.getTime()) / (1000 * 60 * 60 * 24));
    mensaje += `<br><b>🕒 Todavía no se ha cumplido el plazo, faltan ${diasFaltantes} días.</b><br><br>`;
  }

  return mensaje;
}
function mostrarDatosCeldaSeleccionadaSancion() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("Registro entrada sanciones");
  const celda = hoja.getActiveCell();
  const fila = celda.getRow();
  const columna = celda.getColumn();

  if (hoja.getName() !== "Registro entrada sanciones" || columna !== 13 || fila < 2) {
    SpreadsheetApp.getUi().alert("Selecciona una celda en la columna M (M2:M) de la hoja 'Registro entrada sanciones'.");
    return;
  }

  // Guardar fila activa para usar después
  PropertiesService.getScriptProperties().setProperty("filaActiva", fila.toString());

  const html = HtmlService.createHtmlOutputFromFile('ventanaSancion')
    .setWidth(800)
    .setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(html, "Información del expediente");
}
function copiarYActualizar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaSanciones = ss.getSheetByName('Registro entrada sanciones');
  const hojaResolucion = ss.getSheetByName('Resolución');

  const fila = parseInt(PropertiesService.getScriptProperties().getProperty("filaActiva"), 10);
  const valorM = hojaSanciones.getRange(`M${fila}`).getValue();

  hojaResolucion.getRange('F3').setValue(valorM);

  updateResolutionFields();
}
//FIN ventana Sancion//

function IntroducirCantidadAPagarDesdeFormulario(cantidad) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Resolución"); 
  const celdaDestino = hoja.getRange("C136");

  // Usar el valor recibido desde el HTML (en vez de la celda I132)
  celdaDestino.setValue(cantidad);

  // Cambiar el fondo de la celda a blanco
  celdaDestino.setBackground("#FFFFFF");

  // Aplicar bordes blancos
  celdaDestino.setBorder(true, true, true, true, true, true, "#FFFFFF", SpreadsheetApp.BorderStyle.SOLID);

  // Limpiar otras celdas
  hoja.getRange("D136").setValue("");
  hoja.getRange("E136").setValue("");
  hoja.getRange("I132").setValue(""); // Esto lo puedes quitar si ya no usas I132
}


function BorrarResolucion() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Lista de celdas a borrar
  var celdas = ["A8", "C6", "E5", "E4", "F3", "A61", "C60", "F56", "C2", "C54", "A117", "B117", "A120", "A123", "E133", "C136", "B147"];

  // Recorremos y vaciamos cada celda
  celdas.forEach(function(celda) {
    hoja.getRange(celda).setValue("");
  });
  hoja.getRange("F3").setFontColor("red");
  hoja.getRange("E111").setValue(new Date());
  Logger.log("🧹 Celdas limpiadas: " + celdas.join(", "));
}


function eliminarEspaciosEnBlanco() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Procesar cada hoja
  limpiarColumna(ss.getSheetByName("Registro entrada sanciones"), "E", 2);
  limpiarColumna(ss.getSheetByName("Registro pliegos"), "F", 2);
  limpiarColumna(ss.getSheetByName("Testigos"), "G", 2);
}

function limpiarColumna(hoja, letraColumna, filaInicio) {
  if (!hoja) return;
  
  var ultimaFila = hoja.getLastRow();
  if (ultimaFila < filaInicio) return;
  
  var rango = hoja.getRange(`${letraColumna}${filaInicio}:${letraColumna}${ultimaFila}`);
  var valores = rango.getValues();
  
  for (var i = 0; i < valores.length; i++) {
    if (valores[i][0] !== "" && typeof valores[i][0] === "string") {
      // Elimina todos los espacios en blanco, incluso los del medio
      valores[i][0] = valores[i][0].replace(/\s+/g, '');
    }
  }
  
  rango.setValues(valores);
}
function marcarNoFirmado() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registro entrada sanciones");
  if (!hoja) return;

  var ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return; // No hay datos

  // Obtener valores de las columnas N y R
  var rangoN = hoja.getRange("N2:N" + ultimaFila);
  var rangoR = hoja.getRange("R2:R" + ultimaFila);
  var valoresN = rangoN.getValues();
  var valoresR = rangoR.getValues();

  for (var i = 0; i < valoresR.length; i++) {
    var valorN = valoresN[i][0];
    var valorR = valoresR[i][0];

    // 1. Si N = "NO FIRMADO" y R tiene datos → borrar N
    if (valorN === "NO FIRMADO" && valorR !== "" && valorR != null) {
      hoja.getRange(i + 2, 14).setValue(""); // Columna N
    }

    // 2. Si R está vacío → N = "NO FIRMADO"
    else if ((valorR === "" || valorR == null)) {
      hoja.getRange(i + 2, 14).setValue("NO FIRMADO");
    }
  }
}

//*******************************INFO DENUNCIANTE///////////////////////////////////////// */

function mostrarResolucion() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaResol = ss.getSheetByName("Resolución");
  const hojaRegistro = ss.getSheetByName("Registro entrada sanciones");
  const hojaDirectiva = ss.getSheetByName("Directiva");
  const hojaPliegos = ss.getSheetByName("Registro pliegos");

  // === 1. Buscar expediente en hoja Resolución ===
  const valorF3 = hojaResol.getRange("F3").getValue().toString().trim().toLowerCase();

  const datosRegistro = hojaRegistro.getRange("M2:P" + hojaRegistro.getLastRow()).getValues();
  let html = "<div style='font-family:Arial; font-size:13px;'>";

  for (let i = 0; i < datosRegistro.length; i++) {
    const expediente = (datosRegistro[i][0] || "").toString().trim().toLowerCase(); // Col M
    if (expediente.includes(valorF3)) {
      const fila = i + 2;
      const nombre = hojaRegistro.getRange("C" + fila).getValue();
      const dni = hojaRegistro.getRange("D" + fila).getValue();
      const testigos = hojaRegistro.getRange("L" + fila).getValue();
      const contenido = hojaRegistro.getRange("K" + fila).getValue();
      const enlace = hojaRegistro.getRange("P" + fila).getValue();
      const fecha = hojaRegistro.getRange("I" + fila).getValue();

      html += `<div style="background:#f5f5dc; padding:10px; margin:8px; border-radius:6px;">
        <b>Denuncia:</b><br>
        Expediente: ${expediente}<br>
        Nombre: ${nombre}<br>
        DNI: ${dni}<br>
        Testigos: ${testigos}<br>
        Contenido: ${contenido}<br>
        Pruebas: <img src="${enlace}" style="max-width:200px;"><br>
        Fecha de los hechos: ${fecha}<br>`;

      // === 2. Buscar resolución en Directiva ===
      const datosDirectiva = hojaDirectiva.getRange("A2:D" + hojaDirectiva.getLastRow()).getValues();
      for (let j = 0; j < datosDirectiva.length; j++) {
        if ((datosDirectiva[j][0] || "").toString().trim().toLowerCase() === expediente) {
          html += `Resolución: ${datosDirectiva[j][3]}<br><hr>`;
        }
      }
      html += "</div>";
    }
  }

  // === 3. Buscar en Registro Pliegos coincidencias por nombre o DNI (90%) ===
  const datosPliegos = hojaPliegos.getRange("C2:K" + hojaPliegos.getLastRow()).getValues();
  html += `<div style="background:#f9f9f9; padding:10px; margin:8px; border-radius:6px;">`;

  for (let i = 0; i < datosPliegos.length; i++) {
    const dniPliego = (datosPliegos[i][1] || "").toString().toLowerCase(); // Col D
    const nombrePliego = (datosPliegos[i][7] || "").toString().toLowerCase(); // Col J
    if (similarity(nombrePliego, valorF3) >= 0.9 || similarity(dniPliego, valorF3) >= 0.9) {
      const fila = i + 2;
      const dni = hojaPliegos.getRange("D" + fila).getValue();
      const contenido = hojaPliegos.getRange("I" + fila).getValue();
      const enlace = hojaPliegos.getRange("K" + fila).getValue();
      const expedientePliego = hojaPliegos.getRange("C" + fila).getValue();

      html += `<b>Denunciado:</b> ${nombrePliego}<br>
        DNI: ${dni}<br>
        Contenido: ${contenido}<br>
        Enlace: <a href="${enlace}" target="_blank">${enlace}</a><br>`;

      // Buscar expediente en Registro entrada sanciones
      for (let r = 0; r < datosRegistro.length; r++) {
        if ((datosRegistro[r][0] || "").toString().toLowerCase() === expedientePliego.toString().toLowerCase()) {
          html += `Fecha: ${datosRegistro[r][2]}<br>`; // Col I
        }
      }

      // Buscar en Directiva
      const datosDirectiva2 = hojaDirectiva.getRange("A2:D" + hojaDirectiva.getLastRow()).getValues();
      for (let j = 0; j < datosDirectiva2.length; j++) {
        if ((datosDirectiva2[j][0] || "").toString().toLowerCase() === expedientePliego.toString().toLowerCase()) {
          html += `Resolución: ${datosDirectiva2[j][3]}<br>`;
        }
      }
      html += "<hr>";
    }
  }
  html += "</div>";

  // Mostrar ventana
  const output = HtmlService.createHtmlOutput(html)
    .setWidth(600)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(output, "Resultados de Resolución");
}

// === Función para calcular similitud entre strings ===
function similarity(s1, s2) {
  if (!s1 || !s2) return 0;
  const longer = s1.length > s2.length ? s1 : s2;
  const shorter = s1.length > s2.length ? s2 : s1;
  const longerLength = longer.length;
  const editDist = editDistance(longer, shorter);
  return (longerLength - editDist) / parseFloat(longerLength);
}

function editDistance(s1, s2) {
  s1 = s1.toLowerCase();
  s2 = s2.toLowerCase();
  const costs = [];
  for (let i = 0; i <= s1.length; i++) {
    let lastValue = i;
    for (let j = 0; j <= s2.length; j++) {
      if (i === 0) costs[j] = j;
      else {
        if (j > 0) {
          let newValue = costs[j - 1];
          if (s1.charAt(i - 1) !== s2.charAt(j - 1)) {
            newValue = Math.min(Math.min(newValue, lastValue),
              costs[j]) + 1;
          }
          costs[j - 1] = lastValue;
          lastValue = newValue;
        }
      }
    }
    if (i > 0) costs[s2.length] = lastValue;
  }
  return costs[s2.length];
}


//*********************⚠️ FORMULARIO SANCIONES  ⚠️++++++++++++++++++++++++++++ */

// Crear menú en la hoja
/*function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("📌 Registro")
    .addItem("Nuevo registro", "abrirFormulario")
    .addToUi();
}*/

// Abre el formulario en modal
function abrirFormulario() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheetByName("Registro entrada sanciones").activate();
  
  const html = HtmlService.createHtmlOutputFromFile("formulario")
    .setWidth(700)
    .setHeight(1100);
  SpreadsheetApp.getUi().showModalDialog(html, "Nuevo Registro de Sanción");
}
/*                                                                            ***** Cambio realizado por Kapi *****
// Guardar archivo en Drive y devolver enlace
function subirArchivo(obj) {
  try {
    const carpetaId = "1u2aQ8ZuXYSivLmfodcqWqVuT4-vicNLK"; // 👉 pon aquí el ID de la carpeta de destino en Drive
    const carpeta = DriveApp.getFolderById(carpetaId);
    const blob = Utilities.newBlob(obj.datos, obj.tipo, obj.nombre);
    const archivo = carpeta.createFile(blob);
    archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return archivo.getUrl();
  } catch (e) {
    return "Error al subir archivo: " + e.message;
  }
}
*/
// Procesar datos al enviar
function procesarFormulario(datos) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registro entrada sanciones");
  
  // Insertar fila arriba de la fila 2
  hoja.insertRowBefore(2);
  
  // Valores con placeholder por defecto
  const correo = datos.correo || "instructores.taxiarona@gmail.com";
  const nombre = datos.nombre || "DIRECTIVA";
  const telefono = datos.telefono || "922747511";
  const dni = datos.dni || "G38314480";
  const lm = datos.lm || "DIRECTIVA";
  const matricula=datos.matricula;
  const lmDenunciado = datos.lmDenunciado || "";
  const fecha = datos.fecha || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const hora = datos.hora || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm");
  const contenido = datos.contenido || "SIN CONTENIDO";
  const testigos = datos.testigos || "";
  const archivo = datos.archivo || "";

  // Calcular nuevo valor de M2 (M3 +1 antes de "/")
  let m3 = hoja.getRange(3, 13).getValue(); // Columna M fila 3
  let nuevoM = "";
  if (m3 && typeof m3 === "string" && m3.includes("/")) {
    let [num, resto] = m3.split("/");
    nuevoM = (parseInt(num, 10) + 1) + "/" + resto;
  } else {
    nuevoM = "1/0";
  }
  
  // Escribir datos en la nueva fila 2
  hoja.getRange("A2").setValue(new Date()); // Fecha y hora actual
  hoja.getRange("B2").setValue(correo);
  hoja.getRange("C2").setValue(nombre);
  hoja.getRange("D2").setValue(dni);
  hoja.getRange("E2").setValue(telefono);
  hoja.getRange("F2").setValue(lm);
  hoja.getRange("G2").setValue(matricula);
  hoja.getRange("H2").setValue(lmDenunciado);
  hoja.getRange("I2").setValue(fecha);
  hoja.getRange("J2").setValue(hora);
  hoja.getRange("K2").setValue(contenido);
  hoja.getRange("L2").setValue(testigos);
  hoja.getRange("M2").setValue(nuevoM);
  hoja.getRange("P2").setValue(archivo);
}


//********************* ✅ 📄 FORMULARIO PLIEGOS 🙋++++++++++++++++++++++++++++ */
function abrirFormularioPliegoConCopia() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaOrigen = ss.getSheetByName("Registro entrada sanciones");
  var celda = hojaOrigen.getActiveCell();
  var fila = celda.getRow();
  var columna = celda.getColumn();

  // Solo permitir columna M (13) desde fila 2
  if (columna === 13 && fila >= 2) {
    var contenidoCopiado = celda.getValue(); // valor de la celda seleccionada
    // Abrir formulario y pasar el contenido
    abrirFormularioPliego(contenidoCopiado);
  } else {
    SpreadsheetApp.getUi().alert("Debes seleccionar un expediente en la columna M desde la fila 2");
  }
}

/*function copiarExpediente() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaOrigen = ss.getSheetByName("Registro entrada sanciones");
  var hojaDestino = ss.getSheetByName("Registro Pliegos");
  
  var celda = hojaOrigen.getActiveCell();
  var fila = celda.getRow();
  var columna = celda.getColumn();
  
  // Verificar si la celda está en la columna M (13) y desde la fila 2
  if (columna === 13 && fila >= 2) {
    var valor = celda.getValue(); // Obtener valor de la celda
    hojaDestino.activate();        // Activar hoja destino
    
    // Pegar valor en la misma celda activa de la hoja destino
    hojaDestino.getActiveCell().setValue(valor);
    abrirFormularioPliego();
    
  } else {
    SpreadsheetApp.getUi().alert("Debes seleccionar un expediente");
  }
  //abrirFormularioPliego();
}*/


/* Crear menú al abrir la hoja
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("📌 Registro")
    .addItem("Nuevo registro", "abrirFormulario")
    .addToUi();
}*/

/* Abrir la ventana modal con el formulario
function abrirFormularioPliego() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheetByName("Registro pliegos").activate();
  
  const html = HtmlService.createHtmlOutputFromFile("pliego")
    .setWidth(500)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, "Nuevo Registro de Pliego");
}*/
  
  function abrirFormularioPliego(contenidoCopiado) {
  const html = HtmlService.createHtmlOutputFromFile("pliego")
    .setWidth(500)
    .setHeight(1200);

  // Pasamos el valor de la celda al HTML
  html.append(`<script>var contenidoInicial = ${JSON.stringify(contenidoCopiado)};</script>`);

  SpreadsheetApp.getUi().showModalDialog(html, "Nuevo Registro de Pliego");
}

  
// Subir archivo a Drive y devolver el enlace
function subirArchivo(obj) {
  try {
    const carpetaId = "1u2aQ8ZuXYSivLmfodcqWqVuT4-vicNLK"; // Reemplaza con el ID de tu carpeta en Drive
    const carpeta = DriveApp.getFolderById(carpetaId);
    const blob = Utilities.newBlob(obj.datos, obj.tipo, obj.nombre);
    const archivo = carpeta.createFile(blob);
    archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return archivo.getUrl();
  } catch (e) {
    return "Error al subir archivo: " + e.message;
  }
}


// Procesar los datos del formulario de "Registro pliegos"
function procesarFormularioPliegos(datos) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registro pliegos");
  
  // Insertar fila encima de la fila 2
  hoja.insertRowBefore(2);
  
  // Valores con placeholder por defecto
  const correo = datos.correo || "instructores.taxiarona@gmail.com";
  const nombre = datos.nombre || "";
  const telefono = datos.telefono || "";
  const dni = datos.dni || "";
  const lm = datos.lm || "";
  const contenido = datos.contenido || "";
  const matricula = datos.matricula || "";
  const archivo = datos.archivo || "";
  
  // Insertar valores en fila 2
  hoja.getRange("A2").setValue(new Date()); // Fecha y hora actual
  hoja.getRange("B2").setValue(correo);
  hoja.getRange("C2").setValue(datos.copiado || ""); // contenido copiado que quieras pegar
  hoja.getRange("D2").setValue(dni);
  hoja.getRange("F2").setValue(telefono);
  hoja.getRange("G2").setValue(lm);
  hoja.getRange("H2").setValue(matricula);
  hoja.getRange("I2").setValue(contenido);
  hoja.getRange("J2").setValue(nombre);
  hoja.getRange("K2").setValue(archivo);
}


//*****************autocompletar pliego de descargo******************************** */


function doGet() {
  return HtmlService.createHtmlOutput("Método GET no permitido");
}

function doPost(e) {
  return HtmlService.createHtmlOutput("Método POST no permitido");
}

// 🔎 Buscar en hojas
function buscarPersona(nombreBuscado) {
  nombreBuscado = quitarAcentos(nombreBuscado).toLowerCase().trim();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaSanciones = ss.getSheetByName("Registro entrada sanciones");
  const hojaPliegos = ss.getSheetByName("Registro pliegos");

  // --- Buscar en sanciones ---
  const datosSanciones = hojaSanciones.getRange("B2:G" + hojaSanciones.getLastRow()).getValues();
  for (let i = 0; i < datosSanciones.length; i++) {
    let nombre = quitarAcentos(datosSanciones[i][1]).toLowerCase().trim(); // Columna C
    if (nombre === nombreBuscado) {
      return {
        correo: datosSanciones[i][0], // Col B
        dni: datosSanciones[i][2],    // Col D
        telefono: datosSanciones[i][3], // Col E
        matricula: datosSanciones[i][5] // Col G
      };
    }
  }

  // --- Buscar en pliegos ---
  const datosPliegos = hojaPliegos.getRange("B2:H" + hojaPliegos.getLastRow()).getValues();
  for (let i = 0; i < datosPliegos.length; i++) {
    let nombre = quitarAcentos(datosPliegos[i][8]).toLowerCase().trim(); // Col J
    if (nombre === nombreBuscado) {
      return {
        correo: datosPliegos[i][0],   // Col B
        dni: datosPliegos[i][2],      // Col D
        telefono: datosPliegos[i][4], // Col F
        matricula: datosPliegos[i][6] // Col H
      };
    }
  }

  return null; // no encontrado
}

// Función para quitar acentos
function quitarAcentos(str) {
  return str.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}


//+++++++++++++++++buscar en formulario sanciones por histórico++++++++++++++//

function normalizarTexto(texto) {
  if (!texto) return "";
  return texto
    .toString()
    .normalize("NFD") // elimina acentos
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-zA-Z0-9\s]/g, "") // elimina símbolos
    .toLowerCase()
    .trim();
}

function buscarDatosPorNombre(nombre) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nombreNorm = normalizarTexto(nombre);

  // 1️⃣ Buscar en "Registro entrada sanciones"
  const hoja1 = ss.getSheetByName("Registro entrada sanciones");
  const datos1 = hoja1.getRange("B2:G" + hoja1.getLastRow()).getValues();
  for (let i = 0; i < datos1.length; i++) {
    const nombreFila = normalizarTexto(datos1[i][1]); // Columna C (índice 1 en rango B:G)
    if (nombreFila.includes(nombreNorm)) {
      return {
        correo: datos1[i][0],    // B
        dni: datos1[i][2],       // D
        telefono: datos1[i][3],  // E
        lm: datos1[i][4],        // F
        matricula: datos1[i][5]  // G
      };
    }
  }

  // 2️⃣ Buscar en "Registro pliegos"
  const hoja2 = ss.getSheetByName("Registro pliegos");
  const datos2 = hoja2.getRange("B2:H" + hoja2.getLastRow()).getValues();
  for (let i = 0; i < datos2.length; i++) {
    const nombreFila = normalizarTexto(datos2[i][8 - 2]); // Columna J → índice 7 en rango B:H
    if (nombreFila.includes(nombreNorm)) {
      return {
        correo: datos2[i][0],    // B
        dni: datos2[i][2],       // D
        telefono: datos2[i][4],  // F
        lm: datos2[i][5],        // G
        matricula: datos2[i][6]  // H
      };
    }
  }

  return null;
}

//*****************buscar por socio en formulario socios++++++++++++++++++++++++ */

// 🔹 Buscar socio en la hoja "Socios" por LM
function buscarSocioPorLM(lm) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("Socios");
  const datos = hoja.getRange("a2:j" + hoja.getLastRow()).getValues(); // B hasta J

  for (let i = 0; i < datos.length; i++) {
    const lmFila = datos[i][0]; // Columna B
    if (lmFila && lmFila.toString().trim() === lm.toString().trim()) {
      return {
        nombre: datos[i][1],     // B (Nombre y Apellidos según lo que pediste)
        dni: datos[i][2],        // C
        telefono: datos[i][6],   // G
        correo: datos[i][7],     // H
        matricula: datos[i][8]   // I
      };
    }
  }
  return null;
}

// 🔹 Corregir datos en la hoja "Socios"
function corregirSocio(datos) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName("Socios");
  const ultimaFila = hoja.getLastRow();
  const rango = hoja.getRange("a2:j" + ultimaFila).getValues();

  for (let i = 0; i < rango.length; i++) {
    const lmFila = rango[i][0]; // Columna B = LM
    if (lmFila && lmFila.toString().trim() === datos.lm.toString().trim()) {
      const fila = i + 2; // porque empezamos en fila 2

      hoja.getRange(fila, 2).setValue(datos.nombre).setFontColor("blue");   // Col B
      hoja.getRange(fila, 3).setValue(datos.dni).setFontColor("blue");      // Col C
      hoja.getRange(fila, 7).setValue(datos.telefono).setFontColor("blue"); // Col G
      hoja.getRange(fila, 8).setValue(datos.correo).setFontColor("blue");   // Col H
      hoja.getRange(fila, 9).setValue(datos.matricula).setFontColor("blue");// Col I

      // Columna J → fecha actual
      hoja.getRange(fila, 10).setValue(new Date()).setFontColor("blue");

      return true;
    }
  }
  return false;
}

function getAutocompleteData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaActiva = ss.getActiveSheet();
  const hojaRegistro = ss.getSheetByName("Registro pliegos");

  const datos1 = hojaActiva.getRange("C2:C" + hojaActiva.getLastRow()).getValues().flat().filter(String);
  const datos2 = hojaRegistro.getRange("J2:J" + hojaRegistro.getLastRow()).getValues().flat().filter(String);

  const datosUnicos = [...new Set([...datos1, ...datos2])];
  return datosUnicos;
}