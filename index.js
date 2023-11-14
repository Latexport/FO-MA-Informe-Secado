document.addEventListener("DOMContentLoaded", () => {
  console.log("INICIANDO SCRIPT");

  document.getElementById("hojacalculo").addEventListener("submit", (event: Event) => {
    console.log("SE PRECIONO EL BOTON");

    event.preventDefault(); // Evita que se envíe el formulario de forma predeterminada

    submit();
  });
});

async function submit() {
  console.log("submit()");

  // Obtiene los valores ingresados en el formulario
  const datosExtraidos = extraerDatos();
  console.log("LOS DATOS SON : ", datosExtraidos);
  // Establece los valores en las celdas correspondientes
 console.log();
  const data = [
    [
      datosExtraidos.fechaProduccion,
      datosExtraidos.fechaSecado,
      datosExtraidos.secador,
      datosExtraidos.auxiliares,
      datosExtraidos.referencia,
      datosExtraidos.referenciaExtraida,
      datosExtraidos.turno,
      datosExtraidos.lote,
      datosExtraidos.maquina,
      datosExtraidos.registro,
      datosExtraidos.reproceso,
      datosExtraidos.tipoReproceso,
      datosExtraidos.anteriorRegistro,
      datosExtraidos.temperatura,
      datosExtraidos.tiempoSecado,
      datosExtraidos.tiempoAdicional,
      datosExtraidos.tiempoEnfriamiento,
      datosExtraidos.silicona,
      datosExtraidos.pesoSeco,
      datosExtraidos.unidades,
      datosExtraidos.unidadesTeoricas,
      datosExtraidos.diferencia,
      datosExtraidos.consumoMezclas,
      datosExtraidos.observaciones,
      datosExtraidos.totalTiempo,
      datosExtraidos.totalTiempoMinimo
    ]
  ];

  // Nombre de la hoja en la que deseas guardar los datos

  // Guarda los datos en Excel en la hoja especificada
  const nombreHoja = "seguimiento";
  await agregarDatosExcel(nombreHoja, data);
  console.log("TERMINAMOS DE AGREGAR LOS DATOS A LA TABLA");
}

function obtenerLetraColumnaDesdeNumero(numero) {
  if (numero >= 1 && numero <= 1) {
    // 18278 es la cantidad máxima de letras de columna en Excel (26^2 + 26)
    let letra = "";
    while (numero > 0) {
      const modulo = (numero - 1) % 26;
      letra = String.fromCharCode(65 + modulo) + letra; // 65 es el código ASCII de 'A'
      numero = Math.floor((numero - 1) / 26);
    }
    return letra;
  } else {
    return "Número fuera de rango";
  }
}

async function agregarDatosExcel(nombreHoja, data) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(nombreHoja);

      // Obtén el rango donde deseas insertar los datos (puede ser cualquier rango deseado)

      // Carga el rango de destino y sincroniza el contexto
      let fila = await numeroFila(sheet, context);

      console.log("lastRow1", fila);
      console.log("range", `A:${fila}:Z${fila}`);
      let columnaLetra = obtenerLetraColumnaDesdeNumero(data[0].length);
      // Ajusta esta ubicación según tus necesidades
      let dataRange = sheet.getRange(`A${fila}:${columnaLetra}${fila}`);
      //  dataRange.values = data;
      dataRange.values = data;

      return context.sync();
    });
  } catch (error) {
    console.log("Error al agregar datos a Excel:", error);
  }
}

async function numeroFila(sheet, context) {
  const range1 = sheet.getRange("A:A").getUsedRange();
  range1.load("rowIndex");
  range1.load("rowCount");
  await context.sync();

  const lastRow1 = range1.rowIndex + range1.rowCount;

  return lastRow1 + 1;
}

function extraerDatos() {
  try {
    console.log("EXTRAEMOS LOS DATOS");

    const fechaProduccion= (document.getElementById("fechaProduccion") as HTMLInputElement).value;
    console.log("fechaProduccion", fechaProduccion);
  

    const fechaSecado = (document.getElementById("fechaSecado") as HTMLInputElement).value;
    console.log("fechaSecado", fechaSecado);

    const secador = (document.getElementById("secador") as HTMLSelectElement).value;
    console.log("secador", secador);

    const auxiliares = (document.getElementById("auxiliares") as HTMLSelectElement).value;
    console.log("auxiliares", auxiliares);

    const referencia = (document.getElementById("referencia") as HTMLSelectElement).value;
    console.log("referencia", referencia);

    const referenciaExtraida = (document.getElementById("referenciaExtraida") as HTMLInputElement).value;
    console.log("referenciaExtraida", referenciaExtraida);

    const turno = (document.getElementById("turno") as HTMLSelectElement).value;
    console.log("turno", turno);

    const lote = (document.getElementById("lote") as HTMLInputElement).value;
    console.log("lote", lote);

    const maquina = (document.getElementById("maquina") as HTMLSelectElement).value;
    console.log("maquina", maquina);

    const registro = (document.getElementById("registro") as HTMLInputElement).value;
    console.log("registro", registro);

    const reproceso = (document.getElementById("reproceso") as HTMLSelectElement).value;
    console.log("reproceso", reproceso);

    const tipoReproceso = (document.getElementById("tipoReproceso") as HTMLSelectElement).value;
    console.log("tipoReproceso", tipoReproceso);

    const anteriorRegistro = (document.getElementById("anteriorRegistro") as HTMLInputElement).value;
    console.log("anteriorRegistro", anteriorRegistro);

    const temperatura = (document.getElementById("temperatura") as HTMLInputElement).value;
    console.log("temperatura", temperatura);

    const tiempoSecado = (document.getElementById("tiempoSecado") as HTMLInputElement).value;
    console.log("tiempoSecado", tiempoSecado);

    const tiempoAdicional = (document.getElementById("tiempoAdicional") as HTMLInputElement).value;
    console.log("tiempoAdicional", tiempoAdicional);

    const tiempoEnfriamiento = (document.getElementById("tiempoEnfriamiento") as HTMLInputElement).value;
    console.log("tiempoEnfriamiento", tiempoEnfriamiento);
    console.log("si esta bien 1.1");
    const silicona = (document.getElementById("silicona") as HTMLInputElement).value;
    console.log("silicona", silicona);

    const pesoSeco = (document.getElementById("pesoSeco") as HTMLInputElement).value;
    console.log("pesoSeco", pesoSeco);

    const unidades = (document.getElementById("unidades") as HTMLInputElement).value;
    console.log("unidades", unidades);

    const unidadesTeoricas = (document.getElementById("unidadesTeoricas") as HTMLInputElement).value;
    console.log("unidadesTeoricas", unidadesTeoricas);

    const diferencia = (document.getElementById("diferencia") as HTMLInputElement).value;
    console.log("diferencia", diferencia);

    const consumoMezclas = (document.getElementById("consumoMezclas") as HTMLInputElement).value;
    console.log("registro", registro);

    const observaciones = (document.getElementById("observaciones") as HTMLSelectElement).value;
    console.log("observaciones", observaciones);

    const totalTiempo = (document.getElementById("totalTiempo") as HTMLInputElement).value;
    console.log("totalTiempo", totalTiempo);

    const totalTiempoMinimo = (document.getElementById("totalTiempoMinimo") as HTMLInputElement).value;
    console.log("totalTiempoMinimo", totalTiempoMinimo);

    
    const data = {
      fechaProduccion,
      fechaSecado,
      secador,
      auxiliares,
      referencia,
      referenciaExtraida,
      turno,
      lote,
      maquina,
      registro,
      reproceso,
      tipoReproceso,
      anteriorRegistro,
      temperatura,
      tiempoSecado,
      tiempoAdicional,
      tiempoEnfriamiento,
      silicona,
      pesoSeco,
      unidades,
      unidadesTeoricas,
      diferencia,
      consumoMezclas,
      observaciones,
      totalTiempo,
      totalTiempoMinimo
    };
    console.log(data);
    return data;
  } catch (error) {
    console.log("error al recojer  la data del formulario");
    console.log(error);
  }
}

console.log("Script ejecutándose");

document.addEventListener("DOMContentLoaded", function() {
  var maquinaSelect = document.getElementById("maquina");
  var temperaturaInput = document.getElementById("temperatura") as HTMLInputElement;

  if (maquinaSelect) {
    maquinaSelect.addEventListener("change", function() {
      // Obtener el valor seleccionado de la máquina
      var maquinaSeleccionada = maquinaSelect.value;

      // Realizar la búsqueda en el texto
      var textoDeBusqueda = "18"; // Puedes ajustar esto según tu necesidad
      var temperaturaEncontrada = null;

      // Ejemplo de búsqueda
      if (maquinaSeleccionada === "2" && textoDeBusqueda === "18") {
        temperaturaEncontrada = "70";
      }

      // Actualizar el campo de temperatura
      temperaturaInput.value = temperaturaEncontrada !== null ? temperaturaEncontrada : "";
    });

    // Ahora puedes usar maquinaSelect fuera de la función de cambio
    var maquinaSeleccionadaFuera = maquinaSelect.value;
    console.log("Maquina seleccionada fuera de la función:", maquinaSeleccionadaFuera);
  } else {
    console.error("Elemento 'maquina' no encontrado en el documento.");
  }
});



document.addEventListener("DOMContentLoaded", function() {
  var maquinaSelect = document.getElementById("maquina");
  var temperaturaInput = document.getElementById("temperatura") as HTMLInputElement;

  if (maquinaSelect) {
    maquinaSelect.addEventListener("change", function() {
      // Obtener el valor seleccionado de la máquina
      var maquinaSeleccionada = maquinaSelect.value;

      // Realizar la búsqueda en el texto
      var textoDeBusqueda = "18"; // Puedes ajustar esto según tu necesidad
      var temperaturaEncontrada = null;

      // Ejemplo de búsqueda
      if (maquinaSeleccionada === "2" && textoDeBusqueda === "18") {
        temperaturaEncontrada = "70";
      }

      // Actualizar el campo de temperatura
      temperaturaInput.value = temperaturaEncontrada !== null ? temperaturaEncontrada : "";
    });
  } else {
    console.error("Elemento 'maquina' no encontrado en el documento.");
  }
});
 

