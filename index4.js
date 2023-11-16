document.addEventListener("DOMContentLoaded", () => {
  console.log("INICIANDO SCRIPT");

  document.getElementById("hojacalculo").addEventListener("submit", (event) => {
    console.log("SE PRECIONO EL BOTON");

    event.preventDefault(); // Evita que se envíe el formulario de forma predeterminada

    //ejecutamos la funcion principal
    submit();
  });
});

async function submit() {
  console.log("submit()");

  // Obtiene los valores ingresados en el formulario
  const datosExtraidos = extraerDatos();
  console.log("LOS DATOS SON : ", datosExtraidos);
  // Establece los valores en las celdas correspondientes

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
  if (numero >= 1) {
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

async function obtenerNumeroFila(sheet, context) {
  try {
    // Obtener todas las celdas en la columna A
    const dataRange = sheet.getRange("A:A");

    // Cargar las propiedades de la celda y sincronizar el contexto
    dataRange.load("values");
    await context.sync();

    // Verificar si dataRange.values es null o undefined
    if (!dataRange.values) {
      // Si es null, devolver la primera fila (no hay datos en la columna A)
      return 1;
    }

    // Encontrar la última fila no vacía en la columna A
    let ultimaFilaNoVacia = dataRange.values.length;
    while (ultimaFilaNoVacia > 0 && dataRange.values[ultimaFilaNoVacia - 1][0] === "") {
      ultimaFilaNoVacia--;
    }

    // Devolver la siguiente fila después de la última no vacía
    return ultimaFilaNoVacia + 1;
  } catch (error) {
    console.error("Error al obtener el número de fila:", error);
    throw error;
  }
}

async function agregarDatosExcel(nombreHoja, data) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(nombreHoja);

      // Fila específica a la que quieres agregar los datos (19377 en este caso)
      const fila = 19377;

      console.log("Intentando agregar datos en la fila:", fila);

      let columnaLetra = obtenerLetraColumnaDesdeNumero(data[0].length);

      // Ajusta esta ubicación según tus necesidades
      let dataRange = sheet.getRange(`A${fila}:${columnaLetra}${fila}`);
      dataRange.values = data;

      return context.sync();
    });
  } catch (error) {
    console.log("Error al agregar datos a Excel:", error);
  }
}

function extraerDatos() {
  try {
    console.log("[extraerDatos]");

    //lista de los imputs a recojer
    const elementos = [
      `fechaProduccion`,
      `fechaSecado`,
      `secador`,
      `auxiliares`,
      `referencia`,
      `referenciaExtraida`,
      `turno`,
      `lote`,
      `maquina`,
      `registro`,
      `reproceso`,
      `tipoReproceso`,
      `anteriorRegistro`,
      `temperatura`,
      `tiempoSecado`,
      `tiempoAdicional`,
      `tiempoEnfriamiento`,
      `silicona`,
      `pesoSeco`,
      `unidades`,
      `unidadesTeoricas`,
      `diferencia`,
      `consumoMezclas`,
      `observaciones`,
      `totalTiempo`,
      `totalTiempoMinimo`,
    ]
    //recojer el valor de cada uno
    const rta = elementos.map(element => {
      return obtenerElemento(element)
    })

    console.log(`los datos son `, rta)
    // devolvemos un objeto para tenerlo mas adaptativo
    return rta

  } catch (error) {
    console.error("Error al recolectar la data del formulario: ", error);
  }
}

function obtenerElemento(id) {
  const elemento = document.getElementById(id);
  if (!elemento) {
    console.error(`Elemento no encontrado: ${id}`);
    return null;
  }

  return elemento.value;
}