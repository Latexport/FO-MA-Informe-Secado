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
  console.log("[submit()]");

  // Obtiene los valores ingresados en el formulario
  const datosExtraidos = extraerDatos();
  console.log("LOS DATOS SON : ", datosExtraidos);
  // Establece los valores en las celdas correspondientes

  //una funcion manejara todo los dados , creanod nuevas varibles segun lo necesecitados
  const dataExport = convertirDatos(datosExtraidos)

  //devuelve un array en el orden de las columnas
  const data = organisarCeldas(dataExport)

  const dataExcel = [
    ...data
  ];

  // Nombre de la hoja en la que deseas guardar los datos

  // Guarda los datos en Excel en la hoja especificada
  const nombreHoja = "seguimiento";
  await agregarDatosExcel(nombreHoja, dataExcel);
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

function convertirDatos(data) {
  try {
    //agregamos los nuevos parametros y varibles que al final vamos a mandar a la hoja
    const nuevaVariable = ' es un ejemplo '
    //tambien ejecutamos aqui las funiciones de converdion de los datos 
    // ejem: la suma de las horas en una nueva variable

    return { ...data, nuevaVariable }
  } catch (error) {
    throw error
  }
}
function organisarCeldas(data) {
  try {
    //organizamos las variables en el orden que queremos
    const rta = [
      data.nuevaVariable,
      //el resto de la varibles en orden
    ]
    // tambien podemos hacer algo asi pero el orden no quedaria especifiacdo
    // const rta = Object.values(data);

    return rta
  } catch (error) {
    throw error
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