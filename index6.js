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
      console.log("EXTRAEMOS LOS DATOS");
  
      const fechaProduccion = (document.getElementById("fechaProduccion") as HTMLInputElement).value;
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
  
      toggleCampos();
      sumarTiempos();
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
  
    //*****suma tiempo 7 completada face 2*****
  
  }
  
  function toggleCampos() {
    var reprocesoSelect = (document.getElementById("reproceso") as HTMLSelectElement).value;
    var tipoReprocesoInput = document.getElementById("tipoReproceso") as HTMLInputElement;
    var anteriorRegistroInput = document.getElementById("anteriorRegistro") as HTMLInputElement;
  
    // Si la opción seleccionada es "No", deshabilita los campos
    if (reprocesoSelect === "No") {
  
      tipoReprocesoInput.disabled = true;
      anteriorRegistroInput.disabled = true;
    } else {
      // Si la opción seleccionada es "Sí", habilita los campos
  
      tipoReprocesoInput.disabled = false;
      anteriorRegistroInput.disabled = false;
    }
  }
  
  
  function sumarTiempos() {
    // Obtener los valores de los campos de tiempo
    const tiempoSecadoInput = document.getElementById("tiempoSecado") as HTMLInputElement;
    const tiempoAdicionalInput = document.getElementById("tiempoAdicional") as HTMLInputElement;
    const tiempoEnfriamientoInput = document.getElementById("tiempoEnfriamiento") as HTMLInputElement;
  
    // Intentar convertir los valores a números
    const tiempoSecado = parseFloat(tiempoSecadoInput.value) || 0;
    const tiempoAdicional = parseFloat(tiempoAdicionalInput.value) || 0;
    const tiempoEnfriamiento = parseFloat(tiempoEnfriamientoInput.value) || 0;
  
    // Verificar si la conversión fue exitosa
    if (isNaN(tiempoSecado) || isNaN(tiempoAdicional) || isNaN(tiempoEnfriamiento)) {
      console.error("Error: Uno o más valores de tiempo no son números válidos.");
      return;
    }
  
    // Sumar los tiempos
    const totalTiempos = tiempoSecado + tiempoAdicional + tiempoEnfriamiento;
    console.log("totalTiempoww", totalTiempos);
  
    // Mostrar el resultado en el campo "totalTiempo"
    console.log("totalTiempos", totalTiempos);
    (document.getElementById("totalTiempo") as HTMLInputElement).value = totalTiempos.toString();
  
    return totalTiempos;
  }
  
  
  
  
  