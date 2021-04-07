
function main(workbook: ExcelScript.Workbook) {
	
  // Get the current worksheet.
  let datosHidrologicos = workbook.getWorksheet("Datos hidrologicos de caudales");
  
  let curvaDuracionCaudales = workbook.getWorksheet("Curva de duracion de caudales");
  
  //CANTIDAD DE FILAS CON DATOS DE MESES 13-663
  let primerDato = 13;
  let ultimoDato = 663;

  
  //Guardamos los meses en un arreglo para recorrelo:
  let meses = ["C","E","G","I","K","M","O","Q","S","U","W","Y"];
  let resultados = ["D","F","H","J","L","N","P","R","T","V","X","Z"];
  //Vamos a tomar los rangos de intervalos de clase: 6-25

  let minimo = 6;
  let maximo = 25;
  
  
  for (let i =0; i < meses.length;i++){
    let mesActual = meses[i];
    let resultadoActual = resultados[i];
  
    for (let j = primerDato; j <= ultimoDato; j++) {
      for (let k = minimo; k <= maximo; k++) {
        let inferior = curvaDuracionCaudales.getRange("C" + k).getValue();
        let superior = curvaDuracionCaudales.getRange("D" + k).getValue();
        let datoAEvaluar = datosHidrologicos.getRange(mesActual + j).getValue();
        if(datoAEvaluar>=inferior && datoAEvaluar<=superior){        
          datosHidrologicos.getRange(resultadoActual + j).setValue(curvaDuracionCaudales.getRange("B" + k).getValue());
          break;
        }      
      }
    }
  }
  console.log("TERMINO EL TRABAJO CON EXITO");
}
