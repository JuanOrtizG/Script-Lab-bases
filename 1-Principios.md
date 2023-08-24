
Este es el esquema básico para iniciar en script lab
```javascript
//Un codigo para verificar errores
$("#run").click(() => tryCatch(run));

//inicio de la funcion que contendra nuestro codigo
async function run() {
  await Excel.run(async (context) => {

      //aqui va nuestro codigo...

  });
}


/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}
```
Luego en *//aqui va nuestro codigo*,  colocaremos:

```javascript
  const sheet = context.workbook.worksheets.getActiveWorksheet();  //Activa la hoja que queremos utilizar en excel
  const range = sheet.getRange("B1:AS1069"); // Aqui elegimos el rango de tabla que queremos capturar en nuestro codigo
  range.load("values"); // volcamos los datos de nuestra tabla
  const tablaGuardada = range.values; //Guardamos los datos en una lista de sublistas

  console.log(  tablaGuardada[3][4]  ) // tablaGuardada es una lista con sublistas, una estructura de datos normal.
```
 
