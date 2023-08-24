
Este es el esquema básico para iniciar en script lab
```javascript
//Un codigo para verificar errores
$("#run").click(() => tryCatch(run));

//inicio de la funcion que contendra nuestro codigo
async function run() {
  await Excel.run(async (context) => {
      
  });//run()
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
