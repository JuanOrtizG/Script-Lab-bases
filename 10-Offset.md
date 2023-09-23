```javascript

    let range = context.workbook.getSelectedRange(); // selecciono un rango con mouse o teclado
    range.load("values");  // cargo los valores 
    await context.sync();
    
    range.getCell(fila, columna).values = [["SI"]]; // El primer datos seria mi punto 0,0 y puedo desplazarme
```        
