```javascript

    let range = context.workbook.getSelectedRange(); // selecciono un rango con mouse o teclado
    range.load("values");  // cargo los valores 
    await context.sync();
    
    range.getCell(index, 1).values = [["SI"]]; // El primer datos seria mi punto 0,0 y puedo desplazarme
```        
