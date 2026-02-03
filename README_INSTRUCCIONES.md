# ERP Lucas v5.3 - Instrucciones de Actualización

## 🔧 Correcciones Incluidas

### Problema 1: Patrimonio se borra al recargar
**Causa**: El sistema leía de "PN {año}" pero escribía en "Patrimonio" (hojas diferentes).

**Solución**: Nueva hoja dedicada `Patrimonio_App` que sirve tanto para lectura como escritura, garantizando consistencia entre dispositivos.

### Problema 2: Cuotas no se reflejan en saldos futuros
**Causa**: La función `getCuotasPendientes()` filtraba por `idGrupo` que siempre era `null`.

**Solución**: 
1. El trigger ahora genera `idGrupo` para agrupar cuotas de la misma compra
2. `getCuotasPendientes()` ahora funciona con o sin `idGrupo` (agrupa por descripción+fecha si falta)
3. Nueva sección "Proyección Próximos Meses" en la vista de Tarjeta

---

## 📋 Pasos de Instalación

### Paso 1: Actualizar Google Apps Script

1. Abrí tu Google Spreadsheet
2. Ir a **Extensiones > Apps Script**
3. **IMPORTANTE**: Eliminá TODO el código existente
4. Copiá y pegá el contenido de `GOOGLE_APPS_SCRIPT_v5.3.js`
5. Guardá (Ctrl+S)

### Paso 2: Configurar hojas necesarias

En Apps Script, ejecutá estas funciones **UNA SOLA VEZ** (en orden):

1. Seleccioná `setupPatrimonioSheet` en el menú desplegable y clickeá "Ejecutar"
   - Esto crea la hoja `Patrimonio_App` para persistir los saldos
   
2. Seleccioná `setupTrigger` y clickeá "Ejecutar"
   - Esto instala el trigger para expandir cuotas automáticamente
   - También agrega la columna `ID Grupo` a la hoja de gastos

**Nota**: Puede pedirte autorización la primera vez. Aceptá todos los permisos.

### Paso 3: Configurar saldos iniciales de patrimonio

1. Abrí la hoja `Patrimonio_App` que se creó
2. Ingresá los saldos actuales en USD de cada cuenta:
   - BBVA
   - Caja Seguridad
   - Efectivo

### Paso 4: Re-publicar la API

1. En Apps Script, ir a **Implementar > Nueva implementación**
2. Seleccionar "Aplicación web"
3. Configurar:
   - Descripción: "ERP Lucas v5.3"
   - Ejecutar como: "Yo"
   - Acceso: "Cualquier persona"
4. Clickear "Implementar"
5. **IMPORTANTE**: Copiá la nueva URL de implementación

### Paso 5: Actualizar index.html

1. Abrí `index.html` en un editor de texto
2. Buscá `SCRIPT_URL: ''` (aproximadamente línea 360-370)
3. Pegá tu nueva URL de implementación entre las comillas
4. Guardá el archivo
5. Subí el nuevo `index.html` a tu hosting (GitHub Pages, etc.)

---

## 🔍 Verificación

Para verificar que todo funciona:

1. Abrí la app en tu navegador
2. Clickeá el badge "Sincronizado" para forzar una sincronización
3. Verificá:
   - ✅ La solapa Patrimonio muestra los saldos correctos
   - ✅ Al modificar patrimonio desde otro dispositivo, se sincroniza
   - ✅ Las cuotas pendientes aparecen en la vista de Tarjeta
   - ✅ La proyección de próximos meses muestra los saldos futuros

### Test de diagnóstico

Podés abrir la consola del navegador (F12) y ejecutar:
```javascript
runDiagnostic()
```

Esto probará la conexión y mostrará si todo está configurado correctamente.

---

## 📁 Estructura de Hojas

Tu Spreadsheet debería tener estas hojas:

| Hoja | Propósito |
|------|-----------|
| `Formulario Gastos` | Donde Google Forms guarda los gastos |
| `Patrimonio_App` | Saldos USD de cuentas (NUEVO) |
| `Inversiones` | Lista de inversiones |
| `Movimientos` | Historial de movimientos USD |
| `Ingresos Mensuales` | Ingresos por mes |
| `Config` | Configuración general |

---

## ⚠️ Notas Importantes

1. **No modifiques las primeras 8 columnas** de "Formulario Gastos" - son las que usa el Google Form
2. Las columnas I, J, K, L (Cuota, Total Cuotas, Mes Pago, ID Grupo) las maneja el trigger automáticamente
3. Si agregás un gasto con cuotas vía el Form, esperá unos segundos antes de sincronizar
4. Los gastos viejos sin `idGrupo` seguirán funcionando gracias al algoritmo de agrupación por descripción

---

## 🆘 Solución de Problemas

### "Patrimonio siempre en 0"
- Verificá que existe la hoja `Patrimonio_App`
- Ejecutá `setupPatrimonioSheet` de nuevo
- Ingresá los saldos manualmente en la hoja

### "Las cuotas no se expanden"
- Verificá que el trigger está instalado: Ejecutá `setupTrigger`
- Revisá que el nombre de la hoja sea exactamente `Formulario Gastos`

### "Error de conexión"
- Verificá que la URL en `SCRIPT_URL` sea la correcta
- Asegurate de haber dado acceso "Cualquier persona" en la implementación

### "No se ve la proyección de meses"
- Las cuotas deben tener el campo `mesPago` correctamente calculado
- Sincronizá de nuevo después de agregar gastos con cuotas
