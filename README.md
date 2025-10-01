# Actualizador de Sensores - PowerShell

## Descripción del Proyecto

Este script en **PowerShell** automatiza la gestión y actualización de inventarios de sensores médicos (como Abbott y Dexcom).  

Antes, el proceso era tedioso: se tenían que registrar **uno a uno los folios** con sus correspondientes CIPA y unidades en Excel. Este script permite:

- Leer los archivos generados tras el escaneo de folios.
- Detectar automáticamente el mes de envío.
- Validar y corregir CIPA no estándar (con entrada manual si es necesario).
- Actualizar la hoja de cálculo de Excel con los datos de manera automática.
- Mover los archivos TXT y PDF procesados a una carpeta de **histórico**.
- Generar un resumen final de unidades por CIPA y total de sensores.

La obtención de los datos se hace mediante **escaneo de folios físicos**, que luego se convierten a **PDF con OCR**, permitiendo extraer la información automáticamente.

> ⚠️ Este proyecto fue desarrollado en el ordenador del Centro de la Comunidad de Madrid, que está **capado** y no permite instalar lenguajes de programación adicionales. Por eso se usó **PowerShell**, que viene incluido en Windows, trabajando **offline** y sin necesidad de descargas externas.

---

## Configuración

1. **Rutas de carpetas y archivos**: Edita las variables del script en el archivo 'Config-Paths.example.ps1' para establecer las rutas correctas según tu entorno.

2. **Estructura de archivos:**: 

- Los archivos TXT deben generarse a partir del escaneo/OCR de los folios de entrega.

- Los PDF originales se moverán a una carpeta Histórico automáticamente.

- Cada proveedor tiene su propia carpeta y Excel maestro.

---

## Ejecución

1. Abre PowerShell en modo Administrador si fuera necesario.

2. Navega a la carpeta donde tengas el script:
```powershell
cd "RUTA\DEL\SCRIPT"
```

3. Ejecuta el script:
```powershell
.\Update-SensorInventory.ps1
```

4. El script mostrará ventanas de confirmación antes de actualizar cada archivo y antes de volcar datos a Excel.
5. Una vez finalizado, mostrará un resumen final de unidades procesadas.

---

## Notas importantes

- Si el script encuentra CIPA no válidos, pedirá la corrección manual.

- Los datos se insertarán al final de la tabla, dejando algunas filas de separación para no interferir con los datos existentes.

- PowerShell se eligió porque el entorno no permite instalar otros lenguajes y se trabaja offline, asegurando compatibilidad con Windows y seguridad de los datos.

---

## Ventajas

- Evita la entrada manual de folios uno a uno.

- Minimiza errores humanos.

- Genera un histórico automático de archivos procesados.

- Funciona sin instalación de software adicional, solo con Windows y PowerShell.

