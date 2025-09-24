# PLMECO Transporte (Desktop App)

Aplicación WPF para gestionar reuniones de transporte, reemplazo del Excel con macros.

## Funcionalidades
- Importar fichero Excel/CSV de la reunión.
- Mostrar tabla con columnas clave (Matrícula, Destino, Muelle, Estado, Salida Tope).
- Doble clic para sellar horas de Llegada y Salida.
- Incidencias automáticas: "RETRASO TRANSPORTISTA" si Salida > Salida Tope.
- Checkbox LEX por fila.
- Orden automático por Muelle.

## Tecnologías
- C# .NET 8
- WPF
- ClosedXML (importar Excel)
- SQLite (persistencia local opcional)

## Uso
1. `dotnet build`
2. `dotnet run --project src/Plmeco.App`
3. Botón **Importar Excel** para cargar datos.

## Instalador
En la carpeta `installers/` hay un script Inno Setup para generar un `.exe` instalable.

## Licencia
MIT
