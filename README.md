# SANA STUDIO BY PAUNANI S.A.S. — App de contabilidad bancaria

Aplicación Streamlit para importar uno o varios estados de cuenta Excel del mismo banco, reconocer las columnas observadas en el archivo de muestra y proponer asientos contables de partida doble. Incluye libro diario, mayor, balance de comprobación y exportación conjunta Excel.

## Cómo ejecutarla

1. Instala Python 3.10 o superior.
2. Abre una terminal en esta carpeta.
3. Ejecuta `pip install -r requirements.txt`.
4. Ejecuta `streamlit run app.py`.
5. Abre la dirección local que muestra Streamlit y sube tus Excel.

## Procedimiento

1. Sube uno o varios archivos de movimientos bancarios.
2. Escoge un período (todo, día, mes, año o rango).
3. Revisa cada movimiento, corrige **Contracuenta** y **Categoría** antes de descargar el Excel.
4. Analiza libro diario, mayor y balance de comprobación; descarga el archivo.

## Consideraciones contables importantes

- **No sustituye la contabilidad íntegra**: los movimientos bancarios por sí solos no justifican registros de ventas, inventarios, costo de ventas, facturas, IVA, retenciones, nómina ni asientos de cierre.
- Los aportes de las socias van provisionalmente a **Aportes de socias por formalizar**. Solo corresponde capital social si los documentos lo respaldan.
- Los préstamos se proponen en pasivo y los pagos en disminución de pasivo; **se deben separar intereses, capital y cargos** según contratos y documentos.
- Los reintegros y transferencias requieren verificación, pues no siempre son ingresos ni gastos.
- Las cuentas `9.9.01` y `9.9.02` son **transitorias** y deben reclasificarse antes de una contabilidad definitiva.
- Los saldos de mayor y balance de comprobación corresponden exclusivamente a los **movimientos importados del período** y no integran saldos iniciales. No confundirlos con estados financieros cerrados.
- El algoritmo detecta duplicados por fecha, referencia, descripción, signo, valor y saldo. Puede ser necesario distinguir cuentas bancarias distintas y verificar manualmente movimientos similares.
- Los cambios que realices no persisten tras cerrar la sesión; descarga el Excel antes de salir.
- El plan de cuentas del ejemplo es **referencial**, no equivale al catálogo contable aprobado por la empresa ni valida normativas tributarias automáticamente.

## Formato esperado

Encabezados en alguna de las primeras 50 filas: `FECHA`, `REFERENCIA`, `DESCRIPCION`, `+/-`, `VALOR`, `SALDO CONTABLE` y opcional `tipo`. Fechas como `04/21/2026 02:28:00 PM`. Los archivos pueden incluir metadatos del banco antes de la tabla y hojas auxiliares.
