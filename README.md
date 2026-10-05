# Caris

Arma el ZIP de comprobantes de pago para **Odoo › Jornada de Pago › "Importar ZIP de PDFs"**.

Recibe los PDFs de comprobantes que baja el home banking de Credicoop (un comprobante por página) y el Excel de la jornada exportado desde Odoo. Devuelve un ZIP con un PDF por comprobante, con el nombre que espera el wizard: `CUIT NRO.pdf` (ej. `20123456789 00001-00000056.pdf`).

Corre **fuera de Odoo**: no se conecta a ninguna base ni escribe nada en Odoo. Lo usa Tesorería.

## Cómo se usa

1. En Odoo, abrir la jornada › **Pagos** › seleccionar todos › Acciones › Exportar, con la plantilla **"Exportador Caris"** y formato XLSX. Con otra plantilla faltan columnas y Caris no puede cruzar los datos.
2. Abrir `caris.html` en Chrome.
3. Cargar el Excel y los PDFs y apretar **Procesar**. El OCR tarda unos segundos por página. **No cambiar de pestaña mientras procesa**: Chrome frena el OCR en segundo plano.
4. Revisar el reporte y descargar el ZIP.
5. Importar el ZIP en la jornada **una sola vez**. Si se reimporta, Odoo duplica los documentos.

Lo marcado con ❌ no entra en el ZIP y se sube a mano en Odoo.

## Qué procesa y qué va a mano

| Tipo de pago | Caris |
|---|---|
| Facturas de proveedor | ✅ |
| Adelantos de asociados | ✅ |
| Notas de débito | ✅ |
| Pagos con retenciones | ✅ |
| Pagos de servicios por Epagos (sin Observaciones) | ✋ a mano |

**Cómo cruza:** lee con OCR el campo **Observaciones** del comprobante (ej. `FA-A 00001-00000056`), busca ese número en el Excel y toma el CUIT de ahí. Después hace dos controles:

- **Importe:** el del PDF tiene que coincidir al centavo con la columna **Importe** del pago, que es el neto transferido (ya descontadas las retenciones).
- **Cuenta destino:** el CBU o CVU "a acreditar" del PDF tiene que coincidir con la **Cuenta bancaria receptora** del pago. En transferencias entre cuentas Credicoop el comprobante muestra la cuenta interna (`CC$ 191-123-456789/0`), que se verifica contra el CBU.

Si algo no coincide, no adivina: lo marca con ❌. Si el pago no tiene cuenta en Odoo o no se puede leer la del PDF, el comprobante entra en el ZIP con ⚠️ "cuenta no verificada".

## Publicarlo (para quien administra la página de aplicaciones)

`caris.html` es **un único archivo estático**: no tiene backend, no hace falta instalar nada en el servidor y se puede servir desde cualquier servidor web.

- **Necesita internet.** Carga las librerías desde CDN (pdf.js, tesseract.js, pdf-lib, SheetJS, JSZip) y tesseract.js baja el modelo de idioma la primera vez.
- **Los datos no salen de la computadora del usuario.** Todo el procesamiento ocurre en el navegador y los PDFs y el Excel no se envían a ningún lado.
- Se probó servido por `http://`. Abrirlo con doble clic (`file://`) no está probado.
- Conviene limitar el acceso a los usuarios de Tesorería.

## Versión de línea de comandos (opcional)

`caris.py` hace lo mismo desde la terminal y sirve como referencia. Necesita Python 3, `pip install openpyxl pymupdf` y [Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki) instalado en `C:\Program Files\Tesseract-OCR\` (si está en otro lugar, cambiar `TESSERACT` en el script).

```
python caris.py jornada.xlsx comprobantes1.pdf comprobantes2.pdf -o salida.zip
```

## Referencia

`CONTRATO-wizard.md` explica qué valida el wizard de Odoo al importar el ZIP: formato del nombre, cómo busca el CUIT y el número, y sus limitaciones.
