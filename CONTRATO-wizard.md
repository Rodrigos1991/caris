# Contrato del wizard "Importar ZIP de PDFs" (leído de PROD fecoset-v17, 2026-09-25, solo lectura)

Código: `/opt/odoo/custom/src/as-accounting/payment_day_documents/wizards/payment_day_zip_import.py`
(módulo `payment_day_documents`, Abasto). Botón visible solo con la jornada en estado `pending_close`.

## Qué tiene que cumplir cada archivo del ZIP
| Regla | Detalle |
|---|---|
| Nombre | regex `^(\d+)\s+([\d\-]+)\.pdf$` (sin distinguir mayúsculas). Se usa solo el basename; se ignoran las carpetas |
| CUIT | solo dígitos. Busca `res.partner` con `vat = CUIT OR vat ilike CUIT`, `limit 1`, **sin filtrar por compañía** |
| NRO | solo dígitos y guiones. Tiene que aparecer **como substring** en `ml.name + ml.ref + move.name + move.l10n_latam_document_number` de alguna línea de ESA jornada con ese partner |
| Factura proveedor | NRO = `00001-00000056` → matchea contra `move.name` (`FA-A 00001-00000056`) |
| Adelanto asociado | NRO = `2026-09-00000225` → matchea contra `ml.ref` (la Referencia del asiento) |

## Verificado contra la jornada del 14/09 (id 298, compañía 1 Mundo Textil)
- Los 12 CUITs del Excel → **1 partner cada uno** (compartido, `company_id` NULL). El `limit 1` sin compañía no molesta.
- Las 18 filas del Excel → **18 líneas de la jornada** donde el NRO aparece tal cual. 18 de 18.
- La jornada tiene 3 líneas más (de un mismo proveedor) que no están en el Excel ni en los PDFs.

## Consecuencias para Caris
- De Observaciones sacar el NRO: `FA-A 00001-00000056` → `00001-00000056`; `2026 09 00000219` → `2026-09-00000219`.
- El CUIT sale del Excel, cruzando por NRO.
- El NRO completo evita falsos positivos: como el match es por substring, un NRO truncado podría pegarle a otra línea.
- ⚠️ **No importar dos veces el mismo ZIP.** El adjunto se deduplica por nombre, pero el `document.document` se crea igual en cada importación → quedan documentos duplicados (comportamiento de Abasto, no lo tocamos).
- Lo que no matchea lo reporta Odoo en un toast, **y solo muestra los primeros 10 nombres** → Caris tiene que validar todo antes y dar su propio reporte completo.

Verificado con consultas SQL de solo lectura (no se incluyen en el repo porque tienen datos reales).
