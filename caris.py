"""Caris — parte comprobantes de pago Credicoop y arma el ZIP para
Odoo > Jornada de Pago > "Importar ZIP de PDFs".

Uso: python caris.py <excel.xlsx> <pdf> [<pdf> ...] [-o salida.zip]

Contrato del wizard (ver CONTRATO-wizard.md):
  - nombre  ^\\d+\\s+[\\d-]+\\.pdf$   ->  "CUIT NRO.pdf"
  - NRO se busca por SUBSTRING -> solo se acepta el NRO completo, validado contra el Excel
  - Odoo muestra solo 10 no-procesados -> el reporte completo lo da Caris
  - reimportar el mismo ZIP duplica documentos en Odoo -> se avisa en el reporte
Controles: importe del PDF = "Importe" del pago (el neto transferido, ya descontadas las
retenciones) y cuenta destino del PDF = "Cuenta bancaria receptora" del pago (CBU o CVU).
"""
import argparse
import io
import re
import subprocess
import sys
import zipfile
from pathlib import Path

import openpyxl
import pymupdf

TESSERACT = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
WIZARD_RE = re.compile(r"^\d+\s+[\d\-]+\.pdf$", re.IGNORECASE)
NRO_OK = re.compile(r"^(\d{5}-\d{8}|\d{4}-\d{2}-\d{8})$")
# Observaciones: factura "FA-A 00001-00000056" / adelanto "2026 09 00000219".
# En transferencias a otros bancos Credicoop saca los guiones: "FA C 00001 00000062".
FACTURA_RE = re.compile(r"(?<!\d)(\d{5})[\s-]+(\d{8})(?!\d)")
ADELANTO_RE = re.compile(r"(?<!\d)(\d{4})[\s-]+(\d{2})[\s-]+(\d{8})(?!\d)")
IMPORTE_RE = re.compile(r"\$\s*([\d.]+,\d{2})")
# Renglón "...a acreditar:": CBU/CVU "01234567-89012345678901" (22 dígitos) o, entre cuentas
# Credicoop, la cuenta interna "CC$ 191-123-456789/0" (banco-sucursal-cuenta/dígito).
CUENTA_INTERNA_RE = re.compile(r"(\d{3})-(\d{3})-(\d{6})/(\d)")


def leer_excel(path):
    ws = openpyxl.load_workbook(path, data_only=True).active
    filas = list(ws.iter_rows(values_only=True))
    head = [str(h or "") for h in filas[0]]

    def col(*claves):
        for i, h in enumerate(head):
            if all(c.lower() in h.lower() for c in claves):
                return i
        sys.exit(f"ERROR: no encuentro la columna {claves} en el Excel. Columnas: {head}")

    c_nom, c_cuit, c_imp = col("Cliente/proveedor"), col("NIF"), col("Importe")
    c_cbu = col("Cuenta bancaria receptora")
    c_doc, c_ref = col("Número de Documento"), col("Referencia")
    pagos, errores = {}, []
    for n, f in enumerate(filas[1:], start=2):
        if not any(f):
            continue
        cuit = re.sub(r"\D", "", str(f[c_cuit] or ""))
        # Número de Documento si tiene formato válido (facturas, notas de débito);
        # si no (adelanto "AS/2026/1218") -> Referencia. La Referencia puede traer texto libre.
        doc, ref = str(f[c_doc] or "").strip(), str(f[c_ref] or "").strip()
        nro = doc if NRO_OK.match(doc) else ref
        if len(cuit) != 11 or not NRO_OK.match(nro):
            errores.append(f"Excel fila {n}: CUIT '{cuit}' / NRO '{nro}' con formato inválido")
            continue
        if nro in pagos:
            errores.append(f"Excel fila {n}: NRO {nro} repetido")
            continue
        # La celda puede traer "0140...1204" o "0140...1204 - Banco X" (si se exportó la cuenta sin el subcampo)
        m = re.search(r"(?<!\d)\d{22}(?!\d)", str(f[c_cbu] or ""))
        cbu = m.group() if m else ""
        pagos[nro] = {"cuit": cuit, "nombre": f[c_nom], "importe": round(float(f[c_imp] or 0), 2),
                      "cbu": cbu, "fila": n}
    return pagos, errores


def ocr(page):
    png = page.get_pixmap(dpi=300, colorspace=pymupdf.csGRAY).tobytes("png")
    r = subprocess.run([TESSERACT, "stdin", "stdout", "--psm", "6", "-l", "eng"],
                       input=png, capture_output=True, check=True)
    return r.stdout.decode("utf-8", "replace")


def leer_cuenta(texto):
    """Cuenta destino del renglón "a acreditar": ("cbu", 22 dígitos), ("interna", grupos) o None."""
    linea = next((l for l in texto.splitlines() if "acreditar" in l.lower()), "")
    dato = linea.split(":", 1)[-1]
    m = CUENTA_INTERNA_RE.search(dato)
    if m:
        return "interna", m.groups()
    digitos = re.sub(r"\D", "", dato)
    return ("cbu", digitos) if len(digitos) == 22 else None


def chequear_cuenta(texto, cbu_odoo):
    """Devuelve (error, aviso). Error: la cuenta del PDF no es la de Odoo -> no va al ZIP.
    Aviso: no se pudo verificar -> va al ZIP igual, pero se informa."""
    if not cbu_odoo:
        return None, "el pago no tiene cuenta bancaria en Odoo: cuenta no verificada"
    cuenta = leer_cuenta(texto)
    if cuenta is None:
        return None, "no pude leer la cuenta destino del PDF: cuenta no verificada"
    tipo, dato = cuenta
    if tipo == "cbu":
        leido, ok = dato, dato == cbu_odoo
    else:  # CBU Credicoop = 191 + 0 + sucursal + ... + cuenta + dígito + ...
        banco, suc, nro, dv = dato
        leido = f"{banco}-{suc}-{nro}/{dv}"
        ok = cbu_odoo.startswith(banco) and cbu_odoo[3:7] == "0" + suc and nro + dv in cbu_odoo[8:]
    return (None if ok else f"cuenta PDF {leido} ≠ Odoo {cbu_odoo}"), None


def nombre_pdf(p, nro):
    return f"{p['cuit']} {nro}.pdf"


def extraer(texto):
    """Devuelve (nros encontrados en Observaciones, importe leído)."""
    obs = next((l for l in texto.splitlines() if "observ" in l.lower()), "")
    nros = [f"{a}-{b}-{c}" for a, b, c in ADELANTO_RE.findall(obs)]
    if not nros:
        nros = [f"{a}-{b}" for a, b in FACTURA_RE.findall(obs)]
    importes = IMPORTE_RE.findall(texto)
    importe = float(importes[-1].replace(".", "").replace(",", ".")) if importes else None
    return nros, importe, obs.strip()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("excel")
    ap.add_argument("pdfs", nargs="+")
    ap.add_argument("-o", "--salida", default="comprobantes.zip")
    a = ap.parse_args()

    pagos, errores = leer_excel(a.excel)
    usados = {}  # nro -> origen
    avisos = []  # van al ZIP, pero con la cuenta sin verificar
    salida = []  # (nombre_archivo, bytes)

    for pdf in a.pdfs:
        doc = pymupdf.open(pdf)
        for i, page in enumerate(doc):
            origen = f"{Path(pdf).name} pág {i + 1}"
            texto = ocr(page)
            nros, importe, obs = extraer(texto)
            if len(nros) != 1:
                errores.append(f"{origen}: no pude leer UN número en Observaciones (OCR: '{obs}')")
                continue
            nro = nros[0]
            if nro not in pagos:
                errores.append(f"{origen}: NRO {nro} no está en el Excel (OCR: '{obs}')")
                continue
            if nro in usados:
                errores.append(f"{origen}: NRO {nro} ya salió de {usados[nro]} (comprobante duplicado)")
                continue
            p = pagos[nro]
            if importe is None or abs(importe - p["importe"]) > 0.01:
                errores.append(f"{origen}: NRO {nro} importe PDF {importe} ≠ Excel {p['importe']} — revisar")
                continue
            error, aviso = chequear_cuenta(texto, p["cbu"])
            if error:
                errores.append(f"{origen}: NRO {nro} {error} — revisar")
                continue
            nombre = nombre_pdf(p, nro)
            if aviso:
                avisos.append(f"{nombre}: {aviso}")
            assert WIZARD_RE.match(nombre), nombre
            uno = pymupdf.open()
            uno.insert_pdf(doc, from_page=i, to_page=i)
            salida.append((nombre, uno.tobytes()))
            usados[nro] = origen

    for nro, p in pagos.items():
        if nro not in usados:
            errores.append(f"Excel fila {p['fila']}: {p['nombre']} {nro} ${p['importe']:,.2f} sin comprobante en los PDFs")

    if salida:
        with zipfile.ZipFile(a.salida, "w", zipfile.ZIP_DEFLATED) as z:
            for nombre, data in salida:
                z.writestr(nombre, data)

    print(f"\nCaris — {len(salida)} de {len(pagos)} comprobantes del Excel listos → {a.salida if salida else '(sin ZIP)'}")
    for nombre, _ in sorted(salida):
        print(f"  OK  {nombre}")
    if avisos:
        print(f"\n⚠ {len(avisos)} aviso(s) — van en el ZIP, pero revisalos:")
        for av in avisos:
            print(f"  ⚠ {av}")
    if errores:
        print(f"\n⚠ {len(errores)} problema(s) — estos NO van en el ZIP:")
        for e in errores:
            print(f"  ✗ {e}")
    print("\nRecordá: importá este ZIP UNA sola vez (reimportarlo duplica documentos en Odoo).")
    sys.exit(1 if errores else 0)


if __name__ == "__main__":
    main()
