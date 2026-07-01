# -*- coding: utf-8 -*-
"""
COA -> celda del Google Sheet (columna "coa").

Lee un PDF de Certificado de Analisis (formato Kovera Labs / Glacier Aminos),
extrae lote, pureza, mg rotulados, mg reales, fecha, laboratorio, N de reporte y
cantidad de viales testeados, y copia al portapapeles la cadena lista para pegar:

    lote=... | pureza=... | rotulado=... | real=... | fecha=... | lab=... | reporte=... | tests=... | pdf=

Uso:
  - Arrastra el PDF sobre el archivo "COA a Sheet (arrastra el PDF aca).bat", o
  - python coa_a_sheet.py "ruta\\al\\certificado.pdf", o
  - python coa_a_sheet.py   (abre un selector de archivo)

Solo falta pegar el link de Google Drive despues de "pdf=".
"""

import sys
import os
import re
import subprocess


def asegurar_pdfplumber():
    try:
        import pdfplumber  # noqa: F401
        return
    except ImportError:
        print("Instalando pdfplumber (solo la primera vez)...")
        subprocess.run([sys.executable, "-m", "pip", "install", "--quiet", "pdfplumber"])
        import pdfplumber  # noqa: F401


def elegir_pdf():
    if len(sys.argv) > 1 and sys.argv[1].lower().endswith(".pdf"):
        return sys.argv[1]
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        ruta = filedialog.askopenfilename(
            title="Elegi el PDF del COA",
            filetypes=[("PDF", "*.pdf")],
        )
        root.destroy()
        return ruta
    except Exception:
        return ""


def texto_y_filas(ruta):
    import logging
    import warnings
    logging.getLogger("pdfminer").setLevel(logging.ERROR)
    warnings.filterwarnings("ignore")
    import pdfplumber
    texto_partes = []
    filas = []
    with pdfplumber.open(ruta) as pdf:
        for pagina in pdf.pages:
            t = pagina.extract_text() or ""
            texto_partes.append(t)
            for tabla in pagina.extract_tables() or []:
                for fila in tabla:
                    filas.append([(c or "").strip() for c in fila])
    return "\n".join(texto_partes), filas


def celda_despues(filas, etiqueta):
    """Devuelve la celda inmediatamente a la derecha de una etiqueta exacta."""
    et = etiqueta.lower()
    for fila in filas:
        celdas = [c for c in fila if c != ""]
        for i, c in enumerate(celdas):
            if c.lower() == et and i + 1 < len(celdas):
                return celdas[i + 1]
    return None


def fila_promedio(filas):
    """Fila 'Batch Average': devuelve (pureza, contenido)."""
    for fila in filas:
        celdas = [c for c in fila if c != ""]
        for i, c in enumerate(celdas):
            if c.lower().startswith("batch average") and i + 2 < len(celdas):
                return celdas[i + 1], celdas[i + 2]
    return None, None


def coma_decimal(valor):
    """98.915% -> 98,915% ; 54.10 mg -> 54,10 mg (solo el punto entre digitos)."""
    if not valor:
        return valor
    return re.sub(r"(\d)\.(\d)", r"\1,\2", valor)


def fecha_ddmmaaaa(texto):
    # Peptidos: "Certified: MM/DD/AAAA". Agua/insumos: "Analysis Date MM/DD/AAAA".
    m = re.search(r"(?:Certified|Analysis Date|Date of Analysis|Test Date)\s*:?\s*(\d{1,2})/(\d{1,2})/(\d{4})", texto, re.I)
    if not m:
        m = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})", texto)
    if not m:
        return None
    a, b, anio = m.group(1), m.group(2), m.group(3)
    # El COA viene en formato de EE. UU. (MM/DD/AAAA). Si el primer numero es > 12,
    # ya viene como DD/MM y no se toca.
    if int(a) > 12:
        dia, mes = a, b
    else:
        mes, dia = a, b
    return "%02d/%02d/%s" % (int(dia), int(mes), anio)


def normalizar_mg(valor):
    if not valor:
        return valor
    return re.sub(r"(\d)\s*mg", r"\1 mg", valor).strip()


def extraer(texto, filas):
    d = {}

    # --- lote (Batch / Lot) ---
    d["lote"] = celda_despues(filas, "Batch") or celda_despues(filas, "Lot / Batch")
    if not d["lote"]:
        # El codigo de lote siempre tiene al menos un digito (asi no confunde con
        # "Average", "Avg", etc.). Cubre Kovera ("Batch X", "Lot / Batch X") e ILS
        # ("Lot Number: X").
        m = re.search(r"(?:Lot\s*Number|Lot\s*/\s*Batch|Batch|Lot)\s*:?\s*([A-Za-z0-9\-]*\d[A-Za-z0-9\-]*)", texto, re.I)
        d["lote"] = m.group(1) if m else None

    # --- pureza + contenido real (mg) ---
    pureza, real = fila_promedio(filas)
    if not pureza:
        m = re.search(r"Batch Average\s+([\d.,]+\s*%)", texto, re.I)
        pureza = m.group(1) if m else None
    if not pureza:
        # Agua/insumos: "Purity (GC) >=99.5% 99.9%" -> el resultado es el ultimo %.
        m = re.search(r"Purity[^\n%]*?[\d.,]+\s*%\s+([\d.,]+\s*%)", texto, re.I)
        if not m:
            m = re.search(r"Purity[^\n%]*?([\d.,]+\s*%)", texto, re.I)
        pureza = m.group(1) if m else None
    if not real:
        m = re.search(r"Batch Average\s+[\d.,]+\s*%\s+([\d.,]+\s*mg)", texto, re.I)
        real = m.group(1) if m else None
    if not real:
        # ILS: fila "Mean 99.73% 10.54 mg".
        m = re.search(r"\bMean\s+[\d.,]+\s*%\s+([\d.,]+\s*mg)", texto, re.I)
        real = m.group(1) if m else None

    # --- rotulado (mg de etiqueta; no aplica a insumos como el agua) ---
    d["rotulado"] = celda_despues(filas, "Labeled Qty")
    if not d["rotulado"]:
        m = re.search(r"Labeled\s*Qty\s+([\d.,]+\s*mg)", texto, re.I)
        d["rotulado"] = m.group(1) if m else None
    if not d["rotulado"]:
        # ILS: "Concentration: 10mg".
        m = re.search(r"Concentration\s*:?\s*([\d.,]+\s*mg)", texto, re.I)
        d["rotulado"] = m.group(1) if m else None
    if not d["rotulado"]:
        m = re.search(r"\((\d+(?:[.,]\d+)?)\s*mg\s*[±+]", texto)
        d["rotulado"] = (m.group(1) + " mg") if m else None

    d["pureza"] = coma_decimal((pureza or "").replace(" ", "")) or None
    d["real"] = coma_decimal(normalizar_mg(real)) if real else None
    d["rotulado"] = normalizar_mg(d["rotulado"]) if d["rotulado"] else None

    # fecha
    d["fecha"] = fecha_ddmmaaaa(texto)

    # laboratorio: "tested by X using" o, si no, cualquier "* Labs".
    m = re.search(r"tested by\s+(.+?)\s+using", texto, re.I)
    if not m:
        m = re.search(r"\b([A-Z][A-Za-z]+\s+Lab(?:oratories|s)?)\b", texto)
    d["lab"] = m.group(1).strip() if m else None

    # numero de reporte: el codigo Kovera KVR-AAAA-XXXXX (bajo "Report #" o "Sample ID").
    m = re.search(r"\b([A-Z]{2,}-\d{4}-[A-Za-z0-9]+)\b", texto)
    if not m:
        m = re.search(r"(?:Report\s*#|Sample ID)\s*:?\s*#?\s*([A-Za-z0-9][A-Za-z0-9-]{3,})", texto, re.I)
    d["reporte"] = m.group(1) if m else None

    # cantidad de viales testeados ("Vial 1..N" o "(N VIALS)")
    viales = set(re.findall(r"\bVial\s+(\d+)\b", texto))
    if viales:
        d["tests"] = str(len(viales))
    else:
        # Kovera: "(3 VIALS)". ILS: "(3 samples tested)".
        m = re.search(r"\((\d+)\s*(?:VIALS?|samples?\s*tested)\)", texto, re.I)
        d["tests"] = m.group(1) if m else None

    # nombre del producto (solo para mostrar, no va en la cadena)
    d["_producto"] = celda_despues(filas, "Product")

    return d


def construir_cadena(d):
    orden = ["lote", "pureza", "rotulado", "real", "fecha", "lab", "reporte", "tests"]
    partes = ["%s=%s" % (k, d[k]) for k in orden if d.get(k)]
    return " | ".join(partes) + " | pdf="


def copiar(texto):
    try:
        subprocess.run(["clip"], input=texto, text=True)
        return True
    except Exception:
        return False


def main():
    asegurar_pdfplumber()
    ruta = elegir_pdf()
    if not ruta or not os.path.isfile(ruta):
        print("No se eligio ningun PDF.")
        input("\nEnter para cerrar...")
        return

    print("Leyendo: %s\n" % os.path.basename(ruta))
    texto, filas = texto_y_filas(ruta)
    d = extraer(texto, filas)

    if d.get("_producto"):
        print("Producto detectado: %s" % d["_producto"])
    print("\nCampos extraidos:")
    for k in ["lote", "pureza", "rotulado", "real", "fecha", "lab", "reporte", "tests"]:
        estado = d.get(k) if d.get(k) else "(no encontrado)"
        print("  %-9s %s" % (k + ":", estado))

    cadena = construir_cadena(d)
    print("\n" + "=" * 60)
    print("CADENA PARA LA COLUMNA \"coa\" (ya copiada al portapapeles):\n")
    print(cadena)
    print("=" * 60)

    if copiar(cadena):
        print("\n[OK] Copiado. Pegala en la fila del producto y agrega el link de Drive despues de 'pdf='.")
    else:
        print("\n[!] No se pudo copiar solo. Copiala a mano de arriba.")

    input("\nEnter para cerrar...")


if __name__ == "__main__":
    main()
