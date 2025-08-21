# -*- coding: utf-8 -*-
"""
leecfdi.py
Lee XML CFDI (3.3/4.0) de una carpeta (recursivo) y genera un Excel resumido.

Uso:
    python leecfdi.py /ruta/a/carpeta_xml reporte_cfdi.xlsx

Dependencias:
    pip install pandas openpyxl
"""

import os
import sys
import pandas as pd
import xml.etree.ElementTree as ET
from datetime import datetime

# URIs fijos de complementos
URI_TFD    = 'http://www.sat.gob.mx/TimbreFiscalDigital'
URI_PAGO20 = 'http://www.sat.gob.mx/Pagos20'

# ---------- Utilidades ----------

def to_float(val, default=0.0):
    try:
        if val is None or val == "":
            return default
        return float(val)
    except Exception:
        return default

def parse_fecha_iso(s):
    if not s:
        return None
    try:
        return datetime.fromisoformat(s.replace('Z', '+00:00'))
    except Exception:
        return None

def colnum_to_excel_name(n):
    name = ""
    n += 1
    while n > 0:
        n, rem = divmod(n - 1, 26)
        name = chr(65 + rem) + name
    return name

def get_attrib(elem, key, default=""):
    if elem is not None:
        return elem.get(key, default) or default
    return default

def xml_namespace_of_tag(tag):
    """Devuelve el namespace URI de un tag '{uri}LocalName' o ''."""
    if tag and tag.startswith('{'):
        return tag[1:].split('}', 1)[0]
    return ''

def listar_xmls_recursivo(carpeta):
    """Busca .xml/.XML recursivamente."""
    paths = []
    for r, _, files in os.walk(carpeta):
        for f in files:
            if f.lower().endswith('.xml'):
                paths.append(os.path.join(r, f))
    return sorted(paths)

# ---------- Lógica principal ----------

def procesar_xml_a_excel(carpeta_xml, archivo_excel):
    datos = []

    # Validación de carpeta
    if not os.path.isdir(carpeta_xml):
        raise FileNotFoundError(f"La carpeta no existe: {carpeta_xml}")

    archivos = listar_xmls_recursivo(carpeta_xml)

    print("📂 Carpeta:", carpeta_xml)
    print("📄 Archivos XML encontrados (recursivo):", len(archivos))
    for f in archivos:
        print("   -", os.path.relpath(f, carpeta_xml))

    if not archivos:
        print("⚠️  No se encontraron archivos .xml en la carpeta indicada (ni subcarpetas).")
        # Continuamos para generar un Excel vacío si quieres; aquí retorno para claridad.
        return _exportar_excel(pd.DataFrame([]), archivo_excel)

    for file_path in archivos:
        print(f"➡️ Procesando archivo: {file_path}")
        try:
            tree = ET.parse(file_path)
            root = tree.getroot()

            # Detectar dinámicamente el namespace del CFDI (3.3 o 4.0)
            cfdi_uri = xml_namespace_of_tag(root.tag) or 'http://www.sat.gob.mx/cfd/4'
            NS = {
                'cfdi': cfdi_uri,
                'tfd':  URI_TFD,
                'pago20': URI_PAGO20
            }

            comp = {}

            # Emisor / Receptor
            emisor = root.find('.//cfdi:Emisor', NS)
            receptor = root.find('.//cfdi:Receptor', NS)

            comp['RFC_EMISOR']      = get_attrib(emisor, 'Rfc')
            comp['NOMBRE_EMISOR']   = get_attrib(emisor, 'Nombre')
            comp['RFC_RECEPTOR']    = get_attrib(receptor, 'Rfc')
            comp['NOMBRE_RECEPTOR'] = get_attrib(receptor, 'Nombre')

            # Atributos del comprobante (nodo raíz)
            comp['FECHA']             = root.get('Fecha', '') or ''
            comp['SERIE']             = root.get('Serie', '') or ''
            comp['FOLIO']             = root.get('Folio', '') or ''
            comp['N.FACTURA']         = f"{comp['SERIE']}-{comp['FOLIO']}".strip('-')
            comp['SUBTOTAL']          = to_float(root.get('SubTotal'), 0)
            comp['T.FACTURA']         = to_float(root.get('Total'), 0)
            comp['Moneda']            = root.get('Moneda', '') or ''
            comp['TipoDeComprobante'] = root.get('TipoDeComprobante', '') or ''
            comp['MetodoPago']        = root.get('MetodoPago', '') or ''
            comp['FormaPago']         = root.get('FormaPago', '') or ''
            comp['UsoCFDI']           = get_attrib(receptor, 'UsoCFDI')
            comp['TipoCambio']        = root.get('TipoCambio', '1') or '1'

            # Impuestos
            iva_total  = 0.0
            ieps_total = 0.0
            risr_total = 0.0
            riva_total = 0.0

            impuestos = root.find('.//cfdi:Impuestos', NS)
            if impuestos is not None:
                for traslado in impuestos.findall('.//cfdi:Traslado', NS):
                    imp = traslado.get('Impuesto')
                    if imp == '002':
                        iva_total += to_float(traslado.get('Importe'), 0)
                    elif imp == '003':
                        ieps_total += to_float(traslado.get('Importe'), 0)

                for retencion in impuestos.findall('.//cfdi:Retencion', NS):
                    imp = retencion.get('Impuesto')
                    if imp == '001':
                        risr_total += to_float(retencion.get('Importe'), 0)
                    elif imp == '002':
                        riva_total += to_float(retencion.get('Importe'), 0)

            comp['IVA']  = iva_total
            comp['IEPS'] = ieps_total
            comp['RISR'] = risr_total
            comp['RIVA'] = riva_total

            # Timbre Fiscal (UUID)
            tfd = root.find('.//tfd:TimbreFiscalDigital', NS)
            comp['UUID'] = get_attrib(tfd, 'UUID')

            # Complemento de Pagos 2.0
            pago_monto = 0.0
            pagos = root.find('.//pago20:Pagos', NS)
            if pagos is not None:
                pago = pagos.find('.//pago20:Pago', NS)
                if pago is not None:
                    pago_monto = to_float(pago.get('Monto'), 0.0)
            comp['P-Monto'] = pago_monto

            # Mes (YYYY-mm) desde FECHA
            fecha_obj = parse_fecha_iso(comp['FECHA'])
            comp['M'] = fecha_obj.strftime('%Y-%m') if fecha_obj else ''

            datos.append(comp)

        except ET.ParseError as e:
            print(f"❌ XML mal formado '{file_path}': {e}")
        except Exception as e:
            print(f"❌ Error procesando '{file_path}': {e}")

    # ---- DataFrame y export ----
    df = pd.DataFrame(datos)
    column_order = [
        'RFC_EMISOR', 'RFC_RECEPTOR', 'NOMBRE_EMISOR', 'FECHA', 'N.FACTURA',
        'SUBTOTAL', 'IVA', 'RISR', 'RIVA', 'T.FACTURA', 'M', 'UUID', 'IEPS',
        'P-Monto', 'MetodoPago', 'TipoDeComprobante', 'FormaPago', 'UsoCFDI', 'TipoCambio'
    ]
    column_order = [c for c in column_order if c in df.columns]
    df = df[column_order] if not df.empty else df

    rename_map = {
        'RFC_EMISOR': 'RFC EMISOR',
        'RFC_RECEPTOR': 'RFC RECEPTOR',
        'NOMBRE_EMISOR': 'NOMBRE',
        'MetodoPago': 'MetodoP',
        'TipoDeComprobante': 'TipoCFDI',
        'FormaPago': 'FPago',
        'TipoCambio': 'Tcambio'
    }
    df = df.rename(columns=rename_map)

    _exportar_excel(df, archivo_excel)
    print(f"🧾 Total de archivos procesados: {len(datos)}")

def _exportar_excel(df, archivo_excel):
    with pd.ExcelWriter(archivo_excel, engine='openpyxl') as writer:
        sheet_name = 'Reporte CFDI'
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        ws = writer.sheets[sheet_name]
        for idx, col in enumerate(df.columns):
            try:
                max_len_cells = df[col].astype(str).map(len).max() if not df.empty else 0
            except Exception:
                max_len_cells = 0
            max_len = max(len(str(col)), int(max_len_cells)) + 2
            excel_col = colnum_to_excel_name(idx)
            ws.column_dimensions[excel_col].width = max_len
    print(f"✅ Reporte generado: {archivo_excel}")

# ---------- CLI simple ----------
def main():
    # Defaults útiles para debug si no pasas args
    # default_in  = "/home/gustavo/Público/proyectos/leexml/LILI/compras/"
    default_in = "/xml/recepcion/compras/"
    default_out = "reporte_cfdi.xlsx"

    if len(sys.argv) < 3:
        print("⚠️  No se proporcionaron argumentos. Usando valores por defecto:")
        print("   - carpeta_xml:", default_in)
        print("   - archivo_excel:", default_out)
        carpeta_xml = default_in
        archivo_excel = default_out
    else:
        carpeta_xml = sys.argv[1]
        archivo_excel = sys.argv[2]

    procesar_xml_a_excel(carpeta_xml, archivo_excel)

if __name__ == "__main__":
    main()
