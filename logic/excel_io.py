import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


_HEADER_MAPPING = {
    "cliente": "cliente",
    "referencia": "referencia",
    "cantidad": "cantidad",
    "notas del ítem": "notas_item",
    "notas del item": "notas_item",
    "notas_item": "notas_item",
    "notas generales": "notas_generales",
    "notas_generales": "notas_generales",
}


def read_input_excel(filepath):
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active

    raw_headers = [cell.value for cell in ws[1]]
    headers = [str(h).lower().strip() if h else "" for h in raw_headers]

    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(v for v in row if v is not None):
            continue
        row_dict = {}
        for i, h in enumerate(headers):
            key = _HEADER_MAPPING.get(h, h)
            row_dict[key] = row[i] if i < len(row) else None
        rows.append({
            "cliente": str(row_dict.get("cliente", "") or ""),
            "referencia": str(row_dict.get("referencia", "") or ""),
            "cantidad": int(row_dict.get("cantidad", 0) or 0),
            "notas_item": str(row_dict.get("notas_item", "") or ""),
            "notas_generales": str(row_dict.get("notas_generales", "") or ""),
        })
    return rows


def _apply_headers(ws, headers, fill_color):
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    fill = PatternFill("solid", fgColor=fill_color)
    font = Font(bold=True, color="FFFFFF", size=11)
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
    ws.row_dimensions[1].height = 20


def _auto_width(ws):
    for col in ws.columns:
        max_len = max((len(str(cell.value)) for cell in col if cell.value), default=8)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 40)


def write_erp_excel(opp_rows, target):
    from datetime import date
    today = date.today().strftime("%Y%m%d")

    wb = openpyxl.Workbook()

    # --- Hoja Documentos ---
    ws_doc = wb.active
    ws_doc.title = "Documentos"
    doc_headers = [
        "CONSECUTIVO DCTO", "FECHA AAAAMMDD", "PLANIFICADOR",
        "REF1", "REF2", "REF3", "NOTAS",
    ]
    _apply_headers(ws_doc, doc_headers, "1F4E79")
    for r, row in enumerate(opp_rows, 2):
        ws_doc.cell(row=r, column=1, value=row["OPP"])
        ws_doc.cell(row=r, column=2, value=today)
        ws_doc.cell(row=r, column=3, value="")           # PLANIFICADOR — pendiente
        ws_doc.cell(row=r, column=4, value=row["Cliente"])  # REF1
        ws_doc.cell(row=r, column=5, value="")           # REF2 — pendiente
        ws_doc.cell(row=r, column=6, value="")           # REF3 — pendiente
        ws_doc.cell(row=r, column=7, value=row["notas_generales"])
    _auto_width(ws_doc)

    # --- Hoja Items ---
    ws_items = wb.create_sheet("Items")
    item_headers = [
        "NUMERO DCTO", "REGISTRO MVTO", "REFERENCIA", "EXT1", "EXT2",
        "U.M", "CANT PLANEADA", "FECHA INICIO AAAAMMDD", "FECHA TERMINACION AAAAMMDD",
        "METODO LISTA", "METODO RUTA", "MEDIDA REAL", "BODEGA",
    ]
    _apply_headers(ws_items, item_headers, "1F4E79")
    for r, row in enumerate(opp_rows, 2):
        ws_items.cell(row=r, column=1,  value=row["OPP"])
        ws_items.cell(row=r, column=2,  value="")             # REGISTRO MVTO — pendiente
        ws_items.cell(row=r, column=3,  value=row["Referencia"])
        ws_items.cell(row=r, column=4,  value="")             # EXT1 — pendiente
        ws_items.cell(row=r, column=5,  value="")             # EXT2 — pendiente
        ws_items.cell(row=r, column=6,  value="")             # U.M — pendiente
        ws_items.cell(row=r, column=7,  value=row["Cantidad"])
        ws_items.cell(row=r, column=8,  value="")             # FECHA INICIO — pendiente
        ws_items.cell(row=r, column=9,  value="")             # FECHA TERMINACION — pendiente
        ws_items.cell(row=r, column=10, value="")             # METODO LISTA — pendiente
        ws_items.cell(row=r, column=11, value=row["Proceso"]) # METODO RUTA
        ws_items.cell(row=r, column=12, value=row["notas_item"])  # MEDIDA REAL
        ws_items.cell(row=r, column=13, value="")             # BODEGA — pendiente
    _auto_width(ws_items)

    wb.save(target)


def write_sticker_excel(sticker_rows, target):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Stickers"
    headers = ["Cliente", "Numero de documento", "Medida real", "Numero de pieza", "Cantidad"]
    _apply_headers(ws, headers, "375623")
    for r, row in enumerate(sticker_rows, 2):
        for col, key in enumerate(headers, 1):
            ws.cell(row=r, column=col, value=row[key])
    _auto_width(ws)
    wb.save(target)


def create_estructura_template(target):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Estructura"
    headers = ["Referencia", "Descripción", "Proceso 1", "Proceso 2", "Proceso 3",
               "Proceso 4", "Proceso 5", "Proceso 6", "Proceso 7", "Proceso 8"]
    _apply_headers(ws, headers, "2E75B6")
    sample_rows = [
        ["REF001", "Puerta Principal Madera", "Corte", "Lijado", "Pintura", "Terminado", "Empaque", "", "", ""],
        ["REF002", "Marco Metálico", "Corte Metal", "Soldadura", "Pintura", "Empaque", "", "", "", ""],
        ["REF003", "Panel Decorativo", "Corte", "Ensamble", "Lacado", "Control Calidad", "Empaque", "", "", ""],
    ]
    for r, row in enumerate(sample_rows, 2):
        for col, val in enumerate(row, 1):
            ws.cell(row=r, column=col, value=val)
    col_widths = [15, 28, 18, 18, 18, 18, 18, 18, 18, 18]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    wb.save(target)


def read_estructura_excel(filepath):
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active
    referencias = {}
    errors = []
    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
        ref = str(row[0]).strip() if row[0] else ""
        if not ref or ref.lower() == "none":
            continue
        descripcion = str(row[1]).strip() if row[1] else ""
        procesos = [str(v).strip() for v in row[2:] if v and str(v).strip()]
        if not procesos:
            errors.append(f"Fila {row_idx}: '{ref}' no tiene procesos definidos — omitida.")
            continue
        referencias[ref] = {"descripcion": descripcion, "procesos": procesos}
    return referencias, errors


def create_input_template(target):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Entrada"
    headers = ["Cliente", "Referencia", "Cantidad", "Notas del ítem", "Notas generales"]
    _apply_headers(ws, headers, "2E75B6")
    sample_rows = [
        ["CLIENTE A", "REF001", 5, "Medida: 50x30 cm", "Pedido urgente"],
        ["CLIENTE B", "REF002", 3, "Medida: 40x20 cm", ""],
    ]
    for r, row in enumerate(sample_rows, 2):
        for col, val in enumerate(row, 1):
            ws.cell(row=r, column=col, value=val)
    for i, letter in enumerate("ABCDE"):
        ws.column_dimensions[letter].width = [20, 15, 12, 30, 25][i]
    wb.save(target)
