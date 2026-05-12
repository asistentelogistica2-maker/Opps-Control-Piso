import json
import os
from datetime import date, datetime, timedelta
from pathlib import Path

BASE_DIR = Path(__file__).parent.parent
ESTRUCTURA_FILE = BASE_DIR / "config" / "estructura.json"

try:
    from logic import firebase_db as _fdb
except ImportError:
    _fdb = None


def _use_firebase():
    return _fdb is not None and _fdb.is_available()


def load_estructura():
    if _use_firebase():
        return _fdb.load_estructura()
    if not ESTRUCTURA_FILE.exists():
        return {}
    with open(ESTRUCTURA_FILE, encoding="utf-8") as f:
        return json.load(f)


def save_estructura(data):
    if _use_firebase():
        _fdb.save_estructura(data)
        return
    ESTRUCTURA_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(ESTRUCTURA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)



def generate_opps(input_rows, estructura, tipo_opp="Stock"):
    """Returns (opp_rows, sticker_rows, errors)."""
    opp_rows = []
    sticker_rows = []
    errors = []

    for row in input_rows:
        ref = str(row.get("referencia", "")).strip()
        cliente = str(row.get("cliente", "")).strip()
        cantidad = int(row.get("cantidad", 0) or 0)
        notas_item = str(row.get("notas_item", "") or "")

        if not ref:
            continue

        if ref not in estructura:
            errors.append(f"Referencia '{ref}' no encontrada en la estructura productiva.")
            continue

        procesos = estructura[ref].get("procesos", [])
        if not procesos:
            errors.append(f"Referencia '{ref}' no tiene procesos definidos.")
            continue

        last_proceso = procesos[-1]

        notas_generales = str(row.get("notas_generales", "") or "")

        for proceso in procesos:
            opp_num = _next_opp_number()
            opp_rows.append({
                "Tipo": tipo_opp,
                "OPP": opp_num,
                "Cliente": cliente,
                "Referencia": ref,
                "Proceso": proceso,
                "Cantidad": cantidad,
                "notas_item": notas_item,
                "notas_generales": notas_generales,
            })

            if proceso == last_proceso:
                for pieza in range(1, cantidad + 1):
                    sticker_rows.append({
                        "Cliente": cliente,
                        "Numero de documento": opp_num,
                        "Medida real": notas_item,
                        "Numero de pieza": f"{pieza}/{cantidad}",
                        "Cantidad": cantidad,
                    })

    return opp_rows, sticker_rows, errors


def _safe_key(s):
    for ch in ['$', '#', '[', ']', '/', '.']:
        s = s.replace(ch, '-')
    return s.strip().upper()


def load_referencias_stock():
    if _use_firebase():
        raw = _fdb.load_referencias()
        lookup = {}
        for fb_key, data in raw.items():
            ref, color = fb_key.split('|', 1)
            lookup[(_safe_key(ref), _safe_key(color))] = data
        return lookup
    return {}


def generate_opps_stock(input_rows, referencias_lookup):
    opp_list = []
    errors = []
    counter = 0

    for row in input_rows:
        fecha_raw = row.get("fecha")
        ref_input = str(row.get("referencia", "") or "").strip()
        color_input = str(row.get("color", "") or "").strip()
        cantidad = int(row.get("cantidad", 0) or 0)

        if not ref_input or not color_input:
            continue

        key = (_safe_key(ref_input), _safe_key(color_input))
        ref_data = referencias_lookup.get(key)
        if not ref_data:
            errors.append(f"Referencia '{ref_input}' con color '{color_input}' no encontrada.")
            continue

        if isinstance(fecha_raw, datetime):
            fecha_dt = fecha_raw.date()
        elif isinstance(fecha_raw, date):
            fecha_dt = fecha_raw
        else:
            fecha_dt = None
            for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d", "%Y%m%d"):
                try:
                    fecha_dt = datetime.strptime(str(fecha_raw).strip(), fmt).date()
                    break
                except Exception:
                    continue
            if fecha_dt is None:
                fecha_dt = date.today()

        fecha_str = fecha_dt.strftime("%Y%m%d")
        fecha_mas2 = (fecha_dt + timedelta(days=2)).strftime("%Y%m%d")

        tiene_dos = bool(ref_data["ref1"]) and bool(ref_data["ref2_i"])

        counter += 1
        opp_num1 = counter
        opp_list.append({
            "opp": opp_num1,
            "fecha": fecha_str,
            "planificador": "71364487",
            "ref1": ref_data["ref1"],
            "ref2": ref_data["ref2_j"],
            "notas": ref_data["notas1"],
            "referencia_item": ref_data["referencia_b"] if tiene_dos else ref_data["referencia_a"],
            "ext1": ref_data["color_num"],
            "ext2": ref_data["medida"],
            "um": ref_data["um"],
            "cantidad": cantidad,
            "fecha_inicio": fecha_str,
            "fecha_fin": fecha_str,
            "bodega": "80106" if tiene_dos else "80123",
        })

        if tiene_dos:
            counter += 1
            opp_num2 = counter
            opp_list.append({
                "opp": opp_num2,
                "fecha": fecha_str,
                "planificador": "71364487",
                "ref1": ref_data["ref2_i"],
                "ref2": ref_data["ref2_j"],
                "notas": ref_data["notas2"],
                "referencia_item": ref_data["referencia_a"],
                "ext1": ref_data["color_num"],
                "ext2": ref_data["medida"],
                "um": ref_data["um"],
                "cantidad": cantidad,
                "fecha_inicio": fecha_mas2,
                "fecha_fin": fecha_mas2,
                "bodega": "80123",
            })

    return opp_list, errors
