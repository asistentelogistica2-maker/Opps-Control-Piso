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


def _next_day_skip_sunday(d):
    """Suma 1 día hábil: si el resultado es domingo, avanza al lunes."""
    result = d + timedelta(days=1)
    if result.weekday() == 6:  # domingo → lunes
        result += timedelta(days=1)
    return result


def _split_quantity(total, max_qty):
    """Divide total en lotes de máximo max_qty. Sin límite si max_qty es None o 0."""
    if not max_qty or max_qty <= 0:
        return [total]
    chunks = []
    remaining = total
    while remaining > 0:
        chunks.append(min(remaining, max_qty))
        remaining -= min(remaining, max_qty)
    return chunks


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

        # Fechas proceso 1: todas las OPPs del proceso van en paralelo (mismas fechas)
        opp1_inicio_dt = fecha_dt
        opp1_fin_dt    = _next_day_skip_sunday(opp1_inicio_dt)

        tiene_dos = bool(ref_data["ref1"]) and bool(ref_data["ref2_i"])

        # Límites de cantidad por proceso (None = sin límite)
        max_p1   = ref_data.get("max_p1")
        max_p2   = ref_data.get("max_p2")
        max_p1   = int(max_p1) if max_p1 else None
        max_p2   = int(max_p2) if max_p2 else None

        ref_item_p1 = ref_data["referencia_b"] if tiene_dos else ref_data["referencia_a"]
        bodega_p1   = "80106" if tiene_dos else "80123"

        for qty in _split_quantity(cantidad, max_p1):
            counter += 1
            opp_list.append({
                "opp":            counter,
                "fecha":          fecha_str,
                "planificador":   "71364487",
                "ref1":           ref_data["ref1"],
                "ref2":           ref_data["ref2_j"],
                "notas":          ref_data["notas1"],
                "referencia_item": ref_item_p1,
                "ext1":           ref_data["color_num"],
                "ext2":           ref_data["medida"],
                "um":             ref_data["um"],
                "cantidad":       qty,
                "fecha_inicio":   opp1_inicio_dt.strftime("%Y%m%d"),
                "fecha_fin":      opp1_fin_dt.strftime("%Y%m%d"),
                "bodega":         bodega_p1,
            })

        if tiene_dos:
            # Fechas proceso 2: encadenadas desde el fin del proceso 1
            opp2_inicio_dt = _next_day_skip_sunday(opp1_fin_dt)
            opp2_fin_dt    = _next_day_skip_sunday(opp2_inicio_dt)

            for qty in _split_quantity(cantidad, max_p2):
                counter += 1
                opp_list.append({
                    "opp":            counter,
                    "fecha":          fecha_str,
                    "planificador":   "71364487",
                    "ref1":           ref_data["ref2_i"],
                    "ref2":           ref_data["ref2_j"],
                    "notas":          ref_data["notas2"],
                    "referencia_item": ref_data["referencia_a"],
                    "ext1":           ref_data["color_num"],
                    "ext2":           ref_data["medida"],
                    "um":             ref_data["um"],
                    "cantidad":       qty,
                    "fecha_inicio":   opp2_inicio_dt.strftime("%Y%m%d"),
                    "fecha_fin":      opp2_fin_dt.strftime("%Y%m%d"),
                    "bodega":         "80123",
                })

    return opp_list, errors
