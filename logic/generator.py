import json
from pathlib import Path

BASE_DIR = Path(__file__).parent.parent
COUNTER_FILE = BASE_DIR / "data" / "opp_counter.json"
ESTRUCTURA_FILE = BASE_DIR / "config" / "estructura.json"


def load_estructura():
    if not ESTRUCTURA_FILE.exists():
        return {}
    with open(ESTRUCTURA_FILE, encoding="utf-8") as f:
        return json.load(f)


def save_estructura(data):
    ESTRUCTURA_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(ESTRUCTURA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def _next_opp_number():
    COUNTER_FILE.parent.mkdir(parents=True, exist_ok=True)
    if COUNTER_FILE.exists():
        with open(COUNTER_FILE) as f:
            data = json.load(f)
    else:
        data = {"last": 0}
    data["last"] += 1
    with open(COUNTER_FILE, "w") as f:
        json.dump(data, f)
    return data["last"]


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

        for proceso in procesos:
            opp_num = _next_opp_number()
            opp_rows.append({
                "Tipo": tipo_opp,
                "OPP": opp_num,
                "Cliente": cliente,
                "Referencia": ref,
                "Proceso": proceso,
                "Cantidad": cantidad,
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
