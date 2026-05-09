import csv
import io
import os

from reportlab.lib.colors import black, HexColor
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas
import qrcode

MM = 2.8346  # 1 mm en puntos


def pt(mm):
    return mm * MM


# Dimensiones del sticker
W = pt(100)
H = pt(25)

# Límites de zonas (en puntos desde la izquierda)
LOGO_END  = pt(21)
INFO_END  = pt(66)
PIECE_END = pt(80)

GRAY      = HexColor('#CCCCCC')
LOGO_PATH = os.path.join(os.path.dirname(__file__), 'static', 'img',
                         'ISOTIPO-INTERDOORS-POSITIVO.png')


def _sanitize_text(text: str) -> str:
    """Limpia caracteres problemáticos para ReportLab."""
    if not isinstance(text, str):
        text = str(text)
    # Reemplazar caracteres especiales comunes
    text = text.replace('á', 'a').replace('é', 'e').replace('í', 'i')
    text = text.replace('ó', 'o').replace('ú', 'u').replace('ñ', 'n')
    text = text.replace('Á', 'A').replace('É', 'E').replace('Í', 'I')
    text = text.replace('Ó', 'O').replace('Ú', 'U').replace('Ñ', 'N')
    # Limitar a 40 caracteres si es muy largo
    return text[:40].strip()


def _make_qr(data: str) -> ImageReader:
    qr = qrcode.QRCode(
        version=1, border=1,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
    )
    qr.add_data(data)
    qr.make(fit=True)
    buf = io.BytesIO()
    qr.make_image(fill_color='black', back_color='white').save(buf, format='PNG')
    buf.seek(0)
    return ImageReader(buf)


def _draw_sticker(c: canvas.Canvas, cliente: str, doc: str,
                  medida: str, pieza: str) -> None:

    # Sanitizar textos
    cliente = _sanitize_text(cliente)
    doc = _sanitize_text(doc)
    medida = _sanitize_text(medida)
    pieza = _sanitize_text(pieza)

    # ── Borde exterior redondeado ─────────────────────────────────────
    c.setStrokeColor(GRAY)
    c.setLineWidth(0.5)
    c.roundRect(0.5, 0.5, W - 1, H - 1, radius=pt(2), stroke=1, fill=0)

    # ── Zona logo (0 → 21 mm) ─────────────────────────────────────────
    if os.path.exists(LOGO_PATH):
        logo_size = pt(16)
        c.drawImage(
            LOGO_PATH,
            (LOGO_END - logo_size) / 2,
            (H - logo_size) / 2,
            logo_size, logo_size,
            preserveAspectRatio=True, mask='auto',
        )

    # Separador vertical derecho del logo
    c.setStrokeColor(GRAY)
    c.setLineWidth(0.5)
    c.line(LOGO_END, pt(1.5), LOGO_END, H - pt(1.5))

    # ── Zona info (21 → 66 mm) ────────────────────────────────────────
    ix  = LOGO_END
    iw  = INFO_END - LOGO_END
    sec = H / 3          # altura de cada sección
    pad = pt(2.5)

    c.setFont('Helvetica-Bold', 9)
    c.setFillColor(black)

    for i, texto in enumerate([cliente, str(doc), medida]):
        # i=0 → sección superior; i=2 → sección inferior
        y_bot    = H - (i + 1) * sec
        y_center = y_bot + sec / 2
        c.drawString(ix + pad, y_center - pt(1.6), texto)

        # Línea divisoria gris entre secciones
        if i < 2:
            c.setStrokeColor(GRAY)
            c.setLineWidth(0.3)
            c.line(ix, y_bot, ix + iw, y_bot)

    # Acento negro inferior de la zona info
    c.setStrokeColor(black)
    c.setLineWidth(1.2)
    c.line(ix, pt(1.5), ix + iw, pt(1.5))

    # ── Zona pieza (66 → 80 mm) ───────────────────────────────────────
    px = INFO_END
    pw = PIECE_END - INFO_END

    # Separadores verticales a ambos lados
    c.setStrokeColor(GRAY)
    c.setLineWidth(0.5)
    c.line(px,      pt(1.5), px,      H - pt(1.5))
    c.line(px + pw, pt(1.5), px + pw, H - pt(1.5))

    # Número de pieza centrado
    c.setFont('Helvetica-Bold', 12)
    c.setFillColor(black)
    tw = c.stringWidth(pieza, 'Helvetica-Bold', 12)
    c.drawString(px + (pw - tw) / 2, H / 2 - pt(2.2), pieza)

    # ── Zona QR (80 → 100 mm) ─────────────────────────────────────────
    qx      = PIECE_END
    qw      = W - PIECE_END
    qr_size = min(qw, H) - pt(4)
    qr_x    = qx + (qw - qr_size) / 2
    qr_y    = (H - qr_size) / 2

    qr_data = f"Interdoors|{cliente}|{doc}|{medida}|{pieza}"
    c.drawImage(_make_qr(qr_data), qr_x, qr_y, qr_size, qr_size)


def generate_pdf(csv_path: str = 'stickers.csv',
                 output_path: str = 'stickers_output.pdf') -> None:
    c = canvas.Canvas(output_path, pagesize=(W, H))
    total = 0
    with open(csv_path, newline='', encoding='utf-8-sig') as f:
        for row in csv.DictReader(f):
            _draw_sticker(c, row['cliente'], row['doc'],
                          row['medida'], row['pieza'])
            c.showPage()
            total += 1
    c.save()
    print(f"[OK] {total} stickers -> {output_path}")


def verify_installation():
    """Verifica que el sistema esté correctamente configurado."""
    import os
    print("[CHECK] Verificacion de instalacion:")
    
    # Archivos requeridos
    files = {
        'stickers.csv': 'Datos de entrada',
        'static/img/ISOTIPO-INTERDOORS-POSITIVO.png': 'Logo',
    }
    
    for fpath, desc in files.items():
        status = "[OK]" if os.path.exists(fpath) else "[NO]"
        print(f"  {status} {fpath:40} -> {desc}")
    
    # Dependencias
    packages = {
        'reportlab': 'PDF',
        'qrcode': 'QR',
        'PIL': 'Imagenes',
    }
    
    for pkg, desc in packages.items():
        try:
            __import__(pkg)
            print(f"  [OK] {pkg:20} -> {desc}")
        except ImportError:
            print(f"  [NO] {pkg:20} -> {desc} [FALTA]")


def export_stickers_from_orders(orders_data, output_file='stickers_output.pdf'):
    """Helper para integración con Flask: genera PDF desde lista de datos."""
    import os
    csv_temp = 'temp_stickers.csv'
    
    try:
        # Sanitizar datos antes de procesar
        clean_data = [
            {
                'cliente': _sanitize_text(row.get('cliente', '')),
                'doc': _sanitize_text(row.get('doc', '')),
                'medida': _sanitize_text(row.get('medida', '')),
                'pieza': _sanitize_text(row.get('pieza', '')),
            }
            for row in orders_data
        ]
        
        with open(csv_temp, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=['cliente', 'doc', 'medida', 'pieza'])
            writer.writeheader()
            writer.writerows(clean_data)
        
        generate_pdf(csv_temp, output_file)
        return output_file
    finally:
        if os.path.exists(csv_temp):
            os.remove(csv_temp)


if __name__ == '__main__':
    import sys
    
    if len(sys.argv) > 1 and sys.argv[1] == 'verify':
        verify_installation()
    else:
        generate_pdf()
