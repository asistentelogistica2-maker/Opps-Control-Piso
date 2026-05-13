import os
import io
import uuid
from flask import Flask, render_template, request, send_file, redirect, url_for, flash
from logic.generator import load_estructura, save_estructura, generate_opps_stock, load_referencias_stock
from logic.excel_io import (
    read_input_excel, create_input_template, write_jumbo_excel,
    create_estructura_template, read_estructura_excel,
    create_referencias_template, read_referencias_excel,
)

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'dev-secret-key')

_cache = {}
_referencias_stock = None


def get_referencias_stock():
    global _referencias_stock
    if _referencias_stock is None:
        _referencias_stock = load_referencias_stock()
    return _referencias_stock


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/generar', methods=['POST'])
def generar():
    archivo = request.files.get('archivo')
    if not archivo or archivo.filename == '':
        flash('Seleccione un archivo Excel de entrada.', 'warning')
        return redirect(url_for('index'))

    tipo_opp = request.form.get('tipo_opp', 'Stock')

    if tipo_opp != 'Stock':
        flash('Este tipo de OPP aún no está disponible.', 'info')
        return redirect(url_for('index'))

    try:
        stream = io.BytesIO(archivo.read())
        rows = read_input_excel(stream)

        if not rows:
            flash('El archivo Excel no tiene datos.', 'warning')
            return redirect(url_for('index'))

        referencias = get_referencias_stock()
        if not referencias:
            flash('No se encontró el archivo de Referencias Stock.', 'danger')
            return redirect(url_for('index'))

        opp_list, errors = generate_opps_stock(rows, referencias)

        if not opp_list:
            flash('No se generaron OPPs. Verifique las referencias y colores.', 'warning')
            for e in errors:
                flash(e, 'warning')
            return redirect(url_for('index'))

        token = str(uuid.uuid4())
        buf = io.BytesIO()
        write_jumbo_excel(opp_list, buf)
        _cache[f"{token}_jumbo"] = buf.getvalue()

        return render_template(
            'resultados.html',
            opp_list=opp_list,
            errors=errors,
            token=token,
            tipo_opp=tipo_opp,
        )

    except Exception as exc:
        flash(f'Error al procesar el archivo: {exc}', 'danger')
        return redirect(url_for('index'))


@app.route('/descargar/<token>/<tipo>')
def descargar(token, tipo):
    data = _cache.get(f"{token}_{tipo}")
    if not data:
        flash('Archivo no disponible. Vuelva a generar.', 'warning')
        return redirect(url_for('index'))

    if tipo == 'jumbo':
        filename = "Plano OPP Generic.xlsx"
        mimetype = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    elif tipo == 'stickers':
        filename = 'stickers.pdf'
        mimetype = 'application/pdf'
    else:
        filename = 'archivo.xlsx'
        mimetype = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

    return send_file(
        io.BytesIO(data),
        download_name=filename,
        as_attachment=True,
        mimetype=mimetype,
    )


@app.route('/plantilla')
def plantilla():
    buf = io.BytesIO()
    create_input_template(buf)
    buf.seek(0)
    return send_file(
        buf,
        download_name='Generar OPP.xlsx',
        as_attachment=True,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )


@app.route('/referencias/plantilla')
def plantilla_referencias():
    from logic.firebase_db import load_referencias, is_available
    data = load_referencias() if is_available() else None
    buf = io.BytesIO()
    create_referencias_template(buf, data=data)
    buf.seek(0)
    return send_file(
        buf,
        download_name='plantilla_referencias.xlsx',
        as_attachment=True,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )


@app.route('/referencias/importar', methods=['POST'])
def importar_referencias():
    from logic.firebase_db import save_referencias
    archivo = request.files.get('archivo_referencias')
    if not archivo or archivo.filename == '':
        flash('Seleccione un archivo Excel para importar.', 'warning')
        return redirect(url_for('estructura'))

    modo = request.form.get('modo_referencias', 'merge')

    try:
        stream = io.BytesIO(archivo.read())
        nuevas, errors = read_referencias_excel(stream)

        if not nuevas:
            flash('El archivo no contiene referencias válidas.', 'warning')
            return redirect(url_for('estructura'))

        save_referencias(nuevas, modo)
        global _referencias_stock
        _referencias_stock = None

        msg = f'{len(nuevas)} referencia(s) importada(s) correctamente.'
        if errors:
            msg += f' {len(errors)} fila(s) omitida(s).'
        flash(msg, 'success')

    except Exception as exc:
        flash(f'Error al leer el archivo: {exc}', 'danger')

    return redirect(url_for('estructura'))


@app.route('/estructura')
def estructura():
    from logic.firebase_db import load_referencias, is_available
    referencias = load_referencias() if is_available() else {}
    return render_template('estructura.html', estructura=load_estructura(), referencias=referencias)


@app.route('/estructura/guardar', methods=['POST'])
def guardar_estructura():
    ref = request.form.get('ref', '').strip()
    descripcion = request.form.get('descripcion', '').strip()
    procesos_raw = request.form.get('procesos', '').strip()
    procesos = [p.strip() for p in procesos_raw.splitlines() if p.strip()]

    if not ref:
        flash('El código de referencia es obligatorio.', 'danger')
        return redirect(url_for('estructura'))
    if not procesos:
        flash('Ingrese al menos un proceso.', 'danger')
        return redirect(url_for('estructura'))

    data = load_estructura()
    data[ref] = {'descripcion': descripcion, 'procesos': procesos}
    save_estructura(data)
    flash(f"Referencia '{ref}' guardada correctamente.", 'success')
    return redirect(url_for('estructura'))


@app.route('/estructura/plantilla-masiva')
def plantilla_estructura():
    data = load_estructura()
    buf = io.BytesIO()
    create_estructura_template(buf, data=data)
    buf.seek(0)
    return send_file(
        buf,
        download_name='plantilla_estructura_productiva.xlsx',
        as_attachment=True,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )


@app.route('/estructura/importar', methods=['POST'])
def importar_estructura():
    archivo = request.files.get('archivo_masivo')
    if not archivo or archivo.filename == '':
        flash('Seleccione un archivo Excel para importar.', 'warning')
        return redirect(url_for('estructura'))

    modo = request.form.get('modo_importar', 'merge')

    try:
        stream = io.BytesIO(archivo.read())
        nuevas, errors = read_estructura_excel(stream)

        if not nuevas:
            flash('El archivo no contiene referencias válidas.', 'warning')
            return redirect(url_for('estructura'))

        data = {} if modo == 'reemplazar' else load_estructura()
        data.update(nuevas)
        save_estructura(data)

        msg = f'{len(nuevas)} referencia(s) importada(s) correctamente.'
        if errors:
            msg += f' {len(errors)} fila(s) omitida(s) por errores.'
        flash(msg, 'success')

        for e in errors:
            flash(e, 'warning')

    except Exception as exc:
        flash(f'Error al leer el archivo: {exc}', 'danger')

    return redirect(url_for('estructura'))


@app.route('/estructura/eliminar', methods=['POST'])
def eliminar_referencia():
    ref = request.form.get('ref', '').strip()
    data = load_estructura()
    if ref in data:
        del data[ref]
        save_estructura(data)
        flash(f"Referencia '{ref}' eliminada.", 'success')
    return redirect(url_for('estructura'))


if __name__ == '__main__':
    app.run(debug=True)
