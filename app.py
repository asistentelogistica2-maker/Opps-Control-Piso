import os
import io
import uuid
import tempfile
from flask import Flask, render_template, request, send_file, redirect, url_for, flash
from logic.generator import load_estructura, save_estructura, generate_opps
from logic.excel_io import read_input_excel, write_erp_excel, create_input_template, create_estructura_template, read_estructura_excel
from generate_stickers_pdf import export_stickers_from_orders

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'dev-secret-key')

# In-memory cache for generated files (cleared on restart)
_cache = {}


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/generar', methods=['POST'])
def generar():
    archivo = request.files.get('archivo')
    if not archivo or archivo.filename == '':
        flash('Seleccione un archivo Excel de entrada.', 'warning')
        return redirect(url_for('index'))

    gen_erp = 'gen_erp' in request.form
    gen_stickers = 'gen_stickers' in request.form
    tipo_opp = request.form.get('tipo_opp', 'Stock')

    if not gen_erp and not gen_stickers:
        flash('Seleccione al menos una opción de generación.', 'warning')
        return redirect(url_for('index'))

    try:
        estructura = load_estructura()
        if not estructura:
            flash('La estructura productiva está vacía. Configure las referencias primero.', 'warning')
            return redirect(url_for('estructura'))

        stream = io.BytesIO(archivo.read())
        rows = read_input_excel(stream)

        if not rows:
            flash('El archivo Excel no tiene datos.', 'warning')
            return redirect(url_for('index'))

        opp_rows, sticker_rows, errors = generate_opps(rows, estructura, tipo_opp)
        token = str(uuid.uuid4())

        if gen_erp and opp_rows:
            buf = io.BytesIO()
            write_erp_excel(opp_rows, buf)
            _cache[f"{token}_erp"] = buf.getvalue()

        if gen_stickers and sticker_rows:
            # Convertir formato sticker_rows a formato para PDF
            stickers_data = [
                {
                    'cliente': row['Cliente'],
                    'doc': str(row['Numero de documento']),
                    'medida': row['Medida real'],
                    'pieza': row['Numero de pieza'],
                }
                for row in sticker_rows
            ]
            # Generar PDF en lugar de Excel
            with tempfile.NamedTemporaryFile(mode='wb', delete=False, suffix='.pdf') as tmp:
                temp_pdf = tmp.name
            export_stickers_from_orders(stickers_data, temp_pdf)
            with open(temp_pdf, 'rb') as f:
                _cache[f"{token}_stickers"] = f.read()
            os.remove(temp_pdf)

        return render_template(
            'resultados.html',
            opp_rows=opp_rows,
            sticker_rows=sticker_rows,
            errors=errors,
            token=token,
            tipo_opp=tipo_opp,
            has_erp=gen_erp and bool(opp_rows),
            has_stickers=gen_stickers and bool(sticker_rows),
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
    
    if tipo == 'erp':
        filename = 'OPP_ERP.xlsx'
        mimetype = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    else:  # stickers
        filename = 'stickers.pdf'
        mimetype = 'application/pdf'
    
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
        download_name='plantilla_entrada_opp.xlsx',
        as_attachment=True,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )


@app.route('/estructura')
def estructura():
    return render_template('estructura.html', estructura=load_estructura())


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
    buf = io.BytesIO()
    create_estructura_template(buf)
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
