import os
from datetime import datetime, time
from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, send_file
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, logout_user, login_required, current_user
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from weasyprint import HTML
import tempfile

app = Flask(__name__)
app.config['SECRET_KEY'] = 'migab2026_secret_key_prod'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:////home/MIGAB2026/database.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)
login_manager = LoginManager(app)
login_manager.login_view = 'login'

class Usuario(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(50), unique=True, nullable=False)
    password = db.Column(db.String(100), nullable=False)

class Vehiculo(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    placa = db.Column(db.String(20), unique=True, nullable=False)
    modelo = db.Column(db.String(50), nullable=False)
    capacidad = db.Column(db.Float, nullable=False)

class Chofer(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nombre = db.Column(db.String(100), nullable=False)
    cedula = db.Column(db.String(20), unique=True, nullable=False)
    telefono = db.Column(db.String(20), unique=True, nullable=False)

class Viaje(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    origen = db.Column(db.String(100), nullable=False)
    destino = db.Column(db.String(100), nullable=False)
    vehiculo_id = db.Column(db.Integer, db.ForeignKey('vehiculo.id'), nullable=False)
    chofer_id = db.Column(db.Integer, db.ForeignKey('chofer.id'), nullable=False)
    flete = db.Column(db.Float, default=0.0, nullable=False)
    fecha = db.Column(db.DateTime, default=datetime.now, nullable=False)

    vehiculo = db.relationship('Vehiculo', backref='viajes')
    chofer = db.relationship('Chofer', backref='viajes')

class RegistroGPS(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    telefono = db.Column(db.String(20), nullable=False)
    latitud = db.Column(db.Float, nullable=False)
    longitud = db.Column(db.Float, nullable=False)
    fecha = db.Column(db.DateTime, default=datetime.now)

@login_manager.user_loader
def load_user(user_id):
    return Usuario.query.get(int(user_id))

@app.route('/')
def index():
    if current_user.is_authenticated:
        return redirect(url_for('dashboard'))
    return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        user = request.form.get('username')
        password = request.form.get('password')
        usuario = Usuario.query.filter_by(username=user).first()
        if usuario and usuario.password == password:
            login_user(usuario)
            return redirect(url_for('dashboard'))
        flash('Usuario o contraseña incorrectos', 'danger')
    return render_template('login.html')

@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('login'))

@app.route('/dashboard')
@login_required
def dashboard():
    vehiculos = Vehiculo.query.all()
    choferes = Chofer.query.all()
    viajes = Viaje.query.all()
    total_fletes = sum(v.flete for v in viajes)
    return render_template('dashboard.html', vehiculos=vehiculos, choferes=choferes, viajes=viajes, total_fletes=total_fletes)

@app.route('/admin/crear-vehiculo', methods=['POST'])
@login_required
def crear_vehiculo():
    placa = request.form.get('placa')
    modelo = request.form.get('modelo')
    capacidad = request.form.get('capacidad')
    if placa and modelo:
        nuevo_v = Vehiculo(placa=placa, modelo=modelo, capacidad=float(capacidad or 0))
        db.session.add(nuevo_v)
        db.session.commit()
    return redirect(url_for('dashboard'))

@app.route('/admin/crear-chofer', methods=['POST'])
@login_required
def crear_chofer():
    nombre = request.form.get('nombre')
    cedula = request.form.get('cedula')
    telefono = request.form.get('telefono')
    if nombre and cedula and telefono:
        nuevo_c = Chofer(nombre=nombre, cedula=cedula, telefono=telefono)
        db.session.add(nuevo_c)
        db.session.commit()
    return redirect(url_for('dashboard'))

@app.route('/admin/crear-viaje', methods=['POST'])
@login_required
def crear_viaje():
    origen = request.form.get('origen')
    destino = request.form.get('destino')
    vehiculo_id = request.form.get('vehiculo_id')
    chofer_id = request.form.get('chofer_id')
    flete = request.form.get('flete')
    fecha_str = request.form.get('fecha')

    fecha_viaje = datetime.now()
    if fecha_str:
        try:
            fecha_viaje = datetime.strptime(fecha_str, '%Y-%m-%d')
        except:
            pass

    if origen and destino and vehiculo_id and chofer_id:
        nuevo_viaje = Viaje(
            origen=origen,
            destino=destino,
            vehiculo_id=int(vehiculo_id),
            chofer_id=int(chofer_id),
            flete=float(flete or 0.0),
            fecha=fecha_viaje
        )
        db.session.add(nuevo_viaje)
        db.session.commit()
    return redirect(url_for('dashboard'))

@app.route('/admin/reporte/excel', methods=['GET'])
@login_required
def reporte_excel():
    fecha_desde = request.args.get('fecha_desde')
    fecha_hasta = request.args.get('fecha_hasta')

    query = Viaje.query
    if fecha_desde:
        query = query.filter(Viaje.fecha >= datetime.strptime(fecha_desde, '%Y-%m-%d'))
    if fecha_hasta:
        dt_hasta = datetime.combine(datetime.strptime(fecha_hasta, '%Y-%m-%d').date(), time.max)
        query = query.filter(Viaje.fecha <= dt_hasta)

    viajes = query.all()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Reporte de Viajes y Fletes"

    header_fill = PatternFill(start_color="1F2937", end_color="1F2937", fill_type="solid")
    header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    border_thin = Border(left=Side(style='thin', color='CCCCCC'),
                         right=Side(style='thin', color='CCCCCC'),
                         top=Side(style='thin', color='CCCCCC'),
                         bottom=Side(style='thin', color='CCCCCC'))

    ws.append(["ID", "Fecha", "Origen", "Destino", "Chofer", "Cédula", "Placa Vehículo", "Modelo", "Flete ($)"])
    ws.row_dimensions[1].height = 25

    for col_num in range(1, 10):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for v in viajes:
        ws.append([
            v.id,
            v.fecha.strftime('%Y-%m-%d'),
            v.origen,
            v.destino,
            v.chofer.nombre if v.chofer else "N/A",
            v.chofer.cedula if v.chofer else "N/A",
            v.vehiculo.placa if v.vehiculo else "N/A",
            v.vehiculo.modelo if v.vehiculo else "N/A",
            v.flete
        ])

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=9):
        for cell in row:
            cell.border = border_thin
            cell.alignment = Alignment(vertical="center")
            if cell.column == 9:
                cell.number_format = '$#,##0.00'

    total_row = ws.max_row + 1
    ws.cell(row=total_row, column=8, value="TOTAL FLETES:").font = Font(name="Arial", size=11, bold=True)
    ws.cell(row=total_row, column=8).alignment = Alignment(horizontal="right")

    total_cell = ws.cell(row=total_row, column=9, value=f"=SUM(I2:I{total_row-1})")
    total_cell.font = Font(name="Arial", size=11, bold=True)
    total_cell.number_format = '$#,##0.00'

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
    wb.save(tmp_file.name)
    tmp_file.close()

    filename = f"Reporte_Viajes_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(tmp_file.name, as_attachment=True, download_name=filename, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/admin/reporte/pdf', methods=['GET'])
@login_required
def reporte_pdf():
    fecha_desde = request.args.get('fecha_desde')
    fecha_hasta = request.args.get('fecha_hasta')

    query = Viaje.query
    if fecha_desde:
        query = query.filter(Viaje.fecha >= datetime.strptime(fecha_desde, '%Y-%m-%d'))
    if fecha_hasta:
        dt_hasta = datetime.combine(datetime.strptime(fecha_hasta, '%Y-%m-%d').date(), time.max)
        query = query.filter(Viaje.fecha <= dt_hasta)

    viajes = query.all()
    total_fletes = sum(v.flete for v in viajes)

    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="utf-8">
    <style>
        @page {{ size: A4; margin: 15mm; background-color: #ffffff; }}
        body {{ font-family: 'Helvetica', Arial, sans-serif; color: #111827; font-size: 10pt; line-height: 1.4; }}
        .header {{ border-bottom: 2px solid #DC2626; padding-bottom: 10px; margin-bottom: 20px; }}
        h1 {{ color: #1F2937; font-size: 18pt; margin: 0 0 5px 0; }}
        .subtitle {{ color: #6B7280; font-size: 10pt; }}
        .meta {{ margin-bottom: 20px; font-size: 10pt; background: #F9FAFB; padding: 10px; border-radius: 6px; }}
        table {{ width: 100%; border-collapse: collapse; margin-top: 10px; }}
        th {{ background-color: #1F2937; color: white; text-align: left; padding: 8px; font-size: 9pt; }}
        td {{ padding: 8px; border-bottom: 1px solid #E5E7EB; font-size: 9pt; }}
        .total-box {{ margin-top: 20px; text-align: right; font-size: 12pt; font-weight: bold; color: #DC2626; }}
    </style>
    </head>
    <body>
        <div class="header">
            <h1>Reporte Operativo de Viajes y Fletes</h1>
            <div class="subtitle">Sistema de Gestión de Transporte - Maracay</div>
        </div>
        <div class="meta">
            <strong>Filtro de Fechas:</strong> {fecha_desde if fecha_desde else 'Inicio'} al {fecha_hasta if fecha_hasta else 'Actual'} <br>
            <strong>Fecha de Emisión:</strong> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} <br>
            <strong>Total de Viajes:</strong> {len(viajes)}
        </div>
        <table>
            <thead>
                <tr>
                    <th>Fecha</th>
                    <th>Ruta</th>
                    <th>Chofer</th>
                    <th>Vehículo</th>
                    <th style="text-align: right;">Flete</th>
                </tr>
            </thead>
            <tbody>
    """

    for v in viajes:
        chofer_nombre = v.chofer.nombre if v.chofer else 'N/A'
        vehiculo_placa = v.vehiculo.placa if v.vehiculo else 'N/A'
        html_content += f"""
                <tr>
                    <td>{v.fecha.strftime('%Y-%m-%d')}</td>
                    <td>{v.origen} &rarr; {v.destino}</td>
                    <td>{chofer_nombre}</td>
                    <td>{vehiculo_placa}</td>
                    <td style="text-align: right;">${v.flete:,.2f}</td>
                </tr>
        """

    html_content += f"""
            </tbody>
        </table>
        <div class="total-box">
            Total General Fletes: ${total_fletes:,.2f}
        </div>
    </body>
    </html>
    """

    tmp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
    HTML(string=html_content).write_pdf(tmp_pdf.name)
    tmp_pdf.close()

    filename = f"Reporte_Viajes_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    return send_file(tmp_pdf.name, as_attachment=True, download_name=filename, mimetype='application/pdf')

@app.route('/api/gps/traccar', methods=['GET', 'POST'])
def recibir_traccar():
    id_telefono = request.args.get('id') or request.form.get('id')
    lat = request.args.get('lat') or request.form.get('lat')
    lon = request.args.get('lon') or request.form.get('lon')

    if id_telefono and lat and lon:
        punto = RegistroGPS(telefono=str(id_telefono), latitud=float(lat), longitud=float(lon))
        db.session.add(punto)
        db.session.commit()
        return "OK", 200
    return "Bad Request", 400

@app.route('/api/gps/ultimo-punto', methods=['GET'])
@login_required
def ultimo_punto_gps():
    telefono = request.args.get('telefono', '').strip()
    punto = RegistroGPS.query.filter_by(telefono=telefono).order_by(RegistroGPS.fecha.desc()).first()
    chofer = Chofer.query.filter_by(telefono=telefono).first()

    nombre_chofer = chofer.nombre if chofer else "Chofer no registrado"

    if punto:
        return jsonify({
            'success': True,
            'chofer': nombre_chofer,
            'telefono': punto.telefono,
            'latitud': punto.latitud,
            'longitud': punto.longitud,
            'fecha': punto.fecha.strftime('%Y-%m-%d %H:%M:%S')
        })
    return jsonify({'success': False, 'message': 'No hay ubicaciones registradas para este número.'})

if __name__ == '__main__':
    with app.app_context():
        db.create_all()
    app.run(debug=True)
