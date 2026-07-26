import os
import uuid
from datetime import datetime, time, date
from functools import wraps

from flask import (Flask, render_template, request, redirect, url_for, flash,
                   jsonify, send_file, send_from_directory, abort)
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import text
from flask_login import (LoginManager, UserMixin, login_user, logout_user,
                         login_required, current_user)
from werkzeug.security import generate_password_hash, check_password_hash
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from weasyprint import HTML
import tempfile

# --- Rutas base -------------------------------------------------------------
# En PythonAnywhere los datos viven en /home/<usuario>/. En local usamos la
# carpeta del proyecto para que la app funcione en cualquier entorno.
PROD_DIR = '/home/MIGAB2026'
BASE_DIR = PROD_DIR if os.path.isdir(PROD_DIR) else os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif', 'webp', 'heic'}

app = Flask(__name__)
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'migab2026_secret_key_prod')
app.config['SQLALCHEMY_DATABASE_URI'] = f'sqlite:///{os.path.join(BASE_DIR, "database.db")}'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB por archivo
# Cookie de sesión propia y acotada a /transporte para no interferir con Solimotors
app.config['SESSION_COOKIE_NAME'] = 'sanjorge_session'
app.config['SESSION_COOKIE_PATH'] = '/transporte'

db = SQLAlchemy(app)
login_manager = LoginManager(app)
login_manager.login_view = 'login'

# Roles del sistema
ROL_MASTER = 'master'
ROL_ADMIN = 'admin'
ROL_CHOFER = 'chofer'


# --- Modelos ----------------------------------------------------------------
class Usuario(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(50), unique=True, nullable=False)
    password = db.Column(db.String(255), nullable=False)
    rol = db.Column(db.String(20), nullable=False, default=ROL_CHOFER)
    nombre = db.Column(db.String(100), nullable=False)
    # Datos propios del chofer (nulos para master/admin)
    cedula = db.Column(db.String(20), unique=True)
    telefono = db.Column(db.String(20))
    vehiculo_id = db.Column(db.Integer, db.ForeignKey('vehiculo.id'))
    activo = db.Column(db.Boolean, default=True, nullable=False)
    # Estado del rastreo GPS del chofer (por teléfono): encendido/apagado
    gps_activo = db.Column(db.Boolean)
    gps_actualizado = db.Column(db.DateTime)

    vehiculo = db.relationship('Vehiculo', backref='choferes')

    def set_password(self, raw):
        self.password = generate_password_hash(raw)

    def check_password(self, raw):
        # Compatibilidad con contraseñas antiguas guardadas en texto plano
        if self.password and self.password.startswith(('pbkdf2:', 'scrypt:', 'argon2')):
            return check_password_hash(self.password, raw)
        return self.password == raw

    @property
    def es_master(self):
        return self.rol == ROL_MASTER

    @property
    def es_admin(self):
        return self.rol == ROL_ADMIN

    @property
    def es_chofer(self):
        return self.rol == ROL_CHOFER

    @property
    def puede_administrar(self):
        return self.rol in (ROL_MASTER, ROL_ADMIN)


class Vehiculo(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    placa = db.Column(db.String(20), unique=True, nullable=False)
    modelo = db.Column(db.String(50), nullable=False)
    kilometraje_actual = db.Column(db.Float, default=0.0, nullable=False)
    estado = db.Column(db.String(20), default='Activo', nullable=False)


class Viaje(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    origen = db.Column(db.String(100), nullable=False)
    destino = db.Column(db.String(100), nullable=False)
    vehiculo_id = db.Column(db.Integer, db.ForeignKey('vehiculo.id'), nullable=False)
    chofer_id = db.Column(db.Integer, db.ForeignKey('usuario.id'), nullable=False)
    monto_flete = db.Column(db.Float, default=0.0, nullable=False)
    porcentaje_comision = db.Column(db.Float, default=0.0, nullable=False)
    monto_comision = db.Column(db.Float, default=0.0, nullable=False)
    estado = db.Column(db.String(20), default='Asignado', nullable=False)
    guia = db.Column(db.String(255))  # Foto/documento de la guía de despacho
    fecha = db.Column(db.DateTime, default=datetime.now, nullable=False)

    vehiculo = db.relationship('Vehiculo', backref='viajes')
    chofer = db.relationship('Usuario', backref='viajes')


class EventoVehiculo(db.Model):
    """Combustible / mantenimiento reportado por el chofer y confirmado por admin."""
    id = db.Column(db.Integer, primary_key=True)
    vehiculo_id = db.Column(db.Integer, db.ForeignKey('vehiculo.id'), nullable=False)
    chofer_id = db.Column(db.Integer, db.ForeignKey('usuario.id'), nullable=False)
    tipo = db.Column(db.String(40), nullable=False)  # Combustible, Aceite/Filtro, Cauchos, Reparacion, Peaje, Otros...
    viaje_id = db.Column(db.Integer, db.ForeignKey('viaje.id'))  # Gasto asociado a un viaje/despacho
    kilometraje = db.Column(db.Float, default=0.0)
    litros = db.Column(db.Float, default=0.0)  # Cantidad de combustible cargado
    monto_costo = db.Column(db.Float, default=0.0)
    moneda = db.Column(db.String(3), default='USD')  # USD (combustible/aceite) o BS (peaje)
    detalles = db.Column(db.Text)
    foto_ticket = db.Column(db.String(255))
    foto_odometro = db.Column(db.String(255))
    estado = db.Column(db.String(20), default='Pendiente', nullable=False)  # Pendiente / Confirmado / Rechazado
    fecha = db.Column(db.DateTime, default=datetime.now, nullable=False)  # Fecha de la operación
    confirmado_por_id = db.Column(db.Integer, db.ForeignKey('usuario.id'))
    fecha_confirmacion = db.Column(db.DateTime)

    vehiculo = db.relationship('Vehiculo', backref='eventos')
    chofer = db.relationship('Usuario', foreign_keys=[chofer_id])
    confirmado_por = db.relationship('Usuario', foreign_keys=[confirmado_por_id])
    viaje = db.relationship('Viaje', backref='gastos')


class RegistroGPS(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    telefono = db.Column(db.String(20), nullable=False)
    latitud = db.Column(db.Float, nullable=False)
    longitud = db.Column(db.Float, nullable=False)
    fecha = db.Column(db.DateTime, default=datetime.now)


class Auditoria(db.Model):
    """Bitácora de correcciones hechas por admin/master sobre registros de choferes."""
    id = db.Column(db.Integer, primary_key=True)
    usuario_id = db.Column(db.Integer, db.ForeignKey('usuario.id'), nullable=False)  # quién hizo el cambio
    entidad = db.Column(db.String(40), nullable=False)  # 'EventoVehiculo' o 'Viaje'
    entidad_id = db.Column(db.Integer, nullable=False)
    cambios = db.Column(db.Text, nullable=False)  # resumen campo: anterior -> nuevo
    motivo = db.Column(db.Text)
    fecha = db.Column(db.DateTime, default=datetime.now, nullable=False)

    usuario = db.relationship('Usuario')


@login_manager.user_loader
def load_user(user_id):
    return Usuario.query.get(int(user_id))


# --- Utilidades / decoradores ----------------------------------------------
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def guardar_foto(file_storage):
    if not file_storage or file_storage.filename == '':
        return None
    if not allowed_file(file_storage.filename):
        return None
    ext = file_storage.filename.rsplit('.', 1)[1].lower()
    nombre = f"{uuid.uuid4().hex}.{ext}"
    file_storage.save(os.path.join(app.config['UPLOAD_FOLDER'], nombre))
    return nombre


def roles_required(*roles):
    def decorator(fn):
        @wraps(fn)
        @login_required
        def wrapper(*args, **kwargs):
            if current_user.rol not in roles:
                abort(403)
            return fn(*args, **kwargs)
        return wrapper
    return decorator


def parse_fecha(valor, fin_dia=False):
    if not valor:
        return None
    try:
        d = datetime.strptime(valor, '%Y-%m-%d').date()
    except ValueError:
        return None
    return datetime.combine(d, time.max) if fin_dia else datetime.combine(d, time.min)


# --- Autenticación ----------------------------------------------------------
@app.route('/')
def index():
    if current_user.is_authenticated:
        return redirect(url_for('inicio'))
    return redirect(url_for('login'))


@app.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        return redirect(url_for('inicio'))
    if request.method == 'POST':
        user = request.form.get('username')
        password = request.form.get('password')
        usuario = Usuario.query.filter_by(username=user).first()
        if usuario and usuario.activo and usuario.check_password(password):
            login_user(usuario)
            return redirect(url_for('inicio'))
        flash('Usuario o contraseña incorrectos', 'danger')
    return render_template('login.html')


@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('login'))


@app.route('/inicio')
@login_required
def inicio():
    """Punto de entrada que enruta según el rol."""
    if current_user.es_chofer:
        return redirect(url_for('panel_chofer'))
    return redirect(url_for('dashboard'))


# --- Panel administrativo (master / admin) ----------------------------------
@app.route('/dashboard')
@roles_required(ROL_MASTER, ROL_ADMIN)
def dashboard():
    vehiculos = Vehiculo.query.order_by(Vehiculo.placa).all()
    choferes = Usuario.query.filter_by(rol=ROL_CHOFER).order_by(Usuario.nombre).all()
    viajes = Viaje.query.order_by(Viaje.fecha.desc()).all()
    eventos = EventoVehiculo.query.order_by(EventoVehiculo.fecha.desc()).all()
    eventos_pendientes = [e for e in eventos if e.estado == 'Pendiente']

    total_comisiones = sum(v.monto_comision for v in viajes)
    total_gastos = sum(e.monto_costo for e in eventos
                       if e.estado == 'Confirmado' and (e.moneda or 'USD') != 'BS')
    total_peaje_bs = sum(e.monto_costo for e in eventos
                         if e.estado == 'Confirmado' and (e.moneda or 'USD') == 'BS')

    return render_template(
        'dashboard.html',
        vehiculos=vehiculos,
        choferes=choferes,
        viajes=viajes,
        eventos=eventos,
        eventos_pendientes=eventos_pendientes,
        total_comisiones=total_comisiones,
        total_gastos=total_gastos,
        total_peaje_bs=total_peaje_bs,
        hoy=date.today().isoformat(),
    )


@app.route('/admin/crear-vehiculo', methods=['POST'])
@roles_required(ROL_MASTER, ROL_ADMIN)
def crear_vehiculo():
    placa = (request.form.get('placa') or '').strip().upper()
    modelo = (request.form.get('modelo') or '').strip()
    kilometraje = request.form.get('kilometraje')
    if not placa or not modelo:
        flash('Placa y modelo son obligatorios.', 'danger')
        return redirect(url_for('dashboard'))
    if Vehiculo.query.filter_by(placa=placa).first():
        flash(f'Ya existe un vehículo con placa {placa}.', 'warning')
        return redirect(url_for('dashboard'))
    nuevo_v = Vehiculo(placa=placa, modelo=modelo, kilometraje_actual=float(kilometraje or 0))
    db.session.add(nuevo_v)
    db.session.commit()
    flash(f'Vehículo {placa} registrado correctamente.', 'success')
    return redirect(url_for('dashboard'))


@app.route('/admin/crear-chofer', methods=['POST'])
@roles_required(ROL_MASTER, ROL_ADMIN)
def crear_chofer():
    nombre = (request.form.get('nombre') or '').strip()
    username = (request.form.get('username') or '').strip()
    password = request.form.get('password') or ''
    cedula = (request.form.get('cedula') or '').strip() or None
    telefono = (request.form.get('telefono') or '').strip() or None
    vehiculo_id = request.form.get('vehiculo_id')

    if not nombre or not username or not password:
        flash('Nombre, usuario y contraseña son obligatorios.', 'danger')
        return redirect(url_for('dashboard'))

    if Usuario.query.filter_by(username=username).first():
        flash(f'El usuario "{username}" ya existe.', 'warning')
        return redirect(url_for('dashboard'))
    if cedula and Usuario.query.filter_by(cedula=cedula).first():
        flash(f'Ya existe un chofer con cédula {cedula}.', 'warning')
        return redirect(url_for('dashboard'))

    veh_id = None
    if vehiculo_id and vehiculo_id not in ('none', '', '0'):
        veh_id = int(vehiculo_id)

    nuevo_c = Usuario(
        username=username,
        rol=ROL_CHOFER,
        nombre=nombre,
        cedula=cedula,
        telefono=telefono,
        vehiculo_id=veh_id,
    )
    nuevo_c.set_password(password)
    db.session.add(nuevo_c)
    db.session.commit()
    flash(f'Chofer {nombre} creado con acceso de usuario "{username}".', 'success')
    return redirect(url_for('dashboard'))


@app.route('/admin/asignar-vehiculo', methods=['POST'])
@roles_required(ROL_MASTER, ROL_ADMIN)
def asignar_vehiculo():
    chofer_id = request.form.get('chofer_id')
    vehiculo_id = request.form.get('vehiculo_id')
    chofer = Usuario.query.filter_by(id=chofer_id, rol=ROL_CHOFER).first()
    if not chofer:
        flash('Chofer no encontrado.', 'danger')
        return redirect(url_for('dashboard'))
    chofer.vehiculo_id = int(vehiculo_id) if vehiculo_id and vehiculo_id not in ('none', '', '0') else None
    db.session.commit()
    flash(f'Vehículo actualizado para {chofer.nombre}.', 'success')
    return redirect(url_for('dashboard'))


@app.route('/admin/crear-viaje', methods=['POST'])
@roles_required(ROL_MASTER, ROL_ADMIN)
def crear_viaje():
    origen = (request.form.get('origen') or '').strip()
    destino = (request.form.get('destino') or '').strip()
    vehiculo_id = request.form.get('vehiculo_id')
    chofer_id = request.form.get('chofer_id')
    monto_flete = round(float(request.form.get('monto_flete') or 0.0), 2)
    porcentaje_comision = round(float(request.form.get('porcentaje_comision') or 0.0), 2)

    if not (origen and destino and vehiculo_id and chofer_id):
        flash('Faltan datos para asignar el viaje.', 'danger')
        return redirect(url_for('dashboard'))

    monto_comision = round(monto_flete * porcentaje_comision / 100.0, 2)
    nuevo_viaje = Viaje(
        origen=origen,
        destino=destino,
        vehiculo_id=int(vehiculo_id),
        chofer_id=int(chofer_id),
        monto_flete=monto_flete,
        porcentaje_comision=porcentaje_comision,
        monto_comision=monto_comision,
        fecha=datetime.now(),
    )
    db.session.add(nuevo_viaje)
    db.session.commit()
    flash('Viaje asignado correctamente.', 'success')
    return redirect(url_for('dashboard'))


@app.route('/admin/confirmar-evento/<int:evento_id>', methods=['POST'])
@roles_required(ROL_MASTER, ROL_ADMIN)
def confirmar_evento(evento_id):
    evento = EventoVehiculo.query.get_or_404(evento_id)
    accion = request.form.get('accion', 'confirmar')
    if accion == 'rechazar':
        evento.estado = 'Rechazado'
    else:
        evento.estado = 'Confirmado'
        # Actualiza el odómetro del vehículo si el registro es mayor
        if evento.vehiculo and evento.kilometraje and evento.kilometraje > (evento.vehiculo.kilometraje_actual or 0):
            evento.vehiculo.kilometraje_actual = evento.kilometraje
    evento.confirmado_por_id = current_user.id
    evento.fecha_confirmacion = datetime.now()
    db.session.commit()
    flash(f'Registro #{evento.id} {evento.estado.lower()}.', 'success')
    return redirect(url_for('dashboard'))


# --- Corrección de registros con clave de autorización + auditoría ----------
def _registrar_auditoria(entidad, entidad_id, cambios, motivo):
    if not cambios:
        return False
    aud = Auditoria(
        usuario_id=current_user.id,
        entidad=entidad,
        entidad_id=entidad_id,
        cambios="; ".join(cambios),
        motivo=(motivo or '').strip() or None,
    )
    db.session.add(aud)
    return True


@app.route('/admin/editar-evento/<int:evento_id>', methods=['POST'])
@roles_required(ROL_MASTER, ROL_ADMIN)
def editar_evento(evento_id):
    evento = EventoVehiculo.query.get_or_404(evento_id)

    # Autorización: reingresar la clave del propio admin/master
    if not current_user.check_password(request.form.get('clave_autorizacion') or ''):
        flash('Clave de autorización incorrecta. No se guardaron los cambios.', 'danger')
        return redirect(url_for('dashboard'))

    nuevos = {
        'tipo': (request.form.get('tipo') or evento.tipo).strip(),
        'kilometraje': round(float(request.form.get('kilometraje') or 0.0), 2),
        'litros': round(float(request.form.get('litros') or 0.0), 2),
        'monto_costo': round(float(request.form.get('monto_costo') or 0.0), 2),
        'detalles': (request.form.get('detalles') or '').strip() or None,
    }
    cambios = []
    for campo, nuevo in nuevos.items():
        anterior = getattr(evento, campo)
        if (anterior or None) != (nuevo or None):
            cambios.append(f"{campo}: '{anterior}' -> '{nuevo}'")
            setattr(evento, campo, nuevo)

    if _registrar_auditoria('EventoVehiculo', evento.id, cambios, request.form.get('motivo')):
        db.session.commit()
        flash(f'Registro #{evento.id} corregido y guardado en auditoría.', 'success')
    else:
        flash('No hubo cambios que guardar.', 'info')
    return redirect(url_for('dashboard'))


@app.route('/admin/editar-viaje/<int:viaje_id>', methods=['POST'])
@roles_required(ROL_MASTER, ROL_ADMIN)
def editar_viaje(viaje_id):
    viaje = Viaje.query.get_or_404(viaje_id)

    if not current_user.check_password(request.form.get('clave_autorizacion') or ''):
        flash('Clave de autorización incorrecta. No se guardaron los cambios.', 'danger')
        return redirect(url_for('dashboard'))

    origen = (request.form.get('origen') or viaje.origen).strip()
    destino = (request.form.get('destino') or viaje.destino).strip()
    monto_flete = round(float(request.form.get('monto_flete') or 0.0), 2)
    porcentaje_comision = round(float(request.form.get('porcentaje_comision') or 0.0), 2)
    monto_comision = round(monto_flete * porcentaje_comision / 100.0, 2)

    nuevos = {
        'origen': origen,
        'destino': destino,
        'monto_flete': monto_flete,
        'porcentaje_comision': porcentaje_comision,
        'monto_comision': monto_comision,
    }
    cambios = []
    for campo, nuevo in nuevos.items():
        anterior = getattr(viaje, campo)
        if anterior != nuevo:
            cambios.append(f"{campo}: '{anterior}' -> '{nuevo}'")
            setattr(viaje, campo, nuevo)

    if _registrar_auditoria('Viaje', viaje.id, cambios, request.form.get('motivo')):
        db.session.commit()
        flash(f'Despacho #{viaje.id} corregido y guardado en auditoría.', 'success')
    else:
        flash('No hubo cambios que guardar.', 'info')
    return redirect(url_for('dashboard'))


# --- Eliminación de registros (master / admin) ------------------------------
@app.route('/admin/eliminar-evento/<int:evento_id>', methods=['POST'])
@roles_required(ROL_MASTER, ROL_ADMIN)
def eliminar_evento(evento_id):
    evento = EventoVehiculo.query.get_or_404(evento_id)
    _registrar_auditoria('EventoVehiculo', evento.id,
                         [f"eliminado (tipo '{evento.tipo}', costo {evento.monto_costo})"],
                         request.form.get('motivo') or 'Eliminación')
    db.session.delete(evento)
    db.session.commit()
    flash(f'Registro de mantenimiento #{evento_id} eliminado.', 'success')
    return redirect(url_for('dashboard'))


@app.route('/admin/eliminar-viaje/<int:viaje_id>', methods=['POST'])
@roles_required(ROL_MASTER, ROL_ADMIN)
def eliminar_viaje(viaje_id):
    viaje = Viaje.query.get_or_404(viaje_id)
    # Desvincular eventos que referencian este viaje para no dejar huérfanos
    EventoVehiculo.query.filter_by(viaje_id=viaje.id).update({'viaje_id': None})
    _registrar_auditoria('Viaje', viaje.id,
                         [f"eliminado ({viaje.origen} -> {viaje.destino}, flete {viaje.monto_flete})"],
                         request.form.get('motivo') or 'Eliminación')
    db.session.delete(viaje)
    db.session.commit()
    flash(f'Despacho #{viaje_id} eliminado.', 'success')
    return redirect(url_for('dashboard'))


@app.route('/admin/eliminar-vehiculo/<int:vehiculo_id>', methods=['POST'])
@roles_required(ROL_MASTER, ROL_ADMIN)
def eliminar_vehiculo(vehiculo_id):
    veh = Vehiculo.query.get_or_404(vehiculo_id)
    n_viajes = Viaje.query.filter_by(vehiculo_id=veh.id).count()
    n_eventos = EventoVehiculo.query.filter_by(vehiculo_id=veh.id).count()
    if n_viajes or n_eventos:
        flash(f'No se puede eliminar {veh.placa}: tiene {n_viajes} viaje(s) y '
              f'{n_eventos} registro(s) asociados. Elimínalos primero.', 'danger')
        return redirect(url_for('dashboard'))
    # Desasignar choferes que tuvieran este vehículo
    Usuario.query.filter_by(vehiculo_id=veh.id).update({'vehiculo_id': None})
    _registrar_auditoria('Vehiculo', veh.id,
                         [f"eliminado (placa '{veh.placa}', modelo '{veh.modelo}')"],
                         request.form.get('motivo') or 'Eliminación')
    db.session.delete(veh)
    db.session.commit()
    flash(f'Vehículo {veh.placa} eliminado.', 'success')
    return redirect(url_for('dashboard'))


@app.route('/admin/eliminar-chofer/<int:chofer_id>', methods=['POST'])
@roles_required(ROL_MASTER, ROL_ADMIN)
def eliminar_chofer(chofer_id):
    chofer = Usuario.query.filter_by(id=chofer_id, rol=ROL_CHOFER).first_or_404()
    n_viajes = Viaje.query.filter_by(chofer_id=chofer.id).count()
    n_eventos = EventoVehiculo.query.filter_by(chofer_id=chofer.id).count()
    if n_viajes or n_eventos:
        flash(f'No se puede eliminar a {chofer.nombre}: tiene {n_viajes} viaje(s) y '
              f'{n_eventos} registro(s) asociados. Elimínalos primero.', 'danger')
        return redirect(url_for('dashboard'))
    _registrar_auditoria('Usuario', chofer.id,
                         [f"chofer eliminado ('{chofer.nombre}', usuario '{chofer.username}')"],
                         request.form.get('motivo') or 'Eliminación')
    db.session.delete(chofer)
    db.session.commit()
    flash(f'Chofer {chofer.nombre} eliminado.', 'success')
    return redirect(url_for('dashboard'))


@app.route('/admin/auditoria')
@roles_required(ROL_MASTER, ROL_ADMIN)
def ver_auditoria():
    registros = Auditoria.query.order_by(Auditoria.fecha.desc()).limit(200).all()
    return render_template('auditoria.html', registros=registros)


# --- Ayuda / explicación del sistema ----------------------------------------
@app.route('/ayuda')
@login_required
def ayuda():
    return render_template('ayuda.html')


# --- Panel del chofer -------------------------------------------------------
@app.route('/chofer')
@roles_required(ROL_CHOFER)
def panel_chofer():
    vehiculo = current_user.vehiculo
    eventos = (EventoVehiculo.query
               .filter_by(chofer_id=current_user.id)
               .order_by(EventoVehiculo.fecha.desc())
               .limit(20).all())
    viajes = (Viaje.query
              .filter_by(chofer_id=current_user.id)
              .order_by(Viaje.fecha.desc())
              .limit(20).all())
    return render_template('chofer.html', vehiculo=vehiculo, eventos=eventos,
                           viajes=viajes, hoy=date.today().isoformat())


@app.route('/chofer/registrar-evento', methods=['POST'])
@roles_required(ROL_CHOFER)
def registrar_evento():
    if not current_user.vehiculo_id:
        flash('No tienes un vehículo asignado.', 'danger')
        return redirect(url_for('panel_chofer'))

    # El chofer registra en $ (Combustible, Cauchos, Repuestos, Otros) o Peaje en Bs.
    # El cambio de aceite y filtro queda excluido: solo lo hace Luis.
    tipo = request.form.get('tipo') or 'Combustible'
    if tipo not in ('Combustible', 'Cauchos', 'Repuestos', 'Otros', 'Peaje'):
        flash('Tipo de registro no permitido para el chofer.', 'danger')
        return redirect(url_for('panel_chofer'))

    kilometraje = round(float(request.form.get('kilometraje') or 0.0), 2)
    litros = round(float(request.form.get('litros') or 0.0), 2)
    monto_costo = round(float(request.form.get('monto_costo') or 0.0), 2)
    detalles = request.form.get('detalles')
    moneda = 'BS' if tipo == 'Peaje' else 'USD'

    fecha_op = parse_fecha(request.form.get('fecha')) or datetime.now()

    viaje_id = request.form.get('viaje_id')
    veh_viaje = None
    if viaje_id and viaje_id not in ('none', '', '0'):
        # Solo puede asociar sus propios viajes
        v = Viaje.query.filter_by(id=int(viaje_id), chofer_id=current_user.id).first()
        veh_viaje = v.id if v else None

    foto_ticket = guardar_foto(request.files.get('foto_ticket'))
    foto_odometro = guardar_foto(request.files.get('foto_odometro'))

    # El peaje exige foto del recibo
    if tipo == 'Peaje' and not foto_ticket:
        flash('Para registrar un peaje debes anexar la foto del recibo de peaje.', 'danger')
        return redirect(url_for('panel_chofer'))

    evento = EventoVehiculo(
        vehiculo_id=current_user.vehiculo_id,
        chofer_id=current_user.id,
        tipo=tipo,
        viaje_id=veh_viaje,
        kilometraje=kilometraje,
        litros=litros,
        monto_costo=monto_costo,
        moneda=moneda,
        detalles=detalles,
        foto_ticket=foto_ticket,
        foto_odometro=foto_odometro,
        estado='Pendiente',
        fecha=fecha_op,
    )
    db.session.add(evento)
    db.session.commit()
    flash('Registro enviado. Queda pendiente de confirmación por el administrador.', 'success')
    return redirect(url_for('panel_chofer'))


# Gastos que puede registrar Luis (rol admin) directamente, en $.
GASTOS_ADMIN = ('Aceite/Filtro', 'Aceites', 'Cauchos', 'Repuestos',
                'Reparacion', 'Taller', 'Combustible', 'Otros')


@app.route('/admin/registrar-aceite-filtro', methods=['POST'])
@roles_required(ROL_ADMIN)
def registrar_aceite_filtro():
    """Registro de gastos por Luis (rol admin), en $.

    Incluye cambio de aceite y filtro, aceites y otros gastos de mantenimiento.
    """
    vehiculo_id = request.form.get('vehiculo_id')
    if not vehiculo_id or vehiculo_id in ('none', '', '0'):
        flash('Selecciona el vehículo del gasto.', 'danger')
        return redirect(url_for('dashboard'))
    veh = Vehiculo.query.get(int(vehiculo_id))
    if not veh:
        flash('Vehículo no encontrado.', 'danger')
        return redirect(url_for('dashboard'))

    tipo = request.form.get('tipo') or 'Aceite/Filtro'
    if tipo not in GASTOS_ADMIN:
        flash('Tipo de gasto no permitido.', 'danger')
        return redirect(url_for('dashboard'))

    kilometraje = round(float(request.form.get('kilometraje') or 0.0), 2)
    monto_costo = round(float(request.form.get('monto_costo') or 0.0), 2)
    detalles = request.form.get('detalles')
    fecha_op = parse_fecha(request.form.get('fecha')) or datetime.now()
    foto_ticket = guardar_foto(request.files.get('foto_ticket'))

    # chofer asignado al vehículo (si lo hay); si no, se atribuye a quien lo registra
    chofer = veh.choferes[0] if veh.choferes else current_user
    if kilometraje and kilometraje > (veh.kilometraje_actual or 0):
        veh.kilometraje_actual = kilometraje

    evento = EventoVehiculo(
        vehiculo_id=veh.id,
        chofer_id=chofer.id,
        tipo=tipo,
        kilometraje=kilometraje,
        monto_costo=monto_costo,
        moneda='USD',
        detalles=detalles,
        foto_ticket=foto_ticket,
        estado='Confirmado',
        fecha=fecha_op,
        confirmado_por_id=current_user.id,
        fecha_confirmacion=datetime.now(),
    )
    db.session.add(evento)
    db.session.commit()
    flash(f'Gasto ({tipo}) registrado para {veh.placa}.', 'success')
    return redirect(url_for('dashboard'))


@app.route('/chofer/crear-viaje', methods=['POST'])
@roles_required(ROL_CHOFER)
def chofer_crear_viaje():
    if not current_user.vehiculo_id:
        flash('No tienes un vehículo asignado.', 'danger')
        return redirect(url_for('panel_chofer'))

    origen = (request.form.get('origen') or '').strip()
    destino = (request.form.get('destino') or '').strip()
    monto_flete = round(float(request.form.get('monto_flete') or 0.0), 2)
    porcentaje_comision = round(float(request.form.get('porcentaje_comision') or 0.0), 2)

    if not (origen and destino):
        flash('Indica origen y destino del despacho.', 'danger')
        return redirect(url_for('panel_chofer'))

    guia = guardar_foto(request.files.get('guia'))
    monto_comision = round(monto_flete * porcentaje_comision / 100.0, 2)
    viaje = Viaje(
        origen=origen,
        destino=destino,
        vehiculo_id=current_user.vehiculo_id,
        chofer_id=current_user.id,
        monto_flete=monto_flete,
        porcentaje_comision=porcentaje_comision,
        monto_comision=monto_comision,
        guia=guia,
        estado='Reportado',
        fecha=datetime.now(),
    )
    db.session.add(viaje)
    db.session.commit()
    flash(f'Despacho registrado. Tu comisión es ${monto_comision:,.2f}.', 'success')
    return redirect(url_for('panel_chofer'))


@app.route('/uploads/<path:filename>')
@login_required
def uploads(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)


# --- GPS --------------------------------------------------------------------
@app.route('/api/actualizar-gps', methods=['POST'])
@login_required
def actualizar_gps():
    """El chofer transmite su ubicación desde el teléfono."""
    data = request.get_json(silent=True) or {}
    lat = data.get('latitud')
    lon = data.get('longitud')
    telefono = current_user.telefono
    if not telefono:
        return jsonify({'success': False, 'message': 'El chofer no tiene teléfono registrado.'}), 400
    if lat is None or lon is None:
        return jsonify({'success': False, 'message': 'Coordenadas incompletas.'}), 400
    punto = RegistroGPS(telefono=str(telefono), latitud=float(lat), longitud=float(lon))
    db.session.add(punto)
    # El GPS está encendido: la ubicación llegó correctamente
    current_user.gps_activo = True
    current_user.gps_actualizado = datetime.now()
    db.session.commit()
    return jsonify({'success': True})


@app.route('/api/gps/estado', methods=['POST'])
@login_required
def gps_estado():
    """El teléfono del chofer reporta si el rastreo GPS está encendido o apagado."""
    data = request.get_json(silent=True) or {}
    current_user.gps_activo = bool(data.get('activo'))
    current_user.gps_actualizado = datetime.now()
    db.session.commit()
    return jsonify({'success': True})


@app.route('/api/gps/traccar', methods=['GET', 'POST'])
def recibir_traccar():
    """Compatibilidad con apps tipo Traccar que envían por número/ID."""
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
@roles_required(ROL_MASTER, ROL_ADMIN)
def ultimo_punto_gps():
    telefono = request.args.get('telefono', '').strip()
    punto = RegistroGPS.query.filter_by(telefono=telefono).order_by(RegistroGPS.fecha.desc()).first()
    chofer = Usuario.query.filter_by(telefono=telefono, rol=ROL_CHOFER).first()
    nombre_chofer = chofer.nombre if chofer else "Chofer no registrado"
    if punto:
        return jsonify({
            'success': True,
            'chofer': nombre_chofer,
            'telefono': punto.telefono,
            'latitud': punto.latitud,
            'longitud': punto.longitud,
            'fecha': punto.fecha.strftime('%Y-%m-%d %H:%M:%S'),
        })
    return jsonify({'success': False, 'message': 'No hay ubicaciones registradas para este número.'})


@app.route('/api/gps/flota', methods=['GET'])
@roles_required(ROL_MASTER, ROL_ADMIN)
def gps_flota():
    """Última posición conocida de todos los choferes con teléfono."""
    resultado = []
    choferes = Usuario.query.filter(Usuario.rol == ROL_CHOFER, Usuario.telefono.isnot(None)).all()
    for c in choferes:
        punto = (RegistroGPS.query.filter_by(telefono=str(c.telefono))
                 .order_by(RegistroGPS.fecha.desc()).first())
        if punto:
            resultado.append({
                'chofer': c.nombre,
                'telefono': c.telefono,
                'vehiculo': c.vehiculo.placa if c.vehiculo else 'Sin asignar',
                'latitud': punto.latitud,
                'longitud': punto.longitud,
                'fecha': punto.fecha.strftime('%Y-%m-%d %H:%M:%S'),
            })
    return jsonify({'success': True, 'puntos': resultado})


@app.route('/api/gps/estados', methods=['GET'])
@roles_required(ROL_MASTER, ROL_ADMIN)
def gps_estados():
    """Estado de rastreo (encendido/apagado) de cada chofer, para alertar a master/admin."""
    choferes = Usuario.query.filter_by(rol=ROL_CHOFER).order_by(Usuario.nombre).all()
    data = []
    for c in choferes:
        if c.gps_activo is None:
            estado = 'sin_datos'
        elif c.gps_activo:
            estado = 'encendido'
        else:
            estado = 'apagado'
        data.append({
            'chofer': c.nombre,
            'telefono': c.telefono or '—',
            'vehiculo': c.vehiculo.placa if c.vehiculo else 'Sin asignar',
            'estado': estado,
            'actualizado': c.gps_actualizado.strftime('%Y-%m-%d %H:%M:%S') if c.gps_actualizado else None,
        })
    apagados = [d for d in data if d['estado'] != 'encendido']
    return jsonify({'success': True, 'choferes': data, 'apagados': apagados})


# --- Reportes: Viajes -------------------------------------------------------
def _filtrar_viajes(fecha_desde, fecha_hasta):
    query = Viaje.query
    d = parse_fecha(fecha_desde)
    h = parse_fecha(fecha_hasta, fin_dia=True)
    if d:
        query = query.filter(Viaje.fecha >= d)
    if h:
        query = query.filter(Viaje.fecha <= h)
    return query.order_by(Viaje.fecha).all()


@app.route('/admin/reporte/excel', methods=['GET'])
@roles_required(ROL_MASTER, ROL_ADMIN)
def reporte_excel():
    fecha_desde = request.args.get('fecha_desde')
    fecha_hasta = request.args.get('fecha_hasta')
    viajes = _filtrar_viajes(fecha_desde, fecha_hasta)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Viajes y Comisiones"

    header_fill = PatternFill(start_color="1F2937", end_color="1F2937", fill_type="solid")
    header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    border_thin = Border(left=Side(style='thin', color='CCCCCC'),
                         right=Side(style='thin', color='CCCCCC'),
                         top=Side(style='thin', color='CCCCCC'),
                         bottom=Side(style='thin', color='CCCCCC'))

    headers = ["ID", "Fecha", "Origen", "Destino", "Chofer", "Placa", "Modelo",
               "Flete ($)", "% Com.", "Comisión ($)", "Estado", "Guía"]
    ws.append(headers)
    ws.row_dimensions[1].height = 25
    for col_num in range(1, len(headers) + 1):
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
            v.vehiculo.placa if v.vehiculo else "N/A",
            v.vehiculo.modelo if v.vehiculo else "N/A",
            v.monto_flete,
            v.porcentaje_comision,
            v.monto_comision,
            v.estado,
            "Sí" if v.guia else "No",
        ])

    ncols = len(headers)
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ncols):
        for cell in row:
            cell.border = border_thin
            cell.alignment = Alignment(vertical="center")
            if cell.column in (8, 10):
                cell.number_format = '$#,##0.00'

    total_row = ws.max_row + 1
    ws.cell(row=total_row, column=7, value="TOTALES:").font = Font(name="Arial", size=11, bold=True)
    ws.cell(row=total_row, column=7).alignment = Alignment(horizontal="right")
    for col in (8, 10):
        letter = get_column_letter(col)
        c = ws.cell(row=total_row, column=col, value=f"=SUM({letter}2:{letter}{total_row-1})")
        c.font = Font(name="Arial", size=11, bold=True)
        c.number_format = '$#,##0.00'

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
    wb.save(tmp_file.name)
    tmp_file.close()
    filename = f"Reporte_Viajes_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(tmp_file.name, as_attachment=True, download_name=filename,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/admin/reporte/pdf', methods=['GET'])
@roles_required(ROL_MASTER, ROL_ADMIN)
def reporte_pdf():
    fecha_desde = request.args.get('fecha_desde')
    fecha_hasta = request.args.get('fecha_hasta')
    viajes = _filtrar_viajes(fecha_desde, fecha_hasta)
    total_fletes = sum(v.monto_flete for v in viajes)
    total_comisiones = sum(v.monto_comision for v in viajes)

    filas = ""
    for v in viajes:
        chofer_nombre = v.chofer.nombre if v.chofer else 'N/A'
        vehiculo_placa = v.vehiculo.placa if v.vehiculo else 'N/A'
        filas += f"""
            <tr>
                <td>{v.fecha.strftime('%Y-%m-%d')}</td>
                <td>{v.origen} &rarr; {v.destino}</td>
                <td>{chofer_nombre}</td>
                <td>{vehiculo_placa}</td>
                <td style="text-align: right;">${v.monto_flete:,.2f}</td>
                <td style="text-align: right;">${v.monto_comision:,.2f}</td>
                <td>{v.estado}</td>
                <td>{'Sí' if v.guia else 'No'}</td>
            </tr>
        """

    html_content = f"""
    <!DOCTYPE html><html><head><meta charset="utf-8">
    <style>
        @page {{ size: A4; margin: 15mm; }}
        body {{ font-family: 'Helvetica', Arial, sans-serif; color: #111827; font-size: 10pt; }}
        .header {{ border-bottom: 2px solid #DC2626; padding-bottom: 10px; margin-bottom: 20px; }}
        h1 {{ color: #1F2937; font-size: 18pt; margin: 0 0 5px 0; }}
        .subtitle {{ color: #6B7280; font-size: 10pt; }}
        .meta {{ margin-bottom: 20px; font-size: 10pt; background: #F9FAFB; padding: 10px; border-radius: 6px; }}
        table {{ width: 100%; border-collapse: collapse; margin-top: 10px; }}
        th {{ background-color: #1F2937; color: white; text-align: left; padding: 8px; font-size: 9pt; }}
        td {{ padding: 8px; border-bottom: 1px solid #E5E7EB; font-size: 9pt; }}
        .total-box {{ margin-top: 20px; text-align: right; font-size: 12pt; font-weight: bold; color: #DC2626; }}
    </style></head><body>
        <div class="header">
            <h1>Reporte de Viajes y Comisiones</h1>
            <div class="subtitle">San Jorge, C.A.</div>
        </div>
        <div class="meta">
            <strong>Filtro de Fechas:</strong> {fecha_desde or 'Inicio'} al {fecha_hasta or 'Actual'} <br>
            <strong>Fecha de Emisión:</strong> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} <br>
            <strong>Total de Viajes:</strong> {len(viajes)}
        </div>
        <table>
            <thead><tr>
                <th>Fecha</th><th>Ruta</th><th>Chofer</th><th>Vehículo</th>
                <th style="text-align:right;">Flete</th><th style="text-align:right;">Comisión</th>
                <th>Estado</th><th>Guía</th>
            </tr></thead>
            <tbody>{filas}</tbody>
        </table>
        <div class="total-box">
            Total Fletes: ${total_fletes:,.2f} &nbsp;|&nbsp; Total Comisiones: ${total_comisiones:,.2f}
        </div>
    </body></html>
    """

    tmp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
    HTML(string=html_content).write_pdf(tmp_pdf.name)
    tmp_pdf.close()
    filename = f"Reporte_Viajes_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    return send_file(tmp_pdf.name, as_attachment=True, download_name=filename, mimetype='application/pdf')


# --- Reportes: Mantenimiento / Combustible ----------------------------------
def _filtrar_eventos(fecha_desde, fecha_hasta, estado=None):
    query = EventoVehiculo.query
    d = parse_fecha(fecha_desde)
    h = parse_fecha(fecha_hasta, fin_dia=True)
    if d:
        query = query.filter(EventoVehiculo.fecha >= d)
    if h:
        query = query.filter(EventoVehiculo.fecha <= h)
    if estado:
        query = query.filter(EventoVehiculo.estado == estado)
    return query.order_by(EventoVehiculo.fecha).all()


@app.route('/admin/reporte/mantenimiento/excel', methods=['GET'])
@roles_required(ROL_MASTER, ROL_ADMIN)
def reporte_mantenimiento_excel():
    fecha_desde = request.args.get('fecha_desde')
    fecha_hasta = request.args.get('fecha_hasta')
    estado = request.args.get('estado') or None
    eventos = _filtrar_eventos(fecha_desde, fecha_hasta, estado)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Mantenimiento y Combustible"

    header_fill = PatternFill(start_color="1F2937", end_color="1F2937", fill_type="solid")
    header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    border_thin = Border(left=Side(style='thin', color='CCCCCC'),
                         right=Side(style='thin', color='CCCCCC'),
                         top=Side(style='thin', color='CCCCCC'),
                         bottom=Side(style='thin', color='CCCCCC'))

    headers = ["ID", "Fecha", "Tipo", "Vehículo", "Chofer", "Kilometraje",
               "Litros", "Costo", "Moneda", "Estado", "Detalles"]
    ws.append(headers)
    ws.row_dimensions[1].height = 25
    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for e in eventos:
        ws.append([
            e.id,
            e.fecha.strftime('%Y-%m-%d'),
            e.tipo,
            e.vehiculo.placa if e.vehiculo else "N/A",
            e.chofer.nombre if e.chofer else "N/A",
            round(e.kilometraje or 0, 2),
            round(e.litros or 0, 2),
            round(e.monto_costo or 0, 2),
            'Bs' if (e.moneda or 'USD') == 'BS' else '$',
            e.estado,
            e.detalles or "",
        ])

    ncols = len(headers)
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ncols):
        for cell in row:
            cell.border = border_thin
            cell.alignment = Alignment(vertical="center")
            if cell.column in (7, 8):
                cell.number_format = '#,##0.00'

    total_usd = sum((e.monto_costo or 0) for e in eventos if (e.moneda or 'USD') != 'BS')
    total_bs = sum((e.monto_costo or 0) for e in eventos if (e.moneda or 'USD') == 'BS')
    total_row = ws.max_row + 1
    ws.cell(row=total_row, column=7, value="TOTAL $ (combustible/aceite):").font = Font(name="Arial", size=11, bold=True)
    ws.cell(row=total_row, column=7).alignment = Alignment(horizontal="right")
    c = ws.cell(row=total_row, column=8, value=round(total_usd, 2))
    c.font = Font(name="Arial", size=11, bold=True)
    c.number_format = '$#,##0.00'
    ws.cell(row=total_row + 1, column=7, value="TOTAL Bs (peaje):").font = Font(name="Arial", size=11, bold=True)
    ws.cell(row=total_row + 1, column=7).alignment = Alignment(horizontal="right")
    cb = ws.cell(row=total_row + 1, column=8, value=round(total_bs, 2))
    cb.font = Font(name="Arial", size=11, bold=True)
    cb.number_format = '"Bs" #,##0.00'

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
    wb.save(tmp_file.name)
    tmp_file.close()
    filename = f"Reporte_Mantenimiento_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(tmp_file.name, as_attachment=True, download_name=filename,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/admin/reporte/mantenimiento/pdf', methods=['GET'])
@roles_required(ROL_MASTER, ROL_ADMIN)
def reporte_mantenimiento_pdf():
    fecha_desde = request.args.get('fecha_desde')
    fecha_hasta = request.args.get('fecha_hasta')
    estado = request.args.get('estado') or None
    eventos = _filtrar_eventos(fecha_desde, fecha_hasta, estado)
    total_usd = sum(e.monto_costo for e in eventos if (e.moneda or 'USD') != 'BS')
    total_bs = sum(e.monto_costo for e in eventos if (e.moneda or 'USD') == 'BS')

    filas = ""
    for e in eventos:
        es_bs = (e.moneda or 'USD') == 'BS'
        costo_fmt = (f"Bs {e.monto_costo:,.2f}" if es_bs else f"${e.monto_costo:,.2f}")
        filas += f"""
            <tr>
                <td>{e.fecha.strftime('%Y-%m-%d')}</td>
                <td>{e.tipo}</td>
                <td>{e.vehiculo.placa if e.vehiculo else 'N/A'}</td>
                <td>{e.chofer.nombre if e.chofer else 'N/A'}</td>
                <td style="text-align:right;">{e.kilometraje:,.0f}</td>
                <td style="text-align:right;">{(e.litros or 0):,.2f}</td>
                <td style="text-align:right;">{costo_fmt}</td>
                <td>{e.estado}</td>
            </tr>
        """

    html_content = f"""
    <!DOCTYPE html><html><head><meta charset="utf-8">
    <style>
        @page {{ size: A4 landscape; margin: 15mm; }}
        body {{ font-family: 'Helvetica', Arial, sans-serif; color: #111827; font-size: 10pt; }}
        .header {{ border-bottom: 2px solid #F59E0B; padding-bottom: 10px; margin-bottom: 20px; }}
        h1 {{ color: #1F2937; font-size: 18pt; margin: 0 0 5px 0; }}
        .subtitle {{ color: #6B7280; font-size: 10pt; }}
        .meta {{ margin-bottom: 20px; font-size: 10pt; background: #F9FAFB; padding: 10px; border-radius: 6px; }}
        table {{ width: 100%; border-collapse: collapse; margin-top: 10px; }}
        th {{ background-color: #1F2937; color: white; text-align: left; padding: 8px; font-size: 9pt; }}
        td {{ padding: 8px; border-bottom: 1px solid #E5E7EB; font-size: 9pt; }}
        .total-box {{ margin-top: 20px; text-align: right; font-size: 12pt; font-weight: bold; color: #B45309; }}
    </style></head><body>
        <div class="header">
            <h1>Reporte de Mantenimiento y Combustible</h1>
            <div class="subtitle">San Jorge, C.A.</div>
        </div>
        <div class="meta">
            <strong>Filtro de Fechas:</strong> {fecha_desde or 'Inicio'} al {fecha_hasta or 'Actual'} <br>
            <strong>Estado:</strong> {estado or 'Todos'} <br>
            <strong>Fecha de Emisión:</strong> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} <br>
            <strong>Total de Registros:</strong> {len(eventos)}
        </div>
        <table>
            <thead><tr>
                <th>Fecha</th><th>Tipo</th><th>Vehículo</th><th>Chofer</th>
                <th style="text-align:right;">KM</th><th style="text-align:right;">Litros</th><th style="text-align:right;">Costo</th><th>Estado</th>
            </tr></thead>
            <tbody>{filas}</tbody>
        </table>
        <div class="total-box">Total $ (combustible/aceite): ${total_usd:,.2f} &nbsp;|&nbsp; Total Bs (peaje): Bs {total_bs:,.2f}</div>
    </body></html>
    """

    tmp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
    HTML(string=html_content).write_pdf(tmp_pdf.name)
    tmp_pdf.close()
    filename = f"Reporte_Mantenimiento_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    return send_file(tmp_pdf.name, as_attachment=True, download_name=filename, mimetype='application/pdf')


# --- Reportes: Financiero por viaje (ingresos / gastos / ganancias) ---------
def _finanzas_viajes(fecha_desde, fecha_hasta, vehiculo_id=None, chofer_id=None):
    """Devuelve lista de dicts con ingreso, gastos, comisión y ganancia por viaje.

    Ganancia neta = flete - comisión del chofer - gastos asociados confirmados.
    """
    query = Viaje.query
    d = parse_fecha(fecha_desde)
    h = parse_fecha(fecha_hasta, fin_dia=True)
    if d:
        query = query.filter(Viaje.fecha >= d)
    if h:
        query = query.filter(Viaje.fecha <= h)
    if vehiculo_id and str(vehiculo_id) not in ('none', '', '0'):
        query = query.filter(Viaje.vehiculo_id == int(vehiculo_id))
    if chofer_id:
        query = query.filter(Viaje.chofer_id == int(chofer_id))
    viajes = query.order_by(Viaje.fecha).all()

    filas = []
    for v in viajes:
        activos = [g for g in v.gastos if g.estado != 'Rechazado']
        gastos = sum((g.monto_costo or 0.0) for g in activos if (g.moneda or 'USD') != 'BS')
        peaje_bs = sum((g.monto_costo or 0.0) for g in activos if (g.moneda or 'USD') == 'BS')
        ingreso = v.monto_flete or 0.0
        comision = v.monto_comision or 0.0
        ganancia = round(ingreso - comision - gastos, 2)
        filas.append({
            'viaje': v,
            'fecha': v.fecha,
            'ruta': f"{v.origen} → {v.destino}",
            'vehiculo': v.vehiculo.placa if v.vehiculo else 'N/A',
            'chofer': v.chofer.nombre if v.chofer else 'N/A',
            'ingreso': round(ingreso, 2),
            'comision': round(comision, 2),
            'gastos': round(gastos, 2),
            'peaje_bs': round(peaje_bs, 2),
            'ganancia': ganancia,
        })
    return filas


def _finanzas_scope():
    """Filtros según rol: el chofer solo ve sus propios viajes."""
    fd = request.args.get('fecha_desde')
    fh = request.args.get('fecha_hasta')
    veh = request.args.get('vehiculo_id')
    chofer_id = current_user.id if current_user.es_chofer else None
    return _finanzas_viajes(fd, fh, veh, chofer_id), fd, fh


@app.route('/reporte/financiero/excel', methods=['GET'])
@login_required
def reporte_financiero_excel():
    filas, fd, fh = _finanzas_scope()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Ingresos, Gastos y Ganancias"

    header_fill = PatternFill(start_color="065F46", end_color="065F46", fill_type="solid")
    header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")

    headers = ["ID", "Fecha", "Ruta", "Vehículo", "Chofer",
               "Ingreso Flete ($)", "Comisión Chofer ($)", "Gastos ($)", "Ganancia Neta ($)"]
    ws.append(headers)
    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for f in filas:
        ws.append([
            f['viaje'].id,
            f['fecha'].strftime('%Y-%m-%d'),
            f['ruta'],
            f['vehiculo'],
            f['chofer'],
            f['ingreso'],
            f['comision'],
            f['gastos'],
            f['ganancia'],
        ])

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=6, max_col=9):
        for cell in row:
            cell.number_format = '$#,##0.00'

    total_row = ws.max_row + 1
    ws.cell(row=total_row, column=5, value="TOTALES:").font = Font(bold=True)
    for col in (6, 7, 8, 9):
        letter = get_column_letter(col)
        c = ws.cell(row=total_row, column=col, value=f"=SUM({letter}2:{letter}{total_row-1})")
        c.font = Font(bold=True)
        c.number_format = '$#,##0.00'

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = max(max_len + 4, 12)

    tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
    wb.save(tmp_file.name)
    tmp_file.close()
    filename = f"Reporte_Financiero_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(tmp_file.name, as_attachment=True, download_name=filename,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/reporte/financiero/pdf', methods=['GET'])
@login_required
def reporte_financiero_pdf():
    filas, fd, fh = _finanzas_scope()
    tot_ing = sum(f['ingreso'] for f in filas)
    tot_com = sum(f['comision'] for f in filas)
    tot_gas = sum(f['gastos'] for f in filas)
    tot_gan = sum(f['ganancia'] for f in filas)

    cuerpo = ""
    for f in filas:
        cuerpo += f"""
            <tr>
                <td>{f['fecha'].strftime('%Y-%m-%d')}</td>
                <td>{f['ruta']}</td>
                <td>{f['vehiculo']}</td>
                <td>{f['chofer']}</td>
                <td style="text-align:right;">${f['ingreso']:,.2f}</td>
                <td style="text-align:right;">${f['comision']:,.2f}</td>
                <td style="text-align:right;">${f['gastos']:,.2f}</td>
                <td style="text-align:right;">${f['ganancia']:,.2f}</td>
            </tr>
        """

    html_content = f"""
    <!DOCTYPE html><html><head><meta charset="utf-8">
    <style>
        @page {{ size: A4 landscape; margin: 15mm; }}
        body {{ font-family: Arial, sans-serif; color: #111827; font-size: 10pt; }}
        .header {{ border-bottom: 2px solid #059669; padding-bottom: 10px; margin-bottom: 20px; }}
        h1 {{ color: #065F46; font-size: 18pt; margin: 0 0 5px 0; }}
        .meta {{ margin-bottom: 15px; font-size: 10pt; background: #F0FDF4; padding: 10px; border-radius: 6px; }}
        table {{ width: 100%; border-collapse: collapse; }}
        th {{ background-color: #065F46; color: white; text-align: left; padding: 7px; font-size: 9pt; }}
        td {{ padding: 7px; border-bottom: 1px solid #E5E7EB; font-size: 9pt; }}
        tfoot td {{ font-weight: bold; background:#F0FDF4; }}
    </style></head><body>
        <div class="header"><h1>Informe de Ingresos, Gastos y Ganancias</h1>
        <div>San Jorge, C.A.</div></div>
        <div class="meta">
            <strong>Filtro de Fechas:</strong> {fd or 'Inicio'} al {fh or 'Actual'} <br>
            <strong>Emitido por:</strong> {current_user.nombre} ({current_user.rol}) <br>
            <strong>Fecha de Emisión:</strong> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} <br>
            <strong>Viajes:</strong> {len(filas)}
        </div>
        <table>
            <thead><tr>
                <th>Fecha</th><th>Ruta</th><th>Vehículo</th><th>Chofer</th>
                <th style="text-align:right;">Ingreso</th><th style="text-align:right;">Comisión</th>
                <th style="text-align:right;">Gastos</th><th style="text-align:right;">Ganancia Neta</th>
            </tr></thead>
            <tbody>{cuerpo}</tbody>
            <tfoot><tr>
                <td colspan="4" style="text-align:right;">TOTALES:</td>
                <td style="text-align:right;">${tot_ing:,.2f}</td>
                <td style="text-align:right;">${tot_com:,.2f}</td>
                <td style="text-align:right;">${tot_gas:,.2f}</td>
                <td style="text-align:right;">${tot_gan:,.2f}</td>
            </tr></tfoot>
        </table>
    </body></html>
    """

    tmp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
    HTML(string=html_content).write_pdf(tmp_pdf.name)
    tmp_pdf.close()
    filename = f"Reporte_Financiero_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    return send_file(tmp_pdf.name, as_attachment=True, download_name=filename, mimetype='application/pdf')


# --- Inicialización ---------------------------------------------------------
def seed_usuarios():
    """Crea los usuarios principales de San Jorge, C.A. si no existen.

    - Jorge Soares: master (usuario principal), username 'jorge'.
    - Luis Camacho: administrador, username 'luis'.
    Las contraseñas iniciales se toman de las variables de entorno
    JORGE_PASSWORD / LUIS_PASSWORD (con un valor por defecto que debe
    cambiarse tras el primer inicio de sesión).
    """
    if not Usuario.query.filter_by(username='jorge').first():
        jorge = Usuario(username='jorge', rol=ROL_MASTER, nombre='Jorge Soares')
        jorge.set_password(os.environ.get('JORGE_PASSWORD', 'jorge2026'))
        db.session.add(jorge)
    if not Usuario.query.filter_by(username='luis').first():
        luis = Usuario(username='luis', rol=ROL_ADMIN, nombre='Luis Camacho')
        luis.set_password(os.environ.get('LUIS_PASSWORD', 'luis2026'))
        db.session.add(luis)
    db.session.commit()


def migrar_esquema():
    """Agrega columnas nuevas a tablas existentes (SQLite) si aún no existen.

    Idempotente: se puede ejecutar en cada arranque sin perder datos.
    """
    columnas = {
        'evento_vehiculo': [("moneda", "VARCHAR(3) DEFAULT 'USD'")],
        'usuario': [("gps_activo", "BOOLEAN"), ("gps_actualizado", "DATETIME")],
    }
    with db.engine.connect() as conn:
        for tabla, cols in columnas.items():
            existentes = {row[1] for row in conn.execute(text(f"PRAGMA table_info({tabla})"))}
            for nombre, definicion in cols:
                if nombre not in existentes:
                    conn.execute(text(f"ALTER TABLE {tabla} ADD COLUMN {nombre} {definicion}"))
        conn.commit()


def init_db():
    with app.app_context():
        db.create_all()
        migrar_esquema()
        seed_usuarios()


init_db()

if __name__ == '__main__':
    app.run(debug=True)
