"""Migración segura de la base SQLite.

Crea las tablas nuevas (db.create_all) y agrega de forma idempotente las
columnas que se hayan incorporado a modelos ya existentes (SQLite no las
agrega solo). No borra datos. Ejecutar con: python migrate.py
"""
from sqlalchemy import inspect, text

from app import app, db, init_db, BASE_DIR

# Columnas nuevas por tabla -> definición SQL de la columna.
# Se agregan solo si la tabla ya existe y la columna aún no está.
COLUMNAS_NUEVAS = {
    'usuario': {
        'rol': "VARCHAR(20) NOT NULL DEFAULT 'chofer'",
        'telefono': "VARCHAR(20)",
        'vehiculo_id': "INTEGER",
        'activo': "BOOLEAN NOT NULL DEFAULT 1",
    },
    'vehiculo': {
        'kilometraje_actual': "FLOAT NOT NULL DEFAULT 0",
        'estado': "VARCHAR(20) NOT NULL DEFAULT 'Activo'",
    },
    'viaje': {
        'monto_flete': "FLOAT NOT NULL DEFAULT 0",
        'porcentaje_comision': "FLOAT NOT NULL DEFAULT 0",
        'monto_comision': "FLOAT NOT NULL DEFAULT 0",
        'estado': "VARCHAR(20) NOT NULL DEFAULT 'Asignado'",
        'guia': "VARCHAR(255)",
    },
    'evento_vehiculo': {
        'litros': "FLOAT DEFAULT 0",
    },
}


def agregar_columnas_faltantes():
    inspector = inspect(db.engine)
    tablas = set(inspector.get_table_names())
    with db.engine.begin() as conn:
        for tabla, columnas in COLUMNAS_NUEVAS.items():
            if tabla not in tablas:
                continue  # la crea db.create_all()
            existentes = {c['name'] for c in inspector.get_columns(tabla)}
            for col, definicion in columnas.items():
                if col not in existentes:
                    print(f"  + {tabla}.{col}")
                    conn.execute(text(f'ALTER TABLE {tabla} ADD COLUMN {col} {definicion}'))


if __name__ == '__main__':
    print(f"Base de datos en: {BASE_DIR}/database.db")
    with app.app_context():
        print("Agregando columnas nuevas (si faltan)...")
        agregar_columnas_faltantes()
    print("Creando tablas nuevas y usuarios semilla...")
    init_db()
    print("Migración completada.")
