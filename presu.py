# -*- coding: utf-8 -*-
""" Sistema de Formulación Presupuestal y Control de Costos
Versión 2.1.0 - Importación unificada desde un único Excel, BudgetEngine
refactorizado, PDF real y correos segmentados. """

import os
import sys
import sqlite3
import datetime
import smtplib
import webbrowser
import random
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from io import BytesIO
from collections import defaultdict

import pandas as pd
import numpy as np
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, PatternFill, Alignment

import matplotlib
matplotlib.use('TkAgg')
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk

# ReportLab para PDF
try:
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib import colors
    from reportlab.lib.units import inch
    HAS_REPORTLAB = True
except ImportError:
    HAS_REPORTLAB = False

import customtkinter as ctk

# Configuración de customtkinter
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# Constantes
APP_TITLE = "Sistema de Presupuestos y Control de Costos"
VERSION = "2.1.0"
DB_FILE = "budget_db.sqlite"


# ---------------------- BASE DE DATOS ----------------------
class Database:
    def __init__(self, db_file=DB_FILE):
        self.conn = sqlite3.connect(db_file)
        self.cursor = self.conn.cursor()
        self.create_tables()

    def create_tables(self):
        # Tablas principales
        self.cursor.execute('''
        CREATE TABLE IF NOT EXISTS periods (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE,
            start_date TEXT,
            end_date TEXT,
            is_current INTEGER DEFAULT 0
        )
        ''')
        self.cursor.execute('''
        CREATE TABLE IF NOT EXISTS versions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE,
            description TEXT,
            is_active INTEGER DEFAULT 0
        )
        ''')
        self.cursor.execute('''
        CREATE TABLE IF NOT EXISTS centers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT UNIQUE,
            name TEXT,
            type TEXT,
            manager_email TEXT,
            is_active INTEGER DEFAULT 1
        )
        ''')
        self.cursor.execute('''
        CREATE TABLE IF NOT EXISTS elements (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT UNIQUE,
            name TEXT,
            category TEXT,
            is_active INTEGER DEFAULT 1
        )
        ''')
        self.cursor.execute('''
        CREATE TABLE IF NOT EXISTS budget_data (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            period_id INTEGER,
            version_id INTEGER,
            center_id INTEGER,
            element_id INTEGER,
            month INTEGER,
            year INTEGER,
            budget_original REAL DEFAULT 0,
            actual_original REAL DEFAULT 0,
            budget_revised REAL DEFAULT 0,
            actual_revised REAL DEFAULT 0,
            FOREIGN KEY(period_id) REFERENCES periods(id),
            FOREIGN KEY(version_id) REFERENCES versions(id),
            FOREIGN KEY(center_id) REFERENCES centers(id),
            FOREIGN KEY(element_id) REFERENCES elements(id)
        )
        ''')
        self.cursor.execute('''
        CREATE TABLE IF NOT EXISTS forecast_config (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            period_id INTEGER,
            cut_month INTEGER,
            FOREIGN KEY(period_id) REFERENCES periods(id)
        )
        ''')
        self.cursor.execute('''
        CREATE TABLE IF NOT EXISTS investments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            period_id INTEGER,
            description TEXT,
            amount REAL,
            year INTEGER,
            FOREIGN KEY(period_id) REFERENCES periods(id)
        )
        ''')
        self.cursor.execute('''
        CREATE TABLE IF NOT EXISTS financial_statements (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            period_id INTEGER,
            version_id INTEGER,
            statement_type TEXT,
            account_code TEXT,
            account_name TEXT,
            amount REAL,
            year INTEGER,
            FOREIGN KEY(period_id) REFERENCES periods(id),
            FOREIGN KEY(version_id) REFERENCES versions(id)
        )
        ''')
        self.cursor.execute('''
        CREATE TABLE IF NOT EXISTS variances (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            period_id INTEGER,
            center_id INTEGER,
            element_id INTEGER,
            month INTEGER,
            year INTEGER,
            var_type TEXT,
            amount REAL,
            percentage REAL,
            FOREIGN KEY(period_id) REFERENCES periods(id),
            FOREIGN KEY(center_id) REFERENCES centers(id),
            FOREIGN KEY(element_id) REFERENCES elements(id)
        )
        ''')
        # Tablas para productos, drivers y presupuestos detallados
        self.cursor.execute('''
        CREATE TABLE IF NOT EXISTS products (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT UNIQUE,
            name TEXT,
            unit TEXT,
            price REAL,
            standard_cost REAL,
            is_active INTEGER DEFAULT 1
        )
        ''')
        self.cursor.execute('''
        CREATE TABLE IF NOT EXISTS drivers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT UNIQUE,
            description TEXT,
            value REAL,
            is_active INTEGER DEFAULT 1
        )
        ''')
        self.cursor.execute('''
        CREATE TABLE IF NOT EXISTS sales_budget (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            period_id INTEGER,
            version_id INTEGER,
            product_id INTEGER,
            year INTEGER,
            month INTEGER,
            units INTEGER,
            price REAL,
            revenue REAL,
            FOREIGN KEY(period_id) REFERENCES periods(id),
            FOREIGN KEY(version_id) REFERENCES versions(id),
            FOREIGN KEY(product_id) REFERENCES products(id)
        )
        ''')
        self.cursor.execute('''
        CREATE TABLE IF NOT EXISTS inventory_budget (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            period_id INTEGER,
            version_id INTEGER,
            product_id INTEGER,
            year INTEGER,
            month INTEGER,
            initial_stock REAL,
            production REAL,
            sales REAL,
            final_stock REAL,
            FOREIGN KEY(period_id) REFERENCES periods(id),
            FOREIGN KEY(version_id) REFERENCES versions(id),
            FOREIGN KEY(product_id) REFERENCES products(id)
        )
        ''')
        self.cursor.execute('''
        CREATE TABLE IF NOT EXISTS purchases_budget (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            period_id INTEGER,
            version_id INTEGER,
            product_id INTEGER,
            year INTEGER,
            month INTEGER,
            required_raw_material REAL,
            purchases REAL,
            FOREIGN KEY(period_id) REFERENCES periods(id),
            FOREIGN KEY(version_id) REFERENCES versions(id),
            FOREIGN KEY(product_id) REFERENCES products(id)
        )
        ''')
        self.conn.commit()

    def execute(self, sql, params=()):
        self.cursor.execute(sql, params)
        self.conn.commit()
        return self.cursor

    def fetch_all(self, sql, params=()):
        self.cursor.execute(sql, params)
        return self.cursor.fetchall()

    def fetch_one(self, sql, params=()):
        self.cursor.execute(sql, params)
        return self.cursor.fetchone()

    def close(self):
        self.conn.close()


# ---------------------- MODELO DE DATOS ----------------------
class BudgetModel:
    def __init__(self, db):
        self.db = db

    # Períodos
    def add_period(self, name, start, end):
        self.db.execute("INSERT INTO periods (name, start_date, end_date) VALUES (?,?,?)", (name, start, end))
        return self.db.cursor.lastrowid

    def get_periods(self):
        return self.db.fetch_all("SELECT * FROM periods ORDER BY id")

    # Versiones
    def add_version(self, name, desc):
        self.db.execute("INSERT INTO versions (name, description) VALUES (?,?)", (name, desc))
        return self.db.cursor.lastrowid

    def get_versions(self):
        return self.db.fetch_all("SELECT * FROM versions ORDER BY id")

    # Centros
    def add_center(self, code, name, ctype, email):
        existing = self.db.fetch_one("SELECT id FROM centers WHERE code=?", (code,))
        if existing:
            return existing[0]
        self.db.execute("INSERT INTO centers (code, name, type, manager_email) VALUES (?,?,?,?)", (code, name, ctype, email))
        return self.db.cursor.lastrowid

    def get_centers(self):
        return self.db.fetch_all("SELECT * FROM centers WHERE is_active=1")

    # Elementos
    def add_element(self, code, name, category):
        existing = self.db.fetch_one("SELECT id FROM elements WHERE code=?", (code,))
        if existing:
            return existing[0]
        self.db.execute("INSERT INTO elements (code, name, category) VALUES (?,?,?)", (code, name, category))
        return self.db.cursor.lastrowid

    def get_elements(self):
        return self.db.fetch_all("SELECT * FROM elements WHERE is_active=1")

    # Productos
    def add_product(self, code, name, unit, price, standard_cost):
        existing = self.db.fetch_one("SELECT id FROM products WHERE code=?", (code,))
        if existing:
            return existing[0]
        self.db.execute("INSERT INTO products (code, name, unit, price, standard_cost) VALUES (?,?,?,?,?)",
                        (code, name, unit, price, standard_cost))
        return self.db.cursor.lastrowid

    def get_products(self):
        return self.db.fetch_all("SELECT * FROM products WHERE is_active=1")

    # Drivers
    def add_driver(self, code, description, value):
        existing = self.db.fetch_one("SELECT id FROM drivers WHERE code=?", (code,))
        if existing:
            return existing[0]
        self.db.execute("INSERT INTO drivers (code, description, value) VALUES (?,?,?)",
                        (code, description, value))
        return self.db.cursor.lastrowid

    def get_drivers(self):
        return self.db.fetch_all("SELECT * FROM drivers WHERE is_active=1")

    # Ventas
    def save_sales_budget(self, period_id, version_id, product_id, year, month, units, price, revenue):
        self.db.execute('''
            INSERT INTO sales_budget (period_id, version_id, product_id, year, month, units, price, revenue)
            VALUES (?,?,?,?,?,?,?,?)
        ''', (period_id, version_id, product_id, year, month, units, price, revenue))

    def get_sales_budget(self, period_id, version_id, product_id=None, year=None, month=None):
        sql = "SELECT * FROM sales_budget WHERE period_id=? AND version_id=?"
        params = [period_id, version_id]
        if product_id:
            sql += " AND product_id=?"
            params.append(product_id)
        if year:
            sql += " AND year=?"
            params.append(year)
        if month is not None:
            sql += " AND month=?"
            params.append(month)
        return self.db.fetch_all(sql, tuple(params))

    # Inventarios
    def save_inventory_budget(self, period_id, version_id, product_id, year, month,
                              initial_stock, production, sales, final_stock):
        self.db.execute('''
            INSERT INTO inventory_budget
            (period_id, version_id, product_id, year, month, initial_stock, production, sales, final_stock)
            VALUES (?,?,?,?,?,?,?,?,?)
        ''', (period_id, version_id, product_id, year, month, initial_stock, production, sales, final_stock))

    def get_inventory_budget(self, period_id, version_id, product_id=None, year=None):
        sql = "SELECT * FROM inventory_budget WHERE period_id=? AND version_id=?"
        params = [period_id, version_id]
        if product_id:
            sql += " AND product_id=?"
            params.append(product_id)
        if year:
            sql += " AND year=?"
            params.append(year)
        return self.db.fetch_all(sql, tuple(params))

    # Compras
    def save_purchases_budget(self, period_id, version_id, product_id, year, month, required_raw, purchases):
        self.db.execute('''
            INSERT INTO purchases_budget
            (period_id, version_id, product_id, year, month, required_raw_material, purchases)
            VALUES (?,?,?,?,?,?,?)
        ''', (period_id, version_id, product_id, year, month, required_raw, purchases))

    def get_purchases_budget(self, period_id, version_id, product_id=None, year=None):
        sql = "SELECT * FROM purchases_budget WHERE period_id=? AND version_id=?"
        params = [period_id, version_id]
        if product_id:
            sql += " AND product_id=?"
            params.append(product_id)
        if year:
            sql += " AND year=?"
            params.append(year)
        return self.db.fetch_all(sql, tuple(params))

    # Budget data maestro
    def save_budget_data(self, period_id, version_id, center_id, element_id, month, year,
                         budget_orig, actual_orig, budget_rev, actual_rev):
        existing = self.db.fetch_one('''
            SELECT id FROM budget_data
            WHERE period_id=? AND version_id=? AND center_id=? AND element_id=? AND month=? AND year=?
        ''', (period_id, version_id, center_id, element_id, month, year))
        if existing:
            self.db.execute('''
                UPDATE budget_data SET
                budget_original=?, actual_original=?, budget_revised=?, actual_revised=?
                WHERE id=?
            ''', (budget_orig, actual_orig, budget_rev, actual_rev, existing[0]))
        else:
            self.db.execute('''
                INSERT INTO budget_data
                (period_id, version_id, center_id, element_id, month, year,
                budget_original, actual_original, budget_revised, actual_revised)
                VALUES (?,?,?,?,?,?,?,?,?,?)
            ''', (period_id, version_id, center_id, element_id, month, year,
                  budget_orig, actual_orig, budget_rev, actual_rev))

    def get_budget_data(self, period_id, version_id, center_id=None, element_id=None, year=None):
        sql = '''
            SELECT bd.*, c.name as center_name, e.name as element_name
            FROM budget_data bd
            JOIN centers c ON bd.center_id = c.id
            JOIN elements e ON bd.element_id = e.id
            WHERE bd.period_id=? AND bd.version_id=?
        '''
        params = [period_id, version_id]
        if center_id:
            sql += " AND bd.center_id=?"
            params.append(center_id)
        if element_id:
            sql += " AND bd.element_id=?"
            params.append(element_id)
        if year:
            sql += " AND bd.year=?"
            params.append(year)
        return self.db.fetch_all(sql, tuple(params))

    # Forecast
    def set_forecast_cut(self, period_id, cut_month):
        self.db.execute("DELETE FROM forecast_config WHERE period_id=?", (period_id,))
        self.db.execute("INSERT INTO forecast_config (period_id, cut_month) VALUES (?,?)", (period_id, cut_month))

    def get_forecast_cut(self, period_id):
        row = self.db.fetch_one("SELECT cut_month FROM forecast_config WHERE period_id=?", (period_id,))
        return row[0] if row else None

    # Inversiones
    def save_investment(self, period_id, desc, amount, year):
        self.db.execute("INSERT INTO investments (period_id, description, amount, year) VALUES (?,?,?,?)",
                        (period_id, desc, amount, year))

    def get_investments(self, period_id):
        return self.db.fetch_all("SELECT * FROM investments WHERE period_id=?", (period_id,))

    # Estados financieros
    def save_financial_statement(self, period_id, version_id, stype, acc_code, acc_name, amount, year):
        self.db.execute('''
            INSERT INTO financial_statements
            (period_id, version_id, statement_type, account_code, account_name, amount, year)
            VALUES (?,?,?,?,?,?,?)
        ''', (period_id, version_id, stype, acc_code, acc_name, amount, year))

    def get_financial_statements(self, period_id, version_id, stype=None):
        sql = "SELECT * FROM financial_statements WHERE period_id=? AND version_id=?"
        params = [period_id, version_id]
        if stype:
            sql += " AND statement_type=?"
            params.append(stype)
        return self.db.fetch_all(sql, tuple(params))

    # Variaciones
    def save_variance(self, period_id, center_id, element_id, month, year, var_type, amount, percentage):
        self.db.execute('''
            INSERT INTO variances
            (period_id, center_id, element_id, month, year, var_type, amount, percentage)
            VALUES (?,?,?,?,?,?,?,?)
        ''', (period_id, center_id, element_id, month, year, var_type, amount, percentage))

    def get_variances(self, period_id, center_id=None):
        sql = "SELECT * FROM variances WHERE period_id=?"
        params = [period_id]
        if center_id:
            sql += " AND center_id=?"
            params.append(center_id)
        return self.db.fetch_all(sql, tuple(params))


# ---------------------- IMPORTADOR UNIFICADO ----------------------
class DataImporter:
    """Importa datos desde un único archivo Excel con múltiples hojas."""
    REQUIRED_SHEETS = ["Ventas", "Inventarios", "Compras", "Producción", "Costos_Producción",
                       "Gastos_Operativos", "Inversiones"]

    def __init__(self, db_conn, model):
        self.conn = db_conn
        self.model = model

    def import_from_unified_excel(self, filepath):
        try:
            import openpyxl

            # Cargar el libro con data_only=True para evaluar fórmulas
            wb = openpyxl.load_workbook(filepath, data_only=True)

            # Función auxiliar para convertir una hoja a DataFrame
            def sheet_to_df(sheet_name):
                ws = wb[sheet_name]
                data = []
                headers = None
                for row in ws.iter_rows(values_only=True):
                    if headers is None:
                        headers = row
                    else:
                        data.append(row)
                df = pd.DataFrame(data, columns=headers)
                # Eliminar filas sin Periodo
                if 'Periodo' in df.columns:
                    df = df.dropna(subset=['Periodo'])
                return df

            # Validar hojas requeridas
            missing = [s for s in self.REQUIRED_SHEETS if s not in wb.sheetnames]
            if missing:
                raise ValueError(f"Faltan hojas: {', '.join(missing)}")

            df_ventas = sheet_to_df("Ventas")
            df_inventarios = sheet_to_df("Inventarios")
            df_compras = sheet_to_df("Compras")
            df_produccion = sheet_to_df("Producción")
            df_costos_prod = sheet_to_df("Costos_Producción")
            df_gastos_op = sheet_to_df("Gastos_Operativos")
            df_inversiones = sheet_to_df("Inversiones")

            # Procesar cada hoja
            self._process_sales(df_ventas)
            self._process_inventory(df_inventarios)
            self._process_purchases(df_compras)
            self._process_production(df_produccion)
            self._process_production_costs(df_costos_prod)
            self._process_operating_expenses(df_gastos_op)
            self._process_investments(df_inversiones)

            self.conn.commit()
            return True, "Importación completada exitosamente."
        except Exception as e:
            self.conn.rollback()
            return False, f"Error: {str(e)}"

    # Métodos auxiliares de extracción
    def _extract_year_month(self, period_str):
        try:
            y, m = map(int, period_str.split('-'))
            return y, m
        except:
            return 2026, 1

    def _get_or_create_product(self, name):
        cursor = self.conn.cursor()
        cursor.execute("SELECT id FROM products WHERE name=?", (name,))
        row = cursor.fetchone()
        if row:
            return row[0]
        code = f"P{name[:3].upper()}{random.randint(100,999)}"
        cursor.execute("INSERT INTO products (code, name, unit, price, standard_cost) VALUES (?,?,?,?,?)",
                       (code, name, "unidad", 0.0, 0.0))
        return cursor.lastrowid

    def _get_or_create_version(self, name):
        cursor = self.conn.cursor()
        cursor.execute("SELECT id FROM versions WHERE name=?", (name,))
        row = cursor.fetchone()
        if row:
            return row[0]
        cursor.execute("INSERT INTO versions (name, description) VALUES (?,?)", (name, f"Importado de {name}"))
        return cursor.lastrowid

    def _get_or_create_period(self, name):
        cursor = self.conn.cursor()
        cursor.execute("SELECT id FROM periods WHERE name=?", (name,))
        row = cursor.fetchone()
        if row:
            return row[0]
        year, month = self._extract_year_month(name)
        start = f"{year}-{month:02d}-01"
        if month == 12:
            end = f"{year+1}-01-01"
        else:
            end = f"{year}-{month+1:02d}-01"
        cursor.execute("INSERT INTO periods (name, start_date, end_date) VALUES (?,?,?)", (name, start, end))
        return cursor.lastrowid

    def _get_or_create_center(self, name, ctype):
        cursor = self.conn.cursor()
        cursor.execute("SELECT id FROM centers WHERE name=?", (name,))
        row = cursor.fetchone()
        if row:
            return row[0]
        code = f"C{name[:3].upper()}{random.randint(100,999)}"
        cursor.execute("INSERT INTO centers (code, name, type, manager_email) VALUES (?,?,?,?)",
                       (code, name, ctype, ""))
        return cursor.lastrowid

    def _get_or_create_element(self, name, category):
        cursor = self.conn.cursor()
        cursor.execute("SELECT id FROM elements WHERE name=?", (name,))
        row = cursor.fetchone()
        if row:
            return row[0]
        code = f"E{name[:3].upper()}{random.randint(100,999)}"
        cursor.execute("INSERT INTO elements (code, name, category) VALUES (?,?,?)", (code, name, category))
        return cursor.lastrowid

    # Procesamiento de hojas
    def _process_sales(self, df):
        for _, row in df.iterrows():
            year, month = self._extract_year_month(row['Periodo'])
            product_id = self._get_or_create_product(row['Producto'])
            version_id = self._get_or_create_version(row['Versión'])
            period_id = self._get_or_create_period(row['Periodo'])
            self.model.save_sales_budget(
                period_id, version_id, product_id, year, month,
                units=row['Cant. Presupuestada'],
                price=row['Precio Presup. (S/)'],
                revenue=row['Ingreso Presup. (S/)']
            )

    def _process_inventory(self, df):
        for _, row in df.iterrows():
            year, month = self._extract_year_month(row['Periodo'])
            product_id = self._get_or_create_product(row['Ítem / Código'])
            version_id = self._get_or_create_version(row['Versión'])
            period_id = self._get_or_create_period(row['Periodo'])
            self.model.save_inventory_budget(
                period_id, version_id, product_id, year, month,
                initial_stock=row['Inv. Inicial Presup.'],
                production=0,
                sales=0,
                final_stock=row['Inv. Final Presup.']
            )

    def _process_purchases(self, df):
        for _, row in df.iterrows():
            year, month = self._extract_year_month(row['Periodo'])
            product_id = self._get_or_create_product(row['Insumo'])
            version_id = self._get_or_create_version(row['Versión'])
            period_id = self._get_or_create_period(row['Periodo'])
            self.model.save_purchases_budget(
                period_id, version_id, product_id, year, month,
                required_raw=row['Cant. Compra Presup.'],
                purchases=row['Total Presup. (S/)']
            )

    def _process_production(self, df):
        # Se puede usar para actualizar budget_data si se requiere
        pass

    def _process_production_costs(self, df):
        for _, row in df.iterrows():
            year, month = self._extract_year_month(row['Periodo'])
            center_id = self._get_or_create_center(row['Centro de Actividad Productiva'], "Productivo")
            element_id = self._get_or_create_element(row['Elemento de Costo'], "Costo")
            version_id = self._get_or_create_version(row['Versión'])
            period_id = self._get_or_create_period(row['Periodo'])
            self.model.save_budget_data(
                period_id, version_id, center_id, element_id, month, year,
                budget_orig=row['Monto Presupuestado (S/)'],
                actual_orig=row['Monto Real (S/)'],
                budget_rev=0,
                actual_rev=0
            )

    def _process_operating_expenses(self, df):
        for _, row in df.iterrows():
            year, month = self._extract_year_month(row['Periodo'])
            center_id = self._get_or_create_center(row['Centro de Gasto Operativo'], "Gasto")
            element_id = self._get_or_create_element(row['Subcategoría de Gasto'], "Gasto")
            version_id = self._get_or_create_version(row['Versión'])
            period_id = self._get_or_create_period(row['Periodo'])
            self.model.save_budget_data(
                period_id, version_id, center_id, element_id, month, year,
                budget_orig=row['Monto Presupuestado (S/)'],
                actual_orig=row['Monto Real (S/)'],
                budget_rev=0,
                actual_rev=0
            )

    def _process_investments(self, df):
        for _, row in df.iterrows():
            year, _ = self._extract_year_month(row['Periodo'])
            period_id = self._get_or_create_period(row['Periodo'])
            self.model.save_investment(
                period_id,
                desc=row['Proyecto de Inversión (CapEx)'],
                amount=row['Desembolso Presup. (S/)'],
                year=year
            )


# ---------------------- MOTOR DE PRESUPUESTO (REFACTORIZADO) ----------------------
class BudgetEngine:
    def __init__(self, model):
        self.model = model

    def run_cascade(self, period_id, version_id):
        period = self.model.db.fetch_one("SELECT start_date, end_date FROM periods WHERE id=?", (period_id,))
        if not period:
            return
        start_year = int(period[0][:4])
        end_year = int(period[1][:4])
        years = list(range(start_year, end_year + 1))

        centers = self.model.get_centers()
        elements = self.model.get_elements()
        products = self.model.get_products()
        cut_month = self.model.get_forecast_cut(period_id) or 4

        # 1. Ventas
        self._generate_sales_budget(period_id, version_id, products, years)
        # 2. Inventarios
        self._generate_inventory_budget(period_id, version_id, products, years)
        # 3. Compras
        self._generate_purchases_budget(period_id, version_id, products, years)

        # 4. Budget data maestro (híbrido)
        for center in centers:
            for element in elements:
                for year in years:
                    if year == start_year:
                        for month in range(1, 13):
                            self._process_budget_line(period_id, version_id, center[0], element[0],
                                                      month, year, cut_month)
                    else:
                        self._process_budget_line(period_id, version_id, center[0], element[0],
                                                  0, year, cut_month)

        # 5. Estados financieros
        self._generate_financial_statements(period_id, version_id, years)
        # 6. Variaciones
        self._calculate_variances(period_id, version_id)

    def _process_budget_line(self, period_id, version_id, center_id, element_id, month, year, cut_month):
        existing = self.model.db.fetch_one('''
            SELECT id FROM budget_data
            WHERE period_id=? AND version_id=? AND center_id=? AND element_id=? AND month=? AND year=?
        ''', (period_id, version_id, center_id, element_id, month, year))
        if existing:
            return

        if month == 0:
            budget_orig = random.uniform(1200, 12000)
            actual_orig = budget_orig * random.uniform(0.8, 1.1)
            budget_rev = budget_orig * random.uniform(0.95, 1.05)
            actual_rev = actual_orig * random.uniform(0.9, 1.05)
        else:
            budget_orig = random.uniform(100, 1000)
            actual_orig = budget_orig * random.uniform(0.8, 1.1)
            if month <= cut_month:
                budget_rev = actual_orig
            else:
                prev = self.model.db.fetch_all('''
                    SELECT actual_original FROM budget_data
                    WHERE period_id=? AND version_id=? AND center_id=? AND element_id=? AND year=? AND month < ?
                ''', (period_id, version_id, center_id, element_id, year, month))
                if prev:
                    avg_actual = sum([r[0] for r in prev]) / len(prev)
                    budget_rev = avg_actual * 1.05
                else:
                    budget_rev = budget_orig * 1.05
            actual_rev = actual_orig * random.uniform(0.9, 1.05)

        self.model.save_budget_data(period_id, version_id, center_id, element_id,
                                    month, year, budget_orig, actual_orig, budget_rev, actual_rev)

    def _generate_sales_budget(self, period_id, version_id, products, years):
        for product in products:
            product_id = product[0]
            price = product[4]
            first_year_units = []
            for month in range(1, 13):
                units = random.randint(50, 200)
                first_year_units.append(units)
                revenue = units * price
                self.model.save_sales_budget(period_id, version_id, product_id,
                                             years[0], month, units, price, revenue)
            for year in years[1:]:
                total_units = int(sum(first_year_units) * (1.05 ** (year - years[0])))
                self.model.save_sales_budget(period_id, version_id, product_id,
                                             year, 0, total_units, price, total_units * price)

    def _generate_inventory_budget(self, period_id, version_id, products, years):
        initial_stock_map = {}
        for product in products:
            product_id = product[0]
            sales_m1 = self.model.get_sales_budget(period_id, version_id, product_id, years[0], 1)
            initial_stock = sales_m1[0][6] * 0.10 if sales_m1 else 50
            for year in years:
                if year == years[0]:
                    for month in range(1, 13):
                        sales_data = self.model.get_sales_budget(period_id, version_id, product_id, year, month)
                        units_sold = sales_data[0][6] if sales_data else 0
                        desired_ending = units_sold * 0.20
                        production = max(0, units_sold + desired_ending - initial_stock)
                        final_stock = initial_stock + production - units_sold
                        self.model.save_inventory_budget(period_id, version_id, product_id,
                                                         year, month,
                                                         initial_stock, production, units_sold, final_stock)
                        initial_stock = final_stock
                    initial_stock_map[(product_id, years[0])] = initial_stock
                else:
                    sales_data = self.model.get_sales_budget(period_id, version_id, product_id, year, 0)
                    units_sold = sales_data[0][6] if sales_data else 0
                    desired_ending = units_sold * 0.20
                    prev_year = year - 1
                    initial_stock = initial_stock_map.get((product_id, prev_year), 50)
                    production = max(0, units_sold + desired_ending - initial_stock)
                    final_stock = initial_stock + production - units_sold
                    self.model.save_inventory_budget(period_id, version_id, product_id,
                                                     year, 0,
                                                     initial_stock, production, units_sold, final_stock)
                    initial_stock_map[(product_id, year)] = final_stock

    def _generate_purchases_budget(self, period_id, version_id, products, years):
        for product in products:
            product_id = product[0]
            std_cost = product[5]
            for year in years:
                inv_data = self.model.get_inventory_budget(period_id, version_id, product_id, year)
                for record in inv_data:
                    month = record[5]
                    production = record[6]
                    required_raw = production
                    purchases_value = required_raw * std_cost
                    self.model.save_purchases_budget(period_id, version_id, product_id,
                                                     year, month,
                                                     required_raw, purchases_value)

    def _generate_financial_statements(self, period_id, version_id, years):
        self.model.db.execute("DELETE FROM financial_statements WHERE period_id=? AND version_id=?",
                              (period_id, version_id))
        data = self.model.get_budget_data(period_id, version_id)
        if not data:
            for year in years:
                for acc in ['Ingresos', 'Costo de Ventas', 'Gastos Administrativos', 'Gastos de Ventas',
                            'Otros Ingresos', 'Otros Gastos']:
                    amount = random.uniform(10000, 50000) if 'Ingresos' in acc else random.uniform(1000, 20000)
                    self.model.save_financial_statement(period_id, version_id, 'IS', acc[:10], acc, amount, year)
                for acc in ['Activo Corriente', 'Activo No Corriente', 'Pasivo Corriente',
                            'Pasivo No Corriente', 'Patrimonio']:
                    amount = random.uniform(20000, 100000)
                    self.model.save_financial_statement(period_id, version_id, 'BS', acc[:10], acc, amount, year)
            return

        df = pd.DataFrame(data,
                          columns=['id', 'period_id', 'version_id', 'center_id', 'element_id', 'month', 'year',
                                   'budget_original', 'actual_original', 'budget_revised', 'actual_revised',
                                   'center_name', 'element_name'])
        for year in years:
            df_year = df[df['year'] == year]
            if df_year.empty:
                continue
            grouped = df_year.groupby('element_name')['actual_original'].sum()
            for acc in ['Ingresos', 'Costo de Ventas', 'Gastos Administrativos', 'Gastos de Ventas',
                        'Otros Ingresos', 'Otros Gastos']:
                amount = grouped.get(acc, random.uniform(1000, 20000))
                self.model.save_financial_statement(period_id, version_id, 'IS', acc[:10], acc, amount, year)
            for acc in ['Activo Corriente', 'Activo No Corriente', 'Pasivo Corriente',
                        'Pasivo No Corriente', 'Patrimonio']:
                amount = random.uniform(20000, 100000)
                self.model.save_financial_statement(period_id, version_id, 'BS', acc[:10], acc, amount, year)

    def _calculate_variances(self, period_id, version_id):
        self.model.db.execute("DELETE FROM variances WHERE period_id=?", (period_id,))
        data = self.model.get_budget_data(period_id, version_id)
        for row in data:
            budget_orig = row[7]
            actual_orig = row[8]
            budget_rev = row[9]
            actual_rev = row[10]
            var_amount = actual_orig - budget_orig
            var_perc = (var_amount / budget_orig * 100) if budget_orig != 0 else 0
            self.model.save_variance(period_id, row[3], row[4], row[5], row[6],
                                     'Original vs Real', var_amount, var_perc)
            var_amount2 = actual_rev - budget_rev
            var_perc2 = (var_amount2 / budget_rev * 100) if budget_rev != 0 else 0
            self.model.save_variance(period_id, row[3], row[4], row[5], row[6],
                                     'Forecast vs Real', var_amount2, var_perc2)


# ---------------------- ANÁLISIS FINANCIERO ----------------------
class FinancialAnalyzer:
    def __init__(self, model):
        self.model = model

    def vertical_analysis(self, period_id, version_id, stype):
        stmts = self.model.get_financial_statements(period_id, version_id, stype)
        if not stmts:
            return []
        years = set([s[7] for s in stmts])
        results = []
        for year in years:
            year_data = [s for s in stmts if s[7] == year]
            total = sum([s[6] for s in year_data])
            for s in year_data:
                pct = (s[6] / total * 100) if total != 0 else 0
                results.append((s[4], s[5], s[6], pct, year))
        return results

    def horizontal_analysis(self, period_id, version_id, stype):
        stmts = self.model.get_financial_statements(period_id, version_id, stype)
        if not stmts:
            return []
        df = pd.DataFrame(stmts,
                          columns=['id', 'period_id', 'version_id', 'stype', 'code', 'name', 'amount', 'year'])
        df = df.pivot_table(index=['code', 'name'], columns='year', values='amount', aggfunc='sum').reset_index()
        years = sorted([c for c in df.columns if isinstance(c, int)])
        for i in range(1, len(years)):
            prev = years[i - 1]
            curr = years[i]
            df[f'Var_{prev}_{curr}'] = df[curr] - df[prev]
            df[f'VarPct_{prev}_{curr}'] = (df[f'Var_{prev}_{curr}'] / df[prev] * 100).replace([np.inf, -np.inf], 0)
        return df

    def ratios(self, period_id, version_id, year):
        bs = self.model.get_financial_statements(period_id, version_id, 'BS')
        is_ = self.model.get_financial_statements(period_id, version_id, 'IS')
        bs_dict = {s[4]: s[6] for s in bs if s[7] == year}
        is_dict = {s[4]: s[6] for s in is_ if s[7] == year}
        current_assets = bs_dict.get('Activo Corriente', 0)
        current_liab = bs_dict.get('Pasivo Corriente', 0)
        liquidity = current_assets / current_liab if current_liab else 0
        total_assets = bs_dict.get('Activo No Corriente', 0) + current_assets
        total_liab = bs_dict.get('Pasivo No Corriente', 0) + current_liab
        solvency = total_assets / total_liab if total_liab else 0
        net_income = is_dict.get('Ingresos', 0) - is_dict.get('Costo de Ventas', 0) - is_dict.get('Gastos Administrativos', 0) - is_dict.get('Gastos de Ventas', 0)
        equity = bs_dict.get('Patrimonio', 0)
        roe = (net_income / equity * 100) if equity else 0
        return {'Liquidez': liquidity, 'Solvencia': solvency, 'ROE': roe}

    def dupont_analysis(self, period_id, version_id, year):
        bs = self.model.get_financial_statements(period_id, version_id, 'BS')
        is_ = self.model.get_financial_statements(period_id, version_id, 'IS')
        bs_dict = {s[4]: s[6] for s in bs if s[7] == year}
        is_dict = {s[4]: s[6] for s in is_ if s[7] == year}
        ventas = is_dict.get('Ingresos', 0)
        net_income = ventas - is_dict.get('Costo de Ventas', 0) - is_dict.get('Gastos Administrativos', 0) - is_dict.get('Gastos de Ventas', 0)
        total_assets = bs_dict.get('Activo Corriente', 0) + bs_dict.get('Activo No Corriente', 0)
        equity = bs_dict.get('Patrimonio', 0)
        margen = (net_income / ventas) if ventas else 0
        rotacion = (ventas / total_assets) if total_assets else 0
        apalancamiento = (total_assets / equity) if equity else 0
        roe = margen * rotacion * apalancamiento
        return {'Margen': margen, 'Rotacion': rotacion, 'Apalancamiento': apalancamiento, 'ROE': roe}

    def eva(self, period_id, version_id, year, wacc=0.1):
        is_ = self.model.get_financial_statements(period_id, version_id, 'IS')
        bs = self.model.get_financial_statements(period_id, version_id, 'BS')
        is_dict = {s[4]: s[6] for s in is_ if s[7] == year}
        bs_dict = {s[4]: s[6] for s in bs if s[7] == year}
        ebit = is_dict.get('Ingresos', 0) - is_dict.get('Costo de Ventas', 0) - is_dict.get('Gastos Administrativos', 0) - is_dict.get('Gastos de Ventas', 0)
        nopat = ebit * (1 - 0.3)
        capital = bs_dict.get('Activo No Corriente', 0) + bs_dict.get('Activo Corriente', 0)
        eva = nopat - (capital * wacc)
        return {'NOPAT': nopat, 'Capital': capital, 'WACC': wacc, 'EVA': eva}

    def pareto_cost(self, period_id, version_id, year):
        data = self.model.get_budget_data(period_id, version_id)
        df = pd.DataFrame(data,
                          columns=['id', 'period_id', 'version_id', 'center_id', 'element_id', 'month', 'year',
                                   'budget_original', 'actual_original', 'budget_revised', 'actual_revised',
                                   'center_name', 'element_name'])
        df = df[df['year'] == year]
        cost_by_element = df.groupby('element_name')['actual_original'].sum().reset_index()
        cost_by_element = cost_by_element.sort_values('actual_original', ascending=False)
        total_cost = cost_by_element['actual_original'].sum()
        cost_by_element['cum_pct'] = cost_by_element['actual_original'].cumsum() / total_cost * 100
        pareto = cost_by_element[cost_by_element['cum_pct'] <= 80]
        return pareto, cost_by_element

    def generate_variance_multidimensional_analysis(self, period_id):
        query = """
            SELECT
                bd.year,
                bd.month,
                c.name AS center_name,
                c.type AS center_type,
                e.name AS element_name,
                e.category AS element_category,
                bd.budget_original,
                bd.actual_original,
                bd.budget_revised,
                bd.actual_revised
            FROM budget_data bd
            JOIN centers c ON bd.center_id = c.id
            JOIN elements e ON bd.element_id = e.id
            WHERE bd.period_id = ?
        """
        df = pd.read_sql_query(query, self.model.db.conn, params=(period_id,))
        if df.empty:
            return {}
        df['var_abs_orig'] = df['actual_original'] - df['budget_original']
        df['var_pct_orig'] = np.where(df['budget_original'] != 0, (df['var_abs_orig'] / df['budget_original']) * 100, 0)
        df['var_abs_rev'] = df['actual_revised'] - df['budget_revised']
        df['var_pct_rev'] = np.where(df['budget_revised'] != 0, (df['var_abs_rev'] / df['budget_revised']) * 100, 0)
        df['tipo_var'] = df.apply(
            lambda row: 'Desfavorable' if (
                (row['element_category'] in ['Gasto', 'Costo'] and row['var_abs_orig'] > 0) or
                (row['element_category'] == 'Ingresos' and row['var_abs_orig'] < 0)
            ) else 'Favorable',
            axis=1
        )
        por_periodo = df.groupby(['year', 'month'], as_index=False).agg({
            'var_abs_orig': 'sum',
            'var_pct_orig': 'mean',
            'var_abs_rev': 'sum',
            'var_pct_rev': 'mean'
        }).round(2)
        por_centro = df.groupby('center_name', as_index=False).agg({
            'var_abs_orig': 'sum',
            'var_pct_orig': 'mean',
            'var_abs_rev': 'sum',
            'var_pct_rev': 'mean'
        }).round(2)
        por_elemento = df.groupby('element_name', as_index=False).agg({
            'var_abs_orig': 'sum',
            'var_pct_orig': 'mean',
            'var_abs_rev': 'sum',
            'var_pct_rev': 'mean'
        }).round(2)
        matriz = df.pivot_table(
            index='center_name',
            columns='element_name',
            values='var_abs_orig',
            aggfunc='sum',
            fill_value=0
        ).round(2)
        return {
            'por_periodo': por_periodo,
            'por_centro': por_centro,
            'por_elemento': por_elemento,
            'matriz_cruzada': matriz.reset_index()
        }


# ---------------------- EXPORTADOR A EXCEL ----------------------
class ExcelExporter:
    def __init__(self, model):
        self.model = model

    def export_full_budget(self, period_id, version_id, filepath):
        wb = openpyxl.Workbook()
        ws1 = wb.active
        ws1.title = "Presupuesto Maestro"
        data = self.model.get_budget_data(period_id, version_id)
        if data:
            headers = ['Periodo', 'Version', 'Centro', 'Elemento', 'Mes', 'Año',
                       'Presup. Orig', 'Real Orig', 'Presup. Rev', 'Real Rev']
            rows = []
            for row in data:
                rows.append([row[1], row[2], row[11], row[12], row[5], row[6],
                             row[7], row[8], row[9], row[10]])
            df = pd.DataFrame(rows, columns=headers)
            for r in dataframe_to_rows(df, index=False, header=True):
                ws1.append(r)
        ws2 = wb.create_sheet("Estados Financieros")
        stmts = self.model.get_financial_statements(period_id, version_id)
        if stmts:
            headers2 = ['Tipo', 'Cuenta', 'Nombre', 'Monto', 'Año']
            rows2 = []
            for s in stmts:
                rows2.append([s[3], s[4], s[5], s[6], s[7]])
            df2 = pd.DataFrame(rows2, columns=headers2)
            for r in dataframe_to_rows(df2, index=False, header=True):
                ws2.append(r)
        ws3 = wb.create_sheet("Variaciones")
        vars_ = self.model.get_variances(period_id)
        if vars_:
            headers3 = ['Centro', 'Elemento', 'Mes', 'Año', 'Tipo', 'Monto', '%']
            rows3 = []
            for v in vars_:
                center = self.model.db.fetch_one("SELECT name FROM centers WHERE id=?", (v[2],))
                element = self.model.db.fetch_one("SELECT name FROM elements WHERE id=?", (v[3],))
                rows3.append([center[0] if center else '', element[0] if element else '',
                              v[4], v[5], v[6], v[7], v[8]])
            df3 = pd.DataFrame(rows3, columns=headers3)
            for r in dataframe_to_rows(df3, index=False, header=True):
                ws3.append(r)
        wb.save(filepath)


# ---------------------- EXPORTADOR A PDF ----------------------
class PDFExporter:
    def __init__(self, model):
        self.model = model
        if not HAS_REPORTLAB:
            print("ReportLab no instalado. Instale con: pip install reportlab")

    def export_report(self, period_id, version_id, filepath):
        if not HAS_REPORTLAB:
            messagebox.showerror("Error", "ReportLab no instalado. Instale con: pip install reportlab")
            return
        doc = SimpleDocTemplate(filepath, pagesize=letter,
                                rightMargin=72, leftMargin=72,
                                topMargin=72, bottomMargin=72)
        styles = getSampleStyleSheet()
        story = []

        title_style = styles['Title']
        story.append(Paragraph("Informe de Presupuesto", title_style))
        story.append(Spacer(1, 0.25 * inch))

        period = self.model.db.fetch_one("SELECT name FROM periods WHERE id=?", (period_id,))
        version = self.model.db.fetch_one("SELECT name FROM versions WHERE id=?", (version_id,))
        period_name = period[0] if period else "N/A"
        version_name = version[0] if version else "N/A"
        story.append(Paragraph(f"Período: {period_name} | Versión: {version_name}", styles['Normal']))
        story.append(Spacer(1, 0.25 * inch))

        # Tabla de variaciones
        data = self.model.get_budget_data(period_id, version_id)
        if data:
            df = pd.DataFrame(data,
                              columns=['id', 'period_id', 'version_id', 'center_id', 'element_id', 'month', 'year',
                                       'budget_original', 'actual_original', 'budget_revised', 'actual_revised',
                                       'center_name', 'element_name'])
            grouped = df.groupby('element_name').agg({
                'budget_original': 'sum',
                'actual_original': 'sum'
            }).reset_index()
            grouped['var_abs'] = grouped['actual_original'] - grouped['budget_original']
            grouped['var_pct'] = (grouped['var_abs'] / grouped['budget_original'] * 100).round(2)
            grouped = grouped.sort_values('var_abs', ascending=False)

            table_data = [['Elemento', 'Presupuesto', 'Real', 'Variación Abs.', 'Variación %']]
            for _, row in grouped.iterrows():
                table_data.append([
                    row['element_name'],
                    f"{row['budget_original']:,.2f}",
                    f"{row['actual_original']:,.2f}",
                    f"{row['var_abs']:+,.2f}",
                    f"{row['var_pct']:+.2f}%"
                ])
            t = Table(table_data, colWidths=[2 * inch, 1.2 * inch, 1.2 * inch, 1.2 * inch, 1.2 * inch])
            t.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 12),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ]))
            story.append(Paragraph("Análisis de Variaciones", styles['Heading2']))
            story.append(Spacer(1, 0.1 * inch))
            story.append(t)
            story.append(Spacer(1, 0.25 * inch))

            # Estado de Resultados
            is_data = self.model.get_financial_statements(period_id, version_id, 'IS')
            if is_data:
                df_is = pd.DataFrame(is_data,
                                     columns=['id', 'period_id', 'version_id', 'stype', 'code', 'name', 'amount', 'year'])
                pivoted = df_is.pivot(index='name', columns='year', values='amount').fillna(0).round(2)
                table_data2 = [['Cuenta'] + [str(y) for y in pivoted.columns]]
                for idx, row in pivoted.iterrows():
                    table_data2.append([idx] + [f"{v:,.2f}" for v in row])
                t2 = Table(table_data2, colWidths=[2 * inch] + [1.2 * inch] * len(pivoted.columns))
                t2.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 12),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ]))
                story.append(Paragraph("Estado de Resultados", styles['Heading2']))
                story.append(Spacer(1, 0.1 * inch))
                story.append(t2)
                story.append(Spacer(1, 0.25 * inch))

            # Gráfico Pareto
            data = self.model.get_budget_data(period_id, version_id)
            if data:
                df = pd.DataFrame(data,
                                  columns=['id', 'period_id', 'version_id', 'center_id', 'element_id', 'month', 'year',
                                           'budget_original', 'actual_original', 'budget_revised', 'actual_revised',
                                           'center_name', 'element_name'])
                year = df['year'].min()
                df_year = df[df['year'] == year]
                if not df_year.empty:
                    cost_by_element = df_year.groupby('element_name')['actual_original'].sum().sort_values(ascending=False)
                    fig, ax = plt.subplots(figsize=(8, 4))
                    ax.bar(cost_by_element.index, cost_by_element.values, color='skyblue')
                    ax.set_title('Pareto de Costos', fontsize=12)
                    ax.set_xlabel('Elemento')
                    ax.set_ylabel('Monto')
                    plt.xticks(rotation=45, ha='right')
                    plt.tight_layout()
                    img_path = "temp_pareto.png"
                    plt.savefig(img_path)
                    plt.close()
                    story.append(Paragraph("Gráfico Pareto de Costos", styles['Heading2']))
                    story.append(Spacer(1, 0.1 * inch))
                    story.append(Image(img_path, width=6 * inch, height=3 * inch))
                    story.append(Spacer(1, 0.25 * inch))
                    if os.path.exists(img_path):
                        os.remove(img_path)

        try:
            doc.build(story)
            messagebox.showinfo("Éxito", f"PDF exportado a {filepath}")
        except Exception as e:
            messagebox.showerror("Error", f"Error al generar PDF: {e}")


# ---------------------- OPTIMIZATION ADVISOR ----------------------
class OptimizationAdvisor:
    def __init__(self, model):
        self.model = model

    def generate_recommendations(self, period_id, version_id):
        query = """
            SELECT
                bd.year,
                bd.month,
                c.name AS center_name,
                e.name AS element_name,
                e.category AS element_category,
                bd.budget_original,
                bd.actual_original,
                bd.budget_revised,
                bd.actual_revised
            FROM budget_data bd
            JOIN centers c ON bd.center_id = c.id
            JOIN elements e ON bd.element_id = e.id
            WHERE bd.period_id = ? AND bd.version_id = ?
        """
        df = pd.read_sql_query(query, self.model.db.conn, params=(period_id, version_id))
        if df.empty:
            return ["No hay datos para el período y versión seleccionados."]

        df['var_abs_orig'] = df['actual_original'] - df['budget_original']
        df['var_pct_orig'] = np.where(df['budget_original'] != 0, (df['var_abs_orig'] / df['budget_original']) * 100, 0)
        df['var_abs_rev'] = df['actual_revised'] - df['budget_revised']
        df['var_pct_rev'] = np.where(df['budget_revised'] != 0, (df['var_abs_rev'] / df['budget_revised']) * 100, 0)

        df['critica_orig'] = df.apply(
            lambda row: (
                (row['element_category'] in ['Gasto', 'Costo'] and row['var_abs_orig'] > 0 and row['var_pct_orig'] > 10) or
                (row['element_category'] == 'Ingresos' and row['var_abs_orig'] < 0 and row['var_pct_orig'] < -10)
            ),
            axis=1
        )
        df['critica_rev'] = df.apply(
            lambda row: (
                (row['element_category'] in ['Gasto', 'Costo'] and row['var_abs_rev'] > 0 and row['var_pct_rev'] > 10) or
                (row['element_category'] == 'Ingresos' and row['var_abs_rev'] < 0 and row['var_pct_rev'] < -10)
            ),
            axis=1
        )

        df_crit = df[df['critica_orig'] | df['critica_rev']]
        if df_crit.empty:
            return ["✅ No se detectaron desviaciones críticas que requieran acciones inmediatas."]

        alertas = []
        for _, row in df_crit.iterrows():
            if row['critica_orig']:
                tipo = "Original"
                var_abs = row['var_abs_orig']
                var_pct = row['var_pct_orig']
            else:
                tipo = "Forecast"
                var_abs = row['var_abs_rev']
                var_pct = row['var_pct_rev']
            alertas.append({
                'year': row['year'],
                'month': row['month'],
                'center': row['center_name'],
                'element': row['element_name'],
                'categoria': row['element_category'],
                'tipo': tipo,
                'var_abs': f"{var_abs:+.2f}",
                'var_pct': f"{var_pct:+.2f}%",
                'var_abs_num': var_abs,
                'var_pct_num': var_pct
            })
        alertas.sort(key=lambda x: abs(x['var_abs_num']), reverse=True)

        propuestas = {}
        for row in alertas:
            key = (row['center'], row['element'], row['categoria'])
            propuestas.setdefault(key, []).append(row)

        lines = []
        lines.append("=" * 76)
        lines.append(" INFORME EJECUTIVO DE PROPUESTAS DE MEJORA (RENTABILIDAD)")
        lines.append("=" * 76)
        period = self.model.db.fetch_one("SELECT name FROM periods WHERE id=?", (period_id,))
        version = self.model.db.fetch_one("SELECT name FROM versions WHERE id=?", (version_id,))
        period_name = period[0] if period else "N/A"
        version_name = version[0] if version else "N/A"
        lines.append(f"Periodo Presupuestal: {period_name} | Versión: {version_name}")
        lines.append("-" * 76)

        lines.append("\n1. ALERTA CRONOLÓGICA DE FUGAS DE RENTABILIDAD CRÍTICAS")
        lines.append("-" * 76)
        header = f"{'AÑO':>4} | {'MES':>3} | {'CENTRO':<20} | {'ELEMENTO':<15} | {'TIPO PRESUP.':<12} | {'DESVIACIÓN ABS.':<15} | {'DESVIACIÓN %':<12}"
        lines.append(header)
        lines.append("-" * 76)
        for row in alertas[:20]:
            year_str = f"{row['year']:>4}"
            month_str = f"{row['month']:>2}"
            center_str = row['center'][:20].ljust(20)
            element_str = row['element'][:15].ljust(15)
            tipo_str = row['tipo'][:12].ljust(12)
            abs_str = row['var_abs'][:15].rjust(15)
            pct_str = row['var_pct'][:12].rjust(12)
            lines.append(f"{year_str} | {month_str} | {center_str} | {element_str} | {tipo_str} | {abs_str} | {pct_str}")
        lines.append("-" * 76)
        lines.append("[*Nota: Solo se muestran desviaciones que impactan críticamente la rentabilidad (umbral >10%)]")

        lines.append("\n2. PROPUESTAS DE ACCIÓN E IMPACTO ESTRATÉGICO (POR COMPORTAMIENTO)")
        lines.append("-" * 76)
        if propuestas:
            for (center, element, categoria), registros in propuestas.items():
                meses = sorted([(r['year'], r['month']) for r in registros])
                meses_str = ", ".join([f"{y}-M{m:02d}" for y, m in meses])
                avg_abs = sum(r['var_abs_num'] for r in registros) / len(registros)
                avg_pct = sum(r['var_pct_num'] for r in registros) / len(registros)

                if categoria == "Ingresos":
                    icono = "▲ CRÍTICO EN INGRESOS"
                    diagnostico = f"El ingreso real está por debajo del presupuesto en un promedio de {avg_pct:+.1f}% en {len(registros)} período(s)."
                    propuesta = "Reevaluar las estrategias comerciales, revisar precios y activar un plan de contingencia de ventas."
                else:
                    icono = "▼ CONTROL DE COSTOS"
                    diagnostico = f"Los costos/gastos reales superan el presupuesto en un promedio de {avg_pct:+.1f}% en {len(registros)} período(s)."
                    propuesta = "Auditar la asignación presupuestal, identificar ineficiencias y congelar gastos no esenciales en el centro."

                lines.append(f"\n[{icono}] {center} - {element} (Tendencia Detectada en {meses_str}):")
                lines.append(f"  • Diagnóstico: {diagnostico}")
                lines.append(f"  • Propuesta: {propuesta}")
                if len(registros) > 3:
                    lines.append(f"  • Recomendación adicional: Considerar una revisión estructural del centro {center}.")
        else:
            lines.append(" No se generaron propuestas específicas.")

        lines.append("\n" + "=" * 76)
        return lines


# ---------------------- ENVÍO DE CORREOS ----------------------
class EmailSender:
    def __init__(self, smtp_server, smtp_port, username, password):
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        self.username = username
        self.password = password

    def send_budget_report(self, to_email, subject, body, attachment_path=None):
        try:
            msg = MIMEMultipart()
            msg['From'] = self.username
            msg['To'] = to_email
            msg['Subject'] = subject
            msg.attach(MIMEText(body, 'plain'))
            if attachment_path and os.path.exists(attachment_path):
                with open(attachment_path, 'rb') as f:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(f.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(attachment_path)}')
                    msg.attach(part)
            server = smtplib.SMTP(self.smtp_server, self.smtp_port)
            server.starttls()
            server.login(self.username, self.password)
            server.send_message(msg)
            server.quit()
            return True
        except Exception as e:
            print(f"Error enviando correo: {e}")
            return False


# ---------------------- INTERFAZ GRÁFICA (MainApp) ----------------------
class MainApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1400x900")
        self.minsize(1200, 700)

        self.db = Database()
        self.model = BudgetModel(self.db)

        self.tabview = ctk.CTkTabview(self)
        self.tabview.pack(fill="both", expand=True, padx=10, pady=10)

        # Pestañas principales (eliminada "Importación")
        self.tabview.add("Configuración")
        self.tabview.add("Presupuesto Maestro")
        self.tabview.add("Dashboard")
        self.tabview.add("Reportes")
        self.tabview.add("Análisis")
        self.tabview.add("Correo")
        self.tabview.add("Optimización")

        self.preview_df = None
        self.analysis_canvas_frame = None

        # Configuración de cada pestaña
        self.setup_config_tab()  # Ahora incluye la importación
        self.setup_budget_tab()
        self.setup_dashboard_tab()
        self.setup_reports_tab()
        self.setup_analysis_tab()
        self.setup_email_tab()
        self.setup_optimization_tab()

        self.init_sample_data()

    def init_sample_data(self):
        if not self.model.get_periods():
            period_id = self.model.add_period("2026-2028", "2026-01-01", "2028-12-31")
            version_id = self.model.add_version("Versión 1", "Presupuesto inicial")
            self.model.add_center("PROD1", "Planta Producción", "Productivo", "prod@example.com")
            self.model.add_center("ADMIN", "Administración", "Gasto", "admin@example.com")
            self.model.add_center("VENTAS", "Ventas", "Gasto", "ventas@example.com")
            self.model.add_element("VTA", "Ventas", "Ingresos")
            self.model.add_element("COSTO", "Costo de Ventas", "Costo")
            self.model.add_element("GADMIN", "Gastos Administrativos", "Gasto")
            self.model.add_element("GVENTAS", "Gastos de Ventas", "Gasto")
            self.model.add_element("INV", "Inventario", "Activo")
            self.model.set_forecast_cut(period_id, 4)
            engine = BudgetEngine(self.model)
            engine.run_cascade(period_id, version_id)
            messagebox.showinfo("Datos de ejemplo", "Se han creado datos de ejemplo para comenzar.")

    # ---------- Configuración de pestañas (UI) ----------
    def setup_config_tab(self):
        tab = self.tabview.tab("Configuración")
        main_frame = ctk.CTkFrame(tab)
        main_frame.pack(fill="both", expand=True, padx=5, pady=5)
        main_frame.grid_rowconfigure(1, weight=1)  # La fila de configuración se expande
        main_frame.grid_columnconfigure(0, weight=1)

        # ===== SECCIÓN DE IMPORTACIÓN (Paso 1) =====
        import_frame = ctk.CTkFrame(main_frame)
        import_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        import_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(import_frame, text="Paso 1: Cargar Datos Presupuestales (.xlsx)",
                     font=("Arial", 14, "bold")).grid(row=0, column=0, columnspan=3, sticky="w", padx=5, pady=(5, 0))

        # Botón de importación unificada
        btn_import_unified = ctk.CTkButton(import_frame, text="Importar Excel Unificado",
                                           command=self.import_unified_excel)
        btn_import_unified.grid(row=1, column=0, padx=5, pady=5, sticky="w")

        # Etiqueta de estado (resultado de la importación)
        self.import_status = ctk.CTkLabel(import_frame, text="", wraplength=600)
        self.import_status.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        # Área de vista previa (se mantiene la funcionalidad)
        self.preview_textbox = ctk.CTkTextbox(import_frame, wrap="none", font=("Courier", 13), height=150)
        self.preview_textbox.grid(row=2, column=0, columnspan=3, sticky="ew", padx=5, pady=5)
        self.preview_textbox.insert("1.0", "La vista previa aparecerá aquí después de importar.")

        # Separador visual
        separator = ctk.CTkFrame(main_frame, height=2, fg_color="gray")
        separator.grid(row=1, column=0, sticky="ew", padx=5, pady=5)

        # ===== SECCIÓN DE CONFIGURACIÓN (original) =====
        config_frame = ctk.CTkFrame(main_frame)
        config_frame.grid(row=2, column=0, sticky="nsew", padx=5, pady=5)
        config_frame.grid_rowconfigure(0, weight=1)
        config_frame.grid_columnconfigure(0, weight=0, minsize=280)
        config_frame.grid_columnconfigure(1, weight=1)

        # Sub-frame izquierdo (controles)
        left_frame = ctk.CTkFrame(config_frame)
        left_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        left_frame.grid_rowconfigure(6, weight=1)

        ctk.CTkLabel(left_frame, text="Gestión de Períodos", font=("Arial", 14, "bold")).grid(row=0, column=0, columnspan=2, pady=(0, 5), sticky="w")
        btn_add_period = ctk.CTkButton(left_frame, text="Agregar Período", command=self.add_period_dialog)
        btn_add_period.grid(row=1, column=0, padx=5, pady=5, sticky="ew")
        btn_refresh_periods = ctk.CTkButton(left_frame, text="Refrescar", command=self.refresh_periods)
        btn_refresh_periods.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        ctk.CTkLabel(left_frame, text="Versiones", font=("Arial", 14, "bold")).grid(row=2, column=0, columnspan=2, pady=(15, 5), sticky="w")
        btn_add_version = ctk.CTkButton(left_frame, text="Agregar Versión", command=self.add_version_dialog)
        btn_add_version.grid(row=3, column=0, padx=5, pady=5, sticky="ew")
        btn_refresh_versions = ctk.CTkButton(left_frame, text="Refrescar", command=self.refresh_versions)
        btn_refresh_versions.grid(row=3, column=1, padx=5, pady=5, sticky="ew")

        ctk.CTkLabel(left_frame, text="Configuración Forecast", font=("Arial", 14, "bold")).grid(row=4, column=0, columnspan=2, pady=(15, 5), sticky="w")
        ctk.CTkLabel(left_frame, text="Mes de Corte:").grid(row=5, column=0, sticky="e", padx=5)
        self.cut_month_entry = ctk.CTkEntry(left_frame, width=60)
        self.cut_month_entry.grid(row=5, column=1, sticky="w", padx=5)
        btn_set_cut = ctk.CTkButton(left_frame, text="Establecer Corte", command=self.set_forecast_cut)
        btn_set_cut.grid(row=6, column=0, columnspan=2, pady=10, sticky="ew")

        # Sub-frame derecho (listas)
        right_frame = ctk.CTkFrame(config_frame)
        right_frame.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)
        right_frame.grid_rowconfigure(0, weight=1)
        right_frame.grid_rowconfigure(1, weight=1)
        right_frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(right_frame, text="Períodos", font=("Arial", 14, "bold")).grid(row=0, column=0, sticky="w")
        self.period_listbox = tk.Listbox(right_frame, height=8, bg='white', fg='black', font=("Arial", 13))
        self.period_listbox.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)

        ctk.CTkLabel(right_frame, text="Versiones", font=("Arial", 14, "bold")).grid(row=1, column=0, sticky="w", pady=(10, 0))
        self.version_listbox = tk.Listbox(right_frame, height=8, bg='white', fg='black', font=("Arial", 13))
        self.version_listbox.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)

        self.refresh_periods()
        self.refresh_versions()

    # ---------- Métodos auxiliares de configuración ----------
    def refresh_periods(self):
        self.period_listbox.delete(0, "end")
        periods = self.model.get_periods()
        for p in periods:
            self.period_listbox.insert("end", f"{p[1]} ({p[2]} - {p[3]})")

    def refresh_versions(self):
        self.version_listbox.delete(0, "end")
        versions = self.model.get_versions()
        for v in versions:
            self.version_listbox.insert("end", f"{v[1]} - {v[2]}")

    def add_period_dialog(self):
        dialog = ctk.CTkToplevel(self)
        dialog.title("Nuevo Período")
        dialog.geometry("300x200")
        ctk.CTkLabel(dialog, text="Nombre:").grid(row=0, column=0, padx=5, pady=5)
        name_entry = ctk.CTkEntry(dialog)
        name_entry.grid(row=0, column=1, padx=5, pady=5)
        ctk.CTkLabel(dialog, text="Fecha Inicio (YYYY-MM-DD):").grid(row=1, column=0, padx=5, pady=5)
        start_entry = ctk.CTkEntry(dialog)
        start_entry.grid(row=1, column=1, padx=5, pady=5)
        ctk.CTkLabel(dialog, text="Fecha Fin (YYYY-MM-DD):").grid(row=2, column=0, padx=5, pady=5)
        end_entry = ctk.CTkEntry(dialog)
        end_entry.grid(row=2, column=1, padx=5, pady=5)

        def save_period():
            name = name_entry.get()
            start = start_entry.get()
            end = end_entry.get()
            if name and start and end:
                try:
                    self.model.add_period(name, start, end)
                    self.refresh_periods()
                    dialog.destroy()
                except Exception as e:
                    messagebox.showerror("Error", f"No se pudo guardar: {e}")
            else:
                messagebox.showwarning("Advertencia", "Complete todos los campos.")
        ctk.CTkButton(dialog, text="Guardar", command=save_period).grid(row=3, column=0, columnspan=2, pady=10)

    def add_version_dialog(self):
        dialog = ctk.CTkToplevel(self)
        dialog.title("Nueva Versión")
        dialog.geometry("300x150")
        ctk.CTkLabel(dialog, text="Nombre:").grid(row=0, column=0, padx=5, pady=5)
        name_entry = ctk.CTkEntry(dialog)
        name_entry.grid(row=0, column=1, padx=5, pady=5)
        ctk.CTkLabel(dialog, text="Descripción:").grid(row=1, column=0, padx=5, pady=5)
        desc_entry = ctk.CTkEntry(dialog)
        desc_entry.grid(row=1, column=1, padx=5, pady=5)

        def save_version():
            name = name_entry.get()
            desc = desc_entry.get()
            if name:
                try:
                    self.model.add_version(name, desc)
                    self.refresh_versions()
                    dialog.destroy()
                except Exception as e:
                    messagebox.showerror("Error", f"No se pudo guardar: {e}")
            else:
                messagebox.showwarning("Advertencia", "Ingrese un nombre.")
        ctk.CTkButton(dialog, text="Guardar", command=save_version).grid(row=2, column=0, columnspan=2, pady=10)

    def set_forecast_cut(self):
        sel = self.period_listbox.curselection()
        if not sel:
            messagebox.showwarning("Advertencia", "Seleccione un período.")
            return
        period_name = self.period_listbox.get(sel[0])
        period_id = self.model.db.fetch_one("SELECT id FROM periods WHERE name=?", (period_name.split(' (')[0],))[0]
        cut = self.cut_month_entry.get()
        if cut.isdigit() and 1 <= int(cut) <= 12:
            self.model.set_forecast_cut(period_id, int(cut))
            messagebox.showinfo("Éxito", f"Mes de corte establecido a {cut}.")
        else:
            messagebox.showerror("Error", "Ingrese un mes válido (1-12).")

    # ---------- Funciones de importación ----------
    def import_unified_excel(self):
        filepath = filedialog.askopenfilename(
            title="Seleccionar archivo Excel unificado",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not filepath:
            return
        importer = DataImporter(self.db.conn, self.model)
        success, msg = importer.import_from_unified_excel(filepath)
        if success:
            messagebox.showinfo("Éxito", msg)
            self.import_status.configure(text=f"✅ {msg}")
            # Mostrar vista previa del primer DataFrame (ventas como ejemplo)
            try:
                df = pd.read_excel(filepath, sheet_name="Ventas", engine='openpyxl', data_only=True)
                self.show_preview(df)
            except:
                pass
        else:
            messagebox.showerror("Error", msg)
            self.import_status.configure(text=f"❌ {msg}")

    def import_excel(self, table_type):
        filepath = filedialog.askopenfilename(title="Seleccionar archivo Excel", filetypes=[("Excel files", "*.xlsx *.xls")])
        if not filepath:
            return
        try:
            df = pd.read_excel(filepath, header=0)
            if table_type == 'elements':
                col_code = next((c for c in df.columns if c.lower() in ['code', 'codigo', 'código']), None)
                col_name = next((c for c in df.columns if c.lower() in ['name', 'nombre', 'descripcion']), None)
                col_cat = next((c for c in df.columns if c.lower() in ['category', 'categoria', 'tipo']), None)
                if col_code is None or col_name is None or col_cat is None:
                    raise ValueError("Columnas esperadas: code, name, category")
                for _, row in df.iterrows():
                    self.model.add_element(str(row[col_code]), str(row[col_name]), str(row[col_cat]))
                msg = "Plan de cuentas importado."
            elif table_type == 'centers':
                col_code = next((c for c in df.columns if c.lower() in ['code', 'codigo']), None)
                col_name = next((c for c in df.columns if c.lower() in ['name', 'nombre']), None)
                col_type = next((c for c in df.columns if c.lower() in ['type', 'tipo']), None)
                col_email = next((c for c in df.columns if c.lower() in ['manager_email', 'email', 'correo']), None)
                if col_code is None or col_name is None or col_type is None:
                    raise ValueError("Columnas esperadas: code, name, type, manager_email (opcional)")
                for _, row in df.iterrows():
                    email = str(row[col_email]) if col_email and pd.notna(row[col_email]) else ""
                    self.model.add_center(str(row[col_code]), str(row[col_name]), str(row[col_type]), email)
                msg = "Centros importados."
            elif table_type == 'products':
                col_code = next((c for c in df.columns if c.lower() in ['code', 'codigo']), None)
                col_name = next((c for c in df.columns if c.lower() in ['name', 'nombre', 'descripcion']), None)
                col_unit = next((c for c in df.columns if c.lower() in ['unit', 'unidad']), None)
                col_price = next((c for c in df.columns if c.lower() in ['price', 'precio']), None)
                col_cost = next((c for c in df.columns if c.lower() in ['standard_cost', 'costo_estandar']), None)
                if col_code is None or col_name is None:
                    raise ValueError("Columnas mínimas: code, name")
                for _, row in df.iterrows():
                    price = float(row[col_price]) if col_price and pd.notna(row[col_price]) else 0.0
                    std_cost = float(row[col_cost]) if col_cost and pd.notna(row[col_cost]) else 0.0
                    self.model.add_product(str(row[col_code]), str(row[col_name]),
                                           str(row[col_unit]) if col_unit and pd.notna(row[col_unit]) else "",
                                           price, std_cost)
                msg = "Productos importados."
            elif table_type == 'drivers':
                col_code = next((c for c in df.columns if c.lower() in ['code', 'codigo']), None)
                col_desc = next((c for c in df.columns if c.lower() in ['description', 'descripcion']), None)
                col_value = next((c for c in df.columns if c.lower() in ['value', 'valor']), None)
                if col_code is None or col_desc is None or col_value is None:
                    raise ValueError("Columnas esperadas: code, description, value")
                for _, row in df.iterrows():
                    self.model.add_driver(str(row[col_code]), str(row[col_desc]), float(row[col_value]))
                msg = "Drivers importados."
            else:
                msg = "Importación no implementada."

            self.preview_df = df
            self.show_preview(df)
            self.import_status.configure(text=f"✅ {msg}")
        except Exception as e:
            messagebox.showerror("Error", f"Error al importar: {e}")
            self.import_status.configure(text=f"❌ Error: {str(e)}")
            self.preview_textbox.delete("1.0", "end")
            self.preview_textbox.insert("1.0", f"Error al procesar el archivo:\n{str(e)}")

    def show_preview(self, df):
        self.preview_textbox.delete("1.0", "end")
        if df is None or df.empty:
            self.preview_textbox.insert("1.0", "No hay datos para mostrar.")
            return
        preview_df = df.head(20)
        cols = preview_df.columns.tolist()
        col_widths = {c: max(len(str(c)), preview_df[c].astype(str).map(len).max()) for c in cols}
        header = " | ".join([f"{c:<{col_widths[c]}}" for c in cols])
        self.preview_textbox.insert("end", header + "\n")
        self.preview_textbox.insert("end", "-" * len(header) + "\n")
        for _, row in preview_df.iterrows():
            line = " | ".join([f"{str(row[c]):<{col_widths[c]}}" for c in cols])
            self.preview_textbox.insert("end", line + "\n")
        self.preview_textbox.insert("end", f"\n... Mostrando {len(preview_df)} de {len(df)} filas.\n")

    # ---------- Pestaña Presupuesto Maestro ----------
    def setup_budget_tab(self):
        tab = self.tabview.tab("Presupuesto Maestro")
        main_frame = ctk.CTkFrame(tab)
        main_frame.pack(fill="both", expand=True, padx=5, pady=5)
        main_frame.grid_rowconfigure(0, weight=1)
        main_frame.grid_columnconfigure(0, weight=0, minsize=280)
        main_frame.grid_columnconfigure(1, weight=1)

        left_frame = ctk.CTkFrame(main_frame)
        left_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        ctk.CTkLabel(left_frame, text="Ejecutar Cascada", font=("Arial", 14, "bold")).pack(pady=(0, 10))

        ctk.CTkLabel(left_frame, text="Período:").pack(anchor="w")
        self.budget_period_combobox = ctk.CTkComboBox(left_frame, values=[], state="readonly")
        self.budget_period_combobox.pack(pady=5, fill="x")

        ctk.CTkLabel(left_frame, text="Versión:").pack(anchor="w", pady=(10, 0))
        self.budget_version_combobox = ctk.CTkComboBox(left_frame, values=[], state="readonly")
        self.budget_version_combobox.pack(pady=5, fill="x")

        btn_run = ctk.CTkButton(left_frame, text="Ejecutar Presupuesto Maestro", command=self.run_budget)
        btn_run.pack(pady=15, fill="x")

        ctk.CTkLabel(left_frame, text="Ver Detalle de:", font=("Arial", 14, "bold")).pack(pady=(20, 5))
        self.detail_menu = ctk.CTkOptionMenu(left_frame,
                                             values=["Presupuesto de Producción",
                                                     "Costos por Centro Productivo",
                                                     "Gastos por Centro de Actividad",
                                                     "Presupuesto de Inversión (CAPEX)",
                                                     "Estado de Situación Financiera",
                                                     "Estado de Resultados"],
                                             command=self.show_budget_detail)
        self.detail_menu.pack(pady=5, fill="x")
        self.detail_menu.set("Presupuesto de Producción")

        right_frame = ctk.CTkFrame(main_frame)
        right_frame.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)
        ctk.CTkLabel(right_frame, text="Detalle", font=("Arial", 14, "bold")).pack(pady=(0, 5))
        self.detail_textbox = ctk.CTkTextbox(right_frame, wrap="none", font=("Courier", 13))
        self.detail_textbox.pack(fill="both", expand=True, padx=5, pady=5)

        self.refresh_budget_combos()

    def refresh_budget_combos(self):
        periods = self.model.get_periods()
        self.budget_period_combobox.configure(values=[p[1] for p in periods])
        if periods:
            self.budget_period_combobox.set(periods[0][1])
        versions = self.model.get_versions()
        self.budget_version_combobox.configure(values=[v[1] for v in versions])
        if versions:
            self.budget_version_combobox.set(versions[0][1])

    def run_budget(self):
        period_name = self.budget_period_combobox.get()
        version_name = self.budget_version_combobox.get()
        if not period_name or not version_name:
            messagebox.showwarning("Advertencia", "Seleccione período y versión.")
            return
        period = self.model.db.fetch_one("SELECT id FROM periods WHERE name=?", (period_name,))
        version = self.model.db.fetch_one("SELECT id FROM versions WHERE name=?", (version_name,))
        if not period or not version:
            messagebox.showerror("Error", "Período o versión no encontrados.")
            return
        engine = BudgetEngine(self.model)
        try:
            engine.run_cascade(period[0], version[0])
            self.show_budget_detail(self.detail_menu.get())
            messagebox.showinfo("Éxito", "Cascada presupuestal completada.")
        except Exception as e:
            messagebox.showerror("Error", f"Error al ejecutar: {e}")

    def show_budget_detail(self, choice):
        period_name = self.budget_period_combobox.get()
        version_name = self.budget_version_combobox.get()
        if not period_name or not version_name:
            self.detail_textbox.delete("1.0", "end")
            self.detail_textbox.insert("1.0", "Seleccione un período y versión.")
            return
        period = self.model.db.fetch_one("SELECT id FROM periods WHERE name=?", (period_name,))
        version = self.model.db.fetch_one("SELECT id FROM versions WHERE name=?", (version_name,))
        if not period or not version:
            self.detail_textbox.delete("1.0", "end")
            self.detail_textbox.insert("1.0", "Período o versión no válidos.")
            return

        self.detail_textbox.delete("1.0", "end")
        if choice == "Presupuesto de Producción":
            self.show_production_budget(period[0], version[0])
        elif choice == "Costos por Centro Productivo":
            self.show_cost_by_center(period[0], version[0], "Productivo")
        elif choice == "Gastos por Centro de Actividad":
            self.show_cost_by_center(period[0], version[0], "Gasto")
        elif choice == "Presupuesto de Inversión (CAPEX)":
            self.show_investment(period[0])
        elif choice == "Estado de Situación Financiera":
            self.show_financial_statement(period[0], version[0], "BS")
        elif choice == "Estado de Resultados":
            self.show_financial_statement(period[0], version[0], "IS")

    def _format_dataframe_as_table(self, df, title=""):
        if df is None or df.empty:
            return "No hay datos para mostrar."
        lines = []
        if title:
            lines.append(title)
            lines.append("=" * len(title))
        df_str = df.astype(str)
        col_widths = {c: max(len(str(c)), df_str[c].map(len).max()) for c in df_str.columns}
        header = " | ".join([f"{c:<{col_widths[c]}}" for c in df_str.columns])
        lines.append(header)
        lines.append("-" * len(header))
        for _, row in df_str.iterrows():
            line = " | ".join([f"{row[c]:<{col_widths[c]}}" for c in df_str.columns])
            lines.append(line)
        return "\n".join(lines)

    def show_production_budget(self, period_id, version_id):
        data = self.model.get_budget_data(period_id, version_id)
        if not data:
            self.detail_textbox.insert("1.0", "No hay datos de presupuesto.")
            return
        df = pd.DataFrame(data,
                          columns=['id', 'period_id', 'version_id', 'center_id', 'element_id', 'month', 'year',
                                   'budget_original', 'actual_original', 'budget_revised', 'actual_revised',
                                   'center_name', 'element_name'])
        centers = self.model.get_centers()
        prod_center_ids = [c[0] for c in centers if c[3] == "Productivo"]
        df_prod = df[df['center_id'].isin(prod_center_ids)]
        if df_prod.empty:
            self.detail_textbox.insert("1.0", "No hay datos para centros productivos.")
            return
        grouped = df_prod.groupby(['year', 'element_name'])['budget_original'].sum().reset_index()
        pivoted = grouped.pivot(index='element_name', columns='year', values='budget_original').fillna(0)
        pivoted['Total'] = pivoted.sum(axis=1)
        table = self._format_dataframe_as_table(pivoted.reset_index().round(2), "PRESUPUESTO DE PRODUCCIÓN (por elemento y año)")
        self.detail_textbox.insert("1.0", table)

    def show_cost_by_center(self, period_id, version_id, center_type):
        data = self.model.get_budget_data(period_id, version_id)
        if not data:
            self.detail_textbox.insert("1.0", "No hay datos de presupuesto.")
            return
        df = pd.DataFrame(data,
                          columns=['id', 'period_id', 'version_id', 'center_id', 'element_id', 'month', 'year',
                                   'budget_original', 'actual_original', 'budget_revised', 'actual_revised',
                                   'center_name', 'element_name'])
        centers = self.model.get_centers()
        center_ids = [c[0] for c in centers if c[3] == center_type]
        df_filtered = df[df['center_id'].isin(center_ids)]
        if df_filtered.empty:
            self.detail_textbox.insert("1.0", f"No hay datos para centros tipo {center_type}.")
            return
        grouped = df_filtered.groupby(['center_name', 'year'])['budget_original'].sum().reset_index()
        pivoted = grouped.pivot(index='center_name', columns='year', values='budget_original').fillna(0)
        pivoted['Total'] = pivoted.sum(axis=1)
        title = "COSTOS POR CENTRO PRODUCTIVO" if center_type == "Productivo" else "GASTOS POR CENTRO DE ACTIVIDAD"
        table = self._format_dataframe_as_table(pivoted.reset_index().round(2), title)
        self.detail_textbox.insert("1.0", table)

    def show_investment(self, period_id):
        investments = self.model.get_investments(period_id)
        if not investments:
            self.detail_textbox.insert("1.0", "No hay inversiones registradas.")
            return
        df = pd.DataFrame(investments, columns=['id', 'period_id', 'description', 'amount', 'year'])
        df = df[['year', 'description', 'amount']]
        table = self._format_dataframe_as_table(df, "PRESUPUESTO DE INVERSIÓN (CAPEX)")
        self.detail_textbox.insert("1.0", table)

    def show_financial_statement(self, period_id, version_id, stype):
        stmts = self.model.get_financial_statements(period_id, version_id, stype)
        if not stmts:
            self.detail_textbox.insert("1.0", f"No hay datos para {stype}.")
            return
        df = pd.DataFrame(stmts,
                          columns=['id', 'period_id', 'version_id', 'statement_type', 'account_code', 'account_name', 'amount', 'year'])
        pivoted = df.pivot(index='account_name', columns='year', values='amount').fillna(0)
        pivoted['Total'] = pivoted.sum(axis=1)
        title = "ESTADO DE SITUACIÓN FINANCIERA" if stype == "BS" else "ESTADO DE RESULTADOS"
        table = self._format_dataframe_as_table(pivoted.reset_index().round(2), title)
        self.detail_textbox.insert("1.0", table)

    # ---------- Pestaña Dashboard ----------
    def setup_dashboard_tab(self):
        tab = self.tabview.tab("Dashboard")
        main_frame = ctk.CTkFrame(tab)
        main_frame.pack(fill="both", expand=True, padx=5, pady=5)
        main_frame.grid_rowconfigure(0, weight=1)
        main_frame.grid_columnconfigure(0, weight=0, minsize=280)
        main_frame.grid_columnconfigure(1, weight=1)

        left_frame = ctk.CTkFrame(main_frame)
        left_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        ctk.CTkLabel(left_frame, text="Controles", font=("Arial", 14, "bold")).pack(pady=(0, 10))

        ctk.CTkLabel(left_frame, text="Período:").pack(anchor="w")
        self.dash_period_combobox = ctk.CTkComboBox(left_frame, values=[], state="readonly")
        self.dash_period_combobox.pack(pady=5, fill="x")

        ctk.CTkLabel(left_frame, text="Versión:").pack(anchor="w", pady=(10, 0))
        self.dash_version_combobox = ctk.CTkComboBox(left_frame, values=[], state="readonly")
        self.dash_version_combobox.pack(pady=5, fill="x")

        btn_refresh_dash = ctk.CTkButton(left_frame, text="Actualizar Gráficos", command=self.refresh_dashboard)
        btn_refresh_dash.pack(pady=15, fill="x")

        right_frame = ctk.CTkFrame(main_frame)
        right_frame.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)
        self.dash_canvas_frame = ctk.CTkFrame(right_frame)
        self.dash_canvas_frame.pack(fill="both", expand=True)

        self.refresh_dash_combos()

    def refresh_dash_combos(self):
        periods = self.model.get_periods()
        self.dash_period_combobox.configure(values=[p[1] for p in periods])
        if periods:
            self.dash_period_combobox.set(periods[0][1])
        versions = self.model.get_versions()
        self.dash_version_combobox.configure(values=[v[1] for v in versions])
        if versions:
            self.dash_version_combobox.set(versions[0][1])

    def refresh_dashboard(self):
        period_name = self.dash_period_combobox.get()
        version_name = self.dash_version_combobox.get()
        if not period_name or not version_name:
            return
        period = self.model.db.fetch_one("SELECT id FROM periods WHERE name=?", (period_name,))
        version = self.model.db.fetch_one("SELECT id FROM versions WHERE name=?", (version_name,))
        if not period or not version:
            return
        data = self.model.get_budget_data(period[0], version[0])
        if not data:
            messagebox.showinfo("Info", "No hay datos para graficar.")
            return
        df = pd.DataFrame(data,
                          columns=['id', 'period_id', 'version_id', 'center_id', 'element_id', 'month', 'year',
                                   'budget_original', 'actual_original', 'budget_revised', 'actual_revised',
                                   'center_name', 'element_name'])
        grouped = df.groupby('element_name')['actual_original'].sum().sort_values(ascending=False)
        for widget in self.dash_canvas_frame.winfo_children():
            widget.destroy()

        fig, ax = plt.subplots(figsize=(10, 6))
        ax.bar(grouped.index, grouped.values, color='skyblue')
        ax.set_title('Ejecución Presupuestal por Elemento (Real Original)', fontsize=14)
        ax.set_xlabel('Elemento', fontsize=12)
        ax.set_ylabel('Monto', fontsize=12)
        plt.xticks(rotation=45, ha='right')
        plt.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=self.dash_canvas_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

        toolbar = NavigationToolbar2Tk(canvas, self.dash_canvas_frame)
        toolbar.update()
        canvas.get_tk_widget().pack(fill="both", expand=True)

    # ---------- Pestaña Reportes ----------
    def setup_reports_tab(self):
        tab = self.tabview.tab("Reportes")
        main_frame = ctk.CTkFrame(tab)
        main_frame.pack(fill="both", expand=True, padx=5, pady=5)
        main_frame.grid_rowconfigure(0, weight=1)
        main_frame.grid_columnconfigure(0, weight=0, minsize=280)
        main_frame.grid_columnconfigure(1, weight=1)

        left_frame = ctk.CTkFrame(main_frame)
        left_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        ctk.CTkLabel(left_frame, text="Exportar Reportes", font=("Arial", 14, "bold")).pack(pady=(0, 10))

        ctk.CTkLabel(left_frame, text="Período:").pack(anchor="w")
        self.report_period_combobox = ctk.CTkComboBox(left_frame, values=[], state="readonly")
        self.report_period_combobox.pack(pady=5, fill="x")

        ctk.CTkLabel(left_frame, text="Versión:").pack(anchor="w", pady=(10, 0))
        self.report_version_combobox = ctk.CTkComboBox(left_frame, values=[], state="readonly")
        self.report_version_combobox.pack(pady=5, fill="x")

        btn_preview = ctk.CTkButton(left_frame, text="Actualizar Vista Previa", command=self.update_report_preview)
        btn_preview.pack(pady=10, fill="x")

        btn_export_excel = ctk.CTkButton(left_frame, text="Exportar a Excel", command=self.export_excel_report)
        btn_export_excel.pack(pady=5, fill="x")

        btn_export_pdf = ctk.CTkButton(left_frame, text="Exportar a PDF", command=self.export_pdf_report)
        btn_export_pdf.pack(pady=5, fill="x")

        right_frame = ctk.CTkFrame(main_frame)
        right_frame.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)
        ctk.CTkLabel(right_frame, text="Vista Previa del Reporte", font=("Arial", 14, "bold")).pack(pady=(0, 5))
        self.report_preview_textbox = ctk.CTkTextbox(right_frame, wrap="none", font=("Courier", 13))
        self.report_preview_textbox.pack(fill="both", expand=True, padx=5, pady=5)

        self.report_status = ctk.CTkLabel(right_frame, text="", wraplength=600)
        self.report_status.pack(pady=10)

        self.refresh_report_combos()
        self.update_report_preview()

    def refresh_report_combos(self):
        periods = self.model.get_periods()
        self.report_period_combobox.configure(values=[p[1] for p in periods])
        if periods:
            self.report_period_combobox.set(periods[0][1])
        versions = self.model.get_versions()
        self.report_version_combobox.configure(values=[v[1] for v in versions])
        if versions:
            self.report_version_combobox.set(versions[0][1])

    def update_report_preview(self):
        period_name = self.report_period_combobox.get()
        version_name = self.report_version_combobox.get()
        if not period_name or not version_name:
            self.report_preview_textbox.delete("1.0", "end")
            self.report_preview_textbox.insert("1.0", "Seleccione período y versión.")
            return
        period = self.model.db.fetch_one("SELECT id FROM periods WHERE name=?", (period_name,))
        version = self.model.db.fetch_one("SELECT id FROM versions WHERE name=?", (version_name,))
        if not period or not version:
            self.report_preview_textbox.delete("1.0", "end")
            self.report_preview_textbox.insert("1.0", "Período o versión no válidos.")
            return

        data = self.model.get_budget_data(period[0], version[0])
        if not data:
            self.report_preview_textbox.delete("1.0", "end")
            self.report_preview_textbox.insert("1.0", "No hay datos para este período/versión.")
            return

        df = pd.DataFrame(data,
                          columns=['id', 'period_id', 'version_id', 'center_id', 'element_id', 'month', 'year',
                                   'budget_original', 'actual_original', 'budget_revised', 'actual_revised',
                                   'center_name', 'element_name'])
        preview_df = df[['center_name', 'element_name', 'year', 'month',
                         'budget_original', 'actual_original']].head(50)
        cols = preview_df.columns.tolist()
        col_widths = {c: max(len(str(c)), preview_df[c].astype(str).map(len).max()) for c in cols}
        lines = []
        lines.append("VISTA PREVIA DEL REPORTE (primeras 50 filas)")
        lines.append("=" * 80)
        header = " | ".join([f"{c:<{col_widths[c]}}" for c in cols])
        lines.append(header)
        lines.append("-" * len(header))
        for _, row in preview_df.iterrows():
            line = " | ".join([f"{str(row[c]):<{col_widths[c]}}" for c in cols])
            lines.append(line)
        lines.append(f"\n... Mostrando 50 de {len(df)} filas.")
        self.report_preview_textbox.delete("1.0", "end")
        self.report_preview_textbox.insert("1.0", "\n".join(lines))

    def export_excel_report(self):
        period_name = self.report_period_combobox.get()
        version_name = self.report_version_combobox.get()
        if not period_name or not version_name:
            messagebox.showwarning("Advertencia", "Seleccione período y versión.")
            return
        period = self.model.db.fetch_one("SELECT id FROM periods WHERE name=?", (period_name,))
        version = self.model.db.fetch_one("SELECT id FROM versions WHERE name=?", (version_name,))
        if not period or not version:
            return
        filepath = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not filepath:
            return
        exporter = ExcelExporter(self.model)
        try:
            exporter.export_full_budget(period[0], version[0], filepath)
            self.report_status.configure(text=f"✅ Reporte exportado a {filepath}")
        except Exception as e:
            messagebox.showerror("Error", f"Error al exportar: {e}")
            self.report_status.configure(text=f"❌ {str(e)}")

    def export_pdf_report(self):
        period_name = self.report_period_combobox.get()
        version_name = self.report_version_combobox.get()
        if not period_name or not version_name:
            messagebox.showwarning("Advertencia", "Seleccione período y versión.")
            return
        period = self.model.db.fetch_one("SELECT id FROM periods WHERE name=?", (period_name,))
        version = self.model.db.fetch_one("SELECT id FROM versions WHERE name=?", (version_name,))
        if not period or not version:
            return
        filepath = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if not filepath:
            return
        exporter = PDFExporter(self.model)
        exporter.export_report(period[0], version[0], filepath)

    # ---------- Pestaña Análisis ----------
    def setup_analysis_tab(self):
        tab = self.tabview.tab("Análisis")
        main_frame = ctk.CTkFrame(tab)
        main_frame.pack(fill="both", expand=True, padx=5, pady=5)
        main_frame.grid_rowconfigure(0, weight=1)
        main_frame.grid_columnconfigure(0, weight=0, minsize=280)
        main_frame.grid_columnconfigure(1, weight=1)

        left_frame = ctk.CTkFrame(main_frame)
        left_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        ctk.CTkLabel(left_frame, text="Parámetros", font=("Arial", 14, "bold")).pack(pady=(0, 10))

        ctk.CTkLabel(left_frame, text="Período:").pack(anchor="w")
        self.analysis_period_combobox = ctk.CTkComboBox(left_frame, values=[], state="readonly")
        self.analysis_period_combobox.pack(pady=5, fill="x")

        ctk.CTkLabel(left_frame, text="Versión:").pack(anchor="w", pady=(10, 0))
        self.analysis_version_combobox = ctk.CTkComboBox(left_frame, values=[], state="readonly")
        self.analysis_version_combobox.pack(pady=5, fill="x")

        ctk.CTkLabel(left_frame, text="Año:").pack(anchor="w", pady=(10, 0))
        self.analysis_year_combobox = ctk.CTkComboBox(left_frame, values=['2026', '2027', '2028'], state="readonly")
        self.analysis_year_combobox.set('2026')
        self.analysis_year_combobox.pack(pady=5, fill="x")

        ctk.CTkLabel(left_frame, text="Tipo de Análisis:", font=("Arial", 14, "bold")).pack(pady=(20, 5))
        self.analysis_menu = ctk.CTkOptionMenu(left_frame,
                                               values=["Análisis Vertical",
                                                       "Análisis Horizontal",
                                                       "Ratios Financieros",
                                                       "Análisis Dupont",
                                                       "EVA",
                                                       "Pareto (con gráfico)",
                                                       "Análisis de Variaciones Control"],
                                               command=self.on_analysis_menu_changed)
        self.analysis_menu.pack(pady=5, fill="x")
        self.analysis_menu.set("Análisis Vertical")

        self.var_view_seg = ctk.CTkSegmentedButton(left_frame,
                                                   values=["Por Periodo", "Por Centro", "Por Elemento", "Matriz Cruzada"],
                                                   command=self.on_var_view_changed)
        self.var_view_seg.set("Por Periodo")
        self.var_view_seg.pack(pady=5, fill="x")
        self.var_view_seg.pack_forget()

        right_frame = ctk.CTkFrame(main_frame)
        right_frame.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)
        right_frame.grid_rowconfigure(0, weight=1)
        right_frame.grid_rowconfigure(1, weight=0)
        right_frame.grid_columnconfigure(0, weight=1)

        self.analysis_textbox = ctk.CTkTextbox(right_frame, wrap="none", font=("Courier", 13))
        self.analysis_textbox.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)

        self.analysis_canvas_frame = ctk.CTkFrame(right_frame)
        self.analysis_canvas_frame.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        self.analysis_canvas_frame.grid_remove()

        self.refresh_analysis_combos()
        self.on_analysis_menu_changed("Análisis Vertical")

    def refresh_analysis_combos(self):
        periods = self.model.get_periods()
        self.analysis_period_combobox.configure(values=[p[1] for p in periods])
        if periods:
            self.analysis_period_combobox.set(periods[0][1])
        versions = self.model.get_versions()
        self.analysis_version_combobox.configure(values=[v[1] for v in versions])
        if versions:
            self.analysis_version_combobox.set(versions[0][1])

    def on_analysis_menu_changed(self, choice):
        if choice == "Análisis de Variaciones Control":
            self.var_view_seg.pack(pady=5, fill="x")
            self.var_view_seg.set("Por Periodo")
            self.run_variance_analysis("Por Periodo")
        else:
            self.var_view_seg.pack_forget()
            self.run_selected_analysis(choice)

    def on_var_view_changed(self, view):
        self.run_variance_analysis(view)

    def run_variance_analysis(self, view):
        period_name = self.analysis_period_combobox.get()
        version_name = self.analysis_version_combobox.get()
        if not period_name or not version_name:
            self.analysis_textbox.delete("1.0", "end")
            self.analysis_textbox.insert("1.0", "Seleccione período y versión.")
            self.analysis_canvas_frame.grid_remove()
            return
        period = self.model.db.fetch_one("SELECT id FROM periods WHERE name=?", (period_name,))
        version = self.model.db.fetch_one("SELECT id FROM versions WHERE name=?", (version_name,))
        if not period or not version:
            self.analysis_textbox.delete("1.0", "end")
            self.analysis_textbox.insert("1.0", "Período o versión no válidos.")
            self.analysis_canvas_frame.grid_remove()
            return

        analyzer = FinancialAnalyzer(self.model)
        dataframes = analyzer.generate_variance_multidimensional_analysis(period[0])
        if not dataframes:
            self.analysis_textbox.delete("1.0", "end")
            self.analysis_textbox.insert("1.0", "No hay datos para el período seleccionado.")
            self.analysis_canvas_frame.grid_remove()
            return

        if view == "Por Periodo":
            df = dataframes.get('por_periodo')
            title = "VARIACIÓN POR PERIODO (Año/Mes)"
        elif view == "Por Centro":
            df = dataframes.get('por_centro')
            title = "VARIACIÓN POR CENTRO DE ACTIVIDAD"
        elif view == "Por Elemento":
            df = dataframes.get('por_elemento')
            title = "VARIACIÓN POR ELEMENTO DE PRESUPUESTO"
        elif view == "Matriz Cruzada":
            df = dataframes.get('matriz_cruzada')
            title = "MATRIZ CRUZADA: Centro vs Elemento (Variación Absoluta Original)"
        else:
            return

        if df is None or df.empty:
            self.analysis_textbox.delete("1.0", "end")
            self.analysis_textbox.insert("1.0", "No hay datos para la vista seleccionada.")
            self.analysis_canvas_frame.grid_remove()
            return

        table = self._format_dataframe_as_table(df, title)
        self.analysis_textbox.delete("1.0", "end")
        self.analysis_textbox.insert("1.0", table)
        self.analysis_canvas_frame.grid_remove()

    def run_selected_analysis(self, choice):
        period_name = self.analysis_period_combobox.get()
        version_name = self.analysis_version_combobox.get()
        year = int(self.analysis_year_combobox.get())
        if not period_name or not version_name:
            self.analysis_textbox.delete("1.0", "end")
            self.analysis_textbox.insert("1.0", "Seleccione período y versión.")
            self.analysis_canvas_frame.grid_remove()
            return
        period = self.model.db.fetch_one("SELECT id FROM periods WHERE name=?", (period_name,))
        version = self.model.db.fetch_one("SELECT id FROM versions WHERE name=?", (version_name,))
        if not period or not version:
            self.analysis_textbox.delete("1.0", "end")
            self.analysis_textbox.insert("1.0", "Período o versión no válidos.")
            self.analysis_canvas_frame.grid_remove()
            return

        analyzer = FinancialAnalyzer(self.model)
        self.analysis_textbox.delete("1.0", "end")
        self.analysis_canvas_frame.grid_remove()

        if choice == "Análisis Vertical":
            self._show_vertical_analysis(period[0], version[0], analyzer)
        elif choice == "Análisis Horizontal":
            self._show_horizontal_analysis(period[0], version[0], analyzer)
        elif choice == "Ratios Financieros":
            self._show_ratios(period[0], version[0], year, analyzer)
        elif choice == "Análisis Dupont":
            self._show_dupont(period[0], version[0], year, analyzer)
        elif choice == "EVA":
            self._show_eva(period[0], version[0], year, analyzer)
        elif choice == "Pareto (con gráfico)":
            self._show_pareto(period[0], version[0], year, analyzer)

    def _show_vertical_analysis(self, period_id, version_id, analyzer):
        self.analysis_textbox.insert("end", "=== ANÁLISIS VERTICAL - BALANCE ===\n")
        vert_bs = analyzer.vertical_analysis(period_id, version_id, 'BS')
        if vert_bs:
            for item in vert_bs:
                self.analysis_textbox.insert("end", f"{item[1]}: {item[2]:,.2f} ({item[3]:.2f}%) - Año {item[4]}\n")
        else:
            self.analysis_textbox.insert("end", "No hay datos para Balance.\n")

        self.analysis_textbox.insert("end", "\n=== ANÁLISIS VERTICAL - RESULTADOS ===\n")
        vert_is = analyzer.vertical_analysis(period_id, version_id, 'IS')
        if vert_is:
            for item in vert_is:
                self.analysis_textbox.insert("end", f"{item[1]}: {item[2]:,.2f} ({item[3]:.2f}%) - Año {item[4]}\n")
        else:
            self.analysis_textbox.insert("end", "No hay datos para Resultados.\n")

    def _show_horizontal_analysis(self, period_id, version_id, analyzer):
        self.analysis_textbox.insert("end", "=== ANÁLISIS HORIZONTAL - BALANCE ===\n")
        horiz_bs = analyzer.horizontal_analysis(period_id, version_id, 'BS')
        if horiz_bs is not None and not horiz_bs.empty:
            self.analysis_textbox.insert("end", horiz_bs.to_string(index=False) + "\n")
        else:
            self.analysis_textbox.insert("end", "No hay datos para Balance.\n")

        self.analysis_textbox.insert("end", "\n=== ANÁLISIS HORIZONTAL - RESULTADOS ===\n")
        horiz_is = analyzer.horizontal_analysis(period_id, version_id, 'IS')
        if horiz_is is not None and not horiz_is.empty:
            self.analysis_textbox.insert("end", horiz_is.to_string(index=False) + "\n")
        else:
            self.analysis_textbox.insert("end", "No hay datos para Resultados.\n")

    def _show_ratios(self, period_id, version_id, year, analyzer):
        ratios = analyzer.ratios(period_id, version_id, year)
        self.analysis_textbox.insert("end", f"=== RATIOS FINANCIEROS (Año {year}) ===\n")
        for k, v in ratios.items():
            self.analysis_textbox.insert("end", f"{k}: {v:.2f}\n")

    def _show_dupont(self, period_id, version_id, year, analyzer):
        dupont = analyzer.dupont_analysis(period_id, version_id, year)
        self.analysis_textbox.insert("end", f"=== ANÁLISIS DUPONT (Año {year}) ===\n")
        for k, v in dupont.items():
            self.analysis_textbox.insert("end", f"{k}: {v:.4f}\n")

    def _show_eva(self, period_id, version_id, year, analyzer):
        eva = analyzer.eva(period_id, version_id, year, wacc=0.1)
        self.analysis_textbox.insert("end", f"=== EVA (Año {year}, WACC 10%) ===\n")
        for k, v in eva.items():
            self.analysis_textbox.insert("end", f"{k}: {v:,.2f}\n")

    def _show_pareto(self, period_id, version_id, year, analyzer):
        pareto, all_costs = analyzer.pareto_cost(period_id, version_id, year)
        self.analysis_textbox.insert("end", f"=== PARETO DE COSTOS (Año {year}) ===\n")
        if pareto.empty:
            self.analysis_textbox.insert("end", "No hay datos para Pareto.\n")
            return
        self.analysis_textbox.insert("end", "Elementos que acumulan el 80% del costo:\n")
        for _, row in pareto.iterrows():
            self.analysis_textbox.insert("end", f"{row['element_name']}: {row['actual_original']:,.2f} ({row['cum_pct']:.2f}% acumulado)\n")

        self.analysis_canvas_frame.grid()
        for widget in self.analysis_canvas_frame.winfo_children():
            widget.destroy()

        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(10, 4))
        sorted_costs = all_costs.sort_values('actual_original', ascending=False)
        ax1.bar(sorted_costs['element_name'], sorted_costs['actual_original'], color='steelblue')
        ax1.set_title('Costos por Elemento')
        ax1.set_xlabel('Elemento')
        ax1.set_ylabel('Monto')
        plt.setp(ax1.get_xticklabels(), rotation=45, ha='right')

        sorted_costs['cum_pct'] = sorted_costs['actual_original'].cumsum() / sorted_costs['actual_original'].sum() * 100
        ax2.bar(sorted_costs['element_name'], sorted_costs['actual_original'], color='skyblue', alpha=0.7)
        ax2.plot(sorted_costs['element_name'], sorted_costs['cum_pct'], 'r-o', linewidth=2)
        ax2.set_title('Pareto (80/20)')
        ax2.set_xlabel('Elemento')
        ax2.set_ylabel('Porcentaje Acumulado')
        plt.setp(ax2.get_xticklabels(), rotation=45, ha='right')
        ax2.axhline(y=80, color='r', linestyle='--', alpha=0.5, label='80%')
        ax2.legend()

        plt.tight_layout()

        canvas = FigureCanvasTkAgg(fig, master=self.analysis_canvas_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

        toolbar = NavigationToolbar2Tk(canvas, self.analysis_canvas_frame)
        toolbar.update()
        canvas.get_tk_widget().pack(fill="both", expand=True)

    # ---------- Pestaña Correo ----------
    def setup_email_tab(self):
        tab = self.tabview.tab("Correo")
        main_frame = ctk.CTkFrame(tab)
        main_frame.pack(fill="both", expand=True, padx=5, pady=5)
        main_frame.grid_rowconfigure(0, weight=1)
        main_frame.grid_columnconfigure(0, weight=0, minsize=280)
        main_frame.grid_columnconfigure(1, weight=1)

        left_frame = ctk.CTkFrame(main_frame)
        left_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        ctk.CTkLabel(left_frame, text="Configuración SMTP", font=("Arial", 14, "bold")).pack(pady=(0, 10))

        ctk.CTkLabel(left_frame, text="Servidor:").pack(anchor="w")
        self.smtp_server_entry = ctk.CTkEntry(left_frame)
        self.smtp_server_entry.insert(0, "smtp.gmail.com")
        self.smtp_server_entry.pack(pady=5, fill="x")

        ctk.CTkLabel(left_frame, text="Puerto:").pack(anchor="w", pady=(10, 0))
        self.smtp_port_entry = ctk.CTkEntry(left_frame)
        self.smtp_port_entry.insert(0, "587")
        self.smtp_port_entry.pack(pady=5, fill="x")

        ctk.CTkLabel(left_frame, text="Usuario:").pack(anchor="w", pady=(10, 0))
        self.smtp_user_entry = ctk.CTkEntry(left_frame)
        self.smtp_user_entry.pack(pady=5, fill="x")

        ctk.CTkLabel(left_frame, text="Contraseña:").pack(anchor="w", pady=(10, 0))
        self.smtp_pass_entry = ctk.CTkEntry(left_frame, show="*")
        self.smtp_pass_entry.pack(pady=5, fill="x")

        ctk.CTkLabel(left_frame, text="Destinatario:").pack(anchor="w", pady=(10, 0))
        self.email_to_entry = ctk.CTkEntry(left_frame)
        self.email_to_entry.pack(pady=5, fill="x")

        btn_send = ctk.CTkButton(left_frame, text="Enviar Reporte Segmentado", command=self.send_segmented_email)
        btn_send.pack(pady=20, fill="x")

        right_frame = ctk.CTkFrame(main_frame)
        right_frame.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)
        ctk.CTkLabel(right_frame, text="Estado", font=("Arial", 14, "bold")).pack(pady=(0, 5))
        self.email_status = ctk.CTkTextbox(right_frame, wrap="word", font=("Courier", 13))
        self.email_status.pack(fill="both", expand=True, padx=5, pady=5)
        self.email_status.insert("1.0", "Esperando acción...")

    def send_segmented_email(self):
        smtp_server = self.smtp_server_entry.get()
        smtp_port = int(self.smtp_port_entry.get())
        username = self.smtp_user_entry.get()
        password = self.smtp_pass_entry.get()
        if not all([smtp_server, smtp_port, username, password]):
            messagebox.showwarning("Advertencia", "Complete las credenciales SMTP.")
            return

        period_name = self.report_period_combobox.get()
        version_name = self.report_version_combobox.get()
        if not period_name or not version_name:
            messagebox.showwarning("Advertencia", "Seleccione período y versión en la pestaña Reportes.")
            return
        period = self.model.db.fetch_one("SELECT id FROM periods WHERE name=?", (period_name,))
        version = self.model.db.fetch_one("SELECT id FROM versions WHERE name=?", (version_name,))
        if not period or not version:
            return

        centers = self.model.get_centers()
        if not centers:
            messagebox.showinfo("Info", "No hay centros activos para enviar.")
            return

        sender = EmailSender(smtp_server, smtp_port, username, password)
        self.email_status.delete("1.0", "end")
        self.email_status.insert("1.0", "Enviando correos...\n")

        for center in centers:
            center_id = center[0]
            manager_email = center[4]
            if not manager_email:
                continue
            vars_ = self.model.get_variances(period[0], center_id)
            critical = [v for v in vars_ if abs(v[7]) > 0 and abs(v[8]) > 10]
            if not critical:
                continue

            body = f"Reporte de variaciones críticas para el centro {center[1]}\n"
            body += f"Período: {period_name} - Versión: {version_name}\n\n"
            body += "Variaciones críticas (>10%):\n"
            for v in critical:
                element = self.model.db.fetch_one("SELECT name FROM elements WHERE id=?", (v[3],))
                el_name = element[0] if element else "N/A"
                body += f"  - {el_name}: {v[7]:+.2f} ({v[8]:+.2f}%) - {v[5]}-{v[4]:02d} (tipo: {v[6]})\n"
            body += "\nPor favor, revise estas desviaciones y tome acciones correctivas."

            subject = f"Alertas Presupuestarias - {center[1]} - {period_name}"
            success = sender.send_budget_report(manager_email, subject, body, None)
            if success:
                self.email_status.insert("end", f"✅ Enviado a {manager_email}\n")
            else:
                self.email_status.insert("end", f"❌ Error al enviar a {manager_email}\n")

        self.email_status.insert("end", "Proceso completado.\n")

    # ---------- Pestaña Optimización ----------
    def setup_optimization_tab(self):
        tab = self.tabview.tab("Optimización")
        main_frame = ctk.CTkFrame(tab)
        main_frame.pack(fill="both", expand=True, padx=5, pady=5)
        main_frame.grid_rowconfigure(0, weight=1)
        main_frame.grid_columnconfigure(0, weight=0, minsize=280)
        main_frame.grid_columnconfigure(1, weight=1)

        left_frame = ctk.CTkFrame(main_frame)
        left_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        ctk.CTkLabel(left_frame, text="Recomendaciones", font=("Arial", 14, "bold")).pack(pady=(0, 10))

        ctk.CTkLabel(left_frame, text="Período:").pack(anchor="w")
        self.opt_period_combobox = ctk.CTkComboBox(left_frame, values=[], state="readonly")
        self.opt_period_combobox.pack(pady=5, fill="x")

        ctk.CTkLabel(left_frame, text="Versión:").pack(anchor="w", pady=(10, 0))
        self.opt_version_combobox = ctk.CTkComboBox(left_frame, values=[], state="readonly")
        self.opt_version_combobox.pack(pady=5, fill="x")

        btn_optimize = ctk.CTkButton(left_frame, text="Generar Recomendaciones", command=self.generate_optimizations)
        btn_optimize.pack(pady=20, fill="x")

        right_frame = ctk.CTkFrame(main_frame)
        right_frame.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)
        ctk.CTkLabel(right_frame, text="Sugerencias", font=("Arial", 14, "bold")).pack(pady=(0, 5))
        self.opt_textbox = ctk.CTkTextbox(right_frame, wrap="none", font=("Courier", 13))
        self.opt_textbox.pack(fill="both", expand=True, padx=5, pady=5)

        self.refresh_opt_combos()
        self.generate_optimizations()

    def refresh_opt_combos(self):
        periods = self.model.get_periods()
        self.opt_period_combobox.configure(values=[p[1] for p in periods])
        if periods:
            self.opt_period_combobox.set(periods[0][1])
        versions = self.model.get_versions()
        self.opt_version_combobox.configure(values=[v[1] for v in versions])
        if versions:
            self.opt_version_combobox.set(versions[0][1])

    def generate_optimizations(self):
        period_name = self.opt_period_combobox.get()
        version_name = self.opt_version_combobox.get()
        if not period_name or not version_name:
            self.opt_textbox.delete("1.0", "end")
            self.opt_textbox.insert("1.0", "Seleccione período y versión.")
            return
        period = self.model.db.fetch_one("SELECT id FROM periods WHERE name=?", (period_name,))
        version = self.model.db.fetch_one("SELECT id FROM versions WHERE name=?", (version_name,))
        if not period or not version:
            self.opt_textbox.delete("1.0", "end")
            self.opt_textbox.insert("1.0", "Período o versión no válidos.")
            return

        advisor = OptimizationAdvisor(self.model)
        report_lines = advisor.generate_recommendations(period[0], version[0])
        self.opt_textbox.delete("1.0", "end")
        if isinstance(report_lines, list):
            self.opt_textbox.insert("1.0", "\n".join(report_lines))
        else:
            self.opt_textbox.insert("1.0", str(report_lines))


# ---------------------- PUNTO DE ENTRADA ----------------------
if __name__ == "__main__":
    app = MainApp()
    app.mainloop()
