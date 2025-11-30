#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Estadística Hospital - Automatización de descarga y procesamiento de estadísticas
Versión 3.1 con configuración desde la interfaz gráfica

Autor: Automatizado con Claude
Fecha: 2025
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog, font as tkfont
import configparser
import json
import os
import sys
import re
import threading
import subprocess
from pathlib import Path
from datetime import datetime, date, timedelta
from typing import Optional
import time
import calendar

# Intentar importar dependencias
try:
    import pandas as pd
    from openpyxl import load_workbook
except ImportError as e:
    print(f"Error: Falta instalar dependencias. Ejecute: pip install pandas openpyxl")
    sys.exit(1)

try:
    from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout
except ImportError:
    print("Error: Falta instalar playwright. Ejecute: pip install playwright")
    sys.exit(1)

# Intentar importar tkcalendar, si no está disponible usar alternativa
try:
    from tkcalendar import DateEntry
    HAS_TKCALENDAR = True
except ImportError:
    HAS_TKCALENDAR = False


# ============================================
# CONFIGURACIÓN POR DEFECTO
# ============================================

DEFAULT_CONFIG = {
    "General": {
        "URL": "https://hjmvi.orion-labs.com/informes/estadisticos",
        "URLCatalogo": "https://hjmvi.orion-labs.com/informes/catalogos",
        "Headless": "false"
    },
    "Informe": {
        "IdDropdownAgrupar": "agrupar-por",
        "ValorAgrupacion": "SECCION_TIPO_ATENCION",
        "IdFechaDesde": "fecha-orden-desde",
        "IdFechaHasta": "fecha-orden-hasta"
    },
    "Catalogo": {
        "IdDropdownTipo": "tipo",
        "ValorExamenes": "EXAMENES",
        "IdBotonGenerar": "generar-informe-auditorias"
    },
    "Archivos": {
        "CarpetaDescargas": "./ExcelsDescargados",
        "ArchivoSalida": "./Estadistica Hospital.xlsx",
        "ArchivoCatalogo": "./catalogo_examenes.json"
    },
    "ColumnasMerge": {
        "HospitalizacionTotal": "Hospitalización,URGENTE HOSPITALIZACION",
        "ConsultaExternaTotal": "Consulta Externa,URGENTE CONSULTA EXTERNA,REFERENCIA,URGENTE REFERENCIA",
        "Emergencia": "Emergencia,Sin tipo atención"
    }
}

DEFAULT_EXAM_CONFIG = {
    "multipliers": {
        "BIOMETRÍA HEMÁTICA": 18,
        "COPROPARASITARIO": 2,
        "ELEMENTAL Y MICROSCÓPICO DE ORINA": 3,
        "GASOMETRIA ARTERIAL": 14,
        "GASOMETRIA VENOSA": 14,
        "TIPIFICACION SANGUINEA RH (D)": 3
    },
    "cultivo_multiplier": 10,
    "exam_categories": {
        "LEISHMANIA": "Hematologico",
        "CRISTALOGRAFÍA": "Bacteriológico",
        "GRAM (GOTA FRESCA) ORINA": "Bacteriológico",
        "GASOMETRIA ARTERIAL": "Quimica sanguinea",
        "GASOMETRIA VENOSA": "Quimica sanguinea"
    },
    "seccion_categories": {
        # Secciones que van a Serologicos
        "Autoinmunes e Infecciosas": "Serologicos",
        "Drogas y Fármacos": "Serologicos",
        "Inmunología": "Serologicos",
        "Marcadores Tumorales": "Serologicos",
        "Serología": "Serologicos",
        "Estudios de Alergias": "Serologicos",
        "Marcadores Coronarios": "Serologicos",
        "Inmunoquímica Sanguínea": "Serologicos",
        
        # Secciones que van a Quimica sanguinea
        "Bioquímica": "Quimica sanguinea",
        "Electrolitos": "Quimica sanguinea",
        "Gases Arteriales": "Quimica sanguinea",
        
        # Secciones que van a Hematologico
        "Coagulación": "Hematologico",
        "Hematología": "Hematologico",
        "Inmunohematología": "Hematologico",
        "Plaquetas": "Hematologico",
        
        # Secciones que van a Bacteriológico
        "Líquidos Biológicos": "Bacteriológico",
        "Microbiología": "Bacteriológico",
        
        # Secciones que van a Materias fecales
        "Coproanálisis": "Materias fecales",
        
        # Secciones que van a Orina
        "Química Clínica en Orina": "Orina",
        "Uroanálisis": "Orina",
        
        # Secciones que van a Hormonales
        "Estudios Hormonales": "Hormonales",
        
        # Otras secciones (pueden necesitar categorización manual)
        "Biología Molecular": "Serologicos",
        "Citología": "Bacteriológico",
        "Especiales": "Other",
        "Medicina Ocupacional": "Other"
    }
}

NUMERIC_COLUMNS = [
    "REFERENCIA", "Hospitalización", "Emergencia",
    "URGENTE CONSULTA EXTERNA", "Consulta Externa",
    "Sin tipo atención", "URGENTE REFERENCIA",
    "URGENTE HOSPITALIZACION", "Total"
]

CATEGORY_ORDER = [
    "Hematologico", "Bacteriológico", "Quimica sanguinea",
    "Materias fecales", "Orina", "Hormonales", "Serologicos", "Other", "TOTAL"
]


# ============================================
# CLASE PRINCIPAL DE LA APLICACIÓN
# ============================================

class EstadisticaHospitalApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Estadística Hospital v3.6.2")
        self.root.geometry("950x750")
        self.root.minsize(900, 650)
        
        # Determinar directorio base
        if getattr(sys, 'frozen', False):
            self.base_dir = Path(sys.executable).parent
        else:
            self.base_dir = Path(__file__).parent
        
        # Configurar fuentes más grandes
        self.setup_fonts()
        
        # Cargar configuraciones (crea archivos si no existen)
        self.config = self.load_config()
        self.exam_config = self.load_exam_config()
        self.exam_catalog = self.load_exam_catalog()
        
        # Guardar configuraciones por defecto si son nuevas
        self.ensure_config_files_exist()
        
        # Variables de control
        self.should_stop = False
        self.is_running = False
        self.output_file_path = self.base_dir / self.config.get("Archivos", "ArchivoSalida")
        
        # Crear interfaz
        self.create_notebook()
        self.create_main_tab()
        self.create_config_tab()
        self.create_exams_tab()
        self.create_categories_tab()
        
        # Habilitar botón de Excel si el archivo existe
        if self.output_file_path.exists():
            self.open_excel_button.config(state="normal")
    
    def setup_fonts(self):
        """Configura fuentes más grandes para toda la aplicación"""
        default_font = tkfont.nametofont("TkDefaultFont")
        default_font.configure(size=11)
        
        text_font = tkfont.nametofont("TkTextFont")
        text_font.configure(size=11)
        
        fixed_font = tkfont.nametofont("TkFixedFont")
        fixed_font.configure(size=10)
        
        # Configurar estilo para widgets ttk
        style = ttk.Style()
        style.configure(".", font=("Segoe UI", 11))
        style.configure("TLabel", font=("Segoe UI", 11))
        style.configure("TButton", font=("Segoe UI", 11))
        style.configure("TEntry", font=("Segoe UI", 11))
        style.configure("TCombobox", font=("Segoe UI", 11))
        style.configure("TCheckbutton", font=("Segoe UI", 11))
        style.configure("TLabelframe.Label", font=("Segoe UI", 11, "bold"))
        style.configure("TNotebook.Tab", font=("Segoe UI", 11))
        style.configure("Treeview", font=("Segoe UI", 10), rowheight=25)
        style.configure("Treeview.Heading", font=("Segoe UI", 11, "bold"))
    
    def ensure_config_files_exist(self):
        """Crea archivos de configuración si no existen"""
        # config.ini
        config_file = self.base_dir / "config.ini"
        if not config_file.exists():
            self.save_config()
        
        # config_examenes.json
        exam_config_file = self.base_dir / "config_examenes.json"
        if not exam_config_file.exists():
            self.save_exam_config()
        
    def load_config(self) -> configparser.ConfigParser:
        """Carga la configuración general"""
        config = configparser.ConfigParser()
        config_file = self.base_dir / "config.ini"
        
        # Establecer valores por defecto
        for section, values in DEFAULT_CONFIG.items():
            config[section] = values
        
        # Intentar cargar archivo existente
        if config_file.exists():
            try:
                config.read(config_file, encoding='utf-8')
            except:
                pass
        
        return config
    
    def save_config(self):
        """Guarda la configuración general"""
        config_file = self.base_dir / "config.ini"
        with open(config_file, 'w', encoding='utf-8') as f:
            self.config.write(f)
    
    def load_exam_config(self) -> dict:
        """Carga la configuración de exámenes"""
        config_file = self.base_dir / "config_examenes.json"
        
        if config_file.exists():
            try:
                with open(config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                pass
        
        return DEFAULT_EXAM_CONFIG.copy()
    
    def save_exam_config(self):
        """Guarda la configuración de exámenes"""
        config_file = self.base_dir / "config_examenes.json"
        with open(config_file, 'w', encoding='utf-8') as f:
            json.dump(self.exam_config, f, indent=2, ensure_ascii=False)
    
    def load_exam_catalog(self) -> dict:
        """Carga el catálogo de exámenes"""
        catalog_file = self.base_dir / self.config.get("Archivos", "ArchivoCatalogo", fallback="catalogo_examenes.json")
        
        if catalog_file.exists():
            try:
                with open(catalog_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                pass
        
        return {"examenes": [], "ultima_actualizacion": None}
    
    def save_exam_catalog(self, examenes: list):
        """Guarda el catálogo de exámenes"""
        catalog_file = self.base_dir / self.config.get("Archivos", "ArchivoCatalogo", fallback="catalogo_examenes.json")
        catalog = {
            "examenes": examenes,
            "ultima_actualizacion": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        with open(catalog_file, 'w', encoding='utf-8') as f:
            json.dump(catalog, f, indent=2, ensure_ascii=False)
    
    def create_notebook(self):
        """Crea el notebook con pestañas"""
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Crear frames para cada pestaña
        self.main_frame = ttk.Frame(self.notebook)
        self.config_frame = ttk.Frame(self.notebook)
        self.exams_frame = ttk.Frame(self.notebook)
        self.categories_frame = ttk.Frame(self.notebook)
        
        self.notebook.add(self.main_frame, text="  Descarga  ")
        self.notebook.add(self.config_frame, text="  Configuración Web  ")
        self.notebook.add(self.exams_frame, text="  Multiplicadores  ")
        self.notebook.add(self.categories_frame, text="  Categorías  ")
    
    def create_main_tab(self):
        """Crea la pestaña principal de descarga"""
        # Frame de parámetros
        params_frame = ttk.LabelFrame(self.main_frame, text="Parámetros de Descarga", padding=15)
        params_frame.pack(fill="x", padx=10, pady=10)
        
        # Fecha inicial
        ttk.Label(params_frame, text="Fecha inicial:").grid(row=0, column=0, sticky="w", padx=5, pady=8)
        
        # Calcular primer día del mes actual
        today = date.today()
        first_day = date(today.year, today.month, 1)
        
        if HAS_TKCALENDAR:
            self.start_date_entry = DateEntry(params_frame, width=15, date_pattern='yyyy-mm-dd',
                                              font=("Segoe UI", 11), year=first_day.year,
                                              month=first_day.month, day=first_day.day)
            self.start_date_entry.grid(row=0, column=1, sticky="w", padx=5, pady=8)
        else:
            self.start_date_var = tk.StringVar(value=first_day.strftime('%Y-%m-%d'))
            start_entry = ttk.Entry(params_frame, textvariable=self.start_date_var, width=15)
            start_entry.grid(row=0, column=1, sticky="w", padx=5, pady=8)
            ttk.Button(params_frame, text="📅", width=3, 
                      command=lambda: self.show_calendar_popup(self.start_date_var)).grid(row=0, column=2, padx=2)
        
        # Fecha final
        ttk.Label(params_frame, text="Fecha final:").grid(row=0, column=3, sticky="w", padx=(20, 5), pady=8)
        
        if HAS_TKCALENDAR:
            self.end_date_entry = DateEntry(params_frame, width=15, date_pattern='yyyy-mm-dd',
                                            font=("Segoe UI", 11), year=today.year,
                                            month=today.month, day=today.day)
            self.end_date_entry.grid(row=0, column=4, sticky="w", padx=5, pady=8)
        else:
            self.end_date_var = tk.StringVar(value=today.strftime('%Y-%m-%d'))
            end_entry = ttk.Entry(params_frame, textvariable=self.end_date_var, width=15)
            end_entry.grid(row=0, column=4, sticky="w", padx=5, pady=8)
            ttk.Button(params_frame, text="📅", width=3,
                      command=lambda: self.show_calendar_popup(self.end_date_var)).grid(row=0, column=5, padx=2)
        
        # Botones rápidos de fecha
        quick_frame = ttk.Frame(params_frame)
        quick_frame.grid(row=1, column=0, columnspan=6, sticky="w", padx=5, pady=5)
        
        ttk.Label(quick_frame, text="Rápido:").pack(side="left", padx=(0, 10))
        ttk.Button(quick_frame, text="Este mes", command=self.set_this_month).pack(side="left", padx=2)
        ttk.Button(quick_frame, text="Mes anterior", command=self.set_last_month).pack(side="left", padx=2)
        ttk.Button(quick_frame, text="Hoy", command=self.set_today).pack(side="left", padx=2)
        
        # Headless mode
        self.headless_var = tk.BooleanVar(value=self.config.get("General", "Headless", fallback="false").lower() == "true")
        ttk.Checkbutton(params_frame, text="Modo oculto (sin ventana del navegador)", 
                       variable=self.headless_var).grid(row=2, column=0, columnspan=6, sticky="w", padx=5, pady=8)
        
        # Frame de botones de descarga
        buttons_frame = ttk.Frame(self.main_frame)
        buttons_frame.pack(fill="x", padx=10, pady=10)
        
        self.start_button = ttk.Button(buttons_frame, text="▶ Iniciar Descarga", command=self.start_process)
        self.start_button.pack(side="left", padx=5)
        
        self.stop_button = ttk.Button(buttons_frame, text="⏹ Detener", command=self.stop_process, state="disabled")
        self.stop_button.pack(side="left", padx=5)
        
        self.open_excel_button = ttk.Button(buttons_frame, text="📊 Abrir Excel", command=self.open_excel, state="disabled")
        self.open_excel_button.pack(side="left", padx=5)
        
        self.open_folder_button = ttk.Button(buttons_frame, text="📁 Abrir Carpeta", command=self.open_folder)
        self.open_folder_button.pack(side="left", padx=5)
        
        self.recalc_button = ttk.Button(buttons_frame, text="🔄 Recalcular Excel", command=self.recalculate_excel)
        self.recalc_button.pack(side="left", padx=5)
        
        # Frame de catálogo de exámenes
        catalog_frame = ttk.LabelFrame(self.main_frame, text="Catálogo de Exámenes", padding=10)
        catalog_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Button(catalog_frame, text="🔄 Actualizar Catálogo de Exámenes", 
                  command=self.update_exam_catalog).pack(side="left", padx=5)
        
        # Última actualización
        self.last_update_var = tk.StringVar()
        self.update_last_update_label()
        ttk.Label(catalog_frame, textvariable=self.last_update_var, 
                 font=("Segoe UI", 10, "italic")).pack(side="left", padx=20)
        
        # Info del catálogo
        self.catalog_info_var = tk.StringVar()
        self.update_catalog_info()
        ttk.Label(catalog_frame, textvariable=self.catalog_info_var).pack(side="left", padx=10)
        
        # Barra de progreso
        progress_frame = ttk.Frame(self.main_frame)
        progress_frame.pack(fill="x", padx=10, pady=5)
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill="x", expand=True)
        
        self.status_var = tk.StringVar(value="Listo para comenzar")
        ttk.Label(progress_frame, textvariable=self.status_var).pack(anchor="w")
        
        # Log
        log_frame = ttk.LabelFrame(self.main_frame, text="Registro de Actividad", padding=5)
        log_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=12, font=("Consolas", 10))
        self.log_text.pack(fill="both", expand=True)
    
    def get_start_date(self) -> date:
        """Obtiene la fecha de inicio seleccionada"""
        if HAS_TKCALENDAR:
            return self.start_date_entry.get_date()
        else:
            return datetime.strptime(self.start_date_var.get(), '%Y-%m-%d').date()
    
    def get_end_date(self) -> date:
        """Obtiene la fecha final seleccionada"""
        if HAS_TKCALENDAR:
            return self.end_date_entry.get_date()
        else:
            return datetime.strptime(self.end_date_var.get(), '%Y-%m-%d').date()
    
    def set_date(self, start: date, end: date):
        """Establece las fechas en los selectores"""
        if HAS_TKCALENDAR:
            self.start_date_entry.set_date(start)
            self.end_date_entry.set_date(end)
        else:
            self.start_date_var.set(start.strftime('%Y-%m-%d'))
            self.end_date_var.set(end.strftime('%Y-%m-%d'))
    
    def set_this_month(self):
        """Establece las fechas para el mes actual"""
        today = date.today()
        first_day = date(today.year, today.month, 1)
        self.set_date(first_day, today)
    
    def set_last_month(self):
        """Establece las fechas para el mes anterior"""
        today = date.today()
        # Primer día del mes actual
        first_this_month = date(today.year, today.month, 1)
        # Último día del mes anterior
        last_day_prev = first_this_month - timedelta(days=1)
        # Primer día del mes anterior
        first_day_prev = date(last_day_prev.year, last_day_prev.month, 1)
        self.set_date(first_day_prev, last_day_prev)
    
    def set_today(self):
        """Establece las fechas para hoy"""
        today = date.today()
        self.set_date(today, today)
    
    def show_calendar_popup(self, date_var):
        """Muestra un popup de calendario simple (fallback si no hay tkcalendar)"""
        popup = tk.Toplevel(self.root)
        popup.title("Seleccionar Fecha")
        popup.geometry("250x200")
        popup.transient(self.root)
        popup.grab_set()
        
        # Parse current date
        try:
            current = datetime.strptime(date_var.get(), '%Y-%m-%d').date()
        except:
            current = date.today()
        
        # Year and month selection
        frame = ttk.Frame(popup, padding=10)
        frame.pack(fill="both", expand=True)
        
        ttk.Label(frame, text="Año:").grid(row=0, column=0, padx=5, pady=5)
        year_var = tk.StringVar(value=str(current.year))
        year_spin = ttk.Spinbox(frame, from_=2020, to=2030, textvariable=year_var, width=6)
        year_spin.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(frame, text="Mes:").grid(row=1, column=0, padx=5, pady=5)
        month_var = tk.StringVar(value=str(current.month))
        month_spin = ttk.Spinbox(frame, from_=1, to=12, textvariable=month_var, width=6)
        month_spin.grid(row=1, column=1, padx=5, pady=5)
        
        ttk.Label(frame, text="Día:").grid(row=2, column=0, padx=5, pady=5)
        day_var = tk.StringVar(value=str(current.day))
        day_spin = ttk.Spinbox(frame, from_=1, to=31, textvariable=day_var, width=6)
        day_spin.grid(row=2, column=1, padx=5, pady=5)
        
        def apply_date():
            try:
                new_date = date(int(year_var.get()), int(month_var.get()), int(day_var.get()))
                date_var.set(new_date.strftime('%Y-%m-%d'))
                popup.destroy()
            except ValueError as e:
                messagebox.showerror("Error", f"Fecha inválida: {e}")
        
        ttk.Button(frame, text="Aceptar", command=apply_date).grid(row=3, column=0, columnspan=2, pady=10)
    
    def update_last_update_label(self):
        """Actualiza la etiqueta de última actualización"""
        ultima = self.exam_catalog.get("ultima_actualizacion")
        if ultima:
            self.last_update_var.set(f"Última actualización: {ultima}")
        else:
            self.last_update_var.set("Catálogo no descargado")
    
    def update_catalog_info(self):
        """Actualiza la información del catálogo"""
        examenes = self.exam_catalog.get("examenes", {})
        if isinstance(examenes, dict):
            total = sum(len(exams) for exams in examenes.values())
            secciones = len(examenes)
            self.catalog_info_var.set(f"({total} exámenes en {secciones} secciones)")
        elif isinstance(examenes, list):
            self.catalog_info_var.set(f"({len(examenes)} exámenes)")
        else:
            self.catalog_info_var.set("")
    
    def create_config_tab(self):
        """Crea la pestaña de configuración web"""
        # Crear un canvas con scrollbar para permitir scroll
        canvas = tk.Canvas(self.config_frame)
        scrollbar = ttk.Scrollbar(self.config_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Bind mouse wheel
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # Frame de URL y navegación
        web_frame = ttk.LabelFrame(scrollable_frame, text="Configuración del Sitio Web", padding=10)
        web_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(web_frame, text="URL del sitio:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.url_var = tk.StringVar(value=self.config.get("General", "URL"))
        ttk.Entry(web_frame, textvariable=self.url_var, width=60).grid(row=0, column=1, sticky="ew", padx=5, pady=2)
        
        # Frame de elementos del formulario
        form_frame = ttk.LabelFrame(scrollable_frame, text="Elementos del Formulario (IDs HTML)", padding=10)
        form_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(form_frame, text="ID Dropdown Agrupar:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.dropdown_id_var = tk.StringVar(value=self.config.get("Informe", "IdDropdownAgrupar"))
        ttk.Entry(form_frame, textvariable=self.dropdown_id_var, width=30).grid(row=0, column=1, sticky="w", padx=5, pady=2)
        
        ttk.Label(form_frame, text="Valor a seleccionar:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.dropdown_value_var = tk.StringVar(value=self.config.get("Informe", "ValorAgrupacion"))
        ttk.Entry(form_frame, textvariable=self.dropdown_value_var, width=30).grid(row=1, column=1, sticky="w", padx=5, pady=2)
        
        ttk.Label(form_frame, text="ID Campo Fecha Desde:").grid(row=2, column=0, sticky="w", padx=5, pady=2)
        self.fecha_desde_var = tk.StringVar(value=self.config.get("Informe", "IdFechaDesde"))
        ttk.Entry(form_frame, textvariable=self.fecha_desde_var, width=30).grid(row=2, column=1, sticky="w", padx=5, pady=2)
        
        ttk.Label(form_frame, text="ID Campo Fecha Hasta:").grid(row=3, column=0, sticky="w", padx=5, pady=2)
        self.fecha_hasta_var = tk.StringVar(value=self.config.get("Informe", "IdFechaHasta"))
        ttk.Entry(form_frame, textvariable=self.fecha_hasta_var, width=30).grid(row=3, column=1, sticky="w", padx=5, pady=2)
        
        # Frame de archivos
        files_frame = ttk.LabelFrame(scrollable_frame, text="Rutas de Archivos", padding=10)
        files_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(files_frame, text="Carpeta de descargas:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.downloads_folder_var = tk.StringVar(value=self.config.get("Archivos", "CarpetaDescargas"))
        ttk.Entry(files_frame, textvariable=self.downloads_folder_var, width=40).grid(row=0, column=1, sticky="w", padx=5, pady=2)
        
        ttk.Label(files_frame, text="Archivo de salida:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.output_file_var = tk.StringVar(value=self.config.get("Archivos", "ArchivoSalida"))
        ttk.Entry(files_frame, textvariable=self.output_file_var, width=40).grid(row=1, column=1, sticky="w", padx=5, pady=2)
        
        # Frame de columnas merge
        merge_frame = ttk.LabelFrame(scrollable_frame, text="Columnas a Combinar en Estadística Calculada", padding=10)
        merge_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(merge_frame, text="Las columnas se combinan sumándolas. Separe con comas.",
                 font=("", 9, "italic")).grid(row=0, column=0, columnspan=2, sticky="w", padx=5, pady=2)
        
        ttk.Label(merge_frame, text="Hospitalización Total =").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.hosp_total_var = tk.StringVar(value=self.config.get("ColumnasMerge", "HospitalizacionTotal", 
                                           fallback="Hospitalización,URGENTE HOSPITALIZACION"))
        ttk.Entry(merge_frame, textvariable=self.hosp_total_var, width=50).grid(row=1, column=1, sticky="w", padx=5, pady=2)
        
        ttk.Label(merge_frame, text="Consulta Externa Total =").grid(row=2, column=0, sticky="w", padx=5, pady=2)
        self.cons_total_var = tk.StringVar(value=self.config.get("ColumnasMerge", "ConsultaExternaTotal",
                                           fallback="Consulta Externa,URGENTE CONSULTA EXTERNA,REFERENCIA,URGENTE REFERENCIA"))
        ttk.Entry(merge_frame, textvariable=self.cons_total_var, width=50).grid(row=2, column=1, sticky="w", padx=5, pady=2)
        
        ttk.Label(merge_frame, text="Emergencia =").grid(row=3, column=0, sticky="w", padx=5, pady=2)
        self.emerg_var = tk.StringVar(value=self.config.get("ColumnasMerge", "Emergencia",
                                      fallback="Emergencia,Sin tipo atención"))
        ttk.Entry(merge_frame, textvariable=self.emerg_var, width=50).grid(row=3, column=1, sticky="w", padx=5, pady=2)
        
        # Info sobre columnas disponibles
        ttk.Label(merge_frame, text="Columnas disponibles en los datos descargados:",
                 font=("", 9)).grid(row=4, column=0, columnspan=2, sticky="w", padx=5, pady=(10, 2))
        available_cols = "REFERENCIA, Hospitalización, Emergencia, URGENTE CONSULTA EXTERNA, " \
                        "Consulta Externa, Sin tipo atención, URGENTE REFERENCIA, URGENTE HOSPITALIZACION, Total"
        ttk.Label(merge_frame, text=available_cols, font=("", 8, "italic"), 
                 wraplength=500).grid(row=5, column=0, columnspan=2, sticky="w", padx=5, pady=2)
        
        # Botón guardar
        ttk.Button(scrollable_frame, text="💾 Guardar Configuración", 
                  command=self.save_web_config).pack(pady=10)
    
    def create_exams_tab(self):
        """Crea la pestaña de multiplicadores de exámenes"""
        # Frame superior con instrucciones
        info_frame = ttk.Frame(self.exams_frame)
        info_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(info_frame, text="Configure multiplicadores para exámenes específicos.\n"
                 "Los exámenes con 'CULTIVO' en el nombre usan el multiplicador de cultivos automáticamente.",
                 justify="left").pack(anchor="w")
        
        # Frame de multiplicador de cultivos
        cultivo_frame = ttk.LabelFrame(self.exams_frame, text="Multiplicador de Cultivos", padding=10)
        cultivo_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(cultivo_frame, text="Multiplicador para exámenes con 'CULTIVO':").pack(side="left", padx=5)
        self.cultivo_mult_var = tk.StringVar(value=str(self.exam_config.get("cultivo_multiplier", 10)))
        ttk.Entry(cultivo_frame, textvariable=self.cultivo_mult_var, width=10).pack(side="left", padx=5)
        
        # Frame de lista de multiplicadores
        list_frame = ttk.LabelFrame(self.exams_frame, text="Multiplicadores por Examen", padding=10)
        list_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        # Treeview para multiplicadores
        columns = ("exam", "multiplier")
        self.mult_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=10)
        self.mult_tree.heading("exam", text="Nombre del Examen")
        self.mult_tree.heading("multiplier", text="Multiplicador")
        self.mult_tree.column("exam", width=400)
        self.mult_tree.column("multiplier", width=100)
        
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.mult_tree.yview)
        self.mult_tree.configure(yscrollcommand=scrollbar.set)
        
        self.mult_tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Cargar datos
        self.refresh_multipliers_list()
        
        # Frame de edición con búsqueda
        edit_frame = ttk.LabelFrame(self.exams_frame, text="Agregar Multiplicador", padding=10)
        edit_frame.pack(fill="x", padx=10, pady=5)
        
        # Combobox con búsqueda
        ttk.Label(edit_frame, text="Examen:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.new_exam_var = tk.StringVar()
        self.exam_combobox = ttk.Combobox(edit_frame, textvariable=self.new_exam_var, width=50)
        self.exam_combobox.grid(row=0, column=1, sticky="w", padx=5, pady=2)
        
        # Poblar combobox con catálogo
        self.update_exam_combobox()
        
        # Bind para búsqueda fuzzy
        self.exam_combobox.bind('<KeyRelease>', self.filter_exam_combobox)
        
        ttk.Label(edit_frame, text="Multiplicador:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.new_mult_var = tk.StringVar(value="1")
        ttk.Entry(edit_frame, textvariable=self.new_mult_var, width=10).grid(row=1, column=1, sticky="w", padx=5, pady=2)
        
        # Botones
        btn_frame = ttk.Frame(edit_frame)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=5)
        
        ttk.Button(btn_frame, text="➕ Agregar", command=self.add_multiplier).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="🗑️ Eliminar Seleccionado", command=self.delete_multiplier).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="💾 Guardar Multiplicadores", command=self.save_multipliers).pack(side="left", padx=5)
    
    def update_exam_combobox(self):
        """Actualiza el combobox con la lista de exámenes del catálogo"""
        examenes_dict = self.exam_catalog.get("examenes", {})
        
        # Handle both old format (list) and new format (dict by section)
        if isinstance(examenes_dict, list):
            examenes = examenes_dict
        else:
            # Flatten dict to list
            examenes = []
            for section_exams in examenes_dict.values():
                examenes.extend(section_exams)
            examenes = sorted(set(examenes))
        
        self.exam_combobox['values'] = examenes
    
    def filter_exam_combobox(self, event):
        """Filtra el combobox con búsqueda fuzzy"""
        typed = self.new_exam_var.get().upper()
        if not typed:
            self.update_exam_combobox()
            return
        
        examenes_dict = self.exam_catalog.get("examenes", {})
        
        # Handle both formats
        if isinstance(examenes_dict, list):
            examenes = examenes_dict
        else:
            examenes = []
            for section_exams in examenes_dict.values():
                examenes.extend(section_exams)
        
        # Filtro fuzzy: buscar exámenes que contengan las palabras escritas
        words = typed.split()
        filtered = []
        for exam in examenes:
            exam_upper = exam.upper()
            if all(word in exam_upper for word in words):
                filtered.append(exam)
        
        # Eliminar duplicados y ordenar
        filtered = sorted(set(filtered))
        
        self.exam_combobox['values'] = filtered[:20]  # Limitar a 20 resultados
        
        # Mostrar dropdown si hay resultados
        if filtered and len(typed) >= 2:
            self.exam_combobox.event_generate('<Down>')
    
    def create_categories_tab(self):
        """Crea la pestaña de categorías"""
        # Notebook interno para subcategorías
        cat_notebook = ttk.Notebook(self.categories_frame)
        cat_notebook.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Tab de categorías por examen
        exam_cat_frame = ttk.Frame(cat_notebook)
        cat_notebook.add(exam_cat_frame, text=" Por Examen ")
        
        self.create_category_list(exam_cat_frame, "exam")
        
        # Tab de categorías por sección
        section_cat_frame = ttk.Frame(cat_notebook)
        cat_notebook.add(section_cat_frame, text=" Por Sección ")
        
        self.create_category_list(section_cat_frame, "section")
        
        # Tab de exámenes sin categorizar
        uncategorized_frame = ttk.Frame(cat_notebook)
        cat_notebook.add(uncategorized_frame, text=" ⚠ Sin Categorizar ")
        
        self.create_uncategorized_list(uncategorized_frame)
    
    def create_category_list(self, parent, cat_type):
        """Crea una lista de categorías"""
        # Treeview
        columns = ("name", "category")
        tree = ttk.Treeview(parent, columns=columns, show="headings", height=12)
        tree.heading("name", text="Examen" if cat_type == "exam" else "Sección")
        tree.heading("category", text="Categoría")
        tree.column("name", width=350)
        tree.column("category", width=150)
        
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        tree.pack(side="top", fill="both", expand=True, padx=10, pady=5)
        scrollbar.place(relx=0.98, rely=0, relheight=0.8, anchor="ne")
        
        # Cargar datos
        if cat_type == "exam":
            self.exam_cat_tree = tree
            for name, cat in self.exam_config.get("exam_categories", {}).items():
                tree.insert("", "end", values=(name, cat))
        else:
            self.section_cat_tree = tree
            for name, cat in self.exam_config.get("seccion_categories", {}).items():
                tree.insert("", "end", values=(name, cat))
        
        # Frame de edición
        edit_frame = ttk.Frame(parent)
        edit_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(edit_frame, text="Nombre:").pack(side="left", padx=5)
        name_var = tk.StringVar()
        name_entry = ttk.Entry(edit_frame, textvariable=name_var, width=30)
        name_entry.pack(side="left", padx=5)
        
        ttk.Label(edit_frame, text="Categoría:").pack(side="left", padx=5)
        cat_var = tk.StringVar()
        cat_combo = ttk.Combobox(edit_frame, textvariable=cat_var, width=20,
                                values=CATEGORY_ORDER[:-2])  # Excluir Other y TOTAL
        cat_combo.pack(side="left", padx=5)
        
        if cat_type == "exam":
            self.exam_name_var = name_var
            self.exam_cat_var = cat_var
            ttk.Button(edit_frame, text="➕", command=self.add_exam_category).pack(side="left", padx=2)
            ttk.Button(edit_frame, text="🗑️", command=self.delete_exam_category).pack(side="left", padx=2)
        else:
            self.section_name_var = name_var
            self.section_cat_var = cat_var
            ttk.Button(edit_frame, text="➕", command=self.add_section_category).pack(side="left", padx=2)
            ttk.Button(edit_frame, text="🗑️", command=self.delete_section_category).pack(side="left", padx=2)
        
        ttk.Button(edit_frame, text="💾 Guardar", command=self.save_categories).pack(side="right", padx=5)
    
    def create_uncategorized_list(self, parent):
        """Crea la lista de exámenes sin categorizar"""
        info_label = ttk.Label(parent, text="Exámenes del catálogo que no tienen categoría asignada.\n"
                              "Actualice el catálogo desde la pestaña 'Descarga' para obtener la lista completa.")
        info_label.pack(padx=10, pady=5, anchor="w")
        
        # Frame de botones
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Button(btn_frame, text="🔍 Mostrar Sin Categorizar", 
                  command=self.show_uncategorized).pack(side="left", padx=5)
        
        ttk.Button(btn_frame, text="🔄 Auto-Categorizar desde Catálogo", 
                  command=self.auto_categorize_from_catalog).pack(side="left", padx=5)
        
        # Treeview
        columns = ("exam", "section")
        self.uncat_tree = ttk.Treeview(parent, columns=columns, show="headings", height=12)
        self.uncat_tree.heading("exam", text="Examen Sin Categorizar")
        self.uncat_tree.heading("section", text="Sección (del catálogo)")
        self.uncat_tree.column("exam", width=400)
        self.uncat_tree.column("section", width=200)
        
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=self.uncat_tree.yview)
        self.uncat_tree.configure(yscrollcommand=scrollbar.set)
        
        self.uncat_tree.pack(side="top", fill="both", expand=True, padx=10, pady=5)
        scrollbar.place(relx=0.98, rely=0.15, relheight=0.6, anchor="ne")
        
        # Frame para agregar
        add_frame = ttk.Frame(parent)
        add_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(add_frame, text="Asignar categoría:").pack(side="left", padx=5)
        self.uncat_category_var = tk.StringVar()
        ttk.Combobox(add_frame, textvariable=self.uncat_category_var, width=20,
                    values=CATEGORY_ORDER[:-2]).pack(side="left", padx=5)
        
        ttk.Button(add_frame, text="➕ Agregar a Categorías por Examen", 
                  command=self.add_uncategorized_to_exam).pack(side="left", padx=5)
    
    def update_last_update_label(self):
        """Actualiza la etiqueta de última actualización"""
        ultima = self.exam_catalog.get("ultima_actualizacion")
        if ultima:
            self.last_update_var.set(f"Última actualización: {ultima}")
        else:
            self.last_update_var.set("Catálogo no descargado")
    
    def show_uncategorized(self):
        """Muestra los exámenes sin categorizar"""
        # Limpiar lista
        for item in self.uncat_tree.get_children():
            self.uncat_tree.delete(item)
        
        examenes_dict = self.exam_catalog.get("examenes", {})
        
        # Handle both old format (list) and new format (dict by section)
        if isinstance(examenes_dict, list):
            # Old format - just a list of exam names
            examenes_con_seccion = [(exam, "") for exam in examenes_dict]
        else:
            # New format - dict with section -> [exams]
            examenes_con_seccion = []
            for seccion, exams in examenes_dict.items():
                for exam in exams:
                    examenes_con_seccion.append((exam, seccion))
        
        if not examenes_con_seccion:
            messagebox.showinfo("Info", "No hay exámenes en el catálogo.\nHaga clic en 'Actualizar Catálogo de Exámenes' en la pestaña Descarga.")
            return
        
        # Obtener categorías configuradas
        exam_categories = self.exam_config.get("exam_categories", {})
        seccion_categories = self.exam_config.get("seccion_categories", {})
        
        # Encontrar sin categorizar
        uncategorized = []
        for exam, seccion in examenes_con_seccion:
            # Check if exam has category directly
            if exam in exam_categories:
                continue
            # Check if section has category
            if seccion and seccion in seccion_categories:
                continue
            uncategorized.append((exam, seccion))
        
        # Ordenar alfabéticamente
        uncategorized.sort(key=lambda x: x[0])
        
        # Agregar a la lista
        for exam, seccion in uncategorized:
            self.uncat_tree.insert("", "end", values=(exam, seccion))
        
        if not uncategorized:
            total = len(examenes_con_seccion)
            messagebox.showinfo("Completo", f"Todos los {total} exámenes tienen categoría asignada")
        else:
            total = len(examenes_con_seccion)
            messagebox.showinfo("Resultado", f"Se encontraron {len(uncategorized)} exámenes sin categorizar de {total} totales")
    
    def update_exam_catalog(self):
        """Descarga el catálogo de exámenes del sistema"""
        # Ejecutar en hilo separado
        thread = threading.Thread(target=self._download_exam_catalog)
        thread.daemon = True
        thread.start()
    
    def _download_exam_catalog(self):
        """Descarga el catálogo de exámenes por sección (ejecutar en hilo)"""
        try:
            url = self.config.get("General", "URLCatalogo", fallback="https://hjmvi.orion-labs.com/informes/catalogos")
            dropdown_id = self.config.get("Catalogo", "IdDropdownTipo", fallback="tipo")
            dropdown_value = self.config.get("Catalogo", "ValorExamenes", fallback="EXAMENES")
            button_id = self.config.get("Catalogo", "IdBotonGenerar", fallback="generar-informe-auditorias")
            
            browser_data_folder = self.base_dir / "browser_data"
            browser_data_folder.mkdir(exist_ok=True)
            
            downloads_folder = self.base_dir / self.config.get("Archivos", "CarpetaDescargas")
            downloads_folder.mkdir(exist_ok=True)
            
            # Lista de secciones a descargar
            secciones = [
                "Autoinmunes e Infecciosas",
                "Biología Molecular",
                "Bioquímica",
                "Citología",
                "Coagulación",
                "Coproanálisis",
                "Drogas y Fármacos",
                "Electrolitos",
                "Especiales",
                "Estudios de Alergias",
                "Estudios Hormonales",
                "Gases Arteriales",
                "Hematología",
                "Inmunohematología",
                "Inmunología",
                "Inmunoquímica Sanguínea",
                "Líquidos Biológicos",
                "Marcadores Coronarios",
                "Marcadores Tumorales",
                "Medicina Ocupacional",
                "Microbiología",
                "Plaquetas",
                "Química Clínica en Orina",
                "Serología",
                "Uroanálisis"
            ]
            
            self.root.after(0, lambda: self.log("🔄 Descargando catálogo de exámenes por sección..."))
            self.root.after(0, lambda: self.status_var.set("Descargando catálogo de exámenes..."))
            
            examenes_por_seccion = {}
            
            with sync_playwright() as p:
                context = p.chromium.launch_persistent_context(
                    user_data_dir=str(browser_data_folder),
                    headless=False,
                    channel="chrome",
                    accept_downloads=True
                )
                
                page = context.pages[0] if context.pages else context.new_page()
                page.goto(url, timeout=60000)
                
                # Esperar a que cargue
                page.wait_for_load_state("networkidle", timeout=10000)
                
                # Verificar login - esperar a que aparezca el dropdown de tipo
                max_wait = 120
                waited = 0
                while waited < max_wait:
                    try:
                        dropdown = page.query_selector(f"#{dropdown_id}")
                        if dropdown:
                            break
                    except:
                        pass
                    time.sleep(2)
                    waited += 2
                
                if waited >= max_wait:
                    self.root.after(0, lambda: messagebox.showerror("Error", "Timeout esperando la página. ¿Necesita iniciar sesión?"))
                    context.close()
                    return
                
                # Seleccionar "Exámenes" en el dropdown de Tipo
                dropdown = page.query_selector(f"#{dropdown_id}")
                if dropdown:
                    dropdown.select_option(value=dropdown_value)
                    time.sleep(1)
                
                total_secciones = len(secciones)
                secciones_descargadas = 0
                
                for seccion in secciones:
                    try:
                        self.root.after(0, lambda s=seccion, i=secciones_descargadas, t=total_secciones: 
                            self.status_var.set(f"Descargando {s} ({i+1}/{t})..."))
                        self.root.after(0, lambda i=secciones_descargadas, t=total_secciones: 
                            self.progress_var.set((i / t) * 100))
                        
                        # Hacer clic en el dropdown de secciones para abrirlo
                        secciones_input = page.query_selector("#secciones input.vs__search")
                        if secciones_input:
                            secciones_input.click()
                            time.sleep(0.5)
                            
                            # Escribir el nombre de la sección para filtrar
                            secciones_input.fill(seccion)
                            time.sleep(0.5)
                            
                            # Hacer clic en la opción que aparece
                            option = page.query_selector(f".vs__dropdown-menu .vs__dropdown-option")
                            if option:
                                option.click()
                                time.sleep(0.3)
                        
                        # Descargar
                        try:
                            with page.expect_download(timeout=15000) as download_info:
                                generar_btn = page.query_selector(f"#{button_id}")
                                if generar_btn:
                                    generar_btn.click()
                                else:
                                    page.click("button:has-text('Generar informe')")
                            
                            download = download_info.value
                            temp_path = downloads_folder / f"catalogo_{seccion.replace(' ', '_')}.xlsx"
                            download.save_as(temp_path)
                            
                            # Leer el archivo Excel y extraer exámenes
                            df = pd.read_excel(temp_path, skiprows=3)
                            
                            # Buscar la columna de exámenes
                            exam_col = None
                            for col in df.columns:
                                if 'examen' in col.lower():
                                    exam_col = col
                                    break
                            
                            if exam_col is None and len(df.columns) > 0:
                                exam_col = df.columns[0]
                            
                            if exam_col:
                                examenes = df[exam_col].dropna().astype(str).tolist()
                                examenes = [e.strip() for e in examenes 
                                           if e.strip() and not e.startswith('Hospital') and not e.startswith('Generado')]
                                examenes_por_seccion[seccion] = examenes
                                self.root.after(0, lambda s=seccion, n=len(examenes): 
                                    self.log(f"   ✅ {s}: {n} exámenes"))
                            
                            # Eliminar archivo temporal
                            try:
                                temp_path.unlink()
                            except:
                                pass
                            
                        except Exception as e:
                            self.root.after(0, lambda s=seccion, err=str(e): 
                                self.log(f"   ⚠️ {s}: Error - {err}"))
                        
                        # Limpiar la selección de sección para la siguiente iteración
                        # Hacer clic en el botón X para deseleccionar
                        deselect_btns = page.query_selector_all("#secciones .vs__deselect")
                        for btn in deselect_btns:
                            try:
                                btn.click()
                                time.sleep(0.2)
                            except:
                                pass
                        
                        secciones_descargadas += 1
                        
                    except Exception as e:
                        self.root.after(0, lambda s=seccion, err=str(e): 
                            self.log(f"   ❌ {s}: Error - {err}"))
                        continue
                
                context.close()
            
            # Guardar catálogo con estructura por sección
            self.save_exam_catalog(examenes_por_seccion)
            self.exam_catalog = self.load_exam_catalog()
            
            # Contar total de exámenes
            total_examenes = sum(len(exams) for exams in examenes_por_seccion.values())
            
            # Actualizar UI
            self.root.after(0, self.update_last_update_label)
            self.root.after(0, self.update_exam_combobox)
            self.root.after(0, self.update_catalog_info)
            self.root.after(0, lambda: self.progress_var.set(100))
            self.root.after(0, lambda: self.status_var.set("Catálogo actualizado"))
            
            self.root.after(0, lambda: self.log(f"✅ Catálogo actualizado: {total_examenes} exámenes en {len(examenes_por_seccion)} secciones"))
            self.root.after(0, lambda: messagebox.showinfo("Completado", 
                f"Catálogo actualizado:\n{total_examenes} exámenes en {len(examenes_por_seccion)} secciones"))
            
        except Exception as e:
            self.root.after(0, lambda: self.log(f"❌ Error: {str(e)}"))
            self.root.after(0, lambda: messagebox.showerror("Error", f"Error descargando catálogo: {str(e)}"))
    
    # ============================================
    # FUNCIONES DE CONFIGURACIÓN
    # ============================================
    
    def save_web_config(self):
        """Guarda la configuración web"""
        self.config.set("General", "URL", self.url_var.get())
        self.config.set("Informe", "IdDropdownAgrupar", self.dropdown_id_var.get())
        self.config.set("Informe", "ValorAgrupacion", self.dropdown_value_var.get())
        self.config.set("Informe", "IdFechaDesde", self.fecha_desde_var.get())
        self.config.set("Informe", "IdFechaHasta", self.fecha_hasta_var.get())
        self.config.set("Archivos", "CarpetaDescargas", self.downloads_folder_var.get())
        self.config.set("Archivos", "ArchivoSalida", self.output_file_var.get())
        
        # Guardar configuración de columnas merge
        if not self.config.has_section("ColumnasMerge"):
            self.config.add_section("ColumnasMerge")
        self.config.set("ColumnasMerge", "HospitalizacionTotal", self.hosp_total_var.get())
        self.config.set("ColumnasMerge", "ConsultaExternaTotal", self.cons_total_var.get())
        self.config.set("ColumnasMerge", "Emergencia", self.emerg_var.get())
        
        self.save_config()
        messagebox.showinfo("Guardado", "Configuración guardada correctamente")
    
    def refresh_multipliers_list(self):
        """Actualiza la lista de multiplicadores"""
        for item in self.mult_tree.get_children():
            self.mult_tree.delete(item)
        
        for exam, mult in self.exam_config.get("multipliers", {}).items():
            self.mult_tree.insert("", "end", values=(exam, mult))
    
    def add_multiplier(self):
        """Agrega un nuevo multiplicador"""
        exam = self.new_exam_var.get().strip()
        try:
            mult = int(self.new_mult_var.get())
        except:
            messagebox.showerror("Error", "El multiplicador debe ser un número entero")
            return
        
        if not exam:
            messagebox.showerror("Error", "Ingrese el nombre del examen")
            return
        
        if "multipliers" not in self.exam_config:
            self.exam_config["multipliers"] = {}
        
        self.exam_config["multipliers"][exam] = mult
        self.refresh_multipliers_list()
        self.new_exam_var.set("")
    
    def delete_multiplier(self):
        """Elimina el multiplicador seleccionado"""
        selected = self.mult_tree.selection()
        if not selected:
            messagebox.showwarning("Aviso", "Seleccione un examen para eliminar")
            return
        
        exam = self.mult_tree.item(selected[0])['values'][0]
        if exam in self.exam_config.get("multipliers", {}):
            del self.exam_config["multipliers"][exam]
        
        self.refresh_multipliers_list()
    
    def save_multipliers(self):
        """Guarda los multiplicadores"""
        try:
            self.exam_config["cultivo_multiplier"] = int(self.cultivo_mult_var.get())
        except:
            messagebox.showerror("Error", "El multiplicador de cultivos debe ser un número")
            return
        
        self.save_exam_config()
        messagebox.showinfo("Guardado", "Multiplicadores guardados correctamente")
    
    def add_exam_category(self):
        """Agrega una categoría por examen"""
        name = self.exam_name_var.get().strip()
        cat = self.exam_cat_var.get().strip()
        
        if not name or not cat:
            messagebox.showerror("Error", "Complete nombre y categoría")
            return
        
        if "exam_categories" not in self.exam_config:
            self.exam_config["exam_categories"] = {}
        
        self.exam_config["exam_categories"][name] = cat
        self.exam_cat_tree.insert("", "end", values=(name, cat))
        self.exam_name_var.set("")
    
    def delete_exam_category(self):
        """Elimina la categoría de examen seleccionada"""
        selected = self.exam_cat_tree.selection()
        if not selected:
            return
        
        name = self.exam_cat_tree.item(selected[0])['values'][0]
        if name in self.exam_config.get("exam_categories", {}):
            del self.exam_config["exam_categories"][name]
        
        self.exam_cat_tree.delete(selected[0])
    
    def add_section_category(self):
        """Agrega una categoría por sección"""
        name = self.section_name_var.get().strip()
        cat = self.section_cat_var.get().strip()
        
        if not name or not cat:
            messagebox.showerror("Error", "Complete nombre y categoría")
            return
        
        if "seccion_categories" not in self.exam_config:
            self.exam_config["seccion_categories"] = {}
        
        self.exam_config["seccion_categories"][name] = cat
        self.section_cat_tree.insert("", "end", values=(name, cat))
        self.section_name_var.set("")
    
    def delete_section_category(self):
        """Elimina la categoría de sección seleccionada"""
        selected = self.section_cat_tree.selection()
        if not selected:
            return
        
        name = self.section_cat_tree.item(selected[0])['values'][0]
        if name in self.exam_config.get("seccion_categories", {}):
            del self.exam_config["seccion_categories"][name]
        
        self.section_cat_tree.delete(selected[0])
    
    def save_categories(self):
        """Guarda todas las categorías"""
        self.save_exam_config()
        messagebox.showinfo("Guardado", "Categorías guardadas correctamente")
    
    def add_uncategorized_to_exam(self):
        """Agrega el examen sin categorizar a categorías por examen"""
        selected = self.uncat_tree.selection()
        if not selected:
            messagebox.showwarning("Aviso", "Seleccione un examen")
            return
        
        cat = self.uncat_category_var.get()
        if not cat:
            messagebox.showwarning("Aviso", "Seleccione una categoría")
            return
        
        exam = self.uncat_tree.item(selected[0])['values'][0]
        
        if "exam_categories" not in self.exam_config:
            self.exam_config["exam_categories"] = {}
        
        self.exam_config["exam_categories"][exam] = cat
        self.save_exam_config()
        
        # Actualizar listas
        self.exam_cat_tree.insert("", "end", values=(exam, cat))
        self.uncat_tree.delete(selected[0])
        
        messagebox.showinfo("Agregado", f"'{exam}' agregado a categoría '{cat}'")
    
    def auto_categorize_from_catalog(self):
        """Auto-categoriza exámenes basándose en la sección del catálogo y seccion_categories"""
        examenes_dict = self.exam_catalog.get("examenes", {})
        
        if not isinstance(examenes_dict, dict) or not examenes_dict:
            messagebox.showwarning("Aviso", "No hay catálogo con información de secciones.\n"
                                  "Haga clic en 'Actualizar Catálogo de Exámenes' primero.")
            return
        
        seccion_categories = self.exam_config.get("seccion_categories", {})
        exam_categories = self.exam_config.get("exam_categories", {})
        
        nuevas_categorias = 0
        secciones_sin_categoria = set()
        
        for seccion, exams in examenes_dict.items():
            # Check if this section has a category mapping
            if seccion in seccion_categories:
                categoria = seccion_categories[seccion]
                for exam in exams:
                    # Only add if exam doesn't already have a category
                    if exam not in exam_categories:
                        exam_categories[exam] = categoria
                        nuevas_categorias += 1
            else:
                # Section doesn't have a category mapping
                secciones_sin_categoria.add(seccion)
        
        # Save updates
        self.exam_config["exam_categories"] = exam_categories
        self.save_exam_config()
        
        # Refresh the exam categories tree
        for item in self.exam_cat_tree.get_children():
            self.exam_cat_tree.delete(item)
        for name, cat in exam_categories.items():
            self.exam_cat_tree.insert("", "end", values=(name, cat))
        
        # Show results
        msg = f"Auto-categorización completada:\n\n"
        msg += f"✅ {nuevas_categorias} exámenes categorizados automáticamente\n"
        
        if secciones_sin_categoria:
            msg += f"\n⚠️ Las siguientes secciones del catálogo no tienen\n"
            msg += f"una categoría asignada en 'Por Sección':\n"
            for sec in sorted(secciones_sin_categoria):
                msg += f"  • {sec}\n"
            msg += f"\nAgregue estas secciones en la pestaña 'Por Sección'\n"
            msg += f"para poder auto-categorizar sus exámenes."
        
        messagebox.showinfo("Auto-Categorización", msg)
        
        # Refresh uncategorized list
        self.show_uncategorized()
    
    # ============================================
    # FUNCIONES DE CONTROL DE PROCESO
    # ============================================
    
    def log(self, message: str):
        """Agrega mensaje al log"""
        timestamp = datetime.now().strftime("[%H:%M:%S]")
        self.log_text.insert("end", f"{timestamp} {message}\n")
        self.log_text.see("end")
        self.root.update_idletasks()
    
    def update_progress(self, current: int, total: int, status: str = ""):
        """Actualiza la barra de progreso"""
        percentage = (current / total * 100) if total > 0 else 0
        self.progress_var.set(percentage)
        if status:
            self.status_var.set(status)
        self.root.update_idletasks()
    
    def start_process(self):
        """Inicia el proceso de descarga"""
        if self.is_running:
            return
        
        self.is_running = True
        self.should_stop = False
        
        self.start_button.config(state="disabled")
        self.stop_button.config(state="normal")
        self.open_excel_button.config(state="disabled")
        self.log_text.delete(1.0, "end")
        
        # Ejecutar en hilo separado
        thread = threading.Thread(target=self.run_automation)
        thread.daemon = True
        thread.start()
    
    def stop_process(self):
        """Detiene el proceso"""
        self.should_stop = True
        self.log("⏹ Deteniendo proceso...")
    
    def finish_process(self, success: bool = True):
        """Finaliza el proceso y restaura la GUI"""
        self.is_running = False
        self.start_button.config(state="normal")
        self.stop_button.config(state="disabled")
        
        if success and self.output_file_path and self.output_file_path.exists():
            self.open_excel_button.config(state="normal")
            self.status_var.set("Completado - Haga clic en 'Abrir Excel' para ver resultados")
        else:
            self.status_var.set("Proceso finalizado")
    
    def open_excel(self):
        """Abre el archivo Excel generado"""
        if self.output_file_path and self.output_file_path.exists():
            try:
                os.startfile(self.output_file_path)
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo abrir el archivo: {e}")
    
    def open_folder(self):
        """Abre la carpeta de descargas"""
        downloads_folder = self.base_dir / self.config.get("Archivos", "CarpetaDescargas")
        downloads_folder.mkdir(exist_ok=True)
        try:
            os.startfile(downloads_folder)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir la carpeta: {e}")
    
    def recalculate_excel(self):
        """Recalcula el Excel con los archivos descargados existentes"""
        downloads_folder = self.base_dir / self.config.get("Archivos", "CarpetaDescargas")
        
        if not downloads_folder.exists():
            messagebox.showwarning("Aviso", "No existe la carpeta de descargas.\nPrimero descargue los datos.")
            return
        
        # Verificar si hay archivos Excel
        pattern = re.compile(r'^\d.*\.xlsx$')
        xlsx_files = [f for f in downloads_folder.glob('*.xlsx') if pattern.match(f.name)]
        
        if not xlsx_files:
            messagebox.showwarning("Aviso", "No hay archivos Excel en la carpeta de descargas.\nPrimero descargue los datos.")
            return
        
        # Confirmar
        result = messagebox.askyesno("Recalcular Excel", 
            f"Se encontraron {len(xlsx_files)} archivos en la carpeta de descargas.\n\n"
            "¿Desea regenerar el archivo Excel con la configuración actual de multiplicadores y categorías?")
        
        if not result:
            return
        
        # Recalcular
        self.output_file_path = self.base_dir / self.config.get("Archivos", "ArchivoSalida")
        
        self.log("=" * 50)
        self.log("🔄 Recalculando Excel...")
        self.log(f"   Archivos encontrados: {len(xlsx_files)}")
        
        try:
            self.process_excel_files(downloads_folder)
            self.log("✅ Excel regenerado correctamente")
            self.open_excel_button.config(state="normal")
            messagebox.showinfo("Completado", "Excel regenerado correctamente")
        except Exception as e:
            self.log(f"❌ Error: {str(e)}")
            messagebox.showerror("Error", f"Error al recalcular: {str(e)}")
    
    def check_file_locked(self, filepath: Path) -> bool:
        """Verifica si un archivo está bloqueado (abierto en otra aplicación)"""
        if not filepath.exists():
            return False
        
        try:
            with open(filepath, 'a'):
                pass
            return False
        except (IOError, PermissionError):
            return True
    
    # ============================================
    # AUTOMATIZACIÓN PRINCIPAL
    # ============================================
    
    def run_automation(self):
        """Ejecuta el proceso de automatización"""
        try:
            # Obtener parámetros de fechas
            start_date = self.get_start_date()
            end_date = self.get_end_date()
            headless = self.headless_var.get()
            
            url = self.config.get("General", "URL")
            
            dropdown_id = self.config.get("Informe", "IdDropdownAgrupar", fallback="agrupar-por")
            dropdown_value = self.config.get("Informe", "ValorAgrupacion", fallback="SECCION_TIPO_ATENCION")
            id_fecha_desde = self.config.get("Informe", "IdFechaDesde")
            id_fecha_hasta = self.config.get("Informe", "IdFechaHasta")
            
            downloads_folder = self.base_dir / self.config.get("Archivos", "CarpetaDescargas")
            downloads_folder.mkdir(exist_ok=True)
            
            self.output_file_path = self.base_dir / self.config.get("Archivos", "ArchivoSalida")
            
            # Verificar si el archivo de salida está bloqueado
            if self.check_file_locked(self.output_file_path):
                self.log("⚠️ El archivo de salida está abierto en Excel")
                self.log("   Por favor, cierre Excel antes de continuar")
                
                # Mostrar diálogo
                self.root.after(0, lambda: messagebox.showwarning(
                    "Archivo Bloqueado",
                    f"El archivo '{self.output_file_path.name}' está abierto en otra aplicación.\n\n"
                    "Por favor, cierre Excel y haga clic en 'Iniciar Descarga' nuevamente."
                ))
                self.finish_process(success=False)
                return
            
            browser_data_folder = self.base_dir / "browser_data"
            browser_data_folder.mkdir(exist_ok=True)
            
            # Calcular total de días
            total_days = (end_date - start_date).days + 1
            
            self.log("=" * 50)
            self.log("🚀 Iniciando automatización...")
            self.log(f"📅 Rango: {start_date.strftime('%Y-%m-%d')} al {end_date.strftime('%Y-%m-%d')}")
            self.log(f"📁 Carpeta de descargas: {downloads_folder}")
            self.log("=" * 50)
            
            with sync_playwright() as p:
                self.log("🌐 Iniciando navegador Chrome...")
                self.log("   💾 Los datos de sesión se guardarán para futuros usos")
                
                context = p.chromium.launch_persistent_context(
                    user_data_dir=str(browser_data_folder),
                    headless=headless,
                    channel="chrome",
                    accept_downloads=True
                )
                
                page = context.pages[0] if context.pages else context.new_page()
                
                self.log(f"📄 Navegando a {url}")
                page.goto(url, timeout=60000)
                page.wait_for_load_state("networkidle", timeout=30000)
                
                # Verificar login
                max_login_wait = 300
                login_check_interval = 2
                waited_time = 0
                logged_in = False
                first_message = True
                
                while not logged_in:
                    if self.should_stop:
                        context.close()
                        self.finish_process(success=False)
                        return
                    
                    try:
                        page.wait_for_load_state("domcontentloaded", timeout=2000)
                    except:
                        pass
                    
                    try:
                        generar_btn = page.query_selector("button:has-text('Generar informe')")
                        if generar_btn:
                            logged_in = True
                            self.log("✅ Sesión detectada correctamente!")
                            break
                    except:
                        pass
                    
                    if first_message:
                        self.log("🔐 No se detectó sesión activa.")
                        self.log("   👉 Por favor, inicie sesión en la ventana del navegador...")
                        self.log(f"   ⏳ Esperando inicio de sesión (máximo {max_login_wait // 60} minutos)...")
                        first_message = False
                    
                    time.sleep(login_check_interval)
                    waited_time += login_check_interval
                    
                    if waited_time % 10 == 0:
                        self.log(f"   ⏳ Esperando... ({waited_time}s)")
                    
                    if waited_time >= max_login_wait:
                        self.log("❌ Tiempo de espera agotado. No se detectó inicio de sesión.")
                        context.close()
                        self.finish_process(success=False)
                        return
                
                # Estabilizar página
                try:
                    page.wait_for_load_state("networkidle", timeout=5000)
                except:
                    pass
                time.sleep(1)
                
                self.log("🚀 Comenzando descargas...")
                
                downloaded_files = []
                
                # Iterar por cada día en el rango
                current = start_date
                day_index = 0
                while current <= end_date:
                    if self.should_stop:
                        self.log("⏹ Proceso detenido por el usuario")
                        break
                    
                    current_date_str = current.strftime('%Y-%m-%d')
                    self.update_progress(day_index, total_days, f"Descargando {current_date_str}...")
                    self.log(f"📥 Procesando día {day_index + 1}/{total_days}: {current_date_str}")
                    
                    try:
                        # Establecer fechas - Ctrl+A para seleccionar todo, luego escribir
                        fecha_desde = page.locator(f"#{id_fecha_desde}")
                        fecha_hasta = page.locator(f"#{id_fecha_hasta}")
                        
                        # Fecha desde: click, Ctrl+A, escribir, Tab
                        fecha_desde.click()
                        page.keyboard.press("Control+a")
                        page.keyboard.type(current_date_str, delay=20)
                        page.keyboard.press("Tab")
                        
                        # Fecha hasta: click, Ctrl+A, escribir, Tab
                        fecha_hasta.click()
                        page.keyboard.press("Control+a")
                        page.keyboard.type(current_date_str, delay=20)
                        page.keyboard.press("Tab")
                        
                        # Espera para que Vue procese
                        time.sleep(0.3)
                        
                        # Verificar que las fechas se establecieron correctamente
                        actual_desde = fecha_desde.input_value()
                        actual_hasta = fecha_hasta.input_value()
                        
                        if actual_desde != current_date_str or actual_hasta != current_date_str:
                            self.log(f"   ⚠️ Fechas no coinciden: desde={actual_desde}, hasta={actual_hasta}")
                        
                        # Configurar dropdown (solo primer día)
                        if day_index == 0:
                            self.log("   🔧 Configurando dropdown de agrupación...")
                            try:
                                dropdown = page.query_selector(f"#{dropdown_id}")
                                if dropdown:
                                    dropdown.select_option(value=dropdown_value)
                                    self.log(f"   ✅ Dropdown seleccionado: {dropdown_value}")
                                else:
                                    self.log(f"   ⚠️ No se encontró el dropdown #{dropdown_id}")
                            except Exception as e:
                                self.log(f"   ⚠️ Error seleccionando dropdown: {e}")
                        
                        # Hacer clic en "Generar informe" para refrescar datos
                        generar_btn = page.locator("button:has-text('Generar informe')").first
                        generar_btn.click()
                        
                        # Esperar a que aparezca el menú dropdown con Excel
                        try:
                            page.wait_for_selector("a:has-text('Excel')", timeout=15000)
                        except:
                            self.log(f"   ⚠️ No apareció el enlace Excel")
                        
                        time.sleep(0.2)
                        
                        # Buscar y hacer clic en Excel
                        excel_link = page.locator("a:has-text('Excel')").first
                        
                        if excel_link.count() > 0:
                            with page.expect_download(timeout=30000) as download_info:
                                excel_link.click()
                            
                            download = download_info.value
                            download_path = downloads_folder / f"{current_date_str}.xlsx"
                            download.save_as(download_path)
                            downloaded_files.append(download_path)
                            
                            self.log(f"   ✅ Guardado: {current_date_str}.xlsx")
                        else:
                            self.log(f"   ⚠️ No se encontró el enlace Excel para {current_date_str}")
                        
                    except PlaywrightTimeout:
                        self.log(f"   ⚠️ Timeout en {current_date_str}, continuando...")
                    except Exception as e:
                        self.log(f"   ❌ Error en {current_date_str}: {str(e)}")
                    
                    # Avanzar al siguiente día
                    current += timedelta(days=1)
                    day_index += 1
                
                context.close()
            
            if self.should_stop:
                self.finish_process(success=False)
                return
            
            # Procesar archivos
            self.update_progress(90, 100, "Procesando datos...")
            self.log("")
            self.log("=" * 50)
            self.log("📊 Procesando archivos descargados...")
            
            if downloaded_files:
                self.process_excel_files(downloads_folder)
                self.update_progress(100, 100, "¡Completado!")
                self.log("=" * 50)
                self.log(f"✅ ¡Proceso completado! Se procesaron {len(downloaded_files)} archivos.")
                self.finish_process(success=True)
            else:
                self.log("⚠️ No se descargaron archivos")
                self.finish_process(success=False)
                
        except Exception as e:
            self.log(f"❌ Error: {str(e)}")
            import traceback
            self.log(traceback.format_exc())
            self.finish_process(success=False)
    
    def process_excel_files(self, downloads_folder: Path):
        """Procesa los archivos Excel descargados"""
        output_file = self.output_file_path
        
        # Verificar si el archivo está bloqueado
        if self.check_file_locked(output_file):
            self.log("⚠️ El archivo de salida está abierto en Excel")
            self.log("   Por favor, cierre Excel antes de continuar")
            
            # Preguntar al usuario
            self.root.after(0, lambda: messagebox.showwarning(
                "Archivo Bloqueado",
                f"El archivo '{output_file.name}' está abierto en otra aplicación.\n\n"
                "Por favor, cierre Excel e intente procesar nuevamente."
            ))
            return
        
        pattern = re.compile(r'^\d.*\.xlsx$')
        xlsx_files = sorted([f for f in downloads_folder.glob('*.xlsx') if pattern.match(f.name)])
        
        if not xlsx_files:
            self.log("❌ No se encontraron archivos Excel para procesar")
            return
        
        self.log(f"   Encontrados {len(xlsx_files)} archivos")
        
        # Detectar estructura
        first_file = xlsx_files[0]
        sample_df = pd.read_excel(first_file, skiprows=4)
        self.log(f"   📋 Columnas detectadas: {list(sample_df.columns)}")
        
        has_patient_type_columns = any(col in sample_df.columns for col in ['Hospitalización', 'Emergencia', 'Consulta Externa'])
        has_simple_columns = 'Cant. Exámenes' in sample_df.columns
        
        if has_patient_type_columns:
            self.log("   ✅ Informe con desglose por tipo de atención detectado")
        elif has_simple_columns:
            self.log("   ⚠️ Informe simple detectado (sin desglose por tipo de atención)")
            self.log("   ℹ️  El dropdown 'Agrupar por' no se seleccionó correctamente")
        
        # Leer y combinar
        all_dataframes = []
        for filepath in xlsx_files:
            try:
                df = pd.read_excel(filepath, skiprows=4)
                
                if 'Sección' in df.columns:
                    df = df.rename(columns={'Sección': 'Seccion'})
                
                date_str = filepath.stem
                try:
                    df['date'] = pd.to_datetime(date_str, format='%Y-%m-%d')
                except:
                    df['date'] = date_str
                
                if has_simple_columns and not has_patient_type_columns:
                    if 'Cant. Exámenes' in df.columns:
                        df['Total'] = df['Cant. Exámenes']
                    for col in NUMERIC_COLUMNS:
                        if col not in df.columns:
                            df[col] = 0
                else:
                    for col in NUMERIC_COLUMNS:
                        if col not in df.columns:
                            df[col] = 0
                
                all_dataframes.append(df)
            except Exception as e:
                self.log(f"   ⚠️ Error leyendo {filepath.name}: {e}")
        
        if not all_dataframes:
            self.log("❌ No se pudieron leer los archivos")
            return
        
        combined_df = pd.concat(all_dataframes, ignore_index=True)
        datos_descargados = combined_df.copy()
        
        self.log(f"   📊 Total de registros combinados: {len(combined_df)}")
        
        # Aplicar multiplicadores
        multipliers = self.exam_config.get("multipliers", {})
        cultivo_mult = self.exam_config.get("cultivo_multiplier", 10)
        
        def get_multiplier(exam_name):
            if pd.isna(exam_name):
                return 1
            if "CULTIVO" in str(exam_name).upper():
                return cultivo_mult
            return multipliers.get(exam_name, 1)
        
        combined_df['multiplier'] = combined_df['Examen'].apply(get_multiplier)
        
        cols_to_multiply = ['Total'] + [col for col in NUMERIC_COLUMNS if col in combined_df.columns and col != 'Total']
        for col in cols_to_multiply:
            if col in combined_df.columns:
                combined_df[col] = pd.to_numeric(combined_df[col], errors='coerce').fillna(0) * combined_df['multiplier']
        
        # Filtrar
        combined_df = combined_df[combined_df['Examen'].notna()]
        combined_df = combined_df[~combined_df['Examen'].astype(str).str.contains('^Total órdenes|^Generado el|^$', na=False, regex=True)]
        combined_df = combined_df[combined_df['Seccion'].notna() & (combined_df['Seccion'] != '')]
        
        self.log(f"   📊 Registros después de filtrar: {len(combined_df)}")
        
        # Categorizar
        exam_categories = self.exam_config.get("exam_categories", {})
        seccion_categories = self.exam_config.get("seccion_categories", {})
        
        def get_category(row):
            examen = row.get('Examen', '')
            seccion = row.get('Seccion', '')
            if pd.isna(examen):
                return "Other"
            if examen in exam_categories:
                return exam_categories[examen]
            if seccion in seccion_categories:
                return seccion_categories[seccion]
            return "Other"
        
        combined_df['Category'] = combined_df.apply(get_category, axis=1)
        
        # Datos categorizados
        examenes_categorizados = combined_df.copy()
        cols = ['Seccion', 'Examen', 'multiplier', 'Category'] + \
               [c for c in combined_df.columns if c not in ['Seccion', 'Examen', 'multiplier', 'Category']]
        examenes_categorizados = examenes_categorizados[cols]
        
        # Determinar columnas para resumen
        if has_patient_type_columns:
            summary_cols = [col for col in NUMERIC_COLUMNS if col in combined_df.columns]
        else:
            summary_cols = ['Total'] if 'Total' in combined_df.columns else []
        
        summary_cols = list(dict.fromkeys(summary_cols))
        
        self.log(f"   📋 Columnas para resumen: {summary_cols}")
        
        if not summary_cols:
            self.log("   ⚠️ No se encontraron columnas numéricas para el resumen")
            summary_cols = ['Total']
            combined_df['Total'] = 0
        
        # Tabla resumen
        summary_table = combined_df.groupby(['Category', 'date'])[summary_cols].sum().reset_index()
        
        def safe_get_col(df, col):
            return df[col] if col in df.columns else 0
        
        # Obtener configuración de columnas a combinar
        hosp_cols = [c.strip() for c in self.config.get("ColumnasMerge", "HospitalizacionTotal", 
                    fallback="Hospitalización,URGENTE HOSPITALIZACION").split(",")]
        cons_cols = [c.strip() for c in self.config.get("ColumnasMerge", "ConsultaExternaTotal",
                    fallback="Consulta Externa,URGENTE CONSULTA EXTERNA,REFERENCIA,URGENTE REFERENCIA").split(",")]
        emerg_cols = [c.strip() for c in self.config.get("ColumnasMerge", "Emergencia",
                     fallback="Emergencia,Sin tipo atención").split(",")]
        
        self.log(f"   📊 Columnas combinadas:")
        self.log(f"      Hospitalización Total = {' + '.join(hosp_cols)}")
        self.log(f"      Consulta Externa Total = {' + '.join(cons_cols)}")
        self.log(f"      Emergencia = {' + '.join(emerg_cols)}")
        
        # Calcular columnas combinadas
        summary_table['Hospitalización Total'] = sum(safe_get_col(summary_table, col) for col in hosp_cols)
        summary_table['Consulta Externa Total'] = sum(safe_get_col(summary_table, col) for col in cons_cols)
        summary_table['Emergencia'] = sum(safe_get_col(summary_table, col) for col in emerg_cols)
        
        # Totales
        total_cols = list(set(
            [col for col in summary_cols if col in summary_table.columns] + 
            [col for col in ['Hospitalización Total', 'Consulta Externa Total', 'Emergencia', 'Total'] 
             if col in summary_table.columns]
        ))
        
        totals = summary_table.groupby('date', as_index=False)[total_cols].sum()
        totals['Category'] = 'TOTAL'
        
        summary_table = pd.concat([summary_table, totals], ignore_index=True, sort=False)
        
        # Ordenar
        summary_table['cat_order'] = summary_table['Category'].apply(
            lambda x: CATEGORY_ORDER.index(x) if x in CATEGORY_ORDER else 999
        )
        summary_table = summary_table.sort_values(['date', 'cat_order']).drop('cat_order', axis=1)
        
        # Reordenar columnas
        final_cols = ['Category', 'date', 'Hospitalización Total', 'Consulta Externa Total', 
                      'Emergencia', 'Total'] + \
                     [c for c in summary_table.columns if c not in 
                      ['Category', 'date', 'Hospitalización Total', 'Consulta Externa Total', 'Emergencia', 'Total']]
        summary_table = summary_table[[c for c in final_cols if c in summary_table.columns]]
        
        # Formatear fechas
        if 'date' in summary_table.columns:
            summary_table['date'] = pd.to_datetime(summary_table['date']).dt.strftime('%Y-%m-%d')
            summary_table = summary_table.rename(columns={'date': 'Fecha'})
        if 'date' in examenes_categorizados.columns:
            examenes_categorizados['date'] = pd.to_datetime(examenes_categorizados['date']).dt.strftime('%Y-%m-%d')
            examenes_categorizados = examenes_categorizados.rename(columns={'date': 'Fecha'})
        if 'date' in datos_descargados.columns:
            datos_descargados['date'] = pd.to_datetime(datos_descargados['date']).dt.strftime('%Y-%m-%d')
            datos_descargados = datos_descargados.rename(columns={'date': 'Fecha'})
        
        # Guardar Excel
        self.log(f"   Guardando {output_file.name}...")
        
        try:
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                summary_table.to_excel(writer, sheet_name='Estadistica Calculada', index=False)
                examenes_categorizados.to_excel(writer, sheet_name='Examenes Categorizados', index=False)
                datos_descargados.to_excel(writer, sheet_name='Datos Descargados', index=False)
                
                for sheet_name in writer.sheets:
                    ws = writer.sheets[sheet_name]
                    ws.freeze_panes = 'A2'
                    ws.auto_filter.ref = ws.dimensions
                    
                    for column in ws.columns:
                        max_length = 0
                        column_letter = column[0].column_letter
                        
                        for cell in column:
                            try:
                                cell_value = str(cell.value) if cell.value is not None else ""
                                cell_length = len(cell_value)
                                if cell_length > max_length:
                                    max_length = cell_length
                            except:
                                pass
                        
                        adjusted_width = min(max_length + 2, 50)
                        ws.column_dimensions[column_letter].width = adjusted_width
            
            self.log(f"   ✅ Archivo guardado: {output_file}")
            
        except PermissionError:
            self.log(f"   ❌ Error: No se puede escribir el archivo. ¿Está abierto en Excel?")
            self.root.after(0, lambda: messagebox.showerror(
                "Error de Escritura",
                f"No se puede escribir el archivo '{output_file.name}'.\n\n"
                "Por favor, cierre Excel e intente nuevamente."
            ))


def main():
    root = tk.Tk()
    
    style = ttk.Style()
    style.theme_use('clam')
    
    app = EstadisticaHospitalApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
