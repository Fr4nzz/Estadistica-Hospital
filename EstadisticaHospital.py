#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Estadística Hospital - Automatización de descarga y procesamiento de estadísticas
Versión 3.1 con configuración desde la interfaz gráfica

Autor: Automatizado con Claude
Fecha: 2025
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import configparser
import json
import os
import sys
import re
import threading
import subprocess
from pathlib import Path
from datetime import datetime
from typing import Optional
import time

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


# ============================================
# CONFIGURACIÓN POR DEFECTO
# ============================================

DEFAULT_CONFIG = {
    "General": {
        "URL": "https://hjmvi.orion-labs.com/informes/estadisticos",
        "TiempoEspera": "0",
        "TiempoCargaPagina": "5",
        "TimeoutDescarga": "15",
        "Headless": "false"
    },
    "Informe": {
        "IdDropdownAgrupar": "agrupar-por",
        "ValorAgrupacion": "SECCION_TIPO_ATENCION",
        "IdFechaDesde": "fecha-orden-desde",
        "IdFechaHasta": "fecha-orden-hasta"
    },
    "Archivos": {
        "CarpetaDescargas": "./ExcelsDescargados",
        "ArchivoSalida": "./Estadistica Hospital.xlsx"
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
        "Autoinmunes e Infecciosas": "Serologicos",
        "Drogas y Fármacos": "Serologicos",
        "Bioquímica": "Quimica sanguinea",
        "Coagulación": "Hematologico",
        "Coproanálisis": "Materias fecales",
        "Electrolitos": "Quimica sanguinea",
        "Estudios Hormonales": "Hormonales",
        "Gases Arteriales": "Quimica sanguinea",
        "Hematología": "Hematologico",
        "Inmunohematología": "Hematologico",
        "Inmunología": "Serologicos",
        "Líquidos Biológicos": "Bacteriológico",
        "Marcadores Tumorales": "Serologicos",
        "Microbiología": "Bacteriológico",
        "Química Clínica en Orina": "Orina",
        "Serología": "Serologicos",
        "Uroanálisis": "Orina"
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
        self.root.title("Estadística Hospital v3.1")
        self.root.geometry("900x700")
        self.root.minsize(800, 600)
        
        # Determinar directorio base
        if getattr(sys, 'frozen', False):
            self.base_dir = Path(sys.executable).parent
        else:
            self.base_dir = Path(__file__).parent
        
        # Cargar configuraciones
        self.config = self.load_config()
        self.exam_config = self.load_exam_config()
        
        # Variables de control
        self.should_stop = False
        self.is_running = False
        self.output_file_path = None
        
        # Crear interfaz
        self.create_notebook()
        self.create_main_tab()
        self.create_config_tab()
        self.create_exams_tab()
        self.create_categories_tab()
        
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
        params_frame = ttk.LabelFrame(self.main_frame, text="Parámetros de Descarga", padding=10)
        params_frame.pack(fill="x", padx=10, pady=5)
        
        # Año
        ttk.Label(params_frame, text="Año:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.year_var = tk.StringVar(value=str(datetime.now().year))
        ttk.Entry(params_frame, textvariable=self.year_var, width=10).grid(row=0, column=1, sticky="w", padx=5, pady=2)
        
        # Mes
        ttk.Label(params_frame, text="Mes:").grid(row=0, column=2, sticky="w", padx=5, pady=2)
        self.month_var = tk.StringVar(value=str(datetime.now().month))
        month_combo = ttk.Combobox(params_frame, textvariable=self.month_var, width=8, 
                                   values=[str(i) for i in range(1, 13)])
        month_combo.grid(row=0, column=3, sticky="w", padx=5, pady=2)
        
        # Día inicial
        ttk.Label(params_frame, text="Día inicial:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.start_day_var = tk.StringVar(value="1")
        ttk.Entry(params_frame, textvariable=self.start_day_var, width=10).grid(row=1, column=1, sticky="w", padx=5, pady=2)
        
        # Día final
        ttk.Label(params_frame, text="Día final:").grid(row=1, column=2, sticky="w", padx=5, pady=2)
        self.end_day_var = tk.StringVar(value=str(datetime.now().day))
        ttk.Entry(params_frame, textvariable=self.end_day_var, width=10).grid(row=1, column=3, sticky="w", padx=5, pady=2)
        
        # Headless mode
        self.headless_var = tk.BooleanVar(value=self.config.get("General", "Headless", fallback="false").lower() == "true")
        ttk.Checkbutton(params_frame, text="Modo oculto (sin ventana del navegador)", 
                       variable=self.headless_var).grid(row=2, column=0, columnspan=4, sticky="w", padx=5, pady=5)
        
        # Frame de botones
        buttons_frame = ttk.Frame(self.main_frame)
        buttons_frame.pack(fill="x", padx=10, pady=5)
        
        self.start_button = ttk.Button(buttons_frame, text="▶ Iniciar Descarga", command=self.start_process)
        self.start_button.pack(side="left", padx=5)
        
        self.stop_button = ttk.Button(buttons_frame, text="⏹ Detener", command=self.stop_process, state="disabled")
        self.stop_button.pack(side="left", padx=5)
        
        self.open_excel_button = ttk.Button(buttons_frame, text="📊 Abrir Excel", command=self.open_excel, state="disabled")
        self.open_excel_button.pack(side="left", padx=5)
        
        self.open_folder_button = ttk.Button(buttons_frame, text="📁 Abrir Carpeta", command=self.open_folder)
        self.open_folder_button.pack(side="left", padx=5)
        
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
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=15, font=("Consolas", 9))
        self.log_text.pack(fill="both", expand=True)
    
    def create_config_tab(self):
        """Crea la pestaña de configuración web"""
        # Frame de URL y navegación
        web_frame = ttk.LabelFrame(self.config_frame, text="Configuración del Sitio Web", padding=10)
        web_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(web_frame, text="URL del sitio:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.url_var = tk.StringVar(value=self.config.get("General", "URL"))
        ttk.Entry(web_frame, textvariable=self.url_var, width=60).grid(row=0, column=1, sticky="ew", padx=5, pady=2)
        
        # Frame de elementos del formulario
        form_frame = ttk.LabelFrame(self.config_frame, text="Elementos del Formulario (IDs HTML)", padding=10)
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
        
        # Frame de tiempos
        timing_frame = ttk.LabelFrame(self.config_frame, text="Tiempos de Espera (segundos)", padding=10)
        timing_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(timing_frame, text="Entre descargas:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.wait_time_var = tk.StringVar(value=self.config.get("General", "TiempoEspera"))
        ttk.Entry(timing_frame, textvariable=self.wait_time_var, width=10).grid(row=0, column=1, sticky="w", padx=5, pady=2)
        
        ttk.Label(timing_frame, text="Timeout de descarga:").grid(row=0, column=2, sticky="w", padx=5, pady=2)
        self.timeout_var = tk.StringVar(value=self.config.get("General", "TimeoutDescarga"))
        ttk.Entry(timing_frame, textvariable=self.timeout_var, width=10).grid(row=0, column=3, sticky="w", padx=5, pady=2)
        
        ttk.Label(timing_frame, text="Carga de página:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.page_load_var = tk.StringVar(value=self.config.get("General", "TiempoCargaPagina"))
        ttk.Entry(timing_frame, textvariable=self.page_load_var, width=10).grid(row=1, column=1, sticky="w", padx=5, pady=2)
        
        # Frame de archivos
        files_frame = ttk.LabelFrame(self.config_frame, text="Rutas de Archivos", padding=10)
        files_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(files_frame, text="Carpeta de descargas:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.downloads_folder_var = tk.StringVar(value=self.config.get("Archivos", "CarpetaDescargas"))
        ttk.Entry(files_frame, textvariable=self.downloads_folder_var, width=40).grid(row=0, column=1, sticky="w", padx=5, pady=2)
        
        ttk.Label(files_frame, text="Archivo de salida:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.output_file_var = tk.StringVar(value=self.config.get("Archivos", "ArchivoSalida"))
        ttk.Entry(files_frame, textvariable=self.output_file_var, width=40).grid(row=1, column=1, sticky="w", padx=5, pady=2)
        
        # Botón guardar
        ttk.Button(self.config_frame, text="💾 Guardar Configuración", 
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
        
        # Frame de edición
        edit_frame = ttk.Frame(self.exams_frame)
        edit_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(edit_frame, text="Examen:").pack(side="left", padx=5)
        self.new_exam_var = tk.StringVar()
        ttk.Entry(edit_frame, textvariable=self.new_exam_var, width=40).pack(side="left", padx=5)
        
        ttk.Label(edit_frame, text="Multiplicador:").pack(side="left", padx=5)
        self.new_mult_var = tk.StringVar(value="1")
        ttk.Entry(edit_frame, textvariable=self.new_mult_var, width=10).pack(side="left", padx=5)
        
        ttk.Button(edit_frame, text="➕ Agregar", command=self.add_multiplier).pack(side="left", padx=5)
        ttk.Button(edit_frame, text="🗑️ Eliminar", command=self.delete_multiplier).pack(side="left", padx=5)
        
        # Botón guardar
        ttk.Button(self.exams_frame, text="💾 Guardar Multiplicadores", 
                  command=self.save_multipliers).pack(pady=10)
    
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
        info_label = ttk.Label(parent, text="Exámenes encontrados en los archivos descargados que no tienen categoría asignada.\n"
                              "Haga clic en 'Escanear' para analizar los archivos descargados.")
        info_label.pack(padx=10, pady=5, anchor="w")
        
        # Botón escanear
        scan_frame = ttk.Frame(parent)
        scan_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Button(scan_frame, text="🔍 Escanear Archivos Descargados", 
                  command=self.scan_uncategorized).pack(side="left", padx=5)
        
        # Treeview
        columns = ("exam", "section", "count")
        self.uncat_tree = ttk.Treeview(parent, columns=columns, show="headings", height=10)
        self.uncat_tree.heading("exam", text="Examen")
        self.uncat_tree.heading("section", text="Sección")
        self.uncat_tree.heading("count", text="Cantidad")
        self.uncat_tree.column("exam", width=300)
        self.uncat_tree.column("section", width=150)
        self.uncat_tree.column("count", width=80)
        
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
        ttk.Button(add_frame, text="➕ Agregar a Categorías por Sección", 
                  command=self.add_uncategorized_to_section).pack(side="left", padx=5)
    
    # ============================================
    # FUNCIONES DE CONFIGURACIÓN
    # ============================================
    
    def save_web_config(self):
        """Guarda la configuración web"""
        self.config.set("General", "URL", self.url_var.get())
        self.config.set("General", "TiempoEspera", self.wait_time_var.get())
        self.config.set("General", "TimeoutDescarga", self.timeout_var.get())
        self.config.set("General", "TiempoCargaPagina", self.page_load_var.get())
        self.config.set("Informe", "IdDropdownAgrupar", self.dropdown_id_var.get())
        self.config.set("Informe", "ValorAgrupacion", self.dropdown_value_var.get())
        self.config.set("Informe", "IdFechaDesde", self.fecha_desde_var.get())
        self.config.set("Informe", "IdFechaHasta", self.fecha_hasta_var.get())
        self.config.set("Archivos", "CarpetaDescargas", self.downloads_folder_var.get())
        self.config.set("Archivos", "ArchivoSalida", self.output_file_var.get())
        
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
    
    def scan_uncategorized(self):
        """Escanea los archivos descargados para encontrar exámenes sin categorizar"""
        downloads_folder = self.base_dir / self.config.get("Archivos", "CarpetaDescargas")
        
        if not downloads_folder.exists():
            messagebox.showwarning("Aviso", "No existe la carpeta de descargas")
            return
        
        # Limpiar lista
        for item in self.uncat_tree.get_children():
            self.uncat_tree.delete(item)
        
        # Leer archivos
        xlsx_files = list(downloads_folder.glob('*.xlsx'))
        if not xlsx_files:
            messagebox.showinfo("Info", "No hay archivos Excel en la carpeta de descargas")
            return
        
        exam_counts = {}
        exam_sections = {}
        
        for filepath in xlsx_files:
            try:
                df = pd.read_excel(filepath, skiprows=4)
                for _, row in df.iterrows():
                    exam = row.get('Examen', '')
                    section = row.get('Sección', row.get('Seccion', ''))
                    
                    if pd.isna(exam) or not exam or str(exam).startswith('Total'):
                        continue
                    
                    exam_counts[exam] = exam_counts.get(exam, 0) + 1
                    exam_sections[exam] = section
            except:
                continue
        
        # Filtrar los que no tienen categoría
        exam_categories = self.exam_config.get("exam_categories", {})
        seccion_categories = self.exam_config.get("seccion_categories", {})
        
        uncategorized = []
        for exam, count in exam_counts.items():
            section = exam_sections.get(exam, '')
            
            # Verificar si tiene categoría
            if exam in exam_categories:
                continue
            if section in seccion_categories:
                continue
            
            uncategorized.append((exam, section, count))
        
        # Ordenar por cantidad descendente
        uncategorized.sort(key=lambda x: x[2], reverse=True)
        
        # Agregar a la lista
        for exam, section, count in uncategorized:
            self.uncat_tree.insert("", "end", values=(exam, section, count))
        
        if not uncategorized:
            messagebox.showinfo("Completo", "Todos los exámenes tienen categoría asignada")
        else:
            messagebox.showinfo("Resultado", f"Se encontraron {len(uncategorized)} exámenes sin categorizar")
    
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
    
    def add_uncategorized_to_section(self):
        """Agrega la sección sin categorizar a categorías por sección"""
        selected = self.uncat_tree.selection()
        if not selected:
            messagebox.showwarning("Aviso", "Seleccione un examen")
            return
        
        cat = self.uncat_category_var.get()
        if not cat:
            messagebox.showwarning("Aviso", "Seleccione una categoría")
            return
        
        section = self.uncat_tree.item(selected[0])['values'][1]
        
        if not section:
            messagebox.showwarning("Aviso", "Este examen no tiene sección definida")
            return
        
        if "seccion_categories" not in self.exam_config:
            self.exam_config["seccion_categories"] = {}
        
        self.exam_config["seccion_categories"][section] = cat
        self.save_exam_config()
        
        # Actualizar lista de secciones
        self.section_cat_tree.insert("", "end", values=(section, cat))
        
        # Remover todos los exámenes de esa sección de la lista de sin categorizar
        to_remove = []
        for item in self.uncat_tree.get_children():
            item_section = self.uncat_tree.item(item)['values'][1]
            if item_section == section:
                to_remove.append(item)
        
        for item in to_remove:
            self.uncat_tree.delete(item)
        
        messagebox.showinfo("Agregado", f"Sección '{section}' agregada a categoría '{cat}'")
    
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
            # Obtener parámetros
            year = int(self.year_var.get())
            month = int(self.month_var.get())
            start_day = int(self.start_day_var.get())
            end_day = int(self.end_day_var.get())
            headless = self.headless_var.get()
            wait_time = float(self.config.get("General", "TiempoEspera", fallback="0"))
            
            url = self.config.get("General", "URL")
            page_load_time = int(self.config.get("General", "TiempoCargaPagina"))
            download_timeout = int(self.config.get("General", "TimeoutDescarga", fallback="15")) * 1000
            
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
            
            total_days = end_day - start_day + 1
            
            self.log("=" * 50)
            self.log("🚀 Iniciando automatización...")
            self.log(f"📅 Rango: {year}-{month:02d}-{start_day:02d} al {year}-{month:02d}-{end_day:02d}")
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
                page.wait_for_load_state("networkidle", timeout=page_load_time * 1000)
                
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
                    
                    try:
                        time.sleep(login_check_interval)
                    except:
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
                
                for day_index, current_day in enumerate(range(start_day, end_day + 1)):
                    if self.should_stop:
                        self.log("⏹ Proceso detenido por el usuario")
                        break
                    
                    current_date = f"{year}-{month:02d}-{current_day:02d}"
                    self.update_progress(day_index, total_days, f"Descargando {current_date}...")
                    self.log(f"📥 Procesando día {current_day}/{end_day}: {current_date}")
                    
                    try:
                        # Establecer fechas
                        page.evaluate(f"""
                            document.getElementById('{id_fecha_desde}').value = '{current_date}';
                            document.getElementById('{id_fecha_hasta}').value = '{current_date}';
                            document.getElementById('{id_fecha_desde}').dispatchEvent(new Event('input', {{ bubbles: true }}));
                            document.getElementById('{id_fecha_hasta}').dispatchEvent(new Event('input', {{ bubbles: true }}));
                        """)
                        
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
                        
                        # Buscar Excel link
                        excel_link = page.query_selector("a:has-text('Excel'):visible")
                        
                        if not excel_link:
                            generar_btn = page.query_selector("button:has-text('Generar informe')")
                            if generar_btn:
                                generar_btn.click()
                                try:
                                    page.wait_for_selector("a:has-text('Excel')", timeout=2000)
                                except:
                                    pass
                            
                            excel_link = page.query_selector("a:has-text('Excel')")
                        
                        if not excel_link:
                            excel_link = page.query_selector(".dropdown-menu >> text=Excel")
                        if not excel_link:
                            excel_link = page.query_selector("text=Excel")
                        
                        if excel_link:
                            with page.expect_download(timeout=download_timeout) as download_info:
                                excel_link.click()
                            
                            download = download_info.value
                            download_path = downloads_folder / f"{current_date}.xlsx"
                            download.save_as(download_path)
                            downloaded_files.append(download_path)
                            
                            self.log(f"   ✅ Guardado: {current_date}.xlsx")
                        else:
                            self.log(f"   ⚠️ No se encontró el enlace Excel para {current_date}")
                        
                        if wait_time > 0:
                            page.wait_for_timeout(int(wait_time * 1000))
                        
                    except PlaywrightTimeout:
                        self.log(f"   ⚠️ Timeout en {current_date}, continuando...")
                    except Exception as e:
                        self.log(f"   ❌ Error en {current_date}: {str(e)}")
                
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
        
        summary_table['Hospitalización Total'] = (
            safe_get_col(summary_table, 'Hospitalización') + 
            safe_get_col(summary_table, 'URGENTE HOSPITALIZACION') + 
            safe_get_col(summary_table, 'Sin tipo atención')
        )
        
        summary_table['Consulta Externa Total'] = (
            safe_get_col(summary_table, 'Consulta Externa') + 
            safe_get_col(summary_table, 'URGENTE CONSULTA EXTERNA') + 
            safe_get_col(summary_table, 'REFERENCIA') + 
            safe_get_col(summary_table, 'URGENTE REFERENCIA')
        )
        
        if 'Emergencia' not in summary_table.columns:
            summary_table['Emergencia'] = 0
        
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
