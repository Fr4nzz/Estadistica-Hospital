"""
Estadística Hospital - Automatización de descarga y procesamiento de informes
==============================================================================
Aplicación completa para automatizar la descarga de reportes estadísticos
del hospital, procesarlos y generar un informe consolidado.

Autor: Hospital
Versión: 3.0
"""

import os
import re
import sys
import json
import time
import threading
import configparser
from pathlib import Path
from datetime import datetime, timedelta
from typing import Optional, Dict, Any
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext

import pandas as pd
from playwright.sync_api import sync_playwright, Browser, Page, TimeoutError as PlaywrightTimeout

# ============================================
# CONFIGURACIÓN POR DEFECTO
# ============================================

DEFAULT_CONFIG = {
    "General": {
        "URL": "https://hjmvi.orion-labs.com/informes/estadisticos",
        "TiempoEspera": "2",
        "TiempoCargaPagina": "5",
        "Headless": "false"  # true = sin ventana del navegador
    },
    "Fechas": {
        "AñoPorDefecto": "",
        "MesPorDefecto": "",
        "DiaInicialPorDefecto": "1",
        "DiaFinalPorDefecto": ""
    },
    "Informe": {
        "NombreDropdown": "Agrupar por",
        "OpcionAgrupacion": "Sección por tipo atención",
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
        "TIPIFICACION SANGUINEA RH (D)": 3,
    },
    "cultivo_multiplier": 10,
    "exam_categories": {
        "LEISHMANIA": "Hematologico",
        "CRISTALOGRAFÍA": "Bacteriológico",
        "GRAM (GOTA FRESCA) ORINA": "Bacteriológico",
        "GASOMETRIA ARTERIAL": "Quimica sanguinea",
        "GASOMETRIA VENOSA": "Quimica sanguinea",
    },
    "seccion_categories": {
        "Autoinmunes e Infecciosas": "Serologicos",
        "Drogas y Fármacos": "Serologicos",
        "Serología": "Serologicos",
        "Bioquímica": "Quimica sanguinea",
        "Electrolitos": "Quimica sanguinea",
        "Inmunoquímica Sanguínea": "Quimica sanguinea",
        "Química Clínica en Orina": "Quimica sanguinea",
        "Uroanálisis": "Orina",
        "Coproanálisis": "Materias fecales",
        "Biología Molecular": "Hormonales",
        "Estudios Hormonales": "Hormonales",
        "Marcadores Tumorales": "Hormonales",
        "Coagulación": "Hematologico",
        "Hematología": "Hematologico",
        "Inmunohematología": "Hematologico",
        "Microbiología": "Bacteriológico",
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
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Estadística Hospital - Automatizado")
        self.root.geometry("600x700")
        self.root.resizable(True, True)
        
        # Determinar directorio base (funciona tanto en desarrollo como en exe)
        if getattr(sys, 'frozen', False):
            self.base_dir = Path(sys.executable).parent
        else:
            self.base_dir = Path(__file__).parent
        
        # Cargar configuración
        self.config = self.load_config()
        self.exam_config = self.load_exam_config()
        
        # Variables de control
        self.is_running = False
        self.should_stop = False
        
        # Crear interfaz
        self.create_widgets()
        
        # Crear carpeta de descargas si no existe
        downloads_folder = self.base_dir / self.config.get("Archivos", "CarpetaDescargas")
        downloads_folder.mkdir(exist_ok=True)
    
    def load_config(self) -> configparser.ConfigParser:
        """Carga la configuración desde config.ini"""
        config = configparser.ConfigParser()
        config_path = self.base_dir / "config.ini"
        
        # Establecer valores por defecto
        for section, values in DEFAULT_CONFIG.items():
            config[section] = values
        
        # Cargar desde archivo si existe
        if config_path.exists():
            config.read(config_path, encoding='utf-8')
        
        return config
    
    def load_exam_config(self) -> Dict[str, Any]:
        """Carga la configuración de exámenes desde config_examenes.json"""
        config_path = self.base_dir / "config_examenes.json"
        
        if config_path.exists():
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except json.JSONDecodeError:
                self.log("⚠️ Error al leer config_examenes.json, usando valores por defecto")
        
        return DEFAULT_EXAM_CONFIG
    
    def create_widgets(self):
        """Crea la interfaz gráfica"""
        # Frame principal con padding
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky="nsew")
        
        # Configurar el grid
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Título
        title_label = ttk.Label(main_frame, text="🏥 Estadística Hospital", font=('Helvetica', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 15))
        
        # === Sección de Fechas ===
        date_frame = ttk.LabelFrame(main_frame, text="Rango de Fechas", padding="10")
        date_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        date_frame.columnconfigure(1, weight=1)
        date_frame.columnconfigure(3, weight=1)
        
        # Año
        ttk.Label(date_frame, text="Año:").grid(row=0, column=0, sticky="w", padx=(0, 5))
        self.year_var = tk.StringVar(value=self.get_default_year())
        self.year_entry = ttk.Entry(date_frame, textvariable=self.year_var, width=10)
        self.year_entry.grid(row=0, column=1, sticky="w", padx=(0, 20))
        
        # Mes
        ttk.Label(date_frame, text="Mes:").grid(row=0, column=2, sticky="w", padx=(0, 5))
        self.month_var = tk.StringVar(value=self.get_default_month())
        self.month_entry = ttk.Entry(date_frame, textvariable=self.month_var, width=10)
        self.month_entry.grid(row=0, column=3, sticky="w")
        
        # Día inicial
        ttk.Label(date_frame, text="Día inicial:").grid(row=1, column=0, sticky="w", padx=(0, 5), pady=(10, 0))
        self.start_day_var = tk.StringVar(value=self.config.get("Fechas", "DiaInicialPorDefecto") or "1")
        self.start_day_entry = ttk.Entry(date_frame, textvariable=self.start_day_var, width=10)
        self.start_day_entry.grid(row=1, column=1, sticky="w", padx=(0, 20), pady=(10, 0))
        
        # Día final
        ttk.Label(date_frame, text="Día final:").grid(row=1, column=2, sticky="w", padx=(0, 5), pady=(10, 0))
        self.end_day_var = tk.StringVar(value=self.get_default_end_day())
        self.end_day_entry = ttk.Entry(date_frame, textvariable=self.end_day_var, width=10)
        self.end_day_entry.grid(row=1, column=3, sticky="w", pady=(10, 0))
        
        # === Sección de Opciones ===
        options_frame = ttk.LabelFrame(main_frame, text="Opciones", padding="10")
        options_frame.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        options_frame.columnconfigure(1, weight=1)
        
        # Tiempo de espera
        ttk.Label(options_frame, text="Tiempo entre descargas (segundos):").grid(row=0, column=0, sticky="w", padx=(0, 5))
        self.wait_time_var = tk.StringVar(value=self.config.get("General", "TiempoEspera"))
        self.wait_time_entry = ttk.Entry(options_frame, textvariable=self.wait_time_var, width=10)
        self.wait_time_entry.grid(row=0, column=1, sticky="w")
        
        # Modo headless
        self.headless_var = tk.BooleanVar(value=self.config.get("General", "Headless").lower() == "true")
        self.headless_check = ttk.Checkbutton(
            options_frame, 
            text="Modo invisible (sin ventana del navegador)", 
            variable=self.headless_var
        )
        self.headless_check.grid(row=1, column=0, columnspan=2, sticky="w", pady=(10, 0))
        
        # === Botones ===
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, columnspan=2, pady=10)
        
        self.start_button = ttk.Button(
            button_frame, 
            text="▶ Iniciar Descarga", 
            command=self.start_process,
            width=20
        )
        self.start_button.grid(row=0, column=0, padx=5)
        
        self.stop_button = ttk.Button(
            button_frame, 
            text="⏹ Detener", 
            command=self.stop_process,
            width=20,
            state="disabled"
        )
        self.stop_button.grid(row=0, column=1, padx=5)
        
        # === Barra de progreso ===
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(
            main_frame, 
            variable=self.progress_var, 
            maximum=100,
            mode='determinate'
        )
        self.progress_bar.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(0, 5))
        
        self.progress_label = ttk.Label(main_frame, text="Listo para iniciar")
        self.progress_label.grid(row=5, column=0, columnspan=2)
        
        # === Log ===
        log_frame = ttk.LabelFrame(main_frame, text="Registro de Actividad", padding="5")
        log_frame.grid(row=6, column=0, columnspan=2, sticky="nsew", pady=(10, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(6, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=15, state='disabled', wrap='word')
        self.log_text.grid(row=0, column=0, sticky="nsew")
    
    def get_default_year(self) -> str:
        """Obtiene el año por defecto"""
        default = self.config.get("Fechas", "AñoPorDefecto")
        return default if default else str(datetime.now().year)
    
    def get_default_month(self) -> str:
        """Obtiene el mes por defecto"""
        default = self.config.get("Fechas", "MesPorDefecto")
        return default if default else str(datetime.now().month)
    
    def get_default_end_day(self) -> str:
        """Obtiene el día final por defecto (ayer)"""
        default = self.config.get("Fechas", "DiaFinalPorDefecto")
        if default:
            return default
        yesterday = datetime.now() - timedelta(days=1)
        return str(yesterday.day)
    
    def log(self, message: str):
        """Añade un mensaje al log"""
        self.log_text.configure(state='normal')
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert('end', f"[{timestamp}] {message}\n")
        self.log_text.see('end')
        self.log_text.configure(state='disabled')
        self.root.update_idletasks()
    
    def update_progress(self, current: int, total: int, message: str = ""):
        """Actualiza la barra de progreso"""
        percentage = (current / total) * 100 if total > 0 else 0
        self.progress_var.set(percentage)
        if message:
            self.progress_label.config(text=message)
        self.root.update_idletasks()
    
    def validate_inputs(self) -> bool:
        """Valida los inputs del formulario"""
        try:
            year = int(self.year_var.get())
            month = int(self.month_var.get())
            start_day = int(self.start_day_var.get())
            end_day = int(self.end_day_var.get())
            wait_time = float(self.wait_time_var.get())
            
            if not (1900 <= year <= 2100):
                raise ValueError("El año debe estar entre 1900 y 2100")
            if not (1 <= month <= 12):
                raise ValueError("El mes debe estar entre 1 y 12")
            if not (1 <= start_day <= 31):
                raise ValueError("El día inicial debe estar entre 1 y 31")
            if not (1 <= end_day <= 31):
                raise ValueError("El día final debe estar entre 1 y 31")
            if start_day > end_day:
                raise ValueError("El día inicial no puede ser mayor que el día final")
            if wait_time < 0:
                raise ValueError("El tiempo de espera no puede ser negativo")
            
            return True
        
        except ValueError as e:
            messagebox.showerror("Error de validación", str(e))
            return False
    
    def start_process(self):
        """Inicia el proceso de descarga"""
        if not self.validate_inputs():
            return
        
        self.is_running = True
        self.should_stop = False
        self.start_button.config(state="disabled")
        self.stop_button.config(state="normal")
        self.log_text.configure(state='normal')
        self.log_text.delete('1.0', 'end')
        self.log_text.configure(state='disabled')
        
        # Ejecutar en un hilo separado para no bloquear la GUI
        thread = threading.Thread(target=self.run_automation)
        thread.daemon = True
        thread.start()
    
    def stop_process(self):
        """Detiene el proceso de descarga"""
        self.should_stop = True
        self.log("⏹ Deteniendo proceso...")
    
    def run_automation(self):
        """Ejecuta la automatización del navegador"""
        try:
            year = int(self.year_var.get())
            month = int(self.month_var.get())
            start_day = int(self.start_day_var.get())
            end_day = int(self.end_day_var.get())
            wait_time = float(self.wait_time_var.get())
            headless = self.headless_var.get()
            
            url = self.config.get("General", "URL")
            page_load_time = int(self.config.get("General", "TiempoCargaPagina"))
            download_timeout = int(self.config.get("General", "TimeoutDescarga", fallback="15")) * 1000
            
            # Configuración del informe
            dropdown_id = self.config.get("Informe", "IdDropdownAgrupar", fallback="agrupar-por")
            dropdown_value = self.config.get("Informe", "ValorAgrupacion", fallback="SECCION_TIPO_ATENCION")
            id_fecha_desde = self.config.get("Informe", "IdFechaDesde")
            id_fecha_hasta = self.config.get("Informe", "IdFechaHasta")
            
            downloads_folder = self.base_dir / self.config.get("Archivos", "CarpetaDescargas")
            downloads_folder.mkdir(exist_ok=True)
            
            # Carpeta para datos persistentes del navegador (cookies, contraseñas, etc.)
            browser_data_folder = self.base_dir / "browser_data"
            browser_data_folder.mkdir(exist_ok=True)
            
            total_days = end_day - start_day + 1
            
            self.log("=" * 50)
            self.log("🚀 Iniciando automatización...")
            self.log(f"📅 Rango: {year}-{month:02d}-{start_day:02d} al {year}-{month:02d}-{end_day:02d}")
            self.log(f"📁 Carpeta de descargas: {downloads_folder}")
            self.log("=" * 50)
            
            with sync_playwright() as p:
                # Configurar navegador con contexto persistente (guarda cookies y contraseñas)
                self.log("🌐 Iniciando navegador Chrome...")
                self.log("   💾 Los datos de sesión se guardarán para futuros usos")
                
                context = p.chromium.launch_persistent_context(
                    user_data_dir=str(browser_data_folder),
                    headless=headless,
                    channel="chrome",  # Usa Chrome instalado
                    accept_downloads=True
                )
                
                page = context.pages[0] if context.pages else context.new_page()
                
                # Navegar al sitio
                self.log(f"📄 Navegando a {url}")
                page.goto(url, timeout=60000)
                page.wait_for_load_state("networkidle", timeout=page_load_time * 1000)
                
                # Verificar si necesita login - esperar a que el usuario inicie sesión
                max_login_wait = 300  # 5 minutos máximo para iniciar sesión
                login_check_interval = 2  # Revisar cada 2 segundos
                waited_time = 0
                logged_in = False
                first_message = True
                
                while not logged_in:
                    if self.should_stop:
                        self.log("⏹ Proceso cancelado por el usuario")
                        browser.close()
                        self.finish_process(success=False)
                        return
                    
                    try:
                        # Esperar a que la página esté estable después de cualquier navegación
                        page.wait_for_load_state("domcontentloaded", timeout=3000)
                        
                        # Verificar si encontramos el botón de generar informe
                        generar_btn = page.query_selector("button:has-text('Generar informe')")
                        if generar_btn:
                            logged_in = True
                            break
                        
                        # Si estamos en la página correcta pero no hay botón, mostrar mensaje
                        if first_message:
                            self.log("🔐 No se detectó sesión activa.")
                            self.log("   👉 Por favor, inicie sesión en la ventana del navegador...")
                            self.log("   ⏳ Esperando inicio de sesión (máximo 5 minutos)...")
                            self.update_progress(0, 100, "Esperando inicio de sesión...")
                            first_message = False
                        
                        # Esperar antes del siguiente intento
                        page.wait_for_timeout(login_check_interval * 1000)
                        
                    except Exception as e:
                        # Si hay error (como navegación), esperar un momento y reintentar
                        if first_message:
                            self.log("🔐 Página de login detectada.")
                            self.log("   👉 Por favor, inicie sesión en la ventana del navegador...")
                            self.log("   ⏳ Esperando inicio de sesión (máximo 5 minutos)...")
                            self.update_progress(0, 100, "Esperando inicio de sesión...")
                            first_message = False
                        
                        # Usar time.sleep como fallback cuando page.wait_for_timeout falla
                        time.sleep(login_check_interval)
                    
                    waited_time += login_check_interval
                    
                    # Actualizar progreso mientras espera
                    if waited_time % 10 == 0 and waited_time > 0:
                        self.log(f"   ⏳ Esperando... ({waited_time}s)")
                    
                    if waited_time >= max_login_wait:
                        self.log("❌ Tiempo de espera agotado para inicio de sesión.")
                        browser.close()
                        self.finish_process(success=False)
                        return
                
                self.log("✅ Sesión detectada correctamente!")
                
                # Esperar un momento para que la página se estabilice completamente
                try:
                    page.wait_for_load_state("networkidle", timeout=10000)
                except:
                    pass  # Continuar aunque falle
                time.sleep(1)  # Extra safety wait
                
                self.log("🚀 Comenzando descargas...")
                
                # Procesar cada día
                downloaded_files = []
                
                for day_index, current_day in enumerate(range(start_day, end_day + 1)):
                    if self.should_stop:
                        self.log("⏹ Proceso detenido por el usuario")
                        break
                    
                    current_date = f"{year}-{month:02d}-{current_day:02d}"
                    self.update_progress(day_index, total_days, f"Descargando {current_date}...")
                    self.log(f"📥 Procesando día {current_day}/{end_day}: {current_date}")
                    
                    try:
                        # Establecer fechas usando JavaScript (más rápido que fill)
                        page.evaluate(f"""
                            document.getElementById('{id_fecha_desde}').value = '{current_date}';
                            document.getElementById('{id_fecha_hasta}').value = '{current_date}';
                            // Trigger change events
                            document.getElementById('{id_fecha_desde}').dispatchEvent(new Event('input', {{ bubbles: true }}));
                            document.getElementById('{id_fecha_hasta}').dispatchEvent(new Event('input', {{ bubbles: true }}));
                        """)
                        
                        # Configurar dropdown de agrupación (solo primer día)
                        if day_index == 0:
                            self.log("   🔧 Configurando dropdown de agrupación...")
                            
                            try:
                                # El dropdown tiene id configurable, por defecto "agrupar-por"
                                dropdown = page.query_selector(f"#{dropdown_id}")
                                if dropdown:
                                    dropdown.select_option(value=dropdown_value)
                                    self.log(f"   ✅ Dropdown seleccionado: {dropdown_value}")
                                else:
                                    self.log(f"   ⚠️ No se encontró el dropdown #{dropdown_id}")
                            except Exception as e:
                                self.log(f"   ⚠️ Error seleccionando dropdown: {e}")
                        
                        # Buscar el enlace Excel (puede que el menú ya esté abierto)
                        excel_link = page.query_selector("a:has-text('Excel'):visible")
                        
                        # Si no está visible, hacer clic en "Generar informe" para abrir el menú
                        if not excel_link:
                            generar_btn = page.query_selector("button:has-text('Generar informe')")
                            if generar_btn:
                                generar_btn.click()
                                # Esperar a que aparezca el enlace Excel (máximo 2 segundos)
                                try:
                                    page.wait_for_selector("a:has-text('Excel')", timeout=2000)
                                except:
                                    pass
                            
                            # Buscar de nuevo
                            excel_link = page.query_selector("a:has-text('Excel')")
                        
                        if not excel_link:
                            excel_link = page.query_selector(".dropdown-menu >> text=Excel")
                        if not excel_link:
                            excel_link = page.query_selector("text=Excel")
                        
                        if excel_link:
                            # Descargar Excel
                            with page.expect_download(timeout=download_timeout) as download_info:
                                excel_link.click()
                            
                            download = download_info.value
                            download_path = downloads_folder / f"{current_date}.xlsx"
                            download.save_as(download_path)
                            downloaded_files.append(download_path)
                            
                            self.log(f"   ✅ Guardado: {current_date}.xlsx")
                        else:
                            self.log(f"   ⚠️ No se encontró el enlace Excel para {current_date}")
                        
                        # Pequeña pausa entre descargas (configurable)
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
            
            # Procesar archivos descargados
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
        output_file = self.base_dir / self.config.get("Archivos", "ArchivoSalida")
        
        # Obtener archivos xlsx que comienzan con número (fechas)
        pattern = re.compile(r'^\d.*\.xlsx$')
        xlsx_files = sorted([f for f in downloads_folder.glob('*.xlsx') if pattern.match(f.name)])
        
        if not xlsx_files:
            self.log("❌ No se encontraron archivos Excel para procesar")
            return
        
        self.log(f"   Encontrados {len(xlsx_files)} archivos")
        
        # Leer primer archivo para detectar estructura
        first_file = xlsx_files[0]
        sample_df = pd.read_excel(first_file, skiprows=4)
        self.log(f"   📋 Columnas detectadas: {list(sample_df.columns)}")
        
        # Detectar tipo de informe basado en columnas
        has_patient_type_columns = any(col in sample_df.columns for col in ['Hospitalización', 'Emergencia', 'Consulta Externa'])
        has_simple_columns = 'Cant. Exámenes' in sample_df.columns
        
        if has_patient_type_columns:
            self.log("   ✅ Informe con desglose por tipo de atención detectado")
            numeric_cols = NUMERIC_COLUMNS
        elif has_simple_columns:
            self.log("   ⚠️ Informe simple detectado (sin desglose por tipo de atención)")
            self.log("   ℹ️  El dropdown 'Agrupar por' → 'Sección por tipo atención' no se seleccionó correctamente")
            self.log("   ℹ️  Procesando con la columna 'Cant. Exámenes' como Total")
            numeric_cols = ['Cant. Exámenes', 'Cant. Remisiones', 'Cant. Repeticiones']
        else:
            self.log(f"   ⚠️ Estructura de columnas no reconocida")
            numeric_cols = []
        
        # Leer y combinar archivos
        all_dataframes = []
        for filepath in xlsx_files:
            try:
                df = pd.read_excel(filepath, skiprows=4)
                
                # Renombrar columna
                if 'Sección' in df.columns:
                    df = df.rename(columns={'Sección': 'Seccion'})
                
                # Agregar columna de fecha
                date_str = filepath.stem
                try:
                    df['date'] = pd.to_datetime(date_str, format='%Y-%m-%d')
                except:
                    df['date'] = date_str
                
                # Para informe simple, renombrar columnas para compatibilidad
                if has_simple_columns and not has_patient_type_columns:
                    if 'Cant. Exámenes' in df.columns:
                        df['Total'] = df['Cant. Exámenes']
                    # Crear columnas vacías para las que no existen
                    for col in NUMERIC_COLUMNS:
                        if col not in df.columns:
                            df[col] = 0
                else:
                    # Asegurar columnas numéricas
                    for col in NUMERIC_COLUMNS:
                        if col not in df.columns:
                            df[col] = 0
                
                all_dataframes.append(df)
            except Exception as e:
                self.log(f"   ⚠️ Error leyendo {filepath.name}: {e}")
        
        if not all_dataframes:
            self.log("❌ No se pudieron leer los archivos")
            return
        
        # Combinar
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
        
        # Aplicar multiplicadores a columnas numéricas
        cols_to_multiply = ['Total'] + [col for col in NUMERIC_COLUMNS if col in combined_df.columns and col != 'Total']
        for col in cols_to_multiply:
            if col in combined_df.columns:
                combined_df[col] = pd.to_numeric(combined_df[col], errors='coerce').fillna(0) * combined_df['multiplier']
        
        # Filtrar filas inválidas
        combined_df = combined_df[combined_df['Examen'].notna()]
        combined_df = combined_df[~combined_df['Examen'].astype(str).str.contains('^Total órdenes|^Generado el|^$', na=False, regex=True)]
        # También filtrar filas donde Seccion está vacío (filas de totales)
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
        
        # Determinar qué columnas usar para el resumen
        if has_patient_type_columns:
            summary_cols = [col for col in NUMERIC_COLUMNS if col in combined_df.columns]
        else:
            summary_cols = ['Total'] if 'Total' in combined_df.columns else []
        
        # Remover duplicados manteniendo orden
        summary_cols = list(dict.fromkeys(summary_cols))
        
        self.log(f"   📋 Columnas para resumen: {summary_cols}")
        
        if not summary_cols:
            self.log("   ⚠️ No se encontraron columnas numéricas para el resumen")
            summary_cols = ['Total']
            combined_df['Total'] = 0
        
        # Tabla resumen
        summary_table = combined_df.groupby(['Category', 'date'])[summary_cols].sum().reset_index()
        
        # Columnas calculadas (solo si existen las columnas base)
        # Usar 0 si la columna no existe
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
        
        # Asegurar que Emergencia existe
        if 'Emergencia' not in summary_table.columns:
            summary_table['Emergencia'] = 0
        
        # Totales por fecha - asegurar columnas únicas
        total_cols = list(set(
            [col for col in summary_cols if col in summary_table.columns] + 
            [col for col in ['Hospitalización Total', 'Consulta Externa Total', 'Emergencia', 'Total'] 
             if col in summary_table.columns]
        ))
        
        # Crear totales
        totals = summary_table.groupby('date', as_index=False)[total_cols].sum()
        totals['Category'] = 'TOTAL'
        
        # Concatenar asegurando que las columnas coincidan
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
        
        # Formatear columna de fecha para que sea más legible y renombrar a español
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
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            summary_table.to_excel(writer, sheet_name='Estadistica Calculada', index=False)
            examenes_categorizados.to_excel(writer, sheet_name='Examenes Categorizados', index=False)
            datos_descargados.to_excel(writer, sheet_name='Datos Descargados', index=False)
            
            # Aplicar formato a cada hoja
            for sheet_name in writer.sheets:
                ws = writer.sheets[sheet_name]
                ws.freeze_panes = 'A2'
                ws.auto_filter.ref = ws.dimensions
                
                # Auto-ajustar ancho de columnas
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
                    
                    # Añadir un poco de padding y limitar el ancho máximo
                    adjusted_width = min(max_length + 2, 50)
                    ws.column_dimensions[column_letter].width = adjusted_width
        
        self.log(f"   ✅ Archivo guardado: {output_file}")
    
    def finish_process(self, success: bool):
        """Finaliza el proceso y restaura la GUI"""
        self.is_running = False
        self.start_button.config(state="normal")
        self.stop_button.config(state="disabled")
        
        if success:
            messagebox.showinfo("Completado", "¡El proceso se completó exitosamente!")
        else:
            messagebox.showwarning("Proceso terminado", "El proceso terminó con errores o fue cancelado.")


# ============================================
# PUNTO DE ENTRADA
# ============================================

def main():
    root = tk.Tk()
    
    # Estilo
    style = ttk.Style()
    style.theme_use('clam')  # Tema moderno
    
    app = EstadisticaHospitalApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
