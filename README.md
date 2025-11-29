# 🏥 Estadística Hospital - Automatizado

Aplicación de escritorio para automatizar la descarga y consolidación de informes estadísticos de exámenes del hospital.

**Versión 3.0** - Aplicación completamente reescrita en Python, portable como un solo archivo `.exe`.

## ✨ Características

- 🖥️ **Interfaz gráfica** fácil de usar
- 🌐 **Automatización de navegador** con Playwright (más confiable que AutoHotkey)
- 📊 **Procesamiento de datos** con pandas
- 📦 **Ejecutable portable** - un solo archivo `.exe`, sin instalaciones
- ⚙️ **Configurable** - archivos de configuración externos para fácil mantenimiento
- 🔄 **Barra de progreso** y registro de actividad en tiempo real

## 🚀 Instalación

### Opción 1: Descargar el Ejecutable (Recomendado)

1. Descargue `EstadisticaHospital.exe` desde [Releases](../../releases)
2. Descargue también `config.ini` y `config_examenes.json`
3. Coloque los 3 archivos en una carpeta
4. Cree una carpeta llamada `ExcelsDescargados` en el mismo lugar
5. ¡Listo! Ejecute `EstadisticaHospital.exe`

Estructura final:
```
MiCarpeta/
├── EstadisticaHospital.exe
├── config.ini
├── config_examenes.json
└── ExcelsDescargados/
```

### Opción 2: Ejecutar desde Código Fuente

```bash
# Clonar repositorio
git clone https://github.com/usuario/Estadistica-Hospital.git
cd Estadistica-Hospital

# Instalar dependencias
pip install -r requirements.txt

# Ejecutar
python EstadisticaHospital.py
```

### Opción 3: Compilar su Propio Ejecutable

```bash
# Instalar dependencias
pip install -r requirements.txt
pip install pyinstaller

# Compilar (o ejecute build.bat)
pyinstaller --onefile --windowed --name "EstadisticaHospital" EstadisticaHospital.py
```

## 📦 Requisitos

- **Sistema Operativo:** Windows 10 o superior
- **Navegador:** Google Chrome instalado
- **Sesión:** Debe haber iniciado sesión previamente en el sistema del hospital

> **Nota:** El ejecutable `.exe` es portable y no requiere Python instalado. Solo necesita Chrome.

## 🎯 Uso

1. **Ejecute** `EstadisticaHospital.exe`

2. **Complete los parámetros:**
   | Campo | Descripción | Valor por defecto |
   |-------|-------------|-------------------|
   | Año | Año de los reportes | Año actual |
   | Mes | Mes de los reportes (1-12) | Mes actual |
   | Día inicial | Primer día del rango | 1 |
   | Día final | Último día del rango | Día anterior |
   | Tiempo entre descargas | Segundos de espera | 2 |
   | Modo invisible | Ocultar ventana del navegador | No |

3. **Haga clic en "Iniciar Descarga"**

4. **Resultado:** 
   - Archivos individuales en `ExcelsDescargados/`
   - Reporte consolidado: `Estadistica Hospital.xlsx`

## 📊 Archivo de Salida

El archivo `Estadistica Hospital.xlsx` contiene 3 hojas:

| Hoja | Contenido |
|------|-----------|
| **Estadistica Calculada** | Resumen por categoría y fecha con totales |
| **Examenes Categorizados** | Todos los exámenes con su categoría y multiplicador |
| **Datos Descargados** | Datos crudos combinados de todos los archivos |

## ⚙️ Configuración

### config.ini - Parámetros Generales

```ini
[General]
URL=https://hjmvi.orion-labs.com/informes/estadisticos
TiempoEspera=2          ; Segundos entre descargas
TiempoCargaPagina=5     ; Segundos para cargar página
Headless=false          ; true = sin ventana del navegador

[Informe]
NombreDropdown=Agrupar por
OpcionAgrupacion=Sección por tipo atención
```

### config_examenes.json - Multiplicadores y Categorías

```json
{
    "multipliers": {
        "BIOMETRÍA HEMÁTICA": 18,
        "COPROPARASITARIO": 2
    },
    "cultivo_multiplier": 10,
    "exam_categories": {
        "LEISHMANIA": "Hematologico"
    },
    "seccion_categories": {
        "Hematología": "Hematologico"
    }
}
```

#### Agregar un nuevo examen con multiplicador

1. Abra `config_examenes.json` con un editor de texto
2. En `"multipliers"`, agregue:
   ```json
   "NOMBRE EXACTO DEL EXAMEN": 5,
   ```

#### Agregar una nueva categoría

1. Abra `config_examenes.json`
2. En `"seccion_categories"`, agregue:
   ```json
   "Nombre de la Sección": "NombreCategoria",
   ```

## 🔧 Solución de Problemas

### "Chrome no está instalado"
- **Solución:** Instale Google Chrome desde [google.com/chrome](https://www.google.com/chrome/)

### "No se encontró el botón 'Generar informe'"
- **Causa:** No hay sesión activa
- **Solución:** 
  1. Abra Chrome manualmente
  2. Vaya a la URL del sistema e inicie sesión
  3. Cierre Chrome y ejecute el programa nuevamente

### "Timeout" o descargas lentas
- **Solución:** Aumente `TiempoEspera` en `config.ini` (pruebe con 3 o 4)

### La ventana se cierra inmediatamente
- **Causa:** Error al iniciar
- **Solución:** Ejecute desde terminal para ver el error:
  ```bash
  EstadisticaHospital.exe
  ```

### El programa no encuentra los archivos de configuración
- **Causa:** Los archivos .ini y .json no están junto al .exe
- **Solución:** Asegúrese de que `config.ini` y `config_examenes.json` estén en la misma carpeta que el .exe

## 📁 Estructura del Proyecto

```
Estadistica-Hospital/
├── EstadisticaHospital.py      # Código fuente principal
├── EstadisticaHospital.exe     # Ejecutable compilado
├── config.ini                  # Configuración general
├── config_examenes.json        # Multiplicadores y categorías
├── requirements.txt            # Dependencias de Python
├── build.bat                   # Script para compilar .exe
├── README.md                   # Esta documentación
├── .gitignore                  # Archivos ignorados por git
└── ExcelsDescargados/          # Carpeta de descargas
    ├── 2024-01-01.xlsx
    └── ...
```

## 🔄 Migración desde v2.0 (AutoHotkey)

Si viene de la versión anterior con AutoHotkey:

### Archivos a ELIMINAR:
```
❌ EstadisticaAutomatizado.ahk
❌ EstadisticaAutomatizado.exe (el viejo de AHK)
❌ UIA.ahk
❌ UIA_Browser.ahk
❌ UnirExcels.bat
❌ UnirTablas.py
❌ installer.iss
❌ RELEASE_INSTRUCTIONS.md
```

### Archivos a CONSERVAR:
```
✅ config.ini (actualizar si es necesario)
✅ config_examenes.json
✅ .gitignore
✅ ExcelsDescargados/ (la carpeta)
```

### Archivos NUEVOS:
```
📄 EstadisticaHospital.py
📄 EstadisticaHospital.exe
📄 build.bat
📄 requirements.txt
📄 README.md (actualizado)
```

## 📝 Changelog

### v3.0 (Actual)
- 🔄 **Reescritura completa en Python**
- ✨ Nueva interfaz gráfica con tkinter
- 🌐 Automatización con Playwright (reemplaza AutoHotkey + UIA)
- 📦 Ejecutable portable con PyInstaller
- 📊 Procesamiento integrado (no más scripts separados)
- 🎯 Barra de progreso y logs en tiempo real
- 🛑 Botón para detener el proceso

### v2.0
- Añadido config.ini y config_examenes.json
- Reemplazado R por Python para procesamiento

### v1.0
- Versión inicial con AutoHotkey + R

## 🛠️ Desarrollo

### Modificar el código
1. Clone el repositorio
2. Instale dependencias: `pip install -r requirements.txt`
3. Edite `EstadisticaHospital.py`
4. Pruebe: `python EstadisticaHospital.py`
5. Compile: `build.bat` o `pyinstaller ...`

### Crear un Release
1. Actualice la versión en el código si es necesario
2. Ejecute `build.bat` para compilar
3. Cree un ZIP con:
   - `dist/EstadisticaHospital.exe`
   - `config.ini`
   - `config_examenes.json`
   - `README.md`
4. Suba a GitHub Releases

## 📄 Licencia

Este proyecto es de uso interno del hospital.

## 📞 Soporte

Si encuentra problemas:
1. Revise la sección de Solución de Problemas
2. Abra un Issue en este repositorio con:
   - Descripción del problema
   - Captura del log de la aplicación
   - Sistema operativo y versión de Chrome
