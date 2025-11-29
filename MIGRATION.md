# 🔄 Guía de Migración v2.0 → v3.0

## Archivos a ELIMINAR de tu proyecto actual:

```
❌ EstadisticaAutomatizado.ahk      (reemplazado por EstadisticaHospital.py)
❌ EstadisticaAutomatizado.exe      (el viejo .exe de AutoHotkey)
❌ UIA.ahk                          (ya no se usa AutoHotkey)
❌ UIA_Browser.ahk                  (ya no se usa AutoHotkey)
❌ UnirExcels.bat                   (integrado en el nuevo .py)
❌ UnirTablas.py                    (integrado en el nuevo .py)
❌ installer.iss                    (ya no se usa Inno Setup)
❌ RELEASE_INSTRUCTIONS.md          (obsoleto)
```

## Archivos a CONSERVAR:

```
✅ config.ini                       (compatible, quizás actualizar)
✅ config_examenes.json             (compatible)
✅ .gitignore                       (actualizar con el nuevo)
✅ ExcelsDescargados/               (mantener la carpeta)
```

## Archivos NUEVOS a agregar:

```
📄 EstadisticaHospital.py           (código fuente principal)
📄 build.bat                        (para compilar el .exe)
📄 requirements.txt                 (dependencias de Python)
📄 README.md                        (documentación actualizada)
```

## Pasos para migrar:

1. **Hacer backup** de tu proyecto actual (por si acaso)

2. **Eliminar** los archivos marcados con ❌ arriba

3. **Copiar** los nuevos archivos (📄) a tu carpeta del proyecto

4. **Actualizar .gitignore** con el contenido del nuevo archivo

5. **Verificar config.ini** - el nuevo tiene una opción adicional:
   ```ini
   [General]
   Headless=false    ; Nueva opción
   
   [Archivos]
   ArchivoSalida=./Estadistica Hospital.xlsx  ; Nueva opción
   ```

6. **Instalar Python** si no lo tienes:
   - Descargar desde https://www.python.org/downloads/
   - ⚠️ Marcar "Add Python to PATH" durante instalación

7. **Instalar dependencias:**
   ```bash
   pip install -r requirements.txt
   ```

8. **Probar:**
   ```bash
   python EstadisticaHospital.py
   ```

9. **Compilar el .exe** (opcional):
   ```bash
   build.bat
   ```
   El .exe estará en `dist/EstadisticaHospital.exe`

## Comandos útiles:

```bash
# Eliminar archivos viejos (ejecutar en PowerShell desde tu carpeta)
Remove-Item EstadisticaAutomatizado.ahk
Remove-Item EstadisticaAutomatizado.exe
Remove-Item UIA.ahk
Remove-Item UIA_Browser.ahk
Remove-Item UnirExcels.bat
Remove-Item UnirTablas.py
Remove-Item installer.iss
Remove-Item RELEASE_INSTRUCTIONS.md
```

## Estructura final del proyecto:

```
Estadistica-Hospital/
├── EstadisticaHospital.py      # Nuevo código principal
├── config.ini                  # Conservado
├── config_examenes.json        # Conservado
├── requirements.txt            # Nuevo
├── build.bat                   # Nuevo
├── README.md                   # Actualizado
├── .gitignore                  # Actualizado
├── MIGRATION.md                # Este archivo (puede eliminar después)
└── ExcelsDescargados/          # Conservado
```
