# ☠️ XLSBruteforce

**XLSBruteforce** es una herramienta escrita en Python que permite realizar ataques de fuerza bruta sobre hojas de cálculo de Excel (.xlsx) protegidas con contraseña a nivel de hoja (candado). Utiliza procesamiento paralelo para aprovechar todos los núcleos disponibles y probar diccionarios de contraseñas de forma eficiente.

---

## ☠️ Características

- ✅ Selección interactiva del archivo `.xlsx` a atacar
- ✅ Procesamiento multiproceso (`multiprocessing`) para máximo rendimiento
- ✅ Compatibilidad con diccionarios masivos (`SecLists`, `rockyou.txt`, etc.)
- ✅ Guarda automáticamente el archivo desbloqueado si se encuentra la contraseña

---

## ☠️ Captura


---

## ☠️ Requisitos

Instala los módulos necesarios:

```bash
pip install -r requirements.txt
```

---
## uso 

```bash
python3 main.py
```

El script te pedirá que selecciones un archivo .xlsx y luego comenzará a probar diccionarios de contraseñas desde tu carpeta predefinida:

```bash
SECLISTS_PATH = "/Users/usuario/Documents/Tools/SecLists/Passwords"
```

---


