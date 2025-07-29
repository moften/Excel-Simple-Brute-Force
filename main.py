
from multiprocessing import Pool, cpu_count
import openpyxl
import os
import time
import sys
from tqdm import tqdm  
from colorama import init, Fore, Style

SECLISTS_PATH = "/Users/user/Documents/Tools/SecLists/Passwords" #pon tu carpeta de diccionarios aqui
EXCEL_FILE = None

def print_banner():
    init(autoreset=True)
    print(Fore.GREEN + Style.BRIGHT + r""" 

▒██   ██▒ ██▓      ██████  ▄▄▄▄     █████▒
▒▒ █ █ ▒░▓██▒    ▒██    ▒ ▓█████▄ ▓██   ▒ 
░░  █   ░▒██░    ░ ▓██▄   ▒██▒ ▄██▒████ ░ 
 ░ █ █ ▒ ▒██░      ▒   ██▒▒██░█▀  ░▓█▒  ░ 
▒██▒ ▒██▒░██████▒▒██████▒▒░▓█  ▀█▓░▒█░    
▒▒ ░ ░▓ ░░ ▒░▓  ░▒ ▒▓▒ ▒ ░░▒▓███▀▒ ▒ ░    
░░   ░▒ ░░ ░ ▒  ░░ ░▒  ░ ░▒░▒   ░  ░      
 ░    ░    ░ ░   ░  ░  ░   ░    ░  ░ ░    
 ░    ░      ░  ░      ░   ░              
                                ░         

    """)
    print(Fore.YELLOW + Style.BRIGHT + "   XLSBruteforce — fuerza bruta sobre hojas protegidas de Excel")
    print(Fore.CYAN + Style.BRIGHT + "     By m10sec  —  Tremendo flipador de Tools\n")


def cargar_diccionario(diccionario_path):
    try:
        with open(diccionario_path, "r", encoding="latin-1", errors="ignore") as f:
            return [line.strip() for line in f if line.strip()]
    except Exception as e:
        print(f"[!] Error leyendo diccionario: {e}")
        return []


def probar_password(args):
    password, filepath = args
    try:
        wb = openpyxl.load_workbook(filepath)
        for sheet in wb.worksheets:
            if sheet.protection.sheet:
                sheet.protection.set_password(password)
                if not sheet.protection.sheet:
                    output_file = "desprotegido_" + os.path.basename(filepath)
                    wb.save(output_file)
                    return password
    except Exception:
        return None
    return None


def seleccionar_archivo():
    archivos = [f for f in os.listdir() if f.endswith(".xlsx")]
    if not archivos:
        print("☠️ No se encontraron archivos .xlsx en el directorio actual.")
        sys.exit(1)
    print("☠️ Archivos .xlsx disponibles:\n")
    for i, f in enumerate(archivos, 1):
        print(f"{i}. {f}")
    try:
        eleccion = int(input("\nSelecciona el archivo (número): "))
        return archivos[eleccion - 1]
    except:
        print("☠️ Selección inválida.")
        sys.exit(1)


def buscar_diccionarios(base_path):
    diccionarios = []
    for root, _, files in os.walk(base_path):
        for file in files:
            if file.endswith(".txt"):
                diccionarios.append(os.path.join(root, file))
    return diccionarios


def ataque_diccionario_con_paralelismo(diccionario_path, filepath):
    contraseñas = cargar_diccionario(diccionario_path)
    if not contraseñas:
        return False

    print(f"\n[*] Probando {len(contraseñas)} contraseñas desde: {os.path.basename(diccionario_path)}")

    args = [(pw, filepath) for pw in contraseñas]
    with Pool(cpu_count()) as pool:
        with tqdm(total=len(args), desc="⏳ Progreso", ncols=80) as barra:
            for resultado in pool.imap_unordered(probar_password, args):
                barra.update(1)
                if resultado:
                    barra.close()
                    print(f"\n[+] Contraseña encontrada: {resultado}")
                    return True
    return False


def ejecutar():
    global EXCEL_FILE
    print_banner()

    print("=== Excel Brute Force by m10sec ===")
    EXCEL_FILE = seleccionar_archivo()
    diccionarios = buscar_diccionarios(SECLISTS_PATH)

    print(f"\n☠️ Se encontraron {len(diccionarios)} diccionarios en: {SECLISTS_PATH}")

    start_time = time.time()

    for dicc in diccionarios:
        if ataque_diccionario_con_paralelismo(dicc, EXCEL_FILE):
            print("\n✅ Finalizado con éxito.")
            break
    else:
        print("\n☠️ No se encontró la contraseña en los diccionarios.")

    end_time = time.time()
    print(f"\n☠️☠️☠️☠️ Tiempo total de ejecución: {end_time - start_time:.2f} segundos ☠️☠️☠️☠️")


if __name__ == "__main__":
    ejecutar()
