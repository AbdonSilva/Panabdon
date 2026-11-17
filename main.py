#def limpiar pantalla 14
#def inicializar bd17
#def Menu principal 211
#
#
#
import os
import sys
from types import ModuleType
from functools import partial

# Desactivamos los proveedores táctiles problemáticos en Android/Linux para evitar el error de mtdev
os.environ['KIVY_NO_ARGS'] = '1'
os.environ['KIVY_WINDOW'] = 'sdl2'

from kivy.config import Config
# Forzamos a Kivy a ignorar el input mtdev que está dando el fallo
Config.set('input', 'mouse', 'mouse,multitouch_on_demand')
# Si tienes algún input configurado de mtdev en las preferencias, lo removemos o evitamos
try:
    Config.remove_option('input', 'mtdev')
except Exception:
    pass

# Creamos el clon falso para bloquear el mtdev que falta en Android
# Creamos un clon falso avanzado de mtdev con todas las variables que Kivy busca
# Creamos el clon falso blindado con todas las constantes de mtdev que Kivy exige
# --- PEGA ESTO EN SU LUGAR ---
class SmartFakeMtdev(ModuleType):
    def __getattr__(self, name):
        return 0

fake_mtdev = SmartFakeMtdev('kivy.lib.mtdev')
fake_mtdev.Device = object
sys.modules['kivy.lib.mtdev'] = fake_mtdev

import re
import datetime
import sqlite3
# Importamos AMBAS funciones desde encabezado4 y el selector desde diagnostico
from encabezado4 import generar_excel_rp01, generar_consolidado_semanal, guardar_config_db, cargar_config_db
from diagnostico import seleccionar_diagnostico
import actividades
# --- NUEVO SCRIPT DEL RP-66 EN LA CARPETA ABDONL2 ---
from generador_rp66_nuevo import *
from vistas_diagnostico import (
    PantallaMenuDiagnostico, 
    PantallaEPI12,
    PantallaMenuTerreno, 
    PantallaAtencionPrimaria, 
    PantallaTrimestre,
    PantallaSeleccionDireccion,
    ModificarSeleccionDireccion
)
from suscripcion import PantallaCedula
from suscripcion import sincronizacion_manual
from suscripcion import obtener_device_id
from suscripcion import verificar_con_servidor
from suscripcion import guardar_cache_local 
from suscripcion import leer_cache_local
from suscripcion import obtener_medico_vinculado 


# Importaciones de Kivy para tu interfaz gráfica
# Importaciones de Kivy para tu interfaz gráfica
from kivy.app import App
from kivy.core.window import Window
from kivy.uix.label import Label
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
from kivy.uix.scrollview import ScrollView
from kivy.uix.widget import Widget
from kivy.properties import StringProperty
from kivy.metrics import sp
from kivy.uix.popup import Popup
from kivy.clock import Clock
from kivy.uix.textinput import TextInput
CACHE_MANUAL = {}
# ==========================================================

# --- CLASE BLINDADA DE TEXTINPUT PARA ANDROID ---
# ==========================================================
class TextInput(TextInput):
    def __init__(self, **kwargs):
        super(TextInput, self).__init__(**kwargs)
        # Atamos el centrado al cambio de tamaño de la caja para que sea responsivo
        self.bind(size=self._centrar_texto_vertical)

    def _centrar_texto_vertical(self, *args):
        # Solo aplicamos el centrado estricto si es un campo de una sola línea
        if not self.multiline:
            # Calculamos el espacio sobrante entre la altura de la caja y el tamaño de la fuente
            espacio_libre = self.height - self.font_size
            if espacio_libre > 0:
                val = espacio_libre / 2
                # Repartimos el espacio exactamente igual arriba y abajo
                self.padding_y = [val, val]

    def on_touch_up(self, touch):
        # Mantenemos tu toque maestro para asegurar el foco al presionar en el teléfono
        if self.collide_point(*touch.pos):
            self.focus = True
        return super(TextInput, self).on_touch_up(touch)

# ==============================================
# =========================================================

# (y cualquier otro botoncito o elemento de Kivy que vayas a usar)

# Ruta absoluta basada en la ubicación de este script
# --- PARCHE GLOBAL INSTANTÁNEO DE FOCO ---
# ==========================================
# ==========================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "clinica.db")

def limpiar_pantalla():
    os.system('cls' if os.name == 'nt' else 'clear')

def inicializar_bd():
    """Crea la base de datos y la tabla de pacientes si no existen usando ruta absoluta."""
    conexion = sqlite3.connect(DB_PATH)
    cursor = conexion.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS pacientes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT,
            cedula TEXT,
            edad REAL,
            sexo TEXT,
            ambito TEXT,
            embarazada TEXT,
            trimestre TEXT,
            riesgo_prenatal TEXT,
            puerpera TEXT,
            captacion_puerpera TEXT,
            indigena TEXT,
            nacionalidad TEXT,
            direccion TEXT,
            consulta TEXT,
            diagnostico_texto TEXT,
            codigo_epi INTEGER,
            tipo_diagnostico TEXT,
            conducta TEXT,
            fecha_consulta TEXT,
            semana_num TEXT
        )
    ''')
    conexion.commit()
    conexion.close()

def obtener_pacientes_por_fecha(fecha):
    """Consulta los pacientes de un día específico para el RP-01 Diario."""
    conexion = sqlite3.connect(DB_PATH)
    conexion.row_factory = sqlite3.Row
    cursor = conexion.cursor()
    cursor.execute("SELECT * FROM pacientes WHERE fecha_consulta = ?", (fecha,))
    filas = cursor.fetchall()
    conexion.close()
    return [dict(row) for row in filas]

def obtener_pacientes_por_semana(semana_num):
    """Consulta todos los pacientes acumulados en la semana para el Consolidado."""
    conexion = sqlite3.connect(DB_PATH)
    conexion.row_factory = sqlite3.Row
    cursor = conexion.cursor()
    cursor.execute("SELECT * FROM pacientes WHERE semana_num = ?", (str(semana_num),))
    filas = cursor.fetchall()
    conexion.close()
    return [dict(row) for row in filas]


def obtener_pacientes_por_fecha(fecha):
    """Consulta los pacientes de un día específico para el RP-01 Diario."""
    conexion = sqlite3.connect(DB_PATH)
    conexion.row_factory = sqlite3.Row
    cursor = conexion.cursor()
    cursor.execute("SELECT * FROM pacientes WHERE fecha_consulta = ?", (fecha,))
    filas = cursor.fetchall()
    conexion.close()
    return [dict(row) for row in filas]

def obtener_pacientes_por_semana(semana_num):
    """Consulta todos los pacientes acumulados en la semana para el Consolidado."""
    conexion = sqlite3.connect(DB_PATH)
    conexion.row_factory = sqlite3.Row
    cursor = conexion.cursor()
    cursor.execute("SELECT * FROM pacientes WHERE semana_num = ?", (str(semana_num),))
    filas = cursor.fetchall()
    conexion.close()
    return [dict(row) for row in filas]

def revisar_y_corregir_lote(semana_num, filtro_fecha=None):
    """Permite al médico revisar en formato de ficha vertical y corregir antes de exportar."""
    while True:
        limpiar_pantalla
        conexion = sqlite3.connect(DB_PATH)
        conexion.row_factory = sqlite3.Row
        cursor = conexion.cursor()
        
        if filtro_fecha:
            cursor.execute("SELECT * FROM pacientes WHERE semana_num = ? AND fecha_consulta = ?", (str(semana_num), filtro_fecha))
            titulo_contexto = f"SEMANA N° {semana_num} - FECHA: {filtro_fecha}"
        else:
            cursor.execute("SELECT * FROM pacientes WHERE semana_num = ?", (str(semana_num),))
            titulo_contexto = f"CONSOLIDADO SEMANA N° {semana_num}"
            
        pacientes = [dict(row) for row in cursor.fetchall()]
        conexion.close()

        limpiar_pantalla()
        print("=" * 65)
        print(f"   REVISIÓN PREVIA ANTES DE EXPORTAR: {titulo_contexto}")
        print("=" * 65)
        
        if not pacientes:
            print("\n[!] No hay pacientes para revisar en este filtro.")
            input("\nPresione ENTER para continuar...")
            return False

        print(f"{'ID':<4} | {'FECHA':<10} | {'CÉDULA':<12} | {'NOMBRE Y APELLIDO':<25}")
        print("-" * 65)
        for p in pacientes:
            print(f"{p['id']:<4} | {p['fecha_consulta']:<10} | {p['cedula']:<12} | {p['nombre']:<25}")
        print("-" * 65)
        
        accion = input("\n¿Hay algún error que corregir antes de generar el Excel? (S para modificar / N para proceder a exportar) [N]: ").strip().upper()
        
        if accion != "S":
            return True # Todo OK, procedemos a exportar
            
        try:
            id_pac = int(input("Ingrese el ID exacto del paciente que desea revisar y corregir: ").strip())
            
            conexion = sqlite3.connect("clinica.db")
            conexion.row_factory = sqlite3.Row
            cursor = conexion.cursor()
            cursor.execute("SELECT * FROM pacientes WHERE id = ?", (id_pac,))
            p_act = cursor.fetchone()
            conexion.close()
            
            if p_act:
                limpiar_pantalla()
                print("=" * 50)
                print("       FICHA DETALLADA DEL PACIENTE (VERIFICACIÓN)")
                print("=" * 50)
                print(f"ID              : {p_act['id']}")
                print(f"Fecha           : {p_act['fecha_consulta']}")
                print(f"Nombre          : {p_act['nombre']}")
                print(f"Cédula          : {p_act['cedula']}")
                print(f"Edad            : {p_act['edad']}")
                print(f"Sexo            : {p_act['sexo']}")
                print(f"Ámbito          : {p_act['ambito']}")
                print(f"Embarazada      : {p_act['embarazada']} (Trimestre: {p_act['trimestre']}, Riesgo: {p_act['riesgo_prenatal']})")
                print(f"Indígena        : {p_act['indigena']}")
                print(f"Nacionalidad    : {p_act['nacionalidad']}")
                print(f"Dirección       : {p_act['direccion']}")
                print(f"Tipo Consulta   : {p_act['consulta']}")
                print(f"Diagnóstico     : {p_act['diagnostico_texto']} (Cód EPI: {p_act['codigo_epi']})")
                print(f"Tipo Diagnóstico: {p_act['tipo_diagnostico']}")
                print(f"Conducta        : {p_act['conducta']}")
                print("=" * 50)
                
                if input("\n¿Desea editar este registro? (S/N) [S]: ").strip().upper() != "N":
                    nuevo_nombre = input(f"Nombre [{p_act['nombre']}]: ").strip().upper() or p_act['nombre']
                    nueva_cedula = input(f"Cédula [{p_act['cedula']}]: ").strip().upper() or p_act['cedula']
                    
                    nueva_edad_in = input(f"Edad [{p_act['edad']}]: ").strip()
                    nueva_edad = float(nueva_edad_in) if nueva_edad_in else p_act['edad']
                    
                    nuevo_sexo = input(f"Sexo (H/M) [{p_act['sexo']}]: ").strip().upper() or p_act['sexo']
                    nuevo_ambito = input(f"Ámbito (C/T) [{p_act['ambito']}]: ").strip().upper() or p_act['ambito']
                    
                    nueva_embarazada = p_act['embarazada']
                    nuevo_trimestre = p_act['trimestre']
                    nuevo_riesgo = p_act['riesgo_prenatal']
                    
                    if nuevo_sexo == "M":
                        emb_op = input(f"¿Está embarazada? (S/N) [{p_act['embarazada']}]: ").strip().upper()
                        if emb_op in ["S"]:
                            nueva_embarazada = "S"
                            nuevo_trimestre = input(f"  -> Trimestre (1, 2, 3) [{p_act['trimestre']}]: ").strip() or p_act['trimestre']
                            riesgo_op = input(f"  -> ¿Riesgo? (S/N) [{p_act['riesgo_prenatal']}]: ").strip().upper()
                            nuevo_riesgo = "Riesgo" if riesgo_op in ["S"] else "Normal"
                        else:
                            nueva_embarazada = "N"
                            nuevo_trimestre = None
                            nuevo_riesgo = "Normal"
                    
                    nuevo_indigena = input(f"Indígena (S/N) [{p_act['indigena']}]: ").strip().upper() or p_act['indigena']
                    nueva_nacionalidad = input(f"Nacionalidad [{p_act['nacionalidad']}]: ").strip().upper() or p_act['nacionalidad']
                    nueva_direccion = input(f"Dirección [{p_act['direccion']}]: ").strip().upper() or p_act['direccion']
                    nueva_consulta = input(f"Consulta (C/R) [{p_act['consulta']}]: ").strip().upper() or p_act['consulta']
                    
                    cambiar_dx = input(f"¿Cambiar diagnóstico actual? (S/N) [N]: ").strip().upper()
                    if cambiar_dx in ["S"]:
                        dx_datos = seleccionar_diagnostico(nueva_edad)
                        nuevo_dx = dx_datos["diagnostico"]
                        nuevo_codigo_epi = dx_datos["codigo_epi"]
                    else:
                        nuevo_dx = p_act['diagnostico_texto']
                        nuevo_codigo_epi = p_act['codigo_epi']
                    
                    nuevo_tipo_dx = input(f"Tipo Dx (C/P) [{p_act['tipo_diagnostico']}]: ").strip().upper() or p_act['tipo_diagnostico']
                    nueva_conducta = input(f"Conducta [{p_act['conducta']}]: ").strip().upper() or p_act['conducta']
                    
                    conexion = sqlite3.connect(DB_PATH)
                    cursor = conexion.cursor()
                    cursor.execute('''
                        UPDATE pacientes 
                        SET nombre = ?, cedula = ?, edad = ?, sexo = ?, ambito = ?, 
                            embarazada = ?, trimestre = ?, riesgo_prenatal = ?, indigena = ?, 
                            nacionalidad = ?, direccion = ?, consulta = ?, diagnostico_texto = ?, 
                            codigo_epi = ?, tipo_diagnostico = ?, conducta = ?
                        WHERE id = ?
                    ''', (
                        nuevo_nombre, nueva_cedula, nueva_edad, nuevo_sexo, nuevo_ambito,
                        nueva_embarazada, nuevo_trimestre, nuevo_riesgo, nuevo_indigena,
                        nueva_nacionalidad, nueva_direccion, nueva_consulta, nuevo_dx,
                        nuevo_codigo_epi, nuevo_tipo_dx, nueva_conducta, id_pac
                    ))
                    conexion.commit()
                    conexion.close()
                    print("\n[✔] ¡Datos corregidos con éxito!")
                    input("\nPresione ENTER para continuar...")
            else:
                print("\n[!] ID no encontrado.")
                input("\nPresione ENTER para continuar...")
        except ValueError:
            print("\n[!] Debe ingresar un número de ID válido.")
            input("\nPresione ENTER para continuar...")

# ==========================================
# ==========================================
# MENÚ PRINCIPAL
# ==========================================
# ==========================================
# PANTALLA DE BIENVENID

# 🟢 Importas tu función de sincronización manual desde el archivo externo

class PantallaSplash(Screen):
    def __init__(self, **kwargs):
        super(PantallaSplash, self).__init__(**kwargs)
        
        layout = BoxLayout(
            orientation='vertical', 
            padding=[20, 40, 20, 40], 
            spacing=20
        )
        
        lbl_titulo = Label(
            text='ATENCIÓN PRIMARIA', 
            font_size='45sp',         
            bold=True, 
            color=(1, 1, 1, 1),       
            halign='center', 
            valign='middle'
        )
        lbl_titulo.bind(size=lambda s, w: setattr(s, 'text_size', s.size))
        
        lbl_autor = Label(
            text='Panabdon',
            font_size='22sp',         
            bold=True,
            color=(1, 0.84, 0, 1),    
            halign='center', 
            valign='middle'
        )
        lbl_autor.bind(size=lambda s, w: setattr(s, 'text_size', s.size))
        
        # 🟢 El botón verde que ejecuta exclusivamente la Sincronización Manual
        self.btn_entrar = Button(
            text='Verificar licencia',
            size_hint=(0.8, 0.2),
            pos_hint={'center_x': 0.5},
            background_color=(0.1, 0.6, 0.3, 1), # Verde elegante
            color=(1, 1, 1, 1),
            font_size='20sp',
            bold=True
        )
        self.btn_entrar.bind(on_press=self.ejecutar_sincronizacion_manual)
        
        layout.add_widget(lbl_titulo)
        layout.add_widget(lbl_autor)
        layout.add_widget(self.btn_entrar)
        self.add_widget(layout)

    def ejecutar_sincronizacion_manual(self, *args):
        """Se ejecuta UNICAMENTE al presionar el botón verde (carta bajo de la manga)"""
        # Buscamos la cédula guardada en el caché local para hacer la consulta manual real
        cedula_medico = None
        cache = leer_cache_local()
        if cache and cache.get("cedula"):
            cedula_medico = cache.get("cedula")
        
        if cedula_medico:
            # Llamamos a la función de sincronización manual con la cédula real
            resultado = sincronizacion_manual(cedula_medico)
            print(f"Sincronización manual: {resultado.get('mensaje')}")
        else:
            print("No hay cédula previa en caché.")
            
        if self.manager:
            self.manager.current = 'cedula' 

    def on_touch_down(self, touch):
        """Si tocan cualquier parte de la pantalla (excepto el botón), pasa a cédula"""
        # Verificamos si el toque ocurrió dentro de los límites visuales del botón verde
        if self.btn_entrar.collide_point(*touch.pos):
            # Si tocó el botón, dejamos que el botón maneje su propio evento (on_press)
            return super().on_touch_down(touch)
        
        # Si tocó otra parte de la pantalla, salta directo a la pantalla de cédula
        if self.manager and self.manager.current == 'splash':
            self.manager.current = 'cedula'
            return True
            
        return super().on_touch_down(touch)

   
# MENÚ PRINCIPAL
# ==========================================
# ==================================

def menu_principal():
    inicializar_bd()
    semanas_anio = {i: f"Semana Epidemiológica N° {i}" for i in range(1, 56)}

    # --- MÉDICO Y CONSULTORIO OBLIGATORIOS (Se piden una sola vez al arrancar) ---
    limpiar_pantalla()
    print("=" * 60)
    print("   CONTROL DE CONSULTORIO - CONFIGURACIÓN INICIAL")
    print("=" * 60)
    
    while True:
        nombre_medico = input("Nombre y Apellido del Médico (Obligatorio): ").strip().upper()
        if nombre_medico:
            break
        print("  [!] Error: El nombre del médico no puede estar vacío. Por favor, ingréselo.")

    while True:
        nombre_consultorio = input("Nombre del Consultorio (Obligatorio): ").strip().upper()
        if nombre_consultorio:
            break
        print("  [!] Error: El nombre del consultorio no puede estar vacío. Por favor, ingréselo.")
    # -------------------------------------------------------------------------

    while True:
        limpiar_pantalla()
        print("=" * 60)
        print("   CONTROL ANUAL DE CONSULTORIO - 55 SEMANAS EPIDEMIOLÓGICAS")
        print(f"   Médico: {nombre_medico} | Consultorio: {nombre_consultorio}")
        print("=" * 60)
        print("1. Registrar Pacientes en una Semana (Día de Consulta)")
        print("2. Consultar o Modificar Pacientes de una Semana")
        print("3. Generar Consolidado Semanal y EPI-12 de una Semana")
        print("4. Salir")
        
        opcion = input("\nSeleccione una opción (1-4): ").strip()
        
        if opcion == "1":
            limpiar_pantalla()
            print("--- REGISTRO DE PACIENTES POR SEMANA ---")
            
            try:
                sem_input = int(input("Ingrese el número de la semana (1 a 55): ").strip())
                if sem_input not in semanas_anio:
                    print(" [!] Número de semana fuera de rango (1-55).")
                    input("\nPresione ENTER para continuar...")
                    continue
            except ValueError:
                print(" [!] Debe ingresar un número válido.")
                input("\nPresione ENTER para continuar...")
                continue
            
            semana_num = str(sem_input)
            fecha_defecto = datetime.date.today().strftime("%d/%m/%Y")
            fecha = input(f"Fecha de consulta (DD/MM/AAAA) [Por defecto {fecha_defecto}]: ").strip()
            if not fecha:
                fecha = fecha_defecto

            # (Aquí ya no se piden de nuevo porque los tenemos arriba guardaditos)

            while True:
                limpiar_pantalla()
                print(f"=== {semanas_anio[sem_input]} | Fecha: {fecha} ===")
                nombre = input("Nombre y Apellido (ENTER para finalizar el día): ").strip().upper()
                
                if not nombre:
                    break

                cedula = input("Cédula (Ej. V-12345678): ").strip().upper()
                
                while True:
                    edad_in = input("Edad (ej. '5' para años o '2m' para meses): ").strip().lower()
                    try:
                        if "m" in edad_in:
                            meses_val = float(edad_in.replace("m", "").replace(",", ".").strip())
                            val_edad = f"{int(meses_val)}m"
                            break
                        else:
                            val_edad = float(edad_in.replace(",", "."))
                            break
                    except ValueError:
                        print("  [!] Ingrese un valor válido (ej. 5 o 2m).")


                sexo_in = input("Sexo (H = Hombre / M = Mujer): ").strip().upper()
                sexo = "M" if sexo_in in ["M", "F", "MUJER"] else "H"
                
                ambito = input("Ámbito de atención (C = Consultorio / T = Terreno) [C]: ").strip().upper() or "C"
                if ambito not in ["C", "T"]: 
                    ambito = "C"

                embarazada = "N"
                trimestre = None
                riesgo_prenatal = "Normal"
                
                if sexo == "M":
                    emb_op = input("¿Está embarazada? (S/N) [N]: ").strip().upper()
                    if emb_op in ["S", "SI", "SÍ"]:
                        embarazada = "S"
                        trimestre = input("  -> Trimestre de gestación (1, 2 o 3) [1]: ").strip() or "1"
                        riesgo_op = input("  -> ¿Embarazo de Riesgo? (S/N) [N]: ").strip().upper()
                        if riesgo_op in ["S", "SI", "SÍ"]:
                            riesgo_prenatal = "Riesgo"
                puerpera = "N"
                captacion_puerpera = "N"
                
                if sexo == "M":
                    puer_op = input("¿Es consulta de puerperio? (S/N) [N]: ").strip().upper()
                    if puer_op in ["S", "SI", "SÍ"]:
                        puerpera = "S"
                        cap_op = input("  -> ¿Es captación de puérpera? (S/N) [N]: ").strip().upper()
                        if cap_op in ["S", "SI", "SÍ"]:
                            captacion_puerpera = "S"

                indigena = input("Indígena (S/N) [N]: ").strip().upper() or "N"
                indigena = "S" if indigena in ["S", "SI", "SÍ"] else "N"


                nacionalidad = input("Nacionalidad (V/E/C/B/O) [V]: ").strip().upper() or "V"
                direccion = input("Dirección / Comunidad: ").strip().upper()
                consulta = input("Consulta (C = Primera Vez / Reconsulta = R) [C]: ").strip().upper() or "C"
                
                dx_datos = seleccionar_diagnostico(val_edad)
                diagnostico_texto = dx_datos["diagnostico"]
                codigo_epi = dx_datos["codigo_epi"]
                
                tipo_diagnostico = input("Tipo de Diagnóstico (C = Confirmado / P = Presuntivo) [P]: ").strip().upper() or "P"
                conducta = input("Conducta / Tratamiento: ").strip().upper()

                conexion = sqlite3.connect("clinica.db")
                cursor = conexion.cursor()
                cursor.execute('''
                    INSERT INTO pacientes (
                        nombre, cedula, edad, sexo, ambito, embarazada, trimestre, 
                        riesgo_prenatal, puerpera, captacion_puerpera, indigena, nacionalidad, direccion, consulta, 
                        diagnostico_texto, codigo_epi, tipo_diagnostico, conducta, 
                        fecha_consulta, semana_num
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    nombre, cedula, val_edad, sexo, ambito, embarazada, trimestre,
                    riesgo_prenatal, puerpera, captacion_puerpera, indigena, nacionalidad, direccion, consulta,
                    diagnostico_texto, codigo_epi, tipo_diagnostico, conducta,
                    fecha, semana_num
                ))
                conexion.commit()
                conexion.close()

                print(f"\n-> ¡Paciente {nombre} guardado con éxito!")
                
                if input("\n¿Agregar otro paciente en esta misma fecha? (S/N) [S]: ").strip().upper() == "N":
                    break

            # Antes de exportar el Excel diario, pasamos por la revisión previa
            revisar_y_corregir_lote(semana_num, filtro_fecha=fecha)

            # Creamos y aseguramos la carpeta de la semana obligatoriamente aquí
            carpeta_semana = f"Semana_Epidemiologica_{semana_num}"
            os.makedirs(carpeta_semana, exist_ok=True)

            pacientes_dia = obtener_pacientes_por_fecha(fecha)
            if pacientes_dia:
                nombre_salida_diario = os.path.join(carpeta_semana, f"RP-01_Semana{semana_num}_{fecha.replace('/', '-')}.xlsx")
                archivo_diario = generar_excel_rp01(
                    pacientes_dia, 
                    fecha_consulta=fecha, 
                    nombre_archivo=nombre_salida_diario,
                    nombre_medico=nombre_medico,          
                    nombre_consultorio=nombre_consultorio 
                )


        elif opcion == "2":
            limpiar_pantalla()
            print("--- CONSULTAR Y MODIFICAR PACIENTES POR SEMANA ---")
            sem_input = input("Ingrese el número de la semana a revisar (1 a 55): ").strip()
            
            pacientes_semana = obtener_pacientes_por_semana(sem_input)
            if not pacientes_semana:
                print(f"\n[!] No hay pacientes registrados en la Semana N° {sem_input}.")
                input("\nPresione ENTER para volver...")
                continue
            
            print(f"\nSe encontraron {len(pacientes_semana)} pacientes en la Semana N° {sem_input}:\n")
            print(f"{'ID':<4} | {'FECHA':<10} | {'CÉDULA':<12} | {'NOMBRE Y APELLIDO':<25} | {'DIAGNÓSTICO'}")
            print("-" * 75)
            for p in pacientes_semana:
                print(f"{p['id']:<4} | {p['fecha_consulta']:<10} | {p['cedula']:<12} | {p['nombre']:<25} | {p['diagnostico_texto'][:25]}")
            
            print("-" * 75)
            mod_op = input("\n¿Desea modificar o corregir algún paciente de esta lista? (S/N) [N]: ").strip().upper()
            if mod_op == "S":
                try:
                    id_paciente = int(input("Ingrese el ID exacto del paciente que desea actualizar: ").strip())
                    
                    conexion = sqlite3.connect(DB_PATH)
                    conexion.row_factory = sqlite3.Row
                    cursor = conexion.cursor()
                    cursor.execute("SELECT * FROM pacientes WHERE id = ?", (id_paciente,))
                    paciente_act = cursor.fetchone()
                    conexion.close()
                    
                    if paciente_act:
                        limpiar_pantalla()
                        print("\n" + "="*50)
                        print(" MODIFICANDO DATOS DEL PACIENTE (ENTER para conservar)")
                        print("="*50)
                        
                        nuevo_nombre = input(f"Nombre [{paciente_act['nombre']}]: ").strip().upper() or paciente_act['nombre']
                        nueva_cedula = input(f"Cédula [{paciente_act['cedula']}]: ").strip().upper() or paciente_act['cedula']
                        
                        nueva_edad_in = input(f"Edad [{paciente_act['edad']}]: ").strip()
                        nueva_edad = float(nueva_edad_in) if nueva_edad_in else paciente_act['edad']
                        
                        nuevo_sexo = input(f"Sexo (H/M) [{paciente_act['sexo']}]: ").strip().upper() or paciente_act['sexo']
                        nuevo_ambito = input(f"Ámbito (C/T) [{paciente_act['ambito']}]: ").strip().upper() or paciente_act['ambito']
                        
                        nueva_embarazada = paciente_act['embarazada']
                        nuevo_trimestre = paciente_act['trimestre']
                        nuevo_riesgo = paciente_act['riesgo_prenatal']
                        
                        if nuevo_sexo == "M":
                            emb_op = input(f"¿Está embarazada? (SÍ/NO) [{paciente_act['embarazada']}]: ").strip().upper()
                            if emb_op in ["S", "SI", "SÍ"]:
                                nueva_embarazada = "SÍ"
                                nuevo_trimestre = input(f"  -> Trimestre (1, 2, 3) [{paciente_act['trimestre']}]: ").strip() or paciente_act['trimestre']
                                riesgo_op = input(f"  -> ¿Riesgo? (SÍ/NO) [{paciente_act['riesgo_prenatal']}]: ").strip().upper()
                                nuevo_riesgo = "Riesgo" if riesgo_op in ["S", "SI", "SÍ"] else "Normal"
                            else:
                                nueva_embarazada = "NO"
                                nuevo_trimestre = None
                                nuevo_riesgo = "Normal"
                                                

                        # <--- ¡METES EL CÓDIGO DEL PUERPERIO EXACTAMENTE AQUÍ! --->
                        nueva_puerpera = paciente_act['puerpera'] if 'puerpera' in paciente_act else "NO"
                        nueva_captacion_puerpera = paciente_act['captacion_puerpera'] if 'captacion_puerpera' in paciente_act else "NO"
                        
                        if nuevo_sexo == "M":
                            puer_op = input(f"¿Es consulta de puerperio? (SÍ/NO) [{paciente_act.get('puerpera', 'NO')}]: ").strip().upper()
                            if puer_op in ["S", "SI", "SÍ"]:
                                nueva_puerpera = "SÍ"
                                cap_op = input(f"  -> ¿Captación de puérpera? (SÍ/NO) [{paciente_act.get('captacion_puerpera', 'NO')}]: ").strip().upper()
                                nueva_captacion_puerpera = "SÍ" if cap_op in ["S", "SI", "SÍ"] else "NO"
                            else:
                                nueva_puerpera = "NO"
                                nueva_captacion_puerpera = "NO"
                        # -----------------------------------------------                          
                        
                        nuevo_indigena = input(f"Indígena (SÍ/NO) [{paciente_act['indigena']}]: ").strip().upper() or paciente_act['indigena']
                        nueva_nacionalidad = input(f"Nacionalidad [{paciente_act['nacionalidad']}]: ").strip().upper() or paciente_act['nacionalidad']
                        nueva_direccion = input(f"Dirección [{paciente_act['direccion']}]: ").strip().upper() or paciente_act['direccion']
                        nueva_consulta = input(f"Consulta (C/R) [{paciente_act['consulta']}]: ").strip().upper() or paciente_act['consulta']
                        
                        cambiar_dx = input(f"\n¿Cambiar diagnóstico '{paciente_act['diagnostico_texto']}'? (S/N) [N]: ").strip().upper()
                        if cambiar_dx in ["S", "SI", "SÍ"]:
                            dx_datos = seleccionar_diagnostico(nueva_edad)
                            nuevo_dx = dx_datos["diagnostico"]
                            nuevo_codigo_epi = dx_datos["codigo_epi"]
                        else:
                            nuevo_dx = paciente_act['diagnostico_texto']
                            nuevo_codigo_epi = paciente_act['codigo_epi']
                        
                        nuevo_tipo_dx = input(f"Tipo Dx (C/P) [{paciente_act['tipo_diagnostico']}]: ").strip().upper() or paciente_act['tipo_diagnostico']
                        nueva_conducta = input(f"Conducta [{paciente_act['conducta']}]: ").strip().upper() or paciente_act['conducta']
                        
                        conexion = sqlite3.connect(DB_PATH)
                        cursor = conexion.cursor()
                        cursor.execute('''
                            UPDATE pacientes 
                            SET nombre = ?, cedula = ?, edad = ?, sexo = ?, ambito = ?, 
                                embarazada = ?, trimestre = ?, riesgo_prenatal = ?, indigena = ?, 
                                nacionalidad = ?, direccion = ?, consulta = ?, diagnostico_texto = ?, 
                                codigo_epi = ?, tipo_diagnostico = ?, conducta = ?
                            WHERE id = ?
                        ''', (
                            nuevo_nombre, nueva_cedula, nueva_edad, nuevo_sexo, nuevo_ambito,
                            nueva_embarazada, nuevo_trimestre, nuevo_riesgo, nuevo_indigena,
                            nueva_nacionalidad, nueva_direccion, nueva_consulta, nuevo_dx,
                            nuevo_codigo_epi, nuevo_tipo_dx, nueva_conducta, id_paciente
                        ))
                        conexion.commit()
                        conexion.close()
                        print("\n[✔] ¡Registro actualizado con éxito!")
                    else:
                        print("\n[!] ID de paciente no encontrado.")
                except ValueError:
                    print("\n[!] Debe ingresar un número de ID válido.")
            
            input("\nPresione ENTER para volver al menú...")

        elif opcion == "3":
            limpiar_pantalla()
            print("--- GENERAR CONSOLIDADO SEMANAL Y EPI-12 ---")
            sem_input = input("Ingrese el número de la semana a consolidar (1 a 55): ").strip()
                        # --- DATOS MANUALES DEL MÉDICO Y CONSULTORIO ---
            # Validamos primero si hay pacientes para esa semana
            pacientes_semana = obtener_pacientes_por_semana(sem_input)
            if not pacientes_semana:
                print(f"\n[!] No hay registros cargados para la Semana N° {sem_input}.")
                input("\nPresione ENTER para volver...")
                continue

            print(f"\n[i] Usando datos de sesión:")
            print(f"    - Médico: {nombre_medico}")
            print(f"    - Consultorio: {nombre_consultorio}")
            input("\nPresione ENTER para continuar con la revisión previa...")



            pacientes_semana = obtener_pacientes_por_semana(sem_input)
            if not pacientes_semana:
                print(f"\n[!] No hay registros cargados para la Semana N° {sem_input}.")
                input("\nPresione ENTER para volver...")
                continue

            # Antes de exportar el consolidado semanal, activamos la revisión previa
            revisar_y_corregir_lote(sem_input, filtro_fecha=None)

            # Recargamos la lista por si hubo cambios o correcciones y creamos la carpeta
            carpeta_semana = f"Semana_Epidemiologica_{sem_input}"
            os.makedirs(carpeta_semana, exist_ok=True)

            pacientes_semana_actualizados = obtener_pacientes_por_semana(sem_input)
            
            # --- RADIOGRAFÍA PARA CAZAR AL FANTASMA DEL 1ER TRIMESTRE ---
            print("\n" + "="*50)
            print("       RADIOGRAFÍA DE PACIENTES CONTADOS")
            print("="*50)
            for i, p in enumerate(pacientes_semana_actualizados):
                diag = str(p.get("diagnostico_texto", "")).upper()
                edad = p.get("edad", 0)
                emb = str(p.get("embarazada", "")).upper()
                trim = str(p.get("trimestre", ""))
                
                # Vamos a mostrar todo lo que huela a embarazo o trimestre
                if emb == "SÍ" or trim in ["1", "2", "3", "1er", "2do", "3er"]:
                    print(f"[{i}] ID: {p.get('id')} | Nombre: {p.get('nombre')} | Edad: {edad} | Emb: '{emb}' | Trim: '{trim}' | Diag: {diag}")
            print("="*50)
            input("\nPresione ENTER para continuar con la captura de Actividades y Programas...")
            # -------------------------------------------------------------

            # --- AQUÍ INTEGRAMOS EL MENÚ DE ACTIVIDADES Y PROGRAMAS ---
            datos_conceptos_extra = actividades.preguntar_conceptos()      # <--- ¡Esta es la única línea nueva que faltaba aquí!
            datos_actividades_extra = actividades.preguntar_actividades_y_programas()
            # ---------------------------------------------------------

            if pacientes_semana_actualizados:
                nombre_salida_semanal = os.path.join(carpeta_semana, f"Consolidado_Semanal_S{sem_input}.xlsx")
                
                # Le pasamos los datos al generador de Excel incluyendo médico y consultorio
                archivo_semanal = generar_consolidado_semanal(
                    pacientes_semana=pacientes_semana_actualizados,
                    semana_num=sem_input,
                    asic="PORTUARIO",
                    estado="Carabobo",
                    nombre_archivo=nombre_salida_semanal,
                    datos_conceptos=datos_conceptos_extra,     
                    datos_actividades=datos_actividades_extra,  
                    nombre_medico=nombre_medico,          
                    nombre_consultorio=nombre_consultorio 
                )
                print(f"\n[OK] Consolidado guardado en carpeta: {archivo_semanal}")

                # --- NUEVO: GENERAR AUTOMÁTICAMENTE LAS HOJAS DIARIAS (RP-01) DE LA SEMANA ---
                fechas_unicas = set(p.get("fecha_consulta") for p in pacientes_semana_actualizados if p.get("fecha_consulta"))
                for fecha_dia in fechas_unicas:
                    pacientes_dia = obtener_pacientes_por_fecha(fecha_dia)     
                    if pacientes_dia:
                        # Limpiamos la fecha para que el nombre del archivo no falle con las barras (ej: 09/03/2026 -> 09-03-2026)
                        fecha_limpia = fecha_dia.replace("/", "-")
                        nombre_salida_diario = os.path.join(carpeta_semana, f"RP-01_Semana{sem_input}_{fecha_limpia}.xlsx")
                        
                        generar_excel_rp01(
                            pacientes_dia, 
                            fecha_consulta=fecha_dia, 
                            nombre_archivo=nombre_salida_diario,
                            nombre_medico=nombre_medico,          
                            nombre_consultorio=nombre_consultorio 
                        )
                print(f"[OK] ¡Todas las hojas de consulta diaria (RP-01) de la semana fueron exportadas con éxito!")


                   

        elif opcion == "4":
            print("\n¡Saliendo del sistema de consulta. Éxitos en tu guardia, mi amor!")
            break

# --- PANTALLA DE CONFIGURACIÓN INICIAL (Médico y Consultorio) ---
class PantallaConfiguracion(Screen):
    def __init__(self, **kwargs):
        super(PantallaConfiguracion, self).__init__(**kwargs)
        
        # Usamos un ScrollView por si el teclado abre en el teléfono y evita que se tape contenido
        scroll = ScrollView()
        
        # Layout principal vertical centrado
        layout = BoxLayout(
            orientation='vertical', 
            padding=30, 
            spacing=20,
            size_hint_y=None
        )
        # Hacemos que la altura del layout se ajuste automáticamente a los elementos internos
        layout.bind(minimum_height=layout.setter('height'))
        
        # Espaciador superior flexible para empujar el contenido hacia el centro estéticamente
        layout.add_widget(Widget(size_hint_y=None, height=20))
        
        # Título formal centrado correctamente con text_size (altura ajustada a 95)
        self.lbl_titulo = Label(
            text="CONTROL DE CONSULTORIO\nConfiguración Inicial", 
            halign='center', 
            valign='middle',
            font_size=sp(20),
            size_hint_y=None,
            height=95
        )
        # Importante para que Kivy respete el halign='center'
        self.lbl_titulo.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(self.lbl_titulo)
        
        # Espacio
        layout.add_widget(Widget(size_hint_y=None, height=10))
        
        # Campo para el Nombre del Médico (etiqueta con height de 55)
        lbl_med = Label(
            text="Nombre y Apellido del Médico:", 
            halign='left', 
            size_hint_y=None, 
            height=55, 
            font_size=sp(15)
        )
        lbl_med.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_med)
        
        # TextInput con height holgado de 65 para que no corte la letra escrita
        self.input_medico = TextInput(
            text='', 
            multiline=False, 
            size_hint_x=1.0,
            size_hint_y=None, 
            height=65,
            font_size=sp(16),
            write_tab=False,
            padding_y=[8, 8]
        )
        layout.add_widget(self.input_medico)
        
        # Campo para el Nombre del Consultorio (etiqueta con height de 55)
        lbl_cons = Label(
            text="Nombre del Consultorio:", 
            halign='left', 
            size_hint_y=None, 
            height=55, 
            font_size=sp(15)
        )
        lbl_cons.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_cons)
        
        # TextInput con height holgado de 65
        self.input_consultorio = TextInput(
            text='', 
            multiline=False, 
            size_hint_x=1.0,
            size_hint_y=None, 
            height=65,
            font_size=sp(16),
            write_tab=False,
            padding_y=[8, 8]
        )
        layout.add_widget(self.input_consultorio)
        
        # ==========================================
        # AQUÍ METEMOS LOS 3 NUEVOS CAMPOS (ESTADO, MUNICIPIO Y ASIC)
        # ==========================================
        
        # --- CAMPO: Estado ---
        lbl_est = Label(
            text="Estado:", 
            halign='left', 
            size_hint_y=None, 
            height=55, 
            font_size=sp(15)
        )
        lbl_est.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_est)
        
        self.input_estado = TextInput(
            text='Carabobo', 
            multiline=False, 
            size_hint_x=1.0,
            size_hint_y=None, 
            height=65,
            font_size=sp(16),
            write_tab=False,
            padding_y=[8, 8]
        )
        layout.add_widget(self.input_estado)

        # --- CAMPO: Municipio ---
        lbl_mun = Label(
            text="Municipio:", 
            halign='left', 
            size_hint_y=None, 
            height=55, 
            font_size=sp(15)
        )
        lbl_mun.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_mun)
        
        self.input_municipio = TextInput(
            text='', 
            multiline=False, 
            size_hint_x=1.0,
            size_hint_y=None, 
            height=65,
            font_size=sp(16),
            write_tab=False,
            padding_y=[8, 8]
        )
        layout.add_widget(self.input_municipio)

        # --- CAMPO: ASIC ---
        lbl_asic = Label(
            text="ASIC (Área de Salud Integral Com.):", 
            halign='left', 
            size_hint_y=None, 
            height=55, 
            font_size=sp(15)
        )
        lbl_asic.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_asic)
        
        self.input_asic = TextInput(
            text='', 
            multiline=False, 
            size_hint_x=1.0,
            size_hint_y=None, 
            height=65,
            font_size=sp(16),
            write_tab=False,
            padding_y=[8, 8]
        )
        layout.add_widget(self.input_asic)
        
        # ==========================================
        # FIN DE LOS NUEVOS CAMPOS
        # ==========================================

        # Espacio antes del botón
        # Espacio antes de los botones
        layout.add_widget(Widget(size_hint_y=None, height=15))
        
        # Contenedor horizontal para colocar ambos botones lado a lado
        btn_layout = BoxLayout(orientation='horizontal', spacing=10, size_hint_y=None, height=65)
        
        # Botón Anterior (apunta a 'cedula')
        btn_ant = Button(
            text="⬅️ Anterior", 
            background_color=(0.5, 0.5, 0.5, 1), 
            font_size=sp(16), 
            size_hint_y=None, 
            height=65
        )
        btn_ant.bind(on_press=self.ir_anterior)
        
        # Botón Guardar y Continuar
        btn_continuar = Button(
            text="Guardar y Continuar", 
            size_hint_y=None, 
            height=65,
            font_size=sp(16),
            background_color=(0.1, 0.6, 0.8, 1)
        )
        btn_continuar.bind(on_press=self.guardar_configuracion)
        
        # Agregamos los dos botones a la botonera horizontal
        btn_layout.add_widget(btn_ant)
        btn_layout.add_widget(btn_continuar)
        
        # Añadimos la botonera al layout principal
        layout.add_widget(btn_layout)
        
        # Espaciador final para que el scroll respire cómodo abajo
        layout.add_widget(Widget(size_hint_y=None, height=50))
        
        scroll.add_widget(layout)
        self.add_widget(scroll)

        
       
    def on_enter(self):
        """Se ejecuta automáticamente al entrar a esta pantalla para cargar la última configuración guardada."""
        datos = cargar_config_db()  # Viene de encabezado4.py
        
        if datos and len(datos) > 0 and datos[0]:
            # 🟢 ¡BLINDAJE TOTAL! Ya hay un médico registrado. Candado puesto y llave tirada al mar.
            self.input_medico.text = datos[0]
            self.input_medico.readonly = True  # ¡Imposible de modificar por teclado!
            self.input_medico.background_color = (0.85, 0.85, 0.85, 1)  # Un grisecito elegante de solo lectura
            
            # Los demás campos sí quedan abiertos por si el doctor se muda de consultorio o ASIC
            self.input_consultorio.text = datos[1] if len(datos) > 1 and datos[1] else ""
            self.input_estado.text = datos[2] if len(datos) > 2 and datos[2] else "Carabobo"
            self.input_municipio.text = datos[3] if len(datos) > 3 and datos[3] else ""
            self.input_asic.text = datos[4] if len(datos) > 4 and datos[4] else ""
        else:
            # 🟢 ¡ZONA VIRGEN! Primera vez que se instala: se deja escribir el nombre una sola vez.
            self.input_medico.text = ""
            self.input_medico.readonly = False
            self.input_medico.background_color = (1, 1, 1, 1)  # Blanco normal
            
            self.input_consultorio.text = ""
            self.input_estado.text = "Carabobo"
            self.input_municipio.text = ""
            self.input_asic.text = ""

        
    def guardar_configuracion(self, instance):
        medico = self.input_medico.text.strip().upper()
        consultorio = self.input_consultorio.text.strip().upper()
        estado = self.input_estado.text.strip().upper()
        municipio = self.input_municipio.text.strip().upper()
        asic = self.input_asic.text.strip().upper()
        
        if not medico or not consultorio or not estado or not municipio or not asic:
            print("[!] Error: Todos los campos son obligatorios")
            return
            
        # Guardamos en la base de datos física mediante la función de encabezado4.py
        guardar_config_db(medico, consultorio, estado, municipio, asic)
        
        app = App.get_running_app()
        app.nombre_medico = medico
        app.nombre_consultorio = consultorio
        app.estado = estado
        app.municipio = municipio
        app.asic = asic
        
        print(f"[✔] Configuración guardada con éxito para: {medico}")
        self.manager.current = 'menu'
    def ir_anterior(self, instance):
        """Devuelve al usuario a la pantalla de la cédula"""
        if self.manager:
            self.manager.current = 'cedula'
        
class PantallaMenuConducta(Screen):
    def __init__(self, **kwargs):
        super(PantallaMenuConducta, self).__init__(**kwargs)
        
        scroll = ScrollView()
        layout = BoxLayout(orientation='vertical', padding=30, spacing=15, size_hint_y=None)
        layout.bind(minimum_height=layout.setter('height'))
        
        layout.add_widget(Widget(size_hint_y=None, height=10))
        
        lbl_titulo = Label(
            text="[b]Seleccionar Conducta / Tratamiento[/b]",
            markup=True,
            halign='center',
            valign='middle',
            font_size=sp(18),
            size_hint_y=None,
            height=65
        )
        lbl_titulo.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_titulo)
        
        layout.add_widget(Widget(size_hint_y=None, height=10))
        
        opciones_conducta = [
            "Tratamiento medico",
            "Referencia medica",
            "Captacion",
            "Control",
            "Medidas generales",
            "Medidas especificas",
            "Puericultura",
            "medidas preventivas",
            "CARA A CARA",
            "Charla educativa"
        ]
        
        for op in opciones_conducta:
            btn = Button(
                text=op,
                size_hint_y=None,
                height=60,
                font_size=sp(15),
                background_color=(0.2, 0.5, 0.8, 1)
            )
            # Usamos partial para fijar el texto limpiamente sin fallos de lambda
            btn.bind(on_press=partial(self.seleccionar_opcion, op))
            layout.add_widget(btn)
            
        layout.add_widget(Widget(size_hint_y=None, height=15))
        
        btn_volver = Button(
            text="⬅️ Volver",
            size_hint_y=None,
            height=60,
            font_size=sp(15),
            background_color=(0.5, 0.5, 0.5, 1)
        )
        # 🎯 RADIOGRAFÍA: Aquí cambiamos el lambda viejo por tu función volver_atras
        btn_volver.bind(on_press=self.volver_atras)
        layout.add_widget(btn_volver)
        
        layout.add_widget(Widget(size_hint_y=None, height=100))
        scroll.add_widget(layout)
        self.add_widget(scroll)

    def seleccionar_opcion(self, texto, instance):
        app = App.get_running_app()
        app.conducta_seleccionada = texto
        
        # 🛑 Preguntamos de dónde venimos para el menú de selección
        if hasattr(app, 'origen_conducta') and app.origen_conducta == 'modificar':
            self.manager.current = 'modificar_paso4'
        else:
            self.manager.current = 'registro_paso4'

    def volver_atras(self, instance):
        app = App.get_running_app()
        # 🛑 Preguntamos de dónde venimos para el botón volver
        if hasattr(app, 'origen_conducta') and app.origen_conducta == 'modificar':
            self.manager.current = 'modificar_paso4'
        else:
            self.manager.current = 'registro_paso4'

# --- PANTALLA DE MENÚ PRINCIPAL TÁCTIL ---
class PantallaMenu(Screen):
    def __init__(self, **kwargs):
        super(PantallaMenu, self).__init__(**kwargs)
        
        scroll = ScrollView()
        layout = BoxLayout(orientation='vertical', padding=30, spacing=25, size_hint_y=None)
        layout.bind(minimum_height=layout.setter('height'))
        
        layout.add_widget(Widget(size_hint_y=None, height=20))
        
        lbl_menu = Label(
            text="MENÚ PRINCIPAL\nAtención Primaria de Salud", 
            halign='center', 
            valign='middle',
            font_size=sp(20),
            size_hint_y=None,
            height=95
        )
        lbl_menu.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_menu)
        
        layout.add_widget(Widget(size_hint_y=None, height=10))
        
        btn_reg = Button(text="1. Registrar Pacientes en una Semana", size_hint_y=None, height=65, font_size=sp(15))
        btn_reg.bind(on_press=self.ir_registro)
        layout.add_widget(btn_reg)
        
        btn_con = Button(text="2. Consultar o Modificar Pacientes", size_hint_y=None, height=65, font_size=sp(15))
        btn_con.bind(on_press=self.ir_consultar)
        layout.add_widget(btn_con)
        
        btn_excel = Button(
            text="3. Generar Consolidado y Exportar Excel", 
            size_hint_y=None, 
            height=65,
            font_size=sp(15),
            background_color=(0.2, 0.7, 0.3, 1)
        )
        btn_excel.bind(on_press=self.ir_excel)
        layout.add_widget(btn_excel)
        
        # Botón para regresar
        btn_regresar = Button(
            text="← Volver a Configuración", 
            size_hint_y=None, 
            height=55,
            font_size=sp(15),
            background_color=(0.7, 0.3, 0.3, 1)
        )
        btn_regresar.bind(on_press=self.regresar_configuracion)
        layout.add_widget(btn_regresar)

        scroll.add_widget(layout)
        self.add_widget(scroll)


    def ir_registro(self, instance):
        self.manager.current = 'registro_paso1'

    def ir_consultar(self, instance):
        # ¡Cambiado para apuntar al Paso 1 de la nueva modificación en 4 pantallas!
        self.manager.current = 'modificar_paso1'


    def ir_excel(self, instance):
        app = App.get_running_app()
        print(f"[✔] Generando Excel para: {app.nombre_medico} en {app.nombre_consultorio}")
        
        # ¡Aquí estaba el error! Apuntamos al Paso 1 de exportación en lugar de la pantalla vieja
        self.manager.current = 'exportar_paso1'
        
    def regresar_configuracion(self, instance):
        self.manager.current = 'config'

class PantallaExportarPaso1(Screen):
    def __init__(self, **kwargs):
        super(PantallaExportarPaso1, self).__init__(**kwargs)
        
        scroll = ScrollView(scroll_timeout=250, scroll_distance=15)
        layout = BoxLayout(orientation='vertical', padding=30, spacing=20, size_hint_y=None)
        layout.bind(minimum_height=layout.setter('height'))
        
        layout.add_widget(Widget(size_hint_y=None, height=10))
        
        lbl_titulo = Label(
            text="[b]Mortalidad y Población Vulnerable[/b]",
            markup=True,
            halign='center', 
            valign='middle',
            font_size=sp(18),
            size_hint_y=None,
            height=75
        )
        lbl_titulo.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_titulo)
        
        lbl_semana = Label(
            text="[b]N° Semana:[/b]", 
            markup=True, 
            halign='left', 
            size_hint_y=None, 
            height=45, 
            font_size=sp(15)
        )
        lbl_semana.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_semana)

        self.input_semana_reporte = TextInput(
            text="1", 
            multiline=False, 
            input_filter='int', 
            size_hint_x=1.0, 
            size_hint_y=None, 
            height=65, 
            font_size=sp(16),
            write_tab=False
        )
        layout.add_widget(self.input_semana_reporte)

        
        self.inputs = {}
        
        # 🟢 Preguntas corregidas exactamente como en el consolidado
        campos_paso1 = [
            ("fila_6", "De ellos postrados y encamados:"),
            ("fila_10", "Fallecidos de 60 años y más:"),
            ("fila_20", "Embarazadas fallecidas:"),
            ("fila_38", "RN de ellos bajo peso:"),
            ("fila_39", "Fallecidos menores de un año:")
        ]
        
        for clave, texto_label in campos_paso1:
            lbl = Label(text=texto_label, halign='left', size_hint_y=None, height=55, font_size=sp(15))
            lbl.bind(size=lambda s, w: setattr(s, 'text_size', w))
            layout.add_widget(lbl)
            
            valor_guardado = str(CACHE_MANUAL.get(clave, "0"))
            txt = TextInput(
                text=valor_guardado, 
                multiline=False, 
                input_filter='int', 
                size_hint_x=1.0, 
                size_hint_y=None, 
                height=65, 
                font_size=sp(16), 
                write_tab=False
            )
            txt.bind(text=lambda instance, val, k=clave: CACHE_MANUAL.update({k: val}))
            layout.add_widget(txt)
            self.inputs[clave] = txt
            
        layout.add_widget(Widget(size_hint_y=None, height=15))
        
        btn_layout = BoxLayout(orientation='horizontal', spacing=10, size_hint_y=None, height=65)
        btn_volver = Button(text="Volver al Menú", background_color=(0.8, 0.3, 0.3, 1), font_size=sp(16), size_hint_y=None, height=65)
        btn_volver.bind(on_press=lambda x: setattr(self.manager, 'current', 'menu'))
        
        btn_sig = Button(text="Siguiente ➡️", background_color=(0.2, 0.6, 0.8, 1), font_size=sp(16), size_hint_y=None, height=65)
        btn_sig.bind(on_press=lambda x: setattr(self.manager, 'current', 'exportar_paso2'))
        
        btn_layout.add_widget(btn_volver)
        btn_layout.add_widget(btn_sig)
        layout.add_widget(btn_layout)
        
        layout.add_widget(Widget(size_hint_y=None, height=150))
        scroll.add_widget(layout)
        self.add_widget(scroll)

    def limpiar_cache_y_pantalla(self):
        """🟢 Vacía el CACHE_MANUAL y pone los campos en 0 al exportar."""
        for clave in self.inputs:
            CACHE_MANUAL[clave] = "0"
            self.inputs[clave].text = "0"


# --- EXPORTAR / CONSOLIDADO - PASO  ---
            
# --- PANTALLA EXPORTAR PASO 2 (Remitidos y Total 43) ---
# --- PANTALLA EXPORTAR PASO 2 (Remitidos y Total 43) ---
class PantallaExportarPaso2(Screen):
    def __init__(self, **kwargs):
        super(PantallaExportarPaso2, self).__init__(**kwargs)
        
        scroll = ScrollView(scroll_timeout=250, scroll_distance=15)
        layout = BoxLayout(orientation='vertical', padding=30, spacing=20, size_hint_y=None)
        layout.bind(minimum_height=layout.setter('height'))
        
        layout.add_widget(Widget(size_hint_y=None, height=10))
        
        lbl_titulo = Label(
            text="[b]Remitidos y Enfermedades Crónicas[/b]",
            markup=True,
            halign='center', 
            valign='middle',
            font_size=sp(18),
            size_hint_y=None,
            height=75
        )
        lbl_titulo.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_titulo)
        
        self.inputs = {}
        
        # 1. Sección de Pacientes Remitidos (Entrada manual 40, 41, 42)
        lbl_seccion_rem = Label(text="[b]Remitidos:[/b]", markup=True, halign='left', size_hint_y=None, height=55, font_size=sp(16))
        lbl_seccion_rem.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_seccion_rem)
        
        campos_remitidos = [
            ("fila_40", "CDI:"),
            ("fila_41", "Otras instituciones estatales:"),
            ("fila_42", "Instituciones privadas:")
        ]
        
        for clave, texto in campos_remitidos:
            lbl = Label(text=texto, halign='left', size_hint_y=None, height=55, font_size=sp(15))
            lbl.bind(size=lambda s, w: setattr(s, 'text_size', w))
            layout.add_widget(lbl)
            
            txt = TextInput(
                text="0", 
                multiline=False, 
                input_filter='int', 
                size_hint_x=1.0, 
                size_hint_y=None, 
                height=65, 
                font_size=sp(16), 
                write_tab=False, 
                padding_y=[8, 8]
            )
            txt.bind(text=lambda instance, val, k=clave: self.actualizar_remitido(k, val))
            layout.add_widget(txt)
            self.inputs[clave] = txt

        # Motor automático para la Fila 43 (Total Remitidos)
        self.lbl_total_43 = Label(
            text="43. Total Remitidos: 0", 
            halign='left', 
            size_hint_y=None, 
            height=55, 
            font_size=sp(15), 
            color=(0.2, 0.8, 0.4, 1)
        )
        # self.lbl_total_43.bind(size=lambda s, w: setattr(s, 'text_size', w))
        # layout.add_widget(self.lbl_total_43)
        
        CACHE_MANUAL["fila_43"] = CACHE_MANUAL.get("fila_43", 0)
        self.recalcular_total_43()

        layout.add_widget(Widget(size_hint_y=None, height=15))
        
        btn_layout = BoxLayout(orientation='horizontal', spacing=10, size_hint_y=None, height=65)
        btn_ant = Button(text="⬅️ Anterior", background_color=(0.5, 0.5, 0.5, 1), font_size=sp(16), size_hint_y=None, height=65)
        btn_ant.bind(on_press=lambda x: setattr(self.manager, 'current', 'exportar_paso1'))
        
        btn_sig = Button(text="Siguiente ➡️", background_color=(0.2, 0.6, 0.8, 1), font_size=sp(16), size_hint_y=None, height=65)
        btn_sig.bind(on_press=lambda x: setattr(self.manager, 'current', 'exportar_paso4'))
        
        btn_layout.add_widget(btn_ant)
        btn_layout.add_widget(btn_sig)
        layout.add_widget(btn_layout)
        
        layout.add_widget(Widget(size_hint_y=None, height=150))
        scroll.add_widget(layout)
        self.add_widget(scroll)

    # 🟢 AQUÍ ESTá LA FUNCIÓN QUE FALTABA: Limpia y resetea al entrar a la pantalla
    def on_pre_enter(self, *args):
        for clave, txt_input in self.inputs.items():
            txt_input.text = "0"
            CACHE_MANUAL[clave] = 0
        CACHE_MANUAL["fila_43"] = 0
        if hasattr(self, 'lbl_total_43'):
            self.lbl_total_43.text = "43. Total Remitidos: 0"

    def actualizar_remitido(self, clave, val):
        CACHE_MANUAL[clave] = val
        self.recalcular_total_43()

    def recalcular_total_43(self):
        try:
            v40 = int(CACHE_MANUAL.get("fila_40", 0) or 0)
        except ValueError:
            v40 = 0
        try:
            v41 = int(CACHE_MANUAL.get("fila_41", 0) or 0)
        except ValueError:
            v41 = 0
        try:
            v42 = int(CACHE_MANUAL.get("fila_42", 0) or 0)
        except ValueError:
            v42 = 0
            
        total_43 = v40 + v41 + v42
        CACHE_MANUAL["fila_43"] = total_43
        if hasattr(self, 'lbl_total_43'):
            self.lbl_total_43.text = f"43. Total Remitidos: {total_43}"



# --- EXPORTAR / CONSOLIDADO - PASO 4 (Parte 1 de Conceptos: Filas 5 a 9) ---
class PantallaExportarPaso4(Screen):
    def __init__(self, **kwargs):
        super(PantallaExportarPaso4, self).__init__(**kwargs)
        
        scroll = ScrollView(scroll_timeout=250, scroll_distance=15)
        layout = BoxLayout(orientation='vertical', padding=30, spacing=20, size_hint_y=None)
        layout.bind(minimum_height=layout.setter('height'))
        
        layout.add_widget(Widget(size_hint_y=None, height=10))
        
        lbl_titulo = Label(
            text="[b]Conceptos (Parte 1)[/b]",
            markup=True,
            halign='center', 
            valign='middle',
            font_size=sp(18),
            size_hint_y=None,
            height=75
        )
        lbl_titulo.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_titulo)
        
        self.inputs = {}
        
        campos_paso4 = [
            ("fila_5", "Charlas educativas:"),
            ("fila_6", "Partos realizados o asistidos:"),
            ("fila_7", "Vidas salvadas:"),
            ("fila_8", "Nacidos vivos:"),
            ("fila_9", "Encamados niños y adultos):")
        ]

        for clave, texto in campos_paso4:
            lbl = Label(text=texto, halign='left', size_hint_y=None, height=55, font_size=sp(15))
            lbl.bind(size=lambda s, w: setattr(s, 'text_size', w))
            layout.add_widget(lbl)
            
            valor_guardado = str(CACHE_MANUAL.get(clave, "0"))
            txt = TextInput(
                text=valor_guardado, 
                multiline=False, 
                input_filter='int', 
                size_hint_x=1.0, 
                size_hint_y=None, 
                height=65, 
                font_size=sp(16), 
                write_tab=False
            )
            txt.bind(text=lambda instance, val, k=clave: CACHE_MANUAL.update({k: val}))
            layout.add_widget(txt)
            self.inputs[clave] = txt
            
        layout.add_widget(Widget(size_hint_y=None, height=15))
        
        # Botonera de navegación (Anterior hacia el Paso 3 / Siguiente hacia el Paso 5)
        btn_layout = BoxLayout(orientation='horizontal', spacing=10, size_hint_y=None, height=65)
        
        btn_ant = Button(
            text="⬅️ Anterior",
            size_hint_y=None,
            height=65,
            background_color=(0.5, 0.5, 0.5, 1),
            font_size=sp(16)
        )
        btn_ant.bind(on_press=lambda x: setattr(self.manager, 'current', 'exportar_paso2'))
        
        btn_sig = Button(
            text="Siguiente ➡️",
            size_hint_y=None,
            height=65,
            background_color=(0.2, 0.6, 0.8, 1),
            font_size=sp(16)
        )
        btn_sig.bind(on_press=lambda x: setattr(self.manager, 'current', 'exportar_paso5'))
        
        btn_layout.add_widget(btn_ant)
        btn_layout.add_widget(btn_sig)
        layout.add_widget(btn_layout)
        
        layout.add_widget(Widget(size_hint_y=None, height=150))
        scroll.add_widget(layout)
        self.add_widget(scroll)


# --- EXPORTAR / CONSOLIDADO - PASO 5 (Parte 2 de Conceptos: Filas 10 a 14) ---
class PantallaExportarPaso5(Screen):
    def __init__(self, **kwargs):
        super(PantallaExportarPaso5, self).__init__(**kwargs)
        
        scroll = ScrollView(scroll_timeout=250, scroll_distance=15)
        layout = BoxLayout(orientation='vertical', padding=30, spacing=20, size_hint_y=None)
        layout.bind(minimum_height=layout.setter('height'))
        
        layout.add_widget(Widget(size_hint_y=None, height=10))
        
        lbl_titulo = Label(
            text="[b]Conceptos (Parte 2)[/b]",
            markup=True,
            halign='center', 
            valign='middle',
            font_size=sp(18),
            size_hint_y=None,
            height=75
        )
        lbl_titulo.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_titulo)
        
        self.inputs = {}
        
        campos_paso5 = [
            ("fila_10", "Ingresos en el hogar:"),
            ("fila_11", "Altas de ingresos en el hogar:"),
            ("fila_12", "Examen de mama:"),
            ("fila_13", "DIU colocados:"),
            ("fila_14", "Anticonceptivos entregados:")
        ]

        for clave, texto in campos_paso5:
            lbl = Label(text=texto, halign='left', size_hint_y=None, height=55, font_size=sp(15))
            lbl.bind(size=lambda s, w: setattr(s, 'text_size', w))
            layout.add_widget(lbl)
            
            valor_guardado = str(CACHE_MANUAL.get(clave, "0"))
            txt = TextInput(
                text=valor_guardado, 
                multiline=False, 
                input_filter='int', 
                size_hint_x=1.0, 
                size_hint_y=None, 
                height=65, 
                font_size=sp(16), 
                write_tab=False
            )
            txt.bind(text=lambda instance, val, k=clave: CACHE_MANUAL.update({k: val}))
            layout.add_widget(txt)
            self.inputs[clave] = txt
            
        layout.add_widget(Widget(size_hint_y=None, height=15))
        
        # Botonera de navegación (Anterior hacia el Paso 4 / Siguiente hacia el paso final de exportación)
        btn_layout = BoxLayout(orientation='horizontal', spacing=10, size_hint_y=None, height=65)
        
        btn_ant = Button(
            text="⬅️ Anterior",
            size_hint_y=None,
            height=65,
            background_color=(0.5, 0.5, 0.5, 1),
            font_size=sp(16)
        )
        btn_ant.bind(on_press=lambda x: setattr(self.manager, 'current', 'exportar_paso4'))
        
        btn_sig = Button(
            text="Siguiente ➡️",
            size_hint_y=None,
            height=65,
            background_color=(0.2, 0.6, 0.8, 1),
            font_size=sp(16)
        )
        btn_sig.bind(on_press=lambda x: setattr(self.manager, 'current', 'exportar_paso6'))
        
        btn_layout.add_widget(btn_ant)
        btn_layout.add_widget(btn_sig)
        layout.add_widget(btn_layout)
        
        layout.add_widget(Widget(size_hint_y=None, height=150))
        scroll.add_widget(layout)
        self.add_widget(scroll)
        
# --- EXPORTAR / CONSOLIDADO - PASO 6 (Parte 1 de Bases y Urbanismos: 4 preguntas) ---
class PantallaExportarPaso6(Screen):
    def __init__(self, **kwargs):
        super(PantallaExportarPaso6, self).__init__(**kwargs)
        
        scroll = ScrollView(scroll_timeout=250, scroll_distance=15)
        layout = BoxLayout(orientation='vertical', padding=30, spacing=20, size_hint_y=None)
        layout.bind(minimum_height=layout.setter('height'))
        
        layout.add_widget(Widget(size_hint_y=None, height=10))
        
        lbl_titulo = Label(
            text="[b]Bases de Misiones y Urbanismos (Parte 1)[/b]",
            markup=True,
            halign='center', 
            valign='middle',
            font_size=sp(18),
            size_hint_y=None,
            height=75
        )
        lbl_titulo.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_titulo)
        
        self.inputs = {}
        
        campos_paso6 = [
            ("bm_total_consultas", "Base de Misiones - Total Consultas:"),
            ("urb_total_consultas", "Urbanismos - Total Consultas:"),
            ("bm_total_terreno", "Base de Misiones - De ello: Total de terreno:"),
            ("urb_total_terreno", "Urbanismos - De ello: Total de terreno:")
        ]

        for clave, texto in campos_paso6:
            lbl = Label(text=texto, halign='left', size_hint_y=None, height=55, font_size=sp(15))
            lbl.bind(size=lambda s, w: setattr(s, 'text_size', w))
            layout.add_widget(lbl)
            
            valor_guardado = str(CACHE_MANUAL.get(clave, "0"))
            txt = TextInput(
                text=valor_guardado, 
                multiline=False, 
                input_filter='int', 
                size_hint_x=1.0, 
                size_hint_y=None, 
                height=65, 
                font_size=sp(16), 
                write_tab=False
            )
            txt.bind(text=lambda instance, val, k=clave: CACHE_MANUAL.update({k: val}))
            layout.add_widget(txt)
            self.inputs[clave] = txt
            
        layout.add_widget(Widget(size_hint_y=None, height=15))
        
        # Botonera de navegación (Anterior hacia el Paso 5 / Siguiente hacia el Paso 7)
        btn_layout = BoxLayout(orientation='horizontal', spacing=10, size_hint_y=None, height=65)
        
        btn_ant = Button(
            text="⬅️ Anterior",
            size_hint_y=None,
            height=65,
            background_color=(0.5, 0.5, 0.5, 1),
            font_size=sp(16)
        )
        btn_ant.bind(on_press=lambda x: setattr(self.manager, 'current', 'exportar_paso5'))
        
        btn_sig = Button(
            text="Siguiente ➡️",
            size_hint_y=None,
            height=65,
            background_color=(0.2, 0.6, 0.8, 1),
            font_size=sp(16)
        )
        btn_sig.bind(on_press=lambda x: setattr(self.manager, 'current', 'exportar_paso7'))
        
        btn_layout.add_widget(btn_ant)
        btn_layout.add_widget(btn_sig)
        layout.add_widget(btn_layout)
        
        layout.add_widget(Widget(size_hint_y=None, height=150))
        scroll.add_widget(layout)
        self.add_widget(scroll)


# --- EXPORTAR / CONSOLIDADO - PASO 7 (Parte 2 y Cierre con Generación de Excel) ---
class PantallaExportarPaso7(Screen):
    def __init__(self, **kwargs):
        super(PantallaExportarPaso7, self).__init__(**kwargs)
        
        scroll = ScrollView(scroll_timeout=250, scroll_distance=15)
        layout = BoxLayout(orientation='vertical', padding=30, spacing=20, size_hint_y=None)
        layout.bind(minimum_height=layout.setter('height'))
        
        layout.add_widget(Widget(size_hint_y=None, height=10))
        
        lbl_titulo = Label(
            text="[b]Otros Centros y Generación (Paso Final)[/b]",
            markup=True,
            halign='center', 
            valign='middle',
            font_size=sp(18),
            size_hint_y=None,
            height=75
        )
        lbl_titulo.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_titulo)
        
        self.inputs = {}
        
        campos_paso7 = [
            ("privados_libertad", "Privados de Libertad (en prisión):"),
            ("colegios_escuelas", "Colegios (Escuelas):")
        ]

        for clave, texto in campos_paso7:
            lbl = Label(text=texto, halign='left', size_hint_y=None, height=55, font_size=sp(15))
            lbl.bind(size=lambda s, w: setattr(s, 'text_size', w))
            layout.add_widget(lbl)
            
            valor_guardado = str(CACHE_MANUAL.get(clave, "0"))
            txt = TextInput(
                text=valor_guardado, 
                multiline=False, 
                input_filter='int', 
                size_hint_x=1.0, 
                size_hint_y=None, 
                height=65, 
                font_size=sp(16), 
                write_tab=False
            )
            txt.bind(text=lambda instance, val, k=clave: CACHE_MANUAL.update({k: val}))
            layout.add_widget(txt)
            self.inputs[clave] = txt
            
        layout.add_widget(Widget(size_hint_y=None, height=15))

        
                
        # Botón de acción principal (Generar Excel Completo)
        btn_generar = Button(
            text="💾 Generar Excel Completo",
            size_hint_y=None,
            height=65,
            background_color=(0.1, 0.6, 0.3, 1),
            font_size=sp(17)
        )
        btn_generar.bind(on_press=self.procesar_y_exportar)
        layout.add_widget(btn_generar)
        
        # Botón de retorno al Paso 6
        btn_ant = Button(
            text="⬅️ Anterior",
            size_hint_y=None,
            height=65,
            background_color=(0.5, 0.5, 0.5, 1),
            font_size=sp(17)
        )
        btn_ant.bind(on_press=lambda x: setattr(self.manager, 'current', 'exportar_paso6'))
        layout.add_widget(btn_ant)
        
        layout.add_widget(Widget(size_hint_y=None, height=150))
        scroll.add_widget(layout)
        self.add_widget(scroll)

    def procesar_y_exportar(self, instance):
        try:
            import os
            import traceback
            from collections import defaultdict

            print("\n[⏳] Iniciando proceso de exportación completa desde la Pantalla 7...")

            # 1. Capturamos la semana del Paso 1 de forma segura
            # 1. Capturamos la semana del Paso 1 de forma robusta y segura
            semana_actual_num = 1
            if self.manager.has_screen('exportar_paso1'):
                paso1 = self.manager.get_screen('exportar_paso1')
                if hasattr(paso1, 'input_semana_reporte'):
                    texto_semana = paso1.input_semana_reporte.text.strip()
                    if texto_semana.isdigit():
                        semana_actual_num = int(texto_semana)

               
            print(f"[✔] Semana a procesar: {semana_actual_num}")

            # 2. RECOLECCIÓN UNIFICADA DE TODAS LAS PANTALLAS HACIA CACHE_MANUAL
            datos_conceptos = {}
            pasos_a_revisar = ['exportar_paso2', 'exportar_paso4', 'exportar_paso5', 'exportar_paso6', 'exportar_paso7']
            
            for nom_paso in pasos_a_revisar:
                if self.manager.has_screen(nom_paso):
                    pant = self.manager.get_screen(nom_paso)
                    if hasattr(pant, 'inputs'):
                        for k, widget_input in pant.inputs.items():
                            v_txt = widget_input.text.strip() if hasattr(widget_input, 'text') else "0"
                            CACHE_MANUAL[k] = v_txt

            # Mapeos forzados de seguridad para llaves especiales (Remitidos)
            if 'fila_40' not in CACHE_MANUAL and 'cdi' in CACHE_MANUAL:
                CACHE_MANUAL['fila_40'] = CACHE_MANUAL.get('cdi', '0')
            if 'fila_41' not in CACHE_MANUAL and 'estatales' in CACHE_MANUAL:
                CACHE_MANUAL['fila_41'] = CACHE_MANUAL.get('estatales', '0')
            if 'fila_42' not in CACHE_MANUAL and 'privadas' in CACHE_MANUAL:
                CACHE_MANUAL['fila_42'] = CACHE_MANUAL.get('privadas', '0')

            # 🟢 MAPEO COMPLETO A DATOS_CONCEPTOS (Pantallas 4, 5, 6 y 7)
            datos_conceptos["charlas"] = int(str(CACHE_MANUAL.get("fila_5", "0")).strip() or 0)
            datos_conceptos["partos"] = int(str(CACHE_MANUAL.get("fila_6", "0")).strip() or 0)
            datos_conceptos["vidas_salvadas"] = int(str(CACHE_MANUAL.get("fila_7", "0")).strip() or 0)
            datos_conceptos["nacidos_vivos"] = int(str(CACHE_MANUAL.get("fila_8", "0")).strip() or 0)
            datos_conceptos["encamados"] = int(str(CACHE_MANUAL.get("fila_9", "0")).strip() or 0)
            
            # Conceptos de Pantalla 5
            datos_conceptos["ingresos_hogar"] = int(str(CACHE_MANUAL.get("fila_10", "0")).strip() or 0)
            datos_conceptos["altas_hogar"] = int(str(CACHE_MANUAL.get("fila_11", "0")).strip() or 0)
            datos_conceptos["examen_mama"] = int(str(CACHE_MANUAL.get("fila_12", "0")).strip() or 0)
            datos_conceptos["diu"] = int(str(CACHE_MANUAL.get("fila_13", "0")).strip() or 0)
            datos_conceptos["anticonceptivos"] = int(str(CACHE_MANUAL.get("fila_14", "0")).strip() or 0)

            # Bases de Misiones y Urbanismos (Pantalla 6)
            datos_conceptos["bm_total_consultas"] = int(str(CACHE_MANUAL.get("bm_total_consultas", "0")).strip() or 0)
            datos_conceptos["urb_total_consultas"] = int(str(CACHE_MANUAL.get("urb_total_consultas", "0")).strip() or 0)
            datos_conceptos["bm_total_terreno"] = int(str(CACHE_MANUAL.get("bm_total_terreno", "0")).strip() or 0)
            datos_conceptos["urb_total_terreno"] = int(str(CACHE_MANUAL.get("urb_total_terreno", "0")).strip() or 0)

            # Otros Centros (Pantalla 7)
            datos_conceptos["centros_especializados"] = int(str(CACHE_MANUAL.get("centros_especializados", "0")).strip() or 0)
            datos_conceptos["privados_libertad"] = int(str(CACHE_MANUAL.get("privados_libertad", "0")).strip() or 0)
            datos_conceptos["colegios_escuelas"] = int(str(CACHE_MANUAL.get("colegios_escuelas", "0")).strip() or 0)

            print(f"[✔] Datos de todas las pantallas recolectados correctamente.")

            # 3. Configuración del médico y consultorio
            # 3. Configuración del médico, cédula y consultorio
            medico_activo = "DRA. ARELIS PIÑA"
            consultorio_activo = "CMP LOS ROSALES"
            estado_activo = "Carabobo"
            municipio_activo = "Puerto Cabello"
            asic_activo = "PORTUARIO"

            
            if 'cargar_config_db' in globals() or 'cargar_config_db' in locals():
                config = cargar_config_db()
                if config:
                    # 🟢 Usamos índices seguros para los 6 valores de la BD
                    medico_activo = config[0] if len(config) > 0 and config[0] else "DRA. ARELIS PIÑA"
                    consultorio_activo = config[1] if len(config) > 1 and config[1] else "CMP LOS ROSALES"
                    estado_activo = config[2] if len(config) > 2 and config[2] else "Carabobo"
                    municipio_activo = config[3] if len(config) > 3 and config[3] else "Puerto Cabello"
                    asic_activo = config[4] if len(config) > 4 and config[4] else "PORTUARIO"
                    
                    

                 
            # 4. Consulta a la Base de Datos SQLite
            pacientes_data = []
            if 'DB_PATH' in globals() or 'DB_PATH' in locals():
                try:
                    conexion = sqlite3.connect(DB_PATH)
                    conexion.row_factory = sqlite3.Row
                    cursor = conexion.cursor()
                    cursor.execute(
                        "SELECT * FROM pacientes WHERE semana_num = ? OR semana_num = ?", 
                        (str(semana_actual_num), int(semana_actual_num))
                    )
                    filas = cursor.fetchall()
                    for fila in filas:
                        pacientes_data.append(dict(fila))
                    conexion.close()
                    print(f"[✔] Se cargaron {len(pacientes_data)} pacientes de la semana {semana_actual_num} desde SQLite.")
                except Exception as db_err:
                    print(f"[!] Error al consultar SQLite: {db_err}")

            if not pacientes_data:
                print(f"[!] Advertencia: No hay pacientes registrados para la semana {semana_actual_num}.")
                return

            # 5. Creación de directorios y archivos Excel
            carpeta_principal = "Reportes_Epidemiologicos"
            carpeta_semana = os.path.join(carpeta_principal, f"Semana_{semana_actual_num}")
            os.makedirs(carpeta_semana, exist_ok=True)

            pacientes_por_fecha = defaultdict(list)
            for p in pacientes_data:
                fec = str(p.get("fecha_consulta", "")).strip()
                if not fec:
                    fec = datetime.date.today().strftime("%d/%m/%Y")
                pacientes_por_fecha[fec].append(p)

            # A. Generar Hojas Diarias (RP-01)
            for fecha_c, lista_p_dia in pacientes_por_fecha.items():
                fecha_nombre_archivo = fecha_c.replace("/", "-")
                ruta_rp01_dia = os.path.join(carpeta_semana, f"RP01_Diario_{fecha_nombre_archivo}.xlsx")
                
                if 'generar_excel_rp01' in globals() or 'generar_excel_rp01' in locals():
                    generar_excel_rp01(
                        lista_pacientes=lista_p_dia,
                        fecha_consulta=fecha_c,
                        nombre_archivo=ruta_rp01_dia,
                        nombre_medico=medico_activo,
                        nombre_consultorio=consultorio_activo
                    )

            # B. Generar Consolidado Nuevo (RP-66 con sus 2 páginas de EPI-12 integradas)
            ruta_consolidado = os.path.join(carpeta_semana, f"Consolidado_Semana_{semana_actual_num}.xlsx")
            
            # Verificamos si existe la función nueva o la vieja para evitar caídas
            if 'generar_consolidado_rp66_nuevo' in globals() or 'generar_consolidado_rp66_nuevo' in locals():
                generar_consolidado_rp66_nuevo(
                    pacientes_semana=pacientes_data,
                    semana_num=semana_actual_num,
                    asic=asic_activo,
                    municipio=municipio_activo,
                    estado=estado_activo,
                    nombre_archivo=ruta_consolidado,
                    datos_conceptos=datos_conceptos,
                    datos_actividades=datos_conceptos,
                    datos_manuales=CACHE_MANUAL,
                    nombre_medico=medico_activo,
                    nombre_consultorio=consultorio_activo
                )
            elif 'generar_consolidado_semanal' in globals() or 'generar_consolidado_semanal' in locals():
                generar_consolidado_semanal(
                    pacientes_semana=pacientes_data,
                    semana_num=semana_actual_num,
                    asic="PORTUARIO",
                    estado="Carabobo",
                    nombre_archivo=ruta_consolidado,
                    datos_conceptos=datos_conceptos,
                    datos_actividades=datos_conceptos,
                    datos_manuales=CACHE_MANUAL,
                    nombre_medico=medico_activo,
                    nombre_consultorio=consultorio_activo
                )
              
            print(f"[✔] ¡Proceso completado con éxito! Archivos guardados nítidos en: {carpeta_semana}/")
            self.manager.current = 'menu'
            
        except Exception as e:
            print(f"\n[❌ ERROR CRÍTICO EN EXPORTACIÓN]: {e}")
            traceback.print_exc()

      
        
       
class DeslizableView(TextInput):
    """TextInput modificado para evitar que abra el teclado virtual al intentar hacer scroll."""
    def on_touch_down(self, touch):
        # Si el toque no es dentro de este campo, ignorar
        if not self.collide_point(*touch.pos):
            return super(DeslizableView, self).on_touch_down(touch)
        
        # Guardamos la posición inicial del toque para medir si hubo arrastre (scroll)
        touch.ud['initial_pos'] = touch.pos
        return super(DeslizableView, self).on_touch_down(touch)

    def on_touch_move(self, touch):
        # Si el usuario mueve el dedo (está deslizando la pantalla), 
        # quitamos temporalmente el foco para que el teclado no brinque a molestar.
        if 'initial_pos' in touch.ud:
            dx = abs(touch.pos[0] - touch.ud['initial_pos'][0])
            dy = abs(touch.pos[1] - touch.ud['initial_pos'][1])
            if dx > 10 or dy > 10:  # Si se movió más de 10 píxeles, es un desplazamiento
                if self.focus:
                    self.focus = False
        return super(DeslizableView, self).on_touch_move(touch)

        
# --- PANTALLA DE REGISTRO DE PACIENTES ---
# --- PANTALLA DE REGISTRO DE PACIENTES ---


# ==========================================
# PASO 1: Control, Fecha y Identificación Básica (4 campos)
# ==========================================
class PantallaRegistroPaso1(Screen):
    def __init__(self, **kwargs):
        super(PantallaRegistroPaso1, self).__init__(**kwargs)
        
        # 1. Usamos ScrollView para que se adapte al teclado de Android sin recortar nada
        scroll = ScrollView()
        
        layout = BoxLayout(
            orientation='vertical', 
            padding=30, 
            spacing=20,
            size_hint_y=None
        )
        layout.bind(minimum_height=layout.setter('height'))
        
        layout.add_widget(Widget(size_hint_y=None, height=10))
        
        # 2. Título ajustado a REGISTRO PACIENTE (con la misma altura y diseño de la configuración)
        lbl_titulo = Label(
            text="REGISTRO PACIENTE\nControl e Identificación", 
            halign='center', 
            valign='middle',
            font_size=sp(20),
            size_hint_y=None,
            height=95
        )
        lbl_titulo.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_titulo)
        
        layout.add_widget(Widget(size_hint_y=None, height=10))
        
        # --- ETIQUETA DE ALERTA (SALVAVIDAS) ---
        self.lbl_error = Label(
            text="", 
            halign='center', 
            valign='middle',
            font_size=sp(14),
            color=(1, 0.3, 0.3, 1), # Rojo suave
            size_hint_y=None,
            height=55
        )
        self.lbl_error.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(self.lbl_error)
        # --------------------------------------
        layout.add_widget(Widget(size_hint_y=None, height=10))
         
        # 3. Semana (Etiqueta height=55, TextInput height=65 para que quepa todo holgado)
        lbl_sem = Label(text="N° de Semana (1-55):", halign='left', size_hint_y=None, height=55, font_size=sp(15))
        lbl_sem.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_sem)
        
        self.input_semana = DeslizableView(
            text='', multiline=False, size_hint_x=1.0, size_hint_y=None, height=65, font_size=sp(16), write_tab=False, padding_y=[8, 8]
        )
        layout.add_widget(self.input_semana)
        
        # 4. Fecha de Consulta
        fecha_hoy = datetime.date.today().strftime("%d/%m/%Y")
        lbl_fec = Label(text=f"Fecha de Consulta (DD/MM/AAAA):", halign='left', size_hint_y=None, height=55, font_size=sp(15))
        lbl_fec.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_fec)
        
        self.input_fecha = DeslizableView(
            text=fecha_hoy, multiline=False, size_hint_x=1.0, size_hint_y=None, height=65, font_size=sp(16), write_tab=False, padding_y=[8, 8]
        )
        layout.add_widget(self.input_fecha)
        
        # 5. Nombre y Apellido
        lbl_nom = Label(text="Nombre y Apellido:", halign='left', size_hint_y=None, height=55, font_size=sp(15))
        lbl_nom.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_nom)
        
        self.input_nombre = DeslizableView(
            text='', multiline=False, size_hint_x=1.0, size_hint_y=None, height=65, font_size=sp(16), write_tab=False, padding_y=[8, 8]
        )
        layout.add_widget(self.input_nombre)
        
        # 6. Cédula
        lbl_ced = Label(text="Cédula (Ej. V-12345678):", halign='left', size_hint_y=None, height=55, font_size=sp(15))
        lbl_ced.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_ced)
        
        self.input_cedula = DeslizableView(
            text='', multiline=False, size_hint_x=1.0, size_hint_y=None, height=65, font_size=sp(16), write_tab=False, padding_y=[8, 8]
        )
        layout.add_widget(self.input_cedula)
        
        layout.add_widget(Widget(size_hint_y=None, height=15))
        
        # Botones de Navegación (con altura de 65 al igual que los botones principales)
        btn_layout = BoxLayout(orientation='horizontal', spacing=10, size_hint_y=None, height=65)
        
        btn_volver = Button(
            text="Menú Principal", 
            background_color=(0.8, 0.3, 0.3, 1), 
            font_size=sp(16),
            size_hint_y=None,
            height=65
        )
        btn_volver.bind(on_press=lambda x: setattr(self.manager, 'current', 'menu'))
        
        btn_sig = Button(
            text="Siguiente ➡️", 
            background_color=(0.2, 0.5, 0.8, 1), 
            font_size=sp(16),
            size_hint_y=None,
            height=65
        )
        btn_sig.bind(on_press=self.ir_paso_2)
        
        btn_layout.add_widget(btn_volver)
        btn_layout.add_widget(btn_sig)
        layout.add_widget(btn_layout)
        
        layout.add_widget(Widget(size_hint_y=None, height=150))
        
        scroll.add_widget(layout)
        self.add_widget(scroll)

    def ir_paso_2(self, instance):
        cedula = self.input_cedula.text.strip().upper()
        fecha = self.input_fecha.text.strip()
        
        # Validamos que no estén vacíos (mensaje corto)
        if not cedula or not fecha:
            self.lbl_error.text = "⚠️ Ingrese cédula y fecha."
            return

        # --- VALIDACIÓN CON LA BASE DE DATOS OFICIAL (DB_PATH) ---
        try:
            conexion = sqlite3.connect(DB_PATH) # <--- ¡Aquí está la clave, mi pana!
            cursor = conexion.cursor()
            cursor.execute(
                "SELECT COUNT(*) FROM pacientes WHERE cedula = ? AND fecha_consulta = ?",
                (cedula, fecha)
            )
            existe = cursor.fetchone()[0]
            conexion.close()
        except Exception as e:
            existe = 0

        if existe > 0:
            # ¡Freno en seco con texto corto para que no se pique en horizontal!
            self.lbl_error.text = "⚠️ Paciente ya registrado hoy."
            return 

        # Si todo está limpio, borramos el error, guardamos y avanzamos al paso 2
        self.lbl_error.text = ""
        app = App.get_running_app()
        app.datos_registro = {
            'semana': self.input_semana.text.strip(),
            'fecha': fecha,
            'nombre': self.input_nombre.text.strip().upper(),
            'cedula': cedula
        }
        self.manager.current = 'registro_paso2'

        
      
         

# ==========================================
# PASO 2: Demografía y Ubicación (5 campos)
# ==========================================
class PantallaRegistroPaso2(Screen):
    def __init__(self, **kwargs):
        super(PantallaRegistroPaso2, self).__init__(**kwargs)
        
        scroll = ScrollView()
        
        layout = BoxLayout(
            orientation='vertical', 
            padding=30, 
            spacing=20,
            size_hint_y=None
        )
        layout.bind(minimum_height=layout.setter('height'))
        
        layout.add_widget(Widget(size_hint_y=None, height=10))
        
        # Título
        lbl_titulo = Label(
            text="REGISTRO PACIENTE\nDemografía y Ubicación", 
            halign='center', 
            valign='middle',
            font_size=sp(20),
            size_hint_y=None,
            height=95
        )
        lbl_titulo.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_titulo)
        
        layout.add_widget(Widget(size_hint_y=None, height=10))
        
        # 1. Edad
        lbl_eda = Label(text="Edad (ej. '5' años o '2m' meses):", halign='left', size_hint_y=None, height=55, font_size=sp(15))
        lbl_eda.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_eda)
        
        self.input_edad = DeslizableView(
            text='', multiline=False, size_hint_x=1.0, size_hint_y=None, height=65, font_size=sp(16), write_tab=False, padding_y=[8, 8]
        )
        layout.add_widget(self.input_edad)
        
        # 2. Sexo
        lbl_sex = Label(text="Sexo (H = Hombre / M = Mujer):", halign='left', size_hint_y=None, height=55, font_size=sp(15))
        lbl_sex.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_sex)
        
        self.input_sexo = DeslizableView(
            text='H', multiline=False, size_hint_x=1.0, size_hint_y=None, height=65, font_size=sp(16), write_tab=False, padding_y=[8, 8]
        )
        layout.add_widget(self.input_sexo)
        
        # 3. Ámbito
        lbl_amb = Label(text="Ámbito (C = Consultorio / T = Terreno):", halign='left', size_hint_y=None, height=55, font_size=sp(15))
        lbl_amb.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_amb)
        
        self.input_ambito = DeslizableView(
            text='C', multiline=False, size_hint_x=1.0, size_hint_y=None, height=65, font_size=sp(16), write_tab=False, padding_y=[8, 8]
        )
        layout.add_widget(self.input_ambito)
        
        # 4. Dirección
        lbl_dir = Label(text="Dirección:", halign='left', size_hint_y=None, height=55, font_size=sp(15))
        lbl_dir.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_dir)
        
        self.input_direccion = DeslizableView(
            text='', multiline=False, size_hint_x=1.0, size_hint_y=None, height=65, font_size=sp(16), write_tab=False, padding_y=[8, 8]
        )
        layout.add_widget(self.input_direccion)
        
        # 🟢 BOTÓN PARA ABRIR EL MENÚ DE SELECCIÓN DE DIRECCIÓN
        btn_elegir_dir = Button(
            text="📍 Seleccionar Calle San José / Manual", 
            size_hint_y=None, 
            height=50, 
            font_size=sp(14),
            background_normal=' ',
            background_color=(0.8, 0.3, 0.3, 1),
            color=(1, 1, 1, 1),
            bold=True
        )
        btn_elegir_dir.bind(on_press=lambda x: setattr(self.manager, 'current', 'menu_direccion'))
        layout.add_widget(btn_elegir_dir)
        
        # 5. Nacionalidad
        lbl_nac = Label(text="Nacionalidad (V / C / B / O):", halign='left', size_hint_y=None, height=55, font_size=sp(15))
        lbl_nac.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_nac)
        
        self.input_nacionalidad = DeslizableView(
            text='V', multiline=False, size_hint_x=1.0, size_hint_y=None, height=65, font_size=sp(16), write_tab=False, padding_y=[8, 8]
        )
        layout.add_widget(self.input_nacionalidad)
        
        layout.add_widget(Widget(size_hint_y=None, height=10))
        
        # --- ETIQUETA DE ERROR EN ROJO (JUSTO ENCIMA DE LOS BOTONES) ---
        self.lbl_alerta = Label(
            text="", 
            color=(1, 0, 0, 1), # Rojo puro chillón
            bold=True,
            font_size=sp(18),
            halign='center',
            valign='middle',
            size_hint_y=None,
            height=40
        )
        self.lbl_alerta.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(self.lbl_alerta)
        
        # Botones de Navegación
        btn_layout = BoxLayout(orientation='horizontal', spacing=10, size_hint_y=None, height=65)
        
        btn_ant = Button(
            text="⬅️ Anterior", 
            background_color=(0.5, 0.5, 0.5, 1), 
            font_size=sp(16),
            size_hint_y=None,
            height=65
        )
        btn_ant.bind(on_press=self.ir_paso_anterior)
        
        btn_sig = Button(
            text="Siguiente ➡️", 
            background_color=(0.2, 0.5, 0.8, 1), 
            font_size=sp(16),
            size_hint_y=None,
            height=65
        )
        btn_sig.bind(on_press=self.ir_paso_3)
        
        btn_layout.add_widget(btn_ant)
        btn_layout.add_widget(btn_sig)
        layout.add_widget(btn_layout)
        
        layout.add_widget(Widget(size_hint_y=None, height=150))
        
        scroll.add_widget(layout)
        self.add_widget(scroll)

    def on_enter(self):
        self.lbl_alerta.text = ""
        app = App.get_running_app()
        # Si se seleccionó una dirección en la pantalla de menú, la inyectamos aquí
        if hasattr(app, 'datos_registro') and 'direccion' in app.datos_registro:
            self.input_direccion.text = app.datos_registro['direccion']

    def ir_paso_anterior(self, instance):
        self.lbl_alerta.text = ""
        self.manager.current = 'registro_paso1'

    def ir_paso_3(self, instance):
        app = App.get_running_app()
        
        edad = self.input_edad.text.strip()
        sexo = self.input_sexo.text.strip().upper()
        ambito = self.input_ambito.text.strip().upper()
        direccion = self.input_direccion.text.strip().upper()
        nacionalidad = self.input_nacionalidad.text.strip().upper()
        
        # FILTRO ESTRICTO: Si falta algo, prende el cartelón de ERROR en rojo
        if not edad or not sexo or not ambito or not direccion or not nacionalidad:
            self.lbl_alerta.text = "⚠️ ERROR"
            return
        
        self.lbl_alerta.text = ""
        
        app.datos_registro.update({
            'edad': edad,
            'sexo': sexo,
            'ambito': ambito,
            'direccion': direccion,
            'nacionalidad': nacionalidad
        })
        
        self.manager.current = 'registro_paso3'
       
       

# ==========================================
# PASO 3: Condiciones Clínicas y Especiales (5 campos / botones)
# ==========================================
class PantallaRegistroPaso3(Screen):
    def __init__(self, **kwargs):
        super(PantallaRegistroPaso3, self).__init__(**kwargs)
        
        # 1. Usamos ScrollView para evitar recortes y proteger contra el teclado
        scroll = ScrollView()
        
        layout = BoxLayout(
            orientation='vertical', 
            padding=30, 
            spacing=20,
            size_hint_y=None
        )
        layout.bind(minimum_height=layout.setter('height'))
        
        layout.add_widget(Widget(size_hint_y=None, height=10))
        
        # 2. Título unificado idéntico a los pasos anteriores
        lbl_titulo = Label(
            text="REGISTRO PACIENTE\nCondiciones Clínicas y Sociodemográficas", 
            halign='center', 
            valign='middle',
            font_size=sp(20),
            size_hint_y=None,
            height=95
        )
        lbl_titulo.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_titulo)
        
        layout.add_widget(Widget(size_hint_y=None, height=10))
        
        # 3. Embarazo
        lbl_emb = Label(text="¿Está Embarazada? (Si / No):", halign='left', size_hint_y=None, height=55, font_size=sp(15))
        lbl_emb.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_emb)
        
        self.input_embarazo = DeslizableView(
            text='No', multiline=False, size_hint_x=1.0, size_hint_y=None, height=65, font_size=sp(16), write_tab=False, padding_y=[8, 8]
        )
        layout.add_widget(self.input_embarazo)
        
        # 4. Botón Trimestre
        self.btn_trimestre = Button(
            text="🔍 Seleccionar Trimestre ", 
            size_hint_y=None, 
            height=65, 
            font_size=sp(16),
            background_normal=' ',                 
            background_color=(0.2, 0.5, 0.8, 1),  
            color=(0.6, 0.05, 0.05, 1),           
            bold=True
        )                 
        self.btn_trimestre.bind(on_press=lambda x: setattr(self.manager, 'current', 'menu_trimestre'))
        layout.add_widget(self.btn_trimestre)
        
        # 5. Puerperio
        lbl_puer = Label(text="¿Está en Puerperio? (Si / No):", halign='left', size_hint_y=None, height=55, font_size=sp(15))
        lbl_puer.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_puer)
        
        self.input_puerperio = DeslizableView(
            text='No', multiline=False, size_hint_x=1.0, size_hint_y=None, height=65, font_size=sp(16), write_tab=False, padding_y=[8, 8]
        )
        layout.add_widget(self.input_puerperio)
        
        # 6. Riesgo Prenatal
        lbl_rie = Label(text="¿Tiene Riesgo Prenatal? (Si / No):", halign='left', size_hint_y=None, height=55, font_size=sp(15))
        lbl_rie.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_rie)
        
        self.input_riesgo = DeslizableView(
            text='No', multiline=False, size_hint_x=1.0, size_hint_y=None, height=65, font_size=sp(16), write_tab=False, padding_y=[8, 8]
        )
        layout.add_widget(self.input_riesgo)

        # 7. Indígena
        lbl_ind = Label(text="¿Es Indígena? (Si / No):", halign='left', size_hint_y=None, height=55, font_size=sp(15))
        lbl_ind.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_ind)
        
        self.input_indigena = DeslizableView(
            text='No', multiline=False, size_hint_x=1.0, size_hint_y=None, height=65, font_size=sp(16), write_tab=False, padding_y=[8, 8]
        )
        layout.add_widget(self.input_indigena)
        
        # 8. Tipo de Consulta
        lbl_tip = Label(
            text="Tipo de Consulta (1=Primera / 2=Reconsulta):", 
            halign='left', 
            valign='middle',
            size_hint_y=None, 
            height=75,
            font_size=sp(15)
        )
        lbl_tip.bind(width=lambda s, w: setattr(s, 'text_size', (w, None)))
        layout.add_widget(lbl_tip)
        
        self.input_tipo_consulta = DeslizableView(
            text='1', multiline=False, size_hint_x=1.0, size_hint_y=None, height=65, font_size=sp(16), write_tab=False, padding_y=[8, 8]
        )
        layout.add_widget(self.input_tipo_consulta)
        
        layout.add_widget(Widget(size_hint_y=None, height=10))
        
        # --- ETIQUETA DE ERROR EN ROJO (JUSTO ENCIMA DE LOS BOTONES) ---
        self.lbl_alerta = Label(
            text="", 
            color=(1, 0, 0, 1), # Rojo puro chillón
            bold=True,
            font_size=sp(18),
            halign='center',
            valign='middle',
            size_hint_y=None,
            height=40
        )
        self.lbl_alerta.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(self.lbl_alerta)
        
        # Botones de Navegación
        btn_layout = BoxLayout(orientation='horizontal', spacing=10, size_hint_y=None, height=65)
        
        btn_ant = Button(
            text="⬅️ Anterior", 
            background_color=(0.5, 0.5, 0.5, 1), 
            font_size=sp(16),
            size_hint_y=None,
            height=65
        )
        btn_ant.bind(on_press=self.ir_paso_anterior)
        
        btn_sig = Button(
            text="Siguiente ➡️", 
            background_color=(0.2, 0.5, 0.8, 1), 
            font_size=sp(16),
            size_hint_y=None,
            height=65
        )
        btn_sig.bind(on_press=self.ir_paso_4)
        
        btn_layout.add_widget(btn_ant)
        btn_layout.add_widget(btn_sig)
        layout.add_widget(btn_layout)
        
        layout.add_widget(Widget(size_hint_y=None, height=150))
        
        scroll.add_widget(layout)
        self.add_widget(scroll)

    def on_enter(self):
        self.lbl_alerta.text = ""

    def ir_paso_anterior(self, instance):
        self.lbl_alerta.text = ""
        self.manager.current = 'registro_paso2'

    def ir_paso_4(self, instance):
        app = App.get_running_app()
        
        embarazo = self.input_embarazo.text.strip().capitalize()
        puerperio = self.input_puerperio.text.strip().capitalize()
        riesgo = self.input_riesgo.text.strip().capitalize()
        indigena = self.input_indigena.text.strip().capitalize()
        tipo_consulta = self.input_tipo_consulta.text.strip().upper()
        
        # 1. Validación de campos vacíos
        if not embarazo or not puerperio or not riesgo or not indigena or not tipo_consulta:
            self.lbl_alerta.text = "⚠️ ERROR"
            return
            
        # 2. Recuperamos Sexo y Edad del Paso 2 guardados en app.datos_registro
        sexo_paciente = app.datos_registro.get('sexo', '').upper()
        edad_raw = app.datos_registro.get('edad', '').upper()
        
        # Convertimos la edad a años numéricos de forma segura
        anios = 0
        if 'M' not in edad_raw:  # Si no son meses, extraemos los dígitos
            num_str = ''.join(filter(str.isdigit, edad_raw))
            anios = int(num_str) if num_str else 0
        
        # Verificamos si marcaron "Si" (o "Sí")
        esta_embarazada = embarazo.startswith('S')
        esta_puerpera = puerperio.startswith('S')
        
        # =========================================================
        # 🛑 AQUÍ PEGAS LA REGLA EXCLUSIVA (JUSTO AQUÍ)
        # =========================================================
        if esta_embarazada and esta_puerpera:
            self.lbl_alerta.text = "⚠️ ERROR: No puede estar Embarazada y Puérpera a la vez"
            print("[!] Alerta clínica: Conflicto obstétrico (Embarazada + Puérpera).")
            return
        # =====================
        
        # 3. FILTRO CLÍNICO / EPIDEMIOLÓGICO:
        # A) Si es Hombre y marcan Embarazo/Puerperio en 'Si' -> ERROR
        if sexo_paciente == 'H' and (esta_embarazada or esta_puerpera):
            self.lbl_alerta.text = "⚠️ ERROR"
            print("[!] Alerta clínica: Paciente masculino no puede registrar embarazo/puerperio.")
            return
            
        # B) Si es Mujer pero tiene 9 años o menos y marcan Embarazo/Puerperio en 'Si' -> ERROR
        if sexo_paciente == 'M' and anios <= 9 and (esta_embarazada or esta_puerpera):
            self.lbl_alerta.text = "⚠️ ERROR"
            print("[!] Alerta clínica: Paciente femenino menor o igual a 9 años no aplica para embarazo/puerperio.")
            return

        # C) FRENO OBLIGATORIO DE TRIMESTRE
        trimestre_sel = getattr(app, "trimestre_seleccionado", "N/A").strip()
        if esta_embarazada and (trimestre_sel == "N/A" or not trimestre_sel):
            self.lbl_alerta.text = "⚠️ ERROR: Debe seleccionar trimestre"
            return

        # Si supera todos los filtros, limpiamos la alerta y avanzamos
        self.lbl_alerta.text = ""
        
        app.datos_registro.update({
            'embarazo': embarazo,
            'puerperio': puerperio,
            'riesgo_prenatal': riesgo,
            'indigena': indigena,
            'tipo_consulta': tipo_consulta,
            'trimestre': trimestre_sel if esta_embarazada else 'N/A'
        })
        
        self.manager.current = 'registro_paso4'

            
        

# ==========================================
# PASO 4: Diagnóstico, Conducta y Guardado en SQLite (4 campos / botones)
# ==========================================
class PantallaRegistroPaso4(Screen):
    def __init__(self, **kwargs):
        super(PantallaRegistroPaso4, self).__init__(**kwargs)
        
        # 1. ScrollView guardado como atributo para controlar el Auto-Scroll
        self.scroll = ScrollView()
        
        layout = BoxLayout(
            orientation='vertical', 
            padding=30, 
            spacing=20,
            size_hint_y=None
        )
        layout.bind(minimum_height=layout.setter('height'))
        
        layout.add_widget(Widget(size_hint_y=None, height=10))
        
        # 2. Título unificado idéntico a los pasos anteriores
        lbl_titulo = Label(
            text="REGISTRO PACIENTE\nDiagnóstico y Cierre", 
            halign='center', 
            valign='middle',
            font_size=sp(20),
            size_hint_y=None,
            height=95
        )
        lbl_titulo.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_titulo)
        
        layout.add_widget(Widget(size_hint_y=None, height=10))
        
        # 3. Tipo de Diagnóstico
        lbl_tdiag = Label(
            text="Tipo Diagnóstico (C=Confirmado / P=Presuntivo):", 
            halign='left', 
            valign='middle',
            size_hint_y=None, 
            height=75, 
            font_size=sp(15)
        )
        lbl_tdiag.bind(width=lambda s, w: setattr(s, 'text_size', (w, None)))
        layout.add_widget(lbl_tdiag)

        self.input_tipo_diagnostico = DeslizableView(
            text='C', multiline=False, size_hint_x=1.0, size_hint_y=None, height=65, font_size=sp(16), write_tab=False, padding_y=[8, 8]
        )
        layout.add_widget(self.input_tipo_diagnostico)
        
        # 4. Enfermedad Crónica Manual
        lbl_cron = Label(text="¿Es Enfermedad Crónica? (Si / No):", halign='left', size_hint_y=None, height=55, font_size=sp(15))
        lbl_cron.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_cron)
        
        self.input_cronica_manual = DeslizableView(
            text='No', multiline=False, size_hint_x=1.0, size_hint_y=None, height=65, font_size=sp(16), write_tab=False, padding_y=[8, 8]
        )
        layout.add_widget(self.input_cronica_manual)
        
        # 5. Diagnóstico Oficial
        lbl_diag = Label(text="Diagnóstico Oficial :", halign='left', size_hint_y=None, height=55, font_size=sp(15))
        lbl_diag.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_diag)
        
        self.input_diagnostico = DeslizableView(
            text='', multiline=False, size_hint_x=1.0, size_hint_y=None, height=65, font_size=sp(16), write_tab=False, padding_y=[8, 8]
        )
        layout.add_widget(self.input_diagnostico)
        
        # 6. Botón de Selección de Diagnóstico
        self.btn_elegir_diag = Button(
            text="🔍 Seleccionar Diagnóstico ", 
            size_hint_y=None, 
            height=65, 
            font_size=sp(15),
            background_normal=' ',
            background_color=(0.2, 0.5, 0.8, 1),
            color=(0.6, 0.05, 0.05, 1),
            bold=True
        )
        self.btn_elegir_diag.bind(on_press=self.abrir_menu_diagnostico)
        layout.add_widget(self.btn_elegir_diag)
        
        # 7. Conducta / Tratamiento
        lbl_cond = Label(text="Conducta :", halign='left', size_hint_y=None, height=55, font_size=sp(15))
        lbl_cond.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_cond)
        
        self.input_conducta = DeslizableView(
            text='', multiline=False, size_hint_x=1.0, size_hint_y=None, height=65, font_size=sp(16), write_tab=False, padding_y=[8, 8]
        )
        layout.add_widget(self.input_conducta)

        # 8. Botón para Selección de Conducta
        self.btn_elegir_conducta = Button(
            text="🔍 Seleccionar Conducta ", 
            size_hint_y=None, 
            height=65, 
            font_size=sp(15),
            background_normal=' ',
            background_color=(0.2, 0.5, 0.8, 1),
            color=(0.6, 0.05, 0.05, 1),         
            bold=True 
        )
        self.btn_elegir_conducta.bind(on_press=self.abrir_menu_conducta)
        layout.add_widget(self.btn_elegir_conducta)
        
        layout.add_widget(Widget(size_hint_y=None, height=10))
        
        # --- ETIQUETA DE ERROR EN ROJO (JUSTO ENCIMA DE LOS BOTONES) ---
        self.lbl_alerta = Label(
            text="", 
            color=(1, 0, 0, 1), # Rojo puro chillón
            bold=True,
            font_size=sp(15),
            halign='center',
            valign='middle',
            size_hint_y=None,
            height=90
        )
        self.lbl_alerta.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(self.lbl_alerta)
        
        # Botones de Navegación / Guardado (height=65)
        btn_layout = BoxLayout(orientation='horizontal', spacing=10, size_hint_y=None, height=65)
        
        btn_ant = Button(
            text="⬅️ Anterior", 
            background_color=(0.5, 0.5, 0.5, 1), 
            font_size=sp(16),
            size_hint_y=None,
            height=65
        )
        btn_ant.bind(on_press=lambda x: setattr(self.manager, 'current', 'registro_paso3'))
        
        btn_guardar = Button(
            text="💾 Guardar en SQLite", 
            background_color=(0.2, 0.7, 0.3, 1), 
            font_size=sp(16),
            size_hint_y=None,
            height=65
        )
        btn_guardar.bind(on_press=self.guardar_paciente)
        
        btn_layout.add_widget(btn_ant)
        btn_layout.add_widget(btn_guardar)
        layout.add_widget(btn_layout)
        
        layout.add_widget(Widget(size_hint_y=None, height=150))
        
        self.scroll.add_widget(layout)
        self.add_widget(self.scroll)

    def abrir_menu_diagnostico(self, instance):
        app = App.get_running_app()
        app.origen_diagnostico = 'registro'
        self.manager.current = 'menu_diag'

    def abrir_menu_conducta(self, instance):
        app = App.get_running_app()
        # Aseguramos que origen_diagnostico sea None para no pisar el diagnóstico manual al volver
        app.origen_diagnostico = None
        self.manager.current = 'menu_conducta'

    def guardar_paciente(self, instance):
        app = App.get_running_app()
        d = app.datos_registro  # Diccionario acumulado de los pasos 1, 2 y 3
        
        semana_num = d.get('semana', '')
        fecha_consulta = d.get('fecha', '')
        nombre = d.get('nombre', '')
        cedula = d.get('cedula', '')
        edad_raw = d.get('edad', '').upper()
        sexo = d.get('sexo', 'H').upper()
        ambito = d.get('ambito', 'C')
        direccion = d.get('direccion', '')
        
        # Extraemos el valor de embarazo, puerperio y puericultura del Paso 3
        # 1. Leemos el diagnóstico y tipo de diagnóstico primero que nada
        tipo_diagnostico = self.input_tipo_diagnostico.text.strip().upper()
        diagnostico_texto = self.input_diagnostico.text.strip()
        diag_upper = diagnostico_texto.upper()

        # 2. BLINDAJE TOTAL: Verificamos si es un caso de riesgo o terreno
        es_con_riesgo = "RIESGO" in diag_upper or "TERRENO" in diag_upper or "RIESGO" in tipo_diagnostico or "TERRENO" in tipo_diagnostico

        # 3. Leemos las variables del Paso 3
        embarazada_p3 = str(d.get('embarazo', 'No')).strip().upper()
        es_si_embarazada = embarazada_p3 in ["S", "SI", "SÍ", "TRUE", "1"]

        puerpera_p3 = str(d.get('puerperio', 'No')).strip().upper()
        es_si_puerpera = puerpera_p3 in ["S", "SI", "SÍ", "TRUE", "1"]

        # 4. Si es de riesgo, matamos la puericultura a la fuerza; si no, la evaluamos normal
        # 🛑 BLINDAJE TOTAL: Si es de riesgo o terreno, matamos la puericultura a juro
        if es_con_riesgo:
            es_puericultura = False
            es_diag_puericultura = False  # <--- ¡Aquí está la clave que faltaba!
        else:
            puericultura_p3 = str(d.get('puericultura', 'No')).strip().upper()
            es_puericultura = puericultura_p3 in ["S", "SI", "SÍ", "TRUE", "1"]

            # Si hay riesgo, vaciamos los patrones para que 'Escolar' o 'Adolescente' no disparen la falsa alarma
            if es_con_riesgo:
                patrones_puericultura = []
            else:
                patrones_puericultura = [
                    "RECIEN NACIDO", "RECIÉN NACIDO", "RN", 
                    "LACTANTE", "TRANSICIONAL", "PREESCOLAR", 
                    
                ]
            
            es_diag_puericultura = any(p in diag_upper for p in patrones_puericultura) or es_puericultura

        # 🛑 Mételo aquí con la misma sangría exacta de las líneas de arriba:
        if es_con_riesgo:
            es_diag_puericultura = False
            es_puericultura = False

        tipo_consulta = str(d.get('tipo_consulta', '1')).strip().upper()
        is_primera = tipo_consulta in ["1", "C", "PRIMERA", "NUEVA"]


        
        # 5. Lectura de edad en años de forma segura
        anios = 0
        if 'M' not in edad_raw:
            num_str = ''.join(filter(str.isdigit, edad_raw))
            anios = int(num_str) if num_str else 0

        # 6. Leemos la conducta escrita en pantalla
        conducta = self.input_conducta.text.strip()
        conducta_clean = conducta.upper().strip()

        
        # 🛑 FILTRO 1: REGLA ESTRICTA PARA NIÑO SANO (Acepta con o sin acento)
        if es_diag_puericultura:
            if is_primera and conducta_clean not in ["CAPTACIÓN", "CAPTACION"]:
                self.lbl_alerta.text = "⚠️ ERROR: Primera consulta requiere conducta CAPTACIÓN"
                self.scroll.scroll_y = 0.35
                print("[!] Freno activado: Niño Sano en primera consulta requiere CAPTACIÓN.")
                return
            elif not is_primera and conducta_clean not in ["PUERICULTURA/CONTROL", "PUERICULTURA / CONTROL", "CONTROL"]:
                self.lbl_alerta.text = "⚠️ ERROR: Reconsulta requiere conducta PUERICULTURA/CONTROL"
                self.scroll.scroll_y = 0.35
                print("[!] Freno activado: Niño Sano en reconsulta requiere PUERICULTURA/CONTROL.")
                return

           
       
       
       
        # 🛑 FILTRO 2: EL GRAN FILTRO DE MORBILIDAD GENERAL (¡EL NUEVO BLOQUEO!)
        patrones_obstetricos = ["EMBARAZ", "GESTAC", "GESTANT", "PUERPER", "PUÉRPER"]
        patrones_cronicos = ["HIPERTENS", "HTA", "DIABET", "DBT", "ASMA", "EPILEP", "CONVULS", "TUMOR", "NEOPLAS", "CANCER", "CÁNCER", "CEREBROVASC", "ACV", "ICTUS", "RENAL", "NEFRO"]

        es_diag_obstetrico = any(p in diag_upper for p in patrones_obstetricos) or es_si_embarazada or es_si_puerpera
        
        # Detección de crónica manual si existe en la interfaz
        txt_cronica = self.input_cronica_manual.text.strip().upper() if hasattr(self, 'input_cronica_manual') else ""
        es_cronica_manual = txt_cronica in ["S", "SI", "SÍ", "TRUE", "1"]
        es_diag_cronico = any(p in diag_upper for p in patrones_cronicos) or es_cronica_manual

        es_programa_especial = es_diag_puericultura or es_diag_obstetrico or es_diag_cronico

        # Si NO es programa especial y le metieron Captación o Control -> ¡LO REBOTA!
        if not es_programa_especial:
            conductas_prohibidas = ["CAPTACIÓN", "CAPTACION", "CONTROL", "PUERICULTURA/CONTROL", "PUERICULTURA / CONTROL"]
            if conducta_clean in conductas_prohibidas:
                self.lbl_alerta.text = "⚠️ ERROR: Morbilidad general requiere Tratamiento, Referencia, etc. (No Captación/Control)"
                self.scroll.scroll_y = 0.35
                print("[!] Freno activado: No se permite Captación/Control en morbilidad general.")
                return

        # 1. VALIDACIÓN DE CAMPOS VACÍOS OBLIGATORIOS
        if not nombre or not cedula or not diagnostico_texto or not tipo_diagnostico or not conducta:
            self.lbl_alerta.text = "⚠️ ERROR: Faltan campos obligatorios"
            self.scroll.scroll_y = 0.35
            print("[!] Error: Faltan campos obligatorios para guardar.")
            return

        # 2. FILTRO CLÍNICO RIGUROSO Y SEPARADO (EMBARAZO VS PUERPERIO)
        raices_embarazo = ["EMBARAZ", "GESTAC", "GESTANT"]
        raices_puerperio = ["PUERPER", "PUÉRPER"]
        raices_otras = ["PARTO", "CESAR", "ABORTO"]
        
        tiene_embarazo = any(r in diag_upper for r in raices_embarazo)
        tiene_puerperio = any(r in diag_upper for r in raices_puerperio)
        tiene_otro_obstetrico = any(r in diag_upper for r in raices_otras)
        
        contiene_obstetrico = tiene_embarazo or tiene_puerperio or tiene_otro_obstetrico
        
        # 🛑 REGLA SUPREMA: Si en el Paso 3 dijo SÍ embarazada, el diagnóstico OBLIGATORIO debe ser de embarazo
        if es_si_embarazada and not tiene_embarazo:
            self.lbl_alerta.text = "⚠️ ERROR: Paciente embarazada requiere diag. de embarazo"
            self.scroll.scroll_y = 0.35
            return

       

        # A-1) Si el diagnóstico es de EMBARAZO pero en Paso 3 NO indicó estar embarazada -> ERROR
        if tiene_embarazo and not es_si_embarazada:
            self.lbl_alerta.text = "⚠️ ERROR: Paciente NO reportó embarazo en Paso 3"
            self.scroll.scroll_y = 0.35
            print("[!] Alerta clínica: No se puede registrar diagnóstico de embarazo si indicó que NO está embarazada.")
            return

        # A-2) Si el diagnóstico es de PUERPERIO pero en Paso 3 NO indicó ser puérpera -> ERROR
        if tiene_puerperio and not es_si_puerpera:
            self.lbl_alerta.text = "⚠️ ERROR: Paciente NO reportó puerperio en Paso 3"
            self.scroll.scroll_y = 0.35
            print("[!] Alerta clínica: No se puede registrar diagnóstico de puerperio si indicó que NO es puérpera.")
            return
           # <<< ¡AQUÍ ESTÁ LA NUEVA REGLA QUE NECESITABAS! >>>
        # A-3) Si en Paso 3 indicó SÍ SER PUÉRPERA, el diagnóstico OBLIGATORIAMENTE debe ser de puerperio o parto
        if es_si_puerpera and not (tiene_puerperio or tiene_otro_obstetrico):
            self.lbl_alerta.text = "⚠️ ERROR: Paciente puérpera requiere diag. de puerperio"
            self.scroll.scroll_y = 0.35
            return  

        # B) Si es Hombre y el diagnóstico es obstétrico -> ERROR
        if sexo == 'H' and contiene_obstetrico:
            self.lbl_alerta.text = "⚠️ ERROR: Paciente masculino con diag. obstétrico"
            self.scroll.scroll_y = 1
            print("[!] Alerta clínica: Paciente masculino no puede llevar diagnóstico obstétrico.")
            return

        # C) Si es Mujer pero tiene 9 años o menos y el diagnóstico es obstétrico -> ERROR
        if sexo == 'M' and anios <= 9 and contiene_obstetrico:
            self.lbl_alerta.text = "⚠️ ERROR: Niña <= 9 años con diag. obstétrico"
            self.scroll.scroll_y = 0.35
            print("[!] Alerta clínica: Paciente femenino <= 9 años no puede llevar diagnóstico obstétrico.")
            return

        # D) VALIDACIÓN ESTRICTA PARA EMBARAZADAS: Prohibir textos incompletos como "embarazo de semana"
        if es_si_embarazada and tiene_embarazo:
            patron_numero = r"(?:EMBARAZ|GESTAC).*?(\d{1,2})\s*(?:SEMANA|SEM|S)"
            coincidencia_num = re.search(patron_numero, diag_upper)
            
            # Si menciona embarazo/gestación pero NO trae su número de semanas exacto -> ERROR
            if not coincidencia_num:
                self.lbl_alerta.text = "⚠️ ERROR: Ingrese las semanas (Ej: Embarazo de 12 semanas)"
                self.scroll.scroll_y = 0.35
                print("[!] Alerta clínica: Diagnóstico de embarazo incompleto o sin número detectado.")
                return

        # 3. FILTRO CLÍNICO: EXTRACCIÓN Y VALIDACIÓN DE SEMANAS DE EMBARAZO SEGÚN TRIMESTRE
        trimestre = getattr(app, "trimestre_seleccionado", "N/A").strip()
        
        semanas_gestacion = 0
        palabras = diag_upper.split()
        
        # Búsqueda inteligente de semanas en cada palabra o en la anterior
        for i, palabra in enumerate(palabras):
            if any(term in palabra for term in ["SEMANA", "SEM", "S"]):
                if i > 0:
                    num_prev = ''.join(filter(str.isdigit, palabras[i-1]))
                    if num_prev:
                        semanas_gestacion = int(num_prev)
                        break
                num_self = ''.join(filter(str.isdigit, palabra))
                if num_self:
                    semanas_gestacion = int(num_self)
                    break
        
        if semanas_gestacion == 0 and contiene_obstetrico:
            for palabra in palabras:
                num_p = ''.join(filter(str.isdigit, palabra))
                if num_p:
                    val = int(num_p)
                    if 1 <= val <= 42:
                        semanas_gestacion = val
                        break

        print(f"[DEBUG CLÍNICO] Trimestre: '{trimestre}' | Semanas detectadas: {semanas_gestacion}")

        # Aplicar el filtro estricto de trimestres si se detectaron semanas de embarazo
        if semanas_gestacion > 0:
            if "1" in trimestre:
                if semanas_gestacion > 13:
                    self.lbl_alerta.text = "⚠️ ERROR: >13 sem en 1er Trimestre"
                    self.scroll.scroll_y = 0.35
                    print(f"[!] Alerta clínica bloqueada: {semanas_gestacion} semanas no corresponde al 1er Trimestre.")
                    return

            elif "2" in trimestre:
                if semanas_gestacion < 14 or semanas_gestacion > 27:
                    self.lbl_alerta.text = "⚠️ ERROR: Fuera de rango 2do Trimestre"
                    self.scroll.scroll_y = 0.35
                    print(f"[!] Alerta clínica bloqueada: {semanas_gestacion} semanas no corresponde al 2do Trimestre.")
                    return

            elif "3" in trimestre:
                if semanas_gestacion < 28 or semanas_gestacion >= 43:
                    self.lbl_alerta.text = "⚠️ ERROR: Fuera de rango 3er Trimestre"
                    self.scroll.scroll_y = 0.35
                    print(f"[!] Alerta clínica bloqueada: {semanas_gestacion} semanas no corresponde al 3er Trimestre.")
                    return

        # Si pasa todas las validaciones clínicas, limpiamos la alerta
        self.lbl_alerta.text = ""

        # Procesar nacionalidad
        nac_raw = d.get('nacionalidad', 'V')
        if nac_raw in ["V", "VENEZOLANO", "VENEZOLANA"]:
            nacionalidad = "V"
        elif nac_raw in ["C", "COLOMBIANO", "COLOMBIANA"]:
            nacionalidad = "C"
        elif nac_raw in ["B", "BRASILEÑO", "BRASILEÑA", "BRASIL"]:
            nacionalidad = "B"
        elif nac_raw in ["O", "OTRO", "EXTRANJERO", "OTRO EXTRANJERO"]:
            nacionalidad = "O"
        else:
            nacionalidad = "V"

        embarazada = d.get('embarazo', 'No')
        puerpera = d.get('puerperio', 'No')
        riesgo_prenatal = d.get('riesgo', 'No')
        tipo_consulta = d.get('tipo_consulta', '1')

        # --- CAPTACIÓN AUTOMÁTICA DE PUÉRPERAS ---
        if puerpera.upper() in ["S", "SI", "SÍ"] and tipo_consulta in ["1", "C", "PRIMERA", "NUEVA"]:
            captacion_puerpera = "Sí"
        else:
            captacion_puerpera = "No"

        indigena = d.get('indigena', 'No')
        consulta = tipo_consulta

        try:
            codigo_epi = int(getattr(app, "codigo_epi_seleccionado", 99))
        except (ValueError, TypeError):
            codigo_epi = 99

        txt_cronica = self.input_cronica_manual.text.strip().upper()
        es_cronica_manual = txt_cronica in ["S", "SI", "SÍ", "TRUE", "1"]

        # --- FILTRO DE BLOQUEO DE ENFERMEDAD CRÓNICA EN GRUPOS PRIORITARIOS ---
        # --- FILTRO DE BLOQUEO DE ENFERMEDAD CRÓNICA EN GRUPOS PRIORITARIOS ---
        diag_upper = diagnostico_texto.upper()
        
        # 🛑 Si es de riesgo o terreno, NUNCA es grupo especial
        if es_con_riesgo:
            es_grupo_especial = False
        else:
            grupos_especiales = [
                "RECIEN NACIDO", "RECIÉN NACIDO", "LACTANTE", "TRANSICIONAL", 
                "PREESCOLAR", "EMBARAZ", "GESTAC", 
                "PUERPER", "PUÉRPER"
            ]
            es_grupo_especial = any(g in diag_upper for g in grupos_especiales) or es_si_embarazada or es_si_puerpera

        if es_grupo_especial and es_cronica_manual:
            self.lbl_alerta.text = "⚠️ ERROR: Este grupo o diagnóstico no puede registrarse como enfermedad crónica"
            self.scroll.scroll_y = 0.35
            return

        
        
        # ==========================================================
        # BLINDAJE QUIRÚRGICO Y AUTO-DETECCIÓN INTELIGENTE DE CÓDIGOS EPI
        # ==========================================================
        if codigo_epi in [78, 79, 80, 81, 82, 83, 84, 85]:
            pass
        elif codigo_epi in [99, 100]:
            if "HIPERTENS" in diag_upper or "HTA" in diag_upper:
                codigo_epi = 78
            elif "CARDIOPAT" in diag_upper or "ISQUEMIC" in diag_upper or "CORONAR" in diag_upper:
                codigo_epi = 79
            elif "CEREBROVASC" in diag_upper or "ACV" in diag_upper or "ICTUS" in diag_upper:
                codigo_epi = 80
            elif "EPOC" in diag_upper:
                codigo_epi = 81
            elif "ASMA" in diag_upper:
                codigo_epi = 82
            elif "DIABET" in diag_upper or "DBT" in diag_upper:
                codigo_epi = 83
            elif "TUMOR" in diag_upper or "NEOPLAS" in diag_upper or "CANCER" in diag_upper or "CÁNCER" in diag_upper:
                codigo_epi = 84
            elif "EPILEP" in diag_upper or "CONVULS" in diag_upper or "NEUROL" in diag_upper:
                codigo_epi = 85
            elif es_cronica_manual:
                codigo_epi = 95
            else:
                codigo_epi = 100

        try:
            conexion = sqlite3.connect(DB_PATH)
            cursor = conexion.cursor()

            cursor.execute(
                """
                INSERT INTO pacientes (
                    nombre, cedula, edad, sexo, ambito, embarazada, trimestre, 
                    riesgo_prenatal, puerpera, captacion_puerpera, indigena, nacionalidad, 
                    direccion, consulta, diagnostico_texto, codigo_epi, tipo_diagnostico, 
                    conducta, fecha_consulta, semana_num
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
                (
                    nombre, cedula, edad_raw, sexo, ambito, embarazada, trimestre,
                    riesgo_prenatal, puerpera, captacion_puerpera, indigena, nacionalidad,
                    direccion, consulta, diagnostico_texto, codigo_epi, tipo_diagnostico,
                    conducta, fecha_consulta, semana_num
                ),
            )

            conexion.commit()
            conexion.close()

            print(f"[✔] ¡Paciente {nombre} ({nacionalidad}) guardado exitosamente en SQLite!")

            # 🧹 --- LIMPIEZA DE VARIABLES LÓGICAS --- 🧹
            app.codigo_epi_seleccionado = 99
            app.trimestre_seleccionado = "N/A"
            app.conducta_seleccionada = ""
            app.diagnostico_seleccionado = ""
            app.origen_diagnostico = None
            
            if hasattr(app, 'datos_mod_cache'):
                app.datos_mod_cache = {}
            if hasattr(app, 'origen_conducta'):
                app.origen_conducta = {}

            # 🛑 --- LIMPIEZA QUIRÚRGICA DE LA PANTALLA 1 Y 2 (Dejando la 3 intacta) --- 🛑
            if hasattr(app, 'datos_registro') and isinstance(app.datos_registro, dict):
                # Borramos solo lo que pediste: semana, nombre, cédula (Pantalla 1) y edad, dirección (Pantalla 2)
                app.datos_registro.pop('semana', None)
                app.datos_registro.pop('nombre', None)
                app.datos_registro.pop('cedula', None)
                app.datos_registro.pop('edad', None)
                app.datos_registro.pop('direccion', None)

            # 🧹 --- LIMPIAR SOLO DIAGNÓSTICO OFICIAL Y CONDUCTA EN PANTALLA --- 🧹
            self.input_diagnostico.text = ""
            self.input_conducta.text = ""

            # Y saltamos al menú principal
            self.manager.current = 'menu'

        except Exception as e:
            print(f"[❌] Error al guardar en la base de datos: {e}")

                        
                        
                                
        
      
       
       
       
      
    def on_enter(self):
        self.lbl_alerta.text = ""  # Limpia la alerta al entrar
        app = App.get_running_app()
        d = getattr(app, 'datos_registro', {})
        
        # 1. Asegurar que la celda de enfermedad crónica NUNCA esté bloqueada
        self.input_cronica_manual.readonly = False
        
        # Solo actualiza la caja de texto si se seleccionó explícitamente desde el menú de diagnóstico
        if hasattr(app, 'origen_diagnostico') and app.origen_diagnostico == 'registro':
            if hasattr(app, 'diagnostico_seleccionado') and app.diagnostico_seleccionado:
                import re
                diag_bruto = str(app.diagnostico_seleccionado)
                
                # 1. Borra el número inicial y el punto (ej: "012. ")
                diag_limpio = re.sub(r'^\d+\.\s*', '', diag_bruto)
                
                # 2. Borra los paréntesis del final con su código CIE-10 (ej: "(B20-B24)")
                diag_limpio = re.sub(r'\s*\([^)]*\)$', '', diag_limpio).strip()
                
                # Asignamos el texto ya peladito y puro
                self.input_diagnostico.text = diag_limpio
                
            app.origen_diagnostico = None  # Limpiamos para proteger textos manuales futuros

        # ==========================================================
        # 🛑 BLINDAJE DE ENTRADA: Leemos la celda recién pintada
        # ==========================================================
        diag_actual_upper = self.input_diagnostico.text.strip().upper()
        es_de_riesgo = "RIESGO" in diag_actual_upper or "TERRENO" in diag_actual_upper

        # 2. Evaluación de variables de salud materna y puericultura desde Paso 3
        embarazada_p3 = str(d.get('embarazo', 'No')).strip().upper()
        es_si_embarazada = embarazada_p3 in ["S", "SI", "SÍ", "TRUE", "1"]

        puerpera_p3 = str(d.get('puerperio', 'No')).strip().upper()
        es_si_puerpera = puerpera_p3 in ["S", "SI", "SÍ", "TRUE", "1"]

        puericultura_p3 = str(d.get('puericultura', 'No')).strip().upper()
        
        # 🛑 Si es de riesgo, apagamos la puericultura a la fuerza en la entrada
        if es_de_riesgo:
            es_puericultura = False
        else:
            es_puericultura = puericultura_p3 in ["S", "SI", "SÍ", "TRUE", "1"]

        tipo_consulta = str(d.get('tipo_consulta', '1')).strip().upper()
        is_primera = tipo_consulta in ["1", "C", "PRIMERA", "NUEVA"]

        # 3. Blindaje de Conducta según el programa
        if es_puericultura:
            if is_primera:
                self.input_conducta.text = "CAPTACIÓN"
                self.input_conducta.readonly = True
                if hasattr(self, 'btn_elegir_conducta'):
                    self.btn_elegir_conducta.disabled = True
                    self.btn_elegir_conducta.text = "🔒 Conducta Fijada: CAPTACIÓN"
            else:
                self.input_conducta.text = "PUERICULTURA/CONTROL"
                self.input_conducta.readonly = True
                if hasattr(self, 'btn_elegir_conducta'):
                    self.btn_elegir_conducta.disabled = True
                    self.btn_elegir_conducta.text = "🔒 Conducta Fijada: PUERICULTURA/CONTROL"

           
        elif es_si_embarazada or es_si_puerpera:
            if is_primera:
                self.input_conducta.text = "CAPTACIÓN"
                self.input_conducta.readonly = True
                if hasattr(self, 'btn_elegir_conducta'):
                    self.btn_elegir_conducta.disabled = True
                    self.btn_elegir_conducta.text = "🔒 Conducta Fijada: CAPTACIÓN"
            else:
                self.input_conducta.text = "CONTROL"
                self.input_conducta.readonly = True
                if hasattr(self, 'btn_elegir_conducta'):
                    self.btn_elegir_conducta.disabled = True
                    self.btn_elegir_conducta.text = "🔒 Conducta Fijada: CONTROL"

        else:
            # Morbilidad General y Terreno: Desbloqueado para selección o edición libre
            self.input_conducta.readonly = False
            if hasattr(self, 'btn_elegir_conducta'):
                self.btn_elegir_conducta.disabled = False
                self.btn_elegir_conducta.text = "🔍 Seleccionar Conducta "

            # Actualiza la conducta si se seleccionó una en el menú correspondiente
            if hasattr(app, 'conducta_seleccionada') and app.conducta_seleccionada:
                self.input_conducta.text = app.conducta_seleccionada
                app.conducta_seleccionada = ""

        
       
      
        
        
        
# --- PANTALLA DE CONSULTA Y MODIFICACIÓN MAESTRA (TODOS LOS CAMPOS) ---
# ==========================================
# MODIFICACIÓN - PASO 1: Búsqueda y Datos Básicos (4 campos)
# ==========================================

class ModificarPaso1Screen(Screen):
    def __init__(self, **kwargs):
        super(ModificarPaso1Screen, self).__init__(**kwargs)
        
        scroll = ScrollView(scroll_timeout=250, scroll_distance=15)
        layout = BoxLayout(orientation='vertical', padding=30, spacing=20, size_hint_y=None)
        layout.bind(minimum_height=layout.setter('height'))
        
        layout.add_widget(Widget(size_hint_y=None, height=10))
        
        lbl_titulo = Label(
            text="CONSULTA Y MODIFICACIÓN\nPaso 1: Búsqueda y Datos Básicos", 
            halign='center', valign='middle', font_size=sp(20), size_hint_y=None, height=95
        )
        lbl_titulo.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_titulo)
        
        layout.add_widget(Widget(size_hint_y=None, height=10))
        
        # Buscador por Cédula
        lbl_ced_bus = Label(text="Cédula del Paciente:", halign='left', size_hint_y=None, height=55, font_size=sp(15))
        lbl_ced_bus.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_ced_bus)
        
        self.input_buscar_cedula = DeslizableView(
            text='', multiline=False, size_hint_x=1.0, size_hint_y=None, height=65, font_size=sp(16), write_tab=False, padding_y=[8, 8]
        )
        layout.add_widget(self.input_buscar_cedula)
        
        # Fecha de la Consulta
        fecha_hoy = datetime.date.today().strftime("%d/%m/%Y")
        lbl_fec_bus = Label(text="Fecha de la Consulta (DD/MM/AAAA):", halign='left', size_hint_y=None, height=55, font_size=sp(15))
        lbl_fec_bus.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_fec_bus)
        
        self.input_buscar_fecha = DeslizableView(
            text=fecha_hoy, multiline=False, size_hint_x=1.0, size_hint_y=None, height=65, font_size=sp(16), write_tab=False, padding_y=[8, 8]
        )
        layout.add_widget(self.input_buscar_fecha)
        
        # Botón Buscar
        btn_buscar = Button(
            text="🔍 Buscar Registro en SQLite", 
            size_hint_y=None, 
            height=65, 
            font_size=sp(16),
            background_normal='',                 
            background_color=(0.2, 0.5, 0.8, 1),  
            color=(0.6, 0.05, 0.05, 1),           
            bold=True                             
        )
        btn_buscar.bind(on_press=self.buscar_paciente)
        layout.add_widget(btn_buscar)

        # Etiqueta de estado para feedback al usuario
        self.lbl_estado = Label(
            text="", 
            halign='center', 
            size_hint_y=None, 
            height=55,
            font_size=sp(14),
            color=(0.9, 0.9, 0.2, 1)
        )
        self.lbl_estado.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(self.lbl_estado)
        
        layout.add_widget(Widget(size_hint_y=None, height=10))

        # Campo 1: Semana
        lbl_sem = Label(text="N° de Semana :", halign='left', size_hint_y=None, height=55, font_size=sp(15))
        lbl_sem.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_sem)
        self.edit_semana = DeslizableView(text='', multiline=False, size_hint_x=1.0, size_hint_y=None, height=65, font_size=sp(16), write_tab=False, padding_y=[8, 8])
        layout.add_widget(self.edit_semana)
        
        # Campo 2: Fecha Consulta
        lbl_fec = Label(text="Fecha de Consulta (DD/MM/AAAA):", halign='left', size_hint_y=None, height=55, font_size=sp(15))
        lbl_fec.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_fec)
        self.edit_fecha = DeslizableView(text='', multiline=False, size_hint_x=1.0, size_hint_y=None, height=65, font_size=sp(16), write_tab=False, padding_y=[8, 8])
        layout.add_widget(self.edit_fecha)
        
        # Campo 3: Nombre
        lbl_nom = Label(text="Nombre y Apellido:", halign='left', size_hint_y=None, height=55, font_size=sp(15))
        lbl_nom.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_nom)
        self.edit_nombre = DeslizableView(text='', multiline=False, size_hint_x=1.0, size_hint_y=None, height=65, font_size=sp(16), write_tab=False, padding_y=[8, 8])
        layout.add_widget(self.edit_nombre)
        
        # Campo 4: Cédula
        lbl_ced = Label(text="Cédula :", halign='left', size_hint_y=None, height=55, font_size=sp(15))
        lbl_ced.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_ced)
        self.edit_cedula = DeslizableView(text='', multiline=False, size_hint_x=1.0, size_hint_y=None, height=65, font_size=sp(16), write_tab=False, padding_y=[8, 8])
        layout.add_widget(self.edit_cedula)
        
        layout.add_widget(Widget(size_hint_y=None, height=15))
        
        # Botones de Navegación
        btn_layout = BoxLayout(orientation='horizontal', spacing=10, size_hint_y=None, height=65)
        btn_volver = Button(text="Menú Principal", background_color=(0.8, 0.3, 0.3, 1), font_size=sp(16), size_hint_y=None, height=65)
        btn_volver.bind(on_press=lambda x: setattr(self.manager, 'current', 'menu'))
        
        btn_sig = Button(text="Siguiente ➡️", background_color=(0.2, 0.5, 0.8, 1), font_size=sp(16), size_hint_y=None, height=65)
        btn_sig.bind(on_press=self.ir_paso_2)
        
        btn_layout.add_widget(btn_volver)
        btn_layout.add_widget(btn_sig)
        layout.add_widget(btn_layout)
        
        layout.add_widget(Widget(size_hint_y=None, height=150))
        scroll.add_widget(layout)
        self.add_widget(scroll)

    def buscar_paciente(self, instance):
        cedula_buscada = self.input_buscar_cedula.text.strip().upper()
        fecha_buscada = self.input_buscar_fecha.text.strip()
        
        if not cedula_buscada or not fecha_buscada:
            self.lbl_estado.text = "⚠️ Por favor escriba la cédula y la fecha."
            return
            
        try:
            conexion = sqlite3.connect(DB_PATH)
            conexion.row_factory = sqlite3.Row
            cursor = conexion.cursor()
            
            cursor.execute('''
                SELECT id, semana_num, fecha_consulta, nombre, cedula, edad, sexo, ambito, 
                       embarazada, trimestre, riesgo_prenatal, puerpera, captacion_puerpera, 
                       indigena, nacionalidad, direccion, consulta, tipo_diagnostico, 
                       diagnostico_texto, codigo_epi, conducta 
                FROM pacientes WHERE cedula = ? AND fecha_consulta = ?
            ''', (cedula_buscada, fecha_buscada))
 
            resultado = cursor.fetchone()
            conexion.close()
            
            app = App.get_running_app()
            if resultado:
                app.paciente_id = resultado['id']
                
                app.datos_mod_cache = dict(resultado)
                app.datos_mod_cache['cedula_original'] = cedula_buscada
                app.datos_mod_cache['fecha_original'] = fecha_buscada

                self.edit_semana.text = str(resultado['semana_num'] or "")
                self.edit_fecha.text = str(resultado['fecha_consulta'] or "")
                self.edit_nombre.text = str(resultado['nombre'] or "")
                self.edit_cedula.text = str(resultado['cedula'] or "")
                
                self.lbl_estado.text = f"[✔] ¡Ficha de {resultado['nombre']} cargada con éxito!"
            else:
                self.lbl_estado.text = "❌ No existe registro para esa cédula en esa fecha."
                app.paciente_id = None
                app.datos_mod_cache = {}
        except Exception as e:
            self.lbl_estado.text = f"❌ Error al buscar en base de datos: {e}"

    def ir_paso_2(self, instance):
        app = App.get_running_app()
        
        if not (hasattr(app, 'paciente_id') and app.paciente_id):
            self.lbl_estado.text = "⚠️ Debe buscar y cargar un paciente primero."
            return

        nueva_cedula = self.edit_cedula.text.strip().upper()
        nueva_fecha = self.edit_fecha.text.strip()
        nuevo_nombre = self.edit_nombre.text.strip().upper()
        nueva_semana = self.edit_semana.text.strip()

        if not nueva_cedula or not nueva_fecha or not nuevo_nombre:
            self.lbl_estado.text = "⚠️ La cédula, fecha y nombre no pueden quedar vacíos."
            return

        if not isinstance(getattr(app, 'datos_mod_cache', None), dict):
            app.datos_mod_cache = {}

        if 'cedula_original' not in app.datos_mod_cache:
            app.datos_mod_cache['cedula_original'] = self.input_buscar_cedula.text.strip().upper()
        if 'fecha_original' not in app.datos_mod_cache:
            app.datos_mod_cache['fecha_original'] = self.input_buscar_fecha.text.strip()

        app.datos_mod_cache['semana_num'] = nueva_semana
        app.datos_mod_cache['fecha_consulta'] = nueva_fecha
        app.datos_mod_cache['nombre'] = nuevo_nombre
        app.datos_mod_cache['cedula'] = nueva_cedula

        self.lbl_estado.text = ""
        
        if hasattr(self, 'manager') and self.manager:
            self.manager.current = 'modificar_paso2'

     

# ==========================================
# MODIFICACIÓN - PASO 2: Demografía y Clínica (4 campos)
# ==========================================
class ModificarPaso2Screen(Screen):
    def __init__(self, **kwargs):
        super(ModificarPaso2Screen, self).__init__(**kwargs)
        scroll = ScrollView(scroll_timeout=250, scroll_distance=15)
        layout = BoxLayout(orientation='vertical', padding=30, spacing=20, size_hint_y=None)
        layout.bind(minimum_height=layout.setter('height'))
        
        layout.add_widget(Widget(size_hint_y=None, height=10))
        
        lbl_titulo = Label(
            text="CONSULTA Y MODIFICACIÓN\nPaso 2: Demografía y Ubicación", 
            halign='center', valign='middle', font_size=sp(20), size_hint_y=None, height=95
        )
        lbl_titulo.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_titulo)
        
        layout.add_widget(Widget(size_hint_y=None, height=10))
        
        # 1. Edad
        lbl_eda = Label(text="Edad (ej. '5' años o '2m' meses):", halign='left', size_hint_y=None, height=55, font_size=sp(15))
        lbl_eda.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_eda)
        self.edit_edad = DeslizableView(text='', multiline=False, size_hint_x=1.0, size_hint_y=None, height=65, font_size=sp(16), write_tab=False, padding_y=[8, 8])
        layout.add_widget(self.edit_edad)
        
        # 2. Sexo
        lbl_sex = Label(text="Sexo (H = Hombre / M = Mujer):", halign='left', size_hint_y=None, height=55, font_size=sp(15))
        lbl_sex.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_sex)
        self.edit_sexo = DeslizableView(text='H', multiline=False, size_hint_x=1.0, size_hint_y=None, height=65, font_size=sp(16), write_tab=False, padding_y=[8, 8])
        layout.add_widget(self.edit_sexo)
        
        # 3. Ámbito
        lbl_amb = Label(text="Ámbito (C = Consultorio / T = Terreno):", halign='left', size_hint_y=None, height=55, font_size=sp(15))
        lbl_amb.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_amb)
        self.edit_ambito = DeslizableView(text='C', multiline=False, size_hint_x=1.0, size_hint_y=None, height=65, font_size=sp(16), write_tab=False, padding_y=[8, 8])
        layout.add_widget(self.edit_ambito)

        # 4. Dirección
        lbl_dir = Label(text="Dirección:", halign='left', size_hint_y=None, height=55, font_size=sp(15))
        lbl_dir.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_dir)
        self.edit_direccion = DeslizableView(text='', multiline=False, size_hint_x=1.0, size_hint_y=None, height=65, font_size=sp(16), write_tab=False, padding_y=[8, 8])
        layout.add_widget(self.edit_direccion)
        
        # 🟢 BOTÓN SELECTOR DE DIRECCIÓN (Sincronizado)
        btn_elegir_dir = Button(
            text="📍 Seleccionar Calle San José / Manual", 
            size_hint_y=None, 
            height=50, 
            font_size=sp(14),
            background_normal=' ',
            background_color=(0.8, 0.3, 0.3, 1),
            color=(1, 1, 1, 1),
            bold=True
        )
        btn_elegir_dir.bind(on_press=self.abrir_menu_direccion)
        layout.add_widget(btn_elegir_dir)

        # 5. Nacionalidad
        lbl_nac = Label(text="Nacionalidad (V / C / B / O):", halign='left', size_hint_y=None, height=55, font_size=sp(15))
        lbl_nac.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_nac)
        self.edit_nacionalidad = DeslizableView(text='V', multiline=False, size_hint_x=1.0, size_hint_y=None, height=65, font_size=sp(16), write_tab=False, padding_y=[8, 8])
        layout.add_widget(self.edit_nacionalidad)
        
        layout.add_widget(Widget(size_hint_y=None, height=10))

        # --- ETIQUETA DE ALERTA ROJA (Misma del Registro) ---
        self.lbl_alerta = Label(
            text="", 
            color=(1, 0, 0, 1),
            bold=True,
            font_size=sp(18),
            halign='center',
            valign='middle',
            size_hint_y=None,
            height=40
        )
        self.lbl_alerta.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(self.lbl_alerta)

        layout.add_widget(Widget(size_hint_y=None, height=15))
        
        # Botones de Navegación
        btn_layout = BoxLayout(orientation='horizontal', spacing=10, size_hint_y=None, height=65)
        btn_ant = Button(text="⬅️ Anterior", background_color=(0.5, 0.5, 0.5, 1), font_size=sp(16), size_hint_y=None, height=65)
        btn_ant.bind(on_press=self.ir_paso_anterior)
        
        btn_sig = Button(text="Siguiente ➡️", background_color=(0.2, 0.5, 0.8, 1), font_size=sp(16), size_hint_y=None, height=65)
        btn_sig.bind(on_press=self.ir_paso_3)
        
        btn_layout.add_widget(btn_ant)
        btn_layout.add_widget(btn_sig)
        layout.add_widget(btn_layout)
        
        layout.add_widget(Widget(size_hint_y=None, height=150))
        scroll.add_widget(layout)
        self.add_widget(scroll)

    def abrir_menu_direccion(self, instance):
        app = App.get_running_app()
        app.origen_direccion = 'modificar'
        self.manager.current = 'modificar_menu_direccion'

       
    def on_enter(self):
        self.lbl_alerta.text = ""
        app = App.get_running_app()
        d = getattr(app, 'datos_mod_cache', {})
        if not isinstance(d, dict):
            d = {}
            
        self.edit_edad.text = str(d.get('edad', '') if d.get('edad', '') is not None else '')
        self.edit_sexo.text = str(d.get('sexo', 'H')).strip().upper()
        self.edit_ambito.text = str(d.get('ambito', 'C')).strip().upper()
        
        # Cargar dirección desde el selector o desde el caché
        if hasattr(app, 'direccion_seleccionada') and app.direccion_seleccionada and getattr(app, 'origen_direccion', None) == 'modificar':
            self.edit_direccion.text = app.direccion_seleccionada
            app.direccion_seleccionada = None
            app.origen_direccion = None
        else:
            self.edit_direccion.text = str(d.get('direccion', '') if d.get('direccion', '') is not None else '')

        self.edit_nacionalidad.text = str(d.get('nacionalidad', 'V')).strip().upper()

    def ir_paso_anterior(self, instance):
        self.lbl_alerta.text = ""
        self.manager.current = 'modificar_paso1'

    def ir_paso_3(self, instance):
        app = App.get_running_app()
        if not hasattr(app, 'datos_mod_cache') or not isinstance(app.datos_mod_cache, dict):
            app.datos_mod_cache = {}
            
        edad = self.edit_edad.text.strip()
        sexo = self.edit_sexo.text.strip().upper()
        ambito = self.edit_ambito.text.strip().upper()
        direccion = self.edit_direccion.text.strip().upper()
        nacionalidad = self.edit_nacionalidad.text.strip().upper()

        # Filtro estricto
        if not edad or not sexo or not ambito or not direccion or not nacionalidad:
            self.lbl_alerta.text = "⚠️ ERROR"
            return

        self.lbl_alerta.text = ""

        app.datos_mod_cache.update({
            'edad': edad,
            'sexo': sexo,
            'ambito': ambito,
            'direccion': direccion,
            'nacionalidad': nacionalidad
        })
        
        if hasattr(self, 'manager') and self.manager:
            self.manager.current = 'modificar_paso3'


# ==========================================
# MODIFICACIÓN - PASO 3: Gestación, Crónica y Dirección (4 campos)
# ==========================================
class ModificarPaso3Screen(Screen):
    def __init__(self, **kwargs):
        super(ModificarPaso3Screen, self).__init__(**kwargs)
        scroll = ScrollView(scroll_timeout=250, scroll_distance=15)
        layout = BoxLayout(orientation='vertical', padding=30, spacing=20, size_hint_y=None)
        layout.bind(minimum_height=layout.setter('height'))
        
        layout.add_widget(Widget(size_hint_y=None, height=10))
        
        # Título unificado
        lbl_titulo = Label(
            text="CONSULTA Y MODIFICACIÓN\nPaso 3: Condiciones Clínicas", 
            halign='center', valign='middle', font_size=sp(20), size_hint_y=None, height=95
        )
        lbl_titulo.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_titulo)
        
        layout.add_widget(Widget(size_hint_y=None, height=10))
        
        # 1. Embarazo
        lbl_emb = Label(text="¿Está Embarazada? (Si / No):", halign='left', size_hint_y=None, height=55, font_size=sp(15))
        lbl_emb.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_emb)
        self.edit_embarazo = DeslizableView(
            text='No', multiline=False, size_hint_x=1.0, size_hint_y=None, height=65, font_size=sp(16), write_tab=False, padding_y=[8, 8]
        )
        layout.add_widget(self.edit_embarazo)
        
        # 2. Trimestre (Campo de texto visible)
        lbl_tri = Label(text="Trimestre de Gestación:", halign='left', size_hint_y=None, height=55, font_size=sp(15))
        lbl_tri.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_tri)
        self.edit_trimestre = DeslizableView(
            text='N/A', multiline=False, size_hint_x=1.0, size_hint_y=None, height=65, font_size=sp(16), write_tab=False, padding_y=[8, 8]
        )
        layout.add_widget(self.edit_trimestre)

        # 3. Botón para Seleccionar Trimestre
        self.btn_trimestre = Button(
            text="🔍 Seleccionar Trimestre ", 
            size_hint_y=None, 
            height=65, 
            font_size=sp(16),
            background_normal='',                 
            background_color=(0.2, 0.5, 0.8, 1),  
            color=(0.6, 0.05, 0.05, 1),           
            bold=True
        )                 
        self.btn_trimestre.bind(on_press=self.abrir_menu_trimestre)
        layout.add_widget(self.btn_trimestre)
        
        # 4. Puerperio
        lbl_puer = Label(text="¿Está en Puerperio? (Si / No):", halign='left', size_hint_y=None, height=55, font_size=sp(15))
        lbl_puer.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_puer)
        self.edit_puerperio = DeslizableView(
            text='No', multiline=False, size_hint_x=1.0, size_hint_y=None, height=65, font_size=sp(16), write_tab=False, padding_y=[8, 8]
        )
        layout.add_widget(self.edit_puerperio)
        
        # 5. Riesgo Prenatal
        lbl_rie = Label(text="¿Tiene Riesgo Prenatal? (Si / No):", halign='left', size_hint_y=None, height=55, font_size=sp(15))
        lbl_rie.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_rie)
        self.edit_riesgo = DeslizableView(
            text='No', multiline=False, size_hint_x=1.0, size_hint_y=None, height=65, font_size=sp(16), write_tab=False, padding_y=[8, 8]
        )
        layout.add_widget(self.edit_riesgo)

        # 6. Indígena (Añadido para guardar consistencia con el registro)
        lbl_ind = Label(text="¿Es Indígena? (Si / No):", halign='left', size_hint_y=None, height=55, font_size=sp(15))
        lbl_ind.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_ind)
        self.edit_indigena = DeslizableView(
            text='No', multiline=False, size_hint_x=1.0, size_hint_y=None, height=65, font_size=sp(16), write_tab=False, padding_y=[8, 8]
        )
        layout.add_widget(self.edit_indigena)
        
        # 7. Tipo de Consulta
        lbl_tip = Label(
            text="Tipo de Consulta (1=Primera / 2=Reconsulta):", 
            halign='left', 
            valign='middle',
            size_hint_y=None, 
            height=75,
            font_size=sp(15)
        )
        lbl_tip.bind(width=lambda s, w: setattr(s, 'text_size', (w, None)))
        layout.add_widget(lbl_tip)
        self.edit_tipo_consulta = DeslizableView(
            text='1', multiline=False, size_hint_x=1.0, size_hint_y=None, height=65, font_size=sp(16), write_tab=False, padding_y=[8, 8]
        )
        layout.add_widget(self.edit_tipo_consulta)
        
        layout.add_widget(Widget(size_hint_y=None, height=10))

        # --- ETIQUETA DE ERROR EN ROJO (IGUAL QUE EN EL REGISTRO) ---
        self.lbl_alerta = Label(
            text="", 
            color=(1, 0, 0, 1), 
            bold=True,
            font_size=sp(18),
            halign='center',
            valign='middle',
            size_hint_y=None,
            height=40
        )
        self.lbl_alerta.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(self.lbl_alerta)
        
        # Botones de Navegación
        btn_layout = BoxLayout(orientation='horizontal', spacing=10, size_hint_y=None, height=65)
        btn_ant = Button(text="⬅️ Anterior", background_color=(0.5, 0.5, 0.5, 1), font_size=sp(16), size_hint_y=None, height=65)
        btn_ant.bind(on_press=self.ir_paso_anterior)
        
        btn_sig = Button(text="Siguiente ➡️", background_color=(0.2, 0.5, 0.8, 1), font_size=sp(16), size_hint_y=None, height=65)
        btn_sig.bind(on_press=self.ir_paso_4)
        
        btn_layout.add_widget(btn_ant)
        btn_layout.add_widget(btn_sig)
        layout.add_widget(btn_layout)
        
        layout.add_widget(Widget(size_hint_y=None, height=150))
        scroll.add_widget(layout)
        self.add_widget(scroll)

    def abrir_menu_trimestre(self, instance):
        app = App.get_running_app()
        app.origen_trimestre = 'modificar'
        app.viene_de_menu_trimestre = True  # <--- ¡Activamos la bandera de protección!
        
        # Guardamos lo que el usuario acaba de escribir antes de irnos
        if hasattr(app, 'datos_mod_cache') and isinstance(app.datos_mod_cache, dict):
            app.datos_mod_cache['embarazada'] = str(self.edit_embarazo.text).strip()
            app.datos_mod_cache['puerperio'] = str(self.edit_puerperio.text).strip()
            app.datos_mod_cache['riesgo_prenatal'] = str(self.edit_riesgo.text).strip()
            app.datos_mod_cache['indigena'] = str(self.edit_indigena.text).strip()
            app.datos_mod_cache['consulta'] = str(self.edit_tipo_consulta.text).strip()
            app.datos_mod_cache['tipo_consulta'] = str(self.edit_tipo_consulta.text).strip()
            app.datos_mod_cache['trimestre'] = str(self.edit_trimestre.text).strip()

        self.manager.current = 'menu_trimestre'

             
    def on_enter(self):
        try:
            app = App.get_running_app()
            self.lbl_alerta.text = "" 
            
            if not hasattr(app, 'datos_mod_cache') or not isinstance(app.datos_mod_cache, dict):
                app.datos_mod_cache = {}
                
            d = app.datos_mod_cache
            
            # 🛑 SI ACABAMOS DE VOLVER DEL MENÚ DE TRIMESTRE:
            # No tocamos los campos para no borrar lo que el usuario escribió (ej: el "Sí" en embarazo).
            if getattr(app, 'viene_de_menu_trimestre', False):
                app.viene_de_menu_trimestre = False # Apagamos la bandera
                
                # Solo inyectamos el trimestre que se seleccionó de forma limpia
                if hasattr(app, 'trimestre_seleccionado') and app.trimestre_seleccionado:
                    val_tri = str(app.trimestre_seleccionado)
                    d['trimestre'] = val_tri
                    self.edit_trimestre.text = val_tri
                    app.trimestre_seleccionado = None
                return # Salimos aquí mismo y evitamos cualquier choque
            
            # SI ENTRADAS POR PRIMERA VEZ DESDE EL PASO 2: Cargamos la caché normal
            self.edit_embarazo.text = str(d.get('embarazada', 'No') or 'No')
            self.edit_trimestre.text = str(d.get('trimestre', 'N/A') or 'N/A')
            self.edit_puerperio.text = str(d.get('puerperio', d.get('puerpera', 'No')) or 'No')
            self.edit_riesgo.text = str(d.get('riesgo_prenatal', 'No') or 'No')
            self.edit_indigena.text = str(d.get('indigena', 'No') or 'No')
            self.edit_tipo_consulta.text = str(d.get('consulta', '1') or d.get('tipo_consulta', '1') or '1')
            
        except Exception as e:
            print(f"[❌] Error en on_enter ModificarPaso3Screen: {e}")

    def ir_paso_anterior(self, instance):
        self.lbl_alerta.text = ""
        self.manager.current = 'modificar_paso2'

    def ir_paso_4(self, instance):
        try:
            app = App.get_running_app()
            if not hasattr(app, 'datos_mod_cache') or not isinstance(app.datos_mod_cache, dict):
                app.datos_mod_cache = {}
                
            # Extracción limpia de texto
            emb = str(getattr(self.edit_embarazo, 'text', 'No') or 'No').strip().capitalize()
            tri = str(getattr(self.edit_trimestre, 'text', 'N/A') or 'N/A').strip()
            puer = str(getattr(self.edit_puerperio, 'text', 'No') or 'No').strip().capitalize()
            ries = str(getattr(self.edit_riesgo, 'text', 'No') or 'No').strip().capitalize()
            ind = str(getattr(self.edit_indigena, 'text', 'No') or 'No').strip().capitalize()
            tip = str(getattr(self.edit_tipo_consulta, 'text', '1') or '1').strip().upper()

            # Validación de campos vacíos
            if not emb or not puer or not ries or not ind or not tip:
                self.lbl_alerta.text = "⚠️ ERROR: Faltan campos por llenar"
                return

            # RECUPERAMOS SEXO Y EDAD DE LA CACHÉ PARA LOS FILTROS CLÍNICOS
            sexo_paciente = str(app.datos_mod_cache.get('sexo', '')).upper()
            edad_raw = str(app.datos_mod_cache.get('edad', '')).upper()
            
            anios = 0
            if 'M' not in edad_raw:
                num_str = ''.join(filter(str.isdigit, edad_raw))
                anios = int(num_str) if num_str else 0
            
            esta_embarazada = emb.startswith('S')
            esta_puerpera = puer.startswith('S')
            
            # FILTROS CLÍNICOS / EPIDEMIOLÓGICOS
            if sexo_paciente == 'H' and (esta_embarazada or esta_puerpera):
                self.lbl_alerta.text = "⚠️ ERROR: Paciente masculino no aplica"
                return
                
            if sexo_paciente == 'M' and anios <= 9 and (esta_embarazada or esta_puerpera):
                self.lbl_alerta.text = "⚠️ ERROR: Menor de 10 años no aplica"
                return

            # 1. INCOMPATIBILIDAD BIOLÓGICA: No puede ser Embarazada y Puérpera a la vez
            if esta_embarazada and esta_puerpera:
                self.lbl_alerta.text = "⚠️ ERROR: No puede ser Embarazada y Puérpera a la vez"
                return

            # 2. FRENO OBLIGATORIO DE TRIMESTRE: Si está embarazada pero no seleccionó trimestre
            if esta_embarazada and (tri == "N/A" or not tri):
                self.lbl_alerta.text = "⚠️ ERROR: Debe seleccionar el trimestre"
                return

            # Limpiamos alerta si todo está bien
            self.lbl_alerta.text = ""

            # Guardamos en caché asegurando texto plano y compatibilidad total con Pantalla 4
            txt_emb = "Sí" if esta_embarazada else "No"
            txt_puer = "Sí" if esta_puerpera else "No"
            txt_tri = str(tri) if esta_embarazada else 'N/A'

            app.datos_mod_cache['embarazada'] = txt_emb
            app.datos_mod_cache['trimestre'] = txt_tri
            app.datos_mod_cache['puerperio'] = txt_puer  # Clave vital para la Pantalla 4
            app.datos_mod_cache['puerpera'] = txt_puer   # Clave de respaldo
            app.datos_mod_cache['riesgo_prenatal'] = str(ries)
            app.datos_mod_cache['indigena'] = str(ind)
            app.datos_mod_cache['consulta'] = str(tip)
            app.datos_mod_cache['tipo_consulta'] = str(tip)
            
            # ¡CLAVE! Limpiamos el trimestre global para que no rompa la pantalla 4
            if hasattr(app, 'trimestre_seleccionado'):
                app.trimestre_seleccionado = None
            
            # Navegación protegida
            if hasattr(self, 'manager') and self.manager:
                self.manager.current = 'modificar_paso4'
                
        except Exception as e:
            import traceback
            traceback.print_exc()
            print(f"[❌] Error al pasar al Paso 4: {e}")

               
               
# ==========================================
# MODIFICACIÓN - PASO 4: Diagnóstico, Conducta y Guardado (Cierre)
# ==========================================

class ModificarPaso4Screen(Screen):
    def __init__(self, **kwargs):
        super(ModificarPaso4Screen, self).__init__(**kwargs)
        # 1. ScrollView guardado como atributo para controlar el Auto-Scroll
        self.scroll = ScrollView(scroll_timeout=250, scroll_distance=15)
        layout = BoxLayout(orientation='vertical', padding=30, spacing=20, size_hint_y=None)
        layout.bind(minimum_height=layout.setter('height'))
        
        layout.add_widget(Widget(size_hint_y=None, height=10))
        
        # Título unificado
        lbl_titulo = Label(
            text="CONSULTA Y MODIFICACIÓN\nPaso 4: Diagnóstico y Cierre", 
            halign='center', valign='middle', font_size=sp(20), size_hint_y=None, height=95
        )
        lbl_titulo.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_titulo)
        
        layout.add_widget(Widget(size_hint_y=None, height=10))
        
        # 1. Tipo de Diagnóstico
        lbl_tdiag = Label(
            text="Tipo Diagnóstico (C=Confirmado / P=Presuntivo):", 
            halign='left', valign='middle', size_hint_y=None, height=75, font_size=sp(15)
        )
        lbl_tdiag.bind(width=lambda s, w: setattr(s, 'text_size', (w, None)))
        layout.add_widget(lbl_tdiag)
        
        self.edit_tipo_diagnostico = DeslizableView(
            text='C', multiline=False, size_hint_x=1.0, size_hint_y=None, height=65, font_size=sp(16), write_tab=False, padding_y=[8, 8]
        )
        layout.add_widget(self.edit_tipo_diagnostico)

        # 2. Enfermedad Crónica Manual
        lbl_cron = Label(text="¿Es Enfermedad Crónica? (Si / No):", halign='left', size_hint_y=None, height=55, font_size=sp(15))
        lbl_cron.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_cron)
        
        self.edit_cronica_manual = DeslizableView(
            text='No', multiline=False, size_hint_x=1.0, size_hint_y=None, height=65, font_size=sp(16), write_tab=False, padding_y=[8, 8]
        )
        layout.add_widget(self.edit_cronica_manual)

        # 3. Diagnóstico Oficial
        lbl_diag = Label(text="Diagnóstico Oficial :", halign='left', size_hint_y=None, height=55, font_size=sp(15))
        lbl_diag.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_diag)
        
        self.edit_diagnostico = DeslizableView(
            text='', multiline=False, size_hint_x=1.0, size_hint_y=None, height=65, font_size=sp(16), write_tab=False, padding_y=[8, 8]
        )
        layout.add_widget(self.edit_diagnostico)
        
        # Botón de Selección de Diagnóstico
        btn_elegir_diag = Button(
            text="🔍 Seleccionar Diagnóstico ", 
            size_hint_y=None, 
            height=65, 
            font_size=sp(15),
            background_normal='',                 
            background_color=(0.2, 0.5, 0.8, 1),  
            color=(0.6, 0.05, 0.05, 1),           
            bold=True                             
        )
        btn_elegir_diag.bind(on_press=self.abrir_menu_diagnostico)
        layout.add_widget(btn_elegir_diag)
        
        # 4. Conducta
        lbl_cond = Label(text="Conducta :", halign='left', size_hint_y=None, height=55, font_size=sp(15))
        lbl_cond.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(lbl_cond)
        
        self.edit_conducta = DeslizableView(
            text='', multiline=False, size_hint_x=1.0, size_hint_y=None, height=65, font_size=sp(16), write_tab=False, padding_y=[8, 8]
        )
        layout.add_widget(self.edit_conducta)

        # Botón para Selección de Conducta
        self.btn_elegir_conducta = Button(
            text="🔍 Seleccionar Conducta ", 
            size_hint_y=None, 
            height=65, 
            font_size=sp(15),
            background_normal='',
            background_color=(0.2, 0.5, 0.8, 1),
            color=(0.6, 0.05, 0.05, 1),         
            bold=True 
        )
        self.btn_elegir_conducta.bind(on_press=self.abrir_menu_conducta)
        layout.add_widget(self.btn_elegir_conducta)
        
        layout.add_widget(Widget(size_hint_y=None, height=10))

        # --- ETIQUETA DE ERROR EN ROJO (JUSTO ENCIMA DE LOS BOTONES) ---
        self.lbl_alerta = Label(
            text="", 
            color=(1, 0, 0, 1), 
            bold=True,
            font_size=sp(15),
            halign='center',
            valign='middle',
            size_hint_y=None,
            height=90
        )
        self.lbl_alerta.bind(size=lambda s, w: setattr(s, 'text_size', w))
        layout.add_widget(self.lbl_alerta)

        # Botones de Acción
        btn_layout = BoxLayout(orientation='horizontal', spacing=10, size_hint_y=None, height=65)
        
        btn_ant = Button(
            text="⬅️ Anterior", 
            background_color=(0.5, 0.5, 0.5, 1), 
            font_size=sp(16), 
            size_hint_y=None, 
            height=65
        )
        btn_ant.bind(on_press=self.ir_paso_anterior)
        
        btn_guardar_cambios = Button(
            text="💾 Guardar Cambios", 
            background_color=(0.2, 0.7, 0.3, 1), 
            font_size=sp(16), 
            size_hint_y=None, 
            height=65
        )
        btn_guardar_cambios.bind(on_press=self.guardar_modificacion)
        
        btn_layout.add_widget(btn_ant)
        btn_layout.add_widget(btn_guardar_cambios)
        layout.add_widget(btn_layout)
        
        layout.add_widget(Widget(size_hint_y=None, height=150))
        self.scroll.add_widget(layout)
        self.add_widget(self.scroll)

    def on_enter(self):
        try:
            app = App.get_running_app()
            self.lbl_alerta.text = "" # Limpiar alerta
            d = getattr(app, 'datos_mod_cache', {})
            if not isinstance(d, dict):
                d = {}
            
            if hasattr(self, 'edit_tipo_diagnostico') and self.edit_tipo_diagnostico:
                self.edit_tipo_diagnostico.text = str(d.get('tipo_diagnostico', 'C') or 'C').upper()
            
            if hasattr(self, 'edit_diagnostico') and self.edit_diagnostico:
                self.edit_diagnostico.text = str(d.get('diagnostico_texto', '') or '')
                
            if hasattr(self, 'edit_conducta') and self.edit_conducta:
                self.edit_conducta.text = str(d.get('conducta', '') or '')
            
            # Evaluar si la ficha o el código es crónica
            try:
                codigo_cargado = int(d.get('codigo_epi', 100) or 100)
            except (ValueError, TypeError):
                codigo_cargado = 100

            if hasattr(self, 'edit_cronica_manual') and self.edit_cronica_manual:
                if codigo_cargado in [78, 79, 80, 81, 82, 83, 84, 85, 95]:
                    self.edit_cronica_manual.text = "Sí"
                else:
                    self.edit_cronica_manual.text = "No"
            
            # Recuperar selecciones hechas en menús desplegables externos (Con limpieza total de Terreno)
            if hasattr(app, 'diagnostico_seleccionado') and app.diagnostico_seleccionado:
                diag_crudo = str(app.diagnostico_seleccionado).strip()
                # Limpiamos saltos de línea o espacios dobles que traiga el terreno
                diag_limpio = " ".join(diag_crudo.split())
                
                if hasattr(self, 'edit_diagnostico') and self.edit_diagnostico:
                    self.edit_diagnostico.text = diag_limpio
                    d['diagnostico_texto'] = diag_limpio
                app.diagnostico_seleccionado = ""

            # 🎯 --- RECUPERAR CONDUCTA SELECCIONADA DESDE EL MENÚ EXTERNO --- 🎯
            if hasattr(app, 'conducta_seleccionada') and app.conducta_seleccionada:
                conducta_cruda = str(app.conducta_seleccionada).strip()
                conducta_limpia = " ".join(conducta_cruda.split())
                
                if hasattr(self, 'edit_conducta') and self.edit_conducta:
                    self.edit_conducta.text = conducta_limpia
                    d['conducta'] = conducta_limpia # Actualizamos también el caché local
                app.conducta_seleccionada = "" # Limpiamos la variable global
               
            # --- BLINDAJE AUTOMÁTICO DE CONDUCTA SEGÚN DATOS DEL PACIENTE ---
            embarazada_p3 = str(d.get('embarazada', 'No')).strip().upper()
            es_si_embarazada = embarazada_p3 in ["S", "SI", "SÍ", "TRUE", "1"]

            puerpera_p3 = str(d.get('puerperio', 'No')).strip().upper()
            es_si_puerpera = puerpera_p3 in ["S", "SI", "SÍ", "TRUE", "1"]

            puericultura_p3 = str(d.get('puericultura', 'No')).strip().upper()
            es_puericultura = puericultura_p3 in ["S", "SI", "SÍ", "TRUE", "1"]

            tipo_consulta = str(d.get('consulta', d.get('tipo_consulta', '1'))).strip().upper()
            is_primera = tipo_consulta in ["1", "C", "PRIMERA", "NUEVA"]

            if es_puericultura:
                self.edit_conducta.text = "CONTROL"
                if hasattr(self, 'btn_elegir_conducta'):
                    self.btn_elegir_conducta.disabled = True
                    self.btn_elegir_conducta.text = "🔒 Conducta Fijada: CONTROL"
            elif es_si_embarazada or es_si_puerpera:
                if is_primera:
                    self.edit_conducta.text = "CAPTACIÓN"
                    if hasattr(self, 'btn_elegir_conducta'):
                        self.btn_elegir_conducta.disabled = True
                        self.btn_elegir_conducta.text = "🔒 Conducta Fijada: CAPTACIÓN"
                else:
                    self.edit_conducta.text = "CONTROL"
                    if hasattr(self, 'btn_elegir_conducta'):
                        self.btn_elegir_conducta.disabled = True
                        self.btn_elegir_conducta.text = "🔒 Conducta Fijada: CONTROL"
            else:
                if hasattr(self, 'btn_elegir_conducta'):
                    self.btn_elegir_conducta.disabled = False
                    self.btn_elegir_conducta.text = "🔍 Seleccionar Conducta "

            print("[✔] Pantalla Paso 4 cargada con éxito en el on_enter.")

        except Exception as e:
            import traceback
            traceback.print_exc()
            print(f"[❌] Error crítico en el on_enter del Paso 4: {e}")

       
        
    def ir_paso_anterior(self, instance):
        self.lbl_alerta.text = ""
        self.manager.current = 'modificar_paso3'
       
    def abrir_menu_diagnostico(self, instance):
        app = App.get_running_app()
        app.origen_diagnostico = 'modificar'
        if hasattr(self, 'manager') and self.manager:
            self.manager.current = 'menu_diag'
    
    def abrir_menu_conducta(self, instance):
        app = App.get_running_app()
        app.origen_conducta = 'modificar'
        if hasattr(self, 'manager') and self.manager:
            self.manager.current = 'menu_conducta'

    def guardar_modificacion(self, instance):
        app = App.get_running_app()
        pac_id = getattr(app, 'paciente_id', None)
        
        try:
            d = getattr(app, 'datos_mod_cache', {})
            if not isinstance(d, dict):
                d = {}
            
            # 1. Sincronizar UI viva del Paso 4 (Incluyendo Edad y Conducta)
            tipo_diagnostico = str(getattr(self.edit_tipo_diagnostico, 'text', 'C')).strip().upper()
            diagnostico_texto = str(getattr(self.edit_diagnostico, 'text', '')).strip()
            conducta = str(getattr(self.edit_conducta, 'text', '')).strip()
            
            # Capturar la edad de la UI si la pantalla tiene la cajita 'edit_edad'
            if hasattr(self, 'edit_edad') and self.edit_edad:
                edad_ingresada = str(getattr(self.edit_edad, 'text', '')).strip()
                if edad_ingresada:
                    d['edad'] = edad_ingresada

            d['tipo_diagnostico'] = tipo_diagnostico
            d['diagnostico_texto'] = diagnostico_texto
            d['conducta'] = conducta
            
            # Check manual de crónica
            txt_cronica = str(getattr(self.edit_cronica_manual, 'text', 'No')).strip().upper()
            es_cronica_manual = txt_cronica in ["S", "SI", "SÍ", "TRUE", "1"]
            
            # =========================================================================
            # 🛑 FILTRO DE BLOQUEO DE ENFERMEDAD CRÓNICA EN GRUPOS PRIORITARIOS Y SANOS
            # =========================================================================
            diag_upper_cron = diagnostico_texto.upper()
            es_si_emb_mod = str(d.get('embarazada', 'No')).strip().upper() in ["S", "SI", "SÍ", "TRUE", "1"]
            es_si_pue_mod = str(d.get('puerperio', 'No')).strip().upper() in ["S", "SI", "SÍ", "TRUE", "1"]
            es_puericultura_mod = str(d.get('puericultura', 'No')).strip().upper() in ["S", "SI", "SÍ", "TRUE", "1"]
            
            grupos_especiales_mod = [
                "RECIEN NACIDO", "RECIÉN NACIDO", "LACTANTE", "TRANSICIONAL", 
                "PREESCOLAR", "ESCOLAR", "ADOLESCENTE", "EMBARAZ", "GESTAC", 
                "PUERPER", "PUÉRPER"
            ]
            
            es_grupo_especial_mod = any(g in diag_upper_cron for g in grupos_especiales_mod) or es_si_emb_mod or es_si_pue_mod or es_puericultura_mod

            if es_grupo_especial_mod and es_cronica_manual:
                self.lbl_alerta.text = "⚠️ ERROR: Este grupo o diagnóstico no puede registrarse como enfermedad crónica"
                if hasattr(self, 'scroll') and self.scroll:
                    self.scroll.scroll_y = 0.35
                print("[!] Alerta de modificación: Intento inválido de marcar crónica en paciente protegido.")
                return
            # =========================================================================

            # 2. Extracción de datos generales sincronizados
            semana_num = str(d.get('semana_num', d.get('semana', ''))).strip()
            fecha_consulta = str(d.get('fecha_consulta', d.get('fecha', ''))).strip()
            nombre = str(d.get('nombre', '')).strip().upper()
            cedula = str(d.get('cedula', '')).strip().upper()
            cedula_original = str(d.get('cedula_original', cedula)).strip().upper()
            fecha_original = str(d.get('fecha_original', fecha_consulta)).strip()
            
            raw_edad = d.get('edad', 0)
            anios = 0
            
            # Mantenemos el formato textual original (ej: "8m", "15d", "8") para SQLite sin forzar float
            edad = str(raw_edad).strip() if raw_edad is not None else "0"
            
            try:
                edad_str = str(raw_edad).upper().strip()
                if 'M' not in edad_str and 'D' not in edad_str:
                    num_str = ''.join(filter(str.isdigit, edad_str))
                    anios = int(num_str) if num_str else 0
                elif 'M' in edad_str:
                    num_m = ''.join(filter(str.isdigit, edad_str))
                    anios = int(int(num_m) / 12) if num_m else 0
            except (ValueError, TypeError):
                anios = 0

            sexo = str(d.get('sexo', 'H')).strip().upper()
            ambito = str(d.get('ambito', 'C')).strip().upper()
            
            nac_raw = str(d.get('nacionalidad', 'V')).strip().upper()
            if nac_raw in ["V", "VENEZOLANO", "VENEZOLANA"]:
                nacionalidad = "V"
            elif nac_raw in ["C", "COLOMBIANO", "COLOMBIANA"]:
                nacionalidad = "C"
            elif nac_raw in ["B", "BRASILEÑO", "BRASILEÑA"]:
                nacionalidad = "B"
            elif nac_raw in ["O", "OTRO", "EXTRANJERO"]:
                nacionalidad = "O"
            else:
                nacionalidad = "V"

            embarazada_p3 = str(d.get('embarazada', 'No')).strip().upper()
            es_si_embarazada = embarazada_p3 in ["S", "SI", "SÍ", "TRUE", "1"]
            embarazada = "Sí" if es_si_embarazada else "No"
            
            trimestre = str(d.get('trimestre', 'N/A')).strip() if es_si_embarazada else "N/A"

            puerpera_p3 = str(d.get('puerperio', 'No')).strip().upper()
            es_si_puerpera = puerpera_p3 in ["S", "SI", "SÍ", "TRUE", "1"]
            puerpera = "Sí" if es_si_puerpera else "No"

            riesgo_prenatal = str(d.get('riesgo_prenatal', 'No')).strip()
            tipo_consulta = str(d.get('consulta', d.get('tipo_consulta', '1'))).strip()

            captacion_puerpera = "Sí" if es_si_puerpera and tipo_consulta in ["1", "C", "PRIMERA", "NUEVA"] else "No"
            indigena = str(d.get('indigena', 'No')).strip()
            direccion = str(d.get('direccion', '')).strip()
            consulta = tipo_consulta

            # --- VALIDACIÓN DE CAMPOS VACÍOS ---
            if not cedula or not diagnostico_texto or not tipo_diagnostico or not conducta:
                self.lbl_alerta.text = "⚠️ ERROR: Faltan campos obligatorios"
                if hasattr(self, 'scroll') and self.scroll:
                    self.scroll.scroll_y = 0.35
                print("[!] Error: Faltan campos obligatorios para guardar.")
                return

            # --- FILTROS CLÍNICOS Y EPIDEMIOLÓGICOS RIGUROSOS ---
            diag_upper = diagnostico_texto.upper()
            
            raices_embarazo = ["EMBARAZ", "GESTAC", "GESTANT"]
            raices_puerperio = ["PUERPER", "PUÉRPER"]
            raices_otras = ["PARTO", "CESAR", "ABORTO"]
            
            tiene_embarazo = any(r in diag_upper for r in raices_embarazo)
            tiene_puerperio = any(r in diag_upper for r in raices_puerperio)
            tiene_otro_obstetrico = any(r in diag_upper for r in raices_otras)
            
            contiene_obstetrico = tiene_embarazo or tiene_puerperio or tiene_otro_obstetrico

            if tiene_embarazo and not es_si_embarazada:
                self.lbl_alerta.text = "⚠️ ERROR: Paciente NO reportó embarazo en Paso 3"
                if hasattr(self, 'scroll') and self.scroll:
                    self.scroll.scroll_y = 0.35
                return

            if tiene_puerperio and not es_si_puerpera:
                self.lbl_alerta.text = "⚠️ ERROR: Paciente NO reportó puerperio en Paso 3"
                if hasattr(self, 'scroll') and self.scroll:
                    self.scroll.scroll_y = 0.35
                return

            # --- REGLA SUPREMA: Si en el Paso 3 dijo SÍ embarazada, el diagnóstico OBLIGATORIO debe ser de embarazo ---
            if es_si_embarazada and not tiene_embarazo:
                self.lbl_alerta.text = "⚠️ ERROR: Paciente embarazada requiere diag. de embarazo"
                if hasattr(self, 'scroll') and self.scroll:
                    self.scroll.scroll_y = 0.35
                return

            # --- FILTRO CLÍNICO OBLIGATORIO PARA PUÉRPERAS ---
            if es_si_puerpera and not (tiene_puerperio or tiene_otro_obstetrico):
                self.lbl_alerta.text = "⚠️ ERROR: Paciente puérpera requiere diag. de puerperio"
                if hasattr(self, 'scroll') and self.scroll:
                    self.scroll.scroll_y = 0.35
                return    

            if sexo == 'H' and contiene_obstetrico:
                self.lbl_alerta.text = "⚠️ ERROR: Paciente masculino con diag. obstétrico"
                if hasattr(self, 'scroll') and self.scroll:
                    self.scroll.scroll_y = 1
                return

            if sexo == 'M' and anios <= 9 and contiene_obstetrico:
                self.lbl_alerta.text = "⚠️ ERROR: Niña <= 9 años con diag. obstétrico"
                if hasattr(self, 'scroll') and self.scroll:
                    self.scroll.scroll_y = 0.35
                return

            # --- VALIDACIÓN PROTEGIDA DE SEMANAS Y TRIMESTRE ---
            if es_si_embarazada and tiene_embarazo:
                try:
                    patron_numero = r"(?:EMBARAZ|GESTAC).*?(\d{1,2})\s*(?:SEMANA|SEM|S)"
                    coincidencia_num = re.search(patron_numero, diag_upper)
                    if not coincidencia_num:
                        self.lbl_alerta.text = "⚠️ ERROR: Ingrese las semanas (Ej: Embarazo de 12 semanas)"
                        if hasattr(self, 'scroll') and self.scroll:
                            self.scroll.scroll_y = 0.35
                        return
                except Exception as err_re:
                    print(f"[!] Error al buscar patrón de semanas: {err_re}")

            semanas_gestacion = 0
            try:
                palabras = diag_upper.split()
                for i, palabra in enumerate(palabras):
                    if any(term in palabra for term in ["SEMANA", "SEM", "S"]):
                        if i > 0:
                            num_prev = ''.join(filter(str.isdigit, palabras[i-1]))
                            if num_prev:
                                semanas_gestacion = int(num_prev)
                                break
                        num_self = ''.join(filter(str.isdigit, palabra))
                        if num_self:
                            semanas_gestacion = int(num_self)
                            break
            except Exception as err_sem:
                print(f"[!] Error extrayendo número de semanas: {err_sem}")
                semanas_gestacion = 0

            if semanas_gestacion > 0:
                if "1" in trimestre and semanas_gestacion > 13:
                    self.lbl_alerta.text = "⚠️ ERROR: >13 sem en 1er Trimestre"
                    if hasattr(self, 'scroll') and self.scroll:
                        self.scroll.scroll_y = 0.35
                    return
                elif "2" in trimestre and (semanas_gestacion < 14 or semanas_gestacion > 27):
                    self.lbl_alerta.text = "⚠️ ERROR: Fuera de rango 2do Trimestre"
                    if hasattr(self, 'scroll') and self.scroll:
                        self.scroll.scroll_y = 0.35
                    return
                elif "3" in trimestre and (semanas_gestacion < 28 or semanas_gestacion >= 43):
                    self.lbl_alerta.text = "⚠️ ERROR: Fuera de rango 3er Trimestre"
                    if hasattr(self, 'scroll') and self.scroll:
                        self.scroll.scroll_y = 0.35
                    return

            self.lbl_alerta.text = ""

            # 3. AUTO-RECLASIFICACIÓN INTELIGENTE DE CÓDIGO EPI
            if es_cronica_manual:
                codigo_epi = 95  # Otras crónicas
            else:
                cod_menu = getattr(app, "codigo_epi_seleccionado", None)
                if cod_menu and str(cod_menu).isdigit() and int(cod_menu) not in [99, 100]:
                    codigo_epi = int(cod_menu)
                else:
                    if "HIPERTENS" in diag_upper or "HTA" in diag_upper:
                        codigo_epi = 78
                    elif "CARDIOPAT" in diag_upper or "ISQUEMIC" in diag_upper:
                        codigo_epi = 79
                    elif "CEREBROVASC" in diag_upper or "ACV" in diag_upper:
                        codigo_epi = 80
                    elif "EPOC" in diag_upper:
                        codigo_epi = 81
                    elif "ASMA" in diag_upper:
                        codigo_epi = 82
                    elif "DIABET" in diag_upper or "DBT" in diag_upper:
                        codigo_epi = 83
                    elif "TUMOR" in diag_upper or "NEOPLAS" in diag_upper or "CANCER" in diag_upper or "CÁNCER" in diag_upper:
                        codigo_epi = 84
                    elif "EPILEP" in diag_upper or "CONVULS" in diag_upper:
                        codigo_epi = 85
                    elif embarazada == "Sí":
                        codigo_epi = 12  # Gestantes
                    else:
                        codigo_epi = 100  # Adulto general / Sin subgrupo

            # 4. RUTA ABSOLUTA DE BD
            try:
                ruta_db = DB_PATH
            except NameError:
                BASE_DIR = os.path.dirname(os.path.abspath(__file__))
                ruta_db = os.path.join(BASE_DIR, "clinica.db")

            conexion = sqlite3.connect(ruta_db)
            cursor = conexion.cursor()

            if not pac_id:
                cursor.execute("SELECT id FROM pacientes WHERE cedula = ? AND fecha_consulta = ? ORDER BY id DESC LIMIT 1", (cedula_original, fecha_original))
                res = cursor.fetchone()
                if res:
                    pac_id = res[0]
                else:
                    cursor.execute("SELECT id FROM pacientes WHERE cedula = ? ORDER BY id DESC LIMIT 1", (cedula_original,))
                    res = cursor.fetchone()
                    if res:
                        pac_id = res[0]

            if not pac_id:
                self.lbl_alerta.text = "⚠️ ERROR: No se encontró la ficha en BD"
                conexion.close()
                return

            # 5. UPDATE DEFINITIVO Y SINCRONIZADO
            cursor.execute('''
                UPDATE pacientes SET
                    semana_num = ?, 
                    fecha_consulta = ?, 
                    nombre = ?, 
                    cedula = ?, 
                    edad = ?, 
                    sexo = ?, 
                    ambito = ?, 
                    embarazada = ?, 
                    trimestre = ?, 
                    riesgo_prenatal = ?, 
                    puerpera = ?, 
                    captacion_puerpera = ?, 
                    indigena = ?, 
                    nacionalidad = ?, 
                    direccion = ?, 
                    consulta = ?, 
                    tipo_diagnostico = ?, 
                    diagnostico_texto = ?, 
                    codigo_epi = ?, 
                    conducta = ?
                WHERE id = ?
            ''', (
                semana_num, fecha_consulta, nombre, cedula, edad, sexo, ambito, 
                embarazada, trimestre, riesgo_prenatal, puerpera, captacion_puerpera, 
                indigena, nacionalidad, direccion, consulta, tipo_diagnostico, 
                diagnostico_texto, codigo_epi, conducta, pac_id
            ))

            conexion.commit()
            conexion.close()

            print(f"[✔] ¡Ficha del paciente {nombre} (Cédula: {cedula}) ACTUALIZADA PERFECTION (EPI: {codigo_epi}, Trimestre: {trimestre})!")

            # CleanUp final
            app.paciente_id = None
            app.datos_mod_cache = {}
            app.codigo_epi_seleccionado = 99
            app.trimestre_seleccionado = "N/A"
            app.diagnostico_seleccionado = ""
            app.conducta_seleccionada = ""
            
            if hasattr(self, 'manager') and self.manager:
                self.manager.current = 'menu'

        except Exception as e:
            import traceback
            error_completo = traceback.format_exc()
            print(f"[❌] Error en guardar_modificacion: {error_completo}")
            
            # Guardamos el error en un archivo de texto en el celular
            try:
                with open("error_log.txt", "w", encoding="utf-8") as f:
                    f.write(error_completo)
                print("[!] ¡Error guardado en error_log.txt!")
            except Exception as ex_archivo:
                print(f"[!] No se pudo escribir el log: {ex_archivo}")

      
           
                      
        
        

# --- CLASE PRINCIPAL DE LA APLICACIÓN KIVY ---
class ClinicaApp(App):
    # Variables de sesión global para médico y consultorio
    nombre_medico = StringProperty("")
    nombre_consultorio = StringProperty("")
    conducta_seleccionada = StringProperty("")  # <--- ¡Variable que faltaba!

    def build(self):
        # Configuramos el modo del teclado virtual en Android a 'resize' para evitar tirones de layout
        from kivy.core.window import Window
        if Window:
            Window.softinput_mode = 'resize'

        # Inicializamos la base de datos de pacientes si no existe
        inicializar_bd()
        
        # Creamos el Administrador de Pantallas
        sm = ScreenManager()
        
        # Añadimos la pantalla de configuración inicial como pantalla de arranque
        # 1. Agregamos la de bienvenida de primera
        sm.add_widget(PantallaSplash(name='splash'))
        sm.add_widget(PantallaCedula(name='cedula')) 
        sm.add_widget(PantallaConfiguracion(name='config'))
        
        # Añadimos el menú principal al gestor
        sm.add_widget(PantallaMenu(name='menu'))
        
        # --- LOS 7 PASOS DE EXPORTACIÓN (DISTRIBUIDOS EN MÁXIMO 5 PREGUNTAS) ---
        sm.add_widget(PantallaExportarPaso1(name='exportar_paso1'))
        sm.add_widget(PantallaExportarPaso2(name='exportar_paso2'))
        
        sm.add_widget(PantallaExportarPaso4(name='exportar_paso4'))
        sm.add_widget(PantallaExportarPaso5(name='exportar_paso5'))
        sm.add_widget(PantallaExportarPaso6(name='exportar_paso6'))
        sm.add_widget(PantallaExportarPaso7(name='exportar_paso7'))  # 🟢 ¡Paso final con el botón de exportar!
        # ------------------------------------------------------------------
        
        # --- LOS 4 PASOS DEL REGISTRO DE PACIENTES ---
        sm.add_widget(PantallaRegistroPaso1(name='registro_paso1'))
        sm.add_widget(PantallaRegistroPaso2(name='registro_paso2'))
        sm.add_widget(PantallaRegistroPaso3(name='registro_paso3'))
        sm.add_widget(PantallaRegistroPaso4(name='registro_paso4'))
        # ---------------------------------------------

        # --- LOS 4 PASOS DE CONSULTA Y MODIFICACIÓN (NUEVO) ---
        sm.add_widget(ModificarPaso1Screen(name='modificar_paso1'))
        sm.add_widget(ModificarPaso2Screen(name='modificar_paso2'))
        sm.add_widget(ModificarPaso3Screen(name='modificar_paso3'))
        sm.add_widget(ModificarPaso4Screen(name='modificar_paso4'))
        # -------------------------------------------------------

        sm.add_widget(PantallaMenuDiagnostico(name='menu_diag'))
        sm.add_widget(PantallaMenuConducta(name='menu_conducta'))
        sm.add_widget(PantallaEPI12(name='diag_epi'))
        sm.add_widget(PantallaAtencionPrimaria(name='diag_ap'))
        sm.add_widget(PantallaMenuTerreno(name='menu_terreno'))
        
        sm.add_widget(PantallaTrimestre(name='menu_trimestre'))
        # 🟢 Pantalla de dirección
        sm.add_widget(PantallaSeleccionDireccion(name='menu_direccion'))
        sm.add_widget(ModificarSeleccionDireccion(name='modificar_menu_direccion'))

        # Nota: Ya quitamos la vieja `PantallaConsultar` porque ahora usa los 4 pasos de arriba.
  
        return sm


if __name__ == '__main__':
    # Ejecutamos la aplicación móvil/escritorio
    ClinicaApp().run()

           