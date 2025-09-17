import serial
import pandas as pd
from datetime import datetime
import os 
import time
import tkinter as tk
from tkinter import messagebox, simpledialog
import threading
import sys

class BalanzaDataAcquisitionTest:
    def __init__(self):
        self.serial_connection = None
        self.data_list = []
        self.running = False
        self.lote_number = ""
        
    def configurar_puerto_serial(self, puerto="COM1", baudrate=9600, timeout=1):
        """
        Configura la conexión serial con la balanza (MODO PRUEBA)
        """
        try:
            # En modo prueba, simulamos la conexión
            print(f"✓ [MODO PRUEBA] Simulando conexión en {puerto} a {baudrate} baudios")
            print("⚠️  NOTA: No hay balanza real conectada - usando datos simulados")
            return True
        except Exception as e:
            print(f"✗ Error al conectar con la balanza: {e}")
            return False
    
    def leer_datos_balanza(self):
        """
        Lee los datos de la balanza (MODO PRUEBA - datos simulados)
        """
        try:
            # Simular datos de balanza
            import random
            
            # Generar datos aleatorios para prueba
            peso_bruto = round(random.uniform(1.0, 5.0), 3)
            tara = round(random.uniform(0.0, 0.5), 3)
            peso_neto = round(peso_bruto - tara, 3)
            
            datos_procesados = {
                'peso_bruto': peso_bruto,
                'tara': tara,
                'peso_neto': peso_neto
            }
            
            # Agregar timestamp
            timestamp = datetime.now()
            datos_procesados['timestamp'] = timestamp
            datos_procesados['fecha'] = timestamp.strftime('%Y-%m-%d')
            datos_procesados['hora'] = timestamp.strftime('%H:%M:%S')
            datos_procesados['lote'] = self.lote_number
            
            self.data_list.append(datos_procesados)
            print(f"[SIMULADO] Datos leídos: {datos_procesados}")
            
            return datos_procesados
            
        except Exception as e:
            print(f"Error al leer datos: {e}")
            return None
    
    def solicitar_numero_lote(self):
        """
        Solicita el número de lote al usuario mediante interfaz gráfica
        """
        root = tk.Tk()
        root.withdraw()  # Ocultar ventana principal
        
        lote = simpledialog.askstring(
            "Número de Lote", 
            "Ingrese el número de lote:",
            parent=root
        )
        
        root.destroy()
        return lote
    
    def guardar_datos_excel(self):
        """
        Guarda los datos en un archivo Excel
        """
        try:
            if not self.data_list:
                print("No hay datos para guardar")
                return False
            
            # Crear DataFrame
            df = pd.DataFrame(self.data_list)
            
            # Reordenar columnas
            columnas = ['timestamp', 'fecha', 'hora', 'lote', 'peso_bruto', 'tara', 'peso_neto']
            df = df[columnas]
            
            # Nombre del archivo con número de lote
            nombre_archivo = f"datos_balanza_lote_{self.lote_number}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            
            # Crear directorio si no existe
            directorio = "datos_balanza"
            if not os.path.exists(directorio):
                os.makedirs(directorio)
            
            ruta_archivo = os.path.join(directorio, nombre_archivo)
            
            # Guardar en Excel
            df.to_excel(ruta_archivo, index=False, engine='openpyxl')
            
            print(f"✓ Datos guardados en: {ruta_archivo}")
            print(f"✓ Total de registros: {len(self.data_list)}")
            
            return True
            
        except Exception as e:
            print(f"✗ Error al guardar datos: {e}")
            return False
    
    def iniciar_adquisicion(self, puerto="COM1", baudrate=9600, intervalo=1):
        """
        Inicia el proceso de adquisición de datos (MODO PRUEBA)
        """
        print("=== SISTEMA DE ADQUISICIÓN DE DATOS DE BALANZA (MODO PRUEBA) ===")
        print(f"Puerto: {puerto}")
        print(f"Baudrate: {baudrate}")
        print(f"Intervalo de lectura: {intervalo} segundos")
        print("⚠️  MODO PRUEBA: Usando datos simulados")
        print("-" * 50)
        
        # Solicitar número de lote
        self.lote_number = self.solicitar_numero_lote()
        if not self.lote_number:
            print("✗ No se ingresó número de lote. Saliendo...")
            return False
        
        print(f"✓ Número de lote: {self.lote_number}")
        
        # Configurar puerto serial (modo prueba)
        if not self.configurar_puerto_serial(puerto, baudrate):
            return False
        
        # Iniciar adquisición
        self.running = True
        contador = 0
        
        try:
            print("\n=== INICIANDO ADQUISICIÓN DE DATOS (SIMULADA) ===")
            print("Presione Ctrl+C para detener y guardar datos...")
            print("Se generarán 5 lecturas de prueba...")
            
            while self.running and contador < 5:  # Solo 5 lecturas para prueba
                datos = self.leer_datos_balanza()
                if datos:
                    contador += 1
                    print(f"Lectura #{contador}: Peso={datos['peso_bruto']} kg, Tara={datos['tara']} kg, Neto={datos['peso_neto']} kg")
                
                time.sleep(intervalo)
            
            print("\n=== PRUEBA COMPLETADA ===")
            self.running = False
                
        except KeyboardInterrupt:
            print("\n\n=== DETENIENDO ADQUISICIÓN ===")
            self.running = False
        
        finally:
            # Guardar datos
            if self.data_list:
                print("\n=== GUARDANDO DATOS ===")
                if self.guardar_datos_excel():
                    print("✓ Proceso completado exitosamente")
                else:
                    print("✗ Error al guardar datos")
            else:
                print("✗ No se capturaron datos")
        
        return True

def main():
    """
    Función principal del programa (MODO PRUEBA)
    """
    # Configuración por defecto
    PUERTO_DEFAULT = "COM1"
    BAUDRATE_DEFAULT = 9600
    INTERVALO_DEFAULT = 1
    
    # Crear instancia del sistema
    sistema = BalanzaDataAcquisitionTest()
    
    # Iniciar adquisición
    sistema.iniciar_adquisicion(
        puerto=PUERTO_DEFAULT,
        baudrate=BAUDRATE_DEFAULT,
        intervalo=INTERVALO_DEFAULT
    )
    
    # Pausa antes de cerrar
    input("\nPresione Enter para cerrar...")

if __name__ == "__main__":
    main()
