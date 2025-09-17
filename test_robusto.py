import tkinter as tk
from tkinter import messagebox, simpledialog
from datetime import datetime
import os
import time
import random
import sys

class TestBalanzaRobusto:
    def __init__(self):
        self.data_list = []
        self.lote_number = ""
        
    def solicitar_numero_lote(self):
        """
        Solicita el número de lote al usuario mediante interfaz gráfica
        """
        try:
            root = tk.Tk()
            root.title("Sistema de Balanza - Prueba")
            root.geometry("400x300")
            
            # Centrar ventana
            root.eval('tk::PlaceWindow . center')
            
            # Crear interfaz
            frame = tk.Frame(root, padx=20, pady=20)
            frame.pack(expand=True, fill='both')
            
            tk.Label(frame, text="SISTEMA DE ADQUISICIÓN DE DATOS", 
                    font=("Arial", 16, "bold")).pack(pady=10)
            
            tk.Label(frame, text="Ingrese el número de lote:", 
                    font=("Arial", 12)).pack(pady=10)
            
            lote_entry = tk.Entry(frame, font=("Arial", 12), width=20)
            lote_entry.pack(pady=10)
            lote_entry.focus()
            
            def on_ok():
                self.lote_number = lote_entry.get()
                if self.lote_number:
                    root.destroy()
                else:
                    messagebox.showwarning("Advertencia", "Por favor ingrese un número de lote")
            
            def on_cancel():
                self.lote_number = None
                root.destroy()
            
            button_frame = tk.Frame(frame)
            button_frame.pack(pady=20)
            
            tk.Button(button_frame, text="Aceptar", command=on_ok, 
                    font=("Arial", 10), width=10).pack(side=tk.LEFT, padx=5)
            tk.Button(button_frame, text="Cancelar", command=on_cancel, 
                    font=("Arial", 10), width=10).pack(side=tk.LEFT, padx=5)
            
            # Enter para aceptar
            root.bind('<Return>', lambda e: on_ok())
            
            root.mainloop()
            return self.lote_number
            
        except Exception as e:
            print(f"Error en interfaz gráfica: {e}")
            # Fallback a input de consola
            try:
                lote = input("Ingrese el número de lote: ")
                return lote if lote else None
            except:
                return None
    
    def simular_datos(self):
        """
        Simula datos de balanza
        """
        try:
            peso_bruto = round(random.uniform(1.0, 5.0), 3)
            tara = round(random.uniform(0.0, 0.5), 3)
            peso_neto = round(peso_bruto - tara, 3)
            
            timestamp = datetime.now()
            
            datos = {
                'timestamp': timestamp,
                'fecha': timestamp.strftime('%Y-%m-%d'),
                'hora': timestamp.strftime('%H:%M:%S'),
                'lote': self.lote_number,
                'peso_bruto': peso_bruto,
                'tara': tara,
                'peso_neto': peso_neto
            }
            
            self.data_list.append(datos)
            return datos
            
        except Exception as e:
            print(f"Error al simular datos: {e}")
            return None
    
    def guardar_datos_csv(self):
        """
        Guarda los datos en un archivo CSV
        """
        try:
            if not self.data_list:
                print("No hay datos para guardar")
                return False
            
            # Crear directorio si no existe
            directorio = "datos_balanza"
            if not os.path.exists(directorio):
                os.makedirs(directorio)
            
            # Nombre del archivo
            nombre_archivo = f"datos_balanza_lote_{self.lote_number}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            ruta_archivo = os.path.join(directorio, nombre_archivo)
            
            # Escribir CSV
            with open(ruta_archivo, 'w', encoding='utf-8') as f:
                # Escribir encabezados
                f.write("timestamp,fecha,hora,lote,peso_bruto,tara,peso_neto\n")
                
                # Escribir datos
                for datos in self.data_list:
                    f.write(f"{datos['timestamp']},{datos['fecha']},{datos['hora']},{datos['lote']},{datos['peso_bruto']},{datos['tara']},{datos['peso_neto']}\n")
            
            print(f"✓ Datos guardados en: {ruta_archivo}")
            print(f"✓ Total de registros: {len(self.data_list)}")
            
            return True
            
        except Exception as e:
            print(f"✗ Error al guardar datos: {e}")
            return False
    
    def iniciar_prueba(self):
        """
        Inicia la prueba del sistema
        """
        try:
            print("=== SISTEMA DE ADQUISICIÓN DE DATOS DE BALANZA (PRUEBA) ===")
            print("⚠️  MODO PRUEBA: Usando datos simulados")
            print("-" * 50)
            
            # Solicitar número de lote
            self.lote_number = self.solicitar_numero_lote()
            if not self.lote_number:
                print("✗ No se ingresó número de lote. Saliendo...")
                return False
            
            print(f"✓ Número de lote: {self.lote_number}")
            
            # Simular 5 lecturas
            print("\n=== SIMULANDO LECTURAS DE BALANZA ===")
            for i in range(5):
                datos = self.simular_datos()
                if datos:
                    print(f"Lectura #{i+1}: Peso={datos['peso_bruto']} kg, Tara={datos['tara']} kg, Neto={datos['peso_neto']} kg")
                else:
                    print(f"Error en lectura #{i+1}")
                time.sleep(1)
            
            # Guardar datos
            print("\n=== GUARDANDO DATOS ===")
            if self.guardar_datos_csv():
                print("✓ Proceso completado exitosamente")
            else:
                print("✗ Error al guardar datos")
            
            return True
            
        except Exception as e:
            print(f"✗ Error en la prueba: {e}")
            return False

def main():
    """
    Función principal del programa
    """
    try:
        print("Iniciando sistema de prueba...")
        sistema = TestBalanzaRobusto()
        sistema.iniciar_prueba()
        
        # Pausa antes de cerrar
        print("\nPresione Enter para cerrar...")
        try:
            input()
        except:
            print("Programa finalizado.")
            
    except Exception as e:
        print(f"Error crítico: {e}")
        print("Presione Enter para cerrar...")
        try:
            input()
        except:
            pass

if __name__ == "__main__":
    main()
