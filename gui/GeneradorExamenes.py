import tkinter as tk
from tkinter import messagebox, simpledialog
import os
import string
from docx import Document
from docx2pdf import convert
import shutil
import re
import random
import platform
import subprocess

class SoftwareExamenAdmision:
    def __init__(self, root):
        self.root = root
        self.root.title("Software Examen Admision")
        self.root.geometry("450x350")
        self.root.resizable(False, False)
        
        # Configurar el color de fondo de la ventana principal
        self.root.configure(bg="#f0f0f0")
        
        # Título principal
        title_label = tk.Label(
            root, 
            text="Software Examen Admision", 
            font=("Arial", 16, "bold"),
            bg="#f0f0f0"
        )
        title_label.pack(pady=30)
        
        # Frame para los botones
        button_frame = tk.Frame(root, bg="#f0f0f0")
        button_frame.pack(expand=True)
        
        # Botón Ver Examenes
        self.ver_btn = tk.Button(
            button_frame,
            text="Ver Examenes",
            width=15,
            command=self.ver_examenes,
            relief=tk.RAISED,
            bg="#e0e0e0",
            font=("Arial", 10)
        )
        self.ver_btn.pack(pady=10)
        
        # Botón Generar Examenes
        self.generar_btn = tk.Button(
            button_frame,
            text="Generar Examenes",
            width=15,
            command=self.generar_examenes,
            relief=tk.RAISED,
            bg="#e0e0e0",
            font=("Arial", 10)
        )
        self.generar_btn.pack(pady=10)
        
        # Definir rutas
        self.docs_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "docs")
        self.examenes_generados_path = os.path.join(self.docs_path, "Examenes Generados")
        self.examen_original_path = os.path.join(self.docs_path, "Examen Original")
        self.plantilla_path = os.path.join(self.examen_original_path, "Examen Admision.docx")
        
        # Crear carpetas necesarias
        self.crear_carpetas()
        
    def generar_examenes(self):
        """Generar exámenes en PDF reemplazando el [TEMA] con letras desde A hasta Z."""
        try:
            # Verificar que exista la plantilla
            if not os.path.exists(self.plantilla_path):
                messagebox.showerror(
                    "Error", 
                    f"No se encontró la plantilla en {self.plantilla_path}. "
                    "Por favor, asegúrese de colocar la plantilla en esa ubicación."
                )
                return
            
            # Preguntar cantidad de exámenes a generar
            cantidad = simpledialog.askinteger(
                "Cantidad", 
                "¿Cuántos exámenes desea generar? (máximo 26 para temas A-Z)", 
                minvalue=1, 
                maxvalue=26
            )
            
            if not cantidad:
                return
            
            # Crear diálogo de progreso
            progress_dialog = tk.Toplevel(self.root)
            progress_dialog.title("Generando Exámenes")
            progress_dialog.geometry("300x100")
            progress_dialog.transient(self.root)
            progress_dialog.grab_set()
            
            progress_label = tk.Label(progress_dialog, text="Preparando...")
            progress_label.pack(pady=20)
            
            progress_dialog.update()
            
            # Limpiar carpeta de exámenes temporales (si existe)
            temp_dir = os.path.join(self.docs_path, "temp")
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
            os.makedirs(temp_dir)
            
            # Generar los exámenes
            temas = list(string.ascii_uppercase)[:cantidad]  # A, B, C, ..., Z
            
            for i, tema in enumerate(temas):
                # Actualizar etiqueta de progreso
                progress_label.config(text=f"Generando examen con tema {tema}... ({i+1}/{cantidad})")
                progress_dialog.update()
                
                # Crear una copia temporal de la plantilla
                temp_docx = os.path.join(temp_dir, f"Examen_Tema_{tema}.docx")
                
                # Cargar el documento
                doc = Document(self.plantilla_path)

                # Reemplazar tema en el documento actual
                for parrafo in doc.paragraphs:
                    for run in parrafo.runs:
                        texto = run.text
                        if texto and "[TEMA]" in texto:
                            nuevo_texto = texto.replace("[TEMA]", str(tema))
                            run.text = nuevo_texto
                
                # extraer preguntas y alternativas
                preguntas = self.extraer_preguntas_alternativas()

                # Reordenar preguntas y alternativas
                preguntas_reordenadas = self.reordenar_preguntas_alternativas(preguntas)

                # Eliminar párrafos con preguntas y alternativas
                for parrafo in doc.paragraphs:
                    if parrafo.text.strip().startswith("1."):
                        parrafo.clear()
                
                # Agregar preguntas y alternativas reordenadas
                for pregunta in preguntas_reordenadas:
                    doc.add_paragraph(f"{pregunta['pregunta']}")
                    for alternativa in pregunta['alternativas']:
                        doc.add_paragraph(f"{alternativa['letra']}) {alternativa['contenido']}")

                # Guardar la versión modificada
                doc.save(temp_docx)

                # Convertir a PDF directamente
                pdf_output = os.path.join(self.examenes_generados_path, f"Examen_Tema_{tema}.pdf")
                convert(temp_docx, pdf_output)

            # Eliminar archivos temporales
            shutil.rmtree(temp_dir)
            
            # Cerrar diálogo de progreso
            progress_dialog.destroy()
            
            messagebox.showinfo("Éxito", f"Se han generado {cantidad} exámenes con temas del A al {temas[-1]} correctamente.")
            
            # Preguntar si desea abrir la carpeta
            if messagebox.askyesno("Abrir Carpeta", "¿Desea abrir la carpeta con los exámenes generados?"):
                self.ver_examenes()
                
        except Exception as e:
            messagebox.showerror("Error", f"No se pudieron generar los exámenes: {str(e)}")
    
    def crear_carpetas(self):
        """Crear las carpetas necesarias si no existen."""
        for path in [self.docs_path, self.examenes_generados_path, self.examen_original_path]:
            if not os.path.exists(path):
                os.makedirs(path)
        
        # Verificar si existe la plantilla original
        if not os.path.exists(self.plantilla_path):
            messagebox.showwarning(
                "Advertencia", 
                f"No se encontró la plantilla en {self.plantilla_path}. "
                "Por favor, asegúrese de colocar la plantilla en esa ubicación."
            )

    def extraer_preguntas_alternativas(self):
        # Cargar el documento
        doc = Document(self.plantilla_path)
        
        # Lista para almacenar preguntas y alternativas
        preguntas = []
        
        # Variables para seguimiento
        pregunta_actual = None
        alternativas = []
        
        # Patrones para identificar preguntas y alternativas según el formato específico
        patron_pregunta = re.compile(r'^(\d+\.)\s*(.*)', re.IGNORECASE)  # Formato: "1." o "2. "
        patron_alternativa = re.compile(r'^([a-e]\))\s*(.*)', re.IGNORECASE)  # Formato: "a)" hasta "e)"
        
        # Recorrer los párrafos del documento
        for parrafo in doc.paragraphs:
            texto = parrafo.text.strip()
            
            # Saltar párrafos vacíos
            if not texto:
                continue
            
            # Verificar si es una pregunta
            match_pregunta = patron_pregunta.match(texto)
            if match_pregunta:
                # Si hay una pregunta anterior, guardarla primero
                if pregunta_actual:
                    preguntas.append({
                        'pregunta': pregunta_actual,
                        'alternativas': alternativas.copy()
                    })
                
                # Iniciar nueva pregunta
                pregunta_actual = match_pregunta.group(2)
                alternativas = []
                continue
            
            # Verificar si es una alternativa
            match_alternativa = patron_alternativa.match(texto)
            if match_alternativa and pregunta_actual:
                letra = texto[0].upper()
                contenido = match_alternativa.group(2)
                alternativas.append({
                    'letra': letra,
                    'contenido': contenido
                })
        
        # Agregar la última pregunta si existe
        if pregunta_actual:
            preguntas.append({
                'pregunta': pregunta_actual,
                'alternativas': alternativas
            })
        
        return preguntas
    
    def reordenar_preguntas_alternativas(self, preguntas):
        # Copiar las preguntas para no modificar el original
        preguntas_reordenadas = preguntas.copy()
        
        # Reordenar aleatoriamente las preguntas
        random.shuffle(preguntas_reordenadas)
        
        # Para cada pregunta, reordenar sus alternativas
        for pregunta in preguntas_reordenadas:
            # Guardar las alternativas y reordenarlas
            alternativas = pregunta['alternativas'].copy()
            random.shuffle(alternativas)
            
            # Asignar nuevas letras según el orden
            letras = ['A', 'B', 'C', 'D', 'E']
            for i, alternativa in enumerate(alternativas):
                if i < len(letras):
                    alternativa['letra'] = letras[i]
            
            pregunta['alternativas'] = alternativas
        
        return preguntas_reordenadas
    
    def ver_examenes(self):
        """Abrir la carpeta de examenes generados."""
        try:
            if not os.path.exists(self.examenes_generados_path):
                os.makedirs(self.examenes_generados_path)
                
            # Abrir la carpeta según el sistema operativo
            if platform.system() == "Windows":
                os.startfile(self.examenes_generados_path)
            elif platform.system() == "Darwin":  # macOS
                subprocess.run(["open", self.examenes_generados_path])
            else:  # Linux
                subprocess.run(["xdg-open", self.examenes_generados_path])
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir la carpeta: {str(e)}")