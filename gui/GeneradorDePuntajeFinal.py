import pandas as pd
from fpdf import FPDF
import os
import shutil
import customtkinter as ctk
from tkinter import filedialog, messagebox
import time
import threading

# Configuración global de customtkinter con colores basados en el logo UNSLG
ctk.set_appearance_mode("Light")  # Usando modo claro como base

# Paleta de colores del logo UNSLG
# Rojo: #E31E24
# Amarillo: #FFDD00
# Naranja: #F58220
# Negro: #000000

class ExamenApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistema de Evaluación de Exámenes - UNSLG")
        self.root.geometry("1000x700")
        
        # Variables para almacenar rutas de archivos
        self.archivo_respuestas = ""
        self.archivo_claves = ""
        self.df_resultados = None
        self.pdf_filename = ""
        self.total_preguntas = 100
        
        # Definir colores personalizados basados en el logo
        self.color_principal = "#E31E24"  # Rojo
        self.color_secundario = "#F58220"  # Naranja
        self.color_acento = "#FFDD00"      # Amarillo
        self.color_texto = "#000000"       # Negro
        
        # Variable para controlar animación
        self.stop_animation = False
        self.animation_thread = None
        self.animation_completed = False
        
        self.crear_interfaz()
    
    def crear_interfaz(self):
        # Configurar la ventana principal para que sea escalable
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        
        # Frame principal con color de fondo amarillo suave - usando grid en lugar de pack
        self.frame_principal = ctk.CTkFrame(self.root, fg_color="#FFF9E0")
        self.frame_principal.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)
        
        # Configurar el frame_principal para que sea escalable
        self.frame_principal.grid_rowconfigure(4, weight=1)  # Para que frame_resultados se expanda
        self.frame_principal.grid_columnconfigure(0, weight=1)
        
        # Frame para título con efecto de gradiente
        self.frame_titulo = ctk.CTkFrame(self.frame_principal, fg_color="#FFFBEE", height=80)
        self.frame_titulo.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        
        # Crear un marco para contener los elementos del título
        self.titulo_container = ctk.CTkFrame(self.frame_titulo, fg_color="transparent")
        self.titulo_container.pack(pady=10)
        
        # Icono o logo
        self.logo_frame = ctk.CTkFrame(self.titulo_container, width=40, height=40, fg_color=self.color_principal, corner_radius=10)
        self.logo_frame.pack(side="left", padx=(0, 15))
        
        # Título principal con estilo moderno usando fuente más contemporánea
        self.titulo_principal = ctk.CTkLabel(
            self.titulo_container, 
            text="Sistema de Evaluación", 
            font=ctk.CTkFont(family="Arial", size=30, weight="bold"),
            text_color=self.color_principal
        )
        self.titulo_principal.pack(side="left")
        
        # Subtítulo con estilo moderno
        self.subtitulo = ctk.CTkLabel(
            self.titulo_container, 
            text="", 
            font=ctk.CTkFont(family="Arial", size=16),
            text_color=self.color_secundario
        )
        self.subtitulo.pack(side="left", padx=(15, 0))
        
        # Comenzar animación del título solo una vez
        self.animation_thread = threading.Thread(target=self.animar_titulo_una_vez)
        self.animation_thread.daemon = True
        self.animation_thread.start()
        
        # Frame para los botones de carga de archivos
        frame_archivos = ctk.CTkFrame(self.frame_principal, fg_color="#FFFBEE")
        frame_archivos.grid(row=1, column=0, sticky="ew", padx=20, pady=10)
        
        # Configurar frame_archivos para que se expanda horizontalmente
        frame_archivos.grid_columnconfigure(1, weight=1)
        
        # Frame para el botón de respuestas y su indicador
        frame_btn_respuestas = ctk.CTkFrame(frame_archivos, fg_color="transparent")
        frame_btn_respuestas.grid(row=0, column=0, padx=10, pady=15, sticky="w")
        
        # Indicador circular para respuestas (inicialmente rojo del logo)
        self.indicador_respuestas = ctk.CTkLabel(frame_btn_respuestas, text="", width=20, height=20)
        self.indicador_respuestas.configure(fg_color=self.color_principal, corner_radius=10)
        self.indicador_respuestas.grid(row=0, column=0, padx=(0, 10))
        
        # Botón para cargar respuestas con color naranja del logo
        btn_respuestas = ctk.CTkButton(frame_btn_respuestas, text="Cargar Archivo de Respuestas", 
                                    command=self.cargar_respuestas,
                                    height=40, width=300,
                                    font=ctk.CTkFont(size=14),
                                    fg_color=self.color_secundario,
                                    hover_color="#D07018")  # Naranja más oscuro para hover
        btn_respuestas.grid(row=0, column=1)
        
        # Label para mostrar el nombre del archivo de respuestas
        self.lbl_respuestas = ctk.CTkLabel(frame_archivos, text="No se ha seleccionado ningún archivo", 
                                        text_color=self.color_texto)
        self.lbl_respuestas.grid(row=0, column=1, padx=10, pady=10, sticky="w")
        
        # Frame para el botón de claves y su indicador
        frame_btn_claves = ctk.CTkFrame(frame_archivos, fg_color="transparent")
        frame_btn_claves.grid(row=1, column=0, padx=10, pady=15, sticky="w")
        
        # Indicador circular para claves (inicialmente rojo del logo)
        self.indicador_claves = ctk.CTkLabel(frame_btn_claves, text="", width=20, height=20)
        self.indicador_claves.configure(fg_color=self.color_principal, corner_radius=10)
        self.indicador_claves.grid(row=0, column=0, padx=(0, 10))
        
        # Botón para cargar claves con color naranja del logo
        btn_claves = ctk.CTkButton(frame_btn_claves, text="Cargar Archivo de Claves", 
                                command=self.cargar_claves,
                                height=40, width=300,
                                font=ctk.CTkFont(size=14),
                                fg_color=self.color_secundario,
                                hover_color="#D07018")  # Naranja más oscuro para hover
        btn_claves.grid(row=0, column=1)
        
        # Label para mostrar el nombre del archivo de claves
        self.lbl_claves = ctk.CTkLabel(frame_archivos, text="No se ha seleccionado ningún archivo", 
                                    text_color=self.color_texto)
        self.lbl_claves.grid(row=1, column=1, padx=10, pady=10, sticky="w")
        
        # Separador con color amarillo del logo
        separador = ctk.CTkFrame(self.frame_principal, height=2, fg_color=self.color_acento)
        separador.grid(row=2, column=0, sticky="ew", padx=20, pady=10)
        
        # Botón para procesar con color rojo del logo
        btn_procesar = ctk.CTkButton(self.frame_principal, text="Procesar Evaluaciones", 
                                    command=self.procesar_evaluaciones,
                                    font=ctk.CTkFont(size=18, weight="bold"),
                                    height=50, width=400,
                                    fg_color=self.color_principal,
                                    hover_color="#C91A1F")  # Rojo más oscuro para hover
        btn_procesar.grid(row=3, column=0, pady=20)
        
        # Frame para mostrar resultados con fondo amarillo suave
        self.frame_resultados = ctk.CTkFrame(self.frame_principal, fg_color="#FFFBEE")
        self.frame_resultados.grid(row=4, column=0, sticky="nsew", padx=20, pady=10)
        
        # Configurar el frame_resultados para que sea escalable
        self.frame_resultados.grid_rowconfigure(3, weight=1)  # Para que frame_tabla se expanda
        self.frame_resultados.grid_columnconfigure(0, weight=1)
        
        # Etiqueta para el tema seleccionado con color negro
        self.lbl_tema = ctk.CTkLabel(self.frame_resultados, text="Seleccione un tema:", 
                                    font=ctk.CTkFont(size=16),
                                    text_color=self.color_texto)
        self.lbl_tema.grid(row=0, column=0, pady=10)
        
        # ComboBox para selección de tema con colores del logo
        self.combo_tema = ctk.CTkComboBox(self.frame_resultados, values=["Ninguno"], state="readonly",
                                        command=self.actualizar_tabla,
                                        width=200, height=35,
                                        font=ctk.CTkFont(size=14),
                                        fg_color="white",
                                        border_color=self.color_principal,
                                        button_color=self.color_secundario,
                                        button_hover_color="#D07018",
                                        dropdown_fg_color="white",
                                        dropdown_hover_color=self.color_acento)
        self.combo_tema.grid(row=1, column=0, pady=10)
        
        # Frame para encabezados de la tabla (este será fijo)
        self.frame_headers = ctk.CTkFrame(self.frame_resultados, fg_color=self.color_principal)
        self.frame_headers.header_tag = True
        self.frame_headers.grid(row=2, column=0, sticky="ew", padx=15, pady=(5, 0))
        
        # Frame para la tabla con fondo blanco para contraste - ahora usando grid
        self.frame_tabla = ctk.CTkScrollableFrame(self.frame_resultados, fg_color="white")
        self.frame_tabla.grid(row=3, column=0, sticky="nsew", padx=15, pady=0)
        
        # Botón para generar PDF con color naranja del logo
        self.btn_pdf = ctk.CTkButton(self.frame_resultados, text="Generar y Descargar PDF", 
                                    command=self.generar_descargar_pdf,
                                    state="disabled",
                                    font=ctk.CTkFont(size=16),
                                    height=45, width=350,
                                    fg_color=self.color_secundario,
                                    hover_color="#D07018")  # Naranja más oscuro para hover
        self.btn_pdf.grid(row=4, column=0, pady=15)
        
        # Inicialmente ocultar el frame de resultados
        self.frame_resultados.grid_remove()
    
    def animar_titulo_una_vez(self):
        """Función para crear una animación única para el subtítulo"""
        try:
            # Establecer el color del título
            self.titulo_principal.configure(text_color=self.color_principal)
            
            # Animar subtítulo con un efecto de escritura una sola vez
            texto_subtitulo = "Universidad Nacional San Luis Gonzaga"
            for i in range(len(texto_subtitulo) + 1):
                if self.stop_animation:
                    break
                self.subtitulo.configure(text=texto_subtitulo[:i])
                time.sleep(0.05)
            
            # Marcar la animación como completada
            self.animation_completed = True
            
        except Exception:
            # En caso de error (como al cerrar la ventana), salir de la función
            pass

    def cargar_respuestas(self):
        """Función para cargar el archivo de respuestas"""
        archivo = filedialog.askopenfilename(filetypes=[("Archivos CSV", "*.csv")])
        if archivo:
            self.archivo_respuestas = archivo
            nombre_archivo = os.path.basename(archivo)
            self.lbl_respuestas.configure(text=f"Archivo: {nombre_archivo}")
            # Cambiar indicador a amarillo (indicador de éxito)
            self.indicador_respuestas.configure(fg_color=self.color_acento)
    
    def cargar_claves(self):
        """Función para cargar el archivo de claves"""
        archivo = filedialog.askopenfilename(filetypes=[("Archivos CSV", "*.csv")])
        if archivo:
            self.archivo_claves = archivo
            nombre_archivo = os.path.basename(archivo)
            self.lbl_claves.configure(text=f"Archivo: {nombre_archivo}")
            # Cambiar indicador a amarillo (indicador de éxito)
            self.indicador_claves.configure(fg_color=self.color_acento)
    
    def procesar_evaluaciones(self):
        """Función principal para procesar las evaluaciones"""
        if not self.archivo_respuestas or not self.archivo_claves:
            messagebox.showerror("Error", "Debe seleccionar los archivos de respuestas y claves.")
            return
        
        try:
            # Cargar los archivos CSV
            df_respuestas = pd.read_csv(self.archivo_respuestas, sep=";", encoding="utf-8")
            df_claves = pd.read_csv(self.archivo_claves, sep=";", encoding="utf-8")
            
            # Procesar datos
            resultados = self.calcular_resultados(df_respuestas, df_claves)
            
            # Crear dataframe de resultados
            self.df_resultados = pd.DataFrame(resultados)
            
            # Generar archivo Excel en la misma carpeta que el archivo de respuestas
            directorio = os.path.dirname(self.archivo_respuestas)
            excel_path = os.path.join(directorio, "Resultados_Examenes.xlsx")
            self.df_resultados.to_excel(excel_path, index=False)
            
            # Actualizar interfaz con los resultados
            self.mostrar_resultados()
            
            messagebox.showinfo("Éxito", f"Procesamiento completado. Archivo Excel generado en:\n{excel_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error al procesar los archivos: {str(e)}")
    
    def calcular_resultados(self, df_respuestas, df_claves):
        """Calcula los resultados de cada estudiante"""
        resultados = []
        
        # Iterar sobre cada estudiante en el archivo de respuestas
        for index, estudiante in df_respuestas.iterrows():
            tema_estudiante = estudiante["TEMA"]
            total_puntos = 0
            
            claves_tema = df_claves[df_claves["TEMA"] == tema_estudiante]
            
            if claves_tema.empty:
                continue
            
            claves_tema = claves_tema.iloc[0]
            resultado_estudiante = {"LITHO": estudiante["LITHO"], "TEMA": tema_estudiante, "PUNTOS": 0}
            correctas = 0
            incorrectas = 0
            no_respondidas = 0
            
            for i in range(1, self.total_preguntas + 1):
                clave_pregunta = f"PREG_{i:03d}"
            
                if clave_pregunta in estudiante and clave_pregunta in claves_tema:
                    respuesta = estudiante[clave_pregunta]
                
                    if pd.isna(respuesta) or respuesta == "":  # Si la respuesta está vacía
                        resultado_estudiante[clave_pregunta] = "No respondida"
                        no_respondidas += 1
                    else:
                        es_correcta = respuesta == claves_tema[clave_pregunta]
                        resultado_estudiante[clave_pregunta] = "Correcta" if es_correcta else "Incorrecta"
                    
                        if es_correcta:
                            total_puntos += 20
                            correctas += 1
                        else:
                            total_puntos -= 1.125  # Restar si es incorrecta
                            incorrectas += 1
            
            resultado_estudiante["PUNTOS"] = total_puntos
            resultado_estudiante["CORRECTAS"] = correctas
            resultado_estudiante["INCORRECTAS"] = incorrectas
            resultado_estudiante["NO_RESPONDIDAS"] = no_respondidas
            resultados.append(resultado_estudiante)
            
        return resultados
    
    def mostrar_resultados(self):
        """Muestra los resultados en la interfaz"""
        # Mostrar el frame de resultados usando grid en lugar de pack
        self.frame_resultados.grid(row=4, column=0, sticky="nsew", padx=20, pady=10)
        
        # Actualizar el combobox con los temas disponibles
        temas_unicos = self.df_resultados["TEMA"].unique().tolist()
        if not temas_unicos:
            messagebox.showinfo("Sin datos", "No se encontraron datos para mostrar.")
            return
            
        self.combo_tema.configure(values=temas_unicos)
        self.combo_tema.set(temas_unicos[0])
        
        # Actualizar la tabla con el primer tema
        self.actualizar_tabla(temas_unicos[0])
        
        # Habilitar el botón para generar PDF
        self.btn_pdf.configure(state="normal")
    
    def actualizar_tabla(self, tema_seleccionado):
        """Actualiza la tabla con el tema seleccionado"""
        # Limpiar frames anteriores
        for widget in self.frame_tabla.winfo_children():
            widget.destroy()
        
        # Filtrar datos por tema
        df_filtrado = self.df_resultados[self.df_resultados["TEMA"] == tema_seleccionado].sort_values(by="PUNTOS", ascending=False)
        
        # Verificar que hay datos
        if df_filtrado.empty:
            lbl_no_data = ctk.CTkLabel(self.frame_tabla, text="No hay datos para mostrar con este tema", 
                                    font=ctk.CTkFont(size=14),
                                    text_color=self.color_principal)
            lbl_no_data.pack(pady=20)  # Se mantiene pack aquí por ser único elemento
            return
        
        # Primero limpiamos los headers anteriores
        for widget in self.frame_headers.winfo_children():
            widget.destroy()
        
        # Configurar columnas de encabezados con color blanco para texto sobre fondo rojo
        headers = ["Posición", "LITHO", "PUNTOS", "Correctas", "Incorrectas", "No Resp."]
        for col, header in enumerate(headers):
            lbl = ctk.CTkLabel(self.frame_headers, text=header, 
                            font=ctk.CTkFont(weight="bold", size=14),
                            text_color="white",
                            width=120)
            lbl.grid(row=0, column=col, padx=5, pady=5, sticky="nsew")
            self.frame_headers.grid_columnconfigure(col, weight=1)
        
        # Asegurar que todas las filas usen colores visibles y contrastantes
        for i, (_, row) in enumerate(df_filtrado.iterrows(), 1):
            # Alternar colores amarillo suave y blanco para las filas
            color_fila = "#FFFBEE" if i % 2 == 0 else "#FFFFFF"
            
            # Frame para cada fila dentro del área scrollable
            frame_fila = ctk.CTkFrame(self.frame_tabla, fg_color=color_fila, corner_radius=0)
            frame_fila.pack(fill="x", pady=1)  # Mantenemos pack aquí para las filas dentro del ScrollableFrame
            
            # Configurar el grid del frame para alineación con los encabezados
            for j in range(6):
                frame_fila.grid_columnconfigure(j, weight=1)
            
            # Añadir datos a la fila
            texts = [str(i), str(row["LITHO"]), str(round(row["PUNTOS"], 2)), 
                    str(row["CORRECTAS"]), str(row["INCORRECTAS"]), str(row["NO_RESPONDIDAS"])]
            
            for col, text in enumerate(texts):
                # Color de texto oscuro para máxima visibilidad
                text_color = self.color_texto
                # Resaltar puntos en naranja para destacar
                if col == 2:  # Columna de PUNTOS
                    text_color = self.color_secundario
                
                lbl = ctk.CTkLabel(frame_fila, text=text, text_color=text_color, width=120)
                lbl.grid(row=0, column=col, padx=5, pady=2)
    
    def generar_pdf(self):
        """Genera un PDF con los resultados usando los colores del logo UNSLG"""
        if self.df_resultados is None:
            return None
        
        # Crear el PDF en la misma carpeta que el archivo de respuestas
        directorio = os.path.dirname(self.archivo_respuestas)
        pdf_path = os.path.join(directorio, "Ranking_Examenes.pdf")
        
        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=15)
        
        # Definir colores para mejorar la apariencia visual basados en el logo UNSLG
        # Convertir HEX a RGB para FPDF
        # Rojo: #E31E24 -> (227, 30, 36)
        # Amarillo: #FFDD00 -> (255, 221, 0)
        # Naranja: #F58220 -> (245, 130, 32)
        
        temas_unicos = self.df_resultados["TEMA"].unique()
        
        for tema in temas_unicos:
            pdf.add_page()
            pdf.set_font("Arial", style='B', size=16)
            
            # Título con color rojo del logo
            pdf.set_text_color(227, 30, 36)
            pdf.cell(200, 10, f"Ranking de estudiantes - UNSLG", ln=True, align='C')
            pdf.set_text_color(0, 0, 0)
            pdf.cell(200, 10, f"Tema: {tema}", ln=True, align='C')
            pdf.ln(10)
            
            df_tema = self.df_resultados[self.df_resultados["TEMA"] == tema].sort_values(by="PUNTOS", ascending=False)
            
            # Crear tabla con colores del logo
            pdf.set_font("Arial", style='B', size=12)
            
            # Encabezados con fondo rojo y texto blanco
            pdf.set_fill_color(227, 30, 36)  # Rojo del logo
            pdf.set_text_color(255, 255, 255)  # Texto blanco
            
            pdf.cell(40, 10, "POSICIÓN", border=1, align='C', fill=True)
            pdf.cell(40, 10, "LITHO", border=1, align='C', fill=True)
            pdf.cell(40, 10, "PUNTOS", border=1, align='C', fill=True)
            pdf.cell(40, 10, "CORRECTAS", border=1, align='C', fill=True)
            pdf.cell(40, 10, "INCORRECTAS", border=1, align='C', fill=True)
            pdf.ln()
            
            pdf.set_font("Arial", size=12)
            for pos, (_, row) in enumerate(df_tema.iterrows(), 1):
                # Alternar colores de fondo entre amarillo claro y blanco
                if pos % 2 == 0:
                    pdf.set_fill_color(255, 249, 224)  # Amarillo muy claro
                    fill = True
                else:
                    pdf.set_fill_color(255, 255, 255)  # Blanco
                    fill = True
                
                # Texto en negro para datos
                pdf.set_text_color(0, 0, 0)
                
                pdf.cell(40, 10, str(pos), border=1, align='C', fill=fill)
                pdf.cell(40, 10, str(row['LITHO']), border=1, align='C', fill=fill)
                
                # Usar color naranja para destacar los puntos
                pdf.set_text_color(245, 130, 32)  # Naranja del logo
                pdf.cell(40, 10, f"{row['PUNTOS']:.2f}", border=1, align='C', fill=fill)
                
                # Volver a texto negro para el resto
                pdf.set_text_color(0, 0, 0)
                pdf.cell(40, 10, str(row['CORRECTAS']), border=1, align='C', fill=fill)
                pdf.cell(40, 10, str(row['INCORRECTAS']), border=1, align='C', fill=fill)
                pdf.ln()
            pdf.ln(5)
            
            # Agregar estadísticas del tema con color rojo del logo para títulos
            pdf.set_font("Arial", style='B', size=14)
            pdf.set_text_color(227, 30, 36)  # Rojo del logo
            pdf.cell(200, 10, "Estadísticas", ln=True)
            
            pdf.set_font("Arial", size=12)
            pdf.set_text_color(0, 0, 0)  # Negro para texto normal
            
            promedio = df_tema["PUNTOS"].mean()
            maximo = df_tema["PUNTOS"].max()
            minimo = df_tema["PUNTOS"].min()
            
            pdf.cell(200, 10, f"Promedio de puntos: {promedio:.2f}", ln=True)
            pdf.cell(200, 10, f"Puntuación máxima: {maximo:.2f}", ln=True)
            pdf.cell(200, 10, f"Puntuación mínima: {minimo:.2f}", ln=True)
        
        pdf.output(pdf_path)
        return pdf_path
    
    def generar_descargar_pdf(self):
        """Genera el PDF y ofrece descargarlo"""
        pdf_path = self.generar_pdf()
        if not pdf_path:
            messagebox.showerror("Error", "No hay datos para generar el PDF.")
            return
        
        # Mostrar mensaje de éxito
        messagebox.showinfo("PDF Generado", f"El archivo PDF ha sido generado en:\n{pdf_path}")
        
        # Preguntar si desea guardar en otra ubicación
        respuesta = messagebox.askyesno("Guardar PDF", "¿Desea guardar una copia del PDF en otra ubicación?")
        if respuesta:
            destino = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("Archivos PDF", "*.pdf")],
                title="Guardar PDF como"
            )
            if destino:
                shutil.copy2(pdf_path, destino)
                messagebox.showinfo("Descarga Completada", "El archivo PDF ha sido guardado exitosamente.")
    
    def on_closing(self):
        """Maneja el evento de cierre de la ventana"""
        # Detener la animación antes de cerrar
        self.stop_animation = True
        if self.animation_thread and self.animation_thread.is_alive():
            self.animation_thread.join(0.5)  # Esperar a que termine, pero no más de 0.5 segundos
        self.root.destroy()

def main():
    root = ctk.CTk()
    app = ExamenApp(root)
    # Configurar el manejador para el evento de cierre
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()

if __name__ == "__main__":
    main()
