import re
import os
import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter import font as tkFont
import threading
from datetime import datetime

class PDFExtractorGUI:
    def __init__(self, root):
        self.root = root
        self.carpeta_seleccionada = ""
        self.setup_window()
        self.create_widgets()
        
    def setup_window(self):
        """Configurar la ventana principal"""
        self.root.title("Extractor de Datos PDF - SENA")
        self.root.geometry("800x600")
        self.root.minsize(700, 500)
        
        # Configurar colores y estilo
        self.colors = {
            'primary': '#2E86AB',      # Azul principal
            'secondary': '#A23B72',    # Rosa/morado
            'success': '#F18F01',      # Naranja
            'background': '#F5F7FA',   # Gris claro
            'surface': '#FFFFFF',      # Blanco
            'text': '#2D3748',         # Gris oscuro
            'text_light': '#718096'    # Gris medio
        }
        
        self.root.configure(bg=self.colors['background'])
        
        # Configurar fuentes
        self.fonts = {
            'title': tkFont.Font(family="Segoe UI", size=18, weight="bold"),
            'subtitle': tkFont.Font(family="Segoe UI", size=12, weight="bold"),
            'body': tkFont.Font(family="Segoe UI", size=10),
            'button': tkFont.Font(family="Segoe UI", size=11, weight="bold")
        }
        
    def create_widgets(self):
        """Crear todos los widgets de la interfaz"""
        # Frame principal con padding
        main_frame = tk.Frame(self.root, bg=self.colors['background'])
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Header
        self.create_header(main_frame)
        
        # Sección de selección de carpeta
        self.create_folder_section(main_frame)
        
        # Sección de progreso
        self.create_progress_section(main_frame)
        
        # Sección de logs
        self.create_logs_section(main_frame)
        
        # Botones de acción
        self.create_action_buttons(main_frame)
        
    def create_header(self, parent):
        """Crear el header de la aplicación"""
        header_frame = tk.Frame(parent, bg=self.colors['surface'], relief=tk.RAISED, bd=1)
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Título principal
        title_label = tk.Label(
            header_frame,
            text="🔍 Extractor de Datos PDF",
            font=self.fonts['title'],
            bg=self.colors['surface'],
            fg=self.colors['primary'],
            pady=15
        )
        title_label.pack()
        
        # Subtítulo
        subtitle_label = tk.Label(
            header_frame,
            text="Procesamiento automático de fichas de caracterización SENA",
            font=self.fonts['body'],
            bg=self.colors['surface'],
            fg=self.colors['text_light']
        )
        subtitle_label.pack(pady=(0, 10))
        
    def create_folder_section(self, parent):
        """Crear la sección de selección de carpeta"""
        folder_frame = tk.LabelFrame(
            parent,
            text="📁 Selección de Carpeta",
            font=self.fonts['subtitle'],
            bg=self.colors['surface'],
            fg=self.colors['text'],
            padx=15,
            pady=10
        )
        folder_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Frame para el path y botón
        path_frame = tk.Frame(folder_frame, bg=self.colors['surface'])
        path_frame.pack(fill=tk.X, pady=5)
        
        # Entry para mostrar la ruta
        self.path_var = tk.StringVar()
        self.path_var.set("Ninguna carpeta seleccionada")
        
        path_entry = tk.Entry(
            path_frame,
            textvariable=self.path_var,
            font=self.fonts['body'],
            state='readonly',
            bg='#F7FAFC',
            fg=self.colors['text'],
            relief=tk.FLAT,
            bd=5
        )
        path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        # Botón para seleccionar carpeta
        select_btn = tk.Button(
            path_frame,
            text="Seleccionar Carpeta",
            command=self.seleccionar_carpeta,
            font=self.fonts['button'],
            bg=self.colors['primary'],
            fg='white',
            relief=tk.FLAT,
            padx=20,
            pady=8,
            cursor='hand2'
        )
        select_btn.pack(side=tk.RIGHT)
        
        # Hover effects
        def on_enter(e):
            select_btn.configure(bg='#1E5F7A')
        def on_leave(e):
            select_btn.configure(bg=self.colors['primary'])
            
        select_btn.bind("<Enter>", on_enter)
        select_btn.bind("<Leave>", on_leave)
        
    def create_progress_section(self, parent):
        """Crear la sección de progreso"""
        progress_frame = tk.LabelFrame(
            parent,
            text="📊 Progreso del Procesamiento",
            font=self.fonts['subtitle'],
            bg=self.colors['surface'],
            fg=self.colors['text'],
            padx=15,
            pady=10
        )
        progress_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Barra de progreso
        self.progress = ttk.Progressbar(
            progress_frame,
            mode='determinate',
            style='Custom.Horizontal.TProgressbar'
        )
        self.progress.pack(fill=tk.X, pady=5)
        
        # Label de estado
        self.status_var = tk.StringVar()
        self.status_var.set("Listo para procesar")
        
        status_label = tk.Label(
            progress_frame,
            textvariable=self.status_var,
            font=self.fonts['body'],
            bg=self.colors['surface'],
            fg=self.colors['text']
        )
        status_label.pack(pady=(5, 0))
        
        # Configurar estilo de la barra de progreso
        style = ttk.Style()
        style.configure(
            'Custom.Horizontal.TProgressbar',
            background=self.colors['success'],
            troughcolor='#E2E8F0',
            borderwidth=0,
            lightcolor=self.colors['success'],
            darkcolor=self.colors['success']
        )
        
    def create_logs_section(self, parent):
        """Crear la sección de logs"""
        logs_frame = tk.LabelFrame(
            parent,
            text="📝 Registro de Actividad",
            font=self.fonts['subtitle'],
            bg=self.colors['surface'],
            fg=self.colors['text'],
            padx=15,
            pady=10
        )
        logs_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # Frame para el texto y scrollbar
        text_frame = tk.Frame(logs_frame, bg=self.colors['surface'])
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        # Área de texto para logs
        self.logs_text = tk.Text(
            text_frame,
            font=self.fonts['body'],
            bg='#F7FAFC',
            fg=self.colors['text'],
            relief=tk.FLAT,
            bd=5,
            wrap=tk.WORD,
            state=tk.DISABLED
        )
        
        # Scrollbar
        scrollbar = tk.Scrollbar(text_frame, orient=tk.VERTICAL, command=self.logs_text.yview)
        self.logs_text.configure(yscrollcommand=scrollbar.set)
        
        # Pack text y scrollbar
        self.logs_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Mensaje inicial
        self.add_log("🚀 Aplicación iniciada correctamente")
        self.add_log("💡 Selecciona una carpeta con archivos PDF para comenzar")
        
    def create_action_buttons(self, parent):
        """Crear los botones de acción"""
        buttons_frame = tk.Frame(parent, bg=self.colors['background'])
        buttons_frame.pack(fill=tk.X)
        
        self.process_btn = tk.Button(
            buttons_frame,
            text="▶️ EJECUTAR PROCESAMIENTO",
            command=self.procesar_pdfs,
            font=tkFont.Font(family="Segoe UI", size=12, weight="bold"),
            bg=self.colors['success'],
            fg='white',
            relief=tk.FLAT,
            padx=40,
            pady=15,
            cursor='hand2',
            state=tk.DISABLED
        )
        self.process_btn.pack(side=tk.LEFT, padx=(0, 15))
        
        # Botón limpiar logs más pequeño
        clear_btn = tk.Button(
            buttons_frame,
            text="🗑️ Limpiar",
            command=self.limpiar_logs,
            font=self.fonts['body'],
            bg=self.colors['text_light'],
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=10,
            cursor='hand2'
        )
        clear_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Botón salir
        exit_btn = tk.Button(
            buttons_frame,
            text="❌ Salir",
            command=self.root.quit,
            font=self.fonts['body'],
            bg=self.colors['secondary'],
            fg='white',
            relief=tk.FLAT,
            padx=15,
            pady=10,
            cursor='hand2'
        )
        exit_btn.pack(side=tk.RIGHT)
        
        def create_hover_effect(button, normal_color, hover_color):
            def on_enter(e):
                if button['state'] != tk.DISABLED:
                    button.configure(bg=hover_color)
            def on_leave(e):
                if button['state'] != tk.DISABLED:
                    button.configure(bg=normal_color)
            button.bind("<Enter>", on_enter)
            button.bind("<Leave>", on_leave)
        
        create_hover_effect(self.process_btn, self.colors['success'], '#D69E00')
        create_hover_effect(clear_btn, self.colors['text_light'], '#4A5568')
        create_hover_effect(exit_btn, self.colors['secondary'], '#822653')
        
    def add_log(self, mensaje):
        """Agregar mensaje al área de logs"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_message = f"[{timestamp}] {mensaje}\n"
        
        self.logs_text.configure(state=tk.NORMAL)
        self.logs_text.insert(tk.END, log_message)
        self.logs_text.configure(state=tk.DISABLED)
        self.logs_text.see(tk.END)
        self.root.update_idletasks()
        
    def limpiar_logs(self):
        """Limpiar el área de logs"""
        self.logs_text.configure(state=tk.NORMAL)
        self.logs_text.delete(1.0, tk.END)
        self.logs_text.configure(state=tk.DISABLED)
        self.add_log("🧹 Logs limpiados")
        
    def seleccionar_carpeta(self):
        """Seleccionar carpeta con PDFs"""
        carpeta = filedialog.askdirectory(title="Seleccionar carpeta con archivos PDF")
        if carpeta:
            self.carpeta_seleccionada = carpeta
            self.path_var.set(carpeta)
            
            # Contar PDFs en la carpeta
            pdf_count = len([f for f in os.listdir(carpeta) if f.lower().endswith('.pdf')])
            
            if pdf_count > 0:
                self.add_log(f"📁 Carpeta seleccionada: {os.path.basename(carpeta)}")
                self.add_log(f"📄 Se encontraron {pdf_count} archivos PDF")
                self.process_btn.configure(state=tk.NORMAL)
                self.status_var.set(f"Listo para procesar {pdf_count} archivos PDF")
            else:
                self.add_log("⚠️ No se encontraron archivos PDF en la carpeta seleccionada")
                self.process_btn.configure(state=tk.DISABLED)
                self.status_var.set("No hay archivos PDF para procesar")
        
    def procesar_pdfs(self):
        """Procesar todos los PDFs en un hilo separado"""
        if not self.carpeta_seleccionada:
            messagebox.showwarning("Advertencia", "Por favor selecciona una carpeta primero")
            return
            
        # Deshabilitar botón durante el procesamiento
        self.process_btn.configure(state=tk.DISABLED)
        
        # Ejecutar en hilo separado para no bloquear la UI
        thread = threading.Thread(target=self._procesar_pdfs_thread)
        thread.daemon = True
        thread.start()
        
    def _procesar_pdfs_thread(self):
        """Hilo para procesar PDFs sin bloquear la interfaz"""
        try:
            archivos_pdf = [f for f in os.listdir(self.carpeta_seleccionada) if f.lower().endswith('.pdf')]
            total_archivos = len(archivos_pdf)
            
            if total_archivos == 0:
                self.add_log("❌ No se encontraron archivos PDF")
                return
                
            self.add_log(f"🔄 Iniciando procesamiento de {total_archivos} archivos...")
            self.progress['maximum'] = total_archivos
            
            datos = []
            archivos_procesados = 0
            archivos_con_error = 0
            
            for i, archivo in enumerate(archivos_pdf):
                try:
                    ruta_pdf = os.path.join(self.carpeta_seleccionada, archivo)
                    self.status_var.set(f"Procesando: {archivo}")
                    self.add_log(f"📄 Procesando: {archivo}")
                    
                    # Procesar PDF
                    fila = self.procesar_pdf(ruta_pdf)
                    datos.append(fila)
                    archivos_procesados += 1
                    
                    self.add_log(f"✅ {archivo} procesado correctamente")
                    
                except Exception as e:
                    archivos_con_error += 1
                    self.add_log(f"❌ Error procesando {archivo}: {str(e)}")
                
                # Actualizar progreso
                self.progress['value'] = i + 1
                self.root.update_idletasks()
            
            # Guardar resultados
            if datos:
                df = pd.DataFrame(datos)
                salida = os.path.join(self.carpeta_seleccionada, "resultado_total.xlsx")
                df.to_excel(salida, index=False)
                
                self.add_log(f"💾 Archivo Excel guardado: {os.path.basename(salida)}")
                self.add_log(f"📊 Resumen: {archivos_procesados} exitosos, {archivos_con_error} con errores")
                self.status_var.set(f"Completado: {archivos_procesados}/{total_archivos} archivos procesados")
                
                messagebox.showinfo(
                    "Procesamiento Completado",
                    f"Se procesaron {archivos_procesados} archivos correctamente.\n"
                    f"Archivo guardado en: {salida}"
                )
            else:
                self.add_log("❌ No se pudo procesar ningún archivo")
                messagebox.showerror("Error", "No se pudo procesar ningún archivo PDF")
                
        except Exception as e:
            self.add_log(f"❌ Error general: {str(e)}")
            messagebox.showerror("Error", f"Error durante el procesamiento: {str(e)}")
        
        finally:
            # Rehabilitar botón
            self.process_btn.configure(state=tk.NORMAL)
            self.progress['value'] = 0

    # =========================
    # Funciones de procesamiento PDF (copiadas del código original)
    # =========================
    def formatear_cedula(self, raw: str) -> str:
        if not raw:
            return ""
        digits = re.sub(r"[^0-9]", "", raw)
        if not digits:
            return ""
        try:
            return f"{int(digits):,}".replace(",", ".")
        except Exception:
            return digits

    def formatear_fecha(self, dia: str, mes: str, anio: str) -> str:
        return f"{int(dia)}/{int(mes)}/20{int(anio)}"

    def safe_extract_text(self, pdf) -> str:
        pages_text = []
        for page in pdf.pages:
            t = page.extract_text()
            pages_text.append(t or "")
        return "\n".join(pages_text)

    def extraer_horario_maximo(self, pdf) -> str:
        horarios_encontrados = []
        for p_idx, page in enumerate(pdf.pages):
            tables = page.extract_tables()
            if not tables:
                continue
            for t_idx, table in enumerate(tables):
                if not table:
                    continue
                for row_idx, row in enumerate(table):
                    if not row:
                        continue
                    clean_row = [str(cell or "").strip() for cell in row]
                    desde_idx = hasta_idx = None
                    for col_idx, cell in enumerate(clean_row):
                        if "DESDE" in cell.upper():
                            desde_idx = col_idx
                        if "HASTA" in cell.upper():
                            hasta_idx = col_idx
                    if desde_idx is not None and hasta_idx is not None:
                        for data_idx in range(row_idx + 1, len(table)):
                            data_row = table[data_idx]
                            if not data_row or len(data_row) <= max(desde_idx, hasta_idx):
                                continue
                            desde_val = str(data_row[desde_idx] or "").strip()
                            hasta_val = str(data_row[hasta_idx] or "").strip()
                            desde_match = re.search(r'(\d{1,2})', desde_val)
                            hasta_match = re.search(r'(\d{1,2})', hasta_val)
                            if desde_match and hasta_match:
                                desde_hora = int(desde_match.group(1))
                                hasta_hora = int(hasta_match.group(1))
                                rango = hasta_hora - desde_hora
                                horarios_encontrados.append({
                                    'desde': desde_hora,
                                    'hasta': hasta_hora,
                                    'rango': rango,
                                    'formato': f"{desde_hora} A {hasta_hora}"
                                })
                    if "HORARIO" in " ".join(clean_row).upper():
                        numeros = []
                        for cell in clean_row:
                            match = re.search(r'(\d{1,2})', cell)
                            if match:
                                numeros.append(int(match.group(1)))
                        if len(numeros) >= 2:
                            desde_hora = numeros[0]
                            hasta_hora = numeros[1]
                            rango = hasta_hora - desde_hora
                            horarios_encontrados.append({
                                'desde': desde_hora,
                                'hasta': hasta_hora,
                                'rango': rango,
                                'formato': f"{desde_hora} A {hasta_hora}"
                            })
        if horarios_encontrados:
            horario_maximo = max(horarios_encontrados, key=lambda x: x['rango'])
            return horario_maximo['formato']
        return ""

    def procesar_pdf(self, pdf_path):
        with pdfplumber.open(pdf_path) as pdf:
            texto = self.safe_extract_text(pdf)
            horario = self.extraer_horario_maximo(pdf)

        # Código programa
        codigo_programa = ""
        m = re.search(r"c[oó]digo(?:\s+del)?(?:\s+Programa)?(?:\s+o\s+EDT)?[\s:]*([0-9]{4,})", texto, re.I)
        if m:
            codigo_programa = m.group(1)

        # Programa especial / Convenio
        programas_especiales_val = ""
        convenio_val = ""
        m = re.search(r"Programas especiales\s*[:\-]?\s*([^\n\r]*)", texto, re.I)
        if m:
            val = m.group(1).strip()
            if val and val.lower() not in ["programas especiales", "programas especiales:"]:
                programas_especiales_val = val
        m = re.search(r"Convenio\s*[:\-]?\s*([^\n\r]*)", texto, re.I)
        if m:
            val = m.group(1).strip()
            if val and val.lower() not in ["convenio", "convenio:"]:
                convenio_val = val
        if (not programas_especiales_val and not convenio_val) or \
        ((programas_especiales_val.lower() == "no aplica" if programas_especiales_val else True) and
            (convenio_val.lower() == "no aplica" if convenio_val else True)):
            programa_especial = "NINGUNA"
        elif programas_especiales_val and programas_especiales_val.lower() != "no aplica":
            programa_especial = programas_especiales_val
        elif convenio_val and convenio_val.lower() != "no aplica":
            programa_especial = convenio_val
        else:
            programa_especial = "NINGUNA"

        # Cédula
        cedula = ""
        m = re.search(r"cedul[ao]\s*[:\-]?\s*([0-9\.,]+)", texto, re.I)
        if m:
            cedula = self.formatear_cedula(m.group(1))
        else:
            m2 = re.search(r"Instructor[\s\S]{0,80}?([0-9]{6,12})", texto, re.I)
            if m2:
                cedula = self.formatear_cedula(m2.group(1))

        # Fechas
        fecha_inicio_fmt = ""
        fecha_final_fmt = ""
        m = re.search(r"De inicio\s+(\d{1,2})\s+(\d{1,2})\s+(\d{2})", texto, re.I)
        if m:
            fecha_inicio_fmt = self.formatear_fecha(m.group(1), m.group(2), m.group(3))
        m = re.search(r"De finalizaci[oó]n\s+(\d{1,2})\s+(\d{1,2})\s+(\d{2})", texto, re.I)
        if m:
            fecha_final_fmt = self.formatear_fecha(m.group(1), m.group(2), m.group(3))

        # Municipio
        municipio = ""
        m = re.search(r"MUNICIPIO\s*[:\-]?\s*(.*)", texto, re.I)
        if m:
            municipio = m.group(1).strip()

        # Lugar + Vereda
        m_lugar = re.search(r"LUGAR DONDE SE DICTA\s*[:\-]?\s*(.*)", texto, re.I)
        m_vereda = re.search(r"VEREDA\s*[:\-]?\s*(.*)", texto, re.I)
        lugar_val = m_lugar.group(1).strip() if m_lugar and m_lugar.group(1).strip() else ""
        vereda_val = m_vereda.group(1).strip() if m_vereda and m_vereda.group(1).strip() else ""
        if lugar_val and vereda_val:
            if lugar_val.lower() == vereda_val.lower():
                lugar_completo = lugar_val
            else:
                lugar_completo = f"{lugar_val} - {vereda_val}"
        elif lugar_val:
            lugar_completo = lugar_val
        elif vereda_val:
            lugar_completo = vereda_val
        else:
            lugar_completo = ""

        # Días de la semana
        dias_semana = {col: "" for col in ["S", "T", "U", "V", "W", "X", "Y"]}
        dias_mapping = {"LU": "S", "MA": "T", "MI": "U", "JU": "V", "VI": "W", "SA": "X", "DO": "Y"}
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                if not tables:
                    continue
                for table in tables:
                    if not table:
                        continue
                    header = [str(cell or "").strip().upper() for cell in table[0]]
                    if set(dias_mapping.keys()).issubset(header):
                        col_indices = {dia: header.index(dia) for dia in dias_mapping}
                        for row in table[1:]:
                            for dia, col_letter in dias_mapping.items():
                                idx = col_indices[dia]
                                cell_val = str(row[idx] or "").strip()
                                if cell_val.isdigit():
                                    dias_semana[col_letter] = "X"

        # Cupo
        cupo = ""
        m = re.search(r"Cupo\s*[:\-]?\s*(\d+)", texto, re.I)
        if m:
            cupo = int(m.group(1))

        # Diccionario final
        return {
            "D": codigo_programa,
            "H": programa_especial,
            "I": cedula,
            "N": fecha_inicio_fmt,
            "O": fecha_final_fmt,
            "P": municipio,
            "Q": lugar_completo,
            "R": horario,
            **dias_semana,
            "Z": cupo
        }

def main():
    root = tk.Tk()
    app = PDFExtractorGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()