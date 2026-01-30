import tkinter as tk
from tkinter import ttk, messagebox
import xlwings as xw
import warnings

warnings.filterwarnings('ignore')

class AsistenteDirecto:
    def __init__(self, root):
        self.root = root
        self.root.title("⚡ Asistente Directo - Escribe en Excel en Tiempo Real")
        self.root.geometry("500x750")
        self.root.configure(bg='#f5f5f5')
        
        self.wb = None
        self.ws = None
        self.conectado = False
        
        self.crear_interfaz()
    
    def crear_interfaz(self):
        # Header
        header = tk.Frame(self.root, bg='#2c5f2d', height=100)
        header.pack(fill=tk.X)
        header.pack_propagate(False)
        
        tk.Label(header, text="⚡ Asistente Directo", 
                font=('Segoe UI', 16, 'bold'), bg='#2c5f2d', fg='white').pack(pady=10)
        
        self.status_label = tk.Label(header, text="⚠️ Excel no conectado", 
                font=('Segoe UI', 10), bg='#2c5f2d', fg='#ffc107')
        self.status_label.pack()
        
        # Botón conectar
        btn_conectar = tk.Button(header, text="🔌 Conectar a Excel Abierto", 
                                command=self.conectar_excel,
                                bg='#007bff', fg='white', font=('Segoe UI', 10, 'bold'),
                                relief=tk.FLAT, padx=20, pady=8, cursor='hand2')
        btn_conectar.pack(pady=5)
        
        # Frame principal con scroll
        main_frame = tk.Frame(self.root, bg='white')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        canvas = tk.Canvas(main_frame, bg='white')
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        self.scroll_frame = tk.Frame(canvas, bg='white')
        
        self.scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # ===== FORMULARIO COMPACTO =====
        self.campos = {}
        
        # Entidad
        self.crear_campo("Entidad:", "entidad", "entry", 0, default="OTRO")
        self.crear_campo("NIT:", "nit", "entry", 1, default="901145808-5")
        self.crear_campo("ID/Código:", "codigo", "entry", 2)
        
        # Estado Físico Fuste
        self.crear_separador("🌳 Estado Físico Fuste", 3)
        
        fuste_frame = tk.Frame(self.scroll_frame, bg='#f9f9f9', relief=tk.RIDGE, bd=1)
        fuste_frame.grid(row=4, column=0, columnspan=2, sticky='ew', padx=10, pady=5)
        
        self.checks_fuste = {}
        opciones_fuste = [
            ("B", 4), ("Bb", 5), ("BB", 6), ("FR", 7), ("I", 8), ("MI", 9),
            ("To", 10), ("C", 11), ("Rv", 12), ("Ac", 13), ("An", 14), ("Dc", 15),
            ("SB", 16), ("Ag", 17), ("Poe", 18), ("Pe", 19)
        ]
        
        for idx, (nombre, col) in enumerate(opciones_fuste):
            var = tk.IntVar()
            tk.Checkbutton(fuste_frame, text=nombre, variable=var, bg='#f9f9f9',
                          font=('Consolas', 9, 'bold')).grid(row=idx//6, column=idx%6, sticky='w', padx=5, pady=2)
            self.checks_fuste[col] = var
        
        # Estados generales
        self.crear_campo("Estado Fuste General:", "fuste_general", "combo", 5,
                        valores=["", "Bueno", "Regular", "Malo", "Suprimido"])
        
        self.crear_campo("Estado Raíz Específico:", "raiz_especifico", "combo", 6,
                        valores=["", "No apreciable", "Visible", "Superficial", "Profunda"])
        
        self.crear_campo("Estado Raíz General:", "raiz_general", "combo", 7,
                        valores=["", "Bueno", "Regular", "Malo"])
        
        # Estado Sanitario Copa
        self.crear_separador("🍃 Estado Sanitario Copa", 8)
        
        copa_frame = tk.Frame(self.scroll_frame, bg='#f9f9f9', relief=tk.RIDGE, bd=1)
        copa_frame.grid(row=9, column=0, columnspan=2, sticky='ew', padx=10, pady=5)
        
        self.checks_copa = {}
        opciones_copa = [
            ("He", 26), ("An", 27), ("Ag", 28), ("Ne", 29), ("NA", 40)
        ]
        
        for idx, (nombre, col) in enumerate(opciones_copa):
            var = tk.IntVar()
            tk.Checkbutton(copa_frame, text=nombre, variable=var, bg='#f9f9f9',
                          font=('Consolas', 9, 'bold')).grid(row=0, column=idx, sticky='w', padx=8, pady=2)
            self.checks_copa[col] = var
        
        self.crear_campo("Estado Sanitario Copa Específico:", "san_copa_especifico", "combo", 10,
                        valores=["", "Ninguna de las anteriores"])
        
        # Estado Sanitario Fuste
        self.crear_separador("🏥 Estado Sanitario Fuste", 11)
        
        fuste_san_frame = tk.Frame(self.scroll_frame, bg='#f9f9f9', relief=tk.RIDGE, bd=1)
        fuste_san_frame.grid(row=12, column=0, columnspan=2, sticky='ew', padx=10, pady=5)
        
        self.checks_fuste_san = {}
        opciones_fuste_san = [("NA", 47)]
        
        for idx, (nombre, col) in enumerate(opciones_fuste_san):
            var = tk.IntVar()
            tk.Checkbutton(fuste_san_frame, text=nombre, variable=var, bg='#f9f9f9',
                          font=('Consolas', 9, 'bold')).grid(row=0, column=idx, sticky='w', padx=8, pady=2)
            self.checks_fuste_san[col] = var
        
        # Estados sanitarios generales
        self.crear_campo("Estado Sanitario General:", "san_general", "combo", 13,
                        valores=["", "Bueno", "Regular", "Malo"])
        
        self.crear_campo("Estado Sanitario Copa General:", "san_copa_general", "combo", 14,
                        valores=["", "Bueno", "Regular", "Malo"])
        
        self.crear_campo("Estado Sanitario Fuste General:", "san_fuste_general", "combo", 15,
                        valores=["", "Bueno", "Regular", "Malo"])
        
        self.crear_campo("Estado Sanitario Raíz General:", "san_raiz_general", "combo", 16,
                        valores=["", "Bueno", "Regular", "Malo"])
        
        # Causas de Poda
        self.crear_separador("✂️ Causas de Poda", 17)
        
        poda_frame = tk.Frame(self.scroll_frame, bg='#f9f9f9', relief=tk.RIDGE, bd=1)
        poda_frame.grid(row=18, column=0, columnspan=2, sticky='ew', padx=10, pady=5)
        
        self.checks_poda = {}
        opciones_poda = [
            ("Ramas rotas", 61), ("Copa asimétrica", 62), ("Ramas pendulares", 64)
        ]
        
        for idx, (nombre, col) in enumerate(opciones_poda):
            var = tk.IntVar()
            tk.Checkbutton(poda_frame, text=nombre, variable=var, bg='#f9f9f9',
                          font=('Segoe UI', 9)).grid(row=0, column=idx, sticky='w', padx=5, pady=2)
            self.checks_poda[col] = var
        
        self.crear_campo("Tipo de Poda:", "tipo_poda", "combo", 19,
                        valores=["", "De mejoramiento-Estructura", "De mantenimiento", "Especial", "Sanitaria"])
        
        self.crear_campo("Intensidad (%):", "intensidad", "entry", 20)
        self.crear_campo("Residuos (kg):", "residuos", "entry", 21)
        
        # Concepto Técnico
        self.crear_separador("📋 Concepto Técnico", 22)
        
        concepto_frame = tk.Frame(self.scroll_frame, bg='#f9f9f9', relief=tk.RIDGE, bd=1)
        concepto_frame.grid(row=23, column=0, columnspan=2, sticky='ew', padx=10, pady=5)
        
        self.checks_concepto = {}
        opciones_concepto = [("Problemas seguridad", 69)]
        
        for idx, (nombre, col) in enumerate(opciones_concepto):
            var = tk.IntVar()
            tk.Checkbutton(concepto_frame, text=nombre, variable=var, bg='#f9f9f9',
                          font=('Segoe UI', 9)).grid(row=0, column=idx, sticky='w', padx=5, pady=2)
            self.checks_concepto[col] = var
        
        # Establecer valores por defecto
        self.establecer_valores_defecto()
        
        # Botón agregar
        btn_frame = tk.Frame(self.root, bg='#f5f5f5')
        btn_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.btn_agregar = tk.Button(btn_frame, text="➕ Agregar Fila al Excel", 
                                     command=self.agregar_fila,
                                     bg='#28a745', fg='white', font=('Segoe UI', 12, 'bold'),
                                     relief=tk.FLAT, padx=30, pady=12, cursor='hand2',
                                     state='disabled')
        self.btn_agregar.pack(pady=5)
        
        self.info_label = tk.Label(btn_frame, text="Conecta a Excel primero", 
                                   font=('Segoe UI', 9), bg='#f5f5f5', fg='#666')
        self.info_label.pack()
        
        # Configurar columnas
        self.scroll_frame.columnconfigure(1, weight=1)
    
    def establecer_valores_defecto(self):
        """Establece los valores por defecto en los campos"""
        # Estados por defecto
        self.campos['fuste_general'].set("Bueno")
        self.campos['raiz_especifico'].set("No apreciable")
        self.campos['raiz_general'].set("Bueno")
        self.campos['san_copa_especifico'].set("Ninguna de las anteriores")
        self.campos['san_general'].set("Bueno")
        self.campos['san_copa_general'].set("Bueno")
        self.campos['san_fuste_general'].set("Bueno")
        self.campos['san_raiz_general'].set("Bueno")
        self.campos['tipo_poda'].set("De mejoramiento-Estructura")
    
    def crear_campo(self, label, clave, tipo, fila, valores=None, default=""):
        tk.Label(self.scroll_frame, text=label, bg='white', 
                font=('Segoe UI', 9, 'bold')).grid(row=fila, column=0, sticky='w', padx=10, pady=5)
        
        if tipo == "entry":
            widget = tk.Entry(self.scroll_frame, font=('Segoe UI', 9))
            widget.insert(0, default)
            widget.grid(row=fila, column=1, sticky='ew', padx=10, pady=5)
        elif tipo == "combo":
            widget = ttk.Combobox(self.scroll_frame, values=valores or [], 
                                 state='readonly', font=('Segoe UI', 9))
            if default:
                widget.set(default)
            widget.grid(row=fila, column=1, sticky='ew', padx=10, pady=5)
        
        self.campos[clave] = widget
    
    def crear_separador(self, texto, fila):
        frame = tk.Frame(self.scroll_frame, bg='#2c5f2d', height=30)
        frame.grid(row=fila, column=0, columnspan=2, sticky='ew', padx=10, pady=(10, 5))
        frame.grid_propagate(False)
        
        tk.Label(frame, text=texto, font=('Segoe UI', 10, 'bold'), 
                bg='#2c5f2d', fg='white').pack(side=tk.LEFT, padx=10, pady=5)
    
    def conectar_excel(self):
        try:
            # Intentar conectar a Excel abierto
            app = xw.apps.active
            if app is None:
                messagebox.showerror("Error", 
                    "No hay ningún Excel abierto.\n\n"
                    "Por favor:\n"
                    "1. Abre el archivo Excel\n"
                    "2. Ve a la hoja 'BASE DE DATOS'\n"
                    "3. Luego haz clic en Conectar")
                return
            
            # Obtener el workbook activo
            self.wb = app.books.active
            
            # Buscar la hoja BASE DE DATOS
            try:
                self.ws = self.wb.sheets["BASE DE DATOS "]
            except:
                try:
                    self.ws = self.wb.sheets["BASE DE DATOS"]
                except:
                    messagebox.showerror("Error", 
                        "No se encontró la hoja 'BASE DE DATOS'\n\n"
                        "Asegúrate de que la hoja existe en el Excel abierto")
                    return
            
            # Obtener última fila con datos en columna C
            ultima_fila = self.ws.range('C' + str(self.ws.cells.last_cell.row)).end('up').row
            
            # Detectar el último número en columna C y sugerir el siguiente
            try:
                ultimo_codigo = self.ws.range(f'C{ultima_fila}').value
                if ultimo_codigo and str(ultimo_codigo).isdigit():
                    siguiente_codigo = int(ultimo_codigo) + 1
                    self.campos['codigo'].delete(0, tk.END)
                    self.campos['codigo'].insert(0, str(siguiente_codigo))
                    self.campos['codigo'].config(bg='#ffffcc')  # Resaltar
            except:
                pass
            
            self.conectado = True
            self.status_label.config(text=f"✅ Conectado: {self.wb.name} - Fila {ultima_fila}", 
                                    fg='#28a745')
            self.btn_agregar.config(state='normal')
            self.info_label.config(text="Listo para agregar filas", fg='#28a745')
            
            messagebox.showinfo("Conectado", 
                f"Conectado exitosamente a:\n{self.wb.name}\n\n"
                f"Hoja: {self.ws.name}\n"
                f"Última fila: {ultima_fila}\n"
                f"Siguiente código sugerido: {self.campos['codigo'].get()}")
        
        except Exception as e:
            messagebox.showerror("Error", f"Error al conectar:\n{str(e)}")
    
    def agregar_fila(self):
        if not self.conectado:
            messagebox.showwarning("Advertencia", "Conecta a Excel primero")
            return
        
        try:
            codigo = self.campos['codigo'].get()
            
            if not codigo:
                messagebox.showwarning("Advertencia", "Ingresa un ID/Código")
                return
            
            # Encontrar siguiente fila vacía
            ultima_fila = self.ws.range('C' + str(self.ws.cells.last_cell.row)).end('up').row
            nueva_fila = ultima_fila + 1
            
            # Escribir datos básicos DIRECTAMENTE en Excel
            self.ws.range(f'A{nueva_fila}').value = self.campos['entidad'].get()
            self.ws.range(f'B{nueva_fila}').value = self.campos['nit'].get()
            self.ws.range(f'C{nueva_fila}').value = codigo
            
            # Escribir checks de fuste
            for col, var in self.checks_fuste.items():
                if var.get() == 1:
                    self.ws.range((nueva_fila, col)).value = '1'
            
            # Estados generales
            if self.campos['fuste_general'].get():
                self.ws.range((nueva_fila, 23)).value = self.campos['fuste_general'].get()
            if self.campos['raiz_especifico'].get():
                self.ws.range((nueva_fila, 24)).value = self.campos['raiz_especifico'].get()
            if self.campos['raiz_general'].get():
                self.ws.range((nueva_fila, 25)).value = self.campos['raiz_general'].get()
            
            # Checks copa
            for col, var in self.checks_copa.items():
                if var.get() == 1:
                    self.ws.range((nueva_fila, col)).value = '1'
            
            # Estado copa específico
            if self.campos['san_copa_especifico'].get():
                self.ws.range((nueva_fila, 48)).value = self.campos['san_copa_especifico'].get()
            
            # Checks fuste sanitario
            for col, var in self.checks_fuste_san.items():
                if var.get() == 1:
                    self.ws.range((nueva_fila, col)).value = '1'
            
            # Estados sanitarios generales
            if self.campos['san_general'].get():
                self.ws.range((nueva_fila, 49)).value = self.campos['san_general'].get()
            if self.campos['san_copa_general'].get():
                self.ws.range((nueva_fila, 50)).value = self.campos['san_copa_general'].get()
            if self.campos['san_fuste_general'].get():
                self.ws.range((nueva_fila, 51)).value = self.campos['san_fuste_general'].get()
            if self.campos['san_raiz_general'].get():
                self.ws.range((nueva_fila, 52)).value = self.campos['san_raiz_general'].get()
            
            # Checks poda
            for col, var in self.checks_poda.items():
                if var.get() == 1:
                    self.ws.range((nueva_fila, col)).value = '1'
            
            # Tipo e intensidad poda
            if self.campos['tipo_poda'].get():
                self.ws.range((nueva_fila, 66)).value = self.campos['tipo_poda'].get()
            if self.campos['intensidad'].get():
                self.ws.range((nueva_fila, 67)).value = self.campos['intensidad'].get()
            if self.campos['residuos'].get():
                self.ws.range((nueva_fila, 77)).value = self.campos['residuos'].get()
            
            # Checks concepto
            for col, var in self.checks_concepto.items():
                if var.get() == 1:
                    self.ws.range((nueva_fila, col)).value = '1'
            
            # Auto-incrementar código para la siguiente fila
            try:
                codigo_actual = int(codigo)
                siguiente = codigo_actual + 1
                self.campos['codigo'].delete(0, tk.END)
                self.campos['codigo'].insert(0, str(siguiente))
            except:
                self.campos['codigo'].delete(0, tk.END)
            
            self.campos['codigo'].focus()
            
            self.info_label.config(text=f"✅ Fila {nueva_fila} agregada en Excel", fg='#28a745')
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al agregar fila:\n{str(e)}")

def main():
    root = tk.Tk()
    app = AsistenteDirecto(root)
    root.mainloop()

if __name__ == "__main__":
    main()
