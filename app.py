"""Corrugados: Windows desktop product identification, entirely offline."""
from __future__ import annotations
import argparse
import ctypes
import json
import os
import shutil
import subprocess
import sys
import tkinter as tk
import traceback
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
from catalog import Store, FIELDS, LABELS, aliases, normalize_code, read_backup, read_excel

VERSION='1.0.0'
NAVY='#15283F'; BLUE='#245FD3'; BG='#F3F5F9'; INK='#192D44'; MUTED='#52647A'; BORDER='#D5DDE8'; WHITE='#FFFFFF'
ctk.set_appearance_mode('light')
ctk.set_default_color_theme('blue')

def resource(name): return Path(getattr(sys,'_MEIPASS',Path(__file__).parent))/name

def label(parent,text='',size=14,bold=False,color=INK,**kw):
    return ctk.CTkLabel(parent,text=text,font=('Segoe UI',size,'bold' if bold else 'normal'),text_color=color,anchor='w',**kw)

def button(parent,text,command,primary=False,danger=False,**kw):
    bg=BLUE if primary else '#F9FAFC'
    fg=WHITE if primary else ('#AC2232' if danger else INK)
    b=tk.Button(parent,text=text,command=command,font=('Segoe UI',11,'bold' if primary else 'normal'),bg=bg,fg=fg,activebackground='#DDE8FA' if not primary else '#1849AA',activeforeground=fg,relief='flat',bd=0,padx=16,pady=9,cursor='hand2',highlightthickness=2,highlightbackground=bg,highlightcolor='#7C9DDE',**kw)
    b.bind('<Return>',lambda e:command())
    return b

def table(parent,columns,widths,height=8):
    frame=ctk.CTkFrame(parent,fg_color=WHITE,corner_radius=8,border_color=BORDER,border_width=1)
    tree=ttk.Treeview(frame,columns=columns,show='headings',selectmode='browse',height=height)
    for col,width in zip(columns,widths):
        tree.heading(col,text=col);tree.column(col,width=width,minwidth=80,anchor='w',stretch=True)
    scroll=ttk.Scrollbar(frame,orient='vertical',command=tree.yview)
    horizontal=ttk.Scrollbar(frame,orient='horizontal',command=tree.xview)
    tree.configure(yscrollcommand=scroll.set,xscrollcommand=horizontal.set)
    tree.grid(row=0,column=0,sticky='nsew',padx=(1,0),pady=(1,0));scroll.grid(row=0,column=1,sticky='ns')
    horizontal.grid(row=1,column=0,sticky='ew')
    frame.grid_columnconfigure(0,weight=1);frame.grid_rowconfigure(0,weight=1)
    tree.tag_configure('odd',background='#F5F7FB')
    return frame,tree

HELP='''CORRUGADOS · GUÍA RÁPIDA

1. ESCANEAR
Conecta el lector USB en modo teclado y configúralo para enviar Enter al terminar. Abre Escanear (F2), haz clic en el campo y lee la etiqueta. También puedes escribir el código y presionar Enter. El campo queda seleccionado para la siguiente lectura.

La aplicación consulta códigos completos. Conserva ceros iniciales y sólo elimina espacios. No corta dígitos, rellena ceros ni utiliza semejanzas. Un resultado indica lo que dice el catálogo; revisa la etiqueta y las características físicas.

2. INTERPRETAR EL RESULTADO
• Código base: igualdad con NUMERO BASE + NUMERO MOVIL.
• DV en Excel: código base de 12 dígitos seguido del DV registrado.
• ITF14 en Excel: igualdad con el ITF14 registrado, cuando existe.
• DV calculado: posible coincidencia al agregar el verificador GS1 a una base de 12 dígitos sin DV. Se muestra en ámbar; verifica la etiqueta antes de usarla. No se agrega ningún prefijo de caja.
• Código compartido: existen varios registros, incluso entre hojas. Selecciona uno para consultar sus características. No se elige automáticamente.
• No registrado: comprueba la lectura completa. No se adivina el producto.

3. CATÁLOGO
F3 abre el catálogo. Busca por nombre, código, SKU o destino. Selecciona una fila y usa Editar o Eliminar. Nuevo producto (Ctrl+N) permite registrar otro código. Los campos vacíos se muestran como «Sin dato». No se completan características a partir del nombre.

Antes de guardar un código compartido se avisa. Eliminar pide confirmación. En Datos y respaldos puedes recuperar el último eliminado. No se necesita contraseña para editar: asigna el equipo al personal autorizado.

4. COMPARTIR ACTUALIZACIONES
En Datos y respaldos, usa «Exportar catálogo». Envía el archivo .corrugados por USB o correo. En la otra computadora usa «Recibir catálogo». El archivo reemplaza el catálogo local completo, incluyendo las eliminaciones. Los cambios de varias computadoras no se fusionan.

«Importar Excel» admite archivos .xlsx con el encabezado NUMERO BASE + NUMERO MOVIL y conserva los campos de las hojas que lo contienen. Muestra una vista previa antes de reemplazar. La importación sustituye las ediciones locales. Guarda un respaldo si necesitas conservarlas.

5. RESPALDOS Y UBICACIÓN
Los datos se guardan en %LOCALAPPDATA%\\Corrugados\\catalogo.sqlite. No dependen de Excel ni de la ubicación del ejecutable. Antes de cada cambio se guarda un respaldo .corrugados; se conservan los 30 más recientes en la carpeta Respaldos. Recibir un catálogo también crea un respaldo previo. Para restaurar, usa Recibir catálogo y elige uno de esos archivos.

Exportar CSV sirve para consultar en Excel. Los códigos llevan un apóstrofo para conservar ceros y evitar notación científica. El CSV no se utiliza para restaurar; usa .corrugados.

6. INSTALACIÓN
Funciona en Windows 10/11 de 64 bits. Abre Corrugados.exe directamente. En Datos y respaldos puedes elegir «Instalar en este equipo» para copiar el programa a la carpeta de aplicaciones del usuario y crear accesos en el escritorio y menú Inicio, sin permisos de administrador. Puedes retirar la USB después de instalar.

En una instalación nueva, importa tu archivo desde Datos y respaldos → Importar Excel. Cada usuario de Windows tiene datos independientes. Para desinstalar, cierra el programa y elimina su carpeta en %LOCALAPPDATA%\\Programs\\Corrugados y sus accesos directos. La base y sus respaldos se conservan en %LOCALAPPDATA%\\Corrugados.

ATAJOS
F2: escanear    F3: catálogo    Ctrl+N: nuevo producto
Enter: buscar código    Esc: limpiar lectura / cerrar ventana secundaria

Versión 1.0.0 · Sin conexión a internet.
'''

class TextWindow(ctk.CTkToplevel):
    def __init__(self,owner,title,text):
        super().__init__(owner);self.title(title);self.geometry('820x640');self.minsize(600,420);self.transient(owner)
        self.grid_rowconfigure(0,weight=1);self.grid_columnconfigure(0,weight=1)
        box=ctk.CTkTextbox(self,font=('Segoe UI',14),wrap='word');box.grid(row=0,column=0,sticky='nsew',padx=20,pady=20);box.insert('1.0',text);box.configure(state='disabled')
        button(self,'Cerrar',self.destroy).grid(row=1,column=0,sticky='e',padx=20,pady=(0,16))
        self.bind('<Escape>',lambda e:self.destroy());self.after(80,self.focus_force)

class Editor(ctk.CTkToplevel):
    def __init__(self,owner,record=None):
        super().__init__(owner);self.owner=owner;self.record=dict(record or {});self.variables={};self.title('Editar producto' if record else 'Nuevo producto');self.geometry(f'750x{min(740,self.winfo_screenheight()-100)}');self.minsize(650,500);self.transient(owner)
        self.configure(fg_color=BG);self.grid_columnconfigure(0,weight=1);self.grid_rowconfigure(1,weight=1)
        label(self,'Editar producto' if record else 'Nuevo producto',26,True).grid(row=0,column=0,sticky='w',padx=26,pady=(22,10))
        body=ctk.CTkScrollableFrame(self,fg_color=WHITE,corner_radius=10);body.grid(row=1,column=0,sticky='nsew',padx=24,pady=8);body.grid_columnconfigure((0,1),weight=1)
        for ix,key in enumerate(FIELDS):
            var=tk.StringVar(value=self.record.get(key,''));self.variables[key]=var
            rr=(ix//2)*2;cc=ix%2
            label(body,LABELS[key]+(' *' if key in ('code','name') else ''),13,bold=True).grid(row=rr,column=cc,sticky='w',padx=14,pady=(14,4))
            if key=='category':
                if not var.get():var.set('Producto')
                widget=ctk.CTkOptionMenu(body,values=['Producto','Corrugado'],variable=var,fg_color='#E5EDF9',text_color=INK,button_color=BLUE,height=36,font=('Segoe UI',14))
            else:
                widget=ctk.CTkEntry(body,textvariable=var,height=36,font=('Segoe UI',14),border_color=BORDER)
            widget.grid(row=rr+1,column=cc,sticky='ew',padx=14,pady=(0,6))
            if ix==0:self.first_entry=widget
        label(body,'* Código y descripción obligatorios. Los códigos se guardan como texto.',12,color=MUTED,wraplength=600).grid(row=12,column=0,columnspan=2,sticky='w',padx=14,pady=12)
        self.error_var=tk.StringVar();label(self,textvariable=self.error_var,color='#AF2433',wraplength=690).grid(row=2,column=0,sticky='ew',padx=26,pady=6)
        footer=ctk.CTkFrame(self,fg_color='transparent');footer.grid(row=3,column=0,sticky='ew',padx=24,pady=(4,20));footer.grid_columnconfigure(0,weight=1)
        button(footer,'Cancelar',self.destroy).grid(row=0,column=1,padx=8);button(footer,'Guardar producto',self.commit,primary=True).grid(row=0,column=2)
        self.bind('<Escape>',lambda e:self.destroy());self.bind('<Control-s>',lambda e:self.commit())
        self.grab_set();self.after(80,self.first_entry.focus_set)

    def commit(self):
        r=dict(self.record);r.update({k:v.get() for k,v in self.variables.items()})
        try:
            r['code']=normalize_code(r['code'])
            if not r['name'].strip():raise ValueError('Escribe la descripción del producto.')
            from catalog import validate_record
            r=validate_record(r)
            overlaps={m['record']['id']:m['record'] for code,kind in aliases(r) for m in self.owner.store.lookup(code) if m['record']['id']!=r['id']}
            if overlaps and not messagebox.askyesno('Código compartido',f'Uno de los códigos también identifica {len(overlaps)} registro(s). Las lecturas mostrarán varias coincidencias.\n\n¿Guardar así?',parent=self):return
            if self.record:r['edited_manually']=True
            else:r.update(source_sheet='Alta manual',source_row='')
            self.owner.store.save(r);self.owner.catalog_changed('Producto guardado.');self.destroy()
        except (ValueError,OSError) as e:self.error_var.set(str(e))

class App(ctk.CTk):
    def __init__(self,data_dir=None):
        super().__init__();self.title('Corrugados · Consulta de productos');self.geometry(f'{min(1220,self.winfo_screenwidth()-40)}x{min(820,self.winfo_screenheight()-100)}');self.minsize(1000,660);self.configure(fg_color=BG)
        self.data_dir=Path(data_dir or Path(os.environ.get('LOCALAPPDATA',Path.home()))/'Corrugados');self.data_dir.mkdir(parents=True,exist_ok=True)
        self.store=Store(self.data_dir/'catalogo.sqlite')
        self.current_record=None;self.view='scan';self.last_matches={};self._filter_job=None
        self.protocol('WM_DELETE_WINDOW',self.shutdown)
        try:self.iconbitmap(str(resource('corrugados.ico')))
        except tk.TclError:pass
        self.grid_columnconfigure(1,weight=1);self.grid_rowconfigure(0,weight=1)
        side=ctk.CTkFrame(self,fg_color=NAVY,corner_radius=0,width=182);side.grid(row=0,column=0,sticky='nsew');side.grid_propagate(False);side.grid_columnconfigure(0,weight=1);side.grid_rowconfigure(8,weight=1)
        mark=tk.Canvas(side,width=48,height=42,bg=NAVY,highlightthickness=0)
        for x,w in [(3,2),(8,4),(16,2),(21,1),(26,4),(34,2),(40,3)]:mark.create_rectangle(x,6,x+w,34,fill='#AFC8FF',outline='')
        mark.grid(row=0,column=0,sticky='w',padx=24,pady=(30,8))
        label(side,'CORRUGADOS',17,True,WHITE).grid(row=1,column=0,sticky='w',padx=22)
        label(side,'Consulta de productos',12,color='#BCCBDD').grid(row=2,column=0,sticky='w',padx=22,pady=(0,28))
        self.nav={}
        for row,(text,cmd) in enumerate([('Escanear   F2',self.show_scan),('Catálogo   F3',self.show_catalog),('Datos y respaldos',self.show_data)],3):
            b=button(side,text,cmd);b.configure(bg=NAVY,fg='#E7EEF9',activebackground='#284363',activeforeground=WHITE,highlightbackground=NAVY,anchor='w',padx=12);b.grid(row=row,column=0,sticky='ew',padx=12,pady=4);self.nav[text]=b
        button(side,'Ayuda',lambda:TextWindow(self,'Guía de uso',HELP)).grid(row=9,column=0,sticky='ew',padx=18,pady=10)
        label(side,'Sin conexión · v'+VERSION,11,color='#BCCBDD').grid(row=10,column=0,sticky='w',padx=20,pady=(0,20))
        self.content=ctk.CTkFrame(self,fg_color='transparent',corner_radius=0);self.content.grid(row=0,column=1,sticky='nsew',padx=28,pady=24);self.content.grid_rowconfigure(0,weight=1);self.content.grid_columnconfigure(0,weight=1)
        self.pages={}
        self.setup_styles();self.build_scan();self.build_catalog();self.build_data();self.show_scan()
        self.bind('<F2>',lambda e:self.show_scan());self.bind('<F3>',lambda e:self.show_catalog());self.bind('<Control-n>',lambda e:Editor(self));self.bind('<Escape>',lambda e:self.clear_scan() if self.view=='scan' else None)
        self.after(200,self.scan_entry.focus_set)

    def setup_styles(self):
        style=ttk.Style(self);style.theme_use('clam')
        style.configure('Treeview',background=WHITE,foreground=INK,fieldbackground=WHITE,rowheight=38,font=('Segoe UI',10),borderwidth=0)
        style.configure('Treeview.Heading',background='#E9EEF6',foreground=MUTED,font=('Segoe UI',10,'bold'),padding=(10,10),relief='flat')
        style.map('Treeview',background=[('selected','#DCE8FF')],foreground=[('selected','#163C80')])

    def page(self,name):
        p=ctk.CTkFrame(self.content,fg_color='transparent');p.grid(row=0,column=0,sticky='nsew');p.grid_columnconfigure(0,weight=1);self.pages[name]=p;return p

    def switch(self,name):
        self.view=name;self.pages[name].tkraise()
        for text,b in self.nav.items():
            active=(name=='scan' and text.startswith('Escanear')) or (name=='catalog' and text.startswith('Catálogo')) or (name=='data' and text.startswith('Datos'))
            b.configure(bg='#2B4565' if active else NAVY,highlightbackground='#2B4565' if active else NAVY)

    def build_scan(self):
        p=self.page('scan');p.grid_rowconfigure(5,weight=1)
        label(p,'Escanear producto',30,True).grid(row=0,column=0,sticky='w')
        label(p,'Lee la etiqueta y revisa que la presentación sea la que necesitas.',14,color=MUTED).grid(row=1,column=0,sticky='w',pady=(4,20))
        entrycard=ctk.CTkFrame(p,fg_color=WHITE,corner_radius=12,border_color=BORDER,border_width=1);entrycard.grid(row=2,column=0,sticky='ew');entrycard.grid_columnconfigure(0,weight=1)
        label(entrycard,'CÓDIGO DE BARRAS',12,True,color=MUTED).grid(row=0,column=0,sticky='w',padx=20,pady=(16,7))
        self.scan_var=tk.StringVar();self.scan_entry=ctk.CTkEntry(entrycard,textvariable=self.scan_var,height=54,font=('Consolas',25),placeholder_text='Escanea o escribe el código completo',border_width=2,border_color='#A8BEE4',fg_color='#FAFCFF')
        self.scan_entry.grid(row=1,column=0,sticky='ew',padx=(20,12),pady=(0,8));self.scan_entry.bind('<Return>',self.perform_scan);self.scan_entry.bind('<KP_Enter>',self.perform_scan)
        button(entrycard,'Buscar',self.perform_scan,primary=True).grid(row=1,column=1,sticky='ns',padx=(0,20),pady=(0,8))
        label(entrycard,'Enter para buscar · Códigos completos de longitud variable · F2 para volver aquí',12,color=MUTED).grid(row=2,column=0,columnspan=2,sticky='w',padx=20,pady=(0,14))
        self.scan_var.trace_add('write',self.scan_typing)
        self.status_frame=ctk.CTkFrame(p,fg_color='#E7EDF6',corner_radius=8);self.status_frame.grid(row=3,column=0,sticky='ew',pady=(14,12));self.status_frame.grid_columnconfigure(0,weight=1)
        self.status_var=tk.StringVar(value='Listo para escanear');self.status_label=label(self.status_frame,textvariable=self.status_var,size=16,bold=True,wraplength=840);self.status_label.grid(row=0,column=0,sticky='ew',padx=18,pady=14)
        self.status_frame.bind('<Configure>',lambda e:self.status_label.configure(wraplength=max(360,e.width-40)))
        self.results_frame,self.result_tree=table(p,['Tipo','Producto / presentación','Origen de coincidencia'],[110,470,190],height=3)
        self.results_frame.grid(row=4,column=0,sticky='ew',pady=(0,12));self.results_frame.grid_remove()
        self.result_tree.bind('<<TreeviewSelect>>',self.select_match)
        self.scan_detail=ctk.CTkScrollableFrame(p,fg_color=WHITE,corner_radius=12,border_width=1,border_color=BORDER);self.scan_detail.grid(row=5,column=0,sticky='nsew');self.scan_detail.grid_columnconfigure(0,weight=1)
        self.render_detail(self.scan_detail,None)
        footer=ctk.CTkFrame(p,fg_color='transparent');footer.grid(row=6,column=0,sticky='ew',pady=(12,0));footer.grid_columnconfigure(0,weight=1)
        self.count_var=tk.StringVar();label(footer,textvariable=self.count_var,size=12,color=MUTED).grid(row=0,column=0,sticky='w')
        button(footer,'Nueva lectura',self.clear_scan).grid(row=0,column=1);self.update_count()

    def render_detail(self,frame,record,compact=False):
        for widget in frame.winfo_children():widget.destroy()
        if record is None:
            label(frame,'Selecciona un registro' if compact else 'La ficha del producto aparecerá aquí',20 if compact else 22,True,wraplength=210 if compact else 650,justify='left').grid(row=0,column=0,sticky='w',padx=24,pady=(30,10))
            label(frame,'Verás aquí su código y características.' if compact else 'Escanea el código completo. Si hay varias coincidencias, selecciona una para consultar sus características.',14,color=MUTED,wraplength=210 if compact else 600,justify='left').grid(row=1,column=0,sticky='w',padx=24,pady=(0,30));return
        width=210 if compact else min(780,max(430,self.winfo_width()-310))
        label(frame,record['category'].upper(),12,True,color=BLUE).grid(row=0,column=0,sticky='w',padx=22,pady=(18,8))
        label(frame,record['name'] or 'Sin descripción en el catálogo',21 if compact else 25,True,wraplength=width,justify='left').grid(row=1,column=0,sticky='w',padx=22,pady=(0,10))
        label(frame,record['code'],19,True,color=BLUE,wraplength=width).grid(row=2,column=0,sticky='w',padx=22,pady=(0,14))
        for row,key in enumerate(('sku','destination','capacity','alcohol','family','dv','itf14','notes'),3):
            line=ctk.CTkFrame(frame,fg_color='transparent');line.grid(row=row,column=0,sticky='ew',padx=22,pady=(0,10));line.grid_columnconfigure(1,weight=1)
            label(line,LABELS[key],12,bold=True,color=MUTED,width=0 if compact else 168,wraplength=210 if compact else 165).grid(row=0,column=0,sticky='nw',padx=(0,12))
            label(line,record.get(key) or 'Sin dato',14,wraplength=210 if compact else min(570,width-185),justify='left').grid(row=1 if compact else 0,column=0 if compact else 1,sticky='w')
        row=12
        for k,v in record.get('extras',{}).items():
            if v:label(frame,f'{k}: {v}',13,wraplength=width,justify='left').grid(row=row,column=0,sticky='w',padx=22,pady=4);row+=1
        origin=record.get('source_sheet','Alta manual')
        if record.get('source_row'):origin+=f' · fila {record["source_row"]}'
        if record.get('edited_manually'):origin+=' · editado en la aplicación'
        label(frame,origin,12,color=MUTED,wraplength=width).grid(row=row,column=0,sticky='w',padx=22,pady=(12,18))

    def set_status(self,text,mode='neutral'):
        colors={'neutral':('#E7EDF6',INK),'ok':('#E2F1E9','#175B38'),'warn':('#FFF0D5','#744800'),'error':('#FCE8EA','#A72434')}
        bg,fg=colors[mode];self.status_var.set(text);self.status_frame.configure(fg_color=bg);self.status_label.configure(text_color=fg)

    def scan_typing(self,*_):
        had_result=self.current_record is not None or bool(self.last_matches)
        self.current_record=None;self.last_matches={};self.results_frame.grid_remove()
        if had_result:self.render_detail(self.scan_detail,None)
        self.set_status('Capturando código… presiona Enter al terminar.' if self.scan_var.get() else 'Listo para escanear')

    def perform_scan(self,event=None):
        self.current_record=None;self.last_matches={};self.results_frame.grid_remove();self.result_tree.delete(*self.result_tree.get_children());self.render_detail(self.scan_detail,None)
        try:matches=self.store.lookup(self.scan_var.get())
        except ValueError as e:self.set_status(str(e),'error');self.scan_entry.focus_set();self.scan_entry.select_range(0,'end');return 'break'
        if not matches:self.set_status('Código no registrado. Comprueba la lectura o consulta el catálogo.','error');self.bell()
        elif len(matches)>1:
            self.set_status(f'Código compartido: {len(matches)} registros. Selecciona y verifica la presentación.','warn')
            for i,m in enumerate(matches):
                r=m['record'];self.last_matches[r['id']]=m;self.result_tree.insert('','end',iid=r['id'],values=(r['category'],r['name'] or 'Sin descripción',m['kind']),tags=('odd',) if i%2 else ())
            self.results_frame.grid()
        else:
            match=matches[0];self.current_record=match['record'];self.render_detail(self.scan_detail,self.current_record)
            derived=match['kind']=='DV calculado'
            self.set_status('Posible coincidencia · DV calculado. Verifica el código impreso.' if derived else 'Producto encontrado · '+match['kind']+'. Revisa sus características.','warn' if derived else 'ok')
        self.scan_entry.focus_set();self.scan_entry.select_range(0,'end');return 'break'

    def select_match(self,event=None):
        selected=self.result_tree.selection()
        if selected and selected[0] in self.last_matches:
            m=self.last_matches[selected[0]];self.current_record=m['record'];self.render_detail(self.scan_detail,self.current_record)
            self.set_status(f'Código compartido · ficha seleccionada: {m["kind"]}. Verifica la presentación.','warn')
            self.scan_entry.focus_set();self.scan_entry.select_range(0,'end')

    def clear_scan(self):
        self.scan_var.set('');self.scan_entry.focus_set()

    def show_scan(self):self.switch('scan');self.scan_entry.focus_set();self.scan_entry.select_range(0,'end')

    def build_catalog(self):
        p=self.page('catalog');p.grid_rowconfigure(3,weight=1)
        head=ctk.CTkFrame(p,fg_color='transparent');head.grid(row=0,column=0,sticky='ew');head.grid_columnconfigure(0,weight=1)
        label(head,'Catálogo',30,True).grid(row=0,column=0,sticky='w');button(head,'Nuevo producto',lambda:Editor(self),primary=True).grid(row=0,column=1)
        label(p,'Consulta y administra los códigos de productos y corrugados.',14,color=MUTED).grid(row=1,column=0,sticky='w',pady=(4,18))
        filters=ctk.CTkFrame(p,fg_color='transparent');filters.grid(row=2,column=0,sticky='ew',pady=(0,14));filters.grid_columnconfigure(0,weight=1)
        self.filter_var=tk.StringVar();self.filter_entry=ctk.CTkEntry(filters,textvariable=self.filter_var,placeholder_text='Buscar por producto, código, SKU o destino…',height=42,font=('Segoe UI',14),border_color=BORDER)
        self.filter_entry.grid(row=0,column=0,sticky='ew',padx=(0,12));self.filter_var.trace_add('write',self.debounce_filter)
        self.category_var=tk.StringVar(value='Todos');ctk.CTkOptionMenu(filters,values=['Todos','Producto','Corrugado'],variable=self.category_var,command=lambda _:self.refresh_catalog(),height=42,fg_color=BLUE,font=('Segoe UI',14)).grid(row=0,column=1)
        split=ctk.CTkFrame(p,fg_color='transparent');split.grid(row=3,column=0,sticky='nsew');split.grid_columnconfigure(0,weight=3);split.grid_columnconfigure(1,weight=2);split.grid_rowconfigure(0,weight=1)
        listing,self.catalog_tree=table(split,['Código base','Producto / descripción','Tipo'],[150,320,100]);listing.grid(row=0,column=0,sticky='nsew',padx=(0,14))
        self.catalog_tree.bind('<<TreeviewSelect>>',self.select_catalog);self.catalog_tree.bind('<Double-1>',lambda e:self.edit_selected())
        right=ctk.CTkFrame(split,fg_color='transparent');right.grid(row=0,column=1,sticky='nsew');right.grid_rowconfigure(0,weight=1);right.grid_columnconfigure(0,weight=1)
        self.catalog_detail=ctk.CTkScrollableFrame(right,fg_color=WHITE,border_color=BORDER,border_width=1);self.catalog_detail.grid(row=0,column=0,sticky='nsew');self.catalog_detail.grid_columnconfigure(0,weight=1)
        label(self.catalog_detail,'Selecciona un registro para ver sus características.',17,True,wraplength=300,justify='left').pack(padx=22,pady=30,anchor='w')
        actions=ctk.CTkFrame(right,fg_color='transparent');actions.grid(row=1,column=0,sticky='ew',pady=(12,0));actions.grid_columnconfigure((0,1),weight=1)
        self.edit_button=button(actions,'Editar',self.edit_selected,primary=True,state='disabled');self.edit_button.grid(row=0,column=0,sticky='ew',padx=(0,8))
        self.delete_button=button(actions,'Eliminar',self.delete_selected,danger=True,state='disabled');self.delete_button.grid(row=0,column=1,sticky='ew')
        self.catalog_count=tk.StringVar();label(p,textvariable=self.catalog_count,size=12,color=MUTED).grid(row=4,column=0,sticky='w',pady=(14,0))

    def debounce_filter(self,*_):
        if self._filter_job:self.after_cancel(self._filter_job)
        self._filter_job=self.after(150,self.refresh_catalog)

    def refresh_catalog(self):
        self._filter_job=None
        selected=self.catalog_tree.selection();self.catalog_tree.delete(*self.catalog_tree.get_children())
        rows=self.store.list(self.filter_var.get(),self.category_var.get())
        for ix,r in enumerate(rows):self.catalog_tree.insert('','end',iid=r['id'],values=(r['code'],r['name'] or 'Sin descripción',r['category']),tags=('odd',) if ix%2 else ())
        self.catalog_count.set(f'{len(rows)} de {len(self.store.rows)} registros · Doble clic para editar')
        if selected and self.catalog_tree.exists(selected[0]):self.catalog_tree.selection_set(selected[0])
        else:
            self.edit_button.configure(state='disabled');self.delete_button.configure(state='disabled')
            self.render_detail(self.catalog_detail,None,True)

    def selected_record(self):
        sel=self.catalog_tree.selection()
        return next((r for r in self.store.rows if sel and r['id']==sel[0]),None)

    def select_catalog(self,event=None):
        r=self.selected_record();self.render_detail(self.catalog_detail,r,True)
        self.edit_button.configure(state='normal' if r else 'disabled');self.delete_button.configure(state='normal' if r else 'disabled')

    def edit_selected(self):
        r=self.selected_record()
        if r:Editor(self,r)

    def delete_selected(self):
        r=self.selected_record()
        if r and messagebox.askyesno('Eliminar registro',f'{r["code"]}\n{r["name"]}\n\n¿Eliminar este registro del catálogo?\nPuedes recuperarlo en Datos y respaldos.',parent=self):
            self.store.delete(r['id']);self.catalog_changed('Registro eliminado. Puedes recuperarlo en Datos y respaldos.')

    def show_catalog(self):self.switch('catalog');self.refresh_catalog();self.filter_entry.focus_set()

    def update_count(self):self.count_var.set(f'{len(self.store.rows)} registros en este equipo · Datos locales')

    def catalog_changed(self,message):
        self.update_count();self.refresh_catalog();self.clear_scan();self.data_notice.set(message)
        if hasattr(self,'data_summary'):self.data_summary.set(f'{len(self.store.rows)} registros activos · Cada computadora conserva su propio catálogo.')

    def build_data(self):
        p=self.page('data');p.grid_rowconfigure(2,weight=1)
        label(p,'Datos y respaldos',30,True).grid(row=0,column=0,sticky='w')
        self.data_summary=tk.StringVar(value=f'{len(self.store.rows)} registros activos · Cada computadora conserva su propio catálogo.')
        label(p,textvariable=self.data_summary,size=14,color=MUTED).grid(row=1,column=0,sticky='w',pady=(4,20))
        body=ctk.CTkScrollableFrame(p,fg_color=WHITE,corner_radius=12,border_color=BORDER,border_width=1);body.grid(row=2,column=0,sticky='nsew');body.grid_columnconfigure(0,weight=1)
        sections=[('Compartir el catálogo','Envía un archivo .corrugados por USB o correo y recíbelo en la otra computadora.',[('Exportar catálogo',self.export_backup),('Recibir catálogo',self.receive_backup)]),('Actualizar desde Excel','Revisa una vista previa antes de reemplazar el catálogo local con las hojas del archivo .xlsx.',[('Importar Excel',self.import_excel),('Exportar CSV',self.export_csv)]),('Recuperar y revisar','Se conservan los 30 respaldos automáticos más recientes. Puedes revisar los códigos compartidos.',[('Abrir respaldos',self.open_backups),('Recuperar eliminado',self.undo_delete),('Revisar catálogo',lambda:TextWindow(self,'Revisión del catálogo',self.store.report()))]),('Instalar en este equipo','Copia el programa al equipo y crea accesos directos para usarlo sin la USB.',[('Instalar en este equipo',self.install),('Abrir carpeta de datos',lambda:os.startfile(self.data_dir))])]
        for ix,(title,desc,actions) in enumerate(sections):
            box=ctk.CTkFrame(body,fg_color='transparent');box.grid(row=ix,column=0,sticky='ew',padx=22,pady=(22,8));box.grid_columnconfigure(0,weight=1)
            label(box,title,20,True).grid(row=0,column=0,sticky='w');label(box,desc,14,color=MUTED,wraplength=650,justify='left').grid(row=1,column=0,sticky='w',pady=(5,12))
            bar=ctk.CTkFrame(box,fg_color='transparent');bar.grid(row=2,column=0,sticky='w')
            for cc,(text,cmd) in enumerate(actions):button(bar,text,cmd,primary=ix==0 and cc==0).grid(row=0,column=cc,padx=(0,10))
        self.data_notice=tk.StringVar(value='Los respaldos se crean automáticamente antes de guardar, eliminar o reemplazar datos.')
        label(p,textvariable=self.data_notice,size=13,color=MUTED,wraplength=850,justify='left').grid(row=3,column=0,sticky='ew',pady=(16,0))

    def show_data(self):self.switch('data')

    def export_backup(self):
        path=filedialog.asksaveasfilename(parent=self,title='Exportar catálogo para otra computadora',defaultextension='.corrugados',initialfile='Catalogo-Corrugados.corrugados',filetypes=[('Catálogo Corrugados','*.corrugados')])
        if path:self.store.export_backup(path);self.data_notice.set('Catálogo exportado. En la otra computadora usa Recibir catálogo.')

    def receive_backup(self):
        path=filedialog.askopenfilename(parent=self,title='Recibir o restaurar catálogo',filetypes=[('Catálogo Corrugados','*.corrugados')])
        if path:
            try:self.preview_import(read_backup(path),[],Path(path).name)
            except (ValueError,OSError) as e:messagebox.showerror('No se pudo leer el catálogo',str(e),parent=self)

    def import_excel(self):
        path=filedialog.askopenfilename(parent=self,title='Importar base de datos Excel',filetypes=[('Libro Excel','*.xlsx')])
        if not path:return
        self.configure(cursor='watch');self.update_idletasks()
        try:
            rows,issues=read_excel(path);self.preview_import(rows,issues,Path(path).name)
        except Exception as e:messagebox.showerror('No se pudo importar el Excel',str(e),parent=self)
        finally:self.configure(cursor='')

    def preview_import(self,rows,issues,name):
        win=ctk.CTkToplevel(self);win.title('Revisar importación');win.geometry('780x600');win.minsize(600,450);win.transient(self);win.grab_set();win.grid_columnconfigure(0,weight=1);win.grid_rowconfigure(2,weight=1)
        label(win,'Revisar antes de reemplazar',25,True).grid(row=0,column=0,sticky='w',padx=24,pady=(24,12))
        summary=f'{name}\n{len(rows)} registros para importar · {len(issues)} observaciones\nSe reemplazarán los {len(self.store.rows)} registros locales. Se guardará un respaldo previo.'
        label(win,summary,14,wraplength=700,justify='left').grid(row=1,column=0,sticky='w',padx=24,pady=(0,12))
        box=ctk.CTkTextbox(win,font=('Segoe UI',13),wrap='word');box.grid(row=2,column=0,sticky='nsew',padx=24,pady=8)
        box.insert('1.0','\n'.join(issues) if issues else 'Sin observaciones de importación.\n\nSe conservarán los registros con códigos compartidos. La búsqueda mostrará todas sus coincidencias.');box.configure(state='disabled')
        bar=ctk.CTkFrame(win,fg_color='transparent');bar.grid(row=3,column=0,sticky='e',padx=24,pady=20)
        def apply():
            self.store.replace(rows);self.catalog_changed(f'Catálogo actualizado: {len(rows)} registros. Respaldo anterior disponible.');win.destroy()
        button(bar,'Cancelar',win.destroy).grid(row=0,column=0,padx=10);button(bar,'Reemplazar catálogo',apply,primary=True).grid(row=0,column=1)
        win.bind('<Escape>',lambda e:win.destroy())

    def export_csv(self):
        path=filedialog.asksaveasfilename(parent=self,title='Exportar consulta a CSV',defaultextension='.csv',initialfile='Catalogo-Corrugados.csv',filetypes=[('CSV para Excel','*.csv')])
        if path:self.store.export_csv(path);self.data_notice.set('CSV exportado. Para transferir o restaurar usa Exportar catálogo (.corrugados).')

    def open_backups(self):
        folder=self.data_dir/'Respaldos';folder.mkdir(exist_ok=True);os.startfile(folder)

    def undo_delete(self):
        ok=self.store.undo_delete()
        self.catalog_changed('Se recuperó el último registro eliminado.' if ok else 'No hay registros eliminados para recuperar.')

    def install(self):
        if not getattr(sys,'frozen',False):messagebox.showinfo('Instalación','Esta opción está disponible en Corrugados.exe.',parent=self);return
        install_dir=Path(os.environ['LOCALAPPDATA'])/'Programs'/'Corrugados';install_dir.mkdir(parents=True,exist_ok=True);target=install_dir/'Corrugados.exe'
        if Path(sys.executable).resolve()!=target.resolve():shutil.copy2(sys.executable,target)
        quoted=str(target).replace("'","''");working=str(install_dir).replace("'","''")
        script=f"$w=New-Object -ComObject WScript.Shell; $dirs=@($w.SpecialFolders.Item('Desktop'),$w.SpecialFolders.Item('Programs')); foreach($d in $dirs){{$s=$w.CreateShortcut((Join-Path $d 'Corrugados.lnk'));$s.TargetPath='{quoted}';$s.WorkingDirectory='{working}';$s.Description='Consulta de productos y corrugados';$s.Save()}}"
        result=subprocess.run(['powershell.exe','-NoProfile','-NonInteractive','-Command',script],capture_output=True,text=True,creationflags=subprocess.CREATE_NO_WINDOW)
        if result.returncode:raise OSError('Se copió el programa, pero no se pudieron crear los accesos directos: '+result.stderr)
        self.data_notice.set('Instalado. Abre Corrugados desde el escritorio o el menú Inicio. Puedes retirar la USB al cerrar esta ventana.')
        messagebox.showinfo('Instalación terminada','Corrugados ya está instalado para este usuario.\n\nSe crearon accesos en el escritorio y en el menú Inicio.\nCierra esta ventana antes de retirar la USB.',parent=self)

    def report_callback_exception(self,exc,value,tb):
        log=self.data_dir/'errores.log'
        with log.open('a',encoding='utf-8') as f:traceback.print_exception(exc,value,tb,file=f)
        messagebox.showerror('No se pudo completar la acción',str(value)+'\n\nConsulta errores.log en la carpeta de datos para revisar el detalle.',parent=self)

    def shutdown(self):
        if getattr(self,'_filter_job',None):
            try:self.after_cancel(self._filter_job)
            except tk.TclError:pass
        if hasattr(self,'store'):self.store.close()
        # Cancel the interpreter's scheduled UI jobs before destroying its widgets.
        try:
            for job in self.tk.call('after','info'):self.after_cancel(job)
        except tk.TclError:pass
        try:self.destroy()
        except tk.TclError:pass

def main():
    parser=argparse.ArgumentParser();parser.add_argument('--data-dir',type=Path);args=parser.parse_args()
    data=args.data_dir or Path(os.environ.get('LOCALAPPDATA',Path.home()))/'Corrugados'
    # One writer per local catalog. Different Windows users have independent catalogs.
    from hashlib import sha256
    mutex_name='Local\\Corrugados-'+sha256(str(data.resolve()).casefold().encode()).hexdigest()[:24]
    kernel=ctypes.WinDLL('kernel32',use_last_error=True);kernel.CreateMutexW.restype=ctypes.c_void_p;kernel.CreateMutexW.argtypes=[ctypes.c_void_p,ctypes.c_int,ctypes.c_wchar_p]
    handle=kernel.CreateMutexW(None,False,mutex_name)
    if ctypes.get_last_error()==183:
        ctypes.windll.user32.MessageBoxW(None,'Corrugados ya está abierto para este catálogo. Usa la ventana existente.','Corrugados',0x40);return
    try:
        ui=App(data);ui.mainloop()
    except Exception:
        data.mkdir(parents=True,exist_ok=True);(data/'errores-inicio.log').write_text(traceback.format_exc(),encoding='utf-8')
        ctypes.windll.user32.MessageBoxW(None,'No se pudo iniciar Corrugados. Revisa errores-inicio.log en la carpeta de datos.','Corrugados',0x10)
    finally:
        if handle:
            kernel.CloseHandle.argtypes=[ctypes.c_void_p];kernel.CloseHandle(handle)

if __name__=='__main__':main()
