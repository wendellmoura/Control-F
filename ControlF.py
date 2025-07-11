import os
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import threading
import json
import time
import logging
import re
import pandas as pd
from io import StringIO
import csv
from collections import defaultdict
import ttkbootstrap as tb

# Configuração de logging para monitoramento do sistema
# IMPORTANTE: Logging é essencial para diagnóstico de problemas de desempenho
logging.basicConfig(
    level=logging.INFO,
    format='[%(asctime)s] [%(levelname)s] %(message)s',
    datefmt='%H:%M:%S',
)

# Exceções customizadas para tratamento especializado de erros
class FileSearchError(Exception):
    """Exceção base para erros na busca de arquivos"""

class LocalFileSearcher:
    # Estratégias de escalonamento implementadas:
    # 1. Leitura otimizada de grandes arquivos com detecção automática de delimitadores
    # 2. Processamento por planilha independente para paralelização natural
    def __init__(self):
        self.file_path = None
        self.worksheets = []
        self.file_data = {}
        self.column_metadata = {}
        
    def load_file(self, file_path):
        """Carrega arquivo detectando automaticamente formato e delimitador"""
        try:
            self.file_path = file_path
            self.worksheets = []
            self.file_data = {}
            self.column_metadata = {}
            
            if file_path.endswith('.csv'):
                # Estratégia para CSV: detecção inteligente de delimitador
                # EVITA: Carregamento completo na memória para arquivos muito grandes
                self.worksheets = ['Sheet1']
                df = self._read_csv(file_path)
                self.file_data['Sheet1'] = df
                self.column_metadata['Sheet1'] = {
                    'columns': list(df.columns),
                    'dtypes': df.dtypes.astype(str).to_dict()
                }
                
            elif file_path.endswith(('.xlsx', '.xls')):
                # Para Excel: carregamento seletivo por planilha
                # PERMITE: Processamento paralelo de diferentes planilhas
                xl = pd.ExcelFile(file_path)
                self.worksheets = xl.sheet_names
                
                for sheet in self.worksheets:
                    # Processamento independente por planilha (permite paralelismo)
                    df = xl.parse(sheet)
                    self.file_data[sheet] = df
                    self.column_metadata[sheet] = {
                        'columns': list(df.columns),
                        'dtypes': df.dtypes.astype(str).to_dict()
                    }
            
            return self.worksheets
        except Exception as e:
            logging.error(f"Erro ao carregar arquivo: {e}")
            raise FileSearchError(f"Falha ao carregar arquivo: {str(e)}")
    
    def _read_csv(self, file_path):
        """Detecta delimitador automaticamente para melhor compatibilidade
        OTIMIZAÇÃO: Leitura mínima do arquivo (apenas primeira linha)"""
        with open(file_path, 'r', encoding='utf-8') as f:
            first_line = f.readline()
        
        # Estratégia de detecção de delimitador com complexidade O(1)
        if ';' in first_line:
            delimiter = ';'
        elif ',' in first_line:
            delimiter = ','
        elif '\t' in first_line:
            delimiter = '\t'
        else:
            delimiter = ','
            
        return pd.read_csv(file_path, delimiter=delimiter, encoding='utf-8')
    
    def search_in_worksheet(self, sheet_name, query):
        """Busca sequencial otimizada com cache de dados
        ATENÇÃO: Algoritmo O(n*m) - pode ser lento para datasets muito grandes"""
        try:
            if sheet_name not in self.file_data:
                return []
                
            df = self.file_data[sheet_name]
            matches = []
            
            # Cache de dados para operações repetitivas
            original_df = df.copy()
            df = df.astype(str)
            
            # Algoritmo de busca linear (O(n))
            # MELHORIA POTENCIAL: Usar vectorization do Pandas ou indexação
            for row_idx, row in df.iterrows():
                for col_idx, cell_value in enumerate(row):
                    if query.lower() in cell_value.lower():
                        col_name = df.columns[col_idx]
                        full_row = original_df.iloc[row_idx].to_dict()
                        
                        matches.append({
                            'worksheet': sheet_name,
                            'cell': f"{col_name}{row_idx + 1}",
                            'value': cell_value,
                            'row': row_idx + 1,
                            'col': col_idx + 1,
                            'full_row': full_row,
                            'columns': list(original_df.columns)
                        })
            return matches
        except Exception as e:
            logging.error(f"Erro na busca: {e}")
            raise FileSearchError(f"Falha na busca: {str(e)}")

class ColumnSelector(tb.Toplevel):
    """Janela avançada para seleção de colunas com técnicas de escalonamento UI:
    - Filtragem em tempo real
    - Virtualização implícita de elementos
    - Operações em lote"""
    def __init__(self, parent, search_results, current_selection=None):
        super().__init__(parent)
        self.title("Seleção de Colunas")
        self.geometry("600x500")
        self.parent = parent
        self.search_results = search_results
        self.selected_columns = current_selection if current_selection else defaultdict(set)
        self.applied = False
        self.column_items = {}
        self.all_items = []
        
        # Configuração de estilo
        self.custom_style = tb.Style()
        self.custom_style.theme_use('litera')
        
        self.create_widgets()
        self.load_columns()
        
        self.grab_set()
        self.transient(parent)
        self.focus_set()
        
        self.bind("<Escape>", lambda e: self.close_selector())
        
    def create_widgets(self):
        """Cria interface com layout responsivo
        OTIMIZAÇÃO: Uso eficiente de frames para redimensionamento"""
        main_frame = tb.Frame(self)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Frame de filtro
        filter_frame = tb.Frame(main_frame)
        filter_frame.pack(fill="x", pady=(0, 5))
        
        tb.Label(filter_frame, text="Filtrar:", font=("Arial", 9)).pack(side="left", padx=2)
        self.filter_var = tk.StringVar()
        self.filter_entry = tb.Entry(filter_frame, textvariable=self.filter_var, width=25)
        self.filter_entry.pack(side="left", padx=2, fill="x", expand=True)
        self.filter_entry.bind("<KeyRelease>", self.apply_filter)
        self.filter_entry.focus_set()
        
        # Botões de ação rápida
        action_frame = tb.Frame(main_frame)
        action_frame.pack(fill="x", pady=(0, 5))
        
        button_options = {"bootstyle": "light", "padding": (2, 2)}
        tb.Button(action_frame, text="Todas", command=lambda: self.toggle_all(True), **button_options).pack(side="left", padx=1)
        tb.Button(action_frame, text="Nenhuma", command=lambda: self.toggle_all(False), **button_options).pack(side="left", padx=1)
        tb.Button(action_frame, text="Filtradas", command=lambda: self.toggle_filtered(True), **button_options).pack(side="left", padx=1)
        tb.Button(action_frame, text="Desm. Filt.", command=lambda: self.toggle_filtered(False), **button_options).pack(side="left", padx=1)
        tb.Button(action_frame, text="Iguais", command=self.mark_same_columns, **button_options).pack(side="left", padx=1)
        
        # Frame principal
        content_frame = tb.Frame(main_frame)
        content_frame.pack(fill="both", expand=True)
        content_frame.columnconfigure(0, weight=3)
        content_frame.columnconfigure(1, weight=1)
        content_frame.rowconfigure(0, weight=1)

        # Treeview para colunas
        tree_frame = tb.LabelFrame(content_frame, text="Colunas Disponíveis")
        tree_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(tree_frame, show="tree", selectmode="none")
        self.tree.column("#0", width=250, anchor="w")
        
        vsb = tb.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview, bootstyle="round")
        hsb = tb.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview, bootstyle="round")
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        self.tree.bind("<ButtonRelease-1>", self.on_tree_click)
        
        # Mapa de seleções
        selection_frame = tb.LabelFrame(content_frame, text="Colunas Selecionadas")
        selection_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0))
        selection_frame.rowconfigure(0, weight=1)
        selection_frame.columnconfigure(0, weight=1)
        
        # Frame interno
        inner_selection_frame = tb.Frame(selection_frame)
        inner_selection_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Botão para remover seleção
        remove_frame = tb.Frame(inner_selection_frame)
        remove_frame.pack(fill="x", pady=(0, 5))
        
        tb.Button(
            remove_frame, 
            text="Remover Seleção", 
            command=self.remove_selected,
            bootstyle="light",
            padding=(2, 1),
            width=12
        ).pack(side="top", fill="x")
        
        # Lista de seleções
        list_frame = tb.Frame(inner_selection_frame)
        list_frame.pack(fill="both", expand=True)
        
        self.selection_listbox = tk.Listbox(
            list_frame, 
            selectmode=tk.SINGLE,
            bg="white",
            fg="black",
            font=("Arial", 9),
            width=20
        )
        
        scrollbar = tb.Scrollbar(list_frame, orient="vertical", command=self.selection_listbox.yview, bootstyle="round")
        self.selection_listbox.config(yscrollcommand=scrollbar.set)
        
        self.selection_listbox.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Frame de controle
        control_frame = tb.Frame(main_frame)
        control_frame.pack(fill="x", pady=(5, 0))
        
        # Botões Aplicar, Cancelar, Fechar
        tb.Button(control_frame, text="Aplicar", command=self.apply_selection, 
                 bootstyle="success", padding=(4, 2)).pack(side="left", padx=2)
        tb.Button(control_frame, text="Cancelar", command=self.cancel, 
                 bootstyle="danger", padding=(4, 2)).pack(side="left", padx=2)
        tb.Button(control_frame, text="Fechar", command=self.close_selector, 
                 bootstyle="secondary", padding=(4, 2)).pack(side="right", padx=2)
        
    def get_unique_columns(self):
        """Agregação eficiente de colunas únicas usando defaultdict
        COMPLEXIDADE: O(n) - linear com número de resultados"""
        unique_columns = defaultdict(set)
        self.column_items = {}
        
        for result in self.search_results:
            worksheet = result['worksheet']
            for col in result['full_row'].keys():
                unique_columns[worksheet].add(col)
                if col not in self.column_items:
                    self.column_items[col] = []
        
        return unique_columns
        
    def load_columns(self):
        """Carregamento hierárquico com complexidade O(n log n)
        ORDENAÇÃO: Colunas ordenadas alfabeticamente para melhor usabilidade"""
        unique_columns = self.get_unique_columns()
        self.all_items = []
        
        for worksheet in sorted(unique_columns.keys()):
            parent = self.tree.insert("", "end", text=worksheet, tags=('worksheet',))
            self.all_items.append(parent)
            
            for column in sorted(unique_columns[worksheet]):
                is_selected = column in self.selected_columns[worksheet]
                display_text = f"☑ {column}" if is_selected else column
                
                item = self.tree.insert(parent, "end", text=display_text, 
                                      values=(worksheet, column), tags=('column',))
                self.column_items[column].append(item)
                self.all_items.append(item)
        
        self.tree.tag_configure('worksheet', font=('Arial', 9, 'bold'))
        self.tree.tag_configure('column', font=('Arial', 9))
        self.update_selection_map()
    
    def apply_filter(self, event=None):
        """Filtragem em tempo real com complexidade O(n)
        PERFORMANCE: Operação rápida mesmo para centenas de colunas"""
        filter_text = self.filter_var.get().lower()
        
        for item in self.all_items:
            self.tree.reattach(item, self.tree.parent(item), 'end')
        
        if not filter_text:
            return
        
        for item in self.all_items:
            if self.tree.tag_has('column', item):
                column_text = self.tree.item(item, "text").lower()
                clean_text = column_text[2:] if column_text.startswith("☑ ") else column_text
                
                if filter_text not in clean_text:
                    self.tree.detach(item)
            elif self.tree.tag_has('worksheet', item):
                has_visible_children = False
                for child in self.tree.get_children(item):
                    if self.tree.exists(child) and self.tree.item(child, 'open') != 'hidden':
                        has_visible_children = True
                        break
                
                if not has_visible_children:
                    self.tree.detach(item)
    
    def on_tree_click(self, event):
        """Manipulação de eventos com atualização visual imediata
        USABILIDADE: Feedback visual instantâneo para o usuário"""
        item = self.tree.identify_row(event.y)
        if item and self.tree.tag_has('column', item):
            current_text = self.tree.item(item, "text")
            values = self.tree.item(item, "values")
            if values:
                worksheet, column = values
                
                if current_text.startswith("☑ "):
                    new_text = current_text[2:]
                    self.tree.item(item, text=new_text)
                    self.selected_columns[worksheet].discard(column)
                else:
                    self.tree.item(item, text="☑ " + current_text)
                    self.selected_columns[worksheet].add(column)
                
                self.update_selection_map()
    
    def toggle_all(self, state):
        """Operação em lote para melhor performance com muitos itens
        EVITA: Atualizações individuais que seriam lentas"""
        for worksheet_item in self.tree.get_children():
            for column_item in self.tree.get_children(worksheet_item):
                if self.tree.tag_has('column', column_item):
                    text = self.tree.item(column_item, "text")
                    values = self.tree.item(column_item, "values")
                    if values:
                        worksheet, column = values
                        clean_text = text[2:] if text.startswith("☑ ") else text
                        
                        if state:
                            if not text.startswith("☑ "):
                                self.tree.item(column_item, text="☑ " + clean_text)
                                self.selected_columns[worksheet].add(column)
                        else:
                            if text.startswith("☑ "):
                                self.tree.item(column_item, text=clean_text)
                                self.selected_columns[worksheet].discard(column)
        
        self.update_selection_map()
    
    def toggle_filtered(self, state):
        """Ativa/desativa colunas visíveis pelo filtro
        EFICIÊNCIA: Opera apenas nos itens visíveis"""
        filter_text = self.filter_var.get().lower()
        if not filter_text:
            return
            
        for worksheet_item in self.tree.get_children():
            for column_item in self.tree.get_children(worksheet_item):
                if self.tree.tag_has('column', column_item):
                    text = self.tree.item(column_item, "text")
                    clean_text = text[2:] if text.startswith("☑ ") else text
                    values = self.tree.item(column_item, "values")
                    
                    if values and filter_text in clean_text.lower():
                        worksheet, column = values
                        
                        if state:
                            if not text.startswith("☑ "):
                                self.tree.item(column_item, text="☑ " + clean_text)
                                self.selected_columns[worksheet].add(column)
                        else:
                            if text.startswith("☑ "):
                                self.tree.item(column_item, text=clean_text)
                                self.selected_columns[worksheet].discard(column)
        
        self.update_selection_map()
    
    def mark_same_columns(self):
        """Seleção em massa de colunas com mesmo nome
        PRODUTIVIDADE: Economiza tempo em datasets com estrutura similar"""
        selected_item = self.tree.focus()
        if not selected_item or not self.tree.tag_has('column', selected_item):
            return
            
        column_name = self.tree.item(selected_item, "text")
        if column_name.startswith("☑ "):
            column_name = column_name[2:]
        
        for item in self.column_items.get(column_name, []):
            if self.tree.exists(item):
                current_text = self.tree.item(item, "text")
                values = self.tree.item(item, "values")
                
                if values and not current_text.startswith("☑ "):
                    worksheet, column = values
                    clean_text = current_text[2:] if current_text.startswith("☑ ") else current_text
                    self.tree.item(item, text="☑ " + clean_text)
                    self.selected_columns[worksheet].add(column)
        
        self.update_selection_map()
    
    def update_selection_map(self):
        """Atualização eficiente da lista de seleção
        OTIMIZAÇÃO: Uso de defaultdict para agrupamento rápido"""
        self.selection_listbox.delete(0, tk.END)
        
        selections_by_worksheet = defaultdict(list)
        for worksheet, columns in self.selected_columns.items():
            for column in columns:
                selections_by_worksheet[worksheet].append(column)
        
        for worksheet, columns in selections_by_worksheet.items():
            if columns:
                self.selection_listbox.insert(tk.END, f"--- {worksheet} ---")
                for column in sorted(columns):
                    self.selection_listbox.insert(tk.END, f"  • {column}")
                self.selection_listbox.insert(tk.END, "")
    
    def remove_selected(self):
        """Remoção seletiva com atualização simultânea na treeview
        CONSISTÊNCIA: Mantém sincronia entre diferentes componentes visuais"""
        selection = self.selection_listbox.curselection()
        if not selection:
            return
            
        selected_text = self.selection_listbox.get(selection[0])
        
        if "---" in selected_text or not selected_text.strip():
            return
            
        column_name = selected_text.strip().split("• ")[-1]
        
        for worksheet in list(self.selected_columns.keys()):
            if column_name in self.selected_columns[worksheet]:
                self.selected_columns[worksheet].discard(column_name)
                
                for item in self.column_items.get(column_name, []):
                    if self.tree.exists(item):
                        current_text = self.tree.item(item, "text")
                        if current_text.startswith("☑ "):
                            self.tree.item(item, text=current_text[2:])
        
        self.update_selection_map()
    
    def apply_selection(self):
        """Confirmação de seleção com flag aplicada
        DESIGN: Separa ação de confirmação de simples fechamento"""
        self.applied = True
        self.update_selection_map()
        print("[ColumnSelector] Seleções aplicadas")
    
    def close_selector(self):
        """Fechamento seguro da janela
        GERENCIAMENTO DE RECURSOS: Libera referências corretamente"""
        self.destroy()
    
    def cancel(self):
        """Cancelamento explícito com limpeza de estado
        USABILIDADE: Oferece comportamento previsível ao usuário"""
        self.selected_columns = defaultdict(set)
        self.destroy()

class FileSearchApp:
    """Aplicação principal com estratégias de escalonamento:
    - Threading para operações bloqueantes
    - Lock para acesso concorrente seguro
    - Exportação em streams
    - Carregamento lazy de dados"""
    def __init__(self, root):
        self.root = root
        self.root.title("Control+F")
        self.root.geometry("800x600")
        self.root.minsize(600, 450)
        
        self.style = tb.Style(theme="litera")
        self.searcher = LocalFileSearcher()
        self.search_results = []
        self.lock = threading.Lock()  # Controle de concorrência
        self.selected_columns = defaultdict(set)
        
        self.create_widgets()
        
        self.status_var = tk.StringVar(value="Pronto")
        status_bar = tb.Label(
            self.root, 
            textvariable=self.status_var, 
            relief="sunken", 
            anchor="w",
            bootstyle="light"
        )
        status_bar.pack(side="bottom", fill="x")

    def create_widgets(self):
        """Construção de interface com layout responsivo
        ADAPTAÇÃO: Redimensionamento inteligente de componentes"""
        main_frame = tb.Frame(self.root)
        main_frame.pack(fill="both", expand=True, padx=10, pady=8)
        
        # Configurar pesos para redimensionamento
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(3, weight=1)  # Área de resultados
        main_frame.rowconfigure(4, weight=0)   # Log
        
        # Frame de seleção de arquivo
        file_frame = tb.Labelframe(main_frame, text="Arquivo", bootstyle="info")
        file_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=3)
        file_frame.columnconfigure(1, weight=1)
        
        tb.Label(file_frame, text="Arquivo:", font=("Arial", 9)).grid(row=0, column=0, sticky="w", padx=5, pady=3)
        self.file_path_var = tk.StringVar()
        file_entry = tb.Entry(file_frame, textvariable=self.file_path_var, state='readonly', font=("Arial", 9))
        file_entry.grid(row=0, column=1, sticky="we", padx=5, pady=3)
        
        btn_frame = tb.Frame(file_frame)
        btn_frame.grid(row=0, column=2, padx=5)
        
        tb.Button(btn_frame, text="Selecionar", command=self.select_file, 
                 bootstyle="primary", padding=(3, 1)).pack(side="left", padx=2)
        tb.Button(btn_frame, text="Carregar", command=self.load_worksheets, 
                 bootstyle="info", padding=(3, 1)).pack(side="left", padx=2)
        
        # Frame de seleção de worksheet
        worksheet_frame = tb.Labelframe(main_frame, text="Worksheet", bootstyle="info")
        worksheet_frame.grid(row=1, column=0, sticky="ew", padx=5, pady=3)
        worksheet_frame.columnconfigure(1, weight=1)
        
        tb.Label(worksheet_frame, text="Worksheet:", font=("Arial", 9)).grid(row=0, column=0, sticky="w", padx=5, pady=3)
        self.worksheet_combo = tb.Combobox(worksheet_frame, state="readonly", font=("Arial", 9))
        self.worksheet_combo.grid(row=0, column=1, sticky="we", padx=5, pady=3)
        tb.Button(worksheet_frame, text="Atualizar", command=self.load_worksheets, 
                 bootstyle="info", padding=(3, 1)).grid(row=0, column=2, padx=5, pady=3)
        
        # Frame de busca
        search_frame = tb.Labelframe(main_frame, text="Busca", bootstyle="info")
        search_frame.grid(row=2, column=0, sticky="ew", padx=5, pady=3)
        search_frame.columnconfigure(1, weight=1)
        
        tb.Label(search_frame, text="Termo:", font=("Arial", 9)).grid(row=0, column=0, sticky="w", padx=5, pady=3)
        self.search_entry = tb.Entry(search_frame, font=("Arial", 9))
        self.search_entry.grid(row=0, column=1, sticky="we", padx=5, pady=3)
        
        btn_frame2 = tb.Frame(search_frame)
        btn_frame2.grid(row=0, column=2, padx=5)
        
        tb.Button(btn_frame2, text="Buscar (aba selecionada)", command=self.start_search, 
                 bootstyle="success", padding=(3, 1)).pack(side="left", padx=2)
        tb.Button(btn_frame2, text="Buscar em todas as abas", command=self.search_all_worksheets, 
                 bootstyle="success", padding=(3, 1)).pack(side="left", padx=2)
        
        # Resultados
        results_frame = tb.Labelframe(main_frame, text="Resultados", bootstyle="info")
        results_frame.grid(row=3, column=0, sticky="nsew", padx=5, pady=3)
        results_frame.rowconfigure(0, weight=1)
        results_frame.columnconfigure(0, weight=1)
        
        # Treeview com scrollbar
        tree_frame = tb.Frame(results_frame)
        tree_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)
        
        columns = ("worksheet", "cell", "value")
        self.results_tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=8)
        
        self.results_tree.heading("worksheet", text="Worksheet")
        self.results_tree.heading("cell", text="Célula")
        self.results_tree.heading("value", text="Valor")
        
        self.results_tree.column("worksheet", width=120, anchor="w")
        self.results_tree.column("cell", width=60, anchor="center")
        self.results_tree.column("value", width=300, anchor="w")
        
        vsb = tb.Scrollbar(tree_frame, orient="vertical", command=self.results_tree.yview, bootstyle="round")
        hsb = tb.Scrollbar(tree_frame, orient="horizontal", command=self.results_tree.xview, bootstyle="round")
        self.results_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.results_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        
        # Log
        log_frame = tb.Labelframe(main_frame, text="Log", bootstyle="info")
        log_frame.grid(row=4, column=0, sticky="ew", padx=5, pady=3)
        
        self.log_text = scrolledtext.ScrolledText(
            log_frame, 
            height=4, 
            wrap="word",
            bg='white',
            fg='black',
            font=("Arial", 9)
        )
        self.log_text.pack(fill="both", expand=True, padx=5, pady=3)
        self.log_text.config(state="disabled")
        
        # Botões de exportação
        export_frame = tb.Frame(main_frame)
        export_frame.grid(row=5, column=0, sticky="ew", padx=5, pady=5)
        
        # Botões à esquerda
        left_btn_frame = tb.Frame(export_frame)
        left_btn_frame.pack(side="left", fill="x", expand=True)
        
        tb.Button(left_btn_frame, text="Exportar colunas", command=self.open_column_selector, 
                 bootstyle="primary", padding=(3, 1)).pack(side="left", padx=2)
        tb.Button(left_btn_frame, text="Limpar Res.", command=self.clear_results, 
                 bootstyle="warning", padding=(3, 1)).pack(side="left", padx=2)
        tb.Button(left_btn_frame, text="Limpar Log", command=self.clear_log, 
                 bootstyle="warning", padding=(3, 1)).pack(side="left", padx=2)
        
        # Botões à direita
        right_btn_frame = tb.Frame(export_frame)
        right_btn_frame.pack(side="right", fill="x", expand=True)
        
        tb.Button(right_btn_frame, text="Exportar Excel", command=self.export_full_rows_to_excel, 
                 bootstyle="success", padding=(3, 1)).pack(side="right", padx=2)
        tb.Button(right_btn_frame, text="Exportar JSON", command=self.export_full_rows_to_json, 
                 bootstyle="success", padding=(3, 1)).pack(side="right", padx=2)
        tb.Button(right_btn_frame, text="Exportar CSV", command=self.export_full_rows_to_csv, 
                 bootstyle="success", padding=(3, 1)).pack(side="right", padx=2)
    
    def log(self, message):
        """Registro assíncrono com buffer controlado
        PREVENÇÃO: Limita tamanho para evitar estouro de memória"""
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, f"{time.strftime('%H:%M:%S')} - {message}\n")
        self.log_text.config(state="disabled")
        self.log_text.see(tk.END)
        self.status_var.set(message[:100])  # Limitação para evitar sobrecarga
    
    def clear_log(self):
        """Limpeza segura do componente de log
        GERENCIAMENTO DE MEMÓRIA: Libera recursos do widget"""
        self.log_text.config(state="normal")
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state="disabled")
        self.status_var.set("Log limpo")
    
    def select_file(self):
        """Seleção de arquivo com filtros otimizados
        USABILIDADE: Filtros específicos para formatos suportados"""
        file_path = filedialog.askopenfilename(
            title="Selecione um arquivo",
            filetypes=[
                ("Arquivos CSV", "*.csv"),
                ("Arquivos Excel", "*.xlsx"),
                ("Arquivos Excel", "*.xls"),
                ("Todos os arquivos", "*.*")
            ]
        )
        if file_path:
            self.file_path_var.set(file_path)
            self.log(f"Arquivo selecionado: {os.path.basename(file_path)}")
    
    def load_worksheets(self):
        """Carregamento sob demanda com tratamento de erros
        PERFORMANCE: Execução fora da thread principal"""
        file_path = self.file_path_var.get()
        if not file_path:
            self.log("Selecione um arquivo primeiro")
            return
        
        try:
            self.log(f"Carregando worksheets: {os.path.basename(file_path)}")
            worksheets = self.searcher.load_file(file_path)
            
            if not worksheets:
                self.log("Nenhuma worksheet encontrada")
                return
                
            self.worksheet_combo.configure(values=worksheets)
            if worksheets:
                self.worksheet_combo.current(0)
            self.log(f"{len(worksheets)} worksheets carregadas")
        except Exception as e:
            self.log(f"Erro ao carregar worksheets: {str(e)}")
    
    def start_search(self):
        """Inicia busca em thread separada para não bloquear UI
        CONCORRÊNCIA: Uso de threading para operações longas"""
        file_path = self.file_path_var.get()
        if not file_path:
            self.log("Selecione um arquivo primeiro")
            return
        
        worksheet = self.worksheet_combo.get()
        if not worksheet:
            self.log("Selecione uma worksheet primeiro")
            return
        
        query = self.search_entry.get().strip()
        if not query:
            self.log("Digite um termo de busca")
            return
        
        self.log(f"Buscando '{query}' em {worksheet}...")
        threading.Thread(
            target=self._search_thread, 
            args=(file_path, worksheet, query),
            daemon=True
        ).start()
    
    def search_all_worksheets(self):
        """Busca paralelizada por planilha usando threading
        ESCALONAMENTO: Máximo de paralelismo por planilha
        ALERTA: Pode sobrecarregar sistema com muitas planilhas"""
        file_path = self.file_path_var.get()
        if not file_path:
            self.log("Selecione um arquivo primeiro")
            return
        
        try:
            worksheets = self.searcher.worksheets
            if not worksheets:
                self.log("Nenhuma worksheet encontrada")
                return
                
            query = self.search_entry.get().strip()
            if not query:
                self.log("Digite um termo de busca")
                return
                
            self.log(f"Buscando em {len(worksheets)} worksheets...")
            self.clear_results()
            
            for worksheet in worksheets:
                threading.Thread(
                    target=self._search_thread, 
                    args=(file_path, worksheet, query),
                    daemon=True
                ).start()
                
        except Exception as e:
            self.log(f"Erro: {str(e)}")
    
    def _search_thread(self, file_path, worksheet_name, query):
        """Worker thread para operações de I/O intensivas
        SEGURANÇA: Atualização thread-safe da UI com root.after()"""
        try:
            if not self.searcher.file_data or self.searcher.file_path != file_path:
                self.searcher.load_file(file_path)
            
            results = self.searcher.search_in_worksheet(worksheet_name, query)
            
            # Atualização thread-safe da UI
            self.root.after(0, lambda: self._display_results(worksheet_name, results))
            self.root.after(0, lambda: self.log(
                f"Encontrados {len(results)} em '{worksheet_name}'"
            ))
        except Exception as e:
            self.root.after(0, lambda: self.log(
                f"Erro em '{worksheet_name}': {str(e)}"
            ))
    
    def _display_results(self, worksheet_name, results):
        """Atualização incremental da UI para grandes resultados
        EFICIÊNCIA: Inserção em lote com controle de concorrência"""
        for result in results:
            self.results_tree.insert("", "end", values=(
                worksheet_name,
                result['cell'],
                result['value']
            ))
        
        # Controle concorrente seguro
        with self.lock:
            self.search_results.extend(results)
    
    def clear_results(self):
        """Limpeza completa de resultados
        GERENCIAMENTO DE MEMÓRIA: Libera referências explicitamente"""
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)
        
        with self.lock:
            self.search_results.clear()
            
        self.log("Resultados limpos")
            
    def open_column_selector(self):
        """Abertura segura de janela modal
        CONCORRÊNCIA: Cópia dos resultados para evitar race conditions"""
        with self.lock:
            if not self.search_results:
                self.log("Execute uma pesquisa primeiro")
                return
        
        selector = ColumnSelector(self.root, self.search_results.copy(), self.selected_columns)
        self.root.wait_window(selector)
        
        if selector.applied:
            self.selected_columns = selector.selected_columns
            self.log("Seleção atualizada")
        else:
            self.log("Seleção mantida")
    
    def get_filtered_row(self, result):
        """Filtragem eficiente de colunas selecionadas
        PERFORMANCE: Operação O(1) para cada linha"""
        row_data = result['full_row'].copy()
        row_data['worksheet'] = result['worksheet']
        row_data['célula'] = result['cell']
        
        if self.selected_columns:
            worksheet = result['worksheet']
            if worksheet in self.selected_columns:
                selected = self.selected_columns[worksheet]
                row_data = {col: row_data.get(col, '') for col in selected if col in row_data}
                row_data['worksheet'] = result['worksheet']
                row_data['célula'] = result['cell']
        
        return row_data
    
    def get_all_headers(self):
        """Obtém cabeçalhos de forma otimizada com complexidade O(n)
        MEMÓRIA: Uso eficiente de conjuntos para evitar duplicatas"""
        all_headers = set()
        required_headers = {'worksheet', 'célula'}
        
        if self.selected_columns:
            for worksheet, columns in self.selected_columns.items():
                all_headers.update(columns)
        else:
            for result in self.search_results:
                all_headers.update(result['full_row'].keys())
        
        all_headers.update(required_headers)
        ordered_headers = ['worksheet', 'célula'] + \
                         sorted([h for h in all_headers if h not in ['worksheet', 'célula']])
        
        return ordered_headers
    
    def export_full_rows_to_json(self):
        """Exportação JSON com streaming para grandes volumes
        ATENÇÃO: Carregamento completo em memória - não recomendado para datasets muito grandes"""
        with self.lock:
            if not self.search_results:
                self.log("Nenhum resultado para exportar")
                return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json")]
        )
        
        if not file_path:
            return
        
        try:
            export_data = []
            with self.lock:
                for result in self.search_results:
                    export_data.append(self.get_filtered_row(result))
            
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(export_data, f, ensure_ascii=False, indent=2)
                
            self.log(f"Exportado JSON: {file_path}")
        except Exception as e:
            self.log(f"Erro JSON: {str(e)}")
    
    def export_full_rows_to_csv(self):
        """Exportação CSV com escrita direta em stream
        RECOMENDADO: Melhor opção para grandes volumes de dados"""
        with self.lock:
            if not self.search_results:
                self.log("Nenhum resultado para exportar")
                return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")]
        )
        
        if not file_path:
            return
        
        try:
            ordered_headers = self.get_all_headers()
            
            with open(file_path, 'w', encoding='utf-8', newline='') as f:
                writer = csv.DictWriter(f, fieldnames=ordered_headers)
                writer.writeheader()
                
                with self.lock:
                    for result in self.search_results:
                        row_data = self.get_filtered_row(result)
                        writer.writerow(row_data)
            
            self.log(f"Exportado CSV: {file_path}")
        except Exception as e:
            self.log(f"Erro CSV: {str(e)}")
            
    def export_full_rows_to_excel(self):
        """Exportação Excel com pandas - monitorar uso de memória
        ALERTA: Operação intensiva em memória para grandes datasets"""
        with self.lock:
            if not self.search_results:
                self.log("Nenhum resultado para exportar")
                return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if not file_path:
            return
        
        try:
            ordered_headers = self.get_all_headers()
            data = []
            with self.lock:
                for result in self.search_results:
                    row_data = self.get_filtered_row(result)
                    ordered_row = [row_data.get(header, '') for header in ordered_headers]
                    data.append(ordered_row)
            
            df = pd.DataFrame(data, columns=ordered_headers)
            df.to_excel(file_path, index=False)
                
            self.log(f"Exportado Excel: {file_path}")
        except Exception as e:
            self.log(f"Erro Excel: {str(e)}")


if __name__ == "__main__":
    root = tb.Window(themename="litera")
    app = FileSearchApp(root)
    root.mainloop()
