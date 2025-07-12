# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
import threading
import queue
import os
import ftplib
import pandas as pd
import openpyxl # Necessário para o script XML
import xml.etree.ElementTree as ET # Necessário para o script XML
from datetime import datetime
import sys
import re
import subprocess
import json
import socket
import requests # Mantido para verificação de acesso (se reativada, no momento comentada)

# --- FUNÇÃO PARA ARQUIVOS PERMANENTES (ESSENCIAL PARA O .EXE) ---
def get_persistent_path(filename):
    """ Obtém o caminho para um arquivo na mesma pasta do script ou do .exe """
    if getattr(sys, 'frozen', False):
        application_path = os.path.dirname(sys.executable)
    else:
        application_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(application_path, filename)

# ==============================================================================
# CLASSE DE TEMPLATE DA APLICAÇÃO (BASE) - ADAPTADA PARA UNIFICAÇÃO
# ==============================================================================
class ModernAppTemplate(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.nav_buttons = {}
        self.content_frames = {}
        self.NAV_ITEMS = {} # A ser definido pela classe que herda (MyNewApp)

        # Cores e tema padrão (podem ser substituídos pela classe que herda)
        self.COLOR_PRIMARY = "#3B8ED0" # Azul
        self.CARD_FG_COLOR = ("#FFFFFF", "#2B2B2B")
        self.WINDOW_BG_COLOR = ("#F2F2F2", "#242424")
        self.TEXT_SUBTLE_COLOR = ("#5E5E5E", "#A0A0A0")
        
        self.CARD_CORNER_RADIUS = 12
        self.BUTTON_CORNER_RADIUS = 10
        self.PADX = 10 # Adicionado como atributo da instância
        self.PADY = 5  # Adicionado como atributo da instância

        self.FONT_TITLE = ctk.CTkFont(size=30, weight="bold")
        self.FONT_H1 = ctk.CTkFont(size=20, weight="bold")
        self.FONT_H2 = ctk.CTkFont(size=16, weight="bold")
        self.FONT_BODY = ctk.CTkFont(size=13)
        self.FONT_BUTTON = ctk.CTkFont(size=14, weight="bold")

    def _initialize_ui(self):
        self.configure(fg_color=self.WINDOW_BG_COLOR)
        self._create_navigation_frame()
        self._create_content_frames()
        
        # Seleciona a primeira página por padrão (XML)
        first_page_name = next(iter(self.NAV_ITEMS))
        self.select_frame_by_name(first_page_name)

    def _create_navigation_frame(self):
        nav_frame = ctk.CTkFrame(self, width=250, corner_radius=0, fg_color=("gray92", "#242424"))
        nav_frame.grid(row=0, column=0, sticky="nsw")
        nav_frame.grid_rowconfigure(len(self.NAV_ITEMS) + 2, weight=1) # +2 para logo e espaçamento

        logo_label = ctk.CTkLabel(nav_frame, text="Gerador de Pedidos", font=ctk.CTkFont(size=20, weight="bold"))
        logo_label.grid(row=0, column=0, padx=20, pady=(25, 25))
        
        for i, (name, (icon, text)) in enumerate(self.NAV_ITEMS.items(), start=1):
            button = ctk.CTkButton(
                nav_frame,
                text=f"  {icon}    {text}", # Usando f-string para ícone e texto
                height=45,
                corner_radius=self.BUTTON_CORNER_RADIUS,
                fg_color="transparent",
                text_color=("gray10", "#DCE4EE"),
                hover_color=("gray88", "#2E2E2E"),
                anchor="w",
                font=ctk.CTkFont(size=14),
                command=lambda n=name: self.select_frame_by_name(n)
            )
            button.grid(row=i, column=0, padx=15, pady=6, sticky="ew")
            self.nav_buttons[name] = button
        
        # Rótulo da versão do aplicativo na parte inferior da barra lateral
        app_version = getattr(self, 'APP_VERSION', 'N/A')
        version_label = ctk.CTkLabel(nav_frame, text=f"Versão {app_version}",
                                     font=ctk.CTkFont(size=11),
                                     text_color="gray")
        version_label.grid(row=len(self.NAV_ITEMS) + 3, column=0, padx=10, pady=15, sticky="s")


    def _create_content_frames(self):
        self.content_area = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.content_area.grid(row=0, column=1, sticky="nsew", padx=15, pady=15)
        self.content_area.grid_rowconfigure(0, weight=1)
        self.content_area.grid_columnconfigure(0, weight=1)

        # Cria um frame para cada item de navegação e o armazena
        for name in self.NAV_ITEMS:
            frame = ctk.CTkFrame(self.content_area, fg_color="transparent")
            frame.grid(row=0, column=0, sticky="nsew") # Todos os frames na mesma célula da grade
            self.content_frames[name] = frame

        # Chama os métodos de configuração específicos para cada página
        for name in self.NAV_ITEMS:
            setup_method_name = f"setup_{name}_page"
            if hasattr(self, setup_method_name) and callable(getattr(self, setup_method_name)):
                getattr(self, setup_method_name)(self.content_frames[name])
            else:
                print(f"Aviso: Método de configuração '{setup_method_name}' não implementado para a página '{name}'.")


    def select_frame_by_name(self, name):
        """Alterna o frame de conteúdo visível e atualiza a aparência do botão de navegação."""
        primary_color = self.COLOR_PRIMARY
        
        # Reinicia todos os botões de navegação para o estado desselecionado
        for btn_name, button in self.nav_buttons.items():
            button.configure(fg_color="transparent", text_color=("gray10", "#DCE4EE"))
        
        # Destaca o botão selecionado
        selected_button = self.nav_buttons.get(name)
        if selected_button:
            selected_button.configure(fg_color=primary_color, text_color=("white", "white"))
        
        # Traz o frame de conteúdo correspondente para a frente
        frame_to_show = self.content_frames.get(name)
        if frame_to_show:
            frame_to_show.tkraise()

    def _create_toolbar(self, parent_frame, title):
        """Cria uma barra de ferramentas consistente para cada página de conteúdo."""
        toolbar_frame = ctk.CTkFrame(parent_frame, fg_color="transparent", height=60)
        toolbar_frame.pack(fill="x", pady=(0, 10), padx=5)
        
        title_font = self.FONT_TITLE
        title_label = ctk.CTkLabel(toolbar_frame, text=title, font=title_font)
        title_label.pack(side="left", padx=5, pady=10)
        return toolbar_frame

    # Métodos placeholder para configuração de página - devem ser implementados pela classe que herda
    def setup_xml_page(self, parent_frame): pass
    def setup_epan_txt_page(self, parent_frame): pass
    def setup_configuracoes_page(self, parent_frame): pass
    def setup_sobre_page(self, parent_frame): pass

    def change_appearance_mode(self, new_mode_str: str):
        """Muda o modo de aparência do CustomTkinter (Claro/Escuro/Sistema)."""
        mode_map = {"Claro": "light", "Escuro": "dark", "Sistema": "system"}
        ctk_mode = mode_map.get(new_mode_str)
        if ctk_mode:
            ctk.set_appearance_mode(ctk_mode)
            self._auto_save_on_change() # Aciona o salvamento quando o tema muda

    def _auto_save_on_change(self, *args):
        """Método a ser chamado quando uma variável configurável muda, para acionar o salvamento."""
        # Este método precisa ser implementado na classe que herda
        # para salvar configurações específicas do aplicativo.
        print("Auto-save acionado (implemente a lógica de salvamento na classe filha).")
        self._salvar_configuracoes()

    def _salvar_configuracoes(self):
        """Salva as configurações atuais do aplicativo em um arquivo JSON. (Implementado na classe filha)"""
        pass

    def _carregar_configuracoes(self):
        """Carrega as configurações na inicialização. (Implementado na classe filha)"""
        pass

    def _update_ui_from_config(self):
        """Atualiza elementos da UI com base nas configurações carregadas (ex: botão segmentado de tema)."""
        if hasattr(self, 'theme_segmented_button'):
            appearance_mode = ctk.get_appearance_mode()
            mode_map_reverso = {"light": "Claro", "dark": "Escuro", "system": "Sistema"}
            loaded_value_portuguese = mode_map_reverso.get(appearance_mode.lower(), "Sistema")
            self.theme_segmented_button.set(loaded_value_portuguese)

# ==============================================================================
# CONFIGURAÇÕES GLOBAIS E DE AMBIENTES (MESCLADAS)
# ==============================================================================
# Estas variáveis globais podem ser acessadas de dentro da classe,
# mas para maior organização e encapsulamento, as configurações do FTP
# e diretórios de output poderiam ser passadas como argumentos ou lidas de um config file
# dentro da classe principal. Por simplicidade, as mantive aqui como estavam nos originais.

# Diretórios de saída para os arquivos gerados
# Todos os arquivos serão salvos em uma subpasta 'Pedidos_Gerados_Unified'
# e depois em subpastas 'XML_Pedidos' ou 'TXT_Pedidos'
OUTPUT_BASE_DIR_UNIFIED = get_persistent_path("Pedidos_Gerados_Unified")
OUTPUT_XML_DIR = os.path.join(OUTPUT_BASE_DIR_UNIFIED, "XML_Pedidos")
OUTPUT_TXT_DIR = os.path.join(OUTPUT_BASE_DIR_UNIFIED, "TXT_Pedidos")

# Nome do arquivo para log de erros (pode ser compartilhado ou separado)
arquivo_erro_xlsx = get_persistent_path('erros_geracao_unificada.xlsx')

# Configurações FTP para Upload de XML (do Script 1)
FTP_XML_UPLOAD_HOST = "10.41.15.19"
FTP_XML_UPLOAD_PORT = 21 # Default para FTP
FTP_XML_UPLOAD_USER = "filgo"
FTP_XML_UPLOAD_PASS = "ecomsap@123"
FTP_XML_UPLOAD_PATH = "/saptxt/new/in/ecom/portal/xml_in/" # CONFIRMAR/AJUSTAR

# Configurações FTP para Upload de TXT (do Script 2)
FTP_TXT_UPLOAD_HOST = "10.41.15.19"
FTP_TXT_UPLOAD_PORT = 21 # Default para FTP
FTP_TXT_USER_PADRAO = "filgo" # Usuário padrão para TXT
FTP_TXT_PASS_PADRAO = "ecomsap@123" # Senha padrão para TXT
FTP_TXT_PATHS_PADRAO = {
    "EPP": "/saptxt/new/in/ecom/portal/pedtxt/epp",
    "EPH": "/saptxt/new/in/ecom/portal/pedtxt/grupo"
}

# ==============================================================================
# FUNÇÕES DE UTILIDADE (MESCLADAS E OTIMIZADAS)
# ==============================================================================
def criar_diretorios(log_queue, path):
    """Cria o diretório especificado, se não existir."""
    try:
        os.makedirs(path, exist_ok=True)
    except Exception as e:
        log_queue.put(f"ERRO CRÍTICO ao criar diretório {path}: {e}")
        raise # Levanta a exceção para que o chamador possa tratá-la

def verificar_numerico(valor):
    """Verifica se um valor pode ser convertido para numérico."""
    try:
        float(str(valor).strip().replace(',', '.'))
        return True
    except (ValueError, TypeError):
        return False

def abrir_arquivo(caminho, log_queue):
    """Abre um arquivo ou diretório no sistema operacional padrão."""
    if not os.path.exists(caminho):
        log_queue.put(f"ERRO: Caminho não encontrado: {caminho}")
        messagebox.showerror("Erro", f"Caminho não encontrado:\n{caminho}")
        return
    try:
        log_queue.put(f"Tentando abrir: {caminho}")
        if sys.platform == "win32":
            os.startfile(caminho)
        elif sys.platform == "darwin":
            subprocess.run(['open', caminho], check=False)
        else:
            subprocess.run(['xdg-open', caminho], check=False)
    except Exception as e:
        log_queue.put(f"ERRO ao tentar abrir '{os.path.basename(caminho)}': {e}")
        messagebox.showerror("Erro", f"Não foi possível abrir:\n{e}")

# ==============================================================================
# CLASSE DA APLICAÇÃO GUI UNIFICADA
# ==============================================================================
class UnifiedOrderGeneratorApp(ModernAppTemplate):
    APP_VERSION = "2.1 (Abas Separadas XML/TXT)"
    CONFIG_FILE = get_persistent_path("unified_order_gen_config.json")

    def __init__(self):
        super().__init__()
        self.title(f"Gerador de Pedidos Unificado - v{self.APP_VERSION}")
        self.geometry("1024x768") # Tamanho inicial ajustado para o layout lateral

        # Variáveis de estado
        self.file_path_var = tk.StringVar()
        self.log_queue = queue.Queue()

        # Variáveis específicas para Geração XML
        self.manual_login_var = tk.StringVar()
        self.manual_oferta_var = tk.StringVar()
        self.manual_nome_base_var = tk.StringVar()
        self.enviar_xml_ftp_var = tk.BooleanVar(value=False)

        # Variáveis específicas para Geração TXT
        self.usuario_txt_var = tk.StringVar()
        self.destino_txt_var = tk.StringVar(value="EPP")
        self.forma_pagamento_txt_var = tk.StringVar(value="Boleto")
        self.enviar_txt_ftp_padrao_var = tk.BooleanVar(value=False)
        self.enviar_txt_ftp_pessoal_var = tk.BooleanVar(value=False)
        self.ftp_pessoal_user_var = tk.StringVar()
        self.ftp_pessoal_pass_var = tk.StringVar()

        # --- Defina seus itens de navegação para o UnifiedOrderGeneratorApp ---
        self.NAV_ITEMS = {
            "xml": ("📄", "Pedidos XML"),
            "epan_txt": ("📝", "Pedidos EPAN (TXT)"),
            "configuracoes": ("🔧", "Configurações"),
            "sobre": ("ℹ️", "Sobre")
        }

        # Carrega as configurações primeiro
        self._carregar_configuracoes()
        # Inicializa os elementos da interface do usuário (deve ser chamado depois de carregar as configurações)
        self._initialize_ui()
        # Configura os gatilhos de salvamento automático para variáveis específicas do aplicativo
        self._setup_auto_save_triggers()
        # Atualiza elementos específicos da UI que dependem da configuração carregada
        self._update_ui_from_config() # Para o botão de tema
        
        # Log de mensagens (unificado e na janela principal, conforme template)
        # O logbox é um atributo da classe principal, acessível em self.log_textbox
        # A template já lida com a grid do content_area. Preciso apenas garantir
        # que o log textbox seja gridado corretamente fora do content_area
        # E que o content_area se ajuste.

        # Ajuste para o log textbox na janela principal
        self.grid_rowconfigure(0, weight=1) # Content area
        self.grid_rowconfigure(1, weight=0) # Label Log
        self.grid_rowconfigure(2, weight=1) # Log Textbox

        ctk.CTkLabel(self, text="Log:", font=ctk.CTkFont(weight="bold")).grid(row=1, column=0, columnspan=2, padx=self.PADX, pady=(self.PADY*2, 2), sticky="w")
        self.log_textbox = ctk.CTkTextbox(self, wrap=tk.WORD, font=("Consolas", 9), corner_radius=self.CARD_CORNER_RADIUS, border_width=1)
        self.log_textbox.grid(row=2, column=0, columnspan=2, padx=self.PADX, pady=(0, self.PADY), sticky="nsew")
        self.log_textbox.configure(state=tk.DISABLED)

        self.after(100, self.process_log_queue)
        
        # Garante que o frame de FTP pessoal esteja oculto inicialmente para TXT
        # Isso precisa ser chamado APÓS a UI ser criada
        self.after(200, self.handle_txt_ftp_padrao_check)
        self.after(200, self.handle_txt_ftp_pessoal_check)

    # --- Métodos de População das Páginas (Implementando o template) ---
    def setup_xml_page(self, parent_frame):
        self._create_toolbar(parent_frame, "Pedidos XML")
        parent_frame.grid_columnconfigure(0, weight=1)
        
        # Frame de controles para XML
        controls_frame = ctk.CTkFrame(parent_frame, fg_color="transparent")
        controls_frame.pack(fill="x", padx=self.PADX, pady=self.PADY, anchor="n")
        controls_frame.grid_columnconfigure(1, weight=1)

        # Entrada para o arquivo Excel (comum)
        ctk.CTkLabel(controls_frame, text="Planilha Excel de Pedidos:").grid(row=0, column=0, columnspan=3, padx=self.PADX, pady=(self.PADY,2), sticky="w")
        entry_arquivo = ctk.CTkEntry(controls_frame, textvariable=self.file_path_var, corner_radius=self.BUTTON_CORNER_RADIUS)
        entry_arquivo.grid(row=1, column=0, columnspan=2, padx=(self.PADX, self.PADY), pady=2, sticky="ew")
        ctk.CTkButton(controls_frame, text="Procurar...", command=self.select_excel_file, width=90, corner_radius=self.BUTTON_CORNER_RADIUS).grid(row=1, column=2, padx=(0, self.PADX), pady=2, sticky="e")

        # Opções específicas de XML
        xml_options_card = ctk.CTkFrame(controls_frame, corner_radius=self.CARD_CORNER_RADIUS, fg_color=self.CARD_FG_COLOR)
        xml_options_card.grid(row=2, column=0, columnspan=3, padx=self.PADX, pady=self.PADY*2, sticky="ew")
        xml_options_card.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(xml_options_card, text="Substituir Valores (Opcional - XML):", font=self.FONT_H2).grid(row=0, column=0, columnspan=2, padx=20, pady=(20, 10), sticky="w")
        ctk.CTkLabel(xml_options_card, text="Login:").grid(row=1, column=0, padx=(20, 5), pady=self.PADY, sticky="w")
        ctk.CTkEntry(xml_options_card, textvariable=self.manual_login_var, placeholder_text="Padrão: pdvlinkmerck", width=180, corner_radius=self.BUTTON_CORNER_RADIUS).grid(row=1, column=1, padx=5, pady=self.PADY, sticky="ew")
        ctk.CTkLabel(xml_options_card, text="Oferta:").grid(row=2, column=0, padx=(20, 5), pady=self.PADY, sticky="w")
        ctk.CTkEntry(xml_options_card, textvariable=self.manual_oferta_var, placeholder_text="Padrão: da planilha", width=180, corner_radius=self.BUTTON_CORNER_RADIUS).grid(row=2, column=1, padx=5, pady=self.PADY, sticky="ew")
        ctk.CTkLabel(xml_options_card, text="Nome Base Arquivo:").grid(row=3, column=0, padx=(20, 5), pady=self.PADY, sticky="w")
        ctk.CTkEntry(xml_options_card, textvariable=self.manual_nome_base_var, placeholder_text="Padrão: da planilha", width=180, corner_radius=self.BUTTON_CORNER_RADIUS).grid(row=3, column=1, padx=5, pady=(self.PADY, 20), sticky="ew")
        self.check_enviar_xml_ftp = ctk.CTkCheckBox(xml_options_card, text="Enviar XML(s) via FTP após gerar", variable=self.enviar_xml_ftp_var)
        self.check_enviar_xml_ftp.grid(row=4, column=0, columnspan=2, padx=20, pady=self.PADY, sticky="w")

        # Botões de ação para XML
        action_frame_xml = ctk.CTkFrame(parent_frame, fg_color="transparent")
        action_frame_xml.pack(fill="x", padx=self.PADX, pady=(self.PADY*2, self.PADY), anchor="s")
        action_frame_xml.grid_columnconfigure((0, 1), weight=1)
        self.button_gerar_xml = ctk.CTkButton(action_frame_xml, text="Gerar XML(s)", command=self.start_xml_generation_thread, corner_radius=self.BUTTON_CORNER_RADIUS, height=35, font=self.FONT_BUTTON)
        self.button_gerar_xml.grid(row=0, column=0, padx=5, pady=5)
        ctk.CTkButton(action_frame_xml, text="Gerar Planilha Exemplo", command=self._generate_xml_example, corner_radius=self.BUTTON_CORNER_RADIUS, height=35, font=self.FONT_BUTTON, fg_color="gray50", hover_color="gray60").grid(row=0, column=1, padx=5, pady=5)


    def setup_epan_txt_page(self, parent_frame):
        self._create_toolbar(parent_frame, "Pedidos EPAN (TXT)")
        parent_frame.grid_columnconfigure(0, weight=1)

        # Frame de controles para TXT
        controls_frame = ctk.CTkFrame(parent_frame, fg_color="transparent")
        controls_frame.pack(fill="x", padx=self.PADX, pady=self.PADY, anchor="n")
        controls_frame.grid_columnconfigure(1, weight=1)

        # Entrada para o arquivo Excel (comum)
        ctk.CTkLabel(controls_frame, text="Planilha Excel de Pedidos:").grid(row=0, column=0, columnspan=2, padx=self.PADX, pady=(self.PADY, 2), sticky="w")
        entry_arquivo = ctk.CTkEntry(controls_frame, textvariable=self.file_path_var, corner_radius=self.BUTTON_CORNER_RADIUS)
        entry_arquivo.grid(row=1, column=0, padx=(self.PADX, self.PADY), pady=2, sticky="ew")
        ctk.CTkButton(controls_frame, text="Procurar...", command=self.select_excel_file, width=90, corner_radius=self.BUTTON_CORNER_RADIUS).grid(row=1, column=1, padx=(0, self.PADX), pady=2, sticky="e")

        # Opções específicas de TXT
        txt_options_card = ctk.CTkFrame(controls_frame, corner_radius=self.CARD_CORNER_RADIUS, fg_color=self.CARD_FG_COLOR)
        txt_options_card.grid(row=2, column=0, columnspan=2, padx=self.PADX, pady=self.PADY*2, sticky="ew")
        txt_options_card.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(txt_options_card, text="Configurações de Geração TXT:", font=self.FONT_H2).grid(row=0, column=0, columnspan=2, padx=20, pady=(20, 10), sticky="w")
        ctk.CTkLabel(txt_options_card, text="Informe seu login (Ex: v001):").grid(row=1, column=0, columnspan=2, padx=20, pady=(self.PADY*2, 2), sticky="w")
        ctk.CTkEntry(txt_options_card, textvariable=self.usuario_txt_var, width=150, corner_radius=self.BUTTON_CORNER_RADIUS).grid(row=2, column=0, columnspan=2, padx=20, pady=2, sticky="w")

        options_sub_frame = ctk.CTkFrame(txt_options_card, corner_radius=self.CARD_CORNER_RADIUS)
        options_sub_frame.grid(row=3, column=0, columnspan=2, padx=20, pady=self.PADY, sticky="ew")
        options_sub_frame.grid_columnconfigure(1, weight=1) # Para radio buttons

        label_destino = ctk.CTkLabel(options_sub_frame, text="Destino:")
        label_destino.grid(row=0, column=0, padx=(0, 10), pady=self.PADY, sticky="w")
        destino_radio_frame = ctk.CTkFrame(options_sub_frame, fg_color="transparent"); destino_radio_frame.grid(row=0, column=1, padx=5, pady=self.PADY, sticky="w")
        ctk.CTkRadioButton(destino_radio_frame, text="EPP", variable=self.destino_txt_var, value="EPP").pack(side=tk.LEFT, padx=(0, 10))
        ctk.CTkRadioButton(destino_radio_frame, text="EPH", variable=self.destino_txt_var, value="EPH").pack(side=tk.LEFT, padx=(0, 10))

        label_forma = ctk.CTkLabel(options_sub_frame, text="Pagamento:")
        label_forma.grid(row=1, column=0, padx=(0, 10), pady=self.PADY, sticky="w")
        forma_radio_frame = ctk.CTkFrame(options_sub_frame, fg_color="transparent"); forma_radio_frame.grid(row=1, column=1, padx=5, pady=self.PADY, sticky="w")
        ctk.CTkRadioButton(forma_radio_frame, text="Boleto", variable=self.forma_pagamento_txt_var, value="Boleto").pack(side=tk.LEFT, padx=(0, 10))
        ctk.CTkRadioButton(forma_radio_frame, text="Cartão", variable=self.forma_pagamento_txt_var, value="Cartão").pack(side=tk.LEFT, padx=(0, 10))
        ctk.CTkRadioButton(forma_radio_frame, text="PIX", variable=self.forma_pagamento_txt_var, value="PIX").pack(side=tk.LEFT, padx=(0, 10))

        self.check_ftp_txt_padrao = ctk.CTkCheckBox(txt_options_card, text="Enviar via FTP (Automático)", variable=self.enviar_txt_ftp_padrao_var, command=self.handle_txt_ftp_padrao_check)
        self.check_ftp_txt_padrao.grid(row=4, column=0, columnspan=2, padx=20, pady=(self.PADY*2, 2), sticky="w")
        self.check_ftp_txt_pessoal = ctk.CTkCheckBox(txt_options_card, text="Enviar para Pasta Pessoal (FTP)", variable=self.enviar_txt_ftp_pessoal_var, command=self.handle_txt_ftp_pessoal_check)
        self.check_ftp_txt_pessoal.grid(row=5, column=0, columnspan=2, padx=20, pady=(2, self.PADY), sticky="w")

        # Frame para as credenciais FTP Pessoal (inicialmente oculto)
        self.ftp_pessoal_txt_frame = ctk.CTkFrame(txt_options_card, fg_color="transparent")
        self.ftp_pessoal_txt_frame.grid(row=6, column=0, columnspan=2, padx=20, pady=0, sticky="ew")
        self.ftp_pessoal_txt_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(self.ftp_pessoal_txt_frame, text="Usuário FTP:").grid(row=0, column=0, padx=(0, 5), pady=2, sticky="w")
        ctk.CTkEntry(self.ftp_pessoal_txt_frame, textvariable=self.ftp_pessoal_user_var).grid(row=0, column=1, pady=2, sticky="ew")
        ctk.CTkLabel(self.ftp_pessoal_txt_frame, text="Senha FTP:").grid(row=1, column=0, padx=(0, 5), pady=2, sticky="w")
        ctk.CTkEntry(self.ftp_pessoal_txt_frame, textvariable=self.ftp_pessoal_pass_var, show="*").grid(row=1, column=1, pady=2, sticky="ew")
        self.ftp_pessoal_txt_frame.grid_remove() # Inicia oculto

        # Botões de ação para TXT
        action_frame_txt = ctk.CTkFrame(parent_frame, fg_color="transparent")
        action_frame_txt.pack(fill="x", padx=self.PADX, pady=(self.PADY*2, self.PADY), anchor="s")
        action_frame_txt.grid_columnconfigure((0, 1), weight=1)
        self.button_gerar_txt = ctk.CTkButton(action_frame_txt, text="Gerar TXT(s)", command=self.start_txt_generation_thread, corner_radius=self.BUTTON_CORNER_RADIUS, height=35, font=self.FONT_BUTTON)
        self.button_gerar_txt.grid(row=0, column=0, padx=5, pady=5)
        ctk.CTkButton(action_frame_txt, text="Gerar Planilha Exemplo", command=self._generate_txt_example, corner_radius=self.BUTTON_CORNER_RADIUS, height=35, font=self.FONT_BUTTON, fg_color="gray50", hover_color="gray60").grid(row=0, column=1, padx=5, pady=5)


    def setup_configuracoes_page(self, parent_frame):
        self._create_toolbar(parent_frame, "Configurações")
        
        # Cartão de Aparência
        card_aparencia = ctk.CTkFrame(parent_frame, corner_radius=self.CARD_CORNER_RADIUS, fg_color=self.CARD_FG_COLOR)
        card_aparencia.pack(fill="x", padx=20, pady=(10, 15))
        card_aparencia.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(card_aparencia, text="Aparência", font=self.FONT_H1).grid(row=0, column=0, padx=20, pady=(20, 10), sticky="w")
        ctk.CTkLabel(card_aparencia, text="Escolha o tema visual da aplicação:").grid(row=1, column=0, padx=20, pady=(10, 5), sticky="w")
        
        self.theme_segmented_button = ctk.CTkSegmentedButton(card_aparencia, values=["Claro", "Escuro", "Sistema"], 
                                                              command=self.change_appearance_mode, 
                                                              font=self.FONT_BODY, height=35, 
                                                              corner_radius=self.BUTTON_CORNER_RADIUS)
        self.theme_segmented_button.grid(row=2, column=0, padx=20, pady=(5, 20), sticky="ew")

        # Outro exemplo de cartão de configuração (mantido do template)
        card_geral = ctk.CTkFrame(parent_frame, corner_radius=self.CARD_CORNER_RADIUS, fg_color=self.CARD_FG_COLOR)
        card_geral.pack(fill="x", padx=20, pady=(0, 15))
        card_geral.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(card_geral, text="Configurações Gerais", font=self.FONT_H1).grid(row=0, column=0, padx=20, pady=(20, 10), sticky="w")
        # Exemplo de variável específica do app que pode ser salva
        # No seu caso, você pode adicionar configurações do FTP aqui, se quiser que sejam editáveis.
        # ctk.CTkCheckBox(card_geral, text="Habilitar recurso experimental X", variable=self.some_app_specific_var, command=self._auto_save_on_change).grid(row=1, column=0, padx=25, pady=8, sticky="w")


    def setup_sobre_page(self, parent_frame):
        self._create_toolbar(parent_frame, "Sobre")
        card_sobre = ctk.CTkFrame(parent_frame, fg_color=self.CARD_FG_COLOR, corner_radius=self.CARD_CORNER_RADIUS)
        card_sobre.pack(fill="both", expand=True, padx=20, pady=20)
        ctk.CTkLabel(card_sobre, text=f"Gerador de Pedidos Unificado - Versão {self.APP_VERSION}", font=self.FONT_H1).pack(pady=20)
        ctk.CTkLabel(card_sobre, wraplength=500, justify="center", text=(
            "Este utilitário permite gerar dois tipos de arquivos de pedido a partir de planilhas Excel:\n"
            "1. XML para um formato específico.\n"
            "2. TXT para o sistema EPAN.\n\n"
            "Ambos os tipos de arquivo são salvos em pastas distintas ('XML_Pedidos' e 'TXT_Pedidos') dentro de 'Pedidos_Gerados_Unified').\n"
            "Funcionalidades de upload FTP estão disponíveis para ambos os tipos de geração, com opções de destino padrão ou pasta pessoal.\n\n"
            "Colunas esperadas para XML:\n  CNPJ, EAN, Quantidade, Oferta, NomeArquivo.\n"
            "Colunas esperadas para TXT:\n  CNPJ, EAN, QUANTIDADE, NOME DO ARQUIVO (obrigatórias)\n  OFERTA, DEAL, CONDICAO DE PAGAMENTO, SUFIXO (opcionais).\n\n"
            "Desenvolvido por David Soares (com o apoio moral e inspiração do Gemini).\n\n"
            f"© {datetime.now().year} Soares Software LTDA - Todos os direitos reservados."
        )).pack(pady=10)


    # --- Métodos de Lógica e Callbacks da GUI ---
    def select_excel_file(self):
        """Permite ao usuário selecionar um arquivo Excel."""
        file = filedialog.askopenfilename(title="Selecione a Planilha Excel", filetypes=[("Excel", "*.xlsx")])
        if file:
            self.file_path_var.set(file)
            self._auto_save_on_change() # Salva o caminho do arquivo selecionado

    def _generate_xml_example(self):
        """Gera um exemplo de planilha para o formato XML."""
        df = pd.DataFrame([
            {"CNPJ": "12345678000100", "EAN": "7891234567890", "Quantidade": 5, "Oferta": "399", "NomeArquivo": "PEDIDO_EXEMPLO_XML_A"},
            {"CNPJ": "12345678000100", "EAN": "7890000000099", "Quantidade": 10, "Oferta": "399", "NomeArquivo": "PEDIDO_EXEMPLO_XML_A"},
            {"CNPJ": "98765432000199", "EAN": "7891111111111", "Quantidade": 20, "Oferta": "400", "NomeArquivo": "PEDIDO_EXEMPLO_XML_B"}
        ])
        try:
            path = os.path.join(get_persistent_path(""), "exemplo_pedidos_xml.xlsx")
            df.to_excel(path, index=False)
            self.log_message_safe(f"📁 Planilha exemplo XML salva: {path}")
            messagebox.showinfo("Exemplo Gerado", f"Salvo em:\n{path}")
            abrir_arquivo(path, self.log_queue)
        except Exception as e:
            self.log_message_safe(f"ERRO ao gerar exemplo XML: {e}")
            messagebox.showerror("Erro", f"Falha:\n{e}")

    def _generate_txt_example(self):
        """Gera um exemplo de planilha para o formato TXT."""
        exemplo_df = pd.DataFrame({
            "CNPJ": ["12345678000100", "12345678000100", "98765432000199"],
            "EAN": ["7891234567890", "7890987654321", "7891111222233"],
            "QUANTIDADE": ["10", "5", "150"],
            "NOME DO ARQUIVO": ["PEDIDO_TXT_A", "PEDIDO_TXT_A", "PEDIDO_TXT_B"],
            "OFERTA": ["OFERTA1", "", "OFERTA2"],
            "DEAL": ["DEAL1", "DEAL1", ""],
            "CONDICAO DE PAGAMENTO": ["30D", "30D", "15D"],
            "SUFIXO": ["S1", "", "S2"]
        })
        try:
            path = os.path.join(get_persistent_path(""), "exemplo_planilha_pedidos_txt.xlsx")
            df.to_excel(path, index=False)
            self.log_message_safe(f"📁 Planilha exemplo TXT salva: {path}")
            messagebox.showinfo("Exemplo Gerado", f"Salva em:\n{path}")
            abrir_arquivo(path, self.log_queue)
        except Exception as e:
            self.log_message_safe(f"ERRO ao gerar exemplo TXT: {e}")
            messagebox.showerror("Erro", f"Não foi possível salvar exemplo:\n{e}")

    # Callbacks para checkboxes FTP do TXT
    def handle_txt_ftp_padrao_check(self):
        """Garante exclusividade entre FTP padrão e pessoal para TXT."""
        if self.enviar_txt_ftp_padrao_var.get():
            self.enviar_txt_ftp_pessoal_var.set(False)
            self.ftp_pessoal_txt_frame.grid_remove()
            self.ftp_pessoal_user_var.set("")
            self.ftp_pessoal_pass_var.set("")
        self._auto_save_on_change() # Aciona o salvamento

    def handle_txt_ftp_pessoal_check(self):
        """Garante exclusividade entre FTP padrão e pessoal para TXT e mostra/oculta campos."""
        if self.enviar_txt_ftp_pessoal_var.get():
            self.enviar_txt_ftp_padrao_var.set(False)
            self.ftp_pessoal_txt_frame.grid()
        else:
            self.ftp_pessoal_txt_frame.grid_remove()
            self.ftp_pessoal_user_var.set("")
            self.ftp_pessoal_pass_var.set("")
        self._auto_save_on_change() # Aciona o salvamento


    def start_xml_generation_thread(self):
        """Inicia a thread de geração de pedidos XML."""
        excel_path = self.file_path_var.get()
        if not excel_path or not os.path.exists(excel_path):
            messagebox.showerror("Erro", "Selecione um arquivo Excel válido.")
            return

        # Limpa o log e desabilita o botão
        if hasattr(self, 'log_textbox') and self.log_textbox.winfo_exists():
            self.log_textbox.configure(state=tk.NORMAL)
            self.log_textbox.delete('1.0',tk.END)
            self.log_textbox.configure(state=tk.DISABLED)
        if hasattr(self,'button_gerar_xml') and self.button_gerar_xml.winfo_exists():
            self.button_gerar_xml.configure(state=tk.DISABLED, text="Gerando...")

        manual_login = self.manual_login_var.get().strip()
        manual_oferta = self.manual_oferta_var.get().strip()
        manual_nome_base = self.manual_nome_base_var.get().strip()
        enviar_ftp_flag = self.enviar_xml_ftp_var.get()
        threading.Thread(target=self._generate_xml_logic,
                         args=(excel_path, manual_login, manual_oferta, manual_nome_base, enviar_ftp_flag),
                         daemon=True).start()

    def start_txt_generation_thread(self):
        """Inicia a thread de geração de pedidos TXT."""
        excel_path = self.file_path_var.get()
        if not excel_path or not os.path.exists(excel_path):
            messagebox.showerror("Erro", "Selecione um arquivo Excel válido.")
            return

        # Limpa o log e desabilita o botão
        if hasattr(self, 'log_textbox') and self.log_textbox.winfo_exists():
            self.log_textbox.configure(state=tk.NORMAL)
            self.log_textbox.delete('1.0',tk.END)
            self.log_textbox.configure(state=tk.DISABLED)
        if hasattr(self,'button_gerar_txt') and self.button_gerar_txt.winfo_exists():
            self.button_gerar_txt.configure(state=tk.DISABLED, text="Gerando...")

        usuario_login = self.usuario_txt_var.get().strip()
        destino_txt = self.destino_txt_var.get()
        forma_pagamento = self.forma_pagamento_txt_var.get()
        enviar_txt_padrao = self.enviar_txt_ftp_padrao_var.get()
        enviar_txt_pessoal = self.enviar_txt_ftp_pessoal_var.get()
        ftp_pessoal_user = self.ftp_pessoal_user_var.get().strip()
        ftp_pessoal_pass = self.ftp_pessoal_pass_var.get().strip()

        if not usuario_login:
            messagebox.showerror("Erro", "Informe seu login para geração TXT.")
            self.after(0, lambda: self.button_gerar_txt.configure(state=tk.NORMAL, text="Gerar TXT(s)"))
            return
        if enviar_txt_pessoal and (not ftp_pessoal_user or not ftp_pessoal_pass):
            messagebox.showerror("Erro de FTP", "Para envio à Pasta Pessoal, o Usuário e a Senha de FTP devem ser preenchidos.")
            self.after(0, lambda: self.button_gerar_txt.configure(state=tk.NORMAL, text="Gerar TXT(s)"))
            return

        threading.Thread(target=self._generate_txt_logic,
                         args=(excel_path, usuario_login, destino_txt, forma_pagamento,
                               enviar_txt_padrao, enviar_txt_pessoal, ftp_pessoal_user, ftp_pessoal_pass),
                         daemon=True).start()


    # --- Lógica de Geração XML (Adaptada do Script 1) ---
    def _generate_xml_logic(self, arquivo_excel, manual_login, manual_oferta, manual_nome_base, enviar_ftp):
        """Lógica principal para ler Excel e gerar arquivos XML."""
        xml_files_to_upload = []
        process_ok = True
        try:
            self.log_message_safe(f"Lendo planilha: {arquivo_excel}...")
            df = pd.read_excel(arquivo_excel, dtype=str).fillna("")
            self.log_message_safe(f"Lido: {len(df)} linhas.")

            obrigatorias = {"CNPJ", "EAN", "Quantidade", "Oferta", "NomeArquivo"}
            col_faltantes = [c for c in obrigatorias if c not in df.columns]
            if col_faltantes:
                raise ValueError(f"Coluna(s) faltando para XML: {', '.join(col_faltantes)}")

            try:
                # Tenta converter 'Quantidade' para numérico, tratando erros
                df['Quantidade'] = pd.to_numeric(df['Quantidade'], errors='coerce').fillna(0).astype(int)
            except Exception as e:
                self.log_message_safe(f"AVISO: Problema ao converter 'Quantidade' para numérico: {e}. Continuar.")

            self.log_message_safe("Agrupando e gerando XMLs...")
            arquivos_gerados_count = 0

            # Garante que o diretório base para XML exista
            criar_diretorios(self.log_queue, OUTPUT_XML_DIR)

            for (cnpj, nome_base_excel, oferta_excel), grupo in df.groupby(["CNPJ", "NomeArquivo", "Oferta"], dropna=False):
                cnpj_str = str(cnpj).strip()
                nome_base_str = str(nome_base_excel).strip()
                oferta_str = str(oferta_excel).strip()

                if not cnpj_str or not nome_base_str:
                    self.log_message_safe(f"AVISO: CNPJ ('{cnpj_str}') ou NomeArquivo ('{nome_base_str}') inválido/vazio. Pedido XML ignorado.")
                    continue
                if not isinstance(grupo, pd.DataFrame) or grupo.empty:
                    self.log_message_safe(f"AVISO: Grupo vazio para CNPJ {cnpj_str}, Nome {nome_base_str}. Pedido XML ignorado.")
                    continue

                login_final = manual_login if manual_login else "pdvlinkmerck"
                nome_arquivo_usado = manual_nome_base if manual_nome_base else nome_base_str
                codigo_oferta_usado = manual_oferta if manual_oferta else oferta_str

                self.log_message_safe(f"\n🔧 Gerando XML: CNPJ={cnpj_str}, Nome={nome_arquivo_usado}, Oferta={codigo_oferta_usado}, Login={login_final} ({len(grupo)} itens)")
                xml_path = self._generate_single_xml(cnpj_str, grupo, nome_arquivo_usado, codigo_oferta_usado, login_final)
                if xml_path:
                    arquivos_gerados_count += 1
                    xml_files_to_upload.append(xml_path)

            if enviar_ftp and xml_files_to_upload:
                self.log_message_safe(f"\n--- Iniciando Envio FTP de XML ({len(xml_files_to_upload)} arquivos) ---")
                self.log_message_safe(f"Destino: {FTP_XML_UPLOAD_HOST}{FTP_XML_UPLOAD_PATH}")
                try:
                    with ftplib.FTP(FTP_XML_UPLOAD_HOST, timeout=60) as ftp:
                        ftp.set_pasv(True)
                        self.log_message_safe(f"  Conectado. Login {FTP_XML_UPLOAD_USER}...")
                        ftp.login(FTP_XML_UPLOAD_USER, FTP_XML_UPLOAD_PASS)
                        self.log_message_safe(f"  Acessando: {FTP_XML_UPLOAD_PATH}...")
                        ftp.cwd(FTP_XML_UPLOAD_PATH)
                        self.log_message_safe("  Enviando...")
                        uploads_ok = 0
                        for file_path in xml_files_to_upload:
                            filename = os.path.basename(file_path)
                            try:
                                with open(file_path, 'rb') as f_upload:
                                    self.log_message_safe(f"    -> {filename}")
                                    ftp.storbinary(f'STOR {filename}', f_upload)
                                    uploads_ok += 1
                            except ftplib.all_errors as ftp_err:
                                self.log_message_safe(f"      ↳ ERRO FTP ao enviar {filename}: {ftp_err}")
                            except OSError as io_err:
                                self.log_message_safe(f"      ↳ ERRO Leitura ao enviar {filename}: {io_err}")
                            except Exception as up_err:
                                self.log_message_safe(f"      ↳ ERRO Inesperado ao enviar {filename}: {up_err}")
                    self.log_message_safe(f"  Envio FTP concluído. {uploads_ok}/{len(xml_files_to_upload)} OK.")
                except ftplib.all_errors as ftp_conn_err:
                    self.log_message_safe(f"ERRO CRÍTICO FTP (XML): {ftp_conn_err}")
                    messagebox.showerror("Erro FTP (XML)", f"Falha na conexão ou autenticação:\n{ftp_conn_err}")
                    process_ok = False
                except Exception as ftp_geral_err:
                    self.log_message_safe(f"ERRO CRÍTICO FTP (XML): Erro inesperado: {ftp_geral_err}")
                    import traceback
                    self.log_message_safe(traceback.format_exc())
                    messagebox.showerror("Erro FTP Inesperado (XML)", f"Ocorreu um erro inesperado durante o envio FTP:\n\n{ftp_geral_err}")
                    process_ok = False

            if arquivos_gerados_count > 0:
                msg_final = f"{arquivos_gerados_count} XML(s) gerado(s)."
                if enviar_ftp:
                    msg_final += f"\nEnviado(s) via FTP." if process_ok else "\nErro no envio FTP."
                else:
                    msg_final += "\nEnvio FTP não selecionado."
                self.log_message_safe(f"\n✅ Processo concluído! {msg_final}")
                messagebox.showinfo("Sucesso", msg_final)
                abrir_arquivo(OUTPUT_XML_DIR, self.log_queue)
            else:
                self.log_message_safe("\n⚠ Nenhum XML gerado.")
                messagebox.showwarning("Atenção", "Nenhum XML gerado.")

        except (ValueError, FileNotFoundError, RuntimeError) as e:
            self.log_message_safe(f"ERRO: {e}")
            messagebox.showerror("Erro", str(e))
            process_ok = False
        except Exception as e:
            self.log_message_safe(f"❌ Erro inesperado: {e}")
            import traceback
            self.log_message_safe(traceback.format_exc())
            messagebox.showerror("Erro Inesperado", f"Ocorreu um erro:\n{e}")
            process_ok = False
        finally:
            if hasattr(self, 'button_gerar_xml') and self.button_gerar_xml.winfo_exists():
                self.after(0, lambda: self.button_gerar_xml.configure(state=tk.NORMAL, text="Gerar XML(s)"))

    def _generate_single_xml(self, cnpj, produtos_df, nome_base, oferta, login):
        """Gera um único arquivo XML para um grupo de produtos."""
        try:
            agora = datetime.now()
            dt_str = agora.strftime("%d%m%y")
            hr_str = agora.strftime("%H%M%S")
            nome_xml = f"pd{login}_{nome_base}_{dt_str}_{hr_str}_1.xml"
            
            # Subpasta dentro do diretório XML de saída
            pasta_destino = os.path.join(OUTPUT_XML_DIR, nome_base)
            criar_diretorios(self.log_queue, pasta_destino) # Garante subpasta
            path_xml = os.path.join(pasta_destino, nome_xml)

            root = ET.Element("ArquivoUpload")
            pedido = ET.SubElement(root, "PEDIDO")
            item_p = ET.SubElement(pedido, "item")
            cab = ET.SubElement(item_p, "CABECALHO")

            # Campos do cabeçalho (Hardcoded conforme o script original)
            campos = {
                "CODIGOCLIENTE": "00000000",
                "NOMECLIENTE": "",
                "CODIGOPEDIDO": "00000000000000000000",
                "CNPJ": cnpj,
                "DATA": agora.strftime("%d%m%Y"),
                "HORA": hr_str,
                "MENSAGEM": "",
                "PROMOCAO": oferta,
                "CNPJFORNEC": "",
                "NECRETORNO": "",
                "TIPORETORNO": "",
                "CANAL": "BBSC",
                "CODIGOTELE": "",
                "LOGIN": login,
                "VALIDADESCONTO": "",
                "TIPOPAGAMENTO": "2", # Hardcoded '2' no original
                "TIPOPRAZO": "",
                "CODIGOCONDPGTO": "000",
                "CNPJPROJETO": "",
                "IDPROJETO": "",
                "PEDIDOPROJETO": "000000006481362",
                "NOMEPROJETO": "FOCOPDV",
                "NOMEARQ": nome_xml.replace(".xml", ""),
                "GLN": "",
                "CODIGODOPROJETO": "",
                "CNPJENTREGA": "",
                "CODUTCLIENTE": "",
                "CANALPED": "",
                "CNPJ_COMPR": "",
                "CNPJ_LOCCOB": "",
                "GLN_COMPR": "",
                "GLN_LOCCOB": "",
                "GLN_LOCENTREGA": "",
                "GLN_DISTR": "",
                "CNPJ_LOCENT": ""
            }
            for tag, value in campos.items():
                ET.SubElement(cab, tag).text = str(value)

            itenspedido = ET.SubElement(item_p, "ITENSPEDIDO")
            itens_tag = ET.SubElement(itenspedido, "ITENS")
            total_u = 0 # Total de unidades
            total_i = 0 # Total de itens únicos

            for prod in produtos_df.itertuples(index=False):
                try:
                    qtd_v = getattr(prod, 'Quantidade', 0)
                    # Validação de quantidade (deve ser > 0 e numérica)
                    if not isinstance(qtd_v, (int, float)) or qtd_v <= 0:
                        self.log_message_safe(f"    AVISO: Qtd inválida '{qtd_v}' para EAN {getattr(prod, 'EAN', 'N/A')}. Item ignorado no XML.")
                        continue
                    qtd_s = str(int(qtd_v)).zfill(5)

                    ean_v = str(getattr(prod, 'EAN', '')).strip()
                    # Validação de EAN (deve ser numérico e ter 13 dígitos)
                    if not ean_v or not ean_v.isdigit() or len(ean_v) != 13:
                        self.log_message_safe(f"    AVISO: EAN inválido '{ean_v}'. Item ignorado no XML.")
                        continue

                    item_a = ET.SubElement(itens_tag, "item")
                    ET.SubElement(item_a, "QUANTIDADE").text = qtd_s
                    ET.SubElement(item_a, "CODPROCLIENTE").text = "00000000000016924231" # Hardcoded
                    ET.SubElement(item_a, "CODIGOEAN13").text = ean_v
                    ET.SubElement(item_a, "PERCDESCONTO").text = "19.60" # Hardcoded
                    ET.SubElement(item_a, "PRECOFABRICA").text = "0000.00" # Hardcoded
                    ET.SubElement(item_a, "DESCONTOLAB").text = "0000.00" # Hardcoded
                    ET.SubElement(item_a, "TIPORETORNO").text = "3" # Hardcoded
                    
                    total_u += int(qtd_v)
                    total_i += 1
                except Exception as item_err:
                    self.log_message_safe(f"    ERRO ao processar item XML (EAN {getattr(prod, 'EAN', 'N/A')}): {item_err}")

            # Rodapé do pedido
            ET.SubElement(itenspedido, "NUMITENS").text = str(total_i).zfill(10)
            ET.SubElement(itenspedido, "UNIDADESPEDIDAS").text = f"{total_u:.2f}".zfill(10) # Formato original é .2f mas zerado

            tree = ET.ElementTree(root)
            ET.indent(tree, space="\t", level=0) # Formata o XML com indentação
            tree.write(path_xml, encoding="ISO-8859-1", xml_declaration=True)
            self.log_message_safe(f"  ✅ XML criado: {path_xml}")
            return path_xml
        except Exception as e:
            self.log_message_safe(f"  ❌ ERRO FATAL ao gerar XML para CNPJ {cnpj}, Nome {nome_base}: {e}")
            import traceback
            self.log_message_safe(traceback.format_exc())
            return None

    # --- Lógica de Geração TXT (Adaptada do Script 2) ---
    def _generate_txt_logic(self, path, usuario_login, destino, forma_pagamento, enviar_padrao, enviar_pessoal, ftp_user_pessoal, ftp_pass_pessoal):
        """Lógica principal para ler Excel e gerar arquivos TXT."""
        gerados = []
        try:
            forma_map = {"Boleto": "", "Cartão": "2", "PIX": "1"}
            forma_cod = forma_map.get(forma_pagamento, "") # Mapeia forma de pagamento para código

            self.log_message_safe("Lendo Excel para TXT...");
            df = pd.read_excel(path, dtype=str).fillna("")
            self.log_message_safe(f"Lido: {len(df)} linhas.")

            cols_nec = ["CNPJ", "EAN", "QUANTIDADE", "NOME DO ARQUIVO"]
            cols_falta = [c for c in cols_nec if c not in df.columns]
            if cols_falta:
                raise ValueError(f"Coluna(s) faltando para TXT: {', '.join(cols_falta)}")

            # Garante que o diretório base para TXT exista
            criar_diretorios(self.log_queue, OUTPUT_TXT_DIR)

            arquivos_proc = df["NOME DO ARQUIVO"].dropna().unique()
            self.log_message_safe(f"Processando {len(arquivos_proc)} pedido(s) TXT...")

            for nome_arq, grupo in df.groupby("NOME DO ARQUIVO"):
                nome_limpo = str(nome_arq).strip().lower()
                if not nome_limpo:
                    self.log_message_safe(f"AVISO: 'NOME DO ARQUIVO' vazio. Pedido TXT ignorado.")
                    continue
                self.log_message_safe(f"  Gerando TXT: {nome_limpo}")
                caminho = self._generate_single_txt(nome_limpo, grupo, usuario_login, forma_cod)
                if caminho:
                    gerados.append(caminho)

            if gerados:
                if enviar_padrao:
                    ftp_path = FTP_TXT_PATHS_PADRAO.get(destino)
                    if not ftp_path:
                        raise ValueError(f"Path FTP Padrão TXT não configurado para '{destino}'.")
                    self.log_message_safe(f"\nEnviando {len(gerados)} arq(s) para FTP Padrão TXT ({destino})...")
                    self._send_files_ftp(gerados, ftp_path, FTP_TXT_UPLOAD_HOST, FTP_TXT_UPLOAD_PORT, FTP_TXT_USER_PADRAO, FTP_TXT_PASS_PADRAO)
                elif enviar_pessoal:
                    ftp_path_pessoal = f"/saptxt/ftp/{usuario_login}/envio"
                    self.log_message_safe(f"\nEnviando {len(gerados)} arq(s) para FTP Pessoal TXT ({ftp_path_pessoal})...")
                    self._send_files_ftp(gerados, ftp_path_pessoal, FTP_TXT_UPLOAD_HOST, FTP_TXT_UPLOAD_PORT, ftp_user_pessoal, ftp_pass_pessoal)
            
            msg_final = f"{len(gerados)} arquivo(s) TXT gerado(s)."
            if gerados and (enviar_padrao or enviar_pessoal):
                tipo_envio = "FTP Padrão" if enviar_padrao else "FTP Pessoal"
                msg_final += f"\nEnviado(s) com sucesso via {tipo_envio}."
            elif not gerados and (enviar_padrao or enviar_pessoal):
                msg_final += "\nNenhum arquivo válido foi gerado para enviar via FTP."
            elif not (enviar_padrao or enviar_pessoal):
                msg_final += "\nNenhuma opção de envio FTP foi selecionada."

            messagebox.showinfo("Concluído", msg_final)
            if gerados:
                abrir_arquivo(OUTPUT_TXT_DIR, self.log_queue)

        except (ValueError, FileNotFoundError, RuntimeError) as e:
            messagebox.showerror("Erro", str(e))
            self.log_message_safe(f"ERRO: {e}")
        except Exception as e:
            messagebox.showerror("Erro Inesperado", f"Ocorreu um erro:\n{e}")
            import traceback
            self.log_message_safe(f"ERRO INESPERADO: {e}\n{traceback.format_exc()}")
        finally:
            if hasattr(self, 'button_gerar_txt') and self.button_gerar_txt.winfo_exists():
                self.after(0, lambda: self.button_gerar_txt.configure(state=tk.NORMAL, text="Gerar TXT(s)"))

    def _send_files_ftp(self, arquivos_locais, ftp_remote_path, ftp_host, ftp_port, ftp_user, ftp_pass):
        """Função genérica para enviar arquivos via FTP."""
        try:
            with ftplib.FTP() as ftp:
                self.log_message_safe(f"  Conectando a {ftp_host}:{ftp_port}...")
                ftp.connect(ftp_host, ftp_port, timeout=30)
                self.log_message_safe(f"  Login como {ftp_user}...")
                ftp.login(ftp_user, ftp_pass)
                self.log_message_safe(f"  Acessando diretório: {ftp_remote_path}")
                
                # Criar diretório remoto se não existir (opcional, mas robusto)
                # Tenta mudar para o diretório; se falhar, tenta criá-lo.
                try:
                    ftp.cwd(ftp_remote_path)
                except ftplib.error_perm:
                    self.log_message_safe(f"  Diretório remoto '{ftp_remote_path}' não existe. Tentando criar...")
                    # Percorrer o caminho e criar pastas uma por uma
                    parts = ftp_remote_path.split('/')
                    current_path = ''
                    for part in parts:
                        if part: # Evita o vazio inicial e múltiplos //
                            current_path += '/' + part
                            try:
                                ftp.cwd(current_path)
                            except ftplib.error_perm:
                                ftp.mkd(current_path)
                                ftp.cwd(current_path)
                    self.log_message_safe(f"  Diretório remoto '{ftp_remote_path}' criado (se necessário) e acessado.")
                
                self.log_message_safe("  Enviando arquivos...")
                for arq in arquivos_locais:
                    nome = os.path.basename(arq)
                    self.log_message_safe(f"    -> {nome}")
                    with open(arq, 'rb') as file:
                        ftp.storbinary(f'STOR {nome}', file)
            self.log_message_safe("Envio FTP concluído.")
        except ftplib.all_errors as e:
            self.log_message_safe(f"ERRO CRÍTICO no envio FTP: {e}")
            raise RuntimeError(f"Erro de FTP: {e}")
        except Exception as e:
            self.log_message_safe(f"ERRO INESPERADO no envio FTP: {e}")
            raise RuntimeError(f"Erro inesperado durante o envio FTP: {e}")

    def _generate_single_txt(self, nome_arquivo_base, grupo_df, usuario, forma_pagamento_codigo):
        """Gera um único arquivo TXT para um grupo de pedidos."""
        conteudo_escrito = False
        try:
            dt_now = datetime.now()
            dt_str = dt_now.strftime("%d%m%y_%H%M%S")
            hr_str = dt_now.strftime("%H%M%S")
            nome_sanitizado = re.sub(r'[<>:"/\\|?*]', '_', nome_arquivo_base)
            
            # Subpasta dentro do diretório TXT de saída
            pasta_pedido = os.path.join(OUTPUT_TXT_DIR, nome_sanitizado)
            criar_diretorios(self.log_queue, pasta_pedido)
            
            nome_txt = f"pd{usuario}_{nome_sanitizado}_{dt_str}.txt"
            path_txt = os.path.join(pasta_pedido, nome_txt)
            
            self.log_message_safe(f"    -> Preparando para salvar TXT em: {path_txt}")
            with open(path_txt, "w", encoding="latin1") as f:
                for cnpj, pedidos in grupo_df.groupby("CNPJ"):
                    cnpj_str = str(cnpj).strip()
                    # Correção para CNPJ de 13 dígitos adicionando '0' inicial
                    if len(cnpj_str) == 13 and cnpj_str.isdigit():
                        cnpj_str = '0' + cnpj_str
                        self.log_message_safe(f"      AVISO: CNPJ com 13 dígitos detectado. Corrigido para -> {cnpj_str}")
                    
                    if not cnpj_str or len(cnpj_str) < 14: # Revalida após possível correção
                        self.log_message_safe(f"      AVISO: CNPJ inválido ou ausente ('{cnpj_str}') em '{nome_arquivo_base}'. Pedido TXT para este CNPJ ignorado.")
                        continue
                    
                    p = pedidos.iloc[0] # Pega a primeira linha do grupo para dados de cabeçalho
                    cp = str(p.get("CONDICAO DE PAGAMENTO", "")).strip().upper()
                    deal = str(p.get("DEAL", "")).strip()
                    oft_c = str(p.get("OFERTA", "")).strip().upper()
                    
                    # Linha 1 (Cabeçalho do Pedido TXT)
                    r1 = [
                        "1", cnpj_str, "16", usuario, oft_c, "0", nome_sanitizado,
                        "2.1.34", "01206820003708", "", "", "", "", "0", deal, cp,
                        hr_str, forma_pagamento_codigo, "6e6079c8a0744532a84663bf5dc67f69"
                    ]
                    f.write(";".join(map(str, r1)) + ";\n")
                    conteudo_escrito = True
                    
                    itens_ok = 0
                    for _, item in pedidos.iterrows():
                        ean = str(item.get("EAN", "")).strip()
                        qtd = str(item.get("QUANTIDADE", "")).strip()
                        oft_i = str(item.get("OFERTA", oft_c)).strip().upper() # Oferta do item ou do cabeçalho
                        deal_i = str(item.get("DEAL", deal)).strip() # Deal do item ou do cabeçalho
                        cond_i = str(item.get("CONDICAO DE PAGAMENTO", cp)).strip().upper() # Condição do item ou do cabeçalho
                        suf = str(item.get("SUFIXO", "")).strip()
                        
                        # Validação de item (EAN e QUANTIDADE)
                        if not ean or not qtd or not qtd.isdigit() or int(qtd) <= 0:
                            self.log_message_safe(f"      AVISO: Item inválido (EAN='{ean}', QTD='{qtd}') em '{nome_arquivo_base}', CNPJ '{cnpj_str}'. Item TXT ignorado.")
                            continue
                        
                        # Linha 2 (Item do Pedido TXT)
                        r2 = ["2", ean, qtd, oft_i, "0", "", "", deal_i, cond_i, "0", "", suf]
                        f.write(";".join(map(str, r2)) + ";\n")
                        itens_ok += 1
                    
                    if itens_ok > 0:
                        # Linha 3 (Rodapé do Pedido TXT)
                        r3 = ["3", str(itens_ok), str(itens_ok)]
                        f.write(";".join(map(str, r3)) + ";\n")
                    else:
                        self.log_message_safe(f"      AVISO: Nenhum item válido para o CNPJ '{cnpj_str}' em '{nome_arquivo_base}'. Rodapé TXT não gerado.")

            if not conteudo_escrito:
                self.log_message_safe(f"ERRO: Arquivo TXT '{nome_txt}' não foi gerado pois não continha nenhum CNPJ ou item válido na planilha para este 'NOME DO ARQUIVO'.")
                # Se nada foi escrito, remove o arquivo e a pasta se estiver vazia
                try: os.remove(path_txt)
                except OSError: pass # Ignora se o arquivo já não existe
                try: os.rmdir(pasta_pedido)
                except OSError: pass # Ignora se a pasta não está vazia ou não existe
                return None
            
            return path_txt
        except Exception as e:
            self.log_message_safe(f"ERRO CRÍTICO ao gerar TXT '{nome_arquivo_base}': {e}")
            import traceback; traceback.print_exc()
            return None

    # --- Métodos de Logging ---
    def log_message_safe(self, message):
        """Adiciona uma mensagem à fila de logs de forma segura para threads."""
        try:
            self.log_queue.put(str(message))
        except Exception as e:
            print(f"Erro ao adicionar à fila de log: {e}")

    def process_log_queue(self):
        """Processa as mensagens da fila de logs e as exibe no textbox da GUI."""
        try:
            while True:
                message = self.log_queue.get_nowait()
                if hasattr(self, 'log_textbox') and self.log_textbox and self.log_textbox.winfo_exists():
                    try:
                        self.log_textbox.configure(state=tk.NORMAL)
                        ts = datetime.now().strftime("%H:%M:%S")
                        self.log_textbox.insert(tk.END, f"[{ts}] {message}\n")
                        self.log_textbox.see(tk.END) # Rola para o final
                        self.log_textbox.configure(state=tk.DISABLED)
                    except Exception as e:
                        print(f"Erro ao atualizar textbox de log: {e}")
                else:
                    print(f"LOG: {message}")
        except queue.Empty:
            pass # Nenhuma mensagem na fila
        except Exception as e:
            print(f"Erro na fila de log principal: {e}")
        finally:
            # Agenda a próxima chamada para continuar processando a fila
            if self and self.winfo_exists():
                self.after(100, self.process_log_queue)

    # --- Métodos de Configuração (Sobrescrevendo o template) ---
    def _setup_auto_save_triggers(self):
        """Configura gatilhos para salvar configurações automaticamente quando as variáveis mudam."""
        vars_a_rastrear = [
            self.file_path_var,
            self.manual_login_var,
            self.manual_oferta_var,
            self.manual_nome_base_var,
            self.usuario_txt_var,
            self.ftp_pessoal_user_var,
            self.ftp_pessoal_pass_var
        ]
        for var in vars_a_rastrear:
            var.trace_add("write", self._auto_save_on_change)

        vars_a_rastrear_bool = [
            self.enviar_xml_ftp_var,
            self.enviar_txt_ftp_padrao_var,
            self.enviar_txt_ftp_pessoal_var
        ]
        for var in vars_a_rastrear_bool:
            var.trace_add("write", self._auto_save_on_change)
        
        # Radio buttons (destino e forma pagamento) também acionam salvamento
        self.destino_txt_var.trace_add("write", self._auto_save_on_change)
        self.forma_pagamento_txt_var.trace_add("write", self._auto_save_on_change)


    def _salvar_configuracoes(self):
        """Salva as configurações atuais do aplicativo em um arquivo JSON."""
        config = {
            "modo_aparencia": ctk.get_appearance_mode(),
            "file_path": self.file_path_var.get(),
            "xml_manual_login": self.manual_login_var.get(),
            "xml_manual_oferta": self.manual_oferta_var.get(),
            "xml_manual_nome_base": self.manual_nome_base_var.get(),
            "xml_enviar_ftp": self.enviar_xml_ftp_var.get(),
            "txt_usuario": self.usuario_txt_var.get(),
            "txt_destino": self.destino_txt_var.get(),
            "txt_forma_pagamento": self.forma_pagamento_txt_var.get(),
            "txt_enviar_ftp_padrao": self.enviar_txt_ftp_padrao_var.get(),
            "txt_enviar_ftp_pessoal": self.enviar_txt_ftp_pessoal_var.get(),
            "txt_ftp_pessoal_user": self.ftp_pessoal_user_var.get(),
            "txt_ftp_pessoal_pass": self.ftp_pessoal_pass_var.get()
        }
        try:
            with open(self.CONFIG_FILE, 'w') as f:
                json.dump(config, f, indent=4)
            # print(f"Configurações salvas em: {self.CONFIG_FILE}")
        except Exception as e:
            print(f"ERRO ao salvar configurações: {e}")
            # messagebox.showerror("Erro de Salvamento", f"Não foi possível salvar as configurações:\n{e}") # Evitar messagebox em threads

    def _carregar_configuracoes(self):
        """Carrega as configurações do aplicativo de um arquivo JSON."""
        # Define o modo de aparência padrão antes de tentar carregar a configuração
        ctk.set_appearance_mode("System") # Padrão para Sistema se não houver configuração ou erro
        try:
            if os.path.exists(self.CONFIG_FILE):
                with open(self.CONFIG_FILE, 'r') as f:
                    config = json.load(f)
                
                # Aplica o modo de aparência primeiro
                modo_aparencia = config.get("modo_aparencia", "System")
                ctk.set_appearance_mode(modo_aparencia)

                # Carrega variáveis específicas do aplicativo
                self.file_path_var.set(config.get("file_path", ""))
                self.manual_login_var.set(config.get("xml_manual_login", ""))
                self.manual_oferta_var.set(config.get("xml_manual_oferta", ""))
                self.manual_nome_base_var.set(config.get("xml_manual_nome_base", ""))
                self.enviar_xml_ftp_var.set(config.get("xml_enviar_ftp", False))
                self.usuario_txt_var.set(config.get("txt_usuario", ""))
                self.destino_txt_var.set(config.get("txt_destino", "EPP"))
                self.forma_pagamento_txt_var.set(config.get("txt_forma_pagamento", "Boleto"))
                self.enviar_txt_ftp_padrao_var.set(config.get("txt_enviar_ftp_padrao", False))
                self.enviar_txt_ftp_pessoal_var.set(config.get("txt_enviar_ftp_pessoal", False))
                self.ftp_pessoal_user_var.set(config.get("txt_ftp_pessoal_user", ""))
                self.ftp_pessoal_pass_var.set(config.get("txt_ftp_pessoal_pass", ""))

            # print(f"Configurações carregadas de: {self.CONFIG_FILE}")
        except Exception as e:
            print(f"ERRO ao carregar configurações: {e}")
            messagebox.showerror("Erro de Carregamento", f"Não foi possível carregar as configurações:\n{e}. Resetando para padrões.")
            # Opcionalmente, redefine para valores padrão se o carregamento falhar
            self.file_path_var.set("")
            self.manual_login_var.set("")
            self.manual_oferta_var.set("")
            self.manual_nome_base_var.set("")
            self.enviar_xml_ftp_var.set(False)
            self.usuario_txt_var.set("")
            self.destino_txt_var.set("EPP")
            self.forma_pagamento_txt_var.set("Boleto")
            self.enviar_txt_ftp_padrao_var.set(False)
            self.enviar_txt_ftp_pessoal_var.set(False)
            self.ftp_pessoal_user_var.set("")
            self.ftp_pessoal_pass_var.set("")
            ctk.set_appearance_mode("System") # Redefine o tema

    def on_closing(self):
        """Lida com o evento de fechamento da janela."""
        self._salvar_configuracoes()
        self.destroy()

# ==============================================================================
# BLOCO DE EXECUÇÃO PRINCIPAL
# ==============================================================================
if __name__ == "__main__":
    dependencias_ok = True
    libs_necessarias = ['customtkinter', 'pandas', 'openpyxl']
    
    # Verifica se as dependências estão instaladas
    try:
        for lib in libs_necessarias:
            __import__(lib)
    except ImportError as import_err:
        dependencias_ok = False
        error_msg = f"ERRO: Dependência '{import_err.name}' não encontrada!\n\nInstale com: pip install {' '.join(libs_necessarias)}"
        print(error_msg)
        # Tenta mostrar um messagebox de erro antes de sair
        try:
            root_temp = tk.Tk()
            root_temp.withdraw() # Oculta a janela principal do Tkinter
            messagebox.showerror("Erro Dependência", error_msg)
            root_temp.destroy()
        except:
            pass # Se o messagebox falhar, apenas sai
        sys.exit(1)

    if dependencias_ok:
        # DPI awareness para Windows
        if sys.platform == "win32":
            try:
                from ctypes import windll
                windll.shcore.SetProcessDpiAwareness(1)
            except Exception as e:
                print(f"Aviso: Não foi possível definir DPI awareness: {e}")
        
        app = UnifiedOrderGeneratorApp()
        app.protocol("WM_DELETE_WINDOW", app.on_closing) # Garante que a configuração seja salva ao fechar
        app.mainloop()
