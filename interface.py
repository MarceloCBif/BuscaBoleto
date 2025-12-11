"""
Interface gráfica para busca e download de boletos via SFTP.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import subprocess
import sys
import threading
import zipfile
import socket
import configparser
from datetime import datetime
from typing import Optional, List

from ftp_client import SFTPClient, get_config_path


class BuscaBoletoApp:
    """Aplicação principal para busca de boletos."""
    
    def __init__(self, root: tk.Tk):
        """
        Inicializa a aplicação.
        
        Args:
            root: Janela principal do Tkinter.
        """
        self.root = root
        self.root.title("Busca Boleto Ektech")
        self.root.geometry("900x700")
        self.root.minsize(600, 400)
        
        # Cliente SFTP
        self.ftp_client: Optional[SFTPClient] = None
        self.conectado = False
        
        # Lista de resultados da busca
        self.resultados_busca = []
        
        # Controle de extração de clientes
        self.extraindo_clientes = False
        self.cancelar_extracao = False
        
        # Carrega configurações de segurança
        self.faixas_ip_permitidas = self.carregar_faixas_ip()
        
        # Configura o estilo
        self.configurar_estilo()
        
        # Cria os widgets
        self.criar_widgets()
        
        # Centraliza a janela
        self.centralizar_janela()
    
    def configurar_estilo(self):
        """Configura o estilo visual da aplicação."""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Cores personalizadas
        style.configure('TFrame', background='#f0f0f0')
        style.configure('TLabel', background='#f0f0f0', font=('Segoe UI', 10))
        style.configure('TButton', font=('Segoe UI', 10))
        style.configure('Header.TLabel', font=('Segoe UI', 14, 'bold'))
        style.configure('Status.TLabel', font=('Segoe UI', 9))
        
        # Estilo para botões de ação
        style.configure('Accent.TButton', font=('Segoe UI', 10, 'bold'))
    
    def aplicar_mascara_data(self, event):
        """
        Aplica máscara de data DD/MM/AAAA automaticamente.
        Formata a entrada removendo caracteres não numéricos e inserindo barras.
        """
        widget = event.widget
        texto = widget.get()
        
        # Guarda a posição do cursor
        cursor_pos = widget.index(tk.INSERT)
        
        # Remove tudo que não é número
        apenas_numeros = ''.join(c for c in texto if c.isdigit())
        
        # Limita a 8 dígitos (DDMMAAAA)
        apenas_numeros = apenas_numeros[:8]
        
        # Aplica a máscara DD/MM/AAAA
        formatado = ""
        for i, char in enumerate(apenas_numeros):
            if i == 2 or i == 4:  # Adiciona / após dia e mês
                formatado += "/"
            formatado += char
        
        # Atualiza o campo apenas se mudou
        if texto != formatado:
            widget.delete(0, tk.END)
            widget.insert(0, formatado)
            
            # Ajusta a posição do cursor
            # Se adicionou uma barra antes do cursor, move o cursor para frente
            nova_pos = min(cursor_pos, len(formatado))
            
            # Conta quantas barras foram adicionadas antes da posição original
            barras_antes_original = texto[:cursor_pos].count('/')
            barras_antes_novo = formatado[:nova_pos].count('/')
            
            # Ajusta para considerar as barras adicionadas automaticamente
            if len(apenas_numeros) > len(texto.replace('/', '')):
                # Usuário digitou algo
                nova_pos = len(formatado)
            elif barras_antes_novo > barras_antes_original:
                nova_pos += (barras_antes_novo - barras_antes_original)
            
            widget.icursor(nova_pos)
    
    def centralizar_janela(self):
        """Centraliza a janela na tela."""
        # Tamanho da janela definido
        width = 900
        height = 700
        
        # Tamanho da tela
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        # Calcula a posição para centralizar
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
        
        # Garante que não fique negativo (para telas pequenas)
        x = max(0, x)
        y = max(0, y)
        
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def criar_widgets(self):
        """Cria todos os widgets da interface."""
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # === Cabeçalho ===
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(
            header_frame, 
            text="🔍 Busca de Boletos SFTP", 
            style='Header.TLabel'
        ).pack(side=tk.LEFT)
        
        # Indicador de status da conexão
        self.lbl_status_conexao = ttk.Label(
            header_frame,
            text="● Desconectado",
            foreground='red',
            style='Status.TLabel'
        )
        self.lbl_status_conexao.pack(side=tk.RIGHT, padx=10)
        
        # === Frame de Busca por Número ===
        busca_frame = ttk.LabelFrame(main_frame, text="Buscar por Número do Boleto", padding="10")
        busca_frame.pack(fill=tk.X, pady=(0, 5))
        
        # Frame para número
        entrada_frame = ttk.Frame(busca_frame)
        entrada_frame.pack(fill=tk.X, pady=5)
        
        # Campo de busca
        ttk.Label(entrada_frame, text="Número da NF:").pack(side=tk.LEFT)
        
        self.entry_numero = ttk.Entry(entrada_frame, font=('Segoe UI', 12), width=15)
        self.entry_numero.pack(side=tk.LEFT, padx=(5, 10))
        self.entry_numero.bind('<Return>', lambda e: self.buscar_boleto())
        
        self.btn_buscar = ttk.Button(
            entrada_frame,
            text="🔍 Buscar por Número",
            command=self.buscar_boleto,
            style='Accent.TButton'
        )
        self.btn_buscar.pack(side=tk.RIGHT)
        
        # Frame para checkboxes
        opcoes_frame = ttk.Frame(busca_frame)
        opcoes_frame.pack(fill=tk.X, pady=(5, 0))
        
        # Checkbox para busca literal (com máscara)
        self.var_busca_literal = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            opcoes_frame,
            text="Busca literal (número com zeros à esquerda - busca em todas as filiais)",
            variable=self.var_busca_literal
        ).pack(side=tk.LEFT, padx=(0, 20))
        
        # === Frame de Busca por Data ===
        busca_data_frame = ttk.LabelFrame(main_frame, text="Buscar por Data de Criação", padding="10")
        busca_data_frame.pack(fill=tk.X, pady=(0, 5))
        
        # Frame para campos de data
        data_frame = ttk.Frame(busca_data_frame)
        data_frame.pack(fill=tk.X, pady=5)
        
        # Data inicial
        ttk.Label(data_frame, text="De:").pack(side=tk.LEFT)
        self.entry_data_inicio = ttk.Entry(data_frame, font=('Segoe UI', 10), width=12)
        self.entry_data_inicio.pack(side=tk.LEFT, padx=(5, 15))
        self.entry_data_inicio.insert(0, datetime.now().strftime("%d/%m/%Y"))
        self.entry_data_inicio.bind('<KeyRelease>', self.aplicar_mascara_data)
        
        # Data final
        ttk.Label(data_frame, text="Até:").pack(side=tk.LEFT)
        self.entry_data_fim = ttk.Entry(data_frame, font=('Segoe UI', 10), width=12)
        self.entry_data_fim.pack(side=tk.LEFT, padx=(5, 15))
        self.entry_data_fim.insert(0, datetime.now().strftime("%d/%m/%Y"))
        self.entry_data_fim.bind('<KeyRelease>', self.aplicar_mascara_data)
        
        # Botão para data de hoje
        ttk.Button(
            data_frame,
            text="Hoje",
            command=self.definir_data_hoje,
            width=6
        ).pack(side=tk.LEFT, padx=(0, 5))
        
        # Botão para última semana
        ttk.Button(
            data_frame,
            text="Última Semana",
            command=self.definir_ultima_semana,
            width=14
        ).pack(side=tk.LEFT, padx=(0, 5))
        
        # Botão para último mês
        ttk.Button(
            data_frame,
            text="Último Mês",
            command=self.definir_ultimo_mes,
            width=10
        ).pack(side=tk.LEFT, padx=(0, 5))
        
        # Botão cancelar extração (lado direito)
        self.btn_cancelar = ttk.Button(
            data_frame,
            text="✖ Cancelar",
            command=self.cancelar_extracao_clientes,
            width=10
        )
        self.btn_cancelar.pack(side=tk.RIGHT, padx=(5, 0))
        
        # Botão buscar por data (lado direito)
        self.btn_buscar_data = ttk.Button(
            data_frame,
            text="🔍 Buscar por Data",
            command=self.buscar_por_data,
            style='Accent.TButton'
        )
        self.btn_buscar_data.pack(side=tk.RIGHT, padx=(5, 0))
        
        # === Frame de Resultados ===
        resultados_frame = ttk.LabelFrame(main_frame, text="Resultados (Clique em ☐ para marcar/desmarcar | Clique no cabeçalho para ordenar)", padding="10")
        resultados_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Treeview para mostrar resultados - seleção múltipla
        # Ordem: check, tipo, numero, cliente, data, nome, caminho (oculto)
        columns = ('check', 'tipo', 'numero', 'cliente', 'data', 'nome', 'caminho')
        self.tree_resultados = ttk.Treeview(
            resultados_frame, 
            columns=columns, 
            show='headings',
            selectmode='extended'  # Permite seleção múltipla
        )
        
        # Configurar cabeçalhos com ordenação
        self.tree_resultados.heading('check', text='☑', command=self.toggle_todos_checkboxes)
        self.tree_resultados.heading('tipo', text='Tipo ↕', command=lambda: self.ordenar_coluna('tipo', False))
        self.tree_resultados.heading('numero', text='Número ↕', command=lambda: self.ordenar_coluna('numero', False))
        self.tree_resultados.heading('cliente', text='Cliente ↕', command=lambda: self.ordenar_coluna('cliente', False))
        self.tree_resultados.heading('data', text='Data ↕', command=lambda: self.ordenar_coluna('data', False))
        self.tree_resultados.heading('nome', text='Nome do Arquivo ↕', command=lambda: self.ordenar_coluna('nome', False))
        self.tree_resultados.heading('caminho', text='')  # Oculto
        
        self.tree_resultados.column('check', width=30, anchor='center')
        self.tree_resultados.column('tipo', width=60)
        self.tree_resultados.column('numero', width=80, anchor='center')
        self.tree_resultados.column('cliente', width=250)
        self.tree_resultados.column('data', width=110)
        self.tree_resultados.column('nome', width=240)
        self.tree_resultados.column('caminho', width=0, stretch=False)  # Oculta a coluna
        
        # Configurar tags para cores diferentes
        self.tree_resultados.tag_configure('boleto', background='#B3E5FC')   # Azul claro
        self.tree_resultados.tag_configure('nf', background='#B3E5FC')       # Azul claro
        self.tree_resultados.tag_configure('separador', background='white')  # Branco
        
        # Bind para clique no checkbox
        self.tree_resultados.bind('<Button-1>', self.on_tree_click)
        
        # Scrollbar para a lista
        scrollbar = ttk.Scrollbar(
            resultados_frame, 
            orient=tk.VERTICAL, 
            command=self.tree_resultados.yview
        )
        self.tree_resultados.configure(yscrollcommand=scrollbar.set)
        
        self.tree_resultados.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Bind para duplo clique
        self.tree_resultados.bind('<Double-1>', lambda e: self.baixar_selecionado())
        
        # === Frame de Ações ===
        acoes_frame = ttk.Frame(main_frame)
        acoes_frame.pack(fill=tk.X)
        
        self.btn_baixar = ttk.Button(
            acoes_frame,
            text="⬇️ Baixar Selecionado(s)",
            command=self.baixar_selecionado,
            style='Accent.TButton'
        )
        self.btn_baixar.pack(side=tk.LEFT, padx=(0, 10))
        
        self.btn_selecionar_todos = ttk.Button(
            acoes_frame,
            text="☑️ Selecionar Todos",
            command=self.selecionar_todos
        )
        self.btn_selecionar_todos.pack(side=tk.LEFT, padx=(0, 10))
        
        self.btn_abrir_pasta = ttk.Button(
            acoes_frame,
            text="📂 Abrir Pasta de Downloads",
            command=self.abrir_pasta_downloads
        )
        self.btn_abrir_pasta.pack(side=tk.LEFT)
        
        # Label de contagem
        self.lbl_contagem = ttk.Label(
            acoes_frame,
            text="",
            style='Status.TLabel'
        )
        self.lbl_contagem.pack(side=tk.RIGHT)
        
        # === Barra de Status ===
        self.status_bar = ttk.Label(
            self.root,
            text="Pronto. Configure o arquivo config.ini e conecte ao servidor SFTP.",
            relief=tk.SUNKEN,
            anchor=tk.W,
            padding=(5, 2)
        )
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # === Barra de Progresso ===
        self.progress = ttk.Progressbar(
            self.root,
            mode='indeterminate'
        )
    
    def atualizar_status(self, mensagem: str):
        """Atualiza a barra de status."""
        self.status_bar.config(text=mensagem)
        self.root.update_idletasks()
    
    def mostrar_progresso(self, mostrar: bool = True):
        """Mostra ou esconde a barra de progresso."""
        if mostrar:
            self.progress.pack(side=tk.BOTTOM, fill=tk.X, before=self.status_bar)
            self.progress.start(10)
        else:
            self.progress.stop()
            self.progress.pack_forget()
    
    def carregar_faixas_ip(self) -> List[str]:
        """
        Carrega as faixas de IP permitidas do arquivo de configuração.
        
        Returns:
            Lista de prefixos de IP permitidos.
        """
        try:
            config_path = get_config_path()
            if config_path and os.path.exists(config_path):
                config = configparser.ConfigParser()
                config.read(config_path, encoding='utf-8')
                
                if config.has_option('SEGURANCA', 'faixas_ip_permitidas'):
                    faixas = config.get('SEGURANCA', 'faixas_ip_permitidas').strip()
                    if faixas:
                        return [f.strip() for f in faixas.split(',') if f.strip()]
            
            # Valor padrão se não encontrar configuração
            return ['192.168.112.']
        except Exception as e:
            print(f"Erro ao carregar faixas de IP: {e}")
            return ['192.168.112.']
    
    def verificar_ip_permitido(self) -> bool:
        """
        Verifica se o IP da máquina está em uma das faixas permitidas.
        
        Returns:
            True se o IP está na faixa permitida, False caso contrário.
        """
        # Se não há faixas configuradas, permite qualquer IP
        if not self.faixas_ip_permitidas:
            return True
        
        try:
            # Obtém o nome do host
            hostname = socket.gethostname()
            # Obtém todos os IPs da máquina
            ips = socket.gethostbyname_ex(hostname)[2]
            
            # Verifica se algum IP está em alguma faixa permitida
            for ip in ips:
                for faixa in self.faixas_ip_permitidas:
                    if ip.startswith(faixa):
                        return True
            
            # Tenta obter o IP conectando a um servidor externo (mais confiável)
            try:
                s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
                s.connect(('8.8.8.8', 80))
                ip_local = s.getsockname()[0]
                s.close()
                for faixa in self.faixas_ip_permitidas:
                    if ip_local.startswith(faixa):
                        return True
            except:
                pass
            
            return False
        except Exception as e:
            print(f"Erro ao verificar IP: {e}")
            return False
    
    def toggle_conexao(self):
        """Conecta ou desconecta do servidor FTP."""
        if self.conectado:
            self.desconectar()
        else:
            self.conectar_com_modal()
    
    def conectar_com_modal(self, callback=None):
        """
        Conecta ao servidor SFTP exibindo um modal de progresso.
        
        Args:
            callback: Função a ser chamada após conexão bem-sucedida.
        """
        # Cria janela modal
        self.modal_conexao = tk.Toplevel(self.root)
        self.modal_conexao.title("Conectando...")
        self.modal_conexao.geometry("300x100")
        self.modal_conexao.resizable(False, False)
        self.modal_conexao.transient(self.root)
        self.modal_conexao.grab_set()
        
        # Centraliza o modal
        self.modal_conexao.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - 150
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - 50
        self.modal_conexao.geometry(f"+{x}+{y}")
        
        # Conteúdo do modal
        frame = ttk.Frame(self.modal_conexao, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text="🔄 Conectando ao servidor SFTP...", font=('Segoe UI', 11)).pack(pady=(0, 10))
        
        # Barra de progresso indeterminada
        progress = ttk.Progressbar(frame, mode='indeterminate', length=250)
        progress.pack()
        progress.start(10)
        
        # Impede fechar o modal
        self.modal_conexao.protocol("WM_DELETE_WINDOW", lambda: None)
        
        self._conexao_callback_pendente = callback
        
        def conectar_thread():
            try:
                self.ftp_client = SFTPClient()
                sucesso, mensagem = self.ftp_client.conectar()
                
                self.root.after(0, lambda s=sucesso, m=mensagem: self._conectar_callback_modal(s, m))
            except Exception as e:
                erro_msg = str(e)
                self.root.after(0, lambda msg=erro_msg: self._conectar_callback_modal(False, msg))
        
        thread = threading.Thread(target=conectar_thread)
        thread.daemon = True
        thread.start()
    
    def _conectar_callback_modal(self, sucesso: bool, mensagem: str):
        """Callback após tentativa de conexão com modal."""
        # Fecha o modal
        if hasattr(self, 'modal_conexao') and self.modal_conexao:
            self.modal_conexao.destroy()
            self.modal_conexao = None
        
        if sucesso:
            self.conectado = True
            self.lbl_status_conexao.config(text="● Conectado", foreground='green')
            self.atualizar_status(mensagem)
            
            # Executa o callback pendente (a busca)
            if hasattr(self, '_conexao_callback_pendente') and self._conexao_callback_pendente:
                callback = self._conexao_callback_pendente
                self._conexao_callback_pendente = None
                callback()
        else:
            self.conectado = False
            self.lbl_status_conexao.config(text="● Desconectado", foreground='red')
            self.atualizar_status(f"Erro: {mensagem}")
            messagebox.showerror("Erro de Conexão", mensagem)
    
    def conectar(self):
        """Conecta ao servidor SFTP (método legado)."""
        self.conectar_com_modal()
    
    def _conectar_callback(self, sucesso: bool, mensagem: str):
        """Callback após tentativa de conexão (legado)."""
        self._conectar_callback_modal(sucesso, mensagem)
    
    def desconectar(self):
        """Desconecta do servidor SFTP."""
        if self.ftp_client:
            self.ftp_client.desconectar()
        
        self.conectado = False
        self.lbl_status_conexao.config(text="● Desconectado", foreground='red')
        self.atualizar_status("Desconectado do servidor SFTP.")
        
        # Limpa resultados
        self.limpar_resultados()
    
    def limpar_resultados(self):
        """Limpa a lista de resultados."""
        for item in self.tree_resultados.get_children():
            self.tree_resultados.delete(item)
        self.resultados_busca = []
        self.lbl_contagem.config(text="")
    
    def on_tree_click(self, event):
        """Manipula clique no Treeview para toggle do checkbox."""
        # Identifica a região clicada
        region = self.tree_resultados.identify_region(event.x, event.y)
        
        if region == 'cell':
            # Identifica a coluna clicada
            column = self.tree_resultados.identify_column(event.x)
            
            # Se for a primeira coluna (checkbox)
            if column == '#1':
                item = self.tree_resultados.identify_row(event.y)
                if item:
                    self.toggle_checkbox(item)
                return "break"  # Previne a seleção normal
    
    def toggle_checkbox(self, item_id: str):
        """Alterna o estado do checkbox de um item."""
        item = self.tree_resultados.item(item_id)
        valores = list(item['values'])
        
        # Ignora separadores
        if valores[1] == '───':
            return
        
        # Alterna entre ☐ e ☑
        if valores[0] == '☑':
            valores[0] = '☐'
        else:
            valores[0] = '☑'
        
        self.tree_resultados.item(item_id, values=valores)
        self.atualizar_contagem_selecionados()
    
    def toggle_todos_checkboxes(self):
        """Marca ou desmarca todos os checkboxes."""
        itens = self.tree_resultados.get_children('')
        
        # Verifica se todos estão marcados
        todos_marcados = True
        for item_id in itens:
            valores = self.tree_resultados.item(item_id)['values']
            if valores[1] != '───' and valores[0] != '☑':
                todos_marcados = False
                break
        
        # Se todos estão marcados, desmarca todos; senão, marca todos
        novo_estado = '☐' if todos_marcados else '☑'
        
        for item_id in itens:
            item = self.tree_resultados.item(item_id)
            valores = list(item['values'])
            
            # Ignora separadores
            if valores[1] == '───':
                continue
            
            valores[0] = novo_estado
            self.tree_resultados.item(item_id, values=valores)
        
        self.atualizar_contagem_selecionados()
    
    def marcar_todos(self):
        """Marca todos os checkboxes."""
        for item_id in self.tree_resultados.get_children(''):
            item = self.tree_resultados.item(item_id)
            valores = list(item['values'])
            
            # Ignora separadores
            if valores[1] == '───':
                continue
            
            valores[0] = '☑'
            self.tree_resultados.item(item_id, values=valores)
        
        self.atualizar_contagem_selecionados()
    
    def desmarcar_todos(self):
        """Desmarca todos os checkboxes."""
        for item_id in self.tree_resultados.get_children(''):
            item = self.tree_resultados.item(item_id)
            valores = list(item['values'])
            
            # Ignora separadores
            if valores[1] == '───':
                continue
            
            valores[0] = '☐'
            self.tree_resultados.item(item_id, values=valores)
        
        self.atualizar_contagem_selecionados()
    
    def atualizar_contagem_selecionados(self):
        """Atualiza a contagem de itens marcados."""
        marcados = 0
        total = 0
        
        for item_id in self.tree_resultados.get_children(''):
            valores = self.tree_resultados.item(item_id)['values']
            
            # Ignora separadores
            if valores[1] == '───':
                continue
            
            total += 1
            if valores[0] == '☑':
                marcados += 1
        
        texto_atual = self.lbl_contagem.cget('text')
        # Extrai a primeira parte (contagem de boletos/NFs)
        if '|' in texto_atual:
            partes = texto_atual.split('|')
            base = '|'.join(partes[:3])
            self.lbl_contagem.config(text=f"{base} | {marcados}/{total} marcado(s)")
        else:
            self.lbl_contagem.config(text=f"{marcados}/{total} marcado(s)")
    
    def obter_itens_marcados(self) -> list:
        """Retorna lista de itens com checkbox marcado."""
        marcados = []
        
        for item_id in self.tree_resultados.get_children(''):
            valores = self.tree_resultados.item(item_id)['values']
            
            # Ignora separadores e itens não marcados
            if valores[1] == '───' or valores[0] != '☑':
                continue
            
            marcados.append(item_id)
        
        return marcados
    
    def ordenar_coluna(self, coluna: str, reverso: bool):
        """
        Ordena os itens do Treeview pela coluna especificada.
        
        Args:
            coluna: Nome da coluna para ordenar ('tipo', 'nome', 'data', 'caminho').
            reverso: Se True, ordena em ordem decrescente.
        """
        import re
        
        # Coleta todos os itens normais (não separadores) com seus dados
        itens_dados = []
        
        for item in self.tree_resultados.get_children(''):
            valores = self.tree_resultados.item(item)['values']
            tags = self.tree_resultados.item(item)['tags']
            if valores[1] != '───':
                itens_dados.append({
                    'item_id': item,
                    'valores': valores,
                    'tags': tags,
                    'coluna_valor': self.tree_resultados.set(item, coluna)
                })
        
        # Ordena os itens
        if coluna == 'data':
            def parse_data(valor):
                try:
                    if valor == '-':
                        return datetime.min
                    return datetime.strptime(valor, "%d/%m/%Y %H:%M")
                except:
                    return datetime.min
            
            itens_dados.sort(key=lambda x: parse_data(x['coluna_valor']), reverse=reverso)
        else:
            itens_dados.sort(key=lambda x: x['coluna_valor'].lower(), reverse=reverso)
        
        # Limpa o treeview completamente
        for item in self.tree_resultados.get_children(''):
            self.tree_resultados.delete(item)
        
        # Reagrupa os itens ordenados por número do documento
        grupos = {}
        for dados in itens_dados:
            nome = dados['valores'][5]  # Nome do arquivo (índice 5)
            numeros = re.sub(r'[^\d]', '', str(nome))
            if len(numeros) >= 15:
                chave = numeros[:15]
            else:
                chave = numeros
            
            if chave not in grupos:
                grupos[chave] = []
            grupos[chave].append(dados)
        
        # Reinserir itens com separadores entre grupos
        primeiro_grupo = True
        for chave in grupos:
            if not primeiro_grupo:
                # Adiciona separador entre grupos
                self.tree_resultados.insert('', tk.END, values=('', '───', '───', '─' * 30, '───', '─' * 30, ''), tags=('separador',))
            primeiro_grupo = False
            
            for dados in grupos[chave]:
                tag = dados['tags'][0] if dados['tags'] else 'boleto'
                self.tree_resultados.insert('', tk.END, values=dados['valores'], tags=(tag,))
        
        # Atualiza o cabeçalho com indicador de direção
        setas = {'tipo': '↕', 'numero': '↕', 'cliente': '↕', 'data': '↕', 'nome': '↕'}
        titulos = {'tipo': 'Tipo', 'numero': 'Número', 'cliente': 'Cliente', 'data': 'Data', 'nome': 'Nome do Arquivo'}
        
        # Define a seta para a coluna ordenada
        if coluna in setas:
            setas[coluna] = '↓' if reverso else '↑'
        
        # Atualiza os cabeçalhos
        self.tree_resultados.heading('tipo', text=f"{titulos['tipo']} {setas['tipo']}", 
                                     command=lambda: self.ordenar_coluna('tipo', not reverso if coluna == 'tipo' else False))
        self.tree_resultados.heading('numero', text=f"{titulos['numero']} {setas['numero']}", 
                                     command=lambda: self.ordenar_coluna('numero', not reverso if coluna == 'numero' else False))
        self.tree_resultados.heading('cliente', text=f"{titulos['cliente']} {setas['cliente']}", 
                                     command=lambda: self.ordenar_coluna('cliente', not reverso if coluna == 'cliente' else False))
        self.tree_resultados.heading('data', text=f"{titulos['data']} {setas['data']}", 
                                     command=lambda: self.ordenar_coluna('data', not reverso if coluna == 'data' else False))
        self.tree_resultados.heading('nome', text=f"{titulos['nome']} {setas['nome']}", 
                                     command=lambda: self.ordenar_coluna('nome', not reverso if coluna == 'nome' else False))
    
    def formatar_numero_boleto(self, numero: str) -> str:
        """
        Formata o número do boleto com a máscara padrão.
        
        Máscara: FILIAL (6 dígitos) + NUMERO (9 dígitos com zeros à esquerda)
        Exemplo: 5909 -> 010001000005909
        
        Args:
            numero: Número do boleto digitado pelo usuário.
            
        Returns:
            Número formatado com a máscara.
        """
        # Remove caracteres não numéricos
        numero_limpo = ''.join(filter(str.isdigit, numero))
        
        if not numero_limpo:
            return ""
        
        # Formata o número com zeros à esquerda (9 dígitos)
        # A busca será feita em todas as filiais
        numero_formatado = numero_limpo.zfill(9)
        
        return numero_formatado
    
    def _extrair_numero_documento(self, nome_arquivo: str) -> str:
        """
        Extrai o número do documento do nome do arquivo.
        
        O formato é: FILIAL (6 dígitos) + NUMERO (9 dígitos)
        Retorna os 9 dígitos do número sem zeros à esquerda.
        
        Args:
            nome_arquivo: Nome do arquivo PDF.
            
        Returns:
            Número do documento (sem zeros à esquerda).
        """
        import re
        # Extrai apenas dígitos do nome
        numeros = re.sub(r'[^\d]', '', nome_arquivo)
        
        if len(numeros) >= 15:
            # Pega os 9 dígitos após a filial (posições 6-14)
            numero_doc = numeros[6:15]
            # Remove zeros à esquerda
            numero_doc = numero_doc.lstrip('0') or '0'
            return numero_doc
        elif len(numeros) >= 9:
            # Se não tem filial, pega os últimos 9
            numero_doc = numeros[-9:]
            numero_doc = numero_doc.lstrip('0') or '0'
            return numero_doc
        else:
            return numeros.lstrip('0') or '0'
    
    def buscar_boleto(self):
        """Busca boletos no servidor SFTP."""
        # Verifica se está extraindo clientes
        if self.extraindo_clientes:
            if not self._mostrar_modal_extracao():
                return
        
        # Verifica se o IP está na faixa permitida
        if not self.verificar_ip_permitido():
            faixas_str = ', '.join(self.faixas_ip_permitidas) if self.faixas_ip_permitidas else 'nenhuma'
            messagebox.showerror(
                "Acesso Negado", 
                f"Este aplicativo só pode ser utilizado nas redes permitidas ({faixas_str}xxx).\n\n"
                "Verifique sua conexão de rede."
            )
            return
        
        numero = self.entry_numero.get().strip()
        
        if not numero:
            messagebox.showwarning("Aviso", "Digite o número do boleto para buscar.")
            return
        
        # Se não estiver conectado, conecta automaticamente
        if not self.conectado:
            self.conectar_com_modal(callback=self.buscar_boleto)
            return
        
        # Formata o número se busca literal estiver ativada
        busca_literal = self.var_busca_literal.get()
        if busca_literal:
            numero_busca = self.formatar_numero_boleto(numero)
            self.atualizar_status(f"Buscando boletos e NFs (literal): {numero_busca}...")
        else:
            numero_busca = numero
            self.atualizar_status(f"Buscando boletos e NFs (contém): {numero}...")
        
        self.mostrar_progresso(True)
        self.btn_buscar.config(state='disabled')
        self.limpar_resultados()
        
        def buscar_thread():
            try:
                resultados = self.ftp_client.buscar_boleto_e_nf(
                    numero_busca, 
                    busca_recursiva=True,
                    busca_literal=busca_literal
                )
                self.root.after(0, lambda r=resultados, n=numero_busca: self._buscar_callback(r, n))
            except Exception as e:
                erro_msg = str(e)
                self.root.after(0, lambda n=numero_busca, msg=erro_msg: self._buscar_callback([], n, msg))
        
        thread = threading.Thread(target=buscar_thread)
        thread.daemon = True
        thread.start()
    
    def _agrupar_resultados(self, resultados: list) -> list:
        """
        Agrupa os resultados por número do documento (NF + Boleto juntos).
        Mantém apenas o registro mais recente de cada tipo por número.
        
        Args:
            resultados: Lista de tuplas (caminho, nome, data, tipo).
            
        Returns:
            Lista agrupada com separadores.
        """
        import re
        
        # Primeiro, filtra para manter apenas o mais recente de cada (numero + tipo)
        # Chave: (numero_extraido, tipo) -> registro mais recente
        mais_recentes = {}
        
        for caminho, nome, data_mod, tipo in resultados:
            # Extrai o número do arquivo (assume formato FILIAL + NUMERO)
            numeros = re.sub(r'[^\d]', '', nome)
            if len(numeros) >= 15:
                numero_chave = numeros[:15]  # Filial + Numero
            else:
                numero_chave = numeros
            
            chave_unica = (numero_chave, tipo)
            
            # Mantém apenas o mais recente (maior data)
            if chave_unica not in mais_recentes or data_mod > mais_recentes[chave_unica][2]:
                mais_recentes[chave_unica] = (caminho, nome, data_mod, tipo)
        
        # Agora agrupa os registros filtrados por número
        grupos = {}
        for (numero_chave, tipo), (caminho, nome, data_mod, tipo) in mais_recentes.items():
            if numero_chave not in grupos:
                grupos[numero_chave] = []
            grupos[numero_chave].append((caminho, nome, data_mod, tipo))
        
        # Ordena cada grupo (NF primeiro, depois Boleto)
        for chave in grupos:
            grupos[chave].sort(key=lambda x: (0 if x[3] == 'NF' else 1, x[1]))
        
        return grupos
    
    def _buscar_callback(self, resultados: list, numero: str, erro: str = None):
        """Callback após busca de boletos e NFs."""
        self.mostrar_progresso(False)
        self.btn_buscar.config(state='normal')
        
        if erro:
            self.atualizar_status(f"Erro na busca: {erro}")
            messagebox.showerror("Erro", f"Erro ao buscar: {erro}")
            return
        
        # Armazena resultados
        self.resultados_busca = [(caminho, nome, tipo) for caminho, nome, _, tipo in resultados]
        
        # Agrupa resultados
        grupos = self._agrupar_resultados(resultados)
        
        # Filtra grupos não vazios
        grupos_validos = [(chave, itens) for chave, itens in sorted(grupos.items()) if itens]
        
        # Coleta todos os caminhos para extrair clientes
        todos_caminhos = []
        for chave, itens in grupos_validos:
            for caminho, nome, data_mod, tipo in itens:
                todos_caminhos.append(caminho)
        
        # Popula a lista de resultados agrupados (com cliente vazio inicialmente)
        for idx, (chave, itens) in enumerate(grupos_validos):
            # Adiciona separador entre grupos (exceto o primeiro)
            if idx > 0:
                self.tree_resultados.insert('', tk.END, values=('', '───', '───', '─' * 30, '───', '─' * 30, ''), tags=('separador',))
            
            for caminho, nome, data_mod, tipo in itens:
                data_formatada = data_mod.strftime("%d/%m/%Y %H:%M")
                tag = 'boleto' if tipo == 'BOLETO' else 'nf'
                # Extrai número do documento (9 últimos dígitos, sem zeros à esquerda)
                numero_doc = self._extrair_numero_documento(nome)
                self.tree_resultados.insert('', tk.END, values=('☐', tipo, numero_doc, 'Carregando...', data_formatada, nome, caminho), tags=(tag,))
        
        # Atualiza contagem
        qtd_boletos = sum(1 for _, _, _, t in resultados if t == 'BOLETO')
        qtd_nfs = sum(1 for _, _, _, t in resultados if t == 'NF')
        total = qtd_boletos + qtd_nfs
        self.lbl_contagem.config(text=f"{qtd_boletos} boleto(s) | {qtd_nfs} NF(s) | {len(grupos_validos)} grupo(s) | 0/{total} marcado(s)")
        
        if len(resultados) == 0:
            self.atualizar_status(f"Nenhum resultado encontrado para: {numero}")
            messagebox.showinfo("Busca", f"Nenhum boleto ou NF encontrado para o número: {numero}")
        else:
            self.atualizar_status(f"Encontrado(s) {qtd_boletos} boleto(s) e {qtd_nfs} NF(s) para: {numero}. Extraindo nomes de clientes...")
            # Seleciona o primeiro item
            primeiro = self.tree_resultados.get_children()[0]
            self.tree_resultados.selection_set(primeiro)
            self.tree_resultados.focus(primeiro)
            
            # Extrai nomes de clientes em thread separada
            self._extrair_clientes_async()
    
    def _extrair_clientes_async(self):
        """Extrai nomes de clientes dos PDFs em thread separada.
        Extrai apenas do BOLETO e usa o mesmo nome para a NF do mesmo grupo."""
        import re
        
        # Marca que está extraindo
        self.extraindo_clientes = True
        self.cancelar_extracao = False
        
        def extrair_thread():
            try:
                # Agrupa itens por número do documento
                grupos = {}  # chave -> lista de (item_id, tipo, caminho)
                
                for item_id in self.tree_resultados.get_children(''):
                    # Verifica se foi cancelado
                    if self.cancelar_extracao:
                        break
                    
                    valores = self.tree_resultados.item(item_id)['values']
                    # Ignora separadores
                    if valores[1] == '───':
                        continue
                    
                    if valores[3] == 'Carregando...':
                        tipo = valores[1]   # BOLETO ou NF
                        nome = valores[5]   # Nome do arquivo (índice 5)
                        caminho = valores[6]  # Caminho (índice 6)
                        
                        # Extrai número para agrupar
                        numeros = re.sub(r'[^\d]', '', nome)
                        if len(numeros) >= 15:
                            chave = numeros[:15]
                        else:
                            chave = numeros
                        
                        if chave not in grupos:
                            grupos[chave] = []
                        grupos[chave].append((item_id, tipo, caminho))
                
                # Para cada grupo, extrai do BOLETO e aplica para todos
                total_grupos = len(grupos)
                for idx, (chave, itens) in enumerate(grupos.items()):
                    # Verifica se foi cancelado
                    if self.cancelar_extracao:
                        break
                    
                    try:
                        self.root.after(0, lambda i=idx, t=total_grupos: 
                            self.atualizar_status(f"Extraindo cliente {i+1}/{t}..."))
                        
                        # Encontra o BOLETO do grupo para extrair o nome
                        cliente = ""
                        boleto_caminho = None
                        
                        for item_id, tipo, caminho in itens:
                            if tipo == 'BOLETO':
                                boleto_caminho = caminho
                                break
                        
                        # Se tem boleto, extrai dele
                        if boleto_caminho:
                            cliente = self.ftp_client.extrair_cliente_do_pdf(boleto_caminho)
                        
                        # Se não encontrou cliente no boleto, tenta da NF
                        if not cliente:
                            for item_id, tipo, caminho in itens:
                                if tipo == 'NF':
                                    cliente = self.ftp_client.extrair_cliente_do_pdf(caminho)
                                    if cliente:
                                        break
                        
                        if not cliente:
                            cliente = "-"
                        
                        # Aplica o mesmo nome para todos os itens do grupo
                        for item_id, tipo, caminho in itens:
                            def atualizar_item(iid=item_id, cli=cliente):
                                try:
                                    valores = list(self.tree_resultados.item(iid)['values'])
                                    valores[3] = cli  # Coluna cliente
                                    self.tree_resultados.item(iid, values=valores)
                                except:
                                    pass
                            self.root.after(0, atualizar_item)
                        
                    except Exception as e:
                        # Em caso de erro, marca todos do grupo como não encontrado
                        for item_id, tipo, caminho in itens:
                            def marcar_erro(iid=item_id):
                                try:
                                    valores = list(self.tree_resultados.item(iid)['values'])
                                    valores[3] = "-"
                                    self.tree_resultados.item(iid, values=valores)
                                except:
                                    pass
                            self.root.after(0, marcar_erro)
                
                # Finaliza
                def finalizar_extracao():
                    self.extraindo_clientes = False
                    if self.cancelar_extracao:
                        self.atualizar_status("Extração de clientes cancelada.")
                    else:
                        self.atualizar_status("Extração de clientes concluída.")
                
                self.root.after(0, finalizar_extracao)
                
            except Exception as e:
                def finalizar_com_erro():
                    self.extraindo_clientes = False
                    self.atualizar_status(f"Erro ao extrair clientes: {e}")
                self.root.after(0, finalizar_com_erro)
        
        thread = threading.Thread(target=extrair_thread)
        thread.daemon = True
        thread.start()
    
    def _mostrar_modal_extracao(self) -> bool:
        """
        Mostra modal perguntando se deseja cancelar a extração em andamento.
        
        Returns:
            True se o usuário cancelou a extração, False se quer aguardar.
        """
        resposta = messagebox.askyesno(
            "Extração em Andamento",
            "A extração de nomes de clientes ainda está em andamento.\n\n"
            "Deseja cancelar a extração e iniciar uma nova busca?\n\n"
            "• Sim - Cancela a extração e inicia nova busca\n"
            "• Não - Aguarda a conclusão da extração",
            icon='warning'
        )
        
        if resposta:
            # Usuário quer cancelar
            self.cancelar_extracao = True
            self.atualizar_status("Cancelando extração...")
            
            # Aguarda um pouco para a thread terminar
            import time
            timeout = 2  # segundos
            start = time.time()
            while self.extraindo_clientes and (time.time() - start) < timeout:
                self.root.update()
                time.sleep(0.1)
            
            return True
        else:
            # Usuário quer aguardar
            return False
    
    def cancelar_extracao_clientes(self):
        """Cancela a extração de clientes em andamento."""
        if not self.extraindo_clientes:
            self.atualizar_status("Nenhuma extração em andamento.")
            return
        
        # Sinaliza para cancelar
        self.cancelar_extracao = True
        self.atualizar_status("Cancelando extração...")
        
        # Aguarda a thread terminar (com timeout)
        import time
        timeout = 2  # segundos
        start = time.time()
        while self.extraindo_clientes and (time.time() - start) < timeout:
            self.root.update()
            time.sleep(0.1)
        
        if not self.extraindo_clientes:
            self.atualizar_status("Extração cancelada.")
        else:
            self.atualizar_status("Extração será cancelada em breve...")
    
    def definir_data_hoje(self):
        """Define ambos os campos de data para hoje."""
        hoje = datetime.now().strftime("%d/%m/%Y")
        self.entry_data_inicio.delete(0, tk.END)
        self.entry_data_inicio.insert(0, hoje)
        self.entry_data_fim.delete(0, tk.END)
        self.entry_data_fim.insert(0, hoje)
    
    def definir_ultima_semana(self):
        """Define o período para a última semana."""
        from datetime import timedelta
        hoje = datetime.now()
        semana_atras = hoje - timedelta(days=7)
        
        self.entry_data_inicio.delete(0, tk.END)
        self.entry_data_inicio.insert(0, semana_atras.strftime("%d/%m/%Y"))
        self.entry_data_fim.delete(0, tk.END)
        self.entry_data_fim.insert(0, hoje.strftime("%d/%m/%Y"))
    
    def definir_ultimo_mes(self):
        """Define o período para o último mês."""
        from datetime import timedelta
        hoje = datetime.now()
        mes_atras = hoje - timedelta(days=30)
        
        self.entry_data_inicio.delete(0, tk.END)
        self.entry_data_inicio.insert(0, mes_atras.strftime("%d/%m/%Y"))
        self.entry_data_fim.delete(0, tk.END)
        self.entry_data_fim.insert(0, hoje.strftime("%d/%m/%Y"))
    
    def validar_data(self, data_str: str) -> Optional[datetime]:
        """
        Valida e converte uma string de data para datetime.
        
        Args:
            data_str: Data no formato DD/MM/AAAA.
            
        Returns:
            datetime se válido, None caso contrário.
        """
        try:
            return datetime.strptime(data_str.strip(), "%d/%m/%Y")
        except ValueError:
            return None
    
    def buscar_por_data(self):
        """Busca boletos por período de data de criação."""
        # Verifica se está extraindo clientes
        if self.extraindo_clientes:
            if not self._mostrar_modal_extracao():
                return
        
        # Verifica se o IP está na faixa permitida
        if not self.verificar_ip_permitido():
            faixas_str = ', '.join(self.faixas_ip_permitidas) if self.faixas_ip_permitidas else 'nenhuma'
            messagebox.showerror(
                "Acesso Negado", 
                f"Este aplicativo só pode ser utilizado nas redes permitidas ({faixas_str}xxx).\n\n"
                "Verifique sua conexão de rede."
            )
            return
        
        data_inicio_str = self.entry_data_inicio.get().strip()
        data_fim_str = self.entry_data_fim.get().strip()
        
        # Valida as datas
        data_inicio = self.validar_data(data_inicio_str)
        data_fim = self.validar_data(data_fim_str)
        
        if not data_inicio:
            messagebox.showwarning("Aviso", "Data inicial inválida. Use o formato DD/MM/AAAA.")
            self.entry_data_inicio.focus_set()
            return
        
        if not data_fim:
            messagebox.showwarning("Aviso", "Data final inválida. Use o formato DD/MM/AAAA.")
            self.entry_data_fim.focus_set()
            return
        
        if data_inicio > data_fim:
            messagebox.showwarning("Aviso", "A data inicial não pode ser maior que a data final.")
            return
        
        # Se não estiver conectado, conecta automaticamente
        if not self.conectado:
            self.conectar_com_modal(callback=self.buscar_por_data)
            return
        
        periodo = f"{data_inicio_str} a {data_fim_str}"
        self.atualizar_status(f"Buscando boletos e NFs de {periodo}...")
        
        self.mostrar_progresso(True)
        self.btn_buscar_data.config(state='disabled')
        self.limpar_resultados()
        
        def buscar_thread():
            try:
                resultados = self.ftp_client.buscar_por_data(data_inicio, data_fim)
                self.root.after(0, lambda r=resultados, p=periodo: self._buscar_data_callback(r, p))
            except Exception as e:
                erro_msg = str(e)
                self.root.after(0, lambda p=periodo, msg=erro_msg: self._buscar_data_callback([], p, msg))
        
        thread = threading.Thread(target=buscar_thread)
        thread.daemon = True
        thread.start()
    
    def _buscar_data_callback(self, resultados: list, periodo: str, erro: str = None):
        """Callback após busca de boletos e NFs por data."""
        self.mostrar_progresso(False)
        self.btn_buscar_data.config(state='normal')
        
        if erro:
            self.atualizar_status(f"Erro na busca: {erro}")
            messagebox.showerror("Erro", f"Erro ao buscar: {erro}")
            return
        
        # Armazena resultados
        self.resultados_busca = [(caminho, nome, tipo) for caminho, nome, _, tipo in resultados]
        
        # Agrupa resultados
        grupos = self._agrupar_resultados(resultados)
        
        # Filtra grupos não vazios
        grupos_validos = [(chave, itens) for chave, itens in sorted(grupos.items()) if itens]
        
        # Popula a lista de resultados agrupados (com cliente vazio inicialmente)
        for idx, (chave, itens) in enumerate(grupos_validos):
            # Adiciona separador entre grupos (exceto o primeiro)
            if idx > 0:
                self.tree_resultados.insert('', tk.END, values=('', '───', '───', '─' * 30, '───', '─' * 30, ''), tags=('separador',))
            
            for caminho, nome, data_mod, tipo in itens:
                data_formatada = data_mod.strftime("%d/%m/%Y %H:%M")
                tag = 'boleto' if tipo == 'BOLETO' else 'nf'
                # Extrai número do documento (9 últimos dígitos, sem zeros à esquerda)
                numero_doc = self._extrair_numero_documento(nome)
                self.tree_resultados.insert('', tk.END, values=('☐', tipo, numero_doc, 'Carregando...', data_formatada, nome, caminho), tags=(tag,))
        
        # Atualiza contagem
        qtd_boletos = sum(1 for _, _, _, t in resultados if t == 'BOLETO')
        qtd_nfs = sum(1 for _, _, _, t in resultados if t == 'NF')
        total = qtd_boletos + qtd_nfs
        self.lbl_contagem.config(text=f"{qtd_boletos} boleto(s) | {qtd_nfs} NF(s) | {len(grupos_validos)} grupo(s) | 0/{total} marcado(s)")
        
        if len(resultados) == 0:
            self.atualizar_status(f"Nenhum resultado encontrado no período: {periodo}")
            messagebox.showinfo("Busca", f"Nenhum boleto ou NF encontrado no período: {periodo}")
        else:
            self.atualizar_status(f"Encontrado(s) {qtd_boletos} boleto(s) e {qtd_nfs} NF(s) no período: {periodo}. Extraindo nomes de clientes...")
            # Seleciona o primeiro item
            primeiro = self.tree_resultados.get_children()[0]
            self.tree_resultados.selection_set(primeiro)
            self.tree_resultados.focus(primeiro)
            
            # Extrai nomes de clientes em thread separada
            self._extrair_clientes_async()
    
    def selecionar_todos(self):
        """Marca todos os checkboxes na lista de resultados."""
        self.marcar_todos()
    
    def baixar_selecionado(self):
        """Baixa o(s) arquivo(s) marcado(s) na lista."""
        # Usa os itens marcados com checkbox
        marcados = self.obter_itens_marcados()
        
        if not marcados:
            messagebox.showwarning("Aviso", "Marque um ou mais arquivos para baixar.\n\nClique no ☐ para marcar.")
            return
        
        if not self.conectado:
            messagebox.showwarning("Aviso", "Conecte ao servidor SFTP primeiro.")
            return
        
        # Se apenas um arquivo marcado, baixa direto
        if len(marcados) == 1:
            self._baixar_unico(marcados[0])
        else:
            # Múltiplos arquivos - baixa e cria ZIP
            self._baixar_multiplos(marcados)
    
    def _baixar_unico(self, item_id: str):
        """Baixa um único arquivo (boleto ou NF)."""
        item = self.tree_resultados.item(item_id)
        valores = item['values']
        
        # Verifica se é um separador
        if valores[1] == '───':
            messagebox.showwarning("Aviso", "Selecione um arquivo válido, não um separador.")
            return
        
        tipo = valores[1]   # Índice 1 - tipo
        nome = valores[5]   # Índice 5 - nome do arquivo
        caminho = valores[6]  # Índice 6 - caminho (oculto)
        
        self.atualizar_status(f"Baixando {tipo}: {nome}...")
        self.mostrar_progresso(True)
        self.btn_baixar.config(state='disabled')
        
        def baixar_thread():
            try:
                sucesso, resultado = self.ftp_client.baixar_boleto(caminho, nome)
                self.root.after(0, lambda s=sucesso, r=resultado, n=nome: self._baixar_callback(s, r, n))
            except Exception as e:
                erro_msg = str(e)
                self.root.after(0, lambda msg=erro_msg, n=nome: self._baixar_callback(False, msg, n))
        
        thread = threading.Thread(target=baixar_thread)
        thread.daemon = True
        thread.start()
    
    def _baixar_multiplos(self, selecao: tuple):
        """Baixa múltiplos arquivos (boletos e NFs) e cria um arquivo ZIP."""
        # Coleta informações dos arquivos marcados (ignora separadores)
        arquivos = []
        numeros_docs = set()  # Coleta os números dos documentos
        for item_id in selecao:
            item = self.tree_resultados.item(item_id)
            valores = item['values']
            
            # Ignora separadores
            if valores[1] == '───':
                continue
                
            tipo = valores[1]   # Índice 1 - tipo
            numero = valores[2]  # Índice 2 - número do documento
            nome = valores[5]   # Índice 5 - nome do arquivo
            caminho = valores[6]  # Índice 6 - caminho (oculto)
            arquivos.append((caminho, nome, tipo))
            
            # Adiciona número do documento (se não for vazio ou "-")
            if numero and numero != "-":
                numeros_docs.add(str(numero))
        
        if not arquivos:
            messagebox.showwarning("Aviso", "Nenhum arquivo válido marcado.")
            return
        
        # Prepara o identificador para o nome do arquivo
        if numeros_docs:
            # Ordena e junta os números (limitado a 3 para não ficar muito longo)
            numeros_lista = sorted(numeros_docs, key=lambda x: int(x) if x.isdigit() else 0)
            if len(numeros_lista) > 3:
                identificador = f"{numeros_lista[0]}-{numeros_lista[-1]}"
            else:
                identificador = "-".join(numeros_lista)
        else:
            identificador = "doc"
        
        qtd = len(arquivos)
        self.atualizar_status(f"Baixando {qtd} arquivo(s)...")
        self.mostrar_progresso(True)
        self.btn_baixar.config(state='disabled')
        
        def baixar_thread():
            arquivos_baixados = []
            erros = []
            tipos_baixados = set()
            
            try:
                # Baixa cada arquivo
                for i, (caminho, nome, tipo) in enumerate(arquivos, 1):
                    self.root.after(0, lambda idx=i, total=qtd, n=nome, t=tipo: 
                        self.atualizar_status(f"Baixando {idx}/{total} ({t}): {n}..."))
                    
                    sucesso, resultado = self.ftp_client.baixar_boleto(caminho, nome)
                    if sucesso:
                        arquivos_baixados.append(resultado)
                        tipos_baixados.add(tipo)
                    else:
                        erros.append(f"{nome}: {resultado}")
                
                # Cria arquivo ZIP se houver arquivos baixados
                if arquivos_baixados:
                    # Determina o prefixo do nome do ZIP baseado nos tipos
                    if 'NF' in tipos_baixados and 'BOLETO' in tipos_baixados:
                        prefixo = "NFBOL"
                    elif 'NF' in tipos_baixados:
                        prefixo = "NF"
                    else:
                        prefixo = "BOL"
                    
                    # Nome do ZIP: PREFIXO_NUMERO_DD-MM-AAAA.zip
                    data_atual = datetime.now()
                    data_formatada = data_atual.strftime("%d-%m-%Y")
                    zip_nome = f"{prefixo}_{identificador}_{data_formatada}.zip"
                    zip_path = os.path.join(self.ftp_client.pasta_download, zip_nome)
                    
                    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                        for arquivo in arquivos_baixados:
                            zipf.write(arquivo, os.path.basename(arquivo))
                    
                    # Remove os arquivos individuais após criar o ZIP
                    for arquivo in arquivos_baixados:
                        try:
                            os.remove(arquivo)
                        except:
                            pass
                    
                    self.root.after(0, lambda zp=zip_path, qtd_ok=len(arquivos_baixados), qtd_err=len(erros): 
                        self._baixar_multiplos_callback(True, zp, qtd_ok, qtd_err, erros))
                else:
                    self.root.after(0, lambda: 
                        self._baixar_multiplos_callback(False, "", 0, len(erros), erros))
                    
            except Exception as e:
                erro_msg = str(e)
                self.root.after(0, lambda msg=erro_msg: 
                    self._baixar_multiplos_callback(False, "", 0, 1, [msg]))
        
        thread = threading.Thread(target=baixar_thread)
        thread.daemon = True
        thread.start()
    
    def _baixar_multiplos_callback(self, sucesso: bool, zip_path: str, qtd_ok: int, qtd_err: int, erros: List[str]):
        """Callback após download múltiplo."""
        self.mostrar_progresso(False)
        self.btn_baixar.config(state='normal')
        
        if sucesso and qtd_ok > 0:
            msg = f"{qtd_ok} arquivo(s) baixado(s) com sucesso!"
            if qtd_err > 0:
                msg += f"\n{qtd_err} erro(s) encontrado(s)."
            
            self.atualizar_status(f"ZIP criado: {zip_path}")
            
            if messagebox.askyesno(
                "Download Concluído",
                f"{msg}\n\nArquivo ZIP criado: {os.path.basename(zip_path)}\n\n"
                f"Salvo em: {zip_path}\n\nDeseja abrir o arquivo ZIP?"
            ):
                self.abrir_arquivo(zip_path)
        else:
            self.atualizar_status("Erro no download")
            erro_detalhes = "\n".join(erros[:5])  # Mostra até 5 erros
            if len(erros) > 5:
                erro_detalhes += f"\n... e mais {len(erros) - 5} erro(s)"
            messagebox.showerror("Erro", f"Erro ao baixar arquivos:\n{erro_detalhes}")
    
    def _baixar_callback(self, sucesso: bool, resultado: str, nome: str):
        """Callback após download do boleto."""
        self.mostrar_progresso(False)
        self.btn_baixar.config(state='normal')
        
        if sucesso:
            self.atualizar_status(f"Boleto baixado: {resultado}")
            
            # Pergunta se deseja abrir o arquivo
            if messagebox.askyesno(
                "Download Concluído",
                f"Boleto '{nome}' baixado com sucesso!\n\n"
                f"Salvo em: {resultado}\n\n"
                "Deseja abrir o arquivo?"
            ):
                self.abrir_arquivo(resultado)
        else:
            self.atualizar_status(f"Erro no download: {resultado}")
            messagebox.showerror("Erro", f"Erro ao baixar boleto:\n{resultado}")
    
    def abrir_arquivo(self, caminho: str):
        """Abre um arquivo com o programa padrão do sistema."""
        try:
            if sys.platform == 'win32':
                os.startfile(caminho)
            elif sys.platform == 'darwin':
                subprocess.call(['open', caminho])
            else:
                subprocess.call(['xdg-open', caminho])
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao abrir arquivo: {e}")
    
    def abrir_pasta_downloads(self):
        """Abre a pasta de downloads no explorador de arquivos."""
        try:
            pasta = "downloads"
            if self.ftp_client:
                pasta = self.ftp_client.pasta_download
            
            # Cria a pasta se não existir
            if not os.path.exists(pasta):
                os.makedirs(pasta)
            
            # Abre no explorador
            caminho_absoluto = os.path.abspath(pasta)
            
            if sys.platform == 'win32':
                os.startfile(caminho_absoluto)
            elif sys.platform == 'darwin':
                subprocess.call(['open', caminho_absoluto])
            else:
                subprocess.call(['xdg-open', caminho_absoluto])
                
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao abrir pasta: {e}")
    
    def on_closing(self):
        """Evento de fechamento da janela."""
        if self.conectado and self.ftp_client:
            self.ftp_client.desconectar()
        self.root.destroy()


def main():
    """Função principal."""
    root = tk.Tk()
    app = BuscaBoletoApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()


if __name__ == "__main__":
    main()
