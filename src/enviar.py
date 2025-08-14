import os
from dotenv import load_dotenv
from pathlib import Path

import csv
import sys
import threading
import smtplib
import configparser
import time
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                            QLabel, QLineEdit, QPushButton, QTextEdit, QProgressBar,
                            QTreeWidget, QTreeWidgetItem, QFileDialog, QMessageBox,
                            QGroupBox, QFormLayout, QToolBar, QTabWidget, QScrollArea)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QTextCursor, QAction, QPalette, QColor, QIcon


# Configurações iniciais
DEFAULT_CONFIG = {
    'smtp_server': 'smtp.gmail.com',
    'smtp_port': '587',
    'email_subject': 'Documentos Importantes'
}

class EmailSenderApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Ferramenta de Envio de E-mails em Massa")
        self.resize(1200, 800)
        self.dark_mode = False

        try:
            icon_path = os.path.join(os.path.dirname(__file__), 'assets', 'icon-email-sender.ico.ico')
            self.setWindowIcon(QIcon(icon_path))
        except:
            self.log("Ícone não encontrado, usando padrão do sistema")
        
        # Estrutura de dados
        self.lista_para_envio = []
        self.anexos_selecionados = []
        self.attachments_dir = str(Path.home() / "Documents")  # Pasta padrão
        
        self.init_ui()
        self.load_config()
        
    def init_ui(self):
        # Configuração principal da janela
        self.setWindowTitle("Ferramenta de Envio de E-mails em Massa")
        self.resize(1400, 900)
        
        # Layout principal (dividido em esquerda e direita)
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QHBoxLayout(main_widget)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(15)
        
        # --- COLUNA ESQUERDA (30% largura) ---
        left_column = QWidget()
        left_layout = QVBoxLayout(left_column)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(15)
        
        # Seção 1: Adicionar Destinatário
        add_group = QGroupBox("Adicionar Destinatário")
        add_group.setStyleSheet("QGroupBox { font-weight: bold; }")
        add_layout = QFormLayout()
        add_layout.setRowWrapPolicy(QFormLayout.RowWrapPolicy.WrapAllRows)
        
        # Campos de entrada
        self.nome_entry = QLineEdit()
        self.email_entry = QLineEdit()
        self.anexos_label = QLabel("Nenhum arquivo selecionado")
        self.anexos_label.setStyleSheet("color: #666; font-style: italic;")
        
        # Botões
        btn_frame = QWidget()
        btn_layout = QHBoxLayout(btn_frame)
        btn_layout.setContentsMargins(0, 0, 0, 0)
        
        btn_adicionar_anexos = QPushButton("Adicionar Anexos")
        btn_adicionar_anexos.setStyleSheet("background-color: #673AB7; color: white;")
        btn_adicionar_anexos.clicked.connect(self.selecionar_arquivos)
        
        btn_importar = QPushButton("Importar CSV")
        btn_importar.setStyleSheet("background-color: #FF9800; color: white;")
        btn_importar.clicked.connect(self.importar_contatos_csv)
        
        btn_adicionar = QPushButton("Adicionar à Lista")
        btn_adicionar.setStyleSheet("background-color: #4CAF50; color: white;")
        btn_adicionar.clicked.connect(self.adicionar_destinatario)
        
        btn_layout.addWidget(btn_adicionar_anexos)
        btn_layout.addWidget(btn_importar)
        btn_layout.addWidget(btn_adicionar)
        
        # Organização dos elementos
        add_layout.addRow("Nome:", self.nome_entry)
        add_layout.addRow("E-mail:", self.email_entry)
        add_layout.addRow("Anexos:", self.anexos_label)
        add_layout.addRow(btn_frame)
        add_group.setLayout(add_layout)
        
        # Seção 2: Configurações do E-mail
        config_group = QGroupBox("Configurações do E-mail")
        config_group.setStyleSheet("QGroupBox { font-weight: bold; }")
        config_layout = QFormLayout()
        
        self.smtp_server = QLineEdit()
        self.smtp_port = QLineEdit()
        self.email_user = QLineEdit()
        self.email_pass = QLineEdit()
        self.email_pass.setEchoMode(QLineEdit.EchoMode.Password)
        self.email_subject = QLineEdit()
        
        config_layout.addRow("Servidor SMTP:", self.smtp_server)
        config_layout.addRow("Porta:", self.smtp_port)
        config_layout.addRow("Seu E-mail:", self.email_user)
        config_layout.addRow("Senha:", self.email_pass)
        config_layout.addRow("Assunto:", self.email_subject)
        
        btn_save_config = QPushButton("Salvar Configurações")
        btn_save_config.setStyleSheet("background-color: #2196F3; color: white;")
        btn_save_config.clicked.connect(self.save_config)
        config_layout.addRow(btn_save_config)
        
        config_group.setLayout(config_layout)

        # Seção 3: Preferências
        pref_group = QGroupBox("Preferências")
        pref_group.setStyleSheet("QGroupBox { font-weight: bold; }")
        pref_layout = QVBoxLayout()
        
        # Botão de alternar tema
        self.tema_btn = QPushButton("🌙 Ativar Tema Escuro")
        self.tema_btn.setStyleSheet("""
            QPushButton {
                padding: 8px;
                border-radius: 4px;
                background-color: #f0f0f0;
            }
            QPushButton:hover {
                background-color: #e0e0e0;
            }
        """)
        self.tema_btn.clicked.connect(self.alternar_tema)
        
        pref_layout.addWidget(self.tema_btn)
        pref_group.setLayout(pref_layout)
        
        # Adiciona as seções à coluna esquerda
        left_layout.addWidget(add_group)
        left_layout.addWidget(config_group)
        left_layout.addWidget(pref_group)
        left_layout.addStretch()
        
        # --- COLUNA DIREITA (70% largura) ---
        right_column = QWidget()
        right_layout = QVBoxLayout(right_column)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(15)
        
        # Seção 3: Lista de Envios Pendentes
        view_group = QGroupBox("Lista de Envios Pendentes")
        view_group.setStyleSheet("QGroupBox { font-weight: bold; }")
        view_layout = QVBoxLayout()
        
        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["Nome", "E-mail", "Arquivos Anexados"])
        self.tree.setColumnWidth(0, 200)
        self.tree.setColumnWidth(1, 250)
        self.tree.setColumnWidth(2, 350)
        
        view_layout.addWidget(self.tree)
        view_group.setLayout(view_layout)
        
        # Seção 4: Editor de E-mail (com aba para Log)
        tab_widget = QTabWidget()
        
        # Aba 1: Editor de E-mail
        editor_tab = QWidget()
        editor_tab_layout = QVBoxLayout(editor_tab)
        
        # Toolbar de formatação
        toolbar = QToolBar()
        btn_bold = QAction("Negrito", self)
        btn_bold.triggered.connect(self.formatar_texto_negrito)
        toolbar.addAction(btn_bold)
        
        # Área de edição
        self.text_edit = QTextEdit()
        self.text_edit.setAcceptRichText(True)
        
        # Carrega template se existir
        if os.path.exists('corpo_email.txt'):
            with open('corpo_email.txt', 'r', encoding='utf-8') as f:
                self.text_edit.setPlainText(f.read())
        
        # Botões do editor
        editor_btn_frame = QWidget()
        editor_btn_layout = QHBoxLayout(editor_btn_frame)
        editor_btn_layout.setContentsMargins(0, 0, 0, 0)
        
        btn_salvar = QPushButton("💾 Salvar Template")
        btn_salvar.setStyleSheet("background-color: #4CAF50; color: white;")
        btn_salvar.clicked.connect(self.salvar_template_email)
        
        btn_carregar = QPushButton("📂 Carregar Template")
        btn_carregar.setStyleSheet("background-color: #2196F3; color: white;")
        btn_carregar.clicked.connect(self.carregar_template_email)
        
        editor_btn_layout.addWidget(btn_salvar)
        editor_btn_layout.addWidget(btn_carregar)
        
        # Monta o layout do editor
        editor_tab_layout.addWidget(toolbar)
        editor_tab_layout.addWidget(self.text_edit)
        editor_tab_layout.addWidget(editor_btn_frame)
        
        # Aba 2: Log de Execução
        log_tab = QWidget()
        log_tab_layout = QVBoxLayout(log_tab)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setStyleSheet("""
            QTextEdit {
                background-color: #f9f9f9;
                font-family: Consolas, monospace;
                font-size: 10pt;
            }
        """)
        
        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        
        log_tab_layout.addWidget(self.log_text)
        log_tab_layout.addWidget(self.progress)
        
        # Adiciona as abas
        tab_widget.addTab(editor_tab, "Editor de E-mail")
        tab_widget.addTab(log_tab, "Log de Execução")
        
        # Botões de controle
        control_frame = QWidget()
        control_layout = QHBoxLayout(control_frame)
        control_layout.setContentsMargins(0, 0, 0, 0)
        
        self.clear_btn = QPushButton("Limpar Lista")
        self.clear_btn.setStyleSheet("background-color: #f44336; color: white;")
        self.clear_btn.clicked.connect(self.limpar_lista)
        
        self.send_btn = QPushButton("▶ ENVIAR TUDO")
        self.send_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                font-weight: bold;
            }
            QPushButton:disabled {
                background-color: #cccccc;
            }
        """)
        self.send_btn.clicked.connect(self.send_emails)
        
        control_layout.addWidget(self.clear_btn)
        control_layout.addStretch()
        control_layout.addWidget(self.send_btn)
        
        # Adiciona tudo à coluna direita
        right_layout.addWidget(view_group)
        right_layout.addWidget(tab_widget)
        right_layout.addWidget(control_frame)
        
        # Adiciona as colunas ao layout principal
        main_layout.addWidget(left_column, stretch=3)  # 30%
        main_layout.addWidget(right_column, stretch=7) # 70%
        
        # Carrega configurações
        self.load_config()

    def alternar_tema(self):
        """Alterna entre tema claro e escuro"""
        self.dark_mode = not self.dark_mode
        
        palette = QPalette()
        if self.dark_mode:
            # Configuração do tema escuro
            palette.setColor(QPalette.ColorRole.Window, QColor(53, 53, 53))
            palette.setColor(QPalette.ColorRole.WindowText, Qt.GlobalColor.white)
            palette.setColor(QPalette.ColorRole.Base, QColor(25, 25, 25))
            palette.setColor(QPalette.ColorRole.AlternateBase, QColor(53, 53, 53))
            palette.setColor(QPalette.ColorRole.ToolTipBase, Qt.GlobalColor.white)
            palette.setColor(QPalette.ColorRole.ToolTipText, Qt.GlobalColor.white)
            palette.setColor(QPalette.ColorRole.Text, Qt.GlobalColor.white)
            palette.setColor(QPalette.ColorRole.Button, QColor(53, 53, 53))
            palette.setColor(QPalette.ColorRole.ButtonText, Qt.GlobalColor.white)
            palette.setColor(QPalette.ColorRole.BrightText, Qt.GlobalColor.red)
            palette.setColor(QPalette.ColorRole.Highlight, QColor(142, 45, 197).lighter())
            palette.setColor(QPalette.ColorRole.HighlightedText, Qt.GlobalColor.white)
            
            # Ajusta estilos específicos
            self.log_text.setStyleSheet("""
                QTextEdit {
                    background-color: #2d2d2d;
                    color: #000000;
                    font-family: Consolas, monospace;
                    font-size: 10pt;
                    border: 1px solid #444;
                }
            """)
            self.tema_btn.setText("☀ Ativar Tema Claro")
        else:
            # Volta ao tema padrão (claro)
            palette = QApplication.style().standardPalette()
            
            self.log_text.setStyleSheet("""
            QTextEdit {
                background-color: #f9f9f9;
                font-family: Consolas, monospace;
                font-size: 10pt;
                border: 1px solid #ddd;
            }
        """)
        self.tema_btn.setText("🌙 Ativar Tema Escuro")
        
        # Aplica a paleta global
        QApplication.instance().setPalette(palette)
        
        # Atualiza o texto de todos os grupos
        for group in self.findChildren(QGroupBox):
            group.setStyleSheet("""
                QGroupBox {
                    font-weight: bold;
                    border: 1px solid %s;
                    border-radius: 4px;
                    margin-top: 10px;
                    padding-top: 12px;
                }
                QGroupBox::title {
                    subcontrol-origin: margin;
                    left: 10px;
                }
            """ % ("#555" if self.dark_mode else "#aaa"))

    def formatar_texto_negrito(self):
        """Adiciona formatação em negrito ao texto selecionado"""
        cursor = self.text_edit.textCursor()
        
        # Verifica se há texto selecionado
        if cursor.hasSelection():
            selected_text = cursor.selectedText()
            cursor.insertHtml(f"<b>{selected_text}</b>")
        else:
            # Se nada estiver selecionado, insere tags de negrito
            cursor.insertHtml("<b></b>")
            # Move o cursor entre as tags
            cursor.movePosition(QTextCursor.MoveOperation.Left, QTextCursor.MoveMode.MoveAnchor, 4)
            self.text_edit.setTextCursor(cursor)
    
    def load_config(self):
        """Carrega configurações de múltiplas fontes"""
        try: 
            config = configparser.ConfigParser() 
            config.read('config.ini') 
            
            self.smtp_server.setText(config.get('EMAIL', 'servidor_smtp', fallback='smtp.gmail.com')) 
            self.smtp_port.setText(config.get('EMAIL', 'porta_smtp', fallback='587')) 
            self.email_user.setText(config.get('EMAIL', 'usuario', fallback='')) 
            self.email_subject.setText(config.get('EMAIL', 'assunto', fallback='Manual de Uso - WL Pesos Padrão')) 
            self.log("Configurações carregadas do arquivo config.ini")   
        except Exception as e: 
            self.log(f"⚠ Erro ao carregar config: {str(e)}") 
    
    def save_config(self):   
        try: 

            # 1. Tenta carregar do .env
            load_dotenv()

            # 2. Carrega do config.ini se existir
            config = configparser.ConfigParser() 
            if os.path.exists('config.ini'):
                config.read('config.ini')

            # 3. Configurações Padrão + .env + config.ini
            self.smtp_server.setText(
                os.getenv('SMTP_SERVER') or
                config.get('EMAIL', 'servidor_smtp', fallback=DEFAULT_CONFIG['smtp_server'])
            )
            
            with open('config.ini', 'w') as configfile:   
                config.write(configfile) 
            
            self.log("✅ Configurações salvas com sucesso em config.ini")   

        except Exception as e: 
            self.log(f"❌ Erro ao carregar as configurações: {str(e)}")
    
    def log(self, message):   
        timestamp = time.strftime("%H:%M:%S", time.localtime())   
        self.log_text.append(f"[{timestamp}] {message}")   
        self.log_text.moveCursor(QTextCursor.MoveOperation.End)
    
    def selecionar_arquivos(self):
        '''Abre diálogo para seleção de arquivos anexos'''
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Selecione os anexos",
            self.attachments_dir,
            "",
            "Todos os arquivos (*);;PDF (*.pdf);;Imagens (*.png *.jpg)"
        )
        
        if files:
            self.anexos_selecionados = files
            self.anexos_label.setText(f"{len(files)} arquivo(s) selecionado(s)")
            self.anexos_label.setStyleSheet("color: #FFFFFF; font-style: normal;")
        else:
            self.anexos_selecionados = []
            self.anexos_label.setText("Nenhum arquivo selecionado")
            self.anexos_label.setStyleSheet("color: #666; font-style: italic;")
    
    def adicionar_destinatario(self):
        nome = self.nome_entry.text().strip()
        email = self.email_entry.text().strip()

        # Validação básica de e-mail
        if not "@" in email or "." not in email.split("@")[-1]:
            QMessageBox.warning(self, "E-mail inválido", "Por favor, insira um endereço de e-mail válido.")
            return
        
        if not nome or not email:
            QMessageBox.warning(self, "Campos Vazios", "Por favor, preencha o nome e e-mail.")
            return
        
        if not hasattr(self, 'anexos_selecionados') or not self.anexos_selecionados:
            reply = QMessageBox.question(
                self, 
                "Sem Anexos", 
                "Deseja continuar sem anexos?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.No:
                return
            
        # Converte caminhos para absolutos e verifica existência
        arquivos_validos = []
        for arquivo in self.anexos_selecionados:
            arquivo_abs = os.path.abspath(arquivo)
            if os.path.exists(arquivo_abs):
                arquivos_validos.append(arquivo_abs)
            else:
                QMessageBox.warning(self, "Arquivo não encontrado", 
                                  f"O arquivo será ignorado (não encontrado):\n{arquivo}")

        # Adiciona à árvore (mostra apenas nomes dos arquivos)
        nomes_arquivos = [os.path.basename(p) for p in arquivos_validos] if arquivos_validos else ["Nenhum anexo"]
        item = QTreeWidgetItem([nome, email, ", ".join(nomes_arquivos)])
        self.tree.addTopLevelItem(item)
        
        # Armazena os dados
        self.lista_para_envio.append({
            'nome': nome,
            'email': email,
            'arquivos': arquivos_validos
        })
        
        # Limpa os campos
        self.nome_entry.clear()
        self.email_entry.clear()
        self.anexos_selecionados = []
        self.anexos_label.setText("Nenhum arquivo selecionado")
        self.anexos_label.setStyleSheet("color: #666; font-style: italic;")
    
    def send_emails(self):
        if not self.lista_para_envio:
            QMessageBox.warning(self, "Lista Vazia", 
                              "Por favor, adicione pelo menos um destinatário.")
            return
            
        self.send_btn.setEnabled(False)
        self.send_btn.setText("Enviando...")
        
        # Cria e inicia uma thread para envio
        self.email_thread = EmailThread(
            self.lista_para_envio,
            self.smtp_server.text(),
            self.smtp_port.text(),
            self.email_user.text(),
            self.email_pass.text(),
            self.email_subject.text()
        )
        
        self.email_thread.update_signal.connect(self.log)
        self.email_thread.progress_signal.connect(self.progress.setValue)
        self.email_thread.finished_signal.connect(self.envio_finalizado)
        
        self.email_thread.start()
    
    def envio_finalizado(self, success):
        self.send_btn.setEnabled(True)
        self.send_btn.setText("▶ ENVIAR TUDO")
        
        if success:
            reply = QMessageBox.question(
                self, 
                "Envio Concluído", 
                "O processo de envio foi finalizado. Deseja limpar a lista agora?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            
            if reply == QMessageBox.StandardButton.Yes:
                self.limpar_lista()
    
    def limpar_lista(self):
        reply = QMessageBox.question(
            self, 
            "Confirmar", 
            "Tem certeza que deseja limpar toda a lista de envio?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            self.lista_para_envio.clear()
            self.tree.clear()
            self.log("ℹ️ Lista de envio foi limpa.")

    def importar_contatos_csv(self):
        """Importa lista de contatos de um arquivo CSV"""
        filename, _ = QFileDialog.getOpenFileName(
            self,
            "Selecionar arquivo CSV",
            "",
            "Arquivos CSV (*.csv);;Todos os arquivos (*)"
        )
        
        if not filename:
            return
        
        try:
            with open(filename, 'r', encoding='utf-8') as csvfile:
                # Detecta automaticamente o delimitador
                dialect = csv.Sniffer().sniff(csvfile.read(1024))
                csvfile.seek(0)
                reader = csv.reader(csvfile, dialect)
                
                # Pula cabeçalho se existir
                try:
                    if csv.Sniffer().has_header(csvfile.read(1024)):
                        csvfile.seek(0)
                        next(reader)  # Pula a linha de cabeçalho
                    csvfile.seek(0)
                except:
                    csvfile.seek(0)
                
                contatos_importados = 0
                for row in reader:
                    if len(row) >= 2:  # Pelo menos nome e e-mail
                        nome, email = row[0], row[1]
                        if "@" in email:  # Verificação básica de e-mail
                            # Adiciona à lista de envio
                            self.lista_para_envio.append({
                                'nome': nome.strip(),
                                'email': email.strip(),
                            })
                            
                            # Adiciona à visualização
                            item = QTreeWidgetItem([nome.strip(), email.strip(), ", "])
                            self.tree.addTopLevelItem(item)
                            contatos_importados += 1
                
                self.log(f"✅ {contatos_importados} contatos importados de {filename}")
                QMessageBox.information(self, "Importação concluída", 
                                    f"Foram importados {contatos_importados} contatos com sucesso!")
        
        except Exception as e:
            self.log(f"❌ Erro ao importar CSV: {str(e)}")
            QMessageBox.critical(self, "Erro na importação", 
                            f"Não foi possível ler o arquivo CSV:\n{str(e)}")
            
    def criar_editor_email(self):
        """Cria a seção de edição do corpo do e-mail"""
        # Grupo para o editor
        editor_group = QGroupBox("Editor de E-mail")
        editor_layout = QVBoxLayout()
        
        # Barra de ferramentas básica
        toolbar = QToolBar()
        
        # Botões de formatação
        btn_bold = QAction("Negrito", self)
        btn_bold.triggered.connect(lambda: self.text_edit.insertHtml("<b>texto</b>"))
        toolbar.addAction(btn_bold)
        
        # Área de edição
        self.text_edit = QTextEdit()
        self.text_edit.setAcceptRichText(True)
        
        # Carrega template existente se disponível
        if os.path.exists('corpo_email.txt'):
            template_padrao =  '''Prezado(a) %(nome)s,
            
            Segue em anexo os documetnos solicitados.
            
            Atenciosamente,
            Equipe'''
            with open('corpo_email.txt', 'w', encoding='utf-8') as f:
                f.write(template_padrao)

            with open('corpo_email.txt', 'r', encoding='utf-8') as f:
                self.text_edit.setPlainText(f.read())
        
        # Botões de ação
        btn_frame = QWidget()
        btn_layout = QHBoxLayout(btn_frame)
        
        btn_salvar = QPushButton("💾 Salvar Template")
        btn_salvar.clicked.connect(self.salvar_template_email)
        btn_salvar.setStyleSheet("background-color: #4CAF50; color: white;")
        
        btn_carregar = QPushButton("📂 Carregar Template")
        btn_carregar.clicked.connect(self.carregar_template_email)
        btn_carregar.setStyleSheet("background-color: #2196F3; color: white;")
        
        btn_layout.addWidget(btn_salvar)
        btn_layout.addWidget(btn_carregar)
        
        # Monta o layout completo
        editor_layout.addWidget(toolbar)
        editor_layout.addWidget(self.text_edit)
        editor_layout.addWidget(btn_frame)
        editor_group.setLayout(editor_layout)
        
        return editor_group

    def salvar_template_email(self):
        """Salva o template do e-mail em um arquivo"""
        filename, _ = QFileDialog.getSaveFileName(
            self,
            "Salvar template de e-mail",
            "",
            "Arquivos de texto (*.txt);;Todos os arquivos (*)"
        )
        
        if filename:
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(self.text_edit.toPlainText())
                self.log(f"✅ Template salvo em {filename}")
            except Exception as e:
                self.log(f"❌ Erro ao salvar template: {str(e)}")

    def carregar_template_email(self):
        """Carrega um template de e-mail existente"""
        filename, _ = QFileDialog.getOpenFileName(
            self,
            "Carregar template de e-mail",
            "",
            "Arquivos de texto (*.txt);;Todos os arquivos (*)"
        )
        
        if filename:
            try:
                with open(filename, 'r', encoding='utf-8') as f:
                    self.text_edit.setPlainText(f.read())
                self.log(f"✅ Template carregado de {filename}")
            except Exception as e:
                self.log(f"❌ Erro ao carregar template: {str(e)}")

class EmailThread(QThread):
    update_signal = pyqtSignal(str)
    progress_signal = pyqtSignal(int)
    finished_signal = pyqtSignal(bool)
    
    def __init__(self, lista_para_envio, smtp_server, smtp_port, email_user, email_pass, email_subject):
        super().__init__()
        self.lista_para_envio = lista_para_envio
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        self.email_user = email_user
        self.email_pass = email_pass
        self.email_subject = email_subject
    
    def run(self):
        try:
            # Verificar configurações
            if not all([self.smtp_server, self.smtp_port, self.email_user, self.email_pass]):
                self.update_signal.emit("⚠ Erro: Preencha todos os campos de configuração do e-mail e a senha.")
                self.finished_signal.emit(False)
                return

            total_a_enviar = len(self.lista_para_envio)
            if total_a_enviar == 0:
                self.update_signal.emit("ℹ️ A lista de envio está vazia. Adicione destinatários primeiro.")
                self.finished_signal.emit(False)
                return

            # Conectar ao servidor SMTP
            porta = int(self.smtp_port)
            self.update_signal.emit(f"⏳ Conectando a {self.smtp_server} na porta {porta}...")

            if porta == 465:
                server = smtplib.SMTP_SSL(host=self.smtp_server, port=porta, timeout=20)
            else:
                server = smtplib.SMTP(host=self.smtp_server, port=porta, timeout=20)
                server.starttls()
            
            self.update_signal.emit(f"🔑 Autenticando usuário {self.email_user}...")
            server.login(self.email_user, self.email_pass)
            self.update_signal.emit("✅ Conexão e login realizados com sucesso!")

            # Ler template do corpo do e-mail
            with open('corpo_email.txt', 'r', encoding='utf-8') as f:
                corpo_template = self.text_edit.toPlainText()

            # Enviar e-mails
            self.update_signal.emit(f"📤 Iniciando envio para {total_a_enviar} destinatário(s)...")
            
            for i, destinatario in enumerate(self.lista_para_envio, 1):
                try:
                    nome = destinatario['nome']
                    email_dest = destinatario['email']
                    lista_de_anexos = destinatario['arquivos']

                    self.update_signal.emit(f"Preparando e-mail para {email_dest} com {len(lista_de_anexos)} anexo(s)")
                    
                    msg = MIMEMultipart()
                    msg['From'] = self.email_user
                    msg['To'] = email_dest
                    msg['Subject'] = self.email_subject

                    corpo_personalizado = corpo_template.replace('%(nome)s', nome)
                    msg.attach(MIMEText(corpo_personalizado, 'plain'))

                    # Log detalhado dos anexos
                    for idx, caminho_arquivo in enumerate(lista_de_anexos, 1):
                        self.update_signal.emit(f"  Verificando anexo {idx}: {caminho_arquivo}")
                        if not os.path.exists(caminho_arquivo):
                            self.update_signal.emit(f"  ⚠ Arquivo não encontrado: {caminho_arquivo}")
                            continue

                        try:
                            with open(caminho_arquivo, 'rb') as attachment:
                                part = MIMEBase('application', 'octet-stream')
                                part.set_payload(attachment.read())
                                encoders.encode_base64(part)
                                nome_arquivo = os.path.basename(caminho_arquivo)
                                part.add_header(
                                    'Content-Disposition', 
                                    f"attachment; filename={nome_arquivo}"
                                )
                                msg.attach(part)
                                self.update_signal.emit(f"  ✅ Anexo adicionado: {nome_arquivo}")
                        except Exception as e:
                            self.update_signal.emit(f"  ❌ Erro ao processar anexo: {str(e)}")
                            continue

                    server.send_message(msg)
                    self.update_signal.emit(f"✅ ({i}/{total_a_enviar}) E-mail enviado para: {email_dest}")
                    
                    # Atualizar progresso
                    progresso = int((i / total_a_enviar) * 100)
                    self.progress_signal.emit(progresso)
                    
                    time.sleep(1)

                except Exception as e:
                    self.update_signal.emit(f"❌ Falha no envio para {email_dest}: {str(e)}")
                    continue
            
            server.quit()
            self.update_signal.emit("\n🎉 Processo finalizado!")
            self.finished_signal.emit(True)

        except smtplib.SMTPAuthenticationError:
            self.update_signal.emit("⛔ ERRO DE AUTENTICAÇÃO: Usuário ou senha incorretos.")
            self.update_signal.emit("ℹ️ DICA: Se usar Gmail/Outlook com 2FA, você precisa de uma 'Senha de App'.")
            self.finished_signal.emit(False)
        except Exception as e:
            self.update_signal.emit(f"⛔ Erro crítico durante o processo: {str(e)}")
            self.finished_signal.emit(False)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # Aplicar um estilo moderno
    app.setStyle("Fusion")
    
    window = EmailSenderApp()
    window.show()
    sys.exit(app.exec())