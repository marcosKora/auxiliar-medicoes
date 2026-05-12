import time
import csv
import os
import re
import requests
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import sys
import json
import threading

from PyQt6.QtWidgets import *
from PyQt6.QtCore import *
from PyQt6.QtCore import QDate  # 999999: para seletor de data da metrica
from PyQt6.QtGui import *
from qt_material import apply_stylesheet

def resource_path(relative_path):
    """ Retorna o caminho absoluto para o recurso, funciona em dev e no PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# --- CONFIGURAÇÕES DE ACESSO ---
URL_CONTROLE = "https://robomedicoesaccesscontrol-default-rtdb.firebaseio.com/.json"

def verificar_acesso_remoto():
    try:
        try:
            nome_usuario = os.getlogin().upper().replace(".", "_")
        except:
            nome_usuario = os.environ.get('USERNAME', 'DESCONHECIDO').upper().replace(".", "_")

        response = requests.get(URL_CONTROLE, timeout=10)
        dados = response.json()
        
        # Pega a lista de autorizados
        autorizados = dados.get("autorizados", {})
        
        if autorizados and autorizados.get(nome_usuario) is True:
            return True
        else:
            app = QApplication.instance()
            if app is None:
                app = QApplication(sys.argv)
            QMessageBox.critical(None, "Acesso Negado", f"{nome_usuario} - Solicite acesso.")
            return False
    except Exception as e:
        app = QApplication.instance()
        if app is None:
            app = QApplication(sys.argv)
        QMessageBox.critical(None, "Erro de Conexão", "Não foi possível validar o acesso.")
        return False

def verificar_versao():
    try:
        response = requests.get(URL_CONTROLE, timeout=10)
        dados = response.json()
        return dados.get("versao_minima", "0.0.0")
    except:
        return "0.0.0"


# --- CONFIGURAÇÕES DE ARQUIVOS ---
CONFIG_FILE = "credenciais.txt"
BACKUP_FILE = "historico_pedidos.txt"
ERROR_FILE = "erros_pedidos.txt"

def carregar_credenciais():
    creds = {"V360_USER": "", "V360_PASS": "", "SAP_USER": "", "SAP_PASS": "", "KORA_MED_PASS": ""}
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as f:
            for linha in f:
                if "=" in linha:
                    k, v = linha.strip().split("=", 1)
                    if k in creds: creds[k] = v
    return creds

def salvar_credencial(chave, valor):
    creds = carregar_credenciais()
    creds[chave] = valor
    with open(CONFIG_FILE, "w") as f:
        for k, v in creds.items():
            f.write(f"{k}={v}\n")

def salvar_backup(id_v360, resultado):
    horario = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    with open(BACKUP_FILE, "a", encoding="utf-8") as f:
        f.write(f"[{horario}] ID: {id_v360} - {resultado}\n")

def salvar_erro_txt(id_v360, erro):
    horario = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    with open(ERROR_FILE, "a", encoding="utf-8") as f:
        f.write(f"[{horario}] ID: {id_v360} - {erro}\n")

# 999999: salva metrica no csv para historico
def salvar_metrica(id_v, status):
    agora = datetime.now()
    with open("metricas.csv", "a", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerow([agora.strftime("%d/%m/%Y"), agora.strftime("%H:%M:%S"), id_v, status])        

# 999999: carrega metricas do csv com filtro de data
def carregar_metricas(data_inicio=None, data_fim=None):
    if not os.path.exists("metricas.csv"):
        return {"sucesso": 0, "erro": 0, "total": 0}
    sucesso = 0
    erro = 0
    with open("metricas.csv", "r", encoding="utf-8") as f:
        reader = csv.reader(f)
        for row in reader:
            if len(row) < 4:
                continue
            data_linha, hora, id_v, status = row
            if data_inicio and data_fim:
                if data_inicio <= data_linha <= data_fim:
                    if status == "sucesso": sucesso += 1
                    else: erro += 1
            elif data_inicio:
                if data_linha == data_inicio:
                    if status == "sucesso": sucesso += 1
                    else: erro += 1
    return {"sucesso": sucesso, "erro": erro, "total": sucesso + erro}

# --- INTERFACE ---

# ═══ NOVO: Splash Screen Moderno (Tela Cheia) ═══
class SplashScreen(QSplashScreen):
    def __init__(self):
        screen = QApplication.primaryScreen().geometry()
        w = screen.width()
        h = screen.height()
        
        pixmap = QPixmap(w, h)
        pixmap.fill(QColor("#1e1e1e"))
        
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        
        # Título
        font_title = QFont("Segoe UI", 36, QFont.Weight.Bold)
        painter.setFont(font_title)
        painter.setPen(QColor("#123b51"))
        painter.drawText(QRect(0, int(h * 0.24), w, 80), Qt.AlignmentFlag.AlignCenter, "MEDIÇÕES v3")
        
        # Subtítulo
        font_sub = QFont("Segoe UI", 16)
        painter.setFont(font_sub)
        painter.setPen(QColor("#888888"))
        painter.drawText(QRect(0, int(h * 0.35), w, 40), Qt.AlignmentFlag.AlignCenter, "Automação de Processos")
        
        # Linha decorativa
        painter.setPen(QPen(QColor("#123b51"), 2))
        cx = w // 2
        painter.drawLine(cx - 150, int(h * 0.43), cx + 150, int(h * 0.43))
        
        # Ícone
        font_icon = QFont("Segoe UI", 40)
        painter.setFont(font_icon)
        painter.drawText(QRect(0, int(h * 0.48), w, 70), Qt.AlignmentFlag.AlignCenter, "🤖")
        
        # Barra de fundo
        bar_w = 500
        bar_x = (w - bar_w) // 2
        bar_y = int(h * 0.78)
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(QColor("#2d2d2d"))
        painter.drawRoundedRect(bar_x, bar_y, bar_w, 6, 3, 3)
        
        # Versão
        font_version = QFont("Segoe UI", 9)
        painter.setFont(font_version)
        painter.setPen(QColor("#555555"))
        painter.drawText(QRect(0, int(h * 0.90), w, 25), Qt.AlignmentFlag.AlignCenter, "v3.0.0 • Kora Saúde")
        
        painter.end()
        
        super().__init__(pixmap)
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint)
        self.move(0, 0)
        
        self._draw_progress(0)
    
    def _draw_progress(self, percent):
        screen = QApplication.primaryScreen().geometry()
        w = screen.width()
        h = screen.height()
        
        pixmap = self.pixmap()
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        
        bar_w = 500
        bar_x = (w - bar_w) // 2
        bar_y = int(h * 0.78)
        
        # Fundo
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(QColor("#2d2d2d"))
        painter.drawRoundedRect(bar_x, bar_y, bar_w, 6, 3, 3)
        
        # Preenchimento
        if percent > 0:
            painter.setBrush(QColor("#123b51"))
            progress_width = int(bar_w * percent / 100)
            painter.drawRoundedRect(bar_x, bar_y, progress_width, 6, 3, 3)
        
        # Texto
        if percent == 100:
            font_load = QFont("Segoe UI", 9)
            painter.setFont(font_load)
            painter.setPen(QColor("#777777"))
            painter.drawText(QRect(0, int(h * 0.80), w, 25), Qt.AlignmentFlag.AlignCenter, "Pronto!")
        
        painter.end()
        self.setPixmap(pixmap)
    
    def set_progress(self, percent):
        self._draw_progress(percent)
        QApplication.processEvents()



class App(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Medições v3 - Unificado")
        # ═══ NOVO: Centralizar na tela ═══
        screen = QApplication.primaryScreen().geometry()
        window_width = screen.width()    # ← ANTES: 1200
        window_height = screen.height()  # ← ANTES: 750
        x = 0                            # ← ANTES: (screen.width() - 1200) // 2
        y = 0                            # ← ANTES: (screen.height() - 750) // 2
        self.setGeometry(x, y, window_width, window_height)
    # ═══════════════════════════════

        self.ids_processar = []
        self.cont_sucesso = 0
        self.cont_erro = 0
        self.cont_total = 0


        # ═══ NOVO: Variáveis da barra de progresso ═══
        self.progress_bar = None
        self.progress_label = None
        # ═══════════════════════════════════════════
        
        # widget central
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # sidebar
        sidebar = QWidget()
        sidebar.setFixedWidth(240)
        sidebar.setStyleSheet("background-color: #1e1e1e; border-right: 1px solid #333;")
        sidebar_layout = QVBoxLayout(sidebar)
        sidebar_layout.setContentsMargins(20, 30, 20, 20)
        
        logo = QLabel("AUTOMAÇÃO")
        logo.setStyleSheet("color: white; font-size: 22px; font-weight: bold;")
        logo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        sidebar_layout.addWidget(logo)
        sidebar_layout.addSpacing(30)
        
        # dropdown de configs
        btn_config = QPushButton("⚙️ Configurações")
        btn_config.setStyleSheet("""
            QPushButton {
                background-color: #2d2d2d; color: #ccc; border-radius: 10px;
                padding: 10px; font-size: 13px; text-align: center; border: none;
            }
            QPushButton:hover { background-color: #3d3d3d; }
            QPushButton::menu-indicator { image: none; }
        """)

        menu_config = QMenu()
        menu_config.setStyleSheet("""
            QMenu { 
                background-color: #2d2d2d; 
                border: 1px solid #555; 
                border-radius: 6px; 
                padding: 2px;
            }
            QMenu::item { 
                padding: 8px 20px; 
                color: #ccc; 
                font-size: 13px;
                width: 138px;
                text-align: center;
                padding-left: 25px;                  
            }
            QMenu::item:selected { 
                background-color: #123b51; 
                color: white;
                border-radius: 0px;
            }
        """)
        
        action_credenciais = menu_config.addAction("🔑 Credenciais")
        action_credenciais.triggered.connect(self.abrir_credenciais)
        action_limpar = menu_config.addAction("🧹 Limpar Logs")
        action_limpar.triggered.connect(self.limpar_logs)

        btn_config.setMenu(menu_config)
        btn_config.setFixedWidth(200)
        sidebar_layout.addWidget(btn_config)
        sidebar_layout.addSpacing(10)
        

        # ═══ NOVO: Contador de métricas na Sidebar ═══
        self.metrics_label = QLabel("📊 Sessão Atual | ✅ 0 | ❌ 0 | ⏳ 0")
        self.metrics_label.setStyleSheet("""
            QLabel {
                color: #aaa;
                font-size: 11px;
                padding: 8px;
                background-color: #252525;
                border-radius: 6px;
                text-align: center;
            }
        """)
        self.metrics_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.metrics_label.setWordWrap(True)
        sidebar_layout.addWidget(self.metrics_label)
        # ════════════════════════════
        
        # 999999: Dropdown de metricas na Sidebar (so aparece na tab metricas)
        self.metricas_sidebar_btn = QPushButton("📅 Hoje")
        self.metricas_sidebar_btn.setFixedHeight(30)
        self.metricas_sidebar_btn.setVisible(False)
        self.metricas_sidebar_btn.setStyleSheet("""
            QPushButton {
                background-color: #252525; color: white; padding: 5px 10px;
                border: 1px solid #444; border-radius: 5px; font-size: 11px; text-align: left;
            }
            QPushButton:hover { border-color: #123b51; }
            QPushButton::menu-indicator { subcontrol-position: right center; right: 5px; }
        """)
        
        self.metricas_sidebar_menu = QMenu()
        self.metricas_sidebar_menu.setStyleSheet("""
            QMenu {
                background-color: #2d2d2d; border: 1px solid #555;
                border-radius: 4px; padding: 3px; min-width: 180px;
            }
            QMenu::item { padding: 8px 15px; color: #ccc; font-size: 12px; }
            QMenu::item:selected { background-color: #123b51; color: white; border-radius: 3px; }
        """)
        
        opcoes_sidebar = ["🔄 Sessão Atual", "📅 Hoje", "📅 Ontem", "📅 Selecionar Data", "📅 Período"]
        for opcao in opcoes_sidebar:
            action = self.metricas_sidebar_menu.addAction(opcao)
            action.triggered.connect(lambda checked, o=opcao: self.selecionar_metrica(o))
        
        self.metricas_sidebar_btn.setMenu(self.metricas_sidebar_menu)
        sidebar_layout.addWidget(self.metricas_sidebar_btn)         

        # 999999: Datas na Sidebar (so na tab metricas)
        self.metricas_data_widget = QWidget()
        self.metricas_data_widget.setVisible(False)
        self.metricas_data_widget.setStyleSheet("background-color: #252525; border-radius: 6px; padding: 8px;")
        data_layout = QVBoxLayout(self.metricas_data_widget)
        data_layout.setSpacing(5)
        
        self.metricas_data_inicio = QLineEdit()
        self.metricas_data_inicio.setPlaceholderText("dd/mm/aaaa")
        self.metricas_data_inicio.setText(datetime.now().strftime("%d/%m/%Y"))
        self.metricas_data_inicio.setFixedHeight(28)
        self.metricas_data_inicio.setStyleSheet("""
            QLineEdit { background-color: #2d2d2d; color: white; padding: 4px; border: 1px solid #555; border-radius: 4px; font-size: 11px; }
            QLineEdit:hover { border-color: #123b51; }
        """)
        self.metricas_data_inicio.returnPressed.connect(lambda: [self.formatar_data(self.metricas_data_inicio), self.atualizar_tab_metricas()])
        data_layout.addWidget(QLabel("📅 Início:"))
        data_layout.addWidget(self.metricas_data_inicio)
        
        self.metricas_data_fim = QLineEdit()
        self.metricas_data_fim.setPlaceholderText("dd/mm/aaaa")
        self.metricas_data_fim.setText(datetime.now().strftime("%d/%m/%Y"))
        self.metricas_data_fim.setFixedHeight(28)
        self.metricas_data_fim.setStyleSheet("""
            QLineEdit { background-color: #2d2d2d; color: white; padding: 4px; border: 1px solid #555; border-radius: 4px; font-size: 11px; }
            QLineEdit:hover { border-color: #123b51; }
        """)
        self.metricas_data_fim.returnPressed.connect(lambda: [self.formatar_data(self.metricas_data_fim), self.atualizar_tab_metricas()])
        data_layout.addWidget(QLabel("📅 Fim:"))
        data_layout.addWidget(self.metricas_data_fim)
        
        sidebar_layout.addWidget(self.metricas_data_widget)
         

        # ═══ NOVO: Barra de Progresso na Sidebar ═══
        # Label de progresso
        self.progress_label = QLabel("Pronto para iniciar")
        self.progress_label.setStyleSheet("""
            QLabel {
                color: #888;
                font-size: 12px;
                padding: 5px;
                text-align: center;
            }
        """)
        self.progress_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.progress_label.setWordWrap(True)
        sidebar_layout.addWidget(self.progress_label)
        
        # Barra de progresso
        self.progress_bar = QProgressBar()
        self.progress_bar.setFixedWidth(200)
        self.progress_bar.setFixedHeight(20)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                background-color: #2d2d2d;
                border: 2px solid #444;
                border-radius: 10px;
                text-align: center;
                color: white;
                font-size: 11px;
                font-weight: bold;
            }
            QProgressBar::chunk {
                background-color: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #123b51, stop:1 #1a6b8a);
                border-radius: 8px;
            }
        """)
        self.progress_bar.setVisible(False)
        sidebar_layout.addWidget(self.progress_bar, alignment=Qt.AlignmentFlag.AlignCenter)
        # ═══════════════════════════════════════════

        sidebar_layout.addStretch()
        
        main_layout.addWidget(sidebar)
        
        # separador
        separator = QWidget()
        separator.setFixedWidth(1)
        separator.setStyleSheet("background-color: #333;")
        main_layout.addWidget(separator)
        
        # area principal das tabs
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(20, 20, 20, 20)
        
        header = QWidget()
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(0, 0, 0, 10)
        header_label = QLabel("🚀 Painel de Controle")
        header_label.setStyleSheet("color: white; font-size: 16px; font-weight: bold;")
        header_layout.addWidget(header_label)
        header_layout.addStretch()
        right_layout.addWidget(header)
        
        # tabs design
        self.tabview = QTabWidget()
        self.tabview.setStyleSheet("""
            QTabWidget::pane { border: 1px solid #333; background-color: #1e1e1e; }
            QTabBar::tab { 
                background-color: #2d2d2d; color: #999; padding: 10px 20px;
                border-top-left-radius: 8px; border-top-right-radius: 8px;
                margin-right: 2px;
            }
            QTabBar::tab:selected { background-color: #123b51; color: white; }
            QTabBar::tab:hover { background-color: #3d3d3d; }
            QTabBar::tab:last { 
                background-color: #123b51; color: white; font-weight: bold; min-width: 140px;
            }
            QTabBar::tab:last:hover { background-color: #0d2d3f; }
            QTabBar::tab:last:selected { background-color: #123b51; color: white; }
        """)
        
        # tab de entrada de ids
        tab_entrada = QWidget()
        tab_entrada_layout = QVBoxLayout(tab_entrada)
        self.txt_ids = QTextEdit()
        self.txt_ids.setPlaceholderText("# Cole os IDs aqui (um por linha)")
        self.txt_ids.setStyleSheet("""
            QTextEdit {
                background-color: #2d2d2d; color: white; border: 2px solid #555;
                border-radius: 10px; padding: 10px; font-family: Consolas; font-size: 13px;
            }
            QTextEdit:focus { border-color: #123b51; }
        """)
        tab_entrada_layout.addWidget(self.txt_ids)
        self.tabview.addTab(tab_entrada, "Entrada de IDs")
        
        # tab da log de execução
        tab_log = QWidget()
        tab_log_layout = QVBoxLayout(tab_log)
        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setStyleSheet("""
            QTextEdit {
                background-color: #2d2d2d; color: white; border: 2px solid #555;
                border-radius: 10px; padding: 10px; font-family: Consolas; font-size: 12px;
            }
        """)
        tab_log_layout.addWidget(self.log_area)
        self.tabview.addTab(tab_log, "Log de Execução")
        
        # tab da log de ids com sucesso
        tab_success = QWidget()
        tab_success_layout = QVBoxLayout(tab_success)
        self.success_area = QTextEdit()
        self.success_area.setReadOnly(True)
        self.success_area.setStyleSheet("""
            QTextEdit {
                background-color: #2d2d2d; color: #2ecc71; border: 2px solid #27ae60;
                border-radius: 10px; padding: 10px; font-family: Consolas; font-size: 12px;
            }
        """)
        tab_success_layout.addWidget(self.success_area)
        self.tabview.addTab(tab_success, "Pedidos Prontos")
        
        # tab da log de ids com erro
        tab_error = QWidget()
        tab_error_layout = QVBoxLayout(tab_error)
        self.error_area = QTextEdit()
        self.error_area.setReadOnly(True)
        self.error_area.setStyleSheet("""
            QTextEdit {
                background-color: #2d2d2d; color: #e74c3c; border: 2px solid #c0392b;
                border-radius: 10px; padding: 10px; font-family: Consolas; font-size: 12px;
            }
        """)
        tab_error_layout.addWidget(self.error_area)
        self.tabview.addTab(tab_error, "IDs com Erro")
        
        # 999999: NOVA TAB DE MÉTRICAS
        tab_metricas = QWidget()
        tab_metricas_layout = QVBoxLayout(tab_metricas)
        tab_metricas_layout.setSpacing(15)

        

        
        # Label de resultado
        self.metricas_resultado = QLabel("✅ 0 | ❌ 0 | 📊 0")
        self.metricas_resultado.setStyleSheet("""
            QLabel {
                color: white; font-size: 24px; font-weight: bold; padding: 25px;
                background-color: #252525; border-radius: 12px;
            }
        """)
        self.metricas_resultado.setAlignment(Qt.AlignmentFlag.AlignCenter)
        tab_metricas_layout.addWidget(self.metricas_resultado)
        
        # 999999: Label de resumo da data selecionada
        self.metricas_data_resumo = QLabel("")
        self.metricas_data_resumo.setStyleSheet("""
            QLabel {
                color: #aaa; font-size: 13px; padding: 10px;
                background-color: #252525; border-radius: 8px;
            }
        """)
        self.metricas_data_resumo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        tab_metricas_layout.addWidget(self.metricas_data_resumo)
        
        tab_metricas_layout.addStretch()

        self.tabview.addTab(tab_metricas, "📊 Métricas")
        
        # tab do botão de iniciar o processo
        tab_iniciar = QWidget()
        tab_iniciar.setStyleSheet("background-color: #1e1e1e;")
        self.tabview.addTab(tab_iniciar, "▶ INICIAR")
        
        # conector do clique de cada tab
        self.tabview.tabBarClicked.connect(self.tab_clicada)
        self.tabview.currentChanged.connect(self.atualizar_visibilidade_sidebar)  # 999999

        right_layout.addWidget(self.tabview)
        main_layout.addWidget(right_widget)
        
        # barra de status
        self.statusBar().showMessage("Pronto")
        self.statusBar().setStyleSheet("background-color: #1e1e1e; color: #888;")

    def toggle_senha(self, campo, botao, checked):
        if checked:
            campo.setEchoMode(QLineEdit.EchoMode.Normal)
            botao.setText("🔒")
        else:
            campo.setEchoMode(QLineEdit.EchoMode.Password)
            botao.setText("👁️")


    def tab_clicada(self, index):
        if index == 5:
            self.validar_e_rodar()
            self.tabview.setCurrentIndex(1)

    def atualizar_log(self, mensagem, cor=None):
        tag = f"[{datetime.now().strftime('%H:%M:%S')}] "
        full_msg = f"{tag}{mensagem}\n"
        self.log_area.append(full_msg)
        self.log_area.verticalScrollBar().setValue(self.log_area.verticalScrollBar().maximum())
        # QApplication.processEvents()

    def registrar_sucesso(self, id_v, msg):
        texto = f"{datetime.now().strftime('%H:%M:%S')} - ID: {id_v} - {msg}\n"
        self.success_area.append(texto)
        salvar_backup(id_v, msg)
        salvar_metrica(id_v, "sucesso")        
        self.cont_sucesso += 1
        self.atualizar_metricas()

    def registrar_erro(self, id_v, msg):
        texto = f"{datetime.now().strftime('%H:%M:%S')} - ID: {id_v} - {msg}\n"
        self.error_area.append(texto)
        salvar_erro_txt(id_v, msg)
        salvar_metrica(id_v, "erro")
        self.cont_erro += 1
        self.atualizar_metricas()

    def limpar_logs(self):
        reply = QMessageBox.question(self, "Limpar", "Deseja limpar todos os campos de entrada e logs?",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            self.txt_ids.clear()
            self.log_area.clear()
            self.success_area.clear()
            self.error_area.clear()

    def atualizar_metricas(self):
        total = self.cont_sucesso + self.cont_erro
        fila = len(self.ids_processar) - total if len(self.ids_processar) > total else 0
        self.metrics_label.setText(f"📊 Sessão | 📊 {total} | ✅ {self.cont_sucesso} | ❌ {self.cont_erro} | ⏳ {fila}")

    # 999999: funcao chamada quando seleciona opcao no menu dropdown
    def selecionar_metrica(self, opcao):
        self.metricas_sidebar_btn.setText(opcao)
        mostrar_datas = "Data" in opcao or "Período" in opcao
        self.metricas_data_widget.setVisible(mostrar_datas)
        self.metricas_data_fim.setVisible("Período" in opcao)
        self.atualizar_tab_metricas(opcao)   

    # 999999: mostra/esconde dropdown da sidebar conforme a tab
    def atualizar_visibilidade_sidebar(self, index):
        visivel = (index == 4)
        self.metricas_sidebar_btn.setVisible(visivel)
        opcao = self.metricas_sidebar_btn.text()
        mostrar_datas = visivel and ("Data" in opcao or "Período" in opcao)
        self.metricas_data_widget.setVisible(mostrar_datas)   

    # 999999: atualiza a tab de metricas baseado no periodo selecionado
    def atualizar_tab_metricas(self, selecao=None):
        if selecao is None:
            selecao = self.metricas_sidebar_btn.text()
        if hasattr(selecao, 'toString'):
            selecao = self.metricas_sidebar_btn.text()
        
        # 999999: Força pegar o texto do botão para saber qual opção está selecionada
        opcao_atual = self.metricas_sidebar_btn.text()
        
        hoje = datetime.now().strftime("%d/%m/%Y")
        ontem = (datetime.now() - timedelta(days=1)).strftime("%d/%m/%Y")
        
        if "Sessão Atual" in opcao_atual:
            m = {"sucesso": self.cont_sucesso, "erro": self.cont_erro, "total": self.cont_sucesso + self.cont_erro}
        elif "Hoje" in opcao_atual:
            m = carregar_metricas(data_inicio=hoje)
        elif "Ontem" in opcao_atual:
            m = carregar_metricas(data_inicio=ontem)
        elif "Selecionar Data" in opcao_atual:
            data_sel = self.metricas_data_inicio.text().strip()
            m = carregar_metricas(data_inicio=data_sel)
        elif "Período" in opcao_atual:
            data_ini = self.metricas_data_inicio.text().strip()
            data_fim = self.metricas_data_fim.text().strip()
            m = carregar_metricas(data_inicio=data_ini, data_fim=data_fim)
        else:
            m = {"sucesso": 0, "erro": 0, "total": 0}
        
        # 999999: Atualiza resumo da data
        if "Sessão Atual" in opcao_atual:
            self.metricas_data_resumo.setText("")
        elif "Hoje" in opcao_atual:
            self.metricas_data_resumo.setText(f"📅 Data: {hoje}")
        elif "Ontem" in opcao_atual:
            self.metricas_data_resumo.setText(f"📅 Data: {ontem}")
        elif "Selecionar Data" in opcao_atual:
            data_sel = self.metricas_data_inicio.text().strip()
            self.metricas_data_resumo.setText(f"📅 Data: {data_sel}")
        elif "Período" in opcao_atual:
            data_ini = self.metricas_data_inicio.text().strip()
            data_fim = self.metricas_data_fim.text().strip()
            self.metricas_data_resumo.setText(f"📅 Data: {data_ini} - {data_fim}")        

        self.metricas_resultado.setText(f"✅ {m['sucesso']} | ❌ {m['erro']} | 📊 {m['total']}")


    # 999999: formata data automaticamente ao digitar
    def formatar_data(self, campo):
        texto = campo.text().strip().replace("/", "").replace(".", "").replace("-", "")
        if len(texto) == 8 and texto.isdigit():
            campo.setText(f"{texto[0:2]}/{texto[2:4]}/{texto[4:8]}")

    def abrir_credenciais(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Gerenciar Credenciais")
        dialog.setFixedSize(400, 500)
        dialog.setStyleSheet("background-color: #1e1e1e;")
        
        layout = QVBoxLayout(dialog)
        creds = carregar_credenciais()
        
        # campos de credenciais do v360 no painel, para o usuario alterar quando necessário
        layout.addWidget(QLabel("VIRTUAL 360"))
        layout.itemAt(layout.count() - 1).widget().setStyleSheet("color: white; font-weight: bold; font-size: 14px;")
        
        e_v360 = QLineEdit(creds["V360_USER"])
        e_v360.setPlaceholderText("E-mail V360")
        e_v360.setStyleSheet("background-color: #2d2d2d; color: white; padding: 8px; border-radius: 5px;")
        layout.addWidget(e_v360)
        
        senha_v360_widget = QWidget()
        senha_v360_layout = QHBoxLayout(senha_v360_widget)
        senha_v360_layout.setContentsMargins(0,0,0,0)

        s_v360 = QLineEdit(creds["V360_PASS"])
        s_v360.setPlaceholderText("Senha V360")
        s_v360.setEchoMode(QLineEdit.EchoMode.Password)
        s_v360.setStyleSheet("background-color: #2d2d2d; color: white; padding: 8px; border-radius: 5px;")

        btn_olho_v360 = QPushButton("👁️")
        btn_olho_v360.setFixedSize(50,40)
        btn_olho_v360.setCheckable(True)
        btn_olho_v360.setStyleSheet("background-color: #2d2d2d; border: 1px solid #555; border-radius: 5px;")
        btn_olho_v360.clicked.connect(lambda checked: self.toggle_senha(s_v360, btn_olho_v360, checked))

        senha_v360_layout.addWidget(s_v360)
        senha_v360_layout.addWidget(btn_olho_v360)
        layout.addWidget(senha_v360_widget)
        
        btn_save_v360 = QPushButton("Salvar V360")
        btn_save_v360.setStyleSheet("background-color: #123b51; color: white; padding: 8px; border-radius: 5px;")
        btn_save_v360.clicked.connect(lambda: [salvar_credencial("V360_USER", e_v360.text()), 
                                                salvar_credencial("V360_PASS", s_v360.text()), 
                                                QMessageBox.information(dialog, "OK", "V360 Salvo!")])
        layout.addWidget(btn_save_v360)
        
        # campos de credenciais do SAP no painel, para o usuario alterar quando necessário
        layout.addWidget(QLabel("SAP"))
        layout.itemAt(layout.count() - 1).widget().setStyleSheet("color: white; font-weight: bold; font-size: 14px;")
        
        e_sap = QLineEdit(creds["SAP_USER"])
        e_sap.setPlaceholderText("Usuário SAP")
        e_sap.setStyleSheet("background-color: #2d2d2d; color: white; padding: 8px; border-radius: 5px;")
        layout.addWidget(e_sap)
        
        senha_sap_widget = QWidget()
        senha_sap_layout = QHBoxLayout(senha_sap_widget)
        senha_sap_layout.setContentsMargins(0,0,0,0)

        s_sap = QLineEdit(creds["SAP_PASS"])
        s_sap.setPlaceholderText("Senha SAP")
        s_sap.setEchoMode(QLineEdit.EchoMode.Password)
        s_sap.setStyleSheet("background-color: #2d2d2d; color: white; padding: 8px; border-radius: 5px;")

        btn_olho_sap = QPushButton("👁️")
        btn_olho_sap.setFixedSize(50,40)
        btn_olho_sap.setCheckable(True)
        btn_olho_sap.setStyleSheet("background-color: #2d2d2d; border: 1px solid #555; border-radius: 5px;")
        btn_olho_sap.clicked.connect(lambda checked: self.toggle_senha(s_sap, btn_olho_sap, checked))

        senha_sap_layout.addWidget(s_sap)
        senha_sap_layout.addWidget(btn_olho_sap)
        layout.addWidget(senha_sap_widget)
        
        btn_save_sap = QPushButton("Salvar SAP")
        btn_save_sap.setStyleSheet("background-color: #123b51; color: white; padding: 8px; border-radius: 5px;")
        btn_save_sap.clicked.connect(lambda: [salvar_credencial("SAP_USER", e_sap.text()), 
                                               salvar_credencial("SAP_PASS", s_sap.text()), 
                                               QMessageBox.information(dialog, "OK", "SAP Salvo!")])
        layout.addWidget(btn_save_sap)
        
        # campos de credenciais do site Kora-medicoes no painel, para o usuario alterar quando necessário
        layout.addWidget(QLabel("KORA MEDIÇÕES"))
        layout.itemAt(layout.count() - 1).widget().setStyleSheet("color: white; font-weight: bold; font-size: 14px;")
        
        senha_kora_widget = QWidget()
        senha_kora_layout = QHBoxLayout(senha_kora_widget)
        senha_kora_layout.setContentsMargins(0,0,0,0)

        e_kora = QLineEdit(creds["KORA_MED_PASS"])
        e_kora.setPlaceholderText("Senha Kora Medições")
        e_kora.setEchoMode(QLineEdit.EchoMode.Password)
        e_kora.setStyleSheet("background-color: #2d2d2d; color: white; padding: 8px; border-radius: 5px;")

        btn_olho_kora = QPushButton("👁️")
        btn_olho_kora.setFixedSize(50,40)
        btn_olho_kora.setCheckable(True)
        btn_olho_kora.setStyleSheet("background-color: #2d2d2d; border: 1px solid #555; border-radius: 5px;")
        btn_olho_kora.clicked.connect(lambda checked: self.toggle_senha(e_kora, btn_olho_kora, checked))

        senha_kora_layout.addWidget(e_kora)
        senha_kora_layout.addWidget(btn_olho_kora)
        layout.addWidget(senha_kora_widget)
        
        btn_save_kora = QPushButton("Salvar Kora Medições")
        btn_save_kora.setStyleSheet("background-color: #123b51; color: white; padding: 8px; border-radius: 5px;")
        btn_save_kora.clicked.connect(lambda: [salvar_credencial("KORA_MED_PASS", e_kora.text()), 
                                                QMessageBox.information(dialog, "OK", "Kora Medições Salvo!")])
        layout.addWidget(btn_save_kora)
        
        dialog.exec()

    def validar_e_rodar(self):
        ids_raw = self.txt_ids.toPlainText().strip()
        self.ids_processar = [l.strip() for l in ids_raw.split("\n") if l.strip() and not l.startswith("#")]
        if not self.ids_processar:
            QMessageBox.warning(self, "Erro", "Insira IDs válidos.")
            return
        creds = carregar_credenciais()
        if not all(creds.values()):
            QMessageBox.critical(self, "Erro", "Configure as credenciais primeiro!")
            return
        self.tabview.setCurrentIndex(1)
        self.statusBar().showMessage("Executando automação...")
        self.cont_sucesso = 0
        self.cont_erro = 0
        self.atualizar_metricas()

        # ═══ NOVO: Mostrar e configurar barra ═══
        #s elf.progress_bar.setVisible(True)
        # self.progress_bar.setMaximum(len(self.ids_processar))
        # self.progress_bar.setValue(0)
        # self.progress_label.setText("Iniciando...")
        # ════════════════════════════════════

        threading.Thread(target=self.executar_automacao, args=(creds,), daemon=True).start()

    def executar_automacao(self, creds):
        caminho_driver = os.path.join(os.getcwd(), "chromedriver.exe")
        servico = Service(caminho_driver)
        chrome_options = Options()
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--headless=new")
        chrome_options.add_argument("--window-size=1920,1080")
        
        driver = webdriver.Chrome(service=servico, options=chrome_options)
        wait = WebDriverWait(driver, 45)

        def focar_sap():
            driver.switch_to.default_content()
            try:
                iframes = driver.find_elements(By.TAG_NAME, "iframe")
                if iframes: driver.switch_to.frame(0)
            except: pass

        def clicar_com_retry(seletor, tipo_seletor, nome_campo, seletor_espera=None):
            for tentativa in range(3):
                try:
                    elemento = wait.until(EC.element_to_be_clickable((tipo_seletor, seletor)))
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento)
                    time.sleep(0.8)
                    elemento.click()
                    if seletor_espera:
                        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, seletor_espera)))
                    return True
                except:
                    pass
                    time.sleep(1)
            return False

        def enviar_ao_solicitante(id_v, tipo_erro, org_atual=None):
            try:
                driver.switch_to.window(v360_handle)
                driver.get(f"https://kora.virtual360.io/nf/acceptance_terms/{id_v}")
                time.sleep(2)

                # bom dia/tarde/noite dependendo do horário
                hora = datetime.now().hour
                if hora < 12: saudacao = "Bom dia!"
                elif hora < 18: saudacao = "Boa tarde!"
                else: saudacao = "Boa noite!"

                msg = ""

                # carregar mensagens do txt mensagensSolicitante.txt, na pasta do projeto, para permitir personalização da mensagem sem alterar o código
                mensagens_dict = {}
                if os.path.exists("mensagensSolicitante.txt"):
                    with open("mensagensSolicitante.txt", "r", encoding="utf-8") as f:
                        for linha in f:
                            if "::" in linha:
                                k, v = linha.strip().split("::", 1)
                                mensagens_dict[k] = v

                if tipo_erro == "inexistente_sap":
                    msg = mensagens_dict.get("inexistente_sap", f"{saudacao} A solicitação não foi encontrada na base SAP, favor liberar a solicitação caso não esteja liberada ou disponibilizar uma nova.")
                elif tipo_erro == "de_para_errado":
                    if org_atual in ["1418", "2001", "2901"]:
                        msg = mensagens_dict.get("de_para_errado_especifico", f"{saudacao} A solicitação não refletiu corretamente, favor disponibilizar uma nova.")
                    else:
                        msg = mensagens_dict.get("de_para_errado_geral", f"{saudacao} A solicitação MV disponibilizada não está refletindo no SAP, por gentileza poderia gerar uma nova? Favor verificar também o material usado, de acordo com a unidade da medição, na planilha: https://docs.google.com/spreadsheets/d/1ho46nlMd3eDM2Axce8XBx2RNsQ2h_r06/edit?gid=8752513#gid=8752513")
                elif tipo_erro == "cnpj_sem_cadastro":
                    msg = mensagens_dict.get("cnpj_sem_cadastro", f"{saudacao} Informo que o fornecedor não possui cadastro na base SAP, por gentileza solicitar o cadastro do mesmo *para a unidade da medição*. Segue link para realizar a solicitação: https://www.appsheet.com/start/44e01724-465a-42f5-9680-9884977096ad#view=Nova%20Solicita%C3%A7%C3%A3o")
                elif tipo_erro == "cnpj_sem_expansao":
                    msg = mensagens_dict.get("cnpj_sem_expansao", f"{saudacao} Informo que o fornecedor não está expandido na base SAP para a unidade, por gentileza solicitar a expansão do mesmo *para a unidade da medição*. Segue link para realizar a solicitação: https://www.appsheet.com/start/44e01724-465a-42f5-9680-9884977096ad#view=Nova%20Solicita%C3%A7%C3%A3o")

                # botão "ir para ações pendentes" do v360
                btn_pendentes = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.processable-pending-actions-btn")))
                time.sleep(0.5)
                btn_pendentes.click()
                time.sleep(0.5)

                # botão de enviar para o solicitante
                btn_env_sol = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Enviar para solicitante')]")))
                btn_env_sol.click()
                time.sleep(1)
                
                # confirmar pop-up de confirmação que abre, caso abra
                try:
                    alert = driver.switch_to.alert
                    alert.accept()
                except: pass
                
                time.sleep(1)

                # digita mensagem no campo de texto para enviar ao solicitante
                editor = wait.until(EC.presence_of_element_located((By.ID, "note_text_show")))
                driver.execute_script(f"arguments[0].innerHTML = '<div>{msg}</div>';", editor)
                time.sleep(1)

                # confirma envio da mensagem para o solicitante
                btn_confirmar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@type='submit' and contains(text(), 'Enviar para solicitante')]")))
                btn_confirmar.click()
                
                # confirmar pop-up de confirmação que abre, caso abra
                time.sleep(3)
                try:
                    driver.switch_to.alert.accept()
                except: pass
                
                self.atualizar_log(f"ID {id_v}: Enviado para o solicitante", "amarelo")
                
                # --- marcar como enviado ao solicitante no kora-medicoes.web.app ---
                driver.switch_to.window(kora_handle)
                campo_pesquisa = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[placeholder='Pesquisar...']")))
                campo_pesquisa.clear()
                campo_pesquisa.send_keys(id_v)
                time.sleep(0.5)
                
                btn_obs = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@title='Observação']")))
                btn_obs.click()
                time.sleep(1)
                
                btn_enviado_sol = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Enviado para o solicitante')]")))
                time.sleep(0.5)
                btn_enviado_sol.click()
                time.sleep(0.5)
            except Exception as e_envio:
                self.atualizar_log(f"Erro ao enviar para solicitante: {e_envio}", "vermelho")

        def processar_guarda_chuva(id_v360, kora_handle, v360_handle):
            # caso seja pedido guarda-chuva, precisa pegar o número do pedido e o tipo do pedido no Kora para preencher no V360, então aqui tem um processo específico para isso
            driver.switch_to.window(kora_handle)
            campo_pesquisa = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[placeholder='Pesquisar...']")))
            campo_pesquisa.clear()
            campo_pesquisa.send_keys(id_v360)
            time.sleep(1.5)
            
            # pegar o número do pedido
            try:
                num_pedido_el = driver.find_element(By.XPATH, "//button[contains(text(), '45') and contains(@style, 'rgb(19, 62, 81)')]")
                num_pedido = num_pedido_el.text.strip()
            except:
                num_pedido = driver.find_element(By.CSS_SELECTOR, "button[style*='background-color: rgb(19, 62, 81)']").text.strip()
            
            # pegar o tipo do pedido (fixo ou variavel)
            tipo_info = driver.find_element(By.CSS_SELECTOR, "span.bg-slate-100").text
            
            self.atualizar_log(f"Possui contrato GED e pedido pronto ({tipo_info}). Liberando medição.")
            time.sleep(0.5)
            # preencher no v360
            driver.switch_to.window(v360_handle)
            driver.get(f"https://kora.virtual360.io/nf/acceptance_terms/{id_v360}")
            
            # ver se está na etapa certa do setor de contratos...
            try:
                titulo_etapa = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "checkout-bar-item-title"))).text
            except:
                self.registrar_erro(id_v360, "Página do V360 não carregou a tempo")
                self.atualizar_log(f"Aviso: ID {id_v360} - Página não carregou.", "amarelo")
                return False
            
            btn_editar = wait.until(EC.element_to_be_clickable((By.ID, "nav-edit-tab")))
            time.sleep(1)
            btn_editar.click()
            time.sleep(1.5)
            
            # colocar o pedido na medição via js 
            script_final = """
            var pedido = arguments[0];
            var campos = document.querySelectorAll('#acceptance_term_purchase_order');
            var campoAtivo = null;
            for (var i = 0; i < campos.length; i++) {
                if (campos[i].offsetWidth > 0 || campos[i].offsetHeight > 0) {
                    campoAtivo = campos[i]; break;
                }
            }
            if (campoAtivo) {
                campoAtivo.focus(); campoAtivo.value = pedido;
                campoAtivo.dispatchEvent(new Event('input', { bubbles: true }));
                campoAtivo.dispatchEvent(new Event('change', { bubbles: true }));
            }
            """
            driver.execute_script(script_final, num_pedido)
            time.sleep(0.5); time.sleep(0.5); time.sleep(0.5)
            
            # selecionar o tipo de pedido correto no v360 (guarda-chuva variável ou fixo)
            span_select2 = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "[aria-labelledby='select2-acceptance_term_items_attributes_0_cf_tipo_de_pedido-container']")))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", span_select2)
            time.sleep(1)
            span_select2.click()
            
            opcao_texto = "Pedido Guarda-Chuva Variavel" if "Variavel" in tipo_info else "Pedido Guarda-Chuva Fixo"
            wait.until(EC.element_to_be_clickable((By.XPATH, f"//li[contains(text(), '{opcao_texto}')]"))).click()
            time.sleep(0.5)
            # salvar medição
            btn_salvar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@name='status_id' and contains(text(), 'Salvar')]")))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_salvar)
            time.sleep(0.5)
            btn_salvar.click()
            
            # botão "ir para ações pendentes" do v360
            btn_pendentes = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.processable-pending-actions-btn")))
            time.sleep(0.5)
            btn_pendentes.click()
            
            # botão de "tentar novamente" (libera medição)
            btn_tentar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Tentar Novamente__IServ')]")))
            time.sleep(0.5)
            btn_tentar.click()
            
            # verificar se foi para a alçada do solicitante (seguiu) ou se ficou na mesma etapa, com logica de atualizar a pagina algumas vezes para evitar falhas de atualização do status do v360, que é bem frequente.
            status_correto = False
            time.sleep(3)
            
            for i in range(5):
                try:
                    status_txt = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "checkout-bar-item-title"))).text.strip().lower()
                    if "analisar - ciclo de alçada solicitante" in status_txt:
                        status_correto = True
                        break
                except: pass
                time.sleep(1)
            
            if not status_correto:
                driver.refresh()
                time.sleep(3)
                for i in range(5):
                    try:
                        status_txt = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "checkout-bar-item-title"))).text.strip().lower()
                        if "analisar - ciclo de alçada solicitante" in status_txt:
                            status_correto = True
                            break
                    except: pass
                    time.sleep(1)
            
            if not status_correto:
                self.atualizar_log(f"[FALHA] Verificar pedido e medição do id #{id_v360}", "vermelho")
                self.registrar_erro(id_v360, "Status não atualizou")
                return False
            
            # marcar como feito no kora-medicoes.web.app
            driver.switch_to.window(kora_handle)
            campo_pesquisa = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[placeholder='Pesquisar...']")))
            campo_pesquisa.clear()
            campo_pesquisa.send_keys(id_v360)
            time.sleep(1)
            
            btn_feito = wait.until(EC.visibility_of_element_located((By.XPATH, "(//button[text()='FEITO'])[1]")))
            time.sleep(1)
            btn_feito.click()
            time.sleep(0.5)
            
            self.atualizar_log(f"✅ ID {id_v360} LIBERADO!")
            self.registrar_sucesso(id_v360, num_pedido)
            return True

        try:
            self.atualizar_log("Iniciando Logins...")
            
            # abrir site do kora-medicoes e logar
            driver.get("https://kora-medicoes.web.app/")
            driver.execute_script("document.body.style.zoom='90%'")
            
            senha_input = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[type='password']")))
            senha_input.click()
            senha_input.clear()
            time.sleep(0.5)
            senha_input.send_keys(creds["KORA_MED_PASS"])
            time.sleep(1.5)        
            wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Entrar no Painel')]"))).click()
            time.sleep(1)
            kora_handle = driver.current_window_handle
            
            # abrir site do v360 em outra aba e logar
            driver.execute_script("window.open('https://kora.virtual360.io/', '_blank');")
            time.sleep(1)
            for handle in driver.window_handles:
                if handle != kora_handle:
                    v360_handle = handle
                    break
            driver.switch_to.window(v360_handle)
            time.sleep(1) #THASYLA
            wait.until(EC.presence_of_element_located((By.ID, "user_login"))).send_keys(creds["V360_USER"])
            time.sleep(2) #THASYLA
            wait.until(EC.presence_of_element_located((By.ID, "user_password"))).send_keys(creds["V360_PASS"])
            time.sleep(3) #THASYLA
            
            clicar_com_retry("button.v-btn.submit-button", By.CSS_SELECTOR, "Login V360")
            time.sleep(5) #THASYLA

            for idx, id_v360 in enumerate(self.ids_processar, 1):
                # ═══ NOVO: Atualizar barra ═══
                # self.progress_bar.setValue(idx)
                # percentual = int((idx / len(self.ids_processar)) * 98)
                # self.progress_label.setText(f"Processando {idx} de {len(self.ids_processar)} ({percentual}%)")
                # QApplication.processEvents()
                # ════════════════════════════
                
                sap_handle = None
                erro_ja_registrado = False  # 999999: controle de erro duplicado
                try:
                    self.atualizar_log(f"Processando ID: {id_v360}...")
                    
                    # verificar se a medição tem pedido pronto no kora-medicoes.web.app, caso tenha, vai liberar no v360, caso não tenha, vai criar o pedido no SAP (avulso) e depois liberar no v360
                    driver.switch_to.window(kora_handle)
                    campo_pesquisa = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[placeholder='Pesquisar...']")))
                    campo_pesquisa.clear()
                    campo_pesquisa.send_keys(id_v360)
                    time.sleep(1)
                    
                    # tenta encontrar pedido no kora-medicoes.web.app (formato 45xxxxxx)
                    try:
                        num_pedido_el = driver.find_element(By.XPATH, "//button[contains(text(), '45') and contains(@style, 'rgb(19, 62, 81)')]")
                        num_pedido = num_pedido_el.text.strip()
                        is_guarda_chuva = num_pedido.startswith("45")  # verifica se começa com 45 para confirmar que é um pedido, e não outra informação que esteja no botão por acaso
                    except:
                        try:
                            num_pedido = driver.find_element(By.CSS_SELECTOR, "button[style*='background-color: rgb(19, 62, 81)']").text.strip()
                            is_guarda_chuva = num_pedido.startswith("45")  # verifica se começa com 45 para confirmar que é um pedido, e não outra informação que esteja no botão por acaso
                        except:
                            is_guarda_chuva = False
                    if is_guarda_chuva:
                        self.atualizar_log(f"ID {id_v360} → GUARDA-CHUVA (Pedido: {num_pedido})")
                        processar_guarda_chuva(id_v360, kora_handle, v360_handle)
                        continue
                    
                    # ---- LÓGICA NORMAL (SAP + V360) ----
                    self.atualizar_log(f"Não possui contrato no GED. Fazendo avulso.")
                    
                    driver.execute_script(f"window.open('https://prd.sap.korasaude.app.br/sap/bc/ui2/flp?sap-client=400&sap-language=PT#PurchaseOrder-create?sap-ui-tech-hint=GUI&uitype=advanced', '_blank');")
                    time.sleep(1)
                    for handle in driver.window_handles:
                        if handle != v360_handle and handle != kora_handle: sap_handle = handle
                    driver.switch_to.window(sap_handle)
                    
                    try:
                        user_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "USERNAME_FIELD-inner")))
                        user_field.send_keys(creds["SAP_USER"])
                        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "PASSWORD_FIELD-inner"))).send_keys(creds["SAP_PASS"])
                        clicar_com_retry("LOGIN_LINK", By.ID, "Login SAP")
                    except: pass
                    
                    driver.switch_to.window(v360_handle)
                    driver.get(f"https://kora.virtual360.io/nf/acceptance_terms/{id_v360}")
                    
                    # ver se está na etapa certa do setor de contratos, caso esteja, faça, caso contrário, registre erro e vá para o próximo id
                    try:
                        titulo_etapa = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "checkout-bar-item-title"))).text
                    except:
                        self.registrar_erro(id_v360, "Página do V360 não carregou a tempo")
                        self.atualizar_log(f"Aviso: ID {id_v360} - Página não carregou.", "amarelo")
                        continue
                    
                    if "Analisar - Divergência Entre Pedido de Compras e Medição" not in titulo_etapa:
                        self.registrar_erro(id_v360, "ID não está na etapa de criar pedido do zero.")
                        self.atualizar_log(f"Aviso: ID {id_v360} não está na etapa de criar pedido do zero no v360.", "amarelo")
                        continue

                    v_solicitacao = wait.until(EC.presence_of_element_located((By.ID, "acceptance_term_purchase_order"))).get_attribute("value")
                                        # Remove .0 do final da solicitação (ex: 123456.0 → 123456)
                    if v_solicitacao.endswith(".0"):
                        v_solicitacao = v_solicitacao[:-2]
                    v_cnpj = driver.find_element(By.ID, "acceptance_term_supplier_identification_number").get_attribute("value")
                    v_valor = driver.find_element(By.ID, "acceptance_term_total_value").get_attribute("value")
                    v_org_cod = driver.find_element(By.ID, "acceptance_term_erp_purchasing_organization").get_attribute("value")[:4]
                    
                    # verifica se o campo de solicitação tem realmente uma solicitação e não um pedido aleatório preenchido pelo solicitante, caso tenha um numero de 7 ou mais caracteres, classifica como pedido e avisa no painel que tem um pedido no lugar da solicitação, para que o usuário verifique a situação do pedido.
                    if len(v_solicitacao) >= 7:
                        self.atualizar_log(f"Aviso: {id_v360} possui pedido no lugar da solicitação. Verifique a situação do pedido.", "amarelo")
                        self.registrar_erro(id_v360, f"Possui pedido ({v_solicitacao}) no lugar da solicitação. Verifique manualmente.")
                        continue
                        
                    # se a medição for da Kora (1400), avisa que é da Kora pois o processo para fazer pedido da Kora é diferente e no momento deve ser feito manual.
                    if v_org_cod == "1400":
                        self.atualizar_log(f"Aviso: {id_v360} é da Kora, fazer manual.", "amarelo")
                        continue
                        
                    v_iva = "ZZ" 
                    csv_path = resource_path(f"{v_org_cod}.csv")
                    if os.path.exists(csv_path):
                        with open(csv_path, mode='r', encoding='utf-8') as f_csv:
                            reader = csv.reader(f_csv)
                            for row in reader:
                                if row and row[0].strip() == v_cnpj.strip():
                                    v_iva = row[1].strip()
                                    break
                    
                    driver.switch_to.window(sap_handle)
                    time.sleep(2)
                    focar_sap()
                    
                    # Aguarda elemento do SAP carregar (por ID)
                    try:
                        WebDriverWait(driver, 30).until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, "span[id*='78-text']"))
                        )
                    except:
                        pass
                    time.sleep(2)
                    
                    # ativar sintese sap com retry
                    if not clicar_com_retry("div[lsdata*='Ativar síntese de documentos']", By.CSS_SELECTOR, "Ativar síntese"):
                        self.atualizar_log("Botão Ativar síntese não funcionou", "amarelo")
                        continue
                    
                    # Verifica se o dropdown apareceu
                    try:
                        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "[title*='Variante de seleção']")))
                    except:
                        self.atualizar_log("Dropdown da síntese não apareceu", "amarelo")
                        continue

                    # abrir dropdown
                    try:
                        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "[title*='Variante de seleção']")))
                        time.sleep(0.5)
                    except:
                        pass
                    if not clicar_com_retry("[title*='Variante de seleção']", By.CSS_SELECTOR, "Variante de seleção"): continue
                    time.sleep(1)
                    
                    # clicar em requisições de compra
                    try:
                        wait.until(EC.visibility_of_element_located((By.XPATH, "//tr[contains(@aria-label, 'Requisições de compra')]")))
                        time.sleep(0.5)
                    except:
                        pass
                    if not clicar_com_retry("//tr[contains(@aria-label, 'Requisições de compra')]", By.XPATH, "Requisições de compra"): continue
                    
                    # selecionar campo de solicitação e inserir número
                    campo_acomp = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[title*='acompanhamento']")))
                    if not clicar_com_retry("input[title*='acompanhamento']", By.CSS_SELECTOR, "Campo acompanhamento"): continue
                    
                    campo_acomp.clear()
                    campo_acomp.send_keys(v_solicitacao)
                    campo_acomp.send_keys(Keys.F8)
                    time.sleep(1.5)
                    
                    # tratamento de erro para ver se a solicitação existe ou não na base do SAP, caso exista, segue o processo normalmente, caso não exista, manda automaticamente para o solicitante no v360 
                    try:
                        msg_inexistente = driver.find_element(By.ID, "M1:46:::0:5-text").text
                        if "Não existem dados para os critérios de seleção" in msg_inexistente:
                            self.registrar_erro(id_v360, "ID não foi encontrado na base SAP")
                            self.atualizar_log(f"Aviso: ID {id_v360} não foi encontrado na base SAP", "amarelo")
                            enviar_ao_solicitante(id_v360, "inexistente_sap")
                            continue
                    except: pass
                    
                    # selecionar a primeira caixinha para espelhar
                    wait.until(EC.presence_of_element_located((By.CLASS_NAME, "urST5SCMetricInner")))
                    driver.execute_script("""
                        const metric = document.querySelector('.urST5SCMetricInner');
                        if (metric) {
                            const cell = metric.closest('td');
                            ['mousedown', 'mouseup', 'click'].forEach(evtType => {
                                cell.dispatchEvent(new MouseEvent(evtType, { bubbles: true }));
                            });
                        }
                    """)
                    time.sleep(2)
                    
                    # clicar para espelhar a requisição
                    if not clicar_com_retry("[title='Transferir']", By.CSS_SELECTOR, "Botão Transferir"): continue
                    time.sleep(3)
                    
                     # tratamento de erro para verificar se a solicitação espelhou corretamente (hospital correto e como serviço), caso tenha espelhado corretamente, segue o processo normalmente, caso contrário, manda automaticamente para o solicitante no v360.
                    # parte especifica de verificar se o hospital está correto, de acordo com a unidade da medição
                    try:
                        wait.until(EC.presence_of_element_located((By.ID, "M0:46:1:3:2:1:1[1,16]_c")))
                        time.sleep(0.5)
                        json_data = driver.execute_script('return document.getElementById("M0:46:1:3:2:1:1[1,16]_c").getAttribute("lsdata");')
                        nome_campo = json.loads(json_data).get("21", {}).get("value", "")
                        
                        achou_de_para = False
                        valor_esperado = ""
                        if os.path.exists(resource_path("deParaUnidades.csv")):
                            with open(resource_path("deParaUnidades.csv"), mode='r', encoding='utf-8') as f:
                                for linha in f:
                                    partes = linha.strip().split(',')
                                    if len(partes) >= 2 and partes[0].strip() == v_org_cod:
                                        valor_esperado = partes[1].strip()
                                        achou_de_para = True
                                        break
                                        
                        if achou_de_para and valor_esperado.upper() == nome_campo.upper():
                            self.atualizar_log(f"Unidade correta? Sim ({nome_campo})")
                        else:
                            self.atualizar_log(f"Unidade correta? Não ({nome_campo})")
                            self.registrar_erro(id_v360, f"ID {id_v360} com solicitação que não refletiu na unidade correta (Esperado: {valor_esperado}, Lido: {nome_campo})")
                            enviar_ao_solicitante(id_v360, "de_para_errado", v_org_cod)
                            continue
                    except Exception as e:
                        self.atualizar_log(f"Erro na leitura da Unidade: {e}", "vermelho")
                        continue

                    # parte especifica de verificar se é serviço e não material
                    try:
                        wait.until(EC.presence_of_element_located((By.ID, "M0:46:1:3:2:1:1[1,15]_c")))
                        time.sleep(0.5)
                        json_data_cat = driver.execute_script('return document.getElementById("M0:46:1:3:2:1:1[1,15]_c").getAttribute("lsdata");')
                        val_cat = json.loads(json_data_cat).get("21", {}).get("value", "")
                        
                        prefixos_validos = ("REPASSE", "SERV.", "PLANO DE SAUDE", "DESPESAS COM SOFTW")
                        if val_cat.startswith(prefixos_validos):
                            self.atualizar_log(f"Serviço correto? Sim ({val_cat})")
                        else:
                            self.atualizar_log(f"Serviço correto? Não ({val_cat})")
                            self.registrar_erro(id_v360, f"ID {id_v360} com solicitação que não refletiu como serviço: {val_cat}")
                            enviar_ao_solicitante(id_v360, "de_para_errado", v_org_cod)
                            continue
                    except Exception as e:
                        self.atualizar_log(f"Erro na leitura do Serviço: {e}", "vermelho")
                        continue
                    
                    # clicar no campo do fornecedor
                    if not clicar_com_retry("input[title*='Fornecedor/centro fornecedor']", By.CSS_SELECTOR, "Campo fornecedor"): continue
                    
                    # clicar no botão para abrir a telinha de digitar o cnpj
                    try:
                        wait.until(EC.visibility_of_element_located((By.ID, "ls-inputfieldhelpbutton")))
                        time.sleep(0.5)
                    except:
                        pass
                    if not clicar_com_retry("ls-inputfieldhelpbutton", By.ID, "Botão Help Fornecedor"): continue
                    
                    campo_token = wait.until(EC.element_to_be_clickable((By.NAME, "TokenizerComboBox")))
                    campo_token.send_keys(v_cnpj)
                    campo_token.send_keys(Keys.ENTER)
                    
                    # clicar no botão inicio para pesquisar o fornecedor
                    try:
                        wait.until(EC.visibility_of_element_located((By.ID, "btnGO2")))
                        time.sleep(0.5)
                    except:
                        pass
                    if not clicar_com_retry("btnGO2", By.ID, "Botão GO"): continue
                    time.sleep(1.5)
                    
                    # tratamento de erro para verificar se o fornecedor está cadastrado e expandido para a unidade da medição. caso esteja, segue normalmente, caso não esteja, envia para o solicitante no v360.
                    try:
                        msg_erro_fornecedor = driver.find_element(By.ID, "wnd[0]/sbar_msg-txt")
                        texto_erro = msg_erro_fornecedor.get_attribute("title") or msg_erro_fornecedor.text
                        if "Nenhum valor para esta seleção" in texto_erro:
                            self.registrar_erro(id_v360, f"ID {id_v360} está com CNPJ sem cadastro na base SAP")
                            enviar_ao_solicitante(id_v360, "cnpj_sem_cadastro")
                            continue
                        if "não foi criado para organização de compras" in texto_erro or "não foi criado para a organização de compras" in texto_erro:
                            self.registrar_erro(id_v360, f"ID {id_v360} necessita de cadastro do CNPJ para a unidade da medição")
                            enviar_ao_solicitante(id_v360, "cnpj_sem_expansao")
                            continue
                    except:
                        texto_body = driver.find_element(By.TAG_NAME, "body").text
                        if "Nenhum valor para esta seleção" in texto_body:
                            self.registrar_erro(id_v360, f"ID {id_v360} está com CNPJ sem cadastro na base SAP")
                            enviar_ao_solicitante(id_v360, "cnpj_sem_cadastro")
                            continue
                        if "não foi criado para a organização de compras" in texto_body:
                            self.registrar_erro(id_v360, f"ID {id_v360} necessita de cadastro do CNPJ para a unidade da medição")
                            enviar_ao_solicitante(id_v360, "cnpj_sem_expansao")
                            continue
                    
                    # clicar no botão ir para confirmar a seleção do fornecedor
                    try:
                        wait.until(EC.visibility_of_element_located((By.ID, "NSH2_copy")))
                        time.sleep(0.5)
                    except:
                        pass
                    if not clicar_com_retry("NSH2_copy", By.ID, "Botão Copy Fornecedor"): continue
                    time.sleep(2)
                    
                    # ultimo enter para confirmar a seleção do fornecedor depois da tela ser fechada
                    try:
                        wait.until(EC.invisibility_of_element_located((By.ID, "NSH2_copy")))
                        time.sleep(0.5)
                    except:
                        time.sleep(1)
                    driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ENTER)
                    time.sleep(2)

                    # tratamento de erro adicional para se certificar de que o fornecedor está expandido corretamente para a unidade da medição
                    fornecedor_com_erro = driver.execute_script("""
                        const element = document.querySelector('.lsMessageBar__text');
                        if (element) {
                            const mensagemOriginal = element.textContent || element.innerText;
                            const padrao = /O fornecedor \\d+ não foi criado para organização de compras \\d+/;
                            return padrao.test(mensagemOriginal);
                        }
                        return false;
                    """)
                    if fornecedor_com_erro:
                        self.registrar_erro(id_v360, f"ID {id_v360} necessita de cadastro do CNPJ para a unidade da medição")
                        enviar_ao_solicitante(id_v360, "cnpj_sem_expansao")
                        continue

                    # caso não esteja aberta, abrir a primeira aba dentro do SAP onde tem remessa/fatura para colocar a condição de pagamento, usando CTRL+F2 
                    campo_moeda_element = driver.find_elements(By.CSS_SELECTOR, "input[title*='Código da moeda']")
                    if not campo_moeda_element:
                        ActionChains(driver).key_down(Keys.CONTROL).send_keys(Keys.F2).key_up(Keys.CONTROL).perform()
                        time.sleep(2)
                        
                        # clicar em remessa/fatura com retry até aparecer o ZTERM
                        for tentativa_remessa in range(3):
                            clicar_com_retry("//span[contains(text(), 'Remessa/fatura')]", By.XPATH, "Aba Remessa/fatura")
                            time.sleep(1)
                            # Verifica se ZTERM apareceu
                            try:
                                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[lsdata*='MEPO1226-ZTERM']")), timeout=5)
                                break
                            except:
                                if tentativa_remessa == 2:
                                    self.atualizar_log("Aba Remessa/fatura não abriu corretamente", "amarelo")
                                    continue
                    
                    # aguardar o campo de moeda e verificar se está preenchido BRL
                    if campo_moeda_element:
                        campo_moeda = campo_moeda_element[0]
                    else:
                        campo_moeda = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[title*='Código da moeda']")))
                    time.sleep(0.5)
                    moeda_valor = campo_moeda.get_attribute("value")
                    if not moeda_valor or moeda_valor.strip() == "":
                        campo_moeda.clear()
                        campo_moeda.send_keys("BRL")
                        time.sleep(0.5)
                    
                    # colocar z030 no campo condição de pagamento (para pedidos avulsos)
                    z_field = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[lsdata*='MEPO1226-ZTERM']")))
                    if not clicar_com_retry("input[lsdata*='MEPO1226-ZTERM']", By.CSS_SELECTOR, "Campo ZTERM"): continue
                    z_field.send_keys(Keys.CONTROL + "a", Keys.BACKSPACE)
                    z_field.send_keys("Z030", Keys.ENTER)
                    time.sleep(1.5)
                    
                    # abrir a "ultima" aba dentro do SAP para colocar a o valor do pedido, quantidade 1 e o codigo IVA, usando CTRL+4
                    ActionChains(driver).key_down(Keys.CONTROL).send_keys("4").key_up(Keys.CONTROL).perform()
                    time.sleep(2)
                    
                    # clicar no botão de condições para abrir a tela de condições de pagamento e inserir o valor do pedido
                    if not clicar_com_retry("//span[contains(text(), 'Condições')]", By.XPATH, "Aba Condições"): continue
                    time.sleep(1)
                    
                    # colocar o valor
                    try:
                        wait.until(EC.visibility_of_element_located((By.ID, "M0:46:1:4:2:1:2:1:2B264:1:3[1,4]_c")))
                        time.sleep(0.5)
                    except:
                        pass
                    if not clicar_com_retry("M0:46:1:4:2:1:2:1:2B264:1:3[1,4]_c", By.ID, "Valor Condições", "M0:46:1:4:2:1:2:1:2B264:1:3[1,4]_c"): continue
                    driver.execute_script(f'document.getElementById("M0:46:1:4:2:1:2:1:2B264:1:3[1,4]_c").focus(); document.execCommand("insertText", false, "{v_valor}");')
                    time.sleep(1.5)
                    
                    # clicar no botão de quantidade para inserir a quantidade 1
                    try:
                        wait.until(EC.visibility_of_element_located((By.ID, "M0:46:1:4:2:1:2:1::0:4-text")))
                        time.sleep(0.5)
                    except:
                        pass
                    if not clicar_com_retry("M0:46:1:4:2:1:2:1::0:4-text", By.ID, "Quantidade", "M0:46:1:4:2:1:2:1:2B260:1::0:16"): continue
                    time.sleep(1.5)
                    driver.execute_script("document.getElementById('M0:46:1:4:2:1:2:1:2B260:1::0:16').value = '1';")
                    time.sleep(1)
                    
                    # clicar no botão de fatura para inserir o IVA
                    try:
                        wait.until(EC.visibility_of_element_located((By.ID, "M0:46:1:4:2:1:2:1::0:7-text")))
                        time.sleep(0.5)
                    except:
                        pass
                    if not clicar_com_retry("M0:46:1:4:2:1:2:1::0:7-text", By.ID, "Fatura/IVA", "M0:46:1:4:2:1:2:1:2B263::0:68"): continue
                    time.sleep(1.5)
                    
                    # colocar o código do iva com base nos csv salvos localmente na mesma pasta do robô
                    try:
                        wait.until(EC.visibility_of_element_located((By.ID, "M0:46:1:4:2:1:2:1:2B263::0:68")))
                        time.sleep(0.5)
                    except:
                        pass
                    iva_field = wait.until(EC.element_to_be_clickable((By.ID, "M0:46:1:4:2:1:2:1:2B263::0:68")))
                    if not clicar_com_retry("M0:46:1:4:2:1:2:1:2B263::0:68", By.ID, "Campo IVA"): continue
                    iva_field.send_keys(Keys.CONTROL + "a", Keys.BACKSPACE)
                    time.sleep(0.5) 
                    driver.execute_script(f'arguments[0].focus(); document.execCommand("insertText", false, "{v_iva}");', iva_field)
                    time.sleep(0.5)
                    iva_field.send_keys(Keys.ENTER)
                    time.sleep(3)
                    
                    # colocar as datas de inicio e fim no pedido (dia da criação do pedido e 30 dias depois)
                    hoje = datetime.now().strftime("%d.%m.%Y")
                    prox = (datetime.now() + timedelta(days=30)).strftime("%d.%m.%Y")
                    
                    # Espera o campo de data início aparecer
                    try:
                        wait.until(EC.presence_of_element_located((By.ID, "M0:46:1:3:2:1:1[1,42]_c")))
                        time.sleep(0.5)
                    except:
                        pass
                    
                    # Preenche data início
                    driver.execute_script(f"""
                        let d1 = document.getElementById("M0:46:1:3:2:1:1[1,42]_c");
                        if(d1) {{ d1.focus(); d1.value = ''; document.execCommand('insertText', false, '{hoje}'); d1.dispatchEvent(new Event('change', {{bubbles:true}})); }}
                    """)
                    time.sleep(1)
                    
                    # Espera o campo de data fim aparecer
                    try:
                        wait.until(EC.presence_of_element_located((By.ID, "M0:46:1:3:2:1:1[1,43]_c")))
                        time.sleep(0.5)
                    except:
                        pass
                    
                    # Preenche data fim
                    driver.execute_script(f"""
                        let d2 = document.getElementById("M0:46:1:3:2:1:1[1,43]_c");
                        if(d2) {{ d2.focus(); d2.value = ''; document.execCommand('insertText', false, '{prox}'); d2.dispatchEvent(new Event('change', {{bubbles:true}})); }}
                    """)
                    time.sleep(2)
                    
                    # salvar pedido com CTRL+S
                    ActionChains(driver).key_down(Keys.CONTROL).send_keys("s").key_up(Keys.CONTROL).perform()
                    time.sleep(1)
                    focar_sap()
                    
                    # Aguarda o número do pedido OU erro de gravação (10 tentativas de 1s)
                    pedido_gerado = None
                    for tentativa in range(10):
                        # Verifica popup de erro PRIMEIRO
                        try:
                            erro_gravacao = driver.find_element(By.XPATH, "//span[contains(text(), 'Gravar documento incorreto')]")
                            if erro_gravacao.is_displayed():
                                self.atualizar_log(f"ID {id_v360} com erro ao criar pedido. Verificar manual", "vermelho")
                                self.registrar_erro(id_v360, "Erro ao gravar documento no SAP")
                                erro_ja_registrado = True  # 999999: evita duplicar
                                break
                        except:
                            pass
                        
                        # Verifica número do pedido
                        try:
                            status_el = driver.find_element(By.CSS_SELECTOR, "[id$='sbar_msg-txt']")
                            msg_status = status_el.get_attribute("title") or status_el.text
                            match = re.search(r"45\d+", msg_status)
                            if match:
                                pedido_gerado = match.group()
                                break
                        except:
                            pass
                        
                        time.sleep(1)
                    
                    # Se encontrou pedido, continua
                    if pedido_gerado:
                        self.registrar_sucesso(id_v360, pedido_gerado)
                        self.atualizar_log(f"Sucesso: Pedido criado.", "verde")

                        self.atualizar_log("DEBUG: Voltando para V360...")

                        # voltar para o v360 para colocar o numero do pedido, categoria avulso e liberar a medição
                        driver.switch_to.window(v360_handle)
                        
                        btn_editar = wait.until(EC.element_to_be_clickable((By.ID, "nav-edit-tab")))
                        time.sleep(0.5)
                        btn_editar.click()
                        time.sleep(0.5) 

                        # preencher o pedido
                        script_final = """
                        var pedido = arguments[0];
                        var campos = document.querySelectorAll('#acceptance_term_purchase_order');
                        var campoAtivo = null;
                        for (var i = 0; i < campos.length; i++) {
                            if (campos[i].offsetWidth > 0 || campos[i].offsetHeight > 0) {
                                campoAtivo = campos[i]; break;
                            }
                        }
                        if (campoAtivo) {
                            campoAtivo.focus(); campoAtivo.value = pedido;
                            campoAtivo.dispatchEvent(new Event('input', { bubbles: true }));
                            campoAtivo.dispatchEvent(new Event('change', { bubbles: true }));
                        }
                        """
                        driver.execute_script(script_final, pedido_gerado)
                        time.sleep(0.5); time.sleep(0.5); time.sleep(0.5)

                        # selecionar o tipo de pedido como avulso
                        try:
                            span_select2 = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "[aria-labelledby='select2-acceptance_term_items_attributes_0_cf_tipo_de_pedido-container']")))
                            time.sleep(0.5)
                        except:
                            pass
                        span_select2 = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "[aria-labelledby='select2-acceptance_term_items_attributes_0_cf_tipo_de_pedido-container']")))
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", span_select2)
                        time.sleep(0.5)
                        span_select2.click()
                        time.sleep(0.5)
                        
                        opcao_texto = "Pedido Avulso"
                        wait.until(EC.element_to_be_clickable((By.XPATH, f"//li[contains(text(), '{opcao_texto}')]"))).click()

                        # salvar medição
                        btn_salvar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@name='status_id' and contains(text(), 'Salvar')]")))
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_salvar)
                        time.sleep(0.5)
                        btn_salvar.click()

                        # botão "ir para ações pendentes" do v360
                        btn_pendentes = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.processable-pending-actions-btn")))
                        time.sleep(0.5)
                        btn_pendentes.click()

                        # botão de "tentar novamente" (libera medição)
                        btn_tentar = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Tentar Novamente__IServ')]")))
                        time.sleep(0.5)
                        btn_tentar.click()

                        # verificar se foi para a alçada do solicitante (seguiu) ou se ficou na mesma etapa, com logica de atualizar a pagina algumas vezes para evitar falhas de atualização do status do v360, que é bem frequente.
                        status_correto = False
                        time.sleep(3) 

                        for i in range(5):
                            try:
                                status_txt = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "checkout-bar-item-title"))).text.strip().lower()
                                if "analisar - ciclo de alçada solicitante" in status_txt:
                                    status_correto = True
                                    break
                            except: pass
                            time.sleep(1)
                        
                        if not status_correto:
                            driver.refresh()
                            time.sleep(3)
                            for i in range(5):
                                try:
                                    status_txt = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "checkout-bar-item-title"))).text.strip().lower()
                                    if "analisar - ciclo de alçada solicitante" in status_txt:
                                        status_correto = True
                                        break
                                except: pass
                                time.sleep(1)

                        if not status_correto:
                            self.atualizar_log(f"[FALHA] Verificar pedido e medição do id #{id_v360}", "vermelho")
                            self.registrar_erro(id_v360, "Status não atualizou após tentar novamente")

                        # --- marcar AV e FEITO no kora-medicoes.web.app ---
                        driver.switch_to.window(kora_handle)
                        campo_pesquisa = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[placeholder='Pesquisar...']")))
                        campo_pesquisa.clear()
                        campo_pesquisa.send_keys(id_v360)
                        time.sleep(0.5)
                        
                        btn_av = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'AV')]")))
                        btn_av.click()
                        time.sleep(0.5)
                        
                        btn_feito = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'FEITO')]")))
                        btn_feito.click()
                        self.atualizar_log(f"✅ ID {id_v360} LIBERADO!")

                    else:
                        if not erro_ja_registrado:  # 999999: so registra se nao foi antes
                            self.atualizar_log(f"Erro não identificado: {id_v360}: Verificar manual", "vermelho")
                            self.registrar_erro(id_v360, "ALERTA: Pedido não localizado após CTRL+S")

                except Exception as e:
                    import traceback
                    self.atualizar_log(f"Falha ID {id_v360}: {str(e)[:200]}", "vermelho")
                    self.registrar_erro(id_v360, str(e)[:200])
                finally:
                    if sap_handle:
                        try:
                            driver.switch_to.window(sap_handle)
                            driver.close()
                        except: pass
                    driver.switch_to.window(v360_handle)
                    
            self.atualizar_log("🎉 PROCESSO FINALIZADO!")
            self.statusBar().showMessage("✅ Concluído!")
        except Exception as e:
            self.statusBar().showMessage("Erro!")
            erro_msg = f"ERRO CRÍTICO: {str(e)}"
            self.atualizar_log(erro_msg, "vermelho")
            
            # Salvar em arquivo pra não perder
            with open("erro_critico.txt", "w", encoding="utf-8") as f:
                import traceback
                f.write(traceback.format_exc())
            
            QMessageBox.critical(self, "Erro Crítico", str(e))
        finally:
            # self.progress_bar.setValue(len(self.ids_processar))  # 100% real
            # self.progress_label.setText("✅ Concluído!")
            # time.sleep(0.5)  # Pequena pausa pra mostrar 100%
            # self.progress_bar.setVisible(False)
            # self.progress_label.setText("Pronto para iniciar")
            driver.quit()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    splash = SplashScreen()
    splash.show()
    splash.set_progress(5)
    
    if not verificar_acesso_remoto():
        splash.close()
        sys.exit()
    
    splash.set_progress(30)
    
    VERSAO_ATUAL = "3.0.0"
    versao_minima = verificar_versao()
    if versao_minima > VERSAO_ATUAL:
        splash.close()
        QMessageBox.warning(
            None,
            "Atualização Necessária",
            f"Nova versão {versao_minima} disponível!\n"
            f"Sua versão: {VERSAO_ATUAL}\n\n"
            "Baixe a atualização no link enviado."
        )
        sys.exit()
    
    splash.set_progress(55)
    
    apply_stylesheet(app, theme='dark_blue.xml')
    
    splash.set_progress(80)
    
    window = App()
    
    splash.set_progress(100)
    time.sleep(0.4)
    
    splash.close()
    window.showMaximized()
    sys.exit(app.exec())
