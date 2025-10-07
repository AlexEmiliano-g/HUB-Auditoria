# ==============================================================================
# HUB.PY - VERSÃO FINAL COM CORREÇÃO DE REFERÊNCIA DE JANELA
# ==============================================================================

import sys
import os
import importlib

# Imports movidos para a função de carregamento para não bloquear o início
# import pandas
# import openpyxl

from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QListWidget, QPushButton, QTextEdit, QLabel, QListWidgetItem
)
from PyQt6.QtGui import QFont, QPixmap, QIcon
from PyQt6.QtCore import Qt, QPropertyAnimation, QTimer

# (O resto do seu código - resource_path, definições globais, stylesheet, Splash, Hub - permanece o mesmo)
# ...
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

SCRIPTS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "scripts")
LOGO_HUB_PATH = resource_path("logo.png")
LOGO_SPLASH_PATH = resource_path("logo_aber.png")

STYLESHEET = """ 
QWidget#HubWindow { background-color: #F8F9FA; }
QWidget#LeftPanel { background-color: #FFFFFF; border-right: 1px solid #E0E0E0; border-top-left-radius: 10px; border-bottom-left-radius: 10px; }
QListWidget { border: none; font-size: 14px; padding: 5px; outline: 0; }
QListWidget::item { padding: 12px 15px; border-radius: 5px; }
QListWidget::item:hover { background-color: #F1F3F4; }
QListWidget::item:selected { background-color: #E8F0FE; color: #1967D2; font-weight: bold; }
QPushButton { background-color: #4285F4; color: white; font-size: 14px; font-weight: bold; border: none; padding: 12px; border-radius: 5px; }
QPushButton:hover { background-color: #1A73E8; }
QPushButton:pressed { background-color: #1765CC; }
QTextEdit { background-color: #FFFFFF; border: 1px solid #DADCE0; border-radius: 5px; font-family: 'Consolas', 'Courier New', monospace; color: #3C4043; padding: 8px; }
QLabel#TitleLabel { font-size: 16px; font-weight: bold; color: #202124; padding-bottom: 5px; }
"""

class Splash(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Iniciando...")
        self.resize(600, 400)
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.SplashScreen)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        self.container_label = QLabel()
        self.container_label.setStyleSheet("background-color: #333333; border-radius: 15px;")
        layout.addWidget(self.container_label)
        container_layout = QVBoxLayout(self.container_label)
        logo_label = QLabel("Logo de abertura não encontrado")
        if os.path.exists(LOGO_SPLASH_PATH):
            pixmap = QPixmap(LOGO_SPLASH_PATH)
            scaled_pixmap = pixmap.scaled(400, 300, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
            logo_label.setPixmap(scaled_pixmap)
        logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        container_layout.addWidget(logo_label)
        self.status_label = QLabel("Inicializando...")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_label.setStyleSheet("color: #FFFFFF; font-size: 12px; padding: 10px;")
        container_layout.addWidget(self.status_label)
        self.setWindowOpacity(0.0)
        self.fade_anim = QPropertyAnimation(self, b"windowOpacity")
        self.fade_anim.setDuration(1000)
        self.fade_anim.setStartValue(0.0)
        self.fade_anim.setEndValue(1.0)
        self.fade_anim.start()
    def set_status(self, text):
        self.status_label.setText(text)
        QApplication.processEvents()

class Hub(QWidget):
    def __init__(self):
        super().__init__()
        self.open_windows = []
        self.setWindowTitle("DM HUB - Central de Aplicativos")
        if os.path.exists(LOGO_HUB_PATH):
            self.setWindowIcon(QIcon(LOGO_HUB_PATH))
        self.resize(1100, 700)
        self.setObjectName("HubWindow")
        self.setStyleSheet(STYLESHEET)
        main_layout = QHBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        left_panel = QWidget()
        left_panel.setObjectName("LeftPanel")
        left_panel.setFixedWidth(240)
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(10, 10, 10, 10)
        left_layout.setSpacing(10)
        logo_label = QLabel()
        if os.path.exists(LOGO_HUB_PATH):
            pixmap = QPixmap(LOGO_HUB_PATH)
            scaled_pixmap = pixmap.scaled(220, 100, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
            logo_label.setPixmap(scaled_pixmap)
        logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        left_layout.addWidget(logo_label)
        self.sectors_list = QListWidget()
        for setor in ["Contabilidade", "Fiscal"]:
            self.sectors_list.addItem(QListWidgetItem(setor))
        self.sectors_list.itemClicked.connect(self.load_scripts)
        left_layout.addWidget(self.sectors_list)
        main_layout.addWidget(left_panel)
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(25, 20, 25, 20)
        right_layout.setSpacing(15)
        scripts_title = QLabel("Programas Disponíveis")
        scripts_title.setObjectName("TitleLabel")
        right_layout.addWidget(scripts_title)
        self.scripts_list = QListWidget()
        right_layout.addWidget(self.scripts_list, 3)
        self.run_button = QPushButton("Executar Programa Selecionado")
        self.run_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.run_button.clicked.connect(self.run_script)
        right_layout.addWidget(self.run_button)
        log_title = QLabel("Log de Execução")
        log_title.setObjectName("TitleLabel")
        right_layout.addWidget(log_title)
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        right_layout.addWidget(self.log, 1)
        main_layout.addWidget(right_panel, 1)
    def load_scripts(self, item):
        setor = item.text()
        setor_path = resource_path(os.path.join("scripts", setor))
        self.scripts_list.clear()
        if not os.path.exists(setor_path):
            self.log.append(f"AVISO: A pasta para o setor '{setor}' não foi encontrada.")
            return
        for script_filename in sorted(os.listdir(setor_path)):
            if script_filename.endswith("_app.py"):
                display_name = script_filename.replace("_app.py", "").replace("_", " ").capitalize()
                list_item = QListWidgetItem(display_name)
                list_item.setData(Qt.ItemDataRole.UserRole, script_filename)
                self.scripts_list.addItem(list_item)
    def run_script(self):
        setor_item = self.sectors_list.currentItem()
        script_item = self.scripts_list.currentItem()
        if not setor_item or not script_item:
            self.log.append("Selecione um setor e um programa.")
            return
        setor = setor_item.text()
        script_real_name = script_item.data(Qt.ItemDataRole.UserRole)
        module_name = script_real_name.replace('.py', '')
        module_path = f"scripts.{setor}.{module_name}"
        self.log.clear()
        self.log.append(f">>> Carregando: {module_path}")
        QApplication.processEvents()
        try:
            script_module = importlib.import_module(module_path)
            self.log.append(f">>> Executando '{script_item.text()}'...")
            app, window = script_module.main()
            self.open_windows.append(window)
        except ImportError as e:
            self.log.append(f"ERRO de Importação: Não foi possível carregar o módulo '{module_path}'.\nDetalhe: {e}")
        except AttributeError:
            self.log.append(f"ERRO: O script '{module_name}' não possui uma função 'main()'.")
        except Exception as e:
            self.log.append(f"ERRO inesperado ao executar '{module_name}': {e}")


# MUDANÇA AQUI: A função agora aceita 'app' como argumento
def start_application(app, splash):
    """Função que carrega os módulos pesados e inicia o Hub."""
    
    splash.set_status("Carregando bibliotecas de dados (pandas)...")
    import pandas
    
    splash.set_status("Carregando bibliotecas de relatórios (openpyxl)...")
    import openpyxl
    
    splash.set_status("Iniciando interface principal...")
    
    # MUDANÇA AQUI: Anexamos a janela do Hub ao objeto 'app' para criar uma referência persistente
    app.main_hub = Hub()
    
    QTimer.singleShot(500, lambda: (
        app.main_hub.show(),
        splash.close()
    ))
    # Não precisamos mais retornar a instância, pois ela está segura no 'app'


# ---- Ponto de Entrada Principal da Aplicação (MODIFICADO) ----
if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    font = QFont("Segoe UI", 10)
    app.setFont(font)
    
    splash = Splash()
    splash.show()
    
    # MUDANÇA AQUI: Passamos 'app' e 'splash' para a função de carregamento
    QTimer.singleShot(2000, lambda: start_application(app, splash))
    
    sys.exit(app.exec())