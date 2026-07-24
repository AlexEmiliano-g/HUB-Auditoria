# ==============================================================================
# HUB.PY - VERSÃO 0.9.1 (INTEGRAÇÃO COM DESIGN SYSTEM / THEME.QSS)
# ==============================================================================

import sys
import os
import importlib
import getpass

from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QListWidget, QPushButton, QLabel, QListWidgetItem
)
from PyQt6.QtGui import QFont, QPixmap, QIcon # type: ignore
from PyQt6.QtCore import Qt, QPropertyAnimation, QTimer, QDateTime # type: ignore

def resource_path(relative_path):
    """Garante que os arquivos sejam encontrados tanto no VSCode quanto após o build (.exe)"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        # Busca a pasta absoluta onde ESTE arquivo (hub.py) está
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

# Caminhos globais
SCRIPTS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "scripts")
LOGO_HUB_PATH = resource_path("logo.png")
LOGO_SPLASH_PATH = resource_path("logo_aber.png")
THEME_PATH = resource_path("theme_light.qss")

def load_theme(app):
    """Lê o arquivo QSS externo e aplica em toda a aplicação globalmente."""
    if os.path.exists(THEME_PATH):
        try:
            with open(THEME_PATH, "r", encoding="utf-8") as f:
                stylesheet = f.read()
                app.setStyleSheet(stylesheet)
                print("[SISTEMA] Design System (theme_light.qss) carregado e processado com sucesso.")
        except Exception as e:
            print(f"[ERRO] Falha ao ler ou aplicar o arquivo de tema: {e}")
    else:
        print(f"[AVISO] Arquivo de tema não encontrado em: {THEME_PATH}. Usando estilo padrão.")

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
        
        logo_label = QLabel("Logo não encontrado")
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
        
        window_layout = QVBoxLayout(self)
        window_layout.setContentsMargins(0, 0, 0, 0)
        window_layout.setSpacing(0)
        
        # --- ÁREA DE CONTEÚDO PRINCIPAL ---
        content_widget = QWidget()
        content_layout = QHBoxLayout(content_widget)
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(0)
        
        # --- PAINEL ESQUERDO ---
        left_panel = QWidget()
        left_panel.setObjectName("LeftPanel")
        left_panel.setFixedWidth(260)
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(15, 30, 15, 30) 
        left_layout.setSpacing(20)
        
        logo_label = QLabel()
        if os.path.exists(LOGO_HUB_PATH):
            pixmap = QPixmap(LOGO_HUB_PATH)
            scaled_pixmap = pixmap.scaled(220, 100, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
            logo_label.setPixmap(scaled_pixmap)
        logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        left_layout.addWidget(logo_label)
        
        self.sectors_list = QListWidget()
        self.sectors_list.setObjectName("SectorsList")
        
        for setor in ["Contabilidade", "Fiscal", "Analise"]:
            self.sectors_list.addItem(QListWidgetItem(setor))
            
        self.sectors_list.itemClicked.connect(self.load_scripts)
        left_layout.addWidget(self.sectors_list)
        content_layout.addWidget(left_panel)
        
        # --- PAINEL DIREITO ---
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(50, 40, 50, 40) 
        right_layout.setSpacing(20)
        
        self.header_title = QLabel("Selecione um setor no menu ao lado")
        self.header_title.setObjectName("HeaderLabel")
        right_layout.addWidget(self.header_title)
        
        self.scripts_list = QListWidget()
        self.scripts_list.setObjectName("ScriptsList")
        self.scripts_list.itemSelectionChanged.connect(self.check_selection)
        right_layout.addWidget(self.scripts_list)
        
        bottom_action_layout = QHBoxLayout()
        bottom_action_layout.addStretch() 
        
        self.run_button = QPushButton("Executar Programa")
        self.run_button.setObjectName("RunButton")
        self.run_button.setCursor(Qt.CursorShape.PointingHandCursor)
        self.run_button.setEnabled(False) 
        self.run_button.clicked.connect(self.run_script)
        
        bottom_action_layout.addWidget(self.run_button)
        right_layout.addLayout(bottom_action_layout)
        content_layout.addWidget(right_panel, 1)
        
        window_layout.addWidget(content_widget, 1)

        # --- RODAPÉ (FOOTER) ---
        footer_widget = QWidget()
        footer_widget.setObjectName("Footer")
        footer_layout = QHBoxLayout(footer_widget)
        footer_layout.setContentsMargins(20, 8, 20, 8)
        
        version_label = QLabel("Versão 0.9")
        version_label.setObjectName("FooterText")
        
        try:
            username = getpass.getuser()
        except:
            username = "Usuário Desconhecido"
            
        user_label = QLabel(f"Usuário: {username}")
        user_label.setObjectName("FooterText")
        user_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        self.datetime_label = QLabel()
        self.datetime_label.setObjectName("FooterText")
        self.datetime_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        
        footer_layout.addWidget(version_label, 1)
        footer_layout.addWidget(user_label, 1)
        footer_layout.addWidget(self.datetime_label, 1)
        
        window_layout.addWidget(footer_widget, 0)
        
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_datetime)
        self.timer.start(1000)
        self.update_datetime()

    def update_datetime(self):
        current_dt = QDateTime.currentDateTime()
        self.datetime_label.setText(current_dt.toString("dd/MM/yyyy HH:mm:ss"))

    def check_selection(self):
        if self.scripts_list.selectedItems():
            self.run_button.setEnabled(True)
        else:
            self.run_button.setEnabled(False)

    def load_scripts(self, item):
        setor = item.text()
        self.header_title.setText(f"{setor} > Programas Disponíveis")
        
        setor_path = resource_path(os.path.join("scripts", setor.replace("Á", "A").replace("á", "a")))
        self.scripts_list.clear()
        
        if not os.path.exists(setor_path):
            print(f"[AVISO] A pasta para o setor '{setor}' não foi encontrada em: {setor_path}")
            return
            
        for script_filename in sorted(os.listdir(setor_path)):
            if script_filename.endswith("_app.py"):
                display_name = script_filename.replace("_app.py", "").replace("_", " ").capitalize()
                
                nomes_bonitos = {
                    "Analise evolucao": "Análise de Evolução",
                    "Analise fiscal": "Análise Fiscal",
                    "Consolidacao contabil": "Consolidação Contábil",
                    "Tabulador": "Tabulador de Balancetes",
                    "Relatorio c100 efd c": "Relatório C100 (EFD Contribuições)",
                    "Relatorio c100 sped f": "Relatório C100 (SPED Fiscal)",
                    "Relatorio c500 efd c": "Relatório C500 (EFD Contribuições)",
                    "Relatorio c500 sped f": "Relatório C500 (SPED Fiscal)",
                    "Relatorio ciap sped f": "Relatório CIAP (SPED Fiscal)",
                    "Relatorio ct": "Relatório Conhecimento de Transporte (CT)",
                    "Relatorio d100 sped f": "Relatório D100 (SPED Fiscal)",
                    "Relatorio f100 efd c": "Relatório F100 (EFD Contribuições)",
                    "Relatorio xml": "Relatório XML",
                    "Texto anonimo": "Anonimizador de Texto (IA)"
                }
                
                if display_name in nomes_bonitos:
                    display_name = nomes_bonitos[display_name]

                list_item = QListWidgetItem(display_name)
                list_item.setData(Qt.ItemDataRole.UserRole, script_filename)
                self.scripts_list.addItem(list_item)

    def run_script(self):
        setor_item = self.sectors_list.currentItem()
        script_item = self.scripts_list.currentItem()
        
        if not setor_item or not script_item:
            print("[SISTEMA] Tentativa de execução sem programa selecionado ignorada.")
            return
            
        setor = setor_item.text().replace("Á", "A").replace("á", "a")
        script_real_name = script_item.data(Qt.ItemDataRole.UserRole)
        module_name = script_real_name.replace('.py', '')
        module_path = f"scripts.{setor}.{module_name}"
        
        print(f"\n[{setor.upper()}] Inicializando: {module_path}")
        QApplication.processEvents()
        
        try:
            script_module = importlib.import_module(module_path)
            print(f"[{setor.upper()}] Executando '{script_item.text()}'...")
            app, window = script_module.main() 
            self.open_windows.append(window)
            print(f"[{setor.upper()}] Sucesso. Janela aberta.")
        except ImportError as e:
            print(f"[ERRO CRÍTICO] Importação falhou para o módulo '{module_path}'.\n-> Detalhe: {e}")
        except AttributeError:
            print(f"[ERRO CRÍTICO] O script '{module_name}' não possui uma função 'main()'.")
        except Exception as e:
            print(f"[ERRO INESPERADO] Falha ao executar '{module_name}': {e}")

def start_application(app, splash):
    splash.set_status("Carregando bibliotecas de dados (pandas)...")
    import pandas
    
    splash.set_status("Carregando bibliotecas de relatórios (openpyxl)...")
    import openpyxl
    
    splash.set_status("Iniciando interface principal...")
    
    app.main_hub = Hub()
    
    QTimer.singleShot(500, lambda: (
        app.main_hub.show(),
        splash.close()
    ))

if __name__ == "__main__":
    if hasattr(Qt, 'AA_EnableHighDpiScaling'):
        QApplication.setAttribute(Qt.ApplicationAttribute.AA_EnableHighDpiScaling, True)
    if hasattr(Qt, 'AA_UseHighDpiPixmaps'):
        QApplication.setAttribute(Qt.ApplicationAttribute.AA_UseHighDpiPixmaps, True)

    app = QApplication(sys.argv)
    
    # Aplica a Fonte Global
    font = QFont("Segoe UI", 10)
    app.setFont(font)
    
    # INJEÇÃO DO DESIGN SYSTEM AQUI (Antes de abrir qualquer tela)
    load_theme(app)
    
    splash = Splash()
    splash.show()
    
    QTimer.singleShot(2000, lambda: start_application(app, splash))
    
    sys.exit(app.exec())