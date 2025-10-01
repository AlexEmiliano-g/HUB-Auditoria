import sys
import os
import subprocess
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QListWidget, QPushButton, QTextEdit, QLabel, QListWidgetItem
)
from PyQt6.QtGui import QFont, QPixmap, QIcon
from PyQt6.QtCore import Qt, QPropertyAnimation, QTimer, QSize


# --- DEFINIÇÕES GLOBAIS ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SCRIPTS_DIR = os.path.join(BASE_DIR, "scripts")

# --- Caminhos das logos ---
LOGO_HUB_PATH = os.path.join(BASE_DIR, "logo.png")
LOGO_SPLASH_PATH = os.path.join(BASE_DIR, "logo_aber.png")


# ---- Estilo Visual (Stylesheet) ----
STYLESHEET = """
QWidget#HubWindow {
    background-color: #F8F9FA;
}
QWidget#LeftPanel {
    background-color: #FFFFFF;
    border-right: 1px solid #E0E0E0;
    border-top-left-radius: 10px;
    border-bottom-left-radius: 10px;
}
QListWidget {
    border: none;
    font-size: 14px;
    padding: 5px;
    outline: 0;
}
QListWidget::item {
    padding: 12px 15px;
    border-radius: 5px;
}
QListWidget::item:hover {
    background-color: #F1F3F4;
}
QListWidget::item:selected {
    background-color: #E8F0FE;
    color: #1967D2;
    font-weight: bold;
}
QPushButton {
    background-color: #4285F4;
    color: white;
    font-size: 14px;
    font-weight: bold;
    border: none;
    padding: 12px;
    border-radius: 5px;
}
QPushButton:hover {
    background-color: #1A73E8;
}
QPushButton:pressed {
    background-color: #1765CC;
}
QTextEdit {
    background-color: #FFFFFF;
    border: 1px solid #DADCE0;
    border-radius: 5px;
    font-family: 'Consolas', 'Courier New', monospace;
    color: #3C4043;
    padding: 8px;
}
QLabel#TitleLabel {
    font-size: 16px;
    font-weight: bold;
    color: #202124;
    padding-bottom: 5px;
}
"""

# ---- Tela inicial com fundo cinza, logo e fade-in ----
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

        self.setWindowOpacity(0.0)
        self.fade_anim = QPropertyAnimation(self, b"windowOpacity")
        self.fade_anim.setDuration(2000)
        self.fade_anim.setStartValue(0.0)
        self.fade_anim.setEndValue(1.0)
        self.fade_anim.start()

        QTimer.singleShot(3000, self.open_hub)

    def open_hub(self):
        self.close()
        self.hub = Hub()
        self.hub.show()


# ---- Hub principal ----
# ===== CLASSE HUB ATUALIZADA COM LAYOUT AJUSTADO =====

class Hub(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("DM HUB - Central de Aplicativos")
        if os.path.exists(LOGO_HUB_PATH):
            self.setWindowIcon(QIcon(LOGO_HUB_PATH))
        self.resize(1100, 700)
        self.setObjectName("HubWindow")
        self.setStyleSheet(STYLESHEET)

        main_layout = QHBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # ---- Painel Esquerdo (Menu) ----
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
        for setor in ["Contabilidade", "Folha", "Fiscal", "Controle Interno"]:
            self.sectors_list.addItem(QListWidgetItem(setor))
        self.sectors_list.itemClicked.connect(self.load_scripts)
        left_layout.addWidget(self.sectors_list)
        main_layout.addWidget(left_panel)

        # ---- Painel Direito (Conteúdo) ----
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(25, 20, 25, 20)
        right_layout.setSpacing(15)

        scripts_title = QLabel("Programas Disponíveis")
        scripts_title.setObjectName("TitleLabel")
        right_layout.addWidget(scripts_title)
        
        self.scripts_list = QListWidget()
        # MUDANÇA AQUI: Aumentamos o fator de estiramento da lista de programas
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
        # MUDANÇA AQUI: Diminuímos o fator de estiramento da caixa de log
        right_layout.addWidget(self.log, 1)

        main_layout.addWidget(right_panel, 1)

    def load_scripts(self, item):
        setor = item.text()
        setor_path = os.path.join(SCRIPTS_DIR, setor)
        self.scripts_list.clear()
        
        if not os.path.exists(setor_path):
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
        
        if not script_real_name:
            self.log.append(f"ERRO: O item '{script_item.text()}' não tem um arquivo real associado.")
            return

        script_path = os.path.join(SCRIPTS_DIR, setor, script_real_name)

        if not os.path.isfile(script_path):
            self.log.append(f"ERRO: Arquivo principal não encontrado em {script_path}")
            return

        self.log.clear()
        self.log.append(f">>> Iniciando: {setor} / {script_item.text()}")
        QApplication.processEvents()

        if script_real_name.endswith("_app.py"):
            try:
                self.log.append("Iniciando interface em uma nova janela...")
                subprocess.Popen([sys.executable, script_path], cwd=os.path.dirname(script_path))
            except Exception as e:
                self.log.append(f"ERRO ao iniciar a interface: {e}")
        else: # Para scripts de linha de comando
            try:
                completed = subprocess.run(
                    [sys.executable, script_path], capture_output=True,
                    text=True, timeout=60, cwd=os.path.dirname(script_path),
                    encoding='utf-8', errors='replace'
                )
                
                if completed.returncode == 0:
                    self.log.append(">>> Programa executado com sucesso.")
                else:
                    self.log.append(f">>> ERRO: Programa finalizou com o código {completed.returncode}.")
                    if completed.stderr:
                        self.log.append(f"--- Detalhes do Erro ---\n{completed.stderr}")
                    if completed.stdout:
                        self.log.append(f"--- Saída Padrão ---\n{completed.stdout}")
            except Exception as e:
                self.log.append(f"ERRO inesperado ao executar script: {str(e)}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # Define uma fonte padrão mais agradável para toda a aplicação
    font = QFont("Segoe UI", 10)
    app.setFont(font)
    
    splash = Splash()
    splash.show()
    sys.exit(app.exec())