# ==============================================================================
# scripts/Analise/consolidacao_contabil_app.py
# Módulo de Interface: Consolidação e Evolução Contábil
# ==============================================================================

import sys
import os
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QLineEdit, QFileDialog, QTextEdit, QMessageBox
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont

try:
    from . import consolidacao_contabil_logic as logic
except ImportError:
    try:
        from scripts.Analise import consolidacao_contabil_logic as logic
    except ImportError as e:
        print(f"Erro ao importar a lógica: {e}")

STYLESHEET = """
    QWidget { font-family: 'Segoe UI', 'Arial'; font-size: 10pt; }
    QLabel#TitleLabel { font-size: 14pt; font-weight: bold; color: #202124; }
    QLineEdit { padding: 6px; border: 1px solid #DADCE0; border-radius: 4px; background-color: #FFFFFF; }
    QTextEdit { background-color: #F8F9FA; border: 1px solid #DADCE0; border-radius: 4px; font-family: 'Consolas', monospace; color: #3C4043; }
    
    QPushButton { border-radius: 4px; font-weight: bold; padding: 8px 15px; }
    QPushButton#UtilityButton { background-color: #F1F3F4; color: #333333; border: 1px solid #DADCE0; }
    QPushButton#UtilityButton:hover { background-color: #E8EAED; }
    
    QPushButton#HelpButton { background-color: #FFFFFF; color: #4285F4; border: 2px solid #4285F4; border-radius: 15px; font-size: 12pt; font-weight: bold; padding: 0px; }
    QPushButton#HelpButton:hover { background-color: #E8F0FE; }
    
    QPushButton#ActionA { background-color: #4285F4; color: white; border: none; }
    QPushButton#ActionA:hover { background-color: #3367D6; }
    
    QPushButton#ActionB { background-color: #34A853; color: white; border: none; }
    QPushButton#ActionB:hover { background-color: #2B8B45; }
    QPushButton#ActionB:disabled { background-color: #A5D6A7; color: #E8F5E9; }
"""

class ConsolidacaoContabilApp(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Consolidação e Evolução Contábil")
        self.resize(800, 600)
        self.setStyleSheet(STYLESHEET)
        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # Cabeçalho com Título e Botão de Ajuda
        header_layout = QHBoxLayout()
        lbl_title = QLabel("Consolidação e Evolução Contábil")
        lbl_title.setObjectName("TitleLabel")
        header_layout.addWidget(lbl_title)
        header_layout.addStretch()
        
        btn_help = QPushButton("?")
        btn_help.setObjectName("HelpButton")
        btn_help.setFixedSize(30, 30)
        btn_help.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_help.clicked.connect(self.mostrar_ajuda)
        btn_help.setToolTip("Clique aqui para entender como organizar seus arquivos ou baixar planilhas de exemplo.")
        header_layout.addWidget(btn_help)
        
        main_layout.addLayout(header_layout)

        # Campo 1: Pasta dos Balancetes
        main_layout.addWidget(QLabel("1. Selecionar Pasta dos Balancetes (Excel):"))
        box_balancetes = QHBoxLayout()
        self.input_balancetes = QLineEdit()
        self.input_balancetes.setPlaceholderText("Selecione o diretório contendo os balancetes das empresas...")
        self.input_balancetes.setReadOnly(True)
        btn_balancetes = QPushButton("Procurar Pasta")
        btn_balancetes.setObjectName("UtilityButton")
        btn_balancetes.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_balancetes.clicked.connect(self.procurar_pasta_balancetes)
        box_balancetes.addWidget(self.input_balancetes)
        box_balancetes.addWidget(btn_balancetes)
        main_layout.addLayout(box_balancetes)

        # Campo 2: Arquivo de Parâmetros
        main_layout.addWidget(QLabel("2. Selecionar Arquivo de Parâmetros (De-Para):"))
        box_params = QHBoxLayout()
        self.input_params = QLineEdit()
        self.input_params.setPlaceholderText("Selecione a planilha Parametros_Consolidacao.xlsx...")
        self.input_params.setReadOnly(True)
        btn_params = QPushButton("Procurar Arquivo")
        btn_params.setObjectName("UtilityButton")
        btn_params.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_params.clicked.connect(self.procurar_arquivo_parametros)
        box_params.addWidget(self.input_params)
        box_params.addWidget(btn_params)
        main_layout.addLayout(box_params)

        # Campo 3: Pasta de Destino
        main_layout.addWidget(QLabel("3. Selecionar Pasta de Destino (Relatórios):"))
        box_destino = QHBoxLayout()
        self.input_destino = QLineEdit()
        self.input_destino.setPlaceholderText("Onde os relatórios consolidados e logs devem ser salvos...")
        self.input_destino.setReadOnly(True)
        btn_destino = QPushButton("Procurar Pasta")
        btn_destino.setObjectName("UtilityButton")
        btn_destino.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_destino.clicked.connect(self.procurar_pasta_destino)
        box_destino.addWidget(self.input_destino)
        box_destino.addWidget(btn_destino)
        main_layout.addLayout(box_destino)

        # Botões de Ação
        main_layout.addSpacing(10)
        box_acoes = QHBoxLayout()
        
        self.btn_validacao = QPushButton("1. Executar Validação de Contas")
        self.btn_validacao.setObjectName("ActionA")
        self.btn_validacao.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_validacao.clicked.connect(self.iniciar_validacao)
        
        self.btn_consolidacao = QPushButton("2. Gerar Consolidação e Relatórios")
        self.btn_consolidacao.setObjectName("ActionB")
        self.btn_consolidacao.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_consolidacao.setEnabled(False)
        self.btn_consolidacao.clicked.connect(self.iniciar_consolidacao)
        
        box_acoes.addWidget(self.btn_validacao)
        box_acoes.addWidget(self.btn_consolidacao)
        main_layout.addLayout(box_acoes)

        # Log do Sistema
        main_layout.addSpacing(10)
        main_layout.addWidget(QLabel("Log de Execução:"))
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        main_layout.addWidget(self.log_text)

        self.log("Sistema iniciado. Aguardando seleção de arquivos...")

    def log(self, mensagem):
        self.log_text.append(f"> {mensagem}")
        QApplication.processEvents()

    # --- Funções de Ajuda e Templates ---
    def mostrar_ajuda(self):
        msg = QMessageBox(self)
        msg.setWindowTitle("Guia de Uso - Consolidação Contábil")
        msg.setIcon(QMessageBox.Icon.Information)
        
        texto_instrucoes = (
            "<h3>Como preparar seus arquivos para o sistema:</h3>"
            "<p><b>1. Os Balancetes:</b><br>"
            "Mantenha um arquivo Excel para cada empresa. As planilhas internas (abas) "
            "devem obrigatoriamente ser nomeadas com dois dígitos indicando o mês "
            "(ex: '01' para Jan, '02' para Fev, etc.). O layout deve conter as colunas "
            "<i>Conta, Descricao, Saldo_Anterior, Debito, Credito e Saldo_Atual</i>.</p>"
            ""
            "<p><b>2. O Arquivo de Parâmetros:</b><br>"
            "Este arquivo é o 'cérebro' da consolidação. Ele relaciona as contas analíticas "
            "dos balancetes de cada cliente à estrutura sintética do Balanço e DRE.</p>"
            ""
            "<p>Se você não possui estes arquivos formatados, clique no botão abaixo para gerar "
            "os modelos em branco (templates) e preenchê-los.</p>"
        )
        msg.setText(texto_instrucoes)
        
        btn_gerar_exemplo = msg.addButton("Exportar Planilhas de Exemplo", QMessageBox.ButtonRole.ActionRole)
        btn_fechar = msg.addButton("Entendi, Fechar", QMessageBox.ButtonRole.RejectRole)
        
        msg.exec()
        
        if msg.clickedButton() == btn_gerar_exemplo:
            self.exportar_planilhas_exemplo()

    def exportar_planilhas_exemplo(self):
        pasta_destino = QFileDialog.getExistingDirectory(self, "Selecione onde salvar as planilhas de exemplo")
        if pasta_destino:
            self.log("\nGerando planilhas de exemplo...")
            QApplication.processEvents()
            
            sucesso, mensagem = logic.exportar_templates(pasta_destino)
            
            if sucesso:
                self.log(mensagem)
                QMessageBox.information(self, "Sucesso", "Arquivos de exemplo gerados. Verifique a pasta selecionada.")
            else:
                self.log(f"Erro: {mensagem}")
                QMessageBox.critical(self, "Erro", mensagem)

    # --- Funções de Seleção de Diretórios/Arquivos ---
    def procurar_pasta_balancetes(self):
        pasta = QFileDialog.getExistingDirectory(self, "Selecionar Pasta dos Balancetes")
        if pasta:
            self.input_balancetes.setText(pasta)
            self.log(f"Pasta de balancetes selecionada: {pasta}")

    def procurar_arquivo_parametros(self):
        arquivo, _ = QFileDialog.getOpenFileName(self, "Selecionar Arquivo de Parâmetros", "", "Arquivos Excel (*.xlsx *.xls)")
        if arquivo:
            self.input_params.setText(arquivo)
            self.log(f"Arquivo de parâmetros selecionado: {arquivo}")

    def procurar_pasta_destino(self):
        pasta = QFileDialog.getExistingDirectory(self, "Selecionar Pasta de Destino")
        if pasta:
            self.input_destino.setText(pasta)
            self.log(f"Pasta de destino selecionada: {pasta}")

    def validar_campos(self):
        if not self.input_balancetes.text() or not self.input_params.text() or not self.input_destino.text():
            QMessageBox.warning(self, "Aviso", "Por favor, preencha todos os três caminhos antes de continuar.")
            return False
        return True

    # --- Funções de Ação (Conexão com a Lógica) ---
    def iniciar_validacao(self):
        if not self.validar_campos():
            return
            
        self.log("\nIniciando Validação de Contas e Arquivos...")
        self.btn_validacao.setEnabled(False)
        QApplication.processEvents()
        
        sucesso, mensagem = logic.executar_validacao_contas(
            self.input_balancetes.text(),
            self.input_params.text(),
            self.input_destino.text()
        )
        
        self.log(mensagem)
        
        # Como combinamos: se não ocorreu nenhum erro técnico fatal (True),
        # nós sempre liberamos o botão 2, mesmo havendo inconsistências na regra de negócio.
        if sucesso:
            self.btn_consolidacao.setEnabled(True)
            
        self.btn_validacao.setEnabled(True)

    def iniciar_consolidacao(self):
        self.log("\nIniciando Consolidação de Dados...")
        self.btn_consolidacao.setEnabled(False)
        QApplication.processEvents()
        
        logic.gerar_relatorios_consolidacao(
            self.input_balancetes.text(),
            self.input_params.text(),
            self.input_destino.text()
        )
        
        self.log("Consolidação finalizada (Modo Simulação).")
        self.btn_consolidacao.setEnabled(True)

def main():
    app = QApplication.instance()
    if app is None:
        app = QApplication(sys.argv)
    window = ConsolidacaoContabilApp()
    window.show()
    return app, window

if __name__ == "__main__":
    app, window = main()
    sys.exit(app.exec())