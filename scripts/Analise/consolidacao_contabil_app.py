# ==============================================================================
# scripts/Analise/consolidacao_contabil_app.py
# Módulo de Interface: Consolidação e Evolução Contábil
# ==============================================================================

import sys
import os
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QLineEdit, QFileDialog, QTextEdit, QMessageBox,
    QCheckBox, QGroupBox, QDialog, QTableWidget, QTableWidgetItem, QHeaderView
)
from PyQt6.QtCore import Qt

try:
    from . import consolidacao_contabil_logic as logic
except ImportError:
    try:
        from scripts.Analise import consolidacao_contabil_logic as logic
    except ImportError as e:
        print(f"Erro ao importar a camada lógica: {e}")

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

    QGroupBox { font-weight: bold; border: 1px solid #DADCE0; border-radius: 4px; margin-top: 10px; padding-top: 15px; }
    QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 3px 0 3px; color: #202124; }
    
    QTableWidget { background-color: #FFFFFF; border: 1px solid #DADCE0; border-radius: 4px; }
    QHeaderView::section { background-color: #F1F3F4; padding: 4px; border: 1px solid #DADCE0; font-weight: bold; }
"""

class DialogContasNovas(QDialog):
    """ Interface para conciliação de contas órfãs encontradas nos balancetes """
    def __init__(self, contas_novas, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Auditoria: Novas Contas Detectadas")
        self.resize(1100, 500)
        self.setModal(True)
        self.contas_novas = contas_novas
        self.contas_selecionadas = []
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        
        lbl_info = QLabel("Foram identificadas contas nos balancetes que não constam no arquivo de parâmetros.\n"
                          "Preencha o 'Código de Vinculação' (Sintética do Balanço) para as contas que deseja adicionar ao respectivo Plano.")
        lbl_info.setWordWrap(True)
        layout.addWidget(lbl_info)

        self.table = QTableWidget(len(self.contas_novas), 5)
        self.table.setHorizontalHeaderLabels(["Incluir", "Conta", "Descrição", "Plano Destino", "Código Vinculação (Balanço)"])
        
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.Interactive)
        self.table.setColumnWidth(4, 220)

        for row, dados in enumerate(self.contas_novas):
            widget_chk = QWidget()
            layout_chk = QHBoxLayout(widget_chk)
            layout_chk.setContentsMargins(0, 0, 0, 0)
            layout_chk.setAlignment(Qt.AlignmentFlag.AlignCenter)
            
            chk = QCheckBox()
            chk.setChecked(True)
            chk.setCursor(Qt.CursorShape.PointingHandCursor)
            layout_chk.addWidget(chk)
            
            self.table.setCellWidget(row, 0, widget_chk)
            
            conta_item = QTableWidgetItem(dados['conta'])
            conta_item.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            self.table.setItem(row, 1, conta_item)
            
            desc_item = QTableWidgetItem(dados['descricao'])
            desc_item.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            desc_item.setToolTip(dados['descricao'])
            self.table.setItem(row, 2, desc_item)
            
            plano_item = QTableWidgetItem(dados['plano'])
            plano_item.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            self.table.setItem(row, 3, plano_item)
            
            vinculo_item = QTableWidgetItem("")
            self.table.setItem(row, 4, vinculo_item)

        layout.addWidget(self.table)

        box_botoes = QHBoxLayout()
        btn_cancelar = QPushButton("Ignorar e Continuar")
        btn_cancelar.clicked.connect(self.reject)
        
        btn_salvar = QPushButton("Adicionar Selecionadas ao Plano")
        btn_salvar.setObjectName("ActionA")
        btn_salvar.clicked.connect(self.processar_selecao)
        
        box_botoes.addStretch()
        box_botoes.addWidget(btn_cancelar)
        box_botoes.addWidget(btn_salvar)
        
        layout.addLayout(box_botoes)

    def processar_selecao(self):
        for row in range(self.table.rowCount()):
            widget_chk = self.table.cellWidget(row, 0)
            chk = widget_chk.findChild(QCheckBox)
            
            if chk and chk.isChecked():
                conta = self.table.item(row, 1).text()
                descricao = self.table.item(row, 2).text()
                plano = self.table.item(row, 3).text()
                vinculo = self.table.item(row, 4).text().strip()
                
                self.contas_selecionadas.append({
                    'conta': conta,
                    'descricao': descricao,
                    'plano': plano,
                    'vinculo': vinculo
                })
        self.accept()

class ConsolidacaoContabilApp(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Consolidação e Evolução Contábil")
        self.resize(800, 650)
        self.setStyleSheet(STYLESHEET)
        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)

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
        header_layout.addWidget(btn_help)
        
        main_layout.addLayout(header_layout)

        # Seleção de Diretórios e Arquivos
        main_layout.addWidget(QLabel("1. Selecionar Pasta dos Balancetes (Excel):"))
        box_balancetes = QHBoxLayout()
        self.input_balancetes = QLineEdit()
        self.input_balancetes.setReadOnly(True)
        btn_balancetes = QPushButton("Procurar Pasta")
        btn_balancetes.setObjectName("UtilityButton")
        btn_balancetes.clicked.connect(self.procurar_pasta_balancetes)
        box_balancetes.addWidget(self.input_balancetes)
        box_balancetes.addWidget(btn_balancetes)
        main_layout.addLayout(box_balancetes)

        main_layout.addWidget(QLabel("2. Selecionar Arquivo de Parâmetros (De-Para):"))
        box_params = QHBoxLayout()
        self.input_params = QLineEdit()
        self.input_params.setReadOnly(True)
        btn_params = QPushButton("Procurar Arquivo")
        btn_params.setObjectName("UtilityButton")
        btn_params.clicked.connect(self.procurar_arquivo_parametros)
        box_params.addWidget(self.input_params)
        box_params.addWidget(btn_params)
        main_layout.addLayout(box_params)

        main_layout.addWidget(QLabel("3. Selecionar Pasta de Destino (Relatórios):"))
        box_destino = QHBoxLayout()
        self.input_destino = QLineEdit()
        self.input_destino.setReadOnly(True)
        btn_destino = QPushButton("Procurar Pasta")
        btn_destino.setObjectName("UtilityButton")
        btn_destino.clicked.connect(self.procurar_pasta_destino)
        box_destino.addWidget(self.input_destino)
        box_destino.addWidget(btn_destino)
        main_layout.addLayout(box_destino)

        # Seleção de Relatórios
        group_relatorios = QGroupBox("4. Relatórios a Gerar")
        layout_relatorios = QVBoxLayout()
        
        self.chk_rel1 = QCheckBox("Relatório 1: Comparativo entre Empresas por Mês")
        self.chk_rel1.setChecked(True)
        self.chk_rel2 = QCheckBox("Relatório 2: Consolidado Global")
        self.chk_rel2.setChecked(True)
        self.chk_rel3 = QCheckBox("Relatório 3: Evolução Individual por Empresa")
        self.chk_rel3.setChecked(True)
        
        layout_relatorios.addWidget(self.chk_rel1)
        layout_relatorios.addWidget(self.chk_rel2)
        layout_relatorios.addWidget(self.chk_rel3)
        group_relatorios.setLayout(layout_relatorios)
        main_layout.addWidget(group_relatorios)

        # Botões de Ação
        main_layout.addSpacing(10)
        box_acoes = QHBoxLayout()
        
        self.btn_validacao = QPushButton("Validar")
        self.btn_validacao.setObjectName("ActionA")
        self.btn_validacao.clicked.connect(self.iniciar_validacao)
        
        self.btn_consolidacao = QPushButton("Gerar Relatórios")
        self.btn_consolidacao.setObjectName("ActionB")
        self.btn_consolidacao.setEnabled(False)
        self.btn_consolidacao.clicked.connect(self.iniciar_consolidacao)
        
        box_acoes.addWidget(self.btn_validacao)
        box_acoes.addWidget(self.btn_consolidacao)
        main_layout.addLayout(box_acoes)

        # Log
        main_layout.addSpacing(10)
        main_layout.addWidget(QLabel("Log de Execução:"))
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        main_layout.addWidget(self.log_text)

        self.log("Sistema iniciado. Aguardando seleção de arquivos...")

    def log(self, mensagem):
        self.log_text.append(f"> {mensagem}")
        QApplication.processEvents()

    def mostrar_ajuda(self):
        msg = QMessageBox(self)
        msg.setWindowTitle("Guia de Uso - Consolidação Contábil")
        msg.setIcon(QMessageBox.Icon.Information)
        texto = ("<p><b>1. Balancetes:</b> Um arquivo por empresa, com abas mensais (ex: '01', '02').</p>"
                 "<p><b>2. Parâmetros:</b> Arquivo contendo o Cadastro de Empresas e o Plano de Contas (De-Para).</p>")
        msg.setText(texto)
        btn_exemplo = msg.addButton("Exportar Templates", QMessageBox.ButtonRole.ActionRole)
        msg.addButton("Fechar", QMessageBox.ButtonRole.RejectRole)
        msg.exec()
        if msg.clickedButton() == btn_exemplo:
            pasta = QFileDialog.getExistingDirectory(self, "Selecione o destino dos templates")
            if pasta:
                sucesso, msg_retorno = logic.exportar_templates(pasta)
                if sucesso: QMessageBox.information(self, "Sucesso", msg_retorno)
                else: QMessageBox.critical(self, "Erro", msg_retorno)

    def procurar_pasta_balancetes(self):
        pasta = QFileDialog.getExistingDirectory(self, "Selecionar Pasta dos Balancetes")
        if pasta: self.input_balancetes.setText(pasta)

    def procurar_arquivo_parametros(self):
        arquivo, _ = QFileDialog.getOpenFileName(self, "Selecionar Arquivo", "", "Excel (*.xlsx *.xls)")
        if arquivo: self.input_params.setText(arquivo)

    def procurar_pasta_destino(self):
        pasta = QFileDialog.getExistingDirectory(self, "Selecionar Pasta de Destino")
        if pasta: self.input_destino.setText(pasta)

    def validar_campos(self):
        if not self.input_balancetes.text() or not self.input_params.text() or not self.input_destino.text():
            QMessageBox.warning(self, "Aviso", "Preencha todos os diretórios antes de prosseguir.")
            return False
        return True

    def iniciar_validacao(self):
        if not self.validar_campos(): return
            
        self.log("\nIniciando auditoria de contas e arquivos...")
        self.btn_validacao.setEnabled(False)
        QApplication.processEvents()
        
        sucesso, mensagem = logic.executar_validacao_contas(self.input_balancetes.text(), self.input_params.text())
        if not sucesso:
            self.log(mensagem)
            self.btn_validacao.setEnabled(True)
            return

        # Busca por contas órfãs nos balancetes
        sucesso_busca, resultado = logic.verificar_contas_novas(self.input_balancetes.text(), self.input_params.text())
        
        if sucesso_busca and isinstance(resultado, list) and len(resultado) > 0:
            self.log(f"Foram encontradas {len(resultado)} contas não mapeadas. Aguardando intervenção do usuário...")
            dialog = DialogContasNovas(resultado, self)
            if dialog.exec() == QDialog.DialogCode.Accepted and dialog.contas_selecionadas:
                self.log("Processando inclusão de novas contas no arquivo de parâmetros...")
                QApplication.processEvents()
                
                status_inclusao, msg_inclusao = logic.adicionar_contas_plano(self.input_params.text(), dialog.contas_selecionadas)
                self.log(msg_inclusao)
            else:
                self.log("Ação ignorada pelo usuário. As contas não mapeadas não serão consolidadas.")
        else:
            self.log("Nenhuma conta órfã identificada. O Plano de Contas está atualizado.")

        self.btn_consolidacao.setEnabled(True)
        self.btn_validacao.setEnabled(True)

    def iniciar_consolidacao(self):
        rel1, rel2, rel3 = self.chk_rel1.isChecked(), self.chk_rel2.isChecked(), self.chk_rel3.isChecked()
        if not (rel1 or rel2 or rel3):
            QMessageBox.warning(self, "Aviso", "Selecione pelo menos um relatório para gerar.")
            return

        self.log("\nIniciando extração e consolidação de dados...")
        self.btn_consolidacao.setEnabled(False)
        QApplication.processEvents()
        
        try:
            sucesso, mensagem = logic.gerar_relatorios_consolidacao(
                self.input_balancetes.text(), self.input_params.text(), self.input_destino.text(),
                gerar_rel1=rel1, gerar_rel2=rel2, gerar_rel3=rel3
            )
            self.log(mensagem)
        except Exception as e:
            self.log(f"Falha na execução: {str(e)}")
            
        self.btn_consolidacao.setEnabled(True)

def main():
    app = QApplication.instance()
    if app is None: app = QApplication(sys.argv)
    window = ConsolidacaoContabilApp()
    window.show()
    return app, window

if __name__ == "__main__":
    app, window = main()
    sys.exit(app.exec())