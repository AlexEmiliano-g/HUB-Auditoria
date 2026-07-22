import os
from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QLabel, QPushButton, 
                             QFileDialog, QCheckBox, QGroupBox, QHBoxLayout, 
                             QMessageBox, QProgressBar, QScrollArea, QGridLayout, QLineEdit, QApplication)
from PyQt6.QtCore import Qt
from scripts.Contabilidade.analise_evolucao_logic import executar_analise_evolucao

APP_STYLESHEET = """
QWidget#ScrollContent { background-color: #FFFFFF; }
QLabel#AppTitle { font-size: 22px; font-weight: 600; color: #1A73E8; padding: 10px 0; }
QLabel#SectionTitle { font-size: 14px; font-weight: bold; color: #202124; margin-top: 15px; margin-bottom: 5px; }
QLabel { font-size: 12px; color: #3C4043; }
QGroupBox { background-color: #F8F9FA; border: 1px solid #E1E4E8; border-radius: 6px; font-size: 13px; font-weight: 600; color: #333; margin-top: 10px; }
QGroupBox::title { subcontrol-origin: margin; subcontrol-position: top left; padding: 0 10px; }
QPushButton#btn_processar { background-color: #1A73E8; color: white; font-weight: 600; border: none; padding: 8px 15px; border-radius: 4px; }
QPushButton#btn_processar:hover { background-color: #1557B0; }
QPushButton#btn_reprocessar { background-color: #34A853; color: white; font-weight: 600; border: none; padding: 8px 15px; border-radius: 4px; }
QPushButton#btn_reprocessar:hover { background-color: #2b8c44; }
QPushButton:disabled { background-color: #A0C3FF; }
QScrollArea { border: none; }
QLineEdit { background-color: #F1F3F4; border: 1px solid #DADCE0; border-radius: 4px; padding: 4px; color: #202124; font-weight: bold; }
QLineEdit:focus { border: 1px solid #1A73E8; background-color: #FFFFFF; }
"""

class AnaliseEvolucaoApp(QWidget):
    def __init__(self):
        super().__init__()
        self.caminho_arquivo = ""
        self.ultimo_pta_gerado = ""
        self.initUI()

    def initUI(self):
        self.setStyleSheet(APP_STYLESHEET)
        
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        content = QWidget()
        content.setObjectName("ScrollContent")
        layout = QVBoxLayout(content)
        layout.setContentsMargins(25, 20, 25, 20)
        layout.setSpacing(15)
        
        layout.addWidget(QLabel("Análise de Evolução", objectName="AppTitle"))

        g_arq = QGroupBox("1. Arquivo e Master Data")
        lay_arq = QVBoxLayout()
        lay_arq.setContentsMargins(15, 20, 15, 15)
        
        self.lbl_arq = QLabel("Nenhum arquivo selecionado.")
        self.lbl_arq.setStyleSheet("color: #5F6368;")
        btn_arq = QPushButton("Procurar Arquivo Original")
        btn_arq.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_arq.setObjectName("btn_processar")
        btn_arq.setToolTip("Selecione o arquivo Excel bruto contendo o balancete do cliente.")
        btn_arq.clicked.connect(self.selecionar_arquivo)
        
        self.chk_inclusao = QCheckBox("Habilitar Inclusão e Classificação Automática de Contas Órfãs")
        self.chk_inclusao.setChecked(True)
        self.chk_inclusao.setStyleSheet("margin-top: 10px; font-weight: normal;")
        self.chk_inclusao.setToolTip("Identifica contas não cadastradas e as inclui no plano de contas, aplicando inteligência S/A e A/P/R.")
        
        lay_arq.addWidget(btn_arq)
        lay_arq.addWidget(self.lbl_arq)
        lay_arq.addWidget(self.chk_inclusao)
        g_arq.setLayout(lay_arq)
        layout.addWidget(g_arq)

        g_escopo = QGroupBox("2. Parametrização de Risco (Limites de Auditoria)")
        lay_escopo = QVBoxLayout()
        lay_escopo.setContentsMargins(15, 20, 15, 15)
        
        lay_check = QHBoxLayout()
        self.chk_ativo = QCheckBox("PTA - Ativo"); self.chk_ativo.setChecked(True)
        self.chk_passivo = QCheckBox("PTA - Passivo"); self.chk_passivo.setChecked(True)
        self.chk_resultado = QCheckBox("PTA - Resultado"); self.chk_resultado.setChecked(True)
        lay_check.addWidget(self.chk_ativo); lay_check.addWidget(self.chk_passivo); lay_check.addWidget(self.chk_resultado)
        lay_escopo.addLayout(lay_check)

        self.container_ativo, self.regras_ativo_ui = self.criar_formulario_basico("REGRAS ATIVO:", "Chave D&M")
        self.container_passivo, self.regras_passivo_ui = self.criar_formulario_basico("REGRAS PASSIVO:", "Chave Cliente")
        self.container_resultado, self.regras_resultado_ui = self.criar_formulario_composto("REGRAS RESULTADO:")
        
        lay_escopo.addWidget(self.container_ativo)
        lay_escopo.addWidget(self.container_passivo)
        lay_escopo.addWidget(self.container_resultado)
        
        self.chk_ativo.toggled.connect(self.container_ativo.setVisible)
        self.chk_passivo.toggled.connect(self.container_passivo.setVisible)
        self.chk_resultado.toggled.connect(self.container_resultado.setVisible)
        
        g_escopo.setLayout(lay_escopo)
        layout.addWidget(g_escopo)
        layout.addStretch()
        
        self.barra_progresso = QProgressBar()
        self.barra_progresso.setFixedHeight(12)
        self.barra_progresso.setTextVisible(False)
        layout.addWidget(self.barra_progresso)
        
        lay_botoes = QHBoxLayout()
        self.btn_processar = QPushButton("Processar Base Original")
        self.btn_processar.setObjectName("btn_processar")
        self.btn_processar.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_processar.setFixedHeight(45)
        self.btn_processar.setToolTip("Cria o PTA inicial a partir do arquivo bruto do cliente.")
        self.btn_processar.clicked.connect(lambda: self.processar_dados(is_reprocess=False))
        
        self.btn_reprocessar = QPushButton("Atualizar Balancete Histórico")
        self.btn_reprocessar.setObjectName("btn_reprocessar")
        self.btn_reprocessar.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_reprocessar.setFixedHeight(45)
        self.btn_reprocessar.setToolTip("Refaz os cálculos do Balancete usando os parâmetros que você editou no arquivo PTA.")
        self.btn_reprocessar.clicked.connect(lambda: self.processar_dados(is_reprocess=True))
        
        lay_botoes.addWidget(self.btn_processar)
        lay_botoes.addWidget(self.btn_reprocessar)
        layout.addLayout(lay_botoes)

        scroll.setWidget(content)
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.addWidget(scroll)
        self.resize(950, 750)
        self.setWindowTitle("Análise de Evolução")

    def criar_formulario_basico(self, titulo, nome_chave):
        container = QWidget(); lay = QVBoxLayout(container); lay.setContentsMargins(0, 5, 0, 10)
        lbl = QLabel(titulo); lbl.setObjectName("SectionTitle"); lay.addWidget(lbl)
        grid = QGridLayout(); grid.setSpacing(10); grid.setVerticalSpacing(12)
        mins = ["0,10%", "0,30%", "0,60%", "0,90%", "1,20%", "1,50%", "1,80%", "2,00%"]
        maxs = ["0,30%", "0,60%", "0,90%", "1,20%", "1,50%", "1,80%", "2,00%", ""]
        lista_inputs = []
        for i in range(8):
            lbl_faixa = QLabel(f"Faixa {i+1}:"); lbl_faixa.setStyleSheet("font-weight: bold; color: #5F6368;")
            inp_min = QLineEdit(mins[i]); inp_min.setFixedWidth(70); inp_min.setAlignment(Qt.AlignmentFlag.AlignCenter)
            inp_max = QLineEdit(maxs[i]); inp_max.setFixedWidth(70); inp_max.setAlignment(Qt.AlignmentFlag.AlignCenter)
            inp_chave = QLineEdit(""); inp_chave.setFixedWidth(100)
            grid.addWidget(lbl_faixa, i, 0); grid.addWidget(QLabel("Saldo Acum. >"), i, 1); grid.addWidget(inp_min, i, 2)
            grid.addWidget(QLabel("até"), i, 3); grid.addWidget(inp_max, i, 4); grid.addWidget(QLabel(f"|  {nome_chave}:"), i, 5); grid.addWidget(inp_chave, i, 6)
            lista_inputs.append({'min': inp_min, 'max': inp_max, 'chave': inp_chave})
        grid.setColumnStretch(7, 1); lay.addLayout(grid)
        return container, lista_inputs

    def criar_formulario_composto(self, titulo):
        container = QWidget(); lay = QVBoxLayout(container); lay.setContentsMargins(0, 5, 0, 10)
        lbl = QLabel(titulo); lbl.setObjectName("SectionTitle"); lay.addWidget(lbl)
        grid = QGridLayout(); grid.setSpacing(10); grid.setVerticalSpacing(12)
        mins = ["0,20%", "0,40%", "0,70%", "1,00%", "1,30%", "1,60%", "1,90%", "2,10%"]
        maxs = ["0,40%", "0,70%", "1,00%", "1,30%", "1,60%", "1,90%", "2,10%", ""]
        percs = ["100%", "75%", "50%", "25%", "20%", "15%", "10%", "5%"]
        lista_inputs = []
        for i in range(8):
            lbl_faixa = QLabel(f"Faixa {i+1}:"); lbl_faixa.setStyleSheet("font-weight: bold; color: #5F6368;")
            inp_min = QLineEdit(mins[i]); inp_min.setFixedWidth(65); inp_min.setAlignment(Qt.AlignmentFlag.AlignCenter)
            inp_max = QLineEdit(maxs[i]); inp_max.setFixedWidth(65); inp_max.setAlignment(Qt.AlignmentFlag.AlignCenter)
            inp_chave = QLineEdit("30101"); inp_chave.setFixedWidth(80); inp_chave.setAlignment(Qt.AlignmentFlag.AlignCenter)
            inp_perc = QLineEdit(percs[i]); inp_perc.setFixedWidth(60); inp_perc.setAlignment(Qt.AlignmentFlag.AlignCenter)
            grid.addWidget(lbl_faixa, i, 0); grid.addWidget(QLabel("Saldo >"), i, 1); grid.addWidget(inp_min, i, 2)
            grid.addWidget(QLabel("até"), i, 3); grid.addWidget(inp_max, i, 4); grid.addWidget(QLabel("|  Chave D&M:"), i, 5); grid.addWidget(inp_chave, i, 6)
            grid.addWidget(QLabel("E Movimento (mês) >="), i, 7); grid.addWidget(inp_perc, i, 8); grid.addWidget(QLabel("da média"), i, 9)
            lista_inputs.append({'min': inp_min, 'max': inp_max, 'chave': inp_chave, 'perc': inp_perc})
        grid.setColumnStretch(10, 1); lay.addLayout(grid)
        return container, lista_inputs

    def selecionar_arquivo(self):
        caminho, _ = QFileDialog.getOpenFileName(self, "Selecionar Arquivo Original", "", "Excel Files (*.xlsx *.xlsb)")
        if caminho:
            self.caminho_arquivo = caminho
            self.lbl_arq.setText(os.path.basename(caminho))

    def processar_dados(self, is_reprocess=False):
        if not self.caminho_arquivo:
            QMessageBox.warning(self, "Aviso", "Por favor, selecione o arquivo Excel Original primeiro (Base de saldos).")
            return

        opcoes = {
            'ativo': self.chk_ativo.isChecked(),
            'passivo': self.chk_passivo.isChecked(),
            'resultado': self.chk_resultado.isChecked(),
            'inclusao_inteligente': self.chk_inclusao.isChecked(),
            'is_reprocess': is_reprocess
        }

        if is_reprocess:
            caminho_pta_alvo = ""
            if self.ultimo_pta_gerado and os.path.exists(self.ultimo_pta_gerado):
                resp = QMessageBox.question(self, "Atualizar", 
                    f"Deseja reprocessar usando os parâmetros do último arquivo gerado?\n\n{os.path.basename(self.ultimo_pta_gerado)}\n\n(Clique em 'No' para buscar outro arquivo).",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                if resp == QMessageBox.StandardButton.Yes:
                    caminho_pta_alvo = self.ultimo_pta_gerado
            
            if not caminho_pta_alvo:
                caminho_pta_alvo, _ = QFileDialog.getOpenFileName(self, "Selecionar Arquivo PTA Editado", "", "Excel Files (*.xlsx *.xlsb)")
                if not caminho_pta_alvo: return
                
            opcoes['caminho_pta_reprocess'] = caminho_pta_alvo

        nome_sugerido = self.caminho_arquivo.replace(".xlsx", "_Analise_Evolucao_PTA.xlsx").replace(".xlsb", "_Analise_Evolucao_PTA.xlsx")
        caminho_saida, _ = QFileDialog.getSaveFileName(self, "Salvar Relatório Como...", nome_sugerido, "Excel Files (*.xlsx)")
        if not caminho_saida: return
        
        opcoes['caminho_saida'] = caminho_saida

        # Extração das Regras Digitadas na Tela
        regras = {
            'ativo': [{'min': r['min'].text(), 'max': r['max'].text(), 'chave': r['chave'].text()} for r in self.regras_ativo_ui],
            'passivo': [{'min': r['min'].text(), 'max': r['max'].text(), 'chave': r['chave'].text()} for r in self.regras_passivo_ui],
            'resultado': [{'min': r['min'].text(), 'max': r['max'].text(), 'chave': r['chave'].text(), 'perc': r['perc'].text()} for r in self.regras_resultado_ui]
        }

        try:
            self.btn_processar.setEnabled(False)
            self.btn_reprocessar.setEnabled(False)
            self.barra_progresso.setValue(30)
            QApplication.processEvents()
            
            arquivo_gerado, log_erros = executar_analise_evolucao(self.caminho_arquivo, opcoes, regras)
            self.ultimo_pta_gerado = arquivo_gerado
            
            self.barra_progresso.setValue(100)
            msg = f"Processo concluído com sucesso!\n\nSalvo em:\n{os.path.basename(arquivo_gerado)}"
            QMessageBox.information(self, "Sucesso", msg)
            
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Ocorreu um erro durante o processamento:\n{str(e)}")
        finally:
            self.btn_processar.setEnabled(True)
            self.btn_reprocessar.setEnabled(True)
            self.barra_progresso.setValue(0)

def main():
    janela = AnaliseEvolucaoApp()
    janela.show()
    return None, janela