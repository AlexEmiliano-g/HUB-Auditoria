import os
from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QLabel, QPushButton, 
                             QFileDialog, QCheckBox, QGroupBox, QHBoxLayout, 
                             QMessageBox, QProgressBar, QScrollArea, QGridLayout, QLineEdit, QApplication, QDialog, QTextBrowser)
from PyQt6.QtCore import Qt

# Importando as funções do Motor Lógico
from scripts.Contabilidade.analise_evolucao_logic import executar_analise_evolucao, gerar_ptas_excel

APP_STYLESHEET = """
QWidget#MainWindow { background-color: #F1F3F4; }
QWidget#ScrollContent { background-color: #FFFFFF; }
QWidget#Footer { background-color: #FFFFFF; border-top: 1px solid #DADCE0; }
QLabel#AppTitle { font-size: 22px; font-weight: 600; color: #1A73E8; padding: 5px 0; }
QLabel#SectionTitle { font-size: 14px; font-weight: bold; color: #202124; margin-top: 15px; margin-bottom: 5px; }
QLabel#GridHeader { font-size: 12px; font-weight: bold; color: #5F6368; padding-bottom: 5px; }
QLabel { font-size: 12px; color: #3C4043; }

QGroupBox { background-color: #FFFFFF; border: 1px solid #E1E4E8; border-radius: 8px; font-size: 13px; font-weight: 600; color: #202124; margin-top: 15px; padding-top: 15px; }
QGroupBox::title { subcontrol-origin: margin; subcontrol-position: top left; padding: 0 8px; left: 15px; }

/* Botões Principais no Rodapé */
QPushButton#btn_processar { background-color: #1A73E8; color: white; font-size: 13px; font-weight: 600; border: none; padding: 10px 15px; border-radius: 6px; }
QPushButton#btn_processar:hover { background-color: #1557B0; }
QPushButton#btn_reprocessar { background-color: #34A853; color: white; font-size: 13px; font-weight: 600; border: none; padding: 10px 15px; border-radius: 6px; }
QPushButton#btn_reprocessar:hover { background-color: #2b8c44; }
QPushButton#btn_pta { background-color: #FBBC04; color: #202124; font-size: 13px; font-weight: bold; border: none; padding: 10px 15px; border-radius: 6px; }
QPushButton#btn_pta:hover { background-color: #F2A500; }
QPushButton:disabled { background-color: #DADCE0; color: #80868B; }

/* Botão Secundário */
QPushButton#btn_secundario { background-color: #F8F9FA; color: #3C4043; font-weight: bold; border: 1px solid #DADCE0; padding: 7px 15px; border-radius: 4px; }
QPushButton#btn_secundario:hover { background-color: #E8EAED; }

/* Botão Ajuda */
QPushButton#btn_help { 
    background-color: transparent; border: 2px solid #1A73E8; 
    color: #1A73E8; font-size: 16px; font-weight: bold; 
    border-radius: 15px; min-width: 30px; min-height: 30px; max-width: 30px; max-height: 30px; 
}
QPushButton#btn_help:hover { background-color: #E8F0FE; }

QScrollArea { border: none; background-color: transparent; }

QLineEdit { background-color: #F8F9FA; border: 1px solid #DADCE0; border-radius: 4px; padding: 6px; color: #202124; font-weight: bold; }
QLineEdit:focus { border: 1px solid #1A73E8; background-color: #FFFFFF; }
QLineEdit[readOnly="true"] { background-color: #F1F3F4; color: #5F6368; font-weight: normal; border: 1px solid #E8EAED; }

QProgressBar { border: 1px solid #DADCE0; border-radius: 3px; background-color: #F1F3F4; text-align: center; }
QProgressBar::chunk { background-color: #1A73E8; border-radius: 3px; }
QProgressBar[state="success"]::chunk { background-color: #34A853; }
QProgressBar[state="pta"]::chunk { background-color: #FBBC04; }
"""

class AjudaDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Wiki: Análise de Evolução")
        self.resize(700, 550)
        self.setStyleSheet(APP_STYLESHEET)
        
        lay = QVBoxLayout(self)
        lay.setContentsMargins(25, 25, 25, 25)
        lay.setSpacing(15)
        
        lbl_title = QLabel("Documentação da Ferramenta", objectName="AppTitle")
        lay.addWidget(lbl_title)
        
        txt_wiki = QTextBrowser()
        txt_wiki.setHtml("""
        <div style='font-family: Calibri, sans-serif; font-size: 14px; color: #3C4043; line-height: 1.5;'>
            <h3 style='color: #1A73E8;'>Visão Geral</h3>
            <p>O sistema consolida balancetes, gera o <b>Balancete Histórico</b> auditável e constrói planilhas <b>PTA</b> separadas aplicando inteligência sobre saldos e desvios de média (NBC TA 530).</p>
            <h3 style='color: #1A73E8;'>Novo Fluxo de Auditoria</h3>
            <ol>
                <li>Processe a Base ou Atualize o Balancete Histórico.</li>
                <li>Abra o Excel gerado. As contas amostradas estarão pintadas na aba Balancete Histórico.</li>
                <li><b>Julgamento Profissional:</b> Se desejar remover uma conta da malha, apenas delete a tag "X-R" da Coluna A (Seleção) e salve o arquivo.</li>
                <li>Volte no aplicativo e clique no botão Amarelo <b>(Gerar PTAs)</b>. O sistema lerá suas edições e criará um <b>novo arquivo</b> apenas com os Papéis de Trabalho e Justificativas.</li>
            </ol>
        </div>
        """)
        txt_wiki.setStyleSheet("border: 1px solid #DADCE0; border-radius: 6px; background-color: #F8F9FA;")
        lay.addWidget(txt_wiki)
        
        lay_botoes = QHBoxLayout()
        btn_fechar = QPushButton("Entendi")
        btn_fechar.setObjectName("btn_secundario")
        btn_fechar.clicked.connect(self.accept)
        lay_botoes.addStretch()
        lay_botoes.addWidget(btn_fechar)
        lay.addLayout(lay_botoes)

class AnaliseEvolucaoApp(QWidget):
    def __init__(self):
        super().__init__()
        self.caminho_arquivo = ""
        self.ultimo_pta_gerado = ""
        self.initUI()

    def initUI(self):
        self.setObjectName("MainWindow")
        self.setStyleSheet(APP_STYLESHEET)
        
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # --- ÁREA DE ROLAGEM ---
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        content = QWidget()
        content.setObjectName("ScrollContent")
        layout = QVBoxLayout(content)
        layout.setContentsMargins(30, 20, 30, 30)
        layout.setSpacing(20)
        
        lay_header = QHBoxLayout()
        lbl_title = QLabel("Análise de Evolução", objectName="AppTitle")
        btn_ajuda = QPushButton("?")
        btn_ajuda.setObjectName("btn_help")
        btn_ajuda.setToolTip("Abrir documentação e wiki da ferramenta")
        btn_ajuda.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_ajuda.clicked.connect(self.abrir_wiki)
        
        lay_header.addWidget(lbl_title)
        lay_header.addStretch()
        lay_header.addWidget(btn_ajuda)
        layout.addLayout(lay_header)

        g_arq = QGroupBox("Seleção de Arquivo (Master Data)")
        lay_arq = QVBoxLayout()
        lay_arq.setContentsMargins(20, 20, 20, 20)
        
        lay_file = QHBoxLayout()
        self.txt_arquivo = QLineEdit()
        self.txt_arquivo.setReadOnly(True)
        self.txt_arquivo.setPlaceholderText("Selecione o arquivo Excel contendo os balancetes mensais...")
        
        btn_arq = QPushButton("Procurar...")
        btn_arq.setObjectName("btn_secundario")
        btn_arq.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_arq.clicked.connect(self.selecionar_arquivo)
        
        lay_file.addWidget(self.txt_arquivo, stretch=1)
        lay_file.addWidget(btn_arq)
        
        self.chk_inclusao = QCheckBox("Habilitar Inclusão e Classificação Automática de Contas Órfãs")
        self.chk_inclusao.setChecked(True)
        self.chk_inclusao.setStyleSheet("font-weight: normal; color: #3C4043;")
        
        lay_arq.addLayout(lay_file)
        lay_arq.addWidget(self.chk_inclusao)
        g_arq.setLayout(lay_arq)
        layout.addWidget(g_arq)

        g_escopo = QGroupBox("Parametrização de Risco (Limites de Amostragem)")
        lay_escopo = QVBoxLayout()
        lay_escopo.setContentsMargins(20, 20, 20, 20)
        
        lay_check = QHBoxLayout()
        self.chk_ativo = QCheckBox("PTA - Ativo"); self.chk_ativo.setChecked(True)
        self.chk_passivo = QCheckBox("PTA - Passivo"); self.chk_passivo.setChecked(True)
        self.chk_resultado = QCheckBox("PTA - Resultado"); self.chk_resultado.setChecked(True)
        lay_check.addWidget(self.chk_ativo); lay_check.addWidget(self.chk_passivo); lay_check.addWidget(self.chk_resultado)
        lay_check.addStretch()
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
        
        scroll.setWidget(content)
        main_layout.addWidget(scroll, stretch=1)

        # --- RODAPÉ FIXO ---
        footer = QWidget()
        footer.setObjectName("Footer")
        lay_footer = QHBoxLayout(footer)
        lay_footer.setContentsMargins(30, 15, 30, 15)
        lay_footer.setSpacing(15)

        lay_status = QVBoxLayout()
        lay_status.setSpacing(5)
        self.lbl_status = QLabel("Aguardando processamento...")
        self.lbl_status.setStyleSheet("color: #5F6368; font-size: 11px;")
        
        self.barra_progresso = QProgressBar()
        self.barra_progresso.setFixedHeight(6)
        self.barra_progresso.setTextVisible(False)
        self.barra_progresso.setFixedWidth(200)
        self.barra_progresso.setProperty("state", "normal")
        
        lay_status.addWidget(self.lbl_status)
        lay_status.addWidget(self.barra_progresso)
        
        self.btn_processar = QPushButton("Processar Base Original")
        self.btn_processar.setObjectName("btn_processar")
        self.btn_processar.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_processar.clicked.connect(lambda: self.processar_dados(is_reprocess=False))
        
        self.btn_reprocessar = QPushButton("Atualizar Balancete Histórico")
        self.btn_reprocessar.setObjectName("btn_reprocessar")
        self.btn_reprocessar.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_reprocessar.clicked.connect(lambda: self.processar_dados(is_reprocess=True))
        
        self.btn_pta = QPushButton("Gerar PTAs")
        self.btn_pta.setObjectName("btn_pta")
        self.btn_pta.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_pta.setToolTip("Cria um novo arquivo com as abas de Papéis de Trabalho e Justificativas.")
        self.btn_pta.clicked.connect(self.disparar_geracao_pta)

        lay_footer.addLayout(lay_status)
        lay_footer.addStretch()
        lay_footer.addWidget(self.btn_processar)
        lay_footer.addWidget(self.btn_reprocessar)
        lay_footer.addWidget(self.btn_pta)

        main_layout.addWidget(footer)
        self.resize(1050, 750)
        self.setWindowTitle("Análise de Evolução - Auditoria")

    def abrir_wiki(self):
        dlg = AjudaDialog(self)
        dlg.exec()

    def criar_formulario_basico(self, titulo, nome_chave):
        container = QWidget(); lay = QVBoxLayout(container); lay.setContentsMargins(0, 10, 0, 10)
        lbl = QLabel(titulo); lbl.setObjectName("SectionTitle"); lay.addWidget(lbl)
        grid = QGridLayout(); grid.setSpacing(10)
        headers = ["Faixa", "Saldo Mín. >", "Saldo Máx. Até", nome_chave]
        for col, text in enumerate(headers):
            lbl_h = QLabel(text); lbl_h.setObjectName("GridHeader"); lbl_h.setAlignment(Qt.AlignmentFlag.AlignCenter); grid.addWidget(lbl_h, 0, col)
            
        mins = ["0,10%", "0,30%", "0,60%", "0,90%", "1,20%", "1,50%", "1,80%", "2,00%"]
        maxs = ["0,30%", "0,60%", "0,90%", "1,20%", "1,50%", "1,80%", "2,00%", ""]
        lista_inputs = []
        for i in range(8):
            lbl_faixa = QLabel(f"{i+1}"); lbl_faixa.setAlignment(Qt.AlignmentFlag.AlignCenter); lbl_faixa.setStyleSheet("font-weight: bold; color: #5F6368;")
            inp_min = QLineEdit(mins[i]); inp_min.setAlignment(Qt.AlignmentFlag.AlignCenter)
            inp_max = QLineEdit(maxs[i]); inp_max.setAlignment(Qt.AlignmentFlag.AlignCenter)
            inp_chave = QLineEdit(""); inp_chave.setAlignment(Qt.AlignmentFlag.AlignCenter)
            grid.addWidget(lbl_faixa, i+1, 0); grid.addWidget(inp_min, i+1, 1); grid.addWidget(inp_max, i+1, 2); grid.addWidget(inp_chave, i+1, 3)
            lista_inputs.append({'min': inp_min, 'max': inp_max, 'chave': inp_chave})
            
        grid.setColumnStretch(1, 1); grid.setColumnStretch(2, 1); grid.setColumnStretch(3, 3)
        lay.addLayout(grid)
        return container, lista_inputs

    def criar_formulario_composto(self, titulo):
        container = QWidget(); lay = QVBoxLayout(container); lay.setContentsMargins(0, 10, 0, 10)
        lbl = QLabel(titulo); lbl.setObjectName("SectionTitle"); lay.addWidget(lbl)
        grid = QGridLayout(); grid.setSpacing(10)
        headers = ["Faixa", "Saldo Mín. >", "Saldo Máx. Até", "Chave D&M", "Movimento(mês) >= % Média"]
        for col, text in enumerate(headers):
            lbl_h = QLabel(text); lbl_h.setObjectName("GridHeader"); lbl_h.setAlignment(Qt.AlignmentFlag.AlignCenter); grid.addWidget(lbl_h, 0, col)
            
        mins = ["0,20%", "0,40%", "0,70%", "1,00%", "1,30%", "1,60%", "1,90%", "2,10%"]
        maxs = ["0,40%", "0,70%", "1,00%", "1,30%", "1,60%", "1,90%", "2,10%", ""]
        percs = ["100%", "75%", "50%", "25%", "20%", "15%", "10%", "5%"]
        lista_inputs = []
        for i in range(8):
            lbl_faixa = QLabel(f"{i+1}"); lbl_faixa.setAlignment(Qt.AlignmentFlag.AlignCenter); lbl_faixa.setStyleSheet("font-weight: bold; color: #5F6368;")
            inp_min = QLineEdit(mins[i]); inp_min.setAlignment(Qt.AlignmentFlag.AlignCenter)
            inp_max = QLineEdit(maxs[i]); inp_max.setAlignment(Qt.AlignmentFlag.AlignCenter)
            inp_chave = QLineEdit("30101"); inp_chave.setAlignment(Qt.AlignmentFlag.AlignCenter)
            inp_perc = QLineEdit(percs[i]); inp_perc.setAlignment(Qt.AlignmentFlag.AlignCenter)
            grid.addWidget(lbl_faixa, i+1, 0); grid.addWidget(inp_min, i+1, 1); grid.addWidget(inp_max, i+1, 2); grid.addWidget(inp_chave, i+1, 3); grid.addWidget(inp_perc, i+1, 4)
            lista_inputs.append({'min': inp_min, 'max': inp_max, 'chave': inp_chave, 'perc': inp_perc})
            
        grid.setColumnStretch(1, 1); grid.setColumnStretch(2, 1); grid.setColumnStretch(3, 2); grid.setColumnStretch(4, 1)
        lay.addLayout(grid)
        return container, lista_inputs

    def selecionar_arquivo(self):
        caminho, _ = QFileDialog.getOpenFileName(self, "Selecionar Arquivo", "", "Excel Files (*.xlsx *.xlsb)")
        if caminho:
            self.caminho_arquivo = caminho
            self.txt_arquivo.setText(caminho)
            self.lbl_status.setText("Pronto para processar.")

    def set_estado_barra(self, estado):
        self.barra_progresso.setProperty("state", estado)
        self.barra_progresso.style().unpolish(self.barra_progresso)
        self.barra_progresso.style().polish(self.barra_progresso)

    def processar_dados(self, is_reprocess=False):
        if not self.caminho_arquivo:
            QMessageBox.warning(self, "Aviso", "Selecione o arquivo Excel Original primeiro.")
            return

        opcoes = {'ativo': self.chk_ativo.isChecked(), 'passivo': self.chk_passivo.isChecked(), 'resultado': self.chk_resultado.isChecked(), 'inclusao_inteligente': self.chk_inclusao.isChecked(), 'is_reprocess': is_reprocess}

        if is_reprocess:
            caminho_pta_alvo = ""
            if self.ultimo_pta_gerado and os.path.exists(self.ultimo_pta_gerado):
                resp = QMessageBox.question(self, "Atualizar", f"Deseja reprocessar o último arquivo gerado?\n\n{os.path.basename(self.ultimo_pta_gerado)}", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                if resp == QMessageBox.StandardButton.Yes: caminho_pta_alvo = self.ultimo_pta_gerado
            if not caminho_pta_alvo:
                caminho_pta_alvo, _ = QFileDialog.getOpenFileName(self, "Selecionar Balancete Histórico", "", "Excel Files (*.xlsx *.xlsb)")
                if not caminho_pta_alvo: return
            opcoes['caminho_pta_reprocess'] = caminho_pta_alvo

        nome_sugerido = self.caminho_arquivo.replace(".xlsx", "_Balancete_Historico.xlsx").replace(".xlsb", "_Balancete_Historico.xlsx")
        caminho_saida, _ = QFileDialog.getSaveFileName(self, "Salvar Relatório Como...", nome_sugerido, "Excel Files (*.xlsx)")
        if not caminho_saida: return
        
        opcoes['caminho_saida'] = caminho_saida
        regras = {
            'ativo': [{'min': r['min'].text(), 'max': r['max'].text(), 'chave': r['chave'].text()} for r in self.regras_ativo_ui],
            'passivo': [{'min': r['min'].text(), 'max': r['max'].text(), 'chave': r['chave'].text()} for r in self.regras_passivo_ui],
            'resultado': [{'min': r['min'].text(), 'max': r['max'].text(), 'chave': r['chave'].text(), 'perc': r['perc'].text()} for r in self.regras_resultado_ui]
        }

        try:
            self.lbl_status.setText("Lendo bases e calculando amostragem...")
            self.btn_processar.setEnabled(False); self.btn_reprocessar.setEnabled(False); self.btn_pta.setEnabled(False)
            self.set_estado_barra("normal"); self.barra_progresso.setValue(45); QApplication.processEvents() 
            
            arquivo_gerado, log_erros = executar_analise_evolucao(self.caminho_arquivo, opcoes, regras)
            self.ultimo_pta_gerado = arquivo_gerado
            
            self.set_estado_barra("success"); self.barra_progresso.setValue(100)
            self.lbl_status.setText("Processo finalizado com sucesso!")
            msg = f"Balancete Histórico gerado com sucesso!\n\nSalvo em: {os.path.basename(arquivo_gerado)}\n\nAbra o arquivo, revise a seleção na Coluna A e depois use o botão amarelo para gerar os PTAs."
            QMessageBox.information(self, "Sucesso", msg) 
        except Exception as e:
            self.lbl_status.setText("Erro crítico."); QMessageBox.critical(self, "Erro", f"Ocorreu um erro:\n{str(e)}")
        finally:
            self.set_estado_barra("normal"); self.barra_progresso.setValue(0); self.lbl_status.setText("Aguardando ação...")
            self.btn_processar.setEnabled(True); self.btn_reprocessar.setEnabled(True); self.btn_pta.setEnabled(True)

    def disparar_geracao_pta(self):
        caminho_origem = ""
        if self.ultimo_pta_gerado and os.path.exists(self.ultimo_pta_gerado):
            resp = QMessageBox.question(self, "Origem dos Dados", f"Deseja extrair as seleções do último balancete processado?\n\n{os.path.basename(self.ultimo_pta_gerado)}", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if resp == QMessageBox.StandardButton.Yes: caminho_origem = self.ultimo_pta_gerado
        
        if not caminho_origem:
            caminho_origem, _ = QFileDialog.getOpenFileName(self, "Selecione o arquivo Balancete Histórico", "", "Excel Files (*.xlsx)")
            if not caminho_origem: return
            
        nome_sugerido_pta = caminho_origem.replace("_Balancete_Historico", "").replace(".xlsx", "_PTA.xlsx").replace(".xlsb", "_PTA.xlsx")
        caminho_destino, _ = QFileDialog.getSaveFileName(self, "Salvar Novo Arquivo de PTA Como...", nome_sugerido_pta, "Excel Files (*.xlsx)")
        if not caminho_destino: return
        
        try:
            self.lbl_status.setText("Montando estrutura NBC TA 530 em arquivo separado...")
            self.btn_processar.setEnabled(False); self.btn_reprocessar.setEnabled(False); self.btn_pta.setEnabled(False)
            self.set_estado_barra("pta"); self.barra_progresso.setValue(50); QApplication.processEvents()
            
            gerar_ptas_excel(caminho_origem, caminho_destino)
            
            self.barra_progresso.setValue(100)
            self.lbl_status.setText("Papéis de Trabalho gerados com sucesso!")
            QMessageBox.information(self, "Sucesso", f"O arquivo de PTA foi gerado com sucesso em:\n\n{os.path.basename(caminho_destino)}\n\nO balancete histórico original foi mantido intacto.")
        except Exception as e:
            self.lbl_status.setText("Erro na geração."); QMessageBox.critical(self, "Erro", f"Falha ao gerar PTA:\n{str(e)}")
        finally:
            self.set_estado_barra("normal"); self.barra_progresso.setValue(0); self.lbl_status.setText("Aguardando ação...")
            self.btn_processar.setEnabled(True); self.btn_reprocessar.setEnabled(True); self.btn_pta.setEnabled(True)

def main():
    janela = AnaliseEvolucaoApp()
    janela.show()
    return None, janela