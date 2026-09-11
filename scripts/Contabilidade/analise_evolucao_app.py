import os
import sys
import tempfile
import math
import pandas as pd
from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QLabel, QPushButton, 
                             QFileDialog, QCheckBox, QGroupBox, QHBoxLayout, 
                             QMessageBox, QProgressBar, QScrollArea, QGridLayout, QLineEdit, QApplication, QDialog, QTextBrowser)
from PyQt6.QtCore import Qt, QTimer, QUrl
from PyQt6.QtGui import QDesktopServices

# Importando as funções do Motor Lógico
from scripts.Contabilidade.analise_evolucao_logic import executar_analise_evolucao, gerar_ptas_excel

APP_STYLESHEET = """
/* FORÇA MODO CLARO GLOBAL */
QWidget { color: #202124; font-family: "Segoe UI", Arial, sans-serif; }
QDialog, QMessageBox { background-color: #FFFFFF; }

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
QPushButton#btn_processar { background-color: #34A853; color: white; font-size: 13px; font-weight: 600; border: none; padding: 10px 15px; border-radius: 6px; }
QPushButton#btn_processar:hover { background-color: #2b8c44; }

QPushButton#btn_pta { background-color: #FBBC04; color: #202124; font-size: 13px; font-weight: bold; border: none; padding: 10px 15px; border-radius: 6px; }
QPushButton#btn_pta:hover { background-color: #F2A500; }

QPushButton#btn_reprocessar { background-color: #F8F9FA; color: #3C4043; font-size: 13px; font-weight: bold; border: 1px solid #DADCE0; padding: 10px 15px; border-radius: 6px; }
QPushButton#btn_reprocessar:hover { background-color: #E8EAED; }

QPushButton:disabled { background-color: #DADCE0; color: #80868B; border: none; }

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

QLineEdit { background-color: #FFFFFF; border: 1px solid #DADCE0; border-radius: 4px; padding: 6px; color: #202124; font-weight: bold; }
QLineEdit:focus { border: 1px solid #1A73E8; background-color: #FFFFFF; }
QLineEdit[readOnly="true"] { background-color: #F1F3F4; color: #5F6368; font-weight: normal; border: 1px solid #E8EAED; }

QProgressBar { border: 1px solid #DADCE0; border-radius: 3px; background-color: #F1F3F4; text-align: center; color: #202124; }
QProgressBar::chunk { background-color: #1A73E8; border-radius: 3px; }
QProgressBar[state="success"]::chunk { background-color: #34A853; }
QProgressBar[state="pta"]::chunk { background-color: #FBBC04; }
"""

class AjudaDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Documentação: Análise de Evolução")
        self.resize(780, 550)
        self.setStyleSheet(APP_STYLESHEET)
        
        lay = QVBoxLayout(self)
        lay.setContentsMargins(25, 25, 25, 25)
        lay.setSpacing(15)
        
        lbl_title = QLabel("Manual de Operação", objectName="AppTitle")
        lay.addWidget(lbl_title)
        
        txt_wiki = QTextBrowser()
        txt_wiki.setOpenLinks(False)
        txt_wiki.anchorClicked.connect(self.tratar_link_clicado)
        
        txt_wiki.setHtml("""
        <div style='font-family: "Segoe UI", Calibri, sans-serif; font-size: 14px; color: #3C4043; line-height: 1.6;'>
            
            <h3 style='color: #1A73E8; margin-top: 0;'>Visão Geral</h3>
            <p>Este módulo consolida exportações contábeis, gera o <b>Balancete Histórico</b> auditável e constrói planilhas <b>PTA</b> de forma segmentada, aplicando diretrizes de amostragem e análise de desvios (NBC TA 530).</p>
            
            <h3 style='color: #1A73E8;'>Pré-requisitos e Parâmetros</h3>
            <p>Para o processamento correto, o arquivo de entrada deve conter os balancetes tabulados no formato padrão. Se desejar, utilize o botão abaixo para abrir um modelo em Excel.</p>
            
            <p><b>Nota sobre o Plano de Contas:</b> Caso o seu arquivo original não contenha uma aba de Plano de Contas (ou Parâmetros), o sistema analisará os balancetes e gerará o plano de contas automaticamente para você.</p>
            
            <div style='background-color: #F8F9FA; padding: 12px; border-left: 4px solid #5F6368; margin: 15px 0;'>
                <b>Atenção:</b> O programa não insere a Chave D&M para seleção de contas automaticamente. Esta classificação deve estar presente no arquivo de entrada ou ser inserida manualmente após a geração do balancete histórico.
            </div>

            <h3 style='color: #1A73E8;'>Fluxo de Trabalho</h3>
            <ol>
                <li style='margin-bottom: 8px;'><b>Processamento Inicial:</b> Selecione o balancete tabulado na interface principal e execute o processamento para compilar o Balancete Histórico.</li>
                <li style='margin-bottom: 8px;'><b>Revisão de Amostragem:</b> O sistema abrirá o arquivo Excel gerado. As contas selecionadas pela avaliação de risco estarão destacadas.
                    <ul>
                        <li>Para <b>remover</b> uma conta da amostra, apague a codificação de seleção (ex: "X-R1") localizada na Coluna A.</li>
                        <li>Para <b>incluir</b> uma conta mediante julgamento profissional, insira os códigos <b>X-RA</b> (Ativo), <b>X-RP</b> (Passivo) ou <b>X-RR</b> (Resultado) na Coluna A da referida conta.</li>
                    </ul>
                </li>
                <li style='margin-bottom: 8px;'><b>Geração de Relatórios:</b> Após salvar as edições, retorne ao aplicativo e selecione <b>"Gerar PTAs"</b>. O sistema processará as alterações e emitirá os Papéis de Trabalho e as justificativas em um novo arquivo.</li>
            </ol>
        </div>
        """)
        txt_wiki.setStyleSheet("border: 1px solid #DADCE0; border-radius: 6px; background-color: #FFFFFF;")
        lay.addWidget(txt_wiki)
        
        lay_botoes = QHBoxLayout()
        
        btn_modelo = QPushButton("Abrir Modelo em Excel")
        btn_modelo.setObjectName("btn_secundario")
        btn_modelo.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_modelo.setStyleSheet("color: #1A73E8; font-weight: bold;")
        btn_modelo.clicked.connect(self.gerar_excel_de_exemplo)
        
        btn_fechar = QPushButton("Entendi")
        btn_fechar.setObjectName("btn_processar")
        btn_fechar.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_fechar.clicked.connect(self.accept)
        
        lay_botoes.addWidget(btn_modelo)
        lay_botoes.addStretch()
        lay_botoes.addWidget(btn_fechar)
        lay.addLayout(lay_botoes)

    def tratar_link_clicado(self, url: QUrl):
        if url.scheme() == "comando" and url.host() == "gerar_exemplo":
            self.gerar_excel_de_exemplo()
        else:
            QDesktopServices.openUrl(url)

    def gerar_excel_de_exemplo(self):
        try:
            temp_dir = tempfile.gettempdir()
            base_nome = "Modelo_Balancete_Tabulado"
            extensao = ".xlsx"
            caminho_salvar = os.path.join(temp_dir, base_nome + extensao)
            
            counter = 2
            while os.path.exists(caminho_salvar):
                try:
                    with open(caminho_salvar, 'r+'): 
                        break 
                except PermissionError:
                    caminho_salvar = os.path.join(temp_dir, f"{base_nome}_v{counter}{extensao}")
                    counter += 1

            dados_plano = [
                ["1", "", "1", "ATIVO", "S", "A", "1", ""],
                ["11", "", "2", "CIRCULANTE", "S", "A", "2", ""],
                ["111", "", "3", "DISPONIBILIDADES", "S", "A", "3", ""],
                ["11101", "", "4", "CAIXA", "A", "A", "4", ""],
                ["11102", "30101", "5", "BANCOS C/ MOVIMENTO", "A", "A", "5", ""]
            ]
            df_plano = pd.DataFrame(dados_plano, columns=['Chave Cliente', 'Chave D&M', 'Classificação', 'Descrição', 'Sint./An.', 'At/Pas/Res', 'Indice', 'Observação'])

            colunas_mes = ['Atividade', 'Conta', 'Descrição', 'Cod. Reduzido', 'Saldo Anterior', 'Débito', 'Crédito', 'Saldo Acumulado']
            
            df_m1 = pd.DataFrame([
                ["Geral", "1", "ATIVO", "1", 1000.00, 500.00, 200.00, 1300.00],
                ["Geral", "11", "CIRCULANTE", "2", 1000.00, 500.00, 200.00, 1300.00],
                ["Geral", "111", "DISPONIBILIDADES", "3", 1000.00, 500.00, 200.00, 1300.00],
                ["Geral", "11101", "CAIXA", "4", 400.00, 200.00, 100.00, 500.00],
                ["Geral", "11102", "BANCOS C/ MOVIMENTO", "5", 600.00, 300.00, 100.00, 800.00]
            ], columns=colunas_mes)
            
            df_m2 = pd.DataFrame([
                ["Geral", "1", "ATIVO", "1", 1300.00, 100.00, 0.00, 1400.00],
                ["Geral", "11", "CIRCULANTE", "2", 1300.00, 100.00, 0.00, 1400.00],
                ["Geral", "111", "DISPONIBILIDADES", "3", 1300.00, 100.00, 0.00, 1400.00],
                ["Geral", "11101", "CAIXA", "4", 500.00, 50.00, 0.00, 550.00],
                ["Geral", "11102", "BANCOS C/ MOVIMENTO", "5", 800.00, 50.00, 0.00, 850.00]
            ], columns=colunas_mes)
            
            df_m3 = pd.DataFrame([
                ["Geral", "1", "ATIVO", "1", 1400.00, 0.00, 400.00, 1000.00],
                ["Geral", "11", "CIRCULANTE", "2", 1400.00, 0.00, 400.00, 1000.00],
                ["Geral", "111", "DISPONIBILIDADES", "3", 1400.00, 0.00, 400.00, 1000.00],
                ["Geral", "11101", "CAIXA", "4", 550.00, 0.00, 150.00, 400.00],
                ["Geral", "11102", "BANCOS C/ MOVIMENTO", "5", 850.00, 0.00, 250.00, 600.00]
            ], columns=colunas_mes)

            with pd.ExcelWriter(caminho_salvar, engine='openpyxl') as writer:
                df_plano.to_excel(writer, index=False, sheet_name="Plano de Contas")
                df_m1.to_excel(writer, index=False, sheet_name="01")
                df_m2.to_excel(writer, index=False, sheet_name="02")
                df_m3.to_excel(writer, index=False, sheet_name="03")
                
            os.startfile(caminho_salvar)
            
        except Exception as e:
            QMessageBox.critical(self, "Erro na Geração", f"Não foi possível abrir o arquivo de exemplo. Detalhe: {e}")

class AnaliseEvolucaoApp(QWidget):
    def __init__(self):
        super().__init__()
        self.caminho_arquivo = ""
        self.caminho_arquivo_anterior = ""
        self.ultimo_pta_gerado = ""
        
        self.color_step = 0
        self.anim_timer = QTimer(self)
        self.anim_timer.timeout.connect(self.animate_checkbox_color)
        
        self.initUI()
        
        QTimer.singleShot(500, self.abrir_wiki)

    def initUI(self):
        self.setObjectName("MainWindow")
        self.setStyleSheet(APP_STYLESHEET)
        
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
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

        g_arq = QGroupBox("Seleção de Arquivos")
        lay_arq = QVBoxLayout()
        lay_arq.setContentsMargins(20, 20, 20, 20)
        lay_arq.setSpacing(12)
        
        lbl_atual = QLabel("Balancete Tabulado - Ano Atual:")
        lbl_atual.setStyleSheet("font-weight: bold; color: #5F6368;")
        
        lay_file_atual = QHBoxLayout()
        self.txt_arquivo = QLineEdit()
        self.txt_arquivo.setReadOnly(True)
        self.txt_arquivo.setPlaceholderText("Selecione o arquivo Excel contendo os balancetes mensais do ano vigente...")
        
        btn_arq = QPushButton("Procurar...")
        btn_arq.setObjectName("btn_secundario")
        btn_arq.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_arq.clicked.connect(self.selecionar_arquivo)
        
        lay_file_atual.addWidget(self.txt_arquivo, stretch=1)
        lay_file_atual.addWidget(btn_arq)
        
        lbl_anterior = QLabel("Balancete Tabulado - Ano Anterior (Opcional):")
        lbl_anterior.setStyleSheet("font-weight: bold; color: #5F6368; padding-top: 10px;")
        
        lay_file_anterior = QHBoxLayout()
        self.txt_arquivo_anterior = QLineEdit()
        self.txt_arquivo_anterior.setReadOnly(True)
        self.txt_arquivo_anterior.setPlaceholderText("Selecione o balancete tabulado do ano anterior (Para preenchimento automático das comparações)...")
        
        btn_arq_anterior = QPushButton("Procurar...")
        btn_arq_anterior.setObjectName("btn_secundario")
        btn_arq_anterior.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_arq_anterior.clicked.connect(self.selecionar_arquivo_anterior)
        
        lay_file_anterior.addWidget(self.txt_arquivo_anterior, stretch=1)
        lay_file_anterior.addWidget(btn_arq_anterior)
        
        self.chk_inclusao = QCheckBox("Habilitar Inclusão e Classificação Automática de Contas")
        self.chk_inclusao.setChecked(True)
        self.anim_timer.start(50)
        
        lay_arq.addWidget(lbl_atual)
        lay_arq.addLayout(lay_file_atual)
        lay_arq.addWidget(lbl_anterior)
        lay_arq.addLayout(lay_file_anterior)
        lay_arq.addWidget(self.chk_inclusao)
        
        g_arq.setLayout(lay_arq)
        layout.addWidget(g_arq)

        g_escopo = QGroupBox("Parametrização de Risco (Limites de Amostragem)")
        lay_escopo = QVBoxLayout()
        lay_escopo.setContentsMargins(20, 20, 20, 20)
        
        lay_check = QHBoxLayout()
        self.chk_ativo = QCheckBox("PTA - Ativo"); self.chk_ativo.setChecked(False)
        self.chk_passivo = QCheckBox("PTA - Passivo"); self.chk_passivo.setChecked(False)
        self.chk_resultado = QCheckBox("PTA - Resultado"); self.chk_resultado.setChecked(True)
        lay_check.addWidget(self.chk_ativo)
        lay_check.addWidget(self.chk_passivo)
        lay_check.addWidget(self.chk_resultado)
        lay_check.addStretch()
        lay_escopo.addLayout(lay_check)

        self.container_ativo, self.regras_ativo_ui = self.criar_formulario_basico("REGRAS ATIVO:", "Chave D&M")
        self.container_passivo, self.regras_passivo_ui = self.criar_formulario_basico("REGRAS PASSIVO:", "Chave Cliente")
        self.container_resultado, self.regras_resultado_ui = self.criar_formulario_composto("REGRAS RESULTADO:")
        
        self.container_ativo.setVisible(False)
        self.container_passivo.setVisible(False)

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

        footer = QWidget()
        footer.setObjectName("Footer")
        lay_footer = QHBoxLayout(footer)
        lay_footer.setContentsMargins(30, 15, 30, 15)

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

        lay_botoes = QHBoxLayout()
        lay_botoes.setSpacing(15) 
        
        self.btn_reprocessar = QPushButton("Atualizar Balancete Histórico")
        self.btn_reprocessar.setObjectName("btn_reprocessar")
        self.btn_reprocessar.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_reprocessar.clicked.connect(lambda: self.processar_dados(is_reprocess=True))
        
        self.btn_processar = QPushButton("1. Processar Base Original")
        self.btn_processar.setObjectName("btn_processar")
        self.btn_processar.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_processar.clicked.connect(lambda: self.processar_dados(is_reprocess=False))
        
        self.btn_pta = QPushButton("2. Gerar PTAs")
        self.btn_pta.setObjectName("btn_pta")
        self.btn_pta.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_pta.setToolTip("Cria um novo arquivo com as abas de Papéis de Trabalho e Justificativas.")
        self.btn_pta.clicked.connect(self.disparar_geracao_pta)

        lay_botoes.addWidget(self.btn_reprocessar)
        lay_botoes.addWidget(self.btn_processar)
        lay_botoes.addWidget(self.btn_pta)

        lay_footer.addLayout(lay_status)
        lay_footer.addStretch()
        lay_footer.addLayout(lay_botoes)

        main_layout.addWidget(footer)
        self.resize(1050, 750)
        self.setWindowTitle("Análise de Evolução - Auditoria")

    def animate_checkbox_color(self):
        self.color_step += 0.05
        wave = (math.sin(self.color_step) + 1) / 2.0
        r = int(95 + (26 - 95) * wave)
        g = int(99 + (115 - 99) * wave)
        b = int(104 + (232 - 104) * wave)
        color_hex = f"#{r:02X}{g:02X}{b:02X}"
        self.chk_inclusao.setStyleSheet(f"font-weight: 600; font-size: 13px; color: {color_hex}; padding-top: 10px;")

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
        caminho, _ = QFileDialog.getOpenFileName(self, "Selecionar Arquivo do Ano Atual", "", "Excel Files (*.xlsx *.xlsb)")
        if caminho:
            self.caminho_arquivo = caminho
            self.txt_arquivo.setText(caminho)
            self.lbl_status.setText("Pronto para processar.")

    def selecionar_arquivo_anterior(self):
        caminho, _ = QFileDialog.getOpenFileName(self, "Selecionar Arquivo do Ano Anterior", "", "Excel Files (*.xlsx *.xlsb)")
        if caminho:
            self.caminho_arquivo_anterior = caminho
            self.txt_arquivo_anterior.setText(caminho)

    def set_estado_barra(self, estado):
        self.barra_progresso.setProperty("state", estado)
        self.barra_progresso.style().unpolish(self.barra_progresso)
        self.barra_progresso.style().polish(self.barra_progresso)

    def processar_dados(self, is_reprocess=False):
        if not self.caminho_arquivo:
            QMessageBox.warning(self, "Aviso", "Selecione o arquivo Excel do Ano Atual primeiro.")
            return

        opcoes = {
            'ativo': self.chk_ativo.isChecked(), 
            'passivo': self.chk_passivo.isChecked(), 
            'resultado': self.chk_resultado.isChecked(), 
            'inclusao_inteligente': self.chk_inclusao.isChecked(), 
            'is_reprocess': is_reprocess,
            'caminho_anterior': self.caminho_arquivo_anterior
        }

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
            msg = f"Balancete Histórico gerado com sucesso!\n\nSalvo em: {os.path.basename(arquivo_gerado)}\n\nO arquivo será aberto agora para você revisar a seleção na Coluna A. Após revisar, salve, feche e clique no botão '2. Gerar PTAs'."
            QMessageBox.information(self, "Sucesso", msg) 
            
            os.startfile(arquivo_gerado)
            
        except Exception as e:
            self.lbl_status.setText("Erro crítico."); QMessageBox.critical(self, "Erro", f"Ocorreu um erro:\n{str(e)}")
        finally:
            self.set_estado_barra("normal"); self.barra_progresso.setValue(0); self.lbl_status.setText("Aguardando ação...")
            self.btn_processar.setEnabled(True); self.btn_reprocessar.setEnabled(True); self.btn_pta.setEnabled(True)

    def disparar_geracao_pta(self):
        caminho_origem = ""
        caminho_anterior_para_pta = None 
        
        # 1. Tenta usar o arquivo da sessão atual
        if self.ultimo_pta_gerado and os.path.exists(self.ultimo_pta_gerado):
            resp = QMessageBox.question(self, "Origem dos Dados", f"Deseja extrair as seleções do último balancete processado?\n\n{os.path.basename(self.ultimo_pta_gerado)}", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            
            if resp == QMessageBox.StandardButton.Yes: 
                caminho_origem = self.ultimo_pta_gerado
                # AQUI ESTÁ A CHAVE: Pegar diretamente do campo de texto da tela
                arq_ant = self.txt_arquivo_anterior.text().strip()
                if arq_ant and os.path.exists(arq_ant):
                    caminho_anterior_para_pta = arq_ant
        
        # 2. Se o usuário disse NÃO (ou não tem sessão), pede pra ele escolher o arquivo manualmente
        if not caminho_origem:
            caminho_origem, _ = QFileDialog.getOpenFileName(self, "Selecione o arquivo Balancete Histórico", "", "Excel Files (*.xlsx *.xlsb)")
            if not caminho_origem: return
            caminho_anterior_para_pta = None # Garante que NUNCA vai misturar os anos de clientes diferentes
            
        nome_sugerido_pta = caminho_origem.replace("_Balancete_Historico", "").replace(".xlsx", "_PTA.xlsx").replace(".xlsb", "_PTA.xlsx")
        caminho_destino, _ = QFileDialog.getSaveFileName(self, "Salvar Novo Arquivo de PTA Como...", nome_sugerido_pta, "Excel Files (*.xlsx)")
        if not caminho_destino: return
        
        try:
            self.lbl_status.setText("Montando estrutura NBC TA 530 em arquivo separado...")
            self.btn_processar.setEnabled(False); self.btn_reprocessar.setEnabled(False); self.btn_pta.setEnabled(False)
            self.set_estado_barra("pta"); self.barra_progresso.setValue(50); QApplication.processEvents()
            
            # Chama a função passando a terceira variável crucial
            gerar_ptas_excel(caminho_origem, caminho_destino, caminho_anterior_para_pta)
            
            self.barra_progresso.setValue(100)
            self.lbl_status.setText("Papéis de Trabalho gerados com sucesso!")
            QMessageBox.information(self, "Sucesso", f"O arquivo de PTA foi gerado com sucesso em:\n\n{os.path.basename(caminho_destino)}\n\nO arquivo será aberto agora.")
            
            os.startfile(caminho_destino)
            
        except Exception as e:
            self.lbl_status.setText("Erro na geração."); QMessageBox.critical(self, "Erro", f"Falha ao gerar PTA:\n{str(e)}")
        finally:
            self.set_estado_barra("normal"); self.barra_progresso.setValue(0); self.lbl_status.setText("Aguardando ação...")
            self.btn_processar.setEnabled(True); self.btn_reprocessar.setEnabled(True); self.btn_pta.setEnabled(True)

def main():
    app = QApplication.instance()
    if not app:
        app = QApplication(sys.argv)
    janela = AnaliseEvolucaoApp()
    janela.show()
    return None, janela