import sys
import os
import pandas as pd
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QListWidget, QAbstractItemView,
    QFileDialog, QMessageBox, QComboBox, QProgressBar,
    QGroupBox
)
from PyQt6.QtGui import QFont, QCursor
from PyQt6.QtCore import Qt

# Tratamento de caminhos para suportar a execução tanto via HUB quanto standalone
try:
    from scripts.Contabilidade.tabulador_logic import processar_arquivos_selecionados, obter_nomes_sistemas
except ModuleNotFoundError:
    from tabulador_logic import processar_arquivos_selecionados, obter_nomes_sistemas

# Carrega a lista de sistemas configurados no motor de lógica de forma dinâmica
LISTA_SISTEMAS = obter_nomes_sistemas()

class TabuladorPythonApp(QWidget):
    def __init__(self):
        super().__init__()
        self.arquivos_selecionados = []
        self.setObjectName("MainWindow") 
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Tabulador de Balancetes")
        self.setGeometry(300, 300, 650, 550)
        
        master_layout = QVBoxLayout(self)
        master_layout.setContentsMargins(0, 0, 0, 0)
        master_layout.setSpacing(0)

        content_widget = QWidget()
        content_widget.setObjectName("ScrollContent") 
        content_layout = QVBoxLayout(content_widget)
        content_layout.setContentsMargins(30, 20, 30, 20)
        content_layout.setSpacing(15)

        # --- CABEÇALHO ---
        header_layout = QHBoxLayout()
        app_title = QLabel("Tabulador de Balancetes")
        app_title.setObjectName("AppTitle")
        
        btn_help = QPushButton("?")
        btn_help.setObjectName("btn_help")
        btn_help.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        btn_help.setToolTip("Ajuda sobre o Tabulador")
        
        header_layout.addWidget(app_title)
        header_layout.addStretch()
        header_layout.addWidget(btn_help)
        content_layout.addLayout(header_layout)

        # --- ETAPA 1 ---
        group_step1 = QGroupBox("1. Seleção de Arquivo")
        layout_step1 = QVBoxLayout()
        layout_step1.setContentsMargins(15, 25, 15, 15)
        
        self.file_list_widget = QListWidget()
        self.file_list_widget.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        layout_step1.addWidget(self.file_list_widget)

        browse_layout = QHBoxLayout()
        browse_layout.addStretch()
        
        self.browse_button = QPushButton("Procurar Arquivos...")
        self.browse_button.setObjectName("btn_secundario")
        self.browse_button.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.browse_button.clicked.connect(self.browse_files)
        
        browse_layout.addWidget(self.browse_button)
        layout_step1.addLayout(browse_layout)
        
        group_step1.setLayout(layout_step1)
        content_layout.addWidget(group_step1)

        # --- ETAPA 2 ---
        group_step2 = QGroupBox("2. Selecione o Sistema Contábil")
        layout_step2 = QVBoxLayout()
        layout_step2.setContentsMargins(15, 25, 15, 15)
        
        self.system_combo = QComboBox()
        self.system_combo.addItems(LISTA_SISTEMAS)
        self.system_combo.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        layout_step2.addWidget(self.system_combo)
        
        group_step2.setLayout(layout_step2)
        content_layout.addWidget(group_step2)

        content_layout.addStretch() 
        master_layout.addWidget(content_widget, 1)

        # --- RODAPÉ ---
        footer_widget = QWidget()
        footer_widget.setObjectName("Footer")
        footer_layout = QHBoxLayout(footer_widget)
        footer_layout.setContentsMargins(30, 15, 30, 15)
        
        status_layout = QVBoxLayout()
        status_layout.setSpacing(5)
        
        self.status_label = QLabel("Aguardando processamento...")
        self.status_label.setObjectName("StatusLabel")
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setFixedWidth(250)
        
        status_layout.addWidget(self.status_label)
        status_layout.addWidget(self.progress_bar)
        
        footer_layout.addLayout(status_layout)
        footer_layout.addStretch()

        self.run_button = QPushButton("Iniciar Tabulação")
        self.run_button.setObjectName("btn_processar")
        self.run_button.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.run_button.clicked.connect(self.run_pure_python_tabulation)
        
        footer_layout.addWidget(self.run_button, alignment=Qt.AlignmentFlag.AlignBottom)
        master_layout.addWidget(footer_widget, 0)
        
    def set_progress_state(self, state):
        self.progress_bar.setProperty("state", state)
        self.progress_bar.style().unpolish(self.progress_bar)
        self.progress_bar.style().polish(self.progress_bar)

    def browse_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Selecione os Balancetes e Plano de Contas",
            "",
            "Arquivos Excel (*.xlsx *.xls *.csv)"
        )
        if files:
            self.arquivos_selecionados = files
            self.file_list_widget.clear()
            
            self.set_progress_state("default")
            self.progress_bar.setValue(0)
            self.status_label.setText(f"{len(files)} arquivo(s) selecionado(s).")
            
            for file_path in files:
                self.file_list_widget.addItem(os.path.basename(file_path))

    def run_pure_python_tabulation(self):
        if not self.arquivos_selecionados:
            QMessageBox.warning(self, "Atenção", "Nenhum arquivo foi selecionado.")
            return

        sistema_selecionado = self.system_combo.currentText()
        
        # Alerta amigável e proteção para Girando Sol caso tenha Excel
        if sistema_selecionado == "Girando Sol":
            arquivos_excel = [f for f in self.arquivos_selecionados if f.lower().endswith(('.xls', '.xlsx'))]
            if arquivos_excel:
                QMessageBox.warning(self, "Aviso", "Arquivos Excel foram identificados.\n\nO sistema 'Girando Sol' utiliza apenas arquivos CSV. Os arquivos Excel selecionados serão ignorados no processamento.")

        self.run_button.setEnabled(False)
        self.set_progress_state("default")
        self.progress_bar.setRange(0, 0) 
        self.status_label.setText("Processando arquivos... Isso pode levar um momento.")
        QApplication.processEvents()

        try:
            retorno_processamento = processar_arquivos_selecionados(self.arquivos_selecionados, sistema_selecionado)

            if not retorno_processamento:
                raise ValueError("Nenhum dado foi processado. Verifique os arquivos de entrada.")

            if isinstance(retorno_processamento, tuple):
                dicionario_originais, dicionario_recalculados = retorno_processamento
            else:
                dicionario_originais = retorno_processamento
                dicionario_recalculados = None

            self.salvar_planilha_final(dicionario_originais, dicionario_recalculados)
            
            self.progress_bar.setRange(0, 100)
            self.progress_bar.setValue(100)
            self.set_progress_state("success")
            self.status_label.setText("Concluído! Planilha gerada com sucesso.")

        except Exception as e:
            self.progress_bar.setRange(0, 100)
            self.progress_bar.setValue(100)
            self.set_progress_state("error")
            self.status_label.setText("Ocorreu um erro no processamento.")
            QMessageBox.critical(self, "Erro durante o processamento", f"Ocorreu um erro.\n\nDetalhes: {e}")
            
        finally:
            self.run_button.setEnabled(True)
            
    def salvar_planilha_final(self, dicionario_dataframes, dicionario_recalculados=None):
        caminho_salvar, _ = QFileDialog.getSaveFileName(
            self, "Salvar Planilha Tabulada", "", "Arquivos Excel (*.xlsx)"
        )

        if not caminho_salvar:
            return

        def aplicar_formatacao(writer_obj, dicionario):
            for nome_aba in sorted(dicionario.keys()):
                dataframe = dicionario[nome_aba]
                dataframe.to_excel(writer_obj, index=False, sheet_name=nome_aba)
                
                worksheet = writer_obj.sheets[nome_aba]
                
                def get_col_letter(n):
                    string_col = ""
                    while n > 0:
                        n, remainder = divmod(n - 1, 26)
                        string_col = chr(65 + remainder) + string_col
                    return string_col

                for idx, col in enumerate(dataframe.columns):
                    header_len = len(str(col))
                    max_data_len = dataframe[col].astype(str).map(len).max() if not dataframe.empty else 0
                    max_len = max(header_len, max_data_len) + 3
                    
                    if max_len > 60:
                        max_len = 60
                        
                    col_letter = get_col_letter(idx + 1)
                    worksheet.column_dimensions[col_letter].width = max_len

        # 1. Salva o arquivo principal
        with pd.ExcelWriter(caminho_salvar, engine='openpyxl') as writer:
            aplicar_formatacao(writer, dicionario_dataframes)

        # 2. Salva o arquivo recalculado (se o sistema Girando Sol tiver devolvido um)
        caminho_recalculado = ""
        if dicionario_recalculados:
            nome_base, extensao = os.path.splitext(caminho_salvar)
            caminho_recalculado = f"{nome_base}_recalculada{extensao}"
            
            with pd.ExcelWriter(caminho_recalculado, engine='openpyxl') as writer:
                aplicar_formatacao(writer, dicionario_recalculados)

        if caminho_recalculado:
            QMessageBox.information(self, "Sucesso", f"Arquivos salvos em:\n\n1. {caminho_salvar}\n2. {caminho_recalculado}")
        else:
            QMessageBox.information(self, "Sucesso", f"Arquivo salvo em:\n{caminho_salvar}")


def main():
    app = QApplication.instance()
    if not app:
        app = QApplication(sys.argv)
    window = TabuladorPythonApp()
    window.show()
    return app, window

if __name__ == '__main__':
    app = QApplication(sys.argv)
    
    theme_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..', 'theme_light.qss'))
    if os.path.exists(theme_path):
        with open(theme_path, "r", encoding="utf-8") as f:
            app.setStyleSheet(f.read())
            
    window = TabuladorPythonApp()
    window.show()
    sys.exit(app.exec())