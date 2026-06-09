import sys
import os
import pandas as pd
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout,
    QPushButton, QLabel, QListWidget, QAbstractItemView,
    QFileDialog, QMessageBox, QComboBox
)
from PyQt6.QtGui import QFont
from PyQt6.QtCore import Qt

# Tratamento de caminhos para suportar a execucao tanto via HUB quanto standalone
try:
    from scripts.Contabilidade.tabulador_logic import processar_arquivos_selecionados, obter_nomes_sistemas
except ModuleNotFoundError:
    from tabulador_logic import processar_arquivos_selecionados, obter_nomes_sistemas

# Carrega a lista de sistemas configurados no motor de logica de forma dinamica
LISTA_SISTEMAS = obter_nomes_sistemas()

class TabuladorPythonApp(QWidget):
    def __init__(self):
        super().__init__()
        self.arquivos_selecionados = []
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Tabulador de Balancetes")
        self.setGeometry(300, 300, 600, 450)
        
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # Componente de selecao de multiplos arquivos de entrada
        main_layout.addWidget(QLabel("1. Selecione um ou mais arquivos de balancete:"))
        self.file_list_widget = QListWidget()
        self.file_list_widget.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        main_layout.addWidget(self.file_list_widget)

        browse_button = QPushButton("Procurar Arquivos...")
        browse_button.clicked.connect(self.browse_files)
        main_layout.addWidget(browse_button)

        # Menu dropdown alimentado pelo dicionario do tabulador_logic
        main_layout.addWidget(QLabel("2. Selecione o sistema:"))
        self.system_combo = QComboBox()
        self.system_combo.addItems(LISTA_SISTEMAS)
        main_layout.addWidget(self.system_combo)

        # Controle de disparo da tabulacao
        self.run_button = QPushButton("Iniciar Tabulação")
        self.run_button.setFont(QFont("Segoe UI", 12, QFont.Weight.Bold))
        self.run_button.clicked.connect(self.run_pure_python_tabulation)
        main_layout.addWidget(self.run_button)
        
        self.status_label = QLabel("Pronto para começar.")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(self.status_label)
        
    def browse_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Selecione os Balancetes e Plano de Contas",
            "",
            "Arquivos Excel (*.xlsx *.xls)"
        )
        if files:
            self.arquivos_selecionados = files
            self.file_list_widget.clear()
            for file_path in files:
                self.file_list_widget.addItem(os.path.basename(file_path))

    def run_pure_python_tabulation(self):
        if not self.arquivos_selecionados:
            QMessageBox.warning(self, "Atenção", "Nenhum arquivo foi selecionado.")
            return

        self.run_button.setEnabled(False)
        self.status_label.setText("Processando arquivos... Isso pode levar um momento.")
        QApplication.processEvents()

        try:
            # Captura a opcao selecionada na combo para direcionamento do plugin correto
            sistema_selecionado = self.system_combo.currentText()
            
            dicionario_de_dfs = processar_arquivos_selecionados(self.arquivos_selecionados, sistema_selecionado)

            if not dicionario_de_dfs:
                raise ValueError("Nenhum dado foi processado. Verifique os arquivos de entrada.")

            self.salvar_planilha_final(dicionario_de_dfs)
            self.status_label.setText("Concluido! Planilha gerada com sucesso.")

        except Exception as e:
            self.status_label.setText("Ocorreu um erro no processamento.")
            QMessageBox.critical(self, "Erro durante o processamento", f"Ocorreu um erro.\n\nDetalhes: {e}")
        finally:
            self.run_button.setEnabled(True)
            
    def salvar_planilha_final(self, dicionario_dataframes):
        caminho_salvar, _ = QFileDialog.getSaveFileName(
            self, "Salvar Planilha Tabulada", "", "Arquivos Excel (*.xlsx)"
        )

        if caminho_salvar:
            with pd.ExcelWriter(caminho_salvar, engine='openpyxl') as writer:
                for nome_aba in sorted(dicionario_dataframes.keys()):
                    dataframe = dicionario_dataframes[nome_aba]
                    dataframe.to_excel(writer, index=False, sheet_name=nome_aba)
                    
                    # --- INÍCIO DA PERFUMARIA: Auto-Ajuste de Largura das Colunas ---
                    worksheet = writer.sheets[nome_aba]
                    
                    # Função auxiliar para pegar a letra da coluna (1 = A, 2 = B, 27 = AA...)
                    def get_col_letter(n):
                        string_col = ""
                        while n > 0:
                            n, remainder = divmod(n - 1, 26)
                            string_col = chr(65 + remainder) + string_col
                        return string_col

                    for idx, col in enumerate(dataframe.columns):
                        header_len = len(str(col))
                        # Conta o tamanho máximo dos dados na coluna. Se vazia, assume 0.
                        max_data_len = dataframe[col].astype(str).map(len).max() if not dataframe.empty else 0
                        
                        # Define a largura pegando o maior valor (cabeçalho ou dados) + 3 de respiro visual
                        max_len = max(header_len, max_data_len) + 3
                        
                        # Trava em 60 para não deixar uma coluna infinitamente larga (ex: textos de Avisos/Erros)
                        if max_len > 60:
                            max_len = 60
                            
                        col_letter = get_col_letter(idx + 1)
                        worksheet.column_dimensions[col_letter].width = max_len
                    # --- FIM DA PERFUMARIA ---

            QMessageBox.information(self, "Sucesso", f"Arquivo salvo em:\n{caminho_salvar}")


# Ponto de entrada obrigatorio exigido pelo hub.py para carregar a janela na interface central
def main():
    app = QApplication.instance()
    
    # Fallback para criacao de instancia caso o modulo seja chamado fora do HUB
    if not app:
        app = QApplication(sys.argv)
        
    window = TabuladorPythonApp()
    window.show()
    return app, window


# Bloco para execucao direta e isolada do script (Modo de desenvolvimento/testes)
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = TabuladorPythonApp()
    window.show()
    sys.exit(app.exec())