# ===== ANALISADOR SPED EFD CONTRIBUIÇÕES v1.1 - BLOCO F100 (COM DESCRIÇÃO) =====

import sys
import os
import pandas as pd
from datetime import datetime
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout,
    QPushButton, QLabel, QFileDialog, QMessageBox, QListWidget
)
from PyQt6.QtGui import QFont
from PyQt6.QtCore import Qt

# ==============================================================================
# PARTE 1: LÓGICA DE PROCESSAMENTO DO ARQUIVO SPED
# ==============================================================================

def parse_line(line):
    """Lê uma linha do SPED, remove o lixo e a divide em campos."""
    line = line.strip()
    if line and line.startswith('|') and line.endswith('|'):
        return line[1:-1].split('|')
    return []

def safe_get(data_list, index, default=''):
    """Acessa um índice de lista de forma segura."""
    return data_list[index] if len(data_list) > index else default

def safe_float_convert(value_str):
    """Converte uma string para float de forma segura (trata vírgulas e erros)."""
    if not value_str: return 0.0
    try:
        return float(value_str.replace(',', '.'))
    except (ValueError, TypeError):
        return 0.0

def format_date(date_str):
    """Formata uma data do padrão 'ddmmyyyy' para 'dd/mm/yyyy'."""
    if date_str and len(date_str) == 8 and date_str.isdigit():
        try:
            return datetime.strptime(date_str, '%d%m%Y').strftime('%d/%m/%Y')
        except ValueError:
            return ''
    return ''

# **NOVO: Dicionário para a Natureza da Base de Cálculo do Crédito**
NAT_BC_CRED_MAP = {
    '01': '01 - Aquisição de bens para revenda',
    '02': '02 - Aquisição de bens utilizados como insumo',
    '03': '03 - Aquisição de serviços utilizados como insumo',
    '04': '04 - Energia elétrica e térmica, inclusive sob a forma de vapor',
    '05': '05 - Aluguéis de prédios',
    '06': '06 - Aluguéis de máquinas e equipamentos',
    '07': '07 - Armazenagem de mercadoria e frete na operação de venda',
    '08': '08 - Contraprestações de arrendamento mercantil',
    '09': '09 - Máquinas, equip. e outros bens ao ativo imobilizado (crédito sobre depreciação)',
    '10': '10 - Máquinas, equip. e outros bens ao ativo imobilizado (crédito sobre aquisição)',
    '11': '11 - Amortização e Depreciação de edificações e benfeitorias em imóveis',
    '12': '12 - Devolução de Vendas Sujeitas à Incidência Não-Cumulativa',
    '13': '13 - Outras Operações com Direito a Crédito',
    '14': '14 - Transporte de Cargas – Contratação de pessoa física ou PJ optante pelo SIMPLES',
    '15': '15 - Atividade Imobiliária – Custo Incorrido de Unidade Imobiliária',
    '16': '16 - Atividade Imobiliária – Custo Orçado de unidade não concluída',
    '17': '17 - Atividade de Serviços – vale-transporte, vale-refeição ou vale-alimentação, fardamento ou uniforme',
    '18': '18 - Estoque de abertura de bens'
}

def gerar_relatorio_bloco_f100(lista_arquivos):
    """
    Função principal que implementa a lógica de extração para o Bloco F100,
    utilizando o CNPJ do registro F010.
    """
    
    report_rows = []
    print("Iniciando Etapa 1: Processamento do Bloco F100...")
    
    for filepath in lista_arquivos:
        nome_arquivo = os.path.basename(filepath)
        print(f"  - Processando arquivo: {nome_arquivo}")
        
        with open(filepath, 'r', encoding='latin-1') as f:
            cnpj_estabelecimento_atual = ""

            for line in f:
                campos = parse_line(line)
                if not campos: continue
                reg = campos[0]
                
                if reg == 'F010':
                    cnpj_estabelecimento_atual = safe_get(campos, 1)
                
                elif reg == 'F100':
                    nat_bc_cred_cod = safe_get(campos, 14)
                    
                    row_data = {
                        'CNPJ Estabelecimento': cnpj_estabelecimento_atual,
                        'Tipo Operação': safe_get(campos, 1),
                        'Cód. Participante': safe_get(campos, 2),
                        'Cód. Item': safe_get(campos, 3),
                        'Data Operação': format_date(safe_get(campos, 4)),
                        'Vlr. Operação': safe_float_convert(safe_get(campos, 5)),
                        'CST PIS': safe_get(campos, 6),
                        'Vlr. BC PIS': safe_float_convert(safe_get(campos, 7)),
                        'Alíquota PIS (%)': safe_float_convert(safe_get(campos, 8)),
                        'Vlr. PIS': safe_float_convert(safe_get(campos, 9)),
                        'CST COFINS': safe_get(campos, 10),
                        'Vlr. BC COFINS': safe_float_convert(safe_get(campos, 11)),
                        'Alíquota COFINS (%)': safe_float_convert(safe_get(campos, 12)),
                        'Vlr. COFINS': safe_float_convert(safe_get(campos, 13)),
                        'Nat. BC Crédito': nat_bc_cred_cod,
                        # **NOVO: Adiciona a descrição usando o dicionário**
                        'Descrição Natureza Crédito': NAT_BC_CRED_MAP.get(nat_bc_cred_cod, f"Código {nat_bc_cred_cod} não encontrado"),
                        'Ind. Origem Crédito': safe_get(campos, 15),
                        'Cód. Conta Contábil': safe_get(campos, 16),
                        'Cód. Centro Custo': safe_get(campos, 17),
                        'Descrição Compl.': safe_get(campos, 18),
                    }
                    report_rows.append(row_data)

    if not report_rows:
        raise ValueError("Nenhum registro F100 foi encontrado nos arquivos selecionados.")

    df_final = pd.DataFrame(report_rows)
    
    # **ALTERADO: Nova coluna adicionada à lista final**
    colunas_finais = [
        'CNPJ Estabelecimento', 'Data Operação', 'Tipo Operação', 
        'Vlr. Operação', 'Nat. BC Crédito', 'Descrição Natureza Crédito',
        'CST PIS', 'Vlr. BC PIS', 'Alíquota PIS (%)', 'Vlr. PIS',
        'CST COFINS', 'Vlr. BC COFINS', 'Alíquota COFINS (%)', 'Vlr. COFINS',
        'Cód. Participante', 'Cód. Item', 'Ind. Origem Crédito',
        'Cód. Conta Contábil', 'Cód. Centro Custo', 'Descrição Compl.'
    ]
    
    for col in colunas_finais:
        if col not in df_final.columns:
            df_final[col] = ''
            
    return df_final[colunas_finais]

# ==============================================================================
# PARTE 2: INTERFACE GRÁFICA (PyQt6)
# ==============================================================================

class SpedF100App(QWidget):
    def __init__(self):
        super().__init__()
        self.selected_files = []
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Analisador EFD Contribuições - Bloco F100 v1.1")
        self.setGeometry(200, 200, 700, 450)
        
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)
        
        main_layout.addWidget(QLabel("1. Selecione um ou mais arquivos EFD Contribuições (.txt):"))
        self.file_list_widget = QListWidget()
        main_layout.addWidget(self.file_list_widget)
        
        browse_button = QPushButton("Procurar Arquivos...")
        browse_button.clicked.connect(self.select_files)
        main_layout.addWidget(browse_button)
        
        self.run_button = QPushButton("🚀 Gerar Relatório do Bloco F100")
        self.run_button.setFont(QFont("Segoe UI", 12, QFont.Weight.Bold))
        self.run_button.clicked.connect(self.start_processing)
        main_layout.addWidget(self.run_button)
        
        self.status_label = QLabel("Pronto para começar.")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(self.status_label)
        
        main_layout.addStretch()

    def select_files(self):
        filepaths, _ = QFileDialog.getOpenFileNames(self, "Selecione os arquivos SPED", "", "Arquivos de Texto (*.txt)")
        if filepaths:
            self.selected_files = filepaths
            self.file_list_widget.clear()
            for file_path in self.selected_files:
                self.file_list_widget.addItem(os.path.basename(file_path))

    def start_processing(self):
        if not self.selected_files:
            QMessageBox.warning(self, "Atenção", "Nenhum arquivo foi selecionado.")
            return
            
        self.run_button.setEnabled(False)
        self.status_label.setText("Processando... Isso pode levar um momento.")
        QApplication.processEvents()

        try:
            df_result = gerar_relatorio_bloco_f100(self.selected_files)
            self.save_report(df_result)
            self.status_label.setText("✅ Relatório do Bloco F100 gerado com sucesso!")
        except Exception as e:
            self.status_label.setText(f"❌ Erro: {e}")
            QMessageBox.critical(self, "Erro de Processamento", f"Ocorreu um erro.\n\nDetalhes: {e}")
        finally:
            self.run_button.setEnabled(True)

    def save_report(self, dataframe):
        default_filename = f"Relatorio_Bloco_F100_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        save_path, _ = QFileDialog.getSaveFileName(self, "Salvar Relatório", default_filename, "Arquivos Excel (*.xlsx)")
        
        if save_path:
            try:
                self.status_label.setText("Salvando arquivo Excel...")
                QApplication.processEvents()
                dataframe.to_excel(save_path, index=False, sheet_name="Bloco_F100")
                QMessageBox.information(self, "Sucesso", f"Relatório salvo com sucesso em:\n{save_path}")
            except Exception as e:
                self.status_label.setText(f"❌ Erro ao salvar o arquivo: {e}")
                QMessageBox.critical(self, "Erro ao Salvar", f"Não foi possível salvar o arquivo.\n\nDetalhes: {e}")

# ==============================================================================
# PARTE 3: PONTO DE ENTRADA
# ==============================================================================

def main():
    """Função principal que inicia esta aplicação específica."""
    app = QApplication.instance() or QApplication(sys.argv)
    
    window = SpedF100App()
    window.show()

    return app, window

# Este bloco permite que o script ainda seja executável de forma independente para testes
if __name__ == '__main__':
    app, window = main()
    sys.exit(app.exec())