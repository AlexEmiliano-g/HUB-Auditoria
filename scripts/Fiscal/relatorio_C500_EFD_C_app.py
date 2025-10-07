# ===== ANALISADOR SPED EFD CONTRIBUIÇÕES v4.0 - REGISTRO C500 (COM C010) =====

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
# (LÓGICA DO C010 IMPLEMENTADA)
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

def gerar_relatorio_bloco_c500(lista_arquivos):
    """
    Função principal que associa o CNPJ do registro C010 aos registros C500.
    """
    
    # --- ETAPA 1: Mapear dados de apoio de todos os arquivos ---
    mapa_empresa, mapa_participantes = {}, {}
    
    print("Iniciando Etapa 1: Mapeamento de dados de apoio (0000, 0150)...")
    for filepath in lista_arquivos:
        nome_arquivo = os.path.basename(filepath)
        print(f"  - Mapeando arquivo: {nome_arquivo}")
        with open(filepath, 'r', encoding='latin-1') as f:
            for line in f:
                campos = parse_line(line)
                if not campos: continue
                reg = campos[0]
                
                if reg == '0000' and nome_arquivo not in mapa_empresa:
                    mapa_empresa[nome_arquivo] = {
                        # CNPJ do 0000 é do declarante, não necessariamente do estabelecimento
                        'Nome Empresa Declarante': safe_get(campos, 6),
                        'Período': f"{format_date(safe_get(campos, 4))} a {format_date(safe_get(campos, 5))}"
                    }
                elif reg == '0150':
                    cod_part = safe_get(campos, 1)
                    if cod_part and cod_part not in mapa_participantes:
                        mapa_participantes[cod_part] = {
                            'Nome Participante': safe_get(campos, 2), 
                            'CNPJ/CPF Participante': safe_get(campos, 5) or safe_get(campos, 6),
                        }
    
    # --- ETAPA 2: Processar Blocos C010/C500 e montar o relatório ---
    report_rows = []
    print("\nIniciando Etapa 2: Processamento dos Blocos C010 e C500...")
    for filepath in lista_arquivos:
        nome_arquivo = os.path.basename(filepath)
        print(f"  - Processando arquivo: {nome_arquivo}")
        info_empresa = mapa_empresa.get(nome_arquivo, {})
        
        with open(filepath, 'r', encoding='latin-1') as f:
            current_note_data = {}
            cnpj_estabelecimento_atual = "" # Variável para armazenar o CNPJ do C010

            def flush_current_note():
                if current_note_data and 'Número Documento' in current_note_data:
                    report_rows.append(current_note_data.copy())
                    current_note_data.clear()

            for line in f:
                campos = parse_line(line)
                if not campos: continue
                reg = campos[0]
                
                # NOVO: Captura o CNPJ do estabelecimento
                if reg == 'C010':
                    cnpj_estabelecimento_atual = safe_get(campos, 1)

                elif reg == 'C500':
                    flush_current_note()
                    cod_part = safe_get(campos, 1) 
                    participante = mapa_participantes.get(cod_part, {})
                    
                    current_note_data = {
                        **info_empresa,
                        'CNPJ Estabelecimento': cnpj_estabelecimento_atual, # Adiciona o CNPJ correto
                        'Código Participante': cod_part,
                        'Nome Participante': participante.get('Nome Participante', ''),
                        'CNPJ/CPF Participante': participante.get('CNPJ/CPF Participante', ''),
                        'Modelo Documento': safe_get(campos, 2),
                        'Situação Documento': safe_get(campos, 3),
                        'Série': safe_get(campos, 4),
                        'Número Documento': safe_get(campos, 6),
                        'Data Documento': format_date(safe_get(campos, 7)),
                        'Data Entrada': format_date(safe_get(campos, 8)),
                        'Vlr Documento': safe_float_convert(safe_get(campos, 9)),
                        'Vlr Desconto': safe_float_convert(safe_get(campos, 10)),
                        'Vlr PIS Total (Doc)': safe_float_convert(safe_get(campos, 12)),
                        'Vlr COFINS Total (Doc)': safe_float_convert(safe_get(campos, 13)),
                    }
                
                elif reg == 'C501' and current_note_data:
                    current_note_data.update({
                        'CST PIS': safe_get(campos, 1),
                        'Vlr Base Cálculo PIS': safe_float_convert(safe_get(campos, 2)),
                        'Alíquota PIS (%)': safe_float_convert(safe_get(campos, 3)),
                        'Alíquota PIS (R$)': safe_float_convert(safe_get(campos, 5)),
                        'Vlr PIS': safe_float_convert(safe_get(campos, 6)),
                        'Código Conta Analítica': safe_get(campos, 7),
                    })
                
                elif reg == 'C505' and current_note_data:
                    current_note_data.update({
                        'CST COFINS': safe_get(campos, 1),
                        'Vlr Base Cálculo COFINS': safe_float_convert(safe_get(campos, 2)),
                        'Alíquota COFINS (%)': safe_float_convert(safe_get(campos, 3)),
                        'Alíquota COFINS (R$)': safe_float_convert(safe_get(campos, 5)),
                        'Vlr COFINS': safe_float_convert(safe_get(campos, 6)),
                    })
            
            flush_current_note()

    if not report_rows:
        raise ValueError("Nenhum registro C500 foi encontrado nos arquivos selecionados.")

    df_final = pd.DataFrame(report_rows)
    
    colunas_finais = [
        'CNPJ Estabelecimento', 'Nome Empresa Declarante', 'Período', 
        'Data Documento', 'Data Entrada', 'Número Documento', 'Série', 
        'Modelo Documento', 'Situação Documento',
        'Nome Participante', 'CNPJ/CPF Participante', 
        'Vlr Documento', 'Vlr Desconto', 
        'CST PIS', 'Vlr Base Cálculo PIS', 'Alíquota PIS (%)', 'Alíquota PIS (R$)', 'Vlr PIS',
        'CST COFINS', 'Vlr Base Cálculo COFINS', 'Alíquota COFINS (%)', 'Alíquota COFINS (R$)', 'Vlr COFINS',
        'Código Conta Analítica',
        'Vlr PIS Total (Doc)', 'Vlr COFINS Total (Doc)'
    ]
    
    for col in colunas_finais:
        if col not in df_final.columns:
            df_final[col] = ''
            
    return df_final[colunas_finais]

# ==============================================================================
# PARTE 2: INTERFACE GRÁFICA (PyQt6)
# ==============================================================================

class SpedC500App(QWidget):
    def __init__(self):
        super().__init__()
        self.selected_files = []
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Gerador de Relatório EFD Contribuições - Bloco C500 v4.0")
        self.setGeometry(200, 200, 800, 500)
        
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)
        
        main_layout.addWidget(QLabel("1. Selecione um ou mais arquivos SPED EFD Contribuições (.txt):"))
        self.file_list_widget = QListWidget()
        main_layout.addWidget(self.file_list_widget)
        
        browse_button = QPushButton("Procurar Arquivos...")
        browse_button.clicked.connect(self.select_files)
        main_layout.addWidget(browse_button)
        
        self.run_button = QPushButton("🚀 Gerar Relatório C500 (com CNPJ do Estabelecimento)")
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
            QMessageBox.warning(self, "Atenção", "Nenhum arquivo SPED foi selecionado.")
            return
            
        self.run_button.setEnabled(False)
        self.status_label.setText("Processando... Por favor, aguarde.")
        QApplication.processEvents()

        try:
            df_result = gerar_relatorio_bloco_c500(self.selected_files)
            self.save_report(df_result)
            self.status_label.setText("✅ Relatório do Bloco C500 gerado com sucesso!")
        except Exception as e:
            self.status_label.setText(f"❌ Erro: {e}")
            QMessageBox.critical(self, "Erro de Processamento", f"Ocorreu um erro.\n\nDetalhes: {e}")
        finally:
            self.run_button.setEnabled(True)

    def save_report(self, dataframe):
        default_filename = f"Relatorio_C500_EFD_Contrib_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        save_path, _ = QFileDialog.getSaveFileName(self, "Salvar Relatório", default_filename, "Arquivos Excel (*.xlsx)")
        
        if save_path:
            try:
                self.status_label.setText("Salvando arquivo Excel...")
                QApplication.processEvents()
                dataframe.to_excel(save_path, index=False, sheet_name="C500_EFD_Contrib")
                QMessageBox.information(self, "Sucesso", f"Relatório salvo em:\n{save_path}")
            except Exception as e:
                self.status_label.setText(f"❌ Erro ao salvar o arquivo: {e}")
                QMessageBox.critical(self, "Erro ao Salvar", f"Não foi possível salvar o arquivo.\n\nDetalhes: {e}")

# ==============================================================================
# PARTE 3: PONTO DE ENTRADA
# ==============================================================================

def main():
    app = QApplication.instance() or QApplication(sys.argv)
    window = SpedC500App()
    window.show()
    return app, window

if __name__ == '__main__':
    app, window = main()
    sys.exit(app.exec())