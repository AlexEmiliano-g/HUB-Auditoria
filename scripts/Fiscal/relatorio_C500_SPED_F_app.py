# ===== ANALISADOR SPED FISCAL v2.2 - BLOCO C500 (MODIFICADO PARA O HUB) =====

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
# (NENHUMA ALTERAÇÃO NECESSÁRIA AQUI)
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
        return float(value_str.replace('.', '').replace(',', '.'))
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

# Dicionário para traduzir o código da Classe de Consumo
CLASSE_CONSUMO_MAP = {
    '01': '01 - Comercial', '02': '02 - Consumo Próprio',
    '03': '03 - Iluminação Pública', '04': '04 - Industrial',
    '05': '05 - Poder Público', '06': '06 - Residencial',
    '07': '07 - Rural', '08': '08 - Serviço Público'
}

def gerar_relatorio_bloco_c500(lista_arquivos):
    """Função principal que implementa a lógica de extração para o Bloco C500."""
    
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
                        'CNPJ': safe_get(campos, 5), 'Inscrição Estadual': safe_get(campos, 7),
                        'Período': f"{format_date(safe_get(campos, 3))} a {format_date(safe_get(campos, 4))}"
                    }
                elif reg == '0150':
                    cod_part = safe_get(campos, 1)
                    if cod_part and cod_part not in mapa_participantes:
                        mapa_participantes[cod_part] = {
                            'Nome Participante': safe_get(campos, 2), 
                            'CNPJ/CPF Participante': safe_get(campos, 4) or safe_get(campos, 5),
                        }
    
    # --- ETAPA 2: Processar Bloco C500 e montar o relatório ---
    report_rows = []
    print("\nIniciando Etapa 2: Processamento do Bloco C500...")
    for filepath in lista_arquivos:
        nome_arquivo = os.path.basename(filepath)
        print(f"  - Processando arquivo: {nome_arquivo}")
        info_empresa = mapa_empresa.get(nome_arquivo, {})
        
        with open(filepath, 'r', encoding='latin-1') as f:
            current_c500_data = {}
            for line in f:
                campos = parse_line(line)
                if not campos or not campos[0].startswith('C5'): continue
                reg = campos[0]
                
                if reg == 'C500':
                    cod_part = safe_get(campos, 3) 
                    participante = mapa_participantes.get(cod_part, {})
                    cod_classe = safe_get(campos, 8) 
                    
                    current_c500_data = {
                        **info_empresa,
                        'Tipo Operação': 'Entrada' if safe_get(campos, 1) == '0' else 'Saída',
                        'Código Participante': cod_part,
                        'Nome Participante': participante.get('Nome Participante', ''),
                        'Modelo Documento': safe_get(campos, 4),
                        'Situação Documento': safe_get(campos, 5),
                        'Série': safe_get(campos, 6),
                        'SubSérie': safe_get(campos, 7),
                        'Classe de Consumo': CLASSE_CONSUMO_MAP.get(cod_classe, cod_classe),
                        'Número Documento': safe_get(campos, 9),
                        'Data Documento': format_date(safe_get(campos, 10)),
                        'Data Entrada/Saída': format_date(safe_get(campos, 11)),
                        'Vlr Documento': safe_float_convert(safe_get(campos, 12)),
                        'Vlr Desconto': safe_float_convert(safe_get(campos, 13)),
                        'Vlr PIS': safe_float_convert(safe_get(campos, 22)),
                        'Vlr COFINS': safe_float_convert(safe_get(campos, 26)),
                    }
                elif reg == 'C590' and current_c500_data:
                    row = {
                        **current_c500_data,
                        'CST ICMS': safe_get(campos, 1),
                        'CFOP': safe_get(campos, 2),
                        'Alíquota ICMS (%)': safe_float_convert(safe_get(campos, 3)),
                        'Vlr Operação': safe_float_convert(safe_get(campos, 4)),
                        'Vlr Base Cálculo ICMS': safe_float_convert(safe_get(campos, 5)),
                        'Vlr ICMS': safe_float_convert(safe_get(campos, 6)),
                        'Vlr Base Cálculo ICMS ST': safe_float_convert(safe_get(campos, 7)),
                        'Vlr ICMS ST': safe_float_convert(safe_get(campos, 8)),
                    }
                    report_rows.append(row)

    if not report_rows:
        raise ValueError("Nenhum registro C590 (detalhamento de contas) foi encontrado nos arquivos selecionados.")

    df_final = pd.DataFrame(report_rows)
    
    colunas_finais = [
        'CNPJ', 'Inscrição Estadual', 'Período', 'Data Documento', 'Data Entrada/Saída', 
        'Tipo Operação', 'Número Documento', 'Série', 'SubSérie', 'Modelo Documento', 
        'Nome Participante', 'Classe de Consumo', 'Vlr Documento', 'Vlr Desconto', 
        'CFOP', 'CST ICMS', 'Alíquota ICMS (%)', 'Vlr Operação', 
        'Vlr Base Cálculo ICMS', 'Vlr ICMS', 
        'Vlr Base Cálculo ICMS ST', 'Vlr ICMS ST',
        'Vlr PIS', 'Vlr COFINS'
    ]
    # Garante que todas as colunas existam antes de reordenar
    for col in colunas_finais:
        if col not in df_final.columns:
            df_final[col] = ''
            
    return df_final[colunas_finais]

# ==============================================================================
# PARTE 2: INTERFACE GRÁFICA (PyQt6)
# (NENHUMA ALTERAÇÃO NECESSÁRIA AQUI)
# ==============================================================================

class SpedC500App(QWidget):
    def __init__(self):
        super().__init__()
        self.selected_files = []
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Gerador de Relatório SPED Fiscal - Bloco C500")
        self.setGeometry(200, 200, 700, 450)
        
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)
        
        main_layout.addWidget(QLabel("1. Selecione um ou mais arquivos SPED (.txt):"))
        self.file_list_widget = QListWidget()
        main_layout.addWidget(self.file_list_widget)
        
        browse_button = QPushButton("Procurar Arquivos...")
        browse_button.clicked.connect(self.select_files)
        main_layout.addWidget(browse_button)
        
        self.run_button = QPushButton("🚀 Gerar Relatório do Bloco C500")
        self.run_button.setFont(QFont("Segoe UI", 12, QFont.Weight.Bold))
        self.run_button.clicked.connect(self.start_processing)
        main_layout.addWidget(self.run_button)
        
        self.status_label = QLabel("Pronto para começar.")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(self.status_label)
        
        main_layout.addStretch()

    def select_files(self):
        filepaths, _ = QFileDialog.getOpenFileNames(self, "Selecione os arquivos SPED Fiscal", "", "Arquivos de Texto (*.txt)")
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
        self.status_label.setText("Processando múltiplos arquivos... Isso pode levar um momento.")
        QApplication.processEvents()

        try:
            df_result = gerar_relatorio_bloco_c500(self.selected_files)
            self.save_report(df_result)
            self.status_label.setText("✅ Relatório consolidado do Bloco C500 gerado com sucesso!")
        except Exception as e:
            self.status_label.setText(f"❌ Erro: {e}")
            QMessageBox.critical(self, "Erro de Processamento", f"Ocorreu um erro.\n\nDetalhes: {e}")
        finally:
            self.run_button.setEnabled(True)

    def save_report(self, dataframe):
        default_filename = f"Relatorio_Bloco_C500_{datetime.now().strftime('%Y%m%d')}.xlsx"
        save_path, _ = QFileDialog.getSaveFileName(self, "Salvar Relatório", default_filename, "Arquivos Excel (*.xlsx)")
        
        if save_path:
            try:
                self.status_label.setText("Salvando arquivo Excel...")
                QApplication.processEvents()
                dataframe.to_excel(save_path, index=False, sheet_name="Bloco_C500")
                QMessageBox.information(self, "Sucesso", f"Relatório salvo em:\n{save_path}")
            except Exception as e:
                self.status_label.setText(f"❌ Erro ao salvar o arquivo: {e}")
                QMessageBox.critical(self, "Erro ao Salvar", f"Não foi possível salvar o arquivo.\n\nDetalhes: {e}")

# ==============================================================================
# PARTE 3: PONTO DE ENTRADA (MODIFICADO E CORRIGIDO)
# ==============================================================================

def main():
    """Função principal que inicia esta aplicação específica."""
    # Reutiliza a instância da aplicação principal (do Hub) se ela existir.
    # Caso contrário, cria uma nova para execução independente.
    app = QApplication.instance() or QApplication(sys.argv)

    # Cria e exibe a janela da aplicação.
    window = SpedC500App()
    window.show()

    # Retorna a aplicação e a janela para o chamador poder gerenciá-las.
    return app, window

# Este bloco permite que o script ainda seja executável de forma independente para testes
if __name__ == '__main__':
    app, window = main()
    sys.exit(app.exec())